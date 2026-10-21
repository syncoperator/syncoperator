"""RSS/Atom-сборщик (feedparser) + автопоиск фида на HTML-странице."""

from __future__ import annotations

import calendar
import logging
import re
from datetime import datetime
from urllib.parse import urljoin

import feedparser
from bs4 import BeautifulSoup

from ..models import DateSource
from ..textutil import collapse_ws, html_to_text
from ..timeutil import UTC
from .base import CollectResult, RawItem, SourceConfig
from .http import FetchError, Fetcher

log = logging.getLogger(__name__)

FEED_TYPES = ("application/rss+xml", "application/atom+xml", "application/rdf+xml", "application/xml", "text/xml")


def _struct_to_dt(st) -> datetime | None:
    """feedparser отдаёт *_parsed уже в UTC (time.struct_time)."""
    if not st:
        return None
    try:
        return datetime.fromtimestamp(calendar.timegm(st), tz=UTC)
    except (OverflowError, ValueError, TypeError):
        return None


def _image(entry) -> str | None:
    for key in ("media_content", "media_thumbnail"):
        for m in entry.get(key) or []:
            if m.get("url"):
                return m["url"]
    for enc in entry.get("enclosures") or []:
        if str(enc.get("type", "")).startswith("image/") and enc.get("href"):
            return enc["href"]
    for link in entry.get("links") or []:
        if str(link.get("type", "")).startswith("image/") and link.get("href"):
            return link["href"]
    html = ""
    if entry.get("content"):
        html = entry["content"][0].get("value", "")
    html = html or entry.get("summary", "")
    m = re.search(r'<img[^>]+src=["\']([^"\']+)', html or "")
    return m.group(1) if m else None


def entry_to_item(entry, feed_url: str) -> RawItem | None:
    title = html_to_text(entry.get("title", ""))
    link = entry.get("link") or ""
    if link:
        link = urljoin(feed_url, link)
    if not title or not (link or entry.get("id")):
        return None
    content_html = ""
    if entry.get("content"):
        content_html = " ".join(c.get("value", "") for c in entry["content"])
    summary = html_to_text(entry.get("summary", ""))
    content = html_to_text(content_html) if content_html else None
    if content and summary and content.startswith(summary[:60]):
        content = None if len(content) <= len(summary) + 40 else content  # тизер = тот же текст
    published = _struct_to_dt(entry.get("published_parsed"))
    date_source = DateSource.feed_published
    if published is None:
        published = _struct_to_dt(entry.get("updated_parsed"))
        date_source = DateSource.feed_updated if published else DateSource.fetched
    tags = [collapse_ws(t.get("term", "")) for t in entry.get("tags") or [] if t.get("term")]
    return RawItem(
        external_id=(entry.get("id") or link)[:512],
        title=title,
        url=link or None,
        summary=summary or None,
        content=content,
        author=collapse_ws(entry.get("author", "")) or None,
        image_url=_image(entry),
        published_at=published,
        date_source=date_source,
        date_has_tz=True,
        extra={"feed_tags": tags[:12]} if tags else {},
    )


def discover_feed(html: str, base_url: str, match: str | None = None) -> str | None:
    """Ищет ссылку на RSS на странице. match — регулярка по URL/подписи ссылки."""
    soup = BeautifulSoup(html, "lxml")
    candidates: list[tuple[str, str]] = []
    for link in soup.find_all("link", attrs={"rel": True, "href": True}):
        rel = " ".join(link.get("rel") or []).lower()
        typ = str(link.get("type", "")).lower()
        if "alternate" in rel and any(t in typ for t in FEED_TYPES):
            candidates.append((urljoin(base_url, link["href"]), link.get("title") or ""))
    for a in soup.find_all("a", href=True):
        href = str(a["href"])
        if re.search(r"(rss|feed|atom)|\.xml($|\?)", href, re.I):
            candidates.append((urljoin(base_url, href), a.get_text(" ", strip=True)))
    seen = set()
    uniq = []
    for url, label in candidates:
        if url not in seen and url.startswith("http"):
            seen.add(url)
            uniq.append((url, label))
    if match:
        rx = re.compile(match, re.I)
        for url, label in uniq:
            if rx.search(url) or rx.search(label):
                return url
        return None
    return uniq[0][0] if uniq else None


async def collect_rss(src: SourceConfig, state: dict, etag: str | None, last_modified: str | None,
                      fetcher: Fetcher) -> CollectResult:
    result = CollectResult()
    feed_url = src.url
    if src.discover:
        feed_url = state.get("discovered_feed_url") or ""
        if not feed_url:
            page = await fetcher.get(src.url)
            feed_url = discover_feed(page.text, page.url, src.discover_match) or ""
            if not feed_url:
                raise FetchError(f"{src.key}: фид не найден на {src.url} (discover_match={src.discover_match!r})")
            result.state_updates["discovered_feed_url"] = feed_url
            log.info("%s: найден фид %s", src.key, feed_url)
            etag = last_modified = None

    res = await fetcher.get(feed_url, etag=etag, last_modified=last_modified)
    if res.status == 304:
        result.not_modified = True
        result.etag, result.last_modified = etag, last_modified
        return result
    parsed = feedparser.parse(res.content, response_headers={k.lower(): v for k, v in res.headers.items()})
    if parsed.bozo and not parsed.entries:
        raise FetchError(f"{src.key}: битый фид: {parsed.get('bozo_exception')}")
    result.etag = res.headers.get("etag")
    result.last_modified = res.headers.get("last-modified")
    for entry in parsed.entries:
        item = entry_to_item(entry, res.url)
        if item:
            result.items.append(item)
    result.seen = len(parsed.entries)
    return result
