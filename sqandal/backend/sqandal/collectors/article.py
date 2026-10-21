"""Извлечение статьи из HTML: заголовок, ДАТА ПЕРВОИСТОЧНИКА, текст, картинка.

Порядок доверия к дате (от высокого к низкому):
  JSON-LD datePublished → meta article:published_time и аналоги → <time datetime>
  → дата в URL (/2026/10/16/) → нет даты (время обнаружения).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

from bs4 import BeautifulSoup, Tag

from ..models import DateSource
from ..textutil import collapse_ws, detect_lang
from ..timeutil import UTC, date_from_url, parse_datetime

ARTICLE_TYPES = {
    "newsarticle", "article", "blogposting", "reportagenewsarticle",
    "analysisnewsarticle", "backgroundnewsarticle", "socialmediaposting", "webpage",
}

META_DATE_KEYS = [
    ("property", "article:published_time"),
    ("property", "og:article:published_time"),
    ("name", "article:published_time"),
    ("itemprop", "datePublished"),
    ("name", "pubdate"),
    ("name", "publishdate"),
    ("name", "publish-date"),
    ("name", "parsely-pub-date"),
    ("name", "sailthru.date"),
    ("name", "dc.date.issued"),
    ("name", "DC.date.issued"),
    ("name", "date"),
    ("property", "og:updated_time"),  # последний шанс — это дата обновления
]

_BLOCK_JUNK = ["script", "style", "noscript", "iframe", "form", "nav", "aside", "footer", "header", "figure", "button"]
_BODY_SELECTORS = [
    "[itemprop=articleBody]",
    "article",
    ".article-body", ".article__body", ".article-content", ".entry-content", ".post-content",
    ".post__text", ".news-text", ".content-body", ".story-body", "main",
]


@dataclass(slots=True)
class ArticleData:
    title: str | None = None
    description: str | None = None
    text: str | None = None
    published_at: datetime | None = None
    date_source: DateSource = DateSource.fetched
    date_has_tz: bool = False
    modified_at: datetime | None = None
    author: str | None = None
    image_url: str | None = None
    lang: str | None = None
    canonical: str | None = None


def _iter_jsonld(obj: Any):
    if isinstance(obj, list):
        for x in obj:
            yield from _iter_jsonld(x)
    elif isinstance(obj, dict):
        if "@graph" in obj:
            yield from _iter_jsonld(obj["@graph"])
        yield obj


def _types(obj: dict) -> set[str]:
    t = obj.get("@type", [])
    if isinstance(t, str):
        t = [t]
    return {str(x).lower() for x in t}


def _first_str(v: Any) -> str | None:
    if v is None:
        return None
    if isinstance(v, str):
        return v
    if isinstance(v, list) and v:
        return _first_str(v[0])
    if isinstance(v, dict):
        return _first_str(v.get("name") or v.get("url") or v.get("@id"))
    return str(v)


def _parse_jsonld(soup: BeautifulSoup) -> dict | None:
    best = None
    for script in soup.find_all("script", attrs={"type": re.compile(r"ld\+json", re.I)}):
        raw = script.string or script.get_text() or ""
        raw = raw.strip()
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            try:  # частая поломка: управляющие символы внутри строк
                data = json.loads(re.sub(r"[\x00-\x1f]", " ", raw))
            except json.JSONDecodeError:
                continue
        for obj in _iter_jsonld(data):
            types = _types(obj)
            if types & ARTICLE_TYPES and ("datePublished" in obj or "headline" in obj):
                # статья важнее WebPage
                if best is None or ("webpage" in _types(best) and "webpage" not in types):
                    best = obj
    return best


def _meta(soup: BeautifulSoup, attr: str, key: str) -> str | None:
    tag = soup.find("meta", attrs={attr: key})
    if tag and tag.get("content"):
        return str(tag["content"]).strip()
    return None


def _main_text(soup: BeautifulSoup) -> str | None:
    for junk in soup(_BLOCK_JUNK):
        junk.decompose()
    best_text, best_len = None, 0
    candidates: list[Tag] = []
    for sel in _BODY_SELECTORS:
        candidates.extend(soup.select(sel))
    if not candidates and soup.body:
        candidates = [soup.body]
    for c in candidates:
        paras = [collapse_ws(p.get_text(" ", strip=True)) for p in c.find_all(["p", "li", "h2", "blockquote"])]
        paras = [p for p in paras if len(p) >= 30]
        text = "\n".join(paras)
        if len(text) > best_len:
            best_text, best_len = text, len(text)
    return best_text or None


_MONTHS = {
    # ru (родительный падеж — с числом всегда он)
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
    "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
    # de
    "januar": 1, "jänner": 1, "februar": 2, "märz": 3, "maerz": 3, "april": 4, "mai": 5, "juni": 6,
    "juli": 7, "august": 8, "september": 9, "oktober": 10, "okt": 10, "november": 11, "dezember": 12, "dez": 12,
    # en
    "january": 1, "february": 2, "march": 3, "may": 5, "june": 6, "july": 7, "october": 10, "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "sept": 9,
    "oct": 10, "nov": 11, "dec": 12,
}
_MONTH_RX = "|".join(sorted(_MONTHS, key=len, reverse=True))
_DAY_MONTH_YEAR = re.compile(rf"(?<!\d)(\d{{1,2}})\.?\s+({_MONTH_RX})\.?,?\s+(20\d{{2}})(?!\d)", re.I)
_MONTH_DAY_YEAR = re.compile(rf"(?<![\w])({_MONTH_RX})\.?\s+(\d{{1,2}}),?\s+(20\d{{2}})(?!\d)", re.I)
_NUMERIC_DATE = re.compile(r"(?<![\d.])(\d{1,2})\.(\d{1,2})\.(20\d{2})(?![\d])")
_DATE_HINT = re.compile(r"date|time|publish|posted|byline|datum|meta", re.I)


def _parse_day(text: str) -> tuple[int, int, int] | None:
    for rx, order in ((_DAY_MONTH_YEAR, "dmy"), (_MONTH_DAY_YEAR, "mdy"), (_NUMERIC_DATE, "num")):
        m = rx.search(text)
        if not m:
            continue
        if order == "dmy":
            d, mon, y = int(m.group(1)), _MONTHS.get(m.group(2).lower()), int(m.group(3))
        elif order == "mdy":
            mon, d, y = _MONTHS.get(m.group(1).lower()), int(m.group(2)), int(m.group(3))
        else:
            d, mon, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if mon:
            try:
                datetime(y, mon, d)
            except ValueError:
                continue
            return y, mon, d
    return None


def visible_day(soup: BeautifulSoup, assume_tz: str | None = None) -> datetime | None:
    """Дата публикации из видимой шапки статьи (элементы с date/time/byline/meta в class/id/itemprop).

    Только день: возвращает полдень по часовому поясу источника, с пометкой «только день»
    дальше по пайплайну. Даты из текста статьи («концерт 12 декабря») не берём —
    только короткие элементы-метки.
    """
    for el in soup.find_all(True):
        hint = " ".join([*(el.get("class") or []), str(el.get("id") or ""), str(el.get("itemprop") or "")])
        if not hint.strip() or not _DATE_HINT.search(hint):
            continue
        text = collapse_ws(el.get_text(" ", strip=True))
        if not (6 <= len(text) <= 90):
            continue
        ymd = _parse_day(text)
        if ymd:
            tz = ZoneInfo(assume_tz) if assume_tz else UTC
            return datetime(*ymd, 12, 0, tzinfo=tz).astimezone(UTC)
    return None


def extract_article(html: str, url: str, assume_tz: str | None = None) -> ArticleData:
    soup = BeautifulSoup(html, "lxml")
    out = ArticleData()

    html_tag = soup.find("html")
    if html_tag and html_tag.get("lang"):
        out.lang = str(html_tag["lang"]).split("-")[0].lower()[:8]

    canon = soup.find("link", rel=lambda v: v and "canonical" in v)
    if canon and canon.get("href"):
        out.canonical = urljoin(url, str(canon["href"]))

    ld = _parse_jsonld(soup)
    if ld:
        out.title = collapse_ws(_first_str(ld.get("headline")) or "") or None
        out.description = _first_str(ld.get("description"))
        body = ld.get("articleBody")
        if isinstance(body, str) and len(body) > 80:
            out.text = collapse_ws(body)
        out.author = _first_str(ld.get("author"))
        img = ld.get("image")
        out.image_url = _first_str(img.get("url") if isinstance(img, dict) else img)
        dt, has_tz = parse_datetime(_first_str(ld.get("datePublished")), assume_tz)
        if dt:
            out.published_at, out.date_source, out.date_has_tz = dt, DateSource.jsonld, has_tz
        out.modified_at, _ = parse_datetime(_first_str(ld.get("dateModified")), assume_tz)

    if out.published_at is None:
        for attr, key in META_DATE_KEYS:
            dt, has_tz = parse_datetime(_meta(soup, attr, key), assume_tz)
            if dt:
                out.published_at, out.date_source, out.date_has_tz = dt, DateSource.meta, has_tz
                break

    if out.published_at is None:
        scope = soup.find("article") or soup
        t = scope.find("time", attrs={"datetime": True})
        if t:
            dt, has_tz = parse_datetime(str(t["datetime"]), assume_tz)
            if dt:
                out.published_at, out.date_source, out.date_has_tz = dt, DateSource.time_tag, has_tz

    if out.published_at is None:
        # Hiphop.de, The Flow: в разметке даты нет, в шапке — только день («20 Oktober, 2026»).
        # Без этого первый прогон выдал бы новости трёхдневной давности за свежие.
        dt = visible_day(soup, assume_tz)
        if dt:
            out.published_at, out.date_source, out.date_has_tz = dt, DateSource.page_day, False

    if out.published_at is None:
        dt = date_from_url(url)
        if dt:
            out.published_at, out.date_source, out.date_has_tz = dt, DateSource.url, False

    if not out.title:
        out.title = (
            _meta(soup, "property", "og:title")
            or _meta(soup, "name", "twitter:title")
            or (collapse_ws(soup.h1.get_text(" ", strip=True)) if soup.h1 else None)
            or (collapse_ws(soup.title.get_text()) if soup.title else None)
        )
    out.description = out.description or _meta(soup, "property", "og:description") or _meta(soup, "name", "description")
    if not out.image_url:
        out.image_url = _meta(soup, "property", "og:image")
    if out.image_url:
        out.image_url = urljoin(url, out.image_url)

    if not out.text:
        out.text = _main_text(soup)
    if not out.lang:
        out.lang = detect_lang(" ".join(x for x in (out.title, out.text) if x))
    return out
