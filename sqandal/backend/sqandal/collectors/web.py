"""Web-сборщик: страница-листинг → ссылки на статьи → разбор каждой новой статьи.

Без хрупких CSS-селекторов под каждый сайт: ссылки ищутся по регулярке URL
(`link_pattern` в sources.yaml), а дата/заголовок/текст — по стандартной
разметке статьи (JSON-LD / meta / <time>), см. article.py.
"""

from __future__ import annotations

import logging
import re
import time
from collections.abc import Awaitable, Callable
from urllib.parse import urljoin, urldefrag

from bs4 import BeautifulSoup

from ..textutil import canonicalize_url, collapse_ws
from .article import extract_article
from .base import CollectResult, RawItem, SourceConfig
from .browser import BrowserPool
from .http import FetchError, Fetcher, RobotsDisallowed

log = logging.getLogger(__name__)

# статьи, которые не скачались: не долбим их каждые 15 минут (url → monotonic ts)
_FAILED: dict[str, float] = {}
_FAILED_TTL_S = 6 * 3600


def _recently_failed(url: str) -> bool:
    ts = _FAILED.get(url)
    if ts is None:
        return False
    if time.monotonic() - ts > _FAILED_TTL_S:
        _FAILED.pop(url, None)
        return False
    return True


def extract_links(html: str, base_url: str, pattern: str, limit: int) -> list[tuple[str, str]]:
    """Ссылки на статьи в порядке появления на странице (обычно — свежие сверху)."""
    rx = re.compile(pattern)
    soup = BeautifulSoup(html, "lxml")
    out: list[tuple[str, str]] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        href, _ = urldefrag(urljoin(base_url, str(a["href"]).strip()))
        if not rx.search(href):
            continue
        key = canonicalize_url(href) or href
        if key in seen:
            continue
        seen.add(key)
        out.append((href, collapse_ws(a.get_text(" ", strip=True))))
        if len(out) >= limit:
            break
    return out


async def _get_html(url: str, src: SourceConfig, fetcher: Fetcher, browser: BrowserPool | None,
                    listing: bool = False) -> tuple[str, str]:
    if src.render == "browser" and browser is not None:
        if not await fetcher.allowed(url):
            raise RobotsDisallowed(f"robots.txt запрещает {url}")
        # wait_selector описывает листинг (ссылки на статьи); у статьи его может не быть
        return await browser.render(url, wait_selector=src.wait_selector if listing else None), url
    res = await fetcher.get(url)
    return res.text, res.url


async def collect_web(
    src: SourceConfig,
    fetcher: Fetcher,
    browser: BrowserPool | None,
    is_known: Callable[[str], Awaitable[bool]],
    *,
    max_links: int = 30,
    max_new: int = 12,
) -> CollectResult:
    result = CollectResult()
    html, final_url = await _get_html(src.url, src, fetcher, browser, listing=True)
    links = extract_links(html, final_url, src.link_pattern or "", max_links)
    result.seen = len(links)
    if not links:
        raise FetchError(f"{src.key}: на {src.url} нет ссылок по шаблону {src.link_pattern!r} — сменилась вёрстка?")

    fresh = [
        (u, t) for u, t in links
        if not _recently_failed(u) and not await is_known(canonicalize_url(u) or u)
    ]
    for url, anchor in fresh[:max_new]:
        try:
            if src.fetch_full_text:
                art_html, art_url = await _get_html(url, src, fetcher, browser)
                art = extract_article(art_html, art_url, assume_tz=src.tz)
            else:
                art = None
        except (FetchError, RobotsDisallowed) as e:
            log.warning("%s: статья %s не скачалась: %s", src.key, url, e)
            _FAILED[url] = time.monotonic()
            continue
        except Exception as e:  # noqa: BLE001 — одна битая статья не валит источник
            log.warning("%s: ошибка разбора %s: %s", src.key, url, e)
            _FAILED[url] = time.monotonic()
            continue
        title = (art.title if art else None) or anchor
        if not title:
            continue
        item = RawItem(external_id=url[:512], title=title, url=url)
        if art:
            item.summary = art.description
            item.content = art.text
            item.author = art.author
            item.image_url = art.image_url
            item.lang = art.lang
            item.published_at = art.published_at
            item.date_source = art.date_source
            item.date_has_tz = art.date_has_tz
            if art.canonical and art.canonical != url:
                item.extra["canonical"] = art.canonical
        result.items.append(item)
    return result
