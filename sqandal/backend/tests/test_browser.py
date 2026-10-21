"""Playwright: страница, которую рисует JavaScript (как Tilda/SPA), парсится через render: browser."""

from __future__ import annotations

import http.server
import os
import threading
from functools import partial

import pytest

from conftest import run
from sqandal.collectors.base import SourceConfig
from sqandal.collectors.browser import BrowserPool
from sqandal.collectors.http import Fetcher
from sqandal.collectors.web import collect_web
from sqandal.models import DateSource

LISTING = """<!doctype html><html><head><meta charset="utf-8"></head><body><div id="feed"></div>
<script>
  setTimeout(() => {
    document.getElementById('feed').innerHTML =
      '<a href="/news/js-story-one">Один</a><a href="/news/js-story-two">Два</a><a href="/about">о нас</a>';
  }, 150);
</script></body></html>"""

ARTICLE = """<!doctype html><html lang="ru"><head><meta charset="utf-8">
<script type="application/ld+json">{"@type":"NewsArticle","headline":"%s","datePublished":"2026-10-19T12:00:00+03:00"}</script>
</head><body><article><p>Текст статьи, собранной скриптом, достаточно длинный для абзаца.</p></article></body></html>"""


def _browsers_available() -> bool:
    try:
        import playwright  # noqa: F401
    except ImportError:
        return False
    path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", os.path.expanduser("~/.cache/ms-playwright"))
    return os.path.isdir(path) and any(n.startswith("chromium") for n in os.listdir(path))


@pytest.mark.skipif(not _browsers_available(), reason="Chromium для Playwright не установлен")
def test_browser_render_listing_and_articles(settings_factory, tmp_path):
    (tmp_path / "index.html").write_text(LISTING, encoding="utf-8")
    news = tmp_path / "news"
    news.mkdir()
    for slug, title in (("js-story-one", "Первая JS-новость"), ("js-story-two", "Вторая JS-новость")):
        (news / slug).write_text(ARTICLE % title, encoding="utf-8")

    handler = partial(http.server.SimpleHTTPRequestHandler, directory=str(tmp_path))

    class Quiet(handler.func):  # type: ignore[misc]
        def log_message(self, *a):
            pass

        def guess_type(self, path):
            return "text/html"

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), partial(Quiet, directory=str(tmp_path)))
    port = srv.server_address[1]
    threading.Thread(target=srv.serve_forever, daemon=True).start()

    st = settings_factory(browser_enabled=True)
    src = SourceConfig(key="jsite", name="JS Site", kind="web", scene="ru", url=f"http://127.0.0.1:{port}/",
                       link_pattern=r"/news/[a-z\-]+$", render="browser", wait_selector="a[href^='/news/']")

    async def go():
        f = Fetcher(st)
        b = BrowserPool(st)
        try:
            async def unknown(_):
                return False

            return await collect_web(src, f, b, unknown)
        finally:
            await b.close()
            await f.close()

    try:
        res = run(go())
    finally:
        srv.shutdown()
    titles = sorted(i.title for i in res.items)
    assert titles == ["Вторая JS-новость", "Первая JS-новость"]
    assert all(i.date_source == DateSource.jsonld and i.published_at.hour == 9 for i in res.items)
