"""Сборщики: RSS/Atom, извлечение статьи и даты, листинги, Telegram-сообщения."""

from __future__ import annotations

from datetime import datetime, timezone

import httpx
import pytest

from conftest import run
from sqandal.collectors.article import extract_article
from sqandal.collectors.base import SourceConfig
from sqandal.collectors.http import Fetcher, RobotsDisallowed
from sqandal.collectors.rss import collect_rss, discover_feed
from sqandal.collectors.telegram import message_to_item
from sqandal.collectors.web import collect_web, extract_links
from sqandal.models import DateSource

UTC = timezone.utc


def mock_fetcher(settings, routes: dict[str, httpx.Response | callable]):
    calls = []

    def handler(request: httpx.Request):
        calls.append(request)
        url = str(request.url)
        for key, resp in routes.items():
            if url == key or url.startswith(key):
                return resp(request) if callable(resp) else resp
        return httpx.Response(404)

    f = Fetcher(settings, transport=httpx.MockTransport(handler))
    f.calls = calls
    return f


# ── RSS ─────────────────────────────────────────────────────────────
def test_rss_parsing_dates_and_cleanup(settings_factory, fixture_text):
    st = settings_factory()
    feed = fixture_text("xxl_feed.xml").encode()
    f = mock_fetcher(st, {"https://feed.test/rss": httpx.Response(
        200, content=feed, headers={"content-type": "application/rss+xml", "etag": '"v1"'})})
    src = SourceConfig(key="xxl", name="XXL", kind="rss", scene="us", url="https://feed.test/rss")
    res = run(collect_rss(src, {}, None, None, f))
    assert res.etag == '"v1"' and len(res.items) == 3
    first = res.items[0]
    assert first.title == "Test Rapper Transferred to New Prison in Pennsylvania"
    assert first.published_at == datetime(2026, 10, 16, 19, 17, tzinfo=UTC)
    assert first.date_source == DateSource.feed_published
    assert first.author == "Staff & Wire"
    assert first.image_url == "https://cdn.example.com/img1.jpg"
    assert "<b>" not in (first.summary or "")
    assert first.external_id == "https://www.xxlmag.com/?p=111"
    assert res.items[1].title == "Here’s How You Can Win a Gift Card"
    no_date = res.items[2]
    assert no_date.published_at is None and no_date.date_source == DateSource.fetched
    run(f.close())


def test_rss_conditional_get_304(settings_factory):
    st = settings_factory()
    seen = {}

    def handler(req):
        seen["inm"] = req.headers.get("if-none-match")
        return httpx.Response(304)

    f = mock_fetcher(st, {"https://feed.test/rss": handler})
    src = SourceConfig(key="xx", name="X", kind="rss", scene="us", url="https://feed.test/rss")
    res = run(collect_rss(src, {}, '"v1"', None, f))
    assert res.not_modified and seen["inm"] == '"v1"' and res.items == []
    run(f.close())


def test_atom_feed_with_timezone(settings_factory, fixture_text):
    st = settings_factory()
    f = mock_fetcher(st, {"https://de.test/atom": httpx.Response(200, content=fixture_text("atom_feed.xml").encode())})
    src = SourceConfig(key="de", name="DE", kind="rss", scene="de", url="https://de.test/atom")
    res = run(collect_rss(src, {}, None, None, f))
    it = res.items[0]
    # updated без published → feed_updated; +02:00 → UTC
    assert it.published_at == datetime(2026, 10, 19, 9, 30, tzinfo=UTC)
    assert it.date_source == DateSource.feed_updated
    run(f.close())


def test_discover_feed(fixture_text):
    html = fixture_text("listing.html")
    assert discover_feed(html, "https://www.tmz.com/hip-hop/") == "https://www.tmz.com/rss/alle.xml"
    assert discover_feed(html, "https://www.tmz.com/", "unterhaltung") == "https://www.bild.de/rssfeeds/unterhaltung-123.xml"
    assert discover_feed(html, "https://www.tmz.com/", "musik") == "https://www.tmz.com/feeds/musik.xml"
    assert discover_feed(html, "https://www.tmz.com/", "nichts-passt") is None


def test_rss_discovery_flow_caches_feed_url(settings_factory, fixture_text):
    st = settings_factory()
    f = mock_fetcher(st, {
        "https://www.bild.test/rss-feeds": httpx.Response(200, text='<a href="/feed/unterhaltung.xml">Unterhaltung</a>'),
        "https://www.bild.test/feed/unterhaltung.xml": httpx.Response(200, content=fixture_text("xxl_feed.xml").encode()),
    })
    src = SourceConfig(key="bild", name="BILD", kind="rss", scene="de", url="https://www.bild.test/rss-feeds",
                       discover=True, discover_match="unterhaltung", filter="hiphop")
    res = run(collect_rss(src, {}, None, None, f))
    assert res.state_updates["discovered_feed_url"] == "https://www.bild.test/feed/unterhaltung.xml"
    assert len(res.items) == 3
    # второй прогон: страницу списка уже не трогаем
    f.calls.clear()
    run(collect_rss(src, res.state_updates, None, None, f))
    assert [str(c.url) for c in f.calls] == ["https://www.bild.test/feed/unterhaltung.xml"]
    run(f.close())


# ── статьи ──────────────────────────────────────────────────────────
def test_article_jsonld_graph_wins(fixture_text):
    art = extract_article(fixture_text("article_jsonld.html"), "https://hiphop.example/magazin/news/x-418229")
    assert art.title == "Capital Testbra zieht gegen Glücksspielbehörde vor Gericht"
    assert art.published_at == datetime(2026, 10, 18, 7, 15, tzinfo=UTC)  # NewsArticle, а не WebPage 2020
    assert art.date_source == DateSource.jsonld and art.date_has_tz
    assert art.image_url == "https://hiphop.example/img/jsonld.jpg"
    assert art.lang == "de"
    assert "Werbeverbot" in art.text and "Navigation" not in art.text and "Weitere News" not in art.text
    assert art.canonical.endswith("-418229")


def test_article_meta_without_timezone(fixture_text):
    art = extract_article(fixture_text("article_meta.html"), "https://the-test.ru/news/x", assume_tz="Europe/Moscow")
    assert art.date_source == DateSource.meta
    assert art.published_at == datetime(2026, 10, 19, 11, 5, tzinfo=UTC)  # 14:05 МСК
    assert art.date_has_tz is False
    assert art.title == "Рэпера задержали в Москве после концерта"
    assert art.lang == "ru"


def test_article_date_fallbacks():
    html_time = "<html><body><article><time datetime='2026-10-17T19:00:00Z'>17 Oct</time><p>" + "x " * 30 + "</p></article></body></html>"
    a = extract_article(html_time, "https://x.test/a")
    assert a.date_source == DateSource.time_tag and a.published_at.day == 17
    a2 = extract_article("<html><body><p>nothing</p></body></html>", "https://www.tmz.com/2026/10/16/slug/")
    assert a2.date_source == DateSource.url and a2.published_at == datetime(2026, 10, 16, 12, tzinfo=UTC)
    a3 = extract_article("<html><body><h1>Title</h1></body></html>", "https://x.test/no-date")
    assert a3.published_at is None and a3.title == "Title"


# ── листинги ────────────────────────────────────────────────────────
def test_extract_links_dedup_and_pattern(fixture_text):
    links = extract_links(fixture_text("listing.html"), "https://www.tmz.com/hip-hop/",
                          r"^https://www\.tmz\.com/20\d{2}/\d{2}/\d{2}/[a-z0-9\-]+/?$", 30)
    urls = [u for u, _ in links]
    assert urls[0] == "https://www.tmz.com/2026/10/17/hot-97-radio-hosts-criticize-leadership-on-air/"
    assert sum("test-rapper-responds" in u for u in urls) == 1  # дубль с utm/#comments склеен
    assert not any("/photos/" in u for u in urls)


def test_collect_web_skips_known_and_parses(settings_factory, fixture_text):
    st = settings_factory()
    art = fixture_text("article_jsonld.html")
    f = mock_fetcher(st, {
        "https://www.tmz.com/hip-hop/": httpx.Response(200, text=fixture_text("listing.html")),
        "https://www.tmz.com/2026/": httpx.Response(200, text=art, headers={"content-type": "text/html; charset=utf-8"}),
    })
    src = SourceConfig(key="tmz", name="TMZ", kind="web", scene="us", url="https://www.tmz.com/hip-hop/",
                       link_pattern=r"^https://www\.tmz\.com/20\d{2}/\d{2}/\d{2}/[a-z0-9\-]+/?$")

    async def is_known(url):
        return "old-known-story" in url

    res = run(collect_web(src, f, None, is_known))
    assert res.seen == 3 and len(res.items) == 2
    assert all(i.date_source == DateSource.jsonld for i in res.items)
    run(f.close())


def test_robots_disallow(settings_factory):
    st = settings_factory(respect_robots=True)
    f = mock_fetcher(st, {
        "https://blocked.test/robots.txt": httpx.Response(200, text="User-agent: *\nDisallow: /private/",
                                                           headers={"content-type": "text/plain"}),
        "https://blocked.test/": httpx.Response(200, text="ok"),
    })
    with pytest.raises(RobotsDisallowed):
        run(f.get("https://blocked.test/private/news"))
    assert run(f.allowed("https://blocked.test/public"))
    run(f.close())


# ── Telegram ────────────────────────────────────────────────────────
def _tg_message(**kw):
    from telethon.tl.patched import Message
    from telethon.tl.types import PeerChannel

    defaults = dict(id=777, peer_id=PeerChannel(1), date=datetime(2026, 10, 19, 15, 0, tzinfo=UTC),
                    message="Рэпера задержали в Казани после концерта. Подробности позже.", views=15000)
    defaults.update(kw)
    return Message(**defaults)


def test_telegram_message_basic():
    item = message_to_item(_tg_message(), "insider_ru")
    assert item.external_id == "insider_ru:777" and item.url == "https://t.me/insider_ru/777"
    assert item.date_source == DateSource.telegram and item.published_at.hour == 15
    assert item.title.startswith("Рэпера задержали") and item.tg_views == 15000


def test_telegram_forward_uses_original_date_and_links():
    from telethon.tl.types import MessageEntityTextUrl, MessageEntityUrl, MessageFwdHeader

    text = "Подтвердили: читать тут https://tmz.test/a и вот ссылка"
    url_off = text.index("https://")
    fwd = MessageFwdHeader(date=datetime(2026, 10, 19, 9, 0, tzinfo=UTC), from_name="Первоисточник")
    msg = _tg_message(message=text, fwd_from=fwd, entities=[
        MessageEntityUrl(offset=url_off, length=len("https://tmz.test/a")),
        MessageEntityTextUrl(offset=text.index("ссылка"), length=6, url="https://xxl.test/b"),
    ])
    item = message_to_item(msg, "insider_ru")
    assert item.date_source == DateSource.telegram_forward
    assert item.published_at == datetime(2026, 10, 19, 9, 0, tzinfo=UTC)  # время ОРИГИНАЛА
    assert item.fwd_from == "Первоисточник"
    assert item.extra["links"] == ["https://tmz.test/a", "https://xxl.test/b"]


def test_telegram_message_without_text_is_skipped():
    assert message_to_item(_tg_message(message=""), "c") is None


def test_fetch_result_decodes_meta_charset_cp1251():
    from sqandal.collectors.http import FetchResult

    body = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1251"></head>' \
           '<body>Рэпера задержали</body></html>'
    r = FetchResult("https://ru.test/", 200, body.encode("cp1251"), httpx.Headers({"content-type": "text/html"}))
    assert "Рэпера задержали" in r.text


def test_article_visible_day_only_date():
    """Hiphop.de / The Flow: в разметке даты нет, в шапке — только день."""
    de = """<html lang="de"><body>
      <div class="article-header"><h1>Ssynics nächste Gegnerin steht fest</h1>
        <div class="article-meta"><span class="author">Leon Schäfers</span><span class="date">20 Oktober, 2026</span></div>
      </div>
      <article><p>Das Battle steigt am 12. Dezember 2026 in der Live Music Hall in Köln, Tickets kosten knapp 60 Euro.</p></article>
    </body></html>"""
    a = extract_article(de, "https://hiphop.de/magazin/news/x-418244", assume_tz="Europe/Berlin")
    assert a.date_source == DateSource.page_day and a.date_has_tz is False
    assert a.published_at == datetime(2026, 10, 20, 10, 0, tzinfo=UTC)  # полдень по Берлину; не 12 декабря из текста

    ru = """<html><body><div class="news-date">19 октября 2026</div>
      <p>Концерт в Москве — 24 декабря 2026.</p></body></html>"""
    a = extract_article(ru, "https://the-flow.ru/news/x", assume_tz="Europe/Moscow")
    assert a.date_source == DateSource.page_day and a.published_at == datetime(2026, 10, 19, 9, 0, tzinfo=UTC)

    en = '<html><body><p class="byline">Published October 20, 2026 12:33 AM PDT | Updated Oct 21, 2026</p></body></html>'
    a = extract_article(en, "https://x.test/a")
    assert a.date_source == DateSource.page_day and a.published_at.date().isoformat() == "2026-10-20"

    # дата только в тексте статьи, без метки — не берём
    a = extract_article("<html><body><article><p>Am 12. Dezember 2026 in Köln.</p></article></body></html>", "https://x.test/b")
    assert a.published_at is None


def test_day_only_effective_time():
    from sqandal.models import Confidence
    from sqandal.processing.freshness import effective_time

    fetched = datetime(2026, 10, 21, 8, 0, tzinfo=UTC)
    e = effective_time(datetime(2026, 10, 19, 10, 0, tzinfo=UTC), DateSource.page_day, fetched, date_has_tz=False)
    assert e.at.day == 19 and e.confidence == Confidence.low and "day_only" in e.flags
    # тот же день, что и обнаружение — точнее время обнаружения
    e = effective_time(datetime(2026, 10, 21, 10, 0, tzinfo=UTC), DateSource.page_day, fetched, date_has_tz=False)
    assert e.at == fetched and e.flags == ["day_only"]
