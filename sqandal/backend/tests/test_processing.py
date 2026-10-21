"""Дедуп и верификация: текст/URL, теги, свежесть, кластеризация, пайплайн."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import numpy as np
from sqlalchemy import select

from conftest import BACKEND, run
from sqandal.collectors.base import RawItem, SourceConfig
from sqandal.db import Database
from sqandal.models import (
    Cluster,
    Confidence,
    DateSource,
    NewsItem,
    Priority,
    Scene,
    Verification,
)
from sqandal.processing.clustering import ClusterIndex
from sqandal.processing.embeddings import HashingEmbedder
from sqandal.processing.freshness import assess_staleness, effective_time
from sqandal.processing.ingest import IngestPipeline
from sqandal.processing.tagging import ClusterSignals, Tagger
from sqandal.textutil import canonicalize_url, content_hash, detect_lang
from sqandal.timeutil import utcnow

UTC = timezone.utc
TAGGER = Tagger.from_dir(BACKEND / "config")


# ── текст ───────────────────────────────────────────────────────────
def test_canonicalize_url():
    a = canonicalize_url("http://www.TMZ.com/2026/10/16/slug/?utm_source=x&b=2&a=1#frag")
    assert a == "https://tmz.com/2026/10/16/slug?a=1&b=2"
    assert canonicalize_url("https://m.example.com/x/") == "https://example.com/x"
    assert canonicalize_url(None) is None


def test_detect_lang_and_hash():
    assert detect_lang("Рэпера задержали в Москве") == "ru"
    assert detect_lang("Der Rapper zieht gegen die Behörde vor Gericht") == "de"
    assert detect_lang("The rapper was arrested after the show") == "en"
    assert detect_lang("Рэпер i їхній трек") == "uk"
    assert content_hash("Заголовок!", "Текст") == content_hash("заголовок", "текст")


# ── теги и приоритет ────────────────────────────────────────────────
def test_tags_multilang_and_excludes():
    assert TAGGER.tag("Rapper arrested in Atlanta, booked into Fulton County jail") == ["Задержание"]
    assert "Суд" in TAGGER.tag("Capital Bra zieht gegen Glücksspielbehörde vor Gericht")
    assert TAGGER.tag("Snoop Dogg's Death Row Records signs new artist") == []
    assert "Релиз" not in TAGGER.tag("Рэпера выпустили под залог")
    assert "Суд" in TAGGER.tag("Рэпера выпустили под залог")
    assert TAGGER.tag("Kendrick killed it at the show") == []
    assert "Смерть" in TAGGER.tag("Legendary DJ dies at 68")
    assert "Смерть" in TAGGER.tag("Autopsy shows rapper was fatally shot five times")
    assert "Скандал" in TAGGER.tag("Außerdem stichelte er gegen Kianush und Kolja Goldstein")


def test_release_tag_needs_release_context():
    release = [
        "Лёд Восток выпустил альбом «Сибирь»",
        "Gone.Fludd и Cakeboy выпустили сингл «Каждый день»",
        "Kai Angel заменил минуту тишины в альбоме новым треком",
        "Новый альбом выйдет в декабре",
        "Nebelkrähe veröffentlicht EP Asphaltblumen",
        "Capital Bra kündigt neues Album an",
        "Max B Drops New Album 'Public Domain 7'",
        "Lil Baby Unveils 'Fresh' Tracklist",
        "Drake Reveals 'Iceman' Release Date",
        "Big Toasty announces new album Glass House",
    ]
    not_release = [
        "Мира Кай судится с лейблом: спор — о правах на два альбома",
        "Капитал опубликовал пресс-релиз о концерте",
        "Capital Bra veröffentlicht ein Statement zu den Vorwürfen",
        "Der Rapper erscheint nicht vor Gericht",
        "MC Rook drops diss track aimed at Vex Morgan",
        "Police released footage of the arrest",
        "Rapper released from jail on bond",
        "His 2016 album went platinum",
    ]
    for text in release:
        assert "Релиз" in TAGGER.tag(text), text
    for text in not_release:
        assert "Релиз" not in TAGGER.tag(text), text


def test_artist_matching_declension_and_case():
    names = lambda t: [a.name for a in TAGGER.find_artists(t)]  # noqa: E731
    assert names("Оксимирона вызвали в суд") == ["Oxxxymiron"]
    assert names("Басту госпитализировали") == ["Баста"]
    assert names("the future of rap") == []  # cs: true
    assert names("Future drops album") == ["Future"]
    assert names("A$AP Rocky arrested") == ["A$AP Rocky"]


def test_hiphop_filter():
    assert TAGGER.is_hiphop("Rapper sued over unpaid show")
    assert not TAGGER.is_hiphop("He has a long rap sheet in Georgia")
    assert not TAGGER.is_hiphop("Ex-NFL star to pay $4M in wrongful death case")
    assert TAGGER.is_hiphop("Deutschrapper verhaftet")


def test_priority_scoring():
    sig = dict(artists=[], sources_count=1, first_items_within_window=1, insider_only=False, is_stale=False)
    assert TAGGER.score(ClusterSignals(tags=["Задержание"], **sig))[0] == Priority.CRITICAL
    assert TAGGER.score(ClusterSignals(tags=["Суд"], **sig))[0] == Priority.HIGH
    assert TAGGER.score(ClusterSignals(tags=["Релиз"], **sig))[0] == Priority.NORMAL
    stale = {**sig, "is_stale": True}
    assert TAGGER.score(ClusterSignals(tags=["Задержание"], **stale))[0] == Priority.NORMAL
    big = {**sig, "artists": ["Drake"], "sources_count": 3, "first_items_within_window": 3}
    assert TAGGER.score(ClusterSignals(tags=["Релиз"], **big))[0] == Priority.HIGH


# ── свежесть ────────────────────────────────────────────────────────
def test_effective_time_rules():
    now = datetime(2026, 10, 19, 12, tzinfo=UTC)
    e = effective_time(now - timedelta(hours=1), DateSource.jsonld, now)
    assert e.at == now - timedelta(hours=1) and e.confidence == Confidence.high
    e = effective_time(now + timedelta(hours=3), DateSource.meta, now)  # дата из будущего
    assert e.at == now and e.confidence == Confidence.low and "future_date" in e.flags
    e = effective_time(None, DateSource.fetched, now)
    assert e.at == now and e.confidence == Confidence.low
    e = effective_time(now - timedelta(days=30), DateSource.feed_published, now)
    assert "archive" in e.flags
    e = effective_time(now - timedelta(hours=2), DateSource.meta, now, date_has_tz=False)
    assert e.confidence == Confidence.medium and "tz_guessed" in e.flags
    assert assess_staleness(now - timedelta(hours=30), now, 24)[0] is True
    assert assess_staleness(now - timedelta(hours=3), now, 24) == (False, None)


# ── индекс кластеров ────────────────────────────────────────────────
def _unit(*xs):
    v = np.array(xs, dtype=np.float32)
    return v / np.linalg.norm(v)


def test_cluster_index_artist_conflict_raises_bar():
    idx = ClusterIndex(threshold=0.8, window_hours=72)
    t = datetime(2026, 10, 19, tzinfo=UTC)
    idx.add(1, _unit(1, 0, 0), t, {"Суд"}, {"Drake"})
    v = _unit(1, 0.6, 0)  # cos ≈ 0.857
    assert idx.match(v, t, {"Суд"}, {"Drake"}).cluster_id == 1
    assert idx.match(v, t, {"Суд"}, {"Kanye West"}).cluster_id is None  # другой артист → нужно ≥ 0.90
    assert idx.match(v, t + timedelta(days=5), {"Суд"}, {"Drake"}).cluster_id is None  # вне окна


# ── пайплайн ────────────────────────────────────────────────────────
async def _setup(settings):
    db = Database(settings)
    await db.init()
    from sqandal.bootstrap import sync_sources

    cfgs = [
        SourceConfig(key="wire", name="Wire", kind="rss", scene="us", url="https://wire.test/rss"),
        SourceConfig(key="beat", name="Beat", kind="rss", scene="us", url="https://beat.test/rss"),
        SourceConfig(key="beat_web", name="Beat", kind="web", scene="us", url="https://beat.test/",
                     link_pattern="x"),
        SourceConfig(key="tabloid", name="Tabloid", kind="rss", scene="us", url="https://tab.test/rss",
                     filter="hiphop"),
        SourceConfig(key="tg", name="Insider", kind="telegram", scene="ru", channel="insider"),
    ]
    refs = await sync_sources(db, cfgs)
    pipe = IngestPipeline(db, settings, HashingEmbedder(), TAGGER)
    return db, refs, pipe


def _item(ext, title, text, minutes_ago, url=None, **kw):
    now = utcnow()
    return RawItem(external_id=ext, title=title, summary=text, url=url or f"https://x.test/{ext}",
                   published_at=now - timedelta(minutes=minutes_ago), date_source=DateSource.feed_published,
                   fetched_at=now, **kw)


def test_pipeline_merges_same_story_and_tracks_origin(settings_factory):
    st = settings_factory()

    async def go():
        db, refs, pipe = await _setup(st)
        r1 = await pipe.ingest(refs["wire"], [_item(
            "w1", "Rapper Lil Testy arrested in Atlanta after police chase",
            "Rapper Lil Testy was arrested in Atlanta after a police chase on I-85. Bond was set at $25,000.", 30)])
        r2 = await pipe.ingest(refs["beat"], [_item(
            "b1", "Lil Testy arrested in Atlanta after chase, bond set at $25,000",
            "Atlanta police arrested rapper Lil Testy after a chase on I-85; bond is $25,000.", 50)])
        r3 = await pipe.ingest(refs["wire"], [_item(
            "w2", "Big Toasty announces new album Glass House tracklist",
            "Big Toasty revealed the tracklist for his album Glass House due October 31.", 10)])
        # повтор того же external_id не дублируется
        r4 = await pipe.ingest(refs["wire"], [_item("w1", "dup", "dup", 30)])
        async with db.session() as s:
            clusters = list((await s.execute(select(Cluster).order_by(Cluster.id))).scalars())
        await db.dispose()
        return r1, r2, r3, r4, clusters

    r1, r2, r3, r4, clusters = run(go())
    assert r1.created and r2.updated == r1.created and not r2.created
    assert r3.created and r4.new == 0
    arrest = clusters[0]
    assert arrest.items_count == 2 and arrest.sources_count == 2
    assert arrest.verification == Verification.multi_source
    assert arrest.priority == Priority.CRITICAL and "Задержание" in arrest.tags
    # первоисточник — более ранняя публикация Beat (50 мин назад), хоть и пришла второй
    assert (utcnow() - arrest.origin_published_at) > timedelta(minutes=45)
    assert clusters[1].priority == Priority.NORMAL and clusters[1].tags == ["Релиз"]


def test_same_outlet_two_feeds_is_one_source(settings_factory):
    st = settings_factory()

    async def go():
        db, refs, pipe = await _setup(st)
        await pipe.ingest(refs["beat"], [_item("b1", "Rapper sued by label over unpaid royalties",
                                               "The rapper was sued by his label over royalties.", 20,
                                               url="https://beat.test/a")])
        await pipe.ingest(refs["beat_web"], [_item("https://beat.test/a?utm_source=web", "Rapper sued by label",
                                                   "same article", 19, url="https://beat.test/a?utm_source=web")])
        async with db.session() as s:
            c = (await s.execute(select(Cluster))).scalar_one()
        await db.dispose()
        return c

    c = run(go())
    assert c.items_count == 2 and c.sources_count == 1  # одно издание «Beat»
    assert c.verification == Verification.single_source


def test_hiphop_filter_and_archive_are_not_clustered(settings_factory):
    st = settings_factory()

    async def go():
        db, refs, pipe = await _setup(st)
        rep = await pipe.ingest(refs["tabloid"], [
            _item("t1", "Actor buys a mansion in Malibu", "Hollywood actor buys a house.", 10),
            _item("t2", "Rapper arrested at LAX", "A rapper was arrested at LAX on Sunday.", 10),
            _item("t3", "Rapper sued in old case", "Old story from 2024 about a lawsuit.", 60 * 24 * 40),
        ])
        async with db.session() as s:
            items = {i.external_id: i for i in (await s.execute(select(NewsItem))).scalars()}
        await db.dispose()
        return rep, items

    rep, items = run(go())
    assert rep.new == 3 and rep.relevant == 1 and len(rep.created) == 1
    assert items["t1"].is_relevant is False and items["t1"].extra["skip_reason"] == "not_hiphop"
    assert items["t3"].extra["skip_reason"] == "archive" and items["t3"].cluster_id is None


def test_telegram_link_joins_later_article_and_insider_flag(settings_factory):
    st = settings_factory()

    async def go():
        db, refs, pipe = await _setup(st)
        tg = _item("insider:1", "Рэпера задержали — подробности у Wire", "Рэпера задержали. https://wire.test/story",
                   40, url="https://t.me/insider/1", tg_channel="insider", extra={"links": ["https://wire.test/story"]})
        r1 = await pipe.ingest(refs["tg"], [tg])
        async with db.session() as s:
            c1 = (await s.execute(select(Cluster))).scalar_one()
            v1 = c1.verification
        r2 = await pipe.ingest(refs["wire"], [_item("w9", "Rapper detained by police in Moscow",
                                                    "English text with no shared words", 20,
                                                    url="https://www.wire.test/story/?utm_medium=tg")])
        async with db.session() as s:
            c2 = (await s.execute(select(Cluster))).scalar_one()
        await db.dispose()
        return r1, r2, v1, c2

    r1, r2, v1, c2 = run(go())
    assert v1 == Verification.insider_only
    assert r2.updated == r1.created  # склеено по ссылке, хотя тексты на разных языках
    assert c2.verification == Verification.multi_source and c2.items_count == 2
    assert c2.title == "Rapper detained by police in Moscow"  # заголовок от СМИ, не TG
    assert c2.scene == Scene.ru  # сцена первоисточника


def test_first_run_cap_and_reload_index(settings_factory):
    st = settings_factory(first_run_max_items=2)

    async def go():
        db, refs, pipe = await _setup(st)
        heads = [
            ("Alpha Kid drops neon single Midnight Drive", "Alpha Kid released the single Midnight Drive."),
            ("Bravo Fresh announces European tour dates", "Bravo Fresh will tour Europe in spring."),
            ("Charlie Ray sued by former label", "Charlie Ray faces a lawsuit from his old label."),
            ("Delta Moon hospitalized in Miami", "Delta Moon was hospitalized in Miami on Friday."),
            ("Echo Prince wins best newcomer award", "Echo Prince took home the newcomer award."),
        ]
        items = [_item(f"n{i}", h, t, i * 5) for i, (h, t) in enumerate(heads)]
        rep = await pipe.ingest(refs["wire"], items, first_run=True)
        # новый процесс: индекс поднимается из БД и похожая новость клеится к старому сюжету
        pipe2 = IngestPipeline(db, st, HashingEmbedder(), TAGGER)
        await pipe2.load_index()
        rep2 = await pipe2.ingest(refs["beat"], [_item("x0", "Alpha Kid drops neon single Midnight Drive",
                                                       "Alpha Kid released the single Midnight Drive today.", 1)])
        await db.dispose()
        return rep, rep2, len(pipe2.index)

    rep, rep2, n = run(go())
    assert rep.new == 2  # только 2 самых свежих при первом прогоне
    assert n == 2 and rep2.updated and not rep2.created
