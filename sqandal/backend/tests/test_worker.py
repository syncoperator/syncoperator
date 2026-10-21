"""Воркер: прогон источника → БД, здоровье источников, бэкофф, авто-генерация, отложенный постинг."""

from __future__ import annotations

import shutil
from datetime import timedelta

import httpx
from sqlalchemy import select

from conftest import BACKEND, run
from sqandal.bootstrap import build_core
from sqandal.demo import DemoBotTransport, DemoClaudeTransport
from sqandal.models import Cluster, FetchRun, Post, PostStatus, Source
from sqandal.processing.embeddings import HashingEmbedder
from sqandal.timeutil import utcnow
from sqandal.worker import Worker

FEED = """<?xml version="1.0"?><rss version="2.0"><channel><title>T</title>
<item><title>Rapper Lil Testy arrested in Atlanta</title><link>https://feed.test/a1</link>
<pubDate>{d1}</pubDate><description>Rapper Lil Testy was arrested in Atlanta on Sunday.</description></item>
<item><title>Actor buys yacht</title><link>https://feed.test/a2</link><pubDate>{d1}</pubDate>
<description>A movie star bought a yacht in Monaco.</description></item>
</channel></rss>"""


def _config_dir(tmp_path, sources_yaml: str):
    d = tmp_path / "config"
    d.mkdir()
    shutil.copy(BACKEND / "config" / "tags.yaml", d / "tags.yaml")
    shutil.copy(BACKEND / "config" / "artists.yaml", d / "artists.yaml")
    (d / "sources.yaml").write_text(sources_yaml, encoding="utf-8")
    return d


def test_worker_collects_ingests_and_queues_generation(settings_factory, tmp_path):
    cfg = _config_dir(tmp_path, """
sources:
  - {key: tabloid, name: Tabloid, kind: rss, scene: us, url: "https://feed.test/rss", filter: hiphop}
  - {key: broken, name: Broken, kind: rss, scene: de, url: "https://broken.test/rss"}
""")
    st = settings_factory(config_dir=cfg)
    d1 = (utcnow() - timedelta(minutes=20)).strftime("%a, %d %b %Y %H:%M:%S +0000")

    def handler(req: httpx.Request):
        if req.url.host == "feed.test":
            return httpx.Response(200, content=FEED.format(d1=d1).encode())
        return httpx.Response(500)

    async def go():
        core = await build_core(st, claude_transport=DemoClaudeTransport(), bot_transport=DemoBotTransport(),
                                http_transport=httpx.MockTransport(handler))
        w = Worker(core, embedder=HashingEmbedder())
        await w.load_sources()
        rep = await w.collect_source(w.refs["tabloid"])
        bad = await w.collect_source(w.refs["broken"])
        async with core.db.session() as s:
            srcs = {x.key: x for x in (await s.execute(select(Source))).scalars()}
            runs = list((await s.execute(select(FetchRun))).scalars())
        queued = [w.gen_queue.get_nowait() for _ in range(w.gen_queue.qsize())]
        # бэкофф: после ошибки источник не дёргается через 15 минут
        due_ok = w._due(w.refs["tabloid"], utcnow() - timedelta(minutes=16), 0)
        due_bad = w._due(w.refs["broken"], utcnow() - timedelta(minutes=16), 1)
        await core.close()
        return rep, bad, srcs, runs, queued, due_ok, due_bad

    rep, bad, srcs, runs, queued, due_ok, due_bad = run(go())
    assert rep.new == 2 and rep.relevant == 1 and len(rep.created) == 1
    assert bad is None
    assert srcs["tabloid"].last_success_at and srcs["tabloid"].consecutive_failures == 0
    assert srcs["broken"].consecutive_failures == 1 and "500" in srcs["broken"].last_error
    assert {r.ok for r in runs} == {True, False}
    assert queued == rep.created  # CRITICAL-сюжет «arrested» ушёл в очередь авто-генерации
    assert due_ok is True and due_bad is False


def test_scheduled_post_is_published_when_due(settings_factory):
    from sqandal.cli import seed_demo

    st = settings_factory()
    run(seed_demo(st))

    async def go():
        bot = DemoBotTransport()
        core = await build_core(st, claude_transport=DemoClaudeTransport(), bot_transport=bot)
        async with core.db.session() as s:
            p = (await s.execute(select(Post).where(Post.status == PostStatus.scheduled))).scalar_one()
            pid = p.id
            p.scheduled_at = utcnow() - timedelta(seconds=5)  # время пришло
            await s.commit()
        due = await core.newsroom.due_posts()
        for x in due:
            await core.newsroom.publish_post(x, confirm_stale=True)
        async with core.db.session() as s:
            p = await s.get(Post, pid)
            c = await s.get(Cluster, p.cluster_id)
        await core.close()
        return due, pid, p, c, bot

    due, pid, p, c, bot = run(go())
    assert due == [pid] and p.status == PostStatus.published and c.status.value == "published"
    assert any(x["method"] == "sendMessage" for x in bot.sent)
