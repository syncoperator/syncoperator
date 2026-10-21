"""CLI: python -m sqandal.cli <команда>

  init-db                 создать таблицы
  dump-schema [dialect]   DDL схемы (postgresql | sqlite)
  check-sources           проверить sources.yaml и показать источники
  collect [--source KEY]  один прогон сбора (без LLM), печать итогов
  generate ID [--mode M]  сгенерировать пост для сюжета и напечатать
  tg-login                вход в Telegram (Telethon) → TG_SESSION
  set-webhook             зарегистрировать вебхук Bot API (кнопки админ-чата)
  demo [--export FILE]    демо-данные (вымышленные) без сети и ключей
  snapshot DIR --export F снимок реальных новостей для демо (items.json + posts.json)
  calibrate               распределение близостей для подбора CLUSTER_THRESHOLD
"""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import sys
from datetime import timedelta

from sqlalchemy import select
from sqlalchemy.schema import CreateIndex, CreateTable

from .bootstrap import build_core, load_and_validate_sources, sync_sources
from .settings import get_settings
from .timeutil import utcnow


def cmd_dump_schema(dialect_name: str) -> None:
    from sqlalchemy.dialects import postgresql, sqlite

    from .models import Base

    dialect = postgresql.dialect() if dialect_name.startswith("post") else sqlite.dialect()
    print(f"-- SQANDAL RADAR schema ({dialect_name}). Сгенерировано из sqandal/models.py")
    for table in Base.metadata.sorted_tables:
        print(str(CreateTable(table).compile(dialect=dialect)).strip() + ";\n")
        for idx in sorted(table.indexes, key=lambda i: i.name):
            print(str(CreateIndex(idx).compile(dialect=dialect)).strip() + ";")
        if dialect_name.startswith("post"):
            def q(v: str) -> str:
                return "'" + v.replace("'", "''") + "'"

            if table.comment:
                print(f"COMMENT ON TABLE {table.name} IS {q(table.comment)};")
            for col in table.columns:
                if col.comment:
                    print(f"COMMENT ON COLUMN {table.name}.{col.name} IS {q(col.comment)};")
        print()


async def cmd_check_sources() -> None:
    st = get_settings()
    for s in load_and_validate_sources(st):
        flag = "ON " if s.enabled else "off"
        print(f"[{flag}] {s.key:16} {s.kind.value:9} {s.scene.value:3} {s.url}")
        if s.note:
            print(f"        └ {s.note}")


async def cmd_collect(source: str | None) -> None:
    from .worker import Worker

    st = get_settings()
    core = await build_core(st)
    w = Worker(core)
    await w.load_sources()
    await w.pipeline.load_index()
    refs = [r for k, r in w.refs.items() if (source is None or k == source) and r.cfg.enabled
            and r.cfg.kind.value != "telegram"]
    reports = await asyncio.gather(*(w.collect_source(r) for r in refs))
    for r, rep in zip(refs, reports):
        print(f"{r.key:16} " + (f"новых {rep.new:3}  релевантных {rep.relevant:3}  сюжетов +{len(rep.created)}"
                                 if rep else "ошибка — см. лог / sources.last_error"))
    if w.browser:
        await w.browser.close()
    await core.close()


async def cmd_generate(cluster_id: int, mode: str) -> None:
    core = await build_core(get_settings())
    try:
        p = await core.newsroom.generate_post(cluster_id, mode)
        print(p.body_text)
        if p.warnings:
            print("\n⚠️ " + "\n⚠️ ".join(p.warnings))
        print(f"\n[пост #{p.id} v{p.version}, {p.llm_model}, in={p.input_tokens} out={p.output_tokens} "
              f"cache={p.cache_read_tokens}]")
    finally:
        await core.close()


async def cmd_tg_login() -> None:
    from .collectors.telegram import interactive_login

    st = get_settings()
    if not (st.tg_api_id and st.tg_api_hash):
        sys.exit("Задай TG_API_ID и TG_API_HASH (https://my.telegram.org → API development tools)")
    session = await interactive_login(int(st.tg_api_id), st.tg_api_hash)
    print("\nДобавь в .env (это ключ к аккаунту — никому не показывай):\n")
    print(f"TG_SESSION={session}")


async def cmd_set_webhook() -> None:
    st = get_settings()
    if not (st.tg_bot_token and st.public_base_url and st.tg_webhook_secret):
        sys.exit("Нужны TG_BOT_TOKEN, PUBLIC_BASE_URL (https://…) и TG_WEBHOOK_SECRET")
    from .publish.telegram_bot import TelegramBot

    bot = TelegramBot(st.tg_bot_token)
    url = st.public_base_url.rstrip("/") + "/api/telegram/webhook"
    print(await bot.set_webhook(url, st.tg_webhook_secret), url)
    await bot.close()


async def seed_demo(settings, export: str | None = None) -> dict:
    """Наполнить БД демо-сюжетами через настоящий пайплайн. Возвращает экспорт для фронта."""
    from .demo import DEMO_STORIES, DemoBotTransport, DemoClaudeTransport, demo_raw_items, demo_source_configs
    from .demo import story_markers
    from .models import NewsItem, Post
    from .processing.embeddings import HashingEmbedder
    from .processing.ingest import IngestPipeline

    settings = settings.model_copy(update={
        "anthropic_api_key": settings.anthropic_api_key or "demo",
        "tg_bot_token": settings.tg_bot_token or "demo",
        "embedding_backend": "hashing",
        "post_link_preview": False,
    })
    bot_tr = DemoBotTransport()
    core = await build_core(settings, claude_transport=DemoClaudeTransport(), bot_transport=bot_tr)
    core.newsroom.fetcher = None  # демо-ссылки ведут в никуда — не докачиваем статьи
    refs = await sync_sources(core.db, demo_source_configs())
    pipe = IngestPipeline(core.db, settings, HashingEmbedder(), core.tagger)
    now = utcnow()
    for key, raw in demo_raw_items(now):
        await pipe.ingest(refs[key], [raw])

    async with core.db.session() as s:
        rows = list((await s.execute(
            select(NewsItem.url, NewsItem.external_id, NewsItem.cluster_id).where(NewsItem.cluster_id.is_not(None))
        )).all())
    by_marker = {}
    for url, ext, cid in rows:
        by_marker[url] = cid
        by_marker[ext] = cid

    def cluster_of(story) -> int | None:
        for m in story_markers(DEMO_STORIES.index(story)):
            if m in by_marker:
                return by_marker[m]
        return None

    for story in DEMO_STORIES:
        cid = cluster_of(story)
        if cid is None:
            continue
        status = story.get("status", "draft")
        if status == "dismissed":
            await core.newsroom.set_dismissed(cid, True)
            continue
        if status == "none":
            continue
        post = await core.newsroom.generate_post(cid)
        if status == "scheduled":
            await core.newsroom.schedule_post(post.id, now + timedelta(minutes=70))
        elif status == "published":
            await core.newsroom.publish_post(post.id, confirm_stale=True)
            async with core.db.session() as s:  # «опубликовано 4 часа назад»
                p = await s.get(Post, post.id)
                p.published_at = now - timedelta(hours=4)
                await s.commit()

    result = {"stories": len(DEMO_STORIES)}
    if export:
        harder = {}
        for story in DEMO_STORIES:
            cid = cluster_of(story)
            if cid and story.get("harder"):
                harder[cid] = story["harder"]
        data = await export_dashboard(core, settings, source_filter=lambda key: key.startswith("demo_"))
        data.update({"generated_at": now.isoformat(), "harder": harder})
        with open(export, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1, default=str)
        result["export"] = export
    else:
        await core.close()
    return result


async def export_dashboard(core, settings, source_filter=None) -> dict:
    """Всё, что видит дашборд (через настоящий API), одним JSON. Закрывает core."""
    from httpx import ASGITransport, AsyncClient

    from .api.app import create_app

    app = create_app(settings, core_factory=lambda st: _reuse(core))
    async with app.router.lifespan_context(app):
        async with AsyncClient(transport=ASGITransport(app=app), base_url="http://export") as c:
            h = {"Authorization": f"Bearer {settings.dashboard_token}"}
            stories = (await c.get("/api/stories?status=all&hours=720&limit=300", headers=h)).json()["stories"]
            details = {}
            for st_ in stories:
                details[st_["id"]] = (await c.get(f"/api/stories/{st_['id']}", headers=h)).json()
            sources = (await c.get("/api/sources", headers=h)).json()
            return {
                "config": (await c.get("/api/config", headers=h)).json(),
                "stats": (await c.get("/api/stats", headers=h)).json(),
                "sources": [x for x in sources if source_filter is None or source_filter(x["key"])],
                "stories": stories,
                "details": details,
            }


async def _reuse(core):
    return _KeepOpen(core)


class _KeepOpen:
    """Обёртка: API-lifespan не закрывает общий core демо-сидера раньше времени."""

    def __init__(self, core):
        self._core = core

    def __getattr__(self, item):
        return getattr(self._core, item)

    async def close(self):
        await self._core.close()


async def cmd_calibrate(limit: int) -> None:
    import numpy as np

    from .models import NewsItem
    from .processing.embeddings import from_bytes, make_embedder

    st = get_settings()
    core = await build_core(st)
    emb = make_embedder(st)
    async with core.db.session() as s:
        rows = list((await s.execute(
            select(NewsItem.id, NewsItem.title, NewsItem.embedding, NewsItem.cluster_id)
            .where(NewsItem.embedding_model == emb.model_id).order_by(NewsItem.id.desc()).limit(limit)
        )).all())
    if len(rows) < 3:
        sys.exit("Мало данных: дай воркеру собрать новости и повтори")
    vecs = np.vstack([from_bytes(r[2]) for r in rows])
    sims = vecs @ vecs.T
    same, diff = [], []
    pairs = []
    for i in range(len(rows)):
        for j in range(i + 1, len(rows)):
            (same if rows[i][3] == rows[j][3] else diff).append(sims[i, j])
            pairs.append((sims[i, j], rows[i][1], rows[j][1], rows[i][3] == rows[j][3]))
    print(f"Модель: {emb.model_id}; пар внутри сюжетов: {len(same)}, между сюжетами: {len(diff)}")
    if same:
        print(f"  внутри:  min {min(same):.3f}  медиана {np.median(same):.3f}")
    print(f"  между:   max {max(diff):.3f}  p99 {np.percentile(diff, 99):.3f}")
    print("\nПограничные пары (близость 0.55–0.85), проверь глазами:")
    for sim, a, b, same_c in sorted([p for p in pairs if 0.55 <= p[0] <= 0.85], key=lambda x: -x[0])[:25]:
        print(f"  {sim:.3f} {'=' if same_c else '≠'}  {a[:60]}  |  {b[:60]}")
    await core.close()


def main(argv: list[str] | None = None) -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
    ap = argparse.ArgumentParser(prog="sqandal", description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("init-db")
    p = sub.add_parser("dump-schema")
    p.add_argument("dialect", nargs="?", default="postgresql", choices=["postgresql", "sqlite"])
    sub.add_parser("check-sources")
    p = sub.add_parser("collect")
    p.add_argument("--source")
    p = sub.add_parser("generate")
    p.add_argument("cluster_id", type=int)
    p.add_argument("--mode", default="standard", choices=["standard", "harder", "hardest"])
    sub.add_parser("tg-login")
    sub.add_parser("set-webhook")
    p = sub.add_parser("demo")
    p.add_argument("--export")
    p.add_argument("--db", default="sqlite+aiosqlite:///./data/demo.db",
                   help="отдельная БД для демо (по умолчанию data/demo.db — рабочую не трогаем)")
    p = sub.add_parser("snapshot")
    p.add_argument("folder", help="папка с items.json и posts.json (например snapshot/2026-10-21)")
    p.add_argument("--export", required=True)
    p.add_argument("--db", default="sqlite+aiosqlite:///./data/snapshot.db")
    p = sub.add_parser("calibrate")
    p.add_argument("--limit", type=int, default=400)
    a = ap.parse_args(argv)

    if a.cmd == "dump-schema":
        cmd_dump_schema(a.dialect)
    elif a.cmd == "init-db":
        async def _init():
            core = await build_core(get_settings())
            await core.close()
            print("ok")
        asyncio.run(_init())
    elif a.cmd == "check-sources":
        asyncio.run(cmd_check_sources())
    elif a.cmd == "collect":
        asyncio.run(cmd_collect(a.source))
    elif a.cmd == "generate":
        asyncio.run(cmd_generate(a.cluster_id, a.mode))
    elif a.cmd == "tg-login":
        asyncio.run(cmd_tg_login())
    elif a.cmd == "set-webhook":
        asyncio.run(cmd_set_webhook())
    elif a.cmd == "demo":
        st = get_settings().model_copy(update={"database_url": a.db})
        print(asyncio.run(seed_demo(st, a.export)))
    elif a.cmd == "snapshot":
        from pathlib import Path

        from .snapshot import build_snapshot

        db_path = a.db.split("///", 1)[-1]
        if a.db.startswith("sqlite") and Path(db_path).exists():
            Path(db_path).unlink()  # снимок всегда с чистой базы
        st = get_settings().model_copy(update={"database_url": a.db})
        print(asyncio.run(build_snapshot(st, Path(a.folder), Path(a.export))))
    elif a.cmd == "calibrate":
        asyncio.run(cmd_calibrate(a.limit))


if __name__ == "__main__":
    main()
