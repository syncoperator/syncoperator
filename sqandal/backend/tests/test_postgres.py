"""Те же сценарии на настоящем PostgreSQL (JSONB, timestamptz, RETURNING, ILIKE, CHECK-ограничения)."""

from __future__ import annotations

import httpx
from sqlalchemy import text

from conftest import PG_URL, pg_only, run
from sqandal.api.app import create_app
from sqandal.bootstrap import build_core
from sqandal.cli import seed_demo
from sqandal.db import Database
from sqandal.demo import DemoBotTransport, DemoClaudeTransport
from sqandal.models import Base

H = {"Authorization": "Bearer test-token"}


async def _reset(settings):
    db = Database(settings)
    async with db.engine.begin() as conn:
        await conn.run_sync(Base.metadata.drop_all)
    await db.dispose()


@pg_only
def test_postgres_end_to_end(settings_factory):
    st = settings_factory(database_url=PG_URL)
    run(_reset(st))
    run(seed_demo(st))

    async def factory(s):
        core = await build_core(s, claude_transport=DemoClaudeTransport(), bot_transport=DemoBotTransport())
        core.newsroom.fetcher = None
        return core

    app = create_app(st, core_factory=factory)

    async def go():
        async with app.router.lifespan_context(app):
            c = httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="http://t")
            stories = (await c.get("/api/stories?status=all", headers=H)).json()["stories"]
            assert len(stories) == 10
            assert len((await c.get("/api/stories?q=пепл", headers=H)).json()["stories"]) == 1  # ILIKE + кириллица
            static = next(s for s in stories if "Lil Static" in s["title"])
            r = await c.post(f"/api/stories/{static['id']}/generate", json={"mode": "hardest"}, headers=H)
            assert r.status_code == 200 and r.json()["version"] == 2
            r = await c.post(f"/api/posts/{r.json()['id']}/publish", json={}, headers=H)
            assert r.json()["status"] == "published"
            stats = (await c.get("/api/stats", headers=H)).json()
            assert stats["published_24h"] >= 2
            await c.aclose()
        db = Database(st)
        async with db.engine.connect() as conn:
            kinds = (await conn.execute(text(
                "select data_type from information_schema.columns where table_name='posts' and column_name='bullets'"
            ))).scalar_one()
            tz = (await conn.execute(text(
                "select data_type from information_schema.columns where table_name='clusters' "
                "and column_name='origin_published_at'"))).scalar_one()
        await db.dispose()
        assert kinds == "jsonb" and tz == "timestamp with time zone"

    run(go())
