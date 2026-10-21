"""Async-движок SQLAlchemy: PostgreSQL (psycopg 3) или SQLite (aiosqlite)."""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from sqlalchemy import event, select
from sqlalchemy.ext.asyncio import AsyncEngine, AsyncSession, async_sessionmaker, create_async_engine

from .models import AppState, Base
from .settings import Settings
from .timeutil import utcnow


def make_engine(settings: Settings) -> AsyncEngine:
    url = settings.database_url
    kwargs: dict[str, Any] = {"echo": settings.db_echo, "pool_pre_ping": True}
    if url.startswith("sqlite"):
        # каталог под файл БД
        path = url.split("///", 1)[-1]
        if path and path != ":memory:":
            Path(path).parent.mkdir(parents=True, exist_ok=True)
        kwargs["connect_args"] = {"timeout": 30}
    else:
        kwargs.update(pool_size=10, max_overflow=10)
    engine = create_async_engine(url, **kwargs)

    if url.startswith("sqlite"):

        @event.listens_for(engine.sync_engine, "connect")
        def _sqlite_pragmas(dbapi_conn, _record):  # pragma: no cover - инфраструктура
            cur = dbapi_conn.cursor()
            cur.execute("PRAGMA journal_mode=WAL")  # API и воркер пишут одновременно
            cur.execute("PRAGMA foreign_keys=ON")
            cur.execute("PRAGMA busy_timeout=10000")
            cur.execute("PRAGMA synchronous=NORMAL")
            cur.close()

    return engine


def make_session_factory(engine: AsyncEngine) -> async_sessionmaker[AsyncSession]:
    return async_sessionmaker(engine, expire_on_commit=False, autoflush=False)


async def init_db(engine: AsyncEngine) -> None:
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)


class Database:
    """Движок + фабрика сессий. Один экземпляр на процесс."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self.engine = make_engine(settings)
        self.sessions = make_session_factory(self.engine)

    @asynccontextmanager
    async def session(self) -> AsyncIterator[AsyncSession]:
        async with self.sessions() as s:
            yield s

    async def init(self) -> None:
        await init_db(self.engine)

    async def dispose(self) -> None:
        await self.engine.dispose()

    # ── app_state ─────────────────────────────────────────────
    async def get_state(self, key: str) -> dict[str, Any]:
        async with self.session() as s:
            row = await s.get(AppState, key)
            return dict(row.value) if row else {}

    async def set_state(self, key: str, value: dict[str, Any]) -> None:
        async with self.session() as s:
            row = await s.get(AppState, key)
            if row is None:
                s.add(AppState(key=key, value=value, updated_at=utcnow()))
            else:
                row.value = value
                row.updated_at = utcnow()
            await s.commit()

    async def all_state(self) -> dict[str, dict[str, Any]]:
        async with self.session() as s:
            rows = (await s.execute(select(AppState))).scalars().all()
            return {r.key: dict(r.value) for r in rows}
