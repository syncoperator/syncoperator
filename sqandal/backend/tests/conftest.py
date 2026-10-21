"""Общие фикстуры. Async-тесты гоняем через asyncio.run — без плагинов к pytest.

PostgreSQL-тесты включаются переменной SQANDAL_TEST_PG, например:
  SQANDAL_TEST_PG=postgresql+psycopg://postgres@127.0.0.1:5432/sqandal_test pytest
"""

from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

import pytest

BACKEND = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BACKEND))

from sqandal.settings import Settings  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"


def run(coro):
    return asyncio.run(coro)


@pytest.fixture
def fixture_text():
    def _read(name: str) -> str:
        return (FIXTURES / name).read_text(encoding="utf-8")

    return _read


@pytest.fixture
def settings_factory(tmp_path):
    def _make(**overrides) -> Settings:
        base = dict(
            database_url=f"sqlite+aiosqlite:///{tmp_path}/test.db",
            config_dir=BACKEND / "config",
            embedding_backend="hashing",
            dashboard_token="test-token",
            anthropic_api_key="sk-test",
            tg_bot_token="123:TEST",
            respect_robots=False,
            browser_enabled=False,
            frontend_dist=tmp_path / "no-dist",
            _env_file=None,
        )
        base.update(overrides)
        return Settings(**base)

    return _make


PG_URL = os.environ.get("SQANDAL_TEST_PG")
pg_only = pytest.mark.skipif(not PG_URL, reason="SQANDAL_TEST_PG не задан")
