"""Сборка зависимостей процесса (API или воркер) из настроек."""

from __future__ import annotations

import logging
from dataclasses import dataclass

from sqlalchemy import select

from .collectors import SourceConfig, load_sources
from .collectors.http import Fetcher
from .db import Database
from .llm.claude import ClaudeClient
from .llm.generator import PostGenerator
from .models import Source
from .processing.ingest import SourceRef
from .processing.tagging import Tagger
from .publish.telegram_bot import TelegramBot
from .services import Newsroom
from .settings import Settings
from .timeutil import utcnow

log = logging.getLogger(__name__)


@dataclass
class Core:
    settings: Settings
    db: Database
    tagger: Tagger
    fetcher: Fetcher
    claude: ClaudeClient | None
    bot: TelegramBot | None
    newsroom: Newsroom

    async def close(self) -> None:
        await self.fetcher.close()
        if self.claude:
            await self.claude.close()
        if self.bot:
            await self.bot.close()
        await self.db.dispose()


async def build_core(settings: Settings, *, claude_transport=None, bot_transport=None, http_transport=None) -> Core:
    db = Database(settings)
    await db.init()
    tagger = Tagger.from_dir(settings.config_dir)
    fetcher = Fetcher(settings, transport=http_transport)
    claude = None
    generator = None
    if settings.anthropic_api_key:
        claude = ClaudeClient(settings.anthropic_api_key, settings.anthropic_base_url, settings.llm_timeout_s,
                              transport=claude_transport)
        generator = PostGenerator(settings, claude)
    else:
        log.warning("ANTHROPIC_API_KEY не задан — генерация постов выключена")
    bot = TelegramBot(settings.tg_bot_token, transport=bot_transport) if settings.tg_bot_token else None
    if settings.dashboard_token in ("", "change-me"):
        log.warning("DASHBOARD_TOKEN по умолчанию — смени перед выкладкой наружу!")
    newsroom = Newsroom(db, settings, tagger, generator, bot, fetcher)
    return Core(settings, db, tagger, fetcher, claude, bot, newsroom)


async def sync_sources(db: Database, configs: list[SourceConfig]) -> dict[str, SourceRef]:
    """sources.yaml → таблица sources (upsert по key). Возвращает key → SourceRef."""
    refs: dict[str, SourceRef] = {}
    async with db.session() as s:
        existing = {x.key: x for x in (await s.execute(select(Source))).scalars()}
        for cfg in configs:
            row = existing.get(cfg.key)
            data = cfg.model_dump(mode="json")
            if row is None:
                row = Source(key=cfg.key, name=cfg.name, kind=cfg.kind, scene=cfg.scene, url=cfg.url,
                             enabled=cfg.enabled, config=data, state={})
                s.add(row)
            else:
                url_changed = row.url != cfg.url
                row.name, row.kind, row.scene, row.url = cfg.name, cfg.kind, cfg.scene, cfg.url
                row.enabled, row.config = cfg.enabled, data
                if url_changed:  # сменился адрес → сбросить кэш фида и условный GET
                    row.state, row.etag, row.last_modified = {}, None, None
                row.updated_at = utcnow()
        for key, row in existing.items():
            if key not in {c.key for c in configs} and row.enabled:
                row.enabled = False  # удалён из конфига → выключаем, историю не трогаем
        await s.commit()
        rows = {x.key: x for x in (await s.execute(select(Source))).scalars()}
    for cfg in configs:
        r = rows[cfg.key]
        refs[cfg.key] = SourceRef(r.id, r.key, r.name, r.kind, r.scene, cfg)
    return refs


def load_and_validate_sources(settings: Settings) -> list[SourceConfig]:
    return load_sources(settings.config_dir)
