"""Конфигурация из переменных окружения / .env (pydantic-settings)."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Literal

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

BACKEND_DIR = Path(__file__).resolve().parent.parent

Priority = Literal["CRITICAL", "HIGH", "NORMAL"]


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=(".env", "../.env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # ── База данных ──────────────────────────────────────────────
    # SQLite (dev):      sqlite+aiosqlite:///./data/sqandal.db
    # PostgreSQL (prod): postgresql+psycopg://sqandal:pass@db:5432/sqandal
    database_url: str = "sqlite+aiosqlite:///./data/sqandal.db"
    db_echo: bool = False

    config_dir: Path = BACKEND_DIR / "config"

    # ── Сбор ─────────────────────────────────────────────────────
    collect_interval_min: int = 15
    http_timeout_s: float = 20.0
    http_user_agent: str = (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/141.0 Safari/537.36 SQANDALRadar/1.0 (+https://t.me/SQANDALru)"
    )
    respect_robots: bool = True
    max_links_per_listing: int = 30
    max_new_articles_per_run: int = 12
    first_run_max_items: int = 15
    max_item_age_days: int = 7  # старше — не считаем новостью (архив/«популярное»)
    browser_enabled: bool = True  # Playwright для JS-сайтов
    browser_concurrency: int = 2

    # ── Telegram-парсер (Telethon, user-аккаунт) ────────────────
    tg_api_id: int | None = None
    tg_api_hash: str | None = None
    tg_session: str | None = None  # StringSession из `python -m sqandal.cli tg-login`
    tg_catchup_limit: int = 30

    # ── Дедупликация ─────────────────────────────────────────────
    embedding_backend: Literal["local", "voyage", "hashing"] = "local"
    embedding_model: str = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
    voyage_api_key: str | None = None
    voyage_model: str = "voyage-4-lite"
    voyage_base_url: str = "https://ai.mongodb.com/v1"
    cluster_threshold: float | None = None  # None → дефолт под модель
    cluster_window_hours: int = 72
    stale_after_hours: float = 24.0
    republish_lookback_days: int = 14

    # ── LLM (Claude API) ─────────────────────────────────────────
    anthropic_api_key: str | None = None
    anthropic_base_url: str = "https://api.anthropic.com"
    llm_model: str = "claude-sonnet-5-5"
    llm_effort: Literal["low", "medium", "high"] = "low"
    llm_max_tokens: int = 6000
    llm_timeout_s: float = 120.0
    llm_concurrency: int = 3
    llm_source_chars: int = 2500  # сколько текста каждого источника отдаём модели
    llm_max_sources: int = 5
    auto_generate_min_priority: Literal["CRITICAL", "HIGH", "NORMAL", "OFF"] = "NORMAL"
    auto_generate_max_age_hours: float = 6.0  # старее — генерим только по клику (экономия токенов)
    audience_tz: str = "Europe/Moscow"

    # ── Публикация (Bot API) ─────────────────────────────────────
    tg_bot_token: str | None = None
    tg_channel_id: str = "@SQANDALru"
    tg_admin_chat_id: str | None = None  # личка/группа редакции для CRITICAL-алертов
    tg_admin_mode: Literal["off", "polling", "webhook"] = "off"
    tg_webhook_secret: str | None = None
    public_base_url: str | None = None  # https://radar.example.com — для setWebhook
    post_parse_mode: Literal["HTML", "MarkdownV2"] = "HTML"
    post_show_sources: bool = True
    post_link_preview: bool = True
    post_summary_style: Literal["quote", "bold", "italic", "plain"] = "quote"
    post_bullet: str = "▪️"
    brand_line: str = "🔊 SQANDAL"
    brand_url: str | None = "https://t.me/SQANDALru"

    # ── API / дашборд ────────────────────────────────────────────
    dashboard_token: str = "change-me"
    cors_origins: list[str] = Field(default_factory=lambda: ["http://localhost:5173"])
    frontend_dist: Path = BACKEND_DIR.parent / "frontend" / "dist"

    @field_validator("cors_origins", mode="before")
    @classmethod
    def _split_origins(cls, v):
        if isinstance(v, str) and not v.strip().startswith("["):
            return [s.strip() for s in v.split(",") if s.strip()]
        return v

    @property
    def is_sqlite(self) -> bool:
        return self.database_url.startswith("sqlite")

    @property
    def telethon_configured(self) -> bool:
        return bool(self.tg_api_id and self.tg_api_hash and self.tg_session)

    @property
    def bot_configured(self) -> bool:
        return bool(self.tg_bot_token)


@lru_cache
def get_settings() -> Settings:
    return Settings()
