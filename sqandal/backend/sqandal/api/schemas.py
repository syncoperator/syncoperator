"""Pydantic-схемы ответов/запросов API дашборда."""

from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field


class OriginOut(BaseModel):
    published_at: datetime
    confidence: str
    source_name: str | None = None
    source_kind: str | None = None
    url: str | None = None
    date_source: str | None = None


class PostBrief(BaseModel):
    id: int
    version: int
    mode: str
    status: str
    scheduled_at: datetime | None = None
    published_at: datetime | None = None
    created_at: datetime


class StoryCard(BaseModel):
    id: int
    title: str
    scene: str
    scenes: list[str]
    tags: list[str]
    artists: list[str]
    priority: str
    priority_score: int
    status: str
    origin: OriginOut
    first_seen_at: datetime
    last_item_at: datetime
    items_count: int
    sources_count: int
    source_names: list[str] = Field(default_factory=list)
    verification: str
    is_stale: bool
    stale_reason: str | None = None
    republish_of: int | None = None
    needs_refresh: bool
    post: PostBrief | None = None
    updated_at: datetime


class ItemOut(BaseModel):
    id: int
    source_name: str
    source_kind: str
    url: str | None
    title: str
    summary: str | None
    published_at: datetime | None
    effective_at: datetime
    date_source: str
    date_confidence: str
    fetched_at: datetime
    lang: str | None
    tg_views: int | None
    fwd_from: str | None
    similarity: float | None
    date_flags: list[str] = Field(default_factory=list)


class PostOut(BaseModel):
    id: int
    cluster_id: int
    version: int
    mode: str
    status: str
    flag_emoji: str
    topic_emoji: str
    headline: str
    bullets: list[str]
    summary: str
    tags: list[str]
    sources: list[dict]
    footnotes: list[str]
    body_html: str
    body_text: str
    warnings: list[str]
    editor_notes: str | None
    is_publishable: bool
    event_date: str | None
    llm_model: str | None
    input_tokens: int | None
    output_tokens: int | None
    cache_read_tokens: int | None
    scheduled_at: datetime | None
    published_at: datetime | None
    tg_message_id: int | None
    link: str | None = None
    error: str | None
    created_at: datetime
    updated_at: datetime


class StoryDetail(StoryCard):
    items: list[ItemOut]
    current_post: PostOut | None = None
    versions: list[PostBrief] = Field(default_factory=list)


class StoryList(BaseModel):
    stories: list[StoryCard]
    server_time: datetime


class GenerateIn(BaseModel):
    mode: Literal["standard", "harder", "hardest"] = "standard"


class PostPatch(BaseModel):
    flag_emoji: str | None = None
    topic_emoji: str | None = None
    headline: str | None = Field(default=None, max_length=300)
    bullets: list[str] | None = Field(default=None, max_length=6)
    summary: str | None = Field(default=None, max_length=400)
    sources: list[dict] | None = None


class PublishIn(BaseModel):
    confirm_stale: bool = False


class ScheduleIn(BaseModel):
    scheduled_at: datetime


class SourceOut(BaseModel):
    key: str
    name: str
    kind: str
    scene: str
    url: str
    enabled: bool
    last_fetch_at: datetime | None
    last_success_at: datetime | None
    last_error: str | None
    consecutive_failures: int
    items_total: int
    note: str | None = None


class StatsOut(BaseModel):
    by_scene: dict[str, int]
    by_priority: dict[str, int]
    critical_by_scene: dict[str, int] = Field(default_factory=dict)  # красные точки на вкладках сцен
    drafts: int
    scheduled: int
    published_24h: int
    last_collect_at: datetime | None
    sources_ok: int
    sources_failing: int
    server_time: datetime


class ConfigOut(BaseModel):
    channel: str
    parse_mode: str
    brand_line: str
    brand_url: str | None
    bullet: str
    summary_style: str
    show_sources: bool
    tags: list[dict]
    llm_ready: bool
    bot_ready: bool
    telethon_ready: bool
    stale_after_hours: float
    audience_tz: str
    version: str
