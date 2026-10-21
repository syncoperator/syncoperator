"""ORM-модели (SQLAlchemy 2.0). Одна схема для PostgreSQL и SQLite.

Сущности:
  sources      — источники (RSS, сайты, Telegram-каналы) + их здоровье
  news_items   — сырые новости из источников (одна запись = одна публикация)
  clusters     — «сюжеты»: группы news_items об одном событии (дедуп по смыслу)
  posts        — сгенерированные/отредактированные ТГ-посты (версии) и их публикация
  fetch_runs   — журнал прогонов сборщиков
  app_state    — служебные ключи (команды воркеру, метки времени)
"""

from __future__ import annotations

import enum
from datetime import datetime
from typing import Any

from sqlalchemy import (
    JSON,
    MetaData,
    BigInteger,
    Boolean,
    DateTime,
    Enum,
    Float,
    ForeignKey,
    Index,
    Integer,
    LargeBinary,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
from sqlalchemy.types import TypeDecorator

from .timeutil import UTC, utcnow

# BIGINT в PostgreSQL; INTEGER в SQLite (иначе нет автоинкремента rowid)
BigId = BigInteger().with_variant(Integer(), "sqlite")
JsonType = JSON().with_variant(JSONB(), "postgresql")


class UTCDateTime(TypeDecorator):
    """timestamptz в PG; в SQLite — naive UTC. Наружу всегда aware UTC."""

    impl = DateTime(timezone=True)
    cache_ok = True

    def process_bind_param(self, value: datetime | None, dialect):
        if value is None:
            return None
        if value.tzinfo is None:
            raise ValueError("naive datetime запрещён — используйте aware UTC")
        value = value.astimezone(UTC)
        if dialect.name == "sqlite":
            return value.replace(tzinfo=None)
        return value

    def process_result_value(self, value: datetime | None, dialect):
        if value is None:
            return None
        if value.tzinfo is None:
            return value.replace(tzinfo=UTC)
        return value.astimezone(UTC)


def _enum(cls: type[enum.Enum], name: str) -> Enum:
    return Enum(
        cls,
        name=name,
        native_enum=False,
        create_constraint=True,
        values_callable=lambda e: [m.value for m in e],
        length=24,
        validate_strings=True,
    )


class Scene(str, enum.Enum):
    ru = "ru"  # Россия / СНГ
    de = "de"  # Германия / DACH
    us = "us"  # США
    intl = "intl"  # прочее


class SourceKind(str, enum.Enum):
    rss = "rss"
    web = "web"
    telegram = "telegram"


class DateSource(str, enum.Enum):
    """Откуда взята дата первоисточника — от этого зависит доверие к ней."""

    feed_published = "feed_published"
    feed_updated = "feed_updated"
    jsonld = "jsonld"
    meta = "meta"
    time_tag = "time_tag"
    url = "url"
    telegram = "telegram"
    telegram_forward = "telegram_forward"
    page_day = "page_day"  # дата в шапке статьи без времени («20 Oktober, 2026») — только день
    fetched = "fetched"  # даты нет — время, когда мы нашли новость


class Confidence(str, enum.Enum):
    high = "high"
    medium = "medium"
    low = "low"


class Priority(str, enum.Enum):
    CRITICAL = "CRITICAL"
    HIGH = "HIGH"
    NORMAL = "NORMAL"


PRIORITY_RANK = {Priority.CRITICAL: 3, Priority.HIGH: 2, Priority.NORMAL: 1}


class ClusterStatus(str, enum.Enum):
    new = "new"
    drafted = "drafted"
    scheduled = "scheduled"
    published = "published"
    dismissed = "dismissed"


class Verification(str, enum.Enum):
    single_source = "single_source"
    multi_source = "multi_source"
    insider_only = "insider_only"  # только Telegram-каналы


class PostMode(str, enum.Enum):
    standard = "standard"
    harder = "harder"
    hardest = "hardest"
    manual = "manual"


class PostStatus(str, enum.Enum):
    draft = "draft"
    scheduled = "scheduled"
    publishing = "publishing"
    published = "published"
    failed = "failed"
    superseded = "superseded"


NAMING_CONVENTION = {
    "ix": "ix_%(column_0_label)s",
    "uq": "uq_%(table_name)s_%(column_0_name)s",
    "ck": "ck_%(table_name)s_%(constraint_name)s",
    "fk": "fk_%(table_name)s_%(column_0_name)s_%(referred_table_name)s",
    "pk": "pk_%(table_name)s",
}


class Base(DeclarativeBase):
    metadata = MetaData(naming_convention=NAMING_CONVENTION)


class TimestampMixin:
    created_at: Mapped[datetime] = mapped_column(UTCDateTime(), default=utcnow, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        UTCDateTime(), default=utcnow, onupdate=utcnow, nullable=False
    )


class Source(TimestampMixin, Base):
    __tablename__ = "sources"
    __table_args__ = {"comment": "Источники новостей: RSS, сайты, Telegram-каналы"}

    id: Mapped[int] = mapped_column(BigId, primary_key=True, autoincrement=True)
    key: Mapped[str] = mapped_column(String(64), unique=True, nullable=False, comment="Слаг из sources.yaml")
    name: Mapped[str] = mapped_column(String(128), nullable=False)
    kind: Mapped[SourceKind] = mapped_column(_enum(SourceKind, "source_kind"), nullable=False)
    scene: Mapped[Scene] = mapped_column(_enum(Scene, "scene"), nullable=False)
    url: Mapped[str] = mapped_column(Text, nullable=False)
    enabled: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)
    config: Mapped[dict[str, Any]] = mapped_column(JsonType, default=dict, nullable=False)
    state: Mapped[dict[str, Any]] = mapped_column(
        JsonType, default=dict, nullable=False, comment="Состояние сборщика (найденный фид, курсоры)"
    )
    etag: Mapped[str | None] = mapped_column(Text)
    last_modified: Mapped[str | None] = mapped_column(Text)
    tg_last_message_id: Mapped[int | None] = mapped_column(BigInteger, comment="Курсор догона Telegram")
    last_fetch_at: Mapped[datetime | None] = mapped_column(UTCDateTime())
    last_success_at: Mapped[datetime | None] = mapped_column(UTCDateTime())
    last_error: Mapped[str | None] = mapped_column(Text)
    consecutive_failures: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    items_total: Mapped[int] = mapped_column(Integer, default=0, nullable=False)


class Cluster(TimestampMixin, Base):
    __tablename__ = "clusters"
    __table_args__ = (
        Index("ix_clusters_status_priority", "status", "priority"),
        Index("ix_clusters_scene_last_item", "scene", "last_item_at"),
        Index("ix_clusters_origin_published", "origin_published_at"),
        {"comment": "Сюжеты: кластеры новостей об одном событии"},
    )

    id: Mapped[int] = mapped_column(BigId, primary_key=True, autoincrement=True)
    title: Mapped[str] = mapped_column(Text, nullable=False)
    scene: Mapped[Scene] = mapped_column(_enum(Scene, "scene"), nullable=False)
    scenes: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    tags: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    artists: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    priority: Mapped[Priority] = mapped_column(_enum(Priority, "priority"), default=Priority.NORMAL, nullable=False)
    priority_score: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    status: Mapped[ClusterStatus] = mapped_column(
        _enum(ClusterStatus, "cluster_status"), default=ClusterStatus.new, nullable=False
    )
    # мягкая ссылка (без FK): news_items.cluster_id → clusters уже даёт цикл
    origin_item_id: Mapped[int | None] = mapped_column(BigInteger, comment="Самая ранняя публикация сюжета")
    origin_source_id: Mapped[int | None] = mapped_column(ForeignKey("sources.id", ondelete="SET NULL"))
    origin_published_at: Mapped[datetime] = mapped_column(
        UTCDateTime(), nullable=False, comment="Дата/время первоисточника (min по сюжету)"
    )
    origin_date_confidence: Mapped[Confidence] = mapped_column(
        _enum(Confidence, "confidence"), default=Confidence.low, nullable=False
    )
    first_seen_at: Mapped[datetime] = mapped_column(UTCDateTime(), nullable=False)
    last_item_at: Mapped[datetime] = mapped_column(UTCDateTime(), nullable=False)
    items_count: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    sources_count: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    verification: Mapped[Verification] = mapped_column(
        _enum(Verification, "verification"), default=Verification.single_source, nullable=False
    )
    is_stale: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    stale_reason: Mapped[str | None] = mapped_column(Text)
    republish_of_cluster_id: Mapped[int | None] = mapped_column(
        BigInteger, comment="Похож на уже опубликованный сюжет"
    )
    needs_refresh: Mapped[bool] = mapped_column(
        Boolean, default=False, nullable=False, comment="Появились новые источники после черновика"
    )
    centroid: Mapped[bytes | None] = mapped_column(LargeBinary)
    centroid_model: Mapped[str | None] = mapped_column(String(160))

    items: Mapped[list[NewsItem]] = relationship(back_populates="cluster", lazy="raise")
    posts: Mapped[list[Post]] = relationship(back_populates="cluster", lazy="raise")


class NewsItem(Base):
    __tablename__ = "news_items"
    __table_args__ = (
        UniqueConstraint("source_id", "external_id", name="uq_news_items_source_external"),
        Index("ix_news_items_canonical_url", "canonical_url"),
        Index("ix_news_items_content_hash", "content_hash"),
        Index("ix_news_items_effective_at", "effective_at"),
        Index("ix_news_items_cluster", "cluster_id"),
        {"comment": "Сырые публикации из источников"},
    )

    id: Mapped[int] = mapped_column(BigId, primary_key=True, autoincrement=True)
    source_id: Mapped[int] = mapped_column(ForeignKey("sources.id", ondelete="CASCADE"), nullable=False)
    external_id: Mapped[str] = mapped_column(String(512), nullable=False, comment="guid / url / channel:msg_id")
    url: Mapped[str | None] = mapped_column(Text)
    canonical_url: Mapped[str | None] = mapped_column(Text)
    title: Mapped[str] = mapped_column(Text, nullable=False)
    summary: Mapped[str | None] = mapped_column(Text)
    content: Mapped[str | None] = mapped_column(Text, comment="Полный текст (если скачан)")
    author: Mapped[str | None] = mapped_column(String(256))
    image_url: Mapped[str | None] = mapped_column(Text)
    lang: Mapped[str | None] = mapped_column(String(8))
    scene: Mapped[Scene] = mapped_column(_enum(Scene, "scene"), nullable=False)
    published_at: Mapped[datetime | None] = mapped_column(UTCDateTime(), comment="Дата, заявленная источником")
    date_source: Mapped[DateSource] = mapped_column(_enum(DateSource, "date_source"), nullable=False)
    date_confidence: Mapped[Confidence] = mapped_column(_enum(Confidence, "confidence"), nullable=False)
    fetched_at: Mapped[datetime] = mapped_column(UTCDateTime(), nullable=False)
    effective_at: Mapped[datetime] = mapped_column(
        UTCDateTime(), nullable=False, comment="Проверенное время публикации (для первоисточника)"
    )
    content_hash: Mapped[str] = mapped_column(String(64), nullable=False)
    embedding: Mapped[bytes | None] = mapped_column(LargeBinary)
    embedding_model: Mapped[str | None] = mapped_column(String(160))
    cluster_id: Mapped[int | None] = mapped_column(ForeignKey("clusters.id", ondelete="SET NULL"))
    similarity: Mapped[float | None] = mapped_column(Float)
    is_relevant: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)
    tags: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    tg_channel: Mapped[str | None] = mapped_column(String(128))
    tg_message_id: Mapped[int | None] = mapped_column(BigInteger)
    tg_views: Mapped[int | None] = mapped_column(Integer)
    fwd_from: Mapped[str | None] = mapped_column(String(256), comment="Исходный канал/автор репоста")
    extra: Mapped[dict[str, Any]] = mapped_column(JsonType, default=dict, nullable=False)
    created_at: Mapped[datetime] = mapped_column(UTCDateTime(), default=utcnow, nullable=False)

    source: Mapped[Source] = relationship(lazy="raise")
    cluster: Mapped[Cluster | None] = relationship(back_populates="items", lazy="raise")


class Post(TimestampMixin, Base):
    __tablename__ = "posts"
    __table_args__ = (
        UniqueConstraint("cluster_id", "version", name="uq_posts_cluster_version"),
        Index("ix_posts_status_scheduled", "status", "scheduled_at"),
        {"comment": "ТГ-посты: версии генерации, правки, публикации"},
    )

    id: Mapped[int] = mapped_column(BigId, primary_key=True, autoincrement=True)
    cluster_id: Mapped[int] = mapped_column(ForeignKey("clusters.id", ondelete="CASCADE"), nullable=False)
    version: Mapped[int] = mapped_column(Integer, nullable=False)
    mode: Mapped[PostMode] = mapped_column(_enum(PostMode, "post_mode"), nullable=False)
    status: Mapped[PostStatus] = mapped_column(
        _enum(PostStatus, "post_status"), default=PostStatus.draft, nullable=False
    )
    flag_emoji: Mapped[str] = mapped_column(String(32), default="", nullable=False)
    topic_emoji: Mapped[str] = mapped_column(String(32), default="", nullable=False)
    headline: Mapped[str] = mapped_column(Text, nullable=False)
    bullets: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    summary: Mapped[str] = mapped_column(Text, default="", nullable=False)
    tags: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    sources: Mapped[list[dict[str, Any]]] = mapped_column(
        JsonType, default=list, nullable=False, comment="[{name, url}] для строки источников"
    )
    footnotes: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    body_html: Mapped[str] = mapped_column(Text, nullable=False)
    body_text: Mapped[str] = mapped_column(Text, nullable=False)
    warnings: Mapped[list[str]] = mapped_column(JsonType, default=list, nullable=False)
    editor_notes: Mapped[str | None] = mapped_column(Text)
    is_publishable: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)
    event_date: Mapped[str | None] = mapped_column(String(10))
    llm_model: Mapped[str | None] = mapped_column(String(64))
    input_tokens: Mapped[int | None] = mapped_column(Integer)
    output_tokens: Mapped[int | None] = mapped_column(Integer)
    cache_read_tokens: Mapped[int | None] = mapped_column(Integer)
    scheduled_at: Mapped[datetime | None] = mapped_column(UTCDateTime())
    published_at: Mapped[datetime | None] = mapped_column(UTCDateTime())
    tg_chat_id: Mapped[str | None] = mapped_column(String(64))
    tg_message_id: Mapped[int | None] = mapped_column(BigInteger)
    error: Mapped[str | None] = mapped_column(Text)

    cluster: Mapped[Cluster] = relationship(back_populates="posts", lazy="raise")


class FetchRun(Base):
    __tablename__ = "fetch_runs"
    __table_args__ = (
        Index("ix_fetch_runs_source_started", "source_id", "started_at"),
        {"comment": "Журнал прогонов сборщиков"},
    )

    id: Mapped[int] = mapped_column(BigId, primary_key=True, autoincrement=True)
    source_id: Mapped[int] = mapped_column(ForeignKey("sources.id", ondelete="CASCADE"), nullable=False)
    started_at: Mapped[datetime] = mapped_column(UTCDateTime(), nullable=False)
    finished_at: Mapped[datetime | None] = mapped_column(UTCDateTime())
    ok: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    items_seen: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    items_new: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    error: Mapped[str | None] = mapped_column(Text)
    duration_ms: Mapped[int | None] = mapped_column(Integer)


class AppState(Base):
    __tablename__ = "app_state"
    __table_args__ = {"comment": "Служебные ключи: команды воркеру, метки"}

    key: Mapped[str] = mapped_column(String(64), primary_key=True)
    value: Mapped[dict[str, Any]] = mapped_column(JsonType, default=dict, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        UTCDateTime(), default=utcnow, onupdate=utcnow, nullable=False
    )
