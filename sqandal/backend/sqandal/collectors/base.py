"""Общие типы сборщиков."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator

from ..models import DateSource, Scene, SourceKind
from ..timeutil import utcnow


@dataclass(slots=True)
class RawItem:
    """Нормализованная публикация из любого источника — вход пайплайна."""

    external_id: str
    title: str
    url: str | None = None
    summary: str | None = None
    content: str | None = None
    author: str | None = None
    image_url: str | None = None
    published_at: datetime | None = None
    date_source: DateSource = DateSource.fetched
    date_has_tz: bool = True
    lang: str | None = None
    tg_channel: str | None = None
    tg_message_id: int | None = None
    tg_views: int | None = None
    fwd_from: str | None = None
    extra: dict[str, Any] = field(default_factory=dict)
    fetched_at: datetime = field(default_factory=utcnow)

    def text_blob(self) -> str:
        return "\n".join(x for x in (self.title, self.summary, self.content) if x)


@dataclass(slots=True)
class CollectResult:
    items: list[RawItem] = field(default_factory=list)
    not_modified: bool = False
    etag: str | None = None
    last_modified: str | None = None
    state_updates: dict[str, Any] = field(default_factory=dict)
    seen: int = 0


class SourceConfig(BaseModel):
    """Описание источника в config/sources.yaml."""

    key: str = Field(pattern=r"^[a-z0-9_\-]{2,64}$")
    name: str
    kind: SourceKind
    scene: Scene
    url: str = ""
    enabled: bool = True
    interval_min: int | None = None
    # фильтр релевантности для общих СМИ (TMZ, BILD): пропускать только хип-хоп
    filter: Literal["hiphop"] | None = None
    # регулярка по заголовку: совпало → не новость (подборки, рейтинги, промо)
    exclude_pattern: str | None = None
    # RSS: найти фид на HTML-странице (link rel=alternate / ссылки *.xml)
    discover: bool = False
    discover_match: str | None = None
    # web: регулярка для ссылок на статьи на странице-листинге
    link_pattern: str | None = None
    render: Literal["http", "browser"] = "http"
    wait_selector: str | None = None
    fetch_full_text: bool = True
    tz: str | None = None  # часовой пояс для дат без зоны
    # telegram
    channel: str | None = None
    insider: bool = True
    note: str | None = None

    @model_validator(mode="after")
    def _check(self) -> "SourceConfig":
        if self.kind == SourceKind.telegram:
            if not self.channel:
                raise ValueError(f"{self.key}: для telegram нужен channel")
            self.channel = self.channel.lstrip("@").replace("https://t.me/", "").strip("/")
            if not self.url:
                self.url = f"https://t.me/{self.channel}"
        elif not self.url:
            raise ValueError(f"{self.key}: нужен url")
        if self.kind == SourceKind.web and not self.link_pattern:
            raise ValueError(f"{self.key}: для web нужен link_pattern")
        return self


class SourcesFile(BaseModel):
    sources: list[SourceConfig]
