"""Снимок реальных новостей → данные для режима «Реальные новости» в демо-дашборде.

Вход — папка с двумя файлами:
  items.json  публикации из настоящих источников: заголовок, ссылка, дата публикации
              (как её отдаёт источник), короткий пересказ фактов;
  posts.json  подготовленные посты для части сюжетов (ключ — ссылка любой публикации сюжета).

Публикации идут через настоящий пайплайн (релевантность, дата первоисточника, сюжеты,
теги, приоритет, свежесть), посты — через настоящий Newsroom (нормализация, фактчек
чисел, рендер). Подменён только ответ Claude API — подготовленным текстом.

Часы — как у работающего воркера: каждая публикация «приходит» в момент выхода, и сюжет
оценивается на этот момент (приоритет считается один раз, при приходе). Потом, как
housekeeping воркера, сюжеты старше STALE_AFTER_HOURS на момент снимка помечаются
«старым фактом» — приоритет при этом не пересчитывается.

    python -m sqandal.cli snapshot snapshot/2026-10-21 --export ../frontend/src/demo/snapshot-data.json
"""

from __future__ import annotations

import json
import logging
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import date, datetime, time
from pathlib import Path
from zoneinfo import ZoneInfo

import httpx
import yaml
from sqlalchemy import select

from . import services as services_mod
from .bootstrap import build_core, load_and_validate_sources, sync_sources
from .collectors.base import RawItem
from .llm.prompts import (
    ARTICLE_FROM_FACTS,
    ARTICLE_INSTRUCTIONS,
    ARTICLE_SCHEMA,
    FACTS_INSTRUCTIONS,
    FACTS_SCHEMA,
    MODE_INSTRUCTIONS,
    OUTPUT_SCHEMA,
    SYSTEM_PROMPT,
    TAG_NAMES,
    TOPIC_EMOJI,
    _xml_attr,
)
from .models import Cluster, DateSource, NewsItem, Source
from .processing import ingest as ingest_mod
from .processing.embeddings import make_embedder
from .processing.freshness import assess_staleness
from .processing.ingest import IngestPipeline
from .settings import Settings
from .timeutil import UTC, parse_datetime

log = logging.getLogger(__name__)


class PreparedPostsTransport(httpx.AsyncBaseTransport):
    """Вместо Claude API: подготовленный пост для сюжета, найденного по ссылке в промпте."""

    def __init__(self, posts: dict[str, dict], model: str):
        self.posts = posts
        self.model = model
        self.used: set[str] = set()

    async def handle_async_request(self, request: httpx.Request) -> httpx.Response:
        prompt = json.loads(request.content)["messages"][0]["content"]
        for url, post in self.posts.items():
            if f'url="{_xml_attr(url)}"' in prompt:
                self.used.add(url)
                return httpx.Response(200, json={
                    "id": "msg_snapshot", "type": "message", "role": "assistant", "model": self.model,
                    "content": [{"type": "text", "text": json.dumps(post, ensure_ascii=False)}],
                    "stop_reason": "end_turn",
                    "usage": {"input_tokens": 0, "output_tokens": 0},
                })
        return httpx.Response(404, json={"type": "error", "error": {
            "type": "not_found_error", "message": "для этого сюжета нет подготовленного поста"}})


@contextmanager
def _clock(at: datetime) -> Iterator[None]:
    """Остановить часы пайплайна и редакции на моменте `at` (только для пересборки снимка)."""
    saved = ingest_mod.utcnow, services_mod.utcnow
    ingest_mod.utcnow = services_mod.utcnow = lambda: at
    try:
        yield
    finally:
        ingest_mod.utcnow, services_mod.utcnow = saved


def _published(item: dict, tz: str | None) -> tuple[datetime, bool]:
    if item["date_source"] == DateSource.page_day.value:
        d = date.fromisoformat(item["published"])
        return datetime.combine(d, time(12), tzinfo=ZoneInfo(tz or "UTC")).astimezone(UTC), False
    dt, has_tz = parse_datetime(item["published"], tz)
    if dt is None:
        raise ValueError(f"Не разобрать дату {item['published']!r} у {item['url']}")
    return dt, has_tz


async def build_snapshot(settings: Settings, folder: Path, export: Path) -> dict:
    items_doc = json.loads((folder / "items.json").read_text(encoding="utf-8"))
    posts_doc = json.loads((folder / "posts.json").read_text(encoding="utf-8"))
    collected_at, _ = parse_datetime(items_doc["collected_at"])

    settings = settings.model_copy(update={
        "anthropic_api_key": settings.anthropic_api_key or "snapshot",
        "tg_bot_token": None,  # в снимке нет бота: дашборд покажет, что отправка — на сервере
        "post_link_preview": False,
    })
    transport = PreparedPostsTransport(posts_doc["posts"], posts_doc.get("model", "claude"))
    core = await build_core(settings, claude_transport=transport)
    core.newsroom.fetcher = None  # пересказы уже в items.json — статьи не докачиваем
    refs = await sync_sources(core.db, load_and_validate_sources(settings))
    pipe = IngestPipeline(core.db, settings, make_embedder(settings), core.tagger)

    raws: list[tuple[datetime, str, RawItem]] = []
    for it in items_doc["items"]:
        ref = refs[it["source"]]
        published, has_tz = _published(it, ref.cfg.tz)
        raws.append((published, it["source"], RawItem(
            external_id=it["url"], title=it["title"], url=it["url"], summary=it["summary"],
            published_at=published, date_source=DateSource(it["date_source"]), date_has_tz=has_tz,
            fetched_at=collected_at,
        )))
    # в порядке выхода и в момент выхода — как публикации приходили бы в работающий воркер
    for published, key, raw in sorted(raws, key=lambda r: r[0]):
        with _clock(min(published, collected_at)):
            await pipe.ingest(refs[key], [raw])

    async with core.db.session() as s:
        rows = (await s.execute(
            select(NewsItem.url, NewsItem.cluster_id, Source.key, NewsItem.effective_at)
            .join(Source, Source.id == NewsItem.source_id)
            .where(NewsItem.cluster_id.is_not(None))
        )).all()
    per_source: dict[str, int] = {}
    last_at: dict[int, datetime] = {}
    for _, cid, key, eff in rows:
        per_source[key] = per_source.get(key, 0) + 1
        last_at[cid] = max(last_at.get(cid, eff), eff)

    # пост пишется, когда сюжет пришёл (как автогенерация воркера для CRITICAL/HIGH)
    clusters_with_post = sorted({cid for url, cid, _, _ in rows if url in posts_doc["posts"]})
    for cid in clusters_with_post:
        with _clock(min(last_at[cid], collected_at)):
            await core.newsroom.generate_post(cid)
    unused = sorted(set(posts_doc["posts"]) - transport.used)
    if unused:
        log.warning("Подготовленные посты не пригодились (ссылки не в сюжетах): %s", unused)

    # housekeeping к моменту снимка: старые сюжеты → «старый факт», приоритет как был
    async with core.db.session() as s:
        for c in (await s.execute(select(Cluster).where(Cluster.is_stale.is_(False)))).scalars():
            c.is_stale, c.stale_reason = assess_staleness(
                c.origin_published_at, collected_at, settings.stale_after_hours)
        await s.commit()

    data = await export_dashboard_for_snapshot(core, settings)
    at = collected_at.isoformat()
    data["generated_at"] = at
    data["kind"] = "snapshot"
    data["config"]["snapshot_at"] = at
    # здоровье источников — как прошёл сбор этого снимка
    status = items_doc.get("source_status", {})
    ok = bad = 0
    for src in data["sources"]:
        st = status.get(src["key"], {})
        if src["key"] in per_source or st.get("ok"):
            src.update(last_fetch_at=at, last_success_at=at, last_error=None, consecutive_failures=0,
                       items_total=per_source.get(src["key"], 0))
            ok += src["enabled"]
        elif st:
            src.update(last_fetch_at=at, last_error=st["message"], consecutive_failures=1)
            bad += src["enabled"]
        if st.get("ok") and st.get("message"):
            src["snapshot_note"] = st["message"]
    data["stats"].update(last_collect_at=at, sources_ok=ok, sources_failing=bad)
    # приоритет «как в tags.yaml» — дашборд считает его сам для сюжетов, которые соберёт Claude по кнопке
    artists_cfg = yaml.safe_load((settings.config_dir / "artists.yaml").read_text(encoding="utf-8")) or {}
    data["scoring"] = {
        "thresholds": core.tagger.thresholds,
        "bonuses": core.tagger.bonuses,
        "artists": [
            {"name": e["name"], "aliases": e.get("aliases", []), "tier": int(e.get("tier", 2))}
            for entries in artists_cfg.values() for e in (entries or [])
        ],
    }
    data["prompt"] = {
        "system": SYSTEM_PROMPT,
        "modes": MODE_INSTRUCTIONS,
        "tags": TAG_NAMES,
        "topic_emoji": TOPIC_EMOJI,
        "schema": OUTPUT_SCHEMA,
        "article": ARTICLE_INSTRUCTIONS,
        "article_schema": ARTICLE_SCHEMA,
        "facts": FACTS_INSTRUCTIONS,
        "facts_schema": FACTS_SCHEMA,
        "article_from_facts": ARTICLE_FROM_FACTS,
        "audience_tz": settings.audience_tz,
        "max_sources": settings.llm_max_sources,
        "source_chars": settings.llm_source_chars,
    }
    export.parent.mkdir(parents=True, exist_ok=True)
    export.write_text(json.dumps(data, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return {
        "items": len(raws),
        "stories": len(data["stories"]),
        "posts": len(clusters_with_post),
        "unused_posts": unused,
        "export": str(export),
    }


async def export_dashboard_for_snapshot(core, settings: Settings) -> dict:
    from .cli import export_dashboard

    return await export_dashboard(core, settings)
