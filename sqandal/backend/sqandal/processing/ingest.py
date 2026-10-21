"""Пайплайн приёма: RawItem → news_items → сюжеты (clusters) с тегами, приоритетом,
первоисточником и флагами свежести.

Три фазы, чтобы не держать транзакцию во время тяжёлых эмбеддингов:
  A. чтение: какие external_id уже есть;
  B. без БД: релевантность, теги, проверка даты, эмбеддинги;
  C. запись: вставка, привязка к сюжетам, пересчёт агрегатов — коротко.
Один asyncio.Lock на процесс: кластеризация строго последовательная.
"""

from __future__ import annotations

import asyncio
import logging
import re
from collections import Counter
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from datetime import timedelta

import numpy as np
from sqlalchemy import select

from ..collectors.base import RawItem, SourceConfig
from ..db import Database
from ..models import (
    Cluster,
    ClusterStatus,
    Confidence,
    NewsItem,
    Post,
    PostStatus,
    Scene,
    Source,
    SourceKind,
    Verification,
)
from ..settings import Settings
from ..textutil import canonicalize_url, content_hash, detect_lang, truncate
from ..timeutil import utcnow
from .clustering import ClusterIndex, PublishedIndex
from .embeddings import Embedder, embedding_text, from_bytes, threshold_for, to_bytes
from .freshness import assess_staleness, effective_time
from .tagging import ClusterSignals, Tagger, breaking_count

log = logging.getLogger(__name__)


@dataclass(slots=True)
class SourceRef:
    id: int
    key: str
    name: str
    kind: SourceKind
    scene: Scene
    cfg: SourceConfig


@dataclass(slots=True)
class IngestReport:
    received: int = 0
    new: int = 0
    relevant: int = 0
    created: list[int] = field(default_factory=list)
    updated: list[int] = field(default_factory=list)


@dataclass(slots=True)
class _Prepared:
    raw: RawItem
    canonical: str | None
    chash: str
    relevant: bool
    reason: str | None
    lang: str | None
    tags: list[str]
    artists: list[str]
    eff_at: object
    eff_conf: Confidence
    flags: list[str]
    vec: np.ndarray | None = None


ClusterCallback = Callable[[IngestReport], Awaitable[None]]


class IngestPipeline:
    def __init__(self, db: Database, settings: Settings, embedder: Embedder, tagger: Tagger):
        self.db = db
        self.settings = settings
        self.embedder = embedder
        self.tagger = tagger
        thr = threshold_for(embedder, settings)
        self.index = ClusterIndex(thr, settings.cluster_window_hours)
        self.published = PublishedIndex(thr, settings.republish_lookback_days)
        self.lock = asyncio.Lock()
        self._loaded = False
        # ссылки из TG-постов на статьи СМИ: canonical_url статьи → сюжет TG-поста.
        # Инсайд часто выходит раньше статьи и ссылается на неё — склеиваем точно, без эмбеддингов.
        self.link_index: dict[str, int] = {}

    # ── индекс активных сюжетов ───────────────────────────────
    async def load_index(self) -> None:
        now = utcnow()
        cutoff = now - timedelta(hours=self.settings.cluster_window_hours)
        model = self.embedder.model_id
        async with self.db.session() as s:
            clusters = (await s.execute(select(Cluster).where(Cluster.last_item_at >= cutoff))).scalars().all()
            by_id = {c.id: c for c in clusters}
            if by_id:
                rows = await s.execute(
                    select(NewsItem.cluster_id, NewsItem.embedding, NewsItem.effective_at)
                    .where(NewsItem.cluster_id.in_(list(by_id)))
                    .where(NewsItem.embedding_model == model)
                )
                for cid, emb, at in rows:
                    vec = from_bytes(emb)
                    if vec is None:
                        continue
                    c = by_id[cid]
                    self.index.add(cid, vec, at, set(c.tags or []), set(c.artists or []))
            for extra, cid in await s.execute(
                select(NewsItem.extra, NewsItem.cluster_id)
                .where(NewsItem.tg_channel.is_not(None), NewsItem.cluster_id.is_not(None))
                .where(NewsItem.fetched_at >= cutoff)
            ):
                for link in (extra or {}).get("links", []):
                    cl = canonicalize_url(link)
                    if cl:
                        self.link_index.setdefault(cl, cid)
            pub_cut = now - timedelta(days=self.settings.republish_lookback_days)
            pubs = (
                await s.execute(
                    select(Cluster).where(Cluster.status == ClusterStatus.published)
                    .where(Cluster.updated_at >= pub_cut).where(Cluster.centroid_model == model)
                )
            ).scalars().all()
            for c in pubs:
                vec = from_bytes(c.centroid)
                if vec is not None:
                    self.published.add(c.id, vec, c.updated_at, set(c.artists or []))
        self._loaded = True
        log.info("Индекс сюжетов: %d активных, %d опубликованных (порог %.2f)",
                 len(self.index), len(self.published.entries), self.index.threshold)

    def mark_published(self, cluster: Cluster) -> None:
        vec = from_bytes(cluster.centroid)
        if vec is not None and cluster.centroid_model == self.embedder.model_id:
            self.published.add(cluster.id, vec, utcnow(), set(cluster.artists or []))

    # ── основной вход ─────────────────────────────────────────
    async def ingest(self, src: SourceRef, items: list[RawItem], *, first_run: bool = False) -> IngestReport:
        report = IngestReport(received=len(items))
        if not items:
            return report
        async with self.lock:
            if not self._loaded:
                await self.load_index()

            # A. что уже знаем
            uniq: dict[str, RawItem] = {}
            for it in items:
                uniq.setdefault(it.external_id, it)
            async with self.db.session() as s:
                known = set(
                    (await s.execute(
                        select(NewsItem.external_id)
                        .where(NewsItem.source_id == src.id)
                        .where(NewsItem.external_id.in_(list(uniq)))
                    )).scalars()
                )
            fresh = [it for k, it in uniq.items() if k not in known]
            if not fresh:
                return report
            if first_run:  # первый прогон источника: не тащим весь архив ленты
                fresh.sort(key=lambda i: i.published_at or i.fetched_at, reverse=True)
                fresh = fresh[: self.settings.first_run_max_items]
            report.new = len(fresh)

            # B. подготовка без БД
            prepared = [self._prepare(src, it) for it in fresh]
            to_embed = [p for p in prepared if p.relevant]
            if to_embed:
                vecs = await self.embedder.embed(
                    [embedding_text(p.raw.title, p.raw.summary, p.raw.content) for p in to_embed]
                )
                for p, v in zip(to_embed, vecs):
                    p.vec = v
            report.relevant = len(to_embed)

            # C. запись
            async with self.db.session() as s:
                touched: set[int] = set()
                created: set[int] = set()
                rows: list[tuple[_Prepared, NewsItem]] = []
                for p in prepared:
                    row = self._row(src, p)
                    s.add(row)
                    rows.append((p, row))
                await s.flush()

                exact = await self._exact_matches(s, [p for p, _ in rows if p.relevant])
                for p, row in sorted(rows, key=lambda pr: pr[0].eff_at):
                    if not p.relevant or p.vec is None:
                        continue
                    cid = exact.get(p.canonical or "") or exact.get(p.chash)
                    sim = 1.0 if cid else None
                    if cid is None:
                        m = self.index.match(p.vec, p.eff_at, set(p.tags), set(p.artists))
                        cid, sim = m.cluster_id, (m.similarity if m.cluster_id else None)
                    if cid is None:
                        cluster = self._new_cluster(src, p)
                        s.add(cluster)
                        await s.flush()
                        cid = cluster.id
                        created.add(cid)
                    row.cluster_id = cid
                    row.similarity = sim
                    self.index.add(cid, p.vec, p.eff_at, set(p.tags), set(p.artists))
                    touched.add(cid)
                    # следующая публикация с тем же URL/текстом в этом же батче — туда же
                    if p.canonical:
                        exact[p.canonical] = cid
                    exact[p.chash] = cid
                    if p.raw.tg_channel:
                        for link in p.raw.extra.get("links", []):
                            cl = canonicalize_url(link)
                            if cl:
                                self.link_index.setdefault(cl, cid)
                                exact.setdefault(cl, cid)

                await s.flush()
                await self._refresh(s, touched, created)
                await s.commit()
            report.created = sorted(created)
            report.updated = sorted(touched - created)
            self.index.prune(utcnow())
            return report

    def _prepare(self, src: SourceRef, it: RawItem) -> _Prepared:
        text = it.text_blob()
        eff = effective_time(
            it.published_at, it.date_source, it.fetched_at, it.date_has_tz, self.settings.max_item_age_days
        )
        relevant, reason = True, None
        if src.cfg.filter == "hiphop" and not self.tagger.is_hiphop(text):
            relevant, reason = False, "not_hiphop"
        if src.cfg.exclude_pattern and re.search(src.cfg.exclude_pattern, it.title or "", re.I):
            relevant, reason = False, "excluded"
        if "archive" in eff.flags:
            relevant, reason = False, "archive"
        artists = [a.name for a in self.tagger.find_artists(text)]
        return _Prepared(
            raw=it,
            canonical=canonicalize_url(it.extra.get("canonical") or it.url),
            chash=content_hash(it.title, it.summary or it.content),
            relevant=relevant,
            reason=reason,
            lang=it.lang or detect_lang(text),
            tags=self.tagger.tag(text) if relevant else [],
            artists=artists,
            eff_at=eff.at,
            eff_conf=eff.confidence,
            flags=eff.flags,
        )

    def _row(self, src: SourceRef, p: _Prepared) -> NewsItem:
        it = p.raw
        extra = dict(it.extra)
        if p.flags:
            extra["date_flags"] = p.flags
        if p.reason:
            extra["skip_reason"] = p.reason
        if p.artists:
            extra["artists"] = p.artists
        return NewsItem(
            source_id=src.id,
            external_id=it.external_id,
            url=it.url,
            canonical_url=p.canonical,
            title=truncate(it.title, 500),
            summary=it.summary,
            content=it.content,
            author=it.author,
            image_url=it.image_url,
            lang=p.lang,
            scene=src.scene,
            published_at=it.published_at,
            date_source=it.date_source,
            date_confidence=p.eff_conf,
            fetched_at=it.fetched_at,
            effective_at=p.eff_at,
            content_hash=p.chash,
            embedding=to_bytes(p.vec) if p.vec is not None else None,
            embedding_model=self.embedder.model_id if p.vec is not None else None,
            is_relevant=p.relevant,
            tags=p.tags,
            tg_channel=it.tg_channel,
            tg_message_id=it.tg_message_id,
            tg_views=it.tg_views,
            fwd_from=it.fwd_from,
            extra=extra,
        )

    async def _exact_matches(self, s, prepared: list[_Prepared]) -> dict[str, int]:
        """Тот же URL или тот же текст уже лежит в активном сюжете (другой источник)."""
        out: dict[str, int] = {}
        cut = utcnow() - timedelta(hours=self.settings.cluster_window_hours)
        for p in prepared:  # статья, на которую уже ссылался инсайд-пост
            if p.canonical and p.canonical in self.link_index:
                out[p.canonical] = self.link_index[p.canonical]
        urls = [p.canonical for p in prepared if p.canonical]
        hashes = [p.chash for p in prepared]
        if urls:
            for url, cid in await s.execute(
                select(NewsItem.canonical_url, NewsItem.cluster_id)
                .where(NewsItem.canonical_url.in_(urls)).where(NewsItem.cluster_id.is_not(None))
                .where(NewsItem.fetched_at >= cut)
            ):
                out.setdefault(url, cid)
        if hashes:
            for h, cid in await s.execute(
                select(NewsItem.content_hash, NewsItem.cluster_id)
                .where(NewsItem.content_hash.in_(hashes)).where(NewsItem.cluster_id.is_not(None))
                .where(NewsItem.fetched_at >= cut)
            ):
                out.setdefault(h, cid)
        # TG-пост со ссылкой на уже известную статью → её сюжет
        tg_links = {
            p.chash: [canonicalize_url(x) for x in p.raw.extra.get("links", [])]
            for p in prepared if p.raw.tg_channel and p.raw.extra.get("links")
        }
        all_links = [x for xs in tg_links.values() for x in xs if x]
        if all_links:
            found = dict(
                (u, cid) for u, cid in await s.execute(
                    select(NewsItem.canonical_url, NewsItem.cluster_id)
                    .where(NewsItem.canonical_url.in_(all_links)).where(NewsItem.cluster_id.is_not(None))
                )
            )
            for chash, links in tg_links.items():
                for link in links:
                    if link in found:
                        out.setdefault(chash, found[link])
                        break
        return out

    def _new_cluster(self, src: SourceRef, p: _Prepared) -> Cluster:
        return Cluster(
            title=truncate(p.raw.title, 300),
            scene=src.scene,
            scenes=[src.scene.value],
            tags=p.tags,
            artists=p.artists,
            origin_published_at=p.eff_at,
            origin_date_confidence=p.eff_conf,
            first_seen_at=p.raw.fetched_at,
            last_item_at=p.raw.fetched_at,
            items_count=0,
            sources_count=0,
        )

    async def _refresh(self, s, cluster_ids: set[int], created: set[int]) -> None:
        if not cluster_ids:
            return
        now = utcnow()
        clusters = {
            c.id: c for c in (await s.execute(select(Cluster).where(Cluster.id.in_(list(cluster_ids))))).scalars()
        }
        res = await s.execute(
            select(NewsItem, Source.kind, Source.scene, Source.name)
            .join(Source, Source.id == NewsItem.source_id)
            .where(NewsItem.cluster_id.in_(list(cluster_ids)))
            .where(NewsItem.is_relevant.is_(True))
        )
        members: dict[int, list[tuple[NewsItem, SourceKind, Scene, str]]] = {}
        for item, kind, scene, name in res:
            members.setdefault(item.cluster_id, []).append((item, kind, scene, name))

        drafted = set(
            (await s.execute(
                select(Post.cluster_id).where(Post.cluster_id.in_(list(cluster_ids)))
                .where(Post.status.in_([PostStatus.draft, PostStatus.scheduled]))
            )).scalars()
        )

        for cid, c in clusters.items():
            ms = members.get(cid, [])
            if not ms:
                continue
            origin = min(ms, key=lambda m: (m[0].effective_at, m[0].fetched_at))
            item0 = origin[0]
            prev_count = c.items_count or 0
            c.origin_item_id = item0.id
            c.origin_source_id = item0.source_id
            c.origin_published_at = item0.effective_at
            c.origin_date_confidence = item0.date_confidence
            c.first_seen_at = min(m[0].fetched_at for m in ms)
            c.last_item_at = max(m[0].fetched_at for m in ms)
            c.items_count = len(ms)
            # независимые источники считаем по изданию: «TMZ» из RSS и из рубрики — один источник
            c.sources_count = len({m[3] for m in ms})
            kinds = {m[1] for m in ms}
            if kinds == {SourceKind.telegram}:
                c.verification = Verification.insider_only
            elif c.sources_count >= 2:
                c.verification = Verification.multi_source
            else:
                c.verification = Verification.single_source

            # заголовок карточки: первоисточник, но СМИ предпочтительнее TG-поста
            media = [m for m in ms if m[1] != SourceKind.telegram]
            title_src = origin if origin[1] != SourceKind.telegram or not media else min(
                media, key=lambda m: m[0].effective_at
            )
            c.title = truncate(title_src[0].title, 300)

            tag_set = {t for m in ms for t in (m[0].tags or [])}
            c.tags = [t for t in self.tagger.tag_names if t in tag_set]
            c.artists = sorted({a for m in ms for a in (m[0].extra or {}).get("artists", [])})
            scenes = Counter(m[2].value for m in ms)
            c.scenes = sorted(scenes)
            c.scene = Scene(self.tagger.artist_scene(c.artists) or origin[2].value)

            c.is_stale, c.stale_reason = assess_staleness(c.origin_published_at, now, self.settings.stale_after_hours)
            first_by_source: dict[str, object] = {}
            for m in ms:
                outlet = m[3]
                first_by_source[outlet] = min(first_by_source.get(outlet, m[0].effective_at), m[0].effective_at)
            breaking = breaking_count(first_by_source, float(self.tagger.bonuses.get("breaking_window_h", 2)))
            c.priority, c.priority_score = self.tagger.score(
                ClusterSignals(
                    tags=c.tags,
                    artists=c.artists,
                    sources_count=c.sources_count,
                    first_items_within_window=breaking,
                    insider_only=c.verification == Verification.insider_only,
                    is_stale=c.is_stale,
                )
            )
            ac = self.index.clusters.get(cid)
            if ac is not None:
                c.centroid = to_bytes(ac.centroid)
                c.centroid_model = self.embedder.model_id
                if c.republish_of_cluster_id is None and c.status != ClusterStatus.published:
                    c.republish_of_cluster_id = self.published.match(ac.centroid, set(c.artists), exclude_id=cid)
            if cid not in created and c.items_count > prev_count and cid in drafted:
                c.needs_refresh = True
            c.updated_at = now

    # ── правки постов в Telegram-каналах ──────────────────────
    async def apply_edit(self, src: SourceRef, it: RawItem) -> int | None:
        async with self.lock, self.db.session() as s:
            row = (
                await s.execute(
                    select(NewsItem).where(NewsItem.source_id == src.id, NewsItem.external_id == it.external_id)
                )
            ).scalar_one_or_none()
            if row is None:
                return None
            row.title = truncate(it.title, 500)
            row.summary = it.summary
            row.extra = {**(row.extra or {}), "edited_at": utcnow().isoformat()}
            cid = row.cluster_id
            if cid:
                c = await s.get(Cluster, cid)
                if c and c.status in (ClusterStatus.drafted, ClusterStatus.scheduled):
                    c.needs_refresh = True
                    c.updated_at = utcnow()
            await s.commit()
            return cid
