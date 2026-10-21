"""Воркер: сбор каждые 15 минут, Telegram live, дедуп, авто-генерация, отложенный постинг.

Запуск: python -m sqandal.worker
"""

from __future__ import annotations

import asyncio
import logging
import signal
import time
from datetime import timedelta

from sqlalchemy import delete, select, update

from .bootstrap import Core, build_core, load_and_validate_sources, sync_sources
from .collectors import CollectResult, SourceConfig
from .collectors.browser import BrowserPool
from .collectors.http import FetchError, RobotsDisallowed
from .collectors.rss import collect_rss
from .collectors.telegram import TelegramCollector
from .collectors.web import collect_web
from .models import (
    PRIORITY_RANK,
    Cluster,
    ClusterStatus,
    FetchRun,
    NewsItem,
    Priority,
    Source,
    SourceKind,
)
from .processing.embeddings import make_embedder
from .processing.freshness import assess_staleness
from .processing.ingest import IngestPipeline, IngestReport, SourceRef
from .services import NewsroomError
from .settings import Settings, get_settings
from .timeutil import utcnow

log = logging.getLogger("sqandal.worker")


class Worker:
    def __init__(self, core: Core, embedder=None):
        self.core = core
        self.settings: Settings = core.settings
        self.embedder = embedder or make_embedder(self.settings)
        self.pipeline = IngestPipeline(core.db, self.settings, self.embedder, core.tagger)
        self.browser = BrowserPool(self.settings) if self.settings.browser_enabled else None
        self.refs: dict[str, SourceRef] = {}
        self.gen_queue: asyncio.Queue[int] = asyncio.Queue()
        self._queued: set[int] = set()
        self._stop = asyncio.Event()
        self._source_sem = asyncio.Semaphore(4)
        self.telegram: TelegramCollector | None = None

    # ── источники ─────────────────────────────────────────────
    async def load_sources(self) -> None:
        configs = load_and_validate_sources(self.settings)
        self.refs = await sync_sources(self.core.db, configs)
        log.info("Источников: %d (включено %d)", len(self.refs), sum(r.cfg.enabled for r in self.refs.values()))

    async def _is_known_url(self, source_id: int, canonical: str) -> bool:
        async with self.core.db.session() as s:
            return (await s.execute(
                select(NewsItem.id).where(NewsItem.source_id == source_id, NewsItem.canonical_url == canonical).limit(1)
            )).first() is not None

    async def collect_source(self, ref: SourceRef) -> IngestReport | None:
        cfg: SourceConfig = ref.cfg
        started = utcnow()
        t0 = time.monotonic()
        async with self.core.db.session() as s:
            row = await s.get(Source, ref.id)
            state, etag, lm = dict(row.state or {}), row.etag, row.last_modified
            first_run = row.items_total == 0
        result: CollectResult | None = None
        error: str | None = None
        report: IngestReport | None = None
        try:
            async with self._source_sem:
                if cfg.kind == SourceKind.rss:
                    result = await collect_rss(cfg, state, etag, lm, self.core.fetcher)
                elif cfg.kind == SourceKind.web:
                    result = await collect_web(
                        cfg, self.core.fetcher, self.browser,
                        lambda url: self._is_known_url(ref.id, url),
                        max_links=self.settings.max_links_per_listing,
                        max_new=self.settings.max_new_articles_per_run,
                    )
                else:
                    return None
            report = await self.pipeline.ingest(ref, result.items, first_run=first_run)
            await self._after_ingest(report)
        except (FetchError, RobotsDisallowed) as e:
            error = str(e)
        except Exception as e:  # noqa: BLE001 — один источник не роняет воркер
            log.exception("%s: сбой сбора", ref.key)
            error = f"{type(e).__name__}: {e}"
        await self._record_run(ref, started, t0, result, report, error)
        if error:
            log.warning("%s: %s", ref.key, error)
        elif report:
            log.info("%s: найдено %d, новых %d, релевантных %d, сюжетов +%d/~%d", ref.key, report.received,
                     report.new, report.relevant, len(report.created), len(report.updated))
        return report

    async def _record_run(self, ref, started, t0, result, report, error) -> None:
        async with self.core.db.session() as s:
            row = await s.get(Source, ref.id)
            row.last_fetch_at = utcnow()
            if error:
                row.last_error = error[:1000]
                row.consecutive_failures += 1
                if "discovered_feed_url" in (row.state or {}) and row.consecutive_failures >= 3:
                    row.state = {k: v for k, v in row.state.items() if k != "discovered_feed_url"}
            else:
                row.last_success_at = utcnow()
                row.last_error = None
                row.consecutive_failures = 0
                if result is not None:
                    row.etag, row.last_modified = result.etag, result.last_modified
                    if result.state_updates:
                        row.state = {**(row.state or {}), **result.state_updates}
                if report:
                    row.items_total += report.new
            s.add(FetchRun(
                source_id=ref.id, started_at=started, finished_at=utcnow(), ok=error is None,
                items_seen=(result.seen if result else 0), items_new=(report.new if report else 0),
                error=error, duration_ms=int((time.monotonic() - t0) * 1000),
            ))
            await s.commit()

    def _due(self, ref: SourceRef, last_fetch, failures: int) -> bool:
        if not ref.cfg.enabled or ref.cfg.kind == SourceKind.telegram:
            return False
        if last_fetch is None:
            return True
        interval = ref.cfg.interval_min or self.settings.collect_interval_min
        # после ошибок — экспоненциальная пауза, максимум 2 часа
        interval = min(interval * (2 ** min(failures, 3)), 120) if failures else interval
        return utcnow() - last_fetch >= timedelta(minutes=interval)

    async def collect_loop(self) -> None:
        while not self._stop.is_set():
            try:
                req = await self.core.db.get_state("collect_request")
                handled = await self.core.db.get_state("collect_handled")
                forced = req.get("at") and req.get("at") != handled.get("at")
                async with self.core.db.session() as s:
                    rows = {r.key: r for r in (await s.execute(select(Source))).scalars()}
                due = []
                for key, ref in self.refs.items():
                    r = rows.get(key)
                    if r is None:
                        continue
                    want = self._due(ref, r.last_fetch_at, r.consecutive_failures)
                    if forced and ref.cfg.enabled and ref.cfg.kind != SourceKind.telegram and (
                        not req.get("source") or req.get("source") == key
                    ):
                        want = True
                    if want:
                        due.append(ref)
                if forced:
                    await self.core.db.set_state("collect_handled", {"at": req["at"]})
                if due:
                    await asyncio.gather(*(self.collect_source(r) for r in due))
            except Exception:  # noqa: BLE001
                log.exception("collect_loop")
            await self._sleep(30)

    # ── Telegram ──────────────────────────────────────────────
    async def telegram_loop(self) -> None:
        tg_refs = [r for r in self.refs.values() if r.cfg.kind == SourceKind.telegram and r.cfg.enabled]
        if not tg_refs or not self.settings.telethon_configured:
            if tg_refs:
                log.info("Telegram-каналы в конфиге есть, но Telethon не настроен")
            return
        by_key = {r.key: r for r in tg_refs}
        async with self.core.db.session() as s:
            cursors = {r.key: r.tg_last_message_id for r in (await s.execute(
                select(Source).where(Source.id.in_([x.id for x in tg_refs])))).scalars()}

        async def on_items(cfg: SourceConfig, items):
            ref = by_key[cfg.key]
            t0, started = time.monotonic(), utcnow()
            report = await self.pipeline.ingest(ref, items)
            await self._after_ingest(report)
            res = CollectResult(items=items, seen=len(items))
            await self._record_run(ref, started, t0, res, report, None)

        async def on_cursor(cfg: SourceConfig, msg_id: int):
            async with self.core.db.session() as s:
                await s.execute(update(Source).where(Source.id == by_key[cfg.key].id)
                                .values(tg_last_message_id=msg_id))
                await s.commit()

        async def on_edit(cfg: SourceConfig, item):
            await self.pipeline.apply_edit(by_key[cfg.key], item)

        while not self._stop.is_set():
            self.telegram = TelegramCollector(self.settings, [r.cfg for r in tg_refs], cursors, on_items,
                                              on_cursor, on_edit)
            try:
                await self.telegram.start()
                await self.telegram.run_forever()
            except Exception:  # noqa: BLE001
                log.exception("Telegram-парсер упал, перезапуск через 60 с")
            finally:
                await self.telegram.stop()
            await self._sleep(60)

    # ── генерация ─────────────────────────────────────────────
    def _auto_ok(self, c: Cluster) -> bool:
        lvl = self.settings.auto_generate_min_priority
        if lvl == "OFF" or self.core.newsroom.generator is None:
            return False
        if PRIORITY_RANK[c.priority] < PRIORITY_RANK[Priority(lvl)]:
            return False
        if c.is_stale or c.status != ClusterStatus.new:
            return False
        age_h = (utcnow() - c.origin_published_at).total_seconds() / 3600
        return age_h <= self.settings.auto_generate_max_age_hours

    async def _after_ingest(self, report: IngestReport) -> None:
        ids = report.created + report.updated
        if not ids:
            return
        async with self.core.db.session() as s:
            clusters = list((await s.execute(select(Cluster).where(Cluster.id.in_(ids)))).scalars())
        for c in clusters:
            if self._auto_ok(c) and c.id not in self._queued:
                self._queued.add(c.id)
                await self.gen_queue.put(c.id)

    async def gen_worker(self, n: int) -> None:
        while not self._stop.is_set():
            cid = await self.gen_queue.get()
            try:
                async with self.core.db.session() as s:
                    c = await s.get(Cluster, cid)
                if c is None or c.status != ClusterStatus.new:
                    continue
                post = await self.core.newsroom.generate_post(cid)
                log.info("Сгенерирован пост v%d для сюжета #%d [%s]", post.version, cid, c.priority.value)
                if c.priority == Priority.CRITICAL:
                    async with self.core.db.session() as s:
                        c = await s.get(Cluster, cid)
                    await self.core.newsroom.notify_admin(post, c)
            except NewsroomError as e:
                log.warning("Генерация #%d: %s", cid, e)
            except Exception:  # noqa: BLE001
                log.exception("Генерация #%d упала", cid)
            finally:
                self._queued.discard(cid)
                self.gen_queue.task_done()

    async def requeue_missing(self) -> None:
        """После рестарта: свежие сюжеты без черновика — снова в очередь."""
        cut = utcnow() - timedelta(hours=self.settings.auto_generate_max_age_hours)
        async with self.core.db.session() as s:
            clusters = list((await s.execute(
                select(Cluster).where(Cluster.status == ClusterStatus.new, Cluster.origin_published_at >= cut)
            )).scalars())
        for c in clusters:
            if self._auto_ok(c) and c.id not in self._queued:
                self._queued.add(c.id)
                await self.gen_queue.put(c.id)

    # ── отложенные посты, обслуживание, админ-чат ────────────
    async def scheduler_loop(self) -> None:
        while not self._stop.is_set():
            try:
                for pid in await self.core.newsroom.due_posts():
                    try:
                        p = await self.core.newsroom.publish_post(pid, confirm_stale=True)
                        log.info("Отложенный пост #%d опубликован (msg %s)", pid, p.tg_message_id)
                        await self._mark_published_in_index(p.cluster_id)
                    except NewsroomError as e:
                        log.error("Отложенный пост #%d: %s", pid, e)
            except Exception:  # noqa: BLE001
                log.exception("scheduler_loop")
            await self._sleep(20)

    async def _mark_published_in_index(self, cluster_id: int) -> None:
        async with self.core.db.session() as s:
            c = await s.get(Cluster, cluster_id)
            if c:
                self.pipeline.mark_published(c)

    async def housekeeping_loop(self) -> None:
        seen_published: set[int] = set()
        while not self._stop.is_set():
            try:
                now = utcnow()
                async with self.core.db.session() as s:
                    # протухание активных сюжетов пересчитываем по часам
                    for c in (await s.execute(select(Cluster).where(
                        Cluster.status.in_([ClusterStatus.new, ClusterStatus.drafted, ClusterStatus.scheduled]),
                        Cluster.is_stale.is_(False),
                    ))).scalars():
                        stale, reason = assess_staleness(c.origin_published_at, now, self.settings.stale_after_hours)
                        if stale:
                            c.is_stale, c.stale_reason, c.updated_at = True, reason, now
                    # опубликованные из дашборда → в индекс повторов
                    pubs = (await s.execute(select(Cluster).where(
                        Cluster.status == ClusterStatus.published,
                        Cluster.updated_at >= now - timedelta(days=self.settings.republish_lookback_days),
                    ))).scalars()
                    for c in pubs:
                        if c.id not in seen_published:
                            self.pipeline.mark_published(c)
                            seen_published.add(c.id)
                    # нерелевантный шум старше 7 дней — чистим
                    old = now - timedelta(days=7)
                    await s.execute(delete(NewsItem).where(NewsItem.is_relevant.is_(False), NewsItem.fetched_at < old))
                    await s.commit()
                self.pipeline.published.prune(now)
            except Exception:  # noqa: BLE001
                log.exception("housekeeping_loop")
            await self._sleep(300)

    async def admin_poll_loop(self) -> None:
        bot = self.core.bot
        if not (bot and self.settings.tg_admin_mode == "polling" and self.settings.tg_admin_chat_id):
            return
        try:
            await bot.delete_webhook()  # polling и webhook взаимоисключающие
        except Exception:  # noqa: BLE001
            pass
        offset = None
        while not self._stop.is_set():
            try:
                updates = await bot.get_updates(offset, timeout=50)
                for u in updates or []:
                    offset = u["update_id"] + 1
                    if u.get("callback_query"):
                        asyncio.create_task(self.core.newsroom.handle_callback(u["callback_query"]))
            except Exception as e:  # noqa: BLE001
                log.warning("getUpdates: %s", e)
                await self._sleep(10)

    async def _sleep(self, seconds: float) -> None:
        try:
            await asyncio.wait_for(self._stop.wait(), timeout=seconds)
        except asyncio.TimeoutError:
            pass

    def stop(self) -> None:
        self._stop.set()

    async def run(self) -> None:
        await self.load_sources()
        await self.pipeline.load_index()
        await self.requeue_missing()
        tasks = [
            asyncio.create_task(self.collect_loop(), name="collect"),
            asyncio.create_task(self.telegram_loop(), name="telegram"),
            asyncio.create_task(self.scheduler_loop(), name="scheduler"),
            asyncio.create_task(self.housekeeping_loop(), name="housekeeping"),
            asyncio.create_task(self.admin_poll_loop(), name="admin"),
        ]
        tasks += [asyncio.create_task(self.gen_worker(i), name=f"gen{i}") for i in range(self.settings.llm_concurrency)]
        await self._stop.wait()
        for t in tasks:
            t.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)
        if self.browser:
            await self.browser.close()


async def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    settings = get_settings()
    core = await build_core(settings)
    worker = Worker(core)
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, worker.stop)
        except NotImplementedError:  # Windows
            pass
    try:
        await worker.run()
    finally:
        await core.close()


if __name__ == "__main__":
    asyncio.run(main())
