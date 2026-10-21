"""FastAPI: REST + SSE для дашборда, вебхук Telegram для кнопок админ-чата.

Запуск: uvicorn sqandal.api.app:app --host 0.0.0.0 --port 8000
"""

from __future__ import annotations

import asyncio
import json
import logging
import secrets
from contextlib import asynccontextmanager
from datetime import timedelta

from fastapi import BackgroundTasks, Depends, FastAPI, Header, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy import case, func, or_, select

from .. import __version__
from ..bootstrap import Core, build_core, load_and_validate_sources, sync_sources
from ..models import (
    PRIORITY_RANK,
    Cluster,
    ClusterStatus,
    NewsItem,
    Post,
    PostStatus,
    Priority,
    Scene,
    Source,
)
from ..services import NewsroomError, StaleConfirmationRequired
from ..settings import Settings, get_settings
from ..timeutil import utcnow
from .schemas import (
    ConfigOut,
    GenerateIn,
    ItemOut,
    OriginOut,
    PostBrief,
    PostOut,
    PostPatch,
    PublishIn,
    ScheduleIn,
    SourceOut,
    StatsOut,
    StoryCard,
    StoryDetail,
    StoryList,
)

log = logging.getLogger(__name__)

ACTIVE = [ClusterStatus.new, ClusterStatus.drafted, ClusterStatus.scheduled]


def create_app(settings: Settings | None = None, core_factory=None) -> FastAPI:
    settings = settings or get_settings()

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        core: Core = await (core_factory(settings) if core_factory else build_core(settings))
        try:
            await sync_sources(core.db, load_and_validate_sources(settings))
        except Exception as e:  # noqa: BLE001 — кривой YAML не должен класть API
            log.error("sources.yaml: %s", e)
        app.state.core = core
        yield
        await core.close()

    app = FastAPI(title="SQANDAL RADAR", version=__version__, lifespan=lifespan)
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origins,
        allow_credentials=False,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    def core_dep(request: Request) -> Core:
        return request.app.state.core

    def auth(authorization: str | None = Header(default=None), token: str | None = Query(default=None)) -> None:
        supplied = None
        if authorization and authorization.lower().startswith("bearer "):
            supplied = authorization[7:].strip()
        elif token:
            supplied = token
        if not supplied or not secrets.compare_digest(supplied, settings.dashboard_token):
            raise HTTPException(401, "Нужен токен дашборда")

    @app.exception_handler(NewsroomError)
    async def _newsroom_error(_req: Request, exc: NewsroomError):
        code = "stale_confirmation_required" if isinstance(exc, StaleConfirmationRequired) else "newsroom_error"
        return JSONResponse({"detail": str(exc), "code": code}, status_code=exc.status)

    # ── сериализация ──────────────────────────────────────────
    def post_out(core: Core, p: Post) -> PostOut:
        out = PostOut.model_validate(p, from_attributes=True)
        out.mode, out.status = p.mode.value, p.status.value
        out.link = core.newsroom.post_link(p)
        return out

    def brief(p: Post) -> PostBrief:
        return PostBrief(id=p.id, version=p.version, mode=p.mode.value, status=p.status.value,
                         scheduled_at=p.scheduled_at, published_at=p.published_at, created_at=p.created_at)

    async def cards(s, clusters: list[Cluster]) -> list[StoryCard]:
        if not clusters:
            return []
        ids = [c.id for c in clusters]
        posts: dict[int, Post] = {}
        for p in (await s.execute(
            select(Post).where(Post.cluster_id.in_(ids), Post.status != PostStatus.superseded)
            .order_by(Post.cluster_id, Post.version.desc())
        )).scalars():
            posts.setdefault(p.cluster_id, p)
        origin_ids = [c.origin_item_id for c in clusters if c.origin_item_id]
        origins = {}
        if origin_ids:
            for item, sname, skind in await s.execute(
                select(NewsItem, Source.name, Source.kind).join(Source, Source.id == NewsItem.source_id)
                .where(NewsItem.id.in_(origin_ids))
            ):
                origins[item.id] = (item, sname, skind)
        names: dict[int, list[str]] = {}
        for cid, sname in await s.execute(
            select(NewsItem.cluster_id, Source.name).join(Source, Source.id == NewsItem.source_id)
            .where(NewsItem.cluster_id.in_(ids), NewsItem.is_relevant.is_(True))
            .group_by(NewsItem.cluster_id, Source.name)
        ):
            names.setdefault(cid, []).append(sname)
        out = []
        for c in clusters:
            o = origins.get(c.origin_item_id)
            origin = OriginOut(
                published_at=c.origin_published_at,
                confidence=c.origin_date_confidence.value,
                source_name=o[1] if o else None,
                source_kind=o[2].value if o else None,
                url=o[0].url if o else None,
                date_source=o[0].date_source.value if o else None,
            )
            p = posts.get(c.id)
            out.append(StoryCard(
                id=c.id, title=c.title, scene=c.scene.value, scenes=list(c.scenes or []), tags=list(c.tags or []),
                artists=list(c.artists or []), priority=c.priority.value, priority_score=c.priority_score,
                status=c.status.value, origin=origin, first_seen_at=c.first_seen_at, last_item_at=c.last_item_at,
                items_count=c.items_count, sources_count=c.sources_count,
                source_names=sorted(names.get(c.id, []), key=lambda n: n != (o[1] if o else None)),
                verification=c.verification.value, is_stale=c.is_stale, stale_reason=c.stale_reason,
                republish_of=c.republish_of_cluster_id, needs_refresh=c.needs_refresh,
                post=brief(p) if p else None, updated_at=c.updated_at,
            ))
        return out

    def story_query(scene, priority, tag, status, q):
        stmt = select(Cluster)
        if scene and scene != "all":
            stmt = stmt.where(Cluster.scene == Scene(scene))
        if priority:
            stmt = stmt.where(Cluster.priority.in_([Priority(x) for x in priority.split(",") if x]))
        if status == "active":
            stmt = stmt.where(Cluster.status.in_(ACTIVE))
        elif status in {"published", "dismissed", "drafted", "scheduled", "new"}:
            stmt = stmt.where(Cluster.status == ClusterStatus(status))
        if q:
            like = f"%{q.strip()}%"
            stmt = stmt.where(or_(Cluster.title.ilike(like), Cluster.id.in_(
                select(NewsItem.cluster_id).where(NewsItem.title.ilike(like))
            )))
        return stmt

    # ── маршруты ──────────────────────────────────────────────
    @app.get("/api/health")
    async def health(core: Core = Depends(core_dep)):
        async with core.db.session() as s:
            await s.execute(select(func.count()).select_from(Source))
        return {"ok": True, "version": __version__, "time": utcnow()}

    @app.get("/api/config", response_model=ConfigOut, dependencies=[Depends(auth)])
    async def config(core: Core = Depends(core_dep)):
        st = core.settings
        return ConfigOut(
            channel=st.tg_channel_id, parse_mode=st.post_parse_mode, brand_line=st.brand_line,
            brand_url=st.brand_url, bullet=st.post_bullet, summary_style=st.post_summary_style,
            show_sources=st.post_show_sources,
            tags=[{"name": r.name, "emoji": r.emoji, "weight": r.weight} for r in core.tagger.rules],
            llm_ready=core.newsroom.generator is not None, bot_ready=core.bot is not None,
            telethon_ready=st.telethon_configured, stale_after_hours=st.stale_after_hours,
            audience_tz=st.audience_tz, version=__version__,
        )

    @app.get("/api/stories", response_model=StoryList, dependencies=[Depends(auth)])
    async def list_stories(
        scene: str = "all",
        priority: str | None = None,
        tag: str | None = None,
        status: str = "active",
        q: str | None = None,
        sort: str = "time",
        hours: int = Query(72, ge=1, le=24 * 30),
        limit: int = Query(80, ge=1, le=300),
        core: Core = Depends(core_dep),
    ):
        stmt = story_query(scene, priority, tag, status, q)
        stmt = stmt.where(Cluster.last_item_at >= utcnow() - timedelta(hours=hours))
        if sort == "priority":
            rank = case(*((Cluster.priority == p, r) for p, r in PRIORITY_RANK.items()), else_=0)
            stmt = stmt.order_by(rank.desc(), Cluster.origin_published_at.desc())
        else:
            stmt = stmt.order_by(Cluster.origin_published_at.desc())
        async with core.db.session() as s:
            clusters = list((await s.execute(stmt.limit(limit * 2 if tag else limit))).scalars())
            if tag:  # JSON-массив одинаково фильтруется в PG и SQLite — на стороне Python
                clusters = [c for c in clusters if tag in (c.tags or [])][:limit]
            return StoryList(stories=await cards(s, clusters), server_time=utcnow())

    @app.get("/api/stories/{cluster_id}", response_model=StoryDetail, dependencies=[Depends(auth)])
    async def get_story(cluster_id: int, core: Core = Depends(core_dep)):
        async with core.db.session() as s:
            c = await s.get(Cluster, cluster_id)
            if c is None:
                raise HTTPException(404, "Сюжет не найден")
            card = (await cards(s, [c]))[0]
            items = []
            for it, sname, skind in await s.execute(
                select(NewsItem, Source.name, Source.kind).join(Source, Source.id == NewsItem.source_id)
                .where(NewsItem.cluster_id == cluster_id, NewsItem.is_relevant.is_(True))
                .order_by(NewsItem.effective_at)
            ):
                items.append(ItemOut(
                    id=it.id, source_name=sname, source_kind=skind.value, url=it.url, title=it.title,
                    summary=(it.summary or "")[:600] or None, published_at=it.published_at,
                    effective_at=it.effective_at, date_source=it.date_source.value,
                    date_confidence=it.date_confidence.value, fetched_at=it.fetched_at, lang=it.lang,
                    tg_views=it.tg_views, fwd_from=it.fwd_from, similarity=it.similarity,
                    date_flags=list((it.extra or {}).get("date_flags", [])),
                ))
            posts = list((await s.execute(
                select(Post).where(Post.cluster_id == cluster_id).order_by(Post.version.desc())
            )).scalars())
            current = next((p for p in posts if p.status != PostStatus.superseded), posts[0] if posts else None)
            return StoryDetail(**card.model_dump(), items=items,
                               current_post=post_out(core, current) if current else None,
                               versions=[brief(p) for p in posts])

    @app.post("/api/stories/{cluster_id}/generate", response_model=PostOut, dependencies=[Depends(auth)])
    async def generate(cluster_id: int, body: GenerateIn, core: Core = Depends(core_dep)):
        p = await core.newsroom.generate_post(cluster_id, body.mode)
        return post_out(core, p)

    @app.post("/api/stories/{cluster_id}/dismiss", response_model=StoryCard, dependencies=[Depends(auth)])
    async def dismiss(cluster_id: int, core: Core = Depends(core_dep)):
        c = await core.newsroom.set_dismissed(cluster_id, True)
        async with core.db.session() as s:
            return (await cards(s, [c]))[0]

    @app.post("/api/stories/{cluster_id}/restore", response_model=StoryCard, dependencies=[Depends(auth)])
    async def restore(cluster_id: int, core: Core = Depends(core_dep)):
        c = await core.newsroom.set_dismissed(cluster_id, False)
        async with core.db.session() as s:
            return (await cards(s, [c]))[0]

    @app.get("/api/posts/{post_id}", response_model=PostOut, dependencies=[Depends(auth)])
    async def get_post(post_id: int, core: Core = Depends(core_dep)):
        async with core.db.session() as s:
            p = await s.get(Post, post_id)
            if p is None:
                raise HTTPException(404, "Пост не найден")
            return post_out(core, p)

    @app.patch("/api/posts/{post_id}", response_model=PostOut, dependencies=[Depends(auth)])
    async def patch_post(post_id: int, body: PostPatch, core: Core = Depends(core_dep)):
        p = await core.newsroom.update_post(post_id, body.model_dump(exclude_none=True))
        return post_out(core, p)

    @app.post("/api/posts/{post_id}/publish", response_model=PostOut, dependencies=[Depends(auth)])
    async def publish(post_id: int, body: PublishIn, core: Core = Depends(core_dep)):
        p = await core.newsroom.publish_post(post_id, body.confirm_stale)
        return post_out(core, p)

    @app.post("/api/posts/{post_id}/schedule", response_model=PostOut, dependencies=[Depends(auth)])
    async def schedule(post_id: int, body: ScheduleIn, core: Core = Depends(core_dep)):
        if body.scheduled_at.tzinfo is None:
            raise HTTPException(422, "scheduled_at должен быть с часовым поясом (ISO 8601)")
        p = await core.newsroom.schedule_post(post_id, body.scheduled_at)
        return post_out(core, p)

    @app.delete("/api/posts/{post_id}/schedule", response_model=PostOut, dependencies=[Depends(auth)])
    async def unschedule(post_id: int, core: Core = Depends(core_dep)):
        p = await core.newsroom.unschedule_post(post_id)
        return post_out(core, p)

    @app.get("/api/sources", response_model=list[SourceOut], dependencies=[Depends(auth)])
    async def sources(core: Core = Depends(core_dep)):
        async with core.db.session() as s:
            rows = (await s.execute(select(Source).order_by(Source.scene, Source.name))).scalars()
            return [SourceOut(
                key=r.key, name=r.name, kind=r.kind.value, scene=r.scene.value, url=r.url, enabled=r.enabled,
                last_fetch_at=r.last_fetch_at, last_success_at=r.last_success_at, last_error=r.last_error,
                consecutive_failures=r.consecutive_failures, items_total=r.items_total,
                note=(r.config or {}).get("note"),
            ) for r in rows]

    @app.post("/api/collect", dependencies=[Depends(auth)])
    async def collect_now(source: str | None = None, core: Core = Depends(core_dep)):
        await core.db.set_state("collect_request", {"at": utcnow().isoformat(), "source": source})
        return {"queued": True}

    @app.get("/api/stats", response_model=StatsOut, dependencies=[Depends(auth)])
    async def stats(core: Core = Depends(core_dep)):
        return await _stats(core)

    async def _stats(core: Core) -> StatsOut:
        since = utcnow() - timedelta(hours=72)
        async with core.db.session() as s:
            base = select(Cluster).where(Cluster.status.in_(ACTIVE), Cluster.last_item_at >= since).subquery()
            by_scene = dict((await s.execute(select(base.c.scene, func.count()).group_by(base.c.scene))).all())
            by_prio = dict((await s.execute(select(base.c.priority, func.count()).group_by(base.c.priority))).all())
            crit = dict((await s.execute(
                select(base.c.scene, func.count()).where(base.c.priority == Priority.CRITICAL).group_by(base.c.scene)
            )).all())
            drafts = (await s.execute(select(func.count()).select_from(Post).where(Post.status == PostStatus.draft))).scalar_one()
            sched = (await s.execute(select(func.count()).select_from(Post).where(Post.status == PostStatus.scheduled))).scalar_one()
            pub = (await s.execute(select(func.count()).select_from(Post).where(
                Post.status == PostStatus.published, Post.published_at >= utcnow() - timedelta(hours=24)))).scalar_one()
            last = (await s.execute(select(func.max(Source.last_success_at)))).scalar_one()
            ok = (await s.execute(select(func.count()).select_from(Source).where(
                Source.enabled.is_(True), Source.consecutive_failures == 0))).scalar_one()
            bad = (await s.execute(select(func.count()).select_from(Source).where(
                Source.enabled.is_(True), Source.consecutive_failures > 0))).scalar_one()
        def key(v):
            return v.value if hasattr(v, "value") else str(v)
        if last is not None and last.tzinfo is None:  # SQLite: агрегат мимо TypeDecorator
            from ..timeutil import UTC
            last = last.replace(tzinfo=UTC)
        elif isinstance(last, str):
            from ..timeutil import parse_datetime
            last = parse_datetime(last)[0]
        return StatsOut(
            by_scene={key(k): v for k, v in by_scene.items()},
            by_priority={key(k): v for k, v in by_prio.items()},
            critical_by_scene={key(k): v for k, v in crit.items()},
            drafts=drafts, scheduled=sched, published_24h=pub, last_collect_at=last,
            sources_ok=ok, sources_failing=bad, server_time=utcnow(),
        )

    @app.get("/api/stream", dependencies=[Depends(auth)])
    async def stream(request: Request, core: Core = Depends(core_dep)):
        """SSE: обновлённые карточки сюжетов в реальном времени (опрос БД каждые 3 с)."""

        async def gen():
            cursor = utcnow()
            ticks = 0
            yield "retry: 5000\n\n"
            while True:
                if await request.is_disconnected():
                    break
                async with core.db.session() as s:
                    changed = list((await s.execute(
                        select(Cluster).where(Cluster.updated_at > cursor).order_by(Cluster.updated_at).limit(100)
                    )).scalars())
                    if changed:
                        cursor = max(c.updated_at for c in changed)
                        for card in await cards(s, changed):
                            yield f"event: story\ndata: {card.model_dump_json()}\n\n"
                if ticks % 10 == 0:
                    st = await _stats(core)
                    yield f"event: stats\ndata: {st.model_dump_json()}\n\n"
                ticks += 1
                yield ": ping\n\n"
                await asyncio.sleep(3)

        return StreamingResponse(gen(), media_type="text/event-stream",
                                 headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

    @app.post("/api/telegram/webhook")
    async def telegram_webhook(
        request: Request,
        background: BackgroundTasks,
        x_telegram_bot_api_secret_token: str | None = Header(default=None),
        core: Core = Depends(core_dep),
    ):
        secret = core.settings.tg_webhook_secret
        if not secret or not x_telegram_bot_api_secret_token or not secrets.compare_digest(
            x_telegram_bot_api_secret_token, secret
        ):
            raise HTTPException(403, "bad secret")
        update = json.loads(await request.body() or b"{}")
        cq = update.get("callback_query")
        if cq:
            background.add_task(core.newsroom.handle_callback, cq)
        return {"ok": True}

    dist = settings.frontend_dist
    if dist.exists() and (dist / "index.html").exists():
        app.mount("/", StaticFiles(directory=str(dist), html=True), name="dashboard")

    return app


app = create_app()
