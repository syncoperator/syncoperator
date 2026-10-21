"""Newsroom — операции редакции поверх БД: генерация, правка, публикация, планирование.

Используется и API (кнопки дашборда), и воркером (авто-генерация, отложенные посты,
кнопки в админ-чате Telegram).
"""

from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass
from datetime import datetime, timedelta

from sqlalchemy import func, select, update

from .collectors.article import extract_article
from .collectors.http import Fetcher, FetchError
from .db import Database
from .llm.factcheck import check_numbers
from .llm.generator import GenerationError, GeneratedPost, PostGenerator
from .llm.prompts import SourceDoc
from .models import (
    PRIORITY_RANK,
    Cluster,
    ClusterStatus,
    NewsItem,
    Post,
    PostMode,
    PostStatus,
    Source,
    SourceKind,
    Verification,
)
from .processing.freshness import DAY_ONLY
from .processing.tagging import ClusterSignals, Tagger
from .publish.formatter import PostContent, render_html, render_markdown_v2, render_plain, validate_length
from .publish.formatter import MAX_TEXT_UTF16
from .publish.telegram_bot import TelegramBot, TelegramError, admin_keyboard, message_link
from .settings import Settings
from .textutil import truncate
from .timeutil import fmt_local, utcnow

log = logging.getLogger(__name__)

MODES_ORDER = [PostMode.standard, PostMode.harder, PostMode.hardest]


class NewsroomError(RuntimeError):
    status = 400


class NotFound(NewsroomError):
    status = 404


class Conflict(NewsroomError):
    status = 409


class StaleConfirmationRequired(Conflict):
    """Старый факт / повтор: нужна явная галочка «да, публикую»."""


@dataclass(slots=True)
class MemberRow:
    item: NewsItem
    source_name: str
    source_kind: SourceKind


class Newsroom:
    def __init__(
        self,
        db: Database,
        settings: Settings,
        tagger: Tagger,
        generator: PostGenerator | None,
        bot: TelegramBot | None,
        fetcher: Fetcher | None = None,
    ):
        self.db = db
        self.settings = settings
        self.tagger = tagger
        self.generator = generator
        self.bot = bot
        self.fetcher = fetcher
        self._gen_locks: dict[int, asyncio.Lock] = {}

    # ── загрузка сюжета ───────────────────────────────────────
    async def _members(self, s, cluster_id: int) -> list[MemberRow]:
        res = await s.execute(
            select(NewsItem, Source.name, Source.kind)
            .join(Source, Source.id == NewsItem.source_id)
            .where(NewsItem.cluster_id == cluster_id, NewsItem.is_relevant.is_(True))
            .order_by(NewsItem.effective_at)
        )
        return [MemberRow(i, n, k) for i, n, k in res]

    def _pick_docs(self, members: list[MemberRow]) -> list[MemberRow]:
        """Первоисточник + по одному на остальные источники; СМИ раньше TG-репостов."""
        if not members:
            return []
        picked = [members[0]]
        used = {members[0].source_name}
        rest = sorted(members[1:], key=lambda m: (m.source_kind == SourceKind.telegram, m.item.effective_at))
        for m in rest:
            if len(picked) >= self.settings.llm_max_sources:
                break
            if m.source_name in used:  # одно издание (RSS + рубрика) — один документ
                continue
            used.add(m.source_name)
            picked.append(m)
        return picked

    async def _enrich(self, s, rows: list[MemberRow]) -> None:
        """У RSS часто только тизер. Для генерации докачиваем полный текст статьи."""
        if self.fetcher is None:
            return

        async def one(m: MemberRow) -> None:
            it = m.item
            if m.source_kind == SourceKind.telegram or not it.url:
                return
            if len(it.content or "") >= 600 or (it.extra or {}).get("enrich_failed"):
                return
            try:
                res = await asyncio.wait_for(self.fetcher.get(it.url), timeout=25)
                art = extract_article(res.text, res.url)
                if art.text and len(art.text) > len(it.content or ""):
                    it.content = art.text
            except (FetchError, asyncio.TimeoutError, Exception) as e:  # noqa: BLE001
                log.info("Докачка %s не удалась: %s", it.url, e)
                it.extra = {**(it.extra or {}), "enrich_failed": True}

        await asyncio.gather(*(one(m) for m in rows))
        await s.flush()

    def _docs(self, rows: list[MemberRow]) -> list[SourceDoc]:
        lim = self.settings.llm_source_chars
        docs = []
        for i, m in enumerate(rows, 1):
            it = m.item
            body = it.content or it.summary or ""
            if it.summary and it.content and not it.content.startswith(it.summary[:80]):
                body = it.summary + "\n" + it.content
            docs.append(
                SourceDoc(
                    idx=i,
                    outlet=m.source_name,
                    kind="telegram" if m.source_kind == SourceKind.telegram else "media",
                    url=it.url,
                    published_at=it.effective_at,
                    date_confidence=it.date_confidence.value,
                    title=it.title,
                    text=truncate(body, lim),
                    fwd_from=it.fwd_from,
                    day_only=it.date_source in DAY_ONLY,
                )
            )
        return docs

    def _source_links(self, rows: list[MemberRow]) -> list[dict]:
        media = [m for m in rows if m.source_kind != SourceKind.telegram]
        ordered = media + [m for m in rows if m.source_kind == SourceKind.telegram]
        out, seen = [], set()
        for m in ordered:
            if m.source_name in seen:
                continue
            seen.add(m.source_name)
            out.append({"name": m.source_name, "url": m.item.url})
        return out[:4]

    def _warnings(self, cluster: Cluster, gp_parts: list[str], docs: list[SourceDoc], body_html: str) -> list[str]:
        w: list[str] = []
        now = utcnow()
        allow = {str(now.year), str(now.year - 1), str(now.year + 1)}
        missing = check_numbers(gp_parts, [d.title + "\n" + d.text for d in docs], allow)
        if missing:
            w.append("Числа не найдены в источниках: " + ", ".join(missing[:8]))
        if cluster.verification == Verification.insider_only:
            w.append("Только Telegram-инсайды — подтверждения от СМИ нет")
        elif cluster.verification == Verification.single_source:
            w.append("Один источник")
        if cluster.is_stale:
            w.append(f"Устаревание: {cluster.stale_reason}")
        if cluster.republish_of_cluster_id:
            w.append(f"Похоже на уже опубликованный сюжет #{cluster.republish_of_cluster_id}")
        if cluster.origin_date_confidence.value == "low":
            w.append("Точная дата первоисточника не подтверждена")
        if validate_length(body_html) > MAX_TEXT_UTF16:
            w.append("Пост длиннее 4096 символов — Telegram не примет")
        return w

    def _content(self, p: Post) -> PostContent:
        return PostContent(p.flag_emoji, p.topic_emoji, p.headline, list(p.bullets), p.summary,
                           list(p.sources), list(p.footnotes))

    def _render(self, p: Post) -> None:
        content = self._content(p)
        p.body_html = render_html(content, self.settings)
        p.body_text = render_plain(content, self.settings)

    # ── генерация ─────────────────────────────────────────────
    async def generate_post(self, cluster_id: int, mode: str | PostMode = PostMode.standard) -> Post:
        if self.generator is None:
            raise NewsroomError("LLM не настроен: задай ANTHROPIC_API_KEY")
        mode = PostMode(mode)
        lock = self._gen_locks.setdefault(cluster_id, asyncio.Lock())
        async with lock:
            async with self.db.session() as s:
                cluster = await s.get(Cluster, cluster_id)
                if cluster is None:
                    raise NotFound(f"Сюжет #{cluster_id} не найден")
                members = await self._members(s, cluster_id)
                if not members:
                    raise NewsroomError("У сюжета нет релевантных публикаций")
                rows = self._pick_docs(members)
                await self._enrich(s, rows)
                docs = self._docs(rows)
                prev = (
                    await s.execute(
                        select(Post).where(Post.cluster_id == cluster_id).order_by(Post.version.desc()).limit(1)
                    )
                ).scalar_one_or_none()
                previous = None
                if mode in (PostMode.harder, PostMode.hardest) and prev is not None:
                    previous = {"headline": prev.headline, "bullets": list(prev.bullets), "summary": prev.summary}
                await s.commit()  # сохраним докачанные тексты до долгого вызова LLM

            try:
                gp, resp = await self.generator.generate(docs, cluster.scene.value, mode.value, previous)
            except GenerationError as e:
                raise NewsroomError(str(e)) from e

            async with self.db.session() as s:
                cluster = await s.get(Cluster, cluster_id)
                version = (await s.execute(
                    select(func.coalesce(func.max(Post.version), 0)).where(Post.cluster_id == cluster_id)
                )).scalar_one() + 1
                text_for_notes = " ".join([gp.headline, *gp.bullets, gp.summary])
                post = Post(
                    cluster_id=cluster_id,
                    version=version,
                    mode=mode,
                    status=PostStatus.draft,
                    flag_emoji=gp.flag_emoji,
                    topic_emoji=gp.topic_emoji,
                    headline=gp.headline,
                    bullets=gp.bullets,
                    summary=gp.summary,
                    tags=gp.tags,
                    sources=self._source_links(rows),
                    footnotes=self.tagger.footnotes(text_for_notes),
                    body_html="",
                    body_text="",
                    editor_notes=gp.notes_for_editor or None,
                    is_publishable=gp.is_publishable,
                    event_date=gp.event_date,
                    llm_model=resp.model,
                    input_tokens=int(resp.usage.get("input_tokens") or 0),
                    output_tokens=int(resp.usage.get("output_tokens") or 0),
                    cache_read_tokens=int(resp.usage.get("cache_read_input_tokens") or 0),
                )
                self._render(post)
                post.warnings = self._warnings(cluster, [gp.headline, *gp.bullets, gp.summary], docs, post.body_html)
                if len(gp.bullets) < 3:
                    post.warnings.append("Меньше трёх фактов — источники бедные")
                if not gp.is_publishable:
                    post.warnings.insert(0, "Модель считает, что это не пост: " + (gp.notes_for_editor or "без причины"))
                await s.execute(
                    update(Post).where(Post.cluster_id == cluster_id, Post.status == PostStatus.draft)
                    .values(status=PostStatus.superseded, updated_at=utcnow())
                )
                s.add(post)
                self._merge_llm_tags(cluster, gp)
                if cluster.status in (ClusterStatus.new, ClusterStatus.drafted):
                    cluster.status = ClusterStatus.drafted
                cluster.needs_refresh = False
                cluster.updated_at = utcnow()
                await s.commit()
                await s.refresh(post)
                return post

    def _merge_llm_tags(self, cluster: Cluster, gp: GeneratedPost) -> None:
        new_tags = [t for t in self.tagger.tag_names if t in set(cluster.tags or []) | set(gp.tags)]
        if new_tags == list(cluster.tags or []):
            return
        cluster.tags = new_tags
        prio, score = self.tagger.score(
            ClusterSignals(
                tags=new_tags,
                artists=list(cluster.artists or []),
                sources_count=cluster.sources_count,
                first_items_within_window=1,
                insider_only=cluster.verification == Verification.insider_only,
                is_stale=cluster.is_stale,
            )
        )
        if PRIORITY_RANK[prio] > PRIORITY_RANK[cluster.priority]:
            cluster.priority, cluster.priority_score = prio, max(score, cluster.priority_score)

    # ── ручная правка ─────────────────────────────────────────
    async def update_post(self, post_id: int, fields: dict) -> Post:
        async with self.db.session() as s:
            post = await s.get(Post, post_id)
            if post is None:
                raise NotFound(f"Пост #{post_id} не найден")
            if post.status not in (PostStatus.draft, PostStatus.scheduled, PostStatus.failed):
                raise Conflict(f"Пост в статусе {post.status.value} — правка невозможна")
            for key in ("flag_emoji", "topic_emoji", "headline", "summary"):
                if fields.get(key) is not None:
                    setattr(post, key, str(fields[key]).strip())
            if fields.get("bullets") is not None:
                post.bullets = [b.strip() for b in fields["bullets"] if b and b.strip()][:6]
            if fields.get("sources") is not None:
                post.sources = [x for x in fields["sources"] if x.get("name")][:4]
            post.mode = PostMode.manual
            self._render(post)
            cluster = await s.get(Cluster, post.cluster_id)
            members = await self._members(s, post.cluster_id)
            docs = self._docs(self._pick_docs(members))
            keep = [w for w in post.warnings if not w.startswith("Числа не найдены") and "4096" not in w]
            fresh = [w for w in self._warnings(cluster, [post.headline, *post.bullets, post.summary], docs, post.body_html)
                     if w.startswith("Числа не найдены") or "4096" in w]
            post.warnings = fresh + [w for w in keep if w not in fresh]
            post.updated_at = utcnow()
            await s.commit()
            await s.refresh(post)
            return post

    # ── публикация ────────────────────────────────────────────
    async def publish_post(self, post_id: int, confirm_stale: bool = False) -> Post:
        if self.bot is None:
            raise NewsroomError("Bot API не настроен: задай TG_BOT_TOKEN")
        async with self.db.session() as s:
            post = await s.get(Post, post_id)
            if post is None:
                raise NotFound(f"Пост #{post_id} не найден")
            cluster = await s.get(Cluster, post.cluster_id)
            if post.status == PostStatus.published:
                raise Conflict("Уже опубликован")
            if not confirm_stale and (cluster.is_stale or cluster.republish_of_cluster_id):
                reason = cluster.stale_reason or f"похоже на опубликованный сюжет #{cluster.republish_of_cluster_id}"
                raise StaleConfirmationRequired(f"Проверь актуальность: {reason}")
            claimed = await s.execute(
                update(Post)
                .where(Post.id == post_id, Post.status.in_([PostStatus.draft, PostStatus.scheduled, PostStatus.failed]))
                .values(status=PostStatus.publishing, updated_at=utcnow())
                .returning(Post.id)
            )
            if claimed.scalar_one_or_none() is None:
                raise Conflict("Пост уже публикуется или не в публикуемом статусе")
            await s.commit()
            content = self._content(post)

        chat = self.settings.tg_channel_id
        if self.settings.post_parse_mode == "MarkdownV2":
            text, mode = render_markdown_v2(content, self.settings), "MarkdownV2"
        else:
            text, mode = render_html(content, self.settings), "HTML"
        preview_url = None
        if self.settings.post_link_preview:
            preview_url = next((x.get("url") for x in content.sources if x.get("url") and "t.me/" not in x["url"]), None)
        try:
            msg = await self.bot.send_message(
                chat, text, parse_mode=mode, link_preview_url=preview_url, disable_preview=preview_url is None
            )
        except TelegramError as e:
            async with self.db.session() as s:
                p = await s.get(Post, post_id)
                p.status, p.error, p.updated_at = PostStatus.failed, str(e), utcnow()
                await s.commit()
            raise NewsroomError(f"Telegram отклонил пост: {e.description}") from e

        async with self.db.session() as s:
            p = await s.get(Post, post_id)
            p.status = PostStatus.published
            p.published_at = utcnow()
            p.tg_chat_id = str((msg.get("chat") or {}).get("id") or chat)
            p.tg_message_id = int(msg.get("message_id"))
            p.error = None
            c = await s.get(Cluster, p.cluster_id)
            c.status = ClusterStatus.published
            c.updated_at = utcnow()
            await s.execute(
                update(Post).where(Post.cluster_id == p.cluster_id, Post.id != p.id,
                                   Post.status.in_([PostStatus.draft, PostStatus.scheduled]))
                .values(status=PostStatus.superseded, updated_at=utcnow())
            )
            await s.commit()
            await s.refresh(p)
            return p

    def post_link(self, p: Post) -> str | None:
        if not p.tg_message_id:
            return None
        chat = self.settings.tg_channel_id
        return message_link(chat if str(chat).startswith("@") else (p.tg_chat_id or chat), p.tg_message_id)

    async def schedule_post(self, post_id: int, when: datetime) -> Post:
        if when <= utcnow() + timedelta(seconds=30):
            raise NewsroomError("Время публикации должно быть в будущем")
        async with self.db.session() as s:
            p = await s.get(Post, post_id)
            if p is None:
                raise NotFound(f"Пост #{post_id} не найден")
            if p.status not in (PostStatus.draft, PostStatus.failed, PostStatus.scheduled):
                raise Conflict(f"Нельзя запланировать пост в статусе {p.status.value}")
            p.status, p.scheduled_at, p.updated_at = PostStatus.scheduled, when, utcnow()
            c = await s.get(Cluster, p.cluster_id)
            c.status, c.updated_at = ClusterStatus.scheduled, utcnow()
            await s.commit()
            await s.refresh(p)
            return p

    async def unschedule_post(self, post_id: int) -> Post:
        async with self.db.session() as s:
            p = await s.get(Post, post_id)
            if p is None or p.status != PostStatus.scheduled:
                raise Conflict("Пост не запланирован")
            p.status, p.scheduled_at, p.updated_at = PostStatus.draft, None, utcnow()
            c = await s.get(Cluster, p.cluster_id)
            c.status, c.updated_at = ClusterStatus.drafted, utcnow()
            await s.commit()
            await s.refresh(p)
            return p

    async def due_posts(self, now: datetime | None = None) -> list[int]:
        now = now or utcnow()
        async with self.db.session() as s:
            return list((await s.execute(
                select(Post.id).where(Post.status == PostStatus.scheduled, Post.scheduled_at <= now)
                .order_by(Post.scheduled_at)
            )).scalars())

    async def set_dismissed(self, cluster_id: int, dismissed: bool) -> Cluster:
        async with self.db.session() as s:
            c = await s.get(Cluster, cluster_id)
            if c is None:
                raise NotFound(f"Сюжет #{cluster_id} не найден")
            if dismissed:
                c.status = ClusterStatus.dismissed
            else:
                has_draft = (await s.execute(
                    select(func.count()).select_from(Post)
                    .where(Post.cluster_id == cluster_id, Post.status == PostStatus.draft)
                )).scalar_one()
                c.status = ClusterStatus.drafted if has_draft else ClusterStatus.new
            c.updated_at = utcnow()
            await s.commit()
            await s.refresh(c)
            return c

    # ── админ-чат в Telegram ──────────────────────────────────
    async def notify_admin(self, post: Post, cluster: Cluster) -> None:
        if not (self.bot and self.settings.tg_admin_chat_id and self.settings.tg_admin_mode != "off"):
            return
        head = f"<b>{cluster.priority.value}</b> · сюжет #{cluster.id} · {cluster.sources_count} ист. · " \
               f"первоисточник {fmt_local(cluster.origin_published_at, self.settings.audience_tz)}"
        warn = ("\n⚠️ " + "\n⚠️ ".join(post.warnings[:3])) if post.warnings else ""
        text = f"{head}{warn}\n\n{post.body_html}"
        try:
            await self.bot.send_message(self.settings.tg_admin_chat_id, text, disable_preview=True,
                                        reply_markup=admin_keyboard(post.id, cluster.id))
        except TelegramError as e:
            log.warning("Алерт в админ-чат не ушёл: %s", e)

    async def handle_callback(self, cq: dict) -> None:
        """Кнопки под алертом: pub:<post_id> | hard:<post_id> | skip:<cluster_id>."""
        if self.bot is None:
            return
        chat_id = str(((cq.get("message") or {}).get("chat") or {}).get("id", ""))
        if not self.settings.tg_admin_chat_id or chat_id != str(self.settings.tg_admin_chat_id):
            await self.bot.answer_callback_query(cq["id"], "Нет доступа", show_alert=True)
            return
        action, _, raw_id = str(cq.get("data", "")).partition(":")
        msg = cq.get("message") or {}
        try:
            obj_id = int(raw_id)
            if action == "pub":
                p = await self.publish_post(obj_id, confirm_stale=False)
                await self.bot.answer_callback_query(cq["id"], "Опубликовано ✅")
                link = self.post_link(p) or ""
                await self.bot.edit_message_text(msg["chat"]["id"], msg["message_id"],
                                                 f"✅ Опубликовано {link}\n\n{p.body_html}")
            elif action == "hard":
                async with self.db.session() as s:
                    old = await s.get(Post, obj_id)
                    cid = old.cluster_id if old else None
                    mode = old.mode if old else PostMode.standard
                if cid is None:
                    raise NotFound("Пост не найден")
                nxt = MODES_ORDER[min(MODES_ORDER.index(mode) + 1, 2)] if mode in MODES_ORDER else PostMode.harder
                await self.bot.answer_callback_query(cq["id"], "Переписываю жёстче…")
                p = await self.generate_post(cid, nxt)
                await self.bot.edit_message_text(msg["chat"]["id"], msg["message_id"],
                                                 f"🔥 v{p.version} ({p.mode.value})\n\n{p.body_html}",
                                                 reply_markup=admin_keyboard(p.id, cid))
            elif action == "skip":
                await self.set_dismissed(obj_id, True)
                await self.bot.answer_callback_query(cq["id"], "Пропущено")
                await self.bot.edit_message_text(msg["chat"]["id"], msg["message_id"],
                                                 f"⏭ Пропущено (сюжет #{obj_id})")
            else:
                await self.bot.answer_callback_query(cq["id"], "Неизвестная команда")
        except StaleConfirmationRequired as e:
            await self.bot.answer_callback_query(cq["id"], f"{e} — публикуй из дашборда", show_alert=True)
        except NewsroomError as e:
            await self.bot.answer_callback_query(cq["id"], str(e)[:190], show_alert=True)
