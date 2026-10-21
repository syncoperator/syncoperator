"""Парсер Telegram-каналов (Telethon, user-аккаунт).

Режим работы:
  • live: events.NewMessage / MessageEdited по списку каналов из sources.yaml;
  • догон при старте: всё, что вышло, пока воркер был выключен (по курсору msg_id).

Дата первоисточника: message.date (точная, UTC). Если пост — репост
(fwd_from), берём fwd_from.date — это время ОРИГИНАЛА, а не репоста.
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Awaitable, Callable
from typing import Any

from ..models import DateSource
from ..settings import Settings
from ..textutil import first_line
from ..timeutil import ensure_utc
from .base import RawItem, SourceConfig

log = logging.getLogger(__name__)

OnItems = Callable[[SourceConfig, list[RawItem]], Awaitable[None]]
OnCursor = Callable[[SourceConfig, int], Awaitable[None]]
OnEdit = Callable[[SourceConfig, RawItem], Awaitable[None]]


def message_to_item(msg: Any, channel: str, fwd_name: str | None = None) -> RawItem | None:
    """Telethon Message → RawItem. Без текста (голые фото/альбомы) — None."""
    text = (getattr(msg, "message", None) or "").strip()
    if len(text) < 15:
        return None
    published = ensure_utc(msg.date)
    date_source = DateSource.telegram
    fwd = getattr(msg, "fwd_from", None)
    fwd_from = None
    if fwd is not None:
        if getattr(fwd, "date", None):
            published = ensure_utc(fwd.date)
            date_source = DateSource.telegram_forward
        fwd_from = fwd_name or getattr(fwd, "from_name", None) or getattr(fwd, "post_author", None)
    links: list[str] = []
    try:
        for ent, ent_text in msg.get_entities_text() or []:
            url = getattr(ent, "url", None) or (ent_text if ent.__class__.__name__ == "MessageEntityUrl" else None)
            if url and url.startswith(("http://", "https://")):
                links.append(url)
    except Exception:  # noqa: BLE001 — сущности не критичны
        pass
    return RawItem(
        external_id=f"{channel}:{msg.id}",
        title=first_line(text, 160),
        url=f"https://t.me/{channel}/{msg.id}",
        summary=text[:4000],
        content=None,
        published_at=published,
        date_source=date_source,
        date_has_tz=True,
        tg_channel=channel,
        tg_message_id=int(msg.id),
        tg_views=getattr(msg, "views", None),
        fwd_from=(fwd_from or None) and str(fwd_from)[:256],
        extra={"links": links[:10]} if links else {},
    )


class TelegramCollector:
    def __init__(
        self,
        settings: Settings,
        sources: list[SourceConfig],
        cursors: dict[str, int | None],
        on_items: OnItems,
        on_cursor: OnCursor,
        on_edit: OnEdit | None = None,
    ):
        self.settings = settings
        self.sources = {s.channel.lower(): s for s in sources if s.channel}
        self.cursors = cursors
        self.on_items = on_items
        self.on_cursor = on_cursor
        self.on_edit = on_edit
        self.client = None
        self._fwd_names: dict[int, str] = {}

    async def _fwd_name(self, msg: Any) -> str | None:
        fwd = getattr(msg, "fwd_from", None)
        peer = getattr(fwd, "from_id", None) if fwd else None
        cid = getattr(peer, "channel_id", None) or getattr(peer, "user_id", None)
        if not cid:
            return None
        if cid in self._fwd_names:
            return self._fwd_names[cid]
        name = None
        try:
            ent = await self.client.get_entity(peer)
            name = getattr(ent, "username", None) or getattr(ent, "title", None)
            name = f"@{name}" if getattr(ent, "username", None) else name
        except Exception:  # noqa: BLE001 — приватный/недоступный канал
            name = f"id:{cid}"
        self._fwd_names[cid] = name
        return name

    async def _handle(self, src: SourceConfig, messages: list[Any]) -> None:
        items = []
        max_id = self.cursors.get(src.key) or 0
        for msg in messages:
            item = message_to_item(msg, src.channel, await self._fwd_name(msg))
            if item:
                items.append(item)
            max_id = max(max_id, int(msg.id))
        if items:
            await self.on_items(src, items)
        if max_id and max_id != (self.cursors.get(src.key) or 0):
            self.cursors[src.key] = max_id
            await self.on_cursor(src, max_id)

    async def start(self) -> None:
        from telethon import TelegramClient, events
        from telethon.sessions import StringSession

        if not self.settings.telethon_configured:
            log.info("Telethon не настроен (TG_API_ID/TG_API_HASH/TG_SESSION) — парсер каналов выключен")
            return
        self.client = TelegramClient(
            StringSession(self.settings.tg_session),
            int(self.settings.tg_api_id),
            self.settings.tg_api_hash,
            flood_sleep_threshold=120,
        )
        await self.client.connect()
        if not await self.client.is_user_authorized():
            raise RuntimeError("Telethon-сессия не авторизована: запусти `python -m sqandal.cli tg-login`")

        entities = {}
        for name, src in self.sources.items():
            try:
                entities[name] = await self.client.get_entity(name)
            except Exception as e:  # noqa: BLE001
                log.error("Telegram: канал @%s недоступен: %s", name, e)
        if not entities:
            log.warning("Telegram: нет доступных каналов")
            return

        # догон пропущенного
        for name, ent in entities.items():
            src = self.sources[name]
            last = self.cursors.get(src.key)
            msgs = []
            try:
                if last:
                    async for m in self.client.iter_messages(ent, min_id=last, reverse=True, limit=500):
                        msgs.append(m)
                else:
                    async for m in self.client.iter_messages(ent, limit=self.settings.tg_catchup_limit):
                        msgs.append(m)
                    msgs.reverse()
                await self._handle(src, msgs)
                log.info("Telegram: @%s догнали %d сообщений", name, len(msgs))
            except Exception as e:  # noqa: BLE001
                log.error("Telegram: догон @%s упал: %s", name, e)

        chats = list(entities.values())

        @self.client.on(events.NewMessage(chats=chats))
        async def _on_new(event):  # pragma: no cover - требует сети Telegram
            src = await self._src_for(event)
            if src:
                await self._handle(src, [event.message])

        @self.client.on(events.MessageEdited(chats=chats))
        async def _on_edit(event):  # pragma: no cover - требует сети Telegram
            src = await self._src_for(event)
            if src and self.on_edit:
                item = message_to_item(event.message, src.channel)
                if item:
                    await self.on_edit(src, item)

        log.info("Telegram: слушаем %d каналов", len(chats))

    async def _src_for(self, event) -> SourceConfig | None:
        chat = await event.get_chat()
        uname = (getattr(chat, "username", None) or "").lower()
        return self.sources.get(uname)

    async def run_forever(self) -> None:
        if self.client is None:
            await asyncio.Event().wait()  # выключено — просто спим
            return
        await self.client.run_until_disconnected()

    async def stop(self) -> None:
        if self.client is not None:
            await self.client.disconnect()


async def interactive_login(api_id: int, api_hash: str) -> str:
    """Интерактивный вход (телефон → код → 2FA). Возвращает StringSession."""
    from telethon import TelegramClient
    from telethon.sessions import StringSession

    async with TelegramClient(StringSession(), api_id, api_hash) as client:
        await client.start()
        me = await client.get_me()
        print(f"Вошли как {me.first_name} (@{me.username})")
        return client.session.save()
