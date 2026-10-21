"""Telegram Bot API: публикация в канал, админ-чат с кнопками, вебхук."""

from __future__ import annotations

import asyncio
import logging
from typing import Any

import httpx

log = logging.getLogger(__name__)


class TelegramError(RuntimeError):
    def __init__(self, code: int, description: str, method: str):
        super().__init__(f"Bot API {method}: {code} {description}")
        self.code = code
        self.description = description


class TelegramBot:
    def __init__(self, token: str, transport: httpx.AsyncBaseTransport | None = None,
                 base_url: str = "https://api.telegram.org"):
        self._token = token
        self.client = httpx.AsyncClient(base_url=f"{base_url}/bot{token}", timeout=httpx.Timeout(65.0, connect=10.0),
                                        transport=transport)

    async def close(self) -> None:
        await self.client.aclose()

    async def call(self, method: str, **params: Any) -> Any:
        payload = {k: v for k, v in params.items() if v is not None}
        for attempt in range(4):
            try:
                r = await self.client.post(f"/{method}", json=payload)
            except (httpx.TimeoutException, httpx.TransportError) as e:
                if attempt < 3:
                    await asyncio.sleep(1.5 * (attempt + 1))
                    continue
                # не светим токен в логах: str(e) может содержать URL
                raise TelegramError(0, type(e).__name__, method) from None
            try:
                data = r.json()
            except ValueError:
                raise TelegramError(r.status_code, r.text[:200], method) from None
            if data.get("ok"):
                return data.get("result")
            code = int(data.get("error_code", r.status_code))
            desc = str(data.get("description", ""))
            retry_after = (data.get("parameters") or {}).get("retry_after")
            if code == 429 and retry_after and attempt < 3:
                log.warning("Bot API flood-limit, ждём %ss", retry_after)
                await asyncio.sleep(float(retry_after) + 0.5)
                continue
            if code >= 500 and attempt < 3:
                await asyncio.sleep(1.5 * (attempt + 1))
                continue
            raise TelegramError(code, desc, method)
        raise TelegramError(0, "retries_exhausted", method)

    async def get_me(self) -> dict:
        return await self.call("getMe")

    async def send_message(
        self,
        chat_id: str | int,
        text: str,
        *,
        parse_mode: str = "HTML",
        link_preview_url: str | None = None,
        disable_preview: bool = False,
        reply_markup: dict | None = None,
        disable_notification: bool = False,
    ) -> dict:
        if disable_preview:
            preview = {"is_disabled": True}
        elif link_preview_url:
            preview = {"url": link_preview_url, "prefer_large_media": True, "show_above_text": False}
        else:
            preview = None
        return await self.call(
            "sendMessage",
            chat_id=chat_id,
            text=text,
            parse_mode=parse_mode,
            link_preview_options=preview,
            reply_markup=reply_markup,
            disable_notification=disable_notification or None,
        )

    async def edit_message_text(self, chat_id: str | int, message_id: int, text: str, *,
                                parse_mode: str = "HTML", reply_markup: dict | None = None) -> Any:
        return await self.call(
            "editMessageText", chat_id=chat_id, message_id=message_id, text=text, parse_mode=parse_mode,
            reply_markup=reply_markup, link_preview_options={"is_disabled": True},
        )

    async def answer_callback_query(self, callback_query_id: str, text: str | None = None, show_alert: bool = False):
        return await self.call("answerCallbackQuery", callback_query_id=callback_query_id, text=text,
                               show_alert=show_alert or None)

    async def set_webhook(self, url: str, secret_token: str) -> Any:
        return await self.call("setWebhook", url=url, secret_token=secret_token,
                               allowed_updates=["callback_query"], drop_pending_updates=True)

    async def delete_webhook(self) -> Any:
        return await self.call("deleteWebhook", drop_pending_updates=False)

    async def get_updates(self, offset: int | None, timeout: int = 50) -> list[dict]:
        return await self.call("getUpdates", offset=offset, timeout=timeout, allowed_updates=["callback_query"])


def message_link(chat_id: str | int, message_id: int, username: str | None = None) -> str | None:
    """Ссылка на пост: публичный канал — t.me/username/id, приватный — t.me/c/<id>/<msg>."""
    cid = str(chat_id)
    if username:
        return f"https://t.me/{username.lstrip('@')}/{message_id}"
    if cid.startswith("@"):
        return f"https://t.me/{cid[1:]}/{message_id}"
    if cid.startswith("-100"):
        return f"https://t.me/c/{cid[4:]}/{message_id}"
    return None


def admin_keyboard(post_id: int, cluster_id: int) -> dict:
    return {
        "inline_keyboard": [
            [
                {"text": "✅ В канал", "callback_data": f"pub:{post_id}"},
                {"text": "🔥 Жёстче", "callback_data": f"hard:{post_id}"},
            ],
            [{"text": "⏭ Пропустить", "callback_data": f"skip:{cluster_id}"}],
        ]
    }
