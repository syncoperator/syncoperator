"""Минимальный async-клиент Claude Messages API (прямой HTTP-вызов, без SDK).

Структурированный ответ — через `output_config.format` (JSON Schema, GA).
Для claude-sonnet-5-5 принудительный tool_choice и temperature запрещены API (400),
поэтому ни того ни другого здесь нет.
"""

from __future__ import annotations

import asyncio
import json
import logging
import random
from dataclasses import dataclass
from typing import Any

import httpx

log = logging.getLogger(__name__)

API_VERSION = "2023-06-01"
RETRY_STATUSES = {408, 409, 429, 500, 502, 503, 504, 529}


class ClaudeError(RuntimeError):
    def __init__(self, status: int, err_type: str, message: str):
        super().__init__(f"Claude API {status} {err_type}: {message}")
        self.status = status
        self.err_type = err_type


@dataclass(slots=True)
class ClaudeResponse:
    text: str
    stop_reason: str | None
    model: str
    usage: dict[str, Any]
    raw: dict[str, Any]

    def json(self) -> dict[str, Any]:
        return json.loads(self.text)


class ClaudeClient:
    def __init__(
        self,
        api_key: str,
        base_url: str = "https://api.anthropic.com",
        timeout_s: float = 120.0,
        max_retries: int = 3,
        transport: httpx.AsyncBaseTransport | None = None,
    ):
        self.max_retries = max_retries
        self.client = httpx.AsyncClient(
            base_url=base_url.rstrip("/"),
            timeout=httpx.Timeout(timeout_s, connect=15.0),
            headers={
                "x-api-key": api_key,
                "anthropic-version": API_VERSION,
                "content-type": "application/json",
            },
            transport=transport,
        )

    async def close(self) -> None:
        await self.client.aclose()

    async def create(self, body: dict[str, Any]) -> ClaudeResponse:
        last: Exception | None = None
        for attempt in range(self.max_retries + 1):
            try:
                r = await self.client.post("/v1/messages", json=body)
            except (httpx.TimeoutException, httpx.TransportError) as e:
                last = e
                if attempt < self.max_retries:
                    await asyncio.sleep(_backoff(attempt))
                    continue
                raise ClaudeError(0, "network", str(e)) from e
            if r.status_code == 200:
                data = r.json()
                text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
                return ClaudeResponse(
                    text=text,
                    stop_reason=data.get("stop_reason"),
                    model=data.get("model", body.get("model", "")),
                    usage=data.get("usage") or {},
                    raw=data,
                )
            err_type, msg = "http_error", r.text[:500]
            try:
                err = r.json().get("error", {})
                err_type, msg = err.get("type", err_type), err.get("message", msg)
            except ValueError:
                pass
            if r.status_code in RETRY_STATUSES and attempt < self.max_retries:
                ra = r.headers.get("retry-after")
                delay = float(ra) if ra and ra.replace(".", "", 1).isdigit() else _backoff(attempt)
                log.warning("Claude API %s (%s), повтор через %.1fs", r.status_code, err_type, delay)
                await asyncio.sleep(min(delay, 60))
                continue
            raise ClaudeError(r.status_code, err_type, msg)
        raise ClaudeError(0, "retries_exhausted", str(last))


def _backoff(attempt: int) -> float:
    return min(30.0, 2.0 * 2**attempt) + random.random()
