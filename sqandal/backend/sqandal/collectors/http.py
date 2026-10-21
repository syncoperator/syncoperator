"""HTTP-клиент сборщиков: вежливые лимиты, robots.txt, условные GET, ретраи."""

from __future__ import annotations

import asyncio
import logging
import random
import re
import time
from dataclasses import dataclass
from urllib.parse import urlsplit
from urllib.robotparser import RobotFileParser

import httpx

from ..settings import Settings

log = logging.getLogger(__name__)

_META_CHARSET = re.compile(rb"""<meta[^>]+charset=["']?([a-zA-Z0-9_\-]+)""", re.I)


class FetchError(RuntimeError):
    pass


class RobotsDisallowed(FetchError):
    pass


@dataclass(slots=True)
class FetchResult:
    url: str
    status: int
    content: bytes
    headers: httpx.Headers

    @property
    def text(self) -> str:
        enc = None
        ctype = self.headers.get("content-type", "")
        if "charset=" in ctype:
            enc = ctype.split("charset=", 1)[1].split(";")[0].strip().strip('"')
        if not enc:  # кодировка из <meta charset> (у части рунета до сих пор windows-1251)
            m = _META_CHARSET.search(self.content[:4096])
            if m:
                enc = m.group(1).decode("ascii", "ignore")
        try:
            return self.content.decode(enc or "utf-8")
        except (LookupError, UnicodeDecodeError):
            return self.content.decode("utf-8", errors="replace")


class Fetcher:
    RETRY_STATUSES = {429, 500, 502, 503, 504}

    def __init__(self, settings: Settings, transport: httpx.AsyncBaseTransport | None = None):
        self.settings = settings
        self.client = httpx.AsyncClient(
            timeout=httpx.Timeout(settings.http_timeout_s, connect=10.0),
            follow_redirects=True,
            headers={
                "User-Agent": settings.http_user_agent,
                "Accept-Language": "ru,de;q=0.9,en;q=0.8",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/rss+xml,*/*;q=0.8",
            },
            transport=transport,
            http2=False,
        )
        self._host_locks: dict[str, asyncio.Semaphore] = {}
        self._robots: dict[str, tuple[float, RobotFileParser | None]] = {}
        self._robots_lock = asyncio.Lock()

    async def close(self) -> None:
        await self.client.aclose()

    def _sem(self, host: str) -> asyncio.Semaphore:
        if host not in self._host_locks:
            self._host_locks[host] = asyncio.Semaphore(2)
        return self._host_locks[host]

    async def allowed(self, url: str) -> bool:
        if not self.settings.respect_robots:
            return True
        parts = urlsplit(url)
        base = f"{parts.scheme}://{parts.netloc}"
        async with self._robots_lock:
            cached = self._robots.get(base)
            if cached and time.monotonic() - cached[0] < 6 * 3600:
                rp = cached[1]
            else:
                rp = None
                try:
                    r = await self.client.get(base + "/robots.txt", timeout=10.0)
                    if r.status_code == 200 and "text" in r.headers.get("content-type", "text"):
                        rp = RobotFileParser()
                        rp.parse(r.text.splitlines())
                except httpx.HTTPError:
                    rp = None  # robots недоступен → не блокируем
                self._robots[base] = (time.monotonic(), rp)
        if rp is None:
            return True
        # своё имя бота в robots никто не прописывает → применяются правила для «*»
        return rp.can_fetch("SQANDALRadar", url)

    async def get(
        self,
        url: str,
        *,
        etag: str | None = None,
        last_modified: str | None = None,
        check_robots: bool = True,
        retries: int = 2,
    ) -> FetchResult:
        if check_robots and not await self.allowed(url):
            raise RobotsDisallowed(f"robots.txt запрещает {url}")
        headers = {}
        if etag:
            headers["If-None-Match"] = etag
        if last_modified:
            headers["If-Modified-Since"] = last_modified
        host = urlsplit(url).netloc
        last_exc: Exception | None = None
        for attempt in range(retries + 1):
            try:
                async with self._sem(host):
                    r = await self.client.get(url, headers=headers)
                if r.status_code in self.RETRY_STATUSES and attempt < retries:
                    delay = _retry_after(r) or (1.5 * 2**attempt + random.random())
                    await asyncio.sleep(min(delay, 30))
                    continue
                if r.status_code == 304:
                    return FetchResult(str(r.url), 304, b"", r.headers)
                if r.status_code >= 400:
                    raise FetchError(f"HTTP {r.status_code} для {url}")
                return FetchResult(str(r.url), r.status_code, r.content, r.headers)
            except (httpx.TimeoutException, httpx.TransportError) as e:
                last_exc = e
                if attempt < retries:
                    await asyncio.sleep(1.5 * 2**attempt + random.random())
                    continue
        raise FetchError(f"Сеть: {url}: {last_exc}")


def _retry_after(r: httpx.Response) -> float | None:
    v = r.headers.get("retry-after")
    if not v:
        return None
    try:
        return float(v)
    except ValueError:
        return None
