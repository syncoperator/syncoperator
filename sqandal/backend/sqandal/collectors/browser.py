"""Playwright-рендер для сайтов, которые собирают страницу JavaScript'ом (Tilda, SPA)."""

from __future__ import annotations

import asyncio
import logging

from ..settings import Settings

log = logging.getLogger(__name__)

_BLOCKED = {"image", "media", "font"}


class BrowserPool:
    """Один Chromium на процесс, страницы — в отдельных контекстах. Запуск лениво."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self._pw = None
        self._browser = None
        self._lock = asyncio.Lock()
        self._sem = asyncio.Semaphore(max(1, settings.browser_concurrency))

    async def _ensure(self):
        async with self._lock:
            if self._browser is None:
                from playwright.async_api import async_playwright  # тяжёлый импорт — только по требованию

                self._pw = await async_playwright().start()
                self._browser = await self._pw.chromium.launch(
                    headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
                log.info("Playwright Chromium запущен")
        return self._browser

    async def render(self, url: str, wait_selector: str | None = None, timeout_ms: int = 30000) -> str:
        browser = await self._ensure()
        async with self._sem:
            context = await browser.new_context(
                user_agent=self.settings.http_user_agent,
                locale="ru-RU",
                java_script_enabled=True,
                viewport={"width": 1280, "height": 2000},
            )
            try:
                page = await context.new_page()

                async def _route(route):
                    if route.request.resource_type in _BLOCKED:
                        await route.abort()
                    else:
                        await route.continue_()

                await page.route("**/*", _route)
                await page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
                if wait_selector:
                    await page.wait_for_selector(wait_selector, timeout=timeout_ms)
                else:
                    try:
                        await page.wait_for_load_state("networkidle", timeout=min(timeout_ms, 15000))
                    except Exception:  # noqa: BLE001 — долгие трекеры не повод падать
                        pass
                return await page.content()
            finally:
                await context.close()

    async def close(self) -> None:
        if self._browser is not None:
            await self._browser.close()
            self._browser = None
        if self._pw is not None:
            await self._pw.stop()
            self._pw = None
