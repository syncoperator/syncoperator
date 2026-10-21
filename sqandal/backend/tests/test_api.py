"""API дашборда и публикация: авторизация, лента, генерация «жёстче», правка, защита от
старых фактов, публикация в Bot API, планирование, вебхук кнопок админ-чата, SSE."""

from __future__ import annotations

import asyncio
import json
import socket
import threading
import time
from datetime import timedelta

import httpx

from conftest import run
from sqandal.api.app import create_app
from sqandal.bootstrap import build_core
from sqandal.cli import seed_demo
from sqandal.demo import DemoBotTransport, DemoClaudeTransport
from sqandal.timeutil import utcnow

H = {"Authorization": "Bearer test-token"}


def _app(settings, bot_tr):
    async def factory(st):
        core = await build_core(st, claude_transport=DemoClaudeTransport(), bot_transport=bot_tr)
        core.newsroom.fetcher = None
        return core

    return create_app(settings, core_factory=factory)


async def _client(app):
    return httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="http://t")


def _find(stories, needle):
    return next(s for s in stories if needle in s["title"])


def test_dashboard_flow(settings_factory):
    st = settings_factory(tg_admin_chat_id="-100777", tg_admin_mode="webhook", tg_webhook_secret="s3cr3t")
    run(seed_demo(st))
    bot = DemoBotTransport()
    app = _app(st, bot)

    async def go():
        async with app.router.lifespan_context(app):
            c = await _client(app)
            assert (await c.get("/api/stories")).status_code == 401
            assert (await c.get("/api/stories", headers={"Authorization": "Bearer nope"})).status_code == 401
            assert (await c.get("/api/health")).json()["ok"] is True

            stories = (await c.get("/api/stories", headers=H)).json()["stories"]
            assert len(stories) == 8  # активные: без опубликованного и скрытого
            ru = (await c.get("/api/stories?scene=ru", headers=H)).json()["stories"]
            assert {s["scene"] for s in ru} == {"ru"} and len(ru) == 2
            crit = (await c.get("/api/stories?priority=CRITICAL", headers=H)).json()["stories"]
            assert len(crit) == 2
            by_prio = (await c.get("/api/stories?sort=priority", headers=H)).json()["stories"]
            assert by_prio[0]["priority"] == "CRITICAL"
            assert len((await c.get("/api/stories?tag=Суд", headers=H)).json()["stories"]) >= 3
            assert len((await c.get("/api/stories?q=Пепл", headers=H)).json()["stories"]) == 1
            everything = (await c.get("/api/stories?status=all", headers=H)).json()["stories"]
            assert len(everything) == 10
            # красные точки на вкладках считаются по всем сценам, не по текущему фильтру
            assert (await c.get("/api/stats", headers=H)).json()["critical_by_scene"] == {"ru": 1, "us": 1}

            static = _find(stories, "Lil Static")
            det = (await c.get(f"/api/stories/{static['id']}", headers=H)).json()
            assert len(det["items"]) == 3 and det["origin"]["source_name"] == "Демо-инсайд US"
            assert det["source_names"][0] == "Демо-инсайд US"
            post = det["current_post"]
            assert post["status"] == "draft" and post["version"] == 1
            assert post["body_html"].startswith("🇺🇸🚨 <b>Lil Static задержан")

            # «Перегенерировать жёстче» → новая версия, старый черновик вытеснен
            r = await c.post(f"/api/stories/{static['id']}/generate", json={"mode": "harder"}, headers=H)
            assert r.status_code == 200, r.text
            hard = r.json()
            assert hard["version"] == 2 and hard["mode"] == "harder"
            assert hard["headline"] == "Lil Static: погоня, ствол, залог $25 000"
            det = (await c.get(f"/api/stories/{static['id']}", headers=H)).json()
            assert [v["status"] for v in det["versions"]] == ["draft", "superseded"]

            # ручная правка с выдуманной цифрой → предупреждение фактчека
            r = await c.patch(f"/api/posts/{hard['id']}", json={"headline": "Lil Static отпустили за 3 часа"},
                              headers=H)
            patched = r.json()
            assert patched["mode"] == "manual" and "отпустили за 3 часа" in patched["body_html"]
            assert any("Числа не найдены" in w and "3" in w for w in patched["warnings"])

            # публикация в канал
            r = await c.post(f"/api/posts/{hard['id']}/publish", json={}, headers=H)
            assert r.status_code == 200, r.text
            pub = r.json()
            assert pub["status"] == "published" and pub["link"].startswith("https://t.me/SQANDALru/")
            sent = [x for x in bot.sent if x["method"] == "sendMessage"][-1]
            assert sent["chat_id"] == "@SQANDALru" and sent["parse_mode"] == "HTML"
            assert "отпустили за 3 часа" in sent["text"]
            assert (await c.post(f"/api/posts/{hard['id']}/publish", json={}, headers=H)).status_code == 409

            # старый факт: без подтверждения — 409 с кодом, с подтверждением — ок
            kalt = _find(stories, "Kalt Blut")
            kp = (await c.get(f"/api/stories/{kalt['id']}", headers=H)).json()["current_post"]
            r = await c.post(f"/api/posts/{kp['id']}/publish", json={}, headers=H)
            assert r.status_code == 409 and r.json()["code"] == "stale_confirmation_required"
            r = await c.post(f"/api/posts/{kp['id']}/publish", json={"confirm_stale": True}, headers=H)
            assert r.json()["status"] == "published"

            # планирование
            nebel = _find(stories, "Nebelkrähe")
            np_ = (await c.get(f"/api/stories/{nebel['id']}", headers=H)).json()["current_post"]
            past = (utcnow() - timedelta(minutes=5)).isoformat()
            assert (await c.post(f"/api/posts/{np_['id']}/schedule", json={"scheduled_at": past},
                                 headers=H)).status_code == 400
            when = (utcnow() + timedelta(hours=2)).isoformat()
            r = await c.post(f"/api/posts/{np_['id']}/schedule", json={"scheduled_at": when}, headers=H)
            assert r.json()["status"] == "scheduled"
            r = await c.delete(f"/api/posts/{np_['id']}/schedule", headers=H)
            assert r.json()["status"] == "draft"

            # сюжет без поста: генерация по клику
            halo = _find(stories, "Big Halo")
            assert halo["post"] is None
            r = await c.post(f"/api/stories/{halo['id']}/generate", json={}, headers=H)
            assert r.json()["headline"].startswith("Big Halo")

            # скрыть / вернуть
            r = await c.post(f"/api/stories/{halo['id']}/dismiss", headers=H)
            assert r.json()["status"] == "dismissed"
            r = await c.post(f"/api/stories/{halo['id']}/restore", headers=H)
            assert r.json()["status"] == "drafted"

            stats = (await c.get("/api/stats", headers=H)).json()
            assert stats["published_24h"] >= 3 and set(stats["by_scene"]) <= {"ru", "de", "us"}
            cfg = (await c.get("/api/config", headers=H)).json()
            assert cfg["brand_line"] == "🔊 SQANDAL" and cfg["llm_ready"] and cfg["bot_ready"]
            assert [t["name"] for t in cfg["tags"]][:2] == ["Смерть", "Задержание"]
            srcs = (await c.get("/api/sources", headers=H)).json()
            assert any(s["key"] == "rapde" and not s["enabled"] for s in srcs)
            assert (await c.post("/api/collect", headers=H)).json() == {"queued": True}

            # вебхук: чужой секрет — 403; кнопка «В канал» в админ-чате публикует пост
            pepel = _find(stories, "Пепла")
            pp = (await c.get(f"/api/stories/{pepel['id']}", headers=H)).json()["current_post"]
            upd = {"update_id": 1, "callback_query": {
                "id": "cb1", "data": f"pub:{pp['id']}", "from": {"id": 1},
                "message": {"message_id": 55, "chat": {"id": -100777}}}}
            assert (await c.post("/api/telegram/webhook", json=upd,
                                 headers={"X-Telegram-Bot-Api-Secret-Token": "wrong"})).status_code == 403
            r = await c.post("/api/telegram/webhook", json=upd,
                             headers={"X-Telegram-Bot-Api-Secret-Token": "s3cr3t"})
            assert r.status_code == 200
            for _ in range(50):
                if any(x["method"] == "editMessageText" for x in bot.sent):
                    break
                await asyncio.sleep(0.05)
            methods = [x["method"] for x in bot.sent]
            assert "answerCallbackQuery" in methods and "editMessageText" in methods
            assert (await c.get(f"/api/posts/{pp['id']}", headers=H)).json()["status"] == "published"
            # из чужого чата — отказ
            bad = {**upd, "callback_query": {**upd["callback_query"], "id": "cb2",
                                             "message": {"message_id": 1, "chat": {"id": 999}}}}
            await c.post("/api/telegram/webhook", json=bad, headers={"X-Telegram-Bot-Api-Secret-Token": "s3cr3t"})
            await asyncio.sleep(0.2)
            denied = [x for x in bot.sent if x.get("callback_query_id") == "cb2"]
            assert denied and denied[0].get("show_alert") is True
            await c.aclose()

    run(go())


def test_publish_failure_marks_post_failed(settings_factory):
    st = settings_factory()
    run(seed_demo(st))

    class FailingBot(httpx.AsyncBaseTransport):
        async def handle_async_request(self, request):
            return httpx.Response(400, json={"ok": False, "error_code": 400,
                                             "description": "Bad Request: can't parse entities"})

    app = _app(st, FailingBot())

    async def go():
        async with app.router.lifespan_context(app):
            c = await _client(app)
            stories = (await c.get("/api/stories", headers=H)).json()["stories"]
            pid = (await c.get(f"/api/stories/{_find(stories, 'Пепла')['id']}", headers=H)).json()["current_post"]["id"]
            r = await c.post(f"/api/posts/{pid}/publish", json={}, headers=H)
            assert r.status_code == 400 and "can't parse entities" in r.json()["detail"]
            p = (await c.get(f"/api/posts/{pid}", headers=H)).json()
            assert p["status"] == "failed" and "parse entities" in p["error"]
            await c.aclose()

    run(go())


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_sse_stream_pushes_updates(settings_factory):
    """Реальный uvicorn: SSE отдаёт stats и обновлённые карточки сюжетов."""
    import uvicorn

    st = settings_factory()
    run(seed_demo(st))
    app = _app(st, DemoBotTransport())
    port = _free_port()
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning"))
    th = threading.Thread(target=server.run, daemon=True)
    th.start()
    base = f"http://127.0.0.1:{port}"
    try:
        for _ in range(100):
            try:
                if httpx.get(base + "/api/health", timeout=1).status_code == 200:
                    break
            except httpx.HTTPError:
                time.sleep(0.05)
        stories = httpx.get(base + "/api/stories", headers=H).json()["stories"]
        target = stories[0]["id"]
        events: list[tuple[str, dict]] = []
        with httpx.stream("GET", base + "/api/stream?token=test-token", timeout=10) as resp:
            assert resp.headers["content-type"].startswith("text/event-stream")
            ev = None
            dismissed = False
            for line in resp.iter_lines():
                if line.startswith("event: "):
                    ev = line[7:]
                elif line.startswith("data: ") and ev:
                    events.append((ev, json.loads(line[6:])))
                    if ev == "stats" and not dismissed:
                        httpx.post(base + f"/api/stories/{target}/dismiss", headers=H)
                        dismissed = True
                    if ev == "story":
                        break
        kinds = [e for e, _ in events]
        assert kinds[0] == "stats"
        story = next(d for e, d in events if e == "story")
        assert story["id"] == target and story["status"] == "dismissed"
    finally:
        server.should_exit = True
        th.join(timeout=10)
