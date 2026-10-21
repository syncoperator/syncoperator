"""Демо-режим: вымышленные артисты и издания, весь пайплайн без сети и ключей.

`python -m sqandal.cli demo` — наполняет БД сюжетами через настоящий пайплайн
(кластеризация hashing-эмбеддингами, теги, приоритеты, свежесть), генерирует
посты через Newsroom с подменённым Claude-транспортом (заготовленные ответы)
и «публикует» один пост через подменённый Bot API.

Все имена артистов и изданий здесь ВЫДУМАНЫ.
"""

from __future__ import annotations

import json
from datetime import timedelta
from typing import Any

import httpx

from .collectors.base import RawItem, SourceConfig
from .models import DateSource

DEMO_SOURCES: list[dict[str, Any]] = [
    {"key": "demo_wire", "name": "Demo Wire", "kind": "web", "scene": "us", "url": "https://demo.local/wire",
     "link_pattern": r"^https://demo\.local/", "enabled": False, "note": "DEMO: вымышленное издание"},
    {"key": "demo_beat", "name": "Demo Beat", "kind": "rss", "scene": "us", "url": "https://demo.local/beat.xml",
     "enabled": False, "note": "DEMO: вымышленное издание"},
    {"key": "demo_insider_us", "name": "Демо-инсайд US", "kind": "telegram", "scene": "us",
     "channel": "demo_insider_us", "enabled": False, "note": "DEMO: вымышленный канал"},
    {"key": "demo_rundschau", "name": "Demo Rundschau", "kind": "rss", "scene": "de",
     "url": "https://demo.local/rundschau.xml", "enabled": False, "note": "DEMO: вымышленное издание"},
    {"key": "demo_vestnik", "name": "Демо-Вестник", "kind": "rss", "scene": "ru",
     "url": "https://demo.local/vestnik.xml", "enabled": False, "note": "DEMO: вымышленное издание"},
    {"key": "demo_insider_ru", "name": "Демо-инсайд RU", "kind": "telegram", "scene": "ru",
     "channel": "demo_insider_ru", "enabled": False, "note": "DEMO: вымышленный канал"},
]

# minutes — сколько минут назад вышла публикация. Ссылки demo.local — служебные.
DEMO_STORIES: list[dict[str, Any]] = [
    {
        "key": "static-arrest",
        "items": [
            {"source": "demo_insider_us", "minutes": 110,
             "text": "Lil Static задержан в Атланте: по нашим данным, была погоня, в машине нашли оружие. "
                     "Подробности — у Demo Wire: https://demo.local/us/static-arrest"},
            {"source": "demo_wire", "minutes": 95, "url": "https://demo.local/us/static-arrest",
             "title": "Lil Static arrested in Atlanta after police chase on I-85",
             "text": "Atlanta rapper Lil Static was arrested on Saturday night, October 18, after a police chase on "
                     "I-85, according to police records. He was booked into Fulton County Jail on a felony weapons "
                     "charge and a misdemeanor marijuana possession charge. Bond was set at $25,000; he was released "
                     "about six hours later. His Static Season tour is scheduled to start November 1 in Charlotte."},
            {"source": "demo_beat", "minutes": 70, "url": "https://demo.local/beat/static-bond",
             "title": "Lil Static arrested after I-85 chase, released on $25K bond",
             "text": "Rapper Lil Static was arrested in Atlanta after a police chase on I-85 on October 18. "
                     "Records show a felony weapons charge and a misdemeanor marijuana charge; bond was $25,000. "
                     "His team has not commented. The Static Season tour starts November 1."},
        ],
        "post": {
            "scene": "us", "flag_emoji": "🇺🇸", "topic_emoji": "🚨",
            "headline": "Lil Static задержан в Атланте после погони",
            "bullets": [
                "18 октября рэпера задержали после погони полиции по трассе I-85 — по данным Demo Wire.",
                "Обвинения: незаконное хранение оружия (felony — тяжкое преступление) и марихуаны (misdemeanor).",
                "Залог — $25 000, вышел примерно через шесть часов.",
                "Команда артиста ситуацию не комментирует.",
            ],
            "summary": "Тур Static Season стартует 1 ноября — теперь под вопросом каждая дата.",
            "tags": ["Задержание", "Суд"], "event_date": "2026-10-18",
            "notes_for_editor": "Подтверждено двумя СМИ; инсайд вышел раньше всех. Комментария артиста нет.",
        },
        "harder": {
            "headline": "Lil Static: погоня, ствол, залог $25 000",
            "bullets": [
                "18 октября — погоня по I-85 и задержание, данные Demo Wire.",
                "Обвинения: оружие (felony) и марихуана (misdemeanor).",
                "Залог $25 000 — на свободе через шесть часов.",
                "Команда молчит.",
            ],
            "summary": "До тура две недели. Дело — уже сейчас.",
        },
        "status": "draft",
    },
    {
        "key": "jayro-tracklist",
        "items": [
            {"source": "demo_beat", "minutes": 40, "url": "https://demo.local/beat/glass-mansion",
             "title": "Jayro Cartel reveals Glass Mansion tracklist with Vex Morgan and Big Halo",
             "text": "Jayro Cartel has revealed the tracklist for his new album Glass Mansion. The album has 16 songs "
                     "and features Vex Morgan and Big Halo. Glass Mansion drops October 31 via Nightshift Records."},
        ],
        "post": {
            "scene": "us", "flag_emoji": "🇺🇸", "topic_emoji": "💿",
            "headline": "Jayro Cartel показал треклист «Glass Mansion»",
            "bullets": [
                "Альбом выходит 31 октября на Nightshift Records.",
                "В треклисте 17 треков.",
                "Гости — Vex Morgan и Big Halo.",
            ],
            "summary": "Хэллоуин в этом году — с релизом Cartel.",
            "tags": ["Релиз"], "event_date": "2026-10-31",
            "notes_for_editor": "Один источник.",
        },
        "status": "draft",
    },
    {
        "key": "rook-vex-beef",
        "items": [
            {"source": "demo_wire", "minutes": 180, "url": "https://demo.local/us/rook-diss",
             "title": "MC Rook drops diss track aimed at Vex Morgan",
             "text": "MC Rook released a diss track titled Paper Crown aimed at Vex Morgan on October 19. The track "
                     "references their 2025 label dispute. Vex Morgan responded on Instagram within two hours."},
            {"source": "demo_beat", "minutes": 150, "url": "https://demo.local/beat/vex-responds",
             "title": "Vex Morgan fires back at MC Rook after Paper Crown diss",
             "text": "Vex Morgan fired back at MC Rook after the diss track Paper Crown, calling it a publicity stunt "
                     "in an Instagram story. Rook's diss track came out on October 19."},
        ],
        "post": {
            "scene": "us", "flag_emoji": "🇺🇸", "topic_emoji": "🔥",
            "headline": "MC Rook выпустил дисс на Vex Morgan",
            "bullets": [
                "19 октября вышел трек «Paper Crown» — прямой выпад в сторону Vex Morgan.",
                "В тексте — отсылки к их спору с лейблом 2025 года.",
                "Vex Morgan ответил в сторис Instagram через два часа и назвал дисс пиаром.",
            ],
            "summary": "Ответный трек — вопрос дней.",
            "tags": ["Скандал"], "event_date": "2026-10-19", "notes_for_editor": "Два источника.",
        },
        "status": "scheduled",
    },
    {
        "key": "halo-hospital",
        "items": [
            {"source": "demo_wire", "minutes": 20, "url": "https://demo.local/us/halo-crash",
             "title": "Big Halo hospitalized after car crash in Houston",
             "text": "Big Halo was hospitalized after a car crash in Houston on Sunday, his manager confirmed. "
                     "According to the manager, his condition is stable. A show in Dallas on October 22 is postponed."},
        ],
        "status": "none",
    },
    {
        "key": "kaltblut-court",
        "items": [
            {"source": "demo_rundschau", "minutes": 30 * 60, "url": "https://demo.local/de/kalt-blut-prozess",
             "title": "Kalt Blut vor Gericht: Prozess wegen Körperverletzung beginnt",
             "text": "Der Berliner Rapper Kalt Blut muss sich vor dem Amtsgericht Tiergarten verantworten. Die Anklage "
                     "wirft ihm Körperverletzung nach § 223 StGB vor. Der Prozess beginnt am 17. Oktober. "
                     "Sein Anwalt kündigte an, die Vorwürfe zurückzuweisen."},
        ],
        "post": {
            "scene": "de", "flag_emoji": "🇩🇪", "topic_emoji": "⚖️",
            "headline": "Kalt Blut идёт под суд в Берлине",
            "bullets": [
                "Процесс в участковом суде Тиргартена начинается 17 октября.",
                "Обвинение — причинение телесных повреждений по § 223 StGB.",
                "Адвокат рэпера заявил, что обвинения отвергнет.",
            ],
            "summary": "Берлинский рэп снова решает вопросы в зале суда.",
            "tags": ["Суд"], "event_date": "2026-10-17", "notes_for_editor": "Новость вышла больше суток назад.",
        },
        "status": "draft",
    },
    {
        "key": "nebel-ep",
        "items": [
            {"source": "demo_rundschau", "minutes": 190, "url": "https://demo.local/de/nebelkraehe-ep",
             "title": "Nebelkrähe veröffentlicht EP Asphaltblumen",
             "text": "Die Rapperin Nebelkrähe hat ihre EP Asphaltblumen veröffentlicht. Die EP hat sechs Tracks, "
                     "Produzent ist Kranich. Im November folgt eine Club-Tour durch vier Städte."},
        ],
        "post": {
            "scene": "de", "flag_emoji": "🇩🇪", "topic_emoji": "💿",
            "headline": "Nebelkrähe выпустила EP «Asphaltblumen»",
            "bullets": [
                "В EP — шесть треков, продюсер — Kranich.",
                "В ноябре — клубный тур по четырём городам.",
                "Релиз уже на стримингах.",
            ],
            "summary": "Тихий релиз с громкой заявкой на зиму.",
            "tags": ["Релиз", "Тур"], "event_date": None, "notes_for_editor": "Один источник.",
        },
        "status": "draft",
    },
    {
        "key": "sturm-tour",
        "items": [
            {"source": "demo_rundschau", "minutes": 260, "url": "https://demo.local/de/sami-sturm-tour",
             "title": "Sami Sturm sagt Tour ab",
             "text": "Sami Sturm hat seine Herbsttour abgesagt. Als Grund nannte er persönliche Gründe. "
                     "Tickets werden erstattet."},
        ],
        "status": "dismissed",
    },
    {
        "key": "pepel-arrest",
        "items": [
            {"source": "demo_insider_ru", "minutes": 50,
             "text": "Пепла задержали в Казани после концерта в клубе «Гараж». По нашим данным — ч. 1 ст. 228 УК РФ. "
                     "Ждём официальных комментариев."},
            {"source": "demo_vestnik", "minutes": 25, "url": "https://demo.local/ru/pepel-kazan",
             "title": "Рэпера Пепла задержали в Казани после концерта",
             "text": "Рэпера Пепла задержали в Казани 19 октября после концерта в клубе «Гараж». По данным источника, "
                     "против артиста возбуждено дело по ч. 1 ст. 228 УК РФ. Официальных комментариев от полиции и "
                     "самого артиста нет. Концерт в Самаре 21 октября организаторы пока не отменяли."},
        ],
        "post": {
            "scene": "ru", "flag_emoji": "🇷🇺", "topic_emoji": "🚨",
            "headline": "Пепла задержали в Казани после концерта",
            "bullets": [
                "Задержание — 19 октября, после выступления в клубе «Гараж».",
                "Дело — по ч. 1 ст. 228 УК РФ, по данным Демо-Вестника.",
                "Официальных комментариев от полиции и артиста нет.",
                "Концерт в Самаре 21 октября пока не отменён.",
            ],
            "summary": "Судьба тура теперь зависит не от продаж, а от следствия.",
            "tags": ["Задержание", "Суд", "Тур"], "event_date": "2026-10-19",
            "notes_for_editor": "Инсайд подтвердило СМИ. Статью называет только один источник.",
        },
        "harder": {
            "headline": "Пепла приняли в Казани",
            "bullets": [
                "19 октября, сразу после концерта в «Гараже».",
                "Дело по ч. 1 ст. 228 УК РФ — данные Демо-Вестника.",
                "Полиция и артист молчат.",
                "Самара 21 октября — пока в силе.",
            ],
            "summary": "Тур теперь решает следствие.",
        },
        "status": "draft",
    },
    {
        "key": "led-album",
        "items": [
            {"source": "demo_vestnik", "minutes": 300, "url": "https://demo.local/ru/led-sibir",
             "title": "Лёд Восток выпустил альбом «Сибирь»",
             "text": "Лёд Восток выпустил альбом «Сибирь». В релизе 12 треков, среди гостей — Мира Кай. "
                     "Презентация пройдёт 8 ноября в Новосибирске."},
        ],
        "post": {
            "scene": "ru", "flag_emoji": "🇷🇺", "topic_emoji": "💿",
            "headline": "Лёд Восток выпустил альбом «Сибирь»",
            "bullets": ["В альбоме 12 треков.", "Гостья релиза — Мира Кай.", "Презентация — 8 ноября в Новосибирске."],
            "summary": "Холодное название, горячий релиз сезона.",
            "tags": ["Релиз"], "event_date": None, "notes_for_editor": "",
        },
        "status": "published",
    },
    {
        "key": "mira-label",
        "items": [
            {"source": "demo_insider_ru", "minutes": 75,
             "text": "Мира Кай судится с лейблом «Северный ветер»: иск подан в Арбитражный суд Москвы. "
                     "По нашей информации, спор — о правах на два альбома. Лейбл пока молчит."},
        ],
        "post": {
            "scene": "ru", "flag_emoji": "🇷🇺", "topic_emoji": "⚖️",
            "headline": "Мира Кай подала иск к лейблу «Северный ветер»",
            "bullets": [
                "Иск — в Арбитражном суде Москвы, по данным инсайдеров.",
                "Спор, по их же информации, — о правах на два альбома.",
                "Лейбл не комментирует; официального подтверждения нет.",
            ],
            "summary": "Если подтвердится — главный трудовой спор сцены этой осени.",
            "tags": ["Суд", "Скандал"], "event_date": None,
            "notes_for_editor": "Только Telegram-инсайд — дождись подтверждения СМИ или карточки дела.",
        },
        "status": "draft",
    },
]


def demo_source_configs() -> list[SourceConfig]:
    return [SourceConfig.model_validate(x) for x in DEMO_SOURCES]


def tg_external_id(story_idx: int, n: int) -> tuple[str, int]:
    story = DEMO_STORIES[story_idx]
    src = next(s for s in DEMO_SOURCES if s["key"] == story["items"][n]["source"])
    msg_id = 1000 + story_idx * 10 + n
    return f"{src['channel']}:{msg_id}", msg_id


def story_markers(story_idx: int) -> list[str]:
    """URL/external_id публикаций сюжета — по ним демо находит свой кластер."""
    out = []
    for n, it in enumerate(DEMO_STORIES[story_idx]["items"]):
        out.append(it["url"] if it.get("url") else tg_external_id(story_idx, n)[0])
    return out


def demo_raw_items(now) -> list[tuple[str, RawItem]]:
    """(source_key, RawItem) в хронологическом порядке — как приходили бы вживую."""
    out = []
    for si, story in enumerate(DEMO_STORIES):
        for n, it in enumerate(story["items"]):
            src = next(s for s in DEMO_SOURCES if s["key"] == it["source"])
            published = now - timedelta(minutes=it["minutes"])
            fetched = published + timedelta(minutes=3)
            if src["kind"] == "telegram":
                _, msg_id = tg_external_id(si, n)
                links = [w for w in it["text"].split() if w.startswith("https://")]
                raw = RawItem(
                    external_id=f"{src['channel']}:{msg_id}", title=it["text"].split(".")[0][:150],
                    url=f"https://t.me/{src['channel']}/{msg_id}", summary=it["text"], published_at=published,
                    date_source=DateSource.telegram, tg_channel=src["channel"], tg_message_id=msg_id,
                    tg_views=12000 + msg_id, extra={"links": links} if links else {}, fetched_at=fetched,
                )
            else:
                raw = RawItem(
                    external_id=it["url"], title=it["title"], url=it["url"], summary=it["text"],
                    published_at=published,
                    date_source=DateSource.jsonld if src["kind"] == "web" else DateSource.feed_published,
                    fetched_at=fetched,
                )
            out.append((src["key"], raw, published))
    out.sort(key=lambda x: x[2])
    return [(k, r) for k, r, _ in out]


def _story_for_prompt(prompt: str) -> dict | None:
    for story in DEMO_STORIES:
        for it in story["items"]:
            marker = it.get("url") or it["text"][:40]
            if marker in prompt:
                return story
    return None


class DemoClaudeTransport(httpx.AsyncBaseTransport):
    """Подмена Claude API: отдаёт заготовленный пост для сюжета из DEMO_STORIES."""

    async def handle_async_request(self, request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        prompt = body["messages"][0]["content"]
        story = _story_for_prompt(prompt)
        if story is None or "post" not in story:
            post = {
                "is_publishable": True, "scene": "us", "flag_emoji": "🇺🇸", "topic_emoji": "🏥",
                "headline": "Big Halo госпитализирован после ДТП в Хьюстоне",
                "bullets": [
                    "Авария — в воскресенье в Хьюстоне, подтвердил менеджер артиста.",
                    "По словам менеджера, состояние стабильное.",
                    "Шоу в Далласе 22 октября перенесено.",
                ],
                "summary": "Сначала здоровье — потом сцена.", "tags": ["Здоровье", "Тур"],
                "event_date": None, "notes_for_editor": "Один источник; детали травм не раскрыты.",
            }
        else:
            post = {"is_publishable": True, **story["post"]}
            if ("ЖЁСТЧЕ" in prompt or "ЖЁСТКО" in prompt) and story.get("harder"):
                post = {**post, **story["harder"]}
        data = {
            "id": "msg_demo", "type": "message", "role": "assistant", "model": body["model"],
            "content": [{"type": "text", "text": json.dumps(post, ensure_ascii=False)}],
            "stop_reason": "end_turn",
            "usage": {"input_tokens": 1800, "output_tokens": 420, "cache_read_input_tokens": 1500},
        }
        return httpx.Response(200, json=data)


class DemoBotTransport(httpx.AsyncBaseTransport):
    """Подмена Bot API: «публикует» без сети."""

    def __init__(self):
        self.sent: list[dict] = []
        self._msg_id = 4200

    async def handle_async_request(self, request: httpx.Request) -> httpx.Response:
        method = request.url.path.rsplit("/", 1)[-1]
        payload = json.loads(request.content or b"{}")
        self.sent.append({"method": method, **payload})
        if method == "sendMessage":
            self._msg_id += 1
            return httpx.Response(200, json={"ok": True, "result": {
                "message_id": self._msg_id, "chat": {"id": -1001234567890, "username": "SQANDALru", "type": "channel"}}})
        return httpx.Response(200, json={"ok": True, "result": True})
