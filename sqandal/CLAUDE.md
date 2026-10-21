# SQANDAL RADAR — заметки для Claude Code

Мониторинг хип-хоп новостей → сюжеты → посты для @SQANDALru. Python 3.11+ (FastAPI, SQLAlchemy 2.0 async,
feedparser, Telethon, Playwright), React 19 + Vite (JSX, без UI-библиотек). Подробности — README.md.

## Команды
- `make test` — все тесты (async через `asyncio.run`, плагины pytest не нужны). PostgreSQL: `SQANDAL_TEST_PG=…`.
- `make demo` — пересобрать вымышленные демо-данные и `frontend/src/demo/demo-data.json`.
- `python -m sqandal.cli snapshot snapshot/2026-10-21 --export ../frontend/src/demo/snapshot-data.json` — снимок
  реальных новостей для демо-дашборда (режим по умолчанию).
- `make schema` — после любой правки `backend/sqandal/models.py` перегенерировать `backend/db/schema.*.sql`.
- `python -m sqandal.cli collect --source KEY` — прогнать один источник (без LLM).

## Правила
- Время внутри системы — только aware UTC (`timeutil.utcnow`, тип `UTCDateTime`); naive datetime в БД запрещён.
- Формат поста живёт в двух местах: `backend/sqandal/publish/formatter.py` и `frontend/src/lib/format.js` —
  меняешь один, меняй второй.
- Режим «Реальные новости» в артефакте: база артефакта (`db`) + кнопка «Обновить» → задача в облаке
  (`artifact/collector-prompt.md`, ID в `frontend/src/demo/liveApi.js`). Выбор источников, промпт, нормализация
  ответа, приоритет и фактчек чисел продублированы в `frontend/src/demo/newsroom.js` (системный промпт, схема
  и веса приходят из снимка). Меняешь `llm/`, `tags.yaml` или `Newsroom._warnings` — правь и там, и в задании.
- Дизайн — система CITITOOL «Präzision» (`frontend/src/styles.css`): бумага `#ECEDEA`, белые карточки, чернильный
  акцент без синего, JetBrains Mono для меток и чисел, шкала-линейка `--tick-rule`; тёмный вариант — по теме зрителя.
  В интерфейсе только SVG-иконки (`Icons.jsx`), эмодзи и флаги — только в тексте постов. Space Grotesk — лишь знак
  SQANDAL (кириллицы в нём нет), русский текст — Manrope.
- `fire_trigger` вызывать без `text`: этот текст приходит в задачу отдельным сообщением уже после сбора, задача
  прогоняет сбор второй раз, и пустой второй прогон затирает итог первого в `meta/status`.
- Claude API: `claude-sonnet-5-5` отклоняет принудительный `tool_choice` и `temperature`/`top_p` (400).
  Структура ответа — через `output_config.format` (JSON Schema). Тест `test_request_shape_matches_api_constraints`.
- Независимые источники считаются по `Source.name` (издание), не по `source_id`.
- Новый сайт без RSS: `kind: web` + `link_pattern` в `config/sources.yaml`; JS-сайт — `render: browser`.
- Вымышленное демо (`make demo`) — только вымышленные артисты и издания. Снимок реальных новостей — только
  факты из публикаций, пересказанные своими словами, со ссылками на оригиналы.
- В канал ничего не уходит без действия редактора (кнопка, кнопка админ-чата, запланированное время).
