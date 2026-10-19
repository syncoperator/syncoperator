# SQANDAL RADAR — заметки для Claude Code

Мониторинг хип-хоп новостей → сюжеты → посты для @SQANDALru. Python 3.11+ (FastAPI, SQLAlchemy 2.0 async,
feedparser, Telethon, Playwright), React 19 + Vite (JSX, без UI-библиотек). Подробности — README.md.

## Команды
- `make test` — все тесты (async через `asyncio.run`, плагины pytest не нужны). PostgreSQL: `SQANDAL_TEST_PG=…`.
- `make demo` — пересобрать вымышленные демо-данные и `frontend/src/demo/demo-data.json`.
- `make schema` — после любой правки `backend/sqandal/models.py` перегенерировать `backend/db/schema.*.sql`.
- `python -m sqandal.cli collect --source KEY` — прогнать один источник (без LLM).

## Правила
- Время внутри системы — только aware UTC (`timeutil.utcnow`, тип `UTCDateTime`); naive datetime в БД запрещён.
- Формат поста живёт в двух местах: `backend/sqandal/publish/formatter.py` и `frontend/src/lib/format.js` —
  меняешь один, меняй второй.
- Claude API: `claude-sonnet-5-5` отклоняет принудительный `tool_choice` и `temperature`/`top_p` (400).
  Структура ответа — через `output_config.format` (JSON Schema). Тест `test_request_shape_matches_api_constraints`.
- Независимые источники считаются по `Source.name` (издание), не по `source_id`.
- Новый сайт без RSS: `kind: web` + `link_pattern` в `config/sources.yaml`; JS-сайт — `render: browser`.
- Демо-данные — только вымышленные артисты и издания.
- В канал ничего не уходит без действия редактора (кнопка, кнопка админ-чата, запланированное время).
