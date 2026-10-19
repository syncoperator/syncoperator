# Локальная разработка без Docker. Python 3.11+, Node 20+.
VENV ?= .venv
PY := $(VENV)/bin/python
BACK := backend

.PHONY: setup api worker web demo demo-api demo-web test schema

setup:
	python3 -m venv $(VENV)
	$(PY) -m pip install -r $(BACK)/requirements-dev.txt
	$(PY) -m pip install --extra-index-url https://download.pytorch.org/whl/cpu -r $(BACK)/requirements-ml.txt
	$(VENV)/bin/playwright install chromium
	cd frontend && npm install
	@test -f .env || cp .env.example .env

api:        ## API + SSE на :8000
	cd $(BACK) && ../$(PY) -m uvicorn sqandal.api.app:app --reload --port 8000

worker:     ## сбор, дедуп, генерация, отложенный постинг
	cd $(BACK) && ../$(PY) -m sqandal.worker

web:        ## дашборд на :5173 (прокси /api → :8000)
	cd frontend && npm run dev

demo:       ## вымышленные сюжеты в отдельной БД data/demo.db + экспорт для демо-дашборда
	cd $(BACK) && rm -f data/demo.db* && ../$(PY) -m sqandal.cli demo --export ../frontend/src/demo/demo-data.json

demo-api:   ## API поверх демо-БД (вместе с `make web`); ключи обнулены — вымысел не уйдёт в канал
	cd $(BACK) && DATABASE_URL=sqlite+aiosqlite:///./data/demo.db TG_BOT_TOKEN= ANTHROPIC_API_KEY= \
		../$(PY) -m uvicorn sqandal.api.app:app --port 8000

demo-web:   ## дашборд в демо-режиме, без бэкенда
	cd frontend && npm run dev:demo

test:
	cd $(BACK) && ../$(PY) -m pytest

schema:     ## DDL для PostgreSQL и SQLite из моделей
	cd $(BACK) && ../$(PY) -m sqandal.cli dump-schema postgresql > db/schema.postgres.sql
	cd $(BACK) && ../$(PY) -m sqandal.cli dump-schema sqlite > db/schema.sqlite.sql
