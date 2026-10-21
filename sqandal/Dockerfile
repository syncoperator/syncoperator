# syntax=docker/dockerfile:1
# Один образ на API и воркер. Сборка дашборда → статика внутри API.

FROM node:22-alpine AS web
WORKDIR /web
COPY frontend/package.json ./
RUN npm install --no-audit --no-fund
COPY frontend/ ./
RUN npm run build

FROM python:3.12-slim AS app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 \
    HF_HOME=/models PLAYWRIGHT_BROWSERS_PATH=/ms-playwright
WORKDIR /app

COPY backend/requirements.txt backend/requirements-ml.txt ./
RUN pip install -r requirements.txt

# Локальная модель эмбеддингов (CPU). WITH_ML=0 — легче образ, дедуп через Voyage или hashing.
ARG WITH_ML=1
RUN if [ "$WITH_ML" = "1" ]; then \
      pip install --extra-index-url https://download.pytorch.org/whl/cpu -r requirements-ml.txt; \
    fi

# Chromium для сайтов на JavaScript (render: browser). WITH_BROWSER=0 — без него.
ARG WITH_BROWSER=1
RUN if [ "$WITH_BROWSER" = "1" ]; then playwright install --with-deps chromium; fi

COPY backend/ ./
COPY --from=web /web/dist /app/frontend_dist
ENV FRONTEND_DIST=/app/frontend_dist CONFIG_DIR=/app/config

EXPOSE 8000
CMD ["uvicorn", "sqandal.api.app:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers"]
