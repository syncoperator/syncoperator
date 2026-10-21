"""Эмбеддинги для смысловой дедупликации.

Бэкенды:
  local   — sentence-transformers, мультиязычная модель (RU/DE/EN в одном пространстве:
            одна и та же новость на трёх языках склеивается в один сюжет);
  voyage  — Voyage AI API (рекомендация Anthropic для эмбеддингов);
  hashing — без ML: хэшированные n-граммы. Ловит перепечатки на одном языке,
            кросс-язычные дубли — нет. Для тестов и слабого железа.
"""

from __future__ import annotations

import asyncio
import logging
import re
import zlib
from typing import Protocol

import httpx
import numpy as np

from ..settings import Settings

log = logging.getLogger(__name__)

# Пороги косинусной близости «это тот же сюжет». Калибровка: `python -m sqandal.cli calibrate`.
DEFAULT_THRESHOLDS = {
    "hashing": 0.42,
    "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2": 0.72,
    "sentence-transformers/paraphrase-multilingual-mpnet-base-v2": 0.72,
    "intfloat/multilingual-e5-small": 0.88,
    "intfloat/multilingual-e5-base": 0.88,
    "BAAI/bge-m3": 0.78,
    "voyage": 0.80,
}


class Embedder(Protocol):
    model_id: str

    async def embed(self, texts: list[str]) -> np.ndarray:  # (n, dim) float32, L2-норма = 1
        ...


def l2_normalize(m: np.ndarray) -> np.ndarray:
    m = np.asarray(m, dtype=np.float32)
    if m.ndim == 1:
        n = np.linalg.norm(m)
        return m / n if n > 0 else m
    n = np.linalg.norm(m, axis=1, keepdims=True)
    n[n == 0] = 1.0
    return m / n


def to_bytes(v: np.ndarray) -> bytes:
    return np.asarray(v, dtype=np.float32).tobytes()


def from_bytes(b: bytes | None) -> np.ndarray | None:
    if not b:
        return None
    return np.frombuffer(b, dtype=np.float32).copy()


_TOKEN = re.compile(r"\w+", re.UNICODE)


class HashingEmbedder:
    """Детерминированный (crc32, не hash()) bag-of-n-grams в 4096 измерениях."""

    def __init__(self, dim: int = 4096):
        self.dim = dim
        self.model_id = "hashing"

    def _vec(self, text: str) -> np.ndarray:
        v = np.zeros(self.dim, dtype=np.float32)
        text = text.lower().replace("ё", "е")
        words = [w for w in _TOKEN.findall(text) if len(w) > 1]
        feats: list[str] = []
        feats += [f"w:{w}" for w in words]
        feats += [f"b:{a}_{b}" for a, b in zip(words, words[1:])]
        for w in words:
            padded = f"<{w}>"
            feats += [f"c:{padded[i:i + 4]}" for i in range(max(1, len(padded) - 3))]
        for f in feats:
            h = zlib.crc32(f.encode("utf-8"))
            v[h % self.dim] += 1.0 if (h >> 31) & 1 == 0 else -1.0
        v = np.sign(v) * np.log1p(np.abs(v))
        return l2_normalize(v)

    async def embed(self, texts: list[str]) -> np.ndarray:
        if not texts:
            return np.zeros((0, self.dim), dtype=np.float32)
        return np.vstack([self._vec(t) for t in texts])


class SentenceTransformerEmbedder:
    def __init__(self, model_name: str):
        self.model_id = model_name
        self._model = None
        self._lock = asyncio.Lock()
        self._prefix = "query: " if "e5" in model_name.lower() else ""

    async def _load(self):
        async with self._lock:
            if self._model is None:
                from sentence_transformers import SentenceTransformer

                log.info("Загружаю модель эмбеддингов %s…", self.model_id)
                self._model = await asyncio.to_thread(SentenceTransformer, self.model_id, device="cpu")
        return self._model

    async def embed(self, texts: list[str]) -> np.ndarray:
        model = await self._load()
        if not texts:
            dim = model.get_sentence_embedding_dimension()
            return np.zeros((0, dim), dtype=np.float32)
        texts = [self._prefix + t for t in texts]
        vecs = await asyncio.to_thread(
            model.encode, texts, batch_size=32, normalize_embeddings=True, convert_to_numpy=True,
            show_progress_bar=False,
        )
        return l2_normalize(vecs)


class VoyageEmbedder:
    def __init__(self, api_key: str, model: str, base_url: str, transport: httpx.AsyncBaseTransport | None = None):
        self.model_id = f"voyage:{model}"
        self.model = model
        self.client = httpx.AsyncClient(
            base_url=base_url.rstrip("/"),
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=30.0,
            transport=transport,
        )

    async def embed(self, texts: list[str]) -> np.ndarray:
        if not texts:
            return np.zeros((0, 1024), dtype=np.float32)
        out: list[list[float]] = []
        for i in range(0, len(texts), 64):
            batch = texts[i:i + 64]
            for attempt in range(3):
                r = await self.client.post(
                    "/embeddings", json={"input": batch, "model": self.model, "input_type": "document"}
                )
                if r.status_code in (429, 500, 502, 503) and attempt < 2:
                    await asyncio.sleep(2 * (attempt + 1))
                    continue
                r.raise_for_status()
                break
            data = sorted(r.json()["data"], key=lambda d: d["index"])
            out.extend(d["embedding"] for d in data)
        return l2_normalize(np.array(out, dtype=np.float32))


def make_embedder(settings: Settings) -> Embedder:
    if settings.embedding_backend == "hashing":
        return HashingEmbedder()
    if settings.embedding_backend == "voyage":
        if not settings.voyage_api_key:
            raise RuntimeError("EMBEDDING_BACKEND=voyage, но VOYAGE_API_KEY не задан")
        return VoyageEmbedder(settings.voyage_api_key, settings.voyage_model, settings.voyage_base_url)
    try:
        import sentence_transformers  # noqa: F401
    except ImportError:
        log.warning("sentence-transformers не установлен → fallback на hashing (без кросс-язычного дедупа)")
        return HashingEmbedder()
    return SentenceTransformerEmbedder(settings.embedding_model)


def threshold_for(embedder: Embedder, settings: Settings) -> float:
    if settings.cluster_threshold is not None:
        return settings.cluster_threshold
    if embedder.model_id.startswith("voyage"):
        return DEFAULT_THRESHOLDS["voyage"]
    return DEFAULT_THRESHOLDS.get(embedder.model_id, 0.75)


def embedding_text(title: str, summary: str | None, content: str | None, limit: int = 600) -> str:
    """Заголовок важнее всего; добавляем начало текста для контекста."""
    body = (summary or "") + " " + (content or "")
    body = " ".join(body.split())
    return (title.strip() + ". " + body[:limit]).strip()
