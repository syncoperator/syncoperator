"""Онлайн-кластеризация новостей в сюжеты по смысловой близости эмбеддингов.

Новость присоединяется к сюжету, если:
  • косинус с центроидом сюжета ≥ threshold;
  • и с ближайшей публикацией сюжета ≥ threshold − 0.05 (защита от «цепочек»);
  • сюжет активен (в окне cluster_window_hours от его последней публикации).
Повышенный порог, если:
  • в новостях РАЗНЫЕ артисты из вотчлиста (+0.10) — «Drake судится» ≠ «Kanye судится»;
  • теги событий не пересекаются (+0.05) — «Drake: альбом» ≠ «Drake: иск».
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta

import numpy as np

from .embeddings import l2_normalize


@dataclass(slots=True)
class ActiveCluster:
    id: int
    vectors: list[np.ndarray]
    centroid: np.ndarray
    first_at: datetime
    last_at: datetime
    tags: set[str] = field(default_factory=set)
    artists: set[str] = field(default_factory=set)

    def recompute(self) -> None:
        self.centroid = l2_normalize(np.mean(np.vstack(self.vectors), axis=0))


@dataclass(slots=True)
class MatchResult:
    cluster_id: int | None
    similarity: float


class ClusterIndex:
    ARTIST_CONFLICT_PENALTY = 0.10
    TAG_CONFLICT_PENALTY = 0.05
    MEMBER_SLACK = 0.05

    def __init__(self, threshold: float, window_hours: float):
        self.threshold = threshold
        self.window = timedelta(hours=window_hours)
        self.clusters: dict[int, ActiveCluster] = {}

    def __len__(self) -> int:
        return len(self.clusters)

    def required(self, c: ActiveCluster, tags: set[str], artists: set[str]) -> float:
        req = self.threshold
        if artists and c.artists and not (artists & c.artists):
            req += self.ARTIST_CONFLICT_PENALTY
        if tags and c.tags and not (tags & c.tags):
            req += self.TAG_CONFLICT_PENALTY
        return req

    def match(self, vec: np.ndarray, at: datetime, tags: set[str], artists: set[str]) -> MatchResult:
        cands = [
            c for c in self.clusters.values()
            if c.first_at - self.window <= at <= c.last_at + self.window
        ]
        if not cands:
            return MatchResult(None, 0.0)
        cents = np.vstack([c.centroid for c in cands])
        sims = cents @ vec
        order = np.argsort(-sims)
        best = float(sims[order[0]])
        for idx in order:
            c = cands[int(idx)]
            sim = float(sims[idx])
            req = self.required(c, tags, artists)
            if sim < req:
                if sim < self.threshold - 0.1:
                    break  # дальше только хуже
                continue
            member_best = max(float(v @ vec) for v in c.vectors)
            if member_best < req - self.MEMBER_SLACK:
                continue
            return MatchResult(c.id, sim)
        return MatchResult(None, best)

    def add(self, cluster_id: int, vec: np.ndarray, at: datetime, tags: set[str], artists: set[str]) -> ActiveCluster:
        c = self.clusters.get(cluster_id)
        if c is None:
            c = ActiveCluster(cluster_id, [vec], vec.copy(), at, at, set(tags), set(artists))
            self.clusters[cluster_id] = c
            return c
        c.vectors.append(vec)
        if len(c.vectors) > 64:  # большой сюжет (десятки репостов) — хватит последних
            c.vectors = c.vectors[-64:]
        c.first_at = min(c.first_at, at)
        c.last_at = max(c.last_at, at)
        c.tags |= tags
        c.artists |= artists
        c.recompute()
        return c

    def prune(self, now: datetime) -> int:
        old = [cid for cid, c in self.clusters.items() if c.last_at < now - self.window]
        for cid in old:
            del self.clusters[cid]
        return len(old)


@dataclass(slots=True)
class PublishedEntry:
    cluster_id: int
    centroid: np.ndarray
    published_at: datetime
    artists: set[str]


class PublishedIndex:
    """Уже опубликованные сюжеты: ловим повтор старой новости под новым соусом."""

    STRICT_BONUS = 0.03

    def __init__(self, threshold: float, lookback_days: int):
        self.threshold = threshold + self.STRICT_BONUS
        self.lookback = timedelta(days=lookback_days)
        self.entries: dict[int, PublishedEntry] = {}

    def add(self, cluster_id: int, centroid: np.ndarray, published_at: datetime, artists: set[str]) -> None:
        self.entries[cluster_id] = PublishedEntry(cluster_id, centroid, published_at, artists)

    def match(self, vec: np.ndarray, artists: set[str], exclude_id: int | None = None) -> int | None:
        best_id, best = None, self.threshold
        for e in self.entries.values():
            if e.cluster_id == exclude_id:
                continue
            if artists and e.artists and not (artists & e.artists):
                continue
            sim = float(e.centroid @ vec)
            if sim >= best:
                best_id, best = e.cluster_id, sim
        return best_id

    def prune(self, now: datetime) -> None:
        for cid in [c for c, e in self.entries.items() if e.published_at < now - self.lookback]:
            del self.entries[cid]
