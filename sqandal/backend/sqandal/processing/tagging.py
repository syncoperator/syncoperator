"""Авто-теги, вотчлист артистов, релевантность и приоритет CRITICAL/HIGH/NORMAL."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from ..models import Priority

_CYR = re.compile(r"[а-яё]", re.I)


def _ru_word(word: str) -> str:
    """Склонения русских имён: Оксимирон → Оксимирон(а|у|ом|е)?, Баста → Баст(а|ы|е|у|ой)."""
    w = re.escape(word)
    low = word.lower()
    if not _CYR.search(word) or len(word) < 4:
        return w
    if low[-1] in "иоуыэю":  # несклоняемые: Тимати, Мияги, Хаски
        return w
    if low[-1] in "ая":
        stem = re.escape(word[:-1])
        return stem + r"(?:а|я|ы|и|е|у|ю|ой|ей|ою)"
    if low[-1] == "й":
        stem = re.escape(word[:-1])
        return stem + r"(?:й|я|ю|ем|е)"
    if low[-1] == "ь":
        stem = re.escape(word[:-1])
        return stem + r"(?:ь|я|ю|ем|е)"
    return w + r"(?:а|у|ом|е|ы|ов)?"


def name_pattern(name: str, case_sensitive: bool = False) -> re.Pattern:
    parts = name.split()
    body = r"\s+".join(_ru_word(p) for p in parts)
    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile(r"(?<![\w$])" + body + r"(?![\w])", flags)


def _match_outside(patterns: list[re.Pattern], exclude: list[re.Pattern], text: str) -> bool:
    """Есть совпадение с patterns, которое не лежит целиком внутри совпадения с exclude."""
    bad_spans = [m.span() for rx in exclude for m in rx.finditer(text)]
    for rx in patterns:
        for m in rx.finditer(text):
            s, e = m.span()
            if not any(bs <= s and e <= be for bs, be in bad_spans):
                return True
    return False


@dataclass(slots=True)
class Artist:
    name: str
    scene: str
    tier: int
    patterns: list[re.Pattern]
    footnote: str | None = None
    exclude: list[re.Pattern] = field(default_factory=list)

    def matches(self, text: str) -> bool:
        return _match_outside(self.patterns, self.exclude, text)


@dataclass(slots=True)
class TagRule:
    name: str
    emoji: str
    weight: int
    patterns: list[re.Pattern]
    exclude: list[re.Pattern] = field(default_factory=list)

    def matches(self, text: str) -> bool:
        return _match_outside(self.patterns, self.exclude, text)


@dataclass(slots=True)
class ClusterSignals:
    tags: list[str]
    artists: list[str]
    sources_count: int
    first_items_within_window: int  # сколько независимых источников за breaking_window_h
    insider_only: bool
    is_stale: bool


class Tagger:
    def __init__(self, tags_cfg: dict[str, Any], artists_cfg: dict[str, Any]):
        flags = re.IGNORECASE | re.UNICODE
        self.rules: list[TagRule] = []
        for name, spec in (tags_cfg.get("tags") or {}).items():
            self.rules.append(
                TagRule(
                    name=name,
                    emoji=spec.get("emoji", ""),
                    weight=int(spec.get("weight", 10)),
                    patterns=[re.compile(p, flags) for p in spec.get("patterns", [])],
                    exclude=[re.compile(p, flags) for p in spec.get("exclude", [])],
                )
            )
        self.rules.sort(key=lambda r: -r.weight)
        self.thresholds = tags_cfg.get("thresholds", {"CRITICAL": 60, "HIGH": 30})
        self.bonuses = tags_cfg.get("bonuses", {})
        self.hiphop = [re.compile(p, flags) for p in tags_cfg.get("hiphop_keywords", [])]
        self.artists: list[Artist] = []
        for scene, entries in (artists_cfg or {}).items():
            for e in entries or []:
                names = [e["name"], *e.get("aliases", [])]
                cs = bool(e.get("cs", False))
                self.artists.append(
                    Artist(
                        name=e["name"],
                        scene=scene,
                        tier=int(e.get("tier", 2)),
                        patterns=[name_pattern(n, cs) for n in names],
                        footnote=e.get("footnote"),
                        exclude=[re.compile(x, 0 if cs else re.IGNORECASE) for x in e.get("exclude", [])],
                    )
                )

    @classmethod
    def from_dir(cls, config_dir: Path) -> "Tagger":
        tags = yaml.safe_load((config_dir / "tags.yaml").read_text(encoding="utf-8")) or {}
        artists_path = config_dir / "artists.yaml"
        artists = yaml.safe_load(artists_path.read_text(encoding="utf-8")) if artists_path.exists() else {}
        return cls(tags, artists or {})

    @property
    def tag_names(self) -> list[str]:
        return [r.name for r in self.rules]

    def emoji_for(self, tag: str) -> str:
        for r in self.rules:
            if r.name == tag:
                return r.emoji
        return ""

    def tag(self, text: str) -> list[str]:
        return [r.name for r in self.rules if r.matches(text)]

    def find_artists(self, text: str) -> list[Artist]:
        return [a for a in self.artists if a.matches(text)]

    def is_hiphop(self, text: str) -> bool:
        return any(p.search(text) for p in self.hiphop) or bool(self.find_artists(text))

    def artist_scene(self, artists: list[str]) -> str | None:
        scenes = {a.scene for a in self.artists if a.name in artists}
        return scenes.pop() if len(scenes) == 1 else None

    def footnotes(self, text: str) -> list[str]:
        return [f"{a.name} — {a.footnote}" for a in self.find_artists(text) if a.footnote]

    def score(self, sig: ClusterSignals) -> tuple[Priority, int]:
        b = self.bonuses
        weights = sorted((r.weight for r in self.rules if r.name in sig.tags), reverse=True)
        score = 0.0
        if weights:
            score += weights[0] + float(b.get("second_tag_factor", 0.3)) * sum(weights[1:])
        extra = max(0, sig.sources_count - 1) * int(b.get("extra_source", 10))
        score += min(extra, int(b.get("extra_source_cap", 30)))
        if sig.first_items_within_window >= int(b.get("breaking_sources", 3)):
            score += int(b.get("breaking", 20))
        tiers = [a.tier for a in self.artists if a.name in sig.artists]
        if tiers:
            score += int(b.get("artist_tier1", 15)) if min(tiers) == 1 else int(b.get("artist_tier2", 5))
        if sig.insider_only:
            score += int(b.get("insider_only", -5))
        if sig.is_stale:
            score += int(b.get("stale", -40))
        s = int(round(score))
        if s >= int(self.thresholds.get("CRITICAL", 60)):
            return Priority.CRITICAL, s
        if s >= int(self.thresholds.get("HIGH", 30)):
            return Priority.HIGH, s
        return Priority.NORMAL, s


def breaking_count(times_by_source: dict, window_h: float) -> int:
    """Сколько разных источников написали о сюжете в первые window_h часов."""
    if not times_by_source:
        return 0
    first = min(times_by_source.values())
    return sum(1 for t in times_by_source.values() if (t - first).total_seconds() <= window_h * 3600)
