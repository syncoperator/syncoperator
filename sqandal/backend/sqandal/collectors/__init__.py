"""Сборщики: RSS, веб-листинги (httpx/Playwright), Telegram-каналы (Telethon)."""

from __future__ import annotations

from pathlib import Path

import yaml

from .base import CollectResult, RawItem, SourceConfig, SourcesFile


def load_sources(config_dir: Path) -> list[SourceConfig]:
    path = config_dir / "sources.yaml"
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    parsed = SourcesFile.model_validate(data)
    keys = [s.key for s in parsed.sources]
    dupes = {k for k in keys if keys.count(k) > 1}
    if dupes:
        raise ValueError(f"Дубли key в sources.yaml: {sorted(dupes)}")
    return parsed.sources


__all__ = ["CollectResult", "RawItem", "SourceConfig", "load_sources"]
