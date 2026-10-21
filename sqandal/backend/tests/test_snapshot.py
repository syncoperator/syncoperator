"""Снимок реальных новостей (`cli snapshot`): настоящие публикации идут через пайплайн с часами
«как у воркера», подготовленные посты — через Newsroom с фактчеком чисел."""

from __future__ import annotations

import json

from conftest import BACKEND, run
from sqandal.snapshot import build_snapshot

SNAP = BACKEND / "snapshot" / "2026-10-21"


def test_snapshot_builds_from_real_items(settings_factory, tmp_path):
    out = tmp_path / "snapshot.json"
    res = run(build_snapshot(settings_factory(), SNAP, out))
    assert res["posts"] == 19 and res["unused_posts"] == []

    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["kind"] == "snapshot" and data["config"]["snapshot_at"].startswith("2026-10-21T08:10")
    assert data["config"]["bot_ready"] is False  # в снимке нет бота
    assert data["prompt"]["system"].startswith("Ты — выпускающий редактор") and data["prompt"]["schema"]["required"]

    def story(word):
        return next(s for s in data["stories"] if word in s["title"])

    # приоритет считается при приходе публикации, «старый факт» — на момент снимка
    q50 = story("BloodHound")
    assert q50["priority"] == "CRITICAL" and q50["is_stale"]
    assert story("Ивана Золо")["priority"] == "CRITICAL"
    tory = story("Tory Lanez")
    assert tory["sources_count"] == 2 and tory["priority"] == "HIGH"
    # XXL пересказал факт, которому уже три дня: тот же сюжет, и он сразу «старый»
    drake = story("Janice")
    assert drake["sources_count"] == 2 and drake["priority"] == "NORMAL" and drake["is_stale"]

    # дата из шапки статьи — только день, уверенность низкая
    yak = data["details"][str(story("Yakary")["id"])]
    assert yak["origin"]["date_source"] == "page_day" and yak["origin"]["confidence"] == "low"
    assert yak["items"][0]["date_flags"] == ["day_only"]

    # «Future of Battlerap» — не рэпер Future
    assert story("Ssynic")["artists"] == ["PA Sports"]

    # русские посты проходят фактчек: ни одного числа не из источников; кавычки названий целы
    posts = [d["current_post"] for d in data["details"].values() if d.get("current_post")]
    assert len(posts) == 19
    for p in posts:
        assert not any(w.startswith("Числа не найдены") for w in p["warnings"]), (p["headline"], p["warnings"])
    assert any(p["headline"].endswith("«Janice STFU»") for p in posts)
