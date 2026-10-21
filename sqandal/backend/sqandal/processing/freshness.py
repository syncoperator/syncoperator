"""Верификация времени: «когда это реально вышло» и «не протухло ли».

Главная защита от публикации старых фактов:
  1. у каждой публикации — effective_at (проверенная дата) + уверенность;
  2. у сюжета — origin_published_at = самая ранняя effective_at по всем источникам;
  3. сюжет старше STALE_AFTER_HOURS помечается is_stale → кнопка «Отправить»
     требует явного подтверждения;
  4. похожесть на уже опубликованный сюжет (republish_of_cluster_id) — тоже стоп.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta

from ..models import Confidence, DateSource

FUTURE_TOLERANCE = timedelta(minutes=15)

_BASE_CONFIDENCE = {
    DateSource.telegram: Confidence.high,
    DateSource.telegram_forward: Confidence.high,
    DateSource.feed_published: Confidence.high,
    DateSource.jsonld: Confidence.high,
    DateSource.meta: Confidence.high,
    DateSource.feed_updated: Confidence.medium,
    DateSource.time_tag: Confidence.medium,
    DateSource.url: Confidence.low,
    DateSource.page_day: Confidence.low,
    DateSource.fetched: Confidence.low,
}

DAY_ONLY = {DateSource.url, DateSource.page_day}


@dataclass(slots=True)
class EffectiveTime:
    at: datetime
    confidence: Confidence
    flags: list[str]


def effective_time(
    published_at: datetime | None,
    date_source: DateSource,
    fetched_at: datetime,
    date_has_tz: bool = True,
    max_age_days: int = 7,
) -> EffectiveTime:
    flags: list[str] = []
    if published_at is None:
        return EffectiveTime(fetched_at, Confidence.low, ["no_date"])
    conf = _BASE_CONFIDENCE.get(date_source, Confidence.low)
    if not date_has_tz and conf == Confidence.high:
        conf = Confidence.medium
        flags.append("tz_guessed")
    if date_source in DAY_ONLY and published_at.date() == fetched_at.date():
        # известен только день (из URL или шапки статьи), и это день обнаружения —
        # время обнаружения точнее условного полудня
        return EffectiveTime(min(fetched_at, published_at.replace(hour=23, minute=59)), Confidence.low, ["day_only"])
    if published_at > fetched_at + FUTURE_TOLERANCE:
        # дата из будущего = кривой часовой пояс/CMS; верим только факту обнаружения
        return EffectiveTime(fetched_at, Confidence.low, ["future_date"])
    if date_source in DAY_ONLY:
        flags.append("day_only")
    if fetched_at - published_at > timedelta(days=max_age_days):
        flags.append("archive")
    return EffectiveTime(published_at, conf, flags)


def assess_staleness(origin_at: datetime, now: datetime, stale_after_hours: float) -> tuple[bool, str | None]:
    age_h = (now - origin_at).total_seconds() / 3600
    if age_h <= stale_after_hours:
        return False, None
    if age_h < 48:
        return True, f"Первоисточник вышел {int(age_h)} ч назад"
    return True, f"Первоисточник вышел {int(age_h // 24)} дн. назад"


def min_confidence(a: Confidence, b: Confidence) -> Confidence:
    order = {Confidence.low: 0, Confidence.medium: 1, Confidence.high: 2}
    return a if order[a] <= order[b] else b
