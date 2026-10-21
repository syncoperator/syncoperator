"""Работа со временем: всё внутри системы — aware UTC."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo

from dateutil import parser as dateparser

UTC = timezone.utc


def utcnow() -> datetime:
    return datetime.now(UTC)


def ensure_utc(dt: datetime | None, assume_tz: str | None = None) -> datetime | None:
    """Naive → считаем во временной зоне `assume_tz` (или UTC), aware → переводим в UTC."""
    if dt is None:
        return None
    if dt.tzinfo is None:
        tz = ZoneInfo(assume_tz) if assume_tz else UTC
        dt = dt.replace(tzinfo=tz)
    return dt.astimezone(UTC)


def parse_datetime(value: str | None, assume_tz: str | None = None) -> tuple[datetime | None, bool]:
    """Парсит ISO-8601 / RFC-822 / «человеческие» даты.

    Возвращает (datetime UTC | None, had_timezone). had_timezone=False означает,
    что часовой пояс угадан (assume_tz/UTC) — уверенность в дате ниже.
    """
    if not value:
        return None, False
    s = str(value).strip()
    if not s:
        return None, False
    dt: datetime | None = None
    try:
        dt = dateparser.isoparse(s)
    except (ValueError, OverflowError):
        try:
            dt = parsedate_to_datetime(s)
        except (TypeError, ValueError, IndexError):
            try:
                dt = dateparser.parse(s, fuzzy=False)
            except (ValueError, OverflowError, TypeError):
                return None, False
    if dt is None:
        return None, False
    had_tz = dt.tzinfo is not None
    # «2026-10-19» без времени — isoparse даст полночь; это точность «день»
    return ensure_utc(dt, assume_tz), had_tz


_URL_DATE = re.compile(r"/(20\d{2})[/-](0[1-9]|1[0-2])[/-](0[1-9]|[12]\d|3[01])(?:/|-|$)")


def date_from_url(url: str | None) -> datetime | None:
    """/2026/10/16/slug → 2026-10-16 12:00 UTC (середина дня: точность «день»)."""
    if not url:
        return None
    m = _URL_DATE.search(url)
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    try:
        return datetime(y, mo, d, 12, 0, tzinfo=UTC)
    except ValueError:
        return None


def hours_between(a: datetime, b: datetime) -> float:
    return (a - b).total_seconds() / 3600.0


def fmt_local(dt: datetime, tz: str) -> str:
    return dt.astimezone(ZoneInfo(tz)).strftime("%d.%m.%Y %H:%M")


__all__ = [
    "UTC",
    "utcnow",
    "ensure_utc",
    "parse_datetime",
    "date_from_url",
    "hours_between",
    "fmt_local",
    "timedelta",
]
