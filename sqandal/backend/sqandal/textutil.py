"""Нормализация текста и URL, хэши, определение языка."""

from __future__ import annotations

import hashlib
import html as html_lib
import re
import unicodedata
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from bs4 import BeautifulSoup

_TRACKING_PARAMS = {
    "fbclid", "gclid", "yclid", "dclid", "msclkid", "igshid", "mc_cid", "mc_eid",
    "ref", "ref_src", "ref_url", "cmpid", "cmp", "ito", "ocid", "_ga", "spm",
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "utm_id",
}


def canonicalize_url(url: str | None) -> str | None:
    """Канонический URL для дедупа: без трекинга, фрагмента, www, конечного «/»."""
    if not url:
        return None
    url = url.strip()
    try:
        parts = urlsplit(url)
    except ValueError:
        return url
    if not parts.scheme or not parts.netloc:
        return url
    host = parts.hostname.lower() if parts.hostname else ""
    if host.startswith("www."):
        host = host[4:]
    if host.startswith("m.") and host.count(".") >= 2:  # мобильные версии сайтов
        host = host[2:]
    netloc = host + (f":{parts.port}" if parts.port and parts.port not in (80, 443) else "")
    query = [
        (k, v)
        for k, v in parse_qsl(parts.query, keep_blank_values=False)
        if k.lower() not in _TRACKING_PARAMS and not k.lower().startswith("utm_")
    ]
    query.sort()
    path = parts.path or "/"
    if len(path) > 1 and path.endswith("/"):
        path = path.rstrip("/")
    return urlunsplit(("https", netloc, path, urlencode(query), ""))


_WS = re.compile(r"\s+")


def collapse_ws(text: str) -> str:
    return _WS.sub(" ", text).strip()


def html_to_text(value: str | None) -> str:
    """HTML → чистый текст с переносами между блоками."""
    if not value:
        return ""
    if "<" not in value and "&" not in value:
        return collapse_ws(value)
    soup = BeautifulSoup(value, "lxml")
    for tag in soup(["script", "style", "noscript", "iframe"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    return collapse_ws(html_lib.unescape(text))


def normalize_for_hash(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "").lower()
    text = text.replace("ё", "е")
    text = re.sub(r"[^\w\s]", " ", text)
    return collapse_ws(text)


def content_hash(title: str, body: str | None = None) -> str:
    base = normalize_for_hash(title) + "|" + normalize_for_hash(body or "")[:300]
    return hashlib.sha256(base.encode("utf-8")).hexdigest()


def truncate(text: str | None, limit: int) -> str:
    if not text:
        return ""
    text = text.strip()
    if len(text) <= limit:
        return text
    cut = text[:limit]
    space = cut.rfind(" ")
    if space > limit * 0.7:
        cut = cut[:space]
    return cut.rstrip(" ,.;:—-") + "…"


def first_line(text: str, limit: int = 160) -> str:
    for line in (text or "").splitlines():
        line = line.strip(" \t•▪️-—*#")
        if len(line) >= 8:
            return truncate(line, limit)
    return truncate(collapse_ws(text or ""), limit)


_DE_MARKERS = re.compile(
    r"\b(der|die|das|und|nicht|ist|mit|sich|auf|für|eine?n?|wird|wurde|nach|auch|beim|gegen)\b"
)
_EN_MARKERS = re.compile(r"\b(the|and|is|of|to|in|that|with|for|on|was|his|her|after|says|said)\b")
_UK_LETTERS = set("іїєґІЇЄҐ")
_KK_LETTERS = set("әғқңөұүһӘҒҚҢӨҰҮҺ")


def detect_lang(text: str | None) -> str | None:
    """Грубое определение языка: ru / uk / kk / de / en. Без тяжёлых зависимостей."""
    if not text:
        return None
    sample = text[:2000]
    letters = [c for c in sample if c.isalpha()]
    if not letters:
        return None
    cyr = sum(1 for c in letters if "Ѐ" <= c <= "ӿ")
    if cyr / len(letters) > 0.3:
        chars = set(sample)
        if chars & _KK_LETTERS:
            return "kk"
        if chars & _UK_LETTERS:
            return "uk"
        return "ru"
    low = sample.lower()
    de = len(_DE_MARKERS.findall(low)) + 2 * sum(low.count(c) for c in "äöüß")
    en = len(_EN_MARKERS.findall(low))
    if de == 0 and en == 0:
        return None
    return "de" if de > en else "en"
