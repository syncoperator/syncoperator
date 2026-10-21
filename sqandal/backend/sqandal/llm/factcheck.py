"""Пост-проверка сгенерированного текста: каждая цифра должна быть в источниках.

Самая частая галлюцинация LLM в новостях — «уточнённые» числа: суммы, сроки,
возраст, номера статей. Проверка дешёвая и детерминированная: не нашли
число в текстах источников → предупреждение в редакторе.
"""

from __future__ import annotations

import re

_NUM = re.compile(r"(?<![\w.,])\d+(?:[   .,]\d{3})*(?:[.,]\d+)?(?![\w])")

_WORDS = {
    # en
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "fifteen": 15, "twenty": 20, "thirty": 30, "forty": 40,
    "fifty": 50, "hundred": 100, "thousand": 1000, "million": 1000000,
    # de
    "eins": 1, "ein": 1, "eine": 1, "zwei": 2, "drei": 3, "vier": 4, "fünf": 5, "sechs": 6, "sieben": 7,
    "acht": 8, "neun": 9, "zehn": 10, "elf": 11, "zwölf": 12, "zwanzig": 20, "dreißig": 30,
    "hundert": 100, "tausend": 1000, "millionen": 1000000,
    # ru
    "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3, "четыре": 4, "пять": 5, "шесть": 6, "семь": 7,
    "восемь": 8, "девять": 9, "десять": 10, "двадцать": 20, "тридцать": 30, "сто": 100,
    "тысяча": 1000, "тысячи": 1000, "миллион": 1000000, "миллиона": 1000000,
}

_SCALE = re.compile(r"(\d+(?:[.,]\d+)?)\s?(k|m|тыс\.?|млн|млрд|mio\.?|mrd\.?|million|millionen|billion)\b", re.I)
_SCALE_MULT = {"k": 1e3, "тыс": 1e3, "тыс.": 1e3, "m": 1e6, "млн": 1e6, "mio": 1e6, "mio.": 1e6,
               "million": 1e6, "millionen": 1e6, "млрд": 1e9, "mrd": 1e9, "mrd.": 1e9, "billion": 1e9}


def _norm(num: str) -> str:
    """'25 000' / '25,000' / '25.000' → '25000'; '2,5' → '2.5'."""
    s = num.replace(" ", " ").replace(" ", " ")
    if re.fullmatch(r"\d{1,3}([ .,]\d{3})+", s):
        return re.sub(r"[ .,]", "", s)
    return s.replace(",", ".").replace(" ", "")


def numbers_in(text: str) -> set[str]:
    out = {_norm(m.group(0)) for m in _NUM.finditer(text or "")}
    for m in _SCALE.finditer(text or ""):
        out.add(_norm(m.group(1)))  # «$4M» → и «4», и «4000000»
        base = float(m.group(1).replace(",", "."))
        mult = _SCALE_MULT.get(m.group(2).lower().rstrip("."), _SCALE_MULT.get(m.group(2).lower(), 1))
        out.add(_fmt(base * mult))
    return out


def _fmt(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else str(x)


def source_numbers(text: str) -> set[str]:
    nums = numbers_in(text)
    nums |= set(re.findall(r"\d+", text or ""))  # «18.10.2026» → 18, 10, 2026
    low = (text or "").lower()
    for w, v in _WORDS.items():
        if re.search(rf"(?<!\w){re.escape(w)}(?!\w)", low):
            nums.add(str(v))
    return nums


def check_numbers(post_parts: list[str], source_texts: list[str], allow: set[str] | None = None) -> list[str]:
    """Числа из поста, которых нет в источниках (allow — заведомо безопасные, напр. текущий год)."""
    src = set()
    for t in source_texts:
        src |= source_numbers(t)
    allow = allow or set()
    missing: list[str] = []
    for part in post_parts:
        for m in _NUM.finditer(part or ""):
            raw = m.group(0)
            n = _norm(raw)
            if n in src or n in allow:
                continue
            # «18» в «18 октября» найдётся как «18» в «Oct. 18»; «2,5 года» ↔ «2.5 years» — уже норм.
            if n.lstrip("0") in src:
                continue
            if raw not in missing:
                missing.append(raw)
    return missing
