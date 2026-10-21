"""Рендер поста для Telegram: HTML (по умолчанию), MarkdownV2 и plain text.

Структура поста @SQANDALru:
    🇺🇸🚨 <b>Заголовок</b>

    ▪️ факт
    ▪️ факт
    ▪️ факт

    <blockquote>хлёсткое резюме</blockquote>

    📎 TMZ · XXL            ← источники (POST_SHOW_SOURCES)
    🔊 SQANDAL              ← фирменная плашка

Маленькая статья (kind="article") — то же, но вместо пунктов абзацы через пустую строку.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, field

from ..settings import Settings

MAX_TEXT_UTF16 = 4096  # лимит sendMessage после разбора сущностей

_MD2_SPECIAL = re.compile(r"([_*\[\]()~`>#+\-=|{}.!\\])")


@dataclass(slots=True)
class PostContent:
    flag_emoji: str
    topic_emoji: str
    headline: str
    bullets: list[str]
    summary: str
    sources: list[dict] = field(default_factory=list)  # [{"name": "TMZ", "url": "..."}]
    footnotes: list[str] = field(default_factory=list)
    kind: str = "post"  # "post" — пункты, "article" — абзацы маленькой статьи (bullets = абзацы)


def _body(p: PostContent, bullet: str, esc) -> list[str]:
    if p.kind == "article":
        out: list[str] = []
        for i, para in enumerate(x for x in p.bullets if _clean(x)):
            if i:
                out.append("")
            out.append(esc(_clean(para)))
        return out
    return [f"{bullet} {esc(_clean(b))}" for b in p.bullets]


def escape_html(s: str) -> str:
    return html.escape(s or "", quote=False)


def escape_attr(s: str) -> str:
    return html.escape(s or "", quote=True)


def escape_md2(s: str) -> str:
    return _MD2_SPECIAL.sub(r"\\\1", s or "")


def escape_md2_url(url: str) -> str:
    return (url or "").replace("\\", "\\\\").replace(")", "\\)")


def utf16_len(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2


def _clean(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _head(p: PostContent) -> str:
    return (p.flag_emoji or "") + (p.topic_emoji or "")


def render_html(p: PostContent, st: Settings) -> str:
    out = []
    head = _head(p)
    out.append(f"{head} <b>{escape_html(_clean(p.headline))}</b>" if head else f"<b>{escape_html(_clean(p.headline))}</b>")
    out.append("")
    out.extend(_body(p, st.post_bullet, escape_html))
    if p.summary:
        out.append("")
        s = escape_html(_clean(p.summary))
        out.append({
            "quote": f"<blockquote>{s}</blockquote>",
            "bold": f"<b>{s}</b>",
            "italic": f"<i>{s}</i>",
            "plain": s,
        }[st.post_summary_style])
    if p.footnotes:
        out.append("")
        out.extend(f"<i>* {escape_html(_clean(f))}</i>" for f in p.footnotes)
    tail = []
    if st.post_show_sources and p.sources:
        links = []
        for src in p.sources[:4]:
            name = escape_html(src.get("name", ""))
            url = src.get("url")
            links.append(f'<a href="{escape_attr(url)}">{name}</a>' if url else name)
        tail.append("📎 " + " · ".join(links))
    brand = escape_html(st.brand_line)
    tail.append(f'<a href="{escape_attr(st.brand_url)}">{brand}</a>' if st.brand_url else f"<b>{brand}</b>")
    out.append("")
    out.extend(tail)
    return "\n".join(out)


def render_markdown_v2(p: PostContent, st: Settings) -> str:
    out = []
    head = escape_md2(_head(p))
    title = f"*{escape_md2(_clean(p.headline))}*"
    out.append(f"{head} {title}" if head else title)
    out.append("")
    out.extend(_body(p, escape_md2(st.post_bullet), escape_md2))
    if p.summary:
        out.append("")
        s = escape_md2(_clean(p.summary))
        out.append({"quote": f">{s}", "bold": f"*{s}*", "italic": f"_{s}_", "plain": s}[st.post_summary_style])
    if p.footnotes:
        out.append("")
        out.extend(f"_\\* {escape_md2(_clean(f))}_" for f in p.footnotes)
    out.append("")
    if st.post_show_sources and p.sources:
        links = []
        for src in p.sources[:4]:
            name = escape_md2(src.get("name", ""))
            url = src.get("url")
            links.append(f"[{name}]({escape_md2_url(url)})" if url else name)
        out.append("📎 " + " · ".join(links))
    brand = escape_md2(st.brand_line)
    out.append(f"[{brand}]({escape_md2_url(st.brand_url)})" if st.brand_url else f"*{brand}*")
    return "\n".join(out)


def render_plain(p: PostContent, st: Settings) -> str:
    out = [f"{_head(p)} {_clean(p.headline)}".strip(), ""]
    out += _body(p, st.post_bullet, lambda x: x)
    if p.summary:
        out += ["", _clean(p.summary)]
    if p.footnotes:
        out += [""] + [f"* {_clean(f)}" for f in p.footnotes]
    out.append("")
    if st.post_show_sources and p.sources:
        out.append("📎 " + " · ".join(s.get("name", "") for s in p.sources[:4]))
    out.append(st.brand_line)
    return "\n".join(out)


def html_to_visible_text(body_html: str) -> str:
    """Текст, который увидит читатель (для подсчёта лимита 4096)."""
    return html.unescape(re.sub(r"<[^>]+>", "", body_html))


def validate_length(body_html: str) -> int:
    return utf16_len(html_to_visible_text(body_html))
