"""LLM-генератор постов: источники сюжета → Claude (structured output) → валидный пост."""

from __future__ import annotations

import json
import logging
import re
from typing import Literal

from pydantic import BaseModel, ValidationError, field_validator

from ..settings import Settings
from ..timeutil import utcnow
from .claude import ClaudeClient, ClaudeError, ClaudeResponse
from .prompts import OUTPUT_SCHEMA, SYSTEM_PROMPT, TAG_NAMES, SourceDoc, build_user_prompt

log = logging.getLogger(__name__)

SCENE_FLAGS = {"ru": "🇷🇺", "de": "🇩🇪", "us": "🇺🇸", "intl": "🌍"}
_FLAG_RE = re.compile(r"^([\U0001F1E6-\U0001F1FF]{2}|🏴[\U000E0061-\U000E007A]+\U000E007F|🌍|🌎|🌏)$")
_BULLET_PREFIX = re.compile(r"^\s*(?:[-–—•▪●◦*\u25AB\uFE0E\uFE0F]+|\d+[.)])\s*")


class GenerationError(RuntimeError):
    pass


_WRAP = {"«": "»", "“": "”", "„": "“", '"': '"', "'": "'"}


def _balanced(text: str) -> bool:
    depth = 0
    for ch in text:
        if ch == "«":
            depth += 1
        elif ch == "»":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


def unwrap_quotes(v: str) -> str:
    """Снять кавычки, в которые модель обернула ВЕСЬ заголовок. Кавычки названия в конце
    («…украл хук «Janice STFU»») остаются: снимаем пару, только если она внешняя."""
    while len(v) >= 2 and _WRAP.get(v[0]) == v[-1]:
        inner = v[1:-1]
        if not _balanced(inner) or (v[0] in "\"'" and v[0] in inner):
            break
        v = inner.strip()
    return v


class GeneratedPost(BaseModel):
    is_publishable: bool
    scene: Literal["ru", "de", "us", "intl"]
    flag_emoji: str
    topic_emoji: str
    headline: str
    bullets: list[str]
    summary: str
    tags: list[str]
    event_date: str | None
    notes_for_editor: str

    @field_validator("headline")
    @classmethod
    def _headline(cls, v: str) -> str:
        v = unwrap_quotes(re.sub(r"\s+", " ", v).strip())
        return v[:-1] if v.endswith(".") and not v.endswith("..") else v

    @field_validator("bullets")
    @classmethod
    def _bullets(cls, v: list[str]) -> list[str]:
        out, seen = [], set()
        for b in v:
            b = _BULLET_PREFIX.sub("", re.sub(r"\s+", " ", b)).strip()
            if b and b.lower() not in seen:
                seen.add(b.lower())
                out.append(b)
        return out[:4]

    @field_validator("summary", "notes_for_editor")
    @classmethod
    def _strip(cls, v: str) -> str:
        return re.sub(r"\s+", " ", v or "").strip()

    @field_validator("tags")
    @classmethod
    def _tags(cls, v: list[str]) -> list[str]:
        return [t for t in TAG_NAMES if t in set(v)]


class PostGenerator:
    def __init__(self, settings: Settings, client: ClaudeClient):
        self.settings = settings
        self.client = client

    def _body(self, user_prompt: str, max_tokens: int) -> dict:
        return {
            "model": self.settings.llm_model,
            "max_tokens": max_tokens,
            # системный промпт одинаковый для всех сюжетов → кэш (TTL 5 мин продлевается на каждом чтении)
            "system": [{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
            "messages": [{"role": "user", "content": user_prompt}],
            "output_config": {
                "effort": self.settings.llm_effort,
                "format": {"type": "json_schema", "schema": OUTPUT_SCHEMA},
            },
        }

    async def generate(
        self,
        docs: list[SourceDoc],
        scene_hint: str,
        mode: str = "standard",
        previous: dict | None = None,
    ) -> tuple[GeneratedPost, ClaudeResponse]:
        if not docs:
            raise GenerationError("Нет источников для генерации")
        prompt = build_user_prompt(docs, utcnow(), self.settings.audience_tz, scene_hint, mode, previous)
        max_tokens = self.settings.llm_max_tokens
        resp: ClaudeResponse | None = None
        for attempt in range(2):
            try:
                resp = await self.client.create(self._body(prompt, max_tokens))
            except ClaudeError as e:
                raise GenerationError(str(e)) from e
            if resp.stop_reason == "max_tokens" and attempt == 0:
                max_tokens *= 2  # мысли съели бюджет — даём больше и повторяем
                continue
            break
        assert resp is not None
        if resp.stop_reason == "refusal":
            raise GenerationError("Модель отказалась писать этот пост — проверь источник вручную")
        try:
            post = GeneratedPost.model_validate(json.loads(resp.text))
        except (json.JSONDecodeError, ValidationError) as e:
            raise GenerationError(f"Невалидный ответ модели ({resp.stop_reason}): {e}") from e
        if not _FLAG_RE.match(post.flag_emoji.strip()):
            post.flag_emoji = SCENE_FLAGS.get(post.scene, "🌍")
        if not post.bullets:
            raise GenerationError("Модель не вернула ни одного факта")
        return post, resp
