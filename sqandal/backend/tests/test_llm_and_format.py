"""LLM-генератор (формат запроса к Claude API, ретраи, валидация), фактчек, рендер постов."""

from __future__ import annotations

import json
from datetime import datetime, timezone

import httpx
import pytest

from conftest import run
from sqandal.llm.claude import ClaudeClient
from sqandal.llm.factcheck import check_numbers
from sqandal.llm.generator import GenerationError, PostGenerator
from sqandal.llm.prompts import SYSTEM_PROMPT, SourceDoc, build_user_prompt
from sqandal.publish.formatter import (
    PostContent,
    escape_md2,
    render_html,
    render_markdown_v2,
    render_plain,
    utf16_len,
)

UTC = timezone.utc

GOOD = {
    "is_publishable": True, "scene": "us", "flag_emoji": "🇺🇸", "topic_emoji": "🚨",
    "headline": "Lil Testy задержан в Атланте.",
    "bullets": ["▪️ Задержание — 18 октября, по данным Wire.", "- Залог — $25 000.", "• Залог — $25 000.",
                "Комментариев нет.", "Пятый лишний пункт."],
    "summary": "  Тур под   вопросом. ", "tags": ["Задержание", "Выдуманный"], "event_date": "2026-10-18",
    "notes_for_editor": "ok",
}


def _resp(post: dict | str, stop="end_turn", thinking=True):
    text = post if isinstance(post, str) else json.dumps(post, ensure_ascii=False)
    content = [{"type": "thinking", "thinking": "…", "signature": "sig"}] if thinking else []
    content.append({"type": "text", "text": text})
    return httpx.Response(200, json={
        "id": "msg_1", "type": "message", "role": "assistant", "model": "claude-sonnet-5-5",
        "content": content, "stop_reason": stop,
        "usage": {"input_tokens": 900, "output_tokens": 300, "cache_read_input_tokens": 1200},
    })


def _docs():
    return [SourceDoc(1, "Wire", "media", "https://wire.test/a", datetime(2026, 10, 18, 22, tzinfo=UTC), "high",
                      "Rapper Lil Testy arrested", "Bond was set at $25,000 on October 18."),
            SourceDoc(2, "Инсайд", "telegram", "https://t.me/x/1", datetime(2026, 10, 18, 21, tzinfo=UTC), "high",
                      "Задержан", "Слух <b>без</b> подтверждения & детали", fwd_from="@orig")]


def _gen(settings, responses: list[httpx.Response]):
    seen: list[httpx.Request] = []

    def handler(req):
        seen.append(req)
        return responses.pop(0)

    client = ClaudeClient("sk-test", transport=httpx.MockTransport(handler), max_retries=2)
    return PostGenerator(settings, client), seen, client


def test_request_shape_matches_api_constraints(settings_factory):
    st = settings_factory()
    gen, seen, client = _gen(st, [_resp(GOOD)])
    post, resp = run(gen.generate(_docs(), "us"))
    body = json.loads(seen[0].content)
    req = seen[0]
    assert req.url.path == "/v1/messages"
    assert req.headers["x-api-key"] == "sk-test" and req.headers["anthropic-version"] == "2023-06-01"
    assert body["model"] == "claude-sonnet-5-5"
    # Sonnet 5.5: forced tool_choice и сэмплинг-параметры → 400. Их не должно быть.
    for forbidden in ("tool_choice", "tools", "temperature", "top_p", "top_k", "thinking"):
        assert forbidden not in body
    assert body["output_config"]["effort"] == "low"
    fmt = body["output_config"]["format"]
    assert fmt["type"] == "json_schema" and fmt["schema"]["additionalProperties"] is False
    assert set(fmt["schema"]["required"]) == set(fmt["schema"]["properties"])
    assert body["system"][0]["cache_control"] == {"type": "ephemeral"}
    assert body["system"][0]["text"] == SYSTEM_PROMPT
    user = body["messages"][0]["content"]
    assert '<source id="2" outlet="Инсайд" type="telegram"' in user and 'repost_of="@orig"' in user
    assert "&lt;b&gt;без&lt;/b&gt; подтверждения &amp; детали" in user  # текст источника экранирован
    # постобработка
    assert post.headline == "Lil Testy задержан в Атланте"
    assert post.bullets == ["Задержание — 18 октября, по данным Wire.", "Залог — $25 000.", "Комментариев нет.",
                            "Пятый лишний пункт."]
    assert post.summary == "Тур под вопросом." and post.tags == ["Задержание"]
    assert resp.usage["cache_read_input_tokens"] == 1200
    run(client.close())


def test_max_tokens_retry_and_overload_retry(settings_factory):
    st = settings_factory(llm_max_tokens=1000)
    overloaded = httpx.Response(529, json={"type": "error", "error": {"type": "overloaded_error", "message": "busy"}},
                                headers={"retry-after": "0"})
    gen, seen, client = _gen(st, [overloaded, _resp('{"headline": "обре', stop="max_tokens"), _resp(GOOD)])
    post, _ = run(gen.generate(_docs(), "us"))
    assert len(seen) == 3
    assert json.loads(seen[1].content)["max_tokens"] == 1000
    assert json.loads(seen[2].content)["max_tokens"] == 2000
    assert post.flag_emoji == "🇺🇸"
    run(client.close())


def test_refusal_invalid_json_and_flag_fallback(settings_factory):
    st = settings_factory()
    gen, _, client = _gen(st, [_resp("", stop="refusal")])
    with pytest.raises(GenerationError, match="отказалась"):
        run(gen.generate(_docs(), "us"))
    gen, _, client2 = _gen(st, [_resp("not json")])
    with pytest.raises(GenerationError, match="Невалидный"):
        run(gen.generate(_docs(), "us"))
    gen, _, client3 = _gen(st, [_resp({**GOOD, "flag_emoji": "USA", "scene": "de"})])
    post, _ = run(gen.generate(_docs(), "de"))
    assert post.flag_emoji == "🇩🇪"
    gen, _, client4 = _gen(st, [httpx.Response(400, json={"type": "error", "error": {
        "type": "invalid_request_error", "message": "bad"}})])
    with pytest.raises(GenerationError, match="400"):
        run(gen.generate(_docs(), "us"))
    for c in (client, client2, client3, client4):
        run(c.close())


def test_harder_mode_includes_previous_draft():
    p = build_user_prompt(_docs(), datetime(2026, 10, 19, 12, tzinfo=UTC), "Europe/Moscow", "us", "harder",
                          {"headline": "Старый", "bullets": ["факт"], "summary": "итог"})
    assert "<previous_draft>" in p and "Старый" in p and "ЖЁСТЧЕ" in p
    assert "2026-10-19 15:00 (Europe/Moscow)" in p


def test_headline_keeps_quotes_of_titles():
    from sqandal.llm.generator import unwrap_quotes

    assert unwrap_quotes("«Drake отрицает, что украл хук «Janice STFU»»") == "Drake отрицает, что украл хук «Janice STFU»"
    assert unwrap_quotes("Mayot и VisaGangBeatz выпустили «Осень»") == "Mayot и VisaGangBeatz выпустили «Осень»"
    assert unwrap_quotes("«Huh» и «Осень»") == "«Huh» и «Осень»"
    assert unwrap_quotes('"Jay-Z выбрал себя"') == "Jay-Z выбрал себя"
    assert unwrap_quotes("'a' and 'b'") == "'a' and 'b'"


def test_day_only_source_has_no_fake_time():
    doc = SourceDoc(1, "Hiphop.de", "media", "https://hiphop.de/a", datetime(2026, 10, 20, 10, tzinfo=UTC), "low",
                    "Titel", "Text", day_only=True)
    p = build_user_prompt([doc], datetime(2026, 10, 21, 8, tzinfo=UTC), "Europe/Moscow", "de")
    assert 'published="2026-10-20 (известна только дата, время не указано)"' in p
    assert "13:00" not in p


def test_factcheck_flags_invented_numbers():
    src = ["Bond was set at $25,000 on Oct. 18; released after six hours. Termin 18.10.2026, 2,5 Jahre, § 29 BtMG"]
    post = ["Задержан 18 октября 2026", "Залог — $25 000, вышел через 6 часов", "2,5 года, § 29 BtMG",
            "Ему 27 лет", "Штраф 15 000 евро"]
    assert check_numbers(post, src, {"2026"}) == ["27", "15 000"]


# ── форматтер ───────────────────────────────────────────────────────
def _content():
    return PostContent("🇺🇸", "🚨", "Lil <Testy> & Co задержан", ["Залог — $25 000 (felony).", "Тур — 1 ноября."],
                       "Тур под вопросом.", [{"name": "Wire", "url": "https://wire.test/a?x=1&y=2"},
                                             {"name": "Инсайд", "url": None}],
                       ["Lil Testy — пример сноски"])


def test_render_html(settings_factory):
    html = render_html(_content(), settings_factory())
    assert html.startswith("🇺🇸🚨 <b>Lil &lt;Testy&gt; &amp; Co задержан</b>\n\n▪️ Залог")
    assert "<blockquote>Тур под вопросом.</blockquote>" in html
    assert '📎 <a href="https://wire.test/a?x=1&amp;y=2">Wire</a> · Инсайд' in html
    assert html.endswith('<a href="https://t.me/SQANDALru">🔊 SQANDAL</a>')
    assert "<i>* Lil Testy — пример сноски</i>" in html


def test_render_markdown_v2_and_plain(settings_factory):
    st = settings_factory(post_summary_style="bold", brand_url=None)
    md = render_markdown_v2(_content(), st)
    assert "*Lil <Testy\\> & Co задержан*" in md  # «>» в MarkdownV2 экранируется
    assert "Залог — $25 000 \\(felony\\)\\." in md and "Тур — 1 ноября\\." in md
    assert "[Wire](https://wire.test/a?x=1&y=2)" in md and md.endswith("*🔊 SQANDAL*")
    assert escape_md2("a_b*c[d]e(f)g~h`i>j#k+l-m=n|o{p}q.r!s\\t") == \
        "a\\_b\\*c\\[d\\]e\\(f\\)g\\~h\\`i\\>j\\#k\\+l\\-m\\=n\\|o\\{p\\}q\\.r\\!s\\\\t"
    plain = render_plain(_content(), st)
    assert "<" not in plain.replace("<Testy>", "") and plain.splitlines()[-1] == "🔊 SQANDAL"
    assert utf16_len("🔊") == 2


def test_article_renders_paragraphs_not_bullets():
    from sqandal.llm.prompts import ARTICLE_SCHEMA, OUTPUT_SCHEMA
    from sqandal.publish.formatter import PostContent, render_html, render_markdown_v2, render_plain
    from sqandal.settings import Settings

    st = Settings(_env_file=None)
    art = PostContent("🇺🇸", "⚖️", "Cordae получил иск", ["Лид: иск на $25 млн.", "Второй абзац."], "Итог.",
                      [{"name": "XXL", "url": "https://xxl.test/a"}], kind="article")
    h = render_html(art, st)
    assert "Лид: иск на $25 млн.\n\nВторой абзац." in h and st.post_bullet not in h
    assert "Второй абзац\\." in render_markdown_v2(art, st)
    assert render_plain(art, st).count("\n\n") >= 3
    assert "paragraphs" in ARTICLE_SCHEMA["required"] and "bullets" not in ARTICLE_SCHEMA["properties"]
    assert set(ARTICLE_SCHEMA["properties"]) - {"paragraphs"} == set(OUTPUT_SCHEMA["properties"]) - {"bullets"}


def test_facts_schema_and_instructions():
    from sqandal.llm.prompts import ARTICLE_FROM_FACTS, FACTS_INSTRUCTIONS, FACTS_SCHEMA

    item = FACTS_SCHEMA["properties"]["facts"]["items"]
    assert item["required"] == ["text", "sources"] and item["additionalProperties"] is False
    assert "outlet" in FACTS_INSTRUCTIONS and "<facts>" in ARTICLE_FROM_FACTS
