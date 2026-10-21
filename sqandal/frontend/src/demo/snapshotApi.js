// Режим «Реальные новости»: снимок настоящих публикаций (`python -m sqandal.cli snapshot …`),
// прогнанных через пайплайн бэкенда: сюжеты, теги, приоритеты, даты первоисточников — как
// у работающего воркера. Готовые посты есть у CRITICAL и HIGH; остальные (и «жёстче») пишет
// Claude прямо на странице через возможность `sample` — из аккаунта зрителя, с его согласия.
// Бота здесь нет: в канал ничего не уходит, пост копируется в Telegram вручную.
import raw from "./snapshot-data.json";
import { ApiError } from "../lib/api.js";
import { renderPostHtml, renderPostPlain, telegramLength } from "../lib/format.js";
import {
  buildUserPrompt,
  checkNumbers,
  DAY_ONLY,
  normalizePost,
  pickDocs,
  postWarnings,
  sampleErrorText,
  sampleInput,
  sourceLinks,
  truncate,
} from "./newsroom.js";

const clone = (x) => JSON.parse(JSON.stringify(x));
const HOUR = 3600 * 1000;
const ACTIVE = new Set(["new", "drafted", "scheduled"]);
const NO_BOT = "В снимке нет бота: отправка в канал работает на сервере с TG_BOT_TOKEN. Скопируй пост и вставь в Telegram.";
const NO_SAMPLE =
  "Пост пишет Claude из твоего аккаунта — это работает, когда страница открыта в Claude (claude.ai или приложение).";

// `claude.use` отвечает позже первого запуска скрипта; null — эта страница не может звать Claude
let samplePromise = null;
function getSample() {
  if (!samplePromise) {
    samplePromise = window.claude?.use ? window.claude.use("sample").catch(() => null) : Promise.resolve(null);
  }
  return samplePromise;
}

function staleness(originIso, staleAfter, now) {
  const ageH = (now - new Date(originIso).getTime()) / HOUR;
  if (ageH <= staleAfter) return [false, null];
  if (ageH < 48) return [true, `Первоисточник вышел ${Math.floor(ageH)} ч назад`];
  return [true, `Первоисточник вышел ${Math.floor(ageH / 24)} дн. назад`];
}

export function createSnapshotApi() {
  const data = clone(raw);
  const cfg = { ...data.config, bot_ready: false, llm_ready: true, llm_via: "sample" };
  const P = data.prompt;
  const details = new Map(Object.entries(data.details).map(([k, v]) => [Number(k), v]));
  // лента — по-русски: заголовок поста, оригинал — строкой ниже
  const russian = (d) => {
    d.title_orig = d.title_orig ?? d.title;
    d.title = d.current_post?.headline || d.title_orig;
    return d;
  };
  details.forEach(russian);
  let nextPostId = 900;

  // «старый факт» — как housekeeping воркера: по часам зрителя, приоритет не трогаем
  const fresh = (d) => {
    const [stale, reason] = staleness(d.origin.published_at, cfg.stale_after_hours, Date.now());
    if (stale || d.is_stale) {
      d.is_stale = true;
      d.stale_reason = reason ?? d.stale_reason;
    }
    return d;
  };
  const card = (d) => {
    const { items, current_post, versions, ...c } = fresh(d);
    const p = current_post;
    c.post = p
      ? { id: p.id, version: p.version, mode: p.mode, status: p.status, scheduled_at: p.scheduled_at,
          published_at: p.published_at, created_at: p.created_at }
      : null;
    return c;
  };
  const byPost = (pid) => [...details.values()].find((d) => d.current_post?.id === pid);
  const rerender = (p) => {
    p.body_html = renderPostHtml(p, cfg);
    p.body_text = renderPostPlain(p, cfg);
  };
  const docsOf = (d) =>
    pickDocs(d.items, P.max_sources).map((it, i) => ({
      row: it,
      idx: i + 1,
      outlet: it.source_name,
      kind: it.source_kind === "telegram" ? "telegram" : "media",
      url: it.url,
      published: new Date(it.effective_at),
      dateConfidence: it.date_confidence,
      dayOnly: DAY_ONLY.has(it.date_source),
      title: it.title,
      text: truncate(it.summary || "", P.source_chars),
      fwdFrom: it.fwd_from,
    }));

  function addPost(d, gp, mode, docs) {
    const now = new Date().toISOString();
    const p = {
      id: nextPostId++,
      cluster_id: d.id,
      version: (d.versions[0]?.version ?? 0) + 1,
      mode,
      status: "draft",
      flag_emoji: gp.flag_emoji,
      topic_emoji: gp.topic_emoji,
      headline: gp.headline,
      bullets: gp.bullets,
      summary: gp.summary,
      tags: gp.tags,
      sources: sourceLinks(docs.map((x) => x.row)),
      footnotes: [],
      editor_notes: gp.notes_for_editor || null,
      is_publishable: gp.is_publishable,
      event_date: gp.event_date,
      llm_model: "Claude из твоего аккаунта",
      input_tokens: 0,
      output_tokens: 0,
      cache_read_tokens: 0,
      scheduled_at: null,
      published_at: null,
      tg_message_id: null,
      link: null,
      error: null,
      created_at: now,
      updated_at: now,
    };
    rerender(p);
    p.warnings = postWarnings(d, [p.headline, ...p.bullets, p.summary], docs, telegramLength(p, cfg));
    if (gp.bullets.length < 3) p.warnings.push("Меньше трёх фактов — источники бедные");
    if (!gp.is_publishable) p.warnings.unshift("Модель считает, что это не пост: " + (gp.notes_for_editor || "без причины"));
    d.versions = d.versions.map((v) => (v.status === "draft" ? { ...v, status: "superseded" } : v));
    d.versions.unshift({ id: p.id, version: p.version, mode, status: "draft", scheduled_at: null, published_at: null,
                         created_at: now });
    d.current_post = p;
    if (d.status === "new") d.status = "drafted";
    d.needs_refresh = false;
    d.updated_at = now;
    return p;
  }

  const api = {
    demo: true,
    kind: "snapshot",
    async config() {
      return clone(cfg);
    },
    async stats() {
      const all = [...details.values()].map(fresh);
      const active = all.filter((d) => ACTIVE.has(d.status));
      const count = (list, k) => list.reduce((m, d) => ({ ...m, [d[k]]: (m[d[k]] ?? 0) + 1 }), {});
      const posts = all.map((d) => d.current_post).filter(Boolean);
      return {
        ...data.stats,
        by_scene: count(active, "scene"),
        by_priority: count(active, "priority"),
        critical_by_scene: count(active.filter((d) => d.priority === "CRITICAL"), "scene"),
        drafts: posts.filter((p) => p.status === "draft").length,
        scheduled: 0,
        published_24h: 0,
        server_time: new Date().toISOString(),
      };
    },
    async sources() {
      return clone(data.sources).map((x) => ({ ...x, note: x.snapshot_note ?? x.note }));
    },
    async stories({ scene = "all", priority, tag, status = "active", q, sort = "time" } = {}) {
      let list = [...details.values()].map(fresh);
      if (scene !== "all") list = list.filter((d) => d.scene === scene);
      if (priority) list = list.filter((d) => priority.split(",").includes(d.priority));
      if (tag) list = list.filter((d) => d.tags.includes(tag));
      if (status === "active") list = list.filter((d) => ACTIVE.has(d.status));
      else if (status !== "all") list = list.filter((d) => d.status === status);
      if (q) {
        const s = q.toLowerCase();
        list = list.filter(
          (d) => d.title.toLowerCase().includes(s) || (d.title_orig || "").toLowerCase().includes(s)
            || d.artists.some((a) => a.toLowerCase().includes(s)),
        );
      }
      const rank = { CRITICAL: 3, HIGH: 2, NORMAL: 1 };
      const time = (d) => new Date(d.origin.published_at).getTime();
      list.sort((a, b) => (sort === "priority" ? rank[b.priority] - rank[a.priority] || time(b) - time(a) : time(b) - time(a)));
      return { stories: list.map((d) => clone(card(d))), server_time: new Date().toISOString() };
    },
    async story(id) {
      const d = details.get(Number(id));
      if (!d) throw new ApiError(404, "Сюжет не найден");
      return clone(fresh(d));
    },
    async generate(id, mode = "standard") {
      const d = details.get(Number(id));
      if (!d) throw new ApiError(404, "Сюжет не найден");
      const sample = await getSample();
      if (!sample) throw new ApiError(503, NO_SAMPLE);
      fresh(d);
      const docs = docsOf(d);
      const prev = d.current_post;
      const previous = (mode === "harder" || mode === "hardest") && prev
        ? { headline: prev.headline, bullets: prev.bullets, summary: prev.summary }
        : null;
      const user = buildUserPrompt(docs, new Date(), P.audience_tz, d.scene, mode, previous, P.modes);
      let answer;
      try {
        answer = await sample.json(sampleInput(P, user), { modelTier: "default", cache: false });
      } catch (e) {
        throw new ApiError(502, sampleErrorText(e), e?.code);
      }
      let gp;
      try {
        gp = normalizePost(answer, d.scene, P);
      } catch {
        throw new ApiError(502, "Claude ответил не по формату — нажми ещё раз.");
      }
      const p = addPost(d, gp, mode, docs);
      d.tags = data.config.tags.map((x) => x.name).filter((n) => d.tags.includes(n) || gp.tags.includes(n));
      russian(d);
      return clone(p);
    },
    async patchPost(pid, fields) {
      const d = byPost(pid);
      if (!d) throw new ApiError(404, "Пост не найден");
      const p = d.current_post;
      for (const k of ["flag_emoji", "topic_emoji", "headline", "summary"]) {
        if (fields[k] !== undefined && fields[k] !== null) p[k] = String(fields[k]).trim();
      }
      if (Array.isArray(fields.bullets)) p.bullets = fields.bullets.map((b) => b.trim()).filter(Boolean).slice(0, 6);
      p.mode = "manual";
      p.updated_at = new Date().toISOString();
      rerender(p);
      russian(d);
      // как Newsroom.update_post: числа и длину проверяем заново, остальные предупреждения — как были
      const docs = docsOf(d);
      const y = new Date().getUTCFullYear();
      const missing = checkNumbers([p.headline, ...p.bullets, p.summary], docs.map((x) => `${x.title}\n${x.text}`),
                                   new Set([`${y}`, `${y - 1}`, `${y + 1}`]));
      const keep = (p.warnings || []).filter((w) => !w.startsWith("Числа не найдены") && !w.includes("4096"));
      p.warnings = [
        ...(missing.length ? ["Числа не найдены в источниках: " + missing.slice(0, 8).join(", ")] : []),
        ...(telegramLength(p, cfg) > 4096 ? ["Пост длиннее 4096 символов — Telegram не примет"] : []),
        ...keep,
      ];
      d.updated_at = p.updated_at;
      return clone(p);
    },
    async publish() {
      throw new ApiError(501, NO_BOT);
    },
    async schedule() {
      throw new ApiError(501, NO_BOT);
    },
    async unschedule() {
      throw new ApiError(501, NO_BOT);
    },
    async dismiss(id) {
      const d = details.get(Number(id));
      d.status = "dismissed";
      d.updated_at = new Date().toISOString();
      return clone(card(d));
    },
    async restore(id) {
      const d = details.get(Number(id));
      d.status = d.current_post ? "drafted" : "new";
      d.updated_at = new Date().toISOString();
      return clone(card(d));
    },
    async collect() {
      throw new ApiError(501, "Это снимок: новые публикации собирает воркер на сервере.");
    },
    subscribe(onStory, onStats, onStatus) {
      onStatus(false);
      getSample(); // поднять runtime заранее — первая генерация не ждёт
      return () => {};
    },
  };
  return api;
}
