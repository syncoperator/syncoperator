// Демо-API в браузере: вымышленные сюжеты (экспорт `python -m sqandal.cli demo --export …`),
// действия редактора работают в памяти. Через ~12 с «прилетает» новый CRITICAL-сюжет.
import raw from "./demo-data.json";
import { ApiError } from "../lib/api.js";
import { renderPostHtml, renderPostPlain } from "../lib/format.js";

const ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$/;

function shiftDates(obj, delta) {
  if (Array.isArray(obj)) return obj.map((x) => shiftDates(x, delta));
  if (obj && typeof obj === "object") {
    const out = {};
    for (const [k, v] of Object.entries(obj)) out[k] = shiftDates(v, delta);
    return out;
  }
  if (typeof obj === "string" && ISO.test(obj)) return new Date(new Date(obj).getTime() + delta).toISOString();
  return obj;
}

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const clone = (x) => JSON.parse(JSON.stringify(x));

const HALO_POST = {
  flag_emoji: "🇺🇸",
  topic_emoji: "🏥",
  headline: "Big Halo госпитализирован после ДТП в Хьюстоне",
  bullets: [
    "Авария — в воскресенье в Хьюстоне, подтвердил менеджер артиста.",
    "По словам менеджера, состояние стабильное.",
    "Шоу в Далласе 22 октября перенесено.",
  ],
  summary: "Сначала здоровье — потом сцена.",
  tags: ["Здоровье", "Тур"],
  editor_notes: "Один источник; детали травм не раскрыты.",
};

function generic(post, mode) {
  // «жёстче» для сюжетов без заготовки: короче и резче, факты те же
  const cut = (s, n) => (s.length > n ? s.slice(0, n).replace(/[\s,;:—-]+\S*$/, "") + "." : s);
  return {
    headline: mode === "hardest" ? cut(post.headline, 40).replace(/\.$/, "") : post.headline,
    bullets: post.bullets.map((b) => cut(b, mode === "hardest" ? 90 : 120)),
    summary: cut(post.summary, mode === "hardest" ? 70 : 100),
  };
}

export function createDemoApi() {
  const delta = Date.now() - new Date(raw.generated_at).getTime();
  const data = shiftDates(clone(raw), delta);
  const cfg = data.config;
  const details = new Map(Object.entries(data.details).map(([k, v]) => [Number(k), v]));
  let nextPostId = 900;
  let nextMsg = 4300;
  const listeners = new Set();

  const card = (d) => {
    const { items, current_post, versions, ...c } = d;
    const p = current_post;
    c.post = p
      ? { id: p.id, version: p.version, mode: p.mode, status: p.status, scheduled_at: p.scheduled_at,
          published_at: p.published_at, created_at: p.created_at }
      : null;
    return c;
  };
  const byPost = (pid) => [...details.values()].find((d) => d.current_post?.id === pid);
  const touch = (d) => {
    d.updated_at = new Date().toISOString();
    const c = card(d);
    listeners.forEach((l) => l.onStory(clone(c)));
    return c;
  };
  const rerender = (p) => {
    p.body_html = renderPostHtml(p, cfg);
    p.body_text = renderPostPlain(p, cfg);
  };

  function newPost(d, fields, mode) {
    const prev = d.current_post;
    const p = {
      id: nextPostId++,
      cluster_id: d.id,
      version: (d.versions[0]?.version ?? 0) + 1,
      mode,
      status: "draft",
      tags: fields.tags ?? prev?.tags ?? d.tags,
      sources: prev?.sources ?? d.items.map((i) => ({ name: i.source_name, url: i.url })).slice(0, 3),
      footnotes: [],
      warnings: prev?.warnings ?? (d.verification === "single_source" ? ["Один источник"] : []),
      editor_notes: fields.editor_notes ?? prev?.editor_notes ?? "",
      is_publishable: true,
      event_date: null,
      llm_model: "claude-sonnet-5-5 (демо)",
      input_tokens: 1800,
      output_tokens: 410,
      cache_read_tokens: 1500,
      scheduled_at: null,
      published_at: null,
      tg_message_id: null,
      link: null,
      error: null,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
      flag_emoji: fields.flag_emoji ?? prev?.flag_emoji ?? "🇺🇸",
      topic_emoji: fields.topic_emoji ?? prev?.topic_emoji ?? "",
      headline: fields.headline ?? prev.headline,
      bullets: fields.bullets ?? prev.bullets,
      summary: fields.summary ?? prev.summary,
    };
    rerender(p);
    d.versions = d.versions.map((v) => (v.status === "draft" ? { ...v, status: "superseded" } : v));
    d.versions.unshift({ id: p.id, version: p.version, mode, status: "draft", scheduled_at: null,
                         published_at: null, created_at: p.created_at });
    d.current_post = p;
    if (d.status === "new") d.status = "drafted";
    return p;
  }

  // новый сюжет «в эфире» — показать, как выглядит входящий CRITICAL
  const incoming = () => {
    const now = Date.now();
    const iso = (min) => new Date(now - min * 60000).toISOString();
    const d = {
      id: 99,
      title: "Jayro Cartel detained at LAX ahead of Glass Mansion release",
      scene: "us", scenes: ["us"], tags: ["Задержание", "Релиз"], artists: [], priority: "CRITICAL",
      priority_score: 85, status: "new",
      origin: { published_at: iso(4), confidence: "high", source_name: "Demo Beat", source_kind: "media",
                url: "https://demo.local/beat/jayro-lax", date_source: "feed_published" },
      first_seen_at: iso(3), last_item_at: iso(1), items_count: 2, sources_count: 2,
      source_names: ["Demo Beat", "Demo Wire"], verification: "multi_source", is_stale: false,
      stale_reason: null, republish_of: null, needs_refresh: false, updated_at: new Date().toISOString(),
      items: [
        { id: 991, source_name: "Demo Beat", source_kind: "rss", url: "https://demo.local/beat/jayro-lax",
          title: "Jayro Cartel detained at LAX ahead of Glass Mansion release",
          summary: "Jayro Cartel was detained by airport police at LAX on Monday, his publicist confirmed. "
            + "No charges have been announced. Glass Mansion is due October 31.",
          published_at: iso(4), effective_at: iso(4), date_source: "feed_published", date_confidence: "high",
          fetched_at: iso(3), lang: "en", tg_views: null, fwd_from: null, similarity: null, date_flags: [] },
        { id: 992, source_name: "Demo Wire", source_kind: "web", url: "https://demo.local/us/jayro-lax",
          title: "Jayro Cartel held by LAX airport police, publicist says no charges yet",
          summary: "Airport police at LAX held Jayro Cartel on Monday. According to his publicist, no charges "
            + "were filed. The rapper's album Glass Mansion is still set for October 31.",
          published_at: iso(2), effective_at: iso(2), date_source: "jsonld", date_confidence: "high",
          fetched_at: iso(1), lang: "en", tg_views: null, fwd_from: null, similarity: 0.83, date_flags: [] },
      ],
      current_post: null,
      versions: [],
    };
    details.set(d.id, d);
    touch(d);
  };

  const api = {
    demo: true,
    async config() {
      return clone(cfg);
    },
    async stats() {
      const active = [...details.values()].filter((d) => ["new", "drafted", "scheduled"].includes(d.status));
      const by = (k) => active.reduce((m, d) => ({ ...m, [d[k]]: (m[d[k]] ?? 0) + 1 }), {});
      const posts = [...details.values()].map((d) => d.current_post).filter(Boolean);
      return {
        ...data.stats,
        by_scene: by("scene"),
        by_priority: by("priority"),
        critical_by_scene: active
          .filter((d) => d.priority === "CRITICAL")
          .reduce((m, d) => ({ ...m, [d.scene]: (m[d.scene] ?? 0) + 1 }), {}),
        drafts: posts.filter((p) => p.status === "draft").length,
        scheduled: posts.filter((p) => p.status === "scheduled").length,
        published_24h: posts.filter((p) => p.status === "published").length,
        last_collect_at: new Date(Date.now() - 4 * 60000).toISOString(),
        server_time: new Date().toISOString(),
      };
    },
    async sources() {
      return clone(data.sources);
    },
    async stories({ scene = "all", priority, tag, status = "active", q, sort = "time" } = {}) {
      await wait(120);
      let list = [...details.values()];
      if (scene !== "all") list = list.filter((d) => d.scene === scene);
      if (priority) list = list.filter((d) => priority.split(",").includes(d.priority));
      if (tag) list = list.filter((d) => d.tags.includes(tag));
      if (status === "active") list = list.filter((d) => ["new", "drafted", "scheduled"].includes(d.status));
      else if (status !== "all") list = list.filter((d) => d.status === status);
      if (q) {
        const s = q.toLowerCase();
        list = list.filter((d) => d.title.toLowerCase().includes(s) || d.items.some((i) => i.title.toLowerCase().includes(s)));
      }
      const rank = { CRITICAL: 3, HIGH: 2, NORMAL: 1 };
      list.sort((a, b) =>
        sort === "priority"
          ? rank[b.priority] - rank[a.priority] || b.origin.published_at.localeCompare(a.origin.published_at)
          : b.origin.published_at.localeCompare(a.origin.published_at),
      );
      return { stories: list.map((d) => clone(card(d))), server_time: new Date().toISOString() };
    },
    async story(id) {
      await wait(90);
      const d = details.get(Number(id));
      if (!d) throw new ApiError(404, "Сюжет не найден");
      return clone(d);
    },
    async generate(id, mode = "standard") {
      await wait(1300);
      const d = details.get(Number(id));
      if (!d.current_post) {
        const base = id === 99
          ? { flag_emoji: "🇺🇸", topic_emoji: "🚨", headline: "Jayro Cartel задержан в аэропорту Лос-Анджелеса",
              bullets: ["Задержание — в понедельник, полиция аэропорта LAX; подтвердил публицист.",
                        "Обвинений пока не предъявлено.", "Альбом «Glass Mansion» по-прежнему выходит 31 октября."],
              summary: "За десять дней до релиза — разговор с полицией вместо промо.", tags: ["Задержание", "Релиз"],
              editor_notes: "Два источника. Причина задержания не названа — не домысливай." }
          : HALO_POST;
        const p = newPost(d, base, "standard");
        touch(d);
        return clone(p);
      }
      const variant = data.harder[d.id] && mode !== "standard" ? data.harder[d.id] : generic(d.current_post, mode);
      const p = newPost(d, mode === "standard" ? {} : variant, mode);
      touch(d);
      return clone(p);
    },
    async patchPost(pid, fields) {
      await wait(150);
      const d = byPost(pid);
      const p = d.current_post;
      Object.assign(p, Object.fromEntries(Object.entries(fields).filter(([, v]) => v !== undefined)), {
        mode: "manual",
        updated_at: new Date().toISOString(),
      });
      rerender(p);
      touch(d);
      return clone(p);
    },
    async publish(pid, confirmStale = false) {
      await wait(700);
      const d = byPost(pid);
      if ((d.is_stale || d.republish_of) && !confirmStale) {
        throw new ApiError(409, `Проверь актуальность: ${d.stale_reason}`, "stale_confirmation_required");
      }
      const p = d.current_post;
      // в демо ничего не уходит в канал — и ссылки на несуществующий пост нет
      Object.assign(p, { status: "published", published_at: new Date().toISOString(), tg_message_id: nextMsg++,
                         link: null });
      d.versions = d.versions.map((v) => (v.id === p.id ? { ...v, status: "published", published_at: p.published_at } : v));
      d.status = "published";
      touch(d);
      return clone(p);
    },
    async schedule(pid, iso) {
      await wait(200);
      if (new Date(iso).getTime() < Date.now() + 30000) throw new ApiError(400, "Время публикации должно быть в будущем");
      const d = byPost(pid);
      Object.assign(d.current_post, { status: "scheduled", scheduled_at: iso });
      d.versions = d.versions.map((v) => (v.id === pid ? { ...v, status: "scheduled", scheduled_at: iso } : v));
      d.status = "scheduled";
      touch(d);
      return clone(d.current_post);
    },
    async unschedule(pid) {
      await wait(150);
      const d = byPost(pid);
      Object.assign(d.current_post, { status: "draft", scheduled_at: null });
      d.versions = d.versions.map((v) => (v.id === pid ? { ...v, status: "draft", scheduled_at: null } : v));
      d.status = "drafted";
      touch(d);
      return clone(d.current_post);
    },
    async dismiss(id) {
      const d = details.get(Number(id));
      d.status = "dismissed";
      return clone(touch(d));
    },
    async restore(id) {
      const d = details.get(Number(id));
      d.status = d.current_post ? "drafted" : "new";
      return clone(touch(d));
    },
    async collect() {
      await wait(300);
      return { queued: true };
    },
    subscribe(onStory, onStats, onStatus) {
      const l = { onStory, onStats };
      listeners.add(l);
      setTimeout(() => onStatus(true), 300);
      const t1 = setTimeout(incoming, 12000);
      const t2 = setInterval(async () => onStats(await api.stats()), 15000);
      return () => {
        listeners.delete(l);
        clearTimeout(t1);
        clearInterval(t2);
      };
    },
  };
  return api;
}
