// Режим «Реальные новости» в артефакте: сюжеты лежат в базе артефакта (возможность `db`).
// Кнопка «Обновить» запускает задачу в облаке (коннектор Claude Code Remote → fire_trigger):
// Claude собирает свежие новости с сайтов из config/sources.yaml, пишет по каждой русский пост
// и складывает сюжеты в базу — лента обновляется сама (onSnapshot). Задание для сбора —
// artifact/collector-prompt.md. В канал отсюда ничего не уходит: бота нет.
// Нет базы (страница открыта не в Claude) → createLiveApi вернёт null, и откроется встроенный снимок.
import raw from "./snapshot-data.json";
import { ApiError } from "../lib/api.js";
import { renderPostHtml, renderPostPlain, telegramLength } from "../lib/format.js";
import {
  buildArticleFromFactsPrompt,
  buildFactsPrompt,
  buildUserPrompt,
  checkNumbers,
  DAY_ONLY,
  normalizeArticle,
  normalizeFacts,
  normalizePost,
  pickDocs,
  postWarnings,
  sampleErrorText,
  sampleInput,
  scoreStory,
  sourceLinks,
  truncate,
} from "./newsroom.js";

export const TRIGGER_ID = "trig_01YXBCesqS7spDSaUvSzGTDE";
const CONNECTOR = "Claude Code Remote";
const RUN_TIMEOUT_MS = 30 * 60 * 1000; // запуск без ответа дольше — считаем сорвавшимся
const KEEP_MS = 7 * 24 * 3600 * 1000; // в ленте — последние 7 дней
const HOUR = 3600 * 1000;
const ACTIVE = new Set(["new", "drafted", "scheduled"]);
const SCENES = new Set(["ru", "de", "us", "intl"]);
const NO_BOT = "Бота здесь нет: в канал отсюда ничего не уходит. Скопируй пост и вставь в Telegram.";
// что читает облачный сбор по кнопке (artifact/collector-prompt.md); BILD и rap.de ему недоступны
const CLOUD = new Set(["xxl", "tmz_hiphop", "tmz_rss", "hiphopdx", "complex", "thesource", "allhiphop", "hotnewhiphop",
                       "hiphopde", "raptastisch", "rap2soul", "watson", "theflow", "tg_rhymes", "tg_theflow", "tg_nr",
                       "tg_mash", "tg_baza", "tg_shot"]);

const use = (name) => (window.claude?.use ? window.claude.use(name).catch(() => null) : Promise.resolve(null));
const t = (iso) => new Date(iso).getTime();
const clone = (x) => JSON.parse(JSON.stringify(x));

function mcpErrorText(e) {
  switch (e?.code) {
    case "server_not_connected":
      return "Коннектор «Claude Code Remote» не подключён: добавь его в claude.ai → Настройки → Коннекторы.";
    case "needs_reauth":
      return "Переподключи «Claude Code Remote» в настройках claude.ai и нажми «Обновить» ещё раз.";
    case "not_in_manifest":
    case "approval_required":
    case "consent_required":
    case "not_granted":
      return "Запуск сбора не разрешён для этой страницы. Нажми «Обновить» ещё раз и разреши, когда Claude спросит.";
    case "blocked_by_policy":
      return "Запуск закрыт политикой организации.";
    case "selection_required":
      return "Выбери, какой коннектор «Claude Code Remote» использовать, и нажми «Обновить» ещё раз.";
    case "tool_error":
      return "Задача сбора не запустилась: " + String(e?.message || "").slice(0, 160);
    case "server_unavailable":
    case "upstream_error":
      return "Сервис запуска не ответил. Если сбор не начнётся за пару минут — нажми «Обновить» ещё раз.";
    default:
      return "Не удалось запустить сбор.";
  }
}

function staleness(originIso, staleAfter, now) {
  const ageH = (now - t(originIso)) / HOUR;
  if (ageH <= staleAfter) return [false, null];
  if (ageH < 48) return [true, `Первоисточник вышел ${Math.floor(ageH)} ч назад`];
  return [true, `Первоисточник вышел ${Math.floor(ageH / 24)} дн. назад`];
}

export async function createLiveApi() {
  const db = await use("db");
  if (!db) return null;
  getSample(); // поднять runtime заранее — первая генерация не ждёт

  const P = raw.prompt;
  const scoring = raw.scoring;
  const weights = Object.fromEntries(raw.config.tags.map((x) => [x.name, x.weight]));
  const tagOrder = raw.config.tags.map((x) => x.name);
  const cfg = {
    ...raw.config,
    snapshot_at: null,
    bot_ready: false,
    llm_ready: true,
    llm_via: "sample",
    collect_via: "trigger",
    articles: true,
  };
  const docs = new Map(); // id → тело документа
  const details = new Map(); // id → сюжет в форме API
  const listeners = new Set();
  let status = null;
  let quietError = false; // ошибку запуска страница уже показала сама
  const openedAt = Date.now();

  // ── документ базы → сюжет дашборда ─────────────────────────────
  function toDetail(id, doc) {
    const srcs = (Array.isArray(doc.sources) ? doc.sources : []).filter((s) => s && s.url && s.published_at);
    srcs.sort((a, b) => t(a.published_at) - t(b.published_at));
    const items = srcs.map((s, i) => ({
      id: `${id}-${i}`,
      source_name: String(s.name || new URL(s.url).hostname),
      source_kind: /^https?:\/\/t\.me\//.test(s.url) ? "telegram" : "web",
      url: s.url,
      title: String(s.title || s.url),
      summary: String(s.facts || ""),
      published_at: s.published_at,
      effective_at: s.published_at,
      date_source: s.date_source || "fetched",
      date_confidence: s.date_confidence || "low",
      fetched_at: doc.collected_at || s.published_at,
      lang: s.lang || null,
      tg_views: null,
      fwd_from: null,
      similarity: null,
      date_flags: DAY_ONLY.has(s.date_source) ? ["day_only"] : [],
    }));
    const names = [...new Set(items.map((i) => i.source_name))];
    const o = items[0];
    const tags = tagOrder.filter((x) => (doc.tags || []).includes(x));
    const artists = (Array.isArray(doc.artists) ? doc.artists : []).map(String);
    const [priority, score] = scoreStory({ tags, artists, items }, weights, scoring, cfg.stale_after_hours);
    const scene = SCENES.has(doc.scene) ? doc.scene : "intl";
    const d = {
      id,
      title: String(doc.title_ru || doc.post?.headline || o?.title || "Без заголовка"),
      title_orig: o?.title ?? null,
      scene,
      scenes: [scene],
      tags,
      artists,
      priority,
      priority_score: score,
      status: doc.status === "dismissed" ? "dismissed" : "new",
      origin: {
        published_at: o?.published_at ?? doc.collected_at,
        confidence: o?.date_confidence ?? "low",
        source_name: o?.source_name ?? null,
        source_kind: o?.source_kind ?? "web",
        url: o?.url ?? null,
        date_source: o?.date_source ?? "fetched",
      },
      first_seen_at: doc.collected_at,
      last_item_at: doc.updated_at || doc.collected_at,
      items_count: items.length,
      sources_count: names.length,
      source_names: names,
      verification: items.length && items.every((i) => i.source_kind === "telegram")
        ? "insider_only"
        : names.length >= 2 ? "multi_source" : "single_source",
      is_stale: false,
      stale_reason: null,
      republish_of: null,
      needs_refresh: false,
      updated_at: doc.updated_at || doc.collected_at,
      items,
      versions: [],
      current_post: null,
    };
    d.current_post = toPost(d, doc.post, "post");
    d.current_article = toPost(d, doc.article, "article");
    d.facts = Array.isArray(doc.facts?.items)
      ? doc.facts.items.filter((f) => f && f.text).map((f) => ({ text: String(f.text), sources: (f.sources || []).map(String), on: f.on !== false }))
      : null;
    if (d.current_post) {
      const post = d.current_post;
      d.versions = [{ id: post.id, version: post.version, mode: post.mode, status: "draft", scheduled_at: null,
                      published_at: null, created_at: post.created_at }];
    }
    return d;
  }

  function toPost(d, p, kind) {
    const lines = kind === "article" ? p?.paragraphs : p?.bullets;
    if (!p || !p.headline || !Array.isArray(lines)) return null;
    const post = {
      id: `${d.id}:${kind === "article" ? "a" : ""}${p.version || 1}`,
      kind,
      cluster_id: d.id,
      version: p.version || 1,
      mode: p.mode || "standard",
      status: "draft",
      flag_emoji: p.flag_emoji || "",
      topic_emoji: p.topic_emoji || "",
      headline: String(p.headline),
      bullets: lines.map(String),
      summary: String(p.summary || ""),
      tags: d.tags,
      sources: sourceLinks(d.items),
      footnotes: [],
      editor_notes: p.notes || null,
      from_facts: p.from_facts ?? null,
      is_publishable: true,
      event_date: null,
      llm_model: p.model || "Claude",
      input_tokens: 0,
      output_tokens: 0,
      cache_read_tokens: 0,
      scheduled_at: null,
      published_at: null,
      tg_message_id: null,
      link: null,
      error: null,
      created_at: p.created_at || d.updated_at,
      updated_at: p.updated_at || p.created_at || d.updated_at,
      warnings: [],
    };
    post.body_html = renderPostHtml(post, cfg);
    post.body_text = renderPostPlain(post, cfg);
    return post;
  }

  const docsFor = (d) =>
    pickDocs(d.items, P.max_sources).map((it, i) => ({
      row: it,
      idx: i + 1,
      outlet: it.source_name,
      kind: "media",
      url: it.url,
      published: new Date(it.effective_at),
      dateConfidence: it.date_confidence,
      dayOnly: DAY_ONLY.has(it.date_source),
      title: it.title,
      text: truncate(it.summary || "", P.source_chars),
      fwdFrom: null,
    }));

  // свежесть и предупреждения — по часам зрителя, на каждый показ
  function fresh(d) {
    const [stale, reason] = staleness(d.origin.published_at, cfg.stale_after_hours, Date.now());
    d.is_stale = stale;
    d.stale_reason = reason;
    for (const p of [d.current_post, d.current_article]) {
      if (!p) continue;
      p.warnings = postWarnings(d, [p.headline, ...p.bullets, p.summary], docsFor(d), telegramLength(p, cfg));
      if (p.kind === "post" && p.bullets.length < 3) p.warnings.push("Меньше трёх фактов — источники бедные");
    }
    return d;
  }
  const visible = (d) => Date.now() - t(d.origin.published_at) <= KEEP_MS;
  const card = (d) => {
    const { items, current_post, current_article, versions, ...c } = fresh(d);
    const p = current_post;
    c.post = p ? { id: p.id, version: p.version, mode: p.mode, status: p.status, scheduled_at: null,
                   published_at: null, created_at: p.created_at } : null;
    return c;
  };
  const byPost = (pid) =>
    [...details.values()].find((d) => d.current_post?.id === pid || d.current_article?.id === pid);
  const ref = (id) => db.doc(`stories/${id}`);

  async function write(id, patch) {
    try {
      await ref(id).update(patch);
    } catch (e) {
      throw new ApiError(e?.code === "invalid_argument" ? 403 : 503,
        e?.code === "invalid_argument" ? "Менять сюжеты может только владелец страницы." : "База не ответила — попробуй ещё раз.");
    }
  }

  // ── подписки ───────────────────────────────────────────────────
  let firstReady;
  const first = new Promise((r) => (firstReady = r));
  setTimeout(() => firstReady(), 6000); // база молчит — показываем то, что есть
  let loadError = null;
  db.collection("stories").onSnapshot(
    (snap) => {
      for (const ch of snap.docChanges()) {
        const id = ch.doc.id;
        if (ch.type === "removed") {
          docs.delete(id);
          details.delete(id);
          continue;
        }
        const body = ch.doc.data();
        docs.set(id, body);
        let d;
        try {
          d = toDetail(id, body);
        } catch {
          continue; // кривой документ не роняет ленту
        }
        details.set(id, d);
        // новые сюжеты после открытия страницы — в эфир (сирена для CRITICAL); догрузка кэша — тихо
        const late = t(body.collected_at || 0) >= openedAt - 60000 || ch.type === "modified";
        if (late && visible(d)) listeners.forEach((l) => l.onStory(clone(card(d))));
      }
      if (!snap.metadata.fromCache) firstReady();
      listeners.forEach((l) => l.statsSoon());
    },
    (e) => {
      loadError = e;
      firstReady();
    },
  );
  db.doc("meta/status").onSnapshot((s) => {
    const prev = status;
    status = s.exists ? s.data() : null;
    listeners.forEach((l) => l.statsSoon());
    if (prev && status && prev.state !== status.state) {
      if (status.state === "done") listeners.forEach((l) => l.toast?.({ kind: "ok", text: `Сбор готов: ${status.message || "лента обновлена"}` }));
      if (status.state === "error" && !quietError) {
        listeners.forEach((l) => l.toast?.({ kind: "err", text: `Сбор не удался: ${status.message || "без подробностей"}` }));
      }
      quietError = false;
      if (status.state === "running") listeners.forEach((l) => l.toast?.({ text: "Claude начал сбор — новости появятся по мере готовности" }));
    }
  });

  const runBusy = () =>
    status && ["requested", "running"].includes(status.state)
    && Date.now() - t(status.started_at || status.requested_at || 0) < RUN_TIMEOUT_MS;

  const api = {
    demo: true,
    kind: "live",
    async config() {
      await first;
      if (loadError && !details.size) throw new ApiError(503, "База дашборда не открылась — обнови страницу.");
      return clone(cfg);
    },
    async stats() {
      await first;
      const active = [...details.values()].filter((d) => ACTIVE.has(d.status) && visible(d));
      const count = (list, k) => list.reduce((m, d) => ({ ...m, [d[k]]: (m[d[k]] ?? 0) + 1 }), {});
      const failed = new Set(status?.failed_sources || []);
      return {
        by_scene: count(active, "scene"),
        by_priority: count(active, "priority"),
        critical_by_scene: count(active.filter((d) => d.priority === "CRITICAL"), "scene"),
        total: active.length,
        drafts: 0,
        scheduled: 0,
        published_24h: 0,
        last_collect_at: status?.finished_at ?? null,
        sources_ok: raw.sources.filter((s) => s.enabled && !failed.has(s.name)).length,
        sources_failing: raw.sources.filter((s) => s.enabled && failed.has(s.name)).length,
        collect: status ? { ...status, busy: Boolean(runBusy()) } : null,
        server_time: new Date().toISOString(),
      };
    },
    async sources() {
      const failed = new Set(status?.failed_sources || []);
      const at = status?.finished_at ?? null;
      return raw.sources.map((s) => {
        const n = { ...s };
        if (s.enabled && !CLOUD.has(s.key)) {
          return { ...n, enabled: false, note: "Недоступен для сбора по кнопке — читается только серверной версией" };
        }
        if (!s.enabled || !at) return n;
        return failed.has(s.name)
          ? { ...n, last_fetch_at: at, last_error: "Не открылся при последнем сборе", consecutive_failures: 1 }
          : { ...n, last_fetch_at: at, last_success_at: at, last_error: null, consecutive_failures: 0,
              items_total: [...details.values()].filter((d) => d.source_names.includes(s.name)).length };
      });
    },
    async stories({ scene = "all", priority, tag, status: st = "active", q, sort = "time" } = {}) {
      await first;
      let list = [...details.values()].filter(visible).map(fresh);
      if (scene !== "all") list = list.filter((d) => d.scene === scene);
      if (priority) list = list.filter((d) => priority.split(",").includes(d.priority));
      if (tag) list = list.filter((d) => d.tags.includes(tag));
      if (st === "active") list = list.filter((d) => ACTIVE.has(d.status));
      else if (st !== "all") list = list.filter((d) => d.status === st);
      if (q) {
        const s = q.toLowerCase();
        list = list.filter((d) => d.title.toLowerCase().includes(s) || (d.title_orig || "").toLowerCase().includes(s)
          || d.artists.some((a) => a.toLowerCase().includes(s)));
      }
      const rank = { CRITICAL: 3, HIGH: 2, NORMAL: 1 };
      const time = (d) => t(d.origin.published_at);
      list.sort((a, b) => (sort === "priority" ? rank[b.priority] - rank[a.priority] || time(b) - time(a) : time(b) - time(a)));
      return { stories: list.map((d) => clone(card(d))), server_time: new Date().toISOString() };
    },
    async story(id) {
      await first;
      const d = details.get(String(id));
      if (!d) throw new ApiError(404, "Сюжет не найден");
      return clone(fresh(d));
    },
    async collectFacts(id) {
      const d = details.get(String(id));
      if (!d) throw new ApiError(404, "Сюжет не найден");
      const sample = await getSample();
      if (!sample) throw new ApiError(503, "Факты собирает Claude из твоего аккаунта — это работает, когда страница открыта в Claude.");
      const user = buildFactsPrompt(docsFor(d), new Date(), P.audience_tz, d.scene, P.facts);
      let answer;
      try {
        answer = await sample.json(sampleInput({ ...P, schema: P.facts_schema }, user), { modelTier: "default", cache: false });
      } catch (e) {
        throw new ApiError(502, sampleErrorText(e), e?.code);
      }
      let items;
      try {
        items = normalizeFacts(answer);
      } catch {
        throw new ApiError(502, "Claude не вернул список фактов — нажми ещё раз.");
      }
      const now = new Date().toISOString();
      const facts = { items, created_at: now, model: "Claude из твоего аккаунта" };
      await write(d.id, { facts });
      const nd = toDetail(d.id, { ...docs.get(d.id), facts });
      details.set(d.id, nd);
      return clone(nd.facts);
    },
    async generate(id, mode = "standard", kind = "post", opts = {}) {
      const d = details.get(String(id));
      if (!d) throw new ApiError(404, "Сюжет не найден");
      const sample = await getSample();
      if (!sample) throw new ApiError(503, "Текст пишет Claude из твоего аккаунта — это работает, когда страница открыта в Claude.");
      fresh(d);
      const ds = docsFor(d);
      const article = kind === "article";
      const prev = article ? d.current_article : d.current_post;
      const previous = !article && (mode === "harder" || mode === "hardest") && prev
        ? { headline: prev.headline, bullets: prev.bullets, summary: prev.summary }
        : null;
      const picked = article && Array.isArray(opts.facts) ? opts.facts.filter((f) => f.on !== false && f.text) : null;
      if (picked && !picked.length) throw new ApiError(400, "Отметь хотя бы один факт.");
      const user = picked
        ? buildArticleFromFactsPrompt(ds, picked, new Date(), P.audience_tz, d.scene, P)
        : article
          ? buildUserPrompt(ds, new Date(), P.audience_tz, d.scene, "article", null, { article: P.article })
          : buildUserPrompt(ds, new Date(), P.audience_tz, d.scene, mode, previous, P.modes);
      const prompt = article ? { ...P, schema: P.article_schema } : P;
      let answer;
      try {
        answer = await sample.json(sampleInput(prompt, user), { modelTier: "default", cache: false });
      } catch (e) {
        throw new ApiError(502, sampleErrorText(e), e?.code);
      }
      let gp;
      try {
        gp = article ? normalizeArticle(answer, d.scene, P) : normalizePost(answer, d.scene, P);
      } catch {
        throw new ApiError(502, "Claude ответил не по формату — нажми ещё раз.");
      }
      const now = new Date().toISOString();
      const body = {
        flag_emoji: gp.flag_emoji, topic_emoji: gp.topic_emoji, headline: gp.headline,
        [article ? "paragraphs" : "bullets"]: gp.bullets,
        summary: gp.summary, notes: gp.notes_for_editor || "", mode: article ? "standard" : mode,
        version: (prev?.version ?? 0) + 1, model: "Claude из твоего аккаунта", created_at: now, updated_at: now,
      };
      const tags = tagOrder.filter((x) => d.tags.includes(x) || gp.tags.includes(x));
      const field = article ? "article" : "post";
      if (picked) body.from_facts = picked.length;
      const patch = { [field]: body, tags, updated_at: now };
      // выбор галочек сохраняем вместе со статьёй — откроешь сюжет с телефона, отмеченное на месте
      if (picked) patch.facts = { ...(docs.get(d.id)?.facts || {}), items: opts.facts.map((f) => ({ text: f.text, sources: f.sources || [], on: f.on !== false })) };
      await write(d.id, patch);
      const nd = toDetail(d.id, { ...docs.get(d.id), ...patch });
      details.set(d.id, nd);
      fresh(nd);
      return clone(article ? nd.current_article : nd.current_post);
    },
    async patchPost(pid, fields) {
      const d = byPost(pid);
      if (!d) throw new ApiError(404, "Пост не найден");
      const article = d.current_article?.id === pid;
      const field = article ? "article" : "post";
      const lines = article ? "paragraphs" : "bullets";
      const old = docs.get(d.id)?.[field] || {};
      const now = new Date().toISOString();
      const body = {
        ...old,
        flag_emoji: fields.flag_emoji ?? old.flag_emoji,
        topic_emoji: fields.topic_emoji ?? old.topic_emoji,
        headline: fields.headline != null ? String(fields.headline).trim() : old.headline,
        [lines]: Array.isArray(fields.bullets) ? fields.bullets.map((b) => b.trim()).filter(Boolean).slice(0, 6) : old[lines],
        summary: fields.summary != null ? String(fields.summary).trim() : old.summary,
        mode: "manual",
        updated_at: now,
      };
      await write(d.id, { [field]: body, updated_at: now });
      const nd = toDetail(d.id, { ...docs.get(d.id), [field]: body, updated_at: now });
      details.set(d.id, nd);
      const p = article ? fresh(nd).current_article : fresh(nd).current_post;
      // как Newsroom.update_post: после правки заново проверяем числа
      const y = new Date().getUTCFullYear();
      const missing = checkNumbers([p.headline, ...p.bullets, p.summary], docsFor(nd).map((x) => `${x.title}\n${x.text}`),
                                   new Set([`${y}`, `${y - 1}`, `${y + 1}`]));
      if (missing.length && !p.warnings.some((w) => w.startsWith("Числа"))) {
        p.warnings.unshift("Числа не найдены в источниках: " + missing.slice(0, 8).join(", "));
      }
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
      const d = details.get(String(id));
      await write(d.id, { status: "dismissed" });
      d.status = "dismissed";
      return clone(card(d));
    },
    async restore(id) {
      const d = details.get(String(id));
      await write(d.id, { status: "new" });
      d.status = "new";
      return clone(card(d));
    },
    async collect() {
      if (runBusy()) throw new ApiError(409, "Сбор уже идёт — новости появятся сами.");
      const mcp = await use("mcp");
      if (!mcp) {
        throw new ApiError(503, "Запуск сбора работает, когда страница открыта в Claude (claude.ai или приложение).");
      }
      const prev = status || {};
      const now = new Date().toISOString();
      try {
        await db.doc("meta/status").set({ ...prev, state: "requested", requested_at: now, message: "Запускаю сбор…" });
      } catch {
        throw new ApiError(403, "Запускать сбор может только владелец страницы.");
      }
      try {
        // без `text`: дополнительный текст приходит в задачу отдельным сообщением ПОСЛЕ сбора,
        // и задача прогоняла сбор второй раз — пустой прогон затирал итог первого в meta/status
        await mcp.callTool(CONNECTOR, "fire_trigger", { trigger_id: TRIGGER_ID });
      } catch (e) {
        const msg = mcpErrorText(e);
        quietError = true;
        try {
          await db.doc("meta/status").set({ ...prev, state: "error", finished_at: new Date().toISOString(), message: msg });
        } catch {
          /* статус не записался — ошибку всё равно покажем */
        }
        throw new ApiError(502, msg, e?.code);
      }
      return { queued: true };
    },
    subscribe(onStory, onStats, onStatus, toast) {
      let timer = null;
      const l = {
        onStory,
        toast,
        statsSoon: () => {
          clearTimeout(timer);
          timer = setTimeout(async () => onStats(await api.stats()), 300);
        },
      };
      listeners.add(l);
      onStatus(true);
      // «Сбор идёт…» гаснет сам, если запуск не отозвался за полчаса
      const tick = setInterval(() => l.statsSoon(), 60000);
      return () => {
        listeners.delete(l);
        clearTimeout(timer);
        clearInterval(tick);
      };
    },
  };
  return api;
}

let samplePromise = null;
function getSample() {
  if (!samplePromise) samplePromise = use("sample");
  return samplePromise;
}
