// JS-копия редакционной цепочки бэкенда — для режима «Реальные новости», где пост пишет
// Claude прямо на странице. Источники правды (меняешь там — меняй здесь):
//   промпт            backend/sqandal/llm/prompts.py (сам текст приходит в снимке: data.prompt)
//   нормализация      backend/sqandal/llm/generator.py (GeneratedPost)
//   фактчек чисел     backend/sqandal/llm/factcheck.py
//   выбор источников  backend/sqandal/services.py (Newsroom._pick_docs, _source_links, _warnings)

const W = "[\\p{L}\\p{N}_]"; // \w Питона: буквы любого алфавита
const NUM = () => new RegExp(`(?<![\\p{L}\\p{N}_.,])\\d+(?:[ \\u00a0\\u202f.,]\\d{3})*(?:[.,]\\d+)?(?!${W})`, "gu");
const SCALE = () =>
  new RegExp(`(\\d+(?:[.,]\\d+)?)\\s?(k|m|тыс\\.?|млн|млрд|mio\\.?|mrd\\.?|million|millionen|billion)(?!${W})`, "giu");
const SCALE_MULT = {
  k: 1e3, "тыс": 1e3, "тыс.": 1e3, m: 1e6, "млн": 1e6, mio: 1e6, "mio.": 1e6, million: 1e6, millionen: 1e6,
  "млрд": 1e9, mrd: 1e9, "mrd.": 1e9, billion: 1e9,
};
const WORDS = {
  one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12,
  fifteen: 15, twenty: 20, thirty: 30, forty: 40, fifty: 50, hundred: 100, thousand: 1000, million: 1000000,
  eins: 1, ein: 1, eine: 1, zwei: 2, drei: 3, vier: 4, "fünf": 5, sechs: 6, sieben: 7, acht: 8, neun: 9, zehn: 10,
  elf: 11, "zwölf": 12, zwanzig: 20, "dreißig": 30, hundert: 100, tausend: 1000, millionen: 1000000,
  "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3, "четыре": 4, "пять": 5, "шесть": 6, "семь": 7, "восемь": 8,
  "девять": 9, "десять": 10, "двадцать": 20, "тридцать": 30, "сто": 100, "тысяча": 1000, "тысячи": 1000,
  "миллион": 1000000, "миллиона": 1000000,
};
const WORD_RE = Object.entries(WORDS).map(([w, v]) => [new RegExp(`(?<!${W})${w}(?!${W})`, "u"), String(v)]);

function norm(num) {
  const s = num.replace(/[  ]/g, " ");
  if (/^\d{1,3}([ .,]\d{3})+$/.test(s)) return s.replace(/[ .,]/g, "");
  return s.replace(/,/g, ".").replace(/ /g, "");
}

const fmt = (x) => String(x);

function numbersIn(text) {
  const out = new Set();
  for (const m of (text || "").matchAll(NUM())) out.add(norm(m[0]));
  for (const m of (text || "").matchAll(SCALE())) {
    out.add(norm(m[1]));
    const unit = m[2].toLowerCase();
    const mult = SCALE_MULT[unit.replace(/\.+$/, "")] ?? SCALE_MULT[unit] ?? 1;
    out.add(fmt(parseFloat(m[1].replace(",", ".")) * mult));
  }
  return out;
}

function sourceNumbers(text) {
  const nums = numbersIn(text);
  for (const m of (text || "").matchAll(/\d+/g)) nums.add(m[0]); // «18.10.2026» → 18, 10, 2026
  const low = (text || "").toLowerCase();
  for (const [rx, v] of WORD_RE) if (rx.test(low)) nums.add(v);
  return nums;
}

/** Числа из поста, которых нет в источниках (allow — заведомо безопасные, напр. текущий год). */
export function checkNumbers(parts, sourceTexts, allow = new Set()) {
  const src = new Set();
  for (const t of sourceTexts) for (const n of sourceNumbers(t)) src.add(n);
  const missing = [];
  for (const part of parts) {
    for (const m of (part || "").matchAll(NUM())) {
      const n = norm(m[0]);
      if (src.has(n) || allow.has(n) || src.has(n.replace(/^0+/, ""))) continue;
      if (!missing.includes(m[0])) missing.push(m[0]);
    }
  }
  return missing;
}

// ── источники сюжета → документы для модели ────────────────────────
export const DAY_ONLY = new Set(["url", "page_day"]);
const t = (iso) => new Date(iso).getTime();

export function truncate(text, limit) {
  const s = (text || "").trim();
  if (s.length <= limit) return s;
  let cut = s.slice(0, limit);
  const space = cut.lastIndexOf(" ");
  if (space > limit * 0.7) cut = cut.slice(0, space);
  return cut.replace(/[ ,.;:—-]+$/, "") + "…";
}

/** Первоисточник + по одному на остальные издания; СМИ раньше TG-репостов. */
export function pickDocs(items, max) {
  const ms = [...items].sort((a, b) => t(a.effective_at) - t(b.effective_at));
  if (!ms.length) return [];
  const picked = [ms[0]];
  const used = new Set([ms[0].source_name]);
  const tg = (m) => (m.source_kind === "telegram" ? 1 : 0);
  const rest = ms.slice(1).sort((a, b) => tg(a) - tg(b) || t(a.effective_at) - t(b.effective_at));
  for (const m of rest) {
    if (picked.length >= max) break;
    if (used.has(m.source_name)) continue;
    used.add(m.source_name);
    picked.push(m);
  }
  return picked;
}

export function sourceLinks(rows) {
  const ordered = [...rows.filter((m) => m.source_kind !== "telegram"), ...rows.filter((m) => m.source_kind === "telegram")];
  const out = [];
  const seen = new Set();
  for (const m of ordered) {
    if (seen.has(m.source_name)) continue;
    seen.add(m.source_name);
    out.push({ name: m.source_name, url: m.url });
  }
  return out.slice(0, 4);
}

// ── промпт ─────────────────────────────────────────────────────────
const xmlAttr = (v) => String(v).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
const xmlText = (v) => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function at(date, tz) {
  const f = new Intl.DateTimeFormat("en-GB", {
    timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23",
  });
  const p = Object.fromEntries(f.formatToParts(date).map((x) => [x.type, x.value]));
  return { Y: p.year, M: p.month, D: p.day, h: p.hour, m: p.minute };
}

function header(docs, now, tz, sceneHint) {
  const L = at(now, tz);
  const U = at(now, "UTC");
  return [
    `Текущее время: ${L.Y}-${L.M}-${L.D} ${L.h}:${L.m} (${tz}), ${U.Y}-${U.M}-${U.D}T${U.h}:${U.m}Z UTC.`,
    `Сцена по источникам: ${sceneHint}. Источников: ${docs.length}.`,
    "",
  ];
}

function published(d, tz) {
  const u = at(d.published, "UTC");
  const l = at(d.published, tz);
  return d.dayOnly
    ? `${u.Y}-${u.M}-${u.D} (известна только дата, время не указано)`
    : `${u.Y}-${u.M}-${u.D}T${u.h}:${u.m}Z UTC / ${l.D}.${l.M} ${l.h}:${l.m} ${tz}`;
}

function sourceAttrs(d, tz) {
  let attrs = `id="${d.idx}" outlet="${xmlAttr(d.outlet)}" type="${d.kind}" published="${published(d, tz)}" date_confidence="${d.dateConfidence}"`;
  if (d.url) attrs += ` url="${xmlAttr(d.url)}"`;
  if (d.fwdFrom) attrs += ` repost_of="${xmlAttr(d.fwdFrom)}"`;
  return attrs;
}

function sourceBlocks(docs, tz) {
  const lines = [];
  for (const d of docs) {
    lines.push(`<source ${sourceAttrs(d, tz)}>`, `<title>${xmlText(d.title)}</title>`, `<text>${xmlText(d.text)}</text>`, "</source>", "");
  }
  return lines;
}

export function buildUserPrompt(docs, now, tz, sceneHint, mode, previous, modes) {
  const lines = [...header(docs, now, tz, sceneHint), ...sourceBlocks(docs, tz)];
  if (previous) {
    lines.push("<previous_draft>", `headline: ${previous.headline ?? ""}`);
    for (const b of previous.bullets ?? []) lines.push(`- ${b}`);
    lines.push(`summary: ${previous.summary ?? ""}`, "</previous_draft>", "");
  }
  if (modes?.[mode]) lines.push(modes[mode]);
  lines.push("Собери пост по правилам редакции.");
  return lines.join("\n");
}

/** «Собрать факты»: все источники сюжета → отдельные пункты с изданиями. */
export function buildFactsPrompt(docs, now, tz, sceneHint, instruction) {
  return [...header(docs, now, tz, sceneHint), ...sourceBlocks(docs, tz), instruction].join("\n");
}

/** Статья только из отмеченных фактов: тексты источников не даём — только кто и когда опубликовал. */
export function buildArticleFromFactsPrompt(docs, facts, now, tz, sceneHint, prompt) {
  const lines = [...header(docs, now, tz, sceneHint)];
  for (const d of docs) lines.push(`<source ${sourceAttrs(d, tz)} />`);
  lines.push("", "<facts>");
  facts.forEach((f, i) => lines.push(`${i + 1}. ${xmlText(f.text)}${f.sources?.length ? ` (${f.sources.join(", ")})` : ""}`));
  lines.push("</facts>", "", prompt.article, prompt.article_from_facts, "Собери статью по правилам редакции.");
  return lines.join("\n");
}

/** Ответ «факты» → [{text, sources, on}] без пустых и повторов. */
export function normalizeFacts(o) {
  const raw = Array.isArray(o?.facts) ? o.facts : [];
  const seen = new Set();
  const out = [];
  for (const f of raw) {
    const text = ws(f?.text).replace(BULLET_PREFIX, "");
    if (!text || seen.has(text.toLowerCase())) continue;
    seen.add(text.toLowerCase());
    out.push({ text, sources: (Array.isArray(f?.sources) ? f.sources : []).map((x) => ws(x)).filter(Boolean), on: true });
  }
  if (!out.length) throw new Error("нет фактов");
  return out.slice(0, 25);
}

/** Один текст для `sample`: системного промпта у страницы нет — правила идут первыми. */
export function sampleInput(prompt, userPrompt) {
  return [
    prompt.system.trim(),
    "",
    "# Материалы",
    userPrompt,
    "",
    "# Ответ",
    "Верни только один JSON-объект — без пояснений и без Markdown — строго по этой JSON Schema:",
    JSON.stringify(prompt.schema),
  ].join("\n");
}

// ── ответ модели → пост ─────────────────────────────────────────────
const SCENE_FLAGS = { ru: "🇷🇺", de: "🇩🇪", us: "🇺🇸", intl: "🌍" };
const FLAG_RE = /^(?:[\u{1F1E6}-\u{1F1FF}]{2}|🏴[\u{E0061}-\u{E007A}]+\u{E007F}|🌍|🌎|🌏)$/u;
const BULLET_PREFIX = /^\s*(?:[-–—•▪●◦*▫︎️]+|\d+[.)])\s*/u;
const ws = (s) => String(s ?? "").replace(/\s+/g, " ").trim();
const WRAP = { "«": "»", "“": "”", "„": "“", '"': '"', "'": "'" };

function balanced(text) {
  let depth = 0;
  for (const ch of text) {
    if (ch === "«") depth += 1;
    else if (ch === "»" && --depth < 0) return false;
  }
  return depth === 0;
}

/** Снять кавычки, в которые обёрнут ВЕСЬ заголовок; кавычки названия в конце остаются. */
export function unwrapQuotes(v) {
  while (v.length >= 2 && WRAP[v[0]] === v[v.length - 1]) {
    const inner = v.slice(1, -1);
    if (!balanced(inner) || (`"'`.includes(v[0]) && inner.includes(v[0]))) break;
    v = inner.trim();
  }
  return v;
}

export function normalizePost(o, fallbackScene, prompt) {
  if (!o || typeof o !== "object" || Array.isArray(o)) throw new Error("ответ не объект");
  let headline = unwrapQuotes(ws(o.headline));
  if (headline.endsWith(".") && !headline.endsWith("..")) headline = headline.slice(0, -1);
  const seen = new Set();
  const bullets = [];
  for (const raw of Array.isArray(o.bullets) ? o.bullets : []) {
    const b = ws(raw).replace(BULLET_PREFIX, "").trim();
    if (b && !seen.has(b.toLowerCase())) {
      seen.add(b.toLowerCase());
      bullets.push(b);
    }
  }
  if (!headline || !bullets.length) throw new Error("нет заголовка или фактов");
  const scene = ["ru", "de", "us", "intl"].includes(o.scene) ? o.scene : fallbackScene;
  const flag = ws(o.flag_emoji);
  const tags = Array.isArray(o.tags) ? o.tags : [];
  return {
    is_publishable: o.is_publishable !== false,
    scene,
    flag_emoji: FLAG_RE.test(flag) ? flag : SCENE_FLAGS[scene] ?? "🌍",
    topic_emoji: prompt.topic_emoji.includes(o.topic_emoji) ? o.topic_emoji : "",
    headline,
    bullets: bullets.slice(0, 4),
    summary: ws(o.summary),
    tags: prompt.tags.filter((x) => tags.includes(x)),
    event_date: typeof o.event_date === "string" && /^\d{4}-\d{2}-\d{2}$/.test(o.event_date) ? o.event_date : null,
    notes_for_editor: ws(o.notes_for_editor),
  };
}

/** Ответ модели в формате «статья» → поля поста (абзацы лежат в bullets, kind = "article"). */
export function normalizeArticle(o, fallbackScene, prompt) {
  if (!o || typeof o !== "object" || Array.isArray(o)) throw new Error("ответ не объект");
  const paras = (Array.isArray(o.paragraphs) ? o.paragraphs : []).map((x) => ws(x).replace(BULLET_PREFIX, "")).filter(Boolean);
  const post = normalizePost({ ...o, bullets: ["абзацы ниже"] }, fallbackScene, prompt);
  if (!paras.length) throw new Error("нет абзацев");
  return { ...post, bullets: paras.slice(0, 6) };
}

export function postWarnings(story, parts, docs, length) {
  const w = [];
  const y = new Date().getUTCFullYear();
  const missing = checkNumbers(parts, docs.map((d) => `${d.title}\n${d.text}`), new Set([`${y}`, `${y - 1}`, `${y + 1}`]));
  if (missing.length) w.push("Числа не найдены в источниках: " + missing.slice(0, 8).join(", "));
  if (story.verification === "insider_only") w.push("Только Telegram-инсайды — подтверждения от СМИ нет");
  else if (story.verification === "single_source") w.push("Один источник");
  if (story.is_stale) w.push(`Устаревание: ${story.stale_reason}`);
  if (story.republish_of) w.push(`Похоже на уже опубликованный сюжет #${story.republish_of}`);
  if (story.origin?.confidence === "low") w.push("Точная дата первоисточника не подтверждена");
  if (length > 4096) w.push("Пост длиннее 4096 символов — Telegram не примет");
  return w;
}

// Коды ошибок `sample` → что сказать редактору
export function sampleErrorText(e) {
  switch (e?.code) {
    case "not_granted":
      return "Claude не разрешено писать посты на этой странице. Обнови страницу и разреши, когда Claude спросит.";
    case "sampling_disabled":
    case "not_declared":
    case "capability_disabled":
    case "capability_removed":
      return "Генерация через Claude недоступна в этом окне.";
    case "rate_limited":
      return "Слишком много запросов или закончился лимит Claude — попробуй чуть позже.";
    case "session_expired":
      return "Сессия Claude истекла — войди заново и повтори.";
    case "refused":
      return "Claude отказался писать этот пост — проверь источник вручную.";
    case "invalid_json":
    case "empty_completion":
      return "Claude ответил не по формату — нажми ещё раз.";
    case "prompt_too_large":
      return "Источники слишком длинные для одного запроса.";
    case "cancelled":
      return "Генерация остановлена.";
    default:
      return "Claude сейчас не отвечает — попробуй ещё раз.";
  }
}

// ── приоритет: как Tagger.score (backend/sqandal/processing/tagging.py) ──
// Для сюжетов, которые Claude собирает по кнопке: теги и источники → CRITICAL / HIGH / NORMAL.
export function scoreStory({ tags, artists, items }, weights, scoring, staleAfterHours) {
  const b = scoring.bonuses;
  const w = tags.map((x) => weights[x] ?? 0).sort((x, y) => y - x);
  let score = 0;
  if (w.length) score += w[0] + (b.second_tag_factor ?? 0.3) * w.slice(1).reduce((x, y) => x + y, 0);
  const outlets = new Map();
  for (const it of items) outlets.set(it.source_name, Math.min(outlets.get(it.source_name) ?? Infinity, t(it.effective_at)));
  score += Math.min(Math.max(0, outlets.size - 1) * (b.extra_source ?? 10), b.extra_source_cap ?? 30);
  if (outlets.size) {
    const first = Math.min(...outlets.values());
    const within = [...outlets.values()].filter((x) => x - first <= (b.breaking_window_h ?? 2) * 3600e3).length;
    if (within >= (b.breaking_sources ?? 3)) score += b.breaking ?? 15;
  }
  const names = new Map();
  for (const a of scoring.artists) for (const n of [a.name, ...(a.aliases ?? [])]) names.set(n.toLowerCase(), a.tier);
  const tiers = artists.map((a) => names.get(String(a).toLowerCase())).filter(Boolean);
  if (tiers.length) score += Math.min(...tiers) === 1 ? b.artist_tier1 ?? 10 : b.artist_tier2 ?? 3;
  if (items.length && items.every((it) => it.source_kind === "telegram")) score += b.insider_only ?? -5;
  // факт пересказали, когда он уже был старым: последний источник вышел позже порога «свежести»
  if (items.length > 1) {
    const ts = items.map((it) => t(it.effective_at));
    if (Math.max(...ts) - Math.min(...ts) > staleAfterHours * 3600e3) score += b.stale ?? -40;
  }
  const s = Math.round(score);
  const th = scoring.thresholds;
  return [s >= th.CRITICAL ? "CRITICAL" : s >= th.HIGH ? "HIGH" : "NORMAL", s];
}
