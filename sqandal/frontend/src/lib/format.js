// Время, склонения и рендер поста (JS-копия backend/sqandal/publish/formatter.py —
// живой предпросмотр правок до сохранения).

const MIN = 60 * 1000;
const HOUR = 60 * MIN;

export function plural(n, one, few, many) {
  const a = Math.abs(n) % 100;
  const b = a % 10;
  if (a > 10 && a < 20) return many;
  if (b > 1 && b < 5) return few;
  if (b === 1) return one;
  return many;
}

export function ago(iso, now = Date.now()) {
  if (!iso) return "—";
  const diff = now - new Date(iso).getTime();
  if (diff < 0) return "через " + agoShort(-diff);
  return agoShort(diff) + " назад";
}

function agoShort(ms) {
  if (ms < MIN) return "меньше минуты";
  if (ms < HOUR) {
    const m = Math.round(ms / MIN);
    return `${m} мин`;
  }
  if (ms < 48 * HOUR) {
    const h = Math.floor(ms / HOUR);
    const m = Math.round((ms - h * HOUR) / MIN);
    return m && h < 6 ? `${h} ч ${m} мин` : `${h} ч`;
  }
  const d = Math.floor(ms / (24 * HOUR));
  return `${d} ${plural(d, "день", "дня", "дней")}`;
}

export function compactAgo(iso, now = Date.now()) {
  const diff = Math.max(0, now - new Date(iso).getTime());
  if (diff < HOUR) return `${Math.max(1, Math.round(diff / MIN))} мин`;
  if (diff < 48 * HOUR) return `${Math.floor(diff / HOUR)} ч`;
  return `${Math.floor(diff / (24 * HOUR))} д`;
}

const two = (n) => String(n).padStart(2, "0");

export function clock(iso) {
  const d = new Date(iso);
  return `${two(d.getHours())}:${two(d.getMinutes())}`;
}

export function stamp(iso, now = Date.now()) {
  if (!iso) return "—";
  const d = new Date(iso);
  const today = new Date(now);
  const sameDay = d.toDateString() === today.toDateString();
  const yesterday = new Date(now - 24 * HOUR).toDateString() === d.toDateString();
  const day = sameDay ? "сегодня" : yesterday ? "вчера" : `${two(d.getDate())}.${two(d.getMonth() + 1)}`;
  return `${day}, ${clock(iso)}`;
}

// известен только день выхода (дата из адреса или шапки статьи): время не показываем
export const DAY_ONLY = new Set(["url", "page_day"]);

export function day(iso, now = Date.now()) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (d.toDateString() === new Date(now).toDateString()) return "сегодня";
  if (d.toDateString() === new Date(now - 24 * HOUR).toDateString()) return "вчера";
  return `${two(d.getDate())}.${two(d.getMonth() + 1)}`;
}

export function ageHours(iso, now = Date.now()) {
  return (now - new Date(iso).getTime()) / HOUR;
}

export function toLocalInput(date) {
  return `${date.getFullYear()}-${two(date.getMonth() + 1)}-${two(date.getDate())}T${two(date.getHours())}:${two(
    date.getMinutes(),
  )}`;
}

// code — короткая метка сцены в интерфейсе (вместо флага-эмодзи); флаги остаются только в тексте постов
export const SCENES = {
  ru: { flag: "🇷🇺", code: "RU", label: "СНГ" },
  de: { flag: "🇩🇪", code: "DE", label: "GER" },
  us: { flag: "🇺🇸", code: "US", label: "US" },
  intl: { flag: "🌍", code: "INT", label: "INTL" },
};

export const DATE_SOURCE = {
  feed_published: "дата из RSS",
  feed_updated: "дата обновления RSS",
  jsonld: "дата из разметки статьи",
  meta: "дата из meta-тегов",
  time_tag: "дата со страницы",
  url: "дата из адреса — только день",
  page_day: "дата со страницы — только день",
  telegram: "время поста в Telegram",
  telegram_forward: "время оригинала репоста",
  fetched: "даты нет — время обнаружения",
};

export const CONFIDENCE = { high: "точно", medium: "примерно", low: "не подтверждена" };

export const VERIFICATION = {
  single_source: "1 источник",
  multi_source: "подтверждено",
  insider_only: "только инсайд",
};

export const STATUS = {
  new: "без поста",
  drafted: "черновик",
  scheduled: "запланирован",
  published: "опубликован",
  dismissed: "скрыт",
};

export const MODE = { standard: "обычный", harder: "жёстче", hardest: "максимально жёстко", manual: "правка" };

export const NEXT_MODE = { standard: "harder", harder: "hardest", hardest: "hardest", manual: "harder" };

export const TOPIC_EMOJI = ["🚨", "⚖️", "💿", "🔥", "🏥", "🕊", "🎤", "💰", "📈", ""];
export const FLAG_EMOJI = ["🇷🇺", "🇺🇦", "🇧🇾", "🇰🇿", "🇩🇪", "🇦🇹", "🇨🇭", "🇺🇸", "🇬🇧", "🇫🇷", "🌍"];

// Демо-ссылки (demo.local, вымышленные TG-каналы) — не кликабельны, ведут в никуда
export function safeHref(url) {
  if (!url) return null;
  try {
    const u = new URL(url);
    if (u.hostname === "demo.local" || (u.hostname === "t.me" && u.pathname.startsWith("/demo_"))) return null;
    return u.protocol === "https:" || u.protocol === "http:" ? url : null;
  } catch {
    return null;
  }
}

// ── пост ────────────────────────────────────────────────────────────
const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escAttr = (s) => esc(s).replace(/"/g, "&quot;");
const clean = (s) => String(s ?? "").replace(/\s+/g, " ").trim();

// пункты поста или абзацы статьи (kind === "article": абзацы через пустую строку)
function body(p, bullet, e) {
  const items = p.bullets.filter((x) => clean(x));
  if (p.kind !== "article") return items.map((b) => `${bullet} ${e(clean(b))}`);
  return items.flatMap((x, i) => (i ? ["", e(clean(x))] : [e(clean(x))]));
}

export function renderPostHtml(p, cfg) {
  const out = [];
  const head = (p.flag_emoji || "") + (p.topic_emoji || "");
  const title = `<b>${esc(clean(p.headline))}</b>`;
  out.push(head ? `${head} ${title}` : title, "");
  out.push(...body(p, cfg.bullet, esc));
  if (clean(p.summary)) {
    const s = esc(clean(p.summary));
    out.push(
      "",
      { quote: `<blockquote>${s}</blockquote>`, bold: `<b>${s}</b>`, italic: `<i>${s}</i>`, plain: s }[
        cfg.summary_style
      ] ?? s,
    );
  }
  if (p.footnotes?.length) out.push("", ...p.footnotes.map((f) => `<i>* ${esc(clean(f))}</i>`));
  out.push("");
  if (cfg.show_sources && p.sources?.length) {
    const links = p.sources
      .slice(0, 4)
      .map((s) => (s.url ? `<a href="${escAttr(s.url)}">${esc(s.name)}</a>` : esc(s.name)));
    out.push("📎 " + links.join(" · "));
  }
  out.push(cfg.brand_url ? `<a href="${escAttr(cfg.brand_url)}">${esc(cfg.brand_line)}</a>` : `<b>${esc(cfg.brand_line)}</b>`);
  return out.join("\n");
}

export function renderPostPlain(p, cfg) {
  const out = [`${p.flag_emoji || ""}${p.topic_emoji || ""} ${clean(p.headline)}`.trim(), ""];
  out.push(...body(p, cfg.bullet, (x) => x));
  if (clean(p.summary)) out.push("", clean(p.summary));
  if (p.footnotes?.length) out.push("", ...p.footnotes.map((f) => `* ${clean(f)}`));
  out.push("");
  if (cfg.show_sources && p.sources?.length) out.push("📎 " + p.sources.slice(0, 4).map((s) => s.name).join(" · "));
  out.push(cfg.brand_line);
  return out.join("\n");
}

export function telegramLength(p, cfg) {
  // как считает Telegram: видимый текст в UTF-16
  return renderPostPlain(p, cfg).length;
}

export async function copyPost(p, cfg) {
  const html = renderPostHtml(p, cfg).replace(/\n/g, "<br>");
  const plain = renderPostPlain(p, cfg);
  try {
    if (window.ClipboardItem && navigator.clipboard?.write) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([plain], { type: "text/plain" }),
        }),
      ]);
      return true;
    }
    await navigator.clipboard.writeText(plain);
    return true;
  } catch {
    try {
      await navigator.clipboard.writeText(plain);
      return true;
    } catch {
      /* ниже — последний способ */
    }
    const ta = document.createElement("textarea");
    ta.value = plain;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch {
      ok = false;
    }
    ta.remove();
    return ok;
  }
}
