import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  ago,
  ageHours,
  safeHref,
  clock,
  CONFIDENCE,
  copyPost,
  DATE_SOURCE,
  DAY_ONLY,
  day,
  FLAG_EMOJI,
  MODE,
  NEXT_MODE,
  SCENES,
  stamp,
  telegramLength,
  toLocalInput,
  TOPIC_EMOJI,
  VERIFICATION,
} from "../lib/format.js";
import {
  IconBack,
  IconCheck,
  IconChevron,
  IconClock,
  IconCopy,
  IconEye,
  IconFlame,
  IconPlus,
  IconSend,
  IconSpeaker,
  IconWarn,
  IconX,
} from "./Icons.jsx";

const EDITABLE = new Set(["draft", "scheduled", "failed"]);

// Шкала свежести: линейка от 0 до порога «старого факта» (штрих — час, крупный — четверть),
// стрелка — возраст первоисточника. За порогом стрелка упирается в конец и краснеет.
function AgeRuler({ age, limit }) {
  const hours = Math.max(1, Math.round(limit));
  const ratio = Math.min(1, Math.max(0, age / hours));
  const over = age >= hours;
  const quarters = [0, 0.25, 0.5, 0.75, 1];
  return (
    <div className={`ruler ${over ? "over" : ""}`} title={`Возраст новости: ${Math.round(age)} ч, порог «старого факта» — ${hours} ч`}>
      <div className="ruler-scale" style={{ "--n": hours }}>
        <i className="ruler-fill" style={{ width: `${ratio * 100}%` }} />
        <i className="ruler-needle" style={{ left: `${ratio * 100}%` }} />
      </div>
      <div className="ruler-labels" aria-hidden="true">
        {quarters.map((q) => (
          <span key={q} style={{ left: `${q * 100}%` }}>{Math.round(q * hours)}{q === 1 ? " ч" : ""}</span>
        ))}
      </div>
    </div>
  );
}

function Origin({ story, staleAfter, now }) {
  const o = story.origin;
  const age = ageHours(o.published_at, now);
  const dayOnly = DAY_ONLY.has(o.date_source);
  return (
    <section className={`origin ${story.is_stale ? "stale" : ""}`} aria-label="Первоисточник">
      <div className="origin-read">
        <span className="eyebrow">Первоисточник</span>
        <div className="origin-time">{dayOnly ? day(o.published_at, now) : clock(o.published_at)}</div>
      </div>
      <div className="origin-info">
        <div className="origin-what">
          <b>{o.source_name ?? "—"}</b>, {dayOnly ? `${day(o.published_at, now)}, время не указано` : stamp(o.published_at, now)}
        </div>
        <div className="origin-how">
          {dayOnly ? "Известен только день выхода" : ago(o.published_at, now)}, {DATE_SOURCE[o.date_source] ?? "дата"} ({CONFIDENCE[o.confidence]})
          {story.is_stale && <>. Старше {staleAfter} ч — перед отправкой будет вопрос</>}
        </div>
      </div>
      <AgeRuler age={age} limit={staleAfter} />
    </section>
  );
}

function Timeline({ items, now }) {
  return (
    <section>
      <h3 className="section-title">Источники по времени выхода</h3>
      <ol className="timeline">
        {items.map((it, i) => (
          <li key={it.id}>
            {DAY_ONLY.has(it.date_source) ? (
              <div className="tl-time">
                {day(it.effective_at, now)}
                <small>без времени</small>
              </div>
            ) : (
              <div className="tl-time">
                {clock(it.effective_at)}
                <small>{stamp(it.effective_at, now).split(",")[0]}</small>
              </div>
            )}
            <div>
              <div className="tl-src">
                <b>{it.source_name}</b>
                <span>{it.source_kind === "telegram" ? "Telegram" : "СМИ"}</span>
                {i === 0 && <span className="first">первым</span>}
                <span>{DATE_SOURCE[it.date_source]}</span>
                {it.fwd_from && <span>репост: {it.fwd_from}</span>}
                {it.tg_views ? <span>{it.tg_views.toLocaleString("ru-RU")} просмотров</span> : null}
              </div>
              {safeHref(it.url) ? (
                <a className="tl-title" href={safeHref(it.url)} target="_blank" rel="noreferrer">
                  {it.title}
                </a>
              ) : (
                <span className="tl-title">{it.title}</span>
              )}
            </div>
          </li>
        ))}
      </ol>
    </section>
  );
}

function TelegramPreview({ post, cfg }) {
  const bullets = post.bullets.filter((b) => b.trim());
  const len = telegramLength(post, cfg);
  const summary = post.summary.trim();
  const style = cfg.summary_style;
  return (
    <section className="tg" aria-label="Предпросмотр поста в Telegram">
      <div className="tg-chan">
        <span className="tg-ava" aria-hidden="true"><IconSpeaker size={17} /></span>
        <span>
          <b>SQANDAL</b>
          <small>{cfg.channel}</small>
        </span>
        <span className="eyebrow tg-label">Как в канале</span>
      </div>
      <div className="tg-bubble">
        {(post.flag_emoji || post.topic_emoji) && <>{post.flag_emoji}{post.topic_emoji} </>}
        <b>{post.headline.trim() || "Заголовок"}</b>
        {"\n\n"}
        {post.kind === "article"
          ? bullets.map((b, i) => (
              <span key={i}>
                {i > 0 && "\n"}
                {b.trim()}
                {"\n"}
              </span>
            ))
          : bullets.map((b, i) => (
              <span key={i}>
                {cfg.bullet} {b.trim()}
                {"\n"}
              </span>
            ))}
        {summary &&
          (style === "quote" ? (
            <>
              {"\n"}
              <span className="tg-quote">{summary}</span>
            </>
          ) : (
            <>
              {"\n"}
              {style === "bold" ? <b>{summary}</b> : style === "italic" ? <i>{summary}</i> : summary}
              {"\n"}
            </>
          ))}
        {post.footnotes?.length ? (
          <>
            {"\n"}
            {post.footnotes.map((f, i) => (
              <i key={i}>* {f}{"\n"}</i>
            ))}
          </>
        ) : null}
        {"\n"}
        {cfg.show_sources && post.sources?.length ? (
          <>
            📎{" "}
            {post.sources.slice(0, 4).map((s, i) => (
              <span key={s.name}>
                {i > 0 && " · "}
                {safeHref(s.url) ? (
                  <a href={safeHref(s.url)} target="_blank" rel="noreferrer">
                    {s.name}
                  </a>
                ) : (
                  s.name
                )}
              </span>
            ))}
            {"\n"}
          </>
        ) : null}
        {cfg.brand_url ? (
          <a href={cfg.brand_url} target="_blank" rel="noreferrer">
            {cfg.brand_line}
          </a>
        ) : (
          <b>{cfg.brand_line}</b>
        )}
        <div className="tg-foot">
          <span>
            <IconEye /> —
          </span>
          <span>{clock(new Date().toISOString())}</span>
        </div>
      </div>
      <div className={`tg-count ${len > 4096 ? "over" : ""}`}>
        {len} из 4096 символов{len > 4096 ? " — Telegram не примет, сократи" : ""}
      </div>
    </section>
  );
}

function PostFields({ draft, setDraft, readOnly, kind = "post" }) {
  const article = kind === "article";
  const setBullet = (i, v) => setDraft((d) => ({ ...d, bullets: d.bullets.map((b, j) => (j === i ? v : b)) }));
  return (
    <section className="fields" aria-label="Текст поста">
      <div className="field">
        <span className="lbl">Флаг и тема</span>
        <div className="emoji-row">
          {FLAG_EMOJI.map((f) => (
            <button key={f} type="button" className="emoji" aria-pressed={draft.flag_emoji === f}
              disabled={readOnly} onClick={() => setDraft((d) => ({ ...d, flag_emoji: f }))}>
              {f}
            </button>
          ))}
        </div>
        <div className="emoji-row">
          {TOPIC_EMOJI.map((t) => (
            <button key={t || "none"} type="button" className="emoji" aria-pressed={draft.topic_emoji === t}
              disabled={readOnly} onClick={() => setDraft((d) => ({ ...d, topic_emoji: t }))}
              aria-label={t ? `Тема ${t}` : "Без эмодзи темы"}>
              {t || <IconX />}
            </button>
          ))}
        </div>
      </div>
      <div className="field">
        <label htmlFor="pf-head">Заголовок</label>
        <input id="pf-head" className="input" value={draft.headline} readOnly={readOnly}
          onChange={(e) => setDraft((d) => ({ ...d, headline: e.target.value }))} />
      </div>
      <div className="field">
        <span className="lbl">{article ? "Абзацы (первый — лид)" : "Факты (3–4 пункта)"}</span>
        {draft.bullets.map((b, i) => (
          <div className="bullet-row" key={i}>
            <textarea className="textarea" rows={article ? 5 : 2} value={b} readOnly={readOnly}
              aria-label={article ? `Абзац ${i + 1}` : `Пункт ${i + 1}`}
              onChange={(e) => setBullet(i, e.target.value)} />
            {!readOnly && (
              <button type="button" className="btn small ghost" aria-label={`Удалить пункт ${i + 1}`}
                onClick={() => setDraft((d) => ({ ...d, bullets: d.bullets.filter((_, j) => j !== i) }))}>
                <IconX />
              </button>
            )}
          </div>
        ))}
        {!readOnly && draft.bullets.length < (article ? 6 : 4) && (
          <button type="button" className="btn small ghost" style={{ justifySelf: "start" }}
            onClick={() => setDraft((d) => ({ ...d, bullets: [...d.bullets, ""] }))}>
            <IconPlus /> {article ? "Добавить абзац" : "Добавить пункт"}
          </button>
        )}
      </div>
      <div className="field">
        <label htmlFor="pf-sum">Итог одной фразой</label>
        <textarea id="pf-sum" className="textarea" rows={2} value={draft.summary} readOnly={readOnly}
          onChange={(e) => setDraft((d) => ({ ...d, summary: e.target.value }))} />
      </div>
    </section>
  );
}

function FactPicker({ facts, setFacts, busy, hasArticle, onCollect, onWrite }) {
  const [open, setOpen] = useState(!hasArticle);
  useEffect(() => setOpen(!hasArticle), [hasArticle]);
  if (!facts) {
    return (
      <section className="facts">
        <p className="facts-intro">
          Claude соберёт из всех источников сюжета отдельные факты — у каждого видно издание. Отметишь нужные, и статья
          напишется только из них.
        </p>
        <div className="facts-tools">
          <button className="btn primary" onClick={onCollect} disabled={!!busy}>
            {busy === "facts" ? <span className="spin" /> : <IconPlus />}
            {busy === "facts" ? "Claude собирает факты…" : "Собрать факты"}
          </button>
          {!hasArticle && (
            <button className="btn ghost" onClick={() => onWrite(null)} disabled={!!busy}>
              {busy === "generate" ? <span className="spin" /> : null}
              {busy === "generate" ? "Claude пишет статью…" : "Сразу написать статью"}
            </button>
          )}
        </div>
      </section>
    );
  }
  const on = facts.filter((f) => f.on).length;
  const toggle = (i) => setFacts((fs) => fs.map((x, j) => (j === i ? { ...x, on: !x.on } : x)));
  return (
    <section className="facts">
      <button type="button" className="facts-head" onClick={() => setOpen((o) => !o)} aria-expanded={open}>
        <b>Факты из всех источников</b>
        <span>отмечено {on} из {facts.length}</span>
        <span className="chev" aria-hidden="true"><IconChevron /></span>
      </button>
      {open && (
        <>
          <ul className="fact-list">
            {facts.map((f, i) => (
              <li key={i}>
                <label className={`fact ${f.on ? "" : "off"}`}>
                  <input type="checkbox" checked={f.on} onChange={() => toggle(i)} />
                  <span className="fact-text">{f.text}</span>
                  {f.sources.length > 0 && <span className="fact-src">{f.sources.join(" · ")}</span>}
                </label>
              </li>
            ))}
          </ul>
          <div className="facts-tools">
            <button className="btn primary" onClick={() => onWrite(facts)} disabled={!!busy || on === 0}>
              {busy === "generate" ? <span className="spin" /> : <IconFlame />}
              {busy === "generate" ? "Claude пишет статью…" : `${hasArticle ? "Переписать" : "Написать"} статью из отмеченных (${on})`}
            </button>
            <button className="btn small ghost" onClick={() => setFacts((fs) => fs.map((f) => ({ ...f, on: true })))}>Все</button>
            <button className="btn small ghost" onClick={() => setFacts((fs) => fs.map((f) => ({ ...f, on: false })))}>Снять все</button>
            <button className="btn small ghost" onClick={onCollect} disabled={!!busy}>
              {busy === "facts" ? <span className="spin" /> : null} Собрать заново
            </button>
          </div>
        </>
      )}
    </section>
  );
}

function StaleDialog({ reason, onConfirm, onCancel, busy }) {
  const ref = useRef(null);
  useEffect(() => {
    ref.current?.focus();
    const onKey = (e) => e.key === "Escape" && onCancel();
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onCancel]);
  return (
    <div className="overlay" role="dialog" aria-modal="true" aria-labelledby="stale-title">
      <div className="dialog">
        <h2 id="stale-title">Это может быть старый факт</h2>
        <p>{reason}</p>
        <p>Проверь дату у первоисточника. Если новость актуальна — отправляй.</p>
        <div className="row-btns">
          <button className="btn ghost" onClick={onCancel} ref={ref}>Не отправлять</button>
          <button className="btn hot" onClick={onConfirm} disabled={busy}>
            {busy ? <span className="spin" /> : <IconSend />} Всё равно отправить
          </button>
        </div>
      </div>
    </div>
  );
}

const fieldsOf = (p) => ({
  flag_emoji: p.flag_emoji,
  topic_emoji: p.topic_emoji,
  headline: p.headline,
  bullets: [...p.bullets],
  summary: p.summary,
});

export function Editor({ api, cfg, storyId, cardVersion, open, onClose, toast, onChanged, now }) {
  const [detail, setDetail] = useState(null);
  const [error, setError] = useState(null);
  const [draft, setDraft] = useState(null);
  const [busy, setBusy] = useState(null);
  const [schedOpen, setSchedOpen] = useState(false);
  const [schedAt, setSchedAt] = useState("");
  const [stale, setStale] = useState(null);
  const [tab, setTab] = useState("post");
  const [facts, setFacts] = useState(null);
  useEffect(() => {
    setFacts(detail?.facts ? detail.facts.map((f) => ({ ...f })) : null);
  }, [detail]);
  const tabRef = useRef(tab);
  tabRef.current = tab;
  const pick = (d, t = tabRef.current) => (t === "article" ? d?.current_article : d?.current_post) ?? null;
  const post = pick(detail);
  const dirty = useMemo(
    () => post && draft && JSON.stringify(fieldsOf(post)) !== JSON.stringify(draft),
    [post, draft],
  );

  const load = useCallback(
    async (keepDraft = false) => {
      try {
        const d = await api.story(storyId);
        setDetail(d);
        setError(null);
        if (!keepDraft) setDraft(pick(d) ? fieldsOf(pick(d)) : null);
      } catch (e) {
        setError(e.message);
      }
    },
    [api, storyId],
  );

  useEffect(() => {
    if (!storyId) return;
    setDetail(null);
    setDraft(null);
    setSchedOpen(false);
    load();
  }, [storyId, load]);

  // сюжет обновился в эфире (новый источник, публикация с другого устройства)
  const lastVersion = useRef(cardVersion);
  useEffect(() => {
    if (!storyId || cardVersion === lastVersion.current) return;
    lastVersion.current = cardVersion;
    if (!dirty && !busy) load();
  }, [cardVersion, storyId, dirty, busy, load]);

  if (!storyId) {
    return (
      <section className={`editor ${open ? "open" : ""}`}>
        <div className="editor-empty">
          <h2>Выбери сюжет слева</h2>
          <p>Здесь будет готовый пост для канала, время первоисточника и все источники. CRITICAL-сюжеты в ленте
            отмечены красной меткой слева.</p>
        </div>
      </section>
    );
  }
  if (error) {
    return (
      <section className={`editor ${open ? "open" : ""}`}>
        <div className="editor-empty">
          <h2>Сюжет не открылся</h2>
          <p>{error}</p>
          <button className="btn" onClick={() => load()}>Повторить</button>
        </div>
      </section>
    );
  }
  if (!detail) {
    return (
      <section className={`editor ${open ? "open" : ""}`}>
        <div className="editor-empty"><p>Загружаю сюжет…</p></div>
      </section>
    );
  }

  const readOnly = !post || !EDITABLE.has(post.status);
  const viaSample = cfg.llm_via === "sample";
  const isArticle = tab === "article";
  const noun = isArticle ? "Статья" : "Пост";
  const shown = draft && post ? { ...post, ...draft } : post;
  const nextMode = post ? NEXT_MODE[post.mode] : "standard";

  async function run(kind, fn) {
    setBusy(kind);
    try {
      return await fn();
    } catch (e) {
      if (e.code === "stale_confirmation_required") {
        setStale(e.message);
        return null;
      }
      toast({ kind: "err", text: e.message });
      return null;
    } finally {
      setBusy(null);
    }
  }

  async function saveIfDirty() {
    if (!dirty) return post;
    const p = await api.patchPost(post.id, draft);
    setDetail((d) => ({ ...d, [isArticle ? "current_article" : "current_post"]: p }));
    setDraft(fieldsOf(p));
    return p;
  }

  const generate = (mode) =>
    run(mode === "standard" ? "generate" : "harder", async () => {
      if (dirty) await saveIfDirty();
      await (isArticle
        ? api.generate(detail.id, "standard", "article", facts ? { facts } : {})
        : api.generate(detail.id, mode));
      await load();
      onChanged();
      toast({ kind: "ok", text: isArticle ? "Статья готова" : mode === "standard" ? "Пост готов" : `Новый вариант: ${MODE[mode]}` });
    });

  const collectFacts = () =>
    run("facts", async () => {
      const fs = await api.collectFacts(detail.id);
      await load(true);
      setFacts(fs);
      toast({ kind: "ok", text: `Фактов: ${fs.length}. Сними галочки с лишних и пиши статью` });
    });

  const writeArticle = (fs) =>
    run("generate", async () => {
      if (dirty) await saveIfDirty();
      await api.generate(detail.id, "standard", "article", fs ? { facts: fs } : {});
      await load();
      onChanged();
      toast({ kind: "ok", text: fs ? `Статья готова: из ${fs.filter((f) => f.on).length} фактов` : "Статья готова" });
    });

  const switchTab = (t) =>
    run("tab", async () => {
      if (t === tab) return;
      if (dirty) await saveIfDirty();
      tabRef.current = t;
      setTab(t);
      const next = pick(detail, t);
      setDraft(next ? fieldsOf(next) : null);
    });

  const publish = (confirm = false) =>
    run("publish", async () => {
      await saveIfDirty();
      const p = await api.publish(post.id, confirm);
      if (!p) return;
      setStale(null);
      await load();
      onChanged();
      toast({ kind: "ok", text: `Опубликовано в ${cfg.channel}`, link: p.link ? { href: p.link, label: "Открыть пост" } : null });
    });

  const save = () =>
    run("save", async () => {
      await saveIfDirty();
      onChanged();
      toast({ kind: "ok", text: "Правки сохранены" });
    });

  const schedule = () =>
    run("schedule", async () => {
      await saveIfDirty();
      await api.schedule(post.id, new Date(schedAt).toISOString());
      setSchedOpen(false);
      await load();
      onChanged();
      toast({ kind: "ok", text: `Запланировано на ${clock(new Date(schedAt).toISOString())}` });
    });

  const unschedule = () =>
    run("schedule", async () => {
      await api.unschedule(post.id);
      await load();
      onChanged();
      toast({ text: "Публикация снята с очереди" });
    });

  const copy = async () => {
    const ok = await copyPost(shown, cfg);
    if (ok) {
      toast({ kind: "ok", text: `${noun} скопирован${isArticle ? "а" : ""} — вставь в Telegram` });
      return;
    }
    // буфер обмена закрыт (встроенное окно) — выделяем текст поста, остаётся ⌘C / Ctrl+C
    const bubble = document.querySelector(".editor .tg-bubble");
    const sel = window.getSelection?.();
    if (bubble && sel) {
      const range = document.createRange();
      range.selectNodeContents(bubble);
      sel.removeAllRanges();
      sel.addRange(range);
    }
    toast({ kind: "err", text: "Буфер обмена недоступен: текст поста выделен — нажми ⌘C или Ctrl+C" });
  };

  const dismiss = () =>
    run("dismiss", async () => {
      await (detail.status === "dismissed" ? api.restore(detail.id) : api.dismiss(detail.id));
      await load(true);
      onChanged();
    });

  const openSchedule = () => {
    const d = new Date(Date.now() + 60 * 60 * 1000);
    d.setMinutes(Math.ceil(d.getMinutes() / 5) * 5, 0, 0);
    setSchedAt(toLocalInput(d));
    setSchedOpen(true);
  };

  return (
    <section className={`editor ${open ? "open" : ""}`} aria-label="Редактор поста">
      <div className="editor-inner">
        <button className="btn small ghost ed-back" onClick={onClose}>
          <IconBack /> Лента
        </button>

        <header className="ed-head">
          <div className="ed-sub">
            {post?.status !== "published" && (
              <button className="btn small ghost ed-hide" onClick={dismiss} disabled={!!busy}>
                {detail.status === "dismissed" ? "Вернуть в ленту" : "Скрыть сюжет"}
              </button>
            )}
            <span className={`prio ${detail.priority}`}>{detail.priority}</span>
            <span className="scene-code" title={SCENES[detail.scene]?.label}>{SCENES[detail.scene]?.code ?? "—"}</span>
            {detail.tags.map((t) => (
              <span key={t} className={`tag t-${t}`}>#{t}</span>
            ))}
            <span className={`badge ${detail.verification === "insider_only" ? "insider" : ""}`}>
              {VERIFICATION[detail.verification]}
            </span>
          </div>
          <h1 className="ed-title">{detail.title}</h1>
          {detail.title_orig && detail.title_orig !== detail.title && <p className="ed-orig">{detail.title_orig}</p>}
        </header>

        <Origin story={detail} staleAfter={cfg.stale_after_hours} now={now} />

        {(post?.warnings?.length > 0 || detail.republish_of || detail.needs_refresh) && (
          <ul className="warnings">
            {detail.needs_refresh && (
              <li><IconWarn /> После черновика пришли новые источники — перегенерируй, чтобы учесть их.</li>
            )}
            {post?.warnings?.map((w) => (
              <li key={w}><IconWarn /> {w}</li>
            ))}
          </ul>
        )}

        {cfg.articles && (
          <div className="kind-tabs" role="tablist" aria-label="Формат текста">
            <button role="tab" aria-selected={!isArticle} className="kind-tab" onClick={() => switchTab("post")} disabled={!!busy}>
              Пост
            </button>
            <button role="tab" aria-selected={isArticle} className="kind-tab" onClick={() => switchTab("article")} disabled={!!busy}>
              Статья{detail.current_article ? "" : " +"}
            </button>
          </div>
        )}

        {isArticle && (
          <FactPicker facts={facts} setFacts={setFacts} busy={busy} hasArticle={!!post}
            onCollect={collectFacts} onWrite={writeArticle} />
        )}

        {!post && isArticle ? null : !post ? (
          <section className="fields">
            <p className="muted">
              {viaSample
                ? "Поста ещё нет: в снимке готовы посты для CRITICAL и HIGH. Нажми — Claude напишет пост по источникам сюжета прямо здесь, обычно за 10–40 секунд. В первый раз Claude спросит разрешение."
                : "Поста ещё нет — авто-генерация пропускает старые и низкоприоритетные сюжеты. Сгенерируй вручную, это несколько секунд."}
            </p>
            <button className="btn primary" style={{ justifySelf: "start" }} onClick={() => generate("standard")}
              disabled={!!busy || !cfg.llm_ready}>
              {busy === "generate" ? <span className="spin" /> : <IconFlame />}
              {busy === "generate" && viaSample ? "Claude пишет пост…" : "Сгенерировать пост"}
            </button>
            {!cfg.llm_ready && <p className="notes">LLM не настроен — задай ANTHROPIC_API_KEY.</p>}
          </section>
        ) : (
          <>
            <div className="post-grid">
              <TelegramPreview post={shown} cfg={cfg} />
              <PostFields draft={draft ?? fieldsOf(post)} setDraft={setDraft} readOnly={readOnly} kind={post.kind} />
            </div>
            {post.editor_notes && <p className="notes">Заметка модели: {post.editor_notes}</p>}
            <p className="ed-sub">
              Версия {post.version}: {MODE[post.mode]}
              {post.from_facts ? `, из ${post.from_facts} отмеченных фактов` : ""}
              {post.llm_model ? `, ${post.llm_model}` : ""}
              {post.cache_read_tokens ? `, из кэша ${post.cache_read_tokens} токенов` : ""}
            </p>
            {post.status !== "published" && !cfg.bot_ready && !isArticle && (
              <p className="notes no-bot">
                {cfg.snapshot_at
                  ? "Бота в снимке нет: в канал отсюда ничего не уходит. «Отправить» и «Запланировать» работают на сервере, где задан TG_BOT_TOKEN."
                  : "Bot API не настроен (TG_BOT_TOKEN): отправка и расписание выключены. Скопируй пост и вставь в Telegram."}
              </p>
            )}
          </>
        )}

        <Timeline items={detail.items} now={now} />

        <div className={`actions ${post ? "" : "empty"}`}>
            {!post ? null : isArticle ? (
              <>
                <button className="btn primary" onClick={copy}>
                  <IconCopy /> Скопировать<span className="hide-m"> статью</span>
                </button>
                <button className="btn hot" onClick={() => generate("standard")} disabled={!!busy || !cfg.llm_ready}>
                  {busy === "generate" ? <span className="spin" /> : <IconFlame />}
                  {busy === "generate" ? "Claude пишет…" : "Переписать"}
                </button>
                {dirty && (
                  <button className="btn ghost" onClick={save} disabled={!!busy}>
                    {busy === "save" ? <span className="spin" /> : null}
                    Сохранить правки
                  </button>
                )}
              </>
            ) : post.status === "published" ? (
              <span className="done">
                <IconCheck /> Опубликовано {post.published_at ? stamp(post.published_at, now) : ""}
                {post.link && (
                  <a href={post.link} target="_blank" rel="noreferrer" className="link">
                    открыть
                  </a>
                )}
              </span>
            ) : schedOpen ? (
              <div className="sched">
                <label className="sr-only" htmlFor="sched-at">Время публикации</label>
                <input id="sched-at" type="datetime-local" value={schedAt} onChange={(e) => setSchedAt(e.target.value)} />
                <button className="btn primary" onClick={schedule} disabled={!schedAt || !!busy}>
                  {busy === "schedule" ? <span className="spin" /> : <IconClock />}
                  Запланировать на {schedAt ? schedAt.slice(11, 16) : "…"}
                </button>
                <button className="btn ghost" onClick={() => setSchedOpen(false)}>Отмена</button>
              </div>
            ) : (
              <>
                {cfg.bot_ready && (
                  <button className="btn primary" onClick={() => publish(false)} disabled={!!busy || readOnly}>
                    {busy === "publish" ? <span className="spin" /> : <IconSend />}
                    <span>Отправить<span className="hide-m"> в канал</span></span>
                  </button>
                )}
                {!cfg.bot_ready ? null : post.status === "scheduled" ? (
                  <button className="btn" onClick={unschedule} disabled={!!busy}>
                    <IconClock /> В {clock(post.scheduled_at)} — снять
                  </button>
                ) : (
                  <button className="btn" onClick={openSchedule} disabled={!!busy || readOnly}>
                    <IconClock /> Запланировать
                  </button>
                )}
                {!cfg.bot_ready && (
                  <button className="btn primary" onClick={copy}>
                    <IconCopy /> Скопировать<span className="hide-m"> для Telegram</span>
                  </button>
                )}
                <button className="btn hot" onClick={() => generate(nextMode)} disabled={!!busy || readOnly || !cfg.llm_ready}>
                  {busy === "harder" ? <span className="spin" /> : <IconFlame />}
                  {busy === "harder" && viaSample ? "Claude пишет…" : post.mode === "harder" ? "Ещё жёстче" : post.mode === "hardest" ? "Ещё вариант" : (
                    <span><span className="hide-m">Перегенерировать жёстче</span><span className="show-m">Жёстче</span></span>
                  )}
                </button>
                {cfg.bot_ready && (
                  <button className="btn ghost" onClick={copy}>
                    <IconCopy /> Скопировать
                  </button>
                )}
                {dirty && (
                  <button className="btn ghost" onClick={save} disabled={!!busy}>
                    {busy === "save" ? <span className="spin" /> : null}
                    Сохранить правки
                  </button>
                )}
              </>
            )}
        </div>
      </div>
      {stale && (
        <StaleDialog reason={stale} busy={busy === "publish"} onCancel={() => setStale(null)} onConfirm={() => publish(true)} />
      )}
    </section>
  );
}
