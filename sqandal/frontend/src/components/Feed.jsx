import { compactAgo, clock, DAY_ONLY, day, plural, SCENES, STATUS, VERIFICATION } from "../lib/format.js";
import { IconChevron, IconSearch } from "./Icons.jsx";

const TABS = [
  { key: "ru", label: "СНГ" },
  { key: "de", label: "GER" },
  { key: "us", label: "US" },
  { key: "all", label: "ALL" },
];

export function SceneTabs({ scene, onScene, counts, critical }) {
  const total = Object.values(counts ?? {}).reduce((a, b) => a + b, 0);
  return (
    <div className="scenes" role="tablist" aria-label="Сцена">
      {TABS.map((t) => {
        const n = t.key === "all" ? total : counts?.[t.key] ?? 0;
        return (
          <button
            key={t.key}
            role="tab"
            aria-selected={scene === t.key}
            className="scene"
            onClick={() => onScene(t.key)}
          >
            <span className="scene-name">{t.label}</span>
            <span className="scene-count">{n}</span>
            {critical.has(t.key) && <span className="crit-dot" title="Есть CRITICAL" />}
          </button>
        );
      })}
    </div>
  );
}

function Select({ value, onChange, label, children }) {
  return (
    <span className="select-wrap">
      <select className="select" value={value} onChange={onChange} aria-label={label}>
        {children}
      </select>
      <IconChevron />
    </span>
  );
}

export function Filters({ filters, setFilters, tags }) {
  const toggle = (key, value) =>
    setFilters((f) => {
      const set = new Set(f[key]);
      set.has(value) ? set.delete(value) : set.add(value);
      return { ...f, [key]: [...set] };
    });
  return (
    <div className="filters">
      {["CRITICAL", "HIGH", "NORMAL"].map((p) => (
        <button
          key={p}
          className={`chip p-${p}`}
          aria-pressed={filters.priorities.includes(p)}
          onClick={() => toggle("priorities", p)}
        >
          {p}
        </button>
      ))}
      {tags.map((t) => (
        <button
          key={t.name}
          className="chip"
          aria-pressed={filters.tag === t.name}
          onClick={() => setFilters((f) => ({ ...f, tag: f.tag === t.name ? "" : t.name }))}
        >
          #{t.name}
        </button>
      ))}
      <span className="grow" />
      <label className="search">
        <IconSearch />
        <input
          id="feed-search"
          type="search"
          placeholder="Артист или заголовок"
          value={filters.q}
          onChange={(e) => setFilters((f) => ({ ...f, q: e.target.value }))}
          aria-label="Поиск"
        />
      </label>
      <Select value={filters.status} label="Статус" onChange={(e) => setFilters((f) => ({ ...f, status: e.target.value }))}>
        <option value="active">В работе</option>
        <option value="published">Опубликованные</option>
        <option value="dismissed">Скрытые</option>
        <option value="all">Все</option>
      </Select>
      <Select value={filters.sort} label="Сортировка" onChange={(e) => setFilters((f) => ({ ...f, sort: e.target.value }))}>
        <option value="time">Сначала свежие</option>
        <option value="priority">Сначала важные</option>
      </Select>
    </div>
  );
}

export function StoryRow({ s, selected, flash, onOpen, now }) {
  const staleish = s.is_stale || s.republish_of;
  const others = s.sources_count - 1;
  return (
    <button
      className={`row p-${s.priority} ${flash ? "flash" : ""} ${s.status === "dismissed" ? "dim" : ""}`}
      aria-current={selected}
      onClick={() => onOpen(s.id)}
    >
      <div className="row-top">
        <span className={`prio ${s.priority}`}>{s.priority}</span>
        <span className="scene-code" title={SCENES[s.scene]?.label}>{SCENES[s.scene]?.code ?? "—"}</span>
        <span className="row-src">{s.origin.source_name}{others > 0 ? ` и ещё ${others}` : ""}</span>
        {DAY_ONLY.has(s.origin.date_source) ? (
          <span className="row-time" title="Первоисточник: известна только дата">
            <b>{day(s.origin.published_at, now)}</b>
          </span>
        ) : (
          <span className="row-time" title="Время первоисточника">
            <b>{clock(s.origin.published_at)}</b>&nbsp; {compactAgo(s.origin.published_at, now)}
          </span>
        )}
      </div>
      <div className="row-title">{s.title}</div>
      <div className="row-meta">
        {s.tags.map((t) => (
          <span key={t} className={`tag t-${t}`}>#{t}</span>
        ))}
        {s.verification === "insider_only" && <span className="badge insider">{VERIFICATION.insider_only}</span>}
        {s.verification === "multi_source" && (
          <span className="badge">{s.sources_count} {plural(s.sources_count, "источник", "источника", "источников")}</span>
        )}
        {staleish && <span className="badge stale">{s.is_stale ? "старый факт" : "повтор"}</span>}
        {s.post?.status === "scheduled" ? (
          <span className="badge st-scheduled">в {clock(s.post.scheduled_at)}</span>
        ) : s.status !== "new" ? (
          <span className={`badge st-${s.status}`}>{STATUS[s.status]}</span>
        ) : null}
        {s.needs_refresh && <span className="badge insider">новые источники</span>}
      </div>
    </button>
  );
}

export function Feed({ stories, loading, error, selectedId, flashIds, onOpen, onCollect, now }) {
  if (error) {
    return (
      <section className="feed">
        <div className="feed-empty">
          <h3>Ленту не загрузить</h3>
          <p>{error}</p>
        </div>
      </section>
    );
  }
  if (!loading && stories.length === 0) {
    return (
      <section className="feed">
        <div className="feed-empty">
          <h3>Тихо</h3>
          <p>
            {onCollect
              ? "По этим фильтрам сюжетов нет. Сбор идёт каждые 15 минут, Telegram-каналы — сразу."
              : "По этим фильтрам в снимке сюжетов нет."}
          </p>
          {onCollect && <button className="btn small" onClick={onCollect}>Собрать сейчас</button>}
        </div>
      </section>
    );
  }
  const crit = stories.filter((s) => s.priority === "CRITICAL").length;
  return (
    <section className="feed" aria-label="Лента сюжетов" aria-busy={loading}>
      <div className="feed-head">
        <span>Сюжеты · {stories.length}{crit ? ` · CRITICAL ${crit}` : ""}</span>
      </div>
      <div className="feed-list">
        {stories.map((s) => (
          <StoryRow key={s.id} s={s} selected={s.id === selectedId} flash={flashIds.has(s.id)} onOpen={onOpen} now={now} />
        ))}
      </div>
    </section>
  );
}
