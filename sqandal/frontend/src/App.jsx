import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Editor } from "./components/Editor.jsx";
import { Feed, Filters, SceneTabs } from "./components/Feed.jsx";
import { Header, SourcesDrawer } from "./components/Header.jsx";
import { IconInfo } from "./components/Icons.jsx";
import { Toasts } from "./components/Toasts.jsx";
import { stamp } from "./lib/format.js";

const ACTIVE = new Set(["new", "drafted", "scheduled"]);

function matches(s, f) {
  if (f.scene !== "all" && s.scene !== f.scene) return false;
  if (f.priorities.length && !f.priorities.includes(s.priority)) return false;
  if (f.tag && !s.tags.includes(f.tag)) return false;
  if (f.status === "active" && !ACTIVE.has(s.status)) return false;
  if (!["active", "all"].includes(f.status) && s.status !== f.status) return false;
  if (f.q && !s.title.toLowerCase().includes(f.q.toLowerCase())) return false;
  return true;
}

function sortStories(list, sort) {
  const rank = { CRITICAL: 3, HIGH: 2, NORMAL: 1 };
  return [...list].sort((a, b) =>
    sort === "priority"
      ? rank[b.priority] - rank[a.priority] || b.origin.published_at.localeCompare(a.origin.published_at)
      : b.origin.published_at.localeCompare(a.origin.published_at),
  );
}

function useNow(ms = 30000) {
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    const t = setInterval(() => setNow(Date.now()), ms);
    return () => clearInterval(t);
  }, [ms]);
  return now;
}

export default function App({ api, onLogout, onDemoMode }) {
  const [cfg, setCfg] = useState(null);
  const [stats, setStats] = useState(null);
  const [connected, setConnected] = useState(false);
  // в реальных новостях сначала важные
  const [filters, setFilters] = useState({
    scene: "all", priorities: [], tag: "", q: "", status: "active",
    sort: api.kind === "snapshot" || api.kind === "live" ? "priority" : "time",
  });
  const [stories, setStories] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [selectedId, setSelectedId] = useState(null);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [flashIds, setFlashIds] = useState(new Set());
  const [toasts, setToasts] = useState([]);
  const [sourcesOpen, setSourcesOpen] = useState(false);
  const [collecting, setCollecting] = useState(false);
  const known = useRef(new Set());
  const autoPicked = useRef(false);
  const statsTimer = useRef(null);
  const filtersRef = useRef(filters);
  filtersRef.current = filters;
  const now = useNow();

  const toast = useCallback((t) => {
    const id = Math.random().toString(36).slice(2);
    setToasts((xs) => [...xs.slice(-3), { id, ...t }]);
    setTimeout(() => setToasts((xs) => xs.filter((x) => x.id !== id)), t.kind === "err" ? 9000 : 6000);
  }, []);

  useEffect(() => {
    api.config().then(setCfg).catch((e) => setError(e.message));
    api.stats().then(setStats).catch(() => {});
  }, [api]);

  // счётчики и точки на вкладках — из статистики; после событий эфира обновляем её сразу
  // (пачка событий схлопывается в один запрос)
  const refreshStats = useCallback(() => {
    clearTimeout(statsTimer.current);
    statsTimer.current = setTimeout(() => api.stats().then(setStats).catch(() => {}), 400);
  }, [api]);

  const reload = useCallback(async () => {
    const f = filtersRef.current;
    try {
      const res = await api.stories({
        scene: f.scene,
        priority: f.priorities.join(","),
        tag: f.tag,
        q: f.q,
        status: f.status,
        sort: f.sort,
      });
      res.stories.forEach((s) => known.current.add(s.id));
      setStories(res.stories);
      // на широком экране сразу открыть верхний CRITICAL (или первый сюжет) — редактор не пустой
      if (!autoPicked.current && res.stories.length && window.matchMedia?.("(min-width: 901px)").matches) {
        autoPicked.current = true;
        const top = res.stories.find((s) => s.priority === "CRITICAL") ?? res.stories[0];
        setSelectedId((cur) => cur ?? top.id);
      }
      setError(null);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [api]);

  useEffect(() => {
    setLoading(true);
    const t = setTimeout(reload, filters.q ? 250 : 0);
    return () => clearTimeout(t);
  }, [filters, reload]);

  // эфир: SSE с обновлёнными карточками; новый CRITICAL — сирена и уведомление
  useEffect(() => {
    const unsub = api.subscribe(
      (card) => {
        const isNew = !known.current.has(card.id);
        known.current.add(card.id);
        setStories((list) => {
          const rest = list.filter((s) => s.id !== card.id);
          return matches(card, filtersRef.current) ? sortStories([card, ...rest], filtersRef.current.sort) : rest;
        });
        refreshStats();
        if (isNew && card.priority === "CRITICAL") {
          setFlashIds((s) => new Set(s).add(card.id));
          setTimeout(() => setFlashIds((s) => { const n = new Set(s); n.delete(card.id); return n; }), 4000);
          toast({ kind: "crit", text: `CRITICAL: ${card.title}`, action: { label: "Открыть", run: () => { setSelectedId(card.id); setMobileOpen(true); } } });
        }
      },
      setStats,
      setConnected,
      toast,
    );
    const poll = setInterval(() => api.stats().then(setStats).catch(() => {}), 60000);
    return () => {
      unsub();
      clearInterval(poll);
      clearTimeout(statsTimer.current);
    };
  }, [api, toast, refreshStats]);

  // Красная точка — CRITICAL в сцене по ВСЕЙ базе, а не по отфильтрованной ленте:
  // на вкладке СНГ точка у US не должна пропадать.
  const critical = useMemo(() => {
    const set = new Set(
      Object.entries(stats?.critical_by_scene ?? {})
        .filter(([, n]) => n > 0)
        .map(([scene]) => scene),
    );
    for (const s of stories) if (s.priority === "CRITICAL" && ACTIVE.has(s.status)) set.add(s.scene);
    return set;
  }, [stats, stories]);

  const selectedCard = stories.find((s) => s.id === selectedId);

  const collect = async () => {
    setCollecting(true);
    try {
      await api.collect();
      toast({ text: live ? "Сбор запущен: Claude соберёт новости и напишет посты — обычно 5–15 минут, лента обновится сама" : "Сбор запущен — новые сюжеты появятся в ленте сами" });
      api.stats().then(setStats).catch(() => {});
    } catch (e) {
      toast({ kind: "err", text: e.message });
    } finally {
      setTimeout(() => setCollecting(false), 1500);
    }
  };

  if (!cfg) {
    return (
      <div className="login">
        <p className={error ? "error" : "muted"}>{error ?? "Подключаюсь…"}</p>
        {error && <button className="btn" onClick={onLogout}>Сменить токен</button>}
      </div>
    );
  }

  const snapshot = Boolean(cfg.snapshot_at);
  const live = cfg.collect_via === "trigger";

  return (
    <div className="app">
      <div>
        <Header
          stats={stats}
          connected={connected}
          onCollect={snapshot ? null : collect}
          collecting={collecting}
          onSources={() => setSourcesOpen(true)}
          onLogout={onLogout}
          demo={api.demo}
          snapshotAt={cfg.snapshot_at}
          live={live}
          botReady={cfg.bot_ready}
        />
        {api.demo && (
          <div className="demo-note">
            <span className="note-i" aria-hidden="true"><IconInfo /></span>
            {live ? (
              <span>
                Реальные новости по-русски. «Обновить» — Claude соберёт свежие новости с сайтов и напишет посты (5–15 минут,
                лента обновится сама). В канал ничего не уходит.
              </span>
            ) : snapshot ? (
              <span>
                Реальные новости: снимок от {stamp(cfg.snapshot_at)}, лента не обновляется, в канал ничего не уходит.
                Обновлять ленту можно, когда страница открыта в Claude.
              </span>
            ) : (
              <span>
                Демо: артисты и издания вымышлены, «публикация» никуда не уходит. Через несколько секунд в эфир придёт
                новый CRITICAL.
              </span>
            )}
            {onDemoMode && (
              <button type="button" className="demo-switch" onClick={() => onDemoMode(snapshot || live ? "fiction" : "real")}>
                {snapshot || live ? "Вымышленное демо →" : "Реальные новости →"}
              </button>
            )}
          </div>
        )}
      </div>
      <div className="controls">
        <SceneTabs scene={filters.scene} onScene={(scene) => setFilters((f) => ({ ...f, scene }))}
          counts={stats?.by_scene} critical={critical} />
        <Filters filters={filters} setFilters={setFilters} tags={cfg.tags} />
      </div>
      <div className="main">
        <Feed
          stories={stories}
          loading={loading}
          error={error}
          selectedId={selectedId}
          flashIds={flashIds}
          onOpen={(id) => {
            setSelectedId(id);
            setMobileOpen(true);
          }}
          onCollect={snapshot ? null : collect}
          now={now}
        />
        <Editor
          api={api}
          cfg={cfg}
          storyId={selectedId}
          cardVersion={selectedCard?.updated_at}
          open={mobileOpen}
          onClose={() => setMobileOpen(false)}
          toast={toast}
          onChanged={() => {
            reload();
            api.stats().then(setStats).catch(() => {});
          }}
          now={now}
        />
      </div>
      {sourcesOpen && <SourcesDrawer api={api} onClose={() => setSourcesOpen(false)} />}
      <Toasts items={toasts} onClose={(id) => setToasts((xs) => xs.filter((x) => x.id !== id))} />
    </div>
  );
}
