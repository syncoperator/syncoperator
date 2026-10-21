import { useEffect, useState } from "react";
import { ago, clock, plural, SCENES, stamp } from "../lib/format.js";
import { IconAntenna, IconRefresh, IconX } from "./Icons.jsx";
import { Wordmark } from "./Login.jsx";

function CollectState({ info }) {
  const st = info?.busy ? info.state : info?.state === "requested" || info?.state === "running" ? "stuck" : info?.state;
  const label = {
    requested: "Запускаю сбор…",
    running: `Claude собирает новости${info?.started_at ? ` с ${clock(info.started_at)}` : ""}…`,
    done: `Обновлено ${info?.finished_at ? stamp(info.finished_at) : ""}${info?.message ? ` · ${info.message}` : ""}`,
    error: `Сбор не удался${info?.message ? `: ${info.message}` : ""}`,
    stuck: "Прошлый сбор не отозвался — можно запустить снова",
  }[st] ?? "Нажми «Обновить», чтобы собрать свежие новости";
  const short = {
    requested: "Запуск…",
    running: "Идёт сбор…",
    done: info?.finished_at ? `Обновлено ${clock(info.finished_at)}` : "Обновлено",
    error: "Сбор не удался",
    stuck: "Сбор не отозвался",
  }[st] ?? "";
  return (
    <span className={`live collect-${st ?? "idle"}`} title={label}>
      <span className="live-dot" />
      <span className="live-text hide-m">{label}</span>
      <span className="live-text show-m">{short}</span>
    </span>
  );
}

export function Header({ stats, connected, onCollect, collecting, onSources, onLogout, demo, snapshotAt, live, botReady = true }) {
  const crit = stats?.by_priority?.CRITICAL ?? 0;
  const high = stats?.by_priority?.HIGH ?? 0;
  const running = collecting || Boolean(stats?.collect?.busy);
  return (
    <header className="top">
      {/* строка состояния: сбор слева, счётчики справа, под ней — шкала */}
      <div className="statusbar">
        {live ? (
          <CollectState info={stats?.collect} />
        ) : snapshotAt ? (
          <span className="live snap" title="Лента не обновляется: это снимок реальных новостей на момент сбора">
            <span className="live-dot" />
            <span className="live-text">Снимок: {stamp(snapshotAt)}</span>
          </span>
        ) : (
          <span className={`live ${connected ? "on" : ""}`} title={connected ? "Лента обновляется в реальном времени" : "Переподключение…"}>
            <span className="live-dot" />
            <span className="live-text">
              {connected ? "В эфире" : "Нет связи"}
              {stats?.last_collect_at ? `, сбор в ${clock(stats.last_collect_at)}` : ""}
            </span>
          </span>
        )}
        <div className="top-stats">
          <span className="stat crit"><b>{crit}</b> CRITICAL</span>
          <span className="stat high"><b>{high}</b> HIGH</span>
          {botReady ? (
            <>
              <span className="stat"><b>{stats?.drafts ?? 0}</b> черновиков</span>
              <span className="stat"><b>{stats?.scheduled ?? 0}</b> в очереди</span>
              <span className="stat"><b>{stats?.published_24h ?? 0}</b> за сутки</span>
            </>
          ) : (
            <span className="stat"><b>{stats?.total ?? 0}</b> {plural(stats?.total ?? 0, "новость", "новости", "новостей")}</span>
          )}
        </div>
      </div>
      <div className="brandbar">
        <Wordmark />
        <div className="top-actions">
          {onCollect && (
            <button className={`btn ${live ? "primary" : ""}`} onClick={onCollect} disabled={running}
              aria-label={live ? "Обновить новости" : "Собрать сейчас"}>
              {running ? <span className="spin" /> : <IconRefresh />}
              <span className={live ? "" : "hide-m"}>{live ? (running ? "Идёт сбор" : "Обновить") : "Собрать сейчас"}</span>
            </button>
          )}
          <button className="btn icon-text" onClick={onSources} aria-label="Источники">
            <IconAntenna />
            <span className="hide-m">Источники</span>
            {stats?.sources_failing ? <b className="fail-count">{stats.sources_failing}</b> : null}
          </button>
          {!demo && (
            <button className="btn ghost" onClick={onLogout}>
              Выйти
            </button>
          )}
        </div>
      </div>
    </header>
  );
}

export function SourcesDrawer({ api, onClose }) {
  const [rows, setRows] = useState(null);
  const [error, setError] = useState(null);
  useEffect(() => {
    api.sources().then(setRows).catch((e) => setError(e.message));
  }, [api]);
  useEffect(() => {
    const onKey = (e) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
  return (
    <>
      <div className="overlay scrim" onClick={onClose} />
      <aside className="drawer" aria-label="Источники">
        <div className="drawer-head">
          <div>
            <span className="eyebrow">Что читает сбор</span>
            <h2>Источники</h2>
          </div>
          <button className="btn icon" onClick={onClose} aria-label="Закрыть">
            <IconX />
          </button>
        </div>
        <div className="drawer-body">
          {error && <p className="err">{error}</p>}
          {!rows && !error && <p className="muted">Загружаю…</p>}
          {rows?.map((s) => {
            const health = !s.enabled ? "off" : s.consecutive_failures ? "bad" : "";
            return (
              <div className={`src ${health}`} key={s.key}>
                <div className="src-name">
                  <span className={`src-health ${health}`} />
                  <span className="scene-code">{SCENES[s.scene]?.code ?? "—"}</span>
                  <span className="src-title">{s.name}</span>
                  <small className="src-kind">{s.kind === "telegram" ? "Telegram" : s.kind === "rss" ? "RSS" : "сайт"}</small>
                </div>
                <small>
                  {!s.enabled
                    ? "выключен"
                    : s.last_success_at
                      ? `последний успешный сбор ${ago(s.last_success_at)}`
                      : "ещё не собирался"}
                  {s.items_total ? `, всего публикаций: ${s.items_total}` : ""}
                </small>
                {s.last_error && s.enabled && <span className="err">{s.last_error}</span>}
                {s.note && <small>{s.note}</small>}
              </div>
            );
          })}
        </div>
      </aside>
    </>
  );
}
