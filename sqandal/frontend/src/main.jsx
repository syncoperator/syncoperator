import { StrictMode, useCallback, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import { Login } from "./components/Login.jsx";
import { createApi } from "./lib/api.js";
import "./styles.css";

const KEY = "sqandal.token";
const params = new URLSearchParams(location.search);
const wantsDemo = __DEMO__ || params.has("demo");
// демо-сборка: «real» — снимок реальных новостей, «fiction» — вымышленный эфир
// «real»: база артефакта с кнопкой «Обновить»; нет базы (страница не в Claude) — встроенный снимок
const DEMO_APIS = {
  real: async () => {
    const live = await (await import("./demo/liveApi.js")).createLiveApi();
    return live ?? (await import("./demo/snapshotApi.js")).createSnapshotApi();
  },
  fiction: () => import("./demo/demoApi.js").then((m) => m.createDemoApi()),
};

function readToken() {
  try {
    return localStorage.getItem(KEY) || "";
  } catch {
    return "";
  }
}

function Root() {
  const [api, setApi] = useState(null);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const [demoMode, setDemoMode] = useState(params.get("demo") === "fiction" ? "fiction" : "real");

  useEffect(() => {
    if (wantsDemo) {
      // демо-данные грузятся отдельным чанком только по требованию
      let alive = true;
      setApi(null);
      DEMO_APIS[demoMode]().then((a) => alive && setApi(a));
      return () => {
        alive = false;
      };
    }
    const t = readToken();
    if (t) setApi(createApi(t));
  }, [demoMode]);

  const login = useCallback(async (token) => {
    setBusy(true);
    setError(null);
    const candidate = createApi(token);
    try {
      await candidate.config();
      try {
        localStorage.setItem(KEY, token);
      } catch {
        /* приватный режим — живём без запоминания */
      }
      setApi(candidate);
    } catch (e) {
      setError(e.status === 401 ? "Токен не подошёл — проверь DASHBOARD_TOKEN в .env" : e.message);
    } finally {
      setBusy(false);
    }
  }, []);

  const logout = useCallback(() => {
    try {
      localStorage.removeItem(KEY);
    } catch {
      /* ignore */
    }
    setApi(null);
  }, []);

  if (!api) {
    return wantsDemo ? (
      <div className="login"><p style={{ color: "var(--ink-2)" }}>Загружаю ленту…</p></div>
    ) : (
      <Login onLogin={login} error={error} busy={busy} />
    );
  }
  return <App key={wantsDemo ? demoMode : "live"} api={api} onLogout={logout} onDemoMode={wantsDemo ? setDemoMode : null} />;
}

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <Root />
  </StrictMode>,
);
