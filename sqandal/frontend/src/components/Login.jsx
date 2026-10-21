import { useState } from "react";
import { IconSpeaker } from "./Icons.jsx";

export function Wordmark({ sub = true }) {
  return (
    <span className="mark">
      <span className="mark-tile" aria-hidden="true">
        <IconSpeaker size={22} />
      </span>
      <span className="mark-text">
        <span className="mark-word">SQANDAL</span>
        {sub && <span className="mark-sub">Radar</span>}
      </span>
    </span>
  );
}

export function Login({ onLogin, error, busy }) {
  const [token, setToken] = useState("");
  return (
    <main className="login">
      <form
        onSubmit={(e) => {
          e.preventDefault();
          if (token.trim()) onLogin(token.trim());
        }}
      >
        <Wordmark sub={false} />
        <p>Мониторинг хип-хоп повестки и выпуск постов для @SQANDALru.</p>
        <div className="field">
          <label htmlFor="token">Токен дашборда</label>
          <input
            id="token"
            className="input"
            type="password"
            autoComplete="current-password"
            value={token}
            onChange={(e) => setToken(e.target.value)}
            placeholder="DASHBOARD_TOKEN из .env"
            autoFocus
          />
        </div>
        {error && <p className="error">{error}</p>}
        <button className="btn primary" type="submit" disabled={busy || !token.trim()}>
          {busy ? <span className="spin" /> : null}
          Войти
        </button>
        <a href="?demo" className="btn ghost">Открыть демо без сервера</a>
      </form>
    </main>
  );
}
