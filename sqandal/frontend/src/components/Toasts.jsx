import { IconX } from "./Icons.jsx";

export function Toasts({ items, onClose }) {
  return (
    <div className="toasts" role="status" aria-live="polite">
      {items.map((t) => (
        <div key={t.id} className={`toast ${t.kind ?? ""}`}>
          <div>
            <div>{t.text}</div>
            {t.link && (
              <a href={t.link.href} target="_blank" rel="noreferrer" className="link">
                {t.link.label}
              </a>
            )}
            {t.action && (
              <button className="btn small ghost" style={{ marginTop: 6 }} onClick={t.action.run}>
                {t.action.label}
              </button>
            )}
          </div>
          <button aria-label="Закрыть уведомление" onClick={() => onClose(t.id)}>
            <IconX />
          </button>
        </div>
      ))}
    </div>
  );
}
