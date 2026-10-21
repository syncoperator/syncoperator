// Клиент API дашборда. Токен — Bearer в заголовке (для SSE — ?token=).

export class ApiError extends Error {
  constructor(status, detail, code) {
    super(detail || `HTTP ${status}`);
    this.status = status;
    this.code = code;
  }
}

export function createApi(token, base = "") {
  async function req(method, path, body) {
    let res;
    try {
      res = await fetch(base + path, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          ...(body !== undefined ? { "Content-Type": "application/json" } : {}),
        },
        body: body !== undefined ? JSON.stringify(body) : undefined,
      });
    } catch {
      throw new ApiError(0, "Сервер недоступен — проверь, что API запущен");
    }
    let data = null;
    try {
      data = await res.json();
    } catch {
      data = null;
    }
    if (!res.ok) {
      const detail = Array.isArray(data?.detail) ? data.detail.map((d) => d.msg).join("; ") : data?.detail;
      throw new ApiError(res.status, detail, data?.code);
    }
    return data;
  }

  return {
    demo: false,
    config: () => req("GET", "/api/config"),
    stats: () => req("GET", "/api/stats"),
    sources: () => req("GET", "/api/sources"),
    stories: (params = {}) => {
      const q = new URLSearchParams(Object.entries(params).filter(([, v]) => v !== undefined && v !== ""));
      return req("GET", `/api/stories?${q}`);
    },
    story: (id) => req("GET", `/api/stories/${id}`),
    generate: (id, mode = "standard") => req("POST", `/api/stories/${id}/generate`, { mode }),
    dismiss: (id) => req("POST", `/api/stories/${id}/dismiss`),
    restore: (id) => req("POST", `/api/stories/${id}/restore`),
    patchPost: (id, fields) => req("PATCH", `/api/posts/${id}`, fields),
    publish: (id, confirmStale = false) => req("POST", `/api/posts/${id}/publish`, { confirm_stale: confirmStale }),
    schedule: (id, iso) => req("POST", `/api/posts/${id}/schedule`, { scheduled_at: iso }),
    unschedule: (id) => req("DELETE", `/api/posts/${id}/schedule`),
    collect: () => req("POST", "/api/collect"),
    subscribe(onStory, onStats, onStatus) {
      if (typeof EventSource === "undefined") return () => {};
      const es = new EventSource(`${base}/api/stream?token=${encodeURIComponent(token)}`);
      es.addEventListener("story", (e) => onStory(JSON.parse(e.data)));
      es.addEventListener("stats", (e) => onStats(JSON.parse(e.data)));
      es.onopen = () => onStatus(true);
      es.onerror = () => onStatus(false);
      return () => es.close();
    },
  };
}
