import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Дев-сервер проксирует /api на бэкенд (uvicorn :8000) — CORS не нужен.
// VITE_DEMO=1 — демо-режим: встроенные вымышленные данные, без бэкенда.
export default defineConfig({
  plugins: [react()],
  define: {
    __DEMO__: JSON.stringify(process.env.VITE_DEMO === "1"),
  },
  server: {
    port: 5173,
    proxy: {
      "/api": { target: process.env.VITE_API ?? "http://localhost:8000", changeOrigin: true },
    },
  },
  build: { target: "es2022", sourcemap: false },
});
