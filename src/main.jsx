import "./index.css";
import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// ── Google Sheets sync URL ──────────────────────────────────────────────────
// Pega aquí la URL que te entrega Apps Script al implementar Code_Bitacora.gs
// Ejemplo: "https://script.google.com/macros/s/AKfycb.../exec"
window.__SHEETS_URL__ = "";

// ── Storage polyfill (localStorage + sincronización con Google Sheets) ─────
// - set(): guarda local Y envía al Sheet (si hay URL configurada).
// - get(): intenta traer la versión más reciente del Sheet primero (para que
//   todos los dispositivos vean los mismos datos); si no hay internet o no
//   hay URL configurada, usa la copia local guardada en este dispositivo.
const PREFIX = "bitacora_";
window.storage = {
  async get(key, shared) {
    if (window.__SHEETS_URL__) {
      try {
        const res = await fetch(`${window.__SHEETS_URL__}?key=${encodeURIComponent(key)}`);
        const data = await res.json();
        if (data && data.ok && data.value !== null && data.value !== undefined) {
          localStorage.setItem(PREFIX + key, data.value);
          return { key, value: data.value, shared: !!shared };
        }
      } catch {
        // sin internet o el Sheet no respondió — sigue con la copia local
      }
    }
    try {
      const raw = localStorage.getItem(PREFIX + key);
      return raw ? { key, value: raw, shared: !!shared } : null;
    } catch { return null; }
  },
  async set(key, value, shared) {
    try {
      localStorage.setItem(PREFIX + key, value);
      if (window.__SHEETS_URL__) {
        fetch(window.__SHEETS_URL__, {
          method: "POST",
          body: JSON.stringify({ type: "storage", key, value }),
        }).catch(() => {});
      }
      return { key, value, shared: !!shared };
    } catch { return null; }
  },
  async delete(key, shared) {
    try {
      localStorage.removeItem(PREFIX + key);
      if (window.__SHEETS_URL__) {
        fetch(window.__SHEETS_URL__, {
          method: "POST",
          body: JSON.stringify({ type: "storage", key, value: "" }),
        }).catch(() => {});
      }
      return { key, deleted: true, shared: !!shared };
    } catch { return null; }
  },
  async list(prefix, shared) {
    try {
      const keys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(PREFIX + (prefix || "")))
          keys.push(k.slice(PREFIX.length));
      }
      return { keys, prefix, shared: !!shared };
    } catch { return { keys: [] }; }
  },
};

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
