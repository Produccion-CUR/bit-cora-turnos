import "./index.css";
import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// ── Google Sheets sync URL ──────────────────────────────────────────────────
// Pega aquí la URL que te entrega Apps Script al implementar Code_Bitacora.gs
// Ejemplo: "https://script.google.com/macros/s/AKfycb.../exec"
window.__SHEETS_URL__ = "https://script.google.com/macros/s/AKfycbxjT0kYoffxpMcHGbydcRKU05JdSPzg7WEhIF_S3eO0AWux4qlSTvzg2ZZPq_r8iEG7/exec";

// ── Storage polyfill (localStorage + sincronización con Google Sheets) ─────
// - set(): guarda local Y envía al Sheet (si hay URL configurada).
// - get(): intenta traer la versión más reciente del Sheet primero (para que
//   todos los dispositivos vean los mismos datos); si no hay internet o no
//   hay URL configurada, usa la copia local guardada en este dispositivo.
const PREFIX = "bitacora_";
window.storage = {
  async get(key, shared) {
    // Los valores NO compartidos (shared=false: nombre de quien está usando
    // este celular, si tiene modo Jefe activado, etc.) son puramente locales
    // a este dispositivo — nunca necesitan pedirse a Google Sheets. Antes se
    // pedían igual, sumando peticiones de red innecesarias cada vez que se
    // abría la app.
    if (shared && window.__SHEETS_URL__) {
      try {
        // Si Google Apps Script está "frío" o la red está lenta, no dejamos
        // esperando a la persona para siempre: después de 8s se corta el
        // intento y se usa la copia local guardada en el celular/PC.
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 8000);
        const res = await fetch(`${window.__SHEETS_URL__}?key=${encodeURIComponent(key)}`, {
          signal: controller.signal,
        });
        clearTimeout(timeoutId);
        const data = await res.json();
        if (data && data.ok && data.value !== null && data.value !== undefined) {
          localStorage.setItem(PREFIX + key, data.value);
          return { key, value: data.value, shared: !!shared };
        }
      } catch {
        // sin internet, el Sheet no respondió, o se agotó el tiempo — sigue con la copia local
      }
    }
    try {
      const raw = localStorage.getItem(PREFIX + key);
      return raw ? { key, value: raw, shared: !!shared } : null;
    } catch { return null; }
  },
  // Pide VARIAS claves compartidas en una sola petición HTTP, en vez de una
  // por clave — la pantalla de Inicio de turno, por ejemplo, necesita el
  // programa del día + los inicios/cierres de las 3 áreas para las alertas
  // de insumos al mismo tiempo; antes eso eran 8-10 peticiones en paralelo,
  // cada una lenta por el arranque de Apps Script. Devuelve un mapa
  // { clave: valor|null }; las claves no encontradas caen a lo guardado
  // localmente en este dispositivo.
  async getMany(keys) {
    const out = {};
    if (!keys.length) return out;
    let deSheets = null;
    if (window.__SHEETS_URL__) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 8000);
        const res = await fetch(`${window.__SHEETS_URL__}?keys=${encodeURIComponent(keys.join(","))}`, {
          signal: controller.signal,
        });
        clearTimeout(timeoutId);
        const data = await res.json();
        if (data && data.ok && data.values) deSheets = data.values;
      } catch {
        // sin internet o se agotó el tiempo — sigue con lo guardado localmente
      }
    }
    keys.forEach((key) => {
      const v = deSheets ? deSheets[key] : undefined;
      if (v !== undefined && v !== null) {
        try { localStorage.setItem(PREFIX + key, v); } catch { /* noop */ }
        out[key] = v;
      } else {
        try { out[key] = localStorage.getItem(PREFIX + key); } catch { out[key] = null; }
      }
    });
    return out;
  },
  // Envía un POST al Sheet y espera la confirmación real (hasta 8s). Antes
  // este envío era "dispara y olvida": la app declaraba éxito apenas
  // guardaba en el celular, sin esperar ni revisar si Google Sheets
  // realmente lo recibió. Eso hacía que fallas de sincronización (URL mal
  // puesta, sin internet, Apps Script caído) pasaran completamente
  // desapercibidas — la app decía "guardado" igual. Ahora si el POST falla
  // o Apps Script responde con un error, se avisa explícitamente en vez de
  // fallar en silencio.
  async postAlSheet(body) {
    if (!window.__SHEETS_URL__) return { ok: true }; // sin URL configurada: no hay nada que sincronizar
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 8000);
      const res = await fetch(window.__SHEETS_URL__, {
        method: "POST",
        body: JSON.stringify(body),
        signal: controller.signal,
      });
      clearTimeout(timeoutId);
      if (!res.ok) return { ok: false, error: `Google Sheets respondió con error (${res.status})` };
      const data = await res.json().catch(() => null);
      if (!data || data.ok !== true) return { ok: false, error: (data && data.error) || "Google Sheets no confirmó el guardado" };
      return { ok: true };
    } catch (e) {
      return { ok: false, error: e?.name === "AbortError" ? "Se agotó el tiempo de espera (sin internet o Google Sheets muy lento)" : "No se pudo conectar con Google Sheets" };
    }
  },
  async set(key, value, shared) {
    try {
      localStorage.setItem(PREFIX + key, value);
    } catch { return null; }
    const sync = await this.postAlSheet({ type: "storage", key, value });
    if (!sync.ok) {
      // El guardado local sí funcionó (por eso no se pierde lo escrito),
      // pero Sheets no lo recibió — se avisa para que la persona reintente
      // en vez de creer que ya quedó sincronizado.
      return { key, value, shared: !!shared, syncError: sync.error };
    }
    return { key, value, shared: !!shared };
  },
  async delete(key, shared) {
    try {
      localStorage.removeItem(PREFIX + key);
    } catch { return null; }
    const sync = await this.postAlSheet({ type: "storage", key, value: "" });
    if (!sync.ok) return { key, deleted: true, shared: !!shared, syncError: sync.error };
    return { key, deleted: true, shared: !!shared };
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
