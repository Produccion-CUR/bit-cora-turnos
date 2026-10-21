import React, { useState, useEffect, useMemo, useRef } from "react";
import { createPortal } from "react-dom";
import {
  Droplets, Filter, Package, ClipboardCheck, ClipboardList, BarChart3, Settings, X, Lock,
  Save, ArrowLeft, User, Loader2, CheckCircle2, AlertTriangle, Sun, Moon, Clock, Pencil,
  Trash2, LogOut, CalendarDays, Share2, Boxes, Plus, ChevronRight, BookOpen, ChevronDown,
  Truck, ChevronLeft, FileText, Printer,
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, Legend, CartesianGrid,
} from "recharts";

// ---------------------------------------------------------------------------
// Constantes generales
// ---------------------------------------------------------------------------
const TURNOS = [
  { key: "T3", label: "T3 · Noche" },
  { key: "T1", label: "T1 · Día" },
  { key: "T2", label: "T2 · Tarde" },
];
const TURNO_ICONS = { T1: Sun, T2: Clock, T3: Moon };
const SI_NO = ["Sí", "No"];

// Clave de acceso para el modo "Jefe de producción".
// Es un control básico para esta versión de prueba; al conectar con
// SharePoint conviene reemplazarlo por permisos reales de Power Apps / M365.
const JEFE_PASSWORD = "produccion2026";

const ACCENT = {
  blue: { badge: "bg-blue-500/10 text-blue-700 border-blue-500/30", chart: "#60a5fa" },
  purple: { badge: "bg-purple-500/10 text-purple-700 border-purple-500/30", chart: "#c084fc" },
  amber: { badge: "bg-amber-500/10 text-amber-700 border-amber-500/30", chart: "#fbbf24" },
  emerald: { badge: "bg-emerald-500/10 text-emerald-700 border-emerald-500/30", chart: "#34d399" },
  red: { badge: "bg-red-500/10 text-red-700 border-red-500/30", chart: "#f87171" },
};

// Paleta para gráficos de varias series (Indicadores por día x turno)
const SERIES_COLORS = ["#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899", "#06b6d4"];

const inputBase =
  "w-full rounded-lg border border-slate-400 bg-slate-100 px-3 py-2 text-sm text-slate-900 placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-400 focus:border-blue-400";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const num = (v) => (v === "" || v === undefined || v === null ? 0 : Number(v) || 0);

// Formato de número estilo chileno: punto para miles, coma para decimales.
// Ej: fmtNum(12345.6) -> "12.345,6" · fmtNum(950, 0) -> "950" · fmtNum(1234567) -> "1.234.567"
function fmtNum(v, decimales = 0) {
  const n = typeof v === "number" ? v : num(v);
  return n.toLocaleString("es-CL", {
    minimumFractionDigits: decimales,
    maximumFractionDigits: decimales,
  });
}
const fmtPct = (v) => `${fmtNum((v || 0) * 100, 1)}%`;
const today = () => new Date().toISOString().slice(0, 10);
const fmtFecha = (iso) => {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
// Si por un reintento de guardado quedaron varios registros del mismo turno,
// vale el más reciente (mayor id).
const masRecientePorClave = (lista, clave) =>
  (lista || []).reduce((best, r) => (r.claveTurno === clave && (!best || (r.id || 0) >= (best.id || 0)) ? r : best), null);

const claveTurno = (fecha, turno) => (fecha && turno ? `${fecha}|${turno}` : null);
const fmtDateTime = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
};

// Secuencia de turnos dentro de un mismo día de producción: T3 (Noche) -> T1 (Día)
// -> T2 (Tarde) -> T3 (día siguiente)... Se usa para "arrastrar" datos del
// cierre de un turno hacia el inicio del turno siguiente.
const prevDateISO = (fecha) => {
  const d = new Date(`${fecha}T00:00:00`);
  d.setDate(d.getDate() - 1);
  return d.toISOString().slice(0, 10);
};
const turnoAnterior = (fecha, turno) => {
  if (!fecha || !turno) return null;
  if (turno === "T1") return { fecha, turno: "T3" };
  if (turno === "T2") return { fecha, turno: "T1" };
  if (turno === "T3") return { fecha: prevDateISO(fecha), turno: "T2" };
  return null;
};
const turnoLabel = (turno) => (TURNOS.find((t) => t.key === turno)?.label || turno || "—");

// Secuencia hacia adelante (para armar la grilla semanal del programa de producción)
const nextDateISO = (fecha) => {
  const d = new Date(`${fecha}T00:00:00`);
  d.setDate(d.getDate() + 1);
  return d.toISOString().slice(0, 10);
};
const turnoSiguiente = (fecha, turno) => {
  if (!fecha || !turno) return null;
  if (turno === "T3") return { fecha, turno: "T1" };
  if (turno === "T1") return { fecha, turno: "T2" };
  if (turno === "T2") return { fecha: nextDateISO(fecha), turno: "T3" };
  return null;
};

// ---------------------------------------------------------------------------
// Cobertura real de bodega: bodega NUNCA está disponible en T3, así que el
// turno que solicita insumos debe pedir para sí mismo y, según qué turnos
// estén activos en la planta, también para lo que viene después hasta el
// próximo turno que pueda volver a solicitar.
//
// Reglas (derivadas de cómo opera la planta):
//   - T3 nunca solicita (bodega cerrada) — cubre → []
//   - T1 solicita para sí mismo + el próximo turno activo (sea T2 o T3)
//   - T2 solicita para el próximo turno activo (normalmente T3 del día
//     siguiente) + el turno activo que viene después de ese (para que el
//     turno que reabre bodega tenga con qué partir)
//
// Ejemplos:
//   Turnos activos T1+T2+T3 → T1 cubre [T1, T2] · T2 cubre [T3(mañana), T1(mañana)]
//   Turnos activos T1+T3 (sin T2) → T1 cubre [T1, T3(mañana)] · T3 no solicita
// ---------------------------------------------------------------------------
function siguienteTurnoActivo(fecha, turno, turnosActivosKeys) {
  let s = turnoSiguiente(fecha, turno);
  let intentos = 0;
  // El ciclo T1→T2→T3 solo tiene 3 turnos posibles; si en 10 vueltas no
  // encontramos ninguno activo, es que no hay ningún turno activo en el
  // patrón (ej. el usuario desmarcó todos) — cortamos para no congelar la app.
  while (s && !turnosActivosKeys.includes(s.turno) && intentos < 10) {
    s = turnoSiguiente(s.fecha, s.turno);
    intentos++;
  }
  if (!s || !turnosActivosKeys.includes(s.turno)) return null;
  return s;
}

function turnosACubrir(fecha, turnoActual, turnosActivosKeys) {
  if (!fecha || !turnoActual) return [];
  if (turnoActual === "T3") return []; // bodega nunca disponible en T3

  if (turnoActual === "T1") {
    const propio = { fecha, turno: "T1" };
    const siguiente = siguienteTurnoActivo(fecha, "T1", turnosActivosKeys);
    return siguiente ? [propio, siguiente] : [propio];
  }

  if (turnoActual === "T2") {
    const primero = siguienteTurnoActivo(fecha, "T2", turnosActivosKeys); // normalmente T3(mañana)
    if (!primero) return [];
    const segundo = siguienteTurnoActivo(primero.fecha, primero.turno, turnosActivosKeys); // turno que reabre bodega
    return segundo ? [primero, segundo] : [primero];
  }

  return [];
}
const diaCorto = (fecha) => {
  if (!fecha) return "—";
  const d = new Date(`${fecha}T00:00:00`);
  return d.toLocaleDateString("es-CL", { weekday: "short" }).replace(".", "").toUpperCase();
};
const mondayOfWeek = (fecha) => {
  const d = new Date(`${fecha}T00:00:00`);
  const day = d.getDay(); // 0=domingo ... 6=sábado
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
};
const fmtFechaCorta = (iso) => {
  if (!iso) return "—";
  const [, m, d] = iso.split("-");
  return `${d}-${m}`;
};

// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Google Sheets Web App integration
// Pega aquí la URL que te da Apps Script al implementar:
// ---------------------------------------------------------------------------
const SHEETS_URL = "PEGA_AQUI_TU_URL_DE_APPS_SCRIPT";

// Mapeo de clave de lista → tipo para el Apps Script
const SHEETS_TYPE_MAP = {
  "lavado-inicio-records":    "lavado-inicio",
  "lavado-cierre-records":    "lavado-cierre",
  "seleccion-inicio-records": "seleccion-inicio",
  "seleccion-cierre-records": "seleccion-cierre",
  "envasado-inicio-records":  "envasado-inicio",
  "envasado-cierre-records":  "envasado-cierre",
  "programa-records":         "programa",
};

// Envía UN registro al Sheet (upsert por id).
// No lanza error — si falla, la app sigue funcionando con storage local.
async function saveToSheets(listKey, record) {
  const type = SHEETS_TYPE_MAP[listKey];
  if (!type || !SHEETS_URL || SHEETS_URL.startsWith("PEGA")) return;
  try {
    await fetch(SHEETS_URL, {
      method: "POST",
      body: JSON.stringify({ type, record }),
    });
  } catch {
    // silencioso: el fallback es window.storage
  }
}

// ---------------------------------------------------------------------------
// Almacenamiento (compartido entre todas las personas que usan la app)
// ---------------------------------------------------------------------------
// CACHÉ EN MEMORIA (optimización de velocidad): cada pantalla que se abre
// pedía SIEMPRE los datos de nuevo a Google Sheets (una petición de red por
// cada lista, aunque otra pantalla los hubiera pedido hace 2 segundos). Con
// varias listas por pantalla (Inicio pide 5), navegar se sentía lento.
// Ahora: la primera vez que se pide una key se guarda en esta caché; si se
// vuelve a pedir antes de que venza (SHARED_TTL_MS), se muestra al instante
// sin esperar la red. Si ya venció, igual se muestra lo último conocido de
// inmediato (para que la pantalla no quede "cargando...") mientras se pide
// la versión fresca en segundo plano ("stale-while-revalidate"). Además, si
// dos componentes piden la misma key al mismo tiempo, comparten la misma
// petición en vez de duplicarla.
const SHARED_LIST_CACHE = new Map();   // key -> { value, ts }
const SHARED_LIST_INFLIGHT = new Map(); // key -> Promise en curso
const SHARED_TTL_MS = 20000; // 20 segundos

// Estas dos claves no son filas simples de la hoja "Storage": el servidor
// las arma leyendo/combinando otras pestañas (SKU y Procesos del Sheet), así
// que quedan fuera del pedido agrupado de abajo y siguen pidiéndose una por
// una con window.storage.get.
const CLAVES_NO_AGRUPABLES = new Set(["sku-nuevos", "procesos-extra"]);

// Agrupa en UNA sola petición HTTP todas las claves compartidas que varios
// componentes piden "al mismo tiempo" (mismo ciclo de eventos) — por ejemplo
// al abrir Inicio de turno, que dispara 8-10 useSharedList de golpe (programa
// del día + inicios/cierres de las 3 áreas para las alertas de insumos).
// Antes cada una era su propia petición de red a Apps Script, que es lento
// por naturaleza; ahora se juntan y se piden todas juntas.
let LOTE_PENDIENTE = new Set();
let LOTE_ESPERANDO = new Map(); // key -> [{ resolve, reject }, ...]
let LOTE_TIMER = null;

function despacharLote() {
  const claves = Array.from(LOTE_PENDIENTE);
  const esperando = new Map(LOTE_ESPERANDO);
  LOTE_PENDIENTE = new Set();
  LOTE_ESPERANDO = new Map();
  LOTE_TIMER = null;

  window.storage.getMany(claves).then((valores) => {
    claves.forEach((k) => {
      (esperando.get(k) || []).forEach(({ resolve }) => resolve(valores[k]));
    });
  }).catch((err) => {
    claves.forEach((k) => {
      (esperando.get(k) || []).forEach(({ reject }) => reject(err));
    });
  });
}

function pedirClaveAgrupada(key) {
  return new Promise((resolve, reject) => {
    LOTE_PENDIENTE.add(key);
    if (!LOTE_ESPERANDO.has(key)) LOTE_ESPERANDO.set(key, []);
    LOTE_ESPERANDO.get(key).push({ resolve, reject });
    // setTimeout(…, 0) en vez de una microtarea: junta todas las peticiones
    // que se disparan en la misma pasada de renderizado (varios hooks
    // useSharedList montando a la vez), sin esperar tanto como para que se
    // note una demora extra.
    if (!LOTE_TIMER) LOTE_TIMER = setTimeout(despacharLote, 0);
  });
}

function fetchSharedListKey(key) {
  if (SHARED_LIST_INFLIGHT.has(key)) return SHARED_LIST_INFLIGHT.get(key);
  const p = (async () => {
    try {
      const raw = CLAVES_NO_AGRUPABLES.has(key)
        ? (await window.storage.get(key, true))?.value
        : await pedirClaveAgrupada(key);
      const val = raw ? JSON.parse(raw) : [];
      SHARED_LIST_CACHE.set(key, { value: val, ts: Date.now() });
      return val;
    } catch {
      const prev = SHARED_LIST_CACHE.get(key);
      return prev ? prev.value : [];
    } finally {
      SHARED_LIST_INFLIGHT.delete(key);
    }
  })();
  SHARED_LIST_INFLIGHT.set(key, p);
  return p;
}

function useSharedList(key) {
  const cachedEntry = SHARED_LIST_CACHE.get(key);
  const [items, setItems] = useState(cachedEntry ? cachedEntry.value : []);
  const [loading, setLoading] = useState(!cachedEntry);
  const [error, setError] = useState(null);

  useEffect(() => {
    let active = true;
    const entry = SHARED_LIST_CACHE.get(key);
    const isFresh = entry && (Date.now() - entry.ts) < SHARED_TTL_MS;

    if (entry) {
      // Muestra al instante lo último conocido (sin esperar la red).
      setItems(entry.value);
      setLoading(false);
    }

    if (!isFresh) {
      // No había caché, o venció: pide la versión fresca (en segundo plano
      // si ya había algo que mostrar; de entrada si no había nada).
      fetchSharedListKey(key).then((val) => {
        if (active) { setItems(val); setLoading(false); }
      }).catch(() => {
        if (active) setLoading(false);
      });
    }

    return () => { active = false; };
  }, [key]);

  // Devuelve true/false según si realmente se guardó en el almacenamiento persistente.
  // El estado local (items) se actualiza igual, para que la pantalla refleje el cambio
  // aunque la sincronización falle (y el usuario pueda reintentar sin perder lo escrito).
  const save = async (newItems) => {
    setItems(newItems);
    // Actualiza también la caché compartida para que cualquier otra pantalla
    // que pida esta misma key (aunque sea otro componente) vea el cambio
    // de inmediato, sin tener que esperar a que venza el TTL.
    SHARED_LIST_CACHE.set(key, { value: newItems, ts: Date.now() });

    // Sincroniza con Google Sheets (último item = el que se acaba de agregar/editar)
    if (newItems.length > 0) {
      const lastItem = newItems[newItems.length - 1];
      saveToSheets(key, lastItem); // no-await: no bloqueamos la UI
    }

    try {
      const res = await window.storage.set(key, JSON.stringify(newItems), true);
      if (!res) {
        setError("El almacenamiento no respondió. El cambio quedó en esta sesión pero no se sincronizó: vuelve a presionar Guardar.");
        return false;
      }
      if (res.syncError) {
        // Se guardó en este dispositivo, pero Google Sheets no lo confirmó
        // (antes esto pasaba en silencio — ver window.storage.set en main.jsx).
        setError(`Se guardó en este dispositivo, pero no se pudo sincronizar con Google Sheets (${res.syncError}). Vuelve a presionar Guardar cuando tengas señal.`);
        return false;
      }
      setError(null);
      return true;
    } catch (e) {
      setError(`Error al guardar (${e?.message || "desconocido"}). El cambio quedó en esta sesión pero no se sincronizó: vuelve a presionar Guardar.`);
      return false;
    }
  };

  return [items, save, loading, error];
}

// Misma idea de caché que useSharedList, pero para valores individuales
// (no listas) — por ejemplo el nombre del autor o la última área elegida.
const PERSONAL_VALUE_CACHE = new Map();   // key -> { value, ts }
const PERSONAL_VALUE_INFLIGHT = new Map(); // key -> Promise en curso

function fetchPersonalValueKey(key) {
  if (PERSONAL_VALUE_INFLIGHT.has(key)) return PERSONAL_VALUE_INFLIGHT.get(key);
  const p = (async () => {
    try {
      const res = await window.storage.get(key, false);
      if (res) PERSONAL_VALUE_CACHE.set(key, { value: res.value, ts: Date.now() });
      return res ? res.value : undefined;
    } catch {
      return undefined;
    } finally {
      PERSONAL_VALUE_INFLIGHT.delete(key);
    }
  })();
  PERSONAL_VALUE_INFLIGHT.set(key, p);
  return p;
}

function usePersonalValue(key, fallback) {
  const cachedEntry = PERSONAL_VALUE_CACHE.get(key);
  const [value, setValue] = useState(cachedEntry ? cachedEntry.value : fallback);
  const [loaded, setLoaded] = useState(!!cachedEntry);

  useEffect(() => {
    let active = true;
    const entry = PERSONAL_VALUE_CACHE.get(key);
    const isFresh = entry && (Date.now() - entry.ts) < SHARED_TTL_MS;

    if (entry) {
      setValue(entry.value);
      setLoaded(true);
    }

    if (!isFresh) {
      fetchPersonalValueKey(key).then((val) => {
        if (active && val !== undefined) setValue(val);
        if (active) setLoaded(true);
      }).catch(() => {
        if (active) setLoaded(true);
      });
    }

    return () => { active = false; };
  }, [key]);

  const update = async (v) => {
    setValue(v);
    PERSONAL_VALUE_CACHE.set(key, { value: v, ts: Date.now() });
    try { await window.storage.set(key, v, false); } catch { /* noop */ }
  };

  return [value, update, loaded];
}

// ---------------------------------------------------------------------------
// LISTA DE SUPERVISORES — definida por el Jefe de producción y usada como
// selector en los portales de Selección y Envasado. Compartida entre todos
// los dispositivos. Estructura: { id, nombre, area: "Seleccion"|"Envasado"|"Ambas" }
// ---------------------------------------------------------------------------
function useSupervisores() {
  return useSharedList("supervisores-lista");
}

// ---------------------------------------------------------------------------
// HORARIOS DE TURNO (Jefe de producción) — Lunes a Sábado, hora inicio/fin
// por cada turno. A partir de esto se calculan los "minutos efectivos" de
// producción, descontando 30 min de colación + 15 min de inicio de turno +
// 15 min de término de turno (60 min fijos en total).
// ---------------------------------------------------------------------------
const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const DESCUENTO_FIJO_MIN = 30 + 15 + 15; // colación + inicio + término

const HORARIOS_DEFAULT = Object.fromEntries(
  DIAS_SEMANA.map((dia) => [dia, {
    T1: { horaInicio: "07:00", horaFin: "15:00" },
    T2: { horaInicio: "15:00", horaFin: "23:00" },
    T3: { horaInicio: "23:00", horaFin: "07:00" },
  }])
);

function useHorariosTurno() {
  const [stored, save, loading] = useSharedList("horarios-turno-config");
  const horarios = (stored && stored[0]?.horarios) || HORARIOS_DEFAULT;
  const guardar = async (nuevo) => save([{ horarios: nuevo }]);
  return { horarios, guardar, loading };
}

// Turnos que la planta efectivamente opera (T1/T2/T3 habilitados). Se
// configura una vez desde "Horarios de Turno" y determina qué turnos se
// piden configurar horario en esa misma pantalla.
const TURNOS_HABILITADOS_DEFAULT = { T1: true, T2: true, T3: true };

function useTurnosHabilitados() {
  const [stored, save, loading] = useSharedList("turnos-habilitados-config");
  const habilitados = (stored && stored[0]?.habilitados) || TURNOS_HABILITADOS_DEFAULT;
  const guardar = async (nuevo) => save([{ habilitados: nuevo }]);
  return { habilitados, guardar, loading };
}

// Días de la semana que la planta efectivamente trabaja (Lunes a Domingo).
// Se configura una vez desde "Horarios de Turno" — los días deshabilitados
// no piden configuración de horario.
const DIAS_HABILITADOS_DEFAULT = Object.fromEntries(DIAS_SEMANA.map((d) => [d, d !== "Domingo"]));

function useDiasHabilitados() {
  const [stored, save, loading] = useSharedList("dias-habilitados-config");
  const habilitados = (stored && stored[0]?.habilitados) || DIAS_HABILITADOS_DEFAULT;
  const guardar = async (nuevo) => save([{ habilitados: nuevo }]);
  return { habilitados, guardar, loading };
}

// Nombre del día de la semana (Lunes..Sábado, o Domingo) a partir de una
// fecha ISO "YYYY-MM-DD". Usa mediodía para evitar corrimientos de zona horaria.
function diaSemanaDeFecha(fechaISO) {
  if (!fechaISO) return null;
  const d = new Date(`${fechaISO}T12:00:00`);
  const idx = d.getDay(); // 0=Domingo, 1=Lunes, ... 6=Sábado
  const nombres = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  return nombres[idx];
}

// Minutos brutos entre horaInicio y horaFin (HH:MM), soportando turnos que
// cruzan la medianoche (ej. Noche 23:00 → 07:00).
function minutosBrutosTurno(horaInicio, horaFin) {
  if (!horaInicio || !horaFin) return 0;
  const [h1, m1] = horaInicio.split(":").map(Number);
  const [h2, m2] = horaFin.split(":").map(Number);
  let start = h1 * 60 + m1, end = h2 * 60 + m2;
  if (end <= start) end += 24 * 60;
  return end - start;
}

// Minutos efectivos = minutos brutos - 60 (30 colación + 15 inicio + 15 término).
function minutosEfectivosTurno(horaInicio, horaFin) {
  const bruto = minutosBrutosTurno(horaInicio, horaFin);
  if (bruto <= 0) return 0;
  return Math.max(0, bruto - DESCUENTO_FIJO_MIN);
}

// Dado (fecha, turno) y la config de horarios, devuelve { horaInicio, horaFin,
// minutosBrutos, minutosEfectivos } o null si el turno no tiene horario definido
// para ese día de la semana.
function horarioParaTurno(fecha, turno, horarios) {
  const dia = diaSemanaDeFecha(fecha);
  const cfg = horarios?.[dia]?.[turno];
  if (!cfg || !cfg.horaInicio || !cfg.horaFin) return null;
  return {
    dia,
    horaInicio: cfg.horaInicio,
    horaFin: cfg.horaFin,
    minutosBrutos: minutosBrutosTurno(cfg.horaInicio, cfg.horaFin),
    minutosEfectivos: minutosEfectivosTurno(cfg.horaInicio, cfg.horaFin),
  };
}

// Selector de supervisor: si hay lista usa dropdown + botón "Otro" siempre
// visible al lado (para escribir un nombre que no esté en la lista); si no
// hay lista todavía, permite texto libre directamente.
function SupervisorSelect({ area, value, onChange, supervisoresList }) {
  const filtrados = (supervisoresList || []).filter(
    (s) => s.area === area || s.area === "Ambas"
  );
  const estaEnLista = filtrados.some((s) => s.nombre === value);

  if (filtrados.length === 0) {
    return (
      <div className="space-y-1">
        <input
          className={inputBase}
          placeholder="Nombre del supervisor"
          value={value}
          onChange={(e) => onChange(e.target.value)}
        />
        <p className="text-xs text-slate-400">El Jefe de producción puede agregar supervisores desde su menú.</p>
      </div>
    );
  }

  return (
    <select
      className={inputBase}
      value={estaEnLista ? value : ""}
      onChange={(e) => onChange(e.target.value)}
    >
      <option value="">— Seleccionar supervisor —</option>
      {filtrados.map((s) => (
        <option key={s.id} value={s.nombre}>{s.nombre}</option>
      ))}
    </select>
  );
}

// Trae al "Inicio de turno" los valores de cierre del turno anterior
// (p. ej. pallets pendientes -> pallets sucios, materiales piso planta fin -> inicio).
// Solo completa campos vacíos y solo si no se está editando un registro existente.
const esCampoLista = (campo) => /^mtc(Cajas|Bolsas)Sku/.test(String(campo));

function useCarryOver(areaKey, values, setField, editingId, mapping) {
  const [cierres] = useSharedList(`${areaKey}-cierre-records`);

  // Busca el cierre más reciente anterior a este turno. No basta con mirar
  // solo el turno inmediatamente anterior (T3→T1→T2→T3 del día siguiente):
  // si la planta no trabajó todos los turnos (p. ej. se saltó un T3), el
  // cierre real más reciente puede estar más atrás. Se recorre turno por
  // turno hacia atrás hasta encontrar un cierre guardado (o hasta un límite
  // razonable, ~20 días, para no quedar en un loop si nunca hubo cierres).
  let prev = turnoAnterior(values.fecha, values.turno);
  let prevCierre = null;
  for (let pasos = 0; prev && !prevCierre && pasos < 60; pasos++) {
    const clave = claveTurno(prev.fecha, prev.turno);
    prevCierre = masRecientePorClave(cierres, clave);
    if (!prevCierre) prev = turnoAnterior(prev.fecha, prev.turno);
  }

  useEffect(() => {
    if (editingId) return;
    const allEmpty = mapping.every(([, target]) => values[target] === undefined || values[target] === "" || values[target] === null);
    if (!allEmpty) return;
    if (prevCierre) {
      // Hay cierre del turno anterior: los materiales de piso planta del
      // Inicio parten iguales a los del fin de ese cierre.
      mapping.forEach(([source, target]) => {
        const v = prevCierre[source];
        if (v === undefined || v === "") return;
        // Los campos "por SKU" son listas de filas: solo se copian si realmente
        // son listas (y se copian las filas, no la misma referencia).
        if (esCampoLista(target)) {
          if (Array.isArray(v)) setField(target, v.map((r) => ({ ...r })));
          return;
        }
        setField(target, v);
      });
    } else if (values.fecha && values.turno) {
      // No hubo cierre anterior registrado: los materiales de piso planta
      // parten en 0 (explícito, no en blanco, para no confundir "sin dato"
      // con "en cero" al mirar el registro después).
      // (Los campos "por SKU" son listas: se dejan vacíos, nunca en "0".)
      mapping.forEach(([, target]) => { if (!esCampoLista(target)) setField(target, "0"); });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [values.fecha, values.turno, prevCierre?.id, editingId]);

  return { prevCierre, prev };
}

// Devuelve el stock de materiales de piso de planta MÁS RECIENTE conocido,
// comparando el último "Inicio de turno" registrado con el último "Cierre de
// turno" registrado — usa el que sea más nuevo cronológicamente (claveTurno
// es "fecha|turno", que se puede comparar directamente como texto: mismo
// orden que el tiempo real). Esto hace falta porque antes Insumos y Consumo
// solo miraba el último CIERRE: si alguien acababa de completar un Inicio de
// turno con los materiales de piso planta pero ese turno todavía no se
// cerraba, Insumos seguía mostrando el stock de un cierre anterior (más
// viejo), sin considerar lo recién ingresado.
function ultimoStockMateriales(inicioRecords, cierreRecords, materiales) {
  const masReciente = (lista) =>
    [...lista].sort((a, b) => String(b.claveTurno || "").localeCompare(String(a.claveTurno || "")))[0] || null;
  const ultimoInicio = masReciente(inicioRecords);
  const ultimoCierre = masReciente(cierreRecords);

  let fuente = null, prefijo = "finMat";
  if (ultimoInicio && ultimoCierre) {
    if (String(ultimoInicio.claveTurno) > String(ultimoCierre.claveTurno)) {
      fuente = ultimoInicio; prefijo = "inicioMat";
    } else {
      fuente = ultimoCierre; prefijo = "finMat";
    }
  } else if (ultimoInicio) {
    fuente = ultimoInicio; prefijo = "inicioMat";
  } else if (ultimoCierre) {
    fuente = ultimoCierre; prefijo = "finMat";
  }

  if (!fuente) return {};
  const s = {};
  materiales.forEach((m, i) => { s[m] = num(fuente[`${prefijo}_${i}`]); });
  return s;
}

// Traduce las llaves de "materiales piso planta" (nombres crudos, tal como
// los guarda MaterialesTable: "Film Máquina", "Fixo Azul"...) a los nombres
// que usa la configuración de Insumos y Consumo ("Caja de Film Máquina",
// "Caja de Fixo Azul"...), que son los que se comparan contra "Necesito"
// tanto en Insumos y Consumo como en las Alertas de vencimiento del Inicio.
// Sin esta traducción los nombres nunca calzan entre sí, y ni Insumos y
// Consumo ni las alertas pueden saber cuánto stock real hay de cada insumo.
function stockInsumosSeleccion(stockCrudo) {
  const s = { ...stockCrudo };
  s["Caja de Film Máquina"]        = s["Film Máquina"]           ?? null;
  s["Caja de Film Manual"]         = s["Film Manual"]            ?? null;
  s["Caja de Fixo Azul"]           = s["Fixo Azul"]              ?? null;
  s["Caja de Fixo Café"]           = s["Fixo Café"]              ?? null;
  s["Caja de Fixo Transparente"]   = s["Fixo Transparente"]      ?? null;
  // Bolsas 899 / 1744 y Bolsas Bins se ingresan directamente en su unidad final
  // (Pallets y Paquetes respectivamente) en el formulario de Materiales piso
  // planta, así que aquí solo se traspasan tal cual — sin dividir.
  s["Paquete Bolsas Bins Transparentes"] = s["Bolsas Bins"] ?? null;
  s["Pallet de bolsas 899"]        = s["Bolsas 899"] ?? null;
  s["Pallet de bolsas 1744"]       = s["Bolsas 1744 (Totes)"] ?? null;
  s["Pallet de cajas (MTC)"]       = ((s["Pallet MTC1280"] || 0) + (s["Pallet MTC1310"] || 0)) || null;
  return s;
}

function stockInsumosEnvasado(stockCrudo) {
  const s = { ...stockCrudo };
  s["Caja de Film Máquina"] = s["Film Máquina"] ?? null;
  s["Caja de Film Manual"]  = s["Film Manual"]  ?? null;
  s["Caja de Fixo Azul"]    = s["Fixo Azul"]    ?? null;
  return s;
}

// ---------------------------------------------------------------------------
// Datos maestros (basados en las planillas de la planta)
// ---------------------------------------------------------------------------

// Tipos de bandeja / pallet para Lavado de Bandejas, con el factor
// "bandejas por pallet" usado para calcular Cantidad de Pallets a partir
// de Cantidad de Bandejas. factor = 0 => se ingresa manualmente.
const TIPOS_BANDEJA = [
  { key: "b641", label: "Bandejas 641 (Bandeja Azul)", factor: 230 },
  { key: "b057", label: "Bandejas 057 (Bandeja Verde)", factor: 250 },
  { key: "b055", label: "Bandejas 055 (Bandeja Bajo Peso)", factor: 100 },
  { key: "bmec", label: "Bandeja Mecanizada", factor: 100 },
  { key: "b053", label: "Bandejas 053 (Bandeja 3/4)", factor: 55 },
  { key: "totes", label: "Totes", factor: 80 },
  { key: "bins", label: "Bins", factor: 1 },
];

// Materiales de piso de planta (Selección y Envasado)
const MATERIALES_SELECCION = [
  "Pallet Taco Normal", "Pallet MTC1280", "Pallet MTC1310", "Bolsas 899",
  "Bolsas 1744 (Totes)", "Bolsas Bins", "Fixo Azul", "Fixo Transparente",
  "Fixo Café", "Film Máquina", "Film Manual", "Rollo Cotona",
  "Pallet Tote Armados", "Pallet Cajas Armados 1310", "Pallet Cajas Armados 1280", "Pallet Totes x Armar",
  "Pallet Certificado",
];

const MATERIALES_ENVASADO = ["Pallet Certificado", "Film Máquina", "Film Manual", "Cantidad MTC Cajas", "Cantidad MTC Bolsas", "Pallet Taco Normal"];
// Insumos de Lavado que se rastrean como stock (igual patrón que Selección/
// Envasado vía MaterialesTable + finMat_i), para que Insumos y Consumo y las
// alertas de vencimiento puedan comparar Necesito vs. Tengo también aquí.
const MATERIALES_LAVADO = ["Caja de Film Manual", "Paquete Bolsas Bins Transparentes"];

// Aclaración de unidad de medida que se muestra junto a cada campo en el
// formulario de Materiales piso planta (Selección, Envasado y Lavado). Es
// solo informativo — no cambia la llave con la que se guarda cada material,
// para no romper la compatibilidad con registros ya guardados.
const MATERIAL_UNIDAD = {
  "Bolsas 899": "cantidad de Pallets",
  "Bolsas 1744 (Totes)": "cantidad de Pallets",
  "Bolsas Bins": "Paquete",
  "Fixo Azul": "Cajas",
  "Fixo Transparente": "Cajas",
  "Fixo Café": "Cajas",
  "Film Máquina": "Cajas",
  "Film Manual": "Cajas",
  "Caja de Film Manual": "Cajas",
};

// Dotación
const DOTACION_GENERAL_SELECCION = [
  "Estadístico", "Operador de Máquina", "Jefe Línea Cámara",
  "Movilizadores Cámara", "Supervisor Gestión", "Supervisor Producción",
];
const DOTACION_LINEA = ["Operarios", "Movilizadores", "Jefe de Línea"];
const DOTACION_GENERAL_ENVASADO = ["Estadístico", "Operador de Máquina", "Supervisor Producción"];

// Maestro Línea / Proceso / Especie (para Selección y Envasado)
const LINEA_PROCESOS = {
  L1: [
    ["ARAND ORG RABBITEYE IQF",         "ARANDANO ORG"],
    ["ARAND ORG RABBITEYE BINS",         "ARANDANO ORG"],
    ["ARAND CONV RABBITEYE IQF",         "ARANDANO CONV"],
    ["ARAND CONV RABBITEYE BINS",        "ARANDANO CONV"],
    ["ARAND ORG HIGHBUSH IQF",           "ARANDANO ORG"],
    ["ARAND ORG HIGHBUSH BINS",          "ARANDANO ORG"],
    ["ARAND ORG SVAR IQF",               "ARANDANO ORG"],
    ["ARAND ORG SVAR BINS",              "ARANDANO ORG"],
    ["ARAND CONV HIGHBUSH IQF",          "ARANDANO CONV"],
    ["ARAND CONV HIGHBUSH BINS",         "ARANDANO CONV"],
    ["ARAND CONV SVAR IQF",              "ARANDANO CONV"],
    ["ARAND CONV SVAR BINS",             "ARANDANO CONV"],
    ["ARAND JUGO BAYAS DEL SUR BINS",    "ARANDANO CONV"],
    ["ARAND ORG MIX IQF",                "ARANDANO ORG"],
    ["ARAND CONV MIX IQF",               "ARANDANO CONV"],
    ["ARAND ORG MIX BINS",               "ARANDANO ORG"],
    ["ARAND CONV MIX BINS",              "ARANDANO CONV"],
    ["ARAND ORGANICO IQF 1*15KG SE",     "ARANDANO ORG"],
    ["ARAND CONV IQF 1*15KG SE",         "ARANDANO CONV"],
    ["NANA BIO ORG. BLUEBERRIES 1*10KG", "ARANDANO ORG"],
    ["NANA BIO ORG. BLUEBERRIES 1*30LB", "ARANDANO ORG"],
    ["EIGER BLUEBERRIES 1*15 KG",        "ARANDANO CONV"],
    ["F.O. ALIMENTOS BLUEBERRIES 1*15 KG","ARANDANO CONV"],
    // Frutillas (FRUT = abreviación de Frutilla / Strawberry)
    ["FRUTILLA ORGANICA IQF 1*30LB SE",        "FRUTILLA ORG"],
    ["FRUTILLA CONV SLICED IQF 1*30LB SE",     "FRUTILLA CONV"],
    ["CARAVEL ORGANIC WHOLE STRAWBERRY 1*30 LB","FRUTILLA ORG"],
    ["JONES KENT ORG WHOLE STRAWBERRIES 1*30LB","FRUTILLA ORG"],
    ["PARADIESFRUCHT ORG STRAWBERRIES 1*30LB",  "FRUTILLA ORG"],
    ["FRUT ORG C/AP IQF",               "FRUTILLA ORG"],
    ["FRUT ORG IQF",                    "FRUTILLA ORG"],
    ["FRUT SLICED ORG C/AP IQF",        "FRUTILLA ORG"],
    ["FRUT SLICED ORG IQF",             "FRUTILLA ORG"],
    ["FRUT CONV IQF",                   "FRUTILLA CONV"],
    ["FRUT SLICED CONV IQF",            "FRUTILLA CONV"],
    ["FRUT SABRINA IQF",                "FRUTILLA CONV"],
    ["FRUT JAPON IQF",                  "FRUTILLA CONV"],
    // Mora (Blackberry) — agregados a L1
    ["MORA CONV IQF",                   "MORA CONV"],
    ["MORA NAVAJO IQF",                 "MORA CONV"],
    ["MORA ORG IQF",                    "MORA ORG"],
    ["MORA ORG C/AP IQF",               "MORA ORG"],
    ["MORA JAPON IQF",                  "MORA CONV"],
    ["MORA CULT EXP. PMT S.A 1*30 LB.C/N-B/A", "MORA CONV"],
    ["SUNSHINE BLACKBERRIES 1*30LBS",   "MORA CONV"],
    ["BOYSENBERRY IQF 30 LB",           "MORA CONV"],
  ],
  L3: [
    // Mora (Blackberry)
    ["MORA CONV IQF",                   "MORA CONV"],
    ["MORA NAVAJO IQF",                 "MORA CONV"],
    ["MORA ORG IQF",                    "MORA ORG"],
    ["MORA ORG C/AP IQF",               "MORA ORG"],
    ["MORA JAPON IQF",                  "MORA CONV"],
    ["MORA CULT EXP. PMT S.A 1*30 LB.C/N-B/A", "MORA CONV"],
    ["SUNSHINE BLACKBERRIES 1*30LBS",   "MORA CONV"],
    ["BOYSENBERRY IQF 30 LB",           "MORA CONV"],
    // Uva
    ["UVA TRADICIONAL IQF",             "UVA CONV"],
    // Arándanos
    ["ARAND ORG RABBITEYE IQF",         "ARANDANO ORG"],
    ["ARAND ORG RABBITEYE BINS",        "ARANDANO ORG"],
    ["ARAND CONV RABBITEYE IQF",        "ARANDANO CONV"],
    ["ARAND CONV RABBITEYE BINS",       "ARANDANO CONV"],
    ["ARAND ORG HIGHBUSH IQF",          "ARANDANO ORG"],
    ["ARAND ORG HIGHBUSH BINS",         "ARANDANO ORG"],
    ["ARAND ORG SVAR IQF",              "ARANDANO ORG"],
    ["ARAND ORG SVAR BINS",             "ARANDANO ORG"],
    ["ARAND CONV HIGHBUSH IQF",         "ARANDANO CONV"],
    ["ARAND CONV HIGHBUSH BINS",        "ARANDANO CONV"],
    ["ARAND CONV SVAR IQF",             "ARANDANO CONV"],
    ["ARAND CONV SVAR BINS",            "ARANDANO CONV"],
    ["ARAND JUGO BAYAS DEL SUR BINS",   "ARANDANO CONV"],
    ["ARAND ORG MIX IQF",               "ARANDANO ORG"],
    ["ARAND CONV MIX IQF",              "ARANDANO CONV"],
    ["ARAND ORG MIX BINS",              "ARANDANO ORG"],
    ["ARAND CONV MIX BINS",             "ARANDANO CONV"],
    ["ARAND ORGANICO IQF 1*15KG SE",    "ARANDANO ORG"],
    ["ARAND CONV IQF 1*15KG SE",        "ARANDANO CONV"],
    ["NANA BIO ORG. BLUEBERRIES 1*10KG","ARANDANO ORG"],
    ["NANA BIO ORG. BLUEBERRIES 1*30LB","ARANDANO ORG"],
    ["EIGER BLUEBERRIES 1*15 KG",       "ARANDANO CONV"],
    ["F.O. ALIMENTOS BLUEBERRIES 1*15 KG","ARANDANO CONV"],
  ],
  L4: [
    ["FRAMBUESA CONV IQF",              "FRAMBUESA CONV"],
    ["FRAMBUESA DOLOMIA IQF",           "FRAMBUESA CONV"],
    ["FRAMBUESA ORG IQF",               "FRAMBUESA CONV"],
    ["FRAMBUESA ORG C/AP IQF",          "FRAMBUESA ORG"],
    ["FRAMBUESA JAPON IQF",             "FRAMBUESA CONV"],
    ["F.O. ALIMENTOS RASPBERRIES 1*10 KG","FRAMBUESA CONV"],
    // Mora
    ["MORA CONV IQF",                   "MORA CONV"],
    ["MORA NAVAJO IQF",                 "MORA CONV"],
    ["MORA ORG IQF",                    "MORA ORG"],
    ["MORA ORG C/AP IQF",               "MORA ORG"],
    ["MORA JAPON IQF",                  "MORA CONV"],
    ["SUNSHINE BLACKBERRIES 1*30LBS",   "MORA CONV"],
    ["BOYSENBERRY IQF 30 LB",           "MORA CONV"],
  ],
  LINEA_MANUAL: [
    ["UVA LAVADA",                      "—"],
    ["CEREZA CONV. IQF 1X30LB SE",      "CEREZA CONV"],
    ["ARAND ORGANICO IQF 1*15KG SE",    "ARANDANO ORG"],
    ["ARAND CONV IQF 1*15KG SE",        "ARANDANO CONV"],
    ["FRUTILLA ORGANICA IQF 1*30LB SE", "FRUTILLA ORG"],
    ["FRUTILLA CONV SLICED IQF 1*30LB SE","FRUTILLA CONV"],
    ["DURAZNOS CONV DICE IQF 1*30 LBS SE","DURAZNO CONV"],
    ["INGREDION RASP PURE ORG 1*30 LB", "FRAMBUESA ORG PURE"],
    ["F.O. ALIMENTOS RASPBERRIES 1*10KG","FRAMBUESA CONV"],
    ["LOS NIETITOS RASP CRUMBLE CLEAN 1*30 LB","CRUMBLE"],
    ["SUN-IN RASP CRUMBLE CLEAN 1*30 LB","CRUMBLE"],
    ["SUN-IN RASP CRUMBLE DOLOMIA 4*2,5KG","FRAMBUESA CONV"],
    ["SVZ RASP CRUMBLE SEMI-CLEAN 1*30 LB","CRUMBLE"],
    ["SVZ RASP PURE 1*30 LB",          "FRAMBUESA CONV PURE"],
    ["INFRUIT RASPBERRIES PURE CONV 1*30 LB","FRAMBUESA CONV PURE"],
    ["NANA BIO ORG. BLUEBERRIES 1*10KG","ARANDANO ORG"],
    ["NANA BIO ORG. BLUEBERRIES 1*30LB","ARANDANO ORG"],
    ["EIGER BLUEBERRIES 1*15 KG",       "ARANDANO CONV"],
    ["F.O. ALIMENTOS BLUEBERRIES 1*15 KG.","ARANDANO ORG"],
    ["INFRUIT BLUEBERRIES JUGO 1*15 KG","ARANDANO JUGO"],
    ["BOYSENBERRY IQF 30 LB",          "MORA CONV"],
    ["SUNSHINE BLACKBERRIES 1*30LBS",  "MORA CONV"],
    ["INFRUIT JUGO BLACKBERRIES 1*30 LB","MORA JUGO"],
  ],
  // Línea 6 — nueva línea de Selección. Sin procesos predefinidos todavía:
  // el Jefe los agrega desde "Gestionar especies por línea" según se necesite.
  L6: [],
};

// Líneas de Selección y a qué maestro de procesos apuntan.
// Nota: la planilla no define procesos propios para "Línea 5"; se usa el
// listado de "Línea Manual" como aproximación — ajustar si corresponde.
const LINEAS_SELECCION = [
  { key: "linea1", label: "Línea 1", procesos: LINEA_PROCESOS.L1 },
  { key: "linea3", label: "Línea 3", procesos: LINEA_PROCESOS.L3 },
  { key: "linea4", label: "Línea 4", procesos: LINEA_PROCESOS.L4 },
  { key: "linea5", label: "Línea 5", procesos: LINEA_PROCESOS.LINEA_MANUAL },
  { key: "linea6", label: "Línea 6", procesos: LINEA_PROCESOS.L6 },
];

// Cada línea de Selección admite hasta 5 especies/procesos simultáneos en el
// mismo turno. El slot 1 usa el campo "proceso" (sin sufijo, por compatibilidad
// con registros antiguos); los slots 2-5 usan "proceso2".."proceso5".
const ESPECIE_SLOTS = [1, 2, 3, 4, 5];
function sufijoEspecie(slot) {
  return slot === 1 ? "" : String(slot);
}

// Hasta 3 filas de Kg aprobados por especie/proceso — el nombre de cada fila
// se elige de las especies asociadas a ESA línea (LINEA_PROCESOS + extras
// agregadas por el Jefe en "Gestionar especies por línea"), no de una lista
// genérica. Útil cuando una línea procesa fruta mixta en el mismo turno.
const APROBADO_SLOTS = [1, 2, 3];

// Suma los Kg aprobados de los hasta 3 tipos para una línea+especie(slot). Si
// el registro es antiguo (antes de esta función) y no tiene ningún tipo
// cargado, cae de vuelta al campo simple `kg_${key}_e${slot}_apr` para no
// perder datos históricos.
function kgAprobadoTotal(values, lineaKey, slot) {
  let total = 0;
  let algunTipoCargado = false;
  APROBADO_SLOTS.forEach((t) => {
    const kg = values[`kg_${lineaKey}_e${slot}_apr_t${t}_kg`];
    if (kg !== undefined && kg !== "") { algunTipoCargado = true; total += num(kg); }
  });
  if (!algunTipoCargado) return num(values[`kg_${lineaKey}_e${slot}_apr`]);
  return total;
}

// Suma Kg ingresados/aprobados de una línea across TODOS los slots de especie
// con datos (usado en indicadores, rendimiento y resúmenes de WhatsApp).
function lineaKgTotales(values, linea) {
  let ing = 0, apr = 0;
  ESPECIE_SLOTS.forEach((s) => {
    ing += num(values[`kg_${linea.key}_e${s}_ing`]);
    apr += kgAprobadoTotal(values, linea.key, s);
  });
  return { ing, apr };
}

function especiesUnicas(procesos) {
  return [...new Set(procesos.map(([, especie]) => especie))].filter((e) => e && e !== "—");
}

// ---------------------------------------------------------------------------
// Procesos/especies extra agregados por el Jefe de producción (persistente).
// Permite sumar nuevos pares [Proceso, Especie] a una línea sin tocar código.
// Estructura almacenada: [{ id, lineaKey, proceso, especie }]
// ---------------------------------------------------------------------------
function useProcesosExtra() {
  const [items, save, loading] = useSharedList("procesos-extra");

  const porLinea = (lineaKey) => items.filter((p) => p.lineaKey === lineaKey).map((p) => [p.proceso, p.especie]);

  const agregar = async (lineaKey, proceso, especie) => {
    const entry = { id: Date.now(), lineaKey, proceso: proceso.trim(), especie: especie.trim() };
    return save([...items, entry]);
  };

  const eliminar = async (id) => save(items.filter((p) => p.id !== id));

  return { items, porLinea, agregar, eliminar, loading };
}

// Combina los procesos base (hardcodeados) de una línea con los extra agregados por el Jefe.
function procesosConExtra(baseProcesos, extra) {
  return [...baseProcesos, ...extra];
}

// Maestro de SKU de Envasadora. Desde 20/10/2026 solo se consideran estos 16
// SKU (el resto del maestro original fue eliminado). Campos en blanco para los
// SKU que no traían ese dato.
const SKU_MATERIALES = [
  { sku: "T4-12-1117", producto: "PUBLIX CHERRIES 6*48 OZ", codBolsa: "MTC1253", nomBolsa: "BOL DOYPACK PUBLIX CHERRIES 48OZ. V20", codCaja: "MTCG942", nomCaja: "CJ PUBLIX DARK CHERRIES 6/48OZ.", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-1131", producto: "GV CHERRIES 6*40 OZ", codBolsa: "MTC1762", nomBolsa: "BOL 3 SELLO GV CHERRIES 40OZ", codCaja: "MTC1332", nomCaja: "CJ  GV CHERRIES  6X40OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 117, kgXCaja: 6.803885550000001, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T5-01-0192", producto: "BJS ORG BLUEBERRIES 7*3 LB", codBolsa: "MTC1228", nomBolsa: "BOL WELLSLEY FARMS ORG. BLUEBER 3 LBS", codCaja: "MTC1232", nomCaja: "CJ BLUEBER ORG WELLSLEY 3 LB EXHIBIDORA", bolsasXCaja: 7, cajasXPallet: 63, kgXCaja: 9.52543977, cliente: "Wellsy Farms", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-02-0167", producto: "GREENWISE ORG BLUEBERRIES 6*3 LB", codBolsa: "MTC1713", nomBolsa: "BOL DOYPACK GREENWISE BLUEBER 48OZ V25", codCaja: "MTC1068", nomCaja: "CJ GREENWISE ORGANIC BLUEBERRIES 6X48OZ", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0028", producto: "PUBLIX BLUEBERRIES 6*48 OZ", codBolsa: "MTC1092", nomBolsa: "BOL DOYPACK PUBLIX BLUEBERRIES 48OZ V18", codCaja: "MTCG385", nomCaja: "CJ PUBLIX BLUEBERRIES 6/48OZ.", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0130", producto: "CHIQUITA BLUEBERRIES 8*2,5 LB", codBolsa: "MTC1736", nomBolsa: "BOL CHIQUITA BLUEBER  8/2.5 LB V25", codCaja: "MTCG833", nomCaja: "CJ CHIQUITA BLUEBERRIES 2.5Lb.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 9.071847400000001, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0198", producto: "GV BLUEBERRIES 6*48 OZ", codBolsa: "MTC1763", nomBolsa: "BOL 3 SELLO GV BLUEBERRY 48OZ", codCaja: "MTC1315", nomCaja: "CJ GV BLUEBERRY  6X48OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 8.16466266, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T5-12-0204", producto: "GARDEN FOODS BLUEBERRIES 8*3 LB", codBolsa: "MTC1587", nomBolsa: "BOL GARDEN FOODS BLUEBERIES 48OZ", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: 8, cajasXPallet: "", kgXCaja: 10.88621688, cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0351", producto: "WWS BLUEBERRIES 10*1 KG", codBolsa: "MTC1675", nomBolsa: "BOL WW BLUEBERRIES 1 KG CL V25 REC", codCaja: "MTC1373", nomCaja: "CJ WW BLUEBERRIES 10 X 1 KL. C30KRAFT", bolsasXCaja: 10, cajasXPallet: 100, kgXCaja: 10, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n" },
  { sku: "T5-12-0371", producto: "KNOW & LOVE BLUEBERRIES 6*48 OZ", codBolsa: "MTC1589", nomBolsa: "BOL KNOW&LOVE BLUEBERRIES 48OZ V23", codCaja: "MTC1704", nomCaja: "CJ GENERICA 392X292X170 KRAFT", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Know&Love", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0402", producto: "STORE BRAND BLUEBERRIES 4*48 OZ", codBolsa: "MTC1664", nomBolsa: "BOL GIANT BLUEBERRIES 48OZ V24", codCaja: "MTC1703", nomCaja: "CJ GENERICA 392X292X140 KRAFT", bolsasXCaja: 4, cajasXPallet: 140, kgXCaja: 5.44310844, cliente: "Ahold", tipoPallet: "Yugo", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0371", producto: "KNOW & LOVE WHOLE STRAWBERRIES 6*48 OZ", codBolsa: "MTC1594", nomBolsa: "BOL KNOW&LOVE STRAWBERRIES 48OZ V23", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "Know&Love", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1216", producto: "PUBLIX WHOLE STRAWBERRIES 6*64 OZ", codBolsa: "MTC1102", nomBolsa: "BOL DOYPACK PUBLIX STRAWBERRIES 64OZ V18", codCaja: "MTCG265", nomCaja: "CJ PUBLIX STRAWB 6/64OZ.", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1475", producto: "GV WHOLE STRAWBERRIES 6*64 OZ", codBolsa: "MTC1769", nomBolsa: "BOL 3 SELLO GV WHOLE STRAWB 64OZ", codCaja: "MTC1321", nomCaja: "CJ GV WHOLE STRAWB  6X64OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 70, kgXCaja: 10.88621688, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-12-1476", producto: "GV SLCD STRAWBERRIES 6*64 OZ", codBolsa: "MTC1307", nomBolsa: "BOL 3 SELLO GV SLCD STRAWB 64OZ", codCaja: "MTC1316", nomCaja: "CJ GV SLCD STRAWB  6X64OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 70, kgXCaja: 10.88621688, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-12-0351", producto: "WWS SLCD STRAWBERRIES 10*1 KG", codBolsa: "MTC1742", nomBolsa: "BOL WW STRAWBERRIES 1 KG CL V25 REC", codCaja: "MTC1770", nomCaja: "CJ WW STRAWBERRIES  10 X 1 KL. C30KRAFT", bolsasXCaja: 10, cajasXPallet: 90, kgXCaja: 10, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n" },
];

// Receta por SKU (archivo Receta.xlsx): ingredientes y proporción (1 = 100%) de
// cada kilo de producto terminado. Solo los SKU vigentes del maestro.
const RECETA_SKU = {
  "T5-12-0351": [["ARANDANO CONVENCIONAL IQF", 1]],
  "T5-02-0167": [["ARANDANO ORGANICO IQF", 1]],
  "T5-12-0198": [["ARANDANO CONVENCIONAL IQF", 1]],
  "T5-12-0028": [["ARANDANO CONVENCIONAL IQF", 1]],
  "T4-12-1117": [["CEREZA CONVENCIONAL IQF", 1]],
  "T4-12-1131": [["CEREZA CONVENCIONAL IQF", 1]],
  "TA-12-1475": [["FRUTILLA CONVENCIONAL IQF USA", 1]],
  "T5-01-0192": [["ARANDANO ORGANICO IQF", 1]],
  "T5-12-0130": [["ARANDANO CONVENCIONAL IQF", 1]],
  "T5-12-0371": [["ARANDANO CONVENCIONAL IQF", 1]],
  "T5-12-0204": [["ARANDANO CONVENCIONAL IQF", 1]],
  "TA-12-0371": [["FRUTILLA CONVENCIONAL IQF USA", 1]],
  "TA-12-1476": [["FRUTILLA SLICED CONVENCIONAL IQF USA", 1]],
  "TA-12-1216": [["FRUTILLA CONVENCIONAL IQF USA", 1]],
  "T5-12-0402": [["ARANDANO CONVENCIONAL IQF", 1]],
  "TA-12-0351": [["FRUTILLA SLICED CONVENCIONAL IQF", 1]],
};

// Plantilla de codificación (texto de vencimiento/lote) impresa en bolsa y caja, por cliente.
const CODIFICACION_POR_CLIENTE = {
  "AEON": { bolsa: "", caja: "" },
  "Ahold": { bolsa: "CF 6 173 07 BEST IF USED BY Jun 22\n 2028 inserte hora HH:MM PRODUCT OF CHILE", caja: "BEST IF USED BY Jun\n 22 2028 PRODUCT OF CHILE" },
  "Beisia": { bolsa: "", caja: "" },
  "Bogopa": { bolsa: "", caja: "" },
  "Central Market": { bolsa: "", caja: "" },
  "Chiquita": { bolsa: "", caja: "" },
  "Econo": { bolsa: "", caja: "" },
  "Emart": { bolsa: "", caja: "" },
  "Garden Food": { bolsa: "", caja: "" },
  "Global Growers": { bolsa: "", caja: "" },
  "Goya": { bolsa: "17070 173 6\n BEST BY 06/22/28  \n PRODUCT OF CHILE", caja: "" },
  "Greenwise": { bolsa: "CF 6 173 07 BEST BY Jun 22 2028 (Inserte Hora HH:MM)\n PRODUCT OF CHILE", caja: "BEST BY Jun 22 2028\n PRODUCT OF CHILE" },
  "HEB": { bolsa: "", caja: "" },
  "Infruit": { bolsa: "", caja: "" },
  "Know&Love": { bolsa: "CF 6 173 07 BEST BY Jun 22 2028 (inserte hora HH:MM)\n PRODUCT OF CHILE", caja: "BEST BY Jun 22 2028\n PRODUCT OF CHILE" },
  "Kobbe Busan": { bolsa: "", caja: "" },
  "Macro": { bolsa: "22 Jun 28 \nLOT: CF 6 173 07-(Ingresar Cod. Verificador de Hora) (Ingresar Cod. De País de Fruta de origen)", caja: "BEST BEFORE 22 Jun 28 \nLOT: CF 6 173 07-(Ingresar Cod. Verificador de Hora) (Ingresar Cod. De País de Fruta de origen)" },
  "Makro": { bolsa: "", caja: "" },
  "Publix": { bolsa: "CF 6 173 07 BEST BY Jun 22 2028 (inserte hora HH:MM)\n PRODUCT OF CHILE", caja: "BEST BY Jun 22 2028\n PRODUCT OF CHILE" },
  "SEG": { bolsa: "", caja: "" },
  "Seiyu": { bolsa: "", caja: "" },
  "Trader Joe": { bolsa: "", caja: "" },
  "Walmart": { bolsa: "BEST IF USED BY Jun 22 2028\n 6 173 07-(Ingresar Cod. Verificador de Hora) PRODUCT OF CHILE", caja: "BEST IF USED BY Jun 22 2028\n 6 173 07-(Ingresar Cod. Verificador de Hora)" },
  "Walmart CL": { bolsa: "", caja: "" },
  "Wellsy Farms": { bolsa: "", caja: "" },
  "Woolworth": { bolsa: "22 Jun 28 \nLOT: CF 6 173 07-(Ingresar Cod. Verificador de Hora) (Ingresar Cod. De País de Fruta de origen)", caja: "BEST BEFORE 22 Jun 28 \nLOT: CF 6 173 07-(Ingresar Cod. Verificador de Hora) (Ingresar Cod. De País de Fruta de origen)" },
  "Wooyang": { bolsa: "", caja: "" },
};

// Tabla "Verificador de hora": código (1-12) que se usa en la codificación de bolsa/caja
// para indicar en qué bloque horario de 2 horas se envasó el producto.
const VERIFICADOR_HORA = [
  { codigo: 1, desde: "08:00", hasta: "10:00" },
  { codigo: 2, desde: "10:00", hasta: "12:00" },
  { codigo: 3, desde: "12:00", hasta: "14:00" },
  { codigo: 4, desde: "14:00", hasta: "16:00" },
  { codigo: 5, desde: "16:00", hasta: "18:00" },
  { codigo: 6, desde: "18:00", hasta: "20:00" },
  { codigo: 7, desde: "20:00", hasta: "22:00" },
  { codigo: 8, desde: "22:00", hasta: "00:00" },
  { codigo: 9, desde: "00:00", hasta: "02:00" },
  { codigo: 10, desde: "02:00", hasta: "04:00" },
  { codigo: 11, desde: "04:00", hasta: "06:00" },
  { codigo: 12, desde: "06:00", hasta: "08:00" },
];

// Compatibilidad: lista [sku, producto] derivada del maestro real, usada por
// PROGRAMA_LINEAS y por las búsquedas existentes de nombre de producto.
const ENVASADORA_SKUS = SKU_MATERIALES.map((s) => [s.sku, s.producto]);

// Registro de SKU nuevos (agregados por el Jefe en la pestaña "SKU Envasado"
// del Google Sheet). Se llena una vez al cargar la app — ver App() más abajo.
// skuMaterial() revisa primero el maestro hardcodeado y después este registro.
let SKU_NUEVOS = [];

function skuMaterial(sku) {
  return SKU_MATERIALES.find((s) => s.sku === sku) || SKU_NUEVOS.find((s) => s.sku === sku) || null;
}
// Nota: los overrides del Jefe se aplican dentro de los componentes via useSkuOverrides().getMat(sku).
// La función skuMaterial() sigue apuntando al maestro original para cálculos de Insumos y SKU picker.

// Para entradas del Programa de producción de Envasadora, `especie` guarda el
// CÓDIGO DE SKU; para L1/L3/L4/L5, guarda el CÓDIGO DE PROCESO (no el nombre
// de especie deduplicado) — esto permite detectar reglas que dependen del
// proceso exacto (ej. "BINS", "FRAMBUESA"). Esta función resuelve el nombre
// de especie/producto a mostrar en pantalla, cayendo de vuelta al valor crudo
// si no se encuentra (compatibilidad con registros antiguos).
function especieDisplay(entry) {
  if (!entry || entry.especie === "LAVADO") return entry?.especie || "";
  if (entry.lineaKey === "envasadora") {
    const mat = skuMaterial(entry.especie);
    return mat ? mat.producto : entry.especie;
  }
  const lineaCfg = LINEAS_SELECCION.find((l) => l.key === entry.lineaKey);
  if (lineaCfg) {
    const especie = especieFor(lineaCfg.procesos, entry.especie);
    if (especie) return especie;
  }
  return entry.especie;
}

// En el Programa de producción se debe mostrar el PROCESO (código crudo tal
// como se programó, ej. "ARAND ORG RABBITEYE IQF"), no la especie deducida
// (ej. "ARANDANO ORG"). Para Envasadora no aplica esta distinción — ahí se
// muestra el producto real asociado al SKU, igual que especieDisplay().
function procesoDisplay(entry) {
  if (!entry || entry.especie === "LAVADO") return entry?.especie || "";
  if (entry.lineaKey === "envasadora") return especieDisplay(entry);
  return entry.especie || "—";
}

// ---------------------------------------------------------------------------
// Selector de SKU por código (Envasadora). Permite escribir o pegar el código
// directamente — autocompleta contra el maestro de 275 SKU — y muestra chips
// de acceso rápido con los SKU que ya están programados para esa fecha/turno.
// ---------------------------------------------------------------------------
function SkuPicker({ value, onChange, sugeridos, label = "Código SKU" }) {
  const matActual = value ? skuMaterial(value) : null;
  const [texto, setTexto] = useState(matActual ? `${matActual.sku} — ${matActual.producto}` : "");
  const [abierto, setAbierto] = useState(false);

  useEffect(() => {
    const mat = value ? skuMaterial(value) : null;
    setTexto(mat ? `${mat.sku} — ${mat.producto}` : (value || ""));
  }, [value]);

  // Busca por código O por descripción (nombre del producto), sin importar
  // mayúsculas/minúsculas. Muestra hasta 8 resultados.
  const resultados = useMemo(() => {
    const q = texto.trim().toLowerCase();
    if (!q) return [];
    return [...SKU_MATERIALES, ...SKU_NUEVOS]
      .filter((s) => s.sku.toLowerCase().includes(q) || s.producto.toLowerCase().includes(q))
      .slice(0, 8);
  }, [texto]);

  const elegir = (mat) => {
    onChange(mat.sku);
    setTexto(`${mat.sku} — ${mat.producto}`);
    setAbierto(false);
  };

  const onBlurInput = () => {
    // pequeño retraso para que el click en una opción alcance a registrarse
    // antes de cerrar la lista.
    setTimeout(() => {
      setAbierto(false);
      // Si lo que quedó escrito no coincide con ningún SKU válido, limpia.
      const exact = [...SKU_MATERIALES, ...SKU_NUEVOS].find((s) => s.sku.toLowerCase() === texto.trim().toLowerCase());
      if (exact) { onChange(exact.sku); setTexto(`${exact.sku} — ${exact.producto}`); return; }
      if (!value) onChange("");
    }, 150);
  };

  return (
    <div className="relative">
      {sugeridos && sugeridos.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mb-2">
          {sugeridos.map((sku) => {
            const mat = skuMaterial(sku);
            const sel = value === sku;
            return (
              <button key={sku} type="button" onClick={() => mat && elegir(mat)}
                className={`text-xs rounded-full px-3 py-1 border font-medium transition-colors ${sel ? "bg-amber-500 text-white border-amber-500" : "bg-amber-50 text-amber-800 border-amber-300 hover:bg-amber-100"}`}>
                {sku}{mat ? ` · ${mat.producto}` : ""}
              </button>
            );
          })}
        </div>
      )}
      <label className="block text-xs font-medium text-slate-600 mb-1">{label}</label>
      <input
        className={inputBase}
        placeholder="Escribe el código o el nombre del producto…"
        value={texto}
        onChange={(e) => { setTexto(e.target.value); setAbierto(true); onChange(""); }}
        onFocus={() => setAbierto(true)}
        onBlur={onBlurInput}
      />
      {abierto && resultados.length > 0 && (
        <div className="absolute z-20 mt-1 w-full bg-white border border-slate-300 rounded-xl shadow-lg max-h-64 overflow-y-auto">
          {resultados.map((s) => (
            <button
              key={s.sku}
              type="button"
              onMouseDown={(e) => e.preventDefault()} // evita que el blur se dispare antes del click
              onClick={() => elegir(s)}
              className="w-full text-left px-3 py-2 hover:bg-amber-50 border-b border-slate-100 last:border-0"
            >
              <div className="text-xs font-mono text-amber-700 font-semibold">{s.sku}</div>
              <div className="text-sm text-slate-800 truncate">{s.producto}</div>
            </button>
          ))}
        </div>
      )}
      {abierto && texto.trim() && resultados.length === 0 && (
        <div className="absolute z-20 mt-1 w-full bg-white border border-slate-300 rounded-xl shadow-lg px-3 py-2 text-xs text-slate-400">
          Sin resultados para "{texto}"
        </div>
      )}
      {value && matActual && (
        <p className="text-xs text-emerald-700 mt-1">✓ {matActual.producto}</p>
      )}
      {!value && texto && !abierto && (
        <p className="text-xs text-red-600 mt-1">⚠ Código no encontrado en el maestro de SKU.</p>
      )}
    </div>
  );
}


const CATEGORIAS_COMENTARIO = ["Producción", "Calidad", "Mantención", "Seguridad"];

// Puede haber más de un incidente/accidente en el mismo turno, así que el
// detalle se guarda como una lista (record.incAccList). Esta función además
// da compatibilidad con cierres guardados antes de permitir varios (cuando
// el detalle todavía vivía en 4 campos sueltos, uno por turno), y devuelve un
// ítem vacío cuando se marcó "Sí" pero todavía no se ha cargado ningún
// detalle, para que el formulario y el reporte siempre tengan algo que mostrar.
function incidentesDeCierre(record) {
  const hayEvento = record.huboIncidentes === "Sí" || record.huboAccidentes === "Sí";
  if (!hayEvento) return [];
  if (Array.isArray(record.incAccList) && record.incAccList.length > 0) return record.incAccList;
  if (record.incAccNombre || record.incAccDescripcion || record.incAccDocHecha || record.incAccDocEntregada) {
    return [{
      nombre: record.incAccNombre || "",
      descripcion: record.incAccDescripcion || "",
      docHecha: record.incAccDocHecha || "",
      docEntregada: record.incAccDocEntregada || "",
    }];
  }
  return [{ nombre: "", descripcion: "", docHecha: "", docEntregada: "" }];
}

function buildResumenWhatsapp(area, record, inicio, programaEntries, horarios) {
  // Ojo: NO se muestra area.indicator aquí arriba — para Selección ese
  // indicador se calcula solo con el registro de Cierre (sin fusionar el
  // Inicio), que no trae las líneas activas, así que siempre daba 0% y
  // quedaba redundante con "Rendimiento total"/"Cumplimiento" que ya se
  // muestran, correctamente calculados, dentro de *Producción* más abajo.
  const lineas = [
    `*Bitácora de Turnos – ${area.title}*`,
    `Cierre de turno`,
    `Fecha: ${fmtFecha(record.fecha)} · Turno: ${turnoLabel(record.turno)}`,
    `Responsable: ${record.responsable || "—"}`,
    "",
  ];
  // Orden fijo del mensaje: 1) dotación global (no por línea), 2) datos de
  // producción (kg, rendimiento, cumplimiento), 3) materiales de piso, y
  // recién después los comentarios del cierre e incidentes/accidentes.
  if (area.resumenCompletoCierre) {
    const { dotacion = [], produccion = [], materiales = [] } = area.resumenCompletoCierre(record, inicio, programaEntries, horarios) || {};
    if (dotacion.length > 0) {
      lineas.push("*Dotación:*");
      dotacion.forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
      lineas.push("");
    }
    if (produccion.length > 0) {
      lineas.push("*Producción:*");
      // Entradas especiales: "#" = título de una línea (negrita, con espacio
      // arriba); "##" = subtítulo (proceso/SKU dentro de la línea).
      // "=" = encabezado con viñeta (sus datos van debajo, sin viñeta);
      // tercer valor "plano" = línea sin viñeta.
      produccion.forEach(([label, value, estilo]) => {
        if (label === "#") { lineas.push(""); lineas.push(`*${value}*`); }
        else if (label === "##") lineas.push(`_${value}_`);
        else if (label === "=") lineas.push(`• ${value}`);
        else if (estilo === "plano") lineas.push(`${label}: ${value ?? "—"}`);
        else lineas.push(`• ${label}: ${value ?? "—"}`);
      });
      lineas.push("");
    }
    if (materiales.length > 0) {
      lineas.push("*Materiales de Piso:*");
      materiales.forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
      lineas.push("");
    }
  }
  lineas.push("*Comentarios del cierre:*");
  CATEGORIAS_COMENTARIO.forEach((cat) => {
    const texto = record[`com_${cat}`];
    lineas.push(`• ${cat}: ${texto && texto.trim() ? texto.trim() : "Sin observaciones"}`);
  });
  lineas.push("");
  lineas.push(`¿Incidentes? ${record.huboIncidentes || "—"}`);
  lineas.push(`¿Accidentes? ${record.huboAccidentes || "—"}`);
  const incidentesLista = incidentesDeCierre(record);
  if (incidentesLista.length > 0) {
    lineas.push("");
    lineas.push(incidentesLista.length > 1 ? "*Detalle de los incidentes/accidentes:*" : "*Detalle del incidente/accidente:*");
    incidentesLista.forEach((it, i) => {
      if (incidentesLista.length > 1) lineas.push(`_Caso ${i + 1}:_`);
      lineas.push(`• Persona: ${it.nombre || "—"}`);
      lineas.push(`• Qué pasó: ${it.descripcion || "—"}`);
      lineas.push(`• Documentación hecha: ${it.docHecha || "—"}`);
      lineas.push(`• Documentación entregada: ${it.docEntregada || "—"}`);
    });
  }
  return lineas.join("\n");
}

function buildResumenWhatsappInicio(area, record) {
  const lineas = [
    `*Bitácora de Turnos – ${area.title}*`,
    `Inicio de turno`,
    `Fecha: ${fmtFecha(record.fecha)} · Turno: ${turnoLabel(record.turno)}`,
    `Supervisor: ${record.responsable || "—"}`,
    `Hora de inicio: ${record.horaInicio || "—"}`,
    "",
  ];
  // Mismo orden que el cierre: dotación global, producción, materiales de piso.
  if (area.resumenCompletoInicio) {
    const { dotacion = [], produccion = [], materiales = [] } = area.resumenCompletoInicio(record) || {};
    if (dotacion.length > 0) {
      lineas.push("*Dotación:*");
      dotacion.forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
      lineas.push("");
    }
    if (produccion.length > 0) {
      lineas.push("*Producción:*");
      produccion.forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
      lineas.push("");
    }
    if (materiales.length > 0) {
      lineas.push("*Materiales de Piso:*");
      materiales.forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
      lineas.push("");
    }
  } else {
    area.resumenInicio(record).forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
  }
  return lineas.join("\n").replace(/\n+$/, "");
}

function whatsappShareUrl(text) {
  return `https://wa.me/?text=${encodeURIComponent(text)}`;
}

// Líneas/áreas que el Jefe de producción puede programar: para cada una,
// la lista de especies disponibles y la unidad en que se planifica la cantidad.
const PROGRAMA_LINEAS = [
  { key: "linea1",    label: "Línea 1",      procesos: LINEA_PROCESOS.L1,          especies: ["LAVADO", ...especiesUnicas(LINEA_PROCESOS.L1)],          unidad: "Kg",    color: "rose",    categoriaInsumo: "seleccion" },
  { key: "linea3",    label: "Línea 3",      procesos: LINEA_PROCESOS.L3,          especies: ["LAVADO", ...especiesUnicas(LINEA_PROCESOS.L3)],          unidad: "Kg",    color: "emerald", categoriaInsumo: "seleccion" },
  { key: "linea4",    label: "Línea 4",      procesos: LINEA_PROCESOS.L4,          especies: ["LAVADO", ...especiesUnicas(LINEA_PROCESOS.L4)],          unidad: "Kg",    color: "sky",     categoriaInsumo: "seleccion" },
  { key: "linea5",    label: "Línea 5",      procesos: LINEA_PROCESOS.LINEA_MANUAL, especies: ["LAVADO", ...especiesUnicas(LINEA_PROCESOS.LINEA_MANUAL)], unidad: "Kg",    color: "violet",  categoriaInsumo: "envasado" },
  { key: "linea6",    label: "Línea 6",      procesos: LINEA_PROCESOS.L6,          especies: ["LAVADO", ...especiesUnicas(LINEA_PROCESOS.L6)],          unidad: "Kg",    color: "indigo",  categoriaInsumo: "seleccion" },
  { key: "envasadora",label: "Envasadora",   especies: ["LAVADO", ...ENVASADORA_SKUS.map((s) => s[1])],                                                   unidad: "Cajas", color: "amber",   categoriaInsumo: "envasado" },
];

// Mapa lineaKey -> categoría de Insumos ("seleccion" | "envasado"), derivado
// de PROGRAMA_LINEAS. Usarlo en vez de arrays hardcodeados evita que una
// línea nueva (ej. Línea 6, Línea 7...) quede fuera de los cálculos de
// Insumos y Consumo por simple olvido.
const CATEGORIA_INSUMO_POR_LINEA = Object.fromEntries(
  PROGRAMA_LINEAS.map((l) => [l.key, l.categoriaInsumo])
);

// Devuelve PROGRAMA_LINEAS con las especies/procesos extra (agregados por el
// Jefe vía "Gestionar especies por línea") sumados a la lista de cada línea.
function programaLineasConExtra(porLinea) {
  return PROGRAMA_LINEAS.map((l) => {
    const extra = especiesUnicas(porLinea(l.key));
    const nuevasEspecies = extra.filter((e) => !l.especies.includes(e));
    const procesos = l.procesos ? procesosConExtra(l.procesos, porLinea(l.key)) : l.procesos;
    if (nuevasEspecies.length === 0 && procesos === l.procesos) return l;
    return { ...l, especies: [...l.especies, ...nuevasEspecies], procesos };
  });
}

// Clases estáticas por color (Tailwind no admite construir clases dinámicamente con plantillas).
const PROGRAMA_COLOR_CLASSES = {
  rose:    { meta: "bg-rose-200 text-rose-900",    metaCell: "bg-rose-100 text-rose-800",    label: "bg-rose-100 text-rose-900",    cell: "bg-rose-50" },
  emerald: { meta: "bg-emerald-200 text-emerald-900", metaCell: "bg-emerald-100 text-emerald-800", label: "bg-emerald-100 text-emerald-900", cell: "bg-emerald-50" },
  sky:     { meta: "bg-sky-200 text-sky-900",      metaCell: "bg-sky-100 text-sky-800",      label: "bg-sky-100 text-sky-900",      cell: "bg-sky-50" },
  violet:  { meta: "bg-violet-200 text-violet-900", metaCell: "bg-violet-100 text-violet-800", label: "bg-violet-100 text-violet-900", cell: "bg-violet-50" },
  amber:   { meta: "bg-amber-200 text-amber-900",  metaCell: "bg-amber-100 text-amber-800",  label: "bg-amber-100 text-amber-900",  cell: "bg-amber-50" },
  cyan:    { meta: "bg-cyan-200 text-cyan-900",    metaCell: "bg-cyan-100 text-cyan-800",    label: "bg-cyan-100 text-cyan-900",    cell: "bg-cyan-50" },
  indigo:  { meta: "bg-indigo-200 text-indigo-900",metaCell: "bg-indigo-100 text-indigo-800",label: "bg-indigo-100 text-indigo-900",cell: "bg-indigo-50" },
  slate:   { meta: "bg-slate-200 text-slate-900",  metaCell: "bg-slate-100 text-slate-800",  label: "bg-slate-100 text-slate-900",  cell: "bg-slate-50" },
};

// ---------------------------------------------------------------------------
// Insumos y Consumo: tasas de consumo por línea + especie + insumo
// ---------------------------------------------------------------------------
// Cada línea de PROGRAMA_LINEAS pertenece a un área de materiales:
//   linea1, linea3, linea4 -> Selección (MATERIALES_SELECCION, unidad Kg)
//   linea5, envasadora      -> Envasado  (MATERIALES_ENVASADO, unidad Cajas)
const LINEA_AREA_MATERIALES = {
  linea1: { areaKey: "seleccion", materiales: MATERIALES_SELECCION },
  linea3: { areaKey: "seleccion", materiales: MATERIALES_SELECCION },
  linea4: { areaKey: "seleccion", materiales: MATERIALES_SELECCION },
  linea5: { areaKey: "envasado", materiales: MATERIALES_ENVASADO },
  envasadora: { areaKey: "envasado", materiales: MATERIALES_ENVASADO },
};

// Tasa de consumo = unidades de insumo consumidas por cada 100 unidades producidas
// (100 Kg para líneas de Selección, 100 Cajas para Envasado/Línea 5).
// Estructura: { "lineaKey|especie|insumo": tasaPor100 }
// Se persiste en window.storage bajo la clave "tasas-consumo" (ver useSharedTasas).
// ---------------------------------------------------------------------------
// INSUMOS Y CONSUMO — datos reales desde Excel "Conversion_de_insumos.xlsx"
// ---------------------------------------------------------------------------
// Materiales con tasa variable (dependen de los Kg procesados).
// Unidad de referencia: 1 unidad del formato (ej. 1 Caja de Film).
// tasaBase / base = cuántos formatos se necesitan por `base` unidades producidas.
// 1 Pallet de Selección = 70 cajas × 13,62 Kg ≈ 953,4 Kg. Se usa para convertir
// los Kg programados (Selección y Línea 5) a Pallets, que es la unidad en que
// están definidas las tasas de consumo de Film.
const KG_POR_PALLET_SELECCION = 70 * 13.62;

const INSUMOS_VARIABLE_DEFAULT = {
  seleccion: [
    { nombre: "Caja de Film Máquina", formato: "Caja", uniXFormato: 1, tasaBase: 1, base: 32, unidadBase: "Pallets", nota: "1 rollo de Film Máquina cada 32 pallets" },
    { nombre: "Caja de Film Manual",  formato: "Caja", uniXFormato: 6, tasaBase: 1, base: 10, unidadBase: "Pallets", nota: "6 rollos por caja · 1 caja cada 10 pallets" },
  ],
  lavado: [
    { nombre: "Caja de Film Manual",  formato: "Caja", uniXFormato: 6, tasaBase: 1, base: 10, unidadBase: "Pallets de bandejas", nota: "6 rollos por caja · 1 caja cada 10 pallets de bandejas" },
  ],
  envasado: [
    { nombre: "Caja de Film Máquina", formato: "Caja", uniXFormato: 1, tasaBase: 1, base: 32, unidadBase: "Pallets", nota: "1 rollo de Film Máquina cada 32 pallets" },
    { nombre: "Caja de Film Manual",  formato: "Caja", uniXFormato: 6, tasaBase: 1, base: 10, unidadBase: "Pallets", nota: "6 rollos por caja · 1 caja cada 10 pallets" },
  ],
};

// Materiales con consumo fijo por turno (no dependen de la cantidad producida).
// cantXTurno está en la unidad del `formato` (Caja o Paquete) — NO en unidades
// individuales — para que el redondeo hacia arriba dé directamente la solicitud.
// Ej: Bins transparentes = 6 bolsas/turno ÷ 30 bolsas/paquete = 0,2 paquetes/turno
//     → Math.ceil(0,2 × nTurnos) = unidades de Paquete a solicitar a Bodega.
const INSUMOS_FIJO_DEFAULT = {
  seleccion: [
    { nombre: "Caja de Fixo Azul",        formato: "Caja",    uniXFormato: 6,  cantXTurno: 1,   nota: "6 unidades de Fixo por caja · 1 caja por turno" },
    { nombre: "Caja de Fixo Café",         formato: "Caja",    uniXFormato: 6,  cantXTurno: 2,   nota: "6 unidades de Fixo por caja · 2 cajas por turno" },
    { nombre: "Caja de Fixo Transparente", formato: "Caja",    uniXFormato: 6,  cantXTurno: 1,   nota: "6 unidades de Fixo por caja · 1 caja por turno" },
    { nombre: "Paquete Bolsas Bins Transparentes", formato: "Paquete", uniXFormato: 30, cantXTurno: 6/30, nota: "6 bolsas de bins transparentes por turno (30 bolsas/paquete)" },
  ],
  lavado: [
    { nombre: "Paquete Bolsas Bins Transparentes", formato: "Paquete", uniXFormato: 30, cantXTurno: 6/30, nota: "6 bolsas de bins transparentes por turno (30 bolsas/paquete)" },
  ],
  envasado: [
    { nombre: "Caja de Fixo Azul",        formato: "Caja",    uniXFormato: 6,  cantXTurno: 1,   nota: "6 unidades de Fixo por caja · 1 caja por turno" },
  ],
};

// ---------------------------------------------------------------------------
// Configuración editable de insumos (persistente — la edita el Jefe).
// Se guarda como snapshot completo bajo la clave "insumos-config" (compartida).
// Si no hay nada guardado, se usan los valores por defecto de arriba.
// ---------------------------------------------------------------------------
function useInsumosConfig() {
  const [stored, save, loading, error] = useSharedList("insumos-config");

  const config = (stored && stored[0]) || {
    variable: INSUMOS_VARIABLE_DEFAULT,
    fijo: INSUMOS_FIJO_DEFAULT,
  };

  const updateConfig = async (newConfig) => save([newConfig]);

  const resetDefaults = async () => save([{
    variable: INSUMOS_VARIABLE_DEFAULT,
    fijo: INSUMOS_FIJO_DEFAULT,
  }]);

  return { config, updateConfig, resetDefaults, loading, error };
}

// Función auxiliar: cuántos formatos se necesitan para `totalKg` Kg
function calcFormatos(totalKg, item) {
  if (!totalKg || !item.base) return 0;
  // formatos = totalKg × (tasaBase / base)
  return totalKg * (item.tasaBase / item.base);
}

// ---------------------------------------------------------------------------
// ALERTAS DE VENCIMIENTO — aviso proactivo en la pantalla de Inicio cuando
// algún insumo (Film, Fixo, Bolsas Bins) no va a alcanzar para hoy + mañana,
// comparando lo programado en el Programa de producción contra el stock del
// último cierre registrado. Es un aviso "a grandes rasgos": para el detalle
// exacto de cuánto pedir según el turno, siempre está la pantalla de Insumos.
// ---------------------------------------------------------------------------
// Pedidos de insumos hechos DURANTE el turno (ej. se llamó a Bodega a pedir
// más Fixo antes de que se registre un nuevo Cierre con el stock ya
// actualizado). Se guardan aparte del stock real para que, apenas alguien
// registra el pedido, el aviso de "falta material" de la pantalla de Inicio
// se descuente al instante — sin esperar a que quede reflejado en un Cierre.
function usePedidosInsumos() {
  const [pedidos, save, loading, error] = useSharedList("pedidos-insumos");

  const registrarPedido = async ({ area, nombre, cantidad, formato, autor }) => {
    const record = {
      id: Date.now(),
      area, nombre, formato: formato || "",
      cantidad: num(cantidad),
      fecha: today(),
      hora: horaActualStr(),
      autor: autor || "",
    };
    return save([record, ...pedidos]);
  };

  return { pedidos, registrarPedido, loading, error };
}

function useAlertasInsumos() {
  const [programas]        = useSharedList("programa-records");
  const [seleccionInicios] = useSharedList("seleccion-inicio-records");
  const [seleccionCierres] = useSharedList("seleccion-cierre-records");
  const [envasadoInicios]  = useSharedList("envasado-inicio-records");
  const [envasadoCierres]  = useSharedList("envasado-cierre-records");
  const [lavadoInicios]    = useSharedList("lavado-inicio-records");
  const [lavadoCierres]    = useSharedList("lavado-cierre-records");
  const { config: insumosConfig } = useInsumosConfig();
  const { pedidos } = usePedidosInsumos();

  return useMemo(() => {
    const hoy = today();
    const manana = nextDateISO(hoy);
    const ventana = [hoy, manana];

    // Cuánto se pidió durante el turno, por área + insumo, dentro de la
    // ventana hoy/mañana — se resta directamente de lo que falta.
    const pedidoPor = {};
    pedidos.forEach((p) => {
      if (!ventana.includes(p.fecha)) return;
      const k = `${p.area}||${p.nombre}`;
      pedidoPor[k] = (pedidoPor[k] || 0) + num(p.cantidad);
    });
    const pedidoDe = (area, nombre) => pedidoPor[`${area}||${nombre}`] || 0;

    let kgSeleccion = 0;
    const turnosSel = new Set(), turnosEnv = new Set();
    programas.forEach((p) => {
      if (p.especie === "LAVADO" || !ventana.includes(p.fecha)) return;
      const lk = p.lineaKey || p.linea;
      const categoria = CATEGORIA_INSUMO_POR_LINEA[lk];
      if (categoria === "seleccion") { kgSeleccion += num(p.cantidad); turnosSel.add(p.fecha + p.turno); }
      else if (categoria === "envasado") { turnosEnv.add(p.fecha + p.turno); }
    });
    const nTurnosSel = turnosSel.size;
    const nTurnosEnv = turnosEnv.size;
    const pallets = kgSeleccion / KG_POR_PALLET_SELECCION;

    const stockSel    = stockInsumosSeleccion(ultimoStockMateriales(seleccionInicios, seleccionCierres, MATERIALES_SELECCION));
    const stockEnv     = stockInsumosEnvasado(ultimoStockMateriales(envasadoInicios, envasadoCierres, MATERIALES_ENVASADO));
    const stockLavado  = ultimoStockMateriales(lavadoInicios, lavadoCierres, MATERIALES_LAVADO);
    // Pallets pendientes de lavar (backlog) solo se registra al Cierre, no
    // tiene equivalente en Inicio — se sigue tomando del último cierre.
    const ultimoLavado = [...lavadoCierres].sort((a, b) => b.id - a.id)[0] || {};
    const palletsBandejasLavado = totalPalletsManual(ultimoLavado, "pendientes");

    const alertas = [];

    // Empuja una alerta solo si, después de descontar lo ya pedido durante el
    // turno, todavía falta algo.
    const empujarAlerta = (area, item, necesito, tengo) => {
      const pedido = pedidoDe(area, item.nombre);
      const faltaNeta = necesito - tengo - pedido;
      if (faltaNeta > 1e-9) {
        alertas.push({ area, nombre: item.nombre, falta: faltaNeta, formato: item.formato, pedido });
      }
    };

    if (nTurnosSel > 0) {
      (insumosConfig.variable.seleccion || []).forEach((item) => {
        const necesito = calcFormatos(pallets, item);
        const tengo = stockSel[item.nombre];
        if (tengo !== undefined && necesito > tengo) empujarAlerta("Selección", item, necesito, tengo);
      });
      (insumosConfig.fijo.seleccion || []).forEach((item) => {
        const necesito = item.cantXTurno * nTurnosSel;
        const tengo = stockSel[item.nombre];
        if (tengo !== undefined && necesito > tengo) empujarAlerta("Selección", item, necesito, tengo);
      });
    }

    if (nTurnosEnv > 0) {
      (insumosConfig.fijo.envasado || []).forEach((item) => {
        const necesito = item.cantXTurno * nTurnosEnv;
        const tengo = stockEnv[item.nombre];
        if (tengo !== undefined && necesito > tengo) empujarAlerta("Envasado", item, necesito, tengo);
      });
    }

    // Lavado de bandejas — usa el backlog de pallets pendientes (Variable) y
    // los mismos turnos de Selección como referencia (Fijo), ya que Lavado
    // no tiene su propia cantidad programada en el Programa de producción.
    if (palletsBandejasLavado > 0) {
      (insumosConfig.variable.lavado || []).forEach((item) => {
        const necesito = calcFormatos(palletsBandejasLavado, item);
        const tengo = stockLavado[item.nombre];
        if (tengo !== undefined && necesito > tengo) empujarAlerta("Lavado de bandejas", item, necesito, tengo);
      });
    }
    if (nTurnosSel > 0) {
      (insumosConfig.fijo.lavado || []).forEach((item) => {
        const necesito = item.cantXTurno * nTurnosSel;
        const tengo = stockLavado[item.nombre];
        if (tengo !== undefined && necesito > tengo) empujarAlerta("Lavado de bandejas", item, necesito, tengo);
      });
    }

    return alertas;
  }, [programas, seleccionInicios, seleccionCierres, envasadoInicios, envasadoCierres, lavadoInicios, lavadoCierres, insumosConfig, pedidos]);
}

// ---------------------------------------------------------------------------
// INSUMOS Y CONSUMO (pantalla)
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Mapa areaKey ("seleccion"/"envasado"/"lavado", como usa InsumosConsumoScreen)
// -> el mismo texto de área que usa useAlertasInsumos/usePedidosInsumos, para
// que un pedido registrado en una fila quede etiquetado igual que la alerta
// que intenta descontar.
// ---------------------------------------------------------------------------
const AREA_LABEL_INSUMOS = { seleccion: "Selección", envasado: "Envasado", lavado: "Lavado de bandejas" };

// Formulario chico para registrar que un insumo con falta SÍ se pidió a
// bodega durante el turno — usado dentro de RowInsumo/RowFijo cuando hay
// quiebre. Al guardar, useAlertasInsumos() descuenta esa cantidad de
// inmediato (sin esperar a un nuevo Cierre), así el aviso de faltante se
// actualiza solo.
function RegistrarPedidoForm({ area, nombre, formato, registrarPedido, autor, onDone }) {
  const [cantidad, setCantidad] = useState("");
  const [guardando, setGuardando] = useState(false);
  const [ok, setOk] = useState(false);

  const confirmar = async () => {
    const c = num(cantidad);
    if (!c || c <= 0) return;
    setGuardando(true);
    await registrarPedido({ area, nombre, cantidad: c, formato, autor });
    setGuardando(false);
    setOk(true);
    setCantidad("");
    setTimeout(() => { setOk(false); onDone && onDone(); }, 1200);
  };

  return (
    <div className="flex items-center gap-2 mt-1.5">
      <input
        type="text"
        inputMode="decimal"
        className="w-24 text-sm border border-red-300 rounded-lg px-2 py-1.5 bg-white"
        placeholder={`Cant. ${formato}`}
        value={cantidad}
        onChange={(e) => setCantidad(e.target.value.replace(/[^0-9.,]/g, ""))}
      />
      <button
        onClick={confirmar}
        disabled={guardando || !cantidad}
        className="text-xs font-semibold bg-red-600 hover:bg-red-500 disabled:opacity-50 text-white rounded-lg px-3 py-1.5 transition-colors"
      >
        {ok ? "✓ Registrado" : guardando ? "Guardando…" : "Confirmar pedido"}
      </button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Filas de insumos (módulo nivel — no anidar dentro de otro componente)
// `pedido` = cuánto ya se registró como pedido durante el turno para este
// insumo (proviene de usePedidosInsumos) — se suma al stock antes de decidir
// si falta, así que apenas se registra un pedido el aviso se actualiza solo.
// ---------------------------------------------------------------------------
function RowInsumo({ item, kgTotal, stockMap, area, pedido = 0, registrarPedido, autor }) {
  const [abierto, setAbierto] = useState(false);
  const formatos  = kgTotal > 0 ? calcFormatos(kgTotal, item) : 0;
  const solicitar = Math.ceil(formatos - 1e-9);          // unidades enteras a pedir a bodega
  const tengo     = stockMap ? (stockMap[item.nombre] ?? null) : null;
  const tengoEfectivo = tengo !== null ? tengo + pedido : null;
  const diff      = tengoEfectivo !== null ? tengoEfectivo - formatos : null;
  const falta     = diff !== null && diff < 0;
  const fmt       = item.formato;
  const baseLabel = item.unidadBase || "Kg";
  return (
    <div className={`rounded-xl px-3 py-2.5 border mb-2 last:mb-0 ${falta ? "bg-red-50 border-red-200" : tengo !== null ? "bg-emerald-50 border-emerald-200" : "bg-white border-slate-200"}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1.5 flex-wrap">
            {tengo !== null && (falta
              ? <AlertTriangle size={13} className="text-red-600 shrink-0" />
              : <CheckCircle2 size={13} className="text-emerald-600 shrink-0" />
            )}
            <span className="text-sm font-semibold text-slate-800">{item.nombre}</span>
          </div>
          <div className="text-xs text-slate-500 mt-0.5">{item.nota}</div>
          {kgTotal > 0 && <div className="text-xs text-slate-400">{baseLabel}: <b>{baseLabel === "Kg" ? fmtNum(kgTotal, 0) : fmtNum(kgTotal, 2)}</b></div>}
        </div>
        <div className="text-right text-xs shrink-0 space-y-0.5">
          <div className="text-slate-500">Solicitar:</div>
          <div className="text-base font-bold text-slate-900">{solicitar} {fmt}{solicitar !== 1 ? "s" : ""}</div>
          {item.costoUnitario > 0 && (
            <div className="text-slate-400">≈ ${fmtNum(solicitar * item.costoUnitario, 0)}</div>
          )}
          {tengo !== null ? (
            <>
              <div><span className="text-slate-500">Tengo </span><b>{fmtNum(tengo, 2)}</b></div>
              {pedido > 0 && <div className="text-emerald-600">+ pedido {fmtNum(pedido, 2)}</div>}
              <div className={falta ? "text-red-700 font-bold" : "text-emerald-700 font-semibold"}>
                {falta ? `⚠ Faltan ${fmtNum(Math.abs(diff), 2)}` : `✓ OK`}
              </div>
            </>
          ) : <div className="text-slate-400 italic">sin stock</div>}
        </div>
      </div>
      {formatos > 0 && formatos !== solicitar && (
        <div className="text-xs text-slate-400 mt-1">Cálculo exacto: {fmtNum(formatos, 3)} {fmt}s → se redondea arriba</div>
      )}
      {falta && area && registrarPedido && (
        <div className="mt-2 pt-2 border-t border-red-200">
          {!abierto ? (
            <button onClick={() => setAbierto(true)} className="text-xs font-semibold text-red-700 underline">
              ¿Ya se pidió a bodega? Registrar cantidad pedida
            </button>
          ) : (
            <RegistrarPedidoForm area={area} nombre={item.nombre} formato={fmt} registrarPedido={registrarPedido} autor={autor} onDone={() => setAbierto(false)} />
          )}
        </div>
      )}
    </div>
  );
}

function RowFijo({ item, nTurnos, stockMap, area, pedido = 0, registrarPedido, autor }) {
  const [abierto, setAbierto] = useState(false);
  const necesito  = item.cantXTurno * nTurnos;
  const solicitar = Math.ceil(necesito - 1e-9);
  const fmt       = item.formato;
  const tengo     = stockMap ? (stockMap[item.nombre] ?? null) : null;
  const tengoEfectivo = tengo !== null ? (typeof tengo === "number" ? tengo + pedido : tengo) : null;
  const diff      = typeof tengoEfectivo === "number" ? tengoEfectivo - necesito : null;
  const falta     = diff !== null && diff < 0;
  return (
    <div className={`rounded-xl px-3 py-2.5 border mb-2 last:mb-0 ${falta ? "bg-red-50 border-red-200" : tengo !== null ? "bg-emerald-50 border-emerald-200" : "bg-white border-slate-200"}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1.5">
            {tengo !== null && (falta
              ? <AlertTriangle size={13} className="text-red-600 shrink-0" />
              : <CheckCircle2 size={13} className="text-emerald-600 shrink-0" />
            )}
            <span className="text-sm font-semibold text-slate-800">{item.nombre}</span>
          </div>
          <div className="text-xs text-slate-500 mt-0.5">{item.nota} · {nTurnos} turno{nTurnos !== 1 ? "s" : ""}</div>
        </div>
        <div className="text-right text-xs shrink-0 space-y-0.5">
          <div className="text-slate-500">Solicitar:</div>
          <div className="text-base font-bold text-slate-900">{solicitar} {fmt}{solicitar !== 1 ? "s" : ""}</div>
          {item.costoUnitario > 0 && (
            <div className="text-slate-400">≈ ${fmtNum(solicitar * item.costoUnitario, 0)}</div>
          )}
          {tengo !== null ? (
            <>
              <div><span className="text-slate-500">Tengo </span><b>{typeof tengo === "number" ? fmtNum(tengo, 2) : tengo}</b></div>
              {pedido > 0 && <div className="text-emerald-600">+ pedido {fmtNum(pedido, 2)}</div>}
              <div className={falta ? "text-red-700 font-bold" : "text-emerald-700 font-semibold"}>
                {falta ? `⚠ Faltan ${fmtNum(Math.abs(diff), 2)}` : `✓ OK`}
              </div>
            </>
          ) : <div className="text-slate-400 italic">sin stock</div>}
        </div>
      </div>
      {necesito > 0 && necesito !== solicitar && (
        <div className="text-xs text-slate-400 mt-1">Cálculo exacto: {fmtNum(necesito, 2)} {fmt}s</div>
      )}
      {falta && area && registrarPedido && (
        <div className="mt-2 pt-2 border-t border-red-200">
          {!abierto ? (
            <button onClick={() => setAbierto(true)} className="text-xs font-semibold text-red-700 underline">
              ¿Ya se pidió a bodega? Registrar cantidad pedida
            </button>
          ) : (
            <RegistrarPedidoForm area={area} nombre={item.nombre} formato={fmt} registrarPedido={registrarPedido} autor={autor} onDone={() => setAbierto(false)} />
          )}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Glosario de conversión de insumos — recordatorio fijo de cuántas unidades
// trae cada formato de compra. Se usa el mismo lenguaje en toda la app al
// hablar de "Necesito/Tengo/Faltan".
// ---------------------------------------------------------------------------
const GLOSARIO_INSUMOS = [
  ["Caja de Film Manual",                "6 rollos de Film Manual"],
  ["Caja de Film de Máquina",            "1 rollo de Film de Máquina"],
  ["Caja de Fixo (Azul / Café / Trans.)", "6 unidades de Fixo"],
  ["Paquete Bolsas Bins Transparentes",  "30 bolsas transparentes de bins"],
  ["Paquete Bolsas Bins Azules",         "30 bolsas azules de bins"],
  ["Pallet de cajas (MTC1310 / MTC1280)","600 cajas x pallet"],
  ["Pallet de bolsas 899",               "12.000 bolsas x pallet"],
  ["Pallet de bolsas 1744 (Totes)",      "12.000 bolsas x pallet"],
];

function GlosarioInsumos() {
  const [abierto, setAbierto] = useState(false);
  return (
    <Card>
      <button type="button" onClick={() => setAbierto((v) => !v)} className="w-full flex items-center justify-between text-left">
        <span className="text-sm font-semibold text-slate-800 flex items-center gap-2">
          <BookOpen size={16} className="text-amber-600" /> Glosario de conversión de insumos
        </span>
        <ChevronDown size={16} className={`text-slate-400 shrink-0 transition-transform ${abierto ? "rotate-180" : ""}`} />
      </button>
      {abierto && (
        <div className="mt-3 space-y-1.5">
          {GLOSARIO_INSUMOS.map(([a, b], i) => (
            <div key={i} className="flex justify-between gap-3 text-sm border-b border-slate-100 pb-1.5 last:border-0 last:pb-0">
              <span className="text-slate-700 font-medium">{a}</span>
              <span className="text-slate-500 text-right">{b}</span>
            </div>
          ))}
          <div className="pt-2 mt-1 border-t border-slate-200 space-y-1">
            <p className="text-xs text-slate-500">1 Pallet de Selección = 70 cajas × 13,62 Kg ≈ 953,4 Kg.</p>
            <p className="text-xs text-slate-500">1 Pallet de Envasado varía según la especificación del SKU.</p>
            <p className="text-xs text-slate-500">1 rollo de Film de Máquina alcanza para 32 pallets (Selección y Envasado).</p>
            <p className="text-xs text-slate-500">1 rollo de Film Manual alcanza para 10 pallets (Selección, Envasado y Lavado de bandejas).</p>
          </div>
        </div>
      )}
    </Card>
  );
}

// Sección de insumos fijos por turno — componente de módulo (no anidar dentro
// de InsumosConsumoScreen, o React pierde la identidad del componente en
// cada render y lanza el error #130).
function SeccionFijos({ areaKey, stockMap, titulo, insumosFijo, nTurnos, pedidoDe, registrarPedido, autor }) {
  const items = insumosFijo[areaKey] || [];
  if (!items.length) return null;
  const area = AREA_LABEL_INSUMOS[areaKey];
  return (
    <Card title={`${titulo} — Fijos por turno`}>
      {nTurnos === 0
        ? <EmptyNote text="Selecciona al menos un turno." />
        : items.map((item) => (
            <RowFijo
              key={item.nombre} item={item} nTurnos={nTurnos} stockMap={stockMap}
              area={area} pedido={pedidoDe ? pedidoDe(area, item.nombre) : 0}
              registrarPedido={registrarPedido} autor={autor}
            />
          ))}
    </Card>
  );
}

function InsumosConsumoScreen({ isJefe, onBack, areaFiltro, autor }) {
  const [programas]        = useSharedList("programa-records");
  const [lavadoInicios]    = useSharedList("lavado-inicio-records");
  const [lavadoCierres]    = useSharedList("lavado-cierre-records");
  const [seleccionInicios] = useSharedList("seleccion-inicio-records");
  const [seleccionCierres] = useSharedList("seleccion-cierre-records");
  const [envasadoInicios]  = useSharedList("envasado-inicio-records");
  const [envasadoCierres]  = useSharedList("envasado-cierre-records");
  const { config: insumosConfig } = useInsumosConfig();
  const INSUMOS_VARIABLE = insumosConfig.variable;
  const INSUMOS_FIJO     = insumosConfig.fijo;

  // Pedidos hechos a bodega durante el turno — misma ventana (hoy/mañana) que
  // usa useAlertasInsumos, para que "cuánto ya se pidió" se vea igual acá que
  // en el aviso de la pantalla de Inicio.
  const { pedidos, registrarPedido } = usePedidosInsumos();
  const pedidoDe = useMemo(() => {
    const hoy = today();
    const manana = nextDateISO(hoy);
    const ventana = [hoy, manana];
    const mapa = {};
    pedidos.forEach((p) => {
      if (!ventana.includes(p.fecha)) return;
      const k = `${p.area}||${p.nombre}`;
      mapa[k] = (mapa[k] || 0) + num(p.cantidad);
    });
    return (area, nombre) => mapa[`${area}||${nombre}`] || 0;
  }, [pedidos]);

  const [fecha, setFecha]             = useState(today());
  // Turnos que realmente opera la planta (define el patrón de cobertura de bodega).
  const [turnosSelec, setTurnosSelec] = useState({ T1: true, T2: true, T3: true });
  const turnosActivosPlanta = TURNOS.filter((t) => turnosSelec[t.key]).map((t) => t.key);
  // Turno que está haciendo la solicitud ahora mismo (bodega nunca abre en T3).
  const [turnoSolicitante, setTurnoSolicitante] = useState("T1");

  // (fecha, turno) que hay que cubrir con esta solicitud — nunca se extiende
  // más allá de la fecha elegida + el día siguiente.
  const turnosCubrir = useMemo(
    () => turnosACubrir(fecha, turnoSolicitante, turnosActivosPlanta),
    [fecha, turnoSolicitante, turnosActivosPlanta.join(",")]
  );
  const nTurnos = turnosCubrir.length;

  const titulo = areaFiltro === "seleccion" ? "Insumos · Selección"
    : areaFiltro === "envasado" ? "Insumos · Envasado"
    : "Insumos y Consumo";

  // Fechas con programación desde hoy en adelante (acceso rápido). Antes solo
  // se mostraban hoy y mañana, y se perdían las fechas siguientes.
  const fechasConProg = useMemo(() => {
    const hoy = today();
    const fset = new Set();
    programas.forEach((p) => {
      if (p.especie === "LAVADO") return;
      if (!p.fecha || p.fecha < hoy) return;
      const lk = p.lineaKey || p.linea;
      const categoria = CATEGORIA_INSUMO_POR_LINEA[lk];
      if (!categoria) return; // línea desconocida (no debería pasar, pero por seguridad)
      if (!areaFiltro || categoria === areaFiltro) fset.add(p.fecha);
    });
    return [...fset].sort().slice(0, 14);
  }, [programas, areaFiltro]);

  // ── Stock actual (Inicio o Cierre más reciente de cada área — lo que sea
  // más nuevo, así se refleja un Inicio de turno recién completado aunque
  // ese turno todavía no se haya cerrado) ──────────────────────────────────
  const stockSeleccion = useMemo(() => {
    const s = ultimoStockMateriales(seleccionInicios, seleccionCierres, MATERIALES_SELECCION);
    return Object.keys(s).length === 0 ? s : stockInsumosSeleccion(s);
  }, [seleccionInicios, seleccionCierres]);

  const stockEnvasado = useMemo(() => {
    const s = ultimoStockMateriales(envasadoInicios, envasadoCierres, MATERIALES_ENVASADO);
    return Object.keys(s).length === 0 ? s : stockInsumosEnvasado(s);
  }, [envasadoInicios, envasadoCierres]);

  // Pallets de bandejas pendientes según el último cierre de Lavado registrado
  // (Lavado no se programa con cantidad en el Programa de producción, así que
  // se usa el backlog real reportado al cierre como referencia de "Necesito").
  // Este dato solo se registra al Cierre (no tiene equivalente en Inicio).
  const palletsBandejasLavado = useMemo(() => {
    const last = [...lavadoCierres].sort((a, b) => b.id - a.id)[0];
    if (!last) return 0;
    return totalPalletsManual(last, "pendientes");
  }, [lavadoCierres]);

  // Stock actual de Lavado (Film Manual, Bolsas Bins Transparentes) — Inicio
  // o Cierre más reciente, mismo patrón que Selección/Envasado.
  const stockLavado = useMemo(
    () => ultimoStockMateriales(lavadoInicios, lavadoCierres, MATERIALES_LAVADO),
    [lavadoInicios, lavadoCierres]
  );

  // ── Desglose de Selección y Envasado para la fecha/turnos seleccionados ───
  // Selección: separa Kg de Frambuesa en Línea 4 (usa Tote/Bolsa 1744) del
  // resto (usa Bolsa Caja normal), y detecta si hay Arándano en bins (L1).
  // Envasado: Línea 5 se programa en Kg (se convierte a pallets igual que
  // Selección); Envasadora se programa en Cajas (se convierte a pallets según
  // el cajasXPallet propio de cada SKU, que varía por especificación).
  const desglose = useMemo(() => {
    let kgSeleccionTotal = 0;
    let kgFrambuesaL4 = 0;
    let arandanoBinsCount = 0;
    let kgLinea5 = 0;
    let palletsEnvasadora = 0;
    const cajasPorSku = {};

    turnosCubrir.forEach(({ fecha: f, turno }) => {
      programas
        .filter((p) => p.fecha === f && p.turno === turno && p.especie !== "LAVADO")
        .forEach((p) => {
          const lk = p.lineaKey || p.linea;
          const cant = num(p.cantidad);

          if (["linea1", "linea3", "linea4", "linea6"].includes(lk)) {
            kgSeleccionTotal += cant;
            const especieNombre = (especieDisplay(p) || "").toUpperCase();
            const procesoCrudo = (p.especie || "").toUpperCase();
            const esBins = procesoCrudo.includes("BINS");
            if (lk === "linea4" && especieNombre.includes("FRAMBUESA")) {
              kgFrambuesaL4 += cant;
            }
            if (lk === "linea1" && esBins && especieNombre.includes("ARANDANO")) {
              arandanoBinsCount += 1;
            }
          } else if (lk === "linea5") {
            kgLinea5 += cant;
          } else if (lk === "envasadora") {
            const mat = skuMaterial(p.especie);
            const cajasXPallet = mat ? num(mat.cajasXPallet) : 0;
            if (cajasXPallet > 0) palletsEnvasadora += cant / cajasXPallet;
            cajasPorSku[p.especie] = (cajasPorSku[p.especie] || 0) + cant;
          }
        });
    });

    const palletsSeleccion = kgSeleccionTotal / KG_POR_PALLET_SELECCION;
    const palletsLinea5 = kgLinea5 / KG_POR_PALLET_SELECCION;
    const kgCajaResto = Math.max(kgSeleccionTotal - kgFrambuesaL4, 0);

    return {
      kgSeleccionTotal, kgFrambuesaL4, kgCajaResto, palletsSeleccion,
      arandanoBinsCount,
      kgLinea5, palletsLinea5, palletsEnvasadora,
      palletsEnvasadoTotal: palletsLinea5 + palletsEnvasadora,
      cajasPorSku,
    };
  }, [programas, turnosCubrir]);

  const skuNecesidades = useMemo(() =>
    Object.entries(desglose.cajasPorSku).map(([sku, cajas]) => {
      const mat = skuMaterial(sku);
      const bxc = mat ? num(mat.bolsasXCaja) : 0;
      return {
        sku,
        producto:    mat?.producto || sku,
        cajas,
        bolsas:      bxc > 0 ? Math.ceil(cajas * bxc) : 0,
        codBolsa:    mat?.codBolsa || "",
        nomBolsa:    mat?.nomBolsa || "",
        codCaja:     mat?.codCaja  || "",
        nomCaja:     mat?.nomCaja  || "",
        bolsasXCaja: bxc,
      };
    }),
  [desglose.cajasPorSku]);

  const secSel = areaFiltro === "seleccion" || !areaFiltro;
  const secEnv = areaFiltro === "envasado"  || !areaFiltro;
  const totalCajasSku  = skuNecesidades.reduce((s, x) => s + x.cajas,  0);
  const totalBolsasSku = skuNecesidades.reduce((s, x) => s + x.bolsas, 0);

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title={titulo} subtitle="Necesito vs. tengo en piso de planta" onBack={onBack} icon={Boxes} accent="amber" />
      <div className="px-4 pt-4 space-y-4">

        <GlosarioInsumos />

        {/* Selector de fecha, turno solicitante y patrón de turnos activos */}
        <Card title="¿Quién está pidiendo a bodega?">
          <div className="mb-3">
            <TextField label="Fecha de tu turno" type="date" value={fecha} onChange={setFecha} />
          </div>
          {fechasConProg.length > 0 && (
            <div className="mb-3">
              <p className="text-xs text-slate-500 mb-1.5">Fechas con programación:</p>
              <div className="flex flex-wrap gap-1.5">
                {fechasConProg.map((f) => (
                  <button key={f} onClick={() => setFecha(f)}
                    className={`text-xs rounded-full px-3 py-1 border font-medium transition-colors ${f === fecha ? "bg-amber-500 text-white border-amber-500" : "border-amber-300 text-amber-700 bg-amber-50 hover:bg-amber-100"}`}>
                    {fmtFecha(f)}
                  </button>
                ))}
              </div>
            </div>
          )}

          <label className="block text-xs font-medium text-slate-600 mb-2">¿Qué turno eres tú? (el que solicita ahora)</label>
          <div className="flex gap-2 mb-3">
            {TURNOS.map(({ key }) => {
              const TurnoIcon = TURNO_ICONS[key];
              const esT3 = key === "T3";
              return (
                <button key={key}
                  onClick={() => !esT3 && setTurnoSolicitante(key)}
                  disabled={esT3}
                  title={esT3 ? "Bodega nunca está disponible en T3" : undefined}
                  className={`flex-1 flex flex-col items-center gap-1 rounded-xl py-2.5 border transition-all text-xs font-semibold ${
                    esT3
                      ? "bg-slate-100 text-slate-300 border-slate-200 cursor-not-allowed"
                      : turnoSolicitante === key
                        ? "bg-amber-500 text-white border-amber-500 shadow-sm"
                        : "bg-white text-slate-600 border-slate-300 hover:border-amber-400"
                  }`}
                >
                  {TurnoIcon && <TurnoIcon size={16} />}{key}
                </button>
              );
            })}
          </div>
          <p className="text-xs text-slate-400 mb-3">Bodega nunca está disponible en T3 — por eso T1 o T2 deben pedir con anticipación para cubrirlo.</p>

          <label className="block text-xs font-medium text-slate-600 mb-2">Turnos que opera la planta (patrón habitual)</label>
          <div className="flex gap-2">
            {TURNOS.map(({ key }) => {
              const TurnoIcon = TURNO_ICONS[key];
              return (
                <button key={key} onClick={() => setTurnosSelec((p) => ({ ...p, [key]: !p[key] }))}
                  className={`flex-1 flex flex-col items-center gap-1 rounded-xl py-2 border transition-all text-xs font-semibold ${turnosSelec[key] ? "bg-slate-700 text-white border-slate-700" : "bg-white text-slate-500 border-slate-300 hover:border-slate-500"}`}
                >
                  {TurnoIcon && <TurnoIcon size={14} />}{key}
                </button>
              );
            })}
          </div>

          {turnosCubrir.length > 0 ? (
            <div className="mt-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
              <p className="text-xs font-semibold text-amber-800 mb-1">Esta solicitud debe cubrir:</p>
              <p className="text-xs text-amber-700">
                {turnosCubrir.map((t, i) => `${turnoLabel(t.turno)} (${fmtFecha(t.fecha)})`).join(" + ")}
              </p>
            </div>
          ) : (
            <p className="text-xs text-red-600 mt-3">
              {turnoSolicitante === "T3" ? "T3 no solicita — bodega nunca está disponible en ese turno." : "Sin turnos que cubrir según el patrón activo seleccionado."}
            </p>
          )}
        </Card>

        {nTurnos === 0 && (
          <Card><EmptyNote text="Ajusta el turno solicitante o los turnos activos de la planta para ver los insumos." /></Card>
        )}

        {/* ── SELECCIÓN + LAVADO ── */}
        {secSel && nTurnos > 0 && (
          <>
            <Card title="Selección — Film (variable por pallet)">
              {desglose.palletsSeleccion <= 0
                ? <EmptyNote text="Sin Kg de Selección programados para esta fecha/turnos." />
                : (
                  <>
                    <p className="text-xs text-slate-500 mb-2">
                      {fmtNum(desglose.kgSeleccionTotal, 0)} Kg programados ≈ <b>{fmtNum(desglose.palletsSeleccion, 2)} pallets</b> (953,4 Kg/pallet).
                    </p>
                    {(INSUMOS_VARIABLE.seleccion || []).map((item) => (
                      <RowInsumo
                        key={item.nombre} item={item} kgTotal={desglose.palletsSeleccion} stockMap={stockSeleccion}
                        area="Selección" pedido={pedidoDe("Selección", item.nombre)}
                        registrarPedido={registrarPedido} autor={autor}
                      />
                    ))}
                  </>
                )}
            </Card>

            <Card title="Selección — Pallet de bolsas y Pallet de cajas">
              {desglose.kgSeleccionTotal <= 0 ? (
                <EmptyNote text="Sin Kg de Selección programados." />
              ) : (
                <>
                  <RowInsumo
                    item={{ nombre: "Pallet de cajas (MTC)", formato: "Pallet", uniXFormato: 600, tasaBase: 1, base: 13.62 * 600, unidadBase: "Kg", nota: "1 pallet MTC = 600 cajas de 13,62 Kg · todas las especies excepto Frambuesa L4" }}
                    kgTotal={desglose.kgCajaResto}
                    stockMap={stockSeleccion}
                  />
                  <RowInsumo
                    item={{ nombre: "Pallet de bolsas 899", formato: "Pallet", uniXFormato: 12000, tasaBase: 1, base: 13.62 * 12000, unidadBase: "Kg", nota: "1 pallet = 12.000 bolsas 899 · todas las especies excepto Frambuesa L4" }}
                    kgTotal={desglose.kgCajaResto}
                    stockMap={stockSeleccion}
                  />
                  <RowInsumo
                    item={{ nombre: "Pallet de bolsas 1744", formato: "Pallet", uniXFormato: 12000, tasaBase: 1, base: 9 * 12000, unidadBase: "Kg", nota: "1 pallet = 12.000 bolsas 1744 (Totes) · solo Frambuesa en Línea 4" }}
                    kgTotal={desglose.kgFrambuesaL4}
                    stockMap={stockSeleccion}
                  />
                  {desglose.kgFrambuesaL4 > 0 && (
                    <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mt-1">
                      Frambuesa L4: {fmtNum(desglose.kgFrambuesaL4, 0)} Kg → usa Tote y Bolsa 1744, no Bolsa Caja normal.
                    </p>
                  )}
                </>
              )}
            </Card>

            <SeccionFijos areaKey="seleccion" stockMap={stockSeleccion} titulo="Selección" insumosFijo={INSUMOS_FIJO} nTurnos={nTurnos} pedidoDe={pedidoDe} registrarPedido={registrarPedido} autor={autor} />

            <Card title="Selección — Paquete Bolsas Bins Azules (condicional)">
              <RowFijo
                item={{ nombre: "Paquete Bolsas Bins Azules", formato: "Paquete", uniXFormato: 30, cantXTurno: desglose.arandanoBinsCount > 0 ? 1/30 : 0, nota: desglose.arandanoBinsCount > 0 ? "Arándano en bins en L1 — 1 bolsa azul por turno (30 bolsas/paquete)" : "Sin Arándano en bins programado en esta fecha/turnos" }}
                nTurnos={nTurnos || 1}
                stockMap={{}}
              />
            </Card>

            <Card title="Lavado de bandejas — Film Manual">
              <p className="text-xs text-slate-500 mb-2">
                Pallets de bandejas pendientes (según el último cierre de Lavado registrado): <b>{palletsBandejasLavado}</b>
              </p>
              {(INSUMOS_VARIABLE.lavado || []).map((item) => (
                <RowInsumo
                  key={item.nombre} item={item} kgTotal={palletsBandejasLavado} stockMap={stockLavado}
                  area="Lavado de bandejas" pedido={pedidoDe("Lavado de bandejas", item.nombre)}
                  registrarPedido={registrarPedido} autor={autor}
                />
              ))}
            </Card>
            <SeccionFijos areaKey="lavado" stockMap={stockLavado} titulo="Lavado de bandejas" insumosFijo={INSUMOS_FIJO} nTurnos={nTurnos} pedidoDe={pedidoDe} registrarPedido={registrarPedido} autor={autor} />
          </>
        )}

        {/* ── ENVASADO ── */}
        {secEnv && nTurnos > 0 && (
          <>
            <Card title="Envasado — Film (variable por pallet)">
              {desglose.palletsEnvasadoTotal <= 0 ? (
                <EmptyNote text="Sin Línea 5 ni Envasadora programadas para esta fecha/turnos." />
              ) : (
                <>
                  <p className="text-xs text-slate-500 mb-2">
                    {desglose.kgLinea5 > 0 && <>Línea 5: {fmtNum(desglose.kgLinea5, 0)} Kg ≈ {fmtNum(desglose.palletsLinea5, 2)} pallets. </>}
                    {desglose.palletsEnvasadora > 0 && <>Envasadora: {fmtNum(desglose.palletsEnvasadora, 2)} pallets (según cajasXPallet de cada SKU). </>}
                    Total: <b>{fmtNum(desglose.palletsEnvasadoTotal, 2)} pallets</b>.
                  </p>
                  {(INSUMOS_VARIABLE.envasado || []).map((item) => (
                    <RowInsumo
                      key={item.nombre} item={item} kgTotal={desglose.palletsEnvasadoTotal} stockMap={stockEnvasado}
                      area="Envasado" pedido={pedidoDe("Envasado", item.nombre)}
                      registrarPedido={registrarPedido} autor={autor}
                    />
                  ))}
                </>
              )}
            </Card>

            <SeccionFijos areaKey="envasado" stockMap={stockEnvasado} titulo="Envasado" insumosFijo={INSUMOS_FIJO} nTurnos={nTurnos} pedidoDe={pedidoDe} registrarPedido={registrarPedido} autor={autor} />

            <Card title="Envasado — Bolsas y Cajas por SKU">
              {skuNecesidades.length === 0 ? (
                <div className="space-y-2">
                  <EmptyNote text={`Sin Envasadora programada para ${fmtFecha(fecha)} en los turnos seleccionados.`} />
                  {fechasConProg.length > 0 && (
                    <div>
                      <p className="text-xs text-slate-500 mb-1.5">Fechas con programación:</p>
                      <div className="flex flex-wrap gap-1.5">
                        {fechasConProg.map((f) => (
                          <button key={f} onClick={() => setFecha(f)}
                            className="text-xs rounded-full px-3 py-1 border border-amber-300 text-amber-700 bg-amber-50 hover:bg-amber-100 font-medium">
                            {fmtFecha(f)}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                <div className="space-y-3">
                  {skuNecesidades.map((s) => (
                    <div key={s.sku} className="border border-amber-200 bg-amber-50/40 rounded-xl px-3 py-3">
                      <div className="text-sm font-bold text-slate-900 leading-tight">{s.producto}</div>
                      <div className="text-xs text-slate-400 mb-2">{s.sku}</div>
                      <div className="grid grid-cols-2 gap-2">
                        <div className="bg-white border border-amber-200 rounded-lg p-2.5">
                          <div className="text-xs text-slate-500 mb-0.5">Cajas a producir</div>
                          <div className="text-2xl font-bold text-amber-700 leading-none">{fmtNum(s.cajas)}</div>
                          {s.codCaja && <div className="text-xs text-slate-400 mt-1 truncate">{s.codCaja}</div>}
                        </div>
                        <div className="bg-white border border-blue-200 rounded-lg p-2.5">
                          <div className="text-xs text-slate-500 mb-0.5">Bolsas necesarias</div>
                          <div className="text-2xl font-bold text-blue-700 leading-none">{fmtNum(s.bolsas)}</div>
                          <div className="text-xs text-slate-400">{s.bolsasXCaja} bolsas/caja</div>
                          {s.codBolsa && <div className="text-xs text-slate-400 truncate">{s.codBolsa}</div>}
                        </div>
                      </div>
                    </div>
                  ))}
                  <div className="grid grid-cols-2 gap-2 border-t border-amber-200 pt-3 text-center">
                    <div>
                      <div className="text-xs text-slate-500">Total cajas</div>
                      <div className="text-xl font-bold text-amber-700">{fmtNum(totalCajasSku)}</div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500">Total bolsas</div>
                      <div className="text-xl font-bold text-blue-700">{fmtNum(totalBolsasSku)}</div>
                    </div>
                  </div>
                </div>
              )}
            </Card>
          </>
        )}

      </div>
    </div>
  );
}


// ---------------------------------------------------------------------------
// CONFIGURAR INSUMOS (Jefe de producción)
// Permite editar las tasas de consumo (variables por Kg y fijas por turno)
// sin tocar código. Los cambios se guardan compartidos y los usa de inmediato
// la pantalla de Insumos y Consumo.
// ---------------------------------------------------------------------------
const AREAS_INSUMOS = [
  { key: "seleccion", label: "Selección" },
  { key: "lavado",    label: "Lavado de bandejas" },
  { key: "envasado",  label: "Envasado" },
];

function vacioVariable() {
  return { nombre: "", formato: "Caja", uniXFormato: 1, tasaBase: 1, base: 100, unidadBase: "Pallets", costoUnitario: 0, nota: "" };
}
function vacioFijo() {
  return { nombre: "", formato: "Caja", uniXFormato: 1, cantXTurno: 1, costoUnitario: 0, nota: "", unidadMostrar: "" };
}

function InsumosConfigScreen({ onBack }) {
  const { config, updateConfig, resetDefaults, loading } = useInsumosConfig();
  const [areaSel, setAreaSel] = useState("seleccion");
  const [tipo, setTipo] = useState("variable"); // variable | fijo
  const [borrador, setBorrador] = useState(null); // copia editable de config mientras no se guarda
  const [toast, setToast] = useState(null);

  // Usa el borrador si existe (cambios sin guardar), si no usa la config persistida
  const activo = borrador || config;
  const lista = (tipo === "variable" ? activo.variable : activo.fijo)[areaSel] || [];

  const asegurarBorrador = () => {
    if (!borrador) setBorrador(JSON.parse(JSON.stringify(config)));
    return borrador || JSON.parse(JSON.stringify(config));
  };

  const actualizarCampo = (idx, campo, valor) => {
    const next = asegurarBorrador();
    const clone = JSON.parse(JSON.stringify(next));
    const arr = (tipo === "variable" ? clone.variable : clone.fijo)[areaSel];
    arr[idx] = { ...arr[idx], [campo]: ["uniXFormato","tasaBase","base","cantXTurno","costoUnitario"].includes(campo) ? num(valor) : valor };
    setBorrador(clone);
  };

  const agregarItem = () => {
    const next = asegurarBorrador();
    const clone = JSON.parse(JSON.stringify(next));
    const arr = (tipo === "variable" ? clone.variable : clone.fijo);
    arr[areaSel] = [...(arr[areaSel] || []), tipo === "variable" ? vacioVariable() : vacioFijo()];
    setBorrador(clone);
  };

  const eliminarItem = (idx) => {
    const next = asegurarBorrador();
    const clone = JSON.parse(JSON.stringify(next));
    const arr = (tipo === "variable" ? clone.variable : clone.fijo);
    arr[areaSel] = arr[areaSel].filter((_, i) => i !== idx);
    setBorrador(clone);
  };

  const guardar = async () => {
    const aGuardar = borrador || config;
    const ok = await updateConfig(aGuardar);
    setToast({ kind: ok ? "ok" : "error", message: ok ? "Configuración de insumos guardada." : "No se pudo guardar. Intenta nuevamente." });
    if (ok) setBorrador(null);
    setTimeout(() => setToast(null), 3000);
  };

  const restaurar = async () => {
    const ok = await resetDefaults();
    setBorrador(null);
    setToast({ kind: ok ? "ok" : "error", message: ok ? "Se restauraron los valores por defecto." : "No se pudo restaurar." });
    setTimeout(() => setToast(null), 3000);
  };

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Configurar Insumos" subtitle="Tasas de consumo — Jefe de producción" onBack={onBack} icon={Boxes} accent="amber" />

      <div className="px-4 pt-4 space-y-4">
        <Card title="Área y tipo de insumo">
          <div className="flex gap-1.5 mb-3 flex-wrap">
            {AREAS_INSUMOS.map((a) => (
              <button key={a.key} onClick={() => setAreaSel(a.key)}
                className={`text-xs rounded-full px-3 py-1.5 border font-medium transition-colors ${areaSel === a.key ? "bg-slate-800 text-white border-slate-800" : "border-slate-300 text-slate-600 hover:border-slate-500"}`}>
                {a.label}
              </button>
            ))}
          </div>
          <div className="flex gap-2">
            <button onClick={() => setTipo("variable")}
              className={`flex-1 text-xs font-semibold rounded-xl py-2.5 border transition-colors ${tipo === "variable" ? "bg-amber-500 text-white border-amber-500" : "bg-white text-slate-600 border-slate-300"}`}>
              Variable por Pallet
            </button>
            <button onClick={() => setTipo("fijo")}
              className={`flex-1 text-xs font-semibold rounded-xl py-2.5 border transition-colors ${tipo === "fijo" ? "bg-amber-500 text-white border-amber-500" : "bg-white text-slate-600 border-slate-300"}`}>
              Fijo por turno
            </button>
          </div>
        </Card>

        <p className="text-xs text-slate-400 px-1">
          Bolsas 899 / 1744, Bolsa Azul de Bins y otros insumos que dependen de la especie programada (ej. Frambuesa en Línea 4, Arándano en bins) tienen reglas especiales y no se editan aquí — se calculan automáticamente en la pantalla de Insumos y Consumo.
        </p>

        {toast && <Toast {...toast} />}

        {loading ? <Loader /> : (
          <Card title={`${AREAS_INSUMOS.find((a) => a.key === areaSel)?.label} — ${tipo === "variable" ? "Variables por Kg" : "Fijos por turno"}`}>
            {lista.length === 0 && <EmptyNote text="No hay insumos configurados en esta categoría todavía." />}
            <div className="space-y-3">
              {lista.map((item, idx) => (
                <div key={idx} className="border border-slate-200 rounded-xl p-3 space-y-2 bg-slate-50">
                  <div className="flex items-center gap-2">
                    <input
                      className="flex-1 rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm font-semibold"
                      placeholder="Nombre del insumo"
                      value={item.nombre}
                      onChange={(e) => actualizarCampo(idx, "nombre", e.target.value)}
                    />
                    <button onClick={() => eliminarItem(idx)} aria-label="Eliminar" className="p-2 rounded-lg hover:bg-red-100 text-red-500 shrink-0">
                      <Trash2 size={16} />
                    </button>
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-xs text-slate-500 mb-1">Formato</label>
                      <input
                        className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                        placeholder="Caja, Paquete…"
                        value={item.formato}
                        onChange={(e) => actualizarCampo(idx, "formato", e.target.value)}
                      />
                    </div>
                    <div>
                      <label className="block text-xs text-slate-500 mb-1">Unidades x formato</label>
                      <input
                        type="number"
                        className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                        value={item.uniXFormato}
                        onChange={(e) => actualizarCampo(idx, "uniXFormato", e.target.value)}
                      />
                    </div>
                  </div>

                  {tipo === "variable" ? (
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="block text-xs text-slate-500 mb-1">Formatos necesarios</label>
                        <input
                          type="number"
                          className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                          value={item.tasaBase}
                          onChange={(e) => actualizarCampo(idx, "tasaBase", e.target.value)}
                        />
                      </div>
                      <div>
                        <label className="block text-xs text-slate-500 mb-1">Por cada</label>
                        <input
                          type="number"
                          className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                          value={item.base}
                          onChange={(e) => actualizarCampo(idx, "base", e.target.value)}
                        />
                      </div>
                      <div className="col-span-2">
                        <label className="block text-xs text-slate-500 mb-1">Unidad (Pallets, Kg, etc.)</label>
                        <input
                          className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                          placeholder="Pallets"
                          value={item.unidadBase || ""}
                          onChange={(e) => actualizarCampo(idx, "unidadBase", e.target.value)}
                        />
                      </div>
                    </div>
                  ) : (
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="block text-xs text-slate-500 mb-1">Cantidad x turno</label>
                        <input
                          type="number"
                          className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                          value={item.cantXTurno}
                          onChange={(e) => actualizarCampo(idx, "cantXTurno", e.target.value)}
                        />
                      </div>
                      <div>
                        <label className="block text-xs text-slate-500 mb-1">Unidad a mostrar (opcional)</label>
                        <input
                          className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                          placeholder="bolsas, cajas…"
                          value={item.unidadMostrar || ""}
                          onChange={(e) => actualizarCampo(idx, "unidadMostrar", e.target.value)}
                        />
                      </div>
                    </div>
                  )}

                  <div>
                    <label className="block text-xs text-slate-500 mb-1">Costo por {item.formato || "unidad"} ($)</label>
                    <input
                      type="number"
                      className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                      placeholder="0"
                      value={item.costoUnitario || ""}
                      onChange={(e) => actualizarCampo(idx, "costoUnitario", e.target.value)}
                    />
                  </div>

                  <div>
                    <label className="block text-xs text-slate-500 mb-1">Nota / explicación</label>
                    <input
                      className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-sm"
                      placeholder="Ej: 1 rollo cada 6 pallets de 1.100 Kg"
                      value={item.nota}
                      onChange={(e) => actualizarCampo(idx, "nota", e.target.value)}
                    />
                  </div>
                </div>
              ))}
            </div>

            <button onClick={agregarItem}
              className="w-full mt-3 flex items-center justify-center gap-2 border-2 border-dashed border-slate-300 hover:border-amber-400 text-slate-500 hover:text-amber-600 font-semibold rounded-xl py-2.5 transition-colors text-sm">
              <Plus size={16} /> Agregar insumo
            </button>
          </Card>
        )}

        <p className="text-xs text-slate-400 px-1">
          El costo se usa para estimar el gasto en insumos por turno. Si no conoces el precio exacto, déjalo en 0 — no afecta el cálculo de cantidades.
        </p>

        <div className="flex gap-2">
          <button onClick={guardar}
            className="flex-1 flex items-center justify-center gap-2 bg-amber-500 hover:bg-amber-400 text-white font-semibold rounded-xl py-3 transition-colors text-sm">
            <Save size={16} /> Guardar cambios
          </button>
          <button onClick={restaurar}
            className="px-4 rounded-xl border border-slate-300 text-slate-500 hover:border-red-300 hover:text-red-600 text-xs font-medium transition-colors">
            Restaurar
          </button>
        </div>
        <p className="text-xs text-slate-400 text-center">
          "Formatos necesarios" / "Por cada" + "Unidad" define la tasa: ej. 1 rollo cada 32 Pallets.
        </p>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// CALCULADORA DE COSTOS (Jefe de producción)
// Herramienta de estimación rápida: el Jefe ingresa Kg a procesar (Selección)
// y/o turnos a cubrir, y la calculadora usa las mismas tasas y costos
// unitarios configurados en "Configurar Insumos" para estimar el gasto total.
// No guarda nada — es solo para simular escenarios ("¿cuánto me costaría
// procesar 5.000 Kg en 2 turnos?").
// ---------------------------------------------------------------------------
function CalculadoraCostosScreen({ onBack }) {
  const { config, loading } = useInsumosConfig();
  const [area, setArea] = useState("seleccion");
  const [kg, setKg] = useState("");
  const [nTurnos, setNTurnos] = useState("1");

  const kgNum = num(kg);
  const turnosNum = Math.max(1, num(nTurnos) || 1);
  const pallets = area === "seleccion" || area === "lavado" ? kgNum / KG_POR_PALLET_SELECCION : 0;

  const detalle = useMemo(() => {
    const items = [];
    (config.variable[area] || []).forEach((item) => {
      if (!item.costoUnitario) return;
      const formatos = pallets > 0 ? calcFormatos(pallets, item) : 0;
      const solicitar = Math.ceil(formatos - 1e-9);
      if (solicitar > 0) items.push({ nombre: item.nombre, cantidad: solicitar, formato: item.formato, costo: solicitar * item.costoUnitario });
    });
    (config.fijo[area] || []).forEach((item) => {
      if (!item.costoUnitario) return;
      const necesito = item.cantXTurno * turnosNum;
      const solicitar = Math.ceil(necesito - 1e-9);
      if (solicitar > 0) items.push({ nombre: item.nombre, cantidad: solicitar, formato: item.formato, costo: solicitar * item.costoUnitario });
    });
    return items;
  }, [config, area, pallets, turnosNum]);

  const totalEstimado = detalle.reduce((s, d) => s + d.costo, 0);
  const sinCostos = detalle.length === 0 && (kgNum > 0 || turnosNum > 0);

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Calculadora de Costos" subtitle="Estimación rápida de gasto en insumos" onBack={onBack} icon={CheckCircle2} accent="emerald" />
      <div className="px-4 pt-4 space-y-4">

        <Card title="Escenario a estimar">
          <div className="flex gap-1.5 mb-3">
            {AREAS_INSUMOS.map((a) => (
              <button key={a.key} onClick={() => setArea(a.key)}
                className={`text-xs rounded-full px-3 py-1.5 border font-medium transition-colors ${area === a.key ? "bg-slate-800 text-white border-slate-800" : "border-slate-300 text-slate-600 hover:border-slate-500"}`}>
                {a.label}
              </button>
            ))}
          </div>
          {(area === "seleccion" || area === "lavado") && (
            <div className="mb-3">
              <NumField label="Kg a procesar" value={kg} onChange={setKg} unit="Kg" />
            </div>
          )}
          <NumField label="Turnos a cubrir" value={nTurnos} onChange={setNTurnos} />
        </Card>

        {loading ? <Loader /> : (
          <Card title="Estimación">
            {detalle.length === 0 ? (
              <EmptyNote text={sinCostos ? "Ninguno de los insumos de esta área tiene costo unitario configurado todavía (Configurar Insumos)." : "Ingresa Kg y/o turnos para ver la estimación."} />
            ) : (
              <div className="space-y-2">
                {detalle.map((d) => (
                  <div key={d.nombre} className="flex justify-between items-center text-sm border-b border-slate-100 pb-1.5 last:border-0 last:pb-0">
                    <div>
                      <div className="font-medium text-slate-800">{d.nombre}</div>
                      <div className="text-xs text-slate-400">{d.cantidad} {d.formato}{d.cantidad !== 1 ? "s" : ""}</div>
                    </div>
                    <div className="font-semibold text-slate-900">${fmtNum(d.costo, 0)}</div>
                  </div>
                ))}
                <div className="flex justify-between items-center pt-2 mt-2 border-t border-slate-300">
                  <span className="text-sm font-bold text-slate-800">Total estimado</span>
                  <span className="text-lg font-bold text-emerald-700">${fmtNum(totalEstimado, 0)}</span>
                </div>
              </div>
            )}
          </Card>
        )}

        <p className="text-xs text-slate-400 text-center">
          Esta calculadora no guarda nada — solo simula un escenario usando los costos unitarios de "Configurar Insumos".
        </p>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ESPECIFICACIÓN DE SKU — Hook de overrides persistentes
// El Jefe puede modificar cualquier campo del maestro de SKU. Los cambios se
// guardan como un array de { sku, ...camposModificados } y se fusionan con
// SKU_MATERIALES cuando se consulta un SKU.
// ---------------------------------------------------------------------------
function useSkuOverrides() {
  const [overrides, saveOverrides, loading] = useSharedList("sku-overrides");

  const overrideMap = useMemo(() => {
    const m = {};
    (overrides || []).forEach((o) => { if (o.sku) m[o.sku] = o; });
    return m;
  }, [overrides]);

  // Devuelve el material fusionado (base del maestro o SKU nuevo + overrides del Jefe)
  const getMat = (sku) => {
    const base = SKU_MATERIALES.find((s) => s.sku === sku) || SKU_NUEVOS.find((s) => s.sku === sku);
    if (!base) return null;
    const over = overrideMap[sku] || {};
    return { ...base, ...over };
  };

  // Persiste un override para un SKU
  const guardarOverride = async (sku, campos) => {
    const existing = (overrides || []).filter((o) => o.sku !== sku);
    return saveOverrides([...existing, { sku, ...campos }]);
  };

  // Elimina el override (vuelve al maestro original)
  const resetOverride = async (sku) => {
    return saveOverrides((overrides || []).filter((o) => o.sku !== sku));
  };

  const tieneOverride = (sku) => !!overrideMap[sku];

  return { getMat, guardarOverride, resetOverride, tieneOverride, loading };
}

// Campos del maestro de SKU que se muestran y pueden editarse
const CAMPOS_SKU = [
  { key: "producto",      label: "Nombre del producto", tipo: "text",   ancho: "col-span-2" },
  { key: "cliente",       label: "Cliente",              tipo: "text"  },
  { key: "codBolsa",      label: "Código de bolsa",      tipo: "text"  },
  { key: "nomBolsa",      label: "Nombre de bolsa",      tipo: "text"  },
  { key: "codCaja",       label: "Código de caja",       tipo: "text"  },
  { key: "nomCaja",       label: "Nombre de caja",       tipo: "text"  },
  { key: "bolsasXCaja",   label: "Bolsas x Caja",        tipo: "number"},
  { key: "cajasXPallet",  label: "Cajas x Pallet",       tipo: "number"},
  { key: "kgXCaja",       label: "Kg x Caja",            tipo: "number"},
  { key: "tipoPallet",    label: "Tipo de Pallet",        tipo: "text"  },
  { key: "tixhi",         label: "TI x HI",              tipo: "text"  },
  { key: "slipSheet",     label: "Slip Sheet",            tipo: "text"  },
];

// ---------------------------------------------------------------------------
// ESPECIFICACIÓN — Vista de consulta (Envasado portal, solo lectura)
// ---------------------------------------------------------------------------
// Necesidades de materiales e ingredientes por día, según el Programa de
// producción de la Envasadora (cajas por SKU).
function NecesidadesDiaCard({ getMat }) {
  const [programas] = useSharedList("programa-records");
  const [fecha, setFecha] = useState("");
  const [turnosSel, setTurnosSel] = useState(["T1", "T2", "T3"]);

  const fechasProg = useMemo(() => {
    const hoy = today();
    const f = new Set();
    (programas || []).forEach((p) => {
      if ((p.lineaKey || p.linea) !== "envasadora" || p.especie === "LAVADO" || !p.fecha) return;
      if (p.fecha >= hoy) f.add(p.fecha);
    });
    return [...f].sort().slice(0, 14);
  }, [programas]);

  useEffect(() => {
    if (!fecha && fechasProg.length) setFecha(fechasProg[0]);
  }, [fechasProg, fecha]);

  const todos = turnosSel.length === 3;
  // Desde "Todos", tocar un turno deja solo ese turno.
  const toggleTurno = (t) => setTurnosSel((p) => (p.length === 3 ? [t] : p.includes(t) ? p.filter((x) => x !== t) : [...p, t]));

  const calc = useMemo(() => {
    const porSku = {};
    const fuera = [];
    (programas || []).forEach((p) => {
      if ((p.lineaKey || p.linea) !== "envasadora" || p.especie === "LAVADO") return;
      if (p.fecha !== fecha || !turnosSel.includes(p.turno)) return;
      const cajas = num(p.cantidad);
      if (cajas <= 0) return;
      const mat = getMat(p.especie);
      if (!mat) { fuera.push(p.especie); return; }
      porSku[p.especie] = (porSku[p.especie] || 0) + cajas;
    });
    const filas = Object.entries(porSku).map(([sku, cajas]) => {
      const mat = getMat(sku);
      const bxc = num(mat.bolsasXCaja);
      const kgCaja = num(mat.kgXCaja);
      const kgProd = cajas * kgCaja;
      const receta = (RECETA_SKU[sku] || []).map(([ing, prop]) => ({ ing, pct: prop * 100, kg: kgProd * prop }));
      return {
        sku, producto: mat.producto, cajas, bolsas: bxc > 0 ? Math.ceil(cajas * bxc) : 0, bxc,
        pallets: num(mat.cajasXPallet) > 0 ? cajas / num(mat.cajasXPallet) : 0,
        codBolsa: mat.codBolsa || "", nomBolsa: mat.nomBolsa || "",
        codCaja: mat.codCaja || "", nomCaja: mat.nomCaja || "",
        kgCaja, kgProd, receta,
      };
    }).sort((a, b) => a.sku.localeCompare(b.sku));

    const agrup = (cod, nom, campo) => {
      const m = {};
      filas.forEach((f) => { if (f[cod]) { m[f[cod]] = m[f[cod]] || { cod: f[cod], nom: f[nom], total: 0 }; m[f[cod]].total += f[campo]; } });
      return Object.values(m).sort((a, b) => a.cod.localeCompare(b.cod));
    };
    const ingr = {};
    filas.forEach((f) => f.receta.forEach((r) => { ingr[r.ing] = (ingr[r.ing] || 0) + r.kg; }));
    return {
      filas, fuera: [...new Set(fuera)],
      bolsas: agrup("codBolsa", "nomBolsa", "bolsas"),
      cajas: agrup("codCaja", "nomCaja", "cajas"),
      ingredientes: Object.entries(ingr).sort((a, b) => b[1] - a[1]),
      totCajas: filas.reduce((s, f) => s + f.cajas, 0),
      totBolsas: filas.reduce((s, f) => s + f.bolsas, 0),
      totKg: filas.reduce((s, f) => s + f.kgProd, 0),
    };
  }, [programas, fecha, turnosSel, getMat]);

  return (
    <Card title="Necesidades del día según programación">
      {fechasProg.length === 0 ? (
        <EmptyNote text="No hay programación de Envasadora cargada desde hoy en adelante." />
      ) : (
        <>
          <p className="text-xs text-slate-500 mb-1.5">Elige un día programado:</p>
          <div className="flex flex-wrap gap-1.5 mb-3">
            {fechasProg.map((f) => (
              <button key={f} onClick={() => setFecha(f)}
                className={`text-xs rounded-full px-3 py-1 border font-medium transition-colors ${f === fecha ? "bg-amber-500 text-white border-amber-500" : "border-amber-300 text-amber-700 bg-amber-50 hover:bg-amber-100"}`}>
                {fmtFecha(f)}
              </button>
            ))}
          </div>
          <div className="flex gap-1.5 mb-3">
            <button onClick={() => setTurnosSel(todos ? [] : ["T1", "T2", "T3"])}
              className={`flex-1 text-xs rounded-lg py-1.5 border font-semibold ${todos ? "bg-slate-700 text-white border-slate-700" : "bg-white text-slate-600 border-slate-300"}`}>Todos</button>
            {["T1", "T2", "T3"].map((t) => (
              <button key={t} onClick={() => toggleTurno(t)}
                className={`flex-1 text-xs rounded-lg py-1.5 border font-semibold ${turnosSel.includes(t) && !todos ? "bg-amber-500 text-white border-amber-500" : "bg-white text-slate-600 border-slate-300"}`}>{t}</button>
            ))}
          </div>

          {calc.filas.length === 0 ? (
            <EmptyNote text="Sin SKU programados en Envasadora para ese día/turnos." />
          ) : (
            <div className="space-y-3">
              <div className="grid grid-cols-3 gap-2 text-center">
                {[["Cajas", fmtNum(calc.totCajas)], ["Bolsas", fmtNum(calc.totBolsas)], ["Kg producto", fmtNum(calc.totKg, 0)]].map(([l, v]) => (
                  <div key={l} className="border border-slate-200 rounded-xl p-2">
                    <div className="text-xs text-slate-500">{l}</div>
                    <div className="font-bold text-slate-900 text-lg">{v}</div>
                  </div>
                ))}
              </div>

              {calc.filas.map((f) => (
                <div key={f.sku} className="border border-slate-200 rounded-xl p-3">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-mono bg-amber-100 text-amber-800 border border-amber-200 rounded-full px-2 py-0.5">{f.sku}</span>
                    <span className="text-xs text-slate-500">{fmtNum(f.pallets, 2)} pallets</span>
                  </div>
                  <div className="text-sm font-semibold text-slate-800 mt-1">{f.producto}</div>
                  <div className="grid grid-cols-3 gap-2 mt-2 text-center text-xs">
                    <div className="bg-amber-50 rounded-lg p-1.5"><div className="text-slate-500">Cajas</div><div className="font-bold text-amber-800">{fmtNum(f.cajas)}</div></div>
                    <div className="bg-blue-50 rounded-lg p-1.5"><div className="text-slate-500">Bolsas{f.bxc ? ` (×${f.bxc})` : ""}</div><div className="font-bold text-blue-800">{f.bxc ? fmtNum(f.bolsas) : "—"}</div></div>
                    <div className="bg-slate-50 rounded-lg p-1.5"><div className="text-slate-500">Kg producto</div><div className="font-bold text-slate-800">{f.kgCaja ? fmtNum(f.kgProd, 0) : "—"}</div></div>
                  </div>
                  <div className="mt-2 text-xs text-slate-600">
                    {f.codBolsa && <div>Bolsa: <b>{f.codBolsa}</b> · {f.nomBolsa}</div>}
                    {f.codCaja && <div>Caja: <b>{f.codCaja}</b> · {f.nomCaja}</div>}
                  </div>
                  {f.receta.length > 0 && (
                    <table className="w-full text-xs mt-2">
                      <thead><tr className="text-slate-500 text-left"><th className="font-medium py-0.5">Ingrediente</th><th className="font-medium text-right">%</th><th className="font-medium text-right">Kg a usar</th></tr></thead>
                      <tbody>
                        {f.receta.map((r) => (
                          <tr key={r.ing} className="border-t border-slate-100">
                            <td className="py-0.5 text-slate-800">{r.ing}</td>
                            <td className="text-right">{fmtNum(r.pct, 1)}%</td>
                            <td className="text-right font-semibold">{f.kgCaja ? fmtNum(r.kg, 0) : "—"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              ))}

              <div className="border border-blue-100 bg-blue-50 rounded-xl p-3">
                <div className="text-xs font-semibold text-blue-800 mb-1">Total de bolsas por código</div>
                {calc.bolsas.map((b) => (
                  <div key={b.cod} className="flex justify-between gap-2 text-xs py-0.5"><span className="text-slate-700"><b>{b.cod}</b> {b.nom}</span><span className="font-bold text-blue-800">{fmtNum(b.total)}</span></div>
                ))}
              </div>
              <div className="border border-amber-100 bg-amber-50 rounded-xl p-3">
                <div className="text-xs font-semibold text-amber-800 mb-1">Total de cajas por código</div>
                {calc.cajas.map((b) => (
                  <div key={b.cod} className="flex justify-between gap-2 text-xs py-0.5"><span className="text-slate-700"><b>{b.cod}</b> {b.nom}</span><span className="font-bold text-amber-800">{fmtNum(b.total)}</span></div>
                ))}
              </div>
              <div className="border border-emerald-100 bg-emerald-50 rounded-xl p-3">
                <div className="text-xs font-semibold text-emerald-800 mb-1">Consumo de ingredientes (Kg)</div>
                {calc.ingredientes.map(([ing, kg]) => (
                  <div key={ing} className="flex justify-between gap-2 text-xs py-0.5"><span className="text-slate-700">{ing}</span><span className="font-bold text-emerald-800">{fmtNum(kg, 0)} Kg</span></div>
                ))}
              </div>
              {calc.fuera.length > 0 && (
                <p className="text-xs text-red-600">SKU programados que ya no están en el maestro y no se calcularon: {calc.fuera.join(", ")}</p>
              )}
            </div>
          )}
        </>
      )}
    </Card>
  );
}

function EspecificacionScreen({ onBack }) {
  const { getMat } = useSkuOverrides();
  const [skuSel, setSkuSel] = useState("");
  const mat = skuSel ? getMat(skuSel) : null;
  const cod = mat?.cliente ? CODIFICACION_POR_CLIENTE[mat.cliente] : null;

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Especificación de SKU" subtitle="Materiales, codificaciones y configuración de pallet" onBack={onBack} icon={Package} accent="amber" />
      <div className="px-4 pt-4 space-y-4">

        <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3">
          <p className="text-xs text-amber-800">
            ⚠ Recordar validar esta información con las Especificaciones de Producto administradas por el Área de Calidad.
          </p>
        </div>

        <Card title="Buscar SKU">
          <SkuPicker
            label="Código SKU o descripción"
            value={skuSel}
            onChange={setSkuSel}
          />
        </Card>

        <NecesidadesDiaCard getMat={getMat} />

        {mat && (
          <>
            <Card title={`${mat.producto}`}>
              <div className="mb-2 flex items-center gap-2">
                <span className="text-xs font-mono bg-amber-100 text-amber-800 border border-amber-200 rounded-full px-3 py-1">{mat.sku}</span>
                {mat.cliente && <span className="text-xs bg-slate-100 text-slate-600 rounded-full px-3 py-1">{mat.cliente}</span>}
              </div>
              <div className="grid grid-cols-2 gap-3 text-sm">
                {mat.codBolsa && (
                  <div className="border border-blue-100 bg-blue-50 rounded-xl p-2.5">
                    <div className="text-xs text-slate-500 mb-0.5">Bolsa</div>
                    <div className="font-bold text-blue-800 text-base leading-tight">{mat.codBolsa}</div>
                    <div className="text-xs text-slate-600 mt-0.5">{mat.nomBolsa}</div>
                  </div>
                )}
                {mat.codCaja && (
                  <div className="border border-amber-100 bg-amber-50 rounded-xl p-2.5">
                    <div className="text-xs text-slate-500 mb-0.5">Caja</div>
                    <div className="font-bold text-amber-800 text-base leading-tight">{mat.codCaja}</div>
                    <div className="text-xs text-slate-600 mt-0.5">{mat.nomCaja}</div>
                  </div>
                )}
              </div>
              <div className="grid grid-cols-3 gap-2 mt-3 text-sm">
                {[
                  ["Bolsas x Caja", mat.bolsasXCaja],
                  ["Cajas x Pallet", mat.cajasXPallet],
                  ["Kg x Caja", mat.kgXCaja],
                ].filter(([, v]) => v !== "" && v != null).map(([label, val]) => (
                  <div key={label} className="border border-slate-200 rounded-xl p-2 text-center">
                    <div className="text-xs text-slate-500">{label}</div>
                    <div className="font-bold text-slate-900 text-lg">{fmtNum(num(val), num(val) % 1 !== 0 ? 2 : 0)}</div>
                  </div>
                ))}
              </div>
              {(mat.tipoPallet || mat.tixhi || mat.slipSheet) && (
                <div className="grid grid-cols-2 gap-2 mt-2 text-sm">
                  {mat.tipoPallet && (
                    <div className="border border-slate-200 rounded-xl p-2.5">
                      <div className="text-xs text-slate-500">Tipo Pallet</div>
                      <div className="font-semibold text-slate-800">{mat.tipoPallet}{mat.tixhi ? ` · ${mat.tixhi}` : ""}</div>
                    </div>
                  )}
                  {mat.slipSheet && (
                    <div className="border border-slate-200 rounded-xl p-2.5">
                      <div className="text-xs text-slate-500">Slip Sheet</div>
                      <div className="font-semibold text-slate-800">{mat.slipSheet}</div>
                    </div>
                  )}
                </div>
              )}
            </Card>

            {cod && (
              <Card title="Codificación">
                <div className="grid grid-cols-1 gap-3">
                  <div>
                    <div className="text-xs text-slate-500 mb-1">Codificación bolsa</div>
                    <div className="font-mono text-xs text-slate-800 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 whitespace-pre-wrap">{cod.bolsa || "—"}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 mb-1">Codificación caja</div>
                    <div className="font-mono text-xs text-slate-800 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 whitespace-pre-wrap">{cod.caja || "—"}</div>
                  </div>
                </div>
              </Card>
            )}
          </>
        )}

        {skuSel && !mat && (
          <Card><EmptyNote text="Código SKU no encontrado en el maestro de materiales." /></Card>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ESPECIFICACIÓN — Editor (Jefe, campos modificables + persist overrides)
// ---------------------------------------------------------------------------
function EspecificacionEditScreen({ onBack }) {
  const { getMat, guardarOverride, resetOverride, tieneOverride, loading } = useSkuOverrides();
  const [skuSel,  setSkuSel]  = useState("");
  const [campos,  setCampos]  = useState({});
  const [dirty,   setDirty]   = useState(false);
  const [toast,   setToast]   = useState(null);

  const matBase = skuSel ? SKU_MATERIALES.find((s) => s.sku === skuSel) : null;
  const matFull = skuSel ? getMat(skuSel) : null;
  const conOverride = skuSel ? tieneOverride(skuSel) : false;

  // Al seleccionar un SKU, cargar sus valores actuales (base + overrides)
  useEffect(() => {
    if (matFull) {
      const c = {};
      CAMPOS_SKU.forEach(({ key }) => { c[key] = matFull[key] ?? ""; });
      setCampos(c);
      setDirty(false);
      setToast(null);
    } else {
      setCampos({});
    }
  }, [skuSel, matFull?.sku]);

  const setCampo = (key, val) => {
    setCampos((p) => ({ ...p, [key]: val }));
    setDirty(true);
  };

  const guardar = async () => {
    if (!skuSel) return;
    // Solo guarda los campos que difieren del maestro base
    const diff = {};
    CAMPOS_SKU.forEach(({ key }) => {
      if (String(campos[key] ?? "") !== String(matBase?.[key] ?? "")) diff[key] = campos[key];
    });
    const ok = await guardarOverride(skuSel, diff);
    setToast({ kind: ok ? "ok" : "error", message: ok ? "Especificación guardada." : "No se pudo guardar." });
    if (ok) setDirty(false);
    setTimeout(() => setToast(null), 3000);
  };

  const restaurar = async () => {
    const ok = await resetOverride(skuSel);
    if (ok && matBase) {
      const c = {};
      CAMPOS_SKU.forEach(({ key }) => { c[key] = matBase[key] ?? ""; });
      setCampos(c);
      setDirty(false);
    }
    setToast({ kind: ok ? "ok" : "error", message: ok ? "Restaurado al maestro original." : "No se pudo restaurar." });
    setTimeout(() => setToast(null), 3000);
  };

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Editar Especificación SKU" subtitle="Jefe de producción — modificaciones al maestro" onBack={onBack} icon={Settings} accent="amber" />
      <div className="px-4 pt-4 space-y-4">

        <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3">
          <p className="text-xs text-amber-800">
            ⚠ Recordar validar esta información con las Especificaciones de Producto administradas por el Área de Calidad.
          </p>
        </div>

        <Card title="Seleccionar SKU">
          <SkuPicker
            label="Código SKU o descripción a editar"
            value={skuSel}
            onChange={setSkuSel}
          />
          {conOverride && (
            <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-1.5 mt-2">
              ⚠ Este SKU tiene modificaciones guardadas sobre el maestro original.
            </p>
          )}
        </Card>

        {matFull && (
          <>
            <Card title="Campos editables">
              {loading ? <Loader /> : (
                <div className="grid grid-cols-2 gap-3">
                  {CAMPOS_SKU.map(({ key, label, tipo, ancho }) => (
                    <div key={key} className={ancho || ""}>
                      <label className="block text-xs font-medium text-slate-600 mb-1">
                        {label}
                        {String(campos[key] ?? "") !== String(matBase?.[key] ?? "") && (
                          <span className="ml-1.5 text-amber-600 font-semibold">✎ modificado</span>
                        )}
                      </label>
                      <input
                        type={tipo}
                        className={inputBase + (String(campos[key] ?? "") !== String(matBase?.[key] ?? "") ? " border-amber-400 bg-amber-50" : "")}
                        value={campos[key] ?? ""}
                        onChange={(e) => setCampo(key, e.target.value)}
                        placeholder={`${label}…`}
                      />
                      {String(campos[key] ?? "") !== String(matBase?.[key] ?? "") && (
                        <p className="text-xs text-slate-400 mt-0.5">
                          Original: {matBase?.[key] !== undefined && matBase?.[key] !== "" ? String(matBase[key]) : "—"}
                        </p>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </Card>

            {toast && <Toast {...toast} />}

            <div className="flex gap-2">
              <button onClick={guardar} disabled={!dirty}
                className={`flex-1 flex items-center justify-center gap-2 font-semibold rounded-xl py-3 transition-colors text-sm ${dirty ? "bg-amber-500 hover:bg-amber-400 text-white" : "bg-slate-200 text-slate-400"}`}>
                <Save size={16} /> {dirty ? "Guardar cambios" : "Sin cambios"}
              </button>
              {conOverride && (
                <button onClick={restaurar}
                  className="px-4 rounded-xl border border-red-300 text-red-600 hover:bg-red-50 text-xs font-medium transition-colors">
                  Restaurar original
                </button>
              )}
            </div>
            <p className="text-xs text-slate-400 text-center">
              Los cambios afectan la visualización en toda la app. El maestro de SKU original no se modifica permanentemente — se puede restaurar en cualquier momento.
            </p>
          </>
        )}
      </div>
    </div>
  );
}



// ---------------------------------------------------------------------------
// GESTIONAR SUPERVISORES (Jefe de producción)
// ---------------------------------------------------------------------------
const AREAS_SUPERVISOR = ["Seleccion", "Envasado", "Ambas"];

function SupervisoresScreen({ onBack }) {
  const [supervisores, saveSupervisores, loading] = useSupervisores();
  const [nombre, setNombre] = useState("");
  const [area,   setArea]   = useState("Ambas");
  const [toast,  setToast]  = useState(null);

  const agregar = async () => {
    if (!nombre.trim()) {
      setToast({ kind: "error", message: "Escribe el nombre del supervisor." });
      return;
    }
    const nuevo = { id: Date.now(), nombre: nombre.trim(), area };
    const ok = await saveSupervisores([...(supervisores || []), nuevo]);
    if (ok) { setNombre(""); setToast({ kind: "ok", message: `${nuevo.nombre} agregado.` }); }
    else    { setToast({ kind: "error", message: "No se pudo guardar." }); }
    setTimeout(() => setToast(null), 2500);
  };

  const eliminar = async (id) => {
    const ok = await saveSupervisores((supervisores || []).filter((s) => s.id !== id));
    if (!ok) setToast({ kind: "error", message: "No se pudo eliminar." });
  };

  const areaLabel = { Seleccion: "Selección", Envasado: "Envasado", Ambas: "Ambas áreas" };

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Gestionar Supervisores" subtitle="Lista de supervisores disponibles en los portales" onBack={onBack} icon={User} accent="blue" />
      <div className="px-4 pt-4 space-y-4">

        {/* Agregar supervisor */}
        <Card title="Agregar supervisor">
          <div className="space-y-3">
            <TextField label="Nombre completo" value={nombre} onChange={setNombre} placeholder="Ej: Carlos Muñoz" />
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Área asignada</label>
              <div className="flex gap-2">
                {AREAS_SUPERVISOR.map((a) => (
                  <button key={a} onClick={() => setArea(a)}
                    className={`flex-1 text-xs font-semibold rounded-xl py-2.5 border transition-all ${area === a ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-600 border-slate-300 hover:border-blue-400"}`}>
                    {areaLabel[a]}
                  </button>
                ))}
              </div>
            </div>
            {toast && <Toast {...toast} />}
            <button onClick={agregar}
              className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-500 text-white font-semibold rounded-xl py-2.5 transition-colors text-sm">
              <Plus size={16} /> Agregar supervisor
            </button>
          </div>
        </Card>

        {/* Lista actual */}
        <Card title={`Supervisores registrados (${(supervisores || []).length})`}>
          {loading ? <Loader /> : (supervisores || []).length === 0 ? (
            <EmptyNote text="Sin supervisores registrados todavía. Agrega el primero arriba." />
          ) : (
            <div className="space-y-2">
              {[...supervisores].sort((a, b) => a.nombre.localeCompare(b.nombre)).map((s) => (
                <div key={s.id} className="flex items-center justify-between border border-slate-200 rounded-xl px-3 py-2.5 bg-white">
                  <div>
                    <div className="text-sm font-semibold text-slate-900">{s.nombre}</div>
                    <div className="text-xs text-slate-500">{areaLabel[s.area] || s.area}</div>
                  </div>
                  <button onClick={() => eliminar(s.id)}
                    className="p-2 rounded-lg hover:bg-red-50 text-slate-400 hover:text-red-600 transition-colors"
                    aria-label="Eliminar">
                    <Trash2 size={16} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </Card>

        <p className="text-xs text-slate-400 text-center">
          Los supervisores registrados aquí aparecen como selector desplegable en los portales de Selección y Envasado.
        </p>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// GESTIONAR HORARIOS DE TURNO (Jefe de producción)
// Hora de inicio/fin por turno, Lunes a Sábado. De aquí se calculan los
// minutos efectivos (descontando colación + inicio + término) que se usan en
// el Inicio de turno para mostrar Kg/hora y Kg/min necesarios.
// ---------------------------------------------------------------------------
function HorariosScreen({ onBack }) {
  const { horarios, guardar, loading } = useHorariosTurno();
  const { habilitados, guardar: guardarHabilitados, loading: loadingHab } = useTurnosHabilitados();
  const { habilitados: diasHabilitados, guardar: guardarDias, loading: loadingDias } = useDiasHabilitados();
  const [borrador, setBorrador] = useState(null);
  const [toast, setToast] = useState(null);
  const activo = borrador || horarios;

  const turnosHabilitadosKeys = TURNOS.filter((t) => habilitados[t.key]).map((t) => t.key);
  const diasHabilitadosKeys = DIAS_SEMANA.filter((d) => diasHabilitados[d]);

  const toggleTurnoHabilitado = async (key) => {
    const nuevo = { ...habilitados, [key]: !habilitados[key] };
    const ok = await guardarHabilitados(nuevo);
    if (!ok) setToast({ kind: "error", message: "No se pudo actualizar el turno habilitado." });
  };

  const toggleDiaHabilitado = async (dia) => {
    const nuevo = { ...diasHabilitados, [dia]: !diasHabilitados[dia] };
    const ok = await guardarDias(nuevo);
    if (!ok) setToast({ kind: "error", message: "No se pudo actualizar el día habilitado." });
  };

  const asegurarBorrador = () => {
    if (!borrador) { const c = JSON.parse(JSON.stringify(horarios)); setBorrador(c); return c; }
    return borrador;
  };

  const setHora = (dia, turno, campo, valor) => {
    const base = asegurarBorrador();
    const clone = JSON.parse(JSON.stringify(base));
    clone[dia] = clone[dia] || {};
    clone[dia][turno] = { ...(clone[dia][turno] || {}), [campo]: valor };
    setBorrador(clone);
  };

  const guardarCambios = async () => {
    const ok = await guardar(borrador || horarios);
    setToast({ kind: ok ? "ok" : "error", message: ok ? "Horarios guardados." : "No se pudo guardar." });
    if (ok) setBorrador(null);
    setTimeout(() => setToast(null), 2500);
  };

  const restaurar = async () => {
    const ok = await guardar(HORARIOS_DEFAULT);
    setBorrador(null);
    setToast({ kind: ok ? "ok" : "error", message: ok ? "Se restauraron los horarios por defecto." : "No se pudo restaurar." });
    setTimeout(() => setToast(null), 2500);
  };

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Horarios de Turno" subtitle="Lunes a Sábado — usados para calcular minutos efectivos" onBack={onBack} icon={Clock} accent="blue" />
      <div className="px-4 pt-4 space-y-4">

        <Card title="¿Qué días de la semana trabaja la planta?">
          {loadingDias ? <Loader /> : (
            <>
              <div className="grid grid-cols-4 sm:grid-cols-7 gap-1.5">
                {DIAS_SEMANA.map((dia) => {
                  const on = diasHabilitados[dia];
                  return (
                    <button key={dia} onClick={() => toggleDiaHabilitado(dia)}
                      className={`rounded-xl py-2.5 border transition-all text-xs font-semibold ${on ? "bg-blue-600 text-white border-blue-600 shadow-sm" : "bg-white text-slate-400 border-slate-300 hover:border-blue-400"}`}
                    >
                      {dia.slice(0, 3)}
                    </button>
                  );
                })}
              </div>
              <p className="text-xs text-slate-400 mt-2">
                La semana es de Lunes a Domingo. Solo los días habilitados aquí piden configuración de horario abajo.
              </p>
            </>
          )}
        </Card>

        <Card title="¿Qué turnos opera la planta?">
          {loadingHab ? <Loader /> : (
            <>
              <div className="flex gap-2">
                {TURNOS.map(({ key, label }) => {
                  const TurnoIcon = TURNO_ICONS[key];
                  const on = habilitados[key];
                  return (
                    <button key={key} onClick={() => toggleTurnoHabilitado(key)}
                      className={`flex-1 flex flex-col items-center gap-1 rounded-xl py-3 border transition-all text-xs font-semibold ${on ? "bg-blue-600 text-white border-blue-600 shadow-sm" : "bg-white text-slate-400 border-slate-300 hover:border-blue-400"}`}
                    >
                      {TurnoIcon && <TurnoIcon size={18} />}{key}
                    </button>
                  );
                })}
              </div>
              <p className="text-xs text-slate-400 mt-2">
                Solo los turnos habilitados aquí piden configuración de horario abajo, y son los que se usan en Insumos y Consumo para calcular la cobertura de bodega.
              </p>
            </>
          )}
        </Card>

        {turnosHabilitadosKeys.length === 0 || diasHabilitadosKeys.length === 0 ? (
          <Card><EmptyNote text="Habilita al menos un turno y un día para configurar sus horarios." /></Card>
        ) : (
          <>
            <p className="text-xs text-slate-500 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2">
              De cada turno se descuentan 30 min de colación + 15 min de inicio + 15 min de término (60 min fijos) para calcular los minutos efectivos.
            </p>

            {loading ? <Loader /> : diasHabilitadosKeys.map((dia) => (
              <Card key={dia} title={dia}>
                <div className="space-y-3">
                  {TURNOS.filter((t) => turnosHabilitadosKeys.includes(t.key)).map(({ key: turno, label }) => {
                    const cfg = activo[dia]?.[turno] || { horaInicio: "", horaFin: "" };
                    const efectivos = minutosEfectivosTurno(cfg.horaInicio, cfg.horaFin);
                    const bruto = minutosBrutosTurno(cfg.horaInicio, cfg.horaFin);
                    return (
                      <div key={turno} className="border border-slate-200 rounded-lg p-2.5">
                        <div className="flex items-center justify-between mb-2">
                          <span className="text-xs font-semibold text-slate-700">{label}</span>
                          {bruto > 0 && (
                            <span className="text-xs text-slate-500">
                              {bruto} min brutos → <b className="text-emerald-700">{efectivos} min efectivos</b>
                            </span>
                          )}
                        </div>
                        <div className="grid grid-cols-2 gap-2">
                          <TextField label="Hora inicio" type="time" value={cfg.horaInicio} onChange={(v) => setHora(dia, turno, "horaInicio", v)} />
                          <TextField label="Hora fin" type="time" value={cfg.horaFin} onChange={(v) => setHora(dia, turno, "horaFin", v)} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </Card>
            ))}

            {toast && <Toast {...toast} />}

            <div className="flex gap-2">
              <button onClick={guardarCambios}
                className="flex-1 flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-500 text-white font-semibold rounded-xl py-3 transition-colors text-sm">
                <Save size={16} /> Guardar horarios
              </button>
              <button onClick={restaurar}
                className="px-4 rounded-xl border border-slate-300 text-slate-500 hover:border-red-300 hover:text-red-600 text-xs font-medium transition-colors">
                Restaurar
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function Loader({ text = "Cargando…" }) {
  return (
    <p className="text-sm text-slate-500 flex items-center gap-2">
      <Loader2 size={14} className="animate-spin" /> {text}
    </p>
  );
}

function EmptyNote({ text }) {
  return <p className="text-sm text-slate-500 italic">{text}</p>;
}

function Toast({ message, kind = "ok" }) {
  if (!message) return null;
  const styles = kind === "error"
    ? "bg-red-500/10 text-red-700 border-red-500/30"
    : "bg-emerald-500/10 text-emerald-700 border-emerald-500/30";
  return (
    <div className={`flex items-center gap-2 text-sm border rounded-lg px-3 py-2 ${styles}`}>
      {kind === "error" ? <AlertTriangle size={16} /> : <CheckCircle2 size={16} />}
      {message}
    </div>
  );
}

function Card({ title, step, children, className = "" }) {
  return (
    <div className={`bg-white border border-slate-200 shadow-sm rounded-xl p-4 ${className}`}>
      {title && (
        <div className="flex items-center gap-2 mb-3">
          {step && (
            <span className="flex items-center justify-center w-6 h-6 rounded-full border border-slate-500 text-slate-700 text-xs font-semibold flex-shrink-0">
              {step}
            </span>
          )}
          <h3 className="text-xs font-semibold uppercase tracking-wide text-slate-700">{title}</h3>
        </div>
      )}
      {children}
    </div>
  );
}

function TopBar({ title, subtitle, onBack, icon: Icon, accent }) {
  return (
    <div className="sticky top-0 z-10 bg-white border-b border-slate-200 px-4 py-3 flex items-center gap-3">
      {onBack && (
        <button onClick={onBack} aria-label="Volver" className="rounded-lg p-1.5 hover:bg-slate-200 text-slate-700">
          <ArrowLeft size={20} />
        </button>
      )}
      {Icon && (
        <div className={`w-9 h-9 rounded-lg flex items-center justify-center border ${ACCENT[accent]?.badge || "bg-slate-200 text-slate-700 border-slate-300"}`}>
          <Icon size={18} />
        </div>
      )}
      <div>
        <div className="text-sm font-semibold text-slate-900 leading-tight">{title}</div>
        {subtitle && <div className="text-xs text-slate-600 leading-tight">{subtitle}</div>}
      </div>
    </div>
  );
}

function TurnoSelector({ value, onChange }) {
  return (
    <div className="grid grid-cols-3 gap-2">
      {TURNOS.map((t) => {
        const Icon = TURNO_ICONS[t.key];
        const active = value === t.key;
        return (
          <button
            key={t.key}
            type="button"
            onClick={() => onChange(t.key)}
            className={`flex flex-col items-center gap-1 rounded-lg border py-2.5 text-xs font-medium transition-colors ${
              active ? "border-blue-400 bg-blue-500/10 text-blue-700" : "border-slate-300 bg-slate-100 text-slate-600 hover:border-slate-500"
            }`}
          >
            <Icon size={16} />
            {t.label}
          </button>
        );
      })}
    </div>
  );
}

function DataGrid({ items }) {
  return (
    <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-sm">
      {items.map(([label, val], i) => (
        <div key={i}>
          <div className="text-xs text-slate-600">{label}</div>
          <div className="font-medium text-slate-900 break-words">{String(val)}</div>
        </div>
      ))}
    </div>
  );
}

function NumField({ label, value, onChange, unit }) {
  return (
    <div>
      <label className="block text-xs font-medium text-slate-600 mb-1">{label}{unit ? ` (${unit})` : ""}</label>
      <input type="number" className={inputBase} value={value ?? ""} onChange={(e) => onChange(e.target.value)} />
    </div>
  );
}

function TextField({ label, value, onChange, type = "text" }) {
  const dark = type === "date" || type === "time" ? { colorScheme: "light" } : undefined;
  return (
    <div>
      <label className="block text-xs font-medium text-slate-600 mb-1">{label}</label>
      <input type={type} style={dark} className={inputBase} value={value ?? ""} onChange={(e) => onChange(e.target.value)} />
    </div>
  );
}

function SelectField({ label, value, onChange, options, optionsKV, placeholder = "Seleccionar…" }) {
  const opts = optionsKV || (options || []).map((o) => [o, o]);
  return (
    <div>
      <label className="block text-xs font-medium text-slate-600 mb-1">{label}</label>
      <select className={inputBase} value={value ?? ""} onChange={(e) => onChange(e.target.value)}>
        <option value="">{placeholder}</option>
        {opts.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
      </select>
    </div>
  );
}

function TextAreaField({ label, value, onChange }) {
  return (
    <div className="sm:col-span-2">
      <label className="block text-xs font-medium text-slate-600 mb-1">{label}</label>
      <textarea className={inputBase} rows={3} value={value ?? ""} onChange={(e) => onChange(e.target.value)} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Comentarios del cierre (compartido por las 3 áreas)
// ---------------------------------------------------------------------------
const CATEGORIA_COLOR = {
  "Producción": "text-blue-700",
  "Calidad": "text-purple-700",
  "Mantención": "text-amber-700",
  "Seguridad": "text-red-700",
};

function ComentariosFields({ values, setField }) {
  // El cuadro de detalle (persona, descripción, documentación) se despliega
  // apenas se marca "Sí" en Incidentes o en Accidentes. Puede haber más de un
  // caso en el mismo turno, así que se guarda como lista (incAccList) y se
  // puede agregar o quitar casos con los botones de abajo.
  const hayEvento = values.huboIncidentes === "Sí" || values.huboAccidentes === "Sí";
  const incidentesLista = incidentesDeCierre(values);

  const actualizarIncidente = (idx, campo, val) => {
    const nueva = incidentesLista.map((it, i) => (i === idx ? { ...it, [campo]: val } : it));
    setField("incAccList", nueva);
  };
  const agregarIncidente = () => {
    setField("incAccList", [...incidentesLista, { nombre: "", descripcion: "", docHecha: "", docEntregada: "" }]);
  };
  const quitarIncidente = (idx) => {
    setField("incAccList", incidentesLista.filter((_, i) => i !== idx));
  };

  return (
    <Card title="Comentarios del cierre">
      <div className="space-y-3">
        {CATEGORIAS_COMENTARIO.map((cat) => (
          <div key={cat}>
            <label className={`block text-xs font-semibold mb-1 ${CATEGORIA_COLOR[cat]}`}>{cat}</label>
            <textarea
              className={inputBase}
              rows={2}
              placeholder={`Comentarios de ${cat.toLowerCase()}…`}
              value={values[`com_${cat}`] ?? ""}
              onChange={(e) => setField(`com_${cat}`, e.target.value)}
            />
          </div>
        ))}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-3 pt-3 border-t border-slate-300">
        <SelectField label="¿Hubo incidentes en el turno?" value={values.huboIncidentes} onChange={(v) => setField("huboIncidentes", v)} options={SI_NO} />
        <SelectField label="¿Hubo accidentes en el turno?" value={values.huboAccidentes} onChange={(v) => setField("huboAccidentes", v)} options={SI_NO} />
      </div>
      {hayEvento && (
        <div className="mt-3 pt-3 border-t border-red-200 bg-red-50 -mx-4 -mb-4 px-4 pb-4 rounded-b-xl">
          <p className="text-xs font-semibold text-red-700 mb-2 flex items-center gap-1.5">
            <AlertTriangle size={14} /> Detalle del incidente / accidente
          </p>
          <div className="space-y-3">
            {incidentesLista.map((it, idx) => (
              <div key={idx} className="border border-red-200 bg-white rounded-lg p-3">
                {incidentesLista.length > 1 && (
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-semibold text-red-700">Caso {idx + 1}</span>
                    <button
                      type="button"
                      onClick={() => quitarIncidente(idx)}
                      className="text-xs text-red-600 hover:text-red-800 font-medium flex items-center gap-1"
                    >
                      <X size={13} /> Quitar
                    </button>
                  </div>
                )}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <TextField
                    label="Nombre y apellido de la persona"
                    value={it.nombre}
                    onChange={(v) => actualizarIncidente(idx, "nombre", v)}
                  />
                  <SelectField
                    label="¿Se hizo la documentación?"
                    value={it.docHecha}
                    onChange={(v) => actualizarIncidente(idx, "docHecha", v)}
                    options={SI_NO}
                  />
                  <TextAreaField
                    label="¿Qué pasó? (breve descripción)"
                    value={it.descripcion}
                    onChange={(v) => actualizarIncidente(idx, "descripcion", v)}
                  />
                  <SelectField
                    label="¿Se entregó la documentación?"
                    value={it.docEntregada}
                    onChange={(v) => actualizarIncidente(idx, "docEntregada", v)}
                    options={SI_NO}
                  />
                </div>
              </div>
            ))}
          </div>
          <button
            type="button"
            onClick={agregarIncidente}
            className="mt-2 w-full flex items-center justify-center gap-1.5 text-xs font-semibold text-red-700 border border-red-300 border-dashed rounded-lg py-2 hover:bg-red-100"
          >
            <Plus size={14} /> Agregar otro incidente / accidente
          </button>
        </div>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// LAVADO DE BANDEJAS
// ---------------------------------------------------------------------------
function palletsValue(values, prefix, tipo) {
  if (tipo.factor > 0) {
    const cant = num(values[`${prefix}_${tipo.key}_cant`]);
    return Math.round((cant / tipo.factor) * 10) / 10;
  }
  return num(values[`${prefix}_${tipo.key}_pallets`]);
}

function totalPalletsFromPrefix(values, prefix) {
  return TIPOS_BANDEJA.reduce((s, t) => s + palletsValue(values, prefix, t), 0);
}

function totalPalletsManual(values, prefix) {
  return TIPOS_BANDEJA.reduce((s, t) => s + num(values[`${prefix}_${t.key}_pallets`]), 0);
}

function BandejaTable({ values, setField, prefix, soloPallets }) {
  if (soloPallets) {
    return (
      <div className="space-y-2">
        <div className="grid grid-cols-2 gap-2 text-xs font-semibold text-slate-600 uppercase tracking-wide px-1">
          <span>Tipo de bandeja</span>
          <span>Cant. pallets</span>
        </div>
        {TIPOS_BANDEJA.map((t) => (
          <div key={t.key} className="grid grid-cols-2 gap-2 items-center">
            <span className="text-sm text-slate-800">{t.label}</span>
            <input
              type="number"
              className={inputBase}
              value={values[`${prefix}_${t.key}_pallets`] ?? ""}
              onChange={(e) => setField(`${prefix}_${t.key}_pallets`, e.target.value)}
            />
          </div>
        ))}
        <div className="flex justify-between text-sm font-semibold text-slate-800 px-1 pt-1 border-t border-slate-300">
          <span>Total pallets</span>
          <span>{totalPalletsManual(values, prefix)}</span>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-2">
      <div className="grid grid-cols-3 gap-2 text-xs font-semibold text-slate-600 uppercase tracking-wide px-1">
        <span>Tipo de bandeja</span>
        <span>Cant. bandejas</span>
        <span>Cant. pallets</span>
      </div>
      {TIPOS_BANDEJA.map((t) => (
        <div key={t.key} className="grid grid-cols-3 gap-2 items-center">
          <span className="text-sm text-slate-800">{t.label}</span>
          <input
            type="number"
            className={inputBase}
            value={values[`${prefix}_${t.key}_cant`] ?? ""}
            onChange={(e) => setField(`${prefix}_${t.key}_cant`, e.target.value)}
          />
          {t.factor > 0 ? (
            <div className="text-sm text-slate-700 px-3 py-2 bg-slate-100 rounded-lg border border-slate-300">
              {palletsValue(values, prefix, t)} <span className="text-xs text-slate-500">(auto)</span>
            </div>
          ) : (
            <input
              type="number"
              className={inputBase}
              value={values[`${prefix}_${t.key}_pallets`] ?? ""}
              onChange={(e) => setField(`${prefix}_${t.key}_pallets`, e.target.value)}
            />
          )}
        </div>
      ))}
      <div className="flex justify-between text-sm font-semibold text-slate-800 px-1 pt-1 border-t border-slate-300">
        <span>Total pallets</span>
        <span>{totalPalletsFromPrefix(values, prefix)}</span>
      </div>
    </div>
  );
}

function LavadoInicio({ values, setField, editingId }) {
  const mappingBandejas = TIPOS_BANDEJA.map((t) => [`pendientes_${t.key}_pallets`, `sucios_${t.key}_pallets`]);
  const mappingMateriales = MATERIALES_LAVADO.map((_, i) => [`finMat_${i}`, `inicioMat_${i}`]);
  const mapping = [...mappingBandejas, ...mappingMateriales];
  const { prevCierre, prev } = useCarryOver("lavado", values, setField, editingId, mapping);

  return (
    <div className="space-y-4">
      <Card title="Datos generales" step={1}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <TextField label="Fecha" type="date" value={values.fecha} onChange={(v) => setField("fecha", v)} />
          <TextField label="Hora de inicio" type="time" value={values.horaInicio} onChange={(v) => setField("horaInicio", v)} />
          <div className="sm:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1">Turno</label>
            <TurnoSelector value={values.turno} onChange={(v) => setField("turno", v)} />
          </div>
        </div>
      </Card>

      <Card title="Dotación de personal" step={2}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <NumField label="Operarios" value={values.operarios} onChange={(v) => setField("operarios", v)} />
          <NumField label="Movilizadores" value={values.movilizadores} onChange={(v) => setField("movilizadores", v)} />
          <NumField label="Jefe de línea" value={values.jefeLinea} onChange={(v) => setField("jefeLinea", v)} />
          <SelectField label="¿Dotación completa?" value={values.dotacionCompleta} onChange={(v) => setField("dotacionCompleta", v)} options={SI_NO} />
          <TextAreaField label="Comentarios sobre dotación" value={values.comentariosDotacion} onChange={(v) => setField("comentariosDotacion", v)} />
        </div>
      </Card>

      <Card title="Pallets sucios al iniciar" step={3}>
        {prevCierre && (
          <p className="text-xs text-emerald-700 mb-3">
            Se completó automáticamente con los pallets pendientes por lavar del cierre {turnoLabel(prev.turno)} del {fmtFecha(prev.fecha)}. Puedes ajustarlos si es necesario.
          </p>
        )}
        <BandejaTable values={values} setField={setField} prefix="sucios" soloPallets />
      </Card>

      <Card title="Materiales piso planta" step={4}>
        {prevCierre ? (
          <p className="text-xs text-emerald-700 mb-3">
            Se completó automáticamente con los materiales de piso de planta del cierre {turnoLabel(prev.turno)} del {fmtFecha(prev.fecha)}. Puedes ajustarlos si es necesario.
          </p>
        ) : (
          <p className="text-xs text-slate-400 mb-3">
            Sin cierre de turno anterior registrado — se completó en 0. Ajusta si corresponde.
          </p>
        )}
        <MaterialesTable items={MATERIALES_LAVADO} values={values} setField={setField} prefix="inicioMat" />
      </Card>
    </div>
  );
}

function LavadoCierre({ values, setField }) {
  return (
    <div className="space-y-4">
      <Card title="Pallets lavados" step={1}>
        <BandejaTable values={values} setField={setField} prefix="lavados" />
      </Card>
      <Card title="Pallets pendientes por lavar" step={2}>
        <BandejaTable values={values} setField={setField} prefix="pendientes" soloPallets />
      </Card>
      <Card title="Insumos consumidos este turno" step={3}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <NumField label="Rollos de film" value={values.rollosFilm} onChange={(v) => setField("rollosFilm", v)} unit="Rollos" />
          <NumField label="Bolsas de bins" value={values.bolsasBins} onChange={(v) => setField("bolsasBins", v)} unit="Paquetes" />
        </div>
      </Card>
      <Card title="Materiales piso planta (fin de turno)" step={4}>
        <MaterialesTable items={MATERIALES_LAVADO} values={values} setField={setField} prefix="finMat" />
        <p className="text-xs text-slate-400 mt-2">
          Esto es lo que QUEDA disponible al cerrar el turno — se usa para calcular las alertas de vencimiento y "Tengo" en Insumos y Consumo.
        </p>
      </Card>
      <ComentariosFields values={values} setField={setField} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// SELECCIÓN
// ---------------------------------------------------------------------------
function MaterialesTable({ items, values, setField, prefix, excludeIndices }) {
  const excluir = excludeIndices || [];
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
      {items.map((label, i) => {
        if (excluir.includes(i)) return null;
        return (
          <NumField
            key={i}
            label={label}
            unit={MATERIAL_UNIDAD[label]}
            value={values[`${prefix}_${i}`]}
            onChange={(v) => setField(`${prefix}_${i}`, v)}
          />
        );
      })}
    </div>
  );
}

// Suma el campo "cantidad" a través de todas las filas de un desglose de MTC por SKU.
function sumMtcSku(rows) {
  return (Array.isArray(rows) ? rows : []).reduce((sum, r) => sum + (Number(r && r.cantidad) || 0), 0);
}

// Editor de "Cantidad MTC Cajas" o "Cantidad MTC Bolsas" desglosado por
// SKU/MTC. Ojo: son dos MTC distintos e independientes — un MTC que es solo
// de cajas y otro MTC que es solo de bolsas (no el mismo MTC con ambas
// cantidades) — por eso este editor se usa dos veces por separado (una vez
// para Cajas y otra para Bolsas), cada una con su propia lista de filas
// (SKU/MTC + cantidad). En un turno se puede envasar más de un SKU, así que
// se registra una fila por cada MTC programado. El total (suma de todas las
// filas) se guarda además en el campo legacy `legacyField` para mantener
// compatibilidad con el índice posicional de MATERIALES_ENVASADO (usado por
// el histórico de stock, alertas y registros antiguos).
function MtcPorSkuEditor({ title, unitLabel, skusProgramados, values, setField, fieldName, legacyField, referencia }) {
  const rows = Array.isArray(values[fieldName]) ? values[fieldName] : [];
  const skusKey = (skusProgramados || []).join(",");
  // Si el usuario quita a propósito todas las filas, no se vuelven a precargar solas.
  const vacioAProposito = useRef(false);
  useEffect(() => { vacioAProposito.current = false; }, [skusKey]);
  // Cantidad que se registró en el Inicio de turno para cada MTC (solo referencia en el Cierre).
  const refDe = (sku) => {
    const f = (Array.isArray(referencia) ? referencia : []).find((x) => x && x.sku === sku);
    return f && f.cantidad !== "" && f.cantidad != null ? f.cantidad : null;
  };

  // Si aún no hay filas y hay SKU(s) programados para el turno, precarga
  // una fila por cada uno para que el usuario solo tenga que ingresar la cantidad.
  // Depende también de rows.length: si el formulario se reinicia después de
  // elegir el turno (como en el Cierre), las filas se vuelven a precargar.
  useEffect(() => {
    if (rows.length === 0 && !vacioAProposito.current && skusProgramados && skusProgramados.length > 0) {
      setField(fieldName, skusProgramados.map((sku) => ({ sku, cantidad: "" })));
    }
    // Sin lista de dependencias a propósito: el formulario del Cierre se reinicia
    // DESPUÉS de elegir el turno y borraba las filas recién precargadas.
    // La condición evita bucles (solo actúa si no hay filas).
  });

  // Mantiene sincronizado el total legacy con la suma de las filas.
  const rowsKey = JSON.stringify(rows);
  useEffect(() => {
    setField(legacyField, sumMtcSku(rows) || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rowsKey]);

  const updateRow = (i, key, val) => {
    const next = rows.map((r, idx) => (idx === i ? { ...r, [key]: val } : r));
    setField(fieldName, next);
  };
  const addRow = () => setField(fieldName, [...rows, { sku: "", cantidad: "" }]);
  const removeRow = (i) => {
    const next = rows.filter((_, idx) => idx !== i);
    if (next.length === 0) vacioAProposito.current = true;
    setField(fieldName, next);
  };

  return (
    <div className="space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-sm font-semibold text-slate-900">{title} por SKU</span>
        <div className="flex items-center gap-3">
          <button type="button" onClick={addRow} className="text-xs text-emerald-700 font-medium hover:underline">
            + Agregar SKU
          </button>
          {rows.length > 0 && (
            <button type="button" onClick={() => removeRow(rows.length - 1)} className="text-xs text-rose-600 font-medium hover:underline">
              − Quitar SKU
            </button>
          )}
        </div>
      </div>
      {rows.length === 0 ? (
        <p className="text-xs text-slate-400 italic">Sin SKU registrados. Agrega uno manualmente o prográmalo en el turno.</p>
      ) : (
        <div className="space-y-2">
          {rows.map((row, i) => (
            <div key={i} className="grid grid-cols-1 sm:grid-cols-[1fr_1fr_auto] gap-2 items-end border border-slate-200 rounded-lg p-2">
              <TextField label="MTC / SKU" value={row.sku} onChange={(v) => updateRow(i, "sku", v)} />
              <div>
                <NumField label={title} unit={unitLabel} value={row.cantidad} onChange={(v) => updateRow(i, "cantidad", v)} />
                {refDe(row.sku) != null && <div className="text-[11px] text-slate-500 mt-0.5">Al inicio del turno: <b>{fmtNum(num(refDe(row.sku)))}</b> {unitLabel}</div>}
              </div>
              <button type="button" onClick={() => removeRow(i)} className="flex items-center justify-center gap-1 text-xs font-medium text-rose-600 border border-rose-200 rounded-lg px-3 py-2 hover:bg-rose-50">
                <Trash2 size={13} /> Quitar SKU
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function especieFor(procesos, proceso) {
  const found = procesos.find(([p]) => p === proceso);
  return found ? found[1] : "";
}

function LineaCard({ linea, values, setField, procesosExtra, programaEntries }) {
  const activa = values[`linea_${linea.key}_activa`] === "Sí";
  const procesosCompletos = procesosConExtra(linea.procesos, procesosExtra || []);

  // Entradas del Programa para esta línea (sin LAVADO)
  const progLinea = (programaEntries || []).filter(
    (p) => (p.lineaKey === linea.key || p.linea === linea.key) && p.especie !== "LAVADO"
  );

  // Al activar la línea: si hay procesos programados y los slots están vacíos, auto-rellenar.
  useEffect(() => {
    if (!activa || progLinea.length === 0) return;
    const hayAlgoEnSlots = ESPECIE_SLOTS.some(
      (s) => values[`linea_${linea.key}_proceso${sufijoEspecie(s)}`]
    );
    if (hayAlgoEnSlots) return; // no pisar lo que ya escribió el usuario
    progLinea.slice(0, ESPECIE_SLOTS.length).forEach((entry, idx) => {
      const slot = ESPECIE_SLOTS[idx];
      const campo = `linea_${linea.key}_proceso${sufijoEspecie(slot)}`;
      if (entry.especie) setField(campo, entry.especie);
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activa, progLinea.map(e => e.especie).join(","), linea.key]);

  const slots = ESPECIE_SLOTS.map((slot) => {
    const campo = `linea_${linea.key}_proceso${sufijoEspecie(slot)}`;
    const proceso = values[campo] || "";
    return { slot, campo, proceso, especie: especieFor(procesosCompletos, proceso) };
  });

  return (
    <div className="border border-slate-300 rounded-lg p-3 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-sm font-semibold text-slate-900">{linea.label}</span>
        <div className="flex items-center gap-2">
          {progLinea.length > 0 && !activa && (
            <span className="text-xs bg-blue-100 text-blue-700 border border-blue-200 rounded-full px-2 py-0.5">
              {progLinea.length} prog.
            </span>
          )}
          <button
            type="button"
            onClick={() => setField(`linea_${linea.key}_activa`, activa ? "No" : "Sí")}
            className={`text-xs font-medium rounded-full px-3 py-1 border transition-colors ${
              activa ? "bg-blue-500/10 border-blue-400 text-blue-700" : "border-slate-400 text-slate-600"
            }`}
          >
            {activa ? "Activa" : "Inactiva"}
          </button>
        </div>
      </div>

      {activa && (
        <div className="space-y-2">
          {progLinea.length > 0 && (
            <div className="bg-blue-50 border border-blue-200 rounded-lg px-3 py-2 text-xs text-blue-800">
              ✓ Procesos cargados desde el Programa de producción.
              <button type="button" onClick={() => {
                ESPECIE_SLOTS.forEach(s => setField(`linea_${linea.key}_proceso${sufijoEspecie(s)}`, ""));
              }} className="ml-2 underline text-blue-600">Limpiar</button>
            </div>
          )}
          {progLinea.length === 0 && (
            <div className="bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-800">
              ⚠ Esta línea está activa pero el Programa de producción no tiene nada asignado para este turno. Verifica si el Programa cambió a otra línea.
            </div>
          )}

          <SelectField
            label="Especie / proceso 1"
            value={slots[0].proceso}
            onChange={(v) => setField(slots[0].campo, v)}
            options={procesosCompletos.map((p) => p[0])}
          />
          {slots[0].especie && <p className="text-xs text-slate-500">Especie: {slots[0].especie}</p>}

          <p className="text-xs text-slate-500 pt-1">Hasta 4 especies adicionales (opcional):</p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            {slots.slice(1).map(({ slot, campo, proceso, especie }) => (
              <div key={slot}>
                <SelectField
                  label={`Especie / proceso ${slot}`}
                  value={proceso}
                  onChange={(v) => setField(campo, v)}
                  options={procesosCompletos.map((p) => p[0])}
                />
                {especie && <p className="text-xs text-slate-500 mt-1">Especie: {especie}</p>}
              </div>
            ))}
          </div>

          <div className="pt-2 border-t border-slate-300">
            <p className="text-xs font-semibold text-slate-600 uppercase tracking-wide mb-2">Dotación de {linea.label}</p>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              {DOTACION_LINEA.map((label, i) => (
                <NumField
                  key={i}
                  label={label}
                  value={values[`linea_${linea.key}_dot${i}`]}
                  onChange={(v) => setField(`linea_${linea.key}_dot${i}`, v)}
                />
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// "Armado de materiales": funciona igual que una línea (Activar/Desactivar +
// dotación), pero sin selección de especie/proceso — es una tarea de logística
// que no depende de la fruta que se está procesando.
function ArmadoMaterialesCard({ values, setField }) {
  const activa = values.armado_activa === "Sí";
  return (
    <div className="border border-slate-300 rounded-lg p-3 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-sm font-semibold text-slate-900">Armado de materiales</span>
        <button
          type="button"
          onClick={() => setField("armado_activa", activa ? "No" : "Sí")}
          className={`text-xs font-medium rounded-full px-3 py-1 border transition-colors ${
            activa ? "bg-blue-500/10 border-blue-400 text-blue-700" : "border-slate-400 text-slate-600"
          }`}
        >
          {activa ? "Activo" : "Inactivo"}
        </button>
      </div>
      {activa && (
        <div className="pt-2 border-t border-slate-300">
          <p className="text-xs font-semibold text-slate-600 uppercase tracking-wide mb-2">Dotación de Armado de materiales</p>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
            {DOTACION_LINEA.map((label, i) => (
              <NumField
                key={i}
                label={label}
                value={values[`armado_dot${i}`]}
                onChange={(v) => setField(`armado_dot${i}`, v)}
              />
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function lineaDotacionTotal(values, linea) {
  return DOTACION_LINEA.reduce((s, _, i) => s + num(values[`linea_${linea.key}_dot${i}`]), 0);
}

function armadoDotacionTotal(values) {
  return DOTACION_LINEA.reduce((s, _, i) => s + num(values[`armado_dot${i}`]), 0);
}

function totalDotacionSeleccion(values) {
  let t = 0;
  DOTACION_GENERAL_SELECCION.forEach((_, i) => { t += num(values[`dg_${i}`]); });
  LINEAS_SELECCION.forEach((l) => {
    if (values[`linea_${l.key}_activa`] === "Sí") t += lineaDotacionTotal(values, l);
  });
  if (values.armado_activa === "Sí") t += armadoDotacionTotal(values);
  return t;
}

function SeleccionInicio({ values, setField, programaEntries, editingId }) {
  const mapping = MATERIALES_SELECCION.map((_, i) => [`finMat_${i}`, `inicioMat_${i}`]);
  const { prevCierre, prev } = useCarryOver("seleccion", values, setField, editingId, mapping);
  const { porLinea } = useProcesosExtra();
  const { horarios } = useHorariosTurno();

  return (
    <div className="space-y-4">
      <Card title="Datos generales" step={1}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <TextField label="Fecha" type="date" value={values.fecha} onChange={(v) => setField("fecha", v)} />
          <TextField label="Hora de inicio" type="time" value={values.horaInicio} onChange={(v) => setField("horaInicio", v)} />
          <div className="sm:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1">Turno</label>
            <TurnoSelector value={values.turno} onChange={(v) => setField("turno", v)} />
          </div>
        </div>
      </Card>

      <ProgramaResumen entries={programaEntries.filter((e) => ["linea1", "linea3", "linea4", "linea5", "linea6"].includes(e.lineaKey))} fecha={values.fecha} horarios={horarios} />

      <Card title="Materiales piso planta" step={2}>
        {prevCierre ? (
          <p className="text-xs text-emerald-700 mb-3">
            Se completó automáticamente con los materiales de piso de planta del cierre {turnoLabel(prev.turno)} del {fmtFecha(prev.fecha)}. Puedes ajustarlos si es necesario.
          </p>
        ) : (
          <p className="text-xs text-slate-400 mb-3">
            Sin cierre de turno anterior registrado — se completó en 0. Ajusta si corresponde.
          </p>
        )}
        <MaterialesTable items={MATERIALES_SELECCION} values={values} setField={setField} prefix="inicioMat" />
      </Card>

      <Card title="Dotación general" step={3}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {DOTACION_GENERAL_SELECCION.map((label, i) => (
            <NumField key={i} label={label} value={values[`dg_${i}`]} onChange={(v) => setField(`dg_${i}`, v)} />
          ))}
        </div>
      </Card>

      <Card title="Líneas activas, procesos y dotación" step={4}>
        <div className="space-y-3">
          {LINEAS_SELECCION.map((linea) => (
            <LineaCard
              key={linea.key}
              linea={linea}
              values={values}
              setField={setField}
              procesosExtra={porLinea(linea.key)}
              programaEntries={programaEntries}
            />
          ))}
          <ArmadoMaterialesCard values={values} setField={setField} />
        </div>
      </Card>

      <Card title="Dotación total">
        <DataGrid items={[["Total dotación", fmtNum(totalDotacionSeleccion(values))]]} />
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-3">
          <SelectField label="¿Dotación completa?" value={values.dotacionCompleta} onChange={(v) => setField("dotacionCompleta", v)} options={SI_NO} />
        </div>
        <div className="mt-3">
          <TextAreaField label="Comentarios sobre dotación" value={values.comentariosDotacion} onChange={(v) => setField("comentariosDotacion", v)} />
        </div>
      </Card>
    </div>
  );
}

function computeSeleccionRendimiento(record) {
  const activas = LINEAS_SELECCION.filter((l) => record[`linea_${l.key}_activa`] === "Sí");
  if (activas.length === 0) return 0;
  let totalIng = 0, totalApr = 0;
  activas.forEach((l) => {
    const t = lineaKgTotales(record, l);
    totalIng += t.ing;
    totalApr += t.apr;
  });
  return totalIng ? totalApr / totalIng : 0;
}

// Cumplimiento = Kg ingresados real (cierre) / Kg programados (Programa de
// producción) para las líneas activas de Selección en ese turno específico.
function computeSeleccionCumplimiento(values, inicio, programaEntries) {
  const activas = LINEAS_SELECCION.filter((l) => (inicio?.[`linea_${l.key}_activa`] ?? values[`linea_${l.key}_activa`]) === "Sí");
  if (activas.length === 0) return { cumplimiento: 0, kgProgramado: 0, kgIngresado: 0 };
  const kgProgramado = (programaEntries || [])
    .filter((p) => activas.some((l) => l.key === (p.lineaKey || p.linea)) && p.especie !== "LAVADO")
    .reduce((s, p) => s + num(p.cantidad), 0);
  let kgIngresado = 0;
  activas.forEach((l) => { kgIngresado += lineaKgTotales(values, l).ing; });
  const cumplimiento = kgProgramado ? kgIngresado / kgProgramado : 0;
  return { cumplimiento, kgProgramado, kgIngresado };
}

// Para una línea activa, devuelve los slots de especie que tienen un proceso
// elegido en el Inicio de turno (hasta 5). Si no hay ninguno (registros
// antiguos sin especie registrada), devuelve un único slot sin nombre para
// no perder los datos ya cargados.
function especiesActivasLinea(l, inicio, values) {
  const slots = ESPECIE_SLOTS
    .map((s) => {
      const campo = `linea_${l.key}_proceso${sufijoEspecie(s)}`;
      const proceso = (inicio?.[campo] ?? values[campo]) || "";
      return { s, proceso };
    })
    // Las especies que el supervisor quitó en el Cierre no se consideran.
    .filter((x) => x.proceso && !values?.[`quitada_${l.key}_e${x.s}`]);
  return slots.length > 0 ? slots : [{ s: 1, proceso: "—" }];
}

// Hasta 3 filas de "Tipo de aprobado" (nombre desde lista + Kg) por especie.
// Componente de módulo (no anidar) para que React no pierda su identidad.
function KgAprobadoRows({ lineaKey, slot, values, setField, tipos: especiesDisponibles }) {
  const filasActivas = APROBADO_SLOTS.filter((t) => {
    const tipo = values[`kg_${lineaKey}_e${slot}_apr_t${t}_tipo`];
    const kg = values[`kg_${lineaKey}_e${slot}_apr_t${t}_kg`];
    return (tipo !== undefined && tipo !== "") || (kg !== undefined && kg !== "");
  });
  // Si no hay ninguna fila cargada todavía, muestra 1 fila vacía para empezar.
  const mostrar = filasActivas.length > 0 ? filasActivas : [1];
  const puedeAgregar = mostrar.length < APROBADO_SLOTS.length;

  const agregarFila = () => {
    const siguiente = APROBADO_SLOTS.find((t) => !mostrar.includes(t));
    if (siguiente) setField(`kg_${lineaKey}_e${slot}_apr_t${siguiente}_kg`, "0");
  };
  const quitarFila = (t) => {
    setField(`kg_${lineaKey}_e${slot}_apr_t${t}_tipo`, "");
    setField(`kg_${lineaKey}_e${slot}_apr_t${t}_kg`, "");
  };

  return (
    <div className="space-y-2">
      <label className="block text-xs font-medium text-slate-600">Kg aprobados (hasta 3 especies)</label>
      {mostrar.map((t) => (
        <div key={t} className="flex items-end gap-2">
          <div className="flex-1">
            <SelectField
              label={`Especie ${t}`}
              value={values[`kg_${lineaKey}_e${slot}_apr_t${t}_tipo`] || ""}
              onChange={(v) => setField(`kg_${lineaKey}_e${slot}_apr_t${t}_tipo`, v)}
              options={especiesDisponibles}
            />
          </div>
          <div className="flex-1">
            <NumField
              label="Kg"
              value={values[`kg_${lineaKey}_e${slot}_apr_t${t}_kg`]}
              onChange={(v) => setField(`kg_${lineaKey}_e${slot}_apr_t${t}_kg`, v)}
            />
          </div>
          {mostrar.length > 1 && (
            <button type="button" onClick={() => quitarFila(t)} aria-label="Quitar especie"
              className="mb-0.5 p-2 rounded-lg hover:bg-red-50 text-slate-400 hover:text-red-600 transition-colors">
              <Trash2 size={16} />
            </button>
          )}
        </div>
      ))}
      {puedeAgregar && (
        <button type="button" onClick={agregarFila}
          className="text-xs flex items-center gap-1 text-blue-600 hover:text-blue-700 font-medium">
          <Plus size={14} /> Agregar otra especie aprobada
        </button>
      )}
    </div>
  );
}

function SeleccionCierre({ values, setField, inicio, programaEntries }) {
  const activas = LINEAS_SELECCION.filter((l) => (inicio?.[`linea_${l.key}_activa`] ?? values[`linea_${l.key}_activa`]) === "Sí");
  const rendimientoTotal = computeSeleccionRendimiento({ ...(inicio || {}), ...values });
  const { cumplimiento, kgProgramado, kgIngresado } = computeSeleccionCumplimiento(values, inicio, programaEntries);
  const { porLinea } = useProcesosExtra();

  // Detección de discrepancias entre lo activado en Inicio de turno y lo que
  // el Programa de producción realmente tiene asignado para ESTE turno
  // específico. Esto detecta, por ejemplo, cuando el Programa se cambió de
  // una línea a otra (ej. Línea 5 → Línea 6) pero el Inicio de turno no se
  // volvió a editar para reflejar el cambio.
  const progPorLinea = (lk) => (programaEntries || []).filter(
    (p) => (p.lineaKey === lk || p.linea === lk) && p.especie !== "LAVADO"
  );
  const activasSinPrograma = activas.filter((l) => progPorLinea(l.key).length === 0);
  const programadasSinActivar = LINEAS_SELECCION.filter(
    (l) => !activas.some((a) => a.key === l.key) && progPorLinea(l.key).length > 0
  );

  return (
    <div className="space-y-4">
      {(activasSinPrograma.length > 0 || programadasSinActivar.length > 0) && (
        <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3 space-y-2">
          <div className="flex items-start gap-2">
            <AlertTriangle size={18} className="text-amber-600 shrink-0 mt-0.5" />
            <div className="text-sm font-semibold text-amber-800">Revisa las líneas activas de este turno</div>
          </div>
          {activasSinPrograma.map((l) => (
            <p key={l.key} className="text-xs text-amber-700 pl-6">
              <b>{l.label}</b> está activa (desde el Inicio de turno) pero el Programa de producción no tiene nada asignado a esta línea en este turno. Si el Programa cambió a otra línea, vuelve al Inicio de turno y corrígelo.
            </p>
          ))}
          {programadasSinActivar.map((l) => (
            <p key={l.key} className="text-xs text-amber-700 pl-6">
              El Programa tiene <b>{l.label}</b> asignada para este turno, pero no está marcada como activa en el Inicio de turno. Actívala en el Inicio para que aparezca aquí.
            </p>
          ))}
        </div>
      )}

      <Card title="Materiales piso planta (fin de turno)" step={1}>
        <MaterialesTable items={MATERIALES_SELECCION} values={values} setField={setField} prefix="finMat" />
      </Card>

      {activas.length > 0 && (
        <Card title="Resumen del turno">
          <div className="grid grid-cols-2 gap-3">
            <div className="border border-emerald-200 bg-emerald-50 rounded-xl p-3 text-center">
              <div className="text-xs text-slate-500 mb-1">Rendimiento total</div>
              <div className="text-2xl font-bold text-emerald-700">{fmtPct(rendimientoTotal)}</div>
              <div className="text-xs text-slate-400 mt-0.5">Kg aprobados / Kg ingresados</div>
            </div>
            <div className="border border-blue-200 bg-blue-50 rounded-xl p-3 text-center">
              <div className="text-xs text-slate-500 mb-1">Cumplimiento</div>
              <div className="text-2xl font-bold text-blue-700">{kgProgramado ? fmtPct(cumplimiento) : "—"}</div>
              <div className="text-xs text-slate-400 mt-0.5">
                {kgProgramado ? `${fmtNum(kgIngresado, 0)} / ${fmtNum(kgProgramado, 0)} Kg` : "Sin Kg programado"}
              </div>
            </div>
          </div>
        </Card>
      )}

      <Card title="Cierre de procesos por línea" step={2}>
        {activas.length === 0 ? (
          <EmptyNote text="No se marcaron líneas activas en el inicio de turno." />
        ) : (
          <div className="space-y-4">
            {activas.map((l) => {
              const especies = especiesActivasLinea(l, inicio, values);
              // Especies asociadas a esta línea (para elegir el nombre de cada
              // fila de Kg aprobados) — incluye las extra agregadas por el Jefe.
              const procesosCompletos = procesosConExtra(l.procesos || [], porLinea(l.key));
              const especiesLinea = especiesUnicas(procesosCompletos);
              return (
                <div key={l.key} className="border border-slate-300 rounded-lg p-3 space-y-3">
                  <div className="flex items-center justify-between">
                    <div className="text-sm font-semibold text-slate-900">{l.label}</div>
                    {ESPECIE_SLOTS.some((s) => values[`quitada_${l.key}_e${s}`]) && (
                      <button type="button" onClick={() => ESPECIE_SLOTS.forEach((s) => { if (values[`quitada_${l.key}_e${s}`]) setField(`quitada_${l.key}_e${s}`, false); })}
                        className="text-xs text-blue-600 underline">Restaurar especies quitadas</button>
                    )}
                  </div>
                  {especies.map(({ s, proceso }) => {
                    const ing = num(values[`kg_${l.key}_e${s}_ing`]);
                    const apr = kgAprobadoTotal(values, l.key, s);
                    const rend = ing ? apr / ing : 0;
                    return (
                      <div key={s} className="border border-slate-200 bg-slate-50 rounded-lg p-2.5 space-y-2">
                        <div className="flex justify-between items-start gap-2 text-xs">
                          {especies.length > 1 && <span className="font-semibold text-slate-500">Especie {s}</span>}
                          <span className="text-slate-700 font-medium text-right flex-1">{proceso}</span>
                          {especies.length > 1 && (
                            <button type="button" aria-label={`Quitar especie ${s}`} onClick={() => {
                              setField(`quitada_${l.key}_e${s}`, true);
                              setField(`kg_${l.key}_e${s}_ing`, "");
                              APROBADO_SLOTS.forEach((t) => { setField(`kg_${l.key}_e${s}_apr_t${t}_tipo`, ""); setField(`kg_${l.key}_e${s}_apr_t${t}_kg`, ""); });
                            }} className="flex items-center gap-1 font-medium text-rose-600 border border-rose-200 rounded-md px-2 py-0.5 hover:bg-rose-50">
                              <Trash2 size={12} /> Quitar
                            </button>
                          )}
                        </div>
                        <NumField label="Kg ingresados" value={values[`kg_${l.key}_e${s}_ing`]} onChange={(v) => setField(`kg_${l.key}_e${s}_ing`, v)} />
                        <KgAprobadoRows lineaKey={l.key} slot={s} values={values} setField={setField} tipos={especiesLinea} />
                        <p className="text-sm text-slate-700 pt-1 border-t border-slate-200">
                          Kg aprobados total: <span className="font-semibold">{fmtNum(apr, 1)}</span> · Rendimiento: <span className="font-semibold">{fmtPct(rend)}</span>
                        </p>
                      </div>
                    );
                  })}
                  {especies.length > 1 && (
                    <p className="text-xs text-slate-500 pt-1 border-t border-slate-200">
                      Total línea: {lineaKgTotales(values, l).ing} Kg ing. · {lineaKgTotales(values, l).apr} Kg apr. · Rend. {fmtPct(lineaKgTotales(values, l).ing ? lineaKgTotales(values, l).apr / lineaKgTotales(values, l).ing : 0)}
                    </p>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </Card>

      <ComentariosFields values={values} setField={setField} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// ENVASADO
// ---------------------------------------------------------------------------
function skuProducto(sku) {
  const found = ENVASADORA_SKUS.find(([s]) => s === sku);
  return found ? found[1] : "";
}

function CodificacionCard({ sku }) {
  if (!sku) return null;
  const mat = skuMaterial(sku);
  if (!mat) return null;
  const cod = mat.cliente ? CODIFICACION_POR_CLIENTE[mat.cliente] : null;
  return (
    <div className="border border-amber-300 bg-amber-50 rounded-lg p-3 space-y-2 text-sm">
      <p className="text-xs font-semibold text-amber-800 uppercase tracking-wide">Materiales y codificación — {sku}</p>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
        <div>
          <p className="text-xs text-slate-500">Bolsa</p>
          <p className="font-medium text-slate-900">{mat.codBolsa || "—"}{mat.nomBolsa ? ` · ${mat.nomBolsa}` : ""}</p>
        </div>
        <div>
          <p className="text-xs text-slate-500">Caja</p>
          <p className="font-medium text-slate-900">{mat.codCaja || "—"}{mat.nomCaja ? ` · ${mat.nomCaja}` : ""}</p>
        </div>
        {mat.bolsasXCaja !== "" && (
          <div>
            <p className="text-xs text-slate-500">Bolsas x Caja</p>
            <p className="font-medium text-slate-900">{mat.bolsasXCaja}</p>
          </div>
        )}
        {mat.cajasXPallet !== "" && (
          <div>
            <p className="text-xs text-slate-500">Cajas x Pallet</p>
            <p className="font-medium text-slate-900">{mat.cajasXPallet}</p>
          </div>
        )}
        {mat.tipoPallet !== "" && (
          <div>
            <p className="text-xs text-slate-500">Tipo de Pallet</p>
            <p className="font-medium text-slate-900">{mat.tipoPallet}{mat.tixhi !== "" ? ` (${mat.tixhi})` : ""}</p>
          </div>
        )}
        {mat.slipSheet !== "" && (
          <div className="sm:col-span-2">
            <p className="text-xs text-slate-500">Slip Sheet</p>
            <p className="font-medium text-slate-900 whitespace-pre-wrap">{mat.slipSheet}</p>
          </div>
        )}
      </div>
      {cod ? (
        <div className="pt-2 border-t border-amber-300 space-y-2">
          <p className="text-xs text-slate-500">Cliente: <span className="font-medium text-slate-800">{mat.cliente}</span></p>
          <div>
            <p className="text-xs text-slate-500">Codificación bolsa</p>
            <p className="text-xs text-slate-800 whitespace-pre-wrap font-mono bg-white rounded px-2 py-1 border border-amber-200">{cod.bolsa || "—"}</p>
          </div>
          <div>
            <p className="text-xs text-slate-500">Codificación caja</p>
            <p className="text-xs text-slate-800 whitespace-pre-wrap font-mono bg-white rounded px-2 py-1 border border-amber-200">{cod.caja || "—"}</p>
          </div>
        </div>
      ) : (
        <p className="text-xs text-slate-500 italic">No hay cliente/codificación asociada a este SKU en el maestro.</p>
      )}
    </div>
  );
}

function EnvasadoInicio({ values, setField, programaEntries, editingId }) {
  const mapping = [
    ...MATERIALES_ENVASADO.map((_, i) => [`finMat_${i}`, `inicioMat_${i}`]),
    ["mtcCajasSkuFin", "mtcCajasSkuInicio"],
    ["mtcBolsasSkuFin", "mtcBolsasSkuInicio"],
  ];
  const { prevCierre, prev } = useCarryOver("envasado", values, setField, editingId, mapping);
  const { horarios } = useHorariosTurno();
  const { getMat } = useSkuOverrides();

  const activaEnvasadora = values.activa_envasadora === "Sí";
  const activaLinea5 = values.activa_linea5 === "Sí";
  const especiesLinea5 = especiesUnicas(LINEA_PROCESOS.LINEA_MANUAL);

  // SKU(s) que el Jefe programó para Envasadora en este turno específico.
  const skusProgramadosTurno = useMemo(() => {
    const set = new Set();
    programaEntries
      .filter((e) => e.lineaKey === "envasadora" && e.turno === values.turno && e.especie !== "LAVADO")
      .forEach((e) => set.add(e.especie));
    return [...set];
  }, [programaEntries, values.turno]);

  // El MTC de caja y el MTC de bolsa de cada SKU programado ya están
  // definidos en la especificación del SKU (maestro de materiales), así que
  // no hace falta que el supervisor los tipee: se reconocen automáticamente
  // a partir del Programa de producción de Envasadora de este turno. Si dos
  // SKU programados comparten el mismo MTC, se deja una sola fila para ese MTC.
  const mtcCajasProgramados = useMemo(() => {
    const set = new Set();
    skusProgramadosTurno.forEach((sku) => {
      const mat = getMat(sku);
      if (mat?.codCaja) set.add(mat.codCaja);
    });
    return [...set];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [skusProgramadosTurno.join(",")]);

  const mtcBolsasProgramados = useMemo(() => {
    const set = new Set();
    skusProgramadosTurno.forEach((sku) => {
      const mat = getMat(sku);
      if (mat?.codBolsa) set.add(mat.codBolsa);
    });
    return [...set];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [skusProgramadosTurno.join(",")]);

  // Especie(s) programadas para Línea 5 en este turno.
  const especiesLinea5Prog = useMemo(() => {
    const set = new Set();
    programaEntries
      .filter((e) => e.lineaKey === "linea5" && e.turno === values.turno && e.especie !== "LAVADO")
      .forEach((e) => { const esp = especieFor(LINEA_PROCESOS.LINEA_MANUAL, e.especie); if (esp) set.add(esp); });
    return [...set];
  }, [programaEntries, values.turno]);

  // Entradas del Programa para cada línea (para mostrar "N prog." como en Selección).
  const progEnvasadora = programaEntries.filter((e) => e.lineaKey === "envasadora" && e.especie !== "LAVADO");
  const progLinea5 = programaEntries.filter((e) => (e.lineaKey === "linea5" || e.linea === "linea5") && e.especie !== "LAVADO");

  // Si hay un único SKU programado y todavía no se ha elegido ninguno, lo preselecciona.
  useEffect(() => {
    if (activaEnvasadora && !values.envasadora_sku && skusProgramadosTurno.length === 1) {
      setField("envasadora_sku", skusProgramadosTurno[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activaEnvasadora, skusProgramadosTurno.join(","), values.envasadora_sku]);

  // Si hay una única especie programada para L5 y el campo está vacío, preselecciona.
  useEffect(() => {
    if (activaLinea5 && !values.linea5_especie && especiesLinea5Prog.length === 1) {
      setField("linea5_especie", especiesLinea5Prog[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activaLinea5, especiesLinea5Prog.join(","), values.linea5_especie]);

  return (
    <div className="space-y-4">
      <Card title="Datos generales" step={1}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <TextField label="Fecha" type="date" value={values.fecha} onChange={(v) => setField("fecha", v)} />
          <TextField label="Hora de inicio" type="time" value={values.horaInicio} onChange={(v) => setField("horaInicio", v)} />
          <div className="sm:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1">Turno</label>
            <TurnoSelector value={values.turno} onChange={(v) => setField("turno", v)} />
          </div>
        </div>
      </Card>

      <ProgramaResumen entries={programaEntries.filter((e) => ["linea5", "envasadora"].includes(e.lineaKey))} fecha={values.fecha} horarios={horarios} />

      <Card title="Materiales piso planta" step={2}>
        {prevCierre ? (
          <p className="text-xs text-emerald-700 mb-3">
            Se completó automáticamente con los materiales de piso de planta del cierre {turnoLabel(prev.turno)} del {fmtFecha(prev.fecha)}. Puedes ajustarlos si es necesario.
          </p>
        ) : (
          <p className="text-xs text-slate-400 mb-3">
            Sin cierre de turno anterior registrado — se completó en 0. Ajusta si corresponde.
          </p>
        )}
        <MaterialesTable items={MATERIALES_ENVASADO} values={values} setField={setField} prefix="inicioMat" excludeIndices={[3, 4]} />
        <div className="mt-4 pt-4 border-t border-slate-200 space-y-4">
          <MtcPorSkuEditor
            title="Cantidad MTC Cajas"
            unitLabel="cajas"
            skusProgramados={mtcCajasProgramados}
            values={values}
            setField={setField}
            fieldName="mtcCajasSkuInicio"
            legacyField="inicioMat_3"
          />
          <MtcPorSkuEditor
            title="Cantidad MTC Bolsas"
            unitLabel="bolsas"
            skusProgramados={mtcBolsasProgramados}
            values={values}
            setField={setField}
            fieldName="mtcBolsasSkuInicio"
            legacyField="inicioMat_4"
          />
        </div>
      </Card>

      <Card title="Dotación general" step={3}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {DOTACION_GENERAL_ENVASADO.map((label, i) => (
            <NumField key={i} label={label} value={values[`dg_${i}`]} onChange={(v) => setField(`dg_${i}`, v)} />
          ))}
        </div>
      </Card>

      <Card title="Líneas de trabajo" step={4}>
        <p className="text-xs text-slate-600 mb-3">Puedes activar Envasadora, Línea 5, o ambas.</p>
        <div className="space-y-3">
          <div className="border border-slate-300 rounded-lg p-3 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-sm font-semibold text-slate-900">Envasadora</span>
              <div className="flex items-center gap-2">
              {progEnvasadora.length > 0 && !activaEnvasadora && (
                <span className="text-xs bg-blue-100 text-blue-700 border border-blue-200 rounded-full px-2 py-0.5">{progEnvasadora.length} prog.</span>
              )}
              <button
                type="button"
                onClick={() => setField("activa_envasadora", activaEnvasadora ? "No" : "Sí")}
                className={`text-xs font-medium rounded-full px-3 py-1 border transition-colors ${
                  activaEnvasadora ? "bg-blue-500/10 border-blue-400 text-blue-700" : "border-slate-400 text-slate-600"
                }`}
              >
                {activaEnvasadora ? "Activa" : "Inactiva"}
              </button>
              </div>
            </div>
            {activaEnvasadora && (
              <div className="space-y-2">
                {progEnvasadora.length > 0 ? (
                  <div className="bg-blue-50 border border-blue-200 rounded-lg px-3 py-2 text-xs text-blue-800">✓ SKU cargado desde el Programa de producción.</div>
                ) : (
                  <div className="bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-800">⚠ Esta línea está activa pero el Programa de producción no tiene nada asignado para este turno. Verifica si el Programa cambió a otra línea.</div>
                )}
                <SkuPicker
                  label="Código SKU"
                  value={values.envasadora_sku}
                  onChange={(v) => setField("envasadora_sku", v)}
                  sugeridos={skusProgramadosTurno}
                />
                <CodificacionCard sku={values.envasadora_sku} />
                <div className="pt-2 border-t border-slate-300">
                  <p className="text-xs font-semibold text-slate-600 uppercase tracking-wide mb-2">Dotación de Envasadora</p>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                    {DOTACION_LINEA.map((label, i) => (
                      <NumField key={i} label={label} value={values[`envasadora_dot${i}`]} onChange={(v) => setField(`envasadora_dot${i}`, v)} />
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          <div className="border border-slate-300 rounded-lg p-3 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-sm font-semibold text-slate-900">Línea 5</span>
              <div className="flex items-center gap-2">
              {progLinea5.length > 0 && !activaLinea5 && (
                <span className="text-xs bg-blue-100 text-blue-700 border border-blue-200 rounded-full px-2 py-0.5">{progLinea5.length} prog.</span>
              )}
              <button
                type="button"
                onClick={() => setField("activa_linea5", activaLinea5 ? "No" : "Sí")}
                className={`text-xs font-medium rounded-full px-3 py-1 border transition-colors ${
                  activaLinea5 ? "bg-blue-500/10 border-blue-400 text-blue-700" : "border-slate-400 text-slate-600"
                }`}
              >
                {activaLinea5 ? "Activa" : "Inactiva"}
              </button>
              </div>
            </div>
            {activaLinea5 && (
              <div className="space-y-2">
                {progLinea5.length === 0 && (
                  <div className="bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-800">⚠ Esta línea está activa pero el Programa de producción no tiene nada asignado para este turno. Verifica si el Programa cambió a otra línea.</div>
                )}
                {especiesLinea5Prog.length > 0 && (
                  <div className="flex flex-wrap gap-1.5 mb-1">
                    {especiesLinea5Prog.map((esp) => (
                      <button key={esp} type="button" onClick={() => setField("linea5_especie", esp)}
                        className={`text-xs rounded-full px-3 py-1 border font-medium transition-colors ${values.linea5_especie === esp ? "bg-violet-600 text-white border-violet-600" : "bg-violet-50 text-violet-800 border-violet-300 hover:bg-violet-100"}`}>
                        ✓ {esp} <span className="opacity-60">(programado)</span>
                      </button>
                    ))}
                  </div>
                )}
                <SelectField label="Especie" value={values.linea5_especie} onChange={(v) => setField("linea5_especie", v)} options={especiesLinea5} />
                <div className="pt-2 border-t border-slate-300">
                  <p className="text-xs font-semibold text-slate-600 uppercase tracking-wide mb-2">Dotación de Línea 5</p>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                    {DOTACION_LINEA.map((label, i) => (
                      <NumField key={i} label={label} value={values[`linea5_dot${i}`]} onChange={(v) => setField(`linea5_dot${i}`, v)} />
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </Card>

      <Card title="Dotación total">
        <DataGrid items={[["Total dotación", fmtNum(dotacionEnvasadoTotal(values))]]} />
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-3">
          <SelectField label="¿Dotación completa?" value={values.dotacionCompleta} onChange={(v) => setField("dotacionCompleta", v)} options={SI_NO} />
        </div>
        <div className="mt-3">
          <TextAreaField label="Comentarios sobre dotación" value={values.comentariosDotacion} onChange={(v) => setField("comentariosDotacion", v)} />
        </div>
      </Card>
    </div>
  );
}

function dotacionEnvasadoTotal(values) {
  let t = DOTACION_GENERAL_ENVASADO.reduce((s, _, i) => s + num(values[`dg_${i}`]), 0);
  if (values.activa_envasadora === "Sí") t += DOTACION_LINEA.reduce((s, _, i) => s + num(values[`envasadora_dot${i}`]), 0);
  if (values.activa_linea5 === "Sí") t += DOTACION_LINEA.reduce((s, _, i) => s + num(values[`linea5_dot${i}`]), 0);
  return t;
}

function computeEnvasadoMetrics(values) {
  const cumplimiento = num(values.cajasProgramadas) ? num(values.cajasProducidas) / num(values.cajasProgramadas) : 0;
  const mermaCajas = num(values.cajasTeoricas) ? (num(values.cajasConsumidasReal) - num(values.cajasTeoricas)) / num(values.cajasTeoricas) : 0;
  const mermaBolsas = num(values.bolsasTeoricas) ? (num(values.bolsasConsumidasReal) - num(values.bolsasTeoricas)) / num(values.bolsasTeoricas) : 0;
  // Línea 5 ahora usa el mismo esquema que las líneas de Selección (Kg
  // ingresados + hasta 3 tipos de "Kg aprobados"). Si el registro es antiguo
  // y no tiene datos en ese esquema, cae de vuelta a los campos legacy
  // l5_kgIngresados/l5_kgAprobados para no perder el histórico.
  const l5Totales = lineaKgTotales(values, { key: "linea5" });
  let l5Ing = l5Totales.ing;
  let l5Apr = l5Totales.apr;
  if (!l5Ing && !l5Apr) {
    l5Ing = num(values.l5_kgIngresados);
    l5Apr = num(values.l5_kgAprobados);
  }
  const rendimientoLinea5 = l5Ing ? l5Apr / l5Ing : 0;
  return { cumplimiento, mermaCajas, mermaBolsas, rendimientoLinea5, l5Ing, l5Apr };
}

// Suma los minutos de todas las filas de detenciones registradas en el cierre.
function sumDetenciones(rows) {
  return (rows || []).reduce((s, r) => s + (Number(r && r.minutos) || 0), 0);
}

// Tiempo efectivo (T.E) y productividad (bolsas por minuto) del cierre de
// Envasado, siguiendo el mismo criterio que "Horarios de Turno":
//   T.E = minutos efectivos del turno (según horario configurado) − Σ minutos
//         de las detenciones registradas en este cierre.
//   Productividad (bpm) = Bolsas consumidas real / T.E.
function computeTiempoEfectivoEnvasado(values, horarios) {
  const horario = values.fecha && values.turno && horarios ? horarioParaTurno(values.fecha, values.turno, horarios) : null;
  const minutosBase = horario ? horario.minutosEfectivos : 0;
  const totalDetenciones = sumDetenciones(values.detenciones);
  const tiempoEfectivo = horario ? Math.max(0, minutosBase - totalDetenciones) : null;
  const productividad = tiempoEfectivo ? num(values.bolsasConsumidasReal) / tiempoEfectivo : 0;
  return { horario, minutosBase, totalDetenciones, tiempoEfectivo, productividad };
}

// Editor de la lista de "Detenciones" del cierre de Envasado: motivo + minutos
// por fila, usada para calcular el Tiempo Efectivo y la Productividad (bpm).
function DetencionesEditor({ values, setField }) {
  const rows = values.detenciones || [];
  const updateRow = (i, key, val) => {
    const next = rows.map((r, idx) => (idx === i ? { ...r, [key]: val } : r));
    setField("detenciones", next);
  };
  const addRow = () => setField("detenciones", [...rows, { motivo: "", minutos: "" }]);
  const removeRow = (i) => setField("detenciones", rows.filter((_, idx) => idx !== i));

  return (
    <div className="space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-sm font-semibold text-slate-900">Detenciones</span>
        <button type="button" onClick={addRow} className="text-xs text-emerald-700 font-medium hover:underline">
          + Agregar detención
        </button>
      </div>
      {rows.length === 0 ? (
        <p className="text-xs text-slate-400 italic">Sin detenciones registradas en este turno.</p>
      ) : (
        <div className="space-y-2">
          {rows.map((row, i) => (
            <div key={i} className="grid grid-cols-1 sm:grid-cols-[1fr_auto_auto] gap-2 items-end border border-slate-200 rounded-lg p-2">
              <TextField label="Motivo" value={row.motivo} onChange={(v) => updateRow(i, "motivo", v)} />
              <NumField label="Minutos" value={row.minutos} onChange={(v) => updateRow(i, "minutos", v)} />
              <button type="button" onClick={() => removeRow(i)} className="text-xs text-rose-600 hover:underline pb-2">
                Quitar
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function EnvasadoCierre({ values, setField, inicio, programaEntries }) {
  const { horarios } = useHorariosTurno();
  const { getMat } = useSkuOverrides();
  const activaEnvasadora = (inicio?.activa_envasadora ?? values.activa_envasadora) === "Sí";
  const activaLinea5 = (inicio?.activa_linea5 ?? values.activa_linea5) === "Sí";
  const especieLinea5 = inicio?.linea5_especie || values.linea5_especie || "";
  let procesosLinea5 = LINEA_PROCESOS.LINEA_MANUAL.filter(([, especie]) => especie === especieLinea5).map((p) => p[0]);
  if (procesosLinea5.length === 0) procesosLinea5 = LINEA_PROCESOS.LINEA_MANUAL.map((p) => p[0]);
  const m = computeEnvasadoMetrics(values);
  const te = computeTiempoEfectivoEnvasado(values, horarios);
  const cierreStep2 = activaLinea5 ? 2 : null;
  const cierreStep3 = activaEnvasadora ? (activaLinea5 ? 3 : 2) : null;

  // SKU(s) que el Jefe programó para Envasadora en este turno específico.
  const skusProgramadosTurno = useMemo(() => {
    const set = new Set();
    (programaEntries || [])
      .filter((p) => p.lineaKey === "envasadora" && p.especie !== "LAVADO")
      .forEach((p) => set.add(p.especie));
    return [...set];
  }, [programaEntries]);

  // Si hay un único SKU programado y todavía no se ha elegido ninguno, lo preselecciona.
  useEffect(() => {
    if (activaEnvasadora && !values.envasadora_sku && skusProgramadosTurno.length === 1) {
      setField("envasadora_sku", skusProgramadosTurno[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activaEnvasadora, skusProgramadosTurno.join(","), values.envasadora_sku]);

  // El MTC de caja y el MTC de bolsa de cada SKU programado ya están
  // definidos en la especificación del SKU (maestro de materiales) — se usan
  // como respaldo si el Inicio de turno no llegó a registrar ninguno.
  const mtcCajasProgramados = useMemo(() => {
    const set = new Set();
    skusProgramadosTurno.forEach((sku) => {
      const mat = getMat(sku);
      if (mat?.codCaja) set.add(mat.codCaja);
    });
    return [...set];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [skusProgramadosTurno.join(",")]);

  const mtcBolsasProgramados = useMemo(() => {
    const set = new Set();
    skusProgramadosTurno.forEach((sku) => {
      const mat = getMat(sku);
      if (mat?.codBolsa) set.add(mat.codBolsa);
    });
    return [...set];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [skusProgramadosTurno.join(",")]);

  // Los MTC de Cajas/Bolsas que ya se registraron en el Inicio de turno de
  // ESTE mismo turno: el Cierre ya sabe qué MTC están asignados y parte con
  // los mismos, sin que haya que volver a elegirlos. Solo si el Inicio no
  // registró ninguno se recurre, como respaldo, a los MTC que la
  // especificación del SKU define para lo programado en este turno.
  const skusCajasIniciales = useMemo(() => {
    const desdeInicio = (Array.isArray(inicio?.mtcCajasSkuInicio) ? inicio.mtcCajasSkuInicio : []).map((r) => r.sku).filter(Boolean);
    return desdeInicio.length > 0 ? desdeInicio : mtcCajasProgramados;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [inicio, mtcCajasProgramados.join(",")]);

  const skusBolsasIniciales = useMemo(() => {
    const desdeInicio = (Array.isArray(inicio?.mtcBolsasSkuInicio) ? inicio.mtcBolsasSkuInicio : []).map((r) => r.sku).filter(Boolean);
    return desdeInicio.length > 0 ? desdeInicio : mtcBolsasProgramados;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [inicio, mtcBolsasProgramados.join(",")]);

  // Cajas programadas = lo que el Programa de producción indica para ESTE SKU en este turno.
  // (Si hay varios SKU programados en el mismo turno, cada cierre registra uno; cambiar el
  // SKU arriba recalcula automáticamente cuánto correspondía a ese SKU específico.)
  const cajasProgramadasCalc = (programaEntries || [])
    .filter((p) => p.lineaKey === "envasadora" && p.especie === values.envasadora_sku)
    .reduce((s, p) => s + num(p.cantidad), 0);

  const matSku = values.envasadora_sku ? skuMaterial(values.envasadora_sku) : null;
  const bolsasXCaja = matSku ? num(matSku.bolsasXCaja) : 0;

  // "Teórico" = lo que el Programa planificó para este SKU/turno (no lo realmente producido).
  // Así el cierre puede comparar Programado (teórico) vs. Real (lo que efectivamente se produjo/consumió).
  const cajasTeoricasCalc  = cajasProgramadasCalc;
  const bolsasTeoricasCalc = cajasProgramadasCalc * bolsasXCaja;

  // Sincroniza los campos calculados con el registro cada vez que cambian sus insumos,
  // para que queden guardados en el cierre (el supervisor puede seguir ajustándolos a mano).
  useEffect(() => {
    if (!activaEnvasadora) return;
    if (num(values.cajasProgramadas) !== cajasProgramadasCalc) setField("cajasProgramadas", cajasProgramadasCalc || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activaEnvasadora, cajasProgramadasCalc]);

  useEffect(() => {
    if (!activaEnvasadora) return;
    if (num(values.cajasTeoricas) !== cajasTeoricasCalc) setField("cajasTeoricas", cajasTeoricasCalc || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activaEnvasadora, cajasTeoricasCalc]);

  useEffect(() => {
    if (!activaEnvasadora) return;
    if (num(values.bolsasTeoricas) !== bolsasTeoricasCalc) setField("bolsasTeoricas", bolsasTeoricasCalc || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activaEnvasadora, bolsasTeoricasCalc]);

  // Kg programados para Línea 5 en este turno (mismo criterio que el
  // Cumplimiento por línea de Selección) y el cumplimiento resultante.
  const kgProgLinea5 = (programaEntries || [])
    .filter((p) => (p.lineaKey === "linea5" || p.linea === "linea5") && p.especie !== "LAVADO")
    .reduce((s, p) => s + num(p.cantidad), 0);
  const cumplimientoLinea5 = kgProgLinea5 ? m.l5Ing / kgProgLinea5 : 0;

  // Igual que en Selección: qué líneas quedaron activas en el Inicio de turno y
  // si el Programa de producción de este turno coincide con ellas.
  const progEnv = (programaEntries || []).filter((p) => p.lineaKey === "envasadora" && p.especie !== "LAVADO");
  const progL5 = (programaEntries || []).filter((p) => (p.lineaKey === "linea5" || p.linea === "linea5") && p.especie !== "LAVADO");
  const avisos = [];
  if (inicio && activaEnvasadora && progEnv.length === 0) avisos.push("Envasadora está activa (desde el Inicio de turno) pero el Programa de producción no tiene nada asignado a Envasadora en este turno.");
  if (inicio && activaLinea5 && progL5.length === 0) avisos.push("Línea 5 está activa (desde el Inicio de turno) pero el Programa de producción no tiene nada asignado a Línea 5 en este turno.");
  if (inicio && !activaEnvasadora && progEnv.length > 0) avisos.push("El Programa tiene Envasadora asignada para este turno, pero no está marcada como activa en el Inicio de turno. Actívala en el Inicio para que aparezca aquí.");
  if (inicio && !activaLinea5 && progL5.length > 0) avisos.push("El Programa tiene Línea 5 asignada para este turno, pero no está marcada como activa en el Inicio de turno. Actívala en el Inicio para que aparezca aquí.");

  return (
    <div className="space-y-4">
      {avisos.length > 0 && (
        <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3 space-y-2">
          <div className="flex items-start gap-2">
            <AlertTriangle size={18} className="text-amber-600 shrink-0 mt-0.5" />
            <div className="text-sm font-semibold text-amber-800">Revisa las líneas activas de este turno</div>
          </div>
          {avisos.map((a, i) => <p key={i} className="text-xs text-amber-700 pl-6">{a}</p>)}
        </div>
      )}

      <Card title="Líneas activas del turno">
        <div className="flex flex-wrap gap-2">
          {[["Envasadora", activaEnvasadora], ["Línea 5", activaLinea5]].map(([nombre, on]) => (
            <span key={nombre} className={`text-xs font-medium rounded-full px-3 py-1 border ${on ? "bg-emerald-50 border-emerald-300 text-emerald-700" : "bg-slate-100 border-slate-300 text-slate-500"}`}>
              {on ? "✓" : "—"} {nombre} · {on ? "Activa" : "Inactiva"}
            </span>
          ))}
        </div>
        {!inicio && <p className="text-xs text-amber-700 mt-2">No hay inicio de turno registrado para esta fecha y turno, por eso no se conocen las líneas activas.</p>}
      </Card>

      <Card title="Materiales piso planta (fin de turno)" step={1}>
        <MaterialesTable items={MATERIALES_ENVASADO} values={values} setField={setField} prefix="finMat" excludeIndices={[3, 4]} />
        <div className="mt-4 pt-4 border-t border-slate-200 space-y-4">
          <MtcPorSkuEditor
            title="Cantidad MTC Cajas"
            unitLabel="cajas"
            skusProgramados={skusCajasIniciales}
            values={values}
            setField={setField}
            fieldName="mtcCajasSkuFin"
            legacyField="finMat_3"
            referencia={inicio?.mtcCajasSkuInicio}
          />
          <MtcPorSkuEditor
            title="Cantidad MTC Bolsas"
            unitLabel="bolsas"
            skusProgramados={skusBolsasIniciales}
            values={values}
            setField={setField}
            fieldName="mtcBolsasSkuFin"
            legacyField="finMat_4"
            referencia={inicio?.mtcBolsasSkuInicio}
          />
        </div>
      </Card>

      {!activaEnvasadora && !activaLinea5 && (
        <Card><EmptyNote text="Activa Envasadora y/o Línea 5 en el inicio de turno para ver el cierre correspondiente." /></Card>
      )}

      {activaLinea5 && (
        <Card title="Cierre de proceso — Línea 5" step={cierreStep2}>
          <div className="mb-3">
            <SelectField label="Proceso" value={values.l5_proceso} onChange={(v) => setField("l5_proceso", v)} options={procesosLinea5} />
          </div>
          {/* Mismo formato que el cierre de las líneas de Selección: Kg ingresados,
              hasta 3 tipos de Kg aprobados, total, rendimiento y cumplimiento. */}
          <div className="border border-slate-200 bg-slate-50 rounded-lg p-2.5 space-y-2">
            <div className="flex justify-between text-xs">
              <span className="text-slate-700 font-medium">{especieLinea5 || "—"}</span>
            </div>
            <NumField label="Kg ingresados" value={values.kg_linea5_e1_ing} onChange={(v) => setField("kg_linea5_e1_ing", v)} />
            <KgAprobadoRows lineaKey="linea5" slot={1} values={values} setField={setField} tipos={especieLinea5 ? [especieLinea5] : []} />
            <p className="text-sm text-slate-700 pt-1 border-t border-slate-200">
              Kg aprobados total: <span className="font-semibold">{fmtNum(m.l5Apr, 1)}</span> · Rendimiento: <span className="font-semibold">{fmtPct(m.rendimientoLinea5)}</span>
            </p>
          </div>
          <p className="text-sm text-slate-700 mt-2">
            Cumplimiento:{" "}
            <span className="font-semibold">
              {kgProgLinea5 ? `${fmtPct(cumplimientoLinea5)} (${fmtNum(m.l5Ing, 0)}/${fmtNum(kgProgLinea5, 0)} Kg)` : "Sin Kg programado"}
            </span>
          </p>
        </Card>
      )}

      {activaEnvasadora && (
        <Card title="Envasadora" step={cierreStep3}>
          <div className="mb-3">
            <SkuPicker
              label="Código SKU"
              value={values.envasadora_sku}
              onChange={(v) => setField("envasadora_sku", v)}
              sugeridos={skusProgramadosTurno}
            />
          </div>
          <div className="mb-3">
            <CodificacionCard sku={values.envasadora_sku} />
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <NumField label="Cajas producidas" value={values.cajasProducidas} onChange={(v) => setField("cajasProducidas", v)} />
            <div>
              <NumField label="Cajas programadas" value={values.cajasProgramadas} onChange={(v) => setField("cajasProgramadas", v)} />
              <p className="text-xs text-slate-500 mt-1">
                {values.envasadora_sku
                  ? `Calculado automático desde el Programa para este SKU (${cajasProgramadasCalc}). Puedes ajustarlo si es necesario.`
                  : "Elige un código SKU arriba para traer la cantidad programada."}
              </p>
            </div>
            <div>
              <NumField label="Cajas teóricas consumo" value={values.cajasTeoricas} onChange={(v) => setField("cajasTeoricas", v)} />
              <p className="text-xs text-slate-500 mt-1">Calculado automático = Cajas programadas para este SKU ({cajasTeoricasCalc}).</p>
            </div>
            <NumField label="Cajas consumidas real" value={values.cajasConsumidasReal} onChange={(v) => setField("cajasConsumidasReal", v)} />
            <div>
              <NumField label="Bolsas teóricas consumo" value={values.bolsasTeoricas} onChange={(v) => setField("bolsasTeoricas", v)} />
              <p className="text-xs text-slate-500 mt-1">
                {bolsasXCaja > 0
                  ? `Calculado automático = Cajas programadas × ${bolsasXCaja} bolsas/caja (${bolsasTeoricasCalc}).`
                  : "Selecciona un SKU con bolsas/caja definidas para calcularlo automático."}
              </p>
            </div>
            <NumField label="Bolsas consumidas real" value={values.bolsasConsumidasReal} onChange={(v) => setField("bolsasConsumidasReal", v)} />
          </div>
          <div className="mt-3">
            <DataGrid items={[
              ["Cumplimiento", fmtPct(m.cumplimiento)],
              ["Merma cajas", fmtPct(m.mermaCajas)],
              ["Merma bolsas", fmtPct(m.mermaBolsas)],
            ]} />
          </div>
        </Card>
      )}

      {activaEnvasadora && (
        <Card title="Tiempos y detenciones">
          <DetencionesEditor values={values} setField={setField} />
          <div className="mt-3">
            {te.horario ? (
              <DataGrid items={[
                ["T.E", `${fmtNum(te.tiempoEfectivo)} min`],
                ["Productividad", `${fmtNum(te.productividad, 2)} bpm`],
              ]} />
            ) : (
              <p className="text-xs text-slate-400 italic">
                Configura el horario de este turno (Horarios de Turno) para calcular el Tiempo Efectivo y la Productividad.
              </p>
            )}
          </div>
        </Card>
      )}

      <ComentariosFields values={values} setField={setField} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// PROGRAMA DE PRODUCCIÓN (definido por el Jefe de producción)
// ---------------------------------------------------------------------------
function ProgramaResumen({ entries, title = "Programa de producción del día", fecha, horarios }) {
  if (!entries || entries.length === 0) return null;
  const orden = TURNOS.map((t) => t.key);
  const turnosPresentes = [...new Set(entries.map((e) => e.turno))].sort(
    (a, b) => orden.indexOf(a) - orden.indexOf(b)
  );

  return (
    <Card title={title}>
      <div className="space-y-4">
        {turnosPresentes.map((turno) => {
          const delTurno = entries
            .filter((e) => e.turno === turno)
            .sort((a, b) => (a.lineaLabel || "").localeCompare(b.lineaLabel || ""));

          const horario = fecha && horarios ? horarioParaTurno(fecha, turno, horarios) : null;

          // Kg programados de líneas con unidad "Kg" (Selección/Línea 5) — excluye
          // Envasadora (Cajas) y las entradas marcadas como LAVADO.
          const kgProgramados = delTurno
            .filter((e) => e.especie !== "LAVADO" && e.unidad === "Kg")
            .reduce((s, e) => s + num(e.cantidad), 0);
          const cajasProgramadas = delTurno
            .filter((e) => e.especie !== "LAVADO" && e.unidad === "Cajas")
            .reduce((s, e) => s + num(e.cantidad), 0);

          const kgPorMin = horario?.minutosEfectivos ? kgProgramados / horario.minutosEfectivos : 0;
          const kgPorHora = kgPorMin * 60;
          const cajasPorMin = horario?.minutosEfectivos ? cajasProgramadas / horario.minutosEfectivos : 0;
          const cajasPorHora = cajasPorMin * 60;

          return (
            <div key={turno} className="space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-500 uppercase tracking-wide">{turnoLabel(turno)}</span>
                {horario && (
                  <span className="text-xs text-slate-500">
                    {horario.horaInicio}–{horario.horaFin} · <b className="text-emerald-700">{horario.minutosEfectivos} min efectivos</b>
                  </span>
                )}
              </div>

              {!horario && fecha && (
                <p className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-2.5 py-1.5">
                  Sin horario configurado para {diaSemanaDeFecha(fecha)} · {turnoLabel(turno)}. Configúralo en el menú Jefe → Horarios de Turno.
                </p>
              )}

              {horario && kgProgramados > 0 && (
                <div className="grid grid-cols-2 gap-2">
                  <div className="border border-blue-200 bg-blue-50 rounded-lg p-2 text-center">
                    <div className="text-xs text-slate-500">Kg/hora necesario</div>
                    <div className="text-lg font-bold text-blue-700">{fmtNum(kgPorHora, 1)}</div>
                  </div>
                  <div className="border border-blue-200 bg-blue-50 rounded-lg p-2 text-center">
                    <div className="text-xs text-slate-500">Kg/min necesario</div>
                    <div className="text-lg font-bold text-blue-700">{fmtNum(kgPorMin, 2)}</div>
                  </div>
                </div>
              )}
              {horario && cajasProgramadas > 0 && (
                <div className="grid grid-cols-2 gap-2">
                  <div className="border border-amber-200 bg-amber-50 rounded-lg p-2 text-center">
                    <div className="text-xs text-slate-500">Cajas/hora necesario</div>
                    <div className="text-lg font-bold text-amber-700">{fmtNum(cajasPorHora, 1)}</div>
                  </div>
                  <div className="border border-amber-200 bg-amber-50 rounded-lg p-2 text-center">
                    <div className="text-xs text-slate-500">Cajas/min necesario</div>
                    <div className="text-lg font-bold text-amber-700">{fmtNum(cajasPorMin, 2)}</div>
                  </div>
                </div>
              )}

              <div className="space-y-1.5">
                {delTurno.map((e) => (
                  <div key={e.id} className="flex justify-between text-sm border-b border-slate-200/70 pb-1 last:border-0 last:pb-0">
                    <span className="text-slate-700">{e.lineaLabel}</span>
                    {e.especie === "LAVADO" ? (
                      <span className="text-yellow-700 font-semibold text-right">LAVADO</span>
                    ) : (
                      <span className="text-slate-900 font-medium text-right">{procesoDisplay(e)} · {fmtNum(num(e.cantidad))} {e.unidad}</span>
                    )}
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Resumen EJECUTIVO del Programa semanal para compartir por WhatsApp.
// Agrupa por día (Lunes a Domingo) y turno, con totales de Kg y Cajas al
// cierre de cada día, y un resumen general al final de la semana.
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Resumen EJECUTIVO de UN SOLO DÍA para compartir por WhatsApp — usado para
// reenviar el programa de hoy (si hubo alguna modificación) o para adelantar
// el programa de mañana. Incluye las observaciones registradas para ese día.
// ---------------------------------------------------------------------------
function buildResumenProgramaDia(fecha, records, observaciones, programaLineas) {
  try {
    const lineaLabel = (key) => programaLineas.find((l) => l.key === key)?.label || key;
    const entriesDia = (records || []).filter((e) => e && e.fecha === fecha);
    const obsDia = (observaciones || []).filter((o) => o && o.fecha === fecha);

    const diaNombre = diaSemanaDeFecha(fecha) || "";
    const lines = [
      "*PROGRAMA DE PRODUCCIÓN — DÍA*",
      `${diaNombre} ${fmtFecha(fecha)}`,
      "",
    ];

    if (entriesDia.length === 0 && obsDia.length === 0) {
      lines.push("Sin programación registrada para este día.");
      return lines.join("\n");
    }

    const turnosDia = [...new Set(entriesDia.map((e) => e.turno))].sort(
      (a, b) => TURNOS.findIndex((t) => t.key === a) - TURNOS.findIndex((t) => t.key === b)
    );

    let kgDia = 0, cajasDia = 0;

    turnosDia.forEach((turno) => {
      const delTurno = entriesDia.filter((e) => e.turno === turno);
      lines.push(`_${turnoLabel(turno)}_`);
      delTurno
        .sort((a, b) => (a.lineaLabel || lineaLabel(a.lineaKey)).localeCompare(b.lineaLabel || lineaLabel(b.lineaKey)))
        .forEach((e) => {
          const nombreLinea = e.lineaLabel || lineaLabel(e.lineaKey);
          if (e.especie === "LAVADO") {
            lines.push(`  • ${nombreLinea}: LAVADO`);
          } else {
            const cant = num(e.cantidad);
            lines.push(`  • ${nombreLinea}: ${procesoDisplay(e)} — ${fmtNum(cant)} ${e.unidad || ""}`);
            if (e.unidad === "Kg") kgDia += cant;
            else if (e.unidad === "Cajas") cajasDia += cant;
          }
        });

      const obsTurno = obsDia.find((o) => o.turno === turno);
      if (obsTurno) lines.push(`  📝 Observación: ${obsTurno.texto}`);
    });

    const totalesDia = [];
    if (kgDia > 0) totalesDia.push(`${fmtNum(kgDia)} Kg`);
    if (cajasDia > 0) totalesDia.push(`${fmtNum(cajasDia)} Cajas`);
    if (totalesDia.length) {
      lines.push("");
      lines.push(`Total del día: ${totalesDia.join(" · ")}`);
    }

    // Observaciones de turnos sin ninguna entrada de programa (ej. turno libre
    // pero con una nota igualmente) — se listan aparte al final.
    const obsSinTurnoListado = obsDia.filter((o) => !turnosDia.includes(o.turno));
    if (obsSinTurnoListado.length > 0) {
      lines.push("");
      lines.push("*Otras observaciones:*");
      obsSinTurnoListado.forEach((o) => lines.push(`  📝 ${turnoLabel(o.turno)}: ${o.texto}`));
    }

    return lines.join("\n");
  } catch (err) {
    // Nunca dejar que un dato inesperado tumbe toda la pantalla del Programa —
    // el botón de WhatsApp simplemente muestra un mensaje de respaldo.
    return `*PROGRAMA DE PRODUCCIÓN — DÍA*\n${fmtFecha(fecha)}\n\nNo se pudo generar el detalle (revisa los datos de este día).`;
  }
}

function ProgramaProduccionScreen({ onBack }) {
  const [records, saveRecords, loading] = useSharedList("programa-records");
  const [observaciones, saveObservaciones, loadingObs] = useSharedList("programa-observaciones");
  const { porLinea } = useProcesosExtra();
  const programaLineas = programaLineasConExtra(porLinea);
  const [fechaInicio, setFechaInicio] = useState(today());
  const [draft, setDraft] = useState({ fecha: today(), lineaKey: "", turno: "", especie: "", cantidad: "" });
  const [obsTexto, setObsTexto] = useState("");
  const [fechaEnvioWA, setFechaEnvioWA] = useState(today());
  const [toast, setToast] = useState(null);
  const [toastObs, setToastObs] = useState(null);
  // Con cuántos turnos se está trabajando esta semana — solo afecta qué
  // columnas se muestran en "Vista general de la semana" (no borra ni oculta
  // las entradas ya guardadas de un turno que se desmarque acá).
  const [turnosVista, setTurnosVista] = useState({ T1: true, T2: true, T3: true });

  useEffect(() => {
    setDraft((prev) => ({ ...prev, fecha: fechaInicio }));
  }, [fechaInicio]);

  const lineaCfg = programaLineas.find((l) => l.key === draft.lineaKey);
  const esLavado = draft.especie === "LAVADO";
  const esEnvasadora = draft.lineaKey === "envasadora";

  // Semana completa (Lunes a Domingo) que contiene fechaInicio, 21 turnos: T3 → T1 → T2 por cada día.
  const slots = [];
  let cursor = { fecha: mondayOfWeek(fechaInicio), turno: "T3" };
  for (let i = 0; i < 21; i++) {
    slots.push(cursor);
    cursor = turnoSiguiente(cursor.fecha, cursor.turno);
  }
  const slotKey = (s) => `${s.fecha}_${s.turno}`;
  const entriesPorSlot = {};
  records.forEach((e) => {
    const k = `${e.fecha}_${e.turno}`;
    (entriesPorSlot[k] = entriesPorSlot[k] || []).push(e);
  });
  const entriesVentana = slots.flatMap((s) => entriesPorSlot[slotKey(s)] || []);

  // Observaciones de la semana visible (mismo rango de slots que la grilla)
  const observacionesVentana = slots
    .map((s) => observaciones.find((o) => o.fecha === s.fecha && o.turno === s.turno))
    .filter(Boolean);

  // Solo para "Vista general de la semana": columnas de los turnos marcados
  // como activos arriba. El resto de la pantalla (entradas, observaciones,
  // WhatsApp) sigue considerando los 21 turnos completos.
  const slotsVista = slots.filter((s) => turnosVista[s.turno]);

  // El cuadro también se ajusta solo según las líneas realmente programadas:
  // se muestran únicamente las líneas que tienen al menos una entrada dentro
  // de los turnos visibles (evita filas vacías de líneas que no se están
  // trabajando esta semana). Si ninguna línea tiene entradas todavía, se
  // muestran todas para no dejar el cuadro vacío mientras se recién empieza
  // a cargar el programa.
  const lineasConEntradas = programaLineas.filter((linea) =>
    slotsVista.some((s) => (entriesPorSlot[slotKey(s)] || []).some((e) => e.lineaKey === linea.key))
  );
  const lineasVista = lineasConEntradas.length > 0 ? lineasConEntradas : programaLineas;

  // La observación se guarda para la fecha/turno que están seleccionados en
  // el formulario de "Agregar entrada al programa" — no es un formulario
  // aparte, así queda siempre asociada al mismo turno que se está cargando.
  const handleAddObs = async () => {
    if (!draft.fecha || !draft.turno) {
      setToastObs({ kind: "error", message: "Completa fecha y turno antes de guardar la observación." });
      return;
    }
    if (!obsTexto.trim()) {
      setToastObs({ kind: "error", message: "Escribe una observación." });
      return;
    }
    // Una sola observación por (fecha, turno) — si ya existe, la reemplaza.
    const existentes = observaciones.filter((o) => !(o.fecha === draft.fecha && o.turno === draft.turno));
    const nueva = { id: Date.now(), fecha: draft.fecha, turno: draft.turno, texto: obsTexto.trim() };
    const ok = await saveObservaciones([...existentes, nueva]);
    if (ok) {
      setObsTexto("");
      setToastObs({ kind: "ok", message: "Observación guardada." });
    } else {
      setToastObs({ kind: "error", message: "No se pudo guardar. Intenta nuevamente." });
    }
    setTimeout(() => setToastObs(null), 2500);
  };

  const handleEditObs = (o) => {
    setDraft((prev) => ({ ...prev, fecha: o.fecha, turno: o.turno }));
    setObsTexto(o.texto);
  };

  const handleDeleteObs = async (o) => {
    const ok = await saveObservaciones(observaciones.filter((x) => x.id !== o.id));
    if (!ok) setToastObs({ kind: "error", message: "No se pudo eliminar." });
  };

  const setDraftField = (key, val) => setDraft((prev) => ({ ...prev, [key]: val, ...(key === "lineaKey" ? { especie: "" } : {}) }));

  const handleAdd = async () => {
    if (!draft.fecha || !draft.lineaKey || !draft.turno || !draft.especie) {
      setToast({ kind: "error", message: "Completa fecha, línea, turno y especie antes de agregar." });
      return;
    }
    if (esEnvasadora && !esLavado && !skuMaterial(draft.especie)) {
      setToast({ kind: "error", message: "El código SKU ingresado no existe en el maestro. Verifícalo o elígelo de la lista." });
      return;
    }
    if (!esLavado && draft.cantidad === "") {
      setToast({ kind: "error", message: "Indica la cantidad a producir." });
      return;
    }
    const entry = {
      id: Date.now(),
      fecha: draft.fecha,
      lineaKey: draft.lineaKey,
      lineaLabel: lineaCfg?.label || draft.lineaKey,
      turno: draft.turno,
      especie: draft.especie,
      cantidad: esLavado ? 0 : num(draft.cantidad),
      unidad: lineaCfg?.unidad || "",
    };
    const ok = await saveRecords([...records, entry]);
    setDraft((prev) => ({ ...prev, lineaKey: "", turno: "", especie: "", cantidad: "" }));
    setToast(ok
      ? { kind: "ok", message: "Entrada agregada al programa." }
      : { kind: "error", message: "No se pudo sincronizar con el almacenamiento. Vuelve a presionar Agregar." });
  };

  const handleRemove = async (id) => {
    await saveRecords(records.filter((r) => r.id !== id));
  };

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Programa de producción" subtitle="Línea, turno, especie y cantidad a producir" onBack={onBack} icon={CalendarDays} accent="emerald" />
      <div className="p-4 space-y-4">
        <Card title="Semana del programa" step={1}>
          <label className="block text-xs font-medium text-slate-600 mb-1">Fecha de referencia (se muestra la semana completa, Lunes a Domingo)</label>
          <input type="date" style={{ colorScheme: "light" }} className={inputBase} value={fechaInicio} onChange={(e) => setFechaInicio(e.target.value)} />

          <label className="block text-xs font-medium text-slate-600 mt-3 mb-1">¿Con cuántos turnos se está trabajando esta semana?</label>
          <p className="text-xs text-slate-400 mb-2">Define qué columnas se muestran en "Vista general de la semana" — no afecta las entradas ya guardadas.</p>
          <div className="flex gap-2">
            {TURNOS.map((t) => (
              <button
                key={t.key}
                type="button"
                onClick={() => setTurnosVista((prev) => ({ ...prev, [t.key]: !prev[t.key] }))}
                className={`flex-1 text-xs font-semibold rounded-lg px-2 py-2 border transition-colors ${
                  turnosVista[t.key] ? "bg-emerald-500/10 border-emerald-400 text-emerald-700" : "border-slate-300 text-slate-500 hover:border-slate-400"
                }`}
              >
                {t.label}
              </button>
            ))}
          </div>
        </Card>

        <Card title="Enviar Programa por WhatsApp">
          <label className="block text-xs font-medium text-slate-600 mb-1">Fecha a enviar</label>
          <input
            type="date"
            style={{ colorScheme: "light" }}
            className={inputBase}
            value={fechaEnvioWA}
            onChange={(e) => setFechaEnvioWA(e.target.value)}
          />
          <a
            href={whatsappShareUrl(buildResumenProgramaDia(fechaEnvioWA, records, observaciones, programaLineas))}
            target="_blank"
            rel="noopener noreferrer"
            className="w-full mt-2 flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-xl py-3 transition-colors text-sm"
          >
            <Share2 size={18} /> Enviar Programa de {fechaEnvioWA ? fmtFechaCorta(fechaEnvioWA) : "la fecha elegida"}
          </a>
          {toast && <div className="mt-2"><Toast {...toast} /></div>}
        </Card>

        {loading ? (
          <Card><Loader /></Card>
        ) : (
          <>
            <Card title="Vista general de la semana">
              <div className="overflow-x-auto -mx-4 px-4">
                <table className="border-collapse text-[11px]" style={{ minWidth: `${90 + slotsVista.length * 104}px` }}>
                  <thead>
                    <tr>
                      <th className="sticky left-0 z-10 bg-slate-100 border border-slate-300 px-2 py-1 text-left text-slate-600 min-w-[90px]">Línea</th>
                      {slotsVista.map((s, i) => (
                        <th key={i} className="border border-slate-300 px-2 py-1 text-center min-w-[104px]">
                          <div className="text-slate-800 font-semibold">{diaCorto(s.fecha)}</div>
                          <div className="text-slate-600">{turnoLabel(s.turno)}</div>
                          <div className="text-slate-500">{fmtFechaCorta(s.fecha)}</div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {lineasVista.map((linea) => {
                      const filaEntries = slotsVista.map((s) => (entriesPorSlot[slotKey(s)] || []).filter((e) => e.lineaKey === linea.key));
                      const metas = filaEntries.map((es) => es.filter((e) => e.especie !== "LAVADO").reduce((sum, e) => sum + num(e.cantidad), 0));
                      const cc = PROGRAMA_COLOR_CLASSES[linea.color] || PROGRAMA_COLOR_CLASSES.slate;
                      return (
                        <React.Fragment key={linea.key}>
                          <tr>
                            <td className={`sticky left-0 z-10 ${cc.meta} border border-slate-300 px-2 py-1 font-semibold`}>META ({linea.unidad})</td>
                            {metas.map((m, i) => (
                              <td key={i} className={`${cc.metaCell} border border-slate-300 px-2 py-1 text-center`}>
                                {m > 0 ? fmtNum(m) : "—"}
                              </td>
                            ))}
                          </tr>
                          <tr>
                            <td className={`sticky left-0 z-10 ${cc.label} border border-slate-300 px-2 py-1 font-semibold`}>{linea.label}</td>
                            {filaEntries.map((es, i) => (
                              <td key={i} className={`${cc.cell} border border-slate-300 px-1 py-1 align-top`}>
                                {es.length === 0 ? (
                                  <span className="text-slate-400">—</span>
                                ) : (
                                  es.map((e) => (
                                    e.especie === "LAVADO" ? (
                                      <div key={e.id} className="bg-yellow-400 text-slate-900 font-bold rounded px-1 mb-0.5 text-center">LAVADO</div>
                                    ) : (
                                      <div key={e.id} className="text-slate-800 mb-0.5 leading-tight">{procesoDisplay(e)} · {fmtNum(num(e.cantidad))} {e.unidad}</div>
                                    )
                                  ))
                                )}
                              </td>
                            ))}
                          </tr>
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <p className="text-xs text-slate-500 mt-2">Mostrando {slotsVista.length} de 21 turnos de la semana (Lunes a Domingo, T3/T1/T2) — ajusta arriba en "¿Con cuántos turnos se está trabajando?" — y {lineasVista.length} de {programaLineas.length} líneas (solo las que tienen entradas programadas en los turnos visibles). Las celdas en amarillo indican que esa línea está en LAVADO durante ese turno.</p>
            </Card>

            {/* Observaciones por turno — tabla separada del schedule, no ocupa celdas de la grilla */}
            <Card title="Agregar entrada al programa" step={2}>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <TextField label="Fecha" type="date" value={draft.fecha} onChange={(v) => setDraftField("fecha", v)} />
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">Turno</label>
                  <TurnoSelector value={draft.turno} onChange={(v) => setDraftField("turno", v)} />
                </div>
                <SelectField
                  label="Línea"
                  value={draft.lineaKey}
                  onChange={(v) => setDraftField("lineaKey", v)}
                  optionsKV={programaLineas.map((l) => [l.key, l.label])}
                />
                <div className="sm:col-span-2">
                  <button type="button"
                    onClick={() => setDraftField("especie", esLavado ? "" : "LAVADO")}
                    className={`mb-2 text-xs font-semibold rounded-full px-3 py-1.5 border transition-colors ${esLavado ? "bg-yellow-400 border-yellow-400 text-slate-900" : "border-slate-300 text-slate-600 hover:border-slate-500"}`}>
                    {esLavado ? "✓ Marcado como LAVADO (limpieza de línea)" : "Marcar como LAVADO — limpieza de línea"}
                  </button>
                  {!esLavado && (
                    esEnvasadora ? (
                      <SkuPicker
                        label="Código SKU a programar"
                        value={draft.especie}
                        onChange={(v) => setDraftField("especie", v)}
                      />
                    ) : (
                      <>
                        <SelectField
                          label="Proceso / Especie"
                          value={draft.especie}
                          onChange={(v) => setDraftField("especie", v)}
                          optionsKV={(lineaCfg?.procesos || []).map(([p, e]) => [p, `${p} — ${e}`])}
                          placeholder={lineaCfg ? "Seleccionar…" : "Primero elige una línea"}
                        />
                        {draft.especie && lineaCfg && (
                          <p className="text-xs text-emerald-700 mt-1">✓ Especie: {especieFor(lineaCfg.procesos, draft.especie) || "—"}</p>
                        )}
                      </>
                    )
                  )}
                </div>
                {!esLavado && (
                  <div className="sm:col-span-2">
                    <NumField
                      label="Cantidad a producir"
                      unit={lineaCfg ? lineaCfg.unidad : undefined}
                      value={draft.cantidad}
                      onChange={(v) => setDraftField("cantidad", v)}
                    />
                  </div>
                )}
                <div className="sm:col-span-2 pt-2 border-t border-slate-200">
                  <TextAreaField
                    label="Observación de este turno (opcional)"
                    value={obsTexto}
                    onChange={setObsTexto}
                  />
                  {toastObs && <div className="mt-2"><Toast {...toastObs} /></div>}
                  <button
                    type="button"
                    onClick={handleAddObs}
                    className="w-full mt-2 flex items-center justify-center gap-2 bg-slate-700 hover:bg-slate-600 text-white font-semibold rounded-xl py-2 transition-colors text-xs"
                  >
                    <Save size={14} /> Guardar observación de {fmtFechaCorta(draft.fecha)} {draft.turno ? `· ${draft.turno}` : ""}
                  </button>
                </div>
              </div>

              {esLavado && (
                <p className="text-xs text-yellow-700 mt-2">Esta línea quedará marcada como LAVADO en ese turno (sin cantidad a producir).</p>
              )}

              <button
                onClick={handleAdd}
                className="mt-3 w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-xl py-3 transition-colors"
              >
                <Save size={18} /> Agregar al programa
              </button>
              {toast && <div className="mt-3"><Toast {...toast} /></div>}
            </Card>

            <Card title="Observaciones registradas de esta semana">
              {loadingObs ? (
                <Loader />
              ) : observacionesVentana.length === 0 ? (
                <EmptyNote text="Sin observaciones registradas para esta semana." />
              ) : (
                <div className="overflow-x-auto -mx-1">
                  <table className="w-full border-collapse text-sm min-w-[420px]">
                    <thead>
                      <tr className="bg-slate-100">
                        <th className="border border-slate-300 px-2 py-1.5 text-left text-xs font-semibold text-slate-600 uppercase tracking-wide w-[16%]">Fecha</th>
                        <th className="border border-slate-300 px-2 py-1.5 text-left text-xs font-semibold text-slate-600 uppercase tracking-wide w-[14%]">Turno</th>
                        <th className="border border-slate-300 px-2 py-1.5 text-left text-xs font-semibold text-slate-600 uppercase tracking-wide">Observación</th>
                        <th className="border border-slate-300 px-2 py-1.5 w-[70px]"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {observacionesVentana.map((o) => (
                        <tr key={o.id} className="bg-white">
                          <td className="border border-slate-200 px-2 py-1.5 align-top text-slate-700">{fmtFechaCorta(o.fecha)}</td>
                          <td className="border border-slate-200 px-2 py-1.5 align-top text-slate-700">{turnoLabel(o.turno)}</td>
                          <td className="border border-slate-200 px-2 py-1.5 align-top text-slate-800 whitespace-pre-wrap">{o.texto}</td>
                          <td className="border border-slate-200 px-2 py-1.5 align-top">
                            <div className="flex items-center gap-1">
                              <button onClick={() => handleEditObs(o)} aria-label="Editar" className="p-1.5 rounded-md hover:bg-slate-200 text-slate-600">
                                <Pencil size={14} />
                              </button>
                              <button onClick={() => handleDeleteObs(o)} aria-label="Eliminar" className="p-1.5 rounded-md hover:bg-red-100 text-red-500">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </Card>

            <Card title="Entradas de la semana" step={3}>
              {entriesVentana.length === 0 ? (
                <EmptyNote text="Todavía no hay entradas para esta semana." />
              ) : (
                <div className="space-y-2">
                  {entriesVentana.map((e) => (
                    <div key={e.id} className="flex items-start justify-between border border-slate-300 rounded-lg px-3 py-2 text-sm">
                      <div>
                        <div className="text-slate-900 font-medium">{e.lineaLabel} · {diaCorto(e.fecha)} {turnoLabel(e.turno)} ({fmtFechaCorta(e.fecha)})</div>
                        {e.especie === "LAVADO" ? (
                          <div className="text-xs text-yellow-700 font-semibold mt-0.5">LAVADO</div>
                        ) : (
                          <div className="text-xs text-slate-600 mt-0.5">
                            {procesoDisplay(e)} · {fmtNum(num(e.cantidad))} {e.unidad}
                            {e.lineaKey === "envasadora" && skuMaterial(e.especie) && (
                              <span className="text-slate-400"> · {e.especie}</span>
                            )}
                          </div>
                        )}
                      </div>
                      <button onClick={() => handleRemove(e.id)} aria-label="Eliminar" className="p-1.5 rounded-md hover:bg-slate-300 text-red-400">
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// INSUMOS Y CONSUMO
// Compara lo que el Programa de producción dice que se va a necesitar
// (cantidad programada x tasa de consumo configurada) contra el stock

// ---------------------------------------------------------------------------
// GESTIONAR ESPECIES POR LÍNEA (Jefe de producción)
// Permite agregar nuevos pares [Proceso, Especie] a una línea específica,
// sin tocar código. Estos extras se suman automáticamente a las opciones
// disponibles en: Inicio de turno (Selección/Envasado), Programa de
// producción e Insumos y Consumo.
// ---------------------------------------------------------------------------
const LINEAS_TODAS = [
  { key: "linea1", label: "Línea 1" },
  { key: "linea3", label: "Línea 3" },
  { key: "linea4", label: "Línea 4" },
  { key: "linea5", label: "Línea 5" },
  { key: "linea6", label: "Línea 6" },
  { key: "envasadora", label: "Envasadora" },
];

function EspeciesPorLineaScreen({ onBack }) {
  const { items, porLinea, agregar, eliminar, loading } = useProcesosExtra();
  const [lineaKey, setLineaKey] = useState(LINEAS_TODAS[0].key);
  const [proceso, setProceso] = useState("");
  const [especie, setEspecie] = useState("");
  const [toast, setToast] = useState(null);

  const handleAgregar = async () => {
    if (!lineaKey || !proceso.trim() || !especie.trim()) {
      setToast({ kind: "error", message: "Completa línea, proceso y especie antes de agregar." });
      return;
    }
    const ok = await agregar(lineaKey, proceso, especie);
    if (ok) {
      setToast({ kind: "ok", message: "Especie agregada a la línea." });
      setProceso("");
      setEspecie("");
    } else {
      setToast({ kind: "error", message: "No se pudo guardar. Intenta nuevamente." });
    }
  };

  const handleEliminar = async (id) => {
    const ok = await eliminar(id);
    setToast(ok ? { kind: "ok", message: "Especie eliminada." } : { kind: "error", message: "No se pudo eliminar. Intenta nuevamente." });
  };

  const extraDeLinea = items.filter((p) => p.lineaKey === lineaKey);

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Gestionar especies por línea" subtitle="Agrega nuevos procesos/especies sin tocar código" onBack={onBack} icon={Plus} accent="emerald" />

      <div className="px-4 pt-4 space-y-4">
        <Card title="Agregar especie a una línea" step={1}>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <SelectField label="Línea" value={lineaKey} onChange={setLineaKey} optionsKV={LINEAS_TODAS.map((l) => [l.key, l.label])} />
            <TextField label="Proceso (nombre exacto del SKU/proceso)" value={proceso} onChange={setProceso} />
            <TextField label="Especie" value={especie} onChange={setEspecie} />
          </div>
          {toast && <div className="mt-3"><Toast {...toast} /></div>}
          <button
            onClick={handleAgregar}
            className="mt-3 w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-xl py-2.5 transition-colors"
          >
            <Plus size={18} /> Agregar especie
          </button>
        </Card>

        <Card title={`Especies agregadas a ${LINEAS_TODAS.find((l) => l.key === lineaKey)?.label}`}>
          {loading ? (
            <Loader />
          ) : extraDeLinea.length === 0 ? (
            <EmptyNote text="Todavía no se han agregado especies extra a esta línea." />
          ) : (
            <div className="space-y-2">
              {extraDeLinea.map((p) => (
                <div key={p.id} className="flex items-center justify-between border-b border-slate-200 pb-2 last:border-0 last:pb-0">
                  <div className="text-sm">
                    <div className="font-medium text-slate-900">{p.proceso}</div>
                    <div className="text-xs text-slate-500">Especie: {p.especie}</div>
                  </div>
                  <button onClick={() => handleEliminar(p.id)} aria-label="Eliminar" className="p-1.5 rounded-md hover:bg-slate-200 text-red-500">
                    <Trash2 size={16} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// VERIFICADOR DE HORA (Envasado)
// Tabla de códigos 1-12 usados en la codificación de bolsa/caja para indicar
// el bloque horario de 2 horas en que se envasó el producto.
// Sección independiente, accesible a todos (Jefe y Supervisor).
// ---------------------------------------------------------------------------
function codigoVerificadorPara(horaStr) {
  // horaStr en formato "HH:MM". Los bloques son [desde, hasta) de 2 horas,
  // partiendo a las 08:00 y dando la vuelta a la medianoche (bloque 8: 22:00-00:00).
  if (!horaStr) return null;
  const [h, mRaw] = horaStr.split(":").map(Number);
  const minutos = h * 60 + (mRaw || 0);
  for (const v of VERIFICADOR_HORA) {
    const [dh, dm] = v.desde.split(":").map(Number);
    const [hh, hm] = v.hasta.split(":").map(Number);
    let desdeMin = dh * 60 + dm;
    let hastaMin = hh * 60 + hm;
    if (hastaMin <= desdeMin) hastaMin += 24 * 60; // cruza medianoche
    let m = minutos;
    if (m < desdeMin) m += 24 * 60;
    if (m >= desdeMin && m < hastaMin) return v.codigo;
  }
  return null;
}

function horaActualStr() {
  const d = new Date();
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

// ---------------------------------------------------------------------------
// REGISTRO CONTROL ENTREGA DE MATERIALES
// Tabla de entrega de herramientas y equipos al inicio de turno, usada en
// Selección y Envasado. Cada fila tiene: material, cantidad, quien entrega,
// quien recibe y observaciones. El registro queda guardado y se puede
// compartir por WhatsApp.
// ---------------------------------------------------------------------------
const ITEMS_ENTREGA_DEFAULT = [
  "BALANZAS",
  "PORUÑAS",
  "CASCOS",
  "CORTADOR",
  "CHAQUETAS FRIO",
  "RADIOS",
  "TRANSPALETAS",
  "RASTRILLOS",
  "TRANSPALETAS ELÉCTRICAS",
  "HOMBRE A BORDO",
  "LLAVES",
];

// Secuencia de turnos para buscar el registro anterior de recepción
const TURNO_ANTERIOR = { T1: "T3", T2: "T1", T3: "T2" };

function fechaAtras(fecha, dias) {
  const d = new Date(fecha + "T12:00:00");
  d.setDate(d.getDate() - dias);
  return d.toISOString().slice(0, 10);
}

function buildResumenEntrega(record, areaLabel, inicialMap) {
  const tE = record.turnoEntrega || "—";
  const tR = record.turnoRecibe  || "—";
  const lines = [
    `*Entrega de Materiales — ${areaLabel}*`,
    `Fecha: ${fmtFecha(record.fecha)}`,
    `${tE} → ${tR}`,
    `Entrega: ${record.entrega || "—"} · Recibe: ${record.recibe || "—"}`,
    "",
    "*Artículos (Inicial → Entrega):*",
  ];
  // Una línea por artículo con solo cantidad inicial y cantidad de entrega
  // (sin bloque de código — algunos clientes de WhatsApp no muestran bien el
  // texto dentro de ```, y hacía desaparecer tabla y observaciones).
  let hayDiferencias = false;
  (record.items || []).forEach((it) => {
    if (it.cant === "" && !it.obs) return; // sin datos para este artículo — se omite
    const ini = inicialMap ? (inicialMap[it.nombre] ?? "") : "";
    const cambio = ini !== "" && it.cant !== "" && String(ini) !== String(it.cant);
    if (cambio) hayDiferencias = true;
    const iniTxt = ini !== "" ? fmtNum(num(ini)) : "—";
    const entTxt = it.cant !== "" ? fmtNum(num(it.cant)) : "—";
    const flag = cambio ? " ⚠" : "";
    const obsTxt = it.obs ? ` · Obs: ${it.obs}` : "";
    lines.push(`• ${it.nombre}: ${iniTxt} → ${entTxt}${flag}${obsTxt}`);
  });
  if (hayDiferencias) {
    lines.push("");
    lines.push("⚠ Hubo cambios: no se entregaron todos los materiales en la misma cantidad del inicio del turno.");
  }
  return lines.join("\n");
}

function EntregaMaterialesCard({ areaKey, fecha, turno, autor, supervisoresList }) {
  if (areaKey !== "seleccion" && areaKey !== "envasado") return null;

  const areaLabel = areaKey === "seleccion" ? "Selección" : "Envasado";
  const [records, saveRecords, loading] = useSharedList(`entrega-materiales-${areaKey}`);

  // Turno que entrega = el turno que se está cerrando ahora mismo. Turno que
  // recibe se calcula automáticamente (el que sigue cronológicamente).
  const turnoEntrega = turno;
  const turnoRecibeCalc = turnoSiguiente(fecha, turno);
  const turnoRecibe = turnoRecibeCalc?.turno || "";

  const [entrega, setEntrega] = useState(autor || "");
  const [recibe,  setRecibe]  = useState("");
  const [items,   setItems]   = useState(
    ITEMS_ENTREGA_DEFAULT.map((nombre) => ({ nombre, cant: "", obs: "" }))
  );
  const [toast,     setToast]     = useState(null);
  const [lastSaved, setLastSaved] = useState(null);

  useEffect(() => { if (autor && !entrega) setEntrega(autor); }, [autor]);

  // Busca el registro previo donde el turno anterior entregó AL turno actual
  const inicialRecord = useMemo(() => {
    if (!turnoEntrega) return null;
    // Cadena: T3(día X) → T1(X) → T2(X) → T3(X+1). Quien entregó al turno que
    // cierra ahora es el turno anterior en esa cadena, con su fecha exacta
    // (para T3 es el T2 del día anterior).
    const ant = turnoAnterior(fecha, turnoEntrega);
    if (!ant) return null;
    const candidatos = records.filter(
      (r) => r.turnoEntrega === ant.turno &&
             r.turnoRecibe  === turnoEntrega &&
             r.fecha === ant.fecha
    );
    return [...candidatos].sort((a, b) => b.id - a.id)[0] || null;
  }, [records, turnoEntrega, fecha]);

  const inicialMap = useMemo(() => {
    if (!inicialRecord) return {};
    const m = {};
    (inicialRecord.items || []).forEach((it) => { m[it.nombre] = it.cant; });
    return m;
  }, [inicialRecord]);

  const setItemField = (idx, campo, valor) => {
    if (campo === "cant") {
      const v = valor.replace(/[^0-9]/g, "");
      setItems((prev) => prev.map((it, i) => i === idx ? { ...it, cant: v } : it));
    } else {
      setItems((prev) => prev.map((it, i) => i === idx ? { ...it, [campo]: valor } : it));
    }
  };

  const diferencias = items.filter((it) => {
    const ini = inicialMap[it.nombre];
    return ini !== undefined && ini !== "" && it.cant !== "" && String(ini) !== String(it.cant);
  });

  const guardar = async () => {
    if (!fecha || !turnoEntrega) {
      setToast({ kind: "error", message: "Falta fecha o turno del cierre." });
      return;
    }
    const record = { id: Date.now(), fecha, turnoEntrega, turnoRecibe, entrega, recibe, items };
    const ok = await saveRecords([
      record,
      ...records.filter((r) => !(r.fecha === fecha && r.turnoEntrega === turnoEntrega && r.turnoRecibe === turnoRecibe)),
    ]);
    if (ok) { setToast({ kind: "ok", message: "Entrega de materiales guardada." }); setLastSaved(record); }
    else    { setToast({ kind: "error", message: "No se pudo guardar. Intenta nuevamente." }); }
  };

  return (
    <Card title="Entrega de Materiales">
      <p className="text-xs text-slate-500 mb-3">
        {turnoEntrega} ({fmtFecha(fecha)}) → {turnoRecibe || "—"} ({turnoRecibeCalc ? fmtFecha(turnoRecibeCalc.fecha) : "—"})
      </p>

      {inicialRecord ? (
        <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-1.5 mb-3">
          ✓ Cantidades iniciales cargadas desde entrega {inicialRecord.turnoEntrega} → {inicialRecord.turnoRecibe} del {fmtFecha(inicialRecord.fecha)}
        </p>
      ) : (
        <p className="text-xs text-slate-400 bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 mb-3">
          Sin registro previo de recepción — la columna "Inicial" aparecerá vacía.
        </p>
      )}

      <div className="grid grid-cols-2 gap-6 mb-4">
        <div>
          <label className="block text-xs font-medium text-slate-600 mb-1">Quien entrega ({turnoEntrega || "—"})</label>
          <SupervisorSelect
            area={areaKey === "seleccion" ? "Seleccion" : "Envasado"}
            value={entrega}
            onChange={setEntrega}
            supervisoresList={supervisoresList}
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-slate-600 mb-1">Quien recibe ({turnoRecibe || "—"})</label>
          <SupervisorSelect
            area={areaKey === "seleccion" ? "Seleccion" : "Envasado"}
            value={recibe}
            onChange={setRecibe}
            supervisoresList={supervisoresList}
          />
        </div>
      </div>

      {diferencias.length > 0 && (
        <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3 flex items-start gap-3 mb-4">
          <AlertTriangle size={18} className="text-amber-600 shrink-0 mt-0.5" />
          <div>
            <div className="text-sm font-semibold text-amber-800">
              {diferencias.length} artículo{diferencias.length > 1 ? "s" : ""} con diferencia respecto al inicio del turno
            </div>
            <div className="text-xs text-amber-700 mt-0.5">
              {diferencias.map((it) => `${it.nombre}: recibido ${inicialMap[it.nombre]} → entrega ${it.cant}`).join(" · ")}
            </div>
          </div>
        </div>
      )}

      <div className="overflow-x-auto -mx-1 mb-4">
        <table className="w-full border-collapse text-sm min-w-[480px]">
          <thead>
            <tr className="bg-slate-100">
              <th className="border border-slate-300 px-3 py-2 text-left   text-xs font-semibold text-slate-600 uppercase tracking-wide w-[38%]">Artículo</th>
              <th className="border border-slate-300 px-2 py-2 text-center text-xs font-semibold text-slate-500 uppercase tracking-wide w-[16%]">Cant. Inicial</th>
              <th className="border border-slate-300 px-2 py-2 text-center text-xs font-semibold text-slate-600 uppercase tracking-wide w-[16%]">Cant. Entrega</th>
              <th className="border border-slate-300 px-3 py-2 text-left   text-xs font-semibold text-slate-600 uppercase tracking-wide">Observación</th>
            </tr>
          </thead>
          <tbody>
            {items.map((it, idx) => {
              const ini = inicialMap[it.nombre] ?? "";
              const hayDiff = ini !== "" && it.cant !== "" && String(ini) !== String(it.cant);
              return (
                <tr key={it.nombre} className={hayDiff ? "bg-amber-50" : idx % 2 === 0 ? "bg-white" : "bg-slate-50"}>
                  <td className="border border-slate-200 px-3 py-2 font-semibold text-slate-800 align-middle">
                    <div className="flex items-center gap-1.5">
                      {hayDiff && <AlertTriangle size={13} className="text-amber-500 shrink-0" />}
                      {it.nombre}
                    </div>
                  </td>
                  <td className="border border-slate-200 px-2 py-2 text-center align-middle text-slate-500 font-medium bg-slate-50/50">
                    {ini !== "" ? fmtNum(num(ini)) : <span className="text-slate-300">—</span>}
                  </td>
                  <td className={`border p-0 align-middle ${hayDiff ? "border-amber-300" : "border-slate-200"}`}>
                    <input
                      type="text"
                      inputMode="numeric"
                      pattern="[0-9]*"
                      className={`w-full h-full px-2 py-2 text-sm text-center bg-transparent focus:outline-none focus:bg-blue-50 focus:ring-1 focus:ring-blue-400 rounded ${hayDiff ? "text-amber-700 font-bold" : "text-slate-800"}`}
                      placeholder="0"
                      value={it.cant}
                      onChange={(e) => setItemField(idx, "cant", e.target.value)}
                    />
                  </td>
                  <td className="border border-slate-200 p-0 align-middle">
                    <input
                      type="text"
                      className="w-full h-full px-3 py-2 text-sm bg-transparent focus:outline-none focus:bg-blue-50 focus:ring-1 focus:ring-blue-400 rounded"
                      placeholder="—"
                      value={it.obs}
                      onChange={(e) => setItemField(idx, "obs", e.target.value)}
                    />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {toast && <div className="mb-3"><Toast {...toast} /></div>}

      <button onClick={guardar}
        className="w-full flex items-center justify-center gap-2 bg-slate-800 hover:bg-slate-700 text-white font-semibold rounded-xl py-3 transition-colors">
        <Save size={18} /> Guardar Entrega de Materiales
      </button>

      {lastSaved && (
        <a href={whatsappShareUrl(buildResumenEntrega(lastSaved, areaLabel, inicialMap))} target="_blank" rel="noopener noreferrer"
          className="w-full mt-2 flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-xl py-3 transition-colors">
          <Share2 size={18} /> Compartir por WhatsApp
        </a>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Pestaña: Entrega de Materiales (Selección y Envasado) — separada del
// Cierre de turno. La persona elige la fecha y el turno que entrega, igual
// que en Cierre de turno, y aparece la misma tarjeta de siempre.
// ---------------------------------------------------------------------------
function EntregaMaterialesTab({ area, autor, supervisoresList }) {
  const [fecha, setFecha] = useState(today());
  const [turno, setTurno] = useState("");

  return (
    <div className="space-y-4 pb-10">
      <Card title="Selecciona el turno que entrega" step={1}>
        <div className="mb-3">
          <label className="block text-xs font-medium text-slate-600 mb-1">Fecha</label>
          <input type="date" style={{ colorScheme: "light" }} className={inputBase} value={fecha} onChange={(e) => setFecha(e.target.value)} />
        </div>
        <label className="block text-xs font-medium text-slate-600 mb-1">Turno</label>
        <TurnoSelector value={turno} onChange={setTurno} />
      </Card>

      {!turno ? (
        <Card><EmptyNote text="Selecciona la fecha y el turno que entrega para registrar la entrega de materiales." /></Card>
      ) : (
        <EntregaMaterialesCard areaKey={area.key} fecha={fecha} turno={turno} autor={autor} supervisoresList={supervisoresList} />
      )}
    </div>
  );
}

function VerificadorHoraScreen({ onBack }) {
  const [horaConsulta, setHoraConsulta] = useState(horaActualStr());
  const codigoActual = codigoVerificadorPara(horaConsulta);

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Verificador de hora" subtitle="Código de bloque horario para codificación de envasado" onBack={onBack} icon={Clock} accent="amber" />

      <div className="px-4 pt-4 space-y-4">
        <Card title="Consultar código por hora" step={1}>
          <TextField label="Hora" type="time" value={horaConsulta} onChange={setHoraConsulta} />
          <div className="mt-3 rounded-lg border border-amber-300 bg-amber-50 px-4 py-3 text-center">
            <p className="text-xs text-slate-500">Código verificador de hora</p>
            <p className="text-3xl font-bold text-amber-700">{codigoActual ?? "—"}</p>
          </div>
          <button
            onClick={() => setHoraConsulta(horaActualStr())}
            className="mt-3 w-full flex items-center justify-center gap-2 bg-slate-200 hover:bg-slate-300 text-slate-800 font-semibold rounded-xl py-2 transition-colors text-sm"
          >
            <Clock size={16} /> Usar hora actual
          </button>
        </Card>

        <Card title="Tabla completa de códigos">
          <div className="overflow-x-auto -mx-2">
            <table className="w-full text-sm border-collapse">
              <thead>
                <tr className="text-left text-xs text-slate-500 uppercase tracking-wide">
                  <th className="px-2 py-2">Código</th>
                  <th className="px-2 py-2">Desde</th>
                  <th className="px-2 py-2">Hasta</th>
                </tr>
              </thead>
              <tbody>
                {VERIFICADOR_HORA.map((v) => (
                  <tr key={v.codigo} className={`border-t border-slate-200 ${v.codigo === codigoActual ? "bg-amber-100 font-semibold" : ""}`}>
                    <td className="px-2 py-2">{v.codigo}</td>
                    <td className="px-2 py-2">{v.desde}</td>
                    <td className="px-2 py-2">{v.hasta}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Resúmenes COMPLETOS por área (para WhatsApp) — a diferencia de `resumenInicio`
// (que es la versión compacta mostrada en pantalla), estos incluyen TODO lo
// capturado en el formulario: desglose por tipo de bandeja, por línea, por
// material, etc.
// ---------------------------------------------------------------------------
// Cada resumen*Completo devuelve { dotacion, produccion, materiales }: tres
// grupos de [label, valor] que se imprimen en ese orden fijo en el mensaje de
// WhatsApp — dotación global (no desglosada por línea), luego los datos de
// producción (kg, rendimiento, cumplimiento), y por último los materiales de
// piso de planta.
function resumenLavadoInicioCompleto(r) {
  const dotacion = [
    ["Operarios", fmtNum(r.operarios || 0)],
    ["Movilizadores", fmtNum(r.movilizadores || 0)],
    ["Jefe de línea", fmtNum(r.jefeLinea || 0)],
    ["¿Dotación completa?", r.dotacionCompleta || "—"],
    ["Comentarios dotación", r.comentariosDotacion || "—"],
  ];
  const produccion = [];
  TIPOS_BANDEJA.forEach((t) => {
    const v = num(r[`sucios_${t.key}_pallets`]);
    if (v > 0) produccion.push([`Sucios · ${t.label}`, `${fmtNum(v)} pallets`]);
  });
  produccion.push(["Total pallets sucios al iniciar", fmtNum(totalPalletsManual(r, "sucios"))]);
  const materiales = [];
  MATERIALES_LAVADO.forEach((nombre, i) => {
    const v = num(r[`inicioMat_${i}`]);
    if (v > 0) materiales.push([`Material · ${nombre}`, fmtNum(v)]);
  });
  return { dotacion, produccion, materiales };
}

function resumenLavadoCierreCompleto(r, inicio) {
  const src = inicio || r;
  const dotacion = [
    ["Operarios", fmtNum(src.operarios || 0)],
    ["Movilizadores", fmtNum(src.movilizadores || 0)],
    ["Jefe de línea", fmtNum(src.jefeLinea || 0)],
    ["¿Dotación completa?", src.dotacionCompleta || "—"],
    ["Comentarios dotación", src.comentariosDotacion || "—"],
  ];
  const produccion = [];
  produccion.push(["=", "Lavados"]);
  TIPOS_BANDEJA.forEach((t) => {
    const v = palletsValue(r, "lavados", t);
    if (v > 0) produccion.push([t.label, `${fmtNum(v)} pallets`, "plano"]);
  });
  produccion.push(["Total pallets lavados", fmtNum(totalPalletsFromPrefix(r, "lavados"))]);
  produccion.push(["=", "Pendientes"]);
  TIPOS_BANDEJA.forEach((t) => {
    const v = num(r[`pendientes_${t.key}_pallets`]);
    if (v > 0) produccion.push([t.label, `${fmtNum(v)} pallets`, "plano"]);
  });
  produccion.push(["Total pallets pendientes", fmtNum(totalPalletsManual(r, "pendientes"))]);
  const materiales = [
    ["Rollos de film", fmtNum(r.rollosFilm || 0)],
    ["Bolsas de bins", fmtNum(r.bolsasBins || 0)],
  ];
  MATERIALES_LAVADO.forEach((nombre, i) => {
    const v = num(r[`finMat_${i}`]);
    if (v > 0) materiales.push([`Material · ${nombre}`, fmtNum(v)]);
  });
  return { dotacion, produccion, materiales };
}

// Total por puesto (Operarios, Movilizadores, Jefe de Línea) sumando TODAS
// las líneas activas (+ armado en Selección; + Envasadora/Línea 5 en
// Envasado) — a diferencia del desglose por línea, esto da un solo número
// por cargo para todo el turno.
function totalesPorPuestoSeleccion(r) {
  return DOTACION_LINEA.map((label, i) => {
    let t = 0;
    LINEAS_SELECCION.forEach((l) => {
      if (r[`linea_${l.key}_activa`] === "Sí") t += num(r[`linea_${l.key}_dot${i}`]);
    });
    if (r.armado_activa === "Sí") t += num(r[`armado_dot${i}`]);
    return [`Total ${label}`, fmtNum(t)];
  });
}

function totalesPorPuestoEnvasado(r) {
  return DOTACION_LINEA.map((label, i) => {
    let t = 0;
    if (r.activa_envasadora === "Sí") t += num(r[`envasadora_dot${i}`]);
    if (r.activa_linea5 === "Sí") t += num(r[`linea5_dot${i}`]);
    return [`Total ${label}`, fmtNum(t)];
  });
}

function resumenSeleccionInicioCompleto(r) {
  const dotacion = [];
  DOTACION_GENERAL_SELECCION.forEach((label, i) => {
    const v = num(r[`dg_${i}`]);
    if (v > 0) dotacion.push([label, fmtNum(v)]);
  });
  dotacion.push(...totalesPorPuestoSeleccion(r));
  dotacion.push(["Total dotación", fmtNum(totalDotacionSeleccion(r))]);
  dotacion.push(["¿Dotación completa?", r.dotacionCompleta || "—"]);
  dotacion.push(["Comentarios dotación", r.comentariosDotacion || "—"]);

  const produccion = [];
  const activas = LINEAS_SELECCION.filter((l) => r[`linea_${l.key}_activa`] === "Sí");
  if (activas.length === 0) {
    produccion.push(["Líneas activas", "Ninguna"]);
  } else {
    activas.forEach((l) => {
      const procesos = ESPECIE_SLOTS.map((s) => r[`linea_${l.key}_proceso${sufijoEspecie(s)}`]).filter(Boolean);
      produccion.push([`${l.label} · Proceso`, procesos.length ? procesos.join(" + ") : "—"]);
    });
  }
  produccion.push(["Armado de materiales", r.armado_activa === "Sí" ? "Activo" : "Inactivo"]);

  const materiales = [];
  MATERIALES_SELECCION.forEach((nombre, i) => {
    const v = num(r[`inicioMat_${i}`]);
    if (v > 0) materiales.push([`Material · ${nombre}`, fmtNum(v)]);
  });

  return { dotacion, produccion, materiales };
}

function resumenSeleccionCierreCompleto(r, inicio, programaEntries) {
  const src = inicio || r;
  const dotacion = [];
  DOTACION_GENERAL_SELECCION.forEach((label, i) => {
    const v = num(src[`dg_${i}`]);
    if (v > 0) dotacion.push([label, fmtNum(v)]);
  });
  dotacion.push(...totalesPorPuestoSeleccion(src));
  dotacion.push(["Total dotación", fmtNum(totalDotacionSeleccion(src))]);
  dotacion.push(["¿Dotación completa?", src.dotacionCompleta || "—"]);
  dotacion.push(["Comentarios dotación", src.comentariosDotacion || "—"]);

  const produccion = [];
  const activas = LINEAS_SELECCION.filter((l) => (inicio?.[`linea_${l.key}_activa`] ?? r[`linea_${l.key}_activa`]) === "Sí");
  if (activas.length === 0) {
    produccion.push(["Líneas activas", "Ninguna"]);
  } else {
    const rendTotal = computeSeleccionRendimiento({ ...(inicio || {}), ...r });
    const { cumplimiento, kgProgramado, kgIngresado } = computeSeleccionCumplimiento(r, inicio, programaEntries);
    produccion.push(["Rendimiento total", fmtPct(rendTotal)]);
    produccion.push(["Cumplimiento", kgProgramado ? `${fmtPct(cumplimiento)} (${fmtNum(kgIngresado, 0)}/${fmtNum(kgProgramado, 0)} Kg)` : "Sin Kg programado"]);
    // Primero el resultado general del turno (arriba) y después el detalle de
    // cada línea: Rendimiento y Cumplimiento de la línea, y luego sus procesos.
    activas.forEach((l) => {
      const especies = especiesActivasLinea(l, inicio, r);
      const kgProgLinea = (programaEntries || [])
        .filter((p) => (p.lineaKey === l.key || p.linea === l.key) && p.especie !== "LAVADO")
        .reduce((s, p) => s + num(p.cantidad), 0);
      const tot = lineaKgTotales(r, l);
      produccion.push(["#", l.label]);
      produccion.push(["Rendimiento", tot.ing ? fmtPct(tot.apr / tot.ing) : "—"]);
      produccion.push(["Cumplimiento", kgProgLinea ? `${fmtPct(tot.ing / kgProgLinea)} (${fmtNum(tot.ing, 0)}/${fmtNum(kgProgLinea, 0)} Kg)` : "Sin Kg programado"]);
      especies.forEach(({ s, proceso }) => {
        const ing = num(r[`kg_${l.key}_e${s}_ing`]);
        const apr = kgAprobadoTotal(r, l.key, s);
        produccion.push(["##", proceso]);
        produccion.push(["Kg ingresados", fmtNum(ing)]);
        produccion.push(["Kg aprobados", fmtNum(apr)]);
        produccion.push(["Rendimiento", fmtPct(ing ? apr / ing : 0)]);
      });
    });
  }

  const materiales = [];
  MATERIALES_SELECCION.forEach((nombre, i) => {
    const v = num(r[`finMat_${i}`]);
    if (v > 0) materiales.push([`Material · ${nombre}`, fmtNum(v)]);
  });

  return { dotacion, produccion, materiales };
}

function resumenEnvasadoInicioCompleto(r) {
  const dotacion = [];
  DOTACION_GENERAL_ENVASADO.forEach((label, i) => {
    const v = num(r[`dg_${i}`]);
    if (v > 0) dotacion.push([label, v]);
  });
  dotacion.push(...totalesPorPuestoEnvasado(r));
  dotacion.push(["Total dotación", fmtNum(dotacionEnvasadoTotal(r))]);
  dotacion.push(["¿Dotación completa?", r.dotacionCompleta || "—"]);
  dotacion.push(["Comentarios dotación", r.comentariosDotacion || "—"]);

  const produccion = [];
  if (r.activa_envasadora === "Sí") {
    const mat = r.envasadora_sku ? skuMaterial(r.envasadora_sku) : null;
    produccion.push(["Envasadora · SKU", r.envasadora_sku ? `${mat?.producto || "—"} (${r.envasadora_sku})` : "—"]);
  }
  if (r.activa_linea5 === "Sí") {
    produccion.push(["Línea 5 · Especie", r.linea5_especie || "—"]);
  }
  if (r.activa_envasadora !== "Sí" && r.activa_linea5 !== "Sí") produccion.push(["Líneas de trabajo", "Ninguna activa"]);

  const materiales = [];
  MATERIALES_ENVASADO.forEach((nombre, i) => {
    if (i === 3 || i === 4) return; // Cantidad MTC Cajas/Bolsas: se desglosan por SKU abajo.
    const v = num(r[`inicioMat_${i}`]);
    if (v > 0) materiales.push([`Material · ${nombre}`, v]);
  });
  materiales.push(...materialesMtcSkuLines(r.mtcCajasSkuInicio, "Cajas"));
  materiales.push(...materialesMtcSkuLines(r.mtcBolsasSkuInicio, "Bolsas"));

  return { dotacion, produccion, materiales };
}

// Convierte el desglose de MTC por SKU (Cajas o Bolsas, por separado — son
// MTC distintos e independientes) en líneas de mensaje, una por cada
// SKU/MTC que tenga cantidad registrada.
function materialesMtcSkuLines(rows, unitLabel) {
  return (Array.isArray(rows) ? rows : [])
    .filter((row) => num(row.cantidad) > 0)
    .map((row) => [`Material · MTC ${unitLabel} ${row.sku || "—"}`, fmtNum(row.cantidad || 0)]);
}

function resumenEnvasadoCierreCompleto(r, inicio, programaEntries, horarios) {
  const src = inicio || r;
  const dotacion = [];
  DOTACION_GENERAL_ENVASADO.forEach((label, i) => {
    const v = num(src[`dg_${i}`]);
    if (v > 0) dotacion.push([label, v]);
  });
  dotacion.push(...totalesPorPuestoEnvasado(src));
  dotacion.push(["Total dotación", fmtNum(dotacionEnvasadoTotal(src))]);
  dotacion.push(["¿Dotación completa?", src.dotacionCompleta || "—"]);
  dotacion.push(["Comentarios dotación", src.comentariosDotacion || "—"]);

  const produccion = [];
  const activaLinea5 = (inicio?.activa_linea5 ?? r.activa_linea5) === "Sí";
  const activaEnvasadora = (inicio?.activa_envasadora ?? r.activa_envasadora) === "Sí";
  const mEnv = computeEnvasadoMetrics(r);
  const kgProgLinea5 = (programaEntries || [])
    .filter((p) => (p.lineaKey === "linea5" || p.linea === "linea5") && p.especie !== "LAVADO")
    .reduce((s, p) => s + num(p.cantidad), 0);
  // Resultado general del turno (arriba) y después el detalle de cada línea.
  if (activaEnvasadora) produccion.push(["Cumplimiento Envasadora", fmtPct(mEnv.cumplimiento)]);
  if (activaLinea5) {
    produccion.push(["Rendimiento Línea 5", fmtPct(mEnv.rendimientoLinea5)]);
    produccion.push(["Cumplimiento Línea 5", kgProgLinea5 ? `${fmtPct(mEnv.l5Ing / kgProgLinea5)} (${fmtNum(mEnv.l5Ing, 0)}/${fmtNum(kgProgLinea5, 0)} Kg)` : "Sin Kg programado"]);
  }
  if (activaLinea5) {
    const m = mEnv;
    produccion.push(["#", "Línea 5"]);
    produccion.push(["Rendimiento", fmtPct(m.rendimientoLinea5)]);
    produccion.push(["Cumplimiento", kgProgLinea5 ? `${fmtPct(m.l5Ing / kgProgLinea5)} (${fmtNum(m.l5Ing, 0)}/${fmtNum(kgProgLinea5, 0)} Kg)` : "Sin Kg programado"]);
    produccion.push(["Especie", inicio?.linea5_especie || r.linea5_especie || "—"]);
    produccion.push(["Proceso", r.l5_proceso || "—"]);
    produccion.push(["Kg ingresados", fmtNum(m.l5Ing)]);
    produccion.push(["Kg aprobados", fmtNum(m.l5Apr)]);
  }
  if (activaEnvasadora) {
    const m = mEnv;
    const mat = r.envasadora_sku ? skuMaterial(r.envasadora_sku) : null;
    produccion.push(["#", "Envasadora"]);
    produccion.push(["Cumplimiento", fmtPct(m.cumplimiento)]);
    produccion.push(["SKU", r.envasadora_sku ? `${mat?.producto || "—"} (${r.envasadora_sku})` : "—"]);
    produccion.push(["Cajas producidas", fmtNum(r.cajasProducidas || 0)]);
    produccion.push(["Cajas programadas", fmtNum(r.cajasProgramadas || 0)]);
    produccion.push(["Cajas teóricas consumo", fmtNum(r.cajasTeoricas || 0)]);
    produccion.push(["Cajas consumidas real", fmtNum(r.cajasConsumidasReal || 0)]);
    produccion.push(["Bolsas teóricas consumo", fmtNum(r.bolsasTeoricas || 0)]);
    produccion.push(["Bolsas consumidas real", fmtNum(r.bolsasConsumidasReal || 0)]);
    produccion.push(["Merma cajas", fmtPct(m.mermaCajas)]);
    produccion.push(["Merma bolsas", fmtPct(m.mermaBolsas)]);

    // Detenciones + Tiempo Efectivo + Productividad (bolsas por minuto).
    const te = computeTiempoEfectivoEnvasado(r, horarios);
    const detenciones = (r.detenciones || []).filter((d) => (d.motivo && d.motivo.trim()) || num(d.minutos) > 0);
    if (detenciones.length > 0) {
      const detalle = detenciones.map((d) => `  · ${fmtNum(d.minutos || 0)} min ${d.motivo || "—"}`).join("\n");
      produccion.push(["Detenciones", `\n${detalle}`]);
    }
    if (te.horario) {
      produccion.push(["T.E", `${fmtNum(te.tiempoEfectivo)} min`]);
      produccion.push(["Productividad", `${fmtNum(te.productividad, 2)} bpm`]);
    }
  }

  const materiales = [];
  MATERIALES_ENVASADO.forEach((nombre, i) => {
    if (i === 3 || i === 4) return; // Cantidad MTC Cajas/Bolsas: se desglosan por SKU abajo.
    const v = num(r[`finMat_${i}`]);
    if (v > 0) materiales.push([`Material · ${nombre}`, v]);
  });
  materiales.push(...materialesMtcSkuLines(r.mtcCajasSkuFin, "Cajas"));
  materiales.push(...materialesMtcSkuLines(r.mtcBolsasSkuFin, "Bolsas"));

  return { dotacion, produccion, materiales };
}

// ---------------------------------------------------------------------------
// Configuración de las 3 áreas
// ---------------------------------------------------------------------------
const AREAS = {
  lavado: {
    key: "lavado",
    title: "Lavado de bandejas",
    icon: Droplets,
    accent: "blue",
    InicioComponent: LavadoInicio,
    CierreComponent: LavadoCierre,
    resumenInicio: (r) => [
      ["Hora inicio", r.horaInicio || "—"],
      ["Operarios", fmtNum(r.operarios || 0)],
      ["Movilizadores", fmtNum(r.movilizadores || 0)],
      ["Jefe de línea", fmtNum(r.jefeLinea || 0)],
      ["¿Dotación completa?", r.dotacionCompleta || "—"],
      ["Pallets sucios al iniciar", totalPalletsManual(r, "sucios")],
      ["Comentarios dotación", r.comentariosDotacion || "—"],
    ],
    resumenCompletoInicio: resumenLavadoInicioCompleto,
    resumenCompletoCierre: (r, inicio) => resumenLavadoCierreCompleto(r, inicio),
    indicator: { label: "Total pallets lavados", format: "number", compute: (r) => totalPalletsFromPrefix(r, "lavados") },
    detailCharts: [
      {
        title: "Pallets lavados por tipo de bandeja (día x turno)",
        format: "number",
        series: TIPOS_BANDEJA.map((t, i) => ({
          key: `lav_${t.key}`,
          label: t.label.replace(/\s*\(.*\)/, ""),
          color: SERIES_COLORS[i % SERIES_COLORS.length],
          compute: (r) => palletsValue(r, "lavados", t),
        })),
      },
      {
        title: "Pallets pendientes por lavar al cierre (día x turno)",
        format: "number",
        series: TIPOS_BANDEJA.map((t, i) => ({
          key: `pend_${t.key}`,
          label: t.label.replace(/\s*\(.*\)/, ""),
          color: SERIES_COLORS[i % SERIES_COLORS.length],
          compute: (r) => num(r[`pendientes_${t.key}_pallets`]),
        })),
      },
    ],
  },
  seleccion: {
    key: "seleccion",
    title: "Selección",
    icon: Filter,
    accent: "purple",
    InicioComponent: SeleccionInicio,
    CierreComponent: SeleccionCierre,
    resumenInicio: (r) => {
      const activas = LINEAS_SELECCION.filter((l) => r[`linea_${l.key}_activa`] === "Sí");
      return [
        ["Hora inicio", r.horaInicio || "—"],
        ["Total dotación", fmtNum(totalDotacionSeleccion(r))],
        ["¿Dotación completa?", r.dotacionCompleta || "—"],
        ["Líneas activas", activas.length ? activas.map((l) => l.label).join(", ") : "Ninguna"],
        ...activas.map((l) => {
          const procesos = ESPECIE_SLOTS.map((s) => r[`linea_${l.key}_proceso${sufijoEspecie(s)}`]).filter(Boolean);
          return [l.label, `${procesos.length ? procesos.join(" + ") : "—"} · Dot. ${fmtNum(lineaDotacionTotal(r, l))}`];
        }),
        ["Armado de materiales", r.armado_activa === "Sí" ? `Activo · Dot. ${fmtNum(armadoDotacionTotal(r))}` : "Inactivo"],
      ];
    },
    resumenCompletoInicio: resumenSeleccionInicioCompleto,
    resumenCompletoCierre: resumenSeleccionCierreCompleto,
    indicator: { label: "Rendimiento de línea", format: "pct", compute: (r) => computeSeleccionRendimiento(r) },
    detailCharts: [
      {
        title: "Kg ingresados por línea (día x turno)",
        format: "number",
        series: LINEAS_SELECCION.map((l, i) => ({
          key: `ing_${l.key}`,
          label: l.label,
          color: SERIES_COLORS[i % SERIES_COLORS.length],
          compute: (r) => lineaKgTotales(r, l).ing,
        })),
      },
      {
        title: "Rendimiento por línea (día x turno)",
        format: "pct",
        series: LINEAS_SELECCION.map((l, i) => ({
          key: `rend_${l.key}`,
          label: l.label,
          color: SERIES_COLORS[i % SERIES_COLORS.length],
          compute: (r) => {
            const t = lineaKgTotales(r, l);
            return t.ing ? t.apr / t.ing : 0;
          },
        })),
      },
    ],
  },
  envasado: {
    key: "envasado",
    title: "Envasado",
    icon: Package,
    accent: "amber",
    InicioComponent: EnvasadoInicio,
    CierreComponent: EnvasadoCierre,
    resumenInicio: (r) => [
      ["Hora inicio", r.horaInicio || "—"],
      ...(r.activa_envasadora === "Sí" ? [["Envasadora · Especie/SKU", `${r.envasadora_sku ? `${skuProducto(r.envasadora_sku)} (${r.envasadora_sku})` : "—"} · Dot. ${fmtNum(DOTACION_LINEA.reduce((s, _, i) => s + num(r[`envasadora_dot${i}`]), 0))}`]] : []),
      ...(r.activa_linea5 === "Sí" ? [["Línea 5 · Especie", `${r.linea5_especie || "—"} · Dot. ${fmtNum(DOTACION_LINEA.reduce((s, _, i) => s + num(r[`linea5_dot${i}`]), 0))}`]] : []),
      ...(r.activa_envasadora !== "Sí" && r.activa_linea5 !== "Sí" ? [["Líneas de trabajo", "Ninguna activa"]] : []),
      ["Total dotación", fmtNum(dotacionEnvasadoTotal(r))],
      ["¿Dotación completa?", r.dotacionCompleta || "—"],
      ...materialesMtcSkuLines(r.mtcCajasSkuInicio, "Cajas").map(([l, v]) => [l.replace("Material · ", ""), v]),
      ...materialesMtcSkuLines(r.mtcBolsasSkuInicio, "Bolsas").map(([l, v]) => [l.replace("Material · ", ""), v]),
    ],
    resumenCompletoInicio: resumenEnvasadoInicioCompleto,
    resumenCompletoCierre: resumenEnvasadoCierreCompleto,
    indicator: {
      label: "Cumplimiento / rendimiento del turno",
      format: "pct",
      compute: (r) => {
        const m = computeEnvasadoMetrics(r);
        if (r.activa_envasadora === "Sí") return m.cumplimiento;
        if (r.activa_linea5 === "Sí") return m.rendimientoLinea5;
        return 0;
      },
    },
    detailCharts: [
      {
        title: "Cumplimiento y mermas — Envasadora (día x turno)",
        format: "pct",
        series: [
          { key: "cumplimiento", label: "Cumplimiento", color: SERIES_COLORS[0], compute: (r) => computeEnvasadoMetrics(r).cumplimiento },
          { key: "mermaCajas", label: "Merma cajas", color: SERIES_COLORS[3], compute: (r) => computeEnvasadoMetrics(r).mermaCajas },
          { key: "mermaBolsas", label: "Merma bolsas", color: SERIES_COLORS[4], compute: (r) => computeEnvasadoMetrics(r).mermaBolsas },
        ],
      },
    ],
  },
};

// ---------------------------------------------------------------------------
// Pestaña: Inicio de turno
// ---------------------------------------------------------------------------
function InicioTab({ area, values, setField, editingId, onSaved, onCancelEdit, autor }) {
  const [records, saveRecords, loading, storageError] = useSharedList(`${area.key}-inicio-records`);
  const [programas] = useSharedList("programa-records");
  const [toast, setToast] = useState(null);
  const [lastSaved, setLastSaved] = useState(null);

  const programaEntries = programas.filter(
    (p) => p.fecha === values.fecha && (values.turno ? p.turno === values.turno : true)
  );
  const InicioComponent = area.InicioComponent;

  // Si el usuario entra a editar otro registro, ocultamos el botón de WhatsApp del guardado anterior.
  // (Solo al ENTRAR a editar: al terminar de guardar una edición editingId vuelve a null y el
  // botón de WhatsApp debe seguir visible.)
  useEffect(() => { if (editingId) setLastSaved(null); }, [editingId]);

  const handleSave = async () => {
    if (!values.fecha || !values.turno) {
      setToast({ kind: "error", message: "Completa Fecha y Turno antes de guardar." });
      return;
    }
    if (loading) {
      setToast({ kind: "error", message: "Espera unos segundos a que terminen de cargar los registros e inténtalo de nuevo." });
      return;
    }
    const record = {
      ...values,
      id: editingId ?? Date.now(),
      claveTurno: claveTurno(values.fecha, values.turno),
      responsable: autor || "Sin nombre",
    };
    // Un solo inicio por turno: si ya existe uno de esta fecha y turno (p. ej. por
    // reintentar un guardado que dio error de sincronización), se actualiza ese
    // en vez de agregar un duplicado.
    const existente = editingId ? null : masRecientePorClave(records, record.claveTurno);
    if (existente) record.id = existente.id;
    const updated = editingId
      ? records.map((r) => (r.id === editingId ? record : r))
      : existente
        ? records.filter((r) => r.claveTurno !== record.claveTurno).concat(record)
        : [...records, record];
    const ok = await saveRecords(updated);
    // El registro ya quedó guardado en este dispositivo aunque falle la sincronización con
    // Sheets, así que el botón de WhatsApp se ofrece igual.
    setLastSaved(record);
    if (ok) {
      setToast({ kind: "ok", message: editingId ? "Cambios guardados." : "Registro de inicio guardado." });
      onSaved();
    } else {
      setToast({ kind: "error", message: "No se pudo sincronizar con el almacenamiento. Presiona Guardar nuevamente." });
    }
  };

  const recientes = [...records].sort((a, b) => b.id - a.id).slice(0, 5);

  return (
    <div className="space-y-4 pb-6">
      {editingId && (
        <div className="flex items-center justify-between bg-amber-500/10 border border-amber-500/30 text-amber-700 text-sm rounded-lg px-3 py-2">
          <span>Editando registro del {fmtFecha(values.fecha)} · {values.turno || "—"}</span>
          <button onClick={onCancelEdit} className="text-xs underline">Cancelar edición</button>
        </div>
      )}

      <InicioComponent values={values} setField={setField} programaEntries={programaEntries} editingId={editingId} />

      {toast && <Toast {...toast} />}
      {storageError && <Toast kind="error" message={storageError} />}

      <button
        onClick={handleSave}
        className="sticky bottom-2 z-10 w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-500 text-white font-semibold rounded-xl py-3 transition-colors shadow-lg shadow-blue-900/30"
      >
        <Save size={18} /> {editingId ? "Guardar cambios" : "Guardar inicio de turno"}
      </button>

      {lastSaved && (
        <a
          href={whatsappShareUrl(buildResumenWhatsappInicio(area, lastSaved))}
          target="_blank"
          rel="noopener noreferrer"
          className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-xl py-3 transition-colors"
        >
          <Share2 size={18} /> Compartir resumen por WhatsApp
        </a>
      )}

      <Card title="Últimos registros guardados">
        {loading ? <Loader /> : recientes.length === 0 ? (
          <EmptyNote text="Todavía no hay registros de inicio guardados para esta área." />
        ) : (
          <div className="space-y-2 text-sm">
            {recientes.map((r) => (
              <div key={r.id} className="flex justify-between border-b border-slate-300/50 pb-1 last:border-0 last:pb-0">
                <span className="text-slate-700">{fmtFecha(r.fecha)} · {r.turno}{r.responsable ? ` · ${r.responsable}` : ""}</span>
                <span className="text-slate-500">{fmtDateTime(new Date(r.id).toISOString())}</span>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pestaña: Cierre de turno
// ---------------------------------------------------------------------------
function CierreTab({ area, autor, initialTarget, supervisoresList }) {
  const [inicioRecords, , loadingI] = useSharedList(`${area.key}-inicio-records`);
  const [cierreRecords, saveCierres, loadingC, storageError] = useSharedList(`${area.key}-cierre-records`);
  const [programas] = useSharedList("programa-records");
  const { horarios } = useHorariosTurno();
  const [fecha, setFecha] = useState(today());
  const [turno, setTurno] = useState("");
  const [values, setValues] = useState({});
  const [toast, setToast] = useState(null);

  useEffect(() => {
    if (initialTarget) {
      setFecha(initialTarget.fecha);
      setTurno(initialTarget.turno);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialTarget]);

  const loading = loadingI || loadingC;
  const clave = claveTurno(fecha, turno);
  const inicio = masRecientePorClave(inicioRecords, clave) || undefined;
  const cierreActual = masRecientePorClave(cierreRecords, clave) || undefined;

  useEffect(() => {
    if (loading) return;
    const found = masRecientePorClave(cierreRecords, clave);
    setValues(found ? { ...found } : {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [clave, loading]);

  const setField = (key, val) => setValues((prev) => ({ ...prev, [key]: val }));

  const handleFinalizar = async () => {
    if (!turno) return;
    if (loading) {
      setToast({ kind: "error", message: "Espera unos segundos a que terminen de cargar los datos del turno e inténtalo de nuevo." });
      return;
    }
    const record = {
      ...values,
      id: cierreActual?.id ?? Date.now(),
      claveTurno: clave,
      fecha,
      turno,
      responsable: autor || "Sin nombre",
      estado: "Enviado",
      fechaHoraCierre: new Date().toISOString(),
    };
    // Un solo cierre por turno: se reemplaza el existente (y cualquier duplicado de reintentos).
    const updated = cierreActual ? cierreRecords.filter((r) => r.claveTurno !== clave).concat(record) : [...cierreRecords, record];
    const ok = await saveCierres(updated);
    setValues(record);
    if (ok) {
      setToast({ kind: "ok", message: "Cierre de turno finalizado y guardado." });
    } else {
      setToast({ kind: "error", message: "No se pudo sincronizar con el almacenamiento. Presiona el botón nuevamente." });
    }
  };

  const CierreComponent = area.CierreComponent;
  const programaEntries = programas.filter((p) => p.fecha === fecha && p.turno === turno);

  return (
    <div className="space-y-4 pb-10">
      <Card title="Selecciona el turno a cerrar" step={1}>
        <div className="mb-3">
          <label className="block text-xs font-medium text-slate-600 mb-1">Fecha</label>
          <input type="date" style={{ colorScheme: "light" }} className={inputBase} value={fecha} onChange={(e) => setFecha(e.target.value)} />
        </div>
        <label className="block text-xs font-medium text-slate-600 mb-1">Turno</label>
        <TurnoSelector value={turno} onChange={setTurno} />
      </Card>

      {!turno ? (
        <Card><EmptyNote text="Selecciona el turno para ver el resumen y registrar el cierre." /></Card>
      ) : loading ? (
        <Card><Loader text="Cargando información del turno…" /></Card>
      ) : (
        <>
          <Card title="Resumen del inicio de turno" step={2}>
            {!inicio ? (
              <EmptyNote text={`Sin registro de inicio de turno de ${area.title.toLowerCase()} para esta fecha y turno.`} />
            ) : (
              <DataGrid items={area.resumenInicio(inicio)} />
            )}
          </Card>

          <CierreComponent values={values} setField={setField} inicio={inicio} programaEntries={programaEntries} />

          <Card title="Finalizar cierre">
            {cierreActual?.estado === "Enviado" && (
              <p className="text-sm text-slate-600 mb-3">
                Cierre enviado por <span className="text-slate-800 font-medium">{cierreActual.responsable}</span> el {fmtDateTime(cierreActual.fechaHoraCierre)}
              </p>
            )}
            <button
              onClick={handleFinalizar}
              className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-500 text-white font-semibold rounded-xl py-3 transition-colors"
            >
              <ClipboardCheck size={18} /> {cierreActual?.estado === "Enviado" ? "Actualizar cierre" : "Finalizar cierre"}
            </button>
            {cierreActual?.estado === "Enviado" && (
              <a
                href={whatsappShareUrl(buildResumenWhatsapp(area, cierreActual, inicio, programaEntries, horarios))}
                target="_blank"
                rel="noopener noreferrer"
                className="mt-3 w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-xl py-3 transition-colors"
              >
                <Share2 size={18} /> Compartir resumen por WhatsApp
              </a>
            )}
            {toast && <div className="mt-3"><Toast {...toast} /></div>}
            {storageError && <div className="mt-3"><Toast kind="error" message={storageError} /></div>}
          </Card>
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pestaña: Indicadores semanales
// ---------------------------------------------------------------------------
function IndicadoresTab({ area }) {
  const [inicioRecords, , loadingI] = useSharedList(`${area.key}-inicio-records`);
  const [cierreRecords, , loadingC] = useSharedList(`${area.key}-cierre-records`);
  const loading = loadingI || loadingC;
  const charts = area.detailCharts || [];

  const sorted = [...cierreRecords].sort((a, b) => a.id - b.id).slice(-7);
  const merged = sorted.map((c) => ({ ...(masRecientePorClave(inicioRecords, c.claveTurno) || {}), ...c }));

  return (
    <div className="space-y-4">
      {loading ? (
        <Card><Loader /></Card>
      ) : merged.length === 0 ? (
        <Card><EmptyNote text="Todavía no hay cierres de turno para mostrar la tendencia." /></Card>
      ) : (
        <>
          {charts.map((chart, ci) => {
            const chartData = merged.map((r) => {
              const row = { label: `${fmtFecha(r.fecha)} ${r.turno}` };
              chart.series.forEach((s) => {
                const v = s.compute(r);
                row[s.key] = chart.format === "pct" ? Number((v * 100).toFixed(1)) : Number(v.toFixed(2));
              });
              return row;
            });
            return (
              <Card key={ci} title={chart.title} step={ci === 0 ? 1 : undefined}>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={chartData} margin={{ top: 8, right: 8, left: -16, bottom: 0 }}>
                    <CartesianGrid stroke="#e2e8f0" strokeDasharray="3 3" vertical={false} />
                    <XAxis dataKey="label" stroke="#64748b" fontSize={9} />
                    <YAxis stroke="#64748b" fontSize={10} unit={chart.format === "pct" ? "%" : ""} />
                    <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12, color: "#1e293b" }} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    {chart.series.map((s) => (
                      <Bar key={s.key} dataKey={s.key} name={s.label} fill={s.color} radius={[3, 3, 0, 0]} />
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            );
          })}

          <Card title="Detalle por turno">
            <div className="space-y-3 text-sm">
              {[...merged].reverse().map((r) => (
                <div key={r.id} className="border-b border-slate-300/50 pb-2 last:border-0 last:pb-0">
                  <div className="font-medium text-slate-900 mb-1">{fmtFecha(r.fecha)} · {r.turno}</div>
                  {charts.map((chart, ci) => (
                    <div key={ci} className="text-xs text-slate-600">
                      {chart.series.map((s) => {
                        const v = s.compute(r);
                        return `${s.label}: ${chart.format === "pct" ? fmtPct(v) : fmtNum(v, 2)}`;
                      }).join(" · ")}
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pestaña: Registros (solo Jefe de producción)
// ---------------------------------------------------------------------------
function RegistrosTab({ area, onEdit, onEditCierre }) {
  const [inicioRecords, saveInicio, loadingI] = useSharedList(`${area.key}-inicio-records`);
  const [cierreRecords, saveCierre, loadingC] = useSharedList(`${area.key}-cierre-records`);
  const [confirm, setConfirm] = useState(null);

  const removeInicio = async (id) => { await saveInicio(inicioRecords.filter((r) => r.id !== id)); setConfirm(null); };
  const removeCierre = async (id) => { await saveCierre(cierreRecords.filter((r) => r.id !== id)); setConfirm(null); };

  // Para que la app no se vuelva lenta con tanto historial: solo se listan los
  // registros de la semana en curso (lunes a domingo) y los del último día
  // trabajado de la semana pasada. «Ver historial completo» los muestra todos.
  const [verTodo, setVerTodo] = useState(false);
  const lunes = semanaDe(today()).desde;
  const fechasPrevias = [...inicioRecords, ...cierreRecords].map((r) => r.fecha).filter((f) => f && f < lunes).sort();
  const ultimoDiaPrevio = fechasPrevias[fechasPrevias.length - 1];
  const visible = (r) => verTodo || !r.fecha || r.fecha >= lunes || r.fecha === ultimoDiaPrevio;
  const sortedInicio = inicioRecords.filter(visible).sort((a, b) => b.id - a.id);
  const sortedCierre = cierreRecords.filter(visible).sort((a, b) => b.id - a.id);
  const ocultos = inicioRecords.length + cierreRecords.length - sortedInicio.length - sortedCierre.length;

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-600">
        <span>{verTodo ? "Mostrando todo el historial." : `Mostrando la semana en curso${ultimoDiaPrevio ? ` y el último día trabajado de la semana pasada (${fmtFecha(ultimoDiaPrevio)})` : ""}.`}</span>
        {(verTodo || ocultos > 0) && (
          <button type="button" onClick={() => setVerTodo(!verTodo)} className="shrink-0 font-medium text-blue-600 underline">
            {verTodo ? "Ver menos" : `Ver historial completo (${ocultos} más)`}
          </button>
        )}
      </div>
      <Card title="Registros de inicio de turno">
        {loadingI ? <Loader /> : sortedInicio.length === 0 ? <EmptyNote text="Sin registros." /> : (
          <div className="space-y-2">
            {sortedInicio.map((r) => (
              <div key={r.id} className="flex items-center justify-between border border-slate-300 rounded-lg px-3 py-2 text-sm">
                <div className="text-slate-900 font-medium">{fmtFecha(r.fecha)} · {r.turno}{r.responsable ? <span className="text-slate-500 font-normal"> · {r.responsable}</span> : null}</div>
                <div className="flex items-center gap-2">
                  <button onClick={() => onEdit(r)} aria-label="Editar" className="p-1.5 rounded-md hover:bg-slate-300 text-slate-700">
                    <Pencil size={16} />
                  </button>
                  {confirm?.type === "inicio" && confirm.id === r.id ? (
                    <>
                      <button onClick={() => removeInicio(r.id)} className="text-xs text-red-400 font-medium">Confirmar</button>
                      <button onClick={() => setConfirm(null)} className="text-xs text-slate-600">Cancelar</button>
                    </>
                  ) : (
                    <button onClick={() => setConfirm({ type: "inicio", id: r.id })} aria-label="Eliminar" className="p-1.5 rounded-md hover:bg-slate-300 text-red-400">
                      <Trash2 size={16} />
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card title="Cierres de turno">
        {loadingC ? <Loader /> : sortedCierre.length === 0 ? <EmptyNote text="Sin cierres registrados." /> : (
          <div className="space-y-2">
            {sortedCierre.map((r) => (
              <div key={r.id} className="flex items-center justify-between border border-slate-300 rounded-lg px-3 py-2 text-sm">
                <div>
                  <div className="text-slate-900 font-medium">{fmtFecha(r.fecha)} · {r.turno}</div>
                  <div className="text-xs text-slate-500">{r.estado || "—"} · {r.responsable || "—"}</div>
                </div>
                {confirm?.type === "cierre" && confirm.id === r.id ? (
                  <div className="flex items-center gap-2">
                    <button onClick={() => removeCierre(r.id)} className="text-xs text-red-400 font-medium">Confirmar</button>
                    <button onClick={() => setConfirm(null)} className="text-xs text-slate-600">Cancelar</button>
                  </div>
                ) : (
                  <div className="flex items-center gap-2">
                    <button onClick={() => onEditCierre(r)} aria-label="Editar cierre" className="p-1.5 rounded-md hover:bg-slate-300 text-slate-700">
                      <Pencil size={16} />
                    </button>
                    <button onClick={() => setConfirm({ type: "cierre", id: r.id })} aria-label="Eliminar" className="p-1.5 rounded-md hover:bg-slate-300 text-red-400">
                      <Trash2 size={16} />
                    </button>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla: área (Lavado / Selección / Envasado)
// ---------------------------------------------------------------------------
function AreaScreen({ areaKey, isJefe, autor, onBack, supervisoresList }) {
  const area = AREAS[areaKey];
  const [tab, setTab] = useState("inicio");
  const [values, setValues] = useState({ fecha: today(), turno: "" });
  const [editingId, setEditingId] = useState(null);
  const [cierreTarget, setCierreTarget] = useState(null);

  const setField = (key, val) => setValues((prev) => ({ ...prev, [key]: val }));

  const startEdit = (record) => {
    setValues({ ...record });
    setEditingId(record.id);
    setTab("inicio");
  };

  const startEditCierre = (record) => {
    setCierreTarget({ fecha: record.fecha, turno: record.turno, _ts: Date.now() });
    setTab("cierre");
  };

  const cancelEdit = () => {
    setEditingId(null);
    setValues({ fecha: today(), turno: "" });
  };

  const handleSaved = () => {
    if (editingId) {
      // Se queda en "Inicio de turno" para que siga visible el botón de WhatsApp del guardado.
      setEditingId(null);
      setValues({ fecha: today(), turno: "" });
    } else {
      setValues((v) => ({ fecha: v.fecha, turno: "" }));
    }
  };

  const tabs = [
    { key: "inicio", label: "Inicio de turno", icon: ClipboardList },
    { key: "cierre", label: "Cierre de turno", icon: ClipboardCheck },
    // Entrega de Materiales solo aplica a Selección y Envasado (Lavado no maneja
    // esta entrega entre turnos) — pestaña propia, separada del Cierre de turno.
    ...(area.key === "seleccion" || area.key === "envasado"
      ? [{ key: "entrega", label: "Entrega Materiales", icon: Truck }]
      : []),
    { key: "indicadores", label: "Indicadores", icon: BarChart3 },
    { key: "registros", label: "Registros", icon: Pencil },
  ];

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto">
      <TopBar title={area.title} subtitle="Inicio, cierre e indicadores del turno" onBack={onBack} icon={area.icon} accent={area.accent} />
      <div className="px-4 pt-3 flex gap-2 overflow-x-auto">
        {tabs.map((t) => {
          const Icon = t.icon;
          const active = tab === t.key;
          return (
            <button
              key={t.key}
              onClick={() => setTab(t.key)}
              className={`flex items-center gap-1.5 whitespace-nowrap text-xs font-medium rounded-full px-3 py-1.5 border transition-colors ${
                active ? "bg-blue-500/10 border-blue-400 text-blue-700" : "border-slate-300 text-slate-600 hover:border-slate-500"
              }`}
            >
              <Icon size={14} /> {t.label}
            </button>
          );
        })}
      </div>
      <div className="p-4">
        {tab === "inicio" && (
          <InicioTab area={area} values={values} setField={setField} editingId={editingId} onSaved={handleSaved} onCancelEdit={cancelEdit} autor={autor} />
        )}
        {tab === "cierre" && <CierreTab area={area} autor={autor} initialTarget={cierreTarget} supervisoresList={supervisoresList} />}
        {tab === "entrega" && <EntregaMaterialesTab area={area} autor={autor} supervisoresList={supervisoresList} />}
        {tab === "indicadores" && <IndicadoresTab area={area} />}
        {tab === "registros" && <RegistrosTab area={area} onEdit={startEdit} onEditCierre={startEditCierre} />}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Dashboard del Jefe de producción — vista EJECUTIVA + reporte semanal (PDF)
// El reporte PDF reutiliza exactamente el mismo panel que la pantalla.
// ---------------------------------------------------------------------------
function useAreaCierres(areaKey) {
  const [inicio, , loadingI] = useSharedList(`${areaKey}-inicio-records`);
  const [cierres, , loadingC] = useSharedList(`${areaKey}-cierre-records`);
  return { inicio, cierres, loading: loadingI || loadingC };
}

// ---- Fechas (en hora local, formato AAAA-MM-DD como el resto de la app) ----
const parseIso = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); };
const isoDeFecha = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const sumarDias = (iso, n) => { const d = parseIso(iso); d.setDate(d.getDate() + n); return isoDeFecha(d); };
// Semana de lunes a domingo que contiene la fecha dada.
const semanaDe = (iso) => {
  const dow = (parseIso(iso).getDay() + 6) % 7;
  const desde = sumarDias(iso, -dow);
  return { desde, hasta: sumarDias(desde, 6) };
};
const fmtCorto = (iso) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}`;
const MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const sumaPor = (arr, fn) => arr.reduce((s, x) => s + fn(x), 0);
const cap1 = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1) : s);

// Serie día a día entre `desde` y `hasta`. Con `fnDen` entrega una razón
// (num/den) por día; sin ella, la suma. `valor` es null en días sin datos.
function serieDiaria(registros, desde, hasta, fnNum, fnDen) {
  const out = [];
  for (let f = desde; f <= hasta; f = sumarDias(f, 1)) {
    const rs = registros.filter((r) => r.fecha === f);
    let valor = null;
    if (rs.length > 0) {
      const n = sumaPor(rs, fnNum);
      if (fnDen) { const d = sumaPor(rs, fnDen); valor = d ? n / d : null; }
      else valor = n;
    }
    out.push({ fecha: f, label: fmtCorto(f), valor });
  }
  return out;
}

// Incidentes + accidentes por supervisor (responsable del cierre): por día en
// el período, por semana (últimas 8) y por mes (últimos 6). Usa TODO el
// historial, no solo el período seleccionado, para que las barras semanales y
// mensuales sean comparables.
function seguridadPorSupervisor(datos, desde, hasta) {
  const casos = [];
  [datos.lavado, datos.seleccion, datos.envasado].forEach((a) => a.cierres.forEach((c) => {
    if (c.huboIncidentes !== "Sí" && c.huboAccidentes !== "Sí") return;
    const sup = String(c.responsable || "").trim() || "Sin nombre";
    const tipo = c.huboAccidentes === "Sí" ? "acc" : "inc";
    incidentesDeCierre(c).forEach(() => casos.push({ fecha: c.fecha, sup, tipo }));
  }));

  const semanaFin = semanaDe(hasta);
  const semanas = [];
  for (let k = 7; k >= 0; k--) {
    const w = semanaDe(sumarDias(semanaFin.desde, -7 * k));
    semanas.push({ desde: w.desde, hasta: w.hasta, label: fmtCorto(w.desde) });
  }
  const [yH, mH] = hasta.split("-").map(Number);
  const meses = [];
  for (let k = 5; k >= 0; k--) {
    const d = new Date(yH, mH - 1 - k, 1);
    meses.push({ key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`, label: `${MESES_CORTOS[d.getMonth()]} ${String(d.getFullYear()).slice(2)}` });
  }
  const mesActual = meses[meses.length - 1].key;

  const enPeriodo = casos.filter((c) => c.fecha >= desde && c.fecha <= hasta);
  const enSemanas = casos.filter((c) => c.fecha >= semanas[0].desde && c.fecha <= semanas[semanas.length - 1].hasta);
  const enMeses = casos.filter((c) => meses.some((m) => c.fecha.startsWith(m.key)));
  const nombres = [...new Set([...enPeriodo, ...enSemanas, ...enMeses].map((c) => c.sup))];
  const total = (s) => casos.filter((c) => c.sup === s && (
    (c.fecha >= desde && c.fecha <= hasta) || (c.fecha >= semanas[0].desde && c.fecha <= semanas[7].hasta) || c.fecha.startsWith(mesActual)
  )).length;
  const supervisores = nombres.sort((a, b) => total(b) - total(a) || a.localeCompare(b));
  const colores = {};
  supervisores.forEach((s, i) => { colores[s] = SERIES_COLORS[i % SERIES_COLORS.length]; });

  const fila = (label, lista) => {
    const o = { label };
    supervisores.forEach((s) => { o[s] = lista.filter((c) => c.sup === s).length; });
    return o;
  };
  const porDia = [];
  for (let f = desde; f <= hasta; f = sumarDias(f, 1)) porDia.push(fila(fmtCorto(f), enPeriodo.filter((c) => c.fecha === f)));
  const porSemana = semanas.map((w) => fila(w.label, casos.filter((c) => c.fecha >= w.desde && c.fecha <= w.hasta)));
  const porMes = meses.map((m) => fila(m.label, casos.filter((c) => c.fecha.startsWith(m.key))));

  const tabla = supervisores.map((s) => {
    const cuenta = (lista, tipo) => lista.filter((c) => c.sup === s && c.tipo === tipo).length;
    const delMes = casos.filter((c) => c.fecha.startsWith(mesActual));
    return {
      sup: s,
      perInc: cuenta(enPeriodo, "inc"), perAcc: cuenta(enPeriodo, "acc"),
      mesInc: cuenta(delMes, "inc"), mesAcc: cuenta(delMes, "acc"),
    };
  });

  return {
    supervisores, colores, porDia, porSemana, porMes, tabla,
    hayDia: porDia.some((r) => supervisores.some((s) => r[s] > 0)),
    haySemana: porSemana.some((r) => supervisores.some((s) => r[s] > 0)),
    hayMes: porMes.some((r) => supervisores.some((s) => r[s] > 0)),
    mesLabel: meses[meses.length - 1].label,
  };
}

const TODOS_LOS_TURNOS = ["T1", "T2", "T3"];

// Filtra por turno los registros de las 3 áreas y el programa (filtro general del dashboard).
function filtrarDatosPorTurno(datos, turnos) {
  if (turnos.length >= TODOS_LOS_TURNOS.length) return datos;
  const ok = (r) => turnos.includes(r.turno);
  const area = (a) => ({ ...a, inicio: a.inicio.filter(ok), cierres: a.cierres.filter(ok) });
  return { ...datos, lavado: area(datos.lavado), seleccion: area(datos.seleccion), envasado: area(datos.envasado), programas: datos.programas.filter(ok) };
}

// Calcula TODOS los indicadores de un período (cierres con su inicio fusionado).
function resumenPeriodo(datos, desde, hasta, turnosSel = TODOS_LOS_TURNOS) {
  const { lavado, seleccion, envasado, programas, horarios } = datos;
  const enRango = (r) => r.fecha && r.fecha >= desde && r.fecha <= hasta;
  const fusionados = (a) => a.cierres.filter(enRango).map((c) => ({ ...(masRecientePorClave(a.inicio, c.claveTurno) || {}), ...c }));
  const progDe = (c) => programas.filter((p) => p.fecha === c.fecha && p.turno === c.turno);
  const hoy = today();
  const sinCierre = (a) => a.inicio.filter((i) => enRango(i) && i.fecha < hoy && !a.cierres.some((c) => c.claveTurno === i.claveTurno)).length;

  // ---- Lavado ----
  const lav = fusionados(lavado);
  const lavTotal = sumaPor(lav, (r) => totalPalletsFromPrefix(r, "lavados"));
  const lavTipos = TIPOS_BANDEJA
    .map((t) => ({ key: t.key, label: t.label.replace(/\s*\(.*\)/, ""), valor: sumaPor(lav, (r) => palletsValue(r, "lavados", t)) }))
    .filter((t) => t.valor > 0);
  const lavDiario = serieDiaria(lav, desde, hasta, (r) => totalPalletsFromPrefix(r, "lavados"));
  // Pallets lavados por día, separados por tipo de bandeja (barras apiladas).
  const nombreTipo = (t) => t.label.replace(/\s*\(.*\)/, "");
  const lavColores = {};
  TIPOS_BANDEJA.forEach((t, i) => { lavColores[nombreTipo(t)] = SERIES_COLORS[i % SERIES_COLORS.length]; });
  const lavTiposKeys = lavTipos.map((t) => t.label);
  const lavDiarioTipos = [];
  for (let f = desde; f <= hasta; f = sumarDias(f, 1)) {
    const rs = lav.filter((r) => r.fecha === f);
    const fila = { label: fmtCorto(f) };
    TIPOS_BANDEJA.forEach((t) => { fila[nombreTipo(t)] = Math.round(sumaPor(rs, (r) => palletsValue(r, "lavados", t)) * 10) / 10; });
    lavDiarioTipos.push(fila);
  }
  // Pallets pendientes por lavar que quedaron al cierre de cada turno.
  const rangoTurno = { T3: 0, T1: 1, T2: 2 };
  const lavOrden = [...lav].sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)) || (rangoTurno[a.turno] ?? 9) - (rangoTurno[b.turno] ?? 9) || (a.id || 0) - (b.id || 0));
  const pendDe = (c) => c ? {
    fecha: c.fecha, turno: c.turno,
    total: Math.round(totalPalletsManual(c, "pendientes") * 10) / 10,
    tipos: TIPOS_BANDEJA.map((t) => ({ key: t.key, label: nombreTipo(t), valor: num(c[`pendientes_${t.key}_pallets`]) })).filter((t) => t.valor > 0),
  } : null;
  const ultimoPorDia = (lista) => {
    const m = new Map();
    lista.forEach((c) => m.set(c.fecha, c));
    return [...m.values()];
  };
  const pendTotalDe = (c) => totalPalletsManual(c, "pendientes");
  const pendientes = {
    ult: pendDe(lavOrden[lavOrden.length - 1]),
    porTurno: {},
    diario: { ult: serieDiaria(ultimoPorDia(lavOrden), desde, hasta, pendTotalDe) },
  };
  ["T1", "T2", "T3"].forEach((tk) => {
    const delTurno = lavOrden.filter((c) => c.turno === tk);
    pendientes.porTurno[tk] = pendDe(delTurno[delTurno.length - 1]);
    pendientes.diario[tk] = serieDiaria(ultimoPorDia(delTurno), desde, hasta, pendTotalDe);
  });

  // ---- Selección (total y por línea) ----
  const sel = fusionados(seleccion);
  const totSel = (r) => LINEAS_SELECCION.reduce((a, l) => {
    if (r[`linea_${l.key}_activa`] === "Sí") { const t = lineaKgTotales(r, l); a.ing += t.ing; a.apr += t.apr; }
    return a;
  }, { ing: 0, apr: 0 });
  const selLineas = LINEAS_SELECCION.map((l) => ({ key: l.key, label: l.label, ing: 0, apr: 0, ingCump: 0, prog: 0 }));
  let selIngCump = 0, selProg = 0;
  sel.forEach((c) => {
    const prog = progDe(c);
    const cu = computeSeleccionCumplimiento(c, c, prog);
    if (cu.kgProgramado > 0) { selIngCump += cu.kgIngresado; selProg += cu.kgProgramado; }
    LINEAS_SELECCION.forEach((l, i) => {
      if (c[`linea_${l.key}_activa`] !== "Sí") return;
      const t = lineaKgTotales(c, l);
      selLineas[i].ing += t.ing;
      selLineas[i].apr += t.apr;
      const kp = sumaPor(prog.filter((p) => (p.lineaKey || p.linea) === l.key && p.especie !== "LAVADO"), (p) => num(p.cantidad));
      if (kp > 0) { selLineas[i].ingCump += t.ing; selLineas[i].prog += kp; }
    });
  });
  selLineas.forEach((ln, i) => {
    const L = LINEAS_SELECCION[i];
    const act = sel.filter((c) => c[`linea_${L.key}_activa`] === "Sí");
    ln.rend = ln.ing ? ln.apr / ln.ing : null;
    ln.cump = ln.prog ? ln.ingCump / ln.prog : null;
    ln.diarioIng = serieDiaria(act, desde, hasta, (r) => lineaKgTotales(r, L).ing);
    ln.diarioRend = serieDiaria(act, desde, hasta, (r) => lineaKgTotales(r, L).apr, (r) => lineaKgTotales(r, L).ing);
  });
  const selIng = sumaPor(selLineas, (l) => l.ing);
  const selApr = sumaPor(selLineas, (l) => l.apr);
  const selRend = selIng ? selApr / selIng : null;
  const selCump = selProg ? selIngCump / selProg : null;

  // ---- Envasado (Envasadora y Línea 5) ----
  const env = fusionados(envasado);
  let cajasProd = 0, cajasProg = 0, cajasTeor = 0, cajasReal = 0, bolsasTeor = 0, bolsasReal = 0;
  let teTotal = 0, bolsasConTE = 0, l5Ing = 0, l5Apr = 0, l5IngCump = 0, l5Prog = 0, detTotal = 0;
  const detMap = {};
  env.forEach((c) => {
    (c.detenciones || []).forEach((d) => {
      const m = Number(d && d.minutos) || 0;
      if (!m) return;
      const motivo = cap1(String((d && d.motivo) || "").trim().toLowerCase()) || "Sin motivo";
      detMap[motivo] = (detMap[motivo] || 0) + m;
      detTotal += m;
    });
    if (c.activa_envasadora === "Sí") {
      cajasProd += num(c.cajasProducidas); cajasProg += num(c.cajasProgramadas);
      cajasTeor += num(c.cajasTeoricas); cajasReal += num(c.cajasConsumidasReal);
      bolsasTeor += num(c.bolsasTeoricas); bolsasReal += num(c.bolsasConsumidasReal);
      const te = computeTiempoEfectivoEnvasado(c, horarios);
      if (te.tiempoEfectivo) { teTotal += te.tiempoEfectivo; bolsasConTE += num(c.bolsasConsumidasReal); }
    }
    if (c.activa_linea5 === "Sí") {
      const m = computeEnvasadoMetrics(c);
      l5Ing += m.l5Ing; l5Apr += m.l5Apr;
      const kp = sumaPor(progDe(c).filter((p) => (p.lineaKey === "linea5" || p.linea === "linea5") && p.especie !== "LAVADO"), (p) => num(p.cantidad));
      if (kp > 0) { l5IngCump += m.l5Ing; l5Prog += kp; }
    }
  });
  const detPareto = Object.entries(detMap).map(([motivo, min]) => ({ motivo, min })).sort((a, b) => b.min - a.min);
  const envasadoraActiva = env.filter((c) => c.activa_envasadora === "Sí");
  const l5Recs = env.filter((c) => c.activa_linea5 === "Sí");

  // ---- Seguridad y operación ----
  const todos = [
    ...lav.map((c) => ({ ...c, area: "Lavado de bandejas" })),
    ...sel.map((c) => ({ ...c, area: "Selección" })),
    ...env.map((c) => ({ ...c, area: "Envasado" })),
  ];
  const incidentes = todos.filter((c) => c.huboIncidentes === "Sí").length;
  const accidentes = todos.filter((c) => c.huboAccidentes === "Sí").length;
  const eventos = todos
    .filter((c) => c.huboIncidentes === "Sí" || c.huboAccidentes === "Sí")
    .sort((a, b) => new Date(b.fechaHoraCierre || 0) - new Date(a.fechaHoraCierre || 0))
    .flatMap((c) => incidentesDeCierre(c).map((it, i) => ({ ...c, ...it, itemIdx: i })));
  const docsPendientes = eventos.filter((e) => (e.docHecha || "No") !== "Sí" || (e.docEntregada || "No") !== "Sí").length;
  const turnosSinCierre = { lavado: sinCierre(lavado), seleccion: sinCierre(seleccion), envasado: sinCierre(envasado) };

  const comentarios = todos
    .flatMap((c) => ["Producción", "Calidad", "Mantención", "Seguridad"].map((t) => ({
      area: c.area, fecha: c.fecha, turno: c.turno, tema: t, texto: String(c[`com_${t}`] || "").trim(),
    })))
    .filter((x) => x.texto)
    .sort((a, b) => `${b.fecha}${b.turno}`.localeCompare(`${a.fecha}${a.turno}`));

  return {
    desde, hasta, turnosSel,
    turnosCerrados: { lavado: lav.length, seleccion: sel.length, envasado: env.length },
    lavado: { total: lavTotal, promedio: lav.length ? lavTotal / lav.length : null, tipos: lavTipos, diario: lavDiario, diarioTipos: lavDiarioTipos, tiposKeys: lavTiposKeys, colores: lavColores, pendientes },
    seleccion: {
      ing: selIng, apr: selApr, prog: selProg, rend: selRend, cump: selCump, lineas: selLineas,
    },
    envasado: {
      cajasProd, cajasProg,
      cump: cajasProg ? cajasProd / cajasProg : null,
      mermaCajas: cajasTeor ? (cajasReal - cajasTeor) / cajasTeor : null,
      mermaBolsas: bolsasTeor ? (bolsasReal - bolsasTeor) / bolsasTeor : null,
      tiempoEfectivo: teTotal || null,
      productividad: teTotal ? bolsasConTE / teTotal : null,
      detTotal, detPareto,
      diarioCump: serieDiaria(envasadoraActiva, desde, hasta, (r) => num(r.cajasProducidas), (r) => num(r.cajasProgramadas)),
      diarioCajas: serieDiaria(envasadoraActiva, desde, hasta, (r) => num(r.cajasProducidas)),
      l5Ing, l5Apr, l5Prog,
      l5Rend: l5Ing ? l5Apr / l5Ing : null,
      l5Cump: l5Prog ? l5IngCump / l5Prog : null,
      l5DiarioIng: serieDiaria(l5Recs, desde, hasta, (r) => computeEnvasadoMetrics(r).l5Ing),
      l5DiarioRend: serieDiaria(l5Recs, desde, hasta, (r) => computeEnvasadoMetrics(r).l5Apr, (r) => computeEnvasadoMetrics(r).l5Ing),
    },
    seguridad: { incidentes, accidentes, docsPendientes, eventos, supervisores: seguridadPorSupervisor(datos, desde, hasta) },
    turnosSinCierreTotal: turnosSinCierre.lavado + turnosSinCierre.seleccion + turnosSinCierre.envasado,
    turnosSinCierre,
    comentarios,
  };
}

// Variación entre período actual y anterior. pct → puntos porcentuales; num → %.
function variacion(actual, previo, tipo) {
  if (actual == null || previo == null) return null;
  if (tipo === "pct") {
    const d = (actual - previo) * 100;
    if (Math.abs(d) < 0.05) return { txt: "=", signo: 0 };
    return { txt: `${d > 0 ? "+" : ""}${fmtNum(d, 1)} pp`, signo: d > 0 ? 1 : -1 };
  }
  if (!previo) return actual ? { txt: "nuevo", signo: 1 } : { txt: "=", signo: 0 };
  const d = ((actual - previo) / Math.abs(previo)) * 100;
  if (Math.abs(d) < 0.5) return { txt: "=", signo: 0 };
  return { txt: `${d > 0 ? "+" : ""}${fmtNum(d, 0)}%`, signo: d > 0 ? 1 : -1 };
}

function fmtValor(v, tipo) {
  if (v == null) return "—";
  if (tipo === "pct") return fmtPct(v);
  if (tipo === "num1") return fmtNum(v, 1);
  return fmtNum(v, 0);
}

// Verde ≥95 %, ámbar ≥85 %, rojo por debajo (solo para cumplimiento vs programa).
function tonoCumplimiento(v) {
  if (v == null) return undefined;
  return v >= 0.95 ? "ok" : v >= 0.85 ? "warn" : "bad";
}

const TONO_CLASES = {
  ok: "border-emerald-300 bg-emerald-50",
  warn: "border-amber-300 bg-amber-50",
  bad: "border-red-300 bg-red-50",
  none: "border-slate-200 bg-white",
};

function DeltaChip({ actual, previo, tipo, menorEsMejor }) {
  const v = variacion(actual, previo, tipo);
  if (!v) return null;
  const bueno = v.signo === 0 ? null : menorEsMejor ? v.signo < 0 : v.signo > 0;
  const color = bueno === null ? "text-slate-500" : bueno ? "text-emerald-700" : "text-red-600";
  const flecha = v.signo === 0 ? "" : v.signo > 0 ? "▲ " : "▼ ";
  return <span className={`text-[11px] font-medium ${color}`}>{flecha}{v.txt} <span className="text-slate-400 font-normal">vs. anterior</span></span>;
}

function Kpi({ label, valor, tipo = "num", unidad = "", sub, actual, previo, menorEsMejor, tono, grande }) {
  return (
    <div className={`rounded-xl border px-3 py-2.5 ${TONO_CLASES[tono || "none"]}`}>
      <div className="text-[11px] font-medium text-slate-500 leading-tight">{label}</div>
      <div className={`${grande ? "text-3xl" : "text-xl"} font-bold text-slate-900 mt-0.5`}>
        {fmtValor(valor, tipo)}{valor != null && unidad ? <span className="text-xs font-medium text-slate-500 ml-1">{unidad}</span> : null}
      </div>
      {sub && <div className="text-[11px] text-slate-500">{sub}</div>}
      {actual !== undefined && <div><DeltaChip actual={actual} previo={previo} tipo={tipo === "pct" ? "pct" : "num"} menorEsMejor={menorEsMejor} /></div>}
    </div>
  );
}

// ---- Gráficos SVG (se escalan solos; usados en el reporte PDF y en seguridad) ----
function SvgBarras({ data, formato = "num", color = "#3b82f6", alto = 100, ancho = 560 }) {
  const W = ancho;
  const vals = data.map((d) => d.valor).filter((v) => v != null);
  if (vals.length === 0) return <div className="text-[11px] text-slate-400 italic py-2">Sin datos en el período.</div>;
  const max = Math.max(Math.max(...vals) * 1.15, formato === "pct" ? 1 : 0.0001);
  const n = data.length;
  const bw = W / n;
  const fmtV = (v) => (formato === "pct" ? `${fmtNum(v * 100, 0)}%` : fmtNum(v, 0));
  return (
    <svg viewBox={`0 0 ${W} ${alto + 26}`} className="w-full" role="img">
      <line x1="0" y1={alto} x2={W} y2={alto} stroke="#cbd5e1" strokeWidth="1" />
      {data.map((d, i) => {
        const h = d.valor == null ? 0 : (d.valor / max) * (alto - 14);
        const x = i * bw + bw * 0.18;
        return (
          <g key={d.fecha}>
            {d.valor != null && <rect x={x} y={alto - h} width={bw * 0.64} height={Math.max(h, 1)} fill={color} rx="2" />}
            {d.valor != null && n <= 14 && <text x={i * bw + bw / 2} y={alto - h - 3} textAnchor="middle" fontSize="9" fill="#334155">{fmtV(d.valor)}</text>}
            {(n <= 14 || i % 2 === 0) && <text x={i * bw + bw / 2} y={alto + 12} textAnchor="middle" fontSize="9" fill="#64748b">{d.label}</text>}
          </g>
        );
      })}
    </svg>
  );
}

// Barras apiladas por supervisor (una pila por día / semana / mes).
function SvgApilado({ data, keys, colores, alto = 110 }) {
  const W = 560;
  const totales = data.map((d) => keys.reduce((s, k) => s + (d[k] || 0), 0));
  const max = Math.max(...totales, 1) * 1.2;
  const n = data.length;
  const bw = W / n;
  return (
    <svg viewBox={`0 0 ${W} ${alto + 26}`} className="w-full" role="img">
      <line x1="0" y1={alto} x2={W} y2={alto} stroke="#cbd5e1" strokeWidth="1" />
      {data.map((d, i) => {
        let acc = 0;
        const x = i * bw + bw * 0.18;
        return (
          <g key={`${d.label}-${i}`}>
            {keys.map((k) => {
              const v = d[k] || 0;
              if (!v) return null;
              const h = (v / max) * (alto - 14);
              const y = alto - acc - h;
              acc += h;
              return <rect key={k} x={x} y={y} width={bw * 0.64} height={h} fill={colores[k]} />;
            })}
            {totales[i] > 0 && <text x={i * bw + bw / 2} y={alto - acc - 3} textAnchor="middle" fontSize="9" fontWeight="600" fill="#334155">{Math.round(totales[i] * 10) / 10}</text>}
            {(n <= 14 || i % 2 === 0) && <text x={i * bw + bw / 2} y={alto + 12} textAnchor="middle" fontSize="9" fill="#64748b">{d.label}</text>}
          </g>
        );
      })}
    </svg>
  );
}

function LeyendaSupervisores({ keys, colores }) {
  return (
    <div className="flex flex-wrap gap-x-3 gap-y-1 mt-1">
      {keys.map((k) => (
        <span key={k} className="flex items-center gap-1 text-[11px] text-slate-600">
          <span className="inline-block w-2.5 h-2.5 rounded-sm" style={{ background: colores[k] }} />{k}
        </span>
      ))}
    </div>
  );
}

function GraficoApilado({ titulo, data, keys, colores, hay }) {
  return (
    <div className="rep-bloque">
      <p className="text-xs font-semibold text-slate-700 mb-1">{titulo}</p>
      {!hay ? <EmptyNote text="Sin incidentes ni accidentes en este rango." /> : (
        <>
          <SvgApilado data={data} keys={keys} colores={colores} />
          <LeyendaSupervisores keys={keys} colores={colores} />
        </>
      )}
    </div>
  );
}

// Gráfico de barras diario: recharts en pantalla, SVG en el reporte PDF.
function GraficoDiario({ titulo, data, formato = "num", color, svg, mitad }) {
  const hay = data.some((x) => x.valor != null);
  if (svg) {
    return (
      <div className="rep-bloque">
        <p className="text-xs font-semibold text-slate-700 mb-1">{titulo}</p>
        {hay ? <SvgBarras data={data} formato={formato} color={color} ancho={mitad ? 300 : 560} /> : <div className="text-[11px] text-slate-400 italic py-2">Sin datos en el período.</div>}
      </div>
    );
  }
  const d = data.map((x) => ({
    label: x.label,
    valor: x.valor == null ? null : formato === "pct" ? Number((x.valor * 100).toFixed(1)) : Number(x.valor.toFixed(1)),
  }));
  return (
    <div>
      <p className="text-xs font-semibold text-slate-700 mb-1">{titulo}</p>
      {!hay ? <EmptyNote text="Sin datos en el período." /> : (
        <ResponsiveContainer width="100%" height={150}>
          <BarChart data={d} margin={{ top: 6, right: 6, left: -18, bottom: 0 }}>
            <CartesianGrid stroke="#e2e8f0" strokeDasharray="3 3" vertical={false} />
            <XAxis dataKey="label" stroke="#64748b" fontSize={9} interval="preserveStartEnd" minTickGap={14} />
            <YAxis stroke="#64748b" fontSize={10} unit={formato === "pct" ? "%" : ""} />
            <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12, color: "#1e293b" }} />
            <Bar dataKey="valor" name={titulo} fill={color} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      )}
    </div>
  );
}

// Barras apiladas por día (un color por tipo de bandeja): SVG en el reporte, recharts en pantalla.
function GraficoApiladoTipos({ titulo, data, keys, colores, svg, vacio }) {
  const hay = data.some((d) => keys.some((k) => d[k] > 0));
  return (
    <div className="rep-bloque">
      <p className="text-xs font-semibold text-slate-700 mb-1">{titulo}</p>
      {!hay ? <EmptyNote text={vacio || "Sin datos en el período."} /> : svg ? (
        <>
          <SvgApilado data={data} keys={keys} colores={colores} />
          <LeyendaSupervisores keys={keys} colores={colores} />
        </>
      ) : (
        <ResponsiveContainer width="100%" height={190}>
          <BarChart data={data} margin={{ top: 6, right: 6, left: -18, bottom: 0 }}>
            <CartesianGrid stroke="#e2e8f0" strokeDasharray="3 3" vertical={false} />
            <XAxis dataKey="label" stroke="#64748b" fontSize={9} interval="preserveStartEnd" minTickGap={14} />
            <YAxis stroke="#64748b" fontSize={10} allowDecimals={false} />
            <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12, color: "#1e293b" }} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            {keys.map((k) => <Bar key={k} dataKey={k} name={k} stackId="tipos" fill={colores[k]} />)}
          </BarChart>
        </ResponsiveContainer>
      )}
    </div>
  );
}

// Pallets pendientes por lavar al cierre de los turnos elegidos en el filtro general.
function PendientesLavado({ pend, turnosSel, svg }) {
  const detalle = (d) => d.tipos.length ? d.tipos.map((t) => `${t.label}: ${fmtNum(t.valor, 1)}`).join(" · ") : "Sin pallets pendientes";
  const lista = TODOS_LOS_TURNOS.filter((k) => turnosSel.includes(k));
  return (
    <div className="rep-bloque space-y-3">
      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Pallets pendientes por lavar · último cierre del período</p>
      <div className={`grid gap-2 ${lista.length === 1 ? "grid-cols-1" : lista.length === 2 ? "grid-cols-2" : "grid-cols-3"}`}>
        {lista.map((tk) => {
          const d = pend.porTurno[tk];
          return (
            <div key={tk} className={`rounded-xl border px-3 py-2.5 ${d ? TONO_CLASES[d.total > 0 ? "warn" : "ok"] : TONO_CLASES.none}`}>
              <div className="text-[11px] font-medium text-slate-500">{turnoLabel(tk)}</div>
              <div className="text-2xl font-bold text-slate-900">{d ? fmtNum(d.total, 1) : "—"}{d ? <span className="text-xs font-medium text-slate-500 ml-1">pallets</span> : null}</div>
              <div className="text-[10px] text-slate-500">{d ? `Cierre ${fmtFecha(d.fecha)} · ${detalle(d)}` : "Sin cierres en el período"}</div>
            </div>
          );
        })}
      </div>
      <GraficoDiario svg={svg} titulo="Pendientes al cierre por día" data={pend.diario.ult} color="#f59e0b" />
    </div>
  );
}

function ParetoDetenciones({ items, total }) {
  if (!items.length) return <EmptyNote text="Sin detenciones registradas en el período." />;
  const max = items[0].min;
  return (
    <div className="space-y-2">
      {items.slice(0, 8).map((it) => (
        <div key={it.motivo}>
          <div className="flex justify-between text-xs text-slate-700">
            <span className="truncate pr-2">{it.motivo}</span>
            <span className="shrink-0 font-medium">{fmtNum(it.min)} min · {fmtPct(total ? it.min / total : 0)}</span>
          </div>
          <div className="h-2 rounded bg-slate-100">
            <div className="h-2 rounded" style={{ width: `${(it.min / max) * 100}%`, background: "#f59e0b" }} />
          </div>
        </div>
      ))}
    </div>
  );
}

function SeccionArea({ titulo, color, children, salto }) {
  return (
    <div className={`${salto ? "rep-salto " : ""}rounded-2xl border border-slate-200 bg-white overflow-hidden`}>
      <div className="px-4 py-2.5 text-sm font-bold text-white" style={{ background: color }}>{titulo}</div>
      <div className="p-3 space-y-4">{children}</div>
    </div>
  );
}

// Bloque interior de una sección (una línea de Selección, Envasadora, Línea 5…).
function SubBloque({ titulo, color, children }) {
  return (
    <div className="rep-bloque rounded-xl border border-slate-200 p-3 space-y-3">
      <div className="flex items-center gap-2">
        <span className="inline-block w-1.5 h-4 rounded-sm" style={{ background: color }} />
        <p className="text-sm font-bold text-slate-900">{titulo}</p>
      </div>
      {children}
    </div>
  );
}

function SelectorPeriodo({ modo, setModo, offset, setOffset, desde, hasta }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-3 flex flex-wrap items-center gap-3 justify-between">
      <div className="flex items-center gap-1">
        <button onClick={() => setOffset(offset - 1)} aria-label="Semana anterior" className="p-2 rounded-lg hover:bg-slate-100 text-slate-700"><ChevronLeft size={18} /></button>
        <div className="text-center min-w-[170px]">
          <div className="text-sm font-semibold text-slate-900">{fmtFecha(desde)} – {fmtFecha(hasta)}</div>
          <div className="text-[11px] text-slate-500">{offset === 0 ? "Semana actual" : offset === -1 ? "Semana pasada" : `Hace ${-offset} semanas`}{modo === "mes" ? " · 4 semanas" : ""}</div>
        </div>
        <button onClick={() => setOffset(Math.min(0, offset + 1))} disabled={offset >= 0} aria-label="Semana siguiente" className="p-2 rounded-lg hover:bg-slate-100 text-slate-700 disabled:opacity-30"><ChevronRight size={18} /></button>
      </div>
      <div className="flex rounded-lg border border-slate-300 overflow-hidden text-sm">
        {[["semana", "Semana"], ["mes", "4 semanas"]].map(([k, l]) => (
          <button key={k} onClick={() => setModo(k)} className={`px-3 py-1.5 ${modo === k ? "bg-slate-800 text-white" : "bg-white text-slate-700"}`}>{l}</button>
        ))}
      </div>
    </div>
  );
}

// Filtro general de turnos: aplica a TODOS los indicadores (y al reporte PDF).
function SelectorTurnos({ turnos, setTurnos }) {
  const todos = turnos.length === TODOS_LOS_TURNOS.length;
  const alternar = (k) => {
    if (todos) { setTurnos([k]); return; } // desde «Todos», elegir un turno deja solo ese
    if (turnos.includes(k)) { if (turnos.length > 1) setTurnos(turnos.filter((t) => t !== k)); }
    else setTurnos(TODOS_LOS_TURNOS.filter((t) => t === k || turnos.includes(t)));
  };
  const clase = (on) => `px-3 py-1.5 text-sm border-l first:border-l-0 border-slate-300 ${on ? "bg-slate-800 text-white" : "bg-white text-slate-700"}`;
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-3 flex flex-wrap items-center gap-3 justify-between">
      <div>
        <div className="text-sm font-semibold text-slate-900">Turnos</div>
        <div className="text-[11px] text-slate-500">Aplica a todos los indicadores</div>
      </div>
      <div className="flex rounded-lg border border-slate-300 overflow-hidden">
        <button type="button" onClick={() => setTurnos(TODOS_LOS_TURNOS)} className={clase(todos)}>Todos</button>
        {TODOS_LOS_TURNOS.map((k) => (
          <button key={k} type="button" onClick={() => alternar(k)} className={clase(!todos && turnos.includes(k))}>
            {turnoLabel(k)}
          </button>
        ))}
      </div>
    </div>
  );
}

function ListaEventos({ eventos }) {
  if (eventos.length === 0) return <EmptyNote text="Sin incidentes ni accidentes en el período." />;
  return (
    <div className="space-y-2">
      {eventos.map((c) => {
        const tipos = [c.huboAccidentes === "Sí" ? "Accidente" : null, c.huboIncidentes === "Sí" ? "Incidente" : null].filter(Boolean).join(" + ");
        const docHecha = c.docHecha || "No";
        const docEntregada = c.docEntregada || "No";
        const pendiente = docHecha !== "Sí" || docEntregada !== "Sí";
        return (
          <div key={`${c.area}-${c.id}-${c.itemIdx}`} className={`rep-bloque rounded-xl px-3 py-2.5 border ${pendiente ? "bg-red-50 border-red-200" : "bg-emerald-50 border-emerald-200"}`}>
            <div className="flex justify-between items-start gap-2">
              <span className="text-sm font-semibold text-slate-900">{tipos} · {c.area}</span>
              <span className="text-xs text-slate-600 shrink-0">{fmtFecha(c.fecha)} · {turnoLabel(c.turno)}</span>
            </div>
            <div className="text-xs text-slate-600 mt-1">Supervisor responsable: <b>{c.responsable || "—"}</b></div>
            <div className="text-xs text-slate-600">Persona involucrada: <b>{c.nombre || "—"}</b></div>
            {c.descripcion && <div className="text-xs text-slate-600 mt-1">"{c.descripcion}"</div>}
            <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2 text-xs">
              <span className={docHecha === "Sí" ? "text-emerald-700 font-medium" : "text-red-700 font-semibold"}>{docHecha === "Sí" ? "✓" : "⚠"} Documentación hecha: {docHecha}</span>
              <span className={docEntregada === "Sí" ? "text-emerald-700 font-medium" : "text-red-700 font-semibold"}>{docEntregada === "Sí" ? "✓" : "⚠"} Documentación entregada: {docEntregada}</span>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// Panel de indicadores: es lo que se ve en la pantalla Y lo que va en el PDF.
// `svg` = true en el reporte (gráficos que se escalan al ancho de la hoja).
function PanelIndicadores({ r, p, svg }) {
  const cerrados = r.turnosCerrados.lavado + r.turnosCerrados.seleccion + r.turnosCerrados.envasado;
  const lineasSel = r.seleccion.lineas
    .map((l, i) => ({ l, prev: p.seleccion.lineas[i] }))
    .filter(({ l }) => l.ing > 0 || l.apr > 0 || l.prog > 0);
  const sg = r.seguridad.supervisores;

  return (
    <div className="space-y-4">
      {/* ---------- Resumen ejecutivo ---------- */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
        <Kpi grande label="Pallets lavados" valor={r.lavado.total} actual={r.lavado.total} previo={p.lavado.total} sub={`${cerrados} turnos cerrados`} />
        <Kpi grande label="Rendimiento Selección" tipo="pct" valor={r.seleccion.rend} actual={r.seleccion.rend} previo={p.seleccion.rend} sub={`${fmtNum(r.seleccion.ing)} Kg ingresados`} />
        <Kpi grande label="Cumplimiento Envasadora" tipo="pct" valor={r.envasado.cump} tono={tonoCumplimiento(r.envasado.cump)} actual={r.envasado.cump} previo={p.envasado.cump} sub={`${fmtNum(r.envasado.cajasProd)} / ${fmtNum(r.envasado.cajasProg)} cajas`} />
        <Kpi grande label="Incidentes / accidentes" valor={r.seguridad.incidentes + r.seguridad.accidentes} tono={r.seguridad.accidentes > 0 ? "bad" : r.seguridad.incidentes > 0 ? "warn" : "ok"} sub={`${r.seguridad.incidentes} incidentes · ${r.seguridad.accidentes} accidentes`} actual={r.seguridad.incidentes + r.seguridad.accidentes} previo={p.seguridad.incidentes + p.seguridad.accidentes} menorEsMejor />
      </div>

      {r.turnosSinCierreTotal > 0 && (
        <div className="flex items-start gap-2 rounded-xl border border-amber-300 bg-amber-50 px-3 py-2.5 text-sm text-amber-800">
          <AlertTriangle size={16} className="mt-0.5 shrink-0" />
          <span>
            {r.turnosSinCierreTotal} turno{r.turnosSinCierreTotal > 1 ? "s" : ""} con inicio registrado y <b>sin cierre</b>
            {" "}(Lavado {r.turnosSinCierre.lavado} · Selección {r.turnosSinCierre.seleccion} · Envasado {r.turnosSinCierre.envasado}).
          </span>
        </div>
      )}

      {/* ---------- Lavado ---------- */}
      <SeccionArea titulo="Lavado de bandejas" color="#2563eb">
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
          <Kpi label="Pallets lavados" valor={r.lavado.total} actual={r.lavado.total} previo={p.lavado.total} />
          <Kpi label="Promedio por turno" tipo="num1" valor={r.lavado.promedio} actual={r.lavado.promedio} previo={p.lavado.promedio} />
          <Kpi label="Turnos cerrados" valor={r.turnosCerrados.lavado} actual={r.turnosCerrados.lavado} previo={p.turnosCerrados.lavado} />
          {r.lavado.tipos.map((t) => <Kpi key={t.key} label={t.label} tipo="num1" valor={t.valor} sub="pallets" />)}
        </div>
        <GraficoApiladoTipos svg={svg} titulo="Pallets lavados por día, según tipo de bandeja" data={r.lavado.diarioTipos} keys={r.lavado.tiposKeys} colores={r.lavado.colores} vacio="Sin pallets lavados en el período." />
        <PendientesLavado pend={r.lavado.pendientes} turnosSel={r.turnosSel} svg={svg} />
      </SeccionArea>

      {/* ---------- Selección: total y una sección por línea ---------- */}
      <SeccionArea titulo="Selección" color="#7c3aed" salto>
        <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Total Selección</p>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
          <Kpi label="Kg ingresados" valor={r.seleccion.ing} actual={r.seleccion.ing} previo={p.seleccion.ing} />
          <Kpi label="Kg aprobados" valor={r.seleccion.apr} actual={r.seleccion.apr} previo={p.seleccion.apr} />
          <Kpi label="Rendimiento" tipo="pct" valor={r.seleccion.rend} actual={r.seleccion.rend} previo={p.seleccion.rend} />
          <Kpi label="Cumplimiento vs. programa" tipo="pct" valor={r.seleccion.cump} tono={tonoCumplimiento(r.seleccion.cump)} actual={r.seleccion.cump} previo={p.seleccion.cump} sub={r.seleccion.prog ? `${fmtNum(r.seleccion.prog)} Kg programados` : "Sin Kg programado"} />
        </div>
        {lineasSel.length === 0 && <EmptyNote text="Sin cierres de Selección en el período." />}
        {lineasSel.map(({ l, prev }, i) => (
          <SubBloque key={l.key} titulo={l.label} color={SERIES_COLORS[i % SERIES_COLORS.length]}>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
              <Kpi label="Kg ingresados" valor={l.ing} actual={l.ing} previo={prev.ing} />
              <Kpi label="Kg aprobados" valor={l.apr} actual={l.apr} previo={prev.apr} />
              <Kpi label="Rendimiento" tipo="pct" valor={l.rend} actual={l.rend} previo={prev.rend} />
              <Kpi label="Cumplimiento vs. programa" tipo="pct" valor={l.cump} tono={tonoCumplimiento(l.cump)} actual={l.cump} previo={prev.cump} sub={l.prog ? `${fmtNum(l.prog)} Kg programados` : "Sin Kg programado"} />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <GraficoDiario svg={svg} titulo="Kg ingresados por día" data={l.diarioIng} mitad color="#a78bfa" />
              <GraficoDiario svg={svg} titulo="Rendimiento por día" data={l.diarioRend} mitad formato="pct" color="#7c3aed" />
            </div>
          </SubBloque>
        ))}
      </SeccionArea>

      {/* ---------- Envasado: Envasadora y Línea 5 por separado ---------- */}
      <SeccionArea titulo="Envasado" color="#059669" salto>
        <SubBloque titulo="Envasadora" color="#10b981">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <Kpi label="Cajas producidas" valor={r.envasado.cajasProd} actual={r.envasado.cajasProd} previo={p.envasado.cajasProd} />
            <Kpi label="Cajas programadas" valor={r.envasado.cajasProg} />
            <Kpi label="Cumplimiento" tipo="pct" valor={r.envasado.cump} tono={tonoCumplimiento(r.envasado.cump)} actual={r.envasado.cump} previo={p.envasado.cump} />
            <Kpi label="Productividad" tipo="num1" unidad="bolsas/min" valor={r.envasado.productividad} actual={r.envasado.productividad} previo={p.envasado.productividad} />
            <Kpi label="Tiempo efectivo" valor={r.envasado.tiempoEfectivo} unidad="min" actual={r.envasado.tiempoEfectivo} previo={p.envasado.tiempoEfectivo} />
            <Kpi label="Detenciones" valor={r.envasado.detTotal} unidad="min" actual={r.envasado.detTotal} previo={p.envasado.detTotal} menorEsMejor />
            <Kpi label="Merma de cajas" tipo="pct" valor={r.envasado.mermaCajas} actual={r.envasado.mermaCajas} previo={p.envasado.mermaCajas} menorEsMejor />
            <Kpi label="Merma de bolsas" tipo="pct" valor={r.envasado.mermaBolsas} actual={r.envasado.mermaBolsas} previo={p.envasado.mermaBolsas} menorEsMejor />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <GraficoDiario svg={svg} titulo="Cajas producidas por día" mitad data={r.envasado.diarioCajas} color="#34d399" />
            <GraficoDiario svg={svg} titulo="Cumplimiento por día" mitad data={r.envasado.diarioCump} formato="pct" color="#10b981" />
          </div>
          <div className="rep-bloque">
            <p className="text-xs font-semibold text-slate-700 mb-2">Detenciones por motivo (Pareto)</p>
            <ParetoDetenciones items={r.envasado.detPareto} total={r.envasado.detTotal} />
          </div>
        </SubBloque>

        <SubBloque titulo="Línea 5" color="#0d9488">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <Kpi label="Kg ingresados" valor={r.envasado.l5Ing} actual={r.envasado.l5Ing} previo={p.envasado.l5Ing} />
            <Kpi label="Kg aprobados" valor={r.envasado.l5Apr} actual={r.envasado.l5Apr} previo={p.envasado.l5Apr} />
            <Kpi label="Rendimiento" tipo="pct" valor={r.envasado.l5Rend} actual={r.envasado.l5Rend} previo={p.envasado.l5Rend} />
            <Kpi label="Cumplimiento vs. programa" tipo="pct" valor={r.envasado.l5Cump} tono={tonoCumplimiento(r.envasado.l5Cump)} actual={r.envasado.l5Cump} previo={p.envasado.l5Cump} sub={r.envasado.l5Prog ? `${fmtNum(r.envasado.l5Prog)} Kg programados` : "Sin Kg programado"} />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <GraficoDiario svg={svg} titulo="Kg ingresados por día" data={r.envasado.l5DiarioIng} mitad color="#2dd4bf" />
            <GraficoDiario svg={svg} titulo="Rendimiento por día" data={r.envasado.l5DiarioRend} mitad formato="pct" color="#0d9488" />
          </div>
        </SubBloque>
      </SeccionArea>

      {/* ---------- Seguridad ---------- */}
      <SeccionArea titulo="Seguridad" color="#dc2626" salto>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
          <Kpi label="Cierres con incidentes" valor={r.seguridad.incidentes} tono={r.seguridad.incidentes ? "warn" : "ok"} actual={r.seguridad.incidentes} previo={p.seguridad.incidentes} menorEsMejor />
          <Kpi label="Cierres con accidentes" valor={r.seguridad.accidentes} tono={r.seguridad.accidentes ? "bad" : "ok"} actual={r.seguridad.accidentes} previo={p.seguridad.accidentes} menorEsMejor />
          <Kpi label="Documentación pendiente" valor={r.seguridad.docsPendientes} tono={r.seguridad.docsPendientes ? "bad" : "ok"} />
          <Kpi label="Turnos sin cierre" valor={r.turnosSinCierreTotal} tono={r.turnosSinCierreTotal ? "warn" : "ok"} />
        </div>

        <SubBloque titulo="Incidentes y accidentes por supervisor" color="#dc2626">
          <p className="text-[11px] text-slate-500 -mt-1">Cada barra suma los casos de incidente y accidente; el color indica al supervisor responsable del cierre.</p>
          <GraficoApilado titulo="Por día (período seleccionado)" data={sg.porDia} keys={sg.supervisores} colores={sg.colores} hay={sg.hayDia} />
          <GraficoApilado titulo="Acumulado por semana (últimas 8 semanas, desde el lunes indicado)" data={sg.porSemana} keys={sg.supervisores} colores={sg.colores} hay={sg.haySemana} />
          <GraficoApilado titulo="Acumulado por mes (últimos 6 meses)" data={sg.porMes} keys={sg.supervisores} colores={sg.colores} hay={sg.hayMes} />
          {sg.tabla.length > 0 && (
            <div className="rep-bloque overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="text-left text-slate-500 border-b border-slate-200">
                    <th className="py-1.5 font-medium">Supervisor</th>
                    <th className="py-1.5 font-medium text-right">Período · incidentes</th>
                    <th className="py-1.5 font-medium text-right">Período · accidentes</th>
                    <th className="py-1.5 font-medium text-right">{sg.mesLabel} · incidentes</th>
                    <th className="py-1.5 font-medium text-right">{sg.mesLabel} · accidentes</th>
                  </tr>
                </thead>
                <tbody>
                  {sg.tabla.map((t) => (
                    <tr key={t.sup} className="border-b border-slate-100 last:border-0">
                      <td className="py-1.5 font-medium text-slate-900">
                        <span className="inline-block w-2.5 h-2.5 rounded-sm mr-1.5 align-middle" style={{ background: sg.colores[t.sup] }} />{t.sup}
                      </td>
                      <td className="py-1.5 text-right">{t.perInc}</td>
                      <td className={`py-1.5 text-right ${t.perAcc ? "text-red-600 font-semibold" : ""}`}>{t.perAcc}</td>
                      <td className="py-1.5 text-right">{t.mesInc}</td>
                      <td className={`py-1.5 text-right ${t.mesAcc ? "text-red-600 font-semibold" : ""}`}>{t.mesAcc}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </SubBloque>

        <div>
          <p className="rep-titulo text-xs font-semibold text-slate-700 mb-2">Reporte de accidentes e incidentes ({r.seguridad.eventos.length})</p>
          <ListaEventos eventos={r.seguridad.eventos} />
        </div>
      </SeccionArea>
    </div>
  );
}

function JefeDashboard({ onBack }) {
  const lavado = useAreaCierres("lavado");
  const seleccion = useAreaCierres("seleccion");
  const envasado = useAreaCierres("envasado");
  const [programas, , loadingP] = useSharedList("programa-records");
  const { horarios } = useHorariosTurno();
  const [modo, setModo] = useState("semana");
  const [offset, setOffset] = useState(0);
  const [verReporte, setVerReporte] = useState(false);
  const [turnos, setTurnos] = useState(TODOS_LOS_TURNOS);

  const loading = lavado.loading || seleccion.loading || envasado.loading || loadingP;
  const semana = semanaDe(sumarDias(today(), offset * 7));
  const dias = modo === "semana" ? 7 : 28;
  const hasta = semana.hasta;
  const desde = sumarDias(hasta, -(dias - 1));
  const datos = filtrarDatosPorTurno({ lavado, seleccion, envasado, programas, horarios }, turnos);
  const r = loading ? null : resumenPeriodo(datos, desde, hasta, turnos);
  const p = loading ? null : resumenPeriodo(datos, sumarDias(desde, -dias), sumarDias(desde, -1), turnos);

  // El reporte siempre es de UNA semana (lunes a domingo): la semana seleccionada.
  const rSemana = loading ? null : (dias === 7 ? r : resumenPeriodo(datos, semana.desde, semana.hasta, turnos));
  const pSemana = loading ? null : resumenPeriodo(datos, sumarDias(semana.desde, -7), sumarDias(semana.desde, -1), turnos);

  return (
    <div className="w-full max-w-2xl lg:max-w-5xl mx-auto pb-10">
      <TopBar title="Indicadores generales" subtitle="Vista ejecutiva de las 3 áreas" onBack={onBack} icon={BarChart3} accent="emerald" />
      <div className="p-4 space-y-4">
        <SelectorPeriodo modo={modo} setModo={setModo} offset={offset} setOffset={setOffset} desde={desde} hasta={hasta} />
        <SelectorTurnos turnos={turnos} setTurnos={setTurnos} />

        {loading ? <Card><Loader text="Calculando indicadores…" /></Card> : (
          <>
            <button
              onClick={() => setVerReporte(true)}
              className="w-full flex items-center justify-center gap-2 bg-slate-900 hover:bg-slate-800 text-white font-semibold rounded-xl py-3 transition-colors"
            >
              <FileText size={18} /> Generar reporte semanal (PDF)
              <span className="text-xs font-normal text-slate-300">· {fmtCorto(semana.desde)} al {fmtCorto(semana.hasta)}</span>
            </button>

            <PanelIndicadores r={r} p={p} />

            <p className="text-xs text-slate-500">
              Las variaciones comparan con el período anterior de igual duración. El color del cumplimiento es verde desde 95 %, ámbar desde 85 % y rojo bajo eso. Para editar o eliminar registros, entra al área y abre la pestaña «Registros».
            </p>
          </>
        )}
      </div>

      {verReporte && rSemana && (
        <ReporteSemanalModal r={rSemana} p={pSemana} onClose={() => setVerReporte(false)} />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Reporte semanal (se abre en pantalla y se guarda como PDF con «Imprimir»).
// Es el mismo panel de la pantalla de Indicadores, con encabezado y comentarios.
// ---------------------------------------------------------------------------
function ReporteSemanalModal({ r, p, onClose }) {
  const ahora = fmtDateTime(new Date().toISOString());
  const H2 = "text-[13px] font-bold text-slate-900 border-b-2 border-slate-800 pb-1 mb-2 mt-5";

  return createPortal(
    <div className="reporte-portal fixed inset-0 z-[100] bg-slate-200 overflow-auto">
      <style>{`
        @media print {
          #root { display: none !important; }
          .reporte-portal { position: static !important; overflow: visible !important; background: #fff !important; }
          .no-print { display: none !important; }
          #reporte-print { box-shadow: none !important; margin: 0 !important; max-width: none !important; padding: 0 !important; }
          #reporte-print, #reporte-print * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          .rep-bloque { break-inside: avoid; }
          .rep-salto { break-before: page; }
          .rep-titulo { break-after: avoid; }
          #reporte-print { background: #fff !important; }
          @page { size: A4; margin: 10mm; }
        }
      `}</style>

      <div className="no-print sticky top-0 z-10 flex items-center justify-between gap-2 bg-slate-900 text-white px-4 py-2.5">
        <button onClick={onClose} className="flex items-center gap-1.5 text-sm hover:underline"><X size={16} /> Cerrar</button>
        <button onClick={() => window.print()} className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-500 font-semibold text-sm rounded-lg px-4 py-2">
          <Printer size={16} /> Imprimir / Guardar como PDF
        </button>
      </div>
      <p className="no-print text-center text-xs text-slate-600 px-4 pt-3">
        En el cuadro de impresión elige «Guardar como PDF» como destino y luego compártelo.
      </p>

      <div id="reporte-print" className="bg-slate-50 text-slate-900 max-w-[820px] mx-auto my-3 p-5 shadow text-[11.5px] leading-snug">
        <div className="rep-bloque flex items-start justify-between rounded-2xl bg-slate-900 text-white px-5 py-4 mb-4">
          <div>
            <div className="text-lg font-extrabold">Reporte semanal de producción</div>
            <div className="text-slate-300">IQF Congelado · Planta Curicó</div>
          </div>
          <div className="text-right">
            <div className="font-bold">Semana {fmtFecha(r.desde)} – {fmtFecha(r.hasta)}</div>
            <div className="text-slate-300">{r.turnosSel.length === TODOS_LOS_TURNOS.length ? "Todos los turnos" : `Turnos: ${r.turnosSel.map(turnoLabel).join(", ")}`}</div>
            <div className="text-slate-300">Generado el {ahora}</div>
          </div>
        </div>

        <PanelIndicadores r={r} p={p} svg />

        {r.comentarios.length > 0 && (
          <div className="rep-bloque">
            <div className={H2}>Comentarios de los turnos</div>
            <table className="w-full">
              <tbody>
                {r.comentarios.slice(0, 12).map((c, i) => (
                  <tr key={i} className="border-b border-slate-200 align-top">
                    <td className="py-1 px-2 whitespace-nowrap text-slate-500">{fmtCorto(c.fecha)} {c.turno}</td>
                    <td className="py-1 px-2 whitespace-nowrap">{c.area} · {c.tema}</td>
                    <td className="py-1 px-2">{c.texto.length > 220 ? `${c.texto.slice(0, 220)}…` : c.texto}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {r.comentarios.length > 12 && <div className="text-[10px] text-slate-500 mt-1">Se muestran los 12 comentarios más recientes de {r.comentarios.length}.</div>}
          </div>
        )}

        <div className="text-[10px] text-slate-500 mt-4">
          Las variaciones comparan con la semana anterior (puntos porcentuales para porcentajes, % para el resto). Cumplimiento: verde desde 95 %, ámbar desde 85 %, rojo bajo eso.
        </div>
      </div>
    </div>,
    document.body
  );
}

// ---------------------------------------------------------------------------
// Acceso Jefe de producción
// ---------------------------------------------------------------------------
function JefeLoginModal({ onClose, onSuccess }) {
  const [pwd, setPwd] = useState("");
  const [error, setError] = useState(false);

  const submit = () => {
    if (pwd === JEFE_PASSWORD) onSuccess();
    else setError(true);
  };

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4">
      <div className="bg-white border border-slate-200 shadow-xl rounded-xl p-5 w-full max-w-sm">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold text-slate-900 flex items-center gap-2">
            <Lock size={16} /> Acceso jefe de producción
          </h3>
          <button onClick={onClose} aria-label="Cerrar"><X size={18} className="text-slate-600" /></button>
        </div>
        <label className="block text-xs font-medium text-slate-600 mb-1">Clave de acceso</label>
        <input
          type="password"
          autoFocus
          className={inputBase}
          value={pwd}
          onChange={(e) => { setPwd(e.target.value); setError(false); }}
          onKeyDown={(e) => e.key === "Enter" && submit()}
        />
        {error && <p className="text-xs text-red-400 mt-2">Clave incorrecta.</p>}
        <button onClick={submit} className="mt-3 w-full bg-blue-600 hover:bg-blue-500 text-white text-sm font-semibold rounded-lg py-2">
          Ingresar
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pantalla de inicio
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// LAYOUT DE ESCRITORIO — AppShell
// En pantallas pequeñas (móvil) se muestra el contenido directamente.
// En pantallas grandes (lg+) aparece una barra lateral izquierda fija con
// la navegación del portal actual, y el contenido principal a la derecha.
// ---------------------------------------------------------------------------
const ACCENT_SIDEBAR = {
  blue:   "bg-blue-700 text-white",
  rose:   "bg-rose-700 text-white",
  emerald:"bg-emerald-700 text-white",
  amber:  "bg-amber-600 text-white",
  slate:  "bg-slate-800 text-white",
  violet: "bg-violet-700 text-white",
};

function DesktopSidebar({ portal, onNavigate, isJefe, autor }) {
  if (!portal) return null;
  const { label, accent, items } = portal;
  const bg = ACCENT_SIDEBAR[accent] || ACCENT_SIDEBAR.slate;
  return (
    <aside className={`hidden lg:flex flex-col w-56 shrink-0 min-h-screen ${bg}`}>
      <div className="px-5 pt-6 pb-4 border-b border-white/20">
        <div className="text-xs font-bold uppercase tracking-widest opacity-60 mb-1">Portal</div>
        <div className="text-lg font-extrabold leading-tight">{label}</div>
        {autor && <div className="text-xs opacity-70 mt-1 truncate">{autor}</div>}
      </div>
      <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
        {items.map(({ key, label: lbl, Icon }) => (
          <button key={key} onClick={() => onNavigate(key)}
            className="w-full flex items-center gap-2.5 px-3 py-2.5 rounded-xl text-sm font-medium hover:bg-white/20 transition-colors text-left">
            {Icon && <Icon size={16} className="shrink-0 opacity-80" />}
            <span>{lbl}</span>
          </button>
        ))}
      </nav>
      <div className="px-5 py-4 border-t border-white/20">
        <div className="text-xs opacity-50">Bitácora de Turnos</div>
      </div>
    </aside>
  );
}

function AppShell({ children, portal, onNavigate, isJefe, autor }) {
  return (
    <div className="flex min-h-screen bg-slate-100">
      <DesktopSidebar portal={portal} onNavigate={onNavigate} isJefe={isJefe} autor={autor} />
      <main className="flex-1 min-w-0 bg-slate-50 overflow-y-auto">
        {children}
      </main>
    </div>
  );
}

function PortalHeader({ onOpenLogin, onLogoutJefe, isJefe }) {
  return (
    <div className="sticky top-0 z-10 bg-white border-b border-slate-200 px-4 py-3 flex items-center justify-between">
      <div className="flex items-center gap-2">
        <div className="w-9 h-9 rounded-lg bg-blue-600 flex items-center justify-center text-white font-bold text-sm">BT</div>
        <div>
          <div className="text-sm font-semibold text-slate-900 leading-tight">Bitácora de Turnos</div>
          <div className="text-xs text-slate-600 leading-tight">Planta de procesos</div>
        </div>
      </div>
      {isJefe ? (
        <button onClick={onLogoutJefe} className="flex items-center gap-1.5 text-xs font-medium text-emerald-700 bg-emerald-500/10 border border-emerald-500/30 rounded-full px-3 py-1.5">
          <LogOut size={14} /> Jefe de producción
        </button>
      ) : (
        <button onClick={onOpenLogin} aria-label="Acceso jefe de producción" className="rounded-lg p-2 hover:bg-slate-200 text-slate-600">
          <Settings size={18} />
        </button>
      )}
    </div>
  );
}

// Botón de sección grande (usado en HomeScreen y portales de área)
function SectionButton({ label, desc, icon: Icon, accent, onClick, className = "" }) {
  return (
    <button onClick={onClick}
      className={`bg-white border border-slate-200 rounded-2xl p-5 text-left hover:shadow-md hover:border-slate-300 transition-all flex items-center gap-4 ${className}`}>
      <div className={`w-14 h-14 rounded-xl flex items-center justify-center border shrink-0 ${ACCENT[accent]?.badge || ACCENT.blue.badge}`}>
        <Icon size={26} />
      </div>
      <div className="flex-1 min-w-0">
        <div className="font-bold text-slate-900 text-base">{label}</div>
        <div className="text-xs text-slate-500 mt-0.5">{desc}</div>
      </div>
      <ChevronRight size={20} className="text-slate-400 shrink-0" />
    </button>
  );
}

// ── Portal Selección ─────────────────────────────────────────────────────────
function SeleccionPortal({ onNavigate, onBack, autor, setAutor, isJefe, onOpenLogin, onLogoutJefe, supervisoresList }) {
  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-8">
      <PortalHeader onOpenLogin={onOpenLogin} onLogoutJefe={onLogoutJefe} isJefe={isJefe} />
      <div className="px-4 pt-6 pb-2">
        <button onClick={onBack} className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-800 mb-4">
          <ArrowLeft size={14} /> Inicio
        </button>
        <h1 className="text-2xl font-extrabold text-slate-900">Selección</h1>
        <p className="text-sm text-slate-500 mt-0.5">Lavado de bandejas · Líneas de selección · Insumos</p>
      </div>

      <div className="px-4 mb-4">
        <Card title="Supervisor">
          <div className="flex items-center gap-2">
            <User size={18} className="text-slate-500 shrink-0" />
            <div className="flex-1 min-w-0">
              <SupervisorSelect area="Seleccion" value={autor} onChange={setAutor} supervisoresList={supervisoresList} />
            </div>
          </div>
        </Card>
      </div>

      <div className="px-4 space-y-2.5">
        <SectionButton label="Lavado de bandejas" desc="Inicio, cierre e indicadores del área de lavado" icon={AREAS.lavado.icon} accent={AREAS.lavado.accent} onClick={() => onNavigate("lavado")} />
        <SectionButton label="Selección" desc="Inicio, cierre e indicadores de líneas de selección" icon={AREAS.seleccion.icon} accent={AREAS.seleccion.accent} onClick={() => onNavigate("seleccion")} />
        <SectionButton label="Insumos y Consumo" desc="Pallets, film, bolsas — necesito vs. tengo en piso" icon={Boxes} accent="amber" onClick={() => onNavigate("insumos-seleccion")} />
      </div>
    </div>
  );
}

// ── Portal Envasado ──────────────────────────────────────────────────────────
function EnvasadoPortal({ onNavigate, onBack, autor, setAutor, isJefe, onOpenLogin, onLogoutJefe, supervisoresList }) {
  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-8">
      <PortalHeader onOpenLogin={onOpenLogin} onLogoutJefe={onLogoutJefe} isJefe={isJefe} />
      <div className="px-4 pt-6 pb-2">
        <button onClick={onBack} className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-800 mb-4">
          <ArrowLeft size={14} /> Inicio
        </button>
        <h1 className="text-2xl font-extrabold text-slate-900">Envasado</h1>
        <p className="text-sm text-slate-500 mt-0.5">Envasadora · Línea 5 · Verificador de hora · Insumos</p>
      </div>

      <div className="px-4 mb-4">
        <Card title="Supervisor">
          <div className="flex items-center gap-2">
            <User size={18} className="text-slate-500 shrink-0" />
            <div className="flex-1 min-w-0">
              <SupervisorSelect area="Envasado" value={autor} onChange={setAutor} supervisoresList={supervisoresList} />
            </div>
          </div>
        </Card>
      </div>

      <div className="px-4 space-y-2.5">
        <SectionButton label="Envasado" desc="Inicio, cierre e indicadores de la línea de envasado" icon={AREAS.envasado.icon} accent={AREAS.envasado.accent} onClick={() => onNavigate("envasado")} />
        <SectionButton label="Verificador de hora" desc="Código de bloque horario para codificación de bolsa/caja" icon={Clock} accent="amber" onClick={() => onNavigate("verificador")} />
        <SectionButton label="Insumos y Consumo" desc="Cajas, bolsas, film — necesito vs. tengo en piso" icon={Boxes} accent="amber" onClick={() => onNavigate("insumos-envasado")} />
        <SectionButton label="Especificación de SKU" desc="Materiales, codificaciones y config. de pallet por SKU" icon={Package} accent="amber" onClick={() => onNavigate("especificacion")} />
      </div>
    </div>
  );
}

// ── Home principal ───────────────────────────────────────────────────────────
function HomeScreen({ onNavigate, isJefe, onOpenLogin, onLogoutJefe, autor }) {
  const alertasInsumos = useAlertasInsumos();

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-8">
      <PortalHeader onOpenLogin={onOpenLogin} onLogoutJefe={onLogoutJefe} isJefe={isJefe} />

      <div className="px-4 pt-8 pb-4">
        <div className="flex items-center gap-2 text-xs uppercase tracking-widest text-slate-500 mb-3">
          <span className="w-4 h-px bg-slate-400" /> Sistema de registro de turnos
        </div>
        <h1 className="text-4xl font-extrabold text-slate-900 leading-tight">
          Bitácora<br /><span className="text-blue-500">de Turnos</span>
        </h1>
        <p className="text-sm text-slate-500 mt-2">Selecciona el área de producción para comenzar.</p>
      </div>

      {alertasInsumos.length > 0 && (() => {
        // Agrupa por área para que el aviso especifique claramente a cuál
        // sector pertenece cada quiebre/casi-quiebre.
        const porArea = {};
        alertasInsumos.forEach((a) => {
          (porArea[a.area] = porArea[a.area] || []).push(a);
        });
        const areas = Object.keys(porArea);
        return (
          <div className="px-4 mb-4 space-y-2">
            {areas.map((areaNombre) => (
              <button
                key={areaNombre}
                onClick={() => onNavigate(areaNombre === "Envasado" ? "insumos-envasado" : "insumos-seleccion")}
                className="w-full text-left bg-red-50 border border-red-300 rounded-2xl px-4 py-3 flex items-start gap-3 hover:bg-red-100 transition-colors"
              >
                <AlertTriangle size={20} className="text-red-600 shrink-0 mt-0.5" />
                <div className="flex-1 min-w-0">
                  <div className="text-sm font-bold text-red-800">
                    {areaNombre}: {porArea[areaNombre].length} insumo{porArea[areaNombre].length > 1 ? "s" : ""} no va{porArea[areaNombre].length > 1 ? "n" : ""} a alcanzar para hoy/mañana
                  </div>
                  <div className="text-xs text-red-700 mt-0.5 truncate">
                    {porArea[areaNombre].slice(0, 3).map((a) => a.nombre).join(" · ")}
                    {porArea[areaNombre].length > 3 ? ` · +${porArea[areaNombre].length - 3} más` : ""}
                  </div>
                  <div className="text-xs text-red-600 mt-1 underline">Ir a Insumos de {areaNombre} a registrar el pedido →</div>
                </div>
              </button>
            ))}
          </div>
        );
      })()}

      <div className="px-4 space-y-3">
        {/* Dos botones grandes de área */}
        <button onClick={() => onNavigate("portal-seleccion")}
          className="w-full bg-gradient-to-br from-blue-600 to-blue-700 hover:from-blue-500 hover:to-blue-600 text-white rounded-2xl p-6 text-left shadow-lg transition-all flex items-center gap-5">
          <div className="w-16 h-16 rounded-2xl bg-white/20 flex items-center justify-center shrink-0">
            <Filter size={30} className="text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <div className="font-extrabold text-xl leading-tight">Selección</div>
            <div className="text-blue-200 text-sm mt-1">Lavado de bandejas · Líneas L1, L3, L4 · Insumos</div>
          </div>
          <ChevronRight size={24} className="text-white/70 shrink-0" />
        </button>

        <button onClick={() => onNavigate("portal-envasado")}
          className="w-full bg-gradient-to-br from-amber-500 to-orange-600 hover:from-amber-400 hover:to-orange-500 text-white rounded-2xl p-6 text-left shadow-lg transition-all flex items-center gap-5">
          <div className="w-16 h-16 rounded-2xl bg-white/20 flex items-center justify-center shrink-0">
            <Package size={30} className="text-white" />
          </div>
          <div className="flex-1 min-w-0">
            <div className="font-extrabold text-xl leading-tight">Envasado</div>
            <div className="text-amber-100 text-sm mt-1">Envasadora · Línea 5 · Verificador de hora · Insumos</div>
          </div>
          <ChevronRight size={24} className="text-white/70 shrink-0" />
        </button>

        {/* Accesos Jefe de producción */}
        {isJefe && (
          <div className="pt-2 space-y-2">
            <p className="text-xs font-semibold text-slate-500 uppercase tracking-widest px-1">Jefe de producción</p>
            {[
              { key: "dashboard",        label: "Indicadores generales",       desc: "Vista ejecutiva de las 3 áreas y reporte semanal en PDF", Icon: BarChart3 },
              { key: "programa",         label: "Programa de producción",      desc: "Línea, turno, especie y cantidad", Icon: CalendarDays },
              { key: "supervisores",     label: "Gestionar Supervisores",      desc: "Agregar o eliminar supervisores de los portales", Icon: User },
              { key: "horarios",         label: "Horarios de Turno",           desc: "Lunes a Sábado — minutos efectivos por turno", Icon: Clock },
              { key: "especies",         label: "Gestionar especies por línea",desc: "Agregar nuevos procesos/especies", Icon: Plus },
              { key: "insumos-config",   label: "Configurar Insumos",          desc: "Tasas de consumo de Selección, Lavado y Envasado", Icon: Boxes },
              { key: "calculadora-costos", label: "Calculadora de Costos",     desc: "Estima el gasto en insumos para un escenario de Kg/turnos", Icon: CheckCircle2 },
              { key: "espec-jefe",       label: "Editar Especificación SKU",   desc: "Modifica materiales, codificaciones y config. de pallet", Icon: Settings },
            ].map(({ key, label, desc, Icon }) => (
              <button key={key} onClick={() => onNavigate(key)}
                className="w-full bg-white border border-slate-200 rounded-xl p-3.5 text-left hover:border-slate-400 transition-colors flex items-center gap-3">
                <div className={`w-10 h-10 rounded-lg flex items-center justify-center border ${ACCENT.emerald.badge} shrink-0`}>
                  <Icon size={18} />
                </div>
                <div className="flex-1 min-w-0">
                  <div className="font-semibold text-slate-900 text-sm">{label}</div>
                  <div className="text-xs text-slate-500">{desc}</div>
                </div>
                <ChevronRight size={16} className="text-slate-400 shrink-0" />
              </button>
            ))}
          </div>
        )}
      </div>

      <p className="text-xs text-slate-500 text-center px-4 pt-6 pb-4">
        Los datos se guardan de forma compartida para todas las personas que usan esta app.
      </p>
    </div>
  );
}


// ---------------------------------------------------------------------------
// App raíz
// ---------------------------------------------------------------------------
export default function App() {
  const [screen, setScreen] = useState("home");
  const [autor, setAutor, autorLoaded] = usePersonalValue("autor-nombre", "");
  const [supervisoresList]             = useSupervisores();
  const [jefeFlag, setJefeFlag] = usePersonalValue("modo-jefe", "");
  const [showLogin, setShowLogin] = useState(false);
  const isJefe = jefeFlag === "true";

  // SKU nuevos agregados por el Jefe en la pestaña "SKU Envasado" del Google
  // Sheet (si hay conexión configurada). Se actualiza el registro a nivel de
  // módulo ANTES de que rendericen los hijos, para que skuMaterial()/SkuPicker
  // ya los vean disponibles en este mismo render.
  const [skuNuevos] = useSharedList("sku-nuevos");
  SKU_NUEVOS = skuNuevos || [];

  // A qué pantalla volver según dónde estamos
  const backTarget = (s) => {
    if (["lavado", "seleccion", "insumos-seleccion"].includes(s)) return "portal-seleccion";
    if (["envasado", "verificador", "insumos-envasado", "especificacion"].includes(s)) return "portal-envasado";
    return "home";
  };

  // Props compartidas entre los dos portales de área
  const portalProps = {
    onBack: () => setScreen("home"),
    autor: autorLoaded ? autor : "",
    setAutor,
    isJefe,
    onOpenLogin: () => setShowLogin(true),
    onLogoutJefe: () => setJefeFlag(""),
    supervisoresList: supervisoresList || [],
  };

  // Define qué portal/sidebar mostrar según la pantalla activa
  const PORTALS = {
    seleccion: {
      label: "Selección", accent: "emerald",
      items: [
        { key: "lavado",            label: "Lavado de bandejas",    Icon: Droplets },
        { key: "seleccion",         label: "Selección",             Icon: Filter },
        { key: "insumos-seleccion", label: "Insumos y Consumo",     Icon: Boxes },
      ],
    },
    envasado: {
      label: "Envasado", accent: "amber",
      items: [
        { key: "envasado",          label: "Envasado",              Icon: Package },
        { key: "verificador",       label: "Verificador de hora",   Icon: Clock },
        { key: "insumos-envasado",  label: "Insumos y Consumo",     Icon: Boxes },
        { key: "especificacion",    label: "Especificación de SKU", Icon: Package },
      ],
    },
  };
  const activePortion =
    ["lavado","seleccion","insumos-seleccion"].includes(screen) ? "seleccion" :
    ["envasado","verificador","insumos-envasado","especificacion"].includes(screen) ? "envasado" :
    null;
  const activePortal = activePortion ? PORTALS[activePortion] : null;

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900 font-sans">
      <AppShell portal={activePortal} onNavigate={setScreen} isJefe={isJefe} autor={portalProps.autor}>
      {/* Home principal — dos botones grandes */}
      {screen === "home" && (
        <HomeScreen
          onNavigate={setScreen}
          isJefe={isJefe}
          onOpenLogin={() => setShowLogin(true)}
          onLogoutJefe={() => setJefeFlag("")}
          autor={portalProps.autor}
        />
      )}

      {/* Portales de área */}
      {screen === "portal-seleccion" && (
        <SeleccionPortal {...portalProps} onNavigate={setScreen} />
      )}
      {screen === "portal-envasado" && (
        <EnvasadoPortal {...portalProps} onNavigate={setScreen} />
      )}

      {/* Pantallas de área — back vuelve al portal correspondiente */}
      {AREAS[screen] && (
        <AreaScreen
          areaKey={screen}
          isJefe={isJefe}
          autor={autor}
          onBack={() => setScreen(backTarget(screen))}
          supervisoresList={supervisoresList || []}
        />
      )}

      {/* Insumos filtrados por área */}
      {screen === "insumos-seleccion" && (
        <InsumosConsumoScreen isJefe={isJefe} areaFiltro="seleccion" onBack={() => setScreen("portal-seleccion")} autor={portalProps.autor} />
      )}
      {screen === "insumos-envasado" && (
        <InsumosConsumoScreen isJefe={isJefe} areaFiltro="envasado" onBack={() => setScreen("portal-envasado")} autor={portalProps.autor} />
      )}

      {/* Verificador — accesible desde el portal Envasado */}
      {screen === "verificador" && (
        <VerificadorHoraScreen onBack={() => setScreen("portal-envasado")} />
      )}

      {/* Herramientas del Jefe — back siempre al home */}
      {screen === "dashboard" && isJefe && <JefeDashboard onBack={() => setScreen("home")} />}
      {screen === "programa" && isJefe && <ProgramaProduccionScreen onBack={() => setScreen("home")} />}
      {screen === "especies" && isJefe && <EspeciesPorLineaScreen onBack={() => setScreen("home")} />}
      {screen === "insumos-config"  && isJefe && <InsumosConfigScreen onBack={() => setScreen("home")} />}
      {screen === "calculadora-costos" && isJefe && <CalculadoraCostosScreen onBack={() => setScreen("home")} />}
      {screen === "espec-jefe"      && isJefe && <EspecificacionEditScreen onBack={() => setScreen("home")} />}
      {screen === "supervisores"    && isJefe && <SupervisoresScreen onBack={() => setScreen("home")} />}
      {screen === "horarios"        && isJefe && <HorariosScreen onBack={() => setScreen("home")} />}
      {screen === "especificacion"             && <EspecificacionScreen onBack={() => setScreen("portal-envasado")} />}

      {/* Modal de login jefe */}
      {showLogin && (
        <JefeLoginModal onClose={() => setShowLogin(false)} onSuccess={() => { setJefeFlag("true"); setShowLogin(false); }} />
      )}
      </AppShell>
    </div>
  );
}
