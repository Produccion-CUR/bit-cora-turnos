import React, { useState, useEffect, useMemo } from "react";
import {
  Droplets, Filter, Package, ClipboardCheck, ClipboardList, BarChart3, Settings, X, Lock,
  Save, ArrowLeft, User, Loader2, CheckCircle2, AlertTriangle, Sun, Moon, Clock, Pencil,
  Trash2, LogOut, CalendarDays, Share2, Boxes, Plus, ChevronRight, BookOpen, ChevronDown,
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
  while (s && !turnosActivosKeys.includes(s.turno)) {
    s = turnoSiguiente(s.fecha, s.turno);
  }
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
function useSharedList(key) {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        const res = await window.storage.get(key, true);
        if (active) setItems(res ? JSON.parse(res.value) : []);
      } catch {
        if (active) setItems([]);
      }
      if (active) setLoading(false);
    })();
    return () => { active = false; };
  }, [key]);

  // Devuelve true/false según si realmente se guardó en el almacenamiento persistente.
  // El estado local (items) se actualiza igual, para que la pantalla refleje el cambio
  // aunque la sincronización falle (y el usuario pueda reintentar sin perder lo escrito).
  const save = async (newItems) => {
    setItems(newItems);

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
      setError(null);
      return true;
    } catch (e) {
      setError(`Error al guardar (${e?.message || "desconocido"}). El cambio quedó en esta sesión pero no se sincronizó: vuelve a presionar Guardar.`);
      return false;
    }
  };

  return [items, save, loading, error];
}

function usePersonalValue(key, fallback) {
  const [value, setValue] = useState(fallback);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        const res = await window.storage.get(key, false);
        if (active && res) setValue(res.value);
      } catch {
        // sin valor guardado todavía
      }
      if (active) setLoaded(true);
    })();
    return () => { active = false; };
  }, [key]);

  const update = async (v) => {
    setValue(v);
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
  const [modoLibre, setModoLibre] = useState(!estaEnLista && !!value && filtrados.length > 0);

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

  if (modoLibre) {
    return (
      <div className="space-y-1">
        <input
          className={inputBase}
          placeholder="Escribe el nombre del supervisor"
          value={value}
          onChange={(e) => onChange(e.target.value)}
        />
        <button type="button" onClick={() => { setModoLibre(false); onChange(""); }}
          className="text-xs text-blue-600 underline">
          ← Volver al listado de supervisores
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-start gap-2">
      <div className="flex-1">
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
      </div>
      <button type="button" onClick={() => { setModoLibre(true); }}
        className="shrink-0 text-xs font-medium border border-slate-300 rounded-xl px-3 py-2.5 text-slate-600 hover:border-blue-400 hover:text-blue-600 transition-colors">
        ✏ Otro
      </button>
    </div>
  );
}

// Trae al "Inicio de turno" los valores de cierre del turno anterior
// (p. ej. pallets pendientes -> pallets sucios, materiales piso planta fin -> inicio).
// Solo completa campos vacíos y solo si no se está editando un registro existente.
function useCarryOver(areaKey, values, setField, editingId, mapping) {
  const [cierres] = useSharedList(`${areaKey}-cierre-records`);
  const prev = turnoAnterior(values.fecha, values.turno);
  const prevClave = prev ? claveTurno(prev.fecha, prev.turno) : null;
  const prevCierre = prevClave ? cierres.find((c) => c.claveTurno === prevClave) : null;

  useEffect(() => {
    if (editingId || !prevCierre) return;
    const allEmpty = mapping.every(([, target]) => values[target] === undefined || values[target] === "" || values[target] === null);
    if (!allEmpty) return;
    mapping.forEach(([source, target]) => {
      const v = prevCierre[source];
      if (v !== undefined && v !== "") setField(target, v);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [values.fecha, values.turno, prevCierre?.id, editingId]);

  return { prevCierre, prev };
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
];

const MATERIALES_ENVASADO = ["Pallet Certificado", "Film Máquina", "Film Manual", "MTC Cajas", "MTC Bolsas"];

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

// Maestro real de SKU de Envasadora. Fuente: Programa_de_Envasado_Curico_V7.xlsx
// (hojas 'Materiales' + 'Cajas x Pallets'). Campos en blanco para los SKU que no
// traían ese dato en el Excel original.
const SKU_MATERIALES = [
  { sku: "T2-02-0251", producto: "CROFTERS RASP CRUMBLE ORG 1*30LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-02-0401", producto: "NATURES PROMISE ORG RASPBERRIES 8*10 OZ", codBolsa: "MTC1686", nomBolsa: "BOL NATURES PROMISE ORG RASP 10OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-02-2280", producto: "GVO RASPBERRIES 5*10 OZ", codBolsa: "MTC1286", nomBolsa: "BOL DOYPACK GVO ORG. RASPBERRY 10OZ", codCaja: "MTC1338", nomCaja: "CJ  GVO RASPBERRY    5X10OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 253, kgXCaja: 1.41747615625, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T2-02-3143", producto: "FRA IQF ORG WWS (MACRO) 9*450GRS.", codBolsa: "MTC1136", nomBolsa: "BOL MACRO ORG RASPBE. 450G FLAT-BOTTOM", codCaja: "MTC1387", nomCaja: "CJ MACRO ORG. RASPBERRIES 9X450G KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-02-3145", producto: "GREENWISE ORG RASPBERRIES 8*10 OZ", codBolsa: "MTC1634", nomBolsa: "BOL DPACK GREENWISE RASPBERRIES 10OZ V24", codCaja: "MTC1163", nomCaja: "CJ GREENWISE ORG RASPBERRIES 8X10OZ V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.72155422, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-02-3146", producto: "WWS MACRO RASPBERRIES 9*450 GR M.O", codBolsa: "MTC1612", nomBolsa: "BOL MACRO ORG RASPBE. 450G FB CL V23 REC", codCaja: "MTC1542", nomCaja: "CJ MACRO ORG. RASPBERRIES 9X450G K-V23", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.05, cliente: "Macro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T2-06-9380", producto: "INGREDION RASP PURE ORG 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-06-9382", producto: "INFRUIT ORG PUREE RASP SEEDLESS 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTCG304", nomCaja: "CJ BLANCA SIN IMPRESION 392X292X160", bolsasXCaja: 1, cajasXPallet: 90, kgXCaja: 13.62, cliente: "Infruit", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-12-0001", producto: "FRAMBUESA IQF GENERICO 1*10KG", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-0301", producto: "HEB RASPBERRIES 12*12 OZ", codBolsa: "MTC1621", nomBolsa: "BOL HEB RASPBERRIES 12 OZ V24", codCaja: "MTC1246", nomCaja: "CJ GENERICA 390X187X153 12*340 GRS", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-0370", producto: "KNOW & LOVE RASPBERRIES 12*12 OZ", codBolsa: "MTC1593", nomBolsa: "BOL KNOW&LOVE RASPBERRIES  12OZ V23", codCaja: "MTC1246", nomCaja: "CJ GENERICA 390X187X153 12*340 GRS", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-0401", producto: "STORE BRAND RASPBERRIES 8*12 OZ", codBolsa: "MTC1644", nomBolsa: "BOL GIANT RASPBERRIES 12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-2229", producto: "QUALITA RASPBERRIES 12*400 GR", codBolsa: "MTC1174", nomBolsa: "BOL DOYPACK QUALITA FRAMBOESAS 400G V-18", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-2259", producto: "FRA IQF HEB 12*12 OZ C/HEB - B/", codBolsa: "", nomBolsa: "", codCaja: "MTC1246", nomCaja: "CJ GENERICA 390X187X153 12*340 GRS", bolsasXCaja: 12, cajasXPallet: 195, kgXCaja: 4.08233133, cliente: "HEB", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-12-2265", producto: "FRA IQF DOL 4*2.5 KG", codBolsa: "MTCG569", nomBolsa: "BOL I.Q.F. AZUL 35X50", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-2266", producto: "FRA IQF HER 4*2.5 KG", codBolsa: "MTCG569", nomBolsa: "BOL I.Q.F. AZUL 35X50", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-2272", producto: "PUBLIX RASPBERRIES 8*12 OZ", codBolsa: "MTC1099", nomBolsa: "BOL DOYPACK PUBLIX RASPBERRIES 12OZ. V18", codCaja: "MTC1158", nomCaja: "CJ PUBLIX RASPBERRIES 8X12OZ. V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.72155422, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-12-2276", producto: "WWS RASPBERRIES 9*500 GR M.O", codBolsa: "MTC1679", nomBolsa: "BOL WW RASPBERRIES  500G. FB CL V25 REC", codCaja: "MTC1377", nomCaja: "CJ WW RASPB. 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 169, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T2-12-2278", producto: "GV RASPBERRIES 5*12 OZ", codBolsa: "MTC1291", nomBolsa: "BOL DOYPACK GV RASPBERRIES 12OZ", codCaja: "MTC1335", nomCaja: "CJ  GV RASPBERRIES 5X12OZ.  SRP", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-2278-CL", producto: "GV RASPBERRIES 5*12 OZ", codBolsa: "MTC1729", nomBolsa: "BOL DOYPACK GV RASPBERRIES  12OZ  CL V25", codCaja: "MTC1335", nomCaja: "CJ  GV RASPBERRIES 5X12OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 110, kgXCaja: 1.7009713875, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T2-12-2279", producto: "GV RASPBERRIES 3*24 OZ", codBolsa: "MTC1289", nomBolsa: "BOL DOYPACK GV RASPBERRIES 24OZ", codCaja: "MTC1336", nomCaja: "CJ  GV RASPBERRIES 3X24OZ.  SRP", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-2281", producto: "SIAM MAKRO RASPBERRIES 15*500 GR", codBolsa: "MTC1405", nomBolsa: "BOLS 100%PE DP DAILY FFROZEN RASPB 500GR", codCaja: "MTC1406", nomCaja: "CJ DAILY FRESH FROZEN RASPB &#160;15X500", bolsasXCaja: 15, cajasXPallet: 100, kgXCaja: 7.5, cliente: "Makro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-12-2282", producto: "602703-SEG RASP 12x12oz.", codBolsa: "MTC1346", nomBolsa: "BOL DOYPACK RASPBERRIES SE GROCERS 12 OZ", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 195, kgXCaja: 4.08233133, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-12-9370", producto: "LOS NIETITOS RASP CRUMBLE CLEAN 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9372", producto: "MEYER RASP CRUMBLE CLEAN 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9374", producto: "SUN-IN RASP CRUMBLE CLEAN 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9376", producto: "SUN-IN RASPBERRIES DOLOMIA 4*2,5KG", codBolsa: "MTCG569", nomBolsa: "BOL I.Q.F. AZUL 35X50", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9381", producto: "SVZ RASP CRUMBLE SEMI-CLEAN 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9382", producto: "SVZ RASP PURE 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9385", producto: "FRA PURE S/S INFRUIT 1*13.62 CONGELADAS", codBolsa: "", nomBolsa: "", codCaja: "MTCG304", nomCaja: "CJ BLANCA SIN IMPRESION 392X292X160", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9386", producto: "ANDROS RASPBERRIES PURE CONV 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9388", producto: "SUPERIOR FOODS RASPB PURE CONV 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-12-9393", producto: "FLAGFOOD RASPB UNCLEAN CRUMBLE 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-15-9373", producto: "WOOYANG PUREE RASP SEEDLESS 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: 1, cajasXPallet: 90, kgXCaja: 13.62, cliente: "Wooyang", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T2-15-9388", producto: "RASPERRIES CRUMBLE UNCLEAN 1*13.62 KG", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-15-9389", producto: "RASPERRIES CRUMBLE SEMI-CLEAN 1*13.62 KG", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-15-9390", producto: "RASPERRIES CRUMBLE CLEAN 1*13.62 KG", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T2-16-9361", producto: "INFRUIT RASPBERRIES PURE CONV 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTCG304", nomCaja: "CJ BLANCA SIN IMPRESION 392X292X160", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-01-1126", producto: "BJS ORG CHERRIES 7*3 LB", codBolsa: "MTC1227", nomBolsa: "BOL WELLSLEY FARMS ORG S.DARK CHER. 3LBS", codCaja: "MTC1231", nomCaja: "CJ CHERRIES ORG WELLSLEY 3 LB EXHIBIDORA", bolsasXCaja: 7, cajasXPallet: 63, kgXCaja: 9.52543977, cliente: "Wellsy Farms", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-02-0001", producto: "CEREZAS ORG IQF GENERICO 1*30LB SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-02-7042", producto: "GVO CHERRIES 5*10 OZ", codBolsa: "MTC1283", nomBolsa: "BOL DOYPACK GVOrg. CHERRIES 10OZ", codCaja: "MTC1317", nomCaja: "CJ GVO CHERRIES  5X10OZ.  SRP", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-11-1122", producto: "CER IQF AEON 14*150 GRS. CJ-AEON-B/AEON", codBolsa: "MTC1075", nomBolsa: "BOL TOPVALU (A.MET) D.SWEET CHERRY 150G", codCaja: "MTC1083", nomCaja: "CJ TOPVALU DARK SWEET CHERRY 14X150G", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-12-0175", producto: "GLOBAL GROWERS CHERRIES 12*280 GR", codBolsa: "MTC1495", nomBolsa: "BOL GLOBAL GROWERS CHERRIES  280G", codCaja: "MTC1496", nomCaja: "CJ GLOBAL GROWERS CHERRIES 12X280G", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-12-0251", producto: "CEREZA IQF KOBE BUSSAN 24*500 GR", codBolsa: "MTC1470", nomBolsa: "FILM KBUSSAN AMERICAN CHERRIES  500G", codCaja: "MTC1471", nomCaja: "CJ KBUSSAN AMERICAN CHERRIES 24X500G", bolsasXCaja: 24, cajasXPallet: 90, kgXCaja: 12, cliente: "Kobbe Busan", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-0301", producto: "HEB CHERRIES 12*16 OZ", codBolsa: "MTC1455", nomBolsa: "BOL HEB DARK S. CHERRIES 16OZ 2 SELLO", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "HEB", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-0370", producto: "KNOW & LOVE CHERRIES 12*12 OZ", codBolsa: "MTC1590", nomBolsa: "BOL KNOW&LOVE DARK S. CHERRIES 12OZ V23", codCaja: "MTC1246", nomCaja: "CJ GENERICA 390X187X153 12*340 GRS", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-12-0401", producto: "STORE BRAND CHERRIES 8*12 OZ", codBolsa: "MTC1646", nomBolsa: "BOL GIANT DARK S. CHERRIES 12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-12-0402", producto: "STORE BRAND CHERRIES 6*36 OZ", codBolsa: "MTC1662", nomBolsa: "BOL GIANT SWEET   BARK CHERRIES 36OZ V24", codCaja: "MTC1703", nomCaja: "CJ GENERICA 392X292X140 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-12-1117", producto: "PUBLIX CHERRIES 6*48 OZ", codBolsa: "MTC1253", nomBolsa: "BOL DOYPACK PUBLIX CHERRIES 48OZ. V20", codCaja: "MTCG942", nomCaja: "CJ PUBLIX DARK CHERRIES 6/48OZ.", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-1121", producto: "CER IQF K.SHINTOA 1*30 LB.C/N-B/AZUL", codBolsa: "MTC", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T4-12-1124", producto: "PUBLIX CHERRIES 8*12 OZ", codBolsa: "MTC1252", nomBolsa: "BOL DOYPACK PUBLIX CHERRIES 12OZ. V20", codCaja: "MTC1155", nomCaja: "CJ PUBLIX DARK CHERRIES 8X12OZ. V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.72155422, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-1130", producto: "GV CHERRIES 5*16 OZ", codBolsa: "MTC1295", nomBolsa: "BOL DOYPACK GV CHERRIES 16OZ", codCaja: "MTC1326", nomCaja: "CJ  GV CHERRIES   5X16OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 220, kgXCaja: 2.2679618500000003, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T4-12-1131", producto: "GV CHERRIES 6*40 OZ", codBolsa: "MTC1762", nomBolsa: "BOL 3 SELLO GV CHERRIES 40OZ", codCaja: "MTC1332", nomCaja: "CJ  GV CHERRIES  6X40OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 117, kgXCaja: 6.803885550000001, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T4-12-1133", producto: "602702-SEG DK CHERRIES 12x12oz.", codBolsa: "MTC1344", nomBolsa: "BOL DPACK D.SW. CHERRIES SE GROCERS 12OZ", codCaja: "MTC1360", nomCaja: "CJ DK.SWEET CHERRIES SE GROCERS 12X12OZ", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 4.08233133, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-1134", producto: "SIAM MAKRO CHERRIES 15*500 GR", codBolsa: "MTC1394", nomBolsa: "BOLS 100%PE DAILY FRESH CHERRIES 500 GRS", codCaja: "MTC1399", nomCaja: "CJ DAILY FRESH FROZEN PCHERRIES 15X500GR", bolsasXCaja: 15, cajasXPallet: 100, kgXCaja: 7.5, cliente: "Makro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T4-12-1137", producto: "WWS CHERRIES 9*500 GR", codBolsa: "MTC1678", nomBolsa: "BOL WW CHERRIES  500G. FB CL V25 REC", codCaja: "MTC1383", nomCaja: "CJ WW CHERRIES 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T4-12-7031", producto: "TRADER JOES CHERRIES 24*16 OZ", codBolsa: "MTC1423", nomBolsa: "BOL DOYPACK TRADER JOE’S CHERRIES 16 OZ", codCaja: "MTC1425", nomCaja: "CJ TRAD JOE’S P.D.SWEET CHERRIES 24*16OZ", bolsasXCaja: 24, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Trader Joe", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-01-0192", producto: "BJS ORG BLUEBERRIES 7*3 LB", codBolsa: "MTC1228", nomBolsa: "BOL WELLSLEY FARMS ORG. BLUEBER 3 LBS", codCaja: "MTC1232", nomCaja: "CJ BLUEBER ORG WELLSLEY 3 LB EXHIBIDORA", bolsasXCaja: 7, cajasXPallet: 63, kgXCaja: 9.52543977, cliente: "Wellsy Farms", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-02-0001", producto: "ARANDANOS ORGANICO IQF 1*30 LBS", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-02-0167", producto: "GREENWISE ORG BLUEBERRIES 6*3 LB", codBolsa: "MTC1713", nomBolsa: "BOL DOYPACK GREENWISE BLUEBER 48OZ V25", codCaja: "MTC1068", nomCaja: "CJ GREENWISE ORGANIC BLUEBERRIES 6X48OZ", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-02-0183", producto: "GREENWISE ORG BLUEBERRIES 8*10 OZ", codBolsa: "MTC1633", nomBolsa: "BOL DOYPACK GREENWISE BLUEBER 10OZ. V24", codCaja: "MTC1161", nomCaja: "CJ GREENWISE ORG BLUEBERRIES 8X10OZ V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.2679618500000003, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-02-0199", producto: "ARA IQF ORG G.VALUE 5*10 Oz C/G.VL - B/G", codBolsa: "MTC1282", nomBolsa: "BOL DOYPACK GVORG. BLUEBERRY 10OZ", codCaja: "MTC1322", nomCaja: "CJ  GVO BLUEBERRY   5X10OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 253, kgXCaja: 1.41747615625, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T5-02-0204", producto: "602700-SEGNB ORG BLUE 12x10oz.", codBolsa: "MTC1408", nomBolsa: "BOL DOYPACK BLUEBER ORG SE GROCERS 10OZ", codCaja: "MTC1357", nomCaja: "CJ ORG BLUEBERRIES SE GROCERS 12X10OZ", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 3.4019427750000006, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-02-0350", producto: "WWS MACRO BLUEBERRIES 9*450 GR", codBolsa: "MTC1615", nomBolsa: "BOL MACRO ORG BLUEB  450G FB CL V23 REC", codCaja: "MTC1539", nomCaja: "CJ MACRO ORG BLUEBERRIES 9X450G  K-V23", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.05, cliente: "Macro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T5-02-0401", producto: "NATURES PROMISE ORG BLUEBERRIES 8*10 OZ", codBolsa: "MTC1687", nomBolsa: "BOL NATURES PROMISE ORG BLUEB 10OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-02-0402", producto: "NATURES PROMISE ORG BLUEBERRIES 6*32 OZ", codBolsa: "MTC1690", nomBolsa: "BOL NATURES PROMISE ORG BLUEB 32OZ V24", codCaja: "MTC1703", nomCaja: "CJ GENERICA 392X292X140 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-06-0139", producto: "INFRUIT ORG PUREE BLUEB SEEDLESS 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTCG304", nomCaja: "CJ BLANCA SIN IMPRESION 392X292X160", bolsasXCaja: 1, cajasXPallet: 90, kgXCaja: 13.62, cliente: "Infruit", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-11-0070", producto: "ARA IQF KOBE BUSSAN 24*500 GR.C/KB-B/KB", codBolsa: "MTCG934", nomBolsa: "BOL LAMINA KOBE BUSSAN 500G. VER 2016.1", codCaja: "MTC1195", nomCaja: "CJ KOBE BUSSAN BLUEBER 500GX24  BBD  V19", bolsasXCaja: 24, cajasXPallet: 90, kgXCaja: 12, cliente: "Kobbe Busan", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-11-0160", producto: "AEON BLUEBERRIES 10*500 GR", codBolsa: "MTC1618", nomBolsa: "BOL TOPVALU BLUEBERRIES 500G V24", codCaja: "MTC1717", nomCaja: "CJ TOPVALU BLUEBERRIES 10X500 GRS V25", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-11-0163", producto: "ARA IQF AEON 10*350 GRS. CJ-AEON-B/AEON", codBolsa: "MTC1076", nomBolsa: "BOL TOPVALU BLUEBERRIES 350G", codCaja: "MTC1082", nomCaja: "CJ TOPVALU BLUEBERRIES 10X350G", bolsasXCaja: 10, cajasXPallet: 216, kgXCaja: 3.5, cliente: "AEON", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-11-0251", producto: "KOBE BUSSAN BLUEBERRIES 24*500 GR", codBolsa: "", nomBolsa: "", codCaja: "MTC1507", nomCaja: "CJ KOBE BUSSAN BLUEB 500GX24 V23", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0001", producto: "ARANDANOS IQF GENERICO 1*30 LBS", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0028", producto: "PUBLIX BLUEBERRIES 6*48 OZ", codBolsa: "MTC1092", nomBolsa: "BOL DOYPACK PUBLIX BLUEBERRIES 48OZ V18", codCaja: "MTCG385", nomCaja: "CJ PUBLIX BLUEBERRIES 6/48OZ.", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0049", producto: "QUALITA BLUEBERRIES 12*400 GR", codBolsa: "MTC1175", nomBolsa: "BOL DOYPACK QUALITA MIRTILOS 400G. V-18", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0094", producto: "HANMI BLUEBERRIES 1*30LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0095", producto: "ITOCHU BLUEBERRIES 1*15 KG", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0118", producto: "ECONO BLUEBERRIES 12*1,5 LB", codBolsa: "MTC1197", nomBolsa: "BOL ECONO BLUEBER 1.5 LB./680.38G.  V19", codCaja: "MTCG791", nomCaja: "CJ ECONO BLUEBER 12x1.5Lb./680,38g", bolsasXCaja: 12, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Econo", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0130", producto: "CHIQUITA BLUEBERRIES 8*2,5 LB", codBolsa: "MTC1736", nomBolsa: "BOL CHIQUITA BLUEBER  8/2.5 LB V25", codCaja: "MTCG833", nomCaja: "CJ CHIQUITA BLUEBERRIES 2.5Lb.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 9.071847400000001, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0161", producto: "CHIQUITA BLUEBERRIES 8*12 OZ", codBolsa: "MTC1236", nomBolsa: "BOLS 100%PE CHIQUITA BLUEBER  8/12 OZ", codCaja: "MTC1150", nomCaja: "CJ CHIQUITA BLUEBERRIES 8X12OZ. V19", bolsasXCaja: 8, cajasXPallet: 280, kgXCaja: 2.72155422, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0171", producto: "BLUEBERRY BEISIA 10*1 KG", codBolsa: "MTC1499", nomBolsa: "BOL BEISIA BLUEBERRIES 1 KL", codCaja: "MTC1500", nomCaja: "CJ BEISIA BLUEBERRIES 1X10 KL", bolsasXCaja: 10, cajasXPallet: 100, kgXCaja: 10, cliente: "Beisia", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0175", producto: "GLOBAL GROWERS BLUEBERRIES 12*280 GR", codBolsa: "MTC1124", nomBolsa: "BOL GLOBAL GROWERS BLUEBERRIES 280G", codCaja: "MTC1128", nomCaja: "CJ GLOBAL GROWERS BLUEBERRIES 12X280G", bolsasXCaja: 12, cajasXPallet: 198, kgXCaja: 3.36, cliente: "Global Growers", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0184", producto: "PUBLIX BLUEBERRIES 8*12 OZ", codBolsa: "MTC1091", nomBolsa: "BOL DOYPACK PUBLIX BLUEBERRIES 12OZ V18", codCaja: "MTC1154", nomCaja: "CJ PUBLIX BLUEBERRIES 8X12OZ. V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.72155422, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0196", producto: "GV BLUEBERRIES 10*16 OZ", codBolsa: "MTC1293", nomBolsa: "BOL DOYPACK GV BLUEBERRY 16OZ", codCaja: "MTC1314", nomCaja: "CJ GV BLUEBERRY 10X16OZ.  SRP", bolsasXCaja: 10, cajasXPallet: 110, kgXCaja: 4.535923700000001, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T5-12-0196-CL", producto: "GV BLUEBERRIES 10*16 OZ", codBolsa: "MTC1725", nomBolsa: "BOL DOYPACK GV BLUEBERRIES 16OZ  CL V25", codCaja: "MTC1314", nomCaja: "CJ GV BLUEBERRY 10X16OZ.  SRP", bolsasXCaja: 10, cajasXPallet: 55, kgXCaja: 4.535923700000001, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T5-12-0198", producto: "GV BLUEBERRIES 6*48 OZ", codBolsa: "MTC1763", nomBolsa: "BOL 3 SELLO GV BLUEBERRY 48OZ", codCaja: "MTC1315", nomCaja: "CJ GV BLUEBERRY  6X48OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 8.16466266, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "T5-12-0204", producto: "GARDEN FOODS BLUEBERRIES 8*3 LB", codBolsa: "MTC1587", nomBolsa: "BOL GARDEN FOODS BLUEBERIES 48OZ", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0208", producto: "602708-SEG BLUE  12x16oz.", codBolsa: "MTC1343", nomBolsa: "BOL DOYPACK BLUEBERRIES SE GROCERS 16 OZ", codCaja: "MTC1359", nomCaja: "CJ BLUEBERRIES SE GROCERS 12X16OZ", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0209", producto: "602715-SEG BLUE 6x48oz.", codBolsa: "MTC1352", nomBolsa: "BOL DOYPACK BLUEBERRIES SE GROCERS 48 OZ", codCaja: "MTC1368", nomCaja: "CJ BLUEBERRIES SE GROCERS 6X48OZ", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0210", producto: "BOGOPA BLUEBERRIES 6*48 OZ", codBolsa: "MTC1420", nomBolsa: "BOL DOYPACK BOGOPA BLUEBERRIES 3 LB", codCaja: "MTC1438", nomCaja: "CJ GENERICA 392X292X170  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0211", producto: "GARDEN FOODS BLUEBERRIES 12*1 LB", codBolsa: "MTC1224", nomBolsa: "BOL DOYPACK GARDEN FOODS BLUEBER 454 G", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Garden Food", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0230", producto: "PRICESMART BLUEBERRIES 10*48 OZ", codBolsa: "MTC1692", nomBolsa: "BOL MEMBERS BLUEBERRIES 48OZ CL V24", codCaja: "MTC1709", nomCaja: "CJ MEMBERS BLUEB 10X48OZ V24", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0310", producto: "ARCTIC HARVEST BLUEBERRIES 8*1,5 KG", codBolsa: "MTC1554", nomBolsa: "BOL ARTIC HARVEST BLUEB 1.5K", codCaja: "MTC1556", nomCaja: "CJ ARTIC HARVEST BLUEB 8X1,5K", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0350", producto: "WWS BLUEBERRIES 9*500 GR", codBolsa: "MTC1676", nomBolsa: "BOL WW BLUEBER 500G. FB CL V25 REC", codCaja: "MTC1380", nomCaja: "CJ WW BLUEBER. 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T5-12-0351", producto: "WWS BLUEBERRIES 10*1 KG", codBolsa: "MTC1675", nomBolsa: "BOL WW BLUEBERRIES 1 KG CL V25 REC", codCaja: "MTC1373", nomCaja: "CJ WW BLUEBERRIES 10 X 1 KL. C30KRAFT", bolsasXCaja: 10, cajasXPallet: 100, kgXCaja: 10, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n" },
  { sku: "T5-12-0370", producto: "KNOW & LOVE BLUEBERRIES 12*16 OZ", codBolsa: "MTC1588", nomBolsa: "BOL KNOW&LOVE BLUEBERRIES 16OZ V23", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0371", producto: "KNOW & LOVE BLUEBERRIES 6*48 OZ", codBolsa: "MTC1589", nomBolsa: "BOL KNOW&LOVE BLUEBERRIES 48OZ V23", codCaja: "MTC1704", nomCaja: "CJ GENERICA 392X292X170 KRAFT", bolsasXCaja: 6, cajasXPallet: 120, kgXCaja: 8.16466266, cliente: "Know&Love", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T5-12-0401", producto: "STORE BRAND BLUEBERRIES 8*12 OZ", codBolsa: "MTC1645", nomBolsa: "BOL GIANT BLUEBERRIES 12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0402", producto: "STORE BRAND BLUEBERRIES 4*48 OZ", codBolsa: "MTC1664", nomBolsa: "BOL GIANT BLUEBERRIES 48OZ V24", codCaja: "MTC1703", nomCaja: "CJ GENERICA 392X292X140 KRAFT", bolsasXCaja: 4, cajasXPallet: 140, kgXCaja: 5.44310844, cliente: "Ahold", tipoPallet: "Yugo", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0403", producto: "STORE BRAND BLUEBERRIES 4*64 OZ", codBolsa: "MTC1670", nomBolsa: "BOL GIANT BLUEBERRY 64OZ V24", codCaja: "MTC1704", nomCaja: "CJ GENERICA 392X292X170 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0404", producto: "NO USAR", codBolsa: "MTC1670", nomBolsa: "BOL GIANT BLUEBERRY 64OZ V24", codCaja: "MTC1704", nomCaja: "CJ GENERICA 392X292X170 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-12-0405", producto: "STORE BRAND BLUEBERRIES 6*64 OZ", codBolsa: "MTC1670", nomBolsa: "BOL GIANT BLUEBERRY 64OZ V24", codCaja: "MTC1704", nomCaja: "CJ GENERICA 392X292X170 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-16-1011", producto: "INFRUIT BLUEBERRIES JUGO 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T5-16-1012", producto: "INFRUIT BLUEBERRIES JUGO 1*15 KG.", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-0001", producto: "MORA IQF GENERICO 1*30LBS", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-0003", producto: "BOYSENBERRY IQF 30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-0011", producto: "MORA IQF GENERICO 1*30LBS SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-0370", producto: "KNOW & LOVE BLACKBERRIES 12*16 OZ", codBolsa: "MTC1591", nomBolsa: "BOL KNOW&LOVE BLACKBERRIES 16OZ V23", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-0401", producto: "STORE BRAND BLACKBERRIES 8*16 OZ", codBolsa: "MTC1657", nomBolsa: "BOL GIANT BLACKBERRIES 16OZ V24", codCaja: "MTC1702", nomCaja: "CJ GENERICA 295X215X169 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-0995", producto: "MOR CULT IQF CHER KANEMATSU 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-1013", producto: "GLOBAL GROWERS BLACKBERRIES 12*280 GR", codBolsa: "MTC1123", nomBolsa: "BOL GLOBAL GROWERS BLACKBERRIES 280G", codCaja: "MTC1127", nomCaja: "CJ GLOBAL GROWERS BLACKBERRIES 12X280G", bolsasXCaja: 12, cajasXPallet: 198, kgXCaja: 3.36, cliente: "Global Growers", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "T8-12-1118", producto: "WWS BLACKBERRIES 9*500 GR", codBolsa: "MTC1706", nomBolsa: "BOL WW BLACKBERRIES 500G. FB CL V25 REC", codCaja: "MTC1379", nomCaja: "CJ WW BLACKB. 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T8-12-1119", producto: "GV BLACKBERRIES 5*16 OZ", codBolsa: "MTC1288", nomBolsa: "BOL DOYPACK GV BLACKBERRY 16OZ", codCaja: "MTC1324", nomCaja: "CJ  GV BLACKBERRY  5X16OZ.  SRP", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-12-3118", producto: "BOYSENBERR IQF WWS 9*500GRS. 4,5KG", codBolsa: "MTC1278", nomBolsa: "BOL WW BOYSEN 500G FBOTTOM V20", codCaja: "MTC1381", nomCaja: "CJ WW BOYSENB. 9*500G FLATBOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "T8-13-1012", producto: "MORA CULT EXP. PMT S.A 1*30 LB.C/N-B/A", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T8-16-1011", producto: "INFRUIT JUGO BLACKBERRIES  1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T9-12-0001", producto: "KIWI GRANEL 1*30 LBS", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T9-12-0011", producto: "KIWI GRANEL 1*30 LBS SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "T9-12-7006", producto: "GARDEN FOODS KIWI 12*1 LB", codBolsa: "MTC1225", nomBolsa: "BOL DOYPACK GARDEN FOODS KIWI 454G", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Garden Food", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-02-0001", producto: "FRUTILLA ORG IQF GENERICO 1*30LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-02-0011", producto: "FRUTILLA ORG IQF GENERICO 1*30LB SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-02-0401", producto: "NATURES PROMISE ORG STRAWBERRIES 8*10 OZ", codBolsa: "MTC1685", nomBolsa: "BOL NATURES PROMISE ORG STRAW 10OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-02-0402", producto: "NATURES PROMISE ORG STRAWBERRIES 6*32 OZ", codBolsa: "MTC1691", nomBolsa: "BOL NATURES PROMISE ORG STRAW 32OZ V24", codCaja: "MTC1703", nomCaja: "CJ GENERICA 392X292X140 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-02-1464", producto: "HEB ORG WHOLE STRAWBERRIES 12*10 OZ", codBolsa: "MTC1414", nomBolsa: "BOL DPACK HEB ORG. WHOLE STRAW 10OZ&#160", codCaja: "MTCG966", nomCaja: "CJ GENERICA 390X187X143 KRAFT", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 3.4019427750000006, cliente: "HEB", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-02-1467", producto: "GREENWISE ORG WHOLE STRAWBERRIES 8*10 OZ", codBolsa: "MTC1714", nomBolsa: "BOL DOYPACK GREENWISE STRAWB 10OZ V25", codCaja: "MTC1164", nomCaja: "CJ GREENWISE ORG STRAWBERRIES 8X10OZ V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.2679618500000003, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-02-1478", producto: "WWS MACRO WHOLE STRAWBERRIES 9*450 G", codBolsa: "MTC1613", nomBolsa: "BOL MACRO ORG STRAWB 450G FB CL V23 REC", codCaja: "MTC1543", nomCaja: "CJ MACRO ORG STRAWBERRIES 9X450G  K-V23", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.05, cliente: "Macro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TA-02-1480", producto: "602701-SEGNB ORG STRAW 12x10oz.", codBolsa: "MTC1402", nomBolsa: "BOL DPACK WHOLE STRAWB ORG GROCERS 10OZ", codCaja: "MTC1358", nomCaja: "CJ ORG WHOLE STRAWB SE GROCERS 12X10OZ", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-02-1489", producto: "GVO SLCD STRAWBERRIES 10*10 OZ", codBolsa: "MTC1465", nomBolsa: "BOL DOYPACK GVO ORG.  STRAWB 10OZ", codCaja: "MTC1480", nomCaja: "CJ GVO STRAWB   10X10OZ.  SRP", bolsasXCaja: 10, cajasXPallet: 253, kgXCaja: 2.8349523125, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-06-1436", producto: "INFRUIT ORG PUREE STRAW SEEDLESS 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTCG304", nomCaja: "CJ BLANCA SIN IMPRESION 392X292X160", bolsasXCaja: 1, cajasXPallet: 90, kgXCaja: 13.62, cliente: "Infruit", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-11-1443", producto: "PMT DICED STRAWBERRIES 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0001", producto: "FRUTILLA SLIC IQF GENERICO 1*30LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0003", producto: "FRUTILLA IQF GENERICO 1*30LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0011", producto: "FRUTILLA IQF GENERICO 1*30LB SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0012", producto: "FRUTILLA SLICED IQF GENERICO 1*30LB SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0230", producto: "PRICESMART WHOLE STRAWBERRIES 6*80 OZ", codBolsa: "MTC1694", nomBolsa: "BOL MEMBERS WHOLE STRAW 80OZ CL V24", codCaja: "MTC1711", nomCaja: "CJ MEMBERS WHOLE STRAW 5X80OZV24", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0350", producto: "WWS SLCD STRAWBERRIES 9*500 GR M.O", codBolsa: "MTC1677", nomBolsa: "BOL WW SLICED STRAWB 500G FB CL V25 REC", codCaja: "MTC1378", nomCaja: "CJ WW STRAWB. 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TA-12-0370", producto: "KNOW & LOVE WHOLE STRAWBERRIES 12*16 OZ", codBolsa: "MTC1601", nomBolsa: "BOL KNOW&LOVE STRAWBERRIES 16OZ V23", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0371", producto: "KNOW & LOVE WHOLE STRAWBERRIES 6*48 OZ", codBolsa: "MTC1594", nomBolsa: "BOL KNOW&LOVE STRAWBERRIES 48OZ V23", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "Know&Love", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-0401", producto: "STORE BRAND WHOLE STRAWBERRIES 8*16 OZ", codBolsa: "MTC1656", nomBolsa: "BOL GIANT WHOLE STRAWBERRIES 16OZ V24", codCaja: "MTC1702", nomCaja: "CJ GENERICA 295X215X169 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0403", producto: "STORE BRAND WHOLE STRAWBERRIES 4*64 OZ", codBolsa: "MTC1668", nomBolsa: "BOL GIANT WHOLE STRAWBERRIES 64OZ V24", codCaja: "MTC1704", nomCaja: "CJ GENERICA 392X292X170 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0404", producto: "STORE BRAND SLCD STRAWBERRIES 8*16 OZ", codBolsa: "MTC1655", nomBolsa: "BOL GIANT SLICED STRAWBERRIES 16OZ V24", codCaja: "MTC1702", nomCaja: "CJ GENERICA 295X215X169 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-1216", producto: "PUBLIX WHOLE STRAWBERRIES 6*64 OZ", codBolsa: "MTC1102", nomBolsa: "BOL DOYPACK PUBLIX STRAWBERRIES 64OZ V18", codCaja: "MTCG265", nomCaja: "CJ PUBLIX STRAWB 6/64OZ.", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1415", producto: "FRU IQF DICED HANMI 10mm 1*30LB C/N-B/AZ", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-1419", producto: "FRU IQF CUBO ALBION INA 11.1MM 1*13,6 C/", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-1427", producto: "GOYA WHOLE STRAWBERRIES 12*16 OZ", codBolsa: "MTC1435", nomBolsa: "BOL DOYPACK GOYA STRAWBERRIES 16OZ. V21", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Goya", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1429", producto: "FRU IQF KANEMATSU SHINTOA MEDIUM 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-1432", producto: "CHIQUITA WHOLE STRAWBERRIES 8*2,5 LB", codBolsa: "MTC1239", nomBolsa: "BOLS 100%PE CHIQUITA STRAWB  8/2.5 LB", codCaja: "MTCG835", nomCaja: "CJ CHIQUITA STRAWBERRIES 2.5Lb.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 9.071847400000001, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1441", producto: "FRU IQF DICD ALB K.SHINT 3/8x5/8x3/8 30L", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-1451", producto: "CEN MARKET STRAWBERRIES SABRINA 12*10 OZ", codBolsa: "MTC1411", nomBolsa: "BOL DOYP C. MARKET S. STRAW 10OZ V21", codCaja: "MTCG966", nomCaja: "CJ GENERICA 390X187X143 KRAFT", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 3.4019427750000006, cliente: "Central Market", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1459", producto: "CHIQUITA WHOLE STRAWBERRIES 8*16 OZ", codBolsa: "MTC1738", nomBolsa: "BOL CHIQUITA STRAWBERRIES 16OZ  V25", codCaja: "MTC1152", nomCaja: "CJ CHIQUITA STRAWBERRIES 8X16OZ. V19", bolsasXCaja: 8, cajasXPallet: 204, kgXCaja: 3.62873896, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1462", producto: "GLOBAL GROWERS WHOLE STRAW 12*280 GR", codBolsa: "MTC1125", nomBolsa: "BOL GLOBAL GROWERS STRAWBERRIES 280G", codCaja: "MTC1129", nomCaja: "CJ GLOBAL GROWERS STRAWBERRIES 12X280G", bolsasXCaja: 12, cajasXPallet: 198, kgXCaja: 3.36, cliente: "Global Growers", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1463", producto: "PUBLIX WHOLE STRAWBERRIES 8*16 OZ", codBolsa: "MTC1101", nomBolsa: "BOL DOYPACK PUBLIX STRAWBERRIES 16OZ V18", codCaja: "MTC1160", nomCaja: "CJ PUBLIX STRAWBERRIES 8X16OZ. V19", bolsasXCaja: 8, cajasXPallet: 204, kgXCaja: 3.62873896, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1466", producto: "FRU IQF WWS 9*500 GR.C/WW-B/WW FLT B.BAG", codBolsa: "MTC1490", nomBolsa: "BOL WW SLICEDSTRAWB 500G FB MULT-O HSR", codCaja: "MTC1378", nomCaja: "CJ WW STRAWB. 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-1474", producto: "FRU SLICE WWS 9*500 GR (M.ORIGEN) 4.5 KG", codBolsa: "MTC1265", nomBolsa: "BOL WW SLICEDSTRAWB 500G FLATBOTTOM MULT", codCaja: "MTC1378", nomCaja: "CJ WW STRAWB. 9*500G. FLAT-BOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TA-12-1475", producto: "GV WHOLE STRAWBERRIES 6*64 OZ", codBolsa: "MTC1769", nomBolsa: "BOL 3 SELLO GV WHOLE STRAWB 64OZ", codCaja: "MTC1321", nomCaja: "CJ GV WHOLE STRAWB  6X64OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 70, kgXCaja: 10.88621688, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-12-1476", producto: "GV SLCD STRAWBERRIES 6*64 OZ", codBolsa: "MTC1307", nomBolsa: "BOL 3 SELLO GV SLCD STRAWB 64OZ", codCaja: "MTC1316", nomCaja: "CJ GV SLCD STRAWB  6X64OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 70, kgXCaja: 10.88621688, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-12-1477", producto: "GV WHOLE STRAWBERRIES 10*16 OZ", codBolsa: "MTC1294", nomBolsa: "BOL DOYPACK GV WHOLE STRAWB 16OZ", codCaja: "MTC1320", nomCaja: "CJ GV WHOLE STRAWB  10X16OZ.  SRP", bolsasXCaja: 10, cajasXPallet: 110, kgXCaja: 4.535923700000001, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-12-1477-CL", producto: "GV WHOLE STRAWBERRIES 10*16 OZ", codBolsa: "MTC1730", nomBolsa: "BOL DOYPACK GV STRAWBERRIES 16OZ  CL V25", codCaja: "MTC1320", nomCaja: "CJ GV WHOLE STRAWB  10X16OZ.  SRP", bolsasXCaja: 10, cajasXPallet: 55, kgXCaja: 4.535923700000001, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TA-12-1481", producto: "SIAM MAKRO SLCD STRAWBERRIES 15*500 GR", codBolsa: "MTC1395", nomBolsa: "BOL 100%PE DP DAILY FSLICED STRAWB 500GR", codCaja: "MTC1400", nomCaja: "CJ DAILY FRESH FROZEN S. STRAWB 15X500GR", bolsasXCaja: 15, cajasXPallet: 100, kgXCaja: 7.5, cliente: "Makro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1482", producto: "602711-SEG WHL STRAW 12x16oz.", codBolsa: "MTC1347", nomBolsa: "BOL DPACK WHOLE STRAWB SE GROCERS 16 OZ", codCaja: "MTC1363", nomCaja: "CJ WHOLE STRAWBERRIES SE GROCERS 12X16OZ", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1483", producto: "602714-SEG WHL STRAW 6x48oz.", codBolsa: "MTC1353", nomBolsa: "BOL DPACK WHOLE STRAWB SE GROCERS 48 OZ", codCaja: "MTC1369", nomCaja: "CJ WHOLE STRAWBERRIES SE GROCERS 6X48OZ", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1487", producto: "FRUT SLICE GROCERS POTES 12*1.45 LB.C/AZ", codBolsa: "", nomBolsa: "", codCaja: "MTC1431", nomCaja: "CJ SE-GROCERS  SLICED TUB STRAWB 12X15,5", bolsasXCaja: 12, cajasXPallet: 80, kgXCaja: 7.892507238, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1488", producto: "FRUT SLICE GROCERS POTES 12*15.5 OZ.C/AZ", codBolsa: "", nomBolsa: "", codCaja: "MTC1432", nomCaja: "CJ SE- GROCERS SLICED TUB STRAWB 12X23,2", bolsasXCaja: 12, cajasXPallet: 90, kgXCaja: 5.27301130125, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-12-1489", producto: "PUBLIX SLCD STRAWBERRIES POT 12*16 OZ", codBolsa: "", nomBolsa: "", codCaja: "MTC1485", nomCaja: "CJ PUBLIX TUB S-STRAWB 12X16OZ  V22", bolsasXCaja: 12, cajasXPallet: 120, kgXCaja: 5.44310844, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TA-13-0001", producto: "FRUTILLA SLICED TUB  IQF 1*30LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-15-1444", producto: "PMT SLCD STRAWBERRIES 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-16-1452", producto: "INFRUIT PUREE STRAW SEEDLESS 1*30 LB", codBolsa: "", nomBolsa: "", codCaja: "MTCG304", nomCaja: "CJ BLANCA SIN IMPRESION 392X292X160", bolsasXCaja: 1, cajasXPallet: 90, kgXCaja: 13.62, cliente: "Infruit", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TD-12-0001", producto: "DURAZNOS IQF GENERICO 1*30 LBS", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TD-12-0011", producto: "DURAZNOS IQF GENERICO 1*30 LBS SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TD-12-0020", producto: "JUST QUALITY SLCD PEACHES IQF 2*5 LB", codBolsa: "MTCG569", nomBolsa: "BOL I.Q.F. AZUL 35X50", codCaja: "MTC1672", nomCaja: "CJ GENERICA 286mmX220mmX163mm KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TD-12-0125", producto: "PUBLIX SLCD PEACHES 8*20 OZ", codBolsa: "MTC1100", nomBolsa: "BOL DPACK PUBLIX SLICED PEACHER 20OZ V18", codCaja: "MTC1159", nomCaja: "CJ PUBLIX SLICED PEACHES 8X20OZ. V19", bolsasXCaja: 8, cajasXPallet: 156, kgXCaja: 4.535923700000001, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TD-12-0128", producto: "602707-SEG SLD PEACHES 12x16oz.", codBolsa: "MTC1345", nomBolsa: "BOL DPACK SLICED PEACHES SE GROCERS 16OZ", codCaja: "MTC1361", nomCaja: "CJ SLICED PEACHES SE GROCERS 12X16OZ", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TD-12-0370", producto: "KNOW & LOVE SCLD PEACHES 12*16 OZ", codBolsa: "MTC1592", nomBolsa: "BOL KNOW&LOVE SLICED PEACHES 16OZ V23", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TD-12-0401", producto: "STORE BRAND SLCD PEACHES 8*16 OZ", codBolsa: "MTC1659", nomBolsa: "BOL GIANT SLICED PEACHES 16OZ V24", codCaja: "MTC1702", nomCaja: "CJ GENERICA 295X215X169 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TI-02-0112", producto: "GVO PINEAPPLES 5*10 OZ", codBolsa: "MTC1285", nomBolsa: "BOL DOYPACK GVOorg.  PINEAPPLE 10OZ", codCaja: "MTC1331", nomCaja: "CJ GVO PINEAPPLE 5X10OZ", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TI-11-0106", producto: "PIÑ IQF AEON 14*150 GRS. CJ-AEON-B/AEON", codBolsa: "MTC1603", nomBolsa: "BOL TOPVALU PINEAPPLE 150G V23 TRILAMINA", codCaja: "MTC1052", nomCaja: "CJ TOPVALU (AEON) PINEAPPLE 14*150 GR", bolsasXCaja: 14, cajasXPallet: 288, kgXCaja: 2.1, cliente: "AEON", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TI-12-0029", producto: "PIÑA BOGOPA 6*48 OZ", codBolsa: "MTC1420", nomBolsa: "BOL DOYPACK BOGOPA BLUEBERRIES 3 LB", codCaja: "MTC1438", nomCaja: "CJ GENERICA 392X292X170  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TI-12-0107", producto: "PIÑ IQF CHUNKS CHIQ 8*2.5 LB C/CHQ-B/CHQ", codBolsa: "MTC1258", nomBolsa: "BOLS 100%PE CHIQUITA PINEAPPLE  8/2.5 LB", codCaja: "MTC1261", nomCaja: "CJ CHIQUITA PINEAPPLE  8X2,5LB  V20", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TI-12-0108", producto: "PIÑ IQF CHUNKS CHIQ 8*16 OZ. C/CHQ-B/CHQ", codBolsa: "MTC1259", nomBolsa: "BOLS 100%PE CHIQUITA PINEAPPLE  8/16OZ", codCaja: "MTC1260", nomCaja: "CJ CHIQUITA PINEAPPLE  8X16OZ. V20", bolsasXCaja: 8, cajasXPallet: 204, kgXCaja: 3.62873896, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TI-12-7040-CL", producto: "GV PINEAPPLES 5*16 OZ", codBolsa: "MTC1728", nomBolsa: "BOL DOYPACK GV PINEAPLE 16OZ  CL V25", codCaja: "MTC1329", nomCaja: "CJ GV PINEAPPLE 5X16OZ", bolsasXCaja: 5, cajasXPallet: 110, kgXCaja: 2.2679618500000003, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-02-0063", producto: "TRADER JOES ORG MIX BERRIES 24*12 OZ", codBolsa: "MTC1488", nomBolsa: "BOL TJOES ORG MBERRY BLEND 12 OZ 2 SELLO", codCaja: "MTC1489", nomCaja: "CJ TJOES ORG MBERRY BLEND 24X12 OZ", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-02-0164", producto: "GREENWISE ORG MIX BERRIES 6*3 LB", codBolsa: "MTC1715", nomBolsa: "BOL DPACK GREENWISE MIXBERRIES 48OZ V25", codCaja: "MTC1069", nomCaja: "CJ GREENWISE ORG MIXED BERRIES 6X48OZ", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-02-0171", producto: "MIXED BERRIES ORG WWS (MACRO) 9*450GRS.", codBolsa: "MTC1275", nomBolsa: "BOL MACRO ORG MBERRIES 450G FLATBOTT V20", codCaja: "MTC1386", nomCaja: "CJ MACRO ORG MIXED BERRIES 9X450G KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.05, cliente: "Macro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TM-02-0174", producto: "HEB ORG BERRY BLEND 12*10 OZ", codBolsa: "MTC1415", nomBolsa: "BOL DOYP HEB ORG BERRY BLEND 10 OZ V21", codCaja: "MTCG966", nomCaja: "CJ GENERICA 390X187X143 KRAFT", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 3.4019427750000006, cliente: "HEB", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-02-0177", producto: "GREENWISE ORG MIX BERRIES 8*10 OZ", codBolsa: "MTC1716", nomBolsa: "BOL DPACK GREENWISE MIXBERRIES 10OZ V25", codCaja: "MTC1162", nomCaja: "CJ GREENWISE ORG MIXEDBERRIES 8X10OZ V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.2679618500000003, cliente: "Greenwise", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-02-0187", producto: "GVO TRIPLE BERRY 5*10 OZ", codBolsa: "MTC1287", nomBolsa: "BOL DOYPACK GVO ORG.  TRIP-BERRY 10OZ", codCaja: "MTC1339", nomCaja: "CJ  GVO TRIP-BERRY    5X10OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 253, kgXCaja: 1.41747615625, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-02-0350", producto: "WWS MACRO MIX BERRIES 9*450 GR", codBolsa: "MTC1614", nomBolsa: "BOL MACRO ORG MBERRIE 450G FB CL V23 REC", codCaja: "MTC1541", nomCaja: "CJ MACRO ORG MIXED BERRIES 9X450G K-V23", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.05, cliente: "Macro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TM-02-0401", producto: "NATURES PROMISE ORG BERRY MEDLEY 8*10 OZ", codBolsa: "MTC1688", nomBolsa: "BOL NATURES PROMISE ORG BMEDLEY 10OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-11-0178", producto: "4 MIX BERRIES AEON TV 14*150 GRS.", codBolsa: "MTC1057", nomBolsa: "BOL TOPVALU (A.MET) 4 BERRY MIX 150 GR", codCaja: "MTC1218", nomCaja: "CJ TOPVALU (AEON)  4 BERRY MIX 14*150 GR", bolsasXCaja: 14, cajasXPallet: 288, kgXCaja: 2.1, cliente: "AEON", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-11-0180", producto: "4 MIX BERRIES AEON PREMIUM 10*350 GRS.", codBolsa: "MTC1241", nomBolsa: "BOL DOYPACK TOPVALU BERRY MIX 350 GRS", codCaja: "MTC1242", nomCaja: "CJ TOPVALU BERRY MIX  10X350GRS", bolsasXCaja: 10, cajasXPallet: 216, kgXCaja: 3.5, cliente: "AEON", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-11-0181", producto: "AEON 4 MIX BERRIES 10*500 GR", codBolsa: "MTC1617", nomBolsa: "BOL TOPVALU 4 BERRY MIX 500G V24", codCaja: "MTC1626", nomCaja: "CJ TOPVALU 4 BERRYMIX  10X500 GRS V24", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-11-0185", producto: "AEON 4 MIX BERRIES 10*550 GR", codBolsa: "MTC1641", nomBolsa: "BOL TOPVALU 4 BERRY MIX 550G 50años V24", codCaja: "MTC1643", nomCaja: "CJ TOPVALU 4BERRY MIX 10X550G 50años V24", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0090", producto: "QUALITA MIX FRUTAS 12*400 GR", codBolsa: "MTC1192", nomBolsa: "BOL DPACK QUALITA S. DE FRUTAS 400G V-18", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0127", producto: "ECONO 4 BERRY BLEND 8*3 LB", codBolsa: "MTCG784", nomBolsa: "BOL ECONO 4 BERRY BLEND 3 Lb./1.360g.", codCaja: "MTCG790", nomCaja: "CJ ECONO 4 BERRY BLEND 8x3Lb./1.360gr.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Econo", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0128", producto: "ECONO DELIGHT BLEND 8*3 LB", codBolsa: "MTCG786", nomBolsa: "BOL ECONO DELIGHT BLEND 3 Lb./1.360g.", codCaja: "MTCG792", nomCaja: "CJ ECONO DELIGHT BLEND 8x3Lb./1.360gr", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Econo", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0131", producto: "TROPICAL  BLEND GOYA 12*16 oz.", codBolsa: "MTC1436", nomBolsa: "BOL DPACK GOYA TROPICAL BLEND 16OZ. V21", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Goya", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0132", producto: "GOYA BERRY BLEND 12*16 OZ", codBolsa: "MTC1433", nomBolsa: "BOL DOYPACK GOYA BERRY BLEND 16OZ.   V21", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Goya", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0135", producto: "CHIQUITA GOURMET BERRY MEDLEY 8*2,5 LB", codBolsa: "MTC1734", nomBolsa: "BOL CHIQUITA B.GOURMET 8/2.5LB V25", codCaja: "MTCG831", nomCaja: "CJ CHIQUITA BERRIES MEDLEY 2.5Lb.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 9.071847400000001, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0145", producto: "MIXES FRUIT PUBLIX 6*64 OZ. C/ -B/PUBLIX", codBolsa: "MTC1098", nomBolsa: "BOL DOYPACK PUBLIX MIXED FRUIT 64OZ. V18", codCaja: "MTCG945", nomCaja: "CJ PUBLIX MIXED FRUIT 6/64OZ.", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0147", producto: "PUBLIX TRIPLE BERRY 6*48 OZ", codBolsa: "MTC1103", nomBolsa: "BOL DPACK PUBLIX TRIPLE BERRIES 48OZ V18", codCaja: "MTCG947", nomCaja: "CJ PUBLIX TRIPLE BERRIES 6/48OZ.", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0151", producto: "ECONO SLCD STRAWBERRIES & BANANA 8*3 LB", codBolsa: "MTC1196", nomBolsa: "BOL ECONO BANANA & STRAWB BLEND 3LB  V19", codCaja: "MTCG960", nomCaja: "CJ ECONO BANANA & STRAWB BLEND 8X3LB.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Econo", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0152", producto: "ECONO GOURMET BLEND 8*3 LB", codBolsa: "MTCG959", nomBolsa: "BOL ECONO GOURMET BLEND 3LB.", codCaja: "MTCG961", nomCaja: "CJ ECONO GOURMET BLEND 8X3LB.", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 10.88621688, cliente: "Econo", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0154", producto: "WWS TROPICAL FRUITS 9*500 GR", codBolsa: "MTC1642", nomBolsa: "BOL WW TROPICAL FRUIT 500G FB CL V24 REC", codCaja: "MTC1671", nomCaja: "CJ WW TROPICAL FRUITS  9*500G. FB KRAFT", bolsasXCaja: 9, cajasXPallet: 144, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TM-12-0155", producto: "CEN MARKET SUMMER BERRY BLEND 12*10 OZ", codBolsa: "MTC1412", nomBolsa: "BOL DP C. MARKET SUMMER B.BLEND 10OZ V21", codCaja: "MTCG966", nomCaja: "CJ GENERICA 390X187X143 KRAFT", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 3.4019427750000006, cliente: "Central Market", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0158", producto: "CENTRAL MARKET TROPICAL BLEND 12*10 OZ", codBolsa: "MTC1032", nomBolsa: "BOL CENT. MARKET HEB TROPICAL BLEND 10OZ", codCaja: "MTCG966", nomCaja: "CJ GENERICA 390X187X143 KRAFT", bolsasXCaja: 12, cajasXPallet: 210, kgXCaja: 3.4019427750000006, cliente: "Central Market", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0162", producto: "CHIQUITA GOURMET BERRY MEDLEY 8*12 OZ", codBolsa: "MTC1735", nomBolsa: "BOL CHIQUITA B.GOURMET 8/12OZ V25", codCaja: "MTC1149", nomCaja: "CJ CHIQUITA GOURMET BERMEDLEY 8X12OZ V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.72155422, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0163", producto: "FRUIT BLEN GOURM CHIQ 8*16 OZ CJ/CH-B/CH", codBolsa: "MTC1238", nomBolsa: "BOLS 100%PE CHIQUITA FRUIT BLEND  8/16OZ", codCaja: "MTC1151", nomCaja: "CJ CHIQUITA GOURMET FRUITBLEN 8X16OZ V19", bolsasXCaja: 8, cajasXPallet: 204, kgXCaja: 3.62873896, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0169", producto: "GLOBAL GROWERS MIX BERRIES 12*280 GR", codBolsa: "MTC1122", nomBolsa: "BOL GLOBAL GROWERS 3 BERRY BLEND 280G", codCaja: "MTC1126", nomCaja: "CJ GLOBAL GROWERS 3 BERRY MEDLEY12X280G", bolsasXCaja: 12, cajasXPallet: 198, kgXCaja: 3.36, cliente: "Global Growers", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0170", producto: "QUALITA BERRY MIX 12*400 GR", codBolsa: "MTC1138", nomBolsa: "BOL DPACK QUALITA MIX DE FRUTAS 400G V18", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0172", producto: "MIXED BERRIES IQF WWS 10*1 KG. C/WW-E", codBolsa: "MTC1199", nomBolsa: "BOLS WW MIX BERRIES 1 KILO 3 SELLO", codCaja: "MTC1375", nomCaja: "CJ WW MIXED BERRIES 10 X 1 KL. C30KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0173", producto: "PUBLIX MIX BERRIES 8*12 OZ", codBolsa: "MTC1096", nomBolsa: "BOL DPACK PUBLIX MIXED BERRIES 12OZ V18", codCaja: "MTC1157", nomCaja: "CJ PUBLIX MIXED BERRIES 8X12OZ. V19", bolsasXCaja: 8, cajasXPallet: 260, kgXCaja: 2.72155422, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0176", producto: "MIX BERR IQF WWS 9*500 GR.FLAT B.BAGS", codBolsa: "MTC1279", nomBolsa: "BOL WW MIXED BERRIES 500G FLATBOTTOM V20", codCaja: "MTC1250", nomCaja: "CJ WW M.BERRIES 9*500G FLATBOTTOM KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0179", producto: "MIXED FRUIT PUBLIX 8*20 OZ. C/PX -B/PX", codBolsa: "MTC1097", nomBolsa: "BOL DOYPACK PUBLIX MIXED FRUIT 20OZ. V18", codCaja: "MTC1177", nomCaja: "CJ PUBLIX MIXED FRUIT 8X20OZ. V19", bolsasXCaja: 8, cajasXPallet: 156, kgXCaja: 4.535923700000001, cliente: "Publix", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0182", producto: "GV FRUIT SALAD 6*48 OZ", codBolsa: "MTC1302", nomBolsa: "BOL 3 SELLO GV FRUIT SALAD 48OZ", codCaja: "MTC1318", nomCaja: "CJ GV FRUIT SALAD  6X48OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 8.16466266, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0183", producto: "GV BERRY MEDLEY 10*16 OZ", codBolsa: "MTC1292", nomBolsa: "BOL DOYPACK GV BERRY MEDLEY 16OZ", codCaja: "MTC1407", nomCaja: "CJ  GV BERRY MEDLEY  10X16OZ  SRP NEW CB", bolsasXCaja: 10, cajasXPallet: 110, kgXCaja: 4.535923700000001, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0183-CL", producto: "GV BERRY MEDLEY 10*16 OZ", codBolsa: "MTC1724", nomBolsa: "BOL DOYPACK GV BERRY MEDLEY 16OZ  CL V25", codCaja: "MTC1407", nomCaja: "CJ  GV BERRY MEDLEY  10X16OZ  SRP NEW CB", bolsasXCaja: 10, cajasXPallet: 55, kgXCaja: 4.535923700000001, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0184", producto: "GV MIX FRUIT 5*16 OZ", codBolsa: "MTC1297", nomBolsa: "BOL DOYPACK GV MIXED FRUIT 16OZ", codCaja: "MTC1328", nomCaja: "CJ  GV MIXED FRUIT   5X16OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 220, kgXCaja: 2.2679618500000003, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0184-CL", producto: "GV MIX FRUIT 5*16 OZ", codBolsa: "MTC1727", nomBolsa: "BOL DOYPACK GV MIXED FRUIT 16OZ  CL V25", codCaja: "MTC1328", nomCaja: "CJ  GV MIXED FRUIT   5X16OZ.  SRP", bolsasXCaja: 5, cajasXPallet: 110, kgXCaja: 2.2679618500000003, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0185", producto: "GV TRIPLE BERRY 6*48 OZ", codBolsa: "MTC1305", nomBolsa: "BOL 3 SELLO GV TRIPLE BERRY 48OZ", codCaja: "MTC1325", nomCaja: "CJ  GV TRIPLE BERRY  6X48OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 8.16466266, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0186", producto: "GV MIX FRUIT 6*64 OZ", codBolsa: "MTC1306", nomBolsa: "BOL 3 SELLO GV MIXED FRUIT 64OZ", codCaja: "MTC1319", nomCaja: "CJ GV MIXED FRUIT  6X64OZ.  NON-SRP", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0187", producto: "SEIYU BERRY MIX 20*200 GR", codBolsa: "MTC1640", nomBolsa: "BOL SEIYU 4 BERRYMIX 200 GR V24.1", codCaja: "MTC1624", nomCaja: "CJ SEIYU 4 BERRY MIX 20X200 GR V24", bolsasXCaja: 20, cajasXPallet: 180, kgXCaja: 4, cliente: "Seiyu", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0188", producto: "BERRY MIX SEIYU 20*200 GR. C/SY-B/SYU", codBolsa: "MTC1623", nomBolsa: "BOL SEIYU (MET) 4 BERRYMIX 200 GR V24", codCaja: "MTC1624", nomCaja: "CJ SEIYU 4 BERRY MIX 20X200 GR V24", bolsasXCaja: 20, cajasXPallet: 180, kgXCaja: 4, cliente: "Seiyu", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0189", producto: "602706-SEG TROPICAL BLEND 12x16oz.", codBolsa: "MTC1401", nomBolsa: "BOL DPACK TROPICAL BLEND SE GROCERS 16OZ", codCaja: "MTC1403", nomCaja: "CJ TROPICAL BLEND &#160;SE GROCERS 12X16", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0190", producto: "SIAM MAKRO MIX BERRIES 15*500 GR", codBolsa: "MTC1393", nomBolsa: "BOL 100%PE DP DAILY FFROZEN MIXEDB 500GR", codCaja: "MTC1398", nomCaja: "CJ DAILY FRESH FROZEN MIXEDB 15X500GRS", bolsasXCaja: 15, cajasXPallet: 100, kgXCaja: 7.5, cliente: "Makro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0191", producto: "602709-SEG BRY MEDLEY 12x16oz.", codBolsa: "MTC1349", nomBolsa: "BOL DPACK BERRY MEDLEY SE GROCERS 16 OZ", codCaja: "MTC1365", nomCaja: "CJ BERRY MEDLEY SE GROCERS 12X16OZ", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0192", producto: "602710-SEG MIXED FRUIT 12x16oz.", codBolsa: "MTC1350", nomBolsa: "BOL DOYPACK MIXED FRUIT SE GROCERS 16 OZ", codCaja: "MTC1366", nomCaja: "CJ MIXED FRUIT SE GROCERS 12X16OZ", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0193", producto: "602717-SEG MIXED FRUIT 6x48oz.", codBolsa: "MTC1355", nomBolsa: "BOL DOYPACK MIXED FRUIT SE GROCERS 48 OZ", codCaja: "MTC1371", nomCaja: "CJ MIXED FRUIT SE GROCERS 6X48OZ", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0194", producto: "602716-SEG MIX BERRIES 6x48oz.", codBolsa: "MTC1354", nomBolsa: "BOL DOYPACK MIX BERRIES SE GROCERS 48 OZ", codCaja: "MTC1370", nomCaja: "CJ BERRY MEDLEY SE GROCERS 6X48OZ", bolsasXCaja: 6, cajasXPallet: 100, kgXCaja: 8.16466266, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0196", producto: "TROPICAL MIX KOBE BUSAN 24*500 GR", codBolsa: "MTC1426", nomBolsa: "FILM KOBE BUSSAN TROPICAL MIX 500GRS", codCaja: "MTC1427", nomCaja: "CJ KOBE BUSSAN TROPICAL MIX 24X500GRS", bolsasXCaja: 24, cajasXPallet: 90, kgXCaja: 12, cliente: "Kobbe Busan", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0197", producto: "GV CHERRY BERRY BLEND 6*48 OZ", codBolsa: "MTC1301", nomBolsa: "BOL 3 SELLO GV CHERRY BERRY 48OZ", codCaja: "MTC1327", nomCaja: "CJ  GV CHERRY BERRY 6X48OZ.  NON-SRP", bolsasXCaja: 6, cajasXPallet: 90, kgXCaja: 8.16466266, cliente: "Walmart", tipoPallet: "Yugo", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TM-12-0198", producto: "BOGOPA ANTIOX FRUIT BLEND 6*48 OZ", codBolsa: "MTC1418", nomBolsa: "BOL DPACK BOGOPA ANTIOXIDANT BLEND 3 LB", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0199", producto: "BOGOPA MIX BERRIES 6*48 OZ", codBolsa: "MTC1419", nomBolsa: "BOL DOYPACK BOGOPA BERRY MIX 3 LB", codCaja: "MTC1437", nomCaja: "CJ BOGOPA GENERICA 392X292X200", bolsasXCaja: 6, cajasXPallet: 110, kgXCaja: 8.16466266, cliente: "Bogopa", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0201", producto: "GARDEN FOODS 4 BERRY MEDLEY 12*1 LB", codBolsa: "MTC1223", nomBolsa: "BOL DPACK GARDEN FOODS 4 BER.MEDLEY 454G", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Garden Food", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0202", producto: "GARDEN FOODS MANGO MEDLEY 12*1 LB", codBolsa: "MTC1221", nomBolsa: "BOL DOYPACK FOODS MANGO MEDLEY 454G", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Garden Food", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0204", producto: "GARDEN FOODS 4 BERRY MEDLEY 8*3 LB", codBolsa: "MTC1586", nomBolsa: "BOL GARDEN FOODS 4 BER.MEDLEY 48OZ", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0205", producto: "WWS MIX BERRIES 10*1 KG", codBolsa: "MTC1673", nomBolsa: "BOL WW MIX BERRIES 1 KG CL V25 REC", codCaja: "MTC1375", nomCaja: "CJ WW MIXED BERRIES 10 X 1 KL. C30KRAFT", bolsasXCaja: 10, cajasXPallet: 90, kgXCaja: 10, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n" },
  { sku: "TM-12-0206", producto: "GARDEN FOODS STRAW & BANANA 12*1 LB", codBolsa: "MTC1444", nomBolsa: "BOL DOYPACK GARDEN  F STRAW&BANANA 454G", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Garden Food", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0207", producto: "GARDEN FOODS GOURMET MEDLEY 12*1 LB", codBolsa: "MTC1445", nomBolsa: "BOL DOYPACK GARDEN F GOURMET MEDLEY 454G", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Garden Food", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0214", producto: "WWS MIX BERRIES 9*500 GR M.O", codBolsa: "MTC1674", nomBolsa: "BOL WW MIX BERRIES  500G. FB CL V25 REC", codCaja: "MTC1250", nomCaja: "CJ WW M.BERRIES 9*500G FLATBOTTOM KRAFT", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.5, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TM-12-0215", producto: "EMART TRIPLE BERRY 8*1,5 KG", codBolsa: "MTC1733", nomBolsa: "BOL EMART TRIPLE BERRIES 1,5 KL. V25", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: 8, cajasXPallet: 80, kgXCaja: 12, cliente: "Emart", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0216", producto: "MIXED BERRIES EMART 6x2,0 KG", codBolsa: "MTC1494", nomBolsa: "BOL DOY TRIPLE BERRY (EMART) 2 KL", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0221", producto: "TRADERS TRIPLE BERRY 6*2 KG", codBolsa: "MTC1732", nomBolsa: "BOL EMART TRIPLE BERRIES 2 KL. V25", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0230", producto: "PRICESMART TRIPLE BERRY 5x(4*16 OZ)", codBolsa: "MTC1695", nomBolsa: "BOL MEMBERS TRIPLE BBLEND(4-16OZ) CL V24", codCaja: "MTC1705", nomCaja: "CJ MEMBE TRIPLE BBLEND 5X64OZV24", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0251", producto: "KOBE BUSSAN MIX CHERRY-BERRIES 24*500 GR", codBolsa: "MTC1468", nomBolsa: "FILM KBUSSAN MIXED BERRIES-CHERRIES 500G", codCaja: "MTC1469", nomCaja: "CJ KBUSSAN M-BERRIES-CHERRIES 24X500G", bolsasXCaja: 12, cajasXPallet: 90, kgXCaja: 5.44310844, cliente: "Kobbe Busan", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0301", producto: "HEB BERRY MEDLEY 12*16 OZ", codBolsa: "MTC1454", nomBolsa: "BOL HEB BERRY MEDLEY 16OZ 2 SELLO", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "HEB", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TM-12-0310", producto: "ARCTIC HARVEST MIX BERRIES 8*1,5 KG", codBolsa: "MTC1555", nomBolsa: "BOL ARTIC HARVEST 3 BERRY MEDLEY 1.5K", codCaja: "MTC1557", nomCaja: "CJ ARTIC HARVEST BERRY MEDLEY 8X1,5K", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0352", producto: "WWS MIX BERRIES 9*500 GR M.O", codBolsa: "MTC1674", nomBolsa: "BOL WW MIX BERRIES  500G. FB CL V25 REC", codCaja: "MTC1250", nomCaja: "CJ WW M.BERRIES 9*500G FLATBOTTOM KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0353", producto: "WWS MIX BERRIES 10*1 KG", codBolsa: "MTC1673", nomBolsa: "BOL WW MIX BERRIES 1 KG CL V25 REC", codCaja: "MTC1375", nomCaja: "CJ WW MIXED BERRIES 10 X 1 KL. C30KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0370", producto: "KNOW & LOVE BERRY MEDLEY 12*16 OZ", codBolsa: "MTC1596", nomBolsa: "BOL KNOW&LOVE BERRY MEDLEY 16OZ V23", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0371", producto: "KNOW & LOVE BERRY MEDLEY 6*48 OZ", codBolsa: "MTC1597", nomBolsa: "BOL KNOW&LOVE BERRY MEDLEY 48OZ V23", codCaja: "MTC1280", nomCaja: "CJ GRANEL ARANDANO 392X292X200 C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0374", producto: "KNOW & LOVE TROPICAL BLEND 12*16 OZ", codBolsa: "MTC1599", nomBolsa: "BOL KNOW&LOVE TROPICAL BLEND 16OZ V23", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0401", producto: "STORE BRAND BERRY MEDLEY 8*12 OZ", codBolsa: "MTC1652", nomBolsa: "BOL GIANT BERRY MEDLEY 12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0402", producto: "STORE BRAND BERRY MEDLEY 4*48 OZ", codBolsa: "MTC1666", nomBolsa: "BOL GIANT BERRY MEDLEY 48OZ V24", codCaja: "MTC1703", nomCaja: "CJ GENERICA 392X292X140 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0403", producto: "STORE BRAND BERRY BLEND 6*64 OZ", codBolsa: "MTC1669", nomBolsa: "BOL GIANT BERRY BLEND 64OZ V24", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0406", producto: "STORE BRAND ANTIOX BLEND 8*12 OZ", codBolsa: "MTC1651", nomBolsa: "BOL GIANT ANTIOXDENT FRUIT B.  12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0407", producto: "STORE BRAND CHERRY BERRY 8*12 OZ", codBolsa: "MTC1647", nomBolsa: "BOL GIANT CHERRY BERRY MEDLEY 12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-0408", producto: "STORE BRAND BITE SZD BLEND 8*12 OZ", codBolsa: "MTC1649", nomBolsa: "BOL GIANT BITE SIIZED BLEND 12OZ V24", codCaja: "MTC1701", nomCaja: "CJ GENERICA 295X192X156 KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TM-12-7017", producto: "BOGOPA TROPICAL BLEND 6*48 OZ", codBolsa: "MTC1422", nomBolsa: "BOL DOYPACK BOGOPA TROPICAL BLEND &#160;", codCaja: "MTC1438", nomCaja: "CJ GENERICA 392X292X170  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-02-0010", producto: "WWS MACRO MANGO 9*450 GR", codBolsa: "MTC1611", nomBolsa: "BOL MACRO ORG MANGO 450G FB CL V23 REC", codCaja: "MTC1540", nomCaja: "CJ MACRO ORG. MANGO 9X450G KRAFT K-V23", bolsasXCaja: 9, cajasXPallet: 182, kgXCaja: 4.05, cliente: "Macro", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n1 Slip Sheet Negro entre la primer y el segundo piso  \n1 Slip Sheet Negro entre el septimo y octavo piso" },
  { sku: "TQ-02-0350", producto: "MAN CHUNKS ORG MACRO () 9*450GRS.", codBolsa: "MTC1505", nomBolsa: "BOL MACRO ORG MANGO 450G FB V23", codCaja: "MTC1540", nomCaja: "CJ MACRO ORG. MANGO 9X450G KRAFT K-V23", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0009", producto: "MAN CHUNK GOYA 12*16 oz.", codBolsa: "MTC1434", nomBolsa: "BOL DOYPACK GOYA MANGO CHUNKS 16OZ. V21", codCaja: "MTCG604", nomCaja: "CJ GENERICA 390X215X175  KRAFT", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "Goya", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TQ-12-0011", producto: "MAN IQF CHUNKS PUBLIX 8*16Oz C/PX-B/PX", codBolsa: "MTC1095", nomBolsa: "BOL DOY PUBLIX MANGO CHUNKS 16OZ V18", codCaja: "MTC1156", nomCaja: "CJ PUBLIX MANGO CHUNKS 8X16OZ. V19", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0013", producto: "MAN IQF WWS 9*500GRS. FLAT BOTTON BAGS", codBolsa: "MTC1180", nomBolsa: "BOL WW MANGO 500G. FLAT-BOTTOM", codCaja: "MTC1187", nomCaja: "CJ WW MANGO 9*500G. FLAT-BOTTOM", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0014", producto: "MAN IQF CHUNKS HEB 12*12 OZ C/HB - B/-HB", codBolsa: "", nomBolsa: "", codCaja: "MTC1246", nomCaja: "CJ GENERICA 390X187X153 12*340 GRS", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0015", producto: "MAN IQF CHUNKS CHIQ 8*16 OZ. C/CHQ-B/CHQ", codBolsa: "MTC1249", nomBolsa: "BOLS 100%PE CHIQUITA MANGO CHUNKS 16 OZ", codCaja: "MTC1254", nomCaja: "CJ CHIQUITA MANGO CHUNKS 8X16OZ. V20", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0017", producto: "MAN IQF CHUNKS CHIQ 8*2.5 LB C/CHQ-B/CHQ", codBolsa: "MTC1248", nomBolsa: "BOLS 100%PE CHIQUITA MANGO CHUNKS 2,5 LB", codCaja: "MTC1255", nomCaja: "CJ CHIQUITA MANGO CHUNKS  8X2,5LB  V20", bolsasXCaja: 8, cajasXPallet: 90, kgXCaja: 9.071847400000001, cliente: "Chiquita", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TQ-12-0028", producto: "602712-SEG MANGO  12x16oz.", codBolsa: "MTC1348", nomBolsa: "BOL DPACK MANGO CHUNKS SE GROCERS 16 OZ", codCaja: "MTC1364", nomCaja: "CJ MANGO CHUNKS SE GROCERS 12X16OZ", bolsasXCaja: 12, cajasXPallet: 156, kgXCaja: 5.44310844, cliente: "SEG", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet" },
  { sku: "TQ-12-0029", producto: "BOGOPA MANGO 6*48 OZ", codBolsa: "MTC1421", nomBolsa: "BOL DOYPACK BOGOPA CHUNK MAN &#160;3L", codCaja: "MTC1438", nomCaja: "CJ GENERICA 392X292X170  KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0031", producto: "MAN IQF GVALUE 5*16 OZ.C/-B/", codBolsa: "MTC1296", nomBolsa: "BOL DOYPACK GV MANGO CHUNK 16 OZ", codCaja: "MTC1333", nomCaja: "CJ GV MANGO CHUNKS 5X16OZ. SRP", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0031-CL", producto: "GV MANGO 5*16 OZ", codBolsa: "MTC1296", nomBolsa: "BOL DOYPACK GV MANGO CHUNK 16 OZ", codCaja: "MTC1333", nomCaja: "CJ GV MANGO CHUNKS 5X16OZ. SRP", bolsasXCaja: 5, cajasXPallet: 110, kgXCaja: 2.2679618500000003, cliente: "Walmart CL", tipoPallet: "Estandar", tixhi: "", slipSheet: "Sin Slip Sheet" },
  { sku: "TQ-12-0350", producto: "WWS MANGO 9*500 GR", codBolsa: "MTC1180", nomBolsa: "BOL WW MANGO 500G. FLAT-BOTTOM", codCaja: "MTC1187", nomCaja: "CJ WW MANGO 9*500G. FLAT-BOTTOM", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0351", producto: "WWS MANGO 10*1 KG", codBolsa: "MTC1619", nomBolsa: "BOL WW MANGO 1K PACKED CHILE MO V24", codCaja: "MTC1202", nomCaja: "CJ WW MANGO 10 X 1 KL", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TQ-12-0352", producto: "WWS MANGO 10*1 KG VN", codBolsa: "MTC1619", nomBolsa: "BOL WW MANGO 1K PACKED CHILE MO V24", codCaja: "MTC1374", nomCaja: "CJ WW MANGO 10 X 1 KL C30KRAFT", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TU-12-0001", producto: "UVA IQF GENERICO 1*30 LBS", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TU-12-0011", producto: "UVA IQF GENERICO 1*30 LBS SE", codBolsa: "", nomBolsa: "", codCaja: "MTC1310", nomCaja: "CJ GRANEL NEW C32KR", bolsasXCaja: "", cajasXPallet: "", kgXCaja: "", cliente: "", tipoPallet: "", tixhi: "", slipSheet: "" },
  { sku: "TA-12-0351", producto: "WWS SLCD STRAWBERRIES 10*1 KG", codBolsa: "MTC1742", nomBolsa: "BOL WW STRAWBERRIES 1 KG CL V25 REC", codCaja: "MTC1770", nomCaja: "CJ WW STRAWBERRIES  10 X 1 KL. C30KRAFT", bolsasXCaja: 10, cajasXPallet: 90, kgXCaja: 10, cliente: "Woolworth", tipoPallet: "Estandar", tixhi: "", slipSheet: "1 Slip Sheet de Carton en la Base del Pallet\n" },
];

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

function skuMaterial(sku) {
  return SKU_MATERIALES.find((s) => s.sku === sku) || null;
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
function SkuPicker({ value, onChange, sugeridos, label = "Código SKU", listId = "sku-options-global" }) {
  const [texto, setTexto] = useState(value || "");
  useEffect(() => { setTexto(value || ""); }, [value]);

  const matActual = value ? skuMaterial(value) : null;

  const commit = (raw) => {
    const code = (raw || "").trim();
    if (!code) { onChange(""); return; }
    const exact = SKU_MATERIALES.find((s) => s.sku.toLowerCase() === code.toLowerCase());
    if (exact) { onChange(exact.sku); setTexto(exact.sku); return; }
    // El datalist en algunos navegadores antepone "código — producto"; intenta extraer el código
    const soloCodigo = code.split(/[—\-–]\s/)[0].trim();
    const porPrefijo = SKU_MATERIALES.find((s) => s.sku.toLowerCase() === soloCodigo.toLowerCase());
    if (porPrefijo) { onChange(porPrefijo.sku); setTexto(porPrefijo.sku); return; }
    onChange("");
  };

  return (
    <div>
      {sugeridos && sugeridos.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mb-2">
          {sugeridos.map((sku) => {
            const mat = skuMaterial(sku);
            const sel = value === sku;
            return (
              <button key={sku} type="button" onClick={() => onChange(sku)}
                className={`text-xs rounded-full px-3 py-1 border font-medium transition-colors ${sel ? "bg-amber-500 text-white border-amber-500" : "bg-amber-50 text-amber-800 border-amber-300 hover:bg-amber-100"}`}>
                {sku}{mat ? ` · ${mat.producto}` : ""}
              </button>
            );
          })}
        </div>
      )}
      <label className="block text-xs font-medium text-slate-600 mb-1">{label}</label>
      <input
        list={listId}
        className={inputBase}
        placeholder="Escribe o pega el código (ej: T5-12-0351)"
        value={texto}
        onChange={(e) => setTexto(e.target.value)}
        onBlur={(e) => commit(e.target.value)}
      />
      <datalist id={listId}>
        {SKU_MATERIALES.map((s) => (
          <option key={s.sku} value={s.sku}>{s.producto}</option>
        ))}
      </datalist>
      {value && matActual && (
        <p className="text-xs text-emerald-700 mt-1">✓ {matActual.producto}</p>
      )}
      {value && !matActual && (
        <p className="text-xs text-red-600 mt-1">⚠ Código no encontrado en el maestro de SKU.</p>
      )}
      {!value && texto && (
        <p className="text-xs text-slate-400 mt-1">Escribe el código exacto o elígelo de la lista.</p>
      )}
    </div>
  );
}


const CATEGORIAS_COMENTARIO = ["Producción", "Calidad", "Mantención", "Seguridad"];

function buildResumenWhatsapp(area, record, inicio, programaEntries) {
  const ind = area.indicator;
  const valorInd = ind.compute(record);
  const indTexto = ind.format === "pct" ? fmtPct(valorInd) : String(valorInd);
  const lineas = [
    `*Bitácora de Turnos – ${area.title}*`,
    `Cierre de turno`,
    `Fecha: ${fmtFecha(record.fecha)} · Turno: ${turnoLabel(record.turno)}`,
    `Responsable: ${record.responsable || "—"}`,
    "",
    `${ind.label}: ${indTexto}`,
    "",
  ];
  if (area.resumenCompletoCierre) {
    const detalle = area.resumenCompletoCierre(record, inicio, programaEntries);
    if (detalle.length > 0) {
      lineas.push("*Detalle del turno:*");
      detalle.forEach(([label, value]) => lineas.push(`• ${label}: ${value ?? "—"}`));
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
  const detalle = area.resumenCompletoInicio ? area.resumenCompletoInicio(record) : area.resumenInicio(record);
  detalle.forEach(([label, value]) => {
    lineas.push(`• ${label}: ${value ?? "—"}`);
  });
  return lineas.join("\n");
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
// INSUMOS Y CONSUMO (pantalla)
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Filas de insumos (módulo nivel — no anidar dentro de otro componente)
// ---------------------------------------------------------------------------
function RowInsumo({ item, kgTotal, stockMap }) {
  const formatos  = kgTotal > 0 ? calcFormatos(kgTotal, item) : 0;
  const solicitar = Math.ceil(formatos - 1e-9);          // unidades enteras a pedir a bodega
  const tengo     = stockMap ? (stockMap[item.nombre] ?? null) : null;
  const diff      = tengo !== null ? tengo - formatos : null;
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
          {tengo !== null ? (
            <>
              <div><span className="text-slate-500">Tengo </span><b>{fmtNum(tengo, 2)}</b></div>
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
    </div>
  );
}

function RowFijo({ item, nTurnos, stockMap }) {
  const necesito  = item.cantXTurno * nTurnos;
  const solicitar = Math.ceil(necesito - 1e-9);
  const fmt       = item.formato;
  const tengo     = stockMap ? (stockMap[item.nombre] ?? null) : null;
  const diff      = tengo !== null ? tengo - necesito : null;
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
          {tengo !== null ? (
            <>
              <div><span className="text-slate-500">Tengo </span><b>{typeof tengo === "number" ? fmtNum(tengo, 2) : tengo}</b></div>
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
function SeccionFijos({ areaKey, stockMap, titulo, insumosFijo, nTurnos }) {
  const items = insumosFijo[areaKey] || [];
  if (!items.length) return null;
  return (
    <Card title={`${titulo} — Fijos por turno`}>
      {nTurnos === 0
        ? <EmptyNote text="Selecciona al menos un turno." />
        : items.map((item) => (
            <RowFijo key={item.nombre} item={item} nTurnos={nTurnos} stockMap={stockMap} />
          ))}
    </Card>
  );
}

function InsumosConsumoScreen({ isJefe, onBack, areaFiltro }) {
  const [programas]        = useSharedList("programa-records");
  const [lavadoCierres]    = useSharedList("lavado-cierre-records");
  const [seleccionCierres] = useSharedList("seleccion-cierre-records");
  const [envasadoCierres]  = useSharedList("envasado-cierre-records");
  const { config: insumosConfig } = useInsumosConfig();
  const INSUMOS_VARIABLE = insumosConfig.variable;
  const INSUMOS_FIJO     = insumosConfig.fijo;

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

  // Fechas que tienen programación relevante (para acceso rápido). Usa el
  // mapa dinámico de categorías (no arrays hardcodeados) para que ninguna
  // línea programada quede fuera y falten fechas en los chips.
  const fechasConProg = useMemo(() => {
    const fset = new Set();
    programas.forEach((p) => {
      if (p.especie === "LAVADO") return;
      const lk = p.lineaKey || p.linea;
      const categoria = CATEGORIA_INSUMO_POR_LINEA[lk];
      if (!categoria) return; // línea desconocida (no debería pasar, pero por seguridad)
      if (!areaFiltro || categoria === areaFiltro) fset.add(p.fecha);
    });
    return [...fset].sort();
  }, [programas, areaFiltro]);

  // ── Stock actual (último cierre de cada área) ─────────────────────────────
  const stockSeleccion = useMemo(() => {
    const last = [...seleccionCierres].sort((a, b) => b.id - a.id)[0];
    if (!last) return {};
    const s = {};
    MATERIALES_SELECCION.forEach((m, i) => { s[m] = num(last[`finMat_${i}`]); });
    // Conversiones a las unidades de solicitud (para que RowInsumo/RowFijo
    // pueda cruzar "Tengo" contra lo que hay en stock real del piso de planta).
    s["Caja de Film Máquina"]        = s["Film Máquina"]           ?? null;
    s["Caja de Film Manual"]         = s["Film Manual"]            ?? null;
    s["Caja de Fixo Azul"]           = s["Fixo Azul"]              ?? null;
    s["Caja de Fixo Café"]           = s["Fixo Café"]              ?? null;
    s["Caja de Fixo Transparente"]   = s["Fixo Transparente"]      ?? null;
    s["Paquete Bolsas Bins Transparentes"] = s["Bolsas Bins"] != null ? s["Bolsas Bins"] / 30 : null;
    s["Pallet de bolsas 899"]        = s["Bolsas 899"] != null ? s["Bolsas 899"] / 12000 : null;
    s["Pallet de bolsas 1744"]       = s["Bolsas 1744 (Totes)"] != null ? s["Bolsas 1744 (Totes)"] / 12000 : null;
    s["Pallet de cajas (MTC)"]       = ((s["Pallet MTC1280"] || 0) + (s["Pallet MTC1310"] || 0)) || null;
    return s;
  }, [seleccionCierres]);

  const stockEnvasado = useMemo(() => {
    const last = [...envasadoCierres].sort((a, b) => b.id - a.id)[0];
    if (!last) return {};
    const s = {};
    MATERIALES_ENVASADO.forEach((m, i) => { s[m] = num(last[`finMat_${i}`]); });
    s["Caja de Film Máquina"] = s["Film Máquina"] ?? null;
    s["Caja de Film Manual"]  = s["Film Manual"]  ?? null;
    s["Caja de Fixo Azul"]    = s["Fixo Azul"]    ?? null;
    return s;
  }, [envasadoCierres]);

  // Pallets de bandejas pendientes según el último cierre de Lavado registrado
  // (Lavado no se programa con cantidad en el Programa de producción, así que
  // se usa el backlog real reportado al cierre como referencia de "Necesito").
  const palletsBandejasLavado = useMemo(() => {
    const last = [...lavadoCierres].sort((a, b) => b.id - a.id)[0];
    if (!last) return 0;
    return totalPalletsManual(last, "pendientes");
  }, [lavadoCierres]);

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
            const mat = SKU_MATERIALES.find((s) => s.sku === p.especie);
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
      const mat = SKU_MATERIALES.find((s) => s.sku === sku);
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
                      <RowInsumo key={item.nombre} item={item} kgTotal={desglose.palletsSeleccion} stockMap={stockSeleccion} />
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

            <SeccionFijos areaKey="seleccion" stockMap={stockSeleccion} titulo="Selección" insumosFijo={INSUMOS_FIJO} nTurnos={nTurnos} />

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
                <RowInsumo key={item.nombre} item={item} kgTotal={palletsBandejasLavado} stockMap={{}} />
              ))}
            </Card>
            <SeccionFijos areaKey="lavado" stockMap={{}} titulo="Lavado de bandejas" insumosFijo={INSUMOS_FIJO} nTurnos={nTurnos} />
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
                    <RowInsumo key={item.nombre} item={item} kgTotal={desglose.palletsEnvasadoTotal} stockMap={stockEnvasado} />
                  ))}
                </>
              )}
            </Card>

            <SeccionFijos areaKey="envasado" stockMap={stockEnvasado} titulo="Envasado" insumosFijo={INSUMOS_FIJO} nTurnos={nTurnos} />

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
  return { nombre: "", formato: "Caja", uniXFormato: 1, tasaBase: 1, base: 100, unidadBase: "Pallets", nota: "" };
}
function vacioFijo() {
  return { nombre: "", formato: "Caja", uniXFormato: 1, cantXTurno: 1, nota: "", unidadMostrar: "" };
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
    arr[idx] = { ...arr[idx], [campo]: ["uniXFormato","tasaBase","base","cantXTurno"].includes(campo) ? num(valor) : valor };
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

  // Devuelve el material fusionado (base del maestro + overrides del Jefe)
  const getMat = (sku) => {
    const base = SKU_MATERIALES.find((s) => s.sku === sku);
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
function EspecificacionScreen({ onBack }) {
  const { getMat } = useSkuOverrides();
  const [skuSel, setSkuSel] = useState("");
  const mat = skuSel ? getMat(skuSel) : null;
  const cod = mat?.cliente ? CODIFICACION_POR_CLIENTE[mat.cliente] : null;

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Especificación de SKU" subtitle="Materiales, codificaciones y configuración de pallet" onBack={onBack} icon={Package} accent="amber" />
      <div className="px-4 pt-4 space-y-4">

        <Card title="Buscar SKU">
          <SkuPicker
            label="Código SKU"
            value={skuSel}
            onChange={setSkuSel}
            listId="sku-espec-consulta"
          />
        </Card>

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

        <Card title="Seleccionar SKU">
          <SkuPicker
            label="Código SKU a editar"
            value={skuSel}
            onChange={setSkuSel}
            listId="sku-espec-edit"
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
  const mapping = TIPOS_BANDEJA.map((t) => [`pendientes_${t.key}_pallets`, `sucios_${t.key}_pallets`]);
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
      <Card title="Insumos" step={3}>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <NumField label="Rollos de film" value={values.rollosFilm} onChange={(v) => setField("rollosFilm", v)} unit="Rollos" />
          <NumField label="Bolsas de bins" value={values.bolsasBins} onChange={(v) => setField("bolsasBins", v)} unit="Paquetes" />
        </div>
      </Card>
      <ComentariosFields values={values} setField={setField} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// SELECCIÓN
// ---------------------------------------------------------------------------
function MaterialesTable({ items, values, setField, prefix }) {
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
      {items.map((label, i) => (
        <NumField key={i} label={label} value={values[`${prefix}_${i}`]} onChange={(v) => setField(`${prefix}_${i}`, v)} />
      ))}
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
        {prevCierre && (
          <p className="text-xs text-emerald-700 mb-3">
            Se completó automáticamente con los materiales de piso de planta del cierre {turnoLabel(prev.turno)} del {fmtFecha(prev.fecha)}. Puedes ajustarlos si es necesario.
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
    .filter((x) => x.proceso);
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
                  <div className="text-sm font-semibold text-slate-900">{l.label}</div>
                  {especies.map(({ s, proceso }) => {
                    const ing = num(values[`kg_${l.key}_e${s}_ing`]);
                    const apr = kgAprobadoTotal(values, l.key, s);
                    const rend = ing ? apr / ing : 0;
                    return (
                      <div key={s} className="border border-slate-200 bg-slate-50 rounded-lg p-2.5 space-y-2">
                        <div className="flex justify-between text-xs">
                          {especies.length > 1 && <span className="font-semibold text-slate-500">Especie {s}</span>}
                          <span className="text-slate-700 font-medium">{proceso}</span>
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
  const mapping = MATERIALES_ENVASADO.map((_, i) => [`finMat_${i}`, `inicioMat_${i}`]);
  const { prevCierre, prev } = useCarryOver("envasado", values, setField, editingId, mapping);
  const { horarios } = useHorariosTurno();

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

  // Especie(s) programadas para Línea 5 en este turno.
  const especiesLinea5Prog = useMemo(() => {
    const set = new Set();
    programaEntries
      .filter((e) => e.lineaKey === "linea5" && e.turno === values.turno && e.especie !== "LAVADO")
      .forEach((e) => { const esp = especieFor(LINEA_PROCESOS.LINEA_MANUAL, e.especie); if (esp) set.add(esp); });
    return [...set];
  }, [programaEntries, values.turno]);

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
        {prevCierre && (
          <p className="text-xs text-emerald-700 mb-3">
            Se completó automáticamente con los materiales de piso de planta del cierre {turnoLabel(prev.turno)} del {fmtFecha(prev.fecha)}. Puedes ajustarlos si es necesario.
          </p>
        )}
        <MaterialesTable items={MATERIALES_ENVASADO} values={values} setField={setField} prefix="inicioMat" />
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
            {activaEnvasadora && (
              <div className="space-y-2">
                <SkuPicker
                  label="Código SKU"
                  value={values.envasadora_sku}
                  onChange={(v) => setField("envasadora_sku", v)}
                  sugeridos={skusProgramadosTurno}
                  listId="sku-options-inicio"
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
            {activaLinea5 && (
              <div className="space-y-2">
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
  const rendimientoLinea5 = num(values.l5_kgIngresados) ? num(values.l5_kgAprobados) / num(values.l5_kgIngresados) : 0;
  return { cumplimiento, mermaCajas, mermaBolsas, rendimientoLinea5 };
}

function EnvasadoCierre({ values, setField, inicio, programaEntries }) {
  const activaEnvasadora = (inicio?.activa_envasadora ?? values.activa_envasadora) === "Sí";
  const activaLinea5 = (inicio?.activa_linea5 ?? values.activa_linea5) === "Sí";
  const especieLinea5 = inicio?.linea5_especie || values.linea5_especie || "";
  let procesosLinea5 = LINEA_PROCESOS.LINEA_MANUAL.filter(([, especie]) => especie === especieLinea5).map((p) => p[0]);
  if (procesosLinea5.length === 0) procesosLinea5 = LINEA_PROCESOS.LINEA_MANUAL.map((p) => p[0]);
  const m = computeEnvasadoMetrics(values);
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

  return (
    <div className="space-y-4">
      <Card title="Materiales piso planta (fin de turno)" step={1}>
        <MaterialesTable items={MATERIALES_ENVASADO} values={values} setField={setField} prefix="finMat" />
      </Card>

      {!activaEnvasadora && !activaLinea5 && (
        <Card><EmptyNote text="Activa Envasadora y/o Línea 5 en el inicio de turno para ver el cierre correspondiente." /></Card>
      )}

      {activaLinea5 && (
        <Card title="Cierre de proceso — Línea 5" step={cierreStep2}>
          <p className="text-sm text-slate-700 mb-2">Especie: <span className="font-semibold text-slate-900">{especieLinea5 || "—"}</span></p>
          <div className="mb-3">
            <SelectField label="Proceso" value={values.l5_proceso} onChange={(v) => setField("l5_proceso", v)} options={procesosLinea5} />
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <NumField label="Kg ingresados" value={values.l5_kgIngresados} onChange={(v) => setField("l5_kgIngresados", v)} />
            <NumField label="Kg aprobados" value={values.l5_kgAprobados} onChange={(v) => setField("l5_kgAprobados", v)} />
          </div>
          <p className="text-sm text-slate-700 mt-2">Rendimiento: <span className="font-semibold">{fmtPct(m.rendimientoLinea5)}</span></p>
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
              listId="sku-options-cierre"
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
        </Card>

        {loading ? (
          <Card><Loader /></Card>
        ) : (
          <>
            <Card title="Vista general de la semana">
              <div className="overflow-x-auto -mx-4 px-4">
                <table className="border-collapse text-[11px] min-w-[2300px]">
                  <thead>
                    <tr>
                      <th className="sticky left-0 z-10 bg-slate-100 border border-slate-300 px-2 py-1 text-left text-slate-600 min-w-[90px]">Línea</th>
                      {slots.map((s, i) => (
                        <th key={i} className="border border-slate-300 px-2 py-1 text-center min-w-[104px]">
                          <div className="text-slate-800 font-semibold">{diaCorto(s.fecha)}</div>
                          <div className="text-slate-600">{turnoLabel(s.turno)}</div>
                          <div className="text-slate-500">{fmtFechaCorta(s.fecha)}</div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {programaLineas.map((linea) => {
                      const filaEntries = slots.map((s) => (entriesPorSlot[slotKey(s)] || []).filter((e) => e.lineaKey === linea.key));
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
              <p className="text-xs text-slate-500 mt-2">Desliza para ver los 21 turnos de la semana (Lunes a Domingo, T3/T1/T2). Las celdas en amarillo indican que esa línea está en LAVADO durante ese turno.</p>
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
                        listId="sku-options-programa"
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
  "H/ABORDO + LLAVES",
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
  ];
  let hayDiferencias = false;
  (record.items || []).forEach((it) => {
    const ini = inicialMap ? (inicialMap[it.nombre] ?? "") : "";
    const cambio = ini !== "" && it.cant !== "" && String(ini) !== String(it.cant);
    if (cambio) hayDiferencias = true;
    if (it.cant || it.obs || cambio) {
      const flag = cambio ? " ⚠" : "";
      lines.push(`• ${it.nombre}: Ini ${ini !== "" ? fmtNum(num(ini)) : "—"} → Act ${it.cant !== "" ? fmtNum(num(it.cant)) : "—"}${flag}${it.obs ? ` · ${it.obs}` : ""}`);
    }
  });
  if (hayDiferencias) lines.push("\n⚠ Hay diferencias en cantidades respecto al inicio del turno.");
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
    const tAnterior = TURNO_ANTERIOR[turnoEntrega];
    const candidatos = records.filter(
      (r) => r.turnoEntrega === tAnterior &&
             r.turnoRecibe  === turnoEntrega &&
             r.fecha >= fechaAtras(fecha, 2) &&
             r.fecha <= fecha
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
              <th className="border border-slate-300 px-3 py-2 text-left   text-xs font-semibold text-slate-600 uppercase tracking-wide w-[34%]">Artículo</th>
              <th className="border border-slate-300 px-2 py-2 text-center text-xs font-semibold text-slate-500 uppercase tracking-wide w-[12%]">Inicial</th>
              <th className="border border-slate-300 px-2 py-2 text-center text-xs font-semibold text-slate-600 uppercase tracking-wide w-[12%]">Actual</th>
              <th className="border border-slate-300 px-3 py-2 text-left   text-xs font-semibold text-slate-600 uppercase tracking-wide">Observación</th>
              <th className="border border-slate-300 px-2 py-2 text-center text-xs font-semibold text-slate-600 uppercase tracking-wide w-[12%]">Turno</th>
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
                  <td className="border border-slate-200 px-2 py-2 text-center align-middle">
                    <div className="flex flex-col items-center gap-0.5">
                      <span className="text-xs font-bold text-slate-700 leading-none">{turnoEntrega}</span>
                      <span className="text-slate-300 text-xs leading-none">↓</span>
                      <span className="text-xs font-bold text-blue-700 leading-none">{turnoRecibe}</span>
                    </div>
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
function resumenLavadoInicioCompleto(r) {
  const items = [
    ["Operarios", fmtNum(r.operarios || 0)],
    ["Movilizadores", fmtNum(r.movilizadores || 0)],
    ["Jefe de línea", fmtNum(r.jefeLinea || 0)],
    ["¿Dotación completa?", r.dotacionCompleta || "—"],
    ["Comentarios dotación", r.comentariosDotacion || "—"],
  ];
  TIPOS_BANDEJA.forEach((t) => {
    const v = num(r[`sucios_${t.key}_pallets`]);
    if (v > 0) items.push([`Sucios · ${t.label}`, `${fmtNum(v)} pallets`]);
  });
  items.push(["Total pallets sucios al iniciar", fmtNum(totalPalletsManual(r, "sucios"))]);
  return items;
}

function resumenLavadoCierreCompleto(r) {
  const items = [];
  TIPOS_BANDEJA.forEach((t) => {
    const v = palletsValue(r, "lavados", t);
    if (v > 0) items.push([`Lavados · ${t.label}`, `${fmtNum(v)} pallets`]);
  });
  items.push(["Total pallets lavados", fmtNum(totalPalletsFromPrefix(r, "lavados"))]);
  TIPOS_BANDEJA.forEach((t) => {
    const v = num(r[`pendientes_${t.key}_pallets`]);
    if (v > 0) items.push([`Pendientes · ${t.label}`, `${fmtNum(v)} pallets`]);
  });
  items.push(["Total pallets pendientes", fmtNum(totalPalletsManual(r, "pendientes"))]);
  items.push(["Rollos de film", fmtNum(r.rollosFilm || 0)]);
  items.push(["Bolsas de bins", fmtNum(r.bolsasBins || 0)]);
  return items;
}

function resumenSeleccionInicioCompleto(r) {
  const items = [];
  DOTACION_GENERAL_SELECCION.forEach((label, i) => {
    const v = num(r[`dg_${i}`]);
    if (v > 0) items.push([label, fmtNum(v)]);
  });
  items.push(["Total dotación", fmtNum(totalDotacionSeleccion(r))]);
  items.push(["¿Dotación completa?", r.dotacionCompleta || "—"]);
  MATERIALES_SELECCION.forEach((nombre, i) => {
    const v = num(r[`inicioMat_${i}`]);
    if (v > 0) items.push([`Material · ${nombre}`, fmtNum(v)]);
  });
  const activas = LINEAS_SELECCION.filter((l) => r[`linea_${l.key}_activa`] === "Sí");
  if (activas.length === 0) {
    items.push(["Líneas activas", "Ninguna"]);
  } else {
    activas.forEach((l) => {
      const procesos = ESPECIE_SLOTS.map((s) => r[`linea_${l.key}_proceso${sufijoEspecie(s)}`]).filter(Boolean);
      items.push([`${l.label} · Proceso`, procesos.length ? procesos.join(" + ") : "—"]);
      items.push([`${l.label} · Dotación`, `${fmtNum(lineaDotacionTotal(r, l))} (${DOTACION_LINEA.map((d, i) => `${d}: ${fmtNum(num(r[`linea_${l.key}_dot${i}`]))}`).join(", ")})`]);
    });
  }
  items.push(["Armado de materiales", r.armado_activa === "Sí" ? `Activo · Dot. ${fmtNum(armadoDotacionTotal(r))} (${DOTACION_LINEA.map((d, i) => `${d}: ${fmtNum(num(r[`armado_dot${i}`]))}`).join(", ")})` : "Inactivo"]);
  items.push(["Comentarios dotación", r.comentariosDotacion || "—"]);
  return items;
}

function resumenSeleccionCierreCompleto(r, inicio, programaEntries) {
  const items = [];
  MATERIALES_SELECCION.forEach((nombre, i) => {
    const v = num(r[`finMat_${i}`]);
    if (v > 0) items.push([`Material · ${nombre}`, fmtNum(v)]);
  });
  const activas = LINEAS_SELECCION.filter((l) => (inicio?.[`linea_${l.key}_activa`] ?? r[`linea_${l.key}_activa`]) === "Sí");
  if (activas.length === 0) {
    items.push(["Líneas activas", "Ninguna"]);
  } else {
    const rendTotal = computeSeleccionRendimiento({ ...(inicio || {}), ...r });
    const { cumplimiento, kgProgramado, kgIngresado } = computeSeleccionCumplimiento(r, inicio, programaEntries);
    items.push(["Rendimiento total", fmtPct(rendTotal)]);
    items.push(["Cumplimiento", kgProgramado ? `${fmtPct(cumplimiento)} (${fmtNum(kgIngresado, 0)}/${fmtNum(kgProgramado, 0)} Kg)` : "Sin Kg programado"]);
    activas.forEach((l) => {
      const especies = especiesActivasLinea(l, inicio, r);
      especies.forEach(({ s, proceso }) => {
        const ing = num(r[`kg_${l.key}_e${s}_ing`]);
        const apr = kgAprobadoTotal(r, l.key, s);
        const rend = ing ? apr / ing : 0;
        const prefijo = especies.length > 1 ? `${l.label} · ${proceso}` : l.label;
        items.push([`${prefijo} · Kg ingresados`, fmtNum(ing)]);
        APROBADO_SLOTS.forEach((t) => {
          const tipo = r[`kg_${l.key}_e${s}_apr_t${t}_tipo`];
          const kg = r[`kg_${l.key}_e${s}_apr_t${t}_kg`];
          if (tipo && kg !== undefined && kg !== "") {
            items.push([`${prefijo} · Aprobado (${tipo})`, fmtNum(num(kg))]);
          }
        });
        items.push([`${prefijo} · Kg aprobados total`, fmtNum(apr)]);
        items.push([`${prefijo} · Rendimiento`, fmtPct(rend)]);
      });
    });
  }
  return items;
}

function resumenEnvasadoInicioCompleto(r) {
  const items = [];
  DOTACION_GENERAL_ENVASADO.forEach((label, i) => {
    const v = num(r[`dg_${i}`]);
    if (v > 0) items.push([label, v]);
  });
  items.push(["¿Dotación completa?", r.dotacionCompleta || "—"]);
  MATERIALES_ENVASADO.forEach((nombre, i) => {
    const v = num(r[`inicioMat_${i}`]);
    if (v > 0) items.push([`Material · ${nombre}`, v]);
  });
  if (r.activa_envasadora === "Sí") {
    const mat = r.envasadora_sku ? skuMaterial(r.envasadora_sku) : null;
    items.push(["Envasadora · SKU", r.envasadora_sku ? `${mat?.producto || "—"} (${r.envasadora_sku})` : "—"]);
    items.push(["Envasadora · Dotación", DOTACION_LINEA.map((d, i) => `${d}: ${fmtNum(num(r[`envasadora_dot${i}`]))}`).join(", ")]);
  }
  if (r.activa_linea5 === "Sí") {
    items.push(["Línea 5 · Especie", r.linea5_especie || "—"]);
    items.push(["Línea 5 · Dotación", DOTACION_LINEA.map((d, i) => `${d}: ${fmtNum(num(r[`linea5_dot${i}`]))}`).join(", ")]);
  }
  if (r.activa_envasadora !== "Sí" && r.activa_linea5 !== "Sí") items.push(["Líneas de trabajo", "Ninguna activa"]);
  items.push(["Comentarios dotación", r.comentariosDotacion || "—"]);
  return items;
}

function resumenEnvasadoCierreCompleto(r, inicio) {
  const items = [];
  MATERIALES_ENVASADO.forEach((nombre, i) => {
    const v = num(r[`finMat_${i}`]);
    if (v > 0) items.push([`Material · ${nombre}`, v]);
  });
  const activaLinea5 = (inicio?.activa_linea5 ?? r.activa_linea5) === "Sí";
  const activaEnvasadora = (inicio?.activa_envasadora ?? r.activa_envasadora) === "Sí";
  if (activaLinea5) {
    const m = computeEnvasadoMetrics(r);
    items.push(["Línea 5 · Especie", inicio?.linea5_especie || r.linea5_especie || "—"]);
    items.push(["Línea 5 · Proceso", r.l5_proceso || "—"]);
    items.push(["Línea 5 · Kg ingresados", fmtNum(r.l5_kgIngresados || 0)]);
    items.push(["Línea 5 · Kg aprobados", fmtNum(r.l5_kgAprobados || 0)]);
    items.push(["Línea 5 · Rendimiento", fmtPct(m.rendimientoLinea5)]);
  }
  if (activaEnvasadora) {
    const m = computeEnvasadoMetrics(r);
    const mat = r.envasadora_sku ? skuMaterial(r.envasadora_sku) : null;
    items.push(["Envasadora · SKU", r.envasadora_sku ? `${mat?.producto || "—"} (${r.envasadora_sku})` : "—"]);
    items.push(["Cajas producidas", fmtNum(r.cajasProducidas || 0)]);
    items.push(["Cajas programadas", fmtNum(r.cajasProgramadas || 0)]);
    items.push(["Cajas teóricas consumo", fmtNum(r.cajasTeoricas || 0)]);
    items.push(["Cajas consumidas real", fmtNum(r.cajasConsumidasReal || 0)]);
    items.push(["Bolsas teóricas consumo", fmtNum(r.bolsasTeoricas || 0)]);
    items.push(["Bolsas consumidas real", fmtNum(r.bolsasConsumidasReal || 0)]);
    items.push(["Cumplimiento", fmtPct(m.cumplimiento)]);
    items.push(["Merma cajas", fmtPct(m.mermaCajas)]);
    items.push(["Merma bolsas", fmtPct(m.mermaBolsas)]);
  }
  return items;
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
    resumenCompletoCierre: (r) => resumenLavadoCierreCompleto(r),
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
  useEffect(() => { setLastSaved(null); }, [editingId]);

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
    const updated = editingId ? records.map((r) => (r.id === editingId ? record : r)) : [...records, record];
    const ok = await saveRecords(updated);
    if (ok) {
      setToast({ kind: "ok", message: editingId ? "Cambios guardados." : "Registro de inicio guardado." });
      setLastSaved(record);
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
  const inicio = inicioRecords.find((r) => r.claveTurno === clave);
  const cierreActual = cierreRecords.find((r) => r.claveTurno === clave);

  useEffect(() => {
    if (loading) return;
    const found = cierreRecords.find((r) => r.claveTurno === clave);
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
    const updated = cierreActual ? cierreRecords.map((r) => (r.id === cierreActual.id ? record : r)) : [...cierreRecords, record];
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

          <EntregaMaterialesCard areaKey={area.key} fecha={fecha} turno={turno} autor={autor} supervisoresList={supervisoresList} />

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
                href={whatsappShareUrl(buildResumenWhatsapp(area, cierreActual, inicio, programaEntries))}
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
  const merged = sorted.map((c) => ({ ...(inicioRecords.find((r) => r.claveTurno === c.claveTurno) || {}), ...c }));

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

  const sortedInicio = [...inicioRecords].sort((a, b) => b.id - a.id);
  const sortedCierre = [...cierreRecords].sort((a, b) => b.id - a.id);

  return (
    <div className="space-y-4">
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
      setEditingId(null);
      setValues({ fecha: today(), turno: "" });
      setTab("registros");
    } else {
      setValues((v) => ({ fecha: v.fecha, turno: "" }));
    }
  };

  const tabs = [
    { key: "inicio", label: "Inicio de turno", icon: ClipboardList },
    { key: "cierre", label: "Cierre de turno", icon: ClipboardCheck },
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
        {tab === "indicadores" && <IndicadoresTab area={area} />}
        {tab === "registros" && <RegistrosTab area={area} onEdit={startEdit} onEditCierre={startEditCierre} />}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Dashboard del Jefe de producción
// ---------------------------------------------------------------------------
function useAreaCierres(areaKey) {
  const [inicio, , loadingI] = useSharedList(`${areaKey}-inicio-records`);
  const [cierres, , loadingC] = useSharedList(`${areaKey}-cierre-records`);
  return { inicio, cierres, loading: loadingI || loadingC };
}

function MiniChart({ area, inicio, cierres }) {
  const ind = area.indicator;
  const sorted = [...cierres].sort((a, b) => a.id - b.id).slice(-7);
  const merged = sorted.map((c) => ({ ...(inicio.find((r) => r.claveTurno === c.claveTurno) || {}), ...c }));
  const data = merged.map((r) => ({
    label: `${fmtFecha(r.fecha)} ${r.turno}`,
    value: ind.format === "pct" ? Number((ind.compute(r) * 100).toFixed(1)) : ind.compute(r),
  }));
  return (
    <Card title={area.title}>
      <p className="text-xs text-slate-600 mb-2">{ind.label}</p>
      {data.length === 0 ? <EmptyNote text="Sin cierres registrados." /> : (
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={data} margin={{ top: 8, right: 8, left: -16, bottom: 0 }}>
            <CartesianGrid stroke="#e2e8f0" strokeDasharray="3 3" vertical={false} />
            <XAxis dataKey="label" stroke="#64748b" fontSize={9} />
            <YAxis stroke="#64748b" fontSize={10} unit={ind.format === "pct" ? "%" : ""} />
            <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12, color: "#1e293b" }} />
            <Bar dataKey="value" fill={ACCENT[area.accent].chart} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      )}
    </Card>
  );
}

function JefeDashboard({ onBack }) {
  const lavado = useAreaCierres("lavado");
  const seleccion = useAreaCierres("seleccion");
  const envasado = useAreaCierres("envasado");

  const todosCierres = [...lavado.cierres, ...seleccion.cierres, ...envasado.cierres];
  const incidentes = todosCierres.filter((c) => c.huboIncidentes === "Sí").length;
  const accidentes = todosCierres.filter((c) => c.huboAccidentes === "Sí").length;

  const recientes = [
    ...lavado.cierres.map((c) => ({ ...c, area: "Lavado de bandejas" })),
    ...seleccion.cierres.map((c) => ({ ...c, area: "Selección" })),
    ...envasado.cierres.map((c) => ({ ...c, area: "Envasado" })),
  ]
    .filter((c) => c.estado === "Enviado")
    .sort((a, b) => new Date(b.fechaHoraCierre || 0) - new Date(a.fechaHoraCierre || 0))
    .slice(0, 8);

  return (
    <div className="w-full max-w-2xl lg:max-w-4xl mx-auto pb-10">
      <TopBar title="Indicadores generales" subtitle="Vista consolidada de las 3 áreas" onBack={onBack} icon={BarChart3} accent="emerald" />
      <div className="p-4 space-y-4">
        <Card title="Seguridad — acumulado">
          <div className="flex flex-wrap gap-3">
            <span className="flex items-center gap-1.5 text-sm font-medium rounded-lg border px-3 py-1.5 bg-amber-500/10 text-amber-700 border-amber-500/30">
              <AlertTriangle size={16} /> Cierres con incidentes: {incidentes}
            </span>
            <span className="flex items-center gap-1.5 text-sm font-medium rounded-lg border px-3 py-1.5 bg-red-500/10 text-red-700 border-red-500/30">
              <AlertTriangle size={16} /> Cierres con accidentes: {accidentes}
            </span>
          </div>
        </Card>

        <MiniChart area={AREAS.lavado} inicio={lavado.inicio} cierres={lavado.cierres} />
        <MiniChart area={AREAS.seleccion} inicio={seleccion.inicio} cierres={seleccion.cierres} />
        <MiniChart area={AREAS.envasado} inicio={envasado.inicio} cierres={envasado.cierres} />

        <Card title="Cierres recientes">
          {recientes.length === 0 ? <EmptyNote text="Todavía no se han finalizado cierres de turno." /> : (
            <div className="space-y-2 text-sm">
              {recientes.map((c) => (
                <div key={`${c.area}-${c.id}`} className="border border-slate-300 rounded-lg px-3 py-2">
                  <div className="flex justify-between">
                    <span className="font-medium text-slate-900">{c.area}</span>
                    <span className="text-slate-600">{fmtFecha(c.fecha)} · {c.turno}</span>
                  </div>
                  <div className="text-xs text-slate-600 mt-1">
                    Responsable: {c.responsable || "—"} · Incidentes: {c.huboIncidentes || "No"} · Accidentes: {c.huboAccidentes || "No"}
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>

        <p className="text-xs text-slate-500">
          Para editar o eliminar registros de una área, entra a esa área y abre la pestaña «Registros».
        </p>
      </div>
    </div>
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
function HomeScreen({ onNavigate, isJefe, onOpenLogin, onLogoutJefe }) {
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
              { key: "dashboard",        label: "Indicadores generales",       desc: "Dashboard de las 3 áreas", Icon: BarChart3 },
              { key: "programa",         label: "Programa de producción",      desc: "Línea, turno, especie y cantidad", Icon: CalendarDays },
              { key: "supervisores",     label: "Gestionar Supervisores",      desc: "Agregar o eliminar supervisores de los portales", Icon: User },
              { key: "horarios",         label: "Horarios de Turno",           desc: "Lunes a Sábado — minutos efectivos por turno", Icon: Clock },
              { key: "especies",         label: "Gestionar especies por línea",desc: "Agregar nuevos procesos/especies", Icon: Plus },
              { key: "insumos-config",   label: "Configurar Insumos",          desc: "Tasas de consumo de Selección, Lavado y Envasado", Icon: Boxes },
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
        <InsumosConsumoScreen isJefe={isJefe} areaFiltro="seleccion" onBack={() => setScreen("portal-seleccion")} />
      )}
      {screen === "insumos-envasado" && (
        <InsumosConsumoScreen isJefe={isJefe} areaFiltro="envasado" onBack={() => setScreen("portal-envasado")} />
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
