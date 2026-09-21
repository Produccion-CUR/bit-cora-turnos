// =============================================================================
// BITÁCORA DE TURNOS — Apps Script Web App
// =============================================================================
// Guarda cada "clave" de la app en una hoja técnica llamada "Storage" (así
// la app puede volver a leer sus propios datos). ADEMÁS, para las claves que
// son "registros por turno" (Inicio/Cierre de cada área, Programa, Entrega
// de Materiales), arma una PESTAÑA APARTE con una tabla legible — una fila
// por registro, columnas con NOMBRE FÁCIL DE ENTENDER — lista para
// filtrar/ordenar/exportar a Excel (Archivo → Descargar → Microsoft Excel).
// =============================================================================

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateStorageSheet(ss) {
  let sh = ss.getSheetByName("Storage");
  if (!sh) {
    sh = ss.insertSheet("Storage");
    sh.appendRow(["key", "value", "updated_at"]);
    sh.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    sh.setColumnWidth(1, 240);
    sh.setColumnWidth(2, 600);
    sh.setColumnWidth(3, 160);
  }
  return sh;
}

// Claves que son "listas de registros por turno" → se arman como tabla legible.
// (clave que manda la app) → { tab: nombre de la pestaña, area: "seleccion"|"envasado"|null }
// "area" le dice al traductor de nombres qué lista de materiales/dotación usar,
// ya que Selección y Envasado tienen listas distintas.
const TABLAS_LEGIBLES = {
  "lavado-inicio-records":        { tab: "Lavado - Inicio",   area: null },
  "lavado-cierre-records":        { tab: "Lavado - Cierre",   area: null },
  "seleccion-inicio-records":     { tab: "Selección - Inicio", area: "seleccion" },
  "seleccion-cierre-records":     { tab: "Selección - Cierre", area: "seleccion" },
  "envasado-inicio-records":      { tab: "Envasado - Inicio",  area: "envasado" },
  "envasado-cierre-records":      { tab: "Envasado - Cierre",  area: "envasado" },
  "programa-records":             { tab: "Programa de Producción", area: null },
  "programa-observaciones":       { tab: "Observaciones del Programa", area: null },
  "entrega-materiales-seleccion": { tab: "Entrega Materiales - Selección", area: null },
  "entrega-materiales-envasado":  { tab: "Entrega Materiales - Envasado", area: null },
};

const ORDEN_TURNO = { T1: 1, T2: 2, T3: 3 };

// -----------------------------------------------------------------------
// Diccionario de nombres amigables (mismo vocabulario que usa la app)
// -----------------------------------------------------------------------
const MATERIALES_SELECCION = [
  "Pallet Taco Normal", "Pallet MTC1280", "Pallet MTC1310", "Bolsas 899",
  "Bolsas 1744 (Totes)", "Bolsas Bins", "Fixo Azul", "Fixo Transparente",
  "Fixo Café", "Film Máquina", "Film Manual", "Rollo Cotona",
  "Pallet Tote Armados", "Pallet Cajas Armados 1310", "Pallet Cajas Armados 1280", "Pallet Totes x Armar",
];
const MATERIALES_ENVASADO = ["Pallet Certificado", "Film Máquina", "Film Manual", "MTC Cajas", "MTC Bolsas"];
const DOTACION_GENERAL_SELECCION = [
  "Estadístico", "Operador de Máquina", "Jefe Línea Cámara",
  "Movilizadores Cámara", "Supervisor Gestión", "Supervisor Producción",
];
const DOTACION_LINEA = ["Operarios", "Movilizadores", "Jefe de Línea"];
const DOTACION_GENERAL_ENVASADO = ["Estadístico", "Operador de Máquina", "Supervisor Producción"];
const LINEAS_LABEL = { linea1: "Línea 1", linea3: "Línea 3", linea4: "Línea 4", linea5: "Línea 5", linea6: "Línea 6" };

// Campos fijos (mismo nombre sin importar el área) — coincidencia exacta.
const CAMPOS_FIJOS = {
  id: "ID",
  fecha: "Fecha",
  turno: "Turno",
  claveTurno: "Clave de Turno",
  responsable: "Responsable",
  horaInicio: "Hora de Inicio",
  dotacionCompleta: "¿Dotación Completa?",
  comentariosDotacion: "Comentarios de Dotación",
  lineaKey: "Línea",
  lineaLabel: "Línea",
  especie: "Proceso / SKU Programado",
  cantidad: "Cantidad Programada",
  unidad: "Unidad",
  turnoEntrega: "Turno que Entrega",
  turnoRecibe: "Turno que Recibe",
  entrega: "Quien Entrega",
  recibe: "Quien Recibe",
  items: "Detalle de Artículos (JSON)",
  texto: "Observación",
  operarios: "Operarios",
  movilizadores: "Movilizadores",
  jefeLinea: "Jefe de Línea",
  rollosFilm: "Rollos de Film",
  bolsasBins: "Bolsas de Bins",
  activa_envasadora: "¿Envasadora Activa?",
  activa_linea5: "¿Línea 5 Activa?",
  envasadora_sku: "Envasadora · Código SKU",
  linea5_especie: "Línea 5 · Especie",
  l5_proceso: "Línea 5 · Proceso",
  l5_kgIngresados: "Línea 5 · Kg Ingresados",
  l5_kgAprobados: "Línea 5 · Kg Aprobados",
  cajasProducidas: "Cajas Producidas",
  cajasProgramadas: "Cajas Programadas",
  cajasTeoricas: "Cajas Teóricas (Consumo)",
  cajasConsumidasReal: "Cajas Consumidas (Real)",
  bolsasTeoricas: "Bolsas Teóricas (Consumo)",
  bolsasConsumidasReal: "Bolsas Consumidas (Real)",
  huboIncidentes: "¿Hubo Incidentes?",
  huboAccidentes: "¿Hubo Accidentes?",
  armado_activa: "¿Armado de Materiales Activo?",
  editingId: "ID en Edición",
};

// Traduce un nombre de campo técnico (ej. "kg_linea1_e2_apr_t1_kg") a un
// nombre fácil de entender (ej. "Línea 1 · Especie 2 · Aprobado Tipo 1 - Kg").
function nombreAmigable(campo, area) {
  if (CAMPOS_FIJOS[campo]) return CAMPOS_FIJOS[campo];

  let m;

  // Comentarios por categoría: com_Producción, com_Calidad, etc.
  m = campo.match(/^com_(.+)$/);
  if (m) return "Comentario · " + m[1];

  // Materiales de inicio/cierre: inicioMat_3 / finMat_3
  m = campo.match(/^(inicioMat|finMat)_(\d+)$/);
  if (m) {
    const lista = area === "envasado" ? MATERIALES_ENVASADO : MATERIALES_SELECCION;
    const nombre = lista[parseInt(m[2], 10)] || ("Material #" + m[2]);
    const momento = m[1] === "inicioMat" ? "Inicio" : "Cierre";
    return "Material (" + momento + ") · " + nombre;
  }

  // Dotación general: dg_0, dg_1...
  m = campo.match(/^dg_(\d+)$/);
  if (m) {
    const lista = area === "envasado" ? DOTACION_GENERAL_ENVASADO : DOTACION_GENERAL_SELECCION;
    return "Dotación General · " + (lista[parseInt(m[1], 10)] || ("Rol #" + m[1]));
  }

  // Línea: activa / proceso (1..5) / dotación (0..2)
  m = campo.match(/^linea_(\w+?)_activa$/);
  if (m) return (LINEAS_LABEL[m[1]] || m[1]) + " · ¿Activa?";

  m = campo.match(/^linea_(\w+?)_proceso(\d?)$/);
  if (m) {
    const slot = m[2] || "1";
    return (LINEAS_LABEL[m[1]] || m[1]) + " · Proceso (Especie " + slot + ")";
  }

  m = campo.match(/^linea_(\w+?)_dot(\d+)$/);
  if (m) return (LINEAS_LABEL[m[1]] || m[1]) + " · Dotación: " + (DOTACION_LINEA[parseInt(m[2], 10)] || ("Rol #" + m[2]));

  // Kg ingresados / aprobados (legado, 1 solo valor) por línea + especie
  m = campo.match(/^kg_(\w+?)_e(\d+)_ing$/);
  if (m) return (LINEAS_LABEL[m[1]] || m[1]) + " · Especie " + m[2] + " · Kg Ingresados";

  m = campo.match(/^kg_(\w+?)_e(\d+)_apr$/);
  if (m) return (LINEAS_LABEL[m[1]] || m[1]) + " · Especie " + m[2] + " · Kg Aprobados";

  // Kg aprobados con hasta 3 tipos: kg_linea1_e2_apr_t1_tipo / _kg
  m = campo.match(/^kg_(\w+?)_e(\d+)_apr_t(\d+)_tipo$/);
  if (m) return (LINEAS_LABEL[m[1]] || m[1]) + " · Especie " + m[2] + " · Aprobado Tipo " + m[3] + " (Especie Aprobada)";

  m = campo.match(/^kg_(\w+?)_e(\d+)_apr_t(\d+)_kg$/);
  if (m) return (LINEAS_LABEL[m[1]] || m[1]) + " · Especie " + m[2] + " · Aprobado Tipo " + m[3] + " (Kg)";

  // Armado de materiales
  m = campo.match(/^armado_dot(\d+)$/);
  if (m) return "Armado de Materiales · Dotación: " + (DOTACION_LINEA[parseInt(m[1], 10)] || ("Rol #" + m[1]));

  // Línea 5 y Envasadora (dentro de Envasado)
  m = campo.match(/^linea5_dot(\d+)$/);
  if (m) return "Línea 5 · Dotación: " + (DOTACION_LINEA[parseInt(m[1], 10)] || ("Rol #" + m[1]));

  m = campo.match(/^envasadora_dot(\d+)$/);
  if (m) return "Envasadora · Dotación: " + (DOTACION_LINEA[parseInt(m[1], 10)] || ("Rol #" + m[1]));

  // Nada coincidió: devuelve el campo original, un poco más legible
  // (guiones bajos → espacios, primera letra mayúscula).
  return campo.replace(/_/g, " ").replace(/^\w/, function (c) { return c.toUpperCase(); });
}

function ordenarRegistros(registros) {
  return registros.slice().sort(function (a, b) {
    const fa = a.fecha || "", fb = b.fecha || "";
    if (fa !== fb) return fa < fb ? -1 : 1;
    const ta = ORDEN_TURNO[a.turno] || (ORDEN_TURNO[a.turnoEntrega] || 0);
    const tb = ORDEN_TURNO[b.turno] || (ORDEN_TURNO[b.turnoEntrega] || 0);
    if (ta !== tb) return ta - tb;
    return (a.id || 0) - (b.id || 0);
  });
}

// Convierte un array de objetos (con distintos campos entre sí) en una tabla
// con nombres de columna legibles = unión de todas las llaves que aparecen.
function armarTablaLegible(ss, tabName, registrosJson, area) {
  let registros;
  try {
    registros = JSON.parse(registrosJson);
  } catch (e) {
    return;
  }
  if (!Array.isArray(registros)) return;

  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);
  sh.clear();

  if (registros.length === 0) {
    sh.getRange(1, 1).setValue("Sin registros todavía.");
    return;
  }

  registros = ordenarRegistros(registros);

  // Columnas técnicas = unión de todas las llaves, en el orden en que aparecen.
  const columnasTecnicas = [];
  const vistas = {};
  registros.forEach(function (r) {
    Object.keys(r).forEach(function (k) {
      if (!vistas[k]) { vistas[k] = true; columnasTecnicas.push(k); }
    });
  });

  // Encabezado legible (puede repetirse el mismo nombre amigable para llaves
  // distintas — poco probable, pero si pasa, Sheets lo deja igual sin problema).
  const encabezado = columnasTecnicas.map(function (c) { return nombreAmigable(c, area); });

  const filas = [encabezado];
  registros.forEach(function (r) {
    filas.push(columnasTecnicas.map(function (c) {
      const v = r[c];
      if (v === undefined || v === null) return "";
      if (typeof v === "object") return JSON.stringify(v); // ej: items de Entrega
      return v;
    }));
  });

  sh.getRange(1, 1, filas.length, columnasTecnicas.length).setValues(filas);
  sh.getRange(1, 1, 1, columnasTecnicas.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, columnasTecnicas.length);
}

// -----------------------------------------------------------------------
// PESTAÑAS DE TRABAJO — "SKU Envasado" y "Procesos-Especie Selección".
// Si no existen todavía, se crean con encabezados + una fila de ejemplo.
// Si YA existen (como ahora, con tus datos cargados), esta función no las
// toca para nada — no se sobrescriben ni se vuelven a poblar solas.
// -----------------------------------------------------------------------
function asegurarPestañasDeTrabajo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let skuSh = ss.getSheetByName("SKU Envasado (Base de Datos)");
  if (!skuSh) {
    skuSh = ss.insertSheet("SKU Envasado (Base de Datos)");
    const encabezado = [
      "Código SKU", "Producto", "Cliente", "Código Bolsa", "Nombre Bolsa",
      "Código Caja", "Nombre Caja", "Bolsas x Caja", "Cajas x Pallet",
      "Kg x Caja", "Tipo Pallet", "TI x HI", "Slip Sheet",
    ];
    skuSh.appendRow(encabezado);
    skuSh.getRange(1, 1, 1, encabezado.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    skuSh.appendRow(["EJEMPLO-001", "Frutilla IQF 1x30lb (ejemplo, bórrame)", "Cliente X", "B-001", "Bolsa 1kg impresa", "C-001", "Caja MTC1310", 30, 40, 13.62, "MTC1310", "10x5", "Sí"]);
    skuSh.getRange(2, 1, 1, encabezado.length).setFontColor("#94a3b8").setFontStyle("italic");
    skuSh.setFrozenRows(1);
    skuSh.autoResizeColumns(1, encabezado.length);
  }

  let procSh = ss.getSheetByName("Procesos y Especies Selección (Base de Datos)");
  if (!procSh) {
    procSh = ss.insertSheet("Procesos y Especies Selección (Base de Datos)");
    const encabezado2 = ["Línea", "Código de Proceso", "Especie"];
    procSh.appendRow(encabezado2);
    procSh.getRange(1, 1, 1, encabezado2.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    procSh.appendRow(["Línea 1", "FRUT ORG NUEVO IQF (ejemplo, bórrame)", "FRUTILLA ORG"]);
    procSh.getRange(2, 1, 1, encabezado2.length).setFontColor("#94a3b8").setFontStyle("italic");
    procSh.setFrozenRows(1);
    procSh.autoResizeColumns(1, encabezado2.length);
  }
}

// Guarda o actualiza una clave (POST desde la app cada vez que algo cambia).
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.type === "storage") {
      // 1) Respaldo técnico (siempre, para todas las claves)
      const sh = getOrCreateStorageSheet(ss);
      const last = sh.getLastRow();
      let foundRow = -1;
      if (last > 1) {
        const keys = sh.getRange(2, 1, last - 1, 1).getValues();
        for (let i = 0; i < keys.length; i++) {
          if (keys[i][0] === body.key) { foundRow = i + 2; break; }
        }
      }
      const row = [body.key, body.value, new Date().toISOString()];
      if (foundRow > 0) sh.getRange(foundRow, 1, 1, 3).setValues([row]);
      else sh.appendRow(row);

      // 2) Tabla legible aparte, solo para las claves de "registros por turno"
      const cfg = TABLAS_LEGIBLES[body.key];
      if (cfg) armarTablaLegible(ss, cfg.tab, body.value, cfg.area);

      return json({ ok: true });
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

function onOpen() {
  // Se ejecuta solo cuando alguien abre el Sheet en el navegador — crea las
  // pestañas de trabajo si todavía no existen.
  asegurarPestañasDeTrabajo();
}

// Lee una clave (GET desde la app cada vez que abre una pantalla).
const LINEA_LABEL_A_KEY = {
  "Línea 1": "linea1", "Línea 3": "linea3", "Línea 4": "linea4",
  "Línea 5": "linea5", "Línea 6": "linea6",
};

// Lee la pestaña "SKU Envasado (Base de Datos)" y devuelve solo las filas
// nuevas (las que están debajo de la fila separadora "▼▼▼..."), sin contar
// la fila de ejemplo. Convierte cada fila al formato que usa la app.
function leerSkuNuevosDeLaHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("SKU Envasado (Base de Datos)");
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const filas = sh.getRange(2, 1, last - 1, 13).getValues();

  let debajoDelSeparador = false;
  const nuevos = [];
  filas.forEach(function (f) {
    const col1 = String(f[0] || "");
    if (col1.indexOf("▼▼▼") === 0) { debajoDelSeparador = true; return; }
    if (!debajoDelSeparador) return; // todavía estamos en la parte de "solo lectura" (base de datos actual)
    if (!col1.trim()) return; // fila vacía
    if (col1.toUpperCase().indexOf("EJEMPLO") !== -1) return; // fila de ejemplo, no es un SKU real

    nuevos.push({
      sku: col1, producto: f[1] || "", cliente: f[2] || "",
      codBolsa: f[3] || "", nomBolsa: f[4] || "",
      codCaja: f[5] || "", nomCaja: f[6] || "",
      bolsasXCaja: f[7] || "", cajasXPallet: f[8] || "", kgXCaja: f[9] || "",
      tipoPallet: f[10] || "", tixhi: f[11] || "", slipSheet: f[12] || "",
    });
  });
  return nuevos;
}

// Lee la pestaña "Procesos y Especies Selección (Base de Datos)" y devuelve
// solo las filas nuevas (debajo del separador), convertidas al formato
// { id, lineaKey, proceso, especie } que ya usa la app (useProcesosExtra).
function leerProcesosNuevosDeLaHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Procesos y Especies Selección (Base de Datos)");
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const filas = sh.getRange(2, 1, last - 1, 3).getValues();

  let debajoDelSeparador = false;
  const nuevos = [];
  filas.forEach(function (f, idx) {
    const linea = String(f[0] || "");
    if (linea.indexOf("▼▼▼") === 0) { debajoDelSeparador = true; return; }
    if (!debajoDelSeparador) return;
    const proceso = String(f[1] || "");
    if (!proceso.trim()) return;
    if (proceso.toUpperCase().indexOf("EJEMPLO") !== -1) return;

    const lineaKey = LINEA_LABEL_A_KEY[linea] || linea;
    nuevos.push({
      id: "sheet-" + idx, // id estable derivado de la fila, para no duplicar en cada carga
      lineaKey: lineaKey,
      proceso: proceso,
      especie: f[2] || "",
    });
  });
  return nuevos;
}

function doGet(e) {
  const key = e.parameter && e.parameter.key;
  if (!key) {
    asegurarPestañasDeTrabajo();
    return json({ ok: true, app: "Bitácora de Turnos", mensaje: "El script está funcionando." });
  }

  // "sku-nuevos": se arma siempre a partir de la pestaña del Sheet — no hay
  // nada que la app guarde ahí, solo lee lo que el Jefe fue anotando.
  if (key === "sku-nuevos") {
    return json({ ok: true, value: JSON.stringify(leerSkuNuevosDeLaHoja()) });
  }

  // "procesos-extra": fusiona lo que ya estaba guardado (agregado desde la
  // propia app, en "Gestionar especies por línea") con lo nuevo anotado en
  // la pestaña del Sheet, evitando duplicar la misma línea+proceso.
  if (key === "procesos-extra") {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = getOrCreateStorageSheet(ss);
      const last = sh.getLastRow();
      let guardados = [];
      if (last > 1) {
        const data = sh.getRange(2, 1, last - 1, 2).getValues();
        for (let i = 0; i < data.length; i++) {
          if (data[i][0] === key) { guardados = JSON.parse(data[i][1] || "[]"); break; }
        }
      }
      const delSheet = leerProcesosNuevosDeLaHoja();
      const yaExiste = {};
      guardados.forEach(function (p) { yaExiste[p.lineaKey + "|" + p.proceso] = true; });
      const combinados = guardados.concat(delSheet.filter(function (p) { return !yaExiste[p.lineaKey + "|" + p.proceso]; }));
      return json({ ok: true, value: JSON.stringify(combinados) });
    } catch (err) {
      return json({ ok: false, error: err.message });
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Storage");
    if (!sh) return json({ ok: true, value: null });
    const last = sh.getLastRow();
    if (last < 2) return json({ ok: true, value: null });
    const data = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) return json({ ok: true, value: data[i][1] });
    }
    return json({ ok: true, value: null });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}
