// =============================================================================
// BITÁCORA DE TURNOS — Apps Script Web App
// Recibe datos de la app (POST para guardar, GET para leer) y los guarda en
// una sola hoja llamada "Storage" dentro de este Google Sheet. Cada fila es
// una "clave" distinta de la app (ej: "programa-records", "seleccion-cierre-
// records", "insumos-config", etc.) con su contenido completo en JSON.
//
// INSTRUCCIONES (León Sheet):
// 1) Crea un Google Sheet nuevo (o usa uno existente).
// 2) Extensiones → Apps Script.
// 3) Borra todo el código de ejemplo y pega este archivo completo.
// 4) Arriba a la derecha: Implementar → Nueva implementación.
// 5) Tipo: "Aplicación web". Ejecutar como: "Yo". Quién tiene acceso:
//    "Cualquier usuario" (así el celular de cualquiera puede usarla).
// 6) Presiona "Implementar" y copia la URL que te entrega (termina en /exec).
// 7) Pega esa URL en src/main.jsx, en la línea window.__SHEETS_URL__ = "..."
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

// Guarda o actualiza una clave (POST desde la app cada vez que algo cambia).
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.type === "storage") {
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
      return json({ ok: true });
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

// Lee una clave (GET desde la app cada vez que abre una pantalla). Si viene
// ?key=xxx devuelve solo esa fila; si no viene key, solo confirma que el
// script está vivo (útil para probarlo pegando la URL en el navegador).
function doGet(e) {
  const key = e.parameter && e.parameter.key;
  if (!key) {
    return json({ ok: true, app: "Bitácora de Turnos", mensaje: "El script está funcionando." });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Storage");
    if (!sh) return json({ ok: true, value: null });
    const last = sh.getLastRow();
    if (last < 2) return json({ ok: true, value: null });
    const data = sh.getRange(2, 1, last - 1, 2).getValues(); // columnas key, value
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) return json({ ok: true, value: data[i][1] });
    }
    return json({ ok: true, value: null });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}
