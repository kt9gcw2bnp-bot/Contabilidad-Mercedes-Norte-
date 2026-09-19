/* ============================================================
   BACKEND GENÉRICO — Plataforma de Exámenes Digitales
   MSc. Gerardo González González — CTP Mercedes Norte
   (basado en el formato replicado del manual de M.Sc. Ericka Hernández Mena)

   Este archivo NO se modifica de un examen a otro. Cada examen nuevo:
     1. Crea su propia Hoja de Google (sheets.new)
     2. Pega este mismo código en Extensiones → Apps Script
     3. Lo implementa como Aplicación Web (Ejecutar como: Yo / Acceso: Cualquiera)
   Todo el contenido (preguntas, roster, datos administrativos) vive
   únicamente en el archivo HTML del examen, nunca aquí.

   Pestañas que usa (se crean solas la primera vez que se escribe en ellas):
     - "Config"   : Clave | Valor           (ej. Habilitado=SI/NO, Especificaciones=<json>)
     - "Entregas" : una fila por estudiante (autoguardado + entrega final)
   ============================================================ */

function doGet(e) {
  const callback = e.parameter.callback;
  function respond(obj) {
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(obj) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    if (e.parameter.action === 'getConfig') {
      const rows = getRows_('Config');
      const value = {};
      rows.forEach(r => { if (r.Clave) value[r.Clave] = r.Valor; });
      return respond({ ok: true, value: value });
    }

    if (e.parameter.sheet) {
      const rows = getRows_(e.parameter.sheet);
      return respond({ ok: true, rows: rows });
    }

    return respond({ ok: false, error: 'Falta parámetro sheet o action.' });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const p = e.parameter;
    const sheetName = p.sheet;
    if (!sheetName) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Falta sheet' })).setMimeType(ContentService.MimeType.JSON);

    if (p.action === 'delete') {
      deleteRows_(sheetName, p.matchField, p.matchValue);
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }

    // Datos de la fila = todos los parámetros menos los de control
    const row = {};
    Object.keys(p).forEach(k => {
      if (k !== 'sheet' && k !== 'action' && k !== 'keyField') row[k] = p[k];
    });

    upsertRow_(sheetName, p.keyField, row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ---------- helpers internos (operan siempre sobre la Hoja donde vive este script) ---------- */

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getRows_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

function upsertRow_(sheetName, keyField, row) {
  const sheet = getSheet_(sheetName);
  let values = sheet.getDataRange().getValues();
  let headers = values.length ? values[0] : [];

  // Agregar encabezados nuevos que vengan en la fila y todavía no existan
  Object.keys(row).forEach(k => {
    if (headers.indexOf(k) === -1) headers.push(k);
  });
  if (values.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    values = [headers];
  } else if (headers.length > values[0].length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  let targetRowIndex = -1; // índice base 0 dentro de "values" (values[0] = encabezados)
  if (keyField && row[keyField] !== undefined) {
    const keyCol = headers.indexOf(keyField);
    if (keyCol !== -1) {
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][keyCol]).trim() === String(row[keyField]).trim()) {
          targetRowIndex = i;
          break;
        }
      }
    }
  }

  const newRowArr = headers.map(h => (row[h] !== undefined ? row[h] : (targetRowIndex !== -1 ? values[targetRowIndex][headers.indexOf(h)] : '')));

  if (targetRowIndex !== -1) {
    sheet.getRange(targetRowIndex + 1, 1, 1, headers.length).setValues([newRowArr]);
  } else {
    sheet.appendRow(newRowArr);
  }
}

function deleteRows_(sheetName, matchField, matchValue) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0];
  const col = headers.indexOf(matchField);
  if (col === -1) return;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][col]).trim() === String(matchValue).trim()) {
      sheet.deleteRow(i + 1);
    }
  }
}
