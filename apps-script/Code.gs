/**
 * Tortuguitas Daylist — Google Apps Script
 * -----------------------------------------
 * 1. Creá una hoja de cálculo nueva en Google Drive.
 * 2. Extensiones → Apps Script → pegá TODO este archivo.
 * 3. Ejecutá una vez la función setupDaylistSheets() (pedirá permisos).
 * 4. Proyecto → Configuración → Propiedades del script →
 *    agregá: API_TOKEN = (una clave secreta, ej. tortuguitas-2026)
 * 5. Implementar → Nueva implementación → Tipo: Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier persona
 * 6. Copiá la URL terminada en /exec y pegala en index.html (SHEETS_API_URL).
 *
 * Estructura de la hoja (la crea setupDaylistSheets):
 *
 * Hoja "Propietarios"
 *   A:id | B:name | C:color | D:sala
 *
 * Hoja "Tareas"
 *   A:id | B:ownerId | C:desc | D:priority | E:estado |
 *   F:startDate | G:deadline | H:endDate | I:producto | J:order
 */

var OWNER_HEADERS = ['id', 'name', 'color', 'sala'];
var TASK_HEADERS = [
  'id', 'ownerId', 'desc', 'priority', 'estado',
  'startDate', 'deadline', 'endDate', 'producto', 'order'
];

var DEFAULT_OWNERS = [
  { id: 'o1', name: 'Agustina Castillo', color: '#0d7377', sala: 'Florida' },
  { id: 'o2', name: 'Agustín Márquez', color: '#c45c26', sala: 'Florida' },
  { id: 'o3', name: 'Kevin García', color: '#3b6ea5', sala: 'Adrogué' },
  { id: 'o4', name: 'Rafael Barberi', color: '#7b4b94', sala: 'Florida' },
  { id: 'o5', name: 'Gustavo Sotelo', color: '#2d8a5e', sala: 'Merlo' },
  { id: 'o6', name: 'Albana Pais', color: '#b53d5a', sala: 'Adrogué' }
];

/** Ejecutar UNA vez desde el editor de Apps Script. */
function setupDaylistSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abrí este script desde una hoja de cálculo (vinculado).');

  var owners = getOrCreateSheet_(ss, 'Propietarios', OWNER_HEADERS);
  var tasks = getOrCreateSheet_(ss, 'Tareas', TASK_HEADERS);

  // Semilla de propietarios si está vacía
  if (owners.getLastRow() < 2) {
    var rows = DEFAULT_OWNERS.map(function (o) {
      return [o.id, o.name, o.color, o.sala];
    });
    owners.getRange(2, 1, rows.length, OWNER_HEADERS.length).setValues(rows);
  }

  // Ocultar hojas basura tipo "Hoja 1" vacía
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (name !== 'Propietarios' && name !== 'Tareas' && sh.getLastRow() === 0) {
      try { ss.deleteSheet(sh); } catch (e) { /* ignore */ }
    }
  });

  SpreadsheetApp.flush();
  Logger.log('Setup OK. Spreadsheet ID: ' + ss.getId());
}

/** GET ?token=XXX  → { ok, owners, tasks } */
function doGet(e) {
  try {
    e = e || { parameter: {} };
    if (!isAuthorized_(e.parameter && e.parameter.token)) {
      return json_({ ok: false, error: 'Unauthorized' });
    }
    var data = readAll_();
    return json_({ ok: true, owners: data.owners, tasks: data.tasks });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/**
 * POST body (text/plain JSON):
 *   { token, owners: [...], tasks: [...] }
 * Reemplaza el contenido completo de ambas hojas.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var body = JSON.parse(raw);
    if (!isAuthorized_(body.token)) {
      return json_({ ok: false, error: 'Unauthorized' });
    }
    writeAll_(body.owners || [], body.tasks || []);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* ignore */ }
  }
}

/* ——— Internals ——— */

function isAuthorized_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!expected) return true; // sin token configurado: abierto (solo pruebas)
  return String(token || '') === String(expected);
}

function getOrCreateSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var empty = first.every(function (c) { return c === '' || c === null; });
  if (empty || sh.getLastRow() === 0) {
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function readAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ownersSh = ss.getSheetByName('Propietarios');
  var tasksSh = ss.getSheetByName('Tareas');
  if (!ownersSh || !tasksSh) {
    throw new Error('Faltan hojas Propietarios/Tareas. Ejecutá setupDaylistSheets().');
  }
  return {
    owners: sheetToObjects_(ownersSh, OWNER_HEADERS),
    tasks: sheetToObjects_(tasksSh, TASK_HEADERS).map(normalizeTaskRow_)
  };
}

function writeAll_(owners, tasks) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ownersSh = getOrCreateSheet_(ss, 'Propietarios', OWNER_HEADERS);
  var tasksSh = getOrCreateSheet_(ss, 'Tareas', TASK_HEADERS);

  writeObjects_(ownersSh, OWNER_HEADERS, owners.map(function (o) {
    return {
      id: o.id || '',
      name: o.name || '',
      color: o.color || '',
      sala: o.sala || 'Florida'
    };
  }));

  writeObjects_(tasksSh, TASK_HEADERS, (tasks || []).map(function (t) {
    return {
      id: t.id || '',
      ownerId: t.ownerId || '',
      desc: t.desc || '',
      priority: t.priority || 'Media',
      estado: t.estado || 'Sin Comenzar',
      startDate: t.startDate || '',
      deadline: t.deadline || '',
      endDate: t.endDate || '',
      producto: t.producto || '',
      order: typeof t.order === 'number' ? t.order : Number(t.order) || 0
    };
  }));
}

function sheetToObjects_(sh, headers) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  // getRange(row, column, numRows, numColumns)
  var numRows = lastRow - 1;
  var values = sh.getRange(2, 1, numRows, headers.length).getValues();
  return values
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (key, i) {
        var v = row[i];
        obj[key] = v === null || v === undefined ? '' : v;
      });
      return obj;
    });
}

function writeObjects_(sh, headers, objects) {
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  if (!objects.length) return;
  var rows = objects.map(function (obj) {
    return headers.map(function (key) {
      var v = obj[key];
      return v === null || v === undefined ? '' : v;
    });
  });
  // getRange(row, column, numRows, numColumns) — 3er arg = cantidad de filas
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function normalizeTaskRow_(t) {
  var order = t.order;
  if (typeof order === 'string' && order !== '') order = Number(order);
  if (typeof order !== 'number' || isNaN(order)) order = 0;
  return {
    id: String(t.id || ''),
    ownerId: String(t.ownerId || ''),
    desc: String(t.desc || ''),
    priority: String(t.priority || 'Media'),
    estado: String(t.estado || 'Sin Comenzar'),
    startDate: formatDateCell_(t.startDate),
    deadline: formatDateCell_(t.deadline),
    endDate: formatDateCell_(t.endDate),
    producto: String(t.producto || ''),
    order: order
  };
}

function formatDateCell_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
