/**
 * Sevillano Content OS — sincronización con Google Sheets
 * Pegar en Extensiones → Apps Script, guardar y publicar como Aplicación web
 * (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 *
 * Crea solas las hojas: posts, tasks, ideas, config.
 * Las columnas visibles son para leer la planilla; la columna "json" es la que usa la app.
 */

const SHEETS = {
  posts:  ['id', 'fecha', 'pilar', 'seccion', 'titulo', 'estado', 'formato', 'red', 'gancho', 'actualizado', 'json'],
  tasks:  ['id', 'fecha', 'hora', 'tipo', 'titulo', 'hecha', 'lugar', 'actualizado', 'json'],
  ideas:  ['id', 'pilar', 'seccion', 'titulo', 'prioridad', 'gancho', 'actualizado', 'json'],
  config: ['id', 'actualizado', 'json']
};

const PILARES = { valor: 'Valor', exp: 'Experiencia', prod: 'Producto' };

function doGet() {
  return json_({ ok: true, data: readAll_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'all') return json_({ ok: true, data: readAll_() });
    if (!SHEETS[body.sheet]) return json_({ ok: false, error: 'Hoja desconocida' });
    if (body.action === 'upsert') { upsert_(body.sheet, body.item); return json_({ ok: true }); }
    if (body.action === 'delete') { del_(body.sheet, body.id); return json_({ ok: true }); }
    return json_({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const out = {};
  Object.keys(SHEETS).forEach(name => {
    const sh = sheet_(name);
    const last = sh.getLastRow();
    if (last < 2) { out[name] = []; return; }
    const col = SHEETS[name].indexOf('json') + 1;
    out[name] = sh.getRange(2, col, last - 1, 1).getValues()
      .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
      .filter(Boolean);
  });
  return out;
}

function row_(name, it) {
  const now = new Date();
  const map = {
    id: it.id, fecha: it.date || '', hora: it.time || '', pilar: PILARES[it.pilar] || '',
    seccion: it.sub || '', titulo: it.title || '', estado: it.estado || '', formato: it.formato || '',
    red: it.red || '', gancho: it.gancho || '', tipo: it.tipo || '', hecha: it.done ? 'Sí' : 'No',
    lugar: it.lugar || '', prioridad: it.prioridad || '', actualizado: now, json: JSON.stringify(it)
  };
  return SHEETS[name].map(h => map[h]);
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function upsert_(name, item) {
  if (!item || !item.id) throw new Error('Falta el id');
  const sh = sheet_(name);
  const values = [row_(name, item)];
  const r = findRow_(sh, item.id);
  if (r > 0) sh.getRange(r, 1, 1, values[0].length).setValues(values);
  else sh.appendRow(values[0]);
}

function del_(name, id) {
  const sh = sheet_(name);
  const r = findRow_(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
