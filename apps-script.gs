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
    if (body.action === 'upsert') { upsert_(body.sheet, body.item); try { calSync_(body.sheet, body.item); } catch (e) {} return json_({ ok: true }); }
    if (body.action === 'delete') { del_(body.sheet, body.id); try { calDelete_(body.sheet, body.id); } catch (e) {} return json_({ ok: true }); }
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


/* ================= GOOGLE CALENDAR =================
 * Crea y mantiene un calendario "Marketing Ford Sevillano".
 * Primera vez: elegí la función syncAll en el menú de arriba y tocá Ejecutar (pide permiso de Calendar).
 */
const CAL_NAME = 'Marketing Ford Sevillano';
const TZ = 'America/Argentina/Ushuaia';
const PIL_EMOJI = { valor: '🔵', exp: '🟢', prod: '🟠' };
const AREA_TAG = { tradicional: '[0KM] ', ovalo: '[PLAN ÓVALO] ', usados: '[USADOS] ', posventa: '[POSVENTA] ' };

function cal_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('calId');
  let cal = id ? CalendarApp.getCalendarById(id) : null;
  if (!cal) {
    cal = CalendarApp.createCalendar(CAL_NAME, { timeZone: TZ, color: CalendarApp.Color.BLUE });
    props.setProperty('calId', cal.getId());
  }
  return cal;
}

function when_UNUSED(date, time) {
  const [y, m, d] = date.split('-').map(Number);
  const [hh, mm] = (time || '10:00').split(':').map(Number);
  const s = new Date(y, m - 1, d, hh, mm);
  const e = new Date(s.getTime() + (time ? 60 : 30) * 60000);
  return [s, e];
}

function calSync_(sheet, it) {
  if (sheet !== 'posts' && sheet !== 'tasks') return;
  const props = PropertiesService.getScriptProperties();
  const key = 'ev_' + sheet + '_' + it.id;
  const cal = cal_();
  removeEv_(cal, props.getProperty(key));
  props.deleteProperty(key);
  if (!it.date) return;
  if (sheet === 'tasks' && it.done && !it.repeat) return;
  let title, desc;
  if (sheet === 'posts') {
    title = (PIL_EMOJI[it.pilar] || '📲') + ' ' + (AREA_TAG[it.area] || '') + (it.title || 'Publicación') + ' · ' + (it.formato || '') + (it.red ? ' ' + it.red : '');
    desc = 'Estado: ' + (it.estado || 'Idea') + (it.sub ? '\nSección: ' + it.sub : '') + (it.auto ? '\nVehículo: ' + it.auto : '') +
      (it.gancho ? '\n\nGancho: ' + it.gancho : '') + (it.guion ? '\n\nGuion:\n' + it.guion : '');
  } else {
    title = ({ reunion: '👥 ', evento: '📍 ', tarea: '✅ ', recordatorio: '🔔 ' }[it.tipo] || '✅ ') + (it.title || '');
    desc = (it.lugar ? 'Lugar / con quién: ' + it.lugar + '\n' : '') + (it.notas || '');
  }
  const opts = { description: desc };
  if (sheet === 'tasks' && it.lugar) opts.location = it.lugar;
  const d0 = ymd_(it.date);
  const multi = sheet === 'tasks' && it.dateEnd && it.dateEnd > it.date;
  const dEnd = multi ? ymd_(it.dateEnd) : d0;
  let rec = null;
  if (sheet === 'tasks' && it.repeat) {
    rec = CalendarApp.newRecurrence();
    const rule = it.repeat === 'mensual' ? rec.addMonthlyRule() : rec.addWeeklyRule().interval(it.repeat === 'quincenal' ? 2 : 1);
    if (it.repeatUntil) rule.until(new Date(ymd_(it.repeatUntil).getTime() + 86400000));
  }
  let ev;
  if (sheet === 'tasks' && !it.time) {
    const endEx = new Date(dEnd.getTime() + 86400000);
    ev = rec ? cal.createAllDayEventSeries(title, d0, endEx, rec, opts) : cal.createAllDayEvent(title, d0, endEx, opts);
  } else {
    const [hh, mm] = (it.time || '10:00').split(':').map(Number);
    const s = new Date(d0.getFullYear(), d0.getMonth(), d0.getDate(), hh, mm);
    const e = multi ? new Date(dEnd.getFullYear(), dEnd.getMonth(), dEnd.getDate(), hh + 1, mm) : new Date(s.getTime() + (it.time ? 60 : 30) * 60000);
    ev = rec ? cal.createEventSeries(title, s, e, rec, opts) : cal.createEvent(title, s, e, opts);
  }
  ev.removeAllReminders();
  if (it.aviso !== 'no') {
    if (sheet === 'tasks' && !it.time) ev.addPopupReminder(15 * 60);
    else { ev.addPopupReminder(30); ev.addPopupReminder(sheet === 'posts' ? 24 * 60 : 60); }
  }
  props.setProperty(key, (rec ? 'S:' : '') + ev.getId());
}

function ymd_(s) { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }

function removeEv_(cal, stored) {
  if (!stored) return;
  try {
    if (stored.indexOf('S:') === 0) cal.getEventSeriesById(stored.slice(2)).deleteEventSeries();
    else cal.getEventById(stored).deleteEvent();
  } catch (e) {}
}

function calDelete_(sheet, id) {
  const props = PropertiesService.getScriptProperties();
  const key = 'ev_' + sheet + '_' + id;
  removeEv_(cal_(), props.getProperty(key));
  props.deleteProperty(key);
}

/** Ejecutar a mano una vez: pide permisos y pasa todo lo cargado al calendario. */
function syncAll() {
  const data = readAll_();
  data.posts.forEach(p => calSync_('posts', p));
  data.tasks.forEach(t => calSync_('tasks', t));
  Logger.log('Listo: ' + data.posts.length + ' publicaciones y ' + data.tasks.length + ' tareas. Calendario: ' + cal_().getName());
}
