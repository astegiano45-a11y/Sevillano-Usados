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
  config: ['id', 'actualizado', 'json'],
  refs: ['id', 'titulo', 'pilar', 'actualizado', 'json'],
  recursos: ['id', 'categoria', 'titulo', 'actualizado', 'json'],
  drive: ['id', 'nombre', 'carpeta', 'tipo', 'fecha', 'json'],
  comentarios: ['id', 'tipo', 'paso', 'texto', 'estado', 'actualizado', 'json']
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
    if (!SHEETS[body.sheet] || body.sheet === 'drive') return json_({ ok: false, error: 'Hoja desconocida' });
    if (body.action === 'upsert') { const isNewCm = body.sheet === 'comentarios' && findRow_(sheet_('comentarios'), body.item.id) < 0; upsert_(body.sheet, body.item); if (isNewCm) try { avisoComentario_(body.item); } catch (e) {} try { calSync_(body.sheet, body.item); } catch (e) {} return json_({ ok: true }); }
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
    seccion: it.sub || '', titulo: it.title || it.titulo || it.nota || '', estado: it.estado || '', formato: it.formato || '',
    red: it.red || '', gancho: it.gancho || '', tipo: it.tipo || '', hecha: it.done ? 'Sí' : 'No',
    lugar: it.lugar || '', prioridad: it.prioridad || '', categoria: it.cat || '', paso: it.paso || '', texto: it.texto || '', actualizado: now, json: JSON.stringify(it)
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


/* ================= RESÚMENES POR MAIL =================
 * Primera vez: elegí la función setupAvisos y tocá Ejecutar (pide permiso para enviar mails).
 * Para probar sin esperar: ejecutá testDiario, testSemanal o testMensual (te llegan solo a vos).
 */
const APP_URL = 'https://astegiano45-a11y.github.io/mkt/';
const PIL_NAME = { valor: 'Valor', exp: 'Experiencia', prod: 'Producto' };
const AREA_NAME = { tradicional: '0km Tradicional', ovalo: 'Plan Óvalo', usados: 'Usados', posventa: 'Posventa' };
const DONE = ['Publicado', 'Multimedios'];

function setupAvisos() {
  ScriptApp.getProjectTriggers().filter(t => /^mail(Diario|Semanal|Mensual)$/.test(t.getHandlerFunction())).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncDrive').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncDrive').timeBased().everyHours(1).create();
  try { syncDrive(); } catch (e) { Logger.log('Drive: ' + e); }
  ScriptApp.newTrigger('mailDiario').timeBased().everyDays(1).atHour(8).nearMinute(30).inTimezone(TZ).create();
  ScriptApp.newTrigger('mailSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).inTimezone(TZ).create();
  ScriptApp.newTrigger('mailMensual').timeBased().onMonthDay(1).atHour(9).inTimezone(TZ).create();
  Logger.log('Avisos programados: diario 8:30 (lun-vie), semanal viernes 17:00, mensual día 1 9:00.');
}

function cfg_() {
  const d = readAll_();
  const b = (d.config || []).find(x => x.id === 'brand') || {};
  return { data: d, goal: Number(b.goal) || 3, gHist: b.goalHist == null ? 1 : Number(b.goalHist), gFeed: b.goalFeed == null ? 2 : Number(b.goalFeed), gPlaca: Number(b.goalPlaca) || 0, mails: (b.mails || []).filter(m => m.email) };
}
const fmt_ = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
const nice_ = s => { const d = ymd_(s); return ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'][d.getDay()] + ' ' + d.getDate() + '/' + (d.getMonth() + 1); };
function monday_(d) { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; }
function add_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }

function occurs_(t, s) {
  if (!t.date) return false;
  const len = t.dateEnd && t.dateEnd > t.date ? Math.round((ymd_(t.dateEnd) - ymd_(t.date)) / 864e5) : 0;
  const diff = Math.round((ymd_(s) - ymd_(t.date)) / 864e5);
  if (diff < 0) return false;
  if (!t.repeat) return diff <= len;
  if (t.repeatUntil && s > t.repeatUntil) return false;
  if (t.repeat === 'mensual') { const a = ymd_(t.date), b = ymd_(s); return b.getDate() >= a.getDate() && b.getDate() - a.getDate() <= len; }
  return diff % (t.repeat === 'quincenal' ? 14 : 7) <= len;
}

function checks_(done, c) {
  const h = done.filter(p => p.formato === 'Historia').length, f = done.filter(p => p.formato === 'Reel' || p.formato === 'Carrusel').length;
  const it = (ok, t) => '<span style="display:inline-block;margin:2px 4px 2px 0;padding:3px 10px;border-radius:12px;font-size:13px;' + (ok ? 'background:#E1F5EC;color:#0E9F6E' : 'background:#FDEBDF;color:#C2410C') + '">' + (ok ? '✓ ' : '✗ ') + t + '</span>';
  return '<p>' + it(h >= c.gHist, 'Historias ' + h + '/' + c.gHist) + it(f >= c.gFeed, 'Reel o carrusel ' + f + '/' + c.gFeed) +
    Object.keys(PIL_NAME).map(k => it(done.some(p => p.pilar === k), PIL_NAME[k])).join('') +
    (c.gPlaca ? it(done.filter(p => p.formato === 'Placa').length >= c.gPlaca, 'Placas ' + done.filter(p => p.formato === 'Placa').length + '/' + c.gPlaca) : ' <span style="font-size:13px">Placas: <b>' + done.filter(p => p.formato === 'Placa').length + '</b></span>') + '</p>';
}
function postLine_(p) {
  return '<li><b>' + esc_(p.title) + '</b> <span style="color:#5D6880">· ' + nice_(p.date) + ' · ' + (PIL_NAME[p.pilar] || 'Sin pilar') +
    (p.area ? ' · ' + AREA_NAME[p.area] : '') + ' · ' + esc_(p.formato || '') + ' · <i>' + esc_(p.estado || 'Idea') + '</i></span></li>';
}
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]); }
function wrap_(title, body) {
  return '<div style="font-family:Arial,sans-serif;max-width:600px;color:#0D1424">' +
    '<div style="border-bottom:3px solid #0A57FF;padding-bottom:8px;margin-bottom:14px"><b style="font-size:18px">Marketing Ford Sevillano</b><br><span style="color:#5D6880">' + title + '</span></div>' +
    body + '<p style="margin-top:22px"><a href="' + APP_URL + '" style="background:#0A57FF;color:#fff;padding:10px 16px;border-radius:20px;text-decoration:none">Abrir el Content OS</a></p></div>';
}
function bar_(done, goal) {
  const pct = Math.min(100, Math.round(done / goal * 100)); const col = done >= goal ? '#0E9F6E' : '#0A57FF';
  return '<div style="background:#E2E7F0;border-radius:10px;height:12px;margin:6px 0"><div style="width:' + pct + '%;background:' + col + ';height:12px;border-radius:10px"></div></div>';
}
function send_(kind, subject, html, onlyMe) {
  const c = cfg_();
  const to = onlyMe ? [Session.getActiveUser().getEmail()] : c.mails.filter(m => m[kind]).map(m => m.email);
  if (!to.length) return;
  MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html, name: 'Marketing Ford Sevillano' });
}

function buildDiario_() {
  const c = cfg_(); const { data, goal } = c; const today = new Date(); const t = fmt_(today); const ayer = fmt_(add_(today, -1));
  const mon = fmt_(monday_(today)), sun = fmt_(add_(monday_(today), 6));
  const hoyP = data.posts.filter(p => p.date === t);
  const atras = data.posts.filter(p => p.date < t && p.date >= fmt_(add_(today, -14)) && ['Guion', 'Grabado', 'Editado'].indexOf(p.estado) >= 0);
  const hoyT = data.tasks.filter(x => occurs_(x, t) && !(x.done && !x.repeat)).sort((a, b) => (a.time || '') > (b.time || '') ? 1 : -1);
  const week = data.posts.filter(p => p.date >= mon && p.date <= sun); const done = week.filter(p => DONE.indexOf(p.estado) >= 0).length;
  let h = '';
  h += '<h3>📲 Hoy se publica</h3>' + (hoyP.length ? '<ul>' + hoyP.map(postLine_).join('') + '</ul>' : '<p style="color:#5D6880">No hay publicación planificada para hoy.</p>');
  if (atras.length) h += '<h3 style="color:#C2410C">⚠️ Atrasadas (no se marcaron como publicadas)</h3><ul>' + atras.map(postLine_).join('') + '</ul>';
  h += '<h3>📅 Agenda de hoy</h3>' + (hoyT.length ? '<ul>' + hoyT.map(x => '<li>' + (x.time ? '<b>' + x.time + '</b> ' : '') + esc_(x.title) + (x.lugar ? ' <span style="color:#5D6880">· ' + esc_(x.lugar) + '</span>' : '') + '</li>').join('') + '</ul>' : '<p style="color:#5D6880">Sin reuniones ni eventos.</p>');
  h += '<h3>🎯 Objetivo semanal: ' + done + ' de ' + goal + '</h3>' + bar_(done, goal) + checks_(week.filter(p => DONE.indexOf(p.estado) >= 0), c);
  if (today.getDay() === 4 && done + week.filter(p => p.date > t).length < goal) h += '<p style="color:#C2410C"><b>Ojo:</b> con lo planificado no llegan al objetivo de esta semana. Sumen una publicación.</p>';
  return { subject: (atras.length ? '⚠️ ' : '') + 'Hoy en Marketing · ' + nice_(t), html: wrap_('Resumen del día · ' + nice_(t), h) };
}
function mailDiario() { const d = new Date().getDay(); if (d === 0 || d === 6) return; const r = buildDiario_(); send_('diario', r.subject, r.html); }
function testDiario() { const r = buildDiario_(); send_('diario', '[Prueba] ' + r.subject, r.html, true); }

function resumen_(from, to, goalTotal, weekly) {
  const c = cfg_(); const { data } = c;
  const ps = data.posts.filter(p => p.date >= from && p.date <= to);
  const done = ps.filter(p => DONE.indexOf(p.estado) >= 0); const pend = ps.filter(p => ['Guion', 'Grabado', 'Editado'].indexOf(p.estado) >= 0);
  const cnt = (arr, key, names) => Object.keys(names).map(k => names[k] + ': <b>' + arr.filter(p => p[key] === k).length + '</b>').join(' · ') + ' · Sin asignar: <b>' + arr.filter(p => !names[p[key]]).length + '</b>';
  let h = '<h3>🎯 ' + done.length + ' de ' + goalTotal + ' publicaciones (' + Math.round(done.length / goalTotal * 100) + '%)</h3>' + bar_(done.length, goalTotal);
  if (weekly) h += checks_(done, c);
  h += '<p><b>Por pilar:</b> ' + cnt(done, 'pilar', PIL_NAME) + '</p><p><b>Por área:</b> ' + cnt(done, 'area', AREA_NAME) + '</p>';
  const mm = done.filter(p => p.estado === 'Multimedios').length; if (mm) h += '<p><b>En multimedios (radio, diario, TV):</b> ' + mm + '</p>';
  h += '<h3>✅ Publicadas</h3>' + (done.length ? '<ul>' + done.map(postLine_).join('') + '</ul>' : '<p style="color:#5D6880">Ninguna.</p>');
  if (pend.length) h += '<h3 style="color:#C2410C">⏳ Quedaron sin publicar</h3><ul>' + pend.map(postLine_).join('') + '</ul>';
  return h;
}
function buildSemanal_() {
  const { data, goal } = cfg_(); const mon = monday_(new Date());
  let h = resumen_(fmt_(mon), fmt_(add_(mon, 6)), goal, true);
  const nMon = add_(mon, 7); const next = data.posts.filter(p => p.date >= fmt_(nMon) && p.date <= fmt_(add_(nMon, 6)));
  h += '<h3>🗓️ La semana que viene</h3>' + (next.length ? '<ul>' + next.map(postLine_).join('') + '</ul>' : '') +
    (next.length < goal ? '<p style="color:#C2410C">Hay ' + next.length + ' planificadas de ' + goal + '. Completen el calendario antes del lunes.</p>' : '');
  return { subject: 'Cierre de semana · Marketing Ford Sevillano', html: wrap_('Semana del ' + nice_(fmt_(mon)) + ' al ' + nice_(fmt_(add_(mon, 6))), h) };
}
function mailSemanal() { const r = buildSemanal_(); send_('semanal', r.subject, r.html); }
function testSemanal() { const r = buildSemanal_(); send_('semanal', '[Prueba] ' + r.subject, r.html, true); }

function buildMensual_() {
  const { goal } = cfg_(); const now = new Date();
  const first = new Date(now.getFullYear(), now.getMonth() - 1, 1), last = new Date(now.getFullYear(), now.getMonth(), 0);
  const weeks = Math.round(last.getDate() / 7 * 10) / 10; const goalM = Math.round(goal * weeks);
  const mes = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'][first.getMonth()];
  const h = '<p>Objetivo del mes: ' + goal + ' por semana × ' + weeks + ' semanas ≈ <b>' + goalM + '</b> publicaciones.</p>' + resumen_(fmt_(first), fmt_(last), goalM);
  return { subject: 'Resumen de ' + mes + ' · Marketing Ford Sevillano', html: wrap_('Resumen mensual · ' + mes + ' ' + first.getFullYear(), h) };
}
function mailMensual() { const r = buildMensual_(); send_('mensual', r.subject, r.html); }
function testMensual() { const r = buildMensual_(); send_('mensual', '[Prueba] ' + r.subject, r.html, true); }

/* ================= BIBLIOTECA DE DRIVE =================
 * Lee la carpeta configurada en la app (Marca → Carpeta de Drive) y la copia a la hoja "drive".
 * Corre sola cada hora (la programa setupAvisos). Se puede ejecutar a mano: syncDrive.
 */
const MODELOS = ['Ranger', 'Everest', 'Territory', 'Bronco', 'Maverick', 'Kuga', 'EcoSport', 'Focus', 'Ka', 'Transit', 'F-150', 'Mustang', 'Raptor'];
function syncDrive() {
  const b = (readAll_().config || []).find(x => x.id === 'brand') || {};
  const m = String(b.driveFolder || '').match(/folders\/([\w-]+)/) || String(b.driveFolder || '').match(/^([\w-]{20,})$/);
  if (!m) { Logger.log('No hay carpeta de Drive configurada en Marca.'); return; }
  const out = [];
  const walk = (folder, path, depth) => {
    const files = folder.getFiles();
    while (files.hasNext() && out.length < 3000) {
      const f = files.next(); const mime = f.getMimeType();
      const tipo = mime.indexOf('video/') === 0 ? 'Video' : mime.indexOf('image/') === 0 ? 'Imagen' : null;
      if (!tipo) continue;
      const name = f.getName();
      const mod = MODELOS.find(x => name.toLowerCase().replace(/[\s_-]/g, '').indexOf(x.toLowerCase().replace(/[\s_-]/g, '')) >= 0) || '';
      const it = { id: f.getId(), nombre: name, carpeta: path, tipo: tipo, modelo: mod, fecha: Utilities.formatDate(f.getDateCreated(), TZ, 'yyyy-MM-dd') };
      out.push([it.id, it.nombre, it.carpeta, it.tipo, it.fecha, JSON.stringify(it)]);
    }
    if (depth > 4) return;
    const subs = folder.getFolders();
    while (subs.hasNext()) { const s = subs.next(); walk(s, path + ' / ' + s.getName(), depth + 1); }
  };
  const root = DriveApp.getFolderById(m[1]);
  walk(root, root.getName(), 0);
  const sh = sheet_('drive');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, SHEETS.drive.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, SHEETS.drive.length).setValues(out);
  upsert_('config', { id: 'drivemeta', at: new Date().toISOString(), count: out.length });
  Logger.log('Biblioteca: ' + out.length + ' archivos.');
}


/* ================= COMENTARIOS DE LA GUÍA =================
 * Cuando alguien deja una duda o sugerencia en la app, te llega un mail.
 */
function avisoComentario_(c) {
  const to = Session.getEffectiveUser().getEmail();
  const html = '<div style="font-family:Arial,sans-serif;max-width:560px"><p><b>' + esc_(c.quien || 'Alguien') + '</b> dejó una <b>' + esc_(c.tipo || 'duda') + '</b>' +
    (c.paso ? ' sobre <i>' + esc_(c.paso) + '</i>' : '') + ':</p><blockquote style="border-left:3px solid #0A57FF;margin:0;padding:6px 12px;background:#F5F7FB">' + esc_(c.texto) +
    '</blockquote><p>Respondela desde la app: pestaña <b>Guía → Dudas y sugerencias</b>.</p><p><a href="' + APP_URL + '">Abrir el Content OS</a></p></div>';
  MailApp.sendEmail({ to: to, subject: '💬 Nueva ' + (c.tipo || 'duda').toLowerCase() + ' en el Content OS', htmlBody: html, name: 'Marketing Ford Sevillano' });
}
