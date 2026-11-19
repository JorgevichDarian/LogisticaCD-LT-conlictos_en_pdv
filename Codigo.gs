/* =====================================================================
   POE-DIS-014 · Seguimiento de Conflictos en Punto de Entrega
   Apps Script que conecta el reporte web con:
     - la planilla de RESPUESTAS del formulario (solo lectura)
     - la planilla de SEGUIMIENTO (comentarios, planes, acta de mediación)
   Pegar en: planilla de SEGUIMIENTO → Extensiones → Apps Script.
   Implementar → Nueva implementación → Aplicación web
     · Ejecutar como: Yo
     · Quién tiene acceso: Cualquier persona
   Copiar la URL (termina en /exec) y pegarla en index.html (APPS_SCRIPT_URL).
   ===================================================================== */

const SEG_ID   = '1Fx-Dy5xC7iFH0_zYKEDi4SP75jlZf-QY2dm77YbGawY';  // planilla de seguimiento (este archivo)
const RESP_ID  = '1SC9NwwRCbEEcHSyevz8eLY92Iqmq8HyRPqGYXzABOjY';  // respuestas del formulario
const RESP_GID = 321852746;                                       // pestaña "Respuestas de formulario 1"
const ADMIN_PASS = '2026';                                        // misma clave que la pestaña ANALISTA

// Orden de columnas de la hoja "Seguimiento" (clave interna, encabezado)
const COLS = [
  ['id','ID'], ['fr','Fecha reporte'], ['nc','Nro cliente'], ['cli','Cliente'], ['col','Colaborador'], ['mot','Motivo'],
  ['etapa','Etapa'], ['bloq','Bloqueo'], ['fbloq','Fecha bloqueo'], ['cmt','Comentario seguimiento'],
  ['test','Declaración colaborador y testigos'], ['fvis','Fecha visita PDV'], ['relato','Relato del cliente'],
  ['acdesc','Condición de descarga acordada'], ['actrato','Compromiso de trato'], ['plan','Plan de acción'],
  ['sector','Sector'], ['resp','Responsable'], ['fc','Fecha compromiso'], ['status','Status'],
  ['frehab','Fecha rehabilitación'], ['mon1','Monitoreo entrega 1'], ['mon2','Monitoreo entrega 2'],
  ['u','Modificado por'], ['t','Última modificación'], ['cu','Creado por'], ['ct','Creado']
];
// Hoja "Pagos": problemas de pago cargados a mano (sección SANCIONES)
const PAGO_COLS = [
  ['id','ID'], ['f','Fecha'], ['nc','Nro cliente'], ['cli','Cliente'], ['tipo','Tipo de problema'], ['monto','Monto'],
  ['pla','Planilla / Pedido'], ['det','Detalle'], ['u','Cargado por'], ['t','Cargado el'], ['est','Estado'],
  ['du','Descartado por'], ['dm','Motivo descarte'], ['dt','Fecha descarte']
];
const EDITABLES = ['etapa','bloq','fbloq','cmt','test','fvis','relato','acdesc','actrato','plan','sector','resp','fc','status','frehab','mon1','mon2'];
const SNAP = ['fr','nc','cli','col','mot'];
const isDate_ = v => Object.prototype.toString.call(v) === '[object Date]';

function doGet(e){
  const p = (e && e.parameter) || {};
  let out;
  try{
    if (p.action === 'save')      out = save_(JSON.parse(p.data));
    else if (p.action === 'cfg')  out = saveCfg_(JSON.parse(p.data));
    else if (p.action === 'pago') out = savePago_(JSON.parse(p.data));
    else if (p.action === 'pagoDesc') out = descPago_(JSON.parse(p.data));
    else                          out = readAll_();
  }catch(err){ out = { ok:false, error:String(err && err.message || err) }; }
  const txt = JSON.stringify(out);
  const cb = String(p.callback || '');
  if (/^[A-Za-z_$][\w$]{0,60}$/.test(cb))
    return ContentService.createTextOutput(cb + '(' + txt + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- lectura ---------- */
function readAll_(){
  // Respuestas
  const rs = SpreadsheetApp.openById(RESP_ID);
  const sh = rs.getSheets().filter(s => s.getSheetId() === RESP_GID)[0] || rs.getSheets()[0];
  const tz = rs.getSpreadsheetTimeZone();
  const vals = sh.getDataRange().getValues();
  const head = (vals.shift() || []).map(String);
  const rows = vals.filter(r => r.some(v => v !== '' && v !== null)).map(r => r.map(v =>
    isDate_(v) ? Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss") : v));
  // Seguimiento
  const seg = {};
  const ss = segSheet_();
  const sv = ss.getDataRange().getValues(); sv.shift();
  sv.forEach(r => { if (!r[0]) return; const o = {}; COLS.forEach(([k], i) => { const v = r[i]; if (v !== '' && v !== null) o[k] = isDate_(v) ? isoDate_(v, k) : v; }); seg[String(r[0])] = o; });
  // Pagos
  const pagos = [];
  const pv = pagoSheet_().getDataRange().getValues(); pv.shift();
  pv.forEach(r => { if (!r[0]) return; const o = {}; PAGO_COLS.forEach(([k], i) => { const v = r[i]; if (v !== '' && v !== null) o[k] = isDate_(v) ? isoDate_(v, k) : v; }); pagos.push(o); });
  // Config
  let cfg = null;
  const cs = cfgSheet_(); const cv = cs.getDataRange().getValues();
  cv.forEach(r => { if (r[0] === 'cfg' && r[1]) { try{ cfg = JSON.parse(r[1]); }catch(e){} } });
  return { ok:true, head, rows, seg, cfg, pagos, now: new Date().toISOString() };
}
function isoDate_(d, k){
  const tz = 'America/Argentina/Buenos_Aires';
  return (k === 't' || k === 'ct') ? d.getTime() : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/* ---------- escritura ---------- */
function save_(d){
  if (!d || !d.id) throw new Error('Falta id');
  const user = String(d.user || 'sin usuario').slice(0, 60);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh = segSheet_();
    const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
    let row = ids.indexOf(String(d.id)) + 1;
    let cur = {};
    if (row > 0){
      const r = sh.getRange(row, 1, 1, COLS.length).getValues()[0];
      COLS.forEach(([k], i) => { if (r[i] !== '' && r[i] !== null) cur[k] = r[i]; });
    } else {
      row = sh.getLastRow() + 1;
      cur = { id: String(d.id), cu: user, ct: new Date() };
    }
    const snap = d.snap || {};
    SNAP.forEach(k => { if (snap[k] != null && snap[k] !== '') cur[k] = String(snap[k]); });
    const f = d.fields || {}; const log = [];
    Object.keys(f).forEach(k => { if (EDITABLES.indexOf(k) < 0) return; const v = String(f[k] == null ? '' : f[k]).slice(0, 5000); cur[k] = v; log.push(k); });
    cur.u = user; cur.t = new Date();
    const out = COLS.map(([k]) => cur[k] == null ? '' : cur[k]);
    // las fechas escritas como texto ISO para que no las reinterprete la configuración regional
    sh.getRange(row, 1, 1, COLS.length).setNumberFormat('@');
    sh.getRange(row, 1, 1, COLS.length).setValues([out.map((v, i) => isDate_(v) ? Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm') : v)]);
    const hs = histSheet_();
    if (log.length) hs.getRange(hs.getLastRow() + 1, 1, log.length, 6).setValues(log.map(k => [new Date(), user, String(d.id), cur.cli || '', label_(k), String(f[k] == null ? '' : f[k])]));
    const res = {}; COLS.forEach(([k]) => { if (cur[k] !== '' && cur[k] != null) res[k] = isDate_(cur[k]) ? cur[k].getTime() : cur[k]; });
    return { ok:true, id:String(d.id), seg:res };
  } finally { lock.releaseLock(); }
}
function saveCfg_(d){
  if (!d || String(d.pass) !== ADMIN_PASS) throw new Error('Contraseña incorrecta');
  const cs = cfgSheet_(); const cv = cs.getDataRange().getValues();
  let row = cv.findIndex(r => r[0] === 'cfg') + 1; if (row < 1) row = cs.getLastRow() + 1;
  cs.getRange(row, 1, 1, 2).setValues([['cfg', JSON.stringify(d.cfg || {})]]);
  return { ok:true };
}

/* ---------- problemas de pago ---------- */
function savePago_(d){
  const p = (d && d.p) || {};
  if (!p.nc && !p.cli) throw new Error('Falta el cliente');
  const user = String(d.user || 'sin usuario').slice(0, 60);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh = pagoSheet_();
    const now = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
    const o = { id: 'P' + new Date().getTime(), f: String(p.f || now.slice(0,10)), nc: String(p.nc || ''), cli: String(p.cli || ''),
      tipo: String(p.tipo || ''), monto: String(p.monto || ''), pla: String(p.pla || ''), det: String(p.det || '').slice(0, 3000),
      u: user, t: now, est: 'Abierta' };
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, PAGO_COLS.length).setNumberFormat('@').setValues([PAGO_COLS.map(([k]) => o[k] == null ? '' : o[k])]);
    return { ok:true, pago:o };
  } finally { lock.releaseLock(); }
}
function descPago_(d){
  const ids = (d && d.ids) || []; if (!ids.length) throw new Error('Falta id');
  const user = String(d.user || 'sin usuario').slice(0, 60);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh = pagoSheet_(); const v = sh.getDataRange().getValues();
    const now = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
    const col = k => PAGO_COLS.findIndex(c => c[0] === k) + 1; let n = 0;
    v.forEach((r, i) => { if (i && ids.indexOf(String(r[0])) >= 0){
      sh.getRange(i + 1, col('est'), 1, 4).setNumberFormat('@').setValues([['Descartada', user, String(d.motivo || '').slice(0, 1000), now]]); n++; } });
    return { ok:true, n, du:user, dt:now };
  } finally { lock.releaseLock(); }
}

/* ---------- hojas ---------- */
function ensure_(name, head){
  const ss = SpreadsheetApp.openById(SEG_ID);
  let sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0){ sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function segSheet_(){  return ensure_('Seguimiento', COLS.map(c => c[1])); }
function histSheet_(){ return ensure_('Historial', ['Fecha y hora','Usuario','ID reporte','Cliente','Campo','Valor']); }
function cfgSheet_(){  return ensure_('Config', ['Clave','Valor (JSON)']); }
function pagoSheet_(){ return ensure_('Pagos', PAGO_COLS.map(c => c[1])); }
function label_(k){ const c = COLS.filter(x => x[0] === k)[0]; return c ? c[1] : k; }

/* Ejecutar una vez desde el editor para autorizar permisos */
function probar(){ Logger.log(JSON.stringify(readAll_()).slice(0, 500)); }
