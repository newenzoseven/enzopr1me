// ===== Servidor de Enzo · Sensibilidades (Google Apps Script + Google Sheets) =====
// 1) Cambia la contraseña aquí (o créala en Configuración del proyecto > Propiedades del script > PASS)
const PASS_DEFAULT = 'cambia-esto';
const HOJA = 'clientes';

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(HOJA);
  if (!s) { s = ss.insertSheet(HOJA); s.appendRow(['t', 'id', 'nombre', 'celular', 'sensibilidad', 'estado', 'letras', 'precio', 'seed']); }
  return s;
}
function filas_() {
  const v = hoja_().getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ i: i + 2, t: String(r[0]), id: String(r[1]), n: r[2], c: r[3], s: r[4], e: r[5], sfx: String(r[6]), p: r[7], seed: r[8] }));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const P = PropertiesService.getScriptProperties();
  return out_({ cfg: JSON.parse(P.getProperty('cfg') || '{}') });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let b = {};
    try { b = JSON.parse(e.postData.contents); } catch (x) {}
    return out_(accion_(b));
  } catch (x) {
    return out_({ ok: false, err: 'srv' });
  } finally { lock.releaseLock(); }
}

function accion_(b) {
  const P = PropertiesService.getScriptProperties();
  const id = String(b.id || '').replace(/\D/g, '');

  // --- Cliente: registra su ID (aparece solo en el panel del dueño)
  if (b.a === 'reg') {
    if (id.length < 6) return { ok: false };
    const n = String(b.nombre || '').slice(0, 60), c = String(b.cel || '').slice(0, 60), sens = String(b.sens || '').slice(0, 90);
    const pend = filas_().filter(r => r.id === id && r.e === 'pendiente');
    if (pend.length) hoja_().getRange(pend[pend.length - 1].i, 3, 1, 3).setValues([[n, c, sens]]);
    else hoja_().appendRow([Date.now(), id, n, c, sens, 'pendiente', '', '', '']);
    return { ok: true };
  }

  // --- Cliente: pone las 2 letras y pide su sensibilidad (un solo uso)
  if (b.a === 'claim') {
    const C = CacheService.getScriptCache(), k = 'f' + id, f = Number(C.get(k) || 0);
    if (f >= 8) return { ok: false, err: 'wait' };
    const sfx = String(b.sfx || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    const r = filas_().find(r => r.id === id && r.e === 'confirmado' && r.sfx === sfx);
    if (!r) { C.put(k, String(f + 1), 600); return { ok: false, err: 'no' }; }
    hoja_().getRange(r.i, 6).setValue('usado');
    return { ok: true, key: r.t, seed: Number(r.seed) };
  }

  // --- A partir de aquí solo el dueño (con contraseña)
  if (String(b.pw || '') !== (P.getProperty('PASS') || PASS_DEFAULT)) return { ok: false, err: 'pw' };

  if (b.a === 'list') {
    return { ok: true, rows: filas_().slice(-300).map(r => ({ t: r.t, id: r.id, n: r.n, c: r.c, s: r.s, e: r.e, sfx: r.sfx, p: r.p })) };
  }
  if (b.a === 'ok') {
    if (id.length < 6) return { ok: false };
    const mias = filas_().filter(r => r.id === id);
    const pend = mias.filter(r => r.e === 'pendiente').pop();
    const usadas = mias.filter(r => r.e !== 'pendiente').map(r => r.sfx);
    const al = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let sfx; do { sfx = al[Math.floor(Math.random() * 32)] + al[Math.floor(Math.random() * 32)]; } while (usadas.indexOf(sfx) >= 0);
    const cfg = JSON.parse(P.getProperty('cfg') || '{}'), seed = Math.floor(Math.random() * 4e9);
    if (pend) hoja_().getRange(pend.i, 6, 1, 4).setValues([['confirmado', sfx, cfg.precio || '5.00', seed]]);
    else hoja_().appendRow([Date.now(), id, '', '', '', 'confirmado', sfx, cfg.precio || '5.00', seed]);
    return { ok: true, sfx: sfx };
  }
  if (b.a === 'del') {
    const r = filas_().find(r => r.t === String(b.t));
    if (r) hoja_().deleteRow(r.i);
    return { ok: true };
  }
  if (b.a === 'savecfg') { P.setProperty('cfg', JSON.stringify(b.cfg || {})); return { ok: true }; }
  return { ok: false };
}
