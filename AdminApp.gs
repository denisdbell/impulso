/**
 * AdminApp.gs — PANEL ADMIN MÓVIL (?page=admin)
 *
 * Página web protegida por usuario + contraseña, pensada para el TELÉFONO, para que una persona
 * no técnica pueda operar el día a día sin abrir la planilla:
 *   • Aprobar / rechazar solicitudes (reutiliza verifyApplicantV2_ / rejectApplicant_).
 *   • Registrar pagos (reutiliza recordPayment).
 *   • Ver préstamos: saldo, estado, cuotas, mora + botones de WhatsApp para avisar al cliente.
 *   • Mensajear referencias (reutiliza refWaFormula_).
 *   • Ver / aprobar / rechazar prospectos.
 *
 * Comparte el ámbito global con LoanManagerV2.gs / Validaciones.gs / Reactivar.gs — NO redefine
 * ningún helper existente. Reutiliza: recordPayment, verifyApplicantV2_, rejectApplicant_,
 * findLoanById_, loanWaStatus_, cuotasDeLoan_, fundStats_, getClienteById_, remWhatsappFormula_,
 * freezeFormWaFormula_, planFormWaFormula_, paymentWaFormula_, refWaFormula_, waNumberAR_,
 * webAppBaseUrl_, constEq_, genAccessToken_, hexDigest_, settingDays_, getSetting_, setSetting_, guard_.
 *
 * SEGURIDAD: la app web se publica "Ejecutar como: yo / Acceso: cualquiera" (igual que los formularios
 * públicos), así que la puerta NO es el login de Google sino este usuario/contraseña + anti-fuerza-bruta
 * + interruptor "Panel admin activo". La contraseña se guarda CIFRADA (SHA-256 + sal), nunca en texto plano.
 *
 * INSTALACIÓN: pegar este archivo + agregar la rama `?page=admin` en doGet y el ítem de menú (ver plan).
 * Luego ejecutar `configurarClaveAdmin()` UNA vez desde el editor para fijar usuario y contraseña.
 */

/* ========================= CONFIG / CLAVE ========================= */

/** Agrega (sólo si faltan) las claves de configuración del panel admin. Append-only: NUNCA pisa
 *  valores existentes (clave/sal/usuario). Idempotente. Devuelve cuántas claves nuevas creó. */
function ensureAdminSettingsKeys_() {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sh) throw new Error('Falta la hoja "Configuración".');
  const last = sh.getLastRow();
  const existing = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => { const k = String(r[0]).trim(); if (k) existing[k] = true; });
  const defaults = [
    ['Usuario admin', 'admin'],
    ['Hash de contraseña admin (SHA-256)', ''],
    ['Sal de contraseña admin', ''],
    ['Vigencia de sesión admin (minutos)', '60'],
    ['Intentos antes de bloquear (admin)', '5'],
    ['Bloqueo por intentos admin (minutos)', '15'],
    ['Panel admin activo (SÍ/NO)', 'SÍ'],
  ];
  let added = 0;
  defaults.forEach(kv => { if (!existing[kv[0]]) { sh.appendRow([kv[0], kv[1]]); added++; } });
  if (added) { try { invalidateSettingsCache_(); } catch (e) {} }
  return added;
}

/** Hash de contraseña: SHA-256 hex de "sal:contraseña" (reusa hexDigest_). */
function adminHashPw_(salt, pass) { return hexDigest_(String(salt) + ':' + String(pass)); }

/** EDITOR SOLAMENTE: fija usuario + contraseña del panel. Genera una sal nueva y guarda el hash.
 *  Nunca guarda la contraseña en texto plano ni se puede ejecutar desde la web. */
function configurarClaveAdmin() {
  const ui = SpreadsheetApp.getUi();
  ensureAdminSettingsKeys_();
  const actual = String(getSetting_('Usuario admin') || 'admin').trim();
  const u = ui.prompt('Panel admin — usuario', 'Nombre de usuario para entrar al panel (actual: ' + actual + '):', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  const user = String(u.getResponseText() || '').trim();
  if (!user) { ui.alert('Usuario vacío. No se cambió nada.'); return; }
  const p = ui.prompt('Panel admin — contraseña', 'Contraseña (mínimo 6 caracteres; se guarda cifrada, no en texto plano):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const pass = String(p.getResponseText() || '');
  if (pass.length < 6) { ui.alert('La contraseña debe tener al menos 6 caracteres. No se cambió nada.'); return; }
  const salt = genAccessToken_();
  setSetting_('Usuario admin', user);
  setSetting_('Sal de contraseña admin', salt);
  setSetting_('Hash de contraseña admin (SHA-256)', adminHashPw_(salt, pass));
  const base = webAppBaseUrl_();
  ui.alert('Panel admin configurado',
    'Usuario «' + user + '» listo. La contraseña quedó guardada cifrada.\n\n' +
    'Entrá desde el teléfono en:\n' + (base ? (base + '?page=admin') : '(publicá la app web y cargá "URL de la app web (enlaces a clientes)" en Configuración)') +
    '\n\nConsejo: usá una contraseña larga y cambiá la "Vigencia de sesión admin (minutos)" si querés sesiones más cortas.',
    ui.ButtonSet.OK);
}

/* ========================= LÍMITES / SESIÓN ========================= */

function adminEnabled_() { return String(getSetting_('Panel admin activo (SÍ/NO)') || 'SÍ').trim().toUpperCase() !== 'NO'; }
function adminSessionSeconds_() { return Math.min(21600, Math.max(300, Math.round(settingDays_('Vigencia de sesión admin (minutos)', 60) * 60))); }
function adminMaxFails_() { return Math.max(1, Math.round(settingDays_('Intentos antes de bloquear (admin)', 5))); }
function adminLockoutSeconds_() { return Math.max(60, Math.round(settingDays_('Bloqueo por intentos admin (minutos)', 15) * 60)); }

// Throttle propio del admin (no reusa authBlocked_/authNoteFail_ porque esos leen los límites de los ENLACES).
function adminBlocked_(cache, key) { return Number(cache.get(key) || 0) >= adminMaxFails_(); }
function adminNoteFail_(cache, key) { cache.put(key, String(Number(cache.get(key) || 0) + 1), adminLockoutSeconds_()); }

/** Emite una sesión: token en CacheService (rápido) + espejo en ScriptProperties (sobrevive al desalojo del caché). */
function adminIssueSession_(user) {
  const token = genAccessToken_();
  const ttl = adminSessionSeconds_();
  const exp = Date.now() + ttl * 1000;
  CacheService.getScriptCache().put('adminsess:' + token, user + '|' + exp, ttl);
  try { PropertiesService.getScriptProperties().setProperty('ADMINSESS_' + token, JSON.stringify({ user: user, exp: exp })); } catch (e) { logError_('adminIssueSession_', e); }
  return { token: token, ttl: ttl };
}
/** Devuelve {user, exp} si la sesión es válida; null si no. Repuebla el caché desde Properties si hizo falta. */
function adminCheckSession_(session) {
  const tok = String(session || '').trim();
  if (!tok) return null;
  const cache = CacheService.getScriptCache();
  const v = cache.get('adminsess:' + tok);
  if (v) { const i = v.indexOf('|'); const exp = Number(v.slice(i + 1)); return exp > Date.now() ? { user: v.slice(0, i), exp: exp } : null; }
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('ADMINSESS_' + tok);
    if (raw) {
      const o = JSON.parse(raw);
      if (o && o.exp > Date.now()) {
        const left = Math.round((o.exp - Date.now()) / 1000);
        cache.put('adminsess:' + tok, o.user + '|' + o.exp, Math.min(21600, Math.max(1, left)));
        return o;
      }
      PropertiesService.getScriptProperties().deleteProperty('ADMINSESS_' + tok);
    }
  } catch (e) { logError_('adminCheckSession_', e); }
  return null;
}
function adminRequire_(session) { const s = adminCheckSession_(session); if (!s) throw new Error('Sesión vencida. Iniciá sesión de nuevo.'); return s; }
/** Limpia de ScriptProperties las sesiones expiradas (se llama al hacer login, para no acumular). */
function adminGcSessions_() {
  try {
    const props = PropertiesService.getScriptProperties(), all = props.getProperties(), now = Date.now();
    Object.keys(all).forEach(k => {
      if (k.indexOf('ADMINSESS_') === 0) { try { const o = JSON.parse(all[k]); if (!o || !(o.exp > now)) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } }
    });
  } catch (e) {}
}

/* ========================= LOGIN / LOGOUT ========================= */

function adminLogin(user, pass, ip) {
  return guard_('adminLogin', function () {
    if (!adminEnabled_()) return { ok: false, msg: 'El panel está deshabilitado.' };
    ensureAdminSettingsKeys_();
    const cache = CacheService.getScriptCache();
    const kUser = 'adminfail:user:' + String(user || '').slice(0, 40).toLowerCase();
    const kIp = ip ? 'adminfail:ip:' + String(ip).slice(0, 60) : '';
    if (adminBlocked_(cache, kUser) || (kIp && adminBlocked_(cache, kIp))) return { ok: false, msg: 'Demasiados intentos. Esperá unos minutos e intentá de nuevo.' };
    const storedUser = String(getSetting_('Usuario admin') || '').trim();
    const salt = String(getSetting_('Sal de contraseña admin') || '');
    const storedHash = String(getSetting_('Hash de contraseña admin (SHA-256)') || '');
    if (!storedUser || !storedHash || !salt) return { ok: false, msg: 'El panel no está configurado. Ejecutá «Configurar panel admin» desde la planilla.' };
    const okUser = constEq_(storedUser, String(user || '').trim());
    const okPass = constEq_(storedHash, adminHashPw_(salt, String(pass || '')));
    if (!okUser || !okPass) { adminNoteFail_(cache, kUser); if (kIp) adminNoteFail_(cache, kIp); return { ok: false, msg: 'Usuario o contraseña incorrectos.' }; }
    cache.remove(kUser); if (kIp) cache.remove(kIp);
    adminGcSessions_();
    const sess = adminIssueSession_(storedUser);
    let home = null; try { home = adminHomeCached_(); } catch (e) { logError_('adminLogin:home', e); }
    return { ok: true, session: sess.token, expiresInSec: sess.ttl, user: storedUser, home: home };
  });
}

function adminLogout(session) {
  return guard_('adminLogout', function () {
    const tok = String(session || '').trim();
    if (tok) { try { CacheService.getScriptCache().remove('adminsess:' + tok); } catch (e) {} try { PropertiesService.getScriptProperties().deleteProperty('ADMINSESS_' + tok); } catch (e) {} }
    return { ok: true };
  });
}

/* ========================= HELPERS COMPARTIDOS ========================= */

/** Extrae la URL de una fórmula =HYPERLINK("url","label") o devuelve la cadena si ya es una URL. '' si no hay. */
function waUrlFromFormula_(f) {
  const s = String(f || '');
  const m = /HYPERLINK\("([^"]+)"/.exec(s);
  if (m) return m[1];
  return /^https?:\/\//.test(s) ? s : '';
}
/** Clave estable (independiente de la fila) a partir de 2 valores. Reusa el patrón MD5 del "claim" de aprobación. */
function adminStableKey_(a, b) {
  const norm = v => (v instanceof Date) ? String(v.getTime()) : String(v == null ? '' : v).trim();
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, norm(a) + '␟' + norm(b)));
}

/* ------- SOLICITUDES PENDIENTES (Nuevos Prestatarios) ------- */

function adminPendingKeyForRow_(vals, m) {
  const dni = m['DNI'] ? vals[m['DNI'] - 1] : '';
  const fecha = m['Fecha de Envío'] ? vals[m['Fecha de Envío'] - 1] : '';
  return adminStableKey_(normDni_(dni), fecha);
}
function adminResolvePendingRow_(nb, appKey) {
  const last = nb.getLastRow(); if (last < 2) return 0;
  const m = headerMap_(nb);
  const vals = nb.getRange(2, 1, last - 1, nb.getLastColumn()).getValues();
  for (let i = 0; i < vals.length; i++) {
    const nombre = m['Nombre completo'] ? String(vals[i][m['Nombre completo'] - 1] || '').trim() : '';
    if (!nombre) continue;
    if (adminPendingKeyForRow_(vals[i], m) === appKey) return i + 2;
  }
  return 0;
}
function adminCountPending_() {
  const nb = getSS_().getSheetByName(CFG.SHEETS.NEW); if (!nb || nb.getLastRow() < 2) return 0;
  const m = headerMap_(nb), nc = m['Nombre completo'] || 2;
  const vals = nb.getRange(2, nc, nb.getLastRow() - 1, 1).getValues();
  let n = 0; vals.forEach(r => { if (String(r[0] || '').trim()) n++; }); return n;
}

function adminListPending(session) {
  return guard_('adminListPending', function () {
    adminRequire_(session);
    const nb = getSS_().getSheetByName(CFG.SHEETS.NEW);
    if (!nb || nb.getLastRow() < 2) return { ok: true, items: [] };
    const m = headerMap_(nb);
    const g = (vals, name) => m[name] ? vals[m[name] - 1] : '';
    const vals = nb.getRange(2, 1, nb.getLastRow() - 1, nb.getLastColumn()).getValues();
    const items = [];
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      const nombre = String(g(v, 'Nombre completo') || '').trim();
      if (!nombre) continue;
      const refs = [];
      [['Ref 1 Nombre', 'Ref 1 Vínculo', 'Ref 1 Teléfono', 'Ref 1 Validada?'],
       ['Ref 2 Nombre', 'Ref 2 Vínculo', 'Ref 2 Teléfono', 'Ref 2 Validada?']].forEach((rk, idx) => {
        const rn = String(g(v, rk[0]) || '').trim();
        const rp = String(g(v, rk[2]) || '').trim();
        if (rn || rp) refs.push({ idx: idx + 1, nombre: rn, vinculo: String(g(v, rk[1]) || '').trim(), telefono: rp, validada: isRefValidated_(g(v, rk[3])) });
      });
      items.push({
        appKey: adminPendingKeyForRow_(v, m),
        nombre: nombre,
        dni: String(g(v, 'DNI') || '').trim(),
        correo: String(g(v, 'Correo') || '').trim(),
        telefono: String(g(v, 'Teléfono') || '').trim(),
        monto: fmtMoney_(Number(g(v, 'Monto Solicitado')) || 0),
        plazo: String(g(v, 'Plazo (meses)') || g(v, 'Plazo') || '').toString().trim(),
        notas: String(g(v, 'Notas / Motivo') || '').trim(),
        tercero: /tercero|^no\b/i.test(String(g(v, '¿Cuenta propia?') || '')),
        bcra: {
          decision: String(g(v, 'Decisión') || '').trim(),
          peorSit: String(g(v, 'Peor Situación') || '').toString().trim(),
          pctAlDia: String(g(v, '% al día') || '').toString().trim(),
          calif: String(g(v, 'Calificación de Riesgo') || '').trim(),
          resumen: String(g(v, 'Resumen BCRA') || '').trim()
        },
        refs: refs
      });
    }
    return { ok: true, items: items };
  });
}

function adminPendingRefLinks(session, appKey) {
  return guard_('adminPendingRefLinks', function () {
    adminRequire_(session);
    const nb = getSS_().getSheetByName(CFG.SHEETS.NEW);
    if (!nb) return { ok: true, links: [] };
    const row = adminResolvePendingRow_(nb, appKey);
    if (!row) return { ok: false, msg: 'La solicitud cambió o ya no está. Recargá la lista.' };
    const m = headerMap_(nb);
    const vals = nb.getRange(row, 1, 1, nb.getLastColumn()).getValues()[0];
    const g = name => m[name] ? vals[m[name] - 1] : '';
    const who = String(g('Nombre completo') || '').trim();
    const links = [];
    [['Ref 1 Nombre', 'Ref 1 Teléfono', 'Ref 1 Vínculo'], ['Ref 2 Nombre', 'Ref 2 Teléfono', 'Ref 2 Vínculo']].forEach((rk, idx) => {
      const rn = String(g(rk[0]) || '').trim(), rp = String(g(rk[1]) || '').trim(), rl = String(g(rk[2]) || '').trim();
      if (!rn && !rp) return;
      const url = waUrlFromFormula_(refWaFormula_(rn, rp, rl, who));
      links.push({ label: 'Ref ' + (idx + 1) + (rn ? ': ' + rn : ''), url: url, sinTel: !url });
    });
    return { ok: true, links: links };
  });
}

function adminApprove(session, appKey, override) {
  return guard_('adminApprove', function () {
    adminRequire_(session);
    const nb = getSS_().getSheetByName(CFG.SHEETS.NEW);
    if (!nb) return { ok: false, msg: 'No hay solicitudes.' };
    const row = adminResolvePendingRow_(nb, appKey);
    if (!row) return { ok: false, msg: 'La solicitud cambió o ya no está. Recargá la lista.' };
    const res = verifyApplicantV2_(nb, row, !!override);  // re-bloquea + valida el "claim" por contenido
    if (res && res.ok) { adminInvalidateCaches_(); return { ok: true, loanId: res.loanId, name: res.name, msg: 'Aprobado ' + res.name + ' → ' + res.loanId + (res.overridden ? ' (límites anulados)' : '') + '. Préstamo pendiente de firma; se envió el contrato por correo.' }; }
    return { ok: false, msg: (res && res.msg) || 'No se pudo aprobar.' };
  });
}

function adminReject(session, appKey) {
  return guard_('adminReject', function () {
    adminRequire_(session);
    const nb = getSS_().getSheetByName(CFG.SHEETS.NEW);
    if (!nb) return { ok: false, msg: 'No hay solicitudes.' };
    const lock = LockService.getDocumentLock(); try { lock.waitLock(30000); } catch (e) {}
    try {
      const row = adminResolvePendingRow_(nb, appKey);
      if (!row) return { ok: false, msg: 'La solicitud cambió o ya no está. Recargá la lista.' };
      const nombre = String(nb.getRange(row, headerMap_(nb)['Nombre completo'] || 2).getValue() || '').trim();
      rejectApplicant_(nb, row);
      adminInvalidateCaches_();
      return { ok: true, msg: 'Rechazado ' + (nombre || 'solicitante') + '. Se archivó y se le avisó por correo.' };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  });
}

/* ------- PRÉSTAMOS (Prestatarios) ------- */

function adminCountActivos_() {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS); if (!bs || bs.getLastRow() < 2) return 0;
  const H = headerIndex_(bs);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const stC = colByAny_(H, ['Estado']) || PB.STATE;
  const n = bs.getLastRow() - 1;
  const ids = bs.getRange(2, idC, n, 1).getValues(), sts = bs.getRange(2, stC, n, 1).getValues();
  let c = 0;
  for (let i = 0; i < ids.length; i++) {
    if (!/^L-/i.test(String(ids[i][0]).trim())) continue;
    const st = String(sts[i][0] || '').toUpperCase();
    if (st !== ST.PAID && st !== ST.CLEARED) c++;
  }
  return c;
}

// Mapa liviano id Cliente → {name, dni} leyendo SOLO las columnas 1..4 de "Clientes" (estables en
// ambos esquemas). Memoizado por ejecución. Más barato que clientesCache_ (que lee 9-13 columnas +
// headerMap_), y la lista de préstamos sólo necesita nombre/DNI.
var _adminClientND = null;
function adminClientNameDni_() {
  if (_adminClientND) return _adminClientND;
  const map = {};
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS);
  if (sh && sh.getLastRow() >= 2) {
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, PC.DNI).getValues();  // ID, Nombre, Correo, DNI
    for (let i = 0; i < data.length; i++) {
      const id = String(data[i][PC.CLIENT_ID - 1] || '').trim();
      if (id) map[id] = { name: String(data[i][PC.NAME - 1] || ''), dni: String(data[i][PC.DNI - 1] || '') };
    }
  }
  _adminClientND = map;
  return map;
}

// Escaneo de préstamos: UNA lectura de Prestatarios (sólo hasta la última columna necesaria) + el mapa
// liviano de Clientes. O(n). El cliente lo cachea y filtra/busca localmente (sin ida y vuelta por tecla).
function adminScanLoans_(query) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs || bs.getLastRow() < 2) return [];
  const H = headerIndex_(bs);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const cliC = colByAny_(H, ['ID Cliente']) || PB.CLIENT_ID;
  const stC = colByAny_(H, ['Estado']) || PB.STATE;
  const balC = colByAny_(H, ['Saldo con Mora', 'Saldo Pendiente', 'Saldo Pendiente (hoy)']) || PB.BALANCE;
  const venC = colByAny_(H, ['Vencimiento', 'Fecha de Vencimiento']) || PB.DUE;
  const maxC = Math.max(idC, cliC, stC, balC, venC);
  const data = bs.getRange(2, 1, bs.getLastRow() - 1, maxC).getValues();
  const nd = adminClientNameDni_();
  const q = String(query || '').trim().toLowerCase();
  const today = startOfDay_(new Date()).getTime();
  const items = [];
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][idC - 1] || '').trim();
    if (!/^L-/i.test(id)) continue;
    const c = nd[String(data[i][cliC - 1] || '').trim()] || {};
    const nombre = c.name || '', dni = String(c.dni || '');
    if (q && (id.toLowerCase() + ' ' + nombre.toLowerCase() + ' ' + dni).indexOf(q) < 0) continue;
    const estado = String(data[i][stC - 1] || '').trim(), estUp = estado.toUpperCase();
    const saldoNum = round2_(Number(data[i][balC - 1]) || 0);
    const venc = data[i][venC - 1];
    const vd = (venc instanceof Date) ? venc : null;
    const dias = vd ? Math.round((startOfDay_(vd).getTime() - today) / 86400000) : null;
    items.push({
      loanId: id, nombre: nombre, dni: dni, estado: estado,
      saldo: fmtMoney_(saldoNum), saldoNum: saldoNum,
      venc: vd ? fmtDate_(vd) : '', dias: dias,
      payable: saldoNum > 0.009 && estUp !== ST.PAID && estUp !== ST.CLEARED
    });
  }
  return items;
}

// Lista de préstamos. La lista completa (query vacío, lo normal) se sirve de una caché corta (60 s);
// la búsqueda puntual (query) recomputa. El cliente normalmente pide '' y filtra en el navegador.
function adminListLoans(session, query) {
  return guard_('adminListLoans', function () {
    adminRequire_(session);
    const q = String(query || '').trim();
    if (q) return { ok: true, items: adminScanLoans_(q) };
    const cache = CacheService.getScriptCache();
    try { const v = cache.get('admin:loans:v1'); if (v) return { ok: true, items: JSON.parse(v) }; } catch (e) {}
    const items = adminScanLoans_('');
    try { const s = JSON.stringify(items); if (s.length <= 90000) cache.put('admin:loans:v1', s, 60); } catch (e) {}
    return { ok: true, items: items };
  });
}

/** Marca (o desmarca) una referencia como validada en "Nuevos Prestatarios" (columna "Ref N Validada?"). */
function adminValidateRef(session, appKey, refIdx, value) {
  return guard_('adminValidateRef', function () {
    adminRequire_(session);
    const nb = getSS_().getSheetByName(CFG.SHEETS.NEW);
    if (!nb) return { ok: false, msg: 'No hay solicitudes.' };
    const lock = LockService.getDocumentLock(); try { lock.waitLock(30000); } catch (e) {}
    try {
      const row = adminResolvePendingRow_(nb, appKey);
      if (!row) return { ok: false, msg: 'La solicitud cambió o ya no está. Recargá la lista.' };
      const col = headerMap_(nb)['Ref ' + Number(refIdx) + ' Validada?'];
      if (!col) return { ok: false, msg: 'No existe la columna de validación de esa referencia.' };
      nb.getRange(row, col).setValue(value === true);
      return { ok: true, refIdx: Number(refIdx), validada: value === true };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  });
}

/** Lista para la sección "Vencimientos": lee "Estudio de Recordatorios" (vencidos + por vencer) en una
 *  sola lectura, recalcula los días a hoy, y ordena del vencimiento MÁS RECIENTE al más antiguo. */
function adminReminders(session) {
  return guard_('adminReminders', function () {
    adminRequire_(session);
    const sh = getSS_().getSheetByName(CFG.SHEETS.REMINDER);
    if (!sh || sh.getLastRow() < REM_DATA_START) return { ok: true, items: [] };
    const n = sh.getLastRow() - REM_DATA_START + 1;
    const data = sh.getRange(REM_DATA_START, 1, n, REM_HEADERS.length).getValues();
    // Las columnas 📲 (Recordar/Congelar/Plan, 13-15) son fórmulas =HYPERLINK: getValues sólo devuelve la
    // ETIQUETA ("📲 WhatsApp"), por eso hay que leerlas con getFormulas para poder extraer la URL.
    const waForms = sh.getRange(REM_DATA_START, REM_WA_COL, n, 3).getFormulas();
    const today = startOfDay_(new Date()).getTime();
    const items = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const id = String(r[0] || '').trim(), nombre = String(r[1] || '').trim();
      if (!id && !nombre) continue;
      const venc = r[2] instanceof Date ? r[2] : null;
      const dias = venc ? Math.round((startOfDay_(venc).getTime() - today) / 86400000) : (Number(r[3]) || 0);
      // Todos los enlaces de WhatsApp de la fila (sólo los que tienen URL): recordatorio + (si aplican)
      // oferta de congelamiento y de plan de pago — las mismas columnas 📲 de "Estudio de Recordatorios".
      const wf = waForms[i] || [];
      const waLinks = [];
      [[0, '📲 Recordatorio'], [1, '❄ Ofrecer congelar'], [2, '📋 Ofrecer plan']]
        .forEach(function (c) { const u = waUrlFromFormula_(wf[c[0]]); if (u) waLinks.push({ label: c[1], url: u }); });
      items.push({
        loanId: id, nombre: nombre, venc: venc ? fmtDate_(venc) : '', dias: dias,
        estado: String(r[4] || '').trim(), saldo: fmtMoney_(Number(r[6]) || 0),
        waLinks: waLinks, _s: venc ? venc.getTime() : 0
      });
    }
    items.sort((a, b) => b._s - a._s);        // vencimiento más reciente primero
    items.forEach(it => { delete it._s; });
    return { ok: true, items: items };
  });
}

/** Botones de WhatsApp (lender → prestatario) pertinentes al estado del préstamo. */
function adminLoanWaLinks_(loan, s) {
  const out = [];
  const add = (f, lbl) => { const u = waUrlFromFormula_(f); if (u) out.push({ label: lbl, url: u }); };
  const now = new Date();
  const daysUntil = (loan.dueDate instanceof Date) ? Math.round((startOfDay_(loan.dueDate).getTime() - startOfDay_(now).getTime()) / 86400000) : 0;
  const fin = {
    interest: round2_(Math.max(0, (s.totalDue || 0) - (loan.principal || 0))),
    totalDue: s.totalDue, totalPaid: s.totalPaid, owedMora: s.owedMora, toClear: s.toClear,
    daysLate: Math.max(0, -daysUntil)
  };
  if (!s.paidOff) add(remWhatsappFormula_(loan, loan.dueDate, s.toClear, daysUntil, fin), s.late ? '📲 Avisar deuda' : '📲 Recordatorio');
  if (s.late && !s.frozen && !s.onPlan) add(freezeFormWaFormula_(loan), '❄ Ofrecer congelar');
  if (s.late && !s.onPlan) add(planFormWaFormula_(loan), '📋 Ofrecer plan');
  return out;
}

function adminLoanDetail(session, loanId) {
  return guard_('adminLoanDetail', function () {
    adminRequire_(session);
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan) return { ok: false, msg: 'No se encontró el préstamo ' + loanId + '.' };
    const s = loanWaStatus_(loan);
    const cuotas = (s.cuotas || []).map(c => ({
      num: c.num, venc: (c.due instanceof Date) ? fmtDate_(c.due) : '',
      monto: fmtMoney_(c.monto), pagado: fmtMoney_(c.pagado), saldo: fmtMoney_(c.saldo), estado: c.estado
    }));
    return {
      ok: true,
      loan: {
        loanId: loan.loanId, nombre: loan.name, dni: String(loan.dni || ''), telefono: String(loan.phone || ''),
        capital: fmtMoney_(loan.principal), total: fmtMoney_(s.totalDue), pagado: fmtMoney_(s.totalPaid),
        saldo: fmtMoney_(s.baseOwed), mora: fmtMoney_(s.owedMora), toClear: fmtMoney_(s.toClear),
        venc: (loan.dueDate instanceof Date) ? fmtDate_(loan.dueDate) : '',
        paidOff: s.paidOff, frozen: s.frozen, onPlan: s.onPlan, late: s.late
      },
      cuotas: cuotas,
      waLinks: adminLoanWaLinks_(loan, s)
    };
  });
}

function adminRecordPayment(session, loanId, amount, dateIso) {
  return guard_('adminRecordPayment', function () {
    adminRequire_(session);
    loanId = String(loanId || '').trim();
    const msg = recordPayment(loanId, amount, dateIso || '');  // reusa el flujo existente (valida, saldo vivo, recibo)
    adminInvalidateCaches_();
    let waUrl = '';
    try {
      const loan = findLoanById_(loanId), s = loanWaStatus_(loan);
      const amt = Number(String(amount || '').replace(/[^0-9.\-]/g, '')) || 0;
      waUrl = waUrlFromFormula_(paymentWaFormula_(loan, dateIso ? new Date(dateIso) : new Date(), amt, s));
    } catch (e) { logError_('adminRecordPayment:wa', e); }
    return { ok: true, msg: msg, waUrl: waUrl };
  });
}

/* ------- PROSPECTOS ------- */

function adminProspectKeyForRow_(vals, m) {
  const cuil = m['CUIL'] ? vals[m['CUIL'] - 1] : '';
  const fecha = m['Fecha de Captura'] ? vals[m['Fecha de Captura'] - 1] : '';
  return adminStableKey_(String(cuil).replace(/\D/g, ''), fecha);
}
function adminResolveProspectRow_(sh, key) {
  const last = sh.getLastRow(); if (last < 2) return 0;
  const m = headerMap_(sh);
  const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < vals.length; i++) {
    const nombre = m['Nombre'] ? String(vals[i][m['Nombre'] - 1] || '').trim() : '';
    if (!nombre) continue;
    if (adminProspectKeyForRow_(vals[i], m) === key) return i + 2;
  }
  return 0;
}
function adminCountProspects_() {
  const sh = getSS_().getSheetByName(CFG.SHEETS.PROSPECTS); if (!sh || sh.getLastRow() < 2) return 0;
  const m = headerMap_(sh), nc = m['Nombre'] || 2, ec = m['Estado'] || 5;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(nc, ec)).getValues();
  let n = 0; vals.forEach(r => { const nm = String(r[nc - 1] || '').trim(); const est = String(r[ec - 1] || '').trim().toUpperCase(); if (nm && est !== 'APROBADO' && est !== 'RECHAZADO') n++; }); return n;
}

function adminListProspects(session, showAll) {
  return guard_('adminListProspects', function () {
    adminRequire_(session);
    const sh = getSS_().getSheetByName(CFG.SHEETS.PROSPECTS);
    if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };
    const m = headerMap_(sh);
    const g = (v, name) => m[name] ? v[m[name] - 1] : '';
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const items = [];
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      const nombre = String(g(v, 'Nombre') || '').trim();
      if (!nombre) continue;
      const estado = String(g(v, 'Estado') || '').trim();
      if (!showAll && (estado.toUpperCase() === 'APROBADO' || estado.toUpperCase() === 'RECHAZADO')) continue;
      items.push({
        prospectKey: adminProspectKeyForRow_(v, m),
        nombre: nombre,
        cuil: String(g(v, 'CUIL') || '').trim(),
        telefono: String(g(v, 'Teléfono') || '').trim(),
        estado: estado || 'NUEVO',
        contactado: String(g(v, 'Contactado?') || '').trim(),
        nota: String(g(v, 'Nota') || '').trim()
      });
    }
    return { ok: true, items: items };
  });
}

/** Igual que aprobarProspecto_ pero SIN abrir diálogo (getUi() rompe en la web): marca APROBADO/Contactado,
 *  deja el enlace en «WhatsApp 📲» y DEVUELVE la URL de WhatsApp para que la página la abra. */
function aprobarProspectoWeb_(sh, row) {
  const H = headerMap_(sh);
  const name = String((H['Nombre'] ? sh.getRange(row, H['Nombre']).getValue() : '') || '').trim();
  const phone = String((H['Teléfono'] ? sh.getRange(row, H['Teléfono']).getValue() : '') || '').trim();
  if (H['Estado']) sh.getRange(row, H['Estado']).setValue('APROBADO');
  if (H['Contactado?']) sh.getRange(row, H['Contactado?']).setValue('Sí');
  const num = waNumberAR_(phone), waColC = H['WhatsApp 📲'];
  if (!num) { if (waColC) sh.getRange(row, waColC).setValue('sin teléfono'); return { name: name, waUrl: '' }; }
  const firstName = name.split(/\s+/)[0] || name || 'Hola';
  const empresa = companyName_(), url = webAppBaseUrl_(), nl = '\n';
  const msg = '¡Hola ' + firstName + '! 👋 Te escribimos de ' + empresa + '.' + nl +
    '¡Buenas noticias! Ya tenemos cupo disponible y te invitamos a solicitar tu préstamo. 🎉' + nl +
    'Completá tu solicitud acá:' + nl + (url || '(falta configurar la URL de la app web en Configuración)') + nl +
    'Cualquier duda, respondé este mensaje. ¡Gracias por tu paciencia! 💛';
  const waUrl = 'https://wa.me/' + num + '?text=' + encodeURIComponent(msg);
  if (waColC) sh.getRange(row, waColC).setFormula('=HYPERLINK("' + waUrl + '","📲 Enviar solicitud")');
  return { name: name, waUrl: waUrl };
}

function adminApproveProspect(session, prospectKey) {
  return guard_('adminApproveProspect', function () {
    adminRequire_(session);
    const sh = getSS_().getSheetByName(CFG.SHEETS.PROSPECTS);
    if (!sh) return { ok: false, msg: 'No hay prospectos.' };
    const lock = LockService.getDocumentLock(); try { lock.waitLock(30000); } catch (e) {}
    try {
      const row = adminResolveProspectRow_(sh, prospectKey);
      if (!row) return { ok: false, msg: 'El prospecto cambió o ya no está. Recargá la lista.' };
      const out = aprobarProspectoWeb_(sh, row);
      adminInvalidateCaches_();
      return { ok: true, name: out.name, waUrl: out.waUrl, msg: out.waUrl ? 'Prospecto aprobado. Tocá «Abrir WhatsApp» para enviar la invitación.' : 'Prospecto aprobado, pero no tiene un teléfono válido para WhatsApp.' };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  });
}

function adminRejectProspect(session, prospectKey) {
  return guard_('adminRejectProspect', function () {
    adminRequire_(session);
    const sh = getSS_().getSheetByName(CFG.SHEETS.PROSPECTS);
    if (!sh) return { ok: false, msg: 'No hay prospectos.' };
    const lock = LockService.getDocumentLock(); try { lock.waitLock(30000); } catch (e) {}
    try {
      const row = adminResolveProspectRow_(sh, prospectKey);
      if (!row) return { ok: false, msg: 'El prospecto cambió o ya no está. Recargá la lista.' };
      const m = headerMap_(sh);
      if (m['Estado']) sh.getRange(row, m['Estado']).setValue('RECHAZADO');
      const nm = m['Nombre'] ? String(sh.getRange(row, m['Nombre']).getValue() || '').trim() : '';
      adminInvalidateCaches_();
      return { ok: true, msg: 'Prospecto' + (nm ? ' «' + nm + '»' : '') + ' marcado RECHAZADO.' };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  });
}

/* ========================= HOME / BOOTSTRAP ========================= */

/** Snapshot de inicio con el MÍNIMO de lecturas: Prestatarios UNA sola vez (capital prestado +
 *  activos en el mismo recorrido, sin el doble escaneo que hacían fundStats_ + adminCountActivos_),
 *  Pagos una vez (cobrado), y los contadores de "Nuevos Prestatarios"/"Prospectos". O(n+p+a+b). */
function adminHomeData_() {
  const ss = getSS_();
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  let lent = 0, activos = 0;
  if (bs && bs.getLastRow() >= 2) {
    const H = headerIndex_(bs);
    const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
    const capC = colByAny_(H, ['Capital']) || PB.PRINCIPAL;
    const stC = colByAny_(H, ['Estado']) || PB.STATE;
    const maxC = Math.max(idC, capC, stC);
    const data = bs.getRange(2, 1, bs.getLastRow() - 1, maxC).getValues();
    for (let i = 0; i < data.length; i++) {
      const id = String(data[i][idC - 1] || '').trim();
      if (!/^L-/i.test(id)) continue;
      lent += Number(data[i][capC - 1]) || 0;
      const st = String(data[i][stC - 1] || '').toUpperCase();
      if (st !== ST.PAID && st !== ST.CLEARED) activos++;
    }
  }
  let collected = 0;
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (pg && pg.getLastRow() >= 2) {
    const PH = headerIndex_(pg);
    const montoC = colByAny_(PH, ['Monto Pagado']) || PP.AMOUNT;
    const monts = pg.getRange(2, montoC, pg.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < monts.length; i++) collected += Number(monts[i][0]) || 0;
  }
  const total = fondoTotal_();
  const net = round2_(total - lent + collected);
  return {
    fondo: { total: fmtMoney_(total), lent: fmtMoney_(round2_(lent)), collected: fmtMoney_(round2_(collected)),
             available: fmtMoney_(round2_(Math.max(0, net))), net: fmtMoney_(net) },
    pendientes: adminCountPending_(), prospectos: adminCountProspects_(), activos: activos,
    empresa: companyName_()
  };
}
/** Lectura con caché corta (45 s) del snapshot de inicio. Se invalida en cada acción del panel. */
function adminHomeCached_() {
  const cache = CacheService.getScriptCache();
  try { const v = cache.get('admin:home:v1'); if (v) return JSON.parse(v); } catch (e) {}
  const data = adminHomeData_();
  try { cache.put('admin:home:v1', JSON.stringify(data), 45); } catch (e) {}
  return data;
}
/** Invalida las cachés del panel (inicio + lista de préstamos) tras cualquier escritura. */
function adminInvalidateCaches_() {
  try { CacheService.getScriptCache().removeAll(['admin:home:v1', 'admin:loans:v1']); } catch (e) {}
}

/** Inicio: snapshot (cacheado). */
function adminHome(session) {
  return guard_('adminHome', function () {
    adminRequire_(session);
    return Object.assign({ ok: true }, adminHomeCached_());
  });
}
/** ARRANQUE: valida la sesión Y devuelve el inicio en UNA sola ida y vuelta (elimina el doble
 *  adminHome que hacía el arranque: una llamada para validar + otra en loadHome para pintar). */
function adminBootstrap(session) {
  return guard_('adminBootstrap', function () {
    adminRequire_(session);
    return Object.assign({ ok: true }, adminHomeCached_());
  });
}

/* ========================= PÁGINA (SPA MÓVIL) ========================= */

function adminPageHtml_() {
  const company = esc_(companyName_());
  const honey = honeypotTrapHtml_();
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><base target="_top">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<title>Panel Admin — ${company}</title>
<style>
  :root{--brand:#1c4587;--brand2:#2a5db0;--ok:#188038;--okbg:#e6f4ea;--danger:#c5221f;--dangerbg:#fce8e6;
    --wa:#25D366;--warn:#b06000;--warnbg:#fef7e0;--bg:#eef1f6;--card:#fff;--line:#dfe4ec;--txt:#1f2733;--muted:#6b7684;--radius:16px}
  *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
  body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;background:var(--bg);color:var(--txt);font-size:16px}
  .wrap{max-width:560px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column}
  header{background:var(--brand);color:#fff;padding:14px 16px;position:sticky;top:0;z-index:5;display:flex;align-items:center;justify-content:space-between}
  header h1{font-size:17px;margin:0;font-weight:700}
  header .out{background:rgba(255,255,255,.18);border:0;color:#fff;font-size:13px;padding:7px 11px;border-radius:9px;font-weight:600}
  main{flex:1;padding:14px 14px 92px}
  .view{display:none} .view.on{display:block}
  .card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:16px;margin:0 0 12px;box-shadow:0 1px 4px rgba(20,40,80,.05)}
  h2{font-size:15px;margin:2px 0 12px;color:var(--brand)}
  label{display:block;font-size:13px;font-weight:600;color:var(--muted);margin:10px 0 5px}
  input,select{width:100%;padding:14px;border:1.5px solid var(--line);border-radius:12px;font-size:16px;background:#fff}
  input:focus,select:focus{outline:none;border-color:var(--brand2)}
  .btn{display:block;width:100%;border:0;border-radius:13px;padding:15px;font-size:16px;font-weight:700;color:#fff;background:var(--brand);margin-top:12px;cursor:pointer}
  .btn:active{transform:translateY(1px)}
  .btn.ok{background:var(--ok)} .btn.danger{background:var(--danger)} .btn.ghost{background:#eef2f8;color:var(--brand)}
  .btn.wa{background:var(--wa)} .btn.sm{padding:11px;font-size:14px;margin-top:8px}
  .row{display:flex;gap:8px} .row>*{flex:1}
  .tiles{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .tile{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:18px 14px;text-align:center;cursor:pointer}
  .tile .n{font-size:30px;font-weight:800;color:var(--brand);line-height:1}
  .tile .t{font-size:13px;color:var(--muted);margin-top:7px;font-weight:600}
  .kpi{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px dashed var(--line);font-size:14px}
  .kpi:last-child{border:0} .kpi b{color:var(--brand)}
  .chip{display:inline-block;font-size:12px;font-weight:700;padding:3px 9px;border-radius:999px;margin:2px 4px 2px 0}
  .chip.ok{background:var(--okbg);color:var(--ok)} .chip.bad{background:var(--dangerbg);color:var(--danger)}
  .chip.warn{background:var(--warnbg);color:var(--warn)} .chip.info{background:#e8f0fe;color:var(--brand)}
  .nm{font-size:16px;font-weight:700} .sub{font-size:13px;color:var(--muted);margin-top:2px}
  .mono{font-variant-numeric:tabular-nums}
  .ref{display:flex;align-items:center;justify-content:space-between;gap:8px;border-top:1px solid var(--line);padding:9px 0;font-size:14px}
  .ref a{background:var(--wa);color:#fff;text-decoration:none;font-weight:700;padding:8px 10px;border-radius:10px;font-size:13px;white-space:nowrap}
  .refbtns{display:flex;gap:6px;align-items:center;flex:0 0 auto}
  .refval{background:#eef2f8;color:var(--brand);border:1.5px solid var(--brand);font-weight:700;padding:7px 10px;border-radius:10px;font-size:13px;cursor:pointer;white-space:nowrap}
  .refval.on{background:var(--ok);border-color:var(--ok);color:#fff}
  .cuota{display:flex;justify-content:space-between;font-size:13px;padding:6px 0;border-bottom:1px dashed var(--line)}
  .walist a{display:block;background:var(--wa);color:#fff;text-decoration:none;text-align:center;font-weight:700;padding:13px;border-radius:12px;margin-top:9px;font-size:15px}
  .switch{display:flex;align-items:center;justify-content:space-between;background:var(--warnbg);border:1px solid #f0d9a8;border-radius:12px;padding:12px;margin-top:12px}
  .switch span{font-size:13px;font-weight:700;color:var(--warn)}
  .tgl{position:relative;width:52px;height:30px;flex:0 0 auto}
  .tgl input{opacity:0;width:0;height:0} .tgl i{position:absolute;inset:0;background:#c6ccd6;border-radius:999px;transition:.2s}
  .tgl i:before{content:"";position:absolute;width:24px;height:24px;left:3px;top:3px;background:#fff;border-radius:50%;transition:.2s}
  .tgl input:checked+i{background:var(--warn)} .tgl input:checked+i:before{transform:translateX(22px)}
  nav{position:fixed;bottom:0;left:0;right:0;max-width:560px;margin:0 auto;background:#fff;border-top:1px solid var(--line);display:flex;z-index:6}
  nav button{flex:1;border:0;background:none;padding:9px 2px 10px;font-size:11px;color:var(--muted);font-weight:600;cursor:pointer}
  nav button .ic{display:block;font-size:21px;line-height:1;margin-bottom:3px} nav button.on{color:var(--brand)}
  nav .badge{display:inline-block;min-width:17px;background:var(--danger);color:#fff;border-radius:999px;font-size:10px;padding:0 5px;margin-left:3px}
  .empty{text-align:center;color:var(--muted);padding:28px 10px;font-size:14px}
  .toast{position:fixed;left:50%;bottom:104px;transform:translateX(-50%);background:#1f2733;color:#fff;padding:12px 18px;border-radius:12px;font-size:14px;max-width:90%;text-align:center;opacity:0;transition:.25s;z-index:20;pointer-events:none}
  .toast.on{opacity:1} .toast.bad{background:var(--danger)} .toast.ok{background:var(--ok)}
  .over{position:fixed;inset:0;background:rgba(240,243,248,.75);display:none;align-items:center;justify-content:center;z-index:30}
  .over.on{display:flex} .spin{width:42px;height:42px;border:4px solid #cfd8e6;border-top-color:var(--brand);border-radius:50%;animation:sp 1s linear infinite}
  @keyframes sp{to{transform:rotate(360deg)}}
  .login{min-height:100vh;display:flex;flex-direction:column;justify-content:center;padding:24px}
  .login .logo{text-align:center;font-size:26px;font-weight:800;color:var(--brand);margin-bottom:4px}
  .login .sub{text-align:center;margin-bottom:18px}
  .err{color:var(--danger);font-size:13px;margin-top:10px;text-align:center;min-height:16px}
  .back{background:none;border:0;color:var(--brand);font-weight:700;font-size:14px;padding:4px 0;margin-bottom:6px;cursor:pointer}
</style></head>
<body>
<!-- LOGIN -->
<div id="loginWrap" class="login">
  <div class="logo">${company}</div>
  <div class="sub muted" style="color:var(--muted)">Panel de administración</div>
  <div class="card">
    <label>Usuario</label>
    <input id="lu" autocomplete="username" autocapitalize="none" inputmode="text">
    <label>Contraseña</label>
    <input id="lp" type="password" autocomplete="current-password">
    ${honey}
    <button class="btn" id="loginBtn" onclick="doLogin()">Entrar</button>
    <div class="err" id="loginErr"></div>
  </div>
</div>

<!-- APP -->
<div id="app" class="wrap" style="display:none">
  <header><h1 id="hdr">Inicio</h1><button class="out" onclick="doLogout()">Salir</button></header>
  <main>
    <!-- INICIO -->
    <section class="view" id="v-home">
      <div class="card">
        <h2>Caja</h2>
        <div class="kpi"><span>Disponible para prestar</span><b id="k-av">—</b></div>
        <div class="kpi"><span>Capital prestado</span><span id="k-lent" class="mono">—</span></div>
        <div class="kpi"><span>Total cobrado</span><span id="k-col" class="mono">—</span></div>
        <div class="kpi"><span>Fondo total</span><span id="k-tot" class="mono">—</span></div>
      </div>
      <div class="tiles">
        <div class="tile" onclick="go('aprobar')"><div class="n" id="t-pend">0</div><div class="t">Para aprobar</div></div>
        <div class="tile" onclick="go('pagos')"><div class="n">＄</div><div class="t">Registrar pago</div></div>
        <div class="tile" onclick="go('prestamos')"><div class="n" id="t-act">0</div><div class="t">Préstamos</div></div>
        <div class="tile" onclick="go('prospectos')"><div class="n" id="t-prosp">0</div><div class="t">Prospectos</div></div>
      </div>
    </section>

    <!-- APROBAR -->
    <section class="view" id="v-aprobar">
      <h2>Solicitudes para aprobar</h2>
      <div id="pendList"></div>
    </section>

    <!-- PAGOS -->
    <section class="view" id="v-pagos">
      <div id="payPick">
        <h2>Registrar un pago</h2>
        <div class="card">
          <label>Buscá el préstamo (nombre, DNI o código)</label>
          <input id="paySearch" inputmode="search" oninput="paySearch()" placeholder="Ej: Nancy, 29013897, L-0015">
          <div id="payResults" style="margin-top:8px"></div>
        </div>
      </div>
      <div id="payForm" style="display:none">
        <button class="back" onclick="payBack()">‹ Elegir otro préstamo</button>
        <div class="card">
          <div class="nm" id="pf-name">—</div><div class="sub" id="pf-sub">—</div>
          <div class="kpi" style="margin-top:10px"><span>Saldo a cancelar</span><b id="pf-saldo">—</b></div>
          <label>Monto del pago</label>
          <input id="pf-amount" inputmode="decimal" placeholder="0">
          <label>Fecha</label>
          <input id="pf-date" type="date">
          <button class="btn ok" onclick="doPay()">Registrar pago</button>
        </div>
      </div>
    </section>

    <!-- PRESTAMOS -->
    <section class="view" id="v-prestamos">
      <div id="loanPick">
        <h2>Préstamos</h2>
        <div class="card">
          <input id="loanSearch" inputmode="search" oninput="loanSearch()" placeholder="Buscar por nombre, DNI o código">
          <div id="loanResults" style="margin-top:10px"></div>
        </div>
      </div>
      <div id="loanDetail" style="display:none"></div>
    </section>

    <!-- VENCIMIENTOS (Estudio de Recordatorios) -->
    <section class="view" id="v-vencimientos">
      <h2>Vencimientos — vencidos y por vencer</h2>
      <div id="remList"></div>
    </section>

    <!-- PROSPECTOS -->
    <section class="view" id="v-prospectos">
      <h2>Prospectos</h2>
      <div id="prospList"></div>
    </section>
  </main>

  <nav>
    <button data-tab="home" class="on" onclick="go('home')"><span class="ic">🏠</span>Inicio</button>
    <button data-tab="aprobar" onclick="go('aprobar')"><span class="ic">✅</span>Aprobar<span id="nb-pend" class="badge" style="display:none">0</span></button>
    <button data-tab="pagos" onclick="go('pagos')"><span class="ic">＄</span>Pagos</button>
    <button data-tab="prestamos" onclick="go('prestamos')"><span class="ic">📄</span>Préstamos</button>
    <button data-tab="vencimientos" onclick="go('vencimientos')"><span class="ic">⏰</span>Vence<span id="nb-venc" class="badge" style="display:none">0</span></button>
    <button data-tab="prospectos" onclick="go('prospectos')"><span class="ic">🌱</span>Prospectos<span id="nb-prosp" class="badge" style="display:none">0</span></button>
  </nav>
</div>

<div class="toast" id="toast"></div>
<div class="over" id="over"><div class="spin"></div></div>

<script>
var SESS = sessionStorage.getItem('impAdminSess') || '';
var CLIENT_IP = '';
var loanCache = [];     // listado completo de préstamos (cacheado; búsqueda/filtrado en el navegador)
var loansLoaded = false;
var homeFresh = false;  // true = el inicio ya se pintó desde el bootstrap/login (loadHome no re-consulta)
var tDebounce = null;

function $(id){return document.getElementById(id);}
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function show(el,on){el.style.display=on?'':'none';}
function overlay(on){$('over').classList.toggle('on',!!on);}
function toast(msg,kind){var t=$('toast');t.textContent=msg;t.className='toast on'+(kind?(' '+kind):'');setTimeout(function(){t.className='toast';},kind==='bad'?4200:2600);}

// Envoltura de google.script.run: inyecta la sesión, maneja el spinner y el fin de sesión.
function call(fn,args,onOk,opts){
  opts=opts||{};
  if(!opts.quiet) overlay(true);
  google.script.run
    .withSuccessHandler(function(r){ if(!opts.quiet) overlay(false);
      if(r && r.ok===false && /Sesión vencida/i.test(r.msg||'')){ sessionLost(); return; }
      onOk(r);
    })
    .withFailureHandler(function(e){ if(!opts.quiet) overlay(false);
      var m=(e&&e.message)||String(e);
      if(/Sesión vencida/i.test(m)){ sessionLost(); return; }
      toast(m,'bad');
    })[fn].apply(google.script.run,[SESS].concat(args||[]));
}
function sessionLost(){ SESS=''; sessionStorage.removeItem('impAdminSess'); overlay(false); showLogin(); toast('Tu sesión venció. Entrá de nuevo.','bad'); }

/* ---------- IP best-effort (para el anti-fuerza-bruta) ---------- */
(function(){ try{ fetch('https://api.ipify.org?format=json').then(function(r){return r.json();}).then(function(j){CLIENT_IP=(j&&j.ip)||'';}).catch(function(){}); }catch(e){} })();

/* ---------- LOGIN ---------- */
function showLogin(){ $('app').style.display='none'; $('loginWrap').style.display='flex'; }
function showApp(){ $('loginWrap').style.display='none'; $('app').style.display='flex'; go('home'); }
function doLogin(){
  var u=$('lu').value.trim(), p=$('lp').value, trap=($('website')||{}).value||'';
  if(trap){ return; } // honeypot
  if(!u||!p){ $('loginErr').textContent='Completá usuario y contraseña.'; return; }
  $('loginErr').textContent=''; overlay(true);
  google.script.run.withSuccessHandler(function(r){ overlay(false);
    if(r&&r.ok){ SESS=r.session; sessionStorage.setItem('impAdminSess',SESS); $('lp').value='';
      if(r.home){ paintHome(r.home); homeFresh=true; } showApp(); prefetchLoans(); }
    else { $('loginErr').textContent=(r&&r.msg)||'No se pudo entrar.'; }
  }).withFailureHandler(function(e){ overlay(false); $('loginErr').textContent=(e&&e.message)||String(e); })
    .adminLogin(u,p,CLIENT_IP);
}
function doLogout(){ call('adminLogout',[],function(){ sessionLost(); }); }

/* ---------- NAVEGACIÓN ---------- */
var TITLES={home:'Inicio',aprobar:'Aprobar',pagos:'Registrar pago',prestamos:'Préstamos',vencimientos:'Vencimientos',prospectos:'Prospectos'};
function go(tab){
  ['home','aprobar','pagos','prestamos','vencimientos','prospectos'].forEach(function(t){ $('v-'+t).classList.toggle('on',t===tab); });
  var nav=document.querySelectorAll('nav button'); for(var i=0;i<nav.length;i++){ nav[i].classList.toggle('on',nav[i].getAttribute('data-tab')===tab); }
  $('hdr').textContent=TITLES[tab]||'';
  if(tab==='home') loadHome();
  if(tab==='aprobar') loadPending();
  if(tab==='pagos') payEnter();
  if(tab==='prestamos') loanEnter();
  if(tab==='vencimientos') loadReminders();
  if(tab==='prospectos') loadProspects();
}

/* ---------- caché de préstamos (una lectura; búsqueda local = instantánea) ---------- */
function loadLoans(cb,quiet){ if(loansLoaded){ cb(); return; } call('adminListLoans',[''],function(r){ if(r&&r.ok){ loanCache=r.items; loansLoaded=true; } cb(); },{quiet:!!quiet}); }
function invalidateLoans(){ loansLoaded=false; }
// Precarga la lista de préstamos en segundo plano (sin spinner) cuando el navegador está ocioso,
// así Pagos/Préstamos abren sin ida y vuelta visible. No bloquea el pintado del inicio.
function prefetchLoans(){ try{ if(window.requestIdleCallback) requestIdleCallback(function(){ loadLoans(function(){},true); },{timeout:2500}); else setTimeout(function(){ loadLoans(function(){},true); },600); }catch(e){} }
function diasTxt(d){ if(d===null||d===undefined) return ''; if(d>0) return 'vence en '+d+' día'+(d===1?'':'s'); if(d===0) return 'vence hoy'; return 'vencido hace '+(-d)+' día'+(d===-1?'':'s'); }
function diasCls(d){ if(d===null||d===undefined) return 'info'; if(d<0) return 'bad'; if(d<=3) return 'warn'; return 'ok'; }
function renderLoanCard(it,mode){
  var oc=(mode==='pay'?'payOpen':'loanOpen');
  var estCls=/VENCIDO/i.test(it.estado)?'bad':(/PAGADO|SALDADO/i.test(it.estado)?'ok':(/PLAN|CONGEL/i.test(it.estado)?'warn':'info'));
  var dt=it.venc?('<span class="chip '+diasCls(it.dias)+'">'+esc(it.venc)+(it.dias!==null&&it.dias!==undefined?(' · '+diasTxt(it.dias)):'')+'</span>'):'';
  return '<div class="card" style="padding:12px;cursor:pointer" onclick="'+oc+'(\\''+it.loanId+'\\')">'+
    '<div class="nm">'+esc(it.nombre||it.loanId)+'</div>'+
    '<div class="sub">'+esc(it.loanId)+' · DNI '+esc(it.dni||'—')+'</div>'+
    '<div style="margin-top:6px"><span class="chip '+estCls+'">'+esc(it.estado)+'</span><span class="chip info">saldo '+esc(it.saldo)+'</span>'+dt+'</div></div>';
}

/* ---------- INICIO ---------- */
function paintHome(r){ if(!r||!r.fondo) return;
  $('k-av').textContent=r.fondo.available; $('k-lent').textContent=r.fondo.lent;
  $('k-col').textContent=r.fondo.collected; $('k-tot').textContent=r.fondo.total;
  $('t-pend').textContent=r.pendientes; $('t-act').textContent=r.activos; $('t-prosp').textContent=r.prospectos;
  badge('nb-pend',r.pendientes); badge('nb-prosp',r.prospectos);
}
function loadHome(){ if(homeFresh){ homeFresh=false; return; }   // ya pintado por el bootstrap/login
  call('adminHome',[],paintHome,{quiet:true}); }
function badge(id,n){ var b=$(id); if(!b) return; b.textContent=n; b.style.display=n>0?'':'none'; }

/* ---------- APROBAR ---------- */
function loadPending(){
  call('adminListPending',[],function(r){
    var box=$('pendList'); if(!r||!r.ok){box.innerHTML='';return;}
    badge('nb-pend',r.items.length); $('t-pend').textContent=r.items.length;
    if(!r.items.length){ box.innerHTML='<div class="empty">No hay solicitudes pendientes. 🎉</div>'; return; }
    box.innerHTML=r.items.map(renderPend).join('');
  });
}
function renderPend(it){
  var d=it.bcra||{}, dec=(d.decision||'').toUpperCase();
  var bcra = dec? '<span class="chip '+(dec.indexOf('RECHAZ')>=0?'bad':(dec.indexOf('REVIS')>=0?'warn':'ok'))+'">BCRA: '+esc(d.decision)+'</span>':'';
  var extra = (d.peorSit?('<span class="chip info">Peor sit.: '+esc(d.peorSit)+'</span>'):'')+(d.pctAlDia?('<span class="chip info">'+esc(d.pctAlDia)+' al día</span>'):'');
  var terc = it.tercero?'<span class="chip warn">Cuenta de tercero</span>':'';
  var refs = (it.refs||[]).map(function(rf){
    return '<div class="ref"><div>'+esc(rf.nombre||('Ref '+rf.idx))+' <span class="sub">'+esc(rf.vinculo||'')+'</span></div>'+
      '<div class="refbtns">'+
        (rf.telefono?'<a href="#" onclick="refLink(\\''+it.appKey+'\\','+rf.idx+');return false">WA</a>':'')+
        '<button id="rv-'+it.appKey+'-'+rf.idx+'" class="refval'+(rf.validada?' on':'')+'" onclick="refValidate(this,\\''+it.appKey+'\\','+rf.idx+')">'+(rf.validada?'✔ Validada':'Validar')+'</button>'+
      '</div></div>';
  }).join('');
  return '<div class="card">'+
    '<div class="nm">'+esc(it.nombre)+'</div>'+
    '<div class="sub">DNI '+esc(it.dni||'—')+' · '+esc(it.telefono||'sin tel.')+'</div>'+
    '<div style="margin:8px 0"><span class="chip info">'+esc(it.monto)+'</span><span class="chip info">'+esc(it.plazo||'?')+' (plazo)</span>'+terc+'</div>'+
    '<div style="margin-bottom:4px">'+bcra+extra+'</div>'+
    (it.notas?'<div class="sub">📝 '+esc(it.notas)+'</div>':'')+
    (refs?('<div style="margin-top:6px">'+refs+'</div>'):'')+
    '<div class="switch"><span>⚠ Anular límites (aprobar igual)</span><label class="tgl"><input type="checkbox" id="ov-'+it.appKey+'"><i></i></label></div>'+
    '<div class="row"><button class="btn ok" onclick="doApprove(\\''+it.appKey+'\\')">Aprobar</button>'+
    '<button class="btn danger" onclick="doReject(\\''+it.appKey+'\\')">Rechazar</button></div>'+
  '</div>';
}
function refLink(appKey,idx){
  call('adminPendingRefLinks',[appKey],function(r){
    if(!r||!r.ok){ toast(r&&r.msg||'No se pudo.','bad'); return; }
    var link=(r.links||[]).find(function(l){return l.label.indexOf('Ref '+idx)===0;});
    if(link&&link.url) window.open(link.url,'_blank'); else toast('Esa referencia no tiene teléfono válido.','bad');
  });
}
function refValidate(btn,appKey,idx){
  var cur=btn.classList.contains('on');
  call('adminValidateRef',[appKey,idx,!cur],function(r){
    if(r&&r.ok){ btn.classList.toggle('on',r.validada); btn.textContent=r.validada?'✔ Validada':'Validar'; toast(r.validada?'Referencia validada ✔':'Validación quitada','ok'); }
    else toast((r&&r.msg)||'No se pudo.','bad');
  },{quiet:true});
}
function doApprove(appKey){
  var ov=($('ov-'+appKey)||{}).checked;
  if(!confirm(ov?'Vas a APROBAR ANULANDO los límites. ¿Seguro?':'¿Aprobar esta solicitud? Se creará el préstamo y se enviará el contrato para firmar.')) return;
  toast('Aprobando… puede tardar hasta un minuto.');
  call('adminApprove',[appKey,!!ov],function(r){
    if(r&&r.ok){ toast('✅ '+r.msg,'ok'); invalidateLoans(); loadPending(); } else { toast((r&&r.msg)||'No se pudo aprobar.','bad'); }
  });
}
function doReject(appKey){
  if(!confirm('¿Rechazar esta solicitud? Se archiva y se le avisa por correo.')) return;
  call('adminReject',[appKey],function(r){ if(r&&r.ok){ toast('🚫 '+r.msg,'ok'); loadPending(); } else toast((r&&r.msg)||'No se pudo.','bad'); });
}

/* ---------- PAGOS ---------- */
function payEnter(){ show($('payForm'),false); show($('payPick'),true); $('paySearch').value=''; loadLoans(function(){ renderPayList(''); }); }
function payBack(){ payEnter(); }
function paySearch(){ renderPayList($('paySearch').value.trim()); }   // filtra localmente = instantáneo
function renderPayList(q){
  q=(q||'').toLowerCase();
  var list=loanCache.filter(function(it){ return it.payable && (!q || (it.loanId+' '+(it.nombre||'')+' '+(it.dni||'')).toLowerCase().indexOf(q)>=0); });
  $('payResults').innerHTML = list.length ? list.map(function(it){ return renderLoanCard(it,'pay'); }).join('')
    : '<div class="empty">'+(q?'Sin resultados.':'No hay préstamos con saldo para pagar.')+'</div>';
}
function payOpen(loanId){
  call('adminLoanDetail',[loanId],function(r){
    if(!r||!r.ok){ toast((r&&r.msg)||'No se pudo.','bad'); return; }
    var l=r.loan;
    $('pf-name').textContent=l.nombre||l.loanId; $('pf-sub').textContent=l.loanId+' · '+l.estado;
    $('pf-saldo').textContent=l.toClear;
    $('pf-amount').value=''; $('pf-date').value=today();
    $('payForm').setAttribute('data-loan',loanId);
    show($('payPick'),false); show($('payForm'),true);
  });
}
function doPay(){
  var loanId=$('payForm').getAttribute('data-loan'); var amt=$('pf-amount').value.trim(); var date=$('pf-date').value;
  if(!amt||Number(amt.replace(/[^0-9.]/g,''))<=0){ toast('Ingresá un monto válido.','bad'); return; }
  if(!confirm('¿Registrar un pago de '+amt+'?')) return;
  call('adminRecordPayment',[loanId,amt,date],function(r){
    if(r&&r.ok){ toast('✅ '+r.msg,'ok'); invalidateLoans();
      if(r.waUrl && confirm('Pago registrado. ¿Abrir WhatsApp para avisarle al cliente?')) window.open(r.waUrl,'_blank');
      payBack();
    } else toast((r&&r.msg)||'No se pudo registrar.','bad');
  });
}

/* ---------- PRESTAMOS ---------- */
function loanEnter(){ show($('loanDetail'),false); show($('loanPick'),true); $('loanSearch').value=''; loadLoans(function(){ renderLoanList(''); }); }
function loanBack(){ loanEnter(); }
function loanSearch(){ renderLoanList($('loanSearch').value.trim()); }   // filtra localmente = instantáneo
function renderLoanList(q){
  q=(q||'').toLowerCase();
  var list=loanCache.filter(function(it){ return !q || (it.loanId+' '+(it.nombre||'')+' '+(it.dni||'')).toLowerCase().indexOf(q)>=0; });
  $('loanResults').innerHTML = list.length ? list.map(function(it){ return renderLoanCard(it,'open'); }).join('') : '<div class="empty">Sin resultados.</div>';
}
function loanOpen(loanId){
  call('adminLoanDetail',[loanId],function(r){
    if(!r||!r.ok){ toast((r&&r.msg)||'No se pudo.','bad'); return; }
    var l=r.loan;
    var cuotas=(r.cuotas||[]).map(function(c){
      var st=/PAGADA/i.test(c.estado)?'ok':(/VENCIDA/i.test(c.estado)?'bad':'warn');
      return '<div class="cuota"><span>Cuota '+c.num+' <span class="sub">'+esc(c.venc)+'</span></span><span><span class="chip '+st+'">'+esc(c.estado)+'</span> '+esc(c.saldo)+'</span></div>';
    }).join('');
    var wa=(r.waLinks||[]).map(function(w){ return '<a href="'+esc(w.url)+'" target="_blank">'+esc(w.label)+'</a>'; }).join('');
    $('loanDetail').innerHTML =
      '<button class="back" onclick="loanBack()">‹ Volver a la lista</button>'+
      '<div class="card"><div class="nm">'+esc(l.nombre||l.loanId)+'</div><div class="sub">'+esc(l.loanId)+' · DNI '+esc(l.dni||'—')+' · '+esc(l.telefono||'sin tel.')+'</div>'+
        '<div style="margin:8px 0"><span class="chip '+(l.paidOff?'ok':(l.late?'bad':'info'))+'">'+(l.paidOff?'SALDADO':(l.frozen?'CONGELADO':(l.onPlan?'EN PLAN':(l.late?'VENCIDO':'ACTIVO'))))+'</span>'+(l.venc?('<span class="chip info">vence '+esc(l.venc)+'</span>'):'')+'</div>'+
        '<div class="kpi"><span>Capital</span><span class="mono">'+esc(l.capital)+'</span></div>'+
        '<div class="kpi"><span>Total a pagar</span><span class="mono">'+esc(l.total)+'</span></div>'+
        '<div class="kpi"><span>Pagado</span><span class="mono">'+esc(l.pagado)+'</span></div>'+
        '<div class="kpi"><span>Saldo</span><span class="mono">'+esc(l.saldo)+'</span></div>'+
        (l.mora&&l.mora!=='$0.00'?'<div class="kpi"><span>Mora</span><span class="mono">'+esc(l.mora)+'</span></div>':'')+
        '<div class="kpi"><span>A cancelar hoy</span><b>'+esc(l.toClear)+'</b></div>'+
      '</div>'+
      (cuotas?('<div class="card"><h2>Cuotas</h2>'+cuotas+'</div>'):'')+
      '<div class="card walist"><h2>Avisar por WhatsApp</h2>'+(wa||'<div class="sub">Sin teléfono o sin acciones disponibles.</div>')+
        '<button class="btn ghost sm" onclick="go(\\'pagos\\');setTimeout(function(){payOpen(\\''+l.loanId+'\\')},150)">Registrar un pago de este préstamo</button></div>';
    show($('loanPick'),false); show($('loanDetail'),true);
  });
}

/* ---------- VENCIMIENTOS (Estudio de Recordatorios) ---------- */
function loadReminders(){
  call('adminReminders',[],function(r){
    var box=$('remList'); if(!r||!r.ok){ box.innerHTML=''; return; }
    var venc=r.items.filter(function(it){return it.dias!==null&&it.dias<=0;}).length;
    badge('nb-venc',venc);
    if(!r.items.length){ box.innerHTML='<div class="empty">No hay préstamos por vencer ni vencidos. 🎉</div>'; return; }
    box.innerHTML=r.items.map(function(it){
      return '<div class="card" style="padding:12px">'+
        '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px">'+
          '<div><div class="nm">'+esc(it.nombre||it.loanId)+'</div><div class="sub">'+esc(it.loanId)+(it.venc?(' · vence '+esc(it.venc)):'')+' · saldo '+esc(it.saldo)+'</div></div>'+
          '<span class="chip '+diasCls(it.dias)+'" style="white-space:nowrap">'+esc(diasTxt(it.dias))+'</span>'+
        '</div>'+
        ((it.waLinks||[]).length
          ? (it.waLinks||[]).map(function(w){ return '<a class="btn wa sm" href="'+esc(w.url)+'" target="_blank">'+esc(w.label)+'</a>'; }).join('')
          : '<div class="sub" style="margin-top:6px">Sin teléfono para WhatsApp.</div>')+
      '</div>';
    }).join('');
  });
}

/* ---------- PROSPECTOS ---------- */
function loadProspects(){
  call('adminListProspects',[false],function(r){
    var box=$('prospList'); if(!r||!r.ok){box.innerHTML='';return;}
    badge('nb-prosp',r.items.length); $('t-prosp').textContent=r.items.length;
    if(!r.items.length){ box.innerHTML='<div class="empty">No hay prospectos nuevos.</div>'; return; }
    box.innerHTML=r.items.map(function(it){
      return '<div class="card"><div class="nm">'+esc(it.nombre)+'</div><div class="sub">CUIL '+esc(it.cuil||'—')+' · '+esc(it.telefono||'sin tel.')+'</div>'+
        '<div style="margin:6px 0"><span class="chip info">'+esc(it.estado)+'</span>'+(it.contactado?'<span class="chip ok">contactado</span>':'')+'</div>'+
        (it.nota?'<div class="sub">📝 '+esc(it.nota)+'</div>':'')+
        '<div class="row"><button class="btn ok" onclick="doProsp(\\''+it.prospectKey+'\\')">Aprobar y avisar</button>'+
        '<button class="btn danger" onclick="doProspRej(\\''+it.prospectKey+'\\')">Rechazar</button></div></div>';
    }).join('');
  });
}
function doProsp(key){
  if(!confirm('¿Aprobar este prospecto y prepararle la invitación por WhatsApp?')) return;
  call('adminApproveProspect',[key],function(r){
    if(r&&r.ok){ toast('✅ '+r.msg,'ok'); if(r.waUrl) window.open(r.waUrl,'_blank'); loadProspects(); }
    else toast((r&&r.msg)||'No se pudo.','bad');
  });
}
function doProspRej(key){
  if(!confirm('¿Rechazar este prospecto?')) return;
  call('adminRejectProspect',[key],function(r){ if(r&&r.ok){ toast('🚫 '+r.msg,'ok'); loadProspects(); } else toast((r&&r.msg)||'No se pudo.','bad'); });
}

/* ---------- util ---------- */
function today(){ var d=new Date(); var m=('0'+(d.getMonth()+1)).slice(-2), day=('0'+d.getDate()).slice(-2); return d.getFullYear()+'-'+m+'-'+day; }

/* ---------- arranque ---------- */
(function(){
  if(SESS){ call('adminBootstrap',[],function(r){ if(r&&r.ok){ paintHome(r); homeFresh=true; showApp(); prefetchLoans(); } else { showLogin(); } }); }
  else showLogin();
  $('lp').addEventListener('keydown',function(e){ if(e.key==='Enter') doLogin(); });
})();
</script>
</body></html>`;
}
