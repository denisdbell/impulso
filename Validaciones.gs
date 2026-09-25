/************************************************************
 * VALIDACIONES v2  —  Complemento de LoanManagerV2.gs
 * ----------------------------------------------------------
 * Implementa el catálogo de validaciones (V-01 … V-24) de la
 * hoja "impulso-datos-corregidos" (migrada y normalizada) y
 * agrega un FORMULARIO INTELIGENTE que:
 *   • Detecta clientes existentes al ingresar Correo + DNI y
 *     precarga sus datos (entrada rápida).
 *   • Exige TODOS los campos a los clientes nuevos.
 *
 * Se pega COMO ARCHIVO ADICIONAL en el mismo proyecto de Apps
 * Script que LoanManagerV2.gs. Reutiliza sus helpers globales
 * (getSS_, guard_, esc_, normDni_, normEmail_, nbAppend_,
 * savePhoto_, hasFile_, borrowerFolder_, sendBrandedEmail_,
 * fmtMoney_, companyName_, brandHeaderHtml_, fundAvailable_,
 * nextClienteId_, firstEmptyClienteRow_, etc.) — NO los
 * redefine para evitar funciones duplicadas.
 *
 * ESQUEMA "Clientes" (normalizado, 11 columnas):
 *   ID Cliente | Nombre | Correo | DNI | Teléfono |
 *   DNI (norm) | Correo (norm) | Teléfono (norm) |
 *   Correo válido | Duplicado | Nota
 *
 * PUESTA EN MARCHA
 *   1) Ejecute  ->  configurarValidaciones   (crea la hoja
 *      "Validaciones", instala validaciones de datos en las
 *      hojas y normaliza los plazos a 15/30/60 días).
 *   2) Publique el formulario: Implementar ▸ Aplicación web.
 *      La URL servirá el formulario inteligente (doGet ya
 *      enruta a intakeSmartHtml_).
 ************************************************************/

/* ============================ CONFIG ============================ */

// Tasas derivadas del PLAZO EN DÍAS (modelo migrado): 15→25%, 30→50%, 60/90→100%.
// 90 días (3 meses) es el tramo "grande": misma tasa 100% que 60, pero se paga en 3 cuotas.
const TERM_DAYS_RATES = { 15: 0.25, 30: 0.5, 60: 1, 90: 1 };

// V-10: tope de concentración por cliente. Valor por defecto si no está configurado.
const CLIENT_CONCENTRATION_RATE = 0.10;
// Clave del ajuste (Configuración) que define el tope como % del capital disponible.
const CONCENTRATION_SETTING_KEY = 'Tope de concentración por prestatario (% del capital disponible)';

/**
 * Tope de concentración por prestatario como FRACCIÓN (0…1) del capital disponible.
 * Lee el ajuste de "Configuración" (se ingresa como porcentaje, p. ej. 10, 25, 100).
 *   • vacío / no numérico → valor por defecto (10%).
 *   • 0 (o negativo)      → sin tope (se desactiva el chequeo).
 */
function concentrationRate_() {
  const raw = String(getSetting_(CONCENTRATION_SETTING_KEY) || '').trim();
  const cleaned = raw.replace(/[^0-9.\-]/g, '');
  if (cleaned === '') return CLIENT_CONCENTRATION_RATE;
  const n = Number(cleaned);
  if (!isFinite(n)) return CLIENT_CONCENTRATION_RATE;
  if (n <= 0) return 0; // sin tope
  return n / 100;
}

// Máximo de préstamos simultáneos por prestatario.
const MAX_LOANS_PER_CLIENT = 2;

// Encabezados del esquema "Clientes" normalizado.
const CLIENTES_HEADERS_V2 = ['ID Cliente', 'Nombre', 'Correo', 'DNI', 'Teléfono',
  'DNI (norm)', 'Correo (norm)', 'Teléfono (norm)', 'Correo válido', 'Duplicado', 'Nota', 'CUIL', 'Dirección',
  'Ref 1 Nombre', 'Ref 1 Vínculo', 'Ref 1 Teléfono', 'Ref 2 Nombre', 'Ref 2 Vínculo', 'Ref 2 Teléfono'];

// Prefijos de CUIL/CUIT válidos para personas físicas.
const CUIL_PREFIXES = ['20', '23', '24', '27'];

// Catálogo de validaciones (espejo de la hoja "Validaciones" del sheet migrado).
const VALIDATION_RULES = [
  ['V-01', 'Formulario', 'DNI', 'Obligatorio. Solo dígitos tras quitar puntos/espacios. Longitud 7 u 8.', 'Ingresá un DNI válido (7 u 8 dígitos, sin puntos).', 'BLOQUEA', 'ALTA'],
  ['V-02', 'Formulario', 'DNI', 'No puede existir ya con OTRO nombre en Clientes.', 'Ese DNI ya está registrado a nombre de otra persona.', 'BLOQUEA', 'ALTA'],
  ['V-03', 'Formulario', 'CUIL', '11 dígitos, prefijo 20/23/24/27, dígito verificador módulo 11 correcto, y los 8 centrales deben coincidir con el DNI.', 'El CUIL no es válido o no corresponde al DNI ingresado.', 'BLOQUEA', 'ALTA'],
  ['V-04', 'Formulario', 'Correo', 'Formato válido CON dominio de nivel superior. Rechazar "@gmail" sin ".com".', 'El correo no es válido. Verificá que termine en .com, .ar, etc.', 'BLOQUEA', 'ALTA'],
  ['V-05', 'Formulario', 'Correo', 'Si el correo normalizado ya existe con otro DNI → marcar como posible duplicado.', 'Este correo ya figura para otro solicitante.', 'ADVIERTE', 'ALTA'],
  ['V-06', 'Formulario', 'Teléfono', 'Normalizar a E.164 (+54…). 10 dígitos tras el código de área.', 'Ingresá un teléfono válido de 10 dígitos.', 'BLOQUEA', 'MEDIA'],
  ['V-07', 'Formulario', 'Teléfono', 'Si el teléfono normalizado ya existe con otro DNI → marcar como posible duplicado.', 'Este teléfono ya figura para otro solicitante.', 'ADVIERTE', 'ALTA'],
  ['V-08', 'Formulario', 'Fotos DNI', 'Frente y dorso obligatorios.', 'Subí ambas fotos del DNI, nítidas y completas.', 'BLOQUEA', 'MEDIA'],
  ['V-09', 'Aprobación', 'Monto', 'Monto solicitado ≤ Efectivo disponible. Nunca permitir efectivo negativo.', 'No hay fondos suficientes para aprobar este préstamo.', 'BLOQUEA', 'ALTA'],
  ['V-10', 'Aprobación', 'Monto', 'Exposición total del cliente ≤ 10% del fondo.', 'Este cliente superaría el tope de concentración del 10%.', 'BLOQUEA', 'ALTA'],
  ['V-11', 'Aprobación', 'BCRA', 'Verificación BCRA obligatoria. Si Decisión = RECHAZAR → no aprobar.', 'El BCRA marca RECHAZAR. Se requiere autorización expresa.', 'BLOQUEA', 'ALTA'],
  ['V-12', 'Aprobación', 'Duplicado', 'Si comparte correo o teléfono con otro cliente, sumar ambas exposiciones para V-10.', 'Contacto compartido detectado: se agrupan las exposiciones.', 'BLOQUEA', 'ALTA'],
  ['V-13', 'Aprobación', 'Plazo', 'Plazo válido: 15, 30, 60 o 90 días. Tasa derivada del plazo, no se carga a mano.', 'Plazo inválido.', 'BLOQUEA', 'MEDIA'],
  ['V-14', 'Cálculo', 'Días de atraso', 'Si Estado = PAGADO → días de atraso = MAX(0, Fecha de Pago − Vencimiento).', '—', 'AUTOMÁTICO', 'ALTA'],
  ['V-15', 'Cálculo', 'Mora', 'No acumular mora sobre préstamos con saldo cero.', '—', 'AUTOMÁTICO', 'ALTA'],
  ['V-16', 'Cálculo', 'Estado', 'Estado se DERIVA del saldo y la fecha. Nunca se escribe a mano.', '—', 'AUTOMÁTICO', 'ALTA'],
  ['V-17', 'Cálculo', 'Total Pagado', 'Total Pagado se calcula con SUMIF sobre la hoja Pagos.', '—', 'AUTOMÁTICO', 'ALTA'],
  ['V-18', 'Recordatorios', 'Filtro', 'Excluir del listado todo préstamo con Estado = PAGADO.', '—', 'AUTOMÁTICO', 'ALTA'],
  ['V-19', 'Integridad', 'Clientes', 'No crear cliente sin nombre Y sin DNI.', '—', 'BLOQUEA', 'ALTA'],
  ['V-20', 'Integridad', 'Préstamos', 'Todo préstamo debe tener un ID Cliente existente en Clientes.', '—', 'BLOQUEA', 'ALTA'],
  ['V-21', 'Integridad', 'Pagos', 'SUM(Pagos por préstamo) ≤ Total a Pagar del préstamo.', 'El pago excede el saldo pendiente.', 'BLOQUEA', 'MEDIA'],
  ['V-22', 'Integridad', 'Migración', 'Al migrar, comparar cantidad de filas origen vs destino por hoja y abortar si difiere.', '—', 'BLOQUEA', 'ALTA'],
  ['V-23', 'Desembolso', 'Firma', 'Advertir si se desembolsa con Estado de Firma = PENDIENTE.', 'El contrato no está firmado. ¿Desembolsar igual?', 'ADVIERTE', 'ALTA'],
  ['V-24', 'Envío', 'Correo', 'Antes de enviar, verificar Correo válido = SÍ. Si no, registrar en Errores y avisar.', 'No se puede enviar: el correo del cliente es inválido.', 'BLOQUEA', 'ALTA'],
  ['V-26', 'Formulario', 'Compromiso de pago', 'Acuse explícito obligatorio: devolución total al vencimiento, recargo por mora, acciones legales ante el impago y comunicación/plan de pago ante atrasos (condición para préstamos futuros).', 'Debe aceptar los compromisos de pago para enviar la solicitud.', 'BLOQUEA', 'ALTA'],
  ['V-27', 'Formulario', 'Referencias', 'Dos referencias obligatorias (1.ª familiar no conviviente, 2.ª no familiar) con nombre, vínculo y teléfono válido.', 'Completá las dos referencias con un teléfono válido.', 'BLOQUEA', 'MEDIA'],
  ['V-28', 'Formulario', 'Referencias', 'Teléfono de referencia igual al del solicitante o repetido entre referencias → marcar para revisión.', 'Verificá las referencias: el teléfono está repetido o es el del solicitante.', 'ADVIERTE', 'MEDIA'],
  ['V-29', 'Formulario', 'Monto', 'El monto no puede superar el tope configurable (por defecto $500.000).', 'El monto máximo por préstamo es $500.000.', 'BLOQUEA', 'ALTA'],
  ['V-30', 'Formulario', 'Monto/Plazo', 'El plazo/tasa/cuotas se DERIVAN del monto: ≤$150.000→elige 15d/25% o 30d/50% (1 cuota); ≤$300.000→30d/50%/1; >$300.000→90d/100%/3 cuotas.', 'El plazo se calcula según el monto.', 'BLOQUEA', 'ALTA'],
  ['V-31', 'Alta', 'Cuotas', 'Préstamo grande (>$300.000): se genera un cronograma de 3 cuotas mensuales (día 30/60/90), cada una = Total÷3.', '—', 'AUTOMÁTICO', 'ALTA'],
  ['V-32', 'Aprobación', 'Monto', 'Escalera de graduación: el monto ≤ límite por historial de repago del prestatario. Un atraso REINICIA la escalera: solo cuentan los préstamos saldados a tiempo DESPUÉS del último atraso. Mora vigente → límite inicial.', 'Supera el límite por historial del prestatario. Se recupera saldando préstamos a tiempo (aun después de un atraso), o con «Anular límites».', 'BLOQUEA', 'ALTA'],
  ['V-33', 'Formulario y Aprobación', 'Cliente', 'Cliente BLOQUEADO (columna «Bloqueado» = SÍ en Clientes): se rechaza toda solicitud y aprobación que coincida por correo, DNI, CUIL o teléfono. ABSOLUTO: no se anula con «Anular límites». Desbloquear = vaciar la columna «Bloqueado».', 'No es posible procesar solicitudes para este cliente. Ante cualquier duda, comunicate con el prestamista.', 'BLOQUEA', 'ALTA'],
];

/* ==================== VALIDADORES DE CAMPO ==================== */
// Funciones puras: reciben el valor crudo y devuelven { ok, norm, msg }.

/** V-01 — DNI: solo dígitos, 7 u 8 de longitud. */
function vDni_(v) {
  const n = normDni_(v);
  if (!n) return { ok: false, norm: '', msg: 'El DNI es obligatorio.' };
  if (n.length < 7 || n.length > 8) return { ok: false, norm: n, msg: 'Ingresá un DNI válido (7 u 8 dígitos, sin puntos).' };
  return { ok: true, norm: n, msg: '' };
}

/** V-04 — Correo: formato válido CON dominio de nivel superior (rechaza "@gmail"). */
function vEmail_(v) {
  const n = normEmail_(v);
  if (!n) return { ok: false, norm: '', msg: 'El correo electrónico es obligatorio.' };
  // Requiere al menos un punto con TLD de 2+ letras después de la @.
  if (!/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(n)) {
    return { ok: false, norm: n, msg: 'El correo no es válido. Verificá que termine en .com, .ar, etc.' };
  }
  return { ok: true, norm: n, msg: '' };
}

/** V-06 — Teléfono: normaliza a E.164 argentino (+54 + 10 dígitos). */
function normPhoneE164_(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (!d) return '';
  if (d.length > 10 && d.indexOf('54') === 0) d = d.slice(2); // quita código de país
  if (d.length === 11 && d.charAt(0) === '9') d = d.slice(1);  // quita prefijo móvil "9"
  if (d.length !== 10) return '';                              // debe quedar área+número = 10
  return '+54' + d;
}
function vPhone_(v) {
  const e164 = normPhoneE164_(v);
  if (!e164) return { ok: false, norm: '', msg: 'Ingresá un teléfono válido de 10 dígitos.' };
  return { ok: true, norm: e164, msg: '' };
}

/** Dígito verificador de CUIL/CUIT desde los 10 primeros dígitos (módulo 11). Devuelve 0-9.
 *  (Distinto de cuilCheckDigit_(prefix, dni) de LoanManagerV2.gs — no lo pisa.) */
function cuilDv_(first10) {
  const mult = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
  let s = 0;
  for (let i = 0; i < 10; i++) s += parseInt(first10.charAt(i), 10) * mult[i];
  let r = 11 - (s % 11);
  if (r === 11) r = 0;
  if (r === 10) r = 9;
  return r;
}
/** V-03 — CUIL: 11 dígitos, prefijo válido, DV módulo 11 y 8 centrales = DNI. */
function vCuil_(v, dni) {
  const c = String(v == null ? '' : v).replace(/\D/g, '');
  if (!c) return { ok: false, norm: '', msg: 'El CUIL es obligatorio.' };
  if (c.length !== 11) return { ok: false, norm: c, msg: 'El CUIL debe tener 11 dígitos.' };
  if (CUIL_PREFIXES.indexOf(c.slice(0, 2)) < 0) return { ok: false, norm: c, msg: 'El CUIL no es válido o no corresponde al DNI ingresado.' };
  if (cuilDv_(c.slice(0, 10)) !== parseInt(c.charAt(10), 10)) return { ok: false, norm: c, msg: 'El CUIL no es válido o no corresponde al DNI ingresado.' };
  const dn = normDni_(dni);
  if (dn && c.slice(2, 10) !== dn.padStart(8, '0')) {
    return { ok: false, norm: c, msg: 'El CUIL no es válido o no corresponde al DNI ingresado.' };
  }
  return { ok: true, norm: c, msg: '' };
}

/* ============ DIRECTORIO "Clientes" (esquema extendido) ============ */
// Resolución de columnas por nombre de encabezado — robusta al reordenamiento.
let _clientesHdrCache = null;
function clientesHeaderMap_() {
  if (_clientesHdrCache) return _clientesHdrCache;
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS), map = {};
  if (sh && sh.getLastColumn() >= 1) {
    sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach((h, i) => {
      const k = String(h).trim(); if (k) map[k] = i + 1;
    });
  }
  _clientesHdrCache = map; return map;
}
function invalidateClientesHdrCache_() { _clientesHdrCache = null; }
function ccol_(name, fallback) { return clientesHeaderMap_()[name] || fallback || 0; }

/** Normaliza un nombre para comparación (minúsculas, sin acentos, espacios colapsados). */
function normName_(v) {
  return String(v == null ? '' : v).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * Lee todas las filas de Clientes como objetos {id,name,email,dni,phone,row}.
 * Memoizado por ejecución (_clientesRowsCache): en una solicitud se llama 5+
 * veces (clienteByDni_, dniBelongsToOtherName_, emailDupOtherDni_,
 * phoneDupOtherDni_, clientExposure_). El caché se invalida en cada escritura a
 * "Clientes" (invalidateClientesRows_), por lo que el resultado es idéntico.
 */
let _clientesRowsCache = null;
function invalidateClientesRows_() { _clientesRowsCache = null; }
function clientesRows_() {
  if (_clientesRowsCache) return _clientesRowsCache;
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS);
  if (!sh || sh.getLastRow() < 2) { _clientesRowsCache = []; return _clientesRowsCache; }
  const cId = ccol_('ID Cliente', 1), cName = ccol_('Nombre', 2), cEmail = ccol_('Correo', 3),
    cDni = ccol_('DNI', 4), cPhone = ccol_('Teléfono', 5), cDniN = ccol_('DNI (norm)', 6),
    cEmailN = ccol_('Correo (norm)', 7), cPhoneN = ccol_('Teléfono (norm)', 8), cCuil = ccol_('CUIL', 12),
    cAddr = ccol_('Dirección', 13),
    cR1N = ccol_('Ref 1 Nombre', 0), cR1R = ccol_('Ref 1 Vínculo', 0), cR1P = ccol_('Ref 1 Teléfono', 0),
    cR2N = ccol_('Ref 2 Nombre', 0), cR2R = ccol_('Ref 2 Vínculo', 0), cR2P = ccol_('Ref 2 Teléfono', 0),
    cBloq = ccol_('Bloqueado', 0), cBloqM = ccol_('Motivo de bloqueo', 0);
  const width = Math.max(cId, cName, cEmail, cDni, cPhone, cDniN, cEmailN, cPhoneN, cCuil, cAddr,
    cR1N, cR1R, cR1P, cR2N, cR2R, cR2P, cBloq, cBloqM);
  const cell = (r, c) => c ? String(r[c - 1] || '').trim() : '';
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  _clientesRowsCache = data.map((r, i) => ({
    id: String(r[cId - 1] || '').trim(),
    name: r[cName - 1], email: r[cEmail - 1], dni: r[cDni - 1], phone: r[cPhone - 1],
    cuil: cCuil ? String(r[cCuil - 1] || '').trim() : '',
    address: cAddr ? String(r[cAddr - 1] || '').trim() : '',
    ref1Name: cell(r, cR1N), ref1Rel: cell(r, cR1R), ref1Phone: cell(r, cR1P),
    ref2Name: cell(r, cR2N), ref2Rel: cell(r, cR2R), ref2Phone: cell(r, cR2P),
    dniNorm: normDni_(r[cDniN - 1] || r[cDni - 1]),
    emailNorm: normEmail_(r[cEmailN - 1] || r[cEmail - 1]),
    phoneNorm: String(r[cPhoneN - 1] || '').trim() || normPhoneE164_(r[cPhone - 1]),
    bloqueado: (function () { const v = cell(r, cBloq).toUpperCase(); return v === 'SÍ' || v === 'SI'; })(),
    bloqueoMotivo: cell(r, cBloqM),
    row: i + 2,
  })).filter(o => o.id || o.dniNorm || o.emailNorm);
  return _clientesRowsCache;
}

/** Busca un cliente por DNI normalizado. */
function clienteByDni_(dniNorm) {
  if (!dniNorm) return null;
  const hit = clientesRows_().filter(o => o.dniNorm === dniNorm);
  return hit.length ? hit[0] : null;
}
/** V-02 — ¿el DNI ya pertenece a OTRO nombre? */
function dniBelongsToOtherName_(dniNorm, name) {
  const c = clienteByDni_(dniNorm);
  if (!c) return false;
  const a = normName_(name);
  return !!a && !!normName_(c.name) && normName_(c.name) !== a;
}
/** V-05 — clientes que comparten el correo con OTRO DNI. */
function emailDupOtherDni_(emailNorm, dniNorm) {
  if (!emailNorm) return [];
  return clientesRows_().filter(o => o.emailNorm === emailNorm && o.dniNorm && o.dniNorm !== dniNorm);
}
/** V-07 — clientes que comparten el teléfono con OTRO DNI. */
function phoneDupOtherDni_(phoneNorm, dniNorm) {
  if (!phoneNorm) return [];
  return clientesRows_().filter(o => o.phoneNorm === phoneNorm && o.dniNorm && o.dniNorm !== dniNorm);
}

/* ==================== BLOQUEO DE CLIENTES (V-33) ==================== */

// V-33 — mensaje único hacia el solicitante (no revela el motivo del bloqueo).
const BLOQUEO_MSG_CLIENTE = 'No es posible procesar solicitudes para este cliente. Ante cualquier duda, comunicate con el prestamista.';

/**
 * V-33 — devuelve el PRIMER cliente BLOQUEADO ("Bloqueado" = SÍ en Clientes) que
 * coincida por ID Cliente, correo, DNI, CUIL o teléfono. Los identificadores llegan
 * crudos y se normalizan acá; los vacíos nunca coinciden. Solo lectura (usa el caché
 * de clientesRows_; quienes escriben el bloqueo lo invalidan). null = no bloqueado.
 */
function findClienteBloqueado_(ident) {
  ident = ident || {};
  const id = String(ident.clientId || '').trim(),
    em = normEmail_(ident.email || ''), dn = normDni_(ident.dni || ''),
    cu = String(ident.cuil || '').replace(/\D/g, ''), ph = normPhoneE164_(ident.phone || '');
  if (!id && !em && !dn && !cu && !ph) return null;
  return clientesRows_().filter(o => o.bloqueado && (
    (id && o.id === id) || (em && o.emailNorm === em) || (dn && o.dniNorm === dn) ||
    (cu && String(o.cuil || '').replace(/\D/g, '') === cu) || (ph && o.phoneNorm === ph)
  ))[0] || null;
}

/**
 * Bloquea/desbloquea el cliente de la FILA ACTIVA de "Clientes" (menú 🚫). Al bloquear
 * pide el motivo y escribe "SÍ" + motivo con fecha; al desbloquear vacía ambas celdas.
 */
function setClienteBloqueoFilaActiva_(bloquear) {
  const ss = getSS_(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== CFG.SHEETS.CLIENTS) { ui.alert('Abrí la hoja "Clientes", seleccioná la fila del cliente y volvé a usar esta opción.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Seleccioná la fila de un cliente.'); return; }
  const cBloq = ccol_('Bloqueado', 0), cBloqM = ccol_('Motivo de bloqueo', 0);
  if (!cBloq) { ui.alert('Falta la columna "Bloqueado". Ejecutá configurarValidaciones (o el menú de configuración) primero.'); return; }
  const cId = ccol_('ID Cliente', 1), cName = ccol_('Nombre', 2);
  const id = String(sh.getRange(row, cId).getValue()).trim();
  const name = String(sh.getRange(row, cName).getValue()).trim();
  if (!id && !name) { ui.alert('La fila seleccionada no tiene un cliente.'); return; }
  const label = (name || id) + (name && id ? ' (' + id + ')' : '');
  if (bloquear) {
    const resp = ui.prompt('Bloquear cliente', 'Motivo del bloqueo para ' + label + ':', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const motivo = String(resp.getResponseText() || '').trim();
    sh.getRange(row, cBloq).setValue('SÍ');
    if (cBloqM) sh.getRange(row, cBloqM).setValue((motivo ? motivo + ' — ' : '') + 'bloqueado el ' + fmtDate_(new Date()));
  } else {
    sh.getRange(row, cBloq).clearContent();
    if (cBloqM) sh.getRange(row, cBloqM).clearContent();
  }
  invalidateClientesRows_();
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
  try { ss.toast(label + (bloquear ? ' BLOQUEADO: no podrá solicitar préstamos (correo/DNI/CUIL/teléfono).' : ' desbloqueado.'), 'Bloqueo de clientes', 5); } catch (e) { }
}
/** Menú: bloquea el cliente de la fila activa de "Clientes" (pide el motivo). */
function bloquearClienteFilaActiva() { setClienteBloqueoFilaActiva_(true); }
/** Menú: desbloquea el cliente de la fila activa de "Clientes". */
function desbloquearClienteFilaActiva() { setClienteBloqueoFilaActiva_(false); }

/** Crea un cliente en el esquema extendido de 11 columnas. Devuelve ID Cliente. */
function appendClienteV2_(name, email, dni, phone, opts) {
  opts = opts || {};
  const sh = ensureClientesSchema_();
  const emV = vEmail_(email);
  // Una sola lectura de la columna A para el próximo ID y la primera fila vacía
  // (antes: nextClienteId_ + firstEmptyClienteRow_ leían la columna por separado).
  const lastRow = sh.getLastRow();
  const colA = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  let maxId = 0, emptyRow = 0;
  for (let i = 0; i < colA.length; i++) {
    const v = String(colA[i][0]).trim();
    if (v === '') { if (!emptyRow) emptyRow = i + 2; continue; }
    const mm = /^C-(\d+)$/.exec(v); if (mm) maxId = Math.max(maxId, parseInt(mm[1], 10));
  }
  const id = 'C-' + String(maxId + 1).padStart(4, '0');
  const row = emptyRow || (lastRow < 2 ? 2 : lastRow + 1);
  const vals = new Array(CLIENTES_HEADERS_V2.length).fill('');
  vals[ccol_('ID Cliente', 1) - 1] = id;
  vals[ccol_('Nombre', 2) - 1] = name;
  vals[ccol_('Correo', 3) - 1] = email;
  vals[ccol_('DNI', 4) - 1] = dni;
  vals[ccol_('Teléfono', 5) - 1] = phone;
  vals[ccol_('DNI (norm)', 6) - 1] = normDni_(dni);
  vals[ccol_('Correo (norm)', 7) - 1] = emV.norm;
  vals[ccol_('Teléfono (norm)', 8) - 1] = normPhoneE164_(phone);
  vals[ccol_('Correo válido', 9) - 1] = emV.ok ? 'SÍ' : 'NO';
  if (ccol_('Duplicado', 10)) vals[ccol_('Duplicado', 10) - 1] = opts.duplicado || '';
  if (ccol_('Nota', 11)) vals[ccol_('Nota', 11) - 1] = opts.nota || '';
  if (ccol_('CUIL', 12)) vals[ccol_('CUIL', 12) - 1] = opts.cuil || '';
  if (ccol_('Dirección', 13)) vals[ccol_('Dirección', 13) - 1] = opts.address || '';
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  invalidateClientesHdrCache_();
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
  return id;
}
/** Reutiliza el cliente si coincide el DNI; si no, lo crea. Devuelve ID Cliente. */
function getOrCreateClienteV2_(name, email, dni, phone, opts) {
  const existing = clienteByDni_(normDni_(dni));
  return existing ? existing.id : appendClienteV2_(name, email, dni, phone, opts);
}
/** Actualiza el correo de un cliente (y su normalizado/validez) en la fila dada. */
function updateClienteEmail_(row, email) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS); if (!sh || !row) return;
  const emV = vEmail_(email);
  const cEmail = ccol_('Correo', 3), cEmailN = ccol_('Correo (norm)', 7), cValid = ccol_('Correo válido', 9);
  if (cEmail) sh.getRange(row, cEmail).setValue(email);
  if (cEmailN) sh.getRange(row, cEmailN).setValue(emV.norm);
  if (cValid) sh.getRange(row, cValid).setValue(emV.ok ? 'SÍ' : 'NO');
  invalidateClientesHdrCache_();
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
}
/** Guarda / actualiza el CUIL de un cliente en la fila dada. */
function updateClienteCuil_(row, cuil) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS), col = ccol_('CUIL', 12);
  if (!sh || !col || !row || !cuil) return;
  sh.getRange(row, col).setValue(cuil);
  invalidateClientesHdrCache_();
  invalidateClientesRows_();
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
}
/** Guarda / actualiza la Dirección de un cliente en la fila dada. */
function updateClienteDireccion_(row, address) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS), col = ccol_('Dirección', 13);
  if (!sh || !col || !row || !address) return;
  sh.getRange(row, col).setValue(address);
  invalidateClientesHdrCache_();
  invalidateClientesRows_();
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
}
/** Marca la columna "Duplicado" de un cliente (V-05/V-07). */
function markClienteDuplicado_(row, valor, nota) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS), col = ccol_('Duplicado', 10);
  if (!sh || !col || !row) return;
  sh.getRange(row, col).setValue(valor || 'REVISAR');
  const cNota = ccol_('Nota', 11);
  if (cNota && nota) sh.getRange(row, cNota).setValue(nota);
  invalidateClientesRows_();
}

/* ================= HELPERS DE PLAZO (en días) ================= */
function termRateDays_(days) { const r = TERM_DAYS_RATES[Number(days)]; return r == null ? null : r; }
function termLabelDays_(days) { return Number(days) + ' días'; }
function termRatePctDays_(days) { const r = termRateDays_(days); return r == null ? '' : round2_(r * 100) + '%'; }

/* ===== Tramos por MONTO (fuente única de verdad para form, alta y aprobación) ===== */
// Lee un ajuste de dinero de Configuración; usa el valor por defecto si falta/ inválido.
function settingMoney_(key, def) {
  const raw = String(getSetting_(key) || '').replace(/[^0-9.\-]/g, '').trim();
  const n = raw === '' ? NaN : Number(raw);
  return (isFinite(n) && n > 0) ? n : def;
}
function loanMax_() { return settingMoney_('Monto máximo por préstamo', 500000); }
function tier1Max_() { return settingMoney_('Tope tramo 25% (15 días)', 150000); }
function tier2Max_() { return settingMoney_('Tope tramo 50% (30 días)', 300000); }
// Info fija por plazo en días: 15→25%/1 cuota, 30→50%/1 cuota, 90→100%/3 cuotas mensuales.
function termInfoForDays_(days) {
  days = Number(days) || 0;
  if (days === 15) return { days: 15, rate: 0.25, cuotas: 1, cadence: 'once' };
  if (days === 30) return { days: 30, rate: 0.5, cuotas: 1, cadence: 'once' };
  return { days: 90, rate: 1.0, cuotas: 3, cadence: 'monthly' };
}
/**
 * Plazos PERMITIDOS según el monto:
 *   ≤ tramo1 → [15, 30]  (el cliente elige 15 o 30 días)
 *   ≤ tramo2 → [30]      (forzado)
 *   > tramo2 (hasta el tope) → [90]  (forzado, 3 cuotas)
 */
function allowedTermsForAmount_(amount) {
  const a = Number(amount) || 0;
  if (a <= 0) return [];
  if (a <= tier1Max_()) return [15, 30];
  if (a <= tier2Max_()) return [30];
  return [90];
}
/** Plazo por defecto del monto = primer permitido. */
function termForAmount_(amount) { const al = allowedTermsForAmount_(amount); return termInfoForDays_(al.length ? al[0] : 15); }
/** Plazo efectivo: respeta la elección del cliente si es válida para el monto; si no, usa el default. */
function resolveTerm_(amount, chosenDays) {
  const al = allowedTermsForAmount_(amount);
  if (!al.length) return termForAmount_(amount);
  const d = al.indexOf(Number(chosenDays)) >= 0 ? Number(chosenDays) : al[0];
  return termInfoForDays_(d);
}
function dueDateForDays_(days, from) { const d = new Date(from.getTime()); d.setDate(d.getDate() + Number(days || 0)); return d; }

/* ============ API PÚBLICA DEL FORMULARIO INTELIGENTE ============ */

/**
 * Detección de cliente para el formulario. Se llama al ingresar Correo + DNI:
 *   • found  → el DNI y el correo pertenecen al MISMO cliente: se precargan sus
 *              datos (todos menos las fotos).
 *   • conflict → el DNI existe con otro correo, o el correo existe con otro DNI
 *              (se muestra un error; no se avanza).
 *   • ni found ni conflict → cliente nuevo: el formulario sigue EN BLANCO.
 */
function lookupClienteForIntake(email, dni) {
  return guard_('lookupClienteForIntake', function () {
    const dnV = vDni_(dni), emV = vEmail_(email);
    const res = {
      ok: true,
      dniValid: dnV.ok, dniMsg: dnV.msg,
      emailValid: emV.ok, emailMsg: emV.msg,
      found: false, client: null, conflict: false, conflictField: '', hasRefs: false,
    };
    if (!dnV.ok || !emV.ok) return res;   // se necesitan correo y DNI válidos
    // V-33 — cliente bloqueado: se corta en el paso 1 (mensaje genérico, sin detalles).
    const blk = findClienteBloqueado_({ email: emV.norm, dni: dnV.norm });
    if (blk) { res.blocked = true; res.blockedMsg = BLOQUEO_MSG_CLIENTE; return res; }
    // V-32 — límite por historial de repago (escalera de graduación), específico del
    // prestatario. Aplica también a clientes nuevos (aún sin historial → límite inicial).
    const h = repaymentHistory_(dnV.norm);
    res.maxAmount = graduationMax_(dnV.norm);
    res.maxAmountFmt = fmtMoney_(res.maxAmount);
    res.settledCount = h.settledCount;
    res.onTimeCount = h.onTimeCount;
    // Límite reducido HOY: mora vigente, o atrasos sin repagos a tiempo posteriores
    // (un atraso reinicia la escalera; se recupera saldando a tiempo).
    res.hasArrears = !!(h.everDefaulted || (h.lateCount > 0 && h.onTimeCount === 0));
    const cByDni = clienteByDni_(dnV.norm);
    if (cByDni && normEmail_(cByDni.email) === emV.norm) {
      res.found = true;                   // mismo cliente (DNI + correo)
      res.client = { clientId: cByDni.id, name: cByDni.name, phone: cByDni.phone, email: cByDni.email, cuil: cByDni.cuil || '', address: cByDni.address || '',
        ref1Name: cByDni.ref1Name || '', ref1Rel: cByDni.ref1Rel || '', ref1Phone: cByDni.ref1Phone || '',
        ref2Name: cByDni.ref2Name || '', ref2Rel: cByDni.ref2Rel || '', ref2Phone: cByDni.ref2Phone || '' };
      // ¿Ya tiene DOS referencias completas registradas? → se reutilizan (no re-pedir).
      res.hasRefs = !!(cByDni.ref1Name && cByDni.ref1Phone && cByDni.ref2Name && cByDni.ref2Phone);
      return res;
    }
    if (cByDni) { res.conflict = true; res.conflictField = 'dni'; return res; }     // DNI usado con otro correo
    const emailOwner = clientesRows_().filter(o => o.emailNorm === emV.norm)[0];
    if (emailOwner) { res.conflict = true; res.conflictField = 'email'; return res; } // correo usado con otro DNI
    return res;                           // cliente nuevo
  });
}

/** Etiqueta legible del monto (agrupación de miles argentina) para respuestas. */
function fmtAmount_(n) { return fmtMoney_(n); }

/**
 * Recepción del formulario inteligente. Valida (V-01…V-08, V-13) del lado del
 * servidor —nunca confía en el cliente—, detecta el cliente existente por DNI
 * y registra la solicitud en "Nuevos Prestatarios".
 */
function submitIntakeSmart(form) {
  return guard_('submitIntakeSmart', function () {
    if (!acceptingApplications_()) throw new Error('En este momento no estamos aceptando nuevas solicitudes de préstamo.');

    const mode = String(form.mode || 'new').trim();
    const emV = vEmail_(form.email);
    if (!emV.ok) throw new Error(emV.msg);                 // V-04
    const dnV = vDni_(form.dni);
    if (!dnV.ok) throw new Error(dnV.msg);                 // V-01

    // V-33 — cliente bloqueado: rechazo ABSOLUTO por CUALQUIERA de los cuatro
    // identificadores del formulario (correo, DNI, CUIL o teléfono).
    if (findClienteBloqueado_({ email: emV.norm, dni: dnV.norm, cuil: form.cuil, phone: form.phone }))
      throw new Error(BLOQUEO_MSG_CLIENTE);

    const existing = clienteByDni_(dnV.norm);
    const emailOwner = clientesRows_().filter(o => o.emailNorm === emV.norm)[0] || null;
    // Cliente existente = el DNI y el correo pertenecen al MISMO registro → se precarga.
    const isReturning = !!existing && normEmail_(existing.email) === emV.norm;
    // Conflictos (respaldo del bloqueo del formulario): DNI o correo ya usados por otro.
    if (!isReturning) {
      if (existing) throw new Error('Ese DNI ya está registrado con otro correo. Si ya sos cliente, ingresá el correo registrado.');
      if (emailOwner) throw new Error('Ese correo ya está registrado con otro DNI.');
    }

    // Identidad: cliente existente → se toma de "Clientes"; nuevo → del formulario.
    let name, phone, cuil = '';
    if (isReturning) {
      name = String(existing.name || '').trim();
      phone = String(existing.phone || '').trim();
      // Si el correo ingresado es válido y difiere del registrado, actualizarlo.
      if (normEmail_(existing.email) !== emV.norm) updateClienteEmail_(existing.row, emV.norm);
    } else {
      name = String(form.fullName || '').trim();
      phone = String(form.phone || '').trim();
      if (!name) throw new Error('El nombre completo es obligatorio.');
      // V-02 — el DNI no puede pertenecer a otra persona.
      if (dniBelongsToOtherName_(dnV.norm, name)) throw new Error('Ese DNI ya está registrado a nombre de otra persona.');
      const phV = vPhone_(phone);
      if (!phV.ok) throw new Error(phV.msg);               // V-06
      phone = phV.norm;
    }
    // CUIL: obligatorio para TODOS (nuevos y existentes). No se almacena, se pide siempre.
    const cuV = vCuil_(form.cuil, dnV.norm);
    if (!cuV.ok) throw new Error(cuV.msg);                 // V-03
    cuil = cuV.norm;

    // Datos del préstamo (obligatorios para todos).
    const amount = Number(String(form.amount || '').replace(/\D/g, '')) || 0;
    // V-30: el plazo/tasa/cuotas se DERIVAN del monto. Para montos chicos (≤ tramo1) el
    // cliente puede elegir 15 o 30 días; resolveTerm_ respeta la elección si es válida.
    const tinfo = resolveTerm_(amount, Number(form.term) || 0);
    const term = tinfo.days;
    const notes = String(form.notes || '').trim();
    const address = String(form.address || '').trim();
    // Forma de pago elegida por el prestatario (predeterminada: Mercado Pago) + datos opcionales.
    const payMethodSel = String(form.payMethod || '').trim() || 'Mercado Pago';
    const payDetails = String(form.payDetails || '').trim();
    const payMethodFull = payDetails ? (payMethodSel + ' — ' + payDetails) : payMethodSel;
    if (!address) throw new Error('La dirección es obligatoria.');   // obligatoria para nuevos y existentes
    if (!amount) throw new Error('Ingresá un monto de préstamo válido.');
    if (amount > loanMax_()) throw new Error('El monto máximo por préstamo es ' + fmtMoney_(loanMax_()) + '.'); // V-29
    if (termRateDays_(term) == null) throw new Error('Plazo inválido (derivado del monto).'); // V-30
    if (!notes) throw new Error('Las notas / motivo son obligatorias.');

    // V-27 — dos referencias. Cliente que vuelve con referencias completas registradas → se
    // REUTILIZAN (no se re-piden). Cliente nuevo (o existente sin referencias) → obligatorias.
    const reuseRefs = isReturning && String(form.reuseRefs || '') === '1' && existing &&
      String(existing.ref1Name || '').trim() && String(existing.ref1Phone || '').trim() &&
      String(existing.ref2Name || '').trim() && String(existing.ref2Phone || '').trim();
    let ref1Name, ref1Rel, ref1PhoneN, ref2Name, ref2Rel, ref2PhoneN;
    if (reuseRefs) {
      ref1Name = String(existing.ref1Name).trim(); ref1Rel = String(existing.ref1Rel || '').trim(); ref1PhoneN = String(existing.ref1Phone).trim();
      ref2Name = String(existing.ref2Name).trim(); ref2Rel = String(existing.ref2Rel || '').trim(); ref2PhoneN = String(existing.ref2Phone).trim();
    } else {
      ref1Name = String(form.ref1Name || '').trim(); ref1Rel = String(form.ref1Rel || '').trim();
      ref2Name = String(form.ref2Name || '').trim(); ref2Rel = String(form.ref2Rel || '').trim();
      const ref1Ph = vPhone_(form.ref1Phone), ref2Ph = vPhone_(form.ref2Phone);
      if (!ref1Name || !ref1Rel || !ref1Ph.ok) throw new Error('Completá la Referencia 1: nombre, vínculo y un teléfono válido de 10 dígitos.');
      if (!ref2Name || !ref2Rel || !ref2Ph.ok) throw new Error('Completá la Referencia 2: nombre, vínculo y un teléfono válido de 10 dígitos.');
      ref1PhoneN = ref1Ph.norm; ref2PhoneN = ref2Ph.norm;
    }

    // V-26 — compromiso de pago: acuse explícito obligatorio (devolución, mora, consecuencias legales
    // y comunicación ante atrasos como condición para préstamos futuros).
    if (!form.agreeRepay || !form.agreeMora || !form.agreeConseq || !form.agreeContact)
      throw new Error('Debe aceptar todos los compromisos de pago para enviar la solicitud.');
    if (!form.agree) throw new Error('Debe aceptar los Términos y Condiciones para enviar la solicitud.');

    // V-08 — fotos del DNI: obligatorias para TODOS (nuevos y existentes).
    if (!hasFile_(form.dniPhoto) || !hasFile_(form.cuilPhoto)) throw new Error('Subí ambas fotos del DNI, nítidas y completas.');

    // Advertencias no bloqueantes de duplicado (V-05 / V-07).
    const warnings = [];
    const emailDup = emailDupOtherDni_(emV.norm, dnV.norm);
    if (emailDup.length) warnings.push('Este correo ya figura para otro solicitante (' + emailDup.map(x => x.id).join(', ') + ').');
    const phoneDup = phone ? phoneDupOtherDni_(phone, dnV.norm) : [];
    if (phoneDup.length) warnings.push('Este teléfono ya figura para otro solicitante (' + phoneDup.map(x => x.id).join(', ') + ').');
    // V-28 (ADVIERTE) — teléfono de referencia igual al del solicitante o repetido entre referencias (posible referencia falsa).
    if (!reuseRefs && ((phone && (ref1PhoneN === phone || ref2PhoneN === phone)) || (ref1PhoneN && ref1PhoneN === ref2PhoneN)))
      warnings.push('⚠ Teléfono de referencia repetido o igual al del solicitante — verificar las referencias.');

    const rate = termRateDays_(term), interest = round2_(amount * rate), total = round2_(amount * (1 + rate));

    // Alta / reutilización del cliente (esquema normalizado).
    const dupFlag = (emailDup.length || phoneDup.length) ? 'REVISAR' : '';
    const clientId = isReturning ? existing.id
      : getOrCreateClienteV2_(name, emV.norm, dnV.norm, phone, {
        duplicado: dupFlag,
        nota: warnings.length ? 'Contacto compartido detectado en el alta.' : '',
        cuil: cuil,
        address: address,
      });
    // Cliente existente: guardar/actualizar CUIL y Dirección si cambiaron o faltaban.
    if (isReturning && existing && existing.row) {
      if (String(existing.cuil || '').replace(/\D/g, '') !== cuil) updateClienteCuil_(existing.row, cuil);
      if (String(existing.address || '').trim() !== address) updateClienteDireccion_(existing.row, address);
    }
    // Si es duplicado y el cliente ya existía, marcarlo para revisión manual.
    if (dupFlag && existing && existing.row) markClienteDuplicado_(existing.row, 'REVISAR', 'Contacto compartido detectado en solicitud.');

    // Fotos → carpeta del prestatario.
    const folder = borrowerFolder_(borrowerFolderName_(name, dnV.norm));
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const dniUrl = hasFile_(form.dniPhoto) ? savePhoto_(folder, form.dniPhoto, name + ' DNI ' + stamp) : '';
    const cuilUrl = hasFile_(form.cuilPhoto) ? savePhoto_(folder, form.cuilPhoto, name + ' DNI dorso ' + stamp) : '';

    // Registro en "Nuevos Prestatarios" (mismo orden de columnas que NB).
    const ss = getSS_();
    const nb = ss.getSheetByName(CFG.SHEETS.NEW) || setupNew_(ss) || ss.getSheetByName(CFG.SHEETS.NEW);
    // Cols: 1 Fecha,2 Nombre,3 Correo,4 DNI,5 Teléfono,6 Monto,7 Plazo,8 Notas,
    //       9 dniFrente,10 dniDorso,11 Verificado?,12 Rechazar?,13 Resultado,
    //       14 Verificar BCRA?,15 Parámetro BCRA,16 CUIL.
    const rowVals = [new Date(), name, emV.norm, dnV.norm, phone, amount, term, notes, dniUrl, cuilUrl,
      false, false, 'Pendiente de verificación', false, 'DNI', cuil];
    const row = nbAppend_(nb, rowVals);
    // Forma de pago: en NB por nombre de encabezado, y en el cliente (para el contrato).
    try { setSheetFieldByHeader_(nb, row, 'Forma de Pago', payMethodFull); } catch (e) { logError_('submitIntakeSmart:payMethodNb', e); }
    try { setClienteFieldByHeader_(clientId, CLIENTE_PAY_HEADER, payMethodFull); } catch (e) { logError_('submitIntakeSmart:payMethod', e); }
    nb.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    nb.getRange(row, 6).setNumberFormat(CFG.CURRENCY_FMT);
    // Dirección: columna adicional (se resuelve por nombre; se crea si falta en hojas anteriores).
    try {
      const nbMap = headerMap_(nb);
      let dirCol = nbMap['Dirección'];
      if (!dirCol) {
        dirCol = nb.getLastColumn() + 1;
        nb.getRange(1, dirCol).setValue('Dirección').setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff');
      }
      nb.getRange(row, dirCol).setValue(address);
    } catch (e) { logError_('submitIntakeSmart:direccion', e); }

    // Referencias + compromiso: en "Nuevos Prestatarios" (para revisar) y en "Clientes"
    // (persisten tras aprobar, cuando se borra la fila de la solicitud → útiles para cobranza).
    try {
      setSheetFieldByHeader_(nb, row, 'Ref 1 Nombre', ref1Name);
      setSheetFieldByHeader_(nb, row, 'Ref 1 Vínculo', ref1Rel);
      setSheetFieldByHeader_(nb, row, 'Ref 1 Teléfono', ref1PhoneN);
      setSheetFieldByHeader_(nb, row, 'Ref 2 Nombre', ref2Name);
      setSheetFieldByHeader_(nb, row, 'Ref 2 Vínculo', ref2Rel);
      setSheetFieldByHeader_(nb, row, 'Ref 2 Teléfono', ref2PhoneN);
      setSheetFieldByHeader_(nb, row, 'Compromiso aceptado', new Date());
    } catch (e) { logError_('submitIntakeSmart:referenciasNb', e); }
    try {
      setClienteFieldByHeader_(clientId, 'Ref 1 Nombre', ref1Name);
      setClienteFieldByHeader_(clientId, 'Ref 1 Vínculo', ref1Rel);
      setClienteFieldByHeader_(clientId, 'Ref 1 Teléfono', ref1PhoneN);
      setClienteFieldByHeader_(clientId, 'Ref 2 Nombre', ref2Name);
      setClienteFieldByHeader_(clientId, 'Ref 2 Vínculo', ref2Rel);
      setClienteFieldByHeader_(clientId, 'Ref 2 Teléfono', ref2PhoneN);
    } catch (e) { logError_('submitIntakeSmart:referenciasCliente', e); }

    // Correo de confirmación al solicitante.
    try {
      sendBrandedEmail_(emV.norm, 'Solicitud recibida',
        'Estimado/a ' + name + ', recibimos su solicitud de préstamo por ' + fmtMoney_(amount) + '. La revisaremos y nos pondremos en contacto.',
        '<p>Estimado/a ' + esc_(name) + ',</p><p>Recibimos su solicitud de préstamo por <b>' + fmtMoney_(amount) +
        '</b> a ' + termLabelDays_(term) + '. La revisaremos y nos pondremos en contacto a la brevedad.</p>');
    } catch (e) { logError_('submitIntakeSmart:email', e); }

    // Aviso al prestamista con los datos de la solicitud.
    try {
      const recipients = noticeRecipients_();
      if (recipients.length) {
        const avail = fundStats_().net; // neto real: puede ser negativo (sobregiro)
        const warnHtml = warnings.length ? '<p style="color:#a50e0e"><b>⚠ Revisar:</b> ' + esc_(warnings.join(' ')) + '</p>' : '';
        sendBrandedEmail_(recipients.join(','), 'Nueva solicitud de préstamo — ' + name,
          'Nueva solicitud de ' + name + ' por ' + fmtMoney_(amount) + ' a ' + termLabelDays_(term) + '.',
          '<p>Se recibió una <b>nueva solicitud de préstamo</b>' + (isReturning ? ' de un <b>cliente existente</b> (' + esc_(clientId) + ')' : ' de un <b>cliente nuevo</b>') + '.</p>' + warnHtml +
          '<table style="border-collapse:collapse">' +
          row_('ID Cliente', esc_(clientId)) + row_('Nombre completo', esc_(name)) + row_('Correo', esc_(emV.norm)) +
          row_('DNI', esc_(dnV.norm)) + row_('Teléfono', esc_(phone)) + (cuil ? row_('CUIL', esc_(cuil)) : '') +
          row_('Monto solicitado', '<b>' + fmtMoney_(amount) + '</b>') +
          row_('Plazo', termLabelDays_(term) + ' (' + termRatePctDays_(term) + ')') +
          row_('Interés', fmtMoney_(interest)) + row_('Total a devolver', fmtMoney_(total)) +
          row_('Notas / Motivo', esc_(notes)) +
          row_('Forma de pago preferida', esc_(payMethodFull)) +
          row_('Referencia 1 (familiar)', esc_(ref1Name + ' · ' + ref1Rel + ' · ' + ref1PhoneN) + (reuseRefs ? ' <i>(en archivo)</i>' : '')) +
          row_('Referencia 2 (no familiar)', esc_(ref2Name + ' · ' + ref2Rel + ' · ' + ref2PhoneN) + (reuseRefs ? ' <i>(en archivo)</i>' : '')) +
          row_('Compromisos de pago', 'Aceptados ✓') +
          row_('Foto del frente del DNI', dniUrl ? '<a href="' + dniUrl + '">Ver documento</a>' : '— (cliente ya verificado)') +
          row_('Foto del dorso del DNI', cuilUrl ? '<a href="' + cuilUrl + '">Ver documento</a>' : '— (cliente ya verificado)') +
          row_('Efectivo disponible para prestar', (avail < amount ? '<b style="color:#900">' : '<b>') + fmtMoney_(avail) + '</b>' + (avail < amount ? ' ⚠ fondo insuficiente' : '')) +
          '</table>' +
          '<p>Revísela en la hoja <b>Nuevos Prestatarios</b> y tilde <b>Verificado?</b> para aprobar o <b>Rechazar?</b> para archivar.</p>');
      }
    } catch (e) { logError_('submitIntakeSmart:lenderEmail', e); }

    return {
      ok: true, returning: isReturning, clientId: clientId,
      name: name, email: emV.norm, dni: dnV.norm, phone: phone,
      term: term, termLabel: termLabelDays_(term), ratePct: termRatePctDays_(term),
      amountFmt: fmtMoney_(amount), interestFmt: fmtMoney_(interest), totalFmt: fmtMoney_(total),
      warnings: warnings,
      message: '¡Solicitud recibida! Gracias, ' + name + '.',
    };
  });
}

/* ============ VALIDADORES DE APROBACIÓN / INTEGRIDAD ============ */
// Reutilizables desde el flujo de aprobación (onEdit "Verificado?") de
// LoanManagerV2. Devuelven { ok, msg } — NO envían correos.

/** Exposición total (Capital) de un cliente en "Prestatarios", opcionalmente
 *  agrupando contactos compartidos (V-12). */
function clientExposure_(clientId, groupSharedContacts) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs || bs.getLastRow() < 2) return 0;
  let ids = [String(clientId).trim()];
  if (groupSharedContacts) {
    const me = clientesRows_().filter(o => o.id === String(clientId).trim())[0];
    if (me) {
      clientesRows_().forEach(o => {
        if (o.id && o.id !== me.id && ((me.emailNorm && o.emailNorm === me.emailNorm) || (me.phoneNorm && o.phoneNorm === me.phoneNorm))) {
          if (ids.indexOf(o.id) < 0) ids.push(o.id);
        }
      });
    }
  }
  // Resolver ID Cliente y Capital por NOMBRE de encabezado (robusto al layout migrado/nuevo).
  const H = headerIndex_(bs);
  const cliC = colByAny_(H, ['ID Cliente']) || PB.CLIENT_ID;
  const capC = colByAny_(H, ['Capital']) || PB.PRINCIPAL;
  const clis = bs.getRange(2, cliC, bs.getLastRow() - 1, 1).getValues();
  const caps = bs.getRange(2, capC, bs.getLastRow() - 1, 1).getValues();
  let sum = 0;
  for (let i = 0; i < clis.length; i++) { if (ids.indexOf(String(clis[i][0]).trim()) >= 0) sum += Number(caps[i][0]) || 0; }
  return round2_(sum);
}

/** Cantidad de préstamos VIGENTES (con saldo) del cliente en "Prestatarios".
 *  Solo cuentan los que aún se deben (ACTIVO/VENCIDO); los PAGADO/SALDADO NO cuentan
 *  para el tope de préstamos simultáneos (V-25). */
function clientLoanCount_(clientId) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs || bs.getLastRow() < 2) return 0;
  const id = String(clientId || '').trim();
  if (!id) return 0;
  const H = headerIndex_(bs);
  const cliC = colByAny_(H, ['ID Cliente']) || PB.CLIENT_ID;
  const stC = colByAny_(H, ['Estado']) || PB.STATE;
  const data = bs.getRange(2, 1, bs.getLastRow() - 1, Math.max(cliC, stC)).getValues();
  let n = 0;
  data.forEach(r => {
    if (String(r[cliC - 1]).trim() !== id) return;
    const st = String(r[stC - 1] || '').trim().toUpperCase();
    if (st === ST.PAID || st === ST.CLEARED) return; // PAGADO/SALDADO no ocupan cupo
    n++;
  });
  return n;
}

/**
 * V-32 — Historial de repago de un prestatario, identificado por DNI (clave estable
 * que sobrevive al archivado en "Saldados"). Combina tres hojas:
 *   • "Prestatarios": préstamos del cliente (PAGADO = saldado; VENCIDO = mora vigente).
 *   • "Saldados":     préstamos ya archivados (todos saldados), enlazados por DNI.
 *   • "Pagos":        fecha del ÚLTIMO pago por préstamo → decide on-time vs. atrasado.
 * Detección de ATRASO: se marca `lateCount` si CUALQUIER préstamo (en cualquier estado)
 * tiene un pago posterior a su Fecha de Vencimiento. Como los pagos son cronológicos, el
 * último pago > vencimiento ⇔ hubo al menos un pago tardío. `everDefaulted` = tiene una
 * mora VIGENTE hoy (fuerza el límite inicial mientras dure).
 * ESCALERA CON REINICIO: un atraso NO degrada para siempre — reinicia la escalera.
 * `onTimeCount` cuenta SOLO los préstamos saldados a tiempo DESPUÉS del último atraso
 * (`lastLateTime` = fecha del pago tardío más reciente; 0 = nunca se atrasó, cuentan
 * todos). El momento de saldado de cada préstamo es su último pago en "Pagos"
 * (doMoveCleared_ conserva "Pagos" al archivar); si no hay pagos registrados (datos
 * migrados) se usa "Fecha de Saldado" de "Saldados"; sin ninguna de las dos, el
 * préstamo no suma tras un atraso (conservador). Empates: se exige estrictamente
 * posterior (los pagos suelen guardarse solo con fecha). Devuelve
 *   { settledCount, onTimeCount, lateCount, everDefaulted, lastLateTime }.
 * Solo lectura: NO escribe ni envía correos.
 */
function repaymentHistory_(dni) {
  const dniN = normDni_(dni);
  const res = { settledCount: 0, onTimeCount: 0, lateCount: 0, everDefaulted: false, lastLateTime: 0 };
  if (!dniN) return res;
  const ss = getSS_();
  const cli = clienteByDni_(dniN);
  const clientId = cli ? cli.id : '';

  // Último pago (máx. Fecha de Pago) por ID Préstamo, desde "Pagos".
  const lastPayByLoan = {};
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (pg && pg.getLastRow() >= 2) {
    const PH = headerIndex_(pg);
    const lC = colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID;
    const dC = colByAny_(PH, ['Fecha de Pago', 'Fecha']) || PP.DATE;
    const rows = pg.getRange(2, 1, pg.getLastRow() - 1, Math.max(lC, dC)).getValues();
    rows.forEach(r => {
      const id = String(r[lC - 1] || '').trim(); if (!id) return;
      const d = r[dC - 1], t = (d instanceof Date) ? d.getTime() : Date.parse(d);
      if (!isFinite(t)) return;
      if (!(id in lastPayByLoan) || t > lastPayByLoan[id]) lastPayByLoan[id] = t;
    });
  }

  const dayEnd_ = d => { const x = new Date(d); x.setHours(23, 59, 59, 999); return x.getTime(); };
  // ¿El préstamo tuvo algún pago posterior a su vencimiento? (último pago > vencimiento).
  const isLate_ = (loanId, dueDate) => {
    const pay = lastPayByLoan[String(loanId).trim()];
    const due = (dueDate instanceof Date) ? dueDate : (dueDate ? new Date(dueDate) : null);
    if (pay == null || !due || !isFinite(due.getTime())) return false;
    return pay > dayEnd_(due);
  };
  // Acumula los préstamos del cliente (deduplicados por ID: un préstamo puede figurar
  // en ambas hojas a mitad de un archivado) para evaluarlos en dos pasadas.
  const loans = [], seen = {};
  const addLoan_ = (loanId, dueDate, settled, fallbackSettle) => {
    const id = String(loanId || '').trim();
    if (!id || seen[id]) return; seen[id] = true;
    loans.push({ id: id, due: dueDate, settled: !!settled, fallbackSettle: fallbackSettle || null });
  };

  // 1) "Prestatarios": TODOS los préstamos del cliente (por ID Cliente, cualquier estado).
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (clientId && bs && bs.getLastRow() >= 2) {
    const H = headerIndex_(bs);
    const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
    const cliC = colByAny_(H, ['ID Cliente']) || PB.CLIENT_ID;
    const dueC = colByAny_(H, ['Fecha de Vencimiento', 'Vencimiento']) || PB.DUE;
    const stC = colByAny_(H, ['Estado']) || PB.STATE;
    const data = bs.getRange(2, 1, bs.getLastRow() - 1, Math.max(idC, cliC, dueC, stC)).getValues();
    data.forEach(r => {
      if (String(r[cliC - 1] || '').trim() !== clientId) return;
      const st = String(r[stC - 1] || '').trim().toUpperCase();
      if (st === ST.OVERDUE) res.everDefaulted = true;   // mora vigente
      // Detecta un pago tardío en cualquier estado (activo/vencido/pagado); cuenta on-time sólo si PAGADO.
      addLoan_(r[idC - 1], r[dueC - 1], st === ST.PAID);
    });
  }

  // 2) "Saldados": préstamos archivados (todos saldados), por DNI. "Fecha de Saldado"
  //    sirve de respaldo cuando el préstamo no tiene filas en "Pagos" (datos migrados).
  const cs = ss.getSheetByName(CFG.SHEETS.CLEARED);
  if (cs && cs.getLastRow() >= 2) {
    const H = headerIndex_(cs);
    const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || 1;
    const dniC = colByAny_(H, ['DNI']) || 3;
    const dueC = colByAny_(H, ['Fecha de Vencimiento', 'Vencimiento']) || 12;
    const setC = colByAny_(H, ['Fecha de Saldado']) || 13;
    const data = cs.getRange(2, 1, cs.getLastRow() - 1, Math.max(idC, dniC, dueC, setC)).getValues();
    data.forEach(r => { if (normDni_(r[dniC - 1]) === dniN) addLoan_(r[idC - 1], r[dueC - 1], true, r[setC - 1]); });
  }

  // Pasada 1 — atrasos y momento del ÚLTIMO atraso (el pago tardío más reciente).
  loans.forEach(L => {
    L.late = isLate_(L.id, L.due);
    if (L.late) {
      res.lateCount++;
      const t = lastPayByLoan[L.id];
      if (t != null && t > res.lastLateTime) res.lastLateTime = t;
    }
  });

  // Pasada 2 — la escalera se RECONSTRUYE: solo suman on-time los préstamos saldados
  // (sin atraso) DESPUÉS del último atraso. Sin atrasos → cuentan todos, como antes.
  const toTime_ = v => { const t = (v instanceof Date) ? v.getTime() : Date.parse(v); return isFinite(t) ? t : null; };
  loans.forEach(L => {
    if (!L.settled) return;
    res.settledCount++;
    if (L.late) return;                                   // saldado tarde: nunca suma
    if (!res.lastLateTime) { res.onTimeCount++; return; } // nunca se atrasó
    const settleTime = (lastPayByLoan[L.id] != null) ? lastPayByLoan[L.id] : toTime_(L.fallbackSettle);
    if (settleTime != null && settleTime > res.lastLateTime) res.onTimeCount++;
  });

  return res;
}

/**
 * V-32 — Techo de monto por HISTORIAL de repago (escalera de graduación). Los tramos
 * se leen de "Configuración" (ajustables sin tocar el código):
 *   • sin historial / con mora vigente  → "Límite inicial (préstamo nuevo)"  (150k)
 *   • ≥1 préstamo saldado a tiempo      → "Límite tras 1 préstamo saldado"   (300k)
 *   • ≥2 préstamos saldados a tiempo    → "Límite tras 2 préstamos saldados" (500k)
 * Un atraso REINICIA la escalera: solo cuentan los préstamos saldados a tiempo
 * DESPUÉS del último atraso (onTimeCount ya viene depurado de repaymentHistory_),
 * así el prestatario puede volver a subir. Una mora VIGENTE fuerza el límite inicial.
 */
function graduationMax_(dni) {
  const h = repaymentHistory_(dni);
  const starter = settingMoney_('Límite inicial (préstamo nuevo)', 150000);
  if (h.everDefaulted) return starter;
  if (h.onTimeCount >= 2) return settingMoney_('Límite tras 2 préstamos saldados', 500000);
  if (h.onTimeCount >= 1) return settingMoney_('Límite tras 1 préstamo saldado', 300000);
  return starter;
}

/**
 * V-09 + V-10 + V-12 + V-33 — valida un desembolso propuesto para un cliente.
 * Con `override = true` se omiten los topes de cantidad de préstamos y de capital
 * (para el botón "Anular límites" de Nuevos Prestatarios). El bloqueo de cliente
 * (V-33) es ABSOLUTO y se chequea SIEMPRE, aun con override. `ident` (opcional):
 * { email, cuil, phone } para ampliar la coincidencia del bloqueo más allá del DNI.
 */
function validateApprovalV2_(clientId, amount, override, dni, ident) {
  // V-33 — cliente bloqueado: primero y fuera del override (no se anula).
  const blk = findClienteBloqueado_(Object.assign({ clientId: clientId, dni: dni }, ident || {}));
  if (blk) return { ok: false, code: 'V-33',
    msg: 'V-33 — Cliente bloqueado (' + (blk.id || blk.dniNorm) + '). No se anula con «Anular límites».' +
      (blk.bloqueoMotivo ? ' Motivo: ' + blk.bloqueoMotivo : '') };
  const amt = Number(amount) || 0;
  const fs = fundStats_(), avail = fs.available; // avail (acotado ≥ 0) para los topes; fs.net para mostrar el sobregiro
  if (!override) {
    // Tope de 2 préstamos por prestatario.
    if (clientLoanCount_(clientId) >= MAX_LOANS_PER_CLIENT)
      return { ok: false, code: 'V-25', msg: 'El prestatario ya tiene ' + MAX_LOANS_PER_CLIENT + ' préstamos (máximo permitido).' };
    // El nuevo préstamo no puede exceder el capital disponible (= "Efectivo disponible
    // para prestar" del Panel; el capital ya prestado no se descuenta dos veces).
    if (amt > avail)
      return { ok: false, code: 'V-09', msg: 'El préstamo (' + fmtMoney_(amt) + ') supera el capital disponible para prestar (' + fmtMoney_(fs.net) + (fs.net < 0 ? ' — sobregiro' : '') + ').' };
    // V-32 — escalera de graduación: el monto ≤ límite por historial de repago del
    // prestatario. Sin historial o con mora vigente → límite inicial; un atraso reinicia
    // la escalera (sube de nuevo saldando a tiempo). Clave por DNI (sobrevive a "Saldados").
    const gmax = graduationMax_(dni);
    if (gmax > 0 && amt > gmax)
      return { ok: false, code: 'V-32', msg: 'El monto (' + fmtMoney_(amt) + ') supera el límite por historial del prestatario (' + fmtMoney_(gmax) + '). Se amplía saldando préstamos a tiempo (los atrasos reinician la escalera), o usá «Anular límites».' };
    // Tope de concentración por prestatario: % configurable del FONDO TOTAL
    // (Configuración ▸ "Tope de concentración…"). Base estable = fondo total (no el
    // efectivo disponible, que se agota al colocar y bloquearía todo préstamo). 0 = sin tope.
    const rate = concentrationRate_();
    if (rate > 0) {
      const base = fs.total > 0 ? fs.total : avail; // sin fondo configurado → cae a disponible
      const cap = round2_(base * rate);
      const exposure = clientExposure_(clientId, true) + amt; // incluye contactos compartidos (V-12)
      if (exposure > cap) return { ok: false, code: 'V-10', msg: 'Este cliente superaría el tope de concentración del ' + round2_(rate * 100) + '% del fondo total (' + fmtMoney_(cap) + ').' };
    }
  }
  return { ok: true, msg: '' };
}

/** V-20 — verifica que un ID Cliente exista en "Clientes". */
function clientExists_(clientId) {
  return clientesRows_().some(o => o.id === String(clientId || '').trim());
}
/** V-24 — verifica que el correo del cliente esté marcado como válido antes de enviar. */
function clientEmailIsValid_(clientId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS), col = ccol_('Correo válido', 9);
  const c = clientesRows_().filter(o => o.id === String(clientId || '').trim())[0];
  if (!c) return false;
  if (col) { const flag = String(sh.getRange(c.row, col).getValue()).trim().toUpperCase(); return flag === 'SÍ' || flag === 'SI'; }
  return vEmail_(c.email).ok;
}

/* ================= SETUP / VALIDACIONES DE DATOS ================= */

/** Punto de entrada manual: instala validaciones y crea la hoja "Validaciones". */
function configurarValidaciones() {
  return guard_('configurarValidaciones', function () {
    const ss = getSS_();
    ensureClientesSchema_(ss);
    ensureClienteBloqueoColumns_(ss);
    applySheetValidations_(ss);
    buildValidacionesSheet_(ss);
    SpreadsheetApp.flush();
    try { SpreadsheetApp.getActive().toast('Validaciones instaladas y hoja "Validaciones" creada.', 'Listo', 5); } catch (e) { }
    return 'OK';
  });
}

/** Garantiza el esquema extendido de "Clientes" (encabezados) SIN borrar datos. */
function ensureClientesSchema_(ss) {
  ss = ss || getSS_();
  const sh = getOrCreate_(ss, CFG.SHEETS.CLIENTS);
  const head = sh.getRange(1, 1, 1, CLIENTES_HEADERS_V2.length).getValues()[0].map(h => String(h).trim());
  if (head.join('|') !== CLIENTES_HEADERS_V2.join('|')) {
    sh.getRange(1, 1, 1, CLIENTES_HEADERS_V2.length).setValues([CLIENTES_HEADERS_V2])
      .setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff').setWrap(true);
    [90, 200, 230, 110, 140, 110, 220, 140, 100, 110, 260, 130, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  }
  sh.setFrozenRows(1);
  invalidateClientesHdrCache_();
  invalidateClientesRows_();
  return sh;
}

/**
 * V-33 — agrega (una sola vez) las columnas "Bloqueado" y "Motivo de bloqueo" AL FINAL
 * de "Clientes", por nombre de encabezado. NO se agregan a CLIENTES_HEADERS_V2: hojas
 * vivas pueden tener columnas extra al final (p. ej. "Forma de Pago" de Aplicar
 * novedades) y ensureClientesSchema_ reescribiría esos encabezados sobre los datos.
 */
function ensureClienteBloqueoColumns_(ss) {
  ss = ss || getSS_();
  const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  if (!cl) return;
  clientesHeaderCol_(cl, 'Bloqueado', true);
  clientesHeaderCol_(cl, 'Motivo de bloqueo', true);
  invalidateClientesHdrCache_();
  invalidateClientesRows_();
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
}

/** Instala validaciones de datos (listas / casillas) consistentes con el catálogo. */
function applySheetValidations_(ss) {
  ss = ss || getSS_();
  const N = CFG.MAX_ROWS;

  // Clientes: "Correo válido" (SÍ/NO), "Duplicado" (''/REVISAR/DUPLICADO) y "Bloqueado" (''/SÍ).
  const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  if (cl) {
    const cValid = ccol_('Correo válido', 9), cDup = ccol_('Duplicado', 10), cBloq = ccol_('Bloqueado', 0);
    if (cValid) cl.getRange(2, cValid, N, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['SÍ', 'NO'], true).setAllowInvalid(true).build());
    if (cDup) cl.getRange(2, cDup, N, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['', 'REVISAR', 'DUPLICADO'], true).setAllowInvalid(true).build());
    if (cBloq) cl.getRange(2, cBloq, N, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['', 'SÍ'], true).setAllowInvalid(true)
        .setHelpText('SÍ = cliente bloqueado (no puede solicitar ni recibir préstamos). Vaciar = desbloquear.').build());
  }

  // V-13 — Plazo en DÍAS (15/30/60) en "Prestatarios" (col D) y "Nuevos Prestatarios" (col G).
  const plazoDv = SpreadsheetApp.newDataValidation().requireValueInList(['15', '30', '60', '90'], true)
    .setAllowInvalid(false).setHelpText('Elegí un plazo válido: 15, 30 o 60 días.').build();
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (bs) bs.getRange(2, PB.TERM, N, 1).setDataValidation(plazoDv);
  const nb = ss.getSheetByName(CFG.SHEETS.NEW);
  if (nb) nb.getRange(2, 7, N, 1).setDataValidation(plazoDv); // col 7 = Plazo
}

/** Crea/actualiza la hoja "Validaciones" con el catálogo V-01…V-24. */
function buildValidacionesSheet_(ss) {
  ss = ss || getSS_();
  const name = 'Validaciones';
  const sh = getOrCreate_(ss, name);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const H = ['ID', 'Dónde', 'Campo', 'Regla', 'Mensaje al usuario', 'Acción', 'Prioridad'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff').setWrap(true);
  sh.getRange(2, 1, VALIDATION_RULES.length, H.length).setValues(VALIDATION_RULES).setWrap(true).setVerticalAlignment('top');
  [60, 110, 110, 340, 300, 110, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  // Colorea la columna "Acción".
  const acc = sh.getRange(2, 6, VALIDATION_RULES.length, 1);
  acc.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['BLOQUEA', 'ADVIERTE', 'AUTOMÁTICO'], true).setAllowInvalid(true).build());
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('BLOQUEA').setBackground('#f4cccc').setRanges([acc]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ADVIERTE').setBackground('#fff2cc').setRanges([acc]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('AUTOMÁTICO').setBackground('#d9ead3').setRanges([acc]).build(),
  ];
  sh.setConditionalFormatRules(rules);
  return sh;
}

/* ==================== FORMULARIO INTELIGENTE (HTML) ==================== */

function intakeSmartHtml_() {
  const lender = esc_(companyName_());
  const juris = esc_(getSetting_('Jurisdicción') || 'Buenos Aires, Argentina');
  const contractFooter = esc_(getSetting_('Pie del Contrato') || 'Este acuerdo es legalmente vinculante desde la firma de ambas partes.');
  const lateFeePct = Number(lateFeeRate_() * 100) || 5; // % de mora diario (para el popup de confirmación)
  const moraGrace = moraGraceDays_();                   // días de gracia antes de la mora
  const moraCapPct = round2_(moraCapFrac_() * 100);     // tope de mora como % del capital
  const moraClause = esc_(getSetting_('Cláusula de Mora') ||
    ('se aplica un recargo del ' + lateFeePctText_() + ' diario sobre el total a devolver por cada día de atraso posterior a la fecha de vencimiento' + (moraGrace > 0 ? ', tras ' + moraGrace + ' día(s) de gracia' : '') + ', con un tope acumulado del ' + moraCapPct + '% del capital.'));
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    :root{
      --brand:#1c4587;--brand-600:#16375f;--brand-050:#eaf1fb;
      --ok:#2e7d32;--ok-050:#e8f5e9;--danger:#c62828;--danger-050:#fdecec;
      --warn-bg:#fff8ec;--warn-bd:#f0b357;--warn-tx:#8a5300;
      --ink:#1f2733;--muted:#6b7686;--line:#e4e9f2;--field:#cfd6e4;
      --card:#fff;--radius:16px;--radius-sm:10px;
      --shadow:0 8px 30px rgba(23,55,95,.10),0 1px 3px rgba(23,55,95,.06);
      --ring:0 0 0 3px rgba(28,69,135,.18);
    }
    *{box-sizing:border-box}
    body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
      background:linear-gradient(180deg,#e9eef7 0,#eef1f7 240px,#eef1f7 100%);margin:0;padding:24px 16px;color:var(--ink);
      -webkit-font-smoothing:antialiased;line-height:1.5}
    .card{max-width:600px;margin:0 auto;background:var(--card);border-radius:var(--radius);padding:22px;box-shadow:var(--shadow);border-top:4px solid var(--brand)}
    @media(min-width:560px){.card{padding:32px 34px}}
    img{max-width:100%;height:auto}
    h1{font-size:22px;font-weight:700;color:var(--brand);margin:0 0 4px;letter-spacing:-.01em}
    .sub{color:var(--muted);margin:0 0 18px;font-size:14px}
    label{display:block;font-weight:600;margin:14px 0 6px;font-size:13px;color:#39424f}
    .req{color:var(--danger);font-weight:700}
    input:not([type=checkbox]):not([type=radio]):not([type=file]),textarea,select{width:100%;padding:12px 14px;border:1px solid var(--field);
      border-radius:var(--radius-sm);font-size:15px;color:var(--ink);background:#fff;font-family:inherit;transition:border-color .15s,box-shadow .15s;appearance:none;-webkit-appearance:none}
    input::placeholder,textarea::placeholder{color:#aab2c0}
    textarea{min-height:66px;resize:vertical}
    select{background-image:url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%236b7686' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpath d='M6 9l6 6 6-6'/%3e%3c/svg%3e");background-repeat:no-repeat;background-position:right 12px center;background-size:18px;padding-right:40px}
    input:not([type=checkbox]):not([type=radio]):not([type=file]):focus,textarea:focus,select:focus{outline:none;border-color:var(--brand);box-shadow:var(--ring)}
    input[readonly]{background:#eef1f6;color:#5b6675;cursor:not-allowed}
    input[type=file]{width:100%;padding:9px 12px;border:1px dashed var(--field);border-radius:var(--radius-sm);background:#fff;font-size:13.5px;color:var(--muted)}
    input[type=file]::file-selector-button{margin-right:10px;border:0;background:var(--brand-050);color:var(--brand);font-weight:600;padding:8px 12px;border-radius:8px;cursor:pointer;font-size:13px}
    input[type=file]::file-selector-button:hover{background:#dbe7fa}
    .hint{font-weight:400;color:var(--muted);font-size:12px}
    .row2{display:flex;gap:12px}.row2>div{flex:1;min-width:0}
    @media(max-width:460px){.row2{flex-direction:column;gap:0}}
    #identity,#references,#photos,#loanData{border:1px solid var(--line);border-radius:12px;padding:4px 16px 16px;margin:16px 0;background:#fcfdff}
    .secTitle{font-size:15px;font-weight:700;color:var(--brand);margin:14px 0 2px}
    #calc{margin-top:10px;padding:12px 14px;background:var(--brand-050);border:1px solid #d5e3f7;border-radius:var(--radius-sm);font-size:14px;display:none}
    #calc b{color:var(--brand)}
    button{margin-top:20px;width:100%;background:var(--brand);color:#fff;border:0;border-radius:var(--radius-sm);
      padding:14px;font-size:16px;font-weight:700;cursor:pointer;transition:background .15s,transform .05s,box-shadow .15s;box-shadow:0 2px 10px rgba(28,69,135,.28)}
    button:hover{background:var(--brand-600)}
    button:active{transform:translateY(1px)}
    button:focus-visible{outline:none;box-shadow:var(--ring),0 2px 10px rgba(28,69,135,.28)}
    button:disabled{background:#9db4d6;box-shadow:none;cursor:default}
    .link{background:none;color:var(--brand);text-decoration:underline;font-weight:500;width:auto;padding:0;margin:8px 0 0;font-size:13px;box-shadow:none}
    .link:hover{background:none;color:var(--brand-600)}
    #msg{margin-top:16px;padding:12px 14px;border-radius:var(--radius-sm);display:none;font-size:14px;border-left:4px solid transparent}
    .ok{background:var(--ok-050);color:#1e5e22;border-left-color:var(--ok)}
    .err{background:var(--danger-050);color:#a01c1c;border-left-color:var(--danger)}
    .welcome{margin:8px 0 6px;padding:12px 14px;border-radius:12px;font-size:14px}
    .welcome.ret{background:var(--ok-050);color:#1e5e22;border:1px solid #cfe6d0}
    .welcome.newc{background:var(--brand-050);color:#28517f;border:1px solid #d5e3f7}
    .warnbox{margin:4px 0 8px;padding:11px 13px;background:var(--warn-bg);border:1px solid var(--warn-bd);border-left:4px solid #e8891d;border-radius:10px;font-size:13px;color:var(--warn-tx);line-height:1.45}
    .warnbox b{color:#a85800}
    #loadingView{display:none;text-align:center;padding:34px 12px}
    .pbar-track{width:100%;height:22px;background:#e3e9f4;border-radius:12px;overflow:hidden;margin:14px 0 8px}
    .pbar-fill{height:100%;width:0%;border-radius:12px;transition:width .45s ease;background-color:var(--brand);
      background-image:linear-gradient(45deg,rgba(255,255,255,.30) 25%,transparent 25%,transparent 50%,rgba(255,255,255,.30) 50%,rgba(255,255,255,.30) 75%,transparent 75%,transparent);
      background-size:28px 28px;animation:pbar-stripes 1s linear infinite}
    @keyframes pbar-stripes{from{background-position:0 0}to{background-position:28px 0}}
    @media(prefers-reduced-motion:reduce){.pbar-fill{animation:none}button:active{transform:none}}
    #successView{display:none}
    .check{width:76px;height:76px;line-height:76px;margin:6px auto 12px;border-radius:50%;background:var(--ok);color:#fff;font-size:42px;text-align:center;box-shadow:0 6px 18px rgba(46,125,50,.35)}
    .summary{width:100%;border-collapse:separate;border-spacing:0;margin:18px 0;border:1px solid var(--line);border-radius:12px;overflow:hidden}
    .summary td{padding:10px 13px;border-bottom:1px solid var(--line);font-size:14px}
    .summary tr:last-child td{border-bottom:0}
    .summary td.k{background:#f5f8fd;font-weight:600;width:46%;color:#39424f}
    .note{background:var(--brand-050);border-radius:12px;padding:12px 14px;font-size:13px;color:#28517f;margin-top:6px}
    .terms{margin-top:16px;border:1px solid var(--line);border-radius:12px;background:#fcfdff;overflow:hidden}
    .terms h2{font-size:14px;color:var(--brand);margin:0;padding:11px 14px;border-bottom:1px solid var(--line);background:#f5f8fd;font-weight:700}
    .terms .body{max-height:160px;overflow-y:auto;padding:10px 14px;font-size:12.5px;color:#4a5563;line-height:1.55}
    .terms ul{margin:6px 0 0;padding-left:18px}.terms li{margin:6px 0}
    .terms li.mora{color:#8a5300}.terms li.mora b{color:#b45309}
    .agree{display:flex;align-items:flex-start;gap:10px;margin:10px 0 0;padding:11px 13px;font-weight:400;font-size:13.5px;cursor:pointer;
      border:1px solid var(--line);border-radius:10px;background:#fff;transition:border-color .15s,background .15s}
    .agree:hover{border-color:#c3cfe2;background:#fbfcfe}
    .agree input{width:20px;height:20px;flex:0 0 auto;margin-top:0;accent-color:var(--brand);cursor:pointer}
    .terms .agree{margin:10px 12px}
    .fieldErr{display:none;color:var(--danger);font-weight:500;font-size:12px;margin:5px 0 0}
    .fieldErr.show{display:block}
    input.invalid,select.invalid,textarea.invalid{border-color:var(--danger);background:var(--danger-050)}
    input.invalid:focus,select.invalid:focus,textarea.invalid:focus{box-shadow:0 0 0 3px rgba(198,40,40,.15)}
    .agree.invalid{border-color:var(--danger);background:var(--danger-050);color:#a01c1c}
    /* ===== Asistente por pasos (wizard) ===== */
    .wstep{display:none;animation:stepIn .25s ease}.wstep.active{display:block}
    @keyframes stepIn{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:none}}
    #stepper{display:flex;gap:6px;margin:4px 0 8px;overflow-x:auto;-webkit-overflow-scrolling:touch;padding-bottom:2px}
    .schip{display:flex;align-items:center;gap:7px;padding:7px 12px;border:1px solid var(--line);border-radius:999px;background:#fff;font-size:12px;font-weight:600;color:var(--muted);cursor:pointer;white-space:nowrap;transition:background .15s,border-color .15s,color .15s}
    .schip .snum{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;background:#e4e9f2;color:#5b6675;font-size:11px;font-weight:700;flex:0 0 auto}
    .schip:hover{border-color:#c3cfe2}
    .schip.done{color:#1e5e22;border-color:#cfe6d0;background:var(--ok-050)}.schip.done .snum{background:var(--ok);color:#fff}
    .schip.active{color:#fff;background:var(--brand);border-color:var(--brand)}.schip.active .snum{background:#fff;color:var(--brand)}
    @media(max-width:520px){.schip .slbl{display:none}.schip{padding:7px}}
    #progress{height:6px;background:#e4e9f2;border-radius:999px;overflow:hidden;margin:0 0 6px}
    #progressFill{height:100%;width:0;background:var(--brand);border-radius:999px;transition:width .3s ease}
    #stepHint{font-size:12px;color:var(--muted);font-weight:600;margin:2px 0 12px}
    #wnav{display:flex;gap:10px;margin-top:18px}
    #prevBtn{background:#fff;color:var(--brand);border:1px solid var(--field);box-shadow:none;flex:0 0 auto;width:auto;padding:14px 18px;margin-top:0}
    #prevBtn:hover{background:#f5f8fd}
    #nextBtn{margin-top:0}
    @media(prefers-reduced-motion:reduce){.wstep{animation:none}}</style></head><body>
    <div class="card" id="card">${brandHeaderHtml_()}
    <div id="formView">
      <h1>Solicitud de Préstamo</h1>
      <p class="sub">${lender} — ingresá tu <b>correo</b> y <b>DNI</b> para empezar. Si ya sos cliente, completamos tus datos automáticamente (solo faltarán las fotos).</p>
      <form id="f" novalidate autocomplete="off">
        <div id="rest">
          <input type="hidden" name="mode" id="mode" value="new">
          <input type="hidden" name="clientId" id="clientId" value="">
          <input type="hidden" name="reuseRefs" id="reuseRefs" value="">
          <div id="stepper">
            <div class="schip active" data-goto="1"><span class="snum">1</span><span class="slbl">Identificación</span></div>
            <div class="schip" data-goto="2"><span class="snum">2</span><span class="slbl">Tus datos</span></div>
            <div class="schip" data-goto="3"><span class="snum">3</span><span class="slbl">Referencias</span></div>
            <div class="schip" data-goto="4"><span class="snum">4</span><span class="slbl">El préstamo</span></div>
            <div class="schip" data-goto="5"><span class="snum">5</span><span class="slbl">Documentación</span></div>
            <div class="schip" data-goto="6"><span class="snum">6</span><span class="slbl">Confirmación</span></div>
          </div>
          <div id="progress"><div id="progressFill"></div></div>
          <div id="stepHint"></div>

          <!-- Paso 1: Identificación -->
          <div class="wstep active" data-step="1">
            <div class="secTitle">Identificación</div>
          <div class="warnbox">⚠️ Verificá que tu <b>correo</b> y <b>DNI</b> sean correctos: si tienen un error, no podremos completar el proceso.</div>
          <label>Correo electrónico <span class="req">*</span><input name="email" id="email" type="email" required></label>
          <small class="fieldErr" id="e_email"></small>
          <label>DNI <span class="req">*</span><input name="dni" id="dni" required placeholder="XX.XXX.XXX" inputmode="numeric"></label>
          <small class="fieldErr" id="e_dni"></small>
          <div id="lookupMsg" style="display:none;margin-top:12px;padding:10px;border-radius:8px;font-size:13.5px"></div>
          </div><!-- /paso 1 -->

          <!-- Paso 2: Tus datos -->
          <div class="wstep" data-step="2">
          <div id="welcome" class="welcome"></div>

          <div id="identity">
            <div class="secTitle">Tus datos</div>
            <label>Nombre completo <span class="req" id="rqName">*</span><input name="fullName" id="fullName" autocomplete="off"></label>
            <small class="fieldErr" id="e_fullName"></small>
            <div class="row2">
              <div><label>Teléfono <span class="req" id="rqPhone">*</span><input name="phone" id="phone" inputmode="tel" autocomplete="off" placeholder="11 XXXX XXXX"></label>
                <small class="fieldErr" id="e_phone"></small></div>
              <div><label>CUIL <span class="req" id="rqCuil">*</span><input name="cuil" id="cuil" autocomplete="off" placeholder="XX-XXXXXXXX-X" inputmode="numeric"></label>
                <small class="fieldErr" id="e_cuil"></small></div>
            </div>
            <label>Dirección <span class="req" id="rqAddr">*</span><input name="address" id="address" autocomplete="off" placeholder="Calle, número, localidad, provincia"></label>
            <small class="fieldErr" id="e_address"></small>
          </div>
          </div><!-- /paso 2 -->

          <!-- Paso 3: Referencias -->
          <div class="wstep" data-step="3">
          <!-- Referencias registradas (clientes que vuelven): se reutilizan, no se re-piden -->
          <div id="refsNote" class="welcome ret" style="display:none;margin:6px 0"></div>
          <!-- Referencias (obligatorias para nuevos o existentes sin referencias en archivo) -->
          <div id="references">
            <div class="secTitle">Referencias</div>
            <div class="hint" style="margin-bottom:6px">Dos personas que puedan confirmar tus datos. La 1.ª debe ser un familiar directo que <b>no viva con vos</b>; la 2.ª, alguien no familiar (laboral o de confianza).</div>
            <div class="warnbox">⚠️ La información de las referencias debe ser <b>válida y verificable</b> (nombre y teléfono reales). Si no podemos confirmarlas, <b>no se aprobará el préstamo</b>.</div>
            <label>Referencia 1 — Nombre y apellido <span class="req">*</span><input name="ref1Name" id="ref1Name" autocomplete="off" placeholder="Nombre del familiar"></label>
            <small class="fieldErr" id="e_ref1Name"></small>
            <div class="row2">
              <div><label>Vínculo <span class="req">*</span><select name="ref1Rel" id="ref1Rel">
                <option value="">— elegir —</option>
                <option value="Padre/Madre">Padre / Madre</option>
                <option value="Hermano/a">Hermano/a</option>
                <option value="Hijo/a">Hijo/a</option>
                <option value="Otro familiar">Otro familiar (no conviviente)</option></select></label>
                <small class="fieldErr" id="e_ref1Rel"></small></div>
              <div><label>Teléfono <span class="req">*</span><input name="ref1Phone" id="ref1Phone" inputmode="tel" autocomplete="off" placeholder="11 XXXX XXXX"></label>
                <small class="fieldErr" id="e_ref1Phone"></small></div>
            </div>
            <label>Referencia 2 — Nombre y apellido <span class="req">*</span><input name="ref2Name" id="ref2Name" autocomplete="off" placeholder="Nombre de la referencia"></label>
            <small class="fieldErr" id="e_ref2Name"></small>
            <div class="row2">
              <div><label>Vínculo <span class="req">*</span><select name="ref2Rel" id="ref2Rel">
                <option value="">— elegir —</option>
                <option value="Empleador/Jefe">Empleador / Jefe</option>
                <option value="Compañero de trabajo">Compañero de trabajo</option>
                <option value="Vecino/a">Vecino/a</option>
                <option value="Amigo/a de larga data">Amigo/a de larga data</option>
                <option value="Otro">Otro (no familiar)</option></select></label>
                <small class="fieldErr" id="e_ref2Rel"></small></div>
              <div><label>Teléfono <span class="req">*</span><input name="ref2Phone" id="ref2Phone" inputmode="tel" autocomplete="off" placeholder="11 XXXX XXXX"></label>
                <small class="fieldErr" id="e_ref2Phone"></small></div>
            </div>
          </div>
          </div><!-- /paso 3 -->

          <!-- Paso 4: Datos del préstamo -->
          <div class="wstep" data-step="4">
          <!-- Datos del préstamo (todos) -->
          <div id="loanData">
          <div class="secTitle">Datos del préstamo</div>
          <div style="font-size:13px;color:#4a5563;line-height:1.6;margin:2px 0 10px">
            <div style="margin-bottom:5px;color:#39424f">El plazo se calcula según el monto:</div>
            <ul id="tierList" style="margin:0;padding-left:18px">
              <li id="tierLi1">Hasta <b>${fmtMoney_(tier1Max_())}</b>: elegís <b style="color:#1c4587">15 días (25%)</b> o <b style="color:#1c4587">30 días (50%)</b> — <b>1 pago</b>.</li>
              <li id="tierLi2">De <b>${fmtMoney_(tier1Max_())}</b> a <b>${fmtMoney_(tier2Max_())}</b>: <b style="color:#1c4587">30 días (50%)</b> — <b>1 pago</b>.</li>
              <li id="tierLi3">De <b>${fmtMoney_(tier2Max_())}</b> a <b>${fmtMoney_(loanMax_())}</b>: <b style="color:#1c4587">90 días (100%)</b> — <b>3 cuotas mensuales</b> (vencen día 30, 60 y 90).</li>
            </ul>
            <div id="maxBanner" style="margin-top:7px;padding:7px 11px;background:#fff4e5;border:1px solid #f0b357;border-radius:8px;color:#7a4a00">⚠️ <b>Monto máximo por préstamo: ${fmtMoney_(loanMax_())}</b></div>
          </div>
          <label>Monto solicitado (ARS) <span class="req">*</span><input name="amount" id="amount" type="text" inputmode="numeric" autocomplete="off" placeholder="Ej: 100.000" required style="font-size:20px;font-weight:bold;height:auto;padding:12px"></label>
          <div id="amountWords" style="display:none;margin:6px 0 2px;font-size:19px;font-weight:bold;color:#1c4587;line-height:1.35"></div>
          <small class="fieldErr" id="e_amount"></small>
          <input type="hidden" name="term" id="term" value="">
          <label>Plazo <span class="hint">— según el monto (en montos chicos podés elegir)</span></label>
          <div id="termView" style="padding:12px 14px;border:1px solid #cfd6e4;border-radius:10px;background:#eef1f6;color:#5b6675;font-weight:600">Ingresá el monto para ver el plazo y las cuotas.</div>
          <div id="termChoice" style="display:none;margin-top:8px">
            <span style="font-weight:600;font-size:13px;color:#39424f">Elegí el plazo:</span>
            <select id="term15o30" style="width:auto;display:inline-block;margin-left:8px">
              <option value="15">15 días (25% de interés)</option>
              <option value="30">30 días (50% de interés)</option>
            </select>
          </div>
          <small class="fieldErr" id="e_term"></small>
          <div id="calc"></div>
          <label>Notas / Motivo <span class="req">*</span><textarea name="notes" id="notes" rows="2" required></textarea></label>
          <small class="fieldErr" id="e_notes"></small>

          <label>Forma de pago preferida <span class="req">*</span><select name="payMethod" id="payMethod" required>
            <option value="Mercado Pago" selected>Mercado Pago</option>
            <option value="Transferencia bancaria">Transferencia bancaria</option>
            <option value="Efectivo">Efectivo</option>
            <option value="Otro">Otro</option></select></label>
          <label>Datos de la cuenta / alias <small style="font-weight:normal;color:#888">— alias, CBU/CVU o el medio que prefiera (opcional)</small>
            <input name="payDetails" id="payDetails" type="text" autocomplete="off" placeholder="Ej: alias.mercadopago o CBU"></label>
          </div>
          </div><!-- /paso 4 -->

          <!-- Paso 5: Documentación -->
          <div class="wstep" data-step="5">
          <!-- Fotos (obligatorias para nuevos) -->
          <div id="photos">
            <div class="secTitle">Documentación</div>
            <label>Foto del frente del DNI <span class="req" id="rqDni">*</span> <span class="hint">— imagen o PDF, nítida</span><input name="dniPhoto" id="dniPhoto" type="file" accept="image/*,.pdf" capture="environment"></label>
            <small class="fieldErr" id="e_dniPhoto"></small>
            <label>Foto del dorso del DNI <span class="req" id="rqDorso">*</span> <span class="hint">— imagen o PDF, nítida</span><input name="cuilPhoto" id="cuilPhoto" type="file" accept="image/*,.pdf" capture="environment"></label>
            <small class="fieldErr" id="e_cuilPhoto"></small>
          </div>
          </div><!-- /paso 5 -->

          <!-- Paso 6: Confirmación -->
          <div class="wstep" data-step="6">
          <div class="terms">
            <h2>Términos y Condiciones del Préstamo</h2>
            <div class="body"><ul>
              <li><b>Intereses según el plazo:</b> 15 días = 25%, 30 días = 50%, 60 días = 100% sobre el capital.</li>
              <li><b>Devolución:</b> el capital más los intereses se devuelven en su totalidad en la fecha de vencimiento.</li>
              <li class="mora">⚠️ <b>Recargo por mora:</b> ${moraClause}</li>
              <li><b>Verificación:</b> la solicitud está sujeta a revisión y verificación de identidad (DNI y CUIL).</li>
              <li><b>Declaración:</b> declaro que los datos y documentos aportados son verídicos y de mi titularidad.</li>
              <li><b>Jurisdicción:</b> ${juris}. ${contractFooter}</li>
            </ul></div>
          </div>
          <div class="terms" style="border-color:#b45309">
            <h2 style="color:#b45309">Compromiso de pago</h2>
            <label class="agree" id="agreeRepayLabel"><input type="checkbox" name="agreeRepay" id="agreeRepay" required>
              <span>Me comprometo a devolver el <b>monto total (capital + interés)</b> en la fecha de vencimiento. <span class="req">*</span></span></label>
            <small class="fieldErr" id="e_agreeRepay"></small>
            <label class="agree" id="agreeMoraLabel"><input type="checkbox" name="agreeMora" id="agreeMora" required>
              <span>Entiendo el <b>recargo por mora (${lateFeePctText_()} por día${moraGrace > 0 ? `, tras ${moraGrace} día(s) de gracia` : ''}, con tope del ${moraCapPct}% del capital)</b> si pago después del vencimiento. <span class="req">*</span></span></label>
            <small class="fieldErr" id="e_agreeMora"></small>
            <label class="agree" id="agreeConseqLabel"><input type="checkbox" name="agreeConseq" id="agreeConseq" required>
              <span>Entiendo que, ante el impago, se iniciarán <b>acciones legales</b> para el cobro de la deuda. <span class="req">*</span></span></label>
            <small class="fieldErr" id="e_agreeConseq"></small>
            <label class="agree" id="agreeContactLabel"><input type="checkbox" name="agreeContact" id="agreeContact" required>
              <span>Entiendo que si me atraso y <b>no me comunico</b> ni <b>acuerdo un plan de pago</b>, <b>no se me otorgarán préstamos futuros</b>. <span class="req">*</span></span></label>
            <small class="fieldErr" id="e_agreeContact"></small>
          </div>
          <label class="agree" id="agreeLabel"><input type="checkbox" name="agree" id="agree" required>
            <span>He leído y acepto los <b>Términos y Condiciones</b> del préstamo. <span class="req">*</span></span></label>
          <small class="fieldErr" id="e_agree"></small>
          <button type="submit" id="btn">Enviar solicitud</button>
          </div><!-- /paso 6 -->

          <!-- Navegación del asistente -->
          <div id="wnav">
            <button type="button" id="prevBtn">← Anterior</button>
            <button type="button" id="nextBtn">Siguiente →</button>
          </div>
        </div>
      </form>
      <div id="msg"></div>
    </div>
    <div id="confirmOverlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:50;align-items:center;justify-content:center;padding:16px">
      <div style="background:#fff;border-radius:10px;max-width:460px;width:100%;padding:22px 24px;box-shadow:0 6px 30px rgba(0,0,0,.25)">
        <h2 style="margin:0 0 10px;color:#1c4587;font-size:18px">Confirmá tu solicitud</h2>
        <div id="confirmBody" style="font-size:14px;color:#333;line-height:1.55"></div>
        <div style="display:flex;gap:10px;margin-top:18px">
          <button type="button" id="confirmCancel" style="flex:1;background:#e0e0e0;color:#333">Cancelar</button>
          <button type="button" id="confirmOk" style="flex:1">Confirmar y enviar</button>
        </div>
      </div>
    </div>
    <div id="loadingView">
      <p id="loadPhase" style="font-weight:bold;color:#1c4587;margin:0;font-size:16px">Enviando su solicitud…</p>
      <div class="pbar-track"><div class="pbar-fill" id="pbarFill"></div></div>
      <p id="loadPct" style="color:#1c4587;font-weight:bold;margin:0;font-size:20px">0%</p></div>
    <div id="successView"></div>
    </div>
    <script>
      var f=document.getElementById('f'),btn=document.getElementById('btn'),msg=document.getElementById('msg'),calc=document.getElementById('calc');
      var formView=document.getElementById('formView'),loadingView=document.getElementById('loadingView'),successView=document.getElementById('successView');
      var step1=document.getElementById('step1'),rest=document.getElementById('rest'),lookupMsg=document.getElementById('lookupMsg');
      var emailEl=document.getElementById('email'),dniEl=document.getElementById('dni'),modeEl=document.getElementById('mode'),clientIdEl=document.getElementById('clientId');
      var identity=document.getElementById('identity'),photos=document.getElementById('photos'),welcome=document.getElementById('welcome');
      var amountEl=document.getElementById('amount'),termEl=document.getElementById('term');
      // Tramos por monto (el monto determina plazo/tasa/cuotas; se calcula, no se elige).
      var T1MAX=${tier1Max_()}, T2MAX=${tier2Max_()}, LOANMAX=${loanMax_()};
      // Límite del prestatario (escalera de graduación, V-32). Arranca en el tope global y se
      // ajusta al identificarse (según su historial de repago). Nunca supera LOANMAX.
      var MAXAMT=LOANMAX;
      function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
      function fmt(n){try{return n.toLocaleString('es-AR',{style:'currency',currency:'ARS'});}catch(e){return '$'+n.toFixed(2);}}
      function groupNum(d){var out='',c=0;for(var i=d.length-1;i>=0;i--){out=d.charAt(i)+out;c++;if(c%3===0&&i>0){out='.'+out;}}return out;}
      // Fechas de las cuotas PROYECTADAS desde hoy (la fecha real corre desde la firma).
      function fmtDateJS(d){try{return d.toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit',year:'numeric'});}catch(e){return d.getDate()+'/'+(d.getMonth()+1)+'/'+d.getFullYear();}}
      function addDaysJS(base,n){var x=new Date(base.getTime());x.setDate(x.getDate()+n);return x;}
      function cuotaDatesTxt(nc,days){var t=new Date();t.setHours(0,0,0,0);
        if(nc<=1){return 'vence el '+fmtDateJS(addDaysJS(t,days))+' (aprox., '+days+' días desde la firma)';}
        var parts=[];for(var k=1;k<=nc;k++){parts.push('cuota '+k+': '+fmtDateJS(addDaysJS(t,30*k)));}
        return 'vencen '+parts.join(' · ')+' (aprox., desde la firma)';}
      function capFirst(s){return s?s.charAt(0).toUpperCase()+s.slice(1):s;}
      // Número entero (pesos) a letras en español (0..999.999.999).
      function numeroALetras(n){
        n=Math.floor(Math.abs(Number(n)||0)); if(n===0)return 'cero';
        var U=['','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez','once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve','veinte','veintiuno','veintidós','veintitrés','veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve'];
        var D=['','','','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'];
        var C=['','ciento','doscientos','trescientos','cuatrocientos','quinientos','seiscientos','setecientos','ochocientos','novecientos'];
        function sec(x){ if(x===100)return 'cien'; var c=Math.floor(x/100),r=x%100,s=''; if(c)s+=C[c]+' '; if(r<30)s+=U[r]; else{var d=Math.floor(r/10),u=r%10;s+=D[d]+(u?' y '+U[u]:'');} return s.trim(); }
        function apoc(s){ return s.replace(/veintiuno$/,'veintiún').replace(/(^|\s)uno$/,'$1un'); }
        var partes=[], mill=Math.floor(n/1000000), miles=Math.floor((n%1000000)/1000), resto=n%1000;
        if(mill)partes.push(mill===1?'un millón':apoc(sec(mill))+' millones');
        if(miles)partes.push(miles===1?'mil':apoc(sec(miles))+' mil');
        if(resto)partes.push(apoc(sec(resto)));
        return partes.join(' ').replace(/\s+/g,' ').trim();
      }
      function updateAmountWords(){var a=parseAmount(amountEl.value),w=document.getElementById('amountWords'); if(!w)return;
        if(a>0){w.style.display='block';w.textContent=capFirst(numeroALetras(a))+(a===1?' peso':' pesos');}else{w.style.display='none';w.textContent='';}}
      function digitsOnly(s){return String(s==null?'':s).replace(/[^0-9]/g,'');}
      function pad8(s){s=String(s);while(s.length<8)s='0'+s;return s;}
      function isNew(){return modeEl.value!=='existing';}

      /* ===== Validadores del lado del cliente (espejo del servidor, sin backslash en regex) ===== */
      function vEmailC(v){var n=String(v==null?'':v).trim().toLowerCase();
        if(!n)return 'El correo electrónico es obligatorio.';
        if(!/^[^ @]+@[^ @]+[.][a-z]{2,}$/i.test(n))return 'El correo no es válido. Verificá que termine en .com, .ar, etc.';return '';}
      function vDniC(v){var n=digitsOnly(v);if(!n)return 'El DNI es obligatorio.';
        if(n.length<7||n.length>8)return 'Ingresá un DNI válido (7 u 8 dígitos, sin puntos).';return '';}
      function normPhoneC(v){var d=digitsOnly(v);if(!d)return '';
        if(d.length>10&&d.indexOf('54')===0)d=d.slice(2);
        if(d.length===11&&d.charAt(0)==='9')d=d.slice(1);
        if(d.length!==10)return '';return '+54'+d;}
      function vPhoneC(v){if(!String(v==null?'':v).trim())return 'El teléfono es obligatorio.';
        if(!normPhoneC(v))return 'Ingresá un teléfono válido de 10 dígitos.';return '';}
      function vNameC(v){if(!String(v==null?'':v).trim())return 'El nombre completo es obligatorio.';return '';}
      function cuilDvC(fb){var m=[5,4,3,2,7,6,5,4,3,2],s=0;for(var i=0;i<10;i++)s+=parseInt(fb.charAt(i),10)*m[i];var r=11-(s%11);if(r===11)r=0;if(r===10)r=9;return r;}
      function vCuilC(v,dni){var c=digitsOnly(v);if(!c)return 'El CUIL es obligatorio.';
        if(c.length!==11)return 'El CUIL debe tener 11 dígitos.';
        if(['20','23','24','27'].indexOf(c.slice(0,2))<0)return 'El CUIL no es válido o no corresponde al DNI ingresado.';
        if(cuilDvC(c.slice(0,10))!==parseInt(c.charAt(10),10))return 'El CUIL no es válido o no corresponde al DNI ingresado.';
        var dn=digitsOnly(dni);if(dn&&c.slice(2,10)!==pad8(dn))return 'El CUIL no es válido o no corresponde al DNI ingresado.';return '';}
      function vAmountC(v){var n=digitsOnly(v);if(!n||parseInt(n,10)<=0)return 'Ingresá un monto de préstamo válido.';
        if(parseInt(n,10)>MAXAMT)return (MAXAMT<LOANMAX?'Tu límite actual es ':'El monto máximo por préstamo es ')+fmt(MAXAMT)+'.';return '';}
      function vTermC(v){if(['15','30','60','90'].indexOf(String(v))<0)return 'El plazo se calcula según el monto. Ingresá un monto válido.';return '';}
      function vNotesC(v){if(!String(v==null?'':v).trim())return 'Las notas / motivo son obligatorias.';return '';}
      function vAddressC(v){if(!String(v==null?'':v).trim())return 'La dirección es obligatoria.';return '';}
      function vPhotoC(input,required){var fs=input&&input.files;if(!fs||!fs.length)return required?'Subí la foto del DNI (imagen o PDF).':'';return '';}
      function vRefNameC(v){if(!String(v==null?'':v).trim())return 'El nombre de la referencia es obligatorio.';return '';}
      function vRefRelC(v){if(!String(v==null?'':v).trim())return 'Elegí el vínculo de la referencia.';return '';}
      function vRefPhoneC(v){if(!String(v==null?'':v).trim())return 'El teléfono de la referencia es obligatorio.';
        if(!normPhoneC(v))return 'Ingresá un teléfono válido de 10 dígitos.';return '';}

      /* ===== Muestra/oculta el error de un campo de inmediato ===== */
      function setErr(key,m,el){var e=document.getElementById('e_'+key);
        if(e){e.textContent=m||'';if(m)e.classList.add('show');else e.classList.remove('show');}
        if(el){if(m)el.classList.add('invalid');else el.classList.remove('invalid');}
        return !m;}
      function el_(id){return document.getElementById(id);}
      function checkField(key){
        switch(key){
          case 'email':return setErr('email',vEmailC(emailEl.value),emailEl);
          case 'dni':return setErr('dni',vDniC(dniEl.value),dniEl);
          case 'fullName':return setErr('fullName',vNameC(el_('fullName').value),el_('fullName'));
          case 'phone':return setErr('phone',vPhoneC(el_('phone').value),el_('phone'));
          case 'cuil':return setErr('cuil',vCuilC(el_('cuil').value,dniEl.value),el_('cuil'));
          case 'address':return setErr('address',vAddressC(el_('address').value),el_('address'));
          case 'amount':return setErr('amount',vAmountC(amountEl.value),amountEl);
          case 'term':return setErr('term',vTermC(termEl.value),termEl);
          case 'notes':return setErr('notes',vNotesC(el_('notes').value),el_('notes'));
          case 'dniPhoto':return setErr('dniPhoto',vPhotoC(el_('dniPhoto'),true),el_('dniPhoto'));
          case 'cuilPhoto':return setErr('cuilPhoto',vPhotoC(el_('cuilPhoto'),true),el_('cuilPhoto'));
          case 'ref1Name':return setErr('ref1Name',vRefNameC(el_('ref1Name').value),el_('ref1Name'));
          case 'ref1Rel':return setErr('ref1Rel',vRefRelC(el_('ref1Rel').value),el_('ref1Rel'));
          case 'ref1Phone':return setErr('ref1Phone',vRefPhoneC(el_('ref1Phone').value),el_('ref1Phone'));
          case 'ref2Name':return setErr('ref2Name',vRefNameC(el_('ref2Name').value),el_('ref2Name'));
          case 'ref2Rel':return setErr('ref2Rel',vRefRelC(el_('ref2Rel').value),el_('ref2Rel'));
          case 'ref2Phone':return setErr('ref2Phone',vRefPhoneC(el_('ref2Phone').value),el_('ref2Phone'));
          case 'agreeRepay':{var okR=el_('agreeRepay').checked;setErr('agreeRepay',okR?'':'Debe aceptar el compromiso de devolución.');el_('agreeRepayLabel').classList.toggle('invalid',!okR);return okR;}
          case 'agreeMora':{var okM=el_('agreeMora').checked;setErr('agreeMora',okM?'':'Debe aceptar el recargo por mora.');el_('agreeMoraLabel').classList.toggle('invalid',!okM);return okM;}
          case 'agreeConseq':{var okC=el_('agreeConseq').checked;setErr('agreeConseq',okC?'':'Debe aceptar las consecuencias del impago.');el_('agreeConseqLabel').classList.toggle('invalid',!okC);return okC;}
          case 'agreeContact':{var okK=el_('agreeContact').checked;setErr('agreeContact',okK?'':'Debe aceptar la condición sobre comunicación y préstamos futuros.');el_('agreeContactLabel').classList.toggle('invalid',!okK);return okK;}
          case 'agree':{var ok=el_('agree').checked;setErr('agree',ok?'':'Debe aceptar los Términos y Condiciones para enviar la solicitud.');el_('agreeLabel').classList.toggle('invalid',!ok);return ok;}
        }
        return true;
      }
      // Validación en vivo: revisa al salir del campo (blur/change) y mientras se corrige un error visible.
      function bindLive(key,el){if(!el)return;
        el.addEventListener('blur',function(){checkField(key);});
        el.addEventListener('change',function(){checkField(key);});
        el.addEventListener('input',function(){var e=document.getElementById('e_'+key);if(e&&e.classList.contains('show'))checkField(key);});}
      ['email','dni'].forEach(function(k){bindLive(k,el_(k==='email'?'email':'dni'));});
      bindLive('fullName',el_('fullName'));bindLive('phone',el_('phone'));bindLive('cuil',el_('cuil'));bindLive('address',el_('address'));
      bindLive('amount',amountEl);bindLive('term',termEl);bindLive('notes',el_('notes'));
      bindLive('dniPhoto',el_('dniPhoto'));bindLive('cuilPhoto',el_('cuilPhoto'));
      ['ref1Name','ref1Rel','ref1Phone','ref2Name','ref2Rel','ref2Phone'].forEach(function(k){bindLive(k,el_(k));});
      ['agreeRepay','agreeMora','agreeConseq','agreeContact','agree'].forEach(function(k){el_(k).addEventListener('change',function(){checkField(k);});});
      function clearErrors(keys){keys.forEach(function(k){setErr(k,'',el_(k));});}

      // Requerido dinámico por modo.
      function setRequired(el,on){if(!el)return;if(on)el.setAttribute('required','required');else el.removeAttribute('required');}
      function markReqStars(on){['rqName','rqPhone','rqCuil','rqAddr','rqDni','rqDorso'].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display=on?'inline':'none';});}
      // Referencias: si el cliente ya tiene DOS referencias en archivo, se reutilizan (se ocultan y no se piden).
      function applyRefs(r){
        var refs=el_('references'),note=el_('refsNote');
        if(r&&r.found&&r.hasRefs&&r.client){
          refs.style.display='none';el_('reuseRefs').value='1';
          clearErrors(['ref1Name','ref1Rel','ref1Phone','ref2Name','ref2Rel','ref2Phone']);
          var c=r.client;
          note.style.display='block';
          note.innerHTML='<b>Referencias en archivo</b> — no hace falta cargarlas de nuevo:<br>'+
            '1) '+esc(c.ref1Name)+(c.ref1Rel?' ('+esc(c.ref1Rel)+')':'')+' · '+esc(c.ref1Phone)+'<br>'+
            '2) '+esc(c.ref2Name)+(c.ref2Rel?' ('+esc(c.ref2Rel)+')':'')+' · '+esc(c.ref2Phone);
        }else{
          refs.style.display='';el_('reuseRefs').value='';note.style.display='none';
        }
      }

      // Paso 1 (Identificación): valida correo+DNI y consulta el cliente en el servidor.
      var identified=false;
      [emailEl,dniEl].forEach(function(el){ if(el) el.addEventListener('input',function(){ identified=false; }); });
      function runLookup(){
        lookupMsg.style.display='none';
        var okE=checkField('email'),okD=checkField('dni');
        if(!okE||!okD){(okE?dniEl:emailEl).focus();return;}
        var nb=document.getElementById('nextBtn'); if(nb){nb.disabled=true;nb.textContent='Verificando…';}
        google.script.run.withSuccessHandler(onLookup).withFailureHandler(function(e){
          if(nb){nb.disabled=false;nb.textContent='Verificar y continuar →';}
          lookupMsg.style.background='#fdecec';lookupMsg.style.color='#a01c1c';lookupMsg.textContent=(e.message||e);lookupMsg.style.display='block';
        }).lookupClienteForIntake(emailEl.value.trim(),dniEl.value.trim());
      }
      function onLookup(r){
        var nb=document.getElementById('nextBtn'); if(nb){nb.disabled=false;nb.textContent='Verificar y continuar →';}
        // El servidor revalida el formato; si algo falla, mostralo junto al campo (seguimos en el paso 1).
        if(!r.dniValid){setErr('dni',r.dniMsg,dniEl);dniEl.focus();return;}
        if(!r.emailValid){setErr('email',r.emailMsg,emailEl);emailEl.focus();return;}
        if(r.blocked){ // V-33 — cliente bloqueado: mensaje genérico, no se avanza del paso 1
          lookupMsg.style.background='#fdecec';lookupMsg.style.color='#a01c1c';
          lookupMsg.textContent=r.blockedMsg||'No es posible procesar solicitudes para este cliente.';
          lookupMsg.style.display='block';
          return;
        }
        if(r.conflict){
          if(r.conflictField==='dni'){setErr('dni','Este DNI ya está registrado con otro correo. Ingresá el correo registrado.',dniEl);dniEl.focus();}
          else{setErr('email','Este correo ya está registrado con otro DNI.',emailEl);emailEl.focus();}
          return;
        }
        clearErrors(['email','dni']);
        var cuilWrap=el_('cuil').closest('div');
        if(r.found){
          modeEl.value='existing';clientIdEl.value=r.client.clientId;
          welcome.className='welcome ret';
          welcome.innerHTML='<b>¡Hola de nuevo, '+esc(r.client.name)+'!</b><br>Ya te tenemos registrado ('+esc(r.client.clientId)+'). Revisá tus datos y completá los pasos.';
          el_('fullName').value=r.client.name||'';
          el_('phone').value=r.client.phone||'';
          el_('cuil').value=r.client.cuil||'';cuilWrap.style.display='';el_('address').value=r.client.address||'';
          applyRefs(r);
        }else{
          modeEl.value='new';clientIdEl.value='';
          welcome.className='welcome newc';
          welcome.innerHTML='<b>Cliente nuevo.</b> Completá todos los datos para procesar tu primera solicitud.';
          el_('fullName').value='';el_('phone').value='';el_('cuil').value='';cuilWrap.style.display='';el_('address').value='';
          applyRefs(r);
        }
        ['fullName','phone','address','cuil','dniPhoto','cuilPhoto'].forEach(function(id){ setRequired(el_(id),true); });
        markReqStars(true);
        renderLimit(r); // V-32 — muestra y aplica el límite del prestatario según su historial
        identified=true;
        showStep(2); // identificación OK → avanzar a "Tus datos"
      }

      // Monto: agrupación de miles + cálculo.
      amountEl.addEventListener('input',function(){
        var d=amountEl.value.replace(/[^0-9]/g,'').replace(/^0+/,'');
        amountEl.value=d?groupNum(d):'';recalc();updateAmountWords();
      });
      function parseAmount(v){var d=String(v==null?'':v).replace(/[^0-9]/g,'');return d?parseInt(d,10):0;}
      // Info fija por plazo (réplica de termInfoForDays_ del servidor).
      function infoForDays(d){ d=parseInt(d,10); return d===15?{days:15,rate:0.25,cuotas:1}:d===30?{days:30,rate:0.5,cuotas:1}:{days:90,rate:1,cuotas:3}; }
      // Plazos permitidos por monto: ≤T1MAX → [15,30] (elige); ≤T2MAX → [30]; grande → [90].
      function allowedDays(a){ if(!(a>0)) return []; if(a>MAXAMT) return null; if(a<=T1MAX) return [15,30]; if(a<=T2MAX) return [30]; return [90]; }
      function recalc(){
        var a=parseAmount(amountEl.value), tv=el_('termView'), tc=el_('termChoice'), sel=el_('term15o30');
        var al=allowedDays(a);
        if(al&&al.length===0){ termEl.value=''; tc.style.display='none'; tv.style.display='block'; calc.style.display='none'; tv.textContent='Ingresá el monto para ver el plazo y las cuotas.'; tv.style.color='#5b6675'; return; }
        if(al===null){ termEl.value=''; tc.style.display='none'; tv.style.display='block'; calc.style.display='none'; tv.textContent=(MAXAMT<LOANMAX?'Tu límite actual es ':'El monto máximo por préstamo es ')+fmt(MAXAMT)+'.'; tv.style.color='#c62828'; return; }
        var days;
        if(al.length>1){ // tramo chico: el cliente elige 15 o 30
          tc.style.display='block'; tv.style.display='none';
          if(al.indexOf(parseInt(sel.value,10))<0) sel.value=String(al[0]);
          days=parseInt(sel.value,10);
        } else { // forzado
          tc.style.display='none'; tv.style.display='block';
          days=al[0];
          tv.textContent='Plazo: '+days+' días ('+Math.round(infoForDays(days).rate*100)+'% de interés)'; tv.style.color='#1c4587';
        }
        termEl.value=String(days);
        var ti=infoForDays(days), total=a*(1+ti.rate), cuota=total/ti.cuotas;
        var sched=ti.cuotas===1 ? ('1 pago de '+fmt(total)+' — '+cuotaDatesTxt(1,ti.days))
          : (ti.cuotas+' cuotas de '+fmt(cuota)+' — '+cuotaDatesTxt(ti.cuotas,90));
        calc.style.display='block';
        calc.innerHTML='Interés: <b>'+fmt(a*ti.rate)+'</b><br>Total a devolver: <b>'+fmt(total)+'</b><br>Repago: <b>'+sched+'</b>';
      }
      // Al cambiar la elección de plazo (tramo chico), recalcular.
      (function(){ var s=document.getElementById('term15o30'); if(s) s.addEventListener('change',recalc); })();

      // V-32 — refleja el límite del prestatario (escalera de graduación) en el formulario:
      // banner con su tope actual + motivo, y atenúa los tramos por encima de su límite.
      function renderLimit(r){
        MAXAMT=(r&&r.maxAmount>0)?r.maxAmount:LOANMAX;
        var banner=document.getElementById('maxBanner');
        if(banner){
          var reason='';
          if(r){
            if(r.hasArrears) reason=' — por atrasos previos, tu límite volvió al inicial. Se recupera saldando a tiempo';
            else if(r.onTimeCount>=2) reason=' — ampliado por tu historial de pagos a tiempo';
            else if(r.onTimeCount>=1) reason=' — ampliado por tu pago a tiempo';
            else reason=' — límite inicial (aumenta al devolver a tiempo)';
          }
          banner.innerHTML='⚠️ <b>Tu límite actual: '+fmt(MAXAMT)+'</b><span style="font-weight:normal">'+esc(reason)+'</span>';
        }
        function dim(id,off){var li=document.getElementById(id); if(!li)return; li.style.opacity=off?'0.4':'';
          var tag=li.querySelector('.tierTag');
          if(off&&!tag){tag=document.createElement('span');tag.className='tierTag';tag.style.cssText='color:#a06000;font-style:italic;font-weight:normal';tag.textContent=' — disponible al ampliar tu límite';li.appendChild(tag);}
          else if(!off&&tag){tag.parentNode.removeChild(tag);}
        }
        dim('tierLi2', MAXAMT<=T1MAX);
        dim('tierLi3', MAXAMT<=T2MAX);
        recalc();
      }

      // Valida todo el formulario y enfoca el primer error.
      function validateAll(){
        // Todos los campos son obligatorios, sea cliente nuevo o existente.
        var reusing=el_('reuseRefs')&&el_('reuseRefs').value==='1';
        var refKeys=reusing?[]:['ref1Name','ref1Rel','ref1Phone','ref2Name','ref2Rel','ref2Phone'];
        var keys=['email','dni','fullName','phone','cuil','address'].concat(refKeys)
          .concat(['amount','term','notes','agreeRepay','agreeMora','agreeConseq','agreeContact','agree','dniPhoto','cuilPhoto']);
        var firstBad=null;
        keys.forEach(function(k){if(!checkField(k)&&!firstBad)firstBad=k;});
        if(firstBad){var t=el_(firstBad)||document.getElementById('e_'+firstBad);
          if(t){if(t.focus)try{t.focus();}catch(_){}if(t.scrollIntoView)t.scrollIntoView({block:'center',behavior:'smooth'});}}
        return !firstBad;
      }

      /* ===== Barra de progreso (muestra la acción y el porcentaje) ===== */
      var pbarFill=document.getElementById('pbarFill'),loadPhase=document.getElementById('loadPhase'),loadPct=document.getElementById('loadPct'),progTimer=null,progVal=0;
      function setProgress(pct,phase){progVal=Math.max(0,Math.min(100,Math.round(pct)));if(pbarFill)pbarFill.style.width=progVal+'%';if(loadPct)loadPct.textContent=progVal+'%';if(phase&&loadPhase)loadPhase.textContent=phase;}
      function startProgress(){var phases=[[8,'Validando los datos…'],[30,'Subiendo las fotos del DNI…'],[58,'Guardando la solicitud…'],[80,'Registrando al cliente…'],[92,'Enviando los correos…']],i=1;setProgress(phases[0][0],phases[0][1]);clearInterval(progTimer);
        progTimer=setInterval(function(){if(i<phases.length){setProgress(phases[i][0],phases[i][1]);i++;}else if(progVal<96){setProgress(progVal+1);}},700);}
      function stopProgress(){clearInterval(progTimer);progTimer=null;}
      function finishProgress(cb){stopProgress();setProgress(100,'¡Listo!');setTimeout(cb,400);}

      function onSuccess(r){ finishProgress(function(){
        loadingView.style.display='none';
        if(!r||!r.ok){successView.innerHTML='<div style="text-align:center"><div class="check">✓</div><h1 style="color:#38761d">¡Solicitud recibida!</h1><p class="sub">'+esc(r&&r.message?r.message:r)+'</p></div>';successView.style.display='block';return;}
        successView.innerHTML=
          '<div style="text-align:center"><div class="check">&#10003;</div>'+
          '<h1 style="color:#38761d">¡Solicitud recibida!</h1>'+
          '<p class="sub">Gracias, <b>'+esc(r.name)+'</b>'+(r.returning?' — nos alegra tenerte de vuelta':'')+'. Tu solicitud quedó registrada y está pendiente de revisión.</p></div>'+
          '<table class="summary">'+
          '<tr><td class="k">Solicitante</td><td>'+esc(r.name)+' ('+esc(r.clientId)+')</td></tr>'+
          '<tr><td class="k">Correo</td><td>'+esc(r.email)+'</td></tr>'+
          '<tr><td class="k">DNI</td><td>'+esc(r.dni)+'</td></tr>'+
          '<tr><td class="k">Teléfono</td><td>'+esc(r.phone)+'</td></tr>'+
          '<tr><td class="k">Monto solicitado</td><td>'+esc(r.amountFmt)+'</td></tr>'+
          '<tr><td class="k">Plazo</td><td>'+esc(r.termLabel)+' ('+esc(r.ratePct)+' de interés)</td></tr>'+
          '<tr><td class="k">Interés</td><td>'+esc(r.interestFmt)+'</td></tr>'+
          '<tr><td class="k">Total a devolver</td><td><b>'+esc(r.totalFmt)+'</b></td></tr>'+
          '</table>'+
          '<div class="note">Enviamos un correo de confirmación a <b>'+esc(r.email)+'</b>. Revisaremos tu solicitud y nos pondremos en contacto. Podés cerrar esta ventana.</div>';
        successView.style.display='block';
      }); }
      function onFail(err){stopProgress();loadingView.style.display='none';formView.style.display='block';btn.disabled=false;btn.textContent='Enviar solicitud';msg.className='err';msg.style.display='block';msg.textContent=(err.message||err);}
      // Fila de la tabla del popup de confirmación.
      function crow(k,v){return '<tr><td style="padding:5px 8px;border:1px solid #eee;background:#f7f9fc;font-weight:bold;width:42%">'+k+'</td><td style="padding:5px 8px;border:1px solid #eee">'+v+'</td></tr>';}
      var confirmOverlay=document.getElementById('confirmOverlay');
      function showConfirm(){
        var a=parseAmount(amountEl.value),t=parseInt(termEl.value,10);
        if(!(a>0)||!t)return false; // recalc setea termEl.value sólo si el monto es válido
        var ti=infoForDays(t),r=ti.rate;
        var interes=a*r,total=a*(1+r),feePct=${lateFeePct},feeDay=total*feePct/100,pct=Math.round(r*100);
        var grace=${moraGrace},capPct=${moraCapPct},feeCap=a*capPct/100;
        function lateTotal(d){return total+Math.min(feeDay*Math.max(0,d-grace),feeCap);} // recargo tras la gracia, con tope
        var d1=grace+1,d2=grace+3,d3=grace+7;
        var cuota=total/ti.cuotas;
        var schedTxt=ti.cuotas===1?('1 pago de <b>'+fmt(total)+'</b> — '+cuotaDatesTxt(1,t)):(ti.cuotas+' cuotas de <b>'+fmt(cuota)+'</b> — '+cuotaDatesTxt(ti.cuotas,90));
        document.getElementById('confirmBody').innerHTML=
          '<p>Estás por solicitar un préstamo de:</p>'+
          '<table style="width:100%;border-collapse:collapse;font-size:13.5px">'+
          crow('Monto','<b>'+fmt(a)+'</b><br><span style="color:#666;font-size:12px;font-style:italic">'+capFirst(numeroALetras(a))+(a===1?' peso':' pesos')+'</span>')+
          crow('Plazo',t+' días')+
          crow('Interés',fmt(interes)+' ('+pct+'%)')+
          crow('Total a devolver','<b>'+fmt(total)+'</b>')+
          crow('Repago',schedTxt)+
          crow('⚠️ Recargo por mora','<span style="color:#7a5200"><b style="color:#b45309">'+feePct+'% por día</b> sobre el total a devolver ('+fmt(feeDay)+' por día) de atraso posterior al vencimiento'+(grace>0?', tras '+grace+' día(s) de gracia':'')+', con tope del '+capPct+'% del capital</span>')+
          crow('Si pagás tarde','<span style="color:#7a5200">'+d1+' días → <b>'+fmt(lateTotal(d1))+'</b><br>'+d2+' días → <b>'+fmt(lateTotal(d2))+'</b><br>'+d3+' días → <b>'+fmt(lateTotal(d3))+'</b></span>')+
          '</table>'+
          '<p style="margin-top:12px">¿Confirmás que querés pedir este préstamo por <b>'+t+' días</b> a <b>'+pct+'%</b> de interés?</p>';
        confirmOverlay.style.display='flex';
        return true;
      }
      document.getElementById('confirmCancel').addEventListener('click',function(){confirmOverlay.style.display='none';});
      document.getElementById('confirmOk').addEventListener('click',function(){
        confirmOverlay.style.display='none';
        msg.style.display='none';btn.disabled=true;btn.textContent='Subiendo…';formView.style.display='none';loadingView.style.display='block';
        startProgress();
        google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFail).submitIntakeSmart(f);
      });
      f.addEventListener('submit',function(e){
        e.preventDefault();
        if(!validateAll()){msg.className='err';msg.style.display='block';msg.textContent='Revisá los campos marcados en rojo antes de enviar.';return;}
        msg.style.display='none';
        showConfirm(); // muestra el popup; el envío real ocurre al tocar "Confirmar y enviar"
      });

      /* ===== Asistente por pasos ===== */
      var STEP_KEYS={1:['email','dni'],2:['fullName','phone','cuil','address'],3:['ref1Name','ref1Rel','ref1Phone','ref2Name','ref2Rel','ref2Phone'],4:['amount','term','notes'],5:['dniPhoto','cuilPhoto'],6:['agreeRepay','agreeMora','agreeConseq','agreeContact','agree']};
      var STEP_LABELS={1:'Identificación',2:'Tus datos',3:'Referencias',4:'El préstamo',5:'Documentación',6:'Confirmación'};
      var MAXSTEP=6, curStep=1;
      function stepValid(n){
        var keys=STEP_KEYS[n]||[];
        if(n===3 && el_('reuseRefs') && el_('reuseRefs').value==='1') keys=[]; // referencias en archivo → nada que validar
        var firstBad=null;
        keys.forEach(function(k){ if(!checkField(k)&&!firstBad) firstBad=k; });
        if(firstBad){var t=el_(firstBad); if(t){ if(t.focus)try{t.focus();}catch(_){}
          if(t.scrollIntoView)t.scrollIntoView({block:'center',behavior:'smooth'});}}
        return !firstBad;
      }
      function showStep(n){
        n=Math.max(1,Math.min(MAXSTEP,n)); curStep=n;
        var steps=document.querySelectorAll('.wstep');
        for(var i=0;i<steps.length;i++){ steps[i].classList.toggle('active', String(steps[i].getAttribute('data-step'))===String(n)); }
        var chips=document.querySelectorAll('#stepper .schip');
        for(var j=0;j<chips.length;j++){ var s=Number(chips[j].getAttribute('data-goto'));
          chips[j].classList.toggle('active', s===n); chips[j].classList.toggle('done', s<n); }
        var sh=document.getElementById('stepHint'); if(sh) sh.textContent='Paso '+n+' de '+MAXSTEP+' — '+STEP_LABELS[n];
        var pf=document.getElementById('progressFill'); if(pf) pf.style.width=Math.round((n-1)/(MAXSTEP-1)*100)+'%';
        var prev=document.getElementById('prevBtn'), next=document.getElementById('nextBtn');
        if(prev) prev.style.display = n>1 ? 'inline-block':'none';
        if(next){ next.style.display = n<MAXSTEP ? 'inline-block':'none'; next.textContent = (n===1 ? 'Verificar y continuar →' : 'Siguiente →'); }
        // El botón "Enviar solicitud" vive dentro del paso 6, así que sólo se ve ahí.
        try{ card.scrollIntoView({block:'start',behavior:'smooth'}); }catch(_){ window.scrollTo(0,0); }
      }
      function goStep(target){
        if(target===curStep) return;
        if(target<=curStep){ showStep(target); return; }          // atrás: siempre permitido
        if(!identified){ showStep(1); return; }                   // sin identificar → sólo el paso 1
        for(var s=curStep;s<target;s++){ if(!stepValid(s)){ showStep(s); return; } } // adelante: valida lo previo
        showStep(target);
      }
      (function(){
        var nb=document.getElementById('nextBtn'), pb=document.getElementById('prevBtn');
        if(nb) nb.addEventListener('click',function(){ if(curStep===1){ runLookup(); } else if(stepValid(curStep)){ showStep(curStep+1); } });
        if(pb) pb.addEventListener('click',function(){ showStep(curStep-1); });
        var chips=document.querySelectorAll('#stepper .schip');
        for(var i=0;i<chips.length;i++){ (function(c){ c.addEventListener('click',function(){ goStep(Number(c.getAttribute('data-goto'))); }); })(chips[i]); }
        showStep(1); // inicia el asistente en "Identificación"
      })();
    </script></body></html>`;
}
