/************************************************************
 * GESTOR DE PRÉSTAMOS v2  —  Google Sheets + Apps Script
 * ----------------------------------------------------------
 * Versión mejorada y en español. Reemplaza a LoanManager.gs
 * (pegue SOLO este archivo para evitar funciones duplicadas).
 *
 * MEJORAS
 *  - Interfaz 100% en español (incluye estados PAGADO/VENCIDO/ACTIVO).
 *  - Panel lateral (sidebar) con botones reales — sin dibujos.
 *  - Confirmación antes de "Configurar" (evita borrar datos).
 *  - Correos automáticos: solicitud recibida, aprobación (con contrato),
 *    recibo de pago, y recordatorios de vencimiento / mora.
 *  - Hoja "Errores" con registro visible de fallos.
 *  - Hoja "Panel" con indicadores del prestamista.
 *  - Ruta de rechazo/archivo para solicitantes.
 *  - Detección de DNI duplicado y validaciones a prueba de errores.
 *  - Formulario web reforzado + página de consulta de saldo.
 *  - Columnas de fórmula protegidas (aviso al editar).
 *
 * PUESTA EN MARCHA
 *  1) Extensions ▸ Apps Script. Pegue TODO este archivo. Guarde.
 *  2) Ejecute  ->  onOpen  (autorice: Sheets, Gmail, Drive, Forms/ScriptApp).
 *  3) Recargue la hoja. Menú "Gestor de Préstamos" ▸ ① Configurar.
 *  4) Complete "Configuración". Menú ⑥ para publicar el formulario web.
 ************************************************************/

const CFG = {
  CURRENCY_FMT: '"$"#,##0.00',
  LOCALE: 'es-AR', CURRENCY_CODE: 'ARS',
  MAX_ROWS: 300,
  MAIN_FOLDER: 'Documentos de Prestatarios',
  SHEETS: {
    BORROWERS: 'Prestatarios', PAYMENTS: 'Pagos', SUMMARY: 'Resumen', SETTINGS: 'Configuración',
    AGREEMENT: 'Estudio de Contratos', STATEMENTS: 'Estudio de Estados', REMINDER: 'Estudio de Recordatorios', NEW: 'Nuevos Prestatarios',
    PANEL: 'Panel', STATS: 'Estadísticas', LATE: 'Pagos Atrasados', CLEARED: 'Saldados', REJECTED: 'Rechazados', ERRORS: 'Errores', HELP: 'Instrucciones',
    SIGN: 'Firmas', CLIENTS: 'Clientes', INSTALLMENTS: 'Cuotas',
  },
};
const ST = { ACTIVE: 'ACTIVO', OVERDUE: 'VENCIDO', PAID: 'PAGADO', CLEARED: 'SALDADO' };
// Estados de cuota (hoja "Cuotas"): PAGADA / VENCIDA / PENDIENTE.
const CST = { PAID: 'PAGADA', OVERDUE: 'VENCIDA', PENDING: 'PENDIENTE' };
// Estado de firma del contrato (columna P de "Prestatarios", modelo normalizado).
const SIGN = { PENDING: 'PENDIENTE', SIGNED: 'FIRMADO' };

// ============================================================================
// MODELO NORMALIZADO (3NF)
// La identidad del cliente (Nombre/Correo/DNI/Teléfono) vive SOLO en la hoja
// "Clientes". "Prestatarios" referencia al cliente por "ID Cliente" (columna B).
// PB = mapa único de columnas de "Prestatarios" — fuente de verdad para evitar
// índices mágicos dispersos por el código.
// ============================================================================
const PB = {
  LOAN_ID: 1, CLIENT_ID: 2, NAME: 3, DNI: 4, PRINCIPAL: 5, TERM: 6, RATE: 7, LOAN_DATE: 8, DUE: 9,
  INTEREST: 10, TOTAL: 11, PAID: 12, BALANCE: 13, STATE: 14, PDF: 15, SENT: 16, NOTICE: 17,
  SIGN_STATUS: 18, SIGN_DATE: 19, SIGN_PDF: 20, DELETE: 21, SEND_DISB: 22, LOANCOUNT: 23,
  // Columnas de mora: recargo acumulado (derivado de "Pagos Atrasados"), ajuste manual
  // (condonar/cambiar por préstamo) y saldo con mora (= monto a pagar hoy).
  MORA_ACUM: 24, MORA_ADJ: 25, BALANCE_MORA: 26,
};
// Nombre/DNI (col C/D) se muestran en "Prestatarios" (se resuelven de "Clientes"
// por ID Cliente). La fuente de verdad de la identidad sigue siendo "Clientes".
/** Número de columna → letra A1 (1→A, 27→AA). Para construir fórmulas desde PB. */
function colL_(n) { let s = ''; n = Number(n); while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
// Columnas de firma en "Prestatarios" (derivadas de PB).
const COL_SIGN_STATUS = PB.SIGN_STATUS, COL_SIGN_DATE = PB.SIGN_DATE, COL_SIGN_PDF = PB.SIGN_PDF;

// Mapa de columnas de "Clientes".
const PC = { CLIENT_ID: 1, NAME: 2, EMAIL: 3, DNI: 4, PHONE: 5, DNI_NORM: 6, EMAIL_NORM: 7, CREATED: 8, UPDATED: 9 };
const CLIENTES_HEADERS = ['ID Cliente', 'Nombre', 'Correo', 'DNI', 'Teléfono', 'DNI (norm)', 'Correo (norm)', 'Fecha alta', 'Actualizado'];
// Mapa de columnas de "Pagos" (modelo normalizado, sin identidad denormalizada).
const PP = { PAY_ID: 1, LOAN_ID: 2, DATE: 3, AMOUNT: 4, BALANCE: 5, RECEIPT: 6, NAME: 7, DNI: 8, SEND_RECEIPT: 9, GEN_RECEIPT: 10 };
// Hojas de operaciones diarias: siempre visibles. El resto son de configuración/consulta
// y se pueden ocultar/mostrar desde el menú ▸ 🗂 Pestañas.
const MAIN_SHEETS = [
  CFG.SHEETS.BORROWERS,   // Prestatarios — cartera
  CFG.SHEETS.NEW,         // Nuevos Prestatarios — solicitudes a revisar
  CFG.SHEETS.PAYMENTS,    // Pagos — registrar cobros
  CFG.SHEETS.INSTALLMENTS,// Cuotas — cronograma de pagos
  CFG.SHEETS.LATE,        // Pagos Atrasados — mora / avisos
  CFG.SHEETS.PANEL,       // Panel — indicadores
];
// Encabezados de "Cuotas" (cronograma; una fila por cuota de cada préstamo).
const CUOTAS_HEADERS = ['ID Cuota', 'ID Préstamo', 'Nº Cuota', 'Fecha de Vencimiento',
  'Monto Cuota', 'Monto Pagado', 'Saldo Cuota', 'Estado', 'Último Aviso'];
const CU = { CUOTA_ID: 1, LOAN_ID: 2, NUM: 3, DUE: 4, AMOUNT: 5, PAID: 6, BALANCE: 7, STATE: 8, NOTICE: 9 };
// Encabezados de "Nuevos Prestatarios" (el formulario agrega filas aquí, en este orden).
const NB = ['Fecha de Envío', 'Nombre completo', 'Correo', 'DNI', 'Teléfono', 'Monto Solicitado',
  'Plazo (meses)', 'Notas / Motivo', 'Foto del frente del DNI', 'Foto del dorso del DNI', 'Verificado?', 'Rechazar?', 'Resultado',
  'Verificar BCRA?', 'Parámetro BCRA', 'CUIL', 'Nombre BCRA', 'Peor Situación', '% al día', 'Decisión', 'Puntaje BCRA', 'Calificación de Riesgo', 'Resumen BCRA', 'Dirección', 'Forma de Pago',
  'Anular límites',
  'Ref 1 Nombre', 'Ref 1 Vínculo', 'Ref 1 Teléfono', 'Ref 1 Validada?',
  'Ref 2 Nombre', 'Ref 2 Vínculo', 'Ref 2 Teléfono', 'Ref 2 Validada?', 'Compromiso aceptado'];

/* ============================ MENÚ ============================ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gestor de Préstamos')
    .addItem('▶ Abrir panel del prestamista', 'openSidebar')
    .addSeparator()
    .addItem('🆕 Aplicar novedades (columnas y ajustes)', 'applyFeatureUpdates')
    .addItem('① Configurar / reconstruir hojas', 'setupConfirm')
    .addItem('⑥ Formulario web — publicar / ver enlace', 'showWebFormLink')
    .addItem('Configurar marca (nombre + logo)', 'setBranding')
    .addItem('④ Estadísticas y gráficos', 'showStats')
    .addItem('💵 Retiro disponible (sin frenar crecimiento)', 'showWithdrawable')
    .addItem('⑤ Actualizar saldos y resumen', 'refreshAll')
    .addItem('➕ Crear hoja de Recordatorios de pago', 'createRemindersSheet_')
    .addItem('📄 Crear hoja de Estudio de Contratos', 'createAgreementSheet_')
    .addItem('✉ Reenviar enlace de firma (fila seleccionada)', 'resendSigningLink_')
    .addSeparator()
    .addItem('Registrar un pago', 'openSidebar')
    .addItem('✔ Mover préstamos saldados → hoja Saldados', 'moveClearedBorrowersConfirm')
    .addItem('🔄 Reestructurar préstamo vencido (congelar mora + plan de cuotas)', 'restructurarPrestamoUI')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('💸 Mora (fila activa)')
      .addItem('Condonar mora (poner en 0)', 'condonarMoraFilaActiva')
      .addItem('Fijar un importe de mora…', 'fijarMoraFilaActiva')
      .addItem('Restaurar mora automática', 'restaurarMoraFilaActiva'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🚫 Bloqueo de clientes (fila activa en Clientes)')
      .addItem('Bloquear cliente…', 'bloquearClienteFilaActiva')
      .addItem('Desbloquear cliente', 'desbloquearClienteFilaActiva'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🗂 Pestañas')
      .addItem('Ver solo operaciones diarias', 'showDailyOnly')
      .addItem('Mostrar todas las hojas', 'showAllSheets'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧪 Datos de prueba')
      .addItem('Crear 40 prestatarios de prueba', 'seedTestBorrowersConfirm')
      .addItem('Eliminar prestatarios de prueba', 'removeTestBorrowersConfirm')
      .addSeparator()
      .addItem('🧪 Ejecutar pruebas (suite)', 'runAllTests'))
    .addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutput(sidebarHtml_()).setTitle('Panel del Prestamista');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ===================== REGISTRO DE ERRORES ===================== */
function logError_(where, err) {
  try {
    const sh = getSS_().getSheetByName(CFG.SHEETS.ERRORS) || getSS_().insertSheet(CFG.SHEETS.ERRORS);
    if (sh.getLastRow() === 0)
      sh.getRange(1, 1, 1, 3).setValues([['Fecha', 'Función', 'Error']]).setFontWeight('bold');
    sh.appendRow([new Date(), where, (err && err.message) ? err.message : String(err)]);
  } catch (e) { /* nunca fallar por el log */ }
}
function guard_(where, fn) { try { return fn(); } catch (err) { logError_(where, err); throw err; } }

/* ===================== MARCA (nombre + logo) ===================== */
function setBranding() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('Marca', 'Nombre de la empresa:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  setSetting_('Nombre del Prestamista', n.getResponseText().trim() || 'Impulso Crédito');
  const l = ui.prompt('Marca', 'Enlace de Drive del logo (pegue el enlace para compartir; opcional):', ui.ButtonSet.OK_CANCEL);
  if (l.getSelectedButton() === ui.Button.OK) setSetting_('Logo (enlace de Drive)', l.getResponseText().trim());
  _logoBlob = null; _logoLoaded = false; _logoUri = null;
  ui.alert('Marca actualizada. El logo aparecerá en correos, contratos y páginas web.');
}
function setSetting_(key, value) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  const data = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  sh.appendRow([key, value]);
}

var _logoBlob = null, _logoLoaded = false, _logoUri = null;
function getLogoBlob_() {
  if (_logoLoaded) return _logoBlob;
  _logoLoaded = true; _logoBlob = null;
  const ref = String(getSetting_('Logo (enlace de Drive)') || '').trim();
  const ids = extractDriveIds_(ref);
  if (ids.length) { try { const b = DriveApp.getFileById(ids[0]).getBlob(); b.setName('logo'); _logoBlob = b; } catch (e) { logError_('getLogoBlob_', e); } }
  return _logoBlob;
}
function logoDataUri_() {
  if (_logoUri !== null) return _logoUri;
  const b = getLogoBlob_();
  _logoUri = b ? ('data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes())) : '';
  return _logoUri;
}
/**
 * Firma del prestamista (dueño de Impulso Crédito) como data URI, para estampar
 * en TODOS los contratos. Prioridad: 1) "Firma del Prestamista (enlace de Drive)"
 * en Configuración; 2) firma embebida por defecto (OWNER_SIGNATURE_DATAURI,
 * en el archivo SignatureData.gs). Devuelve '' si no hay ninguna.
 */
var _ownerSigUri = null;
function ownerSignatureDataUri_() {
  if (_ownerSigUri !== null) return _ownerSigUri;
  _ownerSigUri = '';
  const ref = String(getSetting_('Firma del Prestamista (enlace de Drive)') || '').trim();
  const ids = extractDriveIds_(ref);
  if (ids.length) {
    try {
      const b = DriveApp.getFileById(ids[0]).getBlob();
      _ownerSigUri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
      return _ownerSigUri;
    } catch (e) { logError_('ownerSignatureDataUri_', e); }
  }
  if (typeof OWNER_SIGNATURE_DATAURI === 'string' && OWNER_SIGNATURE_DATAURI) _ownerSigUri = OWNER_SIGNATURE_DATAURI;
  return _ownerSigUri;
}
/**
 * Firma del segundo titular de Impulso Crédito (Xoana Elizabeth Beron) como data URI.
 * Se estampa, junto con la del titular principal, en TODOS los contratos.
 * Prioridad: 1) "Firma del Prestamista 2 (enlace de Drive)" en Configuración;
 * 2) COOWNER_SIGNATURE_DATAURI (en SignatureData.gs). Devuelve '' si no hay ninguna.
 */
var _coOwnerSigUri = null;
function coOwnerSignatureDataUri_() {
  if (_coOwnerSigUri !== null) return _coOwnerSigUri;
  _coOwnerSigUri = '';
  const ref = String(getSetting_('Firma del Prestamista 2 (enlace de Drive)') || '').trim();
  const ids = extractDriveIds_(ref);
  if (ids.length) {
    try {
      const b = DriveApp.getFileById(ids[0]).getBlob();
      _coOwnerSigUri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
      return _coOwnerSigUri;
    } catch (e) { logError_('coOwnerSignatureDataUri_', e); }
  }
  if (typeof COOWNER_SIGNATURE_DATAURI === 'string' && COOWNER_SIGNATURE_DATAURI) _coOwnerSigUri = COOWNER_SIGNATURE_DATAURI;
  return _coOwnerSigUri;
}
/** Nombres completos de los dos firmantes titulares (configurables). */
function ownerSignatories_() {
  return [
    { name: String(getSetting_('Firmante 1 (nombre completo)') || 'Denis Delroy Bell').trim(), cuit: String(getSetting_('CUIT/CUIL Firmante 1') || '').trim(), sig: ownerSignatureDataUri_() },
    { name: String(getSetting_('Firmante 2 (nombre completo)') || 'Xoana Elizabeth Beron').trim(), cuit: String(getSetting_('CUIT/CUIL Firmante 2') || '').trim(), sig: coOwnerSignatureDataUri_() },
  ];
}
/** Datos de la cuenta de cobro/desembolso del Prestamista (Configuración). */
function lenderAccount_() {
  return {
    cbu: String(getSetting_('Cuenta de cobro/desembolso — CBU/CVU') || '').trim(),
    alias: String(getSetting_('Cuenta de cobro/desembolso — Alias') || '').trim(),
    institucion: String(getSetting_('Cuenta de cobro/desembolso — Institución') || '').trim(),
    titular: String(getSetting_('Cuenta de cobro/desembolso — Titular') || '').trim(),
    mpAlias: String(getSetting_('Mercado Pago — Alias/CVU') || '').trim(),
    mpTitular: String(getSetting_('Mercado Pago — Titular') || '').trim(),
  };
}
function companyName_() {
  const v = String(getSetting_('Nombre del Prestamista') || '').trim();
  return (!v || v === 'Su Nombre / Empresa' || v === 'Su Nombre / Company') ? 'Impulso Crédito' : v;
}
/** Recargo por mora como fracción (5% => 0.05). Predeterminado 5% si no está configurado. */
function lateFeeRate_() {
  const raw = String(getSetting_('Recargo por Mora diario (%)') || getSetting_('Recargo por Mora (%)') || '').trim();
  const v = raw === '' ? 5 : Number(raw.replace(',', '.'));
  return (isFinite(v) && v > 0) ? v / 100 : 0;
}
function lateFeePctText_() { return round2_(lateFeeRate_() * 100) + '%'; }

/** Días de gracia antes de que empiece a correr la mora (Configuración). Predet. 3. */
function moraGraceDays_() {
  const raw = String(getSetting_('Días de gracia antes de mora') || '').replace(/[^0-9.\-]/g, '').trim();
  const n = raw === '' ? 3 : Number(raw);
  return (isFinite(n) && n >= 0) ? Math.floor(n) : 3;
}
/** Tope acumulado de mora como FRACCIÓN del capital (Configuración "Tope de mora (% del capital)").
 *  Predeterminado 100% (comportamiento contractual previo) si el ajuste falta. */
function moraCapFrac_() {
  const raw = String(getSetting_('Tope de mora (% del capital)') || '').replace(/[^0-9.\-]/g, '').trim();
  const n = raw === '' ? 100 : Number(raw);
  return (isFinite(n) && n > 0) ? n / 100 : 1;
}
/** Texto del tope de mora como % del capital (p. ej. "50%"). */
function moraCapPctText_() { return round2_(moraCapFrac_() * 100) + '%'; }

// Colores de marca Impulso Crédito
var BRAND = { navy: '#1c4587', amber: '#e8a33d', tag: 'Préstamos rápidos y confiables' };
function brandInitials_() {
  const ini = companyName_().trim().split(/\s+/).filter(Boolean).map(function (w) { return w.charAt(0); }).join('').slice(0, 2).toUpperCase();
  return ini || 'IC';
}
function brandWordmark_(size) {
  const parts = companyName_().trim().split(/\s+/);
  const first = esc_(parts.shift() || 'Impulso'), rest = esc_(parts.join(' '));
  return '<span style="font-size:' + size + 'px;font-weight:bold;color:' + BRAND.navy + '">' + first + '</span>' +
    (rest ? '<span style="font-size:' + size + 'px;font-weight:bold;color:' + BRAND.amber + '"> ' + rest + '</span>' : '');
}
function brandBadge_(px, cid) {
  if (cid && getLogoBlob_()) return '<img src="cid:logo" style="max-height:' + px + 'px;max-width:' + (px + 8) + 'px">';
  if (!cid) { const u = logoDataUri_(); if (u) return '<img src="' + u + '" style="max-height:' + px + 'px">'; }
  const s = Math.round(px * 0.92);
  return '<span style="display:inline-block;width:' + s + 'px;height:' + s + 'px;line-height:' + s +
    'px;text-align:center;background:' + BRAND.navy + ';color:#fff;border-radius:' + Math.round(s / 4.5) +
    'px;font-size:' + Math.round(s / 2.2) + 'px;font-weight:bold;font-family:Arial,Helvetica,sans-serif">' + esc_(brandInitials_()) + '</span>';
}

/** Cabecera de marca para correos (logo en línea vía cid, o monograma CSS). */
function brandHeaderEmail_() {
  return '<table style="border-collapse:collapse;width:100%;border-bottom:3px solid ' + BRAND.navy + ';margin-bottom:14px"><tr>' +
    '<td style="width:62px;padding:0 10px 8px 0;vertical-align:middle">' + brandBadge_(50, true) + '</td>' +
    '<td style="padding-bottom:8px;vertical-align:middle">' + brandWordmark_(21) +
    '<div style="font-size:11px;color:#888">' + esc_(BRAND.tag) + '</div></td></tr></table>';
}
function brandFooterEmail_() {
  const c = esc_(companyName_()), phone = esc_(getSetting_('Teléfono del Prestamista')),
    email = esc_(getSetting_('Correo del Prestamista')), addr = esc_(getSetting_('Dirección del Prestamista'));
  const bits = [c, addr, phone, email].filter(function (x) { return x; }).join(' · ');
  return '<p style="margin-top:22px;font-size:12px;color:#888;border-top:1px solid #ddd;padding-top:8px">' + bits + '</p>';
}
/** Envía un correo con la cabecera/pie de marca y el logo en línea. opts.attachments opcional. */
function sendBrandedEmail_(to, subject, plainBody, htmlInner, opts) {
  opts = opts || {};
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:640px">' +
    brandHeaderEmail_() + htmlInner + brandFooterEmail_() + '</div>';
  // Modo prueba: captura el correo en vez de enviarlo (sin efectos reales).
  if (TEST_MODE) { _emailOutbox.push({ to: to, subject: subject, plainBody: plainBody, htmlInner: htmlInner, html: html, opts: opts }); return; }
  const o = { name: companyName_(), htmlBody: html };
  const logo = getLogoBlob_(); if (logo) o.inlineImages = { logo: logo };
  if (opts.attachments) o.attachments = opts.attachments;
  GmailApp.sendEmail(to, subject, plainBody, o);
}
/** Cabecera de marca para PDF y páginas web (data URI, o monograma CSS). */
function brandHeaderHtml_() {
  return '<div style="text-align:center;margin-bottom:12px">' + brandBadge_(64, false) +
    '<div style="margin-top:6px">' + brandWordmark_(24) + '</div>' +
    '<div style="font-size:12px;color:#888">' + esc_(BRAND.tag) + '</div></div>';
}

/* ============================ SETUP ============================ */
function setupConfirm() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Configurar', 'Esto RECONSTRUYE las hojas Prestatarios y Pagos y BORRA sus datos. ' +
    '¿Desea continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) setup();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  setupSettings_(ss); setupClientes_(ss); setupBorrowers_(ss); setupPayments_(ss); setupCuotas_(ss); setupSummary_(ss);
  setupAgreement_(ss); setupStatements_(ss); setupReminders_(ss); setupNew_(ss); setupPanel_(ss); setupStats_(ss); setupLate_(ss); setupCleared_(ss); setupRejected_(ss);
  setupErrors_(ss); setupHelp_(ss); setupFirmas_(ss);
  rebuildLateSheet_(ss);
  ensureTriggers_();
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (s1 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch (e) {} }
  applyDailyVisibility_(ss);
  SpreadsheetApp.getUi().alert('Listo. Se dejaron visibles las pestañas de operaciones diarias; ' +
    'las demás (Configuración, Estadísticas, etc.) están ocultas — use el menú ▸ 🗂 Pestañas para mostrarlas. ' +
    'Complete "Configuración" y publique el formulario con el menú ⑥.');
}

/**
 * Aplica las novedades de esta versión SIN borrar datos: agrega los ajustes nuevos
 * a "Configuración", amplía el acta de "Firmas" (IP/hash/CUIL/domicilio) y crea
 * la columna "Forma de Pago" en "Nuevos Prestatarios" y "Clientes" si faltan.
 */
function applyFeatureUpdates() {
  return guard_('applyFeatureUpdates', function () {
    const ss = getSS_(), ui = SpreadsheetApp.getUi(), done = [];
    // 1) Ajustes nuevos en Configuración (no destructivo: solo agrega los que faltan).
    const nAdded = ensureSettingsKeys_(ss, [
      'CUIT/CUIL Firmante 1', 'CUIT/CUIL Firmante 2',
      'Cuenta de cobro/desembolso — CBU/CVU', 'Cuenta de cobro/desembolso — Alias',
      'Cuenta de cobro/desembolso — Institución', 'Cuenta de cobro/desembolso — Titular',
      'Mercado Pago — Alias/CVU', 'Mercado Pago — Titular',
      'Tope de concentración por prestatario (% del capital disponible)',
      'Horizonte de protección del crecimiento (días)',
      'Ventana para estimar el ritmo de colocación (días)',
      'Monto máximo por préstamo', 'Tope tramo 25% (15 días)', 'Tope tramo 50% (30 días)',
      // V-32 escalera de graduación + recuperación (mora): claves nuevas.
      'Límite inicial (préstamo nuevo)', 'Límite tras 1 préstamo saldado', 'Límite tras 2 préstamos saldados',
      'Días de gracia antes de mora', 'Tope de mora (% del capital)',
    ]);
    done.push(nAdded ? (nAdded + ' ajuste(s) nuevo(s) en «Configuración»') : '«Configuración» ya estaba al día');
    // Prefill del tope de concentración (10%) si quedó en blanco, para que sea visible/editable.
    try {
      const stsh = ss.getSheetByName(CFG.SHEETS.SETTINGS);
      if (stsh && String(getSetting_('Tope de concentración por prestatario (% del capital disponible)') || '').trim() === '')
        setSettingValue_(stsh, 'Tope de concentración por prestatario (% del capital disponible)', '10');
      // Prefill de los parámetros de retiro (60 días) si quedaron en blanco.
      if (stsh && String(getSetting_('Horizonte de protección del crecimiento (días)') || '').trim() === '')
        setSettingValue_(stsh, 'Horizonte de protección del crecimiento (días)', '60');
      if (stsh && String(getSetting_('Ventana para estimar el ritmo de colocación (días)') || '').trim() === '')
        setSettingValue_(stsh, 'Ventana para estimar el ritmo de colocación (días)', '60');
      // Prefill de los tramos por monto (tope $500.000; cortes $150.000 / $300.000) si quedaron en blanco.
      if (stsh && String(getSetting_('Monto máximo por préstamo') || '').trim() === '')
        setSettingValue_(stsh, 'Monto máximo por préstamo', '500000');
      if (stsh && String(getSetting_('Tope tramo 25% (15 días)') || '').trim() === '')
        setSettingValue_(stsh, 'Tope tramo 25% (15 días)', '150000');
      if (stsh && String(getSetting_('Tope tramo 50% (30 días)') || '').trim() === '')
        setSettingValue_(stsh, 'Tope tramo 50% (30 días)', '300000');
      // Escalera de graduación (V-32): límite inicial / tras 1 / tras 2 préstamos saldados.
      if (stsh && String(getSetting_('Límite inicial (préstamo nuevo)') || '').trim() === '')
        setSettingValue_(stsh, 'Límite inicial (préstamo nuevo)', '150000');
      if (stsh && String(getSetting_('Límite tras 1 préstamo saldado') || '').trim() === '')
        setSettingValue_(stsh, 'Límite tras 1 préstamo saldado', '300000');
      if (stsh && String(getSetting_('Límite tras 2 préstamos saldados') || '').trim() === '')
        setSettingValue_(stsh, 'Límite tras 2 préstamos saldados', '500000');
      // Recuperación (mora más gradual): 3 días de gracia + tope de mora al 50% del capital.
      if (stsh && String(getSetting_('Días de gracia antes de mora') || '').trim() === '')
        setSettingValue_(stsh, 'Días de gracia antes de mora', '3');
      if (stsh && String(getSetting_('Tope de mora (% del capital)') || '').trim() === '')
        setSettingValue_(stsh, 'Tope de mora (% del capital)', '50');
    } catch (e) { logError_('applyFeatureUpdates:concentracionDefault', e); }
    // 2) Acta de firma ampliada (agrega encabezados IP/hash/CUIL/domicilio; no borra filas).
    setupFirmas_(ss); done.push('columnas de auditoría en «Firmas»');
    // 3) Columna "Forma de Pago" en "Nuevos Prestatarios" y "Clientes" (por nombre de encabezado).
    const nb = ss.getSheetByName(CFG.SHEETS.NEW);
    if (nb) { clientesHeaderCol_(nb, 'Forma de Pago', true); done.push('columna «Forma de Pago» en «Nuevos Prestatarios»'); }
    const cs = ss.getSheetByName(CFG.SHEETS.CLIENTS);
    if (cs) {
      clientesHeaderCol_(cs, 'CUIL', true);
      clientesHeaderCol_(cs, 'Dirección', true);
      clientesHeaderCol_(cs, CLIENTE_PAY_HEADER, true);
      done.push('columnas «CUIL / Dirección / Forma de Pago» en «Clientes»');
    }
    invalidateClientesCache_();
    // 4) Columnas/casillas nuevas (recibos + topes de crédito) — por nombre de encabezado, no destructivo.
    try {
      if (typeof installNewFeatureColumns_ === 'function') {
        const hojas = installNewFeatureColumns_(ss);
        done.push('casillas de recibo + conteo/banner de crédito + «Anular límites» (' + hojas + ')');
      }
    } catch (e) { logError_('applyFeatureUpdates:nuevasColumnas', e); }
    // 5) Referencias + compromiso de pago del formulario (columnas y casillas) — no destructivo.
    try { ensureIntakeRefColumns_(ss).forEach(m => done.push(m)); } catch (e) { logError_('applyFeatureUpdates:referencias', e); }
    // 6) Cronograma de cuotas: crea la hoja "Cuotas" si falta y backfillea préstamos vigentes
    //    (1 cuota a su vencimiento actual; NO cambia plazos ni montos existentes). Aditivo.
    try {
      setupCuotas_(ss);
      const nBf = backfillCuotas_(ss);
      done.push('hoja «Cuotas» lista' + (nBf ? (' · ' + nBf + ' préstamo(s) vigente(s) con cronograma') : ''));
    } catch (e) { logError_('applyFeatureUpdates:cuotas', e); }
    // 7) Pagos: columna "Generar recibo" (casilla) + "Saldo Posterior" como fórmula viva. Aditivo.
    try {
      const psh = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
      if (psh) {
        const genCol = clientesHeaderCol_(psh, 'Generar recibo', true);
        if (genCol > 0) { psh.getRange(2, genCol, CFG.MAX_ROWS, 1).insertCheckboxes(); psh.setColumnWidth(genCol, 110); }
        const balF = []; for (let r = 2; r <= CFG.MAX_ROWS + 1; r++) balF.push([paymentBalanceFormula_(r)]);
        psh.getRange(2, PP.BALANCE, balF.length, 1).setFormulas(balF);
        done.push('casilla «Generar recibo» + «Saldo Posterior» vivo en «Pagos»');
      }
    } catch (e) { logError_('applyFeatureUpdates:pagos', e); }
    // El onEdit (casillas que envían correos) requiere el disparador instalable.
    try { if (typeof ensureTriggers_ === 'function') ensureTriggers_(); } catch (e) { logError_('applyFeatureUpdates:triggers', e); }
    ui.alert('Novedades aplicadas',
      'Se aplicaron sin borrar datos:\n\n• ' + done.join('\n• ') +
      '\n\nComplete los valores de las cuentas del Prestamista y los CUIT en «Configuración», ' +
      'y vuelva a publicar la app web (Implementar ▸ Gestionar implementaciones ▸ Nueva versión) ' +
      'para que el formulario y la página de firma tomen los cambios.', ui.ButtonSet.OK);
  });
}
/**
 * Agrega — sin borrar datos — las columnas de Referencias / Compromiso y las casillas
 * "Ref X Validada?" a "Nuevos Prestatarios", y las columnas de Referencias a "Clientes".
 * Sólo APÉNDA los encabezados que falten (por nombre, al final); nunca reescribe la fila 1
 * ni toca las filas de datos, así toda la información existente queda intacta. Idempotente.
 */
function ensureIntakeRefColumns_(ss) {
  ss = ss || getSS_();
  const done = [];
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const nb = ss.getSheetByName(CFG.SHEETS.NEW);
  if (nb) {
    ['Ref 1 Nombre', 'Ref 1 Vínculo', 'Ref 1 Teléfono', 'Ref 1 Validada?',
      'Ref 2 Nombre', 'Ref 2 Vínculo', 'Ref 2 Teléfono', 'Ref 2 Validada?', 'Compromiso aceptado']
      .forEach(h => clientesHeaderCol_(nb, h, true)); // crea al final sólo si falta
    ['Ref 1 Validada?', 'Ref 2 Validada?'].forEach(h => {
      const c = clientesHeaderCol_(nb, h, true);
      if (c > 0) { nb.getRange(2, c, CFG.MAX_ROWS, 1).setDataValidation(cb); nb.setColumnWidth(c, 100); nb.getRange(1, c).setNote('Tildá cuando hayas contactado y verificado esta referencia.'); }
    });
    // Uniforma el color de los encabezados nuevos con el resto de la hoja (no toca datos).
    nb.getRange(1, 1, 1, nb.getLastColumn()).setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
    done.push('referencias + compromiso + casillas de validación en «Nuevos Prestatarios»');
  }
  const cs = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  if (cs) {
    ['Ref 1 Nombre', 'Ref 1 Vínculo', 'Ref 1 Teléfono', 'Ref 2 Nombre', 'Ref 2 Vínculo', 'Ref 2 Teléfono']
      .forEach(h => clientesHeaderCol_(cs, h, true));
    done.push('referencias en «Clientes»');
  }
  invalidateClientesCache_();
  return done;
}

/** Agrega a "Configuración" las claves que falten (col A), con valor vacío. No borra ni pisa las existentes. */
function ensureSettingsKeys_(ss, keys) {
  const sh = ss.getSheetByName(CFG.SHEETS.SETTINGS) || getOrCreate_(ss, CFG.SHEETS.SETTINGS);
  const last = sh.getLastRow();
  const existing = last >= 1 ? sh.getRange(1, 1, last, 1).getValues().map(r => String(r[0]).trim()) : [];
  let row = last, added = 0;
  keys.forEach(k => {
    if (existing.indexOf(k) === -1) { row++; sh.getRange(row, 1).setValue(k).setFontWeight('bold'); sh.getRange(row, 2).setValue(''); added++; }
  });
  return added;
}

function ensureTriggers_() {
  const fns = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (fns.indexOf('onEditInstallable') === -1)
    ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  if (fns.indexOf('dailyTasks') === -1)
    ScriptApp.newTrigger('dailyTasks').timeBased().everyDays(1).atHour(7).create();
}

function setupSettings_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SETTINGS); sh.clear();
  const rows = [
    ['AJUSTE', 'VALOR'],
    ['Nombre del Prestamista', 'Impulso Crédito'],
    ['Logo (enlace de Drive)', ''],
    ['Firmante 1 (nombre completo)', 'Denis Delroy Bell'],
    ['CUIT/CUIL Firmante 1', ''],
    ['Firma del Prestamista (enlace de Drive)', ''],
    ['Firmante 2 (nombre completo)', 'Xoana Elizabeth Beron'],
    ['CUIT/CUIL Firmante 2', ''],
    ['Firma del Prestamista 2 (enlace de Drive)', ''],
    ['Cuenta de cobro/desembolso — CBU/CVU', ''],
    ['Cuenta de cobro/desembolso — Alias', ''],
    ['Cuenta de cobro/desembolso — Institución', ''],
    ['Cuenta de cobro/desembolso — Titular', ''],
    ['Mercado Pago — Alias/CVU', ''],
    ['Mercado Pago — Titular', ''],
    ['URL de la app web (enlaces a clientes)', ''],
    ['Correo del Prestamista', (Session.getActiveUser().getEmail() || '')],
    ['Correos de aviso adicionales (separados por coma)', ''],
    ['Teléfono del Prestamista', ''],
    ['Dirección del Prestamista', ''],
    ['Jurisdicción', 'Buenos Aires, Argentina'],
    ['Cláusula de Mora', 'En caso de mora se aplica un recargo diario sobre el total a devolver por cada día de atraso posterior al vencimiento, tras un período de gracia, con un tope acumulado como porcentaje del capital. Los valores vigentes (recargo diario, días de gracia y tope) se configuran en esta hoja.'],
    ['Recargo por Mora diario (%)', '5'],
    ['Días de gracia antes de mora', '3'],
    ['Tope de mora (% del capital)', '50'],
    ['Límite inicial (préstamo nuevo)', '150000'],
    ['Límite tras 1 préstamo saldado', '300000'],
    ['Límite tras 2 préstamos saldados', '500000'],
    ['Pie del Contrato', 'Este acuerdo es legalmente vinculante desde la firma de ambas partes.'],
    ['Días de aviso antes del vencimiento', '3'],
    ['Días de aviso de recordatorio', '7'],
    ['Fondo total para prestar', '0'],
    ['Tope de concentración por prestatario (% del capital disponible)', '10'],
    ['Horizonte de protección del crecimiento (días)', '60'],
    ['Ventana para estimar el ritmo de colocación (días)', '60'],
    ['Aceptar solicitudes de préstamo (SÍ/NO)', 'SÍ'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff');
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 500);
  sh.getRange(1, 1, rows.length, 2).setWrap(true); sh.setFrozenRows(1);
}

function setupBorrowers_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.BORROWERS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const H = ['ID Préstamo', 'ID Cliente', 'Nombre', 'DNI', 'Capital', 'Plazo (meses)', 'Tasa',
    'Fecha del Préstamo', 'Fecha de Vencimiento', 'Interés', 'Total a Pagar',
    'Total Pagado', 'Saldo Pendiente (hoy)', 'Estado', 'Contrato PDF', 'Contrato Enviado', 'Último Aviso',
    'Estado de Firma', 'Fecha de Firma', 'Contrato Firmado (PDF)', 'Eliminar', 'Enviar recibo desembolso', 'Préstamos (de ' + MAX_LOANS_PER_CLIENT + ')',
    'Recargo por Mora (acum.)', 'Mora (ajuste)', 'Saldo con Mora'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
  sh.getRange(1, PB.DELETE).setBackground('#990000').setNote('Tildá para ELIMINAR ese préstamo (borra la fila; recuperable con el historial de versiones).');
  sh.getRange(1, PB.SEND_DISB).setNote('Tildá para enviar el recibo de DESEMBOLSO (entrega de fondos) al prestatario. Sólo se envía si el contrato está FIRMADO. Se destilda solo.');
  sh.getRange(1, PB.LOANCOUNT).setNote('Cantidad de préstamos vigentes del prestatario (máximo ' + MAX_LOANS_PER_CLIENT + '). Se pinta de rojo al alcanzar el tope.');
  sh.getRange(1, PB.BALANCE).setNote('Saldo del préstamo: Total a Pagar − Total Pagado (capital + interés). El recargo por mora se muestra aparte en "Recargo por Mora (acum.)".');
  sh.getRange(1, PB.MORA_ACUM).setNote('Recargo por mora acumulado. Automático (derivado de "Pagos Atrasados") salvo que se cargue un valor en "Mora (ajuste)". 0 si el préstamo no está en mora.');
  sh.getRange(1, PB.MORA_ADJ).setNote('AJUSTE MANUAL de la mora de ESTE préstamo. Dejala VACÍA para el cálculo automático. Escribí 0 para CONDONAR (eximir) la mora, o un importe para FIJAR un recargo distinto. Afecta el "Recargo por Mora", el "Saldo con Mora", los avisos y el estado de cuenta.');
  sh.getRange(1, PB.BALANCE_MORA).setNote('Monto a pagar HOY = Saldo Pendiente + Recargo por Mora (ya con el ajuste manual, si lo hay).');
  sh.setFrozenRows(1);
  writeBorrowerFormulas_(sh, CFG.MAX_ROWS);
  formatBorrowerColumns_(sh, CFG.MAX_ROWS);
  protectFormulas_(sh, CFG.MAX_ROWS);
  writeBorrowerCapitalBanner_(sh);
}
/**
 * Banner "en vivo" con el capital disponible (= máximo asignable a un prestatario)
 * en la banda superior congelada, a la derecha de las columnas de datos. No inserta
 * filas ni afecta las sumas por columna (queda fuera de A:columna de datos).
 * Capital disponible = Fondo total − Capital prestado (col E) + Total cobrado (col L).
 */
function writeBorrowerCapitalBanner_(sh) {
  try {
    const ID = colL_(PB.LOAN_ID), CAP = colL_(PB.PRINCIPAL), PA = colL_(PB.PAID), col = PB.BALANCE_MORA + 1; // 1ª columna libre tras las columnas de datos
    if (sh.getMaxColumns() < col + 4) sh.insertColumnsAfter(sh.getMaxColumns(), col + 4 - sh.getMaxColumns()); // asegura lugar para el banner
    sh.getRange(1, col, 1, 4).breakApart();
    const banner = sh.getRange(1, col, 1, 4).merge();
    // SUMIF por "L-*" (no SUM a secas): ignora una eventual fila TOTAL/huérfana en las columnas de datos.
    banner.setFormula('="Capital disponible (máx. asignable): " & TEXT((' + fondoFormula_() +
      ')-SUMIF($' + ID + '$2:$' + ID + ',"L-*",$' + CAP + '$2:$' + CAP +
      ')+SUMIF($' + ID + '$2:$' + ID + ',"L-*",$' + PA + '$2:$' + PA + '),"$#,##0.00")');
    banner.setBackground('#38761d').setFontColor('#fff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    banner.setNote('Máximo que puede asignarse a un prestatario en total (sus préstamos sumados no pueden superar este monto). ' +
      'Baja al prestar capital y sube al cobrar. Cada prestatario admite hasta ' + MAX_LOANS_PER_CLIENT + ' préstamos.');
    sh.setColumnWidth(col, 220);
  } catch (e) { logError_('writeBorrowerCapitalBanner_', e); }
}

/**
 * Escribe SOLO las fórmulas de columnas calculadas de "Prestatarios" en el
 * nuevo layout normalizado. Reutilizable por la migración (NO borra datos).
 * Columnas: A=ID Préstamo, B=ID Cliente, C=Capital, D=Plazo, E=Tasa(f),
 * F=Fecha, G=Vencimiento(f), H=Interés(f), I=Total(f), J=Pagado(f), K=Saldo,
 * L=Estado(f), M=Contrato PDF, N=Enviado, O=Último Aviso, P/Q/R=firma.
 */
function writeBorrowerFormulas_(sh, N) {
  const P = CFG.SHEETS.PAYMENTS, C = CFG.SHEETS.CLIENTS;
  // Letras de columna derivadas del mapa PB (robustas a reordenamientos).
  const A = colL_(PB.LOAN_ID), CLI = colL_(PB.CLIENT_ID), CAP = colL_(PB.PRINCIPAL), TE = colL_(PB.TERM),
    RA = colL_(PB.RATE), FE = colL_(PB.LOAN_DATE), INT = colL_(PB.INTEREST), EST = colL_(PB.STATE),
    PPLOAN = colL_(PP.LOAN_ID), PPAMT = colL_(PP.AMOUNT);
  const fName = [], fRate = [], fDue = [], fITP = [], fState = [], fCount = [], fMora = [], fBalMora = [];
  for (let r = 2; r <= N + 1; r++) {
    fName.push([  // C Nombre y D DNI: se resuelven de "Clientes" por ID Cliente.
      `=IF($${CLI}${r}="","",IFERROR(VLOOKUP($${CLI}${r},'${C}'!$A:$B,2,FALSE),""))`,
      `=IF($${CLI}${r}="","",IFERROR(VLOOKUP($${CLI}${r},'${C}'!$A:$D,4,FALSE),""))`,
    ]);
    fRate.push([`=IF($${TE}${r}="","",IFS($${TE}${r}=15,0.25,OR($${TE}${r}=30,$${TE}${r}=1),0.5,OR($${TE}${r}=60,$${TE}${r}=2),1,OR($${TE}${r}=90,$${TE}${r}=3),1,TRUE,"⚠ PLAZO INVÁLIDO"))`]);
    fDue.push([`=IF(OR($${FE}${r}="",$${TE}${r}=""),"",IF($${TE}${r}=15,$${FE}${r}+15,IF(OR($${TE}${r}=30,$${TE}${r}=60,$${TE}${r}=90),$${FE}${r}+$${TE}${r},EDATE($${FE}${r},$${TE}${r}))))`]);
    fITP.push([
      `=IF(OR($${CAP}${r}="",NOT(ISNUMBER($${RA}${r}))),"",$${CAP}${r}*$${RA}${r})`,   // Interés
      `=IF($${INT}${r}="","",$${CAP}${r}+$${INT}${r})`,                                 // Total a Pagar
      `=IF($${A}${r}="","",SUMIF('${P}'!$${PPLOAN}:$${PPLOAN},$${A}${r},'${P}'!$${PPAMT}:$${PPAMT}))`, // Total Pagado
      borrowerBalanceFormula_(r),                                                        // Saldo Pendiente
    ]);
    fState.push([borrowerStateFormula_(r)]);                                             // Estado
    // Préstamos VIGENTES del prestatario (por ID Cliente, excluyendo PAGADO/SALDADO);
    // numérico para poder pintarlo. Coincide con clientLoanCount_ / el tope V-25.
    fCount.push([`=IF($${CLI}${r}="","",COUNTIFS($${CLI}$2:$${CLI}${N + 1},$${CLI}${r},$${EST}$2:$${EST}${N + 1},"<>${ST.PAID}",$${EST}$2:$${EST}${N + 1},"<>${ST.CLEARED}"))`]);
    fMora.push([borrowerMoraFormula_(r)]);              // Recargo por Mora (acum.) — de "Pagos Atrasados"
    fBalMora.push([borrowerBalanceWithMoraFormula_(r)]); // Saldo con Mora = Saldo + Recargo
  }
  sh.getRange(2, PB.NAME, N, 2).setFormulas(fName); // Nombre + DNI
  sh.getRange(2, PB.RATE, N, 1).setFormulas(fRate);
  sh.getRange(2, PB.DUE, N, 1).setFormulas(fDue);
  sh.getRange(2, PB.INTEREST, N, 4).setFormulas(fITP); // Interés, Total, Pagado, Saldo
  sh.getRange(2, PB.STATE, N, 1).setFormulas(fState);
  sh.getRange(2, PB.LOANCOUNT, N, 1).setFormulas(fCount); // Préstamos (de N)
  sh.getRange(2, PB.MORA_ACUM, N, 1).setFormulas(fMora);          // Recargo por Mora (acum.)
  sh.getRange(2, PB.BALANCE_MORA, N, 1).setFormulas(fBalMora);    // Saldo con Mora ("Mora (ajuste)" queda de entrada)
  sh.getRange(2, PB.MORA_ACUM, N, 3).setNumberFormat(CFG.CURRENCY_FMT); // formato moneda a las 3 columnas de mora
}
/**
 * Fórmula del RECARGO POR MORA acumulado de un préstamo, tomado de la hoja "Pagos Atrasados"
 * (la hoja que centraliza todos los préstamos en mora; se reconstruye con ⑤ Actualizar y a
 * diario). Devuelve 0 si el préstamo no está en mora, o si la hoja aún no existe.
 * `idCellA1` = celda del ID del préstamo (p. ej. "$A2").
 */
function moraLookupFormula_(idCellA1) {
  const L = CFG.SHEETS.LATE;
  const idCol = (LATE_HEADERS.indexOf('ID Préstamo') + 1) || 1;
  const moraCol = (LATE_HEADERS.indexOf('Recargo por Mora (acum.)') + 1) || 9;
  // El ID (columna de búsqueda) debe ser la 1.ª del rango del VLOOKUP.
  return `IFERROR(VLOOKUP(${idCellA1},'${L}'!$${colL_(idCol)}:$${colL_(moraCol)},${moraCol - idCol + 1},FALSE),0)`;
}
/**
 * Saldo pendiente (col Saldo) como fórmula VIVA: Total a Pagar − Total Pagado, nunca negativo
 * (capital + interés). El recargo por mora se lleva APARTE (columna "Recargo por Mora (acum.)")
 * y el total a pagar con mora en "Saldo con Mora".
 */
function borrowerBalanceFormula_(r) {
  const A = colL_(PB.LOAN_ID), TO = colL_(PB.TOTAL), PA = colL_(PB.PAID);
  return `=IF($${A}${r}="","",IF(OR($${TO}${r}="",$${PA}${r}=""),"",MAX(0,ROUND($${TO}${r}-$${PA}${r},2))))`;
}
/**
 * Columna "Recargo por Mora (acum.)": si la columna "Mora (ajuste)" tiene un NÚMERO, se usa
 * ese (0 = condonada; otro = recargo fijado a mano); si está vacía, se toma el recargo
 * automático de la hoja "Pagos Atrasados" (0 si el préstamo no está en mora).
 */
function borrowerMoraFormula_(r) {
  const A = colL_(PB.LOAN_ID), ADJ = colL_(PB.MORA_ADJ);
  return `=IF($${A}${r}="","",IF(ISNUMBER($${ADJ}${r}),MAX(0,$${ADJ}${r}),${moraLookupFormula_('$' + A + r)}))`;
}
/**
 * Columna "Saldo con Mora" = Saldo Pendiente + Recargo por Mora = MONTO A PAGAR HOY. Una vez
 * que el préstamo entra en mora, el recargo queda sumado en esta columna (derivado de "Pagos Atrasados").
 */
function borrowerBalanceWithMoraFormula_(r) {
  const A = colL_(PB.LOAN_ID), BAL = colL_(PB.BALANCE), MOR = colL_(PB.MORA_ACUM);
  return `=IF($${A}${r}="","",N($${BAL}${r})+N($${MOR}${r}))`;
}
/**
 * Estado (col L) como fórmula VIVA. PAGADO en cuanto Total Pagado (J) alcanza el
 * Total a Pagar (I); si no, VENCIDO cuando pasó el vencimiento (G) o ACTIVO.
 * No depende del saldo escrito por script, así que el estado cambia solo al
 * ingresar el pago (en la hoja "Pagos" o desde el panel).
 */
function borrowerStateFormula_(r) {
  const A = colL_(PB.LOAN_ID), TO = colL_(PB.TOTAL), PA = colL_(PB.PAID), VE = colL_(PB.DUE), BM = colL_(PB.BALANCE_MORA);
  // PAGADO sólo cuando el monto a pagar CON MORA ("Saldo con Mora") llega a 0: un préstamo
  // pagado tarde con recargo pendiente NO figura como PAGADO hasta saldar también la mora.
  return `=IF($${A}${r}="","",IF(AND(ISNUMBER($${TO}${r}),ISNUMBER($${PA}${r}),$${PA}${r}>=$${TO}${r}-0.009,N($${BM}${r})<=0.009),"${ST.PAID}",` +
    `IF(OR(${cuotaVencidaExpr_('$' + A + r)},AND($${VE}${r}<>"",TODAY()>$${VE}${r})),"${ST.OVERDUE}",IF(OR($${TO}${r}="",$${PA}${r}=""),"…","${ST.ACTIVE}"))))`;
}
/** Sub-expresión de fórmula: ¿el préstamo (celda con su ID) tiene alguna cuota VENCIDA en "Cuotas"? */
function cuotaVencidaExpr_(loanIdCellA1) {
  const C = CFG.SHEETS.INSTALLMENTS, L = colL_(CU.LOAN_ID), E = colL_(CU.STATE);
  return `COUNTIFS('${C}'!$${L}:$${L},${loanIdCellA1},'${C}'!$${E}:$${E},"${CST.OVERDUE}")>0`;
}

/** Formato, validaciones y formato condicional de "Prestatarios" (nuevo layout). */
function formatBorrowerColumns_(sh, N) {
  sh.getRange(2, PB.PRINCIPAL, N, 1).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(2, PB.RATE, N, 1).setNumberFormat('0%');
  sh.getRange(2, PB.LOAN_DATE, N, 2).setNumberFormat('yyyy-mm-dd'); // Fecha + Vencimiento
  sh.getRange(2, PB.INTEREST, N, 4).setNumberFormat(CFG.CURRENCY_FMT); // Interés..Saldo
  sh.getRange(2, PB.NOTICE, N, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, PB.SIGN_DATE, N, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  // allowInvalid(true): la lista es una AYUDA; no bloquea escrituras del script
  // ni migraciones (el Plazo heredado puede venir como número 1/2/15).
  sh.getRange(2, PB.TERM, N, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['15', '1', '2'], true).setAllowInvalid(true)
    .setHelpText('El plazo debe ser 15 días (25%), 1 mes (50%) o 2 meses (100%).').build());
  sh.getRange(2, PB.LOAN_DATE, N, 1).setDataValidation(datePicker_());
  const widths = [90, 90, 180, 110, 110, 90, 60, 110, 110, 110, 120, 110, 130, 90, 130, 130, 110, 120, 140, 150, 90, 160, 110];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, PB.DELETE, N, 1).insertCheckboxes();    // casilla para eliminar el préstamo
  sh.getRange(2, PB.SEND_DISB, N, 1).insertCheckboxes(); // casilla para enviar el recibo de desembolso
  // Préstamos (de N): número con sufijo " / N" y pintado de rojo al alcanzar el tope.
  const countRange = sh.getRange(2, PB.LOANCOUNT, N, 1);
  countRange.setNumberFormat('0" / ' + MAX_LOANS_PER_CLIENT + '"').setHorizontalAlignment('center');
  // Formato condicional unificado: franja por fila según Estado + firma + tope de préstamos.
  sh.setConditionalFormatRules(borrowerFormatRules_(sh));
  sh.getRange(2, PB.PDF, N, 1).setFontColor('#1155cc');
  sh.getRange(2, PB.SIGN_PDF, N, 1).setFontColor('#1155cc');
}

/**
 * Reglas de formato condicional de "Prestatarios", UNIFICADAS para que todas las rutas
 * (setup, reactivación, alta de columnas de firma) apliquen exactamente lo mismo — así
 * ninguna se pisa con otra (setConditionalFormatRules reemplaza TODAS las reglas):
 *  - Franja de color POR FILA según el Estado, sobre los campos relevantes (A → Estado).
 *  - Color propio de "Estado de Firma" (PENDIENTE/FIRMADO).
 *  - Rojo en "Préstamos (de N)" al alcanzar el tope.
 * Cobertura COMPLETA: filas 2..CFG.MAX_ROWS+1 (no solo las filas con datos actuales),
 * de modo que los préstamos nuevos también quedan coloreados sin re-ejecutar el setup.
 * Columnas resueltas por encabezado (con respaldo en PB) para tolerar layouts migrados.
 */
function borrowerFormatRules_(sh) {
  const N = CFG.MAX_ROWS, HI = headerIndex_(sh);
  // Asegura columnas para el banner "Capital disponible" (5 celdas tras la última de datos).
  const needCols = PB.BALANCE_MORA + 5;
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  const estC = colByAny_(HI, ['Estado']) || PB.STATE;
  const signC = colByAny_(HI, ['Estado de Firma']) || PB.SIGN_STATUS;
  const countC = colByAny_(HI, ['Préstamos (de ' + MAX_LOANS_PER_CLIENT + ')', 'Préstamos']) || PB.LOANCOUNT;
  const estL = colL_(estC);
  const rowRange = sh.getRange(2, 1, N, estC);        // A2:<Estado> — campos relevantes de la fila
  const signRange = sh.getRange(2, signC, N, 1);      // fuera de la franja de Estado (sin conflicto)
  const countRange = sh.getRange(2, countC, N, 1);
  // Banner "Capital disponible" (fila 1, a partir de la 1.ª columna libre tras los datos).
  const bannerRange = sh.getRange(1, PB.BALANCE_MORA + 1, 1, 5);
  // Regla por fila: colorea toda la franja según el valor del Estado (col. absoluta, fila relativa).
  const byState = (state, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + estL + '2="' + state + '"')
    .setBackground(bg).setFontColor(fg).setRanges([rowRange]).build();
  return [
    byState(ST.PAID, '#d9ead3', '#274e13'),
    byState(ST.CLEARED, '#d9ead3', '#274e13'),
    byState(ST.OVERDUE, '#f4cccc', '#990000'),
    byState(ST.ACTIVE, '#e8f0fe', '#1c4587'),
    cc_(signRange, SIGN.PENDING, '#f4cccc'),
    cc_(signRange, SIGN.SIGNED, '#b6d7a8'),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(MAX_LOANS_PER_CLIENT)
      .setBackground('#ea9999').setRanges([countRange]).build(),
    // Sobregiro: el banner muestra "…: -$…" cuando el capital disponible es negativo → fondo rojo.
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('-$')
      .setBackground('#cc0000').setFontColor('#ffffff').setBold(true).setRanges([bannerRange]).build(),
  ];
}

/**
 * Migración NO destructiva: inserta "Nombre" y "DNI" en "Prestatarios" justo
 * después de "ID Cliente" (columnas C y D), con fórmulas que las resuelven de
 * "Clientes". Google Sheets reajusta solo TODAS las referencias existentes
 * (Prestatarios, Resumen, Panel, Estadísticas, Estudio de Estados) al insertar
 * las columnas, por lo que es seguro sobre una hoja con préstamos reales.
 */
function insertNameDniColumns_() {
  const ss = getSS_(), bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs) { SpreadsheetApp.getUi().alert('No existe la hoja "Prestatarios". Ejecute ① Configurar primero.'); return; }
  const head = bs.getRange(1, 1, 1, Math.max(bs.getLastColumn(), 4)).getValues()[0].map(h => String(h).trim());
  if (head[PB.NAME - 1] === 'Nombre' && head[PB.DNI - 1] === 'DNI') {
    try { ss.toast('Las columnas Nombre y DNI ya existen.', 'Sin cambios', 5); } catch (e) {}
    return;
  }
  bs.insertColumnsAfter(PB.CLIENT_ID, 2); // 2 columnas nuevas después de "ID Cliente"
  bs.getRange(1, PB.NAME, 1, 2).setValues([['Nombre', 'DNI']])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
  bs.setColumnWidth(PB.NAME, 180); bs.setColumnWidth(PB.DNI, 110);
  const N = CFG.MAX_ROWS, C = CFG.SHEETS.CLIENTS, CLI = colL_(PB.CLIENT_ID), f = [];
  for (let r = 2; r <= N + 1; r++) f.push([
    `=IF($${CLI}${r}="","",IFERROR(VLOOKUP($${CLI}${r},'${C}'!$A:$B,2,FALSE),""))`,
    `=IF($${CLI}${r}="","",IFERROR(VLOOKUP($${CLI}${r},'${C}'!$A:$D,4,FALSE),""))`,
  ]);
  bs.getRange(2, PB.NAME, N, 2).setFormulas(f);
  try { ss.toast('Columnas "Nombre" y "DNI" agregadas después de ID Cliente.', 'Listo', 6); } catch (e) {}
}

/**
 * Migración NO destructiva de "Prestatarios":
 *  1) Corrige los encabezados O/P/Q que quedaron mal tras la migración (había un
 *     "Estado de Firma" DUPLICADO en la col O; se restaura a "Contrato PDF",
 *     "Contrato Enviado", "Último Aviso"). Queda un único "Estado de Firma" (col R).
 *  2) Limpia de la col "Contrato PDF" los valores de firma (PENDIENTE/FIRMADO)
 *     que se colaron por el duplicado.
 *  3) Agrega la casilla "Eliminar" (col U) para quitar un préstamo.
 */
function upgradeBorrowersColumns_() {
  const ss = getSS_(), bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs) { SpreadsheetApp.getUi().alert('No existe la hoja "Prestatarios". Ejecute ① Configurar primero.'); return; }
  const N = CFG.MAX_ROWS;
  // 1) Encabezados canónicos del bloque contrato/firma.
  const canon = {}; canon[PB.PDF] = 'Contrato PDF'; canon[PB.SENT] = 'Contrato Enviado';
  canon[PB.NOTICE] = 'Último Aviso'; canon[PB.SIGN_STATUS] = 'Estado de Firma';
  canon[PB.SIGN_DATE] = 'Fecha de Firma'; canon[PB.SIGN_PDF] = 'Contrato Firmado (PDF)';
  const head = bs.getRange(1, 1, 1, Math.max(bs.getLastColumn(), PB.DELETE)).getValues()[0].map(h => String(h).trim());
  Object.keys(canon).forEach(k => {
    const c = Number(k);
    if (head[c - 1] !== canon[c]) bs.getRange(1, c).setValue(canon[c]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
  });
  // 2) La col "Contrato PDF" (O) debe tener enlaces, no estados de firma: limpiar PENDIENTE/FIRMADO.
  const pdf = bs.getRange(2, PB.PDF, N, 1).getValues();
  bs.getRange(2, PB.PDF, N, 1).setValues(pdf.map(r => {
    const v = String(r[0]).trim().toUpperCase();
    return (v === SIGN.PENDING || v === SIGN.SIGNED) ? [''] : [r[0]];
  }));
  // 3) Casilla "Eliminar" (col U).
  if (String(bs.getRange(1, PB.DELETE).getValue()).trim() !== 'Eliminar') {
    bs.getRange(1, PB.DELETE).setValue('Eliminar').setFontWeight('bold').setBackground('#990000').setFontColor('#fff').setWrap(true);
    bs.setColumnWidth(PB.DELETE, 90);
    bs.getRange(1, PB.DELETE).setNote('Tildá para ELIMINAR ese préstamo (borra la fila; recuperable con el historial de versiones).');
  }
  bs.getRange(2, PB.DELETE, N, 1).insertCheckboxes();
  try { ss.toast('Prestatarios: "Estado de Firma" duplicado corregido y casilla "Eliminar" agregada.', 'Listo', 7); } catch (e) {}
}

/**
 * Migración NO destructiva: agrega las columnas de firma (S, T, U) a
 * "Prestatarios" si faltan, sin borrar datos. Segura de correr sobre una hoja
 * con préstamos reales. Se ejecuta desde el menú.
 */
function addSigningColumns_() {
  const ss = getSS_(), sh = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (!sh) { SpreadsheetApp.getUi().alert('No existe la hoja "Prestatarios". Ejecute ① Configurar primero.'); return; }
  const N = CFG.MAX_ROWS;
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), COL_SIGN_PDF)).getValues()[0]
    .map(h => String(h).trim());
  const labels = { [COL_SIGN_STATUS]: 'Estado de Firma', [COL_SIGN_DATE]: 'Fecha de Firma', [COL_SIGN_PDF]: 'Contrato Firmado (PDF)' };
  [COL_SIGN_STATUS, COL_SIGN_DATE, COL_SIGN_PDF].forEach(c => {
    if (head[c - 1] !== labels[c]) {
      sh.getRange(1, c).setValue(labels[c]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
    }
  });
  sh.setColumnWidth(COL_SIGN_STATUS, 120); sh.setColumnWidth(COL_SIGN_DATE, 140); sh.setColumnWidth(COL_SIGN_PDF, 150);
  sh.getRange(2, COL_SIGN_DATE, N, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL_SIGN_PDF, N, 1).setFontColor('#1155cc');
  // Marca como PENDIENTE los préstamos existentes sin estado de firma (activos/vencidos).
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    const est = sh.getRange(2, PB.STATE, last - 1, 1).getValues();
    const cur = sh.getRange(2, COL_SIGN_STATUS, last - 1, 1).getValues();
    const out = cur.map((r, i) => {
      const has = String(r[0]).trim();
      if (has) return [has];
      if (!String(ids[i][0]).trim()) return [''];
      const e = String(est[i][0]).toUpperCase();
      return [(e === ST.PAID || e === ST.CLEARED) ? SIGN.SIGNED : SIGN.PENDING];
    });
    sh.getRange(2, COL_SIGN_STATUS, out.length, 1).setValues(out);
  }
  // Reglas de formato condicional completas y consistentes (Estado por fila + firma + tope).
  sh.setConditionalFormatRules(borrowerFormatRules_(sh));
  setupFirmas_(ss);
  // Asegura (sin sobrescribir) las filas de Configuración de la firma. Crea la hoja
  // "Configuración" completa si aún no existe, para que los campos siempre aparezcan.
  if (!ss.getSheetByName(CFG.SHEETS.SETTINGS)) setupSettings_(ss);
  setSetting_('URL de la app web (enlaces a clientes)', getSetting_('URL de la app web (enlaces a clientes)') || '');
  setSetting_('Firmante 1 (nombre completo)', getSetting_('Firmante 1 (nombre completo)') || 'Denis Delroy Bell');
  setSetting_('Firma del Prestamista (enlace de Drive)', getSetting_('Firma del Prestamista (enlace de Drive)') || '');
  setSetting_('Firmante 2 (nombre completo)', getSetting_('Firmante 2 (nombre completo)') || 'Xoana Elizabeth Beron');
  setSetting_('Firma del Prestamista 2 (enlace de Drive)', getSetting_('Firma del Prestamista 2 (enlace de Drive)') || '');
  try { ss.toast('Columnas de firma agregadas y campos de firma en Configuración. Complete la URL de la app web.', 'Listo', 7); } catch (e) {}
}

/** Hoja de auditoría de firmas. */
const FIRMAS_HEADERS = ['Fecha y hora', 'ID Préstamo', 'Nombre', 'DNI', 'Correo', 'Método', 'Dispositivo',
  'Contrato Firmado (PDF)', 'IP', 'Hash SHA-256', 'CUIL', 'Domicilio'];
function setupFirmas_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SIGN);
  // Asegura el encabezado completo (acta de auditoría: IP, hash, CUIL, domicilio).
  const head = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, FIRMAS_HEADERS.length).getValues()[0].map(h => String(h).trim()) : [];
  if (head.join('|') !== FIRMAS_HEADERS.join('|')) {
    sh.getRange(1, 1, 1, FIRMAS_HEADERS.length).setValues([FIRMAS_HEADERS])
      .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff').setWrap(true);
    [150, 100, 180, 110, 200, 90, 260, 150, 120, 260, 120, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Hoja maestra de clientes (modelo normalizado). NO destructiva: crea la hoja y
 * el encabezado si faltan, pero nunca borra filas existentes.
 * Columnas: A=ID Cliente, B=Nombre, C=Correo, D=DNI, E=Teléfono,
 * F=DNI(norm), G=Correo(norm), H=Fecha alta, I=Actualizado.
 */
function setupClientes_(ss) {
  ss = ss || getSS_();
  const sh = getOrCreate_(ss, CFG.SHEETS.CLIENTS);
  // Defensivo: quitar validaciones heredadas del área usada (evita bloqueos al escribir).
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), CFG.MAX_ROWS + 1), CLIENTES_HEADERS.length).clearDataValidations();
  const head = sh.getRange(1, 1, 1, CLIENTES_HEADERS.length).getValues()[0].map(h => String(h).trim());
  if (head.join('|') !== CLIENTES_HEADERS.join('|')) {
    sh.getRange(1, 1, 1, CLIENTES_HEADERS.length).setValues([CLIENTES_HEADERS])
      .setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff').setWrap(true);
    [90, 200, 220, 120, 120, 120, 220, 140, 140].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(2, PC.CREATED, CFG.MAX_ROWS, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  sh.setFrozenRows(1);
  return sh;
}

/** Aviso (no bloqueo) al editar columnas calculadas. */
function protectFormulas_(sh, N) {
  try {
    // Columnas calculadas (derivadas del mapa PB): Nombre, DNI, Tasa, Vencimiento,
    // Interés, Total, Total Pagado, Saldo, Estado.
    [PB.NAME, PB.DNI, PB.RATE, PB.DUE, PB.INTEREST, PB.TOTAL, PB.PAID, PB.BALANCE, PB.STATE, PB.LOANCOUNT]
      .map(colL_).forEach(col => {
        const p = sh.getRange(col + '2:' + col + (N + 1)).protect()
          .setDescription('Columna calculada — no editar');
        p.setWarningOnly(true);
      });
  } catch (e) { /* ignora si no hay permisos */ }
}

/**
 * Fórmula del "Saldo Posterior" de una fila de Pagos: lo que falta pagar del préstamo
 * después de este pago = Total a Pagar − Σ pagos de ese préstamo hasta esta fila (inclusive).
 * Para un préstamo en cuotas equivale a la suma de las cuotas todavía impagas.
 */
function paymentBalanceFormula_(r) {
  const B = colL_(PP.LOAN_ID), D = colL_(PP.AMOUNT), BS = CFG.SHEETS.BORROWERS, bId = colL_(PB.LOAN_ID);
  return `=IF($${B}${r}="","",MAX(0,IFERROR(VLOOKUP($${B}${r},'${BS}'!$${bId}:$${colL_(PB.TOTAL)},${PB.TOTAL},FALSE),0)-SUMIFS($${D}$2:$${D}${r},$${B}$2:$${B}${r},$${B}${r})))`;
}
function setupPayments_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PAYMENTS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); // sh.clear() NO borra validaciones
  // Modelo normalizado: la identidad vive en "Clientes". Nombre y DNI se muestran
  // aquí por comodidad (se autocompletan al elegir el ID Préstamo); no son la
  // fuente de verdad. El ID Pago se autogenera (P-0001, P-0002, …).
  const H = ['ID Pago', 'ID Préstamo', 'Fecha de Pago', 'Monto Pagado', 'Saldo Posterior', 'Recibo Enviado', 'Nombre', 'DNI', 'Enviar recibo', 'Generar recibo'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#38761d').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange(2, PP.DATE, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd')
    .setDataValidation(datePicker_()); // muestra el selector de fecha (calendario)
  sh.getRange(2, PP.AMOUNT, CFG.MAX_ROWS, 2).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(2, PP.SEND_RECEIPT, CFG.MAX_ROWS, 1).insertCheckboxes(); // casilla "Enviar recibo"
  sh.getRange(2, PP.GEN_RECEIPT, CFG.MAX_ROWS, 1).insertCheckboxes(); // casilla "Generar recibo" (sin correo)
  // "Saldo Posterior" VIVO: lo que falta pagar del préstamo tras cada pago (= suma de cuotas impagas).
  const balF = []; for (let r = 2; r <= CFG.MAX_ROWS + 1; r++) balF.push([paymentBalanceFormula_(r)]);
  sh.getRange(2, PP.BALANCE, balF.length, 1).setFormulas(balF);
  [140, 90, 110, 120, 130, 150, 200, 110, 110, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, PP.SEND_RECEIPT).setNote('Tildá para ENVIAR por correo el recibo de pago de esa fila. Se destilda solo y deja el enlace en "Recibo Enviado". (Si el prestatario no tiene correo, usá "Generar recibo".)');
  sh.getRange(1, PP.GEN_RECEIPT).setNote('Tildá para GENERAR el PDF del recibo (con monto pagado y saldo pendiente) sin enviar correo. Se destilda solo y deja el enlace en "Recibo Enviado".');
  sh.getRange(1, PP.BALANCE).setNote('Lo que falta pagar del préstamo después de este pago (suma de las cuotas impagas). Se calcula solo.');
  // Ayuda para el ingreso manual: menú desplegable con los IDs de préstamo válidos.
  // allowInvalid(true) para no bloquear escrituras del script ni importaciones.
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (bs) sh.getRange(2, PP.LOAN_ID, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(bs.getRange('A2:A' + (CFG.MAX_ROWS + 1)), true).setAllowInvalid(true).build());
  sh.getRange(1, PP.LOAN_ID).setNote('Elija el ID del préstamo y escriba el importe en "Monto Pagado".\n' +
    'El "ID Pago" se genera solo y el "Nombre" y "DNI" se autocompletan.\n' +
    'El "Total Pagado", el "Saldo Pendiente" y el "Estado" del prestatario se actualizan solos.\n' +
    'Al saldar el total, el estado pasa a ' + ST.PAID + '.');
}
/**
 * Hoja "Cuotas": cronograma de pagos (una fila por cuota). NO destructiva — sólo asegura
 * encabezados/formatos; las filas de cuota (y sus fórmulas) las escribe cuotasForLoan_ al aprobar.
 * Monto Pagado / Saldo Cuota / Estado son fórmulas vivas que reparten el "Total Pagado" del
 * préstamo entre las cuotas (más antigua primero), reusando la hoja Pagos existente.
 */
function setupCuotas_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.INSTALLMENTS);
  sh.getRange(1, 1, 1, CUOTAS_HEADERS.length).setValues([CUOTAS_HEADERS])
    .setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  sh.getRange(2, CU.DUE, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');       // Fecha de Vencimiento
  sh.getRange(2, CU.AMOUNT, CFG.MAX_ROWS, 3).setNumberFormat(CFG.CURRENCY_FMT); // Monto / Pagado / Saldo
  sh.getRange(2, CU.NOTICE, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');    // Último Aviso
  [110, 100, 70, 150, 130, 130, 130, 110, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const est = sh.getRange(2, CU.STATE, CFG.MAX_ROWS, 1);
  sh.setConditionalFormatRules([
    cc_(est, CST.PAID, '#d9ead3'), cc_(est, CST.OVERDUE, '#f4cccc'), cc_(est, CST.PENDING, '#fff2cc'),
  ]);
  sh.getRange(1, CU.CUOTA_ID).setNote('Cronograma de cuotas. Los pagos se registran en "Pagos"; acá se reparten entre las cuotas (más antigua primero). "Monto Pagado", "Saldo Cuota" y "Estado" son fórmulas.');
  return sh;
}
/**
 * Escribe las N filas de cuota de un préstamo en "Cuotas" (con fórmulas vivas de
 * Pagado/Saldo/Estado). Reparte el Total del préstamo en cuotas iguales; la última
 * absorbe el redondeo para que Σ cuotas = Total exacto.
 * @param {Sheet} cs hoja Cuotas · @param {string} loanId · @param {Date} loanDate
 * @param {number} total Total a Pagar · @param {number} n cantidad de cuotas · @param {number[]} dueDays días de vencimiento (p.ej. [30,60,90] o [15])
 */
function cuotasForLoan_(cs, loanId, loanDate, total, n, dueDays) {
  const LID = colL_(CU.LOAN_ID), NUM = colL_(CU.NUM), AMT = colL_(CU.AMOUNT), DUE = colL_(CU.DUE);
  const B = CFG.SHEETS.BORROWERS, bIdL = colL_(PB.LOAN_ID), bPaidL = colL_(PB.PAID);
  const base = round2_(total / n);
  const rows = [];
  for (let k = 1; k <= n; k++) {
    const monto = (k === n) ? round2_(total - base * (n - 1)) : base;   // la última absorbe el redondeo
    const due = addDays_(loanDate, dueDays[k - 1]);
    rows.push([loanId + '-' + k, loanId, k, due, monto, '', '', '', '']);
  }
  const start = cs.getLastRow() + 1;
  cs.getRange(start, 1, rows.length, CUOTAS_HEADERS.length).setValues(rows);
  // Fórmulas vivas por fila (reparto del Total Pagado del préstamo, cuota más antigua primero).
  for (let i = 0; i < rows.length; i++) {
    const r = start + i;
    const pagadoAntes = `SUMIFS($${AMT}$2:$${AMT},$${LID}$2:$${LID},$${LID}${r},$${NUM}$2:$${NUM},"<"&$${NUM}${r})`;
    const loanPaid = `IFERROR(VLOOKUP($${LID}${r},'${B}'!$${bIdL}:$${bPaidL},${PB.PAID},FALSE),0)`;
    cs.getRange(r, CU.PAID).setFormula(`=MIN($${AMT}${r},MAX(0,${loanPaid}-(${pagadoAntes})))`);
    cs.getRange(r, CU.BALANCE).setFormula(`=$${AMT}${r}-$${colL_(CU.PAID)}${r}`);
    cs.getRange(r, CU.STATE).setFormula(
      `=IF($${colL_(CU.BALANCE)}${r}<=0.009,"${CST.PAID}",IF(TODAY()>$${DUE}${r},"${CST.OVERDUE}","${CST.PENDING}"))`);
  }
  cs.getRange(start, CU.DUE, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  cs.getRange(start, CU.AMOUNT, rows.length, 3).setNumberFormat(CFG.CURRENCY_FMT);
  return rows.length;
}
/**
 * Backfill NO destructivo: crea el cronograma para préstamos vigentes que aún no tienen cuotas.
 * Respeta el plazo EXISTENTE de cada préstamo (1 cuota a su vencimiento actual): no re-escalona ni
 * cambia montos de préstamos ya otorgados. Los préstamos nuevos sí usan los tramos/cuotas nuevos.
 */
function backfillCuotas_(ss) {
  ss = ss || getSS_();
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || setupCuotas_(ss);
  if (!bs || bs.getLastRow() < 2) return 0;
  const has = {};
  if (cs.getLastRow() >= 2) cs.getRange(2, CU.LOAN_ID, cs.getLastRow() - 1, 1).getValues()
    .forEach(r => { const id = String(r[0]).trim(); if (id) has[id] = true; });
  const rows = bs.getRange(2, 1, bs.getLastRow() - 1, PB.STATE).getValues();
  let n = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i], loanId = String(r[PB.LOAN_ID - 1]).trim();
    if (!/^L-/i.test(loanId) || has[loanId]) continue;
    const estado = String(r[PB.STATE - 1]).trim().toUpperCase();
    if (estado === ST.PAID || estado === ST.CLEARED) continue; // saldados: sin cronograma
    const fecha = r[PB.LOAN_DATE - 1], term = Number(r[PB.TERM - 1]) || 0, total = Number(r[PB.TOTAL - 1]) || 0;
    if (!(fecha instanceof Date) || total <= 0) continue;
    cuotasForLoan_(cs, loanId, fecha, round2_(total), 1, [termDays_(term)]);
    has[loanId] = true; n++;
  }
  return n;
}

function setupSummary_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SUMMARY); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); // sh.clear() NO borra validaciones
  const B = CFG.SHEETS.BORROWERS, C = CFG.SHEETS.CLIENTS;
  // Columnas de Prestatarios por NOMBRE de encabezado (robusto al layout migrado o nuevo).
  const bsS = ss.getSheetByName(B), BH = bsS ? headerIndex_(bsS) : {};
  const CLI = colL_(colByAny_(BH, ['ID Cliente']) || PB.CLIENT_ID),
    CAP = colL_(colByAny_(BH, ['Capital']) || PB.PRINCIPAL),
    INT = colL_(colByAny_(BH, ['Interés', 'Interes']) || PB.INTEREST),
    PA = colL_(colByAny_(BH, ['Total Pagado']) || PB.PAID),
    SAL = colL_(colByAny_(BH, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || PB.BALANCE),
    EST = colL_(colByAny_(BH, ['Estado']) || PB.STATE);
  // Se agrupa por ID Cliente; el nombre se resuelve con VLOOKUP a "Clientes".
  sh.getRange(1, 1, 1, 8).setValues([['ID Cliente', 'Prestatario', 'Préstamos', 'Capital Total', 'Interés Total',
    'Total Pagado', 'Saldo Total Pendiente', 'Estado']]).setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('A2').setFormula(`=IFERROR(UNIQUE(FILTER('${B}'!${CLI}2:${CLI},'${B}'!${CLI}2:${CLI}<>"")),"")`);
  const N = CFG.MAX_ROWS, f = [];
  for (let r = 2; r <= N + 1; r++) {
    f.push([
      `=IF($A${r}="","",IFERROR(VLOOKUP($A${r},'${C}'!$A:$B,2,FALSE),""))`,
      `=IF($A${r}="","",COUNTIF('${B}'!$${CLI}:$${CLI},$A${r}))`,
      `=IF($A${r}="","",SUMIF('${B}'!$${CLI}:$${CLI},$A${r},'${B}'!$${CAP}:$${CAP}))`,
      `=IF($A${r}="","",SUMIF('${B}'!$${CLI}:$${CLI},$A${r},'${B}'!$${INT}:$${INT}))`,
      `=IF($A${r}="","",SUMIF('${B}'!$${CLI}:$${CLI},$A${r},'${B}'!$${PA}:$${PA}))`,
      `=IF($A${r}="","",SUMIF('${B}'!$${CLI}:$${CLI},$A${r},'${B}'!$${SAL}:$${SAL}))`,
      `=IF($A${r}="","",IF($G${r}<=0.009,"${ST.CLEARED}",IF(COUNTIFS('${B}'!$${CLI}:$${CLI},$A${r},'${B}'!$${EST}:$${EST},"${ST.OVERDUE}")>0,"${ST.OVERDUE}","${ST.ACTIVE}")))`,
    ]);
  }
  sh.getRange(2, 2, N, 7).setFormulas(f);
  sh.getRange(2, 4, N, 4).setNumberFormat(CFG.CURRENCY_FMT);
  [90, 200, 80, 130, 130, 130, 150, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const range = sh.getRange(2, 8, N, 1);
  sh.setConditionalFormatRules([cc_(range, ST.CLEARED, '#b6d7a8'), cc_(range, ST.OVERDUE, '#ea9999'), cc_(range, ST.ACTIVE, '#fff2cc')]);
}

/** Estudio de Contratos — casillas para vista previa y envío del contrato. */
function setupAgreement_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.AGREEMENT); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const B = CFG.SHEETS.BORROWERS, C = CFG.SHEETS.CLIENTS;
  // Resuelve ID Cliente a partir del ID Préstamo elegido en B4 (modelo normalizado).
  const cId = `VLOOKUP($B$4,'${B}'!$A:$B,2,FALSE)`;
  sh.getRange('A1').setValue('ESTUDIO DE CONTRATOS').setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A2').setValue('Elija un ID de Préstamo, luego tilde ① Vista previa o ② Enviar contrato.').setFontColor('#666');
  sh.getRange('A4').setValue('ID Préstamo:').setFontWeight('bold');
  sh.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(B).getRange('A2:A' + (CFG.MAX_ROWS + 1)), true).setAllowInvalid(false).build())
    .setBackground('#fff2cc').setFontWeight('bold');
  const rows = [
    ['Prestatario', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$B,2,FALSE),"")`],
    ['DNI', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$D,4,FALSE),"")`],
    ['Correo', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$C,3,FALSE),"")`],
    ['Capital', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$E,5,FALSE),"")`],
    ['Plazo (meses)', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$F,6,FALSE),"")`],
    ['Interés', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$J,10,FALSE),"")`],
    ['Total a Pagar', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$K,11,FALSE),"")`],
    ['Fecha de Vencimiento', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$I,9,FALSE),"")`],
  ];
  sh.getRange(6, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(6, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));
  [9, 11, 12].forEach(r => sh.getRange(r, 2).setNumberFormat(CFG.CURRENCY_FMT));
  sh.getRange(13, 2).setNumberFormat('yyyy-mm-dd');
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange('A15').setValue('①  Vista previa del contrato  👁').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B15').setDataValidation(cb);
  sh.getRange('A16').setValue('②  Enviar contrato por correo  ✉').setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange('B16').setDataValidation(cb);
  sh.getRange('A18').setValue('Último PDF:').setFontWeight('bold');
  sh.getRange('A19').setValue('Estado:').setFontWeight('bold');
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 420);
  sh.getRange('A1:B2').setWrap(true);
}

function setupStatements_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.STATEMENTS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const B = CFG.SHEETS.BORROWERS, P = CFG.SHEETS.PAYMENTS, C = CFG.SHEETS.CLIENTS;
  // Resuelve ID Cliente a partir del ID Préstamo elegido en B4 (modelo normalizado).
  const cId = `VLOOKUP($B$4,'${B}'!$A:$B,2,FALSE)`;
  sh.getRange('A1').setValue('ESTUDIO DE ESTADOS DE CUENTA').setFontSize(16).setFontWeight('bold').setFontColor('#674ea7');
  sh.getRange('A2').setValue('Elija un ID de Préstamo; tilde ① para generar el PDF o ② para enviarlo por correo.').setFontColor('#666');
  sh.getRange('A4').setValue('ID Préstamo:').setFontWeight('bold');
  sh.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(B).getRange('A2:A' + (CFG.MAX_ROWS + 1)), true).setAllowInvalid(false).build())
    .setBackground('#fff2cc').setFontWeight('bold');
  const rows = [
    ['Prestatario', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$B,2,FALSE),"")`],     // 6
    ['DNI', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$D,4,FALSE),"")`],             // 7
    ['Correo', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$C,3,FALSE),"")`],          // 8
    ['Capital', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$E,5,FALSE),"")`],           // 9
    ['Total a Pagar', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$K,11,FALSE),"")`],     // 10
    ['Total Pagado', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$L,12,FALSE),"")`],     // 11
    // Días de Atraso: 0 si el préstamo está saldado (Saldo Pendiente B14 ≤ 0);
    // si no, días vencidos = HOY − Vencimiento. Al ser 0 cuando está pago, el
    // Recargo por Mora (B13) queda automáticamente en 0.
    ['Días de Atraso', `=IFERROR(IF($B$14="","",IF($B$14<=0.009,0,IF(VLOOKUP($B$4,'${B}'!$A:$I,9,FALSE)="","",MAX(0,TODAY()-VLOOKUP($B$4,'${B}'!$A:$I,9,FALSE))))),"")`], // 12
    // Recargo por mora = días de atraso MENOS la gracia, × recargo diario × Total, con
    // TOPE = (Tope de mora %) × Capital ($B$9). Espeja computeOutstanding_ (gracia + tope).
    ['Recargo por Mora (acum.)', `=IFERROR(IF(OR($B$12="",$B$12<=0),0,MIN(MAX(0,$B$12-IFERROR(VLOOKUP("Días de gracia antes de mora",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE),3))*$B$10*IFERROR(VLOOKUP("Recargo por Mora diario (%)",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)/100,0.05),$B$9*IFERROR(VLOOKUP("Tope de mora (% del capital)",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)/100,1))),"")`], // 13
    ['Saldo Pendiente', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$M,13,FALSE),"")`],  // 14
    ['Fecha de Vencimiento', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$I,9,FALSE),"")`], // 15
    ['ESTADO', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$N,14,FALSE),"")`],           // 16
  ];
  sh.getRange(6, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(6, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));
  [9, 10, 11, 13, 14].forEach(r => sh.getRange(r, 2).setNumberFormat(CFG.CURRENCY_FMT));
  sh.getRange(12, 2).setNumberFormat('0');
  sh.getRange(15, 2).setNumberFormat('yyyy-mm-dd');
  sh.getRange(13, 2).setFontColor('#900');
  const st = sh.getRange(16, 2); st.setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sh.setConditionalFormatRules([cc_(st, ST.PAID, '#b6d7a8'), cc_(st, ST.OVERDUE, '#ea9999'), cc_(st, ST.ACTIVE, '#fff2cc')]);
  const cbSt = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange('A18').setValue('①  Generar PDF del estado  👁').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B18').setDataValidation(cbSt);
  sh.getRange('A19').setValue('②  Enviar estado por correo  ✉').setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange('B19').setDataValidation(cbSt);
  sh.getRange('A20').setValue('Último PDF:').setFontWeight('bold');
  sh.getRange('A21').setValue('Estado:').setFontWeight('bold');
  sh.getRange('A23:C23').setValues([['Fecha de Pago', 'Monto Pagado', 'Saldo Posterior']])
    .setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff');
  sh.getRange('A24').setFormula(`=IFERROR(FILTER({'${P}'!$C$2:$C,'${P}'!$D$2:$D,'${P}'!$E$2:$E},'${P}'!$B$2:$B=$B$4),"— sin pagos registrados —")`);
  sh.getRange('A24:A' + (CFG.MAX_ROWS + 23)).setNumberFormat('yyyy-mm-dd');
  sh.getRange('B24:C' + (CFG.MAX_ROWS + 23)).setNumberFormat(CFG.CURRENCY_FMT);
  [170, 200, 160].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(23);
}

/* ---- Acciones por casilla en los estudios ---- */
function contractStudioPreview_(sh) {
  const status = m => sh.getRange('B19').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    const file = makeAgreementFile_(loan);
    sh.getRange('B18').setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    status('Vista previa lista — abra el PDF, luego tilde ②');
  } catch (err) { status('Error: ' + err.message); logError_('contractStudioPreview_', err); }
}
function contractStudioSend_(sh) {
  const status = m => sh.getRange('B19').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    if (!loan.email) { status('El prestatario no tiene correo'); return; }
    const file = makeAgreementFile_(loan);
    emailAgreement_(loan, file); markAgreementSent_(loan.loanId, file);
    sh.getRange('B18').setFormula('=HYPERLINK("' + file.getUrl() + '","Ver contrato")');
    status('Contrato enviado a ' + loan.email);
  } catch (err) { status('Error: ' + err.message); logError_('contractStudioSend_', err); }
}
function statementStudioPreview_(sh) {
  const status = m => sh.getRange('B21').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    const file = makeStatementFile_(loan);
    sh.getRange('B20').setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    status('PDF del estado generado — ábralo, luego tilde ② para enviarlo');
  } catch (err) { status('Error: ' + err.message); logError_('statementStudioPreview_', err); }
}
function statementStudioSend_(sh) {
  const status = m => sh.getRange('B21').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    if (!loan.email) { status('El prestatario no tiene correo'); return; }
    const file = makeStatementFile_(loan);
    const msg = sendStatementEmail_(loan, file);
    sh.getRange('B20').setFormula('=HYPERLINK("' + file.getUrl() + '","Ver estado de cuenta")');
    status(msg);
  } catch (err) { status('Error: ' + err.message); logError_('statementStudioSend_', err); }
}

/* ===================== ESTUDIO DE RECORDATORIOS (lista) ===================== */
const REM_HEADERS = ['ID Préstamo', 'Prestatario', 'Fecha de Vencimiento', 'Días (− = atraso)',
  'Estado', 'Total a Pagar', 'Saldo Pendiente', 'Recargo por Mora diario', 'Correo',
  'Generar PDF 👁', 'Enviar ✉', 'Último PDF'];
const REM_GEN_COL = 10, REM_SEND_COL = 11, REM_LINK_COL = 12;
const REM_HEADER_ROW = 3, REM_DATA_START = 4; // fila 1 título, fila 2 control, fila 3 encabezados

/** Construye la estructura (título, casilla de refresco y encabezados) y llena la lista. */
function setupReminders_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.REMINDER); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange('A1').setValue('ESTUDIO DE RECORDATORIOS DE PAGO').setFontSize(16).setFontWeight('bold').setFontColor('#b45f06');
  sh.getRange('C1').setValue('Préstamos vencidos (rojo) o próximos a vencer (amarillo). Tilde 👁 para generar el PDF o ✉ para enviarlo por correo.').setFontColor('#666').setWrap(true);
  sh.getRange('A2').setValue('🔄 Actualizar lista →').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.getRange('A1:A2').setWrap(true);
  sh.getRange(REM_HEADER_ROW, 1, 1, REM_HEADERS.length).setValues([REM_HEADERS])
    .setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(REM_HEADER_ROW);
  [110, 180, 130, 120, 110, 120, 130, 150, 210, 110, 90, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(REM_HEADER_ROW, 1).setNote('Se actualiza con "⑤ Actualizar", con la casilla 🔄 y automáticamente cada día. Incluye préstamos vencidos o que vencen dentro de los "Días de aviso de recordatorio" (Configuración, por defecto 7).');
  rebuildRemindersSheet_(ss);
  return sh;
}

/** Reconstruye la lista: un renglón por préstamo vencido o próximo a vencer, con colores. */
function rebuildRemindersSheet_(ss) {
  ss = ss || getSS_();
  let sh = ss.getSheetByName(CFG.SHEETS.REMINDER);
  if (!sh) { setupReminders_(ss); return; } // setupReminders_ ya llama a rebuild
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS), last = bs ? bs.getLastRow() : 0;
  // limpia datos, casillas y colores previos (conserva encabezados)
  const clearH = Math.max(sh.getMaxRows() - REM_HEADER_ROW, 1);
  sh.getRange(REM_DATA_START, 1, clearH, REM_HEADERS.length).clearContent().clearDataValidations().setBackground(null);
  if (!bs || last < 2) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  // Ventana del recordatorio en días (configurable). Aparecen los préstamos
  // vencidos y los que vencen dentro de esta cantidad de días. Predeterminado 7.
  // Si la fila de Configuración no existe (hojas anteriores), se crea con 7.
  if (String(getSetting_('Días de aviso de recordatorio')).trim() === '') {
    try { setSetting_('Días de aviso de recordatorio', '7'); } catch (e) { logError_('rebuildRemindersSheet_:setting', e); }
  }
  const noticeDays = Number(getSetting_('Días de aviso de recordatorio')) || 7;
  const feePct = lateFeeRate_(), rows = [], overdueFlags = [];
  // Cuotas en memoria: recordamos por CUOTA (la próxima impaga), no sólo por préstamo.
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS), cuotasByLoan = {};
  if (cs && cs.getLastRow() >= 2) {
    cs.getRange(2, 1, cs.getLastRow() - 1, CUOTAS_HEADERS.length).getValues().forEach(cr => {
      const lid = String(cr[CU.LOAN_ID - 1]).trim(); if (!lid) return;
      (cuotasByLoan[lid] = cuotasByLoan[lid] || []).push({
        num: Number(cr[CU.NUM - 1]) || 0, due: cr[CU.DUE - 1],
        monto: Number(cr[CU.AMOUNT - 1]) || 0, saldo: Number(cr[CU.BALANCE - 1]) || 0,
      });
    });
    Object.keys(cuotasByLoan).forEach(k => cuotasByLoan[k].sort((a, b) => a.num - b.num));
  }
  for (let row = 2; row <= last; row++) {
    const loan = readLoan_(bs, row);
    if (!loan.loanId || !(loan.dueDate instanceof Date)) continue;
    // Base de aviso: si el préstamo tiene cronograma, se usa la PRÓXIMA cuota impaga
    // (fecha, monto y saldo de esa cuota); si no, el vencimiento y saldo del préstamo.
    let dueDate = loan.dueDate, totalDue = loan.totalDue, cuotaTxt = '';
    let out = Number(bs.getRange(row, PB.BALANCE).getValue()) || 0;
    const qs = cuotasByLoan[loan.loanId];
    if (qs && qs.length) {
      const nextQ = qs.filter(q => q.saldo > 0.009 && q.due instanceof Date)[0];
      if (!nextQ) continue; // todas las cuotas pagadas
      dueDate = nextQ.due; totalDue = nextQ.monto; out = nextQ.saldo;
      cuotaTxt = ' — cuota ' + nextQ.num + '/' + qs.length;
    } else if (out <= 0.009) { continue; } // saldado (préstamo sin cronograma)
    const dd = new Date(dueDate); dd.setHours(0, 0, 0, 0);
    const daysUntil = Math.round((dd - today) / 86400000);
    if (daysUntil > noticeDays) continue; // ni vencido ni próximo
    const overdue = daysUntil < 0;
    const feeDay = round2_(totalDue * feePct);
    rows.push([loan.loanId, loan.name + cuotaTxt, dueDate, daysUntil, overdue ? ST.OVERDUE : 'PRÓXIMO',
      totalDue, out, feeDay, loan.email, false, false, '']);
    overdueFlags.push(overdue);
  }
  // ordena: más atrasados primero (daysUntil ascendente)
  const order = rows.map((r, i) => i).sort((a, b) => rows[a][3] - rows[b][3]);
  const sorted = order.map(i => rows[i]), sortedFlags = order.map(i => overdueFlags[i]);
  if (!sorted.length) return;
  const n = sorted.length;
  sh.getRange(REM_DATA_START, 1, n, REM_HEADERS.length).setValues(sorted);
  sh.getRange(REM_DATA_START, 3, n, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(REM_DATA_START, 4, n, 1).setNumberFormat('0');
  sh.getRange(REM_DATA_START, 6, n, 3).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(REM_DATA_START, REM_GEN_COL, n, 1).insertCheckboxes();
  sh.getRange(REM_DATA_START, REM_SEND_COL, n, 1).insertCheckboxes();
  // colores por renglón: vencido = rojo, próximo = amarillo (columnas de datos 1..9)
  for (let i = 0; i < n; i++) {
    sh.getRange(REM_DATA_START + i, 1, 1, 9).setBackground(sortedFlags[i] ? '#f4cccc' : '#fff2cc');
  }
}

/* ---- Acciones por renglón en el Estudio de Recordatorios ---- */
function reminderRowGenerate_(sh, row) { // 👁 genera el PDF y deja el enlace (sin correo)
  try {
    const loan = findLoanById_(String(sh.getRange(row, 1).getValue()).trim());
    if (!loan) { getSS_().toast('No se encontró el préstamo', '⚠', 5); return; }
    const file = makeReminderFile_(loan);
    sh.getRange(row, REM_LINK_COL).setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    getSS_().toast('PDF generado para ' + loan.name, 'Recordatorio', 4);
  } catch (err) { logError_('reminderRowGenerate_', err); getSS_().toast(err.message || String(err), '⚠ No se pudo generar', 8); }
}
function reminderRowSend_(sh, row) { // ✉ genera el PDF y lo envía por correo
  try {
    const loan = findLoanById_(String(sh.getRange(row, 1).getValue()).trim());
    if (!loan) { getSS_().toast('No se encontró el préstamo', '⚠', 5); return; }
    if (!loan.email) { getSS_().toast('El prestatario no tiene correo', '⚠', 6); return; }
    const file = makeReminderFile_(loan); emailReminder_(loan, file);
    sh.getRange(row, REM_LINK_COL).setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    getSS_().toast('Recordatorio enviado a ' + loan.email, 'Recordatorio', 5);
  } catch (err) { logError_('reminderRowSend_', err); getSS_().toast(err.message || String(err), '⚠ No se pudo enviar', 8); }
}

/** Crea únicamente la hoja "Estudio de Recordatorios" sin tocar Prestatarios ni Pagos. */
function createRemindersSheet_() {
  const ss = getSS_();
  setupReminders_(ss);
  ensureTriggers_();
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.REMINDER));
  ss.toast('Hoja "Estudio de Recordatorios" creada. Sus datos de Prestatarios y Pagos no fueron modificados.', 'Listo', 6);
}

/**
 * Instala/recrea SOLO la hoja "Estudio de Contratos" (título, selector de ID Préstamo,
 * datos por VLOOKUP y casillas ① Vista previa / ② Enviar) SIN borrar datos de
 * Prestatarios ni Pagos. Útil si la pestaña se eliminó por error.
 */
function createAgreementSheet_() {
  const ss = getSS_();
  setupAgreement_(ss);
  ensureTriggers_();
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.AGREEMENT));
  ss.toast('Hoja "Estudio de Contratos" creada. Sus datos de Prestatarios y Pagos no fueron modificados.', 'Listo', 6);
}

/**
 * Repara los TOTALES heredados de la migración: elimina la fila "TOTAL" incrustada
 * dentro de los datos de "Prestatarios" y "Resumen" (esta app NO usa filas TOTAL en la
 * grilla; los totales viven en Panel/Estadísticas y en el banner superior). Luego re-aplica
 * el banner endurecido y recalcula tableros. NO toca filas de préstamos/pagos/clientes.
 * Sin guion bajo final para que aparezca en la lista "Ejecutar" del editor de Apps Script.
 */
function repairTotals() {
  const ss = getSS_();
  const nB = deleteTotalRows_(ss, CFG.SHEETS.BORROWERS); // fila "TOTAL" incrustada en la grilla
  const nR = deleteTotalRows_(ss, CFG.SHEETS.SUMMARY);
  const b = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (b) writeBorrowerCapitalBanner_(b);                 // re-aplica el banner (SUMIF "L-*")
  try { refreshAll(true); } catch (e) { logError_('repairTotals:refreshAll', e); } // recalcula tableros
  SpreadsheetApp.getUi().alert('Reparación de totales',
    'Filas TOTAL eliminadas — Prestatarios: ' + nB + ', Resumen: ' + nR + '.\n' +
    'Banner y tableros recalculados. Los datos de préstamos, pagos y clientes no se modificaron.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Elimina toda fila cuyo texto en alguna de las primeras columnas sea exactamente
 * "TOTAL" (la fila de totales heredada de la migración, que puede estar en la col A o B).
 * Recorre de abajo hacia arriba para no desplazar índices. Devuelve cuántas eliminó.
 */
function deleteTotalRows_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName); if (!sh) return 0;
  const last = sh.getLastRow(), cols = Math.min(6, sh.getLastColumn());
  if (last < 2 || cols < 1) return 0;
  const vals = sh.getRange(2, 1, last - 1, cols).getValues();
  let deleted = 0;
  for (let i = vals.length - 1; i >= 0; i--) {
    const isTotal = vals[i].some(function (v) { return String(v).trim().toLowerCase() === 'total'; });
    if (isTotal) { try { sh.deleteRow(i + 2); deleted++; } catch (e) { logError_('deleteTotalRows_:' + sheetName, e); } }
  }
  return deleted;
}

/**
 * Instala/actualiza SOLO la hoja "Nuevos Prestatarios" (encabezados, casillas,
 * lista DNI/CUIL, formatos y columnas BCRA) SIN borrar datos de otras hojas ni
 * de las filas existentes. Seguro para correr sobre una hoja con datos reales.
 */
function setupNewOnly() {
  setupNew_(getSS_());
  try { getSS_().toast('Columnas BCRA agregadas a "Nuevos Prestatarios"', 'Listo', 5); } catch (e) {}
}

function setupNew_(ss) {
  let sh = ss.getSheetByName(CFG.SHEETS.NEW) || ss.getSheetByName('New Borrowers');
  if (!sh) sh = ss.insertSheet(CFG.SHEETS.NEW);
  else if (sh.getName() !== CFG.SHEETS.NEW && !ss.getSheetByName(CFG.SHEETS.NEW)) sh.setName(CFG.SHEETS.NEW);
  sh.getRange(1, 1, 1, NB.length).setValues([NB]).setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, 11, CFG.MAX_ROWS, 1).setDataValidation(cb); // Verificado?
  sh.getRange(2, 12, CFG.MAX_ROWS, 1).setDataValidation(cb); // Rechazar?
  sh.getRange(2, 14, CFG.MAX_ROWS, 1).setDataValidation(cb); // Verificar BCRA?
  const overCol = NB.indexOf('Anular límites') + 1;
  if (overCol > 0) sh.getRange(2, overCol, CFG.MAX_ROWS, 1).setDataValidation(cb); // Anular límites (override)
  // Casillas para marcar cada referencia como verificada por el prestamista.
  ['Ref 1 Validada?', 'Ref 2 Validada?'].forEach(function (h) {
    const c = NB.indexOf(h) + 1;
    if (c > 0) { sh.getRange(2, c, CFG.MAX_ROWS, 1).setDataValidation(cb); sh.setColumnWidth(c, 100); sh.getRange(1, c).setNote('Tildá cuando hayas contactado y verificado esta referencia.'); }
  });
  const paramDv = SpreadsheetApp.newDataValidation().requireValueInList(['DNI', 'CUIL'], true).build();
  const paramRange = sh.getRange(2, 15, CFG.MAX_ROWS, 1); // Parámetro BCRA
  paramRange.setDataValidation(paramDv);
  paramRange.setValue('DNI');
  sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, 6, CFG.MAX_ROWS, 1).setNumberFormat(CFG.CURRENCY_FMT);
  [150, 170, 200, 130, 120, 130, 100, 200, 190, 190, 90, 90, 220,
    110, 120, 120, 170, 110, 90, 130, 100, 170, 460].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 23, CFG.MAX_ROWS, 1).setWrap(true); // Resumen BCRA
  sh.getRange('K1').setNote('Tilde "Verificado?" para aprobar (pasa a Prestatarios) o "Rechazar?" para archivar en Rechazados.');
  sh.getRange('N1').setNote('Tilde "Verificar BCRA?" para consultar la Central de Deudores del BCRA. En "Parámetro BCRA" elegí DNI o CUIL: con CUIL se usa el valor de la columna CUIL (si está vacío, se deriva del DNI).');
  if (overCol > 0) sh.getRange(1, overCol).setNote('Tilde "Anular límites" ANTES de tildar "Verificado?" para aprobar aunque el prestatario ya tenga ' + MAX_LOANS_PER_CLIENT + ' préstamos, se exceda el capital disponible, o el BCRA marque RECHAZAR. NO omite las dos referencias con teléfono válido: siguen siendo obligatorias.');
}

function setupRejected_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.REJECTED);
  if (sh.getLastRow() === 0)
    sh.getRange(1, 1, 1, 6).setValues([['Fecha', 'Nombre completo', 'Correo', 'DNI', 'Teléfono', 'Motivo']])
      .setFontWeight('bold').setBackground('#990000').setFontColor('#fff');
  sh.setFrozenRows(1);
}

function setupErrors_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.ERRORS);
  if (sh.getLastRow() === 0)
    sh.getRange(1, 1, 1, 3).setValues([['Fecha', 'Función', 'Error']]).setFontWeight('bold').setBackground('#7f6000').setFontColor('#fff');
  sh.setFrozenRows(1);
}

function setupPanel_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PANEL); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); // sh.clear() NO borra validaciones
  const B = CFG.SHEETS.BORROWERS, C = CFG.SHEETS.CLIENTS, P = CFG.SHEETS.PAYMENTS, fondo = fondoFormula_();
  // Columnas por NOMBRE de encabezado (robusto al layout migrado o nuevo).
  const bsP = ss.getSheetByName(B), pgP = ss.getSheetByName(P);
  const BH = bsP ? headerIndex_(bsP) : {}, PH = pgP ? headerIndex_(pgP) : {};
  const idL = colL_(colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID);
  const cliL = colL_(colByAny_(BH, ['ID Cliente']) || PB.CLIENT_ID);
  const capL = colL_(colByAny_(BH, ['Capital']) || PB.PRINCIPAL);
  const intL = colL_(colByAny_(BH, ['Interés', 'Interes']) || PB.INTEREST);
  const salL = colL_(colByAny_(BH, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || PB.BALANCE);
  const estL = colL_(colByAny_(BH, ['Estado']) || PB.STATE);
  const firmaL = colL_(colByAny_(BH, ['Estado de Firma']) || PB.SIGN_STATUS);
  const venL = colL_(colByAny_(BH, ['Vencimiento', 'Fecha de Vencimiento']) || PB.DUE);
  const payIdL = colL_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID);
  const payAmtL = colL_(colByAny_(PH, ['Monto Pagado']) || PP.AMOUNT);
  const capitalPrestado = `SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${capL}:$${capL})`;
  const totalCobrado = pgP ? `SUMIF('${P}'!$${payIdL}:$${payIdL},"L-*",'${P}'!$${payAmtL}:$${payAmtL})` : '0';
  sh.getRange('A1').setValue('PANEL DEL PRESTAMISTA').setFontSize(18).setFontWeight('bold').setFontColor('#1c4587');
  const kpis = [
    ['Préstamos activos', `=COUNTIF('${B}'!$${estL}:$${estL},"${ST.ACTIVE}")`],           // 3  int
    ['Préstamos vencidos', `=COUNTIF('${B}'!$${estL}:$${estL},"${ST.OVERDUE}")`],         // 4  int
    ['Préstamos pagados', `=COUNTIF('${B}'!$${estL}:$${estL},"${ST.PAID}")`],             // 5  int
    ['Solicitudes pendientes', `=COUNTA('${CFG.SHEETS.NEW}'!$B$2:$B)`],                    // 6  int
    ['Contratos sin firmar', `=COUNTIF('${B}'!$${firmaL}:$${firmaL},"${SIGN.PENDING}")`], // 7  int
    ['Fondo total para prestar', `=${fondo}`],                                            // 8  money
    ['Capital prestado', `=${capitalPrestado}`],                                          // 9  money
    ['Total cobrado', `=${totalCobrado}`],                                                // 10 money
    ['Efectivo disponible para prestar', `=${fondo}-${capitalPrestado}+${totalCobrado}`], // 11 money (negativo = sobregiro)
    ['Interés contratado', `=SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${intL}:$${intL})`],      // 12 money
    ['Saldo pendiente total', `=SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${salL}:$${salL})`],   // 13 money
  ];
  sh.getRange(3, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]])).setFontWeight('bold');
  sh.getRange(3, 2, kpis.length, 1).setFormulas(kpis.map(k => [k[1]]));
  sh.getRange(3, 2, 5, 1).setNumberFormat('0');                 // filas 3–7 enteros
  sh.getRange(8, 2, 6, 1).setNumberFormat(CFG.CURRENCY_FMT);    // filas 8–13 moneda
  sh.getRange('A7').setFontColor('#990000'); sh.getRange('B7').setFontColor('#990000').setFontWeight('bold');
  sh.getRange('A7').setNote('Préstamos aprobados cuyo contrato aún no fue firmado por el prestatario. No desembolsar hasta la firma.');
  sh.getRange('A11').setFontColor('#38761d'); sh.getRange('B11').setFontColor('#38761d').setFontWeight('bold');
  sh.getRange('A11').setNote('Efectivo disponible = Fondo total − Capital prestado + Total cobrado. Baja al prestar y sube al cobrar. Un valor NEGATIVO (en rojo) indica sobregiro: se prestó más capital del disponible.');
  // Sobregiro: si el efectivo disponible es negativo, se muestra en rojo (el verde queda para ≥ 0).
  sh.setConditionalFormatRules([redIfNegativeRule_(sh.getRange('B11'))]);
  // Fila 14: Retiro disponible (instantánea; el cálculo usa ventanas de fechas que una fórmula de celda no expresa bien).
  sh.getRange('A14').setValue('Retiro disponible (sin frenar el crecimiento)').setFontWeight('bold').setFontColor('#38761d');
  try {
    const wd = withdrawableStats_();
    sh.getRange('B14').setValue(wd.withdrawable).setNumberFormat(CFG.CURRENCY_FMT).setFontColor('#38761d').setFontWeight('bold');
    sh.getRange('A14').setNote('Máximo retirable ahora sin frenar el ritmo de colocación. Solo libera ganancia realizada (interés cobrado) y ' +
      'retiene una reserva = colocación proyectada en ' + wd.horizonDays + ' días − cobros esperados + colchón por morosidad. ' +
      'Instantánea: se recalcula con "⑤ Actualizar" y con "💵 Retiro disponible".');
  } catch (e) { logError_('setupPanel_:withdrawable', e); }
  sh.getRange('A16').setValue('Próximos vencimientos (7 días):').setFontWeight('bold');
  // Nombre resuelto desde "Clientes" vía ID Cliente (col B). Vencimiento=G, Saldo=K, Estado=L.
  const nameArr = `ARRAYFORMULA(IFERROR(VLOOKUP('${B}'!$${cliL}2:$${cliL},'${C}'!$A:$B,2,FALSE),""))`;
  sh.getRange('A17').setFormula(
    `=IFERROR(SORT(FILTER({'${B}'!$${idL}2:$${idL},${nameArr},'${B}'!$${venL}2:$${venL},'${B}'!$${salL}2:$${salL}},` +
    `('${B}'!$${estL}2:$${estL}<>"${ST.PAID}")*('${B}'!$${venL}2:$${venL}>=TODAY())*('${B}'!$${venL}2:$${venL}<=TODAY()+7)),3,TRUE),"— sin vencimientos próximos —")`);
  sh.getRange('A16').setNote('Muestra préstamos no pagados que vencen dentro de 7 días.');
  sh.setColumnWidth(1, 240); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 130);
}

/* ===================== FONDO (capital disponible) ===================== */
/** Fórmula que lee el "Fondo total para prestar" de Configuración como número. */
function fondoFormula_() {
  return `IFERROR(VALUE(REGEXREPLACE(TO_TEXT(VLOOKUP("Fondo total para prestar",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)),"[^0-9\\.\\-]","")),0)`;
}
/** ¿El formulario web acepta nuevas solicitudes? Predeterminado: SÍ. */
function acceptingApplications_() {
  const raw = String(getSetting_('Aceptar solicitudes de préstamo (SÍ/NO)') || '').trim();
  if (raw === '') return true; // por defecto abierto si no está configurado
  return /^(s[íi]|si|true|verdadero|1|x|on|abierto)$/i.test(raw);
}
/**
 * Destinatarios de los avisos de nuevas solicitudes: "Correo del Prestamista"
 * más "Correos de aviso adicionales" (separados por coma, punto y coma o espacio).
 * Devuelve una lista de correos válidos, sin duplicados. Ej: ["a@x.com","b@y.com"].
 */
function noticeRecipients_() {
  const raw = String(getSetting_('Correo del Prestamista') || '') + ',' +
    String(getSetting_('Correos de aviso adicionales (separados por coma)') || '');
  const seen = {}, out = [];
  raw.split(/[,;\s]+/).forEach(function (e) {
    e = String(e).trim();
    const key = e.toLowerCase();
    if (!e || seen[key]) return;
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) { seen[key] = true; out.push(e); }
  });
  return out;
}
/** Fondo total configurado (número). */
function fondoTotal_() {
  return Number(String(getSetting_('Fondo total para prestar') || '0').replace(/[^0-9.\-]/g, '')) || 0;
}
/**
 * Efectivo disponible para prestar = Fondo total − Capital prestado + Total cobrado.
 * Baja cuando se presta (capital, col F) y sube cuando se cobra (pagos, col M).
 */
function fundStats_() {
  const ss = getSS_();
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  let lent = 0, collected = 0;
  // Capital prestado = suma de "Capital" de las filas de préstamo (ID "L-…").
  // Columnas por NOMBRE de encabezado (robusto al layout migrado o nuevo).
  if (bs && bs.getLastRow() >= 2) {
    const H = headerIndex_(bs), last = bs.getLastRow();
    const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
    const capC = colByAny_(H, ['Capital']) || PB.PRINCIPAL;
    const ids = bs.getRange(2, idC, last - 1, 1).getValues();
    const caps = bs.getRange(2, capC, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (/^L-/i.test(String(ids[i][0]).trim())) lent += Number(caps[i][0]) || 0;
  }
  // Total cobrado = suma de "Monto Pagado" en la hoja Pagos (misma fuente que el Panel:
  // "Efectivo disponible = MAX(0, Fondo − Capital prestado + Total cobrado)").
  if (pg && pg.getLastRow() >= 2) {
    const PH = headerIndex_(pg);
    const montoC = colByAny_(PH, ['Monto Pagado']) || PP.AMOUNT;
    const monts = pg.getRange(2, montoC, pg.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < monts.length; i++) collected += Number(monts[i][0]) || 0;
  }
  const total = fondoTotal_();
  // net = posición real (puede ser NEGATIVA = sobregiro: se prestó más capital del disponible).
  // available = net acotado a 0 para la capacidad de préstamo (V-09, retiro). El DISPLAY usa net.
  const net = round2_(total - lent + collected);
  return { total: total, lent: round2_(lent), collected: round2_(collected), available: round2_(Math.max(0, net)), net: net };
}
function fundAvailable_() { return fundStats_().available; }

/** Lee un ajuste numérico (días) de Configuración; usa el valor por defecto si falta o es inválido. */
function settingDays_(key, def) {
  const raw = String(getSetting_(key) || '').replace(/[^0-9.\-]/g, '').trim();
  const n = raw === '' ? NaN : Number(raw);
  return (isFinite(n) && n > 0) ? n : def;
}

/**
 * "Retiro disponible sin frenar el crecimiento": cuánto puede retirar el prestamista
 * ahora sin reducir el ritmo de colocación de préstamos. Recolecta las entradas desde
 * las hojas (una pasada por Prestatarios) y delega la aritmética a computeWithdrawable_.
 *
 * Solo se retira ganancia realizada (interés cobrado, base caja). Se reserva efectivo
 * para sostener el ritmo actual de colocación durante el horizonte configurable (H días),
 * neto de los cobros esperados en ese horizonte, más un colchón por morosidad.
 */
function withdrawableStats_() {
  const ss = getSS_();
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const fs = fundStats_();
  const H = settingDays_('Horizonte de protección del crecimiento (días)', 60);
  const W = Math.max(1, settingDays_('Ventana para estimar el ritmo de colocación (días)', 60));
  const DAY = 86400000;
  const now = new Date();
  const today0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const winStart = today0 - W * DAY;    // inicio de la ventana para medir el ritmo
  const horizonEnd = today0 + H * DAY;  // fin del horizonte de protección

  let profit = 0, originatedInWindow = 0, dueSoon = 0, totalContracted = 0, loans = 0, overdue = 0;
  if (bs && bs.getLastRow() >= 2) {
    const HI = headerIndex_(bs), last = bs.getLastRow();
    const idC = colByAny_(HI, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
    const capC = colByAny_(HI, ['Capital']) || PB.PRINCIPAL;
    const paidC = colByAny_(HI, ['Total Pagado']) || PB.PAID;
    const dateC = colByAny_(HI, ['Fecha del Préstamo', 'Fecha del Prestamo']) || PB.LOAN_DATE;
    const dueC = colByAny_(HI, ['Fecha de Vencimiento', 'Vencimiento']) || PB.DUE;
    const totC = colByAny_(HI, ['Total a Pagar', 'Total a pagar']) || PB.TOTAL;
    const balC = colByAny_(HI, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || PB.BALANCE;
    const estC = colByAny_(HI, ['Estado']) || PB.STATE;
    const maxC = Math.max(idC, capC, paidC, dateC, dueC, totC, balC, estC);
    const rows = bs.getRange(2, 1, last - 1, maxC).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!/^L-/i.test(String(r[idC - 1]).trim())) continue;
      loans++;
      const cap = Number(r[capC - 1]) || 0;
      const paid = Number(r[paidC - 1]) || 0;
      const bal = Number(r[balC - 1]) || 0;
      const est = String(r[estC - 1]).trim();
      profit += Math.max(0, paid - cap);        // interés cobrado (base caja)
      totalContracted += Number(r[totC - 1]) || 0;
      if (est === ST.OVERDUE) overdue++;
      const ld = r[dateC - 1];
      if (ld instanceof Date && ld.getTime() >= winStart && ld.getTime() <= today0) originatedInWindow += cap;
      const dd = r[dueC - 1];
      if (est !== ST.PAID && dd instanceof Date && dd.getTime() >= today0 && dd.getTime() <= horizonEnd) dueSoon += bal;
    }
  }
  const dailyRate = originatedInWindow / W;
  const recoveryRate = totalContracted > 0 ? Math.min(1, Math.max(0, fs.collected / totalContracted)) : 1;
  const expectedInflow = dueSoon * recoveryRate;
  const moraRate = loans > 0 ? overdue / loans : 0;
  const res = computeWithdrawable_({
    available: fs.available, profit: profit, dailyRate: dailyRate,
    horizonDays: H, expectedInflow: expectedInflow, moraRate: moraRate
  });
  res.dailyRate = round2_(dailyRate);
  res.expectedInflow = round2_(expectedInflow);
  res.moraRate = moraRate;
  res.recoveryRate = recoveryRate;
  res.horizonDays = H;
  res.windowDays = W;
  return res;
}

/* ===================== ESTADÍSTICAS + GRÁFICOS ===================== */
function showStats() {
  const ss = getSS_(); setupStats_(ss);
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.STATS));
}

/** Muestra el desglose de "Retiro disponible sin frenar el crecimiento" (cálculo en vivo). */
function showWithdrawable() {
  const s = withdrawableStats_();
  const pct = x => (Math.round((Number(x) || 0) * 1000) / 10) + '%';
  const msg =
    'Ganancia realizada (interés cobrado): ' + fmtMoney_(s.profit) + '\n' +
    'Efectivo disponible: ' + fmtMoney_(s.available) + '\n' +
    '\n' +
    'Ritmo de colocación: ' + fmtMoney_(s.dailyRate) + '/día  (ventana ' + s.windowDays + ' días)\n' +
    'Necesidad en ' + s.horizonDays + ' días (ritmo × horizonte): ' + fmtMoney_(s.need) + '\n' +
    'Cobros esperados en el horizonte: ' + fmtMoney_(s.expectedInflow) + '  (recuperación ' + pct(s.recoveryRate) + ')\n' +
    'Colchón por morosidad (' + pct(s.moraRate) + '): ' + fmtMoney_(s.buffer) + '\n' +
    'Reserva total a retener: ' + fmtMoney_(s.reserve) + '\n' +
    '\n' +
    '➤ RETIRO DISPONIBLE: ' + fmtMoney_(s.withdrawable) + '\n' +
    '\n' +
    'Solo se libera ganancia realizada y se conserva la reserva que mantiene el ritmo de colocación.';
  SpreadsheetApp.getUi().alert('Retiro disponible (sin frenar el crecimiento)', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Hoja de estadísticas de la cartera con indicadores y gráficos típicos de una agencia de préstamos. */
function setupStats_(ss) {
  ss = ss || getSS_();
  const sh = getOrCreate_(ss, CFG.SHEETS.STATS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); // sh.clear() NO borra validaciones
  sh.getCharts().forEach(c => sh.removeChart(c));
  const B = CFG.SHEETS.BORROWERS, P = CFG.SHEETS.PAYMENTS, fondo = fondoFormula_();
  const money = CFG.CURRENCY_FMT;
  // Columnas resueltas por NOMBRE de encabezado (robusto al layout migrado o nuevo).
  const bsS = ss.getSheetByName(B), pgS = ss.getSheetByName(P);
  const BH = bsS ? headerIndex_(bsS) : {}, PH = pgS ? headerIndex_(pgS) : {};
  const idL = colL_(colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID);
  const capL = colL_(colByAny_(BH, ['Capital']) || PB.PRINCIPAL);
  const intL = colL_(colByAny_(BH, ['Interés', 'Interes']) || PB.INTEREST);
  const totL = colL_(colByAny_(BH, ['Total a Pagar', 'Total a pagar']) || PB.TOTAL);
  const salL = colL_(colByAny_(BH, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || PB.BALANCE);
  const estL = colL_(colByAny_(BH, ['Estado']) || PB.STATE);
  const payIdL = colL_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID);
  const payAmtL = colL_(colByAny_(PH, ['Monto Pagado']) || PP.AMOUNT);
  // Sumas base coherentes con el Panel (Capital de préstamos "L-…"; cobrado desde Pagos).
  const capitalPrestado = `SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${capL}:$${capL})`;
  const totalCobrado = pgS ? `SUMIF('${P}'!$${payIdL}:$${payIdL},"L-*",'${P}'!$${payAmtL}:$${payAmtL})` : '0';
  const interesTotal = `SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${intL}:$${intL})`;
  const saldoTotal = `SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${salL}:$${salL})`;
  const totalAPagar = `SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${totL}:$${totL})`;
  const nLoans = `COUNTIF('${B}'!$${idL}:$${idL},"L-*")`;
  const efectivoDisp = `${fondo}-${capitalPrestado}+${totalCobrado}`; // negativo = sobregiro (se muestra en rojo)
  sh.getRange('A1').setValue('ESTADÍSTICAS DE LA CARTERA').setFontSize(18).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A2').setValue('Se actualiza con "⑤ Actualizar" y automáticamente cada día. Los gráficos se recalculan solos.').setFontColor('#666');

  // ---- Indicadores clave (A3:B…) ---- type: m=moneda, i=entero, p=porcentaje
  const kpis = [
    ['Fondo total para prestar', `=${fondo}`, 'm'],
    ['Capital prestado (colocado)', `=${capitalPrestado}`, 'm'],
    ['Total cobrado', `=${totalCobrado}`, 'm'],
    ['Efectivo disponible para prestar', `=${efectivoDisp}`, 'm'],
    ['Interés contratado (ganancia proyectada)', `=${interesTotal}`, 'm'],
    ['Saldo pendiente por cobrar', `=${saldoTotal}`, 'm'],
    ['Cartera en mora (saldo vencido)', `=SUMIF('${B}'!$${estL}:$${estL},"${ST.OVERDUE}",'${B}'!$${salL}:$${salL})`, 'm'],
    ['Ticket promedio', `=IFERROR(${capitalPrestado}/${nLoans},0)`, 'm'],
    ['Préstamos activos', `=COUNTIF('${B}'!$${estL}:$${estL},"${ST.ACTIVE}")`, 'i'],
    ['Préstamos vencidos', `=COUNTIF('${B}'!$${estL}:$${estL},"${ST.OVERDUE}")`, 'i'],
    ['Préstamos pagados', `=COUNTIF('${B}'!$${estL}:$${estL},"${ST.PAID}")`, 'i'],
    ['Tasa de morosidad (préstamos vencidos)', `=IFERROR(COUNTIF('${B}'!$${estL}:$${estL},"${ST.OVERDUE}")/(COUNTIF('${B}'!$${estL}:$${estL},"${ST.ACTIVE}")+COUNTIF('${B}'!$${estL}:$${estL},"${ST.OVERDUE}")),0)`, 'p'],
    ['Tasa de recuperación (cobrado / a pagar)', `=IFERROR(${totalCobrado}/${totalAPagar},0)`, 'p'],
    ['Rendimiento sobre capital (ROI)', `=IFERROR(${interesTotal}/${capitalPrestado},0)`, 'p'],
    ['Utilización del fondo', `=IFERROR((${capitalPrestado}-${totalCobrado})/${fondo},0)`, 'p'],
  ];
  const r0 = 4;
  sh.getRange(r0, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]])).setFontWeight('bold');
  sh.getRange(r0, 2, kpis.length, 1).setFormulas(kpis.map(k => [k[1]]));
  kpis.forEach((k, i) => {
    const cell = sh.getRange(r0 + i, 2);
    cell.setNumberFormat(k[2] === 'm' ? money : k[2] === 'p' ? '0.0%' : '0');
  });
  // resaltar el efectivo disponible (verde ≥ 0; rojo si es negativo = sobregiro)
  sh.getRange(r0 + 3, 1, 1, 2).setFontColor('#38761d').setFontWeight('bold');
  sh.getRange(r0 + 3, 1).setNote('Baja al prestar (capital) y sube al cobrar (pagos). Un valor NEGATIVO (en rojo) indica sobregiro: se prestó más capital del disponible.');
  sh.setConditionalFormatRules([redIfNegativeRule_(sh.getRange(r0 + 3, 2))]);

  // ---- Tablas de datos para los gráficos (columnas D:E) ----
  // 1) Estado de la cartera (torta)
  sh.getRange('D3').setValue('Estado de la cartera').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('D4:E7').setValues([
    ['Estado', 'Cantidad'],
    ['Activos', 0], ['Vencidos', 0], ['Pagados', 0],
  ]);
  sh.getRange('E5').setFormula(`=COUNTIF('${B}'!$${estL}:$${estL},"${ST.ACTIVE}")`);
  sh.getRange('E6').setFormula(`=COUNTIF('${B}'!$${estL}:$${estL},"${ST.OVERDUE}")`);
  sh.getRange('E7').setFormula(`=COUNTIF('${B}'!$${estL}:$${estL},"${ST.PAID}")`);

  // 2) Flujo de dinero (barras)
  sh.getRange('D10').setValue('Flujo de dinero').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('D11:E15').setValues([
    ['Concepto', 'Monto'],
    ['Capital prestado', 0], ['Interés contratado', 0], ['Total cobrado', 0], ['Saldo pendiente', 0],
  ]);
  sh.getRange('E12').setFormula(`=${capitalPrestado}`);
  sh.getRange('E13').setFormula(`=${interesTotal}`);
  sh.getRange('E14').setFormula(`=${totalCobrado}`);
  sh.getRange('E15').setFormula(`=${saldoTotal}`);
  sh.getRange('E12:E15').setNumberFormat(money);

  // 3) Liquidez del fondo (torta)
  sh.getRange('D18').setValue('Liquidez del fondo').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('D19:E21').setValues([
    ['Composición', 'Monto'],
    ['Efectivo disponible', 0], ['Por cobrar (saldo)', 0],
  ]);
  sh.getRange('E20').setFormula(`=${efectivoDisp}`);
  sh.getRange('E21').setFormula(`=${saldoTotal}`);
  sh.getRange('E20:E21').setNumberFormat(money);

  sh.setColumnWidth(1, 280); sh.setColumnWidth(2, 150);
  sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 120);

  // ---- Gráficos ----
  const pieStatus = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('D4:E7'))
    .setPosition(4, 7, 0, 0)
    .setOption('title', 'Distribución de préstamos por estado')
    .setOption('width', 460).setOption('height', 260)
    .setOption('pieHole', 0.4)
    .setOption('colors', ['#93c47d', '#e06666', '#f6b26b'])
    .build();
  sh.insertChart(pieStatus);

  const colFlow = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('D11:E15'))
    .setPosition(19, 7, 0, 0)
    .setOption('title', 'Flujo de dinero de la cartera')
    .setOption('width', 460).setOption('height', 260)
    .setOption('legend', 'none')
    .setOption('colors', ['#1c4587'])
    .build();
  sh.insertChart(colFlow);

  const pieFund = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('D19:E21'))
    .setPosition(34, 7, 0, 0)
    .setOption('title', 'Liquidez del fondo: disponible vs. por cobrar')
    .setOption('width', 460).setOption('height', 260)
    .setOption('colors', ['#38761d', '#e8a33d'])
    .build();
  sh.insertChart(pieFund);

  return sh;
}

const LATE_HEADERS = ['ID Préstamo', 'Prestatario', 'DNI', 'Correo', 'Teléfono', 'Fecha de Vencimiento',
  'Días de Atraso', 'Total a Pagar', 'Recargo por Mora (acum.)', 'Total Pagado', 'Saldo Pendiente',
  'Enviar Aviso ✉', 'Último Aviso Enviado'];
// Columnas (1-based) de control en "Pagos Atrasados".
const LATE_SEND_COL = 12, LATE_SENT_COL = 13;
function setupLate_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.LATE); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange(1, 1, 1, LATE_HEADERS.length).setValues([LATE_HEADERS])
    .setFontWeight('bold').setBackground('#990000').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  [110, 170, 120, 200, 120, 130, 100, 120, 160, 120, 130, 110, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('Se actualiza con "⑤ Actualizar" y automáticamente cada día. Lista los préstamos vencidos con saldo.');
  sh.getRange('L1').setNote('Tilde la casilla para enviar un aviso de mora por correo (en español). La casilla se destilda sola y la fecha de envío aparece en "Último Aviso Enviado".');
  return sh;
}

/** Reconstruye "Pagos Atrasados": un renglón por préstamo vencido con saldo, con recargo acumulado. */
function rebuildLateSheet_(ss) {
  ss = ss || getSS_();
  const sh = ss.getSheetByName(CFG.SHEETS.LATE) || setupLate_(ss) || ss.getSheetByName(CFG.SHEETS.LATE);
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  // limpia datos y casillas previas (conserva encabezados)
  const clearH = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, clearH, LATE_HEADERS.length).clearContent();
  sh.getRange(2, LATE_SEND_COL, clearH, 1).clearDataValidations();
  if (last < 2) return;
  const today = new Date(), feePct = lateFeeRate_(), grace = moraGraceDays_(), capFrac = moraCapFrac_(), rows = [];
  for (let row = 2; row <= last; row++) {
    const loan = readLoan_(bs, row);
    if (!loan.loanId || !(loan.dueDate instanceof Date)) continue;
    const days = daysLate_(loan.dueDate, today);
    if (days <= 0) continue;
    const pays = loanPayments_(loan.loanId);
    const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
    const base = Math.max(0, round2_(loan.totalDue - totalPaid));       // capital + interés pendiente
    // Recargo POR CUOTA, respetando el AJUSTE manual (condonar/cambiar) vía accruedMora_.
    const feeAccum = accruedMora_(loan, today);
    const out = round2_(base + feeAccum);                                // monto a pagar hoy = saldo + mora
    if (out <= 0) continue; // vencido pero saldado (y sin mora / mora condonada)
    const lastNotice = bs.getRange(row, PB.NOTICE).getValue(); // "Último Aviso" del prestatario
    rows.push([loan.loanId, loan.name, loan.dni, loan.email, loan.phone, loan.dueDate, days,
      loan.totalDue, feeAccum, totalPaid, out, false, (lastNotice instanceof Date) ? lastNotice : '']);
  }
  rows.sort((a, b) => b[6] - a[6]); // más atrasados primero
  if (rows.length) {
    sh.getRange(2, 1, rows.length, LATE_HEADERS.length).setValues(rows);
    sh.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 8, rows.length, 1).setNumberFormat(CFG.CURRENCY_FMT);
    sh.getRange(2, 9, rows.length, 3).setNumberFormat(CFG.CURRENCY_FMT);
    sh.getRange(2, 7, rows.length, 1).setNumberFormat('0');
    sh.getRange(2, LATE_SEND_COL, rows.length, 1).insertCheckboxes();
    sh.getRange(2, LATE_SENT_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  }
}

function setupHelp_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.HELP); sh.clear();
  const txt = [
    ['INSTRUCCIONES — GESTOR DE PRÉSTAMOS'],
    [''],
    ['1) Configuración: complete la hoja "Configuración" (nombre, contacto, jurisdicción).'],
    ['2) Formulario web: menú ⑥ para publicar y obtener el enlace. Compártalo con solicitantes.'],
    ['   (Si cambia el código, vuelva a implementar: Deploy ▸ Administrar implementaciones ▸ Nueva versión.)'],
    ['3) Las solicitudes llegan a "Nuevos Prestatarios" con fotos de DNI y CUIL.'],
    ['4) Verificar: tilde "Verificado?" para aprobar (pasa a "Prestatarios", crea carpeta y envía contrato).'],
    ['   Tilde "Rechazar?" para archivar en "Rechazados" (envía correo de rechazo).'],
    ['5) Pagos: use el Panel del prestamista (menú ▸ Abrir panel) para registrar pagos y enviar recibos.'],
    ['6) Intereses: 15 días = 25%, 1 mes = 50%, 2 meses = 100%. Tras el vencimiento: recargo por mora de ' + lateFeePctText_() + ' por día sobre el total a devolver' + (moraGraceDays_() > 0 ? ', tras ' + moraGraceDays_() + ' día(s) de gracia' : '') + ', con tope del ' + moraCapPctText_() + ' del capital.'],
    ['   Los préstamos vencidos con saldo aparecen en la hoja "Pagos Atrasados".'],
    ['   En "Pagos Atrasados", tilde "Enviar Aviso ✉" para mandar un aviso de mora por correo; la fecha queda en "Último Aviso Enviado".'],
    ['7) Estados: ACTIVO, VENCIDO, PAGADO. El Panel muestra indicadores y próximos vencimientos.'],
    ['   Saldados: use "✔ Mover préstamos saldados" (menú o Panel) para archivar los préstamos pagados en la hoja "Saldados".'],
    ['8) Recordatorios: se envían automáticamente antes del vencimiento y en mora.'],
    ['9) Fondo: complete "Fondo total para prestar" en "Configuración". El efectivo disponible baja al aprobar un préstamo (capital) y sube al registrar un pago. Se muestra en el Panel y en "Estadísticas".'],
    ['10) Cada nueva solicitud del formulario también avisa por correo al "Correo del Prestamista" con los datos y el efectivo disponible.'],
    ['11) Estadísticas: menú ④ abre la hoja "Estadísticas" con indicadores (morosidad, recuperación, ROI, utilización) y gráficos de la cartera.'],
    ['12) Abrir/cerrar solicitudes: en "Configuración", ponga "Aceptar solicitudes de préstamo (SÍ/NO)" en SÍ para recibir solicitudes o NO para mostrar "no hay préstamos disponibles" en el formulario.'],
    ['13) Errores: revise la hoja "Errores" si algo no funciona.'],
  ];
  sh.getRange(1, 1, txt.length, 1).setValues(txt);
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.setColumnWidth(1, 760); sh.getRange(1, 1, txt.length, 1).setWrap(true);
}

/* ===================== MOTOR DE SALDO ===================== */
/** Cantidad de cuotas del préstamo: 3 para el tramo grande (90 días / 3 meses), si no 1. */
function installmentCount_(loan) {
  const term = Number(loan && loan.term) || 0;
  return (term === 90 || term === 3) ? 3 : 1;
}
/**
 * Cronograma de cuotas de un préstamo, CALCULADO a partir de sus propios datos (Fecha del
 * Préstamo + cantidad de cuotas). No depende de la hoja "Cuotas" (que puede tener 1 sola
 * fila de respaldo) y usa la Fecha del Préstamo vigente (que se fija al firmar). Cuotas
 * mensuales iguales a los días 30/60/90; la última absorbe el redondeo.
 * Devuelve [{due:Date, amount, principalShare}] o null si es de pago único.
 */
function loanSchedule_(loan) {
  if (!loan || !(loan.loanDate instanceof Date)) return null;
  const n = installmentCount_(loan);
  if (n <= 1) return null; // pago único → comportamiento por defecto de computeOutstanding_
  const total = round2_(Number(loan.totalDue) || 0), principal = round2_(Number(loan.principal) || 0);
  const baseAmt = round2_(total / n), basePr = round2_(principal / n), out = [];
  for (let k = 1; k <= n; k++) {
    out.push({
      due: addDays_(loan.loanDate, 30 * k),
      amount: (k === n) ? round2_(total - baseAmt * (n - 1)) : baseAmt,
      principalShare: (k === n) ? round2_(principal - basePr * (n - 1)) : basePr,
    });
  }
  return out;
}

/**
 * Saldo con recargo por mora SIMPLE por día (contrato L-0021, Cláusula 5). Parámetros
 * configurables (Configuración): recargo diario (%), período de gracia (días sin mora tras
 * el vencimiento) y tope acumulado (% del capital).
 *
 * MODELO POR CUOTAS: si se pasa `schedule` (array de {due, amount, principalShare}) — o
 * para préstamos de 1 sola cuota, el pago único derivado del plazo — la mora se acumula
 * POR CUOTA sobre su propio monto, desde el fin de la gracia de CADA vencimiento, con tope
 * por cuota = capFrac × su parte del capital. Los pagos se aplican a la cuota más antigua
 * con saldo, primero a la mora y luego al capital+interés (orden del contrato). Un préstamo
 * de una sola cuota reproduce exactamente el comportamiento anterior.
 */
function computeOutstanding_(principal, rate, loanDate, payments, asOf, paymentCutoff, feeRate, graceDays, feeCapFrac, schedule) {
  const DAY = 86400000;
  const cutoff = (paymentCutoff instanceof Date) ? paymentCutoff : asOf;
  const feePct = (typeof feeRate === 'number') ? feeRate : lateFeeRate_();
  const grace = (typeof graceDays === 'number') ? Math.max(0, Math.floor(graceDays)) : moraGraceDays_();
  const capFrac = (typeof feeCapFrac === 'number') ? feeCapFrac : moraCapFrac_();
  const total = round2_(principal * (1 + rate));
  // Cronograma: por defecto, un solo pago al vencimiento (idéntico al comportamiento previo).
  let sched = (Array.isArray(schedule) && schedule.length) ? schedule : null;
  if (!sched) {
    const dueDate = rate === 0.25 ? addDays_(loanDate, 15) : addMonths_(loanDate, rate === 0.5 ? 1 : 2);
    sched = [{ due: dueDate, amount: total, principalShare: principal }];
  }
  const cuotas = sched.map(c => ({
    dueMs: (c.due instanceof Date ? c.due : new Date(c.due)).getTime(),
    amount: round2_(Number(c.amount) || 0),
    cap: round2_((Number(c.principalShare) || 0) * capFrac),
    balance: round2_(Number(c.amount) || 0),   // capital+interés pendiente de la cuota
    mora: 0,
    cursor: 0,
  })).sort((a, b) => a.dueMs - b.dueMs);
  cuotas.forEach(c => { c.cursor = c.dueMs; });

  // Acumula la mora de cada cuota impaga hasta el instante t (tras su gracia, con su tope).
  const accrueTo = t => {
    for (const c of cuotas) {
      const startMs = c.dueMs + grace * DAY;
      if (c.balance <= 0 || feePct <= 0 || c.mora >= c.cap || t <= startMs) { c.cursor = Math.max(c.cursor, t); continue; }
      const from = Math.max(c.cursor, startMs);
      const days = Math.floor((t - from) / DAY);
      if (days > 0) {
        const add = Math.min(round2_(round2_(c.amount * feePct) * days), round2_(c.cap - c.mora));
        c.mora = round2_(c.mora + add);
      }
      c.cursor = t;
    }
  };
  // Aplica un pago a la cuota más antigua con saldo: primero mora, luego capital+interés.
  const applyPay = amt => {
    let rem = round2_(amt);
    for (const c of cuotas) {
      if (rem <= 0) break;
      if (c.mora > 0) { const m = Math.min(rem, c.mora); c.mora = round2_(c.mora - m); rem = round2_(rem - m); }
      if (rem <= 0) break;
      if (c.balance > 0) { const b = Math.min(rem, c.balance); c.balance = round2_(c.balance - b); rem = round2_(rem - b); }
    }
  };

  const pays = payments.filter(p => new Date(p.date).getTime() <= cutoff.getTime())
    .map(p => ({ t: new Date(p.date).getTime(), amount: p.amount })).sort((a, b) => a.t - b.t);
  for (const p of pays) { accrueTo(p.t); applyPay(p.amount); }
  accrueTo(asOf.getTime());

  let out = 0;
  cuotas.forEach(c => { out = round2_(out + Math.max(0, c.balance) + Math.max(0, c.mora)); });
  return Math.max(0, round2_(out));
}

/**
 * Ajuste MANUAL de mora de un préstamo (columna "Mora (ajuste)" de "Prestatarios"). Devuelve
 * un número (0 = condonada; otro = recargo fijado a mano) o null si está vacía (automático).
 */
function moraOverrideValue_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh ? sh.getLastRow() : 0;
  if (last < 2) return null;
  const id = String(loanId).trim();
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const adjC = colByAny_(H, ['Mora (ajuste)']);
  if (!adjC) return null;
  const data = sh.getRange(2, 1, last - 1, Math.max(idC, adjC)).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idC - 1]).trim() === id) {
      const v = data[i][adjC - 1];
      return (typeof v === 'number' && isFinite(v)) ? Math.max(0, round2_(v)) : null;
    }
  }
  return null;
}
/**
 * Mora acumulada (recargo) de un préstamo a la fecha `asOf`, consistente con el motor por
 * cuotas. Se deriva como (saldo con mora) − (capital+interés pendiente). Si hay un AJUSTE
 * manual ("Mora (ajuste)"), se usa ese en lugar del cálculo. Usado por "Pagos Atrasados",
 * el estado de cuenta y los avisos.
 */
function accruedMora_(loan, asOf) {
  const ov = moraOverrideValue_(loan.loanId);
  if (ov != null) return ov;                 // condonada (0) o fijada a mano
  asOf = (asOf instanceof Date) ? asOf : new Date();
  const rate = loanRateForTerm_(loan.term);
  const pays = loanPayments_(loan.loanId);
  const withMora = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, asOf, new Date(9999, 0, 1), undefined, undefined, undefined, loanSchedule_(loan));
  const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
  const principalDue = Math.max(0, round2_(loan.totalDue - totalPaid));
  return Math.max(0, round2_(withMora - principalDue));
}
/**
 * Núcleo PURO del cálculo de "Retiro disponible sin frenar el crecimiento".
 * Solo aritmética (sin hoja ni fechas) para poder testearlo como computeOutstanding_.
 *
 * Modelo (decisiones con el usuario):
 *  - Se retira SOLO ganancia realizada (profit, base caja).
 *  - Se reserva efectivo para sostener el RITMO actual de colocación durante H días:
 *      need   = dailyRate * horizonDays           (capital a colocar en el horizonte)
 *      buffer = moraRate  * need                  (colchón que crece con la morosidad)
 *      reserve = max(0, need - expectedInflow) + buffer
 *  - Retirable = max(0, min(profit, available - reserve)).
 *
 * @param {{available:number, profit:number, dailyRate:number, horizonDays:number,
 *          expectedInflow:number, moraRate:number}} o
 * @return {{need:number, buffer:number, reserve:number, profit:number,
 *           available:number, withdrawable:number}}
 */
function computeWithdrawable_(o) {
  o = o || {};
  const available = Math.max(0, Number(o.available) || 0);
  const profit = Math.max(0, Number(o.profit) || 0);
  const dailyRate = Math.max(0, Number(o.dailyRate) || 0);
  const horizon = Math.max(0, Number(o.horizonDays) || 0);
  const inflow = Math.max(0, Number(o.expectedInflow) || 0);
  const mora = Math.min(1, Math.max(0, Number(o.moraRate) || 0));
  const need = round2_(dailyRate * horizon);
  const buffer = round2_(mora * need);
  const reserve = round2_(Math.max(0, need - inflow) + buffer);
  const withdrawable = Math.max(0, round2_(Math.min(profit, available - reserve)));
  return { need: need, buffer: buffer, reserve: reserve, profit: profit, available: available, withdrawable: withdrawable };
}
/** Días completos de atraso a hoy (0 si aún no vence). */
function daysLate_(dueDate, asOf) {
  if (!(dueDate instanceof Date)) return 0;
  const DAY = 86400000, a = (asOf || new Date());
  const d0 = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate()).getTime();
  const a0 = new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime();
  return Math.max(0, Math.floor((a0 - d0) / DAY));
}

function updateLoanOutstanding_(loanId) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  if (last < 2) return;
  const ids = bs.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) { writeOutstandingRow_(bs, i + 2); return; }
}
function updateAllOutstanding_() {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  for (let row = 2; row <= last; row++) writeOutstandingRow_(bs, row);
}
function writeOutstandingRow_(bs, row) {
  const loanId = String(bs.getRange(row, 1).getValue()).trim();
  const kCell = bs.getRange(row, PB.BALANCE), lCell = bs.getRange(row, PB.STATE);
  const moraCell = bs.getRange(row, PB.MORA_ACUM), balMoraCell = bs.getRange(row, PB.BALANCE_MORA);
  if (!loanId) { kCell.clearContent(); lCell.clearContent(); moraCell.clearContent(); balMoraCell.clearContent(); return; }
  // Saldo (fórmula), Recargo por Mora (de "Pagos Atrasados"), Saldo con Mora y Estado son
  // fórmulas vivas; las reafirmamos por si la fila fue pegada, migrada o creada al alta.
  kCell.setFormula(borrowerBalanceFormula_(row));
  moraCell.setFormula(borrowerMoraFormula_(row));
  balMoraCell.setFormula(borrowerBalanceWithMoraFormula_(row));
  lCell.setFormula(borrowerStateFormula_(row));
}

/**
 * Elimina el préstamo de una fila de "Prestatarios": limpia SOLO las celdas de
 * entrada y conserva la plantilla de fórmulas (Nombre/DNI en C:D, cálculos en
 * E..N). Recuperable con Archivo ▸ Historial de versiones. Se dispara con la
 * casilla "Eliminar" (columna U) vía onEdit.
 */
function deleteLoanRow_(bs, row) {
  try {
    const loanId = String(bs.getRange(row, PB.LOAN_ID).getValue()).trim();
    if (!loanId) return;
    bs.getRange(row, PB.LOAN_ID, 1, 2).clearContent();    // A:B ID, Cliente
    bs.getRange(row, PB.PRINCIPAL, 1, 2).clearContent();  // E:F Capital, Plazo
    bs.getRange(row, PB.LOAN_DATE).clearContent();        // Fecha del préstamo
    bs.getRange(row, PB.PDF, 1, 6).clearContent();        // bloque contrato/avisos/firma
    getSS_().toast('Préstamo ' + loanId + ' eliminado. (Recuperable con Archivo ▸ Historial de versiones.)', '🗑 Eliminado', 6);
  } catch (e) { logError_('deleteLoanRow_', e); getSS_().toast('No se pudo eliminar: ' + (e.message || e), '⚠', 8); }
}

/* ===================== CONTRATO (PDF + CORREO) ===================== */
/** Convierte HTML a PDF y lo guarda en la carpeta del prestatario. En TEST_MODE
 *  devuelve un archivo simulado (registra el HTML, no escribe en Drive). */
function savePdfToBorrower_(html, filename, loan) {
  if (TEST_MODE) return _fakeFile_(filename, html);
  const pdf = Utilities.newBlob(html, 'text/html', 'doc.html').getAs('application/pdf');
  const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
  const f = folder.createFile(pdf); f.setName(filename);
  return f;
}
function makeAgreementFile_(loan) {
  return savePdfToBorrower_(agreementHtml_(loan), 'Contrato ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}
/** PDF del contrato YA FIRMADO (incluye la imagen de la firma y el bloque de auditoría). */
function makeSignedAgreementFile_(loan, sig) {
  return savePdfToBorrower_(agreementHtml_(loan, sig), 'Contrato Firmado ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}
/** @param {object=} sig  Firma opcional: { dataUrl, signedAt (Date), method, userAgent }. */
function agreementHtml_(loan, sig) {
  const lender = companyName_(), footer = getSetting_('Pie del Contrato');
  const signatories = ownerSignatories_();   // ambos titulares firman todos los contratos
  const days = termDays_(loan.term) || 0;
  // Cronograma de cuotas (tramo grande = 3 cuotas mensuales; tramos chicos = 1 pago).
  const nCuotas = (typeof termForAmount_ === 'function' ? (termForAmount_(loan.principal).cuotas || 1) : 1);
  let cuotasBox = '', cuotasNota = 'Todo se devuelve en un solo pago en la fecha de vencimiento.';
  if (nCuotas > 1 && loan.loanDate) {
    const base = round2_(loan.totalDue / nCuotas);
    let filas = '';
    for (let k = 1; k <= nCuotas; k++) {
      const monto = (k === nCuotas) ? round2_(loan.totalDue - base * (nCuotas - 1)) : base;
      filas += `<tr><td class="k">Cuota ${k} de ${nCuotas}</td><td><b>${fmtMoney_(monto)}</b> — vence ${fmtDate_(addDays_(loan.loanDate, 30 * k))}</td></tr>`;
    }
    cuotasBox = `<div class="box" style="margin-top:8px"><div class="box-title">Plan de pago — ${nCuotas} cuotas mensuales</div><table style="margin:0">${filas}</table></div>`;
    cuotasNota = 'El total se devuelve en ' + nCuotas + ' cuotas mensuales iguales, según el cronograma de arriba (vencimientos a los 30, 60 y 90 días).';
  }
  const feeDay = round2_(loan.totalDue * lateFeeRate_());              // recargo diario sobre el total
  const moraGrace = moraGraceDays_();                                  // días de gracia antes de la mora
  const feeCap = round2_(loan.principal * moraCapFrac_());             // tope acumulado del recargo
  const maxTotal = round2_(loan.totalDue + feeCap);                    // total + tope del recargo
  const daysToCap = feeDay > 0 ? moraGrace + Math.ceil(feeCap / feeDay) : 0;
  const payMethod = loan.payMethod || 'Mercado Pago';
  // Fecha de firma del Prestamista: con valor cuando el contrato está firmado; en blanco si aún no.
  const firmadoEl = (sig && sig.dataUrl) ? fmtDate_(sig.signedAt || new Date()) : '__________';
  const fill = v => (v && String(v).trim()) ? esc_(v) : '<span style="color:#9aa4ad;letter-spacing:1px">________________________</span>';
  const kv = (k, v) => `<tr><td class="k">${k}</td><td>${v}</td></tr>`;
  const borrowerSig = (sig && sig.dataUrl) ? sig.dataUrl : '';
  const sigCell = (img, name, role) =>
    `<td>${img ? `<img src="${img}" alt="firma" style="max-height:44px;max-width:88%;display:block;margin:0 auto 1px">` : '<div style="height:40px"></div>'}` +
    `<div style="border-top:1px solid #333;padding-top:3px">${esc_(name)}<br><span style="font-size:8pt;color:#555">${esc_(role)}</span></div></td>`;
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    @page{size:A4;margin:11mm 12mm 12mm}
    *{box-sizing:border-box}
    body{font-family:"Carlito","Calibri",Arial,Helvetica,sans-serif;color:#1d2630;margin:0;line-height:1.3;font-size:9.2pt}
    h1{text-align:center;font-size:15pt;letter-spacing:.5px;border-bottom:2px solid #0f7b6c;padding-bottom:4px;color:#0f2b27;margin:2px 0 4px}
    h2{font-size:10pt;color:#0f7b6c;margin:10px 0 3px;border-bottom:1px solid #cfe3df;padding-bottom:2px;text-transform:uppercase;letter-spacing:.4px;page-break-after:avoid}
    p{margin:0 0 4px} ul{margin:3px 0 5px;padding-left:18px} li{margin-bottom:2px}
    table{width:100%;border-collapse:collapse;margin:5px 0;font-size:8.7pt} td,th{padding:3px 6px;border:1px solid #d8dee3;text-align:left;vertical-align:top}
    th{background:#eef4f3;color:#0f2b27;font-weight:bold}
    td.k{background:#f4faf9;font-weight:bold;width:46%;color:#3d4952}
    .lbl{font-weight:bold;color:#0f2b27}
    .box{border:1px solid #0f7b6c;border-left:4px solid #0f7b6c;background:#f4faf9;padding:6px 10px;margin:6px 0}
    .warn{border:1px solid #b8860b;border-left:4px solid #b8860b;background:#fdf8ec;padding:6px 10px;margin:6px 0}
    .box-title{font-weight:bold;color:#0f7b6c;text-transform:uppercase;letter-spacing:.4px;font-size:9pt;margin-bottom:3px}
    .warn .box-title{color:#8a6100}
    .meta{text-align:center;font-size:8.6pt;color:#5c6871;margin:0 0 3px}
    .box,.warn,table,.sign,tr{page-break-inside:avoid}
    .sign{margin-top:12px;width:100%} .sign td{border:none;text-align:center;padding:4px 10px;width:50%;vertical-align:bottom}
    .note{font-size:8.6pt;color:#5c6871}
    .foot{margin-top:14px;padding-top:6px;border-top:1px solid #dfe4e8;font-size:8.3pt;color:#5c6871;font-style:italic}</style></head><body>
    <table style="margin:0 0 4px"><tr>
      <td style="border:none;padding:0;vertical-align:middle">${brandBadge_(30, false)} <span style="vertical-align:middle">${brandWordmark_(15)}</span>
        <span style="font-size:8pt;color:#888;margin-left:4px">${esc_(BRAND.tag)}</span></td>
      <td style="border:none;padding:0;text-align:right;vertical-align:middle;font-size:8.4pt;color:#5c6871">Préstamo <b>${esc_(loan.loanId)}</b> · Emitido: ${fmtDate_(new Date())}<br>Moneda: Pesos argentinos (ARS)</td>
    </tr></table>
    <h1>CONTRATO DE PRÉSTAMO</h1>
    <p style="text-align:center;color:#3d4952">Celebrado el <b>${fmtDate_(new Date())}</b> entre ${esc_(signatories[0].name)} y
       ${esc_(signatories[1].name)}, que operan como «${esc_(lender)}» (el Prestamista), y <b>${esc_(loan.name)}</b> (el Prestatario).</p>

    <h2>1. Las partes</h2>
    <table>
      <tr><th style="width:50%">Prestamista</th><th style="width:50%">Prestatario</th></tr>
      <tr>
        <td><b>${esc_(signatories[0].name)}</b><br><b>${esc_(signatories[1].name)}</b></td>
        <td><b>${esc_(loan.name)}</b> · DNI ${esc_(loan.dni)} · CUIL ${fill(loan.cuil)}<br>
            Domicilio: ${fill(loan.direccion)}<br>${esc_(loan.email)} · ${fill(loan.phone)}</td>
      </tr>
    </table>
    <p>Los dos prestamistas prestan en conjunto, como personas particulares. Pagarle a cualquiera de ellos el total salda la deuda con
       ambos, y cualquiera de ellos puede reclamar la totalidad. No son un banco ni una entidad financiera, y el dinero prestado es propio.</p>

    <h2>2. El préstamo</h2>
    <div class="box">
      <div class="box-title">Lo que recibe el Prestatario y lo que devuelve</div>
      <table style="margin:0">
        ${kv('Número de préstamo', esc_(loan.loanId))}
        ${kv('Moneda', 'Pesos argentinos (ARS)')}
        ${kv('Monto que recibe', '<b>' + fmtMoney_(loan.principal) + '</b>')}
        ${kv('Duración', days + ' días')}
        ${kv('Entregado el', fmtDate_(loan.loanDate))}
        ${kv('Vence el', '<b>' + fmtDate_(loan.dueDate) + '</b>')}
        ${kv('Total a devolver', '<b>' + fmtMoney_(loan.totalDue) + '</b>')}
      </table>
    </div>
    ${cuotasBox}
    <p>No hay comisiones, seguros ni ningún otro cargo. ${cuotasNota} El Prestatario
       confirma que recibió estas condiciones antes de firmar y que tuvo tiempo de leerlas.</p>

    <h2>3. Entrega del préstamo</h2>
    <p>El Prestamista le entrega al Prestatario ${fmtMoney_(loan.principal)} por
       ☐ efectivo &nbsp; ☐ transferencia &nbsp; ☐ Mercado Pago. Los datos de la cuenta se le informan al Prestatario por separado y
       están disponibles cuando los pida. Comprobante N.º <span class="note">____________</span>.</p>
    <p>Al firmar, el Prestatario confirma que recibió el monto completo. Este contrato es su recibo. El comprobante de transferencia, o
       el recibo firmado en caso de efectivo, prueba que el dinero fue entregado. El Prestatario confirma que la cuenta usada es suya.</p>

    <h2>4. Devolución</h2>
    <p>El Prestatario debe pagar <b>${fmtMoney_(loan.totalDue)}</b> antes del <b>${fmtDate_(loan.dueDate)}</b>. No hace falta ningún aviso
       ni reclamo para que la deuda venza.</p>
    <p>El Prestatario puede pagar en cualquier momento del préstamo, en forma total o parcial, por cualquiera de los medios de abajo, sin
       aviso previo. Pagar antes no reduce el monto: los ${fmtMoney_(loan.interest)} de interés son fijos, así que el total es siempre
       ${fmtMoney_(loan.totalDue)}. No hay penalidad ni cargo por pagar antes. Esto no afecta el derecho a cancelar dentro de los 10 días
       de la cláusula 8.</p>
    <p><span class="lbl">Forma de pago elegida por el Prestatario:</span> ${esc_(payMethod)}.</p>
    <table>
      <tr><th style="width:22%">Medio</th><th>Cómo funciona</th></tr>
      <tr><td class="lbl">Efectivo</td><td>Entregado en persona a cualquiera de los dos prestamistas. Se da un recibo numerado en el momento. No hay que entregarle efectivo a ninguna otra persona, salvo que el Prestamista la haya autorizado por escrito. Sin recibo, el pago no cuenta.</td></tr>
      <tr><td class="lbl">Transferencia</td><td>Datos de la cuenta a pedido. Se informan cuando se entrega el préstamo, y otra vez en cualquier momento si el Prestatario los pide.</td></tr>
      <tr><td class="lbl">Mercado Pago</td><td>Alias/CVU a pedido, en las mismas condiciones. Se aceptan transferencias desde cuenta o billetera de Mercado Pago y también pagos por link. Las comisiones de Mercado Pago las paga el Prestatario y no tienen nada que ver con este contrato.</td></tr>
    </table>
    <p>El Prestamista informa los datos de pago dentro de las 24 horas de que se los pidan. Pagarle a otra persona no cuenta: solo los
       pagos a las personas y cuentas indicadas acá saldan la deuda. Cuándo cuenta un pago: en efectivo, la fecha del recibo; por
       transferencia y Mercado Pago, la fecha en que el dinero llega a la cuenta del Prestamista. El Prestatario manda el comprobante
       dentro de las 24 horas.</p>
    <p>La fecha de vencimiento es la pactada, sin importar si cae sábado, domingo o feriado. Los pagos parciales se aplican primero a
       gastos de cobranza, después a recargos por mora, después al interés y por último al monto prestado. El Prestamista no está obligado
       a aceptar un pago parcial. Recibos: el Prestamista da recibo de cada pago dentro de las 48 horas, y una carta que confirma que el
       préstamo quedó saldado dentro de los 5 días hábiles del último pago.</p>

    <h2>5. Pago fuera de término</h2>
    <p>Si el Prestatario no pagó en la fecha de vencimiento, ${moraGrace > 0 ? 'cuenta con <b>' + moraGrace + ' día(s) de gracia</b>; pasado ese plazo queda en mora' : 'queda en mora automáticamente desde el día siguiente'}. No hace falta aviso ni
       reclamo. Desde ese momento el interés de arriba deja de correr y lo reemplaza el recargo de abajo.</p>
    <p><span class="lbl">Recargo por mora:</span> ${lateFeePctText_()} por día sobre el total a devolver de ${fmtMoney_(loan.totalDue)} —
       es decir <b>${fmtMoney_(feeDay)} por cada día de atraso</b>${moraGrace > 0 ? ' (a partir del día ' + (moraGrace + 1) + ' de atraso)' : ''}, hasta que el préstamo esté pagado por completo. El recargo es simple:
       siempre se calcula sobre los ${fmtMoney_(loan.totalDue)}, y nunca se suman recargos sobre recargos ya devengados.</p>
    <div class="warn">
      <div class="box-title">Hay un máximo</div>
      <p style="margin:0">Los recargos dejan de acumularse cuando llegan a <b>${fmtMoney_(feeCap)}</b>. Lo máximo que el Prestatario
         puede llegar a deber por este contrato es <b>${fmtMoney_(maxTotal)}</b>, aparte de las costas judiciales.${daysToCap ? ' A ' + fmtMoney_(feeDay) + ' por día, ese máximo se alcanza a los ' + daysToCap + ' días.' : ''}</p>
    </div>
    <p>Un juez puede reducir los intereses que considere excesivos. Si eso pasa, el resto de este contrato sigue vigente.</p>

    <h2>6. Exigir todo antes de tiempo y gastos de cobranza</h2>
    <p>El Prestamista puede exigir todo el monto de inmediato si el Prestatario paga fuera de término, dio información falsa, cae en
       quiebra o concurso, o le trabaron o embargaron los bienes. El Prestatario paga los gastos razonables de cobranza que el
       Prestamista pueda documentar, hasta un 10 % del monto prestado, más las costas judiciales si el asunto llega a los tribunales.</p>

    <h2>7. Cómo nos comunicamos</h2>
    <p>Cualquiera de las dos partes puede usar cualquiera de estas vías, y todas valen como notificación: correo electrónico, WhatsApp y
       SMS (para todo, incluidas las notificaciones formales; se consideran recibidas el día hábil siguiente); llamada telefónica (para
       recordatorios y temas del día a día); carta o carta documento (para notificaciones formales, como la mora, las intimaciones de
       pago y las cesiones); y en persona (con constancia por escrito).</p>
    <p>Un mensaje enviado por una vía no hace falta repetirlo por otra, y el registro de envío o entrega prueba que se mandó. Hay que
       avisar a la otra parte dentro de los 5 días si cambia un domicilio, un correo o un teléfono; hasta entonces valen los datos que
       figuran. Las llamadas y mensajes sobre el pago se hacen entre las 08:00 y las 21:00, solo al Prestatario —nunca a su empleador, su
       familia ni a terceros— y no pueden ser abusivos, intimidatorios ni excesivos.</p>

    <h2>8. Cancelar dentro de los 10 días</h2>
    <p>El Prestatario puede arrepentirse y dejar sin efecto este préstamo dentro de los primeros 10 días. El plazo se cuenta desde que
       firmó o desde que recibió el dinero —vale la fecha más tardía de las dos—. No tiene que dar ningún motivo ni pagar penalidad:
       solo devuelve el capital de ${fmtMoney_(loan.principal)} dentro de las 72 horas. Hecho eso, el contrato queda sin efecto y no debe
       ningún interés.</p>

    <h2>9. Lo que declara el Prestatario</h2>
    <p>El Prestatario confirma que todo lo que le informó al Prestamista es verdadero y está actualizado; que las cuentas que dio son
       suyas; que no está en quiebra, en concurso ni con restricciones sobre sus bienes; que el dinero con el que va a pagar tiene origen
       lícito; que leyó la cláusula 2 antes de firmar; y que recibió una copia completa de este contrato.</p>

    <h2>10. Firma electrónica</h2>
    <p>Las dos partes acuerdan firmar este contrato en forma electrónica y aceptan que es válido. El Prestatario reconoce como propia la
       firma que figura abajo y acepta el registro de firma del Prestamista como prueba de quién firmó y de que el documento no fue
       alterado. Ninguna de las partes va a sostener que este contrato es inválido solo por haberse firmado electrónicamente.</p>

    <h2>11. Otras condiciones</h2>
    <ul>
      <li>Si alguna cláusula resulta inválida, el resto sigue vigente.</li>
      <li>Este documento, junto con su registro de firma, es el acuerdo completo y reemplaza cualquier cosa conversada antes.</li>
      <li>Los cambios tienen que hacerse por escrito y firmarse por las dos partes.</li>
      <li>El contrato se firma en español, que es la única versión con valor legal.</li>
      <li>Los títulos son solo orientativos. Se emite un único ejemplar electrónico y cada parte recibe una copia.</li>
    </ul>

    <h2>Firmas</h2>
    <p>Las dos partes firman de plena conformidad, confirmando que leyeron y entendieron todo el contrato —en particular lo que cuesta el
       préstamo (cláusula 2) y qué pasa si el pago se atrasa (cláusula 5).</p>
    <table class="sign"><tr>
      ${sigCell(signatories[0].sig, signatories[0].name, 'Prestamista · Firmado el ' + firmadoEl)}
      ${sigCell(signatories[1].sig, signatories[1].name, 'Prestamista · Firmado el ' + firmadoEl)}
    </tr></table>
    <table class="sign"><tr>
      ${sigCell(borrowerSig, loan.name, 'Prestatario · DNI ' + esc_(loan.dni) + ' · CUIL ' + (loan.cuil || '—'))}
      <td style="border:none"></td>
    </tr></table>
    ${sig && sig.dataUrl ? `<div class="box" style="margin-top:18px">
      <div class="box-title">Registro de firma</div>
      <table style="margin:0;font-size:10pt">
        ${kv('Firmó', esc_(loan.name) + ' · DNI ' + esc_(loan.dni))}
        ${kv('Correo confirmado', esc_(loan.email))}
        ${kv('Identidad verificada con', 'DNI ' + esc_(loan.dni) + ' y correo registrado; foto del DNI (frente y dorso, Anexo I)')}
        ${kv('Documento', 'Foto del DNI, frente y dorso — Anexo I')}
        ${kv('Fecha y hora', fmtDateTime_(sig.signedAt || new Date()))}
        ${kv('Cómo se firmó', sig.method === 'upload' ? 'Imagen de firma cargada' : 'Firma dibujada en pantalla')}
        ${kv('Dirección IP', fill(sig.ip))}
        ${kv('Dispositivo', esc_(String(sig.userAgent || '').slice(0, 200)))}
        ${kv('Huella del documento (SHA-256)', sig.hash ? '<span style="font-family:monospace;font-size:8.5pt;word-break:break-all">' + esc_(sig.hash) + '</span>' : '<span class="note">Se registra en la hoja «Firmas».</span>')}
      </table>
    </div>` : ''}
    <p class="foot">${esc_(footer)} Este contrato tiene efecto desde que las dos partes lo firman. El Prestatario confirma que lo leyó, que lo acepta y que lo firma electrónicamente.</p></body></html>`;
}
function emailAgreement_(loan, file) {
  sendBrandedEmail_(loan.email, 'Contrato de Préstamo — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', adjuntamos su contrato de préstamo ' + loan.loanId + '.',
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Adjuntamos su contrato de préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
    '<ul><li>Capital: <b>' + fmtMoney_(loan.principal) + '</b></li><li>Plazo: ' + loanTermLabel_(loan.term) + ' (' + loanRatePct_(loan.term) + ')</li>' +
    '<li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li><li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li></ul>',
    { attachments: [file.getAs('application/pdf')] });
}

/* ===================== RECORDATORIO DE PAGO (PDF) ===================== */
function makeReminderFile_(loan) {
  return savePdfToBorrower_(reminderHtml_(loan), 'Recordatorio ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}
function reminderHtml_(loan) {
  const lender = companyName_(), lPhone = getSetting_('Teléfono del Prestamista'),
    lAddr = getSetting_('Dirección del Prestamista'), penalty = getSetting_('Cláusula de Mora'),
    footer = getSetting_('Pie del Contrato');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const due = loan.dueDate instanceof Date ? new Date(loan.dueDate) : null;
  if (due) due.setHours(0, 0, 0, 0);
  const days = due ? Math.round((due - today) / 86400000) : null;
  const late = due ? daysLate_(loan.dueDate, today) : 0;
  const feeDay = round2_(loan.totalDue * lateFeeRate_());
  const moraGrace = moraGraceDays_();       // días de gracia antes de la mora
  const feeCap = round2_(loan.principal * moraCapFrac_());   // tope del recargo (% del capital)
  // Saldo base (Total a Pagar − Total Pagado): consistente con la hoja y el estado de cuenta,
  // y correcto para plazos en días (15/30/60). La mora se muestra aparte en feeLine.
  const totalPaid = loanPayments_(loan.loanId).reduce((s, p) => s + p.amount, 0);
  const outstanding = Math.max(0, round2_(loan.totalDue - totalPaid));
  const moraNow = outstanding > 0 ? accruedMora_(loan) : 0; // recargo POR CUOTA a hoy (0 si no hay atraso)
  // Encabezado y bloque que cambian según falte tiempo o esté en mora.
  let banner, timeLine, feeLine;
  if (days === null) {
    banner = { bg: '#fff2cc', color: '#7f6000', text: 'RECORDATORIO DE PAGO' };
    timeLine = '';
  } else if (days > 0) {
    banner = { bg: '#cfe2f3', color: '#1c4587', text: 'FALTAN ' + days + ' DÍA(S) PARA EL VENCIMIENTO' };
    timeLine = '<p>Su pago vence el <b>' + fmtDate_(loan.dueDate) + '</b>. Le quedan <b>' + days +
      ' día(s)</b> para pagar antes de que se aplique el recargo por mora.</p>';
  } else if (days === 0) {
    banner = { bg: '#fce5cd', color: '#b45f06', text: 'SU PAGO VENCE HOY' };
    timeLine = '<p>Su pago vence <b>hoy, ' + fmtDate_(loan.dueDate) + '</b>. Pague hoy para evitar el recargo por mora.</p>';
  } else {
    banner = { bg: '#f4cccc', color: '#990000', text: 'PAGO VENCIDO — ' + late + ' DÍA(S) DE ATRASO' };
    timeLine = '<p>Su pago venció el <b>' + fmtDate_(loan.dueDate) + '</b>, hace <b>' + late +
      ' día(s)</b>. Ya se está aplicando el recargo por mora indicado abajo.</p>';
  }
  if (moraNow > 0 || late > 0) {
    const feeAccum = moraNow;               // recargo por cuota (correcto para pago único y en cuotas)
    const capped = feeAccum >= feeCap;
    feeLine = '<p><b>Recargo por mora acumulado:</b> ' + fmtMoney_(feeAccum) +
      ' (' + lateFeePctText_() + ' por día sobre las cuotas vencidas' +
      (moraGrace > 0 ? ', tras ' + moraGrace + ' día(s) de gracia' : '') +
      (capped ? ', tope del ' + moraCapPctText_() + ' del capital alcanzado' : '') + ').' +
      (capped ? '' : ' El recargo sigue creciendo mientras haya cuotas impagas vencidas, hasta un tope de ' + fmtMoney_(feeCap) + '.') + '</p>';
  } else {
    feeLine = '<p><b>Recargo por mora:</b> si no paga a tiempo se aplicará un recargo del ' +
      lateFeePctText_() + ' por día sobre el total a devolver (' + fmtMoney_(loan.totalDue) + '), es decir <b>' +
      fmtMoney_(feeDay) + ' por día</b> de atraso' + (moraGrace > 0 ? ', tras ' + moraGrace + ' día(s) de gracia' : '') +
      ', con un tope del ' + moraCapPctText_() + ' del capital (' + fmtMoney_(feeCap) + ').</p>';
  }
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;border-bottom:2px solid #b45f06;padding-bottom:8px}
    h2{font-size:13pt;color:#b45f06;margin-top:24px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px} td{padding:6px 8px;border:1px solid #ccc}
    td.k{background:#faf0e6;font-weight:bold;width:45%}
    .banner{margin:16px 0;padding:12px;text-align:center;font-weight:bold;font-size:14pt;border-radius:6px;
      background:${banner.bg};color:${banner.color}}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>RECORDATORIO DE PAGO</h1>
    <p>Estimado/a <b>${esc_(loan.name)}</b>, le escribimos para recordarle el pago de su préstamo
       <b>${esc_(loan.loanId)}</b> con ${esc_(lender)}.</p>
    <div class="banner">${banner.text}</div>
    ${timeLine}
    <h2>Detalle del préstamo</h2>
    <table>
      <tr><td class="k">ID Préstamo</td><td>${esc_(loan.loanId)}</td></tr>
      <tr><td class="k">Total a pagar</td><td><b>${fmtMoney_(loan.totalDue)}</b></td></tr>
      <tr><td class="k">Saldo pendiente (hoy)</td><td><b>${fmtMoney_(outstanding)}</b></td></tr>
      <tr><td class="k">Fecha de vencimiento</td><td><b>${fmtDate_(loan.dueDate)}</b></td></tr>
    </table>
    <h2>Mora</h2>
    ${feeLine}
    ${penalty ? '<p>' + esc_(penalty) + '</p>' : ''}
    <p style="margin-top:20px">Si ya realizó el pago, por favor ignore este mensaje. Ante cualquier duda,
       contáctenos${lPhone ? ' al ' + esc_(lPhone) : ''}.</p>
    <p style="margin-top:24px">${esc_(lender)}${lAddr ? '<br>' + esc_(lAddr) : ''}</p>
    <p class="foot">${esc_(footer)}</p></body></html>`;
}
function emailReminder_(loan, file) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const due = loan.dueDate instanceof Date ? new Date(loan.dueDate) : null;
  if (due) due.setHours(0, 0, 0, 0);
  const days = due ? Math.round((due - today) / 86400000) : null;
  const when = days === null ? '' : days > 0 ? ' Faltan ' + days + ' día(s) para el vencimiento.'
    : days === 0 ? ' Su pago vence hoy.' : ' Su pago está vencido hace ' + daysLate_(loan.dueDate, today) + ' día(s).';
  sendBrandedEmail_(loan.email, 'Recordatorio de pago — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', le recordamos el pago de su préstamo ' + loan.loanId + '.' + when,
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Le recordamos el pago de su préstamo <b>' + esc_(loan.loanId) + '</b>.' + esc_(when) + '</p>' +
    '<ul><li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li>' +
    '<li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li>' +
    '<li>Recargo por mora: <b>' + fmtMoney_(round2_(loan.totalDue * lateFeeRate_())) + ' por día</b> de atraso (' + lateFeePctText_() + ' sobre el total a devolver, tope 100% del capital)</li></ul>' +
    '<p>Adjuntamos el recordatorio en PDF con el detalle. Si ya realizó el pago, ignore este mensaje.</p>',
    { attachments: [file.getAs('application/pdf')] });
}

/* ===================== PAGOS + RECIBO ===================== */
function recordPayment(loanId, amountStr, dateStr) {
  return guard_('recordPayment', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo "' + loanId + '".');
    const amount = Number(String(amountStr || '').replace(/[^0-9.\-]/g, ''));
    if (!amount || amount <= 0) throw new Error('Monto inválido.');
    const payDate = dateStr ? new Date(dateStr) : new Date();
    if (isNaN(payDate.getTime())) throw new Error('Fecha inválida.');
    // Saldo posterior alineado con las cuotas: Total a Pagar − todo lo pagado (incluido este pago).
    // Usa loanRate_ (día/mes/90) — termRate_ devuelve null para 30/60/90 y daba NaN.
    const existing = loanPayments_(loanId); existing.push({ date: payDate, amount: amount });
    const totalDue = round2_(loan.principal * (1 + loanRate_(loan.term)));
    const paidSoFar = existing.reduce((s, p) => s + (Number(p.amount) || 0), 0);
    const outAfter = Math.max(0, round2_(totalDue - paidSoFar));
    const psh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS);
    const payId = nextPaymentId_(psh);
    // Pagos: [ID Pago, ID Préstamo, Fecha, Monto, Saldo Posterior, Recibo, Nombre, DNI].
    psh.appendRow([payId, loanId, payDate, amount, '', '', loan.name || '', loan.dni || '']);
    const row = psh.getLastRow();
    psh.getRange(row, PP.BALANCE).setFormula(paymentBalanceFormula_(row)); // saldo vivo (no un valor fijo)
    psh.getRange(row, PP.DATE).setNumberFormat('yyyy-mm-dd'); psh.getRange(row, PP.AMOUNT, 1, 2).setNumberFormat(CFG.CURRENCY_FMT);
    updateLoanOutstanding_(loanId);
    if (loan.email) {
      // Recibo de pago como PDF numerado, adjunto al correo de confirmación y enlazado en "Recibo".
      let receiptFile = null;
      try { receiptFile = makePaymentReceiptFile_(loan, { payId: payId, date: payDate, amount: amount, out: outAfter }); }
      catch (e) { logError_('recordPayment:receiptPdf', e); }
      sendBrandedEmail_(loan.email, 'Pago recibido — ' + loanId,
        'Estimado/a ' + loan.name + ', confirmamos su pago de ' + fmtMoney_(amount) + '. Saldo: ' + fmtMoney_(outAfter) + '.',
        receiptHtml_(loan, payDate, amount, outAfter),
        receiptFile ? { attachments: [receiptFile.getAs('application/pdf')] } : {});
      if (receiptFile) psh.getRange(row, PP.RECEIPT).setFormula('=HYPERLINK("' + receiptFile.getUrl() + '","Ver recibo")');
      else psh.getRange(row, PP.RECEIPT).setValue(new Date());
    }
    return 'Pago registrado: ' + fmtMoney_(amount) + '. Saldo posterior: ' + fmtMoney_(outAfter) +
      (loan.email ? '. Recibo enviado a ' + loan.email : '');
  });
}
/** Envía (o reenvía) el recibo de pago de una fila existente de "Pagos". Devuelve el mensaje. */
function sendPaymentReceiptForRow_(psh, row) {
  return guard_('sendPaymentReceiptForRow_', function () {
    const loanId = String(psh.getRange(row, PP.LOAN_ID).getValue()).trim();
    if (!loanId) throw new Error('La fila no tiene ID de préstamo.');
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo "' + loanId + '".');
    if (!loan.email) throw new Error('El prestatario no tiene correo. Usá la casilla "Generar recibo" para crear el PDF sin enviarlo.');
    let payId = String(psh.getRange(row, PP.PAY_ID).getValue()).trim();
    if (!payId) { payId = 'P-' + String(nextPaymentSeq_(psh)).padStart(4, '0'); psh.getRange(row, PP.PAY_ID).setValue(payId); }
    const payDate = psh.getRange(row, PP.DATE).getValue() || new Date();
    const amount = Number(psh.getRange(row, PP.AMOUNT).getValue()) || 0;
    if (amount <= 0) throw new Error('El monto pagado de la fila es inválido.');
    const outAfter = Number(psh.getRange(row, PP.BALANCE).getValue()) || 0;
    let receiptFile = null;
    try { receiptFile = makePaymentReceiptFile_(loan, { payId: payId, date: payDate, amount: amount, out: outAfter }); }
    catch (e) { logError_('sendPaymentReceiptForRow_:receiptPdf', e); }
    sendBrandedEmail_(loan.email, 'Pago recibido — ' + loanId,
      'Estimado/a ' + loan.name + ', confirmamos su pago de ' + fmtMoney_(amount) + '. Saldo: ' + fmtMoney_(outAfter) + '.',
      receiptHtml_(loan, payDate, amount, outAfter),
      receiptFile ? { attachments: [receiptFile.getAs('application/pdf')] } : {});
    if (receiptFile) psh.getRange(row, PP.RECEIPT).setFormula('=HYPERLINK("' + receiptFile.getUrl() + '","Ver recibo")');
    else psh.getRange(row, PP.RECEIPT).setValue(new Date());
    return 'Recibo de pago enviado a ' + loan.email + '.';
  });
}
/** GENERA el PDF del recibo de una fila (monto pagado + saldo pendiente) SIN enviar correo. */
function generatePaymentReceiptForRow_(psh, row) {
  return guard_('generatePaymentReceiptForRow_', function () {
    const loanId = String(psh.getRange(row, PP.LOAN_ID).getValue()).trim();
    if (!loanId) throw new Error('La fila no tiene ID de préstamo.');
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo "' + loanId + '".');
    let payId = String(psh.getRange(row, PP.PAY_ID).getValue()).trim();
    if (!payId) { payId = 'P-' + String(nextPaymentSeq_(psh)).padStart(4, '0'); psh.getRange(row, PP.PAY_ID).setValue(payId); }
    const payDate = psh.getRange(row, PP.DATE).getValue() || new Date();
    const amount = Number(psh.getRange(row, PP.AMOUNT).getValue()) || 0;
    if (amount <= 0) throw new Error('El monto pagado de la fila es inválido.');
    const outAfter = Number(psh.getRange(row, PP.BALANCE).getValue()) || 0;
    const receiptFile = makePaymentReceiptFile_(loan, { payId: payId, date: payDate, amount: amount, out: outAfter });
    psh.getRange(row, PP.RECEIPT).setFormula('=HYPERLINK("' + receiptFile.getUrl() + '","Ver recibo")');
    return 'Recibo ' + payId + ' generado (saldo ' + fmtMoney_(outAfter) + ').';
  });
}
function receiptHtml_(loan, date, amount, out) {
  const cleared = out <= 0.009;
  return '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Confirmamos su pago del préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
    '<table style="border-collapse:collapse">' +
    row_('Fecha de pago', fmtDate_(date)) + row_('Monto pagado', fmtMoney_(amount)) + row_('Saldo pendiente', '<b>' + fmtMoney_(out) + '</b>') +
    '</table>' + (cleared ? '<p style="color:#38761d"><b>¡Su préstamo está totalmente pagado. Gracias!</b></p>' : '') +
    '<p>Saludos,<br>' + esc_(companyName_()) + '</p>';
}
function row_(k, v) { return '<tr><td style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa"><b>' + k + '</b></td><td style="padding:4px 10px;border:1px solid #ccc">' + v + '</td></tr>'; }

/* ===================== RECIBOS (PDF) ===================== */
/** Documento (PDF) genérico de recibo: encabezado de marca + tabla de datos. `rows` es [[clave, valorHTML], …]. */
function receiptDocHtml_(heading, subtitle, rows, footNote) {
  const footer = getSetting_('Pie del Contrato');
  const trs = rows.map(r => `<tr><td class="k">${r[0]}</td><td>${r[1]}</td></tr>`).join('');
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#1d2630;margin:44px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:19pt;letter-spacing:1px;border-bottom:2px solid #0f7b6c;padding-bottom:8px;color:#0f2b27}
    table{width:100%;border-collapse:collapse;margin:10px 0} td{padding:7px 10px;border:1px solid #d8dee3;vertical-align:top}
    td.k{background:#f4faf9;font-weight:bold;width:46%;color:#3d4952}
    .sub{text-align:center;color:#3d4952;margin:0 0 8px}
    .foot{margin-top:30px;padding-top:10px;border-top:1px solid #dfe4e8;font-size:9.5pt;color:#5c6871;font-style:italic}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>${esc_(heading)}</h1>
    ${subtitle ? `<p class="sub">${subtitle}</p>` : ''}
    <table>${trs}</table>
    ${footNote ? `<p>${footNote}</p>` : ''}
    <p class="foot">${esc_(footer || '')}</p></body></html>`;
}
/** Recibo de PAGO (PDF numerado) guardado en la carpeta del prestatario. */
function makePaymentReceiptFile_(loan, pay) {
  const cleared = pay.out <= 0.009;
  const rows = [
    ['Recibo N.º', esc_(pay.payId)],
    ['ID Préstamo', esc_(loan.loanId)],
    ['Prestatario', esc_(loan.name) + ' · DNI ' + esc_(loan.dni)],
    ['Fecha de pago', fmtDate_(pay.date)],
    ['Monto pagado', '<b>' + fmtMoney_(pay.amount) + '</b>'],
    ['Saldo pendiente', '<b>' + fmtMoney_(pay.out) + '</b>'],
  ];
  const note = cleared ? '<b style="color:#38761d">El préstamo está totalmente pagado. ¡Gracias!</b>'
    : 'Comprobante de pago parcial. El saldo pendiente continúa vigente según el contrato.';
  const html = receiptDocHtml_('RECIBO DE PAGO', 'Comprobante de pago — ' + esc_(companyName_()), rows, note);
  return savePdfToBorrower_(html, 'Recibo ' + pay.payId + ' - ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}
/** Recibo de DESEMBOLSO (PDF) guardado en la carpeta del prestatario. `info` = { date, amount, method, ref }. */
function makeDisbursementReceiptFile_(loan, info) {
  const rows = [
    ['ID Préstamo', esc_(loan.loanId)],
    ['Prestatario', esc_(loan.name) + ' · DNI ' + esc_(loan.dni)],
    ['Monto desembolsado', '<b>' + fmtMoney_(info.amount) + '</b>'],
    ['Fecha de desembolso', fmtDate_(info.date)],
    ['Medio de desembolso', esc_(info.method || '—')],
  ];
  if (info.ref) rows.push(['Referencia / Comprobante', esc_(info.ref)]);
  rows.push(['Total a pagar', fmtMoney_(loan.totalDue)]);
  rows.push(['Fecha de vencimiento', fmtDate_(loan.dueDate)]);
  const note = 'Este comprobante acredita la entrega del capital al Prestatario conforme a la Cláusula 5 del contrato ' + esc_(loan.loanId) + '.';
  const html = receiptDocHtml_('RECIBO DE DESEMBOLSO', 'Comprobante de entrega de fondos — ' + esc_(companyName_()), rows, note);
  return savePdfToBorrower_(html, 'Recibo Desembolso ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}
/** Envía por correo el recibo de desembolso al prestatario. */
function emailDisbursementReceipt_(loan, file, info) {
  if (!loan.email) return;
  sendBrandedEmail_(loan.email, 'Recibo de desembolso — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', confirmamos el desembolso de ' + fmtMoney_(info.amount) + ' de su préstamo ' + loan.loanId + '.',
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Confirmamos el <b>desembolso</b> de su préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
    '<ul><li>Monto desembolsado: <b>' + fmtMoney_(info.amount) + '</b></li><li>Fecha: <b>' + fmtDate_(info.date) + '</b></li>' +
    '<li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li><li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li></ul>' +
    '<p>Adjuntamos el recibo de desembolso.</p>', { attachments: [file.getAs('application/pdf')] });
}
/** Estado de firma ("FIRMADO"/"PENDIENTE") de un préstamo por ID, leído de "Prestatarios". */
function loanSignStatus_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh ? sh.getLastRow() : 0;
  if (last < 2) return '';
  const id = String(loanId).trim();
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const firmaC = colByAny_(H, ['Estado de Firma']) || PB.SIGN_STATUS;
  const data = sh.getRange(2, 1, last - 1, Math.max(idC, firmaC)).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][idC - 1]).trim() === id) return String(data[i][firmaC - 1] || '').trim().toUpperCase();
  return '';
}

/** Acción del prestamista: registra el desembolso y envía el recibo por correo. */
function sbSendDisbursementReceipt(loanId) {
  return guard_('sbSendDisbursementReceipt', function () {
    loanId = String(loanId).trim();
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo.');
    if (!loan.email) throw new Error('El prestatario no tiene correo.');
    // V-23 — no desembolsar sin firma. Bloquea también la vía del Panel (la casilla
    // "Enviar recibo desembolso" ya lo verifica en onEdit); enforce a nivel de función
    // para que ningún camino desembolse un contrato PENDIENTE.
    if (loanSignStatus_(loanId) !== SIGN.SIGNED)
      throw new Error('El contrato de ' + loanId + ' no está ' + SIGN.SIGNED + '. No se puede enviar el recibo de desembolso hasta que esté firmado.');
    const info = { date: new Date(), amount: loan.principal, method: (loan.payMethod || 'Transferencia'), ref: '' };
    const file = makeDisbursementReceiptFile_(loan, info);
    emailDisbursementReceipt_(loan, file, info);
    return 'Recibo de desembolso enviado a ' + loan.email + '.';
  });
}

/**
 * RECUPERACIÓN (C3) — Reestructura un préstamo vencido: congela la mora y fija un plan
 * de pago en cuotas. Acción MANUAL del prestamista (menú ▸ "Reestructurar…"). NO reescribe
 * el contrato original (Capital/Interés/Total): registra el plan en una nota auditable y
 * mueve la Fecha de Vencimiento al último vencimiento del plan. Como en este modelo el
 * Estado, el estudio de estado de cuenta, "Pagos Atrasados" y los avisos derivan de esa
 * fecha, la mora deja de crecer mientras el prestatario cumpla el plan.
 *
 * @param {string} [loanId] ID del préstamo (por defecto: el de la fila activa en "Prestatarios").
 * @param {number} [nCuotas] cantidad de cuotas mensuales (por defecto 3).
 * @param {Date}   [primeraFecha] vencimiento de la 1.ª cuota (por defecto: hoy + 30 días).
 * @return {{loanId:string, owed:number, plan:Array}} resumen del plan.
 */
function restructurarPrestamo_(loanId, nCuotas, primeraFecha) {
  return guard_('restructurarPrestamo_', function () {
    const ss = getSS_(), bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    if (!bs || bs.getLastRow() < 2) throw new Error('Falta la hoja "Prestatarios" o no tiene préstamos.');
    // Resolver el ID desde la fila activa si no se pasó explícitamente.
    if (!loanId && ss.getActiveSheet().getName() === CFG.SHEETS.BORROWERS) {
      const ar = ss.getActiveRange();
      if (ar) loanId = String(bs.getRange(ar.getRow(), PB.LOAN_ID).getValue()).trim();
    }
    loanId = String(loanId || '').trim();
    if (!/^L-/i.test(loanId)) throw new Error('Seleccioná la fila de un préstamo (ID "L-…") en "Prestatarios", o pasá el ID.');

    // Ubicar la fila del préstamo.
    let row = 0;
    const ids = bs.getRange(2, PB.LOAN_ID, bs.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) { row = i + 2; break; }
    if (!row) throw new Error('No se encontró el préstamo ' + loanId + '.');
    const loan = readLoan_(bs, row);

    // Monto adeudado HOY con la mora acumulada hasta hoy (se "congela" en el plan).
    const today = new Date();
    const rate = loanRateForTerm_(loan.term);
    const pays = loanPayments_(loanId);
    const owed = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, today, new Date(9999, 0, 1), undefined, undefined, undefined, loanSchedule_(loan));
    if (owed <= 0) throw new Error('El préstamo ' + loanId + ' no tiene saldo pendiente: no requiere reestructuración.');

    // Armar el plan: cuotas mensuales iguales; la última absorbe el redondeo.
    const n = Math.max(1, Math.floor(Number(nCuotas) || 3));
    const first = (primeraFecha instanceof Date) ? primeraFecha : addDays_(today, 30);
    const base = round2_(owed / n), plan = [];
    for (let k = 1; k <= n; k++) {
      const monto = (k === n) ? round2_(owed - base * (n - 1)) : base;
      plan.push({ n: k, due: addMonths_(first, k - 1), monto: monto });
    }
    const finalDue = plan[plan.length - 1].due;

    // CONGELAR MORA: mover la Fecha de Vencimiento al último vencimiento del plan. En este
    // modelo el Estado/estudio/"Pagos Atrasados"/avisos derivan de esta fecha → la mora se detiene.
    const dueC = colByAny_(headerIndex_(bs), ['Fecha de Vencimiento', 'Vencimiento']) || PB.DUE;
    const origDue = bs.getRange(row, dueC).getValue();
    bs.getRange(row, dueC).setValue(finalDue).setNumberFormat('yyyy-mm-dd');

    // Registrar el plan como nota auditable en la celda del ID del préstamo.
    const planTxt = plan.map(p => '  Cuota ' + p.n + '/' + n + ': ' + fmtMoney_(p.monto) + ' — vence ' + fmtDate_(p.due)).join('\n');
    bs.getRange(row, PB.LOAN_ID).setNote(
      'REESTRUCTURADO el ' + fmtDate_(today) + '.\n' +
      'Vencimiento original: ' + fmtDate_(origDue instanceof Date ? origDue : loan.dueDate) + '.\n' +
      'Adeudado congelado (capital + interés + mora a hoy): ' + fmtMoney_(owed) + '.\n' +
      'Plan de pago:\n' + planTxt + '\n' +
      'La mora deja de crecer mientras se cumpla el plan (vencimiento movido al ' + fmtDate_(finalDue) + ').\n' +
      'Registrá cada pago en la hoja "Pagos" como siempre.');

    try { ss.toast(loanId + ' reestructurado: ' + fmtMoney_(owed) + ' en ' + n + ' cuota(s). Mora congelada al ' + fmtDate_(today) + '.', '🔄 Reestructuración', 8); } catch (e) { }
    return { loanId: loanId, owed: owed, plan: plan };
  });
}

/** Menú: reestructura el préstamo de la fila activa (pide la cantidad de cuotas). */
function restructurarPrestamoUI() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Reestructurar préstamo vencido',
    'Se congela la mora del préstamo de la fila activa y se arma un plan de cuotas mensuales.\n\nCantidad de cuotas (por defecto 3):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(String(resp.getResponseText()).replace(/\D/g, ''), 10) || 3;
  const out = restructurarPrestamo_(null, n, null);
  if (out && out.owed) ui.alert('Reestructuración aplicada',
    out.loanId + ': ' + fmtMoney_(out.owed) + ' en ' + out.plan.length + ' cuota(s). El detalle quedó en la nota de la celda del ID. La mora dejó de crecer.', ui.ButtonSet.OK);
}

/**
 * Fija/condona/restaura la mora del préstamo de la FILA ACTIVA de "Prestatarios" escribiendo
 * en la columna "Mora (ajuste)". value === '' restaura el cálculo automático; un número la fija
 * (0 = condonar). El "Recargo por Mora", el "Saldo con Mora" y el Estado se actualizan solos.
 */
function setMoraAjusteFilaActiva_(value, label) {
  const ss = getSS_(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== CFG.SHEETS.BORROWERS) { ui.alert('Abrí la hoja "Prestatarios", seleccioná la fila del préstamo y volvé a usar esta opción.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Seleccioná la fila de un préstamo.'); return; }
  const H = headerIndex_(sh);
  const adjC = colByAny_(H, ['Mora (ajuste)']);
  if (!adjC) { ui.alert('Falta la columna "Mora (ajuste)". Ejecutá ① Configurar o 🆕 Aplicar novedades.'); return; }
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const loanId = String(sh.getRange(row, idC).getValue()).trim();
  if (!loanId) { ui.alert('La fila seleccionada no tiene un préstamo.'); return; }
  const cell = sh.getRange(row, adjC);
  if (value === '') cell.clearContent(); else cell.setValue(value);
  try { ss.toast('Mora de ' + loanId + ': ' + label + '.', 'Mora actualizada', 5); } catch (e) { }
}
/** Menú: condona (pone en 0) la mora del préstamo de la fila activa. */
function condonarMoraFilaActiva() { setMoraAjusteFilaActiva_(0, 'condonada (0)'); }
/** Menú: restaura la mora automática (vacía el ajuste) del préstamo de la fila activa. */
function restaurarMoraFilaActiva() { setMoraAjusteFilaActiva_('', 'automática'); }
/** Menú: fija un recargo por mora a mano para el préstamo de la fila activa (pide el importe). */
function fijarMoraFilaActiva() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Fijar recargo por mora', 'Importe del recargo por mora para el préstamo de la fila activa (0 = condonar):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const n = Number(String(resp.getResponseText()).replace(/[^0-9.,\-]/g, '').replace(',', '.'));
  if (!isFinite(n) || n < 0) { ui.alert('Ingresá un importe válido (≥ 0).'); return; }
  setMoraAjusteFilaActiva_(round2_(n), 'fijada en ' + fmtMoney_(round2_(n)));
}

/** Datos derivados del estado de cuenta (compartidos por el PDF y el correo). */
function statementData_(loan) {
  const pays = loanPayments_(loan.loanId).sort((a, b) => a.date - b.date);
  const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
  // Saldo base (mismo criterio que la hoja "Estudio de Estados": Total a Pagar −
  // Total Pagado). Si está saldado, no hay atraso ni recargo por mora.
  const out = Math.max(0, round2_(loan.totalDue - totalPaid));
  const paid = out <= 0.009;
  const dLate = paid ? 0 : daysLate_(loan.dueDate);
  // Mora POR CUOTA (recargo diario configurable, tras la gracia, con tope % del capital) —
  // consistente con computeOutstanding_. Incluye cuotas vencidas aunque el vencimiento final no llegó.
  const feeAccum = paid ? 0 : accruedMora_(loan);
  const overdue = !paid && (feeAccum > 0 || ((loan.dueDate instanceof Date) && new Date().getTime() > loan.dueDate.getTime()));
  return { pays, out, totalPaid, paid, overdue, dLate, feeAccum };
}

/** PDF del estado de cuenta, guardado en la carpeta del prestatario. */
function makeStatementFile_(loan) {
  return savePdfToBorrower_(statementHtml_(loan), 'Estado de Cuenta ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}

/** HTML del estado de cuenta: banner de estado + detalle + pagos. Mensaje amable si está saldado. */
function statementHtml_(loan) {
  const lender = companyName_(), lPhone = getSetting_('Teléfono del Prestamista'),
    lAddr = getSetting_('Dirección del Prestamista'), footer = getSetting_('Pie del Contrato');
  const d = statementData_(loan);
  let banner, message;
  if (d.paid) {
    banner = { bg: '#d9ead3', color: '#38761d', text: 'PRÉSTAMO SALDADO ✓' };
    message = `<p style="background:#eaf5e6;border-left:4px solid #38761d;padding:12px 16px;border-radius:6px;font-size:12.5pt">
      🎉 <b>¡Felicitaciones, ${esc_(loan.name)}!</b> Su préstamo está <b>totalmente saldado</b>.
      Muchas gracias por su cumplimiento y su confianza — fue un placer acompañarle.
      Esperamos poder ayudarle nuevamente cuando lo necesite.</p>`;
  } else if (d.overdue) {
    banner = { bg: '#f4cccc', color: '#990000', text: 'PAGO VENCIDO — ' + d.dLate + ' DÍA(S) DE ATRASO' };
    message = `<p>Su préstamo presenta un saldo pendiente vencido. Le pedimos regularizar el pago a la brevedad
      para evitar que el recargo por mora siga creciendo.</p>`;
  } else {
    banner = { bg: '#cfe2f3', color: '#1c4587', text: 'PRÉSTAMO ACTIVO' };
    message = `<p>A continuación encontrará el detalle de su préstamo y de los pagos registrados a la fecha.</p>`;
  }
  const cell = 'padding:6px 8px;border:1px solid #ccc';
  let payRows = d.pays.map(p => `<tr><td style="${cell}">${fmtDate_(p.date)}</td><td style="${cell};text-align:right">${fmtMoney_(p.amount)}</td></tr>`).join('');
  if (!payRows) payRows = `<tr><td colspan="2" style="${cell}">— Sin pagos registrados —</td></tr>`;
  const feeRows = d.overdue ?
    `<tr><td class="k">Días de atraso</td><td>${d.dLate}</td></tr>
     <tr><td class="k" style="color:#900">Recargo por mora (${lateFeePctText_()}/día)</td><td style="color:#900"><b>${fmtMoney_(d.feeAccum)}</b></td></tr>` : '';
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;border-bottom:2px solid #674ea7;padding-bottom:8px}
    h2{font-size:13pt;color:#674ea7;margin-top:24px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px} td{padding:6px 8px;border:1px solid #ccc}
    td.k{background:#f3eefb;font-weight:bold;width:45%}
    .banner{margin:16px 0;padding:12px;text-align:center;font-weight:bold;font-size:14pt;border-radius:6px;
      background:${banner.bg};color:${banner.color}}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>ESTADO DE CUENTA</h1>
    <p>Estimado/a <b>${esc_(loan.name)}</b>, este es el estado de cuenta de su préstamo
       <b>${esc_(loan.loanId)}</b> con ${esc_(lender)}, al <b>${fmtDate_(new Date())}</b>.</p>
    <div class="banner">${banner.text}</div>
    ${message}
    <h2>Detalle del préstamo</h2>
    <table>
      <tr><td class="k">ID Préstamo</td><td>${esc_(loan.loanId)}</td></tr>
      <tr><td class="k">DNI</td><td>${esc_(loan.dni)}</td></tr>
      <tr><td class="k">Capital</td><td>${fmtMoney_(loan.principal)}</td></tr>
      <tr><td class="k">Total a pagar</td><td><b>${fmtMoney_(loan.totalDue)}</b></td></tr>
      <tr><td class="k">Fecha de vencimiento</td><td>${fmtDate_(loan.dueDate)}</td></tr>
      <tr><td class="k">Total pagado</td><td><b>${fmtMoney_(d.totalPaid)}</b></td></tr>
      ${feeRows}
      <tr><td class="k">Saldo pendiente</td><td><b>${fmtMoney_(d.out)}</b></td></tr>
    </table>
    <h2>Pagos registrados</h2>
    <table><tr><td class="k" style="width:50%">Fecha</td><td class="k" style="text-align:right">Monto</td></tr>
      ${payRows}
      <tr><td style="${cell}"><b>Total pagado</b></td><td style="${cell};text-align:right"><b>${fmtMoney_(d.totalPaid)}</b></td></tr>
    </table>
    <p style="margin-top:20px">Ante cualquier duda, contáctenos${lPhone ? ' al ' + esc_(lPhone) : ''}.</p>
    <p style="margin-top:14px">${esc_(lender)}${lAddr ? '<br>' + esc_(lAddr) : ''}</p>
    <p class="foot">${esc_(footer)}</p></body></html>`;
}

/** Envía el estado de cuenta con el PDF adjunto. Mensaje amable si el préstamo está saldado. */
function sendStatementEmail_(loan, file) {
  const d = statementData_(loan);
  const cell = 'padding:4px 10px;border:1px solid #ccc';
  let rows = d.pays.map(p => '<tr><td style="' + cell + '">' + fmtDate_(p.date) + '</td><td style="' + cell + ';text-align:right">' + fmtMoney_(p.amount) + '</td></tr>').join('');
  if (!rows) rows = '<tr><td colspan="2" style="' + cell + '">Sin pagos registrados.</td></tr>';
  const feeRows = d.overdue ?
    '<tr><td style="' + cell + '"><b>Días de atraso</b></td><td style="' + cell + ';text-align:right">' + d.dLate + '</td></tr>' +
    '<tr><td style="' + cell + ';color:#900"><b>Recargo por mora (' + lateFeePctText_() + '/día)</b></td><td style="' + cell + ';text-align:right;color:#900"><b>' + fmtMoney_(d.feeAccum) + '</b></td></tr>' : '';
  const intro = d.paid
    ? '<p>Estimado/a ' + esc_(loan.name) + ',</p>' +
      '<p style="background:#eaf5e6;border-left:4px solid #38761d;padding:10px 14px;border-radius:4px">' +
      '🎉 <b>¡Felicitaciones!</b> Su préstamo <b>' + esc_(loan.loanId) + '</b> está <b>totalmente saldado</b>. ' +
      'Muchas gracias por su cumplimiento y su confianza — fue un placer acompañarle. ' +
      'Esperamos poder ayudarle nuevamente cuando lo necesite.</p>'
    : '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Le enviamos el estado de cuenta de su préstamo <b>' + esc_(loan.loanId) + '</b>:</p>';
  const html = intro +
    '<p>Capital ' + fmtMoney_(loan.principal) + ' · Total a pagar ' + fmtMoney_(loan.totalDue) + ' · Vence ' + fmtDate_(loan.dueDate) + '</p>' +
    '<table style="border-collapse:collapse"><tr><th style="' + cell + ';background:#f0f4fa">Fecha</th><th style="' + cell + ';background:#f0f4fa">Monto</th></tr>' +
    rows + '<tr><td style="' + cell + '"><b>Total pagado</b></td><td style="' + cell + ';text-align:right"><b>' + fmtMoney_(d.totalPaid) + '</b></td></tr>' +
    feeRows +
    '<tr><td style="' + cell + '"><b>Saldo pendiente</b></td><td style="' + cell + ';text-align:right"><b>' + fmtMoney_(d.out) + '</b></td></tr></table>' +
    '<p>Adjuntamos el estado de cuenta en PDF.</p>' +
    '<p>Saludos,<br>' + esc_(companyName_()) + '</p>';
  const subject = d.paid ? '¡Préstamo saldado! — ' + loan.loanId : 'Estado de cuenta — ' + loan.loanId;
  const plain = d.paid ? 'Su préstamo ' + loan.loanId + ' está totalmente saldado. ¡Muchas gracias!'
    : 'Estado del préstamo ' + loan.loanId + '. Total pagado ' + fmtMoney_(d.totalPaid) + '. Saldo ' + fmtMoney_(d.out) + '.';
  sendBrandedEmail_(loan.email, subject, plain, html, { attachments: [file.getAs('application/pdf')] });
  return (d.paid ? 'Estado (préstamo saldado) enviado a ' : 'Estado de cuenta enviado a ') + loan.email + '.';
}

function emailStatement(loanId) {
  return guard_('emailStatement', function () {
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan) throw new Error('No se encontró el préstamo.');
    if (!loan.email) throw new Error('El prestatario no tiene correo.');
    return sendStatementEmail_(loan, makeStatementFile_(loan));
  });
}

/* ===================== SIDEBAR (PANEL) ===================== */
function sidebarHtml_() {
  const ctx = activeLoanId_();
  return `<!DOCTYPE html><html><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;margin:0;padding:12px;font-size:13px;color:#222}
    h3{color:#1c4587;margin:14px 0 6px;border-bottom:1px solid #ddd;padding-bottom:3px}
    input{width:100%;box-sizing:border-box;padding:7px;margin:3px 0 8px;border:1px solid #ccc;border-radius:5px}
    button{width:100%;padding:9px;margin:4px 0;border:0;border-radius:5px;color:#fff;font-weight:bold;cursor:pointer}
    .blue{background:#1c4587}.green{background:#38761d}.purple{background:#674ea7}.grey{background:#666}
    #msg{margin-top:10px;padding:8px;border-radius:5px;display:none;white-space:pre-wrap}
    .ok{background:#d9ead3}.err{background:#f4cccc}</style></head><body>
    <h3>Préstamo</h3>
    <label>ID Préstamo</label><input id="loan" value="${esc_(ctx || '')}" placeholder="L-0001">
    <button class="blue" onclick="run('previewAgreement','Generando vista previa…')">👁 Vista previa del contrato</button>
    <button class="blue" onclick="run('sendAgreement','Enviando contrato…')">✉ Enviar contrato por correo</button>
    <button class="green" onclick="run('sendDisbursement','Enviando recibo de desembolso…')">💸 Registrar desembolso + enviar recibo</button>
    <button class="purple" onclick="run('sendStatement','Enviando estado…')">✉ Enviar estado de cuenta</button>
    <h3>Registrar pago</h3>
    <label>Monto (ARS)</label><input id="amount" type="number" min="1" step="any">
    <label>Fecha (opcional)</label><input id="date" type="date">
    <button class="green" onclick="pay()">💵 Registrar pago</button>
    <h3>General</h3>
    <button class="grey" onclick="run('refresh','Actualizando…')">🔄 Actualizar saldos y resumen</button>
    <button class="green" onclick="run('moveCleared','Moviendo saldados…')">✔ Mover saldados → hoja Saldados</button>
    <h3>Pestañas</h3>
    <button class="blue" onclick="run('dailyOnly','Ajustando pestañas…')">🗂 Ver solo operaciones diarias</button>
    <button class="grey" onclick="run('allSheets','Mostrando hojas…')">👁 Mostrar todas las hojas</button>
    <div id="msg"></div>
    <script>
      var msg=document.getElementById('msg');
      function show(t,ok){msg.style.display='block';msg.className=ok?'ok':'err';msg.textContent=t;}
      function loanId(){return document.getElementById('loan').value.trim();}
      function run(kind,busy){
        show(busy,true);
        var f={previewAgreement:'sbPreviewAgreement',sendAgreement:'sbSendAgreement',sendDisbursement:'sbSendDisbursementReceipt',sendStatement:'sbSendStatement',refresh:'sbRefresh',moveCleared:'sbMoveCleared',dailyOnly:'sbDailyOnly',allSheets:'sbAllSheets'}[kind];
        google.script.run.withSuccessHandler(function(r){show(r,true);}).withFailureHandler(function(e){show(e.message||e,false);})[f](loanId());
      }
      function pay(){
        show('Registrando pago…',true);
        google.script.run.withSuccessHandler(function(r){show(r,true);}).withFailureHandler(function(e){show(e.message||e,false);})
          .recordPayment(loanId(),document.getElementById('amount').value,document.getElementById('date').value);
      }
    </script></body></html>`;
}
function activeLoanId_() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== CFG.SHEETS.BORROWERS) return '';
    const row = sh.getActiveRange().getRow();
    return row >= 2 ? String(sh.getRange(row, 1).getValue()).trim() : '';
  } catch (e) { return ''; }
}
function sbPreviewAgreement(loanId) {
  return guard_('sbPreviewAgreement', function () {
    const loan = findLoanById_(String(loanId).trim());
    if (!loan) throw new Error('No se encontró el préstamo.');
    const file = makeAgreementFile_(loan);
    return 'Vista previa creada. Ábrala aquí:\n' + file.getUrl();
  });
}
function sbSendAgreement(loanId) {
  return guard_('sbSendAgreement', function () {
    const loan = findLoanById_(String(loanId).trim());
    if (!loan) throw new Error('No se encontró el préstamo.');
    if (!loan.email) throw new Error('El prestatario no tiene correo.');
    const file = makeAgreementFile_(loan); emailAgreement_(loan, file);
    markAgreementSent_(loan.loanId, file);
    return 'Contrato enviado a ' + loan.email + '.';
  });
}
function sbSendStatement(loanId) { return emailStatement(loanId); }
function sbRefresh() { refreshAll(true); return 'Saldos y resumen actualizados.'; }
function sbMoveCleared() {
  const n = doMoveCleared_();
  return n ? ('Se movieron ' + n + ' préstamo(s) saldado(s) a la hoja "Saldados".')
    : 'No hay préstamos saldados (pagados en su totalidad) para mover.';
}
function sbDailyOnly() {
  const n = applyDailyVisibility_(getSS_());
  return 'Vista de operaciones diarias: ' + n + ' hoja(s) secundaria(s) oculta(s).';
}
function sbAllSheets() {
  const ss = getSS_(); let shown = 0;
  ss.getSheets().forEach(function (sh) { try { if (sh.isSheetHidden()) shown++; sh.showSheet(); } catch (e) {} });
  return shown + ' hoja(s) mostrada(s). Todas las pestañas están visibles.';
}

function markAgreementSent_(loanId, file) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  const ids = bs.getRange(2, 1, Math.max(last - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) {
    bs.getRange(i + 2, PB.PDF).setFormula('=HYPERLINK("' + file.getUrl() + '","Ver contrato")');
    bs.getRange(i + 2, PB.SENT).setValue(new Date()); return;
  }
}

/* ===================== ACTUALIZAR ===================== */
function refreshAll(silent) {
  const ss = getSS_(), N = CFG.MAX_ROWS;
  cleanupLegacySheets_(ss);
  updateAllOutstanding_();
  try { refreshPaymentNames_(); } catch (e) { logError_('refreshAll:refreshPaymentNames_', e); }
  setupSummary_(ss); setupPanel_(ss); setupStats_(ss); setupStatements_(ss); rebuildLateSheet_(ss); rebuildRemindersSheet_(ss);
  SpreadsheetApp.flush();
  if (!silent) SpreadsheetApp.getUi().alert('Actualizado. Saldos, resumen, panel, estadísticas y pagos atrasados recalculados.');
}
function cleanupLegacySheets_(ss) {
  ss.getSheets().forEach(s => {
    const n = s.getName();
    if (/^New Borrowers \(old/i.test(n) || n === 'New Borrowers (pre-form)') { try { ss.deleteSheet(s); } catch (e) {} }
  });
}

/* ===================== FORMULARIO WEB ===================== */
function showWebFormLink() {
  setupNew_(getSS_());
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  if (url) {
    PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', url);
    const cfg = String(getSetting_('URL de la app web (enlaces a clientes)') || '').trim();
    SpreadsheetApp.getUi().alert('Formulario de solicitud (compartir con solicitantes):\n' + url +
      '\n\nConsulta de saldo (para prestatarios):\n' + url + '?page=saldo' +
      '\n\nIMPORTANTE: para que los enlaces de FIRMA que se envían a los clientes usen la URL correcta, ' +
      'pegue la URL de su implementación (la que termina en /exec) en:\n' +
      'Configuración ▸ "URL de la app web (enlaces a clientes)".' +
      (cfg ? '\n\nActualmente configurada:\n' + cfg : '\n\n(Actualmente vacía — se usa la URL detectada automáticamente.)'));
  } else {
    SpreadsheetApp.getUi().alert('Publique el formulario (una vez):\n\n1. Extensions ▸ Apps Script.\n' +
      '2. Deploy ▸ Nueva implementación ▸ ⚙ ▸ Aplicación web.\n3. Ejecutar como: Yo.  Acceso: Cualquiera.\n' +
      '4. Implementar ▸ autorizar ▸ copiar la URL.\n\nLuego ejecute ⑥ otra vez para ver el enlace.');
  }
}

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'saldo') return HtmlService.createHtmlOutput(balancePageHtml_()).setTitle('Consulta de Saldo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (page === 'firmar') return HtmlService.createHtmlOutput(signingPageHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Firmar Contrato').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  // Formulario clásico disponible en ?page=clasico; por defecto se sirve el
  // formulario inteligente (detecta clientes existentes por Correo + DNI).
  if (page === 'clasico') return HtmlService.createHtmlOutput(acceptingApplications_() ? intakeHtml_() : closedHtml_())
    .setTitle('Solicitud de Préstamo').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  const html = acceptingApplications_() ? intakeSmartHtml_() : closedHtml_();
  return HtmlService.createHtmlOutput(html).setTitle('Solicitud de Préstamo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Página que se muestra cuando las solicitudes están cerradas. */
function closedHtml_() {
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:28px;box-shadow:0 2px 12px rgba(0,0,0,.08);text-align:center}
    h1{font-size:20px;color:#1c4587;margin:8px 0}
    .badge{display:inline-block;background:#fce8e6;color:#a50e0e;font-weight:bold;padding:8px 16px;border-radius:999px;margin:8px 0 4px}
    p{color:#555;line-height:1.5}</style></head><body>
    <div class="card">${brandHeaderHtml_()}
      <div class="badge">Solicitudes cerradas</div>
      <h1>Por el momento no hay préstamos disponibles</h1>
      <p>Gracias por su interés. En este momento <b>no estamos aceptando nuevas solicitudes de préstamo</b>.
         Le invitamos a intentarlo nuevamente más adelante.</p>
      <p style="font-size:13px;color:#888">${esc_(companyName_())}</p>
    </div></body></html>`;
}

function intakeHtml_() {
  const lender = esc_(companyName_());
  const feePct = lateFeePctText_();
  const lateFeePct = Number(lateFeeRate_() * 100) || 5; // % de mora diario (para el cálculo en vivo)
  const juris = esc_(getSetting_('Jurisdicción') || 'Buenos Aires, Argentina');
  const moraClause = esc_(getSetting_('Cláusula de Mora') ||
    ('En caso de mora se aplica un recargo del ' + feePct + ' diario sobre el total a devolver por cada día de atraso posterior a la fecha de vencimiento' + (moraGraceDays_() > 0 ? ', tras ' + moraGraceDays_() + ' día(s) de gracia' : '') + ', con un tope acumulado del ' + moraCapPctText_() + ' del capital.'));
  const contractFooter = esc_(getSetting_('Pie del Contrato') ||
    'Este acuerdo es legalmente vinculante desde la firma de ambas partes.');
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#1c4587;margin:0 0 4px}.sub{color:#666;margin:0 0 18px}
    label{display:block;font-weight:bold;margin:14px 0 4px;font-size:14px}.req{color:#c00}
    input,textarea,select{width:100%;box-sizing:border-box;padding:9px 10px;border:1px solid #ccc;border-radius:6px;font-size:14px}
    .hint{font-weight:normal;color:#888;font-size:12px}
    #calc{margin-top:8px;padding:10px;background:#f0f4fa;border-radius:6px;font-size:13px;display:none}
    .amount-words{display:none;margin-top:5px;font-size:12.5px;font-style:italic;color:#1c4587;line-height:1.35}
    button{margin-top:22px;width:100%;background:#1c4587;color:#fff;border:0;border-radius:6px;padding:12px;font-size:15px;font-weight:bold;cursor:pointer}
    button:disabled{background:#9db4d6}
    #msg{margin-top:16px;padding:12px;border-radius:6px;display:none}.ok{background:#d9ead3;color:#274e13}.err{background:#f4cccc;color:#900}
    /* Cargando */
    #loadingView{display:none;text-align:center;padding:34px 0}
    .spinner{width:56px;height:56px;margin:0 auto 18px;border:6px solid #e3e9f4;border-top-color:#1c4587;border-radius:50%;animation:spin .9s linear infinite}
    @keyframes spin{to{transform:rotate(360deg)}}
    /* Éxito */
    #successView{display:none}
    .check{width:72px;height:72px;line-height:72px;margin:6px auto 10px;border-radius:50%;background:#38761d;color:#fff;font-size:40px;font-weight:bold;text-align:center}
    .summary{width:100%;border-collapse:collapse;margin:18px 0}
    .summary td{padding:9px 12px;border:1px solid #e2e6ee;font-size:14px}
    .summary td.k{background:#f0f4fa;font-weight:bold;width:46%}
    .note{background:#e8f0fe;border-radius:8px;padding:12px 14px;font-size:13px;color:#274e6b;margin-top:6px}
    /* Términos y Condiciones */
    .terms{margin-top:18px;border:1px solid #ccd3e0;border-radius:8px;background:#fafbfe}
    .terms h2{font-size:14px;color:#1c4587;margin:0;padding:10px 12px;border-bottom:1px solid #e2e6ee;background:#f0f4fa;border-radius:8px 8px 0 0}
    .terms .body{max-height:160px;overflow-y:auto;padding:8px 14px;font-size:12.5px;color:#444;line-height:1.5}
    .terms ul{margin:6px 0 0;padding-left:18px}.terms li{margin:5px 0}
    .terms li.mora{color:#7a5200}.terms li.mora b{color:#b45309}
    .agree{display:flex;align-items:flex-start;gap:9px;margin:14px 0 0;font-weight:normal;font-size:13.5px;cursor:pointer}
    .agree input{width:18px;height:18px;flex:0 0 auto;margin-top:1px}</style></head><body>
    <div class="card" id="card">${brandHeaderHtml_()}
    <div id="formView">
      <h1>Solicitud de Préstamo</h1>
      <p class="sub">${lender} — todos los campos son obligatorios. Su solicitud será revisada y verificada.</p>
      <form id="f">
        <label>Nombre completo <span class="req">*</span><input name="fullName" required></label>
        <label>Correo electrónico <span class="req">*</span><input name="email" type="email" required></label>
        <label>DNI <span class="req">*</span><input name="dni" required placeholder="XX.XXX.XXX"></label>
        <label>Teléfono <span class="req">*</span><input name="phone" required></label>
        <label>Monto solicitado (ARS) <span class="req">*</span><input name="amount" id="amount" type="text" inputmode="numeric" autocomplete="off" placeholder="Ej: 100.000" required>
          <span id="amountWords" class="amount-words" aria-live="polite"></span></label>
        <label>Plazo <span class="req">*</span><select name="term" id="term" required>
          <option value="">— elegir —</option><option value="15">15 días (25% de interés)</option><option value="1">1 mes (50% de interés)</option><option value="2">2 meses (100% de interés)</option></select></label>
        <div id="calc"></div>
        <label>Notas / Motivo <span class="req">*</span><textarea name="notes" rows="2" required></textarea></label>
        <label>Forma de pago preferida <span class="req">*</span><select name="payMethod" id="payMethod" required>
          <option value="Mercado Pago" selected>Mercado Pago</option>
          <option value="Transferencia bancaria">Transferencia bancaria</option>
          <option value="Efectivo">Efectivo</option>
          <option value="Otro">Otro</option></select></label>
        <label id="payDetailsLabel">Datos de la cuenta / alias <span class="hint">— alias, CBU/CVU o el medio que prefiera (opcional)</span>
          <input name="payDetails" id="payDetails" type="text" autocomplete="off" placeholder="Ej: alias.mercadopago o CBU"></label>
        <label>Foto del frente del DNI <span class="req">*</span> <span class="hint">— imagen o PDF</span><input name="dniPhoto" type="file" accept="image/*,.pdf" capture="environment" required></label>
        <label>Foto del dorso del DNI <span class="req">*</span> <span class="hint">— imagen o PDF</span><input name="cuilPhoto" type="file" accept="image/*,.pdf" capture="environment" required></label>
        <div class="terms">
          <h2>Términos y Condiciones del Préstamo</h2>
          <div class="body">
            <p>Al enviar esta solicitud usted declara haber leído, comprendido y aceptado los siguientes términos con ${lender}:</p>
            <ul>
              <li><b>Intereses según el plazo:</b> 15 días = 25%, 1 mes = 50%, 2 meses = 100% sobre el capital solicitado.</li>
              <li><b>Devolución:</b> el capital más los intereses se devuelven en su totalidad en la fecha de vencimiento.</li>
              <li class="mora">⚠️ <b>Recargo por mora:</b> ${moraClause}</li>
              <li><b>Verificación:</b> la solicitud está sujeta a revisión y verificación de identidad (DNI y CUIL). La aprobación no está garantizada.</li>
              <li><b>Declaración:</b> declaro que los datos y documentos aportados son verídicos y de mi titularidad.</li>
              <li><b>Jurisdicción:</b> ${juris}.</li>
              <li>${contractFooter}</li>
            </ul>
          </div>
        </div>
        <label class="agree"><input type="checkbox" name="agree" id="agree" required>
          <span>He leído y acepto los <b>Términos y Condiciones</b> del préstamo. <span class="req">*</span></span></label>
        <button type="submit" id="btn">Enviar solicitud</button></form>
      <div id="msg"></div>
    </div>
    <div id="loadingView">
      <div class="spinner"></div>
      <p style="font-weight:bold;color:#1c4587;margin:0">Enviando su solicitud…</p>
      <p style="color:#888;font-size:13px;margin:6px 0 0">Subiendo sus documentos. Esto puede tardar unos segundos — por favor no cierre esta ventana.</p>
    </div>
    <div id="successView"></div>
    </div>
    <script>
      var f=document.getElementById('f'),btn=document.getElementById('btn'),msg=document.getElementById('msg'),calc=document.getElementById('calc');
      var formView=document.getElementById('formView'),loadingView=document.getElementById('loadingView'),successView=document.getElementById('successView');
      function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
      // Mensajes de validación en español, sin importar el idioma del navegador.
      Array.prototype.forEach.call(f.elements,function(el){
        if(!el.name)return;
        el.addEventListener('invalid',function(){
          if(el.validity.valueMissing) el.setCustomValidity('Este campo es obligatorio.');
          else if(el.validity.typeMismatch) el.setCustomValidity('Ingrese un '+(el.type==='email'?'correo electrónico válido':'valor válido')+'.');
          else if(el.validity.rangeUnderflow||el.validity.badInput) el.setCustomValidity('Ingrese un monto válido.');
          else el.setCustomValidity('');
        });
        el.addEventListener('input',function(){el.setCustomValidity('');});
        el.addEventListener('change',function(){el.setCustomValidity('');});
      });
      // Mensaje específico para la casilla de Términos y Condiciones.
      var agree=document.getElementById('agree');
      agree.addEventListener('invalid',function(){if(agree.validity.valueMissing)agree.setCustomValidity('Debe aceptar los Términos y Condiciones para enviar la solicitud.');});
      agree.addEventListener('change',function(){agree.setCustomValidity('');});
      function fmt(n){try{return n.toLocaleString('es-AR',{style:'currency',currency:'ARS'});}catch(e){return '$'+n.toFixed(2);}}
      var amountEl=document.getElementById('amount'),amountWords=document.getElementById('amountWords');
      // Convierte solo dígitos a valor numérico (los puntos son separadores de miles).
      function parseAmount(v){var d=String(v==null?'':v).replace(/[^0-9]/g,'');return d?parseInt(d,10):0;}
      // Agrupa los dígitos con puntos de miles al estilo argentino: 1000000 -> 1.000.000
      // (agrupación manual, sin regex con backslash: se pierden dentro de la plantilla de texto).
      function groupNum(d){var out='',c=0;for(var i=d.length-1;i>=0;i--){out=d.charAt(i)+out;c++;if(c%3===0&&i>0){out='.'+out;}}return out;}
      // Convierte un entero a su expresión en palabras (español).
      function numeroALetras(num){
        num=Math.floor(Math.abs(num));if(num===0)return 'cero';
        var U=['','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez','once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve','veinte','veintiuno','veintidós','veintitrés','veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve'];
        var D=['','','','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'];
        var C=['','ciento','doscientos','trescientos','cuatrocientos','quinientos','seiscientos','setecientos','ochocientos','novecientos'];
        function seccion(n){ // n: 0..999, apocopa uno->un (siempre precede a un sustantivo)
          if(n===0)return '';if(n===100)return 'cien';
          var c=Math.floor(n/100),r=n%100,t=c>0?C[c]:'';
          if(r>0){if(t)t+=' ';
            if(r<30){t+=(r===1)?'un':(r===21)?'veintiún':U[r];}
            else{var dd=Math.floor(r/10),u=r%10;t+=D[dd];if(u>0)t+=' y '+((u===1)?'un':U[u]);}}
          return t;
        }
        var p=[],mill=Math.floor(num/1000000),mil=Math.floor((num%1000000)/1000),r=num%1000;
        if(mill>0)p.push(mill===1?'un millón':seccion(mill)+' millones');
        if(mil>0)p.push(mil===1?'mil':seccion(mil)+' mil');
        if(r>0)p.push(seccion(r));
        return p.join(' ').trim();
      }
      function capitalizar(s){return s?s.charAt(0).toUpperCase()+s.slice(1):s;}
      // Reformatea el campo mientras se escribe, preservando la posición del cursor, y muestra el monto en palabras.
      function formatAmountInput(){
        var start=amountEl.selectionStart||0;
        var digitsBefore=amountEl.value.slice(0,start).replace(/[^0-9]/g,'').length;
        var digits=amountEl.value.replace(/[^0-9]/g,'').replace(/^0+/,'');
        var formatted=digits?groupNum(digits):'';
        amountEl.value=formatted;
        var pos=formatted.length,count=0;
        if(digitsBefore>0){for(var i=0;i<formatted.length;i++){if(formatted.charAt(i)!=='.')count++;if(count>=digitsBefore){pos=i+1;break;}}}
        else pos=0;
        try{amountEl.setSelectionRange(pos,pos);}catch(e){}
        var n=digits?parseInt(digits,10):0;
        if(n>0){var w=numeroALetras(n),noun=/(millón|millones)$/.test(w)?' de pesos':(n===1?' peso':' pesos');
          amountWords.textContent=capitalizar(w)+noun;amountWords.style.display='block';}
        else{amountWords.textContent='';amountWords.style.display='none';}
      }
      function recalc(){var a=parseAmount(amountEl.value),t=document.getElementById('term').value;
        if(a>0&&(t==='15'||t==='1'||t==='2')){var r=t==='15'?0.25:t==='1'?0.5:1,total=a*(1+r),feePct=${lateFeePct},feeDay=total*feePct/100;calc.style.display='block';
          calc.innerHTML='Interés: <b>'+fmt(a*r)+'</b><br>Total a devolver: <b>'+fmt(total)+'</b>'+
            '<div style="margin-top:8px;color:#7a5200">⚠️ <b style="color:#b45309">Si pagás tarde</b> ('+feePct+'% por día):<br>1 día → <b>'+fmt(total+feeDay)+'</b><br>3 días → <b>'+fmt(total+feeDay*3)+'</b><br>7 días → <b>'+fmt(total+feeDay*7)+'</b></div>';}else{calc.style.display='none';}}
      amountEl.addEventListener('input',function(){formatAmountInput();recalc();});
      document.getElementById('term').addEventListener('change',recalc);
      function onSuccess(r){
        loadingView.style.display='none';
        if(!r||!r.ok){ // respaldo por si el servidor devolviera texto
          successView.innerHTML='<div style="text-align:center"><div class="check">✓</div><h1 style="color:#38761d">¡Solicitud recibida!</h1><p class="sub">'+esc(r&&r.message?r.message:r)+'</p></div>';
          successView.style.display='block';return;
        }
        successView.innerHTML=
          '<div style="text-align:center"><div class="check">&#10003;</div>'+
          '<h1 style="color:#38761d">¡Solicitud recibida!</h1>'+
          '<p class="sub">Gracias, <b>'+esc(r.name)+'</b>. Su solicitud fue registrada correctamente y está pendiente de revisión.</p></div>'+
          '<table class="summary">'+
          '<tr><td class="k">Solicitante</td><td>'+esc(r.name)+'</td></tr>'+
          '<tr><td class="k">Correo</td><td>'+esc(r.email)+'</td></tr>'+
          '<tr><td class="k">DNI</td><td>'+esc(r.dni)+'</td></tr>'+
          '<tr><td class="k">Teléfono</td><td>'+esc(r.phone)+'</td></tr>'+
          '<tr><td class="k">Monto solicitado</td><td>'+esc(r.amountFmt)+'</td></tr>'+
          '<tr><td class="k">Plazo</td><td>'+esc(r.termLabel)+' ('+esc(r.ratePct)+' de interés)</td></tr>'+
          '<tr><td class="k">Interés</td><td>'+esc(r.interestFmt)+'</td></tr>'+
          '<tr><td class="k">Total a devolver</td><td><b>'+esc(r.totalFmt)+'</b></td></tr>'+
          '</table>'+
          '<div class="note">Enviamos un correo de confirmación a <b>'+esc(r.email)+'</b>. Revisaremos su solicitud y nos pondremos en contacto a la brevedad. Puede cerrar esta ventana.</div>';
        successView.style.display='block';
      }
      function onFail(err){
        loadingView.style.display='none';
        formView.style.display='block';
        btn.disabled=false;btn.textContent='Enviar solicitud';
        msg.className='err';msg.style.display='block';msg.textContent=(err.message||err);
      }
      f.addEventListener('submit',function(e){
        e.preventDefault();
        msg.style.display='none';
        btn.disabled=true;btn.textContent='Subiendo…';
        formView.style.display='none';
        loadingView.style.display='block';
        google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFail).submitIntake(f);
      });
    </script></body></html>`;
}

function submitIntake(form) {
  return guard_('submitIntake', function () {
    if (!acceptingApplications_()) throw new Error('En este momento no estamos aceptando nuevas solicitudes de préstamo.');
    const ss = getSS_(), nb = ss.getSheetByName(CFG.SHEETS.NEW) || setupNew_(ss) || ss.getSheetByName(CFG.SHEETS.NEW);
    const name = String(form.fullName || '').trim(), email = String(form.email || '').trim(),
      dni = String(form.dni || '').trim(), phone = String(form.phone || '').trim(),
      amount = Number(String(form.amount || '').replace(/\D/g, '')) || '',
      term = Number(form.term) || '', notes = String(form.notes || '').trim();
    // Forma de pago elegida por el prestatario (predeterminada: Mercado Pago) + datos opcionales.
    const payMethodSel = String(form.payMethod || '').trim() || 'Mercado Pago';
    const payDetails = String(form.payDetails || '').trim();
    const payMethodFull = payDetails ? (payMethodSel + ' — ' + payDetails) : payMethodSel;
    if (!name) throw new Error('El nombre completo es obligatorio.');
    if (!email) throw new Error('El correo electrónico es obligatorio.');
    if (!dni) throw new Error('El DNI es obligatorio.');
    if (!phone) throw new Error('El teléfono es obligatorio.');
    if (!amount) throw new Error('Ingrese un monto de préstamo válido.');
    if (term !== 1 && term !== 2 && term !== 15) throw new Error('Elija un plazo de 15 días, 1 mes o 2 meses.');
    if (!notes) throw new Error('Las notas / motivo son obligatorias.');
    if (!form.agree) throw new Error('Debe aceptar los Términos y Condiciones para enviar la solicitud.');
    if (!hasFile_(form.dniPhoto)) throw new Error('Adjunte la foto del DNI.');
    if (!hasFile_(form.cuilPhoto)) throw new Error('Adjunte la foto del CUIL.');
    // V-33 — cliente bloqueado (correo/DNI/teléfono): también en el formulario clásico.
    if (typeof findClienteBloqueado_ === 'function' && findClienteBloqueado_({ email: email, dni: dni, phone: phone }))
      throw new Error(typeof BLOQUEO_MSG_CLIENTE !== 'undefined' ? BLOQUEO_MSG_CLIENTE : 'No es posible procesar solicitudes para este cliente.');
    const rate = termRate_(term), interest = round2_(amount * rate), total = round2_(amount * (1 + rate));
    const folder = borrowerFolder_(borrowerFolderName_(name, dni));
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const dniUrl = savePhoto_(folder, form.dniPhoto, name + ' DNI ' + stamp);
    const cuilUrl = savePhoto_(folder, form.cuilPhoto, name + ' CUIL ' + stamp);
    const row = nbAppend_(nb, [new Date(), name, email, dni, phone, amount, term, notes, dniUrl, cuilUrl, false, false, 'Pendiente de verificación']);
    nb.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm'); nb.getRange(row, 6).setNumberFormat(CFG.CURRENCY_FMT);
    // Forma de pago: se escribe por nombre de encabezado (columna al final de NB).
    try { setSheetFieldByHeader_(nb, row, 'Forma de Pago', payMethodFull); } catch (e) { logError_('submitIntake:payMethod', e); }
    try {
      sendBrandedEmail_(email, 'Solicitud recibida',
        'Estimado/a ' + name + ', recibimos su solicitud de préstamo por ' + fmtMoney_(amount) + '. La revisaremos y nos pondremos en contacto.',
        '<p>Estimado/a ' + esc_(name) + ',</p><p>Recibimos su solicitud de préstamo por <b>' + fmtMoney_(amount) +
        '</b> a ' + termLabel_(term) + '. La revisaremos y nos pondremos en contacto a la brevedad.</p>');
    } catch (e) { logError_('submitIntake:email', e); }
    // Aviso al prestamista (Correo del Prestamista) con los datos de la nueva solicitud.
    try {
      const recipients = noticeRecipients_();
      if (recipients.length) {
        const avail = fundStats_().net; // neto real: puede ser negativo (sobregiro)
        sendBrandedEmail_(recipients.join(','), 'Nueva solicitud de préstamo — ' + name,
          'Nueva solicitud de ' + name + ' por ' + fmtMoney_(amount) + ' a ' + termLabel_(term) + '. Efectivo disponible: ' + fmtMoney_(avail) + '.',
          '<p>Se recibió una <b>nueva solicitud de préstamo</b>.</p>' +
          '<table style="border-collapse:collapse">' +
          row_('Nombre completo', esc_(name)) + row_('Correo', esc_(email)) + row_('DNI', esc_(dni)) +
          row_('Teléfono', esc_(phone)) + row_('Monto solicitado', '<b>' + fmtMoney_(amount) + '</b>') +
          row_('Plazo', termLabel_(term) + ' (' + termRatePct_(term) + ')') +
          row_('Interés', fmtMoney_(interest)) + row_('Total a devolver', fmtMoney_(total)) +
          row_('Notas / Motivo', esc_(notes)) +
          row_('Forma de pago preferida', esc_(payMethodFull)) +
          row_('Foto del frente del DNI', dniUrl ? '<a href="' + dniUrl + '">Ver documento</a>' : '—') +
          row_('Foto del dorso del DNI', cuilUrl ? '<a href="' + cuilUrl + '">Ver documento</a>' : '—') +
          row_('Efectivo disponible para prestar', (avail < amount ? '<b style="color:#900">' : '<b>') + fmtMoney_(avail) + '</b>' + (avail < amount ? ' ⚠ fondo insuficiente' : '')) +
          '</table>' +
          '<p>Revísela en la hoja <b>Nuevos Prestatarios</b> y tilde <b>Verificado?</b> para aprobar o <b>Rechazar?</b> para archivar.</p>');
      }
    } catch (e) { logError_('submitIntake:lenderEmail', e); }
    return {
      ok: true,
      name: name,
      email: email,
      dni: dni,
      phone: phone,
      term: term,
      termLabel: termLabel_(term),
      ratePct: termRatePct_(term),
      amountFmt: fmtMoney_(amount),
      interestFmt: fmtMoney_(interest),
      totalFmt: fmtMoney_(total),
      lenderName: companyName_(),
      message: '¡Solicitud recibida! Gracias, ' + name + '. Le enviamos un correo de confirmación.',
    };
  });
}

function balancePageHtml_() {
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px}
    .card{max-width:460px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#674ea7}label{font-weight:bold;display:block;margin:12px 0 4px}
    input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px}
    button{margin-top:16px;width:100%;background:#674ea7;color:#fff;border:0;border-radius:6px;padding:12px;font-weight:bold;cursor:pointer}
    #out{margin-top:16px;padding:12px;border-radius:6px;display:none}.ok{background:#e9e2f5}.err{background:#f4cccc}</style></head><body>
    <div class="card">${brandHeaderHtml_()}<h1>Consulta de Saldo</h1>
    <label>ID Préstamo</label><input id="loan" placeholder="L-0001">
    <label>DNI</label><input id="dni">
    <button onclick="go()">Consultar</button><div id="out"></div>
    <script>
      var out=document.getElementById('out');
      function go(){out.style.display='block';out.className='ok';out.textContent='Consultando…';
        google.script.run.withSuccessHandler(function(r){out.className=r.ok?'ok':'err';out.innerHTML=r.html;})
          .withFailureHandler(function(e){out.className='err';out.textContent=(e.message||e);})
          .lookupBalance(document.getElementById('loan').value,document.getElementById('dni').value);}
    </script></div></body></html>`;
}
function lookupBalance(loanId, dni) {
  return guard_('lookupBalance', function () {
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, html: 'Datos no encontrados. Verifique el ID y el DNI.' };
    const rate = termRate_(loan.term);
    const out = computeOutstanding_(loan.principal, rate, loan.loanDate, loanPayments_(loan.loanId), new Date(), new Date(9999, 0, 1), undefined, undefined, undefined, loanSchedule_(loan));
    return {
      ok: true, html: '<b>' + esc_(loan.name) + '</b><br>Préstamo: ' + esc_(loan.loanId) +
        '<br>Total a pagar: ' + fmtMoney_(loan.totalDue) + '<br>Vencimiento: ' + fmtDate_(loan.dueDate) +
        '<br><b>Saldo pendiente: ' + fmtMoney_(out) + '</b>'
    };
  });
}

/* ===================== FIRMA VIRTUAL DEL CONTRATO ===================== */
/** Ubica la fila de un préstamo en "Prestatarios". Devuelve {sh,row} o null. */
function signStatusRow_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh.getLastRow();
  if (last < 2) return null;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) return { sh: sh, row: i + 2 };
  return null;
}

/** Resumen de términos del contrato (HTML seguro) para la página de firma. */
/** Verifica identidad (ID + DNI) y devuelve el contrato completo a firmar, o el estado "ya firmado". */
function getContractForSigning(loanId, dni) {
  return guard_('getContractForSigning', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, msg: 'Datos no encontrados. Verifique el ID del préstamo y su DNI.' };
    const loc = signStatusRow_(loanId);
    const status = loc ? String(loc.sh.getRange(loc.row, COL_SIGN_STATUS).getValue() || '').toUpperCase() : '';
    if (status === SIGN.SIGNED)
      return { ok: true, signed: true, msg: 'Este contrato ya fue firmado. Le enviamos una copia por correo.' };
    return { ok: true, signed: false, name: loan.name, contractHtml: agreementHtml_(loan) };
  });
}

/** Recibe la firma (dibujada o imagen), genera el PDF firmado, guarda y avisa a ambas partes. */
/** SHA-256 en hexadecimal de una cadena o de un arreglo de bytes. */
function hexDigest_(input) {
  const bytes = (typeof input === 'string') ? Utilities.newBlob(input).getBytes() : input;
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function submitSignature(loanId, dni, signatureDataUrl, agree, meta) {
  return guard_('submitSignature', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      throw new Error('Datos no encontrados. Verifique el ID del préstamo y su DNI.');
    if (agree !== true) throw new Error('Debe aceptar los términos del contrato para firmar.');
    const m = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i.exec(String(signatureDataUrl || ''));
    if (!m) throw new Error('Falta la firma. Dibuje su firma o cargue una imagen.');
    const loc = signStatusRow_(loanId);
    if (!loc) throw new Error('No se encontró el préstamo.');
    if (String(loc.sh.getRange(loc.row, COL_SIGN_STATUS).getValue() || '').toUpperCase() === SIGN.SIGNED)
      return { ok: true, message: 'Este contrato ya estaba firmado. Le enviamos una copia por correo.' };

    const now = new Date();
    const mime = m[1], ext = mime.split('/')[1].replace('+xml', '').replace('jpeg', 'jpg').replace('svg', 'png');
    const bytes = Utilities.base64Decode(m[2]);
    const method = (meta && meta.method === 'upload') ? 'upload' : 'draw';
    const ip = String((meta && meta.ip) || '').slice(0, 60);
    const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
    folder.createFile(Utilities.newBlob(bytes, mime, 'Firma ' + loanId + '.' + ext));

    // El reloj de interés arranca al firmar (Fecha del Préstamo = hoy).
    loc.sh.getRange(loc.row, PB.LOAN_DATE).setValue(now).setNumberFormat('yyyy-mm-dd');
    writeOutstandingRow_(loc.sh, loc.row);
    const freshLoan = readLoan_(loc.sh, loc.row);

    const sig = {
      dataUrl: signatureDataUrl, signedAt: now, method: method,
      userAgent: (meta && meta.userAgent) || '', ip: ip, hash: '',
    };
    const signedFile = makeSignedAgreementFile_(freshLoan, sig);
    // Hash SHA-256 del PDF firmado (huella de integridad del documento).
    let hashHex = '';
    try { hashHex = hexDigest_(signedFile.getAs('application/pdf').getBytes()); } catch (e) { logError_('submitSignature:hash', e); }

    loc.sh.getRange(loc.row, COL_SIGN_STATUS).setValue(SIGN.SIGNED);
    loc.sh.getRange(loc.row, COL_SIGN_DATE).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
    loc.sh.getRange(loc.row, COL_SIGN_PDF).setFormula('=HYPERLINK("' + signedFile.getUrl() + '","Ver contrato firmado")');

    try {
      setupFirmas_(getSS_()).appendRow([now, loanId, loan.name, loan.dni, loan.email,
        method === 'upload' ? 'Imagen cargada' : 'Dibujada en pantalla',
        String((meta && meta.userAgent) || '').slice(0, 250), signedFile.getUrl(),
        ip || 'N/D', hashHex, freshLoan.cuil || '', freshLoan.direccion || '']);
    } catch (e) { logError_('submitSignature:firmas', e); }
    try { emailSignedContract_(freshLoan, signedFile); } catch (e) { logError_('submitSignature:email', e); }
    try { getSS_().toast('Contrato firmado: ' + loan.name + ' → ' + loanId, 'Firma recibida ✍', 6); } catch (e) {}
    return { ok: true, message: '¡Contrato firmado! Gracias, ' + loan.name + '. Le enviamos una copia por correo.' };
  });
}

/** Enlace personal de firma para un préstamo. */
/**
 * URL base de la app web para los enlaces que se envían a clientes.
 * Prioridad: 1) "URL de la app web (enlaces a clientes)" en Configuración,
 * 2) propiedad WEBAPP_URL (menú ⑥), 3) ScriptApp.getService().getUrl().
 * Se limpian parámetros/fragmentos para poder anexar los propios.
 */
function webAppBaseUrl_() {
  let url = String(getSetting_('URL de la app web (enlaces a clientes)') || '').trim();
  if (!url) url = String(PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '');
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; } }
  if (!url) return '';
  url = url.replace(/[?#].*$/, '');                       // quita parámetros/fragmentos
  url = url.replace(/\/macros\/u\/\d+\/s\//, '/macros/s/'); // quita el índice de cuenta (/u/1/): rompe los query params y es específico de una cuenta
  return url;
}
function signingLink_(loanId) {
  const base = webAppBaseUrl_();
  return base ? base + '?page=firmar&loan=' + encodeURIComponent(loanId) : '';
}

/** Correo al prestatario pidiéndole que firme (con enlace) + copia sin firmar para revisión. */
function emailSigningRequest_(loan, file) {
  if (!loan.email) return;
  const link = signingLink_(loan.loanId);
  const btn = link
    ? '<p style="margin:18px 0"><a href="' + link + '" style="background:#1c4587;color:#fff;text-decoration:none;padding:12px 22px;border-radius:6px;font-weight:bold">Revisar y firmar mi contrato</a></p>' +
      '<p style="font-size:12px;color:#888">O copie este enlace: ' + esc_(link) + '</p>'
    : '<p><b>Su contrato está listo para firmar.</b> Solicite el enlace de firma al prestamista.</p>';
  sendBrandedEmail_(loan.email, 'Revise y firme su contrato — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', su contrato ' + loan.loanId + ' está listo para firmar. ' + (link ? ('Ingrese a: ' + link) : ''),
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Su solicitud fue aprobada. <b>Adjuntamos una copia sin firmar del contrato</b> <b>' + esc_(loan.loanId) + '</b> para su revisión.</p>' +
    '<p><b>Le pedimos leer el contrato completo con atención antes de firmar</b>, en particular el costo total del crédito (Cláusula 3) y el régimen de mora (Cláusula 8). Antes de recibir el préstamo debe firmar el contrato.</p>' +
    '<ul><li>Capital: <b>' + fmtMoney_(loan.principal) + '</b></li><li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li><li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li></ul>' +
    btn + '<p style="font-size:12px;color:#888">Para firmar necesitará su DNI (' + esc_(loan.dni) + '). Revise la copia adjunta antes de firmar.</p>',
    file ? { attachments: [file.getAs('application/pdf')] } : {});
}

/** Correo con el PDF firmado, al prestatario y a Impulso Crédito (prestamista). */
function emailSignedContract_(loan, file) {
  const pdf = file.getAs('application/pdf');
  // Cronograma de pago: cuotas (préstamo grande) o pago único. Se calcula desde la Fecha
  // del Préstamo, que se fija al firmar → las fechas del correo son las definitivas.
  const sched = loanSchedule_(loan);
  const planTxt = sched
    ? sched.map((c, i) => 'Cuota ' + (i + 1) + '/' + sched.length + ': ' + fmtMoney_(c.amount) + ' vence ' + fmtDate_(c.due)).join(' · ')
    : 'Vencimiento: ' + fmtDate_(loan.dueDate);
  const planHtml = sched
    ? '<li>Plan de pago — <b>' + sched.length + ' cuotas mensuales</b>:<ul style="margin:4px 0">' +
      sched.map((c, i) => '<li>Cuota <b>' + (i + 1) + ' de ' + sched.length + '</b>: <b>' + fmtMoney_(c.amount) + '</b> — vence <b>' + fmtDate_(c.due) + '</b></li>').join('') +
      '</ul></li>'
    : '<li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li>';
  if (loan.email) {
    sendBrandedEmail_(loan.email, 'Contrato firmado — ' + loan.loanId,
      'Estimado/a ' + loan.name + ', adjuntamos su contrato de préstamo ' + loan.loanId + ' ya firmado. ' + planTxt + '.',
      '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Gracias. Adjuntamos su <b>contrato firmado</b> <b>' + esc_(loan.loanId) + '</b>.</p>' +
      '<ul><li>Capital: <b>' + fmtMoney_(loan.principal) + '</b></li><li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li>' +
      planHtml + '</ul>' +
      (sched ? '<p style="font-size:12px;color:#666">Cada cuota vence el día correspondiente; el recargo por mora se aplica por separado a cada cuota vencida.</p>' : ''),
      { attachments: [pdf] });
  }
  const rec = noticeRecipients_();
  if (rec.length) {
    sendBrandedEmail_(rec.join(','), 'Contrato firmado — ' + loan.loanId + ' — ' + loan.name,
      loan.name + ' firmó el contrato ' + loan.loanId + '. Ya puede desembolsar.',
      '<p>El prestatario <b>' + esc_(loan.name) + '</b> (DNI ' + esc_(loan.dni) + ') <b>firmó</b> el contrato <b>' + esc_(loan.loanId) + '</b>.</p>' +
      '<p>Fecha de firma: <b>' + fmtDateTime_(new Date()) + '</b>. <b>Ya puede desembolsar.</b></p>' +
      '<p>Se adjunta el contrato firmado para su archivo.</p>', { attachments: [pdf] });
  }
}

/** Menú/panel: reenvía el enlace de firma para el préstamo de la fila seleccionada en "Prestatarios". */
function resendSigningLink_() {
  const ss = getSS_(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== CFG.SHEETS.BORROWERS) { ui.alert('Abra la hoja "Prestatarios", seleccione la fila del préstamo y vuelva a usar esta opción.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Seleccione la fila de un préstamo.'); return; }
  const loanId = String(sh.getRange(row, 1).getValue()).trim();
  if (!loanId) { ui.alert('La fila seleccionada no tiene un préstamo.'); return; }
  const loan = readLoan_(sh, row);
  if (!loan.email) { ui.alert('El préstamo no tiene un correo cargado.'); return; }
  const st = String(sh.getRange(row, COL_SIGN_STATUS).getValue() || '').toUpperCase();
  if (st === SIGN.SIGNED && ui.alert('Este contrato ya figura FIRMADO. ¿Reenviar el enlace igualmente?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  try {
    if (st !== SIGN.SIGNED) sh.getRange(row, COL_SIGN_STATUS).setValue(SIGN.PENDING);
    let file = null; try { file = makeAgreementFile_(loan); } catch (e) { logError_('resendSigningLink_:pdf', e); }
    emailSigningRequest_(loan, file);
    ss.toast('Enlace de firma reenviado a ' + loan.email, 'Listo', 5);
  } catch (err) { logError_('resendSigningLink_', err); ui.alert('No se pudo reenviar: ' + (err.message || err)); }
}

/** Página web independiente para firmar el contrato (?page=firmar&loan=L-0001). */
function signingPageHtml_(loanId) {
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><meta charset="utf-8"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#1c4587;margin:6px 0}
    label{font-weight:bold;display:block;margin:12px 0 4px}
    input[type=text]{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px}
    button{margin-top:8px;background:#1c4587;color:#fff;border:0;border-radius:6px;padding:12px 18px;font-weight:bold;cursor:pointer}
    button.sec{background:#eee;color:#333}
    .terms{width:100%;border-collapse:collapse;margin:8px 0 4px}
    .terms td{padding:6px 8px;border:1px solid #ddd;font-size:14px}
    .terms td.k{background:#f0f4fa;font-weight:bold;width:42%}
    .tabs{margin:14px 0 6px}.tabs button{margin-right:6px;background:#eee;color:#333}
    .tabs button.on{background:#1c4587;color:#fff}
    canvas{border:1px dashed #999;border-radius:6px;width:100%;height:180px;touch-action:none;background:#fff}
    #prev{max-width:100%;max-height:180px;border:1px solid #ddd;border-radius:6px;margin-top:8px;display:none}
    .hint{font-size:12px;color:#888}.consent{margin:14px 0}.consent input{transform:scale(1.3);margin-right:8px}
    #out{margin-top:14px;padding:12px;border-radius:6px;display:none}.ok{background:#e6f4ea}.err{background:#f4cccc}
    #loadingView{text-align:center;padding:34px 12px}
    .pbar-track{width:100%;height:24px;background:#e3e9f4;border-radius:12px;overflow:hidden;margin:14px 0 8px}
    .pbar-fill{height:100%;width:0%;border-radius:12px;transition:width .45s ease;background-color:#1c4587;
      background-image:linear-gradient(45deg,rgba(255,255,255,.30) 25%,transparent 25%,transparent 50%,rgba(255,255,255,.30) 50%,rgba(255,255,255,.30) 75%,transparent 75%,transparent);
      background-size:28px 28px;animation:pbar-stripes 1s linear infinite}
    @keyframes pbar-stripes{from{background-position:0 0}to{background-position:28px 0}}
    .check{width:72px;height:72px;line-height:72px;margin:6px auto 10px;border-radius:50%;background:#38761d;color:#fff;font-size:40px;font-weight:bold;text-align:center}
    .sub{color:#555;margin:0 0 8px;line-height:1.45}
    .hide{display:none}</style></head><body>
    <div class="card">${brandHeaderHtml_()}<h1>Firmar Contrato</h1>

    <div id="gate">
      <p class="hint">Ingrese su DNI para ver y firmar el contrato <b>${esc_(loanId)}</b>.</p>
      <label>ID Préstamo</label><input type="text" id="loan" value="${esc_(loanId)}" readonly title="El ID del préstamo no se puede modificar" style="background:#eef1f6;color:#555;cursor:not-allowed">
      <label>DNI</label><input type="text" id="dni" placeholder="Su DNI">
      <button onclick="unlock()">Ver contrato</button>
      <div id="gerr" class="err" style="display:none;margin-top:12px;padding:12px;border-radius:6px"></div>
    </div>

    <div id="signview" class="hide">
      <p class="hint" style="margin:0 0 6px">Lea el contrato completo antes de firmar:</p>
      <iframe id="terms" title="Contrato" style="width:100%;height:62vh;border:1px solid #d8dee3;border-radius:6px;background:#fff"></iframe>
      <div class="tabs">
        <button id="tabDraw" class="on" onclick="showTab('draw')">✍ Dibujar firma</button>
        <button id="tabUp" onclick="showTab('upload')">🖼 Subir imagen</button>
      </div>
      <div id="paneDraw">
        <canvas id="cv"></canvas>
        <button class="sec" onclick="clearCanvas()">Borrar</button>
      </div>
      <div id="paneUp" class="hide">
        <input type="file" id="file" accept="image/*">
        <div class="hint">Tome una foto o cargue una imagen de su firma (fondo claro).</div>
        <img id="prev">
      </div>
      <label class="consent"><input type="checkbox" id="agree">He leído y acepto los términos de este contrato y lo firmo electrónicamente.</label>
      <button id="go" onclick="sign()">Firmar y enviar</button>
      <div id="out"></div>
    </div>

    <div id="loadingView" class="hide">
      <p id="loadPhase" style="font-weight:bold;color:#1c4587;margin:0;font-size:16px">Firmando su contrato…</p>
      <div class="pbar-track"><div class="pbar-fill" id="pbarFill"></div></div>
      <p id="loadPct" style="color:#1c4587;font-weight:bold;margin:0;font-size:20px">0%</p>
      <p class="hint" style="margin-top:6px">No cierre esta ventana.</p>
    </div>

    <div id="done" class="hide" style="text-align:center;padding:8px 0"></div>

    <script>
      var LOAN_ID = ${JSON.stringify(String(loanId))};
      var DNI = '', TAB = 'draw', uploaded = '', drew = false, cv, ctx, drawing = false, CLIENT_IP = '';
      function $(id){return document.getElementById(id);}
      // Captura de IP del firmante (mejor esfuerzo; si el servicio no responde, queda vacío).
      try{ fetch('https://api.ipify.org?format=json').then(function(r){return r.json();}).then(function(j){ CLIENT_IP=(j&&j.ip)||''; }).catch(function(){}); }catch(e){}
      function unlock(){
        DNI = $('dni').value; $('gerr').style.display='none';
        if(!$('loan').value || !DNI){ gerr('Ingrese el ID y el DNI.'); return; }
        google.script.run.withSuccessHandler(function(r){
          if(!r || !r.ok){ gerr((r&&r.msg)||'No se pudo verificar.'); return; }
          if(r.signed){ $('gate').classList.add('hide'); showDone(r.msg); return; }
          $('terms').srcdoc = r.contractHtml;
          $('gate').classList.add('hide'); $('signview').classList.remove('hide'); setupCanvas();
        }).withFailureHandler(function(e){ gerr(e.message||e); }).getContractForSigning($('loan').value, DNI);
      }
      function gerr(m){ var g=$('gerr'); g.textContent=m; g.style.display='block'; }
      function showTab(t){ TAB=t;
        $('tabDraw').className = t==='draw'?'on':''; $('tabUp').className = t==='upload'?'on':'';
        $('paneDraw').className = t==='draw'?'':'hide'; $('paneUp').className = t==='upload'?'':'hide';
        if(t==='draw') setupCanvas();
      }
      function setupCanvas(){
        cv=$('cv'); if(!cv) return; cv.width=cv.clientWidth; cv.height=180;
        ctx=cv.getContext('2d'); ctx.lineWidth=2.5; ctx.lineCap='round'; ctx.strokeStyle='#111';
        cv.onpointerdown=function(e){ drawing=true; drew=true; var p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); cv.setPointerCapture(e.pointerId); };
        cv.onpointermove=function(e){ if(!drawing) return; var p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); };
        cv.onpointerup=function(){ drawing=false; }; cv.onpointerleave=function(){ drawing=false; };
      }
      function pos(e){ var r=cv.getBoundingClientRect(); return {x:e.clientX-r.left, y:e.clientY-r.top}; }
      function clearCanvas(){ if(ctx){ ctx.clearRect(0,0,cv.width,cv.height); drew=false; } }
      $('file') && $('file').addEventListener('change', function(ev){
        var f=ev.target.files[0]; if(!f) return; var rd=new FileReader();
        rd.onload=function(){ var img=new Image(); img.onload=function(){
          var max=1000, s=Math.min(1, max/img.width), w=Math.round(img.width*s), h=Math.round(img.height*s);
          var c=document.createElement('canvas'); c.width=w; c.height=h;
          c.getContext('2d').drawImage(img,0,0,w,h);
          uploaded=c.toDataURL('image/jpeg',0.85);
          var pv=$('prev'); pv.src=uploaded; pv.style.display='block';
        }; img.src=rd.result; };
        rd.readAsDataURL(f);
      });
      /* ===== Barra de progreso (muestra la acción y el porcentaje) ===== */
      var progTimer=null, progVal=0;
      function setProgress(p,ph){ progVal=Math.max(0,Math.min(100,Math.round(p))); var f=$('pbarFill'),pc=$('loadPct'),lp=$('loadPhase'); if(f)f.style.width=progVal+'%'; if(pc)pc.textContent=progVal+'%'; if(ph&&lp)lp.textContent=ph; }
      function startProgress(){ var ph=[[10,'Verificando su identidad…'],[38,'Generando el contrato firmado…'],[66,'Guardando el documento…'],[88,'Enviando su copia por correo…']],i=1; setProgress(ph[0][0],ph[0][1]); clearInterval(progTimer);
        progTimer=setInterval(function(){ if(i<ph.length){ setProgress(ph[i][0],ph[i][1]); i++; } else if(progVal<95){ setProgress(progVal+1); } },700); }
      function stopProgress(){ clearInterval(progTimer); progTimer=null; }
      function finishProgress(cb){ stopProgress(); setProgress(100,'¡Listo!'); setTimeout(cb,400); }
      function sign(){
        var out=$('out'); out.style.display='none';
        if(!$('agree').checked){ msg('Debe aceptar los términos para firmar.', true); return; }
        var data='';
        if(TAB==='draw'){ if(!drew){ msg('Dibuje su firma primero.', true); return; } data=cv.toDataURL('image/png'); }
        else { if(!uploaded){ msg('Cargue una imagen de su firma primero.', true); return; } data=uploaded; }
        $('go').disabled=true;
        $('signview').classList.add('hide'); $('loadingView').classList.remove('hide'); startProgress();
        google.script.run.withSuccessHandler(function(r){
          if(r && r.ok){ finishProgress(function(){ $('loadingView').classList.add('hide'); showDone(r.message); }); }
          else { stopProgress(); $('loadingView').classList.add('hide'); $('signview').classList.remove('hide'); $('go').disabled=false; msg((r&&r.message)||'No se pudo firmar.', true); }
        }).withFailureHandler(function(e){
          stopProgress(); $('loadingView').classList.add('hide'); $('signview').classList.remove('hide'); $('go').disabled=false; msg(e.message||e, true);
        }).submitSignature(LOAN_ID, DNI, data, true, { userAgent: navigator.userAgent, method: TAB, ip: CLIENT_IP });
      }
      function msg(m, isErr){ var o=$('out'); o.className=isErr?'err':'ok'; o.textContent=m; o.style.display='block'; }
      function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
      function showDone(m){
        var d=$('done'); d.classList.remove('hide');
        d.innerHTML='<div class="check">✓</div>'+
          '<h1 style="color:#38761d">¡Contrato firmado!</h1>'+
          '<p class="sub">'+esc(m||'Su contrato fue firmado correctamente.')+'</p>'+
          '<p class="sub"><b>Revise su correo:</b> le enviamos una copia de su contrato firmado (PDF).</p>'+
          '<p class="hint">Ya puede cerrar esta ventana.</p>';
      }
    </script></div></body></html>`;
}

/* ===================== VERIFICAR / RECHAZAR (onEdit) ===================== */
function onEditInstallable(e) {
  try {
    const sh = e.range.getSheet(), name = sh.getName();
    const c0 = e.range.getColumn(), cN = c0 + e.range.getNumColumns() - 1;
    const r0 = Math.max(e.range.getRow(), 2), rN = e.range.getRow() + e.range.getNumRows() - 1;
    if (name === CFG.SHEETS.PAYMENTS) {
      // Al elegir un ID Préstamo (B): autogenerar el ID Pago (A) si falta,
      // autocompletar Nombre (G) y DNI (H), y preparar la Fecha de Pago (C):
      // formato de fecha y, si está vacía, la fecha de hoy.
      if (PP.LOAN_ID >= c0 && PP.LOAN_ID <= cN) {
        let seq = null;
        for (let row = r0; row <= rN; row++) {
          const id = String(sh.getRange(row, PP.LOAN_ID).getValue()).trim();
          if (!id) continue;
          if (!String(sh.getRange(row, PP.PAY_ID).getValue()).trim()) {
            if (seq === null) seq = nextPaymentSeq_(sh);
            sh.getRange(row, PP.PAY_ID).setValue('P-' + String(seq++).padStart(4, '0'));
          }
          const loan = findLoanById_(id);
          sh.getRange(row, PP.NAME, 1, 2).setValues([[loan ? (loan.name || '') : '', loan ? (loan.dni || '') : '']]);
          const fecha = sh.getRange(row, PP.DATE);
          fecha.setNumberFormat('yyyy-mm-dd').setDataValidation(datePicker_()); // selector de fecha
          if (fecha.getValue() === '') fecha.setValue(new Date());
        }
      }
      // Al escribir en la Fecha de Pago (C): garantizar el formato de fecha.
      if (PP.DATE >= c0 && PP.DATE <= cN)
        for (let row = r0; row <= rN; row++) sh.getRange(row, PP.DATE).setNumberFormat('yyyy-mm-dd');
      // Recalcular saldos al editar ID Préstamo (B), Fecha (C) o Monto (D).
      if (c0 <= PP.AMOUNT && cN >= PP.LOAN_ID) {
        const ids = {};
        for (let row = r0; row <= rN; row++) { const id = String(sh.getRange(row, PP.LOAN_ID).getValue()).trim(); if (id) ids[id] = true; }
        Object.keys(ids).forEach(id => { recalcLoanPayments_(id); updateLoanOutstanding_(id); });
      }
      // Casilla "Enviar recibo": envía/reenvía el recibo de pago de esa fila.
      // Se resuelve por NOMBRE de encabezado (funciona en el layout migrado o nuevo).
      const sendRecCol = headerMap_(sh)['Enviar recibo'] || PP.SEND_RECEIPT;
      if (sendRecCol >= c0 && sendRecCol <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, sendRecCol).getValue() !== true) continue;
          sh.getRange(row, sendRecCol).setValue(false);
          try { getSS_().toast(sendPaymentReceiptForRow_(sh, row), 'Recibo de pago', 5); }
          catch (err) { logError_('onEdit:PAGOS:recibo', err); getSS_().toast(err.message || String(err), '⚠ No se pudo enviar', 8); }
        }
      }
      // Casilla "Generar recibo": crea el PDF del recibo (monto + saldo) SIN enviar correo.
      const genRecCol = headerMap_(sh)['Generar recibo'] || PP.GEN_RECEIPT;
      if (genRecCol >= c0 && genRecCol <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, genRecCol).getValue() !== true) continue;
          sh.getRange(row, genRecCol).setValue(false);
          try { getSS_().toast(generatePaymentReceiptForRow_(sh, row), 'Recibo generado', 5); }
          catch (err) { logError_('onEdit:PAGOS:generarRecibo', err); getSS_().toast(err.message || String(err), '⚠ No se pudo generar', 8); }
        }
      }
    } else if (name === CFG.SHEETS.BORROWERS) {
      // Casilla "Eliminar" (col U): quita el préstamo de la fila tildada.
      if (PB.DELETE >= c0 && PB.DELETE <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, PB.DELETE).getValue() !== true) continue;
          sh.getRange(row, PB.DELETE).setValue(false);
          deleteLoanRow_(sh, row);
        }
      }
      // Casilla "Enviar recibo desembolso": envía el recibo de desembolso si el contrato
      // está FIRMADO. Columnas resueltas por NOMBRE (layout migrado o nuevo).
      const bhm = headerMap_(sh);
      const disbCol = bhm['Enviar recibo desembolso'] || PB.SEND_DISB;
      const firmaCol = bhm['Estado de Firma'] || PB.SIGN_STATUS;
      const bLoanCol = bhm['ID Préstamo'] || bhm['ID Prestamo'] || PB.LOAN_ID;
      if (disbCol >= c0 && disbCol <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, disbCol).getValue() !== true) continue;
          sh.getRange(row, disbCol).setValue(false);
          const loanId = String(sh.getRange(row, bLoanCol).getValue()).trim();
          if (!loanId) continue;
          const firma = String(sh.getRange(row, firmaCol).getValue()).trim().toUpperCase();
          if (firma !== SIGN.SIGNED) { getSS_().toast('El contrato de ' + loanId + ' no está ' + SIGN.SIGNED + '. No se envió el recibo de desembolso.', '⚠ Pendiente de firma', 8); continue; }
          try {
            const msg = sbSendDisbursementReceipt(loanId);
            sh.getRange(row, disbCol).setNote('Recibo de desembolso enviado el ' + fmtDate_(new Date()) + '.');
            getSS_().toast(msg, 'Recibo de desembolso', 5);
          } catch (err) { logError_('onEdit:BORROWERS:desembolso', err); getSS_().toast(err.message || String(err), '⚠ No se pudo enviar', 8); }
        }
      }
      if ([PB.LOAN_ID, PB.PRINCIPAL, PB.TERM, PB.LOAN_DATE].some(c => c >= c0 && c <= cN)) for (let row = r0; row <= rN; row++) writeOutstandingRow_(sh, row);
    } else if (name === CFG.SHEETS.AGREEMENT) {
      const a1 = e.range.getA1Notation();
      if (a1 === 'B15' && sh.getRange('B15').getValue() === true) { sh.getRange('B15').setValue(false); contractStudioPreview_(sh); }
      else if (a1 === 'B16' && sh.getRange('B16').getValue() === true) { sh.getRange('B16').setValue(false); contractStudioSend_(sh); }
    } else if (name === CFG.SHEETS.STATEMENTS) {
      const a1 = e.range.getA1Notation();
      if (a1 === 'B18' && sh.getRange('B18').getValue() === true) { sh.getRange('B18').setValue(false); statementStudioPreview_(sh); }
      else if (a1 === 'B19' && sh.getRange('B19').getValue() === true) { sh.getRange('B19').setValue(false); statementStudioSend_(sh); }
    } else if (name === CFG.SHEETS.REMINDER) {
      if (e.range.getA1Notation() === 'B2' && sh.getRange('B2').getValue() === true) {
        sh.getRange('B2').setValue(false); rebuildRemindersSheet_(getSS_());
      } else if ((REM_GEN_COL >= c0 && REM_GEN_COL <= cN) || (REM_SEND_COL >= c0 && REM_SEND_COL <= cN)) {
        for (let row = rN; row >= Math.max(r0, REM_DATA_START); row--) {
          if (sh.getRange(row, 1).getValue() === '') continue;
          if (sh.getRange(row, REM_GEN_COL).getValue() === true) { sh.getRange(row, REM_GEN_COL).setValue(false); reminderRowGenerate_(sh, row); }
          else if (sh.getRange(row, REM_SEND_COL).getValue() === true) { sh.getRange(row, REM_SEND_COL).setValue(false); reminderRowSend_(sh, row); }
        }
      }
    } else if (name === CFG.SHEETS.LATE) {
      if (LATE_SEND_COL >= c0 && LATE_SEND_COL <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, LATE_SEND_COL).getValue() !== true) continue;
          sh.getRange(row, LATE_SEND_COL).setValue(false);
          const loanId = String(sh.getRange(row, 1).getValue()).trim();
          if (!loanId) continue;
          try {
            const res = sendOverdueNotice_(loanId);
            sh.getRange(row, LATE_SENT_COL).setValue(res.date).setNumberFormat('yyyy-mm-dd');
            getSS_().toast(res.message, 'Aviso de mora', 5);
          } catch (err) { logError_('onEdit:LATE', err); getSS_().toast(err.message || String(err), '⚠ No se pudo enviar', 8); }
        }
      }
    } else if (name === CFG.SHEETS.NEW) {
      const m = headerMap_(sh), verCol = m['Verificado?'], rejCol = m['Rechazar?'], bcraCol = m['Verificar BCRA?'], ovrCol = m['Anular límites'];
      const acts = [];
      for (let row = r0; row <= rN; row++) {
        if (bcraCol && bcraCol >= c0 && bcraCol <= cN && sh.getRange(row, bcraCol).getValue() === true) { bcraCheckRow_(sh, row); continue; }
        // Anular límites por sí solo no aprueba: sólo marca la fila para el próximo "Verificado?".
        const override = ovrCol ? (sh.getRange(row, ovrCol).getValue() === true) : false;
        if (verCol && verCol >= c0 && verCol <= cN && sh.getRange(row, verCol).getValue() === true) acts.push({ row: row, kind: 'v', override: override });
        else if (rejCol && rejCol >= c0 && rejCol <= cN && sh.getRange(row, rejCol).getValue() === true) acts.push({ row: row, kind: 'r' });
      }
      // Aprobación: usa el flujo V2 (esquema migrado días + Clientes normalizado) si
      // está disponible; si no, el flujo clásico month-based. `override` anula los topes.
      acts.sort((a, b) => b.row - a.row).forEach(a => a.kind === 'v'
        ? (typeof verifyApplicantV2_ === 'function' ? verifyApplicantV2_(sh, a.row, a.override) : verifyApplicant_(sh, a.row, a.override))
        : rejectApplicant_(sh, a.row));
    }
  } catch (err) { logError_('onEditInstallable', err); }
}

function headerMap_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], m = {};
  h.forEach((x, i) => { const k = String(x).trim(); if (k) m[k] = i + 1; });
  return m;
}

function verifyApplicant_(nb, row) {
  const m = headerMap_(nb);
  // Lee la fila completa de una sola vez (antes: un getRange().getValue() por campo).
  const rowVals = nb.getRange(row, 1, 1, nb.getLastColumn()).getValues()[0];
  const g = t => m[t] ? rowVals[m[t] - 1] : '';
  const resCol = m['Resultado'], verCol = m['Verificado?'];
  const fail = msg => { if (resCol) nb.getRange(row, resCol).setValue(msg); if (verCol) nb.getRange(row, verCol).setValue(false); };
  const name = String(g('Nombre completo')).trim(), email = g('Correo'), dni = String(g('DNI')).trim(),
    phone = g('Teléfono'), amount = Number(g('Monto Solicitado')) || 0, term = Number(g('Plazo (meses)')) || 0,
    natId = g('Foto del frente del DNI') || g('Foto DNI'), workId = g('Foto del dorso del DNI') || g('Foto CUIL');
  if (!name || !dni) { fail('⚠ Falta nombre o DNI'); return; }
  if (term !== 1 && term !== 2 && term !== 15) { fail('⚠ Plazo debe ser 15, 1 o 2'); return; }
  if (!amount) { fail('⚠ Monto inválido'); return; }
  if (!natId || !workId) { fail('⚠ Faltan fotos (DNI/CUIL)'); return; }
  // Efectivo disponible ANTES de colocar este préstamo (baja al prestar el capital).
  const availBefore = fundAvailable_();
  // Duplicado por DNI (contra la hoja maestra "Clientes").
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  if (clienteExistsByDni_(dni)) fail('⚠ DNI ya existe (cliente registrado) — revisar antes de aprobar');
  // Modelo normalizado: reutiliza el cliente (DNI+Correo) o crea uno nuevo.
  const clienteId = getOrCreateCliente_(name, email, dni, phone);
  // Captura completa: forma de pago elegida, y CUIL/Dirección (de la verificación BCRA) al cliente.
  try {
    setClienteFieldByHeader_(clienteId, CLIENTE_PAY_HEADER, String(g('Forma de Pago') || '').trim());
    setClienteFieldByHeader_(clienteId, 'CUIL', String(g('CUIL') || '').trim());
    setClienteFieldByHeader_(clienteId, 'Dirección', String(g('Dirección') || '').trim());
  } catch (e) { logError_('verifyApplicant_:clienteExtra', e); }
  const loanId = nextLoanId_(), tRow = firstEmptyBorrowerRow_(bs);
  bs.getRange(tRow, PB.LOAN_ID, 1, 2).setValues([[loanId, clienteId]]);          // A:B (Nombre/DNI en C:D son fórmulas)
  bs.getRange(tRow, PB.PRINCIPAL, 1, 2).setValues([[amount, term]]);             // E:F Capital, Plazo
  bs.getRange(tRow, PB.LOAN_DATE).setValue(new Date()); bs.getRange(tRow, PB.LOAN_DATE).setNumberFormat('yyyy-mm-dd');
  writeOutstandingRow_(bs, tRow);
  const folder = borrowerFolder_(borrowerFolderName_(name, dni));
  [natId, workId].forEach(u => extractDriveIds_(u).forEach(id => { try { DriveApp.getFileById(id).moveTo(folder); } catch (e) {} }));
  bs.getRange(tRow, PB.CLIENT_ID).setNote('📁 Carpeta: ' + folder.getUrl() + (natId ? '\nDNI: ' + natId : '') + (workId ? '\nCUIL: ' + workId : ''));
  // Firma obligatoria: el préstamo queda PENDIENTE DE FIRMA hasta que el prestatario firme.
  bs.getRange(tRow, COL_SIGN_STATUS).setValue(SIGN.PENDING);
  // Contrato (copia de revisión) + correo pidiendo la FIRMA con enlace.
  try {
    const loan = readLoan_(bs, tRow);
    const file = makeAgreementFile_(loan);
    markAgreementSent_(loanId, file);
    if (email) emailSigningRequest_(loan, file);
  } catch (e) { logError_('verifyApplicant_:contrato', e); }
  nb.deleteRow(row);
  const shortfall = amount > availBefore;
  const availAfter = round2_(availBefore - amount);
  try {
    getSS_().toast(
      (shortfall ? '⚠ Fondo insuficiente. ' : '') + 'Pendiente de firma. Disponible tras prestar: ' + fmtMoney_(availAfter),
      (shortfall ? '⚠ ' : '') + 'Aprobado ' + name + ' → ' + loanId + ' (firma pendiente)', shortfall ? 8 : 6);
  } catch (e) {}
}

function rejectApplicant_(nb, row) {
  const m = headerMap_(nb);
  const g = t => m[t] ? nb.getRange(row, m[t]).getValue() : '';
  const name = String(g('Nombre completo')).trim(), email = g('Correo'), dni = g('DNI'), phone = g('Teléfono');
  const arch = getSS_().getSheetByName(CFG.SHEETS.REJECTED) || setupRejected_(getSS_());
  arch.appendRow([new Date(), name, email, dni, phone, 'Rechazado desde Nuevos Prestatarios']);
  if (email) {
    try {
      sendBrandedEmail_(email, 'Actualización de su solicitud',
        'Estimado/a ' + name + ', gracias por su solicitud. Lamentamos informarle que no podemos avanzar en este momento.',
        '<p>Estimado/a ' + esc_(name) + ',</p><p>Gracias por su solicitud. Lamentamos informarle que no podemos avanzar en este momento.</p>');
    } catch (e) { logError_('rejectApplicant_:email', e); }
  }
  nb.deleteRow(row);
  try { getSS_().toast('Rechazado ' + name, 'Archivado', 5); } catch (e) {}
}

/* ================== BCRA Central de Deudores ==================
 * Consulta la Central de Deudores del BCRA en el momento de la carga,
 * antes de desembolsar. Acepta DNI o CUIL. Los montos que devuelve la
 * API vienen en MILES de pesos; acá se multiplican por 1000 para mostrar
 * pesos reales. También disponible como fórmula: =BCRA_SITUACION(D2).
 */
var BCRA_BASE = 'https://api.bcra.gob.ar/CentralDeDeudores/v1.0';

/** Dígito verificador AFIP/ARCA estándar. Devuelve [prefijo, dígito]. */
function cuilCheckDigit_(prefix, dni) {
  var body = ('' + prefix) + Utilities.formatString('%08d', Number(dni));
  var w = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2], sum = 0;
  for (var i = 0; i < 10; i++) sum += Number(body.charAt(i)) * w[i];
  var rem = sum % 11;
  if (rem === 0) return [prefix, 0];
  if (rem === 1) return [23, prefix === 20 ? 9 : 4];
  return [prefix, 11 - rem];
}

/** Genera candidatos de CUIL a partir de un DNI (prefijos 20 y 27). */
function cuilCandidates_(dni) {
  dni = String(dni).replace(/\D/g, '');
  var out = [];
  [20, 27].forEach(function (p) {
    var r = cuilCheckDigit_(p, dni);
    var c = Utilities.formatString('%02d', r[0]) +
            Utilities.formatString('%08d', Number(dni)) + r[1];
    if (out.indexOf(c) === -1) out.push(c);
  });
  return out;
}

/** GET a un endpoint del BCRA. Devuelve el objeto parseado, o null en 404. */
function bcraGet_(path) {
  // Modo prueba: devuelve la respuesta simulada (objeto o función(path)) sin red.
  if (TEST_MODE) return (typeof _bcraStub === 'function') ? _bcraStub(path) : _bcraStub;
  var res = UrlFetchApp.fetch(BCRA_BASE + path, {
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = res.getResponseCode();
  if (code === 404) return null;
  if (code === 429 || code >= 500) {       // throttled — un reintento lento
    Utilities.sleep(3000);
    res = UrlFetchApp.fetch(BCRA_BASE + path, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
  } else if (code !== 200) {
    return null;
  }
  try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
}

/** Consulta principal. Acepta un DNI o un CUIL completo. */
function consultarBcra_(dniOrCuil) {
  var raw = String(dniOrCuil).replace(/\D/g, '');
  var cands = raw.length === 11 ? [raw] : cuilCandidates_(raw);

  for (var i = 0; i < cands.length; i++) {
    var d = bcraGet_('/Deudas/' + cands[i]);
    var r = d && (d.results || d);
    if (!r || !r.periodos || !r.periodos.length) continue;

    var ents = r.periodos[0].entidades || [];
    var total = 0, perf = 0, worst = 0, deudas = [];
    ents.forEach(function (e) {
      var m = Number(e.monto || 0) * 1000, s = Number(e.situacion || 0);
      total += m;
      if (s <= 2) perf += m;
      if (s > worst) worst = s;
      deudas.push({ entidad: String(e.entidad || '').trim(), situacion: s, monto: m });
    });
    deudas.sort(function (a, b) { return b.situacion - a.situacion; }); // peor primero

    // Tendencia 24 meses
    var h = bcraGet_('/Deudas/Historicas/' + cands[i]);
    var hr = h && (h.results || h), serie = [];
    if (hr && hr.periodos) {
      hr.periodos.forEach(function (p) {
        var w = 0;
        (p.entidades || []).forEach(function (e) {
          w = Math.max(w, Number(e.situacion || 0));
        });
        serie.push({ p: String(p.periodo), s: w });
      });
      serie.sort(function (a, b) { return a.p < b.p ? -1 : 1; });
    }
    var latest = serie.length ? serie[serie.length - 1].s : worst;
    var sixAgo = serie.length >= 7 ? serie[serie.length - 7].s : latest;

    var share = total ? perf / total : 1;
    var decision =
      (worst >= 4 && share < 0.5) ? 'RECHAZAR' :
      (latest - sixAgo >= 2 || worst >= 4) ? 'REVISAR' : 'OK';

    return {
      cuil: cands[i],
      nombre: r.denominacion || '',
      peorSituacion: worst,
      entidades: ents.length,
      deudaTotal: total,
      deudaPerformnte: perf,
      pctPerformante: Math.round(share * 1000) / 10,
      delta6m: latest - sixAgo,
      historia: serie.map(function (x) { return x.s; }).join(''),
      deudas: deudas,
      decision: decision
    };
  }
  return { cuil: '', nombre: '', peorSituacion: 0, entidades: 0, deudaTotal: 0,
           deudaPerformnte: 0, pctPerformante: 100, delta6m: 0, historia: '',
           deudas: [], decision: 'SIN REGISTRO' };
}

/**
 * Fórmula de hoja:  =BCRA_SITUACION(D2)
 * Devuelve un veredicto de una línea para la celda.
 * @customfunction
 */
function BCRA_SITUACION(dniOrCuil) {
  if (!dniOrCuil) return '';
  var r = consultarBcra_(dniOrCuil);
  if (r.decision === 'SIN REGISTRO') return 'SIN REGISTRO';
  return Utilities.formatString('%s · sit %d · %s%% al día · %s',
    r.decision, r.peorSituacion, r.pctPerformante,
    r.delta6m >= 1 ? 'DETERIORANDO' : (r.delta6m <= -1 ? 'MEJORANDO' : 'ESTABLE'));
}

/**
 * Puntaje de desempeño (0–100) y calificación de riesgo según estándares de
 * clasificación de deudores del BCRA (situación 1–6):
 *   1 Normal · 2 Riesgo bajo/seguimiento especial · 3 Con problemas (medio)
 *   4 Alto riesgo de insolvencia · 5 Irrecuperable · 6 Irrec. por disp. técnica.
 * Devuelve { puntaje, rating }.
 */
function bcraScore_(r) {
  if (!r || r.decision === 'SIN REGISTRO') return { puntaje: 85, rating: 'Sin historial' };
  var worst = Number(r.peorSituacion) || 0;
  var sitBase = ({ 1: 100, 2: 75, 3: 50, 4: 25, 5: 5, 6: 0 })[worst];
  if (sitBase === undefined) sitBase = worst > 6 ? 0 : 100; // sin deuda reportada → normal
  var pct = Number(r.pctPerformante); if (isNaN(pct)) pct = 100;
  var puntaje = Math.round(0.6 * sitBase + 0.4 * pct);
  var delta = Number(r.delta6m) || 0;
  if (delta >= 2) puntaje -= 10;        // deteriorando
  else if (delta <= -2) puntaje += 5;   // mejorando
  puntaje = Math.max(0, Math.min(100, puntaje));

  var rating =
    puntaje >= 80 ? 'Riesgo Bajo (A)' :
    puntaje >= 60 ? 'Riesgo Medio-Bajo (B)' :
    puntaje >= 40 ? 'Riesgo Medio (C)' :
    puntaje >= 20 ? 'Riesgo Alto (D)' : 'Riesgo Muy Alto (E)';
  // Piso por situación: la calificación nunca contradice al BCRA.
  var order = ['Riesgo Bajo (A)', 'Riesgo Medio-Bajo (B)', 'Riesgo Medio (C)', 'Riesgo Alto (D)', 'Riesgo Muy Alto (E)'];
  if (worst >= 5 && order.indexOf(rating) < 4) rating = 'Riesgo Muy Alto (E)';
  else if (worst === 4 && order.indexOf(rating) < 3) rating = 'Riesgo Alto (D)';
  return { puntaje: puntaje, rating: rating };
}

/** Etiqueta de la clasificación de deudores del BCRA (situación 1–6). */
function bcraSituLabel_(s) {
  return ({
    1: 'Normal', 2: 'Riesgo bajo (seguimiento especial)', 3: 'Con problemas (riesgo medio)',
    4: 'Alto riesgo de insolvencia', 5: 'Irrecuperable', 6: 'Irrecuperable por disposición técnica'
  })[Number(s)] || 'Sin clasificación';
}

/**
 * Resumen en lenguaje llano que interpreta cada columna para esta persona:
 * situación, % al día, tendencia, puntaje/calificación y la decisión sugerida.
 */
function bcraResumen_(r, s) {
  if (!r || r.decision === 'SIN REGISTRO') {
    return 'Sin registro en la Central de Deudores del BCRA (no se hallaron deudas). ' +
      'Puntaje ' + s.puntaje + '/100 → ' + s.rating + '. ' +
      'Historial crediticio limitado: verificar identidad y CUIL antes de avanzar.';
  }
  var tend = r.delta6m >= 1 ? 'deteriorando' : (r.delta6m <= -1 ? 'mejorando' : 'estable');
  var decTxt = ({
    'OK': 'situación favorable, puede avanzar',
    'REVISAR': 'revisar manualmente antes de prestar',
    'RECHAZAR': 'no se recomienda prestar'
  })[r.decision] || r.decision;
  var nombre = r.nombre ? r.nombre + ' ' : '';
  var ents = Number(r.entidades) || 0;
  var detalle = (r.deudas && r.deudas.length)
    ? ' Detalle por préstamo: ' + r.deudas.map(function (d, i) {
        return (i + 1) + ') ' + (d.entidad || 'Entidad s/nombre') + ' — sit ' + d.situacion +
          ' (' + bcraSituLabel_(d.situacion) + ')';
      }).join('; ') + '.'
    : '';
  return nombre + '(CUIL ' + r.cuil + '). ' +
    'Tiene ' + ents + ' préstamo' + (ents === 1 ? '' : 's') + ' reportado' + (ents === 1 ? '' : 's') + ' en BCRA. ' +
    'Peor situación ' + r.peorSituacion + ' — ' + bcraSituLabel_(r.peorSituacion) + '. ' +
    r.pctPerformante + '% de la deuda al día' +
    (Number(r.pctPerformante) < 100 ? '; el resto en categorías de mayor riesgo' : '') + '. ' +
    'Tendencia 6 meses: ' + tend + '. ' +
    'Puntaje ' + s.puntaje + '/100 → ' + s.rating + '. ' +
    'Decisión: ' + r.decision + ' (' + decTxt + ').' + detalle;
}

/**
 * Ejecuta la consulta BCRA para una fila de "Nuevos Prestatarios" cuando se
 * tilda "Verificar BCRA?". Usa DNI o CUIL según "Parámetro BCRA".
 */
function bcraCheckRow_(sh, row) {
  const m = headerMap_(sh), bcraCol = m['Verificar BCRA?'];
  const uncheck = () => { if (bcraCol) sh.getRange(row, bcraCol).setValue(false); };
  try {
    const param = String((m['Parámetro BCRA'] ? sh.getRange(row, m['Parámetro BCRA']).getValue() : '') || 'DNI').toUpperCase();
    const dni = String((m['DNI'] ? sh.getRange(row, m['DNI']).getValue() : '') || '').trim();
    const cuil = String((m['CUIL'] ? sh.getRange(row, m['CUIL']).getValue() : '') || '').trim();
    const key = (param === 'CUIL' && cuil) ? cuil : dni;
    if (!String(key).replace(/\D/g, '')) { uncheck(); getSS_().toast('Falta DNI o CUIL en la fila ' + row, '⚠ BCRA', 6); return; }
    const r = consultarBcra_(key);
    const s = bcraScore_(r);
    sh.getRange(row, m['CUIL'], 1, 8).setValues([[
      r.cuil, r.nombre, r.peorSituacion, r.pctPerformante, r.decision, s.puntaje, s.rating, bcraResumen_(r, s)
    ]]);
    uncheck();
    getSS_().toast((r.nombre || key) + ' → ' + r.decision + ' · ' + s.rating + ' · ' + s.puntaje + '/100', 'BCRA', 6);
  } catch (err) {
    logError_('bcraCheckRow_', err);
    uncheck();
    try { getSS_().toast(err.message || String(err), '⚠ BCRA', 8); } catch (e) {}
  }
}

/**
 * Escribe el resultado BCRA en una fila y devuelve true solo si es 'OK'.
 * Bloque de 7 columnas: CUIL, Nombre BCRA, Peor Situación, % al día, Decisión, Puntaje, Calificación.
 */
function verificarSolicitante_(sheet, row, dniCol, outCol) {
  var dni = sheet.getRange(row, dniCol).getValue();
  var r = consultarBcra_(dni);
  var s = bcraScore_(r);
  sheet.getRange(row, outCol, 1, 8).setValues([[
    r.cuil, r.nombre, r.peorSituacion, r.pctPerformante, r.decision, s.puntaje, s.rating, bcraResumen_(r, s)
  ]]);
  return r.decision === 'OK';
}

/** Menú ▸ verifica en BCRA todos los clientes (hoja "Clientes", modelo normalizado). */
function backfillPrestatarios() {
  var sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS);
  if (!sh) { getSS_().toast('No existe la hoja "Clientes". Ejecute la migración/Configurar primero.', '⚠ BCRA', 6); return; }
  var last = sh.getLastRow();
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var dniCol = head.indexOf('DNI') + 1;
  if (dniCol < 1) { getSS_().toast('No se encontró la columna DNI', '⚠ BCRA', 6); return; }
  var outCol = sh.getLastColumn() + 1;
  sh.getRange(1, outCol, 1, 8)
    .setValues([['CUIL', 'Nombre BCRA', 'Peor Situación', '% al día', 'Decisión', 'Puntaje BCRA', 'Calificación de Riesgo', 'Resumen BCRA']]);
  for (var r = 2; r <= last; r++) {
    if (!sh.getRange(r, dniCol).getValue()) continue;
    verificarSolicitante_(sh, r, dniCol, outCol);
    Utilities.sleep(1500);   // respetar el límite por IP del BCRA
  }
  try { getSS_().toast('Verificación BCRA completada', 'BCRA', 5); } catch (e) {}
}

/* ===================== TAREAS DIARIAS (recordatorios) ===================== */
function dailyTasks() {
  guard_('dailyTasks', function () {
    updateAllOutstanding_();
    rebuildLateSheet_(getSS_());
    rebuildRemindersSheet_(getSS_());
    const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
    if (last < 2) return;
    const noticeDays = Number(getSetting_('Días de aviso antes del vencimiento')) || 3;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    for (let row = 2; row <= last; row++) {
      const loan = readLoan_(bs, row); if (!loan.loanId || !loan.email) continue;
      const status = bs.getRange(row, PB.STATE).getValue(), outstanding = Number(bs.getRange(row, PB.BALANCE).getValue()) || 0;
      if (status === ST.PAID || outstanding <= 0) continue;
      const due = loan.dueDate instanceof Date ? loan.dueDate : null; if (!due) continue;
      const dd = new Date(due); dd.setHours(0, 0, 0, 0);
      const days = Math.round((dd - today) / 86400000);
      const lastNotice = bs.getRange(row, PB.NOTICE).getValue();
      const lastStr = (lastNotice instanceof Date) ? fmtDate_(lastNotice) : '';
      let send = false, subject = '', intro = '';
      if (days >= 0 && days <= noticeDays && lastStr !== fmtDate_(today)) {
        send = true; subject = 'Recordatorio: su préstamo vence pronto';
        intro = 'Le recordamos que su préstamo ' + loan.loanId + ' vence el ' + fmtDate_(due) + '.';
      } else if (days < 0) {
        // en mora: avisar como máximo cada 7 días
        const daysSince = lastNotice instanceof Date ? Math.round((today - new Date(lastNotice.getFullYear(), lastNotice.getMonth(), lastNotice.getDate())) / 86400000) : 999;
        if (daysSince >= 7) { send = true; subject = 'Aviso de mora — préstamo ' + loan.loanId; intro = 'Su préstamo ' + loan.loanId + ' venció el ' + fmtDate_(due) + ' (' + daysLate_(due, today) + ' día(s) de atraso). Se aplica un recargo por mora del ' + lateFeePctText_() + ' por día (' + fmtMoney_(round2_(loan.totalDue * lateFeeRate_())) + ' por día) sobre el total a devolver, con tope del ' + moraCapPctText_() + ' del capital.'; }
      }
      if (send) {
        try {
          sendBrandedEmail_(loan.email, subject, intro + ' Saldo actual: ' + fmtMoney_(outstanding) + '.',
            '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>' + esc_(intro) + '</p><p><b>Saldo pendiente: ' + fmtMoney_(outstanding) + '</b></p>');
          bs.getRange(row, PB.NOTICE).setValue(today);
        } catch (e) { logError_('dailyTasks:email', e); }
      }
    }
  });
}

/* ===================== HELPERS ===================== */
function recalcLoanPayments_(loanId) {
  // "Saldo Posterior" es una fórmula viva (paymentBalanceFormula_); la reafirmamos en las
  // filas de pago de este préstamo por si fueron pegadas/importadas sin la fórmula.
  const sh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS), last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, PP.LOAN_ID, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === loanId) { const r = i + 2; sh.getRange(r, PP.BALANCE).setFormula(paymentBalanceFormula_(r)); }
  }
}

/** Próximo número de ID Pago (máx. existente + 1), leyendo la columna "ID Pago". */
function nextPaymentSeq_(ps) {
  ps = ps || getSS_().getSheetByName(CFG.SHEETS.PAYMENTS);
  const last = ps.getLastRow(); let max = 0;
  if (last >= 2) ps.getRange(2, PP.PAY_ID, last - 1, 1).getValues().forEach(r => {
    const m = /(\d+)\s*$/.exec(String(r[0] || '')); if (m) max = Math.max(max, Number(m[1]));
  });
  return max + 1;
}
/** ID Pago autogenerado, formato P-0001. */
function nextPaymentId_(ps) { return 'P-' + String(nextPaymentSeq_(ps)).padStart(4, '0'); }

/**
 * Rellena Nombre (G) y DNI (H) de "Pagos" para todas las filas con ID Préstamo,
 * resolviendo préstamo → cliente. Idempotente; se ejecuta en ⑤ Actualizar y en
 * la migración, y así completa también los pagos ya existentes.
 */
function refreshPaymentNames_() {
  const ps = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS); if (!ps) return;
  // Asegura los encabezados Nombre/DNI en hojas creadas antes de esta versión
  // (no destructivo: solo escribe la fila 1 y el ancho de columna).
  if (String(ps.getRange(1, PP.NAME).getValue()).trim() !== 'Nombre') {
    ps.getRange(1, PP.NAME, 1, 2).setValues([['Nombre', 'DNI']])
      .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff');
    ps.setColumnWidth(PP.NAME, 200); ps.setColumnWidth(PP.DNI, 110);
  }
  const last = ps.getLastRow(); if (last < 2) return;
  const ids = ps.getRange(2, PP.LOAN_ID, last - 1, 1).getValues();
  const cache = {};
  const out = ids.map(r => {
    const id = String(r[0] || '').trim(); if (!id) return ['', ''];
    if (!(id in cache)) { const loan = findLoanById_(id); cache[id] = loan ? [loan.name || '', loan.dni || ''] : ['', '']; }
    return cache[id];
  });
  ps.getRange(2, PP.NAME, out.length, 2).setValues(out);
}

/* ==================================================================
 *  MIGRACIÓN AL MODELO NORMALIZADO (3NF)
 * ------------------------------------------------------------------
 *  Convierte el layout heredado (identidad duplicada en cada fila de
 *  "Prestatarios"/"Pagos") al modelo normalizado con hoja maestra
 *  "Clientes" + referencia por "ID Cliente".
 *   - Dedup por DNI+Correo (ambos deben coincidir, normalizados).
 *   - En conflicto, gana el préstamo MÁS RECIENTE (por Fecha del Préstamo).
 *  Es idempotente: si ya está migrada (col B = "ID Cliente"), no hace nada.
 *  DESTRUCTIVA: reorganiza columnas. Requiere respaldo (duplicar la hoja).
 * ================================================================== */
const MIG_PROP = 'MIGRATION_V3NF_STATE';

/** Menú: confirma respaldo y ejecuta la migración. */
function migrateToNormalizedClientesConfirm() {
  const ui = SpreadsheetApp.getUi();
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs) { ui.alert('Migración', 'No existe la hoja "Prestatarios". Ejecute ① Configurar primero.', ui.ButtonSet.OK); return; }
  const state = PropertiesService.getScriptProperties().getProperty(MIG_PROP) || '';
  if (String(bs.getRange(1, 2).getValue()).trim() === 'ID Cliente') {
    // Caso especial: Prestatarios migrado pero "Pagos" quedó sin migrar (ejecución
    // previa interrumpida). Ofrecer terminar solo esa parte, sin volver a empezar.
    if (pagosIsLegacy_(getSS_().getSheetByName(CFG.SHEETS.PAYMENTS))) {
      const rr = ui.alert('Terminar migración de Pagos',
        'Se detectó una migración incompleta: "Prestatarios" ya está en el modelo normalizado, pero "Pagos" ' +
        'aún tiene el layout anterior (por eso "Total cobrado" aparece en $0).\n\n¿Migrar "Pagos" ahora y recalcular las vistas?',
        ui.ButtonSet.YES_NO);
      if (rr === ui.Button.YES) { const msg = migrateToNormalizedClientes_(); ui.alert('Listo', msg, ui.ButtonSet.OK); }
      return;
    }
    if (state && state !== 'columns_removed') {
      ui.alert('Migración incompleta',
        'Una migración anterior NO terminó (estado: ' + state + '). La hoja quedó inconsistente y los datos ' +
        'heredados de "Prestatarios" ya no están en esta copia.\n\nRestaure desde una copia NUEVA de la hoja ' +
        'ORIGINAL (sin migrar) y ejecute de nuevo la migración con el código corregido.', ui.ButtonSet.OK);
    } else {
      ui.alert('Migración', 'La hoja ya está en el modelo normalizado (columna B = "ID Cliente"). No hay nada que migrar.', ui.ButtonSet.OK);
    }
    return;
  }
  const r1 = ui.alert('Migración al modelo normalizado (3NF)',
    'Esta operación es DESTRUCTIVA: crea la hoja "Clientes", reorganiza "Prestatarios" y "Pagos" y ' +
    'elimina las columnas duplicadas de identidad (Nombre/Correo/DNI/Teléfono).\n\n' +
    '⚠ ¿Ya DUPLICÓ la hoja de cálculo como respaldo? (Archivo ▸ Hacer una copia)', ui.ButtonSet.YES_NO);
  if (r1 !== ui.Button.YES) { ui.alert('Migración cancelada', 'Duplique la hoja (Archivo ▸ Hacer una copia) y vuelva a intentarlo.', ui.ButtonSet.OK); return; }
  const r2 = ui.alert('Confirmar migración', 'Se procederá a migrar los datos al modelo normalizado. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r2 !== ui.Button.YES) return;
  const msg = migrateToNormalizedClientes_();
  ui.alert('Migración completada', msg, ui.ButtonSet.OK);
}

/** Núcleo de la migración (idempotente). Devuelve un resumen legible. */
function migrateToNormalizedClientes_() {
  return guard_('migrateToNormalizedClientes_', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    if (!bs) throw new Error('No existe la hoja "Prestatarios".');
    const prevState = PropertiesService.getScriptProperties().getProperty(MIG_PROP) || '';
    if (String(bs.getRange(1, 2).getValue()).trim() === 'ID Cliente') {
      // Prestatarios ya migrado. Si una ejecución previa se interrumpió antes de
      // migrar "Pagos", termínalo ahora (reparación) y recalcula las vistas.
      if (pagosIsLegacy_(ss.getSheetByName(CFG.SHEETS.PAYMENTS))) {
        const n = migratePagosLayout_(ss);
        finishDerivedViews_(ss);
        PropertiesService.getScriptProperties().setProperty(MIG_PROP, 'columns_removed');
        return 'Reparación: se migraron ' + n + ' pago(s) al nuevo layout y se recalcularon las vistas.';
      }
      if (prevState && prevState !== 'columns_removed')
        throw new Error('Migración anterior incompleta (estado: ' + prevState + '). Restaure desde una copia nueva de la hoja original y reintente.');
      return 'La hoja ya estaba migrada. No se realizaron cambios.';
    }

    // ---- Leer TODO el layout heredado a memoria (valores y fórmulas) ----
    // Heredado Prestatarios A..U: 1=ID,2=Nombre,3=Correo,4=DNI,5=Teléfono,6=Capital,
    // 7=Plazo,8=Tasa,9=Fecha,10=Venc,11=Interés,12=Total,13=Pagado,14=Saldo,15=Estado,
    // 16=ContratoPDF,17=Enviado,18=ÚltimoAviso,19=EstadoFirma,20=FechaFirma,21=ContratoFirmado.
    const LEG = 21, lastB = bs.getLastRow();
    const legV = lastB >= 2 ? bs.getRange(2, 1, lastB - 1, LEG).getValues() : [];
    const legF = lastB >= 2 ? bs.getRange(2, 1, lastB - 1, LEG).getFormulas() : [];
    const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    // Heredado Pagos A..I: 1=IDPago,2=IDPréstamo,3=Nombre,4=Apellido,5=DNI,6=Fecha,7=Monto,8=Saldo,9=Recibo.
    const lastP = ps ? ps.getLastRow() : 0;
    const legP = (ps && lastP >= 2) ? ps.getRange(2, 1, lastP - 1, 9).getValues() : [];

    const keyOf = (dni, email, loanId) => {
      const dn = normDni_(dni), en = normEmail_(email);
      return (dn && en) ? (dn + '|' + en) : (' loan:' + String(loanId).trim());
    };
    const parseDate = v => (v instanceof Date) ? v : (v ? new Date(v) : null);

    // ---- Fase 1: construir "Clientes" (dedup DNI+Correo, préstamo más reciente gana) ----
    const cs = setupClientes_(ss);
    if (cs.getLastRow() >= 2) cs.getRange(2, 1, cs.getLastRow() - 1, CLIENTES_HEADERS.length).clearContent();
    const groups = {};
    legV.forEach(r => {
      const loanId = String(r[0] || '').trim(); if (!loanId) return;
      const key = keyOf(r[3], r[2], loanId);
      (groups[key] = groups[key] || []).push({ name: r[1], email: r[2], dni: r[3], phone: r[4], loanDate: parseDate(r[8]), loanId: loanId });
    });
    const keyToClientId = {}, clientesRows = [], now = new Date();
    let seq = 0;
    Object.keys(groups).forEach(key => {
      const rows = groups[key].slice().sort((a, b) => {
        const ta = a.loanDate ? a.loanDate.getTime() : 0, tb = b.loanDate ? b.loanDate.getTime() : 0;
        if (tb !== ta) return tb - ta;                       // más reciente primero
        return String(b.loanId).localeCompare(String(a.loanId)); // desempate por ID
      });
      const c = rows[0], clientId = 'C-' + String(++seq).padStart(4, '0');
      keyToClientId[key] = clientId;
      clientesRows.push([clientId, c.name, c.email, c.dni, c.phone, normDni_(c.dni), normEmail_(c.email), now, now]);
    });
    if (clientesRows.length) cs.getRange(2, 1, clientesRows.length, CLIENTES_HEADERS.length).setValues(clientesRows);
    invalidateClientesCache_();
    PropertiesService.getScriptProperties().setProperty(MIG_PROP, 'clientes_built');

    // ---- Fase 2: mapear préstamo -> ID Cliente ----
    const loanToClient = {};
    legV.forEach(r => {
      const loanId = String(r[0] || '').trim(); if (!loanId) return;
      loanToClient[loanId] = keyToClientId[keyOf(r[3], r[2], loanId)] || '';
    });
    PropertiesService.getScriptProperties().setProperty(MIG_PROP, 'ids_backfilled');

    // ---- Fase 3: reconstruir "Prestatarios" y "Pagos" en el nuevo layout ----
    setupBorrowers_(ss); // plantilla nueva (encabezados + fórmulas); ya leímos los datos a memoria
    let loanCount = 0;
    if (legV.length) {
      const inAB = [], inEF = [], inF = [], inMR = [];
      legV.forEach((r, i) => {
        const loanId = String(r[0] || '').trim();
        if (!loanId) { inAB.push(['', '']); inEF.push(['', '']); inF.push(['']); inMR.push(['', '', '', '', '', '']); return; }
        loanCount++;
        inAB.push([loanId, loanToClient[loanId] || '']);   // A:B = ID Préstamo, ID Cliente (Nombre/DNI en C:D son fórmulas)
        inEF.push([r[5], r[6]]);                            // E:F = Capital, Plazo
        inF.push([r[8]]);                                   // Fecha del Préstamo
        // Bloque firma/PDF (PB.PDF..SIGN_PDF). Preservar HYPERLINK heredados si existen.
        inMR.push([legF[i][15] || r[15], r[16], r[17], r[18] || '', r[19] || '', legF[i][20] || r[20]]);
      });
      const n = inAB.length;
      bs.getRange(2, PB.LOAN_ID, n, 2).setValues(inAB);
      bs.getRange(2, PB.PRINCIPAL, n, 2).setValues(inEF);
      bs.getRange(2, PB.LOAN_DATE, n, 1).setValues(inF);
      bs.getRange(2, PB.PDF, n, 6).setValues(inMR);
    }
    let payCount = 0;
    if (ps) {
      setupPayments_(ss); // ya leímos los pagos a memoria
      if (legP.length) {
        const payRows = legP.map(r => {
          const payId = String(r[0] || '').trim();
          if (!payId) return ['', '', '', '', '', ''];
          payCount++;
          return [payId, r[1], r[5], r[6], r[7], r[8]]; // ID Pago, ID Préstamo, Fecha(6), Monto(7), Saldo(8), Recibo(9)
        });
        ps.getRange(2, 1, payRows.length, 6).setValues(payRows);
      }
    }

    // Núcleo terminado (Clientes + Prestatarios + Pagos). Marcar completo AHORA:
    // así, aunque falle una vista derivada, la migración no se considera incompleta.
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty(MIG_PROP, 'columns_removed');

    // Recalcular saldos y regenerar vistas derivadas (no fatales).
    finishDerivedViews_(ss);

    return 'Migración completada: ' + clientesRows.length + ' cliente(s) creado(s), ' +
      loanCount + ' préstamo(s) y ' + payCount + ' pago(s) migrados al modelo normalizado.';
  });
}

/** ¿La hoja "Pagos" sigue en el layout heredado (con Nombre/Apellido/DNI en C/D/E)? */
function pagosIsLegacy_(ps) {
  if (!ps || ps.getLastRow() < 1) return false;
  const head = ps.getRange(1, 1, 1, Math.max(ps.getLastColumn(), 6)).getValues()[0].map(h => String(h).trim());
  return head[2] === 'Nombre' || head[3] === 'Apellido' || head[4] === 'DNI';
}

/**
 * Migra SOLO "Pagos" del layout heredado (9 col) al normalizado (6 col).
 * Self-contained: no depende de "Prestatarios". Devuelve la cantidad migrada.
 */
function migratePagosLayout_(ss) {
  ss = ss || getSS_();
  const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (!ps || !pagosIsLegacy_(ps)) return 0;
  const last = ps.getLastRow();
  const legP = last >= 2 ? ps.getRange(2, 1, last - 1, 9).getValues() : []; // leer ANTES de limpiar
  setupPayments_(ss); // limpia validaciones + encabezado 6-col
  let n = 0;
  if (legP.length) {
    const payRows = legP.map(r => {
      const payId = String(r[0] || '').trim();
      if (!payId) return ['', '', '', '', '', ''];
      n++;
      return [payId, r[1], r[5], r[6], r[7], r[8]]; // ID Pago, ID Préstamo, Fecha(6), Monto(7), Saldo(8), Recibo(9)
    });
    ps.getRange(2, 1, payRows.length, 6).setValues(payRows);
  }
  return n;
}

/** Recalcula saldos y regenera vistas derivadas. Cada paso es NO fatal (registra en "Errores"). */
function finishDerivedViews_(ss) {
  ss = ss || getSS_();
  const safeStep = (label, fn) => { try { fn(); } catch (e) { logError_('migración:' + label, e); } };
  SpreadsheetApp.flush();
  safeStep('updateAllOutstanding_', () => updateAllOutstanding_());
  safeStep('refreshPaymentNames_', () => refreshPaymentNames_());
  safeStep('setupSummary_', () => setupSummary_(ss));
  safeStep('setupPanel_', () => setupPanel_(ss));
  safeStep('setupStats_', () => setupStats_(ss));
  safeStep('setupStatements_', () => setupStatements_(ss));
  safeStep('setupAgreement_', () => setupAgreement_(ss));
  safeStep('rebuildLateSheet_', () => rebuildLateSheet_(ss));
  safeStep('rebuildRemindersSheet_', () => rebuildRemindersSheet_(ss));
  safeStep('applyDailyVisibility_', () => applyDailyVisibility_(ss));
  SpreadsheetApp.flush();
}

/* ============ CLIENTES (modelo normalizado, identidad) ============ */
// Cache de "Clientes" por ejecución: id -> {id,name,email,dni,phone}.
let _clientesCache = null;
function clientesCache_() {
  if (_clientesCache) return _clientesCache;
  const map = {};
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS);
  if (sh && sh.getLastRow() >= 2) {
    // Las columnas 1..5 (ID/Nombre/Correo/DNI/Teléfono) son idénticas en ambos
    // esquemas de "Clientes". Los campos extra (CUIL, Dirección, Forma de Pago)
    // se resuelven por NOMBRE de encabezado, así funciona con el esquema de 9 o
    // de 13 columnas sin depender de índices que chocan entre versiones.
    const lastCol = Math.max(sh.getLastColumn(), PC.PHONE);
    const hdr = headerMap_(sh);
    const cCuil = hdr['CUIL'] || 0, cDir = hdr['Dirección'] || 0, cPay = hdr[CLIENTE_PAY_HEADER] || 0;
    const at = (r, c) => (c && c <= r.length) ? String(r[c - 1] || '').trim() : '';
    sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues().forEach(r => {
      const id = String(r[PC.CLIENT_ID - 1] || '').trim();
      if (id) map[id] = {
        id: id, name: r[PC.NAME - 1], email: r[PC.EMAIL - 1], dni: r[PC.DNI - 1], phone: r[PC.PHONE - 1],
        cuil: at(r, cCuil), direccion: at(r, cDir), payMethod: at(r, cPay),
      };
    });
  }
  _clientesCache = map;
  return map;
}
// Encabezado (por nombre) donde se guarda la forma de pago elegida por el prestatario.
const CLIENTE_PAY_HEADER = 'Forma de Pago';
/** Fila (1-based) de un cliente por su ID en la hoja "Clientes"; 0 si no está. */
function clienteRowById_(sh, clienteId) {
  const id = String(clienteId || '').trim(); if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, PC.CLIENT_ID, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) return i + 2;
  return 0;
}
/** Nº de columna de un encabezado en "Clientes"; si falta y create=true, la agrega al final. */
function clientesHeaderCol_(sh, header, create) {
  const m = headerMap_(sh);
  if (m[header]) return m[header];
  if (!create) return 0;
  const col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(header)
    .setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff').setWrap(true);
  return col;
}
/** Escribe (row, columna-por-encabezado) en cualquier hoja; crea la columna al final si falta. */
function setSheetFieldByHeader_(sh, row, header, value) {
  const m = headerMap_(sh);
  let col = m[header];
  if (!col) {
    col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue(header).setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
  }
  sh.getRange(row, col).setValue(value);
  return col;
}
/** Escribe un campo "extra" del cliente por nombre de encabezado (creando la columna si falta). */
function setClienteFieldByHeader_(clienteId, header, value) {
  if (value == null || String(value).trim() === '') return;
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS); if (!sh) return;
  const row = clienteRowById_(sh, clienteId); if (!row) return;
  const col = clientesHeaderCol_(sh, header, true);
  sh.getRange(row, col).setValue(value);
  invalidateClientesCache_();
}
function invalidateClientesCache_() {
  _clientesCache = null;
  // Mantiene coherente el caché de filas de "Clientes" en Validaciones.gs:
  // todo sitio que ya invalida este caché invalida ambos.
  if (typeof invalidateClientesRows_ === 'function') invalidateClientesRows_();
}
function getClienteById_(clienteId) {
  const id = String(clienteId || '').trim();
  return id ? (clientesCache_()[id] || null) : null;
}
function normDni_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }
function normEmail_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function nextClienteId_() {
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS); if (!sh) return 'C-0001';
  const last = sh.getLastRow(); let max = 0;
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => { const m = /^C-(\d+)$/.exec(String(r[0]).trim()); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return 'C-' + String(max + 1).padStart(4, '0');
}
function firstEmptyClienteRow_(sh) {
  const last = sh.getLastRow(); if (last < 2) return 2;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === '') return i + 2;
  return last + 1;
}
/** Busca un cliente por DNI+Correo (ambos normalizados y no vacíos). */
function findClienteByDniEmail_(dni, email) {
  const dn = normDni_(dni), en = normEmail_(email);
  if (!dn || !en) return null;
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS); if (!sh || sh.getLastRow() < 2) return null;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, PC.EMAIL_NORM).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][PC.DNI_NORM - 1]).trim() === dn && String(data[i][PC.EMAIL_NORM - 1]).trim() === en) {
      const r = data[i];
      return { id: String(r[PC.CLIENT_ID - 1]).trim(), name: r[PC.NAME - 1], email: r[PC.EMAIL - 1], dni: r[PC.DNI - 1], phone: r[PC.PHONE - 1], row: i + 2 };
    }
  }
  return null;
}
/** Devuelve true si algún cliente ya tiene ese DNI (normalizado). */
function clienteExistsByDni_(dni) {
  const dn = normDni_(dni); if (!dn) return false;
  const sh = getSS_().getSheetByName(CFG.SHEETS.CLIENTS); if (!sh || sh.getLastRow() < 2) return false;
  const data = sh.getRange(2, PC.DNI_NORM, sh.getLastRow() - 1, 1).getValues();
  return data.some(r => String(r[0]).trim() === dn);
}
function appendCliente_(name, email, dni, phone) {
  const sh = setupClientes_();
  const id = nextClienteId_(), now = new Date(), row = firstEmptyClienteRow_(sh);
  sh.getRange(row, 1, 1, CLIENTES_HEADERS.length)
    .setValues([[id, name, email, dni, phone, normDni_(dni), normEmail_(email), now, now]]);
  invalidateClientesCache_();
  return id;
}
/** Reutiliza el cliente si coincide DNI+Correo; si no, crea uno nuevo. Devuelve ID Cliente. */
function getOrCreateCliente_(name, email, dni, phone) {
  const existing = findClienteByDniEmail_(dni, email);
  return existing ? existing.id : appendCliente_(name, email, dni, phone);
}
function nbAppend_(nb, values) {
  const last = Math.max(nb.getLastRow(), 1), dataRows = Math.max(last - 1, 1);
  const colB = nb.getRange(2, 2, dataRows, 1).getValues(); let row = last + 1;
  for (let i = 0; i < colB.length; i++) if (String(colB[i][0]).trim() === '') { row = i + 2; break; }
  nb.getRange(row, 1, 1, values.length).setValues([values]); return row;
}
/** Tasa a partir del plazo, tolerante a días (15/30/60) o meses (15/1/2). */
function loanRateForTerm_(term) {
  const t = Number(term) || 0;
  if (typeof termRateDays_ === 'function') { const r = termRateDays_(t); if (r != null) return r; }
  if (typeof termRate_ === 'function') { const r = termRate_(t); if (r != null) return r; }
  return 0;
}
function readLoan_(sh, row) {
  // Columnas por NOMBRE de encabezado (robusto al layout migrado o nuevo). Interés y
  // Total se CALCULAN si las celdas están en blanco (p. ej. préstamo recién aprobado
  // cuya fila aún no tiene las fórmulas), para que el contrato nunca muestre $0.
  const H = headerIndex_(sh);
  const col = (names, fb) => colByAny_(H, names) || fb;
  const idC = col(['ID Préstamo', 'ID Prestamo'], PB.LOAN_ID);
  const cliC = col(['ID Cliente'], PB.CLIENT_ID);
  const capC = col(['Capital'], PB.PRINCIPAL);
  const plzC = col(['Plazo (días)', 'Plazo (dias)', 'Plazo (meses)', 'Plazo'], PB.TERM);
  const tasC = col(['Tasa'], PB.RATE);
  const fecC = col(['Fecha del Préstamo', 'Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo'], PB.LOAN_DATE);
  const venC = col(['Fecha de Vencimiento', 'Vencimiento'], PB.DUE);
  const intC = col(['Interés', 'Interes'], PB.INTEREST);
  const totC = col(['Total a Pagar', 'Total a pagar'], PB.TOTAL);
  const lastCol = Math.max(idC, cliC, capC, plzC, tasC, fecC, venC, intC, totC);
  const v = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const g = c => v[c - 1];
  const clientId = String(g(cliC) || '').trim();
  const c = getClienteById_(clientId) || {};
  const term = Number(g(plzC)) || 0;
  const ld = g(fecC) instanceof Date ? g(fecC) : new Date(g(fecC));
  const principal = Number(g(capC)) || 0;
  let rate = Number(g(tasC)) || 0; if (!rate) rate = loanRateForTerm_(term);
  let interest = Number(g(intC)) || 0; if (!interest) interest = round2_(principal * rate);
  let totalDue = Number(g(totC)) || 0; if (!totalDue) totalDue = round2_(principal + interest);
  return {
    loanId: g(idC), clientId: clientId,
    name: c.name || '', email: c.email || '', dni: c.dni || '', phone: c.phone || '',
    cuil: c.cuil || '', direccion: c.direccion || '', payMethod: c.payMethod || '',
    principal: principal, term: term,
    loanDate: ld,
    dueDate: g(venC) instanceof Date ? g(venC) : dueDateForTerm_(term || 1, ld),
    interest: interest, totalDue: totalDue,
  };
}
function findLoanById_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][0]).trim() === loanId) return readLoan_(sh, i + 2);
  return null;
}
function loanPayments_(loanId) {
  // Pagos: B=ID Préstamo, C=Fecha, D=Monto.
  const sh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS); if (sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, PP.LOAN_ID, sh.getLastRow() - 1, 3).getValues(), out = [];
  data.forEach(r => { if (String(r[0]).trim() === loanId && typeof r[2] === 'number' && r[2] !== 0) out.push({ date: r[1] instanceof Date ? r[1] : new Date(), amount: r[2] }); });
  return out;
}
function nextLoanId_() {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow(); let max = 0;
  if (last >= 2) bs.getRange(2, 1, last - 1, 1).getValues().forEach(r => { const m = /^L-(\d+)$/.exec(String(r[0]).trim()); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return 'L-' + String(max + 1).padStart(4, '0');
}
function firstEmptyBorrowerRow_(bs) {
  const vals = bs.getRange(2, 1, CFG.MAX_ROWS, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === '') return i + 2;
  return bs.getLastRow() + 1;
}
function getSetting_(key) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return '';
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const r of data) if (String(r[0]).trim() === key) return r[1];
  return '';
}
function spreadsheetParentFolder_() {
  const p = DriveApp.getFileById(getSS_().getId()).getParents();
  return p.hasNext() ? p.next() : DriveApp.getRootFolder();
}
function borrowerMainFolder_() {
  const parent = spreadsheetParentFolder_(), it = parent.getFoldersByName(CFG.MAIN_FOLDER);
  return it.hasNext() ? it.next() : parent.createFolder(CFG.MAIN_FOLDER);
}
function borrowerFolder_(folderName) {
  if (TEST_MODE) return _fakeFolder_(folderName);
  const main = borrowerMainFolder_(), it = main.getFoldersByName(folderName);
  return it.hasNext() ? it.next() : main.createFolder(folderName);
}
function borrowerFolderName_(fullName, dni) {
  return sanitizeName_([firstName_(fullName), lastName_(fullName), String(dni || '').trim()].filter(String).join(' ')) || sanitizeName_(fullName);
}
function savePhoto_(folder, blob, baseName) {
  if (TEST_MODE) return 'test://photo/' + sanitizeName_(baseName || 'foto');
  try {
    if (!blob || !blob.getBytes || blob.getBytes().length === 0) return '';
    const ct = blob.getContentType() || '', ext = ct.indexOf('pdf') >= 0 ? 'pdf' : (ct.split('/')[1] || 'jpg');
    return folder.createFile(blob.setName(sanitizeName_(baseName) + '.' + ext)).getUrl();
  } catch (e) { logError_('savePhoto_', e); return ''; }
}
function extractDriveIds_(cell) {
  const s = String(cell == null ? '' : cell), ids = []; let m;
  const re = /(?:id=|\/d\/)([-\w]{20,})/g; while ((m = re.exec(s))) ids.push(m[1]);
  if (!ids.length) { const re2 = /[-\w]{25,}/g; while ((m = re2.exec(s))) ids.push(m[0]); }
  return ids;
}
function hasFile_(b) { return !!(b && b.getBytes && b.getBytes().length > 0); }
/* ===================== SEAMS DE PRUEBA (Tests.gs) =====================
 * Inertes en operación normal (TEST_MODE = false). El runner de pruebas los
 * activa para correr contra una hoja temporal, capturar correos, evitar
 * escrituras a Drive y stubbear el BCRA. NO borran ni tocan datos reales. */
let TEST_MODE = false;          // interruptor global de modo prueba
let _ssOverride = null;         // hoja de cálculo temporal para las pruebas
const _emailOutbox = [];        // correos capturados (en vez de enviarse)
let _bcraStub = null;           // respuesta BCRA simulada (objeto o función(path))
/** Archivo PDF simulado: registra el HTML y no escribe en Drive. */
function _fakeFile_(filename, html) {
  return {
    _test: true, _html: html, _name: filename,
    getUrl: function () { return 'test://file/' + encodeURIComponent(filename); },
    getName: function () { return this._name; },
    setName: function (n) { this._name = n; return this; },
    getId: function () { return 'test-id'; },
    getAs: function () { return Utilities.newBlob(html, 'text/html', filename); },
  };
}
/** Carpeta simulada: no crea nada en Drive. */
function _fakeFolder_(name) {
  return {
    _test: true, _name: name,
    getName: function () { return name; },
    getUrl: function () { return 'test://folder/' + encodeURIComponent(name); },
    getId: function () { return 'test-folder-id'; },
    createFile: function (blob) { return _fakeFile_((blob && blob.getName && blob.getName()) || 'archivo', ''); },
  };
}
function getSS_() {
  if (_ssOverride) return _ssOverride;
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID'));
}
function getOrCreate_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
// allowInvalid(true): el selector de fecha es una AYUDA, no bloquea escrituras
// del script ni migraciones masivas (evita "viola las reglas de validación").
function datePicker_() { return SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Elija una fecha.').build(); }
function cc_(range, text, color) { return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build(); }
/** Regla: texto rojo y negrita cuando el valor de la celda es negativo (sobregiro del fondo). */
function redIfNegativeRule_(range) { return SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#cc0000').setBold(true).setRanges([range]).build(); }
function addMonths_(date, n) { const d = new Date(date.getFullYear(), date.getMonth(), date.getDate()), day = d.getDate(); d.setMonth(d.getMonth() + n); if (d.getDate() < day) d.setDate(0); return d; }
function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
/* ---- Plazos (términos del préstamo) ----
 * El "Plazo" se codifica como número entero: 15 = 15 días (25%), 1 = 1 mes (50%),
 * 2 = 2 meses (100%). El interés es fijo por el plazo. */
/** Interés fijo del plazo como fracción, o null si el plazo es inválido. */
function termRate_(term) { return term === 15 ? 0.25 : term === 1 ? 0.5 : term === 2 ? 1 : term === 3 ? 1 : null; }
/** Etiqueta legible del plazo (p. ej. "15 días", "1 mes(es)"). */
function termLabel_(term) { return term === 15 ? '15 días' : term + ' mes(es)'; }
/** Porcentaje de interés del plazo como texto (p. ej. "25%"), o '' si es inválido. */
function termRatePct_(term) { const r = termRate_(term); return r === null ? '' : round2_(r * 100) + '%'; }
/** Fecha de vencimiento del plazo a partir de la fecha del préstamo (por días calendario). */
function dueDateForTerm_(term, loanDate) { return addDays_(loanDate, termDays_(term)); }
/**
 * Plazo en días calendario. Robusto a las dos convenciones del proyecto:
 * modelo mes (15=15 días, 1=1 mes≈30, 2=2 meses≈60) y modelo día (15/30/60).
 */
function termDays_(term) {
  term = Number(term);
  return term === 15 ? 15 : term === 30 ? 30 : term === 60 ? 60 : term === 90 ? 90 :
    term === 1 ? 30 : term === 2 ? 60 : term === 3 ? 90 : (term || 0);
}
/** Tasa fija del período como fracción, resolviendo ambas convenciones de plazo. */
function loanRate_(term) {
  const r = termRate_(term);
  if (r != null) return r;
  try { const rd = termRateDays_(term); if (rd != null) return rd; } catch (e) {}
  return 0;
}
/** Etiqueta de plazo en días calendario (robusta a ambas convenciones). */
function loanTermLabel_(term) { return termDays_(term) + ' días calendario'; }
/** Porcentaje de interés del plazo como texto, resolviendo ambas convenciones. */
function loanRatePct_(term) { return round2_(loanRate_(term) * 100) + '%'; }
/** Porcentaje legible con separador local (p. ej. 22711 -> "22.711 %"). */
function pctText_(frac, dp) {
  const v = frac * 100;
  const digits = (typeof dp === 'number') ? dp : (v >= 1000 ? 0 : v >= 1 ? 2 : 4);
  try { return v.toLocaleString(CFG.LOCALE, { minimumFractionDigits: digits, maximumFractionDigits: digits }) + ' %'; }
  catch (e) { return v.toFixed(digits) + ' %'; }
}
/**
 * Disclosures de costo del crédito (Ley 24.240 / comunicación BCRA): tasa diaria,
 * TNA, TEA, tasa mensual y CFT. Sin comisiones ⇒ CFT = TEA. Base 365 días.
 * La TEA y la mensual capitalizan la tasa FIJA del período: (1+tasaPeríodo)^(365/días)−1.
 * Ej. 25% a 15 días ⇒ diaria 1,6667%, TNA 608,33%, TEA 22.711%, mensual 56,25%.
 */
function disclosureRates_(loan) {
  const days = termDays_(loan.term) || 1;
  const ratePeriod = loanRate_(loan.term) || 0;   // fracción fija del período (0.25 = 25%)
  const daily = ratePeriod / days;
  const tea = Math.pow(1 + ratePeriod, 365 / days) - 1;
  const monthly = Math.pow(1 + ratePeriod, 30 / days) - 1;
  return {
    days: days,
    periodPct: pctText_(ratePeriod),
    dailyPct: pctText_(daily, 4),
    tnaPct: pctText_(daily * 365),
    teaPct: pctText_(tea),
    monthlyPct: pctText_(monthly),
    cftPct: pctText_(tea),
  };
}
/** Avanza la fecha al próximo día hábil (salta sábado y domingo). */
function rollToBusinessDay_(date) {
  if (!(date instanceof Date)) date = new Date(date);
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
  return d;
}
function sanitizeName_(s) { return String(s == null ? '' : s).replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim(); }
function firstName_(f) { return String(f == null ? '' : f).trim().split(/\s+/)[0] || ''; }
function lastName_(f) { const p = String(f == null ? '' : f).trim().split(/\s+/); return p.length > 1 ? p.slice(1).join(' ') : ''; }
function fmtMoney_(n) { try { return Number(n).toLocaleString(CFG.LOCALE, { style: 'currency', currency: CFG.CURRENCY_CODE }); } catch (e) { return '$ ' + round2_(Number(n)).toFixed(2); } }
const MESES_ES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
// Fechas para contrato y comunicaciones al cliente: año, mes (en palabra) y día. Ej.: "2026 Agosto 11".
function fmtDate_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  const tz = Session.getScriptTimeZone();
  const y = Utilities.formatDate(d, tz, 'yyyy');
  const m = Number(Utilities.formatDate(d, tz, 'MM'));
  const day = Utilities.formatDate(d, tz, 'd');
  return y + ' ' + MESES_ES[m - 1] + ' ' + day;
}
function fmtDateTime_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return fmtDate_(d) + ' · ' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* ==================================================================
 *  DATOS DE PRUEBA  —  crear / eliminar (create & tear down)
 * ------------------------------------------------------------------
 *  Agrega 40 prestatarios de PRUEBA en distintos estados y casos
 *  límite (ACTIVO, VENCIDO, PAGADO). Los préstamos de prueba usan IDs
 *  con prefijo "T-XXXX" (los reales usan "L-XXXX"), de modo que la
 *  eliminación NUNCA toca datos reales. No se envían correos.
 * ================================================================== */
const TEST_PREFIX = 'T-';               // IDs de préstamos de prueba
const TEST_EMAIL_DOMAIN = 'prueba.test'; // dominio ficticio (no envía correos)

/** ¿Es un ID de préstamo de prueba? (T-0001, etc.) */
function isTestId_(v) { return /^T-\d+/i.test(String(v == null ? '' : v).trim()); }

function countTestBorrowers_(bs) {
  if (!bs || bs.getLastRow() < 2) return 0;
  return bs.getRange(2, 1, bs.getLastRow() - 1, 1).getValues().filter(r => isTestId_(r[0])).length;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { const x = startOfDay_(d); x.setDate(x.getDate() + n); return x; }

/* ---------- CREAR ---------- */
function seedTestBorrowersConfirm() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Datos de prueba',
    'Se agregarán 40 prestatarios de PRUEBA (IDs "T-XXXX") en distintos estados: '
    + 'ACTIVO, VENCIDO y PAGADO, incluyendo casos límite (vence hoy, vencido ayer, '
    + 'pago parcial, sobrepago, capital mínimo/máximo, nombres con acentos).\n\n'
    + 'No se envían correos. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) seedTestBorrowers();
}

function seedTestBorrowers() {
  return guard_('seedTestBorrowers', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    if (!bs || !ps) throw new Error('Faltan las hojas base. Ejecute ① Configurar primero.');

    const existing = countTestBorrowers_(bs);
    if (existing > 0) throw new Error('Ya existen ' + existing + ' prestatarios de prueba. '
      + 'Use "Eliminar prestatarios de prueba" antes de volver a crearlos.');

    let specs = buildTestSpecs_();
    const startRow = firstEmptyBorrowerRow_(bs);
    const room = (CFG.MAX_ROWS + 1) - startRow + 1;
    if (room <= 0) throw new Error('No hay filas libres en "Prestatarios".');
    if (specs.length > room) specs = specs.slice(0, room);

    const rowsAB = [];   // A:B (ID Préstamo, ID Cliente) — Nombre/DNI en C:D son fórmulas
    const rowsEF = [];   // E:F (Capital, Plazo)
    const rowsF = [];    // Fecha del Préstamo
    const payRows = [];  // filas para la hoja de Pagos (6 columnas, modelo normalizado)
    const today = startOfDay_(new Date());

    specs.forEach(function (s, i) {
      const loanId = TEST_PREFIX + String(i + 1).padStart(4, '0');
      const rate = termRate_(s.term);
      const totalDue = round2_(s.capital * (1 + rate));
      const clienteId = getOrCreateCliente_(s.name, s.email, s.dni, s.phone);
      rowsAB.push([loanId, clienteId]);
      rowsEF.push([s.capital, String(s.term)]); // Plazo como texto para respetar la validación de la columna
      rowsF.push([s.loanDate]);
      if (s.payFrac > 0) {
        const amount = round2_(totalDue * s.payFrac);
        let payDate = addDays_(s.loanDate, 5);
        if (payDate.getTime() > today.getTime()) payDate = today; // nunca pagar en el futuro
        payRows.push(['TP-' + String(i + 1).padStart(4, '0'), loanId, payDate, amount,
          Math.max(0, round2_(totalDue - amount)), '']);
      }
    });

    bs.getRange(startRow, PB.LOAN_ID, rowsAB.length, 2).setValues(rowsAB);
    bs.getRange(startRow, PB.PRINCIPAL, rowsEF.length, 2).setValues(rowsEF);
    bs.getRange(startRow, PB.LOAN_DATE, rowsF.length, 1).setValues(rowsF);
    bs.getRange(startRow, PB.PRINCIPAL, rowsEF.length, 1).setNumberFormat(CFG.CURRENCY_FMT);
    bs.getRange(startRow, PB.LOAN_DATE, rowsF.length, 1).setNumberFormat('yyyy-mm-dd');

    if (payRows.length) {
      const pstart = ps.getLastRow() + 1;
      ps.getRange(pstart, 1, payRows.length, 6).setValues(payRows);
      ps.getRange(pstart, PP.DATE, payRows.length, 1).setNumberFormat('yyyy-mm-dd');
      ps.getRange(pstart, PP.AMOUNT, payRows.length, 2).setNumberFormat(CFG.CURRENCY_FMT);
    }

    SpreadsheetApp.flush();
    try { refreshAll(true); } catch (e) { logError_('seedTestBorrowers/refreshAll', e); }

    // Resumen por estado para la confirmación.
    const counts = summarizeTestStates_(bs);
    SpreadsheetApp.getUi().alert('Datos de prueba creados',
      'Se agregaron ' + rowsAG.length + ' prestatarios de prueba y ' + payRows.length + ' pagos.\n\n'
      + 'ACTIVO: ' + counts.ACTIVO + '\nVENCIDO: ' + counts.VENCIDO + '\nPAGADO: ' + counts.PAGADO,
      SpreadsheetApp.getUi().ButtonSet.OK);
  });
}

/** Cuenta los estados actuales de las filas de prueba (columna Estado). */
function summarizeTestStates_(bs) {
  const out = { ACTIVO: 0, VENCIDO: 0, PAGADO: 0, otros: 0 };
  if (!bs || bs.getLastRow() < 2) return out;
  const last = bs.getLastRow();
  const rows = bs.getRange(2, 1, last - 1, PB.STATE).getValues();
  rows.forEach(function (r) {
    if (!isTestId_(r[0])) return;
    const st = String(r[PB.STATE - 1]).trim();
    if (out[st] === undefined) out.otros++; else out[st]++;
  });
  return out;
}

/* ---------- ELIMINAR (tear down) ---------- */
function removeTestBorrowersConfirm() {
  const ui = SpreadsheetApp.getUi();
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  const n = countTestBorrowers_(bs);
  if (n === 0) { ui.alert('Datos de prueba', 'No hay prestatarios de prueba para eliminar.', ui.ButtonSet.OK); return; }
  const r = ui.alert('Datos de prueba',
    'Se eliminarán ' + n + ' prestatarios de prueba (IDs "T-XXXX") y todos sus pagos. '
    + 'Los datos reales ("L-XXXX") no se tocan.\n\n¿Continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) removeTestBorrowers();
}

function removeTestBorrowers() {
  return guard_('removeTestBorrowers', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    let removed = 0, payRemoved = 0;

    // Prestatarios: limpiar solo las celdas de entrada (conserva la plantilla de fórmulas).
    if (bs && bs.getLastRow() >= 2) {
      const last = bs.getLastRow();
      const ids = bs.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (!isTestId_(ids[i][0])) continue;
        const row = i + 2;
        bs.getRange(row, PB.LOAN_ID, 1, 2).clearContent();    // A:B ID, Cliente (C:D Nombre/DNI son fórmulas)
        bs.getRange(row, PB.PRINCIPAL, 1, 2).clearContent();  // E:F Capital, Plazo
        bs.getRange(row, PB.LOAN_DATE).clearContent();        // Fecha del préstamo
        bs.getRange(row, PB.PDF, 1, 6).clearContent();        // bloque contrato/avisos/firma
        removed++;
      }
    }

    // Pagos: eliminar filas de prueba (de abajo hacia arriba para no desplazar índices).
    if (ps && ps.getLastRow() >= 2) {
      const pids = ps.getRange(2, PP.LOAN_ID, ps.getLastRow() - 1, 1).getValues();
      for (let i = pids.length - 1; i >= 0; i--) {
        if (isTestId_(pids[i][0])) { ps.deleteRow(i + 2); payRemoved++; }
      }
    }

    // Clientes de prueba (correo del dominio ficticio): eliminar de abajo hacia arriba.
    const cs = ss.getSheetByName(CFG.SHEETS.CLIENTS);
    if (cs && cs.getLastRow() >= 2) {
      const cmail = cs.getRange(2, PC.EMAIL, cs.getLastRow() - 1, 1).getValues();
      for (let i = cmail.length - 1; i >= 0; i--) {
        if (String(cmail[i][0]).trim().toLowerCase().endsWith('@' + TEST_EMAIL_DOMAIN)) cs.deleteRow(i + 2);
      }
      invalidateClientesCache_();
    }

    SpreadsheetApp.flush();
    try { refreshAll(true); } catch (e) { logError_('removeTestBorrowers/refreshAll', e); }
    SpreadsheetApp.getUi().alert('Datos de prueba eliminados',
      'Se eliminaron ' + removed + ' prestatarios de prueba y ' + payRemoved + ' pagos.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  });
}

/* ---------- CATÁLOGO DE 40 CASOS ---------- */
/**
 * Devuelve 40 especificaciones de prueba con estados y casos límite variados.
 * Cada spec: { name, email, dni, phone, capital, term(1|2), loanDate, payFrac }.
 *  - payFrac 0 = sin pagos; <1 = pago parcial; 1 = saldado; >1 = sobrepago.
 *  - Estado resultante (col O): PAGADO si saldo≤0; VENCIDO si venció; si no ACTIVO.
 *  - Los PAGADO se mantienen NO vencidos para que el recargo por mora no los reactive.
 */
function buildTestSpecs_() {
  const today = startOfDay_(new Date());
  const dueToday = t => addMonths_(today, -t);            // vence exactamente hoy
  const dueYest = t => addDays_(addMonths_(today, -t), -1); // venció ayer
  const ago = d => addDays_(today, -d);                    // fecha del préstamo = hoy - d

  // [nombre, capital, plazo, fechaPréstamo, payFrac, etiqueta]
  const raw = [
    /* ---- ACTIVO (vigente): vence en el futuro o justo hoy ---- */
    ['José Martínez',                 15000,  1, ago(0),        0,    'nuevo hoy'],
    ['Lucía Fernández',               30000,  2, ago(0),        0,    'nuevo hoy (2m)'],
    ['Carlos Gómez',                  12000,  1, ago(5),        0,    'vigente'],
    ['Ana María Suárez',              20000,  1, ago(10),       0.30, 'vigente c/pago parcial'],
    ['Diego Romero',                  18000,  1, ago(20),       0,    'vigente, vence pronto'],
    ['Sofía Ramírez',                 45000,  2, ago(15),       0.50, 'vigente 2m parcial'],
    ['Martín Acosta',                 60000,  2, ago(40),       0,    'vigente 2m'],
    ['Valentina Torres',              25000,  1, dueToday(1),   0,    'límite: vence HOY'],
    ['Nicolás Herrera',               50000,  2, dueToday(2),   0,    'límite: vence HOY (2m)'],
    ['Micaela Rossi',                     1,  1, ago(0),        0,    'capital mínimo ($1)'],
    ['Fernando Del Río',            5000000,  2, ago(3),        0,    'capital máximo'],
    ['María José Rodríguez de la Fuente', 22000, 1, ago(7),     0.10, 'nombre largo c/acentos'],

    /* ---- VENCIDO (en mora): la fecha de vencimiento ya pasó ---- */
    ['Roberto Silva',                 15000,  1, dueYest(1),    0,    'límite: venció AYER'],
    ['Paula Cabrera',                 28000,  2, dueYest(2),    0,    'límite: venció AYER (2m)'],
    ['Gonzalo Medina',                17000,  1, ago(45),       0,    'vencido ~15 días'],
    ['Camila Ortiz',                  19000,  1, ago(60),       0,    'vencido ~30 días'],
    ['Andrés Molina',                 21000,  1, ago(90),       0,    'vencido ~60 días'],
    ['Florencia Vega',                40000,  1, ago(120),      0.50, 'vencido c/pago parcial'],
    ['Sebastián Ríos',                55000,  2, ago(100),      0,    'vencido 2m'],
    ['Julieta Sosa',                  70000,  2, ago(150),      0.25, 'vencido 2m parcial'],
    ['Emiliano Castro',               16000,  1, ago(200),      0,    'vencido, muy atrasado'],
    ['Agustina Peña',                 80000,  2, ago(365),      0,    'vencido ~1 año'],
    ['Tomás Aguirre',                   500,  1, ago(50),       0,    'capital chico vencido'],
    ['Brenda Ledesma',              2000000,  2, ago(80),       0,    'capital grande vencido'],
    ['Ramón Ñáñez',                   14000,  1, ago(70),       0,    'vencido c/acentos'],
    ['Lucas Benítez',                 26000,  1, ago(55),       0.90, 'vencido, casi saldado'],

    /* ---- PAGADO (saldado): no vencido + pago total o sobrepago ---- */
    ['Patricia Núñez',                15000,  1, ago(10),       1.00, 'saldado'],
    ['Hernán Ibáñez',                 30000,  2, ago(5),        1.00, 'saldado 2m'],
    ['Rocío Franco',                  12000,  1, ago(2),        1.00, 'saldado reciente'],
    ['Federico Luna',                 18000,  1, ago(0),        1.00, 'saldado el mismo día'],
    ['Daniela Vera',                  20000,  1, ago(15),       1.05, 'saldado c/sobrepago'],
    ['Ignacio Paz',                   24000,  1, ago(25),       1.00, 'saldado, vencía pronto'],
    ['Antonella Bruno',               35000,  2, ago(20),       1.00, 'saldado 2m'],
    ['Mateo Correa',                      1,  1, ago(1),        1.00, 'saldado capital $1'],
    ['Guadalupe Ferreyra',          3000000,  2, ago(4),        1.00, 'saldado capital grande'],
    ['Ángela Muñoz',                  21000,  1, ago(8),        1.00, 'saldado c/acentos'],
    ['Bautista Godoy',                33000,  2, ago(12),       1.00, 'saldado 2m'],
    ['Renata Ponce',                  16000,  1, ago(6),        1.00, 'saldado'],
    ['Thiago Cardozo',                27000,  2, ago(3),        1.02, 'saldado c/sobrepago'],
    ['Delfina Ojeda',                 19000,  1, ago(9),        1.00, 'saldado'],
  ];

  return raw.map(function (r, i) {
    const name = r[0];
    return {
      name: name,
      email: slugEmail_(name, i),
      dni: 90000001 + i,
      phone: '+54 9 11 ' + String(40000000 + i),
      capital: r[1],
      term: r[2],
      loanDate: r[3],
      payFrac: r[4],
    };
  });
}

/** Correo ficticio único y determinístico a partir del nombre (sin acentos). */
function slugEmail_(name, i) {
  let base = String(name).toLowerCase();
  try { base = base.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { /* normalize no disponible */ }
  const slug = base.replace(/[^a-z0-9]+/g, '.').replace(/^\.+|\.+$/g, '');
  return slug + '.' + (i + 1) + '@' + TEST_EMAIL_DOMAIN;
}

/* ==================================================================
 *  AVISO DE MORA (correo desde "Pagos Atrasados")
 * ================================================================== */
/** Envía un aviso de mora en español para un préstamo y marca la fecha de envío. */
function sendOverdueNotice_(loanId) {
  return guard_('sendOverdueNotice_', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo ' + loanId + '.');
    if (!loan.email) throw new Error('El préstamo ' + loanId + ' no tiene correo cargado.');
    const today = new Date();
    const rate = termRate_(loan.term);
    const pays = loanPayments_(loanId);
    const outstanding = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, today, new Date(9999, 0, 1), undefined, undefined, undefined, loanSchedule_(loan));
    const days = daysLate_(loan.dueDate, today);
    const feeDay = round2_(loan.totalDue * lateFeeRate_());
    const subject = 'Aviso de mora — préstamo ' + loan.loanId;
    const intro = 'Su préstamo ' + loan.loanId + ' venció el ' + fmtDate_(loan.dueDate) + ' (' + days +
      ' día(s) de atraso). Se aplica un recargo por mora del ' + lateFeePctText_() + ' por día (' +
      fmtMoney_(feeDay) + ' por día) sobre el total a devolver, con tope del ' + moraCapPctText_() + ' del capital.';
    sendBrandedEmail_(loan.email, subject, intro + ' Saldo actual: ' + fmtMoney_(outstanding) + '.',
      '<p>Estimado/a ' + esc_(loan.name) + ',</p>' +
      '<p>' + esc_(intro) + '</p>' +
      '<p><b>Saldo pendiente: ' + fmtMoney_(outstanding) + '</b></p>' +
      '<p>Le solicitamos regularizar su situación a la brevedad para evitar mayores recargos. ' +
      'Si ya realizó el pago, por favor ignore este mensaje.</p>');
    // Registrar la fecha del último aviso en "Prestatarios" (col "Último Aviso") para que persista.
    const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
    if (last >= 2) {
      const ids = bs.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === loanId) { bs.getRange(i + 2, PB.NOTICE).setValue(today).setNumberFormat('yyyy-mm-dd'); break; }
      }
    }
    return { date: today, message: 'Aviso de mora enviado a ' + loan.email + ' (' + days + ' día(s) de atraso).' };
  });
}

/* ==================================================================
 *  HOJA "SALDADOS" + MOVER PRÉSTAMOS PAGADOS
 * ================================================================== */
const CLEARED_HEADERS = ['ID Préstamo', 'Prestatario', 'DNI', 'Correo', 'Teléfono', 'Capital',
  'Plazo (meses)', 'Interés', 'Total a Pagar', 'Total Pagado', 'Fecha del Préstamo',
  'Fecha de Vencimiento', 'Fecha de Saldado'];

/** Crea (una vez) la hoja "Saldados". No borra el archivo si ya existe. */
function setupCleared_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.CLEARED);
  // Si ya está inicializada con encabezados, conservar el archivo existente.
  if (sh.getLastRow() >= 1 && String(sh.getRange(1, 1).getValue()).trim() === CLEARED_HEADERS[0]) return sh;
  sh.clear();
  sh.getRange(1, 1, 1, CLEARED_HEADERS.length).setValues([CLEARED_HEADERS])
    .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  [110, 170, 120, 200, 120, 110, 90, 110, 120, 120, 120, 130, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('Préstamos saldados (pagados en su totalidad), archivados desde "Prestatarios". ' +
    'Se completa con el botón "Mover saldados" o el menú. El historial en "Pagos" se conserva.');
  return sh;
}

/** Devuelve los números de fila de "Prestatarios" cuyos préstamos están PAGADOS. */
function clearedLoanRows_(bs) {
  const out = [];
  if (!bs || bs.getLastRow() < 2) return out;
  const data = bs.getRange(2, 1, bs.getLastRow() - 1, PB.STATE).getValues();
  for (let i = 0; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;         // sin ID
    if (String(data[i][PB.STATE - 1]).trim() === ST.PAID) out.push(i + 2); // Estado
  }
  return out;
}

/** Mueve todos los préstamos PAGADOS a "Saldados" y libera sus filas. Devuelve la cantidad movida. */
function doMoveCleared_() {
  return guard_('doMoveCleared_', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const cs = ss.getSheetByName(CFG.SHEETS.CLEARED) || setupCleared_(ss);
    if (!bs) throw new Error('Falta la hoja "Prestatarios". Ejecute ① Configurar.');
    const rowsToMove = clearedLoanRows_(bs);
    if (!rowsToMove.length) return 0;
    const now = new Date();
    const archive = [];
    rowsToMove.forEach(function (row) {
      // Modelo normalizado. v = A..L (0..11): ID Préstamo, ID Cliente, Capital, Plazo,
      // Tasa, FechaPréstamo, FechaVenc, Interés, Total, Pagado, Saldo, Estado.
      const v = bs.getRange(row, 1, 1, PB.STATE).getValues()[0];
      const c = getClienteById_(String(v[PB.CLIENT_ID - 1] || '').trim()) || {};
      // Archivo: ID, Prestatario, DNI, Correo, Teléfono, Capital, Plazo, Interés, Total, Pagado, FechaPréstamo, FechaVenc, FechaSaldado.
      archive.push([v[PB.LOAN_ID - 1], c.name || '', c.dni || '', c.email || '', c.phone || '',
        v[PB.PRINCIPAL - 1], v[PB.TERM - 1], v[PB.INTEREST - 1], v[PB.TOTAL - 1], v[PB.PAID - 1],
        v[PB.LOAN_DATE - 1], v[PB.DUE - 1], now]);
      // Limpiar solo las celdas de entrada (conserva la plantilla de fórmulas, incl. C:D Nombre/DNI).
      bs.getRange(row, PB.LOAN_ID, 1, 2).clearContent();    // A:B ID, Cliente
      bs.getRange(row, PB.PRINCIPAL, 1, 2).clearContent();  // E:F Capital, Plazo
      bs.getRange(row, PB.LOAN_DATE).clearContent();        // Fecha del préstamo
      bs.getRange(row, PB.PDF, 1, 6).clearContent();        // bloque contrato/avisos/firma
    });
    const start = cs.getLastRow() + 1;
    cs.getRange(start, 1, archive.length, CLEARED_HEADERS.length).setValues(archive);
    cs.getRange(start, 6, archive.length, 1).setNumberFormat(CFG.CURRENCY_FMT);   // Capital
    cs.getRange(start, 8, archive.length, 3).setNumberFormat(CFG.CURRENCY_FMT);   // Interés, Total, Pagado
    cs.getRange(start, 11, archive.length, 3).setNumberFormat('yyyy-mm-dd');      // 3 fechas
    SpreadsheetApp.flush();
    try { refreshAll(true); } catch (e) { logError_('doMoveCleared_/refreshAll', e); }
    return archive.length;
  });
}

/** Menú: confirma y mueve los préstamos saldados. */
function moveClearedBorrowersConfirm() {
  const ui = SpreadsheetApp.getUi();
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  const n = clearedLoanRows_(bs).length;
  if (n === 0) { ui.alert('Saldados', 'No hay préstamos saldados (pagados en su totalidad) para mover.', ui.ButtonSet.OK); return; }
  const r = ui.alert('Mover saldados',
    'Se moverán ' + n + ' préstamo(s) saldado(s) de "Prestatarios" a la hoja "Saldados". ' +
    'El historial de pagos se conserva. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) moveClearedBorrowers();
}

function moveClearedBorrowers() {
  const n = doMoveCleared_();
  SpreadsheetApp.getUi().alert('Saldados',
    n ? ('Se movieron ' + n + ' préstamo(s) saldado(s) a la hoja "Saldados".')
      : 'No hay préstamos saldados para mover.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ==================================================================
 *  VISIBILIDAD DE PESTAÑAS (operaciones diarias vs. secundarias)
 * ================================================================== */
function isMainSheet_(name) { return MAIN_SHEETS.indexOf(name) !== -1; }

/** Deja visibles las hojas de operaciones diarias y oculta las secundarias. Devuelve cuántas ocultó. */
function applyDailyVisibility_(ss) {
  ss = ss || getSS_();
  // Activar una hoja principal para no dejar activa una que se va a ocultar.
  const home = ss.getSheetByName(CFG.SHEETS.BORROWERS) || ss.getSheets()[0];
  if (home) { try { home.showSheet(); ss.setActiveSheet(home); } catch (e) {} }
  let hidden = 0;
  ss.getSheets().forEach(function (sh) {
    if (isMainSheet_(sh.getName())) { try { sh.showSheet(); } catch (e) {} }
    else { try { sh.hideSheet(); hidden++; } catch (e) { /* no se puede ocultar la última visible */ } }
  });
  return hidden;
}

/** Menú: ver solo las pestañas de operaciones diarias (oculta el resto). */
function showDailyOnly() {
  return guard_('showDailyOnly', function () {
    const ss = getSS_();
    const n = applyDailyVisibility_(ss);
    ss.toast('Vista de operaciones diarias: ' + n + ' hoja(s) ocultada(s). ' +
      'Menú ▸ 🗂 Pestañas ▸ "Mostrar todas las hojas" para verlas.', 'Pestañas', 6);
  });
}

/** Menú: mostrar todas las hojas (incluidas las secundarias). */
function showAllSheets() {
  return guard_('showAllSheets', function () {
    const ss = getSS_();
    let shown = 0;
    ss.getSheets().forEach(function (sh) {
      try { if (sh.isSheetHidden()) shown++; sh.showSheet(); } catch (e) {}
    });
    ss.toast(shown + ' hoja(s) mostrada(s). Todas las pestañas están visibles.', 'Pestañas', 6);
  });
}
