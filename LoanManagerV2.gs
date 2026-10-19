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
    PANEL: 'Panel', STATS: 'Estadísticas', LATE: 'Pagos Atrasados', PLAN: 'En Plan', CLEARED: 'Saldados', REJECTED: 'Rechazados', ERRORS: 'Errores', HELP: 'Instrucciones',
    SIGN: 'Firmas', CLIENTS: 'Clientes', INSTALLMENTS: 'Cuotas', LEVELS: 'Niveles de Clientes',
    DEVICES: 'Dispositivos', FREEZE_REQ: 'Solicitudes de Congelamiento', PLAN_REQ: 'Solicitudes de Plan de Pago',
  },
};
const ST = { ACTIVE: 'ACTIVO', OVERDUE: 'VENCIDO', PAID: 'PAGADO', CLEARED: 'SALDADO', PLAN: 'EN PLAN', FROZEN: 'CONGELADO' };
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
  // Total del PLAN a pagar (incluye mora congelada en préstamos reestructurados). Suma de
  // "Monto Cuota" de "Cuotas"; ESTABLE ante los pagos (a diferencia de "Saldo con Mora").
  TOTAL_PLAN: 27,
  // Congelamiento del préstamo: fecha en que se detiene el reloj de la mora ("Congelado el";
  // vacío = no congelado) y acumulador de días ya saltados por ciclos previos de descongelado
  // ("Días congelados", gestionado por el script). Ver effectiveMoraAsOf_ / frozenInfo_.
  FROZEN_SINCE: 28, FROZEN_DAYS: 29,
  // Plan de pago (Acuerdo Transaccional): fecha en que se firmó el acuerdo ("Plan desde"; vacío =
  // sin plan). Marca el estado EN PLAN y suspende la mora (se fija "Mora (ajuste)"=0 al activar).
  PLAN_SINCE: 30,
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
const PP = { PAY_ID: 1, LOAN_ID: 2, DATE: 3, AMOUNT: 4, BALANCE: 5, RECEIPT: 6, NAME: 7, DNI: 8, SEND_RECEIPT: 9, GEN_RECEIPT: 10, WA_PAYMENT: 11 };
// Hojas de operaciones diarias: siempre visibles. El resto son de configuración/consulta
// y se pueden ocultar/mostrar desde el menú ▸ 🗂 Pestañas.
const MAIN_SHEETS = [
  CFG.SHEETS.BORROWERS,   // Prestatarios — cartera
  CFG.SHEETS.NEW,         // Nuevos Prestatarios — solicitudes a revisar
  CFG.SHEETS.PAYMENTS,    // Pagos — registrar cobros
  CFG.SHEETS.INSTALLMENTS,// Cuotas — cronograma de pagos
  CFG.SHEETS.LATE,        // Pagos Atrasados — mora / avisos
  CFG.SHEETS.FREEZE_REQ,  // Solicitudes de Congelamiento — aprobar/rechazar
  CFG.SHEETS.PLAN_REQ,    // Solicitudes de Plan de Pago — aprobar/rechazar
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
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestor de Préstamos')
    // — Operaciones diarias —
    .addItem('▶ Abrir panel del prestamista', 'openSidebar')
    .addItem('🆕 Aplicar novedades (columnas y ajustes)', 'applyFeatureUpdates')
    .addItem('⑤ Actualizar saldos y resumen', 'refreshAll')
    .addItem('⑥ Formulario web — publicar / ver enlace', 'showWebFormLink')
    .addSeparator()
    // — Informes —
    .addItem('④ Estadísticas y gráficos', 'showStats')
    .addItem('🎓 Niveles de clientes (escalera de graduación)', 'buildClientLevels')
    .addItem('💵 Retiro disponible (sin frenar crecimiento)', 'showWithdrawable')
    .addSeparator()
    // — Gestión de préstamos (fila activa) —
    .addItem('🔄 Reestructurar préstamo vencido (congelar mora + plan de cuotas)', 'restructurarPrestamoUI')
    .addItem('✔ Mover préstamos saldados → hoja Saldados', 'moveClearedBorrowersConfirm')
    .addItem('✉ Reenviar enlace de firma (fila seleccionada)', 'resendSigningLink_')
    .addSubMenu(ui.createMenu('❄ Congelar préstamo (fila activa)')
      .addItem('❄ Congelar préstamo (la mora deja de crecer)…', 'congelarPrestamoFilaActiva')
      .addItem('☀ Descongelar préstamo (la mora vuelve a correr)', 'descongelarPrestamoFilaActiva'))
    .addSubMenu(ui.createMenu('💸 Mora (fila activa)')
      .addItem('Condonar mora (poner en 0)', 'condonarMoraFilaActiva')
      .addItem('Fijar un importe de mora…', 'fijarMoraFilaActiva')
      .addItem('Restaurar mora automática', 'restaurarMoraFilaActiva'))
    .addSubMenu(ui.createMenu('🚫 Bloqueo de clientes (fila activa en Clientes)')
      .addItem('Bloquear cliente…', 'bloquearClienteFilaActiva')
      .addItem('Desbloquear cliente', 'desbloquearClienteFilaActiva'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🗂 Pestañas')
      .addItem('Ver solo operaciones diarias', 'showDailyOnly')
      .addItem('Mostrar todas las hojas', 'showAllSheets'))
    // — Configuración y mantenimiento (uso poco frecuente) —
    .addSubMenu(ui.createMenu('⚙ Configuración y mantenimiento')
      .addItem('① Configurar / reconstruir hojas', 'setupConfirm')
      .addItem('Configurar marca (nombre + logo)', 'setBranding')
      .addSeparator()
      .addItem('⑤· Actualizar SOLO Pagos Atrasados (mora)', 'refreshLateOnly')
      .addItem('💳 Actualizar Saldo Posterior en «Pagos» (con mora)', 'actualizarSaldoPosteriorPagos')
      .addItem('📅 Actualizar «Vencimientos por período» (semana/mes/mes próximo)', 'actualizarVencimientosPanel')
      .addItem('🧮 Reparar totales y normalizar pagos', 'repairTotals')
      .addItem('🧮 Reparar cuotas (duplicadas + fechas)', 'repararCuotas_')
      .addItem('🔁 Regenerar cuotas (1.ª cuota vence la semana próxima)', 'regenerarCuotasPrestamoUI')
      .addSeparator()
      .addItem('➕ Crear hoja de Recordatorios de pago', 'createRemindersSheet_')
      .addItem('📄 Crear hoja de Estudio de Contratos', 'createAgreementSheet_')
      .addItem('📄 Crear/actualizar hoja de Estudio de Estados', 'createStatementsSheet_')
      .addItem('📊 Crear/actualizar Panel (incluye «Préstamos en plan»)', 'createPanelSheet_')
      .addItem('📋 Crear/actualizar hoja «En Plan» (préstamos reestructurados)', 'createEnPlanSheet_')
      .addSeparator()
      .addSubMenu(ui.createMenu('🧪 Datos de prueba')
        .addItem('Crear 40 prestatarios de prueba', 'seedTestBorrowersConfirm')
        .addItem('Eliminar prestatarios de prueba', 'removeTestBorrowersConfirm')
        .addSeparator()
        .addItem('💬 Crear escenarios de pago (Avisar pago 📲)', 'seedPaymentScenariosConfirm')
        .addItem('💬 Eliminar escenarios de pago', 'removePaymentScenariosConfirm')))
    .addToUi();

  // Regenerar "Niveles de Clientes" en cada apertura para que no quede estático.
  // Silencioso: no cambia la hoja activa ni muestra toast. Un fallo aquí no debe
  // impedir que se cargue el menú.
  try { buildClientLevels({ silent: true }); } catch (e) {}

  // Reconstruir "Pagos Atrasados" en cada apertura. Un préstamo entra en mora por el
  // PASO DEL TIEMPO (nadie edita nada cuando vence), así que onEdit no alcanza: solo una
  // reconstrucción al abrir o el disparador diario la mantienen al día. Silencioso y
  // tolerante a fallos para no impedir la carga del menú.
  try { rebuildLateSheet_(getSS_()); } catch (e) {}

  // Reconstruir "Estudio de Recordatorios" en cada apertura: igual que la mora, un préstamo entra
  // en la lista por el PASO DEL TIEMPO (vence o se acerca el vencimiento), así que conviene
  // refrescarla al abrir. Silencioso y tolerante a fallos; la lista se genera leyendo en bloque.
  try { if (getSS_().getSheetByName(CFG.SHEETS.REMINDER)) rebuildRemindersSheet_(getSS_()); } catch (e) {}
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
  for (let i = 0; i < data.length; i++) if (String(data[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(value); invalidateSettingsCache_(); return; }
  sh.appendRow([key, value]); invalidateSettingsCache_();
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

/** Días de gracia antes de que empiece a correr la mora (Configuración). Predet. 0 (sin gracia). */
function moraGraceDays_() {
  const raw = String(getSetting_('Días de gracia antes de mora') || '').replace(/[^0-9.\-]/g, '').trim();
  const n = raw === '' ? 0 : Number(raw);
  return (isFinite(n) && n >= 0) ? Math.floor(n) : 0;
}
/** Tope acumulado de mora como FRACCIÓN del TOTAL A DEVOLVER (capital + interés).
 *  Ajuste "Tope de mora (% del total a devolver)"; acepta la clave anterior
 *  "(% del capital)" hasta que corra la migración. Predeterminado 100% si falta. */
function moraCapFrac_() {
  const raw = String(getSetting_('Tope de mora (% del total a devolver)') ||
    getSetting_('Tope de mora (% del capital)') || '').replace(/[^0-9.\-]/g, '').trim();
  const n = raw === '' ? 100 : Number(raw);
  return (isFinite(n) && n > 0) ? n / 100 : 1;
}
/** Texto del tope de mora como % (p. ej. "100%"). */
function moraCapPctText_() { return round2_(moraCapFrac_() * 100) + '%'; }
/** Tope de mora en PESOS = % configurado × CAPITAL del préstamo (no el total). El recargo por mora
 *  nunca supera este monto; con el tope al 100% equivale a "la mora no supera el capital prestado".
 *  Es ≤ el tope por contrato (% del total), por lo que cobrar hasta acá queda dentro de lo firmado. */
function moraCapAmount_(loan) { return round2_((Number(loan && loan.principal) || 0) * moraCapFrac_()); }

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
  setupErrors_(ss); setupHelp_(ss); setupFirmas_(ss); setupDispositivos_(ss); setupFreezeRequests_(ss); setupPlanRequests_(ss);
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
      'Límite inicial (préstamo nuevo)', 'Incremento de graduación', 'Límite máximo por historial (graduación)',
      'Atraso que congela el nivel (días)', 'Atraso que reinicia el nivel (días)',
      'Días de gracia antes de mora', 'Tope de mora (% del total a devolver)',
      // Congelamiento (Parte 2): tope de futuros préstamos para clientes que congelaron,
      // y porcentaje que debe estar PAGADO antes de aprobar/congelar.
      'Tope de préstamo tras congelamiento', 'Porcentaje a pagar para congelar (%)',
      // Plan de pago (Acuerdo Transaccional): % inicial a pagar para activar el plan.
      'Porcentaje inicial para plan de pago (%)',
      // Plan flexible (por dificultad): % inicial BAJO para que un vencido pueda arrancar un plan (0–10%).
      'Pago inicial para plan flexible (%)',
      // Tramo grande ($300k–$500k): tamaño objetivo de cada cuota para que sea manejable (↓ default).
      'Cuota manejable objetivo (tramo grande)',
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
      // Escalera de graduación (V-32): inicial 150k, sube de a 50k hasta 500k (tope).
      if (stsh && String(getSetting_('Límite inicial (préstamo nuevo)') || '').trim() === '')
        setSettingValue_(stsh, 'Límite inicial (préstamo nuevo)', '150000');
      if (stsh && String(getSetting_('Incremento de graduación') || '').trim() === '')
        setSettingValue_(stsh, 'Incremento de graduación', '50000');
      // Migra el paso grueso de $100.000 (dejaba el tope en $450k, inalcanzable el de $500k) al
      // paso gradual de $50.000: la escalera sube 150→200→…→500k y el tope SÍ se alcanza. Respeta
      // cualquier otro valor manual (solo corrige el legacy exacto de 100000).
      if (stsh && String(getSetting_('Incremento de graduación') || '').replace(/[^0-9]/g, '').trim() === '100000')
        setSettingValue_(stsh, 'Incremento de graduación', '50000');
      if (stsh && String(getSetting_('Límite máximo por historial (graduación)') || '').trim() === '')
        setSettingValue_(stsh, 'Límite máximo por historial (graduación)', '500000');
      // Tramo grande: tamaño objetivo de cuota manejable (franja de 0% de default). Default $150.000.
      if (stsh && String(getSetting_('Cuota manejable objetivo (tramo grande)') || '').trim() === '')
        setSettingValue_(stsh, 'Cuota manejable objetivo (tramo grande)', '150000');
      // Escalera graduada ante atrasos (V-32): ≤3 días congela, 4–15 días baja un nivel, >15 reinicia.
      if (stsh && String(getSetting_('Atraso que congela el nivel (días)') || '').trim() === '')
        setSettingValue_(stsh, 'Atraso que congela el nivel (días)', '3');
      if (stsh && String(getSetting_('Atraso que reinicia el nivel (días)') || '').trim() === '')
        setSettingValue_(stsh, 'Atraso que reinicia el nivel (días)', '15');
      // Congelamiento (Parte 2): tope de futuros préstamos para clientes que congelaron un préstamo.
      if (stsh && String(getSetting_('Tope de préstamo tras congelamiento') || '').trim() === '')
        setSettingValue_(stsh, 'Tope de préstamo tras congelamiento', '20000');
      // Porcentaje que debe estar pagado antes de aprobar/congelar (default 50%).
      if (stsh && String(getSetting_('Porcentaje a pagar para congelar (%)') || '').trim() === '')
        setSettingValue_(stsh, 'Porcentaje a pagar para congelar (%)', '50');
      // Plan de pago: % inicial para activar (default 20%).
      if (stsh && String(getSetting_('Porcentaje inicial para plan de pago (%)') || '').trim() === '')
        setSettingValue_(stsh, 'Porcentaje inicial para plan de pago (%)', '20');
      // Plan flexible: pago inicial bajo (default 10%). Permite arrancar un plan en cuotas a un vencido.
      if (stsh && String(getSetting_('Pago inicial para plan flexible (%)') || '').trim() === '')
        setSettingValue_(stsh, 'Pago inicial para plan flexible (%)', '10');
      // Limpieza: quita la fila de la redirección de WhatsApp de prueba (quedó obsoleta; ahora todos
      // los avisos van SIEMPRE al teléfono real del prestatario resuelto de «Clientes»).
      if (stsh) {
        const OLD_WA = 'WhatsApp de prueba (redirigir avisos)';
        const wk = stsh.getRange(1, 1, stsh.getLastRow(), 1).getValues();
        for (let i = wk.length - 1; i >= 0; i--)
          if (hkey_(wk[i][0]) === hkey_(OLD_WA)) stsh.deleteRow(i + 1);
      }
      // Mora: sin días de gracia y tope al 100% del TOTAL A DEVOLVER. Migra los valores
      // sembrados por versiones anteriores (3 / 50 y la clave "% del capital"); respeta
      // cualquier otro valor manual.
      if (stsh) {
        const grace = String(getSetting_('Días de gracia antes de mora') || '').trim();
        if (grace === '' || grace === '3') setSettingValue_(stsh, 'Días de gracia antes de mora', '0');
        const NEW_TOPE = 'Tope de mora (% del total a devolver)', OLD_TOPE = 'Tope de mora (% del capital)';
        let tope = String(getSetting_(NEW_TOPE) || '').trim();
        const topeViejo = String(getSetting_(OLD_TOPE) || '').trim();
        if (tope === '' && topeViejo !== '') { setSettingValue_(stsh, NEW_TOPE, topeViejo); tope = topeViejo; }
        // Elimina la fila de la clave anterior para no dejar un ajuste duplicado sin efecto.
        const keys = stsh.getRange(1, 1, stsh.getLastRow(), 1).getValues();
        for (let i = keys.length - 1; i >= 0; i--)
          if (hkey_(keys[i][0]) === hkey_(OLD_TOPE)) stsh.deleteRow(i + 1);
        if (tope === '' || tope === '50') setSettingValue_(stsh, NEW_TOPE, '100');
        // Actualiza la Cláusula de Mora solo si conserva un texto sembrado anterior
        // (mencionaba la gracia o el tope como % del capital).
        const clause = String(getSetting_('Cláusula de Mora') || '').trim();
        if (clause === '' || /tras un período de gracia|porcentaje del capital/.test(clause))
          setSettingValue_(stsh, 'Cláusula de Mora', 'En caso de mora se aplica un recargo diario sobre el total a devolver por cada día de atraso posterior al vencimiento, con un tope acumulado como porcentaje del capital. Los valores vigentes (recargo diario y tope) se configuran en esta hoja.');
      }
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
        // Columna "Avisar pago 📲": enlaces de WhatsApp de aviso de pago (confirmación + saldo + vencimiento).
        SpreadsheetApp.flush(); // que el "Saldo Posterior" ya esté calculado antes de leerlo
        const nWa = fillPaymentWaLinks_(ss);
        done.push('casilla «Generar recibo» + «Saldo Posterior» vivo en «Pagos»');
        done.push('columna «Avisar pago 📲» (WhatsApp) en «Pagos»' + (nWa ? (' · ' + nWa + ' fila(s)') : ''));
      }
    } catch (e) { logError_('applyFeatureUpdates:pagos', e); }
    // 7b) Congelamiento de préstamos: columnas "Congelado el" (fecha) + "Días congelados" (auto)
    //     en "Prestatarios" — permiten frenar la mora de un préstamo (menú ❄ Congelar / ☀ Descongelar).
    //     Aditivo (apéndalas al final por nombre); todo se resuelve por encabezado.
    try {
      const bsf = ss.getSheetByName(CFG.SHEETS.BORROWERS);
      if (bsf) {
        const sinceCol = clientesHeaderCol_(bsf, 'Congelado el', true);
        const daysCol = clientesHeaderCol_(bsf, 'Días congelados', true);
        if (sinceCol > 0) {
          bsf.getRange(2, sinceCol, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd').setDataValidation(datePicker_());
          bsf.setColumnWidth(sinceCol, 110);
          bsf.getRange(1, sinceCol).setNote('CONGELAMIENTO. Fecha desde la que la mora deja de crecer. Vacío = NO congelado. Usá el menú ▸ ❄ Congelar / ☀ Descongelar. Con fecha, el Estado pasa a CONGELADO.');
        }
        if (daysCol > 0) {
          bsf.getRange(2, daysCol, CFG.MAX_ROWS, 1).setNumberFormat('0');
          bsf.setColumnWidth(daysCol, 110);
          bsf.getRange(1, daysCol).setNote('Días de atraso ya saltados por descongelados previos (lo gestiona el menú ☀ Descongelar). No editar a mano.');
        }
        invalidateHeaderIndexCache_(); // para que rebuildLateSheet_ vea las columnas nuevas
        done.push('columnas «Congelado el» + «Días congelados» en «Prestatarios»');
      }
    } catch (e) { logError_('applyFeatureUpdates:congelamiento', e); }
    // 7c) Solicitudes de congelamiento (Parte 2): crea la hoja, la columna "Congelamientos" en
    //     "Clientes" (historial para el tope de futuros préstamos) — todo aditivo/idempotente.
    try {
      setupFreezeRequests_(ss);
      const csf = ss.getSheetByName(CFG.SHEETS.CLIENTS);
      if (csf) { clientesHeaderCol_(csf, 'Congelamientos', true); invalidateClientesCache_(); }
      done.push('hoja «Solicitudes de Congelamiento» + columna «Congelamientos» en «Clientes»');
    } catch (e) { logError_('applyFeatureUpdates:congelamientoSolicitudes', e); }
    // 7d) Plan de pago (Acuerdo Transaccional): columna "Plan desde" en "Prestatarios" + la hoja
    //     "Solicitudes de Plan de Pago". Aditivo/idempotente.
    try {
      const bsp = ss.getSheetByName(CFG.SHEETS.BORROWERS);
      if (bsp) {
        const planCol = clientesHeaderCol_(bsp, 'Plan desde', true);
        if (planCol > 0) {
          bsp.getRange(2, planCol, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');
          bsp.setColumnWidth(planCol, 110);
          bsp.getRange(1, planCol).setNote('PLAN DE PAGO. Fecha del Acuerdo Transaccional firmado. Vacío = sin plan. Con fecha, el Estado pasa a EN PLAN y la mora queda suspendida. Lo gestiona el flujo de firma.');
        }
        invalidateHeaderIndexCache_();
      }
      setupPlanRequests_(ss);
      done.push('columna «Plan desde» en «Prestatarios» + hoja «Solicitudes de Plan de Pago»');
    } catch (e) { logError_('applyFeatureUpdates:planPago', e); }
    // 8) "Pagos Atrasados": columnas "Tope de mora" y "Plan de pago 📅 (3 cuotas)" — la
    //    reconstrucción reafirma los encabezados nuevos y recalcula el % del tope.
    try { rebuildLateSheet_(ss); done.push('columnas «Tope de mora» + «Plan de pago 📅» en «Pagos Atrasados»'); }
    catch (e) { logError_('applyFeatureUpdates:pagosAtrasados', e); }
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
    ['Cláusula de Mora', 'En caso de mora se aplica un recargo diario sobre el total a devolver por cada día de atraso posterior al vencimiento, con un tope acumulado como porcentaje del capital. Los valores vigentes (recargo diario y tope) se configuran en esta hoja.'],
    ['Recargo por Mora diario (%)', '5'],
    ['Días de gracia antes de mora', '0'],
    ['Tope de mora (% del total a devolver)', '100'],
    ['Límite inicial (préstamo nuevo)', '150000'],
    ['Incremento de graduación', '50000'],
    ['Límite máximo por historial (graduación)', '500000'],
    // Escalera graduada ante atrasos (V-32): ≤ N días congela el nivel; de N+1 hasta M días baja
    // un nivel (recuperable); > M días o un préstamo VENCIDO sin pagar reinicia a Inicial.
    ['Atraso que congela el nivel (días)', '3'],
    ['Atraso que reinicia el nivel (días)', '15'],
    ['Pie del Contrato', 'Este acuerdo es legalmente vinculante desde la firma de ambas partes.'],
    ['Días de aviso antes del vencimiento', '3'],
    ['Días de aviso de recordatorio', '7'],
    ['Fondo total para prestar', '0'],
    ['Tope de concentración por prestatario (% del capital disponible)', '10'],
    ['Horizonte de protección del crecimiento (días)', '60'],
    ['Ventana para estimar el ritmo de colocación (días)', '60'],
    ['Aceptar solicitudes de préstamo (SÍ/NO)', 'SÍ'],
    ['Capturar ubicación en la firma (SÍ/NO)', 'SÍ'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff');
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 500);
  sh.getRange(1, 1, rows.length, 2).setWrap(true); sh.setFrozenRows(1);
  invalidateSettingsCache_(); // se reescribió toda la hoja de ajustes
}

function setupBorrowers_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.BORROWERS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const H = ['ID Préstamo', 'ID Cliente', 'Nombre', 'DNI', 'Capital', 'Plazo (meses)', 'Tasa',
    'Fecha del Préstamo', 'Fecha de Vencimiento', 'Interés', 'Total a Pagar',
    'Total Pagado', 'Saldo Pendiente (hoy)', 'Estado', 'Contrato PDF', 'Contrato Enviado', 'Último Aviso',
    'Estado de Firma', 'Fecha de Firma', 'Contrato Firmado (PDF)', 'Eliminar', 'Enviar recibo desembolso', 'Préstamos (de ' + MAX_LOANS_PER_CLIENT + ')',
    'Recargo por Mora (acum.)', 'Mora (ajuste)', 'Saldo con Mora', 'Suma de cuotas', 'Congelado el', 'Días congelados', 'Plan desde'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
  sh.getRange(1, PB.DELETE).setBackground('#990000').setNote('Tildá para ELIMINAR ese préstamo (borra la fila; recuperable con el historial de versiones).');
  sh.getRange(1, PB.SEND_DISB).setNote('Tildá para enviar el recibo de DESEMBOLSO (entrega de fondos) al prestatario. Sólo se envía si el contrato está FIRMADO. Se destilda solo.');
  sh.getRange(1, PB.LOANCOUNT).setNote('Cantidad de préstamos vigentes del prestatario (máximo ' + MAX_LOANS_PER_CLIENT + '). Se pinta de rojo al alcanzar el tope.');
  sh.getRange(1, PB.BALANCE).setNote('Saldo del préstamo: Total a Pagar − Total Pagado (capital + interés). El recargo por mora se muestra aparte en "Recargo por Mora (acum.)".');
  sh.getRange(1, PB.MORA_ACUM).setNote('Recargo por mora acumulado. Automático (derivado de "Pagos Atrasados") salvo que se cargue un valor en "Mora (ajuste)". 0 si el préstamo no está en mora.');
  sh.getRange(1, PB.MORA_ADJ).setNote('AJUSTE MANUAL de la mora de ESTE préstamo. Dejala VACÍA para el cálculo automático. Escribí 0 para CONDONAR (eximir) la mora, o un importe para FIJAR un recargo distinto. Afecta el "Recargo por Mora", el "Saldo con Mora", los avisos y el estado de cuenta.');
  sh.getRange(1, PB.BALANCE_MORA).setNote('Monto a pagar HOY = Saldo Pendiente + Recargo por Mora (ya con el ajuste manual, si lo hay).');
  sh.getRange(1, PB.TOTAL_PLAN).setNote('Total del PLAN a pagar, incluyendo la mora congelada (para préstamos reestructurados = total del plan de cuotas). Es la suma de "Monto Cuota" de "Cuotas"; NO cambia con los pagos. "Saldo con Mora" muestra lo que aún falta pagar.');
  sh.getRange(1, PB.FROZEN_SINCE).setNote('CONGELAMIENTO. Fecha desde la que la mora deja de crecer. Vacío = préstamo NO congelado. Usá el menú ▸ ❄ Congelar / ☀ Descongelar (no hace falta editar a mano). Con fecha, el Estado pasa a CONGELADO.');
  sh.getRange(1, PB.FROZEN_DAYS).setNote('Días de atraso ya SALTADOS por descongelados previos (gestionado por el menú ☀ Descongelar). Se descuentan del atraso para que el período congelado nunca cobre mora. No editar a mano.');
  sh.getRange(1, PB.PLAN_SINCE).setNote('PLAN DE PAGO (Acuerdo Transaccional). Fecha en que se firmó el acuerdo. Vacío = sin plan. Con fecha, el Estado pasa a EN PLAN y la mora queda suspendida (Mora (ajuste)=0). Lo gestiona el flujo de firma; no editar a mano.');
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
    const ID = colL_(PB.LOAN_ID), CAP = colL_(PB.PRINCIPAL), PA = colL_(PB.PAID), col = PB.PLAN_SINCE + 1; // 1ª columna libre tras las columnas de datos
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
  const fName = [], fRate = [], fDue = [], fITP = [], fState = [], fCount = [], fMora = [], fBalMora = [], fPlan = [];
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
    fPlan.push([borrowerPlanTotalFormula_(r)]);          // Suma de cuotas = total del plan (Σ Monto Cuota)
  }
  sh.getRange(2, PB.NAME, N, 2).setFormulas(fName); // Nombre + DNI
  sh.getRange(2, PB.RATE, N, 1).setFormulas(fRate);
  sh.getRange(2, PB.DUE, N, 1).setFormulas(fDue);
  sh.getRange(2, PB.INTEREST, N, 4).setFormulas(fITP); // Interés, Total, Pagado, Saldo
  sh.getRange(2, PB.STATE, N, 1).setFormulas(fState);
  sh.getRange(2, PB.LOANCOUNT, N, 1).setFormulas(fCount); // Préstamos (de N)
  sh.getRange(2, PB.MORA_ACUM, N, 1).setFormulas(fMora);          // Recargo por Mora (acum.)
  sh.getRange(2, PB.BALANCE_MORA, N, 1).setFormulas(fBalMora);    // Saldo con Mora ("Mora (ajuste)" queda de entrada)
  sh.getRange(2, PB.TOTAL_PLAN, N, 1).setFormulas(fPlan);         // Suma de cuotas = total del plan
  sh.getRange(2, PB.MORA_ACUM, N, 4).setNumberFormat(CFG.CURRENCY_FMT); // formato moneda: mora (3) + total del plan (1)
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
 * Columna "Total Pagado" (col J) como fórmula VIVA: SUMIF de todos los pagos de este
 * préstamo en la hoja "Pagos". Reafirmarla en cada recálculo sana las filas que quedaron
 * con un valor ESTÁTICO (migración / pegado) y por eso no acreditaban pagos posteriores.
 */
function borrowerPaidFormula_(r) {
  const A = colL_(PB.LOAN_ID), P = CFG.SHEETS.PAYMENTS, PL = colL_(PP.LOAN_ID), PA = colL_(PP.AMOUNT);
  return `=IF($${A}${r}="","",SUMIF('${P}'!$${PL}:$${PL},$${A}${r},'${P}'!$${PA}:$${PA}))`;
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
 * Columna "Suma de cuotas" = total del plan a pagar. Suma los "Monto Cuota" del
 * préstamo en "Cuotas" (se fijan al armar el plan; los pagos sólo tocan "Monto Pagado"/"Saldo
 * Cuota", nunca "Monto Cuota"), así que es ESTABLE ante los pagos. Para un préstamo
 * reestructurado equivale a capital + interés + mora congelada. Respaldo en "Total a Pagar"
 * si el préstamo aún no tiene cronograma (SUMIF = 0). `aCol`/`totCol` = letras A1 de
 * "ID Préstamo" y "Total a Pagar" (por defecto, las del mapa PB).
 */
function borrowerPlanTotalFormula_(r, aCol, totCol) {
  const A = aCol || colL_(PB.LOAN_ID), TOT = totCol || colL_(PB.TOTAL);
  const C = CFG.SHEETS.INSTALLMENTS, L = colL_(CU.LOAN_ID), E = colL_(CU.AMOUNT);
  const sumf = `SUMIF('${C}'!$${L}:$${L},$${A}${r},'${C}'!$${E}:$${E})`;
  return `=IF($${A}${r}="","",IF(${sumf}>0,${sumf},N($${TOT}${r})))`;
}
/**
 * Fórmula de la columna "Enlace de Firma": muestra la URL COMPLETA de la página de firma
 * del contrato de ese préstamo (clicable y copiable para compartir por WhatsApp/correo).
 * Muestra "✔ Firmado" si ya se firmó; si no hay URL de app web publicada, deja un aviso en
 * vez de un enlace roto. Nunca es estático: se recalcula con el ID del préstamo. `A`/`firmaL`
 * = letras A1 de "ID Préstamo" y "Estado de Firma"; `base` = URL base de la app web
 * (webAppBaseUrl_()), fijada al escribir la fórmula.
 */
function signingLinkFormula_(r, A, firmaL, base) {
  const a = '$' + A + r;
  if (!base) return `=IF(${a}="","","Publicá la app web (menú ⑥) para el enlace de firma")`;
  // URL = texto visible Y destino del HYPERLINK: así la celda muestra el enlace real y, al
  // copiarla, se copia la URL (no una etiqueta).
  const url = `"${base}?page=firmar&loan="&${a}`;
  const link = `HYPERLINK(${url},${url})`;
  if (!firmaL) return `=IF(${a}="","",${link})`;
  return `=IF(${a}="","",IF($${firmaL}${r}="${SIGN.SIGNED}","✔ Firmado",${link}))`;
}
/**
 * Estado (col L) como fórmula VIVA. PAGADO en cuanto Total Pagado (J) alcanza el
 * Total a Pagar (I); si no, VENCIDO cuando pasó el vencimiento (G) o ACTIVO.
 * No depende del saldo escrito por script, así que el estado cambia solo al
 * ingresar el pago (en la hoja "Pagos" o desde el panel).
 */
function borrowerStateFormula_(r) {
  const A = colL_(PB.LOAN_ID), TO = colL_(PB.TOTAL), PA = colL_(PB.PAID), VE = colL_(PB.DUE), BM = colL_(PB.BALANCE_MORA), PL = colL_(PB.TOTAL_PLAN), FS = colL_(PB.FROZEN_SINCE), PS = colL_(PB.PLAN_SINCE);
  // PAGADO sólo cuando el monto a pagar CON MORA ("Saldo con Mora") llega a 0: un préstamo
  // pagado tarde con recargo pendiente NO figura como PAGADO hasta saldar también la mora.
  // CONGELADO: si "Congelado el" tiene fecha y no está saldado, el préstamo está congelado (la
  // mora dejó de crecer) — prevalece sobre VENCIDO/ACTIVO para que se vea de un vistazo.
  // EN PLAN (flag): si "Plan desde" tiene fecha → plan de pago (Acuerdo) vigente, prevalece sobre VENCIDO.
  // EN PLAN (suma): préstamo reestructurado al día — la "Suma de cuotas" supera el "Total a Pagar".
  return `=IF($${A}${r}="","",IF(AND(ISNUMBER($${TO}${r}),ISNUMBER($${PA}${r}),$${PA}${r}>=$${TO}${r}-0.009,N($${BM}${r})<=0.009),"${ST.PAID}",` +
    `IF($${FS}${r}<>"","${ST.FROZEN}",IF($${PS}${r}<>"","${ST.PLAN}",IF(OR(${cuotaVencidaExpr_('$' + A + r)},AND($${VE}${r}<>"",TODAY()>$${VE}${r})),"${ST.OVERDUE}",IF(OR($${TO}${r}="",$${PA}${r}=""),"…",IF(N($${PL}${r})>N($${TO}${r})+0.009,"${ST.PLAN}","${ST.ACTIVE}")))))))`;
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
  // Congelamiento: "Congelado el" (fecha, con selector) + "Días congelados" (entero, auto).
  sh.getRange(2, PB.FROZEN_SINCE, N, 1).setNumberFormat('yyyy-mm-dd').setDataValidation(datePicker_());
  sh.getRange(2, PB.FROZEN_DAYS, N, 1).setNumberFormat('0');
  sh.getRange(2, PB.PLAN_SINCE, N, 1).setNumberFormat('yyyy-mm-dd'); // Plan desde (gestionado por el flujo de firma)
  sh.setColumnWidth(PB.FROZEN_SINCE, 110); sh.setColumnWidth(PB.FROZEN_DAYS, 110); sh.setColumnWidth(PB.PLAN_SINCE, 110);
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
  const needCols = PB.PLAN_SINCE + 5;
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  const estC = colByAny_(HI, ['Estado']) || PB.STATE;
  const signC = colByAny_(HI, ['Estado de Firma']) || PB.SIGN_STATUS;
  const countC = colByAny_(HI, ['Préstamos (de ' + MAX_LOANS_PER_CLIENT + ')', 'Préstamos']) || PB.LOANCOUNT;
  const estL = colL_(estC);
  const rowRange = sh.getRange(2, 1, N, estC);        // A2:<Estado> — campos relevantes de la fila
  const signRange = sh.getRange(2, signC, N, 1);      // fuera de la franja de Estado (sin conflicto)
  const countRange = sh.getRange(2, countC, N, 1);
  // Banner "Capital disponible" (fila 1, a partir de la 1.ª columna libre tras los datos).
  const bannerRange = sh.getRange(1, PB.PLAN_SINCE + 1, 1, 5);
  // Regla por fila: colorea toda la franja según el valor del Estado (col. absoluta, fila relativa).
  const byState = (state, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + estL + '2="' + state + '"')
    .setBackground(bg).setFontColor(fg).setRanges([rowRange]).build();
  return [
    byState(ST.PAID, '#d9ead3', '#274e13'),
    byState(ST.CLEARED, '#d9ead3', '#274e13'),
    byState(ST.OVERDUE, '#f4cccc', '#990000'),
    byState(ST.FROZEN, '#d0e0e3', '#0c343d'),  // CONGELADO (mora detenida): hielo/cian — distinto del ACTIVO
    byState(ST.PLAN, '#d9d2e9', '#351c75'),   // EN PLAN (reestructurado al día): violeta
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

/* ===================== AUDITORÍA DE DISPOSITIVOS ===================== */
/**
 * Hoja "Dispositivos": acta detallada de la huella de dispositivo, las señales de
 * fraude y la ubicación de cada firma. Una fila por evento (modo append, igual que
 * "Firmas"). Se enlaza al cliente por "ID Cliente" y al préstamo por "ID Préstamo".
 * El detalle completo del dispositivo vive AQUÍ (la hoja "Firmas" no se amplía).
 */
const DISPOSITIVOS_HEADERS = ['Fecha y hora', 'ID Préstamo', 'ID Cliente', 'DNI', 'Nombre', 'Correo',
  'Huella de dispositivo (SHA-256)', 'Equipo / SO', 'Modelo', 'Navegador', 'Arquitectura',
  'Pantalla', 'Densidad (DPR)', 'Color (bits)', 'Táctil (puntos)', 'Núcleos CPU', 'Memoria (GB)',
  'Idiomas', 'Zona horaria', 'Desfase UTC', 'Red', 'IP', 'Ubicación (IP)', 'GPS (lat,long)',
  'Precisión GPS (m)', 'Origen de ubicación', 'Automatización (webdriver)', 'Tiempo de lectura (s)',
  'Trazos de firma', 'Duración del trazo (ms)', 'userAgent'];
function setupDispositivos_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.DEVICES);
  const head = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, DISPOSITIVOS_HEADERS.length).getValues()[0].map(h => String(h).trim()) : [];
  if (head.join('|') !== DISPOSITIVOS_HEADERS.join('|')) {
    sh.getRange(1, 1, 1, DISPOSITIVOS_HEADERS.length).setValues([DISPOSITIVOS_HEADERS])
      .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff').setWrap(true);
    [150, 100, 90, 110, 180, 200, 230, 150, 120, 120, 110, 110, 90, 80, 90, 80, 90, 150, 170, 80, 70, 120, 220, 160, 110, 130, 130, 120, 100, 120, 260]
      .forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  sh.setFrozenRows(1);
  return sh;
}

/** ¿Capturar ubicación en la firma? Interruptor de "Configuración" (default: SÍ). */
function locationFlag_() {
  const v = String(getSetting_('Capturar ubicación en la firma (SÍ/NO)') || 'SÍ').trim().toUpperCase();
  return !(v === 'NO' || v === 'N' || v === 'FALSE' || v === '0');
}

/**
 * Huella única y ESTABLE del dispositivo: SHA-256 sobre las señales que no cambian
 * entre firmas (SO, modelo, navegador, pantalla, color, DPR, zona horaria, idiomas,
 * núcleos, memoria, táctil). Se excluyen las volátiles (IP, red, viewport, GPS) para
 * que el mismo equipo produzca el mismo hash → detecta "mismo dispositivo, N préstamos".
 */
function deviceFingerprint_(dev) {
  dev = dev || {};
  const canon = [dev.os, dev.model, dev.browser, dev.screen, dev.color, dev.dpr,
    dev.tz, dev.langs, dev.cores, dev.memory, dev.touch]
    .map(x => String(x == null ? '' : x).trim()).join('|');
  if (!canon.replace(/\|/g, '')) return ''; // sin señales → sin huella
  return hexDigest_(canon);
}

/**
 * Agrega una fila a "Dispositivos" con toda la huella/ubicación de la firma. Best-effort:
 * quien llama lo envuelve en try/catch para que un fallo NUNCA interrumpa la firma.
 */
function logDeviceInfo_(loan, clientId, meta) {
  const sh = setupDispositivos_(getSS_());
  const dev = (meta && meta.device) || {};
  const loc = dev.loc || {};
  const now = (meta && meta.signedAt) || new Date();
  // Cruce liviano de la ubicación IP contra el domicilio declarado (señal de fraude).
  let ipLoc = String(loc.ipText || '');
  const dom = String(loan.direccion || '').toLowerCase();
  if (ipLoc && dom) {
    const hit = ipLoc.toLowerCase().split(/[,·]/).some(t => t.trim().length > 2 && dom.indexOf(t.trim()) !== -1);
    if (!hit) ipLoc = '⚠ (revisar vs domicilio) ' + ipLoc;
  }
  const na = x => (x == null || x === '') ? 'N/D' : x;
  sh.appendRow([
    now, loan.loanId, clientId || loan.clientId || '', loan.dni || '', loan.name || '', loan.email || '',
    (meta && meta.deviceFp) || '',
    na(dev.os), na(dev.model), na(dev.browser), na(dev.arch),
    na(dev.screen), na(dev.dpr), na(dev.color), na(dev.touch),
    na(dev.cores), na(dev.memory), na(dev.langs),
    na(dev.tz), na(dev.tzOffset), na(dev.net),
    (meta && meta.ip) || 'N/D', ipLoc || 'N/D',
    loc.gps || '', (loc.gpsAcc == null ? '' : loc.gpsAcc), na(loc.source),
    dev.webdriver ? 'SÍ' : 'No',
    (meta && meta.dwellMs != null) ? Math.round(meta.dwellMs / 1000) : '',
    (meta && meta.strokes != null) ? meta.strokes : '',
    (meta && meta.drawMs != null) ? meta.drawMs : '',
    String((meta && meta.userAgent) || dev.userAgent || '').slice(0, 300),
  ]);
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
    [PB.NAME, PB.DNI, PB.RATE, PB.DUE, PB.INTEREST, PB.TOTAL, PB.PAID, PB.BALANCE, PB.STATE, PB.LOANCOUNT, PB.TOTAL_PLAN]
      .map(colL_).forEach(col => {
        const p = sh.getRange(col + '2:' + col + (N + 1)).protect()
          .setDescription('Columna calculada — no editar');
        p.setWarningOnly(true);
      });
  } catch (e) { /* ignora si no hay permisos */ }
}

/**
 * Fórmula del "Saldo Posterior" de una fila de Pagos: lo que falta pagar del préstamo
 * después de este pago = (Total a Pagar + Recargo por Mora acum.) − Σ pagos de ese préstamo
 * hasta esta fila (inclusive). Incluye la MORA para que el saldo refleje el monto real
 * adeudado (igual a "Saldo con Mora" de "Prestatarios" en el último pago): un préstamo con
 * capital+interés saldado pero mora pendiente ya NO muestra $0.
 *
 * Columnas resueltas por NOMBRE de encabezado (robusto al layout migrado, donde el orden
 * físico de "Prestatarios" difiere del mapa PB), con respaldo a PP/PB. Si aún no existe la
 * columna de mora, la fórmula equivale al comportamiento anterior (capital+interés).
 */
function paymentBalanceFormula_(r) {
  const ss = getSS_();
  const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS), bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const PH = ps ? headerIndex_(ps) : {}, BH = bs ? headerIndex_(bs) : {};
  const B = colL_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID);
  const D = colL_(colByAny_(PH, ['Monto Pagado']) || PP.AMOUNT);
  const BS = CFG.SHEETS.BORROWERS;
  const bId = colL_(colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID);
  const bTot = colL_(colByAny_(BH, ['Total a Pagar', 'Total a pagar']) || PB.TOTAL);
  const moraC = colByAny_(BH, ['Recargo por Mora (acum.)']);
  const match = `MATCH($${B}${r},'${BS}'!$${bId}:$${bId},0)`;
  const totRef = `IFERROR(INDEX('${BS}'!$${bTot}:$${bTot},${match}),0)`;
  const moraRef = moraC ? `+IFERROR(INDEX('${BS}'!$${colL_(moraC)}:$${colL_(moraC)},${match}),0)` : '';
  const paid = `SUMIFS($${D}$2:$${D}${r},$${B}$2:$${B}${r},$${B}${r})`;
  return `=IF($${B}${r}="","",MAX(0,${totRef}${moraRef}-${paid}))`;
}
function setupPayments_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PAYMENTS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); // sh.clear() NO borra validaciones
  // Modelo normalizado: la identidad vive en "Clientes". Nombre y DNI se muestran
  // aquí por comodidad (se autocompletan al elegir el ID Préstamo); no son la
  // fuente de verdad. El ID Pago se autogenera (P-0001, P-0002, …).
  const H = ['ID Pago', 'ID Préstamo', 'Fecha de Pago', 'Monto Pagado', 'Saldo Posterior', 'Recibo Enviado', 'Nombre', 'DNI', 'Enviar recibo', 'Generar recibo', 'Avisar pago 📲'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#38761d').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange(2, PP.DATE, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd')
    .setDataValidation(datePicker_()); // muestra el selector de fecha (calendario)
  sh.getRange(2, PP.AMOUNT, CFG.MAX_ROWS, 2).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(2, PP.SEND_RECEIPT, CFG.MAX_ROWS, 1).insertCheckboxes(); // casilla "Enviar recibo"
  sh.getRange(2, PP.GEN_RECEIPT, CFG.MAX_ROWS, 1).insertCheckboxes(); // casilla "Generar recibo" (sin correo)
  // "Saldo Posterior" VIVO: lo que falta pagar del préstamo tras cada pago, INCLUIDA la mora.
  const balF = []; for (let r = 2; r <= CFG.MAX_ROWS + 1; r++) balF.push([paymentBalanceFormula_(r)]);
  sh.getRange(2, PP.BALANCE, balF.length, 1).setFormulas(balF);
  [140, 90, 110, 120, 130, 150, 200, 110, 110, 110, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, PP.WA_PAYMENT).setNote('Enlace de WhatsApp con el AVISO DE PAGO listo (confirma el pago recibido, el saldo pendiente y el vencimiento). Se arma solo al registrar el pago; "sin teléfono" si el cliente no tiene un número válido.');
  sh.getRange(1, PP.SEND_RECEIPT).setNote('Tildá para ENVIAR por correo el recibo de pago de esa fila. Se destilda solo y deja el enlace en "Recibo Enviado". (Si el prestatario no tiene correo, usá "Generar recibo".)');
  sh.getRange(1, PP.GEN_RECEIPT).setNote('Tildá para GENERAR el PDF del recibo (con monto pagado y saldo pendiente) sin enviar correo. Se destilda solo y deja el enlace en "Recibo Enviado".');
  sh.getRange(1, PP.BALANCE).setNote('Lo que falta pagar del préstamo después de este pago, INCLUIDA la mora (= "Saldo con Mora" del prestatario en el último pago). Se calcula solo.');
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
function cuotasForLoan_(cs, loanId, loanDate, total, n, dueDays, opts) {
  // opts (opcional): { dates: Date[] }  → fechas de vencimiento explícitas (plan de pago
  //   reestructurado) en lugar de offsets sobre loanDate;
  //   { paidOffset: number } → pagos ANTERIORES al plan: se descuentan del Total Pagado
  //   del préstamo para que las cuotas (que ya son netas de esos pagos) no los dupliquen.
  opts = opts || {};
  const LID = colL_(CU.LOAN_ID), NUM = colL_(CU.NUM), AMT = colL_(CU.AMOUNT), DUE = colL_(CU.DUE);
  const B = CFG.SHEETS.BORROWERS, bIdL = colL_(PB.LOAN_ID), bPaidL = colL_(PB.PAID);
  const base = round2_(total / n);
  const rows = [];
  for (let k = 1; k <= n; k++) {
    const monto = (k === n) ? round2_(total - base * (n - 1)) : base;   // la última absorbe el redondeo
    const due = (opts.dates && opts.dates[k - 1]) ? opts.dates[k - 1] : addDays_(loanDate, dueDays[k - 1]);
    rows.push([loanId + '-' + k, loanId, k, due, monto, '', '', '', '']);
  }
  // Reemplaza, NO apila: borra las cuotas previas de este préstamo antes de escribir.
  // ATÓMICO: un lock de documento serializa las corridas CONCURRENTES (p. ej. tildar
  // "Plan de pago"/regenerar en varias filas, o clics rápidos → varios onEdit a la vez), y
  // flush() tras el borrado hace que getLastRow() refleje el borrado. Sin esto, los borrados y
  // altas se intercalaban → cuotas DUPLICADAS (mismo "ID Cuota") o filas viejas sin borrar.
  const paidOffset = round2_(Number(opts.paidOffset) || 0);
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30000); } catch (e) { logError_('cuotasForLoan_:lock', e); }
  try {
    deleteCuotasRowsForLoan_(cs, loanId);
    SpreadsheetApp.flush();
    const start = cs.getLastRow() + 1;
    cs.getRange(start, 1, rows.length, CUOTAS_HEADERS.length).setValues(rows);
    // Fórmulas vivas por fila (reparto del Total Pagado del préstamo, cuota más antigua primero).
    for (let i = 0; i < rows.length; i++) {
      const r = start + i;
      const pagadoAntes = `SUMIFS($${AMT}$2:$${AMT},$${LID}$2:$${LID},$${LID}${r},$${NUM}$2:$${NUM},"<"&$${NUM}${r})`;
      let loanPaid = `IFERROR(VLOOKUP($${LID}${r},'${B}'!$${bIdL}:$${bPaidL},${PB.PAID},FALSE),0)`;
      if (paidOffset > 0) loanPaid = `MAX(0,${loanPaid}-${paidOffset})`;
      cs.getRange(r, CU.PAID).setFormula(`=MIN($${AMT}${r},MAX(0,${loanPaid}-(${pagadoAntes})))`);
      cs.getRange(r, CU.BALANCE).setFormula(`=$${AMT}${r}-$${colL_(CU.PAID)}${r}`);
      cs.getRange(r, CU.STATE).setFormula(
        `=IF($${colL_(CU.BALANCE)}${r}<=0.009,"${CST.PAID}",IF(TODAY()>$${DUE}${r},"${CST.OVERDUE}","${CST.PENDING}"))`);
    }
    cs.getRange(start, CU.DUE, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    cs.getRange(start, CU.AMOUNT, rows.length, 3).setNumberFormat(CFG.CURRENCY_FMT);
    SpreadsheetApp.flush();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return rows.length;
}
/**
 * Elimina las filas de "Cuotas" de un préstamo (de abajo hacia arriba). Se usa al
 * reestructurar: las cuotas viejas (p. ej. VENCIDAS) se reemplazan por las del plan
 * — si quedaran, la fórmula de Estado de "Prestatarios" seguiría marcando VENCIDO.
 */
function deleteCuotasRowsForLoan_(cs, loanId) {
  if (!cs || cs.getLastRow() < 2) return 0;
  const id = String(loanId || '').trim(); if (!id) return 0;
  const ids = cs.getRange(2, CU.LOAN_ID, cs.getLastRow() - 1, 1).getValues();
  let removed = 0;
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() === id) { cs.deleteRow(i + 2); removed++; }
  }
  return removed;
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

/**
 * REPARACIÓN (idempotente, desde el menú) de la hoja "Cuotas":
 *  1) Elimina filas con "ID Cuota" DUPLICADO (residuo de cuando el generador apilaba en vez
 *     de reemplazar): por cada ID repetido conserva la fila de MAYOR "Monto Cuota" (la cuota
 *     real) y borra las fantasma (p. ej. $12,50 / $1,25).
 *  2) Re-alinea fechas: para préstamos de UNA sola cuota cuyo vencimiento no coincide con el
 *     "Vencimiento" del encabezado (desfase porque la firma movió la Fecha del Préstamo),
 *     regenera esa cuota desde la Fecha del Préstamo + plazo actuales (cuotasForLoan_ borra-y-recrea).
 * NO toca planes de varias cuotas (3 cuotas / reestructurados): tienen IDs distintos (-1/-2/-3).
 */
function repararCuotas_() {
  return guard_('repararCuotas_', function () {
    const ss = getSS_();
    const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS);
    if (!cs || cs.getLastRow() < 2) { ss.toast('No hay cuotas que reparar.', 'Reparar cuotas', 5); return { dups: 0, fixed: 0 }; }

    // 1) Duplicados por "ID Cuota": conservar el de mayor monto, borrar el resto (de abajo hacia arriba).
    let dups = 0;
    const n0 = cs.getLastRow() - 1;
    const vals = cs.getRange(2, 1, n0, CU.AMOUNT).getValues(); // A..E (ID Cuota … Monto Cuota)
    const byId = {};
    for (let i = 0; i < vals.length; i++) {
      const id = String(vals[i][CU.CUOTA_ID - 1]).trim(); if (!id) continue;
      (byId[id] = byId[id] || []).push({ row: i + 2, amount: Number(vals[i][CU.AMOUNT - 1]) || 0 });
    }
    const toDelete = [];
    Object.keys(byId).forEach(id => {
      const list = byId[id]; if (list.length < 2) return;
      list.sort((a, b) => b.amount - a.amount);            // mayor monto primero = cuota real
      for (let j = 1; j < list.length; j++) toDelete.push(list[j].row);
    });
    toDelete.sort((a, b) => b - a).forEach(r => { cs.deleteRow(r); dups++; });

    // 2) Re-alinear fechas de préstamos con UNA sola cuota (tras el dedup).
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    let fixed = 0;
    if (bs && bs.getLastRow() >= 2 && cs.getLastRow() >= 2) {
      const cv = cs.getRange(2, 1, cs.getLastRow() - 1, CU.DUE).getValues(); // A..D
      const byLoan = {};
      for (let i = 0; i < cv.length; i++) {
        const lid = String(cv[i][CU.LOAN_ID - 1]).trim(); if (!lid) continue;
        (byLoan[lid] = byLoan[lid] || []).push(cv[i][CU.DUE - 1]);
      }
      for (let row = 2; row <= bs.getLastRow(); row++) {
        const loan = readLoan_(bs, row);
        if (!loan.loanId) continue;
        const qs = byLoan[loan.loanId];
        if (!qs || qs.length !== 1) continue;              // sólo préstamos de una cuota
        if (!(loan.dueDate instanceof Date)) continue;
        const cur = qs[0] instanceof Date ? startOfDay_(qs[0]).getTime() : null;
        const want = startOfDay_(loan.dueDate).getTime();
        if (cur === want) continue;                        // ya coincide con el encabezado
        cuotasForLoan_(cs, loan.loanId, loan.loanDate, round2_(loan.totalDue), 1, [termDays_(loan.term)]);
        fixed++;
      }
    }
    ss.toast('Cuotas reparadas: ' + dups + ' duplicada(s) eliminada(s), ' + fixed + ' fecha(s) re-alineada(s).', 'Reparar cuotas', 8);
    return { dups: dups, fixed: fixed };
  });
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
  // Días de Atraso congelado-aware: si "Congelado el" / "Días congelados" existen en "Prestatarios"
  // (resueltos por encabezado, posición variable en hojas migradas), se recorta HOY a la fecha de
  // congelamiento y se descuentan los días ya saltados; si no, se usa la fórmula clásica.
  const bsSheet = ss.getSheetByName(B), Hb = bsSheet ? headerIndex_(bsSheet) : {};
  const congIdx = colByAny_(Hb, ['Congelado el']), fdIdx = colByAny_(Hb, ['Días congelados', 'Dias congelados']);
  let diasFormula;
  if (congIdx && fdIdx) {
    const congRef = `N(VLOOKUP($B$4,'${B}'!$A:$${colL_(congIdx)},${congIdx},FALSE))`;
    const fdRef = `N(VLOOKUP($B$4,'${B}'!$A:$${colL_(fdIdx)},${fdIdx},FALSE))`;
    diasFormula = `=IFERROR(LET(id,$B$4,tot,$B$10,due,VLOOKUP($B$4,'${B}'!$A:$I,9,FALSE),cong,${congRef},fdays,${fdRef},IF(due="","",IF($B$14>0.009,MAX(0,INT(IF(cong>0,MIN(TODAY(),cong),TODAY()))-INT(due)-fdays),LET(d,FILTER('${P}'!$C$2:$C,'${P}'!$B$2:$B=id),frz,MIN(MAP(d,LAMBDA(x,IF(SUMIFS('${P}'!$D:$D,'${P}'!$B:$B,id,'${P}'!$C:$C,"<="&x)>=tot-0.009,x,"")))),MAX(0,INT(frz)-INT(due)-fdays))))),"")`;
  } else {
    diasFormula = `=IFERROR(LET(id,$B$4,tot,$B$10,due,VLOOKUP($B$4,'${B}'!$A:$I,9,FALSE),IF(due="","",IF($B$14>0.009,MAX(0,INT(TODAY())-INT(due)),LET(d,FILTER('${P}'!$C$2:$C,'${P}'!$B$2:$B=id),frz,MIN(MAP(d,LAMBDA(x,IF(SUMIFS('${P}'!$D:$D,'${P}'!$B:$B,id,'${P}'!$C:$C,"<="&x)>=tot-0.009,x,"")))),MAX(0,INT(frz)-INT(due)))))),"")`;
  }
  const rows = [
    ['Prestatario', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$B,2,FALSE),"")`],     // 6
    ['DNI', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$D,4,FALSE),"")`],             // 7
    ['Correo', `=IFERROR(VLOOKUP(${cId},'${C}'!$A:$C,3,FALSE),"")`],          // 8
    ['Capital', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$E,5,FALSE),"")`],           // 9
    ['Total a Pagar', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$K,11,FALSE),"")`],     // 10
    ['Total Pagado', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$L,12,FALSE),"")`],     // 11
    // Días de Atraso: respeta la FECHA DE PAGO. Si sigue impago (Saldo Pendiente B14 > 0), HOY −
    // Vencimiento; si está saldado, los días entre el Vencimiento y el pago que CUBRIÓ EL TOTAL
    // BASE (no el último pago: los posteriores sólo cubren mora). El acumulado por fecha se arma
    // con MAP+SUMIFS sobre los pagos del préstamo. Espeja computeOutstanding_ / chargedMora_.
    ['Días de Atraso', diasFormula], // 12
    // Recargo por mora = días de atraso MENOS la gracia, × recargo diario × Total, con
    // TOPE = (Tope de mora %) × Capital ($B$9). Espeja computeOutstanding_ (gracia + tope en capital).
    ['Recargo por Mora (acum.)', `=IFERROR(IF(OR($B$12="",$B$12<=0),0,MIN(MAX(0,$B$12-IFERROR(VLOOKUP("Días de gracia antes de mora",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE),0))*$B$10*IFERROR(VLOOKUP("Recargo por Mora diario (%)",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)/100,0.05),$B$9*IFERROR(VLOOKUP("Tope de mora (% del total a devolver)",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)/100,1))),"")`], // 13
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
  'Generar PDF 👁', 'Enviar ✉', 'Último PDF', 'Recordar por WhatsApp', 'Pedir congelamiento 📲', 'Pedir plan de pago 📋'];
const REM_GEN_COL = 10, REM_SEND_COL = 11, REM_LINK_COL = 12, REM_WA_COL = 13, REM_FREEZE_COL = 14, REM_PLANREQ_COL = 15;
const REM_HEADER_ROW = 3, REM_DATA_START = 4; // fila 1 título, fila 2 control, fila 3 encabezados

/** Construye la estructura (título, casilla de refresco y encabezados) y llena la lista. */
function setupReminders_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.REMINDER); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange('A1').setValue('ESTUDIO DE RECORDATORIOS DE PAGO').setFontSize(16).setFontWeight('bold').setFontColor('#b45f06');
  sh.getRange('C1').setValue('Préstamos vencidos (rojo) o próximos a vencer (amarillo). Tilde 👁 para generar el PDF o ✉ para enviarlo por correo; 📲 abre WhatsApp con el recordatorio listo.').setFontColor('#666').setWrap(true);
  sh.getRange('A2').setValue('🔄 Actualizar lista →').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.getRange('A1:A2').setWrap(true);
  sh.getRange(REM_HEADER_ROW, 1, 1, REM_HEADERS.length).setValues([REM_HEADERS])
    .setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(REM_HEADER_ROW);
  [110, 180, 130, 120, 110, 120, 130, 150, 210, 110, 90, 130, 170].forEach((w, i) => sh.setColumnWidth(i + 1, w));
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
  const feePct = lateFeeRate_(), capFrac = moraCapFrac_(), rows = [], overdueFlags = [], waCells = [], freezeCells = [], planCells = [];
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
  // Lectura EN BLOQUE de "Prestatarios" (una sola vez) + clientes cacheados: evita un getRange por
  // fila (antes: readLoan_ + getValue del saldo por cada préstamo). Columnas por nombre de
  // encabezado con los índices PB de respaldo (mismo criterio que readLoan_).
  const H = headerIndex_(bs);
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
  const balC = PB.BALANCE;
  const maxCol = Math.max(idC, cliC, capC, plzC, tasC, fecC, venC, intC, totC, balC);
  const bdata = bs.getRange(2, 1, last - 1, maxCol).getValues();
  for (let i = 0; i < bdata.length; i++) {
    const v = bdata[i], g = c => v[c - 1];
    const loanId = String(g(idC) || '').trim(); if (!loanId) continue;
    // Vencimiento, tasa, interés y total con el MISMO respaldo que readLoan_ (se calculan si la
    // celda está en blanco, p. ej. préstamo recién aprobado sin fórmulas todavía).
    const term = Number(g(plzC)) || 0;
    const ld = g(fecC) instanceof Date ? g(fecC) : new Date(g(fecC));
    const loanDue = g(venC) instanceof Date ? g(venC) : dueDateForTerm_(term || 1, ld);
    if (!(loanDue instanceof Date) || isNaN(loanDue.getTime())) continue;
    const principal = Number(g(capC)) || 0;
    let rate = Number(g(tasC)) || 0; if (!rate) rate = loanRateForTerm_(term);
    let interest = Number(g(intC)) || 0; if (!interest) interest = round2_(principal * rate);
    let loanTotal = Number(g(totC)) || 0; if (!loanTotal) loanTotal = round2_(principal + interest);
    const cli = getClienteById_(String(g(cliC) || '').trim()) || {};
    const loan = { loanId: loanId, name: cli.name || '', email: cli.email || '', phone: cli.phone || '', dueDate: loanDue, totalDue: loanTotal };
    // Base de aviso: si el préstamo tiene cronograma, se usa la PRÓXIMA cuota impaga
    // (fecha, monto y saldo de esa cuota); si no, el vencimiento y saldo del préstamo.
    let dueDate = loanDue, totalDue = loanTotal, cuotaTxt = '';
    let out = Number(g(balC)) || 0;
    const qs = cuotasByLoan[loanId];
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
    rows.push([loanId, loan.name + cuotaTxt, dueDate, daysUntil, overdue ? ST.OVERDUE : 'PRÓXIMO',
      totalDue, out, feeDay, loan.email, false, false, '', '', '', '']);
    overdueFlags.push(overdue);
    // Enlace de WhatsApp con el recordatorio listo (según el saldo pendiente de esta fila).
    waCells.push(remWhatsappFormula_(loan, dueDate, out, daysUntil));
    // Enlace al formulario de congelamiento SOLO para filas vencidas (daysUntil < 0).
    freezeCells.push(overdue ? (freezeFormWaFormula_(loan) || '') : '');
    // Enlace al formulario de PLAN DE PAGO: sólo cuando (aprox.) llegó al 100% de mora. El formulario
    // vuelve a verificar la elegibilidad exacta. atCap ≈ días de atraso × recargo diario ≥ tope.
    const atCapApprox = overdue && ((-daysUntil) * feePct >= capFrac - 1e-9);
    planCells.push(atCapApprox ? (planFormWaFormula_(loan) || '') : '');
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
  // Columna "Recordar por WhatsApp": enlace clicable por fila; "sin teléfono" si no hay número válido.
  const sortedWa = order.map(i => waCells[i]);
  sh.getRange(REM_DATA_START, REM_WA_COL, n, 1).setFormulas(sortedWa.map(x => [x || '']));
  sortedWa.forEach((x, i) => {
    if (!x) sh.getRange(REM_DATA_START + i, REM_WA_COL).setValue('sin teléfono').setFontColor('#999');
  });
  // Columna "Pedir congelamiento 📲": enlace al formulario, sólo en filas vencidas.
  const sortedFreeze = order.map(i => freezeCells[i]);
  sh.getRange(REM_DATA_START, REM_FREEZE_COL, n, 1).setFormulas(sortedFreeze.map(x => [x || '']));
  // Columna "Pedir plan de pago 📋": enlace al formulario, sólo cuando (aprox.) llegó al 100% de mora.
  const sortedPlan = order.map(i => planCells[i]);
  sh.getRange(REM_DATA_START, REM_PLANREQ_COL, n, 1).setFormulas(sortedPlan.map(x => [x || '']));
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

/** Número para wa.me de un móvil argentino: 549 + área + número (sin el «+»). '' si no es válido.
 *  normPhoneE164_ devuelve «+54» + 10 dígitos (área+número, SIN el 9 móvil); wa.me lo necesita CON el 9.
 *  El número es SIEMPRE el del prestatario (resuelto de «Clientes»). */
function waNumberAR_(phone) {
  const e = (typeof normPhoneE164_ === 'function') ? normPhoneE164_(phone) : '';
  return e ? '549' + e.slice(3) : '';
}

/** Fórmula «=HYPERLINK» que abre WhatsApp con un recordatorio breve (tono informal) para la fila.
 *  Devuelve '' si el prestatario no tiene un teléfono válido (la celda mostrará «sin teléfono»). */
function remWhatsappFormula_(loan, dueDate, amount, daysUntil) {
  const num = waNumberAR_(loan && loan.phone);
  if (!num) return '';
  const name = String((loan && loan.name) || '').trim() || 'cliente';
  const fecha = (dueDate instanceof Date) ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  const monto = fmtMoney_(amount), empresa = companyName_();
  let msg;
  if (daysUntil > 0)      msg = 'Hola ' + name + ' 👋 Te recordamos que tu préstamo vence el ' + fecha + '. Total a pagar: ' + monto + '. ¡Gracias! ' + empresa;
  else if (daysUntil === 0) msg = 'Hola ' + name + ' 👋 Tu préstamo vence hoy (' + fecha + '). Total a pagar: ' + monto + '. ¡Gracias! ' + empresa;
  else                    msg = 'Hola ' + name + ' 👋 Tu préstamo venció el ' + fecha + ' (hace ' + Math.abs(daysUntil) + ' día(s)). Saldo a pagar: ' + monto + '. Por favor regularizá tu pago. ' + empresa;
  return '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📲 WhatsApp")';
}

/** Fórmula «=HYPERLINK» que abre WhatsApp al prestatario con el ENLACE al formulario de
 *  congelamiento (?page=congelar&loan=…). '' si no hay teléfono válido o la app web no está publicada. */
function freezeFormWaFormula_(loan) {
  const num = waNumberAR_(loan && loan.phone);
  const base = webAppBaseUrl_();
  if (!num || !base) return '';
  const name = String((loan && loan.name) || '').trim() || 'cliente';
  const url = base + '?page=congelar&loan=' + encodeURIComponent(loan.loanId);
  const msg = 'Hola ' + name + ' 👋 Tu préstamo ' + loan.loanId + ' está vencido. Si no podés pagar ahora, ' +
    'podés pedir CONGELAR la mora (que deje de crecer) acá, con un plazo de hasta 2 semanas: ' + url + ' — ' + companyName_();
  return '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📲 Congelar")';
}

/** Fórmula «=HYPERLINK» que abre WhatsApp con el ENLACE al formulario de PLAN DE PAGO (?page=plan).
 *  Sólo tiene sentido en préstamos al 100% de mora. '' si no hay teléfono o app web publicada. */
function planFormWaFormula_(loan) {
  const num = waNumberAR_(loan && loan.phone);
  const base = webAppBaseUrl_();
  if (!num || !base) return '';
  const name = String((loan && loan.name) || '').trim() || 'cliente';
  const url = base + '?page=plan&loan=' + encodeURIComponent(loan.loanId);
  const msg = 'Hola ' + name + ' 👋 Tu préstamo ' + loan.loanId + ' llegó al tope de mora. Podés pedir un ' +
    'PLAN DE PAGO (pago único, 4 o 6 cuotas) y frenar la mora acá: ' + url + ' — ' + companyName_();
  return '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📋 Plan de pago")';
}

/** Días de la semana en español (0=domingo) para los textos de fecha límite de cobranza. */
const DIAS_SEMANA_ES = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

/** Aviso formal de COBRANZA (mora tardía) para un préstamo vencido, según la plantilla acordada:
 *  encabezado con datos del préstamo y estado de pagos, DETALLE de la deuda, cláusula del tope del
 *  recargo (= % × CAPITAL, con fecha de alcance y monto fijo final), TRES opciones (1: cancelar sin
 *  mora antes del próximo día hábil; 2: cancelar hoy con mora; 3: congelar), medios de pago, plazos de
 *  recibo/carta de cancelación, derecho a revisar montos, reserva de acciones de cobranza y firma de
 *  los prestamistas. m = { base, mora, totalPaid, days, deadlineStr, out, freezeUrl, isFrozen,
 *  capAmount, maxTotal, dailyAmount, capReachStr, todayStr, loanDateStr, feePctText, alias, contact,
 *  sig1, sig2 }. Devuelve {subject, body}. */
function collectionMessage_(loan, m) {
  const nl = '\n', emp = companyName_();
  const venc = (loan.dueDate instanceof Date) ? fmtDate_(loan.dueDate) : '';
  const subject = 'Gestión de cobranza — Préstamo ' + loan.loanId + ' (vencido) · ' + emp;
  const pagoStatus = (Number(m.totalPaid) || 0) <= 0
    ? 'A la fecha, no registramos ningún pago.'
    : 'A la fecha, registramos pagos por ' + fmtMoney_(m.totalPaid) + '; resta cancelar el saldo indicado.';
  let b = 'Estimado/a ' + (loan.name || 'cliente') + ':' + nl + nl;
  b += 'Le escribimos respecto del préstamo ' + loan.loanId +
    (m.loanDateStr ? (', firmado el ' + m.loanDateStr) : '') +
    (venc ? (', con vencimiento el ' + venc) : '') + '. ' + pagoStatus + nl + nl;
  b += 'Detalle de la deuda al ' + m.todayStr + ':' + nl;
  b += '• Monto del capital: ' + fmtMoney_(loan.principal) + nl;
  b += '• Monto total acordado (capital + interés): ' + fmtMoney_(loan.totalDue) + nl;
  b += '• Pagado a la fecha: ' + fmtMoney_(m.totalPaid) + nl;
  b += '• Recargo por mora (' + m.days + ' día(s) × ' + fmtMoney_(m.dailyAmount) + '): ' + fmtMoney_(m.mora) + nl;
  b += '• Monto total adeudado a hoy: ' + fmtMoney_(m.out) + nl + nl;
  if (m.isFrozen) {
    b += 'Su préstamo tiene la mora CONGELADA: el recargo no sigue creciendo y el total adeudado se mantiene en ' + fmtMoney_(m.out) + '.' + nl + nl;
  } else {
    b += 'Conforme a los términos del contrato que usted firmó, el recargo por mora es del ' + m.feePctText +
      ' diario sobre el saldo de ' + fmtMoney_(loan.totalDue) + ', con un tope máximo de ' + fmtMoney_(m.capAmount) + '. ' +
      (m.capReachStr ? ('Dicho tope se alcanzará el ' + m.capReachStr + '; a partir de esa fecha, el monto total quedará fijo en ' + fmtMoney_(m.maxTotal) + ', ') : ('Alcanzado ese tope, el monto total quedará fijo en ' + fmtMoney_(m.maxTotal) + ', ')) +
      'más los costos de cobranza y honorarios legales que correspondan.' + nl + nl;
  }
  b += 'Usted tiene tres opciones:' + nl + nl;
  b += 'OPCIÓN 1 — Cancelación con bonificación total del recargo por mora' + nl;
  b += 'Si abona ' + fmtMoney_(m.base) + ' antes del cierre del ' + m.deadlineStr + ', le bonificamos la totalidad del recargo por mora (' + fmtMoney_(m.mora) + ') y el préstamo quedará saldado.' + nl + nl;
  b += 'OPCIÓN 2 — Cancelación al día de hoy' + nl;
  b += 'Si abona ' + fmtMoney_(m.out) + ', el préstamo quedará saldado de inmediato.' + nl + nl;
  b += 'OPCIÓN 3 — Congelamiento' + nl;
  if (m.isFrozen) {
    b += 'Su préstamo ya se encuentra congelado. Si acordamos una fecha de pago y necesita ajustarla, responda este mensaje.' + nl + nl;
  } else {
    b += 'Si por el momento no puede afrontar ninguna de las opciones anteriores, puede solicitar el congelamiento de su préstamo: los recargos por mora dejarán de acumularse y se fijará una fecha de pago concreta. Le enviaremos los términos y condiciones para su revisión; nada queda firme hasta su confirmación.' + nl;
    b += 'Solicitar esta opción: ' + (m.freezeUrl || '(solicite el enlace al prestamista)') + nl + nl;
  }
  if (!m.isFrozen && m.planUrl) {
    b += 'OPCIÓN 4 — Plan de pago en cuotas' + nl;
    b += 'Si necesita pagar en cuotas, puede solicitar un plan de pago: elige la frecuencia (semanal, quincenal o mensual) y la cantidad de cuotas, y la mora queda suspendida mientras cumpla el plan. Lo revisamos y, si corresponde, le enviamos el acuerdo para firmar.' + nl;
    b += 'Solicitar esta opción: ' + m.planUrl + nl + nl;
  }
  b += 'Medios de pago' + nl;
  b += (m.alias ? ('Mercado Pago — alias ' + m.alias + ', según lo estipulado en su acuerdo. ') : '') +
    'Si prefiere pagar en efectivo o por transferencia, avísenos y le enviaremos los datos necesarios dentro de las 24 horas.' + nl + nl;
  b += 'Una vez enviado el comprobante de pago, le remitiremos el recibo dentro de las 48 horas. Al cancelarse totalmente el préstamo, le entregaremos una carta de cancelación de deuda dentro de los 5 días hábiles.' + nl + nl;
  b += 'Si considera que alguno de estos montos es incorrecto, responda este mensaje para que podamos revisarlo.' + nl + nl;
  b += 'De no recibir el pago ni respuesta, nos reservamos el derecho de exigir la totalidad del saldo adeudado e iniciar los procedimientos de cobranza previstos en el acuerdo que usted firmó.' + nl + nl;
  b += 'Quedamos a su disposición.' + nl + nl;
  b += (m.sig2 ? (m.sig1 + ' y ' + m.sig2) : m.sig1) + nl;
  b += 'Prestamistas — Préstamo ' + loan.loanId + nl;
  b += emp + (m.contact ? (' · ' + m.contact) : '');
  return { subject: subject, body: b };
}

/** Correo GENÉRICO para proponerle al prestatario un PLAN DE PAGO EN CUOTAS: explica cómo funciona,
 *  los pasos, y el enlace al formulario del plan (?page=plan). Se abre en Gmail (no se envía).
 *  pctInt = % del pago inicial (planDownPct_). Devuelve {subject, body}. */
function planEmailMessage_(loan, planUrl, pctInt) {
  const nl = '\n', emp = companyName_();
  const subject = 'Plan de pago en cuotas — Préstamo ' + loan.loanId + ' · ' + emp;
  let b = 'Hola ' + (loan.name || 'cliente') + ',' + nl + nl;
  b += 'Tu préstamo ' + loan.loanId + ' está vencido. Podés regularizarlo con un PLAN DE PAGO EN CUOTAS: pagás de a poco y la mora deja de crecer.' + nl + nl;
  b += 'Cómo funciona:' + nl;
  b += '• Elegís la frecuencia (semanal, quincenal o mensual) y la cantidad de cuotas.' + nl;
  b += '• Mientras cumplas el plan, la mora queda SUSPENDIDA (no sigue creciendo).' + nl;
  b += '• El plan cubre el total del préstamo (capital + interés), sin la mora.' + nl + nl;
  b += 'Pasos:' + nl;
  b += '1) Abrí este enlace e ingresá tu DNI: ' + (planUrl || '(solicitá el enlace al prestamista)') + nl;
  b += '2) Elegí la frecuencia y la cantidad de cuotas, y enviá la solicitud.' + nl;
  b += '3) Pagá el pago inicial (' + pctInt + '% del total) para activar el plan.' + nl;
  b += '4) Firmá el Acuerdo (te enviamos el enlace cuando aprobemos la solicitud).' + nl;
  b += '5) Pagá las cuotas en fecha. Al terminar, el préstamo queda saldado.' + nl + nl;
  b += 'Si tenés dudas, respondé este correo.' + nl + nl;
  b += emp;
  return { subject: subject, body: b };
}

/** Fecha de "Plan desde" de un préstamo (Acuerdo de plan de pago vigente) o null. Espeja frozenInfo_. */
function loanPlanSince_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh ? sh.getLastRow() : 0;
  if (last < 2) return null;
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const psC = colByAny_(H, ['Plan desde']);
  if (!psC) return null;
  const id = String(loanId).trim();
  const data = sh.getRange(2, 1, last - 1, Math.max(idC, psC)).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][idC - 1]).trim() === id) {
    const v = data[i][psC - 1]; return v instanceof Date ? v : null;
  }
  return null;
}

/** Estado completo de un préstamo para el aviso de pago por WhatsApp. Reúne saldo base, mora (0 si
 *  está en plan por "Mora (ajuste)"=0; congelada si está congelado), congelamiento, plan, cuotas y —
 *  si quedó saldado— el límite vigente del cliente para un nuevo préstamo (graduación V-32). */
function loanWaStatus_(loan) {
  const now = new Date();
  const totalDue = round2_(Number(loan.totalDue) || 0);
  const pays = (typeof loanPayments_ === 'function') ? loanPayments_(loan.loanId) : [];
  const totalPaid = round2_(pays.reduce((s, p) => s + (Number(p.amount) || 0), 0));
  const baseOwed = Math.max(0, round2_(totalDue - totalPaid));          // capital+interés pendiente (sin mora)
  const frz = frozenInfo_(loan.loanId);
  const frozen = frz.since instanceof Date;
  const planSince = loanPlanSince_(loan.loanId);
  const onPlan = planSince instanceof Date;
  const mora = round2_(accruedMora_(loan, now));                        // 0 si plan (ajuste=0); congelada si frozen
  const late = (loan.dueDate instanceof Date) && (typeof daysLate_ === 'function') && daysLate_(loan.dueDate, now) > 0;
  const owedMora = onPlan ? 0 : mora;                                   // en plan la mora queda suspendida
  const toClear = round2_(baseOwed + owedMora);
  const paidOff = toClear <= 0.009;
  const cuotas = (typeof cuotasDeLoan_ === 'function') ? cuotasDeLoan_(loan.loanId) : [];
  const s = { now, totalDue, totalPaid, baseOwed, frz, frozen, planSince, onPlan, mora, owedMora,
    late, toClear, paidOff, cuotas, multiCuota: cuotas.length > 1, limit: null };
  if (paidOff && typeof graduationMax_ === 'function') {
    try { s.limit = graduationMax_(normDni_(loan.dni)); } catch (e) { s.limit = null; }
  }
  return s;
}

/** Líneas de estado por cuota para el mensaje: ✅ pagada · 🔸 parcial · ⬜ pendiente (con vencimiento). */
function cuotaLines_(cuotas) {
  const tz = Session.getScriptTimeZone(), n = cuotas.length;
  return cuotas.map(function (c) {
    const etq = 'Cuota ' + c.num + '/' + n;
    const venc = c.due instanceof Date ? (' (vence ' + Utilities.formatDate(c.due, tz, 'dd/MM/yyyy') + ')') : '';
    if (c.saldo <= 0.009) return '✅ ' + etq + ': ' + fmtMoney_(c.monto) + ' PAGADA';
    if (c.pagado > 0.009) return '🔸 ' + etq + ': pagaste ' + fmtMoney_(c.pagado) + ' de ' + fmtMoney_(c.monto) + ', falta ' + fmtMoney_(c.saldo) + venc;
    return '⬜ ' + etq + ': ' + fmtMoney_(c.monto) + ' pendiente' + venc;
  });
}

/** Arma el TEXTO del aviso de pago según el estado del préstamo (saldado / debe / cuotas / congelado /
 *  en plan). s = loanWaStatus_(loan). Usa saltos de línea reales (encodeURIComponent → %0A en wa.me). */
function paymentWaMsg_(loan, s, payDate, amount) {
  const tz = Session.getScriptTimeZone(), nl = '\n', empresa = companyName_();
  const name = String((loan && loan.name) || '').trim() || 'cliente';
  const fPago = (payDate instanceof Date) ? Utilities.formatDate(payDate, tz, 'dd/MM/yyyy') : String(payDate || '');
  const fVenc = (loan && loan.dueDate instanceof Date) ? Utilities.formatDate(loan.dueDate, tz, 'dd/MM/yyyy') : '';
  let msg = 'Hola ' + name + ' 👋 Confirmamos tu pago de ' + fmtMoney_(amount) + (fPago ? (' del ' + fPago) : '') +
    ' en el préstamo ' + loan.loanId + '.';
  if (s.paidOff) {
    msg += nl + '✅ ¡Tu préstamo quedó SALDADO! Gracias por tu cumplimiento.';
    if (s.limit != null) msg += nl + 'Tu límite actual para un nuevo préstamo es ' + fmtMoney_(s.limit) + '.';
    const base = webAppBaseUrl_();
    if (base) msg += nl + 'Pedí tu próximo préstamo acá: ' + base;
  } else if (s.frozen) {
    const fSince = s.frz.since instanceof Date ? Utilities.formatDate(s.frz.since, tz, 'dd/MM/yyyy') : '';
    msg += nl + '❄ Tu préstamo está CONGELADO' + (fSince ? (' desde el ' + fSince) : '') + ': la mora NO sigue creciendo.';
    msg += nl + 'Saldo pendiente: ' + fmtMoney_(s.baseOwed) + (s.mora > 0.009 ? (' + mora congelada ' + fmtMoney_(s.mora)) : '') + '.';
    msg += nl + 'Total para cancelar: ' + fmtMoney_(s.toClear) + '.';
  } else if (s.onPlan) {
    msg += nl + '📋 Estás en un PLAN DE PAGO: la mora quedó SUSPENDIDA mientras cumplas las cuotas.';
    if (s.cuotas.length) { msg += nl + 'Tus cuotas:'; cuotaLines_(s.cuotas).forEach(function (l) { msg += nl + l; }); }
    msg += nl + 'Total del plan pendiente: ' + fmtMoney_(s.baseOwed) + '.';
  } else if (s.multiCuota) {
    msg += nl + 'Estado de tus cuotas:';
    cuotaLines_(s.cuotas).forEach(function (l) { msg += nl + l; });
    if (s.late && s.mora > 0.009) {
      msg += nl + 'Falta pagar: ' + fmtMoney_(s.baseOwed) + ' + mora por atraso ' + fmtMoney_(s.mora) + '.';
      msg += nl + 'Total para cancelar: ' + fmtMoney_(s.toClear) + '.';
    } else {
      msg += nl + 'Falta pagar en total: ' + fmtMoney_(s.baseOwed) + '.';
    }
    if (fVenc) msg += nl + (s.late ? 'Venció el ' : 'Vence el ') + fVenc + '.';
  } else {
    // Pago único, todavía con saldo.
    if (s.late && s.mora > 0.009) {
      msg += nl + 'Saldo pendiente: ' + fmtMoney_(s.baseOwed) + ' + mora por atraso ' + fmtMoney_(s.mora) + '.';
      msg += nl + 'Te faltan ' + fmtMoney_(s.toClear) + ' para cancelarlo.';
    } else {
      msg += nl + 'Saldo pendiente: ' + fmtMoney_(s.baseOwed) + '. Te faltan ' + fmtMoney_(s.baseOwed) + ' para cancelarlo.';
    }
    if (fVenc) msg += nl + (s.late ? 'Venció el ' : 'Vence el ') + fVenc + '.';
  }
  return msg + nl + empresa;
}

/** Fórmula «=HYPERLINK» que abre WhatsApp al prestatario con el AVISO DE PAGO adecuado al estado del
 *  préstamo (saldado → felicita + límite + enlace a nuevo préstamo; debe → saldo/mora/faltante; cuotas
 *  → detalle por cuota; congelado / en plan → aviso correspondiente). '' si no hay teléfono válido.
 *  @param status (opcional) loanWaStatus_(loan) ya calculado (para reusarlo por préstamo en lotes). */
function paymentWaFormula_(loan, payDate, amount, status) {
  const num = waNumberAR_(loan && loan.phone);
  if (!num) return '';
  const s = status || loanWaStatus_(loan);
  const msg = paymentWaMsg_(loan, s, payDate, amount);
  const label = s.paidOff ? '📲 Avisar saldado' : '📲 Avisar pago';
  return '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","' + label + '")';
}

/** (Re)completa la columna "Avisar pago 📲" de TODA la hoja Pagos con el enlace de WhatsApp por fila
 *  (según el pago y el saldo resultante). "sin teléfono" si el cliente no tiene número válido; vacío si
 *  la fila no tiene préstamo + monto. Crea la columna si falta. Devuelve la cantidad de filas tratadas. */
function fillPaymentWaLinks_(ss) {
  ss = ss || getSS_();
  const sh = ss.getSheetByName(CFG.SHEETS.PAYMENTS); if (!sh) return 0;
  const waCol = clientesHeaderCol_(sh, 'Avisar pago 📲', true);
  if (!waCol) return 0;
  sh.getRange(1, waCol).setBackground('#38761d').setFontColor('#fff'); // iguala el encabezado de Pagos
  const last = sh.getLastRow(); if (last < 2) { sh.setColumnWidth(waCol, 130); return 0; }
  const vals = sh.getRange(2, 1, last - 1, PP.BALANCE).getValues(); // ID Pago..Saldo Posterior
  const out = [], cache = {};
  for (let i = 0; i < vals.length; i++) {
    const id = String(vals[i][PP.LOAN_ID - 1] || '').trim();
    const amt = Number(vals[i][PP.AMOUNT - 1]) || 0;
    if (!id || !amt) { out.push(['']); continue; }
    if (!(id in cache)) { const loan = findLoanById_(id); cache[id] = { loan: loan, status: loan ? loanWaStatus_(loan) : null }; }
    const c = cache[id];
    const f = c.loan ? paymentWaFormula_(c.loan, vals[i][PP.DATE - 1], amt, c.status) : '';
    out.push([f || (c.loan ? 'sin teléfono' : '')]); // setValues trata "=…" como fórmula
  }
  sh.getRange(2, waCol, out.length, 1).setValues(out);
  sh.setColumnWidth(waCol, 130);
  return out.length;
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
 * Instala/recrea SOLO la hoja "Estudio de Estados" (título, selector de ID Préstamo, datos por
 * VLOOKUP y casillas ① PDF / ② Enviar) SIN borrar datos de Prestatarios ni Pagos. Acción LIGERA:
 * reaplica las fórmulas del estado de cuenta (p. ej. el "Recargo por Mora cobrado" que respeta la
 * Fecha de Pago) sin correr el ⑤ completo que puede exceder el tiempo.
 */
function createStatementsSheet_() {
  const ss = getSS_();
  setupStatements_(ss);
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.STATEMENTS));
  ss.toast('Hoja "Estudio de Estados" actualizada. Sus datos de Prestatarios y Pagos no fueron modificados.', 'Listo', 6);
}

/**
 * Instala/recrea SOLO la hoja "Panel" (indicadores del prestamista) SIN borrar datos de
 * Prestatarios ni Pagos. Acción LIGERA: reaplica el layout e indicadores (incluida la nueva
 * categoría "Préstamos en plan") sin correr el ⑤ completo que puede exceder el tiempo.
 */
function createPanelSheet_() {
  const ss = getSS_();
  setupPanel_(ss);
  SpreadsheetApp.flush();
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.PANEL));
  ss.toast('Panel actualizado (incluye "Préstamos en plan"). Sus datos no fueron modificados.', 'Listo', 6);
}

/**
 * Reconstruye el Panel para instalar/actualizar el bloque "Vencimientos por período"
 * (cuotas y monto que vencen esta semana [lun–dom], este mes y el mes próximo). Reutiliza
 * setupPanel_ (que dibuja todo el Panel), activa la hoja y posiciona la vista en el bloque.
 * NO modifica datos de Prestatarios, Pagos ni Cuotas. Sin guion bajo final para que aparezca
 * en la lista "Ejecutar" del editor de Apps Script.
 */
function actualizarVencimientosPanel() {
  const ss = getSS_();
  setupPanel_(ss);
  SpreadsheetApp.flush();
  const sh = ss.getSheetByName(CFG.SHEETS.PANEL);
  ss.setActiveSheet(sh);
  try { sh.setActiveRange(sh.getRange('A21')); } catch (e) {}
  ss.toast('«Vencimientos por período» actualizado (semana/mes/mes próximo). Sus datos no fueron modificados.', 'Listo', 6);
}

/**
 * Instala/actualiza la hoja "En Plan": lista los préstamos EN PLAN (reestructurados) con las
 * casillas "Plan de pago 📅 (3 cuotas)" (= regenerar cuotas, sin correo) y "PDF de cuotas 🧾"
 * (deja el enlace del PDF). Acción LIGERA, SIN borrar datos de Prestatarios ni Pagos.
 */
function createEnPlanSheet_() {
  const ss = getSS_();
  setupEnPlan_(ss);
  rebuildEnPlanSheet_(ss);
  SpreadsheetApp.flush();
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.PLAN));
  ss.toast('Hoja "En Plan" actualizada (préstamos reestructurados).', 'Listo', 6);
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
  const rep = repararPagos_();                           // pagos tipeados como texto → número; IDs sin espacios
  try { refreshAll(true); } catch (e) { logError_('repairTotals:refreshAll', e); } // recalcula tableros
  SpreadsheetApp.getUi().alert('Reparación de totales',
    'Filas TOTAL eliminadas — Prestatarios: ' + nB + ', Resumen: ' + nR + '.\n' +
    'Pagos normalizados — montos texto→número: ' + rep.monto + ', IDs recortados: ' + rep.ids + '.\n' +
    'Banner y tableros recalculados. Los datos de préstamos, pagos y clientes no se modificaron.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Normaliza la hoja "Pagos" para que ningún pago quede "perdido": convierte el
 * "Monto Pagado" que quedó como TEXTO a número real (así el SUMIF de "Total Pagado"
 * y el motor de mora lo cuentan) y recorta espacios del "ID Préstamo". También
 * recorta el "ID Préstamo" de "Prestatarios". Idempotente; no borra ni reasigna pagos.
 * Devuelve {monto, ids} con la cantidad de celdas corregidas.
 */
function repararPagos_() {
  const ss = getSS_(); let fixMonto = 0, fixId = 0;
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (pg && pg.getLastRow() > 1) {
    const PH = (typeof headerIndex_ === 'function') ? headerIndex_(pg) : null;
    const idC = (PH && typeof colByAny_ === 'function' && colByAny_(PH, ['ID Préstamo', 'ID Prestamo'])) || PP.LOAN_ID;
    const amtC = (PH && typeof colByAny_ === 'function' && colByAny_(PH, ['Monto Pagado'])) || PP.AMOUNT;
    const n = pg.getLastRow() - 1;
    const ids = pg.getRange(2, idC, n, 1).getValues();
    const amts = pg.getRange(2, amtC, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      const rawId = ids[i][0], idT = String(rawId == null ? '' : rawId).trim();
      if (idT && idT !== rawId) { pg.getRange(2 + i, idC).setValue(idT); fixId++; }
      const rawAmt = amts[i][0];
      if (rawAmt !== '' && rawAmt != null && typeof rawAmt !== 'number') {
        pg.getRange(2 + i, amtC).setValue(parseMoney_(rawAmt)).setNumberFormat(CFG.CURRENCY_FMT); fixMonto++;
      }
    }
  }
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (bs && bs.getLastRow() > 1) {
    const n = bs.getLastRow() - 1, vals = bs.getRange(2, PB.LOAN_ID, n, 1).getValues();
    for (let i = 0; i < n; i++) { const raw = vals[i][0], t = String(raw == null ? '' : raw).trim(); if (t && t !== raw) { bs.getRange(2 + i, PB.LOAN_ID).setValue(t); fixId++; } }
  }
  return { monto: fixMonto, ids: fixId };
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

// "Solicitudes de Congelamiento": una fila por pedido de congelamiento (Parte 2). El cliente
// la crea desde el formulario web (?page=congelar); el prestamista aprueba/rechaza con casillas.
const FREEZE_REQ_HEADERS = ['Fecha de solicitud', 'ID Préstamo', 'ID Cliente', 'Nombre', 'DNI', 'Correo',
  'Teléfono', 'Capital', 'Saldo con Mora (al pedir)', 'Recargo por Mora (al pedir)',
  'Ya pagado (declarado)', 'Falta para el 50% (24 h)', 'Saldo restante (al pedir)', 'Vencimiento',
  'Fecha de pago prometida', 'Motivo (cliente)', 'Ubicación', 'IP', 'Dispositivo',
  'Aprobar ✔', 'Rechazar ✖', 'Motivo de rechazo', 'Estado', 'Fecha de aprobación',
  'Enlace de firma', 'WhatsApp', 'Estado de Firma', 'Fecha de Firma', 'Contrato de Congelamiento (PDF)'];
const FR_APPROVE_COL = 20, FR_DENY_COL = 21, FR_COVER_COL = 12; // casillas + columna viva "Falta para el 50%"

function setupFreezeRequests_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.FREEZE_REQ);
  const head = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, FREEZE_REQ_HEADERS.length).getValues()[0].map(h => String(h).trim()) : [];
  if (head.join('|') !== FREEZE_REQ_HEADERS.join('|')) {
    sh.getRange(1, 1, 1, FREEZE_REQ_HEADERS.length).setValues([FREEZE_REQ_HEADERS])
      .setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff').setWrap(true);
  }
  sh.setFrozenRows(1);
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, FR_APPROVE_COL, CFG.MAX_ROWS, 1).setDataValidation(cb); // Aprobar ✔
  sh.getRange(2, FR_DENY_COL, CFG.MAX_ROWS, 1).setDataValidation(cb);    // Rechazar ✖
  sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');  // Fecha de solicitud
  sh.getRange(2, 8, CFG.MAX_ROWS, 4).setNumberFormat(CFG.CURRENCY_FMT);    // Capital..Ya pagado (8..11)
  sh.getRange(2, 13, CFG.MAX_ROWS, 1).setNumberFormat(CFG.CURRENCY_FMT);   // Saldo restante (al pedir)
  sh.getRange(2, 14, CFG.MAX_ROWS, 2).setNumberFormat('yyyy-mm-dd');       // Vencimiento + Fecha prometida
  sh.getRange(2, 24, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');       // Fecha de aprobación
  sh.getRange(2, 28, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm'); // Fecha de Firma
  // Columna VIVA "Falta para el 50% (24 h)": fórmula por fila que compara pagos registrados vs el % configurable.
  for (let r = 2; r <= CFG.MAX_ROWS + 1; r++) sh.getRange(r, FR_COVER_COL).setFormula(freezeCoverFormula_(r));
  [150, 100, 90, 170, 110, 200, 130, 110, 140, 150, 130, 160, 130, 110, 140, 220, 200, 120, 230, 90, 90, 220, 110, 130, 220, 160, 120, 150, 180]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  // Franja por Estado (col 23).
  const estL = colL_(23), rng = sh.getRange(2, 1, CFG.MAX_ROWS, FREEZE_REQ_HEADERS.length);
  const byEstado = (val, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + estL + '2="' + val + '"').setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  sh.setConditionalFormatRules([
    byEstado('Pendiente', '#fff2cc', '#7f6000'),
    byEstado('Aprobado', '#d0e0e3', '#0c343d'),
    byEstado('Firmado', '#d9ead3', '#274e13'),
    byEstado('Rechazado', '#f4cccc', '#990000'),
  ]);
  sh.getRange(1, FR_COVER_COL).setNote('VIVO: compara los pagos registrados contra el % configurable del total (Configuración ▸ "Porcentaje a pagar para congelar (%)"). Muestra "✅ cubierto" o "Falta $X". El 50% debe estar PAGADO antes de aprobar y de firmar.');
  sh.getRange(1, FR_APPROVE_COL).setNote('Tildá para APROBAR: SOLO si "Falta para el 50%" muestra ✅ cubierto. Envía por correo y WhatsApp el enlace para firmar. El préstamo se congela recién al firmar.');
  sh.getRange(1, FR_DENY_COL).setNote('Completá antes "Motivo de rechazo" y tildá para RECHAZAR: avisa al cliente por correo y arma el WhatsApp con el motivo.');
  return sh;
}

// "Solicitudes de Plan de Pago": una fila por pedido de plan (Acuerdo Transaccional). El cliente la
// crea desde el formulario web (?page=plan); el prestamista aprueba/rechaza con casillas.
const PLAN_REQ_HEADERS = ['Fecha de solicitud', 'ID Préstamo', 'ID Cliente', 'Nombre', 'DNI', 'Correo',
  'Teléfono', 'Capital', 'Total a Pagar', 'Saldo base (sin mora)', 'Mora suspendida', 'Opción (1/4/6)',
  'Pago inicial', 'Falta para el pago inicial (24 h)', 'Vencimiento', 'Motivo (cliente)', 'Ubicación', 'IP', 'Dispositivo',
  'Aprobar ✔', 'Rechazar ✖', 'Motivo de rechazo', 'Estado', 'Fecha de aprobación',
  'Enlace de firma', 'WhatsApp', 'Estado de Firma', 'Fecha de Firma', 'Acuerdo (PDF)', 'Cadencia'];
const PR_APPROVE_COL = 20, PR_DENY_COL = 21, PR_COVER_COL = 14; // casillas + columna viva "Falta para el 20%"

function setupPlanRequests_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PLAN_REQ);
  if (sh.getMaxColumns() < PLAN_REQ_HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(), PLAN_REQ_HEADERS.length - sh.getMaxColumns()); // ← asegura espacio para «Cadencia»
  const head = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, PLAN_REQ_HEADERS.length).getValues()[0].map(h => String(h).trim()) : [];
  if (head.join('|') !== PLAN_REQ_HEADERS.join('|')) {
    sh.getRange(1, 1, 1, PLAN_REQ_HEADERS.length).setValues([PLAN_REQ_HEADERS])
      .setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff').setWrap(true);
  }
  sh.setFrozenRows(1);
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, PR_APPROVE_COL, CFG.MAX_ROWS, 1).setDataValidation(cb); // Aprobar ✔
  sh.getRange(2, PR_DENY_COL, CFG.MAX_ROWS, 1).setDataValidation(cb);    // Rechazar ✖
  sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');  // Fecha de solicitud
  sh.getRange(2, 8, CFG.MAX_ROWS, 4).setNumberFormat(CFG.CURRENCY_FMT);    // Capital..Mora suspendida (8..11)
  sh.getRange(2, 13, CFG.MAX_ROWS, 1).setNumberFormat(CFG.CURRENCY_FMT);   // Pago inicial (20%)
  sh.getRange(2, 15, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');       // Vencimiento
  sh.getRange(2, 24, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');       // Fecha de aprobación
  sh.getRange(2, 28, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm'); // Fecha de Firma
  for (let r = 2; r <= CFG.MAX_ROWS + 1; r++) sh.getRange(r, PR_COVER_COL).setFormula(planCoverFormula_(r));
  [150, 100, 90, 170, 110, 200, 130, 110, 130, 140, 130, 90, 130, 160, 110, 220, 200, 120, 230, 90, 90, 220, 110, 130, 220, 160, 120, 150, 180, 110]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const estL = colL_(23), rng = sh.getRange(2, 1, CFG.MAX_ROWS, PLAN_REQ_HEADERS.length);
  const byEstado = (val, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + estL + '2="' + val + '"').setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  sh.setConditionalFormatRules([
    byEstado('Pendiente', '#fff2cc', '#7f6000'),
    byEstado('Aprobado', '#d9d2e9', '#351c75'),
    byEstado('Firmado', '#d9ead3', '#274e13'),
    byEstado('Rechazado', '#f4cccc', '#990000'),
  ]);
  sh.getRange(1, PR_COVER_COL).setNote('VIVO: pagos registrados vs el % inicial del total (Configuración ▸ "Pago inicial para plan flexible (%)"). "✅ cubierto" o "Falta $X". El pago inicial debe estar PAGADO antes de aprobar y de firmar.');
  sh.getRange(1, PR_APPROVE_COL).setNote('Tildá para APROBAR: SOLO si "Falta para el pago inicial" muestra ✅ cubierto. Envía el enlace para firmar el Acuerdo. El plan se activa al firmar (reemplaza las cuotas).');
  sh.getRange(1, PR_DENY_COL).setNote('Completá "Motivo de rechazo" y tildá para RECHAZAR: avisa al cliente.');
  return sh;
}

/** Porcentaje (fracción 0–1) que debe estar PAGADO para poder congelar. Configurable; default 50%. */
function freezeDownPct_() {
  let v = Number(String(getSetting_('Porcentaje a pagar para congelar (%)') || '').replace('%', '').replace(',', '.').trim());
  if (!isFinite(v) || v <= 0) v = 50;
  return Math.min(1, v / 100);
}

/**
 * Importes del plan de congelamiento. El requisito se mide sobre los PAGOS REGISTRADOS (hoja "Pagos").
 * - total    = Total a Pagar (capital + interés).
 * - recorded = pagos registrados del préstamo.
 * - balance  = Saldo con Mora a hoy (capital+interés pendiente + mora).
 * - required = porcentaje configurable × total (hito a cubrir ANTES de aprobar/congelar).
 * - downShort= lo que falta pagar para alcanzar el hito = max(0, required − recorded). 0 ⇒ cubierto.
 * - downMet  = el hito está cubierto (downShort ≈ 0).
 * - remaining= saldo restante a abonar por la fecha elegida = max(0, balance − downShort).
 * @param {number=} alreadyPaid  Monto que el cliente DECLARA haber pagado (informativo; no afecta el hito).
 */
function freezeAmounts_(loan, alreadyPaid) {
  const P = Math.max(0, Number(alreadyPaid) || 0);
  const pct = freezeDownPct_();
  const recorded = round2_(loanPayments_(loan.loanId).reduce((s, p) => s + p.amount, 0));
  const total = round2_(loan.totalDue);
  const required = round2_(pct * total);
  const base = Math.max(0, round2_(total - recorded));
  const mora = accruedMora_(loan, new Date());
  const balance = round2_(base + mora);
  const downShort = Math.max(0, round2_(required - recorded));
  const downMet = downShort <= 0.009;
  const remaining = Math.max(0, round2_(balance - downShort));
  return { P: P, pct: pct, pctInt: Math.round(pct * 100), total: total, base: base, mora: mora,
    balance: balance, recorded: recorded, required: required, downShort: downShort, downMet: downMet, remaining: remaining };
}

/** Fórmula VIVA para la columna "Falta para el 50% (24 h)": compara pagos registrados contra el
 *  porcentaje configurable del total; muestra "✅ cubierto" o "Falta $X". Se actualiza sola al registrar pagos. */
function freezeCoverFormula_(r) {
  const B = CFG.SHEETS.BORROWERS, S = CFG.SHEETS.SETTINGS;
  const bs = getSS_().getSheetByName(B), H = bs ? headerIndex_(bs) : {};
  const totN = colByAny_(H, ['Total a Pagar', 'Total a pagar']) || PB.TOTAL;
  const paidN = colByAny_(H, ['Total Pagado']) || PB.PAID;
  const totL = colL_(totN), paidL = colL_(paidN);
  return `=IF($B${r}="","",IFERROR(LET(id,$B${r},` +
    `pct,VLOOKUP("Porcentaje a pagar para congelar (%)",'${S}'!$A:$B,2,FALSE)/100,` +
    `tot,VLOOKUP(id,'${B}'!$A:$${totL},${totN},FALSE),paid,VLOOKUP(id,'${B}'!$A:$${paidL},${paidN},FALSE),need,pct*tot,` +
    `IF(paid>=need-0.009,"✅ cubierto","Falta "&TEXT(need-paid,"$#,##0.00"))),""))`;
}

/* ===================== PLAN DE PAGO (Acuerdo Transaccional) — helpers ===================== */
/** ¿El préstamo llegó al 100% de mora (tope)? Elegibilidad para pedir plan de pago. */
function isLoanAtMoraCap_(loan) {
  if (!loan) return false;
  const capMax = moraCapAmount_(loan);               // tope = % × capital
  if (capMax <= 0) return false;
  return accruedMora_(loan, new Date()) >= capMax - 0.009;
}
/** Porcentaje (fracción 0–1) a pagar para ACTIVAR el plan de pago. Configurable; default 20%. */
function planDownPct_() {
  // Plan flexible (por dificultad): usa "Pago inicial para plan flexible (%)" (0–10%); si está vacío,
  // cae al ajuste anterior "Porcentaje inicial para plan de pago (%)"; default 10%. Permite 0% (sin anticipo).
  let raw = String(getSetting_('Pago inicial para plan flexible (%)') || '').replace('%', '').replace(',', '.').trim();
  if (raw === '') raw = String(getSetting_('Porcentaje inicial para plan de pago (%)') || '').replace('%', '').replace(',', '.').trim();
  let v = Number(raw);
  if (!isFinite(v) || v < 0) v = 10;
  return Math.min(1, v / 100);
}
/**
 * Importes del plan de pago. El plan cubre el TOTAL A PAGAR (capital + interés, SIN mora); la mora
 * se suspende al activar. El requisito inicial se mide sobre los PAGOS REGISTRADOS.
 * - total=Total a Pagar · recorded=pagos registrados · required=pct×total (pago inicial)
 * - base=saldo sin mora pendiente · mora=mora acumulada (se suspende) · downShort/downMet=hito del pago inicial.
 */
function planAmounts_(loan) {
  const pct = planDownPct_();
  const recorded = round2_(loanPayments_(loan.loanId).reduce((s, p) => s + p.amount, 0));
  const total = round2_(loan.totalDue);
  const required = round2_(pct * total);
  const base = Math.max(0, round2_(total - recorded));
  const mora = accruedMora_(loan, new Date());
  const downShort = Math.max(0, round2_(required - recorded));
  const downMet = downShort <= 0.009;
  return { pct: pct, pctInt: Math.round(pct * 100), total: total, recorded: recorded, required: required,
    base: base, mora: mora, downShort: downShort, downMet: downMet };
}
/** Normaliza la cadencia del plan flexible: 'weekly' (semanal), 'biweekly' (quincenal) o 'monthly'
 *  (mensual; valor por defecto y compatible con los planes anteriores). */
function planCadence_(c) { c = String(c || '').toLowerCase().trim(); return (c === 'weekly' || c === 'biweekly') ? c : 'monthly'; }
/** Texto de la opción de plan: "Pago único" o "N cuotas semanales/quincenales/mensuales". */
function planOptTxt_(n, cadence) {
  if (Number(n) <= 1) return 'Pago único';
  const per = planCadence_(cadence) === 'weekly' ? 'semanales' : planCadence_(cadence) === 'biweekly' ? 'quincenales' : 'mensuales';
  return n + ' cuotas ' + per;
}
/** Fechas de vencimiento de las cuotas del plan según la CADENCIA (semanal/quincenal/mensual). Pago
 *  único → ~2 semanas (1 semana si es semanal). Sin cadencia = mensual (comportamiento anterior). */
function planCuotaDates_(signDate, n, cadence) {
  const base = startOfDay_(signDate instanceof Date ? signDate : new Date());
  cadence = planCadence_(cadence);
  if (n <= 1) return [cadence === 'weekly' ? addDays_(base, 7) : addDays_(base, 14)];
  const out = [];
  for (let k = 1; k <= n; k++) {
    if (cadence === 'weekly') out.push(addDays_(base, 7 * k));
    else if (cadence === 'biweekly') out.push(addDays_(base, 14 * k));
    else out.push(addMonths_(base, k));
  }
  return out;
}
/** Fórmula VIVA "Falta para el 20% (24 h)" en "Solicitudes de Plan de Pago": pagos registrados vs % del total. */
function planCoverFormula_(r) {
  const B = CFG.SHEETS.BORROWERS, S = CFG.SHEETS.SETTINGS;
  const bs = getSS_().getSheetByName(B), H = bs ? headerIndex_(bs) : {};
  const totN = colByAny_(H, ['Total a Pagar', 'Total a pagar']) || PB.TOTAL;
  const paidN = colByAny_(H, ['Total Pagado']) || PB.PAID;
  const totL = colL_(totN), paidL = colL_(paidN);
  return `=IF($B${r}="","",IFERROR(LET(id,$B${r},` +
    `pct,IFERROR(VLOOKUP("Pago inicial para plan flexible (%)",'${S}'!$A:$B,2,FALSE),VLOOKUP("Porcentaje inicial para plan de pago (%)",'${S}'!$A:$B,2,FALSE))/100,` +
    `tot,VLOOKUP(id,'${B}'!$A:$${totL},${totN},FALSE),paid,VLOOKUP(id,'${B}'!$A:$${paidL},${paidN},FALSE),need,pct*tot,` +
    `IF(paid>=need-0.009,"✅ cubierto","Falta "&TEXT(need-paid,"$#,##0.00"))),""))`;
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
  // Cuotas: para los vencimientos por período (semana/mes/mes próximo) contamos cuota por cuota.
  const Q = CFG.SHEETS.INSTALLMENTS;          // hoja "Cuotas"
  const qsP = ss.getSheetByName(Q);
  const QH = qsP ? headerIndex_(qsP) : {};
  const qDueL = colL_(colByAny_(QH, ['Fecha de Vencimiento']) || CU.DUE);     // col D
  const qBalL = colL_(colByAny_(QH, ['Saldo Cuota']) || CU.BALANCE);          // col G
  const payIdL = colL_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID);
  const payAmtL = colL_(colByAny_(PH, ['Monto Pagado']) || PP.AMOUNT);
  const paidL = colL_(colByAny_(BH, ['Total Pagado']) || PB.PAID);   // Prestatarios "Total Pagado"
  const capitalPrestado = `SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${capL}:$${capL})`;
  const totalCobrado = pgP ? `SUMIF('${P}'!$${payIdL}:$${payIdL},"L-*",'${P}'!$${payAmtL}:$${payAmtL})` : '0';
  // GANANCIA REAL (base caja, sólo préstamos cuya historia terminó): plata que entró − plata que
  // salió en los PAGADOS y VENCIDOS. Los VENCIDOS NO se pisan en 0: su pérdida (capital que no
  // volvió) baja el neto. Los ACTIVOS (aún corriendo) no entran. El guardia ID≠"" descarta la
  // fila-nota del pie de "Prestatarios" (Estado PAGADO con ID en blanco).
  const paidPagado  = `SUMIFS('${B}'!$${paidL}:$${paidL},'${B}'!$${estL}:$${estL},"${ST.PAID}",'${B}'!$${idL}:$${idL},"<>")`;
  const capPagado   = `SUMIFS('${B}'!$${capL}:$${capL},'${B}'!$${estL}:$${estL},"${ST.PAID}",'${B}'!$${idL}:$${idL},"<>")`;
  const ganPagados   = `(${paidPagado})-(${capPagado})`;    // ganancia limpia de los préstamos ganados
  // "Capital no recuperado": por cada préstamo VENCIDO o EN PLAN, el capital ORIGINAL que aún no
  // volvió = MAX(0, Capital − Total Pagado) por fila, SUMADO. Incluye los préstamos en plan con su
  // capital original a cobrar (sin la mora congelada). Negativo = pérdida (alimenta "Ganancia real").
  const estR = `'${B}'!$${estL}2:$${estL}`, idR = `'${B}'!$${idL}2:$${idL}`;
  const unrecR = `('${B}'!$${capL}2:$${capL}-'${B}'!$${paidL}2:$${paidL})`;
  const perdVencidos = `-SUMPRODUCT(((${estR}="${ST.OVERDUE}")+(${estR}="${ST.PLAN}"))*(${idR}<>"")*(${unrecR}>0)*${unrecR})`;
  const gananciaReal  = `(${ganPagados})+(${perdVencidos})`; // neto de los préstamos cerrados
  sh.getRange('A1').setValue('PANEL DEL PRESTAMISTA').setFontSize(18).setFontWeight('bold').setFontColor('#1c4587');
  const kpis = [
    ['Préstamos activos', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.ACTIVE}",'${B}'!$${idL}:$${idL},"<>")`],   // 3  int
    ['Préstamos vencidos', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.OVERDUE}",'${B}'!$${idL}:$${idL},"<>")`], // 4  int
    ['Préstamos pagados', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.PAID}",'${B}'!$${idL}:$${idL},"<>")`],     // 5  int
    ['Préstamos en plan', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.PLAN}",'${B}'!$${idL}:$${idL},"<>")`],     // 6  int (reestructurados al día)
    ['Solicitudes pendientes', `=COUNTA('${CFG.SHEETS.NEW}'!$B$2:$B)`],                    // 7  int
    ['Contratos sin firmar', `=COUNTIF('${B}'!$${firmaL}:$${firmaL},"${SIGN.PENDING}")`], // 8  int
    ['Fondo total para prestar', `=${fondo}`],                                            // 9  money
    ['Capital prestado', `=${capitalPrestado}`],                                          // 10 money
    ['Total cobrado', `=${totalCobrado}`],                                                // 11 money
    ['Efectivo disponible para prestar', `=${fondo}-${capitalPrestado}+${totalCobrado}`], // 12 money (negativo = sobregiro)
    ['Interés contratado', `=SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${intL}:$${intL})`],      // 13 money
    ['Saldo pendiente total', `=SUMIF('${B}'!$${idL}:$${idL},"L-*",'${B}'!$${salL}:$${salL})`],   // 14 money
    ['Ganancia en préstamos pagados', `=${ganPagados}`],                                 // 15 money (verde)
    ['Pérdida: capital no recuperado (vencidos + en plan)', `=${perdVencidos}`],          // 16 money (rojo)
    ['Ganancia real (neto de préstamos cerrados)', `=${gananciaReal}`],                  // 17 money (rojo si <0)
    ['Ganancia real sobre el fondo (%)', `=IF(${fondo}>0,(${gananciaReal})/(${fondo}),0)`], // 18 %
  ];
  sh.getRange(3, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]])).setFontWeight('bold');
  sh.getRange(3, 2, kpis.length, 1).setFormulas(kpis.map(k => [k[1]]));
  sh.getRange(3, 2, 6, 1).setNumberFormat('0');                 // filas 3–8 enteros
  sh.getRange(9, 2, 6, 1).setNumberFormat(CFG.CURRENCY_FMT);    // filas 9–14 moneda
  sh.getRange(15, 2).setNumberFormat(CFG.CURRENCY_FMT);         // 15 ganancia en pagados $
  sh.getRange(16, 2).setNumberFormat(CFG.CURRENCY_FMT);         // 16 pérdida en vencidos $
  sh.getRange(17, 2).setNumberFormat(CFG.CURRENCY_FMT);         // 17 ganancia real (neto) $
  sh.getRange(18, 2).setNumberFormat('0.00%');                  // 18 ganancia real sobre fondo %
  // Fila 6: "Préstamos en plan" (reestructurados al día) — categoría aparte, en violeta.
  sh.getRange('A6:B6').setFontColor('#351c75'); sh.getRange('B6').setFontWeight('bold');
  sh.getRange('A6').setNote('Préstamos con plan de pago (reestructurados) al día. Categoría aparte: NO se cuentan en "activos" ni "vencidos".');
  sh.getRange('A8').setFontColor('#990000'); sh.getRange('B8').setFontColor('#990000').setFontWeight('bold');
  sh.getRange('A8').setNote('Préstamos aprobados cuyo contrato aún no fue firmado por el prestatario. No desembolsar hasta la firma.');
  sh.getRange('A12').setFontColor('#38761d'); sh.getRange('B12').setFontColor('#38761d').setFontWeight('bold');
  sh.getRange('A12').setNote('Efectivo disponible = Fondo total − Capital prestado + Total cobrado. Baja al prestar y sube al cobrar. Un valor NEGATIVO (en rojo) indica sobregiro: se prestó más capital del disponible.');
  // Ganancia real (base caja, préstamos cerrados): ganados en verde, pérdida en rojo, neto y % en
  // verde salvo que sean negativos (regla condicional los pone en rojo).
  sh.getRange('A15:B15').setFontColor('#38761d'); sh.getRange('B15').setFontWeight('bold');
  sh.getRange('A16:B16').setFontColor('#cc0000'); sh.getRange('B16').setFontWeight('bold');
  sh.getRange('A17:B18').setFontColor('#38761d'); sh.getRange('B17:B18').setFontWeight('bold');
  sh.getRange('A15').setNote('De los préstamos que YA terminaron y se pagaron: lo que te devolvieron − lo que prestaste. La ganancia limpia de los préstamos ganados.');
  sh.getRange('A16').setNote('Capital ORIGINAL que aún no volvió, sumando los préstamos VENCIDOS y EN PLAN: por cada préstamo, MAX(0, Capital − Total Pagado). Incluye a los que están en plan de pago con su capital original a cobrar (sin la mora congelada). Da negativo: es la PÉRDIDA / capital a recuperar.');
  sh.getRange('A17').setNote('La verdad del negocio hasta hoy: ganancia de los pagados menos pérdida de los vencidos. Sólo cuenta préstamos cuya historia terminó; los que siguen vivos no entran todavía. Es plata que entró − plata que salió en préstamos PAGADOS y VENCIDOS.');
  sh.getRange('A18').setNote('Ese neto dividido por tu fondo total. De cada $100 que pusiste, cuánto ganaste o perdiste de verdad hasta hoy.');
  // Rojo automático para valores negativos: sobregiro (B12) y ganancia real neta/% (B17/B18).
  sh.setConditionalFormatRules([redIfNegativeRule_(sh.getRange('B12')), redIfNegativeRule_(sh.getRange('B17')), redIfNegativeRule_(sh.getRange('B18'))]);
  // Fila 19: Retiro disponible (instantánea; el cálculo usa ventanas de fechas que una fórmula de celda no expresa bien).
  sh.getRange('A19').setValue('Retiro disponible (sin frenar el crecimiento)').setFontWeight('bold').setFontColor('#38761d');
  try {
    const wd = withdrawableStats_();
    sh.getRange('B19').setValue(wd.withdrawable).setNumberFormat(CFG.CURRENCY_FMT).setFontColor('#38761d').setFontWeight('bold');
    sh.getRange('A19').setNote('Máximo retirable ahora sin frenar el ritmo de colocación. Solo libera ganancia realizada (interés cobrado) y ' +
      'retiene una reserva = colocación proyectada en ' + wd.horizonDays + ' días − cobros esperados + colchón por morosidad. ' +
      'Instantánea: se recalcula con "⑤ Actualizar" y con "💵 Retiro disponible".');
  } catch (e) { logError_('setupPanel_:withdrawable', e); }
  // Ventanas de fechas (fórmulas). Semana = lunes→domingo (WEEKDAY tipo 2: lun=1…dom=7);
  // mes y mes próximo con EOMONTH. "Semana pasada" = la semana lun–dom anterior a la actual.
  const wkStart = 'TODAY()-WEEKDAY(TODAY(),2)+1';
  const wkEnd   = 'TODAY()-WEEKDAY(TODAY(),2)+7';
  const moStart = 'EOMONTH(TODAY(),-1)+1';
  const moEnd   = 'EOMONTH(TODAY(),0)';
  const nxStart = 'EOMONTH(TODAY(),0)+1';
  const nxEnd   = 'EOMONTH(TODAY(),1)';
  const todStart = 'TODAY()', todNext = 'TODAY()+1';                      // hoy: [hoy, mañana)
  const pwStart = 'TODAY()-WEEKDAY(TODAY(),2)-6';                         // lunes de la semana pasada
  const pwNext  = 'TODAY()-WEEKDAY(TODAY(),2)+1';                         // = lunes de esta semana (exclusivo)
  // Ventanas "siguiente exclusivo" para los cobros por período (cntPay/amtPay usan [inicio, siguiente)).
  const ydStart = 'TODAY()-1', ydNext = 'TODAY()';                        // ayer: [ayer, hoy)
  const wkNext  = 'TODAY()-WEEKDAY(TODAY(),2)+8';                         // lunes de la semana próxima (exclusivo) → esta semana = lun–dom
  const moNext  = 'EOMONTH(TODAY(),0)+1';                                 // 1.º del mes próximo (exclusivo) → este mes completo
  const pmStart = 'EOMONTH(TODAY(),-2)+1', pmNext = 'EOMONTH(TODAY(),-1)+1'; // mes anterior completo

  // ----- 1) Vencimientos por período (TODAS las cuotas no pagadas, por fecha de vencimiento) -----
  const cntDue = (s, e) => qsP ? `=COUNTIFS('${Q}'!$${qDueL}$2:$${qDueL},">="&(${s}),'${Q}'!$${qDueL}$2:$${qDueL},"<="&(${e}),'${Q}'!$${qBalL}$2:$${qBalL},">0.009")` : '=0';
  const amtDue = (s, e) => qsP ? `=SUMIFS('${Q}'!$${qBalL}$2:$${qBalL},'${Q}'!$${qDueL}$2:$${qDueL},">="&(${s}),'${Q}'!$${qDueL}$2:$${qDueL},"<="&(${e}),'${Q}'!$${qBalL}$2:$${qBalL},">0.009")` : '=0';
  sh.getRange('A21').setValue('Vencimientos por período (todas las cuotas)').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A21').setNote('Cuenta las CUOTAS no pagadas (saldo > 0) según su fecha de vencimiento. La semana va de lunes a domingo. El monto es el saldo que queda por cobrar.');
  sh.getRange(22, 1, 1, 3).setValues([['Período', 'Cuotas a cobrar', 'Monto a cobrar']]).setFontWeight('bold');
  const periodRows = [
    ['Esta semana (lun–dom)', cntDue(wkStart, wkEnd), amtDue(wkStart, wkEnd)],
    ['Este mes', cntDue(moStart, moEnd), amtDue(moStart, moEnd)],
    ['Mes próximo', cntDue(nxStart, nxEnd), amtDue(nxStart, nxEnd)],
  ];
  sh.getRange(23, 1, periodRows.length, 1).setValues(periodRows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(23, 2, periodRows.length, 1).setFormulas(periodRows.map(r => [r[1]]));
  sh.getRange(23, 3, periodRows.length, 1).setFormulas(periodRows.map(r => [r[2]]));
  sh.getRange(23, 2, periodRows.length, 1).setNumberFormat('0');               // conteo de cuotas
  sh.getRange(23, 3, periodRows.length, 1).setNumberFormat(CFG.CURRENCY_FMT);  // monto por cobrar

  // ----- 2) Vencimientos por período — préstamos SIN plan de pago (Estado ≠ "EN PLAN") -----
  // Igual que arriba pero excluyendo las cuotas cuyo préstamo está reestructurado ("EN PLAN").
  // Como "Cuotas" no guarda el Estado, se resuelve con VLOOKUP a "Prestatarios" (ID→Estado).
  const bIdCol = colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const bEstCol = colByAny_(BH, ['Estado']) || PB.STATE;
  const vlIdx = bEstCol - bIdCol + 1;                 // índice del Estado dentro del rango ID→Estado
  const qLoanL = colL_(colByAny_(QH, ['ID Préstamo', 'ID Prestamo']) || CU.LOAN_ID);
  const qLast = qsP ? qsP.getLastRow() : 1;
  const noPlanOk = qsP && qLast >= 2 && vlIdx >= 1;
  const dueArr = `'${Q}'!$${qDueL}$2:$${qDueL}${qLast}`;
  const balArr = `'${Q}'!$${qBalL}$2:$${qBalL}${qLast}`;
  const estArr = `ARRAYFORMULA(IFERROR(VLOOKUP('${Q}'!$${qLoanL}$2:$${qLoanL}${qLast},'${B}'!$${idL}:$${estL},${vlIdx},FALSE),""))`;
  const inWin = (s, e) => `(${dueArr}>=(${s}))*(${dueArr}<=(${e}))*(${balArr}>0.009)*(${estArr}<>"${ST.PLAN}")`;
  const cntNoPlan = (s, e) => noPlanOk ? `=SUMPRODUCT(${inWin(s, e)})` : '=0';
  const amtNoPlan = (s, e) => noPlanOk ? `=SUMPRODUCT((${inWin(s, e)})*(${balArr}))` : '=0';
  sh.getRange('A27').setValue('Vencimientos por período — préstamos SIN plan de pago').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A27').setNote('Igual que el bloque anterior pero SOLO préstamos que NO están "EN PLAN" (sin reestructurar). Excluye las cuotas de préstamos reestructurados.');
  sh.getRange(28, 1, 1, 3).setValues([['Período', 'Cuotas a cobrar', 'Monto a cobrar']]).setFontWeight('bold');
  const noPlanRows = [
    ['Esta semana (lun–dom)', cntNoPlan(wkStart, wkEnd), amtNoPlan(wkStart, wkEnd)],
    ['Este mes', cntNoPlan(moStart, moEnd), amtNoPlan(moStart, moEnd)],
    ['Mes próximo', cntNoPlan(nxStart, nxEnd), amtNoPlan(nxStart, nxEnd)],
  ];
  sh.getRange(29, 1, noPlanRows.length, 1).setValues(noPlanRows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(29, 2, noPlanRows.length, 1).setFormulas(noPlanRows.map(r => [r[1]]));
  sh.getRange(29, 3, noPlanRows.length, 1).setFormulas(noPlanRows.map(r => [r[2]]));
  sh.getRange(29, 2, noPlanRows.length, 1).setNumberFormat('0');
  sh.getRange(29, 3, noPlanRows.length, 1).setNumberFormat(CFG.CURRENCY_FMT);

  // ----- 3) Cobros por período (pagos registrados en "Pagos" por Fecha de Pago; día/semana/mes, actual y anterior) -----
  const payDateL = colL_(colByAny_(PH, ['Fecha de Pago']) || PP.DATE);
  const payDateR = `'${P}'!$${payDateL}$2:$${payDateL}`;
  const payAmtR = `'${P}'!$${payAmtL}$2:$${payAmtL}`;
  const cntPay = (s, nx) => pgP ? `=COUNTIFS(${payDateR},">="&(${s}),${payDateR},"<"&(${nx}))` : '=0';
  const amtPay = (s, nx) => pgP ? `=SUMIFS(${payAmtR},${payDateR},">="&(${s}),${payDateR},"<"&(${nx}))` : '=0';
  sh.getRange('A33').setValue('Cobros por período').setFontWeight('bold').setFontColor('#38761d');
  sh.getRange('A33').setNote('Pagos registrados en la hoja "Pagos" según su Fecha de Pago, por día/semana/mes (período actual y anterior). "Pagos cobrados" = cantidad de pagos; "Monto cobrado" = suma de "Monto Pagado". La semana va de lunes a domingo.');
  sh.getRange(34, 1, 1, 3).setValues([['Período', 'Pagos cobrados', 'Monto cobrado']]).setFontWeight('bold');
  const payRows = [
    ['Hoy',                     cntPay(todStart, todNext), amtPay(todStart, todNext)],
    ['Ayer',                    cntPay(ydStart,  ydNext),  amtPay(ydStart,  ydNext)],
    ['Esta semana (lun–dom)',   cntPay(wkStart,  wkNext),  amtPay(wkStart,  wkNext)],
    ['Semana pasada (lun–dom)', cntPay(pwStart,  pwNext),  amtPay(pwStart,  pwNext)],
    ['Este mes',                cntPay(moStart,  moNext),  amtPay(moStart,  moNext)],
    ['Mes anterior',            cntPay(pmStart,  pmNext),  amtPay(pmStart,  pmNext)],
  ];
  sh.getRange(35, 1, payRows.length, 1).setValues(payRows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(35, 2, payRows.length, 1).setFormulas(payRows.map(r => [r[1]]));
  sh.getRange(35, 3, payRows.length, 1).setFormulas(payRows.map(r => [r[2]]));
  sh.getRange(35, 2, payRows.length, 1).setNumberFormat('0');
  sh.getRange(35, 3, payRows.length, 1).setNumberFormat(CFG.CURRENCY_FMT);

  sh.getRange('A42').setValue('Próximos vencimientos (7 días):').setFontWeight('bold');
  // Nombre resuelto desde "Clientes" vía ID Cliente (col B). Vencimiento=G, Saldo=K, Estado=L.
  const nameArr = `ARRAYFORMULA(IFERROR(VLOOKUP('${B}'!$${cliL}2:$${cliL},'${C}'!$A:$B,2,FALSE),""))`;
  sh.getRange('A43').setFormula(
    `=IFERROR(SORT(FILTER({'${B}'!$${idL}2:$${idL},${nameArr},'${B}'!$${venL}2:$${venL},'${B}'!$${salL}2:$${salL}},` +
    `('${B}'!$${estL}2:$${estL}<>"${ST.PAID}")*('${B}'!$${venL}2:$${venL}>=TODAY())*('${B}'!$${venL}2:$${venL}<=TODAY()+7)),3,TRUE),"— sin vencimientos próximos —")`);
  sh.getRange('A42').setNote('Muestra préstamos no pagados que vencen dentro de 7 días.');
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
    ['Préstamos activos', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.ACTIVE}",'${B}'!$${idL}:$${idL},"<>")`, 'i'],
    ['Préstamos vencidos', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.OVERDUE}",'${B}'!$${idL}:$${idL},"<>")`, 'i'],
    ['Préstamos pagados', `=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.PAID}",'${B}'!$${idL}:$${idL},"<>")`, 'i'],
    ['Tasa de morosidad (préstamos vencidos)', `=IFERROR(COUNTIFS('${B}'!$${estL}:$${estL},"${ST.OVERDUE}",'${B}'!$${idL}:$${idL},"<>")/(COUNTIFS('${B}'!$${estL}:$${estL},"${ST.ACTIVE}",'${B}'!$${idL}:$${idL},"<>")+COUNTIFS('${B}'!$${estL}:$${estL},"${ST.OVERDUE}",'${B}'!$${idL}:$${idL},"<>")),0)`, 'p'],
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
  sh.getRange('E5').setFormula(`=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.ACTIVE}",'${B}'!$${idL}:$${idL},"<>")`);
  sh.getRange('E6').setFormula(`=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.OVERDUE}",'${B}'!$${idL}:$${idL},"<>")`);
  sh.getRange('E7').setFormula(`=COUNTIFS('${B}'!$${estL}:$${estL},"${ST.PAID}",'${B}'!$${idL}:$${idL},"<>")`);

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

  // 4) Préstamos vs. cobros por mes del AÑO ACTUAL (barras) + ganancia neta mensual.
  // "Prestado" = capital colocado por Fecha del Préstamo; "Cobrado" = pagos por Fecha de Pago;
  // "Ganancia" = Cobrado − Prestado (flujo neto de caja del mes). Solo préstamos "L-…".
  const fecL = colL_(colByAny_(BH, ['Fecha del Préstamo', 'Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo']) || PB.LOAN_DATE);
  const payDateL = colL_(colByAny_(PH, ['Fecha de Pago']) || PP.DATE);
  const meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const mIni = m => `DATE(YEAR(TODAY()),${m},1)`;              // 1.º del mes m del año en curso
  const mFin = m => `DATE(YEAR(TODAY()),${m + 1},1)`;          // 1.º del mes siguiente (exclusivo); m=12 → enero del próximo año
  const lentMo = m => `=SUMIFS('${B}'!$${capL}:$${capL},'${B}'!$${fecL}:$${fecL},">="&${mIni(m)},'${B}'!$${fecL}:$${fecL},"<"&${mFin(m)},'${B}'!$${idL}:$${idL},"L-*")`;
  const paidMo = m => pgS ? `=SUMIFS('${P}'!$${payAmtL}:$${payAmtL},'${P}'!$${payDateL}:$${payDateL},">="&${mIni(m)},'${P}'!$${payDateL}:$${payDateL},"<"&${mFin(m)},'${P}'!$${payIdL}:$${payIdL},"L-*")` : '=0';
  sh.getRange('A23').setValue('Préstamos vs. cobros por mes (año actual)').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A23').setNote('Por mes del año en curso. "Prestado" = capital colocado según Fecha del Préstamo; "Cobrado" = pagos recibidos según Fecha de Pago; "Ganancia" = Cobrado − Prestado (flujo neto de caja del mes). Solo préstamos "L-…".');
  sh.getRange('A24:D24').setValues([['Mes', 'Prestado', 'Cobrado', 'Ganancia']]).setFontWeight('bold');
  const mesLbl = meses.map(x => [x]), lentF = [], paidF = [], ganF = [];
  for (let m = 1; m <= 12; m++) {
    const rr = 24 + m;                                         // filas 25..36
    lentF.push([lentMo(m)]); paidF.push([paidMo(m)]); ganF.push([`=C${rr}-B${rr}`]);
  }
  sh.getRange(25, 1, 12, 1).setValues(mesLbl).setFontWeight('bold');     // A: Mes
  sh.getRange(25, 2, 12, 1).setFormulas(lentF);                          // B: Prestado
  sh.getRange(25, 3, 12, 1).setFormulas(paidF);                          // C: Cobrado
  sh.getRange(25, 4, 12, 1).setFormulas(ganF);                           // D: Ganancia
  sh.getRange(25, 2, 12, 3).setNumberFormat(money);
  // Ganancia mensual negativa en rojo (sin pisar la regla del efectivo disponible fijada arriba).
  const rulesStats = sh.getConditionalFormatRules();
  rulesStats.push(redIfNegativeRule_(sh.getRange(25, 4, 12, 1)));
  sh.setConditionalFormatRules(rulesStats);

  sh.setColumnWidth(1, 280); sh.setColumnWidth(2, 150); sh.setColumnWidth(3, 130);
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

  // Barras: prestado vs. cobrado por mes del año actual (la diferencia = ganancia neta mensual).
  const colMensual = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('A24:D36'))
    .setPosition(49, 1, 0, 0)
    .setOption('title', 'Prestado vs. cobrado por mes (año actual) — ganancia = cobrado − prestado')
    .setOption('width', 900).setOption('height', 340)
    .setOption('legend', 'top')
    .setOption('colors', ['#e06666', '#6aa84f', '#1c4587'])
    .build();
  sh.insertChart(colMensual);

  return sh;
}

const LATE_HEADERS = ['ID Préstamo', 'Prestatario', 'DNI', 'Correo', 'Teléfono', 'Fecha de Vencimiento',
  'Días de Atraso', 'Total a Pagar', 'Recargo por Mora (acum.)', 'Total Pagado', 'Saldo Pendiente',
  'Enviar Aviso ✉', 'Último Aviso Enviado', 'Tope de mora', 'Plan de pago 📅 (3 cuotas)', 'Plan enviado',
  'PDF de cuotas 🧾', 'PDF generado', 'Pedir congelamiento 📲', 'Pedir plan de pago 📋',
  'Cobranza ✉', 'Cobranza 📲', 'Plan de pago ✉'];
// Columnas (1-based) de control en "Pagos Atrasados".
const LATE_SEND_COL = 12, LATE_SENT_COL = 13, LATE_CAP_COL = 14, LATE_PLAN_COL = 15, LATE_PLAN_SENT_COL = 16,
  LATE_DOC_COL = 17, LATE_DOC_DONE_COL = 18, LATE_FREEZE_COL = 19, LATE_PLANREQ_COL = 20,
  LATE_COB_MAIL_COL = 21, LATE_COB_WA_COL = 22, LATE_PLAN_MAIL_COL = 23;
function setupLate_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.LATE); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange(1, 1, 1, LATE_HEADERS.length).setValues([LATE_HEADERS])
    .setFontWeight('bold').setBackground('#990000').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  [110, 170, 120, 200, 120, 130, 100, 120, 160, 120, 130, 110, 150, 100, 120, 110, 110, 110, 120, 120, 120, 120, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, LATE_COB_MAIL_COL).setNote('Enlace de CORREO (mailto) de la División de Cobranzas, listo para enviar al prestatario: resumen del préstamo, pagos, mora, y 2 opciones (cancelar sin mora el próximo día hábil = saldo base, o congelar la mora hoy y pagar el total con mora). "sin correo" si el cliente no tiene email.');
  sh.getRange(1, LATE_COB_WA_COL).setNote('Enlace de WhatsApp de la División de Cobranzas, con el mismo mensaje que el correo. "sin teléfono" si el cliente no tiene un número válido.');
  sh.getRange(1, LATE_PLAN_MAIL_COL).setNote('Enlace que ABRE GMAIL (no envía) con un correo genérico para proponer un PLAN DE PAGO EN CUOTAS: explica cómo funciona, los pasos y el enlace al formulario del plan. Revisalo y envialo desde Gmail. "sin correo" si el cliente no tiene email.');
  sh.getRange('A1').setNote('Se actualiza con "⑤ Actualizar" y automáticamente cada día. Lista los préstamos vencidos con saldo.');
  sh.getRange('L1').setNote('Tilde la casilla para enviar un aviso de mora por correo (en español). La casilla se destilda sola y la fecha de envío aparece en "Último Aviso Enviado".');
  sh.getRange(1, LATE_CAP_COL).setNote('Mora acumulada como % de su tope ("Tope de mora (% del total a devolver)" en Configuración). SÍ (100%) = llegó al tope: candidato a plan de pago en 3 cuotas.');
  sh.getRange(1, LATE_PLAN_COL).setNote('Tilde para reestructurar el préstamo en 3 cuotas mensuales iguales (la deuda con mora se congela a hoy) y enviar el plan por correo al prestatario. La casilla se destilda sola.');
  sh.getRange(1, LATE_DOC_COL).setNote('Tilde para generar el PDF con el cronograma de cuotas del préstamo (montos, vencimientos, pagado y saldo). Se guarda en la carpeta del prestatario; la fecha queda en "PDF generado". La casilla se destilda sola.');
  return sh;
}

/* ===================== HOJA "EN PLAN" (préstamos reestructurados) ===================== */
const PLAN_HEADERS = ['ID Préstamo', 'Prestatario', 'DNI', 'Correo', 'Teléfono', 'Vencimiento del plan',
  'Cuotas', 'Total del plan', 'Total Pagado', 'Saldo del plan',
  'Plan de pago 📅 (3 cuotas)', 'Plan enviado', 'PDF de cuotas 🧾', 'PDF generado', 'Enlace PDF'];
// Columnas (1-based) de control en "En Plan".
const PLAN_PLAN_COL = 11, PLAN_PLAN_DONE_COL = 12, PLAN_DOC_COL = 13, PLAN_DOC_DONE_COL = 14, PLAN_LINK_COL = 15;

/** Instala/recrea la hoja "En Plan" (lista de préstamos reestructurados, EN PLAN). */
function setupEnPlan_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PLAN); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange(1, 1, 1, PLAN_HEADERS.length).setValues([PLAN_HEADERS])
    .setFontWeight('bold').setBackground('#351c75').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  [110, 170, 120, 200, 120, 140, 70, 140, 130, 140, 150, 110, 120, 110, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('Préstamos EN PLAN (con plan de pago / reestructurados). Se actualiza con "⑤ Actualizar" y con el menú 📋. Tildá "Plan de pago 📅" para regenerar el plan en 3 cuotas (desde la semana próxima, SIN correo) o "PDF de cuotas 🧾" para generar el PDF (su enlace queda en "Enlace PDF").');
  sh.getRange(1, PLAN_PLAN_COL).setNote('Tilde para REGENERAR el plan en 3 cuotas (deuda congelada a hoy; 1.ª cuota la semana próxima). NO envía correo. La casilla se destilda sola.');
  sh.getRange(1, PLAN_DOC_COL).setNote('Tilde para generar el PDF del cronograma de cuotas; el enlace clicable queda en "Enlace PDF". La casilla se destilda sola.');
  return sh;
}

/** Reconstruye "En Plan": un renglón por préstamo con Estado = EN PLAN. Conserva las fechas de
 *  "Plan enviado"/"PDF generado" y el "Enlace PDF" ya generados (por ID de préstamo). */
function rebuildEnPlanSheet_(ss) {
  ss = ss || getSS_();
  const sh = ss.getSheetByName(CFG.SHEETS.PLAN) || setupEnPlan_(ss);
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS); if (!bs) return;
  sh.getRange(1, 1, 1, PLAN_HEADERS.length).setValues([PLAN_HEADERS])
    .setFontWeight('bold').setBackground('#351c75').setFontColor('#fff').setWrap(true);
  // Conservar datos generados previamente, por ID de préstamo.
  const prev = {}, last0 = sh.getLastRow();
  if (last0 >= 2) {
    const vals = sh.getRange(2, 1, last0 - 1, PLAN_HEADERS.length).getValues();
    const linkF = sh.getRange(2, PLAN_LINK_COL, last0 - 1, 1).getFormulas();
    vals.forEach((r, i) => { const id = String(r[0]).trim(); if (id) prev[id] = { planSent: r[PLAN_PLAN_DONE_COL - 1], docDone: r[PLAN_DOC_DONE_COL - 1], link: linkF[i][0] || r[PLAN_LINK_COL - 1] }; });
  }
  const clearH = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, clearH, PLAN_HEADERS.length).clearContent().setBackground(null).setFontWeight(null);
  sh.getRange(2, PLAN_PLAN_COL, clearH, 1).clearDataValidations();
  sh.getRange(2, PLAN_DOC_COL, clearH, 1).clearDataValidations();
  const H = headerIndex_(bs), col = n => colByAny_(H, n);
  const idC = col(['ID Préstamo', 'ID Prestamo']), estC = col(['Estado']), venC = col(['Vencimiento', 'Fecha de Vencimiento']),
    pagC = col(['Total Pagado']), bmC = col(['Saldo con Mora']), planTotC = col(['Suma de cuotas', 'Total a Pagar (con mora)']);
  const last = bs.getLastRow(); if (last < 2 || !idC || !estC) return;
  const estVals = bs.getRange(2, estC, last - 1, 1).getValues();
  const rows = [], links = [];
  for (let i = 0; i < estVals.length; i++) {
    if (String(estVals[i][0]).trim() !== ST.PLAN) continue;
    const row = i + 2, loan = readLoan_(bs, row), id = loan.loanId;
    const venc = venC ? bs.getRange(row, venC).getValue() : '';
    const pagado = pagC ? Number(bs.getRange(row, pagC).getValue()) || 0 : 0;
    const totalPlan = planTotC ? Number(bs.getRange(row, planTotC).getValue()) || 0 : 0;
    const saldoPlan = bmC ? Number(bs.getRange(row, bmC).getValue()) || 0 : 0;
    const nCuotas = cuotasDeLoan_(id).length;
    const p = prev[id] || {};
    rows.push([id, loan.name, loan.dni, loan.email, loan.phone, (venc instanceof Date) ? venc : '', nCuotas,
      totalPlan, pagado, saldoPlan, false, (p.planSent instanceof Date) ? p.planSent : '', false, (p.docDone instanceof Date) ? p.docDone : '', '']);
    links.push(p.link || '');
  }
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, PLAN_HEADERS.length).setValues(rows);
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd');     // Vencimiento del plan
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('0');              // Cuotas
  sh.getRange(2, 8, rows.length, 3).setNumberFormat(CFG.CURRENCY_FMT); // Total/Pagado/Saldo del plan
  sh.getRange(2, PLAN_PLAN_COL, rows.length, 1).insertCheckboxes();
  sh.getRange(2, PLAN_PLAN_DONE_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, PLAN_DOC_COL, rows.length, 1).insertCheckboxes();
  sh.getRange(2, PLAN_DOC_DONE_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  links.forEach((lnk, i) => { if (lnk) { const cell = sh.getRange(i + 2, PLAN_LINK_COL); if (/^=/.test(String(lnk))) cell.setFormula(lnk); else cell.setValue(lnk); } });
}

/** Reconstruye "Pagos Atrasados": un renglón por préstamo vencido con saldo, con recargo acumulado. */
function rebuildLateSheet_(ss) {
  ss = ss || getSS_();
  const sh = ss.getSheetByName(CFG.SHEETS.LATE) || setupLate_(ss) || ss.getSheetByName(CFG.SHEETS.LATE);
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  // Reafirmar encabezados: migra hojas existentes (13 columnas) al layout con
  // "Tope de mora" / "Plan de pago 📅" sin perder datos. Idempotente.
  sh.getRange(1, 1, 1, LATE_HEADERS.length).setValues([LATE_HEADERS])
    .setFontWeight('bold').setBackground('#990000').setFontColor('#fff').setWrap(true);
  // limpia datos y casillas previas (conserva encabezados)
  const clearH = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, clearH, LATE_HEADERS.length).clearContent().setBackground(null).setFontWeight(null);
  sh.getRange(2, LATE_SEND_COL, clearH, 1).clearDataValidations();
  sh.getRange(2, LATE_PLAN_COL, clearH, 1).clearDataValidations();
  sh.getRange(2, LATE_DOC_COL, clearH, 1).clearDataValidations();
  if (last < 2) return;
  const today = new Date(), capFrac = moraCapFrac_(), rows = [];
  // Cobranza: datos fijos para toda la corrida (fecha límite = próximo día hábil, app web, medios de
  // pago, firmantes y recargo diario). Se arman una vez y se reusan en cada fila.
  const cobDeadline = rollToBusinessDay_(addDays_(today, 1));
  const cobDeadlineStr = DIAS_SEMANA_ES[cobDeadline.getDay()] + ' ' + fmtDate_(cobDeadline);
  const cobWebBase = webAppBaseUrl_();
  const feePctFrac = lateFeeRate_();
  const cobAlias = String(getSetting_('Mercado Pago — Alias/CVU') || getSetting_('Cuenta de cobro/desembolso — Alias') || '').trim();
  const cobContact = String(getSetting_('Teléfono del Prestamista') || getSetting_('Correo del Prestamista') || '').trim();
  const cobSig1 = String(getSetting_('Firmante 1 (nombre completo)') || companyName_()).trim();
  const cobSig2 = String(getSetting_('Firmante 2 (nombre completo)') || '').trim();
  const cobPlanPctInt = (typeof planDownPct_ === 'function') ? Math.round(planDownPct_() * 100) : 10; // % del pago inicial del plan
  // --- LECTURA EN BLOQUE (una sola vez cada hoja) para reconstruir rápido ---
  const H = headerIndex_(bs), col = (n, fb) => colByAny_(H, n) || fb;
  const idC = col(['ID Préstamo', 'ID Prestamo'], PB.LOAN_ID), cliC = col(['ID Cliente'], PB.CLIENT_ID),
    capC = col(['Capital'], PB.PRINCIPAL), plzC = col(['Plazo (días)', 'Plazo (dias)', 'Plazo (meses)', 'Plazo'], PB.TERM),
    tasC = col(['Tasa'], PB.RATE), fecC = col(['Fecha del Préstamo', 'Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo'], PB.LOAN_DATE),
    venC = col(['Fecha de Vencimiento', 'Vencimiento'], PB.DUE), intC = col(['Interés', 'Interes'], PB.INTEREST),
    totC = col(['Total a Pagar', 'Total a pagar'], PB.TOTAL), notC = col(['Último Aviso'], PB.NOTICE),
    adjC = colByAny_(H, ['Mora (ajuste)']), sinceC = colByAny_(H, ['Congelado el']), daysAccC = colByAny_(H, ['Días congelados', 'Dias congelados']),
    planSinceC = colByAny_(H, ['Plan desde']);
  const maxCol = Math.max(idC, cliC, capC, plzC, tasC, fecC, venC, intC, totC, notC, adjC || 0, sinceC || 0, daysAccC || 0, planSinceC || 0);
  const bdata = bs.getRange(2, 1, last - 1, maxCol).getValues();
  const payMap = loanPaymentsAll_();   // todos los pagos en UNA lectura
  for (let i = 0; i < bdata.length; i++) {
    const v = bdata[i], g = c => v[c - 1];
    const loanId = String(g(idC) || '').trim(); if (!loanId) continue;
    const term = Number(g(plzC)) || 0;
    const ld = g(fecC) instanceof Date ? g(fecC) : new Date(g(fecC));
    const dueDate = g(venC) instanceof Date ? g(venC) : dueDateForTerm_(term || 1, ld);
    if (!(dueDate instanceof Date) || isNaN(dueDate.getTime())) continue;
    if (daysLate_(dueDate, today) <= 0) continue;  // seguimos listando los vencidos aunque estén congelados
    if (planSinceC && g(planSinceC) instanceof Date) continue;  // EN PLAN (Acuerdo): va en "En Plan", no en Pagos Atrasados
    const principal = Number(g(capC)) || 0;
    let rate = Number(g(tasC)) || 0; if (!rate) rate = loanRateForTerm_(term);
    let interest = Number(g(intC)) || 0; if (!interest) interest = round2_(principal * rate);
    let totalDue = Number(g(totC)) || 0; if (!totalDue) totalDue = round2_(principal + interest);
    const cli = getClienteById_(String(g(cliC) || '').trim()) || {};  // Clientes está cacheado
    const loan = { loanId: loanId, clientId: String(g(cliC) || '').trim(), name: cli.name || '', email: cli.email || '',
      phone: cli.phone || '', dni: cli.dni || '', principal: principal, term: term, rate: rate, loanDate: ld, dueDate: dueDate, interest: interest, totalDue: totalDue };
    const pays = payMap[loanId] || [];
    const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
    const base = Math.max(0, round2_(totalDue - totalPaid));            // capital + interés pendiente
    // Ajuste manual de mora (columna "Mora (ajuste)") e info de congelamiento, LEÍDOS DEL BLOQUE.
    const adjV = adjC ? g(adjC) : '';
    const override = (typeof adjV === 'number' && isFinite(adjV)) ? Math.max(0, round2_(adjV)) : null;
    const sinceV = sinceC ? g(sinceC) : '';
    const frozen = { since: (sinceV instanceof Date) ? sinceV : null, daysAcc: daysAccC ? (Number(g(daysAccC)) || 0) : 0 };
    const feeAccum = accruedMoraFast_(loan, today, pays, override, frozen);   // mora SIN leer hojas
    const out = round2_(base + feeAccum);                                // monto a pagar hoy = saldo + mora
    if (out <= 0) continue; // vencido pero saldado (y sin mora / mora condonada)
    const days = Math.max(0, daysLate_(dueDate, effMoraAsOf2_(frozen, today)));  // atraso congelado-aware
    const lastNotice = notC ? g(notC) : '';
    const isFrozen = frozen.since instanceof Date;
    const nameShown = isFrozen ? '❄ ' + loan.name : loan.name; // marca de congelado
    // "Tope de mora": % del recargo acumulado respecto de su tope (= % × CAPITAL).
    const capMax = round2_(principal * capFrac);
    const pct = capMax > 0 ? Math.round(feeAccum / capMax * 100) : 0;
    const atCap = capMax > 0 && pct >= 100;
    // Enlace WhatsApp al formulario de congelamiento (sólo si NO está ya congelado; filas vencidas).
    const freezeWa = isFrozen ? '' : (freezeFormWaFormula_(loan) || '');
    // Enlace al formulario de PLAN DE PAGO FLEXIBLE: disponible para cualquier vencido no congelado
    // (ya no requiere el 100% de mora; el prestatario elige cuotas y cadencia, el prestamista aprueba).
    const planWa = (!isFrozen) ? (planFormWaFormula_(loan) || '') : '';
    // COBRANZA (mora tardía): mensaje de correo (mailto) + WhatsApp con el detalle y las opciones.
    const freezeUrl = (!isFrozen && cobWebBase) ? (cobWebBase + '?page=congelar&loan=' + encodeURIComponent(loan.loanId)) : '';
    const planUrl = (!isFrozen && cobWebBase) ? (cobWebBase + '?page=plan&loan=' + encodeURIComponent(loan.loanId)) : '';
    const cobDaily = round2_(totalDue * feePctFrac);                  // recargo por día = % diario × total
    const capReach = capMax > 0 && cobDaily > 0 ? fmtDate_(addDays_(dueDate, Math.ceil(capMax / cobDaily))) : '';
    const cm = collectionMessage_(loan, { base: base, mora: feeAccum, totalPaid: totalPaid, days: days,
      deadlineStr: cobDeadlineStr, out: out, freezeUrl: freezeUrl, planUrl: planUrl, isFrozen: isFrozen,
      capAmount: capMax, maxTotal: round2_(base + capMax), dailyAmount: cobDaily, capReachStr: capReach,
      todayStr: fmtDate_(today), loanDateStr: (loan.loanDate instanceof Date ? fmtDate_(loan.loanDate) : ''),
      feePctText: lateFeePctText_(), alias: cobAlias, contact: cobContact, sig1: cobSig1, sig2: cobSig2 });
    const cobWaNum = waNumberAR_(loan.phone);
    const cobWa = cobWaNum ? ('=HYPERLINK("https://wa.me/' + cobWaNum + '?text=' + encodeURIComponent(cm.body) + '","📲 Cobranza")') : 'sin teléfono';
    const cobMail = loan.email ? ('=HYPERLINK("mailto:' + encodeURIComponent(loan.email) + '?subject=' + encodeURIComponent(cm.subject) + '&body=' + encodeURIComponent(cm.body) + '","✉ Cobranza")') : 'sin correo';
    // Correo genérico de PLAN DE PAGO: abre GMAIL (no envía) con los pasos + el enlace del plan.
    const pm = planEmailMessage_(loan, planUrl, cobPlanPctInt);
    const planMail = (loan.email && planUrl)
      ? ('=HYPERLINK("https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(loan.email) + '&su=' + encodeURIComponent(pm.subject) + '&body=' + encodeURIComponent(pm.body) + '","✉ Plan (Gmail)")')
      : (loan.email ? '' : 'sin correo');
    rows.push([loan.loanId, nameShown, loan.dni, loan.email, loan.phone, dueDate, days,
      totalDue, feeAccum, totalPaid, out, false, (lastNotice instanceof Date) ? lastNotice : '',
      atCap ? 'SÍ (100%)' : pct + '%', false, '', false, '', freezeWa, planWa, cobMail, cobWa, planMail]);
  }
  rows.sort((a, b) => b[6] - a[6]); // más atrasados primero
  if (rows.length) {
    sh.getRange(2, 1, rows.length, LATE_HEADERS.length).setValues(rows);   // valores (la col freeze = fórmula por coerción)
    sh.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 8, rows.length, 1).setNumberFormat(CFG.CURRENCY_FMT);
    sh.getRange(2, 9, rows.length, 3).setNumberFormat(CFG.CURRENCY_FMT);
    sh.getRange(2, 7, rows.length, 1).setNumberFormat('0');
    sh.getRange(2, LATE_SEND_COL, rows.length, 1).insertCheckboxes();
    sh.getRange(2, LATE_SENT_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, LATE_PLAN_COL, rows.length, 1).insertCheckboxes();
    sh.getRange(2, LATE_PLAN_SENT_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, LATE_DOC_COL, rows.length, 1).insertCheckboxes();
    sh.getRange(2, LATE_DOC_DONE_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    // Resaltado en BLOQUE (una llamada por columna): nombre congelado + tope de mora alcanzado.
    const nameBg = rows.map(r => [String(r[1]).indexOf('❄') === 0 ? '#d0e0e3' : null]);
    const capBg = rows.map(r => [String(r[LATE_CAP_COL - 1]).indexOf('SÍ') === 0 ? '#f4cccc' : null]);
    sh.getRange(2, 2, rows.length, 1).setBackgrounds(nameBg).setFontWeights(nameBg.map(x => [x[0] ? 'bold' : null]));
    sh.getRange(2, LATE_CAP_COL, rows.length, 1).setBackgrounds(capBg).setFontWeights(capBg.map(x => [x[0] ? 'bold' : null]));
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
    ['   Cuando "Tope de mora" muestra SÍ (100%), el recargo llegó a su máximo: tilde "Plan de pago 📅 (3 cuotas)" para reestructurar la deuda congelada en 3 cuotas mensuales y enviar el plan por correo al prestatario.'],
    ['   Tilde "PDF de cuotas 🧾" en "Pagos Atrasados" para generar el PDF con el cronograma de cuotas del préstamo (montos, vencimientos, pagado y saldo); queda guardado en la carpeta del prestatario.'],
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
 * por cuota = capFrac × su monto (capital+interés). Los pagos se aplican a la cuota más antigua
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
    cap: round2_((Number(c.amount) || 0) * capFrac),
    balance: round2_(Number(c.amount) || 0),   // capital+interés pendiente de la cuota
    mora: 0,            // mora VIGENTE (pendiente) de la cuota
    moraCharged: 0,     // mora TOTAL cobrada en la vida de la cuota (para respetar el tope ACUMULADO)
    cursor: 0,
  })).sort((a, b) => a.dueMs - b.dueMs);
  cuotas.forEach(c => { c.cursor = c.dueMs; });

  // Acumula la mora de cada cuota impaga hasta el instante t (tras su gracia, con su tope).
  // El tope se aplica sobre la mora TOTAL cobrada (moraCharged): la mora sólo corre mientras la
  // cuota tiene capital pendiente y se paga al final (ver applyPay), nunca supera su tope.
  const accrueTo = t => {
    for (const c of cuotas) {
      const startMs = c.dueMs + grace * DAY;
      if (c.balance <= 0 || feePct <= 0 || c.moraCharged >= c.cap || t <= startMs) { c.cursor = Math.max(c.cursor, t); continue; }
      const from = Math.max(c.cursor, startMs);
      const days = Math.floor((t - from) / DAY);
      if (days > 0) {
        const add = Math.min(round2_(round2_(c.amount * feePct) * days), round2_(c.cap - c.moraCharged));
        c.mora = round2_(c.mora + add);
        c.moraCharged = round2_(c.moraCharged + add);
      }
      c.cursor = t;
    }
  };
  // Aplica un pago CAPITAL-PRIMERO: dos pasadas sobre las cuotas más antiguas — primero todo el
  // capital+interés (c.balance), luego la mora (c.mora). Como la mora sólo devenga mientras el
  // capital sigue impago (ver accrueTo), cancelar el capital DETIENE el reloj de mora en la fecha
  // de ese pago: un préstamo pagado tarde sólo acumula mora hasta que se cubre el total base.
  const applyPay = amt => {
    let rem = round2_(amt);
    for (const c of cuotas) { if (rem <= 0) break; if (c.balance > 0) { const b = Math.min(rem, c.balance); c.balance = round2_(c.balance - b); rem = round2_(rem - b); } }
    for (const c of cuotas) { if (rem <= 0) break; if (c.mora > 0) { const m = Math.min(rem, c.mora); c.mora = round2_(c.mora - m); rem = round2_(rem - m); } }
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
 * Estado de CONGELAMIENTO de un préstamo (columnas "Congelado el" / "Días congelados" de
 * "Prestatarios"). `since` es la fecha en que se detuvo el reloj de la mora (Date) o null si
 * el préstamo NO está congelado; `daysAcc` son los días de atraso ya saltados por ciclos
 * PREVIOS de descongelado (número ≥ 0). Se lee por ID, igual que moraOverrideValue_.
 */
function frozenInfo_(loanId) {
  const out = { since: null, daysAcc: 0 };
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh ? sh.getLastRow() : 0;
  if (last < 2) return out;
  const id = String(loanId).trim();
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const sinceC = colByAny_(H, ['Congelado el']);
  const daysC = colByAny_(H, ['Días congelados', 'Dias congelados']);
  if (!sinceC && !daysC) return out;
  const width = Math.max(idC, sinceC || 0, daysC || 0);
  const data = sh.getRange(2, 1, last - 1, width).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idC - 1]).trim() === id) {
      if (sinceC) { const v = data[i][sinceC - 1]; if (v instanceof Date) out.since = v; }
      if (daysC) { const n = Number(data[i][daysC - 1]); if (isFinite(n) && n > 0) out.daysAcc = Math.floor(n); }
      return out;
    }
  }
  return out;
}
/**
 * Fecha EFECTIVA hasta la que debe correr la mora: si el préstamo está congelado, se recorta a
 * la fecha de congelamiento; luego se le restan los días ya saltados en descongelados previos
 * ("Días congelados"). Sólo afecta el DEVENGAMIENTO de la mora: los pagos se aplican igual
 * (los motores pasan el corte de pagos en el año 9999). Sin congelar, devuelve `asOf` tal cual.
 */
function effectiveMoraAsOf_(loanId, asOf) {
  asOf = (asOf instanceof Date) ? asOf : new Date();
  const f = frozenInfo_(loanId);
  let end = asOf;
  if (f.since instanceof Date && f.since.getTime() < end.getTime()) end = f.since;
  if (f.daysAcc > 0) end = new Date(end.getTime() - f.daysAcc * 86400000);
  return end;
}
/**
 * ¿Se puede CONGELAR este préstamo? Sólo los préstamos VENCIDOS (en mora) y con saldo base
 * pendiente. Devuelve {ok, reason}. Usado por el menú (Parte 1) y por el formulario/aprobación/
 * firma de la solicitud de congelamiento (Parte 2), para que la regla sea idéntica en todos lados.
 */
function loanFreezable_(loan, asOf) {
  asOf = (asOf instanceof Date) ? asOf : new Date();
  if (!loan || !(loan.dueDate instanceof Date)) return { ok: false, reason: 'El préstamo no tiene fecha de vencimiento.' };
  if (daysLate_(loan.dueDate, asOf) <= 0) return { ok: false, reason: 'Solo se pueden congelar préstamos VENCIDOS (en mora).' };
  const paid = loanPayments_(loan.loanId).reduce((s, p) => s + p.amount, 0);
  if (round2_(loan.totalDue - paid) <= 0.009) return { ok: false, reason: 'El préstamo ya está saldado.' };
  return { ok: true, reason: '' };
}
/** Fecha efectiva de corte de mora a partir de info de congelamiento EN MEMORIA (sin leer hoja). */
function effMoraAsOf2_(frozen, asOf) {
  let end = (asOf instanceof Date) ? asOf : new Date();
  if (frozen && frozen.since instanceof Date && frozen.since.getTime() < end.getTime()) end = frozen.since;
  if (frozen && frozen.daysAcc > 0) end = new Date(end.getTime() - frozen.daysAcc * 86400000);
  return end;
}
/** Igual que accruedMora_ pero SIN leer la hoja: recibe los pagos, el ajuste manual (override) y la
 *  info de congelamiento ya cargados en memoria. Para reconstrucciones masivas (Pagos Atrasados). */
function accruedMoraFast_(loan, asOf, pays, override, frozen) {
  if (override != null) return override;
  const eff = effMoraAsOf2_(frozen, asOf);
  const rate = loanRateForTerm_(loan.term);
  const sched = loanSchedule_(loan), far = new Date(9999, 0, 1);
  const withMora = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, eff, far, undefined, undefined, undefined, sched);
  const principalDue = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, eff, far, 0, undefined, undefined, sched);
  const m = Math.max(0, round2_(withMora - principalDue));
  const cap = moraCapAmount_(loan);                 // tope = % × capital
  return cap > 0 ? Math.min(m, cap) : m;
}
/** Todos los pagos agrupados por ID de préstamo, en UNA sola lectura de "Pagos". {id:[{date,amount}]}. */
function loanPaymentsAll_() {
  const sh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, PP.LOAN_ID, sh.getLastRow() - 1, 3).getValues();
  for (const r of data) {
    const id = String(r[0]).trim(); if (!id) continue;
    const amt = parseMoney_(r[2]); if (amt === 0) continue;
    (map[id] = map[id] || []).push({ date: r[1] instanceof Date ? r[1] : new Date(), amount: amt });
  }
  return map;
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
  // CONGELAMIENTO: la mora se devenga sólo hasta la fecha efectiva (congelado → se detiene el
  // reloj; menos los días saltados en descongelados previos). Los pagos no se tocan (corte 9999).
  const effAsOf = effectiveMoraAsOf_(loan.loanId, asOf);
  const rate = loanRateForTerm_(loan.term);
  const pays = loanPayments_(loan.loanId);
  const sched = loanSchedule_(loan), far = new Date(9999, 0, 1);
  const withMora = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, effAsOf, far, undefined, undefined, undefined, sched);
  // Capital+interés pendiente SIN mora: el MISMO motor con recargo diario 0. Robusto aunque
  // "Total a Pagar" venga vacío/erróneo en la fila (antes se restaba loan.totalDue y, si
  // faltaba, principalDue=0 y la mora absorbía el capital → duplicaba el "monto a pagar hoy").
  const principalDue = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, effAsOf, far, 0, undefined, undefined, sched);
  const m = Math.max(0, round2_(withMora - principalDue));
  const cap = moraCapAmount_(loan);                 // tope = % × capital (nunca más que el capital al 100%)
  return cap > 0 ? Math.min(m, cap) : m;
}
/**
 * Fecha en que los pagos ACUMULADOS alcanzaron el total base (capital+interés): la mora deja de
 * correr ahí, aunque existan pagos POSTERIORES que cubren la mora. Devuelve la Fecha de Pago del
 * pago que cruzó el total, o null si nunca se alcanzó. Misma regla que applyPay (capital-primero).
 */
function baseSettledDate_(loan, pays) {
  const sorted = pays.slice().sort((a, b) => new Date(a.date) - new Date(b.date));
  let acc = 0;
  for (const p of sorted) { acc = round2_(acc + p.amount); if (acc >= round2_(loan.totalDue) - 0.009) return new Date(p.date); }
  return null;
}
/**
 * Mora COBRADA (histórica) de un préstamo: el recargo por mora que se APLICÓ por los días de
 * atraso, congelado a la fecha en que se saldó (o a hoy si sigue impago). A diferencia de
 * accruedMora_ —que devuelve lo que AÚN se adeuda y por eso da 0 cuando ya se pagó—, esta deja
 * constancia del recargo que corresponde/correspondió, para mostrarlo en el estado de cuenta
 * aunque el préstamo esté saldado. Respeta la FECHA DE PAGO (no se calcula contra hoy si ya se
 * saldó). Espeja la fórmula de la hoja "Estudio de Estados": días × total × recargo diario,
 * menos la gracia, con tope. Un AJUSTE manual ("Mora (ajuste)") manda por encima del cálculo.
 */
function chargedMora_(loan, asOf) {
  const ov = moraOverrideValue_(loan.loanId);
  if (ov != null) return ov;
  asOf = (asOf instanceof Date) ? asOf : new Date();
  const pays = loanPayments_(loan.loanId);
  // Fecha hasta la que corrió la mora: la del pago que cubrió el total base (capital+interés); si
  // nunca se cubrió, hoy. Los pagos POSTERIORES (que sólo cubren la mora) NO la corren: se congela
  // al saldarse el base (misma regla que applyPay, capital-primero).
  let endDate = baseSettledDate_(loan, pays) || asOf;
  // CONGELAMIENTO: si está congelado (o quedan días saltados), recorta el reloj a la fecha efectiva.
  const effAsOf = effectiveMoraAsOf_(loan.loanId, asOf);
  if (effAsOf.getTime() < endDate.getTime()) endDate = effAsOf;
  const effDays = Math.max(0, daysLate_(loan.dueDate, endDate) - moraGraceDays_());
  if (effDays <= 0) return 0;
  const base = round2_(loan.totalDue);
  return Math.max(0, Math.min(round2_(base * lateFeeRate_() * effDays), round2_(base * moraCapFrac_())));
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
  // Columnas por NOMBRE de encabezado (robusto al layout migrado, donde el orden físico
  // difiere del mapa PB), con respaldo a PB para el esquema nuevo.
  const H = headerIndex_(bs);
  const paidCol = colByAny_(H, ['Total Pagado']) || PB.PAID;
  const balCol = colByAny_(H, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || PB.BALANCE;
  const stateCol = colByAny_(H, ['Estado']) || PB.STATE;
  const moraCol = colByAny_(H, ['Recargo por Mora (acum.)']) || PB.MORA_ACUM;
  const balMoraCol = colByAny_(H, ['Saldo con Mora']) || PB.BALANCE_MORA;
  if (!loanId) {
    [balCol, stateCol, moraCol, balMoraCol].forEach(c => bs.getRange(row, c).clearContent());
    return;
  }
  // Total Pagado (SUMIF), Saldo, Recargo por Mora (de "Pagos Atrasados"), Saldo con Mora y
  // Estado son fórmulas vivas; las reafirmamos por si la fila fue pegada, migrada o creada
  // al alta (un "Total Pagado" estático dejaba pagos sin acreditar). Se delega en
  // writeBorrowerRowFormulas_, que resuelve TODAS las columnas (y las letras internas de las
  // fórmulas) por NOMBRE — así las fórmulas no quedan apuntando a columnas equivocadas en el
  // layout migrado (los builders borrower*Formula_ usan letras fijas del mapa PB).
  if (typeof writeBorrowerRowFormulas_ === 'function') { writeBorrowerRowFormulas_(bs, row); return; }
  bs.getRange(row, paidCol).setFormula(borrowerPaidFormula_(row));
  bs.getRange(row, balCol).setFormula(borrowerBalanceFormula_(row));
  bs.getRange(row, moraCol).setFormula(borrowerMoraFormula_(row));
  bs.getRange(row, balMoraCol).setFormula(borrowerBalanceWithMoraFormula_(row));
  bs.getRange(row, stateCol).setFormula(borrowerStateFormula_(row));
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
/** @param {object=} sig  Firma opcional: { dataUrl, signedAt (Date), method, userAgent, ip, hash, device, deviceFp }. */
function agreementHtml_(loan, sig) {
  const lender = companyName_(), footer = getSetting_('Pie del Contrato');
  const signatories = ownerSignatories_();   // ambos titulares firman todos los contratos
  const days = termDays_(loan.term) || 0;
  // Cronograma de cuotas (tramo grande = 3 a 6 cuotas mensuales manejables según el monto; tramos chicos = 1 pago).
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
    cuotasNota = 'El total se devuelve en ' + nCuotas + ' cuotas mensuales iguales, según el cronograma de arriba (vencimientos mensuales a los 30, 60, … días).';
  }
  const feeDay = round2_(loan.totalDue * lateFeeRate_());              // recargo diario sobre el total
  const moraGrace = moraGraceDays_();                                  // días de gracia antes de la mora
  const feeCap = moraCapAmount_(loan);                                 // tope acumulado del recargo (% × capital)
  const maxTotal = round2_(loan.totalDue + feeCap);                    // total + tope del recargo
  const daysToCap = feeDay > 0 ? moraGrace + Math.ceil(feeCap / feeDay) : 0;
  const payMethod = loan.payMethod || 'Mercado Pago';
  // Fecha de firma del Prestamista: con valor cuando el contrato está firmado; en blanco si aún no.
  const firmadoEl = (sig && sig.dataUrl) ? fmtDate_(sig.signedAt || new Date()) : '__________';
  const fill = v => (v && String(v).trim()) ? esc_(v) : '<span style="color:#9aa4ad;letter-spacing:1px">________________________</span>';
  const kv = (k, v) => `<tr><td class="k">${k}</td><td>${v}</td></tr>`;
  // Resumen de ubicación para el panel de firma (prioridad GPS > IP > zona horaria).
  const locSummary = s => {
    const d = (s && s.device) || {}, l = d.loc || {};
    if (l.source === 'GPS' && l.gps && l.gps.indexOf('GPS:') !== 0) return l.gps + (l.gpsAcc != null ? ' (±' + l.gpsAcc + ' m)' : '') + ' · GPS';
    if (l.ipText) return l.ipText + ' · IP';
    if (l.source === 'desactivada') return 'No capturada';
    return d.tz ? ('Zona horaria: ' + d.tz) : '—';
  };
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
      <p style="margin:8px 0 0">Si los recargos alcanzan ese máximo, el Prestatario será incorporado a un <b>plan de pago en 3 cuotas
         mensuales</b>: la deuda total a esa fecha (capital, interés y recargos) queda congelada y se divide en 3 pagos mensuales iguales.
         Mientras el plan se cumpla, no se generan nuevos recargos; el detalle de montos y vencimientos se informa por correo.</p>
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
      <div class="box-title">Constancia de firma electrónica</div>
      <table style="margin:0;font-size:10pt">
        ${kv('Firmó', esc_(loan.name) + ' · DNI ' + esc_(loan.dni))}
        ${kv('Correo confirmado', esc_(loan.email))}
        ${kv('Identidad verificada con', 'DNI ' + esc_(loan.dni) + ' y correo registrado; foto del DNI (frente y dorso, Anexo I)')}
        ${kv('Documento', 'Foto del DNI, frente y dorso — Anexo I')}
        ${kv('Fecha y hora', fmtDateTime_(sig.signedAt || new Date()))}
        ${kv('Cómo se firmó', sig.method === 'upload' ? 'Imagen de firma cargada' : 'Firma dibujada en pantalla')}
        ${kv('Dirección IP', fill(sig.ip))}
        ${kv('Ubicación', esc_(locSummary(sig)))}
        ${kv('Equipo / SO', fill((sig.device && sig.device.os) || ''))}
        ${kv('Navegador', fill((sig.device && sig.device.browser) || ''))}
        ${kv('Pantalla', fill((sig.device && sig.device.screen) ? (sig.device.screen + (sig.device.dpr ? (' · DPR ' + sig.device.dpr) : '')) : ''))}
        ${kv('Automatización detectada', (sig.device && sig.device.webdriver) ? 'Sí (revisar)' : 'No (firmante humano)')}
        ${kv('Dispositivo (completo)', esc_(String(sig.userAgent || '').slice(0, 200)))}
        ${kv('Huella de dispositivo (SHA-256)', sig.deviceFp ? '<span style="font-family:monospace;font-size:8.5pt;word-break:break-all">' + esc_(sig.deviceFp) + '</span>' : '<span class="note">No disponible</span>')}
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
  const feeCap = moraCapAmount_(loan);    // tope del recargo (% del capital)
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
    '<li>Recargo por mora: <b>' + fmtMoney_(round2_(loan.totalDue * lateFeeRate_())) + ' por día</b> de atraso (' + lateFeePctText_() + ' sobre el total a devolver, tope del ' + moraCapPctText_() + ' del capital)</li></ul>' +
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
function restructurarPrestamo_(loanId, nCuotas, primeraFecha, opts) {
  return guard_('restructurarPrestamo_', function () {
    opts = opts || {};
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

    // FIJAR la mora congelada en "Mora (ajuste)": al salir de "Pagos Atrasados" el
    // recargo automático (VLOOKUP a esa hoja) vuelve a 0 y la deuda congelada se
    // perdería de "Saldo con Mora"/Estado. La fórmula deja la mora restante del plan:
    // adeudado congelado − pagos posteriores, sin contar lo que aún cubre capital+interés.
    const paid0 = round2_(pays.reduce((s, p) => s + p.amount, 0));
    const moraCongelada = round2_(Math.max(0, owed - Math.max(0, round2_(loan.totalDue - paid0))));
    const adjC = colByAny_(headerIndex_(bs), ['Mora (ajuste)']) || PB.MORA_ADJ;
    const TO = colL_(PB.TOTAL), PA = colL_(PB.PAID);
    if (adjC) bs.getRange(row, adjC).setFormula(
      `=MAX(0,${round2_(owed + paid0)}-$${PA}${row}-MAX(0,$${TO}${row}-$${PA}${row}))`);

    // Registrar el plan como nota auditable en la celda del ID del préstamo.
    const planTxt = plan.map(p => '  Cuota ' + p.n + '/' + n + ': ' + fmtMoney_(p.monto) + ' — vence ' + fmtDate_(p.due)).join('\n');
    bs.getRange(row, PB.LOAN_ID).setNote(
      'REESTRUCTURADO el ' + fmtDate_(today) + '.\n' +
      'Vencimiento original: ' + fmtDate_(origDue instanceof Date ? origDue : loan.dueDate) + '.\n' +
      'Adeudado congelado (capital + interés + mora a hoy): ' + fmtMoney_(owed) + '.\n' +
      'Plan de pago:\n' + planTxt + '\n' +
      'La mora deja de crecer mientras se cumpla el plan (vencimiento movido al ' + fmtDate_(finalDue) + ').\n' +
      'Mora congelada (' + fmtMoney_(moraCongelada) + ') fijada en "Mora (ajuste)"; el menú 💸 Mora la puede condonar/cambiar.\n' +
      'Registrá cada pago en la hoja "Pagos" como siempre.');

    if (!opts.silent) try { ss.toast(loanId + ' reestructurado: ' + fmtMoney_(owed) + ' en ' + n + ' cuota(s). Mora congelada al ' + fmtDate_(today) + '.', '🔄 Reestructuración', 8); } catch (e) { }
    return { loanId: loanId, owed: owed, plan: plan, paid0: paid0, mora: moraCongelada, row: row };
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
 * Núcleo: REGENERA el cronograma de cuotas de un préstamo con la deuda ACTUAL (capital +
 * interés + mora a hoy) CONGELADA y la 1.ª cuota venciendo la SEMANA PRÓXIMA (hoy + 7 días);
 * las siguientes, mensuales. Reutiliza restructurarPrestamo_ (congela mora, mueve el
 * vencimiento, fija "Mora (ajuste)" y anota el plan) pasándole hoy+7 como 1.ª fecha, y reescribe
 * las filas de "Cuotas" con el nuevo cronograma. Devuelve el resumen del plan.
 */
function regenerarCuotasPrestamo_(loanId, nCuotas) {
  return guard_('regenerarCuotasPrestamo_', function () {
    const ss = getSS_();
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan) throw new Error('No se encontró el préstamo ' + loanId + '.');
    const today = startOfDay_(new Date());
    const firstDue = addDays_(today, 7); // la 1.ª cuota vence la SEMANA SIGUIENTE, no el mismo día
    const n = Math.max(1, Math.floor(Number(nCuotas) || 3));
    // Congela la deuda a hoy y arma el plan con la 1.ª cuota venciendo la semana próxima.
    const out = restructurarPrestamo_(loan.loanId, n, firstDue, { silent: true });
    // Reescribe el cronograma en "Cuotas" con las fechas del plan (cuotasForLoan_ borra las viejas).
    const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || (typeof setupCuotas_ === 'function' ? setupCuotas_(ss) : null);
    if (cs) cuotasForLoan_(cs, loan.loanId, loan.loanDate, out.owed, out.plan.length, null,
      { dates: out.plan.map(p => p.due), paidOffset: out.paid0 });
    try { ss.toast(loan.loanId + ': cuotas regeneradas (' + out.plan.length + ') desde ' + fmtDate_(out.plan[0].due) + '.', '🔁 Cuotas', 6); } catch (e) {}
    return out;
  });
}

/**
 * Menú: regenera el cronograma del préstamo de la FILA ACTIVA (en "Prestatarios" o "Pagos
 * Atrasados") con la deuda congelada a hoy y la 1.ª cuota venciendo la SEMANA PRÓXIMA (hoy + 7
 * días). Pide la cantidad de cuotas (por defecto 3). Luego tildá "PDF de cuotas 🧾" para el PDF.
 */
function regenerarCuotasPrestamoUI() {
  const ss = getSS_(), ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet(), name = sh.getName(), ar = ss.getActiveRange();
  let loanId = '';
  if (ar) {
    if (name === CFG.SHEETS.BORROWERS) loanId = String(sh.getRange(ar.getRow(), PB.LOAN_ID).getValue()).trim();
    else if (name === CFG.SHEETS.LATE) loanId = String(sh.getRange(ar.getRow(), 1).getValue()).trim();
  }
  if (!/^L-/i.test(loanId)) {
    ui.alert('Seleccioná la fila de un préstamo (ID "L-…") en "Prestatarios" o "Pagos Atrasados" y volvé a ejecutar.');
    return;
  }
  const resp = ui.prompt('Regenerar cuotas (desde la semana próxima)',
    'Préstamo ' + loanId + ': se congela la deuda a hoy (capital + interés + mora) y se arma un nuevo cronograma con la 1.ª cuota venciendo la SEMANA PRÓXIMA (hoy + 7 días).\n\nCantidad de cuotas (por defecto 3):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(String(resp.getResponseText()).replace(/\D/g, ''), 10) || 3;
  const out = regenerarCuotasPrestamo_(loanId, n);
  if (out) ui.alert('Cuotas regeneradas',
    out.loanId + ': ' + fmtMoney_(out.owed) + ' en ' + out.plan.length + ' cuota(s), desde ' + fmtDate_(out.plan[0].due) + ' (la 1.ª vence la semana próxima). ' +
    'La mora quedó congelada. Tildá "PDF de cuotas 🧾" en "Pagos Atrasados" para el PDF actualizado.', ui.ButtonSet.OK);
}

/**
 * Casilla "Plan de pago 📅 (3 cuotas)" de "Pagos Atrasados": reestructura el préstamo
 * de la fila en 3 cuotas mensuales iguales (deuda con mora congelada a hoy), reemplaza
 * sus filas de "Cuotas" por las del plan y envía el plan por correo al prestatario.
 * La columna "Tope de mora" indica los candidatos (SÍ = recargo al máximo), pero se
 * puede aplicar antes de llegar al tope si el prestamista lo decide.
 */
function planPagoDesdeAtrasos_(sh, row) {
  const ss = getSS_();
  const loanId = String(sh.getRange(row, 1).getValue()).trim();
  if (!loanId) { ss.toast('La fila seleccionada no tiene un préstamo.', '⚠ Plan de pago', 6); return; }
  const loan = findLoanById_(loanId);
  if (!loan) { ss.toast('No se encontró el préstamo ' + loanId + ' en "Prestatarios".', '⚠ Plan de pago', 6); return; }
  // Guardas: ya reestructurado / al día (el vencimiento movido queda en el futuro), o
  // plan ya enviado hoy desde esta misma hoja (evita doble clic → doble correo).
  if (daysLate_(loan.dueDate, new Date()) <= 0) {
    ss.toast(loanId + ' ya está reestructurado o al día (vencimiento ' + fmtDate_(loan.dueDate) + '). No se aplicó otro plan.', '⚠ Plan de pago', 8);
    return;
  }
  if (sh.getRange(row, LATE_PLAN_SENT_COL).getValue() instanceof Date) {
    ss.toast(loanId + ': el plan ya fue enviado (ver "Plan enviado").', '⚠ Plan de pago', 6);
    return;
  }

  // 1) Reestructurar (congela mora, mueve vencimiento, fija "Mora (ajuste)", anota el plan).
  const out = restructurarPrestamo_(loanId, 3, null, { silent: true });

  // 2) Cuotas: reemplazar las filas viejas (VENCIDAS) por las 3 del plan. paidOffset evita
  //    que los pagos previos al plan se dupliquen en el reparto (el plan ya es neto de ellos).
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || (typeof setupCuotas_ === 'function' ? setupCuotas_(ss) : null);
  if (cs) {
    deleteCuotasRowsForLoan_(cs, loanId);
    cuotasForLoan_(cs, loanId, loan.loanDate, out.owed, 3, null, { dates: out.plan.map(p => p.due), paidOffset: out.paid0 });
  }

  // 3) Sello (evita reaplicar el plan con otro clic) + aviso. NO se envía correo en este paso.
  sh.getRange(row, LATE_PLAN_SENT_COL).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  ss.toast(loanId + ' reestructurado: ' + fmtMoney_(out.owed) + ' en 3 cuotas (sin enviar correo).', '📅 Plan de pago', 8);
}

/** Correo con el plan de pago en 3 cuotas (deuda congelada, fechas y montos). */
function emailPlanPago_(loan, out, toEmail) {
  const to = String(toEmail || loan.email || '').trim();
  if (!to) return;
  const cuotasTxt = out.plan.map(p => 'Cuota ' + p.n + ' de ' + out.plan.length + ': ' + fmtMoney_(p.monto) + ' — vence el ' + fmtDate_(p.due)).join('. ');
  const filas = out.plan.map(p =>
    '<tr><td style="padding:6px 10px;border:1px solid #d7dee8">Cuota ' + p.n + ' de ' + out.plan.length + '</td>' +
    '<td style="padding:6px 10px;border:1px solid #d7dee8;text-align:right"><b>' + fmtMoney_(p.monto) + '</b></td>' +
    '<td style="padding:6px 10px;border:1px solid #d7dee8">' + esc_(fmtDate_(p.due)) + '</td></tr>').join('');
  sendBrandedEmail_(to, 'Plan de pago en 3 cuotas — préstamo ' + loan.loanId,
    'Estimado/a ' + loan.name + ': su préstamo ' + loan.loanId + ' fue reestructurado en un plan de 3 cuotas mensuales. ' +
    'La deuda quedó congelada al día de hoy en ' + fmtMoney_(out.owed) + ' (incluye capital, interés y recargo por mora). ' +
    cuotasTxt + '. El recargo por mora deja de crecer mientras cumpla el plan. ' +
    'Si realizó un pago recientemente, se descuenta de la primera cuota.',
    '<p>Estimado/a ' + esc_(loan.name) + ',</p>' +
    '<p>Le informamos que su préstamo <b>' + esc_(loan.loanId) + '</b> fue reestructurado en un <b>plan de pago de 3 cuotas mensuales</b>. ' +
    'El monto total adeudado quedó <b>congelado al día de hoy en ' + fmtMoney_(out.owed) + '</b> (incluye capital, interés y el recargo por mora acumulado).</p>' +
    '<table style="border-collapse:collapse;margin:10px 0;font-size:14px">' +
    '<tr><th style="padding:6px 10px;border:1px solid #d7dee8;background:#eef3fb;text-align:left">Cuota</th>' +
    '<th style="padding:6px 10px;border:1px solid #d7dee8;background:#eef3fb;text-align:right">Monto</th>' +
    '<th style="padding:6px 10px;border:1px solid #d7dee8;background:#eef3fb;text-align:left">Vencimiento</th></tr>' +
    filas + '</table>' +
    '<p><b>El recargo por mora deja de crecer mientras cumpla el plan.</b> Le pedimos abonar cada cuota a más tardar en su fecha de vencimiento.</p>' +
    '<p>Si ya realizó un pago reciente, será descontado de la primera cuota. Ante cualquier duda, responda este correo.</p>');
}

/** Cuotas registradas del préstamo en la hoja "Cuotas" (ordenadas por número). */
function cuotasDeLoan_(loanId) {
  const cs = getSS_().getSheetByName(CFG.SHEETS.INSTALLMENTS);
  if (!cs || cs.getLastRow() < 2) return [];
  const out = [];
  cs.getRange(2, 1, cs.getLastRow() - 1, CU.STATE).getValues().forEach(cr => {
    if (String(cr[CU.LOAN_ID - 1]).trim() !== loanId) return;
    out.push({
      num: Number(cr[CU.NUM - 1]) || 0,
      due: (cr[CU.DUE - 1] instanceof Date) ? cr[CU.DUE - 1] : null,
      monto: round2_(Number(cr[CU.AMOUNT - 1]) || 0),
      pagado: round2_(Number(cr[CU.PAID - 1]) || 0),
      saldo: round2_(Number(cr[CU.BALANCE - 1]) || 0),
      estado: String(cr[CU.STATE - 1] || '').trim(),
    });
  });
  return out.sort((a, b) => a.num - b.num);
}

/**
 * Casilla "PDF de cuotas 🧾" de "Pagos Atrasados": genera el PDF con el cronograma de
 * cuotas del préstamo de la fila (tomado de la hoja "Cuotas"; si el préstamo no tiene
 * cronograma, su pago único) y lo guarda en la carpeta del prestatario. Sella la fecha
 * en "PDF generado".
 */
function docCuotasDesdeAtrasos_(sh, row) {
  const ss = getSS_();
  const loanId = String(sh.getRange(row, 1).getValue()).trim();
  if (!loanId) { ss.toast('La fila seleccionada no tiene un préstamo.', '⚠ PDF de cuotas', 6); return; }
  const f = generateCuotasPdfFile_(loanId);
  sh.getRange(row, LATE_DOC_DONE_COL).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  ss.toast('PDF de cuotas de ' + loanId + ' guardado en la carpeta del prestatario.', '🧾 PDF de cuotas', 8);
  return f;
}

/**
 * Genera el PDF del cronograma de cuotas de un préstamo (de la hoja "Cuotas"; si no tiene
 * cronograma, su pago único) y lo guarda en la carpeta del prestatario. Devuelve el archivo.
 * Reutilizado por "Pagos Atrasados" y por la hoja "En Plan". Lanza si no se encuentra el préstamo.
 */
function generateCuotasPdfFile_(loanId) {
  const loan = findLoanById_(String(loanId || '').trim());
  if (!loan) throw new Error('No se encontró el préstamo ' + loanId + ' en "Prestatarios".');
  let cuotas = cuotasDeLoan_(loan.loanId);
  if (!cuotas.length) {
    const paid = round2_(loanPayments_(loan.loanId).reduce((s, p) => s + p.amount, 0));
    const total = round2_(Number(loan.totalDue) || 0);
    cuotas = [{
      num: 1, due: (loan.dueDate instanceof Date) ? loan.dueDate : null, monto: total,
      pagado: Math.min(paid, total), saldo: Math.max(0, round2_(total - paid)),
      estado: daysLate_(loan.dueDate, new Date()) > 0 ? CST.OVERDUE : CST.PENDING,
    }];
  }
  return savePdfToBorrower_(cuotasDocHtml_(loan, cuotas), 'Cuotas ' + loan.loanId + ' - ' + loan.name + '.pdf', loan);
}

/** HTML del documento de cuotas: detalle del préstamo + cronograma con pagado/saldo por cuota. */
function cuotasDocHtml_(loan, cuotas) {
  const lender = companyName_(), footer = getSetting_('Pie del Contrato');
  const totMonto = round2_(cuotas.reduce((s, c) => s + c.monto, 0));
  const totPag = round2_(cuotas.reduce((s, c) => s + c.pagado, 0));
  const totSaldo = round2_(cuotas.reduce((s, c) => s + c.saldo, 0));
  // Capital+interés ORIGINAL (sin mora). Si "Total a Pagar" viene vacío, se deriva del préstamo.
  const baseTotal = round2_(Number(loan.totalDue) || round2_(loan.principal * (1 + loanRateForTerm_(loan.term))));
  // Total realmente PAGADO del préstamo. En un préstamo reestructurado, los pagos ANTERIORES al
  // plan YA redujeron el cronograma (no figuran en el "Pagado" por cuota): se incorporan en el
  // resumen para que la cuenta cierre (deuda original + mora − pagado = cronograma).
  const totalPaidReal = round2_(loanPayments_(loan.loanId).reduce((s, p) => s + p.amount, 0));
  // ¿Préstamo REESTRUCTURADO (plan de pago)? Tiene la mora CONGELADA en "Mora (ajuste)" (la fija
  // restructurarPrestamo_); entonces la mora ya está DENTRO de las cuotas y NO se vuelve a sumar.
  const adj = (typeof moraOverrideValue_ === 'function') ? moraOverrideValue_(loan.loanId) : null;
  const restructured = (adj != null && adj > 0.009);
  const feeAccum = restructured ? round2_(adj) : round2_(accruedMora_(loan) || 0);
  const montoHoy = restructured ? totSaldo : round2_(totSaldo + feeAccum);
  const cell = 'padding:6px 8px;border:1px solid #ccc';
  const filas = cuotas.map(c => {
    const late = c.saldo > 0.009 && (c.due instanceof Date) && daysLate_(c.due, new Date()) > 0;
    return `<tr${late ? ' style="background:#fdecec"' : ''}>` +
      `<td style="${cell};text-align:center">${c.num} de ${cuotas.length}</td>` +
      `<td style="${cell}">${c.due instanceof Date ? fmtDate_(c.due) : '—'}</td>` +
      `<td style="${cell};text-align:right">${fmtMoney_(c.monto)}</td>` +
      `<td style="${cell};text-align:right">${fmtMoney_(c.pagado)}</td>` +
      `<td style="${cell};text-align:right"><b>${fmtMoney_(c.saldo)}</b></td>` +
      `<td style="${cell};text-align:center">${esc_(c.estado)}</td></tr>`;
  }).join('');
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;border-bottom:2px solid #1c4587;padding-bottom:8px}
    h2{font-size:13pt;color:#1c4587;margin-top:24px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px} td,th{padding:6px 8px;border:1px solid #ccc}
    th{background:#eef3fb;text-align:left}
    td.k{background:#eef3fb;font-weight:bold;width:45%}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>CRONOGRAMA DE CUOTAS</h1>
    <p>Préstamo <b>${esc_(loan.loanId)}</b> de <b>${esc_(loan.name)}</b> (DNI ${esc_(loan.dni)}) con ${esc_(lender)}, al <b>${fmtDate_(new Date())}</b>.</p>
    <h2>Resumen</h2>
    <table>
      <tr><td class="k">Capital</td><td>${fmtMoney_(loan.principal)}</td></tr>
      ${restructured ? `<tr><td class="k">Total a pagar original (capital + interés)</td><td>${fmtMoney_(baseTotal)}</td></tr>` : ''}
      ${feeAccum > 0 ? `<tr><td class="k" style="color:#900">Recargo por mora ${restructured ? 'congelado (incluido en las cuotas)' : 'acumulado'}</td><td style="color:#900"><b>${fmtMoney_(feeAccum)}</b></td></tr>` : ''}
      ${totalPaidReal > 0.009 ? `<tr><td class="k">Total pagado${restructured ? ' (descontado del cronograma)' : ''}</td><td>${fmtMoney_(totalPaidReal)}</td></tr>` : `<tr><td class="k">Total pagado</td><td>${fmtMoney_(totPag)}</td></tr>`}
      <tr><td class="k">Total del cronograma (${cuotas.length} cuota${cuotas.length === 1 ? '' : 's'})</td><td><b>${fmtMoney_(totMonto)}</b></td></tr>
      <tr><td class="k">Saldo pendiente del cronograma</td><td><b>${fmtMoney_(totSaldo)}</b></td></tr>
      ${feeAccum > 0 ? `<tr><td class="k">Monto a pagar hoy (saldo + mora)</td><td><b>${fmtMoney_(montoHoy)}</b></td></tr>` : ''}
    </table>
    <h2>Cuotas</h2>
    <table><tr><th style="text-align:center">Cuota</th><th>Vencimiento</th><th style="text-align:right">Monto</th><th style="text-align:right">Pagado</th><th style="text-align:right">Saldo</th><th style="text-align:center">Estado</th></tr>
    ${filas}</table>
    <p style="margin-top:14px;font-size:11pt;color:#555">Los pagos se aplican a la cuota más antigua con saldo. El recargo por mora (si corresponde) se abona junto con la cuota.</p>
    <p class="foot">${esc_(footer)}</p></body></html>`;
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

/**
 * Menú: CONGELA el préstamo de la fila activa de "Prestatarios" — la mora deja de crecer desde
 * la fecha elegida (por defecto hoy; se puede retro-fechar). Escribe esa fecha en "Congelado el";
 * el Estado pasa a CONGELADO (con su color) y "Recargo por Mora"/"Saldo con Mora" se congelan solos.
 */
function congelarPrestamoFilaActiva() {
  const ss = getSS_(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== CFG.SHEETS.BORROWERS) { ui.alert('Abrí la hoja "Prestatarios", seleccioná la fila del préstamo y volvé a usar esta opción.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Seleccioná la fila de un préstamo.'); return; }
  const H = headerIndex_(sh);
  const sinceC = colByAny_(H, ['Congelado el']);
  if (!sinceC) { ui.alert('Falta la columna "Congelado el". Ejecutá ① Configurar o 🆕 Aplicar novedades.'); return; }
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const loanId = String(sh.getRange(row, idC).getValue()).trim();
  if (!loanId) { ui.alert('La fila seleccionada no tiene un préstamo.'); return; }
  // Solo se pueden congelar préstamos vencidos (en mora) y con saldo pendiente.
  const fz = loanFreezable_(findLoanById_(loanId));
  if (!fz.ok) { ui.alert(fz.reason); return; }
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const resp = ui.prompt('Congelar préstamo ' + loanId,
    'Fecha desde la que la mora deja de crecer (AAAA-MM-DD o DD/MM/AAAA). Dejá vacío para HOY (' + fmtDate_(hoy) + '):',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const txt = String(resp.getResponseText()).trim();
  let fecha = hoy;
  if (txt) { const d = parseFreezeDate_(txt); if (!d) { ui.alert('Fecha inválida. Usá AAAA-MM-DD o DD/MM/AAAA (p. ej. ' + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ').'); return; } fecha = d; }
  sh.getRange(row, sinceC).setValue(fecha).setNumberFormat('yyyy-mm-dd');
  try { ss.toast('Préstamo ' + loanId + ' CONGELADO desde ' + fmtDate_(fecha) + '. La mora dejó de crecer.', '❄ Congelado', 6); } catch (e) { }
}
/**
 * Menú: DESCONGELA el préstamo de la fila activa. La mora vuelve a correr desde hoy, pero los
 * días que estuvo congelado NO se cobran: se acumulan en "Días congelados" y se descuentan del
 * atraso para siempre (ver effectiveMoraAsOf_). Vacía "Congelado el".
 */
function descongelarPrestamoFilaActiva() {
  const ss = getSS_(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== CFG.SHEETS.BORROWERS) { ui.alert('Abrí la hoja "Prestatarios", seleccioná la fila del préstamo y volvé a usar esta opción.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Seleccioná la fila de un préstamo.'); return; }
  const H = headerIndex_(sh);
  const sinceC = colByAny_(H, ['Congelado el']);
  const daysC = colByAny_(H, ['Días congelados', 'Dias congelados']);
  if (!sinceC || !daysC) { ui.alert('Faltan las columnas de congelamiento. Ejecutá ① Configurar o 🆕 Aplicar novedades.'); return; }
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const loanId = String(sh.getRange(row, idC).getValue()).trim();
  if (!loanId) { ui.alert('La fila seleccionada no tiene un préstamo.'); return; }
  const since = sh.getRange(row, sinceC).getValue();
  if (!(since instanceof Date)) { ui.alert('El préstamo ' + loanId + ' no está congelado.'); return; }
  const loan = findLoanById_(loanId);
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  // Días de atraso saltados DURANTE el congelamiento (sólo los que habrían devengado mora).
  const gap = loan ? Math.max(0, daysLate_(loan.dueDate, hoy) - daysLate_(loan.dueDate, since)) : Math.max(0, daysLate_(since, hoy));
  const prev = Number(sh.getRange(row, daysC).getValue()) || 0;
  sh.getRange(row, daysC).setValue(prev + gap);
  sh.getRange(row, sinceC).clearContent();
  try { ss.toast('Préstamo ' + loanId + ' DESCONGELADO. No se cobran ' + gap + ' día(s) de atraso del período congelado.', '☀ Descongelado', 6); } catch (e) { }
}
/** Parsea la fecha de congelamiento ingresada por el usuario: AAAA-MM-DD o DD/MM/AAAA (o con "-"). */
function parseFreezeDate_(txt) {
  txt = String(txt || '').trim();
  let y, mo, da, m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(txt);   // AAAA-MM-DD
  if (m) { y = +m[1]; mo = +m[2]; da = +m[3]; }
  else { m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(txt); if (m) { da = +m[1]; mo = +m[2]; y = +m[3]; } } // DD/MM/AAAA
  if (!m || mo < 1 || mo > 12 || da < 1 || da > 31) return null;
  const d = new Date(y, mo - 1, da);
  return (d.getFullYear() === y && d.getMonth() === mo - 1 && d.getDate() === da) ? d : null;
}

/** Datos derivados del estado de cuenta (compartidos por el PDF y el correo). */
function statementData_(loan) {
  const pays = loanPayments_(loan.loanId).sort((a, b) => a.date - b.date);
  const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
  // Saldo base (mismo criterio que la hoja "Estudio de Estados": Total a Pagar −
  // Total Pagado). Si está saldado, no hay atraso ni recargo por mora.
  const out = Math.max(0, round2_(loan.totalDue - totalPaid));
  const paid = out <= 0.009;
  // Mora COBRADA (histórica): respeta la Fecha de Pago y queda como constancia aunque el
  // préstamo esté saldado. Días de atraso = hasta el pago que cubrió el total base (capital+
  // interés), o hasta hoy si sigue impago. Espeja la hoja "Estudio de Estados" y chargedMora_.
  let endDate = baseSettledDate_(loan, pays) || new Date();
  // CONGELAMIENTO: corta los días de atraso en la fecha efectiva, igual que chargedMora_, para
  // que los "días de atraso" mostrados coincidan con el recargo congelado.
  const effAsOf = effectiveMoraAsOf_(loan.loanId, new Date());
  if (effAsOf.getTime() < endDate.getTime()) endDate = effAsOf;
  const dLate = daysLate_(loan.dueDate, endDate);
  const feeAccum = chargedMora_(loan);
  const overdue = !paid && (loan.dueDate instanceof Date) && new Date().getTime() > loan.dueDate.getTime();
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
  const feeRows = d.feeAccum > 0 ?
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
/**
 * Actualiza SOLO la hoja "Pagos Atrasados" (recargo por mora) — versión LIGERA de ⑤ para
 * evitar el tiempo de espera de refreshAll, que además rearma Resumen, Panel, Estadísticas,
 * Estados y Recordatorios. Reconstruye el detalle de mora; las columnas "Recargo por Mora
 * (acum.)" / "Saldo con Mora" de "Prestatarios" (VLOOKUP a esta hoja) y el Estado se
 * recalculan solos. Úsela cuando un préstamo pagado tarde no muestre la mora.
 */
function refreshLateOnly() {
  guard_('refreshLateOnly', function () {
    const ss = getSS_();
    rebuildLateSheet_(ss);
    SpreadsheetApp.flush();
    try { ss.toast('Pagos Atrasados (mora) actualizado.', '⑤· Mora', 5); } catch (e) {}
  });
}
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
  setupSummary_(ss); setupPanel_(ss); setupStats_(ss); setupStatements_(ss); rebuildLateSheet_(ss); rebuildEnPlanSheet_(ss); rebuildRemindersSheet_(ss);
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
  if (page === 'firmar-congelamiento') return HtmlService.createHtmlOutput(freezeSigningPageHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Firmar Congelamiento').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (page === 'firmar-plan') return HtmlService.createHtmlOutput(planSigningPageHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Firmar Acuerdo de Plan de Pago').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (page === 'firmar') return HtmlService.createHtmlOutput(signingPageHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Firmar Contrato').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (page === 'congelar') return HtmlService.createHtmlOutput(freezeRequestFormHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Solicitud de congelamiento').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (page === 'plan') return HtmlService.createHtmlOutput(planRequestFormHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Solicitud de plan de pago').addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
      dni = String(form.dni || '').trim(),
      amount = Number(String(form.amount || '').replace(/\D/g, '')) || '',
      term = Number(form.term) || '', notes = String(form.notes || '').trim();
    let phone = String(form.phone || '').trim();
    // Forma de pago elegida por el prestatario (predeterminada: Mercado Pago) + datos opcionales.
    const payMethodSel = String(form.payMethod || '').trim() || 'Mercado Pago';
    const payDetails = String(form.payDetails || '').trim();
    const payMethodFull = payDetails ? (payMethodSel + ' — ' + payDetails) : payMethodSel;
    if (!name) throw new Error('El nombre completo es obligatorio.');
    if (name.length < 10) throw new Error('El nombre completo debe tener al menos 10 caracteres.'); // V-35
    if (!email) throw new Error('El correo electrónico es obligatorio.');
    if (!dni) throw new Error('El DNI es obligatorio.');
    if (!phone) throw new Error('El teléfono es obligatorio.');
    { const phV = vPhone_(phone); if (!phV.ok) throw new Error(phV.msg); phone = phV.norm; } // V-06: teléfono argentino válido
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

/* ===================== SOLICITUD DE CONGELAMIENTO (formulario web) ===================== */
/** Formulario público para pedir el congelamiento de un préstamo vencido (?page=congelar&loan=…).
 *  Mismo tema e interacción que la página de firma (barra de progreso, pantalla de éxito). */
function freezeRequestFormHtml_(loanId) {
  const lid = esc_(String(loanId || '').trim());
  const pctInt = Math.round(freezeDownPct_() * 100);
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><meta charset="utf-8"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#1c4587;margin:6px 0}
    label{font-weight:bold;display:block;margin:12px 0 4px}
    input[type=text],input[type=number],input[type=date],textarea{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px;font-size:14px}
    button{margin-top:10px;background:#1c4587;color:#fff;border:0;border-radius:6px;padding:12px 18px;font-weight:bold;cursor:pointer;width:100%}
    button:disabled{opacity:.5;cursor:not-allowed}
    .terms{width:100%;border-collapse:collapse;margin:8px 0}
    .terms td{padding:6px 8px;border:1px solid #ddd;font-size:14px}
    .terms td.k{background:#f0f4fa;font-weight:bold;width:58%}
    .terms td.big{font-size:15px;font-weight:bold;color:#0b2f52}
    .warn{background:#fff2cc;border-left:4px solid #bf9000;padding:10px 12px;border-radius:6px;margin:12px 0;font-size:13px}
    .consent{display:flex;gap:8px;align-items:flex-start;margin:10px 0;font-size:13px}
    .consent input{transform:scale(1.25);margin-top:3px}
    #gerr,#out{margin-top:12px;padding:12px;border-radius:6px;display:none}.err{background:#f4cccc}.ok{background:#e6f4ea}
    #loadingView{text-align:center;padding:34px 12px}
    .pbar-track{width:100%;height:24px;background:#e3e9f4;border-radius:12px;overflow:hidden;margin:14px 0 8px}
    .pbar-fill{height:100%;width:0%;border-radius:12px;transition:width .45s ease;background-color:#1c4587;
      background-image:linear-gradient(45deg,rgba(255,255,255,.30) 25%,transparent 25%,transparent 50%,rgba(255,255,255,.30) 50%,rgba(255,255,255,.30) 75%,transparent 75%,transparent);
      background-size:28px 28px;animation:pbar-stripes 1s linear infinite}
    @keyframes pbar-stripes{from{background-position:0 0}to{background-position:28px 0}}
    .check{width:72px;height:72px;line-height:72px;margin:6px auto 10px;border-radius:50%;background:#38761d;color:#fff;font-size:40px;font-weight:bold;text-align:center}
    .sub{color:#555;margin:0 0 8px;line-height:1.45}.hint{font-size:12px;color:#888}
    .hide{display:none}</style></head><body>
    <div class="card">${brandHeaderHtml_()}<h1>Solicitud de congelamiento</h1>

    <div id="gate">
      <p class="hint">Si tu préstamo está <b>vencido</b>, podés pedir <b>congelar la mora</b> (que deje de crecer) mientras acordás un pago. Ingresá tu DNI para ver tu préstamo.</p>
      <label>ID Préstamo</label><input type="text" id="loan" value="${lid}" placeholder="L-0001">
      <label>DNI</label><input type="text" id="dni" inputmode="numeric" placeholder="Tu DNI">
      <button onclick="look()">Ver mi préstamo</button>
      <div id="gerr"></div>
    </div>

    <div id="planview" class="hide">
      <div id="det"></div>
      <div class="warn">⚠ <b>Importante:</b> para congelar, primero debe estar <b>pagado el ${pctInt}%</b> del préstamo (dentro de las 24 h). Además, si congelás, es muy probable que un préstamo futuro, de otorgarse, sea por un <b>monto mucho menor</b>.</div>
      <label>¿Ya pagaste algo y no figura? Indicá el monto (lo verificaremos)</label>
      <input type="number" id="paid" min="0" step="0.01" value="0" placeholder="0">
      <table class="terms">
        <tr><td class="k">Pagado hasta ahora (registrado)</td><td id="paidRec">—</td></tr>
        <tr><td class="k">A pagar dentro de 24 h (${pctInt}% del préstamo)</td><td class="big" id="due24">—</td></tr>
        <tr><td class="k">Saldo restante (antes de la fecha que elijas)</td><td class="big" id="rem">—</td></tr>
      </table>
      <label>Fecha de pago del saldo restante (máximo 3 semanas)</label><input type="date" id="pay">
      <label>Motivo (opcional)</label><textarea id="reason" rows="2" placeholder="Contanos brevemente tu situación"></textarea>
      <label class="consent"><input type="checkbox" id="c1"><span>Me comprometo a pagar el <b>${pctInt}% dentro de las 24 horas</b> (antes de que se apruebe el congelamiento).</span></label>
      <label class="consent"><input type="checkbox" id="c2"><span>Me comprometo a pagar el <b>saldo restante</b> antes de la fecha que elijo (máximo 3 semanas).</span></label>
      <label class="consent"><input type="checkbox" id="c3"><span>Entiendo que un futuro préstamo, de haberlo, será por un <b>monto mucho menor</b>.</span></label>
      <button id="btnSend" onclick="send()" disabled>Enviar solicitud</button>
      <div id="out"></div>
    </div>

    <div id="loadingView" class="hide">
      <p id="loadPhase" style="font-weight:bold;color:#1c4587;margin:0;font-size:16px">Enviando tu solicitud…</p>
      <div class="pbar-track"><div class="pbar-fill" id="pbarFill"></div></div>
      <p id="loadPct" style="color:#1c4587;font-weight:bold;margin:0;font-size:20px">0%</p>
      <p class="hint" style="margin-top:6px">No cierres esta ventana.</p>
    </div>

    <div id="done" class="hide" style="text-align:center;padding:8px 0"></div>

    <script>
      var CLIENT_IP='', DEVICE={loc:{}}, PAGE_OPEN=Date.now(), LOANDATA=null;
      function $(id){return document.getElementById(id);}
      function money(n){ n=Math.round((Number(n)||0)*100)/100; try{return '$'+n.toLocaleString('es-AR',{minimumFractionDigits:2,maximumFractionDigits:2});}catch(e){return '$'+n.toFixed(2);} }
      function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
      (function(){ var t=new Date(); t.setHours(0,0,0,0); var mx=new Date(t.getTime()+21*864e5);
        function iso(d){return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);}
        $('pay').min=iso(t); $('pay').max=iso(mx); })();
      try{ fetch('https://api.ipify.org?format=json').then(function(r){return r.json();}).then(function(j){CLIENT_IP=(j&&j.ip)||'';}).catch(function(){});}catch(e){}
      (function(){ try{ var n=navigator, d=DEVICE;
        d.userAgent=n.userAgent||''; d.os=n.platform||''; d.langs=(n.languages&&n.languages.join)?n.languages.join(','):(n.language||'');
        d.screen=(screen.width+'x'+screen.height); d.dpr=window.devicePixelRatio||''; d.color=screen.colorDepth||'';
        d.touch=n.maxTouchPoints||0; d.cores=n.hardwareConcurrency||''; d.memory=n.deviceMemory||'';
        try{d.tz=Intl.DateTimeFormat().resolvedOptions().timeZone||'';}catch(e){d.tz='';} d.tzOffset=(new Date()).getTimezoneOffset();
        d.net=(n.connection&&n.connection.effectiveType)||''; d.webdriver=!!n.webdriver;
      }catch(e){}
        try{ fetch('https://ipwho.is/').then(function(r){return r.json();}).then(function(j){ if(j&&j.success!==false){ DEVICE.loc.ipText=[j.city,j.region,j.country].filter(Boolean).join(', '); if(!DEVICE.loc.source)DEVICE.loc.source='IP'; } }).catch(function(){}); }catch(e){}
        try{ if(navigator.geolocation) navigator.geolocation.getCurrentPosition(function(p){ DEVICE.loc.gps=p.coords.latitude.toFixed(5)+','+p.coords.longitude.toFixed(5); DEVICE.loc.gpsAcc=Math.round(p.coords.accuracy); DEVICE.loc.source='GPS'; }, function(){}, {timeout:5000,maximumAge:600000}); }catch(e){}
      })();
      function gerr(m){ var g=$('gerr'); g.className='err'; g.textContent=m; g.style.display='block'; }
      function out(m,err){ var o=$('out'); o.className=err?'err':'ok'; o.textContent=m; o.style.display='block'; }
      ['c1','c2','c3'].forEach(function(id){ $(id).onchange=revalidate; });
      $('pay').onchange=revalidate; $('paid').oninput=recalc;
      function recalc(){
        if(!LOANDATA) return;
        var P=Math.max(0, Number($('paid').value)||0);
        var due=Math.max(0, Math.round((LOANDATA.downShort - P)*100)/100);
        $('paidRec').textContent=money(LOANDATA.recorded);
        $('due24').textContent = due<=0 ? '✅ ya cubierto' : money(due);
        $('rem').textContent=money(LOANDATA.remaining);
      }
      function revalidate(){ $('btnSend').disabled=!(LOANDATA && $('c1').checked && $('c2').checked && $('c3').checked && $('pay').value); }
      function look(){
        $('gerr').style.display='none';
        if(!$('loan').value || !$('dni').value){ gerr('Ingresá el ID del préstamo y tu DNI.'); return; }
        $('gate').classList.add('hide'); $('loadingView').classList.remove('hide');
        startProgress([[25,'Verificando tu identidad…'],[65,'Calculando tu saldo y la mora…']]);
        google.script.run.withSuccessHandler(function(r){
          if(!r||!r.ok){ stopProgress(); $('loadingView').classList.add('hide'); $('gate').classList.remove('hide'); gerr((r&&r.reason)||'No se encontró el préstamo.'); return; }
          LOANDATA=r;
          $('det').innerHTML='<table class="terms">'+
            '<tr><td class="k">Préstamo</td><td>'+esc(r.name)+' — '+esc(r.loanId)+'</td></tr>'+
            '<tr><td class="k">Capital</td><td>'+money(r.principal)+'</td></tr>'+
            '<tr><td class="k">Vencimiento</td><td>'+esc(r.dueDateFmt)+' ('+r.daysLate+' día(s) de atraso)</td></tr>'+
            '<tr><td class="k">Recargo por mora</td><td>'+money(r.mora)+'</td></tr>'+
            '<tr><td class="k">Saldo con mora (hoy)</td><td class="big">'+money(r.balance)+'</td></tr></table>';
          finishProgress(function(){ $('loadingView').classList.add('hide'); $('planview').classList.remove('hide'); recalc(); revalidate(); });
        }).withFailureHandler(function(e){ stopProgress(); $('loadingView').classList.add('hide'); $('gate').classList.remove('hide'); gerr((e&&e.message)||e); }).lookupFreezeLoan($('loan').value,$('dni').value);
      }
      /* Barra de progreso (idéntica a la firma del contrato). */
      var progTimer=null, progVal=0;
      function setProgress(p,ph){ progVal=Math.max(0,Math.min(100,Math.round(p))); var f=$('pbarFill'),pc=$('loadPct'),lp=$('loadPhase'); if(f)f.style.width=progVal+'%'; if(pc)pc.textContent=progVal+'%'; if(ph&&lp)lp.textContent=ph; }
      function startProgress(ph){ ph = ph || [[12,'Validando tu préstamo…'],[48,'Registrando tu solicitud…'],[80,'Avisando al prestamista…']]; var i=1; setProgress(ph[0][0],ph[0][1]); clearInterval(progTimer);
        progTimer=setInterval(function(){ if(i<ph.length){ setProgress(ph[i][0],ph[i][1]); i++; } else if(progVal<95){ setProgress(progVal+1); } },650); }
      function stopProgress(){ clearInterval(progTimer); progTimer=null; }
      function finishProgress(cb){ stopProgress(); setProgress(100,'¡Listo!'); setTimeout(cb,400); }
      function showDone(m){ var d=$('done'); d.classList.remove('hide');
        d.innerHTML='<div class="check">✓</div><h1 style="color:#38761d">¡Solicitud enviada!</h1>'+
          '<p class="sub">'+esc(m||'Revisaremos tu pedido y te contactaremos.')+'</p>'+
          '<p class="hint">Ya podés cerrar esta ventana.</p>'; }
      function send(){
        $('out').style.display='none'; $('btnSend').disabled=true;
        $('planview').classList.add('hide'); $('loadingView').classList.remove('hide'); startProgress();
        var payload={ loanId:$('loan').value, dni:$('dni').value, promisedDate:$('pay').value,
          alreadyPaid:Number($('paid').value)||0, reason:$('reason').value, agree:true,
          ip:CLIENT_IP, device:DEVICE, dwellMs:Date.now()-PAGE_OPEN };
        google.script.run.withSuccessHandler(function(r){
          if(r&&r.ok){ finishProgress(function(){ $('loadingView').classList.add('hide'); showDone(r.message); }); }
          else { stopProgress(); $('loadingView').classList.add('hide'); $('planview').classList.remove('hide'); $('btnSend').disabled=false; out((r&&r.reason)||'No se pudo enviar.', true); }
        }).withFailureHandler(function(e){ stopProgress(); $('loadingView').classList.add('hide'); $('planview').classList.remove('hide'); $('btnSend').disabled=false; out((e&&e.message)||e, true); }).submitFreezeRequest(payload);
      }
    </script></div></body></html>`;
}

/** Verifica identidad (ID + DNI) y que el préstamo sea congelable; devuelve sus datos NUMÉRICOS para el formulario. */
function lookupFreezeLoan(loanId, dni) {
  return guard_('lookupFreezeLoan', function () {
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, reason: 'Datos no encontrados. Verificá el ID del préstamo y tu DNI.' };
    const fz = loanFreezable_(loan);
    if (!fz.ok) return { ok: false, reason: fz.reason };
    const a = freezeAmounts_(loan, 0);
    return {
      ok: true, name: loan.name, loanId: loan.loanId, principal: round2_(loan.principal),
      dueDateFmt: fmtDate_(loan.dueDate), daysLate: daysLate_(loan.dueDate, new Date()),
      mora: a.mora, balance: a.balance, total: a.total, required: a.required,
      recorded: a.recorded, downShort: a.downShort, remaining: a.remaining, pctInt: a.pctInt,
    };
  });
}

/** ¿Ya hay una solicitud de congelamiento Pendiente o Aprobada para este préstamo? (evita duplicados). */
function hasPendingFreezeRequest_(sh, loanId) {
  const last = sh.getLastRow(); if (last < 2) return false;
  const H = headerMap_(sh), idC = H['ID Préstamo'] || 2, stC = H['Estado'] || 23;
  const data = sh.getRange(2, 1, last - 1, Math.max(idC, stC)).getValues(), id = String(loanId).trim();
  for (let i = 0; i < data.length; i++)
    if (String(data[i][idC - 1]).trim() === id) { const st = String(data[i][stC - 1]).trim(); if (st === 'Pendiente' || st === 'Aprobado') return true; }
  return false;
}

/** Recibe la solicitud del formulario: valida, guarda en "Solicitudes de Congelamiento" y avisa a ambas partes. */
function submitFreezeRequest(p) {
  return guard_('submitFreezeRequest', function () {
    p = p || {};
    const loan = findLoanById_(String(p.loanId || '').trim());
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(p.dni || '').replace(/\D/g, ''))
      return { ok: false, reason: 'Datos no encontrados. Verificá el ID del préstamo y tu DNI.' };
    const fz = loanFreezable_(loan);
    if (!fz.ok) return { ok: false, reason: fz.reason };
    if (p.agree !== true) return { ok: false, reason: 'Debés aceptar las condiciones para continuar.' };
    const promised = parseFreezeDate_(String(p.promisedDate || '').trim());
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const max = new Date(today.getTime() + 21 * 86400000);
    if (!promised || promised.getTime() < today.getTime() || promised.getTime() > max.getTime())
      return { ok: false, reason: 'Elegí una fecha de pago válida, dentro de las próximas 3 semanas.' };
    const ss = getSS_();
    const sh = ss.getSheetByName(CFG.SHEETS.FREEZE_REQ) || setupFreezeRequests_(ss);
    if (hasPendingFreezeRequest_(sh, loan.loanId))
      return { ok: false, reason: 'Ya hay una solicitud de congelamiento en curso para este préstamo.' };
    const a = freezeAmounts_(loan, p.alreadyPaid);
    const loc = (p.device && p.device.loc) || {};
    const ubic = loc.gps ? ('GPS ' + loc.gps) : (loc.ipText || '');
    // Orden EXACTO de FREEZE_REQ_HEADERS (29 col). La col 12 ("Falta para el 50%") es una FÓRMULA viva (se escribe aparte).
    const rowVals = [new Date(), loan.loanId, loan.clientId, loan.name, loan.dni, loan.email, loan.phone,
      loan.principal, a.balance, a.mora, a.P, '', a.remaining, loan.dueDate, promised,
      String(p.reason || '').slice(0, 500), ubic, String(p.ip || ''), String((p.device && p.device.userAgent) || '').slice(0, 250),
      false, false, '', 'Pendiente', '', '', '', 'PENDIENTE', '', ''];
    const row = nbAppend_(sh, rowVals);
    sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, 8, 1, 4).setNumberFormat(CFG.CURRENCY_FMT);   // Capital..Ya pagado
    sh.getRange(row, 13, 1, 1).setNumberFormat(CFG.CURRENCY_FMT);  // Saldo restante
    sh.getRange(row, 14, 1, 2).setNumberFormat('yyyy-mm-dd');
    sh.getRange(row, FR_COVER_COL).setFormula(freezeCoverFormula_(row)); // "Falta para el 50% (24 h)" viva
    try { logDeviceInfo_(loan, loan.clientId, { device: p.device || {}, ip: p.ip || '', userAgent: (p.device && p.device.userAgent) || '', dwellMs: p.dwellMs, signedAt: new Date() }); } catch (e) { logError_('submitFreezeRequest:device', e); }
    const lender = getSetting_('Correo del Prestamista');
    const amtRows = '<tr><td><b>A pagar en 24 h (' + a.pctInt + '%)</b>&nbsp;</td><td>' + fmtMoney_(a.downShort) + '</td></tr>' +
      '<tr><td><b>Saldo restante</b>&nbsp;</td><td>' + fmtMoney_(a.remaining) + '</td></tr>' +
      (a.P > 0 ? '<tr><td><b>Ya pagado (declarado)</b>&nbsp;</td><td>' + fmtMoney_(a.P) + '</td></tr>' : '');
    try {
      if (loan.email) sendBrandedEmail_(loan.email, 'Recibimos tu solicitud de congelamiento — ' + loan.loanId,
        'Hola ' + loan.name + ', recibimos tu solicitud de congelamiento del préstamo ' + loan.loanId + '. Te avisaremos cuando la revisemos.',
        '<p>Hola ' + esc_(loan.name) + ',</p><p>Recibimos tu <b>solicitud de congelamiento</b> del préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
        '<table style="border-collapse:collapse">' +
        '<tr><td><b>Saldo con mora</b>&nbsp;</td><td>' + fmtMoney_(a.balance) + '</td></tr>' + amtRows +
        '<tr><td><b>Fecha de pago del saldo</b>&nbsp;</td><td>' + fmtDate_(promised) + '</td></tr></table>' +
        '<p>Te avisaremos cuando la revisemos.</p>');
      if (lender) sendBrandedEmail_(lender, 'Nueva solicitud de congelamiento — ' + loan.loanId,
        'Nueva solicitud de congelamiento de ' + loan.name + ' (' + loan.loanId + '). Revisala en «Solicitudes de Congelamiento».',
        '<p>Nueva <b>solicitud de congelamiento</b>:</p><table style="border-collapse:collapse">' +
        '<tr><td><b>Préstamo</b>&nbsp;</td><td>' + esc_(loan.loanId) + '</td></tr>' +
        '<tr><td><b>Cliente</b>&nbsp;</td><td>' + esc_(loan.name) + ' (DNI ' + esc_(loan.dni) + ')</td></tr>' +
        '<tr><td><b>Saldo con mora</b>&nbsp;</td><td>' + fmtMoney_(a.balance) + '</td></tr>' + amtRows +
        '<tr><td><b>Fecha de pago del saldo</b>&nbsp;</td><td>' + fmtDate_(promised) + '</td></tr></table>' +
        '<p>Aprobala o rechazala en la hoja «Solicitudes de Congelamiento».</p>');
    } catch (e) { logError_('submitFreezeRequest:email', e); }
    return { ok: true, message: '¡Solicitud enviada! Revisaremos tu pedido y te contactaremos.' };
  });
}

/* ===================== SOLICITUD DE PLAN DE PAGO (formulario web) ===================== */
/** Formulario público para pedir un plan de pago (Acuerdo) de un préstamo al 100% de mora
 *  (?page=plan&loan=…). Mismo tema/interacción que el formulario de congelamiento. */
function planRequestFormHtml_(loanId) {
  const lid = esc_(String(loanId || '').trim());
  const pctInt = Math.round(planDownPct_() * 100);
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><meta charset="utf-8"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#674ea7;margin:6px 0}
    label{font-weight:bold;display:block;margin:12px 0 4px}
    input[type=text],input[type=number],textarea{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px;font-size:14px}
    button{margin-top:10px;background:#674ea7;color:#fff;border:0;border-radius:6px;padding:12px 18px;font-weight:bold;cursor:pointer;width:100%}
    button:disabled{opacity:.5;cursor:not-allowed}
    .terms{width:100%;border-collapse:collapse;margin:8px 0}
    .terms td{padding:6px 8px;border:1px solid #ddd;font-size:14px}
    .terms td.k{background:#f3f0fa;font-weight:bold;width:58%}
    .terms td.big{font-size:15px;font-weight:bold;color:#351c75}
    .opt{display:flex;gap:8px;align-items:flex-start;margin:8px 0;font-weight:normal;font-size:14px;border:1px solid #d8d0ec;border-radius:6px;padding:8px 10px}
    .opt input{transform:scale(1.2);margin-top:3px}
    .warn{background:#fff2cc;border-left:4px solid #bf9000;padding:10px 12px;border-radius:6px;margin:12px 0;font-size:13px}
    .consent{display:flex;gap:8px;align-items:flex-start;margin:10px 0;font-size:13px}
    .consent input{transform:scale(1.25);margin-top:3px}
    #gerr,#out{margin-top:12px;padding:12px;border-radius:6px;display:none}.err{background:#f4cccc}.ok{background:#e6f4ea}
    #loadingView{text-align:center;padding:34px 12px}
    .pbar-track{width:100%;height:24px;background:#ece7f6;border-radius:12px;overflow:hidden;margin:14px 0 8px}
    .pbar-fill{height:100%;width:0%;border-radius:12px;transition:width .45s ease;background-color:#674ea7;
      background-image:linear-gradient(45deg,rgba(255,255,255,.30) 25%,transparent 25%,transparent 50%,rgba(255,255,255,.30) 50%,rgba(255,255,255,.30) 75%,transparent 75%,transparent);
      background-size:28px 28px;animation:pbar-stripes 1s linear infinite}
    @keyframes pbar-stripes{from{background-position:0 0}to{background-position:28px 0}}
    .check{width:72px;height:72px;line-height:72px;margin:6px auto 10px;border-radius:50%;background:#38761d;color:#fff;font-size:40px;font-weight:bold;text-align:center}
    .sub{color:#555;margin:0 0 8px;line-height:1.45}.hint{font-size:12px;color:#888}
    .hide{display:none}</style></head><body>
    <div class="card">${brandHeaderHtml_()}<h1>Solicitud de plan de pago</h1>

    <div id="gate">
      <p class="hint">Si estás <b>atrasado/a</b> y no podés pagar todo junto, podés acordar un <b>plan de pago en cuotas</b> (semanal, quincenal o mensual) y frenar la mora. Ingresá tu DNI para ver tu préstamo.</p>
      <label>ID Préstamo</label><input type="text" id="loan" value="${lid}" placeholder="L-0001">
      <label>DNI</label><input type="text" id="dni" inputmode="numeric" placeholder="Tu DNI">
      <button onclick="look()">Ver mi préstamo</button>
      <div id="gerr"></div>
    </div>

    <div id="planview" class="hide">
      <div id="det"></div>
      <div class="warn">⚠ <b>Importante:</b> para activar el plan, primero debe estar <b>pagado el ${pctInt}%</b> del préstamo (dentro de las 24 h). La mora acumulada queda <b>suspendida</b> mientras cumplas el plan; si incumplís, se restablece.</div>
      <label>Frecuencia de pago</label>
      <select id="cad" onchange="pickOpt()" style="width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px;font-size:14px">
        <option value="monthly">Mensual (una cuota por mes)</option>
        <option value="biweekly">Quincenal (cada 15 días)</option>
        <option value="weekly">Semanal (una cuota por semana)</option>
      </select>
      <label>Elegí una opción de pago</label>
      <label class="opt"><input type="radio" name="opt" value="1" onchange="pickOpt()"><span><b>Pago único</b> — todo el saldo en un solo pago.</span></label>
      <label class="opt"><input type="radio" name="opt" value="4" onchange="pickOpt()"><span><b>4 cuotas</b></span></label>
      <label class="opt"><input type="radio" name="opt" value="6" onchange="pickOpt()"><span><b>6 cuotas</b></span></label>
      <table class="terms" id="sched"></table>
      <table class="terms">
        <tr><td class="k">Pagado hasta ahora (registrado)</td><td id="paidRec">—</td></tr>
        <tr><td class="k">A pagar dentro de 24 h (${pctInt}% del préstamo)</td><td class="big" id="due24">—</td></tr>
      </table>
      <label>Motivo (opcional)</label><textarea id="reason" rows="2" placeholder="Contanos brevemente tu situación"></textarea>
      <label class="consent"><input type="checkbox" id="c1"><span>Me comprometo a pagar el <b>${pctInt}% dentro de las 24 horas</b> para activar el plan.</span></label>
      <label class="consent"><input type="checkbox" id="c2"><span>Me comprometo a pagar las <b>cuotas en fecha</b>. Reconozco la deuda.</span></label>
      <label class="consent"><input type="checkbox" id="c3"><span>Entiendo que si incumplo un pago, la <b>mora se restablece</b> y el saldo total se vuelve exigible.</span></label>
      <button id="btnSend" onclick="send()" disabled>Enviar solicitud</button>
      <div id="out"></div>
    </div>

    <div id="loadingView" class="hide">
      <p id="loadPhase" style="font-weight:bold;color:#674ea7;margin:0;font-size:16px">Enviando tu solicitud…</p>
      <div class="pbar-track"><div class="pbar-fill" id="pbarFill"></div></div>
      <p id="loadPct" style="color:#674ea7;font-weight:bold;margin:0;font-size:20px">0%</p>
      <p class="hint" style="margin-top:6px">No cierres esta ventana.</p>
    </div>

    <div id="done" class="hide" style="text-align:center;padding:8px 0"></div>

    <script>
      var CLIENT_IP='', DEVICE={loc:{}}, PAGE_OPEN=Date.now(), LOANDATA=null, OPT=0;
      function $(id){return document.getElementById(id);}
      function money(n){ n=Math.round((Number(n)||0)*100)/100; try{return '$'+n.toLocaleString('es-AR',{minimumFractionDigits:2,maximumFractionDigits:2});}catch(e){return '$'+n.toFixed(2);} }
      function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
      var MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
      function fmtD(d){ return d.getDate()+' '+MESES[d.getMonth()]+' '+d.getFullYear(); }
      try{ fetch('https://api.ipify.org?format=json').then(function(r){return r.json();}).then(function(j){CLIENT_IP=(j&&j.ip)||'';}).catch(function(){});}catch(e){}
      (function(){ try{ var n=navigator, d=DEVICE;
        d.userAgent=n.userAgent||''; d.os=n.platform||''; d.langs=(n.languages&&n.languages.join)?n.languages.join(','):(n.language||'');
        d.screen=(screen.width+'x'+screen.height); d.dpr=window.devicePixelRatio||''; d.color=screen.colorDepth||'';
        d.touch=n.maxTouchPoints||0; d.cores=n.hardwareConcurrency||''; d.memory=n.deviceMemory||'';
        try{d.tz=Intl.DateTimeFormat().resolvedOptions().timeZone||'';}catch(e){d.tz='';} d.tzOffset=(new Date()).getTimezoneOffset();
        d.net=(n.connection&&n.connection.effectiveType)||''; d.webdriver=!!n.webdriver;
      }catch(e){}
        try{ fetch('https://ipwho.is/').then(function(r){return r.json();}).then(function(j){ if(j&&j.success!==false){ DEVICE.loc.ipText=[j.city,j.region,j.country].filter(Boolean).join(', '); if(!DEVICE.loc.source)DEVICE.loc.source='IP'; } }).catch(function(){}); }catch(e){}
        try{ if(navigator.geolocation) navigator.geolocation.getCurrentPosition(function(p){ DEVICE.loc.gps=p.coords.latitude.toFixed(5)+','+p.coords.longitude.toFixed(5); DEVICE.loc.gpsAcc=Math.round(p.coords.accuracy); DEVICE.loc.source='GPS'; }, function(){}, {timeout:5000,maximumAge:600000}); }catch(e){}
      })();
      function gerr(m){ var g=$('gerr'); g.className='err'; g.textContent=m; g.style.display='block'; }
      function out(m,err){ var o=$('out'); o.className=err?'err':'ok'; o.textContent=m; o.style.display='block'; }
      ['c1','c2','c3'].forEach(function(id){ $(id).onchange=revalidate; });
      function pickOpt(){ var r=document.querySelector('input[name=opt]:checked'); OPT=r?Number(r.value):0; drawSched(); revalidate(); }
      function schedDates(n){ var today=new Date(); today.setHours(0,0,0,0); var cad=$('cad')?$('cad').value:'monthly'; var out=[];
        if(n<=1){ var d=new Date(today.getTime()+(cad==='weekly'?7:14)*864e5); return [d]; }
        for(var k=1;k<=n;k++){ var d;
          if(cad==='weekly') d=new Date(today.getTime()+7*k*864e5);
          else if(cad==='biweekly') d=new Date(today.getTime()+14*k*864e5);
          else d=new Date(today.getFullYear(), today.getMonth()+k, today.getDate());
          out.push(d); } return out; }
      function drawSched(){
        if(!LOANDATA || !OPT){ $('sched').innerHTML=''; return; }
        var n=OPT, base=Math.round((LOANDATA.total/n)*100)/100, paid=LOANDATA.recorded, dates=schedDates(n);
        var html='<tr><td class="k">Cuota</td><td class="k">Vencimiento (aprox.)</td><td class="k">Monto</td></tr>';
        var left=paid;
        for(var k=0;k<n;k++){
          var monto=(k===n-1)?Math.round((LOANDATA.total-base*(n-1))*100)/100:base;
          var ap=Math.min(left,monto); var neto=Math.round((monto-ap)*100)/100; left=Math.round((left-ap)*100)/100;
          var et=(n===1)?'Pago único':(k+1)+' de '+n;
          html+='<tr><td>'+et+'</td><td>'+fmtD(dates[k])+'</td><td>'+money(neto)+(ap>0?' <span style="color:#888">(−'+money(ap)+' ya pagado)</span>':'')+'</td></tr>';
        }
        $('sched').innerHTML=html;
      }
      function revalidate(){ $('btnSend').disabled=!(LOANDATA && OPT && $('c1').checked && $('c2').checked && $('c3').checked); }
      function look(){
        $('gerr').style.display='none';
        if(!$('loan').value || !$('dni').value){ gerr('Ingresá el ID del préstamo y tu DNI.'); return; }
        $('gate').classList.add('hide'); $('loadingView').classList.remove('hide');
        startProgress([[25,'Verificando tu identidad…'],[65,'Calculando tu saldo y la mora…']]);
        google.script.run.withSuccessHandler(function(r){
          if(!r||!r.ok){ stopProgress(); $('loadingView').classList.add('hide'); $('gate').classList.remove('hide'); gerr((r&&r.reason)||'No se encontró el préstamo.'); return; }
          LOANDATA=r;
          $('det').innerHTML='<table class="terms">'+
            '<tr><td class="k">Préstamo</td><td>'+esc(r.name)+' — '+esc(r.loanId)+'</td></tr>'+
            '<tr><td class="k">Total a pagar (capital + interés)</td><td>'+money(r.total)+'</td></tr>'+
            '<tr><td class="k">Mora acumulada (se suspende)</td><td>'+money(r.mora)+'</td></tr>'+
            '<tr><td class="k">Saldo base pendiente (sin mora)</td><td class="big">'+money(r.base)+'</td></tr></table>';
          $('paidRec').textContent=money(r.recorded);
          var due=Math.max(0, Math.round((r.downShort)*100)/100);
          $('due24').textContent = due<=0 ? '✅ ya cubierto' : money(due);
          finishProgress(function(){ $('loadingView').classList.add('hide'); $('planview').classList.remove('hide'); revalidate(); });
        }).withFailureHandler(function(e){ stopProgress(); $('loadingView').classList.add('hide'); $('gate').classList.remove('hide'); gerr((e&&e.message)||e); }).lookupPlanLoan($('loan').value,$('dni').value);
      }
      var progTimer=null, progVal=0;
      function setProgress(p,ph){ progVal=Math.max(0,Math.min(100,Math.round(p))); var f=$('pbarFill'),pc=$('loadPct'),lp=$('loadPhase'); if(f)f.style.width=progVal+'%'; if(pc)pc.textContent=progVal+'%'; if(ph&&lp)lp.textContent=ph; }
      function startProgress(ph){ ph = ph || [[12,'Validando…'],[48,'Registrando tu solicitud…'],[80,'Avisando al prestamista…']]; var i=1; setProgress(ph[0][0],ph[0][1]); clearInterval(progTimer);
        progTimer=setInterval(function(){ if(i<ph.length){ setProgress(ph[i][0],ph[i][1]); i++; } else if(progVal<95){ setProgress(progVal+1); } },650); }
      function stopProgress(){ clearInterval(progTimer); progTimer=null; }
      function finishProgress(cb){ stopProgress(); setProgress(100,'¡Listo!'); setTimeout(cb,400); }
      function showDone(m){ var d=$('done'); d.classList.remove('hide');
        d.innerHTML='<div class="check">✓</div><h1 style="color:#38761d">¡Solicitud enviada!</h1>'+
          '<p class="sub">'+esc(m||'Revisaremos tu pedido y te contactaremos.')+'</p>'+
          '<p class="hint">Ya podés cerrar esta ventana.</p>'; }
      function send(){
        $('out').style.display='none'; $('btnSend').disabled=true;
        $('planview').classList.add('hide'); $('loadingView').classList.remove('hide'); startProgress();
        var payload={ loanId:$('loan').value, dni:$('dni').value, option:OPT, cadence:($('cad')?$('cad').value:'monthly'), reason:$('reason').value, agree:true,
          ip:CLIENT_IP, device:DEVICE, dwellMs:Date.now()-PAGE_OPEN };
        google.script.run.withSuccessHandler(function(r){
          if(r&&r.ok){ finishProgress(function(){ $('loadingView').classList.add('hide'); showDone(r.message); }); }
          else { stopProgress(); $('loadingView').classList.add('hide'); $('planview').classList.remove('hide'); $('btnSend').disabled=false; out((r&&r.reason)||'No se pudo enviar.', true); }
        }).withFailureHandler(function(e){ stopProgress(); $('loadingView').classList.add('hide'); $('planview').classList.remove('hide'); $('btnSend').disabled=false; out((e&&e.message)||e, true); }).submitPlanRequest(payload);
      }
    </script></div></body></html>`;
}

/** Verifica identidad (ID+DNI) y que el préstamo esté VENCIDO con saldo; devuelve datos para el plan de pago flexible. */
function lookupPlanLoan(loanId, dni) {
  return guard_('lookupPlanLoan', function () {
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, reason: 'Datos no encontrados. Verificá el ID del préstamo y tu DNI.' };
    // Plan de pago FLEXIBLE: disponible para cualquier préstamo VENCIDO con saldo (no requiere el
    // 100% de mora). El prestamista aprueba cada solicitud.
    if (!(loan.dueDate instanceof Date) || daysLate_(loan.dueDate, new Date()) <= 0)
      return { ok: false, reason: 'El plan de pago está disponible para préstamos vencidos.' };
    const a = planAmounts_(loan);
    if (a.base <= 0.009) return { ok: false, reason: 'Este préstamo ya está saldado.' };
    return {
      ok: true, name: loan.name, loanId: loan.loanId, principal: round2_(loan.principal),
      dueDateFmt: fmtDate_(loan.dueDate), daysLate: daysLate_(loan.dueDate, new Date()),
      total: a.total, base: a.base, mora: a.mora, recorded: a.recorded, required: a.required,
      downShort: a.downShort, pctInt: a.pctInt,
    };
  });
}

/** ¿Ya hay una solicitud de plan Pendiente o Aprobada para este préstamo? (evita duplicados). */
function hasPendingPlanRequest_(sh, loanId) {
  const last = sh.getLastRow(); if (last < 2) return false;
  const H = headerMap_(sh), idC = H['ID Préstamo'] || 2, stC = H['Estado'] || 23;
  const data = sh.getRange(2, 1, last - 1, Math.max(idC, stC)).getValues(), id = String(loanId).trim();
  for (let i = 0; i < data.length; i++)
    if (String(data[i][idC - 1]).trim() === id) { const st = String(data[i][stC - 1]).trim(); if (st === 'Pendiente' || st === 'Aprobado') return true; }
  return false;
}

/** Recibe la solicitud de plan del formulario: valida, guarda en "Solicitudes de Plan de Pago", avisa. */
function submitPlanRequest(p) {
  return guard_('submitPlanRequest', function () {
    p = p || {};
    const loan = findLoanById_(String(p.loanId || '').trim());
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(p.dni || '').replace(/\D/g, ''))
      return { ok: false, reason: 'Datos no encontrados. Verificá el ID del préstamo y tu DNI.' };
    if (!(loan.dueDate instanceof Date) || daysLate_(loan.dueDate, new Date()) <= 0)
      return { ok: false, reason: 'El plan de pago está disponible para préstamos vencidos.' };
    if (Math.max(0, round2_(loan.totalDue - loanPayments_(loan.loanId).reduce((s, pp) => s + pp.amount, 0))) <= 0.009)
      return { ok: false, reason: 'Este préstamo ya está saldado.' };
    const option = Math.floor(Number(p.option) || 0);
    if ([1, 4, 6].indexOf(option) === -1) return { ok: false, reason: 'Elegí una opción de pago (único, 4 o 6 cuotas).' };
    const cadence = planCadence_(p.cadence);  // semanal / quincenal / mensual (default)
    if (p.agree !== true) return { ok: false, reason: 'Debés aceptar las condiciones para continuar.' };
    const ss = getSS_();
    let sh = ss.getSheetByName(CFG.SHEETS.PLAN_REQ);
    if (!sh) { sh = setupPlanRequests_(ss); }
    else if (headerMap_(sh)['Cadencia'] == null) {
      // Alta LIVIANA de la columna «Cadencia» (sheets viejos de 29 col): NO corre el setup pesado
      // (300 fórmulas + validaciones) durante el envío del formulario — eso colgaba la barra al 95%.
      if (sh.getMaxColumns() < PLAN_REQ_HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(), PLAN_REQ_HEADERS.length - sh.getMaxColumns());
      sh.getRange(1, PLAN_REQ_HEADERS.length).setValue('Cadencia').setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff').setWrap(true);
    }
    if (hasPendingPlanRequest_(sh, loan.loanId))
      return { ok: false, reason: 'Ya hay una solicitud de plan de pago en curso para este préstamo.' };
    const a = planAmounts_(loan);
    const loc = (p.device && p.device.loc) || {};
    const ubic = loc.gps ? ('GPS ' + loc.gps) : (loc.ipText || '');
    // Orden EXACTO de PLAN_REQ_HEADERS (30 col). Col 14 ("Falta para el 20%") es FÓRMULA viva; col 30 = Cadencia.
    const rowVals = [new Date(), loan.loanId, loan.clientId, loan.name, loan.dni, loan.email, loan.phone,
      loan.principal, a.total, a.base, a.mora, option, a.required, '', loan.dueDate,
      String(p.reason || '').slice(0, 500), ubic, String(p.ip || ''), String((p.device && p.device.userAgent) || '').slice(0, 250),
      false, false, '', 'Pendiente', '', '', '', 'PENDIENTE', '', '', cadence];
    const row = nbAppend_(sh, rowVals);
    sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, 8, 1, 4).setNumberFormat(CFG.CURRENCY_FMT);   // Capital..Mora suspendida
    sh.getRange(row, 13, 1, 1).setNumberFormat(CFG.CURRENCY_FMT);  // Pago inicial (20%)
    sh.getRange(row, 15, 1, 1).setNumberFormat('yyyy-mm-dd');      // Vencimiento
    sh.getRange(row, PR_COVER_COL).setFormula(planCoverFormula_(row));
    try { logDeviceInfo_(loan, loan.clientId, { device: p.device || {}, ip: p.ip || '', userAgent: (p.device && p.device.userAgent) || '', dwellMs: p.dwellMs, signedAt: new Date() }); } catch (e) { logError_('submitPlanRequest:device', e); }
    const optTxt = planOptTxt_(option, cadence);
    const lender = getSetting_('Correo del Prestamista');
    try {
      if (loan.email) sendBrandedEmail_(loan.email, 'Recibimos tu solicitud de plan de pago — ' + loan.loanId,
        'Hola ' + loan.name + ', recibimos tu solicitud de plan de pago (' + optTxt + ') del préstamo ' + loan.loanId + '.',
        '<p>Hola ' + esc_(loan.name) + ',</p><p>Recibimos tu <b>solicitud de plan de pago</b> (' + esc_(optTxt) + ') del préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
        '<table style="border-collapse:collapse"><tr><td><b>Total a pagar</b>&nbsp;</td><td>' + fmtMoney_(a.total) + '</td></tr>' +
        '<tr><td><b>Pago inicial (' + a.pctInt + '%)</b>&nbsp;</td><td>' + fmtMoney_(a.required) + '</td></tr></table>' +
        '<p>Pagá el ' + a.pctInt + '% dentro de las 24 h para activar el plan. Te avisaremos cuando lo revisemos.</p>');
      if (lender) sendBrandedEmail_(lender, 'Nueva solicitud de plan de pago — ' + loan.loanId,
        'Nueva solicitud de plan de ' + loan.name + ' (' + loan.loanId + ', ' + optTxt + '). Revisala en «Solicitudes de Plan de Pago».',
        '<p>Nueva <b>solicitud de plan de pago</b>:</p><table style="border-collapse:collapse">' +
        '<tr><td><b>Préstamo</b>&nbsp;</td><td>' + esc_(loan.loanId) + '</td></tr>' +
        '<tr><td><b>Cliente</b>&nbsp;</td><td>' + esc_(loan.name) + ' (DNI ' + esc_(loan.dni) + ')</td></tr>' +
        '<tr><td><b>Opción</b>&nbsp;</td><td>' + esc_(optTxt) + '</td></tr>' +
        '<tr><td><b>Total a pagar</b>&nbsp;</td><td>' + fmtMoney_(a.total) + '</td></tr>' +
        '<tr><td><b>Pago inicial (' + a.pctInt + '%)</b>&nbsp;</td><td>' + fmtMoney_(a.required) + '</td></tr></table>' +
        '<p>Aprobala o rechazala en la hoja «Solicitudes de Plan de Pago».</p>');
    } catch (e) { logError_('submitPlanRequest:email', e); }
    return { ok: true, message: '¡Solicitud enviada! Revisaremos tu pedido y te contactaremos.' };
  });
}

/** APRUEBA la solicitud de la fila: arma el enlace para firmar el nuevo contrato, el WhatsApp de
 *  aviso y avisa por correo a ambas partes. NO congela todavía: el préstamo se congela al FIRMAR. */
function approveFreezeRequest_(sh, row) {
  const m = headerMap_(sh);
  const g = h => m[h] ? sh.getRange(row, m[h]).getValue() : '';
  const setv = (h, v) => { if (m[h]) sh.getRange(row, m[h]).setValue(v); };
  const setf = (h, f) => { if (m[h]) sh.getRange(row, m[h]).setFormula(f); };
  const loanId = String(g('ID Préstamo')).trim();
  if (!loanId) return;
  const loan = findLoanById_(loanId);
  if (!loan) { setv('Estado', 'Rechazado'); setv('Motivo de rechazo', 'No se encontró el préstamo.'); return; }
  const fz = loanFreezable_(loan);
  if (!fz.ok) { getSS_().toast(loanId + ': ' + fz.reason + ' No se aprueba.', '⚠', 8); return; }
  // GATE: el % configurable debe estar PAGADO (registrado) antes de aprobar. Si falta, NO aprueba y
  // deja constancia VISIBLE (nota en la fila + toast persistente) — getUi().alert no corre en el onEdit.
  const amt = freezeAmounts_(loan);
  if (!amt.downMet) {
    const msg = 'Falta pagar el ' + amt.pctInt + '% (' + fmtMoney_(amt.required) + '): faltan ' + fmtMoney_(amt.downShort) +
      '. Registrá el pago en «Pagos» y volvé a tildar «Aprobar».';
    if (m['Aprobar ✔']) sh.getRange(row, m['Aprobar ✔']).setNote('⛔ ' + msg);
    if (m['Estado']) sh.getRange(row, m['Estado']).setNote('⛔ No aprobado — ' + msg);
    try { getSS_().toast('Préstamo ' + loanId + ' — ' + msg, '⛔ NO SE APROBÓ · Falta el pago inicial', -1); } catch (e) { }
    try { SpreadsheetApp.getUi().alert('No se aprobó el congelamiento ' + loanId + '.\n\n' + msg); } catch (e) { }
    return; // no se aprueba: queda Pendiente
  }
  if (m['Aprobar ✔']) sh.getRange(row, m['Aprobar ✔']).clearNote();   // cubierto: limpia la nota de bloqueo
  if (m['Estado']) sh.getRange(row, m['Estado']).clearNote();
  const promised = g('Fecha de pago prometida');
  setv('Estado', 'Aprobado');
  if (m['Fecha de aprobación']) sh.getRange(row, m['Fecha de aprobación']).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  setv('Estado de Firma', 'PENDIENTE');
  const base = webAppBaseUrl_();
  const url = base ? base + '?page=firmar-congelamiento&loan=' + encodeURIComponent(loanId) : '';
  if (url) setf('Enlace de firma', '=HYPERLINK("' + url + '","✍ Firmar congelamiento")');
  else setv('Enlace de firma', 'Publicá la app web (menú ⑥) para obtener el enlace.');
  // Importes RECALCULADOS a la fecha de aprobación (no se leen de la hoja, para no depender de
  // nombres de columna ni de valores viejos). `amt` ya se computó para el control del %.
  const fechaTxt = (promised instanceof Date) ? fmtDate_(promised) : 'la fecha comprometida';
  const num = waNumberAR_(loan.phone);
  if (num && url) {
    const msg = 'Hola ' + (loan.name || 'cliente') + ' 👋 Aprobamos tu congelamiento del préstamo ' + loanId +
      '. El ' + amt.pctInt + '% requerido ya está cubierto. Resta pagar el saldo (' + fmtMoney_(amt.remaining) + ') antes del ' + fechaTxt +
      '. Para que tenga efecto, firmá el nuevo contrato acá: ' + url + ' — ' + companyName_();
    setf('WhatsApp', '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📲 Avisar aprobación")');
  }
  const lender = getSetting_('Correo del Prestamista');
  const btn = url ? '<p><a href="' + url + '" style="background:#0b5394;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold">✍ Firmar el congelamiento</a></p>' : '';
  const promTxt = (promised instanceof Date) ? (' con fecha de pago comprometida <b>' + fmtDate_(promised) + '</b>') : '';
  const amtTable = '<table style="border-collapse:collapse;margin:6px 0">' +
    '<tr><td><b>Saldo con mora</b>&nbsp;</td><td>' + fmtMoney_(amt.balance) + '</td></tr>' +
    '<tr><td><b>Pago requerido (' + amt.pctInt + '%)</b>&nbsp;</td><td>' + fmtMoney_(amt.required) + '</td></tr>' +
    '<tr><td><b>Pagado (registrado)</b>&nbsp;</td><td>' + fmtMoney_(amt.recorded) + '</td></tr>' +
    '<tr><td><b>Saldo restante</b>&nbsp;</td><td>' + fmtMoney_(amt.remaining) + '</td></tr>' +
    (promised instanceof Date ? '<tr><td><b>Fecha de pago del saldo</b>&nbsp;</td><td>' + fmtDate_(promised) + '</td></tr>' : '') + '</table>';
  try {
    if (loan.email) sendBrandedEmail_(loan.email, 'Tu congelamiento fue aprobado — ' + loanId,
      'Hola ' + loan.name + ', aprobamos tu congelamiento del préstamo ' + loanId + '. Firmá el nuevo contrato: ' + (url || '(enlace pendiente)'),
      '<p>Hola ' + esc_(loan.name) + ',</p><p>Aprobamos tu pedido de <b>congelamiento</b> del préstamo <b>' + esc_(loanId) + '</b>' + promTxt + '.</p>' +
      amtTable +
      '<p>El <b>' + amt.pctInt + '% requerido ya está cubierto</b>. Resta pagar el <b>saldo restante</b> antes de la fecha comprometida.</p>' +
      '<p>Para que el congelamiento tenga efecto, <b>firmá el nuevo contrato</b> con los nuevos términos. El préstamo queda congelado al firmar.</p>' + btn);
    if (lender) sendBrandedEmail_(lender, 'Congelamiento aprobado — ' + loanId,
      'Aprobaste el congelamiento de ' + loan.name + ' (' + loanId + '). Se envió el enlace de firma al cliente.',
      '<p>Aprobaste el congelamiento de <b>' + esc_(loan.name) + '</b> (' + esc_(loanId) + ').</p>' + amtTable +
      '<p>Se envió el enlace de firma al cliente; el préstamo se congelará al firmar.</p>' + btn);
  } catch (e) { logError_('approveFreezeRequest_:email', e); }
  try { getSS_().toast(loanId + ': aprobado. Enlace de firma enviado al cliente.', '✔ Congelamiento aprobado', 6); } catch (e) { }
}

/** RECHAZA la solicitud de la fila: usa "Motivo de rechazo", avisa por correo a ambas partes y
 *  arma el WhatsApp con el motivo. No toca el préstamo. */
function denyFreezeRequest_(sh, row) {
  const m = headerMap_(sh);
  const g = h => m[h] ? sh.getRange(row, m[h]).getValue() : '';
  const setv = (h, v) => { if (m[h]) sh.getRange(row, m[h]).setValue(v); };
  const setf = (h, f) => { if (m[h]) sh.getRange(row, m[h]).setFormula(f); };
  const loanId = String(g('ID Préstamo')).trim();
  const name = String(g('Nombre')).trim(), email = g('Correo'), phone = g('Teléfono');
  const reason = String(g('Motivo de rechazo')).trim() || 'No cumple las condiciones para el congelamiento en este momento.';
  setv('Estado', 'Rechazado');
  const num = waNumberAR_(phone);
  if (num) {
    const msg = 'Hola ' + (name || 'cliente') + ' 👋 Revisamos tu pedido de congelamiento del préstamo ' + loanId +
      ' y, por ahora, no podemos aprobarlo. Motivo: ' + reason + '. — ' + companyName_();
    setf('WhatsApp', '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📲 Avisar rechazo")');
  }
  const lender = getSetting_('Correo del Prestamista');
  try {
    if (email) sendBrandedEmail_(email, 'Sobre tu solicitud de congelamiento — ' + loanId,
      'Hola ' + name + ', revisamos tu pedido de congelamiento del préstamo ' + loanId + ' y no podemos aprobarlo. Motivo: ' + reason,
      '<p>Hola ' + esc_(name) + ',</p><p>Revisamos tu pedido de <b>congelamiento</b> del préstamo <b>' + esc_(loanId) + '</b> y, por ahora, <b>no podemos aprobarlo</b>.</p><p><b>Motivo:</b> ' + esc_(reason) + '</p><p>Comunicate con nosotros para regularizar tu situación.</p>');
    if (lender) sendBrandedEmail_(lender, 'Congelamiento rechazado — ' + loanId,
      'Rechazaste el congelamiento de ' + name + ' (' + loanId + '). Motivo: ' + reason,
      '<p>Rechazaste el congelamiento de <b>' + esc_(name) + '</b> (' + esc_(loanId) + ').</p><p><b>Motivo:</b> ' + esc_(reason) + '</p>');
  } catch (e) { logError_('denyFreezeRequest_:email', e); }
  try { getSS_().toast(loanId + ': rechazado.', '✖ Congelamiento rechazado', 6); } catch (e) { }
}

/** APRUEBA la solicitud de PLAN DE PAGO de la fila: exige el % inicial pagado, arma el enlace para
 *  firmar el Acuerdo + WhatsApp + correos. NO activa el plan todavía: se activa al FIRMAR. */
function approvePlanRequest_(sh, row) {
  const m = headerMap_(sh);
  const g = h => m[h] ? sh.getRange(row, m[h]).getValue() : '';
  const setv = (h, v) => { if (m[h]) sh.getRange(row, m[h]).setValue(v); };
  const setf = (h, f) => { if (m[h]) sh.getRange(row, m[h]).setFormula(f); };
  const loanId = String(g('ID Préstamo')).trim();
  if (!loanId) return;
  const loan = findLoanById_(loanId);
  if (!loan) { setv('Estado', 'Rechazado'); setv('Motivo de rechazo', 'No se encontró el préstamo.'); return; }
  // GATE: el PAGO INICIAL (% configurable) debe estar PAGADO (registrado en «Pagos») ANTES de aprobar.
  // Si falta, NO aprueba y deja constancia VISIBLE (nota en la fila + aviso) — porque getUi().alert no
  // se muestra desde el disparador onEdit, y el toast puede pasar desapercibido.
  const amt = planAmounts_(loan);
  if (!amt.downMet) {
    const msg = 'Falta pagar el ' + amt.pctInt + '% (' + fmtMoney_(amt.required) + '): faltan ' + fmtMoney_(amt.downShort) +
      '. Registrá el pago en «Pagos» y volvé a tildar «Aprobar».';
    if (m['Aprobar ✔']) sh.getRange(row, m['Aprobar ✔']).setNote('⛔ ' + msg);
    if (m['Estado']) sh.getRange(row, m['Estado']).setNote('⛔ No aprobado — ' + msg);
    // Mensaje VISIBLE. Desde el disparador onEdit no se puede abrir un diálogo (getUi().alert falla),
    // así que mostramos un toast PERSISTENTE (timeout -1 = queda hasta que lo cierres). Si se corre
    // desde el menú (hay UI), además se abre el alert modal.
    try { getSS_().toast('Préstamo ' + loanId + ' — ' + msg, '⛔ NO SE APROBÓ · Falta el pago inicial', -1); } catch (e) { }
    try { SpreadsheetApp.getUi().alert('No se aprobó el plan ' + loanId + '.\n\n' + msg); } catch (e) { }
    return; // queda Pendiente hasta que se registre el pago inicial
  }
  if (m['Aprobar ✔']) sh.getRange(row, m['Aprobar ✔']).clearNote();   // cubierto: limpia la nota de bloqueo
  if (m['Estado']) sh.getRange(row, m['Estado']).clearNote();
  const option = Math.floor(Number(g('Opción (1/4/6)')) || 0);
  const optTxt = planOptTxt_(option, planCadence_(g('Cadencia')));
  setv('Estado', 'Aprobado');
  if (m['Fecha de aprobación']) sh.getRange(row, m['Fecha de aprobación']).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  setv('Estado de Firma', 'PENDIENTE');
  const base = webAppBaseUrl_();
  const url = base ? base + '?page=firmar-plan&loan=' + encodeURIComponent(loanId) : '';
  if (url) setf('Enlace de firma', '=HYPERLINK("' + url + '","✍ Firmar Acuerdo")');
  else setv('Enlace de firma', 'Publicá la app web (menú ⑥) para obtener el enlace.');
  const downTxt = amt.downMet
    ? ('El ' + amt.pctInt + '% (' + fmtMoney_(amt.required) + ') ya está cubierto.')
    : ('Para activar: pagá el ' + amt.pctInt + '% (' + fmtMoney_(amt.required) + '; falta ' + fmtMoney_(amt.downShort) + ') y firmá el Acuerdo.');
  const num = waNumberAR_(loan.phone);
  if (num && url) {
    const msg = 'Hola ' + (loan.name || 'cliente') + ' 👋 Aprobamos tu plan de pago (' + optTxt + ') del préstamo ' + loanId +
      '. ' + downTxt + ' Firmá el Acuerdo acá: ' + url + ' — ' + companyName_();
    setf('WhatsApp', '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📲 Avisar aprobación")');
  }
  const lender = getSetting_('Correo del Prestamista');
  const btn = url ? '<p><a href="' + url + '" style="background:#674ea7;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold">✍ Firmar el Acuerdo</a></p>' : '';
  const amtTable = '<table style="border-collapse:collapse;margin:6px 0">' +
    '<tr><td><b>Opción</b>&nbsp;</td><td>' + esc_(optTxt) + '</td></tr>' +
    '<tr><td><b>Total a pagar</b>&nbsp;</td><td>' + fmtMoney_(amt.total) + '</td></tr>' +
    '<tr><td><b>Pago inicial (' + amt.pctInt + '%)</b>&nbsp;</td><td>' + fmtMoney_(amt.required) + (amt.downMet ? ' — cubierto' : ' — falta ' + fmtMoney_(amt.downShort)) + '</td></tr></table>';
  try {
    if (loan.email) sendBrandedEmail_(loan.email, 'Tu plan de pago fue aprobado — ' + loanId,
      'Hola ' + loan.name + ', aprobamos tu plan de pago del préstamo ' + loanId + '. Firmá el Acuerdo: ' + (url || '(enlace pendiente)'),
      '<p>Hola ' + esc_(loan.name) + ',</p><p>Aprobamos tu <b>plan de pago</b> del préstamo <b>' + esc_(loanId) + '</b>.</p>' + amtTable +
      '<p>' + esc_(downTxt) + ' Al firmar se <b>suspende la mora</b> y se reemplazan las cuotas.</p>' + btn);
    if (lender) sendBrandedEmail_(lender, 'Plan de pago aprobado — ' + loanId,
      'Aprobaste el plan de pago de ' + loan.name + ' (' + loanId + '). Se envió el enlace de firma.',
      '<p>Aprobaste el plan de pago de <b>' + esc_(loan.name) + '</b> (' + esc_(loanId) + ').</p>' + amtTable +
      '<p>Se envió el enlace de firma; el plan se activa al firmar (reemplaza las cuotas).</p>' + btn);
  } catch (e) { logError_('approvePlanRequest_:email', e); }
  try { getSS_().toast(loanId + ': plan aprobado. Enlace de firma enviado.', '✔ Plan aprobado', 6); } catch (e) { }
}

/** RECHAZA la solicitud de plan de la fila: usa "Motivo de rechazo", avisa por correo + WhatsApp. */
function denyPlanRequest_(sh, row) {
  const m = headerMap_(sh);
  const g = h => m[h] ? sh.getRange(row, m[h]).getValue() : '';
  const setv = (h, v) => { if (m[h]) sh.getRange(row, m[h]).setValue(v); };
  const setf = (h, f) => { if (m[h]) sh.getRange(row, m[h]).setFormula(f); };
  const loanId = String(g('ID Préstamo')).trim();
  const name = String(g('Nombre')).trim(), email = g('Correo'), phone = g('Teléfono');
  const reason = String(g('Motivo de rechazo')).trim() || 'No cumple las condiciones para el plan de pago en este momento.';
  setv('Estado', 'Rechazado');
  const num = waNumberAR_(phone);
  if (num) {
    const msg = 'Hola ' + (name || 'cliente') + ' 👋 Revisamos tu pedido de plan de pago del préstamo ' + loanId +
      ' y, por ahora, no podemos aprobarlo. Motivo: ' + reason + '. — ' + companyName_();
    setf('WhatsApp', '=HYPERLINK("https://wa.me/' + num + '?text=' + encodeURIComponent(msg) + '","📲 Avisar rechazo")');
  }
  const lender = getSetting_('Correo del Prestamista');
  try {
    if (email) sendBrandedEmail_(email, 'Sobre tu solicitud de plan de pago — ' + loanId,
      'Hola ' + name + ', revisamos tu pedido de plan de pago del préstamo ' + loanId + ' y no podemos aprobarlo. Motivo: ' + reason,
      '<p>Hola ' + esc_(name) + ',</p><p>Revisamos tu pedido de <b>plan de pago</b> del préstamo <b>' + esc_(loanId) + '</b> y, por ahora, <b>no podemos aprobarlo</b>.</p><p><b>Motivo:</b> ' + esc_(reason) + '</p>');
    if (lender) sendBrandedEmail_(lender, 'Plan de pago rechazado — ' + loanId,
      'Rechazaste el plan de pago de ' + name + ' (' + loanId + '). Motivo: ' + reason,
      '<p>Rechazaste el plan de pago de <b>' + esc_(name) + '</b> (' + esc_(loanId) + ').</p><p><b>Motivo:</b> ' + esc_(reason) + '</p>');
  } catch (e) { logError_('denyPlanRequest_:email', e); }
  try { getSS_().toast(loanId + ': plan rechazado.', '✖ Plan rechazado', 6); } catch (e) { }
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

/**
 * Columnas de firma de "Prestatarios" resueltas por NOMBRE de encabezado (robusto al
 * layout migrado, donde el orden físico difiere del mapa PB). Con respaldo a PB para el
 * esquema nuevo. Úsese en el flujo de firma en lugar de los índices fijos COL_SIGN_*.
 */
function firmaCols_(sh) {
  const H = headerIndex_(sh);
  return {
    status: colByAny_(H, ['Estado de Firma']) || PB.SIGN_STATUS,
    date: colByAny_(H, ['Fecha de Firma']) || PB.SIGN_DATE,
    pdf: colByAny_(H, ['Contrato Firmado (PDF)']) || PB.SIGN_PDF,
    loanDate: colByAny_(H, ['Fecha del Préstamo', 'Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo']) || PB.LOAN_DATE,
  };
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
    const status = loc ? String(loc.sh.getRange(loc.row, firmaCols_(loc.sh).status).getValue() || '').toUpperCase() : '';
    if (status === SIGN.SIGNED)
      return { ok: true, signed: true, msg: 'Este contrato ya fue firmado. Le enviamos una copia por correo.' };
    return { ok: true, signed: false, name: loan.name, contractHtml: agreementHtml_(loan), captureLocation: locationFlag_() };
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
    const fc = firmaCols_(loc.sh);
    if (String(loc.sh.getRange(loc.row, fc.status).getValue() || '').toUpperCase() === SIGN.SIGNED)
      return { ok: true, message: 'Este contrato ya estaba firmado. Le enviamos una copia por correo.' };

    const now = new Date();
    const mime = m[1], ext = mime.split('/')[1].replace('+xml', '').replace('jpeg', 'jpg').replace('svg', 'png');
    const bytes = Utilities.base64Decode(m[2]);
    const method = (meta && meta.method === 'upload') ? 'upload' : 'draw';
    const ip = String((meta && meta.ip) || '').slice(0, 60);
    const devInfo = (meta && meta.device) || {};
    const deviceFp = deviceFingerprint_(devInfo); // huella única y estable del equipo
    const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
    folder.createFile(Utilities.newBlob(bytes, mime, 'Firma ' + loanId + '.' + ext));

    // El reloj de interés arranca al firmar (Fecha del Préstamo = hoy).
    loc.sh.getRange(loc.row, fc.loanDate).setValue(now).setNumberFormat('yyyy-mm-dd');
    writeOutstandingRow_(loc.sh, loc.row);
    const freshLoan = readLoan_(loc.sh, loc.row);
    // Regenera el cronograma con la fecha DEFINITIVA (la firma mueve "Fecha del Préstamo"
    // a hoy). Las cuotas se habían creado al aprobar con la fecha de aprobación; sin esto
    // quedaban desfasadas respecto del "Vencimiento" del encabezado. cuotasForLoan_ borra-y-recrea.
    try {
      const cs = getSS_().getSheetByName(CFG.SHEETS.INSTALLMENTS) || setupCuotas_(getSS_());
      const ti = resolveTerm_(freshLoan.principal, freshLoan.term);
      const totalPlan = round2_(freshLoan.principal * (1 + ti.rate));
      const dueDays = ti.cadence === 'monthly' ? Array.from({ length: ti.cuotas }, (_, i) => (i + 1) * 30) : [ti.days];
      cuotasForLoan_(cs, loanId, freshLoan.loanDate, totalPlan, ti.cuotas, dueDays);
    } catch (e) { logError_('submitSignature:cuotas', e); }

    const sig = {
      dataUrl: signatureDataUrl, signedAt: now, method: method,
      userAgent: (meta && meta.userAgent) || '', ip: ip, hash: '',
      device: devInfo, deviceFp: deviceFp, // resumen de dispositivo para el panel del PDF
    };
    const signedFile = makeSignedAgreementFile_(freshLoan, sig);
    // Hash SHA-256 del PDF firmado (huella de integridad del documento).
    let hashHex = '';
    try { hashHex = hexDigest_(signedFile.getAs('application/pdf').getBytes()); } catch (e) { logError_('submitSignature:hash', e); }

    loc.sh.getRange(loc.row, fc.status).setValue(SIGN.SIGNED);
    loc.sh.getRange(loc.row, fc.date).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
    loc.sh.getRange(loc.row, fc.pdf).setFormula('=HYPERLINK("' + signedFile.getUrl() + '","Ver contrato firmado")');

    try {
      setupFirmas_(getSS_()).appendRow([now, loanId, loan.name, loan.dni, loan.email,
        method === 'upload' ? 'Imagen cargada' : 'Dibujada en pantalla',
        String((meta && meta.userAgent) || '').slice(0, 250), signedFile.getUrl(),
        ip || 'N/D', hashHex, freshLoan.cuil || '', freshLoan.direccion || '']);
    } catch (e) { logError_('submitSignature:firmas', e); }
    // Detalle completo de dispositivo/ubicación/fraude → hoja "Dispositivos" (best-effort).
    try {
      logDeviceInfo_(freshLoan, freshLoan.clientId, {
        device: devInfo, deviceFp: deviceFp, ip: ip, userAgent: (meta && meta.userAgent) || '',
        method: method, signedAt: now,
        dwellMs: meta && meta.dwellMs, strokes: meta && meta.strokes, drawMs: meta && meta.drawMs,
      });
    } catch (e) { logError_('submitSignature:dispositivos', e); }
    let emailOk = true;
    try { emailSignedContract_(freshLoan, signedFile); } catch (e) { emailOk = false; logError_('submitSignature:email', e); }
    try { getSS_().toast('Contrato firmado: ' + loan.name + ' → ' + loanId + (emailOk ? '' : ' (⚠ no se pudo enviar el correo)'), 'Firma recibida ✍', 6); } catch (e) {}
    return { ok: true, message: emailOk
      ? '¡Contrato firmado! Gracias, ' + loan.name + '. Le enviamos una copia por correo.'
      : '¡Contrato firmado! Gracias, ' + loan.name + '. (No pudimos enviar la copia por correo; el prestamista ya tiene su contrato firmado.)' };
  });
}

/* ===================== FIRMA DEL ACUERDO DE CONGELAMIENTO (Parte 2) ===================== */
/** Página para firmar el acuerdo de congelamiento (reusa la de firma en modo 'freeze'). */
function freezeSigningPageHtml_(loanId) { return signingPageHtml_(loanId, 'freeze'); }

/** Ubica la solicitud de congelamiento por ID de préstamo y estado(s). Devuelve {sh,row,m} o null. */
function findFreezeRequestRow_(loanId, states) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.FREEZE_REQ); if (!sh || sh.getLastRow() < 2) return null;
  const m = headerMap_(sh), idC = m['ID Préstamo'] || 2, stC = m['Estado'] || 23;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(idC, stC)).getValues();
  const id = String(loanId).trim(), set = states ? states.map(String) : null;
  for (let i = data.length - 1; i >= 0; i--)   // la más reciente primero
    if (String(data[i][idC - 1]).trim() === id && (!set || set.indexOf(String(data[i][stC - 1]).trim()) !== -1))
      return { sh: sh, row: i + 2, m: m };
  return null;
}

/** Verifica identidad + solicitud APROBADA; devuelve el acuerdo de congelamiento a firmar. */
function getFreezeContractForSigning(loanId, dni) {
  return guard_('getFreezeContractForSigning', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, msg: 'Datos no encontrados. Verificá el ID del préstamo y tu DNI.' };
    const req = findFreezeRequestRow_(loanId, ['Aprobado', 'Firmado']);
    if (!req) return { ok: false, msg: 'No hay una solicitud de congelamiento aprobada para este préstamo.' };
    const g = h => req.m[h] ? req.sh.getRange(req.row, req.m[h]).getValue() : '';
    if (String(g('Estado de Firma')).toUpperCase() === SIGN.SIGNED || String(g('Estado')).trim() === 'Firmado')
      return { ok: true, signed: true, msg: 'El acuerdo de congelamiento ya fue firmado. Te enviamos una copia por correo.' };
    const fz = loanFreezable_(loan);
    if (!fz.ok) return { ok: false, msg: fz.reason };
    return { ok: true, signed: false, name: loan.name, contractHtml: freezeAgreementHtml_(loan, { promised: g('Fecha de pago prometida'), alreadyPaid: g('Ya pagado (declarado)') }), captureLocation: locationFlag_() };
  });
}

/** Recibe la firma del acuerdo: CONGELA el préstamo (hoy), guarda el PDF, registra auditoría,
 *  marca la solicitud como Firmada y suma el congelamiento al historial del cliente. */
function submitFreezeSignature(loanId, dni, signatureDataUrl, agree, meta) {
  return guard_('submitFreezeSignature', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      throw new Error('Datos no encontrados. Verificá el ID del préstamo y tu DNI.');
    if (agree !== true) throw new Error('Debés aceptar los términos para firmar.');
    const req = findFreezeRequestRow_(loanId, ['Aprobado', 'Firmado']);
    if (!req) throw new Error('No hay una solicitud de congelamiento aprobada para este préstamo.');
    const g = h => req.m[h] ? req.sh.getRange(req.row, req.m[h]).getValue() : '';
    const setv = (h, v) => { if (req.m[h]) req.sh.getRange(req.row, req.m[h]).setValue(v); };
    const setf = (h, f) => { if (req.m[h]) req.sh.getRange(req.row, req.m[h]).setFormula(f); };
    if (String(g('Estado de Firma')).toUpperCase() === SIGN.SIGNED)
      return { ok: true, message: 'El acuerdo ya estaba firmado. Te enviamos una copia por correo.' };
    const m = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i.exec(String(signatureDataUrl || ''));
    if (!m) throw new Error('Falta la firma. Dibujá tu firma o cargá una imagen.');
    const fz = loanFreezable_(loan);
    if (!fz.ok) throw new Error(fz.reason);
    // GATE: el % configurable debe estar PAGADO (registrado) antes de congelar.
    const amtGate = freezeAmounts_(loan);
    if (!amtGate.downMet) throw new Error('Para congelar, primero debe estar pagado el ' + amtGate.pctInt + '% del préstamo. Falta ' + fmtMoney_(amtGate.downShort) + '.');

    const now = new Date();
    const mime = m[1], ext = mime.split('/')[1].replace('+xml', '').replace('jpeg', 'jpg').replace('svg', 'png');
    const bytes = Utilities.base64Decode(m[2]);
    const method = (meta && meta.method === 'upload') ? 'upload' : 'draw';
    const ip = String((meta && meta.ip) || '').slice(0, 60);
    const devInfo = (meta && meta.device) || {};
    const deviceFp = deviceFingerprint_(devInfo);
    const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
    folder.createFile(Utilities.newBlob(bytes, mime, 'Firma Congelamiento ' + loanId + '.' + ext));

    // CONGELA el préstamo: "Congelado el" = hoy en "Prestatarios" (reutiliza la columna de la Parte 1).
    const loc = signStatusRow_(loanId);
    if (loc) { const sinceC = colByAny_(headerIndex_(loc.sh), ['Congelado el']);
      if (sinceC) loc.sh.getRange(loc.row, sinceC).setValue(now).setNumberFormat('yyyy-mm-dd'); }

    const promised = g('Fecha de pago prometida');
    const sig = { dataUrl: signatureDataUrl, signedAt: now, method: method, userAgent: (meta && meta.userAgent) || '', ip: ip, hash: '', device: devInfo, deviceFp: deviceFp };
    const signedFile = savePdfToBorrower_(freezeAgreementHtml_(loan, { promised: promised, alreadyPaid: g('Ya pagado (declarado)') }, sig), 'Acuerdo de Congelamiento ' + loanId + ' - ' + loan.name + '.pdf', loan);
    let hashHex = ''; try { hashHex = hexDigest_(signedFile.getAs('application/pdf').getBytes()); } catch (e) { logError_('submitFreezeSignature:hash', e); }

    try { setupFirmas_(getSS_()).appendRow([now, loanId, loan.name, loan.dni, loan.email, 'Congelamiento (' + (method === 'upload' ? 'imagen' : 'dibujada') + ')',
      String((meta && meta.userAgent) || '').slice(0, 250), signedFile.getUrl(), ip || 'N/D', hashHex, loan.cuil || '', loan.direccion || '']); } catch (e) { logError_('submitFreezeSignature:firmas', e); }
    try { logDeviceInfo_(loan, loan.clientId, { device: devInfo, deviceFp: deviceFp, ip: ip, userAgent: (meta && meta.userAgent) || '', method: 'congelamiento', signedAt: now, dwellMs: meta && meta.dwellMs, strokes: meta && meta.strokes, drawMs: meta && meta.drawMs }); } catch (e) { logError_('submitFreezeSignature:dispositivos', e); }

    setv('Estado', 'Firmado'); setv('Estado de Firma', SIGN.SIGNED);
    if (req.m['Fecha de Firma']) req.sh.getRange(req.row, req.m['Fecha de Firma']).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
    setf('Contrato de Congelamiento (PDF)', '=HYPERLINK("' + signedFile.getUrl() + '","Ver acuerdo firmado")');

    try { incrementClienteFreezeCount_(loan.clientId); } catch (e) { logError_('submitFreezeSignature:clienteFlag', e); }

    const lender = getSetting_('Correo del Prestamista');
    let emailOk = true;
    try {
      const att = signedFile.getAs('application/pdf');
      if (loan.email) sendBrandedEmail_(loan.email, 'Acuerdo de congelamiento firmado — ' + loanId,
        'Hola ' + loan.name + ', tu acuerdo de congelamiento del préstamo ' + loanId + ' quedó firmado. La mora dejó de crecer.',
        '<p>Hola ' + esc_(loan.name) + ',</p><p>Tu <b>acuerdo de congelamiento</b> del préstamo <b>' + esc_(loanId) + '</b> quedó firmado. A partir de hoy <b>la mora deja de crecer</b>' + (promised instanceof Date ? (' y te comprometiste a pagar antes del <b>' + fmtDate_(promised) + '</b>') : '') + '. Adjuntamos tu copia.</p>',
        { attachments: [att] });
      if (lender) sendBrandedEmail_(lender, 'Congelamiento firmado — ' + loanId,
        loan.name + ' firmó el acuerdo de congelamiento del préstamo ' + loanId + '. El préstamo quedó congelado.',
        '<p><b>' + esc_(loan.name) + '</b> firmó el acuerdo de congelamiento del préstamo <b>' + esc_(loanId) + '</b>. El préstamo quedó <b>CONGELADO</b> desde hoy.</p>',
        { attachments: [att] });
    } catch (e) { emailOk = false; logError_('submitFreezeSignature:email', e); }
    try { getSS_().toast(loanId + ': congelamiento firmado y aplicado.', '❄ Congelado', 6); } catch (e) { }
    return { ok: true, message: emailOk ? '¡Acuerdo firmado! La mora dejó de crecer. Te enviamos una copia por correo.' : '¡Acuerdo firmado! La mora dejó de crecer. (No pudimos enviar la copia por correo.)' };
  });
}

/** Suma 1 al contador "Congelamientos" del cliente (crea la columna si falta). Historial durable. */
function incrementClienteFreezeCount_(clientId) {
  clientId = String(clientId || '').trim(); if (!clientId) return;
  const cs = getSS_().getSheetByName(CFG.SHEETS.CLIENTS); if (!cs || cs.getLastRow() < 2) return;
  const m = headerMap_(cs), idC = m['ID Cliente'] || PC.CLIENT_ID;
  let col = m['Congelamientos']; if (!col) col = clientesHeaderCol_(cs, 'Congelamientos', true);
  const ids = cs.getRange(2, idC, cs.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === clientId) {
    const cell = cs.getRange(i + 2, col); cell.setValue((Number(cell.getValue()) || 0) + 1); return;
  }
}

/** HTML del ACUERDO DE CONGELAMIENTO (nuevos términos) — breve, con firma del prestatario + titulares y auditoría. */
function freezeAgreementHtml_(loan, req, sig) {
  const lender = companyName_(), footer = getSetting_('Pie del Contrato');
  const signatories = ownerSignatories_();
  const a = freezeAmounts_(loan, req && req.alreadyPaid);
  const promised = (req && req.promised instanceof Date) ? req.promised : null;
  const signedAt = (sig && sig.signedAt instanceof Date) ? sig.signedAt : new Date();
  const deadline24 = new Date(signedAt.getTime() + 24 * 3600 * 1000);
  const kv = (k, v) => `<tr><td class="k">${k}</td><td>${v}</td></tr>`;
  const borrowerSig = (sig && sig.dataUrl) ? sig.dataUrl : '';
  const sigCell = (img, name, role) =>
    `<td>${img ? `<img src="${img}" alt="firma" style="max-height:44px;max-width:88%;display:block;margin:0 auto 1px">` : '<div style="height:40px"></div>'}` +
    `<div style="border-top:1px solid #333;padding-top:3px">${esc_(name)}<br><span style="font-size:8pt;color:#555">${esc_(role)}</span></div></td>`;
  const locSummary = s => { const d = (s && s.device) || {}, l = d.loc || {};
    if (l.source === 'GPS' && l.gps && l.gps.indexOf('GPS:') !== 0) return l.gps + (l.gpsAcc != null ? ' (±' + l.gpsAcc + ' m)' : '') + ' · GPS';
    if (l.ipText) return l.ipText + ' · IP'; if (l.source === 'desactivada') return 'No capturada'; return d.tz ? ('Zona horaria: ' + d.tz) : '—'; };
  const fill = v => (v && String(v).trim()) ? esc_(v) : '<span style="color:#9aa4ad">—</span>';
  const mono = '<span style="font-family:monospace;font-size:8.5pt;word-break:break-all">';
  // Constancia de firma electrónica COMPLETA — mismos campos que el contrato de préstamo.
  const auditPanel = (sig && sig.dataUrl) ? `
    <h2>Constancia de firma electrónica</h2>
    <table>
      ${kv('Firmó', esc_(loan.name) + ' · DNI ' + esc_(loan.dni || '—'))}
      ${kv('Correo confirmado', esc_(loan.email || '—'))}
      ${kv('Identidad verificada con', 'DNI ' + esc_(loan.dni || '—') + ' y correo registrado; foto del DNI (frente y dorso, Anexo I)')}
      ${kv('Documento', 'Foto del DNI, frente y dorso — Anexo I')}
      ${kv('Fecha y hora', fmtDateTime_(sig.signedAt || new Date()))}
      ${kv('Cómo se firmó', sig.method === 'upload' ? 'Imagen de firma cargada' : 'Firma dibujada en pantalla')}
      ${kv('Dirección IP', fill(sig.ip))}
      ${kv('Ubicación', esc_(locSummary(sig)))}
      ${kv('Equipo / SO', fill((sig.device && sig.device.os) || ''))}
      ${kv('Navegador', fill((sig.device && sig.device.browser) || ''))}
      ${kv('Pantalla', fill((sig.device && sig.device.screen) ? (sig.device.screen + (sig.device.dpr ? (' · DPR ' + sig.device.dpr) : '')) : ''))}
      ${kv('Automatización detectada', (sig.device && sig.device.webdriver) ? 'Sí (revisar)' : 'No (firmante humano)')}
      ${kv('Dispositivo (completo)', esc_(String(sig.userAgent || '').slice(0, 200)))}
      ${kv('Huella de dispositivo (SHA-256)', sig.deviceFp ? (mono + esc_(sig.deviceFp) + '</span>') : '<span style="color:#9aa4ad">No disponible</span>')}
      ${kv('Huella del documento (SHA-256)', sig.hash ? (mono + esc_(sig.hash) + '</span>') : '<span style="color:#9aa4ad">Se registra en la hoja «Firmas».</span>')}
    </table>` : '';
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    @page{size:A4;margin:14mm 14mm}
    body{font-family:"Carlito","Calibri",Arial,sans-serif;color:#1d2630;margin:0;line-height:1.35;font-size:10pt}
    h1{text-align:center;font-size:15pt;border-bottom:2px solid #0b5394;padding-bottom:4px;color:#0b2f52;margin:2px 0 6px}
    h2{font-size:10.5pt;color:#0b5394;margin:12px 0 3px;border-bottom:1px solid #cfe0f0;padding-bottom:2px;text-transform:uppercase;letter-spacing:.4px}
    p{margin:0 0 5px} table{width:100%;border-collapse:collapse;margin:5px 0;font-size:9.2pt} td,th{padding:4px 7px;border:1px solid #d8dee3;vertical-align:top}
    td.k{background:#eef4fb;font-weight:bold;width:46%}
    .box{border:1px solid #0b5394;border-left:4px solid #0b5394;background:#eef4fb;padding:7px 11px;margin:7px 0}
    .sign{margin-top:16px;width:100%} .sign td{border:none;text-align:center;padding:6px 10px;width:33%;vertical-align:bottom}
    .foot{margin-top:16px;padding-top:6px;border-top:1px solid #dfe4e8;font-size:8.6pt;color:#5c6871;font-style:italic}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>ACUERDO DE CONGELAMIENTO DE PRÉSTAMO</h1>
    <p style="text-align:center;color:#3d4952">Celebrado el <b>${fmtDate_(new Date())}</b> entre ${esc_(signatories[0].name)} y ${esc_(signatories[1].name)}, que operan como «${esc_(lender)}» (el Prestamista), y <b>${esc_(loan.name)}</b> (el Prestatario), respecto del préstamo <b>${esc_(loan.loanId)}</b>.</p>
    <h2>1. Antecedente</h2>
    <p>El préstamo <b>${esc_(loan.loanId)}</b>, con vencimiento original el <b>${fmtDate_(loan.dueDate)}</b>, se encuentra vencido. El Prestatario solicitó, y el Prestamista aceptó, <b>congelar la mora</b> (detener el recargo diario) bajo las condiciones de abajo.</p>
    <h2>2. Saldo y plan de pago</h2>
    <div class="box"><table style="margin:0">
      ${kv('Préstamo', esc_(loan.loanId))}
      ${kv('Total del préstamo (capital + interés)', fmtMoney_(a.total))}
      ${kv('Recargo por mora acumulado a la fecha', fmtMoney_(a.mora))}
      ${kv('Saldo al congelar (capital + interés + mora)', '<b>' + fmtMoney_(a.balance) + '</b>')}
      ${kv('Pago mínimo requerido (' + a.pctInt + '% del préstamo, en 24 h)', '<b>' + fmtMoney_(a.required) + '</b>')}
      ${kv('Pagado a la fecha (registrado)', fmtMoney_(a.recorded))}
      ${a.downShort > 0 ? kv('Falta para el ' + a.pctInt + '%', '<b style="color:#900">' + fmtMoney_(a.downShort) + '</b>') : ''}
      ${kv('Saldo restante (a abonar por la fecha pactada)', '<b>' + fmtMoney_(a.remaining) + '</b>' + (promised ? (' — antes del ' + fmtDate_(promised)) : ''))}
    </table></div>
    <h2>3. Condiciones (aceptadas por el Prestatario)</h2>
    <p><b>a)</b> La mora queda <b>CONGELADA</b> desde la firma (${fmtDate_(signedAt)}): el recargo diario deja de crecer mientras rija este acuerdo.</p>
    <p><b>b)</b> Como condición del congelamiento, el Prestatario debe tener pagado el <b>${a.pctInt}% del préstamo (${fmtMoney_(a.required)})</b> dentro de las <b>24 horas</b> (antes de la aprobación). ${a.downMet ? 'Este requisito se encuentra <b>cumplido</b>.' : 'Pendiente: falta <b>' + fmtMoney_(a.downShort) + '</b>.'}</p>
    <p><b>c)</b> El Prestatario pagará el <b>saldo restante (${fmtMoney_(a.remaining)})</b> ${promised ? ('antes del <b>' + fmtDate_(promised) + '</b>') : 'antes de la fecha comprometida'} (plazo máximo de 3 semanas desde la firma).</p>
    <p><b>d)</b> Si el Prestatario no cumple con alguno de estos pagos en término, el Prestamista podrá <b>descongelar</b> el préstamo: la mora volverá a correr desde ese momento (los días transcurridos durante el congelamiento no se cobran) y podrá exigir el total adeudado.</p>
    <p><b>e)</b> El Prestatario entiende y acepta que, habiendo congelado este préstamo, es muy probable que un futuro préstamo, de otorgarse, sea por un <b>monto considerablemente menor</b>.</p>
    <p>Las demás cláusulas del contrato original del préstamo ${esc_(loan.loanId)} siguen vigentes en todo lo que no se oponga a este acuerdo.</p>
    <table class="sign"><tr>
      ${sigCell(borrowerSig, loan.name, 'Prestatario · DNI ' + (loan.dni || ''))}
      ${sigCell(signatories[0].sig, signatories[0].name, 'Prestamista')}
      ${sigCell(signatories[1].sig, signatories[1].name, 'Prestamista')}
    </tr></table>
    ${auditPanel}
    <p class="foot">${esc_(footer || '')}</p>
  </body></html>`;
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
  const statusCol = firmaCols_(sh).status;
  const st = String(sh.getRange(row, statusCol).getValue() || '').toUpperCase();
  if (st === SIGN.SIGNED && ui.alert('Este contrato ya figura FIRMADO. ¿Reenviar el enlace igualmente?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  try {
    if (st !== SIGN.SIGNED) sh.getRange(row, statusCol).setValue(SIGN.PENDING);
    let file = null; try { file = makeAgreementFile_(loan); } catch (e) { logError_('resendSigningLink_:pdf', e); }
    emailSigningRequest_(loan, file);
    ss.toast('Enlace de firma reenviado a ' + loan.email, 'Listo', 5);
  } catch (err) { logError_('resendSigningLink_', err); ui.alert('No se pudo reenviar: ' + (err.message || err)); }
}

/** Página web independiente para firmar el contrato (?page=firmar&loan=L-0001). */
/* ===================== FIRMA DEL ACUERDO DE PLAN DE PAGO ===================== */
/** Página para firmar el Acuerdo de plan de pago (reusa la página de firma en modo 'plan'). */
function planSigningPageHtml_(loanId) { return signingPageHtml_(loanId, 'plan'); }

/** Ubica la solicitud de plan por ID de préstamo y estado(s). Devuelve {sh,row,m} o null. */
function findPlanRequestRow_(loanId, states) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.PLAN_REQ); if (!sh || sh.getLastRow() < 2) return null;
  const m = headerMap_(sh), idC = m['ID Préstamo'] || 2, stC = m['Estado'] || 23;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(idC, stC)).getValues();
  const id = String(loanId).trim(), set = states ? states.map(String) : null;
  for (let i = data.length - 1; i >= 0; i--)
    if (String(data[i][idC - 1]).trim() === id && (!set || set.indexOf(String(data[i][stC - 1]).trim()) !== -1))
      return { sh: sh, row: i + 2, m: m };
  return null;
}

/** Verifica identidad + solicitud APROBADA; devuelve el Acuerdo de plan de pago a firmar. */
function getPlanContractForSigning(loanId, dni) {
  return guard_('getPlanContractForSigning', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, msg: 'Datos no encontrados. Verificá el ID del préstamo y tu DNI.' };
    const req = findPlanRequestRow_(loanId, ['Aprobado', 'Firmado']);
    if (!req) return { ok: false, msg: 'No hay una solicitud de plan de pago aprobada para este préstamo.' };
    const g = h => req.m[h] ? req.sh.getRange(req.row, req.m[h]).getValue() : '';
    if (String(g('Estado de Firma')).toUpperCase() === SIGN.SIGNED || String(g('Estado')).trim() === 'Firmado')
      return { ok: true, signed: true, msg: 'El Acuerdo ya fue firmado. Te enviamos una copia por correo.' };
    const option = Math.floor(Number(g('Opción (1/4/6)')) || 4);
    const cadence = planCadence_(g('Cadencia'));
    return { ok: true, signed: false, name: loan.name, contractHtml: planAgreementHtml_(loan, { option: option, cadence: cadence }), captureLocation: locationFlag_() };
  });
}

/** Recibe la firma del Acuerdo: ACTIVA el plan (suspende mora, reemplaza cuotas, marca "Plan desde"),
 *  guarda el PDF, registra auditoría (Firmas/Dispositivos) y marca la solicitud como Firmada. */
function submitPlanSignature(loanId, dni, signatureDataUrl, agree, meta) {
  return guard_('submitPlanSignature', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      throw new Error('Datos no encontrados. Verificá el ID del préstamo y tu DNI.');
    if (agree !== true) throw new Error('Debés aceptar los términos para firmar.');
    const req = findPlanRequestRow_(loanId, ['Aprobado', 'Firmado']);
    if (!req) throw new Error('No hay una solicitud de plan de pago aprobada para este préstamo.');
    const g = h => req.m[h] ? req.sh.getRange(req.row, req.m[h]).getValue() : '';
    const setv = (h, v) => { if (req.m[h]) req.sh.getRange(req.row, req.m[h]).setValue(v); };
    const setf = (h, f) => { if (req.m[h]) req.sh.getRange(req.row, req.m[h]).setFormula(f); };
    if (String(g('Estado de Firma')).toUpperCase() === SIGN.SIGNED)
      return { ok: true, message: 'El Acuerdo ya estaba firmado. Te enviamos una copia por correo.' };
    const m = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i.exec(String(signatureDataUrl || ''));
    if (!m) throw new Error('Falta la firma. Dibujá tu firma o cargá una imagen.');
    // GATE: el % inicial debe estar PAGADO (registrado) antes de activar el plan.
    const amtGate = planAmounts_(loan);
    if (!amtGate.downMet) throw new Error('Para activar el plan, primero debe estar pagado el ' + amtGate.pctInt + '%. Falta ' + fmtMoney_(amtGate.downShort) + '.');

    const now = new Date();
    const mime = m[1], ext = mime.split('/')[1].replace('+xml', '').replace('jpeg', 'jpg').replace('svg', 'png');
    const bytes = Utilities.base64Decode(m[2]);
    const method = (meta && meta.method === 'upload') ? 'upload' : 'draw';
    const ip = String((meta && meta.ip) || '').slice(0, 60);
    const devInfo = (meta && meta.device) || {};
    const deviceFp = deviceFingerprint_(devInfo);
    const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
    folder.createFile(Utilities.newBlob(bytes, mime, 'Firma Acuerdo ' + loanId + '.' + ext));

    const option = Math.floor(Number(g('Opción (1/4/6)')) || 4);
    const cadence = planCadence_(g('Cadencia'));
    // ACTIVA el plan: suspende mora, reemplaza cuotas (según la cadencia), marca "Plan desde".
    activarPlanPago_(loan, option, now, cadence);

    const sig = { dataUrl: signatureDataUrl, signedAt: now, method: method, userAgent: (meta && meta.userAgent) || '', ip: ip, hash: '', device: devInfo, deviceFp: deviceFp };
    const signedFile = savePdfToBorrower_(planAgreementHtml_(loan, { option: option, cadence: cadence }, sig), 'Acuerdo de Plan de Pago ' + loanId + ' - ' + loan.name + '.pdf', loan);
    let hashHex = ''; try { hashHex = hexDigest_(signedFile.getAs('application/pdf').getBytes()); } catch (e) { logError_('submitPlanSignature:hash', e); }

    try { setupFirmas_(getSS_()).appendRow([now, loanId, loan.name, loan.dni, loan.email, 'Plan de Pago (' + (method === 'upload' ? 'imagen' : 'dibujada') + ')',
      String((meta && meta.userAgent) || '').slice(0, 250), signedFile.getUrl(), ip || 'N/D', hashHex, loan.cuil || '', loan.direccion || '']); } catch (e) { logError_('submitPlanSignature:firmas', e); }
    try { logDeviceInfo_(loan, loan.clientId, { device: devInfo, deviceFp: deviceFp, ip: ip, userAgent: (meta && meta.userAgent) || '', method: 'plan de pago', signedAt: now, dwellMs: meta && meta.dwellMs, strokes: meta && meta.strokes, drawMs: meta && meta.drawMs }); } catch (e) { logError_('submitPlanSignature:dispositivos', e); }

    setv('Estado', 'Firmado'); setv('Estado de Firma', SIGN.SIGNED);
    if (req.m['Fecha de Firma']) req.sh.getRange(req.row, req.m['Fecha de Firma']).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
    setf('Acuerdo (PDF)', '=HYPERLINK("' + signedFile.getUrl() + '","Ver acuerdo firmado")');

    const optTxt = planOptTxt_(option, cadence);
    const lender = getSetting_('Correo del Prestamista');
    let emailOk = true;
    try {
      const att = signedFile.getAs('application/pdf');
      if (loan.email) sendBrandedEmail_(loan.email, 'Acuerdo de plan de pago firmado — ' + loanId,
        'Hola ' + loan.name + ', tu Acuerdo de plan de pago (' + optTxt + ') del préstamo ' + loanId + ' quedó firmado.',
        '<p>Hola ' + esc_(loan.name) + ',</p><p>Tu <b>Acuerdo de plan de pago</b> (' + esc_(optTxt) + ') del préstamo <b>' + esc_(loanId) + '</b> quedó firmado y <b>activo</b>. La mora quedó suspendida mientras cumplas el plan. Adjuntamos tu copia con el cronograma de cuotas.</p>',
        { attachments: [att] });
      if (lender) sendBrandedEmail_(lender, 'Plan de pago firmado — ' + loanId,
        loan.name + ' firmó el Acuerdo del préstamo ' + loanId + '. El plan quedó activo (cuotas reemplazadas).',
        '<p><b>' + esc_(loan.name) + '</b> firmó el Acuerdo del préstamo <b>' + esc_(loanId) + '</b> (' + esc_(optTxt) + '). El plan quedó <b>EN PLAN</b>: la mora se suspendió y se reemplazaron las cuotas.</p>',
        { attachments: [att] });
    } catch (e) { emailOk = false; logError_('submitPlanSignature:email', e); }
    try { getSS_().toast(loanId + ': plan firmado y activado.', '📋 Plan activo', 6); } catch (e) { }
    return { ok: true, message: emailOk ? '¡Acuerdo firmado! Tu plan está activo y la mora quedó suspendida. Te enviamos una copia por correo.' : '¡Acuerdo firmado! Tu plan está activo. (No pudimos enviar la copia por correo.)' };
  });
}

/** ACTIVA el plan: suspende la mora (Mora (ajuste)=0), reemplaza las cuotas por el plan elegido
 *  (sobre el Total a Pagar, SIN mora) y marca "Plan desde" (→ estado EN PLAN). */
function activarPlanPago_(loan, option, signDate, cadence) {
  const ss = getSS_();
  const n = (option === 1) ? 1 : (option === 6 ? 6 : 4);
  cadence = planCadence_(cadence);
  const loc = signStatusRow_(loan.loanId);
  if (loc) {
    const H = headerIndex_(loc.sh);
    const adjC = colByAny_(H, ['Mora (ajuste)']) || PB.MORA_ADJ;
    loc.sh.getRange(loc.row, adjC).setValue(0);  // suspende la mora mientras rija el plan
    const planC = colByAny_(H, ['Plan desde']) || PB.PLAN_SINCE;
    loc.sh.getRange(loc.row, planC).setValue(signDate).setNumberFormat('yyyy-mm-dd');
  }
  // Reemplaza las cuotas: el plan cubre el TOTAL A PAGAR (sin mora). Los pagos registrados (incluido
  // el pago inicial) se imputan a la cuota más antigua por la fórmula por cuota de cuotasForLoan_.
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || setupCuotas_(ss);
  deleteCuotasRowsForLoan_(cs, loan.loanId);
  cuotasForLoan_(cs, loan.loanId, loan.loanDate, round2_(loan.totalDue), n, null, { dates: planCuotaDates_(signDate, n, cadence) });
}

/** HTML del ACUERDO TRANSACCIONAL — Reconocimiento de deuda y plan de pagos (modelo del PDF de ejemplo). */
function planAgreementHtml_(loan, req, sig) {
  const lender = companyName_(), footer = getSetting_('Pie del Contrato');
  const signatories = ownerSignatories_();
  const option = (req && req.option) ? Math.floor(req.option) : 4;
  const n = (option === 1) ? 1 : (option === 6 ? 6 : 4);
  const cadence = planCadence_(req && req.cadence);
  const a = planAmounts_(loan);
  const signedAt = (sig && sig.signedAt instanceof Date) ? sig.signedAt : new Date();
  const dates = planCuotaDates_(signedAt, n, cadence);
  const baseCuota = round2_(a.total / n);
  let left = a.recorded; const cuotaRows = [];
  for (let k = 0; k < n; k++) {
    const monto = (k === n - 1) ? round2_(a.total - baseCuota * (n - 1)) : baseCuota;
    const ap = Math.min(left, monto); const neto = round2_(monto - ap); left = round2_(left - ap);
    cuotaRows.push({ n: k + 1, due: dates[k], monto: monto, neto: neto, ap: ap });
  }
  const optTxt = planOptTxt_(n, cadence);
  const kv = (k, v) => `<tr><td class="k">${k}</td><td>${v}</td></tr>`;
  const fill = v => (v && String(v).trim()) ? esc_(v) : '<span style="color:#9aa4ad">—</span>';
  const mono = '<span style="font-family:monospace;font-size:8.5pt;word-break:break-all">';
  const borrowerSig = (sig && sig.dataUrl) ? sig.dataUrl : '';
  const sigCell = (img, name, role) =>
    `<td>${img ? `<img src="${img}" alt="firma" style="max-height:44px;max-width:88%;display:block;margin:0 auto 1px">` : '<div style="height:40px"></div>'}` +
    `<div style="border-top:1px solid #333;padding-top:3px">${esc_(name)}<br><span style="font-size:8pt;color:#555">${esc_(role)}</span></div></td>`;
  const locSummary = s => { const d = (s && s.device) || {}, l = d.loc || {};
    if (l.source === 'GPS' && l.gps && l.gps.indexOf('GPS:') !== 0) return l.gps + (l.gpsAcc != null ? ' (±' + l.gpsAcc + ' m)' : '') + ' · GPS';
    if (l.ipText) return l.ipText + ' · IP'; if (l.source === 'desactivada') return 'No capturada'; return d.tz ? ('Zona horaria: ' + d.tz) : '—'; };
  const auditPanel = (sig && sig.dataUrl) ? `
    <h2>Constancia de firma electrónica</h2>
    <table>
      ${kv('Firmó', esc_(loan.name) + ' · DNI ' + esc_(loan.dni || '—'))}
      ${kv('Correo confirmado', esc_(loan.email || '—'))}
      ${kv('Identidad verificada con', 'DNI ' + esc_(loan.dni || '—') + ' y correo registrado; foto del DNI (frente y dorso, Anexo II)')}
      ${kv('Documento', 'Foto del DNI, frente y dorso — Anexo II')}
      ${kv('Fecha y hora', fmtDateTime_(sig.signedAt || new Date()))}
      ${kv('Cómo se firmó', sig.method === 'upload' ? 'Imagen de firma cargada' : 'Firma dibujada en pantalla')}
      ${kv('Dirección IP', fill(sig.ip))}
      ${kv('Ubicación', esc_(locSummary(sig)))}
      ${kv('Equipo / SO', fill((sig.device && sig.device.os) || ''))}
      ${kv('Navegador', fill((sig.device && sig.device.browser) || ''))}
      ${kv('Pantalla', fill((sig.device && sig.device.screen) ? (sig.device.screen + (sig.device.dpr ? (' · DPR ' + sig.device.dpr) : '')) : ''))}
      ${kv('Automatización detectada', (sig.device && sig.device.webdriver) ? 'Sí (revisar)' : 'No (firmante humano)')}
      ${kv('Dispositivo (completo)', esc_(String(sig.userAgent || '').slice(0, 200)))}
      ${kv('Huella de dispositivo (SHA-256)', sig.deviceFp ? (mono + esc_(sig.deviceFp) + '</span>') : '<span style="color:#9aa4ad">No disponible</span>')}
      ${kv('Huella del documento (SHA-256)', sig.hash ? (mono + esc_(sig.hash) + '</span>') : '<span style="color:#9aa4ad">Se registra en la hoja «Firmas».</span>')}
    </table>` : '';
  const cuotaTable = `<table><tr><th>Cuota</th><th>Vencimiento</th><th style="text-align:right">Monto</th></tr>` +
    cuotaRows.map(c => `<tr><td>${n === 1 ? 'Pago único' : (c.n + ' de ' + n)}</td><td>${fmtDate_(c.due)}</td><td style="text-align:right">${fmtMoney_(c.neto)}${c.ap > 0 ? ' <span style="color:#777">(−' + fmtMoney_(c.ap) + ' ya pagado)</span>' : ''}</td></tr>`).join('') + `</table>`;
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    @page{size:A4;margin:14mm 14mm}
    body{font-family:"Carlito","Calibri",Arial,sans-serif;color:#1d2630;margin:0;line-height:1.35;font-size:10pt}
    h1{text-align:center;font-size:15pt;border-bottom:2px solid #674ea7;padding-bottom:4px;color:#351c75;margin:2px 0 6px}
    h2{font-size:10.5pt;color:#674ea7;margin:12px 0 3px;border-bottom:1px solid #d8d0ec;padding-bottom:2px;text-transform:uppercase;letter-spacing:.4px}
    p{margin:0 0 5px} ol{margin:3px 0 5px 18px} table{width:100%;border-collapse:collapse;margin:5px 0;font-size:9.2pt} td,th{padding:4px 7px;border:1px solid #d8dee3;vertical-align:top}
    th{background:#f3f0fa;text-align:left} td.k{background:#f3f0fa;font-weight:bold;width:46%}
    .box{border:1px solid #674ea7;border-left:4px solid #674ea7;background:#f3f0fa;padding:7px 11px;margin:7px 0}
    .sign{margin-top:16px;width:100%} .sign td{border:none;text-align:center;padding:6px 10px;width:33%;vertical-align:bottom}
    .foot{margin-top:16px;padding-top:6px;border-top:1px solid #dfe4e8;font-size:8.6pt;color:#5c6871;font-style:italic}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>ACUERDO TRANSACCIONAL — RECONOCIMIENTO DE DEUDA Y PLAN DE PAGOS</h1>
    <p style="text-align:center;color:#3d4952">Celebrado el <b>${fmtDate_(signedAt)}</b> entre ${esc_(signatories[0].name)} y ${esc_(signatories[1].name)}, que operan como «${esc_(lender)}» (el Prestamista), y <b>${esc_(loan.name)}</b> — DNI ${esc_(loan.dni || '—')}${loan.cuil ? (' · CUIL ' + esc_(loan.cuil)) : ''} (la Deudora/el Deudor), respecto del préstamo <b>${esc_(loan.loanId)}</b>.</p>
    <h2>1. Reconocimiento de deuda</h2>
    <p>La parte deudora reconoce, libremente y sin reservas, que recibió el préstamo <b>${esc_(loan.loanId)}</b> por un capital de <b>${fmtMoney_(loan.principal)}</b> y que adeuda el total pactado (capital + interés) de <b>${fmtMoney_(a.total)}</b>, del cual ya pagó <b>${fmtMoney_(a.recorded)}</b>. Este reconocimiento no constituye un nuevo préstamo (art. 733 CCyCN).</p>
    <h2>2. Qué se adeuda y a qué renuncia el Prestamista</h2>
    <p>El préstamo venció impago y la parte deudora se encuentra en mora. El cargo punitorio alcanzó su tope contractual (<b>${fmtMoney_(a.mora)}</b>) y <b>dejó de devengarse</b>.</p>
    <p>El Prestamista acepta <b>SUSPENDER</b> el cargo punitorio de ${fmtMoney_(a.mora)} mientras este acuerdo se cumpla. La suspensión es <b>condicional y no implica condonación</b>: el punitorio no se extingue, queda en suspenso, y la cláusula 4 establece qué ocurre si se incumple. El monto bajo este acuerdo es, por lo tanto, de <b>${fmtMoney_(a.total)}</b> (capital + interés), sin gastos ni comisiones.</p>
    <h2>3. Cómo pagará — opción elegida: ${esc_(optTxt)}</h2>
    <div class="box">
      ${kv('Total del acuerdo', '<b>' + fmtMoney_(a.total) + '</b>')}
      ${kv('Pago inicial requerido (' + a.pctInt + '%)', fmtMoney_(a.required) + ' — pagado antes de activar el acuerdo')}
    </div>
    ${cuotaTable}
    <p>El vencimiento es la fecha pactada, aunque caiga fin de semana o feriado. La parte deudora puede pagar anticipadamente, total o parcialmente, sin penalidad. Los pagos se imputan primero a la cuota impaga más antigua.</p>
    <h2>4. Si se incumple un pago</h2>
    <p>La suspensión del cargo punitorio (cláusula 2) depende del cumplimiento. Si cualquier pago no se realiza íntegramente en su fecha: <b>(a)</b> la suspensión cesa automáticamente, sin necesidad de notificación; <b>(b)</b> el cargo punitorio de ${fmtMoney_(a.mora)} se restablece en forma íntegra y retroactiva; <b>(c)</b> la totalidad del saldo impago se torna exigible de inmediato. Este acuerdo no se renegocia por segunda vez.</p>
    <h2>5. Pagos y comunicación</h2>
    <p>Todo pago se realiza a la cuenta que el Prestamista indique; el pago a cualquier otra persona no cancela la deuda. La parte deudora envía el comprobante dentro de las 24 h y responde los mensajes del Prestamista dentro de las 48 h. Las partes se comunican por correo, WhatsApp o SMS, válidos también para notificaciones formales.</p>
    <h2>6. Firma y ley aplicable</h2>
    <p>Las partes firman electrónicamente y aceptan su validez. La parte deudora adopta como propia la firma de abajo y acepta el registro de firma del Prestamista —fecha y hora, IP, dispositivo y hash— como prueba. Este acuerdo se rige por la ley argentina.</p>
    <table class="sign"><tr>
      ${sigCell(borrowerSig, loan.name, 'Deudor/a · DNI ' + (loan.dni || ''))}
      ${sigCell(signatories[0].sig, signatories[0].name, 'Prestamista')}
      ${sigCell(signatories[1].sig, signatories[1].name, 'Prestamista')}
    </tr></table>
    ${auditPanel}
    <p class="foot">${esc_(footer || '')} Este acuerdo tiene efecto desde que se firma. La parte deudora confirma que lo leyó, lo acepta y lo firma electrónicamente.</p>
  </body></html>`;
}

function signingPageHtml_(loanId, mode) {
  const freeze = (mode === 'freeze'), plan = (mode === 'plan');
  const TITLE = plan ? 'Firmar Acuerdo de Plan de Pago' : (freeze ? 'Firmar Congelamiento' : 'Firmar Contrato');
  const FN_GET = plan ? 'getPlanContractForSigning' : (freeze ? 'getFreezeContractForSigning' : 'getContractForSigning');
  const FN_SUB = plan ? 'submitPlanSignature' : (freeze ? 'submitFreezeSignature' : 'submitSignature');
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
    <div class="card">${brandHeaderHtml_()}<h1>${TITLE}</h1>

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
      var PAGE_OPEN = Date.now(), CAPTURE_LOC = true, STROKES = 0, DRAW_START = 0, DRAW_END = 0;
      function $(id){return document.getElementById(id);}
      // Captura de IP del firmante (mejor esfuerzo; si el servicio no responde, queda vacío).
      try{ fetch('https://api.ipify.org?format=json').then(function(r){return r.json();}).then(function(j){ CLIENT_IP=(j&&j.ip)||''; }).catch(function(){}); }catch(e){}
      function unlock(){
        DNI = $('dni').value; $('gerr').style.display='none';
        if(!$('loan').value || !DNI){ gerr('Ingrese el ID y el DNI.'); return; }
        google.script.run.withSuccessHandler(function(r){
          if(!r || !r.ok){ gerr((r&&r.msg)||'No se pudo verificar.'); return; }
          if(r.signed){ $('gate').classList.add('hide'); showDone(r.msg); return; }
          CAPTURE_LOC = (r.captureLocation !== false);
          $('terms').srcdoc = r.contractHtml; PAGE_OPEN = Date.now();
          $('gate').classList.add('hide'); $('signview').classList.remove('hide'); setupCanvas();
        }).withFailureHandler(function(e){ gerr(e.message||e); }).${FN_GET}($('loan').value, DNI);
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
        cv.onpointerdown=function(e){ drawing=true; drew=true; STROKES++; if(!DRAW_START) DRAW_START=Date.now(); var p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); cv.setPointerCapture(e.pointerId); };
        cv.onpointermove=function(e){ if(!drawing) return; var p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); };
        cv.onpointerup=function(){ drawing=false; DRAW_END=Date.now(); }; cv.onpointerleave=function(){ drawing=false; };
      }
      function pos(e){ var r=cv.getBoundingClientRect(); return {x:e.clientX-r.left, y:e.clientY-r.top}; }
      function clearCanvas(){ if(ctx){ ctx.clearRect(0,0,cv.width,cv.height); drew=false; STROKES=0; DRAW_START=0; DRAW_END=0; } }
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
      function withTimeout(p, ms){ return Promise.race([p, new Promise(function(res){ setTimeout(res, ms); })]); }
      // Reúne la huella de dispositivo, señales de fraude y ubicación. SIEMPRE resuelve
      // (best-effort): si algo se deniega, agota el tiempo o falla, se continúa igual.
      function collectDeviceInfo(){
        var dev = { loc: { source:'desactivada', ipText:'', gps:'', gpsAcc:null, gpsStatus:'' } };
        try {
          var s = window.screen||{}, n = navigator||{};
          dev.screen = (s.width||'')+'x'+(s.height||'');
          dev.color = s.colorDepth||'';
          dev.dpr = window.devicePixelRatio||'';
          dev.touch = (n.maxTouchPoints!=null)?n.maxTouchPoints:'';
          dev.cores = n.hardwareConcurrency||'';
          dev.memory = n.deviceMemory||'';
          dev.langs = (n.languages&&n.languages.join)?n.languages.join(','):(n.language||'');
          dev.webdriver = !!n.webdriver;
          try { dev.tz = Intl.DateTimeFormat().resolvedOptions().timeZone||''; } catch(e){ dev.tz=''; }
          dev.tzOffset = (new Date()).getTimezoneOffset();
          dev.net = (n.connection&&n.connection.effectiveType)||'';
          dev.os = n.platform||''; dev.browser=''; dev.model=''; dev.arch='';
          dev.userAgent = n.userAgent||'';
        } catch(e){}
        var jobs = [];
        // Client Hints (solo Chromium): SO/modelo/navegador/arquitectura de alta entropía.
        try {
          if (navigator.userAgentData && navigator.userAgentData.getHighEntropyValues) {
            jobs.push(navigator.userAgentData.getHighEntropyValues(
              ['platform','platformVersion','architecture','bitness','model','fullVersionList'])
              .then(function(h){ if(!h) return;
                if(h.platform) dev.os = h.platform + (h.platformVersion?(' '+h.platformVersion):'');
                if(h.model) dev.model = h.model;
                if(h.architecture) dev.arch = h.architecture + (h.bitness?('/'+h.bitness):'');
                var b = h.fullVersionList && h.fullVersionList.filter(function(x){ return !/Not.?A.?Brand/i.test(x.brand); });
                if(b && b.length) dev.browser = b[b.length-1].brand+' '+b[b.length-1].version;
              }).catch(function(){}));
          }
        } catch(e){}
        // Ubicación: escalera de respaldo (best-effort), solo si está habilitada.
        if (CAPTURE_LOC) {
          // Nivel 2: geo-IP (sin permiso).
          jobs.push(withTimeout(fetch('https://ipwho.is/').then(function(r){return r.json();}).then(function(j){
            if(j && j.success!==false){
              var parts=[j.city,j.region,j.country].filter(Boolean).join(', ');
              var isp=(j.connection&&j.connection.isp)||j.isp||'';
              dev.loc.ipText = parts + (isp?(' · '+isp):'');
            }
          }).catch(function(){}), 4000));
          // Nivel 1: GPS (con permiso del navegador; timeout 5 s; nunca bloquea).
          jobs.push(new Promise(function(res){
            if(!navigator.geolocation){ dev.loc.gpsStatus='no disponible'; return res(); }
            navigator.geolocation.getCurrentPosition(function(p){
              dev.loc.gps = p.coords.latitude.toFixed(5)+','+p.coords.longitude.toFixed(5);
              dev.loc.gpsAcc = Math.round(p.coords.accuracy); dev.loc.gpsStatus='ok'; res();
            }, function(err){
              dev.loc.gpsStatus = (err&&err.code===1)?'no consentida':'no disponible';
              dev.loc.gps = 'GPS: '+dev.loc.gpsStatus; res();
            }, { timeout:5000, maximumAge:600000, enableHighAccuracy:false });
          }));
        }
        return Promise.all(jobs).then(function(){
          // Origen = mejor nivel disponible (GPS > IP > zona horaria).
          if(!CAPTURE_LOC) dev.loc.source='desactivada';
          else if(dev.loc.gpsStatus==='ok') dev.loc.source='GPS';
          else if(dev.loc.ipText) dev.loc.source='IP';
          else dev.loc.source='zona horaria';
          return dev;
        }).catch(function(){ return dev; });
      }
      function sign(){
        var out=$('out'); out.style.display='none';
        if(!$('agree').checked){ msg('Debe aceptar los términos para firmar.', true); return; }
        var data='';
        if(TAB==='draw'){ if(!drew){ msg('Dibuje su firma primero.', true); return; } data=cv.toDataURL('image/png'); }
        else { if(!uploaded){ msg('Cargue una imagen de su firma primero.', true); return; } data=uploaded; }
        $('go').disabled=true;
        $('signview').classList.add('hide'); $('loadingView').classList.remove('hide'); startProgress();
        collectDeviceInfo().then(function(device){
          var meta = { userAgent: navigator.userAgent, method: TAB, ip: CLIENT_IP, device: device,
            dwellMs: Date.now()-PAGE_OPEN, strokes: STROKES, drawMs: (DRAW_END>DRAW_START ? DRAW_END-DRAW_START : 0) };
          google.script.run.withSuccessHandler(function(r){
            if(r && r.ok){ finishProgress(function(){ $('loadingView').classList.add('hide'); showDone(r.message); }); }
            else { stopProgress(); $('loadingView').classList.add('hide'); $('signview').classList.remove('hide'); $('go').disabled=false; msg((r&&r.message)||'No se pudo firmar.', true); }
          }).withFailureHandler(function(e){
            stopProgress(); $('loadingView').classList.add('hide'); $('signview').classList.remove('hide'); $('go').disabled=false; msg(e.message||e, true);
          }).${FN_SUB}(LOAN_ID, DNI, data, true, meta);
        });
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
        const touched = Object.keys(ids);
        touched.forEach(id => { recalcLoanPayments_(id); updateLoanOutstanding_(id); });
        // Un pago puede saldar el capital pero dejar mora pendiente: reconstruir "Pagos Atrasados"
        // para que "Recargo por Mora (acum.)" / "Saldo con Mora" (VLOOKUP) y el Estado queden al día.
        if (touched.length) { try { rebuildLateSheet_(getSS_()); } catch (err) { logError_('onEdit:PAGOS:rebuildLate', err); } }
        // Aviso de pago por WhatsApp: arma el enlace por fila (confirma el pago, el saldo resultante
        // y el vencimiento). Se calcula tras recalcular el "Saldo Posterior" de la(s) fila(s).
        try {
          SpreadsheetApp.flush();
          const waCol = headerMap_(sh)['Avisar pago 📲'] || PP.WA_PAYMENT;
          for (let row = r0; row <= rN; row++) {
            const id = String(sh.getRange(row, PP.LOAN_ID).getValue()).trim();
            const amt = Number(sh.getRange(row, PP.AMOUNT).getValue()) || 0;
            const cell = sh.getRange(row, waCol);
            if (!id || !amt) { cell.clearContent(); continue; }
            const loan = findLoanById_(id);
            const f = loan ? paymentWaFormula_(loan, sh.getRange(row, PP.DATE).getValue(), amt) : '';
            if (f) cell.setFormula(f); else cell.setValue(loan ? 'sin teléfono' : '');
          }
        } catch (err) { logError_('onEdit:PAGOS:waPago', err); }
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
      // Casilla "Generar enlace 🔗": genera el enlace de firma de ESE préstamo, lo escribe en
      // "Enlace de Firma" (clicable) y lo muestra para copiar. Se destilda solo.
      const genLinkCol = bhm['Generar enlace 🔗'];
      const linkCol = bhm['Enlace de Firma'];
      if (genLinkCol && genLinkCol >= c0 && genLinkCol <= cN) {
        const base = (typeof webAppBaseUrl_ === 'function') ? webAppBaseUrl_() : '';
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, genLinkCol).getValue() !== true) continue;
          sh.getRange(row, genLinkCol).setValue(false);
          const loanId = String(sh.getRange(row, bLoanCol).getValue()).trim();
          if (!loanId) continue;
          if (!base) { getSS_().toast('Configurá la "URL de la app web (enlaces a clientes)" en Configuración (o publicá la app web) para generar enlaces.', '⚠ Falta la URL de la app', 8); continue; }
          try {
            const link = signingLink_(loanId);
            if (linkCol) sh.getRange(row, linkCol).setFormula('=HYPERLINK("' + link + '","' + link + '")');
            sh.getRange(row, genLinkCol).setNote('Enlace de firma generado el ' + fmtDate_(new Date()) + ':\n' + link);
            getSS_().toast('Enlace de firma de ' + loanId + ': ' + link, '🔗 Enlace generado', 10);
          } catch (err) { logError_('onEdit:BORROWERS:genEnlace', err); getSS_().toast(err.message || String(err), '⚠ No se pudo generar', 8); }
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
      // Casilla "Plan de pago 📅 (3 cuotas)": reestructura la deuda congelada en 3 cuotas
      // mensuales y envía el plan por correo. La columna "Tope de mora" marca los candidatos.
      if (LATE_PLAN_COL >= c0 && LATE_PLAN_COL <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, LATE_PLAN_COL).getValue() !== true) continue;
          sh.getRange(row, LATE_PLAN_COL).setValue(false);
          try { planPagoDesdeAtrasos_(sh, row); }
          catch (err) { logError_('onEdit:LATE:planPago', err); getSS_().toast(err.message || String(err), '⚠ No se pudo reestructurar', 8); }
        }
      }
      // Casilla "PDF de cuotas 🧾": genera el PDF con el cronograma de cuotas del préstamo.
      if (LATE_DOC_COL >= c0 && LATE_DOC_COL <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, LATE_DOC_COL).getValue() !== true) continue;
          sh.getRange(row, LATE_DOC_COL).setValue(false);
          try { docCuotasDesdeAtrasos_(sh, row); }
          catch (err) { logError_('onEdit:LATE:docCuotas', err); getSS_().toast(err.message || String(err), '⚠ No se pudo generar el PDF', 8); }
        }
      }
    } else if (name === CFG.SHEETS.FREEZE_REQ) {
      // Casillas "Aprobar ✔" / "Rechazar ✖" de una solicitud de congelamiento.
      const mF = headerMap_(sh), appC = mF['Aprobar ✔'] || FR_APPROVE_COL, denC = mF['Rechazar ✖'] || FR_DENY_COL;
      for (let row = rN; row >= Math.max(r0, 2); row--) {
        if (appC >= c0 && appC <= cN && sh.getRange(row, appC).getValue() === true) {
          sh.getRange(row, appC).setValue(false);
          try { approveFreezeRequest_(sh, row); } catch (err) { logError_('onEdit:FREEZE:approve', err); getSS_().toast(err.message || String(err), '⚠ No se pudo aprobar', 8); }
        } else if (denC >= c0 && denC <= cN && sh.getRange(row, denC).getValue() === true) {
          sh.getRange(row, denC).setValue(false);
          try { denyFreezeRequest_(sh, row); } catch (err) { logError_('onEdit:FREEZE:deny', err); getSS_().toast(err.message || String(err), '⚠ No se pudo rechazar', 8); }
        }
      }
    } else if (name === CFG.SHEETS.PLAN_REQ) {
      // Casillas "Aprobar ✔" / "Rechazar ✖" de una solicitud de plan de pago.
      const mP = headerMap_(sh), appC = mP['Aprobar ✔'] || PR_APPROVE_COL, denC = mP['Rechazar ✖'] || PR_DENY_COL;
      for (let row = rN; row >= Math.max(r0, 2); row--) {
        if (appC >= c0 && appC <= cN && sh.getRange(row, appC).getValue() === true) {
          sh.getRange(row, appC).setValue(false);
          try { approvePlanRequest_(sh, row); } catch (err) { logError_('onEdit:PLAN_REQ:approve', err); getSS_().toast(err.message || String(err), '⚠ No se pudo aprobar', 8); }
        } else if (denC >= c0 && denC <= cN && sh.getRange(row, denC).getValue() === true) {
          sh.getRange(row, denC).setValue(false);
          try { denyPlanRequest_(sh, row); } catch (err) { logError_('onEdit:PLAN_REQ:deny', err); getSS_().toast(err.message || String(err), '⚠ No se pudo rechazar', 8); }
        }
      }
    } else if (name === CFG.SHEETS.PLAN) {
      // Casilla "Plan de pago 📅 (3 cuotas)": REGENERA el plan (= "Regenerar cuotas": deuda
      // congelada a hoy, 1.ª cuota la semana próxima, SIN correo). Sella "Plan enviado".
      if (PLAN_PLAN_COL >= c0 && PLAN_PLAN_COL <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, PLAN_PLAN_COL).getValue() !== true) continue;
          sh.getRange(row, PLAN_PLAN_COL).setValue(false);
          const loanId = String(sh.getRange(row, 1).getValue()).trim();
          if (!loanId) continue;
          try {
            const out = regenerarCuotasPrestamo_(loanId, 3);
            sh.getRange(row, PLAN_PLAN_DONE_COL).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
            if (out) getSS_().toast(loanId + ': plan regenerado en ' + out.plan.length + ' cuotas desde ' + fmtDate_(out.plan[0].due) + '.', '📅 Plan regenerado', 6);
          } catch (err) { logError_('onEdit:PLAN:regenerar', err); getSS_().toast(err.message || String(err), '⚠ No se pudo regenerar el plan', 8); }
        }
      }
      // Casilla "PDF de cuotas 🧾": genera el PDF del cronograma y deja el enlace en "Enlace PDF".
      if (PLAN_DOC_COL >= c0 && PLAN_DOC_COL <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, PLAN_DOC_COL).getValue() !== true) continue;
          sh.getRange(row, PLAN_DOC_COL).setValue(false);
          const loanId = String(sh.getRange(row, 1).getValue()).trim();
          if (!loanId) continue;
          try {
            const f = generateCuotasPdfFile_(loanId);
            if (f) {
              sh.getRange(row, PLAN_LINK_COL).setFormula('=HYPERLINK("' + f.getUrl() + '","📄 Ver PDF")');
              sh.getRange(row, PLAN_DOC_DONE_COL).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
            }
            getSS_().toast('PDF de cuotas de ' + loanId + ' generado (ver "Enlace PDF").', '🧾 PDF de cuotas', 6);
          } catch (err) { logError_('onEdit:PLAN:docCuotas', err); getSS_().toast(err.message || String(err), '⚠ No se pudo generar el PDF', 8); }
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
      // Congelados: la mora está detenida → no enviamos recordatorios ni avisos de mora.
      if (status === ST.PAID || status === ST.FROZEN || outstanding <= 0) continue;
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

/**
 * Reafirma el "Saldo Posterior" (con mora) en TODAS las filas de pago con ID Préstamo.
 * Resuelve las columnas por NOMBRE de encabezado (robusto al layout migrado). Devuelve la
 * cantidad de filas actualizadas. Fuente de la fórmula: paymentBalanceFormula_.
 */
function writePaymentBalanceFormulas_(ss) {
  ss = ss || getSS_();
  const sh = ss.getSheetByName(CFG.SHEETS.PAYMENTS); if (!sh) return 0;
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID;
  const balC = colByAny_(H, ['Saldo Posterior']) || PP.BALANCE;
  const last = sh.getLastRow(); if (last < 2) return 0;
  const ids = sh.getRange(2, idC, last - 1, 1).getValues();
  let n = 0;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === '') continue;
    const r = i + 2;
    sh.getRange(r, balC).setFormula(paymentBalanceFormula_(r));
    n++;
  }
  return n;
}
/**
 * Menú: recalcula el "Saldo Posterior" de la hoja "Pagos" para que incluya la MORA (monto real
 * adeudado del préstamo tras cada pago). Acción LIGERA e independiente — no toca Resumen, Panel,
 * Estadísticas ni otras hojas, así que no afecta otras funcionalidades.
 */
function actualizarSaldoPosteriorPagos() {
  guard_('actualizarSaldoPosteriorPagos', function () {
    const ss = getSS_();
    const n = writePaymentBalanceFormulas_(ss);
    SpreadsheetApp.flush();
    let nWa = 0; try { nWa = fillPaymentWaLinks_(ss); } catch (e) { logError_('actualizarSaldoPosteriorPagos:wa', e); }
    try { ss.toast(n + ' fila(s): «Saldo Posterior» (con mora) + «Avisar pago 📲» (' + nWa + ') actualizados en «Pagos».', '💳 Pagos · saldo', 6); } catch (e) {}
  });
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
  const key = String(loanId).trim(); // compara ambos lados sin espacios (V-17)
  data.forEach(r => { const amt = parseMoney_(r[2]); if (String(r[0]).trim() === key && amt !== 0) out.push({ date: r[1] instanceof Date ? r[1] : new Date(), amount: amt }); });
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
// Caché por ejecución de toda la hoja "Configuración" (clave→valor; gana la primera aparición,
// igual que antes). getSetting_ se llama muchísimo (p. ej. graduationLevelFor_ por cada cliente en
// "Niveles de Clientes"); leer la hoja una sola vez acelera mucho esas listas. Los que ESCRIBEN
// ajustes (setSetting_, setSettingValue_, setupSettings_) invalidan el caché.
let _settingsCache = null;
function invalidateSettingsCache_() { _settingsCache = null; }
function getSetting_(key) {
  if (!_settingsCache) {
    _settingsCache = {};
    const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
    if (sh && sh.getLastRow() >= 2) {
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      for (const r of data) { const k = String(r[0]).trim(); if (k && !(k in _settingsCache)) _settingsCache[k] = r[1]; }
    }
  }
  return (key in _settingsCache) ? _settingsCache[key] : '';
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
function ccContains_(range, text, color) { return SpreadsheetApp.newConditionalFormatRule().whenTextContains(text).setBackground(color).setRanges([range]).build(); }
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
/**
 * Convierte a número un importe que puede venir como número real o como TEXTO
 * (ej. "$262.500,00", "262,500.00", "'262500"). Tolera separadores es-AR (miles ".",
 * decimales ",") y en-US. Devuelve 0 si no hay número. Evita que un monto tipeado
 * como texto quede fuera del SUMIF / del motor de mora (pago "perdido").
 */
function parseMoney_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v == null || v instanceof Date) return 0;
  let s = String(v).trim().replace(/[^\d.,\-]/g, ''); // deja dígitos, separadores y signo
  if (!s || s === '-') return 0;
  const decPos = Math.max(s.lastIndexOf('.'), s.lastIndexOf(','));
  if (decPos > -1 && /^[.,]\d{1,2}$/.test(s.slice(decPos))) {   // el último separador son centavos
    s = s.slice(0, decPos).replace(/[.,]/g, '') + '.' + s.slice(decPos + 1);
  } else {
    s = s.replace(/[.,]/g, '');                                 // todos son separadores de miles
  }
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
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
    // Congelado: la mora está detenida → no corresponde avisar. No es error (se puede tildar
    // la casilla en "Pagos Atrasados" sin querer): devolvemos un mensaje y no enviamos nada.
    if (frozenInfo_(loanId).since instanceof Date) {
      return { date: '', message: 'Préstamo ' + loanId + ' CONGELADO — no se envía aviso de mora.' };
    }
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

/**
 * PRUEBA — crea en "Prestatarios" un préstamo que YA está al TOPE DE MORA (100%),
 * para probar el flujo de PLAN DE PAGO sin esperar ~20 días de atraso. Retrodata la
 * "Fecha Préstamo" lo necesario para que la mora llegue al tope, arma las cuotas y
 * resuelve Nombre/DNI del cliente (default C-0028 = Denis Bell). El préstamo queda
 * FIRMADO (activo) y VENCIDO. Corré esto desde el editor de Apps Script (Ejecutar ▸
 * crearPrestamoPruebaMora100) — todos los argumentos tienen valor por defecto.
 *
 * Cada corrida crea UN préstamo nuevo (L-#### siguiente). Para limpiar: borrá su fila
 * en "Prestatarios" y sus cuotas en "Cuotas" (se reconoce por la nota 🧪 en el ID).
 *
 * @param {string} [clienteId='C-0028'] ID del cliente en "Clientes".
 * @param {number} [capital=200000]     Capital del préstamo (ARS).
 * @param {number} [plazoDias=30]       Plazo elegido (15/30/60/90); se ajusta al tramo.
 * @return {string} ID del préstamo creado.
 */
function crearPrestamoPruebaMora100(clienteId, capital, plazoDias) {
  return guard_('crearPrestamoPruebaMora100', function () {
    clienteId = String(clienteId || 'C-0028').trim();
    capital = Number(capital) || 200000;
    plazoDias = Number(plazoDias) || 30;

    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    if (!bs) throw new Error('Falta la hoja "' + CFG.SHEETS.BORROWERS + '".');

    // El cliente debe existir (Nombre/DNI se resuelven de "Clientes" por ID Cliente).
    const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
    let cliName = '';
    if (cl && cl.getLastRow() >= 2) {
      const cv = cl.getRange(2, PC.CLIENT_ID, cl.getLastRow() - 1, PC.NAME).getValues();
      for (let i = 0; i < cv.length; i++) if (String(cv[i][PC.CLIENT_ID - 1]).trim() === clienteId) { cliName = String(cv[i][PC.NAME - 1]); break; }
    }
    if (cl && cl.getLastRow() >= 2 && !cliName) throw new Error('No encontré el cliente ' + clienteId + ' en "' + CFG.SHEETS.CLIENTS + '". Verificá el ID.');

    // Tasa/cuotas/plazo reales según el monto (misma fuente que una aprobación normal).
    const ti = resolveTerm_(capital, plazoDias);
    const total = round2_(capital * (1 + ti.rate));
    // Para llegar al tope (100% al 5%/día) hacen falta ~20 días de atraso. Dejamos un
    // margen extra para quedar holgadamente en el tope. Vencimiento = Fecha + plazo.
    const overdueDays = Math.ceil(moraCapFrac_() / 0.05) + 5;
    const loanDate = addDays_(startOfDay_(new Date()), -(Number(ti.days) + overdueDays));

    // Alta de la fila (mismo patrón de columnas que approveApplicantV2_).
    const H = headerIndex_(bs);
    const idCol = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
    const cliCol = colByAny_(H, ['ID Cliente']) || PB.CLIENT_ID;
    const capCol = colByAny_(H, ['Capital']) || PB.PRINCIPAL;
    const plzCol = colByAny_(H, ['Plazo (días)', 'Plazo (dias)', 'Plazo']) || PB.TERM;
    const fecCol = colByAny_(H, ['Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo']) || PB.LOAN_DATE;
    const firmaCol = colByAny_(H, ['Estado de Firma']) || PB.SIGN_STATUS;
    const nameCol = colByAny_(H, ['Nombre']);
    const dniCol = colByAny_(H, ['DNI']);

    const loanId = nextLoanId_();
    const tRow = insertBorrowerRow_(bs, idCol);
    bs.getRange(tRow, idCol).setValue(loanId)
      .setNote('🧪 Préstamo de PRUEBA (mora 100%) creado el ' + fmtDate_(new Date()) + '. Borralo (fila + cuotas) cuando termines de probar.');
    bs.getRange(tRow, cliCol).setValue(clienteId);
    bs.getRange(tRow, capCol).setValue(capital);
    bs.getRange(tRow, plzCol).setValue(ti.days);
    bs.getRange(tRow, fecCol).setValue(loanDate).setNumberFormat('yyyy-mm-dd');
    if (firmaCol) bs.getRange(tRow, firmaCol).setValue(SIGN.SIGNED); // préstamo activo (ya firmado)

    // Fórmulas calculadas de la fila (Tasa/Venc/Interés/Total/Pagado/Saldo/Mora/Estado/Enlace…).
    writeBorrowerRowFormulas_(bs, tRow);
    // Estado con reconocimiento de CONGELADO/EN PLAN si ya existen esas columnas
    // (las agrega «🆕 Aplicar novedades»): así el estado vira a EN PLAN al activar el plan.
    if (colByAny_(H, ['Plan desde']) && colByAny_(H, ['Congelado el'])) {
      bs.getRange(tRow, colByAny_(H, ['Estado']) || PB.STATE).setFormula(borrowerStateFormula_(tRow));
    }
    // Nombre/DNI resueltos de "Clientes" por ID Cliente (modelo normalizado).
    if ((nameCol || dniCol) && cl) {
      const CH = headerIndex_(cl), cliL = a1col_(cliCol), CSN = CFG.SHEETS.CLIENTS;
      const cIdL = a1col_(colByAny_(CH, ['ID Cliente']) || 1);
      if (nameCol) { const cNameL = a1col_(colByAny_(CH, ['Nombre']) || 2);
        bs.getRange(tRow, nameCol).setFormula(`=IFERROR(INDEX('${CSN}'!$${cNameL}:$${cNameL},MATCH($${cliL}${tRow},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`); }
      if (dniCol) { const cDniL = a1col_(colByAny_(CH, ['DNI']) || 4);
        bs.getRange(tRow, dniCol).setFormula(`=IFERROR(INDEX('${CSN}'!$${cDniL}:$${cDniL},MATCH($${cliL}${tRow},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`); }
    }

    // Cronograma de cuotas (misma lógica que una aprobación): 1 cuota (tramo chico) o N mensuales (grande).
    // Se usa la Fecha Préstamo retrodatada → la(s) cuota(s) también quedan VENCIDAS.
    const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || (typeof setupCuotas_ === 'function' ? setupCuotas_(ss) : null);
    if (cs) {
      const dueDays = ti.cadence === 'monthly' ? Array.from({ length: ti.cuotas }, (_, i) => (i + 1) * 30) : [ti.days];
      cuotasForLoan_(cs, loanId, loanDate, total, ti.cuotas, dueDays);
    }
    SpreadsheetApp.flush();

    // Refrescá "Pagos Atrasados" para que la mora aparezca en "Prestatarios" y el
    // préstamo quede listo con su enlace «Pedir plan de pago 📋».
    try { if (typeof rebuildLateSheet_ === 'function') rebuildLateSheet_(); } catch (e) { logError_('crearPrestamoPruebaMora100:late', e); }

    const mora = round2_(total * moraCapFrac_()); // al tope
    const msg = '🧪 Préstamo de PRUEBA ' + loanId + ' creado para ' + (cliName || clienteId) +
      ' — Capital ' + fmtMoney_(capital) + ', Total ' + fmtMoney_(total) +
      ', atraso ~' + overdueDays + ' días, mora ' + fmtMoney_(mora) + ' (tope ' + Math.round(moraCapFrac_() * 100) + '%). ' +
      'Ya califica para PLAN DE PAGO (y para congelamiento).';
    try { ss.toast(msg, '🧪 Préstamo de prueba', 12); } catch (e) {}
    Logger.log(msg);
    return loanId;
  });
}

/* ============================================================================
 * PRUEBA — ESCENARIOS DE PAGO Y ESTADOS DE PRÉSTAMO (para probar "Avisar pago 📲")
 * Crea clientes + préstamos de prueba que cubren TODOS los estados: SALDADO, debe al
 * día, debe atrasado (con mora), en CUOTAS (una paga / otra parcial / resto pendiente),
 * CONGELADO y EN PLAN (mora suspendida). Menú ▸ ⚙ Configuración y mantenimiento ▸
 * 🧪 Datos de prueba. Idempotente: borra los de prueba anteriores antes de recrearlos.
 * ==========================================================================*/
const SCENARIO_TEST_PREFIX = '🧪 PRUEBA — ';

/** Menú: confirma y crea los préstamos de prueba de todos los escenarios de pago. */
function seedPaymentScenariosConfirm() {
  return guard_('seedPaymentScenariosConfirm', function () {
    const ui = SpreadsheetApp.getUi();
    const r = ui.alert('Crear escenarios de pago de prueba',
      'Se crearán 6 clientes y préstamos de prueba (prefijo «' + SCENARIO_TEST_PREFIX + '») que cubren todos los estados:\n' +
      '• SALDADO  • Debe (al día)  • Debe (atrasado + mora)\n• En cuotas  • CONGELADO  • EN PLAN (mora suspendida)\n\n' +
      'Primero se eliminan los de prueba anteriores (si los hay). ¿Continuar?',
      ui.ButtonSet.OK_CANCEL);
    if (r !== ui.Button.OK) return;
    const res = seedPaymentScenarios_();
    ui.alert('Escenarios de prueba creados',
      res.resumen + '\n\nAbrí la hoja «Pagos» y mirá la columna «Avisar pago 📲» de cada fila.',
      ui.ButtonSet.OK);
  });
}

/** Crea (idempotente) los préstamos de prueba de todos los escenarios. Devuelve {loans, resumen}. */
function seedPaymentScenarios_() {
  const ss = getSS_();
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (!bs || !ps) throw new Error('Faltan las hojas «Prestatarios» y/o «Pagos».');
  const BH = headerIndex_(bs);
  if (!colByAny_(BH, ['Congelado el']) || !colByAny_(BH, ['Plan desde']))
    throw new Error('Faltan las columnas de congelamiento/plan en «Prestatarios». Ejecutá primero 🆕 Aplicar novedades.');

  removePaymentScenarios_(); // limpia los de prueba anteriores para no duplicar

  const P = SCENARIO_TEST_PREFIX, today = startOfDay_(new Date());
  const scenarios = [
    { estado: 'SALDADO', cliente: { name: P + 'Saldado', email: 'prueba.saldado@example.com', dni: '90000001', phone: '1150000001' },
      capital: 100000, plazoDias: 30, fechaOffsetDays: -20, cuotas: { n: 1 }, pagos: [{ offsetDays: -1, monto: 150000 }] },
    { estado: 'Debe (al día)', cliente: { name: P + 'Debe al dia', email: 'prueba.aldia@example.com', dni: '90000002', phone: '1150000002' },
      capital: 200000, plazoDias: 30, fechaOffsetDays: -5, cuotas: { n: 1 }, pagos: [{ offsetDays: -1, monto: 100000 }] },
    { estado: 'Debe (atrasado + mora)', cliente: { name: P + 'Debe atrasado', email: 'prueba.atraso@example.com', dni: '90000003', phone: '1150000003' },
      capital: 200000, plazoDias: 30, fechaOffsetDays: -35, cuotas: { n: 1 }, pagos: [{ offsetDays: -20, monto: 50000 }] },
    { estado: 'En cuotas', cliente: { name: P + 'En cuotas', email: 'prueba.cuotas@example.com', dni: '90000004', phone: '1150000004' },
      capital: 400000, plazoDias: 90, rate: 1.0, fechaOffsetDays: -20, cuotas: { n: 4, dueDays: [30, 60, 90, 120] }, pagos: [{ offsetDays: -2, monto: 300000 }] },
    { estado: 'CONGELADO', cliente: { name: P + 'Congelado', email: 'prueba.congelado@example.com', dni: '90000005', phone: '1150000005' },
      capital: 250000, plazoDias: 30, fechaOffsetDays: -36, congeladoOffset: 0, cuotas: { n: 1 }, pagos: [{ offsetDays: -1, monto: 187500 }] },
    { estado: 'EN PLAN (mora suspendida)', cliente: { name: P + 'En plan', email: 'prueba.plan@example.com', dni: '90000006', phone: '1150000006' },
      capital: 200000, plazoDias: 30, fechaOffsetDays: -40, planOffset: 0, moraAjuste: 0,
      cuotas: { n: 3, dates: [addDays_(today, 30), addDays_(today, 60), addDays_(today, 90)] }, pagos: [{ offsetDays: -1, monto: 60000 }] },
  ];

  const loans = [];
  scenarios.forEach(function (sc) { const r = seedScenarioLoan_(ss, sc); loans.push(sc.estado + ' → ' + r.loanId); });

  SpreadsheetApp.flush();
  try { rebuildLateSheet_(ss); } catch (e) { logError_('seedPaymentScenarios_:late', e); }
  try { fillPaymentWaLinks_(ss); } catch (e) { logError_('seedPaymentScenarios_:wa', e); }
  try { if (typeof buildClientLevels === 'function') buildClientLevels({ silent: true }); } catch (e) {}

  const resumen = loans.length + ' escenario(s) creados:\n• ' + loans.join('\n• ');
  try { ss.toast(loans.length + ' escenarios de prueba creados.', '🧪 Escenarios de pago', 8); } catch (e) {}
  Logger.log(resumen);
  return { loans: loans, resumen: resumen };
}

/** Crea UN préstamo de prueba (+ cuotas + pagos) según opts. Devuelve {loanId, clienteId, total}. */
function seedScenarioLoan_(ss, opts) {
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || (typeof setupCuotas_ === 'function' ? setupCuotas_(ss) : null);
  const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  const today = startOfDay_(new Date());

  // Cliente (reusa/crea) + invalida cachés para que las fórmulas Nombre/DNI lo resuelvan.
  const c = opts.cliente;
  const clienteId = getOrCreateClienteV2_(c.name, c.email, normDni_(c.dni), c.phone, {});
  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
  if (typeof invalidateClientesRows_ === 'function') invalidateClientesRows_();
  if (typeof invalidateClientesHdrCache_ === 'function') invalidateClientesHdrCache_();

  const rate = (opts.rate != null) ? opts.rate : (termRateDays_(opts.plazoDias) || 0.5);
  const total = round2_(opts.capital * (1 + rate));
  const loanDate = addDays_(today, opts.fechaOffsetDays || 0);

  const H = headerIndex_(bs);
  const idCol = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const cliCol = colByAny_(H, ['ID Cliente']) || PB.CLIENT_ID;
  const capCol = colByAny_(H, ['Capital']) || PB.PRINCIPAL;
  const plzCol = colByAny_(H, ['Plazo (días)', 'Plazo (dias)', 'Plazo']) || PB.TERM;
  const fecCol = colByAny_(H, ['Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo']) || PB.LOAN_DATE;
  const firmaCol = colByAny_(H, ['Estado de Firma']) || PB.SIGN_STATUS;
  const nameCol = colByAny_(H, ['Nombre']);
  const dniCol = colByAny_(H, ['DNI']);
  const congCol = colByAny_(H, ['Congelado el']);
  const planCol = colByAny_(H, ['Plan desde']);
  const adjCol = colByAny_(H, ['Mora (ajuste)']);

  const loanId = nextLoanId_();
  const tRow = insertBorrowerRow_(bs, idCol);
  bs.getRange(tRow, idCol).setValue(loanId).setNote('🧪 Préstamo de PRUEBA (' + (opts.estado || '') + '). Borralo con «Eliminar escenarios de pago».');
  bs.getRange(tRow, cliCol).setValue(clienteId);
  bs.getRange(tRow, capCol).setValue(opts.capital);
  bs.getRange(tRow, plzCol).setValue(opts.plazoDias);
  bs.getRange(tRow, fecCol).setValue(loanDate).setNumberFormat('yyyy-mm-dd');
  if (firmaCol) bs.getRange(tRow, firmaCol).setValue(SIGN.SIGNED);
  if (opts.congeladoOffset != null && congCol) bs.getRange(tRow, congCol).setValue(addDays_(today, opts.congeladoOffset)).setNumberFormat('yyyy-mm-dd');
  if (opts.planOffset != null && planCol) bs.getRange(tRow, planCol).setValue(addDays_(today, opts.planOffset)).setNumberFormat('yyyy-mm-dd');
  if (opts.moraAjuste != null && adjCol) bs.getRange(tRow, adjCol).setValue(opts.moraAjuste);

  writeBorrowerRowFormulas_(bs, tRow);
  if (congCol && planCol) bs.getRange(tRow, colByAny_(H, ['Estado']) || PB.STATE).setFormula(borrowerStateFormula_(tRow));

  if ((nameCol || dniCol) && cl) {
    const CH = headerIndex_(cl), cliL = a1col_(cliCol), CSN = CFG.SHEETS.CLIENTS;
    const cIdL = a1col_(colByAny_(CH, ['ID Cliente']) || 1);
    if (nameCol) { const cNameL = a1col_(colByAny_(CH, ['Nombre']) || 2);
      bs.getRange(tRow, nameCol).setFormula(`=IFERROR(INDEX('${CSN}'!$${cNameL}:$${cNameL},MATCH($${cliL}${tRow},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`); }
    if (dniCol) { const cDniL = a1col_(colByAny_(CH, ['DNI']) || 4);
      bs.getRange(tRow, dniCol).setFormula(`=IFERROR(INDEX('${CSN}'!$${cDniL}:$${cDniL},MATCH($${cliL}${tRow},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`); }
  }

  // Cuotas: fechas explícitas (plan) o por offsets sobre la Fecha Préstamo.
  const cuo = opts.cuotas || { n: 1 }, n = cuo.n || 1;
  if (cuo.dates && cuo.dates.length) cuotasForLoan_(cs, loanId, loanDate, total, n, null, { dates: cuo.dates });
  else cuotasForLoan_(cs, loanId, loanDate, total, n, cuo.dueDays || [opts.plazoDias]);

  // Pagos.
  let seq = nextPaymentSeq_(ps);
  (opts.pagos || []).forEach(function (p) { addTestPayment_(ps, seq++, loanId, c.name, c.dni, addDays_(today, p.offsetDays), p.monto); });

  return { loanId: loanId, clienteId: clienteId, total: total };
}

/** Agrega una fila de pago de prueba a «Pagos» (con su «Saldo Posterior» vivo). */
function addTestPayment_(ps, seq, loanId, name, dni, date, monto) {
  const row = ps.getLastRow() + 1;
  ps.getRange(row, PP.PAY_ID).setValue('P-' + String(seq).padStart(4, '0'));
  ps.getRange(row, PP.LOAN_ID).setValue(loanId);
  ps.getRange(row, PP.DATE).setValue(date).setNumberFormat('yyyy-mm-dd');
  ps.getRange(row, PP.AMOUNT).setValue(monto).setNumberFormat(CFG.CURRENCY_FMT);
  ps.getRange(row, PP.NAME).setValue(name);
  ps.getRange(row, PP.DNI).setValue(normDni_(dni));
  ps.getRange(row, PP.BALANCE).setFormula(paymentBalanceFormula_(row));
  SpreadsheetApp.flush();
}

/** Menú: confirma y elimina los clientes/préstamos/pagos/cuotas de los escenarios de prueba. */
function removePaymentScenariosConfirm() {
  return guard_('removePaymentScenariosConfirm', function () {
    const ui = SpreadsheetApp.getUi();
    const r = ui.alert('Eliminar escenarios de pago de prueba',
      'Se eliminarán los clientes de prueba (prefijo «' + SCENARIO_TEST_PREFIX + '») y TODOS sus préstamos, pagos y cuotas. ¿Continuar?',
      ui.ButtonSet.OK_CANCEL);
    if (r !== ui.Button.OK) return;
    const n = removePaymentScenarios_();
    ui.alert('Listo', n.loans + ' préstamo(s) y ' + n.clientes + ' cliente(s) de prueba eliminados.', ui.ButtonSet.OK);
  });
}

/** Elimina clientes de prueba (por prefijo) + sus préstamos/pagos/cuotas. Devuelve {clientes, loans}. */
function removePaymentScenarios_() {
  const ss = getSS_();
  const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS);
  if (!cl || !bs) return { clientes: 0, loans: 0 };

  // 1) Clientes de prueba (por prefijo en Nombre).
  const CH = headerIndex_(cl);
  const cIdC = colByAny_(CH, ['ID Cliente']) || 1, cNameC = colByAny_(CH, ['Nombre']) || 2;
  const clientRows = [], clientIds = {};
  if (cl.getLastRow() >= 2) {
    const cv = cl.getRange(2, 1, cl.getLastRow() - 1, Math.max(cIdC, cNameC)).getValues();
    for (let i = 0; i < cv.length; i++) if (String(cv[i][cNameC - 1] || '').indexOf(SCENARIO_TEST_PREFIX) === 0) {
      clientRows.push(i + 2); clientIds[String(cv[i][cIdC - 1]).trim()] = true;
    }
  }
  // 2) Préstamos de esos clientes.
  const BH = headerIndex_(bs);
  const bIdC = colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || PB.LOAN_ID;
  const bCliC = colByAny_(BH, ['ID Cliente']) || PB.CLIENT_ID;
  const loanRows = [], loanIds = {};
  if (bs.getLastRow() >= 2) {
    const bv = bs.getRange(2, 1, bs.getLastRow() - 1, Math.max(bIdC, bCliC)).getValues();
    for (let i = 0; i < bv.length; i++) if (clientIds[String(bv[i][bCliC - 1]).trim()]) {
      loanRows.push(i + 2); loanIds[String(bv[i][bIdC - 1]).trim()] = true;
    }
  }
  // 3) Cuotas de esos préstamos.
  if (cs) Object.keys(loanIds).forEach(function (id) { try { deleteCuotasRowsForLoan_(cs, id); } catch (e) {} });
  // 4) Pagos de esos préstamos (de abajo hacia arriba).
  if (ps && ps.getLastRow() >= 2) {
    const pIdC = colByAny_(headerIndex_(ps), ['ID Préstamo', 'ID Prestamo']) || PP.LOAN_ID;
    const pv = ps.getRange(2, pIdC, ps.getLastRow() - 1, 1).getValues();
    for (let i = pv.length - 1; i >= 0; i--) if (loanIds[String(pv[i][0]).trim()]) ps.deleteRow(i + 2);
  }
  // 5) Préstamos y 6) Clientes (de abajo hacia arriba).
  loanRows.sort(function (a, b) { return b - a; }).forEach(function (r) { bs.deleteRow(r); });
  clientRows.sort(function (a, b) { return b - a; }).forEach(function (r) { cl.deleteRow(r); });

  if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
  if (typeof invalidateClientesRows_ === 'function') invalidateClientesRows_();
  if (typeof invalidateClientesHdrCache_ === 'function') invalidateClientesHdrCache_();
  SpreadsheetApp.flush();
  try { rebuildLateSheet_(ss); } catch (e) {}
  try { fillPaymentWaLinks_(ss); } catch (e) {}
  return { clientes: clientRows.length, loans: loanRows.length };
}