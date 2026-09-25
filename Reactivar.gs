/************************************************************
 * REACTIVAR v1  —  Complemento de LoanManagerV2.gs + Validaciones.gs
 * ----------------------------------------------------------
 * La migración dejó los datos correctos y normalizados, pero las
 * hojas quedaron con VALORES ESTÁTICOS. Este módulo vuelve a
 * instalar la FUNCIONALIDAD VIVA de cada hoja (fórmulas,
 * validaciones, formato condicional, protecciones e indicadores)
 * SIN BORRAR datos: resuelve las columnas por el NOMBRE del
 * encabezado, detecta el bloque de datos y se detiene en filas en
 * blanco o de TOTAL.
 *
 * Idempotente: se puede ejecutar varias veces sin duplicar nada.
 *
 * ESQUEMA MIGRADO (días):
 *   Prestatarios: ID Préstamo | ID Cliente | Capital | Plazo (días) |
 *     Tasa | Fecha Préstamo | Vencimiento | Interés | Total a Pagar |
 *     Total Pagado | Saldo Pendiente | Estado | Estado de Firma
 *   Pagos: ID Pago | ID Préstamo | Fecha de Pago | Monto Pagado |
 *     Saldo Posterior | Recibo Enviado
 *   Clientes: …| Correo válido | Duplicado | Nota
 *   Resumen: ID Cliente | Prestatario | Préstamos | Capital Total |
 *     Interés Total | Total Pagado | Saldo Pendiente | Estado
 *   Nuevos Prestatarios: solicitudes del formulario (casillas
 *     Verificado?/Rechazar?/Verificar BCRA?, Parámetro BCRA, CUIL, …).
 *
 * USO
 *   1) Ejecute  ->  reactivarFuncionalidad   (autorice Sheets).
 *      Reinstala fórmulas/validaciones/formato en TODAS las hojas
 *      (incluida "Nuevos Prestatarios": casillas + Plazo + formato).
 *   2) Ejecute  ->  activarDisparadores   (autorice Drive/Gmail/UrlFetch).
 *      Instala el onEdit para que las casillas ejecuten acciones.
 *   3) APROBAR: tildá "Verificado?" en una solicitud → corre la
 *      verificación BCRA (V-11), valida (V-02/V-08/V-09/V-10/V-13),
 *      crea/reutiliza el cliente y MUEVE el préstamo a "Prestatarios"
 *      (Estado de Firma = PENDIENTE). "Rechazar?" → archiva en Rechazados.
 *      Alternativa por lote (sin disparador): procesarNuevosPrestatarios.
 *      Reemplaza al flujo month-based de LoanManagerV2.gs.
 ************************************************************/

/* ============================ HELPERS ============================ */

// Número de columna -> letra(s) A1 (1->A, 27->AA).
function a1col_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// Clave normalizada de encabezado/etiqueta: minúsculas, sin acentos, espacios colapsados.
function hkey_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

// Mapa {claveNormalizada -> índiceColumna} del encabezado (fila 1).
// Memoizado por ejecución y por nombre de hoja. Se auto-invalida si cambia la
// cantidad de columnas (p. ej. migraciones que reestructuran la hoja), evitando
// leer un mapa obsoleto. El mapa es de solo lectura para quien lo consume.
const _headerIndexCache = {};
function invalidateHeaderIndexCache_() { for (const k in _headerIndexCache) delete _headerIndexCache[k]; }
function headerIndex_(sh) {
  const name = sh.getName(), last = sh.getLastColumn();
  const cached = _headerIndexCache[name];
  if (cached && cached.last === last) return cached.map;
  const map = {};
  if (last >= 1) sh.getRange(1, 1, 1, last).getValues()[0].forEach((h, i) => { const k = hkey_(h); if (k) map[k] = i + 1; });
  _headerIndexCache[name] = { last: last, map: map };
  return map;
}
// Primer índice de columna que coincide con alguno de los nombres candidatos.
function colByAny_(H, candidates) { for (const c of candidates) { const k = hkey_(c); if (H[k]) return H[k]; } return 0; }

// Bloque de datos contiguo por una columna clave: desde la fila 2 hasta la
// primera vacía (evita la fila TOTAL, que tiene la clave en blanco).
function dataBlock_(sh, keyCol) {
  const maxR = sh.getLastRow(); if (maxR < 2 || !keyCol) return { first: 2, last: 1, count: 0 };
  const vals = sh.getRange(2, keyCol, maxR - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < vals.length; i++) { if (String(vals[i][0]).trim() === '') break; count++; }
  return { first: 2, last: 1 + count, count: count };
}

// Escribe fórmulas en una columna del bloque (fn(row) -> fórmula A1).
function setColFormulas_(sh, col, first, count, fn) {
  if (!col || count <= 0) return;
  const arr = []; for (let i = 0; i < count; i++) arr.push([fn(first + i)]);
  sh.getRange(first, col, count, 1).setFormulas(arr);
}

// Aplica formato de número a una columna del bloque.
function fmtCol_(sh, col, first, count, fmt) { if (col && count > 0) sh.getRange(first, col, count, 1).setNumberFormat(fmt); }

// Asegura que exista una columna con ese encabezado; si falta, la agrega a la
// derecha (sin borrar datos). Devuelve el índice 1-based. Resuelve por NOMBRE, así
// funciona en cualquier layout (migrado o nuevo).
function ensureHeaderColumn_(sh, header, bg) {
  const found = colByAny_(headerIndex_(sh), [header]);
  if (found) return found;
  const col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(header).setFontWeight('bold')
    .setBackground(bg || '#1c4587').setFontColor('#fff').setWrap(true);
  invalidateHeaderIndexCache_();
  return col;
}

/** Inserta (una sola vez) las columnas "Recargo por Mora (acum.)", "Mora (ajuste)" y
 *  "Saldo con Mora" en "Prestatarios", justo DESPUÉS de "Préstamos (de N)" — mismo orden que
 *  el layout nuevo (PB.MORA_ACUM / PB.MORA_ADJ / PB.BALANCE_MORA). Desplaza a la derecha lo
 *  que hubiera (banner). Idempotente (crea sólo las que falten). */
function ensureMoraColumns_(bs) {
  const setHead = (c, t) => bs.getRange(1, c).setValue(t).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
  if (!colByAny_(headerIndex_(bs), ['Recargo por Mora (acum.)'])) {
    const cntC = colByAny_(headerIndex_(bs), ['Préstamos (de ' + MAX_LOANS_PER_CLIENT + ')', 'Préstamos']) || PB.LOANCOUNT;
    bs.insertColumnsAfter(cntC, 3);
    setHead(cntC + 1, 'Recargo por Mora (acum.)');
    setHead(cntC + 2, 'Mora (ajuste)');
    setHead(cntC + 3, 'Saldo con Mora');
    invalidateHeaderIndexCache_();
    return;
  }
  // "Recargo" ya existe: agrega "Mora (ajuste)" si falta (p. ej. versión previa sin ajuste),
  // insertándola justo después de "Recargo por Mora (acum.)".
  if (!colByAny_(headerIndex_(bs), ['Mora (ajuste)'])) {
    const recC = colByAny_(headerIndex_(bs), ['Recargo por Mora (acum.)']);
    bs.insertColumnsAfter(recC, 1);
    setHead(recC + 1, 'Mora (ajuste)');
    invalidateHeaderIndexCache_();
  }
}

/** Escribe las fórmulas de "Recargo por Mora (acum.)" y "Saldo con Mora" en todo el bloque de
 *  datos de "Prestatarios" (respetando "Mora (ajuste)"). "Mora (ajuste)" queda como entrada.
 *  Usado por 🆕 Aplicar novedades para que las columnas no queden vacías sin un reactivar completo. */
function writeMoraFormulasBlock_(sh) {
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']); if (!idC) return;
  const b = dataBlock_(sh, idC); if (!b.count) return;
  const A = a1col_(idC);
  const K = a1col_(colByAny_(H, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || 11);
  const morC = colByAny_(H, ['Recargo por Mora (acum.)']);
  const adjC = colByAny_(H, ['Mora (ajuste)']);
  const bmC = colByAny_(H, ['Saldo con Mora']);
  const MOR = morC ? a1col_(morC) : '';
  const ADJ = adjC ? a1col_(adjC) : '';
  if (morC) setColFormulas_(sh, morC, b.first, b.count, r => {
    const auto = (typeof moraLookupFormula_ === 'function') ? moraLookupFormula_('$' + A + r) : '0';
    return ADJ ? `=IF($${A}${r}="","",IF(ISNUMBER($${ADJ}${r}),MAX(0,$${ADJ}${r}),${auto}))` : `=IF($${A}${r}="","",${auto})`;
  });
  if (bmC) setColFormulas_(sh, bmC, b.first, b.count, r => MOR ? `=IF($${A}${r}="","",N($${K}${r})+N($${MOR}${r}))` : `=IF($${A}${r}="","",N($${K}${r}))`);
  if (morC) fmtCol_(sh, morC, b.first, b.count, CFG.CURRENCY_FMT);
  if (adjC) { fmtCol_(sh, adjC, b.first, b.count, CFG.CURRENCY_FMT); sh.getRange(1, adjC).setNote('AJUSTE MANUAL de la mora. Vacío = automático. 0 = condonar. Un importe = fijar ese recargo.'); }
  if (bmC) fmtCol_(sh, bmC, b.first, b.count, CFG.CURRENCY_FMT);
}

/** Agrega/actualiza (idempotente) una regla de formato condicional "≥ n" sobre una
 *  columna, sin duplicarla ni pisar las reglas de otras columnas. */
function setColConditionalGteRule_(sh, col, first, count, n, bg) {
  const rng = sh.getRange(first, col, count, 1);
  const kept = sh.getConditionalFormatRules().filter(rule =>
    !(rule.getRanges() || []).some(rr => rr.getColumn() === col));
  kept.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(n).setBackground(bg).setRanges([rng]).build());
  sh.setConditionalFormatRules(kept);
}

/**
 * Instala (SIN borrar datos, idempotente) las columnas/casillas nuevas de esta versión
 * en cualquier layout — resuelve todo por NOMBRE de encabezado, así funciona igual en el
 * esquema migrado o en el nuevo. Lo llaman "Aplicar novedades" y "reactivarFuncionalidad".
 *   • Prestatarios: casilla "Enviar recibo desembolso", columna "Préstamos (de N)"
 *     (conteo por prestatario, rojo al tope) y banner "Capital disponible" arriba.
 *   • Pagos: casilla "Enviar recibo".
 *   • Nuevos Prestatarios: casilla "Anular límites" (override de topes).
 */
function installNewFeatureColumns_(ss) {
  ss = ss || getSS_();
  const out = [];

  // --- Prestatarios ---
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (bs) {
    const H = headerIndex_(bs);
    const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']);
    const cliC = colByAny_(H, ['ID Cliente']);
    const capC = colByAny_(H, ['Capital']);
    const pagC = colByAny_(H, ['Total Pagado']);
    const b = dataBlock_(bs, idC);

    const disbCol = ensureHeaderColumn_(bs, 'Enviar recibo desembolso');
    if (b.count) bs.getRange(b.first, disbCol, b.count, 1).insertCheckboxes();
    bs.getRange(1, disbCol).setNote('Tildá para enviar el recibo de DESEMBOLSO al prestatario (sólo si Estado de Firma = ' + SIGN.SIGNED + '). Se destilda solo.');

    const cntCol = ensureHeaderColumn_(bs, 'Préstamos (de ' + MAX_LOANS_PER_CLIENT + ')');
    if (cliC && b.count) {
      const cliL = a1col_(cliC);
      const estC = colByAny_(H, ['Estado']);
      const estL = estC ? a1col_(estC) : '';
      // Préstamos VIGENTES (excluye PAGADO/SALDADO): coincide con clientLoanCount_ y el tope V-25.
      setColFormulas_(bs, cntCol, b.first, b.count, r =>
        estL
          ? `=IF($${cliL}${r}="","",COUNTIFS($${cliL}$${b.first}:$${cliL}$${b.last},$${cliL}${r},$${estL}$${b.first}:$${estL}$${b.last},"<>${ST.PAID}",$${estL}$${b.first}:$${estL}$${b.last},"<>${ST.CLEARED}"))`
          : `=IF($${cliL}${r}="","",COUNTIF($${cliL}$${b.first}:$${cliL}$${b.last},$${cliL}${r}))`);
      fmtCol_(bs, cntCol, b.first, b.count, '0" / ' + MAX_LOANS_PER_CLIENT + '"');
      bs.getRange(b.first, cntCol, b.count, 1).setHorizontalAlignment('center');
      setColConditionalGteRule_(bs, cntCol, b.first, b.count, MAX_LOANS_PER_CLIENT, '#ea9999');
      protectFormulaCols_(bs, [cntCol], b.first, b.count);
    }
    bs.getRange(1, cntCol).setNote('Préstamos vigentes del prestatario (máximo ' + MAX_LOANS_PER_CLIENT + ').');
    ensureMoraColumns_(bs); // columnas separadas Recargo / Mora (ajuste) / Saldo con Mora (tras "Préstamos (de N)")
    writeMoraFormulasBlock_(bs); // rellena las fórmulas para que no queden vacías con 🆕 solo

    if (capC && pagC) {
      try {
        const capL = a1col_(capC), pagL = a1col_(pagC), bannerCol = cntCol + 4; // deja lugar a las 3 columnas de mora
        bs.getRange(1, bannerCol, 1, 3).breakApart();
        const banner = bs.getRange(1, bannerCol, 1, 3).merge();
        banner.setFormula('="Capital disponible (máx. asignable): " & TEXT((' + fondoFormula_() +
          ')-SUM($' + capL + '$' + b.first + ':$' + capL + ')+SUM($' + pagL + '$' + b.first + ':$' + pagL + '),"$#,##0.00")');
        banner.setBackground('#38761d').setFontColor('#fff').setFontWeight('bold')
          .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
        banner.setNote('Máximo asignable a un prestatario (sus préstamos sumados no pueden superarlo). Cada prestatario admite hasta ' + MAX_LOANS_PER_CLIENT + ' préstamos.');
        bs.setColumnWidth(bannerCol, 240);
      } catch (e) { logError_('installNewFeatureColumns_:banner', e); }
    }
    out.push('Prestatarios');
  }

  // --- Pagos ---
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (pg) {
    const idC = colByAny_(headerIndex_(pg), ['ID Préstamo', 'ID Prestamo']);
    const b = dataBlock_(pg, idC);
    const sendCol = ensureHeaderColumn_(pg, 'Enviar recibo', '#38761d');
    if (b.count) pg.getRange(b.first, sendCol, b.count, 1).insertCheckboxes();
    pg.getRange(1, sendCol).setNote('Tildá para enviar/reenviar el recibo de pago de esa fila. Se destilda solo y deja el enlace en "Recibo Enviado".');
    out.push('Pagos');
  }

  // --- Nuevos Prestatarios ---
  const nb = ss.getSheetByName(CFG.SHEETS.NEW);
  if (nb) {
    const ovrCol = ensureHeaderColumn_(nb, 'Anular límites', '#b45f06');
    nb.getRange(2, ovrCol, CFG.MAX_ROWS, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build());
    nb.getRange(1, ovrCol).setNote('Tildá "Anular límites" ANTES de "Verificado?" para aprobar aunque el prestatario ya tenga ' + MAX_LOANS_PER_CLIENT + ' préstamos, se exceda el capital disponible, o el BCRA marque RECHAZAR.');
    out.push('Nuevos Prestatarios');
  }

  return out.length ? out.join(' · ') : 'sin hojas';
}

// Quita las protecciones creadas por este módulo (idempotencia).
function clearReactivarProtections_(sh) {
  try {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
      if (String(p.getDescription() || '').indexOf('REACTIVAR') === 0) { try { p.remove(); } catch (e) { } }
    });
  } catch (e) { }
}
// Protege columnas de fórmula (aviso, no bloqueo).
function protectFormulaCols_(sh, cols, first, count) {
  if (!count) return;
  cols.filter(Boolean).forEach(col => {
    try {
      sh.getRange(first, col, count, 1).protect().setWarningOnly(true)
        .setDescription('REACTIVAR: columna calculada — no editar');
    } catch (e) { }
  });
}
// Busca una etiqueta en toda la hoja; devuelve {row, valCol, valA1} donde valCol
// es la columna inmediata a la derecha. null si no la encuentra.
function findLabelCell_(sh, label) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 1 || lc < 1) return null;
  const vals = sh.getRange(1, 1, lr, lc).getValues(), key = hkey_(label);
  for (let r = 0; r < vals.length; r++) for (let c = 0; c < vals[r].length; c++) {
    const cell = hkey_(vals[r][c]);
    if (cell === key || (key.length > 4 && cell.indexOf(key) === 0)) {
      const valCol = c + 2; // columna a la derecha de la etiqueta
      return { row: r + 1, valCol: valCol, valA1: a1col_(valCol) + (r + 1) };
    }
  }
  return null;
}

/* ============================ ENTRADA ============================ */

function reactivarFuncionalidad() {
  return guard_('reactivarFuncionalidad', function () {
    const ss = getSS_(), report = [];
    const run = (label, fn) => { try { report.push(label + ': ' + fn()); } catch (e) { logError_('reactivar:' + label, e); report.push(label + ': ⚠ ' + e.message); } };
    run('Configuración', () => reactivateSettings_(ss));
    run('Prestatarios', () => reactivateBorrowers_(ss));
    run('Nuevos Prestatarios', () => reactivateNew_(ss));
    run('Pagos', () => reactivatePayments_(ss));
    run('Clientes', () => reactivateClientes_(ss));
    run('Resumen', () => reactivateResumen_(ss));
    run('Panel', () => reactivatePanel_(ss));
    run('Estadísticas', () => reactivateStats_(ss));
    run('Novedades (recibos + topes)', () => installNewFeatureColumns_(ss));
    // Validaciones de datos (listas/casillas) del módulo Validaciones.
    try { if (typeof applySheetValidations_ === 'function') applySheetValidations_(ss); } catch (e) { logError_('reactivar:validaciones', e); }
    SpreadsheetApp.flush();
    const msg = report.join('\n');
    try { SpreadsheetApp.getActive().toast('Funcionalidad reactivada. Ver detalle en el registro.', 'Listo', 6); } catch (e) { }
    return msg;
  });
}

/* ======================= CONFIGURACIÓN ======================= */
/**
 * Crea la hoja "Configuración" (ajustes clave→valor) si NO existe — la migración
 * no la trajo, y sin ella fondoTotal_() = 0 y V-09 bloquea toda aprobación.
 * Precarga el "Fondo total para prestar" con el valor del Panel migrado
 * (5.330.000) y el correo del prestamista. NO toca una hoja ya existente.
 */
function reactivateSettings_(ss) {
  const name = CFG.SHEETS.SETTINGS;
  if (ss.getSheetByName(name)) return 'ya existe (sin cambios)';
  if (typeof setupSettings_ !== 'function') return '⚠ setupSettings_ no disponible';
  setupSettings_(ss);
  const sh = ss.getSheetByName(name); if (!sh) return 'no se pudo crear';
  setSettingValue_(sh, 'Fondo total para prestar', 5330000); // Panel: 5.330.000
  try { const me = Session.getActiveUser().getEmail(); if (me) setSettingValue_(sh, 'Correo del Prestamista', me); } catch (e) { }
  return 'creada (Fondo=5.330.000 — verificá los demás ajustes)';
}
/** Asigna el valor (col B) de un ajuste buscándolo por su nombre (col A). */
function setSettingValue_(sh, key, val) {
  const last = sh.getLastRow(); if (last < 1) return false;
  const rows = sh.getRange(1, 1, last, 1).getValues();
  for (let i = 0; i < rows.length; i++) if (hkey_(rows[i][0]) === hkey_(key)) { sh.getRange(i + 1, 2).setValue(val); return true; }
  return false;
}

/* ======================= PRESTATARIOS ======================= */
function reactivateBorrowers_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.BORROWERS); if (!sh) return 'hoja ausente';
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  ensureMoraColumns_(sh); // columnas separadas de mora (Recargo por Mora / Saldo con Mora)
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']);
  const cliC = colByAny_(H, ['ID Cliente']);
  const capC = colByAny_(H, ['Capital']);
  const plzC = colByAny_(H, ['Plazo (días)', 'Plazo (dias)', 'Plazo']);
  const tasC = colByAny_(H, ['Tasa']);
  const fecC = colByAny_(H, ['Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo']);
  const venC = colByAny_(H, ['Vencimiento']);
  const intC = colByAny_(H, ['Interés', 'Interes']);
  const totC = colByAny_(H, ['Total a Pagar', 'Total a pagar']);
  const pagC = colByAny_(H, ['Total Pagado']);
  const salC = colByAny_(H, ['Saldo Pendiente']);
  const morC = colByAny_(H, ['Recargo por Mora (acum.)']);
  const adjC = colByAny_(H, ['Mora (ajuste)']);
  const bmC = colByAny_(H, ['Saldo con Mora']);
  const estC = colByAny_(H, ['Estado']);
  const firmaC = colByAny_(H, ['Estado de Firma']);
  if (!idC) return 'falta ID Préstamo';
  const b = dataBlock_(sh, idC); if (!b.count) return 'sin datos';

  // Letras A1 de columnas usadas en fórmulas.
  const A = a1col_(idC), C = a1col_(capC), D = a1col_(plzC), E = a1col_(tasC),
    F = a1col_(fecC), G = a1col_(venC), Hh = a1col_(intC), I = a1col_(totC), J = a1col_(pagC), K = a1col_(salC);
  const MOR = morC ? a1col_(morC) : '';
  const ADJ = adjC ? a1col_(adjC) : '';
  const BM = bmC ? a1col_(bmC) : K; // Estado cae al saldo base si aún no existe "Saldo con Mora"
  // Columnas de "Pagos" (para SUMIF de Total Pagado, V-17).
  let payLoanL = 'B', payAmtL = 'D';
  if (pg) { const PH = headerIndex_(pg); payLoanL = a1col_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || 2); payAmtL = a1col_(colByAny_(PH, ['Monto Pagado']) || 4); }
  const PGN = CFG.SHEETS.PAYMENTS;

  // Fórmulas por columna calculada.
  if (tasC) setColFormulas_(sh, tasC, b.first, b.count, r => `=IF($${D}${r}="","",IF($${D}${r}=15,0.25,IF($${D}${r}=30,0.5,IF($${D}${r}=60,1,IF($${D}${r}=90,1,"")))))`);
  if (venC) setColFormulas_(sh, venC, b.first, b.count, r => `=IF($${F}${r}="","",$${F}${r}+$${D}${r})`);
  if (intC) setColFormulas_(sh, intC, b.first, b.count, r => `=IF($${C}${r}="","",$${C}${r}*$${E}${r})`);
  if (totC) setColFormulas_(sh, totC, b.first, b.count, r => `=IF($${C}${r}="","",$${C}${r}+$${Hh}${r})`);
  if (pagC) setColFormulas_(sh, pagC, b.first, b.count, r => `=IF($${A}${r}="","",SUMIF('${PGN}'!$${payLoanL}:$${payLoanL},$${A}${r},'${PGN}'!$${payAmtL}:$${payAmtL}))`);
  if (salC) setColFormulas_(sh, salC, b.first, b.count, r => `=IF($${A}${r}="","",MAX(0,$${I}${r}-$${J}${r}))`);
  // Recargo por Mora (acum.): usa el AJUSTE manual si es número (0 = condonar); si no, el
  // recargo automático derivado de "Pagos Atrasados".
  if (morC) {
    setColFormulas_(sh, morC, b.first, b.count, r => {
      const auto = (typeof moraLookupFormula_ === 'function') ? moraLookupFormula_('$' + A + r) : '0';
      return ADJ
        ? `=IF($${A}${r}="","",IF(ISNUMBER($${ADJ}${r}),MAX(0,$${ADJ}${r}),${auto}))`
        : `=IF($${A}${r}="","",${auto})`;
    });
  }
  // Saldo con Mora = Saldo Pendiente + Recargo por Mora = monto a pagar HOY.
  if (bmC) setColFormulas_(sh, bmC, b.first, b.count, r => MOR ? `=IF($${A}${r}="","",N($${K}${r})+N($${MOR}${r}))` : `=IF($${A}${r}="","",N($${K}${r}))`);
  // V-16: Estado se DERIVA del monto a pagar (Saldo con Mora) y la fecha. PAGADO sólo si nada
  // se debe, incluida la mora. V-15: sin mora si saldo 0.
  if (estC) setColFormulas_(sh, estC, b.first, b.count, r => `=IF($${A}${r}="","",IF($${BM}${r}<=0,"${ST.PAID}",IF(OR(${cuotaVencidaExpr_('$' + A + r)},TODAY()>$${G}${r}),"${ST.OVERDUE}","${ST.ACTIVE}")))`);

  // Nombre/DNI (si "Prestatarios" los muestra como columnas): se resuelven del cliente por
  // "ID Cliente" con INDEX/MATCH. Repuebla toda la cartera, sanando filas que quedaron en blanco.
  const nameC = colByAny_(H, ['Nombre']), dniC = colByAny_(H, ['DNI']);
  const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  if ((nameC || dniC) && cliC && cl) {
    const CH = headerIndex_(cl), CSN = CFG.SHEETS.CLIENTS;
    const Bc = a1col_(cliC), cIdL = a1col_(colByAny_(CH, ['ID Cliente']) || 1);
    if (nameC) { const cNameL = a1col_(colByAny_(CH, ['Nombre']) || 2);
      setColFormulas_(sh, nameC, b.first, b.count, r => `=IF($${A}${r}="","",IFERROR(INDEX('${CSN}'!$${cNameL}:$${cNameL},MATCH($${Bc}${r},'${CSN}'!$${cIdL}:$${cIdL},0)),""))`); }
    if (dniC) { const cDniL = a1col_(colByAny_(CH, ['DNI']) || 4);
      setColFormulas_(sh, dniC, b.first, b.count, r => `=IF($${A}${r}="","",IFERROR(INDEX('${CSN}'!$${cDniL}:$${cDniL},MATCH($${Bc}${r},'${CSN}'!$${cIdL}:$${cIdL},0)),""))`); }
  }

  // Formatos.
  fmtCol_(sh, capC, b.first, b.count, CFG.CURRENCY_FMT);
  fmtCol_(sh, intC, b.first, b.count, CFG.CURRENCY_FMT);
  fmtCol_(sh, totC, b.first, b.count, CFG.CURRENCY_FMT);
  fmtCol_(sh, pagC, b.first, b.count, CFG.CURRENCY_FMT);
  fmtCol_(sh, salC, b.first, b.count, CFG.CURRENCY_FMT);
  if (morC) fmtCol_(sh, morC, b.first, b.count, CFG.CURRENCY_FMT);
  if (adjC) { fmtCol_(sh, adjC, b.first, b.count, CFG.CURRENCY_FMT); sh.getRange(1, adjC).setNote('AJUSTE MANUAL de la mora. Vacío = automático. 0 = condonar. Un importe = fijar ese recargo.'); }
  if (bmC) fmtCol_(sh, bmC, b.first, b.count, CFG.CURRENCY_FMT);
  fmtCol_(sh, tasC, b.first, b.count, '0%');
  fmtCol_(sh, fecC, b.first, b.count, 'yyyy-mm-dd');
  fmtCol_(sh, venC, b.first, b.count, 'yyyy-mm-dd');

  // Validaciones de datos: Plazo (V-13) y Estado de Firma.
  if (plzC) sh.getRange(2, plzC, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['15', '30', '60', '90'], true).setAllowInvalid(false)
      .setHelpText('Elegí un plazo válido: 15, 30 o 60 días.').build());
  if (firmaC) sh.getRange(2, firmaC, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([SIGN.PENDING, SIGN.SIGNED], true).setAllowInvalid(true).build());

  // Formato condicional por Estado, POR FILA (franja en los campos relevantes A→Estado),
  // más el color propio de la firma y el tope de préstamos. Cobertura completa de filas
  // y reglas unificadas con el setup para que ninguna ruta pise a la otra.
  sh.setConditionalFormatRules(borrowerFormatRules_(sh));

  // Protección (aviso) de columnas de fórmula.
  clearReactivarProtections_(sh);
  protectFormulaCols_(sh, [tasC, venC, intC, totC, pagC, salC, estC], b.first, b.count);

  return b.count + ' préstamos · fórmulas Tasa/Venc/Interés/Total/Pagado/Saldo/Estado + validaciones + formato';
}

/**
 * Escribe las fórmulas de columnas calculadas (Tasa, Vencimiento, Interés, Total a
 * Pagar, Total Pagado, Saldo Pendiente, Estado) en UNA fila de "Prestatarios".
 * Resuelve las columnas por NOMBRE de encabezado (mismas fórmulas que reactivateBorrowers_).
 * Se usa al APROBAR para que la fila nueva quede completa al instante, sin esperar un refresh.
 */
function writeBorrowerRowFormulas_(bs, r) {
  const ss = getSS_(), pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  const H = headerIndex_(bs);
  const A = a1col_(colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || 1);
  const C = a1col_(colByAny_(H, ['Capital']) || 3);
  const D = a1col_(colByAny_(H, ['Plazo (días)', 'Plazo (dias)', 'Plazo (meses)', 'Plazo']) || 4);
  const E = a1col_(colByAny_(H, ['Tasa']) || 5);
  const F = a1col_(colByAny_(H, ['Fecha Préstamo', 'Fecha de Préstamo', 'Fecha del Préstamo', 'Fecha Prestamo']) || 6);
  const G = a1col_(colByAny_(H, ['Vencimiento', 'Fecha de Vencimiento']) || 7);
  const Hh = a1col_(colByAny_(H, ['Interés', 'Interes']) || 8);
  const I = a1col_(colByAny_(H, ['Total a Pagar', 'Total a pagar']) || 9);
  const J = a1col_(colByAny_(H, ['Total Pagado']) || 10);
  const K = a1col_(colByAny_(H, ['Saldo Pendiente', 'Saldo Pendiente (hoy)']) || 11);
  const morC = colByAny_(H, ['Recargo por Mora (acum.)']);
  const adjC = colByAny_(H, ['Mora (ajuste)']);
  const bmC = colByAny_(H, ['Saldo con Mora']);
  const MOR = morC ? a1col_(morC) : '';
  const ADJ = adjC ? a1col_(adjC) : '';
  const BM = bmC ? a1col_(bmC) : K; // si aún no existe la columna, el Estado cae al saldo base
  let payLoanL = 'B', payAmtL = 'D';
  if (pg) { const PH = headerIndex_(pg); payLoanL = a1col_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || 2); payAmtL = a1col_(colByAny_(PH, ['Monto Pagado']) || 4); }
  const PGN = CFG.SHEETS.PAYMENTS;
  const setF = (names, formula) => { const c = colByAny_(H, names); if (c) bs.getRange(r, c).setFormula(formula); };
  // Tasa tolerante a días (15/30/60) y meses (15/1/2): 15→25%, 30/1→50%, 60/2→100%.
  setF(['Tasa'], `=IF($${D}${r}="","",IF($${D}${r}=15,0.25,IF(OR($${D}${r}=30,$${D}${r}=1),0.5,IF(OR($${D}${r}=60,$${D}${r}=2,$${D}${r}=90,$${D}${r}=3),1,""))))`);
  setF(['Vencimiento', 'Fecha de Vencimiento'], `=IF($${F}${r}="","",IF($${D}${r}=15,$${F}${r}+15,IF(OR($${D}${r}=30,$${D}${r}=60,$${D}${r}=90),$${F}${r}+$${D}${r},EDATE($${F}${r},$${D}${r}))))`);
  setF(['Interés', 'Interes'], `=IF($${C}${r}="","",$${C}${r}*$${E}${r})`);
  setF(['Total a Pagar', 'Total a pagar'], `=IF($${C}${r}="","",$${C}${r}+$${Hh}${r})`);
  setF(['Total Pagado'], `=IF($${A}${r}="","",SUMIF('${PGN}'!$${payLoanL}:$${payLoanL},$${A}${r},'${PGN}'!$${payAmtL}:$${payAmtL}))`);
  // Saldo Pendiente = Total − Pagado (≥0), capital + interés. El recargo por mora va APARTE.
  setF(['Saldo Pendiente', 'Saldo Pendiente (hoy)'], `=IF($${A}${r}="","",MAX(0,$${I}${r}-$${J}${r}))`);
  // Recargo por Mora (acum.): usa el AJUSTE manual si es número (0 = condonar); si no, el auto.
  const moraFrag = (typeof moraLookupFormula_ === 'function') ? moraLookupFormula_('$' + A + r) : '0';
  setF(['Recargo por Mora (acum.)'], ADJ
    ? `=IF($${A}${r}="","",IF(ISNUMBER($${ADJ}${r}),MAX(0,$${ADJ}${r}),${moraFrag}))`
    : `=IF($${A}${r}="","",${moraFrag})`);
  // Saldo con Mora = Saldo Pendiente + Recargo por Mora = monto a pagar HOY.
  setF(['Saldo con Mora'], MOR ? `=IF($${A}${r}="","",N($${K}${r})+N($${MOR}${r}))` : `=IF($${A}${r}="","",N($${K}${r}))`);
  // Estado: PAGADO sólo cuando el "Saldo con Mora" llega a 0 (incluye la mora pendiente).
  setF(['Estado'], `=IF($${A}${r}="","",IF($${BM}${r}<=0,"${ST.PAID}",IF(OR(${cuotaVencidaExpr_('$' + A + r)},TODAY()>$${G}${r}),"${ST.OVERDUE}","${ST.ACTIVE}")))`);
}

/* ============================ PAGOS ============================ */
function reactivatePayments_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.PAYMENTS); if (!sh) return 'hoja ausente';
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Préstamo', 'ID Prestamo']);
  const fecC = colByAny_(H, ['Fecha de Pago']);
  const montoC = colByAny_(H, ['Monto Pagado']);
  const saldoC = colByAny_(H, ['Saldo Posterior']);
  const reciboC = colByAny_(H, ['Recibo Enviado']);
  if (!idC) return 'falta ID Préstamo';
  const b = dataBlock_(sh, idC); if (!b.count) return 'sin datos';

  const idL = a1col_(idC), montoL = a1col_(montoC);
  // Saldo posterior: Total a Pagar del préstamo − pagos acumulados hasta esta fila.
  if (saldoC && bs) {
    const BH = headerIndex_(bs);
    const bIdL = a1col_(colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || 1);
    const bTotL = a1col_(colByAny_(BH, ['Total a Pagar', 'Total a pagar']) || 9);
    const BSN = CFG.SHEETS.BORROWERS;
    setColFormulas_(sh, saldoC, b.first, b.count, r =>
      `=IF($${idL}${r}="","",IFERROR(INDEX('${BSN}'!$${bTotL}:$${bTotL},MATCH($${idL}${r},'${BSN}'!$${bIdL}:$${bIdL},0)),0)-SUMIFS($${montoL}$${b.first}:$${montoL}${r},$${idL}$${b.first}:$${idL}${r},$${idL}${r}))`);
    fmtCol_(sh, saldoC, b.first, b.count, CFG.CURRENCY_FMT);
  }
  fmtCol_(sh, montoC, b.first, b.count, CFG.CURRENCY_FMT);
  fmtCol_(sh, fecC, b.first, b.count, 'yyyy-mm-dd');

  // Validaciones: ID Préstamo (lista de préstamos) y Recibo Enviado (SÍ/NO).
  if (idC && bs) {
    const BH = headerIndex_(bs), bIdC = colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || 1;
    const src = bs.getRange(2, bIdC, CFG.MAX_ROWS, 1);
    sh.getRange(2, idC, CFG.MAX_ROWS, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(src, true).setAllowInvalid(true).build());
  }
  if (reciboC) sh.getRange(2, reciboC, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SÍ', 'NO'], true).setAllowInvalid(true).build());

  clearReactivarProtections_(sh);
  protectFormulaCols_(sh, [saldoC], b.first, b.count);
  return b.count + ' pagos · Saldo Posterior + validaciones + formato';
}

/* =========================== CLIENTES =========================== */
function reactivateClientes_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.CLIENTS); if (!sh) return 'hoja ausente';
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Cliente']);
  const validC = colByAny_(H, ['Correo válido', 'Correo valido']);
  const dupC = colByAny_(H, ['Duplicado']);
  if (!idC) return 'falta ID Cliente';
  const b = dataBlock_(sh, idC); if (!b.count) return 'sin datos';

  // Validaciones de lista.
  if (validC) sh.getRange(2, validC, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SÍ', 'NO'], true).setAllowInvalid(true).build());
  if (dupC) sh.getRange(2, dupC, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['', 'REVISAR', 'DUPLICADO'], true).setAllowInvalid(true).build());

  // Formato condicional por fila (correo inválido / duplicado).
  const width = sh.getLastColumn();
  const rng = sh.getRange(b.first, 1, b.count, width), rules = [];
  if (dupC) {
    const dL = a1col_(dupC);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${dL}${b.first}="DUPLICADO"`).setBackground('#f9cb9c').setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${dL}${b.first}="REVISAR"`).setBackground('#fff2cc').setRanges([rng]).build());
  }
  if (validC) {
    const vL = a1col_(validC);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${vL}${b.first}="NO"`).setBackground('#f4cccc').setRanges([rng]).build());
  }
  if (rules.length) sh.setConditionalFormatRules(rules);
  return b.count + ' clientes · validaciones + formato (correo inválido / duplicado)';
}

/* =========================== RESUMEN =========================== */
function reactivateResumen_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.SUMMARY); if (!sh) return 'hoja ausente';
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS); if (!bs) return 'falta Prestatarios';
  const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS);
  const H = headerIndex_(sh);
  const idC = colByAny_(H, ['ID Cliente']);
  const nameC = colByAny_(H, ['Prestatario']);
  const nC = colByAny_(H, ['Préstamos', 'Prestamos']);
  const capC = colByAny_(H, ['Capital Total']);
  const intC = colByAny_(H, ['Interés Total', 'Interes Total']);
  const pagC = colByAny_(H, ['Total Pagado']);
  const salC = colByAny_(H, ['Saldo Pendiente']);
  const estC = colByAny_(H, ['Estado']);
  if (!idC) return 'falta ID Cliente';
  const b = dataBlock_(sh, idC); if (!b.count) return 'sin datos';

  const BH = headerIndex_(bs), BSN = CFG.SHEETS.BORROWERS;
  const bCliL = a1col_(colByAny_(BH, ['ID Cliente']) || 2);
  const bCapL = a1col_(colByAny_(BH, ['Capital']) || 3);
  const bIntL = a1col_(colByAny_(BH, ['Interés', 'Interes']) || 8);
  const bPagL = a1col_(colByAny_(BH, ['Total Pagado']) || 10);
  const bSalL = a1col_(colByAny_(BH, ['Saldo Pendiente']) || 11);
  const A = a1col_(idC), G = a1col_(salC);

  if (nameC && cl) {
    const CH = headerIndex_(cl), CSN = CFG.SHEETS.CLIENTS;
    const cIdL = a1col_(colByAny_(CH, ['ID Cliente']) || 1), cNameL = a1col_(colByAny_(CH, ['Nombre']) || 2);
    setColFormulas_(sh, nameC, b.first, b.count, r => `=IFERROR(INDEX('${CSN}'!$${cNameL}:$${cNameL},MATCH($${A}${r},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`);
  }
  if (nC) setColFormulas_(sh, nC, b.first, b.count, r => `=COUNTIF('${BSN}'!$${bCliL}:$${bCliL},$${A}${r})`);
  if (capC) setColFormulas_(sh, capC, b.first, b.count, r => `=SUMIF('${BSN}'!$${bCliL}:$${bCliL},$${A}${r},'${BSN}'!$${bCapL}:$${bCapL})`);
  if (intC) setColFormulas_(sh, intC, b.first, b.count, r => `=SUMIF('${BSN}'!$${bCliL}:$${bCliL},$${A}${r},'${BSN}'!$${bIntL}:$${bIntL})`);
  if (pagC) setColFormulas_(sh, pagC, b.first, b.count, r => `=SUMIF('${BSN}'!$${bCliL}:$${bCliL},$${A}${r},'${BSN}'!$${bPagL}:$${bPagL})`);
  if (salC) setColFormulas_(sh, salC, b.first, b.count, r => `=SUMIF('${BSN}'!$${bCliL}:$${bCliL},$${A}${r},'${BSN}'!$${bSalL}:$${bSalL})`);
  if (estC) setColFormulas_(sh, estC, b.first, b.count, r => `=IF($${G}${r}<=0,"${ST.CLEARED}","${ST.ACTIVE}")`);

  // Formatos: conteo entero (arregla el "$1,00") y monedas.
  fmtCol_(sh, nC, b.first, b.count, '0');
  [capC, intC, pagC, salC].forEach(c => fmtCol_(sh, c, b.first, b.count, CFG.CURRENCY_FMT));

  clearReactivarProtections_(sh);
  protectFormulaCols_(sh, [nameC, nC, capC, intC, pagC, salC, estC], b.first, b.count);
  return b.count + ' clientes · agregados por ID Cliente (conteo entero + SUMIF)';
}

/* ============================ PANEL ============================ */
function reactivatePanel_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.PANEL); if (!sh) return 'hoja ausente';
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS); if (!bs) return 'falta Prestatarios';
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  const BH = headerIndex_(bs), BSN = CFG.SHEETS.BORROWERS;
  const bIdL = a1col_(colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || 1);
  const bCapL = a1col_(colByAny_(BH, ['Capital']) || 3);
  const bIntL = a1col_(colByAny_(BH, ['Interés', 'Interes']) || 8);
  const bSalL = a1col_(colByAny_(BH, ['Saldo Pendiente']) || 11);
  const bEstL = a1col_(colByAny_(BH, ['Estado']) || 12);
  const bFirmaL = a1col_(colByAny_(BH, ['Estado de Firma']) || 13);
  let payLoanL = 'B', payAmtL = 'D';
  if (pg) { const PH = headerIndex_(pg); payLoanL = a1col_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || 2); payAmtL = a1col_(colByAny_(PH, ['Monto Pagado']) || 4); }
  const PGN = CFG.SHEETS.PAYMENTS;

  const capitalPrestado = `SUMIF('${BSN}'!$${bIdL}:$${bIdL},"L-*",'${BSN}'!$${bCapL}:$${bCapL})`;
  const totalCobrado = pg ? `SUMIF('${PGN}'!$${payLoanL}:$${payLoanL},"L-*",'${PGN}'!$${payAmtL}:$${payAmtL})` : '0';

  const set = [];
  const put = (label, formula) => { const c = findLabelCell_(sh, label); if (c) { sh.getRange(c.row, c.valCol).setFormula('=' + formula); set.push(label); return c; } return null; };

  put('Préstamos activos', `COUNTIF('${BSN}'!$${bEstL}:$${bEstL},"${ST.ACTIVE}")`);
  put('Préstamos vencidos', `COUNTIF('${BSN}'!$${bEstL}:$${bEstL},"${ST.OVERDUE}")`);
  put('Préstamos pagados', `COUNTIF('${BSN}'!$${bEstL}:$${bEstL},"${ST.PAID}")`);
  put('Contratos sin firmar', `COUNTIF('${BSN}'!$${bFirmaL}:$${bFirmaL},"${SIGN.PENDING}")`);
  const capCell = put('Capital prestado', capitalPrestado);
  const cobCell = put('Total cobrado', totalCobrado);
  put('Interés contratado', `SUMIF('${BSN}'!$${bIdL}:$${bIdL},"L-*",'${BSN}'!$${bIntL}:$${bIntL})`);
  put('Saldo pendiente total', `SUMIF('${BSN}'!$${bIdL}:$${bIdL},"L-*",'${BSN}'!$${bSalL}:$${bSalL})`);

  // Efectivo disponible = Fondo − Capital prestado + Total cobrado. Un valor NEGATIVO (rojo) = sobregiro.
  const fondoCell = findLabelCell_(sh, 'Fondo total para prestar');
  if (fondoCell && capCell && cobCell) {
    const ed = findLabelCell_(sh, 'Efectivo disponible');
    if (ed) {
      const cell = sh.getRange(ed.row, ed.valCol);
      cell.setFormula(`=${fondoCell.valA1}-${capCell.valA1}+${cobCell.valA1}`).setFontColor('#38761d').setFontWeight('bold');
      sh.setConditionalFormatRules([redIfNegativeRule_(cell)]);
      set.push('Efectivo disponible');
    }
  }
  return set.length ? ('indicadores: ' + set.join(', ')) : 'sin etiquetas reconocidas';
}

/* ========================= ESTADÍSTICAS ========================= */
function reactivateStats_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.STATS); if (!sh) return 'hoja ausente';
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS); if (!bs) return 'falta Prestatarios';
  const pg = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  const rs = ss.getSheetByName(CFG.SHEETS.SUMMARY);
  const panel = ss.getSheetByName(CFG.SHEETS.PANEL);
  const BH = headerIndex_(bs), BSN = CFG.SHEETS.BORROWERS;
  const bIdL = a1col_(colByAny_(BH, ['ID Préstamo', 'ID Prestamo']) || 1);
  const bCapL = a1col_(colByAny_(BH, ['Capital']) || 3);
  const bIntL = a1col_(colByAny_(BH, ['Interés', 'Interes']) || 8);
  const bTotL = a1col_(colByAny_(BH, ['Total a Pagar', 'Total a pagar']) || 9);
  const bSalL = a1col_(colByAny_(BH, ['Saldo Pendiente']) || 11);
  const bEstL = a1col_(colByAny_(BH, ['Estado']) || 12);
  let payLoanL = 'B', payAmtL = 'D';
  if (pg) { const PH = headerIndex_(pg); payLoanL = a1col_(colByAny_(PH, ['ID Préstamo', 'ID Prestamo']) || 2); payAmtL = a1col_(colByAny_(PH, ['Monto Pagado']) || 4); }
  const PGN = CFG.SHEETS.PAYMENTS;

  const nLoans = `COUNTIF('${BSN}'!$${bIdL}:$${bIdL},"L-*")`;
  const nVenc = `COUNTIF('${BSN}'!$${bEstL}:$${bEstL},"${ST.OVERDUE}")`;
  const capital = `SUMIF('${BSN}'!$${bIdL}:$${bIdL},"L-*",'${BSN}'!$${bCapL}:$${bCapL})`;
  const interes = `SUMIF('${BSN}'!$${bIdL}:$${bIdL},"L-*",'${BSN}'!$${bIntL}:$${bIntL})`;
  const totalAPagar = `SUMIF('${BSN}'!$${bIdL}:$${bIdL},"L-*",'${BSN}'!$${bTotL}:$${bTotL})`;
  const cobrado = pg ? `SUMIF('${PGN}'!$${payLoanL}:$${payLoanL},"L-*",'${PGN}'!$${payAmtL}:$${payAmtL})` : '0';
  const moraSaldo = `SUMIF('${BSN}'!$${bEstL}:$${bEstL},"${ST.OVERDUE}",'${BSN}'!$${bSalL}:$${bSalL})`;

  // Fondo: desde Configuración si existe; si no, desde la celda del Panel.
  const cfg = ss.getSheetByName(CFG.SHEETS.SETTINGS);
  let fondoExpr = null;
  if (cfg) fondoExpr = `IFERROR(VLOOKUP("Fondo total para prestar",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE),0)`;
  else if (panel) { const fc = findLabelCell_(panel, 'Fondo total para prestar'); if (fc) fondoExpr = `'${CFG.SHEETS.PANEL}'!${fc.valA1}`; }

  // Exposición máxima por cliente (desde Resumen, excluye la fila TOTAL con "C-*").
  let expMax = null;
  if (rs) {
    const RH = headerIndex_(rs), RSN = CFG.SHEETS.SUMMARY;
    const rIdL = a1col_(colByAny_(RH, ['ID Cliente']) || 1), rCapL = a1col_(colByAny_(RH, ['Capital Total']) || 4);
    expMax = `MAXIFS('${RSN}'!$${rCapL}:$${rCapL},'${RSN}'!$${rIdL}:$${rIdL},"C-*")`;
  }

  const set = [];
  const put = (label, formula, fmt) => { const c = findLabelCell_(sh, label); if (c) { const cell = sh.getRange(c.row, c.valCol); cell.setFormula('=' + formula); if (fmt) cell.setNumberFormat(fmt); set.push(label); } };

  put('Tasa de morosidad', `IFERROR(${nVenc}/${nLoans},0)`, '0.0%');
  put('Tasa de recuperación', `IFERROR(${cobrado}/${totalAPagar},0)`, '0.0%');
  put('Rendimiento sobre capital', `IFERROR(${interes}/${capital},0)`, '0.0%');
  if (fondoExpr) put('Utilización neta del fondo', `IFERROR((${capital}-${cobrado})/(${fondoExpr}),0)`, '0.0%');
  put('Cartera en mora', moraSaldo, CFG.CURRENCY_FMT);
  put('Ticket promedio', `IFERROR(${capital}/${nLoans},0)`, CFG.CURRENCY_FMT);
  if (expMax) put('Exposición máxima por cliente', expMax, CFG.CURRENCY_FMT);

  return set.length ? ('indicadores: ' + set.join(', ')) : 'sin etiquetas reconocidas';
}

/* ===================== NUEVOS PRESTATARIOS ===================== */
/**
 * Reactiva la hoja de solicitudes: asegura la estructura base (encabezados,
 * casillas Verificado?/Rechazar?/Verificar BCRA?, Parámetro BCRA, formatos),
 * agrega la validación de Plazo en días (V-13) y el formato condicional por
 * estado de la solicitud. La APROBACIÓN/RECHAZO se procesa con
 * procesarNuevosPrestatarios() (ver más abajo).
 */
function reactivateNew_(ss) {
  let sh = ss.getSheetByName(CFG.SHEETS.NEW);
  const nueva = !sh;
  // setupNew_ crea/normaliza encabezados, casillas y Parámetro BCRA (no borra filas).
  if (typeof setupNew_ === 'function') { setupNew_(ss); sh = ss.getSheetByName(CFG.SHEETS.NEW); }
  if (!sh) return 'no se pudo crear';
  const H = headerIndex_(sh);

  // V-13 — Plazo en días (15/30/60). allowInvalid: no bloquea datos históricos.
  const plzC = colByAny_(H, ['Plazo (días)', 'Plazo (dias)', 'Plazo (meses)', 'Plazo']);
  // El modelo migrado usa DÍAS; corrige el encabezado heredado "Plazo (meses)".
  if (plzC && hkey_(sh.getRange(1, plzC).getValue()) !== hkey_('Plazo (días)')) sh.getRange(1, plzC).setValue('Plazo (días)');
  if (plzC) sh.getRange(2, plzC, CFG.MAX_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['15', '30', '60', '90'], true).setAllowInvalid(true)
      .setHelpText('Plazo en días: 15, 30 o 60.').build());

  // Formato condicional por estado de la solicitud (usa toda la fila de datos).
  const verC = colByAny_(H, ['Verificado?']), rejC = colByAny_(H, ['Rechazar?']);
  const width = sh.getMaxColumns(), rng = sh.getRange(2, 1, CFG.MAX_ROWS, width), rules = [];
  if (rejC) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${a1col_(rejC)}2=TRUE`).setBackground('#f4cccc').setRanges([rng]).build());
  if (verC) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${a1col_(verC)}2=TRUE`).setBackground('#d9ead3').setRanges([rng]).build());
  if (rules.length) sh.setConditionalFormatRules(rules);

  return (nueva ? 'creada' : 'actualizada') + ' · casillas + Plazo(15/30/60) + Parámetro BCRA + formato condicional';
}

/**
 * Procesa las solicitudes marcadas en "Nuevos Prestatarios" contra el esquema
 * MIGRADO (días + Clientes normalizado). Reemplaza al onEdit month-based:
 *   • Verificado? = TRUE  → valida (V-02/V-08/V-09/V-10/V-11/V-13), crea/reutiliza
 *     el cliente (11 col), agrega el préstamo a "Prestatarios", deja Estado de
 *     Firma = PENDIENTE, mueve las fotos a la carpeta y borra la fila.
 *   • Rechazar?  = TRUE  → archiva en "Rechazados" y avisa por correo.
 * Ejecutar manualmente (o desde un menú). Devuelve un resumen.
 */
function procesarNuevosPrestatarios() {
  return guard_('procesarNuevosPrestatarios', function () {
    const ss = getSS_(), sh = ss.getSheetByName(CFG.SHEETS.NEW);
    if (!sh || sh.getLastRow() < 2) return 'No hay solicitudes para procesar.';
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    if (!bs) throw new Error('Falta la hoja "Prestatarios".');
    const m = headerMap_(sh);
    const verCol = m['Verificado?'], rejCol = m['Rechazar?'], ovrCol = m['Anular límites'];
    const last = sh.getLastRow();
    const out = { aprobados: [], rechazados: [], rebotados: [] };

    // De abajo hacia arriba: borrar/insertar filas no corre los índices pendientes.
    for (let row = last; row >= 2; row--) {
      const idCell = String(sh.getRange(row, m['Nombre completo'] || 2).getValue()).trim();
      const verOn = verCol && sh.getRange(row, verCol).getValue() === true;
      const rejOn = rejCol && sh.getRange(row, rejCol).getValue() === true;
      if (!verOn && !rejOn) continue;
      if (rejOn) { rejectApplicant_(sh, row); out.rechazados.push(idCell); continue; }
      const override = ovrCol ? (sh.getRange(row, ovrCol).getValue() === true) : false;
      const res = approveApplicantV2_(ss, bs, sh, row, m, override);
      if (res.ok) out.aprobados.push(res.loanId + ' ' + res.name);
      else out.rebotados.push((idCell || 'fila ' + row) + ': ' + res.msg);
    }

    // Si se aprobó algo, reinstala fórmulas/formatos sobre el bloque ampliado.
    if (out.aprobados.length) { try { reactivateBorrowers_(ss); } catch (e) { logError_('procesarNuevos:reactivar', e); } }
    SpreadsheetApp.flush();
    const parts = [];
    if (out.aprobados.length) parts.push('Aprobados: ' + out.aprobados.join(', '));
    if (out.rechazados.length) parts.push('Rechazados: ' + out.rechazados.join(', '));
    if (out.rebotados.length) parts.push('Rebotados: ' + out.rebotados.join(' | '));
    const msg = parts.length ? parts.join('\n') : 'Ninguna solicitud tenía Verificado?/Rechazar? tildado.';
    try { SpreadsheetApp.getActive().toast(msg, 'Solicitudes procesadas', 8); } catch (e) { }
    return msg;
  });
}

/** Inserta una fila de préstamo justo debajo del último "L-…" (preserva la fila TOTAL). */
function insertBorrowerRow_(bs, idCol) {
  const maxR = bs.getLastRow(); let lastLoan = 0;
  if (maxR >= 2) {
    const vals = bs.getRange(2, idCol, maxR - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) if (/^L-/i.test(String(vals[i][0]).trim())) lastLoan = i + 2;
  }
  if (lastLoan < 2) return Math.max(maxR + 1, 2);
  bs.insertRowAfter(lastLoan);
  return lastLoan + 1;
}

/** Aprueba una solicitud contra el esquema migrado. Devuelve {ok, loanId, name, msg}. */
function approveApplicantV2_(ss, bs, nb, row, m, override) {
  // Lee la fila completa de una sola vez (antes: ~12 getRange().getValue()).
  const rowVals = nb.getRange(row, 1, 1, nb.getLastColumn()).getValues()[0];
  const g = t => m[t] ? rowVals[m[t] - 1] : '';
  const resCol = m['Resultado'], verCol = m['Verificado?'];
  const fail = msg => { if (resCol) nb.getRange(row, resCol).setValue('⚠ ' + msg); if (verCol) nb.getRange(row, verCol).setValue(false); return { ok: false, msg: msg }; };

  const name = String(g('Nombre completo')).trim();
  const email = String(g('Correo')).trim();
  const dni = String(g('DNI')).trim();
  const phone = String(g('Teléfono')).trim();
  const amount = Number(g('Monto Solicitado')) || 0;
  const term = Number(g('Plazo (días)') || g('Plazo (meses)') || g('Plazo')) || 0;
  const natId = g('Foto del frente del DNI'), workId = g('Foto del dorso del DNI');

  // V-19 / V-01: identidad mínima.
  const dnV = vDni_(dni);
  if (!name || !dnV.ok) return fail('Falta nombre o DNI válido');
  // V-13: plazo en días.
  if (termRateDays_(term) == null) return fail('Plazo debe ser 15, 30 o 60 días');
  if (!amount) return fail('Monto inválido');
  // V-08: ambas fotos.
  if (!natId || !workId) return fail('Faltan las fotos del DNI (frente/dorso)');
  // V-02: el DNI no puede pertenecer a otra persona.
  if (typeof dniBelongsToOtherName_ === 'function' && dniBelongsToOtherName_(dnV.norm, name)) return fail('El DNI ya está registrado a nombre de otra persona');
  // V-33: cliente bloqueado (correo/DNI/CUIL/teléfono). ABSOLUTO: se chequea temprano
  // (antes que referencias/BCRA, para que el mensaje de bloqueo prevalezca) y NO se
  // anula con "Anular límites". Desbloquear = vaciar "Bloqueado" en Clientes.
  if (typeof findClienteBloqueado_ === 'function') {
    const blk = findClienteBloqueado_({ email: email, dni: dnV.norm, phone: phone, cuil: String(g('CUIL') || '') });
    if (blk) return fail('V-33 — Cliente bloqueado (' + (blk.id || blk.dniNorm) + '). No se anula con «Anular límites».' + (blk.bloqueoMotivo ? ' Motivo: ' + blk.bloqueoMotivo : ''));
  }

  // V-27: DOS referencias con teléfono válido son OBLIGATORIAS para aprobar.
  // NO se anula con "Anular límites" (override sólo omite fondos/concentración/tope/BCRA).
  const r1Name = String(g('Ref 1 Nombre') || '').trim(), r2Name = String(g('Ref 2 Nombre') || '').trim();
  const r1Ph = vPhone_(g('Ref 1 Teléfono')), r2Ph = vPhone_(g('Ref 2 Teléfono'));
  if (!r1Name || !r1Ph.ok || !r2Name || !r2Ph.ok) return fail('Faltan dos referencias con teléfono válido (obligatorias; no se anulan con «Anular límites»)');

  // BCRA MANUAL: la verificación ya NO se corre automáticamente al aprobar. Si el
  // prestamista la corrió antes (casilla "Verificar BCRA?") y quedó "RECHAZAR",
  // se respeta salvo que se apruebe con anulación de límites (override).
  const decision = String(g('Decisión') || '').trim().toUpperCase();
  if (decision === 'RECHAZAR' && !override) return fail('El BCRA marca RECHAZAR — requiere autorización expresa (usá "Anular límites")');

  // V-09 / V-10 / tope de 2 préstamos: fondos y concentración (sobre el cliente
  // existente, si lo hay). Con override se omiten estos topes.
  const existing = (typeof clienteByDni_ === 'function') ? clienteByDni_(dnV.norm) : null;
  if (typeof validateApprovalV2_ === 'function') {
    // dnV.norm habilita V-32 (escalera de graduación por historial de repago), que
    // aplica también a prestatarios nuevos (aún sin ID Cliente). El 5º argumento
    // amplía la coincidencia del bloqueo V-33 a correo/teléfono/CUIL.
    const chk = validateApprovalV2_(existing ? existing.id : null, amount, override, dnV.norm,
      { email: email, phone: phone, cuil: String(g('CUIL') || '') });
    if (!chk.ok) return fail(chk.msg);
  }

  // Cliente (reutiliza o crea). Se guardan también CUIL y Dirección de la solicitud
  // para que queden en "Clientes" y se recuperen en próximas solicitudes.
  const cuil = String(g('CUIL') || '').trim();
  const address = String(g('Dirección') || '').trim();
  let clienteId;
  if (typeof getOrCreateClienteV2_ === 'function') {
    clienteId = getOrCreateClienteV2_(name, email, dnV.norm, phone, { cuil: cuil, address: address });
    if (existing && existing.row) { // cliente reutilizado → completar CUIL/Dirección si faltaban o cambiaron
      if (cuil && typeof updateClienteCuil_ === 'function' && String(existing.cuil || '').replace(/\D/g, '') !== cuil.replace(/\D/g, '')) updateClienteCuil_(existing.row, cuil);
      if (address && typeof updateClienteDireccion_ === 'function' && String(existing.address || '').trim() !== address) updateClienteDireccion_(existing.row, address);
    }
  } else {
    clienteId = existing ? existing.id : '';
  }

  // Alta del préstamo (columnas base; las calculadas quedan a cargo de las fórmulas).
  const H = headerIndex_(bs);
  const idCol = colByAny_(H, ['ID Préstamo', 'ID Prestamo']) || 1;
  const cliCol = colByAny_(H, ['ID Cliente']) || 2;
  const capCol = colByAny_(H, ['Capital']) || 5;
  const plzCol = colByAny_(H, ['Plazo (días)', 'Plazo (dias)', 'Plazo (meses)', 'Plazo']) || 6;
  const fecCol = colByAny_(H, ['Fecha Préstamo', 'Fecha de Préstamo', 'Fecha Prestamo']) || 8;
  const firmaCol = colByAny_(H, ['Estado de Firma']) || 18;
  // Si "Prestatarios" muestra Nombre/DNI como columnas, hay que poblarlas en la fila nueva
  // (el instalador de fórmulas migrado no las toca): antes quedaban en blanco al aprobar.
  const nameCol = colByAny_(H, ['Nombre']);
  const dniCol = colByAny_(H, ['DNI']);
  const loanId = nextLoanId_(), tRow = insertBorrowerRow_(bs, idCol);
  bs.getRange(tRow, idCol).setValue(loanId);
  if (override) bs.getRange(tRow, idCol).setNote('Aprobado con anulación de límites (override) el ' + fmtDate_(new Date()) + '.');
  bs.getRange(tRow, cliCol).setValue(clienteId);
  bs.getRange(tRow, capCol).setValue(amount);
  bs.getRange(tRow, plzCol).setValue(term);
  bs.getRange(tRow, fecCol).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  if (firmaCol) bs.getRange(tRow, firmaCol).setValue(SIGN.PENDING);
  // Fórmulas calculadas (Tasa/Venc/Interés/Total/Pagado/Saldo/Estado) en la fila nueva,
  // para que el préstamo quede completo al instante (Total ≠ $0 en el contrato/panel).
  try { writeBorrowerRowFormulas_(bs, tRow); } catch (e) { logError_('approveApplicantV2_:formulasFila', e); }
  // V-31 — cronograma de cuotas: tramos chicos = 1 pago; tramo grande (>$300.000) = 3 cuotas mensuales.
  try {
    const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || (typeof setupCuotas_ === 'function' ? setupCuotas_(ss) : null);
    if (cs && typeof cuotasForLoan_ === 'function' && typeof resolveTerm_ === 'function') {
      const ti = resolveTerm_(amount, term), total = round2_(amount * (1 + ti.rate)); // respeta el plazo elegido (15/30 en tramo chico)
      const dueDays = ti.cuotas === 3 ? [30, 60, 90] : [ti.days];
      cuotasForLoan_(cs, loanId, new Date(), total, ti.cuotas, dueDays);
    }
  } catch (e) { logError_('approveApplicantV2_:cuotas', e); }
  // Nombre/DNI: se resuelven del cliente por "ID Cliente" con INDEX/MATCH (modelo normalizado);
  // si no se puede resolver "Clientes", se escribe el valor literal como respaldo.
  if (nameCol || dniCol) {
    try {
      const cl = ss.getSheetByName(CFG.SHEETS.CLIENTS), CSN = CFG.SHEETS.CLIENTS;
      const CH = cl ? headerIndex_(cl) : null, cliL = a1col_(cliCol);
      if (cl && CH) {
        const cIdL = a1col_(colByAny_(CH, ['ID Cliente']) || 1);
        if (nameCol) { const cNameL = a1col_(colByAny_(CH, ['Nombre']) || 2);
          bs.getRange(tRow, nameCol).setFormula(`=IFERROR(INDEX('${CSN}'!$${cNameL}:$${cNameL},MATCH($${cliL}${tRow},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`); }
        if (dniCol) { const cDniL = a1col_(colByAny_(CH, ['DNI']) || 4);
          bs.getRange(tRow, dniCol).setFormula(`=IFERROR(INDEX('${CSN}'!$${cDniL}:$${cDniL},MATCH($${cliL}${tRow},'${CSN}'!$${cIdL}:$${cIdL},0)),"")`); }
      } else {
        if (nameCol) bs.getRange(tRow, nameCol).setValue(name);
        if (dniCol) bs.getRange(tRow, dniCol).setValue(dnV.norm);
      }
    } catch (e) { logError_('approveApplicantV2_:nombreDni', e); }
  }

  // Mueve las fotos a la carpeta del prestatario.
  try {
    const folder = borrowerFolder_(borrowerFolderName_(name, dnV.norm));
    [natId, workId].forEach(u => extractDriveIds_(u).forEach(id => { try { DriveApp.getFileById(id).moveTo(folder); } catch (e) { } }));
    bs.getRange(tRow, cliCol).setNote('📁 ' + folder.getUrl());
  } catch (e) { logError_('approveApplicantV2_:fotos', e); }

  // Aviso al prestatario: aprobado + ENLACE para revisar y firmar el contrato.
  // El enlace apunta a la página de firma (?page=firmar&loan=…); requiere la
  // "URL de la app web" en Configuración (o un despliegue publicado de la app).
  if (email) {
    try {
      const rate = termRateDays_(term), interest = round2_(amount * rate), total = round2_(amount * (1 + rate));
      const payMethod = String(g('Forma de Pago') || '').trim();
      const link = (typeof signingLink_ === 'function') ? signingLink_(loanId) : '';
      const btn = link
        ? '<p style="margin:18px 0"><a href="' + link + '" style="background:#1c4587;color:#fff;text-decoration:none;padding:12px 22px;border-radius:6px;font-weight:bold">Revisar y firmar mi contrato</a></p>' +
          '<p style="font-size:12px;color:#888">O copie este enlace: ' + esc_(link) + '</p>' +
          '<p style="font-size:12px;color:#888">Para firmar necesitará su DNI (' + esc_(dni) + ').</p>'
        : '<p><b>Su contrato está listo para firmar.</b> Solicite el enlace de firma al prestamista.</p>';
      // Contrato SIN firmar (PDF): se adjunta al correo y queda archivado en la carpeta del prestatario.
      let contractFile = null;
      try {
        const loanObj = {
          loanId: loanId, name: name, dni: dni, email: email, phone: phone,
          cuil: cuil || '', direccion: address || '', payMethod: payMethod || '',
          principal: amount, term: term, interest: interest, totalDue: total,
          loanDate: new Date(), dueDate: addDays_(new Date(), termDays_(term)),
        };
        contractFile = makeAgreementFile_(loanObj);
      } catch (e) { logError_('approveApplicantV2_:contratoPDF', e); }
      sendBrandedEmail_(email, 'Revise y firme su contrato — ' + loanId,
        'Estimado/a ' + name + ', su préstamo ' + loanId + ' fue aprobado por ' + fmtMoney_(amount) + '. ' +
        'Adjuntamos el contrato sin firmar: léalo en detalle antes de firmar. ' + (link ? ('Firme aquí: ' + link) : 'Queda pendiente de firma.'),
        '<p>Estimado/a ' + esc_(name) + ',</p>' +
        '<p>Su solicitud fue <b>aprobada</b> (' + esc_(loanId) + '). <b>Adjuntamos el contrato sin firmar</b> para su revisión.</p>' +
        '<p><b>Le pedimos leer el contrato completo y en detalle antes de firmarlo</b>, en particular el costo total del crédito (Cláusula 3) y el régimen de mora e interés punitorio (Cláusula 8). Recién después de leerlo y firmarlo se libera el préstamo.</p>' +
        '<table style="border-collapse:collapse">' +
        row_('Capital', fmtMoney_(amount)) + row_('Plazo', term + ' días (' + round2_(rate * 100) + '%)') +
        row_('Interés', fmtMoney_(interest)) + row_('Total a devolver', '<b>' + fmtMoney_(total) + '</b>') +
        '</table>' + btn,
        contractFile ? { attachments: [contractFile.getAs('application/pdf')] } : {});
    } catch (e) { logError_('approveApplicantV2_:email', e); }
  }

  nb.deleteRow(row);
  return { ok: true, loanId: loanId, name: name, overridden: !!override };
}

/**
 * Entrada llamada por el onEdit al tildar "Verificado?": aprueba la solicitud
 * y la MUEVE a "Prestatarios" (esquema migrado, días + Clientes normalizado),
 * corriendo la verificación BCRA (V-11) y demás validaciones. Instala las
 * fórmulas de la fila nueva. Rebota (avisa) si alguna validación falla.
 */
function verifyApplicantV2_(nb, row, override) {
  return guard_('verifyApplicantV2_', function () {
    const ss = getSS_(), bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    if (!bs) throw new Error('Falta la hoja "Prestatarios".');
    const res = approveApplicantV2_(ss, bs, nb, row, headerMap_(nb), override);
    if (res.ok) {
      try { reactivateBorrowers_(ss); } catch (e) { logError_('verifyApplicantV2_:reactivar', e); }
      try { ss.toast('Aprobado ' + res.name + ' → ' + res.loanId + (res.overridden ? ' · límites anulados' : '') + ' · movido a Prestatarios (pendiente de firma).', '✔ Aprobado', 6); } catch (e) { }
    } else {
      try { ss.toast(res.msg, '⚠ No aprobado', 8); } catch (e) { }
    }
    return res;
  });
}

/**
 * Instala el disparador onEdit (necesario para que las casillas
 * "Verificado?"/"Rechazar?"/"Verificar BCRA?" ejecuten acciones con permisos de
 * Drive/Gmail/UrlFetch). Ejecutar una vez y autorizar. Idempotente.
 */
function activarDisparadores() {
  return guard_('activarDisparadores', function () {
    if (typeof ensureTriggers_ === 'function') { ensureTriggers_(); return 'Disparadores instalados (onEdit + tareas diarias).'; }
    const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onEditInstallable');
    if (!has) ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
    return 'Disparador onEdit instalado.';
  });
}
