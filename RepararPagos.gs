/************************************************************
 * RepararPagos.gs — arreglo de la ubicación de los pagos
 * ----------------------------------------------------------
 * Archivo ADICIONAL del mismo proyecto de Apps Script (comparte el
 * scope global con LoanManagerV2.gs). NO redefine funciones existentes.
 *
 * PROBLEMA
 *   recordPayment() usaba psh.appendRow(...). En "Pagos" las casillas
 *   (insertCheckboxes de "Enviar/Generar recibo") están pre-creadas hasta
 *   la fila MAX_ROWS+1 (301), así que getLastRow() devuelve 301 y appendRow
 *   escribe en la fila 302 — el pago queda al final, con un hueco enorme,
 *   en vez de JUSTO DESPUÉS del último pago.
 *
 * SOLUCIÓN
 *   • appendPaymentRowFixed_ / firstEmptyPaymentRow_ : insertan en la primera
 *     fila realmente vacía del bloque (A y B sin valor). recordPayment ahora
 *     los usa en lugar de appendRow (ver LoanManagerV2.gs).
 *   • compactarPagos() : arreglo de UNA SOLA VEZ para los pagos que ya quedaron
 *     mal ubicados — los sube para que queden contiguos desde la fila 2,
 *     reinstala la fórmula de "Saldo Posterior" (depende de la fila) y los enlaces
 *     de WhatsApp. Correr a mano desde el editor o el menú.
 ************************************************************/

/** Primera fila VACÍA del bloque de datos de "Pagos" (ID Pago e ID Préstamo sin valor), para que un
 *  pago nuevo entre justo después del último y NO al final de la hoja pre-formateada. */
function firstEmptyPaymentRow_(psh) {
  psh = psh || getSS_().getSheetByName(CFG.SHEETS.PAYMENTS);
  const last = psh.getLastRow();
  if (last < 2) return 2;
  const vals = psh.getRange(2, 1, last - 1, PP.LOAN_ID).getValues(); // columnas A (ID Pago) y B (ID Préstamo)
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][PP.PAY_ID - 1] || '').trim() === '' && String(vals[i][PP.LOAN_ID - 1] || '').trim() === '') return i + 2;
  }
  return last + 1;
}

/** Escribe un pago en la fila correcta (la primera vacía) en vez de al final. Devuelve la fila usada.
 *  Lo usa recordPayment(); el llamador instala la fórmula de "Saldo Posterior" y los formatos. */
function appendPaymentRowFixed_(psh, values) {
  psh = psh || getSS_().getSheetByName(CFG.SHEETS.PAYMENTS);
  const row = firstEmptyPaymentRow_(psh);
  psh.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}

/**
 * ONE-OFF: compacta la hoja "Pagos" — sube los pagos que quedaron al final (tras filas vacías) para
 * que queden contiguos desde la fila 2, conservando el ORDEN actual. Reinstala la fórmula de
 * "Saldo Posterior" (que depende de la posición de la fila) y reconstruye "Avisar pago" (WhatsApp).
 * Idempotente: si ya está compacto, no cambia nada. Correr a mano una vez.
 */
function compactarPagos() {
  return guard_('compactarPagos', function () {
    const ss = getSS_(), psh = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    if (!psh) throw new Error('No existe la hoja "' + CFG.SHEETS.PAYMENTS + '".');
    const last = psh.getLastRow();
    if (last < 2) return 'Pagos: sin datos para compactar.';
    const W = PP.DNI; // bloque de datos A..H (ID Pago .. DNI)
    const data = psh.getRange(2, 1, last - 1, W).getValues();
    // Filas con pago real (tienen ID Pago o ID Préstamo), en su orden actual.
    const filled = data.filter(function (r) {
      return String(r[PP.PAY_ID - 1] || '').trim() !== '' || String(r[PP.LOAN_ID - 1] || '').trim() !== '';
    });
    const n = filled.length;
    if (n === 0) return 'Pagos: no hay pagos cargados.';

    // ¿Ya está compacto? (las primeras n filas son justamente las llenas y no hay nada debajo)
    let yaCompacto = (last - 1) === n;
    for (let i = 0; yaCompacto && i < n; i++) {
      if (!(String(data[i][PP.PAY_ID - 1] || '').trim() || String(data[i][PP.LOAN_ID - 1] || '').trim())) yaCompacto = false;
    }

    // Reescribe el bloque contiguo desde la fila 2 y limpia lo que sobra debajo.
    psh.getRange(2, 1, n, W).setValues(filled);
    if (last - 1 > n) psh.getRange(2 + n, 1, (last - 1) - n, W).clearContent();

    // Reinstala la fórmula de "Saldo Posterior" por fila (referencia acumulada hasta esa fila) y formatos.
    for (let r = 2; r < 2 + n; r++) psh.getRange(r, PP.BALANCE).setFormula(paymentBalanceFormula_(r));
    psh.getRange(2, PP.DATE, n, 1).setNumberFormat('yyyy-mm-dd');
    psh.getRange(2, PP.AMOUNT, n, 2).setNumberFormat(CFG.CURRENCY_FMT);

    // Reconstruye los enlaces "Avisar pago 📲" (quedan alineados a la nueva posición).
    try { if (typeof fillPaymentWaLinks_ === 'function') fillPaymentWaLinks_(ss); } catch (e) { logError_('compactarPagos:wa', e); }
    SpreadsheetApp.flush();

    const msg = yaCompacto
      ? ('Pagos ya estaban contiguos (' + n + '); se reinstalaron fórmulas y enlaces.')
      : ('Pagos compactados: ' + n + ' pago(s) contiguos desde la fila 2.');
    try { ss.toast(msg, '💳 Reparar Pagos', 6); } catch (e) { }
    return msg;
  });
}
