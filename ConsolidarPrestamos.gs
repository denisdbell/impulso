/**
 * ConsolidarPrestamos.gs — Script A MEDIDA para CONSOLIDAR (refinanciar) varios préstamos de un
 * mismo cliente en UNO nuevo, con un plan de cuotas mensuales y enlace de firma del contrato.
 *
 * Caso pedido — Nancy Roxana Riccobene (C-0016): fusionar L-0015 y L-0019 en un préstamo nuevo.
 *   • Monto = SALDO LIMPIO combinado (capital + interés − pagos, SIN mora): $50.000 + $1.500.000
 *     = $1.550.000 (coincide con "Saldo Posterior" del pago P-0044).
 *   • 0% — sólo reprogramación: 4 cuotas IGUALES de $387.500.
 *   • Vencen el DÍA 10 de cada mes: 10-nov-2026, 10-dic-2026, 10-ene-2027, 10-feb-2027.
 *   • Recargo por mora ESTÁNDAR (5% diario) automático → "Mora (ajuste)" se deja VACÍA.
 *   • Se genera token + "Enlace de Firma"; el contrato queda PENDIENTE de firma. NO se envía correo.
 *   • L-0015 y L-0019 se CIERRAN (SALDADO): pago de traslado por su saldo, mora condonada (ajuste=0)
 *     y se borran sus cuotas (el cronograma viejo queda obsoleto). No entra ni sale caja real: el
 *     pago de traslado compensa, en "Efectivo disponible", el capital del préstamo nuevo (neto 0).
 *
 * Reutiliza helpers de LoanManagerV2.gs/Reactivar.gs (getSS_, nextLoanId_, firstEmptyBorrowerRow_,
 * cuotasForLoan_, deleteCuotasRowsForLoan_, writeBorrowerRowFormulas_, ensureLoanToken_,
 * signingLink_, makeAgreementFile_, markAgreementSent_, paymentBalanceFormula_, …). NO redefine
 * ninguna función existente (todo comparte un único ámbito global).
 *
 * CÓMO USAR: pegá este archivo en el editor de Apps Script (Extensions ▸ Apps Script), elegí la
 * función `consolidarNancyL15L19` en la barra de funciones y ejecutá UNA sola vez (no es
 * idempotente: una segunda corrida crearía OTRO préstamo nuevo). Verificá el resultado en la hoja.
 */

/**
 * Regenera el PDF del contrato (y refresca el "Enlace de Firma") de un préstamo YA existente, SIN
 * volver a consolidar ni enviar correo. Úsalo tras corregir el cronograma del contrato: ahora el
 * contrato toma las cuotas REALES de la hoja "Cuotas" (p. ej. las 4 cuotas de $387.500 de L-0083),
 * en lugar de derivarlas del monto. El PDF viejo (6 cuotas) queda reemplazado por el nuevo link.
 * @param {string} [loanId='L-0083'] ID del préstamo consolidado (usá el que mostró la corrida).
 * @param {boolean} [enviarCorreo=false] true para reenviarle el enlace por correo al prestatario.
 */
function regenerarContratoConsolidado_(loanId, enviarCorreo) {
  return guard_('regenerarContratoConsolidado_', function () {
    loanId = String(loanId || 'L-0083').trim();
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo ' + loanId + '.');
    ensureLoanToken_(loanId, false);        // asegura el token (no lo rota salvo que falte)
    const file = makeAgreementFile_(loan);  // PDF con el cronograma REAL (cuotas de la hoja)
    markAgreementSent_(loanId, file);
    if (enviarCorreo && loan.email) emailSigningRequest_(loan, file);
    const enlace = signingLink_(loanId);
    try { getSS_().toast(loanId + ': contrato regenerado' + (enviarCorreo ? ' y reenviado por correo.' : ' (sin correo).'), '🧾 Contrato', 6); } catch (e) {}
    return { loanId: loanId, enlaceFirma: enlace };
  });
}

/** Atajo: regenera el contrato de L-0083 (Nancy) con las 4 cuotas reales, SIN enviar correo. */
function regenerarContratoNancy() { return regenerarContratoConsolidado_('L-0083', false); }

/** Atajo one-shot para el caso concreto pedido (L-0015 + L-0019 → nuevo préstamo de Nancy). */
function consolidarNancyL15L19() {
  const base = new Date(2026, 10, 10); // 10 de noviembre de 2026 (mes 10 = noviembre, base 0)
  const vencimientos = [0, 1, 2, 3].map(function (k) { return addMonths_(base, k); });
  const out = consolidarPrestamos_({
    loanIds: ['L-0015', 'L-0019'],
    clienteEsperado: 'C-0016',
    dueDates: vencimientos,   // 10-nov, 10-dic, 10-ene-2027, 10-feb-2027
    rate: 0,                  // 0% — sólo reprogramación del saldo
    enviarCorreo: false,      // generar el enlace, NO enviar correo
    motivo: 'Consolidación de L-0015 y L-0019'
  });
  try {
    SpreadsheetApp.getUi().alert('Consolidación aplicada',
      out.loanId + ': ' + fmtMoney_(out.total) + ' en ' + out.cuotas + ' cuotas de ' +
      fmtMoney_(round2_(out.total / out.cuotas)) + '.\n\n' +
      'Préstamos cerrados: ' + out.origen.map(function (o) { return o.id + ' (' + fmtMoney_(o.saldo) + ')'; }).join(', ') + '.\n\n' +
      'Enlace de firma (también en la columna "Enlace de Firma"):\n' + (out.enlaceFirma || '(ver la fila del préstamo)'),
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* sin UI (ejecución desde el editor): el resultado va al toast y al valor de retorno */ }
  return out;
}

/**
 * Núcleo reutilizable. opts:
 *   loanIds          Array<string>  IDs de los préstamos a fusionar (≥ 2, mismo cliente).
 *   dueDates         Array<Date>    Fechas de vencimiento de las cuotas del préstamo nuevo.
 *   rate             number         Tasa del préstamo nuevo (0 = sólo reprogramar el saldo).
 *   clienteEsperado  string?        Si se pasa, valida que el cliente resuelto coincida.
 *   enviarCorreo     boolean?       true = enviar el correo de firma al prestatario.
 *   motivo           string?        Texto para las notas de auditoría.
 * Devuelve { loanId, row, total, cuotas, origen:[{id,saldo}], enlaceFirma }.
 */
function consolidarPrestamos_(opts) {
  return guard_('consolidarPrestamos_', function () {
    opts = opts || {};
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || (typeof setupCuotas_ === 'function' ? setupCuotas_(ss) : null);
    if (!bs) throw new Error('Falta la hoja "Prestatarios".');
    if (!ps) throw new Error('Falta la hoja "Pagos".');
    if (!cs) throw new Error('Falta la hoja "Cuotas".');

    const loanIds = (opts.loanIds || []).map(function (s) { return String(s).trim(); }).filter(Boolean);
    if (loanIds.length < 2) throw new Error('Pasá al menos 2 IDs de préstamo a consolidar.');
    const dueDates = (opts.dueDates || []).map(function (d) { return startOfDay_(d instanceof Date ? d : new Date(d)); });
    if (!dueDates.length) throw new Error('Pasá las fechas de vencimiento de las cuotas (opts.dueDates).');
    const hoy = startOfDay_(new Date());
    const motivo = String(opts.motivo || 'Consolidación').trim();

    // 1) Leer los préstamos a fusionar; saldo LIMPIO = Total a Pagar − pagos (capital + interés,
    //    SIN mora: la mora se condona al consolidar). Verificar que sean del MISMO cliente.
    const origen = [];
    let clienteId = '';
    let total = 0;
    loanIds.forEach(function (id) {
      const loan = findLoanById_(id);
      if (!loan) throw new Error('No se encontró el préstamo ' + id + ' en "Prestatarios".');
      const cid = String(loan.clientId || '').trim();
      if (!clienteId) clienteId = cid;
      else if (cid !== clienteId) throw new Error('Los préstamos no son del mismo cliente (' + id + ' → ' + cid + ' ≠ ' + clienteId + ').');
      const paid = round2_(loanPayments_(id).reduce(function (s, p) { return s + p.amount; }, 0));
      const owed = round2_(Math.max(0, loan.totalDue - paid));
      origen.push({ id: id, loan: loan, owed: owed });
      total = round2_(total + owed);
    });
    if (opts.clienteEsperado && clienteId !== String(opts.clienteEsperado).trim())
      throw new Error('El cliente resuelto (' + clienteId + ') no coincide con el esperado (' + opts.clienteEsperado + ').');
    if (total <= 0) throw new Error('El saldo combinado es 0: no hay nada que consolidar.');

    // 2) Borrar las cuotas de los préstamos de origen (cronograma viejo obsoleto; además, en estos
    //    préstamos las cuotas estaban DUPLICADas — bug N-1 — y mantendrían el estado en VENCIDO).
    origen.forEach(function (o) { deleteCuotasRowsForLoan_(cs, o.id); });
    SpreadsheetApp.flush();

    // 3) Crear el préstamo nuevo en la primera fila vacía (igual que el alta normal). Nombre/DNI
    //    (C:D) NO se tocan: son ARRAYFORMULA que se resuelven solos desde "Clientes" por ID Cliente.
    const loanId = nextLoanId_();
    const tRow = firstEmptyBorrowerRow_(bs);
    const H = headerIndex_(bs);
    const col = function (names, fb) { return colByAny_(H, names) || fb; };
    const idC  = col(['ID Préstamo', 'ID Prestamo'], PB.LOAN_ID);
    const cliC = col(['ID Cliente'], PB.CLIENT_ID);
    const capC = col(['Capital'], PB.PRINCIPAL);
    const plzC = col(['Plazo (días)', 'Plazo (dias)', 'Plazo (meses)', 'Plazo'], PB.TERM);
    const tasC = col(['Tasa'], PB.RATE);
    const fecC = col(['Fecha Préstamo', 'Fecha de Préstamo', 'Fecha del Préstamo', 'Fecha Prestamo'], PB.LOAN_DATE);
    const intC = col(['Interés', 'Interes'], PB.INTEREST);
    const totC = col(['Total a Pagar', 'Total a pagar'], PB.TOTAL);
    const sgnC = col(['Estado de Firma'], COL_SIGN_STATUS);
    const adjC = col(['Mora (ajuste)'], PB.MORA_ADJ);
    const bIdC = idC;

    const lastDue = dueDates[dueDates.length - 1];
    const spanDias = Math.max(1, Math.round((lastDue.getTime() - hoy.getTime()) / 86400000));

    bs.getRange(tRow, idC).setValue(loanId);
    bs.getRange(tRow, cliC).setValue(clienteId);
    bs.getRange(tRow, capC).setValue(total);
    bs.getRange(tRow, plzC).setValue(spanDias);
    bs.getRange(tRow, fecC).setValue(hoy).setNumberFormat('yyyy-mm-dd');

    // 4) Cronograma de cuotas del préstamo nuevo (fechas explícitas; la última absorbe el redondeo).
    cuotasForLoan_(cs, loanId, hoy, total, dueDates.length, null, { dates: dueDates });

    // 5) Fórmulas vivas de la fila (Saldo, Estado, Total Pagado, Mora auto, Saldo con Mora, Suma de
    //    cuotas, Vencimiento = última cuota, Enlace de Firma), resolviendo columnas por NOMBRE…
    writeBorrowerRowFormulas_(bs, tRow);
    SpreadsheetApp.flush();
    // …y FORZAR 0% (sólo reprogramación): Tasa 0, Interés 0, Total a Pagar = saldo combinado.
    //    (writeBorrowerRowFormulas_ deriva la Tasa del plazo; para un plazo de 4 meses daría "".)
    //    Si alguna vez se reconstruyen las fórmulas (reactivarFuncionalidad), REAFIRMAR estos 3.
    bs.getRange(tRow, tasC).setValue(Number(opts.rate) || 0);
    bs.getRange(tRow, intC).setValue(0);
    bs.getRange(tRow, totC).setValue(total);
    // Mora (ajuste) VACÍA en el préstamo nuevo → recargo por mora AUTOMÁTICO (5% diario estándar).
    bs.getRange(tRow, adjC).clearContent();

    // 6) Firma pendiente + token + enlace. El "Enlace de Firma" de la fila es fórmula (lo arma
    //    signingLinkFormula_ con el token); acá devolvemos también el enlace directo para la mano.
    bs.getRange(tRow, sgnC).setValue(SIGN.PENDING);
    const token = ensureLoanToken_(loanId, true);
    const enlaceFirma = signingLink_(loanId, token);

    bs.getRange(tRow, bIdC).setNote(
      'CONSOLIDACIÓN (' + motivo + ') el ' + fmtDate_(hoy) + '.\n' +
      'Fusiona: ' + origen.map(function (o) { return o.id + ' (' + fmtMoney_(o.saldo) + ')'; }).join(' + ') + '.\n' +
      'Saldo refinanciado: ' + fmtMoney_(total) + ' al 0% en ' + dueDates.length + ' cuotas de ' +
      fmtMoney_(round2_(total / dueDates.length)) + ', con vencimientos ' +
      dueDates.map(function (d) { return fmtDate_(d); }).join(', ') + '.\n' +
      'Recargo por mora estándar (5% diario) si se atrasa.');

    // 7) Contrato (copia de revisión) + marca de enviado. Correo de firma SÓLO si opts.enviarCorreo.
    let loanNuevo = null;
    try {
      loanNuevo = readLoan_(bs, tRow);
      const file = makeAgreementFile_(loanNuevo);
      markAgreementSent_(loanId, file);
      if (opts.enviarCorreo && loanNuevo.email) emailSigningRequest_(loanNuevo, file);
    } catch (e) { logError_('consolidarPrestamos_:contrato', e); }

    // 8) CERRAR los préstamos de origen: pago de TRASLADO por su saldo (→ Saldo 0) y condonar su
    //    mora (ajuste = 0 → Saldo con Mora 0). Con las cuotas ya borradas (paso 2), el Estado
    //    resuelve a SALDADO/PAGADO. El pago de traslado NO es cobro real (nota aclaratoria).
    const PH = headerIndex_(ps);
    const pcol = function (names, fb) { return colByAny_(PH, names) || fb; };
    const pIdC   = pcol(['ID Pago'], PP.PAY_ID);
    const pLoanC = pcol(['ID Préstamo', 'ID Prestamo'], PP.LOAN_ID);
    const pDateC = pcol(['Fecha de Pago'], PP.DATE);
    const pAmtC  = pcol(['Monto Pagado'], PP.AMOUNT);
    const pBalC  = pcol(['Saldo Posterior'], PP.BALANCE);
    const pNameC = pcol(['Nombre'], PP.NAME);
    const pDniC  = pcol(['DNI'], PP.DNI);

    // Mapa ID préstamo → fila en "Prestatarios" (una sola lectura).
    const idsCol = bs.getRange(2, bIdC, Math.max(0, bs.getLastRow() - 1), 1).getValues();
    const rowOf = {};
    for (let i = 0; i < idsCol.length; i++) { const v = String(idsCol[i][0]).trim(); if (v) rowOf[v] = i + 2; }

    let seq = nextPaymentSeq_(ps);
    origen.forEach(function (o) {
      if (o.owed > 0) {
        const pr = ps.getLastRow() + 1;
        const payId = 'P-' + String(seq++).padStart(4, '0');
        ps.getRange(pr, pIdC).setValue(payId);
        ps.getRange(pr, pLoanC).setValue(o.id);
        ps.getRange(pr, pDateC).setValue(hoy).setNumberFormat('yyyy-mm-dd');
        ps.getRange(pr, pAmtC).setValue(o.owed).setNumberFormat(CFG.CURRENCY_FMT);
        if (pNameC) ps.getRange(pr, pNameC).setValue(o.loan.name || '');
        if (pDniC)  ps.getRange(pr, pDniC).setValue(o.loan.dni || '');
        if (pBalC && typeof paymentBalanceFormula_ === 'function')
          ps.getRange(pr, pBalC).setFormula(paymentBalanceFormula_(pr)).setNumberFormat(CFG.CURRENCY_FMT);
        ps.getRange(pr, pIdC).setNote('Traslado a ' + loanId + ' (' + motivo + ') el ' + fmtDate_(hoy) +
          '. NO es un cobro real: el saldo se refinancia en ' + loanId + '.');
      }
      const or = rowOf[o.id];
      if (or) {
        bs.getRange(or, adjC).setValue(0); // condonar mora (0 = eximir) → Estado SALDADO/PAGADO
        bs.getRange(or, bIdC).setNote('CONSOLIDADO en ' + loanId + ' el ' + fmtDate_(hoy) + '.\n' +
          'Saldo trasladado: ' + fmtMoney_(o.owed) + '. Mora condonada (ajuste = 0). Cuotas dadas de baja.\n' +
          'Este préstamo queda SALDADO; el saldo continúa en ' + loanId + '.');
      }
    });
    SpreadsheetApp.flush();

    const resumen = loanId + ': ' + fmtMoney_(total) + ' en ' + dueDates.length + ' cuotas de ' +
      fmtMoney_(round2_(total / dueDates.length)) + ' (' + loanIds.join(' + ') + ' consolidados y SALDADOS).';
    try { ss.toast(resumen, '🔗 Consolidación', 10); } catch (e) {}
    return {
      loanId: loanId, row: tRow, total: total, cuotas: dueDates.length,
      origen: origen.map(function (o) { return { id: o.id, saldo: o.owed }; }),
      enlaceFirma: enlaceFirma
    };
  });
}
