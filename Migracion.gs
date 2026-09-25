/************************************************************
 * MIGRACION v1 — Trae los DATOS DE PRODUCCIÓN al sistema nuevo
 * ----------------------------------------------------------
 * Copia los datos de la planilla de producción
 * ("impulso-datos-corregidos") a ESTA planilla (la del sistema
 * nuevo), REEMPLAZANDO los datos de prueba. Diferencias
 * verificadas antes de escribir este script:
 *   • Todo lo que está SOLO en esta planilla es dato de PRUEBA
 *     (Denis Bell 12345678, Xoana 33570168, "Impulso Credito"
 *     12345677): préstamos L-0034…L-0039, pagos P-0014…P-0021,
 *     clientes C-0026…C-0028, fila L-0038 de Firmas y 3
 *     solicitudes de prueba en "Nuevos Prestatarios".
 *   • En producción esos MISMOS IDs son datos reales (L-0034
 *     Joana Troncoso, L-0035 Cristina Mabel avila, L-0036 /
 *     C-0026 Verónica Cardozo; pagos reales a L-0004/L-0022/
 *     L-0012/L-0016; firmas de L-0035/L-0036).
 * Por eso la estrategia es REEMPLAZO TOTAL de las hojas de
 * datos (no merge): Clientes, Prestatarios, Pagos y Firmas se
 * vacían y se recargan desde producción; "Nuevos Prestatarios"
 * solo se vacía (producción no tiene solicitudes pendientes).
 * La copia es POSICIONAL (col A en adelante): "Clientes" tiene
 * un encabezado 'Ref 2 Teléfono' DUPLICADO en ambas planillas,
 * así que mapear por nombre no es seguro; los layouts son
 * idénticos en el prefijo copiado (se verifica antes de escribir).
 *
 * USO (una sola vez; es re-ejecutable — cada paso recarga todo)
 *   1) Pegue este archivo en el proyecto de Apps Script de ESTA
 *      planilla (la nueva) junto a LoanManagerV2/Validaciones/
 *      Reactivar. Guarde.
 *   2) Ejecute  ->  migrarDesdeProduccion  (autorice Drive+Sheets).
 *      Antes de tocar nada crea una COPIA DE RESPALDO del archivo.
 *   3) Revise el resumen final y la hoja "Errores".
 ************************************************************/

// Planilla de ORIGEN (producción): impulso-datos-corregidos.
const MIGRACION_SOURCE_ID = '1PWHRO56lhB5TPPiD--NoQQmi6oQznah6Gz0tUuKkP1c';

// Hojas de datos a copiar y el ANCHO del bloque de origen (columnas desde A).
// El ancho es el del layout de producción; las columnas extra del destino
// (mora X–Z de Prestatarios, Bloqueado/Motivo de Clientes, "Generar recibo"
// de Pagos) quedan vacías y las reinstala reactivarFuncionalidad().
const MIGRACION_HOJAS = [
  { name: 'Clientes', width: 21, idCol: 1 },
  { name: 'Prestatarios', width: 23, idCol: 1 },
  { name: 'Pagos', width: 9, idCol: 1 },
  { name: 'Firmas', width: 14, idCol: 2 }, // ID Préstamo en col B; col A = fecha
];

/** Punto de entrada: migra producción → esta planilla y reconstruye todo. */
function migrarDesdeProduccion() {
  return guard_('migrarDesdeProduccion', function () {
    const ss = getSS_(), ui = SpreadsheetApp.getUi();
    const report = [];
    const run = (label, fn) => { try { report.push(label + ': ' + fn()); } catch (e) { logError_('migracion:' + label, e); report.push(label + ': ⚠ ' + e.message); throw e; } };

    // 0) Confirmación explícita: esto BORRA los datos actuales de esta planilla.
    const resp = ui.alert('Migrar datos de producción',
      'Se REEMPLAZAN los datos de Clientes, Prestatarios, Pagos y Firmas de ESTA planilla ' +
      'con los de producción (impulso-datos-corregidos), y se vacían las solicitudes de prueba. ' +
      'Primero se crea una copia de respaldo del archivo completo.\n\n¿Continuar?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return 'Cancelado por el usuario.';

    // 1) Respaldo del archivo completo ANTES de tocar nada.
    let backupName = '';
    run('Respaldo', () => {
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
      backupName = 'Impulso — respaldo pre-migración ' + stamp;
      DriveApp.getFileById(ss.getId()).makeCopy(backupName);
      return backupName;
    });

    // 2) Abrir producción.
    const src = SpreadsheetApp.openById(MIGRACION_SOURCE_ID);

    // 3) Copiar las hojas de datos (posicional, solo valores).
    MIGRACION_HOJAS.forEach(h => run(h.name, () => migrarHoja_(src, ss, h)));

    // 3b) "Nuevos Prestatarios": producción no tiene solicitudes pendientes →
    //     solo se vacían las filas de prueba del destino (encabezados intactos).
    run('Nuevos Prestatarios', () => {
      const nb = ss.getSheetByName(CFG.SHEETS.NEW);
      if (!nb || nb.getLastRow() < 2) return 'sin filas que limpiar';
      nb.getRange(2, 1, nb.getLastRow() - 1, nb.getLastColumn()).clearContent();
      return 'solicitudes de prueba eliminadas';
    });

    // 4) Configuración: los valores de producción pisan las claves existentes;
    //    las claves nuevas del destino (tope de mora, límites, etc.) se conservan.
    run('Configuración', () => migrarConfiguracion_(src, ss));

    // 5) Cuotas: regenerar el cronograma desde los préstamos migrados.
    run('Cuotas', () => migrarCuotas_(ss));

    // 6) Reconstrucción: cachés, fórmulas, validaciones y vistas derivadas.
    run('Reconstrucción', () => {
      if (typeof invalidateHeaderIndexCache_ === 'function') invalidateHeaderIndexCache_();
      if (typeof invalidateClientesRows_ === 'function') invalidateClientesRows_();
      if (typeof invalidateClientesCache_ === 'function') invalidateClientesCache_();
      if (typeof reactivarFuncionalidad === 'function') reactivarFuncionalidad();
      rebuildLateSheet_(ss);
      rebuildRemindersSheet_(ss);
      if (typeof refreshAll === 'function') refreshAll(true);
      return 'fórmulas, validaciones y vistas reconstruidas';
    });

    SpreadsheetApp.flush();
    const msg = report.join('\n');
    try {
      ui.alert('Migración completada',
        msg + '\n\nRespaldo: "' + backupName + '" (en su Drive).\n' +
        'Verifique el Panel, "Pagos Atrasados" y ejecute la suite 🧪 si desea.', ui.ButtonSet.OK);
    } catch (e) { }
    return msg;
  });
}

/**
 * Copia una hoja de datos de producción al destino: valida encabezados
 * (posicional, claves normalizadas con hkey_), vacía TODAS las filas de datos
 * del destino (a lo ancho completo: borra también columnas de prueba/mora) y
 * escribe el bloque de origen. V-22: la cantidad de filas con ID del destino
 * debe igualar a la del origen, si no se aborta.
 */
function migrarHoja_(src, ss, h) {
  const sSh = src.getSheetByName(h.name), tSh = ss.getSheetByName(h.name);
  if (!sSh) throw new Error('Falta la hoja "' + h.name + '" en producción.');
  if (!tSh) throw new Error('Falta la hoja "' + h.name + '" en esta planilla.');

  // Encabezados: el prefijo copiado debe coincidir posición por posición.
  const sHead = sSh.getRange(1, 1, 1, h.width).getValues()[0];
  const tHead = tSh.getRange(1, 1, 1, h.width).getValues()[0];
  for (let i = 0; i < h.width; i++) {
    if (hkey_(sHead[i]) !== hkey_(tHead[i]))
      throw new Error(h.name + ': el encabezado de la columna ' + (i + 1) + ' difiere ("' + sHead[i] + '" vs "' + tHead[i] + '"). Migración abortada.');
  }

  // Leer origen (solo valores) y contar filas con ID.
  const sLast = sSh.getLastRow();
  const data = sLast >= 2 ? sSh.getRange(2, 1, sLast - 1, h.width).getValues() : [];
  const conId = rows => rows.reduce((n, r) => n + (String(r[h.idCol - 1] || '').trim() ? 1 : 0), 0);
  const srcCount = conId(data);

  // Vaciar TODAS las filas de datos del destino, a lo ancho completo (borra
  // también fórmulas de mora, casillas y las columnas que producción no tiene).
  const tRows = Math.max(tSh.getMaxRows() - 1, 1);
  tSh.getRange(2, 1, tRows, tSh.getMaxColumns()).clearContent().clearDataValidations();

  // Escribir el bloque de producción.
  if (data.length) tSh.getRange(2, 1, data.length, h.width).setValues(data);

  // V-22 — comparar filas con ID origen vs destino.
  const tLast = tSh.getLastRow();
  const written = tLast >= 2 ? conId(tSh.getRange(2, 1, tLast - 1, h.width).getValues()) : 0;
  if (written !== srcCount)
    throw new Error('V-22 — ' + h.name + ': origen ' + srcCount + ' fila(s) con ID, destino ' + written + '. Migración abortada.');
  return srcCount + ' fila(s) migradas (V-22 OK)';
}

/**
 * Copia los VALORES de "Configuración" de producción sobre las claves del
 * destino (coincidencia por hkey_); las claves que falten se agregan al final.
 * Las claves que solo existen en el destino no se tocan.
 */
function migrarConfiguracion_(src, ss) {
  const sSh = src.getSheetByName(CFG.SHEETS.SETTINGS), tSh = ss.getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sSh || !tSh) return 'hoja "Configuración" faltante — sin cambios';
  const last = sSh.getLastRow();
  if (last < 2) return 'producción sin ajustes — sin cambios';
  const rows = sSh.getRange(2, 1, last - 1, 2).getValues();
  let updated = 0, added = 0;
  rows.forEach(r => {
    const key = String(r[0] || '').trim();
    if (!key) return;
    if (setSettingValue_(tSh, key, r[1])) updated++;
    else { tSh.appendRow([key, r[1]]); added++; }
  });
  return updated + ' ajuste(s) actualizados' + (added ? ', ' + added + ' agregados' : '');
}

/**
 * Regenera "Cuotas" desde los préstamos migrados: se vacía la hoja (las cuotas
 * previas referían préstamos de prueba) y se crea el cronograma de cada
 * préstamo NO saldado — 3 cuotas mensuales (30/60/90) para plazo 90, si no
 * 1 cuota a su plazo. (No se usa backfillCuotas_: ese siempre escribe 1 cuota.)
 */
function migrarCuotas_(ss) {
  const cs = ss.getSheetByName(CFG.SHEETS.INSTALLMENTS) || setupCuotas_(ss);
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (!bs || bs.getLastRow() < 2) return 'sin préstamos';
  if (cs.getLastRow() >= 2) cs.getRange(2, 1, cs.getLastRow() - 1, cs.getMaxColumns()).clearContent();
  const rows = bs.getRange(2, 1, bs.getLastRow() - 1, PB.STATE).getValues();
  let n = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i], loanId = String(r[PB.LOAN_ID - 1]).trim();
    if (!/^L-/i.test(loanId)) continue;
    const estado = String(r[PB.STATE - 1]).trim().toUpperCase();
    if (estado === ST.PAID || estado === ST.CLEARED) continue; // saldados: sin cronograma
    const fecha = r[PB.LOAN_DATE - 1], term = Number(r[PB.TERM - 1]) || 0, total = Number(r[PB.TOTAL - 1]) || 0;
    if (!(fecha instanceof Date) || total <= 0) continue;
    const cuotas = installmentCount_({ term: term });
    cuotasForLoan_(cs, loanId, fecha, round2_(total), cuotas, cuotas === 3 ? [30, 60, 90] : [termDays_(term)]);
    n++;
  }
  return n + ' cronograma(s) regenerados';
}
