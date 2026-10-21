/*************************************************************
 * VerificacionDNI.gs — Verificación de DNI por OCR (menú + disparador diario)
 *
 * Qué hace:
 *   Una vez por día (disparador verificarDNIDiario) recorre una carpeta de Drive
 *   con fotos de DNI subidas (frente/"frente" y dorso/"dorso"), les pasa OCR,
 *   vuelca cada atributo leído a la hoja "Verificación de DNI", lo coteja contra
 *   la hoja "Clientes" y marca anomalías (DNI que no coincide, nombre distinto,
 *   DNI ausente en Clientes, ilegible, frente/dorso discrepantes, etc.). Al final
 *   envía un informe por correo al "Correo para resumen de avisos".
 *
 * IMPORTANTE — la lectura por OCR es ORIENTATIVA: puede contener errores y NO
 * reemplaza la verificación humana. El estado "Revisar" es un aviso para mirar,
 * nunca un rechazo/aprobación automática.
 *
 * Comparte el espacio global con LoanManagerV2.gs / Validaciones.gs / Reactivar.gs:
 * reutiliza getSS_, CFG, getSetting_/setSetting_, guard_, logError_, getOrCreate_,
 * cc_, esc_, fmtDate_, sendBrandedEmail_, extractDriveIds_, normDni_, normName_,
 * clientesRows_, clienteByDni_, y el seam de prueba _ocrStub/TEST_MODE.
 *
 * REQUISITO DE INSTALACIÓN (una sola vez):
 *   En el editor de Apps Script ▸ Servicios ▸ + ▸ "Drive API" (identificador
 *   "Drive", versión v2). Lo usa ocrImageText_ para OCR vía Documentos de Google.
 *   La primera corrida pedirá autorización de Drive y Documentos.
 *************************************************************/

/** Encabezados de la hoja "Verificación de DNI". La última columna (ID de archivo)
 *  es la clave de de-duplicación para el escaneo incremental. */
var DNI_CHECK_HEADERS = ['Fecha de escaneo', 'Persona (del archivo)', 'Lado', 'Apellido', 'Nombre',
  'N° de DNI', 'Sexo', 'Fecha de nacimiento', 'Nacionalidad', 'CUIL', 'Fecha de vencimiento',
  'DNI en Clientes', 'Nombre en Clientes', 'Coincidencia', 'Estado', 'Anomalías',
  'Archivo frente', 'Archivo dorso', 'ID de archivo (Drive)'];
// Índices 1-based que se usan para formato y lectura:
var DNK = { DATE: 1, VIA: 14, STATE: 15, FILEIDS: 19 };

/* ===================== DISPARADOR + ORQUESTACIÓN ===================== */

/**
 * Instala/actualiza el disparador diario (verificarDNIDiario). Siembra ajustes
 * por defecto (solo si están vacíos) y evita duplicar el disparador. Correr una
 * vez desde el menú o el editor (pide autorización de Google).
 */
function instalarVerificacionDNI() {
  return guard_('instalarVerificacionDNI', function () {
    const defs = [
      ['Verificación de DNI activa (SÍ/NO)', 'SÍ'],
      ['Carpeta de documentos de DNI (ID o enlace)', '10lrgWdEkQdykjocMpRR-DlPeRQjg_a43'],
      ['Incluir subcarpetas de DNI (SÍ/NO)', 'SÍ'],
      ['Palabras clave de archivos de DNI', 'DNI,CUIL'],
      ['Hora de escaneo de DNI', '6'],
      ['Máx. archivos por corrida', '40'],
      ['Idioma OCR', 'es'],
    ];
    defs.forEach(function (d) { if (!String(getSetting_(d[0]) || '').trim()) setSetting_(d[0], d[1]); });
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'verificarDNIDiario') ScriptApp.deleteTrigger(t); });
    let hora = parseInt(String(getSetting_('Hora de escaneo de DNI') || '6'), 10);
    if (!isFinite(hora) || hora < 0 || hora > 23) hora = 6;
    ScriptApp.newTrigger('verificarDNIDiario').timeBased().everyDays(1).atHour(hora).create();
    const msg = 'Verificación de DNI: escaneo diario instalado (~' + hora + ':00 h). ' +
      'Recordá habilitar el servicio avanzado "Drive API" (Editor ▸ Servicios ▸ +) para el OCR. ' +
      'El informe se envía al "Correo para resumen de avisos".';
    try { getSS_().toast(msg, '🪪 Verificación de DNI instalada', 8); } catch (e) { }
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
    return msg;
  });
}

/** Handler del disparador diario: escaneo incremental + informe por correo. */
function verificarDNIDiario() {
  return guard_('verificarDNIDiario', function () {
    if (!verificacionDNIActiva_()) return 'Verificación de DNI desactivada (Configuración).';
    const stats = scanDNIFolderCore_({ force: false });
    enviarReporteDNI_(stats);
    return stats.mensaje;
  });
}

/** Menú: escanear ahora (solo archivos nuevos) + informe. */
function escanearDNIAhora() {
  return guard_('escanearDNIAhora', function () {
    const stats = scanDNIFolderCore_({ force: false });
    enviarReporteDNI_(stats);
    try { getSS_().toast(stats.mensaje, '🪪 Verificación de DNI', 8); } catch (e) { }
    return stats.mensaje;
  });
}

/** Menú: re-escanear TODO (ignora los ya procesados), con confirmación. */
function reescanearDNITodo() {
  return guard_('reescanearDNITodo', function () {
    let ui; try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
    if (ui) {
      const r = ui.alert('Re-escanear TODO',
        'Esto vuelve a procesar TODOS los archivos de la carpeta, incluso los ya escaneados (consume cuota de OCR). ¿Continuar?',
        ui.ButtonSet.YES_NO);
      if (r !== ui.Button.YES) return 'Cancelado.';
    }
    const stats = scanDNIFolderCore_({ force: true });
    enviarReporteDNI_(stats);
    try { getSS_().toast(stats.mensaje, '🪪 Verificación de DNI', 8); } catch (e) { }
    return stats.mensaje;
  });
}

/** ¿Está activa la verificación de DNI? (Configuración) */
function verificacionDNIActiva_() {
  const v = String(getSetting_('Verificación de DNI activa (SÍ/NO)') || '').trim().toUpperCase();
  return v === 'SÍ' || v === 'SI';
}

/* ===================== NÚCLEO DEL ESCANEO ===================== */

/**
 * Recorre la carpeta, OCR + cotejo, escribe filas nuevas y devuelve un objeto de
 * estadísticas con la forma que esperan los constructores del informe.
 * @param {{force?:boolean}} opts  force=true ignora la de-duplicación por ID.
 */
function scanDNIFolderCore_(opts) {
  opts = opts || {};
  const force = !!opts.force;
  const ss = getSS_();
  const sh = setupVerificacionDNI_(ss);
  const hoy = new Date();

  const stats = {
    escaneados: 0, nuevos: 0, ok: 0, revisar: 0, sinCoincidencia: 0, errores: 0,
    filasNuevas: [], anomaliasAbiertas: [], hoy: hoy, mensaje: '',
  };

  let folder;
  try { folder = dniScanFolder_(); }
  catch (e) {
    logError_('scanDNIFolderCore_:folder', e);
    stats.errores++;
    stats.mensaje = 'No se pudo abrir la carpeta de documentos de DNI: ' + ((e && e.message) || e);
    return stats;
  }

  const recurse = /^(SÍ|SI)$/i.test(String(getSetting_('Incluir subcarpetas de DNI (SÍ/NO)') || 'SÍ').trim());
  const keywords = String(getSetting_('Palabras clave de archivos de DNI') || 'DNI,CUIL')
    .split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
  let maxFiles = parseInt(String(getSetting_('Máx. archivos por corrida') || '40'), 10);
  if (!isFinite(maxFiles) || maxFiles <= 0) maxFiles = 40;

  // Re-escaneo total: vaciar las filas de datos para reconstruir sin duplicar.
  if (force && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, DNI_CHECK_HEADERS.length).clearContent();
    SpreadsheetApp.flush();
  }

  // Mapas sobre las filas existentes:
  //  - 'procesados': IDs ya OCR'd (de-duplicación del escaneo incremental).
  //  - 'idToRow': ID → fila (superseder al llegar el otro lado o una re-subida).
  //  - 'keyToRows': persona normalizada → filas (dejar UNA fila por persona: la más reciente).
  const procesados = {};
  const idToRow = {};
  const keyToRows = {};
  if (!force && sh.getLastRow() >= 2) {
    const n = sh.getLastRow() - 1;
    const idVals = sh.getRange(2, DNK.FILEIDS, n, 1).getValues();
    const perVals = sh.getRange(2, 2, n, 1).getValues();   // "Persona (del archivo)"
    idVals.forEach(function (r, i) {
      String(r[0] || '').split(',').forEach(function (id) {
        id = id.trim(); if (id) { procesados[id] = true; idToRow[id] = i + 2; }
      });
      const pk = normName_(perVals[i][0]);
      if (pk) (keyToRows[pk] = keyToRows[pk] || []).push(i + 2);
    });
  }

  const files = collectDniFiles_(folder, recurse, [], keywords);
  const groups = pairDniFiles_(files);
  const nuevasFilas = [];
  const filasASuperseder = {};   // nº de fila previa → true (partial→completo / re-subidas)
  let enCorrida = 0;
  let topeAlcanzado = false;

  for (let gi = 0; gi < groups.length; gi++) {
    const g = groups[gi];
    const ids = [];
    if (g.frente) ids.push(g.frente.id);
    if (g.dorso) ids.push(g.dorso.id);
    if (!ids.length) continue;

    const pendiente = force || ids.some(function (id) { return !procesados[id]; });
    if (!pendiente) continue;
    if (enCorrida >= maxFiles) { topeAlcanzado = true; break; }  // tope (límite de 6 min por corrida)
    enCorrida += ids.length;   // contar lados: cada lado es una llamada OCR

    let textoFrente = '', textoDorso = '';
    try { if (g.frente) textoFrente = ocrImageText_(g.frente.file.getBlob()) || ''; }
    catch (e) { logError_('scanDNIFolderCore_:ocrFrente:' + (g.frente && g.frente.name), e); stats.errores++; }
    try { if (g.dorso) textoDorso = ocrImageText_(g.dorso.file.getBlob()) || ''; }
    catch (e) { logError_('scanDNIFolderCore_:ocrDorso:' + (g.dorso && g.dorso.name), e); stats.errores++; }

    const front = parseDniFront_(textoFrente);
    const back = parseDniBack_(textoDorso);
    const fields = mergeDniFields_(front, back);
    const match = matchClienteForDni_(fields, g.persona);
    const an = dniAnomalies_(fields, match, g);
    const cli = match.cliente;
    const lado = (g.frente && g.dorso) ? 'Frente+Dorso' : (g.frente ? 'Frente' : 'Dorso');

    nuevasFilas.push([
      hoy, g.persona || '', lado,
      fields.apellido || '', fields.nombre || '', fields.dni || '', fields.sexo || '',
      fields.nacimiento || '', fields.nacionalidad || '', fields.cuil || '', fields.vencimiento || '',
      cli ? (cli.dni || '') : '', cli ? (cli.name || '') : '', match.via || 'Sin coincidencia',
      an.estado, an.labels.join('; '),
      g.frente ? g.frente.url : '', g.dorso ? g.dorso.url : '', ids.join(','),
    ]);
    // Marcar filas previas obsoletas: por ID compartido y por misma persona
    // (deja una sola fila por persona, la del DNI más reciente).
    ids.forEach(function (id) { if (idToRow[id]) filasASuperseder[idToRow[id]] = true; });
    const gk = normName_(g.persona);
    if (gk && keyToRows[gk]) keyToRows[gk].forEach(function (rn) { filasASuperseder[rn] = true; });

    stats.escaneados++;
    if (an.estado === 'OK') stats.ok++; else stats.revisar++;
    if (match.via === 'Sin coincidencia') stats.sinCoincidencia++;
    stats.filasNuevas.push({
      persona: g.persona || '',
      apenom: ((fields.apellido || '') + ' ' + (fields.nombre || '')).trim(),
      dni: fields.dni || '', via: match.via || 'Sin coincidencia',
      estado: an.estado, anomalias: an.labels.join('; '),
    });
  }

  if (nuevasFilas.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, nuevasFilas.length, DNI_CHECK_HEADERS.length).setValues(nuevasFilas);
    stats.nuevos = nuevasFilas.length;
  }

  // Borrar las filas previas superseded (de mayor a menor nº para no desplazar índices).
  Object.keys(filasASuperseder).map(Number).sort(function (a, b) { return b - a; })
    .forEach(function (rn) { sh.deleteRow(rn); });

  // Anomalías abiertas: TODAS las filas con Estado = "Revisar" (incluye corridas previas).
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, DNI_CHECK_HEADERS.length).getValues().forEach(function (r) {
      if (String(r[DNK.STATE - 1] || '').trim() === 'Revisar') {
        stats.anomaliasAbiertas.push({
          persona: r[1], apenom: ((r[3] || '') + ' ' + (r[4] || '')).trim(),
          dni: r[5], via: r[DNK.VIA - 1], anomalias: r[15],
        });
      }
    });
  }

  stats.mensaje = 'Verificación de DNI: ' + stats.nuevos + ' nuevos escaneados, ' +
    stats.revisar + ' a revisar (de los nuevos), ' + stats.anomaliasAbiertas.length + ' anomalías abiertas, ' +
    stats.errores + ' errores.' + (topeAlcanzado ? ' Se alcanzó el tope por corrida; el resto continúa mañana.' : '');
  stats.topeAlcanzado = topeAlcanzado;
  return stats;
}

/** Crea/formatea la hoja "Verificación de DNI" (idempotente; no borra datos). */
function setupVerificacionDNI_(ss) {
  ss = ss || getSS_();
  const sh = getOrCreate_(ss, CFG.SHEETS.DNI_CHECK);
  const head = sh.getRange(1, 1, 1, DNI_CHECK_HEADERS.length).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head.join('|') !== DNI_CHECK_HEADERS.join('|')) {
    sh.getRange(1, 1, 1, DNI_CHECK_HEADERS.length).setValues([DNI_CHECK_HEADERS])
      .setFontWeight('bold').setBackground('#0b8043').setFontColor('#fff').setWrap(true);
    [140, 180, 95, 150, 150, 110, 55, 130, 120, 120, 130, 120, 180, 150, 85, 280, 150, 150, 220]
      .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(2, DNK.DATE, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  sh.setFrozenRows(1);
  const est = sh.getRange(2, DNK.STATE, CFG.MAX_ROWS, 1);
  sh.setConditionalFormatRules([cc_(est, 'OK', '#d9ead3'), cc_(est, 'Revisar', '#f4cccc')]);
  sh.getRange(1, DNK.STATE).setNote('La lectura por OCR es orientativa y puede contener errores; ' +
    'revisá manualmente los casos marcados «Revisar». No reemplaza la verificación humana.');
  return sh;
}

/** Resuelve la carpeta de Drive a escanear desde Configuración (acepta ID pelado o enlace). */
function dniScanFolder_() {
  const ref = String(getSetting_('Carpeta de documentos de DNI (ID o enlace)') || '').trim();
  if (!ref) throw new Error('Configurá "Carpeta de documentos de DNI (ID o enlace)" en la hoja Configuración.');
  const ids = extractDriveIds_(ref);
  const id = ids.length ? ids[0] : ref;
  return DriveApp.getFolderById(id);
}

/**
 * Enumera archivos de imagen/PDF de la carpeta (y subcarpetas si recurse). Si
 * 'keywords' tiene elementos, SOLO incluye archivos cuyo nombre contenga alguna
 * de esas palabras (p. ej. "DNI"/"CUIL"): así se omiten contratos, recibos, etc.
 * que conviven en la misma carpeta del prestatario.
 */
function collectDniFiles_(folder, recurse, acc, keywords) {
  acc = acc || [];
  const kw = (keywords && keywords.length) ? keywords : null;
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const mime = f.getMimeType() || '';
    if (mime.indexOf('image/') !== 0 && mime !== 'application/pdf') continue; // ignorar Docs nativos
    const nm = (f.getName() || '').toLowerCase();
    if (kw && !kw.some(function (k) { return nm.indexOf(k) >= 0; })) continue;  // filtrar por nombre
    let t = 0; try { t = f.getLastUpdated().getTime(); } catch (e) { t = 0; }   // fecha real para "más reciente"
    acc.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), mime: mime, time: t, file: f });
  }
  if (recurse) {
    const fit = folder.getFolders();
    while (fit.hasNext()) collectDniFiles_(fit.next(), true, acc, keywords);
  }
  return acc;
}

/* ===================== APAREAMIENTO FRENTE/DORSO ===================== */

/**
 * Descompone un nombre de archivo en {persona, key, side, stamp}.
 *  - persona: etiqueta legible (lo que queda tras quitar "DNI", lado y timestamp).
 *  - key: persona normalizada (clave de agrupación frente+dorso).
 *  - side: 'dorso' si hay token de dorso/reverso; si no, 'frente'.
 *  - stamp: dígitos del timestamp tipo 20261021-112335 (para desempatar duplicados).
 */
function parseDniFileName_(name) {
  let base = String(name || '').replace(/\.[a-z0-9]{2,5}$/i, ''); // sin extensión
  const side = /\b(dorso|reverso|back|trasero|posterior)\b/i.test(base) ? 'dorso' : 'frente';
  const mStamp = base.match(/(\d{8}[-_ ]\d{6})/);
  const stamp = mStamp ? mStamp[1].replace(/\D/g, '') : '';
  const persona = base
    .replace(/(\d{8}[-_ ]\d{6})/g, ' ')
    .replace(/\b(dorso|reverso|back|trasero|posterior|frente|front|anverso|delantero)\b/gi, ' ')
    .replace(/\bDNI\b/gi, ' ')
    .replace(/[_\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return { persona: persona, key: normName_(persona), side: side, stamp: stamp };
}

/** Agrupa archivos por persona → [{persona, frente, dorso}]. Por cada lado conserva
 *  SOLO el más reciente (por fecha real del archivo; el timestamp del nombre desempata),
 *  de modo que las subidas viejas de DNI quedan descartadas. */
function pairDniFiles_(files) {
  const groups = {};
  files.forEach(function (f) {
    const meta = parseDniFileName_(f.name);
    const key = meta.key || normName_(f.name);
    const g = groups[key] || (groups[key] = { persona: meta.persona || f.name, frente: null, dorso: null });
    if (!g.persona && meta.persona) g.persona = meta.persona;
    const slot = meta.side; // 'frente' | 'dorso'
    const cur = g[slot];
    const ft = f.time || 0;
    const newer = !cur || ft > (cur._time || 0) ||
      (ft === (cur._time || 0) && String(meta.stamp) > String(cur._stamp || ''));
    if (newer) {
      g[slot] = { id: f.id, name: f.name, url: f.url, mime: f.mime, file: f.file, _stamp: meta.stamp, _time: ft };
    }
  });
  return Object.keys(groups).map(function (k) { return groups[k]; });
}

/* ===================== OCR (seam único e intercambiable) ===================== */

/**
 * Devuelve el texto OCR de una imagen. Implementación por defecto: servicio
 * avanzado Drive API (v2) — inserta un Documento de Google temporal con OCR,
 * lee su texto y borra el temporal.
 *
 * ALTERNATIVA (no implementada): Cloud Vision DOCUMENT_TEXT_DETECTION vía
 * UrlFetchApp POST a https://vision.googleapis.com/v1/images:annotate con una
 * API key (más preciso en la MRZ del dorso; requiere proyecto GCP con facturación
 * y la Vision API habilitada). Al estar aislado acá, cambiar de motor es tocar
 * solo esta función.
 */
function ocrImageText_(blob) {
  if (TEST_MODE) return (typeof _ocrStub === 'function') ? _ocrStub(blob) : String(_ocrStub || '');
  const lang = String(getSetting_('Idioma OCR') || 'es').trim() || 'es';
  const tmp = Drive.Files.insert(
    { title: 'ocr-tmp-' + (new Date()).getTime(), mimeType: 'application/vnd.google-apps.document' },
    blob, { ocr: true, ocrLanguage: lang });
  try {
    return DocumentApp.openById(tmp.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(tmp.id); } catch (e) { logError_('ocrImageText_:cleanup', e); }
  }
}

/* ===================== PARSERS ===================== */

/** Limpia un campo de texto (nombre/apellido/nacionalidad): corta en el primer dígito. */
function cleanField_(s) {
  return String(s || '').replace(/[:|]+/g, ' ').replace(/\s+/g, ' ').trim().replace(/\d.*$/, '').trim();
}

/** Extrae un DNI (7-8 dígitos) de un texto. loose=true permite grupos separados por espacios. */
function extractDni_(s, loose) {
  if (!s) return '';
  const str = String(s);
  let best = '', m;
  const reDot = /\b\d{1,2}\.\d{3}\.\d{3}\b/g;                 // 32.456.789
  while ((m = reDot.exec(str))) { const d = m[0].replace(/\D/g, ''); if (d.length >= 7 && d.length <= 8 && d.length > best.length) best = d; }
  const reNum = /\b\d{7,8}\b/g;                               // 32456789
  while ((m = reNum.exec(str))) { if (m[0].length > best.length) best = m[0]; }
  if (loose && !best) {
    const reSp = /\b\d{1,2}[ ]\d{3}[ ]\d{3}\b/g;              // 32 456 789 (solo en línea rotulada)
    while ((m = reSp.exec(str))) { const d2 = m[0].replace(/\D/g, ''); if (d2.length >= 7 && d2.length <= 8 && d2.length > best.length) best = d2; }
  }
  return best;
}

/** Parser del FRENTE por etiquetas en español. Devuelve campos (vacíos si no se leen). */
function parseDniFront_(text) {
  const out = { apellido: '', nombre: '', dni: '', sexo: '', nacimiento: '', nacionalidad: '', cuil: '', vencimiento: '' };
  if (!text) return out;
  const U = String(text).replace(/\r/g, '').toUpperCase();
  const lines = U.split('\n').map(function (s) { return s.trim(); }).filter(function (s) { return s; });

  function afterLabel(reLabel) {
    for (let i = 0; i < lines.length; i++) {
      const m = lines[i].match(reLabel);
      if (m) {
        const rest = lines[i].slice(m.index + m[0].length).replace(/^[:\s]+/, '').trim();
        if (rest) return rest;
        if (i + 1 < lines.length) return lines[i + 1];
      }
    }
    return '';
  }

  out.apellido = cleanField_(afterLabel(/APELLIDO/));
  out.nombre = cleanField_(afterLabel(/NOMBRES?/));
  const sx = afterLabel(/SEXO/);
  out.sexo = (/\bF\b|FEMENINO/.test(sx)) ? 'F' : (/\bM\b|MASCULINO/.test(sx)) ? 'M' : '';
  out.nacionalidad = cleanField_(afterLabel(/NACIONALIDAD/));
  out.nacimiento = afterLabel(/FECHA DE NAC|NACIMIENTO/).replace(/[|]+/g, ' ').replace(/\s+/g, ' ').trim();
  out.vencimiento = afterLabel(/VENCIMIENTO|VENC\.|EXPIRY/).replace(/[|]+/g, ' ').replace(/\s+/g, ' ').trim();
  out.dni = extractDni_(afterLabel(/DOCUMENTO|D\.?N\.?I\.?|N[º°O]\.?\s*DOC/), true) || extractDni_(U, false);
  const mCuil = U.match(/\b(\d{2}-?\d{8}-?\d)\b/);
  if (mCuil) out.cuil = mCuil[1].replace(/\D/g, '');
  return out;
}

/** Encuentra las líneas candidatas de la MRZ (dominadas por [A-Z0-9<], con '<'). */
function extractMrzLines_(text) {
  if (!text) return [];
  return String(text).toUpperCase().split('\n')
    .map(function (s) { return s.replace(/\s+/g, ''); }) // la MRZ no lleva espacios internos
    .filter(function (s) {
      if (s.length < 20 || s.length > 40 || s.indexOf('<') < 0) return false;
      const good = (s.match(/[A-Z0-9<]/g) || []).length;
      return good >= s.length * 0.8;
    });
}

/** Convierte YYMMDD de la MRZ a texto crudo dd/mm/aa (sin inferir el siglo). */
function mrzDate_(yymmdd) {
  if (!/^\d{6}$/.test(String(yymmdd))) return '';
  return yymmdd.slice(4, 6) + '/' + yymmdd.slice(2, 4) + '/' + yymmdd.slice(0, 2);
}

/** Parser del DORSO: MRZ (TD1, 3 líneas); si no hay MRZ legible, cae a etiquetas del frente. */
function parseDniBack_(text) {
  let out = { apellido: '', nombre: '', dni: '', sexo: '', nacimiento: '', nacionalidad: '', cuil: '', vencimiento: '' };
  if (!text) return out;
  const mrz = extractMrzLines_(text);
  if (mrz.length >= 3) {
    const L = mrz.slice(-3);                 // la MRZ suele ir al pie
    const l1 = L[0], l2 = L[1], l3 = L[2];
    const parts = l3.split('<<');
    out.apellido = String(parts[0] || '').replace(/</g, ' ').replace(/\s+/g, ' ').trim();
    out.nombre = String(parts.slice(1).join(' ') || '').replace(/</g, ' ').replace(/\s+/g, ' ').trim();
    const m2 = l2.match(/(\d{6})\d?([MFX<])(\d{6})/);
    if (m2) {
      out.nacimiento = mrzDate_(m2[1]);
      out.sexo = (m2[2] === 'M' || m2[2] === 'F') ? m2[2] : '';
      out.vencimiento = mrzDate_(m2[3]);
    }
    const mNac = l2.match(/[MFX<]\d{6}\d?([A-Z]{3})/);
    if (mNac) out.nacionalidad = mNac[1];
    out.dni = extractDni_(l1.replace(/</g, ' '), false) || extractDni_(L.join(' ').replace(/</g, ' '), false);
  } else {
    out = parseDniFront_(text);             // sin MRZ: reutilizar heurística de etiquetas
  }
  return out;
}

/** Combina frente + dorso (prefiere lo no vacío). Marca frenteDorsoDisiente si los DNI difieren. */
function mergeDniFields_(front, back) {
  front = front || {}; back = back || {};
  const pick = function (a, b) { a = String(a || '').trim(); return a || String(b || '').trim(); };
  const merged = {
    apellido: pick(front.apellido, back.apellido),
    nombre: pick(front.nombre, back.nombre),
    dni: '',
    sexo: pick(front.sexo, back.sexo),
    nacimiento: pick(front.nacimiento, back.nacimiento),
    nacionalidad: pick(front.nacionalidad, back.nacionalidad),
    cuil: pick(front.cuil, back.cuil),
    vencimiento: pick(front.vencimiento, back.vencimiento),
    frenteDorsoDisiente: false,
  };
  const df = normDni_(front.dni), db = normDni_(back.dni);
  if (df && db) { merged.dni = df; merged.frenteDorsoDisiente = (df !== db); }
  else { merged.dni = df || db; }
  return merged;
}

/* ===================== COTEJO CON CLIENTES ===================== */

/** ¿El texto 'haystack' contiene TODOS los tokens de 'needle' (≥2 tokens)? */
function containsAllTokens_(haystack, needle) {
  const toks = String(needle || '').split(' ').filter(Boolean);
  if (toks.length < 2) return false;  // evitar coincidencias por un solo token
  return toks.every(function (t) { return haystack.indexOf(t) >= 0; });
}

/**
 * Empareja el documento escaneado con un cliente.
 *  - PRIMARIO: por DNI (normDni_ contra dniNorm).
 *  - FALLBACK: por nombre (del archivo, y si no del OCR apellido+nombre) contra Clientes.
 *  - Varias coincidencias de nombre ⇒ sin coincidencia + marca ambiguo.
 */
function matchClienteForDni_(fields, personaFromFile) {
  const dn = normDni_(fields && fields.dni);
  if (dn) {
    const c = clienteByDni_(dn);
    if (c) return { cliente: c, via: 'Por DNI' };
  }
  const rows = clientesRows_();
  const byFile = normName_(personaFromFile);
  const byOcr = normName_([fields.apellido, fields.nombre].filter(Boolean).join(' '));
  const matchName = function (target) {
    if (!target) return [];
    return rows.filter(function (o) {
      const n = normName_(o.name);
      return n && (n === target || containsAllTokens_(n, target) || containsAllTokens_(target, n));
    });
  };
  let list = matchName(byFile);
  if (!list.length) list = matchName(byOcr);
  const uniq = {};
  list.forEach(function (o) { uniq[o.id || o.row] = o; });
  const dedup = Object.keys(uniq).map(function (k) { return uniq[k]; });
  if (dedup.length === 1) return { cliente: dedup[0], via: 'Por nombre (archivo)' };
  if (dedup.length > 1) return { cliente: null, via: 'Sin coincidencia', ambiguo: true };
  return { cliente: null, via: 'Sin coincidencia' };
}

/** Catálogo de anomalías. Devuelve {estado:'OK'|'Revisar', labels:[]}. */
function dniAnomalies_(fields, match, group) {
  const labels = [];
  const dn = normDni_(fields && fields.dni);
  const cli = match && match.cliente;

  if (!dn) labels.push('DNI ilegible');
  if (match && match.ambiguo) labels.push('Coincidencia ambigua por nombre');

  if (!cli) {
    if (dn && !(match && match.ambiguo)) labels.push('DNI no encontrado en Clientes');
  } else {
    if (dn && cli.dniNorm && dn !== cli.dniNorm) labels.push('DNI no coincide');
    const ocrName = normName_([fields.apellido, fields.nombre].filter(Boolean).join(' '));
    const cliName = normName_(cli.name);
    if (ocrName && cliName && ocrName !== cliName &&
      !containsAllTokens_(cliName, ocrName) && !containsAllTokens_(ocrName, cliName)) {
      labels.push('Nombre no coincide');
    }
  }

  if (fields && fields.frenteDorsoDisiente) labels.push('Frente y dorso no coinciden');
  if (!group.frente) labels.push('Falta el frente');
  if (!group.dorso) labels.push('Falta el dorso');
  if (!String(fields.apellido || '').trim() || !String(fields.nombre || '').trim()) labels.push('OCR incompleto');

  return { estado: labels.length ? 'Revisar' : 'OK', labels: labels };
}

/* ===================== INFORME POR CORREO ===================== */

/** Envía el informe al "Correo para resumen de avisos" (remitente opcional si es alias verificado). */
function enviarReporteDNI_(stats) {
  const to = String(getSetting_('Correo para resumen de avisos') || '').trim();
  if (!to) return;
  const wantFrom = String(getSetting_('Correo remitente de avisos') || '').trim();
  let fromAddr = '';
  if (wantFrom) { try { if (GmailApp.getAliases().indexOf(wantFrom) !== -1) fromAddr = wantFrom; } catch (e) { } }
  const opts = { from: fromAddr, replyTo: wantFrom || '' };
  try {
    sendBrandedEmail_(to, reporteDNIAsunto_(stats), reporteDNIPlano_(stats), reporteDNIHtml_(stats), opts);
  } catch (e) { logError_('enviarReporteDNI_', e); }
}

function reporteDNIAsunto_(stats) {
  return 'Verificación de DNI — ' + stats.nuevos + ' nuevos, ' + stats.anomaliasAbiertas.length +
    ' a revisar (' + fmtDate_(stats.hoy) + ')';
}

function reporteDNIPlano_(stats) {
  let t = 'Verificación de DNI (OCR)\nEjecución: ' + fmtDate_(stats.hoy) + '\n\n' +
    'Nuevos escaneados: ' + stats.nuevos + '\n' +
    'A revisar (de los nuevos): ' + stats.revisar + '\n' +
    'Sin coincidencia en Clientes (nuevos): ' + stats.sinCoincidencia + '\n' +
    'Anomalías abiertas (total): ' + stats.anomaliasAbiertas.length + '\n' +
    'Errores: ' + stats.errores + '\n';
  if (stats.topeAlcanzado) t += 'Nota: se alcanzó el tope por corrida; el resto continúa en la próxima.\n';
  t += '\n';
  if (stats.filasNuevas.length) {
    t += 'Nuevos:\n';
    stats.filasNuevas.forEach(function (f) {
      t += '- ' + f.persona + ' | ' + f.apenom + ' | DNI ' + f.dni + ' | ' + f.via + ' | ' + f.estado +
        (f.anomalias ? (' | ' + f.anomalias) : '') + '\n';
    });
    t += '\n';
  }
  if (stats.anomaliasAbiertas.length) {
    t += 'A revisar (abiertas):\n';
    stats.anomaliasAbiertas.forEach(function (f) {
      t += '- ' + f.persona + ' | ' + f.apenom + ' | DNI ' + f.dni + ' | ' + f.via + ' | ' + f.anomalias + '\n';
    });
  }
  t += '\nLa lectura por OCR es orientativa y puede contener errores; revisá manualmente los casos marcados «Revisar». No reemplaza la verificación humana.';
  return t;
}

function reporteDNIHtml_(stats) {
  const tile = function (n, label, color, bg) {
    return '<td style="padding:0 5px;width:25%"><div style="background:' + bg + ';border-radius:8px;padding:14px 6px;text-align:center">' +
      '<div style="font-size:26px;font-weight:bold;color:' + color + '">' + n + '</div>' +
      '<div style="font-size:11px;color:#667">' + label + '</div></div></td>';
  };
  const estadoCell = function (estado) {
    const rev = estado === 'Revisar';
    const col = rev ? '#c0392b' : '#2e7d46';
    const txt = rev ? '⚠ Revisar' : '✓ OK';
    return '<td align="center" style="padding:8px 10px;border-top:1px solid #eef0f3;color:' + col + ';font-weight:bold">' + txt + '</td>';
  };

  let nuevos = '';
  stats.filasNuevas.forEach(function (f, i) {
    nuevos += '<tr style="background:' + (i % 2 ? '#fafbfc' : '#fff') + '">' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.persona) + '</td>' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.apenom) + '</td>' +
      '<td align="center" style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.dni) + '</td>' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.via) + '</td>' +
      estadoCell(f.estado) +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.anomalias) + '</td></tr>';
  });
  if (!nuevos) nuevos = '<tr><td colspan="6" style="padding:12px 10px;color:#888;text-align:center">No se escanearon documentos nuevos en esta corrida.</td></tr>';

  let abiertas = '';
  stats.anomaliasAbiertas.forEach(function (f, i) {
    abiertas += '<tr style="background:' + (i % 2 ? '#fafbfc' : '#fff') + '">' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.persona) + '</td>' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.apenom) + '</td>' +
      '<td align="center" style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.dni) + '</td>' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3">' + esc_(f.via) + '</td>' +
      '<td style="padding:8px 10px;border-top:1px solid #eef0f3;color:#c0392b">' + esc_(f.anomalias) + '</td></tr>';
  });
  if (!abiertas) abiertas = '<tr><td colspan="5" style="padding:12px 10px;color:#2e7d46;text-align:center">No hay anomalías abiertas. 🎉</td></tr>';

  return '<p style="font-size:15px;margin:0 0 2px;color:#444">Ejecución del <b>' + esc_(fmtDate_(stats.hoy)) + '</b></p>' +
    '<p style="font-size:13px;margin:0 0 16px;color:#888">Cotejo por OCR de documentos de DNI subidos contra la hoja «Clientes».</p>' +
    '<table role="presentation" width="100%" style="border-collapse:collapse"><tr>' +
    tile(stats.nuevos, 'Nuevos', '#1c4587', '#eef3fb') +
    tile(stats.ok, 'OK (nuevos)', '#2e7d46', '#e7f5ec') +
    tile(stats.anomaliasAbiertas.length, 'A revisar', '#c0392b', '#fbeceb') +
    tile(stats.errores, 'Errores', '#8a7b00', '#fdf6e3') +
    '</tr></table>' +
    (stats.topeAlcanzado ? '<p style="font-size:12px;color:#8a7b00;margin:12px 0 0">Se alcanzó el tope de archivos por corrida; el resto continúa en la próxima ejecución.</p>' : '') +
    '<p style="font-size:14px;color:#333;margin:18px 0 6px"><b>Documentos escaneados en esta corrida</b></p>' +
    '<table role="presentation" width="100%" style="border-collapse:collapse;font-size:13px;border:1px solid #e3e6ec;border-radius:8px;overflow:hidden">' +
    '<tr style="background:#0b8043;color:#fff">' +
    '<th align="left" style="padding:9px 10px">Persona</th><th align="left" style="padding:9px 10px">Apellido/Nombre (OCR)</th>' +
    '<th align="center" style="padding:9px 10px">N° DNI</th><th align="left" style="padding:9px 10px">Coincidencia</th>' +
    '<th align="center" style="padding:9px 10px">Estado</th><th align="left" style="padding:9px 10px">Anomalías</th></tr>' +
    nuevos + '</table>' +
    '<p style="font-size:14px;color:#333;margin:20px 0 6px"><b>Todas las anomalías abiertas</b></p>' +
    '<table role="presentation" width="100%" style="border-collapse:collapse;font-size:13px;border:1px solid #e3e6ec;border-radius:8px;overflow:hidden">' +
    '<tr style="background:#c0392b;color:#fff">' +
    '<th align="left" style="padding:9px 10px">Persona</th><th align="left" style="padding:9px 10px">Apellido/Nombre (OCR)</th>' +
    '<th align="center" style="padding:9px 10px">N° DNI</th><th align="left" style="padding:9px 10px">Coincidencia</th>' +
    '<th align="left" style="padding:9px 10px">Anomalías</th></tr>' +
    abiertas + '</table>' +
    '<p style="margin-top:16px;font-size:12px;color:#888">La lectura por OCR es <b>orientativa</b> y puede contener errores; ' +
    'revisá manualmente los casos marcados «Revisar». No reemplaza la verificación humana.</p>';
}
