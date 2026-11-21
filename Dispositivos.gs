/************************************************************
 * Dispositivos.gs — Módulo de captura anti-fraude (ampliado)
 * ----------------------------------------------------------
 * Enriquece la hoja "Dispositivos" con MUCHAS más señales por envío, de forma
 * CENTRALIZADA: ningún formulario se edita. Todos los formularios ya inyectan
 * `deviceCaptureJs_()` y llaman a `collectDeviceInfo()` + `logDeviceInfo_()`; acá:
 *
 *   • deviceCaptureJs_()  (en LoanManagerV2.gs)  →  agrega `antifraudExtrasJs_()`,
 *     un IIFE que ENVUELVE collectDeviceInfo (le suma huellas canvas/WebGL/audio,
 *     token persistente, Intl/conexión/reloj) e INSTALA, por sí solo, escuchas de
 *     comportamiento (pegado/teclas/foco/scroll) y de inputs de archivo (EXIF de
 *     las fotos, leído ANTES de la compresión). Nada de esto toca cada formulario.
 *   • logDeviceInfo_()  (en LoanManagerV2.gs)  →  delega en `recordDispositivo_()`,
 *     que escribe las 31 columnas base + columnas nuevas (aditivas, idempotentes).
 *
 * Todo es best-effort y con topes de tiempo: una API bloqueada por el iframe
 * (WebRTC/GPS) o un parser que falle NUNCA interrumpe el envío ni la captura base.
 *
 * NOTA (Ley 25.326): la recolección de datos técnicos del dispositivo/foto para
 * prevención de fraude debería estar declarada en los Términos y Condiciones.
 ************************************************************/

/** Columnas nuevas de "Dispositivos" (se agregan al final, después de las 31 base). */
const DISPOSITIVOS_COLS_V2 = [
  'Huella v2 (canvas+webgl+audio)', 'Canvas (hash)', 'WebGL (GPU)', 'WebGL (hash)', 'Audio (hash)',
  'Token persistente', 'Conexión (downlink/rtt)', 'Local (Intl)', 'Consistencia local',
  'Desfase de reloj (ms)', 'IP (WebRTC)', 'VPN/Datacenter?', 'Comportamiento', 'Pegado en identidad',
  'EXIF foto (GPS)', 'EXIF foto (equipo)', 'EXIF foto (fecha)', 'EXIF foto (software)',
];

/** ISPs que sugieren hosting/VPN/datacenter (indicativo; espeja analyze_devices.py). */
const AF_HOSTING_RE = /\b(vpn|proxy|hosting|datacenter|data center|cloud|ovh|digitalocean|linode|amazon|aws|azure|google llc|google cloud|hetzner|vultr|m247|leaseweb|choopa|colo|server)\b/i;

/**
 * Asegura (idempotente) que la hoja "Dispositivos" tenga las columnas nuevas al
 * final del encabezado. Devuelve el objeto hoja. No destructivo.
 */
function ensureDispositivosV2_(ss) {
  const sh = (typeof setupDispositivos_ === 'function')
    ? setupDispositivos_(ss || getSS_())
    : (ss || getSS_()).getSheetByName(CFG.SHEETS.DEVICES);
  if (!sh) return null;
  const lastCol = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = DISPOSITIVOS_COLS_V2.filter(h => head.indexOf(h) === -1);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
      .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff').setWrap(true);
    missing.forEach((h, i) => sh.setColumnWidth(lastCol + 1 + i, 140));
  }
  return sh;
}

/**
 * Sink ampliado. Reproduce la fila base (mismas 31 columnas y mismo cruce IP↔domicilio
 * que el logDeviceInfo_ original) y agrega las columnas nuevas derivadas de `meta.device`.
 * Firma idéntica a logDeviceInfo_(loan, clientId, meta) — es su implementación real.
 */
function recordDispositivo_(loan, clientId, meta) {
  const sh = ensureDispositivosV2_(getSS_());
  if (!sh) return null;
  loan = loan || {};
  const dev = (meta && meta.device) || {};
  const loc = dev.loc || {};
  const now = (meta && meta.signedAt) || new Date();
  const na = x => (x == null || x === '') ? 'N/D' : x;

  // Cruce liviano de la ubicación IP contra el domicilio declarado (igual que el original).
  let ipLoc = String(loc.ipText || '');
  const dom = String(loan.direccion || '').toLowerCase();
  if (ipLoc && dom) {
    const hit = ipLoc.toLowerCase().split(/[,·]/).some(t => t.trim().length > 2 && dom.indexOf(t.trim()) !== -1);
    if (!hit) ipLoc = '⚠ (revisar vs domicilio) ' + ipLoc;
  }

  // --- fila BASE (columnas 1..31, posicional — idéntica al logDeviceInfo_ previo) ---
  const base = [
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
  ];

  // --- columnas NUEVAS (derivadas) ---
  const canvas = String(dev.canvas || '');
  const webgl = dev.webgl || {};
  const audio = String(dev.audio || '');
  const fpParts = [(meta && meta.deviceFp) || '', canvas, webgl.hash || '', audio].filter(String);
  const fpv2 = (fpParts.length && typeof hexDigest_ === 'function') ? hexDigest_(fpParts.join('|')) : '';

  const conn = [dev.downlink ? (dev.downlink + 'Mbps') : '', dev.rtt ? (dev.rtt + 'ms') : '',
    dev.saveData ? 'saveData' : ''].filter(String).join(' · ');

  // Consistencia local (indicativo): idioma/zona del dispositivo vs país de la IP.
  const langs = String(dev.langs || '').toLowerCase();
  const tz = String(dev.tz || '');
  const ipCountryAR = /argentina/i.test(ipLoc);
  const deviceAR = /\bes(-|_)?ar\b|\bes\b/.test(langs) || /argentina|buenos_aires|cordoba|mendoza|catamarca|salta|tucuman|ushuaia|jujuy|san_juan|san_luis|la_rioja/i.test(tz);
  let consistencia = '';
  if (ipLoc && ipLoc !== 'N/D' && !ipCountryAR && tz && !deviceAR) consistencia = '⚠ idioma/zona/IP fuera de AR';
  else if (ipLoc && ipLoc !== 'N/D' && !ipCountryAR) consistencia = '⚠ IP fuera de AR';

  // Desfase de reloj: hora del cliente (cuando corrió la captura) vs hora del servidor.
  // Incluye la latencia de red (no se corrige): sólo desfases grandes (≥2 min) son señal.
  let skew = '';
  const clientNow = (meta && meta.clientNow) || dev.clientNow;
  if (clientNow) { const d = now.getTime() - Number(clientNow); if (isFinite(d)) skew = Math.round(d); }

  const rtcIp = String(dev.rtcIp || '');
  const vpn = (ipLoc && AF_HOSTING_RE.test(ipLoc)) ? 'SÍ' : (dev.vpn ? 'SÍ' : 'No');

  const b = dev.behav || {};
  const comportamiento = [
    (b.keys != null ? ('teclas ' + b.keys) : ''),
    (b.back != null ? ('borrados ' + b.back) : ''),
    (b.pastes ? ('pegados ' + b.pastes) : ''),
    (b.hiddenMs ? ('oculto ' + Math.round(b.hiddenMs / 1000) + 's') : ''),
    (b.scrollPct != null ? ('scroll ' + b.scrollPct + '%') : ''),
    (b.autofill ? 'autocompletado' : ''),
    (b.pointer === false ? 'sin-mouse' : ''),
  ].filter(String).join(' · ');
  const pasteId = (b.pasteId ? 'SÍ' : 'No');

  // EXIF de la primera foto con datos (si la hay). Marca ⚠ si GPS/equipo no cuadran.
  const img = (Array.isArray(dev.images) ? dev.images : []).filter(x => x && (x.gps || x.make || x.model || x.dateTime))[0] || {};
  let exifGps = String(img.gps || '');
  if (exifGps && dom) {
    const city = String(img.city || '').toLowerCase();
    if (city && dom.indexOf(city) === -1) exifGps = '⚠ (revisar vs domicilio) ' + exifGps;
  }
  const exifDevTxt = [img.make, img.model].filter(String).join(' ');
  let exifDev = exifDevTxt;
  if (exifDevTxt && dev.model && String(dev.model) !== 'N/D') {
    const a = exifDevTxt.toLowerCase(), m = String(dev.model).toLowerCase();
    if (m && a.indexOf(m) === -1 && m.indexOf(a) === -1) exifDev = '⚠ (≠ equipo de envío) ' + exifDevTxt;
  }

  const extra = [
    fpv2, canvas, String(webgl.gpu || ''), String(webgl.hash || ''), audio,
    String(dev.token || ''), conn, String(dev.locale || ''), consistencia,
    skew, rtcIp, vpn, comportamiento, pasteId,
    exifGps, exifDev, String(img.dateTime || ''), String(img.software || ''),
  ];

  // Escribe por NOMBRE de encabezado: base en 1..31, nuevas donde estén (robusto si el orden cambia).
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const row = new Array(head.length).fill('');
  for (let i = 0; i < Math.min(base.length, head.length); i++) row[i] = base[i];
  DISPOSITIVOS_COLS_V2.forEach((name, i) => {
    const col = head.indexOf(name);
    if (col >= 0) row[col] = extra[i];
  });
  sh.appendRow(row);
  invalidateDispositivosRows_(); // coherencia dentro de la misma ejecución (V-39)
  return sh;
}

/* ============== V-39 — señales por persona (misma IP/dispositivo/GPS) ==============
 * Capa de LECTURA de la hoja "Dispositivos" para el límite configurable de personas
 * distintas que comparten una señal (ver checkUbicacionV39_ en Validaciones.gs).
 * Una sola lectura getDataRange por ejecución, memoizada (misma convención que
 * _clientesRowsCache). Si recordDispositivo_ agrega una fila, invalida la caché. */

/** 'lat,long' → {lat, lon} o null si no es un par numérico válido. */
function parseGps_(s) {
  const m = /^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$/.exec(String(s || ''));
  if (!m) return null;
  const lat = Number(m[1]), lon = Number(m[2]);
  return (isFinite(lat) && isFinite(lon)) ? { lat: lat, lon: lon } : null;
}

/** Distancia haversine en metros entre dos puntos {lat, lon}. */
function haversineM_(a, b) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (b.lat - a.lat) * rad, dLon = (b.lon - a.lon) * rad;
  const s = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return 2 * R * Math.asin(Math.sqrt(s));
}

let _dispSignalRowsCache = null;
function invalidateDispositivosRows_() { _dispSignalRowsCache = null; }

/** Filas de "Dispositivos" reducidas a las señales que usa V-39. Columnas base por
 *  POSICIÓN (mismo orden que recordDispositivo_): 1 Fecha, 3 clientId, 4 DNI, 7 huella,
 *  22 IP, 24 GPS, 25 precisión GPS. 'Token persistente' se resuelve por encabezado. */
function dispositivosSignalRows_() {
  if (_dispSignalRowsCache) return _dispSignalRowsCache;
  const sh = getSS_().getSheetByName(CFG.SHEETS.DEVICES);
  if (!sh || sh.getLastRow() < 2) { _dispSignalRowsCache = []; return _dispSignalRowsCache; }
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const cTok = head.indexOf('Token persistente'); // -1 si la hoja aún no tiene la columna
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  _dispSignalRowsCache = data.map(r => {
    const ip = String(r[21] || '').trim();
    return {
      t: (r[0] instanceof Date) ? r[0] : null,
      dniNorm: (typeof normDni_ === 'function') ? normDni_(r[3]) : String(r[3] || '').replace(/\D/g, ''),
      clientId: String(r[2] || '').trim(),
      ip: (ip && ip !== 'N/D') ? ip : '',
      fp: String(r[6] || '').trim(),
      token: cTok >= 0 ? String(r[cTok] || '').trim() : '',
      gps: parseGps_(r[23]),
      gpsAcc: Number(r[24]) || 0,
    };
  }).filter(o => o.dniNorm || o.clientId);
  return _dispSignalRowsCache;
}

/**
 * Cuenta, por señal, las personas DISTINTAS (por DNI normalizado) que comparten la señal
 * del solicitante en la hoja "Dispositivos". Excluye las filas propias (mismo dniNorm; o
 * mismo clientId cuando la fila no tiene DNI).
 *   sig = { dniNorm, clientId, ip, deviceFp, token, gps:{lat,lon}|null }
 *   cfg = { radioM, ipDias }  — IP: solo filas dentro de la ventana; GPS: distancia ≤
 *         max(radioM, precisión informada); huella/token: igualdad exacta.
 * Devuelve { ip:Set, fp:Set, token:Set, gps:Set } de dniNorm ajenos por señal.
 */
function contarCoincidenciasSenales_(sig, cfg) {
  const out = { ip: new Set(), fp: new Set(), token: new Set(), gps: new Set() };
  const rows = dispositivosSignalRows_();
  const now = new Date();
  const ipMs = Math.max(1, Number(cfg.ipDias) || 30) * 24 * 3600 * 1000;
  const radioM = Math.max(1, Number(cfg.radioM) || 50);
  // Identidad canónica SIEMPRE por DNI: una fila sin DNI pero con ID Cliente se resuelve
  // vía "Clientes" para que la misma persona no cuente dos veces (una por DNI y otra por ID).
  let idToDni = null;
  try {
    if (typeof clientesRows_ === 'function') {
      idToDni = {};
      clientesRows_().forEach(o => { if (o.id && o.dniNorm) idToDni[o.id] = o.dniNorm; });
    }
  } catch (e) { idToDni = null; }
  for (const r of rows) {
    const own = (sig.dniNorm && r.dniNorm === sig.dniNorm) ||
      (sig.clientId && r.clientId && r.clientId === sig.clientId);
    if (own) continue;
    const quien = r.dniNorm || (idToDni && idToDni[r.clientId]) || r.clientId;
    if (!quien) continue;
    if (sig.ip && r.ip === sig.ip && r.t && (now.getTime() - r.t.getTime()) <= ipMs) out.ip.add(quien);
    if (sig.deviceFp && r.fp === sig.deviceFp) out.fp.add(quien);
    if (sig.token && r.token === sig.token) out.token.add(quien);
    if (sig.gps && r.gps && haversineM_(sig.gps, r.gps) <= Math.max(radioM, r.gpsAcc || 0)) out.gps.add(quien);
  }
  return out;
}

/**
 * JS cliente AMPLIADO que se agrega a `deviceCaptureJs_()`. Un IIFE que:
 *  1) envuelve collectDeviceInfo para sumarle huellas y señales nuevas al objeto `dev`;
 *  2) instala, por delegación en document, la captura de comportamiento y de EXIF de
 *     los inputs de archivo (sin referenciar ningún formulario).
 * Devuelve una cadena (sin `${` ni backticks adentro) lista para concatenar.
 */
function antifraudExtrasJs_() {
  return `
      /* ===== Captura anti-fraude ampliada (módulo Dispositivos) — best-effort ===== */
      (function(){
        if (typeof collectDeviceInfo !== 'function') return;
        function h32(s){ var h=0x811c9dc5; s=String(s||''); for(var i=0;i<s.length;i++){ h^=s.charCodeAt(i); h=(h*0x01000193)>>>0; } return ('0000000'+h.toString(16)).slice(-8); }
        function withTO(p,ms){ return Promise.race([Promise.resolve(p), new Promise(function(r){ setTimeout(function(){r(null);},ms); })]); }

        /* --- comportamiento (listeners globales, sin tocar cada formulario) --- */
        var BEHAV={ keys:0, back:0, pastes:0, pasteId:false, hiddenMs:0, scrollPct:0, autofill:false, pointer:false, _hid:0 };
        var ID_RE=/dni|cuil|nombre|name|correo|email|tel|phone/i;
        try{
          document.addEventListener('keydown', function(e){ BEHAV.keys++; if(e.key==='Backspace'||e.key==='Delete') BEHAV.back++; }, true);
          document.addEventListener('paste', function(e){ BEHAV.pastes++; var t=e.target||{}; if(ID_RE.test((t.id||'')+' '+(t.name||''))) BEHAV.pasteId=true; }, true);
          document.addEventListener('pointermove', function(){ BEHAV.pointer=true; }, {passive:true, once:true});
          document.addEventListener('touchstart', function(){ BEHAV.pointer=true; }, {passive:true, once:true});
          document.addEventListener('animationstart', function(e){ if(e.animationName && /autofill/i.test(e.animationName)) BEHAV.autofill=true; }, true);
          window.addEventListener('scroll', function(){ try{ var d=document.documentElement; var p=Math.round((window.scrollY+window.innerHeight)/Math.max(1,d.scrollHeight)*100); if(p>BEHAV.scrollPct) BEHAV.scrollPct=Math.min(100,p); }catch(e){} }, {passive:true});
          document.addEventListener('visibilitychange', function(){ if(document.hidden){ BEHAV._hid=Date.now(); } else if(BEHAV._hid){ BEHAV.hiddenMs+=Date.now()-BEHAV._hid; BEHAV._hid=0; } }, true);
        }catch(e){}

        /* --- EXIF de las fotos (leído del File ORIGINAL, antes de comprimir) --- */
        var IMAGES=[];
        function rat(dv,o,le){ var n=dv.getUint32(o,le), d=dv.getUint32(o+4,le); return d? n/d : 0; }
        function tlen(t){ return t===1||t===2||t===7?1 : t===3?2 : t===4||t===9?4 : t===5||t===10?8 : 1; }
        function parseExif(buf){
          try{
            var dv=new DataView(buf); if(dv.getUint16(0)!==0xFFD8) return null;
            var off=2, L=dv.byteLength, app1=-1;
            while(off+4<L){ var mk=dv.getUint16(off); if(mk===0xFFE1){ app1=off; break; } if((mk&0xFF00)!==0xFF00) break; off+=2+dv.getUint16(off+2); }
            if(app1<0) return null;
            if(dv.getUint32(app1+4)!==0x45786966) return null; // 'Exif'
            var tiff=app1+10, le=dv.getUint16(tiff)===0x4949;
            function u16(o){return dv.getUint16(o,le);} function u32(o){return dv.getUint32(o,le);}
            function str(o,c){ var s=''; for(var i=0;i<c;i++){ var ch=dv.getUint8(o+i); if(!ch)break; s+=String.fromCharCode(ch); } return s.trim(); }
            var out={}, exifPtr=0, gpsPtr=0;
            function walk(dir, tags){ var n=u16(dir); for(var i=0;i<n;i++){ var e=dir+2+i*12; var tag=u16(e), ty=u16(e+2), cnt=u32(e+4); var vo=(tlen(ty)*cnt<=4)?(e+8):(tiff+u32(e+8)); tags(tag,ty,cnt,vo); } }
            var ifd0=tiff+u32(tiff+4);
            walk(ifd0, function(tag,ty,cnt,vo){
              if(tag===0x010F) out.make=str(vo,cnt);
              else if(tag===0x0110) out.model=str(vo,cnt);
              else if(tag===0x0131) out.software=str(vo,cnt);
              else if(tag===0x0132) out.dateTime=str(vo,cnt);
              else if(tag===0x8769) exifPtr=tiff+u32(vo);
              else if(tag===0x8825) gpsPtr=tiff+u32(vo);
            });
            if(exifPtr) walk(exifPtr, function(tag,ty,cnt,vo){ if(tag===0x9003) out.dateTime=str(vo,cnt); });
            if(gpsPtr){ var gps={}; walk(gpsPtr, function(tag,ty,cnt,vo){
              if(tag===0x0001) gps.latRef=str(vo,cnt); else if(tag===0x0003) gps.lonRef=str(vo,cnt);
              else if(tag===0x0002) gps.lat=[rat(dv,vo,le),rat(dv,vo+8,le),rat(dv,vo+16,le)];
              else if(tag===0x0004) gps.lon=[rat(dv,vo,le),rat(dv,vo+8,le),rat(dv,vo+16,le)];
            });
              if(gps.lat&&gps.lon){ var la=gps.lat[0]+gps.lat[1]/60+gps.lat[2]/3600; var lo=gps.lon[0]+gps.lon[1]/60+gps.lon[2]/3600;
                if((gps.latRef||'N').toUpperCase()==='S') la=-la; if((gps.lonRef||'E').toUpperCase()==='W') lo=-lo;
                out.gps=la.toFixed(5)+','+lo.toFixed(5); }
            }
            return out;
          }catch(e){ return null; }
        }
        function readExif(input){
          try{ var f=input&&input.files&&input.files[0]; if(!f||!/image\\/jpe?g/i.test(f.type||'')) return;
            var fr=new FileReader();
            fr.onload=function(){ var ex=parseExif(fr.result); IMAGES.push({ field:(input.id||input.name||''), make:(ex&&ex.make)||'', model:(ex&&ex.model)||'', dateTime:(ex&&ex.dateTime)||'', software:(ex&&ex.software)||'', gps:(ex&&ex.gps)||'' }); };
            fr.readAsArrayBuffer(f.slice(0, 131072)); // los 128 KB iniciales alcanzan para el bloque EXIF
          }catch(e){}
        }
        try{ document.addEventListener('change', function(e){ var t=e.target; if(t&&t.type==='file') readExif(t); }, true); }catch(e){}

        /* --- huellas de hardware --- */
        function afCanvas(){ try{ var c=document.createElement('canvas'); c.width=240; c.height=60; var x=c.getContext('2d'); if(!x) return ''; x.textBaseline='top'; x.font="14px 'Arial'"; x.fillStyle='#f60'; x.fillRect(10,10,120,30); x.fillStyle='#069'; x.fillText('Impulso \\uD83D\\uDD12 fp',12,14); x.strokeStyle='rgba(0,80,160,.6)'; x.beginPath(); x.arc(60,30,20,0,Math.PI*1.7); x.stroke(); return h32(c.toDataURL()); }catch(e){ return ''; } }
        function afWebgl(){ try{ var c=document.createElement('canvas'); var gl=c.getContext('webgl')||c.getContext('experimental-webgl'); if(!gl) return {}; var gpu=''; var dbg=gl.getExtension('WEBGL_debug_renderer_info'); if(dbg){ gpu=(gl.getParameter(dbg.UNMASKED_VENDOR_WEBGL)||'')+' · '+(gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL)||''); } var p=[gl.getParameter(gl.VENDOR),gl.getParameter(gl.RENDERER),gl.getParameter(gl.VERSION),gl.getParameter(gl.SHADING_LANGUAGE_VERSION),gl.getParameter(gl.MAX_TEXTURE_SIZE)].join('|'); return { gpu:String(gpu).slice(0,120), hash:h32(p) }; }catch(e){ return {}; } }
        function afAudio(){ return new Promise(function(res){ try{ var OC=window.OfflineAudioContext||window.webkitOfflineAudioContext; if(!OC){ return res(''); } var ctx=new OC(1,5000,44100); var osc=ctx.createOscillator(); osc.type='triangle'; osc.frequency.value=10000; var comp=ctx.createDynamicsCompressor(); osc.connect(comp); comp.connect(ctx.destination); osc.start(0); ctx.startRendering(); ctx.oncomplete=function(ev){ try{ var d=ev.renderedBuffer.getChannelData(0); var s=0; for(var i=4000;i<5000;i++) s+=Math.abs(d[i]); res(h32(String(s))); }catch(e){ res(''); } }; setTimeout(function(){ res(''); },800); }catch(e){ res(''); } }); }
        function afToken(){ try{ var k='ic_fp_token', t=localStorage.getItem(k); if(!t){ t=(Date.now().toString(36)+Math.random().toString(36).slice(2,10)); localStorage.setItem(k,t); } return t; }catch(e){ return ''; } }
        function afWebrtc(){ return new Promise(function(res){ try{ var RTC=window.RTCPeerConnection||window.webkitRTCPeerConnection; if(!RTC){ return res(''); } var pc=new RTC({iceServers:[]}); var ips={}; pc.onicecandidate=function(e){ if(!e||!e.candidate){ return; } var m=/([0-9]{1,3}(\\.[0-9]{1,3}){3})/.exec(e.candidate.candidate||''); if(m) ips[m[1]]=1; }; pc.createDataChannel('x'); pc.createOffer().then(function(o){ return pc.setLocalDescription(o); }).catch(function(){}); setTimeout(function(){ try{pc.close();}catch(e){} res(Object.keys(ips).join(',')); },700); }catch(e){ res(''); } }); }

        var _orig=collectDeviceInfo;
        collectDeviceInfo=function(){
          return Promise.resolve(_orig()).then(function(dev){
            dev=dev||{};
            try{ dev.canvas=afCanvas(); }catch(e){}
            try{ dev.webgl=afWebgl(); }catch(e){}
            try{ dev.token=afToken(); }catch(e){}
            try{ var n=navigator.connection||{}; dev.downlink=n.downlink||''; dev.rtt=n.rtt||''; dev.saveData=!!n.saveData; }catch(e){}
            try{ var r=Intl.DateTimeFormat().resolvedOptions(); dev.locale=[r.locale,r.calendar,r.hourCycle,r.numberingSystem].filter(Boolean).join('/'); }catch(e){}
            try{ dev.clientNow=Date.now(); }catch(e){}
            try{ dev.behav={ keys:BEHAV.keys, back:BEHAV.back, pastes:BEHAV.pastes, pasteId:BEHAV.pasteId, hiddenMs:BEHAV.hiddenMs, scrollPct:BEHAV.scrollPct, autofill:BEHAV.autofill, pointer:BEHAV.pointer }; }catch(e){}
            try{ dev.images=IMAGES.slice(0,4); }catch(e){}
            return Promise.all([ withTO(afAudio(),900), withTO(afWebrtc(),800) ]).then(function(r){
              try{ dev.audio=r[0]||''; dev.rtcIp=r[1]||''; }catch(e){}
              return dev;
            }).catch(function(){ return dev; });
          });
        };
      })();`;
}
