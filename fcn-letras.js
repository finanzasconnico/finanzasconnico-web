/* ═══════════════════════════════════════════════════════════════════════════
   fcn-letras.js — "¿Cuánto cobro?" para LECAPs y BONCAPs (letras/bonos capitalizables
   en pesos). Núcleo SIN DOM, compartido por Cuanto_Cobro_FCN.html (Suite) y por la
   página pública cuanto-cobro.html (repo web: es una COPIA de este archivo — la
   fuente de verdad es este, en la Suite. Si lo tocás acá, copiálo allá).

   Tres piezas:
     1) Datos de cada letra (emisión, vencimiento, valor final) — se LEEN de las hojas
        por ticker del Monitor (data/monitor.xlsx), así una letra nueva aparece sola
        cuando se agrega su hoja al Monitor. Nada hardcodeado.
     2) Precios en vivo de BYMA (API pública, misma que usa fcn-byma.js), con caché
        de 90 s en localStorage. Si falla, el llamador cae al precio del Monitor.
     3) calcular(): cuántos VN se compran con un monto, cuánto se cobra al vencimiento,
        ganancia directa, TNA, TEM, TEA y comparación con plazo fijo.

   Convenciones (verificadas contra acuantoesta.com.ar el 18/09/2026):
     - Precios de BYMA vienen cada 100 VN; acá todo se maneja POR 1 VN (precio/100).
     - Un LECAP/BONCAP paga UNA sola vez, al vencimiento: valor final por 1 VN
       (ej. S30S6 = 1,1753). Sale del flujo de la hoja del Monitor (Total / valor residual).
     - Ganancia directa = vf / precioConComision − 1. TNA = ganancia × 365 / días.
       TEM = (vf/precioConComision)^(30/días) − 1.
     - settlementType de BYMA: '1' = CI (contado inmediato), '2' = 24 hs.
   ═══════════════════════════════════════════════════════════════════════════ */
(function (global) {
  'use strict';

  var BASE = 'https://open.bymadata.com.ar/vanoms-be-core/rest/api/bymadata/free';
  var LS_KEY = 'fcn_letras_px_v1';
  var TTL_MS = 90 * 1000;
  var RE_LETRA = /^[ST]\d{1,2}[A-Z]\d$/;          // S30S6, S16O6, T15E7, T30J7…
  var DAY = 86400000;

  // Feriados nacionales de días hábiles (para contar la liquidación a 24 hs).
  // Solo los que están confirmados; los "puentes turísticos" se deciden por decreto
  // y no están acá — si uno cae en el medio, la liquidación se corre 1 día. Ampliar a mano.
  var FERIADOS = [
    '2026-10-12', '2026-11-23', '2026-12-08', '2026-12-25',
    '2027-01-01', '2027-02-15', '2027-02-16', '2027-03-24', '2027-03-26', '2027-04-02',
    '2027-05-25'
  ];

  function num(v) { v = parseFloat(v); return isFinite(v) ? v : 0; }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  var RE_ACENTOS = new RegExp('[' + String.fromCharCode(0x300) + '-' + String.fromCharCode(0x36f) + ']', 'g');   // marcas diacríticas tras normalize('NFD')
  function noAcc(s) { return String(s == null ? '' : s).normalize('NFD').replace(RE_ACENTOS, '').toLowerCase().trim(); }

  // fecha de una celda de Excel: Date, serial numérico o 'dd/mm/aaaa' / 'aaaa-mm-dd'
  function fechaCelda(v) {
    if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      var d = new Date(Math.round((v - 25569) * DAY));          // serial de Excel -> UTC
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    if (typeof v === 'string') {
      var m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
      m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    }
    return null;
  }
  function diasEntre(a, b) { return Math.round((Date.UTC(b.getFullYear(), b.getMonth(), b.getDate()) - Date.UTC(a.getFullYear(), a.getMonth(), a.getDate())) / DAY); }
  function esHabil(d) { var w = d.getDay(); return w !== 0 && w !== 6 && FERIADOS.indexOf(iso(d)) < 0; }
  function sumarHabiles(d, n) {
    var r = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    while (n > 0) { r.setDate(r.getDate() + 1); if (esHabil(r)) n--; }
    return r;
  }

  // Fecha y hora de "ahora" en Buenos Aires (quien mira la página puede estar en otro huso).
  function ahoraBA(d) {
    d = d || new Date();
    try {
      var p = {};
      new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Argentina/Buenos_Aires', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
        .formatToParts(d).forEach(function (x) { p[x.type] = x.value; });
      return { fecha: new Date(+p.year, +p.month - 1, +p.day), hora: +p.hour + (+p.minute) / 60 };
    } catch (e) {
      return { fecha: new Date(d.getFullYear(), d.getMonth(), d.getDate()), hora: d.getHours() + d.getMinutes() / 60 };
    }
  }
  var CIERRE_RUEDA = 17;   // BYMA cierra a las 17:00; después de eso la orden se ejecuta el próximo día hábil
  // Día en que se ejecutaría una orden dada HOY: hoy si es hábil y la rueda sigue abierta; si no, el próximo hábil.
  function fechaOperacion(hoy) {
    var b = ahoraBA(hoy);
    if (esHabil(b.fecha) && b.hora < CIERRE_RUEDA) return b.fecha;
    return sumarHabiles(b.fecha, 1);
  }

  // ─────────────────────────────────────────────────────────────────────────
  // 1) Datos de las letras desde el Monitor
  // ─────────────────────────────────────────────────────────────────────────

  // Nombres de hoja del Monitor que son una letra (S…/T…), sin abrir el libro entero.
  function nombresLetras(sheetNames) {
    return (sheetNames || []).filter(function (n) { return RE_LETRA.test(n); });
  }

  // rows = sheet_to_json(ws, {header:1, raw:true}) de UNA hoja de letra del Monitor.
  // OJO: las hojas NO son todas iguales (verificado 18/09/2026: en S16O6 faltan las
  // etiquetas de la izquierda y las columnas del flujo están corridas una posición).
  // Por eso el flujo se ubica por su ENCABEZADO ("Fecha" / "Valor residual" / "Total"), no por
  // posición fija, y el vencimiento sale de la fecha del último pago, no de una etiqueta.
  function parseHoja(ticker, rows) {
    var cF = -1, cR = -1, cT = -1, hdr = -1, i, j, r;
    for (i = 0; i < rows.length && hdr < 0; i++) {
      r = rows[i] || [];
      var f = -1, t = -1, v = -1;
      for (j = 0; j < r.length; j++) {
        var c = noAcc(r[j]);
        if (c === 'fecha') f = j; else if (c === 'total') t = j; else if (c === 'valor residual') v = j;
      }
      if (f >= 0 && t >= 0 && v >= 0) { hdr = i; cF = f; cR = v; cT = t; }
    }
    if (hdr < 0) return null;

    var flujo = [];
    for (i = hdr + 1; i < rows.length; i++) {
      r = rows[i] || [];
      var d = fechaCelda(r[cF]);
      if (d && typeof r[cT] === 'number') flujo.push({ fecha: d, residual: r[cR], total: r[cT] });
    }
    // Estructura esperada de una letra capitalizable: fila de emisión (residual = base, total 0),
    // fila de compra (total negativo, la usa el Monitor para su TIR) y UN solo pago positivo al final.
    var pagos = flujo.filter(function (x) { return x.total > 0; });
    var emisionRow = flujo.filter(function (x) { return typeof x.residual === 'number' && x.residual > 0; })[0];
    if (pagos.length !== 1 || !emisionRow) return null;      // no es un pago único: no la tratamos como letra
    var pago = pagos[0];

    var tna = null, precio = null;
    for (i = 0; i < rows.length; i++) {          // datos "de cortesía": la etiqueta puede estar en cualquier columna
      r = rows[i] || [];                         // (según la hoja); el valor es la celda de al lado
      for (j = 0; j < r.length - 1; j++) {
        var lab = noAcc(r[j]);
        if (lab === 'interes anual') tna = num(r[j + 1]);
        else if (lab === 'precio') precio = num(r[j + 1]);
      }
    }
    return {
      ticker: ticker,
      tipo: ticker.charAt(0) === 'S' ? 'LECAP' : 'BONCAP',
      emision: emisionRow.fecha,
      vto: pago.fecha,
      tnaEmision: tna > 0 ? tna : null,
      vf: pago.total / emisionRow.residual,      // valor final por 1 VN
      precioMonitor: precio > 0 ? precio : null  // por 1 VN, sin comisión (solo como respaldo si falla BYMA)
    };
  }

  // Valor final teórico: capitaliza por mes en base 30/360 (convención US: el día 31 pasa a 30 solo si el
  // otro día ya es ≥30). Solo para CONTROL cruzado contra el valor del Monitor, no para mostrar.
  function vfTeorico(emision, vto, tnaEmision) {
    if (!emision || !vto || !tnaEmision) return null;
    var d1 = emision.getDate(), d2 = vto.getDate();
    if (d1 === 31) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
    var d360 = (vto.getFullYear() - emision.getFullYear()) * 360 + (vto.getMonth() - emision.getMonth()) * 30 + (d2 - d1);
    return Math.pow(1 + tnaEmision / 12, d360 / 30);
  }

  // Lee un libro de SheetJS ya parseado (XLSX.read) y devuelve las letras vigentes, ordenadas por vto.
  function letrasDesdeLibro(XLSX, wb, hoy) {
    hoy = hoy || new Date();
    var out = [];
    nombresLetras(wb.SheetNames).forEach(function (n) {
      var ws = wb.Sheets[n];
      if (!ws) return;
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      var l = parseHoja(n, rows);
      if (l && diasEntre(hoy, l.vto) > 0) out.push(l);
    });
    out.sort(function (a, b) { return a.vto - b.vto || (a.ticker < b.ticker ? -1 : 1); });
    return out;
  }

  // Versión compacta para publicar (data/cuanto_cobro.json). La genera .github/scripts/generar_cuanto_cobro.js con ESTE
  // mismo parser, así la web pública no tiene que bajar el Monitor entero (1,7 MB) para leer 10 hojas.
  function letrasAJSON(letras) {
    return letras.map(function (l) {
      return { t: l.ticker, tp: l.tipo, e: l.emision ? iso(l.emision) : null, v: iso(l.vto), tna: l.tnaEmision, vf: l.vf, pm: l.precioMonitor };
    });
  }
  function letrasDesdeJSON(arr, hoy) {
    hoy = hoy || new Date();
    var out = (arr || []).map(function (x) {
      return { ticker: x.t, tipo: x.tp, emision: x.e ? fechaCelda(x.e) : null, vto: fechaCelda(x.v), tnaEmision: x.tna, vf: x.vf, precioMonitor: x.pm };
    }).filter(function (l) { return l.vto && diasEntre(hoy, l.vto) > 0; });      // se vuelve a filtrar por vigencia: el JSON puede tener días
    out.sort(function (a, b) { return a.vto - b.vto || (a.ticker < b.ticker ? -1 : 1); });
    return out;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // 2) Precios en vivo (BYMA)
  // ─────────────────────────────────────────────────────────────────────────
  // POST a BYMA con memoria de 60 s: fcn-bonos.js también pide 'public-bonds' y son ~1 MB; así se baja una sola vez.
  var _memo = {};
  function bymaPost(ep, body, force) {
    var k = ep + '|' + body, m = _memo[k];
    if (!force && m && Date.now() - m.ts < 60000) return m.p;          // "Actualizar" (force) siempre vuelve a pedir
    var p = fetch(BASE + '/' + ep, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: body })
      .then(function (r) { if (!r.ok) throw new Error('BYMA HTTP ' + r.status); return r.json(); });
    _memo[k] = { ts: Date.now(), p: p };
    p.catch(function () { if (_memo[k] && _memo[k].p === p) delete _memo[k]; });     // un fallo no se recuerda
    return p;
  }
  function fetchEndpoint(ep, force) {
    return bymaPost(ep, '{"page_size":2000}', force);    // sin page_size devuelve solo 189 filas y se pierden las S…/T…
  }

  // { S30S6: { ci:{px,bid,offer,ult,prev,vol,hora}, h24:{…} }, … }  (px = precio cada 100 VN)
  function agruparPrecios(payloads) {
    var out = {};
    payloads.forEach(function (p) {
      ((p && p.data) || []).forEach(function (x) {
        if (!RE_LETRA.test(x.symbol)) return;
        var pata = x.settlementType === '1' ? 'ci' : (x.settlementType === '2' ? 'h24' : null);
        if (!pata) return;
        var ult = num(x.trade) || num(x.closingPrice) || num(x.previousClosingPrice);
        var offer = num(x.offerPrice), bid = num(x.bidPrice);
        // Lo que pagaría un comprador es la punta vendedora; si falta o está lejos del último
        // (libro vacío/roto), se usa el último operado.
        var px = (offer > 0 && (!ult || Math.abs(offer / ult - 1) <= 0.03)) ? offer : ult;
        if (!(px > 0)) return;
        (out[x.symbol] = out[x.symbol] || {})[pata] = {
          px: px, ult: ult, bid: bid, offer: offer, prev: num(x.previousClosingPrice),
          vol: num(x.volume), hora: x.tradeHour || null
        };
      });
    });
    return out;
  }

  function leerCache() { try { return JSON.parse(localStorage.getItem(LS_KEY) || 'null'); } catch (e) { return null; } }
  function guardarCache(o) { try { localStorage.setItem(LS_KEY, JSON.stringify(o)); } catch (e) {} }

  // Promise<{ok, ts, stale, data}>. Nunca rechaza: si falla y hay caché vieja la devuelve con stale:true.
  var _inflight = null;
  function getPrecios(opts) {
    opts = opts || {};
    var c = leerCache();
    if (!opts.force && c && c.ts && (Date.now() - c.ts) < TTL_MS) return Promise.resolve({ ok: true, ts: c.ts, stale: false, data: c.data });
    if (_inflight) return _inflight;
    _inflight = Promise.all([fetchEndpoint('lebacs', opts.force), fetchEndpoint('public-bonds', opts.force)])
      .then(function (res) {
        var data = agruparPrecios(res);
        if (!Object.keys(data).length) throw new Error('BYMA sin datos de letras');
        var o = { ts: Date.now(), data: data };
        guardarCache(o);
        return { ok: true, ts: o.ts, stale: false, data: data };
      })
      .catch(function (e) {
        if (c && c.data) return { ok: true, ts: c.ts, stale: true, data: c.data, error: String(e && e.message || e) };
        return { ok: false, ts: null, stale: false, data: {}, error: String(e && e.message || e) };
      })
      .then(function (r) { _inflight = null; return r; });
    return _inflight;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // 3) Cálculo
  // ─────────────────────────────────────────────────────────────────────────
  // Precio (cada 100 VN) a usar para una letra según el plazo elegido. Devuelve {px, fuente, hora} o null.
  //   plazo: 'CI' | '24'.  Si la pata pedida no operó, se cae a la otra; si no hay BYMA, al Monitor.
  function elegirPrecio(letra, preciosBYMA, plazo) {
    var b = preciosBYMA && preciosBYMA[letra.ticker];
    if (b) {
      var pref = plazo === 'CI' ? 'ci' : 'h24', alt = plazo === 'CI' ? 'h24' : 'ci';
      var p = b[pref] || b[alt];
      if (p) return { px: p.px, fuente: 'BYMA', hora: p.hora, pata: b[pref] ? pref : alt, variacion: p.prev > 0 ? p.px / p.prev - 1 : null };
    }
    if (letra.precioMonitor) return { px: letra.precioMonitor * 100, fuente: 'Monitor', hora: null, pata: null, variacion: null };
    return null;
  }

  // o = { monto, comisionPct, plazo:'CI'|'24', tnaPF (en %), hoy:Date, precio:{px,…} }
  function calcular(letra, o) {
    var pr = o.precio;
    if (!pr || !(pr.px > 0)) return null;
    var op = fechaOperacion(o.hoy || new Date());          // cuándo se ejecuta la orden
    var liq = o.plazo === 'CI' ? op : sumarHabiles(op, 1);  // CI liquida el mismo día; 24 hs, el hábil siguiente
    var dias = diasEntre(liq, letra.vto);
    if (dias <= 0) return null;
    var com = (num(o.comisionPct) || 0) / 100;
    var pCom = pr.px / 100 * (1 + com);                 // precio por 1 VN, con comisión
    var ganDirecta = letra.vf / pCom - 1;
    var r = {
      ticker: letra.ticker, tipo: letra.tipo, vto: letra.vto, operacion: op, liquidacion: liq, dias: dias,
      precio: pr.px / 100, precioCom: pCom, vf: letra.vf, fuente: pr.fuente, hora: pr.hora, variacion: pr.variacion,
      ganDirecta: ganDirecta,
      tna: ganDirecta * 365 / dias,
      tem: Math.pow(letra.vf / pCom, 30 / dias) - 1,
      tea: Math.pow(letra.vf / pCom, 365 / dias) - 1
    };
    var monto = num(o.monto);
    if (monto > 0) {
      var vn = Math.floor(monto / pCom);                // los VN se compran de a 1: no hay fracciones
      var costo = vn * pCom;
      r.vn = vn;
      r.costo = costo;
      r.sobrante = monto - costo;
      r.cobro = vn * letra.vf;
      r.ganancia = r.cobro - costo;
      var tnaPF = num(o.tnaPF) / 100;
      if (tnaPF > 0) {
        r.pfCobro = monto * (1 + tnaPF * dias / 365);
        // mismo monto, mismos días: letra (lo cobrado + lo que sobró sin invertir) vs plazo fijo
        r.vsPF = (r.cobro + r.sobrante) - r.pfCobro;
      }
    }
    return r;
  }

  global.FCNLetras = {
    RE_LETRA: RE_LETRA,
    nombresLetras: nombresLetras,
    parseHoja: parseHoja,
    letrasDesdeLibro: letrasDesdeLibro,
    letrasAJSON: letrasAJSON,
    letrasDesdeJSON: letrasDesdeJSON,
    vfTeorico: vfTeorico,
    getPrecios: getPrecios,
    agruparPrecios: agruparPrecios,
    elegirPrecio: elegirPrecio,
    calcular: calcular,
    util: { num: num, noAcc: noAcc, fechaCelda: fechaCelda, diasEntre: diasEntre, sumarHabiles: sumarHabiles, esHabil: esHabil, bymaPost: bymaPost },   // los usa fcn-bonos.js
    fechaOperacion: fechaOperacion,
    fechaLiquidacion: function (hoy, plazo) { var op = fechaOperacion(hoy); return plazo === 'CI' ? op : sumarHabiles(op, 1); },
    ahoraBA: ahoraBA,
    diasEntre: diasEntre,
    FERIADOS: FERIADOS
  };
})(typeof window !== 'undefined' ? window : globalThis);
