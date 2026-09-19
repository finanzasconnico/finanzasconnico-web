/* ═══════════════════════════════════════════════════════════════════════════
   fcn-bonos.js — "¿Cuánto cobro?" Fase 2: bonos y ONs en DÓLARES (soberanos AL/GD/AE/AN/AO,
   provinciales, BOPREALES y obligaciones negociables). Núcleo SIN DOM, hermano de
   fcn-letras.js (que debe cargarse ANTES: reusa sus utilidades y el cálculo de fechas de
   liquidación). Compartido por Cuanto_Cobro_FCN.html (Suite) y cuanto-cobro.html (repo
   web: COPIA de este archivo — la fuente de verdad es este; si lo tocás, copiálo allá).

   Qué hace:
     1) Lee de las hojas por ticker del Monitor (data/monitor.xlsx) el CALENDARIO de cobros
        completo de cada bono: fecha, cupón (interés), amortización y total.
     2) Precios en vivo de BYMA (endpoints public-bonds + negociable-obligations).
     3) calcularBono(): con un monto en USD, cuántos VN se compran, cuánto se cobra en los
        próximos 12 meses y hasta el vencimiento, y la TIR con esos flujos.

   VERIFICADO (18/09/2026, sobre las 231 hojas con precio y TIR):
     - Con los flujos de la hoja y el precio del Monitor, la TIR recalculada reproduce la
       "TIR efectiva" del Monitor con error mediano ~1e-9 en 228 de 231. → los flujos se
       leen bien. Los que no cierran (>0,5 pt) se marcan validado:false y no se muestran.
     - El precio de mercado ya es SUCIO (incluye intereses corridos): sumarlos aparte
       daba error de hasta 37 pt de TIR (228 vs 1 a favor del precio tal cual).
     - El precio del Monitor incluye una comisión de 0,5% (ratio BYMA/Monitor: mediana 0,996).
     - Las hojas NO son todas iguales: la base de los importes es 1000 en casi todas pero
       100 en otras (AE38…) → se normaliza por "Nominales a comprar". Hay mínimos de
       compra ("Nominales mínimos": p.ej. CO32 = 10.000 VN) y múltiplos.
     - El símbolo de BYMA para comprar con dólar MEP es el "…D" (el Monitor lo marca en el
       campo "Ticker" de cada hoja), incluso para los que dice "Cable".
   ═══════════════════════════════════════════════════════════════════════════ */
(function (global) {
  'use strict';
  var L = global.FCNLetras;
  if (!L || !L.util) throw new Error('fcn-bonos.js necesita fcn-letras.js cargado antes');
  var U = L.util;
  var num = U.num, noAcc = U.noAcc, fechaCelda = U.fechaCelda, diasEntre = U.diasEntre;

  var LS_KEY = 'fcn_bonos_px_v1';
  var TTL_MS = 90 * 1000;
  var COM_MONITOR = 0.005;               // el precio del Monitor ya trae 0,5% de comisión
  var TOL_TIR = 0.005;                   // discrepancia máxima aceptada entre la TIR recalculada y la del Monitor
  var TOL_PRECIO = 0.15;                 // BYMA vs Monitor: más lejos que esto => símbolo/precio dudoso
  // Hojas del Monitor que no son un instrumento (resúmenes, cajas de ahorro, tasas…)
  var NO_INSTRUMENTO = { 'Hoja3': 1, 'FX': 1, 'Soberanos': 1, 'Corporativos': 1, 'Letras-Bonos $': 1, 'OFFSHORE': 1 };

  // ─────────────────────────────────────────────────────────────────────────
  // TIR (efectiva anual, base 365) por bisección
  // ─────────────────────────────────────────────────────────────────────────
  function xirr(cfs) {                   // cfs: [{t: años desde hoy, v: importe}] — el primero es la compra (negativo)
    var lo = -0.99, hi = 10;
    function f(r) { var s = 0; for (var i = 0; i < cfs.length; i++) s += cfs[i].v / Math.pow(1 + r, cfs[i].t); return s; }
    var flo = f(lo), fhi = f(hi);
    if (!isFinite(flo) || !isFinite(fhi) || flo * fhi > 0) return null;
    for (var k = 0; k < 200; k++) {
      var m = (lo + hi) / 2, fm = f(m);
      if (flo * fm <= 0) { hi = m; } else { lo = m; flo = fm; }
    }
    return (lo + hi) / 2;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // 1) Lectura de una hoja de bono
  // ─────────────────────────────────────────────────────────────────────────
  function categoria(emisor) {
    var e = String(emisor || '');
    if (e === 'Tesoro Nacional') return 'Soberano';
    if (e === 'BCRA') return 'BOPREAL';
    if (/^(Provincia|Ciudad)/.test(e)) return 'Provincial';
    return 'ON';
  }

  // rows = sheet_to_json(ws, {header:1, raw:true, defval:null}). Devuelve el bono o
  // {descartado:'motivo'} (para poder informar POR QUÉ una hoja no entra), o null si no es de bono.
  function parseHoja(ticker, rows) {
    var i, j, r, hdr = -1, c = {};
    for (i = 0; i < rows.length && hdr < 0; i++) {           // encabezado del flujo, ubicado por sus títulos
      r = rows[i] || [];
      var m = {};
      for (j = 0; j < r.length; j++) {
        var k = noAcc(r[j]);
        if (k === 'fecha' || k === 'valor residual' || k === 'amortizacion usd' || k === 'interes' || k === 'total') m[k] = j;
      }
      if (m.fecha != null && m.total != null && m['valor residual'] != null && m['amortizacion usd'] != null && m.interes != null) { hdr = i; c = m; }
    }
    if (hdr < 0) return null;

    var meta = {};                                            // etiqueta -> celda de al lado (en cualquier columna)
    for (i = 0; i < rows.length; i++) {
      r = rows[i] || [];
      // en la fila del encabezado del flujo solo se lee lo que está a la IZQUIERDA de él (ahí vive, p.ej., "Ley")
      for (j = 0; j < r.length - 1 && (i !== hdr || j < c.fecha); j++) {
        if (typeof r[j] !== 'string' || r[j + 1] === null || r[j + 1] === '') continue;
        var lab = noAcc(r[j]);
        if (lab && meta[lab] === undefined) meta[lab] = r[j + 1];
      }
    }
    var moneda = meta['moneda de cobro'];
    if (moneda !== 'MEP' && moneda !== 'Cable') return null; // solo instrumentos que pagan en dólares

    var base = num(meta['nominales a comprar']);              // escala de TODOS los importes de la hoja (1000, 100…)
    if (!(base > 0)) return { descartado: 'sin "Nominales a comprar"' };

    var liqFile = fechaCelda(meta['liquidacion']);
    var flujos = [], compra = null;
    for (i = hdr + 1; i < rows.length; i++) {
      r = rows[i] || [];
      var d = fechaCelda(r[c.fecha]), tot = r[c.total];
      if (!d || typeof tot !== 'number') continue;
      if (tot < 0) { compra = -tot; continue; }               // fila "compra": lo que el Monitor paga (precio con comisión)
      if (tot === 0) continue;
      flujos.push({ fecha: d, cupon: num(r[c.interes]) / base, amort: num(r[c['amortizacion usd']]) / base, total: tot / base });
    }
    if (!flujos.length) return { descartado: 'sin cobros futuros' };

    var vto = fechaCelda(meta['fecha de vencimiento']) || flujos[flujos.length - 1].fecha;
    var bono = {
      ticker: ticker,
      simbolo: typeof meta['ticker'] === 'string' ? meta['ticker'] : ticker,   // símbolo BYMA en dólar MEP (…D)
      emisor: meta['emisor'] || '', cat: categoria(meta['emisor']), calif: meta['calificacion'] || '',
      ley: meta['ley'] || '', cobra: moneda, tipoTasa: meta['tipo de tasa'] || '',
      cuponAnual: num(meta['interes anual']), frecuencia: meta['frecuencia de cobro de intereses'] || '',
      amortTxt: meta['amortizacion'] || '', mesesCupon: meta['fechas de cobro de intereses'] || '',
      vto: vto, minNom: num(meta['nominales minimos']) || 1, multiplo: num(meta['multiplo']) || 1,
      flujos: flujos,
      precioMonitor: compra ? compra / base : null,           // por 1 VN, CON 0,5% de comisión
      tirMonitor: typeof meta['tir efectiva'] === 'number' ? meta['tir efectiva'] : null,
      liqMonitor: liqFile, validado: true, motivo: ''
    };

    // ── Validación: (1) la amortización futura debe sumar el valor residual; (2) los flujos deben
    //    reproducir la TIR del Monitor al precio del Monitor.
    var res = num(meta['valor residual']), amortFut = 0, kk;
    for (kk = 0; kk < flujos.length; kk++) if (!liqFile || flujos[kk].fecha > liqFile) amortFut += flujos[kk].amort;
    if (res > 0 && Math.abs(amortFut * base - res) > Math.max(0.5, res * 0.005)) { bono.validado = false; bono.motivo = 'la amortización futura no suma el valor residual'; }
    if (bono.validado && compra && bono.tirMonitor !== null && liqFile) {
      var cfs = [{ t: 0, v: -compra / base }];
      flujos.forEach(function (f) { if (f.fecha > liqFile) cfs.push({ t: diasEntre(liqFile, f.fecha) / 365, v: f.total }); });
      var t = xirr(cfs);
      bono.tirCalc = t;
      if (t === null || Math.abs(t - bono.tirMonitor) > TOL_TIR) { bono.validado = false; bono.motivo = 'la TIR recalculada no coincide con la del Monitor'; }
    }
    return bono;
  }

  // Nombres de hoja que pueden ser un bono (todas menos las de resumen y las letras en pesos S…/T…).
  function nombresBonos(sheetNames) {
    return (sheetNames || []).filter(function (n) { return !NO_INSTRUMENTO[n] && !L.RE_LETRA.test(n); });
  }

  // Lee un libro de SheetJS y devuelve { bonos:[...], descartados:[{ticker,motivo}] }. Solo bonos vigentes.
  function bonosDesdeLibro(XLSX, wb, hoy) {
    hoy = hoy || new Date();
    var bonos = [], desc = [];
    nombresBonos(wb.SheetNames).forEach(function (n) {
      var ws = wb.Sheets[n]; if (!ws) return;
      var b = parseHoja(n, XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }));
      if (!b) return;
      if (b.descartado) { desc.push({ ticker: n, motivo: b.descartado }); return; }
      if (diasEntre(hoy, b.vto) <= 0) return;                  // ya venció
      bonos.push(b);
    });
    bonos.sort(function (a, b) { return a.vto - b.vto || (a.ticker < b.ticker ? -1 : 1); });
    return { bonos: bonos, descartados: desc };
  }

  // Versión compacta para publicar (data/cuanto_cobro.json). Ver nota en fcn-letras.js.
  function r8(v) { return Math.round(v * 1e8) / 1e8; }
  function iso(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function bonosAJSON(bonos) {
    return bonos.map(function (b) {
      return { t: b.ticker, s: b.simbolo, em: b.emisor, c: b.cat, ca: b.calif, l: b.ley, co: b.cobra, tt: b.tipoTasa, cu: b.cuponAnual,
        fr: b.frecuencia, at: b.amortTxt, mc: b.mesesCupon, v: iso(b.vto), mn: b.minNom, mu: b.multiplo,
        f: b.flujos.map(function (f) { return [iso(f.fecha), r8(f.cupon), r8(f.amort), r8(f.total)]; }),
        pm: b.precioMonitor, lm: b.liqMonitor ? iso(b.liqMonitor) : null, tm: b.tirMonitor, ok: b.validado ? 1 : 0, mo: b.motivo };
    });
  }
  function bonosDesdeJSON(arr, hoy) {
    hoy = hoy || new Date();
    var out = (arr || []).map(function (x) {
      return { ticker: x.t, simbolo: x.s, emisor: x.em, cat: x.c, calif: x.ca, ley: x.l, cobra: x.co, tipoTasa: x.tt, cuponAnual: x.cu,
        frecuencia: x.fr, amortTxt: x.at, mesesCupon: x.mc, vto: fechaCelda(x.v), minNom: x.mn, multiplo: x.mu,
        flujos: x.f.map(function (f) { return { fecha: fechaCelda(f[0]), cupon: f[1], amort: f[2], total: f[3] }; }),
        precioMonitor: x.pm, liqMonitor: x.lm ? fechaCelda(x.lm) : null, tirMonitor: x.tm, validado: !!x.ok, motivo: x.mo || '' };
    }).filter(function (b) { return b.vto && diasEntre(hoy, b.vto) > 0; });
    out.sort(function (a, b) { return a.vto - b.vto || (a.ticker < b.ticker ? -1 : 1); });
    return out;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // 2) Precios en vivo (BYMA)
  // ─────────────────────────────────────────────────────────────────────────
  function post(ep, body, force) { return U.bymaPost(ep, body, force); }         // comparte la descarga (y su memoria de 60 s) con fcn-letras.js
  function filas(p) { return Array.isArray(p) ? p : ((p && p.data) || []); }

  // Símbolos BYMA candidatos para un bono, en orden de preferencia (dólar MEP primero).
  function candidatos(b) {
    var out = [], t = b.ticker, stem = t.slice(0, 4);
    function add(s) { if (s && out.indexOf(s) < 0) out.push(s); }
    add(b.simbolo); add(t + 'D'); add(stem + 'D'); add(b.simbolo.replace(/D$/, 'C')); add(t + 'C'); add(stem + 'C');
    return out;
  }

  // { SIMBOLO: { ci:{px,ult,bid,offer,prev,vol,hora}, h24:{…} } } — solo de los símbolos pedidos
  function agrupar(rows, wanted) {
    var out = {};
    rows.forEach(function (x) {
      if (!wanted[x.symbol]) return;
      var pata = x.settlementType === '1' ? 'ci' : (x.settlementType === '2' ? 'h24' : null);
      if (!pata) return;
      if (x.denominationCcy === 'ARS') return;               // la punta en pesos (símbolo sin D/C) cotiza en otra unidad: nunca es la que buscamos
      var ult = num(x.trade) || num(x.closingPrice) || num(x.previousClosingPrice);
      var offer = num(x.offerPrice), bid = num(x.bidPrice);
      var px = (offer > 0 && (!ult || Math.abs(offer / ult - 1) <= 0.03)) ? offer : ult;
      if (!(px > 0)) return;
      (out[x.symbol] = out[x.symbol] || {})[pata] = { px: px, ult: ult, bid: bid, offer: offer, prev: num(x.previousClosingPrice), vol: num(x.volume), hora: x.tradeHour || null };
    });
    return out;
  }
  function leerCache() { try { return JSON.parse(localStorage.getItem(LS_KEY) || 'null'); } catch (e) { return null; } }
  function guardarCache(o) { try { localStorage.setItem(LS_KEY, JSON.stringify(o)); } catch (e) {} }

  var _inflight = null;
  // Promise<{ok, ts, stale, data}> — nunca rechaza. `bonos` = lista de bonos ya leída (define qué símbolos se guardan).
  function getPrecios(bonos, opts) {
    opts = opts || {};
    var c = leerCache();
    if (!opts.force && c && c.ts && (Date.now() - c.ts) < TTL_MS) return Promise.resolve({ ok: true, ts: c.ts, stale: false, data: c.data });
    if (_inflight) return _inflight;
    var wanted = {};
    bonos.forEach(function (b) { candidatos(b).forEach(function (s) { wanted[s] = 1; }); });
    _inflight = Promise.all([post('public-bonds', '{"page_size":2000}', opts.force), post('negociable-obligations', '{}', opts.force)])
      .then(function (res) {
        var data = agrupar(filas(res[0]).concat(filas(res[1])), wanted);
        if (!Object.keys(data).length) throw new Error('BYMA sin datos de bonos');
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

  // Precio a usar (cada 100 VN, en USD). Prefiere el símbolo …D; valida contra el precio del Monitor.
  // → { px, fuente:'BYMA'|'Monitor', simbolo, hora, variacion, sinOperar, aviso } o null
  function elegirPrecio(b, preciosBYMA, plazo, hoy) {
    var pref = plazo === 'CI' ? 'ci' : 'h24', alt = plazo === 'CI' ? 'h24' : 'ci';
    var refMon = b.precioMonitor ? b.precioMonitor / (1 + COM_MONITOR) * 100 : null;   // precio Monitor sin comisión, cada 100 VN
    var cands = candidatos(b), i, aviso = '';
    for (i = 0; i < cands.length; i++) {
      var d = preciosBYMA && preciosBYMA[cands[i]];
      var p = d && (d[pref] || d[alt]);
      if (!p) continue;
      if (!refMon && (p.px < 0.5 || p.px > 250)) continue;        // sin precio de referencia: descartar cotizaciones que no pueden ser USD cada 100 VN
      if (refMon && Math.abs(p.px / refMon - 1) > TOL_PRECIO) { aviso = 'BYMA difiere más de ' + Math.round(TOL_PRECIO * 100) + '% del Monitor'; continue; }
      return { px: p.px, fuente: 'BYMA', simbolo: cands[i], hora: p.hora, variacion: p.prev > 0 ? p.px / p.prev - 1 : null, sinOperar: !(p.vol > 0) };
    }
    // El precio del Monitor es de SU fecha. Si entre esa fecha y la liquidación de hoy cayó un pago, ese precio todavía
    // incluía el cobro (cotiza "con cupón") y usarlo daría una TIR absurda: en ese caso no hay precio confiable.
    var liqHoy = L.fechaLiquidacion(hoy || new Date(), plazo);
    if (refMon && b.liqMonitor && b.flujos.some(function (f) { return f.fecha > b.liqMonitor && f.fecha <= liqHoy; })) return null;
    if (refMon) return { px: refMon, fuente: 'Monitor', simbolo: null, hora: null, variacion: null, sinOperar: false, aviso: aviso };
    return null;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // 3) Cálculo
  // ─────────────────────────────────────────────────────────────────────────
  // o = { monto (USD), comisionPct, plazo:'CI'|'24', hoy:Date, precio:{px,…} }
  function calcularBono(b, o) {
    var pr = o.precio;
    if (!pr || !(pr.px > 0)) return null;
    var hoy = o.hoy || new Date();
    var liq = L.fechaLiquidacion(hoy, o.plazo);
    var fut = b.flujos.filter(function (f) { return f.fecha > liq; });     // cobros posteriores a la liquidación
    if (!fut.length) return null;
    var com = (num(o.comisionPct) || 0) / 100;
    var pCom = pr.px / 100 * (1 + com);                                     // por 1 VN, con comisión
    var cfs = [{ t: 0, v: -pCom }];
    fut.forEach(function (f) { cfs.push({ t: diasEntre(liq, f.fecha) / 365, v: f.total }); });
    var r = {
      ticker: b.ticker, cat: b.cat, emisor: b.emisor, calif: b.calif, cobra: b.cobra, vto: b.vto,
      liquidacion: liq, precio: pr.px / 100, precioCom: pCom, fuente: pr.fuente, simbolo: pr.simbolo, hora: pr.hora,
      variacion: pr.variacion, sinOperar: pr.sinOperar, aviso: pr.aviso || '',
      tir: xirr(cfs), dias: diasEntre(liq, b.vto), proximo: fut[0], minNom: b.minNom, multiplo: b.multiplo
    };
    var totPorVN = 0, hasta = new Date(liq.getTime()); hasta.setFullYear(hasta.getFullYear() + 1);
    var monto = num(o.monto);
    if (monto > 0) {
      var vn = Math.floor(monto / pCom / b.multiplo) * b.multiplo;          // lotes enteros
      r.vn = vn;
      r.minimoUSD = b.minNom * pCom;                                        // lo mínimo para poder comprar
      if (vn < b.minNom) { r.vn = 0; r.bajoMinimo = true; return r; }
      r.costo = vn * pCom;
      r.calendario = fut.map(function (f) { return { fecha: f.fecha, cupon: f.cupon * vn, amort: f.amort * vn, total: f.total * vn }; });
      var c12 = 0, i12 = 0, a12 = 0, cT = 0, iT = 0, aT = 0;
      r.calendario.forEach(function (x) {
        cT += x.total; iT += x.cupon; aT += x.amort;
        if (x.fecha <= hasta) { c12 += x.total; i12 += x.cupon; a12 += x.amort; }
      });
      r.cobro12 = c12; r.renta12 = i12; r.capital12 = a12;
      r.cobroTotal = cT; r.rentaTotal = iT; r.capitalTotal = aT;
      r.ganancia = cT - r.costo;
      r.proximoUSD = r.calendario[0];
      r.rentaCorriente = r.costo > 0 ? i12 / r.costo : null;               // intereses de los próximos 12 meses / lo invertido
    }
    return r;
  }

  // Cobros de los próximos 12 meses agrupados por mes (para la tira de barras del detalle).
  function flujoMensual(calendario, liq, meses) {
    meses = meses || 12;
    var out = [], i;
    for (i = 0; i < meses; i++) {
      var d = new Date(liq.getFullYear(), liq.getMonth() + i, 1);
      out.push({ anio: d.getFullYear(), mes: d.getMonth(), total: 0, cupon: 0, amort: 0 });
    }
    (calendario || []).forEach(function (x) {
      for (var k = 0; k < out.length; k++) {
        if (out[k].anio === x.fecha.getFullYear() && out[k].mes === x.fecha.getMonth()) { out[k].total += x.total; out[k].cupon += x.cupon; out[k].amort += x.amort; break; }
      }
    });
    return out;
  }

  global.FCNBonos = {
    xirr: xirr, categoria: categoria, parseHoja: parseHoja, nombresBonos: nombresBonos, bonosDesdeLibro: bonosDesdeLibro,
    bonosAJSON: bonosAJSON, bonosDesdeJSON: bonosDesdeJSON,
    candidatos: candidatos, agrupar: agrupar, getPrecios: getPrecios, elegirPrecio: elegirPrecio,
    calcularBono: calcularBono, flujoMensual: flujoMensual
  };
})(typeof window !== 'undefined' ? window : globalThis);
