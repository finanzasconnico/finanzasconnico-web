/* ═══════════════════════════════════════════════════════════════════════════
   fcn-vivo.js — precios EN VIVO para las herramientas de la página (analisis.html).
   Va DESPUÉS de fcn-letras.js y fcn-bonos.js (usa sus motores: cronogramas del Monitor, precios de
   BYMA con respaldo en data912, y el mismo cálculo de TIR que ¿Cuánto cobro?).

   Por qué existe: el Monitor (monitor.xlsx) NO siempre se refresca (verificado 17/11/2026: el archivo se
   republicaba a diario pero su contenido era del 15/09, dos meses viejo). Con esos precios un cliente podía
   decidir mirando números de hace dos meses. Entonces:
     - De cada bono se toma del Monitor SOLO lo que no cambia con el tiempo: el calendario de cobros.
     - El precio sale del mercado (BYMA o, si no responde, data912) y la TIR y la duration se RECALCULAN con
       ese precio, la misma comisión de 0,5% que usa el Monitor y la fecha de liquidación de hoy.
     - Lo que NO se puede recalcular (bonos en pesos CER / TAMAR / Dual: su TIR depende de proyecciones de
       inflación y tasas) no se muestra si el Monitor está viejo, en vez de mostrar un número de hace meses.
     - Si no hay precio en vivo NI un Monitor reciente, no se arma ninguna recomendación y se dice por qué.
   Todo es de solo lectura: nunca opera ni manda nada.
   ═══════════════════════════════════════════════════════════════════════════ */
(function (global) {
  'use strict';
  var L = global.FCNLetras, B = global.FCNBonos;
  if (!L || !B) throw new Error('fcn-vivo.js necesita fcn-letras.js y fcn-bonos.js cargados antes');
  var U = L.util, diasEntre = L.diasEntre;

  var COM_BONO = 0.5;      // % — el precio del Monitor para bonos ya trae 0,5% de comisión
  var COM_LETRA = 0.1;     // % — y 0,1% para letras
  var PLAZO = '24';        // 24 hs: el plazo donde más se opera
  var MIN_DIAS_TIR = 30;   // a menos de 30 días del vencimiento la TIR anualizada no dice nada (±80% por un centavo de precio)
  var MAX_CACHE_VIEJA_MS = 30 * 60 * 1000;

  var est = null, prom = null;

  function r2(x) { return Math.round(x * 100) / 100; }
  function hoyFecha() { var h = new Date(); return new Date(h.getFullYear(), h.getMonth(), h.getDate()); }
  function p2(n) { return (n < 10 ? '0' : '') + n; }
  function fDM(d) { return d ? p2(d.getDate()) + '/' + p2(d.getMonth() + 1) : '—'; }
  function hhmm(ts) {
    var b = L.ahoraBA(new Date(ts)), h = Math.floor(b.hora), m = Math.round((b.hora - h) * 60);
    if (m === 60) { h += 1; m = 0; }
    return p2(h) + ':' + p2(m);
  }

  // Duration de Macaulay (años) con la TIR vivo — es la convención del Monitor (verificado contra la celda "Duration" de cada hoja).
  function duracion(b, liq, tir) {
    var pv = 0, mac = 0;
    b.flujos.forEach(function (f) {
      if (f.fecha <= liq) return;
      var t = diasEntre(liq, f.fecha) / 365, v = f.total / Math.pow(1 + tir, t);
      pv += v; mac += t * v;
    });
    return pv > 0 ? mac / pv : null;
  }

  // ── Bonos CER de pago único (TZXD6, TZXM7, X29E7…): del Monitor solo se toma ticker y vencimiento; el precio sale de data912.
  // Pagan capital ajustado por inflación al vencimiento, sin cupones. Su cobro real depende de la inflación que haya, así que
  // en cerComoLetras() se ESTIMA con lo que hoy descuenta el mercado (ver ahí). No se usa el cobro proyectado de la hoja
  // del Monitor porque es de la fecha del Monitor: con el precio de hoy daría ganancias sin sentido.
  var RE_HOJA_CER = /^(TZX\w+|TX\d+|X\d{2}[A-Z]\d)$/;
  function cerDesdeLibro(XLSX, wb, hoy) {
    var out = [];
    (wb.SheetNames || []).forEach(function (n) {
      if (!RE_HOJA_CER.test(n) || !wb.Sheets[n]) return;
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }), tipo = null, interes = null, vto = null;
      rows.forEach(function (r) {
        for (var j = 0; r && j < r.length - 1; j++) {
          var k = U.noAcc(r[j]);
          if (k === 'tipo de tasa' && tipo === null) tipo = U.noAcc(r[j + 1]);
          else if (k === 'interes anual' && interes === null) interes = r[j + 1];
          else if (k === 'fecha de vencimiento' && vto === null) vto = U.fechaCelda(r[j + 1]);
        }
      });
      if (tipo === 'cer' && interes === 0 && vto && diasEntre(hoy, vto) > 0) out.push({ ticker: n, vto: vto });
    });
    return out;
  }
  function cerPrecios(lista) {
    if (!lista.length) return Promise.resolve({});
    return Promise.all([U.d912Get('arg_bonds'), U.d912Get('arg_notes')]).then(function (res) {
      var by = {}, rueda = L.enRueda(new Date());
      U.d912Filas(res[0]).concat(U.d912Filas(res[1])).forEach(function (x) { by[x.symbol] = x; });
      var out = {};
      lista.forEach(function (b) {
        var x = by[b.ticker]; if (!x) return;
        var ult = x.trade, offer = x.offerPrice;
        var px = (rueda && offer > 0 && (!ult || Math.abs(offer / ult - 1) <= 0.03)) ? offer : ult;
        if (px > 0) out[b.ticker] = px;
      });
      return out;
    }).catch(function () { return {}; });
  }

  function armar(XLSX, wb, opts) {
    var hoy = hoyFecha();
    var bonos = B.bonosDesdeLibro(XLSX, wb, hoy).bonos;
    var letras = L.letrasDesdeLibro(XLSX, wb, hoy);
    var cerBase = cerDesdeLibro(XLSX, wb, hoy);
    return Promise.all([B.getPrecios(bonos, opts), L.getPrecios(opts), cerPrecios(cerBase)]).then(function (res) {
      var rb = res[0], rl = res[1], i;

      // ── antigüedad real del Monitor: la fecha de liquidación guardada en cada hoja (las celdas solo dicen "Updated at 17:11")
      var asOf = null;
      bonos.forEach(function (b) { if (b.liqMonitor && (!asOf || b.liqMonitor > asOf)) asOf = b.liqMonitor; });
      letras.forEach(function (l) { if (l.fechaPrecio && (!asOf || l.fechaPrecio > asOf)) asOf = l.fechaPrecio; });
      var edad = asOf ? diasEntre(asOf, hoy) : null;
      var monitorViejo = edad === null || edad > L.MAX_EDAD_MONITOR;

      var bonosOk = rb.ok && (!rb.stale || (rb.ts && Date.now() - rb.ts < MAX_CACHE_VIEJA_MS));
      var letrasOk = rl.ok && (!rl.stale || (rl.ts && Date.now() - rl.ts < MAX_CACHE_VIEJA_MS));
      var out = { bonos: {}, letras: {}, nBonosTotal: 0, nBonosVivo: 0, nLetrasTotal: letras.length, nLetrasVivo: 0 };

      // ── bonos / ONs en dólares
      bonos.forEach(function (b) {
        if (!b.validado) return;                         // sus flujos no reproducen la TIR del Monitor: no se usan para nada
        out.nBonosTotal++;
        var pr = B.elegirPrecio(b, bonosOk ? rb.data : {}, PLAZO, new Date());
        if (!pr) return;
        var c = B.calcularBono(b, { monto: 0, comisionPct: COM_BONO, plazo: PLAZO, hoy: new Date(), precio: pr });
        if (!c) return;
        var vivo = pr.fuente === 'BYMA';
        var tirOk = c.tir !== null && isFinite(c.tir) && c.dias >= MIN_DIAS_TIR;
        out.bonos[b.ticker] = {
          vivo: vivo, fuente: pr.fuente, via: pr.via || null, simbolo: pr.simbolo, hora: pr.hora, sinOperar: !!pr.sinOperar,
          px: pr.px, precioCom: c.precioCom, dias: c.dias,
          tir: tirOk ? c.tir * 100 : null,
          duration: tirOk ? duracion(b, c.liquidacion, c.tir) : null
        };
        if (vivo) out.nBonosVivo++;
      });

      // ── letras en pesos (LECAP / BONCAP)
      letras.forEach(function (l) {
        var pr = L.elegirPrecio(l, letrasOk ? rl.data : {}, PLAZO);
        if (!pr) return;
        var c = L.calcular(l, { monto: 0, comisionPct: COM_LETRA, plazo: PLAZO, hoy: new Date(), precio: pr });
        if (!c) return;
        var vivo = pr.fuente === 'BYMA';
        out.letras[l.ticker] = { vivo: vivo, fuente: pr.fuente, via: pr.via || null, hora: pr.hora, px: pr.px, precioCom: c.precioCom, tna: c.tna, tea: c.tea, dias: c.dias };
        if (vivo) out.nLetrasVivo++;
      });

      // ── bonos CER de pago único (solo con curva de letras vivas: de ahí sale la estimación)
      var liqHoy = L.fechaLiquidacion(new Date(), PLAZO);
      out.cer = {};
      cerBase.forEach(function (b) {
        var px = res[2][b.ticker], dias = diasEntre(liqHoy, b.vto);
        if (px > 0 && dias > 0) out.cer[b.ticker] = { px: px, vto: b.vto, dias: dias };
      });

      var ts = Math.max(bonosOk && rb.ts ? rb.ts : 0, letrasOk && rl.ts ? rl.ts : 0) || null;
      var via = (rb.ok && rb.fuente) || (rl.ok && rl.fuente) || null;
      est = {
        bonos: out.bonos, letras: out.letras, cer: out.cer,
        bonosVivo: out.nBonosVivo > 0, letrasVivo: out.nLetrasVivo > 0,
        nBonosTotal: out.nBonosTotal, nBonosVivo: out.nBonosVivo, nLetrasTotal: out.nLetrasTotal, nLetrasVivo: out.nLetrasVivo,
        fuente: via, ts: ts, enRueda: L.enRueda(new Date()),
        asOf: asOf, edadMonitor: edad, monitorViejo: monitorViejo,
        error: (!rb.ok && rb.error) || (!rl.ok && rl.error) || null
      };
      return est;
    });
  }

  // Promise<estado>. Nunca rechaza. Se arma una sola vez (los motores ya cachean los precios 90 s).
  function cargar(XLSX, wb, opts) {
    opts = opts || {};
    if (prom && !opts.force) return prom;
    prom = armar(XLSX, wb, opts).catch(function (e) {
      est = { bonos: {}, letras: {}, bonosVivo: false, letrasVivo: false, nBonosTotal: 0, nBonosVivo: 0, nLetrasTotal: 0, nLetrasVivo: 0,
        fuente: null, ts: null, enRueda: false, asOf: null, edadMonitor: null, monitorViejo: true, error: String(e && e.message || e) };
      return est;
    });
    return prom;
  }
  function estado() { return est; }

  // ── Aplicar a las tablas de datos de la página ───────────────────────────────────────────────────────────
  // BOND_DB / FLUJO_DB: pisa precio, TIR y duration con los vivos. Un bono sin precio de hoy NO se deja con el
  // precio viejo: sale de las bases (se avisa). Devuelve { bloqueado, quitados:[...] }.
  function aplicarDB(BOND_DB, FLUJO_DB) {
    var quitados = [];
    if (!est) return { bloqueado: false, quitados: quitados };
    if (!est.bonosVivo) {
      if (!est.monitorViejo) return { bloqueado: false, quitados: quitados };   // sin vivo, pero el Monitor es reciente: se usa tal cual
      Object.keys(BOND_DB).forEach(function (t) { delete BOND_DB[t]; });
      Object.keys(FLUJO_DB).forEach(function (t) { delete FLUJO_DB[t]; });
      return { bloqueado: true, quitados: quitados };
    }
    Object.keys(BOND_DB).forEach(function (tk) {
      var e = est.bonos[tk], b = BOND_DB[tk], f = FLUJO_DB[tk];
      if (!e) { quitados.push(tk); delete BOND_DB[tk]; delete FLUJO_DB[tk]; return; }
      b.precioMep = e.precioCom; b.precioCable = null;
      b.tirMep = e.tir === null ? null : r2(e.tir); b.tirCable = null;
      b.duration = e.duration; b.vivo = e.vivo; b.fuentePrecio = e.fuente; b.sinOperar = e.sinOperar;
      if (f) { f.tirEfectiva = e.tir === null ? null : r2(e.tir); f.duration = e.duration; }
    });
    Object.keys(FLUJO_DB).forEach(function (tk) { if (!est.bonos[tk]) delete FLUJO_DB[tk]; });   // hojas sin ficha en Corporativos/Soberanos y sin precio
    return { bloqueado: false, quitados: quitados };
  }

  // Fila de las tablas "Monitor de bonos" (ONs y Soberanos). null = no mostrar.
  function aplicarFilaBono(o) {
    if (!est) return o;
    if (!est.bonosVivo) return est.monitorViejo ? null : o;
    var e = est.bonos[o.ticker];
    if (!e) return null;
    return Object.assign({}, o, {
      precio: e.precioCom, tir: e.tir === null ? null : e.tir / 100,
      duration: e.duration === null ? NaN : e.duration, vivo: e.vivo, sinOperar: e.sinOperar
    });
  }

  // Letras (Renta mensual en pesos): precioMonitor pasa a ser el precio vivo (por 1 VN, sin comisión, igual que antes).
  function aplicarLetras(lista) {
    if (!est) return lista;
    if (!est.letrasVivo) return est.monitorViejo ? [] : lista;
    return lista.filter(function (l) {
      var e = est.letras[l.ticker];
      if (!e) return false;
      l.precioMonitor = e.px / 100; l.vivo = e.vivo; l.tnaVivo = e.tna; l.diasVivo = e.dias;
      return true;
    });
  }

  // Bonos CER de pago único, con el mismo formato que una letra (para la escalera de Renta mensual en pesos), con el cobro
  // ESTIMADO: precio de hoy × (1 + TEA)^(años), donde TEA es el rendimiento que hoy dan las LECAPs/BONCAP de plazo parecido
  // (interpolado entre las más cercanas; más allá de la última se usa la última). Es decir: "si la inflación termina siendo la
  // que hoy descuenta el mercado". Si es mayor, el CER paga más; si es menor, menos.
  function cerComoLetras() {
    if (!est || !est.letrasVivo || !est.cer) return [];
    var curva = Object.keys(est.letras).map(function (t) { return { d: est.letras[t].dias, tea: est.letras[t].tea }; })
      .filter(function (p) { return p.d > 0 && isFinite(p.tea); }).sort(function (a, b) { return a.d - b.d; });
    if (!curva.length) return [];
    function tea(d) {
      if (d <= curva[0].d) return curva[0].tea;
      for (var i = 1; i < curva.length; i++) {
        if (d <= curva[i].d) { var a = curva[i - 1], b = curva[i]; return a.tea + (b.tea - a.tea) * (d - a.d) / (b.d - a.d); }
      }
      return curva[curva.length - 1].tea;
    }
    return Object.keys(est.cer).map(function (t) {
      var c = est.cer[t], p = c.px / 100, te = tea(c.dias);
      return { ticker: t, tipo: 'Bono CER', vto: c.vto, vf: p * Math.pow(1 + te, c.dias / 365), precioMonitor: p, vivo: true, diasVivo: c.dias, estimado: true, teaEstimada: te };
    });
  }

  // Fila de la tabla de bonos en pesos. LECAP/BONCAP: recalculada en vivo. CER/TAMAR/Dual/otras: no se puede
  // recalcular → si el Monitor está viejo se marca sinVivo (la tabla muestra "—" y lo explica).
  function aplicarFilaPesos(b) {
    if (!est) return b;
    var e = L.RE_LETRA.test(b.ticker) ? est.letras[b.ticker] : null;
    if (e && est.letrasVivo) {
      return Object.assign({}, b, { precio: e.precioCom, tir: e.tea, tna: e.tna, duration: e.dias / 365, vivo: e.vivo });
    }
    if (est.monitorViejo) return Object.assign({}, b, { sinVivo: true, precio: null, tir: null, tna: NaN, duration: null });
    return b;
  }

  // ── Cartel de estado ──────────────────────────────────────────────────────────────────────────────────────
  // { nivel: 'ok' | 'cerrado' | 'warn' | 'bad', html }
  function cartel() {
    if (!est) return { nivel: 'info', html: 'Trayendo precios del mercado…' };
    var vivo = est.bonosVivo || est.letrasVivo;
    var origen = est.fuente === 'data912' ? 'data912.com (BYMA no respondió)' : 'BYMA';
    // Solo una parte con precio de hoy: se dice cuál, para no mostrar "en vivo" a secas cuando faltan los bonos
    if (est.letrasVivo && !est.bonosVivo) return { nivel: 'warn', html: '⚠️ <strong>Solo pudimos traer los precios de las letras en pesos.</strong> Los de bonos y ONs en dólares no están disponibles ahora' + (est.monitorViejo ? ' y no mostramos los guardados porque son del ' + fDM(est.asOf) + ' (hace ' + est.edadMonitor + ' días)' : '') + '. Probá de nuevo en unos minutos.' };
    if (est.bonosVivo && !est.letrasVivo) return { nivel: 'warn', html: '⚠️ <strong>Solo pudimos traer los precios de bonos y ONs en dólares.</strong> Los de las letras en pesos no están disponibles ahora' + (est.monitorViejo ? ' y no mostramos los guardados porque son del ' + fDM(est.asOf) + ' (hace ' + est.edadMonitor + ' días)' : '') + '. Probá de nuevo en unos minutos.' };
    if (vivo) {
      // Con la página abierta mucho rato los precios cargados se van quedando viejos
      var min = est.ts ? Math.floor((Date.now() - est.ts) / 60000) : 0;
      if (min >= 20) return { nivel: 'warn', html: '⏱️ <strong>Estos precios se cargaron hace ' + min + ' minutos</strong> (a las ' + hhmm(est.ts) + ' hs) y pueden haber cambiado. Actualizá la página para ver los de ahora.' };
      if (est.enRueda) return { nivel: 'ok', html: '🟢 <strong>Precios en vivo</strong> · actualizados a las ' + hhmm(est.ts) + ' hs · fuente: ' + origen };
      return { nivel: 'cerrado', html: '🌙 <strong>Mercado cerrado</strong> · se muestran los últimos precios operados (cargados a las ' + hhmm(est.ts) + ' hs · fuente: ' + origen + '). Se renuevan en la próxima rueda (11 a 17 hs).' };
    }
    if (est.monitorViejo) {
      return { nivel: 'bad', html: '⛔ <strong>No pudimos traer los precios en vivo</strong> y los últimos guardados son del ' + fDM(est.asOf) + (est.edadMonitor != null ? ' (hace ' + est.edadMonitor + ' días)' : '') +
        '. Por seguridad no mostramos precios ni armamos recomendaciones con datos viejos. Probá de nuevo en unos minutos.' };
    }
    return { nivel: 'warn', html: '⚠️ <strong>Sin conexión con el mercado</strong>: se muestran los últimos precios guardados (' + fDM(est.asOf) + '). Verificalos antes de operar.' };
  }

  global.PreciosVivo = {
    cargar: cargar, estado: estado, cartel: cartel,
    aplicarDB: aplicarDB, aplicarFilaBono: aplicarFilaBono, aplicarLetras: aplicarLetras, cerComoLetras: cerComoLetras, aplicarFilaPesos: aplicarFilaPesos,
    COM_BONO: COM_BONO, COM_LETRA: COM_LETRA, PLAZO: PLAZO
  };
})(typeof window !== 'undefined' ? window : globalThis);
