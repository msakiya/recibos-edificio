/*
 * calculo.js — Cálculo de los recibos del edificio (funciones puras, sin internet).
 * Es el MISMO cálculo validado en la hoja de Google (setiembre 2026: dpto 101 = S/ 471.99).
 * Funciona en el navegador (window.Calculo) y en node (module.exports) para las pruebas.
 */
(function (global) {
'use strict';
var MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
             'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre'];

/* =====================================================================
 *  FUNCIONES PURAS (sin servicios de Google; se prueban fuera de Apps Script)
 * ===================================================================== */

/** Redondeo a n decimales "como Sheets" (mitad hacia arriba, sin errores binarios). */
function redondear(x, n) {
  n = (n === undefined) ? 2 : n;
  x = Number(x) || 0;
  var f = Math.pow(10, n);
  var s = x < 0 ? -1 : 1;
  return s * Math.round(Number((Math.abs(x) * f).toPrecision(15))) / f;
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }

/** Normaliza un mes a 'AAAA-MM'. Acepta Date, '2026-9', '2026-09', '09/2026', '9-2026'. */
function normalizarMes(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + '-' + pad2(v.getMonth() + 1);
  }
  var s = String(v).trim(), m;
  if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.]\d{1,2})?$/))) return m[1] + '-' + pad2(Number(m[2]));
  if ((m = s.match(/^(\d{1,2})[-\/.](\d{4})$/))) return m[2] + '-' + pad2(Number(m[1]));
  return s;
}

function partesMes(mes) {
  var p = normalizarMes(mes).split('-');
  return { anio: Number(p[0]), mes: Number(p[1]) };
}

function mesDesplazado(mes, delta) {
  var p = partesMes(mes);
  var idx = p.anio * 12 + (p.mes - 1) + delta;
  return Math.floor(idx / 12) + '-' + pad2((idx % 12) + 1);
}
function mesAnterior(mes) { return mesDesplazado(mes, -1); }
function mesSiguiente(mes) { return mesDesplazado(mes, 1); }

function nombreMes(mes) { return MESES[partesMes(mes).mes - 1]; }

/** Nombre de la carpeta del mes, igual que en Drive: '9-Setiembre', '10-Octubre'. */
function nombreCarpetaMes(mes) { var p = partesMes(mes); return p.mes + '-' + MESES[p.mes - 1]; }

function fmtNum(x, dec) {
  x = redondear(x, dec);
  var neg = x < 0; x = Math.abs(x);
  var partes = x.toFixed(dec).split('.');
  partes[0] = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg ? '-' : '') + partes.join('.');
}
function fmtSoles(x) { var v = redondear(x, 2); return (v < 0 ? '-S/ ' : 'S/ ') + fmtNum(Math.abs(v), 2); }

function aFecha(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var s = String(v).trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return null;
}
function fmtFecha(d) { return d ? pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear() : ''; }

function esSi(v) { return /^s[ií]?$|^si$|^sí$|^true$|^1$|^x$/i.test(String(v || '').trim()); }
function num(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function mismaUnidad(a, b) { return String(a).trim() === String(b).trim(); }

/**
 * Cálculo completo de un mes.
 * datos = {mes, config, deptos[], lecturas[], servicios[], gastos[], ajustes[], pagos[], fondo[]}
 * Devuelve totales del edificio y el detalle de cada unidad.
 */
function calcularMes(datos) {
  var mes = normalizarMes(datos.mes);
  var cfg = datos.config || {};
  var errores = [];
  var serv = (datos.servicios || []).filter(function (s) { return normalizarMes(s.mes) === mes; })[0];
  if (!serv) { errores.push('No hay fila del mes ' + mes + ' en la pestaña Servicios.'); serv = {}; }
  var nDptos = num(cfg.NUM_DPTOS) || datos.deptos.length;
  var aguaImporte = num(serv.aguaImporte), aguaM3 = num(serv.aguaM3);
  var precioM3 = aguaM3 ? aguaImporte / aguaM3 : 0;

  // --- Agua privada por unidad
  var unidades = datos.deptos.map(function (d) {
    var l = (datos.lecturas || []).filter(function (x) {
      return normalizarMes(x.mes) === mes && mismaUnidad(x.unidad, d.unidad);
    })[0] || {};
    var ant = num(l.anterior), act = num(l.actual);
    var tieneLectura = l.actual !== undefined && l.actual !== null && l.actual !== '';
    var consumo = tieneLectura ? redondear(act - ant, 3) : 0;
    return {
      unidad: d.unidad, propietario: d.propietario || '', telefono: d.telefono || '', correo: d.correo || '',
      anterior: ant, actual: act, tieneLectura: tieneLectura, consumo: consumo,
      aguaPrivada: redondear(consumo * precioM3, 2), ajusteRedondeo: 0
    };
  });
  var consumoDptos = redondear(unidades.reduce(function (a, u) { return a + u.consumo; }, 0), 3);
  var m3Comun = redondear(aguaM3 - consumoDptos, 3);
  var aguaComun = redondear(m3Comun * precioM3, 2);
  var aguaPrivadaTotal = redondear(aguaImporte - aguaComun, 2);
  var sumaRedondeada = redondear(unidades.reduce(function (a, u) { return a + u.aguaPrivada; }, 0), 2);
  var diferencia = redondear(aguaPrivadaTotal - sumaRedondeada, 2);
  var unidadAjuste = null;
  if (diferencia !== 0 && unidades.length) {
    var mayor = unidades.reduce(function (a, u) { return u.consumo > a.consumo ? u : a; }, unidades[0]);
    mayor.ajusteRedondeo = diferencia; unidadAjuste = mayor.unidad;
  }

  // --- Gastos comunes
  var gastos = (datos.gastos || []).filter(function (g) { return normalizarMes(g.mes) === mes; })
    .map(function (g) {
      var tipo = String(g.tipo || 'FIJO').toUpperCase().trim();
      var monto = num(g.monto);
      if (tipo === 'AGUA_COMUN') monto = aguaComun;
      if (tipo === 'LUZ' && (g.monto === '' || g.monto === null || g.monto === undefined)) monto = num(serv.luzImporte);
      return { num: num(g.num), concepto: String(g.concepto || ''), tipo: tipo, monto: redondear(monto, 2),
               porUnidad: redondear(monto / nDptos, 2), copiar: g.copiar };
    })
    .filter(function (g) { return g.monto !== 0; })
    .sort(function (a, b) { return a.num - b.num; });
  var comunPorUnidad = redondear(gastos.reduce(function (a, g) { return a + g.porUnidad; }, 0), 2);
  var totalFijos = redondear(gastos.filter(function (g) { return g.tipo !== 'EXTRA'; })
    .reduce(function (a, g) { return a + g.monto; }, 0), 2);
  var totalExtras = redondear(gastos.filter(function (g) { return g.tipo === 'EXTRA'; })
    .reduce(function (a, g) { return a + g.monto; }, 0), 2);
  var totalComun = redondear(totalFijos + totalExtras, 2);

  // --- Total por unidad
  var mesPrev = mesAnterior(mes);
  unidades.forEach(function (u) {
    u.ajustes = (datos.ajustes || []).filter(function (a) {
      return normalizarMes(a.mes) === mes && mismaUnidad(a.unidad, u.unidad) && num(a.monto) !== 0;
    }).map(function (a) { return { concepto: String(a.concepto || 'Ajuste'), monto: redondear(num(a.monto), 2) }; });
    u.totalAjustes = redondear(u.ajustes.reduce(function (a, x) { return a + x.monto; }, 0), 2);
    u.aguaPrivadaFinal = redondear(u.aguaPrivada + u.ajusteRedondeo, 2);
    u.comun = comunPorUnidad;
    u.total = redondear(u.comun + u.aguaPrivadaFinal + u.totalAjustes, 2);
    u.ultimoPago = redondear((datos.pagos || []).filter(function (p) {
      return normalizarMes(p.mes) === mesPrev && mismaUnidad(p.unidad, u.unidad);
    }).reduce(function (a, p) { return a + num(p.monto); }, 0), 2);
  });

  // --- Fondo de contingencia
  var aporte = redondear(gastos.filter(function (g) { return g.tipo === 'FONDO'; })
    .reduce(function (a, g) { return a + g.monto; }, 0), 2);
  var filaFondo = (datos.fondo || []).filter(function (f) { return normalizarMes(f.mes) === mes; })[0];
  var saldoAnterior;
  if (filaFondo && filaFondo.saldoAnterior !== '' && filaFondo.saldoAnterior !== undefined) {
    saldoAnterior = num(filaFondo.saldoAnterior);
  } else {
    var prev = (datos.fondo || []).filter(function (f) { return normalizarMes(f.mes) === mesPrev; })[0];
    saldoAnterior = prev ? num(prev.saldoAnterior) + num(prev.aporte) - num(prev.gastosFondo) : num(cfg.SALDO_INICIAL_FONDO);
  }
  var gastosFondo = filaFondo ? num(filaFondo.gastosFondo) : 0;
  var fondo = {
    saldoAnterior: redondear(saldoAnterior, 2), aporte: aporte, gastos: redondear(gastosFondo, 2),
    saldo: redondear(saldoAnterior + aporte - gastosFondo, 2), detalle: filaFondo ? (filaFondo.detalle || '') : ''
  };

  // --- Fechas
  var emision = aFecha(serv.emision) || new Date();
  var venc = aFecha(serv.vencimiento);
  if (!venc) { var pm = partesMes(mesSiguiente(mes)); venc = new Date(pm.anio, pm.mes - 1, num(cfg.DIA_VENCIMIENTO) || 5); }

  return {
    mes: mes, nombreMes: nombreMes(mes), anio: partesMes(mes).anio, nDptos: nDptos, errores: errores,
    emision: emision, vencimiento: venc,
    agua: { importe: aguaImporte, m3: aguaM3, precioM3: precioM3, consumoDptos: consumoDptos, m3Comun: m3Comun,
            comun: aguaComun, privadaTotal: aguaPrivadaTotal, sumaRedondeada: sumaRedondeada,
            diferenciaRedondeo: diferencia, unidadAjuste: unidadAjuste },
    luz: { importe: num(serv.luzImporte) },
    gastos: gastos, comunPorUnidad: comunPorUnidad,
    totalFijos: totalFijos, totalExtras: totalExtras, totalComun: totalComun,
    totalFacturado: redondear(totalComun + aguaPrivadaTotal, 2),
    fondo: fondo, unidades: unidades
  };
}

/** Arma el objeto que usa la plantilla Recibo.html para UNA unidad. */
function datosPlantilla(res, u, cfg, imgs) {
  imgs = imgs || {};
  var nm = res.nombreMes, nmMin = nm.toLowerCase(), nmMay = nm.toUpperCase();
  var lineas = res.gastos.map(function (g) { return { concepto: g.concepto, monto: fmtSoles(g.porUnidad) }; });
  u.ajustes.forEach(function (a) { lineas.push({ concepto: a.concepto, monto: fmtSoles(a.monto) }); });
  var textoAgua = 'Cobro Individual - Consumo de agua privada - 100.00% (Valor S/ ' + fmtNum(res.agua.precioM3, 6) +
    ') Ant. ' + fmtNum(u.anterior, 3) + ' | Act. ' + fmtNum(u.actual, 3) + ' | Cons. ' + fmtNum(u.consumo, 3);
  if (u.ajusteRedondeo) textoAgua += ' (incluye ajuste por redondeo ' + fmtSoles(u.ajusteRedondeo) + ')';
  var n1 = 0, n2 = 0;
  var etiquetaReporte = function (g) {
    return g.tipo === 'FONDO' ? '+ ' + g.concepto + ' (Mes de ' + nm + ')' : g.concepto;
  };
  return {
    edificio: cfg.EDIFICIO_NOMBRE || '', dir1: cfg.DIRECCION_LINEA_1 || '', dir2: cfg.DIRECCION_LINEA_2 || '',
    logo: imgs.logo || '',
    emision: fmtFecha(res.emision), cuotaMes: nmMin + ' - ' + res.anio, vencimiento: fmtFecha(res.vencimiento),
    ultimoPago: fmtSoles(u.ultimoPago),
    propietario: u.propietario, unidad: String(u.unidad), mesMin: nmMin,
    lineas: lineas, aguaTexto: textoAgua, aguaMonto: fmtSoles(u.aguaPrivadaFinal), total: fmtSoles(u.total),
    foto: imgs.foto || '',
    tesorero: cfg.TESORERO_NOMBRE || '', tesoreroCel: cfg.TESORERO_CELULAR || '',
    presidente: cfg.PRESIDENTE_NOMBRE || '', presidenteCel: cfg.PRESIDENTE_CELULAR || '',
    banco: cfg.BANCO || '', cuenta: cfg.CUENTA_BANCARIA || '', cci: cfg.CCI || '', nota: cfg.NOTA_PAGO || '',
    tituloReporte: nmMay + ' - ' + res.anio,
    fijos: res.gastos.filter(function (g) { return g.tipo !== 'EXTRA'; })
      .map(function (g) { return { n: ++n1, concepto: etiquetaReporte(g), monto: fmtSoles(g.monto) }; }),
    extras: res.gastos.filter(function (g) { return g.tipo === 'EXTRA'; })
      .map(function (g) { return { n: ++n2, concepto: g.concepto, monto: fmtSoles(g.monto) }; }),
    totalFijos: fmtSoles(res.totalFijos), totalExtras: fmtSoles(res.totalExtras),
    totalComun: fmtSoles(res.totalComun), totalAguaPrivada: fmtSoles(res.agua.privadaTotal),
    totalFacturado: fmtSoles(res.totalFacturado),
    fondoAnterior: fmtSoles(res.fondo.saldoAnterior), fondoAporte: fmtSoles(res.fondo.aporte),
    fondoGastos: fmtSoles(res.fondo.gastos), fondoSaldo: fmtSoles(res.fondo.saldo), fondoDetalle: res.fondo.detalle,
    mesAnio: nm + ' ' + res.anio, mesAnioMay: nmMay + ' ' + res.anio,
    empresaAgua: cfg.EMPRESA_AGUA || 'Agua', empresaLuz: cfg.EMPRESA_LUZ || 'Luz',
    imgAgua: imgs.agua || '', imgLuz: imgs.luz || ''
  };
}

function nombrePdf(res, unidad) { return 'Recibo - ' + unidad + ' - ' + res.nombreMes + ' ' + res.anio + '.pdf'; }

/** ¿El nombre de archivo corresponde a la unidad? Acepta 101, 101.jpg, 101.jpeg, 101.png, '101 - foto.jpg', 'Dpto 101'. */
function nombreEsDeUnidad(nombre, unidad) {
  var base = String(nombre).toLowerCase().trim().replace(/\.(jpe?g|png|webp|heic|heif|gif|pdf)$/i, '').trim();
  var u = String(unidad).toLowerCase().trim();
  if (base === u) return true;
  return new RegExp('(^|[^0-9])' + u + '([^0-9]|$)').test(base);
}

function textoWhatsApp(plantilla, v) {
  return String(plantilla || '').replace(/\{(\w+)\}/g, function (_, k) { return v[k] !== undefined ? v[k] : ''; });
}

/** Lista legible de unidades: "101, 102 y 201". */
function listaUnidades(arr) {
  arr = arr.map(String);
  if (arr.length <= 1) return arr.join('');
  return arr.slice(0, -1).join(', ') + ' y ' + arr[arr.length - 1];
}

/** Nombre con el que se guarda un archivo subido (es el que busca el generador). */
function nombreArchivoSubida(tipo, unidad, ext, fecha) {
  ext = String(ext || 'jpg').toLowerCase().replace(/^\./, '').replace('jpeg', 'jpg');
  if (tipo === 'medidor' || tipo === 'pago') return String(unidad).trim() + '.' + ext;
  if (tipo === 'luz' || tipo === 'agua') {
    var d = fecha || new Date();
    return 'recibo-de-' + tipo + '-' + d.getFullYear() + '.' + pad2(d.getMonth() + 1) + '.' + pad2(d.getDate()) + '.' + ext;
  }
  throw new Error('Tipo de archivo desconocido: ' + tipo);
}

function extensionDe(nombre, mime) {
  if (/pdf/i.test(mime || '') || /\.pdf$/i.test(nombre || '')) return 'pdf';
  if (/png/i.test(mime || '')) return 'png';
  if (/webp/i.test(mime || '')) return 'webp';
  return 'jpg';
}

/** Convierte lo que escribe la usuaria ("1290,107", "1,290.107", " 38.5 ") en número, o null. */
function parseNumero(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  var s = String(v).trim().replace(/\s/g, '').replace(/^S\/\.?/i, '');
  if (!s) return null;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/,/g, '');
  else s = s.replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? null : n;
}

/** Último mes (AAAA-MM) anterior a `mes` que aparece en la columna 0 de las filas. */
function ultimoMesAntes(filas, mes) {
  var m = normalizarMes(mes), mejor = '';
  filas.forEach(function (r) { var x = normalizarMes(r[0]); if (x && x < m && x > mejor) mejor = x; });
  return mejor;
}

/** Lista de meses para los desplegables: los que existen + el actual y el siguiente. */
function armarListaMeses(existentes, hoy) {
  var set = {};
  existentes.forEach(function (m) { if (/^\d{4}-\d{2}$/.test(m)) set[m] = true; });
  var todos = Object.keys(set).concat([hoy, mesSiguiente(hoy)])
    .filter(function (m, i, a) { return a.indexOf(m) === i; }).sort().reverse();
  var meses = todos.map(function (m) {
    return { v: m, label: nombreMes(m) + ' ' + partesMes(m).anio + (set[m] ? '' : '  (nuevo)'), existe: !!set[m] };
  });
  var ex = Object.keys(set).sort();
  return {
    meses: meses,
    sugeridoSubida: hoy,
    sugeridoGenerar: set[hoy] ? hoy : (ex.length ? ex[ex.length - 1] : hoy)
  };
}


var api = {MESES: MESES, redondear: redondear, pad2: pad2, normalizarMes: normalizarMes, partesMes: partesMes, mesDesplazado: mesDesplazado, mesAnterior: mesAnterior, mesSiguiente: mesSiguiente, nombreMes: nombreMes, fmtNum: fmtNum, fmtSoles: fmtSoles, aFecha: aFecha, fmtFecha: fmtFecha, esSi: esSi, num: num, mismaUnidad: mismaUnidad, calcularMes: calcularMes, datosPlantilla: datosPlantilla, nombrePdf: nombrePdf, nombreEsDeUnidad: nombreEsDeUnidad, listaUnidades: listaUnidades, parseNumero: parseNumero};
if (typeof module !== 'undefined' && module.exports) module.exports = api;
else global.Calculo = api;
})(this);
