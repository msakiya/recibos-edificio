/*
 * excel.js — Lee y escribe el Excel "base de datos" del edificio (con SheetJS).
 * Compatible con las mismas pestañas/columnas de la hoja de Google anterior.
 * Navegador: window.Excel (usa window.XLSX y window.Calculo). Node: module.exports.
 */
(function (global) {
'use strict';
var C = (typeof module !== 'undefined' && module.exports) ? require('./calculo.js') : global.Calculo;
var XLSX = null;
function usar(lib) { XLSX = lib; }

var HOJAS = {
  CONFIG: 'Config', DEPTOS: 'Departamentos', LECTURAS: 'Lecturas', SERVICIOS: 'Servicios',
  GASTOS: 'Gastos', AJUSTES: 'Ajustes', PAGOS: 'Pagos', FONDO: 'Fondo', LOGO: 'Logo', INSTR: 'Instrucciones',
  MES: 'Mes a facturar'
};
var ENCABEZADOS = {
  Config: ['Clave (no cambiar)', 'Valor (editable)', 'Descripción'],
  Departamentos: ['Unidad', 'Propietario', 'Celular / WhatsApp', 'Correo(s) del propietario (varios: separar con coma)', 'Notas'],
  Lecturas: ['Mes (AAAA-MM)', 'Unidad', 'Lectura anterior (m³)', 'Lectura actual (m³)', 'Consumo m³', 'Foto del medidor (archivo)', 'Observación'],
  Servicios: ['Mes (AAAA-MM)', 'Fecha de emisión', 'Fecha de vencimiento', 'Agua: importe total recibo', 'Agua: m³ facturados', 'Luz: importe total recibo', 'Archivos', 'Notas'],
  Gastos: ['Mes (AAAA-MM)', '#', 'Concepto (tal como sale en el recibo)', 'Monto total edificio', 'Tipo', '¿Copiar al mes siguiente?', 'Notas'],
  Ajustes: ['Mes (AAAA-MM)', 'Unidad', 'Concepto (sale en el recibo)', 'Monto (+ cargo / − crédito)', 'Notas'],
  Pagos: ['Mes de la cuota (AAAA-MM)', 'Unidad', 'Fecha de pago', 'Monto pagado', 'Medio', 'N° operación', 'Comprobante', '¿Verificado?', 'Notas'],
  Fondo: ['Mes (AAAA-MM)', 'Saldo anterior', 'Aporte del mes', '(−) Gastos pagados con el fondo', 'Saldo final', 'Detalle de gastos del fondo'],
  Logo: ['Logo del edificio (no editar: imagen guardada por la página)']
};
var CLAVES_CONFIG = [
  ['EDIFICIO_NOMBRE', 'Nombre que aparece en el encabezado del recibo'],
  ['DIRECCION_LINEA_1', 'Dirección, primera línea'],
  ['DIRECCION_LINEA_2', 'Dirección, segunda línea (distrito)'],
  ['PRESIDENTE_NOMBRE', 'Presidente(a) de la junta'],
  ['PRESIDENTE_CELULAR', 'Celular / WhatsApp del presidente(a)'],
  ['TESORERO_NOMBRE', 'Tesorero(a)'],
  ['TESORERO_CELULAR', 'Celular / WhatsApp del tesorero(a)'],
  ['BANCO', 'Banco de la cuenta del edificio'],
  ['CUENTA_BANCARIA', 'Número de cuenta'],
  ['CCI', 'Código de cuenta interbancario'],
  ['TITULAR_CUENTA', 'Titular de la cuenta (nombre completo; sale en el correo y en el WhatsApp)'],
  ['WHATSAPP_CONSTANCIA', 'WhatsApp donde se envía la constancia de pago (vacío = celular del tesorero)'],
  ['NOTA_PAGO', 'Nota que sale junto a los datos de pago'],
  ['CORREO_NOMBRE_EDIFICIO', 'Opcional: nombre en el encabezado del correo (vacío = "Edificio" + nombre del edificio)'],
  ['EMPRESA_AGUA', 'Empresa de agua (ej. Sedapal)'],
  ['EMPRESA_LUZ', 'Empresa de luz (ej. Luz del Sur)'],
  ['NUM_DPTOS', 'Número de departamentos entre los que se dividen los gastos (vacío = los de la pestaña Departamentos)'],
  ['DIA_VENCIMIENTO', 'Día del mes siguiente en que vence el recibo (ej. 5)'],
  ['SALDO_INICIAL_FONDO', 'Saldo del fondo de contingencia antes del primer mes registrado'],
  ['LOGO_BASE64', 'Opcional: logo como texto base64 (más fácil: súbelo en la página y se guarda en la pestaña Logo)']
];
var TIPOS = ['FIJO', 'EXTRA', 'FONDO', 'LUZ', 'AGUA_COMUN'];

/* ----------------------------- utilidades ----------------------------- */
function sinAcentos(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(); }
function hoja(wb, nombre) {
  var k = sinAcentos(nombre);
  for (var i = 0; i < wb.SheetNames.length; i++) if (sinAcentos(wb.SheetNames[i]) === k) return wb.SheetNames[i];
  return null;
}
function filas(wb, nombre) {
  var n = hoja(wb, nombre);
  if (!n) return null;
  return XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
}
function vacio(v) { return v === '' || v === null || v === undefined; }
function serialAFecha(n) {            // número de serie de Excel -> {y,m,d} (sin problemas de zona horaria)
  var ms = Math.round((n - 25569) * 86400000);
  var d = new Date(ms);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}
function aMes(v) {
  if (vacio(v)) return '';
  if (typeof v === 'number' && v > 20000 && v < 80000) { var f = serialAFecha(v); return f.y + '-' + C.pad2(f.m); }
  if (v instanceof Date) return v.getFullYear() + '-' + C.pad2(v.getMonth() + 1);
  return C.normalizarMes(String(v).trim());
}
function aFechaISO(v) {
  if (vacio(v)) return '';
  if (typeof v === 'number' && v > 20000 && v < 80000) { var f = serialAFecha(v); return f.y + '-' + C.pad2(f.m) + '-' + C.pad2(f.d); }
  if (v instanceof Date) return v.getFullYear() + '-' + C.pad2(v.getMonth() + 1) + '-' + C.pad2(v.getDate());
  var d = C.aFecha(String(v).trim());
  return d ? d.getFullYear() + '-' + C.pad2(d.getMonth() + 1) + '-' + C.pad2(d.getDate()) : '';
}
function aNum(v) { if (vacio(v)) return ''; var n = C.parseNumero(v); return n === null ? '' : n; }
function aUnidad(v) { if (vacio(v)) return ''; var s = String(v).trim(); if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10)); return s; }
function hoyISO() { var d = new Date(); return d.getFullYear() + '-' + C.pad2(d.getMonth() + 1) + '-' + C.pad2(d.getDate()); }

/* ------------------------------- lectura ------------------------------- */
/** Lee el libro y devuelve {datos, avisos, errores}. `datos` tiene el formato de Calculo.calcularMes. */
function leerLibro(wb) {
  var errores = [], avisos = [];
  var datos = { config: {}, deptos: [], lecturas: [], servicios: [], gastos: [], ajustes: [], pagos: [], fondo: [], logo: '' };
  var r;
  if ((r = filas(wb, HOJAS.CONFIG))) {
    r.slice(1).forEach(function (f) {
      var k = String(f[0] || '').trim().toUpperCase();
      if (/^[A-Z0-9_]+$/.test(k) && !vacio(f[1])) datos.config[k] = typeof f[1] === 'string' ? f[1].trim() : f[1];
    });
  } else errores.push('Falta la pestaña "Config" (datos del edificio).');
  if ((r = filas(wb, HOJAS.DEPTOS))) {
    // Columna de correo por encabezado ("Correo", "Email"...); Excel antiguo sin esa columna = sin correo
    var colCorreo = -1;
    (r[0] || []).forEach(function (h, i) { if (colCorreo < 0 && /correo|e-?mail/.test(sinAcentos(h))) colCorreo = i; });
    r.slice(1).forEach(function (f) {
      var u = aUnidad(f[0]);
      if (u) datos.deptos.push({ unidad: u, propietario: String(f[1] || '').trim(), telefono: String(f[2] || '').trim(),
        correo: colCorreo >= 0 ? String(f[colCorreo] || '').trim() : '' });
    });
  } else errores.push('Falta la pestaña "Departamentos".');
  if (!datos.deptos.length && !errores.length) errores.push('La pestaña "Departamentos" está vacía: escribe los números de departamento (101, 102…) y sus propietarios.');
  (filas(wb, HOJAS.LECTURAS) || []).slice(1).forEach(function (f) {
    var mes = aMes(f[0]), u = aUnidad(f[1]);
    if (/^\d{4}-\d{2}$/.test(mes) && u) datos.lecturas.push({ mes: mes, unidad: u, anterior: aNum(f[2]), actual: aNum(f[3]), foto: String(f[5] || ''), obs: String(f[6] || '') });
  });
  (filas(wb, HOJAS.SERVICIOS) || []).slice(1).forEach(function (f) {
    var mes = aMes(f[0]);
    if (/^\d{4}-\d{2}$/.test(mes)) datos.servicios.push({ mes: mes, emision: aFechaISO(f[1]), vencimiento: aFechaISO(f[2]),
      aguaImporte: aNum(f[3]), aguaM3: aNum(f[4]), luzImporte: aNum(f[5]), notas: String(f[7] || '') });
  });
  (filas(wb, HOJAS.GASTOS) || []).slice(1).forEach(function (f) {
    var mes = aMes(f[0]);
    if (/^\d{4}-\d{2}$/.test(mes) && (String(f[2] || '').trim() || !vacio(f[3]))) {
      var tipo = String(f[4] || 'FIJO').toUpperCase().trim().replace(/\s+/g, '_');
      if (TIPOS.indexOf(tipo) < 0) tipo = 'FIJO';
      datos.gastos.push({ mes: mes, num: aNum(f[1]), concepto: String(f[2] || '').trim(), monto: aNum(f[3]), tipo: tipo,
        copiar: String(f[5] || ''), notas: String(f[6] || '') });
    }
  });
  (filas(wb, HOJAS.AJUSTES) || []).slice(1).forEach(function (f) {
    var mes = aMes(f[0]), u = aUnidad(f[1]);
    if (/^\d{4}-\d{2}$/.test(mes) && u && !vacio(f[3])) datos.ajustes.push({ mes: mes, unidad: u, concepto: String(f[2] || 'Ajuste').trim(), monto: aNum(f[3]), notas: String(f[4] || '') });
  });
  (filas(wb, HOJAS.PAGOS) || []).slice(1).forEach(function (f) {
    var mes = aMes(f[0]), u = aUnidad(f[1]);
    if (/^\d{4}-\d{2}$/.test(mes) && u) datos.pagos.push({ mes: mes, unidad: u, fecha: aFechaISO(f[2]), monto: aNum(f[3]) });
  });
  (filas(wb, HOJAS.FONDO) || []).slice(1).forEach(function (f) {
    var mes = aMes(f[0]);
    if (/^\d{4}-\d{2}$/.test(mes)) datos.fondo.push({ mes: mes, saldoAnterior: aNum(f[1]), aporte: aNum(f[2]), gastosFondo: aNum(f[3]), saldo: aNum(f[4]), detalle: String(f[5] || '') });
  });
  // logo: pestaña Logo (trozos en la columna A) o clave LOGO_BASE64 de Config
  var lg = filas(wb, HOJAS.LOGO);
  if (lg) datos.logo = lg.slice(1).map(function (f) { return String(f[0] || '').trim(); }).join('');
  if (!datos.logo && datos.config.LOGO_BASE64) datos.logo = String(datos.config.LOGO_BASE64).trim();
  if (datos.logo && !/^data:image\//.test(datos.logo)) datos.logo = /^iVBOR/.test(datos.logo) ? 'data:image/png;base64,' + datos.logo : (/^\/9j\//.test(datos.logo) ? 'data:image/jpeg;base64,' + datos.logo : '');
  ['EDIFICIO_NOMBRE', 'TESORERO_NOMBRE', 'BANCO', 'CUENTA_BANCARIA'].forEach(function (k) {
    if (!datos.config[k]) avisos.push('En la pestaña Config falta ' + k + '.');
  });
  return { datos: datos, errores: errores, avisos: avisos };
}

/* ------------------------- meses y borrador ------------------------- */
function mesesConDatos(datos) {
  var s = {};
  datos.lecturas.forEach(function (l) { s[l.mes] = true; });
  datos.servicios.forEach(function (l) { s[l.mes] = true; });
  return Object.keys(s).sort();
}
/** Mes sugerido: el último registrado si aún no se generó (no tiene fila en Fondo); si no, el siguiente. */
function mesSugerido(datos, hoyMes) {
  var ms = mesesConDatos(datos);
  if (!ms.length) return hoyMes;
  var ult = ms[ms.length - 1];
  var generado = datos.fondo.some(function (f) { return f.mes === ult; });
  return generado ? C.mesSiguiente(ult) : ult;
}
function opcionesMeses(datos, hoyMes) {
  var ms = mesesConDatos(datos);
  var sig = ms.length ? C.mesSiguiente(ms[ms.length - 1]) : hoyMes;
  var todos = ms.concat([sig]);
  if (todos.indexOf(hoyMes) < 0 && !ms.length) todos.push(hoyMes);
  return todos.filter(function (m, i, a) { return a.indexOf(m) === i; }).sort().reverse().map(function (m) {
    return { v: m, label: C.nombreMes(m) + ' ' + C.partesMes(m).anio + (ms.indexOf(m) >= 0 ? '' : ' (mes nuevo)'), existe: ms.indexOf(m) >= 0 };
  });
}
function copiaSiguienteMes(g) {
  if (vacio(g.copiar)) return g.tipo !== 'EXTRA';
  return C.esSi(g.copiar);
}
/** Lo que se muestra en el paso 2 para el mes `mes` (si ya existe en el Excel, sus datos; si no, copiado del mes anterior). */
function prepararBorrador(datos, mes) {
  var existe = mesesConDatos(datos).indexOf(mes) >= 0;
  var b = { mes: mes, existe: existe, lecturas: [], servicios: {}, gastos: [], ajustes: [], fondo: {} };
  datos.deptos.forEach(function (d) {
    var l = datos.lecturas.filter(function (x) { return x.mes === mes && C.mismaUnidad(x.unidad, d.unidad); })[0];
    var prev = datos.lecturas.filter(function (x) { return x.mes < mes && C.mismaUnidad(x.unidad, d.unidad) && !vacio(x.actual); })
      .sort(function (a, c) { return a.mes < c.mes ? 1 : -1; })[0];
    b.lecturas.push({ unidad: d.unidad, propietario: d.propietario,
      anterior: l && !vacio(l.anterior) ? l.anterior : (prev ? prev.actual : ''),
      actual: l ? l.actual : '', obs: l ? l.obs : '' });
  });
  var s = datos.servicios.filter(function (x) { return x.mes === mes; })[0];
  b.servicios = s ? { emision: s.emision || hoyISO(), vencimiento: s.vencimiento || '', aguaImporte: s.aguaImporte, aguaM3: s.aguaM3, luzImporte: s.luzImporte }
                  : { emision: hoyISO(), vencimiento: '', aguaImporte: '', aguaM3: '', luzImporte: '' };
  if (!b.servicios.vencimiento) {
    var pm = C.partesMes(C.mesSiguiente(mes));
    b.servicios.vencimiento = pm.anio + '-' + C.pad2(pm.mes) + '-' + C.pad2(Number(datos.config.DIA_VENCIMIENTO) || 5);
  }
  var gm = datos.gastos.filter(function (g) { return g.mes === mes; });
  if (gm.length) {
    b.gastos = gm.map(function (g) { return { concepto: g.concepto, monto: g.tipo === 'AGUA_COMUN' ? '' : g.monto, tipo: g.tipo, copiar: copiaSiguienteMes(g) }; });
  } else {
    var antes = datos.gastos.filter(function (g) { return g.mes < mes; }).map(function (g) { return g.mes; }).sort();
    var mp = antes[antes.length - 1];
    b.gastos = datos.gastos.filter(function (g) { return g.mes === mp && copiaSiguienteMes(g); })
      .sort(function (a, c) { return C.num(a.num) - C.num(c.num); })
      .map(function (g) { return { concepto: g.concepto, monto: (g.tipo === 'LUZ' || g.tipo === 'AGUA_COMUN') ? '' : g.monto, tipo: g.tipo, copiar: true }; });
    if (!b.gastos.some(function (g) { return g.tipo === 'LUZ'; })) b.gastos.unshift({ concepto: 'Consumo de luz', monto: '', tipo: 'LUZ', copiar: true });
    if (!b.gastos.some(function (g) { return g.tipo === 'AGUA_COMUN'; })) b.gastos.push({ concepto: 'Agua común', monto: '', tipo: 'AGUA_COMUN', copiar: true });
  }
  // Si la luz se escribió como monto del gasto LUZ (y no en Servicios), se pasa al recibo de luz
  b.gastos.forEach(function (g) {
    if (g.tipo === 'LUZ') { if (vacio(b.servicios.luzImporte) && !vacio(g.monto)) b.servicios.luzImporte = g.monto; g.monto = ''; }
  });
  b.ajustes = datos.ajustes.filter(function (a) { return a.mes === mes; }).map(function (a) { return { unidad: a.unidad, concepto: a.concepto, monto: a.monto }; });
  var f = datos.fondo.filter(function (x) { return x.mes === mes; })[0];
  b.fondo = { saldoAnterior: f && !vacio(f.saldoAnterior) ? f.saldoAnterior : saldoFondoAntes(datos, mes),
              gastosFondo: f ? f.gastosFondo : '', detalle: f ? f.detalle : '' };
  return b;
}
function saldoFondoAntes(datos, mes) {
  var prev = datos.fondo.filter(function (f) { return f.mes < mes; }).sort(function (a, c) { return a.mes < c.mes ? 1 : -1; })[0];
  if (!prev) return C.num(datos.config.SALDO_INICIAL_FONDO);
  if (!vacio(prev.saldo)) return C.num(prev.saldo);
  return C.redondear(C.num(prev.saldoAnterior) + C.num(prev.aporte) - C.num(prev.gastosFondo), 2);
}
/** Devuelve una copia de `datos` con el mes del borrador reemplazado (lista para calcular y guardar). */
function aplicarBorrador(datos, b) {
  var mes = b.mes, otro = function (x) { return x.mes !== mes; };
  var n = JSON.parse(JSON.stringify(datos));
  n.mes = mes;
  n.lecturas = n.lecturas.filter(otro).concat(b.lecturas.map(function (l) {
    return { mes: mes, unidad: l.unidad, anterior: C.parseNumero(l.anterior) === null ? '' : C.parseNumero(l.anterior),
             actual: C.parseNumero(l.actual) === null ? '' : C.parseNumero(l.actual), foto: l.foto || '', obs: l.obs || '' };
  }));
  var s = b.servicios;
  n.servicios = n.servicios.filter(otro).concat([{ mes: mes, emision: s.emision || '', vencimiento: s.vencimiento || '',
    aguaImporte: aNum(s.aguaImporte), aguaM3: aNum(s.aguaM3), luzImporte: aNum(s.luzImporte), notas: s.notas || '' }]);
  n.gastos = n.gastos.filter(otro).concat(b.gastos.filter(function (g) { return String(g.concepto || '').trim() || !vacio(g.monto); })
    .map(function (g, i) {
      return { mes: mes, num: i + 1, concepto: String(g.concepto || '').trim(), tipo: g.tipo || 'FIJO',
               monto: (g.tipo === 'LUZ' || g.tipo === 'AGUA_COMUN') ? '' : aNum(g.monto), copiar: g.copiar ? 'SÍ' : 'NO',
               notas: g.tipo === 'LUZ' ? 'Automático: recibo de luz (pestaña Servicios)' : (g.tipo === 'AGUA_COMUN' ? 'Automático: agua de áreas comunes' : '') };
    }));
  n.ajustes = n.ajustes.filter(otro).concat(b.ajustes.filter(function (a) { return a.unidad && C.parseNumero(a.monto); })
    .map(function (a) { return { mes: mes, unidad: String(a.unidad), concepto: String(a.concepto || 'Ajuste').trim(), monto: C.parseNumero(a.monto), notas: '' }; }));
  var res = C.calcularMes(n);
  n.fondo = n.fondo.filter(otro).concat([{ mes: mes, saldoAnterior: C.num(C.parseNumero(b.fondo.saldoAnterior)),
    aporte: res.fondo.aporte, gastosFondo: C.num(C.parseNumero(b.fondo.gastosFondo)), saldo: 0, detalle: b.fondo.detalle || '' }]);
  res = C.calcularMes(n);
  n.fondo[n.fondo.length - 1].saldo = res.fondo.saldo;
  return { datos: n, res: res };
}

/** Revisión antes de generar: errores graves (bloquean) y avisos (se puede continuar). */
function revisar(b, res, archivos) {
  var errores = [], avisos = [];
  var faltan = [], negativos = [], cero = [], sinAnt = [];
  b.lecturas.forEach(function (l) {
    var a = C.parseNumero(l.actual), p = C.parseNumero(l.anterior);
    if (a === null) faltan.push(l.unidad);
    else if (p !== null && a < p) negativos.push(l.unidad);
    else if (p !== null && a === p) cero.push(l.unidad);
    if (p === null) sinAnt.push(l.unidad);
  });
  if (faltan.length) errores.push('Falta la lectura actual del medidor de: ' + C.listaUnidades(faltan) + '.');
  if (negativos.length) errores.push('La lectura actual es MENOR que la anterior (consumo negativo) en: ' + C.listaUnidades(negativos) + '. Revisa los números.');
  if (sinAnt.length) errores.push('Falta la lectura anterior de: ' + C.listaUnidades(sinAnt) + '.');
  if (!C.parseNumero(b.servicios.aguaImporte)) errores.push('Falta el monto total del recibo de agua.');
  if (!C.parseNumero(b.servicios.aguaM3)) errores.push('Faltan los m³ facturados del recibo de agua.');
  if (res && res.agua.m3Comun < 0 && !faltan.length && !negativos.length)
    errores.push('La suma de consumos de los dptos (' + C.fmtNum(res.agua.consumoDptos, 3) + ' m³) es mayor que los m³ del recibo de agua (' + C.fmtNum(res.agua.m3, 0) + '). Revisa las lecturas o los m³.');
  if (!b.gastos.some(function (g) { return g.tipo !== 'LUZ' && g.tipo !== 'AGUA_COMUN' && C.parseNumero(g.monto); })) errores.push('No hay gastos con monto. Revisa la lista de gastos.');
  if (C.parseNumero(b.servicios.luzImporte) === null && b.gastos.some(function (g) { return g.tipo === 'LUZ'; })) avisos.push('No escribiste el monto del recibo de luz (saldrá en S/ 0.00).');
  if (cero.length) avisos.push('Consumo 0 m³ (lectura igual a la anterior) en: ' + C.listaUnidades(cero) + '.');
  if (archivos) {
    var sinFoto = b.lecturas.filter(function (l) { return !archivos.medidor[l.unidad]; }).map(function (l) { return l.unidad; });
    if (sinFoto.length) avisos.push('Sin foto del medidor (el recibo saldrá sin foto): ' + C.listaUnidades(sinFoto) + '.');
    if (!archivos.agua) avisos.push('Falta el recibo de agua del edificio (la página 3 saldrá sin él).');
    if (!archivos.luz) avisos.push('Falta el recibo de luz del edificio (la página 3 saldrá sin él).');
  }
  return { errores: errores, avisos: avisos };
}


/* ===================== Pestaña "Mes a facturar" (lo que se llena cada mes) ===================== */
var MES_ETIQ = {
  mes: 'Mes a facturar (AAAA-MM)',
  aguaImporte: 'Agua (Sedapal): importe total del recibo (S/)',
  aguaM3: 'Agua (Sedapal): m³ facturados',
  luzImporte: 'Luz (Luz del Sur): importe total del recibo (S/)',
  emision: 'Fecha de emisión del recibo (DD/MM/AAAA)',
  vencimiento: 'Fecha de vencimiento (DD/MM/AAAA)',
  saldoAnterior: 'Fondo de contingencia: saldo anterior (S/)',
  gastosFondo: 'Fondo de contingencia: gastos pagados con el fondo este mes (S/)',
  detalle: 'Fondo de contingencia: detalle de esos gastos'
};
var MES_AYUDA = {
  mes: '← mes a cobrar, formato AAAA-MM (ej. 2026-10)',
  aguaImporte: '← monto total del recibo de Sedapal (S/)',
  aguaM3: '← m³ facturados en el recibo de Sedapal',
  luzImporte: '← monto total del recibo de Luz del Sur (S/)',
  emision: '← opcional; vacío = fecha del día en que generas',
  vencimiento: '← opcional; vacío = día DIA_VENCIMIENTO del mes siguiente',
  saldoAnterior: '← viene del mes anterior; cámbialo solo si no coincide',
  gastosFondo: '← lo pagado con el fondo este mes (0 si nada)',
  detalle: '← en qué se usó el fondo (opcional)'
};
var MESES_TXT = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'setiembre', 'octubre', 'noviembre', 'diciembre'];
function letraCol(i) { return String.fromCharCode(65 + i); }
function celda(c, f) { return letraCol(c) + (f + 1); }
function mesDeTexto(v) {
  var m = aMes(v);
  if (/^\d{4}-\d{2}$/.test(m)) return m;
  var t = sinAcentos(v).replace('septiembre', 'setiembre');
  var anio = (t.match(/(20\d{2})/) || [])[1];
  for (var i = 0; i < 12; i++) if (anio && t.indexOf(MESES_TXT[i]) >= 0) return anio + '-' + C.pad2(i + 1);
  return '';
}
function tipoPorConcepto(concepto) {
  var c = sinAcentos(concepto);
  if (/\bluz\b/.test(c)) return 'LUZ';
  if (/agua/.test(c)) return 'AGUA_COMUN';
  if (/fondo|contingencia/.test(c)) return 'FONDO';
  return 'FIJO';
}
function claveDeEtiqueta(a) {
  var t = sinAcentos(a);
  if (!t) return '';
  if (/^mes\b/.test(t) || /mes a facturar/.test(t)) return 'mes';
  if (/agua/.test(t) && /(m3|m³|metros|facturad)/.test(t) && !/importe|monto|s\/|soles/.test(t)) return 'aguaM3';
  if (/agua/.test(t)) return 'aguaImporte';
  if (/luz/.test(t)) return 'luzImporte';
  if (/emision/.test(t)) return 'emision';
  if (/vencimiento/.test(t)) return 'vencimiento';
  if (/saldo anterior/.test(t)) return 'saldoAnterior';
  if (/gastos pagados|gasto del fondo|gastos del fondo/.test(t)) return 'gastosFondo';
  if (/detalle/.test(t)) return 'detalle';
  return '';
}

/** Lee la pestaña "Mes a facturar". Devuelve null si no existe. Cada dato guarda su celda para los mensajes. */
function leerHojaMes(wb, datos) {
  var nombre = hoja(wb, HOJAS.MES);
  if (!nombre) return null;
  var r = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: '' });
  var H = '"' + nombre + '"';
  var campos = {}, ubic = {}, seccion = '', lect = [], gastos = [], ajustes = [];
  r.forEach(function (f, i) {
    var a = f[0], ta = sinAcentos(a);
    if (/^lecturas/.test(ta)) { seccion = 'lecturas'; return; }
    if (/^gastos/.test(ta)) { seccion = 'gastos'; return; }
    if (/^ajustes|^creditos/.test(ta)) { seccion = 'ajustes'; return; }
    if (!seccion) {
      var k = claveDeEtiqueta(a);
      if (k && (!(k in campos) || vacio(campos[k]))) { campos[k] = f[1]; ubic[k] = celda(1, i); }
      return;
    }
    if (seccion === 'lecturas') {
      var u = aUnidad(a);
      if (!u || !/\d/.test(u) || /dpto|unidad|departamento/.test(ta)) return;
      lect.push({ unidad: u, anterior: f[2], actual: f[3], fila: i, obs: String(f[5] || '') });
    } else if (seccion === 'gastos') {
      if (/^concepto/.test(ta) || /^total/.test(ta)) return;
      var con = String(a || '').trim();
      if (!con && vacio(f[1])) return;
      var tipo = String(f[2] || '').toUpperCase().trim().replace(/\s+/g, '_');
      if (TIPOS.indexOf(tipo) < 0) tipo = tipoPorConcepto(con);
      gastos.push({ concepto: con, monto: f[1], tipo: tipo, copiar: tipo !== 'EXTRA', fila: i });
    } else if (seccion === 'ajustes') {
      var ua = aUnidad(a);
      if (!ua || /dpto|unidad|departamento/.test(ta)) return;
      if (vacio(f[2])) return;
      ajustes.push({ unidad: ua, concepto: String(f[1] || 'Ajuste').trim() || 'Ajuste', monto: f[2], fila: i });
    }
  });
  return { hoja: nombre, H: H, campos: campos, ubic: ubic, lecturas: lect, gastos: gastos, ajustes: ajustes };
}

/**
 * Decide el mes a facturar y arma el borrador SOLO con lo que dice el Excel (sin pedir nada a mano).
 * Devuelve {borrador, errores[], avisos[], origen: 'hoja'|'historial'|null}.
 */
function cargarMes(wb, datos) {
  var errores = [], avisos = [];
  var M = leerHojaMes(wb, datos);
  if (!M) return cargarMesHistorial(datos);
  var H = M.H, ub = M.ubic, cp = M.campos;
  function donde(k) { return 'pestaña ' + H + ', celda ' + (ub[k] || ('junto a “' + MES_ETIQ[k] + '”')); }
  var mes = mesDeTexto(cp.mes);
  if (!mes) {
    errores.push('Falta el mes a facturar (ej. 2026-10) → ' + donde('mes') + '.');
    return { borrador: null, errores: errores, avisos: avisos, origen: 'hoja' };
  }
  var b = { mes: mes, existe: mesesConDatos(datos).indexOf(mes) >= 0, lecturas: [], servicios: {}, gastos: [], ajustes: [], fondo: {} };
  // Servicios
  var aI = C.parseNumero(cp.aguaImporte), aM = C.parseNumero(cp.aguaM3), lz = C.parseNumero(cp.luzImporte);
  if (!aI) errores.push('Falta el importe del recibo de agua → ' + donde('aguaImporte') + '.');
  if (!aM) errores.push('Faltan los m³ facturados del recibo de agua → ' + donde('aguaM3') + '.');
  var pm = C.partesMes(C.mesSiguiente(mes));
  b.servicios = {
    aguaImporte: aI === null ? '' : aI, aguaM3: aM === null ? '' : aM, luzImporte: lz === null ? '' : lz,
    emision: aFechaISO(cp.emision) || hoyISO(),
    vencimiento: aFechaISO(cp.vencimiento) || (pm.anio + '-' + C.pad2(pm.mes) + '-' + C.pad2(Number(datos.config.DIA_VENCIMIENTO) || 5))
  };
  // Lecturas: una por departamento de la pestaña Departamentos
  var prevDe = function (u) {
    var p = datos.lecturas.filter(function (x) { return x.mes < mes && C.mismaUnidad(x.unidad, u) && !vacio(x.actual); })
      .sort(function (a, c) { return a.mes < c.mes ? 1 : -1; })[0];
    return p ? p.actual : '';
  };
  var faltan = [], negativas = [], sinAnt = [], cero = [];
  datos.deptos.forEach(function (d) {
    var l = M.lecturas.filter(function (x) { return C.mismaUnidad(x.unidad, d.unidad); })[0];
    if (!l) { errores.push('El Dpto ' + d.unidad + ' no está en la tabla de lecturas → pestaña ' + H + ', sección LECTURAS DE AGUA (columna A).'); return; }
    var ant = C.parseNumero(l.anterior), act = C.parseNumero(l.actual);
    if (ant === null) { var pv = C.parseNumero(prevDe(d.unidad)); if (pv !== null) ant = pv; }
    if (act === null) faltan.push(d.unidad + ' (' + celda(3, l.fila) + ')');
    if (ant === null) sinAnt.push(d.unidad + ' (' + celda(2, l.fila) + ')');
    if (act !== null && ant !== null && act < ant) negativas.push(d.unidad + ' (' + celda(3, l.fila) + ')');
    if (act !== null && ant !== null && act === ant) cero.push(d.unidad);
    b.lecturas.push({ unidad: d.unidad, propietario: d.propietario, anterior: ant === null ? '' : ant, actual: act === null ? '' : act, obs: l.obs || '' });
  });
  if (faltan.length) errores.push('Falta la lectura actual del medidor de: ' + faltan.join(', ') + ' → pestaña ' + H + ', columna D.');
  if (sinAnt.length) errores.push('Falta la lectura anterior de: ' + sinAnt.join(', ') + ' → pestaña ' + H + ', columna C.');
  if (negativas.length) errores.push('La lectura actual es MENOR que la anterior en: ' + negativas.join(', ') + '. Revisa esos números en la pestaña ' + H + '.');
  if (cero.length) avisos.push('Consumo 0 m³ (lectura actual igual a la anterior) en: ' + C.listaUnidades(cero) + '. Si no es correcto, revisa la pestaña ' + H + '.');
  // Gastos
  var malos = [];
  M.gastos.forEach(function (g) {
    var n = C.parseNumero(g.monto);
    if (!vacio(g.monto) && n === null) malos.push(celda(1, g.fila));
    var monto = n === null ? '' : n;
    if (g.tipo === 'LUZ') { if (lz === null && n !== null) b.servicios.luzImporte = n; monto = ''; }
    if (g.tipo === 'AGUA_COMUN') monto = '';
    b.gastos.push({ concepto: g.concepto || (g.tipo === 'LUZ' ? 'Consumo de luz' : g.tipo === 'AGUA_COMUN' ? 'Agua común' : 'Gasto'), monto: monto, tipo: g.tipo, copiar: g.copiar });
  });
  if (malos.length) errores.push('Monto no válido en la sección GASTOS (celdas ' + malos.join(', ') + ') de la pestaña ' + H + '. Escribe solo números (ej. 349.34).');
  if (!b.gastos.some(function (g) { return g.tipo === 'AGUA_COMUN'; })) b.gastos.push({ concepto: 'Agua común', monto: '', tipo: 'AGUA_COMUN', copiar: true });
  if (!b.gastos.some(function (g) { return g.tipo !== 'LUZ' && g.tipo !== 'AGUA_COMUN' && C.parseNumero(g.monto); }))
    errores.push('No hay gastos con monto → pestaña ' + H + ', sección GASTOS COMUNES DEL MES (columna B).');
  if (b.servicios.luzImporte === '') {
    if (b.gastos.some(function (g) { return g.tipo === 'LUZ'; })) avisos.push('No hay monto del recibo de luz (saldrá S/ 0.00) → ' + donde('luzImporte') + '.');
  } else if (!b.gastos.some(function (g) { return g.tipo === 'LUZ'; })) {
    b.gastos.unshift({ concepto: 'Consumo de luz', monto: '', tipo: 'LUZ', copiar: true });
  }
  // Ajustes
  M.ajustes.forEach(function (a) {
    var n = C.parseNumero(a.monto);
    if (n === null) { errores.push('Monto no válido en AJUSTES, celda ' + celda(2, a.fila) + ' de la pestaña ' + H + '.'); return; }
    if (!datos.deptos.some(function (d) { return C.mismaUnidad(d.unidad, a.unidad); }))
      { errores.push('El ajuste de la celda ' + celda(0, a.fila) + ' es para el Dpto ' + a.unidad + ', que no está en la pestaña Departamentos.'); return; }
    b.ajustes.push({ unidad: a.unidad, concepto: a.concepto, monto: n });
  });
  // Fondo
  var sa = C.parseNumero(cp.saldoAnterior), gf = C.parseNumero(cp.gastosFondo);
  b.fondo = { saldoAnterior: sa === null ? saldoFondoAntes(datos, mes) : sa, gastosFondo: gf === null ? '' : gf, detalle: vacio(cp.detalle) ? '' : String(cp.detalle) };
  if (b.existe && datos.fondo.some(function (f) { return f.mes === mes; }))
    avisos.push('El mes ' + C.nombreMes(mes) + ' ' + C.partesMes(mes).anio + ' ya se había generado antes: al generar se reemplaza en el historial.');
  // Agua común negativa
  if (!errores.length) {
    var res = aplicarBorrador(datos, b).res;
    if (res.agua.m3Comun < 0) errores.push('La suma de consumos de los dptos (' + C.fmtNum(res.agua.consumoDptos, 3) + ' m³) es mayor que los m³ del recibo de agua (' +
      C.fmtNum(res.agua.m3, 0) + ') → revisa las lecturas o ' + donde('aguaM3') + '.');
  }
  return { borrador: errores.length ? null : b, errores: errores, avisos: avisos, origen: 'hoja' };
}

/** Excel antiguo (sin pestaña "Mes a facturar"): usa el último mes de las pestañas Lecturas/Servicios si está completo. */
function cargarMesHistorial(datos) {
  var errores = [], avisos = [];
  var ms = mesesConDatos(datos);
  var ult = ms[ms.length - 1];
  var NOTA = 'Tu Excel no tiene la pestaña "Mes a facturar". Descarga tu Excel con esa pestaña (botón de abajo), llénala con los datos del mes y vuelve a subirlo.';
  if (!ult) return { borrador: null, errores: [NOTA], avisos: avisos, origen: null, sinHoja: true };
  var b = prepararBorrador(datos, ult);
  var generado = datos.fondo.some(function (f) { return f.mes === ult; });
  var res = aplicarBorrador(datos, b).res;
  var r = revisar(b, res, null);
  if (generado || r.errores.length) {
    return { borrador: null, errores: [NOTA], avisos: avisos, origen: null, sinHoja: true, mesSugerido: generado ? C.mesSiguiente(ult) : ult };
  }
  avisos.push('Datos leídos de las pestañas Lecturas y Servicios (' + C.nombreMes(ult) + ' ' + C.partesMes(ult).anio + '). El Excel que se descarga ya trae la pestaña "Mes a facturar" para el mes siguiente.');
  return { borrador: b, errores: [], avisos: avisos.concat(r.avisos), origen: 'historial' };
}

/** Escribe la pestaña "Mes a facturar" con el borrador `b` (o el del mes `mes` preparado desde el historial). Queda como 1ª pestaña. */
function escribirHojaMes(wb, datos, mes, b) {
  var explicito = !!b;
  b = b || prepararBorrador(datos, mes);
  var s = b.servicios || {};
  function fechaTxt(iso) { var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + '/' + m[2] + '/' + m[1] : ''; }
  function nv(v) { return vacio(v) || v === null ? '' : v; }
  var aoa = [
    ['DATOS DEL MES A FACTURAR — llena esta pestaña cada mes y sube el Excel a la página'],
    ['Solo cambia las celdas en blanco o con números. La página lee todo de aquí: no hay que escribir nada más en la página.'],
    [''],
    [MES_ETIQ.mes, b.mes, MES_AYUDA.mes],
    [MES_ETIQ.aguaImporte, nv(s.aguaImporte), MES_AYUDA.aguaImporte],
    [MES_ETIQ.aguaM3, nv(s.aguaM3), MES_AYUDA.aguaM3],
    [MES_ETIQ.luzImporte, nv(s.luzImporte), MES_AYUDA.luzImporte],
    [MES_ETIQ.emision, (b.existe || explicito) ? fechaTxt(s.emision) : '', MES_AYUDA.emision],
    [MES_ETIQ.vencimiento, fechaTxt(s.vencimiento), MES_AYUDA.vencimiento],
    [MES_ETIQ.saldoAnterior, nv(b.fondo && b.fondo.saldoAnterior), MES_AYUDA.saldoAnterior],
    [MES_ETIQ.gastosFondo, nv(b.fondo && b.fondo.gastosFondo), MES_AYUDA.gastosFondo],
    [MES_ETIQ.detalle, (b.fondo && b.fondo.detalle) || '', MES_AYUDA.detalle],
    [''],
    ['LECTURAS DE AGUA (m³) — escribe la lectura ACTUAL de cada medidor (columna D)'],
    ['Dpto', 'Propietario', 'Lectura anterior', 'Lectura actual', 'Consumo m³', 'Observación']
  ];
  var filaLect0 = aoa.length;
  b.lecturas.forEach(function (l, i) {
    aoa.push([isNaN(Number(l.unidad)) ? l.unidad : Number(l.unidad), l.propietario || '', nv(l.anterior), nv(l.actual), '', l.obs || '']);
  });
  aoa.push(['']);
  aoa.push(['GASTOS COMUNES DEL MES — monto TOTAL del edificio (la página lo divide entre los dptos)']);
  aoa.push(['Concepto (tal como sale en el recibo)', 'Monto total (S/)', 'Tipo', 'Nota']);
  b.gastos.forEach(function (g) {
    var nota = g.tipo === 'LUZ' ? 'Automático: usa el importe de luz de arriba' : g.tipo === 'AGUA_COMUN' ? 'Automático: se calcula con el agua' :
      g.tipo === 'EXTRA' ? 'Solo este mes' : g.tipo === 'FONDO' ? 'Aporte al fondo de contingencia' : '';
    aoa.push([g.concepto, (g.tipo === 'LUZ' || g.tipo === 'AGUA_COMUN') ? '' : nv(g.monto), g.tipo, nota]);
  });
  for (var k = 0; k < 3; k++) aoa.push(['', '', 'EXTRA', k === 0 ? '← filas libres: escribe concepto y monto de un gasto EXTRA de este mes' : '']);
  aoa.push(['']);
  aoa.push(['AJUSTES / CRÉDITOS (opcional) — monto negativo = crédito a favor (ej. -5)']);
  aoa.push(['Dpto', 'Concepto (sale en el recibo)', 'Monto (S/)']);
  b.ajustes.forEach(function (a) { aoa.push([isNaN(Number(a.unidad)) ? a.unidad : Number(a.unidad), a.concepto, a.monto]); });
  for (var j = 0; j < 4; j++) aoa.push(['', '', '']);
  aoa.push(['']);
  aoa.push(['TIPOS de gasto: FIJO (se repite cada mes) · EXTRA (solo este mes) · FONDO (aporte al fondo) · LUZ y AGUA_COMUN (se calculan solos, deja el monto vacío).']);
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  // Consumo = actual − anterior (fórmula, solo como ayuda visual)
  b.lecturas.forEach(function (l, i) {
    var f = filaLect0 + i + 1;
    ws['E' + f] = { t: 'n', f: 'IF(D' + f + '="","",ROUND(D' + f + '-C' + f + ',3))' };
  });
  ws['!cols'] = [{ wch: 58 }, { wch: 30 }, { wch: 18 }, { wch: 40 }, { wch: 12 }, { wch: 24 }];
  var n = hoja(wb, HOJAS.MES);
  if (n) wb.Sheets[n] = ws;
  else { XLSX.utils.book_append_sheet(wb, ws, HOJAS.MES); n = HOJAS.MES; }
  wb.SheetNames.splice(wb.SheetNames.indexOf(n), 1);
  wb.SheetNames.unshift(n);
  if (wb.Workbook && wb.Workbook.Sheets) delete wb.Workbook.Sheets; // evita índices de hojas desfasados
  return wb;
}

/* ------------------------------- escritura ------------------------------- */
function reemplazarMes(filasViejas, mes, nuevas) {
  var cuerpo = filasViejas.slice(1).filter(function (f) { return aMes(f[0]) !== mes && f.some(function (x) { return !vacio(x); }); });
  return cuerpo.concat(nuevas);
}
function ponerHoja(wb, nombre, aoa, anchos) {
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  if (anchos) ws['!cols'] = anchos.map(function (w) { return { wch: w }; });
  // Meses como texto (evita que Excel los convierta en fecha)
  var n = hoja(wb, nombre);
  if (n) wb.Sheets[n] = ws; else XLSX.utils.book_append_sheet(wb, ws, nombre);
}
function fechasATexto(f) { return f.map(function (x) { return x instanceof Date ? aFechaISO(x) : x; }); }
function filasDe(wb, nombre) { var r = filas(wb, nombre) || [[]]; return r.map(fechasATexto); }
/** Escribe el mes `mes` de `datos` en el libro (lo modifica y lo devuelve). Las demás pestañas no se tocan. */
function escribirMes(wb, datos, mes) {
  var es = function (x) { return x.mes === mes; };
  ponerHoja(wb, HOJAS.LECTURAS, [ENCABEZADOS.Lecturas].concat(reemplazarMes(filasDe(wb, HOJAS.LECTURAS), mes,
    datos.lecturas.filter(es).map(function (l) {
      var cons = (vacio(l.actual) || vacio(l.anterior)) ? '' : C.redondear(l.actual - l.anterior, 3);
      return [mes, isNaN(Number(l.unidad)) ? l.unidad : Number(l.unidad), l.anterior, l.actual, cons, l.foto || '', l.obs || ''];
    }))), [14, 8, 18, 18, 12, 24, 30]);
  ponerHoja(wb, HOJAS.SERVICIOS, [ENCABEZADOS.Servicios].concat(reemplazarMes(filasDe(wb, HOJAS.SERVICIOS), mes,
    datos.servicios.filter(es).map(function (s) { return [mes, s.emision, s.vencimiento, s.aguaImporte, s.aguaM3, s.luzImporte, s.archivos || '', s.notas || '']; }))),
    [14, 14, 14, 16, 14, 16, 30, 30]);
  ponerHoja(wb, HOJAS.GASTOS, [ENCABEZADOS.Gastos].concat(reemplazarMes(filasDe(wb, HOJAS.GASTOS), mes,
    datos.gastos.filter(es).map(function (g) { return [mes, g.num, g.concepto, g.monto, g.tipo, g.copiar, g.notas || '']; }))),
    [14, 5, 60, 14, 12, 12, 40]);
  ponerHoja(wb, HOJAS.AJUSTES, [ENCABEZADOS.Ajustes].concat(reemplazarMes(filasDe(wb, HOJAS.AJUSTES), mes,
    datos.ajustes.filter(es).map(function (a) { return [mes, isNaN(Number(a.unidad)) ? a.unidad : Number(a.unidad), a.concepto, a.monto, a.notas || '']; }))),
    [14, 8, 40, 16, 30]);
  ponerHoja(wb, HOJAS.FONDO, [ENCABEZADOS.Fondo].concat(reemplazarMes(filasDe(wb, HOJAS.FONDO), mes,
    datos.fondo.filter(es).map(function (f) { return [mes, f.saldoAnterior, f.aporte, f.gastosFondo, f.saldo, f.detalle || '']; }))),
    [14, 14, 14, 18, 14, 40]);
  return wb;
}
function ponerLogo(wb, dataUrl) {
  var aoa = [ENCABEZADOS.Logo];
  for (var i = 0; i < (dataUrl || '').length; i += 30000) aoa.push([dataUrl.slice(i, i + 30000)]);
  ponerHoja(wb, HOJAS.LOGO, aoa, [60]);
}
var TEXTO_INSTRUCCIONES = [
  ['RECIBOS DEL EDIFICIO — cómo usar este Excel'],
  [''],
  ['1. Este archivo es la "base de datos". Guárdalo en tu Drive o en tu computadora y NO lo pierdas.'],
  ['2. Cada mes llena SOLO la pestaña "Mes a facturar": mes, recibo de agua (S/ y m³), recibo de luz, lecturas actuales de los medidores, gastos y créditos.'],
  ['3. Abre la página de recibos y sube este Excel. La página muestra un resumen para revisar (no hay que escribir nada en la página).'],
  ['4. Sube las fotos y pulsa "Generar recibos": aparece la lista de departamentos con su PDF, y los botones para enviar por correo y WhatsApp.'],
  ['   Descarga también el Excel ACTUALIZADO (botón "Descargar Excel actualizado"): ya trae la pestaña "Mes a facturar" lista para el mes siguiente.'],
  ['   Guarda el Excel nuevo en lugar de este.'],
  [''],
  ['Pestañas:'],
  ['  Mes a facturar: lo único que se llena cada mes.'],
  ['  Config: datos del edificio (presidente, tesorero, banco, cuenta, CCI, titular, WhatsApp de constancias, día de vencimiento...). Cambia solo la columna B.'],
  ['  Departamentos: un renglón por departamento (unidad, propietario, celular/WhatsApp de 9 dígitos y correo; si hay varios correos, sepáralos con coma).'],
  ['  Lecturas, Servicios, Gastos, Ajustes, Fondo: historial de los meses ya generados (la página los llena sola; no hace falta tocarlos).'],
  ['  Tipos de gasto: FIJO, EXTRA (no se repite), FONDO (aporte al fondo de contingencia), LUZ (= recibo de luz), AGUA_COMUN (se calcula solo).'],
  ['  Logo: lo llena la página si subes el logo. No la edites.'],
  [''],
  ['Cálculo: cada gasto se divide entre el número de dptos; el agua de cada dpto = consumo × (monto del recibo de agua ÷ m³ facturados);'],
  ['el agua común = (m³ del recibo − m³ de los dptos) × precio; el céntimo de redondeo del agua va al dpto de mayor consumo.']
];
/** Libro nuevo con todas las pestañas (para "Excel en blanco" y la plantilla de ejemplo). */
function libroNuevo(datos, opc) {
  opc = opc || {};
  var wb = XLSX.utils.book_new();
  datos = datos || { config: {}, deptos: [], lecturas: [], servicios: [], gastos: [], ajustes: [], pagos: [], fondo: [] };
  var ws = XLSX.utils.aoa_to_sheet(TEXTO_INSTRUCCIONES); ws['!cols'] = [{ wch: 120 }];
  XLSX.utils.book_append_sheet(wb, ws, HOJAS.INSTR);
  ponerHoja(wb, HOJAS.CONFIG, [ENCABEZADOS.Config].concat(CLAVES_CONFIG.map(function (c) {
    var v = datos.config[c[0]]; return [c[0], vacio(v) ? '' : v, c[1]];
  })), [24, 50, 70]);
  ponerHoja(wb, HOJAS.DEPTOS, [ENCABEZADOS.Departamentos].concat(datos.deptos.map(function (d) {
    return [isNaN(Number(d.unidad)) ? d.unidad : Number(d.unidad), d.propietario || '', d.telefono || '', d.correo || '', ''];
  })), [8, 40, 18, 28, 30]);
  [HOJAS.LECTURAS, HOJAS.SERVICIOS, HOJAS.GASTOS, HOJAS.AJUSTES, HOJAS.FONDO].forEach(function (h) { ponerHoja(wb, h, [ENCABEZADOS[h]]); });
  ponerHoja(wb, HOJAS.PAGOS, [ENCABEZADOS.Pagos].concat((datos.pagos || []).map(function (p) { return [p.mes, Number(p.unidad), p.fecha || '', p.monto]; })), [14, 8, 14, 14, 14, 14, 24, 12, 30]);
  var meses = mesesConDatos(datos);
  meses.forEach(function (m) { escribirMes(wb, datos, m); });
  if (datos.logo) ponerLogo(wb, datos.logo);
  if (opc.mes || opc.borrador) escribirHojaMes(wb, datos, opc.mes || opc.borrador.mes, opc.borrador);
  return wb;
}

var api = { usar: usar, HOJAS: HOJAS, ENCABEZADOS: ENCABEZADOS, CLAVES_CONFIG: CLAVES_CONFIG, TIPOS: TIPOS,
  leerLibro: leerLibro, mesesConDatos: mesesConDatos, mesSugerido: mesSugerido, opcionesMeses: opcionesMeses,
  prepararBorrador: prepararBorrador, aplicarBorrador: aplicarBorrador, revisar: revisar, saldoFondoAntes: saldoFondoAntes,
  escribirMes: escribirMes, ponerLogo: ponerLogo, libroNuevo: libroNuevo,
  leerHojaMes: leerHojaMes, cargarMes: cargarMes, escribirHojaMes: escribirHojaMes, mesDeTexto: mesDeTexto, aMes: aMes, aFechaISO: aFechaISO, hoyISO: hoyISO };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else global.Excel = api;
})(this);
