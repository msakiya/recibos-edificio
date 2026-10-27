/*
 * recibo-html.js — Arma el HTML de las 3 páginas del recibo (mismo diseño del ejemplo).
 * Escapa el texto; las imágenes van como data URL.
 */
(function (global) {
'use strict';
function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function filasConceptos(d) {
  var h = '';
  for (var i = 0; i < d.lineas.length; i++) h += '<tr><td>' + esc(d.lineas[i].concepto) + '</td><td class="m">' + esc(d.lineas[i].monto) + '</td></tr>';
  h += '<tr><td>' + esc(d.aguaTexto) + '</td><td class="m">' + esc(d.aguaMonto) + '</td></tr>';
  return h;
}
function filasRep(arr) {
  var h = '';
  for (var i = 0; i < arr.length; i++) h += '<tr><td class="n">' + esc(arr[i].n) + '</td><td>' + esc(arr[i].concepto) + '</td><td class="m">' + esc(arr[i].monto) + '</td></tr>';
  return h;
}
var CSS = '@page{size:A4;margin:0}' +
  'body{font-family:Arial,Helvetica,sans-serif;font-size:10.5px;color:#222;margin:0;background:#fff}' +
  '.pagina{width:210mm;min-height:297mm;box-sizing:border-box;padding:14mm 14mm 12mm 14mm;background:#fff;position:relative}' +
  'table{border-collapse:collapse;width:100%}td{vertical-align:top;padding:0}' +
  '.teal{color:#1d5e5a}.titulo-edif{font-size:20px;font-weight:bold;color:#1d5e5a;text-align:center}' +
  '.dir{text-align:center;font-size:10.5px;color:#333}' +
  '.meta td{font-size:10.5px;color:#444;padding:1px 0}.meta td.v{text-align:right;color:#222}' +
  '.banda{border-top:1.5px solid #1d5e5a;border-bottom:1px solid #1d5e5a;text-align:center;font-size:13px;font-weight:bold;color:#1d5e5a;padding:5px 0;margin-top:10px}' +
  '.lbl{font-size:10px;color:#444}.prop{font-size:13px;font-weight:bold}.unidad{font-size:22px;font-weight:bold;text-align:right}' +
  '.sec{font-size:12px;font-weight:bold;margin-top:10px}.sub{font-size:9.5px;color:#555;margin-bottom:6px}' +
  '.conc td{padding:2.5px 0;font-size:10.5px}.conc td.m{text-align:right;white-space:nowrap;width:90px;padding-left:12px}' +
  '.total-mes{font-size:12px;font-weight:bold;margin-top:8px}.total-pagar{font-size:19px;font-weight:bold;color:#1d5e5a;text-align:right;margin-top:14px}' +
  '.foto{text-align:center;margin-top:12px}.foto img{max-height:190px;max-width:100%}.foto-cap{font-size:9.5px;color:#1d5e5a;font-style:italic;text-align:center;margin-top:3px}' +
  '.dep td{font-size:10.5px;padding:2px 0}.nota{font-size:9.5px;font-style:italic;color:#333}' +
  '.rep-titulo{border-top:1.5px solid #1d5e5a;border-bottom:1px solid #1d5e5a;text-align:center;font-size:13px;font-weight:bold;color:#1d5e5a;padding:5px 0}' +
  '.rep-sec{font-size:12px;font-weight:bold;margin:14px 0 5px 0}' +
  '.rep th{font-size:9.5px;color:#1d5e5a;text-align:left;padding:3px 4px 3px 0;border-bottom:1px solid #ccc}' +
  '.rep th.m,.rep td.m{text-align:right;white-space:nowrap;width:90px}.rep td{font-size:10px;padding:2.5px 4px 2.5px 0}.rep td.n{width:28px}' +
  '.rep-total td{font-size:11px;font-weight:bold;padding-top:6px}.doble{border-bottom:3px double #e07b39}' +
  '.res td{font-size:10.5px;padding:3px 0 3px 12px}.res td.m{text-align:right}' +
  '.fondo{border:1px solid #333;padding:6px 12px}.fondo td{font-size:10.5px;padding:3px 0}.fondo td.m{text-align:right;font-weight:bold}' +
  '.verde{color:#2e7d32}.rojo{color:#c62828}' +
  '.comp-titulo{border-top:1px dashed #1d5e5a;text-align:center;font-size:13px;font-weight:bold;color:#1d5e5a;padding-top:8px}' +
  '.comp-lbl{font-size:13px;font-weight:bold;color:#1d5e5a;padding:6px 0 4px 0}' +
  '.vacio{font-size:10px;color:#999;font-style:italic;padding:30px 0;text-align:center;border:1px dashed #bbb}';

function pagina1(d) {
  return '<div class="pagina">' +
    '<table><tr><td style="width:110px">' + (d.logo ? '<img src="' + d.logo + '" width="80">' : '') + '</td>' +
    '<td><div class="titulo-edif">' + esc(d.edificio) + '</div><div class="dir">' + esc(d.dir1) + '<br>' + esc(d.dir2) + '</div></td>' +
    '<td style="width:200px"><table class="meta">' +
    '<tr><td>Emisión:</td><td class="v">' + esc(d.emision) + '</td></tr>' +
    '<tr><td>Cuota Mes:</td><td class="v">' + esc(d.cuotaMes) + '</td></tr>' +
    '<tr><td>Vencimiento:</td><td class="v">' + esc(d.vencimiento) + '</td></tr>' +
    '<tr><td>Último Pago por:</td><td class="v">' + esc(d.ultimoPago) + '</td></tr>' +
    '</table></td></tr></table>' +
    '<div class="banda">DETALLE DE GASTOS COMUNES</div>' +
    '<table style="margin-top:8px"><tr><td><div class="lbl">Propietario:</div><div class="prop">' + esc(d.propietario) + '</div></td>' +
    '<td style="width:120px"><div class="lbl" style="text-align:right">Unidad:</div><div class="unidad">' + esc(d.unidad) + '</div></td></tr></table>' +
    '<div class="sec">B. Conceptos de la Cuota del Mes de ' + esc(d.mesMin) + '</div>' +
    '<div class="sub">A continuación se detallan los conceptos incluidos en la cuota de su recibo.</div>' +
    '<table class="conc">' + filasConceptos(d) + '</table>' +
    '<div class="total-mes">Total del Mes: Monto exacto: ' + esc(d.total) + '</div>' +
    '<table class="conc"><tr><td>Fecha de Vencimiento: ' + esc(d.vencimiento) + '</td><td class="m">' + esc(d.total) + '</td></tr></table>' +
    '<div class="total-pagar">TOTAL A PAGAR: ' + esc(d.total) + '</div>' +
    '<div class="foto">' + (d.foto ? '<img src="' + d.foto + '">' : '<div class="vacio">(Foto del medidor no disponible)</div>') + '</div>' +
    '<div class="foto-cap">Evidencia fotográfica de la lectura del medidor de agua</div>' +
    '<div class="banda" style="font-size:11.5px">DATOS PARA DEPÓSITO O TRANSFERENCIA</div>' +
    '<table class="dep" style="margin-top:6px"><tr><td style="width:50%">Tesorero(a): ' + esc(d.tesorero) + '<br>Celular / WhatsApp: ' + esc(d.tesoreroCel) +
    '<br>Banco: ' + esc(d.banco) + '<br>Cuenta Bancaria: ' + esc(d.cuenta) + '<br>CCI: ' + esc(d.cci) +
    '</td><td>Presidente(a): ' + esc(d.presidente) + '<br>Celular / WhatsApp: ' + esc(d.presidenteCel) +
    '<br><div class="nota">' + esc(d.nota) + '</div></td></tr></table></div>';
}
function pagina2(d) {
  var extras = d.extras.length
    ? ('<table class="rep"><tr><th class="n">#</th><th>NOMBRE DEL CONCEPTO</th><th class="m">SUBTOTAL</th></tr>' + filasRep(d.extras) + '</table>' +
       '<table class="rep-total"><tr><td>TOTAL GASTOS EXTRAS</td><td style="text-align:right;width:110px"><span class="doble">' + esc(d.totalExtras) + '</span></td></tr></table>')
    : '<div class="sub">No hubo gastos extras este mes.</div>';
  return '<div class="pagina"><div class="rep-titulo">REPORTE DE GASTOS Y FONDO DE CONTINGENCIA ' + esc(d.tituloReporte) + '</div>' +
    '<div class="rep-sec">1. GASTOS FIJOS DEL MES</div>' +
    '<table class="rep"><tr><th class="n">#</th><th>NOMBRE DEL CONCEPTO</th><th class="m">SUBTOTAL</th></tr>' + filasRep(d.fijos) + '</table>' +
    '<table class="rep-total"><tr><td>TOTAL GASTOS FIJOS</td><td style="text-align:right;width:110px"><span class="doble">' + esc(d.totalFijos) + '</span></td></tr></table>' +
    '<div class="rep-sec">2. GASTOS EXTRAS DEL MES</div>' + extras +
    '<div class="rep-sec">3. RESUMEN GENERAL A COBRAR EN RECIBOS</div>' +
    '<table class="res"><tr><td>Total Gasto Común (Fijos + Extras)</td><td class="m">' + esc(d.totalComun) + '</td></tr>' +
    '<tr><td>Total Consumo de agua privada</td><td class="m">' + esc(d.totalAguaPrivada) + '</td></tr>' +
    '<tr><td style="font-size:12px;font-weight:bold;padding-top:8px">TOTAL FACTURADO EN EL MES</td>' +
    '<td class="m" style="font-size:12px;font-weight:bold;padding-top:8px"><span class="doble">' + esc(d.totalFacturado) + '</span></td></tr></table>' +
    '<div class="rep-sec">4. ESTADO DEL FONDO DE CONTINGENCIA</div><div class="fondo"><table>' +
    '<tr><td>Saldo Anterior del Fondo:</td><td class="m">' + esc(d.fondoAnterior) + '</td></tr>' +
    '<tr><td>Aporte del mes ' + esc(d.mesAnio) + ' (Fondo de contingencia):</td><td class="m verde">' + esc(d.fondoAporte) + '</td></tr>' +
    '<tr><td>(−) Gastos Extras / Mantenimientos' + (d.fondoDetalle ? ' (' + esc(d.fondoDetalle) + ')' : '') + ':</td><td class="m rojo">' + esc(d.fondoGastos) + '</td></tr>' +
    '<tr><td class="teal" style="border-top:1px solid #999;font-weight:bold;font-size:11.5px;padding-top:6px">Fondo de contingencia:</td>' +
    '<td class="m teal" style="border-top:1px solid #999;font-size:11.5px;padding-top:6px">' + esc(d.fondoSaldo) + '</td></tr>' +
    '</table></div></div>';
}
function pagina3(d) {
  return '<div class="pagina"><div class="comp-titulo">COMPROBANTES DE SERVICIOS - ' + esc(d.mesAnioMay) + '</div>' +
    '<table style="margin-top:8px"><tr>' +
    '<td style="width:50%;padding-right:6px"><div class="comp-lbl">Recibo de Agua (' + esc(d.empresaAgua) + ')</div>' +
    (d.imgAgua ? '<img src="' + d.imgAgua + '" style="width:100%;max-width:330px">' : '<div class="vacio">(Recibo de agua no encontrado)</div>') + '</td>' +
    '<td style="width:50%;padding-left:6px"><div class="comp-lbl">Recibo de Luz (' + esc(d.empresaLuz) + ')</div>' +
    (d.imgLuz ? '<img src="' + d.imgLuz + '" style="width:100%;max-width:330px">' : '<div class="vacio">(Recibo de luz no encontrado)</div>') + '</td>' +
    '</tr></table></div>';
}
function documento(d) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + CSS + '</style></head><body>' +
    pagina1(d) + pagina2(d) + pagina3(d) + '</body></html>';
}
var api = { CSS: CSS, pagina1: pagina1, pagina2: pagina2, pagina3: pagina3, documento: documento, esc: esc };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else global.ReciboHtml = api;
})(this);
