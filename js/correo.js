/*
 * correo.js — Arma el correo y el WhatsApp de cada departamento (asunto, cuerpo HTML/texto, links).
 * No envía nada: la persona revisa y envía. Navegador: window.Correo. Node: module.exports.
 */
(function (global) {
'use strict';

var COLOR = '#5b5fe0';
var ROMANOS = /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)$/i;

function esc(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/** "NÁPOLES II" -> "Edificio Nápoles II". Usa CORREO_NOMBRE_EDIFICIO si existe en Config. */
function nombreEdificio(cfg) {
  cfg = cfg || {};
  var propio = String(cfg.CORREO_NOMBRE_EDIFICIO || '').trim();
  if (propio) return propio;
  var n = String(cfg.EDIFICIO_NOMBRE || '').trim().replace(/\s+/g, ' ');
  if (!n) return 'Edificio';
  if (n === n.toUpperCase()) {
    n = n.split(' ').map(function (w) {
      if (ROMANOS.test(w)) return w.toUpperCase();
      if (/^(de|del|la|las|los|el|y)$/i.test(w)) return w.toLowerCase();
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    }).join(' ');
  }
  if (!/^edificio\b/i.test(n)) n = 'Edificio ' + n;
  return n;
}

/** Separa "a@x.com, b@y.com; c@z.com" en una lista de correos válidos (sin repetir). */
function listaCorreos(txt) {
  var vistos = {};
  return String(txt || '').split(/[,;\s]+/).map(function (s) { return s.trim().replace(/^<|>$/g, ''); })
    .filter(function (s) {
      var ok = /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s) && !vistos[s.toLowerCase()];
      if (ok) vistos[s.toLowerCase()] = true;
      return ok;
    });
}

/** Teléfono para WhatsApp: "987 654 321" -> "51987654321"; "+51 943..." -> "51943..."; vacío/raro -> "". */
function telefonoWhatsapp(txt) {
  var s = String(txt === null || txt === undefined ? '' : txt).split(/[\/,;]| y /)[0];
  var d = s.replace(/\D/g, '');
  if (!d) return '';
  if (d.length === 9 && d.charAt(0) === '9') return '51' + d;       // celular peruano
  if (d.length === 11 && d.slice(0, 2) === '51') return d;
  if (d.length === 13 && d.slice(0, 4) === '0051') return d.slice(2);
  if (d.length >= 10 && d.length <= 15 && /^\+/.test(s.trim())) return d;  // otro país con +
  return '';
}

/** "987654321" -> "987 654 321" (solo para mostrar). */
function telefonoBonito(txt) {
  var d = String(txt === null || txt === undefined ? '' : txt).replace(/\D/g, '');
  if (d.length === 11 && d.slice(0, 2) === '51') d = d.slice(2);
  if (d.length === 9) return d.slice(0, 3) + ' ' + d.slice(3, 6) + ' ' + d.slice(6);
  return String(txt || '').trim();
}

/** Datos de pago desde Config. */
function datosPago(cfg) {
  cfg = cfg || {};
  var wa = String(cfg.WHATSAPP_CONSTANCIA || '').trim() || String(cfg.TESORERO_CELULAR || '').trim();
  return {
    banco: String(cfg.BANCO || '').trim(),
    cuenta: String(cfg.CUENTA_BANCARIA || '').trim(),
    cci: String(cfg.CCI || '').trim(),
    titular: String(cfg.TITULAR_CUENTA || '').trim(),
    whatsapp: wa ? telefonoBonito(wa) : ''
  };
}

/** d = {edificio, unidad, mesTexto ("Setiembre 2026"), total ("S/ 471.99"), link, pago} */
function asunto(d) { return 'Recibo de Gastos Comunes - ' + d.mesTexto + ' - Dpto ' + d.unidad; }

function html(d) {
  var f = 'font-family:Arial,Helvetica,sans-serif;';
  var p = d.pago || {};
  function fila(contenido, estilo) {
    return '<tr><td align="center" style="text-align:center;' + f + (estilo || '') + '">' + contenido + '</td></tr>';
  }
  var cuentas = '';
  if (p.cuenta) cuentas += fila('<b>Cuenta' + (p.banco ? ' ' + esc(p.banco) : '') + ':</b> ' + esc(p.cuenta), 'font-size:13px;color:#222222;padding:6px 0;');
  if (p.cci) cuentas += fila('<b>CCI:</b> ' + esc(p.cci), 'font-size:13px;color:#222222;padding:6px 0;');
  if (p.titular) cuentas += fila('<b>Titular de la Cuenta:</b> ' + esc(p.titular), 'font-size:13px;color:#222222;padding:6px 0;');
  var boton = d.link
    ? '<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;border-collapse:separate;width:auto;display:inline-table;">' +
        '<tr><td align="center" bgcolor="' + COLOR + '" style="background:' + COLOR + ';border-radius:10px;">' +
          '<a href="' + esc(d.link) + '" target="_blank" style="display:inline-block;padding:18px 56px;' + f +
          'font-size:22px;font-weight:bold;letter-spacing:1px;color:#ffffff;text-decoration:none;border-radius:10px;">VER RECIBO</a>' +
        '</td></tr></table>'
    : '<span style="' + f + 'font-size:13px;color:#999999;">(el botón VER RECIBO aparece al conectar Google Drive)</span>';
  return '<div style="' + f + 'color:#222222;font-size:13px;max-width:980px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;width:100%;">' +
      fila('<b>' + esc(d.edificio) + '</b>', 'font-size:20px;color:#111111;padding:10px 0 0;') +
      fila('Hola, (Dpto ' + esc(d.unidad) + ')', 'font-size:13px;color:#222222;padding:16px 0 0;') +
    '</table>' +
    '<hr style="border:0;border-top:2px solid ' + COLOR + ';height:0;margin:24px 0 16px;">' +
    '<p style="' + f + 'font-size:13px;color:#222222;margin:0 0 18px;">Se ha generado tu recibo de gastos comunes correspondiente al mes de <b>' + esc(d.mesTexto) + '</b>.</p>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;width:100%;">' +
      fila('Total a pagar: ' + esc(d.total), 'font-size:24px;font-weight:bold;color:' + COLOR + ';padding:4px 0 18px;') +
      fila(boton, 'padding:4px 0 22px;') +
      cuentas +
    '</table>' +
    '<hr style="border:0;border-top:1px solid #cccccc;height:0;margin:16px 0;">' +
    (p.whatsapp ? '<p style="' + f + 'font-size:13px;color:#222222;text-align:center;margin:0 0 8px;"><b>Importante:</b> Enviar constancia de pago al WhatsApp: <b>' + esc(p.whatsapp) + '</b></p>' : '') +
  '</div>';
}

function texto(d) {
  var p = d.pago || {};
  var t = d.edificio + '\n\nHola, (Dpto ' + d.unidad + ')\n\n' +
    'Se ha generado tu recibo de gastos comunes correspondiente al mes de ' + d.mesTexto + '.\n\n' +
    'Total a pagar: ' + d.total + '\n\n';
  if (d.link) t += 'VER RECIBO: ' + d.link + '\n\n';
  if (p.cuenta) t += 'Cuenta' + (p.banco ? ' ' + p.banco : '') + ': ' + p.cuenta + '\n';
  if (p.cci) t += 'CCI: ' + p.cci + '\n';
  if (p.titular) t += 'Titular de la Cuenta: ' + p.titular + '\n';
  if (p.whatsapp) t += '\nImportante: Enviar constancia de pago al WhatsApp: ' + p.whatsapp + '\n';
  return t;
}

/** Mensaje de WhatsApp (texto con *negritas*). */
function textoWhatsapp(d) {
  var p = d.pago || {};
  var t = 'Hola, Dpto ' + d.unidad + ' 👋\n' +
    'Se ha generado tu recibo de gastos comunes de *' + d.mesTexto + '* (' + d.edificio + ').\n\n' +
    '*Total a pagar: ' + d.total + '*\n\n';
  if (d.link) t += 'Puedes ver tu recibo aquí 👉 ' + d.link + '\n\n';
  if (p.cuenta) t += '*Cuenta' + (p.banco ? ' ' + p.banco : '') + ':* ' + p.cuenta + '\n';
  if (p.cci) t += '*CCI:* ' + p.cci + '\n';
  if (p.titular) t += '*Titular de la Cuenta:* ' + p.titular + '\n';
  if (p.whatsapp) t += '\n*Importante:* Enviar constancia de pago al WhatsApp: *' + p.whatsapp + '*';
  return t.replace(/\n+$/, '');
}

/** Gmail en el navegador: destinatario y asunto llenos; el cuerpo se pega (va con formato). */
function urlGmail(correos, su) {
  return 'https://mail.google.com/mail/?view=cm&fs=1' +
    '&to=' + encodeURIComponent((correos || []).join(',')) +
    '&su=' + encodeURIComponent(su);
}
/** Cualquier app de correo (Gmail del celular, Outlook, Apple Mail…): cuerpo en texto simple. */
function urlMailto(correos, su, cuerpo) {
  return 'mailto:' + (correos || []).map(encodeURIComponent).join(',').replace(/%40/g, '@') +
    '?subject=' + encodeURIComponent(su) + '&body=' + encodeURIComponent(cuerpo || '');
}
/** WhatsApp: Web en computadora, wa.me en celular. */
function urlWhatsapp(tel, msg, movil) {
  var base = movil ? 'https://wa.me/' + tel + '?text=' : 'https://web.whatsapp.com/send?phone=' + tel + '&text=';
  return base + encodeURIComponent(msg);
}

/** o = {cfg, unidad, mesTexto, total, correos (texto), telefono, link, movil} */
function armar(cfg, unidad, mesTexto, total, correosTxt, extra) {
  extra = extra || {};
  var d = { edificio: nombreEdificio(cfg), unidad: String(unidad), mesTexto: mesTexto, total: total,
            link: extra.link || '', pago: datosPago(cfg) };
  var to = listaCorreos(correosTxt);
  var su = asunto(d);
  var tel = telefonoWhatsapp(extra.telefono);
  var wa = textoWhatsapp(d);
  return { datos: d, correos: to, asunto: su, html: html(d), texto: texto(d),
           gmail: urlGmail(to, su), mailto: urlMailto(to, su, texto(d)),
           telefono: tel, whatsappTexto: wa, whatsapp: tel ? urlWhatsapp(tel, wa, !!extra.movil) : '' };
}

var api = { COLOR: COLOR, nombreEdificio: nombreEdificio, listaCorreos: listaCorreos, asunto: asunto,
  html: html, texto: texto, textoWhatsapp: textoWhatsapp, telefonoWhatsapp: telefonoWhatsapp, telefonoBonito: telefonoBonito,
  datosPago: datosPago, urlGmail: urlGmail, urlMailto: urlMailto, urlWhatsapp: urlWhatsapp, armar: armar };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else global.Correo = api;
})(this);
