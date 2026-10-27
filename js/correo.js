/*
 * correo.js — Arma el correo de cada departamento (asunto, cuerpo HTML y texto) y los enlaces
 * para abrir Gmail o la app de correo. No envía nada: la persona revisa, pega y envía.
 * Navegador: window.Correo. Node: module.exports.
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

/** d = {edificio, unidad, mesTexto ("Setiembre 2026"), total ("S/ 471.99")} */
function asunto(d) { return 'Recibo de Gastos Comunes - ' + d.mesTexto + ' - Dpto ' + d.unidad; }

function html(d) {
  var f = 'font-family:Arial,Helvetica,sans-serif;';
  return '<div style="' + f + 'color:#222222;font-size:13px;max-width:980px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;width:100%;">' +
      '<tr><td style="width:50%;"></td>' +
      '<td align="center" style="width:50%;text-align:center;padding:10px 0 0;' + f + '">' +
        '<div style="' + f + 'font-size:20px;font-weight:bold;color:#111111;">' + esc(d.edificio) + '</div>' +
        '<div style="' + f + 'font-size:13px;color:#222222;margin-top:16px;">Hola, (Dpto ' + esc(d.unidad) + ')</div>' +
      '</td></tr></table>' +
    '<hr style="border:0;border-top:2px solid ' + COLOR + ';height:0;margin:24px 0 16px;">' +
    '<p style="' + f + 'font-size:13px;color:#222222;margin:0 0 18px;">Se ha generado tu recibo de gastos comunes correspondiente al mes de <b>' + esc(d.mesTexto) + '</b>.</p>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;width:100%;">' +
      '<tr><td style="width:50%;"></td>' +
      '<td align="center" style="width:50%;text-align:center;' + f + 'font-size:24px;font-weight:bold;color:' + COLOR + ';padding:4px 0 10px;">' +
        'Total a pagar: ' + esc(d.total) + '</td></tr></table>' +
  '</div>';
}

function texto(d) {
  return d.edificio + '\n\nHola, (Dpto ' + d.unidad + ')\n\n' +
    'Se ha generado tu recibo de gastos comunes correspondiente al mes de ' + d.mesTexto + '.\n\n' +
    'Total a pagar: ' + d.total + '\n';
}

/** Gmail en el navegador: destinatario y asunto llenos; el cuerpo se pega (va con formato). */
function urlGmail(correos, su) {
  return 'https://mail.google.com/mail/?view=cm&fs=1' +
    '&to=' + encodeURIComponent((correos || []).join(',')) +
    '&su=' + encodeURIComponent(su);
}
/** Cualquier app de correo (Outlook, Apple Mail…): cuerpo en texto simple. */
function urlMailto(correos, su, cuerpo) {
  return 'mailto:' + (correos || []).map(encodeURIComponent).join(',').replace(/%40/g, '@') +
    '?subject=' + encodeURIComponent(su) + '&body=' + encodeURIComponent(cuerpo || '');
}

function armar(cfg, unidad, mesTexto, total, correosTxt) {
  var d = { edificio: nombreEdificio(cfg), unidad: String(unidad), mesTexto: mesTexto, total: total };
  var to = listaCorreos(correosTxt);
  var su = asunto(d);
  return { datos: d, correos: to, asunto: su, html: html(d), texto: texto(d),
           gmail: urlGmail(to, su), mailto: urlMailto(to, su, texto(d)) };
}

var api = { COLOR: COLOR, nombreEdificio: nombreEdificio, listaCorreos: listaCorreos, asunto: asunto,
  html: html, texto: texto, urlGmail: urlGmail, urlMailto: urlMailto, armar: armar };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else global.Correo = api;
})(this);
