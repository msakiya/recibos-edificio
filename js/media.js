/*
 * media.js — Reduce fotos con canvas y convierte la 1ª página de un PDF a imagen (pdf.js).
 * Todo en el navegador. Navegador: window.Media. Node (solo reconocimiento de nombres): module.exports.
 */
(function (global) {
'use strict';
var C = (typeof module !== 'undefined' && module.exports) ? require('./calculo.js') : global.Calculo;
var MAX_LADO = 1600, CALIDAD = 0.85;

function esPdf(file) { return /pdf/i.test(file.type || '') || /\.pdf$/i.test(file.name || ''); }
function esHeic(file) { return /heic|heif/i.test(file.type || '') || /\.(heic|heif)$/i.test(file.name || ''); }
function esImagen(file) { return /^image\//i.test(file.type || '') || /\.(jpe?g|png|webp|gif)$/i.test(file.name || ''); }

/** ¿El nombre apunta a un recibo de luz o de agua del edificio? */
function tipoReciboEdificio(nombre) {
  var n = String(nombre || '').toLowerCase();
  if (/recibo[\s_-]*de[\s_-]*luz|luz[\s_-]*del[\s_-]*sur|recibo[\s_-]*luz|^luz[\s._-]/i.test(n)) return 'luz';
  if (/recibo[\s_-]*de[\s_-]*agua|sedapal|recibo[\s_-]*agua|^agua[\s._-]/i.test(n)) return 'agua';
  return null;
}
/** ¿Parece un comprobante de pago (no medidor)? */
function parecePago(nombre) {
  return /pago|comprobante|yape|transfer|voucher|deposito|depósito/i.test(String(nombre || ''));
}
/** Unidades cuyo número aparece en el nombre (sin confundir 101 con 1010 ni 201). */
function unidadesEnNombre(nombre, unidades) {
  var out = [];
  (unidades || []).forEach(function (u) { if (C.nombreEsDeUnidad(nombre, u)) out.push(String(u)); });
  return out;
}

function leerComoDataUrl(file) {
  return new Promise(function (ok, bad) {
    var r = new FileReader();
    r.onload = function () { ok(r.result); };
    r.onerror = function () { bad(new Error('No se pudo leer ' + file.name)); };
    r.readAsDataURL(file);
  });
}
function dataUrlABlob(dataUrl) {
  var m = dataUrl.match(/^data:([^;]+);base64,(.*)$/);
  if (!m) throw new Error('data URL inválida');
  var bin = atob(m[2]), u8 = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return new Blob([u8], { type: m[1] });
}

/** Reduce una imagen (File o data URL) a JPEG ≤ MAX_LADO px. Devuelve {dataUrl, blob, w, h}. */
function reducirImagen(entrada, maxLado) {
  maxLado = maxLado || MAX_LADO;
  return Promise.resolve().then(function () {
    if (typeof entrada === 'string') return entrada;
    return leerComoDataUrl(entrada);
  }).then(function (src) {
    return new Promise(function (ok, bad) {
      var img = new Image();
      img.onload = function () {
        var w = img.naturalWidth, h = img.naturalHeight;
        var s = Math.min(1, maxLado / Math.max(w, h));
        var cw = Math.max(1, Math.round(w * s)), ch = Math.max(1, Math.round(h * s));
        var canvas = document.createElement('canvas');
        canvas.width = cw; canvas.height = ch;
        var ctx = canvas.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cw, ch);
        ctx.drawImage(img, 0, 0, cw, ch);
        var dataUrl = canvas.toDataURL('image/jpeg', CALIDAD);
        ok({ dataUrl: dataUrl, blob: dataUrlABlob(dataUrl), w: cw, h: ch });
      };
      img.onerror = function () { bad(new Error('No se pudo abrir la imagen')); };
      img.src = src;
    });
  });
}

/** Primera página de un PDF (File) → JPEG. Usa pdfjsLib global. */
function pdfAImagen(file, maxLado) {
  maxLado = maxLado || MAX_LADO;
  if (!global.pdfjsLib) return Promise.reject(new Error('pdf.js no está cargado'));
  return file.arrayBuffer().then(function (buf) {
    return global.pdfjsLib.getDocument({ data: buf }).promise;
  }).then(function (pdf) {
    return pdf.getPage(1).then(function (page) {
      var base = page.getViewport({ scale: 1 });
      var scale = Math.min(2.2, maxLado / Math.max(base.width, base.height));
      var vp = page.getViewport({ scale: scale });
      var canvas = document.createElement('canvas');
      canvas.width = vp.width; canvas.height = vp.height;
      return page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise.then(function () {
        var dataUrl = canvas.toDataURL('image/jpeg', CALIDAD);
        return { dataUrl: dataUrl, blob: dataUrlABlob(dataUrl), w: canvas.width, h: canvas.height };
      });
    });
  });
}

/** Procesa un File: si es PDF → imagen de la pág. 1; si es imagen → JPEG reducido. HEIC se rechaza. */
function aJpeg(file, maxLado) {
  if (esHeic(file)) return Promise.reject(new Error('Las fotos HEIC del iPhone no se pueden leer aquí. Compártelas como JPG o haz una captura de pantalla.'));
  if (esPdf(file)) return pdfAImagen(file, maxLado);
  if (esImagen(file)) return reducirImagen(file, maxLado);
  return Promise.reject(new Error('Formato no soportado: ' + (file.name || file.type) + '. Usa JPG, PNG o PDF.'));
}

/**
 * Clasifica una lista de Files arrastrados.
 * Devuelve {asignados:[{file,tipo,unidad,sugerido}], sinReconocer:[{file}]}.
 * tipo ∈ 'medidor'|'pago'|'luz'|'agua'.
 */
function clasificar(files, unidades) {
  var asignados = [], sin = [];
  Array.prototype.forEach.call(files, function (f) {
    var rec = tipoReciboEdificio(f.name);
    if (rec) { asignados.push({ file: f, tipo: rec, unidad: '', sugerido: true }); return; }
    var us = unidadesEnNombre(f.name, unidades);
    if (us.length === 1) {
      asignados.push({ file: f, tipo: parecePago(f.name) ? 'pago' : 'medidor', unidad: us[0], sugerido: true });
      return;
    }
    if (us.length > 1) { sin.push({ file: f, candidatos: us }); return; }
    sin.push({ file: f, candidatos: [] });
  });
  return { asignados: asignados, sinReconocer: sin };
}

var api = { MAX_LADO: MAX_LADO, esPdf: esPdf, esHeic: esHeic, esImagen: esImagen,
  tipoReciboEdificio: tipoReciboEdificio, parecePago: parecePago, unidadesEnNombre: unidadesEnNombre,
  clasificar: clasificar, reducirImagen: reducirImagen, pdfAImagen: pdfAImagen, aJpeg: aJpeg,
  leerComoDataUrl: leerComoDataUrl };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else global.Media = api;
})(this);
