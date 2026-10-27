/*
 * pdf.js — Genera el PDF A4 de 3 páginas de un departamento (html2canvas + jsPDF).
 * Requiere en el navegador: jsPDF, html2canvas, ReciboHtml, Calculo.
 */
(function (global) {
'use strict';
var C = global.Calculo, RH = global.ReciboHtml;
var A4_MM = { w: 210, h: 297 };
var ESCALA = 2; // 2× = ~150 dpi; suficiente y más liviano en el celular

function esperarImagenes(root) {
  var imgs = Array.prototype.slice.call(root.querySelectorAll('img'));
  return Promise.all(imgs.map(function (img) {
    if (img.complete && img.naturalWidth) return null;
    return new Promise(function (ok) {
      img.onload = img.onerror = function () { ok(null); };
    });
  }));
}
function mmAPx(mm) { return Math.round(mm * 96 / 25.4); }

/** Renderiza un div .pagina a JPEG (data URL) con html2canvas. */
function capturarPagina(el) {
  return global.html2canvas(el, {
    scale: ESCALA, backgroundColor: '#ffffff', useCORS: true, allowTaint: true,
    logging: false, width: el.offsetWidth, height: el.offsetHeight,
    windowWidth: el.offsetWidth, windowHeight: el.offsetHeight
  }).then(function (canvas) {
    return canvas.toDataURL('image/jpeg', 0.92);
  });
}

/**
 * Genera el PDF de UNA unidad.
 * imgs = {logo, foto, agua, luz} (data URLs).
 * Devuelve {nombre, blob, bytes}.
 */
function generarUno(res, unidad, cfg, imgs) {
  var u = res.unidades.filter(function (x) { return C.mismaUnidad(x.unidad, unidad); })[0];
  if (!u) return Promise.reject(new Error('No está la unidad ' + unidad));
  var d = C.datosPlantilla(res, u, cfg, imgs || {});
  var host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-10000px;top:0;width:210mm;background:#fff;z-index:-1';
  host.innerHTML = RH.pagina1(d) + RH.pagina2(d) + RH.pagina3(d);
  // Inyecta el CSS una sola vez
  if (!document.getElementById('recibo-css')) {
    var st = document.createElement('style'); st.id = 'recibo-css'; st.textContent = RH.CSS; document.head.appendChild(st);
  }
  document.body.appendChild(host);
  return esperarImagenes(host).then(function () {
    // Forzar layout a mm exactos
    var pags = host.querySelectorAll('.pagina');
    Array.prototype.forEach.call(pags, function (p) {
      p.style.width = A4_MM.w + 'mm';
      p.style.minHeight = A4_MM.h + 'mm';
      p.style.height = A4_MM.h + 'mm';
      p.style.overflow = 'hidden';
    });
    var cadena = Promise.resolve([]);
    Array.prototype.forEach.call(pags, function (p) {
      cadena = cadena.then(function (arr) {
        return capturarPagina(p).then(function (img) { arr.push(img); return arr; });
      });
    });
    return cadena;
  }).then(function (imgsPag) {
    document.body.removeChild(host);
    var JsPDF = (global.jspdf && global.jspdf.jsPDF) || global.jsPDF;
    var doc = new JsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
    imgsPag.forEach(function (img, i) {
      if (i) doc.addPage();
      doc.addImage(img, 'JPEG', 0, 0, A4_MM.w, A4_MM.h);
    });
    var bytes = doc.output('arraybuffer');
    var blob = new Blob([bytes], { type: 'application/pdf' });
    return { nombre: C.nombrePdf(res, unidad), blob: blob, bytes: bytes, total: u.total, propietario: u.propietario, unidad: String(u.unidad) };
  }).catch(function (e) {
    if (host.parentNode) document.body.removeChild(host);
    throw e;
  });
}

/**
 * Genera los PDF de todas las unidades + el Excel actualizado, empaquetados en un ZIP.
 * onProgreso(texto, pct 0-1).
 * Devuelve {blob, nombre, resumen:[{unidad,total,nombre}]}.
 */
function generarZip(datos, res, archivos, excelBytes, onProgreso) {
  onProgreso = onProgreso || function () {};
  var JsZip = global.JSZip;
  if (!JsZip) return Promise.reject(new Error('JSZip no está cargado'));
  var zip = new JsZip();
  var cfg = datos.config || {};
  var comun = { logo: (archivos && archivos.logo) || datos.logo || '', agua: (archivos && archivos.agua) || '', luz: (archivos && archivos.luz) || '' };
  var i = 0, n = res.unidades.length, resumen = [];
  function siguiente() {
    if (i >= n) {
      onProgreso('Empaquetando ZIP…', 0.95);
      if (excelBytes) zip.file('Base de datos - ' + res.nombreMes + ' ' + res.anio + '.xlsx', excelBytes);
      return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } })
        .then(function (blob) {
          onProgreso('Listo', 1);
          return { blob: blob, nombre: 'Recibos - ' + res.nombreMes + ' ' + res.anio + '.zip', resumen: resumen };
        });
    }
    var u = res.unidades[i++];
    onProgreso('Generando recibo ' + u.unidad + ' (' + i + '/' + n + ')…', (i - 1) / (n + 1));
    var imgs = { logo: comun.logo, agua: comun.agua, luz: comun.luz,
                 foto: (archivos && archivos.medidor && archivos.medidor[u.unidad]) || '' };
    return generarUno(res, u.unidad, cfg, imgs).then(function (pdf) {
      zip.file(pdf.nombre, pdf.bytes);
      resumen.push({ unidad: pdf.unidad, propietario: pdf.propietario, total: pdf.total, nombre: pdf.nombre, blob: pdf.blob });
      return siguiente();
    });
  }
  return siguiente();
}

var api = { generarUno: generarUno, generarZip: generarZip };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else global.PdfGen = api;
})(this);
