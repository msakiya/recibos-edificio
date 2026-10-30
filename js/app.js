/* app.js — Asistente de 3 pasos. Todo en el navegador. */
(function () {
'use strict';
Excel.usar(XLSX);

var estado = {
  paso: 1,
  wb: null,           // libro SheetJS
  datos: null,        // datos leídos
  nombreArchivo: '',
  borrador: null,     // prepararBorrador
  logoDataUrl: '',
  archivos: { medidor: {}, pago: {}, agua: '', luz: '' }, // data URLs
  cola: []            // [{id,file,tipo,unidad,preview,status}]
};

function $(id) { return document.getElementById(id); }
function mostrar(id, si) { $(id).classList.toggle('oculto', !si); }
function esc(s) { return ReciboHtml.esc(s); }

function irA(paso) {
  if (paso > 1 && !estado.borrador) return;
  estado.paso = paso;
  [1, 3].forEach(function (n) {
    mostrar('paso' + n, n === paso);
    var t = $('tab' + n);
    t.classList.toggle('activo', n === paso);
    t.classList.toggle('hecho', n < paso);
    t.disabled = n > 1 && !estado.borrador;
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ---------------- drop helpers ---------------- */
function setupDrop(zonaId, inputId, onFiles) {
  var zona = $(zonaId), input = $(inputId);
  function abrir() { input.click(); }
  zona.addEventListener('click', abrir);
  zona.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); abrir(); } });
  input.addEventListener('change', function () { if (input.files && input.files.length) onFiles(input.files); input.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) {
    zona.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); zona.classList.add('sobre'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    zona.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); zona.classList.remove('sobre'); });
  });
  zona.addEventListener('drop', function (e) { if (e.dataTransfer.files.length) onFiles(e.dataTransfer.files); });
}

/* ---------------- Paso 1: Excel ---------------- */
function cargarExcel(files) {
  var f = files[0];
  if (!f) return;
  if (!/\.xlsx$/i.test(f.name) && f.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    alert('Elige un archivo .xlsx (Excel).');
    return;
  }
  var info = $('excelInfo');
  info.classList.remove('oculto');
  info.innerHTML = '<div class="banner info">Leyendo <b>' + esc(f.name) + '</b>…</div>';
  f.arrayBuffer().then(function (buf) {
    var wb = XLSX.read(buf, { type: 'array', cellDates: true });
    var L = Excel.leerLibro(wb);
    if (L.errores.length) {
      info.innerHTML = '<div class="banner error"><b>No se pudo usar este Excel:</b><ul>' +
        L.errores.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul></div>';
      return;
    }
    estado.wb = wb; estado.datos = L.datos; estado.nombreArchivo = f.name;
    estado.logoDataUrl = L.datos.logo || '';
    estado.archivos = { medidor: {}, pago: {}, agua: '', luz: '' };
    estado.cola = [];
    estado.borrador = null;
    var M = Excel.cargarMes(wb, L.datos);
    var cab = '<div class="banner ok">✔ Excel cargado: <b>' + esc(f.name) + '</b> · ' +
      esc(L.datos.config.EDIFICIO_NOMBRE || '(sin nombre)') + ' · ' + L.datos.deptos.length + ' departamentos</div>';
    var avisos = L.avisos.concat(M.avisos);
    var htmlAvisos = avisos.length ? '<div class="banner aviso"><ul>' + avisos.map(function (a) { return '<li>' + esc(a) + '</li>'; }).join('') + '</ul></div>' : '';
    if (M.errores.length) {
      info.innerHTML = cab + '<div class="banner error"><b>Falta completar el Excel.</b> Corrige esto en tu Excel, guárdalo y vuelve a subirlo:<ul>' +
        M.errores.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul></div>' + htmlAvisos +
        (M.sinHoja ? '<div class="botones"><button type="button" class="boton primario" id="btnConHoja">⬇ Descargar mi Excel con la pestaña “Mes a facturar”</button></div>' : '');
      if (M.sinHoja) $('btnConHoja').onclick = function () { descargarConHojaMes(M.mesSugerido || Excel.hoyISO().slice(0, 7)); };
      irA(1);
      return;
    }
    estado.borrador = M.borrador;
    var ap = Excel.aplicarBorrador(estado.datos, estado.borrador);
    info.innerHTML = cab + htmlAvisos +
      '<h3 class="sub-titulo">Revisa los datos de ' + esc(ap.res.nombreMes + ' ' + ap.res.anio) + ' (leídos del Excel)</h3>' +
      '<p class="mini">Si algo está mal, corrígelo en la pestaña “Mes a facturar” de tu Excel y vuelve a subirlo.</p>' +
      htmlResumen(estado.borrador, ap.res) +
      '<div class="botones"><button type="button" class="boton primario" id="btnAPaso3b">Todo bien → Subir fotos y generar</button></div>';
    pintarLogo();
    $('btnAPaso3b').onclick = function () { irA(3); pintarEstadoArchivos(); };
    irA(1);
  }).catch(function (e) {
    info.innerHTML = '<div class="banner error">No se pudo leer el Excel: ' + esc(e.message || e) + '</div>';
  });
}
setupDrop('dropExcel', 'fileExcel', cargarExcel);

// "Excel de ejemplo" y "Excel en blanco" son archivos fijos del sitio (plantilla-ejemplo.xlsx / plantilla-en-blanco.xlsx), con formato.


/* ---------------- Resumen del mes (solo lectura, todo sale del Excel) ---------------- */
$('tab1').onclick = function () { irA(1); };
$('tab3').onclick = function () { if (estado.borrador) irA(3); };
$('btnVolver2').onclick = function () { irA(1); };

$('logoFile').onchange = function () {
  var f = this.files && this.files[0]; if (!f) return;
  Media.aJpeg(f, 400).then(function (r) {
    estado.logoDataUrl = r.dataUrl;
    $('logoPreview').innerHTML = '<img src="' + r.dataUrl + '" height="40" alt="logo"> Logo listo';
  }).catch(function (e) { alert(e.message || e); });
};
function pintarLogo() {
  if (estado.logoDataUrl) $('logoPreview').innerHTML = '<img src="' + estado.logoDataUrl + '" height="40" alt="logo"> Logo del Excel';
  else $('logoPreview').textContent = 'Sin logo (opcional)';
}

function fmtFechaIso(iso) { var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + '/' + m[2] + '/' + m[1] : '—'; }

function htmlResumen(b, res) {
  var S = Calculo.fmtSoles, N = Calculo.fmtNum;
  var filas = res.unidades.map(function (u) {
    return '<tr><td><b>' + esc(u.unidad) + '</b></td><td>' + esc(u.propietario) + '</td>' +
      '<td class="m">' + N(u.anterior, 3) + '</td><td class="m">' + N(u.actual, 3) + '</td><td class="m">' + N(u.consumo, 3) + '</td>' +
      '<td class="m">' + S(u.aguaPrivadaFinal) + '</td><td class="m">' + S(u.comun) + '</td>' +
      '<td class="m">' + (u.totalAjustes ? S(u.totalAjustes) : '') + '</td><td class="m"><b>' + S(u.total) + '</b></td></tr>';
  }).join('');
  var gastos = res.gastos.map(function (g) {
    return '<tr><td>' + esc(g.concepto) + (g.tipo === 'EXTRA' ? ' <span class="pill aviso">extra</span>' : '') + '</td><td class="m">' + S(g.monto) + '</td><td class="m">' + S(g.porUnidad) + '</td></tr>';
  }).join('');
  return '<div class="resumen-cab">' +
      '<div class="dato"><span>Mes</span><b>' + esc(res.nombreMes + ' ' + res.anio) + '</b></div>' +
      '<div class="dato"><span>Agua (recibo)</span><b>' + S(res.agua.importe) + '</b><small>' + N(res.agua.m3, 0) + ' m³ · S/ ' + N(res.agua.precioM3, 4) + ' por m³</small></div>' +
      '<div class="dato"><span>Luz (recibo)</span><b>' + S(res.luz.importe) + '</b></div>' +
      '<div class="dato"><span>Emisión · Vence</span><b>' + fmtFechaIso(b.servicios.emision) + '</b><small>vence ' + fmtFechaIso(b.servicios.vencimiento) + '</small></div>' +
      '<div class="dato"><span>Total del mes</span><b>' + S(res.totalFacturado) + '</b></div>' +
      '<div class="dato"><span>Fondo de contingencia</span><b>' + S(res.fondo.saldo) + '</b><small>antes ' + S(res.fondo.saldoAnterior) + '</small></div>' +
    '</div>' +
    '<h3 class="sub-titulo">Lecturas y total por departamento</h3>' +
    '<div class="tabla-wrap"><table class="datos resumen-tabla"><thead><tr><th>Dpto</th><th>Propietario</th><th class="m">Anterior</th><th class="m">Actual</th><th class="m">Consumo m³</th>' +
    '<th class="m">Agua</th><th class="m">Común</th><th class="m">Ajuste</th><th class="m">Total</th></tr></thead><tbody>' + filas + '</tbody></table></div>' +
    '<p class="mini">Agua común: ' + N(res.agua.m3Comun, 3) + ' m³ = ' + S(res.agua.comun) + (res.agua.unidadAjuste ? ' · redondeo de ' + S(res.agua.diferenciaRedondeo) + ' al Dpto ' + esc(res.agua.unidadAjuste) : '') + '</p>' +
    '<h3 class="sub-titulo">Gastos comunes</h3>' +
    '<div class="tabla-wrap"><table class="datos resumen-tabla gastos-tabla"><thead><tr><th>Concepto</th><th class="m">Total edificio</th><th class="m">Por dpto</th></tr></thead><tbody>' + gastos +
    '<tr><td><b>Total por dpto (común)</b></td><td class="m"><b>' + S(res.totalComun) + '</b></td><td class="m"><b>' + S(res.comunPorUnidad) + '</b></td></tr></tbody></table></div>';
}

/** Excel sin pestaña "Mes a facturar": descarga el mismo Excel con la pestaña agregada (para llenarla). */
function descargarConHojaMes(mes) {
  Excel.escribirHojaMes(estado.wb, estado.datos, mes);
  var base = estado.nombreArchivo.replace(/\.xlsx$/i, '');
  XLSX.writeFile(estado.wb, base + ' (con Mes a facturar).xlsx');
}

/* ---------------- Paso 3: fotos ---------------- */
var uid = 0;
setupDrop('dropFotos', 'fileFotos', function (files) { agregarArchivos(files); });

function agregarArchivos(files) {
  var unidades = estado.datos.deptos.map(function (d) { return d.unidad; });
  var clas = Media.clasificar(files, unidades);
  var pendientes = clas.asignados.map(function (a) {
    return { file: a.file, tipo: a.tipo, unidad: a.unidad };
  }).concat(clas.sinReconocer.map(function (s) {
    return { file: s.file, tipo: '', unidad: '', candidatos: s.candidatos };
  }));
  pendientes.forEach(function (p) {
    var item = { id: ++uid, file: p.file, tipo: p.tipo || '', unidad: p.unidad || '', preview: '', status: 'procesando', error: '' };
    estado.cola.push(item);
    pintarCola();
    Media.aJpeg(p.file).then(function (r) {
      item.preview = r.dataUrl; item.status = 'ok';
      aplicarAsignacion(item);
      pintarCola(); pintarEstadoArchivos();
    }).catch(function (e) {
      item.status = 'error'; item.error = e.message || String(e);
      pintarCola();
    });
  });
}

function aplicarAsignacion(item) {
  // Limpia asignaciones previas de este id
  Object.keys(estado.archivos.medidor).forEach(function (u) {
    if (estado.archivos.medidor[u + '__id'] === item.id) { delete estado.archivos.medidor[u]; delete estado.archivos.medidor[u + '__id']; }
  });
  Object.keys(estado.archivos.pago).forEach(function (u) {
    if (estado.archivos.pago[u + '__id'] === item.id) { delete estado.archivos.pago[u]; delete estado.archivos.pago[u + '__id']; }
  });
  if (estado.archivos.aguaId === item.id) { estado.archivos.agua = ''; estado.archivos.aguaId = null; }
  if (estado.archivos.luzId === item.id) { estado.archivos.luz = ''; estado.archivos.luzId = null; }
  if (!item.preview || item.status !== 'ok') return;
  if (item.tipo === 'agua') { estado.archivos.agua = item.preview; estado.archivos.aguaId = item.id; }
  else if (item.tipo === 'luz') { estado.archivos.luz = item.preview; estado.archivos.luzId = item.id; }
  else if (item.tipo === 'medidor' && item.unidad) { estado.archivos.medidor[item.unidad] = item.preview; estado.archivos.medidor[item.unidad + '__id'] = item.id; }
  else if (item.tipo === 'pago' && item.unidad) { estado.archivos.pago[item.unidad] = item.preview; estado.archivos.pago[item.unidad + '__id'] = item.id; }
}

function pintarCola() {
  var box = $('listaArchivos');
  var optsTipo = [['', '— elegir —'], ['medidor', 'Foto medidor'], ['pago', 'Comprobante pago'], ['agua', 'Recibo de agua'], ['luz', 'Recibo de luz']];
  var optsUnidad = [['', '— dpto —']].concat(estado.datos.deptos.map(function (d) { return [String(d.unidad), String(d.unidad)]; }));
  box.innerHTML = estado.cola.map(function (it) {
    var necesitaUnidad = (it.tipo === 'medidor' || it.tipo === 'pago');
    return '<div class="archivo-card" data-id="' + it.id + '">' +
      (it.preview ? '<img src="' + it.preview + '" alt="">' : '<div style="width:56px;height:56px;background:#eee;border-radius:6px"></div>') +
      '<div class="meta"><div class="nombre">' + esc(it.file.name) + '</div>' +
      '<div class="sub">' + (it.status === 'procesando' ? 'Procesando…' : it.status === 'error' ? esc(it.error) : 'Listo') + '</div></div>' +
      '<select class="tipo">' + optsTipo.map(function (o) { return '<option value="' + o[0] + '"' + (it.tipo === o[0] ? ' selected' : '') + '>' + o[1] + '</option>'; }).join('') + '</select>' +
      '<select class="unidad"' + (necesitaUnidad ? '' : ' disabled') + '>' +
        optsUnidad.map(function (o) { return '<option value="' + o[0] + '"' + (String(it.unidad) === o[0] ? ' selected' : '') + '>' + o[1] + '</option>'; }).join('') +
      '</select>' +
      '<button type="button" class="x" title="Quitar">×</button></div>';
  }).join('');
  box.querySelectorAll('.archivo-card').forEach(function (card) {
    var id = +card.getAttribute('data-id');
    var it = estado.cola.filter(function (x) { return x.id === id; })[0];
    card.querySelector('.tipo').onchange = function () { it.tipo = this.value; if (it.tipo === 'agua' || it.tipo === 'luz') it.unidad = ''; aplicarAsignacion(it); pintarCola(); pintarEstadoArchivos(); };
    card.querySelector('.unidad').onchange = function () { it.unidad = this.value; aplicarAsignacion(it); pintarEstadoArchivos(); };
    card.querySelector('.x').onclick = function () {
      estado.cola = estado.cola.filter(function (x) { return x.id !== id; });
      // limpia
      var dummy = { id: id, preview: '', status: 'ok', tipo: '', unidad: '' };
      aplicarAsignacion(dummy);
      pintarCola(); pintarEstadoArchivos();
    };
  });
}

function pintarEstadoArchivos() {
  var b = estado.borrador;
  var conFoto = b.lecturas.filter(function (l) { return estado.archivos.medidor[l.unidad]; }).length;
  var html = 'Medidores: <b>' + conFoto + '/' + b.lecturas.length + '</b> con foto. · ' +
    'Recibo de agua: ' + (estado.archivos.agua ? '<span class="pill ok">✔</span>' : '<span class="pill falta">falta</span>') + ' · ' +
    'Recibo de luz: ' + (estado.archivos.luz ? '<span class="pill ok">✔</span>' : '<span class="pill falta">falta</span>');
  $('estadoArchivos').innerHTML = html;
}

/* ---------------- Generar ---------------- */
$('btnGenerar').onclick = function () {
  var aplicado = Excel.aplicarBorrador(estado.datos, estado.borrador);
  var rev = Excel.revisar(estado.borrador, aplicado.res, {
    medidor: estado.archivos.medidor, agua: estado.archivos.agua, luz: estado.archivos.luz
  });
  var panel = $('panelRevision');
  panel.classList.remove('oculto');
  mostrar('panelResultado', false);
  mostrar('panelProgreso', false);
  var html = '';
  if (rev.errores.length) {
    html += '<div class="banner error"><b>Hay errores. Corrige esto antes de generar:</b><ul>' +
      rev.errores.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul></div>';
    panel.innerHTML = html;
    return;
  }
  if (rev.avisos.length) {
    html += '<div class="banner aviso"><b>Avisos (no impiden generar):</b><ul>' +
      rev.avisos.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul>' +
      '<p>¿Generar los recibos de <b>' + esc(aplicado.res.nombreMes + ' ' + aplicado.res.anio) + '</b> de todas formas?</p>' +
      '<div class="botones"><button type="button" class="boton naranja" id="btnSiGen">Sí, generar recibos</button>' +
      '<button type="button" class="boton sec" id="btnNoGen">No, voy a corregir</button></div></div>';
    panel.innerHTML = html;
    $('btnSiGen').onclick = function () { panel.classList.add('oculto'); ejecutarGeneracion(aplicado); };
    $('btnNoGen').onclick = function () { panel.classList.add('oculto'); };
    return;
  }
  panel.classList.add('oculto');
  ejecutarGeneracion(aplicado);
};

function esMovil() {
  return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent || '') ||
    (window.matchMedia && window.matchMedia('(pointer: coarse)').matches && Math.min(screen.width, screen.height) < 820);
}

var gen = null;   // resultado de la última generación: {res, pdfs, excelBytes, excelNombre, links, carpetaId, subiendo, errorDrive}

function ejecutarGeneracion(aplicado) {
  // Si Drive está conectado, pide el token YA (dentro del clic) para que Google no bloquee la ventanita.
  var dEst = Drive.estado();
  if (dEst.conectado && Drive.configurado() && !Drive.tokenValido()) Drive.pedirToken(false).catch(function () {});
  mostrar('panelProgreso', true);
  mostrar('panelResultado', false);
  $('btnGenerar').disabled = true;
  var res = aplicado.res;
  var wb = estado.wb;
  if (estado.logoDataUrl) {
    Excel.ponerLogo(wb, estado.logoDataUrl);
    aplicado.datos.logo = estado.logoDataUrl;
  }
  Excel.escribirMes(wb, aplicado.datos, res.mes);
  // Deja la pestaña "Mes a facturar" lista para el mes siguiente (lecturas anteriores, gastos fijos, saldo del fondo)
  var sig = Calculo.mesSiguiente(res.mes);
  Excel.escribirHojaMes(wb, aplicado.datos, sig);
  var excelBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  var archivos = {
    logo: estado.logoDataUrl || aplicado.datos.logo || '',
    medidor: estado.archivos.medidor,
    agua: estado.archivos.agua,
    luz: estado.archivos.luz
  };
  PdfGen.generarTodos(aplicado.datos, res, archivos, function (txt, pct) {
    $('txtProgreso').textContent = txt;
    $('barraProgreso').style.width = Math.round(pct * 100) + '%';
  }).then(function (pdfs) {
    $('btnGenerar').disabled = false;
    mostrar('panelProgreso', false);
    mostrar('panelResultado', true);
    estado.datos = aplicado.datos;
    gen = { res: res, pdfs: pdfs, excelBytes: excelBytes, links: {}, carpetaId: '', subiendo: false, errorDrive: '', pctDrive: 0, txtDrive: '',
            excelNombre: 'Excel actualizado - para ' + Calculo.nombreMes(sig) + ' ' + Calculo.partesMes(sig).anio + '.xlsx',
            sigTexto: Calculo.nombreMes(sig) + ' ' + Calculo.partesMes(sig).anio };
    enviados = {};
    pintarResultado();
    if (Drive.estado().conectado && Drive.configurado()) subirADrive();
  }).catch(function (e) {
    $('btnGenerar').disabled = false;
    mostrar('panelProgreso', false);
    mostrar('panelResultado', true);
    $('panelResultado').innerHTML = '<div class="banner error">Error al generar: ' + esc(e.message || e) + '</div>';
    console.error(e);
  });
}

function descargarBlob(blob, nombre) {
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = nombre;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function () { URL.revokeObjectURL(a.href); }, 30000);
}

function pintarResultado() {
  var res = gen.res;
  $('panelResultado').innerHTML =
    '<div class="banner ok"><b>✔ Listo:</b> ' + gen.pdfs.length + ' recibos de ' + esc(res.nombreMes + ' ' + res.anio) +
    ' · Total ' + esc(Calculo.fmtSoles(res.totalFacturado)) + ' · Fondo ' + esc(Calculo.fmtSoles(res.fondo.saldo)) + '</div>' +
    '<div class="excel-caja"><button type="button" class="boton primario" id="btnExcel">⬇ Descargar Excel actualizado (para el próximo mes)</button>' +
    '<p class="mini">Guárdalo en lugar del anterior: ya trae la pestaña “Mes a facturar” lista para <b>' + esc(gen.sigTexto) + '</b>. El próximo mes llénalo y súbelo aquí.</p></div>' +
    '<div id="panelDrive"></div>' +
    '<div id="panelEnvio"></div>';
  $('btnExcel').onclick = function () {
    descargarBlob(new Blob([gen.excelBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), gen.excelNombre);
  };
  pintarPanelDrive();
  pintarEnvio();
}

function pintarPanelDrive() {
  var box = $('panelDrive'); if (!box || !gen) return;
  var d = Drive.estado(), n = Object.keys(gen.links).length;
  var html;
  if (gen.subiendo) {
    html = '<div class="banner info">☁ ' + esc(gen.txtDrive || 'Subiendo a Google Drive…') + '<div class="progreso"><span id="barraDrive" style="width:' + Math.round(gen.pctDrive * 100) + '%"></span></div></div>';
  } else if (n && n === gen.pdfs.length) {
    html = '<div class="banner ok">☁ Los ' + n + ' recibos están en Google Drive (cualquiera con el link puede verlos). ' +
      (gen.carpetaId ? '<a href="https://drive.google.com/drive/folders/' + esc(gen.carpetaId) + '" target="_blank" rel="noopener">Abrir la carpeta</a>' : '') + '</div>';
  } else if (gen.errorDrive) {
    html = '<div class="banner error">No se pudieron subir los recibos a Google Drive: ' + esc(gen.errorDrive) +
      '<div class="botones"><button type="button" class="boton primario" id="btnSubirDrive">☁ Reintentar subida a Drive</button></div></div>';
  } else if (d.conectado && Drive.configurado()) {
    html = '<div class="banner info">☁ Drive conectado (' + esc(d.email || 'cuenta de Google') + ').' +
      '<div class="botones"><button type="button" class="boton primario" id="btnSubirDrive">☁ Subir recibos a Drive</button></div></div>';
  } else {
    html = '<div class="banner aviso"><b>Conecta Google Drive para enviar el link del recibo.</b> Los PDF se pueden descargar igual con <b>⬇ PDF</b>.' +
      '<div class="botones"><button type="button" class="boton primario" id="btnConectarRes">☁ Conectar Google Drive</button></div></div>';
  }
  box.innerHTML = html;
  if ($('btnSubirDrive')) $('btnSubirDrive').onclick = function () { Drive.pedirToken(false).catch(function () {}); subirADrive(); };
  if ($('btnConectarRes')) $('btnConectarRes').onclick = function () { conectarDrive(true); };
}

function subirADrive() {
  if (!gen || gen.subiendo) return;
  var res = gen.res;
  gen.subiendo = true; gen.errorDrive = ''; gen.pctDrive = 0; gen.txtDrive = 'Preparando carpetas en Google Drive…';
  pintarPanelDrive();
  var info = { edificio: Correo.nombreEdificio(estado.datos.config || {}), anio: res.anio,
               mesNum: Calculo.partesMes(res.mes).mes, mesNombre: res.nombreMes };
  return Drive.subirRecibos(gen.pdfs, info, function (txt, pct) {
    gen.txtDrive = txt; gen.pctDrive = pct;
    var b = $('barraDrive');
    if (b) { b.style.width = Math.round(pct * 100) + '%'; b.parentNode.previousSibling.textContent = '☁ ' + txt; }
    else pintarPanelDrive();
  }).then(function (out) {
    gen.subiendo = false; gen.carpetaId = out.carpetaId;
    Object.keys(out.archivos).forEach(function (u) { gen.links[u] = out.archivos[u].link; });
    pintarPanelDrive(); pintarEnvio();
    toast('☁ Recibos guardados en Google Drive. Ya puedes enviar los links.');
  }).catch(function (e) {
    gen.subiendo = false; gen.errorDrive = e.message || String(e);
    pintarPanelDrive(); pintarEnvio();
    console.error(e);
  });
}

/* ---------------- Enviar por correo / WhatsApp (sin servidor: la persona revisa y envía) ---------------- */
var enviados = {};      // {unidad: {correo:true, wa:true}} solo durante esta sesión
var correosEscritos = {};
var urlsPdf = [];

function toast(html, ms) {
  var t = $('toast');
  if (!t) { t = document.createElement('div'); t.id = 'toast'; t.className = 'toast'; t.setAttribute('role', 'status'); document.body.appendChild(t); }
  t.innerHTML = html; t.classList.add('visible');
  clearTimeout(toast._t); toast._t = setTimeout(function () { t.classList.remove('visible'); }, ms || 7000);
}

/** Copia el cuerpo con formato (HTML) y en texto. Devuelve Promise<boolean>. */
function copiarCorreo(c) {
  function respaldo() {
    var div = document.createElement('div');
    div.contentEditable = 'true';
    div.style.cssText = 'position:fixed;left:-10000px;top:0;width:900px;background:#fff';
    div.innerHTML = c.html;
    document.body.appendChild(div);
    var sel = window.getSelection(), rango = document.createRange();
    rango.selectNodeContents(div); sel.removeAllRanges(); sel.addRange(rango);
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    sel.removeAllRanges(); div.remove();
    return ok;
  }
  if (navigator.clipboard && window.ClipboardItem && window.isSecureContext) {
    try {
      var item = new ClipboardItem({
        'text/html': new Blob([c.html], { type: 'text/html' }),
        'text/plain': new Blob([c.texto], { type: 'text/plain' })
      });
      return navigator.clipboard.write([item]).then(function () { return true; }, function () { return respaldo(); });
    } catch (e) { return Promise.resolve(respaldo()); }
  }
  return Promise.resolve(respaldo());
}

function abrir(url) {
  var w = window.open(url, '_blank');
  if (w) { try { w.opener = null; } catch (e) {} }
  return w;
}

function pintarEnvio() {
  var panel = $('panelEnvio'); if (!panel || !gen) return;
  urlsPdf.forEach(function (u) { URL.revokeObjectURL(u); }); urlsPdf = [];
  var res = gen.res, out = gen.pdfs;
  var cfg = estado.datos.config || {};
  var mesTexto = res.nombreMes + ' ' + res.anio;
  var movil = esMovil();
  function depto(unidad) { return estado.datos.deptos.filter(function (x) { return Calculo.mismaUnidad(x.unidad, unidad); })[0] || {}; }
  function correoDe(unidad) { return correosEscritos[unidad] !== undefined ? correosEscritos[unidad] : (depto(unidad).correo || ''); }
  function armar(r) {
    return Correo.armar(cfg, r.unidad, mesTexto, Calculo.fmtSoles(r.total), correoDe(r.unidad),
      { link: gen.links[r.unidad] || '', telefono: depto(r.unidad).telefono || '', movil: movil });
  }
  var hayLinks = Object.keys(gen.links).length > 0;
  var sinCorreo = out.filter(function (r) { return !Correo.listaCorreos(correoDe(r.unidad)).length; }).length;
  var sinTel = out.filter(function (r) { return !Correo.telefonoWhatsapp(depto(r.unidad).telefono); }).length;
  var e0 = function (u, k) { return enviados[u] && enviados[u][k]; };

  panel.innerHTML =
    '<h3 class="sub-titulo">Recibos por departamento</h3>' +
    '<p class="ayuda"><b>⬇ PDF</b> descarga el recibo. <b>✉ Correo</b> abre ' + (movil ? 'tu app de correo' : 'Gmail') + ' con el destinatario, el asunto y el mensaje' +
    (movil ? '' : ' (el mensaje se copia: <b>pega con Ctrl+V</b> y envía)') + '. <b>🟢 WhatsApp</b> abre ' + (movil ? 'WhatsApp' : 'WhatsApp Web') + ' con el mensaje listo. ' +
    'El correo y el WhatsApp llevan el link para ver el recibo.</p>' +
    (hayLinks ? '' : '<div class="banner aviso">Conecta Google Drive para enviar el link del recibo.</div>') +
    (sinCorreo || sinTel ? '<div class="banner aviso">' + (sinCorreo ? sinCorreo + ' dpto(s) sin correo. ' : '') + (sinTel ? sinTel + ' dpto(s) sin teléfono. ' : '') +
      'Para que quede guardado, escríbelo en la pestaña <b>Departamentos</b> de tu Excel.</div>' : '') +
    '<div class="lista-envio">' + out.map(function (r, i) {
      var url = URL.createObjectURL(r.blob); urlsPdf.push(url);
      var c = armar(r);
      var waBtn = c.telefono
        ? '<button type="button" class="boton wa-btn' + (e0(r.unidad, 'wa') ? ' hecho' : '') + (c.datos.link ? '' : ' sin-link') + '">' + (e0(r.unidad, 'wa') ? '🟢 Enviado ✓' : '🟢 WhatsApp') + '</button>'
        : '<button type="button" class="boton wa-btn no-disp" disabled title="Sin teléfono en el Excel">WhatsApp no disponible</button>';
      return '<div class="envio-fila' + (e0(r.unidad, 'correo') || e0(r.unidad, 'wa') ? ' enviado' : '') + '" data-i="' + i + '">' +
        '<div class="envio-info"><div class="envio-dpto">Dpto ' + esc(r.unidad) + ' <span class="envio-total">' + esc(Calculo.fmtSoles(r.total)) + '</span></div>' +
        '<div class="mini">' + esc(r.propietario || '') + (c.telefono ? ' · 📱 ' + esc(Correo.telefonoBonito(depto(r.unidad).telefono)) : '') + '</div>' +
        '<input type="email" class="envio-correo" inputmode="email" autocomplete="off" placeholder="sin correo · escríbelo aquí" value="' + esc(correoDe(r.unidad)) + '">' +
        '<div class="envio-estado mini">' + (c.correos.length ? '' : '<span class="pill falta">sin correo</span>') + '</div></div>' +
        '<div class="envio-botones">' +
          '<a class="boton sec envio-pdf" href="' + url + '" download="' + esc(r.nombre) + '">⬇ PDF</a>' +
          '<button type="button" class="boton envio-btn' + (e0(r.unidad, 'correo') ? ' hecho' : '') + (c.datos.link ? '' : ' sin-link') + '">' + (e0(r.unidad, 'correo') ? '✉ Enviado ✓' : '✉ Correo') + '</button>' +
          waBtn +
          '<div class="envio-links"><button type="button" class="enlace envio-ver">Ver mensajes</button>' +
          (c.datos.link ? ' · <a class="enlace envio-link" href="' + esc(c.datos.link) + '" target="_blank" rel="noopener">Ver en Drive</a>' : '') +
          (movil ? '' : ' · <a class="enlace envio-mailto" href="' + esc(c.mailto) + '">Otra app de correo</a>') + '</div>' +
        '</div></div>';
    }).join('') + '</div>' +
    '<div class="previas"><div><h3 class="sub-titulo">Vista previa del correo</h3>' +
    '<div class="correo-prev"><div class="correo-prev-cab"><span class="mini">Para:</span> <b id="prevPara"></b><br>' +
    '<span class="mini">Asunto:</span> <b id="prevAsunto"></b></div><div id="prevCuerpo" class="correo-prev-cuerpo"></div></div></div>' +
    '<div><h3 class="sub-titulo">Vista previa del WhatsApp</h3><div class="wa-prev"><div class="wa-cab" id="prevWaPara"></div><div class="wa-burbuja" id="prevWa"></div></div></div></div>';

  function waHtml(t) {
    return esc(t).replace(/\*([^*\n]+)\*/g, '<b>$1</b>').replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
  }
  function mostrarPrev(r) {
    var c = armar(r);
    $('prevPara').textContent = c.correos.length ? c.correos.join(', ') : '(sin correo)';
    $('prevAsunto').textContent = c.asunto;
    $('prevCuerpo').innerHTML = c.html;
    $('prevWaPara').textContent = 'Para: ' + (c.telefono ? '+' + c.telefono + ' (Dpto ' + r.unidad + ')' : 'Dpto ' + r.unidad + ' — sin teléfono en el Excel');
    $('prevWa').innerHTML = waHtml(c.whatsappTexto);
  }
  mostrarPrev(out[0]);

  function sinLink() {
    if (gen.subiendo) { toast('⏳ Espera a que terminen de subir los recibos a Google Drive.'); return true; }
    toast('<b>Conecta Google Drive para enviar el link del recibo.</b> Mientras tanto puedes descargar el PDF con ⬇ PDF.', 8000);
    var b = $('btnConectarRes') || $('btnSubirDrive'); if (b) b.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return true;
  }

  panel.querySelectorAll('.envio-fila').forEach(function (fila) {
    var r = out[+fila.getAttribute('data-i')];
    var inp = fila.querySelector('.envio-correo');
    inp.addEventListener('input', function () {
      correosEscritos[r.unidad] = inp.value;
      var c = armar(r);
      var m = fila.querySelector('.envio-mailto'); if (m) m.href = c.mailto;
      fila.querySelector('.envio-estado').innerHTML = c.correos.length ? '' : (inp.value.trim() ? '<span class="pill aviso">correo no válido</span>' : '<span class="pill falta">sin correo</span>');
    });
    fila.querySelector('.envio-ver').onclick = function () { mostrarPrev(r); $('prevCuerpo').scrollIntoView({ behavior: 'smooth', block: 'center' }); };
    function marcar(k, btn) {
      enviados[r.unidad] = enviados[r.unidad] || {}; enviados[r.unidad][k] = true;
      fila.classList.add('enviado');
      btn.textContent = (k === 'wa' ? '🟢' : '✉') + ' Enviado ✓'; btn.classList.add('hecho');
    }
    var m = fila.querySelector('.envio-mailto');
    if (m) m.addEventListener('click', function (ev) { if (!gen.links[r.unidad]) { ev.preventDefault(); sinLink(); return; } marcar('correo', fila.querySelector('.envio-btn')); });
    fila.querySelector('.envio-btn').onclick = function () {
      var btn = this, c = armar(r);
      mostrarPrev(r);
      if (!c.datos.link) { sinLink(); return; }
      if (movil) {
        if (c.correos.length || !navigator.share) {
          window.location.href = c.mailto;   // abre la app de correo con todo escrito (texto simple + link)
        } else {
          navigator.share({ title: c.asunto, text: c.texto }).catch(function () {});
        }
        marcar('correo', btn);
        return;
      }
      var copia = copiarCorreo(c);              // dentro del clic (permiso del portapapeles)
      var w = abrir(c.gmail);                    // también dentro del clic (evita bloqueo de ventanas)
      copia.then(function (ok) {
        toast((ok ? '✔ Mensaje copiado. ' : '⚠ No se pudo copiar el mensaje. ') +
          '<b>En Gmail, haz clic en el cuerpo, pega con Ctrl+V y presiona Enviar.</b>' +
          (w ? '' : '<br>El navegador bloqueó la ventana de Gmail: <a href="' + esc(c.gmail) + '" target="_blank" rel="noopener">ábrela aquí</a>.'), 10000);
      });
      marcar('correo', btn);
    };
    var wa = fila.querySelector('.wa-btn');
    if (!wa.disabled) wa.onclick = function () {
      var c = armar(r);
      mostrarPrev(r);
      if (!c.datos.link) { sinLink(); return; }
      var w = abrir(c.whatsapp);
      if (!w) toast('El navegador bloqueó la ventana de WhatsApp: <a href="' + esc(c.whatsapp) + '" target="_blank" rel="noopener">ábrela aquí</a>.', 10000);
      marcar('wa', wa);
    };
  });
}
window.__pintarEnvio = pintarEnvio;

/* ---------------- Google Drive: Conectar / Desconectar (barra de arriba) ---------------- */
function conectarDrive(desdeResultado) {
  if (!Drive.configurado()) { toast('Esta página aún no tiene configurado Google Drive (falta GOOGLE_CLIENT_ID en js/config.js).', 9000); return; }
  if (!Drive.gisListo()) { toast('Google todavía está cargando. Espera unos segundos e inténtalo otra vez.'); return; }
  Drive.conectar().then(function (e) {
    toast('✔ Google Drive conectado' + (e.email ? ': <b>' + esc(e.email) + '</b>' : '') + '.');
    if (gen && !Object.keys(gen.links).length) subirADrive();
  }).catch(function (err) {
    toast('⚠ ' + esc(err.message || err), 9000);
  });
}
function pintarBarraDrive() {
  var b = $('barraDrive0'); if (!b) return;
  var e = Drive.estado();
  if (!Drive.configurado()) {
    b.innerHTML = '<span class="drive-ico">☁</span><div class="drive-txt"><b>Google Drive</b><span class="mini">Falta configurar la página (GOOGLE_CLIENT_ID). Puedes generar y descargar los PDF igual.</span></div>' +
      '<button type="button" class="boton sec" id="btnConectar" disabled>Conectar Google Drive</button>';
  } else if (e.conectado) {
    b.innerHTML = '<span class="drive-ico ok">☁</span><div class="drive-txt"><b>Google Drive conectado</b><span class="mini" id="driveEmail">' + esc(e.email || 'cuenta de Google') + '</span></div>' +
      '<button type="button" class="boton sec" id="btnDesconectar">Desconectar</button>';
  } else {
    b.innerHTML = '<span class="drive-ico">☁</span><div class="drive-txt"><b>Google Drive</b><span class="mini">Conéctalo para guardar los recibos y enviar el link a cada propietario.</span></div>' +
      '<button type="button" class="boton primario" id="btnConectar">Conectar Google Drive</button>';
  }
  var c = $('btnConectar'); if (c && !c.disabled) c.onclick = function () { conectarDrive(false); };
  var d = $('btnDesconectar'); if (d) d.onclick = function () {
    if (!window.confirm('¿Desconectar Google Drive (' + (e.email || 'esta cuenta') + ')?\nLos recibos ya guardados siguen en ese Drive. Después puedes conectar otra cuenta.')) return;
    Drive.desconectar().then(function () { toast('Google Drive desconectado. Puedes conectar otra cuenta cuando quieras.'); });
  };
}
Drive.alCambiar(function () { pintarBarraDrive(); if (gen) { pintarPanelDrive(); } });
pintarBarraDrive();


// Expone para pruebas (Playwright)
window.__RecibosApp = {
  estado: estado, cargarExcel: cargarExcel, agregarArchivos: agregarArchivos, irA: irA,
  gen: function () { return gen; }, pintarBarraDrive: pintarBarraDrive, subirADrive: subirADrive
};
})();
