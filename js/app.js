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
  if (paso > 1 && !estado.datos) return;
  if (paso > 2 && !estado.borrador) return;
  estado.paso = paso;
  [1, 2, 3].forEach(function (n) {
    mostrar('paso' + n, n === paso);
    var t = $('tab' + n);
    t.classList.toggle('activo', n === paso);
    t.classList.toggle('hecho', n < paso);
    t.disabled = (n > 1 && !estado.datos) || (n > 2 && !estado.borrador);
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
    var ms = Excel.mesesConDatos(L.datos);
    var sug = Excel.mesSugerido(L.datos, Excel.hoyISO().slice(0, 7));
    info.innerHTML = '<div class="banner ok">✔ Excel cargado: <b>' + esc(f.name) + '</b><br>' +
      esc(L.datos.config.EDIFICIO_NOMBRE || '(sin nombre)') + ' · ' + L.datos.deptos.length + ' departamentos' +
      (ms.length ? ' · meses registrados: ' + ms.join(', ') : ' · sin meses aún') +
      '<br>Mes sugerido: <b>' + esc(Calculo.nombreMes(sug) + ' ' + Calculo.partesMes(sug).anio) + '</b></div>' +
      (L.avisos.length ? '<div class="banner aviso"><ul>' + L.avisos.map(function (a) { return '<li>' + esc(a) + '</li>'; }).join('') + '</ul></div>' : '') +
      '<div class="botones"><button type="button" class="boton primario" id="btnAPaso2">Continuar → Datos del mes</button></div>';
    $('btnAPaso2').onclick = function () { abrirPaso2(sug); };
  }).catch(function (e) {
    info.innerHTML = '<div class="banner error">No se pudo leer el Excel: ' + esc(e.message || e) + '</div>';
  });
}
setupDrop('dropExcel', 'fileExcel', cargarExcel);

$('btnBlanco').onclick = function () {
  var wb = Excel.libroNuevo({
    config: {
      EDIFICIO_NOMBRE: '', DIRECCION_LINEA_1: '', DIRECCION_LINEA_2: '',
      PRESIDENTE_NOMBRE: '', PRESIDENTE_CELULAR: '', TESORERO_NOMBRE: '', TESORERO_CELULAR: '',
      BANCO: '', CUENTA_BANCARIA: '', CCI: '',
      NOTA_PAGO: 'Por favor, envíe su comprobante de pago al tesorero vía WhatsApp indicando su número de departamento.',
      EMPRESA_AGUA: 'Sedapal', EMPRESA_LUZ: 'Luz del Sur', NUM_DPTOS: 10, DIA_VENCIMIENTO: 5, SALDO_INICIAL_FONDO: 0
    },
    deptos: [101, 102, 201, 202, 301, 302, 401, 402, 501, 502].map(function (u) { return { unidad: u, propietario: '' }; }),
    lecturas: [], servicios: [], gastos: [
      { mes: '1970-01', num: 1, concepto: 'Salario del personal', monto: 0, tipo: 'FIJO', copiar: 'SÍ' },
      { mes: '1970-01', num: 2, concepto: 'Consumo de luz', monto: '', tipo: 'LUZ', copiar: 'SÍ' },
      { mes: '1970-01', num: 3, concepto: 'Mantenimiento de ascensor', monto: 0, tipo: 'FIJO', copiar: 'SÍ' },
      { mes: '1970-01', num: 4, concepto: 'Jardinero / Mantenimiento de áreas verdes', monto: 0, tipo: 'FIJO', copiar: 'SÍ' },
      { mes: '1970-01', num: 5, concepto: 'Agua común', monto: '', tipo: 'AGUA_COMUN', copiar: 'SÍ' },
      { mes: '1970-01', num: 6, concepto: 'Fondo de contingencia', monto: 800, tipo: 'FONDO', copiar: 'SÍ' },
      { mes: '1970-01', num: 7, concepto: 'Recarga de celular conserje', monto: 0, tipo: 'FIJO', copiar: 'SÍ' },
      { mes: '1970-01', num: 8, concepto: 'Domingos - días', monto: 0, tipo: 'FIJO', copiar: 'SÍ' },
      { mes: '1970-01', num: 9, concepto: 'Materiales de limpieza', monto: 0, tipo: 'FIJO', copiar: 'SÍ' }
    ], ajustes: [], pagos: [], fondo: []
  });
  // Quita el mes ficticio 1970-01 dejando solo la plantilla de gastos como "modelo" en un comentario... mejor: gastos base sin mes
  // Reescribimos Gastos con mes vacío → la primera vez el usuario los completa; prepararBorrador copia del último.
  // Truco: usamos un mes "modelo" que prepararBorrador encontrará si no hay otro.
  XLSX.writeFile(wb, 'Base de datos - Edificio (en blanco).xlsx');
};

/* ---------------- Paso 2 ---------------- */
function abrirPaso2(mes) {
  var ops = Excel.opcionesMeses(estado.datos, Excel.hoyISO().slice(0, 7));
  var sel = $('selMes');
  sel.innerHTML = ops.map(function (o) {
    return '<option value="' + o.v + '"' + (o.v === mes ? ' selected' : '') + '>' + esc(o.label) + '</option>';
  }).join('');
  cargarBorrador(mes || sel.value);
  irA(2);
}
$('selMes').onchange = function () { cargarBorrador(this.value); };
$('btnVolver1').onclick = function () { irA(1); };
$('tab1').onclick = function () { irA(1); };
$('tab2').onclick = function () { if (estado.datos) irA(2); };
$('tab3').onclick = function () { if (estado.borrador) irA(3); };

function cargarBorrador(mes) {
  estado.borrador = Excel.prepararBorrador(estado.datos, mes);
  var b = estado.borrador;
  var av = $('avisoMes');
  if (b.existe) {
    av.className = 'banner aviso'; av.classList.remove('oculto');
    av.textContent = 'Este mes ya está en el Excel. Si generas de nuevo, se actualizarán sus filas.';
  } else {
    av.className = 'banner info'; av.classList.remove('oculto');
    av.textContent = 'Mes nuevo: se copiaron los gastos fijos del mes anterior y la lectura anterior = lectura actual del mes previo.';
  }
  $('aguaImporte').value = b.servicios.aguaImporte === '' ? '' : b.servicios.aguaImporte;
  $('aguaM3').value = b.servicios.aguaM3 === '' ? '' : b.servicios.aguaM3;
  $('luzImporte').value = b.servicios.luzImporte === '' ? '' : b.servicios.luzImporte;
  $('emision').value = b.servicios.emision || '';
  $('vencimiento').value = b.servicios.vencimiento || '';
  $('fondoAnterior').value = b.fondo.saldoAnterior === '' ? '' : b.fondo.saldoAnterior;
  $('fondoGastos').value = b.fondo.gastosFondo === '' ? '' : b.fondo.gastosFondo;
  $('fondoDetalle').value = b.fondo.detalle || '';
  if (estado.logoDataUrl) $('logoPreview').innerHTML = '<img src="' + estado.logoDataUrl + '" height="40" alt="logo"> Logo listo';
  else $('logoPreview').textContent = 'Sin logo (opcional)';
  pintarLecturas();
  pintarGastos();
  pintarAjustes();
}

function pintarLecturas() {
  var tb = $('tablaLecturas').querySelector('tbody');
  tb.innerHTML = estado.borrador.lecturas.map(function (l, i) {
    var ant = l.anterior === '' || l.anterior === null ? '' : l.anterior;
    var act = l.actual === '' || l.actual === null ? '' : l.actual;
    var cons = '';
    var a = Calculo.parseNumero(act), p = Calculo.parseNumero(ant);
    if (a !== null && p !== null) cons = Calculo.fmtNum(Calculo.redondear(a - p, 3), 3);
    return '<tr data-i="' + i + '"><td><b>' + esc(l.unidad) + '</b></td><td>' + esc(l.propietario) + '</td>' +
      '<td><input type="text" inputmode="decimal" class="lec-ant" value="' + esc(ant) + '"></td>' +
      '<td><input type="text" inputmode="decimal" class="lec-act" value="' + esc(act) + '"><div class="aviso-campo"></div></td>' +
      '<td class="cons">' + esc(cons) + '</td></tr>';
  }).join('');
  tb.querySelectorAll('input').forEach(function (inp) {
    inp.addEventListener('input', function () {
      var tr = inp.closest('tr'), i = +tr.getAttribute('data-i');
      var ant = tr.querySelector('.lec-ant').value, act = tr.querySelector('.lec-act').value;
      estado.borrador.lecturas[i].anterior = ant; estado.borrador.lecturas[i].actual = act;
      var a = Calculo.parseNumero(act), p = Calculo.parseNumero(ant);
      var av = tr.querySelector('.aviso-campo');
      tr.querySelector('.lec-act').classList.remove('malo');
      if (act && a === null) { av.textContent = 'número inválido'; tr.querySelector('.lec-act').classList.add('malo'); }
      else if (a !== null && p !== null && a < p) { av.textContent = 'menor que la anterior'; tr.querySelector('.lec-act').classList.add('malo'); }
      else av.textContent = '';
      tr.querySelector('.cons').textContent = (a !== null && p !== null) ? Calculo.fmtNum(Calculo.redondear(a - p, 3), 3) : '';
    });
  });
}

function pintarGastos() {
  var box = $('listaGastos');
  box.innerHTML = estado.borrador.gastos.map(function (g, i) {
    var montoDisabled = (g.tipo === 'LUZ' || g.tipo === 'AGUA_COMUN') ? ' disabled placeholder="auto"' : '';
    return '<div class="fila" data-i="' + i + '">' +
      '<input type="text" class="concepto" value="' + esc(g.concepto) + '" placeholder="Concepto">' +
      '<input type="text" inputmode="decimal" class="monto" value="' + esc(g.monto === '' || g.monto === null ? '' : g.monto) + '"' + montoDisabled + '>' +
      '<select class="tipo">' + Excel.TIPOS.map(function (t) {
        return '<option value="' + t + '"' + (g.tipo === t ? ' selected' : '') + '>' + t + '</option>';
      }).join('') + '</select>' +
      '<button type="button" class="boton sec x" title="Quitar">×</button></div>';
  }).join('');
  box.querySelectorAll('.fila').forEach(function (fila) {
    var i = +fila.getAttribute('data-i');
    fila.querySelector('.concepto').oninput = function () { estado.borrador.gastos[i].concepto = this.value; };
    fila.querySelector('.monto').oninput = function () { estado.borrador.gastos[i].monto = this.value; };
    fila.querySelector('.tipo').onchange = function () {
      estado.borrador.gastos[i].tipo = this.value;
      var m = fila.querySelector('.monto');
      if (this.value === 'LUZ' || this.value === 'AGUA_COMUN') { m.value = ''; m.disabled = true; estado.borrador.gastos[i].monto = ''; }
      else m.disabled = false;
    };
    fila.querySelector('.x').onclick = function () { estado.borrador.gastos.splice(i, 1); pintarGastos(); };
  });
}
$('btnAddGasto').onclick = function () {
  estado.borrador.gastos.push({ concepto: '', monto: '', tipo: 'EXTRA', copiar: false });
  pintarGastos();
};

function pintarAjustes() {
  var box = $('listaAjustes');
  var opts = estado.datos.deptos.map(function (d) { return '<option value="' + esc(d.unidad) + '">' + esc(d.unidad) + '</option>'; }).join('');
  box.innerHTML = estado.borrador.ajustes.map(function (a, i) {
    return '<div class="fila" data-i="' + i + '">' +
      '<select class="unidad">' + opts.replace('value="' + esc(a.unidad) + '"', 'value="' + esc(a.unidad) + '" selected') + '</select>' +
      '<input type="text" class="concepto" value="' + esc(a.concepto || '') + '" placeholder="Concepto">' +
      '<input type="text" inputmode="decimal" class="monto" value="' + esc(a.monto === '' || a.monto === null ? '' : a.monto) + '" placeholder="Monto">' +
      '<button type="button" class="boton sec x">×</button></div>';
  }).join('');
  // Fix selected: rebuild properly
  box.innerHTML = estado.borrador.ajustes.map(function (a, i) {
    var o = estado.datos.deptos.map(function (d) {
      return '<option value="' + esc(d.unidad) + '"' + (String(d.unidad) === String(a.unidad) ? ' selected' : '') + '>' + esc(d.unidad) + '</option>';
    }).join('');
    return '<div class="fila" data-i="' + i + '"><select class="unidad">' + o + '</select>' +
      '<input type="text" class="concepto" value="' + esc(a.concepto || '') + '" placeholder="Concepto">' +
      '<input type="text" inputmode="decimal" class="monto" value="' + esc(a.monto === '' || a.monto === null ? '' : a.monto) + '">' +
      '<button type="button" class="boton sec x">×</button></div>';
  }).join('');
  box.querySelectorAll('.fila').forEach(function (fila) {
    var i = +fila.getAttribute('data-i');
    fila.querySelector('.unidad').onchange = function () { estado.borrador.ajustes[i].unidad = this.value; };
    fila.querySelector('.concepto').oninput = function () { estado.borrador.ajustes[i].concepto = this.value; };
    fila.querySelector('.monto').oninput = function () { estado.borrador.ajustes[i].monto = this.value; };
    fila.querySelector('.x').onclick = function () { estado.borrador.ajustes.splice(i, 1); pintarAjustes(); };
  });
}
$('btnAddAjuste').onclick = function () {
  estado.borrador.ajustes.push({ unidad: estado.datos.deptos[0].unidad, concepto: 'Crédito a favor', monto: '' });
  pintarAjustes();
};

$('logoFile').onchange = function () {
  var f = this.files && this.files[0]; if (!f) return;
  Media.aJpeg(f, 400).then(function (r) {
    estado.logoDataUrl = r.dataUrl;
    $('logoPreview').innerHTML = '<img src="' + r.dataUrl + '" height="40" alt="logo"> Logo listo';
  }).catch(function (e) { alert(e.message || e); });
};

function recolectarPaso2() {
  var b = estado.borrador;
  b.servicios.aguaImporte = $('aguaImporte').value;
  b.servicios.aguaM3 = $('aguaM3').value;
  b.servicios.luzImporte = $('luzImporte').value;
  b.servicios.emision = $('emision').value;
  b.servicios.vencimiento = $('vencimiento').value;
  b.fondo.saldoAnterior = $('fondoAnterior').value;
  b.fondo.gastosFondo = $('fondoGastos').value;
  b.fondo.detalle = $('fondoDetalle').value;
}

$('btnAPaso3').onclick = function () {
  recolectarPaso2();
  irA(3);
  pintarEstadoArchivos();
};
$('btnVolver2').onclick = function () { irA(2); };

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
  recolectarPaso2();
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

function ejecutarGeneracion(aplicado) {
  mostrar('panelProgreso', true);
  mostrar('panelResultado', false);
  $('btnGenerar').disabled = true;
  var res = aplicado.res;
  // Escribe el mes en el libro y el logo
  var wb = estado.wb;
  if (estado.logoDataUrl) {
    Excel.ponerLogo(wb, estado.logoDataUrl);
    aplicado.datos.logo = estado.logoDataUrl;
  }
  Excel.escribirMes(wb, aplicado.datos, res.mes);
  // Actualiza Config.NUM_DPTOS si hace falta
  var excelBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  var archivos = {
    logo: estado.logoDataUrl || aplicado.datos.logo || '',
    medidor: estado.archivos.medidor,
    agua: estado.archivos.agua,
    luz: estado.archivos.luz
  };
  PdfGen.generarZip(aplicado.datos, res, archivos, excelBytes, function (txt, pct) {
    $('txtProgreso').textContent = txt;
    $('barraProgreso').style.width = Math.round(pct * 100) + '%';
  }).then(function (out) {
    $('btnGenerar').disabled = false;
    mostrar('panelProgreso', false);
    mostrar('panelResultado', true);
    // Descarga automática
    var a = document.createElement('a');
    a.href = URL.createObjectURL(out.blob);
    a.download = out.nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 30000);
    var filas = out.resumen.map(function (r) {
      return '<tr><td>' + esc(r.unidad) + '</td><td>' + esc(r.propietario) + '</td><td class="m">' + esc(Calculo.fmtSoles(r.total)) + '</td></tr>';
    }).join('');
    $('panelResultado').innerHTML =
      '<div class="banner ok"><b>✔ Listo:</b> ' + out.resumen.length + ' recibos de ' + esc(res.nombreMes + ' ' + res.anio) +
      '. Se descargó <b>' + esc(out.nombre) + '</b> (PDF + Excel actualizado).<br>' +
      'Guarda el Excel del ZIP: es el que debes subir el próximo mes.</div>' +
      '<div class="tabla-wrap"><table class="datos resumen-tabla"><thead><tr><th>Dpto</th><th>Propietario</th><th class="m">Total</th></tr></thead><tbody>' +
      filas +
      '<tr><td colspan="2"><b>Total facturado</b></td><td class="m"><b>' + esc(Calculo.fmtSoles(res.totalFacturado)) + '</b></td></tr>' +
      '<tr><td colspan="2">Fondo de contingencia</td><td class="m">' + esc(Calculo.fmtSoles(res.fondo.saldo)) + '</td></tr>' +
      '</tbody></table></div>' +
      '<div class="botones"><button type="button" class="boton primario" id="btnReDescargar">⬇ Volver a descargar el ZIP</button></div>';
    $('btnReDescargar').onclick = function () {
      var a2 = document.createElement('a');
      a2.href = URL.createObjectURL(out.blob); a2.download = out.nombre;
      document.body.appendChild(a2); a2.click(); a2.remove();
      setTimeout(function () { URL.revokeObjectURL(a2.href); }, 30000);
    };
    // Actualiza estado.datos con el mes escrito (para generar de nuevo o cambiar de mes)
    estado.datos = aplicado.datos;
  }).catch(function (e) {
    $('btnGenerar').disabled = false;
    mostrar('panelProgreso', false);
    mostrar('panelResultado', true);
    $('panelResultado').innerHTML = '<div class="banner error">Error al generar: ' + esc(e.message || e) + '</div>';
    console.error(e);
  });
}

// Expone para pruebas (Playwright)
window.__RecibosApp = {
  estado: estado, cargarExcel: cargarExcel, abrirPaso2: abrirPaso2, agregarArchivos: agregarArchivos,
  recolectarPaso2: recolectarPaso2, irA: irA
};
})();
