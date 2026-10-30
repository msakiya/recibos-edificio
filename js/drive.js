/*
 * drive.js — Conexión con Google Drive (Google Identity Services, token en el navegador) y subida de PDF.
 * Permiso único: drive.file (la página solo ve los archivos y carpetas que ella misma crea).
 * Se recuerda en localStorage solo "conectado" + correo; el token vive en memoria (dura ~1 hora).
 */
(function (global) {
'use strict';
var CLAVE = 'recibos.drive';
var API = 'https://www.googleapis.com/drive/v3';
var UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
var CARPETA = 'application/vnd.google-apps.folder';

var token = null, vence = 0, cliente = null, pendiente = null, manejador = null;
var oyentes = [];

function cfg() { return global.RECIBOS_CONFIG || {}; }
function clientId() { return String(cfg().GOOGLE_CLIENT_ID || '').trim(); }
function scope() { return cfg().DRIVE_SCOPE || 'https://www.googleapis.com/auth/drive.file'; }
function configurado() { return !!clientId(); }
function gisListo() { return !!(global.google && global.google.accounts && global.google.accounts.oauth2); }

function leer() { try { return JSON.parse(global.localStorage.getItem(CLAVE) || 'null'); } catch (e) { return null; } }
function guardar(v) { try { if (v) global.localStorage.setItem(CLAVE, JSON.stringify(v)); else global.localStorage.removeItem(CLAVE); } catch (e) {} }
function estado() { var v = leer(); return { conectado: !!(v && v.conectado), email: (v && v.email) || '', nombre: (v && v.nombre) || '' }; }
function avisar() { var e = estado(); oyentes.forEach(function (f) { try { f(e); } catch (x) {} }); }
function alCambiar(f) { oyentes.push(f); }
function tokenValido() { return !!token && Date.now() < vence - 60000; }

function obtenerCliente() {
  if (cliente) return cliente;
  cliente = global.google.accounts.oauth2.initTokenClient({
    client_id: clientId(),
    scope: scope(),
    callback: function (r) { if (manejador) manejador.ok(r); },
    error_callback: function (err) { if (manejador) manejador.mal(err); }
  });
  return cliente;
}

/**
 * Pide un token. interactivo=true: elige cuenta (botón Conectar). false: silencioso (prompt '').
 * IMPORTANTE: llamarlo directo desde un clic (Google abre una ventanita).
 */
function pedirToken(interactivo) {
  if (!configurado()) return Promise.reject(new Error('Falta configurar GOOGLE_CLIENT_ID en js/config.js.'));
  if (!gisListo()) return Promise.reject(new Error('Google todavía está cargando. Espera unos segundos e inténtalo otra vez.'));
  if (!interactivo && tokenValido()) return Promise.resolve(token);
  if (pendiente && !interactivo) return pendiente;
  var c = obtenerCliente();
  var e = estado();
  pendiente = new Promise(function (ok, mal) {
    manejador = {};
    manejador.ok = function (r) {
      pendiente = null;
      if (!r || r.error || !r.access_token) { mal(new Error(textoError(r && r.error))); return; }
      if (global.google.accounts.oauth2.hasGrantedAllScopes && !global.google.accounts.oauth2.hasGrantedAllScopes(r, scope())) {
        mal(new Error('No se dio el permiso para guardar en Drive. Vuelve a conectar y marca la casilla de Google Drive.')); return;
      }
      token = r.access_token; vence = Date.now() + (Number(r.expires_in) || 3600) * 1000;
      ok(token);
    };
    manejador.mal = function (err) {
      pendiente = null;
      mal(new Error(err && err.type === 'popup_closed' ? 'Se cerró la ventana de Google antes de terminar.' :
        err && err.type === 'popup_failed_to_open' ? 'El navegador bloqueó la ventana de Google. Permite las ventanas emergentes para esta página.' :
        'No se pudo conectar con Google.'));
    };
    var opc = interactivo ? { prompt: 'select_account' } : { prompt: '' };
    if (!interactivo && e.email) opc.hint = e.email;
    c.requestAccessToken(opc);
  });
  return pendiente;
}
function textoError(code) {
  if (code === 'access_denied') return 'No se dio el permiso en Google.';
  if (code === 'interaction_required' || code === 'consent_required' || code === 'login_required') return 'Google pide volver a elegir la cuenta. Toca “Conectar Google Drive”.';
  return 'No se pudo conectar con Google' + (code ? ' (' + code + ')' : '') + '.';
}

function api(metodo, url, cuerpo, tipo) {
  return pedirToken(false).then(function (t) {
    var h = { Authorization: 'Bearer ' + t };
    if (tipo) h['Content-Type'] = tipo;
    return global.fetch(url, { method: metodo, headers: h, body: cuerpo });
  }).then(function (r) {
    if (r.status === 401) { token = null; vence = 0; throw new Error('La sesión de Google venció. Vuelve a tocar el botón.'); }
    if (!r.ok) return r.text().then(function (t) { throw new Error('Google Drive respondió ' + r.status + ': ' + String(t).slice(0, 200)); });
    return r.status === 204 ? {} : r.json();
  });
}

/** Botón "Conectar Google Drive": elige cuenta, pide permiso y guarda el correo. */
function conectar() {
  return pedirToken(true).then(function () {
    return api('GET', API + '/about?fields=user(emailAddress,displayName)').catch(function () { return {}; });
  }).then(function (a) {
    var u = (a && a.user) || {};
    guardar({ conectado: true, email: u.emailAddress || '', nombre: u.displayName || '' });
    avisar();
    return estado();
  });
}

/** Botón "Desconectar": revoca el permiso y borra lo guardado. */
function desconectar() {
  var t = token;
  token = null; vence = 0; pendiente = null;
  guardar(null);
  avisar();
  return new Promise(function (ok) {
    if (t && gisListo() && global.google.accounts.oauth2.revoke) {
      try { global.google.accounts.oauth2.revoke(t, function () { ok(true); }); } catch (e) { ok(false); }
      setTimeout(function () { ok(true); }, 4000);
    } else ok(false);
  });
}

function q(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

function buscar(nombre, padre, esCarpeta) {
  var consulta = "name = '" + q(nombre) + "' and '" + q(padre) + "' in parents and trashed = false" +
    (esCarpeta ? " and mimeType = '" + CARPETA + "'" : " and mimeType != '" + CARPETA + "'");
  return api('GET', API + '/files?q=' + encodeURIComponent(consulta) + '&fields=' + encodeURIComponent('files(id,name,webViewLink)') + '&spaces=drive&pageSize=10')
    .then(function (r) { return (r.files || [])[0] || null; });
}

function carpeta(nombre, padre) {
  return buscar(nombre, padre, true).then(function (f) {
    if (f) return f.id;
    return api('POST', API + '/files?fields=id', JSON.stringify({ name: nombre, mimeType: CARPETA, parents: [padre] }), 'application/json; charset=UTF-8')
      .then(function (r) { return r.id; });
  });
}

/** Crea (o reutiliza) Recibos <Edificio>/<Año>/<n-Mes>/RECIBOS. Devuelve el id de RECIBOS. */
function carpetasMes(edificio, anio, mesNum, mesNombre) {
  var ruta = ['Recibos ' + edificio, String(anio), mesNum + '-' + mesNombre, 'RECIBOS'];
  var p = Promise.resolve('root');
  ruta.forEach(function (n) { p = p.then(function (padre) { return carpeta(n, padre); }); });
  return p;
}

/** Sube un PDF (reemplaza el contenido si ya existe uno con el mismo nombre) y lo deja visible con el link. */
function subirPdf(blob, nombre, padre) {
  var campos = 'id,name,webViewLink,webContentLink';
  return buscar(nombre, padre, false).then(function (existe) {
    if (existe) {
      return api('PATCH', UPLOAD + '/files/' + existe.id + '?uploadType=media&fields=' + campos, blob, 'application/pdf');
    }
    var b = 'recibos' + Math.random().toString(36).slice(2);
    var cuerpo = new Blob([
      '--' + b + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify({ name: nombre, parents: [padre], mimeType: 'application/pdf' }) + '\r\n' +
      '--' + b + '\r\nContent-Type: application/pdf\r\n\r\n', blob, '\r\n--' + b + '--'
    ]);
    return api('POST', UPLOAD + '/files?uploadType=multipart&fields=' + campos, cuerpo, 'multipart/related; boundary=' + b);
  }).then(function (f) {
    return api('POST', API + '/files/' + f.id + '/permissions?fields=id', JSON.stringify({ role: 'reader', type: 'anyone' }), 'application/json; charset=UTF-8')
      .then(function () {
        return { id: f.id, nombre: f.name || nombre,
                 link: f.webViewLink || ('https://drive.google.com/file/d/' + f.id + '/view?usp=sharing'),
                 descarga: 'https://drive.google.com/uc?export=download&id=' + f.id };
      });
  });
}

/**
 * Sube todos los PDF. pdfs = [{unidad, nombre, blob}]. onProgreso(texto, pct).
 * Devuelve {carpetaId, archivos: {unidad: {id, link, descarga}}}.
 */
function subirRecibos(pdfs, info, onProgreso) {
  onProgreso = onProgreso || function () {};
  onProgreso('Preparando carpetas en Google Drive…', 0);
  return carpetasMes(info.edificio, info.anio, info.mesNum, info.mesNombre).then(function (padre) {
    var out = { carpetaId: padre, archivos: {} }, i = 0;
    function sig() {
      if (i >= pdfs.length) { onProgreso('Recibos guardados en Google Drive', 1); return out; }
      var p = pdfs[i++];
      onProgreso('Subiendo a Google Drive ' + i + '/' + pdfs.length + ' (Dpto ' + p.unidad + ')…', (i - 1) / pdfs.length);
      return subirPdf(p.blob, p.nombre, padre).then(function (r) { out.archivos[p.unidad] = r; return sig(); });
    }
    return sig();
  });
}

var apiPub = { configurado: configurado, gisListo: gisListo, estado: estado, alCambiar: alCambiar, tokenValido: tokenValido,
  pedirToken: pedirToken, conectar: conectar, desconectar: desconectar, carpetasMes: carpetasMes, subirPdf: subirPdf,
  subirRecibos: subirRecibos, CLAVE: CLAVE };
global.Drive = apiPub;
})(this);
