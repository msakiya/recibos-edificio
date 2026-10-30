"""Simulación de Google Identity Services + Drive REST para pruebas con Playwright (sin red ni cuenta real)."""
import json, re, urllib.parse, itertools

GIS_JS = r"""
(function(){
  window.__gis = { reqs: [], revoked: [], n: 0, cuenta: 0 };
  window.google = { accounts: { oauth2: {
    initTokenClient: function (cfg) {
      window.__gis.cfg = { client_id: cfg.client_id, scope: cfg.scope };
      return { requestAccessToken: function (opc) {
        window.__gis.reqs.push(opc || {});
        var n = ++window.__gis.n;
        setTimeout(function () {
          if (window.__gis.fallar) { cfg.error_callback && cfg.error_callback({ type: 'popup_closed' }); return; }
          cfg.callback({ access_token: 'tok-' + n, expires_in: 3599, scope: cfg.scope, token_type: 'Bearer' });
        }, 80);
      } };
    },
    hasGrantedAllScopes: function () { return true; },
    revoke: function (t, cb) { window.__gis.revoked.push(t); cb && cb({ successful: true }); }
  } } };
})();
"""

CONFIG_JS = "window.RECIBOS_CONFIG = { GOOGLE_CLIENT_ID: '123-prueba.apps.googleusercontent.com', DRIVE_SCOPE: 'https://www.googleapis.com/auth/drive.file' };"


class DriveFalso:
    """Guarda carpetas/archivos en memoria y responde como la API de Drive v3."""
    def __init__(self):
        self.files = {}          # id -> {id,name,mimeType,parents,size}
        self.perms = {}          # id -> [perm]
        self.log = []            # (method, path)
        self.tokens = []
        self.emails = ['tesoreria.prueba@gmail.com']
        self._ids = itertools.count(1)

    def _nuevo_id(self):
        return 'f%04d' % next(self._ids)

    def _vista(self, f):
        d = dict(f)
        if f['mimeType'] != 'application/vnd.google-apps.folder':
            d['webViewLink'] = 'https://drive.google.com/file/d/%s/view?usp=drivesdk' % f['id']
            d['webContentLink'] = 'https://drive.google.com/uc?id=%s&export=download' % f['id']
        return d

    def handler(self, route, request):
        url = urllib.parse.urlparse(request.url)
        qs = urllib.parse.parse_qs(url.query)
        m = request.method
        self.log.append((m, url.path, url.query))
        self.tokens.append(request.headers.get('authorization', ''))
        def ok(obj, status=200):
            route.fulfill(status=status, content_type='application/json', body=json.dumps(obj),
                          headers={'Access-Control-Allow-Origin': '*'})
        if m == 'OPTIONS':
            return route.fulfill(status=204, headers={'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': '*',
                                                       'Access-Control-Allow-Methods': '*'})
        if not request.headers.get('authorization', '').startswith('Bearer tok-'):
            return ok({'error': 'unauthorized'}, 401)
        p = url.path
        if p == '/drive/v3/about':
            return ok({'user': {'emailAddress': self.emails[0], 'displayName': 'Prueba'}})
        if p == '/drive/v3/files' and m == 'GET':
            q = qs['q'][0]
            name = re.search(r"name = '((?:[^'\\]|\\.)*)'", q).group(1).replace("\\'", "'").replace('\\\\', '\\')
            parent = re.search(r"'([^']+)' in parents", q).group(1)
            folder = "mimeType = 'application/vnd.google-apps.folder'" in q
            res = [self._vista(f) for f in self.files.values() if f['name'] == name and parent in f['parents']
                   and (f['mimeType'] == 'application/vnd.google-apps.folder') == folder]
            return ok({'files': res})
        if p == '/drive/v3/files' and m == 'POST':
            meta = json.loads(request.post_data)
            f = {'id': self._nuevo_id(), 'name': meta['name'], 'mimeType': meta['mimeType'], 'parents': meta.get('parents', ['root'])}
            self.files[f['id']] = f
            return ok(self._vista(f))
        if p == '/upload/drive/v3/files' and m == 'POST':
            assert qs.get('uploadType') == ['multipart'], qs
            body = request.post_data_buffer
            meta = json.loads(re.search(rb'\r\n\r\n(\{.*?\})\r\n--', body, re.S).group(1))
            assert b'%PDF' in body
            f = {'id': self._nuevo_id(), 'name': meta['name'], 'mimeType': 'application/pdf', 'parents': meta['parents'], 'size': len(body), 'versiones': 1}
            self.files[f['id']] = f
            return ok(self._vista(f))
        mm = re.match(r'/upload/drive/v3/files/([^/]+)$', p)
        if mm and m == 'PATCH':
            f = self.files[mm.group(1)]
            assert request.post_data_buffer[:4] == b'%PDF'
            f['versiones'] = f.get('versiones', 1) + 1
            return ok(self._vista(f))
        mm = re.match(r'/drive/v3/files/([^/]+)/permissions$', p)
        if mm and m == 'POST':
            self.perms.setdefault(mm.group(1), []).append(json.loads(request.post_data))
            return ok({'id': 'anyoneWithLink'})
        return ok({'error': 'no simulado: %s %s' % (m, p)}, 404)

    # ayudas
    def carpetas(self):
        return [f for f in self.files.values() if f['mimeType'] == 'application/vnd.google-apps.folder']
    def pdfs(self):
        return [f for f in self.files.values() if f['mimeType'] == 'application/pdf']
    def ruta(self, f):
        partes = [f['name']]
        par = f['parents'][0]
        while par != 'root':
            partes.insert(0, self.files[par]['name']); par = self.files[par]['parents'][0]
        return '/'.join(partes)


def instalar(ctx, drive, configurado=True):
    ctx.route('https://accounts.google.com/gsi/client', lambda r, q: r.fulfill(status=200, content_type='text/javascript', body=GIS_JS))
    if configurado:
        ctx.route('**/js/config.js', lambda r, q: r.fulfill(status=200, content_type='text/javascript', body=CONFIG_JS))
    ctx.route('https://www.googleapis.com/**', drive.handler)
