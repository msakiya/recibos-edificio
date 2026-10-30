"""Plantillas públicas (datos ficticios): descarga 'Excel de ejemplo' y 'Excel en blanco', sube el de ejemplo,
genera los 10 PDF (sin ZIP), revisa el estado SIN Drive (aviso y botones), conecta Drive (simulado) desde el
resultado, sube, y prueba el WhatsApp en celular (wa.me). Guarda capturas ficticias para la guía en guia/."""
import time, subprocess, signal, tempfile, sys, urllib.parse
from pathlib import Path
from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright
sys.path.insert(0, str(Path(__file__).parent))
from mock_google import DriveFalso, instalar

ROOT = Path('/workspace/recibos_web')
GUIA = ROOT / 'guia'; GUIA.mkdir(exist_ok=True)
TMP = Path(tempfile.mkdtemp())
URL = 'http://127.0.0.1:8767/'

def img(nombre, texto, color):
    im = Image.new('RGB', (800, 1000), color); d = ImageDraw.Draw(im)
    d.text((40, 40), texto, fill='black'); p = TMP / nombre; im.save(p, quality=85); return str(p)

def sin_toast(page):
    page.evaluate("() => { const t = document.getElementById('toast'); if (t) t.classList.remove('visible'); }")
    page.wait_for_timeout(400)

def cargar_y_generar(page, excel):
    page.set_input_files('#fileExcel', str(excel))
    page.wait_for_selector('#btnAPaso3b', timeout=15000)
    page.click('#btnAPaso3b')
    fotos = [img('%d.jpg' % u, 'Medidor %d (ejemplo)' % u, (230, 240, 250)) for u in (101, 102, 201, 202, 301, 302, 401, 402, 501, 502)]
    fotos += [img('recibo-de-agua-ejemplo.jpg', 'Recibo de agua (ejemplo)', (220, 235, 255)),
              img('recibo-de-luz-ejemplo.jpg', 'Recibo de luz (ejemplo)', (255, 245, 210))]
    page.set_input_files('#fileFotos', fotos)
    page.wait_for_function("() => window.__RecibosApp.estado.cola.length === 12 && window.__RecibosApp.estado.cola.every(c => c.status === 'ok')", timeout=60000)
    page.wait_for_function("() => !!window.__RecibosApp.estado.archivos.agua && !!window.__RecibosApp.estado.archivos.luz")
    page.click('#btnGenerar')
    try: page.wait_for_selector('#btnSiGen', timeout=4000); page.click('#btnSiGen')
    except Exception: pass
    page.wait_for_selector('.envio-fila', timeout=300000)

def main():
    srv = subprocess.Popen(['python3', '-m', 'http.server', '8767', '--bind', '127.0.0.1'], cwd=str(ROOT),
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(0.6)
    drive = DriveFalso(); drive.emails[0] = 'edificio.ejemplo@gmail.com'
    try:
        with sync_playwright() as p:
            b = p.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
            # --- 0) Página real sin GOOGLE_CLIENT_ID: el botón avisa que falta configurar
            c0 = b.new_context(viewport={'width': 1100, 'height': 900})
            c0.route('https://accounts.google.com/gsi/client', lambda r, q: r.fulfill(status=200, content_type='text/javascript', body=''))
            p0 = c0.new_page(); p0.goto(URL, wait_until='networkidle')
            assert 'Falta configurar' in p0.inner_text('#barraDrive0') and p0.locator('#btnConectar').is_disabled()
            c0.close()

            ctx = b.new_context(accept_downloads=True, viewport={'width': 1100, 'height': 900})
            ctx.grant_permissions(['clipboard-read', 'clipboard-write'], origin=URL.rstrip('/'))
            instalar(ctx, drive)
            page = ctx.new_page()
            errs, descargas = [], []
            page.on('pageerror', lambda e: errs.append(str(e)))
            page.on('download', lambda d: descargas.append(d.suggested_filename))
            page.on('dialog', lambda d: d.accept())
            page.goto(URL, wait_until='networkidle')
            page.evaluate("() => { window.__abiertas = []; window.open = (u) => { window.__abiertas.push(u); return {}; }; }")
            page.wait_for_selector('#btnConectar:not([disabled])')
            page.screenshot(path=str(GUIA / 'g01-inicio.png'))
            files = {}
            for etiqueta in ['Excel de ejemplo', 'Excel en blanco']:
                with page.expect_download() as d:
                    page.click('a:has-text("%s")' % etiqueta)
                dst = TMP / d.value.suggested_filename; d.value.save_as(str(dst)); files[etiqueta] = dst
            page.set_input_files('#fileExcel', str(files['Excel en blanco']))
            page.wait_for_selector('text=Falta completar el Excel')
            assert 'celda B4' in page.inner_text('#excelInfo')
            descargas.clear()

            # --- 1) Sin Drive: generar, PDFs individuales, avisos
            cargar_y_generar(page, files['Excel de ejemplo'])
            assert descargas == [], descargas
            t = page.inner_text('#panelResultado')
            assert 'Conecta Google Drive para enviar el link del recibo' in t and 'S/ 471.99' in t
            filas = page.locator('.envio-fila'); assert filas.count() == 10
            for i in range(10):
                with page.expect_download() as d:
                    filas.nth(i).locator('.envio-pdf').click()
                assert d.value.path().read_bytes()[:4] == b'%PDF'
            assert descargas[0] == 'Recibo - 101 - Setiembre 2026.pdf' and len(descargas) == 10, descargas
            filas.nth(0).locator('.envio-btn').click()
            page.wait_for_selector('#toast.visible:has-text("Conecta Google Drive")')
            assert 'Conecta Google Drive para enviar el link del recibo' in page.inner_text('#toast')
            assert page.evaluate('() => window.__abiertas.length') == 0
            for u in (3, 6):   # 202 y 401 sin teléfono en el ejemplo
                w = filas.nth(u).locator('.wa-btn'); assert w.is_disabled() and 'no disponible' in w.inner_text()
            sin_toast(page)
            page.locator('#panelResultado').screenshot(path=str(GUIA / 'g05-sin-drive.png'))

            # --- 2) Conectar desde el resultado → sube solo
            page.click('#btnConectarRes')
            page.wait_for_selector('text=Los 10 recibos están en Google Drive', timeout=60000)
            assert len(drive.pdfs()) == 10 and 'edificio.ejemplo@gmail.com' in page.inner_text('#barraDrive0')
            assert drive.ruta(sorted(drive.pdfs(), key=lambda f: f['name'])[0]) == 'Recibos Edificio Ejemplo/2026/9-Setiembre/RECIBOS/Recibo - 101 - Setiembre 2026.pdf'
            sin_toast(page)
            page.evaluate('() => window.scrollTo(0,0)')
            page.screenshot(path=str(GUIA / 'g03-conectado.png'))
            page.locator('#panelDrive').screenshot(path=str(GUIA / 'g06-drive-ok.png'))
            page.locator('.lista-envio').screenshot(path=str(GUIA / 'g07-lista.png'))
            filas.nth(1).locator('.envio-btn').click(); page.wait_for_selector('#toast.visible:has-text("Ctrl+V")')
            page.locator('#toast').screenshot(path=str(GUIA / 'g08-toast.png'))
            sin_toast(page)
            page.locator('.correo-prev').screenshot(path=str(GUIA / 'g09-correo.png'))
            filas.nth(1).locator('.wa-btn').click()
            u = page.evaluate('() => window.__abiertas[window.__abiertas.length-1]')
            assert u.startswith('https://web.whatsapp.com/send?phone=51999001102&text='), u
            sin_toast(page)
            page.locator('.wa-prev').screenshot(path=str(GUIA / 'g10-whatsapp.png'))
            page.locator('.envio-fila').nth(1).screenshot(path=str(GUIA / 'g11-fila-enviada.png'))
            page.evaluate('() => window.scrollTo(0,0)')
            page.locator('#barraDrive0').screenshot(path=str(GUIA / 'g12-desconectar.png'))
            page.locator('header.app').screenshot(path=str(GUIA / 'g13-cabecera.png'))
            assert not errs, errs

            # --- 3) Celular: WhatsApp por wa.me y correo por mailto
            m = b.new_context(viewport={'width': 390, 'height': 844}, is_mobile=True, has_touch=True, accept_downloads=True,
                user_agent='Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1')
            instalar(m, drive)
            mp = m.new_page(); merrs = []; mp.on('pageerror', lambda e: merrs.append(str(e)))
            mp.goto(URL, wait_until='networkidle')
            mp.evaluate("() => { window.__abiertas = []; window.open = (u) => { window.__abiertas.push(u); return {}; }; }")
            mp.screenshot(path=str(GUIA / 'g02-movil-inicio.png'))
            mp.click('#btnConectar'); mp.wait_for_selector('text=Google Drive conectado')
            cargar_y_generar(mp, files['Excel de ejemplo'])
            mp.wait_for_selector('text=Los 10 recibos están en Google Drive', timeout=60000)
            assert all(f.get('versiones') == 2 for f in drive.pdfs())   # reemplaza, no duplica
            mf = mp.locator('.envio-fila')
            mf.nth(0).locator('.wa-btn').click()
            u = mp.evaluate('() => window.__abiertas[0]')
            assert u.startswith('https://wa.me/51999001101?text='), u
            assert 'Puedes ver tu recibo aquí 👉 https://drive.google.com/file/d/' in urllib.parse.unquote(u)
            # Correo en celular: abre la app de correo (mailto, texto simple con el link) y marca Enviado
            mf.nth(0).locator('.envio-btn').click()
            mp.wait_for_function("() => document.querySelector('.envio-fila .envio-btn').textContent.trim().endsWith('Enviado ✓')")
            sin_toast(mp)
            mp.evaluate('() => window.scrollTo(0,0)')
            mp.screenshot(path=str(GUIA / 'g14-movil-conectado.png'))
            mp.locator('.envio-fila').nth(0).screenshot(path=str(GUIA / 'g15-movil-fila.png'))
            assert not merrs, merrs
            b.close()
            print('PLANTILLA + SIN DRIVE + CELULAR OK')
    finally:
        srv.send_signal(signal.SIGTERM); srv.wait(timeout=5)

if __name__ == '__main__':
    main()
