"""Prueba del envío por correo: genera setiembre (Excel de prueba con correos) y revisa
la lista de envío, el portapapeles (HTML + texto) y la URL de Gmail. Datos fuera del repo."""
import time, subprocess, signal, urllib.parse
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = Path('/workspace/recibos_web')
TEST = Path('/workspace/recibos_web_test')
FOTOS = TEST / 'fotos'
SET = Path('/workspace/sistema_recibos/set_2026')
URL = 'http://127.0.0.1:8765/'

def main():
    srv = subprocess.Popen(['python3', '-m', 'http.server', '8765', '--bind', '127.0.0.1'],
                           cwd=str(ROOT), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(0.6)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
            ctx = browser.new_context(viewport={'width': 1100, 'height': 900}, accept_downloads=True)
            ctx.grant_permissions(['clipboard-read', 'clipboard-write'], origin=URL.rstrip('/'))
            page = ctx.new_page()
            errs = []
            page.on('pageerror', lambda e: errs.append(str(e)))
            page.goto(URL, wait_until='networkidle')
            # window.open falso: guarda la URL (no abre Gmail de verdad)
            page.evaluate("() => { window.__abiertas = []; window.open = (u) => { window.__abiertas.push(u); return {}; }; }")

            page.set_input_files('#fileExcel', str(TEST / 'setiembre-real-correos.xlsx'))
            page.wait_for_selector('text=Excel cargado', timeout=15000)
            page.click('#btnAPaso2')
            page.select_option('#selMes', '2026-09')
            page.wait_for_function("() => window.__RecibosApp.estado.borrador.mes === '2026-09'")
            page.click('#btnAPaso3')
            fotos = sorted([str(f) for f in FOTOS.iterdir() if f.suffix.lower() in ('.jpg', '.jpeg', '.png')])
            fotos += [str(SET / 'recibo-de-agua-2026.09.19.jpg'), str(SET / 'Luz_del_sur_Setiembre.pdf')]
            page.set_input_files('#fileFotos', fotos)
            page.wait_for_function("() => window.__RecibosApp.estado.cola.length && window.__RecibosApp.estado.cola.every(c => c.status !== 'procesando')", timeout=120000)
            page.wait_for_function("() => !!window.__RecibosApp.estado.archivos.agua && !!window.__RecibosApp.estado.archivos.luz", timeout=30000)
            with page.expect_download(timeout=300000):
                page.click('#btnGenerar')
                try:
                    page.wait_for_selector('#btnSiGen', timeout=5000); page.click('#btnSiGen')
                except Exception:
                    pass
            page.wait_for_selector('.envio-fila', timeout=30000)
            filas = page.locator('.envio-fila')
            assert filas.count() == 10, filas.count()
            f101 = filas.nth(0)
            assert 'Dpto 101' in f101.inner_text() and 'S/ 471.99' in f101.inner_text(), f101.inner_text()

            # PDF individual
            with page.expect_download() as d:
                f101.locator('.envio-pdf').click()
            assert d.value.suggested_filename == 'Recibo - 101 - Setiembre 2026.pdf', d.value.suggested_filename
            assert d.value.path().read_bytes()[:4] == b'%PDF'

            # Enviar 101
            f101.locator('.envio-btn').click()
            page.wait_for_selector('#toast.visible', timeout=5000)
            toast = page.inner_text('#toast'); print('TOAST:', toast)
            assert 'Pega el mensaje' in toast and 'Mensaje copiado' in toast
            assert f101.locator('.envio-btn').inner_text().strip() == 'Enviado ✓'
            url = page.evaluate('() => window.__abiertas[0]')
            print('GMAIL:', url)
            q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
            assert url.startswith('https://mail.google.com/mail/?view=cm&fs=1&')
            assert q['to'] == ['dpto101@prueba.test'], q
            assert q['su'] == ['Recibo de Gastos Comunes - Setiembre 2026 - Dpto 101'], q
            assert 'body' not in q
            clip = page.evaluate("""async () => {
                const items = await navigator.clipboard.read(); const out = {};
                for (const it of items) for (const t of it.types) out[t] = await (await it.getType(t)).text();
                return out; }""")
            html, txt = clip.get('text/html', ''), clip.get('text/plain', '')
            print('CLIP TEXT:\n' + txt)
            for s in ['Edificio Nápoles II', 'Hola, (Dpto 101)', '#5b5fe0',
                      'correspondiente al mes de <b>Setiembre 2026</b>.', 'Total a pagar: S/ 471.99']:
                assert s in html, s
            assert 'Total a pagar: S/ 471.99' in txt and 'Hola, (Dpto 101)' in txt

            # Dpto 102 con 2 correos
            filas.nth(1).locator('.envio-btn').click()
            url2 = page.evaluate('() => window.__abiertas[1]')
            q2 = urllib.parse.parse_qs(urllib.parse.urlparse(url2).query)
            assert q2['to'] == ['dpto102@prueba.test,otro102@prueba.test'], q2
            assert q2['su'] == ['Recibo de Gastos Comunes - Setiembre 2026 - Dpto 102']
            # mailto
            href = filas.nth(0).locator('.envio-mailto').get_attribute('href')
            assert href.startswith('mailto:dpto101@prueba.test?subject=Recibo%20de%20Gastos%20Comunes'), href
            # Correo borrado -> "sin correo", y escribir uno nuevo
            inp = filas.nth(2).locator('.envio-correo')
            inp.fill(''); assert 'sin correo' in filas.nth(2).inner_text()
            inp.fill('nuevo@prueba.test'); assert 'sin correo' not in filas.nth(2).inner_text()

            # Vista previa del 101 para la captura
            filas.nth(0).locator('.envio-ver').click()
            page.wait_for_timeout(600)
            page.evaluate("() => document.getElementById('toast').classList.remove('visible')")
            page.wait_for_timeout(500)
            page.locator('#panelEnvio').screenshot(path=str(ROOT / 'captura-envio.png'))
            # Móvil
            page.set_viewport_size({'width': 390, 'height': 844})
            filas.nth(3).locator('.envio-copiar').click()
            page.wait_for_selector('#toast.visible')
            page.screenshot(path=str(ROOT / 'captura-envio-movil-toast.png'))
            page.evaluate("() => document.getElementById('toast').classList.remove('visible')")
            page.wait_for_timeout(500)
            page.locator('#panelEnvio').screenshot(path=str(ROOT / 'captura-envio-movil.png'))
            assert not errs, errs
            browser.close()
            print('CORREO OK')
    finally:
        srv.send_signal(signal.SIGTERM); srv.wait(timeout=5)

if __name__ == '__main__':
    main()
