"""Plantillas públicas: descarga 'Excel de ejemplo' y 'Excel en blanco' desde la página,
sube el de ejemplo, revisa el resumen y genera los 10 PDF (con imágenes inventadas)."""
import time, subprocess, signal, zipfile, tempfile
from pathlib import Path
from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

ROOT = Path('/workspace/recibos_web')
TMP = Path(tempfile.mkdtemp())

def img(nombre, texto, color):
    im = Image.new('RGB', (800, 1000), color); d = ImageDraw.Draw(im)
    d.text((40, 40), texto, fill='black'); p = TMP / nombre; im.save(p, quality=85); return str(p)

def main():
    srv = subprocess.Popen(['python3', '-m', 'http.server', '8767', '--bind', '127.0.0.1'], cwd=str(ROOT),
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(0.6)
    try:
        with sync_playwright() as p:
            b = p.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
            page = b.new_context(accept_downloads=True, viewport={'width': 1100, 'height': 900}).new_page()
            errs = []; page.on('pageerror', lambda e: errs.append(str(e)))
            page.goto('http://127.0.0.1:8767/', wait_until='networkidle')
            files = {}
            for etiqueta in ['Excel de ejemplo', 'Excel en blanco']:
                with page.expect_download() as d:
                    page.click('a:has-text("%s")' % etiqueta)
                dst = TMP / d.value.suggested_filename; d.value.save_as(str(dst)); files[etiqueta] = dst
                print('descarga:', etiqueta, '→', d.value.suggested_filename)
            # En blanco: pide el mes (B4)
            page.set_input_files('#fileExcel', str(files['Excel en blanco']))
            page.wait_for_selector('text=Falta completar el Excel')
            assert 'celda B4' in page.inner_text('#excelInfo')
            # Ejemplo: resumen válido
            page.set_input_files('#fileExcel', str(files['Excel de ejemplo']))
            page.wait_for_selector('#btnAPaso3b', timeout=15000)
            t = page.inner_text('#excelInfo')
            for s in ['Setiembre 2026', 'S/ 679.40', 'S/ 510.40', 'S/ 471.99', 'S/ 527.62', 'S/ 4,987.48', 'S/ 4,777.67', '25/09/2026', 'Juan Demo']:
                assert s in t, s
            page.click('#btnAPaso3b')
            fotos = [img('%d.jpg' % u, 'Medidor %d (ejemplo)' % u, (230, 240, 250)) for u in (101, 102, 201, 202, 301, 302, 401, 402, 501, 502)]
            fotos += [img('recibo-de-agua-ejemplo.jpg', 'Recibo de agua (ejemplo)', (220, 235, 255)),
                      img('recibo-de-luz-ejemplo.jpg', 'Recibo de luz (ejemplo)', (255, 245, 210))]
            page.set_input_files('#fileFotos', fotos)
            page.wait_for_function("() => window.__RecibosApp.estado.cola.length === 12 && window.__RecibosApp.estado.cola.every(c => c.status === 'ok')", timeout=60000)
            page.wait_for_function("() => !!window.__RecibosApp.estado.archivos.agua && !!window.__RecibosApp.estado.archivos.luz")
            with page.expect_download(timeout=300000) as dz:
                page.click('#btnGenerar')
                try: page.wait_for_selector('#btnSiGen', timeout=4000); page.click('#btnSiGen')
                except Exception: pass
            z = TMP / 'ejemplo.zip'; dz.value.save_as(str(z))
            with zipfile.ZipFile(z) as zf:
                pdfs = [n for n in zf.namelist() if n.endswith('.pdf')]
                assert len(pdfs) == 10, pdfs
                assert all(zf.read(n)[:4] == b'%PDF' for n in pdfs)
                assert sum(n.endswith('.xlsx') for n in zf.namelist()) == 1
            page.wait_for_selector('.envio-fila')
            assert not errs, errs
            b.close()
            print('PLANTILLA OK')
    finally:
        srv.send_signal(signal.SIGTERM); srv.wait(timeout=5)

if __name__ == '__main__':
    main()
