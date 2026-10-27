"""Prueba de punta a punta: carga Excel real + fotos + genera ZIP con 10 PDF."""
import os, re, time, zipfile, io, subprocess, signal
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = Path('/workspace/recibos_web')
OUT = ROOT  # capturas y previews aquí
TEST = Path('/workspace/recibos_web_test')
FOTOS = TEST / 'fotos'
SET = Path('/workspace/sistema_recibos/set_2026')

def main():
    # Servidor estático
    srv = subprocess.Popen(['python3', '-m', 'http.server', '8765', '--bind', '127.0.0.1'],
                           cwd=str(ROOT), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(0.6)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path='/usr/bin/google-chrome',
                                        args=['--no-sandbox', '--disable-web-security'])
            ctx = browser.new_context(viewport={'width': 1100, 'height': 900}, accept_downloads=True)
            page = ctx.new_page()
            errs = []
            page.on('pageerror', lambda e: errs.append(str(e)))
            page.goto('http://127.0.0.1:8765/', wait_until='networkidle')
            page.screenshot(path=str(OUT / 'captura-paso1.png'), full_page=True)

            # Sube Excel real
            page.set_input_files('#fileExcel', str(TEST / 'setiembre-real.xlsx'))
            page.wait_for_selector('text=Excel cargado', timeout=15000)
            page.screenshot(path=str(OUT / 'captura-paso1-cargado.png'), full_page=True)
            page.click("#btnAPaso2")
            page.wait_for_selector("#paso2:not(.oculto)")
            page.select_option("#selMes", "2026-09")
            page.wait_for_function("() => window.__RecibosApp.estado.borrador.mes === '2026-09'")
            page.wait_for_selector('#paso2:not(.oculto)')
            # Verifica mes setiembre y total de agua precargado
            assert page.input_value('#aguaImporte') in ('679.4', '679.40')
            page.screenshot(path=str(OUT / 'captura-paso2.png'), full_page=True)
            page.click('#btnAPaso3')
            page.wait_for_selector('#paso3:not(.oculto)')

            # Fotos: medidores + comprobantes + recibos
            fotos = sorted([str(f) for f in FOTOS.iterdir() if f.suffix.lower() in ('.jpg', '.jpeg', '.png')])
            fotos += [str(SET / 'recibo-de-agua-2026.09.19.jpg'), str(SET / 'Luz_del_sur_Setiembre.pdf')]
            page.set_input_files('#fileFotos', fotos)
            page.wait_for_function(
                "() => window.__RecibosApp.estado.cola.length && window.__RecibosApp.estado.cola.every(c => c.status !== 'procesando')",
                timeout=120000)
            # Espera a que agua y luz estén asignados
            page.wait_for_function(
                "() => !!window.__RecibosApp.estado.archivos.agua && !!window.__RecibosApp.estado.archivos.luz",
                timeout=30000)
            page.screenshot(path=str(OUT / 'captura-paso3.png'), full_page=True)
            n_med = page.evaluate("() => Object.keys(window.__RecibosApp.estado.archivos.medidor).filter(k => !k.includes('__')).length")
            print('medidores asignados:', n_med)
            assert n_med >= 9, n_med

            # Generar
            with page.expect_download(timeout=300000) as di:
                page.click('#btnGenerar')
                # Si hay avisos, confirmar
                try:
                    page.wait_for_selector('#btnSiGen', timeout=5000)
                    page.click('#btnSiGen')
                except Exception:
                    pass
            download = di.value
            zip_path = TEST / 'recibos-setiembre.zip'
            download.save_as(str(zip_path))
            print('ZIP:', zip_path, zip_path.stat().st_size)

            page.wait_for_selector('text=Listo:', timeout=30000)
            page.screenshot(path=str(OUT / 'captura-resultado.png'), full_page=True)

            # Inspecciona ZIP
            with zipfile.ZipFile(zip_path) as z:
                names = z.namelist()
                print('ZIP contiene:', names)
                pdfs = [n for n in names if n.endswith('.pdf')]
                xlsx = [n for n in names if n.endswith('.xlsx')]
                assert len(pdfs) == 10, pdfs
                assert any('101' in n and 'Setiembre' in n for n in pdfs), pdfs
                assert len(xlsx) == 1, xlsx
                # Extrae PDF 101
                pdf101 = [n for n in pdfs if '101' in n][0]
                data = z.read(pdf101)
                (TEST / 'recibo-101.pdf').write_bytes(data)
                print('PDF 101 bytes:', len(data))

            assert not errs, errs
            browser.close()
    finally:
        srv.send_signal(signal.SIGTERM)
        srv.wait(timeout=5)

    # Renderiza PDF 101 a PNG con pdftoppm
    subprocess.check_call(['pdftoppm', '-png', '-r', '110', str(TEST / 'recibo-101.pdf'), str(OUT / 'preview-101')])
    # Renombra preview-101-1.png → preview-101-p1.png
    for i, src in enumerate(sorted(OUT.glob('preview-101-*.png')), 1):
        if re.search(r'preview-101-\d+\.png$', src.name):
            dest = OUT / f'preview-101-p{i}.png'
            src.rename(dest)
            print('→', dest)
    print('E2E OK')

if __name__ == '__main__':
    main()
