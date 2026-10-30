"""Crea Guia-Conectar-Google-Drive.pdf (A4, letra grande) con capturas FICTICIAS de guia/ (las genera
tests/test_plantilla.py) e ilustraciones de las pantallas de Google. Uso: python3 tests/crear_guia.py"""
from pathlib import Path
from PIL import Image
from playwright.sync_api import sync_playwright

ROOT = Path('/workspace/recibos_web'); G = ROOT / 'guia'
URL_PAGINA = 'https://msakiya.github.io/recibos-edificio/'

def recorte(src, dst, caja):
    im = Image.open(G / src); im.crop(caja).save(G / dst); return dst

recorte('g01-inicio.png', 'c-inicio.png', (0, 0, 1100, 150))
recorte('g02-movil-inicio.png', 'c-movil.png', (0, 0, 390, 430))
recorte('g03-conectado.png', 'c-conectado.png', (60, 64, 1040, 150))
im = Image.open(G / 'g07-lista.png'); recorte('g07-lista.png', 'c-lista.png', (0, 0, im.width, min(im.height, 176)))
im = Image.open(G / 'g05-sin-drive.png'); recorte('g05-sin-drive.png', 'c-sin-drive.png', (0, 0, im.width, 250))

ILU = '<div class="ilu">🖼 Ilustración. La pantalla real de Google se verá parecida; esta captura se actualizará cuando la página tenga su Client ID configurado.</div>'

CSS = """
@page { size: A4; margin: 14mm 14mm 14mm 14mm; }
* { box-sizing: border-box; }
body { font-family: Arial, Helvetica, sans-serif; color: #1f2a2a; font-size: 16px; line-height: 1.4; margin: 0; }
.pag { page-break-after: always; }
.pag:last-child { page-break-after: auto; }
h1 { color: #1d5e5a; font-size: 36px; margin: 0 0 6px; line-height: 1.15; }
h2 { color: #1d5e5a; font-size: 23px; margin: 14px 0 6px; display: flex; align-items: center; gap: 12px; }
.num { background: #e07b39; color: #fff; border-radius: 50%; width: 42px; height: 42px; display: inline-grid; place-items: center; font-size: 22px; flex: 0 0 auto; }
p { margin: 6px 0 10px; }
.cap { border: 1px solid #cfd8d6; border-radius: 10px; max-width: 100%; display: block; margin: 8px auto 12px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
.flecha { color: #e07b39; font-weight: bold; }
.caja { background: #eef6f5; border-left: 6px solid #1d5e5a; border-radius: 8px; padding: 10px 14px; margin: 10px 0; }
.aviso { background: #fff6e0; border-left: 6px solid #e0a000; border-radius: 8px; padding: 10px 14px; margin: 10px 0; }
.ilu { font-size: 12px; color: #8a5a00; background: #fff6e0; border: 1px dashed #e0b040; border-radius: 6px; padding: 4px 8px; margin: 4px auto 10px; max-width: 520px; text-align: center; }
.portada { background: #1d5e5a; color: #fff; border-radius: 16px; padding: 20px 26px; margin-bottom: 14px; }
.portada h1 { color: #fff; font-size: 40px; }
.portada p { font-size: 19px; opacity: .95; }
.dos { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }
ul.grande li { margin: 6px 0; }
.boton-dib { display: inline-block; background: #1d5e5a; color: #fff; border-radius: 10px; padding: 6px 14px; font-weight: bold; }
.boton-gris { display: inline-block; background: #eef2f1; color: #333; border-radius: 10px; padding: 6px 14px; font-weight: bold; }
/* Ilustraciones de Google */
.goo { width: 420px; margin: 4px auto 2px; border: 1px solid #dadce0; border-radius: 10px; padding: 14px 22px; font-family: Arial, sans-serif; background: #fff; box-shadow: 0 2px 10px rgba(0,0,0,.08); }
.goo .g { font-size: 22px; font-weight: bold; letter-spacing: -1px; }
.goo .g b:nth-child(1){color:#4285f4}.goo .g b:nth-child(2){color:#ea4335}.goo .g b:nth-child(3){color:#fbbc05}.goo .g b:nth-child(4){color:#4285f4}.goo .g b:nth-child(5){color:#34a853}.goo .g b:nth-child(6){color:#ea4335}
.goo h3 { font-weight: normal; font-size: 20px; margin: 8px 0 2px; color: #202124; }
.goo .sub { color: #5f6368; font-size: 14px; margin-bottom: 12px; }
.cta { display: flex; align-items: center; gap: 12px; border-top: 1px solid #e8eaed; padding: 7px 4px; font-size: 15px; }
.cta .av { width: 32px; height: 32px; border-radius: 50%; color: #fff; display: grid; place-items: center; font-weight: bold; }
.cta small { display: block; color: #5f6368; font-size: 13px; }
.cta.sel { background: #e8f0fe; border-radius: 6px; outline: 3px solid #e07b39; }
.perm { border: 1px solid #e8eaed; border-radius: 8px; padding: 10px 12px; font-size: 14px; display: flex; gap: 10px; align-items: flex-start; margin: 8px 0; }
.chk { width: 18px; height: 18px; background: #1a73e8; border-radius: 3px; color: #fff; font-size: 13px; display: grid; place-items: center; flex: 0 0 auto; }
.gbtns { display: flex; justify-content: flex-end; gap: 10px; margin-top: 14px; }
.gb { border-radius: 18px; padding: 8px 18px; font-size: 14px; border: 1px solid #dadce0; color: #1a73e8; }
.gb.azul { background: #1a73e8; color: #fff; border-color: #1a73e8; outline: 3px solid #e07b39; outline-offset: 2px; }
.drv { width: 520px; margin: 6px auto; border: 1px solid #dadce0; border-radius: 10px; background: #fff; padding: 14px 18px; font-size: 15px; }
.drv .ruta { color: #5f6368; font-size: 14px; margin-bottom: 8px; }
.drv .it { display: flex; gap: 10px; align-items: center; padding: 6px 0; border-bottom: 1px solid #f1f3f4; }
.carp { width: 22px; height: 17px; background: #5f6368; border-radius: 2px 2px 3px 3px; position: relative; flex: 0 0 auto; }
.pdfi { width: 18px; height: 22px; background: #ea4335; border-radius: 2px; color: #fff; font-size: 8px; display: grid; place-items: center; font-weight: bold; flex: 0 0 auto; }
.dlg { width: 470px; margin: 6px auto; border-radius: 10px; background: #fff; border: 1px solid #ccc; padding: 16px 20px; font-size: 15px; box-shadow: 0 4px 16px rgba(0,0,0,.15); }
.dlg .bts { text-align: right; margin-top: 10px; } .dlg .bts span { display: inline-block; padding: 6px 14px; border-radius: 6px; margin-left: 8px; background: #1d5e5a; color: #fff; }
.dlg .bts span.n { background: #eee; color: #333; }
"""

def html():
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><style>{CSS}</style></head><body>

<section class="pag">
  <div class="portada">
    <h1>Conectar Google Drive y enviar los recibos</h1>
    <p>Guía paso a paso · Recibos del edificio</p>
    <p style="font-size:16px">Página: <b>{URL_PAGINA}</b></p>
  </div>
  <div class="caja"><b>¿Para qué sirve?</b> Al conectar Google Drive, la página guarda el PDF de cada departamento en tu Drive
  y prepara un <b>link</b> para que cada propietario vea y descargue su recibo desde el correo o el WhatsApp.
  Se hace <b>una sola vez</b> por computadora o celular.</div>
  <p><b>Necesitas:</b> una cuenta de Google (Gmail) del edificio o la que ustedes elijan. Puede ser la de la tesorera o una cuenta nueva solo para el edificio.</p>

  <h2><span class="num">1</span> Toca “Conectar Google Drive”</h2>
  <p>Abre la página. Arriba verás el cuadro <b>Google Drive</b>. Toca el botón <span class="boton-dib">Conectar Google Drive</span>.</p>
  <img class="cap" src="c-inicio.png" style="width:100%">
  <div class="dos">
    <div><p>En el celular se ve así <span class="flecha">→</span></p>
      <p style="font-size:15px">Si el navegador avisa que “bloqueó una ventana emergente”, elige <b>Permitir</b> y vuelve a tocar el botón.</p></div>
    <img class="cap" src="c-movil.png" style="width:190px">
  </div>
</section>

<section class="pag">
  <h2><span class="num">2</span> Elige la cuenta de Google</h2>
  <p>Se abre una ventanita de Google. Toca la cuenta donde quieren guardar los recibos.
  Si no aparece, toca <b>“Usar otra cuenta”</b> y escribe el correo y la contraseña.</p>
  <div class="goo">
    <div class="g"><b>G</b><b>o</b><b>o</b><b>g</b><b>l</b><b>e</b></div>
    <h3>Elige una cuenta</h3><div class="sub">para ir a <b>Recibos del edificio</b></div>
    <div class="cta sel"><span class="av" style="background:#e07b39">E</span><div>Edificio Ejemplo<small>edificio.ejemplo@gmail.com</small></div></div>
    <div class="cta"><span class="av" style="background:#7b61ff">L</span><div>Luis Gómez<small>luis.gomez@gmail.com</small></div></div>
    <div class="cta"><span class="av" style="background:#9aa0a6">+</span><div>Usar otra cuenta</div></div>
  </div>
  {ILU}

  <h2><span class="num">3</span> Acepta el permiso</h2>
  <p>Google pregunta si la página puede guardar archivos en tu Drive. Si ves una casilla, márcala. Toca <b>Continuar</b> (o <b>Permitir</b>).</p>
  <div class="goo">
    <div class="g"><b>G</b><b>o</b><b>o</b><b>g</b><b>l</b><b>e</b></div>
    <h3 style="font-size:19px">Recibos del edificio quiere acceder a tu Cuenta de Google</h3>
    <div class="sub">edificio.ejemplo@gmail.com</div>
    <div class="perm"><span class="chk">✓</span><div>Ver, editar, crear y borrar <b>solo los archivos específicos de Google Drive que uses con esta app</b></div></div>
    <div class="gbtns"><span class="gb">Cancelar</span><span class="gb azul">Continuar</span></div>
  </div>
  {ILU}
  <div class="caja">🔒 <b>Tranquila:</b> la página <b>solo</b> puede ver los archivos que ella misma crea (los recibos). No ve tus fotos, documentos ni correos.</div>
</section>

<section class="pag">
  <div class="aviso"><b>Solo si aparece “Google no verificó esta app”:</b> toca <b>Avanzado</b> (o “Configuración avanzada”) y luego
  <b>“Ir a Recibos del edificio (no seguro)”</b>. Es normal en páginas pequeñas y gratuitas; normalmente <b>no</b> aparece.</div>
  <div class="goo" style="width:420px;padding:16px 22px">
    <h3 style="font-size:19px;margin-top:0">Google no verificó esta app</h3>
    <div class="sub">La app solicita acceso a información de tu Cuenta de Google…</div>
    <div style="font-size:14px;color:#1a73e8;outline:3px solid #e07b39;display:inline-block;padding:2px 6px;border-radius:4px">Avanzado</div>
    <div style="font-size:14px;color:#1a73e8;margin-top:8px">Ir a Recibos del edificio (no seguro) <span class="flecha">← 2.º</span></div>
  </div>
  {ILU}

  <h2><span class="num">4</span> ¡Listo! Ya está conectado</h2>
  <p>Arriba aparece <b>“Google Drive conectado”</b> con el correo de la cuenta. La página lo recuerda: no tienes que conectarlo cada mes.</p>
  <img class="cap" src="c-conectado.png" style="width:100%">

  <h2><span class="num">5</span> Genera los recibos como siempre</h2>
  <p>Sube el Excel, las fotos y toca <b>Generar recibos</b>. Al terminar, la página <b>sube sola</b> los PDF a tu Drive y muestra:</p>
  <img class="cap" src="g06-drive-ok.png" style="width:100%">
  <p>Debajo está la lista de departamentos con sus botones, y el botón <b>“Descargar Excel actualizado (para el próximo mes)”</b>: guárdalo, es el que usarás el próximo mes.</p>
  <img class="cap" src="c-lista.png" style="width:100%">
</section>

<section class="pag">
  <h2><span class="num">6</span> ¿Dónde quedan los recibos en Drive?</h2>
  <p>En <b>Mi unidad</b> se crea una carpeta con el nombre del edificio, y dentro una por año y por mes:</p>
  <div class="drv">
    <div class="ruta">Mi unidad › Recibos Edificio Ejemplo › 2026 › 9-Setiembre › RECIBOS</div>
    <div class="it"><span class="pdfi">PDF</span> Recibo - 101 - Setiembre 2026.pdf</div>
    <div class="it"><span class="pdfi">PDF</span> Recibo - 102 - Setiembre 2026.pdf</div>
    <div class="it"><span class="pdfi">PDF</span> Recibo - 201 - Setiembre 2026.pdf</div>
    <div class="it" style="color:#5f6368">… un PDF por departamento</div>
  </div>
  <div class="ilu">🖼 Ilustración de Google Drive (las carpetas y nombres son los que crea la página).</div>
  <ul class="grande">
    <li>Cada PDF queda como <b>“Cualquier persona con el link puede ver”</b>, para que el propietario lo abra sin cuenta.</li>
    <li>Si generas el mismo mes otra vez, el PDF se <b>reemplaza</b> (no se duplica) y el link sigue igual.</li>
    <li>Toca <b>“Abrir la carpeta”</b> en la página para ir directo.</li>
  </ul>

  <h2><span class="num">7</span> Enviar por correo</h2>
  <p>Toca <b>✉ Correo</b> en el departamento. Se abre Gmail con el destinatario y el asunto, y el mensaje queda copiado.
  <b>Haz clic en el cuerpo del correo, pega (Ctrl+V) y toca Enviar.</b> En el celular se abre tu app de correo con todo escrito.</p>
  <div class="dos">
    <img class="cap" src="g09-correo.png" style="width:100%">
    <div><img class="cap" src="g08-toast.png" style="width:100%">
      <p style="font-size:15px">El botón grande <b>VER RECIBO</b> abre el PDF del departamento en Drive.</p>
      <p style="font-size:15px">Los datos de la cuenta, el titular y el WhatsApp para la constancia salen de la pestaña <b>Config</b> del Excel.</p></div>
  </div>
</section>

<section class="pag">
  <h2><span class="num">8</span> Enviar por WhatsApp</h2>
  <p>Toca <b>🟢 WhatsApp</b>. Se abre WhatsApp Web (en la computadora) o la app (en el celular) con el mensaje listo; solo toca <b>Enviar</b>.
  El número sale de la pestaña <b>Departamentos</b> del Excel (9 dígitos, ej. 999 001 101).</p>
  <div class="dos">
    <img class="cap" src="g10-whatsapp.png" style="width:100%">
    <div><img class="cap" src="g15-movil-fila.png" style="width:100%">
    <p style="font-size:15px">Después de tocar, el botón dice <b>Enviado ✓</b>, para saber a quién ya le mandaste.</p>
    <p style="font-size:15px">Si un departamento no tiene teléfono en el Excel, el botón sale gris: <span class="boton-gris">WhatsApp no disponible</span>.</p></div>
  </div>

  <h2><span class="num">9</span> Desconectar o cambiar de cuenta</h2>
  <p>Arriba, toca <span class="boton-gris">Desconectar</span> y confirma. La página olvida la cuenta y quita el permiso.
  Los recibos ya guardados <b>siguen</b> en ese Drive. Luego toca <span class="boton-dib">Conectar Google Drive</span> y elige otra cuenta (pasos 1 a 4).</p>
  <img class="cap" src="g12-desconectar.png" style="width:100%">
  <div class="caja"><b>Si no conectas Drive</b> igual puedes generar y descargar cada PDF con <b>⬇ PDF</b>; los botones de correo y WhatsApp te recordarán:
  “Conecta Google Drive para enviar el link del recibo”.</div>
  <p style="font-size:14px;color:#678">¿Cambió la tesorera o el presidente? Solo cambia sus datos en la pestaña Config del Excel. Si la nueva persona usa otra cuenta de Google, desconecta y conecta la suya.</p>
</section>
</body></html>"""

def main():
    (G / 'guia.html').write_text(html(), encoding='utf-8')
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
        pg = b.new_page()
        pg.goto((G / 'guia.html').as_uri(), wait_until='networkidle')
        pg.pdf(path=str(ROOT / 'Guia-Conectar-Google-Drive.pdf'), format='A4', print_background=True,
               margin={'top': '12mm', 'bottom': '12mm', 'left': '12mm', 'right': '12mm'})
        b.close()
    print('guía OK')

if __name__ == '__main__':
    main()
