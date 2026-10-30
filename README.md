# Recibos del edificio

Página web **estática** (sin servidor) para generar los recibos mensuales de un edificio a partir de un Excel.

- Todo el cálculo y los PDF se hacen **en tu navegador**. Solo se envían a internet los PDF que subes a *tu* Google Drive (si lo conectas).
- Interfaz en español, en 3 pasos, usable en celular.
- Se publica gratis con **GitHub Pages**.

## Cómo usarla

En la página hay dos descargas: **Excel de ejemplo** (todas las celdas con números de muestra inventados; las amarillas son las que se llenan cada mes) y **Excel en blanco** (misma estructura, vacía).

1. **Cada mes, en el Excel:** llena la pestaña **“Mes a facturar”** (es la primera):
   - B4 mes (ej. `2026-10`), B5 importe del recibo de agua, B6 m³ facturados, B7 importe del recibo de luz,
     B8/B9 fechas de emisión y vencimiento (opcionales), B10–B12 fondo de contingencia.
   - Sección **LECTURAS DE AGUA**: la lectura **actual** de cada medidor (columna D). La anterior ya viene puesta.
   - Sección **GASTOS COMUNES DEL MES**: montos totales del edificio (los fijos ya vienen copiados; agrega los EXTRA).
   - Sección **AJUSTES / CRÉDITOS**: opcional (crédito = monto negativo).
2. **Paso 1 en la página:** sube el Excel. La página lee todo y muestra un resumen (no hay que escribir nada).
   Si falta algo, dice exactamente qué celda llenar.
3. **Paso 2:** arrastra las fotos de los medidores y los recibos de luz/agua (JPG, PNG o PDF) y pulsa **Generar recibos**.
4. Aparece la lista de departamentos (sin ZIP): cada uno con **⬇ PDF**, **✉ Correo** y **🟢 WhatsApp**.
   Descarga también **“Descargar Excel actualizado (para el próximo mes)”**: historial + pestaña “Mes a facturar”
   lista para el mes siguiente. **Guarda ese Excel**: es el que llenarás y subirás el próximo mes.

## Google Drive (un botón) y envío por correo / WhatsApp

Guía ilustrada: [`Guia-Conectar-Google-Drive.pdf`](Guia-Conectar-Google-Drive.pdf).

- **Conectar Google Drive** (arriba): elige la cuenta y acepta el permiso. Permiso único `drive.file`: la página solo ve
  los archivos que ella crea. Se recuerda en el navegador el correo y que está conectado (el token no se guarda).
  **Desconectar** quita el permiso y permite conectar otra cuenta.
- Al generar, sube cada PDF a `Recibos <Edificio>/<Año>/<n-Mes>/RECIBOS/Recibo - 101 - <Mes Año>.pdf`
  (reutiliza las carpetas y reemplaza el PDF si ya existe) y lo deja como “cualquiera con el link puede ver”.
- **✉ Correo**: en computadora copia el mensaje con formato (botón grande **VER RECIBO**, cuenta, CCI, titular,
  “Importante: Enviar constancia…”) y abre Gmail con destinatario y asunto: pegar con Ctrl+V y enviar.
  En celular abre la app de correo con el mensaje en texto simple y el link.
- **🟢 WhatsApp**: abre WhatsApp Web (computadora) o wa.me (celular) con el mensaje y el link. El número sale de la
  pestaña Departamentos (9 dígitos → 51XXXXXXXXX). Sin número: botón gris “WhatsApp no disponible”.
- Sin Drive conectado se pueden descargar los PDF, y los botones de envío avisan “Conecta Google Drive para enviar el link del recibo”.
- Config del Excel: `TITULAR_CUENTA` y `WHATSAPP_CONSTANCIA` (vacío = celular del tesorero).

### Configuración única (dueño de la página)

En `js/config.js` va `GOOGLE_CLIENT_ID` (cliente OAuth “Aplicación web” de Google Cloud, con la API de Google Drive
habilitada, origen JavaScript autorizado `https://msakiya.github.io`, scope `.../auth/drive.file`, pantalla de
consentimiento publicada “En producción”). No es secreto.

## Cálculo (validado)

Misma lógica que la hoja de Google anterior. Ejemplo setiembre 2026:

- Dpto 101 = **S/ 471.99**
- Dpto 102 = **S/ 597.80**
- Total facturado = **S/ 4,987.48**
- Fondo de contingencia = **S/ 4,897.67**
- El céntimo de redondeo del agua va al departamento de mayor consumo.

## Privacidad del repositorio

Este repo es **público** (GitHub Pages gratis lo requiere en cuentas free). La `plantilla-ejemplo.xlsx` usa **nombres y cuentas inventados**. No subas tu Excel real con datos de propietarios ni cuentas bancarias.

## Desarrollo

```bash
# servir local
python3 -m http.server 8765

# regenerar plantillas (datos ficticios) y darles formato
node tests/crear_plantillas.js && python3 tests/estilo_plantillas.py

# pruebas de cálculo
node tests/test_calculo.js

# plantillas descargables + 10 PDF sin ZIP + sin Drive / con Drive simulado / celular (capturas para la guía)
python3 tests/test_plantilla.py

# punta a punta (Chrome + Playwright; necesita fotos en /workspace/recibos_web_test)
python3 tests/test_e2e.py

# Drive simulado + correo + WhatsApp + desconectar (datos reales fuera del repo)
python3 tests/test_envio.py
node tests/test_correo_unit.js

# guía PDF (usa las capturas ficticias que deja test_plantilla.py en guia/)
python3 tests/crear_guia.py
```

Librerías incluidas en `vendor/` (sin CDN): SheetJS 0.20.3, jsPDF 2.5.2, pdf.js 3.11.174, html2canvas 1.4.1. Google Identity Services se carga desde accounts.google.com.

## Licencias de terceros

Ver `vendor/licencias/`.
