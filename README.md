# Recibos del edificio

Página web **estática** (sin servidor) para generar los recibos mensuales de un edificio a partir de un Excel.

- Todo el cálculo y los PDF se hacen **en tu navegador**. Nada se envía a internet.
- Interfaz en español, en 3 pasos, usable en celular.
- Se publica gratis con **GitHub Pages**.

## Cómo usarla

1. **Cada mes, en el Excel:** llena la pestaña **“Mes a facturar”** (es la primera):
   - B4 mes (ej. `2026-10`), B5 importe del recibo de agua, B6 m³ facturados, B7 importe del recibo de luz,
     B8/B9 fechas de emisión y vencimiento (opcionales), B10–B12 fondo de contingencia.
   - Sección **LECTURAS DE AGUA**: la lectura **actual** de cada medidor (columna D). La anterior ya viene puesta.
   - Sección **GASTOS COMUNES DEL MES**: montos totales del edificio (los fijos ya vienen copiados; agrega los EXTRA).
   - Sección **AJUSTES / CRÉDITOS**: opcional (crédito = monto negativo).
2. **Paso 1 en la página:** sube el Excel. La página lee todo y muestra un resumen (no hay que escribir nada).
   Si falta algo, dice exactamente qué celda llenar.
3. **Paso 2:** arrastra las fotos de los medidores y los recibos de luz/agua (JPG, PNG o PDF) y pulsa **Generar recibos**.
4. Se descarga un ZIP con un PDF por departamento y el **Excel actualizado** (historial + pestaña “Mes a facturar”
   lista para el mes siguiente). **Guarda ese Excel**: es el que llenarás y subirás el próximo mes.

## Enviar los recibos por correo

Después de generar, aparece la lista de los departamentos con su total, un botón **✉ Enviar** y **⬇ PDF**.

1. Escribe los correos en la pestaña **Departamentos** del Excel (columna *Correo*; varios correos se separan con coma).
2. Toca **✉ Enviar**: se copia el mensaje con formato y se abre Gmail con el destinatario y el asunto
   (`Recibo de Gastos Comunes - <Mes> <Año> - Dpto <N>`).
3. Pega el mensaje (Ctrl+V o mantener presionado › Pegar), adjunta el PDF descargado con **⬇ PDF** y envía.

¿No usas Gmail? Usa **Otra app de correo** (abre tu app de correo con el mensaje en texto simple).
La página no envía correos sola: no hay servidor.

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

# pruebas de cálculo
node tests/test_calculo.js

# punta a punta (Chrome + Playwright; necesita fotos en /workspace/recibos_web_test)
python3 tests/test_e2e.py

# envío por correo (portapapeles + URL de Gmail)
python3 tests/test_correo.py
```

Librerías incluidas en `vendor/` (sin CDN): SheetJS 0.20.3, jsPDF 2.5.2, JSZip 3.10.1, pdf.js 3.11.174, html2canvas 1.4.1.

## Licencias de terceros

Ver `vendor/licencias/`.
