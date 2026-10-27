# Recibos del edificio

Página web **estática** (sin servidor) para generar los recibos mensuales de un edificio a partir de un Excel.

- Todo el cálculo y los PDF se hacen **en tu navegador**. Nada se envía a internet.
- Interfaz en español, en 3 pasos, usable en celular.
- Se publica gratis con **GitHub Pages**.

## Cómo usarla

1. Abre la página.
2. **Paso 1:** sube tu Excel (o descarga la plantilla de ejemplo / un Excel en blanco).
3. **Paso 2:** elige el mes, escribe las lecturas nuevas, los montos de agua y luz, y revisa los gastos.
4. **Paso 3:** arrastra las fotos de los medidores y los recibos de luz/agua (JPG, PNG o PDF). Pulsa **Generar recibos**.
5. Se descarga un ZIP con un PDF por departamento y el Excel actualizado. **Guarda ese Excel**: es el que subirás el próximo mes.

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
```

Librerías incluidas en `vendor/` (sin CDN): SheetJS 0.20.3, jsPDF 2.5.2, JSZip 3.10.1, pdf.js 3.11.174, html2canvas 1.4.1.

## Licencias de terceros

Ver `vendor/licencias/`.
