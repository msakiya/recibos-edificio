"""Da formato a las plantillas (celdas a llenar en amarillo, notas en gris).
SheetJS (versión libre) no escribe estilos, por eso se hace aquí con openpyxl. No cambia celdas ni posiciones."""
import sys
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

AMARILLO = PatternFill('solid', fgColor='FFF7C2')
GRIS = PatternFill('solid', fgColor='EEF1F0')
CAB = PatternFill('solid', fgColor='1D5E5A')
NOTA = Font(color='808080', italic=True, size=10)
BLANCA = Font(color='FFFFFF', bold=True)
TITULO = Font(color='1D5E5A', bold=True, size=14)
SECCION = Font(color='1D5E5A', bold=True, size=12)
FINO = Side(style='thin', color='C9D6D3')
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)

def llenar(c, fill=AMARILLO):
    c.fill = fill; c.border = BORDE

def hoja_mes(ws):
    ws['A1'].font = TITULO
    ws['A2'].font = NOTA
    for f in range(4, 13):
        ws.cell(f, 1).font = Font(bold=True)
        llenar(ws.cell(f, 2), GRIS if f == 10 else AMARILLO)
        ws.cell(f, 3).font = NOTA
    for f in (5, 7, 10, 11):
        ws.cell(f, 2).number_format = '#,##0.00'
    seccion = None
    for f in range(13, ws.max_row + 1):
        a = ws.cell(f, 1).value
        ta = str(a or '').upper()
        if ta.startswith('LECTURAS'): seccion = 'L'; ws.cell(f, 1).font = SECCION; continue
        if ta.startswith('GASTOS'): seccion = 'G'; ws.cell(f, 1).font = SECCION; continue
        if ta.startswith('AJUSTES'): seccion = 'A'; ws.cell(f, 1).font = SECCION; continue
        if ta.startswith('TIPOS'): ws.cell(f, 1).font = NOTA; seccion = None; continue
        if ta in ('DPTO', 'CONCEPTO (TAL COMO SALE EN EL RECIBO)'):
            for c in range(1, 7 if seccion == 'L' else (5 if seccion == 'G' else 4)):
                ws.cell(f, c).fill = CAB; ws.cell(f, c).font = BLANCA
            continue
        if seccion == 'L' and a is not None:
            llenar(ws.cell(f, 3), GRIS); llenar(ws.cell(f, 4)); ws.cell(f, 5).fill = GRIS
            for c in (3, 4, 5): ws.cell(f, c).number_format = '0.000'
            if ws.cell(f - 1, 1).value == 'Dpto':
                ws.cell(f, 7, '← escribe la lectura ACTUAL de cada medidor en la columna D (amarillo)').font = NOTA
        elif seccion == 'G':
            tipo = str(ws.cell(f, 3).value or '')
            if a is None and tipo != 'EXTRA': continue
            llenar(ws.cell(f, 1)); ws.cell(f, 3).border = BORDE
            llenar(ws.cell(f, 2), GRIS if tipo in ('LUZ', 'AGUA_COMUN') else AMARILLO)
            ws.cell(f, 2).number_format = '#,##0.00'
            ws.cell(f, 4).font = NOTA
        elif seccion == 'A':
            if a is None and ws.cell(f, 3).value is None and ws.cell(f - 1, 1).value is None and ws.cell(f - 2, 1).value is None \
               and ws.cell(f + 1, 1).value is None and str(ws.cell(f + 1, 1).value or '').upper().startswith('TIPOS'):
                continue
            if str(ws.cell(f + 1, 1).value or '').upper().startswith('TIPOS') and a is None:
                continue  # fila vacía antes de la nota final
            for c in (1, 2, 3): llenar(ws.cell(f, c))
            ws.cell(f, 3).number_format = '#,##0.00'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.column_dimensions['G'].width = 62
    ws.sheet_view.zoomScale = 100

def hoja_config(ws):
    for c in range(1, 4): ws.cell(1, c).fill = CAB; ws.cell(1, c).font = BLANCA
    for f in range(2, ws.max_row + 1):
        ws.cell(f, 1).font = Font(bold=True)
        llenar(ws.cell(f, 2)); ws.cell(f, 3).font = NOTA

def hoja_deptos(ws):
    for c in range(1, 6): ws.cell(1, c).fill = CAB; ws.cell(1, c).font = BLANCA
    for f in range(2, ws.max_row + 1):
        ws.cell(f, 1).font = Font(bold=True)
        for c in range(2, 6): llenar(ws.cell(f, c))

def estilo(path):
    wb = openpyxl.load_workbook(path)
    hoja_mes(wb['Mes a facturar'])
    hoja_config(wb['Config'])
    hoja_deptos(wb['Departamentos'])
    for n in ('Lecturas', 'Servicios', 'Gastos', 'Ajustes', 'Fondo', 'Pagos'):
        ws = wb[n]
        for c in range(1, ws.max_column + 1):
            ws.cell(1, c).fill = CAB; ws.cell(1, c).font = BLANCA
    wb['Instrucciones']['A1'].font = TITULO
    wb.active = 0
    wb.save(path)

if __name__ == '__main__':
    for p in sys.argv[1:] or ['plantilla-ejemplo.xlsx', 'plantilla-en-blanco.xlsx']:
        estilo(p); print('estilo OK', p)
