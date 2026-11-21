"""Plantilla Excel de comunicaciones internas (depositos a otros bancos): generacion y lectura tolerante.

El cajero anota durante el dia, una fila por factura: N° factura | Banco | Codigo de asignacion, y la sube al final del dia.
El importe, el cliente y la fecha NUNCA se digitan: salen del movimiento del sistema de facturacion.
Este modulo solo conoce el archivo (no la base ni el motor); las validaciones de negocio estan en servicio.Almacen.aplicar_plantilla."""
import os, re, unicodedata
import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

from . import catalogo

HOJA = 'COMUNICACIONES INTERNAS'
HOJA_AYUDA = 'INSTRUCCIONES'
ENCABEZADOS = ['N° factura', 'Banco', 'Código de asignación']
BANCOS_PLANTILLA = [b for b in catalogo.LISTA_BANCOS if not catalogo.es_qr(b)]    # BNB, BMSC, BISA, UNION
FILAS_PLANTILLA = 200                    # filas 2..201 con formato de texto
MAX_FILAS_LEIDAS = 2000                  # tope de filas con datos al importar
NOMBRE_DESCARGA = 'Plantilla_Comunicaciones_Internas.xlsx'
RUTA_ESTATICA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'plantilla_comunicaciones_internas.xlsx')

GUINDA = '870B42'
ALFANUM = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'

INSTRUCCIONES = [
    ('Plantilla de comunicaciones internas (depósitos a otros bancos)', 'titulo'),
    ('Use esta plantilla para anotar, durante el día, las facturas que el cliente pagó con un depósito a OTRO banco (BNB, BMSC, BISA o UNION). '
     'Al final del día la sube junto con el movimiento del cajero y el reporte QR BANECO, y el sistema reclasifica esas facturas automáticamente como «Comunicación interna».', 'texto'),
    ('Paso a paso', 'sub'),
    ('1. Cuándo anotar una factura: cada vez que cobre con un depósito hecho a otro banco que no sea QR BANECO y que le entreguen un código de asignación (comprobante o comunicación interna).', 'texto'),
    ('2. Abra la hoja «COMUNICACIONES INTERNAS» (pestaña de abajo) y escriba UNA factura por fila, empezando en la fila 2 (debajo de los títulos).', 'texto'),
    ('3. Columna «N° factura»: el número de la factura tal como sale en el sistema de facturación. Solo el número.', 'texto'),
    ('4. Columna «Banco»: elija de la lista (BNB, BMSC, BISA o UNION). Use el desplegable de la celda; no escriba otros nombres.', 'texto'),
    ('5. Columna «Código de asignación»: el código del depósito, tal como figura en el comprobante. Las celdas están en formato texto: no pierde ceros ni letras.', 'texto'),
    ('6. Al final del día guarde el archivo como .xlsx (Excel: Archivo > Guardar como > Libro de Excel). No use .xls, .csv ni PDF.', 'texto'),
    ('7. En el sistema, en «Nuevo cierre», suba los TRES archivos juntos: movimiento del cajero, reporte QR BANECO y esta plantilla.', 'texto'),
    ('Formato del código por banco (los ejemplos son inventados)', 'sub'),
    ('BNB: letras y números, debe incluir una P.  Ejemplo: P12345 o 7P2T91', 'ejemplo'),
    ('BMSC: letras y números, debe incluir una T.  Ejemplo: T98765 o 4T0B12', 'ejemplo'),
    ('BISA: solo números.  Ejemplo: 4567890', 'ejemplo'),
    ('UNION: solo números.  Ejemplo: 1234567', 'ejemplo'),
    ('Si un código no cumple su formato, la celda se pinta de ámbar. No es un error que impida presentar: revise que lo escribió bien; si es correcto, déjelo así y Ingresos lo verá como alerta de revisión.', 'texto'),
    ('Importante', 'sub'),
    ('• NO anote facturas de QR BANECO ni de tarjeta. Solo depósitos a otros bancos.', 'texto'),
    ('• Una factura por fila. No repita la misma factura ni el mismo código del mismo banco.', 'texto'),
    ('• No borre, mueva ni cambie los títulos de la fila 1, ni el nombre de las hojas.', 'texto'),
    ('• No escriba importe, cliente ni fecha: el importe, el cliente y la fecha se toman del movimiento de Univalle (sistema de facturación). No se digitan.', 'texto'),
    ('• Las filas vacías se ignoran. Si una fila tiene un problema (factura que no existe, anulada, código repetido, etc.), el sistema le mostrará cuál y por qué; las demás filas sí se aplican. '
     'Puede corregir el Excel y volver a subirlo con «Reemplazar plantilla», o resolver esa factura a mano en la pantalla de revisión.', 'texto'),
]


# ---------------------------------------------------------------- generacion
def _formula_formato(fila=2):
    """Formulas de Excel (formato condicional de la columna C) equivalentes a las reglas de catalogo.BANCOS, a partir de la fila dada."""
    c, b = f'$C{fila}', f'$B{fila}'
    cuerpo = f'SUMPRODUCT((ROW($1:$40)<=LEN({c}))*ISNUMBER(FIND(MID(UPPER({c}),ROW($1:$40),1),"{ALFANUM}")))<LEN({c})'
    digitos = f'SUMPRODUCT((ROW($1:$40)<=LEN({c}))*ISNUMBER(FIND(MID({c},ROW($1:$40),1),"0123456789")))<LEN({c})'
    return {
        'BNB': f'AND({b}="BNB",{c}<>"",OR(ISERROR(FIND("P",UPPER({c}))),{cuerpo}))',
        'BMSC': f'AND({b}="BMSC",{c}<>"",OR(ISERROR(FIND("T",UPPER({c}))),{cuerpo}))',
        'NUMERICO': f'AND(OR({b}="BISA",{b}="UNION"),{c}<>"",{digitos})',
    }


def generar_plantilla(ruta):
    """Escribe la plantilla .xlsx vacia (sin filas de ejemplo) en `ruta` y la devuelve."""
    wb = openpyxl.Workbook()
    ayuda = wb.active; ayuda.title = HOJA_AYUDA
    ws = wb.create_sheet(HOJA)

    # --- hoja de datos
    relleno = PatternFill('solid', start_color=GUINDA, end_color=GUINDA)
    for i, t in enumerate(ENCABEZADOS, start=1):
        c = ws.cell(row=1, column=i, value=t)
        c.font = Font(bold=True, color='FFFFFF', size=12); c.fill = relleno
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 28
    for col, ancho in zip('ABC', (20, 14, 30)): ws.column_dimensions[col].width = ancho
    ws.freeze_panes = 'A2'
    ultima = FILAS_PLANTILLA + 1
    for fila in range(2, ultima + 1):
        for col in (1, 2, 3):
            c = ws.cell(row=fila, column=col); c.number_format = '@'; c.alignment = Alignment(horizontal='left')
    dv = DataValidation(type='list', formula1='"' + ','.join(BANCOS_PLANTILLA) + '"', allow_blank=True, showErrorMessage=True, showInputMessage=True,
                        errorStyle='stop', errorTitle='Banco no válido', error='Elija BNB, BMSC, BISA o UNION de la lista.',
                        promptTitle='Banco', prompt='Elija el banco de la lista.')
    dv.add(f'B2:B{ultima}'); ws.add_data_validation(dv)
    ambar = PatternFill('solid', start_color='FFD966', end_color='FFD966', bgColor='FFD966')
    f = _formula_formato()
    for k in ('BNB', 'BMSC', 'NUMERICO'):
        ws.conditional_formatting.add(f'C2:C{ultima}', FormulaRule(formula=[f[k]], fill=ambar, font=Font(color='5E3D00')))

    # --- hoja de instrucciones
    ayuda.column_dimensions['A'].width = 100
    ayuda.sheet_view.showGridLines = False
    for i, (txt, tipo) in enumerate(INSTRUCCIONES, start=1):
        c = ayuda.cell(row=i, column=1, value=txt)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if tipo == 'titulo': c.font = Font(bold=True, size=16, color=GUINDA)
        elif tipo == 'sub': c.font = Font(bold=True, size=12, color='FFFFFF'); c.fill = relleno
        elif tipo == 'ejemplo': c.font = Font(name='Consolas', size=11); c.fill = PatternFill('solid', start_color='F6E7EE', end_color='F6E7EE')
        else: c.font = Font(size=11)
        # altura aproximada para el texto ajustado (Excel no la recalcula al abrir)
        lineas = max(1, -(-len(txt) // 92)); ayuda.row_dimensions[i].height = 22 if tipo in ('titulo', 'sub') else max(18, 16 * lineas + 4)
    for hoja in (ayuda, ws):      # al imprimir, todo el ancho en una pagina
        hoja.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True); hoja.page_setup.fitToWidth = 1; hoja.page_setup.fitToHeight = 0
    wb.active = 0
    wb.save(ruta)
    return ruta


# ---------------------------------------------------------------- lectura
def _norm(s):
    """Minusculas, sin tildes ni espacios ni signos: 'N° factura' -> 'nfactura'."""
    s = unicodedata.normalize('NFD', str(s if s is not None else '')).lower()
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return re.sub(r'[^a-z0-9]', '', s)


_ALIAS = {
    'factura': {'nfactura', 'nrofactura', 'numerofactura', 'nofactura', 'factura', 'nrodefactura', 'numerodefactura'},
    'banco': {'banco'},
    'codigo': {'codigodeasignacion', 'codigoasignacion', 'codigo', 'codasignacion', 'codigodeasignacion'},
}


def texto_celda(v):
    """Convierte el valor de una celda a texto: los numeros enteros sin '.0', espacios recortados."""
    if v is None: return ''
    if isinstance(v, bool): return str(v)
    if isinstance(v, int): return str(v)
    if isinstance(v, float):
        if v != v: return ''
        return str(int(v)) if v.is_integer() else repr(v)
    if hasattr(v, 'isoformat'): return v.isoformat()
    t = str(v).strip()
    return re.sub(r'^(\d+)\.0+$', r'\1', t)


def banco_norm(v):
    """Banco en MAYUSCULAS y sin tildes ('Unión ' -> 'UNION')."""
    t = unicodedata.normalize('NFD', texto_celda(v).upper())
    return ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn').strip()


def leer(ruta):
    """Lee la plantilla. Devuelve una lista de dicts {fila, factura, banco, codigo} (fila = numero de fila de Excel) para las filas con algun dato.
    Lanza ValueError con un mensaje claro si el archivo no es una plantilla legible."""
    try: wb = openpyxl.load_workbook(ruta, data_only=True)
    except Exception: raise ValueError('El archivo no es una plantilla de comunicaciones internas válida (.xlsx). Descargue la plantilla desde «Nuevo cierre».')
    try:
        nombres = {_norm(n): n for n in wb.sheetnames}
        ws = wb[nombres[_norm(HOJA)]] if _norm(HOJA) in nombres else wb[wb.sheetnames[0]]
        filas = []; cab = None; pos = {}
        for n, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if cab is None:
                if n > 30: break
                celdas = [_norm(x) for x in row]
                hallados = {k: next((i for i, c in enumerate(celdas) if c in al), None) for k, al in _ALIAS.items()}
                if all(v is not None for v in hallados.values()): cab = n; pos = hallados
                continue
            def g(k): return row[pos[k]] if pos[k] < len(row) else None
            d = {'fila': n, 'factura': texto_celda(g('factura')), 'banco': banco_norm(g('banco')), 'codigo': texto_celda(g('codigo'))}
            if not (d['factura'] or d['banco'] or d['codigo']): continue
            filas.append(d)
            if len(filas) > MAX_FILAS_LEIDAS: raise ValueError(f'La plantilla tiene más de {MAX_FILAS_LEIDAS} filas con datos; revise que sea el archivo correcto.')
    finally: wb.close()
    if cab is None:
        raise ValueError('La plantilla no tiene los encabezados esperados («N° factura», «Banco», «Código de asignación»). Use la plantilla de «Nuevo cierre» sin cambiar la primera fila.')
    return filas
