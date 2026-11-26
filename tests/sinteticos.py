"""Generador de reportes SINTETICOS con nombres ficticios (para pruebas sin datos reales)."""
import datetime as dt
import openpyxl

FECHA = dt.date(2026, 10, 16)
# (nro_factura, tipo_pago, monto, razon social, estudiante, estado, hora)
FACTURAS = [
    ('30001', 'EFECTIVO', 500, 'LUCIA FERNANDEZ ROCA', 'FERNANDEZ ROCA LUCIA', 'Válido', '10:00:10'),
    ('30002', 'BANCO', 300, 'MARTIN QUISPE LEON', 'QUISPE LEON MARTIN', 'Válido', '10:30:00'),
    ('30003', 'EFECTIVO', 700, 'EMPRESA ALFA SRL', 'SOTO VEGA CARLOS', 'Válido', '11:00:00'),
    ('30004', 'TARJETA', 200, 'RAUL MENDEZ PAZ', 'MENDEZ PAZ RAUL', 'Válido', '11:30:00'),
    ('30005', 'EFECTIVO', 150, 'ELENA ROJAS DIAZ', 'ROJAS DIAZ ELENA', 'Anulado', '12:00:00'),
]
# (id, titular, hora, monto)
BANCO = [
    ('700000001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500),
    ('700000002', 'QUISPE LEON MARTIN', '10:29:40', 300),
    ('700000003', 'SOTO VEGA ANDREA', '10:58:50', 700),
    ('700000004', 'PERSONA AJENA LIMA', '12:00:00', 900),
]


def crear_a1(ruta, facturas=FACTURAS, caja='SFC999', fecha=FECHA):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'RIngresoDiario'
    ws.append([]); ws.append(['UNIVERSIDAD PRIVADA DEL VALLE', 'Usuario:', 'CAJERO1']); ws.append(['SISTEMA DE FACTURACIÓN', 'Fecha:', '16/10/2026 17:00'])
    ws.append(['Caja:', caja]); ws.append([]); ws.append(['INGRESO DIARIO']); ws.append([]); ws.append([])
    cols = ['Nro', 'Fecha', 'Número Factura', 'Nit/C.I.', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago', 'Banco / Cuenta bancaria',
            'Código de Bancarización', 'Nro. Cuenta SAP asociada', 'Monto', 'Cuenta', 'Estado']
    ws.append(cols)
    for i, (nf, tp, m, rs, es, est, h) in enumerate(facturas, start=1):
        hh = dt.datetime.strptime(h, '%H:%M:%S').time()
        ws.append([i, dt.datetime.combine(fecha, hh), nf, '123', rs, es, tp, '-', '-', '-', float(m), 'USR', est])
    wb.save(ruta); return ruta


def crear_a2(ruta, banco=BANCO, quitar=(), fecha=FECHA):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(['ID QR', 'ID TRANSACCION', 'TITULAR ORIGEN', 'CUENTA ORIGEN', 'GLOSA', 'CUENTA DESTINO', 'FECHA / HORA', 'PUNTOS DE VENTA', 'MONTO'])
    tot = 0
    for idt, tit, h, m in banco:
        if idt in quitar: continue
        ws.append(['999', idt, tit, '******0001', 'Ingresos', '3041', f"{fecha:%d/%m/%Y} {h}", 'SEDE X', f'{m:,.2f} BOB']); tot += m
    ws.append([None] * 7 + ['TOTAL', f'{tot:,.2f} BOB'])
    wb.save(ruta); return ruta


# ---------- formato NUEVO del reporte (sin columnas bancarias, 'Valido' sin tilde) y PDF ----------
COLS_NUEVO = ['Nro', 'Fecha', 'Número Factura', 'Nit/C.I.', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago', 'Monto', 'Cuenta', 'Estado']
FACT_NUEVO = [  # (nro factura, tipo, monto, razon social, estudiante, estado, hora)
    ('41001', 'EFECTIVO', 701.00, 'CLARA MONTES VIDAL', 'MONTES VIDAL CLARA SOFIA', 'Valido', '10:42:00'),
    ('41002', 'EFECTIVO', 1220.76, 'COMERCIAL BETA SRL', 'DURAN ROCA PEDRO JOSE LUIS', 'Valido', '10:55:00'),
    ('41003', 'EFECTIVO', 300.00, 'IVAN PEREZ SOTO', 'PEREZ SOTO IVAN', 'Valido', '11:10:00'),
    ('41004', 'EFECTIVO', 1500.50, 'NORA CAMPOS LEON', 'CAMPOS LEON NORA ALEJANDRA MARIA', 'Valido', '11:30:00'),
    ('41005', 'EFECTIVO', 485.90, 'OSCAR RIOS PAZ', 'RIOS PAZ OSCAR', 'Valido', '11:45:00'),
]


def crear_a1_nuevo_xlsx(ruta, facturas=FACT_NUEVO, hoja='Reporte', fecha=FECHA, caja='SFC777'):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = hoja
    ws.append(['UNIVERSIDAD PRIVADA DEL VALLE']); ws.append(['SEDE FICTICIA']); ws.append(['Caja:', caja]); ws.append(['INGRESO DIARIO']); ws.append([])
    ws.append(COLS_NUEVO)
    for i, (nf, tp, m, rs, es, est, h) in enumerate(facturas, start=1):
        ws.append([i, dt.datetime.combine(fecha, dt.datetime.strptime(h, '%H:%M:%S').time()), nf, '1234567', rs, es, tp, float(m), 'USR', est])
    wb.save(ruta); return ruta


def crear_a1_pdf(ruta, facturas=FACT_NUEVO, por_pagina=None, bordes=True, fecha=FECHA, caja="SFC777", pie=True, ajuste_total=0):
    """PDF sintetico que imita el reporte: cabecera, tabla con bordes (celdas de varias lineas, fecha y hora en dos lineas), pie con totales."""
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    ss = getSampleStyleSheet(); c = ss['BodyText']; c.fontSize = 7; c.leading = 9
    doc = SimpleDocTemplate(ruta, pagesize=landscape(letter), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    filas = [(i, *f) for i, f in enumerate(facturas, start=1)]
    paginas = [filas] if not por_pagina else [filas[i:i + por_pagina] for i in range(0, len(filas), por_pagina)]
    vs = [f for f in facturas if f[5] == 'Valido']
    def suma(tp): return sum(f[2] for f in vs if f[1] == tp)
    el = []; total_pag = len(paginas)
    for n, pag in enumerate(paginas, start=1):
        el += [Paragraph('UNIVERSIDAD PRIVADA DEL VALLE', ss['Title']), Paragraph('SISTEMA DE FACTURACIÓN FICTICIO', c), Paragraph('SEDE FICTICIA', c),
               Paragraph('Usuario: CAJERO1', c), Paragraph(f'Fecha: {fecha:%d/%m/%Y} 17:05', c), Paragraph(f'Caja: {caja}', c),
               Paragraph('INGRESO DIARIO', ss['Heading2']), Paragraph(f'Fecha: {fecha:%d/%m/%Y}', c), Spacer(1, 6)]
        data = [[Paragraph(h, c) for h in COLS_NUEVO]]
        for (i, nf, tp, m, rs, es, est, h) in pag:
            f1 = f'{fecha:%d/%m/%Y}\n{h[:5]}'
            data.append([str(i), f1, nf, '1234567', rs, es, tp, f'{m:.2f}', 'USR', est])
        data = [data[0]] + [[Paragraph(str(x).replace('\n', '<br/>'), c) for x in r] for r in data[1:]]
        t = Table(data, colWidths=[32, 52, 66, 52, 100, 110, 52, 50, 40, 40], repeatRows=1)
        t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')] + ([('GRID', (0, 0), (-1, -1), 0.5, colors.black)] if bordes else [])))
        el.append(t)
        if n == total_pag and pie:
            el += [Spacer(1, 8), Paragraph(f'Monto Credito: {0:.2f}', c), Paragraph(f'Monto Deposito: {suma("BANCO"):,.2f}', c),
                   Paragraph(f'Monto Efectivo: {suma("EFECTIVO"):,.2f}', c), Paragraph(f'Monto Tarjeta: {suma("TARJETA"):,.2f}', c),
                   Paragraph(f'Total: {sum(f[2] for f in vs) + ajuste_total:,.2f}', c)]
        el += [Spacer(1, 6), Paragraph(f'Página: {n}/{total_pag}', c)]
        if n < total_pag: el.append(PageBreak())
    doc.build(el); return ruta


# ---------- reporte del banco en el formato NUEVO (y variantes flexibles) ----------
PAGOS_NUEVO = [(f'7100000{i}', es[:22], h, m) for i, (nf, tp, m, rs, es, est, h) in enumerate(FACT_NUEVO, start=1)]   # (id, titular, hora, monto)
ENC_NUEVO = {'idqr': 'ID QR', 'trans': 'TRANSACCION PAGO', 'titular': 'TITULAR', 'glosa': 'GLOSA', 'punto': 'PUNTO VENTA', 'fecha': 'FECHA PAGO',
             'hora': 'HORA PAGO', 'estado': 'ESTADO', 'importe': 'IMPORTE'}


def _monto_txt(m, fmt):
    if fmt == 'coma': return f'{m:.2f}'.replace('.', ',') + ' Bs'
    if fmt == 'miles_punto': return f'{m:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return f'{m:,.2f} BOB'


def crear_a2_nuevo(ruta, pagos=PAGOS_NUEVO, fecha=FECHA, fechas='texto', hoja='Pagos', filas_previas=0, estados=None, total=None, total_en='ESTADO',
                   encabezados=None, orden=None, extra=False, monto_fmt='std', con_total=True, quitar=()):
    """fechas: 'texto' (FECHA PAGO y HORA PAGO como texto) | 'real' (fecha y hora reales) | 'serial' (numeros) | 'junta' (una columna 'FECHA HORA' texto) |
    'iso' (aaaa-mm-dd). estados: {id: estado}. total: importe de la fila TOTAL (por defecto la suma de los pagados). total_en: columna con la palabra TOTAL."""
    enc = {**ENC_NUEVO, **(encabezados or {})}
    orden = orden or ['idqr', 'trans', 'titular', 'glosa', 'punto', 'fecha', 'hora', 'estado', 'importe']
    if fechas == 'junta': orden = [k for k in orden if k != 'hora']
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = hoja
    for i in range(filas_previas): ws.append(['REPORTE DE PAGOS QR FICTICIO'] if i == 0 else [])
    ws.append([enc[k] for k in orden] + (['OBSERVACION EXTRA'] if extra else []))
    suma = 0.0
    for idt, tit, h, m in pagos:
        if idt in quitar: continue
        est = (estados or {}).get(idt, 'Pagado'); hh = dt.datetime.strptime(h, '%H:%M:%S').time()
        if est.lower() == 'pagado': suma += m
        f = {'texto': f'{fecha:%d/%m/%Y}', 'real': fecha, 'serial': (fecha - dt.date(1899, 12, 30)).days, 'junta': f'{fecha:%d/%m/%Y} {h}',
             'iso': f'{fecha:%Y-%m-%d}'}[fechas]
        ho = {'texto': h, 'real': hh, 'serial': (hh.hour * 3600 + hh.minute * 60 + hh.second) / 86400.0, 'iso': h}.get(fechas)
        val = {'idqr': '999', 'trans': idt, 'titular': tit, 'glosa': 'Ingresos', 'punto': 'SEDE X', 'fecha': f, 'hora': ho, 'estado': est,
               'importe': _monto_txt(m, monto_fmt)}
        ws.append([val[k] for k in orden] + (['x'] if extra else []))
    if con_total:
        fila = {k: None for k in orden}
        fila[{'ESTADO': 'estado', 'PUNTO': 'punto'}.get(total_en, 'estado')] = 'TOTAL'; fila['importe'] = _monto_txt(suma if total is None else total, monto_fmt)
        ws.append([fila[k] for k in orden])
    wb.save(ruta); return ruta
