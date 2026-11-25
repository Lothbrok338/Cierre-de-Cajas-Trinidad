"""Movimiento del cajero en el formato NUEVO del reporte 'Ingreso Diario' (sin columnas bancarias, 'Valido' sin tilde,
hoja con otro nombre) y en PDF. Todos los datos son ficticios."""
import datetime as dt, glob, json, os
import pandas as pd, pytest, openpyxl
from tests.conftest import login, post, subir
from tests import sinteticos as S
from tests.test_app import bid_de, pag, estado_txt, presentar, db, sha
from app import motor, servicio
from app.servicio import ErrorUsuario

DIR_REAL = os.environ.get('DATOS_PRUEBA_DIR', '')
COLS_OBLIG = ['Nro', 'Fecha', 'Número Factura', 'Tipo Pago', 'Monto', 'Estado']
MULTI = S.FACT_NUEVO + [('41006', 'TARJETA', 250.00, 'ELSA VEGA RUIZ', 'VEGA RUIZ ELSA', 'Valido', '12:00:00'),
                        ('41007', 'EFECTIVO', 99.00, 'JOSE LUNA MAR', 'LUNA MAR JOSE', 'Anulado', '12:10:00')]


def esperado(facturas):
    """DataFrame esperado (tal como lo produce el lector de Excel) para las facturas dadas."""
    f = pd.DataFrame([{'Nro': i, 'Fecha': dt.datetime.combine(S.FECHA, dt.datetime.strptime(h, '%H:%M:%S').time()), 'Número Factura': nf,
                       'Razon Social': rs, 'Nombre Estudiante': es, 'Tipo Pago': tp, 'Monto': float(m),
                       'Estado': 'Válido' if est == 'Valido' else est} for i, (nf, tp, m, rs, es, est, h) in enumerate(facturas, start=1)])
    return f


def comparar(df, facturas):
    e = esperado(facturas)
    cols = list(e.columns)
    pd.testing.assert_frame_equal(df[cols].reset_index(drop=True), e, check_dtype=False)
    assert df['Nro'].dtype.kind == 'i' and df['Monto'].dtype.kind == 'f' and str(df['Fecha'].dtype).startswith('datetime64')
    for c in ('Banco / Cuenta bancaria', 'Código de Bancarización', 'Nro. Cuenta SAP asociada'): assert (df[c] == '').all()
    assert list(df['_fila_excel']) == list(range(1, len(df) + 1)) or df['_fila_excel'].min() >= 1


# ---------- Excel formato nuevo ----------
def test_xlsx_formato_nuevo_hoja_otro_nombre(tmp_path):
    p = S.crear_a1_nuevo_xlsx(str(tmp_path / 'n.xlsx'), hoja='Reporte')
    df = motor.parse_a1(p); comparar(df, S.FACT_NUEVO)
    assert motor.metadatos_a1(p)['caja'] == 'SFC777'
    assert servicio.validar_a1(p)[0] == 'xlsx'


def test_xls_formato_nuevo_otro_nombre_de_hoja(tmp_path):
    xlwt = pytest.importorskip('xlwt')
    wb = xlwt.Workbook(); ws = wb.add_sheet('RIngresoDiarioCS'); fmt = xlwt.easyxf(num_format_str='DD/MM/YYYY HH:MM')
    ws.write(0, 0, 'Caja:'); ws.write(0, 1, 'SFC777')
    for c, h in enumerate(S.COLS_NUEVO): ws.write(2, c, h)
    for i, (nf, tp, m, rs, es, est, h) in enumerate(S.FACT_NUEVO, start=1):
        r = 2 + i
        ws.write(r, 0, i); ws.write(r, 1, dt.datetime.combine(S.FECHA, dt.datetime.strptime(h, '%H:%M:%S').time()), fmt); ws.write(r, 2, nf)
        ws.write(r, 4, rs); ws.write(r, 5, es); ws.write(r, 6, tp); ws.write(r, 7, float(m)); ws.write(r, 9, est)
    p = str(tmp_path / 'n.xls'); wb.save(p)
    comparar(motor.parse_a1(p), S.FACT_NUEVO); assert servicio.validar_a1(p)[0] == 'xls'


def test_encabezados_sin_tildes_ni_mayusculas(tmp_path):
    p = str(tmp_path / 'h.xlsx'); wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'X'
    ws.append(['NRO', 'FECHA', 'Numero\nFactura', 'TIPO  PAGO', 'monto', 'ESTADO'])
    ws.append([1, dt.datetime(2026, 10, 16, 9, 0), 5001, ' banco ', 10, 'VALIDO'])
    wb.save(p); df = motor.parse_a1(p)
    assert list(df['Número Factura']) == ['5001'] and list(df['Tipo Pago']) == ['BANCO'] and list(df['Estado']) == ['Válido']
    assert df.loc[0, 'Razon Social'] == '' and df.loc[0, 'Nombre Estudiante'] == ''


def test_estado_y_tipo_pago_normalizados_y_desconocidos_alertan():
    a1 = esperado([('1', 'EFECTIVO', 10, 'A B', 'B A', 'Valido', '10:00:00'), ('2', ' tarjeta ', 5, 'A B', 'B A', 'VÁLIDO', '10:01:00'),
                   ('3', 'EFECTIVO', 7, 'A B', 'B A', 'ANULADO', '10:02:00'), ('4', 'EFECTIVO', 8, 'A B', 'B A', 'Pendiente', '10:03:00'),
                   ('5', 'CHEQUE', 9, 'A B', 'B A', 'Valido', '10:04:00')])
    a1['Estado'] = ['Valido', 'VÁLIDO', 'ANULADO', 'Pendiente', 'Valido']
    f, al = motor.preparar_a1(a1)
    assert list(f['estado_prep']) == ['VALIDA', 'VALIDA', 'EXCLUIDA (Anulado)', 'NO PROCESADA - valor no reconocido', 'NO PROCESADA - valor no reconocido']
    assert list(f['tipo'][:2]) == ['QR', 'TARJETA']
    assert sum(1 for a in al if a[0] == 'ALERTA') == 2


def test_sin_nombres_nivel_debil_con_alerta(tmp_path):
    p = str(tmp_path / 's.xlsx'); wb = openpyxl.Workbook(); ws = wb.active
    ws.append(['Nro', 'Fecha', 'Número Factura', 'Tipo Pago', 'Monto', 'Estado'])
    ws.append([1, dt.datetime(2026, 10, 16, 10, 0), '9001', 'EFECTIVO', 300.0, 'Valido']); wb.save(p)
    a2 = S.crear_a2(str(tmp_path / 'b.xlsx'), [('700000001', 'PERSONA UNO DOS', '10:00:30', 300)])
    r = motor.conciliar(p, a2, 0)
    assert any('no trae nombres' in d for d in r['alertas']['detalle'])
    assert 'DEBIL' in r['facturas'].loc[0, 'candidatos']


# ---------- PDF ----------
def test_pdf_sintetico_igual_al_esperado(tmp_path):
    p = S.crear_a1_pdf(str(tmp_path / 'r.pdf'))
    df = motor.parse_a1(p); comparar(df, S.FACT_NUEVO)
    assert df.attrs['alertas_lectura'] == []
    m = motor.metadatos_a1(p); assert m['caja'] == 'SFC777' and m['usuario'] == 'CAJERO1' and m['fecha_emision'] == '16/10/2026 17:05'
    assert servicio.validar_a1(p)[0] == 'pdf'
    x = motor.parse_a1(S.crear_a1_nuevo_xlsx(str(tmp_path / 'x.xlsx')))
    pd.testing.assert_frame_equal(df[list(x.columns)].drop(columns='_fila_excel'), x.drop(columns='_fila_excel'), check_dtype=False)   # _fila_excel: fila del detalle vs fila de la hoja


def test_pdf_multipagina_cabecera_repetida_tarjeta_y_anulada(tmp_path):
    p = S.crear_a1_pdf(str(tmp_path / 'm.pdf'), MULTI, por_pagina=3)
    import pdfplumber
    assert len(pdfplumber.open(p).pages) >= 2
    df = motor.parse_a1(p); comparar(df, MULTI)
    assert df.attrs['alertas_lectura'] == [] and list(df['Estado'])[-1] == 'Anulado' and 'TARJETA' in set(df['Tipo Pago'])


def test_pdf_sin_bordes_por_posicion(tmp_path):
    comparar(motor.parse_a1(S.crear_a1_pdf(str(tmp_path / 'sb.pdf'), bordes=False, por_pagina=3)), S.FACT_NUEVO)


def test_pdf_total_no_coincide_alerta_sin_bloqueo(tmp_path):
    p = S.crear_a1_pdf(str(tmp_path / 't.pdf'), ajuste_total=100)
    df = motor.parse_a1(p); comparar(df, S.FACT_NUEVO)
    assert len(df.attrs['alertas_lectura']) == 1 and 'no coincide con el Total' in df.attrs['alertas_lectura'][0][2]
    assert servicio.validar_a1(p)[0] == 'pdf'


# ---------- mensajes de error ----------
def msg(path):
    with pytest.raises(ErrorUsuario) as e: servicio.validar_a1(path)
    return str(e.value)


def test_mensajes_de_error(tmp_path):
    p = str(tmp_path / 'sm.xlsx'); wb = openpyxl.Workbook(); wb.active.title = 'Hoja1'; wb.active.append(['Nro', 'Fecha']); wb.save(p)
    m = msg(p); assert m.startswith(servicio.MSG_A1) and "'Nro' y 'Monto'" in m and 'Hoja1' in m
    p = str(tmp_path / 'v.xlsx'); wb = openpyxl.Workbook(); wb.active.title = 'Vacia'; wb.save(p)
    m = msg(p); assert m.startswith(servicio.MSG_A1) and 'Vacia' in m
    p = str(tmp_path / 'fc.xlsx'); wb = openpyxl.Workbook(); wb.active.append(['Nro', 'Fecha', 'Número Factura', 'Tipo Pago', 'Monto']); wb.active.append([1, 'x', 'y', 'z', 1]); wb.save(p)
    m = msg(p); assert 'Faltan las columnas: Estado' in m and 'Columnas encontradas: Nro, Fecha' in m
    from reportlab.pdfgen import canvas
    p = str(tmp_path / 'img.pdf'); c = canvas.Canvas(p); c.rect(50, 50, 200, 100); c.save()
    m = msg(p); assert m.startswith(servicio.MSG_A1) and 'El PDF no tiene texto legible; exporte el reporte desde el sistema de facturación (Excel o PDF).' in m
    p = str(tmp_path / 'otro.pdf'); c = canvas.Canvas(p); c.drawString(50, 700, 'Hola mundo'); c.save()
    assert "'Nro' y 'Monto'" in msg(p)


# ---------- web de punta a punta ----------
def banco_nuevo(d):
    return S.crear_a2(str(d / 'banco_n.xlsx'), [(f'7100000{i}', es, h[:8], int(m) if m == int(m) else m)
                                                for i, (nf, tp, m, rs, es, est, h) in enumerate(S.FACT_NUEVO, start=1)])


def test_flujo_web_pdf_nuevo_formato(app, tmp_path):
    pdf = S.crear_a1_pdf(str(tmp_path / 'ingreso diario.pdf')); a2 = banco_nuevo(tmp_path)
    c = login(app, 'ana'); r = subir(c, pdf, a2, '0'); assert r.status_code == 303, r.text
    bid = bid_de(r); html = pag(c, bid).text
    assert estado_txt(html) == 'CUADRA', html[:300]
    assert os.path.isfile(os.path.join(app.state.almacen.dir, 'borradores', bid, 'movimiento_original.pdf'))
    r = presentar(c, bid); assert r.status_code == 303
    row = db(app).execute('select ruta, sha256_movimiento, caja, qr_registrado from cierres').fetchone()
    dest = os.path.join(app.state.almacen.dir, row[0])
    assert row[2] == 'SFC777' and row[3] == pytest.approx(4208.16)
    assert sha(os.path.join(dest, 'movimiento_original.pdf')) == row[1] == sha(pdf)
    cj = json.load(open(os.path.join(dest, 'cierre.json'), encoding='utf-8')); assert cj['archivo_movimiento'] == 'movimiento_original.pdf'
    assert open(os.path.join(dest, 'cierre.pdf'), 'rb').read(5) == b'%PDF-'
    cid = db(app).execute('select id from cierres').fetchone()[0]
    d = c.get(f'/cierres/{cid}/archivo/movimiento_original.pdf'); assert d.status_code == 200 and d.headers['content-type'] == 'application/pdf'
    assert d.content == open(pdf, 'rb').read()


def test_web_reemplazo_de_movimiento_por_pdf_en_editar_y_reenviar(app, sint, tmp_path):
    from tests.test_v2 import rechazado
    from tests.conftest import token
    c, ci, cid = rechazado(app, sint)
    bid = bid_de(post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}'))
    pdf = S.crear_a1_pdf(str(tmp_path / 'r.pdf'))
    with open(pdf, 'rb') as f:
        r = c.post(f'/borrador/{bid}/archivos', data={'csrf': token(c, f'/borrador/{bid}')}, files={'movimiento': ('r.pdf', f)})
    assert r.status_code == 303, r.text
    d = os.path.join(app.state.almacen.dir, 'borradores', bid)
    assert sha(os.path.join(d, 'movimiento_original.pdf')) == sha(pdf) and not os.path.exists(os.path.join(d, 'movimiento_original.xlsx'))
    falso = tmp_path / 'falso.pdf'; falso.write_bytes(b'esto no es un pdf')
    with open(falso, 'rb') as f:
        r = c.post(f'/borrador/{bid}/archivos', data={'csrf': token(c, f'/borrador/{bid}')}, files={'movimiento': ('falso.pdf', f)})
    assert 'no parece ser el reporte de Ingreso Diario' in c.get(f'/borrador/{bid}').text
    assert sha(os.path.join(d, 'movimiento_original.pdf')) == sha(pdf)     # el inválido no reemplaza nada


@pytest.mark.skipif(not DIR_REAL or not glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls')), reason='DATOS_PRUEBA_DIR sin datos reales')
def test_real_formato_antiguo_y_equivalente_nuevo(tmp_path):
    import xlrd
    a1 = glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls'))[0]; a2 = glob.glob(os.path.join(DIR_REAL, '*Reporte_pagos*.xlsx'))[0]
    viejo = motor.parse_a1(a1); r1 = motor.conciliar(a1, a2, 1191)
    # copia temporal con las columnas del formato nuevo: sin las 3 bancarias, 'Valido' sin tilde, otra hoja
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Reporte'; ws.append([]); ws.append(['Caja:', motor.metadatos_a1(a1)['caja']]); ws.append(S.COLS_NUEVO)
    for _, r in viejo.iterrows():
        ws.append([int(r['Nro']), r['Fecha'].to_pydatetime(), r['Número Factura'], r['Nit/C.I.'], r['Razon Social'], r['Nombre Estudiante'], r['Tipo Pago'], r['Monto'], r['Cuenta'],
                   'Valido' if r['Estado'] == 'Válido' else r['Estado']])
    nuevo = str(tmp_path / 'nuevo.xlsx'); wb.save(nuevo)
    r2 = motor.conciliar(nuevo, a2, 1191)
    quita = lambda d: d[d['concepto'] != 'Generado el'].reset_index(drop=True)
    pd.testing.assert_frame_equal(quita(r1['resumen']), quita(r2['resumen']))
    pd.testing.assert_frame_equal(r1['facturas'].drop(columns=['_fila_excel'], errors='ignore'), r2['facturas'].drop(columns=['_fila_excel'], errors='ignore'))
    pd.testing.assert_frame_equal(r1['banco'], r2['banco'])
    assert r1['totales'] == r2['totales'] if isinstance(r1['totales'], dict) else True
    res = dict(zip(r2['resumen']['concepto'], r2['resumen']['valor']))
    assert (res['n_CONCILIADA'], res['n_POR CONFIRMAR'], res['QR registrado']) == (13, 1, 10784.0)
    print(r1['totales'], r1['resumen'])
