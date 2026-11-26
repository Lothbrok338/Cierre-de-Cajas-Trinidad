"""Reporte QR del banco: formato nuevo (TRANSACCION PAGO / TITULAR / FECHA PAGO + HORA PAGO / ESTADO / IMPORTE) y lectura flexible
(sinonimos, palabras clave, orden libre, columnas extra, deteccion por contenido); movimiento (parse_a1) con encabezados y valores flexibles.
Todos los datos son ficticios."""
import datetime as dt, os, re
import openpyxl, pytest
from tests.conftest import login, subir
from tests import sinteticos as S
from tests.test_app import bid_de, pag, estado_txt, presentar, db
from app import motor, servicio
from app.servicio import ErrorUsuario


def a2(tmp_path, **kw): return motor.parse_a2(S.crear_a2_nuevo(str(tmp_path / 'b.xlsx'), **kw))
def al(df, nivel=None): return [a for a in df.attrs['alertas_lectura'] if nivel is None or a[0] == nivel]
def msg(tmp_path, **kw):
    with pytest.raises(ErrorUsuario) as e: servicio.validar_a2(S.crear_a2_nuevo(str(tmp_path / 'b.xlsx'), **kw))
    return str(e.value)


def test_formato_nuevo_feliz_5_pagos(tmp_path):
    df = a2(tmp_path)
    assert len(df) == 5 and list(df['ID TRANSACCION']) == [p[0] for p in S.PAGOS_NUEVO]
    assert list(df.columns) == ['ID QR', 'ID TRANSACCION', 'TITULAR ORIGEN', 'CUENTA ORIGEN', 'GLOSA', 'CUENTA DESTINO', 'FECHA / HORA', 'PUNTOS DE VENTA', 'MONTO']
    assert (df['CUENTA ORIGEN'] == '').all() and df.loc[0, 'FECHA / HORA'] == dt.datetime(2026, 10, 16, 10, 42, 0)
    assert sum(motor.num(x) for x in df['MONTO']) == pytest.approx(4208.16) and df.attrs['alertas_lectura'] == []
    servicio.validar_a2(S.crear_a2_nuevo(str(tmp_path / 'b.xlsx')))


def test_estado_distinto_de_pagado_se_excluye_con_alerta(tmp_path):
    df = a2(tmp_path, estados={'71000002': 'Anulado', '71000003': 'pendiente', '71000001': 'PAGADO'})
    assert len(df) == 3 and '71000002' not in set(df['ID TRANSACCION'])
    a = al(df); assert [x[0] for x in a] == ['INFO', 'ALERTA']
    assert 'tiene estado Anulado y no se cuenta' in a[0][2] and '71000002' in a[0][2] and '1220.76' in a[0][2]


@pytest.mark.parametrize('donde', ['ESTADO', 'PUNTO'])
def test_fila_total_en_cualquier_columna(tmp_path, donde):
    df = a2(tmp_path, total_en=donde); assert len(df) == 5 and df.attrs['alertas_lectura'] == []


@pytest.mark.parametrize('modo', ['texto', 'real', 'serial', 'junta', 'iso'])
def test_fecha_y_hora_en_todas_las_variantes(tmp_path, modo):
    df = a2(tmp_path, fechas=modo)
    assert list(df['FECHA / HORA']) == [dt.datetime.combine(S.FECHA, dt.datetime.strptime(p[2], '%H:%M:%S').time()) for p in S.PAGOS_NUEVO]
    assert (len(al(df, 'INFO')) == 1) == (modo == 'iso') and not al(df, 'ALERTA')
    motor.preparar_a2(df)


def test_cabecera_no_en_fila_1_y_hoja_con_otro_nombre(tmp_path):
    df = a2(tmp_path, filas_previas=4, hoja='Hoja3'); assert len(df) == 5
    p = S.crear_a2_nuevo(str(tmp_path / 'dos.xlsx')); wb = openpyxl.load_workbook(p)
    wb.create_sheet('Notas', 0).append(['nada aqui']); wb.active = 0; wb.save(p)       # la hoja activa no es la de datos
    assert len(motor.parse_a2(p)) == 5


def test_total_del_archivo_que_no_cuadra_alerta_sin_bloquear(tmp_path):
    df = a2(tmp_path, total=4000.00); assert len(df) == 5
    assert 'no coincide con la fila TOTAL' in al(df, 'ALERTA')[0][2]
    servicio.validar_a2(S.crear_a2_nuevo(str(tmp_path / 'b.xlsx'), total=4000.00))


def test_columnas_faltantes_y_sin_cabecera_dan_mensaje_claro(tmp_path):
    m = msg(tmp_path, orden=['idqr', 'trans', 'glosa', 'punto', 'fecha', 'hora', 'estado', 'importe'])
    assert m.startswith(servicio.MSG_A2) and 'Faltan las columnas: TITULAR ORIGEN. Columnas encontradas:' in m and 'IMPORTE' in m and 'Ingresos' not in m
    p = str(tmp_path / 'x.xlsx'); wb = openpyxl.Workbook(); wb.active.append(['uno', 'dos']); wb.save(p)
    with pytest.raises(ErrorUsuario) as e: servicio.validar_a2(p)
    assert 'No se encontró la fila de cabecera' in str(e.value) and 'Hojas revisadas: Sheet' in str(e.value)


def test_encabezados_con_otras_palabras_orden_libre_y_columna_extra(tmp_path):
    enc = {'trans': 'Nro. de Transacción QR', 'titular': 'Nombre del Titular', 'importe': 'Importe pagado (Bs)', 'fecha': 'Fecha de pago', 'hora': 'Hora de pago',
           'estado': 'Estado del pago', 'punto': 'Punto de Venta'}
    df = a2(tmp_path, encabezados=enc, extra=True, orden=['importe', 'estado', 'hora', 'fecha', 'titular', 'trans', 'punto', 'glosa', 'idqr'])
    assert len(df) == 5 and list(df['ID TRANSACCION']) == [p[0] for p in S.PAGOS_NUEVO] and df.loc[0, 'MONTO'] == '701.00 BOB'
    assert list(df['FECHA / HORA'])[2].hour == 11
    info = ' '.join(x[2] for x in al(df, 'INFO'))
    assert '«Nro. de Transacción QR» se interpretó como ID TRANSACCION' in info and '«Importe pagado (Bs)» se interpretó como MONTO' in info
    assert not al(df, 'ALERTA')


@pytest.mark.parametrize('fmt', ['coma', 'miles_punto'])
def test_montos_con_coma_decimal(tmp_path, fmt):
    df = a2(tmp_path, monto_fmt=fmt); assert [round(motor.num(x), 2) for x in df['MONTO']] == [m for *_, m in S.PAGOS_NUEVO]
    assert any('coma decimal' in x[2] for x in al(df, 'INFO')) and not al(df, 'ALERTA')    # tampoco salta la alerta del TOTAL


def test_num_variantes():
    for t, v in [('1,220.76 BOB', 1220.76), ('1.220,76', 1220.76), ('Bs 1 220,5', 1220.5), ('1220,76', 1220.76), ('1,220', 1220.0), ('1.5', 1.5),
                 ('12,345,678.9', 12345678.9), ('0,500', 0.5), (' 300 ', 300.0), ('-5,25', -5.25)]:
        assert motor.num(t) == pytest.approx(v), t
    with pytest.raises(ValueError): motor.num('abc')


def test_deteccion_por_contenido_con_aviso(tmp_path):
    enc = {'trans': 'Cod', 'titular': 'Persona', 'importe': 'Valor'}
    df = a2(tmp_path, encabezados=enc)
    assert len(df) == 5 and df.loc[1, 'ID TRANSACCION'] == '71000002' and motor.num(df.loc[1, 'MONTO']) == 1220.76
    a = al(df, 'INFO'); assert len(a) == 1 and 'se detectaron las columnas por su contenido' in a[0][2]
    for c in ('ID TRANSACCION', 'TITULAR ORIGEN', 'MONTO'): assert c in a[0][2]
    servicio.validar_a2(S.crear_a2_nuevo(str(tmp_path / 'b.xlsx'), encabezados=enc))


def test_deteccion_por_contenido_ambigua_da_error_claro(tmp_path):
    m = msg(tmp_path, encabezados={'titular': 'Persona A', 'punto': 'Lugar'})
    assert m.startswith(servicio.MSG_A2) and 'No se pudo decidir' in m and 'TITULAR ORIGEN' in m and 'columna' in m and 'Lugar' in m


def test_titular_truncado_22_empareja_media(tmp_path):
    # 'TORREZ ALMANZA NATALI' (21, truncado, orden NOMBRE APELLIDOS) vs razon 'NATALIA TORREZ' y estudiante 'BASCOPE TORREZ HUGO'
    sc, cl = motor.clasif_nombre('TORREZ ALMANZA NATALI', 'NATALIA TORREZ', 'BASCOPE TORREZ HUGO')
    assert cl == 'parcial' and sc == pytest.approx(2 / 3)
    d = tmp_path
    p1 = S.crear_a1(str(d / 'm.xlsx'), [('50001', 'EFECTIVO', 999.90, 'NATALIA TORREZ', 'BASCOPE TORREZ HUGO', 'Válido', '15:43:30')])
    p2 = S.crear_a2_nuevo(str(d / 'b.xlsx'), [('88000001', 'TORREZ ALMANZA NATALI', '15:46:44', 999.90)])
    r = motor.conciliar(p1, p2, 0); f = r['facturas'].iloc[0]
    assert (f['estado_conciliacion'], f['nivel']) == ('POR CONFIRMAR', 'MEDIA')


# ---------- movimiento (parse_a1) flexible ----------
def a1_xlsx(tmp_path, enc, filas, nombre='m.xlsx'):
    p = str(tmp_path / nombre); wb = openpyxl.Workbook(); ws = wb.active; ws.append(enc)
    for f in filas: ws.append(f)
    wb.save(p); return p


def test_movimiento_encabezados_flexibles_orden_libre_y_extra(tmp_path):
    enc = ['Observaciones', 'Situación', 'Forma de Pago', 'Importe', 'Nombre del Estudiante', 'Razón Social', 'N° Factura', 'Fecha Emisión', 'Hora', 'N°']
    p = a1_xlsx(tmp_path, enc, [
        ['x', 'ANULADA', 'Efectivo ', '1.220,76 Bs', 'GOMEZ RUIZ ANA', 'ANA GOMEZ', 20001, '16-10-2026', '10:15:30', 1],
        ['y', 'VÁLIDA', 'EFECTIVO', '300,5', 'PAZ LUNA JOSE', 'JOSE PAZ', '20002', '2026-10-16', dt.time(11, 5), 2]])
    df = motor.parse_a1(p)
    assert list(df['Número Factura']) == ['20001', '20002'] and list(df['Monto']) == [1220.76, 300.5]
    assert list(df['Estado']) == ['Anulado', 'Válido'] and list(df['Tipo Pago']) == ['EFECTIVO', 'EFECTIVO']
    assert list(df['Fecha']) == [dt.datetime(2026, 10, 16, 10, 15, 30), dt.datetime(2026, 10, 16, 11, 5)]
    assert df.loc[0, 'Razon Social'] == 'ANA GOMEZ' and df.loc[1, 'Nombre Estudiante'] == 'PAZ LUNA JOSE' and 'Hora' not in df.columns
    info = ' '.join(a[2] for a in df.attrs['alertas_lectura']); assert all(a[0] == 'INFO' for a in df.attrs['alertas_lectura'])
    assert '«Importe» se interpretó como Monto' in info and '«Situación» se interpretó como Estado' in info and 'coma decimal' in info and 'aaaa-mm-dd' in info
    assert servicio.validar_a1(p)[0] == 'xlsx'


def test_movimiento_monto_gana_a_total_y_sin_nro_se_numera(tmp_path):
    p = a1_xlsx(tmp_path, ['Nro Factura', 'Fecha', 'Tipo de Pago', 'Total', 'Monto', 'Estado'],
                [['30001', dt.datetime(2026, 10, 16, 9, 0), 'Tarjeta', 999, 10, 'Valido'], ['30002', dt.datetime(2026, 10, 16, 9, 5), 'Banco', 999, 20, 'Valido']])
    df = motor.parse_a1(p); assert list(df['Monto']) == [10.0, 20.0] and list(df['Nro']) == [1, 2]
    assert any('no hay columna Nro' in a[2] for a in df.attrs['alertas_lectura'])


def test_movimiento_faltan_columnas_mensaje(tmp_path):
    p = a1_xlsx(tmp_path, ['Nro', 'Fecha', 'Número Factura', 'Monto', 'Estado'], [[1, dt.datetime(2026, 10, 16), '1', 5, 'Valido']])
    with pytest.raises(ErrorUsuario) as e: servicio.validar_a1(p)
    assert 'Faltan las columnas: Tipo Pago' in str(e.value)


# ---------- punta a punta ----------
def test_flujo_web_banco_formato_nuevo_cuadra_y_presenta(app, tmp_path):
    pdf = S.crear_a1_pdf(str(tmp_path / 'ingreso diario.pdf')); banco = S.crear_a2_nuevo(str(tmp_path / 'banco nuevo.xlsx'), fechas='iso', hoja='Pagos QR', filas_previas=2)
    c = login(app, 'ana'); r = subir(c, pdf, banco, '0'); assert r.status_code == 303, r.text
    bid = bid_de(r); html = pag(c, bid).text
    assert estado_txt(html) == 'CUADRA', html[:300]
    assert 'se interpretó como' not in html and 'aaaa-mm-dd' not in html          # las INFO de lectura no se muestran al cajero
    assert presentar(c, bid).status_code == 303
    row = db(app).execute('select qr_registrado from cierres').fetchone()
    assert row[0] == pytest.approx(4208.16)
    assert db(app).execute("select count(*) from alertas_cierre where detalle like '%se interpretó%'").fetchone()[0] == 0


def test_web_banco_ilegible_muestra_causa(app, tmp_path):
    pdf = S.crear_a1_pdf(str(tmp_path / 'ingreso diario.pdf'))
    banco = S.crear_a2_nuevo(str(tmp_path / 'b.xlsx'), orden=['trans', 'glosa', 'punto', 'fecha', 'hora', 'estado', 'importe'])
    c = login(app, 'ana'); r = subir(c, pdf, banco, '0')
    assert r.status_code in (200, 400, 422) and 'Faltan las columnas: TITULAR ORIGEN' in r.text
