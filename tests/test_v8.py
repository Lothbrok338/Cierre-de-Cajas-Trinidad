"""Pruebas v8: plantilla Excel de comunicaciones internas (depositos a otros bancos). Datos ficticios; un caso con reales si hay DATOS_PRUEBA_DIR."""
import glob, hashlib, json, os, re
import openpyxl
import pytest
from tests.conftest import login, nuevo_cliente, post, token
from tests import sinteticos as S
from tests.test_app import bid_de, db, pag, presentar
from tests.test_v2 import flash_de, leer_estado, presentado_con_ci, respaldo
from tests.test_v3 import revisar
from tests.test_v5 import carpeta, texto_pdf
from app import motor, plantilla as PL, servicio

DIR_REAL = os.environ.get('DATOS_PRUEBA_DIR', '')
real = pytest.mark.skipif(not DIR_REAL or not glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls')), reason='DATOS_PRUEBA_DIR sin datos reales')
OBS_LARGA = 'Observacion de prueba larga para rechazar'


def construir(d, extras=0):
    """60001 CONCILIADA (FUERTE, 800000001) · 60002/60003/60005 SIN RESPALDO · 60004 POR CONFIRMAR (800000004) · 60006 TARJETA 200 · 60007 anulada
    (+ `extras` facturas SIN RESPALDO 60008...)."""
    d.mkdir()
    fac = [('60001', 'EFECTIVO', 500, 'LUCIA FERNANDEZ ROCA', 'FERNANDEZ ROCA LUCIA', 'Válido', '10:00:10'),
           ('60002', 'EFECTIVO', 700, 'EMPRESA ALFA SRL', 'SOTO VEGA CARLOS', 'Válido', '11:00:00'),
           ('60003', 'BANCO', 410, 'OSCAR LUNA SOTO', 'LUNA SOTO OSCAR', 'Válido', '12:30:00'),
           ('60004', 'EFECTIVO', 125, 'DIEGO PAZ RIOS', 'PAZ RIOS DIEGO', 'Válido', '12:45:00'),
           ('60005', 'EFECTIVO', 90, 'NORMA VEGA LIMA', 'VEGA LIMA NORMA', 'Válido', '13:00:00'),
           ('60006', 'TARJETA', 200, 'RAUL MENDEZ PAZ', 'MENDEZ PAZ RAUL', 'Válido', '13:30:00'),
           ('60007', 'EFECTIVO', 150, 'ELENA ROJAS DIAZ', 'ROJAS DIAZ ELENA', 'Anulado', '14:00:00')]
    fac += [(str(60008 + i), 'EFECTIVO', 11 + i, f'CLIENTE EXTRA {i}', f'EXTRA CLIENTE {i}', 'Válido', '14:10:00') for i in range(extras)]
    banco = [('800000001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500), ('800000004', 'PAZ RIOS SILVIA', '12:44:30', 125), ('800000099', 'PERSONA AJENA LIMA', '12:00:00', 900)]
    return {'a1': S.crear_a1(str(d / 'mov.xlsx'), fac), 'a2': S.crear_a2(str(d / 'banco.xlsx'), banco), 'dir': d}


@pytest.fixture
def caso(tmp_path): return construir(tmp_path / 'c8')


@pytest.fixture
def caso_g(tmp_path): return construir(tmp_path / 'c8g', 6)


def hacer_plantilla(ruta, filas, cabecera=None, hoja=None):
    """Plantilla real (generada) con las filas dadas [(factura, banco, codigo)] desde la fila 2; los valores se escriben tal cual (numeros incluidos)."""
    PL.generar_plantilla(str(ruta)); wb = openpyxl.load_workbook(ruta); ws = wb[PL.HOJA]
    if cabecera:
        for i, t in enumerate(cabecera, start=1): ws.cell(row=1, column=i, value=t)
    for n, fila in enumerate(filas, start=2):
        for i, v in enumerate(fila, start=1):
            if v is not None: ws.cell(row=n, column=i, value=v)
    if hoja: ws.title = hoja
    wb.save(ruta); return str(ruta)


def subir_pl(c, a1, a2, pl=None, tarjeta='200'):
    tk = token(c, '/nuevo'); fs = {'movimiento': (os.path.basename(a1), open(a1, 'rb')), 'banco': (os.path.basename(a2), open(a2, 'rb'))}
    if pl: fs['plantilla'] = (os.path.basename(pl), open(pl, 'rb'))
    try: return c.post('/nuevo', data={'csrf': tk, 'tarjeta': tarjeta}, files=fs)
    finally:
        for _, f in fs.values(): f.close()


def nuevo(app, caso, filas, usuario='ana', tarjeta='200', **kw):
    c = login(app, usuario); pl = hacer_plantilla(caso['dir'] / f'pl_{usuario}.xlsx', filas, **kw)
    r = subir_pl(c, caso['a1'], caso['a2'], pl, tarjeta); assert r.status_code == 303, r.text
    return c, bid_de(r), pl


def problemas(app, bid): return {p['fila']: p for p in (leer_estado(app, bid)['plantilla']['problemas'])}
def estados(app, bid): return dict(zip(*[app.state.almacen.conciliar(bid, leer_estado(app, bid))['facturas'][k] for k in ('Número Factura', 'estado_conciliacion')]))
def resumen(app, bid): e = leer_estado(app, bid)['plantilla']; return e['leidas'], e['aplicadas'], e['con_problemas']
def sha(ruta): return hashlib.sha256(open(ruta, 'rb').read()).hexdigest()

TRES = [('60002', 'BNB', 'P12345'), ('60003', 'BMSC', 'T98765'), ('60005', 'BISA', '4567890')]


# ---------- plantilla generada ----------
def verificar_plantilla(ruta):
    wb = openpyxl.load_workbook(ruta); assert wb.sheetnames == [PL.HOJA_AYUDA, PL.HOJA]
    ws = wb[PL.HOJA]
    assert [c.value for c in ws[1]][:4] == ['N° factura', 'Banco', 'Código de asignación', 'Fecha del depósito'] and ws.max_column == 4
    assert ws['A1'].font.bold and ws['A1'].font.color.rgb.endswith('FFFFFF') and ws['A1'].fill.start_color.rgb.endswith('870B42')
    assert ws.freeze_panes == 'A2'
    for fila in (2, 100, 201):
        for col in 'ABCD': assert ws[f'{col}{fila}'].number_format == '@' and ws[f'{col}{fila}'].value is None
    assert all(c.value is None for r in ws.iter_rows(min_row=2) for c in r)       # sin filas de ejemplo
    dvs = ws.data_validations.dataValidation; assert len(dvs) == 1
    assert dvs[0].type == 'list' and dvs[0].formula1 == '"BNB,BMSC,BISA,UNION,BANECO"' and 'B2:B201' in str(dvs[0].sqref)
    assert len(ws.conditional_formatting) >= 1
    assert 'cliente' in ' '.join(str(r[0].value) for r in wb[PL.HOJA_AYUDA].iter_rows() if r[0].value).lower()
    assert PL.leer(ruta) == []


def test_plantilla_generada_encabezados_validacion_y_formato_texto(tmp_path):
    verificar_plantilla(PL.generar_plantilla(str(tmp_path / 'p.xlsx')))


def test_plantilla_del_repo_es_valida_y_la_instruccion_cubre_lo_pedido():
    assert os.path.isfile(PL.RUTA_ESTATICA); verificar_plantilla(PL.RUTA_ESTATICA)
    txt = ' '.join(str(r[0].value) for r in openpyxl.load_workbook(PL.RUTA_ESTATICA)[PL.HOJA_AYUDA].iter_rows() if r[0].value)
    for k in ('QR BANECO', 'tarjeta', 'Una factura por fila', '.xlsx', 'importe, el cliente y la fecha', 'BNB', 'BMSC', 'BISA', 'UNION', 'TRES archivos'): assert k in txt


def test_descarga_por_ruta_y_permisos(app):
    assert nuevo_cliente(app).get('/plantilla-comunicaciones-internas').headers['location'] == '/login'
    for u in ('ana', 'ing', 'sup'):
        r = login(app, u).get('/plantilla-comunicaciones-internas'); assert r.status_code == 200
        assert 'attachment' in r.headers['content-disposition'] and 'Plantilla_Comunicaciones_Internas.xlsx' in r.headers['content-disposition']
        assert 'spreadsheetml' in r.headers['content-type'] and r.content[:2] == b'PK'
    open('/tmp/_pl_descarga.xlsx', 'wb').write(r.content); verificar_plantilla('/tmp/_pl_descarga.xlsx')
    assert 'plantilla-comunicaciones-internas' in login(app, 'ana').get('/nuevo').text


# ---------- lector tolerante ----------
def test_lector_tolerante(tmp_path):
    wb = openpyxl.Workbook(); wb.active.title = 'Hoja de ayuda'; ws = wb.create_sheet('  comunicaciones  internas ')
    ws.append(['titulo cualquiera']); ws.append([None, 'NRO. FACTURA', 'BANCO ', 'Codigo de Asignacion'])
    ws.append([None, 20048.0, ' bnb ', 'p123.0']); ws.append([None, None, None, None]); ws.append([None, '20049', 'Unión', 4567.0]); ws.append([None, '   ', None, '  '])
    ws.append([None, 20050, 'bisa', None])
    p = str(tmp_path / 'x.xlsx'); wb.save(p)
    assert PL.leer(p) == [{'fila': 3, 'factura': '20048', 'banco': 'BNB', 'codigo': 'p123.0', 'fecha_deposito': ''}, {'fila': 5, 'factura': '20049', 'banco': 'UNION', 'codigo': '4567', 'fecha_deposito': ''},
                          {'fila': 7, 'factura': '20050', 'banco': 'BISA', 'codigo': '', 'fecha_deposito': ''}]
    wb2 = openpyxl.Workbook(); wb2.active.append(['otra', 'cosa']); p2 = str(tmp_path / 'y.xlsx'); wb2.save(p2)     # primera hoja si no existe la de C.I.; sin encabezados -> error
    with pytest.raises(ValueError, match='encabezados'): PL.leer(p2)
    (tmp_path / 'z.xlsx').write_bytes(b'no es excel')
    with pytest.raises(ValueError): PL.leer(str(tmp_path / 'z.xlsx'))


# ---------- importacion ----------
def test_importacion_feliz_tres_filas(app, caso):
    c, bid, pl = nuevo(app, caso, TRES)
    assert resumen(app, bid) == (3, 3, 0)
    est = leer_estado(app, bid)
    assert est['reclasificaciones'] == {'60002': {'banco': 'BNB', 'codigo': 'P12345'}, '60003': {'banco': 'BMSC', 'codigo': 'T98765'}, '60005': {'banco': 'BISA', 'codigo': '4567890'}}
    assert sorted(est['plantilla_origen']) == ['60002', '60003', '60005'] and est['plantilla_descartes'] == []
    assert sha(os.path.join(app.state.almacen.dir, 'borradores', bid, 'comunicaciones_internas_original.xlsx')) == sha(pl)
    res = app.state.almacen.conciliar(bid, est); t = res['totales']
    assert t['n_COMUNICACIÓN INTERNA'] == 3 and t['C.I. según sistema'] == 1200 and t['QR registrado'] == 625
    assert t['Total cierre'] == t['QR registrado'] + t['C.I. según sistema'] + 200 == 2025
    html = pag(c, bid).text
    assert 'Plantilla:</strong> 3 filas leídas · 3 aplicadas · 0 con problemas' in html
    assert html.count('(desde plantilla)') == 3 and 'Comunicación interna — BNB P12345' in html and 'Filas de la plantilla con problemas' not in html
    assert 'Plantilla leída: 3 filas · 3 aplicadas · 0 con problemas.' in html
    assert est['nombres_subidos']['plantilla'] == 'pl_ana.xlsx'


def test_cada_error_de_fila_con_su_mensaje_y_el_resto_se_aplica(app, caso_g):
    caso = caso_g
    # otro cajero ya presento el codigo BNB P77777 en su cierre vigente (misma fecha, otro usuario)
    presentado_con_ci(app, {'a1': S.crear_a1(str(caso['dir'] / 'b1.xlsx')), 'a2': S.crear_a2(str(caso['dir'] / 'b2.xlsx')), 'dir': caso['dir']}, 'beto', 'P77777')
    filas = [('60002', 'BNB', 'P12345'),      # 2 ok
             ('60003', 'XXX', 'P1111'),       # 3 banco invalido
             ('99999', 'BNB', 'P2222'),       # 4 inexistente
             ('60007', 'BNB', 'P3333'),       # 5 anulada
             ('60006', 'BNB', 'P4444'),       # 6 tarjeta
             ('60005', 'BISA', '111222'),     # 7 factura repetida (con 8)
             ('60005', 'UNION', '333444'),    # 8
             ('60004', 'BMSC', 'T5555'),      # 9 POR CONFIRMAR: se permite
             ('60001', 'BNB', 'P6666'),       # 10 ya conciliada FUERTE
             ('60008', 'BNB', 'P7777'),       # 11 codigo repetido en el banco (con 12)
             ('60009', 'BNB', 'p7777'),       # 12
             (None, 'BNB', 'P8888'),          # 13 factura vacia
             ('60010', None, 'P8888'),        # 14 banco vacio
             ('60011', 'BNB', None),          # 15 codigo vacio
             ('60012', 'BNB', 'P77777'),      # 16 ya usado en otro cierre vigente
             ('60013', 'BNB', 'P1')]          # 17 codigo demasiado corto
    c, bid, _ = nuevo(app, caso, filas)
    p = problemas(app, bid); est = leer_estado(app, bid)
    assert resumen(app, bid) == (16, 2, 14)
    assert 'no es válido' in p[3]['motivo'] and '«XXX»' in p[3]['motivo']
    assert 'no existe en el movimiento' in p[4]['motivo'] and '99999' in p[4]['motivo']
    assert 'ANULADA' in p[5]['motivo']
    assert 'TARJETA' in p[6]['motivo']
    assert 'repetida en la plantilla' in p[7]['motivo'] and 'fila 8' in p[7]['motivo'] and 'fila 7' in p[8]['motivo']
    assert 9 not in p and est['reclasificaciones']['60004'] == {'banco': 'BMSC', 'codigo': 'T5555'}
    assert 'pago QR del banco' in p[10]['motivo'] and '800000001' in p[10]['motivo']
    assert 'repetido en la plantilla' in p[11]['motivo'] and 'fila 12' in p[11]['motivo'] and 'fila 11' in p[12]['motivo']
    assert p[13]['motivo'].startswith('Falta el N° de factura')
    assert 'falta el banco' in p[14]['motivo'].lower()
    assert 'falta el código de asignación' in p[15]['motivo'].lower()
    assert 'ya está registrado en el cierre CC-' in p[16]['motivo'] and 'beto' in p[16]['motivo']
    assert 'entre 3 y 40 caracteres' in p[17]['motivo']
    assert 2 not in p                                                         # la fila 2 si se aplico
    assert est['reclasificaciones'].get('60002') == {'banco': 'BNB', 'codigo': 'P12345'}
    assert set(est['reclasificaciones']) == {'60002', '60004'} and sorted(est['plantilla_origen']) == ['60002', '60004']   # las filas con error NO se aplican
    html = pag(c, bid).text
    assert 'Filas de la plantilla con problemas (14)' in html and 'Fila</th>' in html and 'pago QR del banco' in html
    e = estados(app, bid); assert e['60001'] == 'CONCILIADA' and e['60003'] == 'SIN RESPALDO' and e['60002'] == 'COMUNICACIÓN INTERNA'


def test_ya_confirmada_por_cajero_con_qr_es_conflicto_pero_por_confirmar_se_permite(app, caso):
    c, bid, _ = nuevo(app, caso, [('60004', 'BNB', 'P1000')])
    assert problemas(app, bid) == {} and leer_estado(app, bid)['reclasificaciones'] == {'60004': {'banco': 'BNB', 'codigo': 'P1000'}}
    c2 = login(app, 'ana'); r = subir_pl(c2, caso['a1'], caso['a2'], None); b2 = bid_de(r)        # sin plantilla: confirmo 60004 y subo la plantilla luego
    post(c2, f'/borrador/{b2}/confirmar', f'/borrador/{b2}', factura='60004')
    pl = hacer_plantilla(caso['dir'] / 'tarde.xlsx', [('60004', 'BNB', 'P1000')])
    tk = token(c2, f'/borrador/{b2}')
    r = c2.post(f'/borrador/{b2}/plantilla', data={'csrf': tk}, files={'plantilla': ('tarde.xlsx', open(pl, 'rb'))}); assert r.status_code == 303
    p = problemas(app, b2); assert '800000004' in p[2]['motivo'] and 'pago QR del banco' in p[2]['motivo']


def test_codigo_ya_registrado_en_otra_factura_manual_del_mismo_cierre(app, caso):
    c = login(app, 'ana'); b = bid_de(subir_pl(c, caso['a1'], caso['a2'], None))
    respaldo(c, b, '60003', 'BNB', 'P5000')
    pl = hacer_plantilla(caso['dir'] / 'm.xlsx', [('60002', 'BNB', 'P5000'), ('60003', 'BISA', '123456')])
    c.post(f'/borrador/{b}/plantilla', data={'csrf': token(c, f'/borrador/{b}')}, files={'plantilla': ('m.xlsx', open(pl, 'rb'))})
    p = problemas(app, b); assert 'ya está registrado en la factura 60003' in p[2]['motivo'] and 'reclasificada a mano' in p[3]['motivo']
    assert set(leer_estado(app, b)['reclasificaciones']) == {'60003'}


def test_alertas_de_formato_no_bloquean(app, caso):
    # BNB sin P, BMSC sin T, BISA con letras: se aplican igual y generan alertas; el cierre se puede presentar
    c, bid, _ = nuevo(app, caso, [('60002', 'BNB', 'ZZZ99'), ('60003', 'BMSC', 'ABC12'), ('60005', 'BISA', '12AB34')])
    assert resumen(app, bid) == (3, 3, 0)
    html = pag(c, bid).text; assert html.count('FORMATO A REVISAR') >= 3
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='60004')
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    assert db(app).execute('select count(*) from alertas_cierre').fetchone()[0] == 3


def test_facturas_numericas_guardadas_como_numero_se_normalizan(app, caso):
    c, bid, _ = nuevo(app, caso, [(60002.0, 'bnb', 'p12345'), (60003, ' BISA ', 123456.0), (60005.0, 'UNION', 7890123)])
    est = leer_estado(app, bid); assert resumen(app, bid) == (3, 3, 0)
    assert est['reclasificaciones'] == {'60002': {'banco': 'BNB', 'codigo': 'P12345'}, '60003': {'banco': 'BISA', 'codigo': '123456'}, '60005': {'banco': 'UNION', 'codigo': '7890123'}}


def test_plantilla_vacia_es_valida(app, caso):
    c, bid, _ = nuevo(app, caso, [])
    assert resumen(app, bid) == (0, 0, 0) and leer_estado(app, bid)['reclasificaciones'] == {}
    assert 'Plantilla:</strong> 0 filas leídas · 0 aplicadas · 0 con problemas' in pag(c, bid).text
    # y la plantilla recien descargada, tal cual, tambien
    r = subir_pl(c, caso['a1'], caso['a2'], PL.RUTA_ESTATICA); assert r.status_code == 303


def test_sin_plantilla_el_flujo_es_el_de_siempre(app, caso):
    c = login(app, 'ana'); r = subir_pl(c, caso['a1'], caso['a2']); bid = bid_de(r); est = leer_estado(app, bid)
    assert 'plantilla' not in est and 'Plantilla:' not in pag(c, bid).text and 'Subir plantilla de comunicaciones internas' in pag(c, bid).text
    respaldo(c, bid, '60002', 'BNB', 'P12345'); respaldo(c, bid, '60003', 'BNB', 'P12346'); respaldo(c, bid, '60005', 'BNB', 'P12347')
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='60004')
    r = presentar(c, bid); cid = int(r.headers['location'].rsplit('/', 1)[1]); d = carpeta(app, cid)
    assert not os.path.exists(os.path.join(d, 'comunicaciones_internas_original.xlsx'))
    cj = json.load(open(os.path.join(d, 'cierre.json'), encoding='utf-8'))
    assert 'plantilla' not in cj and 'sha256_plantilla' not in cj and db(app).execute('select sha256_plantilla from cierres').fetchone()[0] is None
    assert 'Plantilla de comunicaciones' not in texto_pdf(os.path.join(d, 'cierre.pdf')) and 'Plantilla de comunicaciones' not in c.get(f'/cierres/{cid}').text


@pytest.mark.parametrize('contenido,nombre,msg', [(b'esto no es excel', 'x.xlsx', 'xlsx válido'), (None, 'sin_encabezados.xlsx', 'encabezados')])
def test_plantilla_ilegible_rechaza_la_subida(app, caso, contenido, nombre, msg):
    p = caso['dir'] / nombre
    if contenido: p.write_bytes(contenido)
    else:
        wb = openpyxl.Workbook(); wb.active.append(['a', 'b']); wb.save(p)
    c = login(app, 'ana'); r = subir_pl(c, caso['a1'], caso['a2'], str(p))
    assert r.status_code == 400 and msg in r.text and not os.listdir(os.path.join(app.state.almacen.dir, 'borradores'))
    p2 = caso['dir'] / 'p.csv'; p2.write_text('a,b'); r = subir_pl(c, caso['a1'], caso['a2'], str(p2)); assert r.status_code == 400 and '.xlsx' in r.text


# ---------- reemplazar / deshacer ----------
def test_reemplazar_plantilla_rehace_la_importacion(app, caso):
    c, bid, _ = nuevo(app, caso, [('60002', 'BNB', 'P12345'), ('99999', 'BNB', 'P22222')])
    assert resumen(app, bid) == (2, 1, 1) and '99999' in pag(c, bid).text
    # el cajero resuelve a mano 60003 y luego reemplaza la plantilla corregida
    respaldo(c, bid, '60003', 'BISA', '555666')
    pl2 = hacer_plantilla(caso['dir'] / 'v2.xlsx', [('60005', 'UNION', '9876543')])
    r = c.post(f'/borrador/{bid}/plantilla', data={'csrf': token(c, f'/borrador/{bid}')}, files={'plantilla': ('Corregida.xlsx', open(pl2, 'rb'))}); assert r.status_code == 303
    est = leer_estado(app, bid)
    assert set(est['reclasificaciones']) == {'60003', '60005'} and est['plantilla_origen'] == ['60005']        # la de la plantilla vieja se fue; la manual se conserva
    assert resumen(app, bid) == (1, 1, 0) and est['nombres_subidos']['plantilla'] == 'Corregida.xlsx'
    assert sha(os.path.join(app.state.almacen.dir, 'borradores', bid, 'comunicaciones_internas_original.xlsx')) == sha(pl2)
    assert 'Filas de la plantilla con problemas' not in pag(c, bid).text
    # un archivo malo al reemplazar no toca nada
    mala = caso['dir'] / 'mala.xlsx'; mala.write_bytes(b'xx')
    c.post(f'/borrador/{bid}/plantilla', data={'csrf': token(c, f'/borrador/{bid}')}, files={'plantilla': ('mala.xlsx', open(mala, 'rb'))})
    assert 'xlsx válido' in pag(c, bid).text and set(leer_estado(app, bid)['reclasificaciones']) == {'60003', '60005'}


def test_deshacer_reclasificacion_de_plantilla_no_se_reaplica(app, caso):
    c, bid, _ = nuevo(app, caso, TRES)
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='60003')
    est = leer_estado(app, bid)
    assert '60003' not in est['reclasificaciones'] and est['plantilla_descartes'] == ['60003'] and sorted(est['plantilla_origen']) == ['60002', '60005']
    assert estados(app, bid)['60003'] == 'SIN RESPALDO'
    # recalcular (reaplicar la plantilla guardada): la deshecha no vuelve
    nuevo_est, res = app.state.almacen.aplicar_plantilla(bid, est, 'ana'); assert '60003' not in nuevo_est['reclasificaciones'] and res['descartadas'] == 1 and res['aplicadas'] == 2
    html = pag(c, bid).text; assert '1 deshecha por usted' in html and html.count('(desde plantilla)') == 2
    # la puede resolver a mano de otro modo; ya no cuenta como "de plantilla"
    respaldo(c, bid, '60003', 'BNB', 'P99999'); est = leer_estado(app, bid)
    assert est['reclasificaciones']['60003'] == {'banco': 'BNB', 'codigo': 'P99999'} and '60003' not in est['plantilla_origen']
    # reemplazar la plantilla reinicia los descartes
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='60003')
    c.post(f'/borrador/{bid}/plantilla', data={'csrf': token(c, f'/borrador/{bid}')}, files={'plantilla': ('otra.xlsx', open(hacer_plantilla(caso['dir'] / 'o.xlsx', TRES), 'rb'))})
    assert leer_estado(app, bid)['plantilla_descartes'] == [] and '60003' in leer_estado(app, bid)['reclasificaciones']


# ---------- presentar / detalle / PDF ----------
def presentado(app, caso, filas=TRES):
    c, bid, pl = nuevo(app, caso, filas)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='60004')
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    return c, bid, pl, int(r.headers['location'].rsplit('/', 1)[1])


def test_presentar_guarda_archivo_sha_detalle_y_pdf(app, caso):
    c, bid, pl, cid = presentado(app, caso); d = carpeta(app, cid); h = sha(pl)
    assert sha(os.path.join(d, 'comunicaciones_internas_original.xlsx')) == h
    cj = json.load(open(os.path.join(d, 'cierre.json'), encoding='utf-8'))
    assert cj['sha256_plantilla'] == h and cj['sha256_archivos']['comunicaciones_internas_original.xlsx'] == h
    assert cj['plantilla']['leidas'] == 3 and cj['plantilla']['aplicadas'] == 3 and sorted(cj['plantilla_origen']) == ['60002', '60003', '60005']
    assert db(app).execute('select sha256_plantilla from cierres where id=?', (cid,)).fetchone()[0] == h
    assert db(app).execute('select count(*) from cierre_ci').fetchone()[0] == 3
    for u in ('ana', 'ing'):
        html = login(app, u).get(f'/cierres/{cid}').text
        assert h in html and 'comunicaciones_internas_original.xlsx' in html and html.count('(desde plantilla)') == 3
    r = login(app, 'ana').get(f'/cierres/{cid}/archivo/comunicaciones_internas_original.xlsx'); assert r.status_code == 200 and r.content == open(pl, 'rb').read()
    for pdf in ('cierre.pdf', 'cierre_revision.pdf'):
        if pdf == 'cierre_revision.pdf': revisar(login(app, 'ing'), cid, 'aprobar')
        t = texto_pdf(os.path.join(d, pdf))
        assert f'Plantilla de comunicaciones internas: {h[:16]}…{h[-8:]}' in t and 'Cargadas desde la plantilla' in t and '60002' in t


def test_otro_cierre_vigente_bloquea_el_mismo_codigo_en_la_plantilla_siguiente(app, caso):
    presentado(app, caso)      # ana presento BNB P12345 / BMSC T98765 / BISA 4567890
    d2 = caso['dir'] / 'otra'; d2.mkdir()
    a1 = S.crear_a1(str(d2 / 'm.xlsx'), fecha=__import__('datetime').date(2026, 10, 17))
    a2 = S.crear_a2(str(d2 / 'b.xlsx'), fecha=__import__('datetime').date(2026, 10, 17))
    c = login(app, 'beto'); pl = hacer_plantilla(d2 / 'p.xlsx', [('30003', 'BNB', 'p12345')])
    b = bid_de(subir_pl(c, a1, a2, pl)); p = problemas(app, b)
    assert 'ya está registrado en el cierre CC-' in p[2]['motivo'] and 'ana' in p[2]['motivo']


# ---------- editar y reenviar ----------
def test_editar_y_reenviar_copia_y_permite_reemplazar_la_plantilla(app, caso):
    c, bid, pl, cid = presentado(app, caso); h = sha(pl)
    revisar(login(app, 'ing'), cid, 'rechazar', OBS_LARGA)
    r = post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}'); assert r.status_code == 303
    b2 = bid_de(r); d2 = os.path.join(app.state.almacen.dir, 'borradores', b2); est = leer_estado(app, b2)
    assert sha(os.path.join(d2, 'comunicaciones_internas_original.xlsx')) == h
    assert sorted(est['plantilla_origen']) == ['60002', '60003', '60005'] and est['plantilla']['aplicadas'] == 3 and set(est['reclasificaciones']) == {'60002', '60003', '60005'}
    html = pag(c, b2).text; assert 'Plantilla:</strong> 3 filas leídas' in html and 'nuevo-plantilla' in html and html.count('(desde plantilla)') == 3
    # reemplazo por /archivos (solo la plantilla)
    nueva = hacer_plantilla(caso['dir'] / 'n.xlsx', [('60002', 'BNB', 'P12345'), ('60003', 'BMSC', 'T11111'), ('60005', 'BISA', '4567890')])
    r = c.post(f'/borrador/{b2}/archivos', data={'csrf': token(c, f'/borrador/{b2}')}, files={'plantilla': ('n.xlsx', open(nueva, 'rb'))}); assert r.status_code == 303
    est = leer_estado(app, b2); assert est['reclasificaciones']['60003']['codigo'] == 'T11111' and resumen(app, b2) == (3, 3, 0)
    assert sha(os.path.join(d2, 'comunicaciones_internas_original.xlsx')) == sha(nueva) != h        # el codigo de su propio cierre rechazado no cuenta como conflicto
    r = presentar(c, b2); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, b2)
    cid2 = int(r.headers['location'].rsplit('/', 1)[1]); d = carpeta(app, cid2)
    assert sha(os.path.join(d, 'comunicaciones_internas_original.xlsx')) == sha(nueva) and db(app).execute('select sha256_plantilla from cierres where id=?', (cid2,)).fetchone()[0] == sha(nueva)


def test_editar_reenviar_sin_plantilla_previa_permite_subirla(app, sint):
    from tests.test_v2 import presentado_con_ci
    c, bid, cid = presentado_con_ci(app, sint)
    revisar(login(app, 'ing'), cid, 'rechazar', OBS_LARGA)
    b2 = bid_de(post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}')); assert 'plantilla' not in leer_estado(app, b2)
    pl = hacer_plantilla(sint['dir'] / 'q.xlsx', [('30003', 'BNB', 'DEP-12345')])      # el mismo codigo de su cierre rechazado: no es conflicto
    c.post(f'/borrador/{b2}/archivos', data={'csrf': token(c, f'/borrador/{b2}')}, files={'plantilla': ('q.xlsx', open(pl, 'rb'))})
    est = leer_estado(app, b2); assert est['plantilla']['con_problemas'] == 1 and 'reclasificada a mano' in est['plantilla']['problemas'][0]['motivo']


# ---------- permisos ----------
def test_permisos_de_la_plantilla_en_el_borrador(app, caso):
    c, bid, pl = nuevo(app, caso, TRES)
    anon = nuevo_cliente(app)
    assert anon.post(f'/borrador/{bid}/plantilla', files={'plantilla': ('x.xlsx', open(pl, 'rb'))}).status_code == 303
    otro = login(app, 'beto')
    r = otro.post(f'/borrador/{bid}/plantilla', data={'csrf': token(otro, '/nuevo')}, files={'plantilla': ('x.xlsx', open(pl, 'rb'))}); assert r.status_code == 404
    r = c.post(f'/borrador/{bid}/plantilla', data={'csrf': 'malo'}, files={'plantilla': ('x.xlsx', open(pl, 'rb'))}); assert r.status_code == 403
    ing = login(app, 'ing'); assert ing.get('/nuevo').status_code == 403
    assert ing.post(f'/borrador/{bid}/plantilla', data={'csrf': token(ing, '/cuenta')}, files={'plantilla': ('x.xlsx', open(pl, 'rb'))}).status_code == 404
    assert otro.get(f'/borrador/{bid}').status_code == 404


# ---------- migracion ----------
def test_migracion_v5_idempotente(tmp_path):
    import sqlite3
    from app import db as dbm
    ruta = str(tmp_path / 'm.db'); dbm.inicializar(ruta)
    con = sqlite3.connect(ruta); con.execute('ALTER TABLE cierres DROP COLUMN sha256_plantilla'); con.execute('PRAGMA user_version = 4'); con.commit(); con.close()
    for _ in range(2):
        dbm.inicializar(ruta); con = sqlite3.connect(ruta)
        assert 'sha256_plantilla' in {r[1] for r in con.execute('PRAGMA table_info(cierres)')} and con.execute('PRAGMA user_version').fetchone()[0] == 6 == dbm.VERSION_ESQUEMA; con.close()


# ---------- datos reales ----------
@real
def test_real_plantilla_con_facturas_del_movimiento_real(app, tmp_path):
    a1 = glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls'))[0]; a2 = glob.glob(os.path.join(DIR_REAL, '*Reporte_pagos*.xlsx'))[0]
    f = motor.conciliar(a1, a2, 0)['facturas']
    obs = list(f[f['estado_conciliacion'].isin(['SIN RESPALDO', 'AMBIGUO', 'POR CONFIRMAR'])]['Número Factura'])[:2]
    fuerte = list(f[f['estado_conciliacion'] == 'CONCILIADA']['Número Factura'])[:1]
    if len(obs) < 1 or not fuerte: pytest.skip('el movimiento real no tiene el caso necesario')
    filas = [(obs[0], 'BNB', 'P40001'), (fuerte[0], 'BNB', 'P40002'), ('00000', 'BISA', '123456')]
    pl = hacer_plantilla(tmp_path / 'real.xlsx', filas); c = login(app, 'ana'); r = subir_pl(c, a1, a2, pl, '0'); assert r.status_code == 303, r.text
    p = problemas(app, bid_de(r)); assert resumen(app, bid_de(r)) == (3, 1, 2) and 'pago QR del banco' in p[3]['motivo'] and 'no existe' in p[4]['motivo']
