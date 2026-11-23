"""Pruebas v2: pestañas, respaldo manual / C.I., revisión por Ingresos, editar y reenviar, roles y migración."""
import json, os, re, sqlite3
import pytest
from tests.conftest import login, nuevo_cliente, token, post, subir
from tests import sinteticos as S
from tests.test_app import bid_de, db, pag, presentar, estado_txt, sha
from app import catalogo, motor


def respaldo(c, bid, nf, banco, codigo):
    return post(c, f'/borrador/{bid}/respaldo', f'/borrador/{bid}', factura=nf, banco=banco, codigo=codigo)


def flash_de(c, bid):
    m = re.search(r'role="status">(.*?)</div>', pag(c, bid).text, re.S)
    return m.group(1).strip() if m else ''


def leer_estado(app, bid):
    return json.load(open(os.path.join(app.state.almacen.dir, 'borradores', bid, 'estado.json'), encoding='utf-8'))


def borrador_nuevo(app, sint, usuario='ana', a2=None, tarjeta='200'):
    c = login(app, usuario); r = subir(c, sint['a1'], a2 or sint['a2'], tarjeta); assert r.status_code == 303, r.text
    return c, bid_de(r)


def presentado_con_ci(app, sint, usuario='ana', codigo='DEP-12345'):
    """Cierre presentado donde la factura 30003 (700) se reclasifica como C.I. BNB."""
    c, bid = borrador_nuevo(app, sint, usuario)
    respaldo(c, bid, '30003', 'BNB', codigo)
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    return c, bid, int(r.headers['location'].rsplit('/', 1)[1])


@pytest.fixture
def sint_extra(tmp_path):
    """Sintético con una factura QR extra (30006, 300) sin movimiento del banco: dos observadas."""
    d = tmp_path / 'sx'; d.mkdir()
    fac = S.FACTURAS + [('30006', 'EFECTIVO', 300, 'OSCAR LUNA SOTO', 'LUNA SOTO OSCAR', 'Válido', '12:30:00')]
    return {'a1': S.crear_a1(str(d / 'mov.xlsx'), fac), 'a2': S.crear_a2(str(d / 'banco.xlsx'))}


# ---------- catálogo ----------
def test_catalogo_bancos_y_formato():
    assert catalogo.LISTA_BANCOS == ['QR BANECO', 'BNB', 'BMSC', 'BISA', 'UNION', 'BANECO']
    assert (catalogo.BANCOS['BNB']['cuenta_bancaria'], catalogo.BANCOS['BNB']['cuenta_sap']) == ('8000006425', '110103222')
    assert (catalogo.BANCOS['BMSC']['cuenta_bancaria'], catalogo.BANCOS['BMSC']['cuenta_sap']) == ('4010884831', '110103242')
    assert (catalogo.BANCOS['BISA']['cuenta_bancaria'], catalogo.BANCOS['BISA']['cuenta_sap']) == ('696870021', '110103232')
    assert (catalogo.BANCOS['UNION']['cuenta_bancaria'], catalogo.BANCOS['UNION']['cuenta_sap']) == ('1-56089490', '110103382')
    assert catalogo.BANCOS['QR BANECO']['cuenta_bancaria'] is None and catalogo.BANCOS['QR BANECO']['cuenta_sap'] is None
    assert all(b['formato_asignacion'] and b['ayuda'] for b in catalogo.BANCOS.values())
    # el formato por banco es solo alerta: validar_codigo solo bloquea vacío, longitud 3-40 y caracteres de control
    assert catalogo.validar_codigo('BNB', ' dep-12/3.4 ') == (True, None, 'DEP-12/3.4')
    assert catalogo.validar_codigo('BNB', 'ñandú1')[0] and catalogo.validar_codigo('BNB', 'a b c')[0]
    for malo in ('', '   ', 'ab', 'x' * 41, 'AB\nC12'):
        assert not catalogo.validar_codigo('BNB', malo)[0], malo
    assert not catalogo.validar_codigo('NOEXISTE', 'ABC123')[0] and not catalogo.validar_codigo('QR BANECO', 'ABC123')[0]


def test_catalogo_formato_por_banco_alerta(monkeypatch):
    assert catalogo.formato_incumplido('BNB', 'ab12P') is None and catalogo.formato_incumplido('BNB', 'ABC-123') == 'BNB: letras y números, debe incluir una P'
    monkeypatch.setitem(catalogo.BANCOS['BISA'], 'formato_asignacion', r'[0-9]{8}')
    assert catalogo.formato_incumplido('BISA', '12345678') is None and catalogo.formato_incumplido('BISA', '123')
    assert catalogo.validar_codigo('BISA', 'ABC-123')[0]    # nunca bloquea


# ---------- motor: reclasificación ----------
def test_motor_reclasificacion_ci(sint):
    base = motor.conciliar(sint['a1'], sint['a2'], 200)['totales']
    assert base['QR registrado'] == 1500 and base['n_COMUNICACIÓN INTERNA'] == 0 and base['Total cierre'] == 1700
    r = motor.conciliar(sint['a1'], sint['a2'], 200, reclasificaciones={'30003': {'banco': 'BNB', 'codigo': 'DEP-1'}})
    t = r['totales']
    assert t['QR registrado'] == 800 and t['n_facturas_qr'] == 2 and t['Estado QR'] == 'CUADRA' and t['Diferencia QR'] == 0
    assert t['C.I. según sistema'] == 700 and t['C.I. con código de asignación'] == 700 and t['n_COMUNICACIÓN INTERNA'] == 1
    assert t['Total cierre'] == t['QR registrado'] + t['C.I. según sistema'] + t['Tarjeta declarada'] == 1700
    f = r['facturas'].set_index('Número Factura')
    assert f.loc['30003', 'estado_conciliacion'] == 'COMUNICACIÓN INTERNA' and f.loc['30003', 'Tipo (motor)'] == 'C.I.' and f.loc['30003', 'ID TRANSACCION'] == ''
    assert r['ci'] == [{'factura': '30003', 'razon_social': 'EMPRESA ALFA SRL', 'estudiante': 'SOTO VEGA CARLOS', 'monto': 700.0, 'banco': 'BNB', 'codigo': 'DEP-1', 'fecha_deposito': ''}]
    # el movimiento que antes estaba sugerido para 30003 queda libre (ajeno), ya no asignado
    assert (r['banco'].set_index('ID TRANSACCION').loc['700000003', 'estado_banco']) == 'AJENO / NO CORRESPONDE AL CAJERO'
    # reclasificar algo que no es QR válido no se aplica
    r2 = motor.conciliar(sint['a1'], sint['a2'], 200, reclasificaciones={'30004': {'banco': 'BNB', 'codigo': 'X123'}, '99999': {'banco': 'BNB', 'codigo': 'X123'}})
    assert r2['totales']['n_COMUNICACIÓN INTERNA'] == 0 and sum('Reclasificación no aplicada' in a for a in r2['alertas']['detalle']) == 2
    from app import servicio
    assert [c['concepto'] for c in servicio.conceptos(t)] == ['QR BANECO', 'Comunicaciones internas', 'Tarjetas']
    ci = servicio.conceptos(t)[1]
    assert (ci['segun_sistema'], ci['segun_respaldo'], ci['diferencia'], ci['estado']) == (700, 700, 0, 'PREVALIDADO POR INGRESOS') and 'BNB DEP-1' in ci['nota']
    tot = servicio.total_conceptos(servicio.conceptos(t))
    assert tot['segun_sistema'] == 1700 and tot['estado'] == 'CUADRA'


# ---------- pestañas ----------
def test_pestanas_y_conteos(app, sint):
    c, bid = borrador_nuevo(app, sint)
    html = pag(c, bid).text
    assert 'role="tablist"' in html and html.count('role="tab"') == 2 and html.count('role="tabpanel"') == 2
    assert 'Conciliadas (2)' in html and 'Observadas (1)' in html and 'data-tab-inicial="observadas"' in html
    assert re.search(r'id="tab-observadas"[^>]*aria-selected="true"', html) and 'aria-controls="panel-observadas"' in html
    assert 'aria-labelledby="tab-conciliadas"' in html
    assert 'Respaldo manual' in html and 'name="banco"' in html and html.count('<option value="') >= 5
    for b in catalogo.LISTA_BANCOS[1:]: assert f'<option value="{b}">' in html
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    html = pag(c, bid).text
    assert 'Conciliadas (3)' in html and 'Observadas (0)' in html and 'data-tab-inicial="conciliadas"' in html
    assert 'Confirmada por cajero' in html and 'Deshacer' in html
    # deshacer desde Conciliadas devuelve la factura a Observadas
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30003')
    html = pag(c, bid).text
    assert 'Conciliadas (2)' in html and 'Observadas (1)' in html and 'data-tab-inicial="observadas"' in html
    assert '/static/app.js' in html     # sin JS los dos paneles quedan visibles (los hidden los pone el script)
    assert 'hidden' not in re.search(r'<section class="revision".*?</section>', html, re.S).group(0).split('<tbody>')[0].replace('type="button"', '')


def test_pestana_por_defecto_conciliadas_si_no_hay_observadas(app, sint):
    c, bid = borrador_nuevo(app, sint)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    assert 'data-tab-inicial="conciliadas"' in pag(c, bid).text


# ---------- respaldo manual QR BANECO ----------
@pytest.fixture
def sint_manual(tmp_path):
    """30002 sin su movimiento + un movimiento libre (700000009, 300) con otro titular y uno de otra fecha."""
    d = tmp_path / 'sm'; d.mkdir()
    banco = [b for b in S.BANCO if b[0] != '700000002'] + [('700000009', 'ZZZ XXX WWW', '10:29:00', 300)]
    import datetime as dt
    a2 = S.crear_a2(str(d / 'b.xlsx'), banco)
    return {'a1': S.crear_a1(str(d / 'm.xlsx')), 'a2': a2}


def test_respaldo_manual_qr_valido(app, sint_manual):
    c, bid = borrador_nuevo(app, sint_manual)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    assert estado_txt(pag(c, bid).text) == 'NO CUADRA' and 'Observadas (1)' in pag(c, bid).text
    r = respaldo(c, bid, '30002', 'QR BANECO', '700000009'); assert r.status_code == 303
    html = pag(c, bid).text
    assert estado_txt(html) == 'CUADRA' and 'Conciliadas (3)' in html and 'Observadas (0)' in html
    assert 'Respaldo manual</span>' in html and '700000009' in html
    est = leer_estado(app, bid)
    assert est['elecciones'] == {'30002': '700000009'} and est['respaldos_manuales'] == ['30002']
    res = motor.conciliar(os.path.join(app.state.almacen.dir, 'borradores', bid, 'movimiento_original.xlsx'),
                          os.path.join(app.state.almacen.dir, 'borradores', bid, 'banco_qr_original.xlsx'), 200, est['confirmaciones'], est['elecciones'],
                          respaldos_manuales=est['respaldos_manuales'])
    f = res['facturas'].set_index('Número Factura')
    assert f.loc['30002', 'estado_conciliacion'] == 'CONFIRMADA POR CAJERO' and f.loc['30002', 'nivel'] == 'RESPALDO MANUAL'
    ant = [json.loads(a[0]) for a in db(app).execute("select detalle_json from auditoria where accion='respaldo_manual'")]
    assert ant and ant[0]['factura'] == '30002' and ant[0]['banco'] == 'QR BANECO' and ant[0]['codigo'] == '700000009' and 'anterior' in ant[0]
    assert presentar(c, bid).status_code == 303 and db(app).execute('select count(*) from cierres').fetchone()[0] == 1


def test_respaldo_manual_qr_invalido(app, sint_manual):
    c, bid = borrador_nuevo(app, sint_manual)
    casos = [('NOEXISTE1', 'inexistente'), ('700000004', 'monto distinto'), ('', 'Escriba el número de transacción')]
    for codigo, esperado in casos:
        r = respaldo(c, bid, '30002', 'QR BANECO', codigo); assert r.status_code == 303
        assert esperado in flash_de(c, bid), (codigo, flash_de(c, bid))
        assert leer_estado(app, bid)['elecciones'] == {}
    # POR CONFIRMAR: QR BANECO no se ofrece ni se acepta (debe usar «Sí, corresponde»)
    r = respaldo(c, bid, '30003', 'QR BANECO', '700000003'); assert 'Sí, corresponde' in flash_de(c, bid)
    assert leer_estado(app, bid)['elecciones'] == {}
    # factura ya conciliada o inexistente
    respaldo(c, bid, '30001', 'BNB', 'ABC-1'); assert 'no está en Observadas' in flash_de(c, bid)
    respaldo(c, bid, '30002', '', 'ABC-1'); assert 'Seleccione el banco' in flash_de(c, bid)
    assert db(app).execute("select count(*) from auditoria where accion='respaldo_manual_rechazado'").fetchone()[0] >= 6


def test_respaldo_manual_qr_id_ya_usado_por_otra_factura(app, sint_extra):
    c, bid = borrador_nuevo(app, sint_extra)
    respaldo(c, bid, '30006', 'QR BANECO', '700000002')      # asignado automáticamente a 30002 (mismo monto 300)
    assert 'ya fue usado por la factura 30002' in flash_de(c, bid) and leer_estado(app, bid)['elecciones'] == {}


def test_respaldo_manual_qr_id_en_otro_cierre(app, sint_manual):
    ca, ba = borrador_nuevo(app, sint_manual, 'ana')
    respaldo(ca, ba, '30002', 'QR BANECO', '700000009'); post(ca, f'/borrador/{ba}/confirmar', f'/borrador/{ba}', factura='30003')
    assert presentar(ca, ba).status_code == 303
    cb, bb = borrador_nuevo(app, sint_manual, 'beto')
    respaldo(cb, bb, '30002', 'QR BANECO', '700000009')
    assert 'ya está en el cierre' in flash_de(cb, bb) and leer_estado(app, bb)['elecciones'] == {}


# ---------- reclasificación a C.I. por la UI ----------
def test_reclasificar_a_ci(app, sint):
    c, bid = borrador_nuevo(app, sint)
    assert estado_txt(pag(c, bid).text) == 'PENDIENTE DE CONFIRMACIÓN'
    r = respaldo(c, bid, '30003', 'BNB', ' dep-12345 '); assert r.status_code == 303
    html = pag(c, bid).text
    assert estado_txt(html) == 'CUADRA' and 'Conciliadas (3)' in html and 'Observadas (0)' in html
    assert 'Comunicación interna — BNB DEP-12345' in html and 'Comunicaciones internas' in html and 'prevalidado por Ingresos' in html
    assert leer_estado(app, bid)['reclasificaciones'] == {'30003': {'banco': 'BNB', 'codigo': 'DEP-12345'}}
    # deshacer devuelve la factura a Observadas
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30003')
    assert leer_estado(app, bid)['reclasificaciones'] == {} and 'Observadas (1)' in pag(c, bid).text
    # editar = deshacer y volver a registrar; la auditoría guarda el valor anterior
    respaldo(c, bid, '30003', 'BNB', 'DEP-12345'); post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30003'); respaldo(c, bid, '30003', 'BISA', 'XYZ-777')
    filas = [json.loads(a[0]) for a in db(app).execute("select detalle_json from auditoria where accion in ('respaldo_manual','deshacer_confirmacion') order by id")]
    assert filas[-2]['anterior']['reclasificacion'] == {'banco': 'BNB', 'codigo': 'DEP-12345'}
    assert filas[-1]['banco'] == 'BISA' and filas[-1]['resultado'] == 'COMUNICACIÓN INTERNA'


def test_ci_codigo_invalido_o_repetido_en_borrador(app, sint_extra):
    c, bid = borrador_nuevo(app, sint_extra, tarjeta='200')
    for malo in ('', 'ab', 'x' * 41):
        respaldo(c, bid, '30003', 'BNB', malo)
        assert leer_estado(app, bid)['reclasificaciones'] == {}, malo
    respaldo(c, bid, '30003', 'BNB', 'ABC-1'); assert 'reclasificaciones' in leer_estado(app, bid) and len(leer_estado(app, bid)['reclasificaciones']) == 1
    respaldo(c, bid, '30006', 'BNB', 'abc-1')     # mismo par (banco, código), sin distinguir mayúsculas
    assert 'ya está registrado en la factura 30003' in flash_de(c, bid) and list(leer_estado(app, bid)['reclasificaciones']) == ['30003']
    respaldo(c, bid, '30006', 'BMSC', 'ABC-1')    # otro banco con el mismo código: permitido
    assert len(leer_estado(app, bid)['reclasificaciones']) == 2


def test_ci_no_presentable_si_codigo_invalido(app, sint):
    """Defensa en servidor: aunque el estado se manipule, evaluar() exige código válido en cada C.I."""
    from app import servicio
    res = motor.conciliar(sint['a1'], sint['a2'], 200, reclasificaciones={'30003': {'banco': 'BNB', 'codigo': 'x'}})
    ev = servicio.evaluar(res)
    assert not ev['presentable'] and 'Comunicación interna de la factura 30003' in ' '.join(ev['razones'])
    ok = motor.conciliar(sint['a1'], sint['a2'], 200, reclasificaciones={'30003': {'banco': 'BNB', 'codigo': 'DEP-9'}})
    assert servicio.evaluar(ok)['presentable']


def test_ci_presentado_totales_y_pdf(app, sint):
    c, bid, cid = presentado_con_ci(app, sint)
    row = db(app).execute('select qr_registrado, qr_conciliado, tarjeta_declarada, ci_total, total_cierre, estado_revision, ruta from cierres').fetchone()
    assert row[:6] == (800, 800, 200, 700, 1700, 'PENDIENTE DE REVISIÓN')
    assert db(app).execute('select banco, codigo, factura, monto from cierre_ci').fetchall() == [('BNB', 'DEP-12345', '30003', 700.0)]
    assert db(app).execute("select count(*) from cierre_movimientos where factura='30003'").fetchone()[0] == 0     # no cuenta como QR
    dest = os.path.join(app.state.almacen.dir, row[6])
    cj = json.load(open(os.path.join(dest, 'cierre.json'), encoding='utf-8'))
    assert [x['concepto'] for x in cj['conceptos']] == ['QR BANECO', 'Comunicaciones internas', 'Tarjetas']
    assert cj['reclasificaciones'] == {'30003': {'banco': 'BNB', 'codigo': 'DEP-12345'}} and cj['comunicaciones_internas'][0]['codigo'] == 'DEP-12345'
    html = c.get(f'/cierres/{cid}').text
    assert 'Comunicaciones internas' in html and 'DEP-12345' in html and 'PREVALIDADO POR INGRESOS' in html
    pypdf = pytest.importorskip('pypdf')
    txt = ' '.join(p.extract_text() for p in pypdf.PdfReader(os.path.join(dest, 'cierre.pdf')).pages).replace('\n', ' ')
    assert 'Comunicaciones internas' in txt and 'PREVALIDADO POR INGRESOS' in txt and 'DEP-12345' in txt and 'Comprobante de cierre de caja' in txt
    assert 'PENDIENTE DE REVISIÓN' in txt


def test_ci_duplicado_entre_cierres_bloqueado(app, sint):
    presentado_con_ci(app, sint, 'ana', 'DEP-12345')
    cb, bb = borrador_nuevo(app, sint, 'beto')
    respaldo(cb, bb, '30003', 'BNB', 'dep-12345')
    assert 'ya está registrado en el cierre' in flash_de(cb, bb) and leer_estado(app, bb)['reclasificaciones'] == {}
    respaldo(cb, bb, '30003', 'BISA', 'DEP-12345')     # otro banco: es otro par
    assert len(leer_estado(app, bb)['reclasificaciones']) == 1
    # forzado: aunque el estado traiga el par repetido, la vista y la presentación lo bloquean
    est = leer_estado(app, bb); est['reclasificaciones']['30003'] = {'banco': 'BNB', 'codigo': 'DEP-12345'}
    app.state.almacen.guardar_estado(bb, est)
    assert 'data-abrir="dlg-presentar"' not in pag(cb, bb).text and 'ya está en el cierre' in pag(cb, bb).text
    presentar(cb, bb)
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 1


# ---------- tarjeta editable ----------
def test_corregir_tarjeta_en_borrador(app, sint):
    c, bid = borrador_nuevo(app, sint, tarjeta='150')
    respaldo(c, bid, '30003', 'BNB', 'DEP-1')
    html = pag(c, bid).text
    assert estado_txt(html) == 'NO CUADRA' and 'Las tarjetas no cuadran' in html and 'Corregir el total declarado con tarjetas' in html
    post(c, f'/borrador/{bid}/tarjeta', f'/borrador/{bid}', tarjeta='abc'); assert leer_estado(app, bid)['tarjeta'] == 150
    post(c, f'/borrador/{bid}/tarjeta', f'/borrador/{bid}', tarjeta='200'); assert leer_estado(app, bid)['tarjeta'] == 200
    assert estado_txt(pag(c, bid).text) == 'CUADRA'
    a = [json.loads(x[0]) for x in db(app).execute("select detalle_json from auditoria where accion='tarjeta_corregida'")]
    assert a == [{'borrador': bid, 'anterior': 150, 'nuevo': 200}]


# ---------- revisión por Ingresos ----------
def test_ingresos_aprueba(app, sint):
    c, bid, cid = presentado_con_ci(app, sint)
    ci = login(app, 'ing')
    html = ci.get(f'/cierres/{cid}').text
    assert 'Aprobar (Conforme)' in html and 'Rechazar' in html
    r = post(ci, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='aprobar', observacion=''); assert r.status_code == 303
    row = db(app).execute('select estado_revision, revisado_por, revisado_en, observacion_revision, ruta from cierres').fetchone()
    assert row[0] == 'CONFORME' and row[1] == 'ing' and row[2] and row[3] is None
    assert 'Aprobar (Conforme)' not in ci.get(f'/cierres/{cid}').text
    # un CONFORME no se toca: ni rechazar, ni aprobar de nuevo, ni editar por el cajero
    post(ci, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='rechazar', observacion='Intento tardío de rechazo')
    assert db(app).execute('select estado_revision from cierres').fetchone()[0] == 'CONFORME'
    post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}')
    assert db(app).execute('select estado_revision from cierres').fetchone()[0] == 'CONFORME' and len(os.listdir(os.path.join(app.state.almacen.dir, 'borradores'))) == 1
    acc = [a[0] for a in db(app).execute('select accion from auditoria')]
    assert 'cierre_aprobado' in acc and 'editar_bloqueado' in acc
    # PDF con estado de revisión (el cierre.pdf original no se modifica)
    dest = os.path.join(app.state.almacen.dir, row[4]); assert os.path.isfile(os.path.join(dest, 'cierre_revision.pdf'))
    pypdf = pytest.importorskip('pypdf')
    txt = ' '.join(p.extract_text() for p in pypdf.PdfReader(os.path.join(dest, 'cierre_revision.pdf')).pages)
    assert 'CONFORME' in txt and 'Revisado por' in txt and 'ing' in txt
    assert 'cierre_revision.pdf' in c.get(f'/cierres/{cid}').text and c.get(f'/cierres/{cid}/archivo/cierre_revision.pdf').content[:5] == b'%PDF-'
    # el cajero ve el estado en su lista
    assert 'CONFORME' in c.get('/').text


def test_ingresos_rechaza_observacion_obligatoria(app, sint):
    c, bid, cid = presentado_con_ci(app, sint)
    ci = login(app, 'ing')
    for corta in ('', '   ', 'muy corta', '         x          '):
        post(ci, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='rechazar', observacion=corta)
        assert db(app).execute('select estado_revision from cierres').fetchone()[0] == 'PENDIENTE DE REVISIÓN', corta
    assert 'al menos 10 caracteres' in ci.get(f'/cierres/{cid}').text
    post(ci, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='rechazar', observacion='El código del depósito no coincide con el comprobante.')
    row = db(app).execute('select estado_revision, revisado_por, revisado_en, observacion_revision, ruta from cierres').fetchone()
    assert row[0] == 'RECHAZADO' and row[1] == 'ing' and row[2] and row[3] == 'El código del depósito no coincide con el comprobante.'
    d = [json.loads(a[0]) for a in db(app).execute("select detalle_json from auditoria where accion='cierre_rechazado'")]
    assert d and d[0]['observacion'] == row[3]
    # el cajero ve la observación destacada y el botón Editar y reenviar
    html = c.get(f'/cierres/{cid}').text
    assert 'Rechazado por Ingresos' in html and 'El código del depósito no coincide' in html and 'Editar y reenviar' in html and 'class="rechazo"' in html
    assert 'Observación de Ingresos' in c.get('/').text
    pypdf = pytest.importorskip('pypdf')
    txt = ' '.join(p.extract_text() for p in pypdf.PdfReader(os.path.join(app.state.almacen.dir, row[4], 'cierre_revision.pdf')).pages)
    assert 'RECHAZADO' in txt and 'El código del depósito no coincide' in txt.replace('\n', ' ')


def test_ingresos_filtros_y_vista(app, sint, sint_manual):
    c, bid, cid = presentado_con_ci(app, sint)
    ci = login(app, 'ing')
    assert 'Ana Prueba' in ci.get('/').text
    assert 'CC-' in ci.get('/?estado=PENDIENTE DE REVISIÓN').text
    assert 'No hay cierres con esos filtros' in ci.get('/?estado=CONFORME').text
    assert 'CC-' in ci.get('/?desde=2026-10-16&hasta=2026-10-16').text and 'No hay cierres con esos filtros' in ci.get('/?desde=2026-10-17').text
    assert ci.get('/?estado=XX&desde=basura').status_code == 200      # filtros inválidos se ignoran
    for e in ('PENDIENTE DE REVISIÓN', 'CONFORME', 'RECHAZADO', 'REEMPLAZADO'): assert f'<option value="{e}"' in ci.get('/').text
    assert 'name="estado"' not in c.get('/').text    # el cajero no tiene filtros ni ve cierres ajenos


# ---------- roles ----------
def test_roles(app, sint):
    c, bid, cid = presentado_con_ci(app, sint)
    ci = login(app, 'ing')
    # cajero no puede aprobar ni rechazar
    for dec, obs in (('aprobar', ''), ('rechazar', 'Observación suficientemente larga')):
        assert post(c, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision=dec, observacion=obs).status_code == 403
    assert db(app).execute('select estado_revision from cierres').fetchone()[0] == 'PENDIENTE DE REVISIÓN'
    html = c.get(f'/cierres/{cid}').text
    assert 'Aprobar (Conforme)' not in html and 'dlg-rechazar' not in html
    # ingresos no puede crear cierres ni operar borradores ni editar
    assert ci.get('/nuevo').status_code == 403
    assert post(ci, '/nuevo', '/', tarjeta='1').status_code == 403
    assert post(ci, f'/cierres/{cid}/editar', f'/cierres/{cid}').status_code == 403
    assert ci.get(f'/borrador/{bid}').status_code == 404 and respaldo(ci, bid, '30003', 'BNB', 'ABC-1').status_code == 404
    # otro cajero no ve ni revisa
    cb = login(app, 'beto')
    assert cb.get(f'/cierres/{cid}').status_code == 404
    assert post(cb, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='aprobar').status_code == 403
    # sin sesión
    anon = nuevo_cliente(app)
    assert anon.post(f'/cierres/{cid}/revisar', data={'decision': 'aprobar'}).status_code == 303
    # CSRF
    assert ci.post(f'/cierres/{cid}/revisar', data={'decision': 'aprobar', 'csrf': 'mal'}).status_code == 403


def test_cajero_no_edita_pendiente(app, sint):
    c, bid, cid = presentado_con_ci(app, sint)
    n0 = len(os.listdir(os.path.join(app.state.almacen.dir, 'borradores')))
    r = post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}'); assert r.status_code == 303 and r.headers['location'] == f'/cierres/{cid}'
    assert 'Solo se puede editar un cierre RECHAZADO' in c.get(f'/cierres/{cid}').text
    assert len(os.listdir(os.path.join(app.state.almacen.dir, 'borradores'))) == n0
    assert 'Editar y reenviar' not in c.get(f'/cierres/{cid}').text


# ---------- editar y reenviar ----------
OBS = 'El código del depósito no coincide con el comprobante.'


def rechazado(app, sint):
    c, bid, cid = presentado_con_ci(app, sint)
    ci = login(app, 'ing')
    post(ci, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='rechazar', observacion=OBS)
    return c, ci, cid


def test_editar_y_reenviar_crea_v2(app, sint):
    c, ci, cid = rechazado(app, sint)
    # un cierre nuevo para la misma fecha se bloquea y dirige a «Editar y reenviar»
    r = subir(c, sint['a1'], sint['a2']); assert r.status_code == 400 and 'Editar y reenviar' in r.text
    r = post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}'); assert r.status_code == 303 and r.headers['location'].startswith('/borrador/')
    bid2 = bid_de(r)
    est = leer_estado(app, bid2)
    assert est['reclasificaciones'] == {'30003': {'banco': 'BNB', 'codigo': 'DEP-12345'}} and est['tarjeta'] == 200 and est['origen']['cierre_id'] == cid
    d = os.path.join(app.state.almacen.dir, 'borradores', bid2)
    assert sha(os.path.join(d, 'movimiento_original.xlsx')) == sha(sint['a1']) and sha(os.path.join(d, 'banco_qr_original.xlsx')) == sha(sint['a2'])
    html = pag(c, bid2).text
    assert OBS in html and 'Corrigiendo el cierre' in html and estado_txt(html) == 'CUADRA' and 'Conciliadas (3)' in html
    # volver a pulsar «Editar y reenviar» reutiliza el mismo borrador
    assert post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}').headers['location'] == f'/borrador/{bid2}'
    # corregir la tarjeta (bloquea) y volver a dejarla bien
    post(c, f'/borrador/{bid2}/tarjeta', f'/borrador/{bid2}', tarjeta='999')
    assert estado_txt(pag(c, bid2).text) == 'NO CUADRA' and presentar(c, bid2).status_code == 303
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 1
    post(c, f'/borrador/{bid2}/tarjeta', f'/borrador/{bid2}', tarjeta='200')
    # cambiar el código de la C.I.: editar = deshacer y volver a registrar (el par anterior era de v1, que será reemplazado)
    r = presentar(c, bid2); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/')
    cid2 = int(r.headers['location'].rsplit('/', 1)[1])
    rows = db(app).execute('select id, correlativo, version, estado_revision, reemplaza_a, motivo_reemplazo, ruta from cierres order by id').fetchall()
    assert [(x[2], x[3]) for x in rows] == [(1, 'REEMPLAZADO'), (2, 'PENDIENTE DE REVISIÓN')]
    assert rows[1][4] == cid and rows[1][5] == f'Corrección de observación de Ingresos: {OBS}' and os.path.basename(rows[1][6]) == 'v2'
    base = app.state.almacen.dir
    assert all(os.path.isfile(os.path.join(base, x[6], 'cierre.pdf')) for x in rows)         # nada se borra
    cj = json.load(open(os.path.join(base, rows[1][6], 'cierre.json'), encoding='utf-8'))
    assert cj['reclasificaciones'] == {'30003': {'banco': 'BNB', 'codigo': 'DEP-12345'}} and cj['motivo_reemplazo'] == rows[1][5]
    # el par (BNB, DEP-12345) de v1 no bloquea a v2 y solo queda vigente el de v2
    assert db(app).execute("select count(*) from cierre_ci").fetchone()[0] == 2
    assert 'REEMPLAZADO' in c.get(f'/cierres/{cid}').text and f'/cierres/{cid2}' in c.get(f'/cierres/{cid}').text
    # el REEMPLAZADO no se puede editar y el v2 PENDIENTE tampoco
    for x in (cid, cid2):
        n = len(os.listdir(os.path.join(base, 'borradores'))); post(c, f'/cierres/{x}/editar', f'/cierres/{x}')
        assert len(os.listdir(os.path.join(base, 'borradores'))) == n
    ac = [a[0] for a in db(app).execute('select accion from auditoria')]
    assert 'cierre_rechazado' in ac and 'editar_reenviar_iniciado' in ac and ac.count('presentar_cierre') == 2
    # ingresos revisa la v2
    assert 'Aprobar (Conforme)' in ci.get(f'/cierres/{cid2}').text and OBS in ci.get(f'/cierres/{cid2}').text


def test_editar_reenviar_otro_cajero_no_puede(app, sint):
    c, ci, cid = rechazado(app, sint)
    cb = login(app, 'beto')
    assert post(cb, f'/cierres/{cid}/editar', f'/cierres/{cid}').status_code == 404


def test_editar_reenviar_reemplazar_archivos_y_fecha(app, sint, tmp_path):
    c, ci, cid = rechazado(app, sint)
    bid2 = bid_de(post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}'))
    # reemplazar el banco por otro archivo del mismo contenido lógico (nueva copia)
    nuevo = S.crear_a2(str(tmp_path / 'banco_nuevo.xlsx'))
    tk = token(c, f'/borrador/{bid2}')
    with open(nuevo, 'rb') as f:
        r = c.post(f'/borrador/{bid2}/archivos', data={'csrf': tk}, files={'banco': ('banco_nuevo.xlsx', f)})
    assert r.status_code == 303
    d = os.path.join(app.state.almacen.dir, 'borradores', bid2)
    assert sha(os.path.join(d, 'banco_qr_original.xlsx')) == sha(nuevo) and leer_estado(app, bid2)['nombres_subidos']['banco'] == 'banco_nuevo.xlsx'
    assert db(app).execute("select count(*) from auditoria where accion='archivos_reemplazados'").fetchone()[0] == 1
    # archivo inválido: se rechaza y no cambia nada
    malo = tmp_path / 'malo.xlsx'; malo.write_bytes(b'no es excel')
    with open(malo, 'rb') as f:
        c.post(f'/borrador/{bid2}/archivos', data={'csrf': token(c, f'/borrador/{bid2}')}, files={'banco': ('malo.xlsx', f)})
    assert sha(os.path.join(d, 'banco_qr_original.xlsx')) == sha(nuevo)
    # otra fecha: no se puede presentar como corrección del cierre rechazado
    otra = S.crear_a1(str(tmp_path / 'otro_dia.xlsx'), fecha=S.dt.date(2026, 10, 17))
    otro_banco = S.crear_a2(str(tmp_path / 'banco_otro.xlsx'), fecha=S.dt.date(2026, 10, 17))
    with open(otra, 'rb') as f1, open(otro_banco, 'rb') as f2:
        c.post(f'/borrador/{bid2}/archivos', data={'csrf': token(c, f'/borrador/{bid2}')}, files={'movimiento': ('otro_dia.xlsx', f1), 'banco': ('banco_otro.xlsx', f2)})
    html = pag(c, bid2).text
    assert 'debe ser la misma del cierre rechazado' in html and 'data-abrir="dlg-presentar"' not in html
    presentar(c, bid2); assert db(app).execute('select count(*) from cierres').fetchone()[0] == 1


def test_archivos_solo_en_correccion(app, sint):
    c, bid = borrador_nuevo(app, sint)
    r = c.post(f'/borrador/{bid}/archivos', data={'csrf': token(c, f'/borrador/{bid}')})
    assert r.status_code == 403


# ---------- migración ----------
ESQUEMA_V1 = """
CREATE TABLE cierres (
  id INTEGER PRIMARY KEY AUTOINCREMENT, correlativo TEXT NOT NULL UNIQUE, fecha_cierre TEXT NOT NULL, usuario TEXT NOT NULL, nombre TEXT, caja TEXT,
  qr_registrado REAL, qr_conciliado REAL, tarjeta_declarada REAL, tarjeta_sistema REAL, total_cierre REAL,
  estado_revision TEXT NOT NULL DEFAULT 'PENDIENTE DE REVISIÓN', version INTEGER NOT NULL, reemplaza_a INTEGER, motivo_reemplazo TEXT,
  presentado_en TEXT NOT NULL, ruta TEXT NOT NULL, sha256_movimiento TEXT, sha256_banco TEXT);
CREATE INDEX ix_cierres_fu ON cierres(fecha_cierre, usuario);
CREATE TABLE auditoria (id INTEGER PRIMARY KEY AUTOINCREMENT, momento TEXT NOT NULL, usuario TEXT, accion TEXT NOT NULL, detalle_json TEXT);
CREATE TABLE cierre_movimientos (cierre_id INTEGER NOT NULL REFERENCES cierres(id), id_transaccion TEXT NOT NULL, factura TEXT, monto REAL);
CREATE INDEX ix_cm_id ON cierre_movimientos(id_transaccion);
"""


def crear_db_v1(ruta):
    con = sqlite3.connect(ruta); con.executescript(ESQUEMA_V1)
    con.execute("INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, total_cierre, "
                "estado_revision, version, presentado_en, ruta) VALUES ('CC-2026-00001','2026-10-15','ana','Ana Prueba','SFC1',1000,1000,100,100,1100,'REEMPLAZADO',1,'2026-10-15 17:00:00','cierres/x/v1')")
    con.execute("INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, total_cierre, "
                "estado_revision, version, reemplaza_a, motivo_reemplazo, presentado_en, ruta) VALUES ('CC-2026-00002','2026-10-15','ana','Ana Prueba','SFC1',1200,1200,100,100,1300,'PENDIENTE DE REVISIÓN',2,1,'corrección','2026-10-15 18:00:00','cierres/x/v2')")
    con.execute("INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES ('2026-10-15 17:00:00','ana','presentar_cierre','{}')")
    con.execute("INSERT INTO cierre_movimientos VALUES (2,'700000001','30001',500.0)")
    con.commit(); con.close()


def test_migracion_desde_esquema_v1(tmp_path, monkeypatch):
    data = tmp_path / 'data'; data.mkdir(); ruta = str(data / 'cierres.db'); crear_db_v1(ruta)
    from app import db as dbm
    from app.servicio import Almacen
    for _ in range(2):    # idempotente: arrancar dos veces no falla ni duplica
        Almacen(str(data))
        con = sqlite3.connect(ruta); con.row_factory = sqlite3.Row
        cols = {r[1] for r in con.execute('PRAGMA table_info(cierres)')}
        assert {'revisado_por', 'revisado_en', 'observacion_revision', 'ci_total'} <= cols
        assert con.execute('PRAGMA user_version').fetchone()[0] == dbm.VERSION_ESQUEMA
        assert con.execute("select name from sqlite_master where name='cierre_ci'").fetchone()
        filas = con.execute('select correlativo, estado_revision, total_cierre, version, reemplaza_a, revisado_por, ci_total from cierres order by id').fetchall()
        assert [tuple(f) for f in filas] == [('CC-2026-00001', 'REEMPLAZADO', 1100, 1, None, None, None), ('CC-2026-00002', 'PENDIENTE DE REVISIÓN', 1300, 2, 1, None, None)]
        assert con.execute('select count(*) from auditoria').fetchone()[0] == 1 and con.execute('select count(*) from cierre_movimientos').fetchone()[0] == 1
        con.close()
    # los cierres v1 se pueden revisar con el esquema nuevo (la aplicación arranca sobre la DB migrada)
    monkeypatch.setenv('DATA_DIR', str(data)); monkeypatch.setenv('APP_USERS', json.dumps(__import__('tests.conftest', fromlist=['x']).USUARIOS))
    monkeypatch.setenv('SECRET_KEY', 'k'); monkeypatch.setenv('COOKIE_SECURE', '0')
    from app.main import create_app
    app = create_app(); ci = login(app, 'ing')
    assert 'CC-2026-00002' in ci.get('/?estado=PENDIENTE DE REVISIÓN').text
    os.makedirs(data / 'cierres' / 'x' / 'v2')       # cierre.json ausente en v1 sintético: el detalle igual abre
    assert ci.get('/cierres/2').status_code == 200
    r = post(ci, '/cierres/2/revisar', '/cierres/2', decision='rechazar', observacion='Observación de prueba larga')
    assert dbm.conectar(ruta).execute('select estado_revision, revisado_por from cierres where id=2').fetchone()[:] == ('RECHAZADO', 'ing')
