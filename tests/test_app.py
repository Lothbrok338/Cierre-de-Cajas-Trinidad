import glob, hashlib, json, os, re, sqlite3
import pytest
from fastapi.testclient import TestClient
from tests.conftest import login, nuevo_cliente, token, post, subir, CLAVES
from tests import sinteticos as S

DIR_REAL = os.environ.get('DATOS_PRUEBA_DIR', '')


def sha(p): return hashlib.sha256(open(p, 'rb').read()).hexdigest()
def bid_de(r): return r.headers['location'].rsplit('/', 1)[1]
def db(app): return sqlite3.connect(os.path.join(app.state.almacen.dir, 'cierres.db'))
def pag(c, bid): return c.get(f'/borrador/{bid}')
def estado_txt(html):
    for k in ('PENDIENTE DE CONFIRMACIÓN', 'NO CUADRA', 'CUADRA'):
        if re.search(rf'<strong>{k}</strong>', html): return k
    return None
def presentar(c, bid, **kw): return post(c, f'/borrador/{bid}/presentar', f'/borrador/{bid}', declaro='1', **kw)


# ---------- login ----------
def test_salud_y_redireccion(app):
    c = nuevo_cliente(app)
    assert c.get('/salud').status_code == 200
    for ruta in ('/', '/nuevo', '/borrador/x', '/cierres/1'):
        r = c.get(ruta); assert r.status_code == 303 and r.headers['location'] == '/login'


def test_login_correcto_e_incorrecto(app):
    c = nuevo_cliente(app)
    r = c.post('/login', data={'usuario': 'ana', 'clave': 'mala', 'csrf': token(c)})
    assert r.status_code == 401 and 'Usuario o clave incorrectos' in r.text
    assert c.get('/').status_code == 303
    r = c.post('/login', data={'usuario': 'ana', 'clave': CLAVES['ana'], 'csrf': token(c)})
    assert r.status_code == 303 and c.get('/').status_code == 200
    assert 'httponly' in r.headers['set-cookie'].lower()


def test_login_sin_csrf_rechazado(app):
    c = nuevo_cliente(app); c.get('/login')
    assert c.post('/login', data={'usuario': 'ana', 'clave': CLAVES['ana']}).status_code == 403


def test_sistema_no_configurado(tmp_path, monkeypatch):
    monkeypatch.setenv('DATA_DIR', str(tmp_path / 'd')); monkeypatch.delenv('APP_USERS', raising=False)
    from app.main import create_app
    c = TestClient(create_app(), follow_redirects=False)
    assert 'Sistema no configurado' in c.get('/login').text
    assert 'Sistema no configurado' in c.post('/login', data={'usuario': 'a', 'clave': 'b'}).text


def test_bloqueo_por_intentos(app):
    c = nuevo_cliente(app)
    for _ in range(5): c.post('/login', data={'usuario': 'beto', 'clave': 'x', 'csrf': token(c)})
    r = c.post('/login', data={'usuario': 'beto', 'clave': CLAVES['beto'], 'csrf': token(c)})
    assert r.status_code == 429


def test_hash_clave_script():
    from app.auth import hash_clave, verificar_clave
    h = hash_clave('secreto-123', 1000)
    assert h.startswith('pbkdf2_sha256$') and verificar_clave('secreto-123', h) and not verificar_clave('otro', h)


# ---------- validacion de archivos ----------
def test_validaciones_de_subida(app, sint, tmp_path):
    c = login(app, 'ana')
    r = subir(c, sint['a2'], sint['a2'])    # banco como movimiento
    assert r.status_code == 400 and 'no parece ser el reporte de Ingreso Diario' in r.text
    r = subir(c, sint['a1'], sint['a1'])    # movimiento como banco
    assert r.status_code == 400 and 'QR BANECO' in r.text
    malo = tmp_path / 'x.txt'; malo.write_text('hola')
    assert 'debe ser' in subir(c, str(malo), sint['a2']).text
    grande = tmp_path / 'grande.xlsx'; grande.write_bytes(b'PK\x03\x04' + b'0' * (10 * 1024 * 1024 + 10))
    assert 'supera el máximo de 10 MB' in subir(c, sint['a1'], str(grande)).text
    assert 'total cobrado con tarjetas' in subir(c, sint['a1'], sint['a2'], tarjeta='abc').text
    falso = tmp_path / 'falso.xls'; falso.write_bytes(b'esto no es excel')
    assert 'no parece ser el reporte de Ingreso Diario' in subir(c, str(falso), sint['a2']).text
    assert not os.listdir(os.path.join(app.state.almacen.dir, 'borradores'))


# ---------- flujo completo (sintetico) ----------
def flujo_hasta_cuadra(app, sint, usuario='ana'):
    c = login(app, usuario)
    r = subir(c, sint['a1'], sint['a2']); assert r.status_code == 303, r.text
    bid = bid_de(r); html = pag(c, bid).text
    assert estado_txt(html) == 'PENDIENTE DE CONFIRMACIÓN' and 'Sí, corresponde' in html
    assert 'disabled' in html
    r = post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003'); assert r.status_code == 303
    html = pag(c, bid).text
    assert estado_txt(html) == 'CUADRA' and 'puede presentar su cierre' in html
    return c, bid


def test_flujo_completo_sintetico(app, sint):
    c, bid = flujo_hasta_cuadra(app, sint)
    d = os.path.join(app.state.almacen.dir, 'borradores', bid)
    assert sha(os.path.join(d, 'movimiento_original.xlsx')) == sha(sint['a1'])
    assert sha(os.path.join(d, 'banco_qr_original.xlsx')) == sha(sint['a2'])
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/')
    rows = db(app).execute('select correlativo, fecha_cierre, usuario, caja, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, '
                           'total_cierre, estado_revision, version, ruta, sha256_movimiento, sha256_banco from cierres').fetchall()
    assert len(rows) == 1
    (corr, fecha, usu, caja, qrr, qrc, td, ts, tot, est, ver, ruta, hm, hb) = rows[0]
    assert corr.startswith('CC-') and corr.endswith('-00001') and fecha == '2026-10-16' and usu == 'ana' and caja == 'SFC999'
    assert (qrr, qrc, td, ts, tot) == (1500, 1500, 200, 200, 1700) and est == 'PENDIENTE DE REVISIÓN' and ver == 1
    dest = os.path.join(app.state.almacen.dir, ruta)
    assert ruta.replace(os.sep, '/') == 'cierres/2026/10/2026-10-16/ana/v1'
    assert sorted(os.listdir(dest)) == ['banco_qr_original.xlsx', 'cierre.json', 'cierre.pdf', 'conciliacion.xlsx', 'movimiento_original.xlsx']
    assert sha(os.path.join(dest, 'movimiento_original.xlsx')) == hm == sha(sint['a1'])
    assert sha(os.path.join(dest, 'banco_qr_original.xlsx')) == hb == sha(sint['a2'])
    assert open(os.path.join(dest, 'cierre.pdf'), 'rb').read(5) == b'%PDF-'
    cj = json.load(open(os.path.join(dest, 'cierre.json'), encoding='utf-8'))
    assert cj['sha256_movimiento'] == hm and cj['confirmaciones'] == ['30003'] and cj['usuario'] == 'ana'
    assert [x['concepto'] for x in cj['conceptos']] == ['QR BANECO', 'Tarjetas'] and cj['conceptos'][0]['diferencia'] == 0
    import openpyxl
    assert 'Resumen' in openpyxl.load_workbook(os.path.join(dest, 'conciliacion.xlsx')).sheetnames
    acc = [a[0] for a in db(app).execute('select accion from auditoria').fetchall()]
    for a in ('login_ok', 'borrador_creado', 'confirmar_coincidencia', 'presentar_cierre'): assert a in acc
    # detalle y descargas
    cid = r.headers['location'].rsplit('/', 1)[1]
    assert corr in c.get(f'/cierres/{cid}').text
    assert c.get(f'/cierres/{cid}/archivo/cierre.pdf').content[:5] == b'%PDF-'
    assert c.get(f'/cierres/{cid}/archivo/movimiento_original.xlsx').content == open(sint['a1'], 'rb').read()
    assert c.get(f'/cierres/{cid}/archivo/../../cierres.db').status_code in (404, 405, 422)
    assert 'Mis cierres' in c.get('/').text and corr in c.get('/').text
    assert c.get(f'/borrador/{bid}').status_code == 404      # ya presentado


# ---------- reglas de presentacion ----------
def test_presentar_bloqueado_si_no_cuadra(app, sint):
    c = login(app, 'ana')
    bid = bid_de(subir(c, sint['a1'], sint['a2_sin']))
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    html = pag(c, bid).text
    assert estado_txt(html) == 'NO CUADRA' and 'no puede presentar' in html and 'Revise las facturas sin respaldo' in html
    assert 'data-abrir="dlg-presentar"' not in html and 'disabled' in html and 'Asignar' not in html   # no hay candidatos
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'] == f'/borrador/{bid}'
    assert 'no se puede presentar' in pag(c, bid).text.lower() or 'No se puede presentar' in pag(c, bid).text
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 0


def test_presentar_bloqueado_si_tarjeta_difiere(app, sint):
    c = login(app, 'ana')
    bid = bid_de(subir(c, sint['a1'], sint['a2'], tarjeta='150'))
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    html = pag(c, bid).text
    assert 'data-abrir="dlg-presentar"' not in html and 'Las tarjetas no cuadran' in html and 'declaró Bs 150.00' in html
    presentar(c, bid)
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 0


def test_presentar_pendiente_bloqueado(app, sint):
    c = login(app, 'ana'); bid = bid_de(subir(c, sint['a1'], sint['a2']))
    presentar(c, bid); assert db(app).execute('select count(*) from cierres').fetchone()[0] == 0


def test_deshacer_y_asignar(app, sint, tmp_path):
    from tests import sinteticos as s
    banco = [b for b in s.BANCO if b[0] != '700000002'] + [('700000009', 'ZZZ XXX WWW', '10:29:00', 300)]
    a2 = s.crear_a2(str(tmp_path / 'b9.xlsx'), banco)
    c = login(app, 'ana'); bid = bid_de(subir(c, sint['a1'], a2))
    html = pag(c, bid).text
    assert estado_txt(html) == 'NO CUADRA' and '700000009' in html and 'Asignar' in html and 'ZZZ XXX WWW' in html
    post(c, f'/borrador/{bid}/asignar', f'/borrador/{bid}', factura='30002', id_transaccion='700000009')
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    html = pag(c, bid).text
    assert estado_txt(html) == 'CUADRA' and 'Deshacer' in html and 'CONFIRMADA POR USTED' in html
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30002')
    assert estado_txt(pag(c, bid).text) == 'NO CUADRA'
    est = json.load(open(os.path.join(app.state.almacen.dir, 'borradores', bid, 'estado.json')))
    assert est['elecciones'] == {} and est['confirmaciones'] == ['30003']
    acc = [a[0] for a in db(app).execute('select accion from auditoria').fetchall()]
    assert 'asignar_movimiento' in acc and 'deshacer_confirmacion' in acc


def test_reemplazo_exige_motivo(app, sint):
    c, bid = flujo_hasta_cuadra(app, sint); presentar(c, bid)
    r = subir(c, sint['a1'], sint['a2']); bid2 = bid_de(r)
    post(c, f'/borrador/{bid2}/confirmar', f'/borrador/{bid2}', factura='30003')
    html = pag(c, bid2).text
    assert 'Motivo del reemplazo' in html and 'reemplazará' in html
    r = presentar(c, bid2, motivo='   ')
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 1
    assert 'motivo del reemplazo' in pag(c, bid2).text.lower()
    r = presentar(c, bid2, motivo='Se corrigió un monto'); assert r.headers['location'].startswith('/cierres/')
    rows = db(app).execute('select correlativo, version, estado_revision, reemplaza_a, motivo_reemplazo, ruta from cierres order by id').fetchall()
    assert [(x[1], x[2]) for x in rows] == [(1, 'REEMPLAZADO'), (2, 'PENDIENTE DE REVISIÓN')]
    assert rows[1][3] == 1 and rows[1][4] == 'Se corrigió un monto' and rows[1][0].endswith('00002')
    base = app.state.almacen.dir
    assert os.path.isfile(os.path.join(base, rows[0][5], 'cierre.pdf')) and os.path.isfile(os.path.join(base, rows[1][5], 'cierre.pdf'))   # el anterior no se borra
    assert os.path.basename(rows[1][5]) == 'v2'


def test_antiduplicado_entre_usuarios(app, sint):
    c, bid = flujo_hasta_cuadra(app, sint, 'ana'); presentar(c, bid)
    corr = db(app).execute('select correlativo from cierres').fetchone()[0]
    cb = login(app, 'beto'); b2 = bid_de(subir(cb, sint['a1'], sint['a2']))
    post(cb, f'/borrador/{b2}/confirmar', f'/borrador/{b2}', factura='30003')
    html = pag(cb, b2).text
    assert corr in html and '700000001' in html and 'ya está en el cierre' in html and 'data-abrir="dlg-presentar"' not in html
    presentar(cb, b2)    # forzado por POST directo: tambien se bloquea en servidor
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 1
    assert db(app).execute("select count(*) from auditoria where accion='presentar_bloqueado'").fetchone()[0] == 1


def test_borrador_y_cierre_ajenos(app, sint):
    c, bid = flujo_hasta_cuadra(app, sint, 'ana')
    cb = login(app, 'beto')
    assert cb.get(f'/borrador/{bid}').status_code in (403, 404)
    assert post(cb, f'/borrador/{bid}/confirmar', '/nuevo', factura='30003').status_code in (403, 404)
    assert presentar(cb, bid).status_code in (403, 404)
    assert cb.get('/borrador/no-es-uuid').status_code == 404
    r = presentar(c, bid); cid = r.headers['location'].rsplit('/', 1)[1]
    assert cb.get(f'/cierres/{cid}').status_code in (403, 404) and cb.get(f'/cierres/{cid}/archivo/cierre.pdf').status_code in (403, 404)
    ci = login(app, 'ing')
    assert ci.get(f'/cierres/{cid}').status_code == 200 and ci.get(f'/cierres/{cid}/archivo/cierre.pdf').status_code == 200
    assert ci.get('/nuevo').status_code == 403


def test_pdf_una_pagina_mas_anexo(app, sint):
    c, bid = flujo_hasta_cuadra(app, sint); presentar(c, bid)
    pdf = os.path.join(app.state.almacen.dir, 'cierres/2026/10/2026-10-16/ana/v1/cierre.pdf')
    data = open(pdf, 'rb').read()
    assert len(re.findall(rb'/Type\s*/Page[^s]', data)) == 2


# ---------- datos reales ----------
real = pytest.mark.skipif(not DIR_REAL or not glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls')), reason='DATOS_PRUEBA_DIR sin datos reales')


@real
def test_flujo_real_tarjeta_1191(app):
    a1 = glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls'))[0]; a2 = glob.glob(os.path.join(DIR_REAL, '*Reporte_pagos*.xlsx'))[0]
    c = login(app, 'ana'); r = subir(c, a1, a2, '1191'); assert r.status_code == 303, r.text
    bid = bid_de(r); html = pag(c, bid).text
    assert estado_txt(html) == 'PENDIENTE DE CONFIRMACIÓN' and '10,784.00' in html and '615.00' in html
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='20048')
    assert estado_txt(pag(c, bid).text) == 'CUADRA'
    r = presentar(c, bid); assert r.status_code == 303
    row = db(app).execute('select correlativo, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, total_cierre, caja, ruta, sha256_movimiento, sha256_banco from cierres').fetchone()
    assert row[1:7] == (10784, 10784, 1191, 1191, 11975, 'SFC401')
    dest = os.path.join(app.state.almacen.dir, row[7])
    assert sha(os.path.join(dest, 'movimiento_original.xls')) == row[8] == sha(a1) and sha(os.path.join(dest, 'banco_qr_original.xlsx')) == row[9] == sha(a2)
    assert open(os.path.join(dest, 'cierre.pdf'), 'rb').read(5) == b'%PDF-'
    # reutiliza todas las facturas: otro cajero bloqueado
    cb = login(app, 'beto'); b2 = bid_de(subir(cb, a1, a2, '1191'))
    post(cb, f'/borrador/{b2}/confirmar', f'/borrador/{b2}', factura='20048')
    assert 'ya está en el cierre' in pag(cb, b2).text and row[0] in pag(cb, b2).text
