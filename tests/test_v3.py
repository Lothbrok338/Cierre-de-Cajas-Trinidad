"""Pruebas v3: usuarios en base de datos, rol superadmin, revertir revisión, administración, auditoría, mi cuenta y migración v2 -> v3."""
import json, os, re, sqlite3
import pytest
from fastapi.testclient import TestClient
from tests.conftest import login, nuevo_cliente, token, post, subir, USUARIOS, CLAVES
from tests.test_app import bid_de, db, presentar
from tests.test_v2 import borrador_nuevo, ESQUEMA_V1, crear_db_v1
from app import auth, db as dbm
from app.servicio import Almacen


def cierre_presentado(app, sint, usuario='ana'):
    c, bid = borrador_nuevo(app, sint, usuario)
    from tests.test_v2 import flash_de
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    return c, int(r.headers['location'].rsplit('/', 1)[1])


def revisar(c, cid, decision, obs=''):
    return post(c, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision=decision, observacion=obs)


def estado(app, cid):
    return db(app).execute('select estado_revision, revisado_por, revisado_en, observacion_revision from cierres where id=?', (cid,)).fetchone()


def entrar(app, usuario, clave):
    c = nuevo_cliente(app); r = c.post('/login', data={'usuario': usuario, 'clave': clave, 'csrf': token(c)}); return c, r


def flash(c, ruta='/'):
    m = re.search(r'role="status">(.*?)</div>', c.get(ruta).text, re.S); return m.group(1).strip() if m else ''


# ---------- login / usuarios en BD ----------
def test_login_sin_distinguir_mayusculas_y_nombre_visible(app):
    for escrito in ('ANA', 'Ana', ' aNa '):
        c, r = entrar(app, escrito, CLAVES['ana']); assert r.status_code == 303, escrito
        assert 'Ana Prueba' in c.get('/').text
    assert db(app).execute("select usuario from usuarios where usuario='ana'").fetchone()


def test_usuario_inactivo_no_entra_y_sesion_abierta_cae(app):
    c_ana = login(app, 'ana'); sup = login(app, 'sup')
    post(sup, '/admin/usuarios/ana/activo', '/admin/usuarios', activo='0')
    assert c_ana.get('/').status_code == 303 and c_ana.get('/').headers['location'] == '/login'   # la sesión abierta deja de valer
    c, r = entrar(app, 'ana', CLAVES['ana']); assert r.status_code == 401
    post(sup, '/admin/usuarios/ana/activo', '/admin/usuarios', activo='1')
    assert entrar(app, 'ana', CLAVES['ana'])[1].status_code == 303


def test_bloqueo_por_intentos_insensible_a_mayusculas(app):
    c = nuevo_cliente(app)
    for u in ('BETO', 'beto', 'Beto', 'bETO', 'beto'): c.post('/login', data={'usuario': u, 'clave': 'x', 'csrf': token(c)})
    assert c.post('/login', data={'usuario': 'beto', 'clave': CLAVES['beto'], 'csrf': token(c)}).status_code == 429


def test_siembra_no_pisa_cambios_de_la_ui(tmp_path, monkeypatch):
    monkeypatch.setenv('DATA_DIR', str(tmp_path / 'd')); monkeypatch.setenv('SECRET_KEY', 'k'); monkeypatch.setenv('COOKIE_SECURE', '0')
    monkeypatch.setenv('APP_USERS', json.dumps(USUARIOS))
    from app.main import create_app
    app = create_app(); sup = login(app, 'sup')
    post(sup, '/admin/usuarios/ana/rol', '/admin/usuarios', rol='ingresos')
    post(sup, '/admin/usuarios/beto/activo', '/admin/usuarios', activo='0')
    post(sup, '/admin/usuarios/ana/clave', '/admin/usuarios', clave='NuevaClave-9', clave2='NuevaClave-9')
    app2 = create_app()    # reinicio con el mismo APP_USERS: no debe revertir nada
    r = db(app2).execute("select rol, activo from usuarios where usuario in ('ana','beto') order by usuario").fetchall()
    assert r == [('ingresos', 1), ('cajero', 0)]
    assert entrar(app2, 'ana', 'NuevaClave-9')[1].status_code == 303 and entrar(app2, 'ana', CLAVES['ana'])[1].status_code == 401
    # con "forzar": true, esa entrada sí se vuelve a aplicar (y reactiva)
    forzado = {**USUARIOS, 'ana': {**USUARIOS['ana'], 'forzar': True}, 'nuevo': {'nombre': 'Nuevo', 'rol': 'cajero', 'clave_hash': USUARIOS['ana']['clave_hash']}}
    monkeypatch.setenv('APP_USERS', json.dumps(forzado))
    app3 = create_app()
    assert db(app3).execute("select rol from usuarios where usuario='ana'").fetchone()[0] == 'cajero'
    assert entrar(app3, 'ana', CLAVES['ana'])[1].status_code == 303
    assert db(app3).execute("select activo from usuarios where usuario='beto'").fetchone()[0] == 0     # beto no estaba forzado
    assert db(app3).execute("select count(*) from usuarios where usuario='nuevo'").fetchone()[0] == 1


def test_cargar_usuarios_minusculas_y_rol_superadmin():
    h = auth.hash_clave('x', 1000)
    u = auth.cargar_usuarios(json.dumps({'Gabo': {'nombre': 'G', 'rol': 'superadmin', 'clave_hash': h, 'forzar': True},
                                         'x': {'nombre': 'X', 'rol': 'otro', 'clave_hash': h}}))
    assert list(u) == ['gabo'] and u['gabo']['forzar'] is True


# ---------- superadmin ----------
def test_superadmin_crea_cierre_y_aprueba(app, sint):
    sup, cid = cierre_presentado(app, sint, 'sup')
    assert sup.get('/').status_code == 200 and 'CC-' in sup.get('/').text
    html = sup.get('/nuevo'); assert html.status_code == 200
    d = sup.get(f'/cierres/{cid}').text
    assert 'data-abrir="dlg-aprobar"' in d and 'class="btn btn-ok"' in d
    assert revisar(sup, cid, 'aprobar').status_code == 303
    assert estado(app, cid)[:2] == ('CONFORME', 'sup')
    assert sup.get(f'/cierres/{cid}/archivo/cierre.pdf').status_code == 200


def test_superadmin_ve_cierres_de_otros_y_puede_rechazar(app, sint):
    ana, cid = cierre_presentado(app, sint, 'ana'); sup = login(app, 'sup')
    assert 'CC-' in sup.get('/').text and sup.get(f'/cierres/{cid}').status_code == 200
    assert revisar(sup, cid, 'rechazar', 'Falta el voucher del POS').status_code == 303
    assert estado(app, cid)[0] == 'RECHAZADO'


def test_revertir_conforme_a_pendiente_con_motivo(app, sint):
    ana, cid = cierre_presentado(app, sint); ing = login(app, 'ing'); sup = login(app, 'sup')
    revisar(ing, cid, 'aprobar', 'Todo bien'); assert estado(app, cid)[0] == 'CONFORME'
    corto = post(sup, f'/cierres/{cid}/revertir', f'/cierres/{cid}', motivo='corto')
    assert corto.status_code == 303 and estado(app, cid)[0] == 'CONFORME' and 'al menos 10' in flash(sup, f'/cierres/{cid}')
    assert post(sup, f'/cierres/{cid}/revertir', f'/cierres/{cid}', motivo='Se aprobó por error, falta verificar').status_code == 303
    assert estado(app, cid) == ('PENDIENTE DE REVISIÓN', None, None, None)
    h = db(app).execute('select usuario, accion, estado_anterior, estado_nuevo, observacion from revisiones_historial where cierre_id=? order by id', (cid,)).fetchall()
    assert h == [('ing', 'aprobar', 'PENDIENTE DE REVISIÓN', 'CONFORME', 'Todo bien'),
                 ('sup', 'revertir', 'CONFORME', 'PENDIENTE DE REVISIÓN', 'Se aprobó por error, falta verificar')]
    a = db(app).execute("select detalle_json from auditoria where accion='revision_revertida'").fetchone()[0]
    assert 'Todo bien' in a and '"revisado_por_anterior": "ing"' in a and 'falta verificar' in a
    # historial visible para ingresos y superadmin; el cajero no lo ve
    assert 'Historial de revisiones' in ing.get(f'/cierres/{cid}').text and 'Revirtió la revisión' in sup.get(f'/cierres/{cid}').text
    assert 'Historial de revisiones' not in ana.get(f'/cierres/{cid}').text
    assert 'Esperando revisión de Ingresos' in ana.get(f'/cierres/{cid}').text
    # y vuelve a poder revisarse
    assert revisar(ing, cid, 'rechazar', 'Ahora sí: observación completa').status_code == 303 and estado(app, cid)[0] == 'RECHAZADO'


def test_revertir_rechazado(app, sint):
    ana, cid = cierre_presentado(app, sint); ing = login(app, 'ing'); sup = login(app, 'sup')
    revisar(ing, cid, 'rechazar', 'Observación de prueba larga')
    post(sup, f'/cierres/{cid}/revertir', f'/cierres/{cid}', motivo='Rechazo equivocado del revisor')
    assert estado(app, cid) == ('PENDIENTE DE REVISIÓN', None, None, None)
    assert 'Observación de prueba larga' in sup.get(f'/cierres/{cid}').text    # sigue en el historial


def test_revertir_bloqueado_para_otros_roles_y_estados(app, sint):
    ana, cid = cierre_presentado(app, sint); ing = login(app, 'ing'); sup = login(app, 'sup')
    # PENDIENTE: no se puede revertir
    post(sup, f'/cierres/{cid}/revertir', f'/cierres/{cid}', motivo='Motivo suficientemente largo')
    assert 'solo se puede revertir un cierre CONFORME o RECHAZADO' in flash(sup, f'/cierres/{cid}')
    revisar(ing, cid, 'aprobar')
    assert 'dlg-revertir' not in ing.get(f'/cierres/{cid}').text and 'dlg-revertir' not in ana.get(f'/cierres/{cid}').text
    for c in (ing, ana):
        assert post(c, f'/cierres/{cid}/revertir', f'/cierres/{cid}', motivo='Motivo suficientemente largo').status_code == 403
    assert estado(app, cid)[0] == 'CONFORME'


def test_revertir_bloqueado_si_reemplazado(app, sint):
    ing = login(app, 'ing'); sup = login(app, 'sup')
    # rechazar -> el cajero corrige y reenvía -> v1 queda REEMPLAZADO
    _, cid2 = cierre_presentado(app, sint, 'beto')
    revisar(ing, cid2, 'rechazar', 'Corrija la tarjeta declarada')
    beto = login(app, 'beto'); r = post(beto, f'/cierres/{cid2}/editar', f'/cierres/{cid2}'); assert r.status_code == 303
    from tests.test_v2 import flash_de
    post(beto, f'/borrador/{bid_de(r)}/confirmar', f'/borrador/{bid_de(r)}', factura='30003')
    p = presentar(beto, bid_de(r)); assert p.status_code == 303 and p.headers['location'].startswith('/cierres/'), flash_de(beto, bid_de(r))
    assert estado(app, cid2)[0] == 'REEMPLAZADO'
    post(sup, f'/cierres/{cid2}/revertir', f'/cierres/{cid2}', motivo='Intento de revertir un reemplazado')
    assert estado(app, cid2)[0] == 'REEMPLAZADO' and 'REEMPLAZADO' in flash(sup, f'/cierres/{cid2}')
    assert 'dlg-revertir' not in sup.get(f'/cierres/{cid2}').text


def test_ingresos_y_cajero_siguen_como_antes(app, sint):
    ana, cid = cierre_presentado(app, sint); ing = login(app, 'ing')
    assert ing.get('/nuevo').status_code == 403 and ana.get('/nuevo').status_code == 200
    for ruta in ('/admin/usuarios', '/admin/auditoria'):
        assert ing.get(ruta).status_code == 403 and ana.get(ruta).status_code == 403
    assert ana.post('/admin/usuarios/crear', data={'csrf': token(ana, '/'), 'usuario': 'zzz'}).status_code == 403
    assert ing.get('/').status_code == 200 and 'Usuarios' not in ing.get('/').text and 'Mi cuenta' in ing.get('/').text
    assert ana.get(f'/cierres/{cid}').status_code == 200
    assert login(app, 'beto').get(f'/cierres/{cid}').status_code == 404


# ---------- administración de usuarios ----------
def test_admin_usuarios_crud(app):
    sup = login(app, 'sup'); assert 'Administración de usuarios' in sup.get('/admin/usuarios').text
    crear = lambda **kw: post(sup, '/admin/usuarios/crear', '/admin/usuarios', **{'usuario': 'Nuevo.User', 'nombre': 'Nuevo Usuario', 'rol': 'cajero', 'clave': 'ClaveLarga-1', 'clave2': 'ClaveLarga-1', **kw})
    crear(clave2='otra-cosa-1'); assert 'no coincide' in flash(sup, '/admin/usuarios')
    crear(clave='corta', clave2='corta'); assert 'al menos 8' in flash(sup, '/admin/usuarios')
    crear(rol='dios'); assert 'Rol no válido' in flash(sup, '/admin/usuarios')
    crear(usuario='a b'); assert 'letras' in flash(sup, '/admin/usuarios')
    assert db(app).execute("select count(*) from usuarios where usuario like 'nuevo%'").fetchone()[0] == 0
    crear(); assert db(app).execute("select usuario, nombre, rol, activo from usuarios where usuario='nuevo.user'").fetchone() == ('nuevo.user', 'Nuevo Usuario', 'cajero', 1)
    hash_guardado = db(app).execute("select clave_hash from usuarios where usuario='nuevo.user'").fetchone()[0]
    assert hash_guardado.startswith('pbkdf2_sha256$') and 'ClaveLarga-1' not in hash_guardado
    crear(usuario='NUEVO.user'); assert 'ya existe' in flash(sup, '/admin/usuarios')
    assert entrar(app, 'NUEVO.USER', 'ClaveLarga-1')[1].status_code == 303
    post(sup, '/admin/usuarios/nuevo.user/rol', '/admin/usuarios', rol='ingresos')
    assert db(app).execute("select rol from usuarios where usuario='nuevo.user'").fetchone()[0] == 'ingresos'
    post(sup, '/admin/usuarios/nuevo.user/clave', '/admin/usuarios', clave='Restablecida-2', clave2='Restablecida-2')
    assert entrar(app, 'nuevo.user', 'ClaveLarga-1')[1].status_code == 401 and entrar(app, 'nuevo.user', 'Restablecida-2')[1].status_code == 303
    post(sup, '/admin/usuarios/nuevo.user/clave', '/admin/usuarios', clave='corta', clave2='corta'); assert 'al menos 8' in flash(sup, '/admin/usuarios')
    post(sup, '/admin/usuarios/nuevo.user/activo', '/admin/usuarios', activo='0')
    assert entrar(app, 'nuevo.user', 'Restablecida-2')[1].status_code == 401
    # auditoría: queda registro y nunca la clave
    acc = [r[0] for r in db(app).execute("select accion from auditoria where accion like 'usuario_%'")]
    assert {'usuario_creado', 'usuario_rol_cambiado', 'usuario_clave_restablecida', 'usuario_desactivado'} <= set(acc)
    todo = ' '.join(r[0] or '' for r in db(app).execute('select detalle_json from auditoria'))
    assert 'ClaveLarga-1' not in todo and 'Restablecida-2' not in todo


def test_no_se_puede_dejar_el_sistema_sin_superadmin(app):
    sup = login(app, 'sup')
    post(sup, '/admin/usuarios/sup/activo', '/admin/usuarios', activo='0')
    assert 'al menos un superadministrador activo' in flash(sup, '/admin/usuarios')
    post(sup, '/admin/usuarios/sup/rol', '/admin/usuarios', rol='cajero')
    assert 'al menos un superadministrador activo' in flash(sup, '/admin/usuarios')
    assert db(app).execute("select rol, activo from usuarios where usuario='sup'").fetchone() == ('superadmin', 1)
    # con un segundo superadmin sí se puede quitar el primero
    post(sup, '/admin/usuarios/ana/rol', '/admin/usuarios', rol='superadmin')
    post(sup, '/admin/usuarios/sup/rol', '/admin/usuarios', rol='ingresos')
    assert db(app).execute("select rol from usuarios where usuario='sup'").fetchone()[0] == 'ingresos'
    ana = login(app, 'ana'); post(ana, '/admin/usuarios/ana/activo', '/admin/usuarios', activo='0')   # ahora ana es la última
    assert db(app).execute("select activo from usuarios where usuario='ana'").fetchone()[0] == 1


# ---------- auditoría ----------
def test_auditoria_solo_superadmin_con_filtros_y_paginas(app):
    for u in ('ana', 'ing'): assert login(app, u).get('/admin/auditoria').status_code == 403
    assert nuevo_cliente(app).get('/admin/auditoria').status_code == 303
    con = sqlite3.connect(app.state.almacen.dbp)
    con.executemany('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                    [(f'2026-01-{(i % 28) + 1:02d} 10:00:00', 'relleno' if i % 2 else 'otro', 'accion_x' if i % 2 else 'accion_y', '{}') for i in range(120)])
    con.commit(); con.close()
    sup = login(app, 'sup'); r = sup.get('/admin/auditoria'); assert r.status_code == 200 and 'Página 1 de' in r.text
    assert r.text.count('<code>') == 50
    p3 = sup.get('/admin/auditoria?pagina=3'); assert 'Página 3 de' in p3.text and p3.text.count('<code>') <= 50
    assert sup.get('/admin/auditoria?pagina=999').status_code == 200 and sup.get('/admin/auditoria?pagina=abc').status_code == 200
    f = sup.get('/admin/auditoria?usuario=RELLENO&accion=accion_x'); assert '60 registros' in f.text and 'accion_y' not in f.text.split('<tbody>')[1]
    n = sup.get('/admin/auditoria?desde=2026-01-05&hasta=2026-01-05').text
    assert re.search(r'(\d+) registros', n) and int(re.search(r'(\d+) registros?', n).group(1)) in (4, 5)
    assert 'No hay registros' in sup.get('/admin/auditoria?usuario=nadie-existe').text


# ---------- mi cuenta ----------
def test_cambio_de_clave_propio(app):
    c = login(app, 'ana'); assert c.get('/cuenta').status_code == 200
    def cambiar(actual, nueva, nueva2): return post(c, '/cuenta/clave', '/cuenta', actual=actual, nueva=nueva, nueva2=nueva2)
    assert 'no es correcta' in cambiar('mala', 'NuevaClave-1', 'NuevaClave-1').text
    assert 'no coincide' in cambiar(CLAVES['ana'], 'NuevaClave-1', 'otra').text
    assert 'al menos 8' in cambiar(CLAVES['ana'], 'corta', 'corta').text
    assert 'distinta' in cambiar(CLAVES['ana'], CLAVES['ana'], CLAVES['ana']).text
    assert cambiar(CLAVES['ana'], 'NuevaClave-1', 'NuevaClave-1').status_code == 303
    assert entrar(app, 'ana', CLAVES['ana'])[1].status_code == 401 and entrar(app, 'ana', 'NuevaClave-1')[1].status_code == 303
    assert nuevo_cliente(app).get('/cuenta').status_code == 303
    assert c.post('/cuenta/clave', data={'actual': 'a', 'nueva': 'b', 'nueva2': 'b'}).status_code == 403    # sin CSRF
    todo = ' '.join(r[0] or '' for r in db(app).execute('select detalle_json from auditoria'))
    assert 'NuevaClave-1' not in todo and db(app).execute("select count(*) from auditoria where accion='clave_cambiada'").fetchone()[0] == 1


def test_menu_segun_rol(app):
    menu = lambda u: login(app, u).get('/').text.split('class="nav"')[1].split('</nav>')[0]
    assert 'Usuarios' in menu('sup') and 'Auditoría' in menu('sup') and 'Nuevo cierre' in menu('sup')
    assert 'Usuarios' not in menu('ing') and 'Nuevo cierre' not in menu('ing')
    assert 'Nuevo cierre' in menu('ana') and 'Auditoría' not in menu('ana') and 'Mi cuenta' in menu('ana')


# ---------- migración v2 -> v3 ----------
ESQUEMA_V2_EXTRA = """
ALTER TABLE cierres ADD COLUMN ci_total REAL; ALTER TABLE cierres ADD COLUMN revisado_por TEXT;
ALTER TABLE cierres ADD COLUMN revisado_en TEXT; ALTER TABLE cierres ADD COLUMN observacion_revision TEXT;
CREATE TABLE cierre_ci (cierre_id INTEGER NOT NULL REFERENCES cierres(id), factura TEXT, banco TEXT NOT NULL, codigo TEXT NOT NULL, monto REAL);
CREATE INDEX ix_ci_bc ON cierre_ci(banco, codigo);
PRAGMA user_version = 2;
"""


def crear_db_v2(ruta):
    con = sqlite3.connect(ruta); con.executescript(ESQUEMA_V1 + ESQUEMA_V2_EXTRA)
    base = ("INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, total_cierre, "
            "estado_revision, version, reemplaza_a, presentado_en, ruta, revisado_por, revisado_en, observacion_revision) VALUES ")
    con.execute(base + "('CC-2026-00001','2026-10-15','ana','Ana Prueba','SFC1',1000,1000,100,100,1100,'REEMPLAZADO',1,NULL,'2026-10-15 17:00:00','cierres/x/v1','ing','2026-10-15 17:30:00','Falta voucher del POS')")
    con.execute(base + "('CC-2026-00002','2026-10-15','ana','Ana Prueba','SFC1',1200,1200,100,100,1300,'CONFORME',2,1,'2026-10-15 18:00:00','cierres/x/v2','ing','2026-10-15 18:30:00','Conforme')")
    con.execute(base + "('CC-2026-00003','2026-10-16','beto','Beto Prueba','SFC2',500,500,0,0,500,'PENDIENTE DE REVISIÓN',1,NULL,'2026-10-16 18:00:00','cierres/x/v3',NULL,NULL,NULL)")
    con.execute("INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES ('2026-10-15 17:00:00','ana','presentar_cierre','{}')")
    con.execute("INSERT INTO cierre_movimientos VALUES (2,'700000001','30001',500.0)")
    con.commit(); con.close()


def test_migracion_v2_a_v3_idempotente_y_siembra(tmp_path, monkeypatch):
    data = tmp_path / 'data'; data.mkdir(); ruta = str(data / 'cierres.db'); crear_db_v2(ruta)
    assert sqlite3.connect(ruta).execute('PRAGMA user_version').fetchone()[0] == 2 and dbm.VERSION_ESQUEMA == 4
    monkeypatch.setenv('DATA_DIR', str(data)); monkeypatch.setenv('APP_USERS', json.dumps(USUARIOS))
    monkeypatch.setenv('SECRET_KEY', 'k'); monkeypatch.setenv('COOKIE_SECURE', '0')
    from app.main import create_app
    for _ in range(2):      # arrancar dos veces: no falla, no duplica, no pierde nada
        app = create_app()
        con = sqlite3.connect(ruta)
        assert con.execute('PRAGMA user_version').fetchone()[0] == 4
        assert {r[0] for r in con.execute("select name from sqlite_master where type='table'")} >= {'usuarios', 'revisiones_historial', 'cierres', 'auditoria', 'cierre_ci', 'alertas_cierre'}
        assert con.execute('select correlativo, estado_revision, total_cierre, revisado_por from cierres order by id').fetchall() == [
            ('CC-2026-00001', 'REEMPLAZADO', 1100, 'ing'), ('CC-2026-00002', 'CONFORME', 1300, 'ing'), ('CC-2026-00003', 'PENDIENTE DE REVISIÓN', 500, None)]
        assert con.execute("select count(*) from auditoria where accion='presentar_cierre'").fetchone()[0] == 1
        assert con.execute('select count(*) from cierre_movimientos').fetchone()[0] == 1
        assert [tuple(r) for r in con.execute('select usuario, rol, activo from usuarios order by usuario')] == [
            ('ana', 'cajero', 1), ('beto', 'cajero', 1), ('ing', 'ingresos', 1), ('sup', 'superadmin', 1)]
        # las revisiones v2 pasaron al historial una sola vez
        assert con.execute('select cierre_id, usuario, accion, estado_nuevo, observacion from revisiones_historial order by cierre_id').fetchall() == [
            (1, 'ing', 'rechazar', 'RECHAZADO', 'Falta voucher del POS'), (2, 'ing', 'aprobar', 'CONFORME', 'Conforme')]
        con.close()
    # los usuarios sembrados pueden entrar y el superadmin puede revertir un CONFORME migrado
    sup = login(app, 'sup'); os.makedirs(data / 'cierres' / 'x' / 'v2')
    assert sup.get('/cierres/2').status_code == 200 and 'Historial de revisiones' in sup.get('/cierres/2').text
    post(sup, '/cierres/2/revertir', '/cierres/2', motivo='Prueba de reversión sobre dato migrado')
    assert sqlite3.connect(ruta).execute('select estado_revision, revisado_por from cierres where id=2').fetchone() == ('PENDIENTE DE REVISIÓN', None)
