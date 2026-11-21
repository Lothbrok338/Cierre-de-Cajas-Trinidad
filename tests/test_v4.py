"""Pruebas v4: formato del código de asignación por banco como alerta de revisión NO bloqueante."""
import json, os, re, sqlite3
from datetime import datetime
import pytest
from tests.conftest import login, post, USUARIOS
from tests import sinteticos as S
from tests.test_app import db, pag, presentar
from tests.test_v2 import borrador_nuevo, respaldo, flash_de, leer_estado, sint_extra, sint_manual   # noqa: F401  (fixtures)
from app import catalogo, db as dbm

ANIO = datetime.now().year


def antes_kpis(html): return html.split('<section class="kpis">')[0]


# ---------- regla por banco (unidad) ----------
@pytest.mark.parametrize('banco,validos,invalidos', [
    ('BNB', ['P123', 'p123', 'AB12P9', 'abcp', '12p', 'PPP'], ['ABC123', '123', 'P-123', 'P.123', 'P 12', 'PÑ1', 'P12_']),
    ('BMSC', ['T123', 't123', 'AB12T9', 'abct', '12t', 'TTT'], ['ABC123', 'P123', 'T-123', 'T/123', 'T 12']),
    ('BISA', ['123', '0001234567', '000'], ['12A', 'P123', '12-3', '1.23', '12 3', '١٢٣']),
    ('UNION', ['123', '9999999'], ['U123', '12-3', '1/23']),
    ('QR BANECO', ['700000001', '123'], ['70000000A', '7000-1', 'QR123']),
])
def test_formato_por_banco(banco, validos, invalidos):
    for c in validos: assert catalogo.formato_incumplido(banco, c) is None, (banco, c)
    for c in invalidos: assert catalogo.formato_incumplido(banco, c) == catalogo.regla_formato(banco), (banco, c)
    assert catalogo.regla_formato('BNB') == 'BNB: letras y números, debe incluir una P'
    assert catalogo.regla_formato('BMSC') == 'BMSC: letras y números, debe incluir una T'
    if banco != 'QR BANECO':     # el formato nunca bloquea: los códigos con símbolos se aceptan
        for c in invalidos: assert catalogo.validar_codigo(banco, c)[0], (banco, c)


# ---------- cajero ----------
def test_cajero_ve_alerta_y_no_se_bloquea(app, sint_extra):
    c, bid = borrador_nuevo(app, sint_extra)
    assert respaldo(c, bid, '30003', 'BNB', 'abc-123').status_code == 303
    assert leer_estado(app, bid)['reclasificaciones'] == {'30003': {'banco': 'BNB', 'codigo': 'ABC-123'}}      # se registró igual
    msg = flash_de(c, bid)
    assert 'FORMATO A REVISAR' in msg and 'BNB: letras y números, debe incluir una P' in msg
    html = pag(c, bid).text
    assert 'Alertas de revisión (1)' in html and 'class="alerta-fmt"' in html                  # banner + junto a la factura
    assert 'FORMATO A REVISAR: factura 30003, BNB, código ABC-123' in html                    # en la sección Alertas existente
    assert db(app).execute("select count(*) from auditoria where accion='alerta_formato'").fetchone()[0] == 1
    respaldo(c, bid, '30006', 'BMSC', 't-77'); assert 'Alertas de revisión (2)' in pag(c, bid).text
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30006'); respaldo(c, bid, '30006', 'BMSC', 'a7t')   # corregido: ya sin alerta
    assert 'respaldo manual registrado (BMSC A7T)' in flash_de(c, bid) and 'FORMATO A REVISAR' not in flash_de(c, bid)
    html = pag(c, bid).text
    assert 'Alertas de revisión (1)' in html and 'A7T' not in re.search(r'<section class="alertas-rev".*?</section>', html, re.S).group(0)
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30003')      # deshacer quita la alerta
    assert 'alertas-rev' not in antes_kpis(pag(c, bid).text)


def test_codigos_validos_sin_alerta(app, sint_extra):
    c, bid = borrador_nuevo(app, sint_extra)
    respaldo(c, bid, '30003', 'BNB', 'xp1'); respaldo(c, bid, '30006', 'BMSC', ' 5t5 ')
    html = pag(c, bid).text
    assert 'alertas-rev' not in html and 'FORMATO A REVISAR' not in html
    assert leer_estado(app, bid)['reclasificaciones'] == {'30003': {'banco': 'BNB', 'codigo': 'XP1'}, '30006': {'banco': 'BMSC', 'codigo': '5T5'}}
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/')
    cid = int(r.headers['location'].rsplit('/', 1)[1])
    assert db(app).execute('select count(*) from alertas_cierre').fetchone()[0] == 0
    ing = login(app, 'ing'); assert 'alertas-revision' not in ing.get(f'/cierres/{cid}').text and 'ins-alerta' not in ing.get('/').text


def test_aviso_en_dialogo_de_presentar(app, sint_extra):
    c, bid = borrador_nuevo(app, sint_extra)
    respaldo(c, bid, '30003', 'BISA', 'x12'); respaldo(c, bid, '30006', 'UNION', '5555')
    assert 'Hay 1 código con formato a revisar; Ingresos será notificado' in pag(c, bid).text
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='30006'); respaldo(c, bid, '30006', 'UNION', '55-5')
    html = pag(c, bid).text
    assert 'Hay 2 códigos con formato a revisar; Ingresos será notificado' in html and 'data-abrir="dlg-presentar"' in html     # sigue siendo presentable


def test_qr_manual_formato(app, tmp_path, sint_manual):
    c, bid = borrador_nuevo(app, sint_manual)      # ID numérico: sin alerta
    respaldo(c, bid, '30002', 'QR BANECO', '700000009')
    assert leer_estado(app, bid)['respaldos_manuales'] == ['30002'] and 'alertas-rev' not in antes_kpis(pag(c, bid).text)
    d = tmp_path / 'qrx'; d.mkdir()      # ID no numérico en el reporte: se registra pero alerta
    banco = [b for b in S.BANCO if b[0] != '700000002'] + [('ZQ9X', 'ZZZ XXX WWW', '10:29:00', 300)]
    s2 = {'a1': S.crear_a1(str(d / 'm.xlsx')), 'a2': S.crear_a2(str(d / 'b.xlsx'), banco)}
    c2, b2 = borrador_nuevo(app, s2, 'beto')
    respaldo(c2, b2, '30002', 'QR BANECO', 'ZQ9X')
    assert leer_estado(app, b2)['respaldos_manuales'] == ['30002']
    html = pag(c2, b2).text
    assert 'Alertas de revisión (1)' in html and 'QR BANECO: solo números' in html and 'class="alerta-fmt"' in html
    # una confirmación o asignación automática (ID del banco, no escrito por el cajero) no genera alerta
    c3, b3 = borrador_nuevo(app, s2, 'ana'); post(c3, f'/borrador/{b3}/confirmar', f'/borrador/{b3}', factura='30003')
    assert 'alertas-rev' not in antes_kpis(pag(c3, b3).text)


# ---------- presentar: se guardan; ingresos / superadmin las ven ----------
def presentado_con_alertas(app, sint_extra, usuario='ana'):
    c, bid = borrador_nuevo(app, sint_extra, usuario)
    respaldo(c, bid, '30003', 'BNB', 'DEP-12345'); respaldo(c, bid, '30006', 'BISA', '12-34')     # dos alertas
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    return c, bid, int(r.headers['location'].rsplit('/', 1)[1])


def test_alertas_se_guardan_al_presentar(app, sint_extra):
    c, bid, cid = presentado_con_alertas(app, sint_extra)
    filas = db(app).execute('select cierre_id, tipo, factura, banco, codigo, detalle, creado_en, revisada_por, revisada_en from alertas_cierre order by id').fetchall()
    assert [f[:6] for f in filas] == [(cid, 'FORMATO', '30003', 'BNB', 'DEP-12345', 'BNB: letras y números, debe incluir una P'),
                                      (cid, 'FORMATO', '30006', 'BISA', '12-34', 'BISA: solo números')]
    assert all(f[6] and f[7] is None and f[8] is None for f in filas)
    ruta = db(app).execute('select ruta from cierres where id=?', (cid,)).fetchone()[0]
    cj = json.load(open(os.path.join(app.state.almacen.dir, ruta, 'cierre.json'), encoding='utf-8'))
    assert [(a['factura'], a['banco'], a['codigo']) for a in cj['alertas_formato']] == [('30003', 'BNB', 'DEP-12345'), ('30006', 'BISA', '12-34')]
    aud = [json.loads(a[0]) for a in db(app).execute("select detalle_json from auditoria where accion='presentar_cierre'")]
    assert aud[-1]['alertas_formato'] == 2


def test_lista_insignia_filtro_y_detalle(app, sint_extra, tmp_path):
    c, bid, cid = presentado_con_alertas(app, sint_extra)
    # cierre limpio de otro cajero, con IDs de banco distintos para no chocar con el anterior
    d = tmp_path / 'otro'; d.mkdir()
    banco = [('8' + b[0][1:], *b[1:]) for b in S.BANCO]
    otro = {'a1': S.crear_a1(str(d / 'm.xlsx')), 'a2': S.crear_a2(str(d / 'b.xlsx'), banco)}
    cb, bb = borrador_nuevo(app, otro, 'beto'); post(cb, f'/borrador/{bb}/confirmar', f'/borrador/{bb}', factura='30003')
    assert presentar(cb, bb).status_code == 303
    for rol in ('ing', 'sup'):
        u = login(app, rol); h = u.get('/').text
        assert '⚠ 2 alertas' in h and h.count('ins-alerta') == 1 and 'Con alertas' in h and h.count(f'CC-{ANIO}-0000') == 2
        f = u.get('/?alertas=1').text
        assert f'CC-{ANIO}-00001' in f and f'CC-{ANIO}-00002' not in f and 'value="1" selected' in f
        d_ = u.get(f'/cierres/{cid}').text
        assert 'Alertas de revisión (2)' in d_ and 'BNB: letras y números, debe incluir una P' in d_ and 'BISA: solo números' in d_
        assert 'DEP-12345' in d_ and '12-34' in d_ and d_.count('Marcar como revisada') == 2
    # el cajero dueño ve el bloque (sin botón) y su lista sin insignia; otro cajero no ve el cierre
    d_ = c.get(f'/cierres/{cid}').text
    assert 'Alertas de revisión (2)' in d_ and 'Marcar como revisada' not in d_ and 'Por revisar' in d_ and 'ins-alerta' not in c.get('/').text
    assert cb.get(f'/cierres/{cid}').status_code == 404
    assert 'Con alertas' not in c.get('/').text


def test_marcar_revisada_con_auditoria(app, sint_extra):
    c, bid, cid = presentado_con_alertas(app, sint_extra)
    ing = login(app, 'ing')
    assert 'Hay 2 alertas sin revisar' in ing.get(f'/cierres/{cid}').text      # advertencia en el diálogo de aprobación
    aid, aid2 = [r[0] for r in db(app).execute('select id from alertas_cierre order by id')]
    assert post(c, f'/cierres/{cid}/alertas/{aid}/revisar', f'/cierres/{cid}').status_code == 403      # el cajero no puede
    assert post(ing, f'/cierres/{cid}/alertas/{aid}/revisar', f'/cierres/{cid}').status_code == 303
    assert db(app).execute('select revisada_por, revisada_en from alertas_cierre where id=?', (aid,)).fetchone()[0] == 'ing'
    assert db(app).execute('select revisada_en from alertas_cierre where id=?', (aid,)).fetchone()[0]
    a = [json.loads(x[0]) for x in db(app).execute("select detalle_json from auditoria where accion='alerta_revisada'")]
    assert len(a) == 1 and a[0]['factura'] == '30003' and a[0]['codigo'] == 'DEP-12345' and a[0]['cierre'] == cid and a[0]['alerta'] == aid
    h = ing.get('/').text; d = ing.get(f'/cierres/{cid}').text
    assert '⚠ 1 alerta' in h and 'Hay 1 alerta sin revisar' in d and 'Revisada' in d and d.count('Marcar como revisada') == 1
    post(ing, f'/cierres/{cid}/alertas/{aid}/revisar', f'/cierres/{cid}')      # repetir no duplica ni pisa
    post(ing, f'/cierres/{cid}/alertas/9999/revisar', f'/cierres/{cid}')
    assert db(app).execute("select count(*) from auditoria where accion='alerta_revisada'").fetchone()[0] == 1
    sup = login(app, 'sup'); post(sup, f'/cierres/{cid}/alertas/{aid2}/revisar', f'/cierres/{cid}')
    assert db(app).execute('select revisada_por from alertas_cierre where id=?', (aid2,)).fetchone()[0] == 'sup'
    h = sup.get('/').text; assert '✓ 2 revisadas' in h and 'ins-alerta' not in h
    assert 'Hay ' not in re.search(r'<dialog id="dlg-aprobar">.*?</dialog>', sup.get(f'/cierres/{cid}').text, re.S).group(0).replace('Hay 0', '')


def test_aprobar_permitido_con_alertas_sin_revisar(app, sint_extra):
    c, bid, cid = presentado_con_alertas(app, sint_extra)
    ing = login(app, 'ing')
    assert post(ing, f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='aprobar', observacion='').status_code == 303
    assert db(app).execute('select estado_revision from cierres where id=?', (cid,)).fetchone()[0] == 'CONFORME'
    assert db(app).execute('select count(*) from alertas_cierre where revisada_en is null').fetchone()[0] == 2


def test_pdf_marca_y_lista_alertas(app, sint_extra):
    pypdf = pytest.importorskip('pypdf')
    c, bid, cid = presentado_con_alertas(app, sint_extra)
    ruta = db(app).execute('select ruta from cierres where id=?', (cid,)).fetchone()[0]
    leer = lambda n: ' '.join(p.extract_text() for p in pypdf.PdfReader(os.path.join(app.state.almacen.dir, ruta, n)).pages)
    txt = leer('cierre.pdf')
    assert txt.count('⚠') >= 2 and 'Alertas de revisión' in txt and 'BNB: letras y números, debe incluir una P' in txt and 'DEP-12345' in txt and '12-34' in txt
    post(login(app, 'ing'), f'/cierres/{cid}/revisar', f'/cierres/{cid}', decision='aprobar', observacion='')     # el PDF con estado de revisión las conserva
    txt = leer('cierre_revision.pdf'); assert '⚠' in txt and 'BISA: solo números' in txt


def test_pdf_sin_alertas_no_las_menciona(app, sint):
    pypdf = pytest.importorskip('pypdf')
    c, bid = borrador_nuevo(app, sint); post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    cid = int(presentar(c, bid).headers['location'].rsplit('/', 1)[1])
    ruta = db(app).execute('select ruta from cierres where id=?', (cid,)).fetchone()[0]
    txt = ' '.join(p.extract_text() for p in pypdf.PdfReader(os.path.join(app.state.almacen.dir, ruta, 'cierre.pdf')).pages)
    assert '⚠' not in txt and 'Alertas de revisión' not in txt


# ---------- migración v3 -> v4 ----------
def test_migracion_v3_a_v4_con_filas(tmp_path, monkeypatch):
    data = tmp_path / 'data'; data.mkdir(); ruta = str(data / 'cierres.db')
    dbm.inicializar(ruta)      # se deja como una base v3 (sin la tabla nueva) con filas
    con = sqlite3.connect(ruta)
    con.execute("INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, total_cierre, "
                "estado_revision, version, presentado_en, ruta) VALUES ('CC-2026-00001','2026-10-15','ana','Ana Prueba','SFC1',1000,1000,100,100,1100,'CONFORME',1,'2026-10-15 17:00:00','cierres/x/v1')")
    con.execute("INSERT INTO cierre_ci VALUES (1,'30003','BNB','ABC-1',700)")
    con.execute("INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES ('2026-10-15 17:00:00','ana','presentar_cierre','{}')")
    con.execute('DROP TABLE alertas_cierre'); con.execute('PRAGMA user_version = 3'); con.commit(); con.close()
    monkeypatch.setenv('DATA_DIR', str(data)); monkeypatch.setenv('APP_USERS', json.dumps(USUARIOS))
    monkeypatch.setenv('SECRET_KEY', 'k'); monkeypatch.setenv('COOKIE_SECURE', '0')
    from app.main import create_app
    for _ in range(2):      # idempotente
        app = create_app(); con = sqlite3.connect(ruta)
        assert con.execute('PRAGMA user_version').fetchone()[0] == 4
        cols = [r[1] for r in con.execute('PRAGMA table_info(alertas_cierre)')]
        assert cols[:8] == ['id', 'cierre_id', 'tipo', 'factura', 'banco', 'codigo', 'detalle', 'creado_en'] and 'revisada_por' in cols and 'revisada_en' in cols
        assert con.execute('select correlativo, estado_revision from cierres').fetchall() == [('CC-2026-00001', 'CONFORME')]
        assert con.execute('select count(*) from cierre_ci').fetchone()[0] == 1 and con.execute('select count(*) from alertas_cierre').fetchone()[0] == 0
        con.close()
    ing = login(app, 'ing'); os.makedirs(data / 'cierres' / 'x' / 'v1')      # los cierres antiguos no se revalidan
    h = ing.get('/').text; assert 'CC-2026-00001' in h and 'ins-alerta' not in h
    assert 'CC-2026-00001' not in ing.get('/?alertas=1').text and ing.get('/cierres/1').status_code == 200
