"""Copiar para el extracto / para liberar: contenido TSV y permisos."""
import glob, json, os, re
import pytest
from tests.conftest import login, post
from tests.test_app import bid_de, db, presentar, subir, DIR_REAL
from tests.test_v2 import borrador_nuevo, respaldo, flash_de
from tests.test_v3 import cierre_presentado
from app import servicio


def tsv_de(html, ident):
    m = re.search(rf'<script type="application/json" id="{ident}">(.*?)</script>', html, re.S)
    return json.loads(m.group(1).replace('\\u0026', '&')) if m else None


def filas(t): return [l.split('\t') for l in t.split('\n')] if t else []


def test_tsv_funcion_pura():
    movs = [{'id_transaccion': '111', 'factura': '1', 'monto': 1592.0}, {'id_transaccion': '222', 'factura': '2', 'monto': 1234567.5}]
    t, n, tot = servicio.tsv_extracto('CC-2026-00007', '2026-10-16', movs)
    assert t == '\t1592.00\t111\t1\tCC-2026-00007\t16/10/2026\n\t1234567.50\t222\t2\tCC-2026-00007\t16/10/2026'
    assert (n, tot) == (2, 1236159.5)
    assert servicio.tsv_extracto('C', '2026-01-02', movs, 'LIBERAR')[0].startswith('LIBERAR\t1592.00\t111\t1\tC\t02/01/2026')
    assert servicio.tsv_extracto('C', '2026-01-02', []) == ('', 0, 0)


def test_boton_y_contenido_ingresos_y_super(app, sint):
    c, cid = cierre_presentado(app, sint)   # 30001 (500) y 30002 (300) conciliadas; 30003 confirmada; tarjeta aparte
    ing = login(app, 'ing'); h = ing.get(f'/cierres/{cid}').text
    assert 'Copiar para el extracto' in h and 'Copiar para liberar' not in h and 'datos-liberar' not in h
    f = filas(tsv_de(h, 'datos-marcar'))
    assert [x[2] for x in f] == ['700000001', '700000002', '700000003']
    assert [x[3] for x in f] == ['30001', '30002', '30003'] and [x[1] for x in f] == ['500.00', '300.00', '700.00']
    corr = db(app).execute('select correlativo from cierres').fetchone()[0]
    assert all(x[0] == '' and x[4] == corr and x[5] == '16/10/2026' and len(x) == 6 for x in f)
    assert '<strong>3</strong> facturas QR' in h and '1,500.00' in h
    sup = login(app, 'sup'); hs = sup.get(f'/cierres/{cid}').text
    assert 'Copiar para liberar' in hs
    assert all(x[0] == 'LIBERAR' for x in filas(tsv_de(hs, 'datos-liberar')))
    assert tsv_de(hs, 'datos-marcar') == tsv_de(h, 'datos-marcar')


def test_cajero_no_ve_botones(app, sint):
    c, cid = cierre_presentado(app, sint)
    assert 'Copiar para el extracto' not in c.get(f'/cierres/{cid}').text


def test_respaldo_manual_incluido_y_ci_excluida(app, sint_manual_ci):
    c, bid = borrador_nuevo(app, sint_manual_ci)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    assert respaldo(c, bid, '30002', 'QR BANECO', '700000009').status_code == 303
    r = presentar(c, bid); assert r.status_code == 303, flash_de(c, bid)
    cid = int(r.headers['location'].rsplit('/', 1)[1])
    f = filas(tsv_de(login(app, 'ing').get(f'/cierres/{cid}').text, 'datos-marcar'))
    assert sorted(x[2] for x in f) == ['700000001', '700000003', '700000009']


@pytest.fixture
def sint_manual_ci(tmp_path):
    from tests import sinteticos as S
    d = tmp_path / 'sm'; d.mkdir()
    banco = [b for b in S.BANCO if b[0] != '700000002'] + [('700000009', 'ZZZ XXX WWW', '10:29:00', 300)]
    return {'a1': S.crear_a1(str(d / 'm.xlsx')), 'a2': S.crear_a2(str(d / 'b.xlsx'), banco)}


def test_comunicacion_interna_no_va(app, sint):
    c, bid = borrador_nuevo(app, sint)
    respaldo(c, bid, '30003', 'BNB', 'DEP-12345')
    r = presentar(c, bid); assert r.status_code == 303, flash_de(c, bid)
    cid = int(r.headers['location'].rsplit('/', 1)[1])
    f = filas(tsv_de(login(app, 'ing').get(f'/cierres/{cid}').text, 'datos-marcar'))
    assert [x[3] for x in f] == ['30001', '30002'] and all('DEP-12345' not in ','.join(x) for x in f)


real = pytest.mark.skipif(not DIR_REAL or not glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls')), reason='DATOS_PRUEBA_DIR sin datos reales')


@real
def test_tsv_datos_reales(app):
    a1 = glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls'))[0]; a2 = glob.glob(os.path.join(DIR_REAL, '*Reporte_pagos*.xlsx'))[0]
    c = login(app, 'ana'); r = subir(c, a1, a2, '1191'); assert r.status_code == 303, r.text
    bid = bid_de(r)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='20048')
    r = presentar(c, bid); assert r.status_code == 303
    cid = int(r.headers['location'].rsplit('/', 1)[1])
    row = db(app).execute('select correlativo, qr_conciliado, fecha_cierre from cierres where id=?', (cid,)).fetchone()
    h = login(app, 'sup').get(f'/cierres/{cid}').text
    f = filas(tsv_de(h, 'datos-marcar'))
    assert len(f) == 14 and all(len(x) == 6 and x[0] == '' and x[4] == row[0] for x in f)
    assert abs(sum(float(x[1]) for x in f) - row[1]) < 0.005
    assert len({x[2] for x in f}) == 14 and all(re.fullmatch(r'\d{6,}', x[2]) for x in f)
    assert all(re.fullmatch(r'\d+\.\d\d', x[1]) for x in f) and {x[5] for x in f} == {'16/10/2026'}
    assert len(filas(tsv_de(h, 'datos-liberar'))) == 14
