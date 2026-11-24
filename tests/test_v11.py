"""Pruebas v11: eliminar cierre (solo superadmin), auditoría y correlativo no reutilizado."""
import json, os, sqlite3
from tests.conftest import login, post
from tests.test_app import db, presentar
from tests.test_v2 import flash_de
from tests.test_v3 import cierre_presentado, revisar

MOT = 'Presentado por error, se repite'


def eliminar(c, cid, corr, motivo=MOT):
    return post(c, f'/cierres/{cid}/eliminar', f'/cierres/{cid}', motivo=motivo, correlativo=corr)


def fila(app, cid):
    con = db(app); con.row_factory = sqlite3.Row
    return con.execute('select * from cierres where id=?', (cid,)).fetchone()


def n(app, t, cid): return db(app).execute(f'select count(*) from {t} where cierre_id=?', (cid,)).fetchone()[0]


def test_solo_superadmin(app, sint):
    ana, cid = cierre_presentado(app, sint); corr = fila(app, cid)['correlativo']
    for u in ('ana', 'ing'):
        c = login(app, u)
        assert eliminar(c, cid, corr).status_code in (403, 404)
        assert 'Eliminar cierre' not in c.get(f'/cierres/{cid}').text
    assert fila(app, cid) is not None
    assert 'Eliminar cierre' in login(app, 'sup').get(f'/cierres/{cid}').text


def test_requiere_motivo_y_correlativo(app, sint):
    ana, cid = cierre_presentado(app, sint); corr = fila(app, cid)['correlativo']; sup = login(app, 'sup')
    eliminar(sup, cid, corr, 'corto'); eliminar(sup, cid, 'CC-2026-99999'); eliminar(sup, cid, '')
    assert fila(app, cid) is not None
    assert db(app).execute("select count(*) from auditoria where accion='eliminar_cierre'").fetchone()[0] == 0


def test_elimina_filas_archivos_audita_y_permite_repetir(app, sint):
    ana, cid = cierre_presentado(app, sint); sup = login(app, 'sup')
    f = fila(app, cid); corr = f['correlativo']; carpeta = os.path.join(os.environ['DATA_DIR'], f['ruta']); assert os.path.isdir(carpeta)
    assert n(app, 'cierre_movimientos', cid) > 0
    r = eliminar(sup, cid, corr); assert r.status_code == 303 and r.headers['location'] == '/'
    assert f'Cierre {corr} eliminado. Quedó registrado en auditoría.' in sup.get('/').text
    assert not os.path.exists(carpeta)
    for t in ('cierre_movimientos', 'cierre_ci', 'alertas_cierre', 'revisiones_historial'): assert n(app, t, cid) == 0
    assert db(app).execute('select count(*) from cierres').fetchone()[0] == 0
    a = json.loads(db(app).execute("select detalle_json from auditoria where accion='eliminar_cierre'").fetchone()[0])
    assert a['correlativo'] == corr and a['cajero'] == 'ana' and a['motivo'] == MOT and a['sha256_movimiento'] == f['sha256_movimiento']
    assert a['sha256_banco'] == f['sha256_banco'] and a['version'] == 1 and a['estado_revision'] == f['estado_revision'] and a['cierre'] == cid
    assert len(a['sha256_archivos']) >= 2
    ana2, cid2 = cierre_presentado(app, sint)   # mismos archivos/IDs: sin bloqueo anti-duplicado
    assert int(fila(app, cid2)['correlativo'].split('-')[-1]) == int(corr.split('-')[-1]) + 1


def test_carpeta_inexistente_igual_elimina(app, sint):
    import shutil
    ana, cid = cierre_presentado(app, sint); sup = login(app, 'sup'); f = fila(app, cid)
    shutil.rmtree(os.path.join(os.environ['DATA_DIR'], f['ruta']))
    assert eliminar(sup, cid, f['correlativo']).status_code == 303
    assert fila(app, cid) is None and n(app, 'cierre_movimientos', cid) == 0
    assert 'no existía' in sup.get('/').text


def test_ruta_fuera_de_data_no_se_borra(app, sint, tmp_path):
    ana, cid = cierre_presentado(app, sint); sup = login(app, 'sup'); corr = fila(app, cid)['correlativo']
    ajena = tmp_path / 'ajena'; ajena.mkdir(); (ajena / 'x.txt').write_text('x')
    w = sqlite3.connect(os.path.join(os.environ['DATA_DIR'], 'cierres.db')); w.execute('update cierres set ruta=? where id=?', (str(ajena), cid)); w.commit(); w.close()
    assert eliminar(sup, cid, corr).status_code == 303
    assert (ajena / 'x.txt').exists() and fila(app, cid) is None


def test_eliminar_reemplazante_no_restaura_ni_afecta_otros(app, sint):
    ana, c1 = cierre_presentado(app, sint, 'ana'); sup = login(app, 'sup'); ing = login(app, 'ing')
    revisar(ing, c1, 'rechazar', 'Falta el voucher del POS')
    r = post(ana, f'/cierres/{c1}/editar', f'/cierres/{c1}'); assert r.status_code == 303
    bid = r.headers['location'].rsplit('/', 1)[1]
    p = presentar(ana, bid); assert p.status_code == 303 and p.headers['location'].startswith('/cierres/'), flash_de(ana, bid)
    c2 = int(p.headers['location'].rsplit('/', 1)[1]); f1 = fila(app, c1); f2 = fila(app, c2); assert f2['version'] == 2
    assert eliminar(sup, c2, f2['correlativo']).status_code == 303
    assert fila(app, c2) is None and fila(app, c1)['estado_revision'] == 'REEMPLAZADO'
    base = os.environ['DATA_DIR']
    assert os.path.isdir(os.path.join(base, f1['ruta'])) and not os.path.exists(os.path.join(base, f2['ruta']))
    assert n(app, 'cierre_movimientos', c1) > 0
