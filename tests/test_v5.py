"""Pruebas v5: comprobante PDF del cajero (banner de exito, botones, version vigente, 'hace N dias', pie del PDF, permisos)."""
import hashlib, os, re
from datetime import datetime, timedelta
import pytest
from tests.conftest import login, post
from tests.test_app import db
from tests.test_v2 import borrador_nuevo
from tests.test_v3 import cierre_presentado, revisar
from tests.test_app import presentar
from tests.test_v2 import flash_de

pypdf = pytest.importorskip('pypdf')


def texto_pdf(ruta):
    return ' '.join(p.extract_text() for p in pypdf.PdfReader(ruta).pages).replace('\n', ' ')


def carpeta(app, cid):
    return os.path.join(app.state.almacen.dir, db(app).execute('select ruta from cierres where id=?', (cid,)).fetchone()[0])


def presentar_y_ver(app, sint):
    c, bid = borrador_nuevo(app, sint)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    r = presentar(c, bid); assert r.status_code == 303, flash_de(c, bid)
    return c, r.headers['location']


def test_tras_presentar_banner_y_botones_una_sola_vez(app, sint):
    c, url = presentar_y_ver(app, sint)
    html = c.get(url).text; cid = url.rsplit('/', 1)[1]
    corr = db(app).execute('select correlativo from cierres').fetchone()[0]
    assert f'Cierre presentado — {corr}' in html and 'Guarde o imprima este comprobante como respaldo' in html
    assert 'algunos días' in html
    assert f'href="/cierres/{cid}/archivo/cierre.pdf"' in html and 'target="_blank"' in html
    assert 'Ver e imprimir comprobante (PDF)' in html and 'Descargar PDF' in html and f'cierre.pdf?descargar=1' in html
    assert 'Cierre presentado —' not in c.get(url).text            # el banner solo aparece tras presentar
    assert 'Ver e imprimir comprobante (PDF)' in c.get(url).text   # la tarjeta Comprobante queda siempre


def test_descarga_y_vista_inline(app, sint):
    c, url = presentar_y_ver(app, sint); cid = url.rsplit('/', 1)[1]
    r = c.get(f'/cierres/{cid}/archivo/cierre.pdf'); assert r.content[:5] == b'%PDF-' and r.headers['content-disposition'].startswith('inline')
    r = c.get(f'/cierres/{cid}/archivo/cierre.pdf?descargar=1'); assert r.headers['content-disposition'].startswith('attachment')
    r = c.get(f'/cierres/{cid}/comprobante'); assert r.status_code == 303 and r.headers['location'].endswith('/archivo/cierre.pdf')


def test_enlace_usa_revision_pdf_cuando_existe_y_cierre_pdf_inmutable(app, sint):
    c, cid = cierre_presentado(app, sint)
    d = carpeta(app, cid); h0 = hashlib.sha256(open(os.path.join(d, 'cierre.pdf'), 'rb').read()).hexdigest()
    ing = login(app, 'ing')
    assert revisar(ing, cid, 'aprobar', 'Todo bien').status_code == 303
    html = c.get(f'/cierres/{cid}').text
    assert f'/cierres/{cid}/archivo/cierre_revision.pdf"' in html and f'/archivo/cierre.pdf"' not in html.split('Archivos de respaldo')[0]
    assert 'Aprobado por Ingresos Prueba' in html
    assert c.get(f'/cierres/{cid}/comprobante').headers['location'].endswith('cierre_revision.pdf')
    assert hashlib.sha256(open(os.path.join(d, 'cierre.pdf'), 'rb').read()).hexdigest() == h0     # inmutable
    t = texto_pdf(os.path.join(d, 'cierre_revision.pdf'))
    assert 'Revisado y conforme por Ingresos Prueba' in t and 'CONFORME' in t and 'no constituye aprobación' not in t
    # revertir regenera con pie de pendiente; rechazar muestra la observacion
    sup = login(app, 'sup'); post(sup, f'/cierres/{cid}/revertir', f'/cierres/{cid}', motivo='Se aprobó por error aquí')
    assert 'no constituye aprobación hasta la revisión de Ingresos' in texto_pdf(os.path.join(d, 'cierre_revision.pdf'))
    revisar(ing, cid, 'rechazar', 'Falta el comprobante del depósito')
    assert 'Rechazado por Ingresos: Falta el comprobante del depósito' in texto_pdf(os.path.join(d, 'cierre_revision.pdf'))
    html = c.get(f'/cierres/{cid}').text
    assert 'Falta el comprobante del depósito' in html and 'Editar y reenviar' in html and html.count('Editar y reenviar') == 1
    assert hashlib.sha256(open(os.path.join(d, 'cierre.pdf'), 'rb').read()).hexdigest() == h0


def test_pdf_presentado_titulo_estado_y_pie_en_todas_las_paginas(app, sint):
    c, cid = cierre_presentado(app, sint)
    r = pypdf.PdfReader(os.path.join(carpeta(app, cid), 'cierre.pdf'))
    assert len(r.pages) >= 2
    for p in r.pages: assert 'no constituye aprobación hasta la revisión de Ingresos' in p.extract_text().replace('\n', ' ')
    p1 = r.pages[0].extract_text()
    assert 'Comprobante de cierre de caja' in p1 and 'PENDIENTE DE REVISIÓN' in p1


def test_hace_n_dias_en_detalle_y_lista(app, sint):
    c, cid = cierre_presentado(app, sint)
    assert 'presentado hoy' in c.get(f'/cierres/{cid}').text and '>hoy<' in c.get('/').text
    antes = (datetime.now() - timedelta(days=5)).strftime('%Y-%m-%d %H:%M:%S')
    con = db(app); con.execute('update cierres set presentado_en=? where id=?', (antes, cid)); con.commit()
    assert 'Esperando revisión de Ingresos — presentado hace 5 días' in c.get(f'/cierres/{cid}').text
    assert 'hace 5 días' in c.get('/').text
    login(app, 'ing').get('/')   # no falla para ingresos
    revisar(login(app, 'ing'), cid, 'aprobar')
    assert 'hace 5 días' not in c.get('/').text      # solo en PENDIENTE


def test_lista_boton_pdf_y_archivos_tecnicos_plegados(app, sint):
    c, cid = cierre_presentado(app, sint)
    h = c.get('/').text
    assert f'href="/cierres/{cid}/comprobante"' in h and 'target="_blank"' in h and '>PDF<' in h
    d = c.get(f'/cierres/{cid}').text
    assert '<details class="respaldo-tec"><summary>Archivos de respaldo (avanzado)</summary>' in d
    plegable = d.split('<details class="respaldo-tec">')[1].split('</details>')[0]
    for a in ('cierre.json', 'conciliacion.xlsx', 'movimiento_original.xlsx', 'banco_qr_original.xlsx'): assert a in plegable
    assert 'cierre.json' not in d.split('<details class="respaldo-tec">')[0]
    # ingresos ve la lista de archivos sin plegar y tambien el comprobante
    di = login(app, 'ing').get(f'/cierres/{cid}').text
    assert 'respaldo-tec' not in di and 'Ver e imprimir comprobante (PDF)' in di


def test_permisos_ajenos_404_e_ingresos_todos(app, sint):
    c, cid = cierre_presentado(app, sint)
    beto = login(app, 'beto')
    for ruta in (f'/cierres/{cid}', f'/cierres/{cid}/comprobante', f'/cierres/{cid}/archivo/cierre.pdf', f'/cierres/{cid}/archivo/cierre.pdf?descargar=1'):
        assert beto.get(ruta).status_code == 404, ruta
    for u in ('ing', 'sup'): assert login(app, u).get(f'/cierres/{cid}/archivo/cierre.pdf').content[:5] == b'%PDF-'
    assert f'/cierres/{cid}/comprobante' not in beto.get('/').text
