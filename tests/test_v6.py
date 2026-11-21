"""Pruebas v6: alertas INFO ocultas, textos de presentar, confirmar todas, nombres de archivo legibles, nombre completo del revisor."""
import json, os, re, shutil
import pytest
from tests.conftest import login, post, subir
from tests import sinteticos as S
from tests.test_app import bid_de, db, pag, presentar
from tests.test_v2 import borrador_nuevo, flash_de, leer_estado
from tests.test_v3 import cierre_presentado, revisar
from tests.test_v5 import carpeta, texto_pdf
from app import motor

TEXTO_PRESENTAR = ('Al presentar se genera su comprobante en PDF para imprimir o guardar. El cierre no podrá modificarse mientras esté pendiente de revisión; '
                   'si Ingresos lo rechaza, podrá corregirlo y reenviarlo.')


@pytest.fixture
def sint_dudosas(tmp_path):
    """Tres facturas POR CONFIRMAR (titular parecido pero distinto), una SIN RESPALDO y una conciliada. Nombres ficticios."""
    d = tmp_path / 'sd'; d.mkdir()
    fac = [('40001', 'EFECTIVO', 500, 'LUCIA FERNANDEZ ROCA', 'FERNANDEZ ROCA LUCIA', 'Válido', '10:00:10'),
           ('40002', 'EFECTIVO', 700, 'EMPRESA ALFA SRL', 'SOTO VEGA CARLOS', 'Válido', '11:00:00'),
           ('40003', 'EFECTIVO', 410, 'OSCAR LUNA SOTO', 'LUNA SOTO OSCAR', 'Válido', '12:30:00'),
           ('40004', 'EFECTIVO', 125, 'DIEGO PAZ RIOS', 'PAZ RIOS DIEGO', 'Válido', '12:45:00'),
           ('40005', 'EFECTIVO', 90, 'NORMA VEGA LIMA', 'VEGA LIMA NORMA', 'Válido', '13:00:00')]
    banco = [('800000001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500), ('800000002', 'SOTO VEGA ANDREA', '10:58:50', 700),
             ('800000003', 'LUNA SOTO MARIA', '12:29:30', 410), ('800000004', 'PAZ RIOS SILVIA', '12:44:30', 125)]
    return {'a1': S.crear_a1(str(d / 'mov.xlsx'), fac), 'a2': S.crear_a2(str(d / 'banco.xlsx'), banco), 'dir': d}


@pytest.fixture
def sint_dup(tmp_path):
    """Reporte del banco con un número de transacción repetido: genera una alerta real de nivel ALERTA."""
    d = tmp_path / 'sdup'; d.mkdir()
    banco = S.BANCO + [('700000001', 'FERNANDEZ ROCA LUCIA', '09:59:40', 500)]
    return {'a1': S.crear_a1(str(d / 'mov.xlsx')), 'a2': S.crear_a2(str(d / 'banco.xlsx'), banco)}


# ---------- 1) alertas ----------
def test_alertas_info_no_se_muestran_ni_cuentan(app, sint):
    res = motor.conciliar(sint['a1'], sint['a2'], 200)
    assert (res['alertas']['nivel'] == 'INFO').any()               # la alerta de factura anulada sigue en el motor (cierre.json / hoja Alertas)
    c, bid = borrador_nuevo(app, sint); html = pag(c, bid).text
    assert '<summary>Alertas' not in html and 'Estos avisos no impiden presentar' not in html
    assert 'excluida' not in html.split('Facturas anuladas')[0].split('<section class="kpis">')[-1]
    assert 'Facturas anuladas <span class="cuenta">1</span>' in html     # la información sigue en su sección


def test_seccion_alertas_con_alerta_real(app, sint_dup):
    c, bid = borrador_nuevo(app, sint_dup); html = pag(c, bid).text
    assert re.search(r'<summary>Alertas <span class="cuenta">1</span>', html)
    assert 'Estos avisos no impiden presentar, pero Ingresos los verán' not in html and 'Estos avisos no impiden presentar, pero Ingresos los verá.' in html
    assert 'El número de transacción del banco 700000001 aparece repetido' in html and 'Revise que el banco no haya duplicado' in html
    assert 'ID TRANSACCION' not in re.search(r'<ul class="alertas">.*?</ul>', html, re.S).group(0)


def test_textos_de_alertas_en_lenguaje_simple(app, sint):
    res = motor.conciliar(sint['a1'], sint['a2'], 200, ['30001'], {'30002': 'NOEXISTE', '30003': '700000003'})
    textos = ' | '.join(res['alertas']['detalle'])
    assert 'Confirmación no aplicada: la factura 30001' in textos
    assert 'Elección no aplicada (factura 30002, número de transacción NOEXISTE)' in textos and 'ID TRANSACCION' not in textos
    assert 'está Anulado y no se cuenta en el cierre' in textos


# ---------- 2) texto de presentar ----------
def test_texto_de_presentar_en_zona_y_dialogo(app, sint):
    c, bid = borrador_nuevo(app, sint)
    assert pag(c, bid).text.count(TEXTO_PRESENTAR) == 1               # no presentable aún: solo el diálogo (oculto), sin el texto de la zona
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='30003')
    html = pag(c, bid).text
    assert html.count(TEXTO_PRESENTAR) == 2                           # zona "Presentar el cierre" y diálogo
    assert 'name="declaro"' in html and 'Declaro que el cierre es correcto' in html


# ---------- 3) confirmar todas ----------
def test_boton_solo_con_dos_o_mas_por_confirmar(app, sint, sint_dudosas):
    c, bid = borrador_nuevo(app, sint)                                # una sola POR CONFIRMAR
    assert 'Confirmar todas las coincidencias dudosas' not in pag(c, bid).text
    c, bid = borrador_nuevo(app, sint_dudosas, 'beto'); html = pag(c, bid).text
    assert 'Confirmar todas las coincidencias dudosas (3)' in html
    dlg = re.search(r'<dialog id="dlg-confirmar-todas">.*?</dialog>', html, re.S).group(0)
    for esperado in ('40002', 'EMPRESA ALFA SRL', 'SOTO VEGA ANDREA', '700.00', '40003', 'LUNA SOTO MARIA', '40004', 'PAZ RIOS SILVIA'):
        assert esperado in dlg, esperado
    assert '40005' not in dlg and '40001' not in dlg                  # SIN RESPALDO y conciliada no se ofrecen
    assert html.count('Sí, corresponde') == 3                         # la confirmación individual se mantiene


def test_confirmar_todas_confirma_audita_y_no_toca_otras(app, sint_dudosas):
    c, bid = borrador_nuevo(app, sint_dudosas)
    pend = {'40002', '40003', '40004'}
    r = post(c, f'/borrador/{bid}/confirmar-todas', f'/borrador/{bid}', facturas=['40001', '40002', '40003', '40004', '40005', 'NOEXISTE'])
    assert r.status_code == 303 and r.headers['location'] == f'/borrador/{bid}'
    est = leer_estado(app, bid)
    assert set(est['confirmaciones']) == pend and est['elecciones'] == {} and est['reclasificaciones'] == {} and est['respaldos_manuales'] == []
    assert 'Se confirmaron 3 coincidencias dudosas.' in flash_de(c, bid)
    con = db(app)
    aud = con.execute("select detalle_json from auditoria where accion='confirmar_coincidencia'").fetchall()
    assert len(aud) == 3 and {json.loads(a[0])['factura'] for a in aud} == pend and all(json.loads(a[0])['lote'] for a in aud)
    assert con.execute("select count(*) from auditoria where accion='confirmar_todas'").fetchone()[0] == 1
    html = pag(c, bid).text
    assert 'Confirmar todas las coincidencias dudosas' not in html
    assert 'Observadas (1)' in html and 'Conciliadas (4)' in html           # solo queda la SIN RESPALDO (40005)
    # repetir el envío no vuelve a confirmar nada
    r = post(c, f'/borrador/{bid}/confirmar-todas', f'/borrador/{bid}', facturas=['40002'])
    assert 'No hay coincidencias dudosas' in flash_de(c, bid) and set(leer_estado(app, bid)['confirmaciones']) == pend
    assert con.execute("select count(*) from auditoria where accion='confirmar_coincidencia'").fetchone()[0] == 3


def test_confirmar_todas_ajeno_y_sin_sesion(app, sint_dudosas):
    c, bid = borrador_nuevo(app, sint_dudosas)
    otro = login(app, 'beto')
    assert post(otro, f'/borrador/{bid}/confirmar-todas', '/nuevo', facturas=['40002', '40003']).status_code == 404
    assert leer_estado(app, bid)['confirmaciones'] == []


# ---------- 4) nombres de archivo legibles ----------
def test_nombre_de_archivo_sin_prefijo_hexadecimal(app, sint):
    d = sint['dir']; p1 = d / 'a1b2c3d4-Ingreso Diario.xlsx'; p2 = d / '0f9e8d7c-Reporte pagos.xlsx'
    shutil.copy(sint['a1'], p1); shutil.copy(sint['a2'], p2)
    c = login(app, 'ana'); r = subir(c, str(p1), str(p2)); assert r.status_code == 303
    bid = bid_de(r); html = pag(c, bid).text
    assert 'Archivos: Ingreso Diario.xlsx y Reporte pagos.xlsx' in html and 'a1b2c3d4' not in html
    assert leer_estado(app, bid)['nombres_subidos'] == {'movimiento': 'a1b2c3d4-Ingreso Diario.xlsx', 'banco': '0f9e8d7c-Reporte pagos.xlsx'}   # lo guardado no cambia
    # un nombre sin prefijo de 8 hex + guion no se toca
    sub = d / 'mov-20261016.xlsx'; shutil.copy(sint['a1'], sub)
    cb = login(app, 'beto'); r = subir(cb, str(sub), sint['a2']); assert r.status_code == 303
    assert 'Archivos: mov-20261016.xlsx y banco.xlsx' in pag(cb, bid_de(r)).text


# ---------- 5) nombre completo del revisor ----------
def test_nombre_completo_del_revisor_en_detalle_historial_y_pdf(app, sint):
    ana, cid = cierre_presentado(app, sint); ing = login(app, 'ing'); sup = login(app, 'sup')
    assert revisar(ing, cid, 'aprobar', 'Todo bien').status_code == 303
    assert db(app).execute('select revisado_por from cierres').fetchone()[0] == 'ing'             # lo guardado sigue siendo el login
    assert db(app).execute('select usuario from revisiones_historial').fetchone()[0] == 'ing'
    assert db(app).execute("select usuario from auditoria where accion='cierre_aprobado'").fetchone()[0] == 'ing'
    for cli in (ana, ing, sup):
        h = cli.get(f'/cierres/{cid}').text
        assert 'Aprobado por Ingresos Prueba' in h and 'Revisado por</dt><dd>Ingresos Prueba' in h and 'Aprobado por ing ' not in h
    hist = sup.get(f'/cierres/{cid}').text.split('Historial de revisiones')[1]
    assert '<td data-l="Usuario">Ingresos Prueba</td>' in hist
    txt = texto_pdf(os.path.join(carpeta(app, cid), 'cierre_revision.pdf'))
    assert 'Revisado por Ingresos Prueba' in txt and 'Revisado y conforme por Ingresos Prueba' in txt
    # si el usuario ya no existe, se muestra el usuario guardado
    con = db(app); con.execute("delete from usuarios where usuario='ing'"); con.commit()
    h = sup.get(f'/cierres/{cid}').text
    assert 'Aprobado por ing ' in h and 'Aprobado por Ingresos Prueba' not in h
    app.state.almacen.regenerar_pdf(cid)
    txt = texto_pdf(os.path.join(carpeta(app, cid), 'cierre_revision.pdf'))
    assert 'Revisado y conforme por ing el' in txt
