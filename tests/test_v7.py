"""Pruebas v7: «No corresponde» (descartar la coincidencia sugerida de una factura POR CONFIRMAR). Datos ficticios; y con reales si hay DATOS_PRUEBA_DIR."""
import glob, json, os, re
import pytest
from tests.conftest import login, post
from tests import sinteticos as S
from tests.test_app import db, pag, presentar
from tests.test_v2 import borrador_nuevo, flash_de, leer_estado, respaldo
from tests.test_v3 import revisar
from tests.test_v5 import carpeta, texto_pdf
from app import motor, servicio

DIR_REAL = os.environ.get('DATOS_PRUEBA_DIR', '')
real = pytest.mark.skipif(not DIR_REAL or not glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls')), reason='DATOS_PRUEBA_DIR sin datos reales')


@pytest.fixture
def caso(tmp_path):
    """3 POR CONFIRMAR (40002, 40003, 40004), 1 conciliada (40001), 40005 conciliada. Hay un pago libre de 125 (800000014) para reasignar."""
    d = tmp_path / 'c7'; d.mkdir()
    fac = [('40001', 'EFECTIVO', 500, 'LUCIA FERNANDEZ ROCA', 'FERNANDEZ ROCA LUCIA', 'Válido', '10:00:10'),
           ('40002', 'EFECTIVO', 700, 'EMPRESA ALFA SRL', 'SOTO VEGA CARLOS', 'Válido', '11:00:00'),
           ('40003', 'EFECTIVO', 410, 'OSCAR LUNA SOTO', 'LUNA SOTO OSCAR', 'Válido', '12:30:00'),
           ('40004', 'EFECTIVO', 125, 'DIEGO PAZ RIOS', 'PAZ RIOS DIEGO', 'Válido', '12:45:00'),
           ('40005', 'EFECTIVO', 90, 'NORMA VEGA LIMA', 'VEGA LIMA NORMA', 'Válido', '13:00:00')]
    banco = [('800000001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500), ('800000002', 'SOTO VEGA ANDREA', '10:58:50', 700),
             ('800000003', 'LUNA SOTO MARIA', '12:29:30', 410), ('800000004', 'PAZ RIOS SILVIA', '12:44:30', 125),
             ('800000005', 'VEGA LIMA NORMA', '12:59:00', 90), ('800000014', 'JUAN TORRES VEGA', '12:40:00', 125)]
    return {'a1': S.crear_a1(str(d / 'mov.xlsx'), fac), 'a2': S.crear_a2(str(d / 'banco.xlsx'), banco), 'dir': d, 'fac': fac, 'banco': banco}


def estados(res): return dict(zip(res['facturas']['Número Factura'], res['facturas']['estado_conciliacion']))


def descartar(c, bid, nf): return post(c, f'/borrador/{bid}/descartar', f'/borrador/{bid}', factura=nf)


# ---------- motor ----------
def test_motor_sin_rechazos_resultado_identico(caso):
    base = motor.conciliar(caso['a1'], caso['a2'], 0)
    for r in (None, {}, {'40002': []}):
        x = motor.conciliar(caso['a1'], caso['a2'], 0, rechazos=r)
        assert x['facturas'].equals(base['facturas']) and x['banco'].drop(columns=[]).equals(base['banco']) and x['descartes'] == []
        assert x['totales']['Estado QR'] == base['totales']['Estado QR']
    # con confirmaciones/elecciones previas tampoco cambia
    a = motor.conciliar(caso['a1'], caso['a2'], 0, ['40002'], {})
    b = motor.conciliar(caso['a1'], caso['a2'], 0, ['40002'], {}, rechazos={})
    assert a['facturas'].equals(b['facturas'])


def test_motor_rechazo_excluye_par_y_libera_movimiento(caso):
    base = motor.conciliar(caso['a1'], caso['a2'], 0)
    assert estados(base)['40004'] == 'POR CONFIRMAR' and base['facturas'].set_index('Número Factura').loc['40004', 'ID TRANSACCION'] == '800000004'
    r = motor.conciliar(caso['a1'], caso['a2'], 0, rechazos={'40002': ['800000002'], '40003': ['800000003']})
    f = r['facturas'].set_index('Número Factura'); b = r['banco'].set_index('ID TRANSACCION')
    assert f.loc['40002', 'estado_conciliacion'] == 'SIN RESPALDO' and f.loc['40003', 'estado_conciliacion'] == 'SIN RESPALDO'
    assert b.loc['800000002', 'estado_banco'].startswith('AJENO') and b.loc['800000003', 'estado_banco'].startswith('AJENO')
    assert {d['id'] for d in r['descartes']} == {'800000002', '800000003'}
    assert f.loc['40001', 'estado_conciliacion'] == 'CONCILIADA'
    assert 'candidato sugerido' not in f.loc['40002', 'candidatos']       # el descartado no se vuelve a sugerir ni como DEBIL


def test_motor_reevalua_con_otra_coincidencia_y_libera_para_otra_factura():
    fac = [('50001', 'EFECTIVO', 100, 'ANA GOMEZ RUIZ', 'GOMEZ RUIZ ANA', 'Válido', '10:00:00'),
           ('50002', 'EFECTIVO', 100, 'LUIS PAZ MORA', 'PAZ MORA LUIS', 'Válido', '10:05:00')]
    banco = [('900000001', 'GOMEZ RUIZ ANA', '09:59:00', 100), ('900000002', 'PAZ MORA LUIS', '10:04:00', 100)]
    import tempfile, os as _os
    d = tempfile.mkdtemp(); a1 = S.crear_a1(_os.path.join(d, 'a.xlsx'), fac); a2 = S.crear_a2(_os.path.join(d, 'b.xlsx'), banco)
    # mismo monto: si el cajero descarta el par fuerte (50001, 900000001) sin ser POR CONFIRMAR, el motor igual lo excluye (la regla de negocio vive en servicio)
    r = motor.conciliar(a1, a2, 0, rechazos={'50001': ['900000001']}); f = r['facturas'].set_index('Número Factura')
    assert f.loc['50002', 'ID TRANSACCION'] == '900000002' and f.loc['50001', 'estado_conciliacion'] == 'SIN RESPALDO'
    # un movimiento descartado para una factura queda libre para otra
    r = motor.conciliar(a1, a2, 0, rechazos={'50002': ['900000002']}); f = r['facturas'].set_index('Número Factura')
    assert f.loc['50001', 'ID TRANSACCION'] == '900000001' and f.loc['50002', 'estado_conciliacion'] == 'SIN RESPALDO'


def test_motor_eleccion_del_descartado_no_se_aplica_y_rechazo_inexistente_alerta(caso):
    r = motor.conciliar(caso['a1'], caso['a2'], 0, elecciones={'40003': '800000003'}, rechazos={'40003': ['800000003'], '99999': ['1']})
    txt = ' | '.join(r['alertas']['detalle'])
    assert 'usted descartó ese pago para esta factura' in txt and 'Descarte no aplicado' in txt
    assert estados(r)['40003'] == 'SIN RESPALDO'


# ---------- servicio / web ----------
def test_descartar_reglas_y_estado(app, caso):
    c, bid = borrador_nuevo(app, caso, tarjeta='0')
    html = pag(c, bid).text
    assert html.count('No corresponde</button>') == 3 and html.count('Sí, corresponde') == 3
    assert 'El pago <strong class="mono">800000003</strong> de <strong>LUNA SOTO MARIA</strong> no corresponde a la factura <strong>40003</strong>. Podrá elegir otro pago o reclasificarla.' in html
    assert 'Si alguna no corresponde, cierre este aviso y use «No corresponde» en esa factura.' in html
    # factura conciliada FUERTE: error, sin cambios
    descartar(c, bid, '40001'); assert 'no tiene una coincidencia dudosa' in flash_de(c, bid) and leer_estado(app, bid)['rechazos'] == {}
    assert descartar(c, bid, 'NOEXISTE').status_code == 303 and 'No se encontró la factura' in flash_de(c, bid)
    # confirmada: primero deshacer
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='40002')
    descartar(c, bid, '40002'); assert 'primero use «Deshacer»' in flash_de(c, bid) and leer_estado(app, bid)['rechazos'] == {}
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='40002')
    descartar(c, bid, '40002')
    est = leer_estado(app, bid)
    assert est['rechazos'] == {'40002': ['800000002']} and '40002' not in est['confirmaciones']
    assert 'Descartó el pago 800000002 de SOTO VEGA ANDREA para la factura 40002' in flash_de(c, bid)
    html = pag(c, bid).text
    assert 'Observadas (3)' in html and 'Usted descartó el pago <span class="mono">800000002</span>' in html
    assert html.count('No corresponde</button>') == 2
    # ajeno
    assert descartar(login(app, 'beto'), bid, '40003').status_code == 404


def test_caso_del_usuario_cuadra_y_se_presenta(app, caso):
    c, bid = borrador_nuevo(app, caso, tarjeta='0')
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='40002')              # 1 confirmada
    descartar(c, bid, '40003'); descartar(c, bid, '40004')                                  # 2 descartadas
    res = app.state.almacen.conciliar(bid, leer_estado(app, bid)); est = estados(res)
    assert est['40002'] == 'CONFIRMADA POR CAJERO' and est['40003'] == 'SIN RESPALDO'
    assert est['40004'] == 'SIN RESPALDO'
    # el descartado figura en el desplegable, marcado, sin botón de asignar
    html = pag(c, bid).text
    assert html.count('Descartado por usted') == 2 and 'Sin movimientos posibles' in html and 'Elegir movimiento del banco' in html
    cands = servicio.candidatos(res, '40003'); assert [(x['id'], x['descartado']) for x in cands] == [('800000003', True)]
    # 40003: reclasificar a comunicación interna; 40004: asignar otro movimiento (800000014) por elección manual
    assert sorted((x['id'], x['descartado']) for x in servicio.candidatos(res, '40004')) == [('800000004', True), ('800000014', False)]
    assert 'No cuadra' not in pag(c, bid).text
    # manual con el descartado: rechazado
    respaldo(c, bid, '40003', 'QR BANECO', '800000003'); assert 'descartó' in flash_de(c, bid)
    post(c, f'/borrador/{bid}/asignar', f'/borrador/{bid}', factura='40003', id_transaccion='800000003')
    assert estados(app.state.almacen.conciliar(bid, leer_estado(app, bid)))['40003'] == 'SIN RESPALDO'
    # asignar 40004 a 800000014 y reclasificar 40003
    post(c, f'/borrador/{bid}/asignar', f'/borrador/{bid}', factura='40004', id_transaccion='800000014')
    respaldo(c, bid, '40003', 'BNB', 'P12345')
    est = leer_estado(app, bid); assert est['rechazos'] == {'40003': ['800000003'], '40004': ['800000004']}
    html = pag(c, bid).text
    assert '<strong>CUADRA</strong>' in html
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    cid = int(r.headers['location'].rsplit('/', 1)[1])
    # el movimiento descartado 800000003 queda libre (ajeno) y 800000004 también
    cj = json.load(open(os.path.join(carpeta(app, cid), 'cierre.json'), encoding='utf-8'))
    assert set(cj['rechazos']) == {'40003', '40004'} and len(cj['notas_descartes']) == 2
    usados = {r[0] for r in db(app).execute('select id_transaccion from cierre_movimientos where cierre_id=?', (cid,))}
    assert '800000003' not in usados and '800000004' not in usados and '800000014' in usados
    # PDF y detalle
    for nombre in ('cierre.pdf',):
        txt = texto_pdf(os.path.join(carpeta(app, cid), nombre))
        assert 'descartó la coincidencia sugerida (800000003' in txt and 'reclasificó como comunicación interna (BNB P12345)' in txt
        assert 'descartó la coincidencia sugerida (800000004' in txt and 'asignó el pago 800000014' in txt
    for u in ('ing', 'sup'):
        h = login(app, u).get(f'/cierres/{cid}').text
        assert 'Coincidencias descartadas por el cajero' in h and 'El cajero descartó la coincidencia sugerida (800000004' in h
    assert 'Coincidencias descartadas por el cajero' not in c.get(f'/cierres/{cid}').text
    assert revisar(login(app, 'ing'), cid, 'aprobar', 'ok').status_code == 303
    assert 'descartó la coincidencia sugerida' in texto_pdf(os.path.join(carpeta(app, cid), 'cierre_revision.pdf'))


def test_deshacer_descarte_restaura_sugerencia_y_auditoria(app, caso):
    c, bid = borrador_nuevo(app, caso, tarjeta='0')
    descartar(c, bid, '40003')
    assert 'Deshacer descarte' in pag(c, bid).text
    # con respaldo asignado no se deshace el descarte
    respaldo(c, bid, '40003', 'BNB', 'P12345')
    post(c, f'/borrador/{bid}/deshacer-descarte', f'/borrador/{bid}', factura='40003')
    assert 'ya tiene respaldo' in flash_de(c, bid) and '40003' in leer_estado(app, bid)['rechazos']
    post(c, f'/borrador/{bid}/deshacer', f'/borrador/{bid}', factura='40003')
    post(c, f'/borrador/{bid}/deshacer-descarte', f'/borrador/{bid}', factura='40003')
    assert leer_estado(app, bid)['rechazos'] == {} and 'Se deshizo el descarte' in flash_de(c, bid)
    assert estados(app.state.almacen.conciliar(bid, leer_estado(app, bid)))['40003'] == 'POR CONFIRMAR'
    post(c, f'/borrador/{bid}/deshacer-descarte', f'/borrador/{bid}', factura='40003'); assert 'no tiene pagos descartados' in flash_de(c, bid)
    aud = db(app).execute("select usuario, accion, detalle_json, momento from auditoria where accion in ('descartar_coincidencia','deshacer_descarte') order by id").fetchall()
    assert [a[1] for a in aud] == ['descartar_coincidencia', 'deshacer_descarte'] and all(a[0] == 'ana' and a[3] for a in aud)
    d0 = json.loads(aud[0][2]); assert d0['factura'] == '40003' and d0['id_transaccion'] == '800000003' and d0['titular'] == 'LUNA SOTO MARIA'
    assert json.loads(aud[1][2])['ids_transaccion'] == ['800000003']


def test_confirmar_todas_con_restantes_y_frase(app, caso):
    c, bid = borrador_nuevo(app, caso, tarjeta='0')
    descartar(c, bid, '40003')
    html = pag(c, bid).text
    assert 'Confirmar todas las coincidencias dudosas (2)' in html
    post(c, f'/borrador/{bid}/confirmar-todas', f'/borrador/{bid}', facturas=['40002', '40003', '40004'])
    est = leer_estado(app, bid); assert set(est['confirmaciones']) == {'40002', '40004'}
    assert '40003' not in est['confirmaciones'] and est['rechazos'] == {'40003': ['800000003']}


def test_editar_reenviar_conserva_rechazos(app, caso):
    c, bid = borrador_nuevo(app, caso, tarjeta='0')
    est = leer_estado(app, bid); assert est['rechazos'] == {}      # esquema nuevo
    # estado antiguo sin la clave sigue funcionando
    est.pop('rechazos'); app.state.almacen.guardar_estado(bid, est)
    assert pag(c, bid).status_code == 200
    d = app.state.almacen
    nuevo, info = d.descartar_coincidencia(bid, est, '40003'); assert nuevo['rechazos'] == {'40003': ['800000003']}


# ---------- datos reales ----------
@real
def test_real_rechazo_sin_cambio_con_vacio_y_efecto_con_rechazo():
    a1 = glob.glob(os.path.join(DIR_REAL, '*IngresosDiarios*.xls'))[0]; a2 = glob.glob(os.path.join(DIR_REAL, '*Reporte_pagos*.xlsx'))[0]
    base = motor.conciliar(a1, a2, 0)
    assert motor.conciliar(a1, a2, 0, rechazos={})['facturas'].equals(base['facturas'])
    pc = base['facturas'][base['facturas']['estado_conciliacion'] == 'POR CONFIRMAR']
    for _, r in pc.iterrows():
        x = motor.conciliar(a1, a2, 0, rechazos={r['Número Factura']: [r['ID TRANSACCION']]})
        f = x['facturas'].set_index('Número Factura').loc[r['Número Factura']]
        assert f['ID TRANSACCION'] != r['ID TRANSACCION'] and f['estado_conciliacion'] in ('SIN RESPALDO', 'POR CONFIRMAR', 'CONCILIADA')
        assert x['banco'].set_index('ID TRANSACCION').loc[r['ID TRANSACCION'], 'estado_banco'] != 'ASIGNADO POR CONFIRMAR' or f['ID TRANSACCION'] == ''
