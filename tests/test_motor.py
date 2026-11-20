"""Tests del motor con datos REALES (privados). Se leen desde DATOS_PRUEBA_DIR y se omiten si no existe.
Los archivos sinteticos derivados se generan en un directorio temporal, nunca en el repositorio.
Ejecutar: DATOS_PRUEBA_DIR=/ruta/con/reportes python -m pytest tests -v"""
import os, glob, tempfile, datetime as dt
import pandas as pd, openpyxl, pytest
from app import motor as mc

_DIR = os.environ.get('DATOS_PRUEBA_DIR', '')
_a1 = sorted(glob.glob(os.path.join(_DIR, '*IngresosDiarios*.xls'))) if _DIR else []
_a2 = sorted(glob.glob(os.path.join(_DIR, '*Reporte_pagos*.xlsx'))) if _DIR else []
pytestmark = pytest.mark.skipif(not (_a1 and _a2), reason='DATOS_PRUEBA_DIR no definido o sin los reportes reales')
A1 = _a1[0] if _a1 else ''
A2 = _a2[0] if _a2 else ''
DATOS = tempfile.mkdtemp(prefix='sinteticos_')
# nombres tomados en tiempo de ejecucion de los datos reales (no se guardan en el repositorio)
RS, ES = {}, {}
if A1:
    _df = mc.parse_a1(A1); RS = dict(zip(_df['Número Factura'], _df['Razon Social'])); ES = dict(zip(_df['Número Factura'], _df['Nombre Estudiante']))

def a2_sintetico(nombre, agregar=(), quitar_ids=()):
    """Copia el A2 real con openpyxl, agrega filas (antes de TOTAL) y/o elimina por ID TRANSACCION."""
    wb = openpyxl.load_workbook(A2); ws = wb.active
    for r in range(ws.max_row, 1, -1):
        if str(ws.cell(r, 2).value) in quitar_ids: ws.delete_rows(r)
    fila_total = next(r for r in range(2, ws.max_row + 1) if str(ws.cell(r, 8).value).upper() == 'TOTAL')
    for k, (idt, tit, fh, monto) in enumerate(agregar):
        ws.insert_rows(fila_total + k)
        for c, v in enumerate(['26070801016676059507', idt, tit, '******0001', 'Ingresos Cursos Pregrado', '3041210569',
                               fh, 'SEDE TRINIDAD', f'{monto:,.2f} BOB'], start=1):
            ws.cell(fila_total + k, c, v)
    ruta = f'{DATOS}/{nombre}.xlsx'; wb.save(ruta); return ruta

def est(res, fact): return res['facturas'].set_index('Número Factura').loc[fact]
def mov(res, idt): return res['banco'].set_index('ID TRANSACCION').loc[idt]

def test_T1_caso_real():
    r = mc.conciliar(A1, A2, 1191); t = r['totales']
    assert t['n_facturas_qr'] == 14 and t['QR registrado'] == 10784 and t['n_movimientos_ajenos'] == 0
    assert t['Diferencia QR'] == 615 and t['Tarjeta según sistema'] == 1191 and t['Diferencia tarjeta (declarada - sistema)'] == 0
    assert t['Total cierre'] == 10784 + 1191
    assert est(r, '20043')['estado_conciliacion'].startswith('EXCLUIDA')
    assert est(r, '20045')['estado_conciliacion'] == 'TARJETA' and est(r, '20051')['estado_conciliacion'] == 'TARJETA'
    niveles = {f: est(r, f)['nivel'] for f in r['facturas']['Número Factura'] if est(r, f)['tipo' if False else 'Tipo (motor)'] == 'QR'}
    print('\nT1 niveles:', niveles, '| estado QR:', t['Estado QR'])
    assert niveles['20048'] == 'MEDIA' and all(v == 'FUERTE' for k, v in niveles.items() if k != '20048')
    assert est(r, '20048')['estado_conciliacion'] == 'POR CONFIRMAR'
    assert (r['facturas'].query("`Tipo (motor)`=='QR'")['ID TRANSACCION'] != '').all()
    assert t['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN'
    assert t['QR conciliado automático (CONCILIADA)'] == 10169 and t['QR pendiente de confirmar (POR CONFIRMAR + AMBIGUO)'] == 615
    assert t['Conformidad del cajero'] == 'SIN CONFORMIDAD'

def test_T1b_con_confirmacion():
    r = mc.conciliar(A1, A2, 1191, confirmaciones=['20048']); t = r['totales']
    assert est(r, '20048')['estado_conciliacion'] == 'CONFIRMADA POR CAJERO'
    assert t['Diferencia QR'] == 0 and t['Estado QR'] == 'CUADRA' and t['QR confirmado por cajero'] == 615
    r2 = mc.conciliar(A1, A2, 1191, confirmaciones=['20048'], conformidad='CONFORME')
    assert r2['totales']['Conformidad del cajero'] == 'CONFORME'
    r3 = mc.conciliar(A1, A2, 1191, confirmaciones=['20035', '99999'])      # 20035 no esta POR CONFIRMAR
    assert r3['totales']['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN' and (r3['alertas']['detalle'].str.contains('Confirmación no aplicada')).sum() == 2

def test_T2_ajenos():
    ruta = a2_sintetico('A2_T2_ajenos', [('900000001', 'PEDRO PEREZ LOPEZ', '16/10/2026 16:52:00', 933),   # monto 933 a 3.7 min de la 20044
                                          ('900000002', 'MARIA GOMEZ RIOS', '16/10/2026 12:00:00', 450),
                                          ('900000003', 'LUIS VARGAS SOTO', '16/10/2026 13:30:00', 1000)])
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048']); t = r['totales']
    assert est(r, '20044')['ID TRANSACCION'] == '419795667' and est(r, '20044')['estado_conciliacion'] == 'CONCILIADA'
    for i in ['900000001', '900000002', '900000003']: assert mov(r, i)['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO'
    assert t['n_movimientos_ajenos'] == 3 and t['Diferencia QR'] == 0
    print('\nT2: ajenos=3, dif QR=0, estado', t['Estado QR'])

def test_T3_sin_respaldo():
    ruta = a2_sintetico('A2_T3_sin_respaldo', quitar_ids=['419885341'])
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048']); t = r['totales']
    assert est(r, '20047')['estado_conciliacion'] == 'SIN RESPALDO'
    assert t['Diferencia QR'] == 1191 and t['Estado QR'] == 'NO CUADRA'
    print('\nT3: 20047 SIN RESPALDO, dif', t['Diferencia QR'], t['Estado QR'])

def test_T4_duplicado():
    ruta = a2_sintetico('A2_T4_dup', [('419885341', ES['20047'], '16/10/2026 18:12:28', 1191)])
    r = mc.conciliar(A1, ruta, 1191); b = r['banco']; t = r['totales']
    d = b[b['ID TRANSACCION'] == '419885341']
    assert sorted(d['estado_banco']) == ['ASIGNADO CONCILIADA', 'DUPLICADO EN ARCHIVO BANCO']
    assert t['n_duplicados'] == 1 and t['QR conciliado automático (CONCILIADA)'] == 10169 and t['n_movimientos_ajenos'] == 0
    assert (r['facturas']['ID TRANSACCION'] == '419885341').sum() == 1
    print('\nT4: duplicados=1, usado una vez, conciliado auto', t['QR conciliado automático (CONCILIADA)'])

def test_T5_fuera_de_fecha():
    ruta = a2_sintetico('A2_T5_fuera_fecha', [('900000005', RS['20047'], '17/10/2026 18:12:00', 1191)])
    r = mc.conciliar(A1, ruta, 1191)
    assert mov(r, '900000005')['estado_banco'] == 'FUERA DE FECHA'
    assert est(r, '20047')['ID TRANSACCION'] == '419885341' and r['totales']['n_movimientos_ajenos'] == 0
    print('\nT5: FUERA DE FECHA, 20047 conserva su movimiento real')

def test_T5b_fuera_de_fecha_sin_original():
    """Si ademas se quita el movimiento real, el de otra fecha NO puede respaldar la factura."""
    ruta = a2_sintetico('A2_T5b', [('900000005', RS['20047'], '17/10/2026 18:12:00', 1191)], quitar_ids=['419885341'])
    r = mc.conciliar(A1, ruta, 1191)
    assert est(r, '20047')['estado_conciliacion'] == 'SIN RESPALDO' and r['totales']['Estado QR'] == 'NO CUADRA'

def test_T6_empate():
    # Dos movimientos de 25, mismo titular que 20041 (16:14:29), equidistantes: 16:14:02 (real) y 16:14:56
    ruta = a2_sintetico('A2_T6_empate', [('900000006', RS['20041'], '16/10/2026 16:14:56', 25)])
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048']); t = r['totales']
    print('\nT6: 20041 ->', est(r, '20041')['estado_conciliacion'], '| 419761158 ->', mov(r, '419761158')['estado_banco'],
          '| 900000006 ->', mov(r, '900000006')['estado_banco'], '| dif QR', t['Diferencia QR'], t['Estado QR'])
    assert est(r, '20041')['estado_conciliacion'] == 'AMBIGUO'
    assert mov(r, '419761158')['estado_banco'].startswith('AMBIGUO') and mov(r, '900000006')['estado_banco'].startswith('AMBIGUO')
    assert t['Diferencia QR'] == 25 and t['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN'   # sin SIN RESPALDO -> pendiente (regla D)
    assert est(r, '20049')['estado_conciliacion'] == 'CONCILIADA'      # el otro 25 no se contamina

def test_T7_tipo_banco_y_no_reconocidos():
    a1 = mc.parse_a1(A1)
    a1.loc[a1['Número Factura'] == '20036', 'Tipo Pago'] = 'BANCO'
    a1 = pd.concat([a1, a1.iloc[[0]].assign(**{'Número Factura': '29999', 'Tipo Pago': 'CHEQUE', 'Monto': 7.0}),
                    a1.iloc[[0]].assign(**{'Número Factura': '29998', 'Estado': 'Pendiente', 'Monto': 8.0})], ignore_index=True)
    f, al = mc.preparar_a1(a1)
    assert f.set_index('Número Factura').loc['20036', 'tipo'] == 'QR'
    assert sum('valor no reconocido' in a[2] for a in al) == 2
    r = mc.conciliar_df(a1, mc.parse_a2(A2), 1191); t = r['totales']
    assert est(r, '20036')['estado_conciliacion'] == 'CONCILIADA' and t['QR registrado'] == 10784
    assert est(r, '29999')['estado_conciliacion'].startswith('NO PROCESADA')
    assert est(r, '29998')['estado_conciliacion'].startswith('NO PROCESADA')

def test_T6b_empate_resuelto_con_eleccion():
    ruta = a2_sintetico('A2_T6_empate', [('900000006', RS['20041'], '16/10/2026 16:14:56', 25)])
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048'], elecciones={'20041': '419761158'}); t = r['totales']
    assert est(r, '20041')['estado_conciliacion'] == 'CONFIRMADA POR CAJERO' and est(r, '20041')['ID TRANSACCION'] == '419761158'
    assert mov(r, '900000006')['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO'
    assert t['Diferencia QR'] == 0 and t['Estado QR'] == 'CUADRA' and t['n_movimientos_ajenos'] == 1
    r2 = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048'], elecciones={'20041': '900000006'})   # la otra eleccion tambien vale
    assert mov(r2, '419761158')['estado_banco'].startswith('AJENO') and r2['totales']['Estado QR'] == 'CUADRA'

def test_eleccion_invalida():
    ruta = a2_sintetico('A2_T3_sin_respaldo', quitar_ids=['419885341'])
    # monto distinto (movimiento de 800 para factura 1191)
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048'], elecciones={'20047': '419919681'})
    assert est(r, '20047')['estado_conciliacion'] == 'SIN RESPALDO' and r['alertas']['detalle'].str.contains('monto distinto').any()
    # ID ya usado por otra factura (mismo monto 25: 419889217 pertenece a 20049)
    ruta6 = a2_sintetico('A2_T6_empate', [('900000006', RS['20041'], '16/10/2026 16:14:56', 25)])
    r = mc.conciliar(A1, ruta6, 1191, elecciones={'20041': '419889217'})
    assert est(r, '20041')['estado_conciliacion'] == 'AMBIGUO' and r['alertas']['detalle'].str.contains('ya usado').any()
    # ID inexistente
    r = mc.conciliar(A1, ruta, 1191, elecciones={'20047': '123'})
    assert r['alertas']['detalle'].str.contains('inexistente').any() and est(r, '20047')['estado_conciliacion'] == 'SIN RESPALDO'
    # fecha distinta
    ruta5 = a2_sintetico('A2_T5b', [('900000005', RS['20047'], '17/10/2026 18:12:00', 1191)], quitar_ids=['419885341'])
    r = mc.conciliar(A1, ruta5, 1191, elecciones={'20047': '900000005'})
    assert r['alertas']['detalle'].str.contains('fecha distinta').any() and est(r, '20047')['estado_conciliacion'] == 'SIN RESPALDO'

def test_eleccion_valida_sin_respaldo():
    ruta = a2_sintetico('A2_T8_robo', [('900000008', 'PEDRO PEREZ LOPEZ', '16/10/2026 18:10:00', 1191)], quitar_ids=['419885341'])
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048'], elecciones={'20047': '900000008'})
    assert est(r, '20047')['estado_conciliacion'] == 'CONFIRMADA POR CAJERO' and r['totales']['Estado QR'] == 'CUADRA'

def test_conformidad():
    with pytest.raises(ValueError, match='CONFORME'):
        mc.conciliar(A1, A2, 1191, conformidad='CONFORME')                      # PENDIENTE
    ruta = a2_sintetico('A2_T3_sin_respaldo', quitar_ids=['419885341'])
    with pytest.raises(ValueError):
        mc.conciliar(A1, ruta, 1191, confirmaciones=['20048'], conformidad='CONFORME')   # NO CUADRA
    with pytest.raises(ValueError, match='motivo'):
        mc.conciliar(A1, ruta, 1191, conformidad='NO CONFORME', motivo='  ')
    r = mc.conciliar(A1, ruta, 1191, conformidad='NO CONFORME', motivo='falta un pago')
    rs = r['resumen'].set_index('concepto')['valor']
    assert rs['Conformidad del cajero'] == 'NO CONFORME' and rs['Motivo conformidad'] == 'falta un pago' and rs['Generado el']

def test_T8_nombre_nulo_con_hora_ya_no_es_media():
    """v1.1: deposito ajeno de 1191 a <=10 min de la factura es solo DEBIL: no se asigna, queda como candidato sugerido."""
    ruta = a2_sintetico('A2_T8_robo', [('900000008', 'PEDRO PEREZ LOPEZ', '16/10/2026 18:10:00', 1191)], quitar_ids=['419885341'])
    r = mc.conciliar(A1, ruta, 1191, confirmaciones=['20048']); t = r['totales']
    f = est(r, '20047')
    print('\nT8 v1.1: 20047 ->', f['estado_conciliacion'], '|', f['candidatos'], '| dif QR', t['Diferencia QR'], '|', t['Estado QR'])
    assert f['estado_conciliacion'] == 'SIN RESPALDO' and '900000008' in f['candidatos']
    assert t['Diferencia QR'] == 1191 and t['Estado QR'] == 'NO CUADRA'
    assert mov(r, '900000008')['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO'
