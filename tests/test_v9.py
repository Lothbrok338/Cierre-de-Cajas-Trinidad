"""Pruebas v9: banco BANECO como comunicacion interna (deposito de fecha anterior) distinto de QR BANECO, antiduplicados cruzados con el reporte QR
y con otros cierres, y fecha del deposito opcional. Datos ficticios."""
import json, os, sqlite3
import openpyxl
import pytest
from tests.conftest import login, token, post
from tests import sinteticos as S
from tests.test_app import bid_de, db, pag, presentar
from tests.test_v2 import flash_de, leer_estado
from tests.test_v5 import carpeta, texto_pdf
from tests.test_v8 import caso, caso_g, construir, hacer_plantilla, nuevo, subir_pl, problemas, estados, resumen, TRES
from app import catalogo, db as dbm, motor, plantilla as PL, servicio


def resp(c, bid, nf, banco, codigo, fecha=''):
    return post(c, f'/borrador/{bid}/respaldo', f'/borrador/{bid}', factura=nf, banco=banco, codigo=codigo, fecha_deposito=fecha)


def caso_con_banco(d, banco):
    """Mismo movimiento que el caso base de v8, con otro reporte QR (lista [(id, titular, hora, monto)])."""
    c = construir(d); S.crear_a2(str(d / 'banco.xlsx'), banco); return c


BANCO_BASE = [('800000001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500), ('800000004', 'PAZ RIOS SILVIA', '12:44:30', 125), ('800000099', 'PERSONA AJENA LIMA', '12:00:00', 900)]
RESTO = [('60002', 'BNB', 'P12345'), ('60003', 'BMSC', 'T98765'), ('60005', 'BISA', '4567890')]    # deja 60004 POR CONFIRMAR por resolver


def presentado(app, caso, usuario, filas):
    """Cierre presentado por `usuario` con las C.I. de `filas` desde plantilla (60001 queda CONCILIADA; 60004 se confirma)."""
    c, bid, _ = nuevo(app, caso, filas, usuario=usuario)
    post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='60004') if not any(f[0] == '60004' for f in filas) else None
    r = presentar(c, bid); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, bid)
    return c, bid, int(r.headers['location'].rsplit('/', 1)[1])


# ---------- catalogo ----------
def test_catalogo_baneco_ci_distinto_de_qr_baneco():
    assert 'BANECO' in catalogo.LISTA_BANCOS and 'QR BANECO' in catalogo.LISTA_BANCOS
    b = catalogo.BANCOS['BANECO']
    assert b['es_qr'] is False and b['cuenta_bancaria'] is None and b['cuenta_sap'] is None and not catalogo.es_qr('BANECO') and catalogo.es_qr('QR BANECO')
    assert catalogo.regla_formato('BANECO') == 'BANECO: solo números (número de transacción del banco)'
    assert catalogo.formato_incumplido('BANECO', '123456') is None and catalogo.formato_incumplido('BANECO', 'AB123') == catalogo.regla_formato('BANECO')
    assert catalogo.validar_codigo('BANECO', ' 5551234 ') == (True, None, '5551234') and catalogo.validar_codigo('BANECO', '5551234.0')[2] == '5551234'
    assert catalogo.validar_codigo('QR BANECO', '123')[0] is False                      # QR BANECO no es C.I.
    assert catalogo.etiqueta('QR BANECO') == 'QR BANECO (pago del día, está en el reporte)' and catalogo.etiqueta('BANECO') == 'BANECO (depósito de fecha anterior)'
    assert catalogo.normalizar_codigo('p123.0') == 'P123.0' and catalogo.normalizar_codigo(' 98765.00 ') == '98765'
    assert PL.BANCOS_PLANTILLA == ['BNB', 'BMSC', 'BISA', 'UNION', 'BANECO']


def test_fechas_del_deposito():
    from datetime import date
    fc = date(2026, 10, 16)
    assert catalogo.normalizar_fecha_deposito('5/9/2026') == '05/09/2026' and catalogo.normalizar_fecha_deposito('2026-09-05T00:00:00') == '05/09/2026'
    assert catalogo.alerta_fecha_deposito('BANECO', '', fc) == catalogo.MSG_FECHA_FALTA and catalogo.alerta_fecha_deposito('BNB', '', fc) is None
    assert catalogo.alerta_fecha_deposito('BANECO', '05/09/2026', fc) is None and catalogo.alerta_fecha_deposito('BANECO', '16/10/2026', fc) is None
    assert 'posterior a la fecha del cierre' in catalogo.alerta_fecha_deposito('BANECO', '17/10/2026', fc)
    assert 'no es válida' in catalogo.alerta_fecha_deposito('BNB', '31/02/2026', fc) and 'no es válida' in catalogo.alerta_fecha_deposito('BANECO', 'ayer', fc)


# ---------- formulario manual ----------
def test_formulario_muestra_ambas_opciones_con_ayuda(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2'])); h = pag(c, bid).text
    assert 'value="QR BANECO">QR BANECO (pago del día, está en el reporte)<' in h and 'value="BANECO">BANECO (depósito de fecha anterior)<' in h
    assert 'BANECO: solo números (número de transacción del banco)' in h and 'name="fecha_deposito"' in h and 'Fecha del depósito' in h
    assert 'depósito hecho a BANECO en una fecha anterior' in h


def test_baneco_por_formulario_feliz_total_y_cierre(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    resp(c, bid, '60002', 'BANECO', '5551234', '5/9/2026'); m = flash_de(c, bid)
    assert 'respaldo manual registrado (BANECO 5551234)' in m and 'Fecha del depósito: 05/09/2026' in m and 'FORMATO A REVISAR' not in m and 'Falta la fecha' not in m
    est = leer_estado(app, bid); assert est['reclasificaciones']['60002'] == {'banco': 'BANECO', 'codigo': '5551234', 'fecha_deposito': '05/09/2026'}
    resp(c, bid, '60003', 'BMSC', 'T98765'); resp(c, bid, '60005', 'BISA', '4567890'); post(c, f'/borrador/{bid}/confirmar', f'/borrador/{bid}', factura='60004')
    t = app.state.almacen.conciliar(bid, leer_estado(app, bid))['totales']
    assert t['C.I. según sistema'] == 1200 and t['QR registrado'] == 625 and t['Total cierre'] == 625 + 1200 + 200
    h = pag(c, bid).text; assert 'Comunicación interna — BANECO 5551234' in h and 'Fecha del depósito: 05/09/2026' in h and 'class="alerta-fmt"' not in h
    r = presentar(c, bid); assert r.status_code == 303, flash_de(c, bid); cid = int(r.headers['location'].rsplit('/', 1)[1])
    cj = json.load(open(os.path.join(carpeta(app, cid), 'cierre.json'), encoding='utf-8'))
    ci = {x['factura']: x for x in cj['comunicaciones_internas']}; assert ci['60002']['banco'] == 'BANECO' and ci['60002']['fecha_deposito'] == '05/09/2026'
    assert cj['alertas_formato'] == [] and cj['totales']['Total cierre'] == 2025
    con = db(app); assert con.execute("select banco, codigo, fecha_deposito, monto from cierre_ci where factura='60002'").fetchone() == ('BANECO', '5551234', '05/09/2026', 700.0)
    assert con.execute('select total_cierre, ci_total from cierres').fetchone() == (2025.0, 1200.0)
    for u in ('ing', 'sup'):
        d = login(app, u).get(f'/cierres/{cid}').text
        assert 'Fecha del depósito' in d and '05/09/2026' in d and 'BANECO' in d and 'PREVALIDADO' in d
    txt = texto_pdf(os.path.join(carpeta(app, cid), 'cierre.pdf'))
    assert 'BANECO' in txt and '5551234' in txt and '05/09/2026' in txt and 'Fecha depósito' in txt


def test_qr_baneco_manual_sigue_validandose_contra_el_reporte(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    resp(c, bid, '60002', 'QR BANECO', '800000099', '05/09/2026'); assert 'no es válido para la factura 60002' in flash_de(c, bid)      # monto distinto: el motor valida
    resp(c, bid, '60002', 'QR BANECO', '123456789'); assert 'inexistente en el reporte' in flash_de(c, bid)
    assert leer_estado(app, bid)['reclasificaciones'] == {}
    resp(c, bid, '60002', 'BANECO', '123456789'); assert 'respaldo manual registrado (BANECO 123456789)' in flash_de(c, bid)      # el mismo numero como BANECO (no esta en el reporte) si pasa
    assert estados(app, bid)['60002'] == 'COMUNICACIÓN INTERNA'


def test_codigo_no_numerico_es_alerta_no_bloqueante(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    resp(c, bid, '60002', 'BANECO', 'AB-77', '05/09/2026'); m = flash_de(c, bid)
    assert 'respaldo manual registrado (BANECO AB-77)' in m and 'FORMATO A REVISAR' in m and 'BANECO: solo números' in m
    assert leer_estado(app, bid)['reclasificaciones']['60002']['codigo'] == 'AB-77'
    h = pag(c, bid).text; assert 'class="alerta-fmt"' in h and 'Alertas de revisión (1)' in h


def test_baneco_sin_fecha_fecha_invalida_y_futura_son_alertas(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    resp(c, bid, '60002', 'BANECO', '1110001'); m = flash_de(c, bid)
    assert 'Falta la fecha del depósito: Ingresos la necesita para ubicarlo en el extracto del mes correcto' in m
    resp(c, bid, '60003', 'BANECO', '1110002', '31/12/2030'); assert 'posterior a la fecha del cierre' in flash_de(c, bid)
    resp(c, bid, '60005', 'BANECO', '1110003', 'ayer'); assert 'no es válida' in flash_de(c, bid)
    resp(c, bid, '60004', 'BNB', 'P99999'); assert 'Falta la fecha' not in flash_de(c, bid) and 'no es válida' not in flash_de(c, bid)    # otros bancos: la fecha no es necesaria
    assert estados(app, bid)['60002'] == estados(app, bid)['60003'] == estados(app, bid)['60005'] == 'COMUNICACIÓN INTERNA'
    h = pag(c, bid).text; assert 'Alertas de revisión (3)' in h and h.count('FECHA DEL DEPÓSITO') >= 3 and 'ayer' in h
    r = presentar(c, bid); assert r.status_code == 303, flash_de(c, bid); cid = int(r.headers['location'].rsplit('/', 1)[1])
    assert sorted(x[0] for x in db(app).execute("select tipo from alertas_cierre where cierre_id=?", (cid,))) == ['FECHA DEPÓSITO'] * 3
    d = login(app, 'ing').get(f'/cierres/{cid}').text; assert 'Alertas de revisión (3' in d and 'Falta la fecha del depósito' in d
    txt = texto_pdf(os.path.join(carpeta(app, cid), 'cierre.pdf')); assert 'Alertas de revisión' in txt and 'ayer' in txt


# ---------- antiduplicados ----------
def test_baneco_igual_a_id_del_reporte_qr_del_borrador_se_bloquea(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    for idt in ('800000099', '800000004', '800000001'):
        resp(c, bid, '60002', 'BANECO', idt)
        assert 'Ese número de transacción aparece en el reporte QR de este día: respalde la factura como QR BANECO, no como comunicación interna' in flash_de(c, bid)
    assert leer_estado(app, bid)['reclasificaciones'] == {}
    # por plantilla: la fila se rechaza con el mismo mensaje
    c2, b2, _ = nuevo(app, caso, [('60002', 'BANECO', '800000099'), ('60003', 'BANECO', '555.0')], usuario='beto')
    p = problemas(app, b2); assert list(p) == [2] and 'aparece en el reporte QR de este día' in p[2]['motivo']
    assert leer_estado(app, b2)['reclasificaciones']['60003']['codigo'] == '555'


def test_baneco_distinto_de_bancos_de_otras_cuentas_con_mismo_numero(app, caso):
    """El numero del reporte QR solo bloquea a BANECO: BISA usa otra cuenta, otro catalogo de numeros."""
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    resp(c, bid, '60002', 'BISA', '800000099'); assert 'respaldo manual registrado (BISA 800000099)' in flash_de(c, bid)


def test_baneco_ya_usado_como_respaldo_qr_en_otro_cierre(app, tmp_path):
    otro = caso_con_banco(tmp_path / 'o', [('7770001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500), ('800000004', 'PAZ RIOS SILVIA', '12:44:30', 125)])
    _, _, cid = presentado(app, otro, 'beto', RESTO)
    assert db(app).execute("select count(*) from cierre_movimientos where id_transaccion='7770001'").fetchone()[0] == 1
    ana = construir(tmp_path / 'a')
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, ana['a1'], ana['a2']))
    resp(c, bid, '60002', 'BANECO', '7770001'); m = flash_de(c, bid)
    assert 'ya se usó como respaldo QR en el cierre CC-' in m and 'beto' in m and leer_estado(app, bid)['reclasificaciones'] == {}
    c2, b2, _ = nuevo(app, ana, [('60002', 'BANECO', '7770001', '05/09/2026')], usuario='ana')
    p = problemas(app, b2); assert 'ya se usó como respaldo QR' in p[2]['motivo']


def test_baneco_ya_usado_como_ci_baneco_en_otro_cierre(app, tmp_path):
    otro = construir(tmp_path / 'o')
    _, bid_b, cid_b = presentado(app, otro, 'beto', [('60002', 'BANECO', '7770001', '01/09/2026'), ('60003', 'BMSC', 'T98765'), ('60005', 'BISA', '4567890')])
    ana = caso_con_banco(tmp_path / 'a', BANCO_BASE)
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, ana['a1'], ana['a2']))
    resp(c, bid, '60002', 'BANECO', '7770001', '02/09/2026'); m = flash_de(c, bid)
    assert 'El código 7770001 de BANECO ya está registrado en el cierre CC-' in m and 'beto' in m and leer_estado(app, bid)['reclasificaciones'] == {}
    c2, b2, _ = nuevo(app, ana, [('60002', 'BANECO', '7770001.0')], usuario='ana'); p = problemas(app, b2)
    assert 'ya está registrado en el cierre CC-' in p[2]['motivo'] and p[2]['codigo'] == '7770001'
    # el codigo queda guardado una sola vez, con su fecha
    assert db(app).execute("select banco, codigo, fecha_deposito from cierre_ci where banco='BANECO'").fetchall() == [('BANECO', '7770001', '01/09/2026')]


def test_qr_con_id_ya_usado_como_ci_baneco_en_otro_cierre_se_bloquea(app, tmp_path):
    otro = construir(tmp_path / 'o')
    presentado(app, otro, 'beto', [('60002', 'BANECO', '7770001', '01/09/2026'), ('60003', 'BMSC', 'T98765'), ('60005', 'BISA', '4567890')])
    # ana: el reporte trae el 7770001 (ajeno, 700) y su factura 60002 lo toma por respaldo manual QR; 60001 (conciliada) usa otro id
    ana = caso_con_banco(tmp_path / 'a', BANCO_BASE + [('7770001', 'PERSONA SIN RELACION', '10:58:50', 700)])
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, ana['a1'], ana['a2']))
    resp(c, bid, '60002', 'QR BANECO', '7770001'); m = flash_de(c, bid)
    assert 'ya figura como comunicación interna BANECO en el cierre CC-' in m and 'beto' in m
    # tambien en una conciliacion automatica: se bloquea la presentacion con el mismo mensaje
    ana2 = caso_con_banco(tmp_path / 'a2', [('7770001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500), ('800000004', 'PAZ RIOS SILVIA', '12:44:30', 125)])
    c2, b2, _ = nuevo(app, ana2, [('60002', 'BNB', 'P22222'), ('60003', 'BMSC', 'T22222'), ('60005', 'BISA', '2222222'), ('60004', 'UNION', '2223334')], usuario='ana')
    h = pag(c2, b2).text; assert 'ya figura como comunicación interna BANECO en el cierre CC-' in h
    r = presentar(c2, b2); assert r.status_code == 303 and r.headers['location'].startswith('/borrador/')
    assert 'ya figura como comunicación interna BANECO' in flash_de(c2, b2) and db(app).execute("select count(*) from cierres where usuario='ana'").fetchone()[0] == 0


def test_baneco_repetido_dentro_del_mismo_borrador(app, caso):
    c = login(app, 'ana'); bid = bid_de(subir_pl(c, caso['a1'], caso['a2']))
    resp(c, bid, '60002', 'BANECO', '1234567'); assert 'respaldo manual registrado' in flash_de(c, bid)
    resp(c, bid, '60003', 'BANECO', '1234567'); assert 'ya está registrado en la factura 60002 de este cierre' in flash_de(c, bid)
    c2, b2, _ = nuevo(app, caso, [('60002', 'BANECO', '1234567'), ('60003', 'BANECO', '1234567')], usuario='beto')
    p = problemas(app, b2); assert 'repetido en la plantilla' in p[2]['motivo'] and 'repetido en la plantilla' in p[3]['motivo']
    assert resumen(app, b2) == (2, 0, 2)


# ---------- plantilla ----------
def test_plantilla_nueva_con_columna_de_fecha_y_baneco(app, caso):
    filas = [('60002', 'baneco', 5551234.0, '05/09/2026'), ('60003', 'BMSC', 'T98765', None), ('60005', 'BANECO', '5551235', '2026-09-06')]
    c, bid, pl = nuevo(app, caso, filas)
    assert resumen(app, bid) == (3, 3, 0)
    est = leer_estado(app, bid)
    assert est['reclasificaciones'] == {'60002': {'banco': 'BANECO', 'codigo': '5551234', 'fecha_deposito': '05/09/2026'}, '60003': {'banco': 'BMSC', 'codigo': 'T98765'},
                                        '60005': {'banco': 'BANECO', 'codigo': '5551235', 'fecha_deposito': '06/09/2026'}}
    t = app.state.almacen.conciliar(bid, est)['totales']; assert t['Total cierre'] == t['QR registrado'] + t['C.I. según sistema'] + 200
    h = pag(c, bid).text; assert 'Comunicación interna — BANECO 5551234' in h and 'Fecha del depósito: 06/09/2026' in h and 'Alertas de revisión' not in h
    assert PL.leer(pl)[0]['fecha_deposito'] == '05/09/2026'
    r = servicio.alertas_revision(app.state.almacen.conciliar(bid, est), est); assert r == []


def test_plantilla_baneco_sin_fecha_o_con_fecha_mala_genera_alertas_y_se_aplica(app, caso):
    c, bid, _ = nuevo(app, caso, [('60002', 'BANECO', '5551234'), ('60003', 'BANECO', '5551235', '01/01/2031'), ('60005', 'BANECO', '5551236', 'xx')])
    assert resumen(app, bid) == (3, 3, 0)
    est = leer_estado(app, bid); al = servicio.alertas_revision(app.state.almacen.conciliar(bid, est), est)
    assert [(a['factura'], a['tipo']) for a in al] == [('60002', 'FECHA DEPÓSITO'), ('60003', 'FECHA DEPÓSITO'), ('60005', 'FECHA DEPÓSITO')]
    assert 'Falta la fecha del depósito' in al[0]['detalle'] and 'posterior' in al[1]['detalle'] and 'no es válida' in al[2]['detalle']


def test_plantilla_antigua_sin_columna_de_fecha_sigue_funcionando(app, caso, tmp_path):
    pl = hacer_plantilla(tmp_path / 'vieja.xlsx', TRES); wb = openpyxl.load_workbook(pl); wb[PL.HOJA].delete_cols(4); wb.save(pl)
    assert openpyxl.load_workbook(pl)[PL.HOJA].max_column == 3 and PL.leer(pl)[0]['fecha_deposito'] == ''
    c = login(app, 'ana'); r = subir_pl(c, caso['a1'], caso['a2'], pl); assert r.status_code == 303; bid = bid_de(r)
    assert resumen(app, bid) == (3, 3, 0) and 'fecha_deposito' not in json.dumps(leer_estado(app, bid)['reclasificaciones'])


def test_qr_baneco_en_la_plantilla_no_es_banco_valido(app, caso):
    c, bid, _ = nuevo(app, caso, [('60002', 'QR BANECO', '800000099')])
    assert 'no es válido' in problemas(app, bid)[2]['motivo'] and 'BANECO' in problemas(app, bid)[2]['motivo']


def test_plantilla_generada_trae_baneco_fecha_y_formato_ambar(tmp_path):
    ruta = PL.generar_plantilla(str(tmp_path / 'p.xlsx')); wb = openpyxl.load_workbook(ruta); ws = wb[PL.HOJA]
    assert [c.value for c in ws[1]] == ['N° factura', 'Banco', 'Código de asignación', 'Fecha del depósito'] and ws['D1'].comment and 'OPCIONAL' in ws['D1'].comment.text
    assert all(ws[f'D{n}'].number_format == '@' for n in (2, 100, 201))
    assert ws.data_validations.dataValidation[0].formula1 == '"BNB,BMSC,BISA,UNION,BANECO"'
    reglas = [r.formula[0] for rs in ws.conditional_formatting for r in rs.rules]; assert any('"BANECO"' in f and '0123456789' in f for f in reglas)
    txt = ' '.join(str(r[0].value) for r in wb[PL.HOJA_AYUDA].iter_rows() if r[0].value)
    for k in ('BANECO: solo números', 'SOLO para un depósito', 'NO aparece en el reporte QR', 'Fecha del depósito', 'dd/mm/aaaa', 'QR BANECO'): assert k in txt
    est = openpyxl.load_workbook(PL.RUTA_ESTATICA)[PL.HOJA]; assert est['D1'].value == 'Fecha del depósito'


# ---------- migracion ----------
def test_migracion_v5_a_v6_con_filas_e_idempotente(tmp_path):
    ruta = str(tmp_path / 'm.db'); dbm.inicializar(ruta)
    con = sqlite3.connect(ruta)
    con.execute("INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, tarjeta_sistema, total_cierre, "
                "estado_revision, version, presentado_en, ruta) VALUES ('CC-2026-00001','2026-10-15','ana','Ana','SFC1',1000,1000,100,100,1100,'CONFORME',1,'2026-10-15 17:00:00','cierres/x/v1')")
    con.execute("INSERT INTO cierre_ci (cierre_id, factura, banco, codigo, monto, fecha_deposito) VALUES (1,'30003','BNB','ABC-1',700,NULL)")
    con.execute('ALTER TABLE cierre_ci DROP COLUMN fecha_deposito'); con.execute('PRAGMA user_version = 5'); con.commit(); con.close()
    for _ in range(2):
        dbm.inicializar(ruta); con = sqlite3.connect(ruta)
        assert 'fecha_deposito' in {r[1] for r in con.execute('PRAGMA table_info(cierre_ci)')}
        assert con.execute('PRAGMA user_version').fetchone()[0] == 6 == dbm.VERSION_ESQUEMA
        assert con.execute('select banco, codigo, monto, fecha_deposito from cierre_ci').fetchall() == [('BNB', 'ABC-1', 700.0, None)]
        assert con.execute('select count(*) from cierres').fetchone()[0] == 1; con.close()


def test_editar_y_reenviar_conserva_baneco_y_fecha_y_no_choca_con_su_propio_cierre(app, caso):
    from tests.test_v3 import revisar
    c, bid, cid = presentado(app, caso, 'ana', [('60002', 'BANECO', '7770001', '01/09/2026'), ('60003', 'BMSC', 'T98765'), ('60005', 'BISA', '4567890')])
    ing = login(app, 'ing'); revisar(ing, cid, 'rechazar', 'Falta aclarar el deposito BANECO')
    r = post(c, f'/cierres/{cid}/editar', f'/cierres/{cid}'); assert r.status_code == 303; b2 = bid_de(r)
    assert leer_estado(app, b2)['reclasificaciones']['60002'] == {'banco': 'BANECO', 'codigo': '7770001', 'fecha_deposito': '01/09/2026'}
    assert 'ya está registrado' not in pag(c, b2).text
    r = presentar(c, b2); assert r.status_code == 303 and r.headers['location'].startswith('/cierres/'), flash_de(c, b2)
    assert db(app).execute("select count(*) from cierre_ci where banco='BANECO' and codigo='7770001'").fetchone()[0] == 2        # el viejo (REEMPLAZADO) y el nuevo
    c3 = login(app, 'beto'); b3 = bid_de(subir_pl(c3, caso['a1'], caso['a2']))
    resp(c3, b3, '60002', 'BANECO', '7770001'); assert 'ya está registrado en el cierre CC-' in flash_de(c3, b3)
