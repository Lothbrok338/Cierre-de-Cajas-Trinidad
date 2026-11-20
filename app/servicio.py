"""Logica de borradores, evaluacion del cierre y presentacion (versiones inmutables)."""
import hashlib, json, os, re, shutil, uuid
from datetime import datetime

from . import db, motor
from .formato import fecha

MAX_BYTES = 10 * 1024 * 1024
COLS_A1 = ['Nro', 'Fecha', 'Número Factura', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago', 'Estado', 'Monto']
COLS_A2 = ['ID TRANSACCION', 'TITULAR ORIGEN', 'FECHA / HORA', 'MONTO']
MSG_A1 = 'El archivo no parece ser el reporte de Ingreso Diario del sistema de facturación.'
MSG_A2 = 'El archivo no parece ser el reporte de pagos QR BANECO del banco.'
ESTADO_INICIAL = 'PENDIENTE DE REVISIÓN'
ARCHIVOS_CIERRE = {  # nombre logico -> tipo mime
    'conciliacion.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'cierre.pdf': 'application/pdf', 'cierre.json': 'application/json',
    'banco_qr_original.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}


class ErrorUsuario(Exception):
    pass


def sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, 'rb') as f:
        for b in iter(lambda: f.read(65536), b''): h.update(b)
    return h.hexdigest()


def _tipo_real(path):
    with open(path, 'rb') as f: m = f.read(8)
    if m[:4] == b'PK\x03\x04': return 'xlsx'
    if m[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1': return 'xls'
    return None


def validar_a1(path):
    """Devuelve (extension_real, metadatos). Lanza ErrorUsuario si no es el reporte esperado."""
    ext = _tipo_real(path)
    if ext is None: raise ErrorUsuario(MSG_A1)
    try:
        df = motor.parse_a1(path)
        meta = motor.metadatos_a1(path)
    except Exception:
        raise ErrorUsuario(MSG_A1)
    if any(c not in df.columns for c in COLS_A1) or df.empty: raise ErrorUsuario(MSG_A1)
    return ext, meta


def validar_a2(path):
    if _tipo_real(path) != 'xlsx': raise ErrorUsuario(MSG_A2)
    try:
        df = motor.parse_a2(path)
        if any(c not in df.columns for c in COLS_A2): raise ValueError
        motor.preparar_a2(df)
    except Exception:
        raise ErrorUsuario(MSG_A2)


class Almacen:
    def __init__(self, data_dir):
        self.dir = os.path.abspath(data_dir)
        self.dbp = os.path.join(self.dir, 'cierres.db')
        os.makedirs(os.path.join(self.dir, 'borradores'), exist_ok=True)
        os.makedirs(os.path.join(self.dir, 'cierres'), exist_ok=True)
        db.inicializar(self.dbp)

    def auditar(self, usuario, accion, detalle=None): db.auditar(self.dbp, usuario, accion, detalle)
    def con(self): return db.conectar(self.dbp)

    # ---------------- borradores ----------------
    def ruta_borrador(self, bid):
        try: bid = str(uuid.UUID(str(bid)))
        except ValueError: return None
        return os.path.join(self.dir, 'borradores', bid)

    def crear_borrador(self, usuario, ruta_a1_tmp, ext_a1, ruta_a2_tmp, tarjeta, nombres):
        bid = str(uuid.uuid4()); d = self.ruta_borrador(bid); os.makedirs(d)
        a1 = os.path.join(d, f'movimiento_original.{ext_a1}'); a2 = os.path.join(d, 'banco_qr_original.xlsx')
        shutil.copyfile(ruta_a1_tmp, a1); shutil.copyfile(ruta_a2_tmp, a2)
        meta = motor.metadatos_a1(a1)
        est = {'usuario': usuario, 'tarjeta': tarjeta, 'confirmaciones': [], 'elecciones': {}, 'creado': db.ahora(),
               'archivo_movimiento': os.path.basename(a1), 'archivo_banco': 'banco_qr_original.xlsx',
               'nombres_subidos': nombres, 'caja': meta['caja'], 'presentado': None}
        self.guardar_estado(bid, est)
        return bid

    def estado(self, bid):
        d = self.ruta_borrador(bid)
        if not d or not os.path.isfile(os.path.join(d, 'estado.json')): return None
        with open(os.path.join(d, 'estado.json'), encoding='utf-8') as f: return json.load(f)

    def guardar_estado(self, bid, est):
        d = self.ruta_borrador(bid); tmp = os.path.join(d, 'estado.json.tmp')
        with open(tmp, 'w', encoding='utf-8') as f: json.dump(est, f, ensure_ascii=False, indent=1)
        os.replace(tmp, os.path.join(d, 'estado.json'))

    def conciliar(self, bid, est):
        d = self.ruta_borrador(bid)
        return motor.conciliar(os.path.join(d, est['archivo_movimiento']), os.path.join(d, est['archivo_banco']),
                               est['tarjeta'], est['confirmaciones'], est['elecciones'])

    # ---------------- consultas ----------------
    def vigente(self, fecha_cierre, usuario):
        con = self.con()
        try:
            return con.execute("SELECT * FROM cierres WHERE fecha_cierre=? AND usuario=? AND estado_revision!='REEMPLAZADO' "
                               "ORDER BY version DESC LIMIT 1", (str(fecha_cierre), usuario)).fetchone()
        finally: con.close()

    def conflictos_ids(self, res, excluir_id=None):
        """IDs TRANSACCION de este cierre ya usados en otro cierre vigente. Lista de dicts."""
        f = res['facturas']; usados = f[f['estado_conciliacion'].isin(['CONCILIADA', 'CONFIRMADA POR CAJERO'])]
        ids = {str(r['ID TRANSACCION']): str(r['Número Factura']) for _, r in usados.iterrows() if r['ID TRANSACCION']}
        if not ids: return []
        con = self.con(); out = []
        try:
            q = ','.join('?' * len(ids))
            for r in con.execute(f"SELECT m.id_transaccion, m.factura, c.correlativo, c.usuario FROM cierre_movimientos m JOIN cierres c ON c.id=m.cierre_id "
                                 f"WHERE c.estado_revision!='REEMPLAZADO' AND m.id_transaccion IN ({q})", list(ids)):
                out.append({'factura': ids[r['id_transaccion']], 'id_transaccion': r['id_transaccion'],
                            'correlativo': r['correlativo'], 'usuario': r['usuario']})
        finally: con.close()
        return out

    def conflictos_para(self, res, usuario):
        """Excluye el cierre vigente del mismo usuario y fecha (sera reemplazado)."""
        prev = self.vigente(res['totales']['fecha_cierre'], usuario)
        return [c for c in self.conflictos_ids(res) if not (prev and c['correlativo'] == prev['correlativo'])]

    def cierres_de(self, usuario=None):
        con = self.con()
        try:
            if usuario: rows = con.execute('SELECT * FROM cierres WHERE usuario=? ORDER BY id DESC', (usuario,)).fetchall()
            else: rows = con.execute('SELECT * FROM cierres ORDER BY id DESC').fetchall()
            return rows
        finally: con.close()

    def cierre(self, cid):
        con = self.con()
        try: return con.execute('SELECT * FROM cierres WHERE id=?', (cid,)).fetchone()
        finally: con.close()

    # ---------------- presentacion ----------------
    def presentar(self, bid, est, usuario, nombre, motivo_reemplazo=None):
        from . import pdf
        res = self.conciliar(bid, est)
        ev = evaluar(res)
        if not ev['presentable']: raise ErrorUsuario('No se puede presentar: ' + ' / '.join(ev['razones']))
        if est.get('presentado'): raise ErrorUsuario('Este borrador ya fue presentado.')
        t = res['totales']; fc = str(t['fecha_cierre'])
        con = self.con(); destino = None
        try:
            con.execute('BEGIN IMMEDIATE')
            prev = con.execute("SELECT * FROM cierres WHERE fecha_cierre=? AND usuario=? AND estado_revision!='REEMPLAZADO' "
                               "ORDER BY version DESC LIMIT 1", (fc, usuario)).fetchone()
            if prev and not (motivo_reemplazo or '').strip():
                raise ErrorUsuario(f"Ya presentó un cierre para el {fecha(fc)} ({prev['correlativo']}). Indique el motivo del reemplazo.")
            confl = [c for c in self.conflictos_ids(res) if not (prev and c['correlativo'] == prev['correlativo'])]
            if confl:
                c = confl[0]
                raise ErrorUsuario(f"La factura {c['factura']} usa el ID de transacción {c['id_transaccion']}, que ya está en el cierre "
                                   f"{c['correlativo']} (cajero {c['usuario']}). No se puede presentar.")
            ahora = datetime.now(); presentado_en = ahora.strftime('%Y-%m-%d %H:%M:%S')
            version = (prev['version'] + 1) if prev else 1
            if prev:  # versiones previas (incluidas ya reemplazadas) -> max
                version = con.execute('SELECT MAX(version)+1 FROM cierres WHERE fecha_cierre=? AND usuario=?', (fc, usuario)).fetchone()[0]
            ult = con.execute("SELECT correlativo FROM cierres WHERE correlativo LIKE ? ORDER BY correlativo DESC LIMIT 1",
                              (f'CC-{ahora.year}-%',)).fetchone()
            corr = f"CC-{ahora.year}-{(int(ult[0].split('-')[-1]) + 1) if ult else 1:05d}"
            seguro = re.sub(r'[^A-Za-z0-9._-]', '_', usuario)
            rel = os.path.join('cierres', fc[:4], fc[5:7], fc, seguro, f'v{version}')
            destino = os.path.join(self.dir, rel)
            if os.path.exists(destino): raise ErrorUsuario('Conflicto al crear la versión del cierre; intente de nuevo.')
            os.makedirs(destino)
            d = self.ruta_borrador(bid)
            mov = f"movimiento_original.{est['archivo_movimiento'].rsplit('.', 1)[1]}"
            shutil.copyfile(os.path.join(d, est['archivo_movimiento']), os.path.join(destino, mov))
            shutil.copyfile(os.path.join(d, est['archivo_banco']), os.path.join(destino, 'banco_qr_original.xlsx'))
            h_mov = sha256(os.path.join(destino, mov)); h_ban = sha256(os.path.join(destino, 'banco_qr_original.xlsx'))
            if h_mov != sha256(os.path.join(d, est['archivo_movimiento'])): raise ErrorUsuario('Error de copia de archivos.')
            motor.exportar(res, os.path.join(destino, 'conciliacion.xlsx'))
            info = {'correlativo': corr, 'fecha_cierre': fc, 'caja': est.get('caja', ''), 'usuario': usuario, 'nombre': nombre,
                    'presentado_en': presentado_en, 'version': version, 'sha256_movimiento': h_mov, 'sha256_banco': h_ban}
            pdf.generar(res, info, os.path.join(destino, 'cierre.pdf'))
            cj = {**info, 'archivo_movimiento': mov, 'estado_revision': ESTADO_INICIAL, 'reemplaza_a': prev['correlativo'] if prev else None,
                  'motivo_reemplazo': (motivo_reemplazo or '').strip() or None, 'estado_qr': t['Estado QR'],
                  'conceptos': conceptos(t), 'totales': t, 'confirmaciones': est['confirmaciones'], 'elecciones': est['elecciones'],
                  'tarjeta_declarada_por_cajero': est['tarjeta'], 'declaracion': 'El cajero declara conforme el presente cierre.',
                  'facturas': registros(res['facturas']),
                  'sha256_archivos': {n: sha256(os.path.join(destino, n)) for n in ('conciliacion.xlsx', 'cierre.pdf')}}
            with open(os.path.join(destino, 'cierre.json'), 'w', encoding='utf-8') as f:
                json.dump(cj, f, ensure_ascii=False, indent=1, default=_json)
            for n in os.listdir(destino): os.chmod(os.path.join(destino, n), 0o444)
            qr_conc = t['QR conciliado automático (CONCILIADA)'] + t['QR confirmado por cajero']
            cur = con.execute('INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, '
                              'tarjeta_sistema, total_cierre, estado_revision, version, reemplaza_a, motivo_reemplazo, presentado_en, ruta, '
                              'sha256_movimiento, sha256_banco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                              (corr, fc, usuario, nombre, est.get('caja', ''), t['QR registrado'], qr_conc, t['Tarjeta declarada'],
                               t['Tarjeta según sistema'], t['Total cierre'], ESTADO_INICIAL, version, prev['id'] if prev else None,
                               cj['motivo_reemplazo'], presentado_en, rel, h_mov, h_ban))
            cid = cur.lastrowid
            f = res['facturas']
            for _, r in f[f['estado_conciliacion'].isin(['CONCILIADA', 'CONFIRMADA POR CAJERO'])].iterrows():
                con.execute('INSERT INTO cierre_movimientos VALUES (?,?,?,?)', (cid, str(r['ID TRANSACCION']), str(r['Número Factura']), float(r['Monto'])))
            if prev: con.execute("UPDATE cierres SET estado_revision='REEMPLAZADO' WHERE id=?", (prev['id'],))
            con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                        (presentado_en, usuario, 'presentar_cierre', json.dumps({'correlativo': corr, 'version': version, 'borrador': bid,
                         'reemplaza': prev['correlativo'] if prev else None, 'motivo': cj['motivo_reemplazo']}, ensure_ascii=False)))
            con.execute('COMMIT')
        except BaseException:
            try: con.execute('ROLLBACK')
            except Exception: pass
            if destino and os.path.isdir(destino):
                for n in os.listdir(destino): os.chmod(os.path.join(destino, n), 0o644)
                shutil.rmtree(destino, ignore_errors=True)
            raise
        finally:
            con.close()
        est['presentado'] = corr; self.guardar_estado(bid, est)
        return cid, corr


def registros(df):
    import pandas as pd
    d = df.copy()
    for c in d.columns:
        if pd.api.types.is_datetime64_any_dtype(d[c]): d[c] = d[c].dt.strftime('%Y-%m-%d %H:%M:%S').where(d[c].notna(), None)
    return json.loads(d.to_json(orient='records', force_ascii=False))


def _json(o):
    import numpy as np, pandas as pd
    if isinstance(o, (np.integer,)): return int(o)
    if isinstance(o, (np.floating,)): return float(o)
    if isinstance(o, (np.bool_,)): return bool(o)
    if o is pd.NaT or (isinstance(o, float) and o != o): return None
    if hasattr(o, 'isoformat'): return o.isoformat()
    return str(o)


def conceptos(t):
    """Cuadro principal del cierre como lista de conceptos; agregar filas aqui para nuevos conceptos."""
    qr_conc = t['QR conciliado automático (CONCILIADA)'] + t['QR confirmado por cajero']
    dq = round(qr_conc - t['QR registrado'], 2)
    dt = round(t['Tarjeta declarada'] - t['Tarjeta según sistema'], 2)
    return [
        {'concepto': 'QR BANECO', 'segun_sistema': t['QR registrado'], 'segun_respaldo': qr_conc, 'diferencia': dq,
         'estado': t['Estado QR'], 'nota': 'Respaldo = conciliado con el reporte del banco'},
        {'concepto': 'Tarjetas', 'segun_sistema': t['Tarjeta según sistema'], 'segun_respaldo': t['Tarjeta declarada'], 'diferencia': dt,
         'estado': 'CUADRA' if abs(dt) < 0.005 else 'NO CUADRA', 'nota': 'Pendiente de verificación por Ingresos'},
    ]


def total_conceptos(lista):
    s = round(sum(c['segun_sistema'] for c in lista), 2); r = round(sum(c['segun_respaldo'] for c in lista), 2)
    return {'concepto': 'TOTAL', 'segun_sistema': s, 'segun_respaldo': r, 'diferencia': round(r - s, 2),
            'estado': 'CUADRA' if all(c['estado'] == 'CUADRA' for c in lista) else 'NO CUADRA', 'nota': ''}


def evaluar(res):
    """Regla de presentacion: Estado QR == CUADRA y diferencia de tarjetas == 0."""
    t = res['totales']; razones = []
    dt = round(t['Diferencia tarjeta (declarada - sistema)'], 2)
    if t['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN':
        razones.append(f"Hay facturas QR por confirmar o asignar (diferencia QR Bs {t['Diferencia QR']:,.2f}).")
    elif t['Estado QR'] != 'CUADRA':
        razones.append(f"El QR no cuadra: faltan Bs {t['Diferencia QR']:,.2f} sin respaldo del banco.")
    if abs(dt) >= 0.005:
        razones.append(f"Las tarjetas no cuadran: declaró Bs {t['Tarjeta declarada']:,.2f} y el sistema registra Bs {t['Tarjeta según sistema']:,.2f} (diferencia Bs {dt:,.2f}).")
    if t['fecha_cierre'] is None: razones.append('No hay facturas válidas en el movimiento del cajero.')
    if t['Estado QR'] == 'NO CUADRA' or (t['Estado QR'] == 'CUADRA' and abs(dt) >= 0.005): nivel = 'rojo'
    elif t['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN': nivel = 'ambar'
    else: nivel = 'verde'
    return {'presentable': not razones, 'razones': razones, 'nivel': nivel}


def candidatos(res, nf):
    f = res['facturas'].set_index('Número Factura')
    if nf not in f.index: return []
    m = f.loc[nf, 'Monto']; b = res['banco']
    c = b[(abs(b['MONTO'] - m) < 1e-9) & (b['estado_banco'].str.startswith('AJENO') | b['estado_banco'].str.startswith('AMBIGUO'))]
    return [{'id': r['ID TRANSACCION'], 'titular': r['TITULAR ORIGEN'], 'hora': r['FECHA / HORA']} for _, r in c.iterrows()]
