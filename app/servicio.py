"""Logica de borradores, evaluacion del cierre y presentacion (versiones inmutables)."""
import hashlib, json, os, re, shutil, uuid
from datetime import datetime

from . import catalogo, db, motor
from .formato import fecha

MAX_BYTES = 10 * 1024 * 1024
COLS_A1 = ['Nro', 'Fecha', 'Número Factura', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago', 'Estado', 'Monto']
COLS_A2 = ['ID TRANSACCION', 'TITULAR ORIGEN', 'FECHA / HORA', 'MONTO']
MSG_A1 = 'El archivo no parece ser el reporte de Ingreso Diario del sistema de facturación.'
MSG_A2 = 'El archivo no parece ser el reporte de pagos QR BANECO del banco.'
ESTADO_INICIAL = 'PENDIENTE DE REVISIÓN'
ESTADO_CONFORME = 'CONFORME'
ESTADO_RECHAZADO = 'RECHAZADO'
ESTADO_REEMPLAZADO = 'REEMPLAZADO'
ESTADOS_REVISION = (ESTADO_INICIAL, ESTADO_CONFORME, ESTADO_RECHAZADO, ESTADO_REEMPLAZADO)
OBSERVADAS = ('POR CONFIRMAR', 'SIN RESPALDO', 'AMBIGUO')
MIN_OBSERVACION = 10
ESTADO_CI = 'PREVALIDADO POR INGRESOS'
ARCHIVOS_CIERRE = {  # nombre logico -> tipo mime
    'cierre_revision.pdf': 'application/pdf',
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
        est = {'usuario': usuario, 'tarjeta': tarjeta, 'confirmaciones': [], 'elecciones': {}, 'reclasificaciones': {},
               'respaldos_manuales': [], 'origen': None, 'creado': db.ahora(),
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
                               est['tarjeta'], est['confirmaciones'], est['elecciones'],
                               reclasificaciones=est.get('reclasificaciones') or None, respaldos_manuales=est.get('respaldos_manuales') or None)

    def crear_borrador_desde(self, c, usuario):
        """Borrador para 'Editar y reenviar' a partir de un cierre RECHAZADO: copia archivos originales y decisiones.
        Si ya hay un borrador abierto para ese cierre, lo reutiliza."""
        base = os.path.join(self.dir, 'borradores')
        for nombre in sorted(os.listdir(base)):
            est = self.estado(nombre)
            if est and not est.get('presentado') and est.get('usuario') == usuario and (est.get('origen') or {}).get('cierre_id') == c['id']:
                return nombre
        dc = os.path.join(self.dir, c['ruta'])
        with open(os.path.join(dc, 'cierre.json'), encoding='utf-8') as f: cj = json.load(f)
        mov = cj['archivo_movimiento']; ext = mov.rsplit('.', 1)[1]
        bid = str(uuid.uuid4()); d = self.ruta_borrador(bid); os.makedirs(d)
        shutil.copyfile(os.path.join(dc, mov), os.path.join(d, f'movimiento_original.{ext}'))
        shutil.copyfile(os.path.join(dc, 'banco_qr_original.xlsx'), os.path.join(d, 'banco_qr_original.xlsx'))
        est = {'usuario': usuario, 'tarjeta': cj['tarjeta_declarada_por_cajero'], 'confirmaciones': list(cj.get('confirmaciones') or []),
               'elecciones': dict(cj.get('elecciones') or {}), 'reclasificaciones': dict(cj.get('reclasificaciones') or {}),
               'respaldos_manuales': list(cj.get('respaldos_manuales') or []), 'creado': db.ahora(),
               'origen': {'cierre_id': c['id'], 'correlativo': c['correlativo'], 'fecha_cierre': c['fecha_cierre'], 'observacion': c['observacion_revision']},
               'archivo_movimiento': f'movimiento_original.{ext}', 'archivo_banco': 'banco_qr_original.xlsx',
               'nombres_subidos': {'movimiento': mov, 'banco': 'banco_qr_original.xlsx'}, 'caja': cj.get('caja', ''), 'presentado': None}
        self.guardar_estado(bid, est)
        return bid

    def reemplazar_archivos(self, bid, est, ruta_a1=None, ext_a1=None, nombre_a1=None, ruta_a2=None, nombre_a2=None):
        """Reemplaza uno o ambos archivos del borrador (ya validados). Devuelve el estado actualizado."""
        d = self.ruta_borrador(bid)
        if ruta_a1:
            nuevo = f'movimiento_original.{ext_a1}'
            if est['archivo_movimiento'] != nuevo and os.path.isfile(os.path.join(d, est['archivo_movimiento'])): os.remove(os.path.join(d, est['archivo_movimiento']))
            shutil.copyfile(ruta_a1, os.path.join(d, nuevo)); est['archivo_movimiento'] = nuevo; est['nombres_subidos']['movimiento'] = nombre_a1
            est['caja'] = motor.metadatos_a1(os.path.join(d, nuevo))['caja']
        if ruta_a2:
            shutil.copyfile(ruta_a2, os.path.join(d, est['archivo_banco'])); est['nombres_subidos']['banco'] = nombre_a2
        self.guardar_estado(bid, est)
        return est

    # ---------------- respaldo manual ----------------
    def registrar_respaldo(self, bid, est, usuario, nf, banco, codigo):
        """Valida y aplica un respaldo manual a una factura OBSERVADA. Devuelve (est_nuevo, resumen).
        Lanza ErrorUsuario con un mensaje claro si no es valido (no modifica `est`)."""
        nf = (nf or '').strip(); banco = (banco or '').strip(); codigo = (codigo or '').strip()
        if banco not in catalogo.BANCOS: raise ErrorUsuario('Seleccione el banco.')
        res = self.conciliar(bid, est); f = res['facturas'].set_index('Número Factura')
        if nf not in f.index: raise ErrorUsuario('No se encontró la factura.')
        estado = f.loc[nf, 'estado_conciliacion']
        if estado not in OBSERVADAS:
            raise ErrorUsuario(f'La factura {nf} no está en Observadas (está {estado}). Para cambiarla, primero deshaga su respaldo actual.')
        nuevo = json.loads(json.dumps(est))
        for k, v in (('reclasificaciones', {}), ('respaldos_manuales', [])): nuevo.setdefault(k, v)
        if catalogo.es_qr(banco):
            if not codigo: raise ErrorUsuario('Escriba el ID de transacción del reporte QR.')
            if estado == 'POR CONFIRMAR':
                raise ErrorUsuario(f'La factura {nf} ya tiene un movimiento del banco sugerido: use «Sí, corresponde», o reclasifíquela con otro banco.')
            nuevo['elecciones'][nf] = codigo
            if nf not in nuevo['respaldos_manuales']: nuevo['respaldos_manuales'].append(nf)
            r2 = self.conciliar(bid, nuevo); f2 = r2['facturas'].set_index('Número Factura')
            if f2.loc[nf, 'estado_conciliacion'] != 'CONFIRMADA POR CAJERO':
                al = [a['detalle'] for a in r2['alertas'].to_dict('records') if a['referencia'] == nf and 'Elección no aplicada' in a['detalle']]
                motivo = al[0].split('): ', 1)[-1] if al else 'ID de transacción no válido'
                raise ErrorUsuario(f'El ID de transacción {codigo} no es válido para la factura {nf}: {motivo}.')
            c = [x for x in self.conflictos_para(r2, usuario) if x['factura'] == nf]
            if c: raise ErrorUsuario(f"El ID de transacción {codigo} ya está en el cierre {c[0]['correlativo']} (cajero {c[0]['usuario']}).")
            return nuevo, {'banco': banco, 'codigo': codigo, 'resultado': 'CONFIRMADA POR CAJERO (RESPALDO MANUAL)'}
        ok, msg, cn = catalogo.validar_codigo(banco, codigo)
        if not ok: raise ErrorUsuario(msg)
        for otra, d in nuevo['reclasificaciones'].items():
            if otra != nf and d['banco'] == banco and catalogo.normalizar_codigo(d['codigo']) == cn:
                raise ErrorUsuario(f'El código {cn} de {banco} ya está registrado en la factura {otra} de este cierre.')
        fc = res['totales']['fecha_cierre']; prev = self.vigente(fc, usuario) if fc else None
        c = self.conflictos_ci_banco(banco, cn, prev['correlativo'] if prev else None)
        if c: raise ErrorUsuario(f"El código {cn} de {banco} ya está registrado en el cierre {c['correlativo']} (cajero {c['usuario']}).")
        nuevo['confirmaciones'] = [x for x in nuevo['confirmaciones'] if x != nf]; nuevo['elecciones'].pop(nf, None)
        nuevo['respaldos_manuales'] = [x for x in nuevo['respaldos_manuales'] if x != nf]
        nuevo['reclasificaciones'][nf] = {'banco': banco, 'codigo': cn}
        return nuevo, {'banco': banco, 'codigo': cn, 'resultado': 'COMUNICACIÓN INTERNA'}

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

    def conflictos_ci_banco(self, banco, codigo_norm, excluir_correlativo=None):
        """Primer cierre vigente (cualquier usuario) que ya usa el par (banco, codigo) de una C.I., o None."""
        con = self.con()
        try:
            for r in con.execute("SELECT c.correlativo, c.usuario FROM cierre_ci k JOIN cierres c ON c.id=k.cierre_id "
                                 "WHERE c.estado_revision!='REEMPLAZADO' AND k.banco=? AND k.codigo=?", (banco, codigo_norm)):
                if r['correlativo'] != excluir_correlativo: return {'correlativo': r['correlativo'], 'usuario': r['usuario']}
        finally: con.close()
        return None

    def conflictos_ci_para(self, res, usuario):
        """C.I. de este borrador cuyo (banco, codigo) ya esta en otro cierre vigente (se excluye el que sera reemplazado)."""
        fc = res['totales']['fecha_cierre']; prev = self.vigente(fc, usuario) if fc else None; out = []
        for r in res.get('ci', []):
            c = self.conflictos_ci_banco(r['banco'], catalogo.normalizar_codigo(r['codigo']), prev['correlativo'] if prev else None)
            if c: out.append({'factura': r['factura'], 'banco': r['banco'], 'codigo': r['codigo'], **c})
        return out

    def cierres_de(self, usuario=None, estado=None, desde=None, hasta=None):
        q = 'SELECT * FROM cierres WHERE 1=1'; a = []
        if usuario: q += ' AND usuario=?'; a.append(usuario)
        if estado: q += ' AND estado_revision=?'; a.append(estado)
        if desde: q += ' AND fecha_cierre>=?'; a.append(desde)
        if hasta: q += ' AND fecha_cierre<=?'; a.append(hasta)
        con = self.con()
        try: return con.execute(q + ' ORDER BY id DESC', a).fetchall()
        finally: con.close()

    def cierre(self, cid):
        con = self.con()
        try: return con.execute('SELECT * FROM cierres WHERE id=?', (cid,)).fetchone()
        finally: con.close()

    # ---------------- presentacion ----------------
    def presentar(self, bid, est, usuario, nombre):
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
            origen = est.get('origen')
            msg = regla_vigente(prev, origen, fc)
            if msg: raise ErrorUsuario(msg)
            motivo_reemplazo = f"Corrección de observación de Ingresos: {prev['observacion_revision']}" if origen else None
            confl = [c for c in self.conflictos_ids(res) if not (prev and c['correlativo'] == prev['correlativo'])]
            if confl:
                c = confl[0]
                raise ErrorUsuario(f"La factura {c['factura']} usa el ID de transacción {c['id_transaccion']}, que ya está en el cierre "
                                   f"{c['correlativo']} (cajero {c['usuario']}). No se puede presentar.")
            confl_ci = self.conflictos_ci_para(res, usuario)
            if confl_ci:
                c = confl_ci[0]
                raise ErrorUsuario(f"La comunicación interna de la factura {c['factura']} ({c['banco']} {c['codigo']}) ya está registrada en el cierre "
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
                    'presentado_en': presentado_en, 'version': version, 'sha256_movimiento': h_mov, 'sha256_banco': h_ban,
                    'estado_revision': ESTADO_INICIAL, 'reemplaza_a': prev['correlativo'] if prev else None,
                    'motivo_reemplazo': (motivo_reemplazo or '').strip() or None}
            pdf.generar(res, info, os.path.join(destino, 'cierre.pdf'))
            cj = {**info, 'archivo_movimiento': mov, 'estado_revision': ESTADO_INICIAL, 'reemplaza_a': prev['correlativo'] if prev else None,
                  'motivo_reemplazo': (motivo_reemplazo or '').strip() or None, 'estado_qr': t['Estado QR'],
                  'conceptos': conceptos(t), 'totales': t, 'confirmaciones': est['confirmaciones'], 'elecciones': est['elecciones'],
                  'reclasificaciones': est.get('reclasificaciones') or {}, 'respaldos_manuales': est.get('respaldos_manuales') or [],
                  'comunicaciones_internas': res.get('ci', []),
                  'tarjeta_declarada_por_cajero': est['tarjeta'], 'declaracion': 'El cajero declara conforme el presente cierre.',
                  'facturas': registros(res['facturas']),
                  'sha256_archivos': {n: sha256(os.path.join(destino, n)) for n in ('conciliacion.xlsx', 'cierre.pdf')}}
            with open(os.path.join(destino, 'cierre.json'), 'w', encoding='utf-8') as f:
                json.dump(cj, f, ensure_ascii=False, indent=1, default=_json)
            for n in os.listdir(destino): os.chmod(os.path.join(destino, n), 0o444)
            qr_conc = t['QR conciliado automático (CONCILIADA)'] + t['QR confirmado por cajero']
            cur = con.execute('INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, '
                              'tarjeta_sistema, total_cierre, estado_revision, version, reemplaza_a, motivo_reemplazo, presentado_en, ruta, '
                              'sha256_movimiento, sha256_banco, ci_total) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                              (corr, fc, usuario, nombre, est.get('caja', ''), t['QR registrado'], qr_conc, t['Tarjeta declarada'],
                               t['Tarjeta según sistema'], t['Total cierre'], ESTADO_INICIAL, version, prev['id'] if prev else None,
                               cj['motivo_reemplazo'], presentado_en, rel, h_mov, h_ban, t['C.I. según sistema']))
            cid = cur.lastrowid
            f = res['facturas']
            for _, r in f[f['estado_conciliacion'].isin(['CONCILIADA', 'CONFIRMADA POR CAJERO'])].iterrows():
                con.execute('INSERT INTO cierre_movimientos VALUES (?,?,?,?)', (cid, str(r['ID TRANSACCION']), str(r['Número Factura']), float(r['Monto'])))
            for r in res.get('ci', []):
                con.execute('INSERT INTO cierre_ci VALUES (?,?,?,?,?)', (cid, r['factura'], r['banco'], catalogo.normalizar_codigo(r['codigo']), r['monto']))
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

    # ---------------- revision por Ingresos ----------------
    def revisar(self, cid, revisor, decision, observacion=''):
        """Ingresos aprueba (CONFORME) o rechaza (RECHAZADO, observacion >= 10 caracteres) un cierre PENDIENTE DE REVISIÓN."""
        if decision not in (ESTADO_CONFORME, ESTADO_RECHAZADO): raise ErrorUsuario('Decisión no válida.')
        obs = ' '.join((observacion or '').split()) if decision == ESTADO_RECHAZADO else (observacion or '').strip()
        if decision == ESTADO_RECHAZADO and len(obs) < MIN_OBSERVACION:
            raise ErrorUsuario(f'Para rechazar escriba una observación de al menos {MIN_OBSERVACION} caracteres.')
        con = self.con(); momento = db.ahora()
        try:
            con.execute('BEGIN IMMEDIATE')
            c = con.execute('SELECT * FROM cierres WHERE id=?', (cid,)).fetchone()
            if c is None: raise ErrorUsuario('No se encontró el cierre.')
            if c['estado_revision'] != ESTADO_INICIAL:
                raise ErrorUsuario(f"El cierre {c['correlativo']} está {c['estado_revision']}: solo se puede revisar un cierre PENDIENTE DE REVISIÓN.")
            con.execute('UPDATE cierres SET estado_revision=?, revisado_por=?, revisado_en=?, observacion_revision=? WHERE id=?',
                        (decision, revisor, momento, obs or None, cid))
            con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                        (momento, revisor, 'cierre_aprobado' if decision == ESTADO_CONFORME else 'cierre_rechazado',
                         json.dumps({'correlativo': c['correlativo'], 'cierre': cid, 'estado_anterior': c['estado_revision'], 'observacion': obs or None}, ensure_ascii=False)))
            con.execute('COMMIT')
        except BaseException:
            try: con.execute('ROLLBACK')
            except Exception: pass
            raise
        finally: con.close()
        try: self.regenerar_pdf(cid)
        except Exception as e:   # la revision ya quedo registrada; el PDF original sigue disponible
            self.auditar(revisor, 'pdf_revision_error', {'cierre': cid, 'error': str(e)[:200]})
        return self.cierre(cid)

    def regenerar_pdf(self, cid):
        """PDF con el estado de revision (cierre_revision.pdf). El cierre.pdf original, inmutable, no se toca."""
        from . import pdf
        c = self.cierre(cid); d = os.path.join(self.dir, c['ruta'])
        with open(os.path.join(d, 'cierre.json'), encoding='utf-8') as f: cj = json.load(f)
        res = motor.conciliar(os.path.join(d, cj['archivo_movimiento']), os.path.join(d, 'banco_qr_original.xlsx'), cj['tarjeta_declarada_por_cajero'],
                              cj['confirmaciones'], cj['elecciones'], reclasificaciones=cj.get('reclasificaciones') or None,
                              respaldos_manuales=cj.get('respaldos_manuales') or None)
        info = {k: cj.get(k) for k in ('correlativo', 'fecha_cierre', 'caja', 'usuario', 'nombre', 'presentado_en', 'version',
                                       'sha256_movimiento', 'sha256_banco', 'motivo_reemplazo')}
        info.update(estado_revision=c['estado_revision'], revisado_por=c['revisado_por'], revisado_en=c['revisado_en'],
                    observacion_revision=c['observacion_revision'], reemplaza_a=cj.get('reemplaza_a'))
        destino = os.path.join(d, 'cierre_revision.pdf')
        if os.path.exists(destino): os.chmod(destino, 0o644); os.remove(destino)
        pdf.generar(res, info, destino); os.chmod(destino, 0o444)


def regla_vigente(prev, origen, fc):
    """Mensaje de bloqueo (o None) segun el cierre vigente de la misma fecha y usuario.
    Solo se puede crear otro cierre para esa fecha corrigiendo uno RECHAZADO (Editar y reenviar)."""
    if origen:
        if str(origen['fecha_cierre']) != str(fc):
            return f"La fecha de los archivos ({fecha(fc)}) debe ser la misma del cierre rechazado ({fecha(origen['fecha_cierre'])})."
        if not prev or prev['id'] != origen['cierre_id'] or prev['estado_revision'] != ESTADO_RECHAZADO:
            return f"El cierre {origen['correlativo']} ya no está RECHAZADO o ya fue reemplazado; no se puede reenviar esta corrección."
        return None
    if prev:
        if prev['estado_revision'] == ESTADO_RECHAZADO:
            return (f"El cierre {prev['correlativo']} del {fecha(fc)} fue RECHAZADO por Ingresos. Para corregirlo use «Editar y reenviar» "
                    f"desde ese cierre; no se puede crear uno nuevo para la misma fecha.")
        return (f"Ya existe el cierre {prev['correlativo']} para el {fecha(fc)} (estado {prev['estado_revision']}). "
                f"No se puede crear otro para la misma fecha.")
    return None


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
    lista = [
        {'concepto': 'QR BANECO', 'segun_sistema': t['QR registrado'], 'segun_respaldo': qr_conc, 'diferencia': dq,
         'estado': t['Estado QR'], 'nota': 'Respaldo = conciliado con el reporte del banco'}]
    if t.get('n_COMUNICACIÓN INTERNA'):
        ci_s = t['C.I. según sistema']; ci_r = t['C.I. con código de asignación']; dci = round(ci_r - ci_s, 2)
        lista.append({'concepto': 'Comunicaciones internas', 'segun_sistema': ci_s, 'segun_respaldo': ci_r, 'diferencia': dci,
                      'estado': ESTADO_CI if abs(dci) < 0.005 else 'SIN CÓDIGO',
                      'nota': 'Depósitos a otras cuentas, prevalidados por Ingresos (sin cruce bancario): ' + (t.get('Detalle C.I.') or '')})
    return lista + [
        {'concepto': 'Tarjetas', 'segun_sistema': t['Tarjeta según sistema'], 'segun_respaldo': t['Tarjeta declarada'], 'diferencia': dt,
         'estado': 'CUADRA' if abs(dt) < 0.005 else 'NO CUADRA', 'nota': 'Pendiente de verificación por Ingresos'},
    ]


def total_conceptos(lista):
    s = round(sum(c['segun_sistema'] for c in lista), 2); r = round(sum(c['segun_respaldo'] for c in lista), 2)
    return {'concepto': 'TOTAL', 'segun_sistema': s, 'segun_respaldo': r, 'diferencia': round(r - s, 2),
            'estado': 'CUADRA' if all(c['estado'] in ('CUADRA', ESTADO_CI) for c in lista) else 'NO CUADRA', 'nota': ''}


def evaluar(res):
    """Regla de presentacion: Estado QR == CUADRA (ya sin las C.I. reclasificadas), diferencia de tarjetas == 0 y todas las C.I. con codigo valido."""
    t = res['totales']; razones = []
    dt = round(t['Diferencia tarjeta (declarada - sistema)'], 2)
    if t['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN':
        razones.append(f"Hay facturas QR por confirmar o asignar (diferencia QR Bs {t['Diferencia QR']:,.2f}).")
    elif t['Estado QR'] != 'CUADRA':
        razones.append(f"El QR no cuadra: faltan Bs {t['Diferencia QR']:,.2f} sin respaldo del banco.")
    if abs(dt) >= 0.005:
        razones.append(f"Las tarjetas no cuadran: declaró Bs {t['Tarjeta declarada']:,.2f} y el sistema registra Bs {t['Tarjeta según sistema']:,.2f} (diferencia Bs {dt:,.2f}).")
    if t['fecha_cierre'] is None: razones.append('No hay facturas válidas en el movimiento del cajero.')
    vistos = {}
    for r in res.get('ci', []):
        ok, msg, cn = catalogo.validar_codigo(r['banco'], r['codigo'])
        if not ok: razones.append(f"Comunicación interna de la factura {r['factura']}: {msg}")
        elif (r['banco'], cn) in vistos:
            razones.append(f"El código {cn} de {r['banco']} está repetido en las facturas {vistos[(r['banco'], cn)]} y {r['factura']}.")
        else: vistos[(r['banco'], cn)] = r['factura']
    if t['Estado QR'] == 'NO CUADRA' or (t['Estado QR'] == 'CUADRA' and abs(dt) >= 0.005): nivel = 'rojo'
    elif t['Estado QR'] == 'PENDIENTE DE CONFIRMACIÓN': nivel = 'ambar'
    else: nivel = 'verde'
    if razones and nivel == 'verde': nivel = 'rojo'
    return {'presentable': not razones, 'razones': razones, 'nivel': nivel}


def candidatos(res, nf):
    f = res['facturas'].set_index('Número Factura')
    if nf not in f.index: return []
    m = f.loc[nf, 'Monto']; b = res['banco']
    c = b[(abs(b['MONTO'] - m) < 1e-9) & (b['estado_banco'].str.startswith('AJENO') | b['estado_banco'].str.startswith('AMBIGUO'))]
    return [{'id': r['ID TRANSACCION'], 'titular': r['TITULAR ORIGEN'], 'hora': r['FECHA / HORA']} for _, r in c.iterrows()]
