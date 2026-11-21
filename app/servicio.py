"""Logica de borradores, evaluacion del cierre y presentacion (versiones inmutables)."""
import hashlib, json, os, re, shutil, sqlite3, uuid
from datetime import datetime

from . import auth, catalogo, db, motor, plantilla
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
    'comunicaciones_internas_original.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}
ARCH_PLANTILLA = 'comunicaciones_internas_original.xlsx'


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


def validar_plantilla(path):
    """Valida que el archivo sea una plantilla legible (.xlsx con los encabezados). Devuelve las filas leidas. Lanza ErrorUsuario si no."""
    if _tipo_real(path) != 'xlsx': raise ErrorUsuario('Plantilla de comunicaciones internas: el archivo debe ser un .xlsx válido. Descargue la plantilla desde «Nuevo cierre».')
    try: return plantilla.leer(path)
    except ValueError as e: raise ErrorUsuario(str(e))


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

    def crear_borrador(self, usuario, ruta_a1_tmp, ext_a1, ruta_a2_tmp, tarjeta, nombres, ruta_plantilla_tmp=None):
        bid = str(uuid.uuid4()); d = self.ruta_borrador(bid); os.makedirs(d)
        a1 = os.path.join(d, f'movimiento_original.{ext_a1}'); a2 = os.path.join(d, 'banco_qr_original.xlsx')
        shutil.copyfile(ruta_a1_tmp, a1); shutil.copyfile(ruta_a2_tmp, a2)
        if ruta_plantilla_tmp: shutil.copyfile(ruta_plantilla_tmp, os.path.join(d, ARCH_PLANTILLA))
        meta = motor.metadatos_a1(a1)
        est = {'usuario': usuario, 'tarjeta': tarjeta, 'confirmaciones': [], 'elecciones': {}, 'reclasificaciones': {},
               'respaldos_manuales': [], 'rechazos': {}, 'origen': None, 'creado': db.ahora(),
               'archivo_movimiento': os.path.basename(a1), 'archivo_banco': 'banco_qr_original.xlsx',
               'nombres_subidos': nombres, 'caja': meta['caja'], 'presentado': None}
        if ruta_plantilla_tmp: est.update(plantilla={'archivo': ARCH_PLANTILLA, 'nombre': nombres.get('plantilla')}, plantilla_origen=[], plantilla_descartes=[])
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
                               reclasificaciones=est.get('reclasificaciones') or None, respaldos_manuales=est.get('respaldos_manuales') or None,
                               rechazos=est.get('rechazos') or None)

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
               'respaldos_manuales': list(cj.get('respaldos_manuales') or []),
               'rechazos': {k: list(v) for k, v in (cj.get('rechazos') or {}).items()}, 'creado': db.ahora(),
               'origen': {'cierre_id': c['id'], 'correlativo': c['correlativo'], 'fecha_cierre': c['fecha_cierre'], 'observacion': c['observacion_revision']},
               'archivo_movimiento': f'movimiento_original.{ext}', 'archivo_banco': 'banco_qr_original.xlsx',
               'nombres_subidos': {'movimiento': mov, 'banco': 'banco_qr_original.xlsx'}, 'caja': cj.get('caja', ''), 'presentado': None}
        if cj.get('plantilla') and os.path.isfile(os.path.join(dc, ARCH_PLANTILLA)):
            shutil.copyfile(os.path.join(dc, ARCH_PLANTILLA), os.path.join(d, ARCH_PLANTILLA))
            est['nombres_subidos']['plantilla'] = ARCH_PLANTILLA
            est.update(plantilla=dict(cj['plantilla'], archivo=ARCH_PLANTILLA), plantilla_origen=list(cj.get('plantilla_origen') or []),
                       plantilla_descartes=list(cj.get('plantilla_descartes') or []))
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

    def reemplazar_plantilla(self, bid, est, usuario, ruta_tmp, nombre):
        """Sube/reemplaza la plantilla de comunicaciones internas del borrador (ya validada) y la vuelve a aplicar desde cero
        (las C.I. que vinieron de la plantilla anterior se quitan y se reinician los descartes). Devuelve (est_nuevo, resumen)."""
        d = self.ruta_borrador(bid)
        shutil.copyfile(ruta_tmp, os.path.join(d, ARCH_PLANTILLA))
        nuevo = json.loads(json.dumps(est)); nuevo.setdefault('nombres_subidos', {})['plantilla'] = nombre
        nuevo['plantilla'] = {'archivo': ARCH_PLANTILLA, 'nombre': nombre}; nuevo['plantilla_descartes'] = []
        nuevo, resumen = self.aplicar_plantilla(bid, nuevo, usuario)
        self.guardar_estado(bid, nuevo)
        return nuevo, resumen

    def aplicar_plantilla(self, bid, est, usuario):
        """Lee la plantilla guardada en el borrador y reclasifica como C.I. las filas validas (mismo mecanismo y validaciones que el respaldo manual).
        Las C.I. que vinieron de una aplicacion anterior de la plantilla se rehacen desde cero; las que el cajero deshizo (`plantilla_descartes`) no se reaplican.
        Las filas con error NO se aplican y quedan en est['plantilla']['problemas']. Devuelve (est_nuevo, resumen); no guarda el estado."""
        nuevo = json.loads(json.dumps(est)); pl = nuevo.get('plantilla')
        ruta = os.path.join(self.ruta_borrador(bid), ARCH_PLANTILLA)
        recl = nuevo.setdefault('reclasificaciones', {})
        for nf in nuevo.get('plantilla_origen') or []: recl.pop(nf, None)
        nuevo['plantilla_origen'] = []; nuevo.setdefault('plantilla_descartes', [])
        if not pl or not os.path.isfile(ruta): return nuevo, None
        filas = plantilla.leer(ruta)
        res = self.conciliar(bid, nuevo)
        idx = {}
        for _, r in res['facturas'].iterrows(): idx.setdefault(str(r['Número Factura']), r)
        fc = res['totales']['fecha_cierre']; prev = self.vigente(fc, usuario) if fc else None
        excluir = prev['correlativo'] if prev else None
        cuenta_nf, cuenta_bc = {}, {}
        for r in filas:
            if r['factura']: cuenta_nf.setdefault(r['factura'], []).append(r['fila'])
            cn = catalogo.normalizar_codigo(r['codigo'])
            if r['banco'] in plantilla.BANCOS_PLANTILLA and cn: cuenta_bc.setdefault((r['banco'], cn), []).append(r['fila'])
        def otras(lista, fila): return ', '.join(str(x) for x in lista if x != fila)
        problemas, aplicadas, descartadas = [], [], 0
        for r in filas:
            nf, banco, cn = r['factura'], r['banco'], catalogo.normalizar_codigo(r['codigo'])
            def mal(motivo): problemas.append({'fila': r['fila'], 'factura': nf, 'banco': banco, 'codigo': cn, 'motivo': motivo})
            faltan = []
            if not nf: faltan.append('falta el N° de factura')
            if not banco: faltan.append('falta el banco')
            elif banco not in plantilla.BANCOS_PLANTILLA: faltan.append(f"el banco «{banco}» no es válido (use {', '.join(plantilla.BANCOS_PLANTILLA)})")
            if not cn: faltan.append('falta el código de asignación')
            if faltan: mal(faltan[0][0].upper() + '; '.join(faltan)[1:] + '.'); continue
            if nf in (nuevo.get('plantilla_descartes') or []): descartadas += 1; continue
            ok, msg, cn = catalogo.validar_codigo(banco, cn)
            if not ok: mal(msg); continue
            if len(cuenta_nf[nf]) > 1: mal(f'La factura {nf} está repetida en la plantilla (también en la fila {otras(cuenta_nf[nf], r["fila"])}).'); continue
            if len(cuenta_bc[(banco, cn)]) > 1: mal(f'El código {cn} de {banco} está repetido en la plantilla (también en la fila {otras(cuenta_bc[(banco, cn)], r["fila"])}).'); continue
            f = idx.get(nf)
            if f is None: mal(f'La factura {nf} no existe en el movimiento del cajero.'); continue
            e = str(f['estado_conciliacion'])
            if e.startswith('EXCLUIDA'): mal(f'La factura {nf} está ANULADA en el movimiento; no puede ser comunicación interna.'); continue
            if e == 'TARJETA': mal(f'La factura {nf} es de TARJETA; no puede ser comunicación interna.'); continue
            if e in ('CONCILIADA', 'CONFIRMADA POR CAJERO'):
                mal(f"Esta factura ya tiene un pago QR del banco (transacción {f['ID TRANSACCION']}); no puede ser comunicación interna."
                    + (' Si no corresponde, use «Deshacer» en la pantalla de revisión.' if e == 'CONFIRMADA POR CAJERO' else '')); continue
            if e == 'COMUNICACIÓN INTERNA': mal(f'La factura {nf} ya fue reclasificada a mano como comunicación interna; deshaga ese respaldo si prefiere usar la plantilla.'); continue
            if e not in OBSERVADAS: mal(f"La factura {nf} tiene un estado que no permite reclasificarla ({e})."); continue
            otra = next((o for o, d in recl.items() if d['banco'] == banco and catalogo.normalizar_codigo(d['codigo']) == cn), None)
            if otra: mal(f'El código {cn} de {banco} ya está registrado en la factura {otra} de este cierre.'); continue
            c = self.conflictos_ci_banco(banco, cn, excluir)
            if c: mal(f"El código {cn} de {banco} ya está registrado en el cierre {c['correlativo']} (cajero {c['usuario']}).".replace('Ya', 'ya')); continue
            nuevo['confirmaciones'] = [x for x in nuevo['confirmaciones'] if x != nf]; nuevo['elecciones'].pop(nf, None)
            nuevo['respaldos_manuales'] = [x for x in nuevo.get('respaldos_manuales', []) if x != nf]
            recl[nf] = {'banco': banco, 'codigo': cn}; nuevo['plantilla_origen'].append(nf); aplicadas.append(nf)
        resumen = {'leidas': len(filas), 'aplicadas': len(aplicadas), 'con_problemas': len(problemas), 'descartadas': descartadas, 'problemas': problemas}
        pl.update(resumen); nuevo['plantilla'] = pl
        return nuevo, resumen

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
            if not codigo: raise ErrorUsuario('Escriba el número de transacción del reporte QR del banco.')
            if estado == 'POR CONFIRMAR':
                raise ErrorUsuario(f'La factura {nf} ya tiene un movimiento del banco sugerido: use «Sí, corresponde», o reclasifíquela con otro banco.')
            nuevo['elecciones'][nf] = codigo
            if nf not in nuevo['respaldos_manuales']: nuevo['respaldos_manuales'].append(nf)
            r2 = self.conciliar(bid, nuevo); f2 = r2['facturas'].set_index('Número Factura')
            if f2.loc[nf, 'estado_conciliacion'] != 'CONFIRMADA POR CAJERO':
                al = [a['detalle'] for a in r2['alertas'].to_dict('records') if a['referencia'] == nf and 'Elección no aplicada' in a['detalle']]
                motivo = al[0].split('): ', 1)[-1].rstrip('.') if al else 'número de transacción no válido'
                raise ErrorUsuario(f'El número de transacción {codigo} no es válido para la factura {nf}: {motivo}.')
            c = [x for x in self.conflictos_para(r2, usuario) if x['factura'] == nf]
            if c: raise ErrorUsuario(f"El número de transacción {codigo} ya está en el cierre {c[0]['correlativo']} (cajero {c[0]['usuario']}).")
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

    # ---------------- descarte de coincidencias sugeridas ----------------
    def descartar_coincidencia(self, bid, est, nf):
        """El cajero dice que el pago sugerido (POR CONFIRMAR) no corresponde a la factura. Devuelve (est_nuevo, info).
        Solo aplica a POR CONFIRMAR: una CONCILIADA automática o una ya CONFIRMADA (primero «Deshacer») no se descarta."""
        nf = (nf or '').strip()
        res = self.conciliar(bid, est); f = res['facturas'].set_index('Número Factura')
        if nf not in f.index: raise ErrorUsuario('No se encontró la factura.')
        e = f.loc[nf, 'estado_conciliacion']
        if e == 'CONFIRMADA POR CAJERO':
            raise ErrorUsuario(f'La factura {nf} ya fue confirmada: primero use «Deshacer» y luego «No corresponde».')
        if e != 'POR CONFIRMAR':
            raise ErrorUsuario(f'La factura {nf} no tiene una coincidencia dudosa para descartar (está {e}).')
        idt = str(f.loc[nf, 'ID TRANSACCION']); titular = str(f.loc[nf, 'TITULAR ORIGEN'])
        nuevo = json.loads(json.dumps(est)); nuevo.setdefault('rechazos', {})
        lista = nuevo['rechazos'].setdefault(nf, [])
        if idt not in lista: lista.append(idt)
        nuevo['confirmaciones'] = [x for x in nuevo['confirmaciones'] if x != nf]
        return nuevo, {'id_transaccion': idt, 'titular': titular}

    def deshacer_descarte(self, bid, est, nf):
        """Quita los descartes de la factura (vuelve la sugerencia original). Solo si sigue en Observadas. Devuelve (est_nuevo, ids)."""
        nf = (nf or '').strip()
        ids = list((est.get('rechazos') or {}).get(nf) or [])
        if not ids: raise ErrorUsuario(f'La factura {nf} no tiene pagos descartados.')
        res = self.conciliar(bid, est); f = res['facturas'].set_index('Número Factura')
        if nf in f.index and f.loc[nf, 'estado_conciliacion'] not in OBSERVADAS:
            raise ErrorUsuario(f'La factura {nf} ya tiene respaldo ({f.loc[nf, "estado_conciliacion"]}). Primero use «Deshacer» y luego «Deshacer descarte».')
        nuevo = json.loads(json.dumps(est)); nuevo['rechazos'].pop(nf, None)
        return nuevo, ids

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

    def cierres_de(self, usuario=None, estado=None, desde=None, hasta=None, con_alertas=False):
        q = ('SELECT c.*, (SELECT COUNT(*) FROM alertas_cierre a WHERE a.cierre_id=c.id) AS n_alertas, '
             '(SELECT COUNT(*) FROM alertas_cierre a WHERE a.cierre_id=c.id AND a.revisada_en IS NULL) AS n_alertas_pend '
             'FROM cierres c WHERE 1=1'); a = []
        if con_alertas: q += ' AND EXISTS (SELECT 1 FROM alertas_cierre a WHERE a.cierre_id=c.id)'

        if usuario: q += ' AND c.usuario=?'; a.append(usuario)
        if estado: q += ' AND c.estado_revision=?'; a.append(estado)
        if desde: q += ' AND c.fecha_cierre>=?'; a.append(desde)
        if hasta: q += ' AND c.fecha_cierre<=?'; a.append(hasta)
        con = self.con()
        try: return con.execute(q + ' ORDER BY c.id DESC', a).fetchall()
        finally: con.close()

    def alertas_de(self, cid):
        con = self.con()
        try: return con.execute('SELECT * FROM alertas_cierre WHERE cierre_id=? ORDER BY id', (cid,)).fetchall()
        finally: con.close()

    def revisar_alerta(self, cid, aid, usuario):
        """Ingresos/superadmin marca una alerta del cierre como Revisada (usuario y momento; queda en auditoria)."""
        con = self.con(); momento = db.ahora()
        try:
            con.execute('BEGIN IMMEDIATE')
            a = con.execute('SELECT a.*, c.correlativo FROM alertas_cierre a JOIN cierres c ON c.id=a.cierre_id WHERE a.id=? AND a.cierre_id=?', (aid, cid)).fetchone()
            if a is None: raise ErrorUsuario('No se encontró la alerta.')
            if a['revisada_en']: raise ErrorUsuario(f"La alerta de la factura {a['factura']} ya fue revisada por {a['revisada_por']}.")
            con.execute('UPDATE alertas_cierre SET revisada_por=?, revisada_en=? WHERE id=?', (usuario, momento, aid))
            con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                        (momento, usuario, 'alerta_revisada', json.dumps({'correlativo': a['correlativo'], 'cierre': cid, 'alerta': aid, 'tipo': a['tipo'],
                                                                         'factura': a['factura'], 'banco': a['banco'], 'codigo': a['codigo']}, ensure_ascii=False)))
            con.execute('COMMIT')
        except BaseException:
            try: con.execute('ROLLBACK')
            except Exception: pass
            raise
        finally: con.close()
        return a

    def cierre(self, cid):
        con = self.con()
        try: return con.execute('SELECT * FROM cierres WHERE id=?', (cid,)).fetchone()
        finally: con.close()

    def movimientos_cierre(self, cid):
        """Facturas QR BANECO con ID TRANSACCION del cierre (conciliadas automaticas, confirmadas por el cajero o con respaldo manual QR).
        Las comunicaciones internas y las tarjetas no estan aqui. Lista de dicts en el orden en que se registraron."""
        con = self.con()
        try:
            return [dict(r) for r in con.execute('SELECT id_transaccion, factura, monto FROM cierre_movimientos WHERE cierre_id=? ORDER BY rowid', (cid,))]
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
                raise ErrorUsuario(f"La factura {c['factura']} usa el número de transacción {c['id_transaccion']}, que ya está en el cierre "
                                   f"{c['correlativo']} (cajero {c['usuario']}). No se puede presentar.")
            confl_ci = self.conflictos_ci_para(res, usuario)
            if confl_ci:
                c = confl_ci[0]
                raise ErrorUsuario(f"La comunicación interna de la factura {c['factura']} ({c['banco']} {c['codigo']}) ya está registrada en el cierre "
                                   f"{c['correlativo']} (cajero {c['usuario']}). No se puede presentar.")
            alertas = alertas_formato(res, est)
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
            h_pl = None
            if est.get('plantilla') and os.path.isfile(os.path.join(d, ARCH_PLANTILLA)):
                shutil.copyfile(os.path.join(d, ARCH_PLANTILLA), os.path.join(destino, ARCH_PLANTILLA)); h_pl = sha256(os.path.join(destino, ARCH_PLANTILLA))
                if h_pl != sha256(os.path.join(d, ARCH_PLANTILLA)): raise ErrorUsuario('Error de copia de archivos.')
            if h_mov != sha256(os.path.join(d, est['archivo_movimiento'])): raise ErrorUsuario('Error de copia de archivos.')
            motor.exportar(res, os.path.join(destino, 'conciliacion.xlsx'))
            info = {'correlativo': corr, 'fecha_cierre': fc, 'caja': est.get('caja', ''), 'usuario': usuario, 'nombre': nombre,
                    'presentado_en': presentado_en, 'version': version, 'sha256_movimiento': h_mov, 'sha256_banco': h_ban,
                    'estado_revision': ESTADO_INICIAL, 'reemplaza_a': prev['correlativo'] if prev else None,
                    'motivo_reemplazo': (motivo_reemplazo or '').strip() or None, 'alertas_formato': alertas,
                    'notas_descartes': notas_descartes(res)}
            if h_pl: info.update(sha256_plantilla=h_pl, plantilla_origen=list(est.get('plantilla_origen') or []))
            pdf.generar(res, info, os.path.join(destino, 'cierre.pdf'))
            cj = {**info, 'archivo_movimiento': mov, 'estado_revision': ESTADO_INICIAL, 'reemplaza_a': prev['correlativo'] if prev else None,
                  'motivo_reemplazo': (motivo_reemplazo or '').strip() or None, 'estado_qr': t['Estado QR'],
                  'conceptos': conceptos(t), 'totales': t, 'confirmaciones': est['confirmaciones'], 'elecciones': est['elecciones'],
                  'rechazos': est.get('rechazos') or {},
                  'reclasificaciones': est.get('reclasificaciones') or {}, 'respaldos_manuales': est.get('respaldos_manuales') or [],
                  'comunicaciones_internas': res.get('ci', []), 'alertas_formato': alertas,
                  'tarjeta_declarada_por_cajero': est['tarjeta'], 'declaracion': 'El cajero declara conforme el presente cierre.',
                  'facturas': registros(res['facturas']),
                  'sha256_archivos': {n: sha256(os.path.join(destino, n)) for n in ('conciliacion.xlsx', 'cierre.pdf')}}
            if h_pl:
                cj.update(plantilla={k: v for k, v in est['plantilla'].items()}, plantilla_descartes=list(est.get('plantilla_descartes') or []))
                cj['sha256_archivos'][ARCH_PLANTILLA] = h_pl
            with open(os.path.join(destino, 'cierre.json'), 'w', encoding='utf-8') as f:
                json.dump(cj, f, ensure_ascii=False, indent=1, default=_json)
            for n in os.listdir(destino): os.chmod(os.path.join(destino, n), 0o444)
            qr_conc = t['QR conciliado automático (CONCILIADA)'] + t['QR confirmado por cajero']
            cur = con.execute('INSERT INTO cierres (correlativo, fecha_cierre, usuario, nombre, caja, qr_registrado, qr_conciliado, tarjeta_declarada, '
                              'tarjeta_sistema, total_cierre, estado_revision, version, reemplaza_a, motivo_reemplazo, presentado_en, ruta, '
                              'sha256_movimiento, sha256_banco, ci_total, sha256_plantilla) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                              (corr, fc, usuario, nombre, est.get('caja', ''), t['QR registrado'], qr_conc, t['Tarjeta declarada'],
                               t['Tarjeta según sistema'], t['Total cierre'], ESTADO_INICIAL, version, prev['id'] if prev else None,
                               cj['motivo_reemplazo'], presentado_en, rel, h_mov, h_ban, t['C.I. según sistema'], h_pl))
            cid = cur.lastrowid
            f = res['facturas']
            for _, r in f[f['estado_conciliacion'].isin(['CONCILIADA', 'CONFIRMADA POR CAJERO'])].iterrows():
                con.execute('INSERT INTO cierre_movimientos VALUES (?,?,?,?)', (cid, str(r['ID TRANSACCION']), str(r['Número Factura']), float(r['Monto'])))
            for r in res.get('ci', []):
                con.execute('INSERT INTO cierre_ci VALUES (?,?,?,?,?)', (cid, r['factura'], r['banco'], catalogo.normalizar_codigo(r['codigo']), r['monto']))
            for a in alertas:
                con.execute('INSERT INTO alertas_cierre (cierre_id, tipo, factura, banco, codigo, detalle, creado_en) VALUES (?,?,?,?,?,?,?)',
                            (cid, a['tipo'], a['factura'], a['banco'], a['codigo'], a['detalle'], presentado_en))
            if prev: con.execute("UPDATE cierres SET estado_revision='REEMPLAZADO' WHERE id=?", (prev['id'],))
            con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                        (presentado_en, usuario, 'presentar_cierre', json.dumps({'correlativo': corr, 'version': version, 'borrador': bid,
                         'reemplaza': prev['correlativo'] if prev else None, 'motivo': cj['motivo_reemplazo'],
                         'alertas_formato': len(alertas)}, ensure_ascii=False)))
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
            con.execute('INSERT INTO revisiones_historial (cierre_id, momento, usuario, accion, estado_anterior, estado_nuevo, observacion) VALUES (?,?,?,?,?,?,?)',
                        (cid, momento, revisor, 'aprobar' if decision == ESTADO_CONFORME else 'rechazar', c['estado_revision'], decision, obs or None))
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

    # ---------------- revertir revision (superadmin) ----------------
    def revertir(self, cid, usuario, motivo):
        """Superadmin: un cierre CONFORME o RECHAZADO vuelve a PENDIENTE DE REVISIÓN. Motivo obligatorio; el historial se conserva."""
        mot = ' '.join((motivo or '').split())
        if len(mot) < MIN_OBSERVACION:
            raise ErrorUsuario(f'Para revertir la revisión escriba un motivo de al menos {MIN_OBSERVACION} caracteres.')
        con = self.con(); momento = db.ahora()
        try:
            con.execute('BEGIN IMMEDIATE')
            c = con.execute('SELECT * FROM cierres WHERE id=?', (cid,)).fetchone()
            if c is None: raise ErrorUsuario('No se encontró el cierre.')
            if c['estado_revision'] == ESTADO_REEMPLAZADO:
                raise ErrorUsuario(f"El cierre {c['correlativo']} está REEMPLAZADO: no se puede revertir su revisión.")
            if c['estado_revision'] not in (ESTADO_CONFORME, ESTADO_RECHAZADO):
                raise ErrorUsuario(f"El cierre {c['correlativo']} está {c['estado_revision']}: solo se puede revertir un cierre CONFORME o RECHAZADO.")
            con.execute('UPDATE cierres SET estado_revision=?, revisado_por=NULL, revisado_en=NULL, observacion_revision=NULL WHERE id=?', (ESTADO_INICIAL, cid))
            con.execute('INSERT INTO revisiones_historial (cierre_id, momento, usuario, accion, estado_anterior, estado_nuevo, observacion) VALUES (?,?,?,?,?,?,?)',
                        (cid, momento, usuario, 'revertir', c['estado_revision'], ESTADO_INICIAL, mot))
            con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                        (momento, usuario, 'revision_revertida',
                         json.dumps({'correlativo': c['correlativo'], 'cierre': cid, 'estado_anterior': c['estado_revision'], 'motivo': mot,
                                     'revisado_por_anterior': c['revisado_por'], 'revisado_en_anterior': c['revisado_en'],
                                     'observacion_anterior': c['observacion_revision']}, ensure_ascii=False)))
            con.execute('COMMIT')
        except BaseException:
            try: con.execute('ROLLBACK')
            except Exception: pass
            raise
        finally: con.close()
        try: self.regenerar_pdf(cid)
        except Exception as e:
            self.auditar(usuario, 'pdf_revision_error', {'cierre': cid, 'error': str(e)[:200]})
        return self.cierre(cid)

    def historial(self, cid):
        con = self.con()
        try: return con.execute('SELECT * FROM revisiones_historial WHERE cierre_id=? ORDER BY id', (cid,)).fetchall()
        finally: con.close()

    # ---------------- usuarios ----------------
    def sembrar_usuarios(self, entradas):
        """Siembra desde APP_USERS solo los usuarios que no existen (no pisa cambios hechos en la UI), salvo entradas con forzar=true.
        Devuelve (creados, forzados)."""
        creados, forzados = [], []
        con = self.con(); t = db.ahora()
        try:
            for u, d in (entradas or {}).items():
                ex = con.execute('SELECT 1 FROM usuarios WHERE usuario=?', (u,)).fetchone()
                if ex is None:
                    con.execute('INSERT OR IGNORE INTO usuarios (usuario, nombre, rol, clave_hash, activo, creado_en, actualizado_en) VALUES (?,?,?,?,1,?,?)',
                                (u, d['nombre'], d['rol'], d['clave_hash'], t, t)); creados.append(u)
                elif d.get('forzar'):
                    con.execute('UPDATE usuarios SET nombre=?, rol=?, clave_hash=?, activo=1, actualizado_en=? WHERE usuario=?',
                                (d['nombre'], d['rol'], d['clave_hash'], t, u)); forzados.append(u)
        finally: con.close()
        for u in creados: self.auditar('sistema', 'usuario_sembrado', {'usuario': u, 'rol': entradas[u]['rol']})
        for u in forzados: self.auditar('sistema', 'usuario_forzado_desde_env', {'usuario': u, 'rol': entradas[u]['rol']})
        return creados, forzados

    def usuario_get(self, usuario):
        con = self.con()
        try: return con.execute('SELECT * FROM usuarios WHERE usuario=?', ((usuario or '').strip().lower(),)).fetchone()
        finally: con.close()

    def usuarios(self):
        con = self.con()
        try: return con.execute('SELECT * FROM usuarios ORDER BY activo DESC, usuario').fetchall()
        finally: con.close()

    def hay_usuarios_activos(self):
        con = self.con()
        try: return con.execute('SELECT 1 FROM usuarios WHERE activo=1 LIMIT 1').fetchone() is not None
        finally: con.close()

    def usuario_crear(self, actor, usuario, nombre, rol, clave, clave2):
        u = (usuario or '').strip().lower(); n = ' '.join((nombre or '').split())
        if not re.fullmatch(r'[a-z0-9._-]{3,32}', u): raise ErrorUsuario('El usuario debe tener de 3 a 32 caracteres: letras, números, punto, guion o guion bajo.')
        if not n or len(n) > 80: raise ErrorUsuario('Escriba el nombre completo (hasta 80 caracteres).')
        if rol not in auth.ROLES: raise ErrorUsuario('Rol no válido.')
        err = auth.validar_clave_nueva(clave or '', clave2 or '')
        if err: raise ErrorUsuario(err)
        t = db.ahora(); con = self.con()
        try:
            try:
                con.execute('INSERT INTO usuarios (usuario, nombre, rol, clave_hash, activo, creado_en, actualizado_en) VALUES (?,?,?,?,1,?,?)',
                            (u, n, rol, auth.hash_clave(clave), t, t))
            except sqlite3.IntegrityError: raise ErrorUsuario(f'El usuario «{u}» ya existe.')
        finally: con.close()
        self.auditar(actor, 'usuario_creado', {'usuario': u, 'nombre': n, 'rol': rol})

    def _cambiar_usuario(self, actor, usuario, accion, aplicar, detalle):
        """Aplica un cambio sobre un usuario garantizando que siempre quede al menos un superadmin activo."""
        con = self.con(); t = db.ahora()
        try:
            con.execute('BEGIN IMMEDIATE')
            r = con.execute('SELECT * FROM usuarios WHERE usuario=?', (usuario,)).fetchone()
            if r is None: raise ErrorUsuario('No existe ese usuario.')
            cambios = aplicar(r)
            sets = ', '.join(f'{k}=?' for k in cambios) + ', actualizado_en=?'
            con.execute(f'UPDATE usuarios SET {sets} WHERE usuario=?', [*cambios.values(), t, usuario])
            if not con.execute("SELECT 1 FROM usuarios WHERE rol='superadmin' AND activo=1").fetchone():
                raise ErrorUsuario('Debe quedar al menos un superadministrador activo; no se puede dejar el sistema sin uno.')
            con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                        (t, actor, accion, json.dumps({'usuario': usuario, **detalle(r, cambios)}, ensure_ascii=False)))
            con.execute('COMMIT')
        except BaseException:
            try: con.execute('ROLLBACK')
            except Exception: pass
            raise
        finally: con.close()

    def usuario_cambiar_rol(self, actor, usuario, rol):
        if rol not in auth.ROLES: raise ErrorUsuario('Rol no válido.')
        self._cambiar_usuario(actor, usuario, 'usuario_rol_cambiado', lambda r: {'rol': rol}, lambda r, c: {'rol_anterior': r['rol'], 'rol_nuevo': rol})

    def usuario_set_activo(self, actor, usuario, activo):
        a = 1 if activo else 0
        self._cambiar_usuario(actor, usuario, 'usuario_activado' if a else 'usuario_desactivado', lambda r: {'activo': a}, lambda r, c: {})

    def usuario_reset_clave(self, actor, usuario, clave, clave2):
        err = auth.validar_clave_nueva(clave or '', clave2 or '')
        if err: raise ErrorUsuario(err)
        self._cambiar_usuario(actor, usuario, 'usuario_clave_restablecida', lambda r: {'clave_hash': auth.hash_clave(clave)}, lambda r, c: {})

    def cambiar_clave_propia(self, usuario, actual, nueva, nueva2):
        r = self.usuario_get(usuario)
        if r is None or not auth.verificar_clave(actual or '', r['clave_hash']): raise ErrorUsuario('La clave actual no es correcta.')
        err = auth.validar_clave_nueva(nueva or '', nueva2 or '')
        if err: raise ErrorUsuario(err)
        if nueva == actual: raise ErrorUsuario('La clave nueva debe ser distinta de la actual.')
        self._cambiar_usuario(usuario, usuario, 'clave_cambiada', lambda r: {'clave_hash': auth.hash_clave(nueva)}, lambda r, c: {})

    # ---------------- auditoria (lectura) ----------------
    def auditoria(self, usuario=None, accion=None, desde=None, hasta=None, pagina=1, por_pagina=50):
        w, a = [], []
        if usuario: w.append("LOWER(COALESCE(usuario,'')) LIKE ? ESCAPE '\\'"); a.append('%' + usuario.lower().replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_') + '%')
        if accion: w.append('accion=?'); a.append(accion)
        if desde: w.append('substr(momento,1,10)>=?'); a.append(desde)
        if hasta: w.append('substr(momento,1,10)<=?'); a.append(hasta)
        where = (' WHERE ' + ' AND '.join(w)) if w else ''
        con = self.con()
        try:
            total = con.execute('SELECT COUNT(*) FROM auditoria' + where, a).fetchone()[0]
            paginas = max(1, -(-total // por_pagina)); pagina = min(max(1, pagina), paginas)
            filas = con.execute('SELECT * FROM auditoria' + where + ' ORDER BY id DESC LIMIT ? OFFSET ?', [*a, por_pagina, (pagina - 1) * por_pagina]).fetchall()
            acciones = [r[0] for r in con.execute('SELECT DISTINCT accion FROM auditoria ORDER BY accion')]
        finally: con.close()
        return {'filas': filas, 'total': total, 'pagina': pagina, 'paginas': paginas, 'acciones': acciones}

    def regenerar_pdf(self, cid):
        """PDF con el estado de revision (cierre_revision.pdf). El cierre.pdf original, inmutable, no se toca."""
        from . import pdf
        c = self.cierre(cid); d = os.path.join(self.dir, c['ruta'])
        with open(os.path.join(d, 'cierre.json'), encoding='utf-8') as f: cj = json.load(f)
        res = motor.conciliar(os.path.join(d, cj['archivo_movimiento']), os.path.join(d, 'banco_qr_original.xlsx'), cj['tarjeta_declarada_por_cajero'],
                              cj['confirmaciones'], cj['elecciones'], reclasificaciones=cj.get('reclasificaciones') or None,
                              respaldos_manuales=cj.get('respaldos_manuales') or None, rechazos=cj.get('rechazos') or None)
        info = {k: cj.get(k) for k in ('correlativo', 'fecha_cierre', 'caja', 'usuario', 'nombre', 'presentado_en', 'version',
                                       'sha256_movimiento', 'sha256_banco', 'motivo_reemplazo')}
        if cj.get('sha256_plantilla'): info['sha256_plantilla'] = cj['sha256_plantilla']; info['plantilla_origen'] = cj.get('plantilla_origen') or []
        info['alertas_formato'] = cj.get('alertas_formato') or []; info['notas_descartes'] = cj.get('notas_descartes') or []
        u_rev = self.usuario_get(c['revisado_por']) if c['revisado_por'] else None   # nombre completo solo para mostrar; revisado_por sigue siendo el login
        info['revisado_por_nombre'] = u_rev['nombre'] if u_rev is not None and u_rev['nombre'] else c['revisado_por']
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


def tsv_extracto(correlativo, fecha_cierre, movimientos, accion=''):
    """Filas TSV SIN encabezado para pegar en la hoja del lote (Office Script validar_y_marcar):
    Accion | Importe | Codigo de asignacion (ID TRANSACCION) | N° factura | Correlativo | Fecha cierre dd/mm/aaaa.
    Devuelve (texto, n_filas, total). `accion` vacio = marcar; 'LIBERAR' = liberar."""
    fc = fecha(fecha_cierre)
    filas = [f"{accion}\t{float(m['monto']):.2f}\t{m['id_transaccion']}\t{m['factura']}\t{correlativo}\t{fc}" for m in movimientos]
    return '\n'.join(filas), len(filas), round(sum(float(m['monto']) for m in movimientos), 2)


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


TIPO_FORMATO = 'FORMATO'


def alertas_formato(res, est):
    """Alertas NO bloqueantes de formato del codigo de asignacion (regla por banco, ver catalogo).
    Revisa las C.I. (codigo de otro banco) y los respaldos manuales QR BANECO (ID de transaccion escrito por el cajero).
    `est` es el estado del borrador (o el cierre.json): usa 'respaldos_manuales'. Lista de dicts {tipo, factura, banco, codigo, detalle}."""
    out = []
    for r in res.get('ci', []):
        cn = catalogo.normalizar_codigo(r['codigo']); regla = catalogo.formato_incumplido(r['banco'], cn)
        if regla: out.append({'tipo': TIPO_FORMATO, 'factura': str(r['factura']), 'banco': r['banco'], 'codigo': cn, 'detalle': regla})
    manuales = set(est.get('respaldos_manuales') or [])
    if manuales:
        f = res['facturas']
        for _, r in f[f['Número Factura'].astype(str).isin(manuales) & (f['estado_conciliacion'] == 'CONFIRMADA POR CAJERO')].iterrows():
            cn = catalogo.normalizar_codigo(str(r['ID TRANSACCION'])); regla = catalogo.formato_incumplido(catalogo.BANCO_QR, cn)
            if regla: out.append({'tipo': TIPO_FORMATO, 'factura': str(r['Número Factura']), 'banco': catalogo.BANCO_QR, 'codigo': cn, 'detalle': regla})
    return out


def candidatos(res, nf):
    f = res['facturas'].set_index('Número Factura')
    if nf not in f.index: return []
    m = f.loc[nf, 'Monto']; b = res['banco']
    desc = {d['id'] for d in res.get('descartes', []) if d['factura'] == nf}
    c = b[(abs(b['MONTO'] - m) < 1e-9) & (b['estado_banco'].str.startswith('AJENO') | b['estado_banco'].str.startswith('AMBIGUO'))]
    return [{'id': r['ID TRANSACCION'], 'titular': r['TITULAR ORIGEN'], 'hora': r['FECHA / HORA'], 'descartado': r['ID TRANSACCION'] in desc}
            for _, r in c.iterrows()]


def notas_descartes(res):
    """Una nota por cada pago descartado por el cajero, segun como quedo la factura. Lista de dicts {factura, id, titular, texto}."""
    f = res['facturas'].set_index('Número Factura'); ci = {r['factura']: r for r in res.get('ci', [])}; out = []
    for d in res.get('descartes', []):
        nf = d['factura']; base = f"El cajero descartó la coincidencia sugerida ({d['id']}, titular «{d['titular']}») de la factura {nf}"
        if nf in ci: fin = f" y la reclasificó como comunicación interna ({ci[nf]['banco']} {ci[nf]['codigo']})."
        elif nf in f.index and f.loc[nf, 'estado_conciliacion'] == 'CONFIRMADA POR CAJERO':
            r = f.loc[nf]
            fin = (f" y asignó el pago {r['ID TRANSACCION']} (titular «{r['TITULAR ORIGEN']}»)"
                   + (' como respaldo manual.' if r['nivel'] == 'RESPALDO MANUAL' else ('.' if r['nivel'] == 'ELECCIÓN CAJERO' else ' tras revisar otra coincidencia.')))
        else: fin = '; la factura no tiene respaldo asignado.'
        out.append({'factura': nf, 'id': d['id'], 'titular': d['titular'], 'texto': base + fin})
    return out
