#!/usr/bin/env python3
"""Motor de conciliacion v1 (perspectiva cajero) - especificacion ejecutable.

Uso CLI:
  python3 -m app.motor A1.xls A2.xlsx --tarjeta 1191 [--salida RUTA.xlsx]
API:
  conciliar(ruta_a1, ruta_a2, tarjeta_declarada) -> dict(facturas, banco, resumen, alertas, tarjetas, totales)
  conciliar_df(a1_df, a2_df, tarjeta_declarada)  -> idem, con DataFrames ya parseados (testeable)
Los parsers parse_a1 / parse_a2 / num provienen de etapa1_analisis.py (copiados aqui porque ese
script ejecuta todo al importarse). parse_a1 acepta .xls (xlrd) y .xlsx (openpyxl), detectado por contenido."""
import sys, re, os, argparse, unicodedata
import xlrd, openpyxl, pandas as pd

OUT = os.environ.get('MOTOR_SALIDAS', './salidas')   # solo para la CLI

CONFIG = {
    'estados_validos': ['Válido'],
    'estados_excluidos': ['Anulado'],          # conocidos: se excluyen con aviso informativo (no es "no reconocido")
    'mapa_tipo_pago': {'EFECTIVO': 'QR', 'BANCO': 'QR', 'TARJETA': 'TARJETA'},
    'min_len_token_parcial': 4,                # R5 nombre_parcial
    'umbral_nombre_fuerte': 0.75,              # R5
    'len_titular_truncado': 19,                # R5: titular >= este largo => ultimo token puede ser prefijo
    'redondeo_score': 6, 'redondeo_dif': 4,    # tolerancia de comparacion para empates (R8)
}
NIVEL_RANK = {'FUERTE': 0, 'MEDIA': 1, 'DEBIL': 2}

# ---------------- parsers (de etapa1_analisis.py) ----------------
def num(s):
    return float(str(s).replace('BOB', '').replace(',', '').strip())

def _es_xlsx(path):
    with open(path, 'rb') as fh: return fh.read(4) == b'PK\x03\x04'

def leer_filas_a1(path):
    """Devuelve (filas, datemode). filas = lista de listas (celdas vacias = ''). datemode None si es xlsx."""
    if _es_xlsx(path):
        with open(path, 'rb') as fh: wb = openpyxl.load_workbook(fh, data_only=True)   # objeto archivo: no depende de la extension
        ws = wb['RIngresoDiario'] if 'RIngresoDiario' in wb.sheetnames else wb.active
        return [['' if v is None else v for v in r] for r in ws.iter_rows(values_only=True)], None
    wb = xlrd.open_workbook(path); sh = wb.sheet_by_name('RIngresoDiario')
    return [sh.row_values(r) for r in range(sh.nrows)], wb.datemode

def parse_a1(path):
    rows, datemode = leer_filas_a1(path)
    hdr = next(r for r, row in enumerate(rows) if any(isinstance(x, str) and x.strip() == 'Nro' for x in row))
    cols = {}
    for c, x in enumerate(rows[hdr]):
        if x != '': cols[re.sub(r'\s+', ' ', str(x)).strip()] = c
    det = []; r = hdr + 1
    while r < len(rows) and any(x != '' for x in rows[r]):
        d = {k: (rows[r][c] if c < len(rows[r]) else '') for k, c in cols.items()}; d['_fila_excel'] = r + 1; det.append(d); r += 1
    df = pd.DataFrame(det)
    df['Nro'] = df['Nro'].astype(int)
    df['Fecha'] = df['Fecha'].apply(lambda x: x if hasattr(x, 'year') else xlrd.xldate_as_datetime(x, datemode))
    df['Número Factura'] = df['Número Factura'].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and float(x).is_integer() else str(x))
    df['Monto'] = df['Monto'].astype(float)
    return df

def metadatos_a1(path):
    """Metadatos de la cabecera del reporte (p. ej. 'Caja:'). Devuelve dict con 'caja' ('' si no existe)."""
    rows, _ = leer_filas_a1(path)
    caja = ''
    for row in rows[:hdr_limite(rows)]:
        for c, x in enumerate(row):
            if isinstance(x, str) and x.strip().rstrip(':').strip().lower() == 'caja':
                caja = next((str(v).strip() for v in row[c + 1:] if v != ''), ''); break
        if caja: break
    return {'caja': caja}

def hdr_limite(rows):
    return next((r for r, row in enumerate(rows) if any(isinstance(x, str) and x.strip() == 'Nro' for x in row)), min(len(rows), 20))

def parse_a2(path):
    ws = openpyxl.load_workbook(path).active
    it = list(ws.iter_rows(values_only=True)); head = list(it[0]); rows = []
    for r in it[1:]:
        if all(x is None for x in r): continue
        d = dict(zip(head, r))
        if str(d.get('PUNTOS DE VENTA', '')).strip().upper() == 'TOTAL' or d['ID TRANSACCION'] is None:
            continue                                   # R2: excluir fila TOTAL
        rows.append(d)
    return pd.DataFrame(rows)

# ---------------- preparacion (R1, R2) ----------------
def preparar_a1(a1, cfg=CONFIG):
    """R1. Devuelve (facturas_df con columnas 'tipo' y 'estado_prep', alertas list)."""
    f = a1.copy(); alertas = []; tipos = []; est = []
    for _, r in f.iterrows():
        e = str(r['Estado']).strip(); tp = str(r['Tipo Pago']).strip().upper(); ref = str(r['Número Factura'])
        if e in cfg['estados_excluidos']:
            tipos.append(None); est.append(f'EXCLUIDA ({e})')
            alertas.append(('INFO', ref, f"Factura {ref} Bs {r['Monto']:.2f} excluida por estado {e}"))
        elif e not in cfg['estados_validos']:
            tipos.append(None); est.append('NO PROCESADA - valor no reconocido')
            alertas.append(('ALERTA', ref, f"valor no reconocido: Estado='{r['Estado']}' (factura {ref}, no se procesa)"))
        elif tp not in cfg['mapa_tipo_pago']:
            tipos.append(None); est.append('NO PROCESADA - valor no reconocido')
            alertas.append(('ALERTA', ref, f"valor no reconocido: Tipo Pago='{r['Tipo Pago']}' (factura {ref}, no se procesa)"))
        else:
            tipos.append(cfg['mapa_tipo_pago'][tp]); est.append('VALIDA')
    f['tipo'] = tipos; f['estado_prep'] = est
    return f, alertas

def preparar_a2(a2):
    """R2. Monto numerico, fecha datetime, duplicados de ID TRANSACCION. Devuelve DataFrame con 'dup'."""
    b = a2.copy().reset_index(drop=True)
    b['ID TRANSACCION'] = b['ID TRANSACCION'].astype(str)
    b['MONTO'] = b['MONTO'].apply(lambda x: x if isinstance(x, (int, float)) else num(x)).astype(float)
    if not pd.api.types.is_datetime64_any_dtype(b['FECHA / HORA']):
        b['FECHA / HORA'] = pd.to_datetime(b['FECHA / HORA'], format='%d/%m/%Y %H:%M:%S')
    b['dup'] = b['ID TRANSACCION'].duplicated(keep='first')
    return b

# ---------------- nombre (R5) ----------------
def norm(s):
    s = unicodedata.normalize('NFD', str(s).upper().replace('Ñ', 'N'))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.findall(r'[A-Z0-9]+', s)

def _score_vs(titular, target, cfg):
    tt = norm(titular); tg = norm(target)
    if not tt: return 0.0, False
    trunc = len(str(titular).strip()) >= cfg['len_titular_truncado']
    hit = 0; parcial = False
    for i, t in enumerate(tt):
        ok = t in tg or (trunc and i == len(tt) - 1 and any(g.startswith(t) for g in tg))
        if ok:
            hit += 1
            if len(t) >= cfg['min_len_token_parcial']: parcial = True
    return hit / len(tt), parcial

def clasif_nombre(titular, razon, estudiante, cfg=CONFIG):
    """Devuelve (score, 'fuerte'|'parcial'|'nulo'). Score = mayor entre Razon Social y Nombre Estudiante."""
    s1, p1 = _score_vs(titular, razon, cfg); s2, p2 = _score_vs(titular, estudiante, cfg)
    s = max(s1, s2)
    if s >= cfg['umbral_nombre_fuerte']: return s, 'fuerte'
    if p1 or p2: return s, 'parcial'
    return s, 'nulo'

# ---------------- motor ----------------
def conciliar(ruta_a1, ruta_a2, tarjeta_declarada, confirmaciones=None, elecciones=None, conformidad=None, motivo=None, cfg=CONFIG):
    return conciliar_df(parse_a1(ruta_a1), parse_a2(ruta_a2), tarjeta_declarada, confirmaciones, elecciones, conformidad, motivo, cfg)

def conciliar_df(a1, a2, tarjeta_declarada, confirmaciones=None, elecciones=None, conformidad=None, motivo=None, cfg=CONFIG):
    """confirmaciones: lista de Número Factura (POR CONFIRMAR). elecciones: {Número Factura: ID TRANSACCION}.
    conformidad: None | 'CONFORME' | 'NO CONFORME' (+ motivo). Lanza ValueError si no es valida."""
    alertas = []
    fa, al = preparar_a1(a1, cfg); alertas += al
    bk = preparar_a2(a2)
    qr = fa[(fa['estado_prep'] == 'VALIDA') & (fa['tipo'] == 'QR')]
    tj = fa[(fa['estado_prep'] == 'VALIDA') & (fa['tipo'] == 'TARJETA')]

    # R3 fecha del cierre = fecha(s) de las facturas validas
    fechas = sorted({d.date() for d in fa[fa['estado_prep'] == 'VALIDA']['Fecha']})
    if len(fechas) > 1: alertas.append(('ALERTA', '', f'Facturas con mas de una fecha: {fechas}; la ventana usa todas'))
    if not fechas: alertas.append(('ALERTA', '', 'Sin facturas validas: no hay fecha de cierre'))

    # estado de cada movimiento
    bk['estado_banco'] = ''; bk['factura_asignada'] = ''; bk['nivel_asignado'] = ''
    usable = []
    for i, m in bk.iterrows():
        if m['dup']:
            bk.at[i, 'estado_banco'] = 'DUPLICADO EN ARCHIVO BANCO'
            alertas.append(('ALERTA', m['ID TRANSACCION'], f"ID TRANSACCION {m['ID TRANSACCION']} repetido; solo la primera ocurrencia es utilizable"))
        elif m['FECHA / HORA'].date() not in fechas:
            bk.at[i, 'estado_banco'] = 'FUERA DE FECHA'
            alertas.append(('INFO', m['ID TRANSACCION'], f"Movimiento {m['ID TRANSACCION']} ({m['TITULAR ORIGEN']}, Bs {m['MONTO']:.2f}) fecha {m['FECHA / HORA']:%d/%m/%Y} fuera de la fecha del cierre"))
        else: usable.append(i)

    # pares con monto exacto (R4) y nivel (R7)
    pares = []
    for fi, f in qr.iterrows():
        hf = f['Fecha'].replace(microsecond=0)
        for bi in usable:
            m = bk.loc[bi]
            if abs(f['Monto'] - m['MONTO']) > 1e-9: continue
            sc, cl = clasif_nombre(m['TITULAR ORIGEN'], f['Razon Social'], f['Nombre Estudiante'], cfg)
            dif = (hf - m['FECHA / HORA']).total_seconds() / 60.0     # R6: + = factura despues del pago
            if cl == 'fuerte': nivel = 'FUERTE'
            elif cl == 'parcial': nivel = 'MEDIA'     # v1.1: nombre nulo + hora cercana ya NO sube a MEDIA
            else: nivel = 'DEBIL'
            pares.append(dict(fi=fi, bi=bi, nivel=nivel, score=sc, dif=dif, nombre=cl,
                              key=(NIVEL_RANK[nivel], -round(sc, cfg['redondeo_score']), round(abs(dif), cfg['redondeo_dif']))))
    # R8 asignacion greedy 1 a 1 con deteccion de empates
    fuertes = sorted([p for p in pares if p['nivel'] != 'DEBIL'], key=lambda p: p['key'])
    asign = {}                 # fi -> par
    f_amb = {}; m_amb = {}     # fi -> [bi], bi -> [fi]
    usadas_f = set(); usados_m = set(); bloq_f = set(); bloq_m = set()
    j = 0
    while j < len(fuertes):
        k = j
        while k < len(fuertes) and fuertes[k]['key'] == fuertes[j]['key']: k += 1
        grupo = [p for p in fuertes[j:k] if p['fi'] not in usadas_f | bloq_f and p['bi'] not in usados_m | bloq_m]
        for p in grupo:
            conflicto = [q for q in grupo if q is not p and (q['fi'] == p['fi'] or q['bi'] == p['bi'])]
            if conflicto:
                f_amb.setdefault(p['fi'], []).append(p['bi']); m_amb.setdefault(p['bi'], []).append(p['fi'])
            else:
                asign[p['fi']] = p
        for p in grupo:
            if p['fi'] in asign and asign[p['fi']] is p:
                usadas_f.add(p['fi']); usados_m.add(p['bi'])
            else:
                bloq_f.add(p['fi']); bloq_m.add(p['bi'])
        j = k

    # estados por factura (R9)
    fa['estado_conciliacion'] = fa['estado_prep']; fa['nivel'] = ''; fa['ID TRANSACCION'] = ''; fa['TITULAR ORIGEN'] = ''
    fa['hora_banco'] = pd.NaT; fa['score_nombre'] = None; fa['dif_min'] = None; fa['candidatos'] = ''
    for fi in tj.index: fa.at[fi, 'estado_conciliacion'] = 'TARJETA'
    def poner(fi, bi, estado, nivel, score=None, dif=None):
        m = bk.loc[bi]
        fa.at[fi, 'estado_conciliacion'] = estado; fa.at[fi, 'nivel'] = nivel; fa.at[fi, 'ID TRANSACCION'] = m['ID TRANSACCION']
        fa.at[fi, 'TITULAR ORIGEN'] = m['TITULAR ORIGEN']; fa.at[fi, 'hora_banco'] = m['FECHA / HORA']
        fa.at[fi, 'score_nombre'] = None if score is None else round(score, 3); fa.at[fi, 'dif_min'] = None if dif is None else round(dif, 2)
        fa.at[fi, 'candidatos'] = ''
        bk.at[bi, 'estado_banco'] = 'ASIGNADO ' + estado; bk.at[bi, 'factura_asignada'] = fa.at[fi, 'Número Factura']; bk.at[bi, 'nivel_asignado'] = nivel
    for fi in qr.index:
        if fi in asign:
            p = asign[fi]
            poner(fi, p['bi'], 'CONCILIADA' if p['nivel'] == 'FUERTE' else 'POR CONFIRMAR', p['nivel'], p['score'], p['dif'])
        elif fi in f_amb:
            fa.at[fi, 'estado_conciliacion'] = 'AMBIGUO'; fa.at[fi, 'nivel'] = 'AMBIGUO'
        else:
            fa.at[fi, 'estado_conciliacion'] = 'SIN RESPALDO'

    # B. confirmaciones y elecciones del cajero
    idx_f = {str(r['Número Factura']): i for i, r in fa.iterrows()}
    for nf in (confirmaciones or []):
        nf = str(nf); fi = idx_f.get(nf)
        if fi is None or fa.at[fi, 'estado_conciliacion'] != 'POR CONFIRMAR':
            alertas.append(('ALERTA', nf, f"Confirmación no aplicada: factura {nf} no está en POR CONFIRMAR"))
            continue
        bi = bk.index[(bk['ID TRANSACCION'] == fa.at[fi, 'ID TRANSACCION']) & (~bk['dup'])][0]
        fa.at[fi, 'estado_conciliacion'] = 'CONFIRMADA POR CAJERO'; bk.at[bi, 'estado_banco'] = 'ASIGNADO CONFIRMADA POR CAJERO'
    for nf, idt in (elecciones or {}).items():
        nf = str(nf); idt = str(idt); fi = idx_f.get(nf)
        def rechazo(msg): alertas.append(('ALERTA', nf, f"Elección no aplicada ({nf} -> {idt}): {msg}"))
        if fi is None or fa.at[fi, 'estado_conciliacion'] not in ('SIN RESPALDO', 'AMBIGUO'):
            rechazo('la factura no está SIN RESPALDO ni AMBIGUO'); continue
        cand = bk.index[(bk['ID TRANSACCION'] == idt) & (~bk['dup'])]
        if len(cand) == 0: rechazo('ID TRANSACCION inexistente'); continue
        bi = cand[0]
        if abs(bk.at[bi, 'MONTO'] - fa.at[fi, 'Monto']) > 1e-9: rechazo(f"monto distinto (banco {bk.at[bi, 'MONTO']:.2f} vs factura {fa.at[fi, 'Monto']:.2f})"); continue
        if bk.at[bi, 'FECHA / HORA'].date() not in fechas: rechazo('fecha distinta a la del cierre'); continue
        if bk.at[bi, 'estado_banco'].startswith('ASIGNADO'): rechazo(f"movimiento ya usado por la factura {bk.at[bi, 'factura_asignada']}"); continue
        poner(fi, bi, 'CONFIRMADA POR CAJERO', 'ELECCIÓN CAJERO')
    # facturas AMBIGUO restantes: candidatos que sigan libres
    for fi, bis in f_amb.items():
        if fa.at[fi, 'estado_conciliacion'] != 'AMBIGUO': continue
        libres = [b for b in bis if not bk.at[b, 'estado_banco'].startswith('ASIGNADO')]
        nf = fa.at[fi, 'Número Factura']
        if not libres:
            fa.at[fi, 'estado_conciliacion'] = 'SIN RESPALDO'; fa.at[fi, 'nivel'] = ''
        else:
            ids = [bk.at[b, 'ID TRANSACCION'] for b in libres]
            fa.at[fi, 'candidatos'] = 'AMBIGUO - requiere elección del cajero: ' + ', '.join(ids)
            alertas.append(('ALERTA', nf, f"AMBIGUO - requiere elección del cajero: factura {nf} empata con movimientos {ids}"))
    for fi in qr.index:
        if fa.at[fi, 'estado_conciliacion'] == 'SIN RESPALDO':
            deb = [q for q in pares if q['fi'] == fi and not bk.at[q['bi'], 'estado_banco'].startswith('ASIGNADO')]
            if deb:
                fa.at[fi, 'candidatos'] = 'candidato sugerido (DEBIL, no asignado): ' + '; '.join(
                    f"{bk.at[q['bi'], 'ID TRANSACCION']} {bk.at[q['bi'], 'TITULAR ORIGEN']} dif {q['dif']:.1f} min" for q in deb)
    # movimientos no asignados (R10)
    amb_vivas = {fi for fi in f_amb if fa.at[fi, 'estado_conciliacion'] == 'AMBIGUO'}
    for bi in usable:
        if bk.at[bi, 'estado_banco'] == '':
            fs = [x for x in m_amb.get(bi, []) if x in amb_vivas]
            if fs:
                bk.at[bi, 'estado_banco'] = 'AMBIGUO - requiere elección del cajero'
                bk.at[bi, 'factura_asignada'] = ', '.join(fa.at[x, 'Número Factura'] for x in fs)
            else:
                bk.at[bi, 'estado_banco'] = 'AJENO / NO CORRESPONDE AL CAJERO'

    # R11 / D totales y estado
    s = lambda est: float(fa[fa['estado_conciliacion'] == est]['Monto'].sum())
    n = lambda est: int((fa['estado_conciliacion'] == est).sum())
    reg = float(qr['Monto'].sum()); auto = s('CONCILIADA'); conf = s('CONFIRMADA POR CAJERO')
    pend = s('POR CONFIRMAR') + s('AMBIGUO'); sinr = s('SIN RESPALDO')
    dif_qr = reg - (auto + conf)
    if abs(dif_qr) < 0.005: estado_qr = 'CUADRA'
    elif n('SIN RESPALDO') == 0: estado_qr = 'PENDIENTE DE CONFIRMACIÓN'
    else: estado_qr = 'NO CUADRA'
    # E. conformidad del cajero
    if conformidad is not None:
        if conformidad not in ('CONFORME', 'NO CONFORME'): raise ValueError("conformidad debe ser 'CONFORME' o 'NO CONFORME'")
        if conformidad == 'CONFORME' and estado_qr != 'CUADRA':
            raise ValueError(f"No se puede registrar CONFORME: Estado QR = {estado_qr} (diferencia QR {dif_qr:.2f})")
        if conformidad == 'NO CONFORME' and not (motivo or '').strip():
            raise ValueError("NO CONFORME exige un motivo no vacío")
    tsis = float(tj['Monto'].sum())
    tot = {'fecha_cierre': fechas[0] if fechas else None, 'n_facturas_qr': len(qr), 'QR registrado': reg,
           'QR conciliado automático (CONCILIADA)': auto, 'QR confirmado por cajero': conf,
           'QR pendiente de confirmar (POR CONFIRMAR + AMBIGUO)': pend, 'QR sin respaldo': sinr,
           'Diferencia QR': dif_qr, 'Estado QR': estado_qr, 'Tarjeta según sistema': tsis,
           'Tarjeta declarada': float(tarjeta_declarada), 'Diferencia tarjeta (declarada - sistema)': float(tarjeta_declarada) - tsis,
           'Total cierre': reg + float(tarjeta_declarada),
           'n_CONCILIADA': n('CONCILIADA'), 'n_CONFIRMADA POR CAJERO': n('CONFIRMADA POR CAJERO'), 'n_POR CONFIRMAR': n('POR CONFIRMAR'),
           'n_SIN RESPALDO': n('SIN RESPALDO'), 'n_AMBIGUO': n('AMBIGUO'),
           'n_movimientos_ajenos': int((bk['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO').sum()),
           'n_movimientos_ambiguos': int(bk['estado_banco'].str.startswith('AMBIGUO').sum()),
           'n_duplicados': int(bk['dup'].sum()), 'n_fuera_de_fecha': int((bk['estado_banco'] == 'FUERA DE FECHA').sum()),
           'Conformidad del cajero': conformidad or 'SIN CONFORMIDAD', 'Motivo conformidad': (motivo or '') if conformidad else '',
           'Generado el': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
    resumen = pd.DataFrame([{'concepto': k, 'valor': (str(v) if k == 'fecha_cierre' else v)} for k, v in tot.items()])
    cols_f = ['Nro', 'Número Factura', 'Fecha', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago', 'Estado', 'tipo', 'Monto',
              'estado_conciliacion', 'nivel', 'ID TRANSACCION', 'TITULAR ORIGEN', 'hora_banco', 'score_nombre', 'dif_min', 'candidatos']
    facturas = fa[cols_f].rename(columns={'tipo': 'Tipo (motor)', 'Fecha': 'hora_factura'})
    banco = bk.drop(columns=['dup'])
    alert_df = pd.DataFrame(alertas, columns=['nivel', 'referencia', 'detalle'])
    return {'facturas': facturas, 'banco': banco, 'resumen': resumen, 'alertas': alert_df,
            'tarjetas': facturas[facturas['estado_conciliacion'] == 'TARJETA'].copy(), 'totales': tot}

def exportar(res, ruta):
    f = res['facturas']; b = res['banco']
    with pd.ExcelWriter(ruta) as w:
        res['resumen'].to_excel(w, sheet_name='Resumen', index=False)
        f[f['Tipo (motor)'] == 'QR'].to_excel(w, sheet_name='Facturas_QR', index=False)
        b[~b['estado_banco'].str.startswith('ASIGNADO')].to_excel(w, sheet_name='Movimientos_ajenos', index=False)
        res['tarjetas'].to_excel(w, sheet_name='Tarjetas', index=False)
        res['alertas'].to_excel(w, sheet_name='Alertas', index=False)
    return ruta

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('a1', help='Movimiento del cajero (.xls/.xlsx)')
    ap.add_argument('a2', help='Reporte QR del banco (.xlsx)')
    ap.add_argument('--tarjeta', type=float, required=True, help='total de tarjetas declarado por el cajero')
    ap.add_argument('--salida', default=None)
    ap.add_argument('--confirmar', action='append', default=[], help='Número Factura POR CONFIRMAR que el cajero acepta (repetible)')
    ap.add_argument('--elegir', action='append', default=[], help='FACTURA=ID_TRANSACCION (repetible)')
    ap.add_argument('--conformidad', choices=['CONFORME', 'NO CONFORME'], default=None)
    ap.add_argument('--motivo', default=None)
    a = ap.parse_args()
    try:
        el = dict(x.split('=', 1) for x in a.elegir)
        res = conciliar(a.a1, a.a2, a.tarjeta, a.confirmar, el, a.conformidad, a.motivo)
    except ValueError as e:
        sys.exit(f'ERROR: {e}')
    os.makedirs(OUT, exist_ok=True)
    ruta = a.salida or f"{OUT}/cierre_{res['totales']['fecha_cierre']}_v1.xlsx"
    exportar(res, ruta)
    pd.set_option('display.width', 250); pd.set_option('display.max_columns', 30); pd.set_option('display.max_colwidth', 60)
    print(res['resumen'].to_string(index=False)); print()
    print(res['facturas'][['Número Factura', 'Tipo Pago', 'Monto', 'estado_conciliacion', 'nivel', 'ID TRANSACCION', 'score_nombre', 'dif_min']].to_string(index=False))
    print(); print(res['alertas'].to_string(index=False)); print('\nExportado:', ruta)

if __name__ == '__main__':
    main()
