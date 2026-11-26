"""Lector del reporte 'Ingreso Diario' del sistema de facturacion exportado en PDF.

leer_pdf_a1(path) -> (filas, columnas, meta, alertas)
  filas: lista de dicts {columna canonica: texto, '_fila_excel': n} (misma forma que el lector de Excel; motor.parse_a1 los normaliza)
  meta: {'caja', 'usuario', 'fecha_emision', 'sede'}; alertas: [(nivel, ref, mensaje)] (p. ej. el detalle no suma el total del pie).
Estrategia: tabla por lineas (bordes) con pdfplumber; si la pagina no tiene lineas, por posicion de palabras bajo la cabecera.
Las celdas de varias lineas se unen con espacio (la fecha y la hora en dos lineas quedan 'dd/mm/aaaa hh:mm')."""
import re
from . import motor

MSG_SIN_TEXTO = 'El PDF no tiene texto legible; exporte el reporte desde el sistema de facturación (Excel o PDF).'


def _celda(x):
    return re.sub(r'\s+', ' ', str(x or '')).strip()


def _num(s):
    try: return motor.num(s)
    except ValueError: return None


def _mapa_cabecera(celdas):
    """{columna canonica: indice} si la fila es la cabecera del detalle (tiene 'Nro' y 'Monto'), si no None."""
    cols, _ = motor._mapear_a1(celdas)
    return cols if 'Nro' in cols and 'Monto' in cols and len(cols) >= 3 else None


def _filas_por_tablas(page, estado):
    """Detalle de una pagina usando las tablas con bordes. estado = {'cols': mapa, 'n': ncolumnas} se arrastra entre paginas."""
    filas = []
    for tabla in page.extract_tables({'vertical_strategy': 'lines', 'horizontal_strategy': 'lines'}) or []:
        for row in tabla:
            celdas = [_celda(x) for x in row]
            cab = _mapa_cabecera(celdas)
            if cab: estado['cols'] = cab; estado['n'] = len(celdas); continue
            if not estado.get('cols') or len(celdas) != estado['n']: continue
            if not re.fullmatch(r'\d+', celdas[estado['cols']['Nro']] or 'x'): continue     # totales / pie / filas vacias
            filas.append({k: celdas[i] for k, i in estado['cols'].items()})
    return filas


def _filas_por_posicion(page, estado):
    """Sin lineas de tabla: columnas por la posicion x de la cabecera; una fila nueva empieza cuando la columna Nro trae un entero."""
    words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    if not words: return []
    # lineas por coordenada vertical
    lineas = []
    for w in sorted(words, key=lambda w: (round(w['top']), w['x0'])):
        if lineas and abs(lineas[-1][0] - w['top']) <= 2.5: lineas[-1][1].append(w)
        else: lineas.append([w['top'], [w]])
    hdr = None
    for i, (_, ws) in enumerate(lineas):
        grupos = []                                   # palabras contiguas (hueco pequeno) forman una celda de cabecera
        for w in sorted(ws, key=lambda w: w['x0']):
            if grupos and w['x0'] - grupos[-1]['x1'] < 6: grupos[-1] = {'text': grupos[-1]['text'] + ' ' + w['text'], 'x0': grupos[-1]['x0'], 'x1': w['x1']}
            else: grupos.append({'text': w['text'], 'x0': w['x0'], 'x1': w['x1']})
        cab = _mapa_cabecera([g['text'] for g in grupos])
        if cab: hdr = (i, grupos, cab); break
    if hdr is None and estado.get('pos'): grupos, cab = estado['pos']; ini = 0
    elif hdr is None: return []
    else:
        ini = hdr[0] + 1; grupos, cab = hdr[1], hdr[2]; estado['pos'] = (grupos, cab)
    inv = {i: k for k, i in cab.items()}
    limites = [(g['x0'] - 3) for g in grupos]
    def columna(w):
        c = None
        for i, x in enumerate(limites):
            if w['x0'] >= x or w['x1'] - 1 > x + 1 and i == 0: c = i
        return c
    filas = []; cur = None
    for _, ws in lineas[ini:]:
        porcol = {}
        for w in sorted(ws, key=lambda w: w['x0']):
            c = columna(w)
            if c is not None: porcol.setdefault(c, []).append(w['text'])
        nro_i = cab['Nro']
        nro = ' '.join(porcol.get(nro_i, []))
        if re.fullmatch(r'\d+', nro):
            cur = {k: '' for k in cab}; filas.append(cur)
        elif cur is None or re.match(r'(?i)^(monto|total|p[aá]gina)', ' '.join(t for v in porcol.values() for t in v)):
            if filas: cur = None
            continue
        if cur is None: continue
        for c, toks in porcol.items():
            if c in inv: cur[inv[c]] = (cur[inv[c]] + ' ' + ' '.join(toks)).strip()
    return filas


def _meta(texto):
    def buscar(pat, flags=re.I):
        m = re.search(pat, texto, flags); return m.group(1).strip() if m else ''
    return {
        'caja': buscar(r'\bCaja\s*:?\s*([A-Za-z0-9_\-]+)'),
        'usuario': buscar(r'\bUsuario\s*:?\s*([^\s]+)'),
        'fecha_emision': buscar(r'\bFecha\s*:?\s*(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?)'),
        'sede': buscar(r'\bSEDE\s+([^\n]+)', 0),
    }


_PIE = {'credito': r'cr[eé]dito', 'deposito': r'dep[oó]sito', 'efectivo': r'efectivo', 'tarjeta': r'tarjeta'}


def _totales_pie(texto):
    t = {}
    for k, pat in _PIE.items():
        m = re.search(r'(?i)monto\s+' + pat + r'\s*:?\s*(-?[\d.,]+)', texto)
        if m and _num(m.group(1)) is not None: t[k] = _num(m.group(1))
    ms = re.findall(r'(?im)^\s*(?:monto\s+)?total\s*:?\s*(-?[\d][\d.,]*)\s*$', texto)
    if ms and _num(ms[-1]) is not None: t['total'] = _num(ms[-1])
    return t


def _validar_totales(filas, tot):
    """Alertas (no bloqueo) si el detalle de Validos no coincide con el pie del reporte."""
    alertas = []
    validos = [f for f in filas if motor.canon_estado(f.get('Estado', '')) in motor.CONFIG['estados_validos'] and _num(f.get('Monto')) is not None]
    suma = round(sum(_num(f['Monto']) for f in validos), 2)
    def por_tipo(tp): return round(sum(_num(f['Monto']) for f in validos if motor.canon_tipo_pago(f.get('Tipo Pago', '')) == tp), 2)
    def avisa(txt): alertas.append(('ALERTA', '', 'Lectura del PDF: ' + txt + ' Revise que el PDF esté completo (todas las páginas) o use el Excel del sistema.'))
    if 'total' in tot:
        if abs(suma - tot['total']) > 0.005: avisa(f"la suma de las facturas válidas del detalle (Bs {suma:,.2f}) no coincide con el Total del pie del reporte (Bs {tot['total']:,.2f}).")
    elif tot:
        comp = round(sum(tot.get(k, 0) for k in _PIE), 2)
        if abs(suma - comp) > 0.005: avisa(f"la suma de las facturas válidas del detalle (Bs {suma:,.2f}) no coincide con la suma de los montos del pie (Bs {comp:,.2f}).")
    for k, tp in (('efectivo', 'EFECTIVO'), ('tarjeta', 'TARJETA'), ('deposito', 'BANCO')):
        if k in tot and abs(por_tipo(tp) - tot[k]) > 0.005:
            avisa(f"el detalle de facturas válidas con tipo de pago {tp} suma Bs {por_tipo(tp):,.2f} y el pie del reporte indica Monto {k.capitalize()} Bs {tot[k]:,.2f}.")
    if not tot:
        avisa('no se encontraron los totales del pie del reporte, no se pudo verificar que el detalle esté completo.')
    return alertas


def leer_pdf_a1(path):
    import pdfplumber
    try:
        pdf = pdfplumber.open(path)
    except Exception:
        raise motor.ErrorLecturaA1('No se pudo abrir el PDF (archivo dañado o protegido con contraseña).')
    with pdf:
        textos = []; filas = []; estado = {}
        for page in pdf.pages:
            textos.append(page.extract_text() or '')
            f = _filas_por_tablas(page, estado)
            filas += f or _filas_por_posicion(page, estado)
        texto = '\n'.join(textos)
    if not texto.strip(): raise motor.ErrorLecturaA1(MSG_SIN_TEXTO)
    cols = list((estado.get('cols') or (estado.get('pos') or (None, {}))[1]).keys())
    if not cols:
        raise motor.ErrorLecturaA1(f"No se encontró la tabla del reporte en el PDF (no hay una fila con 'Nro' y 'Monto'). Páginas: {len(textos)}.")
    for n, f in enumerate(filas, start=1): f['_fila_excel'] = n
    return filas, cols, _meta(texto), _validar_totales(filas, _totales_pie(texto))
