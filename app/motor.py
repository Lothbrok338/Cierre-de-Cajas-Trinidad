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
    """Monto flexible: acepta 'BOB'/'Bs', espacios, miles y decimales con coma o punto. Si hay ambos separadores, el ultimo es el decimal;
    si hay uno solo: repetido => miles; una vez con exactamente 3 digitos tras el separador (y 1-3 antes, sin cero inicial) => miles; si no, decimal."""
    if isinstance(s, (int, float)): return float(s)
    t = re.sub(r'(?i)bob|bs\.?', '', str(s).replace('\xa0', ' ')); t = re.sub(r'\s+', '', t)
    if not re.fullmatch(r'-?[\d.,]*\d', t): raise ValueError(f'monto no reconocido: {s!r}')
    if ',' in t and '.' in t:
        dec = ',' if t.rfind(',') > t.rfind('.') else '.'; mil = '.' if dec == ',' else ','
        t = t.replace(mil, '').replace(dec, '.')
    elif ',' in t or '.' in t:
        sep = ',' if ',' in t else '.'; a, *resto = t.split(sep)
        if len(resto) > 1 or (len(resto[0]) == 3 and a.lstrip('-') not in ('', '0') and len(a.lstrip('-')) <= 3): t = t.replace(sep, '')
        else: t = t.replace(sep, '.')
    return float(t)

def _estilo_monto(x):
    """Descripcion del formato de monto si no es el estandar del banco ('1,220.76 BOB'); None si es estandar o no es texto."""
    if not isinstance(x, str): return None
    try: std = float(re.sub(r'(?i)bob', '', x).replace(',', '').strip())
    except ValueError: std = None
    if std is None or abs(std - num(x)) > 1e-9: return 'los montos vienen con coma decimal o punto de miles (se interpretó el último separador como decimal)'
    if re.search(r'(?i)\bbs', x): return "los montos vienen con «Bs»"
    return None

def _es_xlsx(path):
    with open(path, 'rb') as fh: return fh.read(4) == b'PK\x03\x04'

def _es_pdf(path):
    with open(path, 'rb') as fh: return b'%PDF' in fh.read(1024)

class ErrorLecturaA1(ValueError):
    """El movimiento del cajero no se pudo leer; el mensaje dice que falta (solo nombres de hojas/columnas, nunca datos personales)."""

def _nk(s):
    """Clave de comparacion: sin tildes, mayusculas ni espacios/saltos/signos ('Número\nFactura' == 'numero factura')."""
    s = unicodedata.normalize('NFD', str(s).lower())
    return re.sub(r'[^a-z0-9]', '', ''.join(c for c in s if unicodedata.category(c) != 'Mn'))

# columnas del reporte: clave normalizada -> nombre canonico
COLS_A1_CANON = {_nk(c): c for c in ['Nro', 'Fecha', 'Número Factura', 'Nit/C.I.', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago',
                                      'Banco / Cuenta bancaria', 'Código de Bancarización', 'Nro. Cuenta SAP asociada', 'Monto', 'Cuenta', 'Estado']}
COLS_A1_OBLIG = ['Nro', 'Fecha', 'Número Factura', 'Tipo Pago', 'Monto', 'Estado']
COLS_A1_OPC = ['Nit/C.I.', 'Razon Social', 'Nombre Estudiante', 'Banco / Cuenta bancaria', 'Código de Bancarización', 'Nro. Cuenta SAP asociada', 'Cuenta']

_SIN_A1 = {
    'Nro': ['nro', 'n', 'no', 'item', 'correlativo'],
    'Fecha': ['fecha', 'fechafactura', 'fechaemision', 'fechadeemision'],
    'Número Factura': ['numerofactura', 'nrofactura', 'nfactura', 'numerodefactura', 'nrodefactura', 'nofactura', 'numfactura'],
    'Nit/C.I.': ['nitci', 'nit', 'ci'],
    'Razon Social': ['razonsocial'],
    'Nombre Estudiante': ['nombreestudiante', 'nombredelestudiante', 'estudiante'],
    'Tipo Pago': ['tipopago', 'tipodepago', 'formapago', 'formadepago'],
    'Banco / Cuenta bancaria': ['bancocuentabancaria'],
    'Código de Bancarización': ['codigodebancarizacion', 'codigobancarizacion'],
    'Nro. Cuenta SAP asociada': ['nrocuentasapasociada'],
    'Monto': ['monto'], 'Cuenta': ['cuenta'], 'Estado': ['estado'], 'Hora': ['hora', 'horafactura'],
}
_CANON_A1 = {k: c for c, ks in _SIN_A1.items() for k in ks}
_REGLAS_A1 = [
    ('Estado', lambda k: 'estado' in k),
    ('Fecha', lambda k: 'fecha' in k),
    ('Hora', lambda k: 'hora' in k),
    ('Monto', lambda k: 'monto' in k),
    ('Código de Bancarización', lambda k: 'bancariz' in k),
    ('Nro. Cuenta SAP asociada', lambda k: 'sap' in k),
    ('Banco / Cuenta bancaria', lambda k: 'banco' in k or 'cuentabancaria' in k),
    ('Tipo Pago', lambda k: 'pago' in k and ('tipo' in k or 'forma' in k)),
    ('Número Factura', lambda k: 'factura' in k and 'nit' not in k),
    ('Nombre Estudiante', lambda k: 'estudiante' in k),
    ('Razon Social', lambda k: 'razon' in k or 'cliente' in k),
    ('Nit/C.I.', lambda k: 'nit' in k),
]
_FALLBACK_A1 = [('Monto', lambda k: k in ('importe', 'total', 'montototal', 'importetotal')), ('Estado', lambda k: 'situacion' in k)]

def _mapear_a1(row):
    return _mapear(row, _CANON_A1, _REGLAS_A1, _FALLBACK_A1)

def _es_cabecera(celdas):
    m, _ = _mapear_a1(celdas)
    return 'Monto' in m and ('Nro' in m or 'Número Factura' in m) and len(m) >= 3

def _sheet_rows_a1(path):
    """Todas las hojas como (nombre, filas, datemode); 'RIngresoDiario' primero si existe."""
    if _es_xlsx(path):
        with open(path, 'rb') as fh: wb = openpyxl.load_workbook(fh, data_only=True)   # objeto archivo: no depende de la extension
        hojas = [(ws.title, [['' if v is None else v for v in r] for r in ws.iter_rows(values_only=True)], None) for ws in wb.worksheets]
    else:
        wb = xlrd.open_workbook(path)
        hojas = [(sh.name, [sh.row_values(r) for r in range(sh.nrows)], wb.datemode) for sh in wb.sheets()]
    return sorted(hojas, key=lambda h: h[0] != 'RIngresoDiario')   # sort estable

def _hoja_a1(path):
    """(filas, datemode, indice_cabecera) de la primera hoja con una fila de cabecera ('Nro' y 'Monto')."""
    hojas = _sheet_rows_a1(path)
    for _, rows, dm in hojas:
        for i, row in enumerate(rows):
            if _es_cabecera(row): return rows, dm, i
    nombres = ', '.join(h[0] for h in hojas) or '(ninguna)'
    raise ErrorLecturaA1(f"No se encontró la tabla del reporte (no hay una fila con 'Nro' y 'Monto'). Hojas encontradas: {nombres}.")

def leer_filas_a1(path):
    """Devuelve (filas, datemode) de la hoja del reporte. filas = lista de listas (celdas vacias = ''). datemode None si es xlsx."""
    rows, dm, _ = _hoja_a1(path)
    return rows, dm

_ESTADOS = {_nk(v): v for v in CONFIG['estados_validos'] + CONFIG['estados_excluidos']}
_ESTADOS.update({'valida': 'Válido', 'anulada': 'Anulado'})

def canon_estado(x):
    """'Valido'/'VÁLIDO' -> 'Válido'; 'ANULADO' -> 'Anulado'; otro valor queda tal cual (se alerta en preparar_a1)."""
    return _ESTADOS.get(_nk(x), str(x).strip())

def canon_tipo_pago(x):
    return re.sub(r'\s+', ' ', str(x)).strip().upper()

def _fecha_texto(x, estilos=None):
    """Texto dd/mm/aaaa, dd-mm-aaaa, aaaa-mm-dd (con hora opcional hh:mm[:ss]) -> datetime."""
    import datetime as dt
    t = str(x).strip()
    m = re.search(r'(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})', t)
    if m: y, mo, d = m.groups(); est = 'fechas en formato aaaa-mm-dd'
    else:
        m = re.search(r'(\d{1,2})([/\-.])(\d{1,2})\2(\d{4})', t)
        if not m: raise ValueError(f'fecha no reconocida: {x!r}')
        d, sep, mo, y = m.groups(); est = None if sep == '/' else 'fechas con guiones o puntos (dd-mm-aaaa)'
    h = re.search(r'(\d{1,2}):(\d{2})(?::(\d{2}))?', t[m.end():])
    if est and estilos is not None: estilos.add(est)
    return dt.datetime(int(y), int(mo), int(d), int(h.group(1)) if h else 0, int(h.group(2)) if h else 0, int(h.group(3) or 0) if h else 0)

def _fecha(x, datemode, estilos=None):
    import datetime as dt
    if isinstance(x, dt.datetime): return x
    if isinstance(x, dt.date): return dt.datetime.combine(x, dt.time(0, 0))
    if isinstance(x, str): return _fecha_texto(x, estilos)
    return xlrd.xldate_as_datetime(x, datemode or 0)

def _hora(x):
    """Hora separada: texto hh:mm[:ss], hora real, datetime o fraccion de dia (serial) -> datetime.time."""
    import datetime as dt
    if _vacio(x): return dt.time(0, 0, 0)
    if isinstance(x, dt.datetime): return x.time().replace(microsecond=0)
    if isinstance(x, dt.time): return x.replace(microsecond=0)
    if isinstance(x, (int, float)):
        seg = int(round((float(x) % 1) * 86400)) % 86400
        return dt.time(seg // 3600, seg % 3600 // 60, seg % 60)
    m = re.search(r'(\d{1,2}):(\d{2})(?::(\d{2}))?', str(x))
    if not m: raise ValueError('hora no reconocida')
    return dt.time(int(m.group(1)), int(m.group(2)), int(m.group(3) or 0))

def _dia_a2(x, estilos=None):
    import datetime as dt
    if isinstance(x, (dt.datetime, dt.date, str)): return _fecha(x, None, estilos).date()
    return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(x))).date()

def _monto(x):
    if isinstance(x, (int, float)): return float(x)
    return num(x)

def _df_a1(det, hdr_cols):
    """det: lista de dicts {columna canonica: valor crudo, '_fila_excel': n}. Mismo resultado para Excel y PDF."""
    faltan = [c for c in COLS_A1_OBLIG if c not in hdr_cols]
    if faltan:
        raise ErrorLecturaA1(f"Faltan las columnas: {', '.join(faltan)}. Columnas encontradas: {', '.join(hdr_cols) or '(ninguna)'}.")
    if not det: raise ErrorLecturaA1('La tabla del reporte no tiene filas de detalle.')
    df = pd.DataFrame(det)
    for c in COLS_A1_OPC:
        if c not in df.columns: df[c] = ''
    return df

def _normalizar_a1(df, datemode):
    raw_montos = list(df['Monto']); est = set()
    try:
        df['Nro'] = df['Nro'].apply(lambda x: int(float(str(x).replace(',', '').strip())))
        df['Fecha'] = df['Fecha'].apply(lambda x: _fecha(x, datemode, est))
        if 'Hora' in df.columns:
            df['Fecha'] = [f if _vacio(h) else f.replace(hour=_hora(h).hour, minute=_hora(h).minute, second=_hora(h).second) for f, h in zip(df['Fecha'], df['Hora'])]
            df = df.drop(columns='Hora')
        df['Número Factura'] = df['Número Factura'].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and float(x).is_integer() else str(x).strip())
        df['Monto'] = df['Monto'].apply(_monto).astype(float)
    except (ValueError, TypeError) as e:
        raise ErrorLecturaA1(f'Hay filas con datos que no se pueden interpretar (Nro, Fecha o Monto inválidos): {str(e)[:80]}')
    df.attrs['_estilos'] = est | {e for e in map(_estilo_monto, raw_montos) if e}
    df['Estado'] = df['Estado'].apply(canon_estado)
    df['Tipo Pago'] = df['Tipo Pago'].apply(canon_tipo_pago)
    for c in COLS_A1_OPC: df[c] = df[c].apply(lambda x: '' if x is None else x)
    return df

def parse_a1(path):
    """Lee el reporte 'Ingreso Diario' (.xls, .xlsx o .pdf, detectado por contenido). df.attrs['alertas_lectura'] = alertas de la lectura (PDF)."""
    if _es_pdf(path):
        from . import lector_pdf
        det, cols, _meta, alertas = lector_pdf.leer_pdf_a1(path)
        df = _normalizar_a1(_df_a1(det, cols), None)
        df.attrs['alertas_lectura'] = list(alertas) + [('INFO', '', f'Reporte de facturación: {e}.') for e in sorted(df.attrs.pop('_estilos', set()))]
        return df
    rows, datemode, hdr = _hoja_a1(path)
    cols, notas = _mapear_a1(rows[hdr])
    sin_nro = 'Nro' not in cols
    det = []; r = hdr + 1
    while r < len(rows) and any(x != '' for x in rows[r]):
        d = {k: (rows[r][c] if c < len(rows[r]) else '') for k, c in cols.items()}; d['_fila_excel'] = r + 1
        if sin_nro: d['Nro'] = len(det) + 1
        det.append(d); r += 1
    df = _normalizar_a1(_df_a1(det, list(cols) + (['Nro'] if sin_nro else [])), datemode)
    estilos = df.attrs.pop('_estilos', set())
    al = _alertas_cabecera(notas, 'Reporte de facturación')
    if sin_nro: al.append(('INFO', '', "Reporte de facturación: no hay columna Nro; se numeraron las filas en orden."))
    df.attrs['alertas_lectura'] = al + [('INFO', '', f'Reporte de facturación: {e}.') for e in sorted(estilos)]
    return df

def metadatos_a1(path):
    """Metadatos de la cabecera del reporte. Devuelve dict con 'caja' ('' si no existe) y, en PDF, usuario/fecha_emision/sede."""
    if _es_pdf(path):
        from . import lector_pdf
        return lector_pdf.leer_pdf_a1(path)[2]
    rows, _, hdr = _hoja_a1(path)
    caja = ''
    for row in rows[:hdr]:
        for c, x in enumerate(row):
            if isinstance(x, str) and x.strip().rstrip(':').strip().lower() == 'caja':
                caja = next((str(v).strip() for v in row[c + 1:] if v != ''), ''); break
        if caja: break
    return {'caja': caja}

class ErrorLecturaA2(ValueError):
    """El reporte QR del banco no se pudo leer; el mensaje dice que falta (solo nombres de hojas/columnas, nunca datos personales)."""

# columna canonica -> sinonimos exactos (claves normalizadas con _nk); lo demas se resuelve por palabras clave y, al final, por contenido
_SIN_A2 = {
    'ID TRANSACCION': ['idtransaccion', 'transaccionpago', 'nrotransaccion', 'transaccion', 'idtransaccionpago'],
    'TITULAR ORIGEN': ['titularorigen', 'titular'],
    'PUNTOS DE VENTA': ['puntosdeventa', 'puntodeventa', 'puntoventa'],
    'MONTO': ['monto', 'importe'],
    'FECHA / HORA': ['fechahora'],
    '_FECHA': ['fechapago'],
    '_HORA': ['horapago'],
    'ID QR': ['idqr'], 'GLOSA': ['glosa'], 'CUENTA ORIGEN': ['cuentaorigen'], 'CUENTA DESTINO': ['cuentadestino'],
    '_ESTADO': ['estado'],
}
_CANON_A2 = {k: c for c, ks in _SIN_A2.items() for k in ks}
_REGLAS_A2 = [   # (columna, predicado sobre la clave normalizada); el orden resuelve los casos con varias palabras clave
    ('_ESTADO', lambda k: 'estado' in k),
    ('FECHA / HORA', lambda k: 'fecha' in k and 'hora' in k),
    ('_FECHA', lambda k: 'fecha' in k),
    ('_HORA', lambda k: 'hora' in k),
    ('MONTO', lambda k: 'monto' in k or 'importe' in k),
    ('PUNTOS DE VENTA', lambda k: 'punto' in k and 'venta' in k),
    ('CUENTA ORIGEN', lambda k: 'cuenta' in k and 'origen' in k),
    ('CUENTA DESTINO', lambda k: 'cuenta' in k and 'destino' in k),
    ('TITULAR ORIGEN', lambda k: 'titular' in k),
    ('GLOSA', lambda k: 'glosa' in k),
    ('ID QR', lambda k: 'qr' in k and 'id' in k),
    ('ID TRANSACCION', lambda k: 'transac' in k),
]
_OPC_A2 = ['ID QR', 'GLOSA', 'CUENTA ORIGEN', 'CUENTA DESTINO', 'PUNTOS DE VENTA']
_ESTADOS_A2_INFO = {'anulado', 'devuelto', 'reversado', 'rechazado', 'cancelado', 'extornado'}
_NOMBRE_COL = {'_FECHA': 'FECHA PAGO', '_HORA': 'HORA PAGO', '_ESTADO': 'ESTADO'}

def _mapear(row, exactas, reglas, fallbacks=()):
    """Cabecera flexible. 1) sinonimos exactos; 2) palabras clave (encabezados de hasta 40 caracteres) sobre las columnas libres;
    3) fallbacks [(columna, predicado)] solo si la columna sigue sin asignar. Devuelve ({columna: indice}, [(encabezado, columna)] de lo no exacto)."""
    m = {}; notas = []
    claves = [(i, _nk(v)) for i, v in enumerate(row) if isinstance(v, str)]
    for i, k in claves:
        c = exactas.get(k)
        if c and c not in m: m[c] = i
    for etapa in (reglas, fallbacks):
        for i, k in claves:
            if i in m.values() or not k or len(row[i]) > 40: continue
            for c, pred in etapa:
                if c not in m and pred(k): m[c] = i; notas.append((str(row[i]).strip(), c)); break
    return m, notas

def _faltan_a2(m):
    f = [c for c in ('ID TRANSACCION', 'TITULAR ORIGEN', 'MONTO') if c not in m]
    if 'FECHA / HORA' not in m:
        if '_FECHA' not in m: f.append('FECHA / HORA (o FECHA PAGO)')
        elif '_HORA' not in m: f.append('HORA PAGO')
    return f

def _vacio(x): return x is None or (isinstance(x, str) and not x.strip())

_RE_MONTO_TXT = re.compile(r'(?i)^\s*(bs\.?\s*)?-?\d[\d.,\s]*\s*(bob|bs\.?)?\s*$')
def _c_monto(v):
    if isinstance(v, float): return not v.is_integer()
    if not isinstance(v, str) or not _RE_MONTO_TXT.match(v): return False
    return bool(re.search(r'(?i)bob|bs', v) or re.search(r'[.,]\d{1,2}\s*$', v))
def _c_fechahora(v):
    import datetime as dt
    if isinstance(v, dt.datetime): return v.time() != dt.time(0, 0)
    return isinstance(v, str) and bool(re.search(r'\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}', v) and re.search(r'\d{1,2}:\d{2}', v))
def _c_fecha(v):
    import datetime as dt
    if isinstance(v, dt.datetime): return v.time() == dt.time(0, 0)
    if isinstance(v, dt.date): return True
    return isinstance(v, str) and bool(re.fullmatch(r'\s*(\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4})\s*', v))
def _c_hora(v):
    import datetime as dt
    return isinstance(v, dt.time) or (isinstance(v, str) and bool(re.fullmatch(r'\s*\d{1,2}:\d{2}(:\d{2})?\s*', v)))
def _c_trans(v):
    if isinstance(v, (int, float)) and float(v).is_integer(): v = str(int(v))
    return isinstance(v, str) and bool(re.fullmatch(r'\d{6,12}', v.strip()))
def _c_titular(v):
    return isinstance(v, str) and ' ' in v.strip() and v == v.upper() and bool(re.fullmatch(r"[A-ZÁÉÍÓÚÑÜ.'\s]+", v.strip()))

def _detectar_por_contenido(rows, m, h, faltan, notas):
    """Ultimo recurso: asigna columnas obligatorias que faltan por el contenido de las columnas libres. Lanza ErrorLecturaA2 si es ambiguo."""
    datos = [r for r in rows[h + 1:] if not all(_vacio(x) for x in r) and not any(isinstance(x, str) and _nk(x) == 'total' for x in r)]
    if not datos: return
    libres = [i for i in range(max(len(r) for r in rows[h:h + 1] + datos)) if i not in m.values()]
    def col(i): return [r[i] for r in datos if i < len(r) and not _vacio(r[i])]
    def todos(i, f):
        v = col(i); return bool(v) and sum(1 for x in v if f(x)) >= 0.9 * len(v)
    def unica(i, f):
        v = col(i); return todos(i, f) and len({str(x) for x in v}) == len(v)
    pedidos = []
    if 'MONTO' in faltan: pedidos.append(('MONTO', lambda i: todos(i, _c_monto)))
    if 'ID TRANSACCION' in faltan: pedidos.append(('ID TRANSACCION', lambda i: unica(i, _c_trans)))
    if 'TITULAR ORIGEN' in faltan: pedidos.append(('TITULAR ORIGEN', lambda i: todos(i, _c_titular)))
    if 'FECHA / HORA' not in m and '_FECHA' not in m:
        pedidos.append(('FECHA / HORA', lambda i: todos(i, _c_fechahora)))
        pedidos.append(('_FECHA', lambda i: todos(i, _c_fecha)))
    if 'FECHA / HORA' not in m and '_HORA' not in m: pedidos.append(('_HORA', lambda i: todos(i, _c_hora)))
    nom = lambda i: f"columna {i + 1}" + (f" ('{str(rows[h][i]).strip()}')" if i < len(rows[h]) and isinstance(rows[h][i], str) and rows[h][i].strip() else '')
    elegidas = {}
    for c, f in pedidos:
        if c in m: continue
        cand = [i for i in libres if f(i)]
        if len(cand) > 1:
            raise ErrorLecturaA2(f"No se pudo decidir cuál es la columna de {_NOMBRE_COL.get(c, c)} por su contenido (candidatas: {', '.join(nom(i) for i in cand)}). "
                                 f"Columnas encontradas: {', '.join(str(x).strip() for x in rows[h] if isinstance(x, str) and x.strip()) or '(ninguna)'}.")
        if cand:
            m[c] = cand[0]; libres.remove(cand[0]); elegidas[c] = cand[0]
            if c == 'FECHA / HORA': m.pop('_FECHA', None); m.pop('_HORA', None)
    for c, i in elegidas.items(): notas.append(('contenido', f"{_NOMBRE_COL.get(c, c)} = {nom(i)}"))

def _fecha_junta(rows, m, h):
    """Columna de fecha sin columna de hora aparte cuyos valores ya traen la hora (fecha real con hora o texto 'dd/mm/aaaa hh:mm') = FECHA / HORA."""
    if '_FECHA' in m and '_HORA' not in m and 'FECHA / HORA' not in m:
        v = [r[m['_FECHA']] for r in rows[h + 1:] if m['_FECHA'] < len(r) and not _vacio(r[m['_FECHA']]) and not (isinstance(r[m['_FECHA']], str) and _nk(r[m['_FECHA']]) == 'total')]
        if v and all(_c_fechahora(x) for x in v): m['FECHA / HORA'] = m.pop('_FECHA')

def _hoja_a2(path):
    """(filas, mapa, indice_cabecera, notas) de la hoja activa o, si no, de la primera que tenga la cabecera (primeras 15 filas)."""
    with open(path, 'rb') as fh: wb = openpyxl.load_workbook(fh, data_only=True)
    hojas = [wb.active] + [w for w in wb.worksheets if w is not wb.active]
    mejor = None; nombres = []
    for ws in hojas:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        for i, row in enumerate(rows[:15]):
            m, notas = _mapear(row, _CANON_A2, _REGLAS_A2); _fecha_junta(rows, m, i)
            if not _faltan_a2(m): return rows, m, i, notas
            if len(m) >= 2 and (mejor is None or len(m) > len(mejor[1])): mejor = (rows, m, i, notas)
        nombres.append(ws.title)
    if not mejor:
        raise ErrorLecturaA2(f"No se encontró la fila de cabecera (debe tener transacción, titular, monto y fecha en las primeras 15 filas). Hojas revisadas: {', '.join(nombres) or '(ninguna)'}.")
    rows, m, h, notas = mejor
    enc = ', '.join(str(x).strip() for x in rows[h] if isinstance(x, str) and x.strip()) or '(ninguna)'
    faltan = _faltan_a2(m)
    _detectar_por_contenido(rows, m, h, faltan, notas)
    if _faltan_a2(m):
        raise ErrorLecturaA2(f"Faltan las columnas: {', '.join(_faltan_a2(m))}. Columnas encontradas: {enc}.")
    return rows, m, h, notas

def _alertas_cabecera(notas, fuente):
    """Alertas INFO (trazabilidad) de las decisiones flexibles de cabecera: [(encabezado, columna)] y detecciones por contenido."""
    al = []; cont = [d for k, d in notas if k == 'contenido']
    for k, c in notas:
        if k != 'contenido': al.append(('INFO', '', f"{fuente}: el encabezado «{k}» se interpretó como {_NOMBRE_COL.get(c, c)}."))
    if cont: al.append(('INFO', '', f"{fuente}: se detectaron las columnas por su contenido: {'; '.join(cont)}."))
    return al

def _idt(x):
    return str(int(x)) if isinstance(x, (int, float)) and float(x).is_integer() else str(x).strip()

def parse_a2(path):
    """Lee el reporte QR del banco (formato original o nuevo; cabecera flexible, orden libre, columnas extra ignoradas). Devuelve DataFrame
    con las columnas canonicas del original. df.attrs['alertas_lectura'] = alertas (decisiones flexibles INFO, estados no pagados, total que no cuadra)."""
    import datetime as dt
    rows, m, h, notas = _hoja_a2(path)
    comb = 'FECHA / HORA' not in m
    alertas = _alertas_cabecera(notas, 'Reporte del banco'); out = []; total = None; suma_todas = 0.0; estilos = set()
    def cel(r, c): return r[m[c]] if c in m and m[c] < len(r) else None
    for r in rows[h + 1:]:
        if all(_vacio(x) for x in r): continue
        idt = cel(r, 'ID TRANSACCION')
        if _vacio(idt):
            if any(isinstance(x, str) and _nk(x) == 'total' for x in r):     # R2: excluir fila TOTAL
                try: total = num(cel(r, 'MONTO'))
                except (ValueError, TypeError): pass
            continue
        d = {c: ('' if _vacio(cel(r, c)) else cel(r, c)) for c in _OPC_A2}
        d['ID TRANSACCION'] = _idt(idt); d['TITULAR ORIGEN'] = cel(r, 'TITULAR ORIGEN'); d['MONTO'] = cel(r, 'MONTO')
        try:
            if comb: d['FECHA / HORA'] = dt.datetime.combine(_dia_a2(cel(r, '_FECHA'), estilos), _hora(cel(r, '_HORA')))
            else: d['FECHA / HORA'] = _fecha(cel(r, 'FECHA / HORA'), None, estilos)
            mt = num(d['MONTO']); e = _estilo_monto(d['MONTO'])
            if e: estilos.add(e)
        except (ValueError, TypeError):
            raise ErrorLecturaA2(f"El movimiento {d['ID TRANSACCION']} tiene fecha, hora o monto que no se pueden interpretar.")
        suma_todas += mt
        est = cel(r, '_ESTADO')
        if '_ESTADO' in m and _nk(est or '') not in ('pagado', 'pagada'):
            k = _nk(est or '')
            alertas.append(('INFO' if k in _ESTADOS_A2_INFO else 'ALERTA', d['ID TRANSACCION'],
                            f"El movimiento {d['ID TRANSACCION']} (Bs {mt:.2f}) del banco tiene estado {str(est).strip() or '(vacío)'} y no se cuenta."))
            continue
        out.append(d)
    alertas += [('INFO', '', f'Reporte del banco: {e}.') for e in sorted(estilos)]
    cols = ['ID QR', 'ID TRANSACCION', 'TITULAR ORIGEN', 'CUENTA ORIGEN', 'GLOSA', 'CUENTA DESTINO', 'FECHA / HORA', 'PUNTOS DE VENTA', 'MONTO']
    df = pd.DataFrame(out, columns=cols)
    if total is not None:
        suma = sum(num(x) for x in df['MONTO'])
        if abs(suma - total) > 0.005 and abs(suma_todas - total) > 0.005:
            alertas.append(('ALERTA', '', f"La suma de los pagos del banco (Bs {suma:.2f}) no coincide con la fila TOTAL del archivo (Bs {total:.2f}). Verifique que el archivo esté completo."))
    df.attrs['alertas_lectura'] = alertas
    return df

# ---------------- preparacion (R1, R2) ----------------
def preparar_a1(a1, cfg=CONFIG):
    """R1. Devuelve (facturas_df con columnas 'tipo' y 'estado_prep', alertas list)."""
    f = a1.copy(); alertas = []; tipos = []; est = []
    for _, r in f.iterrows():
        e = canon_estado(r['Estado']); tp = canon_tipo_pago(r['Tipo Pago']); ref = str(r['Número Factura'])
        if e in cfg['estados_excluidos']:
            tipos.append(None); est.append(f'EXCLUIDA ({e})')
            alertas.append(('INFO', ref, f"La factura {ref} (Bs {r['Monto']:.2f}) está {e} y no se cuenta en el cierre."))
        elif e not in cfg['estados_validos']:
            tipos.append(None); est.append('NO PROCESADA - valor no reconocido')
            alertas.append(('ALERTA', ref, f"La factura {ref} tiene un estado con valor no reconocido ('{r['Estado']}') y no se procesó. Revise el reporte del sistema de facturación."))
        elif tp not in cfg['mapa_tipo_pago']:
            tipos.append(None); est.append('NO PROCESADA - valor no reconocido')
            alertas.append(('ALERTA', ref, f"La factura {ref} tiene un tipo de pago con valor no reconocido ('{r['Tipo Pago']}') y no se procesó. Revise el reporte del sistema de facturación."))
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
def conciliar(ruta_a1, ruta_a2, tarjeta_declarada, confirmaciones=None, elecciones=None, conformidad=None, motivo=None, cfg=CONFIG,
              reclasificaciones=None, respaldos_manuales=None, rechazos=None):
    return conciliar_df(parse_a1(ruta_a1), parse_a2(ruta_a2), tarjeta_declarada, confirmaciones, elecciones, conformidad, motivo, cfg,
                        reclasificaciones, respaldos_manuales, rechazos)

def conciliar_df(a1, a2, tarjeta_declarada, confirmaciones=None, elecciones=None, conformidad=None, motivo=None, cfg=CONFIG,
                 reclasificaciones=None, respaldos_manuales=None, rechazos=None):
    """rechazos: {Número Factura: [ID TRANSACCION descartados por el cajero]}. Un par (factura, movimiento) rechazado se excluye del cruce
    como si no existiera (el movimiento queda libre para otras facturas). Con rechazos vacío el resultado no cambia.
    confirmaciones: lista de Número Factura (POR CONFIRMAR). elecciones: {Número Factura: ID TRANSACCION}.
    reclasificaciones: {Número Factura: {'banco': ..., 'codigo': ...}} -> la factura QR se reclasifica como Comunicación interna
    (C.I.) y sale del cruce QR. respaldos_manuales: lista de facturas cuya elección es un respaldo manual (nivel RESPALDO MANUAL).
    conformidad: None | 'CONFORME' | 'NO CONFORME' (+ motivo). Lanza ValueError si no es valida."""
    alertas = []
    fa, al = preparar_a1(a1, cfg); alertas += al
    alertas += list(a1.attrs.get('alertas_lectura') or [])
    alertas += list(a2.attrs.get('alertas_lectura') or [])
    if not any(str(v).strip() for col in ('Razon Social', 'Nombre Estudiante') for v in fa[col]):
        alertas.append(('ALERTA', '', 'El reporte de facturación no trae nombres (ni Razón Social ni Nombre Estudiante): el cruce con el banco solo puede usar el monto y todas las coincidencias quedan en nivel DÉBIL; confírmelas una a una.'))
    # C.I.: facturas QR validas reclasificadas por el cajero salen del cruce QR (antes de conciliar)
    ci_decl = {}
    for nf, d in (reclasificaciones or {}).items():
        nf = str(nf); m = fa.index[(fa['Número Factura'] == nf) & (fa['estado_prep'] == 'VALIDA') & (fa['tipo'] == 'QR')]
        if len(m) == 0:
            alertas.append(('ALERTA', nf, f"Reclasificación no aplicada: la factura {nf} ya no es una factura QR válida. Revise el banco y el código que registró para ella.")); continue
        fa.at[m[0], 'tipo'] = 'C.I.'; ci_decl[nf] = {'banco': str(d.get('banco', '')), 'codigo': str(d.get('codigo', '')).strip(), 'fecha_deposito': str(d.get('fecha_deposito') or '').strip()}
    manuales = {str(x) for x in (respaldos_manuales or [])}
    bk = preparar_a2(a2)
    qr = fa[(fa['estado_prep'] == 'VALIDA') & (fa['tipo'] == 'QR')]
    tj = fa[(fa['estado_prep'] == 'VALIDA') & (fa['tipo'] == 'TARJETA')]

    # R3 fecha del cierre = fecha(s) de las facturas validas
    fechas = sorted({d.date() for d in fa[fa['estado_prep'] == 'VALIDA']['Fecha']})
    if len(fechas) > 1: alertas.append(('ALERTA', '', f"El reporte de facturas tiene más de una fecha ({', '.join(f'{d:%d/%m/%Y}' for d in fechas)}). Se usan todas; verifique que cargó el reporte del día correcto."))
    if not fechas: alertas.append(('ALERTA', '', 'No hay facturas válidas en el reporte del sistema, así que no se puede determinar la fecha del cierre. Verifique el archivo cargado.'))

    # estado de cada movimiento
    bk['estado_banco'] = ''; bk['factura_asignada'] = ''; bk['nivel_asignado'] = ''
    usable = []
    for i, m in bk.iterrows():
        if m['dup']:
            bk.at[i, 'estado_banco'] = 'DUPLICADO EN ARCHIVO BANCO'
            alertas.append(('ALERTA', m['ID TRANSACCION'], f"El número de transacción del banco {m['ID TRANSACCION']} aparece repetido en el reporte del banco; solo se usa la primera vez. Revise que el banco no haya duplicado el movimiento."))
        elif m['FECHA / HORA'].date() not in fechas:
            bk.at[i, 'estado_banco'] = 'FUERA DE FECHA'
            alertas.append(('INFO', m['ID TRANSACCION'], f"El movimiento del banco {m['ID TRANSACCION']} ({m['TITULAR ORIGEN']}, Bs {m['MONTO']:.2f}) es del {m['FECHA / HORA']:%d/%m/%Y}, distinto a la fecha del cierre, y no se usó."))
        else: usable.append(i)

    # descartes del cajero: pares (factura, movimiento) que no deben cruzarse
    rech = {}
    for nf, ids in (rechazos or {}).items():
        for idt in (ids if isinstance(ids, (list, tuple, set)) else [ids]):
            rech.setdefault(str(nf), set()).add(str(idt))
    descartes = []
    for nf, ids in rech.items():
        mf = fa.index[(fa['Número Factura'] == nf) & (fa['estado_prep'] == 'VALIDA')]
        for idt in sorted(ids):
            mb = bk.index[(bk['ID TRANSACCION'] == idt) & (~bk['dup'])]
            if len(mf) == 0 or len(mb) == 0:
                alertas.append(('ALERTA', nf, f"Descarte no aplicado: la factura {nf} o el número de transacción {idt} ya no existen en los archivos cargados."))
                continue
            m = bk.loc[mb[0]]
            descartes.append({'factura': nf, 'id': idt, 'titular': m['TITULAR ORIGEN'], 'hora': m['FECHA / HORA'], 'monto': float(m['MONTO'])})

    # pares con monto exacto (R4) y nivel (R7)
    pares = []
    for fi, f in qr.iterrows():
        hf = f['Fecha'].replace(microsecond=0)
        rf = rech.get(str(f['Número Factura']), ())
        for bi in usable:
            m = bk.loc[bi]
            if abs(f['Monto'] - m['MONTO']) > 1e-9: continue
            if m['ID TRANSACCION'] in rf: continue         # descartado por el cajero: se trata como si no existiera
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
    for fi in fa.index[fa['tipo'] == 'C.I.']: fa.at[fi, 'estado_conciliacion'] = 'COMUNICACIÓN INTERNA'
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
            alertas.append(('ALERTA', nf, f"Confirmación no aplicada: la factura {nf} ya no tiene una coincidencia pendiente de confirmar. Revise su estado en Observadas."))
            continue
        bi = bk.index[(bk['ID TRANSACCION'] == fa.at[fi, 'ID TRANSACCION']) & (~bk['dup'])][0]
        fa.at[fi, 'estado_conciliacion'] = 'CONFIRMADA POR CAJERO'; bk.at[bi, 'estado_banco'] = 'ASIGNADO CONFIRMADA POR CAJERO'
    for nf, idt in (elecciones or {}).items():
        nf = str(nf); idt = str(idt); fi = idx_f.get(nf)
        def rechazo(msg): alertas.append(('ALERTA', nf, f"Elección no aplicada (factura {nf}, número de transacción {idt}): {msg}"))
        if fi is None or fa.at[fi, 'estado_conciliacion'] not in ('SIN RESPALDO', 'AMBIGUO'):
            rechazo('la factura no está pendiente de asignar movimiento del banco'); continue
        cand = bk.index[(bk['ID TRANSACCION'] == idt) & (~bk['dup'])]
        if len(cand) == 0: rechazo('ese número de transacción es inexistente en el reporte del banco'); continue
        bi = cand[0]
        if abs(bk.at[bi, 'MONTO'] - fa.at[fi, 'Monto']) > 1e-9: rechazo(f"el movimiento tiene un monto distinto al de la factura (banco {bk.at[bi, 'MONTO']:.2f}, factura {fa.at[fi, 'Monto']:.2f})"); continue
        if bk.at[bi, 'FECHA / HORA'].date() not in fechas: rechazo('el movimiento es de una fecha distinta a la del cierre'); continue
        if idt in rech.get(nf, ()): rechazo('usted descartó ese pago para esta factura'); continue
        if bk.at[bi, 'estado_banco'].startswith('ASIGNADO'): rechazo(f"ese movimiento del banco ya fue usado por la factura {bk.at[bi, 'factura_asignada']}"); continue
        poner(fi, bi, 'CONFIRMADA POR CAJERO', 'RESPALDO MANUAL' if nf in manuales else 'ELECCIÓN CAJERO')
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
            alertas.append(('ALERTA', nf, f"La factura {nf} coincide con varios movimientos del banco ({', '.join(ids)}). Elija en Observadas cuál le corresponde."))
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
    ci_rows = []
    for fi in fa.index[fa['tipo'] == 'C.I.']:
        nf = str(fa.at[fi, 'Número Factura']); d = ci_decl[nf]
        ci_rows.append({'factura': nf, 'razon_social': fa.at[fi, 'Razon Social'], 'estudiante': fa.at[fi, 'Nombre Estudiante'],
                        'monto': float(fa.at[fi, 'Monto']), 'banco': d['banco'], 'codigo': d['codigo'], 'fecha_deposito': d['fecha_deposito']})
    ci_sis = float(sum(r['monto'] for r in ci_rows)); ci_resp = float(sum(r['monto'] for r in ci_rows if r['codigo']))
    tot = {'fecha_cierre': fechas[0] if fechas else None, 'n_facturas_qr': len(qr), 'QR registrado': reg,
           'QR conciliado automático (CONCILIADA)': auto, 'QR confirmado por cajero': conf,
           'QR pendiente de confirmar (POR CONFIRMAR + AMBIGUO)': pend, 'QR sin respaldo': sinr,
           'Diferencia QR': dif_qr, 'Estado QR': estado_qr, 'Tarjeta según sistema': tsis,
           'Tarjeta declarada': float(tarjeta_declarada), 'Diferencia tarjeta (declarada - sistema)': float(tarjeta_declarada) - tsis,
           'C.I. según sistema': ci_sis, 'C.I. con código de asignación': ci_resp, 'n_COMUNICACIÓN INTERNA': len(ci_rows),
           'Detalle C.I.': '; '.join(f"{r['banco']} {r['codigo']}" for r in ci_rows),
           'Total cierre': reg + ci_sis + float(tarjeta_declarada),
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
    return {'facturas': facturas, 'banco': banco, 'resumen': resumen, 'alertas': alert_df, 'ci': ci_rows, 'descartes': descartes,
            'tarjetas': facturas[facturas['estado_conciliacion'] == 'TARJETA'].copy(), 'totales': tot}

def exportar(res, ruta):
    f = res['facturas']; b = res['banco']
    with pd.ExcelWriter(ruta) as w:
        res['resumen'].to_excel(w, sheet_name='Resumen', index=False)
        f[f['Tipo (motor)'] == 'QR'].to_excel(w, sheet_name='Facturas_QR', index=False)
        b[~b['estado_banco'].str.startswith('ASIGNADO')].to_excel(w, sheet_name='Movimientos_ajenos', index=False)
        if res.get('ci'): pd.DataFrame(res['ci']).to_excel(w, sheet_name='Comunicaciones_internas', index=False)
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
