import math, re
from datetime import date, datetime


def bs(x):
    """10784 -> '10,784.00'"""
    return f"{float(x):,.2f}"


def fecha(x):
    if x is None or x == '': return ''
    if isinstance(x, str):
        try: x = datetime.fromisoformat(x)
        except ValueError: return x
    return x.strftime('%d/%m/%Y')


def fecha_hora(x):
    if x is None or x == '': return ''
    if isinstance(x, str):
        try: x = datetime.fromisoformat(x)
        except ValueError: return x
    return x.strftime('%d/%m/%Y %H:%M')


def hora(x):
    try: return x.strftime('%H:%M:%S')
    except Exception: return ''


def parsear_monto(txt):
    """Acepta '1191', '1,191.00', '1191.50', '1191,50'. Lanza ValueError si no es un monto valido (>= 0)."""
    t = re.sub(r'\s|Bs\.?|BOB', '', str(txt or ''), flags=re.I)
    if not t: raise ValueError('vacío')
    if re.fullmatch(r'\d+,\d{1,2}', t): t = t.replace(',', '.')
    else: t = t.replace(',', '')
    if not re.fullmatch(r'\d+(\.\d+)?', t): raise ValueError('formato')
    v = float(t)
    if not math.isfinite(v) or v < 0 or v > 1e9: raise ValueError('rango')
    return round(v, 2)
