"""Catalogo de bancos para el respaldo manual de facturas observadas.

- 'QR BANECO': el codigo es un ID TRANSACCION del reporte QR cargado (lo valida el motor).
- Los demas bancos son depositos a otra cuenta: la factura se reclasifica como Comunicacion interna (C.I.)
  y el codigo es el codigo de asignacion del deposito.
- 'BANECO' (C.I.) es DISTINTO de 'QR BANECO': es un deposito hecho a la cuenta BANECO en una fecha anterior, que NO esta en el reporte QR
  del dia; el codigo es el numero de transaccion del banco y no puede coincidir con ningun ID del reporte QR ni usarse en otro cierre vigente.

Formato del codigo por banco (`formato_asignacion`, regex que se aplica con fullmatch sobre el codigo ya normalizado
a MAYUSCULAS, y `ayuda`, el texto de la regla). Un codigo que NO cumple el formato NO se rechaza: se registra igual y
genera la alerta de revision "FORMATO A REVISAR" (ver `formato_incumplido`). Lo unico bloqueante es: codigo vacio,
longitud fuera de 3-40, caracteres de control y los antiduplicados (en servicio)."""
import re
from datetime import date, datetime

BANCO_QR = 'QR BANECO'

BANCOS = {
    'QR BANECO': {'cuenta_bancaria': None, 'cuenta_sap': None, 'formato_asignacion': r'[0-9]+', 'es_qr': True,
                  'ayuda': 'QR BANECO: solo números (ID de transacción)'},
    'BNB': {'cuenta_bancaria': '8000006425', 'cuenta_sap': '110103222', 'formato_asignacion': r'(?=.*P)[A-Z0-9]+', 'es_qr': False,
            'ayuda': 'BNB: letras y números, debe incluir una P'},
    'BMSC': {'cuenta_bancaria': '4010884831', 'cuenta_sap': '110103242', 'formato_asignacion': r'(?=.*T)[A-Z0-9]+', 'es_qr': False,
            'ayuda': 'BMSC: letras y números, debe incluir una T'},
    'BISA': {'cuenta_bancaria': '696870021', 'cuenta_sap': '110103232', 'formato_asignacion': r'[0-9]+', 'es_qr': False,
            'ayuda': 'BISA: solo números'},
    'UNION': {'cuenta_bancaria': '1-56089490', 'cuenta_sap': '110103382', 'formato_asignacion': r'[0-9]+', 'es_qr': False,
            'ayuda': 'UNION: solo números'},
    'BANECO': {'cuenta_bancaria': None, 'cuenta_sap': None, 'formato_asignacion': r'[0-9]+', 'es_qr': False,
               'ayuda': 'BANECO: solo números (número de transacción del banco)'},
}
# Texto de las listas desplegables (el valor interno sigue siendo la clave de BANCOS)
ETIQUETAS = {'QR BANECO': 'QR BANECO (pago del día, está en el reporte)', 'BANECO': 'BANECO (depósito de fecha anterior)'}
BANCO_BANECO_CI = 'BANECO'
TIPO_FECHA = 'FECHA DEPÓSITO'
MSG_FECHA_FALTA = 'Falta la fecha del depósito: Ingresos la necesita para ubicarlo en el extracto del mes correcto.'
MSG_BANECO_EN_REPORTE = ('Ese número de transacción aparece en el reporte QR de este día: respalde la factura como QR BANECO, '
                         'no como comunicación interna.')
LISTA_BANCOS = list(BANCOS)

LARGO_MIN, LARGO_MAX = 3, 40
MSG_LARGO = f'El código de asignación debe tener entre {LARGO_MIN} y {LARGO_MAX} caracteres.'


def es_qr(banco):
    return banco == BANCO_QR


def etiqueta(banco):
    return ETIQUETAS.get(banco, banco)


def normalizar_codigo(codigo):
    """Recorta espacios, pasa a mayusculas y quita un '.0' final cuando el resto son solo digitos (Excel: 123456.0 -> 123456)."""
    c = str(codigo if codigo is not None else '').strip().upper()
    m = re.fullmatch(r'(\d+)\.0+', c)
    return m.group(1) if m else c


_FORMATOS_FECHA = ('%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%Y-%m-%d')


def parsear_fecha_deposito(texto):
    """date o None si el texto no es una fecha valida (dd/mm/aaaa; tambien se aceptan dd-mm-aaaa, aaaa-mm-dd y fecha ISO con hora)."""
    t = str(texto or '').strip()
    if 'T' in t[:11] or (len(t) > 10 and t[10:11] == ' '): t = t[:10]
    for f in _FORMATOS_FECHA:
        try: return datetime.strptime(t, f).date()
        except ValueError: pass
    return None


def normalizar_fecha_deposito(texto):
    """Fecha valida -> 'dd/mm/aaaa'; vacia -> ''; texto no valido -> el mismo texto recortado (maximo 30 caracteres, sin control) para mostrarlo en la alerta."""
    t = ' '.join(str(texto or '').split())
    if not t: return ''
    d = parsear_fecha_deposito(t)
    return d.strftime('%d/%m/%Y') if d else ''.join(ch for ch in t if ch.isprintable())[:30]


def alerta_fecha_deposito(banco, texto, fecha_cierre):
    """Alerta NO bloqueante (texto) sobre la fecha de deposito de una C.I., o None: BANECO sin fecha, fecha invalida o posterior a la fecha del cierre."""
    t = normalizar_fecha_deposito(texto)
    if not t: return MSG_FECHA_FALTA if banco == BANCO_BANECO_CI else None
    d = parsear_fecha_deposito(t)
    if d is None: return f'La fecha del depósito «{t}» no es válida: use el formato dd/mm/aaaa.'
    fc = fecha_cierre.date() if isinstance(fecha_cierre, datetime) else fecha_cierre
    if isinstance(fc, date) and d > fc: return f'La fecha del depósito ({d:%d/%m/%Y}) es posterior a la fecha del cierre ({fc:%d/%m/%Y}).'
    return None


def regla_formato(banco):
    """Texto de ayuda de la regla de formato del banco (o '' si no hay)."""
    return (BANCOS.get(banco) or {}).get('ayuda', '')


def formato_incumplido(banco, codigo):
    """Regla esperada (texto) si el codigo NO cumple el formato del banco; None si cumple (o el banco no tiene formato).
    Es solo una alerta de revision: nunca bloquea."""
    patron = (BANCOS.get(banco) or {}).get('formato_asignacion')
    if not patron: return None
    return None if re.fullmatch(patron, normalizar_codigo(codigo)) else regla_formato(banco)


def validar_codigo(banco, codigo):
    """Validacion BLOQUEANTE del codigo de una C.I. (vacio, longitud 3-40, sin caracteres de control).
    El formato por banco NO se valida aqui (es alerta, ver `formato_incumplido`). Devuelve (ok, mensaje_o_None, codigo_normalizado)."""
    if banco not in BANCOS or es_qr(banco): return False, 'Banco no válido para comunicación interna.', ''
    c = normalizar_codigo(codigo)
    if not c: return False, 'Escriba el código de asignación.', ''
    if not (LARGO_MIN <= len(c) <= LARGO_MAX) or not c.isprintable(): return False, MSG_LARGO, c
    return True, None, c
