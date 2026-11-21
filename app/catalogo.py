"""Catalogo de bancos para el respaldo manual de facturas observadas.

- 'QR BANECO': el codigo es un ID TRANSACCION del reporte QR cargado (lo valida el motor).
- Los demas bancos son depositos a otra cuenta: la factura se reclasifica como Comunicacion interna (C.I.)
  y el codigo es el codigo de asignacion del deposito.

Formato del codigo por banco (`formato_asignacion`, regex que se aplica con fullmatch sobre el codigo ya normalizado
a MAYUSCULAS, y `ayuda`, el texto de la regla). Un codigo que NO cumple el formato NO se rechaza: se registra igual y
genera la alerta de revision "FORMATO A REVISAR" (ver `formato_incumplido`). Lo unico bloqueante es: codigo vacio,
longitud fuera de 3-40, caracteres de control y los antiduplicados (en servicio)."""
import re

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
}
LISTA_BANCOS = list(BANCOS)

LARGO_MIN, LARGO_MAX = 3, 40
MSG_LARGO = f'El código de asignación debe tener entre {LARGO_MIN} y {LARGO_MAX} caracteres.'


def es_qr(banco):
    return banco == BANCO_QR


def normalizar_codigo(codigo):
    return (codigo or '').strip().upper()


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
