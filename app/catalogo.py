"""Catalogo de bancos para el respaldo manual de facturas observadas.

- 'QR BANECO': el codigo es un ID TRANSACCION del reporte QR cargado (lo valida el motor).
- Los demas bancos son depositos a otra cuenta: la factura se reclasifica como Comunicacion interna (C.I.)
  y el codigo es el codigo de asignacion del deposito.
`formato_asignacion` es un regex (str) por banco. Hoy es None para todos porque los formatos oficiales AUN NO
estan definidos; con None solo se exige 3-40 caracteres alfanumericos y - / .  Para activar un formato oficial,
basta poner el regex del banco (se aplica con fullmatch sobre el codigo en MAYUSCULAS, sin espacios laterales)."""
import re

BANCO_QR = 'QR BANECO'

BANCOS = {
    'QR BANECO': {'cuenta_bancaria': None, 'cuenta_sap': None, 'formato_asignacion': None, 'es_qr': True},
    'BNB': {'cuenta_bancaria': '8000006425', 'cuenta_sap': '110103222', 'formato_asignacion': None, 'es_qr': False},
    'BMSC': {'cuenta_bancaria': '4010884831', 'cuenta_sap': '110103242', 'formato_asignacion': None, 'es_qr': False},
    'BISA': {'cuenta_bancaria': '696870021', 'cuenta_sap': '110103232', 'formato_asignacion': None, 'es_qr': False},
    'UNION': {'cuenta_bancaria': '1-56089490', 'cuenta_sap': '110103382', 'formato_asignacion': None, 'es_qr': False},
}
LISTA_BANCOS = list(BANCOS)

FORMATO_GENERICO = re.compile(r'[A-Za-z0-9\-/.]{3,40}')
MSG_GENERICO = 'El código de asignación debe tener entre 3 y 40 caracteres: letras, números, guion, barra o punto.'


def es_qr(banco):
    return banco == BANCO_QR


def normalizar_codigo(codigo):
    return (codigo or '').strip().upper()


def validar_codigo(banco, codigo):
    """Valida el codigo de asignacion de una C.I. Devuelve (ok, mensaje_o_None, codigo_normalizado)."""
    if banco not in BANCOS or es_qr(banco): return False, 'Banco no válido para comunicación interna.', ''
    c = normalizar_codigo(codigo)
    if not c: return False, 'Escriba el código de asignación.', ''
    patron = BANCOS[banco]['formato_asignacion']
    if patron:
        if not re.fullmatch(patron, c): return False, f'El código no tiene el formato de {banco}.', c
    elif not FORMATO_GENERICO.fullmatch(c):
        return False, MSG_GENERICO, c
    return True, None, c
