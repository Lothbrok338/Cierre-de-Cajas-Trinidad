"""Autenticacion provisional (hasta integrar Microsoft Entra ID).
Los usuarios viven en la tabla `usuarios` (ver servicio.Almacen). APP_USERS solo SIEMBRA usuarios al arrancar:
APP_USERS = JSON {"usuario": {"nombre": "...", "rol": "superadmin"|"ingresos"|"cajero", "clave_hash": "pbkdf2_sha256$iter$salt$hash",
                              "forzar": false}}   (forzar=true vuelve a aplicar esa entrada aunque el usuario ya exista)"""
import base64, hashlib, hmac, json, os, secrets, time

ROLES = ('superadmin', 'ingresos', 'cajero')
MIN_CLAVE = 8
ITER_DEFECTO = 600_000


def hash_clave(clave: str, iteraciones: int = ITER_DEFECTO) -> str:
    sal = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac('sha256', clave.encode('utf-8'), sal, iteraciones)
    return f"pbkdf2_sha256${iteraciones}${base64.b64encode(sal).decode()}${base64.b64encode(dk).decode()}"


def verificar_clave(clave: str, almacenado: str) -> bool:
    try:
        alg, it, sal, h = almacenado.split('$')
        if alg != 'pbkdf2_sha256': return False
        dk = hashlib.pbkdf2_hmac('sha256', clave.encode('utf-8'), base64.b64decode(sal), int(it))
        return hmac.compare_digest(dk, base64.b64decode(h))
    except Exception:
        return False


def cargar_usuarios(texto):
    """Devuelve dict de usuarios validos, o None si APP_USERS no esta definido / es invalido."""
    if not texto or not texto.strip(): return None
    try:
        data = json.loads(texto)
    except ValueError:
        return None
    if not isinstance(data, dict) or not data: return None
    out = {}
    for u, d in data.items():
        if not isinstance(d, dict) or d.get('rol') not in ROLES or not d.get('clave_hash'): continue
        k = str(u).strip().lower()
        if not k: continue
        out[k] = {'nombre': str(d.get('nombre') or u), 'rol': d['rol'], 'clave_hash': d['clave_hash'], 'forzar': d.get('forzar') is True}
    return out or None


_HASH_FALSO = hash_clave('x', 1000)


class LimiteIntentos:
    """Bloqueo simple en memoria: max_fallos por usuario en ventana de segundos."""
    def __init__(self, max_fallos=5, ventana=300):
        self.max, self.ventana, self.fallos = max_fallos, ventana, {}

    def _limpiar(self, u):
        ahora = time.time()
        self.fallos[u] = [t for t in self.fallos.get(u, []) if ahora - t < self.ventana]
        return self.fallos[u]

    def bloqueado(self, u): return len(self._limpiar(u)) >= self.max
    def registrar_fallo(self, u): self._limpiar(u).append(time.time())
    def limpiar(self, u): self.fallos.pop(u, None)


def autenticar(buscar, usuario, clave):
    """`buscar(usuario_en_minusculas)` devuelve el registro (dict/Row con clave_hash y activo) o None.
    Devuelve ese registro si las credenciales son validas y el usuario esta activo, si no None (tiempo similar en todos los casos)."""
    u = buscar(usuario.strip().lower())
    if u is None:
        verificar_clave(clave, _HASH_FALSO); return None
    ok = verificar_clave(clave, u['clave_hash'])
    return u if ok and u['activo'] else None


def validar_clave_nueva(clave, confirmacion):
    """Mensaje de error o None."""
    if clave != confirmacion: return 'La confirmación de la clave no coincide.'
    if len(clave) < MIN_CLAVE: return f'La clave debe tener al menos {MIN_CLAVE} caracteres.'
    return None
