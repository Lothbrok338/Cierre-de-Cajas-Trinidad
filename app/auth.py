"""Autenticacion provisional (hasta integrar Microsoft Entra ID).
APP_USERS = JSON {"usuario": {"nombre": "...", "rol": "cajero"|"ingresos", "clave_hash": "pbkdf2_sha256$iter$salt$hash"}}"""
import base64, hashlib, hmac, json, os, secrets, time

ROLES = ('cajero', 'ingresos')
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
        out[str(u)] = {'nombre': str(d.get('nombre') or u), 'rol': d['rol'], 'clave_hash': d['clave_hash']}
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


def autenticar(usuarios, usuario, clave):
    """Devuelve el dict de usuario si las credenciales son validas, si no None (tiempo similar en ambos casos)."""
    u = usuarios.get(usuario)
    if u is None:
        verificar_clave(clave, _HASH_FALSO); return None
    return u if verificar_clave(clave, u['clave_hash']) else None
