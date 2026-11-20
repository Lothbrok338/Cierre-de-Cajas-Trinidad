import os, sys, re, glob, pytest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fastapi.testclient import TestClient
from app.auth import hash_clave
import json

USUARIOS = {
    'ana': {'nombre': 'Ana Prueba', 'rol': 'cajero', 'clave_hash': hash_clave('clave-ana-1', 1000)},
    'beto': {'nombre': 'Beto Prueba', 'rol': 'cajero', 'clave_hash': hash_clave('clave-beto-1', 1000)},
    'ing': {'nombre': 'Ingresos Prueba', 'rol': 'ingresos', 'clave_hash': hash_clave('clave-ing-1', 1000)},
}
CLAVES = {'ana': 'clave-ana-1', 'beto': 'clave-beto-1', 'ing': 'clave-ing-1'}


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setenv('DATA_DIR', str(tmp_path / 'data'))
    monkeypatch.setenv('APP_USERS', json.dumps(USUARIOS))
    monkeypatch.setenv('SECRET_KEY', 'test-secret')
    monkeypatch.setenv('COOKIE_SECURE', '0')
    from app.main import create_app
    return create_app()


def nuevo_cliente(app):
    return TestClient(app, follow_redirects=False)


def token(c, ruta='/login'):
    r = c.get(ruta); m = re.search(r'name="csrf" value="([^"]+)"', r.text); return m.group(1)


def login(app, usuario):
    c = nuevo_cliente(app)
    r = c.post('/login', data={'usuario': usuario, 'clave': CLAVES[usuario], 'csrf': token(c)})
    assert r.status_code == 303
    return c


def post(c, url, ruta_csrf, **data):
    return c.post(url, data={'csrf': token(c, ruta_csrf), **data})


@pytest.fixture
def sint(tmp_path):
    from tests import sinteticos as s
    d = tmp_path / 'sint'; d.mkdir()
    return {'a1': s.crear_a1(str(d / 'mov.xlsx')), 'a2': s.crear_a2(str(d / 'banco.xlsx')),
            'a2_sin': s.crear_a2(str(d / 'banco_sin.xlsx'), quitar=('700000002',)), 'dir': d}


def subir(c, a1, a2, tarjeta='200'):
    tk = token(c, '/nuevo')
    with open(a1, 'rb') as f1, open(a2, 'rb') as f2:
        return c.post('/nuevo', data={'csrf': tk, 'tarjeta': tarjeta},
                      files={'movimiento': (os.path.basename(a1), f1), 'banco': (os.path.basename(a2), f2)})
