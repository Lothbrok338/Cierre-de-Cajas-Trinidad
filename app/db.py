import json, sqlite3, os
from datetime import datetime

ESQUEMA = """
CREATE TABLE IF NOT EXISTS cierres (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  correlativo TEXT NOT NULL UNIQUE,
  fecha_cierre TEXT NOT NULL,
  usuario TEXT NOT NULL,
  nombre TEXT,
  caja TEXT,
  qr_registrado REAL, qr_conciliado REAL, tarjeta_declarada REAL, tarjeta_sistema REAL, total_cierre REAL,
  estado_revision TEXT NOT NULL DEFAULT 'PENDIENTE DE REVISIÓN',
  version INTEGER NOT NULL,
  reemplaza_a INTEGER,
  motivo_reemplazo TEXT,
  presentado_en TEXT NOT NULL,
  ruta TEXT NOT NULL,
  sha256_movimiento TEXT, sha256_banco TEXT
);
CREATE INDEX IF NOT EXISTS ix_cierres_fu ON cierres(fecha_cierre, usuario);
CREATE TABLE IF NOT EXISTS auditoria (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  momento TEXT NOT NULL, usuario TEXT, accion TEXT NOT NULL, detalle_json TEXT
);
CREATE TABLE IF NOT EXISTS cierre_movimientos (
  cierre_id INTEGER NOT NULL REFERENCES cierres(id),
  id_transaccion TEXT NOT NULL, factura TEXT, monto REAL
);
CREATE INDEX IF NOT EXISTS ix_cm_id ON cierre_movimientos(id_transaccion);
"""


def ahora():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def conectar(ruta):
    con = sqlite3.connect(ruta, timeout=30, isolation_level=None)
    con.row_factory = sqlite3.Row
    return con


def inicializar(ruta):
    os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
    con = conectar(ruta)
    con.executescript(ESQUEMA); con.close()


def auditar(ruta, usuario, accion, detalle=None):
    con = conectar(ruta)
    try:
        con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                    (ahora(), usuario, accion, json.dumps(detalle or {}, ensure_ascii=False, default=str)))
    finally:
        con.close()
