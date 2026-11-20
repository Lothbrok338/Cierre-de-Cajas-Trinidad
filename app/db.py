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
  sha256_movimiento TEXT, sha256_banco TEXT,
  ci_total REAL,
  revisado_por TEXT, revisado_en TEXT, observacion_revision TEXT
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
CREATE TABLE IF NOT EXISTS cierre_ci (
  cierre_id INTEGER NOT NULL REFERENCES cierres(id),
  factura TEXT, banco TEXT NOT NULL, codigo TEXT NOT NULL, monto REAL
);
CREATE INDEX IF NOT EXISTS ix_ci_bc ON cierre_ci(banco, codigo);
CREATE TABLE IF NOT EXISTS usuarios (
  usuario TEXT PRIMARY KEY,
  nombre TEXT NOT NULL,
  rol TEXT NOT NULL,
  clave_hash TEXT NOT NULL,
  activo INTEGER NOT NULL DEFAULT 1,
  creado_en TEXT NOT NULL,
  actualizado_en TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS revisiones_historial (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  cierre_id INTEGER NOT NULL REFERENCES cierres(id),
  momento TEXT NOT NULL, usuario TEXT, accion TEXT NOT NULL,
  estado_anterior TEXT, estado_nuevo TEXT, observacion TEXT
);
CREATE INDEX IF NOT EXISTS ix_rh_cierre ON revisiones_historial(cierre_id);
CREATE INDEX IF NOT EXISTS ix_aud_momento ON auditoria(momento);
"""


def ahora():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def conectar(ruta):
    con = sqlite3.connect(ruta, timeout=30, isolation_level=None)
    con.row_factory = sqlite3.Row
    return con


# columnas agregadas despues del esquema v1: (tabla, columna, definicion)
MIGRACIONES = [
    ('cierres', 'ci_total', 'REAL'),
    ('cierres', 'revisado_por', 'TEXT'),
    ('cierres', 'revisado_en', 'TEXT'),
    ('cierres', 'observacion_revision', 'TEXT'),
]
VERSION_ESQUEMA = 3


def migrar(con):
    """Migracion idempotente: agrega las columnas que falten (no toca los datos existentes)."""
    for tabla, col, definicion in MIGRACIONES:
        cols = {r[1] for r in con.execute(f'PRAGMA table_info({tabla})')}
        if col not in cols:
            try: con.execute(f'ALTER TABLE {tabla} ADD COLUMN {col} {definicion}')
            except sqlite3.OperationalError as e:
                if 'duplicate column' not in str(e).lower(): raise    # otro proceso la agrego antes
    # v3: las revisiones ya hechas (v2) pasan al historial, una sola vez por cierre
    con.execute("INSERT INTO revisiones_historial (cierre_id, momento, usuario, accion, estado_anterior, estado_nuevo, observacion) "
                "SELECT id, revisado_en, revisado_por, CASE estado_revision WHEN 'CONFORME' THEN 'aprobar' ELSE 'rechazar' END, "
                "'PENDIENTE DE REVISIÓN', CASE estado_revision WHEN 'CONFORME' THEN 'CONFORME' ELSE 'RECHAZADO' END, observacion_revision FROM cierres "
                "WHERE revisado_por IS NOT NULL "
                "AND id NOT IN (SELECT cierre_id FROM revisiones_historial)")
    con.execute(f'PRAGMA user_version = {VERSION_ESQUEMA}')


def inicializar(ruta):
    os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
    con = conectar(ruta)
    try:
        con.executescript(ESQUEMA)   # CREATE ... IF NOT EXISTS: en una DB v1 no cambia las tablas existentes
        migrar(con)
    finally: con.close()


def auditar(ruta, usuario, accion, detalle=None):
    con = conectar(ruta)
    try:
        con.execute('INSERT INTO auditoria (momento, usuario, accion, detalle_json) VALUES (?,?,?,?)',
                    (ahora(), usuario, accion, json.dumps(detalle or {}, ensure_ascii=False, default=str)))
    finally:
        con.close()
