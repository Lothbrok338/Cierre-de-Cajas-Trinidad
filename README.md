# Cierre Diario de Caja (perspectiva cajero)

FastAPI + Jinja2 + SQLite. Conciliación de facturas QR BANECO contra el reporte del banco; presentación de cierres inmutables con PDF.

## Ejecutar local
    pip install -r requirements-dev.txt
    python scripts/hash_clave.py --json ana "Ana Pérez" cajero     # genera la entrada para APP_USERS
    export APP_USERS='{"ana": {...}}' SECRET_KEY=$(python -c "import secrets;print(secrets.token_hex(32))")
    uvicorn app.main:app --reload
Pruebas: `python -m pytest tests` (los tests con datos reales requieren `DATOS_PRUEBA_DIR`, si no se omiten).

## Variables de entorno
- `APP_USERS` (JSON usuarios; si falta, el login muestra "Sistema no configurado")
- `SECRET_KEY` (firma de cookie de sesión; obligatoria en producción)
- `DATA_DIR` (default `./data`; en Railway `/data` sobre un volumen)
- `COOKIE_SECURE` (`auto` por defecto: seguro si hay variables RAILWAY_*; `1`/`0` para forzar)
- `PORT` (lo define Railway)

Los reportes reales nunca van al repositorio (ver `.gitignore`).
