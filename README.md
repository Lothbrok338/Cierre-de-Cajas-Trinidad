# Cierre Diario de Caja (perspectiva cajero)

FastAPI + Jinja2 + SQLite. Conciliación de facturas QR BANECO contra el reporte del banco; presentación de cierres inmutables con PDF.

## Ejecutar local
    pip install -r requirements-dev.txt
    python scripts/hash_clave.py --json ana "Ana Pérez" cajero     # genera la entrada para APP_USERS
    export APP_USERS='{"ana": {...}}' SECRET_KEY=$(python -c "import secrets;print(secrets.token_hex(32))")
    uvicorn app.main:app --reload
Pruebas: `python -m pytest tests` (los tests con datos reales requieren `DATOS_PRUEBA_DIR`, si no se omiten).

## Variables de entorno
- `APP_USERS` (JSON usuarios; solo SIEMBRA la tabla `usuarios` al arrancar con los que aún no existen; con `"forzar": true` en una entrada se vuelve a aplicar siempre; si no hay ningún usuario activo, el login muestra "Sistema no configurado")
- `SECRET_KEY` (firma de cookie de sesión; obligatoria en producción)
- `DATA_DIR` (default `./data`; en Railway `/data` sobre un volumen)
- `COOKIE_SECURE` (`auto` por defecto: seguro si hay variables RAILWAY_*; `1`/`0` para forzar)
- `PORT` (lo define Railway)

Los reportes reales nunca van al repositorio (ver `.gitignore`).

## v2: roles, revisión y respaldo manual
- Roles: `cajero` (crea cierres, corrige los RECHAZADOS) e `ingresos` (ve todos, filtra por estado/fecha, aprueba o rechaza con observación de al menos 10 caracteres).
- Estados: PENDIENTE DE REVISIÓN → CONFORME | RECHAZADO; un RECHAZADO se corrige con «Editar y reenviar» (crea la versión siguiente y deja la anterior como REEMPLAZADO). Un CONFORME no se modifica.
- Respaldo manual de facturas observadas (`app/catalogo.py`): QR BANECO (ID de transacción del reporte) o depósito a otro banco (se reclasifica como Comunicación interna; el par banco/código no puede repetirse en cierres vigentes). El formato del código por banco (`formato_asignacion` + `ayuda`) es una alerta de revisión no bloqueante: BNB alfanumérico con P, BMSC alfanumérico con T, BISA y UNION solo dígitos, QR BANECO (ID manual) solo dígitos.
- Base de datos: `app/db.py` migra el esquema v1 al arrancar de forma idempotente (columnas de revisión y tabla `cierre_ci`), sin perder datos.
- Iconos: `python scripts/generar_iconos.py` regenera favicon y apple-touch-icon desde `scripts/escudo_fuente.png` (reemplace ese archivo por la versión a color).

## v3: tres roles y administración
- Roles: `superadmin` (todo lo de cajero e ingresos, más revertir revisiones, administración de usuarios y auditoría), `ingresos` y `cajero`.
- Usuarios en la tabla `usuarios` (login sin distinguir mayúsculas; se guarda en minúsculas). `APP_USERS` solo siembra los que faltan, así los cambios hechos en la UI no se pisan al reiniciar. `python scripts/hash_clave.py --json usuario "Nombre" superadmin` genera entradas (agregue `"forzar": true` solo si quiere reaplicarla).
- Superadmin: «Revertir revisión» (CONFORME o RECHAZADO a PENDIENTE, motivo de 10+ caracteres, no si está REEMPLAZADO; el historial queda en `revisiones_historial` y en la auditoría), `/admin/usuarios` (crear, rol, activar/desactivar, restablecer clave; siempre queda al menos un superadmin activo) y `/admin/auditoria` (50 por página, filtros por usuario, acción y fecha).
- «Mi cuenta» (`/cuenta`): cualquier usuario cambia su clave (mínimo 8 caracteres).
- Base de datos: esquema 3 (tablas `usuarios` y `revisiones_historial`); la migración es idempotente y pasa las revisiones v2 al historial.
