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
- Respaldo manual de facturas observadas (`app/catalogo.py`): QR BANECO (ID de transacción del reporte) o depósito a otro banco (se reclasifica como Comunicación interna; el par banco/código no puede repetirse en cierres vigentes). El formato del código por banco (`formato_asignacion` + `ayuda`) es una alerta de revisión no bloqueante: BNB alfanumérico con P, BMSC alfanumérico con T, BISA y UNION solo dígitos, QR BANECO (ID manual) solo dígitos. **BANECO** (C.I., depósito de fecha anterior que NO está en el reporte QR del día) es distinto de **QR BANECO**: solo dígitos; su código no puede coincidir con ningún ID del reporte QR del borrador ni estar ya usado (como respaldo QR o como C.I. BANECO) en otro cierre vigente, y a la inversa un ID QR no puede coincidir con una C.I. BANECO de otro cierre vigente. Las C.I. admiten una *Fecha del depósito* opcional (dd/mm/aaaa; columna `cierre_ci.fecha_deposito`, esquema v6); BANECO sin fecha, fecha inválida o posterior al cierre generan alerta no bloqueante.
- Base de datos: `app/db.py` migra el esquema v1 al arrancar de forma idempotente (columnas de revisión y tabla `cierre_ci`), sin perder datos.
- Iconos: `python scripts/generar_iconos.py` regenera favicon y apple-touch-icon desde `scripts/escudo_fuente.png` (reemplace ese archivo por la versión a color).

## v3: tres roles y administración
- Roles: `superadmin` (todo lo de cajero e ingresos, más revertir revisiones, administración de usuarios y auditoría), `ingresos` y `cajero`.
- Usuarios en la tabla `usuarios` (login sin distinguir mayúsculas; se guarda en minúsculas). `APP_USERS` solo siembra los que faltan, así los cambios hechos en la UI no se pisan al reiniciar. `python scripts/hash_clave.py --json usuario "Nombre" superadmin` genera entradas (agregue `"forzar": true` solo si quiere reaplicarla).
- Superadmin: «Revertir revisión» (CONFORME o RECHAZADO a PENDIENTE, motivo de 10+ caracteres, no si está REEMPLAZADO; el historial queda en `revisiones_historial` y en la auditoría), `/admin/usuarios` (crear, rol, activar/desactivar, restablecer clave; siempre queda al menos un superadmin activo) y `/admin/auditoria` (50 por página, filtros por usuario, acción y fecha).
- «Mi cuenta» (`/cuenta`): cualquier usuario cambia su clave (mínimo 8 caracteres).
- Base de datos: esquema 3 (tablas `usuarios` y `revisiones_historial`); la migración es idempotente y pasa las revisiones v2 al historial.

## v8: plantilla Excel de comunicaciones internas
- El cajero llena durante el día la plantilla (`/plantilla-comunicaciones-internas`, copia en `app/static/plantilla_comunicaciones_internas.xlsx`, regenerable con `python scripts/generar_plantilla.py`): N° factura | Banco | Código de asignación. La sube en «Nuevo cierre» (campo opcional) junto al movimiento y el reporte QR; importe, cliente y fecha salen del movimiento.
- Lectura tolerante (`app/plantilla.py`) y validación por fila (`Almacen.aplicar_plantilla`): las filas válidas se aplican como `reclasificaciones` (mismas reglas que el respaldo manual; el formato por banco sigue siendo alerta no bloqueante); las filas con error no se aplican y se listan en «Filas de la plantilla con problemas». «Reemplazar plantilla» rehace la importación; lo deshecho por el cajero queda en `plantilla_descartes` y no se reaplica.
- Se guarda como `comunicaciones_internas_original.xlsx` (con SHA-256 en `cierre.json`, columna `cierres.sha256_plantilla` y PDF). Esquema 5 (migración idempotente). «Editar y reenviar» la copia y permite reemplazarla.

## Copiar para el extracto (Ingresos y superadmin)
En el detalle del cierre, **Copiar para el extracto** copia en una sola pegada (hoja TRINIDAD del Excel, desde A2, TSV sin encabezado, siempre 10 columnas) las filas QR BANECO (facturas con ID de transacción: conciliadas, confirmadas o con respaldo manual QR) y las comunicaciones internas del cierre (`cierre_ci`: BNB, BMSC, BISA, UNION, BANECO). Columnas: A Acción | B Importe | C Código | D N° factura | E Correlativo | F Fecha cierre dd/mm/aaaa | G-I vacías | J Banco. Las filas QR llevan A vacía y J = `QR BANECO`; las C.I. llevan A = `REGISTRAR` y J = banco. Orden: primero QR y luego C.I., por factura. La fecha del depósito no se copia. Junto a los botones se muestra «N QR BANECO + M comunicaciones internas · total Bs X». Solo el superadmin ve **Copiar para liberar** (mismas filas con A = `LIBERAR`); ese TSV no se incrusta en la página para otros roles. Se procesa con el Office Script de `office_scripts/` (hoja REGISTRO con control de duplicados).
