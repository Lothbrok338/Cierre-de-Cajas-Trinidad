# Office Script: validar y marcar el extracto BANECO

`validar_y_marcar.ts` corre en **Excel para la web** (Automatizar > Nuevo script; pegar el archivo tal cual).
Valida un lote de pagos QR contra el extracto bancario y marca (o libera) las filas usadas.

## Uso
1. En la web de cierres, abra el cierre y pulse **Copiar para el extracto** (o **Copiar para liberar**, solo superadmin).
2. Pegue en la hoja `TRINIDAD`, celda A2 (sin encabezado). Si la hoja no existe, ejecute el script una vez: la crea.
3. Ejecute el script. Cada fila del lote (G vacío, con importe, código y N° factura) recibe su resultado en G, la fila del extracto en H y la hora de proceso en I. El resumen queda en K1.
4. Para reprocesar filas con error, corrija y limpie G de esas filas.

Columnas de la hoja del lote: A Acción (vacío/`MARCAR`/`LIBERAR`) · B Importe · C Código de asignación · D N° factura · E Correlativo · F Fecha cierre (dd/mm/aaaa) · G Resultado · H Fila extracto · I Procesado el.

Reglas: todo o nada por lote (si una fila falla no se escribe nada en el extracto); no se mezclan `MARCAR` y `LIBERAR`; una fila del extracto está libre si H, I y J están vacías; la marca es H = N° factura, I = `TRINIDAD - <mes>`.

## Constantes (arriba del script)
`HOJA_EXTRACTO` ("Hoja1"; si no existe usa la primera hoja que no sea la del lote), `HOJA_LOTE` ("TRINIDAD"), `PREFIJO_SEDE` ("TRINIDAD").

## Pruebas (sin Excel)
```
cd office_scripts
npm install
npm test
```
Transpilan el `.ts` con `typescript` y lo ejecutan en `vm` contra un fake de `ExcelScript.Workbook` (`tests/harness.js`) con datos ficticios.
Prueba opcional con la copia real del extracto (no está en el repo; requiere `python3` con `openpyxl`):
```
EXTRACTO_REAL=/ruta/extracto.xlsx REPORTE_REAL=/ruta/Reporte_pagos.xlsx npm test
```
Sin esas variables esa prueba se omite.
