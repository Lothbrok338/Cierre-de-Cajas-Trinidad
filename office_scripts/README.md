# Office Script: validar y marcar el extracto BANECO + REGISTRO

`validar_y_marcar.ts` corre en **Excel para la web** (Automatizar > Nuevo script; pegar el archivo tal cual).
Valida un lote de pagos (QR del extracto y comunicaciones internas, C.I.), marca/libera el extracto y lleva la hoja `REGISTRO` acumulada para detectar duplicados.

## Uso
1. En la web de cierres, abra el cierre y pulse **Copiar para el extracto** (o **Copiar para liberar**, solo superadmin).
2. Pegue en la hoja `TRINIDAD`, celda A2 (sin encabezado). Si la hoja no existe, ejecute el script una vez: la crea (y crea `REGISTRO`).
3. Ejecute el script. Cada fila del lote (G vacío, con importe, código y N° factura) recibe su resultado en G, la fila del extracto en H y la hora de proceso en I. El resumen queda en K1: `Lote: n filas · ✔ k · ✖ m · Extracto: escrito sí/no · Registro: +r filas` (o `r liberadas`).
4. Para reprocesar filas con error, corrija y limpie G de esas filas.

## Contrato de pegado (hoja TRINIDAD)
A Acción · B Importe · C Código de asignación · D N° factura · E Correlativo cierre · F Fecha cierre (dd/mm/aaaa) · G Resultado · H Fila extracto · I Procesado el · **J Banco** · K1 resumen.
Desde la web se pegan A..F, tres celdas vacías (G, H, I) y J (TSV).

| Tipo | A Acción | J Banco |
|---|---|---|
| QR (extracto) | vacío o `MARCAR` / `LIBERAR` | vacío o `QR BANECO` |
| C.I. | `REGISTRAR` (o vacío) / `LIBERAR` | `BNB`, `BMSC`, `BISA`, `UNION` (o `UNIÓN`), `BANECO` |

Un banco desconocido da `✖ BANCO NO VÁLIDO`. Las hojas TRINIDAD antiguas (sin J) siguen funcionando como QR; el script escribe `Banco` en J1.

## Reglas
- Todo o nada por lote: si una fila falla no se escribe nada (ni extracto ni REGISTRO).
- `LIBERAR` no se mezcla con `MARCAR`/`REGISTRAR` (sí pueden convivir QR y C.I. en un lote de alta).
- QR: una fila del extracto está libre si H, I y J están vacías; la marca es H = N° factura, I = `TRINIDAD - <mes>`. Además se agrega a REGISTRO.
- C.I.: no se busca en el extracto; exige importe numérico, código, N° factura y fecha de cierre; se agrega a REGISTRO.
- Duplicados (solo filas VIGENTE de REGISTRO): misma factura con otro banco/código -> `FACTURA YA REGISTRADA`; mismo banco+código con otra factura -> `DUPLICADO EN REGISTRO`; igual factura+banco+código -> `YA REGISTRADO` / `YA MARCADO` (idempotente; si falta la fila en REGISTRO se agrega). En el lote: se rechaza repetir banco+código o factura.
- LIBERAR: QR limpia H/I (si la marca es de esa factura) y pone `LIBERADO` en REGISTRO (si no hay fila: `(sin fila en REGISTRO)`, no es error); C.I. pone `LIBERADO` o falla con `no está registrada`.

## Hoja REGISTRO (la crea el script)
A Fecha cierre · B N° factura · C Importe · D Banco · E Código de asignación · F Correlativo cierre · G Registrado el · H Estado (`VIGENTE`/`LIBERADO`) · I Duplicado (fórmula por fila: `COUNTIFS`/`IF`/`OR`; formato condicional rojo si dice `DUPLICADO`). Encabezado guinda `#870B42`, fila 1 congelada, B/E/F como texto.

## Constantes (arriba del script)
`HOJA_EXTRACTO` ("Hoja1"; si no existe usa la primera hoja que no sea la del lote ni REGISTRO), `HOJA_LOTE` ("TRINIDAD"), `HOJA_REGISTRO` ("REGISTRO"), `PREFIJO_SEDE` ("TRINIDAD").

## Pruebas (sin Excel)
```
cd office_scripts
npm install
npm test
```
Transpilan el `.ts` con `typescript` y lo ejecutan en `vm` contra un fake de `ExcelScript.Workbook` (`tests/harness.js`, con fórmulas y formatos condicionales simulados) con datos ficticios. `tests/registro.test.js` cubre REGISTRO y C.I.
Prueba opcional con la copia real del extracto (no está en el repo; requiere `python3` con `openpyxl`):
```
EXTRACTO_REAL=/ruta/extracto.xlsx REPORTE_REAL=/ruta/Reporte_pagos.xlsx npm test
```
Sin esas variables esa prueba se omite. Reglas de Office Scripts respetadas: solo funciones flecha en callbacks de arreglos, tipos a nivel superior, sin `any`, sin export/imports.
