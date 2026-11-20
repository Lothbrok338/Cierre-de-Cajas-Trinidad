# Instructivo: Validar y marcar QR BANECO en el extracto

## Para qué sirve

Marca en el extracto de BANECO (Excel en la nube) los movimientos QR que respaldan un cierre de caja, pero solo si están libres. Si alguno de los movimientos del lote está ocupado o no coincide, no marca nada y le dice por qué en la misma hoja.

## Qué escribe en el extracto

- **Columna H** = N° de factura.
- **Columna I** = `TRINIDAD - <mes>` (por ejemplo, `TRINIDAD - 10` para octubre).
- Un movimiento está **libre** cuando H, I y J están vacías. Si cualquiera de las tres tiene algo, está ocupado.

---

## Instalación (una sola vez, la hace quien administra el Excel)

1. Abra el Excel del extracto en **Excel para la web** (en el navegador).
2. Vaya a la pestaña **Automatizar** y pulse **Nuevo script** (también puede aparecer como "Script nuevo"). Si no ve la pestaña Automatizar, es probable que Office Scripts esté desactivado para su organización: pida a TI que lo active.
3. Borre el código de ejemplo que aparece. Abra el archivo `validar_y_marcar.ts` (se lo entrega el administrador), copie todo su contenido y péguelo en el editor.
4. Cambie el nombre del script a **Validar y marcar** (haga clic en el nombre que aparece arriba) y pulse **Guardar**.
5. Pulse **Ejecutar** una vez. El script crea la hoja **TRINIDAD** con sus encabezados y muestra el mensaje "Hoja creada: pegue las filas y vuelva a ejecutar".
6. Agregue un botón en el libro: en el panel del script, pulse el menú **"..."** (Más opciones) y elija **"Agregar en el libro"**. Si no lo ve, busque la opción **"Agregar botón"**. El botón aparece en la hoja actual; puede moverlo a la hoja TRINIDAD.
7. Comparta el libro con quienes van a usar el botón. El botón en el libro se comparte con quienes editan el archivo.

**Nota para el administrador:** el script toma el extracto de la hoja `Hoja1`. Si esa hoja no existe, usa la primera hoja que no sea TRINIDAD. Guarde el Excel en un sitio SharePoint del área (recomendado), no en un OneDrive personal.

---

## Uso diario (Ingresos)

1. En la página **Cierre Diario de Caja**, abra el cierre pendiente y pulse **Copiar para el extracto**.
2. En el Excel, vaya a la hoja **TRINIDAD**. Haga clic en la primera celda vacía de la **columna A**, debajo de lo último procesado (en un lote nuevo, la celda A2). Pegue con **Ctrl+V**.
   - Las columnas se completan solas: A Acción, B Importe, C Código de asignación, D N° factura, E Correlativo, F Fecha cierre.
   - La columna **A (Acción)** puede quedar vacía: vacía significa MARCAR.
3. Pulse el botón **Validar y marcar**.
4. Revise el resultado:
   - **Columna G (Resultado)**: el mensaje de cada fila.
   - **Celda K1 (resumen)**: por ejemplo, `Lote: 5 filas · ✔ 5 · ✖ 0 · Escrito: sí`.
5. Según el resultado:
   - **Todas las filas con ✔ y "Escrito: sí"**: los movimientos quedaron marcados. Vuelva a la página del cierre y pulse **Aprobar**.
   - **Alguna fila con ✖ o "Escrito: no"**: no se marcó nada en el extracto. Vuelva a la página del cierre y pulse **Rechazar**, copiando en la observación el motivo de la columna G.

Regla importante: el lote es **todo o nada**. Si una sola fila tiene error, ninguna fila se marca en el extracto.

---

## Significado de los resultados (columna G)

| Mensaje | Qué significa | Qué hacer |
|---|---|---|
| ✔ MARCADO | El movimiento estaba libre y se marcó con la factura. | Nada, si el lote está completo. |
| ✔ YA MARCADO (misma factura) | El movimiento ya estaba marcado con esta misma factura. No es un error. | Nada. |
| ✔ LIBERADO | (Solo con acción LIBERAR) Se liberó la marca de esa factura. | Nada. |
| ✖ OCUPADO: H=… \| I=… \| J=… | El movimiento ya tiene otra marca. El texto muestra qué hay en H, I y J. | Revisar el movimiento en el extracto. Si la marca es de otro cierre, no se puede usar ese QR. |
| ✖ NO EXISTE EN EXTRACTO | El código de asignación no aparece en el extracto. | Verificar el código y que el extracto esté actualizado. |
| ✖ DUPLICADO EN EXTRACTO (filas …) | El mismo código aparece más de una vez en el extracto. | Avisar a quien administra el extracto. No se marca. |
| ✖ MONTO DISTINTO (extracto …) | El importe del cierre no coincide con el del extracto. El texto muestra el monto del extracto. | Revisar el importe del cierre antes de rechazar. |
| ✖ REPETIDO EN LA LISTA | El mismo código aparece dos veces en el lote pegado. | Quitar la fila repetida y reprocesar. |
| ✖ INCOMPLETO: falta importe, código o N° factura | A la fila le falta algún dato obligatorio. | Completar el dato en el lote y reprocesar. |
| ✖ FECHA DE CIERRE NO VÁLIDA (use dd/mm/aaaa) | La fecha de cierre no se pudo leer. | Corregir la columna F con formato dd/mm/aaaa y reprocesar. |
| ✖ SIN FECHA: no se pudo determinar el mes | No hay fecha de cierre y tampoco se pudo leer la fecha del extracto. | Completar la fecha de cierre y reprocesar. |
| ✖ IMPORTE NO VÁLIDO | El importe no es un número. | Corregir la columna B y reprocesar. |
| ✖ ACCIÓN NO VÁLIDA (use MARCAR, LIBERAR o déjela vacía) | La columna A tiene un texto distinto de MARCAR o LIBERAR. | Corregir o borrar el texto de la columna A. |
| ✖ LOTE MEZCLADO | El lote tiene filas MARCAR y LIBERAR juntas. | Separar en dos pegados (ver "Liberar marcas"). |
| ✖ NO SE PUEDE LIBERAR: la marca no corresponde a esta factura | La marca del extracto es de otra factura o de otro mes. | No liberar. Avisar al superadministrador. |
| — NO MARCADO: hay errores en el lote | Aparece al final del mensaje de las filas que estaban bien. Ninguna fila se escribió. | Corregir las filas con ✖ y reprocesar. |
| — NO LIBERADO: hay errores en el lote | Igual que el anterior, en un lote de liberación. | Corregir las filas con ✖ y reprocesar. |

En la celda K1 también puede aparecer:

- `Lote: 0 filas …`: no hay filas nuevas para procesar (columna G vacía con importe, código y factura).
- `ERROR: no se encontró la hoja del extracto` o `ERROR: no se encontró el encabezado del extracto`: el extracto no tiene la forma esperada. Avisar a TI o al administrador del Excel. No se escribe nada.

Si no ve estos mensajes en G, revise que la hoja se llame TRINIDAD y que el botón haya terminado de ejecutarse. Si el mensaje sigue sin aparecer, avise al administrador.

---

## Reprocesar un lote con errores

1. Corrija el dato que causó el error (importe, código, factura, fecha o acción).
2. Borre el contenido de la **columna G** de esas filas. Si prefiere, borre las filas completas.
3. Pulse **Validar y marcar** de nuevo.

El script no vuelve a procesar las filas que ya tienen algo en G. Por eso, si una fila ya dice ✔ MARCADO, no la borre.

---

## Liberar marcas (solo superadministrador, cuando revierte una aprobación)

1. En la página del cierre, pulse **Copiar para liberar**.
2. En la hoja TRINIDAD, pegue el lote. La columna A dirá **LIBERAR**.
3. Pulse **Validar y marcar**.

El script libera solo si la marca del extracto corresponde a esa misma factura. Si no coincide, responde "NO SE PUEDE LIBERAR" y no toca el extracto.

No mezcle filas MARCAR y LIBERAR en un mismo pegado: el script lo rechaza con "LOTE MEZCLADO". Pegue primero un lote y después el otro.

---

## Buenas prácticas

- No escriba a mano en las columnas H, I y J de las filas QR de Trinidad. Esas columnas las escribe solo el script.
- No ordene ni inserte filas en la hoja del extracto mientras el script se está ejecutando.
- Guarde el Excel en un sitio SharePoint del área, no en un OneDrive personal.
- Si el resultado no coincide con lo que espera, no corrija el extracto a mano. Consulte antes con quien administra el Excel.
