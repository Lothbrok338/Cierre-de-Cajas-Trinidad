# Instructivo: Validar y marcar (QR BANECO y comunicaciones internas) + hoja REGISTRO

## Para qué sirve

1. **Pagos QR**: marca en el extracto de BANECO (Excel en la nube) los movimientos QR que respaldan un cierre de caja, pero solo si están libres. Si alguno del lote está ocupado o no coincide, no marca nada y le dice por qué en la misma hoja.
2. **Comunicaciones internas (C.I.)** de BNB, BMSC, BISA, UNIÓN o BANECO: no están en el extracto, por eso solo se **registran**.
3. **Control de duplicados**: todo lo marcado o registrado queda anotado en la hoja **REGISTRO**, que avisa si una factura o un código se usó dos veces.

## Qué escribe

**En el extracto (solo filas QR):**
- **Columna H** = N° de factura.
- **Columna I** = `TRINIDAD - <mes>` (por ejemplo, `TRINIDAD - 10` para octubre).
- Un movimiento está **libre** cuando H, I y J están vacías. Si cualquiera de las tres tiene algo, está ocupado.

**En la hoja REGISTRO (QR y C.I.)**, una fila por pago:

| Columna | Contenido |
|---|---|
| A Fecha cierre | Fecha del cierre (dd/mm/aaaa) |
| B N° factura | Factura del cierre |
| C Importe | Monto |
| D Banco | `QR BANECO`, `BNB`, `BMSC`, `BISA`, `UNION` o `BANECO` |
| E Código de asignación | Código del pago |
| F Correlativo cierre | Correlativo del cierre |
| G Registrado el | Fecha y hora en que el script lo anotó |
| H Estado | `VIGENTE` (en uso) o `LIBERADO` (ya no cuenta) |
| I Duplicado | Fórmula automática: vacía, o la palabra **DUPLICADO** |

**Cómo se ve un duplicado:** la celda de la columna I dice **DUPLICADO** y se pinta de **rojo**. Aparece cuando dos filas VIGENTE tienen el mismo banco y código, o la misma factura. Las filas LIBERADO nunca cuentan. Normalmente no verá ninguno, porque el script bloquea los duplicados antes de escribir; la columna I es una segunda revisión por si alguien edita a mano.

---

## Instalación (una sola vez, la hace quien administra el Excel)

1. Abra el Excel del extracto en **Excel para la web** (en el navegador).
2. Vaya a la pestaña **Automatizar** y pulse **Nuevo script** (también puede aparecer como "Script nuevo"). Si no ve la pestaña Automatizar, es probable que Office Scripts esté desactivado para su organización: pida a TI que lo active.
3. Borre el código de ejemplo que aparece. Abra el archivo `validar_y_marcar.ts` (se lo entrega el administrador), copie todo su contenido y péguelo en el editor.
4. Cambie el nombre del script a **Validar y marcar** (haga clic en el nombre que aparece arriba) y pulse **Guardar**.
5. Pulse **Ejecutar** una vez. El script crea las hojas **TRINIDAD** y **REGISTRO** con sus encabezados y muestra el mensaje "Hoja creada: pegue las filas y vuelva a ejecutar". Si TRINIDAD ya existía de antes, el script solo agrega el encabezado **Banco** en J1 y crea REGISTRO; no borra nada.
6. Agregue un botón en el libro: en el panel del script, pulse el menú **"..."** (Más opciones) y elija **"Agregar en el libro"**. Si no lo ve, busque la opción **"Agregar botón"**. El botón aparece en la hoja actual; puede moverlo a la hoja TRINIDAD.
7. Comparta el libro con quienes van a usar el botón. El botón en el libro se comparte con quienes editan el archivo.

**Nota para el administrador:** el script toma el extracto de la hoja `Hoja1`. Si esa hoja no existe, usa la primera hoja que no sea TRINIDAD ni REGISTRO. Guarde el Excel en un sitio SharePoint del área (recomendado), no en un OneDrive personal.

---

## Uso diario (Ingresos)

1. En la página **Cierre Diario de Caja**, abra el cierre pendiente y pulse **Copiar para el extracto**. Ese único botón copia las filas **QR BANECO** y también las **comunicaciones internas** del cierre con su banco (BNB, BMSC, BISA, UNION o BANECO); junto al botón se muestra el conteo, por ejemplo «5 QR BANECO + 2 comunicaciones internas · total Bs 3,500.00». Si el cierre no tiene comunicaciones internas, solo salen las filas QR.
2. En el Excel, vaya a la hoja **TRINIDAD**. Haga clic en la primera celda vacía de la **columna A**, debajo de lo último procesado (en un lote nuevo, la celda A2). Pegue con **Ctrl+V**.
   - Las columnas se completan solas: A Acción, B Importe, C Código de asignación, D N° factura, E Correlativo, F Fecha cierre; G, H e I quedan vacías (las llena el script) y **J Banco**.
   - **Columna A (Acción)**: para QR queda vacía o dice `MARCAR`; para comunicaciones internas dice `REGISTRAR`.
   - **Columna J (Banco)**: vacía o `QR BANECO` = pago QR del extracto. `BNB`, `BMSC`, `BISA`, `UNION` (o `UNIÓN`) o `BANECO` = comunicación interna.
   - En un mismo pegado pueden venir filas QR y C.I. juntas.
3. Pulse el botón **Validar y marcar**.
4. Revise el resultado:
   - **Columna G (Resultado)**: el mensaje de cada fila.
   - **Celda K1 (resumen)**: por ejemplo, `Lote: 5 filas · ✔ 5 · ✖ 0 · Extracto: escrito sí · Registro: +5 filas`.
5. Según el resultado:
   - **Todas las filas con ✔**: el extracto quedó marcado (si había filas QR) y REGISTRO actualizado. Vuelva a la página del cierre y pulse **Aprobar**.
   - **Alguna fila con ✖**: no se escribió nada, ni en el extracto ni en REGISTRO. Vuelva a la página del cierre y pulse **Rechazar**, copiando en la observación el motivo de la columna G.

Regla importante: el lote es **todo o nada**. Si una sola fila tiene error, no se marca ni se registra ninguna fila.

---

## Significado de los resultados (columna G)

| Mensaje | Qué significa | Qué hacer |
|---|---|---|
| ✔ MARCADO | El movimiento QR estaba libre y se marcó con la factura (y se anotó en REGISTRO). | Nada, si el lote está completo. |
| ✔ REGISTRADO | Comunicación interna anotada en REGISTRO. | Nada. |
| ✔ YA MARCADO (misma factura) | El movimiento QR ya estaba marcado con esta misma factura. No es un error. Si faltaba en REGISTRO, se agrega. | Nada. |
| ✔ YA REGISTRADO | La comunicación interna ya estaba en REGISTRO con la misma factura, banco y código. No es un error. | Nada. |
| ✔ LIBERADO | (Solo con acción LIBERAR) Se liberó la marca y/o el registro de esa factura. | Nada. |
| ✔ LIBERADO (sin fila en REGISTRO) | Se liberó el extracto, pero esa factura no estaba en REGISTRO (cierre anterior a REGISTRO). No es un error. | Nada. |
| ✖ OCUPADO: H=… \| I=… \| J=… | El movimiento ya tiene otra marca. El texto muestra qué hay en H, I y J. | Revisar el movimiento en el extracto. Si la marca es de otro cierre, no se puede usar ese QR. |
| ✖ FACTURA YA REGISTRADA (banco X, código Y, cierre Z) | Esa factura ya está VIGENTE en REGISTRO con otro banco u otro código. | Revisar el cierre Z; no registrar la misma factura dos veces. |
| ✖ DUPLICADO EN REGISTRO: el código C de B ya se usó en la factura F (cierre Z) | Ese código de ese banco ya se usó en otra factura. | Revisar la factura F y el cierre Z. Si fue un error, un superadministrador debe liberar la anterior. |
| ✖ BANCO NO VÁLIDO | La columna J tiene un banco desconocido. | Use vacío, `QR BANECO`, `BNB`, `BMSC`, `BISA`, `UNION` o `BANECO`. |
| ✖ NO EXISTE EN EXTRACTO | El código de asignación de un QR no aparece en el extracto. | Verificar el código y que el extracto esté actualizado. |
| ✖ DUPLICADO EN EXTRACTO (filas …) | El mismo código aparece más de una vez en el extracto. | Avisar a quien administra el extracto. No se marca. |
| ✖ MONTO DISTINTO (extracto …) | El importe del cierre no coincide con el del extracto. | Revisar el importe del cierre antes de rechazar. |
| ✖ REPETIDO EN LA LISTA | El mismo banco y código aparece dos veces en el lote pegado. | Quitar la fila repetida y reprocesar. |
| ✖ REPETIDO EN LA LISTA (N° factura) | La misma factura aparece dos veces en el lote. | Quitar la fila repetida y reprocesar. |
| ✖ INCOMPLETO: falta importe, código o N° factura | A la fila le falta algún dato obligatorio. | Completar el dato y reprocesar. |
| ✖ FECHA DE CIERRE NO VÁLIDA (use dd/mm/aaaa) | La fecha de cierre no se pudo leer. | Corregir la columna F con formato dd/mm/aaaa y reprocesar. |
| ✖ SIN FECHA: … | En QR: no hay fecha de cierre ni fecha legible en el extracto. En C.I.: falta la fecha de cierre. | Completar la columna F y reprocesar. |
| ✖ IMPORTE NO VÁLIDO | El importe no es un número. | Corregir la columna B y reprocesar. |
| ✖ ACCIÓN NO VÁLIDA … | La columna A no corresponde: QR usa MARCAR o LIBERAR (o vacía); C.I. usa REGISTRAR o LIBERAR (o vacía). | Corregir o borrar el texto de la columna A. |
| ✖ LOTE MEZCLADO | El lote mezcla LIBERAR con MARCAR/REGISTRAR. | Separar en dos pegados (ver "Liberar"). |
| ✖ NO SE PUEDE LIBERAR: la marca no corresponde a esta factura | (QR) La marca del extracto es de otra factura o de otro mes. | No liberar. Avisar al superadministrador. |
| ✖ NO SE PUEDE LIBERAR: no está registrada | (C.I.) No hay una fila VIGENTE en REGISTRO con ese banco, código y factura. | No liberar. Avisar al superadministrador. |
| — NO MARCADO / NO REGISTRADO / NO LIBERADO: hay errores en el lote | Aparece al final del mensaje de las filas que estaban bien. No se escribió ninguna fila. | Corregir las filas con ✖ y reprocesar. |

En la celda K1 también puede aparecer:

- `Lote: 0 filas …`: no hay filas nuevas para procesar (columna G vacía con importe, código y factura).
- `ERROR: no se encontró la hoja del extracto` o `ERROR: no se encontró el encabezado del extracto`: el extracto no tiene la forma esperada. Avisar a TI o al administrador. No se escribe nada. (Un lote solo de comunicaciones internas no necesita el extracto.)

---

## Reprocesar un lote con errores

1. Corrija el dato que causó el error (importe, código, factura, fecha, banco o acción).
2. Borre el contenido de la **columna G** de esas filas. Si prefiere, borre las filas completas.
3. Pulse **Validar y marcar** de nuevo.

El script no vuelve a procesar las filas que ya tienen algo en G. Por eso, si una fila ya dice ✔, no la borre. Volver a procesar algo ya hecho es seguro: dice "YA MARCADO" o "YA REGISTRADO" y no duplica filas en REGISTRO.

---

## Liberar (solo superadministrador, cuando revierte una aprobación)

1. En la página del cierre, pulse **Copiar para liberar**.
2. En la hoja TRINIDAD, pegue el lote. La columna A dirá **LIBERAR**.
3. Pulse **Validar y marcar**.

Qué pasa:
- **QR**: el script limpia H e I del extracto solo si la marca corresponde a esa misma factura, y además cambia a **LIBERADO** la fila VIGENTE de REGISTRO. Si esa factura no estaba en REGISTRO (cierre anterior), igual libera y lo avisa con "(sin fila en REGISTRO)".
- **C.I.**: cambia a **LIBERADO** la fila de REGISTRO. Si no está registrada, responde "NO SE PUEDE LIBERAR: no está registrada".
- Una fila LIBERADO queda como historia y ya no cuenta para duplicados: el código y la factura pueden volver a usarse.

No mezcle LIBERAR con MARCAR/REGISTRAR en un mismo pegado: el script lo rechaza con "LOTE MEZCLADO". Pegue primero un lote y después el otro.

---

## Buenas prácticas

- No escriba a mano en las columnas H, I y J de las filas QR de Trinidad. Esas columnas las escribe solo el script.
- **No edite REGISTRO a mano**, salvo para corregir un dato mal anotado (por ejemplo, un importe). No borre filas ni cambie el Estado a mano: para liberar use la acción LIBERAR.
- **No ordene ni inserte filas** en el extracto ni en REGISTRO mientras el script se está ejecutando (ni filtre REGISTRO con filas ocultas que luego se pierdan).
- Guarde el Excel en un sitio SharePoint del área, no en un OneDrive personal.
- Si el resultado no coincide con lo que espera, no corrija el extracto a mano. Consulte antes con quien administra el Excel.
