// Office Script (Excel para la web): valida un lote de pagos (QR del extracto y comunicaciones internas "C.I.")
// y marca/libera el extracto BANECO + lleva la hoja REGISTRO acumulada (control de duplicados).
// Pegar tal cual en Automatizar > Nuevo script. Sin dependencias, sin export, sin fetch.
//
// Hoja del lote (HOJA_LOTE), fila 1 = encabezados:
//   A Acción | B Importe | C Código de asignación | D N° factura | E Correlativo cierre | F Fecha cierre (dd/mm/aaaa)
//   G Resultado | H Fila extracto | I Procesado el | J Banco      (el resumen del lote queda en K1)
// Banco: vacío o "QR BANECO" = fila QR del extracto; BNB, BMSC, BISA, UNION (o UNIÓN), BANECO = comunicación interna (C.I.).
// Acciones: QR = MARCAR (o vacía) / LIBERAR; C.I. = REGISTRAR (o vacía) / LIBERAR. LIBERAR no se mezcla con las otras.
// Se procesan SOLO las filas con G vacío y B, C y D con datos. TODO O NADA: si una fila del lote falla,
// no se escribe nada (ni en el extracto ni en REGISTRO).
//
// Hoja REGISTRO (la crea el script): A Fecha cierre | B N° factura | C Importe | D Banco | E Código de asignación |
//   F Correlativo cierre | G Registrado el | H Estado (VIGENTE / LIBERADO) | I Duplicado (fórmula)

const HOJA_EXTRACTO = "Hoja1";      // si no existe, se usa la primera hoja que no sea la del lote ni REGISTRO
const HOJA_LOTE = "TRINIDAD";
const HOJA_REGISTRO = "REGISTRO";
const PREFIJO_SEDE = "TRINIDAD";    // la marca en la columna I del extracto es "<PREFIJO_SEDE> - <MM>"

const TOLERANCIA_MONTO = 0.005;
const MAX_FILAS_BUSCAR_ENCABEZADO = 60;
// Columnas (índice base 0) del extracto
const COL_FECHA = 0, COL_NRO = 2, COL_MONTO = 5, COL_H = 7, COL_I = 8, COL_J = 9, ANCHO_EXTRACTO = 10;
// Columnas (índice base 0) de la hoja del lote
const L_ACCION = 0, L_IMPORTE = 1, L_CODIGO = 2, L_FACTURA = 3, L_CORRELATIVO = 4, L_FECHA = 5, L_RESULTADO = 6, L_FILA = 7, L_PROCESADO = 8, L_BANCO = 9, ANCHO_LOTE = 10;
const ENCABEZADOS_LOTE = ["Acción", "Importe", "Código de asignación", "N° factura", "Correlativo cierre", "Fecha cierre (dd/mm/aaaa)", "Resultado", "Fila extracto", "Procesado el", "Banco"];
// Columnas (índice base 0) de REGISTRO
const R_FECHA = 0, R_FACTURA = 1, R_IMPORTE = 2, R_BANCO = 3, R_CODIGO = 4, R_CORREL = 5, R_REGISTRADO = 6, R_ESTADO = 7, R_DUP = 8, ANCHO_REGISTRO_DATOS = 8;
const ENCABEZADOS_REGISTRO = ["Fecha cierre", "N° factura", "Importe", "Banco", "Código de asignación", "Correlativo cierre", "Registrado el", "Estado", "Duplicado"];
const FILAS_FORMATO_DUP = 20000;    // el formato condicional de la columna I cubre I2:I20000
const BANCO_QR = "QR BANECO";
const BANCOS_CI = ["BNB", "BMSC", "BISA", "UNION", "BANECO"];

type Celda = string | number | boolean;

interface Fila {
  r: number; accion: string; importe: number; codigo: string; factura: string; correl: string;
  banco: string; tipo: string; msg: string; ok: boolean; filaExtracto: number
}
interface RegFila { fila: number; fecha: string; factura: string; banco: string; codigo: string; correl: string; estado: string }
interface NuevoReg { fecha: string; factura: string; importe: number; banco: string; codigo: string; correl: string }
interface Marca { fila: number; factura: string; marca: string }
interface Extracto { hoja: ExcelScript.Worksheet; E: Celda[][]; indice: { [codigo: string]: number[] }; textosHJ: string[][] | null }
interface Plan { escrituras: Marca[]; altas: NuevoReg[]; bajas: number[]; liberar: boolean; siguiente: number }

function main(workbook: ExcelScript.Workbook): string {
  // ---------- hojas del lote y de REGISTRO ----------
  let lote = workbook.getWorksheet(HOJA_LOTE);
  if (!lote) {
    lote = workbook.addWorksheet(HOJA_LOTE);
    prepararHojaLote(lote);
    if (!workbook.getWorksheet(HOJA_REGISTRO)) prepararRegistro(workbook.addWorksheet(HOJA_REGISTRO));
    const msg = "Hoja creada: pegue las filas y vuelva a ejecutar";
    lote.getRange("K1").setValue(msg);
    return msg;
  }
  asegurarEncabezadoBanco(lote);
  let registro = workbook.getWorksheet(HOJA_REGISTRO);
  if (!registro) {
    registro = workbook.addWorksheet(HOJA_REGISTRO);
    prepararRegistro(registro);
  }

  // ---------- leer el lote (una sola lectura) ----------
  const ultimaL = ultimaFila(lote);
  const L = ultimaL > 1 ? lote.getRangeByIndexes(0, 0, ultimaL, ANCHO_LOTE).getValues() : [];
  const filas: Fila[] = [];
  for (let r = 1; r < L.length; r++) {
    const f = L[r];
    if (!vacio(f[L_RESULTADO])) continue;
    const tieneB = !vacio(f[L_IMPORTE]), tieneC = !vacio(f[L_CODIGO]), tieneD = !vacio(f[L_FACTURA]);
    if (!tieneB && !tieneC && !tieneD) continue;
    const cb = clasificarBanco(f[L_BANCO]);
    const crudo = norm(f[L_ACCION]).toUpperCase();
    const defecto = cb.tipo === "QR" ? "MARCAR" : (cb.tipo === "CI" ? "REGISTRAR" : "");
    filas.push({
      r: r, accion: crudo === "" ? defecto : crudo, importe: parseMonto(f[L_IMPORTE]), codigo: normCodigo(f[L_CODIGO]),
      factura: normCodigo(f[L_FACTURA]), correl: normCodigo(f[L_CORRELATIVO]), banco: cb.nombre, tipo: cb.tipo,
      msg: "", ok: false, filaExtracto: -1
    });
    if (!(tieneB && tieneC && tieneD)) filas[filas.length - 1].msg = "✖ INCOMPLETO: falta importe, código o N° factura";
  }
  if (filas.length === 0) return terminar(lote, "Lote: 0 filas · ✔ 0 · ✖ 0 · Extracto: escrito no · Registro: +0 filas (no hay filas nuevas con importe, código y N° factura)");

  const ahora = fechaHoraTexto(new Date());

  // ---------- acción del lote: LIBERAR no se mezcla con MARCAR/REGISTRAR ----------
  let hayLiberar = false, hayAlta = false;
  for (const f of filas) {
    if (f.accion === "LIBERAR") hayLiberar = true;
    if (f.accion === "MARCAR" || f.accion === "REGISTRAR") hayAlta = true;
  }
  const mezcla = hayLiberar && hayAlta;
  const modoLiberar = hayLiberar && !mezcla;
  const plan: Plan = { escrituras: [], altas: [], bajas: [], liberar: modoLiberar, siguiente: 1 };

  if (mezcla) {
    for (const f of filas) f.msg = "✖ LOTE MEZCLADO: no combine LIBERAR con MARCAR/REGISTRAR en un mismo lote";
    return cerrarLote(lote, null, registro, filas, ahora, false, plan, "Error general: el lote mezcla LIBERAR con MARCAR/REGISTRAR");
  }

  // ---------- hoja del extracto (solo si hay filas QR) ----------
  let ex: Extracto | null = null;
  if (filas.some((f) => f.tipo === "QR")) {
    const res = cargarExtracto(workbook);
    if (typeof res === "string") return terminar(lote, res);
    ex = res;
  }

  // ---------- leer REGISTRO ----------
  const ultimaR = ultimaFila(registro);
  const R = ultimaR > 1 ? registro.getRangeByIndexes(0, 0, ultimaR, ANCHO_REGISTRO_DATOS).getValues() : [];
  const vig: RegFila[] = [];
  for (let i = 1; i < R.length; i++) {
    if (vacio(R[i][R_FACTURA]) && vacio(R[i][R_CODIGO])) continue;
    plan.siguiente = i + 1;
    const estado = norm(R[i][R_ESTADO]).toUpperCase();
    if (estado !== "VIGENTE") continue;
    vig.push({
      fila: i, fecha: String(R[i][R_FECHA]).trim(), factura: normCodigo(R[i][R_FACTURA]), banco: norm(R[i][R_BANCO]).toUpperCase(),
      codigo: normCodigo(R[i][R_CODIGO]), correl: String(R[i][R_CORREL]).trim(), estado: estado
    });
  }

  // repetidos dentro del lote (banco + código; y N° de factura en MARCAR/REGISTRAR)
  const vecesCodigo: { [c: string]: number } = {};
  const vecesFactura: { [c: string]: number } = {};
  for (const f of filas) {
    vecesCodigo[f.banco + "|" + f.codigo] = (vecesCodigo[f.banco + "|" + f.codigo] || 0) + 1;
    if (f.accion === "MARCAR" || f.accion === "REGISTRAR") vecesFactura[f.factura] = (vecesFactura[f.factura] || 0) + 1;
  }

  // ---------- validación fila por fila ----------
  for (const f of filas) {
    if (f.msg !== "") continue;
    if (f.tipo === "X") { f.msg = "✖ BANCO NO VÁLIDO"; continue; }
    if (f.tipo === "QR" && f.accion !== "MARCAR" && f.accion !== "LIBERAR") { f.msg = "✖ ACCIÓN NO VÁLIDA (use MARCAR, LIBERAR o déjela vacía)"; continue; }
    if (f.tipo === "CI" && f.accion !== "REGISTRAR" && f.accion !== "LIBERAR") { f.msg = "✖ ACCIÓN NO VÁLIDA para C.I. (use REGISTRAR, LIBERAR o déjela vacía)"; continue; }
    if (vecesCodigo[f.banco + "|" + f.codigo] > 1) { f.msg = "✖ REPETIDO EN LA LISTA"; continue; }
    if (f.accion !== "LIBERAR" && vecesFactura[f.factura] > 1) { f.msg = "✖ REPETIDO EN LA LISTA (N° factura)"; continue; }

    const rawF = L[f.r][L_FECHA];
    let fecha = vacio(rawF) ? null : parseFecha(rawF);
    if (!vacio(rawF) && fecha === null) { f.msg = "✖ FECHA DE CIERRE NO VÁLIDA (use dd/mm/aaaa)"; continue; }

    if (f.tipo === "CI") {
      // comunicación interna: no se busca en el extracto
      if (f.accion === "LIBERAR") {
        const coinciden = vig.filter((x) => x.banco === f.banco && x.codigo === f.codigo && x.factura === f.factura);
        if (coinciden.length === 0) { f.msg = "✖ NO SE PUEDE LIBERAR: no está registrada"; continue; }
        for (const x of coinciden) plan.bajas.push(x.fila);
        f.ok = true; f.msg = "✔ LIBERADO"; continue;
      }
      if (isNaN(f.importe)) { f.msg = "✖ IMPORTE NO VÁLIDO"; continue; }
      if (fecha === null) { f.msg = "✖ SIN FECHA: complete la fecha de cierre"; continue; }
      const ctl = controlRegistro(vig, f.banco, f.codigo, f.factura);
      if (ctl.msg !== "") { f.msg = ctl.msg; continue; }
      f.ok = true;
      if (ctl.existe) { f.msg = "✔ YA REGISTRADO"; continue; }
      f.msg = "✔ REGISTRADO";
      plan.altas.push({ fecha: fechaTexto(fecha), factura: f.factura, importe: f.importe, banco: f.banco, codigo: f.codigo, correl: f.correl });
      continue;
    }

    // ----- QR: contra el extracto -----
    if (ex === null) { f.msg = "✖ ERROR: extracto no disponible"; continue; }
    const E = ex.E;
    const ubic = ex.indice[f.codigo];
    if (!ubic || ubic.length === 0) { f.msg = "✖ NO EXISTE EN EXTRACTO"; continue; }
    if (ubic.length > 1) { f.msg = "✖ DUPLICADO EN EXTRACTO (filas " + ubic.map((x) => x + 1).join(", ") + ")"; continue; }
    const i = ubic[0];
    f.filaExtracto = i + 1;

    // mes de la marca: de la fecha de cierre (F) o, si está vacía, de la fecha de la fila del extracto
    if (fecha === null) fecha = parseFecha(E[i][COL_FECHA]);
    if (fecha === null) { f.msg = "✖ SIN FECHA: no se pudo determinar el mes (complete la fecha de cierre)"; continue; }
    const marca = PREFIJO_SEDE + " - " + fecha.m;

    const h = normCodigo(E[i][COL_H]), iv = normMarca(E[i][COL_I]), jv = vacio(E[i][COL_J]);
    const hVacio = vacio(E[i][COL_H]), iVacio = vacio(E[i][COL_I]);
    const esMia = h === f.factura && iv === normMarca(marca);

    if (f.accion === "LIBERAR") {
      if (!esMia) { f.msg = "✖ NO SE PUEDE LIBERAR: la marca no corresponde a esta factura"; continue; }
      const coinciden = vig.filter((x) => x.banco === f.banco && x.codigo === f.codigo && x.factura === f.factura);
      for (const x of coinciden) plan.bajas.push(x.fila);
      f.ok = true; f.msg = coinciden.length > 0 ? "✔ LIBERADO" : "✔ LIBERADO (sin fila en REGISTRO)";
      plan.escrituras.push({ fila: i, factura: "", marca: "" });
      continue;
    }

    // MARCAR
    if (isNaN(f.importe)) { f.msg = "✖ IMPORTE NO VÁLIDO"; continue; }
    const montoExt = parseMonto(E[i][COL_MONTO]);
    if (isNaN(montoExt) || Math.abs(montoExt - f.importe) > TOLERANCIA_MONTO) {
      f.msg = "✖ MONTO DISTINTO (extracto " + (isNaN(montoExt) ? String(E[i][COL_MONTO]).trim() : montoExt.toFixed(2)) + ")"; continue;
    }
    const ctl = controlRegistro(vig, f.banco, f.codigo, f.factura);
    if (ctl.msg !== "") { f.msg = ctl.msg; continue; }
    const nuevoReg: NuevoReg = { fecha: fechaTexto(fecha), factura: f.factura, importe: f.importe, banco: f.banco, codigo: f.codigo, correl: f.correl };
    const ocupada = !(hVacio && iVacio && jv);
    if (ocupada) {
      if (esMia) {
        f.ok = true; f.msg = "✔ YA MARCADO (misma factura)";
        if (!ctl.existe) plan.altas.push(nuevoReg);   // marca del extracto sin fila en REGISTRO: se agrega
        continue;
      }
      f.msg = "✖ OCUPADO: H=" + mostrarHJ(ex, i, 0) + " | I=" + mostrarHJ(ex, i, 1) + " | J=" + mostrarHJ(ex, i, 2); continue;
    }
    f.ok = true; f.msg = "✔ MARCADO";
    plan.escrituras.push({ fila: i, factura: f.factura, marca: marca });
    if (!ctl.existe) plan.altas.push(nuevoReg);
  }

  const hayErrores = filas.some((f) => !f.ok);
  if (hayErrores) {
    for (const f of filas) {
      if (!f.ok || f.msg.indexOf("YA MARCADO") >= 0 || f.msg.indexOf("YA REGISTRADO") >= 0) continue;
      const noHecho = modoLiberar ? "— NO LIBERADO: hay errores en el lote" : (f.tipo === "CI" ? "— NO REGISTRADO: hay errores en el lote" : "— NO MARCADO: hay errores en el lote");
      f.msg = f.msg + " " + noHecho;
    }
    return cerrarLote(lote, ex, registro, filas, ahora, false, plan, "");
  }
  return cerrarLote(lote, ex, registro, filas, ahora, true, plan, "");
}

// ---------- extracto ----------
function cargarExtracto(workbook: ExcelScript.Workbook): Extracto | string {
  let hoja = workbook.getWorksheet(HOJA_EXTRACTO);
  if (!hoja) {
    const hojas = workbook.getWorksheets();
    for (let i = 0; i < hojas.length; i++) {
      const n = hojas[i].getName();
      if (n !== HOJA_LOTE && n !== HOJA_REGISTRO) { hoja = hojas[i]; break; }
    }
  }
  if (!hoja) return "ERROR: no se encontró la hoja del extracto. Escrito: no";

  const ultimaE = ultimaFila(hoja);
  const E = ultimaE > 0 ? hoja.getRangeByIndexes(0, 0, ultimaE, ANCHO_EXTRACTO).getValues() : [];
  let encabezado = -1;
  for (let i = 0; i < E.length && i < MAX_FILAS_BUSCAR_ENCABEZADO; i++) {
    if (norm(E[i][COL_FECHA]) === "fecha" && norm(E[i][COL_NRO]).indexOf("transac") >= 0) { encabezado = i; break; }
  }
  if (encabezado < 0) return 'ERROR: no se encontró el encabezado del extracto (fila con "Fecha" en A y "Nro Transacción" en C). Escrito: no';

  // índice código -> filas del extracto (índices base 0)
  const indice: { [codigo: string]: number[] } = {};
  for (let i = encabezado + 1; i < E.length; i++) {
    const cod = normCodigo(E[i][COL_NRO]);
    if (cod === "") continue;
    if (!indice[cod]) indice[cod] = [];
    indice[cod].push(i);
  }
  // texto mostrado de H:J (para mensajes: una fecha real se ve como fecha, no como número de serie)
  let textosHJ: string[][] | null = null;
  try {
    if (ultimaE > 0) textosHJ = hoja.getRangeByIndexes(0, COL_H, ultimaE, 3).getText();
  } catch (e) { textosHJ = null; }
  return { hoja: hoja, E: E, indice: indice, textosHJ: textosHJ };
}

function mostrarHJ(ex: Extracto, i: number, k: number): string {
  if (ex.textosHJ && ex.textosHJ[i] && ex.textosHJ[i][k] !== undefined) return String(ex.textosHJ[i][k]).trim();
  const v = ex.E[i][COL_H + k];
  return String(v === null || v === undefined ? "" : v).trim();
}

// ---------- control de duplicados contra REGISTRO (solo filas VIGENTE) ----------
function cierreDe(x: RegFila): string {
  return x.correl !== "" ? x.correl : x.fecha;
}

function controlRegistro(vig: RegFila[], banco: string, codigo: string, factura: string): { msg: string; existe: boolean } {
  let existe = false;
  // (a) misma factura ya vigente con otro banco u otro código
  for (const x of vig) {
    if (x.factura !== factura) continue;
    if (x.banco === banco && x.codigo === codigo) { existe = true; continue; }
    return { msg: "✖ FACTURA YA REGISTRADA (banco " + x.banco + ", código " + x.codigo + ", cierre " + cierreDe(x) + ")", existe: false };
  }
  // (b) mismo banco + código ya vigente con otra factura
  for (const x of vig) {
    if (x.banco === banco && x.codigo === codigo && x.factura !== factura) {
      return { msg: "✖ DUPLICADO EN REGISTRO: el código " + codigo + " de " + banco + " ya se usó en la factura " + x.factura + " (cierre " + cierreDe(x) + ")", existe: false };
    }
  }
  // (c) misma factura + banco + código ya vigente: existe = true (idempotente)
  return { msg: "", existe: existe };
}

// ---------- escritura de resultados ----------
function cerrarLote(lote: ExcelScript.Worksheet, ex: Extracto | null, registro: ExcelScript.Worksheet,
  filas: Fila[], ahora: string, escribir: boolean, plan: Plan, errorGeneral: string): string {
  let escritoExtracto = false, altas = 0, bajas = 0;
  if (escribir) {
    // 1) extracto
    if (ex !== null) {
      for (const w of plan.escrituras) {
        const rango = ex.hoja.getRangeByIndexes(w.fila, COL_H, 1, 2);
        if (w.marca === "") {
          rango.setValues([["", ""]]);                 // LIBERAR: limpiar H e I
        } else {
          rango.setNumberFormat([["@", "@"]]);         // texto: el N° de factura no debe convertirse en número/fecha
          rango.setValues([[w.factura, w.marca]]);
        }
        escritoExtracto = true;
      }
    }
    // 2) REGISTRO: estado LIBERADO y filas nuevas (con la fórmula de duplicados en I)
    for (const fila of plan.bajas) {
      registro.getRangeByIndexes(fila, R_ESTADO, 1, 1).setValues([["LIBERADO"]]);
      bajas++;
    }
    if (plan.altas.length > 0) {
      const n = plan.altas.length;
      const bloque = registro.getRangeByIndexes(plan.siguiente, 0, n, ANCHO_REGISTRO_DATOS);
      bloque.setNumberFormat(plan.altas.map(() => ["@", "@", "0.00", "@", "@", "@", "@", "@"]));
      bloque.setValues(plan.altas.map((a) => [a.fecha, a.factura, a.importe, a.banco, a.codigo, a.correl, ahora, "VIGENTE"]));
      registro.getRangeByIndexes(plan.siguiente, R_DUP, n, 1).setFormulas(plan.altas.map((a, k) => [formulaDuplicado(plan.siguiente + k + 1)]));
      altas = n;
    }
  }
  // 3) lote: resultado, fila del extracto y fecha-hora de proceso
  for (const f of filas) {
    lote.getRangeByIndexes(f.r, L_RESULTADO, 1, 3).setValues([[f.msg, f.filaExtracto > 0 ? f.filaExtracto : "", ahora]]);
  }
  const k = filas.filter((f) => f.ok).length;
  const resumen = "Lote: " + filas.length + " filas · ✔ " + k + " · ✖ " + (filas.length - k)
    + " · Extracto: escrito " + (escritoExtracto ? "sí" : "no")
    + " · Registro: " + (plan.liberar ? bajas + " liberadas" : "+" + altas + " filas")
    + (errorGeneral ? " · " + errorGeneral : "");
  return terminar(lote, resumen);
}

// fórmula de la columna I de REGISTRO para la fila n (base 1)
function formulaDuplicado(n: number): string {
  return '=IF(H' + n + '="LIBERADO","",IF(OR(COUNTIFS($D:$D,D' + n + ',$E:$E,E' + n + ',$H:$H,"VIGENTE")>1,COUNTIFS($B:$B,B' + n + ',$H:$H,"VIGENTE")>1),"DUPLICADO",""))';
}

function terminar(lote: ExcelScript.Worksheet, resumen: string): string {
  lote.getRange("K1").setValue(resumen);
  return resumen;
}

function prepararHojaLote(hoja: ExcelScript.Worksheet) {
  // columnas de texto: evita que Excel convierta códigos, facturas y fechas al pegar
  hoja.getRange("A:A").setNumberFormat("@");
  hoja.getRange("C:F").setNumberFormat("@");
  hoja.getRange("G:H").setNumberFormat("@");
  hoja.getRange("J:J").setNumberFormat("@");
  const enc = hoja.getRangeByIndexes(0, 0, 1, ANCHO_LOTE);
  enc.setNumberFormat([ENCABEZADOS_LOTE.map(() => "@")]);
  enc.setValues([ENCABEZADOS_LOTE]);
  const fmt = enc.getFormat();
  fmt.getFont().setBold(true);
  fmt.getFill().setColor("#D9E1F2");
  const anchos = [90, 90, 170, 110, 150, 170, 380, 110, 150, 110];
  for (let c = 0; c < anchos.length; c++) hoja.getRangeByIndexes(0, c, 1, 1).getFormat().setColumnWidth(anchos[c]);
  hoja.getRange("K1").getFormat().getFont().setBold(true);
  hoja.getFreezePanes().freezeRows(1);
}

// hoja TRINIDAD de una versión anterior (sin columna Banco): escribe el encabezado en J1 sin borrar nada
function asegurarEncabezadoBanco(lote: ExcelScript.Worksheet) {
  const j1 = lote.getRange("J1");
  if (!vacio(j1.getValue())) return;
  lote.getRange("J:J").setNumberFormat("@");
  j1.setValue("Banco");
  const fmt = j1.getFormat();
  fmt.getFont().setBold(true);
  fmt.getFill().setColor("#D9E1F2");
  fmt.setColumnWidth(110);
}

function prepararRegistro(hoja: ExcelScript.Worksheet) {
  hoja.getRange("A:B").setNumberFormat("@");
  hoja.getRange("D:H").setNumberFormat("@");
  hoja.getRange("C:C").setNumberFormat("0.00");
  const enc = hoja.getRangeByIndexes(0, 0, 1, ENCABEZADOS_REGISTRO.length);
  enc.setNumberFormat([ENCABEZADOS_REGISTRO.map(() => "@")]);
  enc.setValues([ENCABEZADOS_REGISTRO]);
  const fmt = enc.getFormat();
  fmt.getFont().setBold(true);
  fmt.getFont().setColor("#FFFFFF");
  fmt.getFill().setColor("#870B42");
  const anchos = [110, 110, 100, 100, 170, 150, 150, 100, 110];
  for (let c = 0; c < anchos.length; c++) hoja.getRangeByIndexes(0, c, 1, 1).getFormat().setColumnWidth(anchos[c]);
  hoja.getFreezePanes().freezeRows(1);
  // formato condicional: celdas de I que digan DUPLICADO en rojo (comparación exacta; el encabezado no se ve afectado)
  try {
    const cf = hoja.getRange("I2:I" + FILAS_FORMATO_DUP).addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    cf.getCustom().getRule().setFormula('=$I2="DUPLICADO"');
    cf.getCustom().getFormat().getFill().setColor("#FFC7CE");
    cf.getCustom().getFormat().getFont().setColor("#9C0006");
    cf.getCustom().getFormat().getFont().setBold(true);
  } catch (e) { /* si el formato condicional no se pudo crear, el valor DUPLICADO igual se calcula */ }
}

// vacío o "QR BANECO" = QR del extracto; BNB/BMSC/BISA/UNION(UNIÓN)/BANECO = C.I.; otro texto = no válido
function clasificarBanco(v: string | number | boolean | null | undefined): { tipo: string; nombre: string } {
  const b = norm(v).toUpperCase();
  if (b === "" || b === BANCO_QR) return { tipo: "QR", nombre: BANCO_QR };
  if (BANCOS_CI.indexOf(b) >= 0) return { tipo: "CI", nombre: b };
  return { tipo: "X", nombre: b };
}

// dd/mm/aaaa (texto, sin ambigüedad de configuración regional)
function fechaTexto(f: { d: number; m: number; y: number }): string {
  const p = (n: number): string => (n < 10 ? "0" : "") + n;
  return p(f.d) + "/" + p(f.m) + "/" + f.y;
}

// ---------- utilidades ----------
function ultimaFila(hoja: ExcelScript.Worksheet): number {
  // cantidad de filas hasta la última con valores (0 si la hoja está vacía)
  try {
    const u = hoja.getUsedRange(true);
    if (!u) return 0;
    return u.getRowIndex() + u.getRowCount();
  } catch (e) { return 0; }
}

function vacio(v: string | number | boolean | null | undefined): boolean {
  return v === null || v === undefined || String(v).trim() === "";
}

function norm(v: string | number | boolean | null | undefined): string {
  return String(v === null || v === undefined ? "" : v).normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[°º.]/g, "").replace(/\s+/g, " ").trim();
}

function normAccion(v: string | number | boolean | null | undefined): string {
  const a = norm(v).toUpperCase();
  return a === "" ? "MARCAR" : a;
}

function normMarca(v: string | number | boolean | null | undefined): string {
  return norm(v).toUpperCase().replace(/\s*-\s*/g, " - ");
}

// código o factura como texto: trim y sin ".0" final (el extracto trae "1592.0"; números sueltos llegan como número)
function normCodigo(v: string | number | boolean | null | undefined): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isInteger(v) ? v.toFixed(0) : String(v);
  return String(v).trim().replace(/^(\d+)\.0+$/, "$1");
}

// "1592.0" | 1592 | "1,592.00" | "1592,50" -> número (NaN si no es un monto)
function parseMonto(v: string | number | boolean | null | undefined): number {
  if (typeof v === "number") return v;
  let t = String(v === null || v === undefined ? "" : v).replace(/\s|Bs\.?|BOB/gi, "");
  if (t === "") return NaN;
  const c = t.lastIndexOf(","), p = t.lastIndexOf(".");
  if (c >= 0 && p >= 0) t = c > p ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else if (c >= 0) t = /,\d{1,2}$/.test(t) ? t.replace(",", ".") : t.replace(/,/g, "");
  return /^\d+(\.\d+)?$/.test(t) ? parseFloat(t) : NaN;
}

const MESES: { [m: string]: number } = {
  ene: 1, jan: 1, feb: 2, mar: 3, abr: 4, apr: 4, may: 5, jun: 6, jul: 7, ago: 8, aug: 8, sep: 9, set: 9, oct: 10, nov: 11, dic: 12, dec: 12
};

// dd/mm/aaaa | dd/Mes/aaaa | aaaa-mm-dd | número de serie de Excel -> {d, m, y}
function parseFecha(v: string | number | boolean | null | undefined): { d: number; m: number; y: number } | null {
  if (typeof v === "number") {
    if (!(v > 1 && v < 100000)) return null;
    const dt = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return { d: dt.getUTCDate(), m: dt.getUTCMonth() + 1, y: dt.getUTCFullYear() };
  }
  const t = String(v === null || v === undefined ? "" : v).trim().toLowerCase();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return valida(+m[3], +m[2], +m[1]);
  m = t.match(/^(\d{1,2})[\/\-. ]([a-z]{3,})[\/\-. ](\d{2,4})/);
  if (m) { const mes = MESES[m[2].substring(0, 3)]; return mes ? valida(+m[1], mes, +m[3]) : null; }
  m = t.match(/^(\d{1,2})[\/\-. ](\d{1,2})[\/\-. ](\d{2,4})/);
  if (m) return valida(+m[1], +m[2], +m[3]);
  return null;
}

function valida(d: number, mes: number, y: number): { d: number; m: number; y: number } | null {
  if (mes < 1 || mes > 12 || d < 1 || d > 31) return null;
  return { d: d, m: mes, y: y < 100 ? 2000 + y : y };
}

function fechaHoraTexto(d: Date): string {
  const p = (n: number): string => (n < 10 ? "0" : "") + n;
  return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear() + " " + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
}
