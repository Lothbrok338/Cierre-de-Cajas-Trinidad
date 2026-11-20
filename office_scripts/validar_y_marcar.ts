// Office Script (Excel para la web): valida un lote de pagos QR contra el extracto bancario y marca/libera
// las filas usadas. Pegar tal cual en Automatizar > Nuevo script. Sin dependencias, sin export, sin fetch.
//
// Hoja del lote (HOJA_LOTE), fila 1 = encabezados:
//   A Acción | B Importe | C Código de asignación | D N° factura | E Correlativo cierre | F Fecha cierre (dd/mm/aaaa)
//   G Resultado | H Fila extracto | I Procesado el
// Se procesan SOLO las filas con G vacío y B, C y D con datos. TODO O NADA: si una fila del lote falla,
// no se escribe nada en el extracto.

const HOJA_EXTRACTO = "Hoja1";      // si no existe, se usa la primera hoja que no sea la del lote
const HOJA_LOTE = "TRINIDAD";
const PREFIJO_SEDE = "TRINIDAD";    // la marca en la columna I del extracto es "<PREFIJO_SEDE> - <MM>"

const TOLERANCIA_MONTO = 0.005;
const MAX_FILAS_BUSCAR_ENCABEZADO = 60;
// Columnas (índice base 0) del extracto
const COL_FECHA = 0, COL_NRO = 2, COL_MONTO = 5, COL_H = 7, COL_I = 8, COL_J = 9, ANCHO_EXTRACTO = 10;
// Columnas (índice base 0) de la hoja del lote
const L_ACCION = 0, L_IMPORTE = 1, L_CODIGO = 2, L_FACTURA = 3, L_CORRELATIVO = 4, L_FECHA = 5, L_RESULTADO = 6, L_FILA = 7, L_PROCESADO = 8, ANCHO_LOTE = 9;
const ENCABEZADOS_LOTE = ["Acción", "Importe", "Código de asignación", "N° factura", "Correlativo cierre", "Fecha cierre (dd/mm/aaaa)", "Resultado", "Fila extracto", "Procesado el"];

function main(workbook: ExcelScript.Workbook): string {
  // ---------- hoja del lote ----------
  let lote = workbook.getWorksheet(HOJA_LOTE);
  if (!lote) {
    lote = workbook.addWorksheet(HOJA_LOTE);
    prepararHojaLote(lote);
    const msg = "Hoja creada: pegue las filas y vuelva a ejecutar";
    lote.getRange("K1").setValue(msg);
    return msg;
  }

  // ---------- hoja del extracto ----------
  let extracto = workbook.getWorksheet(HOJA_EXTRACTO);
  if (!extracto) {
    const hojas = workbook.getWorksheets();
    for (let i = 0; i < hojas.length; i++) {
      if (hojas[i].getName() !== HOJA_LOTE) { extracto = hojas[i]; break; }
    }
  }
  if (!extracto) return terminar(lote, "ERROR: no se encontró la hoja del extracto. Escrito: no");

  const ultimaE = ultimaFila(extracto);
  const E = ultimaE > 0 ? extracto.getRangeByIndexes(0, 0, ultimaE, ANCHO_EXTRACTO).getValues() : [];
  let encabezado = -1;
  for (let i = 0; i < E.length && i < MAX_FILAS_BUSCAR_ENCABEZADO; i++) {
    if (norm(E[i][COL_FECHA]) === "fecha" && norm(E[i][COL_NRO]).indexOf("transac") >= 0) { encabezado = i; break; }
  }
  if (encabezado < 0) return terminar(lote, 'ERROR: no se encontró el encabezado del extracto (fila con "Fecha" en A y "Nro Transacción" en C). Escrito: no');

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
    if (ultimaE > 0) textosHJ = extracto.getRangeByIndexes(0, COL_H, ultimaE, 3).getText();
  } catch (e) { textosHJ = null; }
  const mostrar = (i: number, k: number): string => {
    if (textosHJ && textosHJ[i] && textosHJ[i][k] !== undefined) return String(textosHJ[i][k]).trim();
    return String(E[i][COL_H + k] === null || E[i][COL_H + k] === undefined ? "" : E[i][COL_H + k]).trim();
  };

  // ---------- leer el lote (una sola lectura) ----------
  const ultimaL = ultimaFila(lote);
  const L = ultimaL > 1 ? lote.getRangeByIndexes(0, 0, ultimaL, ANCHO_LOTE).getValues() : [];
  interface Fila { r: number; accion: string; importe: number; codigo: string; factura: string; fecha: string; msg: string; ok: boolean; filaExtracto: number; mes: number }
  const filas: Fila[] = [];
  for (let r = 1; r < L.length; r++) {
    const f = L[r];
    if (!vacio(f[L_RESULTADO])) continue;
    const tieneB = !vacio(f[L_IMPORTE]), tieneC = !vacio(f[L_CODIGO]), tieneD = !vacio(f[L_FACTURA]);
    if (!tieneB && !tieneC && !tieneD) continue;
    filas.push({
      r: r, accion: normAccion(f[L_ACCION]), importe: parseMonto(f[L_IMPORTE]), codigo: normCodigo(f[L_CODIGO]),
      factura: normCodigo(f[L_FACTURA]), fecha: "", msg: "", ok: false, filaExtracto: -1, mes: 0
    });
    // guardar si la fila está completa (mensaje de incompleta se resuelve más abajo)
    if (!(tieneB && tieneC && tieneD)) filas[filas.length - 1].msg = "✖ INCOMPLETO: falta importe, código o N° factura";
  }
  if (filas.length === 0) return terminar(lote, "Lote: 0 filas · ✔ 0 · ✖ 0 · Escrito: no (no hay filas nuevas con importe, código y N° factura)");

  const fechaLote = (r: number) => L[r][L_FECHA];
  const ahora = fechaHoraTexto(new Date());

  // ---------- acción del lote ----------
  const acciones: { [a: string]: boolean } = {};
  for (const f of filas) {
    if (f.accion === "MARCAR" || f.accion === "LIBERAR") acciones[f.accion] = true;
  }
  const mezcla = acciones["MARCAR"] === true && acciones["LIBERAR"] === true;
  const modoLiberar = acciones["LIBERAR"] === true && !mezcla;

  if (mezcla) {
    for (const f of filas) f.msg = "✖ LOTE MEZCLADO: no combine MARCAR y LIBERAR en un mismo lote";
    return cerrarLote(lote, extracto, filas, ahora, false, [], "Error general: el lote mezcla MARCAR y LIBERAR");
  }

  // repetidos dentro del lote
  const veces: { [c: string]: number } = {};
  for (const f of filas) veces[f.codigo] = (veces[f.codigo] || 0) + 1;

  // ---------- validación fila por fila ----------
  const escrituras: { fila: number; factura: string; marca: string }[] = [];
  for (const f of filas) {
    if (f.msg !== "") continue;
    if (f.accion !== "MARCAR" && f.accion !== "LIBERAR") { f.msg = "✖ ACCIÓN NO VÁLIDA (use MARCAR, LIBERAR o déjela vacía)"; continue; }
    if (veces[f.codigo] > 1) { f.msg = "✖ REPETIDO EN LA LISTA"; continue; }
    const ubic = indice[f.codigo];
    if (!ubic || ubic.length === 0) { f.msg = "✖ NO EXISTE EN EXTRACTO"; continue; }
    if (ubic.length > 1) { f.msg = "✖ DUPLICADO EN EXTRACTO (filas " + ubic.map((x) => x + 1).join(", ") + ")"; continue; }
    const i = ubic[0];
    f.filaExtracto = i + 1;

    // mes de la marca: de la fecha de cierre (F) o, si está vacía, de la fecha de la fila del extracto
    const rawF = fechaLote(f.r);
    let fecha = vacio(rawF) ? null : parseFecha(rawF);
    if (!vacio(rawF) && fecha === null) { f.msg = "✖ FECHA DE CIERRE NO VÁLIDA (use dd/mm/aaaa)"; continue; }
    if (fecha === null) fecha = parseFecha(E[i][COL_FECHA]);
    if (fecha === null) { f.msg = "✖ SIN FECHA: no se pudo determinar el mes (complete la fecha de cierre)"; continue; }
    f.mes = fecha.m;
    const marca = PREFIJO_SEDE + " - " + fecha.m;

    const h = normCodigo(E[i][COL_H]), iv = normMarca(E[i][COL_I]), jv = vacio(E[i][COL_J]);
    const hVacio = vacio(E[i][COL_H]), iVacio = vacio(E[i][COL_I]);
    const esMia = h === f.factura && iv === normMarca(marca);

    if (f.accion === "LIBERAR") {
      if (!esMia) { f.msg = "✖ NO SE PUEDE LIBERAR: la marca no corresponde a esta factura"; continue; }
      f.ok = true; f.msg = "✔ LIBERADO"; escrituras.push({ fila: i, factura: "", marca: "" });
      continue;
    }

    // MARCAR
    if (isNaN(f.importe)) { f.msg = "✖ IMPORTE NO VÁLIDO"; continue; }
    const montoExt = parseMonto(E[i][COL_MONTO]);
    if (isNaN(montoExt) || Math.abs(montoExt - f.importe) > TOLERANCIA_MONTO) {
      f.msg = "✖ MONTO DISTINTO (extracto " + (isNaN(montoExt) ? String(E[i][COL_MONTO]).trim() : montoExt.toFixed(2)) + ")"; continue;
    }
    const ocupada = !(hVacio && iVacio && jv);
    if (ocupada) {
      if (esMia) { f.ok = true; f.msg = "✔ YA MARCADO (misma factura)"; continue; }
      f.msg = "✖ OCUPADO: H=" + mostrar(i, 0) + " | I=" + mostrar(i, 1) + " | J=" + mostrar(i, 2); continue;
    }
    f.ok = true; f.msg = "✔ MARCADO"; escrituras.push({ fila: i, factura: f.factura, marca: marca });
  }

  const hayErrores = filas.some((f) => !f.ok);
  if (hayErrores) {
    const noHecho = modoLiberar ? "— NO LIBERADO: hay errores en el lote" : "— NO MARCADO: hay errores en el lote";
    for (const f of filas) if (f.ok && f.msg.indexOf("YA MARCADO") < 0) f.msg = f.msg + " " + noHecho;
    return cerrarLote(lote, extracto, filas, ahora, false, [], "");
  }
  return cerrarLote(lote, extracto, filas, ahora, true, escrituras, "");
}

// ---------- escritura de resultados ----------
function cerrarLote(lote: ExcelScript.Worksheet, extracto: ExcelScript.Worksheet,
  filas: { r: number; msg: string; ok: boolean; filaExtracto: number }[], ahora: string,
  escribir: boolean, escrituras: { fila: number; factura: string; marca: string }[], errorGeneral: string): string {
  // 1) extracto: solo si todo el lote es válido
  let escrito = false;
  if (escribir) {
    for (const w of escrituras) {
      const rango = extracto.getRangeByIndexes(w.fila, COL_H, 1, 2);
      if (w.marca === "") {
        rango.setValues([["", ""]]);                 // LIBERAR: limpiar H e I
      } else {
        rango.setNumberFormat([["@", "@"]]);         // texto: el N° de factura no debe convertirse en número/fecha
        rango.setValues([[w.factura, w.marca]]);
      }
      escrito = true;
    }
  }
  // 2) lote: resultado, fila del extracto y fecha-hora de proceso
  for (const f of filas) {
    lote.getRangeByIndexes(f.r, L_RESULTADO, 1, 3).setValues([[f.msg, f.filaExtracto > 0 ? f.filaExtracto : "", ahora]]);
  }
  const k = filas.filter((f) => f.ok).length;
  const resumen = "Lote: " + filas.length + " filas · ✔ " + k + " · ✖ " + (filas.length - k) + " · Escrito: " + (escrito ? "sí" : "no")
    + (errorGeneral ? " · " + errorGeneral : "");
  return terminar(lote, resumen);
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
  const enc = hoja.getRangeByIndexes(0, 0, 1, ANCHO_LOTE);
  enc.setNumberFormat([ENCABEZADOS_LOTE.map(() => "@")]);
  enc.setValues([ENCABEZADOS_LOTE]);
  const fmt = enc.getFormat();
  fmt.getFont().setBold(true);
  fmt.getFill().setColor("#D9E1F2");
  const anchos = [90, 90, 170, 110, 150, 170, 380, 110, 150];
  for (let c = 0; c < anchos.length; c++) hoja.getRangeByIndexes(0, c, 1, 1).getFormat().setColumnWidth(anchos[c]);
  hoja.getRange("K1").getFormat().getFont().setBold(true);
  hoja.getFreezePanes().freezeRows(1);
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
  return String(v === null || v === undefined ? "" : v).normalize("NFD").replace(/[̀-ͯ]/g, "")
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
  const p = function (n: number) { return (n < 10 ? "0" : "") + n; };
  return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear() + " " + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
}
