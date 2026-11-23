'use strict';
// Harness sin Excel: transpila validar_y_marcar.ts con `typescript` y lo ejecuta en `vm` contra un FAKE de ExcelScript.Workbook.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ts = require('typescript');

const RUTA_TS = path.join(__dirname, '..', 'validar_y_marcar.ts');

function cargarMain() {
  const src = fs.readFileSync(RUTA_TS, 'utf8');
  const out = ts.transpileModule(src, {
    reportDiagnostics: true,
    compilerOptions: { target: ts.ScriptTarget.ES2019, module: ts.ModuleKind.None },
  });
  const errores = (out.diagnostics || []).filter((d) => d.category === ts.DiagnosticCategory.Error);
  if (errores.length) throw new Error('TS: ' + errores.map((d) => ts.flattenDiagnosticMessageText(d.messageText, '\n')).join('; '));
  if (/^\s*export\s/m.test(src)) throw new Error('El script no debe usar export');
  const ctx = vm.createContext({ console, ExcelScript: { ConditionalFormatType: { custom: 'Custom' } } });
  vm.runInContext(out.outputText, ctx);
  return ctx.main;
}

// ---- Celdas: valor primitivo, o {v: serial, t: 'texto mostrado'} para una fecha real de Excel ----
const val = (c) => (c !== null && typeof c === 'object' && 'v' in c ? c.v : c === undefined ? '' : c);
const txt = (c) => (c !== null && typeof c === 'object' && 't' in c ? c.t : c === null || c === undefined ? '' : String(c));

function colLetra(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function colIndice(l) { let n = 0; for (const ch of l) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }

class Hoja {
  constructor(nombre, filas) {
    this.nombre = nombre; this.celdas = new Map(); this.formatos = new Map(); this.escrituras = []; this.formulas = new Map(); this.condFormats = [];
    (filas || []).forEach((f, r) => f.forEach((c, k) => { if (c !== null && c !== undefined && c !== '') this.celdas.set(r + ',' + k, c); }));
  }
  get(r, c) { const x = this.celdas.get(r + ',' + c); return x === undefined ? '' : x; }
  set(r, c, v) { if (v === '' || v === null || v === undefined) this.celdas.delete(r + ',' + c); else this.celdas.set(r + ',' + c, v); }
  valorTexto(r, c) { return txt(this.get(r, c)); }
  getName() { return this.nombre; }
  _limites() {
    let maxR = -1, maxC = -1;
    for (const k of this.celdas.keys()) { const [r, c] = k.split(',').map(Number); if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
    return { maxR, maxC };
  }
  getUsedRange() {
    const { maxR, maxC } = this._limites();
    if (maxR < 0) return this.getRangeByIndexes(0, 0, 1, 1);
    return this.getRangeByIndexes(0, 0, maxR + 1, maxC + 1);
  }
  getRangeByIndexes(r, c, nr, nc) { return new Rango(this, r, c, nr, nc); }
  getCell(r, c) { return new Rango(this, r, c, 1, 1); }
  getRange(dir) {
    let m = dir.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (m) { const c1 = colIndice(m[1]), r1 = +m[2] - 1, c2 = m[3] ? colIndice(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1; return new Rango(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1); }
    m = dir.match(/^([A-Z]+):([A-Z]+)$/);
    if (m) { const c1 = colIndice(m[1]), c2 = colIndice(m[2]); return new Rango(this, 0, c1, 1048576, c2 - c1 + 1, true); }
    throw new Error('Dirección no soportada por el fake: ' + dir);
  }
  getFreezePanes() { return noop(); }
}

function noop() {
  const f = function () { return p; };
  const p = new Proxy(f, { get: (t, k) => (k === 'then' ? undefined : p), apply: () => p });
  return p;
}

class Rango {
  constructor(h, r, c, nr, nc, columnas) { this.h = h; this.r = r; this.c = c; this.nr = nr; this.nc = nc; this.columnas = !!columnas; }
  getRowIndex() { return this.r; }
  getColumnIndex() { return this.c; }
  getRowCount() { return this.nr; }
  getColumnCount() { return this.nc; }
  _mat(fn) { const o = []; for (let i = 0; i < this.nr; i++) { const f = []; for (let j = 0; j < this.nc; j++) f.push(fn(this.r + i, this.c + j)); o.push(f); } return o; }
  getValues() { return this._mat((r, c) => val(this.h.get(r, c))); }
  getText() { return this._mat((r, c) => txt(this.h.get(r, c))); }
  getValue() { return val(this.h.get(this.r, this.c)); }
  getFormulas() { return this._mat((r, c) => this.h.formulas.get(r + ',' + c) || ''); }
  setFormulas(m) {
    if (m.length !== this.nr || m[0].length !== this.nc) throw new Error(`setFormulas: dimensiones ${m.length}x${m[0].length} vs rango ${this.nr}x${this.nc}`);
    m.forEach((f, i) => f.forEach((v, j) => this.h.formulas.set((this.r + i) + ',' + (this.c + j), v)));
  }
  addConditionalFormat(tipo) {
    const reg = { tipo, filaDesde: this.r, col: this.c, filas: this.nr, formula: null };
    this.h.condFormats.push(reg);
    return { getCustom: () => ({ getRule: () => ({ setFormula: (f) => { reg.formula = f; } }), getFormat: () => noop() }) };
  }
  setValues(m) {
    if (m.length !== this.nr || m[0].length !== this.nc) throw new Error(`setValues: dimensiones ${m.length}x${m[0].length} vs rango ${this.nr}x${this.nc}`);
    m.forEach((f, i) => f.forEach((v, j) => { this.h.set(this.r + i, this.c + j, v); this.h.escrituras.push({ r: this.r + i, c: this.c + j, v }); }));
  }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.h.set(this.r + i, this.c + j, v); }
  setNumberFormat(f) {
    if (this.columnas) { for (let j = 0; j < this.nc; j++) this.h.formatos.set('col' + (this.c + j), f); return; }
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.h.formatos.set((this.r + i) + ',' + (this.c + j), Array.isArray(f) ? f[i][j] : f);
  }
  getFormat() { return noop(); }
}

class Libro {
  constructor(hojas) { this.hojas = hojas; }
  getWorksheet(n) { return this.hojas.find((h) => h.nombre === n); }
  getWorksheets() { return this.hojas; }
  addWorksheet(n) { const h = new Hoja(n, []); this.hojas.push(h); return h; }
}

// ---- Constructores de datos ficticios con el formato real del extracto ----
const ENC_EXTRACTO = ['Fecha', 'Hora', 'Nro Transacción', 'Transacción', 'Nota', 'Monto', 'Saldo', 'COD DE TRANF', null, null];

// fila: {nro, monto, fecha='16/Oct/2026', H, I, J}  (todo como TEXTO, como en el extracto real)
function hojaExtracto(filas, { encabezadoEnFila = 4, nombre = 'Hoja1' } = {}) {
  const m = [];
  for (let i = 1; i < encabezadoEnFila; i++) m.push([]);
  m[0] = ['ECONOMICO CTA CTE '];
  if (encabezadoEnFila >= 3) m[encabezadoEnFila - 2] = [null, null, 'cod. Bancarizacion SAP'];
  m.push(ENC_EXTRACTO.slice());
  filas.forEach((f, i) => m.push([f.fecha || '16/Oct/2026', '10:0' + (i % 10) + ':00', f.nro, 'CREDITO ACH QR', 'PERSONA FICTICIA ' + i,
    typeof f.monto === 'number' ? f.monto.toFixed(1) : f.monto, '1000.00', f.H === undefined ? null : f.H, f.I === undefined ? null : f.I, f.J === undefined ? null : f.J]));
  return new Hoja(nombre, m);
}

const ENC_LOTE = ['Acción', 'Importe', 'Código de asignación', 'N° factura', 'Correlativo cierre', 'Fecha cierre (dd/mm/aaaa)', 'Resultado', 'Fila extracto', 'Procesado el'];
// fila de lote: [accion, importe, codigo, factura, correlativo, fecha, (G)]
function hojaLote(filas, nombre = 'TRINIDAD') { return new Hoja(nombre, [ENC_LOTE.slice(), ...filas]); }

const ENC_LOTE_V2 = [...ENC_LOTE, 'Banco'];
// lote con columna Banco: filas [accion, importe, codigo, factura, correlativo, fecha, banco] -> A..F, G/H/I vacías, J
function hojaLoteV2(filas, nombre = 'TRINIDAD') {
  return new Hoja(nombre, [ENC_LOTE_V2.slice(), ...filas.map((f) => [f[0], f[1], f[2], f[3], f[4], f[5], '', '', '', f[6]])]);
}
const ENC_REGISTRO = ['Fecha cierre', 'N° factura', 'Importe', 'Banco', 'Código de asignación', 'Correlativo cierre', 'Registrado el', 'Estado', 'Duplicado'];
// REGISTRO preexistente: filas [fecha, factura, importe, banco, codigo, correlativo, registradoEl, estado]
function hojaRegistro(filas) { return new Hoja('REGISTRO', [ENC_REGISTRO.slice(), ...filas]); }

const correr = (libro) => cargarMain()(libro);

module.exports = { cargarMain, Hoja, Libro, hojaExtracto, hojaLote, hojaLoteV2, hojaRegistro, correr, ENC_LOTE, ENC_LOTE_V2, ENC_REGISTRO, val, txt };
