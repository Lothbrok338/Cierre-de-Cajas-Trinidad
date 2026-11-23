'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { Libro, hojaExtracto, hojaLote, correr, Hoja } = require('./harness');

// Datos ficticios. Extracto: filas 5.. (índice 4..). Códigos 9000000NN.
function extractoBase(extra = {}) {
  return hojaExtracto([
    { nro: '900000001', monto: 1592 },                                   // libre
    { nro: '900000002', monto: 615 },                                    // libre
    { nro: '900000003', monto: 300.5 },                                  // libre
    { nro: '900000004', monto: 820, H: '29/10/2026', I: 'VRIS' },        // ocupada
    { nro: 900000005, monto: 100 },                                      // código numérico, libre
    { nro: '900000006', monto: 410, I: 'TRINIDAD - 10', H: '40001' },    // ya marcada por la factura 40001
    { nro: '900000007', monto: 55, I: 'LIDIA EN OFICINAS', H: { v: 46324, t: '29/10/2026' } }, // fecha real en H
    { nro: '900000008', monto: 77, J: 'TRINIDAD - 10' },                 // solo J
    ...(extra.filas || []),
  ]);
}
const LIBRE = (i) => ({ r: 4 + i });   // fila Excel = índice + 1 -> 5 + i
const libro = (ext, lote) => new Libro([ext, ...(lote ? [lote] : [])]);
const G = (lote, n) => lote.valorTexto(n, 6);        // resultado de la fila n del lote (índice 0 = encabezado)
const HI = (ext, fila) => [ext.valorTexto(fila - 1, 7), ext.valorTexto(fila - 1, 8)];

test('todo libre: marca H (factura como texto) e I con el mes sin cero', () => {
  const ext = extractoBase();
  const lote = hojaLote([
    ['', 1592, '900000001', '50001', 'CC-2026-00001', '16/10/2026'],
    ['MARCAR', '615', '900000002', '50002', 'CC-2026-00001', '16/10/2026'],
    ['', 100, '900000005', '50003', 'CC-2026-00001', '16/10/2026'],     // el código del extracto es numérico
  ]);
  const r = correr(libro(ext, lote));
  assert.ok(String(r).startsWith('Lote: 3 filas · ✔ 3 · ✖ 0 · Extracto: escrito sí'));
  assert.deepEqual(HI(ext, 5), ['50001', 'TRINIDAD - 10']);
  assert.deepEqual(HI(ext, 6), ['50002', 'TRINIDAD - 10']);
  assert.deepEqual(HI(ext, 9), ['50003', 'TRINIDAD - 10']);
  assert.equal(typeof ext.get(4, 7), 'string');
  assert.equal(ext.formatos.get('4,7'), '@');
  assert.equal(G(lote, 1), '✔ MARCADO');
  assert.equal(lote.get(1, 7), 5); assert.equal(lote.get(2, 7), 6); assert.equal(lote.get(3, 7), 9);
  assert.match(lote.valorTexto(1, 8), /^\d\d\/\d\d\/\d{4} \d\d:\d\d:\d\d$/);
  assert.equal(lote.valorTexto(0, 10), r);   // K1
});

test('mes de la marca: de F; si F vacía, de la fecha de la fila del extracto', () => {
  const ext = hojaExtracto([{ nro: '900000011', monto: 10, fecha: '03/Nov/2026' }, { nro: '900000012', monto: 20, fecha: '16/Oct/2026' }]);
  const lote = hojaLote([['', 10, '900000011', '1', 'CC', ''], ['', 20, '900000012', '2', 'CC', '05/01/2027']]);
  correr(libro(ext, lote));
  assert.equal(ext.valorTexto(4, 8), 'TRINIDAD - 11');
  assert.equal(ext.valorTexto(5, 8), 'TRINIDAD - 1');
});

test('fecha de cierre como número de serie de Excel', () => {
  const ext = hojaExtracto([{ nro: '900000011', monto: 10 }]);
  const lote = hojaLote([['', 10, '900000011', '1', 'CC', 46311]]);   // 2026-10-16
  assert.match(correr(libro(ext, lote)), /✔ 1 · ✖ 0 · Extracto: escrito sí/);
  assert.equal(ext.valorTexto(4, 8), 'TRINIDAD - 10');
});

test('una fila ocupada: no escribe nada y las válidas quedan "NO MARCADO"', () => {
  const ext = extractoBase();
  const antes = JSON.stringify([...ext.celdas]);
  const lote = hojaLote([
    ['', 1592, '900000001', '50001', 'CC', '16/10/2026'],
    ['', 820, '900000004', '50002', 'CC', '16/10/2026'],
  ]);
  const r = correr(libro(ext, lote));
  assert.ok(String(r).startsWith('Lote: 2 filas · ✔ 1 · ✖ 1 · Extracto: escrito no'));
  assert.equal(JSON.stringify([...ext.celdas]), antes);
  assert.equal(G(lote, 1), '✔ MARCADO — NO MARCADO: hay errores en el lote');
  assert.equal(G(lote, 2), '✖ OCUPADO: H=29/10/2026 | I=VRIS | J=');
});

test('ocupado muestra la fecha real de H tal como se ve y J ocupada', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 55, '900000007', '1', 'CC', '16/10/2026'], ['', 77, '900000008', '2', 'CC', '16/10/2026']]);
  correr(libro(ext, lote));
  assert.equal(G(lote, 1), '✖ OCUPADO: H=29/10/2026 | I=LIDIA EN OFICINAS | J=');
  assert.equal(G(lote, 2), '✖ OCUPADO: H= | I= | J=TRINIDAD - 10');
});

test('no existe en el extracto', () => {
  const ext = extractoBase(); const lote = hojaLote([['', 5, '999999999', '1', 'CC', '16/10/2026']]);
  assert.ok(String(correr(libro(ext, lote))).startsWith('Lote: 1 filas · ✔ 0 · ✖ 1 · Extracto: escrito no'));
  assert.equal(G(lote, 1), '✖ NO EXISTE EN EXTRACTO');
});

test('monto distinto (compara numérico, el extracto trae texto)', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 1592.01, '900000001', '1', 'CC', '16/10/2026'], ['', '1592.00', '900000001', '1', 'CC', '16/10/2026']]);
  correr(libro(ext, lote));
  // el mismo código dos veces es REPETIDO; probamos el monto aparte
  const ext2 = extractoBase(); const lote2 = hojaLote([['', 1592.01, '900000001', '1', 'CC', '16/10/2026']]);
  correr(libro(ext2, lote2));
  assert.equal(G(lote2, 1), '✖ MONTO DISTINTO (extracto 1592.00)');
  assert.deepEqual(HI(ext2, 5), ['', '']);
  const ext3 = extractoBase(); const lote3 = hojaLote([['', '300,5', '900000003', '1', 'CC', '16/10/2026']]);   // tolerante a coma decimal
  assert.match(correr(libro(ext3, lote3)), /✔ 1/);
});

test('repetido dentro del lote', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 1592, '900000001', '1', 'CC', '16/10/2026'], ['', 1592, '900000001', '2', 'CC', '16/10/2026'], ['', 615, '900000002', '3', 'CC', '16/10/2026']]);
  assert.ok(String(correr(libro(ext, lote))).startsWith('Lote: 3 filas · ✔ 1 · ✖ 2 · Extracto: escrito no'));
  assert.equal(G(lote, 1), '✖ REPETIDO EN LA LISTA'); assert.equal(G(lote, 2), '✖ REPETIDO EN LA LISTA');
  assert.deepEqual(HI(ext, 6), ['', '']);
});

test('idempotencia: re-ejecutar las mismas filas tras limpiar G da YA MARCADO y no reescribe', () => {
  const ext = extractoBase();
  const fila = () => ['', 1592, '900000001', '50001', 'CC', '16/10/2026'];
  const lote = hojaLote([fila()]);
  const lib = libro(ext, lote);
  assert.match(correr(lib), /Extracto: escrito sí/);
  for (let c = 6; c <= 8; c++) lote.set(1, c, '');            // el usuario limpia G (y el resto) para reprocesar
  const n = ext.escrituras.length;
  const r = correr(lib);
  assert.equal(G(lote, 1), '✔ YA MARCADO (misma factura)');
  assert.ok(String(r).startsWith('Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito no'));
  assert.deepEqual(HI(ext, 5), ['50001', 'TRINIDAD - 10']);
  // misma fila, otra factura -> ocupada
  const lote2 = hojaLote([['', 1592, '900000001', '50099', 'CC', '16/10/2026']]);
  correr(new Libro([ext, lote2]));
  assert.equal(G(lote2, 1), '✖ OCUPADO: H=50001 | I=TRINIDAD - 10 | J=');
  void n;
});

test('YA MARCADO preexistente (factura en H como texto e I con la marca) y el N° como número', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 410, '900000006', 40001, 'CC', '16/10/2026']]);
  correr(libro(ext, lote));
  assert.equal(G(lote, 1), '✔ YA MARCADO (misma factura)');
});

test('YA MARCADO exige el mismo mes', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 410, '900000006', '40001', 'CC', '16/11/2026']]);
  correr(libro(ext, lote));
  assert.match(G(lote, 1), /^✖ OCUPADO/);
});

test('LIBERAR correcto: limpia H e I', () => {
  const ext = extractoBase();
  const lote = hojaLote([['LIBERAR', 410, '900000006', '40001', 'CC', '16/10/2026']]);
  assert.ok(String(correr(libro(ext, lote))).startsWith('Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito sí'));
  assert.equal(G(lote, 1), '✔ LIBERADO (sin fila en REGISTRO)');
  assert.deepEqual(HI(ext, 10), ['', '']);
});

test('LIBERAR incorrecto (otra factura, o marca de otra sede/mes): no escribe', () => {
  const ext = extractoBase({ filas: [{ nro: '900000010', monto: 9, H: '40002', I: 'TRINIDAD - 10' }] });
  const lote = hojaLote([['LIBERAR', 410, '900000006', '49999', 'CC', '16/10/2026'], ['LIBERAR', 820, '900000004', '1', 'CC', '16/10/2026'], ['LIBERAR', 9, '900000010', '40002', 'CC', '16/10/2026']]);
  const r = correr(libro(ext, lote));
  assert.ok(String(r).startsWith('Lote: 3 filas · ✔ 1 · ✖ 2 · Extracto: escrito no'));
  assert.equal(G(lote, 1), '✖ NO SE PUEDE LIBERAR: la marca no corresponde a esta factura');
  assert.equal(G(lote, 2), '✖ NO SE PUEDE LIBERAR: la marca no corresponde a esta factura');
  assert.equal(G(lote, 3), '✔ LIBERADO (sin fila en REGISTRO) — NO LIBERADO: hay errores en el lote');
  assert.deepEqual(HI(ext, 10), ['40001', 'TRINIDAD - 10']);
  assert.deepEqual(HI(ext, 13), ['40002', 'TRINIDAD - 10']);
});

test('mezcla MARCAR/LIBERAR: error general y no escribe', () => {
  const ext = extractoBase(); const antes = JSON.stringify([...ext.celdas]);
  const lote = hojaLote([['MARCAR', 1592, '900000001', '1', 'CC', '16/10/2026'], ['LIBERAR', 410, '900000006', '40001', 'CC', '16/10/2026']]);
  const r = correr(libro(ext, lote));
  assert.match(r, /Extracto: escrito no · Registro: \+0 filas · Error general: el lote mezcla LIBERAR con MARCAR\/REGISTRAR/);
  assert.equal(JSON.stringify([...ext.celdas]), antes);
  assert.match(G(lote, 1), /^✖ LOTE MEZCLADO/);
});

test('solo se procesan filas con G vacío; las ya procesadas no se tocan', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 1592, '900000001', '1', 'CC', '16/10/2026', '✔ MARCADO'], ['', 615, '900000002', '2', 'CC', '16/10/2026']]);
  assert.ok(String(correr(libro(ext, lote))).startsWith('Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito sí'));
  assert.deepEqual(HI(ext, 5), ['', '']);
  assert.deepEqual(HI(ext, 6), ['2', 'TRINIDAD - 10']);
});

test('fila incompleta (falta factura) es un error del lote', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 1592, '900000001', '', 'CC', '16/10/2026'], ['', 615, '900000002', '2', 'CC', '16/10/2026']]);
  assert.match(correr(libro(ext, lote)), /✖ 1 · Extracto: escrito no/);
  assert.match(G(lote, 1), /^✖ INCOMPLETO/);
});

test('acción no válida', () => {
  const ext = extractoBase(); const lote = hojaLote([['BORRAR', 1592, '900000001', '1', 'CC', '16/10/2026']]);
  correr(libro(ext, lote));
  assert.match(G(lote, 1), /^✖ ACCIÓN NO VÁLIDA/);
});

test('hoja TRINIDAD inexistente: se crea con encabezados y termina con el mensaje', () => {
  const ext = extractoBase(); const lib = libro(ext);
  const r = correr(lib);
  assert.equal(r, 'Hoja creada: pegue las filas y vuelva a ejecutar');
  const h = lib.getWorksheet('TRINIDAD');
  assert.ok(h);
  assert.deepEqual(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].map((_, i) => h.valorTexto(0, i)),
    ['Acción', 'Importe', 'Código de asignación', 'N° factura', 'Correlativo cierre', 'Fecha cierre (dd/mm/aaaa)', 'Resultado', 'Fila extracto', 'Procesado el']);
  assert.equal(h.formatos.get('col2'), '@');   // C como texto
  assert.deepEqual(HI(ext, 5), ['', '']);
});

test('encabezado del extracto en otra fila y hoja con otro nombre', () => {
  const ext = hojaExtracto([{ nro: '900000001', monto: 1592 }], { encabezadoEnFila: 7, nombre: 'Extracto octubre' });
  const lote = hojaLote([['', 1592, '900000001', '1', 'CC', '16/10/2026']]);
  assert.match(correr(new Libro([ext, lote])), /✔ 1 · ✖ 0 · Extracto: escrito sí/);
  assert.equal(lote.get(1, 7), 8);                         // fila Excel 8 (encabezado 7)
  assert.deepEqual(HI(ext, 8), ['1', 'TRINIDAD - 10']);
});

test('encabezado en fila 4 (formato real): la fila escrita es 5', () => {
  const ext = hojaExtracto([{ nro: '900000001', monto: 1592 }]);
  const lote = hojaLote([['', 1592, '900000001', '1', 'CC', '16/10/2026']]);
  correr(libro(ext, lote));
  assert.equal(lote.get(1, 7), 5);
});

test('sin encabezado reconocible: error y no escribe', () => {
  const ext = new Hoja('Hoja1', [['nada'], ['otra cosa']]); const lote = hojaLote([['', 1, '9', '1', 'CC', '16/10/2026']]);
  assert.match(correr(libro(ext, lote)), /^ERROR: no se encontró el encabezado/);
});

test('encabezado tolerante a mayúsculas, tildes y espacios', () => {
  const ext = hojaExtracto([{ nro: '900000001', monto: 1592 }]);
  ext.set(3, 0, '  FECHA '); ext.set(3, 2, 'nro   TRANSACCION');
  const lote = hojaLote([['', 1592, '900000001', '1', 'CC', '16/10/2026']]);
  assert.match(correr(libro(ext, lote)), /Extracto: escrito sí/);
});

test('lote vacío', () => {
  const ext = extractoBase(); const lote = hojaLote([]);
  assert.match(correr(libro(ext, lote)), /^Lote: 0 filas · ✔ 0 · ✖ 0 · Extracto: escrito no/);
});
