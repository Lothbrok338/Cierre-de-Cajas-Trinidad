'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { Libro, Hoja, hojaExtracto, hojaLoteV2, hojaLote, hojaRegistro, correr } = require('./harness');

const extractoBase = () => hojaExtracto([
  { nro: '900000001', monto: 1592 },
  { nro: '900000002', monto: 615 },
  { nro: '900000003', monto: 300.5 },
  { nro: '900000004', monto: 820, H: '29/10/2026', I: 'VRIS' },
  { nro: '900000006', monto: 410, I: 'TRINIDAD - 10', H: '40001' },       // ya marcada por la factura 40001
]);
// fila de lote V2: [accion, importe, codigo, factura, correl, fecha, banco]
const F = '16/10/2026', CC = 'CC-2026-00001';
const G = (lote, n) => lote.valorTexto(n, 6);
const HI = (ext, fila) => [ext.valorTexto(fila - 1, 7), ext.valorTexto(fila - 1, 8)];
const reg = (lib) => lib.getWorksheet('REGISTRO');
const filasReg = (lib) => { const h = reg(lib); const out = []; for (let r = 1; ; r++) { if (h.valorTexto(r, 1) === '' && h.valorTexto(r, 4) === '') break; out.push([0, 1, 2, 3, 4, 5, 7].map((c) => h.valorTexto(r, c))); } return out; };
const formulaEsperada = (n) => `=IF(H${n}="LIBERADO","",IF(OR(COUNTIFS($D:$D,D${n},$E:$E,E${n},$H:$H,"VIGENTE")>1,COUNTIFS($B:$B,B${n},$H:$H,"VIGENTE")>1),"DUPLICADO",""))`;
const snap = (h) => JSON.stringify([...h.celdas]);

test('QR feliz: marca el extracto y agrega la fila a REGISTRO (VIGENTE, banco QR BANECO, fórmula en I)', () => {
  const ext = extractoBase();
  const lote = hojaLoteV2([['', 1592, '900000001', '50001', CC, F, ''], ['MARCAR', 615, '900000002', '50002', CC, F, 'QR BANECO']]);
  const lib = new Libro([ext, lote]);
  const r = correr(lib);
  assert.equal(r, 'Lote: 2 filas · ✔ 2 · ✖ 0 · Extracto: escrito sí · Registro: +2 filas');
  assert.deepEqual(HI(ext, 5), ['50001', 'TRINIDAD - 10']);
  assert.deepEqual(HI(ext, 6), ['50002', 'TRINIDAD - 10']);
  assert.deepEqual(filasReg(lib), [
    ['16/10/2026', '50001', '1592', 'QR BANECO', '900000001', CC, 'VIGENTE'],
    ['16/10/2026', '50002', '615', 'QR BANECO', '900000002', CC, 'VIGENTE'],
  ]);
  const h = reg(lib);
  assert.equal(h.get(1, 2), 1592);                 // importe numérico
  assert.match(h.valorTexto(1, 6), /^\d\d\/\d\d\/\d{4} \d\d:\d\d:\d\d$/);
  assert.equal(h.formulas.get('1,8'), formulaEsperada(2));
  assert.equal(h.formulas.get('2,8'), formulaEsperada(3));
  assert.equal(h.formatos.get('1,1'), '@'); assert.equal(h.formatos.get('1,4'), '@'); assert.equal(h.formatos.get('1,5'), '@');
  assert.equal(lote.valorTexto(0, 10), r);
  assert.equal(G(lote, 1), '✔ MARCADO');
});

test('REGISTRO se crea si no existe: encabezados, formato condicional y J1 de TRINIDAD', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 1592, '900000001', '50001', CC, F]]);   // lote antiguo: sin columna Banco
  const lib = new Libro([ext, lote]);
  assert.equal(reg(lib), undefined);
  correr(lib);
  const h = reg(lib);
  assert.ok(h);
  assert.deepEqual([0, 1, 2, 3, 4, 5, 6, 7, 8].map((c) => h.valorTexto(0, c)),
    ['Fecha cierre', 'N° factura', 'Importe', 'Banco', 'Código de asignación', 'Correlativo cierre', 'Registrado el', 'Estado', 'Duplicado']);
  assert.equal(h.condFormats.length, 1);
  assert.equal(h.condFormats[0].formula, '=$I2="DUPLICADO"');
  assert.equal(h.formatos.get('col1'), '@'); assert.equal(h.formatos.get('col4'), '@'); assert.equal(h.formatos.get('col5'), '@');
  assert.equal(lote.valorTexto(0, 9), 'Banco');       // J1 escrito sin borrar nada
  assert.equal(lote.valorTexto(0, 0), 'Acción');
  // lote antiguo se comporta como QR
  assert.deepEqual(HI(ext, 5), ['50001', 'TRINIDAD - 10']);
  assert.equal(filasReg(lib)[0][3], 'QR BANECO');
});

test('hoja TRINIDAD inexistente: se crea junto con REGISTRO', () => {
  const lib = new Libro([extractoBase()]);
  assert.equal(correr(lib), 'Hoja creada: pegue las filas y vuelva a ejecutar');
  assert.equal(lib.getWorksheet('TRINIDAD').valorTexto(0, 9), 'Banco');
  assert.ok(reg(lib));
});

test('C.I. feliz con cada banco (BNB, BMSC, BISA, UNION, UNIÓN, BANECO, minúsculas): no toca el extracto', () => {
  const ext = extractoBase(); const antes = snap(ext);
  const lote = hojaLoteV2([
    ['', 100, 'CI-1', '60001', CC, F, 'BNB'],
    ['REGISTRAR', 200.5, 'CI-2', '60002', CC, F, 'BMSC'],
    ['', '300', 'CI-3', '60003', CC, F, 'bisa'],
    ['', 400, 'CI-4', '60004', CC, F, 'UNION'],
    ['', 500, 'CI-5', '60005', CC, F, 'UNIÓN'],
    ['', 600, 'CI-6', '60006', CC, F, 'BANECO'],
  ]);
  const lib = new Libro([ext, lote]);
  const r = correr(lib);
  assert.equal(r, 'Lote: 6 filas · ✔ 6 · ✖ 0 · Extracto: escrito no · Registro: +6 filas');
  assert.equal(snap(ext), antes);
  assert.deepEqual(filasReg(lib).map((f) => f[3]), ['BNB', 'BMSC', 'BISA', 'UNION', 'UNION', 'BANECO']);
  assert.equal(G(lote, 1), '✔ REGISTRADO');
  assert.equal(reg(lib).get(2, 2), 200.5);
  for (let i = 0; i < 6; i++) assert.equal(reg(lib).formulas.get((i + 1) + ',8'), formulaEsperada(i + 2));
  assert.equal(lote.valorTexto(1, 7), '');            // sin fila de extracto
});

test('lote solo con C.I. funciona aunque el extracto no sea válido', () => {
  const ext = new Hoja('Hoja1', [['nada']]);
  const lote = hojaLoteV2([['', 100, 'CI-1', '60001', CC, F, 'BNB']]);
  assert.match(correr(new Libro([ext, lote])), /✔ 1 · ✖ 0 · Extracto: escrito no · Registro: \+1 filas/);
});

test('mezcla QR y C.I. en un lote: ambos se procesan', () => {
  const ext = extractoBase();
  const lote = hojaLoteV2([['', 1592, '900000001', '50001', CC, F, ''], ['', 100, 'CI-1', '60001', CC, F, 'BNB']]);
  const lib = new Libro([ext, lote]);
  assert.equal(correr(lib), 'Lote: 2 filas · ✔ 2 · ✖ 0 · Extracto: escrito sí · Registro: +2 filas');
  assert.equal(filasReg(lib).length, 2);
});

test('mismo banco+código ya VIGENTE con otra factura: bloquea TODO el lote y no escribe nada', () => {
  const ext = extractoBase();
  const registro = hojaRegistro([['10/10/2026', '55555', 100, 'BNB', 'CI-1', 'CC-2026-00009', '10/10/2026 10:00:00', 'VIGENTE']]);
  const lote = hojaLoteV2([['', 1592, '900000001', '50001', CC, F, ''], ['', 100, 'CI-1', '60001', CC, F, 'BNB']]);
  const lib = new Libro([ext, lote, registro]);
  const a = snap(ext), b = snap(registro);
  const r = correr(lib);
  assert.equal(r, 'Lote: 2 filas · ✔ 1 · ✖ 1 · Extracto: escrito no · Registro: +0 filas');
  assert.equal(G(lote, 2), '✖ DUPLICADO EN REGISTRO: el código CI-1 de BNB ya se usó en la factura 55555 (cierre CC-2026-00009)');
  assert.equal(G(lote, 1), '✔ MARCADO — NO MARCADO: hay errores en el lote');
  assert.equal(snap(ext), a); assert.equal(snap(registro), b);
});

test('el mismo código en OTRO banco no es duplicado', () => {
  const registro = hojaRegistro([['10/10/2026', '55555', 100, 'BNB', 'X-1', 'CC9', 'x', 'VIGENTE']]);
  const lote = hojaLoteV2([['', 100, 'X-1', '60001', CC, F, 'BISA']]);
  assert.match(correr(new Libro([extractoBase(), lote, registro])), /✔ 1 · ✖ 0/);
});

test('misma factura ya VIGENTE con otro código o banco: bloquea', () => {
  const registro = hojaRegistro([['10/10/2026', '55555', 100, 'BNB', 'CI-1', 'CC-2026-00009', 'x', 'VIGENTE']]);
  const ext = extractoBase(); const a = snap(ext);
  const esperado = '✖ FACTURA YA REGISTRADA (banco BNB, código CI-1, cierre CC-2026-00009)';
  // otro código, mismo banco
  const l1 = hojaLoteV2([['', 100, 'CI-9', '55555', CC, F, 'BNB'], ['', 615, '900000002', '55556', CC, F, '']]);
  const lib = new Libro([ext, l1, registro]);
  assert.match(correr(lib), /✔ 1 · ✖ 1 · Extracto: escrito no · Registro: \+0 filas/);
  assert.equal(G(l1, 1), esperado);
  assert.equal(G(l1, 2), '✔ MARCADO — NO MARCADO: hay errores en el lote');
  // otro banco
  const l2 = hojaLoteV2([['', 1592, '900000001', '55555', CC, F, '']]);
  correr(new Libro([ext, l2, registro]));
  assert.equal(G(l2, 1), esperado);
  assert.equal(snap(ext), a);
  assert.equal(filasReg(new Libro([registro])).length, 1);
});

test('duplicados dentro del lote: código repetido por banco y factura repetida', () => {
  const lote = hojaLoteV2([
    ['', 1, 'CI-1', '1', CC, F, 'BNB'], ['', 2, 'CI-1', '2', CC, F, 'BNB'],       // mismo banco+código
    ['', 3, 'CI-1', '3', CC, F, 'BISA'],                                            // otro banco: bien
    ['', 4, 'CI-4', '9', CC, F, 'BNB'], ['', 5, 'CI-5', '9', CC, F, 'BISA'],       // misma factura
  ]);
  const lib = new Libro([extractoBase(), lote]);
  assert.match(correr(lib), /✔ 1 · ✖ 4 · Extracto: escrito no · Registro: \+0 filas/);
  assert.equal(G(lote, 1), '✖ REPETIDO EN LA LISTA'); assert.equal(G(lote, 2), '✖ REPETIDO EN LA LISTA');
  assert.match(G(lote, 3), /^✔ REGISTRADO — NO REGISTRADO/);
  assert.equal(G(lote, 4), '✖ REPETIDO EN LA LISTA (N° factura)'); assert.equal(G(lote, 5), '✖ REPETIDO EN LA LISTA (N° factura)');
  assert.equal(filasReg(lib).length, 0);
});

test('idempotencia C.I.: YA REGISTRADO, no agrega otra fila', () => {
  const registro = hojaRegistro([['16/10/2026', '60001', 100, 'BNB', 'CI-1', CC, 'x', 'VIGENTE']]);
  const lote = hojaLoteV2([['', 100, 'CI-1', '60001', CC, F, 'BNB']]);
  const lib = new Libro([extractoBase(), lote, registro]);
  assert.equal(correr(lib), 'Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito no · Registro: +0 filas');
  assert.equal(G(lote, 1), '✔ YA REGISTRADO');
  assert.equal(filasReg(lib).length, 1);
});

test('idempotencia QR: extracto ya marcado y falta la fila en REGISTRO -> YA MARCADO y la agrega (una sola vez)', () => {
  const ext = extractoBase();
  const lote = hojaLoteV2([['', 410, '900000006', '40001', CC, F, '']]);
  const lib = new Libro([ext, lote]);
  assert.equal(correr(lib), 'Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito no · Registro: +1 filas');
  assert.equal(G(lote, 1), '✔ YA MARCADO (misma factura)');
  assert.deepEqual(filasReg(lib), [['16/10/2026', '40001', '410', 'QR BANECO', '900000006', CC, 'VIGENTE']]);
  const lote2 = hojaLoteV2([['', 410, '900000006', '40001', CC, F, '']]);
  assert.equal(correr(new Libro([ext, lote2, reg(lib)])), 'Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito no · Registro: +0 filas');
  assert.equal(filasReg(new Libro([reg(lib)])).length, 1);
});

test('QR: duplicado en REGISTRO (otra factura con el mismo código) se informa antes que OCUPADO', () => {
  const ext = extractoBase();
  const registro = hojaRegistro([['10/10/2026', '40001', 410, 'QR BANECO', '900000006', 'CC-9', 'x', 'VIGENTE']]);
  const lote = hojaLoteV2([['', 410, '900000006', '40002', CC, F, '']]);
  correr(new Libro([ext, lote, registro]));
  assert.equal(G(lote, 1), '✖ DUPLICADO EN REGISTRO: el código 900000006 de QR BANECO ya se usó en la factura 40001 (cierre CC-9)');
});

test('LIBERADO no cuenta como duplicado: el código puede volver a usarse; el duplicado se detecta solo entre VIGENTE', () => {
  const registro = hojaRegistro([['10/10/2026', '55555', 100, 'BNB', 'CI-1', 'CC9', 'x', 'LIBERADO']]);
  const lote = hojaLoteV2([['', 100, 'CI-1', '60001', CC, F, 'BNB'], ['', 50, 'CI-2', '55555', CC, F, 'BISA']]);   // mismo código y misma factura que la fila liberada
  const lib = new Libro([extractoBase(), lote, registro]);
  assert.equal(correr(lib), 'Lote: 2 filas · ✔ 2 · ✖ 0 · Extracto: escrito no · Registro: +2 filas');
  assert.equal(filasReg(lib).length, 3);
  assert.equal(registro.valorTexto(1, 7), 'LIBERADO');
  assert.equal(registro.formulas.get('2,8'), formulaEsperada(3));   // las filas nuevas empiezan en la fila 3
});

test('LIBERAR QR: limpia el extracto y pone LIBERADO en REGISTRO', () => {
  const ext = extractoBase();
  const registro = hojaRegistro([['10/10/2026', '40001', 410, 'QR BANECO', '900000006', 'CC9', 'x', 'VIGENTE']]);
  const lote = hojaLoteV2([['LIBERAR', 410, '900000006', '40001', CC, F, '']]);
  const lib = new Libro([ext, lote, registro]);
  assert.equal(correr(lib), 'Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito sí · Registro: 1 liberadas');
  assert.equal(G(lote, 1), '✔ LIBERADO');
  assert.deepEqual(HI(ext, 9), ['', '']);
  assert.equal(registro.valorTexto(1, 7), 'LIBERADO');
});

test('LIBERAR QR sin fila en REGISTRO: no es error, se informa', () => {
  const ext = extractoBase();
  const lote = hojaLoteV2([['LIBERAR', 410, '900000006', '40001', CC, F, 'QR BANECO']]);
  const lib = new Libro([ext, lote]);
  assert.match(correr(lib), /✔ 1 · ✖ 0 · Extracto: escrito sí/);
  assert.equal(G(lote, 1), '✔ LIBERADO (sin fila en REGISTRO)');
  assert.deepEqual(HI(ext, 9), ['', '']);
});

test('LIBERAR C.I.: pone LIBERADO; si no está registrada -> error y no escribe', () => {
  const registro = hojaRegistro([['10/10/2026', '60001', 100, 'BNB', 'CI-1', 'CC9', 'x', 'VIGENTE'], ['10/10/2026', '60002', 100, 'BISA', 'CI-2', 'CC9', 'x', 'VIGENTE']]);
  const lote = hojaLoteV2([['LIBERAR', 100, 'CI-1', '60001', CC, F, 'BNB']]);
  const lib = new Libro([extractoBase(), lote, registro]);
  assert.equal(correr(lib), 'Lote: 1 filas · ✔ 1 · ✖ 0 · Extracto: escrito no · Registro: 1 liberadas');
  assert.equal(registro.valorTexto(1, 7), 'LIBERADO'); assert.equal(registro.valorTexto(2, 7), 'VIGENTE');
  // ahora ya no está vigente
  const lote2 = hojaLoteV2([['LIBERAR', 100, 'CI-1', '60001', CC, F, 'BNB'], ['LIBERAR', 100, 'CI-2', '60002', CC, F, 'BISA']]);
  const r = correr(new Libro([extractoBase(), lote2, registro]));
  assert.match(r, /✔ 1 · ✖ 1 · Extracto: escrito no · Registro: 0 liberadas/);
  assert.equal(G(lote2, 1), '✖ NO SE PUEDE LIBERAR: no está registrada');
  assert.equal(registro.valorTexto(2, 7), 'VIGENTE');
});

test('LIBERAR mezclado con MARCAR/REGISTRAR: error general y no escribe', () => {
  for (const alta of [['', 1592, '900000001', '50001', CC, F, ''], ['', 100, 'CI-1', '60001', CC, F, 'BNB']]) {
    const ext = extractoBase(); const a = snap(ext);
    const lote = hojaLoteV2([alta, ['LIBERAR', 410, '900000006', '40001', CC, F, '']]);
    const lib = new Libro([ext, lote]);
    const r = correr(lib);
    assert.match(r, /Extracto: escrito no · Registro: \+0 filas · Error general: el lote mezcla LIBERAR con MARCAR\/REGISTRAR/);
    assert.match(G(lote, 1), /^✖ LOTE MEZCLADO/); assert.match(G(lote, 2), /^✖ LOTE MEZCLADO/);
    assert.equal(snap(ext), a); assert.equal(filasReg(lib).length, 0);
  }
});

test('filas C.I. inválidas: sin fecha, importe no numérico, sin importe, banco inválido, acciones cruzadas', () => {
  const lote = hojaLoteV2([
    ['', 100, 'C1', '1', CC, '', 'BNB'],
    ['', 'abc', 'C2', '2', CC, F, 'BNB'],
    ['', '', 'C3', '3', CC, F, 'BNB'],
    ['', 100, 'C4', '4', CC, F, 'PAYPAL'],
    ['MARCAR', 100, 'C5', '5', CC, F, 'BNB'],
    ['REGISTRAR', 100, '900000001', '6', CC, F, ''],
    ['', 100, 'C7', '7', CC, 'xx', 'BISA'],
  ]);
  const lib = new Libro([extractoBase(), lote]);
  assert.match(correr(lib), /✔ 0 · ✖ 7 · Extracto: escrito no · Registro: \+0 filas/);
  assert.equal(G(lote, 1), '✖ SIN FECHA: complete la fecha de cierre');
  assert.equal(G(lote, 2), '✖ IMPORTE NO VÁLIDO');
  assert.match(G(lote, 3), /^✖ INCOMPLETO/);
  assert.equal(G(lote, 4), '✖ BANCO NO VÁLIDO');
  assert.match(G(lote, 5), /^✖ ACCIÓN NO VÁLIDA para C\.I\./);
  assert.match(G(lote, 6), /^✖ ACCIÓN NO VÁLIDA/);
  assert.match(G(lote, 7), /^✖ FECHA DE CIERRE NO VÁLIDA/);
  assert.equal(filasReg(lib).length, 0);
});

test('lote antiguo (9 columnas, sin J): se comporta como QR y no rompe los datos', () => {
  const ext = extractoBase();
  const lote = hojaLote([['', 1592, '900000001', '50001', CC, F], ['LIBERAR', 410, '900000006', '40001', CC, F]]);
  correr(new Libro([ext, lote]));
  assert.match(G(lote, 1), /^✖ LOTE MEZCLADO/);
  const lote2 = hojaLote([['', 1592, '900000001', '50001', CC, F]]);
  const lib = new Libro([ext, lote2]);
  assert.match(correr(lib), /✔ 1 · ✖ 0 · Extracto: escrito sí · Registro: \+1 filas/);
  assert.equal(lote2.valorTexto(0, 9), 'Banco');
});

test('varias ejecuciones acumulan en REGISTRO y las fórmulas siguen la fila', () => {
  const ext = extractoBase();
  const lote1 = hojaLoteV2([['', 1592, '900000001', '50001', CC, F, '']]);
  const lib = new Libro([ext, lote1]);
  correr(lib);
  const lote2 = hojaLoteV2([['', 100, 'CI-1', '60001', CC, F, 'BNB'], ['', 615, '900000002', '50002', CC, F, '']]);
  correr(new Libro([ext, lote2, reg(lib)]));
  assert.deepEqual(filasReg(lib).map((f) => f[1]), ['50001', '60001', '50002']);
  assert.equal(reg(lib).formulas.get('3,8'), formulaEsperada(4));
});
