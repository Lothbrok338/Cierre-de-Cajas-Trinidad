'use strict';
// Extremo a extremo: el TSV que genera la app (servicio.tsv_extracto, datos ficticios) pasa por el script.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const { execFileSync } = require('child_process');
const { Hoja, Libro, hojaExtracto, ENC_LOTE_V2, correr } = require('./harness');

const REPO = path.join(__dirname, '..', '..');
const PY = `
import sys, json
sys.path.insert(0, sys.argv[1])
from app import servicio
qr = [{'id_transaccion': '900000001', 'factura': '50001', 'monto': 1592.0}, {'id_transaccion': '900000002', 'factura': '50002', 'monto': 615.0}]
ci = [{'factura': '50004', 'banco': 'UNION', 'codigo': '4567', 'monto': 300.0}, {'factura': '50003', 'banco': 'BNB', 'codigo': 'P123', 'monto': 410.5}]
print(json.dumps({'marcar': servicio.tsv_extracto('CC-2026-00009', '2026-10-16', qr, '', ci)[0],
                  'liberar': servicio.tsv_extracto('CC-2026-00009', '2026-10-16', qr, 'LIBERAR', ci)[0]}))
`;
let tsv = null, motivo = false;
try { tsv = JSON.parse(execFileSync('python3', ['-I', '-c', PY, REPO]).toString()); } catch (e) { motivo = 'no se pudo importar la app Python: ' + String(e.message).split('\n')[0]; }

const loteDe = (t) => new Hoja('TRINIDAD', [ENC_LOTE_V2.slice(), ...t.split('\n').map((l) => l.split('\t'))]);
const extracto = () => hojaExtracto([{ nro: '900000001', monto: 1592 }, { nro: '900000002', monto: 615 }]);

test('el TSV de la app: QR se marca en el extracto y las C.I. se registran', { skip: motivo }, () => {
  assert.ok(tsv.marcar.split('\n').every((l) => l.split('\t').length === 10));
  const ext = extracto(), lote = loteDe(tsv.marcar), lib = new Libro([ext, lote]);
  const r = correr(lib);
  assert.equal(r, 'Lote: 4 filas · ✔ 4 · ✖ 0 · Extracto: escrito sí · Registro: +4 filas');
  assert.equal(ext.valorTexto(4, 7), '50001'); assert.equal(ext.valorTexto(5, 7), '50002');
  const reg = lib.getWorksheet('REGISTRO');
  assert.deepEqual([1, 2, 3, 4].map((i) => reg.valorTexto(i, 3) + ':' + reg.valorTexto(i, 1)), ['QR BANECO:50001', 'QR BANECO:50002', 'BNB:50003', 'UNION:50004']);
  assert.equal(reg.valorTexto(3, 4), 'P123');
});

test('el TSV de liberar de la app libera todo lo anterior', { skip: motivo }, () => {
  const ext = extracto(), lib = new Libro([ext, loteDe(tsv.marcar)]);
  correr(lib);
  const lote2 = loteDe(tsv.liberar); lib.hojas[1] = lote2;
  const r = correr(lib);
  assert.match(r, /^Lote: 4 filas · ✔ 4 · ✖ 0/);
  assert.equal(ext.valorTexto(4, 7), ''); assert.equal(lib.getWorksheet('REGISTRO').valorTexto(1, 7), 'LIBERADO');
});
