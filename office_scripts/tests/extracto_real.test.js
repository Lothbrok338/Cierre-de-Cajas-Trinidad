'use strict';
// Prueba con la copia REAL del extracto (no está en el repo). Se salta si no existe.
//   EXTRACTO_REAL=/ruta/extracto.xlsx  REPORTE_REAL=/ruta/Reporte_pagos.xlsx  npm test
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { Hoja, Libro, hojaLote, correr } = require('./harness');

const EXTRACTO = process.env.EXTRACTO_REAL || '';
const REPORTE = process.env.REPORTE_REAL || '';
const hay = EXTRACTO && REPORTE && fs.existsSync(EXTRACTO) && fs.existsSync(REPORTE);

const PY = `
import sys, json, datetime, openpyxl
def cel(v):
    if isinstance(v, datetime.datetime):
        return {'v': (v - datetime.datetime(1899, 12, 30)).days, 't': v.strftime('%d/%m/%Y')}
    return v
ex = openpyxl.load_workbook(sys.argv[1], data_only=True); we = ex[ex.sheetnames[0]]
filas = [[cel(c) for c in r] for r in we.iter_rows(max_col=10, values_only=True)]
rp = openpyxl.load_workbook(sys.argv[2], data_only=True).active
ids = []
for r in rp.iter_rows(min_row=2, values_only=True):
    if r[1] and r[6]:
        ids.append([str(r[1]).strip(), str(r[6]).strip()[:10], str(r[8]).replace('BOB', '').replace(',', '').strip()])
print(json.dumps({'hoja': we.title, 'filas': filas, 'ids': ids}))
`;

test('extracto real: los códigos del reporte del 16/10 existen y salen OCUPADO (ya usados a mano con "TRINIDAD - 10")', { skip: hay ? false : 'EXTRACTO_REAL/REPORTE_REAL no definidos o no existen' }, () => {
  const salida = execFileSync('python3', ['-I', '-c', PY, EXTRACTO, REPORTE], { maxBuffer: 64 * 1024 * 1024 }).toString();
  const d = JSON.parse(salida);
  const ids = d.ids.filter((x) => x[1] === '16/10/2026');
  assert.equal(ids.length, 14, 'el reporte debería traer 14 pagos del 16/10');
  const ext = new Hoja(d.hoja, d.filas);
  const lote = hojaLote(ids.map((x, i) => ['', x[2], x[0], String(60000 + i), 'CC-2026-99999', '16/10/2026']));
  const antes = JSON.stringify([...ext.celdas]);
  const lib = new Libro([ext, lote]);
  const r = correr(lib);
  assert.ok(String(r).startsWith('Lote: 14 filas · ✔ 0 · ✖ 14 · Extracto: escrito no'));
  for (let i = 1; i <= 14; i++) {
    const g = lote.valorTexto(i, 6);
    assert.match(g, /^✖ OCUPADO: .*I=TRINIDAD - 10/, `fila ${i}: ${g}`);
  }
  assert.equal(JSON.stringify([...ext.celdas]), antes);
  // nada se escribe en REGISTRO: solo el encabezado
  const registro = lib.getWorksheet('REGISTRO');
  assert.ok(registro);
  assert.equal(registro.valorTexto(1, 1), '');
  assert.equal(registro.valorTexto(1, 4), '');
});
