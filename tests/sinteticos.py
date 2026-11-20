"""Generador de reportes SINTETICOS con nombres ficticios (para pruebas sin datos reales)."""
import datetime as dt
import openpyxl

FECHA = dt.date(2026, 10, 16)
# (nro_factura, tipo_pago, monto, razon social, estudiante, estado, hora)
FACTURAS = [
    ('30001', 'EFECTIVO', 500, 'LUCIA FERNANDEZ ROCA', 'FERNANDEZ ROCA LUCIA', 'Válido', '10:00:10'),
    ('30002', 'BANCO', 300, 'MARTIN QUISPE LEON', 'QUISPE LEON MARTIN', 'Válido', '10:30:00'),
    ('30003', 'EFECTIVO', 700, 'EMPRESA ALFA SRL', 'SOTO VEGA CARLOS', 'Válido', '11:00:00'),
    ('30004', 'TARJETA', 200, 'RAUL MENDEZ PAZ', 'MENDEZ PAZ RAUL', 'Válido', '11:30:00'),
    ('30005', 'EFECTIVO', 150, 'ELENA ROJAS DIAZ', 'ROJAS DIAZ ELENA', 'Anulado', '12:00:00'),
]
# (id, titular, hora, monto)
BANCO = [
    ('700000001', 'FERNANDEZ ROCA LUCIA', '09:59:30', 500),
    ('700000002', 'QUISPE LEON MARTIN', '10:29:40', 300),
    ('700000003', 'SOTO VEGA ANDREA', '10:58:50', 700),
    ('700000004', 'PERSONA AJENA LIMA', '12:00:00', 900),
]


def crear_a1(ruta, facturas=FACTURAS, caja='SFC999', fecha=FECHA):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'RIngresoDiario'
    ws.append([]); ws.append(['UNIVERSIDAD PRIVADA DEL VALLE', 'Usuario:', 'CAJERO1']); ws.append(['SISTEMA DE FACTURACIÓN', 'Fecha:', '16/10/2026 17:00'])
    ws.append(['Caja:', caja]); ws.append([]); ws.append(['INGRESO DIARIO']); ws.append([]); ws.append([])
    cols = ['Nro', 'Fecha', 'Número Factura', 'Nit/C.I.', 'Razon Social', 'Nombre Estudiante', 'Tipo Pago', 'Banco / Cuenta bancaria',
            'Código de Bancarización', 'Nro. Cuenta SAP asociada', 'Monto', 'Cuenta', 'Estado']
    ws.append(cols)
    for i, (nf, tp, m, rs, es, est, h) in enumerate(facturas, start=1):
        hh = dt.datetime.strptime(h, '%H:%M:%S').time()
        ws.append([i, dt.datetime.combine(fecha, hh), nf, '123', rs, es, tp, '-', '-', '-', float(m), 'USR', est])
    wb.save(ruta); return ruta


def crear_a2(ruta, banco=BANCO, quitar=(), fecha=FECHA):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(['ID QR', 'ID TRANSACCION', 'TITULAR ORIGEN', 'CUENTA ORIGEN', 'GLOSA', 'CUENTA DESTINO', 'FECHA / HORA', 'PUNTOS DE VENTA', 'MONTO'])
    tot = 0
    for idt, tit, h, m in banco:
        if idt in quitar: continue
        ws.append(['999', idt, tit, '******0001', 'Ingresos', '3041', f"{fecha:%d/%m/%Y} {h}", 'SEDE X', f'{m:,.2f} BOB']); tot += m
    ws.append([None] * 7 + ['TOTAL', f'{tot:,.2f} BOB'])
    wb.save(ruta); return ruta
