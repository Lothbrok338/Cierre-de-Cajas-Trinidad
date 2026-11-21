"""PDF del cierre: pagina 1 (resumen) + anexo (detalle)."""
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether

from .formato import bs, fecha, fecha_hora, hora

import os
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'logo-univalle.png')
AZUL = colors.HexColor('#870B42'); GRIS = colors.HexColor('#535352'); LINEA = colors.HexColor('#d9cdd3'); FONDO = colors.HexColor('#F6E7EE')
AMBAR = colors.HexColor('#fff1cc'); VERDE = colors.HexColor('#dff3e4'); ROJO = colors.HexColor('#f9dcdc')

# Simbolo de alerta: Helvetica no tiene el glifo, se usa DejaVu Sans (incluida en app/fonts); sin la fuente se usa "(!)".
ALERTA = '(!)'
try:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pdfmetrics.registerFont(TTFont('DejaVuSans', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts', 'DejaVuSans.ttf')))
    ALERTA = '<font name="DejaVuSans" color="#b45f00">\u26a0</font>'
except Exception: pass

_ss = getSampleStyleSheet()
H1 = ParagraphStyle('h1', parent=_ss['Title'], fontName='Helvetica-Bold', fontSize=16, textColor=AZUL, alignment=0, spaceAfter=2)
H2 = ParagraphStyle('h2', parent=_ss['Heading2'], fontName='Helvetica-Bold', fontSize=11, textColor=AZUL, spaceBefore=10, spaceAfter=4)
N = ParagraphStyle('n', parent=_ss['Normal'], fontName='Helvetica', fontSize=9, leading=11.5)
C = ParagraphStyle('c', parent=N, fontSize=7.5, leading=9)
CB = ParagraphStyle('cb', parent=C, fontName='Helvetica-Bold')
NOTA = ParagraphStyle('nota', parent=N, fontSize=7.5, textColor=GRIS, leading=9.5)
DER = ParagraphStyle('der', parent=N, alignment=TA_RIGHT)


def _p(t, st=N):
    t = str(t if t is not None else '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return Paragraph(t, st)


def _pa(t, st=N):
    """Como _p pero antepone el simbolo de alerta (el texto se escapa)."""
    t = str(t if t is not None else '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return Paragraph(f'{ALERTA} {t}', st)


def _col_estado(e):
    e = str(e)
    if e.startswith('PREVALIDADO'): return FONDO
    if e.startswith('CUADRA') or e in ('CONCILIADA', 'CONFORME'): return VERDE
    if 'CONFIRMADA' in e or 'PENDIENTE' in e or 'CONFIRMAR' in e or 'REEMPLAZADO' in e: return AMBAR
    return ROJO


def generar(res, info, ruta):
    from .servicio import conceptos, total_conceptos
    t = res['totales']; f = res['facturas']; b = res['banco']
    doc = SimpleDocTemplate(ruta, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=14 * mm, bottomMargin=20 * mm,
                            title=f"Comprobante de cierre de caja {info['correlativo']}", author='UNIVALLE')
    S = []
    alertas = info.get('alertas_formato') or []
    marcadas = {(a['factura'], a['banco']) for a in alertas}
    if os.path.isfile(LOGO):
        S.append(Image(LOGO, width=52 * mm, height=52 * mm * 134 / 526, hAlign='LEFT')); S.append(Spacer(1, 3))
    else: S.append(_p('UNIVALLE', H1))
    estado = info.get('estado_revision') or 'PENDIENTE DE REVISIÓN'
    # Recuadro de estado: caja propia de altura fija, apoyada sobre la linea inferior del titulo (no dentro de su celda, que la estiraba)
    caja_estado = Table([[_p(estado, ParagraphStyle('ins', parent=CB, fontSize=8.5, leading=10.5, alignment=1))]], colWidths=[52 * mm], rowHeights=[8 * mm])
    caja_estado.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), _col_estado(estado)), ('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    cab = Table([[_p('Comprobante de cierre de caja', H1), caja_estado]], colWidths=[122 * mm, 52 * mm])
    cab.setStyle(TableStyle([('LINEBELOW', (0, 0), (-1, -1), 1.5, AZUL), ('LEFTPADDING', (0, 0), (0, -1), 0), ('RIGHTPADDING', (1, 0), (1, -1), 0),
                             ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 4), ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'), ('ALIGN', (1, 0), (1, 0), 'RIGHT')]))
    S.append(cab)
    S.append(Spacer(1, 3))
    S.append(_p(f"Correlativo {info['correlativo']}", ParagraphStyle('c2', parent=N, fontName='Helvetica-Bold', fontSize=11, textColor=AZUL)))
    S.append(Spacer(1, 4))
    enc = [[_p('Fecha del cierre', CB), _p(fecha(info['fecha_cierre']), C), _p('Caja', CB), _p(info.get('caja') or '-', C)],
           [_p('Cajero', CB), _p(f"{info['nombre']} ({info['usuario']})", C), _p('Presentado el', CB), _p(fecha_hora(info['presentado_en']), C)],
           [_p('Versión', CB), _p(f"v{info['version']}", C), _p('Estado de revisión', CB), _p(info.get('estado_revision') or 'PENDIENTE DE REVISIÓN', CB)]]
    if info.get('motivo_reemplazo'):
        enc.append([_p('Reemplaza a', CB), _p(info.get('reemplaza_a') or '-', C), _p('Motivo', CB), _p(info['motivo_reemplazo'], C)])
    if info.get('revisado_por'):
        enc.append([_p('Revisado por', CB), _p(info.get('revisado_por_nombre') or info['revisado_por'], C), _p('Revisado el', CB), _p(fecha_hora(info['revisado_en']), C)])
        if info.get('observacion_revision'): enc.append([_p('Observación de Ingresos', CB), _p(info['observacion_revision'], C), '', ''])
    te = Table(enc, colWidths=[30 * mm, 60 * mm, 28 * mm, 56 * mm])
    te.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (0, -1), FONDO),
                            ('BACKGROUND', (2, 0), (2, -1), FONDO), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    if info.get('observacion_revision'): te.setStyle(TableStyle([('SPAN', (1, len(enc) - 1), (3, len(enc) - 1)), ('BACKGROUND', (1, len(enc) - 1), (3, len(enc) - 1), colors.white)]))
    S.append(te)

    S.append(_p('Cuadro principal', H2))
    lista = conceptos(t); tot = total_conceptos(lista)
    filas = [[_p(h, CB) for h in ('Medio de pago', 'Según sistema (Bs)', 'Según respaldo (Bs)', 'Diferencia (Bs)', 'Estado')]]
    estilos = [('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (-1, 0), FONDO),
               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (1, 1), (3, -1), 'RIGHT')]
    for i, c in enumerate(lista + [tot], start=1):
        st = CB if c['concepto'] == 'TOTAL' else C
        filas.append([_p(c['concepto'], st), _p(bs(c['segun_sistema']), st), _p(bs(c['segun_respaldo']), st), _p(bs(c['diferencia']), st), _p(c['estado'], st)])
        estilos.append(('BACKGROUND', (4, i), (4, i), _col_estado(c['estado'])))
    tc = Table(filas, colWidths=[40 * mm, 30 * mm, 30 * mm, 28 * mm, 46 * mm]); tc.setStyle(TableStyle(estilos))
    for r in filas[1:]:
        for cell in r[1:4]: cell.style = ParagraphStyle('rr', parent=cell.style, alignment=TA_RIGHT)
    S.append(tc)
    for c in lista:
        if c['nota']: S.append(_p(f"{c['concepto']}: {c['nota']}.", NOTA))

    S.append(_p('Resumen de conciliación QR BANECO', H2))
    ajenos = b[b['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO']
    anul = f[f['estado_conciliacion'].str.startswith('EXCLUIDA')]
    res_rows = [['Facturas QR BANECO del cierre', str(t['n_facturas_qr'])],
                ['Conciliadas automáticamente', str(t['n_CONCILIADA'])],
                ['Confirmadas por el cajero', str(t['n_CONFIRMADA POR CAJERO'])],
                ['Depósitos de otros cajeros excluidos', f"{len(ajenos)}  (Bs {bs(ajenos['MONTO'].sum())})"],
                *([['Comunicaciones internas (otras cuentas)', f"{t['n_COMUNICACIÓN INTERNA']}  (Bs {bs(t['C.I. según sistema'])})"]] if t['n_COMUNICACIÓN INTERNA'] else []),
                ['Facturas anuladas excluidas', f"{len(anul)}  (Bs {bs(anul['Monto'].sum())})"]]
    tr = Table([[_p(a, C), _p(v, CB)] for a, v in res_rows], colWidths=[90 * mm, 60 * mm])
    tr.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (0, -1), FONDO)]))
    S.append(tr)

    S.append(_p('Declaración', H2))
    S.append(_p('El cajero declara conforme el presente cierre.', ParagraphStyle('d', parent=N, fontName='Helvetica-Bold')))
    S.append(_p(f"Firmado electrónicamente por sesión de usuario: {info['usuario']} — {info['nombre']}, el {fecha_cierre_hora(info['presentado_en'])}.", N))
    S.append(Spacer(1, 4))
    S.append(_p('Huellas de los archivos originales (SHA-256, abreviadas)', CB))
    S.append(_p(f"Movimiento del cajero: {info['sha256_movimiento'][:16]}…{info['sha256_movimiento'][-8:]}", C))
    S.append(_p(f"Reporte QR BANECO del banco: {info['sha256_banco'][:16]}…{info['sha256_banco'][-8:]}", C))
    if info.get('sha256_plantilla'):
        S.append(_p(f"Plantilla de comunicaciones internas: {info['sha256_plantilla'][:16]}…{info['sha256_plantilla'][-8:]}", C))
    S.append(Spacer(1, 6))
    S.append(_p('Las tarjetas se declaran por el cajero y quedan pendientes de verificación por Ingresos.', NOTA))

    # ---- anexo ----
    S.append(PageBreak())
    S.append(_p(f"Anexo — Detalle de facturas QR BANECO · {info['correlativo']}", H1))
    qr = f[f['Tipo (motor)'] == 'QR']
    filas = [[_p(h, CB) for h in ('N° factura', 'Cliente', 'Monto (Bs)', 'ID TRANSACCION', 'Titular banco', 'Estado')]]
    est = [('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (-1, 0), FONDO), ('VALIGN', (0, 0), (-1, -1), 'TOP')]
    for i, (_, r) in enumerate(qr.iterrows(), start=1):
        idc = _pa(r['ID TRANSACCION'], C) if (str(r['Número Factura']), 'QR BANECO') in marcadas else _p(r['ID TRANSACCION'] or '—', C)
        filas.append([_p(r['Número Factura'], C), _p(r['Razon Social'], C), _p(bs(r['Monto']), ParagraphStyle('r', parent=C, alignment=TA_RIGHT)),
                      idc, _p(r['TITULAR ORIGEN'] or '—', C), _p(r['estado_conciliacion'], C)])
        if r['estado_conciliacion'] == 'CONFIRMADA POR CAJERO': est.append(('BACKGROUND', (0, i), (-1, i), AMBAR))
    ta = Table(filas, colWidths=[20 * mm, 44 * mm, 22 * mm, 26 * mm, 40 * mm, 28 * mm], repeatRows=1); ta.setStyle(TableStyle(est)); S.append(ta)
    conf = qr[qr['estado_conciliacion'] == 'CONFIRMADA POR CAJERO']
    S.append(_p('Confirmaciones manuales del cajero', H2))
    if conf.empty: S.append(_p('Ninguna.', C))
    else:
        for _, r in conf.iterrows():
            how = {'ELECCIÓN CAJERO': 'asignación elegida por el cajero', 'RESPALDO MANUAL': 'respaldo manual (ID de transacción indicado por el cajero)'}.get(
                r['nivel'], 'coincidencia parcial confirmada por el cajero')
            S.append(_p(f"Factura {r['Número Factura']} (Bs {bs(r['Monto'])}) ↔ ID {r['ID TRANSACCION']} · titular banco «{r['TITULAR ORIGEN']}» · cliente «{r['Razon Social']}» / estudiante «{r['Nombre Estudiante']}» · {how}.".replace('↔', '-'), C))
    notas = info.get('notas_descartes') or []
    if notas:
        S.append(_p('Coincidencias sugeridas descartadas por el cajero', H2))
        for n in notas: S.append(_p(n['texto'], C))
    if res.get('ci'):
        S.append(_p('Comunicaciones internas (depósitos a otras cuentas, prevalidados por Ingresos)', H2))
        fi = [[_p(h, CB) for h in ('N° factura', 'Cliente', 'Monto (Bs)', 'Banco', 'Código de asignación')]]
        for r in res['ci']: fi.append([_p(r['factura'], C), _p(r['razon_social'], C), _p(bs(r['monto']), ParagraphStyle('r2', parent=C, alignment=TA_RIGHT)), _p(r['banco'], C),
                                _pa(r['codigo'], C) if (str(r['factura']), r['banco']) in marcadas else _p(r['codigo'], C)])
        x = Table(fi, colWidths=[22 * mm, 66 * mm, 24 * mm, 22 * mm, 40 * mm], repeatRows=1)
        x.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (-1, 0), FONDO)])); S.append(x)
        desde = [str(r['factura']) for r in res['ci'] if str(r['factura']) in set(map(str, info.get('plantilla_origen') or []))]
        if desde: S.append(_p('Cargadas desde la plantilla de comunicaciones internas: facturas ' + ', '.join(desde) + '.', NOTA))
    if alertas:
        S.append(_p('Alertas de revisión: formato del código de asignación', H2))
        S.append(_p('Estos códigos se registraron igual, pero no cumplen el formato esperado del banco; Ingresos debe revisarlos.', NOTA))
        fl = [[_p(h, CB) for h in ('N° factura', 'Banco', 'Código', 'Regla esperada')]]
        for a in alertas: fl.append([_p(a['factura'], C), _p(a['banco'], C), _pa(a['codigo'], C), _p(a['detalle'], C)])
        x = Table(fl, colWidths=[22 * mm, 24 * mm, 46 * mm, 82 * mm], repeatRows=1)
        x.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (-1, 0), FONDO),
                               ('BACKGROUND', (0, 1), (-1, -1), AMBAR)])); S.append(x)
    S.append(_p('Facturas anuladas (excluidas)', H2))
    if anul.empty: S.append(_p('Ninguna.', C))
    else:
        fa = [[_p(h, CB) for h in ('N° factura', 'Cliente', 'Monto (Bs)', 'Estado')]]
        for _, r in anul.iterrows(): fa.append([_p(r['Número Factura'], C), _p(r['Razon Social'], C), _p(bs(r['Monto']), C), _p(r['estado_conciliacion'], C)])
        x = Table(fa, colWidths=[25 * mm, 80 * mm, 30 * mm, 45 * mm], repeatRows=1)
        x.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (-1, 0), FONDO)])); S.append(x)
    tj = res['tarjetas']
    S.append(_p('Facturas con tarjeta (según sistema)', H2))
    if tj.empty: S.append(_p('Ninguna.', C))
    else:
        ft = [[_p(h, CB) for h in ('N° factura', 'Cliente', 'Monto (Bs)')]]
        for _, r in tj.iterrows(): ft.append([_p(r['Número Factura'], C), _p(r['Razon Social'], C), _p(bs(r['Monto']), C)])
        x = Table(ft, colWidths=[25 * mm, 110 * mm, 30 * mm], repeatRows=1)
        x.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.5, LINEA), ('INNERGRID', (0, 0), (-1, -1), 0.25, LINEA), ('BACKGROUND', (0, 0), (-1, 0), FONDO)])); S.append(x)

    def pie(canvas, d):
        canvas.saveState(); canvas.setFillColor(GRIS)
        est_r = info.get('estado_revision') or 'PENDIENTE DE REVISIÓN'
        if est_r == 'CONFORME' and info.get('revisado_por'):
            txt = f"Revisado y conforme por {info.get('revisado_por_nombre') or info['revisado_por']} el {fecha_hora(info['revisado_en'])}."
        elif est_r == 'RECHAZADO':
            txt = f"Rechazado por Ingresos: {info.get('observacion_revision') or '-'}"
        else:
            txt = 'Comprobante del cajero. Declaración del cajero; no constituye aprobación hasta la revisión de Ingresos.'
        pf = _p(txt, ParagraphStyle('pie', parent=N, fontSize=7.5, leading=9, textColor=GRIS))
        w, h = pf.wrap(A4[0] - 32 * mm, 12 * mm); pf.drawOn(canvas, 16 * mm, 12 * mm)
        canvas.setStrokeColor(LINEA); canvas.line(16 * mm, 12 * mm + h + 1.5 * mm, A4[0] - 16 * mm, 12 * mm + h + 1.5 * mm)
        canvas.setFont('Helvetica', 7)
        canvas.drawString(16 * mm, 8 * mm, f"{info['correlativo']} · v{info['version']} · generado {fecha_hora(info['presentado_en'])}")
        canvas.drawRightString(A4[0] - 16 * mm, 8 * mm, f"Página {d.page}"); canvas.restoreState()
    doc.build(S, onFirstPage=pie, onLaterPages=pie)


def fecha_cierre_hora(x):
    return fecha_hora(x).replace(' ', ' a las ')
