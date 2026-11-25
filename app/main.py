import json, logging, os, re, secrets, tempfile
from datetime import datetime, timedelta
from urllib.parse import quote

from fastapi import FastAPI, Form, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.middleware.sessions import SessionMiddleware

from . import auth, catalogo, formato, plantilla as plantilla_xl, servicio
from .servicio import Almacen, ErrorUsuario, MAX_BYTES

log = logging.getLogger('cierre')
AQUI = os.path.dirname(os.path.abspath(__file__))


class NoAutenticado(Exception):
    pass


def create_app():
    app = FastAPI(title='Cierre Diario de Caja', docs_url=None, redoc_url=None, openapi_url=None)
    secret = os.environ.get('SECRET_KEY')
    if not secret:
        secret = secrets.token_hex(32); log.warning('SECRET_KEY no definido: se usa una clave temporal (las sesiones se pierden al reiniciar).')
    cs = os.environ.get('COOKIE_SECURE', 'auto').lower()
    seguro = cs in ('1', 'true', 'si') or (cs == 'auto' and bool(os.environ.get('RAILWAY_ENVIRONMENT') or os.environ.get('RAILWAY_PROJECT_ID')))
    app.add_middleware(SessionMiddleware, secret_key=secret, session_cookie='cierre_sesion', max_age=8 * 3600,
                       same_site='lax', https_only=seguro)
    app.state.almacen = Almacen(os.environ.get('DATA_DIR', './data'))
    app.state.almacen.sembrar_usuarios(auth.cargar_usuarios(os.environ.get('APP_USERS')))   # solo crea los que faltan (o los marcados forzar)
    app.state.limite = auth.LimiteIntentos()
    app.mount('/static', StaticFiles(directory=os.path.join(AQUI, 'static')), name='static')
    T = Jinja2Templates(directory=os.path.join(AQUI, 'templates'))
    T.env.filters.update(bs=formato.bs, fecha=formato.fecha, fecha_hora=formato.fecha_hora, hora=formato.hora)
    T.env.filters['rol_txt'] = lambda r: {'superadmin': 'Superadmin', 'ingresos': 'Ingresos', 'cajero': 'Cajero'}.get(r, r)
    T.env.globals.update(ROLES=auth.ROLES, MIN_CLAVE=auth.MIN_CLAVE, ESTADOS=servicio.ESTADOS_REVISION, BANCOS=catalogo.LISTA_BANCOS, BANCOS_UI=[(b, catalogo.etiqueta(b)) for b in catalogo.LISTA_BANCOS], BANCO_QR=catalogo.BANCO_QR, REGLAS={b: catalogo.regla_formato(b) for b in catalogo.LISTA_BANCOS}, MIN_OBS=servicio.MIN_OBSERVACION)
    def nombre_completo(login):
        """Nombre completo del usuario (tabla usuarios) para mostrar en pantalla; si ya no existe, el usuario tal cual."""
        if not login: return login
        r = alm.usuario_get(login)
        return r['nombre'] if r is not None and r['nombre'] else login
    T.env.filters['nombre_completo'] = nombre_completo
    T.env.filters['nombre_archivo'] = lambda n: re.sub(r'^[0-9a-f]{8}-', '', str(n or ''))   # solo en pantalla: quita el prefijo hexadecimal de origen
    T.env.filters['clase_estado'] = lambda e: {'PENDIENTE DE REVISIÓN': 'ins-ambar', 'CONFORME': 'ins-verde', 'RECHAZADO': 'ins-rojo'}.get(e, 'ins-gris')
    alm: Almacen = app.state.almacen

    def render(request, nombre, status=200, **ctx):
        ctx.update(user=request.session.get('user'), csrf=csrf_token(request), flash=request.session.pop('flash', None))
        return T.TemplateResponse(request, nombre, ctx, status_code=status)

    def csrf_token(request):
        if 'csrf' not in request.session: request.session['csrf'] = secrets.token_urlsafe(24)
        return request.session['csrf']

    def flash(request, msg, tipo='info'): request.session['flash'] = {'msg': msg, 'tipo': tipo}

    def usuario(request):
        """Usuario de la sesion, revalidado contra la base: un usuario desactivado o con otro rol lo nota de inmediato."""
        u = request.session.get('user')
        if not u: raise NoAutenticado()
        r = alm.usuario_get(u.get('usuario'))
        if r is None or not r['activo']:
            request.session.clear(); raise NoAutenticado()
        u = {'usuario': r['usuario'], 'nombre': r['nombre'], 'rol': r['rol']}
        request.session['user'] = u
        return u

    def verificar_csrf(request, token):
        if not token or not secrets.compare_digest(token, request.session.get('csrf', '')):
            raise StarletteHTTPException(403, 'Sesión de formulario vencida. Recargue la página e intente de nuevo.')

    def cajero(request):   # crear/corregir cierres: cajero y superadmin
        u = usuario(request)
        if u['rol'] not in ('cajero', 'superadmin'): raise StarletteHTTPException(403, 'Esta pantalla es solo para cajeros.')
        return u

    def ingresos(request):   # revisar cierres: ingresos y superadmin
        u = usuario(request)
        if u['rol'] not in ('ingresos', 'superadmin'): raise StarletteHTTPException(403, 'Esta acción es solo para el personal de Ingresos.')
        return u

    def superadmin(request):
        u = usuario(request)
        if u['rol'] != 'superadmin': raise StarletteHTTPException(403, 'Esta pantalla es solo para el superadministrador.')
        return u

    @app.exception_handler(NoAutenticado)
    async def _no_auth(request, exc):
        return RedirectResponse('/login', status_code=303)

    @app.exception_handler(StarletteHTTPException)
    async def _http(request, exc):
        msgs = {404: 'No se encontró lo que busca.', 403: 'No tiene permiso para ver esto.'}
        return render(request, 'error.html', exc.status_code, codigo=exc.status_code, mensaje=exc.detail if exc.status_code != 404 else msgs[404])

    # ---------------- salud / login ----------------
    @app.get('/salud')
    def salud(): return {'estado': 'ok'}

    @app.get('/login', response_class=HTMLResponse)
    def login_form(request: Request):
        if request.session.get('user'): return RedirectResponse('/', 303)
        return render(request, 'login.html', configurado=alm.hay_usuarios_activos())

    @app.post('/login')
    def login(request: Request, usuario_: str = Form('', alias='usuario'), clave: str = Form(''), csrf: str = Form('')):
        if not alm.hay_usuarios_activos(): return render(request, 'login.html', 200, configurado=False)
        verificar_csrf(request, csrf)
        u = usuario_.strip().lower()
        if app.state.limite.bloqueado(u):
            alm.auditar(u, 'login_bloqueado')
            return render(request, 'login.html', 429, configurado=True, error='Demasiados intentos. Espere unos minutos e intente de nuevo.')
        d = auth.autenticar(alm.usuario_get, u, clave)
        if not d:
            app.state.limite.registrar_fallo(u); alm.auditar(u, 'login_fallido')
            return render(request, 'login.html', 401, configurado=True, error='Usuario o clave incorrectos.')
        app.state.limite.limpiar(u); request.session.clear()
        request.session['user'] = {'usuario': d['usuario'], 'nombre': d['nombre'], 'rol': d['rol']}
        alm.auditar(u, 'login_ok')
        return RedirectResponse('/', 303)

    @app.post('/logout')
    def logout(request: Request, csrf: str = Form('')):
        u = request.session.get('user')
        if u: verificar_csrf(request, csrf); alm.auditar(u['usuario'], 'logout')
        request.session.clear()
        return RedirectResponse('/login', 303)

    # ---------------- mis cierres ----------------
    @app.get('/', response_class=HTMLResponse)
    def inicio(request: Request):
        u = usuario(request); q = request.query_params
        estado = q.get('estado', '') if q.get('estado', '') in servicio.ESTADOS_REVISION else ''
        def f_ok(x):
            try: datetime.strptime(x, '%Y-%m-%d'); return x
            except ValueError: return ''
        desde, hasta = f_ok(q.get('desde', '')), f_ok(q.get('hasta', ''))
        es_ing = u['rol'] in ('ingresos', 'superadmin')
        con_al = es_ing and q.get('alertas', '') == '1'
        filas = alm.cierres_de(None if es_ing else u['usuario'], estado or None, desde or None, hasta or None, con_al)
        return render(request, 'mis_cierres.html', cierres=filas, todos=es_ing, f_estado=estado, f_desde=desde, f_hasta=hasta,
                      f_alertas=con_al, filtrado=bool(estado or desde or hasta or con_al))

    # ---------------- nuevo cierre ----------------
    @app.get('/plantilla-comunicaciones-internas')
    def descargar_plantilla(request: Request):
        u = usuario(request)
        if not os.path.isfile(plantilla_xl.RUTA_ESTATICA): plantilla_xl.generar_plantilla(plantilla_xl.RUTA_ESTATICA)
        alm.auditar(u['usuario'], 'descarga_plantilla', {})
        return FileResponse(plantilla_xl.RUTA_ESTATICA, media_type=servicio.ARCHIVOS_CIERRE['conciliacion.xlsx'], filename=plantilla_xl.NOMBRE_DESCARGA,
                            content_disposition_type='attachment')

    @app.get('/nuevo', response_class=HTMLResponse)
    def nuevo(request: Request):
        cajero(request); return render(request, 'nuevo.html', max_mb=MAX_BYTES // 1048576)

    async def guardar_subida(archivo: UploadFile, tmpdir, etiqueta, permitidas):
        nombre = os.path.basename(archivo.filename or '')
        if not nombre: raise ErrorUsuario(f'Falta el archivo: {etiqueta}.')
        ext = nombre.rsplit('.', 1)[-1].lower() if '.' in nombre else ''
        if ext not in permitidas: raise ErrorUsuario(f'{etiqueta}: el archivo debe ser {" o ".join("." + e for e in permitidas)}.')
        ruta = os.path.join(tmpdir, etiqueta.split()[0] + '_' + secrets.token_hex(4) + '.' + ext); n = 0
        with open(ruta, 'wb') as f:
            while chunk := await archivo.read(1 << 16):
                n += len(chunk)
                if n > MAX_BYTES: raise ErrorUsuario(f'{etiqueta}: el archivo supera el máximo de {MAX_BYTES // 1048576} MB.')
                f.write(chunk)
        if n == 0: raise ErrorUsuario(f'{etiqueta}: el archivo está vacío.')
        return ruta, nombre

    def aplicar_y_avisar(request, u, bid, est, resumen, nombre, extra=''):
        """Aplica la plantilla del borrador (si `resumen` es None) y la deja guardada; auditoria y aviso con el resumen."""
        if resumen is None:
            est, resumen = alm.aplicar_plantilla(bid, est, u['usuario']); alm.guardar_estado(bid, est)
        d = alm.ruta_borrador(bid)
        alm.auditar(u['usuario'], 'plantilla_aplicada', {'borrador': bid, 'archivo': nombre, 'sha256': servicio.sha256(os.path.join(d, servicio.ARCH_PLANTILLA)),
                                                         'leidas': resumen['leidas'], 'aplicadas': resumen['aplicadas'], 'con_problemas': resumen['con_problemas']})
        n, a, p = resumen['leidas'], resumen['aplicadas'], resumen['con_problemas']
        flash(request, f"Plantilla leída: {n} {'fila' if n == 1 else 'filas'} · {a} {'aplicada' if a == 1 else 'aplicadas'} · {p} con {'problema' if p == 1 else 'problemas'}."
                       + (' Revise el cuadro «Filas de la plantilla con problemas».' if p else '') + (' ' + extra if extra else ''), 'ambar' if p else 'ok')

    @app.post('/nuevo')
    async def nuevo_post(request: Request, movimiento: UploadFile = File(None), banco: UploadFile = File(None),
                         plantilla: UploadFile = File(None), tarjeta: str = Form(''), csrf: str = Form('')):
        u = cajero(request); verificar_csrf(request, csrf)
        with tempfile.TemporaryDirectory() as tmp:
            try:
                if movimiento is None or banco is None: raise ErrorUsuario('Seleccione los dos archivos.')
                try: tj = formato.parsear_monto(tarjeta)
                except ValueError: raise ErrorUsuario('Escriba el total cobrado con tarjetas en Bs (use 0 si no hubo). Ejemplo: 1191 o 1,191.50')
                p1, n1 = await guardar_subida(movimiento, tmp, 'Movimiento del cajero', ('xls', 'xlsx', 'pdf'))
                p2, n2 = await guardar_subida(banco, tmp, 'Reporte QR BANECO', ('xlsx',))
                ext, _ = servicio.validar_a1(p1); servicio.validar_a2(p2)
                p3 = n3 = None
                if plantilla is not None and plantilla.filename:
                    p3, n3 = await guardar_subida(plantilla, tmp, 'Plantilla de comunicaciones internas', ('xlsx',)); servicio.validar_plantilla(p3)
                from . import motor
                try: r = motor.conciliar(p1, p2, tj)
                except Exception: raise ErrorUsuario('No se pudo conciliar: ' + servicio.MSG_A1)
                if r['totales']['fecha_cierre'] is None: raise ErrorUsuario('El movimiento del cajero no tiene facturas válidas.')
                fc = str(r['totales']['fecha_cierre']); bloqueo = servicio.regla_vigente(alm.vigente(fc, u['usuario']), None, fc)
                if bloqueo: raise ErrorUsuario(bloqueo)
                bid = alm.crear_borrador(u['usuario'], p1, ext, p2, tj, {'movimiento': n1, 'banco': n2, **({'plantilla': n3} if n3 else {})}, p3)
            except ErrorUsuario as e:
                return render(request, 'nuevo.html', 400, max_mb=MAX_BYTES // 1048576, error=str(e), tarjeta=tarjeta)
        alm.auditar(u['usuario'], 'borrador_creado', {'borrador': bid, 'tarjeta': tj, 'archivos': {'movimiento': n1, 'banco': n2, **({'plantilla': n3} if n3 else {})}})
        if n3: aplicar_y_avisar(request, u, bid, alm.estado(bid), None, n3)
        return RedirectResponse(f'/borrador/{bid}', 303)

    # ---------------- borrador ----------------
    def cargar_borrador(request, bid):
        u = usuario(request); est = alm.estado(bid)
        if est is None or est['usuario'] != u['usuario']:
            if est is not None: alm.auditar(u['usuario'], 'acceso_denegado', {'borrador': bid})
            raise StarletteHTTPException(404, 'No se encontró el borrador.')
        return u, est

    def vista(bid, est, usr):
        res = alm.conciliar(bid, est); ev = servicio.evaluar(res); t = res['totales']; fc = t['fecha_cierre']
        if fc:
            conf = alm.conflictos_para(res, usr['usuario'])
            if conf:
                c = conf[0]; ev['presentable'] = False; ev['nivel'] = 'rojo'
                ev['razones'].append(f"La factura {c['factura']}: " + servicio.msg_qr_usado(c['id_transaccion'], c))
            for c in alm.conflictos_ci_para(res, usr['usuario']):
                ev['presentable'] = False; ev['nivel'] = 'rojo'
                ev['razones'].append(f"La comunicación interna de la factura {c['factura']} ({c['banco']} {c['codigo']}): "
                                     + servicio.msg_ci_usada(c['banco'], catalogo.normalizar_codigo(c['codigo']), c))
            bloqueo = servicio.regla_vigente(alm.vigente(str(fc), usr['usuario']), est.get('origen'), str(fc))
            if bloqueo: ev['presentable'] = False; ev['nivel'] = 'rojo'; ev['razones'].append(bloqueo)
        f = res['facturas']; b = res['banco']; qr = f[f['Tipo (motor)'].isin(['QR', 'C.I.'])]
        ci = {r['factura']: r for r in res.get('ci', [])}
        af = servicio.alertas_revision(res, est)
        af_por_nf = {a['factura']: a for a in af if a['tipo'] == servicio.TIPO_FORMATO}; ad_por_nf = {a['factura']: a for a in af if a['tipo'] == catalogo.TIPO_FECHA}
        desde_pl = set(est.get('plantilla_origen') or [])
        pl = est.get('plantilla'); pl_vista = None
        if pl:
            en_ci = {str(r['factura']) for r in res.get('ci', [])}
            vivos = [p for p in (pl.get('problemas') or []) if p['factura'] not in en_ci]    # los que el cajero ya resolvio a mano no se listan
            pl_vista = {**pl, 'problemas_vivos': vivos, 'resueltos': len(pl.get('problemas') or []) - len(vivos), 'descartadas_vivas': len(est.get('plantilla_descartes') or [])}
        desc_por_nf = {}
        for d in res.get('descartes', []): desc_por_nf.setdefault(d['factura'], []).append(d)
        filas = []
        for _, r in qr.iterrows():
            nf = r['Número Factura']; e = r['estado_conciliacion']; nivel = r['nivel']
            obs = e in servicio.OBSERVADAS
            filas.append({'nf': nf, 'cliente': r['Razon Social'], 'estudiante': r['Nombre Estudiante'], 'monto': r['Monto'], 'hora': r['hora_factura'],
                          'estado': e, 'nivel': nivel, 'id': r['ID TRANSACCION'], 'titular': r['TITULAR ORIGEN'], 'hora_banco': r['hora_banco'],
                          'observada': obs, 'manual': nivel == 'RESPALDO MANUAL', 'ci': ci.get(nf), 'alerta_fmt': af_por_nf.get(str(nf)), 'alerta_fecha': ad_por_nf.get(str(nf)),
                          'deshacer': e == 'CONFIRMADA POR CAJERO' or e == 'COMUNICACIÓN INTERNA', 'desde_plantilla': str(nf) in desde_pl,
                          'candidatos': servicio.candidatos(res, nf) if e in ('SIN RESPALDO', 'AMBIGUO') else [],
                          'descartes': desc_por_nf.get(str(nf), [])})
        observadas = [x for x in filas if x['observada']]; conciliadas = [x for x in filas if not x['observada']]
        lista = servicio.conceptos(t)
        return dict(bid=bid, est=est, t=t, ev=ev, pl=pl_vista, observadas=observadas, conciliadas=conciliadas, tab_inicial='observadas' if observadas else 'conciliadas',
                    conceptos=lista, total=servicio.total_conceptos(lista), origen=est.get('origen'),
                    qr_conc=t['QR conciliado automático (CONCILIADA)'] + t['QR confirmado por cajero'],
                    ajenos=b[b['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO'].to_dict('records'),
                    anuladas=f[f['estado_conciliacion'].str.startswith('EXCLUIDA')].to_dict('records'),
                    tarjetas=res['tarjetas'].to_dict('records'), alertas_fmt=af,
                    por_confirmar=[x for x in observadas if x['estado'] == 'POR CONFIRMAR'],
                    # las alertas INFO (anuladas, otra fecha) no se muestran al cajero: siguen en cierre.json y en la hoja Alertas del xlsx
                    alertas=[({'nivel': 'ALERTA', 'detalle': f"Factura {a['factura']}: el código {a['codigo']} de {a['banco']} no cumple el formato esperado ({a['detalle']}). "
                                                             f"Si lo escribió mal, deshaga el respaldo y regístrelo de nuevo; si es correcto, puede presentar igual."}
                              if a['tipo'] == servicio.TIPO_FORMATO else
                              {'nivel': 'ALERTA', 'detalle': f"Factura {a['factura']} ({a['banco']} {a['codigo']}): {a['detalle']} Si la tiene, deshaga el respaldo y regístrelo de nuevo con la fecha; si no, puede presentar igual."})
                             for a in af]
                    + [a for a in res['alertas'].to_dict('records') if a['nivel'] != 'INFO'])

    @app.get('/borrador/{bid}', response_class=HTMLResponse)
    def ver_borrador(request: Request, bid: str):
        u, est = cargar_borrador(request, bid)
        if est.get('presentado'):
            raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        return render(request, 'borrador.html', **vista(bid, est, u))

    def estado_de(est, nf):
        return {'confirmada': nf in est['confirmaciones'], 'eleccion': est['elecciones'].get(nf), 'respaldo_manual': nf in est.get('respaldos_manuales', []),
                'reclasificacion': (est.get('reclasificaciones') or {}).get(nf), 'rechazos': (est.get('rechazos') or {}).get(nf)}

    def accion(nombre, aplicar):
        def h(request: Request, bid: str, factura: str = Form(''), id_transaccion: str = Form(''), csrf: str = Form('')):
            u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
            if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
            antes = estado_de(est, factura.strip())
            aplicar(est, factura.strip(), id_transaccion.strip())
            alm.guardar_estado(bid, est)
            res = alm.conciliar(bid, est)
            alm.auditar(u['usuario'], nombre, {'borrador': bid, 'factura': factura, 'id_transaccion': id_transaccion, 'estado_qr': res['totales']['Estado QR'],
                                               'anterior': antes})
            avisos = [a['detalle'] for a in res['alertas'].to_dict('records') if 'no aplicada' in a['detalle']]
            if avisos: flash(request, ' '.join(avisos), 'error')
            return RedirectResponse(f'/borrador/{bid}', 303)
        return h

    def _conf(est, nf, _):
        if nf and nf not in est['confirmaciones']: est['confirmaciones'].append(nf)
    def _des(est, nf, _):
        est['confirmaciones'] = [x for x in est['confirmaciones'] if x != nf]; est['elecciones'].pop(nf, None)
        est['respaldos_manuales'] = [x for x in est.get('respaldos_manuales', []) if x != nf]; (est.get('reclasificaciones') or {}).pop(nf, None)
        if nf in (est.get('plantilla_origen') or []):    # si venia de la plantilla, no se vuelve a aplicar sola al recalcular
            est['plantilla_origen'] = [x for x in est['plantilla_origen'] if x != nf]
            est.setdefault('plantilla_descartes', [])
            if nf not in est['plantilla_descartes']: est['plantilla_descartes'].append(nf)
    def _asig(est, nf, idt):
        if nf and idt: est['elecciones'][nf] = idt
    @app.post('/borrador/{bid}/confirmar-todas')
    def confirmar_todas(request: Request, bid: str, facturas: list[str] = Form([]), csrf: str = Form('')):
        """Confirma de una vez las coincidencias dudosas (solo POR CONFIRMAR) listadas en el diálogo; una auditoría por factura."""
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        pedidas = list(dict.fromkeys(x.strip() for x in facturas if x.strip()))
        res = alm.conciliar(bid, est); f = res['facturas'].set_index('Número Factura')
        validas = [nf for nf in pedidas if nf in f.index and f.loc[nf, 'estado_conciliacion'] == 'POR CONFIRMAR']
        if not validas:
            flash(request, 'No hay coincidencias dudosas para confirmar (ya cambiaron de estado). Revise la lista.', 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        nuevo = json.loads(json.dumps(est))
        for nf in validas:
            if nf not in nuevo['confirmaciones']: nuevo['confirmaciones'].append(nf)
        r2 = alm.conciliar(bid, nuevo); f2 = r2['facturas'].set_index('Número Factura')
        ok = [nf for nf in validas if f2.loc[nf, 'estado_conciliacion'] == 'CONFIRMADA POR CAJERO']
        nuevo['confirmaciones'] = [x for x in nuevo['confirmaciones'] if x not in validas or x in ok]
        alm.guardar_estado(bid, nuevo)
        for nf in ok:
            alm.auditar(u['usuario'], 'confirmar_coincidencia', {'borrador': bid, 'factura': nf, 'id_transaccion': str(f.loc[nf, 'ID TRANSACCION']), 'lote': True,
                                                                 'estado_qr': r2['totales']['Estado QR'], 'anterior': estado_de(est, nf)})
        alm.auditar(u['usuario'], 'confirmar_todas', {'borrador': bid, 'facturas': ok, 'n': len(ok)})
        flash(request, f"Se confirmaron {len(ok)} coincidencias dudosas." if len(ok) != 1 else 'Se confirmó 1 coincidencia dudosa.', 'ok')
        return RedirectResponse(f'/borrador/{bid}', 303)

    app.post('/borrador/{bid}/confirmar')(accion('confirmar_coincidencia', _conf))
    app.post('/borrador/{bid}/deshacer')(accion('deshacer_confirmacion', _des))
    app.post('/borrador/{bid}/asignar')(accion('asignar_movimiento', _asig))

    def _descarte(nombre, aplicar):
        def h(request: Request, bid: str, factura: str = Form(''), csrf: str = Form('')):
            u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
            if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
            nf = factura.strip()
            try: nuevo, info = aplicar(bid, est, nf)
            except ErrorUsuario as e:
                flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
            alm.guardar_estado(bid, nuevo)
            res = alm.conciliar(bid, nuevo)
            alm.auditar(u['usuario'], nombre, {'borrador': bid, 'factura': nf, **info, 'estado_qr': res['totales']['Estado QR']})
            if nombre == 'descartar_coincidencia':
                flash(request, f"Descartó el pago {info['id_transaccion']} de {info['titular']} para la factura {nf}. Ahora puede elegir otro pago o reclasificarla.", 'ok')
            else: flash(request, f"Se deshizo el descarte de la factura {nf}.", 'ok')
            return RedirectResponse(f'/borrador/{bid}', 303)
        return h

    def _desc(bid, est, nf): return alm.descartar_coincidencia(bid, est, nf)
    def _undesc(bid, est, nf):
        nuevo, ids = alm.deshacer_descarte(bid, est, nf); return nuevo, {'ids_transaccion': ids}
    app.post('/borrador/{bid}/descartar')(_descarte('descartar_coincidencia', _desc))
    app.post('/borrador/{bid}/deshacer-descarte')(_descarte('deshacer_descarte', _undesc))

    @app.post('/borrador/{bid}/respaldo')
    def respaldo_manual(request: Request, bid: str, factura: str = Form(''), banco: str = Form(''), codigo: str = Form(''), fecha_deposito: str = Form(''), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        nf = factura.strip(); antes = estado_de(est, nf)
        try:
            nuevo, info = alm.registrar_respaldo(bid, est, u['usuario'], nf, banco, codigo, fecha_deposito)
        except ErrorUsuario as e:
            alm.auditar(u['usuario'], 'respaldo_manual_rechazado', {'borrador': bid, 'factura': nf, 'banco': banco, 'codigo': codigo.strip(), 'fecha_deposito': fecha_deposito.strip(), 'motivo': str(e), 'anterior': antes})
            flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        alm.guardar_estado(bid, nuevo)
        alm.auditar(u['usuario'], 'respaldo_manual', {'borrador': bid, 'factura': nf, **info, 'anterior': antes})
        msg = f"Factura {nf}: respaldo manual registrado ({info['banco']} {info['codigo']})."
        regla = catalogo.formato_incumplido(info['banco'], info['codigo'])
        fecha_al = None if catalogo.es_qr(info['banco']) else catalogo.alerta_fecha_deposito(info['banco'], info.get('fecha_deposito'), alm.conciliar(bid, nuevo)['totales']['fecha_cierre'])
        if info.get('fecha_deposito'): msg += f" Fecha del depósito: {info['fecha_deposito']}."
        if regla:
            alm.auditar(u['usuario'], 'alerta_formato', {'borrador': bid, 'factura': nf, 'banco': info['banco'], 'codigo': info['codigo'], 'regla': regla})
            msg += f" FORMATO A REVISAR: se registró igual, pero no cumple la regla esperada ({regla}). Ingresos será notificado al presentar."
        if fecha_al:
            alm.auditar(u['usuario'], 'alerta_fecha_deposito', {'borrador': bid, 'factura': nf, 'banco': info['banco'], 'codigo': info['codigo'], 'detalle': fecha_al})
            msg += f" {fecha_al} Se registró igual; Ingresos será notificado al presentar."
        flash(request, msg, 'ambar' if (regla or fecha_al) else 'ok')
        return RedirectResponse(f'/borrador/{bid}', 303)

    @app.post('/borrador/{bid}/tarjeta')
    def cambiar_tarjeta(request: Request, bid: str, tarjeta: str = Form(''), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        try: tj = formato.parsear_monto(tarjeta)
        except ValueError:
            flash(request, 'Escriba el total cobrado con tarjetas en Bs (use 0 si no hubo). Ejemplo: 1191 o 1,191.50', 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        anterior = est['tarjeta']; est['tarjeta'] = tj; alm.guardar_estado(bid, est)
        alm.auditar(u['usuario'], 'tarjeta_corregida', {'borrador': bid, 'anterior': anterior, 'nuevo': tj})
        flash(request, f'Tarjeta declarada actualizada a Bs {formato.bs(tj)}.', 'ok')
        return RedirectResponse(f'/borrador/{bid}', 303)

    @app.post('/borrador/{bid}/plantilla')
    async def subir_plantilla(request: Request, bid: str, plantilla: UploadFile = File(None), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        with tempfile.TemporaryDirectory() as tmp:
            try:
                if plantilla is None or not plantilla.filename: raise ErrorUsuario('Seleccione el archivo de la plantilla (.xlsx).')
                p3, n3 = await guardar_subida(plantilla, tmp, 'Plantilla de comunicaciones internas', ('xlsx',)); servicio.validar_plantilla(p3)
                nuevo, resumen = alm.reemplazar_plantilla(bid, est, u['usuario'], p3, n3)
            except ErrorUsuario as e:
                flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        aplicar_y_avisar(request, u, bid, nuevo, resumen, n3)
        return RedirectResponse(f'/borrador/{bid}', 303)

    @app.post('/borrador/{bid}/archivos')
    async def reemplazar_archivos(request: Request, bid: str, movimiento: UploadFile = File(None), banco: UploadFile = File(None),
                                  plantilla: UploadFile = File(None), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        if not est.get('origen'): raise StarletteHTTPException(403, 'Solo se pueden reemplazar archivos al corregir un cierre rechazado.')
        with tempfile.TemporaryDirectory() as tmp:
            try:
                p1 = p2 = ext = n1 = n2 = None
                if movimiento is not None and movimiento.filename:
                    p1, n1 = await guardar_subida(movimiento, tmp, 'Movimiento del cajero', ('xls', 'xlsx', 'pdf')); ext, _ = servicio.validar_a1(p1)
                if banco is not None and banco.filename:
                    p2, n2 = await guardar_subida(banco, tmp, 'Reporte QR BANECO', ('xlsx',)); servicio.validar_a2(p2)
                p3 = n3 = None
                if plantilla is not None and plantilla.filename:
                    p3, n3 = await guardar_subida(plantilla, tmp, 'Plantilla de comunicaciones internas', ('xlsx',)); servicio.validar_plantilla(p3)
                if not p1 and not p2 and not p3: raise ErrorUsuario('Seleccione al menos un archivo para reemplazar.')
                from . import motor
                a1 = p1 or os.path.join(alm.ruta_borrador(bid), est['archivo_movimiento']); a2 = p2 or os.path.join(alm.ruta_borrador(bid), est['archivo_banco'])
                try: r = motor.conciliar(a1, a2, est['tarjeta'])
                except Exception: raise ErrorUsuario('No se pudo conciliar con esos archivos: ' + servicio.MSG_A1)
                if r['totales']['fecha_cierre'] is None: raise ErrorUsuario('El movimiento del cajero no tiene facturas válidas.')
                est = alm.reemplazar_archivos(bid, est, p1, ext, n1, p2, n2)
                resumen = None
                if p3: est, resumen = alm.reemplazar_plantilla(bid, est, u['usuario'], p3, n3)
                elif est.get('plantilla') and (p1 or p2):    # archivos nuevos: se revalida la plantilla ya cargada (sin reiniciar lo que el cajero deshizo)
                    est, resumen = alm.aplicar_plantilla(bid, est, u['usuario']); alm.guardar_estado(bid, est)
            except ErrorUsuario as e:
                flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        alm.auditar(u['usuario'], 'archivos_reemplazados', {'borrador': bid, 'movimiento': n1, 'banco': n2, 'plantilla': n3, 'origen': est['origen']['correlativo']})
        if resumen is not None: aplicar_y_avisar(request, u, bid, est, resumen, n3 or est['plantilla'].get('nombre'), 'Archivos reemplazados: revise de nuevo las facturas observadas y las decisiones guardadas.')
        else: flash(request, 'Archivos reemplazados. Revise de nuevo las facturas observadas y las decisiones guardadas.', 'ok')
        return RedirectResponse(f'/borrador/{bid}', 303)

    @app.post('/borrador/{bid}/presentar')
    def presentar(request: Request, bid: str, declaro: str = Form(''), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if declaro != '1':
            flash(request, 'Debe marcar la declaración para presentar el cierre.', 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        try:
            cid, corr = alm.presentar(bid, est, u['usuario'], u['nombre'])
        except ErrorUsuario as e:
            alm.auditar(u['usuario'], 'presentar_bloqueado', {'borrador': bid, 'motivo': str(e)})
            flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        request.session['presentado'] = cid
        return RedirectResponse(f'/cierres/{cid}', 303)

    # ---------------- detalle de cierre ----------------
    def pdf_vigente(c):
        """Nombre del PDF vigente: cierre_revision.pdf si existe (lleva el estado de revision), si no cierre.pdf."""
        d = os.path.join(alm.dir, c['ruta'])
        return 'cierre_revision.pdf' if os.path.isfile(os.path.join(d, 'cierre_revision.pdf')) else 'cierre.pdf'

    def hace_dias(presentado_en):
        try: n = (datetime.now().date() - datetime.fromisoformat(str(presentado_en)).date()).days
        except ValueError: return ''
        n = max(n, 0)
        return 'hoy' if n == 0 else ('hace 1 día' if n == 1 else f'hace {n} días')
    T.env.filters['hace_dias'] = hace_dias

    def cargar_cierre(request, cid):
        u = usuario(request); c = alm.cierre(cid)
        if c is None or (u['rol'] not in ('ingresos', 'superadmin') and c['usuario'] != u['usuario']):
            if c is not None: alm.auditar(u['usuario'], 'acceso_denegado', {'cierre': cid})
            raise StarletteHTTPException(404, 'No se encontró el cierre.')
        return u, c

    def q_presentado(request, cid): return request.session.pop('presentado', None) == cid

    @app.get('/cierres/{cid}', response_class=HTMLResponse)
    def detalle(request: Request, cid: int):
        u, c = cargar_cierre(request, cid)
        d = os.path.join(alm.dir, c['ruta'])
        archivos = [n for n in sorted(os.listdir(d))] if os.path.isdir(d) else []
        cj = None
        try:
            with open(os.path.join(d, 'cierre.json'), encoding='utf-8') as f: cj = json.load(f)
        except Exception: pass
        rem = alm.cierre(c['reemplaza_a']) if c['reemplaza_a'] else None
        posterior = None
        if c['estado_revision'] == servicio.ESTADO_REEMPLAZADO:
            con = alm.con()
            try: posterior = con.execute('SELECT id, correlativo FROM cierres WHERE reemplaza_a=?', (c['id'],)).fetchone()
            finally: con.close()
        conc = (cj or {}).get('conceptos', [])
        alertas_c = [dict(a) for a in alm.alertas_de(cid)]
        extracto = None
        if u['rol'] in ('ingresos', 'superadmin'):
            movs = alm.movimientos_cierre(cid); cis = alm.ci_cierre(cid)
            if movs or cis:
                t_m, n_m, n_c, tot_m = servicio.tsv_extracto(c['correlativo'], c['fecha_cierre'], movs, '', cis)
                extracto = {'n': n_m, 'n_ci': n_c, 'total': tot_m, 'marcar': t_m,
                            'liberar': servicio.tsv_extracto(c['correlativo'], c['fecha_cierre'], movs, 'LIBERAR', cis)[0] if u['rol'] == 'superadmin' else None}
        notas_desc = (cj or {}).get('notas_descartes', []) if u['rol'] in ('ingresos', 'superadmin') else []
        return render(request, 'cierre.html', notas_desc=notas_desc, c=c, pdf_vigente=pdf_vigente(c), recien=q_presentado(request, cid) and u['usuario'] == c['usuario'],
                      pendiente_hace=hace_dias(c['presentado_en']), archivos=archivos, cj=cj, anterior=rem, posterior=posterior, conceptos=conc, extracto=extracto,
                      total=servicio.total_conceptos(conc) if cj else None, ci=(cj or {}).get('comunicaciones_internas', []),
                      pl=(cj or {}).get('plantilla'), desde_pl=set((cj or {}).get('plantilla_origen') or []),
                      alertas_c=alertas_c, n_alertas_pend=sum(1 for a in alertas_c if not a['revisada_en']),
                      es_ingresos=u['rol'] in ('ingresos', 'superadmin'), es_super=u['rol'] == 'superadmin',
                      es_dueno=u['rol'] in ('cajero', 'superadmin') and c['usuario'] == u['usuario'],
                      historial=alm.historial(cid) if u['rol'] in ('ingresos', 'superadmin') else [])

    @app.post('/cierres/{cid}/revisar')
    def revisar(request: Request, cid: int, decision: str = Form(''), observacion: str = Form(''), csrf: str = Form('')):
        u = ingresos(request); verificar_csrf(request, csrf)
        c = alm.cierre(cid)
        if c is None: raise StarletteHTTPException(404, 'No se encontró el cierre.')
        d = {'aprobar': servicio.ESTADO_CONFORME, 'rechazar': servicio.ESTADO_RECHAZADO}.get(decision)
        try:
            if not d: raise ErrorUsuario('Decisión no válida.')
            alm.revisar(cid, u['usuario'], d, observacion)
        except ErrorUsuario as e:
            alm.auditar(u['usuario'], 'revision_rechazada_por_sistema', {'cierre': cid, 'decision': decision, 'motivo': str(e)})
            flash(request, str(e), 'error'); return RedirectResponse(f'/cierres/{cid}', 303)
        flash(request, f"Cierre {c['correlativo']} marcado como {d}.", 'ok')
        return RedirectResponse(f'/cierres/{cid}', 303)

    @app.post('/cierres/{cid}/alertas/{aid}/revisar')
    def revisar_alerta(request: Request, cid: int, aid: int, csrf: str = Form('')):
        u = ingresos(request); verificar_csrf(request, csrf)
        if alm.cierre(cid) is None: raise StarletteHTTPException(404, 'No se encontró el cierre.')
        try:
            a = alm.revisar_alerta(cid, aid, u['usuario'])
            flash(request, f"Alerta de la factura {a['factura']} ({a['banco']} {a['codigo']}) marcada como revisada.", 'ok')
        except ErrorUsuario as e:
            flash(request, str(e), 'error')
        return RedirectResponse(f'/cierres/{cid}#alertas-revision', 303)

    @app.post('/cierres/{cid}/eliminar')
    def eliminar(request: Request, cid: int, motivo: str = Form(''), correlativo: str = Form(''), csrf: str = Form('')):
        u = superadmin(request); verificar_csrf(request, csrf)
        if alm.cierre(cid) is None: raise StarletteHTTPException(404, 'No se encontró el cierre.')
        try: corr, existia = alm.eliminar(cid, u['usuario'], motivo, correlativo)
        except ErrorUsuario as e:
            flash(request, str(e), 'error'); return RedirectResponse(f'/cierres/{cid}', 303)
        flash(request, f"Cierre {corr} eliminado. Quedó registrado en auditoría." + ('' if existia else ' (La carpeta de archivos no existía; se eliminaron solo los registros.)'), 'ok')
        return RedirectResponse('/', 303)

    @app.post('/cierres/{cid}/revertir')
    def revertir(request: Request, cid: int, motivo: str = Form(''), csrf: str = Form('')):
        u = superadmin(request); verificar_csrf(request, csrf)
        c = alm.cierre(cid)
        if c is None: raise StarletteHTTPException(404, 'No se encontró el cierre.')
        try: alm.revertir(cid, u['usuario'], motivo)
        except ErrorUsuario as e:
            alm.auditar(u['usuario'], 'reversion_rechazada_por_sistema', {'cierre': cid, 'motivo': str(e)})
            flash(request, str(e), 'error'); return RedirectResponse(f'/cierres/{cid}', 303)
        flash(request, f"La revisión del cierre {c['correlativo']} se revirtió: vuelve a PENDIENTE DE REVISIÓN.", 'ok')
        return RedirectResponse(f'/cierres/{cid}', 303)

    @app.post('/cierres/{cid}/editar')
    def editar_reenviar(request: Request, cid: int, csrf: str = Form('')):
        u = cajero(request); verificar_csrf(request, csrf)
        _, c = cargar_cierre(request, cid)
        if c['estado_revision'] != servicio.ESTADO_RECHAZADO:
            alm.auditar(u['usuario'], 'editar_bloqueado', {'cierre': c['correlativo'], 'estado': c['estado_revision']})
            flash(request, f"Solo se puede editar un cierre RECHAZADO (este está {c['estado_revision']}).", 'error')
            return RedirectResponse(f'/cierres/{cid}', 303)
        vig = alm.vigente(c['fecha_cierre'], u['usuario'])
        if not vig or vig['id'] != c['id']:
            flash(request, 'Este cierre ya fue reemplazado.', 'error'); return RedirectResponse(f'/cierres/{cid}', 303)
        bid = alm.crear_borrador_desde(c, u['usuario'])
        alm.auditar(u['usuario'], 'editar_reenviar_iniciado', {'cierre': c['correlativo'], 'borrador': bid})
        return RedirectResponse(f'/borrador/{bid}', 303)

    @app.get('/cierres/{cid}/comprobante')
    def comprobante(request: Request, cid: int):
        u, c = cargar_cierre(request, cid)
        return RedirectResponse(f'/cierres/{cid}/archivo/{pdf_vigente(c)}', 303)

    @app.get('/cierres/{cid}/archivo/{nombre}')
    def descargar(request: Request, cid: int, nombre: str):
        u, c = cargar_cierre(request, cid)
        ok = nombre in servicio.ARCHIVOS_CIERRE or nombre in ('movimiento_original.xls', 'movimiento_original.xlsx', 'movimiento_original.pdf')
        ruta = os.path.join(alm.dir, c['ruta'], nombre)
        if not ok or not os.path.isfile(ruta): raise StarletteHTTPException(404, 'Archivo no encontrado.')
        alm.auditar(u['usuario'], 'descarga', {'cierre': c['correlativo'], 'archivo': nombre})
        mt = servicio.ARCHIVOS_CIERRE.get(nombre, 'application/vnd.ms-excel' if nombre.endswith('.xls') else 'application/pdf' if nombre.endswith('.pdf') else servicio.ARCHIVOS_CIERRE['conciliacion.xlsx'])
        return FileResponse(ruta, media_type=mt, filename=f"{c['correlativo']}_{nombre}",
                            content_disposition_type='inline' if nombre in ('cierre.pdf', 'cierre_revision.pdf') and request.query_params.get('descargar') != '1' else 'attachment')

    # ---------------- mi cuenta ----------------
    @app.get('/cuenta', response_class=HTMLResponse)
    def cuenta(request: Request):
        usuario(request); return render(request, 'cuenta.html')

    @app.post('/cuenta/clave')
    def cuenta_clave(request: Request, actual: str = Form(''), nueva: str = Form(''), nueva2: str = Form(''), csrf: str = Form('')):
        u = usuario(request); verificar_csrf(request, csrf)
        if app.state.limite.bloqueado(u['usuario']):
            alm.auditar(u['usuario'], 'cambio_clave_bloqueado')
            return render(request, 'cuenta.html', 429, error='Demasiados intentos. Espere unos minutos e intente de nuevo.')
        try: alm.cambiar_clave_propia(u['usuario'], actual, nueva, nueva2)
        except ErrorUsuario as e:
            if 'actual' in str(e): app.state.limite.registrar_fallo(u['usuario'])
            alm.auditar(u['usuario'], 'cambio_clave_fallido', {'motivo': str(e)})
            return render(request, 'cuenta.html', 400, error=str(e))
        app.state.limite.limpiar(u['usuario'])
        flash(request, 'Su clave se cambió correctamente.', 'ok')
        return RedirectResponse('/cuenta', 303)

    # ---------------- administracion (superadmin) ----------------
    @app.get('/admin/usuarios', response_class=HTMLResponse)
    def admin_usuarios(request: Request):
        u = superadmin(request)
        return render(request, 'admin_usuarios.html', usuarios=alm.usuarios(), yo=u['usuario'])

    def admin_accion(request, csrf, ejecutar, ok):
        u = superadmin(request); verificar_csrf(request, csrf)
        try: ejecutar(u['usuario']); flash(request, ok, 'ok')
        except ErrorUsuario as e: flash(request, str(e), 'error')
        return RedirectResponse('/admin/usuarios', 303)

    @app.post('/admin/usuarios/crear')
    def admin_crear(request: Request, usuario_: str = Form('', alias='usuario'), nombre: str = Form(''), rol: str = Form(''),
                    clave: str = Form(''), clave2: str = Form(''), csrf: str = Form('')):
        return admin_accion(request, csrf, lambda a: alm.usuario_crear(a, usuario_, nombre, rol, clave, clave2),
                            f'Usuario «{usuario_.strip().lower()}» creado.')

    @app.post('/admin/usuarios/{nombre_usuario}/rol')
    def admin_rol(request: Request, nombre_usuario: str, rol: str = Form(''), csrf: str = Form('')):
        return admin_accion(request, csrf, lambda a: alm.usuario_cambiar_rol(a, nombre_usuario, rol), f'Rol de «{nombre_usuario}» actualizado.')

    @app.post('/admin/usuarios/{nombre_usuario}/activo')
    def admin_activo(request: Request, nombre_usuario: str, activo: str = Form(''), csrf: str = Form('')):
        on = activo == '1'
        return admin_accion(request, csrf, lambda a: alm.usuario_set_activo(a, nombre_usuario, on),
                            f'Usuario «{nombre_usuario}» {"activado" if on else "desactivado"}.')

    @app.post('/admin/usuarios/{nombre_usuario}/clave')
    def admin_clave(request: Request, nombre_usuario: str, clave: str = Form(''), clave2: str = Form(''), csrf: str = Form('')):
        def hacer(a):
            alm.usuario_reset_clave(a, nombre_usuario, clave, clave2); app.state.limite.limpiar(nombre_usuario)
        return admin_accion(request, csrf, hacer, f'Clave de «{nombre_usuario}» restablecida.')

    @app.get('/admin/auditoria', response_class=HTMLResponse)
    def admin_auditoria(request: Request):
        superadmin(request); q = request.query_params
        def f_ok(x):
            try: datetime.strptime(x, '%Y-%m-%d'); return x
            except ValueError: return ''
        try: pag = int(q.get('pagina', '1'))
        except ValueError: pag = 1
        f = {'usuario': q.get('usuario', '').strip()[:60], 'accion': q.get('accion', '').strip()[:60], 'desde': f_ok(q.get('desde', '')), 'hasta': f_ok(q.get('hasta', ''))}
        r = alm.auditoria(f['usuario'] or None, f['accion'] or None, f['desde'] or None, f['hasta'] or None, pag)
        base = '&'.join(f'{k}={quote(v)}' for k, v in f.items() if v)
        return render(request, 'admin_auditoria.html', f=f, base=base, **r)

    return app


app = create_app()
