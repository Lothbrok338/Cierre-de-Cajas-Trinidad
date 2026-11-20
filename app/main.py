import logging, os, secrets, tempfile
from datetime import datetime
from urllib.parse import quote

from fastapi import FastAPI, Form, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.middleware.sessions import SessionMiddleware

from . import auth, catalogo, formato, servicio
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
    app.state.usuarios = auth.cargar_usuarios(os.environ.get('APP_USERS'))
    app.state.limite = auth.LimiteIntentos()
    app.mount('/static', StaticFiles(directory=os.path.join(AQUI, 'static')), name='static')
    T = Jinja2Templates(directory=os.path.join(AQUI, 'templates'))
    T.env.filters.update(bs=formato.bs, fecha=formato.fecha, fecha_hora=formato.fecha_hora, hora=formato.hora)
    T.env.globals.update(ESTADOS=servicio.ESTADOS_REVISION, BANCOS=catalogo.LISTA_BANCOS, BANCO_QR=catalogo.BANCO_QR, MIN_OBS=servicio.MIN_OBSERVACION)
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
        u = request.session.get('user')
        if not u: raise NoAutenticado()
        return u

    def verificar_csrf(request, token):
        if not token or not secrets.compare_digest(token, request.session.get('csrf', '')):
            raise StarletteHTTPException(403, 'Sesión de formulario vencida. Recargue la página e intente de nuevo.')

    def cajero(request):
        u = usuario(request)
        if u['rol'] != 'cajero': raise StarletteHTTPException(403, 'Esta pantalla es solo para cajeros.')
        return u

    def ingresos(request):
        u = usuario(request)
        if u['rol'] != 'ingresos': raise StarletteHTTPException(403, 'Esta acción es solo para el personal de Ingresos.')
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
        return render(request, 'login.html', configurado=bool(app.state.usuarios))

    @app.post('/login')
    def login(request: Request, usuario_: str = Form('', alias='usuario'), clave: str = Form(''), csrf: str = Form('')):
        if not app.state.usuarios: return render(request, 'login.html', 200, configurado=False)
        verificar_csrf(request, csrf)
        u = usuario_.strip()
        if app.state.limite.bloqueado(u):
            alm.auditar(u, 'login_bloqueado')
            return render(request, 'login.html', 429, configurado=True, error='Demasiados intentos. Espere unos minutos e intente de nuevo.')
        d = auth.autenticar(app.state.usuarios, u, clave)
        if not d:
            app.state.limite.registrar_fallo(u); alm.auditar(u, 'login_fallido')
            return render(request, 'login.html', 401, configurado=True, error='Usuario o clave incorrectos.')
        app.state.limite.limpiar(u); request.session.clear()
        request.session['user'] = {'usuario': u, 'nombre': d['nombre'], 'rol': d['rol']}
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
        es_ing = u['rol'] != 'cajero'
        filas = alm.cierres_de(None if es_ing else u['usuario'], estado or None, desde or None, hasta or None)
        return render(request, 'mis_cierres.html', cierres=filas, todos=es_ing, f_estado=estado, f_desde=desde, f_hasta=hasta,
                      filtrado=bool(estado or desde or hasta))

    # ---------------- nuevo cierre ----------------
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

    @app.post('/nuevo')
    async def nuevo_post(request: Request, movimiento: UploadFile = File(None), banco: UploadFile = File(None),
                         tarjeta: str = Form(''), csrf: str = Form('')):
        u = cajero(request); verificar_csrf(request, csrf)
        with tempfile.TemporaryDirectory() as tmp:
            try:
                if movimiento is None or banco is None: raise ErrorUsuario('Seleccione los dos archivos.')
                try: tj = formato.parsear_monto(tarjeta)
                except ValueError: raise ErrorUsuario('Escriba el total cobrado con tarjetas en Bs (use 0 si no hubo). Ejemplo: 1191 o 1,191.50')
                p1, n1 = await guardar_subida(movimiento, tmp, 'Movimiento del cajero', ('xls', 'xlsx'))
                p2, n2 = await guardar_subida(banco, tmp, 'Reporte QR BANECO', ('xlsx',))
                ext, _ = servicio.validar_a1(p1); servicio.validar_a2(p2)
                from . import motor
                try: r = motor.conciliar(p1, p2, tj)
                except Exception: raise ErrorUsuario('No se pudo conciliar: ' + servicio.MSG_A1)
                if r['totales']['fecha_cierre'] is None: raise ErrorUsuario('El movimiento del cajero no tiene facturas válidas.')
                fc = str(r['totales']['fecha_cierre']); bloqueo = servicio.regla_vigente(alm.vigente(fc, u['usuario']), None, fc)
                if bloqueo: raise ErrorUsuario(bloqueo)
                bid = alm.crear_borrador(u['usuario'], p1, ext, p2, tj, {'movimiento': n1, 'banco': n2})
            except ErrorUsuario as e:
                return render(request, 'nuevo.html', 400, max_mb=MAX_BYTES // 1048576, error=str(e), tarjeta=tarjeta)
        alm.auditar(u['usuario'], 'borrador_creado', {'borrador': bid, 'tarjeta': tj, 'archivos': {'movimiento': n1, 'banco': n2}})
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
                ev['razones'].append(f"La factura {c['factura']} usa el ID {c['id_transaccion']}, que ya está en el cierre {c['correlativo']} (cajero {c['usuario']}).")
            for c in alm.conflictos_ci_para(res, usr['usuario']):
                ev['presentable'] = False; ev['nivel'] = 'rojo'
                ev['razones'].append(f"La comunicación interna de la factura {c['factura']} ({c['banco']} {c['codigo']}) ya está en el cierre {c['correlativo']} (cajero {c['usuario']}).")
            bloqueo = servicio.regla_vigente(alm.vigente(str(fc), usr['usuario']), est.get('origen'), str(fc))
            if bloqueo: ev['presentable'] = False; ev['nivel'] = 'rojo'; ev['razones'].append(bloqueo)
        f = res['facturas']; b = res['banco']; qr = f[f['Tipo (motor)'].isin(['QR', 'C.I.'])]
        ci = {r['factura']: r for r in res.get('ci', [])}
        filas = []
        for _, r in qr.iterrows():
            nf = r['Número Factura']; e = r['estado_conciliacion']; nivel = r['nivel']
            obs = e in servicio.OBSERVADAS
            filas.append({'nf': nf, 'cliente': r['Razon Social'], 'estudiante': r['Nombre Estudiante'], 'monto': r['Monto'], 'hora': r['hora_factura'],
                          'estado': e, 'nivel': nivel, 'id': r['ID TRANSACCION'], 'titular': r['TITULAR ORIGEN'], 'hora_banco': r['hora_banco'],
                          'observada': obs, 'manual': nivel == 'RESPALDO MANUAL', 'ci': ci.get(nf),
                          'deshacer': e == 'CONFIRMADA POR CAJERO' or e == 'COMUNICACIÓN INTERNA',
                          'candidatos': servicio.candidatos(res, nf) if e in ('SIN RESPALDO', 'AMBIGUO') else []})
        observadas = [x for x in filas if x['observada']]; conciliadas = [x for x in filas if not x['observada']]
        lista = servicio.conceptos(t)
        return dict(bid=bid, est=est, t=t, ev=ev, observadas=observadas, conciliadas=conciliadas, tab_inicial='observadas' if observadas else 'conciliadas',
                    conceptos=lista, total=servicio.total_conceptos(lista), origen=est.get('origen'),
                    qr_conc=t['QR conciliado automático (CONCILIADA)'] + t['QR confirmado por cajero'],
                    ajenos=b[b['estado_banco'] == 'AJENO / NO CORRESPONDE AL CAJERO'].to_dict('records'),
                    anuladas=f[f['estado_conciliacion'].str.startswith('EXCLUIDA')].to_dict('records'),
                    tarjetas=res['tarjetas'].to_dict('records'), alertas=res['alertas'].to_dict('records'))

    @app.get('/borrador/{bid}', response_class=HTMLResponse)
    def ver_borrador(request: Request, bid: str):
        u, est = cargar_borrador(request, bid)
        if est.get('presentado'):
            raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        return render(request, 'borrador.html', **vista(bid, est, u))

    def estado_de(est, nf):
        return {'confirmada': nf in est['confirmaciones'], 'eleccion': est['elecciones'].get(nf), 'respaldo_manual': nf in est.get('respaldos_manuales', []),
                'reclasificacion': (est.get('reclasificaciones') or {}).get(nf)}

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
    def _asig(est, nf, idt):
        if nf and idt: est['elecciones'][nf] = idt
    app.post('/borrador/{bid}/confirmar')(accion('confirmar_coincidencia', _conf))
    app.post('/borrador/{bid}/deshacer')(accion('deshacer_confirmacion', _des))
    app.post('/borrador/{bid}/asignar')(accion('asignar_movimiento', _asig))

    @app.post('/borrador/{bid}/respaldo')
    def respaldo_manual(request: Request, bid: str, factura: str = Form(''), banco: str = Form(''), codigo: str = Form(''), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        nf = factura.strip(); antes = estado_de(est, nf)
        try:
            nuevo, info = alm.registrar_respaldo(bid, est, u['usuario'], nf, banco, codigo)
        except ErrorUsuario as e:
            alm.auditar(u['usuario'], 'respaldo_manual_rechazado', {'borrador': bid, 'factura': nf, 'banco': banco, 'codigo': codigo.strip(), 'motivo': str(e), 'anterior': antes})
            flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        alm.guardar_estado(bid, nuevo)
        alm.auditar(u['usuario'], 'respaldo_manual', {'borrador': bid, 'factura': nf, **info, 'anterior': antes})
        flash(request, f"Factura {nf}: respaldo manual registrado ({info['banco']} {info['codigo']}).", 'ok')
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

    @app.post('/borrador/{bid}/archivos')
    async def reemplazar_archivos(request: Request, bid: str, movimiento: UploadFile = File(None), banco: UploadFile = File(None), csrf: str = Form('')):
        u, est = cargar_borrador(request, bid); verificar_csrf(request, csrf)
        if est.get('presentado'): raise StarletteHTTPException(404, 'Este borrador ya fue presentado.')
        if not est.get('origen'): raise StarletteHTTPException(403, 'Solo se pueden reemplazar archivos al corregir un cierre rechazado.')
        with tempfile.TemporaryDirectory() as tmp:
            try:
                p1 = p2 = ext = n1 = n2 = None
                if movimiento is not None and movimiento.filename:
                    p1, n1 = await guardar_subida(movimiento, tmp, 'Movimiento del cajero', ('xls', 'xlsx')); ext, _ = servicio.validar_a1(p1)
                if banco is not None and banco.filename:
                    p2, n2 = await guardar_subida(banco, tmp, 'Reporte QR BANECO', ('xlsx',)); servicio.validar_a2(p2)
                if not p1 and not p2: raise ErrorUsuario('Seleccione al menos un archivo para reemplazar.')
                from . import motor
                a1 = p1 or os.path.join(alm.ruta_borrador(bid), est['archivo_movimiento']); a2 = p2 or os.path.join(alm.ruta_borrador(bid), est['archivo_banco'])
                try: r = motor.conciliar(a1, a2, est['tarjeta'])
                except Exception: raise ErrorUsuario('No se pudo conciliar con esos archivos: ' + servicio.MSG_A1)
                if r['totales']['fecha_cierre'] is None: raise ErrorUsuario('El movimiento del cajero no tiene facturas válidas.')
                alm.reemplazar_archivos(bid, est, p1, ext, n1, p2, n2)
            except ErrorUsuario as e:
                flash(request, str(e), 'error'); return RedirectResponse(f'/borrador/{bid}', 303)
        alm.auditar(u['usuario'], 'archivos_reemplazados', {'borrador': bid, 'movimiento': n1, 'banco': n2, 'origen': est['origen']['correlativo']})
        flash(request, 'Archivos reemplazados. Revise de nuevo las facturas observadas y las decisiones guardadas.', 'ok')
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
        flash(request, f'Cierre {corr} presentado correctamente.', 'ok')
        return RedirectResponse(f'/cierres/{cid}', 303)

    # ---------------- detalle de cierre ----------------
    def cargar_cierre(request, cid):
        u = usuario(request); c = alm.cierre(cid)
        if c is None or (u['rol'] != 'ingresos' and c['usuario'] != u['usuario']):
            if c is not None: alm.auditar(u['usuario'], 'acceso_denegado', {'cierre': cid})
            raise StarletteHTTPException(404, 'No se encontró el cierre.')
        return u, c

    @app.get('/cierres/{cid}', response_class=HTMLResponse)
    def detalle(request: Request, cid: int):
        u, c = cargar_cierre(request, cid)
        d = os.path.join(alm.dir, c['ruta'])
        archivos = [n for n in sorted(os.listdir(d))] if os.path.isdir(d) else []
        cj = None
        try:
            import json
            with open(os.path.join(d, 'cierre.json'), encoding='utf-8') as f: cj = json.load(f)
        except Exception: pass
        rem = alm.cierre(c['reemplaza_a']) if c['reemplaza_a'] else None
        posterior = None
        if c['estado_revision'] == servicio.ESTADO_REEMPLAZADO:
            con = alm.con()
            try: posterior = con.execute('SELECT id, correlativo FROM cierres WHERE reemplaza_a=?', (c['id'],)).fetchone()
            finally: con.close()
        conc = (cj or {}).get('conceptos', [])
        return render(request, 'cierre.html', c=c, archivos=archivos, cj=cj, anterior=rem, posterior=posterior, conceptos=conc,
                      total=servicio.total_conceptos(conc) if cj else None, ci=(cj or {}).get('comunicaciones_internas', []),
                      es_ingresos=u['rol'] == 'ingresos', es_dueno=u['rol'] == 'cajero' and c['usuario'] == u['usuario'])

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

    @app.get('/cierres/{cid}/archivo/{nombre}')
    def descargar(request: Request, cid: int, nombre: str):
        u, c = cargar_cierre(request, cid)
        ok = nombre in servicio.ARCHIVOS_CIERRE or nombre in ('movimiento_original.xls', 'movimiento_original.xlsx')
        ruta = os.path.join(alm.dir, c['ruta'], nombre)
        if not ok or not os.path.isfile(ruta): raise StarletteHTTPException(404, 'Archivo no encontrado.')
        alm.auditar(u['usuario'], 'descarga', {'cierre': c['correlativo'], 'archivo': nombre})
        mt = servicio.ARCHIVOS_CIERRE.get(nombre, 'application/vnd.ms-excel' if nombre.endswith('.xls') else servicio.ARCHIVOS_CIERRE['conciliacion.xlsx'])
        return FileResponse(ruta, media_type=mt, filename=f"{c['correlativo']}_{nombre}",
                            content_disposition_type='inline' if nombre in ('cierre.pdf', 'cierre_revision.pdf') else 'attachment')

    return app


app = create_app()
