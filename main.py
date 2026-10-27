import hashlib
import hmac
import json
import os
import re
import secrets
import smtplib
import ssl
import traceback
import warnings
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from email.message import EmailMessage

from fastapi import BackgroundTasks, Depends, FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sqlalchemy import func, inspect, or_, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session, joinedload
from starlette.middleware.sessions import SessionMiddleware

import models
from database import BASE_DIR, SessionLocal, engine

warnings.filterwarnings("ignore", message="Dialect sqlite.*Decimal")

CLAVE_MAESTRA = os.getenv("CLAVE_MAESTRA", "ADMIN2026")
CORREO_OFICIAL = "support1@eagleaeronauticalsupport.com"
CLAVE_CORREO_PRUEBA = "Colombia2024*"
NOMBRE_ARQUITECTO = "Arquitecto Maestro · Juan Esteban"

ROLES_SISTEMA = ["Administrador", "Gerencia", "Jefe de area", "Supervisión", "Asesor"]
ROLES_SUPERIORES = ["Administrador", "Gerencia", "Jefe de area"]
ROLES_ACCESO_USUARIOS = ["Administrador", "Gerencia", "Jefe de area", "Supervisión"]

ESTADOS = ["En proceso", "Facturado", "Aduana", "En tránsito", "Entregado"]
ESTADOS_ADUANA = ["No aplica / Pendiente", "En revisión aduanera", "Pago de aranceles", "Liberado / Nacionalizado"]
METODOS_PAGO_OPCIONES = ["Visa", "American", "PayPal", "Transferencia"]
UNIDADES_MEDIDA = ["UND", "EA", "GAL", "LT", "KIT", "LB", "ROLLO", "CAJA", "MT", "PC"]

PESOS_PROBABILIDAD = {
    "En proceso": 25,
    "Aduana": 55,
    "En tránsito": 75,
    "Facturado": 90,
    "Entregado": 100,
}


def cargar_secret_key() -> str:
    env = os.getenv("SECRET_KEY")
    if env:
        return env
    ruta = os.path.join(BASE_DIR, ".secret_key")
    if os.path.exists(ruta):
        with open(ruta) as f:
            return f.read().strip()
    clave = secrets.token_hex(32)
    with open(ruta, "w") as f:
        f.write(clave)
    return clave


GOOGLE_DRIVE_DISPONIBLE = False
try:
    from google.auth.transport.requests import Request as GoogleRequest
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    GOOGLE_DRIVE_DISPONIBLE = True
except ImportError:
    pass

SCOPES = ["https://www.googleapis.com/auth/drive.file"]
TOKEN_PATH = os.path.join(BASE_DIR, "token.json")
FOLDER_MIME = "application/vnd.google-apps.folder"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_drive_service = None


def obtener_drive():
    global _drive_service
    if not GOOGLE_DRIVE_DISPONIBLE:
        return None
    if _drive_service:
        return _drive_service
    if not os.path.exists(TOKEN_PATH):
        return None
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(GoogleRequest())
            with open(TOKEN_PATH, "w") as f:
                f.write(creds.to_json())
        else:
            return None
    _drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
    return _drive_service


def _q(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace("'", "\\'")


def carpeta_drive(service, nombre, padre_id=None):
    q = f"name='{_q(nombre)}' and mimeType='{FOLDER_MIME}' and trashed=false"
    if padre_id:
        q += f" and '{padre_id}' in parents"
    res = service.files().list(q=q, fields="files(id)", pageSize=1).execute()
    if res["files"]:
        return res["files"][0]["id"]
    meta = {"name": nombre, "mimeType": FOLDER_MIME}
    if padre_id:
        meta["parents"] = [padre_id]
    return service.files().create(body=meta, fields="id").execute()["id"]


def subir_a_drive(service, ruta_local, nombre_archivo, tipo, asesor):
    raiz = "Pedidos Nacionales" if tipo == "Nacional" else "Pedidos Internacionales"
    id_raiz = carpeta_drive(service, raiz)
    id_asesor = carpeta_drive(service, asesor, id_raiz)
    media = MediaFileUpload(ruta_local, mimetype=XLSX_MIME)
    existente = service.files().list(
        q=f"name='{_q(nombre_archivo)}' and '{id_asesor}' in parents and trashed=false",
        fields="files(id)", pageSize=1).execute()["files"]
    if existente:
        return service.files().update(fileId=existente[0]["id"], media_body=media,
                                      fields="id, webViewLink").execute()
    return service.files().create(body={"name": nombre_archivo, "parents": [id_asesor]},
                                  media_body=media, fields="id, webViewLink").execute()


ENC_BASE = [
    "1. Fecha inicial",
    "2. PO",
    "3. Cliente",
    "4. Nombre del asesor",
    "5. Fecha de compra a proveedores",
    "6. Proveedor (Origen)",
    "7. Parte número por producto (SKU)",
    "8. Parte número alterno",
    "9. Descripción del producto",
    "10. Unidad de medida",
    "11. Cantidad",
    "12. Método de pago",
    "13. Precio unitario inicial ($)",
    "14. Precio total ($)",
    "15. IVA / TAX (%)",
    "16. Valor después de IVA / TAX ($)",
    "17. Plazo de entrega prometido",
]

ENC_NACIONAL = ENC_BASE + [
    "18. Días Hábiles (Semáforo)",
    "19. Fecha de entrega al cliente",
    "20. Número de remisión",
    "21. Documentación",
    "22. Estado",
    "23. Observaciones",
]

ENC_INTERNACIONAL = ENC_BASE + [
    "18. ¿Llega a Click Mail?",
    "19. Fecha llegada Click Mail",
    "20. Tracking",
    "21. Nacionalización (Estado)",
    "22. Fecha de Nacionalización",
    "23. Valor de Nacionalización ($)",
    "24. Días Hábiles (Semáforo)",
    "25. Fecha de entrega al cliente",
    "26. Número de remisión",
    "27. Documentación",
    "28. Estado",
    "29. Observaciones",
]


def generar_excel(ruta, p: models.Pedido, intl: models.PedidoInternacional | None):
    wb = Workbook()
    ws = wb.active
    ws.title = "FLY2026"

    if p.tipo == "Nacional":
        ws.merge_cells("A1:W1")
        titulo, encabezados = "CUADRO DE COMPRAS NACIONALES FLY LOGISTIC 2026", ENC_NACIONAL
    else:
        ws.merge_cells("A1:AC1")
        titulo, encabezados = "CUADRO DE COMPRAS EN EL EXTERIOR FLY LOGISTIC 2026", ENC_INTERNACIONAL

    t = ws["A1"]
    t.value = titulo
    t.font = Font(bold=True, size=15, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="002060")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.append(encabezados)
    relleno_enc = PatternFill("solid", fgColor="2563EB")
    for c in ws[2]:
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = relleno_enc
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    semaforo = p.dias_habiles_info["detalle"]
    val_nac = float(p.valor_nacionalizacion or 0)

    for idx, item in enumerate(p.items, start=3):
        iva_item = float(item.iva_porcentaje if item.iva_porcentaje is not None else p.iva_porcentaje)
        medio = []
        if p.tipo == "Internacional":
            llega_cm = (intl.llega_click_mail if intl and intl.llega_click_mail else ("Si" if (intl and intl.fecha_click_mail) else "No"))
            fecha_cm = intl.fecha_click_mail if (intl and llega_cm == "Si") else None
            medio = [
                llega_cm,
                fecha_cm,
                intl.tracking if intl else "",
                intl.estado_nacionalizacion if intl else "",
                intl.fecha_nacionalizacion if intl else None,
                val_nac if idx == 3 else 0.0,
            ]
        cola = [
            semaforo,
            p.fecha_entrega_cliente,
            p.numero_remision,
            p.documentacion,
            p.estado,
            p.observaciones,
        ]
        fila = [
            p.fecha_inicial,
            p.po,
            p.cliente.nombre if p.cliente else "",
            p.asesor.nombre if p.asesor else "",
            item.fecha_compra_proveedor,
            item.proveedor,
            item.parte_numero,
            item.parte_numero_alterno or "",
            item.descripcion,
            item.unidad_medida or "UND",
            int(item.cantidad),
            p.metodo_pago or "",
            float(item.precio_unitario),
            f"=K{idx}*M{idx}",
            iva_item,
            f"=N{idx}*(1+O{idx}/100)",
            p.plazo_entrega,
        ] + medio + cola
        ws.append(fila)

    ultima_fila_items = 2 + len(p.items)
    fila_tot = ultima_fila_items + 1
    ws[f"L{fila_tot}"] = "TOTALES PO:"
    ws[f"L{fila_tot}"].font = Font(bold=True, color="002060")
    ws[f"N{fila_tot}"] = f"=SUM(N3:N{ultima_fila_items})"
    ws[f"N{fila_tot}"].font = Font(bold=True)
    ws[f"N{fila_tot}"].number_format = '"$"#,##0.00'
    ws[f"P{fila_tot}"] = f"=SUM(P3:P{ultima_fila_items})"
    ws[f"P{fila_tot}"].font = Font(bold=True, color="002060")
    ws[f"P{fila_tot}"].number_format = '"$"#,##0.00'

    if p.tipo == "Internacional":
        ws[f"W{fila_tot}"] = f"=SUM(W3:W{ultima_fila_items})"
        ws[f"W{fila_tot}"].font = Font(bold=True, color="002060")
        ws[f"W{fila_tot}"].number_format = '"$"#,##0.00'

    for i, enc in enumerate(encabezados, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(16, len(enc) * 0.92)

    for row in range(3, 3 + len(p.items)):
        for celda in (f"M{row}", f"N{row}", f"P{row}"):
            ws[celda].number_format = '"$"#,##0.00'
        if p.tipo == "Internacional":
            ws[f"W{row}"].number_format = '"$"#,##0.00'
        for c in ws[row]:
            if hasattr(c.value, "year"):
                c.number_format = "YYYY-MM-DD"

    wb.save(ruta)


def nombre_seguro(texto: str) -> str:
    s = re.sub(r"[^\w\-. ]", "_", (texto or "").strip(), flags=re.UNICODE).strip(" .")
    return s or "sin_nombre"


def ruta_excel(p):
    carpeta = os.path.join(BASE_DIR, f"Pedidos_{p.tipo}", nombre_seguro(p.asesor.nombre))
    os.makedirs(carpeta, exist_ok=True)
    nombre = f"Pedido_{nombre_seguro(p.po)}.xlsx"
    return os.path.join(carpeta, nombre), nombre


def generar_salidas(db: Session, p) -> str:
    try:
        ruta, nombre = ruta_excel(p)
        generar_excel(ruta, p, p.internacional)
    except Exception as e:
        return f"Pedido guardado, pero falló el Excel: {e}"
    try:
        service = obtener_drive()
        if service:
            archivo = subir_a_drive(service, ruta, nombre, p.tipo, p.asesor.nombre)
            p.drive_link = archivo.get("webViewLink", "")
            db.commit()
            return "Excel actualizado y subido a Google Drive"
        return "Excel generado localmente"
    except Exception as e:
        return f"Excel generado localmente; aviso Drive: {e}"


ITER = 200_000


def hash_clave(clave: str) -> str:
    sal = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", clave.encode(), bytes.fromhex(sal), ITER).hex()
    return f"pbkdf2${ITER}${sal}${h}"


def verificar_clave(clave: str, guardado) -> bool:
    try:
        _, iters, sal, h = (guardado or "").split("$")
        calc = hashlib.pbkdf2_hmac("sha256", clave.encode(), bytes.fromhex(sal), int(iters)).hex()
        return hmac.compare_digest(calc, h)
    except ValueError:
        hashlib.pbkdf2_hmac("sha256", clave.encode(), b"x" * 16, ITER)
        return False


def buscar_autorizador_por_clave(db: Session, clave_ingresada: str, jefe_id_seleccionado: int | None = None):
    if not clave_ingresada:
        return False, None
    superiores = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).all()
    if jefe_id_seleccionado:
        jefe_sel = db.get(models.Asesor, jefe_id_seleccionado)
        if jefe_sel and jefe_sel.password_hash and verificar_clave(clave_ingresada, jefe_sel.password_hash):
            return True, jefe_sel
    for sup in superiores:
        if sup.password_hash and verificar_clave(clave_ingresada, sup.password_hash):
            return True, sup
    if hmac.compare_digest(clave_ingresada.encode(), CLAVE_MAESTRA.encode()):
        jefe_sel = db.get(models.Asesor, jefe_id_seleccionado) if jefe_id_seleccionado else None
        return True, jefe_sel
    return False, None


def migrar_bd():
    insp = inspect(engine)
    with engine.begin() as con:
        for tabla in models.Base.metadata.sorted_tables:
            if not insp.has_table(tabla.name):
                continue
            existentes = {c["name"] for c in insp.get_columns(tabla.name)}
            for col in tabla.columns:
                if col.name not in existentes:
                    tipo = col.type.compile(engine.dialect)
                    con.execute(text(f'ALTER TABLE {tabla.name} ADD COLUMN "{col.name}" {tipo}'))

        if insp.has_table("asesores"):
            con.execute(text("UPDATE asesores SET rol='Asesor' WHERE rol IS NULL OR rol=''"))
            con.execute(text("UPDATE asesores SET email=:correo WHERE email IS NULL OR email=''"),
                        {"correo": CORREO_OFICIAL})
            filas = con.execute(text(
                "SELECT id, nombre FROM asesores WHERE nombre_norm IS NULL OR nombre_norm = ''")).fetchall()
            for i, nombre in filas:
                con.execute(text("UPDATE asesores SET nombre_norm=:n WHERE id=:i"),
                            {"n": models.normalizar(nombre), "i": i})
            con.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS ix_asesores_username ON asesores(username)"))

        if insp.has_table("pedidos"):
            con.execute(text("UPDATE pedidos SET valor_nacionalizacion = 0 WHERE valor_nacionalizacion IS NULL"))

        if insp.has_table("pedido_items"):
            con.execute(text("UPDATE pedido_items SET parte_numero_alterno = '' WHERE parte_numero_alterno IS NULL"))
            con.execute(text("UPDATE pedido_items SET unidad_medida = 'UND' WHERE unidad_medida IS NULL OR unidad_medida = ''"))
            con.execute(text("UPDATE pedido_items SET iva_porcentaje = 19 WHERE iva_porcentaje IS NULL"))

        if insp.has_table("pedidos_internacionales"):
            con.execute(text("UPDATE pedidos_internacionales SET llega_click_mail = 'Si' WHERE fecha_click_mail IS NOT NULL AND (llega_click_mail IS NULL OR llega_click_mail = '')"))
            con.execute(text("UPDATE pedidos_internacionales SET llega_click_mail = 'No' WHERE llega_click_mail IS NULL OR llega_click_mail = ''"))


def inicializar_admin_y_correo():
    ruta_principal = os.path.join(BASE_DIR, "correo_config.json")
    if not os.path.exists(ruta_principal):
        datos_smtp = {
            "host": "mail.eagleaeronauticalsupport.com",
            "port": 465,
            "user": CORREO_OFICIAL,
            "password": CLAVE_CORREO_PRUEBA,
            "from": f"Fly Logistic Command Center <{CORREO_OFICIAL}>",
            "app_url": "http://127.0.0.1:8000"
        }
        for nombre_cfg in ("correo_config.json", "correo_config.json.json"):
            try:
                with open(os.path.join(BASE_DIR, nombre_cfg), "w", encoding="utf-8") as f:
                    json.dump(datos_smtp, f, indent=2, ensure_ascii=False)
            except Exception:
                pass

    db = SessionLocal()
    try:
        n_norm = models.normalizar(NOMBRE_ARQUITECTO)
        admin = db.query(models.Asesor).filter(func.lower(models.Asesor.username) == "admin").first()
        if not admin:
            admin = db.query(models.Asesor).filter_by(nombre_norm=n_norm).first()
        if not admin:
            admin = models.Asesor(
                nombre=NOMBRE_ARQUITECTO,
                nombre_norm=n_norm,
                username="admin",
                email=CORREO_OFICIAL,
                rol="Administrador"
            )
            db.add(admin)
        else:
            admin.nombre = NOMBRE_ARQUITECTO
            admin.nombre_norm = n_norm
            admin.username = "admin"
            admin.rol = "Administrador"
            admin.email = admin.email or CORREO_OFICIAL
        if not admin.password_hash:
            admin.password_hash = hash_clave(CLAVE_CORREO_PRUEBA)
        db.commit()
    except Exception as e:
        db.rollback()
        print(f"[AVISO ADMIN]: {e}")
    finally:
        db.close()


models.Base.metadata.create_all(bind=engine)
migrar_bd()
inicializar_admin_y_correo()

app = FastAPI(title="Fly Logistic Command Center")


@app.middleware("http")
async def interceptor_de_errores(request: Request, call_next):
    try:
        return await call_next(request)
    except Exception:
        trace = traceback.format_exc()
        print("\n[ERROR INTERNO DETECTADO]\n", trace)
        return HTMLResponse(
            content=f"""
            <div style="background:#0f172a; color:#38bdf8; padding:32px; font-family:monospace; min-height:100vh;">
                <h2 style="color:#f87171; margin-bottom:12px;">Diagnóstico Técnico del Servidor</h2>
                <p style="color:#e2e8f0; margin-bottom:16px;">Ruta: <strong>{request.url.path}</strong></p>
                <pre style="background:#1e293b; color:#4ade80; padding:20px; border-radius:8px; overflow-x:auto;">{trace}</pre>
                <p style="margin-top:20px;"><a href="/pedidos" style="color:#fff; background:#2563eb; padding:10px 18px; border-radius:6px; text-decoration:none;">← Volver al Centro de Mando</a></p>
            </div>
            """,
            status_code=500
        )


app.add_middleware(SessionMiddleware, secret_key=cargar_secret_key(), max_age=8 * 3600, same_site="lax")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))
templates.env.filters["money"] = lambda v: f"${Decimal(str(v or 0)):,.2f}"


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


class NoAutenticado(Exception):
    pass


@app.exception_handler(NoAutenticado)
async def _redirigir_login(request: Request, exc: NoAutenticado):
    return RedirectResponse("/login", status_code=303)


def usuario_actual(request: Request, db: Session = Depends(get_db)) -> models.Asesor:
    uid = request.session.get("asesor_id")
    user = db.get(models.Asesor, uid) if uid else None
    if not user:
        request.session.clear()
        raise NoAutenticado()
    return user


def ids_asesores_visibles(db: Session, user: models.Asesor) -> list[int] | None:
    if user.rol in ["Administrador", "Gerencia"]:
        return None
    if user.rol == "Jefe de area":
        sub_ids = [a.id for a in db.query(models.Asesor).filter(models.Asesor.jefe_id == user.id).all()]
        sub_ids.append(user.id)
        return sub_ids
    return [user.id]


def tiene_permiso_gestion(user: models.Asesor, objetivo: models.Asesor) -> bool:
    if user.rol in ["Administrador", "Gerencia"]:
        return True
    if user.rol == "Jefe de area" and objetivo.jefe_id == user.id:
        return True
    if user.rol == "Supervisión":
        return True
    return False


def render(request: Request, nombre: str, ctx: dict | None = None, status: int = 200):
    ctx = dict(ctx or {})
    ctx.setdefault("flash", request.session.pop("flash", None))
    ctx["request"] = request
    return templates.TemplateResponse(request=request, name=nombre, context=ctx, status_code=status)


EMAIL_RE = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")
RESET_MINUTOS = 30


def config_correo() -> dict:
    cfg = {
        "host": "mail.eagleaeronauticalsupport.com",
        "port": 465,
        "user": CORREO_OFICIAL,
        "password": CLAVE_CORREO_PRUEBA,
        "from": f"Fly Logistic Command Center <{CORREO_OFICIAL}>",
        "app_url": "http://127.0.0.1:8000"
    }
    for filename in ("correo_config.json", "correo_config.json.json"):
        ruta = os.path.join(BASE_DIR, filename)
        if os.path.exists(ruta):
            try:
                with open(ruta, encoding="utf-8") as f:
                    leido = json.load(f)
                    cfg.update({k: v for k, v in leido.items() if v is not None and str(v).strip() != ""})
                break
            except Exception:
                pass
    for k, env in (("host", "SMTP_HOST"), ("port", "SMTP_PORT"), ("user", "SMTP_USER"),
                   ("password", "SMTP_PASSWORD"), ("from", "SMTP_FROM"), ("app_url", "APP_URL")):
        if os.getenv(env):
            cfg[k] = os.getenv(env)
    return cfg


def guardar_config_correo(host: str, port: int, user_smtp: str, password_smtp: str):
    cfg = config_correo()
    cfg.update({
        "host": host.strip(),
        "port": int(port),
        "user": user_smtp.strip(),
        "password": password_smtp,
        "from": f"Fly Logistic Command Center <{user_smtp.strip()}>"
    })
    for filename in ("correo_config.json", "correo_config.json.json"):
        try:
            with open(os.path.join(BASE_DIR, filename), "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception:
            pass


def enviar_correo_con_diagnostico(destino: str, asunto: str, cuerpo: str) -> tuple[bool, str]:
    """Intenta enviar correo negociando automáticamente SSL (465) y STARTTLS (587)
    en el servidor configurado y servidores alternativos del dominio para evitar desconexiones."""
    cfg = config_correo()
    remitente = cfg.get("user") or CORREO_OFICIAL
    clave = cfg.get("password") or CLAVE_CORREO_PRUEBA
    host_principal = (cfg.get("host") or "mail.eagleaeronauticalsupport.com").strip()
    puerto_principal = int(cfg.get("port") or 465)

    msg = EmailMessage()
    msg["From"] = cfg.get("from") or f"Fly Logistic Command Center <{remitente}>"
    msg["To"] = destino
    msg["Subject"] = asunto
    msg.set_content(cuerpo)

    intentos = [
        (host_principal, puerto_principal),
        ("mail.eagleaeronauticalsupport.com", 465),
        ("mail.eagleaeronauticalsupport.com", 587),
        ("smtp.office365.com", 587),
        ("smtp.gmail.com", 465),
        ("smtp.gmail.com", 587),
    ]
    vistos = set()
    errores = []

    ctx_ssl = ssl.create_default_context()
    ctx_sin_verificar = ssl._create_unverified_context()

    for host, port in intentos:
        if (host, port) in vistos:
            continue
        vistos.add((host, port))
        for contexto in (ctx_ssl, ctx_sin_verificar):
            try:
                if int(port) == 465:
                    with smtplib.SMTP_SSL(host, int(port), context=contexto, timeout=12) as s:
                        s.ehlo()
                        s.login(remitente, clave)
                        s.send_message(msg)
                else:
                    with smtplib.SMTP(host, int(port), timeout=12) as s:
                        s.ehlo()
                        s.starttls(context=contexto)
                        s.ehlo()
                        s.login(remitente, clave)
                        s.send_message(msg)
                detalle_ok = f"Correo entregado vía {host}:{port} a {destino}"
                print(f"\n[SMTP OK] {detalle_ok}\n")
                return True, detalle_ok
            except Exception as e:
                errores.append(f"{host}:{port} -> {e}")

    resumen_err = " | ".join(errores[:3])
    print(
        f"\n==================== [DIAGNÓSTICO SMTP / RESPALDO EN CONSOLA] ====================\n"
        f"De: {remitente} | Para: {destino}\n"
        f"Asunto: {asunto}\n"
        f"Detalle Técnico SMTP: {resumen_err}\n"
        f"--- MENSAJE GENERADO ---\n{cuerpo}\n"
        f"==================================================================================\n"
    )
    return False, resumen_err


def enviar_correo(destino: str, asunto: str, cuerpo: str) -> bool:
    ok, _ = enviar_correo_con_diagnostico(destino, asunto, cuerpo)
    return ok


def construir_cuerpo_alerta_po(p: models.Pedido) -> tuple[str, str]:
    info_dias = p.dias_habiles_info
    estado_semaforo = "EN CONTRA (VENCIDO / RETRASADO)" if info_dias["en_contra"] else "A FAVOR (EN TIEMPO)"
    asunto = f"[Fly Logistic Command Center] Seguimiento {info_dias['etiqueta']} - Cliente: {p.cliente.nombre}"

    lineas_items = []
    for idx, it in enumerate(p.items, start=1):
        alt_txt = f" (Alterno: {it.parte_numero_alterno})" if it.parte_numero_alterno else ""
        lineas_items.append(
            f"  {idx}. P/N: {it.parte_numero}{alt_txt} | {it.descripcion} | "
            f"Cant: {it.cantidad} {it.unidad_medida} | Unit: ${Decimal(str(it.precio_unitario)):,.2f} | "
            f"IVA/TAX: {Decimal(str(it.iva_porcentaje)):g}% | Total Línea: ${it.total_item:,.2f}"
        )

    intl_txt = ""
    if p.tipo == "Internacional":
        i = p.internacional
        llega_cm = i.llega_click_mail if i else "No"
        fecha_cm = i.fecha_click_mail.isoformat() if (i and i.fecha_click_mail) else "No aplica"
        fecha_nac = i.fecha_nacionalizacion.isoformat() if (i and i.fecha_nacionalizacion) else "Pendiente"
        est_nac = i.estado_nacionalizacion if (i and i.estado_nacionalizacion) else "Pendiente"
        intl_txt = (
            f"\n--- DATOS INTERNACIONALES Y NACIONALIZACIÓN ---\n"
            f"• ¿Llega a Click Mail?: {llega_cm} (Fecha: {fecha_cm})\n"
            f"• Tracking: {(i.tracking if i and i.tracking else 'Sin asignar')}\n"
            f"• Estado de Nacionalización: {est_nac}\n"
            f"• Fecha de Nacionalización: {fecha_nac}\n"
            f"• Valor de la Nacionalización del Producto: ${p.valor_nacionalizacion_dec:,.2f}\n"
        )

    cuerpo = (
        f"Hola {p.asesor.nombre},\n\n"
        f"Reporte de seguimiento operativo para la Orden de Compra PO {p.po}:\n\n"
        f"• Número de PO: {p.po}\n"
        f"• Cliente: {p.cliente.nombre}\n"
        f"• Tipo de Operación: {p.tipo}\n"
        f"• Estado General: {p.estado}\n"
        f"• Fecha Inicial: {p.fecha_inicial.isoformat()}\n"
        f"• Plazo de Entrega Prometido: {p.plazo_entrega.isoformat() if p.plazo_entrega else 'Sin definir'}\n"
        f"• Semáforo de Días Hábiles: {info_dias['etiqueta']} ({estado_semaforo})\n"
        f"• Método(s) de Pago: {p.metodo_pago or 'Sin definir'}\n"
        f"{intl_txt}\n"
        f"--- DETALLE DE PRODUCTOS ---\n"
        + "\n".join(lineas_items) +
        f"\n\n--- RESUMEN DE VALORES ---\n"
        f"• Subtotal Inicial Productos: ${p.subtotal:,.2f}\n"
        f"• Total Impuestos (IVA 19% / TAX 7%): ${p.total_impuestos:,.2f}\n"
        f"• Valor de Nacionalización: ${p.valor_nacionalizacion_dec:,.2f}\n"
        f"• GRAN TOTAL PO: ${p.total:,.2f}\n\n"
        f"Fly Logistic Command Center ({CORREO_OFICIAL})"
    )
    return asunto, cuerpo


CAMPOS = [
    "fecha_inicial", "po_numero", "cliente", "tipo_pedido", "metodo_pago", "iva_porcentaje",
    "valor_nacionalizacion", "plazo_entrega", "fecha_entrega_cliente", "numero_remision",
    "llega_click_mail", "fecha_click_mail", "tracking", "estado_nacionalizacion",
    "fecha_nacionaliza", "documentacion", "estado_general", "observaciones", "asesor_asignado_id"
]
CAMPOS_ITEM = [
    "fecha_compra_prov", "proveedor", "parte_numero", "parte_numero_alterno",
    "unidad_medida", "descripcion", "cantidad", "precio_unitario", "iva_item"
]


def datos_desde_form(f) -> dict:
    d = {k: (f.get(k) or "").strip() for k in CAMPOS}
    metodos_lista = [m.strip() for m in f.getlist("metodos_pago_check") if m.strip()]
    if metodos_lista:
        d["metodo_pago"] = ", ".join(metodos_lista)
    listas = [f.getlist(k) for k in CAMPOS_ITEM]
    d["items"] = [
        dict(zip(CAMPOS_ITEM, [str(v).strip() for v in fila]))
        for fila in zip(*listas)
        if any(str(v).strip() for v in fila)
    ]
    return d


def datos_desde_pedido(p: models.Pedido) -> dict:
    iso = lambda v: v.isoformat() if v else ""
    i = p.internacional
    llega_cm = i.llega_click_mail if (i and i.llega_click_mail) else ("Si" if (i and i.fecha_click_mail) else "No")
    return {
        "fecha_inicial": iso(p.fecha_inicial),
        "po_numero": p.po,
        "cliente": p.cliente.nombre,
        "tipo_pedido": p.tipo,
        "metodo_pago": p.metodo_pago or "",
        "iva_porcentaje": f"{Decimal(str(p.iva_porcentaje)):g}",
        "valor_nacionalizacion": f"{Decimal(str(p.valor_nacionalizacion or 0)):.2f}",
        "plazo_entrega": iso(p.plazo_entrega),
        "fecha_entrega_cliente": iso(p.fecha_entrega_cliente),
        "numero_remision": p.numero_remision or "",
        "llega_click_mail": llega_cm,
        "fecha_click_mail": iso(i.fecha_click_mail) if i else "",
        "tracking": (i.tracking or "") if i else "",
        "estado_nacionalizacion": (i.estado_nacionalizacion or "") if i else "",
        "fecha_nacionaliza": iso(i.fecha_nacionalizacion) if i else "",
        "documentacion": p.documentacion or "",
        "estado_general": p.estado,
        "observaciones": p.observaciones or "",
        "asesor_asignado_id": str(p.asesor_id),
        "items": [{
            "fecha_compra_prov": iso(it.fecha_compra_proveedor),
            "proveedor": it.proveedor,
            "parte_numero": it.parte_numero,
            "parte_numero_alterno": it.parte_numero_alterno or "",
            "unidad_medida": it.unidad_medida or "UND",
            "descripcion": it.descripcion,
            "cantidad": str(it.cantidad),
            "precio_unitario": f"{Decimal(str(it.precio_unitario)):.2f}",
            "iva_item": f"{Decimal(str(it.iva_porcentaje if it.iva_porcentaje is not None else p.iva_porcentaje)):g}",
        } for it in p.items],
    }


def _fecha(txt, campo, obligatoria=False):
    if not txt:
        if obligatoria:
            raise ValueError(f"La fecha «{campo}» es obligatoria.")
        return None
    try:
        return datetime.strptime(txt, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"Fecha inválida en «{campo}».")


def _decimal(txt, campo):
    try:
        v = Decimal(txt)
    except (InvalidOperation, TypeError):
        raise ValueError(f"Número inválido en «{campo}».")
    if not v.is_finite():
        raise ValueError(f"Número inválido en «{campo}».")
    return v


def parsear(d: dict, tipo_fijo: str | None = None) -> dict:
    if not d.get("po_numero"):
        raise ValueError("El número de PO es obligatorio.")
    if not d.get("cliente"):
        raise ValueError("El cliente es obligatorio.")
    tipo = tipo_fijo or d.get("tipo_pedido")
    if tipo not in ("Nacional", "Internacional"):
        raise ValueError("Selecciona el tipo de pedido.")

    iva_defecto = "19" if tipo == "Nacional" else "7"
    iva_cabecera = _decimal(d.get("iva_porcentaje") or iva_defecto, "IVA / TAX")
    if not 0 <= iva_cabecera <= 100:
        raise ValueError("El porcentaje de IVA / TAX debe estar entre 0 y 100.")

    val_nac = _decimal(d.get("valor_nacionalizacion") or "0", "Valor de Nacionalización")
    if val_nac < 0:
        raise ValueError("El valor de nacionalización no puede ser negativo.")

    if d.get("estado_general") not in ESTADOS:
        raise ValueError("Estado general inválido.")
    if not d.get("items"):
        raise ValueError("Agrega al menos un producto.")

    items = []
    for n, it in enumerate(d["items"], start=1):
        for k, etiqueta in (("proveedor", "proveedor"), ("parte_numero", "parte número (SKU)"), ("descripcion", "descripción")):
            if not it.get(k):
                raise ValueError(f"Producto {n}: falta {etiqueta}.")
        try:
            cant = int(it.get("cantidad", 0))
        except ValueError:
            raise ValueError(f"Producto {n}: cantidad inválida.")
        if cant < 1:
            raise ValueError(f"Producto {n}: la cantidad debe ser al menos 1.")
        precio = _decimal(it.get("precio_unitario", 0), f"precio unitario del producto {n}")
        if precio < 0:
            raise ValueError(f"Producto {n}: el precio no puede ser negativo.")

        iva_it_txt = it.get("iva_item")
        iva_it = _decimal(iva_it_txt if iva_it_txt != "" else str(iva_cabecera), f"IVA/TAX del producto {n}")
        if not 0 <= iva_it <= 100:
            raise ValueError(f"Producto {n}: el IVA/TAX debe estar entre 0 y 100.")

        items.append({
            "fecha_compra_proveedor": _fecha(it.get("fecha_compra_prov"), f"compra a proveedor (producto {n})"),
            "proveedor": it["proveedor"],
            "parte_numero": it["parte_numero"],
            "parte_numero_alterno": it.get("parte_numero_alterno") or "",
            "unidad_medida": it.get("unidad_medida") or "UND",
            "descripcion": it["descripcion"],
            "cantidad": cant,
            "precio_unitario": precio,
            "iva_porcentaje": iva_it,
        })

    llega_cm = "Si" if d.get("llega_click_mail") == "Si" else "No"
    fecha_cm = _fecha(d.get("fecha_click_mail"), "Fecha Click Mail") if llega_cm == "Si" else None

    return {
        "po": d["po_numero"],
        "tipo": tipo,
        "fecha_inicial": _fecha(d.get("fecha_inicial"), "Fecha inicial", obligatoria=True),
        "cliente": d["cliente"],
        "metodo_pago": d.get("metodo_pago"),
        "iva_porcentaje": iva_cabecera,
        "valor_nacionalizacion": val_nac if tipo == "Internacional" else Decimal("0"),
        "plazo_entrega": _fecha(d.get("plazo_entrega"), "Plazo de entrega"),
        "fecha_entrega_cliente": _fecha(d.get("fecha_entrega_cliente"), "Entrega al cliente"),
        "numero_remision": d.get("numero_remision"),
        "documentacion": d.get("documentacion"),
        "estado": d["estado_general"],
        "observaciones": d.get("observaciones"),
        "items": items,
        "intl": {
            "llega_click_mail": llega_cm,
            "fecha_click_mail": fecha_cm,
            "tracking": d.get("tracking"),
            "estado_nacionalizacion": d.get("estado_nacionalizacion"),
            "fecha_nacionalizacion": _fecha(d.get("fecha_nacionaliza"), "Fecha de Nacionalización"),
        } if tipo == "Internacional" else None,
    }


def aplicar(db: Session, pedido: models.Pedido, v: dict):
    cli = db.query(models.Cliente).filter(func.lower(models.Cliente.nombre) == v["cliente"].lower()).first()
    if not cli:
        cli = models.Cliente(nombre=v["cliente"])
        db.add(cli)
        db.flush()
    pedido.cliente_id = cli.id
    for k in ("fecha_inicial", "metodo_pago", "iva_porcentaje", "valor_nacionalizacion", "plazo_entrega",
              "fecha_entrega_cliente", "numero_remision", "documentacion", "estado", "observaciones"):
        setattr(pedido, k, v[k])
    pedido.items.clear()
    for it in v["items"]:
        pedido.items.append(models.ItemPedido(**it))

    if v["intl"]:
        if pedido.internacional:
            for k, val in v["intl"].items():
                setattr(pedido.internacional, k, val)
        else:
            pedido.internacional = models.PedidoInternacional(**v["intl"])
    else:
        pedido.internacional = None


def pedido_del_usuario(db: Session, pid: int, user: models.Asesor) -> models.Pedido | None:
    visibles = ids_asesores_visibles(db, user)
    q = db.query(models.Pedido).options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.internacional)
    ).filter(models.Pedido.id == pid)
    if visibles is not None:
        q = q.filter(models.Pedido.asesor_id.in_(visibles))
    return q.first()


@app.get("/")
def inicio(request: Request):
    return RedirectResponse("/pedidos" if request.session.get("asesor_id") else "/login", status_code=303)


@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request):
    if request.session.get("asesor_id"):
        return RedirectResponse("/pedidos", status_code=303)
    return render(request, "login.html", {"username": "", "error": None})


@app.post("/login", response_class=HTMLResponse)
async def login_post(request: Request, db: Session = Depends(get_db)):
    f = await request.form()
    raw_user = (f.get("username") or "").strip()
    raw_pass = (f.get("password") or "").strip()
    username_norm = models.normalizar(raw_user)
    user = db.query(models.Asesor).filter_by(username=username_norm).first() if username_norm else None

    if not user or not verificar_clave(raw_pass, user.password_hash):
        return render(request, "login.html", {"username": raw_user, "error": "Usuario o contraseña incorrectos."}, 401)

    request.session.clear()
    request.session["asesor_id"] = user.id
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


@app.get("/registro", response_class=HTMLResponse)
def registro_form(request: Request, db: Session = Depends(get_db)):
    superiores = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).order_by(models.Asesor.nombre).all()
    return render(request, "registro.html", {"v": {"email": CORREO_OFICIAL}, "superiores": superiores, "error": None})


@app.post("/registro", response_class=HTMLResponse)
async def registro_post(request: Request, db: Session = Depends(get_db)):
    f = await request.form()
    v = {k: (f.get(k) or "").strip() for k in ("nombre", "username", "email", "jefe_id")}
    if not v["email"]:
        v["email"] = CORREO_OFICIAL
    usuario = models.normalizar(v["username"])
    password = (f.get("password") or "").strip()
    password2 = (f.get("password2") or "").strip()
    clave_autorizacion = (f.get("clave_maestra") or "").strip()
    jefe_id_int = int(v["jefe_id"]) if v["jefe_id"].isdigit() else None
    superiores = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).order_by(models.Asesor.nombre).all()

    def error(msg, status=400):
        return render(request, "registro.html", {"v": v, "superiores": superiores, "error": msg}, status)

    autorizado, sup_encontrado = buscar_autorizador_por_clave(db, clave_autorizacion, jefe_id_int)
    if not autorizado:
        return error("La clave de autorización no coincide con ningún Administrador, Gerencia o Jefe de área.", 403)
    if len(v["nombre"]) < 3:
        return error("Escribe tu nombre completo.")
    if not v["username"]:
        return error("El usuario es obligatorio.")
    if not EMAIL_RE.fullmatch(v["email"]):
        return error("Escribe un correo electrónico válido.")
    if not 3 <= len(usuario) <= 100:
        return error("El usuario debe tener entre 3 y 100 caracteres.")
    if len(password) < 8:
        return error("La contraseña debe tener al menos 8 caracteres.")
    if password != password2:
        return error("Las contraseñas no coinciden.")
    if db.query(models.Asesor).filter_by(username=usuario).first():
        return error("Ese nombre de usuario ya está en uso.", 409)

    asesor = db.query(models.Asesor).filter_by(nombre_norm=models.normalizar(v["nombre"])).first()
    if asesor and asesor.password_hash:
        return error("Ya existe una cuenta registrada para ese nombre. Inicia sesión.", 409)
    if not asesor:
        asesor = models.Asesor(nombre=v["nombre"], nombre_norm=models.normalizar(v["nombre"]), rol="Asesor")
        db.add(asesor)

    asesor.username = usuario
    asesor.password_hash = hash_clave(password)
    asesor.email = v["email"]
    asesor.rol = "Asesor"
    asesor.jefe_id = jefe_id_int or (sup_encontrado.id if sup_encontrado else None)
    db.commit()

    request.session.clear()
    request.session["asesor_id"] = asesor.id
    jefe_txt = f" asignado a {sup_encontrado.nombre}" if sup_encontrado else ""
    request.session["flash"] = {"ok": True, "msg": f"Cuenta creada exitosamente{jefe_txt}. ¡Bienvenido, {asesor.nombre}!"}
    return RedirectResponse("/pedidos", status_code=303)


def _hash_token(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


def asesor_por_token(db: Session, token: str):
    a = db.query(models.Asesor).filter_by(reset_token_hash=_hash_token(token)).first()
    if a and a.reset_expira and a.reset_expira > datetime.utcnow():
        return a
    return None


@app.get("/olvide", response_class=HTMLResponse)
def olvide_form(request: Request):
    return render(request, "olvide.html", {
        "enviado": False, "valor": CORREO_OFICIAL, "bloqueado_asesor": False,
        "error": None, "enlace_directo_prueba": None
    })


@app.post("/olvide", response_class=HTMLResponse)
async def olvide_post(request: Request, db: Session = Depends(get_db)):
    valor = ((await request.form()).get("identificador") or "").strip()
    if not valor:
        return render(request, "olvide.html", {
            "enviado": False, "valor": "", "bloqueado_asesor": False,
            "error": "Ingresa tu usuario o correo.", "enlace_directo_prueba": None
        })

    cuentas = db.query(models.Asesor).filter(
        or_(models.Asesor.username == models.normalizar(valor),
            func.lower(models.Asesor.email) == valor.lower())
    ).all()

    cuentas_superiores = [a for a in cuentas if a.rol in ROLES_SUPERIORES]
    if cuentas and not cuentas_superiores:
        jefe_nombre = cuentas[0].jefe.nombre if cuentas[0].jefe else "tu Jefe de Área, Supervisión o Arquitecto Maestro"
        return render(request, "olvide.html", {
            "enviado": False,
            "valor": valor,
            "bloqueado_asesor": True,
            "jefe_nombre": jefe_nombre,
            "error": None,
            "enlace_directo_prueba": None
        }, 403)

    ahora = datetime.utcnow()
    base = (config_correo().get("app_url") or str(request.base_url)).rstrip("/")
    enlace_generado = None
    estado_smtp = ""
    for a in cuentas_superiores:
        destino_correo = a.email or CORREO_OFICIAL
        token = secrets.token_urlsafe(32)
        a.reset_token_hash = _hash_token(token)
        a.reset_expira = ahora + timedelta(minutes=RESET_MINUTOS)
        db.commit()
        enlace_generado = f"{base}/restablecer/{token}"
        ok, detalle = enviar_correo_con_diagnostico(
            destino_correo,
            "Restablecer contraseña Directiva · Fly Logistic Command Center",
            f"Hola {a.nombre} ({a.rol}),\n\nRecibimos una solicitud para restablecer tu contraseña directiva.\n"
            f"Ingresa al siguiente enlace (válido por {RESET_MINUTOS} minutos):\n\n"
            f"{enlace_generado}\n\n"
            f"Enviado por: {CORREO_OFICIAL}"
        )
        estado_smtp = detalle

    return render(request, "olvide.html", {
        "enviado": True, "valor": valor, "bloqueado_asesor": False,
        "error": None, "enlace_directo_prueba": enlace_generado, "estado_smtp": estado_smtp
    })


@app.get("/restablecer/{token}", response_class=HTMLResponse)
def restablecer_form(token: str, request: Request, db: Session = Depends(get_db)):
    valido = asesor_por_token(db, token) is not None
    return render(request, "restablecer.html", {"valido": valido, "error": None}, 200 if valido else 400)


@app.post("/restablecer/{token}", response_class=HTMLResponse)
async def restablecer_post(token: str, request: Request, db: Session = Depends(get_db)):
    a = asesor_por_token(db, token)
    if not a:
        return render(request, "restablecer.html", {"valido": False, "error": None}, 400)
    f = await request.form()
    p1, p2 = (f.get("password") or "").strip(), (f.get("password2") or "").strip()
    if len(p1) < 8:
        return render(request, "restablecer.html", {"valido": True, "error": "Mínimo 8 caracteres."}, 400)
    if p1 != p2:
        return render(request, "restablecer.html", {"valido": True, "error": "Las contraseñas no coinciden."}, 400)
    a.password_hash = hash_clave(p1)
    a.reset_token_hash = None
    a.reset_expira = None
    db.commit()
    request.session.clear()
    request.session["flash"] = {"ok": True, "msg": "Contraseña actualizada correctamente. Ya puedes iniciar sesión."}
    return RedirectResponse("/login", status_code=303)


@app.get("/perfil", response_class=HTMLResponse)
def perfil(request: Request, user: models.Asesor = Depends(usuario_actual)):
    return render(request, "perfil.html", {"user": user, "cfg_smtp": config_correo(), "error": None})


@app.post("/perfil/correo", response_class=HTMLResponse)
async def perfil_correo(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    email = ((await request.form()).get("email") or "").strip()
    if not EMAIL_RE.fullmatch(email):
        return render(request, "perfil.html", {"user": user, "cfg_smtp": config_correo(), "error": "Formato de correo inválido."}, 400)
    user.email = email
    db.commit()
    request.session["flash"] = {"ok": True, "msg": "Correo electrónico actualizado."}
    return RedirectResponse("/perfil", status_code=303)


@app.post("/perfil/smtp", response_class=HTMLResponse)
async def perfil_smtp(request: Request, user: models.Asesor = Depends(usuario_actual)):
    if user.rol not in ROLES_SUPERIORES:
        return RedirectResponse("/perfil", status_code=303)
    f = await request.form()
    host = (f.get("smtp_host") or "mail.eagleaeronauticalsupport.com").strip()
    port = int(f.get("smtp_port") or 465)
    smtp_user = (f.get("smtp_user") or CORREO_OFICIAL).strip()
    smtp_pass = (f.get("smtp_pass") or CLAVE_CORREO_PRUEBA).strip()
    destino_prueba = (f.get("destino_prueba") or user.email or CORREO_OFICIAL).strip()

    guardar_config_correo(host, port, smtp_user, smtp_pass)

    if f.get("accion_smtp") == "probar":
        ok, detalle = enviar_correo_con_diagnostico(
            destino_prueba,
            "Prueba Real de Correo · Fly Logistic Command Center",
            f"Hola {user.nombre},\n\nEste es un correo de validación en vivo enviado desde Fly Logistic Command Center.\n"
            f"Servidor utilizado: {host}:{port}\nCuenta remitente: {smtp_user}\n"
        )
        if ok:
            request.session["flash"] = {"ok": True, "msg": f"¡Prueba SMTP exitosa! {detalle}"}
        else:
            request.session["flash"] = {"ok": False, "msg": f"Configuración guardada, pero el servidor SMTP respondió: {detalle}"}
    else:
        request.session["flash"] = {"ok": True, "msg": "Parámetros SMTP actualizados correctamente."}
    return RedirectResponse("/perfil", status_code=303)


@app.post("/perfil/clave", response_class=HTMLResponse)
async def perfil_clave(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    f = await request.form()
    actual, p1, p2 = (f.get("actual") or "").strip(), (f.get("password") or "").strip(), (f.get("password2") or "").strip()
    err = None
    if not verificar_clave(actual, user.password_hash):
        err = "La contraseña actual no es correcta."
    elif len(p1) < 8:
        err = "La nueva contraseña debe tener al menos 8 caracteres."
    elif p1 != p2:
        err = "Las contraseñas nuevas no coinciden."
    if err:
        return render(request, "perfil.html", {"user": user, "cfg_smtp": config_correo(), "error": err}, 400)
    user.password_hash = hash_clave(p1)
    user.reset_token_hash, user.reset_expira = None, None
    db.commit()
    request.session["flash"] = {"ok": True, "msg": "Contraseña actualizada exitosamente."}
    return RedirectResponse("/perfil", status_code=303)


@app.get("/gestion_usuarios", response_class=HTMLResponse)
def gestion_usuarios(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_ACCESO_USUARIOS:
        return RedirectResponse("/pedidos", status_code=303)

    if user.rol in ["Administrador", "Gerencia", "Supervisión"]:
        usuarios = db.query(models.Asesor).order_by(models.Asesor.id).all()
        jefes_disponibles = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).all()
    else:
        usuarios = db.query(models.Asesor).filter(
            or_(models.Asesor.jefe_id == user.id, models.Asesor.id == user.id)
        ).order_by(models.Asesor.id).all()
        jefes_disponibles = [user]

    return render(request, "Usuarios.html", {
        "user": user,
        "usuarios": usuarios,
        "jefes_disponibles": jefes_disponibles,
        "roles_sistema": ROLES_SISTEMA
    })


@app.post("/gestion_usuarios/accion", response_class=HTMLResponse)
async def accion_usuario(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_ACCESO_USUARIOS:
        return RedirectResponse("/pedidos", status_code=303)

    f = await request.form()
    accion = (f.get("accion") or "").strip()
    objetivo_id = (f.get("usuario_id") or "").strip()
    objetivo = db.get(models.Asesor, int(objetivo_id)) if objetivo_id.isdigit() else None

    if user.rol == "Supervisión" and accion != "modificar_clave":
        request.session["flash"] = {"ok": False, "msg": "El perfil de Supervisión únicamente está autorizado para restablecer contraseñas."}
        return RedirectResponse("/gestion_usuarios", status_code=303)

    if accion == "crear":
        nombre = (f.get("nombre") or "").strip()
        username = models.normalizar(f.get("username") or "")
        email = (f.get("email") or "").strip() or CORREO_OFICIAL
        password = (f.get("password") or "").strip() or CLAVE_CORREO_PRUEBA
        rol_elegido = (f.get("rol") or "Asesor").strip()
        jefe_id_str = (f.get("jefe_id") or "").strip()

        if len(nombre) < 3 or len(username) < 3 or len(password) < 8:
            request.session["flash"] = {"ok": False, "msg": "Verifica nombre, usuario (mín. 3 letras) y clave (mín. 8 caracteres)."}
            return RedirectResponse("/gestion_usuarios", status_code=303)

        if db.query(models.Asesor).filter(
            or_(models.Asesor.username == username, models.Asesor.nombre_norm == models.normalizar(nombre))
        ).first():
            request.session["flash"] = {"ok": False, "msg": "Ya existe un usuario con ese nombre o login."}
            return RedirectResponse("/gestion_usuarios", status_code=303)

        rol_final = rol_elegido if (user.rol in ["Administrador", "Gerencia"] and rol_elegido in ROLES_SISTEMA) else "Asesor"
        if user.rol == "Jefe de area":
            jefe_final = user.id
        else:
            jefe_final = int(jefe_id_str) if jefe_id_str.isdigit() else None

        nuevo_user = models.Asesor(
            nombre=nombre,
            nombre_norm=models.normalizar(nombre),
            username=username,
            email=email,
            rol=rol_final,
            jefe_id=jefe_final,
            password_hash=hash_clave(password)
        )
        db.add(nuevo_user)
        db.commit()
        request.session["flash"] = {"ok": True, "msg": f"Usuario «{nuevo_user.username}» ({nuevo_user.rol}) creado correctamente."}
        return RedirectResponse("/gestion_usuarios", status_code=303)

    if not objetivo or not tiene_permiso_gestion(user, objetivo):
        request.session["flash"] = {"ok": False, "msg": "No tienes permisos sobre este usuario."}
        return RedirectResponse("/gestion_usuarios", status_code=303)

    if accion == "modificar_clave":
        nueva_clave = (f.get("nueva_clave") or "").strip()
        if len(nueva_clave) < 8:
            request.session["flash"] = {"ok": False, "msg": "La nueva contraseña debe tener mínimo 8 caracteres."}
        else:
            objetivo.password_hash = hash_clave(nueva_clave)
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Contraseña de «{objetivo.nombre}» actualizada por {user.nombre}."}

    elif accion == "eliminar":
        if objetivo.id == user.id or objetivo.username == "admin":
            request.session["flash"] = {"ok": False, "msg": "La cuenta del Arquitecto Maestro / Sesión activa está protegida."}
            return RedirectResponse("/gestion_usuarios", status_code=303)
        try:
            nombre_borrado = objetivo.nombre
            uid_borrar = objetivo.id
            db.query(models.Pedido).filter(models.Pedido.asesor_id == uid_borrar).update(
                {"asesor_id": user.id}, synchronize_session=False
            )
            db.query(models.Asesor).filter(models.Asesor.jefe_id == uid_borrar).update(
                {"jefe_id": user.id}, synchronize_session=False
            )
            db.execute(text("DELETE FROM asesores WHERE id = :uid"), {"uid": uid_borrar})
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Cuenta de «{nombre_borrado}» eliminada. Sus POs fueron transferidas a tu perfil."}
        except Exception as e:
            db.rollback()
            request.session["flash"] = {"ok": False, "msg": f"No se pudo eliminar: {e}"}

    elif accion == "modificar_correo":
        nuevo_correo = (f.get("nuevo_correo") or "").strip()
        if not EMAIL_RE.fullmatch(nuevo_correo):
            request.session["flash"] = {"ok": False, "msg": "Correo electrónico inválido."}
        else:
            objetivo.email = nuevo_correo
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Correo de «{objetivo.nombre}» actualizado a {nuevo_correo}."}

    elif accion == "modificar_rol_jefe":
        nuevo_rol = (f.get("nuevo_rol") or "").strip()
        nuevo_jefe_str = (f.get("nuevo_jefe_id") or "").strip()
        if user.rol in ["Administrador", "Gerencia"] and nuevo_rol in ROLES_SISTEMA:
            objetivo.rol = nuevo_rol
            objetivo.jefe_id = int(nuevo_jefe_str) if nuevo_jefe_str.isdigit() else None
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Perfil y jerarquía de «{objetivo.nombre}» actualizados."}
        elif user.rol == "Jefe de area" and nuevo_rol in ["Asesor", "Supervisión", "Jefe de area"]:
            objetivo.rol = nuevo_rol
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Rol de «{objetivo.nombre}» actualizado a {nuevo_rol}."}

    return RedirectResponse("/gestion_usuarios", status_code=303)


def _construir_analitica(db: Session, user: models.Asesor, asesor_filtro: str = "", tipo_filtro: str = "", desde: str = "", hasta: str = ""):
    visibles = ids_asesores_visibles(db, user)
    q_asesores = db.query(models.Asesor)
    if visibles is not None:
        q_asesores = q_asesores.filter(models.Asesor.id.in_(visibles))
    asesores_equipo = q_asesores.order_by(models.Asesor.nombre).all()

    q_pedidos = db.query(models.Pedido).options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.internacional)
    )
    if visibles is not None:
        q_pedidos = q_pedidos.filter(models.Pedido.asesor_id.in_(visibles))

    if asesor_filtro.isdigit():
        q_pedidos = q_pedidos.filter(models.Pedido.asesor_id == int(asesor_filtro))
    if tipo_filtro in ("Nacional", "Internacional"):
        q_pedidos = q_pedidos.filter(models.Pedido.tipo == tipo_filtro)
    if desde:
        try:
            q_pedidos = q_pedidos.filter(models.Pedido.fecha_inicial >= datetime.strptime(desde, "%Y-%m-%d").date())
        except ValueError:
            pass
    if hasta:
        try:
            q_pedidos = q_pedidos.filter(models.Pedido.fecha_inicial <= datetime.strptime(hasta, "%Y-%m-%d").date())
        except ValueError:
            pass

    pedidos = q_pedidos.order_by(models.Pedido.fecha_inicial.desc(), models.Pedido.id.desc()).all()
    hoy = date.today()

    total_pedidos = len(pedidos)
    total_nacionales = sum(1 for p in pedidos if p.tipo == "Nacional")
    total_internacionales = sum(1 for p in pedidos if p.tipo == "Internacional")
    valor_total_ventas = sum(float(p.total) for p in pedidos)
    valor_subtotal = sum(float(p.subtotal) for p in pedidos)
    valor_total_nacionalizacion = sum(float(p.valor_nacionalizacion_dec) for p in pedidos)

    suma_probabilidad = 0
    tiempos_gestion = []
    entregas_a_tiempo = 0
    entregas_evaluadas = 0
    compras_pendientes_po = []
    items_sin_compra = 0

    ventas_mensuales_nac = {}
    ventas_mensuales_int = {}
    estado_counts = {e: 0 for e in ESTADOS}
    proveedores_top = {}

    metricas_por_asesor = {}
    for a in asesores_equipo:
        metricas_por_asesor[a.id] = {
            "id": a.id,
            "nombre": a.nombre,
            "rol": a.rol,
            "jefe": a.jefe.nombre if a.jefe else "Directo",
            "total_pos": 0,
            "entregados": 0,
            "pendientes": 0,
            "items_pendientes_compra": 0,
            "volumen": 0.0,
            "nacionalizacion": 0.0,
            "dias_acum": 0,
            "dias_count": 0,
            "prob_acum": 0,
            "sla_ok": 0,
            "sla_total": 0,
        }

    for p in pedidos:
        prob_p = PESOS_PROBABILIDAD.get(p.estado, 40)
        suma_probabilidad += prob_p
        estado_counts[p.estado] = estado_counts.get(p.estado, 0) + 1

        mes_key = p.fecha_inicial.strftime("%Y-%m")
        if p.tipo == "Nacional":
            ventas_mensuales_nac[mes_key] = ventas_mensuales_nac.get(mes_key, 0.0) + float(p.total)
        else:
            ventas_mensuales_int[mes_key] = ventas_mensuales_int.get(mes_key, 0.0) + float(p.total)

        cli_nom = p.cliente.nombre if p.cliente else "Sin Cliente"
        fecha_fin = p.fecha_entrega_cliente or hoy
        dias_gest = max(0, (fecha_fin - p.fecha_inicial).days)
        tiempos_gestion.append(dias_gest)

        info_habiles = p.dias_habiles_info
        en_tiempo = not info_habiles["en_contra"]
        if p.plazo_entrega:
            entregas_evaluadas += 1
            if en_tiempo:
                entregas_a_tiempo += 1

        faltan_compras_item = 0
        for it in p.items:
            prov_nom = it.proveedor or "Sin Proveedor"
            proveedores_top[prov_nom] = proveedores_top.get(prov_nom, 0.0) + float(it.subtotal_item)
            if not it.fecha_compra_proveedor:
                faltan_compras_item += 1
                items_sin_compra += 1

        if p.estado != "Entregado" or faltan_compras_item > 0:
            compras_pendientes_po.append({
                "id": p.id,
                "po": p.po,
                "cliente": cli_nom,
                "asesor": p.asesor.nombre if p.asesor else "",
                "asesor_email": (p.asesor.email if p.asesor and p.asesor.email else CORREO_OFICIAL),
                "tipo": p.tipo,
                "estado": p.estado,
                "dias": dias_gest,
                "plazo": p.plazo_entrega.isoformat() if p.plazo_entrega else "Sin definir",
                "estado_plazo": info_habiles["etiqueta"],
                "dias_habiles_info": info_habiles,
                "items_pendientes": faltan_compras_item,
                "valor_nacionalizacion": float(p.valor_nacionalizacion_dec),
                "total": float(p.total),
                "probabilidad": prob_p
            })

        m = metricas_por_asesor.get(p.asesor_id)
        if m:
            m["total_pos"] += 1
            m["volumen"] += float(p.total)
            m["nacionalizacion"] += float(p.valor_nacionalizacion_dec)
            m["prob_acum"] += prob_p
            m["dias_acum"] += dias_gest
            m["dias_count"] += 1
            m["items_pendientes_compra"] += faltan_compras_item
            if p.estado == "Entregado":
                m["entregados"] += 1
            else:
                m["pendientes"] += 1
            if p.plazo_entrega:
                m["sla_total"] += 1
                if en_tiempo:
                    m["sla_ok"] += 1

    probabilidad_global = round(suma_probabilidad / total_pedidos, 1) if total_pedidos else 0.0
    tiempo_promedio_dias = round(sum(tiempos_gestion) / len(tiempos_gestion), 1) if tiempos_gestion else 0.0
    cumplimiento_sla = round((entregas_a_tiempo / entregas_evaluadas) * 100, 1) if entregas_evaluadas else 100.0
    pendientes_gestion = sum(1 for p in pedidos if p.estado != "Entregado")

    tabla_asesores = []
    for aid, m in metricas_por_asesor.items():
        t_pos = m["total_pos"]
        m["tiempo_promedio"] = round(m["dias_acum"] / m["dias_count"], 1) if m["dias_count"] else 0.0
        m["probabilidad"] = round(m["prob_acum"] / t_pos, 1) if t_pos else 0.0
        m["sla_pct"] = round((m["sla_ok"] / m["sla_total"]) * 100, 1) if m["sla_total"] else 100.0
        score = (m["probabilidad"] * 0.6) + (m["sla_pct"] * 0.4) if t_pos else 0
        if t_pos == 0:
            m["nivel"] = "Sin casos"
        elif score >= 85:
            m["nivel"] = "Sobresaliente"
        elif score >= 65:
            m["nivel"] = "Óptimo"
        else:
            m["nivel"] = "En seguimiento"
        tabla_asesores.append(m)

    tabla_asesores.sort(key=lambda x: x["volumen"], reverse=True)
    todos_meses = sorted(list(set(ventas_mensuales_nac.keys()) | set(ventas_mensuales_int.keys())))
    top_prov_sorted = sorted(proveedores_top.items(), key=lambda x: x[1], reverse=True)[:6]

    return {
        "user": user,
        "asesores_equipo": asesores_equipo,
        "f_asesor": asesor_filtro,
        "f_tipo": tipo_filtro,
        "f_desde": desde,
        "f_hasta": hasta,
        "total_pedidos": total_pedidos,
        "total_nacionales": total_nacionales,
        "total_internacionales": total_internacionales,
        "valor_total_ventas": valor_total_ventas,
        "valor_subtotal": valor_subtotal,
        "valor_total_nacionalizacion": valor_total_nacionalizacion,
        "pendientes_gestion": pendientes_gestion,
        "items_sin_compra": items_sin_compra,
        "probabilidad_global": probabilidad_global,
        "tiempo_promedio_dias": tiempo_promedio_dias,
        "cumplimiento_sla": cumplimiento_sla,
        "tabla_asesores": tabla_asesores,
        "compras_pendientes_po": compras_pendientes_po,
        "meses_labels": json.dumps(todos_meses),
        "meses_nac_data": json.dumps([round(ventas_mensuales_nac.get(m, 0), 2) for m in todos_meses]),
        "meses_int_data": json.dumps([round(ventas_mensuales_int.get(m, 0), 2) for m in todos_meses]),
        "estados_labels": json.dumps(list(estado_counts.keys())),
        "estados_data": json.dumps(list(estado_counts.values())),
        "asesores_labels": json.dumps([a["nombre"] for a in tabla_asesores]),
        "asesores_vol_data": json.dumps([round(a["volumen"], 2) for a in tabla_asesores]),
        "prov_labels": json.dumps([k for k, _ in top_prov_sorted]),
        "prov_data": json.dumps([round(v, 2) for _, v in top_prov_sorted]),
    }


@app.get("/informes", response_class=HTMLResponse)
def vista_informes(request: Request, asesor_id: str = "", tipo: str = "", desde: str = "", hasta: str = "",
                   user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        return RedirectResponse("/pedidos", status_code=303)
    ctx = _construir_analitica(db, user, asesor_id, tipo, desde, hasta)
    return render(request, "Informes.html", ctx)


@app.get("/informes/excel")
def descargar_informe_gerencial(asesor_id: str = "", tipo: str = "", desde: str = "", hasta: str = "",
                                user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        return RedirectResponse("/pedidos", status_code=303)

    ctx = _construir_analitica(db, user, asesor_id, tipo, desde, hasta)
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Metricas_Asesores"
    ws1.merge_cells("A1:K1")
    ws1["A1"] = "REPORTE GERENCIAL DE HABILIDADES, PROBABILIDAD Y TIEMPOS DE GESTIÓN - FLY LOGISTIC"
    ws1["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws1["A1"].fill = PatternFill("solid", fgColor="002060")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

    enc1 = ["Asesor", "Rol", "Jefe Asignado", "Total POs", "Entregados", "Pendientes",
            "Ítems Sin Compra", "Tiempo Prom. (Días)", "Probabilidad Cierre (%)", "Nacionalización ($)", "Volumen Total ($)"]
    ws1.append(enc1)
    for c in ws1[2]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1E40AF")

    for a in ctx["tabla_asesores"]:
        ws1.append([
            a["nombre"], a["rol"], a["jefe"], a["total_pos"], a["entregados"],
            a["pendientes"], a["items_pendientes_compra"], a["tiempo_promedio"],
            a["probabilidad"], a["nacionalizacion"], a["volumen"]
        ])
    for i in range(1, 12):
        ws1.column_dimensions[get_column_letter(i)].width = 20

    ws2 = wb.create_sheet("Compras_y_Semaforo")
    enc2 = ["PO", "Cliente", "Asesor", "Tipo", "Estado Actual", "Semáforo Días Hábiles",
            "Plazo Entrega", "Ítems Pendientes Compra", "Valor Nacionalización ($)", "Probabilidad (%)", "Total PO ($)"]
    ws2.append(enc2)
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="002060")

    for cp in ctx["compras_pendientes_po"]:
        ws2.append([
            cp["po"], cp["cliente"], cp["asesor"], cp["tipo"], cp["estado"],
            cp["dias_habiles_info"]["etiqueta"], cp["plazo"], cp["items_pendientes"],
            cp["valor_nacionalizacion"], cp["probabilidad"], cp["total"]
        ])
    for i in range(1, 12):
        ws2.column_dimensions[get_column_letter(i)].width = 22

    carpeta_rep = os.path.join(BASE_DIR, "Reportes_Gerenciales")
    os.makedirs(carpeta_rep, exist_ok=True)
    ruta_rep = os.path.join(carpeta_rep, "Informe_Gerencial_FlyLogistic.xlsx")
    wb.save(ruta_rep)
    return FileResponse(ruta_rep, filename="Informe_Gerencial_FlyLogistic.xlsx", media_type=XLSX_MIME)


@app.get("/pedidos", response_class=HTMLResponse)
def mis_pedidos(request: Request, q: str = "", estado: str = "", tipo: str = "",
                user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    visibles = ids_asesores_visibles(db, user)
    query = (db.query(models.Pedido)
             .options(
                 joinedload(models.Pedido.items),
                 joinedload(models.Pedido.cliente),
                 joinedload(models.Pedido.asesor),
                 joinedload(models.Pedido.internacional)
             )
             .join(models.Cliente))

    if visibles is not None:
        query = query.filter(models.Pedido.asesor_id.in_(visibles))
    if estado in ESTADOS:
        query = query.filter(models.Pedido.estado == estado)
    if tipo in ("Nacional", "Internacional"):
        query = query.filter(models.Pedido.tipo == tipo)
    if q.strip():
        like = f"%{q.strip()}%"
        query = query.filter(models.Pedido.po.ilike(like) | models.Cliente.nombre.ilike(like))

    pedidos = query.order_by(models.Pedido.id.desc()).all()
    return render(request, "pedidos.html", {
        "user": user, "pedidos": pedidos, "q": q, "estado_f": estado, "tipo_f": tipo, "estados": ESTADOS
    })


@app.post("/pedidos/{pid}/notificar")
def notificar_po_responsable(pid: int, request: Request,
                             user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Pedido no encontrado."}
        return RedirectResponse("/pedidos", status_code=303)
    destino = (pedido.asesor.email if (pedido.asesor and pedido.asesor.email) else None) or user.email or CORREO_OFICIAL

    asunto, cuerpo = construir_cuerpo_alerta_po(pedido)
    ok, detalle = enviar_correo_con_diagnostico(destino, asunto, cuerpo)
    if ok:
        request.session["flash"] = {
            "ok": True,
            "msg": f"Correo enviado exitosamente a {destino} ({pedido.dias_habiles_info['etiqueta']})."
        }
    else:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Reporte generado en consola. Para salida real en vivo verifica la clave de {CORREO_OFICIAL} en Mi Perfil ({detalle[:90]})."
        }
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/pedidos/notificar_proximos")
def notificar_pos_proximas(request: Request, tareas: BackgroundTasks,
                           user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    visibles = ids_asesores_visibles(db, user)
    q = db.query(models.Pedido).options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.internacional)
    ).filter(models.Pedido.estado != "Entregado")
    if visibles is not None:
        q = q.filter(models.Pedido.asesor_id.in_(visibles))

    enviados = 0
    for p in q.all():
        destino = (p.asesor.email if (p.asesor and p.asesor.email) else None) or CORREO_OFICIAL
        asunto, cuerpo = construir_cuerpo_alerta_po(p)
        tareas.add_task(enviar_correo, destino, asunto, cuerpo)
        enviados += 1

    request.session["flash"] = {
        "ok": True,
        "msg": f"Se procesaron {enviados} correo(s) de seguimiento diario de POs a través de {CORREO_OFICIAL}."
    }
    return RedirectResponse("/pedidos", status_code=303)


def _form(request, db: Session, user: models.Asesor, d, modo, pedido=None, error=None, status=200):
    visibles = ids_asesores_visibles(db, user)
    q_as = db.query(models.Asesor)
    if visibles is not None:
        q_as = q_as.filter(models.Asesor.id.in_(visibles))
    asesores_disponibles = q_as.order_by(models.Asesor.nombre).all()

    return render(request, "pedido_form.html", {
        "user": user, "d": d, "modo": modo, "pedido": pedido, "error": error,
        "estados": ESTADOS, "estados_aduana": ESTADOS_ADUANA,
        "metodos_pago_opciones": METODOS_PAGO_OPCIONES,
        "unidades_medida": UNIDADES_MEDIDA,
        "asesores_disponibles": asesores_disponibles
    }, status)


@app.get("/pedidos/nuevo", response_class=HTMLResponse)
def nuevo_form(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    hoy = datetime.now().strftime("%Y-%m-%d")
    d = {k: "" for k in CAMPOS}
    d.update(
        fecha_inicial=hoy,
        tipo_pedido="Nacional",
        iva_porcentaje="19",
        valor_nacionalizacion="0.00",
        llega_click_mail="No",
        estado_general="En proceso",
        asesor_asignado_id=str(user.id),
        items=[{
            "fecha_compra_prov": "", "proveedor": "", "parte_numero": "",
            "parte_numero_alterno": "", "unidad_medida": "UND", "descripcion": "",
            "cantidad": "1", "precio_unitario": "", "iva_item": "19"
        }]
    )
    return _form(request, db, user, d, "nuevo")


@app.post("/pedidos/nuevo", response_class=HTMLResponse)
async def nuevo_post(request: Request, tareas: BackgroundTasks,
                     user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    d = datos_desde_form(await request.form())
    try:
        v = parsear(d)
        if db.query(models.Pedido).filter(func.lower(models.Pedido.po) == v["po"].lower()).first():
            raise ValueError(f"El PO «{v['po']}» ya existe en el sistema.")

        asesor_id_final = user.id
        if user.rol in ROLES_SUPERIORES and d.get("asesor_asignado_id", "").isdigit():
            asesor_id_final = int(d["asesor_asignado_id"])

        pedido = models.Pedido(po=v["po"], tipo=v["tipo"], asesor_id=asesor_id_final)
        aplicar(db, pedido, v)
        db.add(pedido)
        db.commit()
    except ValueError as e:
        db.rollback()
        return _form(request, db, user, d, "nuevo", error=str(e), status=400)
    except IntegrityError:
        db.rollback()
        return _form(request, db, user, d, "nuevo", error="El número de PO ya se encuentra registrado.", status=400)

    db.refresh(pedido)
    estado = generar_salidas(db, pedido)

    if pedido.plazo_entrega:
        destino = (pedido.asesor.email if (pedido.asesor and pedido.asesor.email) else None) or CORREO_OFICIAL
        asunto, cuerpo = construir_cuerpo_alerta_po(pedido)
        tareas.add_task(enviar_correo, destino, asunto, cuerpo)

    request.session["flash"] = {"ok": True, "msg": f"PO {pedido.po} registrada ({pedido.dias_habiles_info['etiqueta']}). {estado}."}
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/pedidos/{pid}/editar", response_class=HTMLResponse)
def editar_form(pid: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Pedido no encontrado o sin permisos de visualización."}
        return RedirectResponse("/pedidos", status_code=303)
    return _form(request, db, user, datos_desde_pedido(pedido), "editar", pedido)


@app.post("/pedidos/{pid}/editar", response_class=HTMLResponse)
async def editar_post(pid: int, request: Request,
                      user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Pedido no encontrado."}
        return RedirectResponse("/pedidos", status_code=303)
    d = datos_desde_form(await request.form())
    d["po_numero"], d["tipo_pedido"] = pedido.po, pedido.tipo
    try:
        aplicar(db, pedido, parsear(d, tipo_fijo=pedido.tipo))
        if user.rol in ROLES_SUPERIORES and d.get("asesor_asignado_id", "").isdigit():
            pedido.asesor_id = int(d["asesor_asignado_id"])
        db.commit()
    except ValueError as e:
        db.rollback()
        return _form(request, db, user, d, "editar", pedido, error=str(e), status=400)
    except IntegrityError:
        db.rollback()
        return _form(request, db, user, d, "editar", pedido, error="Error de integridad al actualizar el pedido.", status=400)

    db.refresh(pedido)
    estado = generar_salidas(db, pedido)
    request.session["flash"] = {"ok": True, "msg": f"PO {pedido.po} actualizada ({pedido.dias_habiles_info['etiqueta']}). {estado}."}
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/pedidos/{pid}/eliminar")
def eliminar_pedido(pid: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {"ok": False, "msg": "Solo Administradores, Gerencia o Jefes de área pueden eliminar POs."}
        return RedirectResponse("/pedidos", status_code=303)
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Pedido no encontrado."}
        return RedirectResponse("/pedidos", status_code=303)
    po_num = pedido.po
    db.delete(pedido)
    db.commit()
    request.session["flash"] = {"ok": True, "msg": f"PO {po_num} eliminada definitivamente."}
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/pedidos/{pid}/excel")
def descargar_excel(pid: int, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        return RedirectResponse("/pedidos", status_code=303)
    ruta, nombre = ruta_excel(pedido)
    generar_excel(ruta, pedido, pedido.internacional)
    return FileResponse(ruta, filename=nombre, media_type=XLSX_MIME)