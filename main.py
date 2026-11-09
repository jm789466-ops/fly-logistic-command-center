import asyncio
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import smtplib
import ssl
import threading
import time as time_lib
import traceback
import warnings
import zipfile
from xml.sax.saxutils import escape as _xml_escape
from copy import copy
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal, InvalidOperation
from email.message import EmailMessage

try:
    from zoneinfo import ZoneInfo
    BOGOTA_TZ = ZoneInfo("America/Bogota")
except Exception:
    BOGOTA_TZ = timezone(timedelta(hours=-5))

from fastapi import BackgroundTasks, Depends, FastAPI, Request, UploadFile, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from sqlalchemy import func, inspect, or_, text
from sqlalchemy.exc import IntegrityError, OperationalError
from sqlalchemy.orm import Session, joinedload
from starlette.middleware.sessions import SessionMiddleware

import models
from database import BASE_DIR, SessionLocal, engine
from excel_com_fm_08 import (
    generar_excel_com_fm_08,
    _escribir_bloque_po_en_hoja,
    _preparar_hoja_cuadro,
    generar_excel_consolidado_usuario,
    generar_excel_maestro_productos,
    COLORES_FILAS_PO
)

warnings.filterwarnings("ignore", message="Dialect sqlite.*Decimal")

CLAVE_MAESTRA = os.getenv("CLAVE_MAESTRA", "FLY01")
CORREO_OFICIAL = "notificaciones@flylogistic.aero"
CLAVE_CORREO_PRUEBA = "Colombia2024*"
NOMBRE_ADMIN = "Juan Esteban Martínez"

ROLES_SISTEMA = ["Administrador", "Gerencia", "Jefe de area", "Asesor", "Almacenista"]
ROLES_SUPERIORES = ["Administrador", "Gerencia", "Jefe de area"]
ROLES_ACCESO_USUARIOS = ["Administrador", "Gerencia", "Jefe de area"]
ROLES_GERENCIA_ADMIN = ["Administrador", "Gerencia"]
ROLES_OPERATIVOS_HORARIO = ["Asesor", "Almacenista"]

ESTADOS = ["En proceso", "Entregado", "Cancelado"]
ESTADOS_ADUANA = ["No aplica / Pendiente", "En revisión aduanera", "Pago de aranceles", "Liberado / Nacionalizado"]
ESTADOS_PRODUCTO = [
    "PENDIENTE DE PAGO",
    "COMPRADO",
    "SOLICITADO PARA IMPORTACIÓN",
    "LLEGADO EN FLY",
    "EN CLICKMAIL",
    "PENDIENTE DESPACHO",
    "DESPACHADO",
    "EN OPERADOR LOGÍSTICO",
    "ENTREGADO AL CLIENTE"
]
MAPA_ESTADOS_PRODUCTO_LEGACY = {
    "PENDIENTE": "PENDIENTE DE PAGO",
    "EN FLY": "LLEGADO EN FLY",
    "LLEGADO FLY": "LLEGADO EN FLY",
    "ENTREGADO": "ENTREGADO AL CLIENTE",
}
TRANSPORTADORAS_OPCIONES = ["FedEx", "Coordinadora", "DHL", "Clickmail", "Inter Rapidísimo"]
METODOS_PAGO_OPCIONES = ["Visa", "American", "PayPal", "Transferencia", "Efectivo"]
UNIDADES_MEDIDA = ["UND", "EA", "GAL", "LT", "KIT", "LB", "ROLLO", "CAJA", "MT", "PC", "SET", "CANECA 55 GAL", "GARRAFA 5 GAL"]

MAX_ARCHIVOS_PO = 10
MAX_ITEMS_COTIZADOR = 13
REGISTROS_POR_PAGINA = 10
HORA_LIMITE_ASESOR = time(17, 30)
HORA_LIMITE_GERENCIA = time(21, 0)

PESOS_PROBABILIDAD = {
    "En proceso": 25,
    "Entregado": 100,
    "Cancelado": 0,
}


def ahora_bogota_naive() -> datetime:
    return datetime.now(BOGOTA_TZ).replace(tzinfo=None)


def desbloqueo_horario_vigente(user: models.Asesor, hora_prueba: time | None = None) -> bool:
    hasta = getattr(user, "desbloqueo_horario_hasta", None)
    if not hasta:
        return False
    if hora_prueba is not None:
        referencia = datetime.combine(ahora_bogota_naive().date(), hora_prueba)
    else:
        referencia = ahora_bogota_naive()
    return referencia < hasta


def estado_horario_usuario(user) -> dict:
    ahora_dt = datetime.now(BOGOTA_TZ)
    hora_str = ahora_dt.strftime("%H:%M:%S")
    rol_u = getattr(user, "rol", "") if user else ""
    estado = {
        "aplica": False,
        "bloqueado": False,
        "permitido": True,
        "desbloqueado": False,
        "autorizacion_temporal": False,
        "hasta": "",
        "desbloqueo_hasta_str": "",
        "hora_actual": hora_str,
        "hora_limite": "24/7 Sin restricción",
        "regla_texto": "Ventana operativa: Asesores y Almacén hasta 05:30 p.m. | Jefatura y Gerencia hasta 09:00 p.m. | Administración 24/7",
    }
    if not user:
        return estado
    if rol_u == "Administrador":
        estado["hora_limite"] = "Acceso Continuo 24/7"
        estado["regla_texto"] = "Nivel Ejecutivo Administrador: Operación habilitada 24/7 | Asesores y Almacén hasta 05:30 p.m. | Jefatura y Gerencia hasta 09:00 p.m."
        return estado
    if rol_u in ("Gerencia", "Jefe de area"):
        estado["hora_limite"] = "09:00 p.m. (21:00)"
        estado["regla_texto"] = f"Nivel {rol_u}: Ventana extendida hasta las 09:00 p.m. | Asesores y Almacén hasta las 05:30 p.m."
        if os.getenv("DESACTIVAR_RESTRICCION_HORARIA", "0") != "1":
            bloq = ahora_dt.time() >= HORA_LIMITE_GERENCIA
            estado["bloqueado"] = bloq
            estado["permitido"] = not bloq
        return estado
    # Asesor o Almacenista
    estado["aplica"] = True
    estado["aplica_restriccion"] = True
    estado["hora_limite"] = "17:30"
    estado["regla_texto"] = f"Perfil {rol_u}: Horario operativo autorizado hasta las 05:30 p.m. (17:30 America/Bogota)"
    if desbloqueo_horario_vigente(user):
        hasta_fmt = user.desbloqueo_horario_hasta.strftime("%H:%M")
        estado["desbloqueado"] = True
        estado["autorizacion_temporal"] = True
        estado["hasta"] = hasta_fmt
        estado["desbloqueo_hasta_str"] = hasta_fmt
        estado["bloqueado"] = False
        estado["permitido"] = True
        return estado
    if os.getenv("DESACTIVAR_RESTRICCION_HORARIA", "0") != "1":
        bloq = ahora_dt.time() >= HORA_LIMITE_ASESOR
        estado["bloqueado"] = bloq
        estado["permitido"] = not bloq
    estado["aplica_restriccion"] = estado["aplica"]
    return estado


def validar_horario_asesor(user: models.Asesor, hora_prueba: time | None = None) -> tuple[bool, str]:
    if os.getenv("DESACTIVAR_RESTRICCION_HORARIA", "0") == "1":
        return True, ""
    if not user:
        return False, "Restricción de seguridad: Sesión no autenticada."
    ahora_col = hora_prueba if hora_prueba is not None else datetime.now(BOGOTA_TZ).time()
    if user.rol not in ROLES_SUPERIORES:
        if ahora_col >= HORA_LIMITE_ASESOR:
            if desbloqueo_horario_vigente(user, hora_prueba):
                return True, ""
            return (
                False,
                f"Restricción de seguridad (America/Bogota): El horario operativo para {user.rol} finalizó a las 17:30 (5:30 p.m.). "
                "Si cuenta con una contraseña maestra autorizada por su Jefe de Área o Gerencia, ingrésela desde el aviso de horario; de lo contrario, contacte a su Jefatura."
            )
    elif user.rol in ("Gerencia", "Jefe de area"):
        if ahora_col >= HORA_LIMITE_GERENCIA:
            return (
                False,
                "Restricción de seguridad (America/Bogota): La extensión horaria para Gerencia y Jefatura de Área finalizó a las 21:00 (9:00 p.m.)."
            )
    return True, ""


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


_drive_service = None
_drive_creds = None


def obtener_drive(forzar_refresh: bool = False):
    global _drive_service, _drive_creds
    if not GOOGLE_DRIVE_DISPONIBLE or os.getenv("OMS_DISABLE_DRIVE") == "1":
        return None
    if not os.path.exists(TOKEN_PATH):
        return None
    try:
        if _drive_creds is None or forzar_refresh:
            _drive_creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        if forzar_refresh or not _drive_creds.valid:
            if _drive_creds.expired and _drive_creds.refresh_token:
                _drive_creds.refresh(GoogleRequest())
                try:
                    with open(TOKEN_PATH, "w") as f:
                        f.write(_drive_creds.to_json())
                except Exception as w_err:
                    print(f"[GOOGLE DRIVE] Aviso persistiendo token renovado: {w_err}")
                _drive_service = build("drive", "v3", credentials=_drive_creds, cache_discovery=False)
            elif not _drive_creds.valid:
                return None
        elif _drive_service is None:
            _drive_service = build("drive", "v3", credentials=_drive_creds, cache_discovery=False)
        return _drive_service
    except Exception as exc:
        print(f"[GOOGLE DRIVE AUTH ERROR]: {exc}")
        return None


def _q(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace("'", "\\'")


def carpeta_drive(service, nombre, padre_id=None):
    q = f"name='{_q(nombre)}' and mimeType='{FOLDER_MIME}' and trashed=false"
    if padre_id:
        q += f" and '{padre_id}' in parents"
    res = service.files().list(q=q, fields="files(id, webViewLink)", pageSize=1).execute()
    if res["files"]:
        return res["files"][0]["id"], res["files"][0].get("webViewLink", "")
    meta = {"name": nombre, "mimeType": FOLDER_MIME}
    if padre_id:
        meta["parents"] = [padre_id]
    creada = service.files().create(body=meta, fields="id, webViewLink").execute()
    return creada["id"], creada.get("webViewLink", "")


def subir_archivo_a_carpeta_drive(service, ruta_local, nombre_archivo, id_carpeta, mime_type=None):
    media = MediaFileUpload(ruta_local, mimetype=mime_type, resumable=False)
    existente = service.files().list(
        q=f"name='{_q(nombre_archivo)}' and '{id_carpeta}' in parents and trashed=false",
        fields="files(id)", pageSize=1
    ).execute()["files"]
    if existente:
        return service.files().update(
            fileId=existente[0]["id"], media_body=media, fields="id, webViewLink"
        ).execute()
    return service.files().create(
        body={"name": nombre_archivo, "parents": [id_carpeta]},
        media_body=media, fields="id, webViewLink"
    ).execute()


def eliminar_archivo_de_drive(service, file_id: str) -> bool:
    if not service or not file_id:
        return False
    try:
        service.files().delete(fileId=file_id).execute()
        return True
    except Exception:
        return False


def sincronizar_po_en_drive(
    service,
    p: models.Pedido,
    ruta_excel_local: str,
    nombre_excel: str,
    ruta_doc_local: str | None = None,
    rutas_docs_locales: list[str] | None = None
):
    raiz = "Pedidos Nacionales" if p.tipo == "Nacional" else "Pedidos Internacionales"
    id_raiz, _ = carpeta_drive(service, raiz)
    nombre_asesor = p.asesor.nombre if p.asesor else "General"
    id_asesor, _ = carpeta_drive(service, nombre_asesor, id_raiz)
    id_po_folder, link_po_folder = carpeta_drive(service, f"PO_{p.po}", id_asesor)

    res_excel = subir_archivo_a_carpeta_drive(service, ruta_excel_local, nombre_excel, id_po_folder, XLSX_MIME)
    link_excel = res_excel.get("webViewLink", "")

    link_doc = p.documentacion_drive_link or ""
    if ruta_doc_local and os.path.exists(ruta_doc_local):
        nombre_doc = os.path.basename(ruta_doc_local)
        res_doc = subir_archivo_a_carpeta_drive(service, ruta_doc_local, nombre_doc, id_po_folder)
        link_doc = res_doc.get("webViewLink", "")

    carpeta_local = carpeta_po_local(p)
    for doc_obj in getattr(p, "documentos", []) or []:
        ruta_cand = os.path.join(carpeta_local, doc_obj.nombre_archivo)
        if os.path.exists(ruta_cand):
            res_d = subir_archivo_a_carpeta_drive(service, ruta_cand, doc_obj.nombre_archivo, id_po_folder)
            doc_obj.drive_file_id = res_d.get("id", "") or doc_obj.drive_file_id
            doc_obj.drive_link = res_d.get("webViewLink", "") or doc_obj.drive_link
            if not link_doc:
                link_doc = doc_obj.drive_link

    if rutas_docs_locales:
        for r_local in rutas_docs_locales:
            if r_local and os.path.exists(r_local):
                nom_d = os.path.basename(r_local)
                res_d = subir_archivo_a_carpeta_drive(service, r_local, nom_d, id_po_folder)
                if not link_doc:
                    link_doc = res_d.get("webViewLink", "")

    return link_excel, link_po_folder, link_doc


def nombre_seguro(texto: str) -> str:
    s = re.sub(r"[^\w\-. ]", "_", (texto or "").strip(), flags=re.UNICODE).strip(" .")
    return s or "sin_nombre"


def carpeta_po_local(p: models.Pedido) -> str:
    raiz_local = "Pedidos Nacionales" if p.tipo == "Nacional" else "Pedidos Internacionales"
    nombre_asesor = p.asesor.nombre if p.asesor else "General"
    carpeta = os.path.join(
        BASE_DIR,
        raiz_local,
        nombre_seguro(nombre_asesor),
        f"PO_{nombre_seguro(p.po)}"
    )
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


def ruta_excel(p: models.Pedido):
    carpeta = carpeta_po_local(p)
    codigo_po = p.po_codigo if hasattr(p, "po_codigo") else p.po
    nombre = f"Pedido_{nombre_seguro(codigo_po)}.xlsx"
    return os.path.join(carpeta, nombre), nombre


def sincronizar_resumen_documentos_po(p: models.Pedido):
    nombres = [d.nombre_archivo for d in (p.documentos or []) if d.nombre_archivo]
    p.documentacion_archivo = " | ".join(nombres) if nombres else None
    if nombres and not p.documentacion:
        p.documentacion = ", ".join(nombres)


async def guardar_archivo_documentacion(p: models.Pedido, archivo_subido: UploadFile | None) -> str | None:
    if not archivo_subido or not archivo_subido.filename or not archivo_subido.filename.strip():
        if p.documentacion_archivo:
            primer_arch = p.documentacion_archivo.split("|")[0].strip()
            posible = os.path.join(carpeta_po_local(p), primer_arch)
            if os.path.exists(posible):
                return posible
        return None
    nombre_limpio = nombre_seguro(os.path.basename(archivo_subido.filename))
    carpeta = carpeta_po_local(p)
    ruta_destino = os.path.join(carpeta, nombre_limpio)
    contenido = await archivo_subido.read()
    if contenido:
        with open(ruta_destino, "wb") as f_out:
            f_out.write(contenido)
        ya_existe = any(d.nombre_archivo == nombre_limpio for d in (p.documentos or []))
        if not ya_existe:
            p.documentos.append(models.PedidoDocumento(nombre_archivo=nombre_limpio))
        sincronizar_resumen_documentos_po(p)
        return ruta_destino
    return None


async def guardar_archivos_documentacion_multiple(
    db: Session,
    p: models.Pedido,
    archivos_subidos: list[UploadFile]
) -> list[str]:
    archivos_validos = [
        a for a in (archivos_subidos or [])
        if isinstance(a, UploadFile) and a.filename and a.filename.strip()
    ]
    if not archivos_validos:
        return []

    existentes_nombres = {d.nombre_archivo for d in (p.documentos or [])}
    nuevos_netos = [
        a for a in archivos_validos
        if nombre_seguro(os.path.basename(a.filename)) not in existentes_nombres
    ]
    if len(existentes_nombres) + len(nuevos_netos) > MAX_ARCHIVOS_PO:
        raise ValueError(
            f"Límite superado: Cada Orden de Compra permite máximo {MAX_ARCHIVOS_PO} archivos adjuntos "
            f"(actualmente tiene {len(existentes_nombres)} e intentó agregar {len(nuevos_netos)})."
        )

    carpeta = carpeta_po_local(p)
    rutas_guardadas = []
    for archivo in archivos_validos:
        nombre_limpio = nombre_seguro(os.path.basename(archivo.filename))
        ruta_destino = os.path.join(carpeta, nombre_limpio)
        contenido = await archivo.read()
        if contenido:
            with open(ruta_destino, "wb") as f_out:
                f_out.write(contenido)
            rutas_guardadas.append(ruta_destino)
            if nombre_limpio not in existentes_nombres:
                p.documentos.append(models.PedidoDocumento(nombre_archivo=nombre_limpio))
                existentes_nombres.add(nombre_limpio)

    sincronizar_resumen_documentos_po(p)
    db.flush()
    return rutas_guardadas


def generar_salidas(
    db: Session,
    p: models.Pedido,
    ruta_doc_local: str | None = None,
    rutas_docs_locales: list[str] | None = None
) -> str:
    try:
        ruta, nombre = ruta_excel(p)
        generar_excel_com_fm_08(p, ruta)
    except Exception as e:
        return f"Orden guardada, error al generar Excel: {e}"
    try:
        service = obtener_drive()
        if service:
            if not ruta_doc_local and p.documentacion_archivo:
                primer_doc = p.documentacion_archivo.split("|")[0].strip()
                posible_doc = os.path.join(carpeta_po_local(p), primer_doc)
                if os.path.exists(posible_doc):
                    ruta_doc_local = posible_doc
            link_excel, link_folder, link_doc = sincronizar_po_en_drive(
                service, p, ruta, nombre, ruta_doc_local, rutas_docs_locales
            )
            p.drive_link = link_excel or link_folder
            p.drive_folder_link = link_folder
            if link_doc:
                p.documentacion_drive_link = link_doc
            db.commit()
            return "Expediente, documentos y Excel sincronizados en Google Drive"
        return "Expediente y Excel generados correctamente"
    except Exception as e:
        return f"Excel generado localmente ({e})"


def sugerir_siguiente_po(db: Session) -> str:
    try:
        todas_pos = [r[0] for r in db.query(models.Pedido.po).all() if r[0]]
        max_num = 0
        prefijo_detectado = "PO-"
        ancho_num = 4
        for po_str in todas_pos:
            m = re.search(r"^(.*?)(\d+)$", po_str.strip())
            if m:
                num_val = int(m.group(2))
                if num_val > max_num:
                    max_num = num_val
                    prefijo_detectado = m.group(1)
                    ancho_num = len(m.group(2))
        if max_num == 0:
            return "PO-0001"
        return f"{prefijo_detectado}{str(max_num + 1).zfill(ancho_num)}"
    except Exception:
        return f"PO-{secrets.token_hex(2).upper()}"


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
    if hmac.compare_digest(clave_ingresada.strip().encode(), CLAVE_MAESTRA.encode()):
        jefe_sel = db.get(models.Asesor, jefe_id_seleccionado) if jefe_id_seleccionado else None
        return True, jefe_sel
    superiores = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).all()
    if jefe_id_seleccionado:
        jefe_sel = db.get(models.Asesor, jefe_id_seleccionado)
        if jefe_sel and jefe_sel.password_hash and verificar_clave(clave_ingresada, jefe_sel.password_hash):
            return True, jefe_sel
    for sup in superiores:
        if sup.password_hash and verificar_clave(clave_ingresada, sup.password_hash):
            return True, sup
    return False, None


def migrar_bd():
    models.Base.metadata.create_all(bind=engine)
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

        if insp.has_table("pedidos_internacionales") and insp.has_table("pedidos"):
            con.execute(text("DELETE FROM pedidos_internacionales WHERE pedido_id NOT IN (SELECT id FROM pedidos)"))
        if insp.has_table("pedido_items") and insp.has_table("pedidos"):
            con.execute(text("DELETE FROM pedido_items WHERE pedido_id NOT IN (SELECT id FROM pedidos)"))
        if insp.has_table("pedido_documentos") and insp.has_table("pedidos"):
            con.execute(text("DELETE FROM pedido_documentos WHERE pedido_id NOT IN (SELECT id FROM pedidos)"))

        if insp.has_table("asesores"):
            con.execute(text("UPDATE asesores SET rol='Asesor' WHERE rol IS NULL OR rol='' OR rol='Supervisión'"))
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
            con.execute(text("UPDATE pedidos SET transportadora_nacional = '' WHERE transportadora_nacional IS NULL"))
            con.execute(text("UPDATE pedidos SET guia_nacional = '' WHERE guia_nacional IS NULL"))
            con.execute(text("UPDATE pedidos SET estado='En proceso' WHERE estado NOT IN ('En proceso', 'Entregado', 'Cancelado')"))

        if insp.has_table("pedido_items"):
            con.execute(text("UPDATE pedido_items SET parte_numero_alterno = '' WHERE parte_numero_alterno IS NULL"))
            con.execute(text("UPDATE pedido_items SET unidad_medida = 'UND' WHERE unidad_medida IS NULL OR unidad_medida = ''"))
            con.execute(text("UPDATE pedido_items SET iva_porcentaje = 19 WHERE iva_porcentaje IS NULL"))
            con.execute(text("UPDATE pedido_items SET cobro_hazmat = 0 WHERE cobro_hazmat IS NULL"))
            con.execute(text("UPDATE pedido_items SET cobro_certificado = 0 WHERE cobro_certificado IS NULL"))
            con.execute(text("UPDATE pedido_items SET sin_iva = 0 WHERE sin_iva IS NULL"))
            con.execute(text("UPDATE pedido_items SET estado_producto = 'PENDIENTE DE PAGO' WHERE estado_producto IS NULL OR estado_producto = '' OR estado_producto = 'PENDIENTE'"))

        if insp.has_table("pedidos_internacionales"):
            con.execute(text("UPDATE pedidos_internacionales SET llega_click_mail = 'Si' WHERE fecha_click_mail IS NOT NULL AND (llega_click_mail IS NULL OR llega_click_mail = '')"))
            con.execute(text("UPDATE pedidos_internacionales SET llega_click_mail = 'No' WHERE llega_click_mail IS NULL OR llega_click_mail = ''"))
            con.execute(text("UPDATE pedidos_internacionales SET transportadora = 'FedEx' WHERE transportadora IS NULL OR transportadora = ''"))

        if insp.has_table("pedidos") and insp.has_table("pedido_documentos"):
            filas_docs = con.execute(
                text("SELECT id, documentacion_archivo, documentacion_drive_link FROM pedidos WHERE documentacion_archivo IS NOT NULL AND documentacion_archivo != ''")
            ).fetchall()
            for pid, doc_arch, doc_link in filas_docs:
                existentes_doc = {
                    r[0] for r in con.execute(
                        text("SELECT nombre_archivo FROM pedido_documentos WHERE pedido_id = :pid"),
                        {"pid": pid}
                    ).fetchall()
                }
                for nombre_ind in [x.strip() for x in (doc_arch or "").split("|") if x.strip()]:
                    if nombre_ind not in existentes_doc:
                        con.execute(
                            text("INSERT INTO pedido_documentos (pedido_id, nombre_archivo, drive_file_id, drive_link) VALUES (:pid, :nom, '', :lnk)"),
                            {"pid": pid, "nom": nombre_ind, "lnk": doc_link or ""}
                        )


def inicializar_admin_y_correo():
    ruta_principal = os.path.join(BASE_DIR, "correo_config.json")
    if not os.path.exists(ruta_principal):
        datos_smtp = {
            "host": "mail.flylogistic.aero",
            "port": 465,
            "user": CORREO_OFICIAL,
            "password": CLAVE_CORREO_PRUEBA,
            "from": f"Fly Logistic S.A.S. <{CORREO_OFICIAL}>",
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
        n_norm = models.normalizar(NOMBRE_ADMIN)
        admin = db.query(models.Asesor).filter(func.lower(models.Asesor.username) == "admin").first()
        if not admin:
            admin = db.query(models.Asesor).filter_by(nombre_norm=n_norm).first()
        if not admin:
            admin = models.Asesor(
                nombre=NOMBRE_ADMIN,
                nombre_norm=n_norm,
                username="admin",
                email=CORREO_OFICIAL,
                rol="Administrador"
            )
            db.add(admin)
        else:
            admin.nombre = NOMBRE_ADMIN
            admin.nombre_norm = n_norm
            admin.username = "admin"
            admin.rol = "Administrador"
            admin.email = admin.email or CORREO_OFICIAL
        admin.password_hash = hash_clave(CLAVE_CORREO_PRUEBA)
        db.commit()
        db.refresh(admin)

        leo_norm = models.normalizar("Leonardo Garzón")
        leonardo = db.query(models.Asesor).filter(
            or_(models.Asesor.username == "lgarzon", models.Asesor.nombre_norm == leo_norm)
        ).first()
        if not leonardo:
            leonardo = models.Asesor(
                nombre="Leonardo Garzón",
                nombre_norm=leo_norm,
                username="lgarzon",
                email="leonardo.garzon@flylogistic.aero",
                rol="Jefe de area",
                jefe_id=admin.id,
                password_hash=hash_clave("Colombia2024*")
            )
            db.add(leonardo)
        else:
            leonardo.nombre = "Leonardo Garzón"
            leonardo.nombre_norm = leo_norm
            leonardo.username = "lgarzon"
            leonardo.rol = "Jefe de area"
            leonardo.jefe_id = admin.id
            leonardo.password_hash = hash_clave("Colombia2024*")
        db.commit()
        db.refresh(leonardo)

        miguel_norm = models.normalizar("Miguel Mosquera")
        miguel = db.query(models.Asesor).filter(
            or_(models.Asesor.username == "mmosquera", models.Asesor.nombre_norm == miguel_norm)
        ).first()
        if not miguel:
            miguel = models.Asesor(
                nombre="Miguel Mosquera",
                nombre_norm=miguel_norm,
                username="mmosquera",
                email="miguel.mosquera@flylogistic.aero",
                rol="Asesor",
                jefe_id=leonardo.id,
                password_hash=hash_clave("Colombia2024*")
            )
            db.add(miguel)
        else:
            miguel.nombre = "Miguel Mosquera"
            miguel.nombre_norm = miguel_norm
            miguel.username = "mmosquera"
            miguel.rol = "Asesor"
            miguel.jefe_id = leonardo.id
            miguel.password_hash = hash_clave("Colombia2024*")
        db.commit()

        carlos_norm = models.normalizar("Carlos Mendoza")
        carlos = db.query(models.Asesor).filter(
            or_(models.Asesor.username == "cmendoza", models.Asesor.nombre_norm == carlos_norm)
        ).first()
        if not carlos:
            carlos = models.Asesor(
                nombre="Carlos Mendoza",
                nombre_norm=carlos_norm,
                username="cmendoza",
                email="carlos.mendoza@flylogistic.aero",
                rol="Asesor",
                jefe_id=leonardo.id,
                password_hash=hash_clave("Colombia2024*")
            )
            db.add(carlos)
        else:
            carlos.nombre = "Carlos Mendoza"
            carlos.nombre_norm = carlos_norm
            carlos.username = "cmendoza"
            carlos.rol = "Asesor"
            carlos.jefe_id = leonardo.id
            carlos.password_hash = hash_clave("Colombia2024*")
        db.commit()

        almacen_norm = models.normalizar("Andrés Almacén")
        almacenista = db.query(models.Asesor).filter(
            or_(models.Asesor.username == "almacen", models.Asesor.nombre_norm == almacen_norm)
        ).first()
        if not almacenista:
            almacenista = models.Asesor(
                nombre="Andrés Almacén",
                nombre_norm=almacen_norm,
                username="almacen",
                email="almacen@flylogistic.aero",
                rol="Almacenista",
                jefe_id=leonardo.id,
                password_hash=hash_clave("Colombia2024*")
            )
            db.add(almacenista)
        else:
            almacenista.nombre = "Andrés Almacén"
            almacenista.nombre_norm = almacen_norm
            almacenista.username = "almacen"
            almacenista.rol = "Almacenista"
            almacenista.jefe_id = leonardo.id
            almacenista.password_hash = hash_clave("Colombia2024*")
        db.commit()

        for old_u in db.query(models.Asesor).filter(models.Asesor.username.in_(["qa_tester", "asesor_qa"])).all():
            db.query(models.Pedido).filter(models.Pedido.asesor_id == old_u.id).update(
                {"asesor_id": leonardo.id}, synchronize_session=False
            )
            db.delete(old_u)
        db.commit()

        # Semilla inicial de Inventario Aeronáutico Multibodega (Warehouses 1, 2 y 3) si está vacío
        if db.query(models.InventarioItem).count() == 0:
            hoy_seed = date.today()
            lotes_iniciales = [
                models.InventarioItem(
                    warehouse=1,
                    codigo_barras="770999000101",
                    parte_numero="MS20995C32",
                    descripcion="Alambre de frenado acero inoxidable 0.032 in (Rollo 1 lb)",
                    medida="ROLLO",
                    marca="Malin Aerospace",
                    normativa="MIL-SPEC MS20995",
                    lote="LOT-WA1-2409",
                    dom=hoy_seed - timedelta(days=180),
                    doe=hoy_seed + timedelta(days=540),
                    existencia=24,
                    ubicacion="PAS-A1 / EST-02",
                    actualizado_por_id=admin.id,
                    actualizado_por_nombre=admin.nombre,
                ),
                models.InventarioItem(
                    warehouse=1,
                    codigo_barras="770999000102",
                    parte_numero="AEROSHELL-500",
                    descripcion="Aceite sintético para turbina AeroShell Turbine Oil 500",
                    medida="LT",
                    marca="Shell Aviation",
                    normativa="MIL-PRF-23699 / HAZMAT",
                    lote="LOT-SH-8821",
                    dom=hoy_seed - timedelta(days=300),
                    doe=hoy_seed + timedelta(days=18),
                    existencia=12,
                    ubicacion="PAS-A2 / EST-HAZ-01",
                    actualizado_por_id=admin.id,
                    actualizado_por_nombre=admin.nombre,
                ),
                models.InventarioItem(
                    warehouse=2,
                    codigo_barras="770999000201",
                    parte_numero="SKYDROL-LD4-5G",
                    descripcion="Fluido hidráulico resistente al fuego Skydrol LD-4",
                    medida="GARRAFA 5 GAL",
                    marca="Eastman Aviation",
                    normativa="HAZMAT UN1993",
                    lote="LOT-SK-5510",
                    dom=hoy_seed - timedelta(days=420),
                    doe=hoy_seed - timedelta(days=12),
                    existencia=4,
                    ubicacion="PAS-B1 / EST-HAZ-04",
                    actualizado_por_id=leonardo.id,
                    actualizado_por_nombre=leonardo.nombre,
                ),
                models.InventarioItem(
                    warehouse=2,
                    codigo_barras="770999000202",
                    parte_numero="PR-1422-B2",
                    descripcion="Sellante integral de tanques de combustible aeronáutico B-2",
                    medida="KIT",
                    marca="PPG Aerospace",
                    normativa="MIL-S-8802 / HAZMAT",
                    lote="LOT-PPG-9012",
                    dom=hoy_seed - timedelta(days=90),
                    doe=hoy_seed + timedelta(days=75),
                    existencia=15,
                    ubicacion="PAS-B2 / EST-REFRIG-01",
                    actualizado_por_id=leonardo.id,
                    actualizado_por_nombre=leonardo.nombre,
                ),
                models.InventarioItem(
                    warehouse=3,
                    codigo_barras="770999000301",
                    parte_numero="CH48108-1",
                    descripcion="Filtro de aceite de flujo completo para motor recíproco",
                    medida="UND",
                    marca="Champion Aerospace",
                    normativa="FAA-PMA",
                    lote="LOT-CH-3319",
                    dom=hoy_seed - timedelta(days=120),
                    doe=hoy_seed + timedelta(days=365),
                    existencia=30,
                    ubicacion="PAS-C1 / EST-05",
                    actualizado_por_id=miguel.id,
                    actualizado_por_nombre=miguel.nombre,
                ),
            ]
            for it_inv in lotes_iniciales:
                db.add(it_inv)
            db.commit()

    except Exception as e:
        db.rollback()
        print(f"[AVISO INICIALIZACIÓN]: {e}")
    finally:
        db.close()


models.Base.metadata.create_all(bind=engine)
migrar_bd()
inicializar_admin_y_correo()

app = FastAPI(title="Fly Logistic S.A.S. - OMS")


@app.on_event("startup")
async def sincronizacion_automatica_drive_startup():
    def _tarea_silenciosa_segundo_plano():
        try:
            with SessionLocal() as db_sesion:
                sincronizar_todo_pendiente_drive(db_sesion)
        except Exception:
            pass
    threading.Thread(target=_tarea_silenciosa_segundo_plano, daemon=True).start()


@app.exception_handler(Exception)
async def interceptor_de_errores(request: Request, exc: Exception):
    if isinstance(exc, (asyncio.CancelledError, KeyboardInterrupt)):
        raise exc
    trace = traceback.format_exc()
    print("\n[ERROR INTERNO DETECTADO]\n", trace)
    return HTMLResponse(
        content=f"""
        <div style="background:#0f172a; color:#38bdf8; padding:32px; font-family:monospace; min-height:100vh;">
            <h2 style="color:#f87171; margin-bottom:12px;">Diagnóstico Técnico del Servidor</h2>
            <p style="color:#e2e8f0; margin-bottom:16px;">Ruta: <strong>{request.url.path}</strong></p>
            <pre style="background:#1e293b; color:#4ade80; padding:20px; border-radius:8px; overflow-x:auto;">{trace}</pre>
            <p style="margin-top:20px;"><a href="/pedidos" style="color:#fff; background:#2563eb; padding:10px 18px; border-radius:6px; text-decoration:none;">Volver a Órdenes de Compra</a></p>
        </div>
        """,
        status_code=500
    )


app.add_middleware(SessionMiddleware, secret_key=cargar_secret_key(), max_age=8 * 3600, same_site="lax")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))
templates.env.filters["money"] = lambda v: f"${Decimal(str(v or 0)):,.2f}"
templates.env.globals["estado_horario"] = estado_horario_usuario


def obtener_logo_base64() -> str:
    for fname in ("logo_fly.webp", "logo_fly.webp.webp"):
        r = os.path.join(BASE_DIR, fname)
        if os.path.exists(r):
            try:
                import base64
                with open(r, "rb") as fh:
                    return base64.b64encode(fh.read()).decode("ascii")
            except Exception:
                pass
    return ""


LOGO_BASE64 = obtener_logo_base64()
templates.env.globals["LOGO_BASE64"] = LOGO_BASE64
templates.env.globals["LOGO_SRC"] = f"data:image/webp;base64,{LOGO_BASE64}" if LOGO_BASE64 else "/logo_fly.webp"


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
    if user.rol == "Jefe de area" and (objetivo.jefe_id == user.id or objetivo.id == user.id):
        return True
    return False


def ruta_inicio_rol(user: models.Asesor | None) -> str:
    if user and getattr(user, "rol", "") == "Almacenista":
        return "/inventario?warehouse=1"
    return "/pedidos"


def render(request: Request, nombre: str, ctx: dict | None = None, status: int = 200):
    ctx = dict(ctx or {})
    ctx.setdefault("flash", request.session.pop("flash", None))
    ctx["request"] = request
    ctx["roles_superiores"] = ROLES_SUPERIORES
    ctx["roles_gerencia_admin"] = ROLES_GERENCIA_ADMIN
    eh_dict = estado_horario_usuario(ctx.get("user"))
    ctx["estado_horario_actual"] = eh_dict
    ctx["estado_horario_json"] = json.dumps(eh_dict, ensure_ascii=False)
    return templates.TemplateResponse(request=request, name=nombre, context=ctx, status_code=status)


EMAIL_RE = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")
RESET_MINUTOS = 30


def config_correo() -> dict:
    cfg = {
        "host": "mail.flylogistic.aero",
        "port": 465,
        "user": CORREO_OFICIAL,
        "password": CLAVE_CORREO_PRUEBA,
        "from": f"Fly Logistic S.A.S. <{CORREO_OFICIAL}>",
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
        "from": f"Fly Logistic S.A.S. <{user_smtp.strip()}>"
    })
    for filename in ("correo_config.json", "correo_config.json.json"):
        try:
            with open(os.path.join(BASE_DIR, filename), "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception:
            pass


def enviar_correo_con_diagnostico(
    destino: str,
    asunto: str,
    cuerpo: str,
    archivos_adjuntos: list[tuple[str, bytes, str]] | None = None,
    html_cuerpo: str | None = None
) -> tuple[bool, str]:
    cfg = config_correo()
    remitente = cfg.get("user") or CORREO_OFICIAL
    clave = cfg.get("password") or CLAVE_CORREO_PRUEBA
    host_principal = (cfg.get("host") or "mail.flylogistic.aero").strip()
    puerto_principal = int(cfg.get("port") or 465)

    msg = EmailMessage()
    msg["From"] = cfg.get("from") or f"Fly Logistic S.A.S. <{remitente}>"
    msg["To"] = destino
    msg["Subject"] = asunto
    msg.set_content(cuerpo)

    if html_cuerpo:
        msg.add_alternative(html_cuerpo, subtype="html")

    if archivos_adjuntos:
        for nombre_adj, contenido_bytes, mime in archivos_adjuntos:
            if "/" in mime:
                maintype, subtype = mime.split("/", 1)
            else:
                maintype, subtype = ("application", "octet-stream")
            msg.add_attachment(
                contenido_bytes,
                maintype=maintype,
                subtype=subtype,
                filename=nombre_adj
            )

    intentos = [
        (host_principal, puerto_principal),
        ("mail.flylogistic.aero", 465),
        ("mail.flylogistic.aero", 587),
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
                return True, detalle_ok
            except Exception as e:
                errores.append(f"{host}:{port} -> {e}")

    resumen_err = " | ".join(errores[:3])
    return False, resumen_err


def enviar_correo(destino: str, asunto: str, cuerpo: str) -> bool:
    ok, _ = enviar_correo_con_diagnostico(destino, asunto, cuerpo)
    return ok


def construir_cuerpo_alerta_po(p: models.Pedido) -> tuple[str, str]:
    info_dias = p.dias_habiles_info
    estado_semaforo = "VENCIDO / RETRASADO" if info_dias["en_contra"] else "EN TIEMPO"
    moneda = p.moneda
    asunto = f"[Fly Logistic S.A.S.] Seguimiento {info_dias['etiqueta']} - Cliente: {p.cliente.nombre}"

    lineas_items = []
    for idx, it in enumerate(p.items, start=1):
        alt_txt = f" (Alterno: {it.parte_numero_alterno})" if it.parte_numero_alterno else ""
        est_prod = it.estado_producto or "PENDIENTE DE PAGO"
        imp_txt = "EXENTO (0%)" if it.es_exento_iva else f"{Decimal(str(it.iva_porcentaje)):g}%"
        adicionales_txt = ""
        if it.cobro_hazmat_dec > 0 or it.cobro_certificado_dec > 0:
            adicionales_txt = f" | Hazmat: ${it.cobro_hazmat_dec:,.2f} | Cert.: ${it.cobro_certificado_dec:,.2f}"
        lineas_items.append(
            f"  {idx}. [{est_prod}] Part Number: {it.parte_numero}{alt_txt} | {it.descripcion} | "
            f"Cant: {it.cantidad} {it.unidad_medida} | Unit ({moneda}): ${Decimal(str(it.precio_unitario)):,.2f}"
            f"{adicionales_txt} | Impuesto: {imp_txt} | Total Línea ({moneda}): ${it.total_item:,.2f}"
        )

    intl_txt = ""
    if p.tipo == "Internacional":
        i = p.internacional
        llega_cm = i.llega_click_mail if i else "No"
        fecha_cm = i.fecha_click_mail.isoformat() if (i and i.fecha_click_mail) else "No aplica"
        fecha_nac = i.fecha_nacionalizacion.isoformat() if (i and i.fecha_nacionalizacion) else "Pendiente"
        est_nac = i.estado_nacionalizacion if (i and i.estado_nacionalizacion) else "Pendiente"
        intl_txt = (
            f"\nDATOS INTERNACIONALES Y NACIONALIZACIÓN\n"
            f"- Recepción Warehouse EE. UU.: {llega_cm} (Fecha: {fecha_cm})\n"
            f"- Tracking: {(i.tracking if i and i.tracking else 'Sin asignar')}\n"
            f"- Estado de Nacionalización: {est_nac}\n"
            f"- Fecha de Nacionalización: {fecha_nac}\n"
            f"- Costo de Nacionalización (USD): ${p.valor_nacionalizacion_dec:,.2f}\n"
        )

    cuerpo = (
        f"Cordial saludo, {p.asesor.nombre}.\n\n"
        f"Se adjunta el resumen operativo de la Orden de Compra {p.po_codigo}:\n\n"
        f"- Orden de Compra: {p.po_codigo}\n"
        f"- Cliente: {p.cliente.nombre}\n"
        f"- Operación: {p.tipo} ({moneda})\n"
        f"- Estado General: {p.estado}\n"
        f"- Fecha de Inicio: {p.fecha_inicial.isoformat()}\n"
        f"- Fecha Prometida: {p.plazo_entrega.isoformat() if p.plazo_entrega else 'Sin definir'}\n"
        f"- Estado de Plazo: {info_dias['etiqueta']} ({estado_semaforo})\n"
        f"- Condición de Pago: {p.metodo_pago or 'Sin definir'}\n"
        f"{intl_txt}\n"
        f"DETALLE DE ÍTEMS\n"
        + "\n".join(lineas_items) +
        f"\n\nTOTALES DE LA ORDEN ({moneda})\n"
        f"- Subtotal (incl. Hazmat y Certificados): ${p.subtotal:,.2f}\n"
        f"- Impuestos: ${p.total_impuestos:,.2f}\n"
        f"- Nacionalización: ${p.valor_nacionalizacion_dec:,.2f}\n"
        f"- Total General ({moneda}): ${p.total:,.2f}\n\n"
        f"Fly Logistic S.A.S."
    )
    return asunto, cuerpo


def construir_cuerpo_consolidado_faltantes(pedidos: list[models.Pedido]) -> tuple[str, str]:
    asunto = f"[Fly Logistic S.A.S.] Reporte Consolidado de Faltantes de Compra ({date.today().isoformat()})"
    lineas = []
    total_faltantes = 0
    for p in pedidos:
        moneda = p.moneda
        for it in p.items:
            if not it.fecha_compra_proveedor or it.estado_producto in ("PENDIENTE DE PAGO", "PENDIENTE", "COMPRADO"):
                total_faltantes += 1
                lineas.append(
                    f"- PO: {p.po_codigo} ({p.tipo}) | Asesor: {p.asesor.nombre if p.asesor else 'N/A'} | "
                    f"Cliente: {p.cliente.nombre if p.cliente else 'N/A'} | "
                    f"SKU: {it.parte_numero} | Desc: {it.descripcion} | "
                    f"Prov: {it.proveedor} | Cant: {it.cantidad} {it.unidad_medida} | "
                    f"Subtotal ({moneda}): ${it.subtotal_item:,.2f} | Estado: {it.estado_producto}"
                )

    if not lineas:
        lineas.append("No se registran ítems pendientes de compra a proveedor en este momento.")

    cuerpo = (
        f"REPORTE EJECUTIVO CONSOLIDADO DE ÍTEMS Y COMPRAS FALTANTES\n"
        f"Fecha de corte: {datetime.now(BOGOTA_TZ).strftime('%Y-%m-%d %H:%M')}\n"
        f"Total de ítems pendientes de compra / despacho: {total_faltantes}\n\n"
        + "\n".join(lineas) +
        "\n\nFly Logistic S.A.S. - Command Center OMS"
    )
    return asunto, cuerpo


def construir_notificacion_po_opciones(p: models.Pedido, tipo_notif: str, mensaje_extra: str = "") -> tuple[str, str, str]:
    info_dias = p.dias_habiles_info
    moneda = p.moneda
    extra_txt = f"\nMENSAJE ADICIONAL DE SUPERVISIÓN:\n{mensaje_extra.strip()}\n" if mensaje_extra.strip() else ""
    extra_html = f"<div style='margin: 16px 0; padding: 12px; background: #FEF3C7; border-left: 4px solid #D97706; border-radius: 8px; font-size: 13px; color: #92400E;'><strong>Nota de Supervisión:</strong><br>{mensaje_extra.strip()}</div>" if mensaje_extra.strip() else ""

    if tipo_notif == "productos_por_comprar":
        asunto = f"[Fly Logistic S.A.S.] Alerta de Repuestos Pendientes por Comprar — Orden {p.po_codigo} ({p.cliente.nombre})"
        items_pendientes = [it for it in p.items if (not it.fecha_compra_proveedor or it.estado_producto in ("PENDIENTE", "PENDIENTE DE PAGO"))]
        if not items_pendientes:
            items_pendientes = list(p.items)

        lineas_txt = []
        filas_html = ""
        for idx, it in enumerate(items_pendientes, start=1):
            lineas_txt.append(
                f"  {idx}. Part Number: {it.parte_numero} | {it.descripcion} | "
                f"Cant: {it.cantidad} {it.unidad_medida} | Prov: {it.proveedor or 'Sin asignar'} | "
                f"Valor Unit ({moneda}): ${Decimal(str(it.precio_unitario)):,.2f} | Estado: {it.estado_producto}"
            )
            filas_html += f"<tr><td style='padding:8px;border:1px solid #E2E8F0;font-weight:bold;'>{it.parte_numero}</td><td style='padding:8px;border:1px solid #E2E8F0;'>{it.descripcion}</td><td style='padding:8px;border:1px solid #E2E8F0;text-align:center;'>{it.cantidad} {it.unidad_medida}</td><td style='padding:8px;border:1px solid #E2E8F0;'>{it.proveedor or 'Sin asignar'}</td><td style='padding:8px;border:1px solid #E2E8F0;text-align:right;font-weight:bold;'>${Decimal(str(it.precio_unitario)):,.2f}</td><td style='padding:8px;border:1px solid #E2E8F0;text-align:center;'><span style='background:#FEF3C7;color:#92400E;padding:2px 8px;border-radius:6px;font-weight:bold;font-size:11px;'>{it.estado_producto}</span></td></tr>"

        cuerpo_txt = (
            f"ALERTA OPERATIVA: REPUESTOS PENDIENTES POR COMPRAR\n\n"
            f"Orden de Compra: {p.po_codigo}\n"
            f"Cliente: {p.cliente.nombre}\n"
            f"Asesor Responsable: {p.asesor.nombre if p.asesor else 'Sin asignar'}\n"
            f"Fecha Inicial: {p.fecha_inicial.isoformat()}\n"
            f"Fecha Prometida: {p.plazo_entrega.isoformat() if p.plazo_entrega else 'Sin definir'}\n"
            f"Plazo SLA: {info_dias['etiqueta']}\n"
            f"{extra_txt}\n"
            f"REPUESTOS REQUERIDOS:\n"
            + "\n".join(lineas_txt) +
            f"\n\nPor favor gestionar la compra y suministro de estos repuestos a la brevedad.\n"
            f"Se adjunta el formato Excel COM-FM-08 de la orden.\n\n"
            f"Fly Logistic S.A.S. - Trazabilidad Aeronáutica"
        )
        cuerpo_html = f"""
        <div style="font-family: Arial, sans-serif; max-width: 680px; margin: 0 auto; color: #1E293B;">
            <div style="background: #002060; color: white; padding: 20px; border-radius: 12px 12px 0 0;">
                <h2 style="margin: 0; font-size: 18px;">Alerta de Repuestos Pendientes por Comprar</h2>
                <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Orden de Compra {p.po_codigo} &middot; {p.cliente.nombre}</p>
            </div>
            <div style="padding: 24px; border: 1px solid #E2E8F0; border-top: none; border-radius: 0 0 12px 12px; background: #FFFFFF;">
                <p style="font-size: 13px; margin-top: 0;">Se requiere la gestión inmediata de compra y confirmación de los siguientes repuestos aeronáuticos:</p>
                {extra_html}
                <table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 14px;">
                    <thead><tr style="background: #F1F5F9; color: #002060;"><th style="padding:8px;border:1px solid #E2E8F0;text-align:left;">Part Number</th><th style="padding:8px;border:1px solid #E2E8F0;text-align:left;">Descripción</th><th style="padding:8px;border:1px solid #E2E8F0;text-align:center;">Cant.</th><th style="padding:8px;border:1px solid #E2E8F0;text-align:left;">Proveedor</th><th style="padding:8px;border:1px solid #E2E8F0;text-align:right;">Precio ({moneda})</th><th style="padding:8px;border:1px solid #E2E8F0;text-align:center;">Estado</th></tr></thead>
                    <tbody>{filas_html}</tbody>
                </table>
                <div style="margin-top: 20px; padding: 12px; background: #F8FAFC; border-radius: 8px; font-size: 12px; color: #64748B;">
                    <strong>Compromiso SLA:</strong> {info_dias['etiqueta']} &middot; Fecha Prometida: {p.plazo_entrega.isoformat() if p.plazo_entrega else 'Sin definir'}<br>
                    <strong>Adjunto:</strong> Se incluye el archivo oficial de la orden COM-FM-08.
                </div>
            </div>
        </div>
        """
        return asunto, cuerpo_txt, cuerpo_html

    elif tipo_notif == "cumplimiento_sla":
        asunto = f"[Fly Logistic S.A.S.] Estado de Cumplimiento de Plazo y SLA — Orden {p.po_codigo} ({p.cliente.nombre})"
        cuerpo_txt = (
            f"SEGUIMIENTO DE PLAZO Y CUMPLIMIENTO SLA\n\n"
            f"Orden de Compra: {p.po_codigo}\n"
            f"Cliente: {p.cliente.nombre}\n"
            f"Asesor Responsable: {p.asesor.nombre if p.asesor else 'Sin asignar'}\n"
            f"Fecha de Inicio: {p.fecha_inicial.isoformat()}\n"
            f"Fecha Prometida: {p.plazo_entrega.isoformat() if p.plazo_entrega else 'Sin definir'}\n"
            f"Estado de Plazo: {info_dias['etiqueta']}\n"
            f"Estado Operativo: {p.estado}\n"
            f"Transportadora / Tracking: {p.tracking_info.get('etiqueta', 'Sin asignar')}\n"
            f"Enlace de Rastreo: {p.tracking_info.get('url', 'N/A')}\n"
            f"{extra_txt}\n"
            f"Se adjunta el formato Excel COM-FM-08 con el detalle completo de la orden.\n\n"
            f"Fly Logistic S.A.S."
        )
        cuerpo_html = f"""
        <div style="font-family: Arial, sans-serif; max-width: 680px; margin: 0 auto; color: #1E293B;">
            <div style="background: #002060; color: white; padding: 20px; border-radius: 12px 12px 0 0;">
                <h2 style="margin: 0; font-size: 18px;">Seguimiento de Plazo y Cumplimiento SLA</h2>
                <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Orden de Compra {p.po_codigo} &middot; {p.cliente.nombre}</p>
            </div>
            <div style="padding: 24px; border: 1px solid #E2E8F0; border-top: none; border-radius: 0 0 12px 12px; background: #FFFFFF;">
                <div style="padding: 14px; background: #EFF6FF; border-left: 4px solid #2563EB; border-radius: 8px; margin-bottom: 16px;">
                    <div style="font-size: 12px; text-transform: uppercase; font-weight: bold; color: #1E40AF;">Indicador de Compromiso</div>
                    <div style="font-size: 16px; font-weight: bold; color: #1E3A8A; margin-top: 2px;">{info_dias['etiqueta']}</div>
                </div>
                {extra_html}
                <ul style="font-size: 13px; line-height: 1.7; color: #334155; margin: 0; padding-left: 20px;">
                    <li><strong>Cliente:</strong> {p.cliente.nombre}</li>
                    <li><strong>Asesor a Cargo:</strong> {p.asesor.nombre if p.asesor else 'Sin asignar'}</li>
                    <li><strong>Fecha Inicial:</strong> {p.fecha_inicial.isoformat()}</li>
                    <li><strong>Fecha Prometida:</strong> {p.plazo_entrega.isoformat() if p.plazo_entrega else 'Sin definir'}</li>
                    <li><strong>Despacho y Rastreo:</strong> {p.tracking_info.get('etiqueta', 'Sin asignar')}</li>
                    <li><strong>Enlace de Rastreo:</strong> <a href="{p.tracking_info.get('url', '#')}" style="color: #2563EB;">{p.tracking_info.get('url', 'Sin enlace')}</a></li>
                </ul>
            </div>
        </div>
        """
        return asunto, cuerpo_txt, cuerpo_html

    else:
        # seguimiento_po (default general)
        asunto_def, cuerpo_def = construir_cuerpo_alerta_po(p)
        cuerpo_txt = cuerpo_def + extra_txt + "\nSe adjunta el formato Excel COM-FM-08 de la orden.\n"
        cuerpo_html = f"""
        <div style="font-family: Arial, sans-serif; max-width: 680px; margin: 0 auto; color: #1E293B;">
            <div style="background: #002060; color: white; padding: 20px; border-radius: 12px 12px 0 0;">
                <h2 style="margin: 0; font-size: 18px;">Seguimiento Operativo de Orden</h2>
                <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Orden de Compra {p.po_codigo} &middot; {p.cliente.nombre}</p>
            </div>
            <div style="padding: 24px; border: 1px solid #E2E8F0; border-top: none; border-radius: 0 0 12px 12px; background: #FFFFFF;">
                {extra_html}
                <div style="font-size: 13px; line-height: 1.6; color: #334155;">
                    <p><strong>Estado:</strong> {p.estado} &middot; <strong>Plazo SLA:</strong> {info_dias['etiqueta']}</p>
                    <p><strong>Total Orden ({moneda}):</strong> ${p.total:,.2f}</p>
                    <p><strong>Guía y Despacho:</strong> {p.tracking_info.get('etiqueta', 'Sin asignar')}</p>
                </div>
                <div style="margin-top: 18px; padding: 12px; background: #F8FAFC; border-radius: 8px; font-size: 12px; color: #64748B;">
                    Se adjunta el archivo oficial COM-FM-08 con los ítems, proveedores y soportes documentales.
                </div>
            </div>
        </div>
        """
        return asunto_def, cuerpo_txt, cuerpo_html


CAMPOS = [
    "fecha_inicial", "po_numero", "cliente", "tipo_pedido", "metodo_pago", "iva_porcentaje",
    "valor_nacionalizacion", "plazo_entrega", "fecha_entrega_cliente", "numero_remision",
    "transportadora", "numero_guia_nacional",
    "transportadora_nacional", "guia_nacional", "transportadora_intl",
    "llega_click_mail", "fecha_click_mail", "tracking", "estado_nacionalizacion",
    "fecha_nacionaliza", "documentacion", "estado_general", "observaciones",
    "motivo_cancelacion", "asesor_asignado_id", "cotizacion_origen_id"
]
CAMPOS_ITEM = [
    "fecha_compra_prov", "proveedor", "parte_numero", "parte_numero_alterno",
    "unidad_medida", "descripcion", "cantidad", "precio_unitario",
    "cobro_hazmat", "cobro_certificado", "sin_iva", "iva_item", "estado_producto"
]


def datos_desde_form(f) -> dict:
    d = {k: (f.get(k) or "").strip() if isinstance(f.get(k), str) else "" for k in CAMPOS}
    if d.get("transportadora") and not d.get("transportadora_nacional"):
        d["transportadora_nacional"] = d["transportadora"]
        d["transportadora_intl"] = d["transportadora"]
    if d.get("numero_guia_nacional") and not d.get("guia_nacional"):
        d["guia_nacional"] = d["numero_guia_nacional"]
    metodos_lista = [m.strip() for m in f.getlist("metodos_pago_check") if isinstance(m, str) and m.strip()]
    desglose_dict = {}
    if metodos_lista:
        desgloses = []
        for m in metodos_lista:
            monto_m = (f.get(f"monto_{m}") or "").strip()
            if monto_m:
                desgloses.append(f"{m}: ${monto_m}")
                desglose_dict[m] = monto_m
            else:
                desgloses.append(m)
                desglose_dict[m] = ""
        d["metodo_pago"] = ", ".join(desgloses)
    else:
        for item_p in models.parsear_desglose_pagos(d.get("metodo_pago")):
            desglose_dict[item_p["metodo"]] = item_p["monto_str"]
    d["desglose_pagos_dict"] = desglose_dict

    listas_raw = {k: [str(v).strip() for v in f.getlist(k)] for k in CAMPOS_ITEM}
    num_filas = max((len(listas_raw["parte_numero"]), len(listas_raw["descripcion"]), len(listas_raw["proveedor"])), default=0)

    items = []
    for idx in range(num_filas):
        fila_dict = {}
        for k in CAMPOS_ITEM:
            vals = listas_raw.get(k, [])
            if idx < len(vals):
                fila_dict[k] = vals[idx]
            else:
                if k == "estado_producto":
                    fila_dict[k] = "PENDIENTE DE PAGO"
                elif k == "unidad_medida":
                    fila_dict[k] = "UND"
                elif k in ("cobro_hazmat", "cobro_certificado"):
                    fila_dict[k] = "0.00"
                elif k == "sin_iva":
                    fila_dict[k] = "0"
                else:
                    fila_dict[k] = ""
        if any(fila_dict.get(c) for c in ("proveedor", "parte_numero", "descripcion", "precio_unitario")):
            items.append(fila_dict)
    d["items"] = items
    return d


def datos_desde_pedido(p: models.Pedido) -> dict:
    iso = lambda v: v.isoformat() if v else ""
    i = p.internacional
    llega_cm = i.llega_click_mail if (i and i.llega_click_mail) else ("Si" if (i and i.fecha_click_mail) else "No")
    desglose_dict = {x["metodo"]: x["monto_str"] for x in p.desglose_pagos}
    trk_info = p.tracking_info
    transp_val = getattr(p, "transportadora_nacional", "") or (getattr(i, "transportadora", "") if i else "") or trk_info["carrier"]
    guia_nac_val = getattr(p, "guia_nacional", "") or ""
    return {
        "fecha_inicial": iso(p.fecha_inicial),
        "po_numero": p.po,
        "cliente": p.cliente.nombre,
        "tipo_pedido": p.tipo,
        "moneda": p.moneda,
        "metodo_pago": p.metodo_pago or "",
        "desglose_pagos_dict": desglose_dict,
        "iva_porcentaje": f"{Decimal(str(p.iva_porcentaje)):g}",
        "valor_nacionalizacion": f"{Decimal(str(p.valor_nacionalizacion or 0)):.2f}",
        "plazo_entrega": iso(p.plazo_entrega),
        "fecha_entrega_cliente": iso(p.fecha_entrega_cliente),
        "numero_remision": p.numero_remision or "",
        "transportadora": transp_val,
        "numero_guia_nacional": guia_nac_val,
        "transportadora_nacional": getattr(p, "transportadora_nacional", "") or (trk_info["transportadora"] if p.tipo == "Nacional" else ""),
        "guia_nacional": guia_nac_val,
        "transportadora_intl": (getattr(i, "transportadora", "") if i else "") or (trk_info["transportadora"] if p.tipo == "Internacional" else "FedEx"),
        "llega_click_mail": llega_cm,
        "fecha_click_mail": iso(i.fecha_click_mail) if i else "",
        "tracking": (i.tracking or "") if i else "",
        "estado_nacionalizacion": (i.estado_nacionalizacion or "") if i else "",
        "fecha_nacionaliza": iso(i.fecha_nacionalizacion) if i else "",
        "documentacion": p.documentacion or "",
        "documentacion_archivo": p.documentacion_archivo or "",
        "documentacion_drive_link": p.documentacion_drive_link or "",
        "documentos_adjuntos": [
            {
                "id": doc.id,
                "nombre_archivo": doc.nombre_archivo,
                "drive_link": doc.drive_link or ""
            }
            for doc in p.lista_documentos
        ],
        "estado_general": p.estado,
        "observaciones": p.observaciones or "",
        "motivo_cancelacion": getattr(p, "motivo_cancelacion", "") or "",
        "asesor_asignado_id": str(p.asesor_id),
        "cotizacion_origen_id": str(p.cotizacion_origen_id or ""),
        "items": [{
            "fecha_compra_prov": iso(it.fecha_compra_proveedor),
            "proveedor": it.proveedor,
            "parte_numero": it.parte_numero,
            "parte_numero_alterno": it.parte_numero_alterno or "",
            "unidad_medida": it.unidad_medida or "UND",
            "descripcion": it.descripcion,
            "cantidad": str(it.cantidad),
            "precio_unitario": f"{Decimal(str(it.precio_unitario)):.2f}",
            "cobro_hazmat": f"{Decimal(str(getattr(it, 'cobro_hazmat', 0) or 0)):.2f}",
            "cobro_certificado": f"{Decimal(str(getattr(it, 'cobro_certificado', 0) or 0)):.2f}",
            "sin_iva": "1" if (getattr(it, "sin_iva", 0) or Decimal(str(it.iva_porcentaje or 0)) == 0) else "0",
            "iva_item": "0" if (getattr(it, "sin_iva", 0) or Decimal(str(it.iva_porcentaje or 0)) == 0) else f"{Decimal(str(it.iva_porcentaje if it.iva_porcentaje is not None else p.iva_porcentaje)):g}",
            "estado_producto": it.estado_producto or "PENDIENTE DE PAGO",
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
    po_limpio = (d.get("po_numero") or "").strip()
    if not po_limpio:
        raise ValueError("El número de PO (Orden de Compra) es obligatorio.")
    if not (d.get("cliente") or "").strip():
        raise ValueError("El nombre del cliente es obligatorio.")
    tipo = tipo_fijo or d.get("tipo_pedido")
    if tipo not in ("Nacional", "Internacional"):
        raise ValueError("Seleccione el tipo de operación (Nacional o Internacional).")

    fecha_ini = _fecha(d.get("fecha_inicial"), "Fecha inicial", obligatoria=True)
    plazo_ent = _fecha(d.get("plazo_entrega"), "Plazo de entrega prometido", obligatoria=True)
    if not (d.get("metodo_pago") or "").strip():
        raise ValueError("Seleccione al menos un método de pago.")

    iva_defecto = "19" if tipo == "Nacional" else "7"
    iva_cab_raw = (d.get("iva_porcentaje") or "").strip().lower()
    if iva_cab_raw in ("exento", "sin_iva", "sin iva", "0", "0.0"):
        iva_cabecera = Decimal("0")
    else:
        iva_cabecera = _decimal(d.get("iva_porcentaje") or iva_defecto, "IVA / TAX")
    if not 0 <= iva_cabecera <= 100:
        raise ValueError("El porcentaje de impuesto debe estar entre 0 y 100.")

    val_nac = _decimal(d.get("valor_nacionalizacion") or "0", "Valor de Nacionalización")
    if val_nac < 0:
        raise ValueError("El valor de nacionalización no puede ser negativo.")

    estado_gen = (d.get("estado_general") or "En proceso").strip()
    if estado_gen.upper() == "CANCELADO":
        estado_gen = "Cancelado"
    if estado_gen not in ESTADOS:
        raise ValueError("Estado general inválido (En proceso, Entregado o Cancelado).")
    if not d.get("items"):
        raise ValueError("Registre al menos un producto en la orden.")

    items = []
    for n, it in enumerate(d["items"], start=1):
        for k, etiqueta in (("proveedor", "proveedor"), ("parte_numero", "Part Number"), ("descripcion", "descripción")):
            if not (it.get(k) or "").strip():
                raise ValueError(f"Ítem {n}: falta completar el campo «{etiqueta}».")
        try:
            cant = int(float(it.get("cantidad", 0)))
        except ValueError:
            raise ValueError(f"Ítem {n}: cantidad inválida.")
        if cant < 1:
            raise ValueError(f"Ítem {n}: la cantidad mínima es 1.")
        if (it.get("precio_unitario") or "").strip() == "":
            raise ValueError(f"Ítem {n}: el precio unitario es obligatorio.")
        precio = _decimal(it.get("precio_unitario"), f"precio unitario del ítem {n}")
        if precio <= 0:
            raise ValueError(f"Ítem {n}: el precio unitario debe ser mayor a 0.")

        hazmat_txt = (it.get("cobro_hazmat") or "0").strip() or "0"
        cobro_hazmat = _decimal(hazmat_txt, f"cobro Hazmat del ítem {n}")
        if cobro_hazmat < 0:
            raise ValueError(f"Ítem {n}: el cobro Hazmat no puede ser negativo.")

        cert_txt = (it.get("cobro_certificado") or "0").strip() or "0"
        cobro_certificado = _decimal(cert_txt, f"cobro Certificado del ítem {n}")
        if cobro_certificado < 0:
            raise ValueError(f"Ítem {n}: el cobro Certificado no puede ser negativo.")

        sin_iva_flag = str(it.get("sin_iva") or "").strip().lower() in ("1", "true", "si", "sí", "on", "exento", "sin_iva")
        iva_it_txt = str(it.get("iva_item") if it.get("iva_item") is not None else "").strip().lower()

        if sin_iva_flag or iva_it_txt in ("0", "0.0", "0.00", "exento", "sin_iva", "sin iva"):
            iva_it = Decimal("0")
            es_sin_iva = 1
        else:
            iva_it = _decimal(iva_it_txt if iva_it_txt != "" else str(iva_cabecera), f"impuesto del ítem {n}")
            if not 0 <= iva_it <= 100:
                raise ValueError(f"Ítem {n}: el impuesto debe estar entre 0 y 100.")
            es_sin_iva = 1 if iva_it == Decimal("0") else 0

        est_prod_raw = (it.get("estado_producto") or "PENDIENTE DE PAGO").strip().upper()
        est_prod = MAPA_ESTADOS_PRODUCTO_LEGACY.get(est_prod_raw, est_prod_raw)
        if est_prod not in ESTADOS_PRODUCTO:
            est_prod = "PENDIENTE DE PAGO"

        items.append({
            "fecha_compra_proveedor": _fecha(it.get("fecha_compra_prov"), f"compra a proveedor (ítem {n})"),
            "proveedor": it["proveedor"].strip(),
            "parte_numero": it["parte_numero"].strip(),
            "parte_numero_alterno": (it.get("parte_numero_alterno") or "").strip(),
            "unidad_medida": it.get("unidad_medida") or "UND",
            "descripcion": it["descripcion"].strip(),
            "cantidad": cant,
            "precio_unitario": precio,
            "cobro_hazmat": cobro_hazmat,
            "cobro_certificado": cobro_certificado,
            "sin_iva": es_sin_iva,
            "iva_porcentaje": iva_it,
            "estado_producto": est_prod,
        })

    transp_nac = (d.get("transportadora_nacional") or d.get("transportadora") or "").strip()
    guia_nac = (d.get("guia_nacional") or d.get("numero_guia_nacional") or "").strip()

    if tipo == "Internacional":
        fecha_cm = _fecha(d.get("fecha_click_mail"), "Fecha Warehouse EE. UU.")
        llega_cm = "Si" if (d.get("llega_click_mail") == "Si" or fecha_cm is not None) else "No"
        trk_str = (d.get("tracking") or "").strip()
        transp_intl_manual = (d.get("transportadora_intl") or d.get("transportadora") or "").strip()
        det_intl = models.detectar_transportadora_tracking(
            tracking=trk_str,
            llega_click_mail=llega_cm,
            transportadora_manual=transp_intl_manual,
            tipo_operacion="Internacional"
        )
        intl_data = {
            "llega_click_mail": llega_cm,
            "fecha_click_mail": fecha_cm,
            "transportadora": det_intl["transportadora"] if det_intl["transportadora"] in TRANSPORTADORAS_OPCIONES else (transp_intl_manual or "FedEx"),
            "tracking": trk_str,
            "estado_nacionalizacion": d.get("estado_nacionalizacion") or "No aplica / Pendiente",
            "fecha_nacionalizacion": _fecha(d.get("fecha_nacionaliza"), "Fecha de Nacionalización"),
        }
    else:
        if guia_nac and not transp_nac:
            det_nac = models.detectar_transportadora_tracking(
                tracking=guia_nac,
                llega_click_mail="No",
                transportadora_manual="",
                tipo_operacion="Nacional"
            )
            transp_nac = det_nac["transportadora"]
        intl_data = None

    return {
        "po": po_limpio,
        "tipo": tipo,
        "fecha_inicial": fecha_ini,
        "cliente": d["cliente"].strip(),
        "metodo_pago": d.get("metodo_pago"),
        "iva_porcentaje": iva_cabecera,
        "valor_nacionalizacion": val_nac if tipo == "Internacional" else Decimal("0"),
        "plazo_entrega": plazo_ent,
        "fecha_entrega_cliente": _fecha(d.get("fecha_entrega_cliente"), "Entrega al cliente"),
        "numero_remision": (d.get("numero_remision") or "").strip(),
        "transportadora_nacional": transp_nac,
        "guia_nacional": guia_nac,
        "documentacion": (d.get("documentacion") or "").strip(),
        "estado": estado_gen,
        "observaciones": (d.get("observaciones") or "").strip(),
        "motivo_cancelacion": (d.get("motivo_cancelacion") or "").strip(),
        "cotizacion_origen_id": int(d["cotizacion_origen_id"]) if str(d.get("cotizacion_origen_id") or "").isdigit() else None,
        "items": items,
        "intl": intl_data,
    }


def obtener_o_crear_cliente(db: Session, nombre_raw: str) -> models.Cliente:
    nombre_limpio = (nombre_raw or "").strip()
    cli = db.query(models.Cliente).filter(models.Cliente.nombre == nombre_limpio).first()
    if cli:
        return cli
    norm_buscado = models.normalizar(nombre_limpio)
    for c in db.query(models.Cliente).all():
        if models.normalizar(c.nombre) == norm_buscado:
            return c
    cli = models.Cliente(nombre=nombre_limpio)
    db.add(cli)
    db.flush()
    return cli


def registrar_historial_pedido(
    db: Session,
    pedido: models.Pedido,
    arg3=None,
    arg4=None,
    *,
    user: models.Asesor | None = None,
    estado_anterior: str | None = None,
    estado_nuevo: str | None = None,
    accion: str = "CAMBIO_ESTADO",
    motivo: str = ""
):
    if not pedido or not pedido.id:
        return
    if isinstance(arg3, models.Asesor):
        user = arg3
    elif isinstance(arg3, str):
        estado_anterior = arg3
    if isinstance(arg4, str):
        estado_nuevo = arg4

    est_nuevo = estado_nuevo or pedido.estado or "En proceso"
    items_snapshot = [
        {
            "parte_numero": it.parte_numero,
            "descripcion": it.descripcion,
            "proveedor": it.proveedor,
            "cantidad": int(it.cantidad or 0),
            "unidad_medida": it.unidad_medida or "UND",
            "precio_unitario": str(it.precio_unitario),
            "cobro_hazmat": str(it.cobro_hazmat_dec),
            "cobro_certificado": str(it.cobro_certificado_dec),
            "sin_iva": int(getattr(it, "sin_iva", 0) or 0),
            "iva_porcentaje": str(it.iva_porcentaje),
            "total_item": str(it.total_item),
            "estado_producto": it.estado_producto,
        }
        for it in (pedido.items or [])
    ]
    reg = models.PedidoHistorial(
        pedido_id=pedido.id,
        po_numero=pedido.po,
        accion=accion,
        estado_anterior=estado_anterior or "",
        estado_nuevo=est_nuevo,
        usuario_id=user.id if user else pedido.asesor_id,
        usuario_nombre=user.nombre if user else (pedido.asesor.nombre if pedido.asesor else "Sistema"),
        motivo=motivo or getattr(pedido, "motivo_cancelacion", "") or pedido.observaciones or "",
        resumen_items=json.dumps(items_snapshot, ensure_ascii=False),
        total_snapshot=pedido.total,
        moneda=pedido.moneda,
    )
    db.add(reg)


def aplicar(db: Session, pedido: models.Pedido, v: dict, user: models.Asesor | None = None):
    estado_previo = pedido.estado if pedido.id else None
    cli = obtener_o_crear_cliente(db, v["cliente"])
    pedido.cliente_id = cli.id
    pedido.tipo = v["tipo"]
    for k in ("fecha_inicial", "metodo_pago", "iva_porcentaje", "valor_nacionalizacion", "plazo_entrega",
              "fecha_entrega_cliente", "numero_remision", "transportadora_nacional", "guia_nacional",
              "documentacion", "estado", "observaciones"):
        if k in v:
            setattr(pedido, k, v[k])
    if v.get("cotizacion_origen_id"):
        pedido.cotizacion_origen_id = v["cotizacion_origen_id"]

    ahora_dt = datetime.now()
    if v.get("estado") == "Cancelado":
        if not isinstance(pedido.fecha_cancelacion, datetime):
            pedido.fecha_cancelacion = ahora_dt
        motivo_canc = v.get("motivo_cancelacion") or "Orden cancelada con retención íntegra de historial."
        pedido.motivo_cancelacion = motivo_canc
        sello_hist = f"[HISTORIAL RETENIDO - CANCELADO {datetime.now(BOGOTA_TZ).strftime('%Y-%m-%d %H:%M')}]"
        if sello_hist not in (pedido.observaciones or ""):
            pedido.observaciones = f"{(pedido.observaciones or '').strip()} {sello_hist}".strip()

    pedido.actualizado_en = ahora_dt
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

    db.flush()
    if pedido.id and (estado_previo != pedido.estado or pedido.estado == "Cancelado"):
        registrar_historial_pedido(
            db,
            pedido,
            user=user,
            estado_anterior=estado_previo,
            estado_nuevo=pedido.estado,
            motivo=v.get("motivo_cancelacion") or f"Transición de estado a {pedido.estado}"
        )


def pedido_del_usuario(db: Session, pid: int, user: models.Asesor, bloquear: bool = False) -> models.Pedido | None:
    visibles = ids_asesores_visibles(db, user)
    q = db.query(models.Pedido)
    if bloquear:
        q = q.with_for_update()
    q = q.options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.internacional),
        joinedload(models.Pedido.documentos),
        joinedload(models.Pedido.historial)
    ).filter(models.Pedido.id == pid)
    if visibles is not None:
        q = q.filter(models.Pedido.asesor_id.in_(visibles))
    return q.first()


@app.get("/pedidos/consolidado/excel")
def descargar_consolidado_individual(
    asesor: str = "",
    asesor_id: str = "",
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    visibles = ids_asesores_visibles(db, user)
    objetivo_user = user
    filtro_asesor = asesor or asesor_id

    if user.rol in ROLES_SUPERIORES and filtro_asesor.isdigit():
        cand = db.get(models.Asesor, int(filtro_asesor))
        if cand and (visibles is None or cand.id in visibles):
            objetivo_user = cand

    q_ped = db.query(models.Pedido).options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.internacional),
        joinedload(models.Pedido.documentos)
    )

    if filtro_asesor.lower() == "todos" and user.rol in ROLES_SUPERIORES:
        if visibles is not None:
            q_ped = q_ped.filter(models.Pedido.asesor_id.in_(visibles))
        pedidos_usuario = q_ped.all()
        nombre_arch = "Consolidado_POs_Equipo_Global.xlsx"
        ruta_cons = os.path.join(BASE_DIR, "Consolidados_Usuarios", nombre_arch)
        generar_excel_consolidado_usuario(ruta_cons, pedidos_usuario, "Equipo Global")
        return FileResponse(ruta_cons, filename=nombre_arch, media_type=XLSX_MIME)
    else:
        if user.rol in ROLES_SUPERIORES and asesor_id.isdigit():
            q_ped = q_ped.filter(models.Pedido.asesor_id == objetivo_user.id)
        else:
            q_ped = q_ped.filter(models.Pedido.asesor_id == user.id)
            objetivo_user = user

        pedidos_usuario = q_ped.all()
        carpeta_cons = os.path.join(BASE_DIR, "Consolidados_Usuarios")
        os.makedirs(carpeta_cons, exist_ok=True)
        nombre_arch = f"Consolidado_POs_{nombre_seguro(objetivo_user.nombre)}.xlsx"
        ruta_cons = os.path.join(carpeta_cons, nombre_arch)
        generar_excel_consolidado_usuario(ruta_cons, pedidos_usuario, objetivo_user.nombre)
        return FileResponse(ruta_cons, filename=nombre_arch, media_type=XLSX_MIME)


@app.api_route("/healthz", methods=["GET", "HEAD"])
def health_check():
    return JSONResponse({"status": "ok", "timestamp": datetime.utcnow().isoformat()})


@app.api_route("/", methods=["GET", "HEAD"])
def inicio(request: Request, db: Session = Depends(get_db)):
    if request.method == "HEAD":
        return HTMLResponse(content="", status_code=200)
    uid = request.session.get("asesor_id")
    if not uid:
        return RedirectResponse("/login", status_code=303)
    user = db.get(models.Asesor, uid) if isinstance(db, Session) else None
    return RedirectResponse(ruta_inicio_rol(user), status_code=303)


@app.api_route("/login", methods=["GET", "HEAD"], response_class=HTMLResponse)
def login_form(request: Request, db: Session = Depends(get_db)):
    if request.method == "HEAD":
        return HTMLResponse(content="", status_code=200)
    uid = request.session.get("asesor_id")
    if uid:
        user = db.get(models.Asesor, uid) if isinstance(db, Session) else None
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)
    return render(request, "login.html", {"username": "", "nombre": "", "error": None})


@app.post("/login", response_class=HTMLResponse)
async def login_post(request: Request, db: Session = Depends(get_db)):
    f = await request.form()
    raw_user = (f.get("username") or f.get("nombre") or "").strip()
    raw_pass = (f.get("password") or f.get("clave") or "").strip()
    username_norm = models.normalizar(raw_user)
    user = db.query(models.Asesor).filter(
        or_(models.Asesor.username == username_norm, models.Asesor.nombre_norm == username_norm)
    ).first() if username_norm else None

    if not user or not verificar_clave(raw_pass, user.password_hash):
        return render(request, "login.html", {"username": raw_user, "nombre": raw_user, "error": "Credenciales de acceso inválidas."}, 401)

    request.session.clear()
    request.session["asesor_id"] = user.id
    return RedirectResponse(ruta_inicio_rol(user), status_code=303)


@app.api_route("/logout", methods=["GET", "POST"])
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


@app.get("/registro", response_class=HTMLResponse)
def registro_form(request: Request, db: Session = Depends(get_db)):
    superiores = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).order_by(models.Asesor.nombre).all()
    return render(request, "registro.html", {"v": {"email": "", "rol": "Asesor"}, "superiores": superiores, "jefes_disponibles": superiores, "error": None})


@app.post("/registro", response_class=HTMLResponse)
async def registro_post(request: Request, db: Session = Depends(get_db)):
    f = await request.form()
    v = {k: (f.get(k) or "").strip() for k in ("nombre", "username", "email", "jefe_id")}
    rol_solicitado = (f.get("rol") or "Asesor").strip()
    v["rol"] = rol_solicitado if rol_solicitado in ("Asesor", "Almacenista") else "Asesor"
    if not v["email"]:
        v["email"] = CORREO_OFICIAL
    usuario = models.normalizar(v["username"])
    password = (f.get("password") or f.get("clave") or "").strip()
    password2 = (f.get("password2") or f.get("confirmar") or "").strip()
    clave_autorizacion = (f.get("clave_maestra") or f.get("clave_admin") or "").strip()
    jefe_id_int = int(v["jefe_id"]) if v["jefe_id"].isdigit() else None
    superiores = db.query(models.Asesor).filter(models.Asesor.rol.in_(ROLES_SUPERIORES)).order_by(models.Asesor.nombre).all()

    def error(msg, status=400):
        return render(request, "registro.html", {"v": v, "superiores": superiores, "jefes_disponibles": superiores, "error": msg}, status)

    autorizado, sup_encontrado = buscar_autorizador_por_clave(db, clave_autorizacion, jefe_id_int)
    if not autorizado:
        return error("La clave de autorización ingresada no es válida.", 403)
    if len(v["nombre"]) < 3:
        return error("Ingrese el nombre y apellido completo.")
    if not v["username"]:
        return error("El nombre de usuario es obligatorio.")
    if not EMAIL_RE.fullmatch(v["email"]):
        return error("Ingrese un correo electrónico corporativo válido.")
    if not 3 <= len(usuario) <= 100:
        return error("El usuario debe contener entre 3 y 100 caracteres.")
    if len(password) < 8:
        return error("La contraseña debe tener mínimo 8 caracteres.")
    if password != password2:
        return error("Las contraseñas no coinciden.")
    if db.query(models.Asesor).filter_by(username=usuario).first():
        return error("El nombre de usuario ya se encuentra registrado.", 409)

    asesor = db.query(models.Asesor).filter_by(nombre_norm=models.normalizar(v["nombre"])).first()
    if asesor and asesor.password_hash:
        return error("Ya existe un perfil activo para este colaborador.", 409)
    if not asesor:
        asesor = models.Asesor(nombre=v["nombre"], nombre_norm=models.normalizar(v["nombre"]), rol=v["rol"])
        db.add(asesor)

    asesor.username = usuario
    asesor.password_hash = hash_clave(password)
    asesor.email = v["email"]
    asesor.rol = v["rol"]
    asesor.jefe_id = jefe_id_int or (sup_encontrado.id if sup_encontrado else None)
    db.commit()

    request.session.clear()
    request.session["asesor_id"] = asesor.id
    request.session["flash"] = {"ok": True, "msg": f"Perfil corporativo ({asesor.rol}) habilitado para {asesor.nombre}."}
    return RedirectResponse(ruta_inicio_rol(asesor), status_code=303)


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
        "enviado": False, "valor": "", "bloqueado_asesor": False,
        "error": None, "enlace_directo_prueba": None
    })


@app.post("/olvide", response_class=HTMLResponse)
async def olvide_post(request: Request, db: Session = Depends(get_db)):
    valor = ((await request.form()).get("identificador") or "").strip()
    if not valor:
        return render(request, "olvide.html", {
            "enviado": False, "valor": "", "bloqueado_asesor": False,
            "error": "Ingrese su usuario o correo electrónico.", "enlace_directo_prueba": None
        })

    cuentas = db.query(models.Asesor).filter(
        or_(models.Asesor.username == models.normalizar(valor),
            func.lower(models.Asesor.email) == valor.lower())
    ).all()

    cuentas_superiores = [a for a in cuentas if a.rol in ROLES_SUPERIORES]
    if cuentas and not cuentas_superiores:
        jefe_nombre = cuentas[0].jefe.nombre if cuentas[0].jefe else "Jefatura de Área o Administración"
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
            "Restablecimiento de Credenciales - Fly Logistic S.A.S.",
            f"Hola {a.nombre},\n\nSe ha solicitado el restablecimiento de su contraseña.\n"
            f"Enlace de acceso (válido por {RESET_MINUTOS} minutos):\n\n"
            f"{enlace_generado}\n\n"
            f"Fly Logistic S.A.S."
        )
        estado_smtp = detalle

    return render(request, "olvide.html", {
        "enviado": True, "valor": valor, "bloqueado_asesor": False,
        "error": None, "enlace_directo_prueba": enlace_generado, "estado_smtp": estado_smtp
    })


@app.get("/restablecer/{token}", response_class=HTMLResponse)
def restablecer_form(token: str, request: Request, db: Session = Depends(get_db)):
    valido = asesor_por_token(db, token) is not None
    return render(request, "restablecer.html", {"token": token, "valido": valido, "invalido": not valido, "error": None}, 200 if valido else 400)


@app.post("/restablecer/{token}", response_class=HTMLResponse)
async def restablecer_post(token: str, request: Request, db: Session = Depends(get_db)):
    a = asesor_por_token(db, token)
    if not a:
        return render(request, "restablecer.html", {"token": token, "valido": False, "invalido": True, "error": None}, 400)
    f = await request.form()
    p1 = (f.get("password") or f.get("clave") or "").strip()
    p2 = (f.get("password2") or f.get("confirmar") or "").strip()
    if len(p1) < 8:
        return render(request, "restablecer.html", {"token": token, "valido": True, "invalido": False, "error": "Mínimo 8 caracteres."}, 400)
    if p1 != p2:
        return render(request, "restablecer.html", {"token": token, "valido": True, "invalido": False, "error": "Las contraseñas no coinciden."}, 400)
    a.password_hash = hash_clave(p1)
    a.reset_token_hash = None
    a.reset_expira = None
    db.commit()
    request.session.clear()
    request.session["flash"] = {"ok": True, "msg": "Contraseña actualizada correctamente."}
    return RedirectResponse("/login", status_code=303)


ALFABETO_CLAVE_HORARIO = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
VIGENCIAS_CLAVE_HORARIO = [1, 2, 3, 4]
MAX_INTENTOS_DESBLOQUEO = 5
BLOQUEO_INTENTOS_SEGUNDOS = 300


def hash_codigo_horario(codigo: str) -> str:
    normalizado = re.sub(r"[^A-Z0-9]", "", (codigo or "").upper())
    return hashlib.sha256(f"fly-horario:{normalizado}".encode()).hexdigest()


def generar_codigo_horario() -> str:
    parte1 = "".join(secrets.choice(ALFABETO_CLAVE_HORARIO) for _ in range(4))
    parte2 = "".join(secrets.choice(ALFABETO_CLAVE_HORARIO) for _ in range(4))
    return f"FLY-{parte1}-{parte2}"


def listar_claves_horario_para_perfil(db: Session, user: models.Asesor) -> list:
    consulta = db.query(models.ClaveHorarioExtendido)
    if user.rol not in ROLES_GERENCIA_ADMIN:
        consulta = consulta.filter(models.ClaveHorarioExtendido.generado_por_id == user.id)
    claves = consulta.order_by(models.ClaveHorarioExtendido.id.desc()).limit(15).all()
    ahora = ahora_bogota_naive()
    resultado = []
    for c in claves:
        if c.revocada:
            estado = "Revocada"
        elif c.expira_en <= ahora:
            estado = "Vencida"
        else:
            estado = "Activa"
        resultado.append({
            "id": c.id,
            "sufijo": c.codigo_sufijo,
            "destino": c.asesor_destino.nombre if c.asesor_destino else "Cualquier asesor",
            "generado_por": c.generado_por.nombre if c.generado_por else "",
            "motivo": c.motivo or "",
            "creada": c.creado_en.strftime("%Y-%m-%d %H:%M") if c.creado_en else "",
            "expira": c.expira_en.strftime("%Y-%m-%d %H:%M"),
            "estado": estado,
            "usos": c.usos or 0,
            "ultimo_uso_por": c.ultimo_uso_por or "",
        })
    return resultado


@app.get("/perfil", response_class=HTMLResponse)
def perfil(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    ctx = {"user": user, "cfg_smtp": config_correo(), "error": None}
    if user.rol in ROLES_SUPERIORES and isinstance(db, Session):
        ctx["claves_horario"] = listar_claves_horario_para_perfil(db, user)
        consulta_asesores = db.query(models.Asesor).filter(models.Asesor.rol.in_(["Asesor", "Almacenista"]))
        if user.rol not in ROLES_GERENCIA_ADMIN:
            consulta_asesores = consulta_asesores.filter(models.Asesor.jefe_id == user.id)
        ctx["asesores_para_clave"] = consulta_asesores.order_by(models.Asesor.nombre).all()
        ctx["clave_horario_nueva"] = request.session.pop("clave_horario_nueva", None)
        ctx["vigencias_clave"] = VIGENCIAS_CLAVE_HORARIO
    return render(request, "perfil.html", ctx)


@app.post("/perfil/clave_horario/generar")
async def generar_clave_horario(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {"ok": False, "msg": "Solo Administración, Gerencia y Jefes de Área pueden generar contraseñas maestras."}
        return RedirectResponse("/perfil", status_code=303)
    f = await request.form()
    try:
        horas = int(f.get("vigencia_horas") or 2)
    except ValueError:
        horas = 2
    horas = horas if horas in VIGENCIAS_CLAVE_HORARIO else 2
    destino_raw = (f.get("asesor_destino_id") or "").strip()
    destino = None
    if destino_raw:
        destino = db.get(models.Asesor, int(destino_raw)) if destino_raw.isdigit() else None
        if not destino or destino.rol not in ("Asesor", "Almacenista") or not tiene_permiso_gestion(user, destino):
            request.session["flash"] = {"ok": False, "msg": "El colaborador seleccionado no pertenece a su equipo o no es válido."}
            return RedirectResponse("/perfil", status_code=303)
    ahora = ahora_bogota_naive()
    fin_de_dia = ahora.replace(hour=23, minute=59, second=0, microsecond=0)
    expira = min(ahora + timedelta(hours=horas), fin_de_dia)
    codigo = ""
    for _ in range(10):
        candidato = generar_codigo_horario()
        if not db.query(models.ClaveHorarioExtendido).filter(
            models.ClaveHorarioExtendido.codigo_hash == hash_codigo_horario(candidato)
        ).first():
            codigo = candidato
            break
    if not codigo:
        request.session["flash"] = {"ok": False, "msg": "No fue posible generar un código único. Intente nuevamente."}
        return RedirectResponse("/perfil", status_code=303)
    clave = models.ClaveHorarioExtendido(
        codigo_hash=hash_codigo_horario(codigo),
        codigo_sufijo=codigo[-4:],
        generado_por_id=user.id,
        asesor_destino_id=destino.id if destino else None,
        motivo=(f.get("motivo") or "").strip()[:255],
        creado_en=datetime.utcnow(),
        expira_en=expira,
    )
    db.add(clave)
    db.commit()
    request.session["clave_horario_nueva"] = {
        "codigo": codigo,
        "expira": expira.strftime("%Y-%m-%d %H:%M"),
        "destino": destino.nombre if destino else "Cualquier asesor o almacenista",
    }
    return RedirectResponse("/perfil", status_code=303)


@app.post("/perfil/clave_horario/{clave_id}/revocar")
def revocar_clave_horario(clave_id: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    clave = (
        db.query(models.ClaveHorarioExtendido)
        .filter(models.ClaveHorarioExtendido.id == clave_id)
        .with_for_update()
        .first()
    )
    if not clave or user.rol not in ROLES_SUPERIORES or (user.rol not in ROLES_GERENCIA_ADMIN and clave.generado_por_id != user.id):
        request.session["flash"] = {"ok": False, "msg": "No tiene permisos para revocar esta contraseña maestra."}
        return RedirectResponse("/perfil", status_code=303)
    clave.revocada = True
    afectados = (
        db.query(models.Asesor)
        .filter(models.Asesor.desbloqueo_clave_id == clave.id)
        .all()
    )
    for a in afectados:
        a.desbloqueo_horario_hasta = None
        a.desbloqueo_clave_id = None
    db.commit()
    request.session["flash"] = {"ok": True, "msg": f"La contraseña maestra terminada en {clave.codigo_sufijo} fue revocada y las autorizaciones asociadas dejaron de estar vigentes."}
    return RedirectResponse("/perfil", status_code=303)


@app.post("/horario/desbloquear")
async def desbloquear_horario(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    f = await request.form()
    destino = (f.get("next") or ruta_inicio_rol(user)).strip()
    if not destino.startswith("/") or destino.startswith("//"):
        destino = ruta_inicio_rol(user)
    if user.rol == "Almacenista" and not destino.startswith("/inventario"):
        destino = "/inventario?warehouse=1"
    if user.rol in ROLES_SUPERIORES:
        request.session["flash"] = {"ok": True, "msg": "Su perfil no requiere contraseña maestra para operar."}
        return RedirectResponse(destino, status_code=303)

    ahora_ts = datetime.now(timezone.utc).timestamp()
    bloqueo_hasta = float(request.session.get("desbloqueo_bloqueado_hasta") or 0)
    if bloqueo_hasta > ahora_ts:
        minutos = int((bloqueo_hasta - ahora_ts) // 60) + 1
        request.session["flash"] = {"ok": False, "msg": f"Demasiados intentos fallidos. Espere {minutos} minuto(s) antes de intentarlo de nuevo."}
        return RedirectResponse(destino, status_code=303)

    codigo = (f.get("clave_horario") or "").strip()
    clave = None
    if codigo:
        clave = (
            db.query(models.ClaveHorarioExtendido)
            .filter(models.ClaveHorarioExtendido.codigo_hash == hash_codigo_horario(codigo))
            .with_for_update()
            .first()
        )
    ahora = ahora_bogota_naive()
    valida = bool(
        clave
        and not clave.revocada
        and clave.expira_en > ahora
        and (clave.asesor_destino_id is None or clave.asesor_destino_id == user.id)
    )
    if not valida:
        intentos = int(request.session.get("desbloqueo_intentos") or 0) + 1
        if intentos >= MAX_INTENTOS_DESBLOQUEO:
            request.session["desbloqueo_intentos"] = 0
            request.session["desbloqueo_bloqueado_hasta"] = ahora_ts + BLOQUEO_INTENTOS_SEGUNDOS
            msg = "Se alcanzó el máximo de intentos. El ingreso de contraseñas maestras quedó pausado por 5 minutos."
        else:
            request.session["desbloqueo_intentos"] = intentos
            msg = f"La contraseña maestra no es válida, expiró o no fue emitida para su usuario. Intento {intentos} de {MAX_INTENTOS_DESBLOQUEO}."
        request.session["flash"] = {"ok": False, "msg": msg}
        return RedirectResponse(destino, status_code=303)

    request.session["desbloqueo_intentos"] = 0
    user.desbloqueo_horario_hasta = clave.expira_en
    user.desbloqueo_clave_id = clave.id
    clave.usos = (clave.usos or 0) + 1
    clave.ultimo_uso_en = ahora
    clave.ultimo_uso_por = user.nombre
    db.commit()
    request.session["flash"] = {
        "ok": True,
        "msg": f"Autorización especial activada. Podrá operar fuera de horario hasta las {clave.expira_en.strftime('%H:%M')} (America/Bogota).",
    }
    return RedirectResponse(destino, status_code=303)


@app.post("/perfil/correo", response_class=HTMLResponse)
async def perfil_correo(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    email = ((await request.form()).get("email") or "").strip()
    if not EMAIL_RE.fullmatch(email):
        return render(request, "perfil.html", {"user": user, "cfg_smtp": config_correo(), "error": "Formato de correo inválido."}, 400)
    user.email = email
    db.commit()
    request.session["flash"] = {"ok": True, "msg": "Correo electrónico actualizado."}
    return RedirectResponse("/perfil", status_code=303)


@app.post("/perfil/datos", response_class=HTMLResponse)
async def perfil_datos(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    f = await request.form()
    email = (f.get("email") or "").strip()
    telefono = (f.get("telefono") or "").strip()
    if not EMAIL_RE.fullmatch(email):
        return render(request, "perfil.html", {"user": user, "cfg_smtp": config_correo(), "error": "Formato de correo inválido."}, 400)
    user.email = email
    user.telefono = telefono
    db.commit()
    request.session["flash"] = {"ok": True, "msg": "Datos de contacto actualizados."}
    return RedirectResponse("/perfil", status_code=303)


@app.post("/perfil/smtp", response_class=HTMLResponse)
async def perfil_smtp(request: Request, user: models.Asesor = Depends(usuario_actual)):
    if user.rol != "Administrador":
        return RedirectResponse("/perfil", status_code=303)
    f = await request.form()
    host = (f.get("smtp_host") or "mail.flylogistic.aero").strip()
    port = int(f.get("smtp_port") or 465)
    smtp_user = (f.get("smtp_user") or CORREO_OFICIAL).strip()
    smtp_pass = (f.get("smtp_pass") or CLAVE_CORREO_PRUEBA).strip()
    destino_prueba = (f.get("destino_prueba") or user.email or CORREO_OFICIAL).strip()

    guardar_config_correo(host, port, smtp_user, smtp_pass)

    if f.get("accion_smtp") == "probar":
        ok, detalle = enviar_correo_con_diagnostico(
            destino_prueba,
            "Prueba de Servicio SMTP - Fly Logistic S.A.S.",
            f"Hola {user.nombre},\n\nConfirmación de salida de correo desde el sistema OMS.\n"
            f"Servidor: {host}:{port}\nCuenta: {smtp_user}\n"
        )
        if ok:
            request.session["flash"] = {"ok": True, "msg": f"Conexión SMTP verificada: {detalle}"}
        else:
            request.session["flash"] = {"ok": False, "msg": f"Parámetros guardados. Respuesta del servidor: {detalle}"}
    else:
        request.session["flash"] = {"ok": True, "msg": "Configuración SMTP actualizada."}
    return RedirectResponse("/perfil", status_code=303)


@app.post("/perfil/clave", response_class=HTMLResponse)
async def perfil_clave(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    f = await request.form()
    actual, p1, p2 = (f.get("actual") or "").strip(), (f.get("password") or "").strip(), (f.get("password2") or "").strip()
    err = None
    if not verificar_clave(actual, user.password_hash):
        err = "La contraseña actual es incorrecta."
    elif len(p1) < 8:
        err = "La nueva contraseña debe tener mínimo 8 caracteres."
    elif p1 != p2:
        err = "Las contraseñas nuevas no coinciden."
    if err:
        return render(request, "perfil.html", {"user": user, "cfg_smtp": config_correo(), "error": err}, 400)
    user.password_hash = hash_clave(p1)
    user.reset_token_hash, user.reset_expira = None, None
    db.commit()
    request.session["flash"] = {"ok": True, "msg": "Contraseña actualizada."}
    return RedirectResponse("/perfil", status_code=303)


@app.get("/gestion_usuarios", response_class=HTMLResponse)
def gestion_usuarios(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_ACCESO_USUARIOS:
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)

    if user.rol in ["Administrador", "Gerencia"]:
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
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)

    f = await request.form()
    accion = (f.get("accion") or "").strip()
    objetivo_id = (f.get("usuario_id") or "").strip()
    objetivo = (
        db.query(models.Asesor).with_for_update().filter(models.Asesor.id == int(objetivo_id)).first()
        if objetivo_id.isdigit() else None
    )

    if accion == "crear":
        nombre = (f.get("nombre") or "").strip()
        username = models.normalizar(f.get("username") or "")
        email = (f.get("email") or "").strip() or CORREO_OFICIAL
        password = (f.get("password") or "").strip() or CLAVE_CORREO_PRUEBA
        rol_elegido = (f.get("rol") or "Asesor").strip()
        jefe_id_str = (f.get("jefe_id") or "").strip()

        if len(nombre) < 3 or len(username) < 3 or len(password) < 8:
            request.session["flash"] = {"ok": False, "msg": "Verifique nombre, usuario (mín. 3 caracteres) y contraseña (mín. 8 caracteres)."}
            return RedirectResponse("/gestion_usuarios", status_code=303)

        if db.query(models.Asesor).filter(
            or_(models.Asesor.username == username, models.Asesor.nombre_norm == models.normalizar(nombre))
        ).first():
            request.session["flash"] = {"ok": False, "msg": "Ya existe un colaborador registrado con ese nombre o usuario."}
            return RedirectResponse("/gestion_usuarios", status_code=303)

        if user.rol in ["Administrador", "Gerencia"] and rol_elegido in ROLES_SISTEMA:
            rol_final = rol_elegido
        elif user.rol == "Jefe de area" and rol_elegido in ("Asesor", "Almacenista"):
            rol_final = rol_elegido
        else:
            rol_final = "Asesor"
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
        request.session["flash"] = {"ok": True, "msg": f"Colaborador {nuevo_user.nombre} ({nuevo_user.rol}) registrado."}
        return RedirectResponse("/gestion_usuarios", status_code=303)

    if not objetivo or not tiene_permiso_gestion(user, objetivo):
        request.session["flash"] = {"ok": False, "msg": "Permisos insuficientes para modificar este perfil."}
        return RedirectResponse("/gestion_usuarios", status_code=303)

    if accion in ("modificar_clave", "cambiar_clave"):
        nueva_clave = (f.get("nueva_clave") or "").strip()
        if len(nueva_clave) < 8:
            request.session["flash"] = {"ok": False, "msg": "La contraseña debe tener al menos 8 caracteres."}
        else:
            objetivo.password_hash = hash_clave(nueva_clave)
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Contraseña actualizada para {objetivo.nombre}."}

    elif accion == "modificar_identidad":
        nuevo_nom = (f.get("nuevo_nombre") or "").strip()
        nuevo_user = models.normalizar(f.get("nuevo_username") or "")
        if len(nuevo_nom) < 3 or len(nuevo_user) < 3:
            request.session["flash"] = {"ok": False, "msg": "Nombre y usuario deben tener al menos 3 caracteres."}
        else:
            existente = db.query(models.Asesor).filter(
                models.Asesor.id != objetivo.id,
                or_(models.Asesor.username == nuevo_user, models.Asesor.nombre_norm == models.normalizar(nuevo_nom))
            ).first()
            if existente:
                request.session["flash"] = {"ok": False, "msg": "Ya existe otro colaborador con ese nombre o usuario."}
            else:
                objetivo.nombre = nuevo_nom
                objetivo.nombre_norm = models.normalizar(nuevo_nom)
                objetivo.username = nuevo_user
                db.commit()
                request.session["flash"] = {"ok": True, "msg": f"Identidad actualizada para {objetivo.nombre}."}

    elif accion == "eliminar":
        if objetivo.id == user.id or objetivo.username == "admin":
            request.session["flash"] = {"ok": False, "msg": "La cuenta administrativa principal no puede ser eliminada."}
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
            request.session["flash"] = {"ok": True, "msg": f"Usuario {nombre_borrado} dado de baja. Sus órdenes fueron reasignadas."}
        except Exception as e:
            db.rollback()
            request.session["flash"] = {"ok": False, "msg": f"Error al eliminar cuenta: {e}"}

    elif accion == "modificar_correo":
        nuevo_correo = (f.get("nuevo_correo") or "").strip()
        if not EMAIL_RE.fullmatch(nuevo_correo):
            request.session["flash"] = {"ok": False, "msg": "Formato de correo electrónico inválido."}
        else:
            objetivo.email = nuevo_correo
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Correo actualizado para {objetivo.nombre}."}

    elif accion == "modificar_rol_jefe":
        nuevo_rol = (f.get("nuevo_rol") or "").strip()
        nuevo_jefe_str = (f.get("nuevo_jefe_id") or "").strip()
        if user.rol in ["Administrador", "Gerencia"] and nuevo_rol in ROLES_SISTEMA:
            objetivo.rol = nuevo_rol
            objetivo.jefe_id = int(nuevo_jefe_str) if nuevo_jefe_str.isdigit() else None
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Cargo y asignación actualizados para {objetivo.nombre}."}
        elif user.rol == "Jefe de area" and nuevo_rol in ["Asesor", "Almacenista", "Jefe de area"]:
            objetivo.rol = nuevo_rol
            db.commit()
            request.session["flash"] = {"ok": True, "msg": f"Cargo actualizado para {objetivo.nombre}."}

    return RedirectResponse("/gestion_usuarios", status_code=303)


@app.get("/supervision", response_class=HTMLResponse)
def vista_supervision(
    request: Request,
    asesor_id: str = "",
    tipo: str = "",
    estado: str = "",
    riesgo: str = "",
    q: str = "",
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db)
):
    if user.rol not in ROLES_SUPERIORES:
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)

    visibles = ids_asesores_visibles(db, user)
    q_asesores = db.query(models.Asesor).options(
        joinedload(models.Asesor.pedidos).joinedload(models.Pedido.cliente),
        joinedload(models.Asesor.pedidos).joinedload(models.Pedido.items),
        joinedload(models.Asesor.pedidos).joinedload(models.Pedido.documentos),
        joinedload(models.Asesor.pedidos).joinedload(models.Pedido.internacional),
        joinedload(models.Asesor.cotizaciones).joinedload(models.Cotizacion.items),
    )
    if visibles is not None:
        q_asesores = q_asesores.filter(models.Asesor.id.in_(visibles))
    asesores_lista = q_asesores.order_by(models.Asesor.nombre).all()

    q_pedidos = db.query(models.Pedido).join(models.Cliente).options(
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.internacional),
        joinedload(models.Pedido.documentos),
        joinedload(models.Pedido.historial),
    )
    if visibles is not None:
        q_pedidos = q_pedidos.filter(models.Pedido.asesor_id.in_(visibles))
    if asesor_id.strip().isdigit():
        q_pedidos = q_pedidos.filter(models.Pedido.asesor_id == int(asesor_id.strip()))
    if tipo in ("Nacional", "Internacional"):
        q_pedidos = q_pedidos.filter(models.Pedido.tipo == tipo)
    if estado in ESTADOS:
        q_pedidos = q_pedidos.filter(models.Pedido.estado == estado)
    if q.strip():
        like = f"%{q.strip()}%"
        q_pedidos = q_pedidos.filter(
            models.Pedido.po.ilike(like)
            | models.Cliente.nombre.ilike(like)
            | models.Pedido.items.any(models.ItemPedido.parte_numero.ilike(like))
            | models.Pedido.items.any(models.ItemPedido.proveedor.ilike(like))
        )

    pedidos_candidatos = q_pedidos.order_by(models.Pedido.id.desc()).all()

    riesgo_limpio = (riesgo or "").strip().lower()
    if riesgo_limpio == "vencidas":
        pedidos_supervision = [
            p for p in pedidos_candidatos
            if p.estado == "En proceso" and p.dias_habiles_info["en_contra"]
        ]
    elif riesgo_limpio == "criticas":
        pedidos_supervision = [
            p for p in pedidos_candidatos
            if p.estado == "En proceso"
            and not p.dias_habiles_info["en_contra"]
            and p.dias_habiles_info["dias"] is not None
            and p.dias_habiles_info["dias"] <= 3
        ]
    elif riesgo_limpio == "en_tiempo":
        pedidos_supervision = [
            p for p in pedidos_candidatos
            if p.estado == "En proceso" and not p.dias_habiles_info["en_contra"]
        ]
    elif riesgo_limpio == "sin_compra":
        pedidos_supervision = [
            p for p in pedidos_candidatos
            if p.estado == "En proceso" and any(
                not it.fecha_compra_proveedor or it.estado_producto in ("PENDIENTE DE PAGO", "PENDIENTE")
                for it in p.items
            )
        ]
    else:
        pedidos_supervision = pedidos_candidatos

    hay_filtro_activo = bool(
        asesor_id.strip() or tipo.strip() or estado.strip() or riesgo_limpio or q.strip()
    )
    ids_pedidos_filtrados = {p.id for p in pedidos_supervision}

    resumen_asesores = []
    total_vencidas_equipo = 0
    total_proximas_equipo = 0
    total_cop_equipo = Decimal("0.00")
    total_usd_equipo = Decimal("0.00")
    total_items_sin_comprar_equipo = 0

    for a in asesores_lista:
        if asesor_id.strip().isdigit() and a.id != int(asesor_id.strip()):
            continue
        pedidos_asesor = [
            p for p in a.pedidos
            if (not hay_filtro_activo or p.id in ids_pedidos_filtrados)
        ]
        if hay_filtro_activo and not asesor_id.strip().isdigit() and len(pedidos_asesor) == 0:
            continue

        pos_activas = [p for p in pedidos_asesor if p.estado == "En proceso"]
        pos_vencidas = [p for p in pos_activas if p.dias_habiles_info["en_contra"]]
        pos_proximas = [
            p for p in pos_activas
            if not p.dias_habiles_info["en_contra"]
            and p.dias_habiles_info["dias"] is not None
            and p.dias_habiles_info["dias"] <= 3
        ]
        pos_entregadas = [p for p in pedidos_asesor if p.estado == "Entregado"]
        pos_canceladas = [p for p in pedidos_asesor if p.estado == "Cancelado"]
        vol_cop = sum((p.total for p in pedidos_asesor if p.tipo == "Nacional" and p.estado != "Cancelado"), Decimal("0.00"))
        vol_usd = sum((p.total for p in pedidos_asesor if p.tipo == "Internacional" and p.estado != "Cancelado"), Decimal("0.00"))
        items_sin_comprar = sum(
            1 for p in pos_activas for it in p.items
            if not it.fecha_compra_proveedor or it.estado_producto in ("PENDIENTE DE PAGO", "PENDIENTE")
        )
        total_vencidas_equipo += len(pos_vencidas)
        total_proximas_equipo += len(pos_proximas)
        total_cop_equipo += vol_cop
        total_usd_equipo += vol_usd
        total_items_sin_comprar_equipo += items_sin_comprar

        sla_salud = 100
        if len(pos_activas) > 0:
            sla_salud = max(0, int(round(((len(pos_activas) - len(pos_vencidas)) / len(pos_activas)) * 100)))

        resumen_asesores.append({
            "asesor": a,
            "total_pos": len(pedidos_asesor),
            "pos_activas": len(pos_activas),
            "pos_vencidas": len(pos_vencidas),
            "pos_proximas": len(pos_proximas),
            "pos_entregadas": len(pos_entregadas),
            "pos_canceladas": len(pos_canceladas),
            "cotizaciones_count": len(a.cotizaciones or []),
            "volumen_cop": vol_cop,
            "volumen_usd": vol_usd,
            "items_sin_comprar": items_sin_comprar,
            "sla_salud": sla_salud,
            "pedidos": sorted(pedidos_asesor, key=lambda x: x.dias_habiles_info["orden_urgencia"])
        })

    # Incluir todos los pedidos visibles en el mapa de vista previa para que cualquier badge funcione al instante
    todos_visibles_sup = {p.id: p for a in asesores_lista for p in a.pedidos}
    for p in pedidos_supervision:
        todos_visibles_sup[p.id] = p
    preview_map_sup = {f"PO-{p.id}": _serializar_po_preview(p) for p in todos_visibles_sup.values()}

    q_hist = db.query(models.PedidoHistorial)
    if hay_filtro_activo and ids_pedidos_filtrados:
        q_hist = q_hist.filter(models.PedidoHistorial.pedido_id.in_(ids_pedidos_filtrados))
    historial_reciente = q_hist.order_by(models.PedidoHistorial.id.desc()).limit(15).all()
    movimientos_bodega = db.query(models.MovimientoInventario).order_by(models.MovimientoInventario.id.desc()).limit(12).all()

    return render(request, "Supervision.html", {
        "user": user,
        "asesores_lista": asesores_lista,
        "resumen_asesores": resumen_asesores,
        "pedidos_supervision": pedidos_supervision,
        "preview_map_json": json.dumps(preview_map_sup, ensure_ascii=False),
        "historial_reciente": historial_reciente,
        "movimientos_bodega": movimientos_bodega,
        "total_vencidas_equipo": total_vencidas_equipo,
        "total_proximas_equipo": total_proximas_equipo,
        "total_cop_equipo": total_cop_equipo,
        "total_usd_equipo": total_usd_equipo,
        "total_items_sin_comprar_equipo": total_items_sin_comprar_equipo,
        "estados": ESTADOS,
        "estados_aduana": ESTADOS_ADUANA,
        "f_asesor": asesor_id.strip(),
        "f_tipo": tipo.strip(),
        "f_estado": estado.strip(),
        "f_riesgo": riesgo_limpio,
        "q": q.strip(),
        "hay_filtro_activo": hay_filtro_activo,
    })


@app.post("/supervision/actualizar_estado_po")
async def actualizar_estado_rapido_supervision(
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db)
):
    if user.rol not in ROLES_SUPERIORES:
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/supervision", status_code=303)

    f = await request.form()
    pid_str = (f.get("pedido_id") or "").strip()
    nuevo_estado = (f.get("nuevo_estado") or "").strip()
    nota_sup = (f.get("nota_supervision") or "").strip()

    if not pid_str.isdigit() or nuevo_estado not in ESTADOS:
        request.session["flash"] = {"ok": False, "msg": "Seleccione una orden y un estado operativo válido."}
        return RedirectResponse("/supervision", status_code=303)

    pedido = pedido_del_usuario(db, int(pid_str), user, bloquear=True)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "No se encontró la orden o no tiene permisos sobre ella."}
        return RedirectResponse("/supervision", status_code=303)

    ahora_dt = datetime.now()
    estado_prev = pedido.estado
    pedido.estado = nuevo_estado
    pedido.actualizado_en = ahora_dt
    if nuevo_estado == "Entregado" and not pedido.fecha_entrega_cliente:
        pedido.fecha_entrega_cliente = date.today()
    elif nuevo_estado == "Cancelado":
        pedido.motivo_cancelacion = nota_sup or "Cancelación ejecutada desde panel de supervisión."
        pedido.fecha_cancelacion = ahora_dt

    if nota_sup:
        sello = f"[Supervisión {user.nombre}: {nota_sup}]"
        pedido.observaciones = f"{(pedido.observaciones or '').strip()} {sello}".strip()

    registrar_historial_pedido(
        db, pedido, estado_prev, nuevo_estado,
        accion=f"SUPERVISION_{user.rol.upper().replace(' ', '_')}",
        motivo=nota_sup or f"Actualización directa desde Supervisión ({user.rol})",
        user=user,
    )
    db.commit()
    db.refresh(pedido)
    generar_salidas(db, pedido)
    request.session["flash"] = {
        "ok": True,
        "msg": f"Orden {pedido.po_codigo} actualizada a «{nuevo_estado}» por {user.nombre}."
    }
    return RedirectResponse("/supervision", status_code=303)


@app.post("/pedidos/{pid}/reactivar")
async def reactivar_pedido_cancelado(
    pid: int,
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db)
):
    if user.rol not in ROLES_GERENCIA_ADMIN:
        request.session["flash"] = {
            "ok": False,
            "msg": "La reactivación de órdenes canceladas es exclusiva para Gerencia y Administración."
        }
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)

    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    pedido = pedido_del_usuario(db, pid, user, bloquear=True)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada."}
        return RedirectResponse("/pedidos", status_code=303)

    f = await request.form()
    justificacion = (f.get("motivo_reactivacion") or "Reactivación autorizada por Gerencia/Administración").strip()
    estado_prev = pedido.estado
    pedido.estado = "En proceso"
    pedido.actualizado_en = datetime.now()
    registrar_historial_pedido(
        db, pedido, estado_prev, "En proceso",
        accion="REACTIVACION_EJECUTIVA",
        motivo=justificacion,
        user=user,
    )
    db.commit()
    db.refresh(pedido)
    generar_salidas(db, pedido)
    request.session["flash"] = {
        "ok": True,
        "msg": f"Orden {pedido.po_codigo} reactivada a estado En proceso. Autorización registrada en historial."
    }
    return RedirectResponse("/supervision", status_code=303)


@app.post("/supervision/reasignar")
async def reasignar_po_supervision(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)

    f = await request.form()
    modo_transferencia = (f.get("modo") or "individual").strip()
    nuevo_asesor_id = (f.get("nuevo_asesor_id") or "").strip()

    if not nuevo_asesor_id.isdigit():
        request.session["flash"] = {"ok": False, "msg": "Seleccione un asesor destino válido."}
        return RedirectResponse("/supervision", status_code=303)

    asesor_destino = db.get(models.Asesor, int(nuevo_asesor_id))
    if not asesor_destino:
        request.session["flash"] = {"ok": False, "msg": "El asesor destino no existe."}
        return RedirectResponse("/supervision", status_code=303)

    visibles = ids_asesores_visibles(db, user)
    if visibles is not None and asesor_destino.id not in visibles:
        request.session["flash"] = {"ok": False, "msg": "No tiene permisos sobre el asesor destino."}
        return RedirectResponse("/supervision", status_code=303)

    if modo_transferencia == "masivo":
        origen_id = (f.get("asesor_origen_id") or "").strip()
        if not origen_id.isdigit() or int(origen_id) == asesor_destino.id:
            request.session["flash"] = {"ok": False, "msg": "Seleccione un asesor origen diferente al destino."}
            return RedirectResponse("/supervision", status_code=303)
        asesor_origen = db.get(models.Asesor, int(origen_id))
        if not asesor_origen or (visibles is not None and asesor_origen.id not in visibles):
            request.session["flash"] = {"ok": False, "msg": "Sin permisos sobre el asesor origen."}
            return RedirectResponse("/supervision", status_code=303)

        pedidos_mover = db.query(models.Pedido).with_for_update().options(
            joinedload(models.Pedido.items),
            joinedload(models.Pedido.cliente),
            joinedload(models.Pedido.asesor),
            joinedload(models.Pedido.internacional),
            joinedload(models.Pedido.documentos)
        ).filter(models.Pedido.asesor_id == asesor_origen.id).all()

        cantidad_movida = len(pedidos_mover)
        for p in pedidos_mover:
            p.asesor_id = asesor_destino.id
            p.actualizado_en = datetime.utcnow()
        db.commit()

        for p in pedidos_mover:
            db.refresh(p)
            generar_salidas(db, p)

        request.session["flash"] = {
            "ok": True,
            "msg": f"Se transfirieron {cantidad_movida} órdenes de {asesor_origen.nombre} a {asesor_destino.nombre}."
        }
        return RedirectResponse("/supervision", status_code=303)

    pedido_id = (f.get("pedido_id") or "").strip()
    if not pedido_id.isdigit():
        request.session["flash"] = {"ok": False, "msg": "Seleccione una orden válida."}
        return RedirectResponse("/supervision", status_code=303)

    pedido = pedido_del_usuario(db, int(pedido_id), user, bloquear=True)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada o sin permisos."}
        return RedirectResponse("/supervision", status_code=303)

    asesor_anterior = pedido.asesor.nombre if pedido.asesor else "Anterior"
    pedido.asesor_id = asesor_destino.id
    pedido.actualizado_en = datetime.utcnow()
    db.commit()
    db.refresh(pedido)
    generar_salidas(db, pedido)

    request.session["flash"] = {
        "ok": True,
        "msg": f"Orden {pedido.po_codigo} reasignada de {asesor_anterior} a {asesor_destino.nombre}."
    }
    return RedirectResponse("/supervision", status_code=303)


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
        joinedload(models.Pedido.internacional),
        joinedload(models.Pedido.documentos)
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
    valor_total_ventas_cop = sum(float(p.total) for p in pedidos if p.tipo == "Nacional")
    valor_total_ventas_usd = sum(float(p.total) for p in pedidos if p.tipo == "Internacional")
    valor_subtotal_cop = sum(float(p.subtotal) for p in pedidos if p.tipo == "Nacional")
    valor_subtotal_usd = sum(float(p.subtotal) for p in pedidos if p.tipo == "Internacional")
    valor_hazmat_cop = sum(float(p.total_hazmat) for p in pedidos if p.tipo == "Nacional")
    valor_hazmat_usd = sum(float(p.total_hazmat) for p in pedidos if p.tipo == "Internacional")
    valor_certificados_cop = sum(float(p.total_certificados) for p in pedidos if p.tipo == "Nacional")
    valor_certificados_usd = sum(float(p.total_certificados) for p in pedidos if p.tipo == "Internacional")

    suma_probabilidad = 0
    tiempos_gestion = []
    entregas_a_tiempo = 0
    entregas_evaluadas = 0
    compras_pendientes_po = []
    todos_items_global = []
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
            "cancelados": 0,
            "items_pendientes_compra": 0,
            "volumen": 0.0,
            "volumen_cop": 0.0,
            "volumen_usd": 0.0,
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
        if p.plazo_entrega and p.estado != "Cancelado":
            entregas_evaluadas += 1
            if en_tiempo:
                entregas_a_tiempo += 1

        faltan_compras_item = 0
        for it in p.items:
            prov_nom = it.proveedor or "Sin Proveedor"
            proveedores_top[prov_nom] = proveedores_top.get(prov_nom, 0.0) + float(it.subtotal_item)
            if p.estado != "Cancelado" and (not it.fecha_compra_proveedor or it.estado_producto in ("PENDIENTE DE PAGO", "PENDIENTE", "COMPRADO")):
                faltan_compras_item += 1
                items_sin_compra += 1

            todos_items_global.append({
                "pedido_id": p.id,
                "po": p.po_codigo,
                "tipo": p.tipo,
                "moneda": p.moneda,
                "asesor": p.asesor.nombre if p.asesor else "",
                "cliente": cli_nom,
                "proveedor": prov_nom,
                "fecha_compra": it.fecha_compra_proveedor.isoformat() if it.fecha_compra_proveedor else "FALTA COMPRAR",
                "comprado": bool(it.fecha_compra_proveedor),
                "parte_numero": it.parte_numero,
                "descripcion": it.descripcion,
                "cantidad": it.cantidad,
                "unidad_medida": it.unidad_medida or "UND",
                "precio_unitario": float(it.precio_unitario or 0),
                "cobro_hazmat": float(it.cobro_hazmat_dec),
                "cobro_certificado": float(it.cobro_certificado_dec),
                "subtotal_item": float(it.subtotal_item),
                "iva_porcentaje": 0.0 if it.es_exento_iva else float(it.iva_porcentaje or 0),
                "total_item": float(it.total_item),
                "estado_producto": it.estado_producto or "PENDIENTE DE PAGO",
                "estado_po": p.estado,
                "semaforo": info_habiles["etiqueta"],
            })

        if p.estado == "En proceso" or faltan_compras_item > 0:
            compras_pendientes_po.append({
                "id": p.id,
                "po": p.po_codigo,
                "cliente": cli_nom,
                "asesor": p.asesor.nombre if p.asesor else "",
                "asesor_email": (p.asesor.email if p.asesor and p.asesor.email else CORREO_OFICIAL),
                "tipo": p.tipo,
                "moneda": p.moneda,
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
            if p.tipo == "Nacional":
                m["volumen_cop"] += float(p.total)
            else:
                m["volumen_usd"] += float(p.total)
            m["nacionalizacion"] += float(p.valor_nacionalizacion_dec)
            m["prob_acum"] += prob_p
            m["dias_acum"] += dias_gest
            m["dias_count"] += 1
            m["items_pendientes_compra"] += faltan_compras_item
            if p.estado == "Entregado":
                m["entregados"] += 1
            elif p.estado == "Cancelado":
                m["cancelados"] += 1
            else:
                m["pendientes"] += 1
            if p.plazo_entrega and p.estado != "Cancelado":
                m["sla_total"] += 1
                if en_tiempo:
                    m["sla_ok"] += 1

    probabilidad_global = round(suma_probabilidad / total_pedidos, 1) if total_pedidos else 0.0
    tiempo_promedio_dias = round(sum(tiempos_gestion) / len(tiempos_gestion), 1) if tiempos_gestion else 0.0
    cumplimiento_sla = round((entregas_a_tiempo / entregas_evaluadas) * 100, 1) if entregas_evaluadas else 100.0
    pendientes_gestion = sum(1 for p in pedidos if p.estado == "En proceso")

    tabla_asesores = []
    for aid, m in metricas_por_asesor.items():
        t_pos = m["total_pos"]
        m["tiempo_promedio"] = round(m["dias_acum"] / m["dias_count"], 1) if m["dias_count"] else 0.0
        m["probabilidad"] = round(m["prob_acum"] / t_pos, 1) if t_pos else 0.0
        m["sla_pct"] = round((m["sla_ok"] / m["sla_total"]) * 100, 1) if m["sla_total"] else 100.0
        m["sla_porcentaje"] = m["sla_pct"]
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

    # Analítica avanzada: Ticket promedio, IVA, Splits de pago, Transportadoras, Conversión y Bodegas
    ticket_promedio_cop = round(valor_total_ventas_cop / total_nacionales, 2) if total_nacionales else 0.0
    ticket_promedio_usd = round(valor_total_ventas_usd / total_internacionales, 2) if total_internacionales else 0.0
    total_impuestos_cop = sum(float(p.total_impuestos) for p in pedidos if p.tipo == "Nacional")
    total_impuestos_usd = sum(float(p.total_impuestos) for p in pedidos if p.tipo == "Internacional")
    items_exentos_count = sum(1 for p in pedidos for it in p.items if it.es_exento_iva)

    riesgo_vencidas = sum(1 for p in pedidos if p.estado == "En proceso" and p.dias_habiles_info["en_contra"])
    riesgo_criticas = sum(
        1 for p in pedidos
        if p.estado == "En proceso"
        and not p.dias_habiles_info["en_contra"]
        and p.dias_habiles_info["dias"] is not None
        and p.dias_habiles_info["dias"] <= 3
    )
    riesgo_holgadas = sum(
        1 for p in pedidos
        if p.estado == "En proceso"
        and not p.dias_habiles_info["en_contra"]
        and (p.dias_habiles_info["dias"] is None or p.dias_habiles_info["dias"] > 3)
    )

    desglose_metodos_global = {}
    distribucion_carriers = {"FedEx": 0, "Coordinadora": 0, "DHL": 0, "Clickmail": 0, "Por asignar": 0}
    for p in pedidos:
        for sp in p.desglose_pagos:
            m_nom = sp["metodo"]
            desglose_metodos_global[m_nom] = desglose_metodos_global.get(m_nom, 0.0) + float(sp["monto"] or 0)
        c_nom = p.tracking_info.get("carrier") or "Por asignar"
        distribucion_carriers[c_nom] = distribucion_carriers.get(c_nom, 0) + 1

    q_cots_an = db.query(models.Cotizacion).options(joinedload(models.Cotizacion.items))
    if visibles is not None:
        q_cots_an = q_cots_an.filter(models.Cotizacion.asesor_id.in_(visibles))
    if asesor_filtro.isdigit():
        q_cots_an = q_cots_an.filter(models.Cotizacion.asesor_id == int(asesor_filtro))
    if tipo_filtro in ("Nacional", "Internacional"):
        q_cots_an = q_cots_an.filter(models.Cotizacion.tipo_operacion == tipo_filtro)
    if desde:
        try:
            q_cots_an = q_cots_an.filter(models.Cotizacion.fecha >= datetime.strptime(desde, "%Y-%m-%d").date())
        except ValueError:
            pass
    if hasta:
        try:
            q_cots_an = q_cots_an.filter(models.Cotizacion.fecha <= datetime.strptime(hasta, "%Y-%m-%d").date())
        except ValueError:
            pass
    cotizaciones_analitica = q_cots_an.all()
    total_cotizaciones_an = len(cotizaciones_analitica)
    valor_cotizado_cop = sum(float(c.total) for c in cotizaciones_analitica if (c.moneda or "COP") == "COP")
    valor_cotizado_usd = sum(float(c.total) for c in cotizaciones_analitica if (c.moneda or "COP") == "USD")
    tasa_conversion_comercial = round(
        (total_pedidos / (total_pedidos + total_cotizaciones_an)) * 100, 1
    ) if (total_pedidos + total_cotizaciones_an) > 0 else 0.0

    items_inv_all = db.query(models.InventarioItem).all()
    resumen_inventario = {
        "total_skus": len(items_inv_all),
        "total_unidades": sum(int(i.existencia or 0) for i in items_inv_all),
        "vencidos": sum(1 for i in items_inv_all if i.convencion_info["vencido"]),
        "criticos": sum(1 for i in items_inv_all if i.convencion_info["color"] == "amber"),
        "precaucion": sum(1 for i in items_inv_all if i.convencion_info["color"] == "yellow"),
        "vigentes": sum(1 for i in items_inv_all if i.convencion_info["color"] == "emerald"),
        "w1_unidades": sum(int(i.existencia or 0) for i in items_inv_all if i.warehouse == 1),
        "w2_unidades": sum(int(i.existencia or 0) for i in items_inv_all if i.warehouse == 2),
        "w3_unidades": sum(int(i.existencia or 0) for i in items_inv_all if i.warehouse == 3),
    }

    return {
        "user": user,
        "pedidos_filtrados": pedidos,
        "todos_items_global": todos_items_global,
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
        "valor_total_cop": valor_total_ventas_cop,
        "valor_total_usd": valor_total_ventas_usd,
        "valor_total_ventas_cop": valor_total_ventas_cop,
        "valor_total_ventas_usd": valor_total_ventas_usd,
        "valor_subtotal_cop": valor_subtotal_cop,
        "valor_subtotal_usd": valor_subtotal_usd,
        "valor_hazmat_cop": valor_hazmat_cop,
        "valor_hazmat_usd": valor_hazmat_usd,
        "valor_certificados_cop": valor_certificados_cop,
        "valor_certificados_usd": valor_certificados_usd,
        "valor_total_nacionalizacion": valor_total_nacionalizacion,
        "ticket_promedio_cop": ticket_promedio_cop,
        "ticket_promedio_usd": ticket_promedio_usd,
        "total_impuestos_cop": total_impuestos_cop,
        "total_impuestos_usd": total_impuestos_usd,
        "items_exentos_count": items_exentos_count,
        "riesgo_vencidas": riesgo_vencidas,
        "riesgo_criticas": riesgo_criticas,
        "riesgo_holgadas": riesgo_holgadas,
        "desglose_metodos_global": desglose_metodos_global,
        "distribucion_carriers": distribucion_carriers,
        "total_cotizaciones_an": total_cotizaciones_an,
        "valor_cotizado_cop": valor_cotizado_cop,
        "valor_cotizado_usd": valor_cotizado_usd,
        "tasa_conversion_comercial": tasa_conversion_comercial,
        "resumen_inventario": resumen_inventario,
        "pendientes_gestion": pendientes_gestion,
        "total_en_proceso": pendientes_gestion,
        "total_pedidos": total_pedidos,
        "items_sin_compra": items_sin_compra,
        "probabilidad_global": probabilidad_global,
        "tiempo_promedio_dias": tiempo_promedio_dias,
        "promedio_dias_gestion": tiempo_promedio_dias,
        "cumplimiento_sla": cumplimiento_sla,
        "tabla_asesores": tabla_asesores,
        "metricas_asesores": tabla_asesores,
        "total_entregas_evaluadas": entregas_evaluadas,
        "entregas_a_tiempo": entregas_a_tiempo,
        "entregas_con_retraso": max(0, entregas_evaluadas - entregas_a_tiempo),
        "top_prov_sorted": top_prov_sorted,
        "todos_items_global": todos_items_global,
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
        return RedirectResponse(ruta_inicio_rol(user), status_code=303)
    ctx = _construir_analitica(db, user, asesor_id, tipo, desde, hasta)
    return render(request, "Informes.html", ctx)


@app.get("/informes/excel")
@app.get("/informes/exportar")
@app.get("/informes/exportar/gestion_sla")
def descargar_informe_gerencial(
    asesor: str = "",
    asesor_id: str = "",
    tipo: str = "",
    desde: str = "",
    hasta: str = "",
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)

    filtro_asesor = asesor or asesor_id
    if filtro_asesor.lower() == "todos":
        filtro_asesor = ""

    ctx = _construir_analitica(db, user, filtro_asesor, tipo, desde, hasta)
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Metricas_Asesores"
    ws1.merge_cells("A1:M1")
    ws1["A1"] = "INFORME GERENCIAL DE DESEMPEÑO E INDICADORES OPERATIVOS - FLY LOGISTIC S.A.S."
    ws1["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws1["A1"].fill = PatternFill("solid", fgColor="002060")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

    enc1 = [
        "Asesor", "Cargo", "Superior", "Total POs", "Entregadas", "En Proceso",
        "Ítems Sin Compra", "Tiempo Prom. (Días)", "% Cumplimiento SLA",
        "Nivel de Desempeño", "Avance Ponderado (%)", "Nacionalización ($)", "Volumen Total ($)"
    ]
    ws1.append(enc1)
    for c in ws1[2]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1E40AF")

    for a in ctx["tabla_asesores"]:
        ws1.append([
            a["nombre"], a["rol"], a["jefe"], a["total_pos"], a["entregados"],
            a["pendientes"], a["items_pendientes_compra"], a["tiempo_promedio"],
            a["sla_pct"], a["nivel"], a["probabilidad"], a["nacionalizacion"], a["volumen"]
        ])
    for i in range(1, len(enc1) + 1):
        ws1.column_dimensions[get_column_letter(i)].width = 22

    ws2 = wb.create_sheet("Seguimiento_POs")
    enc2 = [
        "PO", "Cliente", "Asesor", "Operación", "Moneda", "Estado", "Indicador Plazo (Semáforo)",
        "Fecha Prometida", "Ítems Pendientes", "Nacionalización ($)", "Avance (%)", "Total PO ($)"
    ]
    ws2.append(enc2)
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="002060")

    for idx_row, cp in enumerate(ctx["compras_pendientes_po"], start=2):
        ws2.append([
            cp["po"], cp["cliente"], cp["asesor"], cp["tipo"], cp["moneda"], cp["estado"],
            cp["dias_habiles_info"]["etiqueta"], cp["plazo"], cp["items_pendientes"],
            cp["valor_nacionalizacion"], cp["probabilidad"], cp["total"]
        ])
        color_sem = "FEE2E2" if cp["dias_habiles_info"]["en_contra"] else "DCFCE7"
        ws2[f"G{idx_row}"].fill = PatternFill("solid", fgColor=color_sem)
    for i in range(1, len(enc2) + 1):
        ws2.column_dimensions[get_column_letter(i)].width = 22

    ws3 = wb.create_sheet("Detalle_Global_Items")
    enc3 = [
        "PO", "Operación", "Moneda", "Asesor", "Cliente", "Proveedor", "Fecha Compra", "¿Comprado?",
        "Part Number", "Descripción", "Cantidad", "Unidad", "Precio Unit.", "Cobro Hazmat",
        "Cobro Certificado", "Subtotal Ítem", "IVA / TAX (%)", "Total Ítem", "Estado Ítem", "Estado PO", "Indicador Plazo"
    ]
    ws3.append(enc3)
    for c in ws3[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="002060")
    for idx_it, it_g in enumerate(ctx["todos_items_global"], start=2):
        ws3.append([
            it_g["po"], it_g["tipo"], it_g["moneda"], it_g["asesor"], it_g["cliente"], it_g["proveedor"],
            it_g["fecha_compra"], "SÍ" if it_g["comprado"] else "FALTA COMPRAR",
            it_g["parte_numero"], it_g["descripcion"], it_g["cantidad"], it_g["unidad_medida"],
            it_g["precio_unitario"], it_g["cobro_hazmat"], it_g["cobro_certificado"],
            it_g["subtotal_item"], it_g["iva_porcentaje"], it_g["total_item"],
            it_g["estado_producto"], it_g["estado_po"], it_g["semaforo"]
        ])
        color_c = "DCFCE7" if it_g["comprado"] else "FEE2E2"
        ws3[f"H{idx_it}"].fill = PatternFill("solid", fgColor=color_c)
    for i in range(1, len(enc3) + 1):
        ws3.column_dimensions[get_column_letter(i)].width = 22

    carpeta_rep = os.path.join(BASE_DIR, "Reportes_Gerenciales")
    os.makedirs(carpeta_rep, exist_ok=True)
    ruta_rep = os.path.join(carpeta_rep, "Informe_Gerencial_FlyLogistic.xlsx")
    wb.save(ruta_rep)
    return FileResponse(ruta_rep, filename="Informe_Gerencial_FlyLogistic.xlsx", media_type=XLSX_MIME)


@app.get("/informes/productos_excel")
@app.get("/informes/exportar/productos_po")
@app.get("/informes/exportar/maestro_productos")
def descargar_excel_maestro_productos(
    asesor: str = "",
    asesor_id: str = "",
    tipo: str = "",
    estado_producto: str = "",
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)

    filtro_asesor = asesor or asesor_id
    if filtro_asesor.lower() == "todos":
        filtro_asesor = ""

    visibles = ids_asesores_visibles(db, user)
    q_ped = db.query(models.Pedido).options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor),
        joinedload(models.Pedido.internacional),
        joinedload(models.Pedido.documentos)
    )
    if visibles is not None:
        q_ped = q_ped.filter(models.Pedido.asesor_id.in_(visibles))
    if filtro_asesor.isdigit():
        q_ped = q_ped.filter(models.Pedido.asesor_id == int(filtro_asesor))
    if tipo in ("Nacional", "Internacional"):
        q_ped = q_ped.filter(models.Pedido.tipo == tipo)
    if estado_producto in ESTADOS_PRODUCTO:
        q_ped = q_ped.filter(models.Pedido.items.any(models.ItemPedido.estado_producto == estado_producto))

    pedidos = q_ped.order_by(models.Pedido.id.desc()).all()
    carpeta_rep = os.path.join(BASE_DIR, "Reportes_Gerenciales")
    os.makedirs(carpeta_rep, exist_ok=True)
    ruta_rep = os.path.join(carpeta_rep, "Control_Maestro_Productos_PO_FlyLogistic.xlsx")
    generar_excel_maestro_productos(ruta_rep, pedidos)
    return FileResponse(ruta_rep, filename="Control_Maestro_Productos_PO_FlyLogistic.xlsx", media_type=XLSX_MIME)


def _serializar_po_preview(p: models.Pedido) -> dict:
    trk = p.tracking_info
    return {
        "uid": f"PO-{p.id}",
        "id": p.id,
        "kind": "PO",
        "registro_tipo": "PO",
        "codigo": p.po_codigo,
        "cliente": p.cliente.nombre if p.cliente else "",
        "asesor": p.asesor.nombre if p.asesor else "",
        "asesor_rol": p.asesor.rol if p.asesor else "",
        "tipo": p.tipo,
        "tipo_operacion": p.tipo,
        "moneda": p.moneda,
        "estado": p.estado,
        "fecha_inicial": p.fecha_inicial.isoformat() if p.fecha_inicial else "",
        "plazo_entrega": p.plazo_entrega.isoformat() if p.plazo_entrega else "Sin definir",
        "fecha_entrega": p.fecha_entrega_cliente.isoformat() if p.fecha_entrega_cliente else "",
        "indicador": p.dias_habiles_info,
        "semaforo_etiqueta": p.dias_habiles_info["etiqueta"],
        "semaforo_color": p.dias_habiles_info["color"],
        "metodo_pago": p.metodo_pago or "Por definir",
        "desglose_pagos": [
            {"metodo": x["metodo"], "monto": float(x["monto"])}
            for x in p.desglose_pagos
        ],
        "subtotal": float(p.subtotal),
        "hazmat": float(p.total_hazmat),
        "certificados": float(p.total_certificados),
        "impuestos": float(p.total_impuestos),
        "nacionalizacion": float(p.valor_nacionalizacion_dec),
        "total": float(p.total),
        "tracking": trk,
        "tracking_transportadora": trk["transportadora"],
        "tracking_codigo": trk["codigo"],
        "tracking_url": trk["url"],
        "tracking_etiqueta": trk["etiqueta"],
        "observaciones": p.observaciones or "",
        "motivo_cancelacion": getattr(p, "motivo_cancelacion", "") or "",
        "fecha_cancelacion": (
            p.fecha_cancelacion.strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(p.fecha_cancelacion, "strftime")
            else str(p.fecha_cancelacion or "")
        ),
        "url_abrir": f"/pedidos/{p.id}/editar",
        "url_editar": f"/pedidos/{p.id}/editar",
        "url_excel": f"/pedidos/{p.id}/excel",
        "documentos": [
            {
                "id": d.id,
                "nombre": d.nombre_archivo,
                "url": f"/pedidos/{p.id}/documento?nombre={d.nombre_archivo}",
                "url_preview": f"/pedidos/{p.id}/documento/preview?nombre={d.nombre_archivo}",
                "drive_link": d.drive_link or p.documentacion_drive_link or "",
                "extension": os.path.splitext(d.nombre_archivo)[1].lower().lstrip("."),
                "es_previsualizable": os.path.splitext(d.nombre_archivo)[1].lower() in (".pdf", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".txt"),
            }
            for d in p.lista_documentos
        ],
        "historial": [
            {
                "accion": getattr(h, "accion", "") or "CAMBIO_ESTADO",
                "fecha": h.creado_en.strftime("%Y-%m-%d %H:%M") if h.creado_en else "",
                "anterior": h.estado_anterior or "Inicial",
                "estado_anterior": h.estado_anterior or "Inicial",
                "nuevo": h.estado_nuevo,
                "estado_nuevo": h.estado_nuevo,
                "usuario": h.usuario_nombre or "",
                "motivo": h.motivo or "",
                "total": f"{Decimal(str(h.total_snapshot or 0)):,.2f} {h.moneda}"
            }
            for h in (p.historial or [])
        ],
        "items": [
            {
                "parte_numero": it.parte_numero,
                "descripcion": it.descripcion,
                "proveedor": it.proveedor,
                "cantidad": it.cantidad,
                "unidad": it.unidad_medida or "UND",
                "precio_unitario": float(it.precio_unitario or 0),
                "cobro_hazmat": float(it.cobro_hazmat_dec),
                "cobro_certificado": float(it.cobro_certificado_dec),
                "sin_iva": bool(it.es_exento_iva),
                "iva": "EXENTO 0%" if it.es_exento_iva else f"{Decimal(str(it.iva_porcentaje or 0)):g}%",
                "total": float(it.total_item),
                "estado": it.estado_producto or "PENDIENTE DE PAGO",
            }
            for it in p.items
        ],
    }


def _serializar_cotizacion_preview(c: models.Cotizacion) -> dict:
    info = c.dias_habiles_info
    return {
        "uid": f"COT-{c.id}",
        "id": c.id,
        "kind": "COTIZACION",
        "registro_tipo": "COT",
        "codigo": c.codigo_cotizacion,
        "cliente": c.empresa,
        "asesor": c.asesor.nombre if c.asesor else (c.prepared_by or ""),
        "asesor_rol": c.asesor.rol if c.asesor else "Asesor",
        "tipo": c.tipo_operacion or "Cotización",
        "tipo_operacion": c.tipo_operacion or "Cotización",
        "moneda": c.moneda or "COP",
        "estado": c.estado or "Activa",
        "fecha_inicial": c.fecha.isoformat() if c.fecha else "",
        "plazo_entrega": c.fecha_expiracion.isoformat() if c.fecha_expiracion else "Sin definir",
        "fecha_entrega": "",
        "indicador": info,
        "semaforo_etiqueta": info["etiqueta"],
        "semaforo_color": info["color"],
        "metodo_pago": c.terms or "Según cotización",
        "desglose_pagos": [],
        "subtotal": float(c.subtotal),
        "hazmat": float(c.total_hazmat),
        "certificados": float(c.total_certificados),
        "impuestos": float(c.total_impuestos),
        "nacionalizacion": float(c.transporte_dec),
        "total": float(c.total),
        "tracking": {"carrier": "", "guia": "", "url": ""},
        "tracking_transportadora": "No aplica (Cotización)",
        "tracking_codigo": c.ref_no or "Sin referencia",
        "tracking_url": "",
        "tracking_etiqueta": f"Ref: {c.ref_no}" if c.ref_no else "Cotización Comercial",
        "observaciones": f"Contacto: {c.contact or 'N/A'} | Tel: {c.phone or 'N/A'} | Email: {c.email or 'N/A'}",
        "motivo_cancelacion": "",
        "url_abrir": f"/cotizador/nuevo?cot_id={c.id}",
        "url_editar": f"/pedidos/nuevo?cot_id={c.id}",
        "url_excel": f"/cotizaciones/{c.id}/excel",
        "url_convertir": f"/pedidos/nuevo?cot_id={c.id}",
        "documentos": [],
        "historial": [],
        "items": [
            {
                "parte_numero": it.description,
                "descripcion": it.remarks or "Ítem de cotización",
                "proveedor": c.site or "Fly Logistic",
                "cantidad": float(it.qty or 1),
                "unidad": it.unit or "UND",
                "precio_unitario": float(it.unit_price or 0),
                "cobro_hazmat": float(it.cobro_hazmat_dec),
                "cobro_certificado": float(it.cobro_certificado_dec),
                "sin_iva": bool(it.es_exento_iva),
                "iva": "EXENTO 0%" if it.es_exento_iva else f"{Decimal(str(it.iva_porcentaje or 19)):g}%",
                "total": float(it.total_item),
                "estado": "COTIZADO",
            }
            for it in c.items
        ],
    }


@app.get("/pedidos", response_class=HTMLResponse)
def mis_pedidos(request: Request, q: str = "", po: str = "", cliente: str = "", proveedor: str = "",
                asesor_id: str = "", estado: str = "", tipo: str = "", estado_producto: str = "",
                vista: str = "", vista_registro: str = "todos", kpi: str = "",
                page: int = 1, cot_page: int = 1,
                user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    visibles = ids_asesores_visibles(db, user)
    query = (db.query(models.Pedido)
             .options(
                 joinedload(models.Pedido.items),
                 joinedload(models.Pedido.cliente),
                 joinedload(models.Pedido.asesor),
                 joinedload(models.Pedido.internacional),
                 joinedload(models.Pedido.documentos),
                 joinedload(models.Pedido.historial)
             )
             .join(models.Cliente))

    q_cot = db.query(models.Cotizacion).options(
        joinedload(models.Cotizacion.items),
        joinedload(models.Cotizacion.asesor)
    )

    if visibles is not None:
        query = query.filter(models.Pedido.asesor_id.in_(visibles))
        q_cot = q_cot.filter(models.Cotizacion.asesor_id.in_(visibles))

    asesor_f_limpio = (asesor_id or "").strip()
    if asesor_f_limpio == "":
        query = query.filter(models.Pedido.asesor_id == user.id)
        q_cot = q_cot.filter(models.Cotizacion.asesor_id == user.id)
        asesor_f_activo = str(user.id)
    elif asesor_f_limpio.lower() in ("todos", "all", "0"):
        if user.rol in ROLES_SUPERIORES:
            asesor_f_activo = "todos"
        else:
            query = query.filter(models.Pedido.asesor_id == user.id)
            q_cot = q_cot.filter(models.Cotizacion.asesor_id == user.id)
            asesor_f_activo = str(user.id)
    elif asesor_f_limpio.isdigit():
        if user.rol in ROLES_SUPERIORES:
            query = query.filter(models.Pedido.asesor_id == int(asesor_f_limpio))
            q_cot = q_cot.filter(models.Cotizacion.asesor_id == int(asesor_f_limpio))
            asesor_f_activo = asesor_f_limpio
        else:
            query = query.filter(models.Pedido.asesor_id == user.id)
            q_cot = q_cot.filter(models.Cotizacion.asesor_id == user.id)
            asesor_f_activo = str(user.id)
    else:
        query = query.filter(models.Pedido.asesor_id == user.id)
        q_cot = q_cot.filter(models.Cotizacion.asesor_id == user.id)
        asesor_f_activo = str(user.id)

    estado_limpio = (estado or "").strip()
    kpi_limpio = (kpi or "").strip().lower()
    busqueda_especifica = bool(po.strip() or q.strip() or cliente.strip() or proveedor.strip() or estado_producto.strip())

    if estado_limpio in ESTADOS:
        query = query.filter(models.Pedido.estado == estado_limpio)
    elif estado_limpio.lower() in ("todos", "all", "historial"):
        pass  # Ver todo el histórico sin filtro de estado
    elif kpi_limpio in ("entregados", "entregado"):
        query = query.filter(models.Pedido.estado == "Entregado")
    elif kpi_limpio in ("cancelados", "cancelado"):
        query = query.filter(models.Pedido.estado == "Cancelado")
    elif not busqueda_especifica:
        # Por defecto, en la vista activa del Centro de Operaciones no se muestran las órdenes en estado final (Entregadas / Canceladas)
        query = query.filter(models.Pedido.estado == "En proceso")

    if tipo in ("Nacional", "Internacional"):
        query = query.filter(models.Pedido.tipo == tipo)
        q_cot = q_cot.filter(models.Cotizacion.tipo_operacion == tipo)

    # En cotizaciones: en la vista activa por defecto no se muestran las que ya alcanzaron estado final (convertidas a PO o canceladas)
    if not busqueda_especifica and estado_limpio.lower() not in ("todos", "all", "historial"):
        q_cot = q_cot.filter(
            (models.Cotizacion.po_convertida == "") | (models.Cotizacion.po_convertida.is_(None)),
            models.Cotizacion.estado != "Cancelada"
        )

    if po.strip():
        query = query.filter(models.Pedido.po.ilike(f"%{po.strip()}%"))
        q_cot = q_cot.filter(models.Cotizacion.numero.ilike(f"%{po.strip()}%"))
    if cliente.strip():
        query = query.filter(models.Cliente.nombre.ilike(f"%{cliente.strip()}%"))
        q_cot = q_cot.filter(models.Cotizacion.empresa.ilike(f"%{cliente.strip()}%"))
    if proveedor.strip():
        query = query.filter(models.Pedido.items.any(models.ItemPedido.proveedor.ilike(f"%{proveedor.strip()}%")))
    if estado_producto in ESTADOS_PRODUCTO:
        query = query.filter(models.Pedido.items.any(models.ItemPedido.estado_producto == estado_producto))
    if q.strip():
        like = f"%{q.strip()}%"
        query = query.filter(
            models.Pedido.po.ilike(like)
            | models.Cliente.nombre.ilike(like)
            | models.Pedido.items.any(models.ItemPedido.proveedor.ilike(like))
            | models.Pedido.items.any(models.ItemPedido.parte_numero.ilike(like))
        )
        q_cot = q_cot.filter(
            models.Cotizacion.numero.ilike(like)
            | models.Cotizacion.empresa.ilike(like)
            | models.Cotizacion.ref_no.ilike(like)
            | models.Cotizacion.items.any(models.CotizacionItem.description.ilike(like))
        )

    pedidos_sin_ordenar = query.all()
    pedidos_base = sorted(
        pedidos_sin_ordenar,
        key=lambda x: (
            1 if x.estado in ("Entregado", "Cancelado") else 0,
            x.dias_habiles_info["orden_urgencia"],
            -x.id
        )
    )
    cotizaciones_base = q_cot.order_by(models.Cotizacion.id.desc()).all()

    kpi_limpio = (kpi or "").strip().lower()
    vista_raw = (vista or vista_registro or "todos").strip().lower()
    if vista_raw in ("po", "pos"):
        vista_limpia = "pos"
        vista_f = "po"
    elif vista_raw in ("cotizacion", "cotizaciones"):
        vista_limpia = "cotizaciones"
        vista_f = "cotizacion"
    else:
        vista_limpia = "todos"
        vista_f = ""

    if kpi_limpio == "cotizaciones":
        vista_limpia = "cotizaciones"
        vista_f = "cotizacion"

    pedidos_ordenados = list(pedidos_base)
    if kpi_limpio == "vencidas":
        pedidos_ordenados = [p for p in pedidos_ordenados if p.estado == "En proceso" and p.dias_habiles_info["en_contra"]]
    elif kpi_limpio == "en_plazo":
        pedidos_ordenados = [p for p in pedidos_ordenados if p.estado == "En proceso" and not p.dias_habiles_info["en_contra"]]
    elif kpi_limpio == "entregadas":
        pedidos_ordenados = [p for p in pedidos_ordenados if p.estado == "Entregado"]
    elif kpi_limpio == "canceladas":
        pedidos_ordenados = [p for p in pedidos_ordenados if p.estado == "Cancelado"]

    registros_unificados_todos = []
    if vista_limpia in ("todos", "pos"):
        for p in pedidos_ordenados:
            registros_unificados_todos.append({
                "kind": "PO",
                "es_cotizacion": False,
                "registro_tipo": "PO",
                "uid": f"PO-{p.id}",
                "id": p.id,
                "codigo": p.po_codigo,
                "cliente": p.cliente.nombre if p.cliente else "",
                "tipo": p.tipo,
                "moneda": p.moneda,
                "estado": p.estado,
                "asesor_nombre": p.asesor.nombre if p.asesor else "",
                "asesor_rol": p.asesor.rol if p.asesor else "",
                "fecha_inicial": str(p.fecha_inicial or ""),
                "plazo_entrega": str(p.plazo_entrega or ""),
                "fecha_entrega_cliente": str(p.fecha_entrega_cliente or ""),
                "total": p.total,
                "indicador": p.dias_habiles_info,
                "preview": _serializar_po_preview(p),
                "obj": p,
                "pedido": p,
                "cotizacion": None,
            })
    if vista_limpia in ("todos", "cotizaciones") and kpi_limpio not in ("vencidas", "en_plazo", "entregadas", "canceladas"):
        if not proveedor.strip() and not estado_producto.strip() and not estado.strip():
            for c in cotizaciones_base:
                registros_unificados_todos.append({
                    "kind": "COTIZACION",
                    "es_cotizacion": True,
                    "registro_tipo": "COT",
                    "uid": f"COT-{c.id}",
                    "id": c.id,
                    "codigo": c.codigo_cotizacion,
                    "cliente": c.empresa or "",
                    "tipo": "Cotización",
                    "moneda": c.moneda or "COP",
                    "estado": c.estado or "Activa",
                    "asesor_nombre": c.asesor.nombre if c.asesor else (c.prepared_by or ""),
                    "asesor_rol": c.asesor.rol if c.asesor else "Comercial",
                    "fecha_inicial": str(c.fecha or ""),
                    "plazo_entrega": str(c.fecha_expiracion or ""),
                    "fecha_entrega_cliente": "",
                    "total": c.total,
                    "indicador": c.dias_habiles_info,
                    "preview": _serializar_cotizacion_preview(c),
                    "obj": c,
                    "pedido": None,
                    "cotizacion": c,
                })

    por_pagina = REGISTROS_POR_PAGINA
    total_pedidos_filtrados = len(pedidos_ordenados)
    total_pedidos = total_pedidos_filtrados
    total_cotizaciones = len(cotizaciones_base)
    total_unificados = len(registros_unificados_todos)

    total_paginas_po = max(1, (total_pedidos_filtrados + por_pagina - 1) // por_pagina)
    pagina_actual_po = max(1, min(page, total_paginas_po))
    inicio_po = (pagina_actual_po - 1) * por_pagina
    fin_po = inicio_po + por_pagina
    pedidos_paginados = pedidos_ordenados[inicio_po:fin_po]

    total_paginas_cot = max(1, (total_cotizaciones + por_pagina - 1) // por_pagina)
    pagina_actual_cot = max(1, min(cot_page, total_paginas_cot))
    inicio_cot = (pagina_actual_cot - 1) * por_pagina
    fin_cot = inicio_cot + por_pagina
    cotizaciones_paginadas = cotizaciones_base[inicio_cot:fin_cot]

    total_paginas = total_paginas_po if vista_limpia != "cotizaciones" else total_paginas_cot
    pagina_actual = pagina_actual_po if vista_limpia != "cotizaciones" else pagina_actual_cot

    inicio = (pagina_actual - 1) * por_pagina
    fin = inicio + por_pagina
    registros_paginados = registros_unificados_todos[inicio:fin]

    preview_map = {}
    for p in pedidos_base:
        preview_map[f"PO-{p.id}"] = _serializar_po_preview(p)
    for c in cotizaciones_base:
        preview_map[f"COT-{c.id}"] = _serializar_cotizacion_preview(c)

    total_cop_cartera = sum((p.total for p in pedidos_ordenados if p.tipo == "Nacional" and p.estado != "Cancelado"), Decimal("0"))
    total_usd_cartera = sum((p.total for p in pedidos_ordenados if p.tipo == "Internacional" and p.estado != "Cancelado"), Decimal("0"))
    total_cot_cop = sum((c.total for c in cotizaciones_base if (c.moneda or "COP") == "COP"), Decimal("0"))
    total_cot_usd = sum((c.total for c in cotizaciones_base if (c.moneda or "COP") == "USD"), Decimal("0"))

    q_as = db.query(models.Asesor)
    if visibles is not None:
        q_as = q_as.filter(models.Asesor.id.in_(visibles))
    asesores_disponibles = q_as.order_by(models.Asesor.nombre).all()

    clientes_disponibles = [c[0] for c in db.query(models.Cliente.nombre).order_by(models.Cliente.nombre).all() if c[0]]
    proveedores_disponibles = sorted(list({
        r[0] for r in db.query(models.ItemPedido.proveedor).all() if r[0]
    }))

    total_vencidas = sum(1 for p in pedidos_base if p.estado == "En proceso" and p.dias_habiles_info["en_contra"])
    total_proximas = sum(
        1 for p in pedidos_base
        if p.estado == "En proceso"
        and not p.dias_habiles_info["en_contra"]
        and p.dias_habiles_info["dias"] is not None
        and p.dias_habiles_info["dias"] <= 3
    )
    total_en_proceso = sum(1 for p in pedidos_base if p.estado == "En proceso")
    total_entregados = sum(1 for p in pedidos_base if p.estado == "Entregado")
    total_cancelados = sum(1 for p in pedidos_base if p.estado == "Cancelado")

    nombre_asesor_sel = user.nombre
    if asesor_f_activo.isdigit():
        as_obj = db.get(models.Asesor, int(asesor_f_activo))
        if as_obj:
            nombre_asesor_sel = as_obj.nombre

    semaforo_resumen = [
        {
            "uid": f"PO-{p.id}",
            "etiqueta": p.dias_habiles_info["etiqueta"],
            "color": p.dias_habiles_info["color"],
        }
        for p in pedidos_ordenados[:6]
    ]

    return render(request, "pedidos.html", {
        "user": user,
        "pedidos": pedidos_paginados,
        "todos_pedidos": pedidos_base,
        "cotizaciones": cotizaciones_paginadas,
        "cotizaciones_cartera": cotizaciones_base,
        "cotizaciones_paginadas": cotizaciones_paginadas,
        "registros_unificados": registros_paginados,
        "todos_registros_unificados": registros_unificados_todos,
        "preview_map_json": json.dumps(preview_map, ensure_ascii=False),
        "pagina_actual": pagina_actual,
        "total_paginas": total_paginas,
        "pagina_actual_po": pagina_actual_po,
        "total_paginas_po": total_paginas_po,
        "pagina_actual_cot": pagina_actual_cot,
        "total_paginas_cot": total_paginas_cot,
        "cot_pagina_actual": pagina_actual_cot,
        "cot_total_paginas": total_paginas_cot,
        "total_pedidos": total_pedidos,
        "total_pedidos_filtrados": total_pedidos_filtrados,
        "total_cotizaciones": total_cotizaciones,
        "total_cotizaciones_filtradas": total_cotizaciones,
        "total_unificados": total_unificados,
        "total_registros_unificados": total_unificados,
        "por_pagina": por_pagina,
        "registros_por_pagina": por_pagina,
        "total_cop_cartera": total_cop_cartera,
        "total_usd_cartera": total_usd_cartera,
        "total_cot_cop": total_cot_cop,
        "total_cot_usd": total_cot_usd,
        "volumen_cop": total_cop_cartera,
        "volumen_usd": total_usd_cartera,
        "volumen_cot_cop": total_cot_cop,
        "volumen_cot_usd": total_cot_usd,
        "total_vencidas": total_vencidas,
        "total_proximas": total_proximas,
        "total_en_proceso": total_en_proceso,
        "total_entregados": total_entregados,
        "total_cancelados": total_cancelados,
        "semaforo_resumen": semaforo_resumen,
        "q": q,
        "po_f": po,
        "f_po": po,
        "cliente_f": cliente,
        "f_cliente": cliente,
        "proveedor_f": proveedor,
        "asesor_f": asesor_f_activo,
        "asesor_f_activo": asesor_f_activo,
        "nombre_asesor_seleccionado": nombre_asesor_sel,
        "es_rol_superior": user.rol in ROLES_SUPERIORES,
        "estado_f": estado,
        "f_estado": estado,
        "tipo_f": tipo,
        "f_tipo": tipo,
        "estado_prod_f": estado_producto,
        "vista_f": vista_f,
        "vista_registro_f": vista_limpia,
        "kpi_f": kpi_limpio,
        "f_kpi": kpi_limpio,
        "estados": ESTADOS,
        "estados_producto": ESTADOS_PRODUCTO,
        "asesores_disponibles": asesores_disponibles,
        "asesores_visibles": asesores_disponibles,
        "clientes_disponibles": clientes_disponibles,
        "lista_clientes": clientes_disponibles,
        "proveedores_disponibles": proveedores_disponibles,
        "lista_proveedores": proveedores_disponibles,
        "correo_oficial": CORREO_OFICIAL,
    })


@app.post("/pedidos/{pid}/notificar")
def notificar_po_responsable(pid: int, request: Request,
                             user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {
            "ok": False,
            "msg": "Acción reservada para Administración, Gerencia y Jefatura de Área."
        }
        return RedirectResponse("/pedidos", status_code=303)

    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada."}
        return RedirectResponse("/pedidos", status_code=303)
    destino = (pedido.asesor.email if (pedido.asesor and pedido.asesor.email) else None) or user.email or CORREO_OFICIAL

    asunto, cuerpo = construir_cuerpo_alerta_po(pedido)
    ok, detalle = enviar_correo_con_diagnostico(destino, asunto, cuerpo)
    if ok:
        request.session["flash"] = {
            "ok": True,
            "msg": f"Notificación enviada a {destino} ({pedido.dias_habiles_info['etiqueta']})."
        }
    else:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Aviso SMTP: verifique las credenciales del servidor de correo en Mi Perfil ({detalle[:90]})."
        }
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/pedidos/notificar_proximos")
def notificar_pos_proximas(request: Request, tareas: BackgroundTasks,
                           user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {
            "ok": False,
            "msg": "Acción reservada para Administración, Gerencia y Jefatura de Área."
        }
        return RedirectResponse("/pedidos", status_code=303)

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
        "msg": f"Se programó el envío de {enviados} notificaciones de seguimiento operativo."
    }
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/pedidos/notificar_faltantes")
def notificar_consolidado_faltantes(request: Request,
                                    user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_GERENCIA_ADMIN:
        request.session["flash"] = {
            "ok": False,
            "msg": "El envío del Reporte Consolidado de Faltantes está restringido a Gerencia y Administradores."
        }
        return RedirectResponse("/pedidos", status_code=303)

    pedidos = db.query(models.Pedido).options(
        joinedload(models.Pedido.items),
        joinedload(models.Pedido.cliente),
        joinedload(models.Pedido.asesor)
    ).filter(models.Pedido.estado != "Entregado").all()

    destino = user.email or CORREO_OFICIAL
    asunto, cuerpo = construir_cuerpo_consolidado_faltantes(pedidos)
    ok, detalle = enviar_correo_con_diagnostico(destino, asunto, cuerpo)
    if ok:
        request.session["flash"] = {
            "ok": True,
            "msg": f"Reporte Consolidado de Faltantes enviado a {destino}."
        }
    else:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Aviso SMTP al enviar Consolidado de Faltantes: {detalle[:90]}"
        }
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/pedidos/selector_correos")
async def selector_correos_post(request: Request, tareas: BackgroundTasks,
                                user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    f = await request.form()
    tipo_correo = (f.get("tipo_correo") or "").strip()
    pedido_id_str = (f.get("pedido_id") or "").strip()

    if tipo_correo == "consolidado_faltantes":
        return notificar_consolidado_faltantes(request, user, db)
    elif tipo_correo == "individual_po" and pedido_id_str.isdigit():
        return notificar_po_responsable(int(pedido_id_str), request, user, db)
    else:
        return notificar_pos_proximas(request, tareas, user, db)


@app.post("/pedidos/notificar_alertas")
def notificar_alertas_post(request: Request, tareas: BackgroundTasks,
                           user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {
            "ok": False,
            "msg": "Acción restringida: El envío de notificaciones por correo está reservado exclusivamente para Gerencia, Jefatura y Administradores."
        }
        return RedirectResponse("/pedidos", status_code=303)
    return notificar_pos_proximas(request, tareas, user, db)


@app.post("/pedidos/{pid}/enviar_correo")
async def pedidos_enviar_correo_individual(
    pid: int,
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db)
):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {
            "ok": False,
            "msg": "Acción restringida: El envío de notificaciones oficiales por correo está reservado exclusivamente a Gerencia, Jefatura de Área y Administradores."
        }
        return RedirectResponse("/pedidos", status_code=303)

    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada."}
        return RedirectResponse("/pedidos", status_code=303)

    f = await request.form()
    tipo_notif = (f.get("tipo_notificacion") or "seguimiento_po").strip()
    destino = (f.get("destino") or "").strip() or (pedido.asesor.email if pedido.asesor else "") or CORREO_OFICIAL
    asunto_form = (f.get("asunto") or "").strip()
    mensaje_extra = (f.get("mensaje_adicional") or "").strip()
    adjuntar_excel = str(f.get("adjuntar_po", "1")).strip() in ("1", "true", "on", "si")

    asunto_def, cuerpo_txt, cuerpo_html = construir_notificacion_po_opciones(pedido, tipo_notif, mensaje_extra)
    asunto = asunto_form or asunto_def

    archivos_adj = []
    if adjuntar_excel:
        try:
            ruta_excel = ruta_excel_po_local(pedido)
            generar_excel_com_fm_08(pedido, ruta_excel)
            if os.path.exists(ruta_excel):
                with open(ruta_excel, "rb") as ef:
                    archivos_adj.append((
                        f"Orden_Compra_{pedido.po_codigo}.xlsx",
                        ef.read(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ))
        except Exception:
            pass

    ok, detalle = enviar_correo_con_diagnostico(
        destino=destino,
        asunto=asunto,
        cuerpo=cuerpo_txt,
        archivos_adjuntos=archivos_adj,
        html_cuerpo=cuerpo_html
    )

    if ok:
        request.session["flash"] = {
            "ok": True,
            "msg": f"Notificación enviada exitosamente a {destino} ({asunto[:45]}...)."
        }
    else:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Aviso SMTP al enviar correo: {detalle[:95]}"
        }
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/cotizaciones/exportar/excel")
def exportar_cotizaciones_excel(user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    q = db.query(models.Cotizacion).options(joinedload(models.Cotizacion.asesor), joinedload(models.Cotizacion.items))
    if user.rol not in ROLES_SUPERIORES:
        q = q.filter(models.Cotizacion.asesor_id == user.id)
    cots = q.order_by(models.Cotizacion.id.desc()).all()

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "Cotizaciones Fly"

    ws.append(["NÚMERO", "FECHA EMISIÓN", "FECHA EXPIRA", "CLIENTE", "CONTACTO", "CORREO", "ASESOR", "MONEDA", "OPERACIÓN", "ESTADO", "TOTAL COTIZADO"])
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    for col in range(1, 12):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
    for idx, c in enumerate(cots, start=2):
        ws.append([
            c.quote_number,
            c.quote_date.isoformat() if c.quote_date else "",
            c.expiration_date.isoformat() if c.expiration_date else "",
            c.to_company,
            c.contact or "",
            c.email or "",
            c.asesor.nombre if c.asesor else c.prepared_by,
            c.moneda,
            c.tipo_operacion,
            c.estado,
            float(c.total or 0)
        ])
        for col in range(1, 12):
            cell = ws.cell(row=idx, column=col)
            cell.border = thin_border
            if col == 11:
                cell.number_format = "$#,##0.00"

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    os.makedirs(os.path.join(BASE_DIR, "Reportes_Gerenciales"), exist_ok=True)
    ruta = os.path.join(BASE_DIR, "Reportes_Gerenciales", "Cartera_Cotizaciones_FlyLogistic.xlsx")
    wb.save(ruta)
    return FileResponse(ruta, filename="Cartera_Cotizaciones_FlyLogistic.xlsx", media_type=XLSX_MIME)


@app.get("/api/validar_po")
def api_validar_po(po: str = "", excluir_id: int = 0, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    po_limpio = po.strip()
    if not po_limpio:
        return JSONResponse({"existe": False, "sugerido": sugerir_siguiente_po(db)})
    q = db.query(models.Pedido).filter(func.lower(models.Pedido.po) == po_limpio.lower())
    if excluir_id:
        q = q.filter(models.Pedido.id != excluir_id)
    existente = q.first()
    if existente:
        return JSONResponse({
            "existe": True,
            "po": existente.po,
            "tipo": existente.tipo,
            "asesor": existente.asesor.nombre if existente.asesor else "",
            "sugerido": sugerir_siguiente_po(db)
        })
    return JSONResponse({"existe": False, "sugerido": sugerir_siguiente_po(db)})


@app.get("/api/tracking")
def api_tracking(guia: str = "", transportadora: str = "", clickmail: str = "0", user: models.Asesor = Depends(usuario_actual)):
    es_cm = str(clickmail).strip().lower() in ("1", "true", "si", "sí")
    info = models.detectar_transportadora_tracking(
        tracking=guia,
        llega_click_mail="Si" if es_cm else "No",
        transportadora_manual=transportadora,
    )
    return JSONResponse(info)


def _form(request, db: Session, user: models.Asesor, d, modo, pedido=None, error=None, status=200):
    visibles = ids_asesores_visibles(db, user)
    q_as = db.query(models.Asesor)
    if visibles is not None:
        q_as = q_as.filter(models.Asesor.id.in_(visibles))
    asesores_disponibles = q_as.order_by(models.Asesor.nombre).all()

    pos_existentes = [
        {"po": p.po, "tipo": p.tipo}
        for p in db.query(models.Pedido).order_by(models.Pedido.id.desc()).limit(30).all()
    ]

    return render(request, "pedido_form.html", {
        "user": user,
        "d": d,
        "modo": modo,
        "pedido": pedido,
        "error": error,
        "estados": ESTADOS,
        "estados_aduana": ESTADOS_ADUANA,
        "estados_producto": ESTADOS_PRODUCTO,
        "metodos_pago_opciones": METODOS_PAGO_OPCIONES,
        "transportadoras_opciones": TRANSPORTADORAS_OPCIONES,
        "unidades_medida": UNIDADES_MEDIDA,
        "asesores_disponibles": asesores_disponibles,
        "siguiente_po_sugerida": sugerir_siguiente_po(db),
        "pos_existentes": pos_existentes,
        "max_archivos_po": MAX_ARCHIVOS_PO,
    }, status)


@app.get("/pedidos/nuevo", response_class=HTMLResponse)
def nuevo_form(request: Request, cot_id: int = 0, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    hoy = datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d")
    sugerida = sugerir_siguiente_po(db)
    d = {k: "" for k in CAMPOS}
    d.update(
        fecha_inicial=hoy,
        po_numero=sugerida,
        tipo_pedido="Nacional",
        moneda="COP",
        iva_porcentaje="19",
        valor_nacionalizacion="0.00",
        llega_click_mail="No",
        transportadora="",
        numero_guia_nacional="",
        motivo_cancelacion="",
        estado_general="En proceso",
        asesor_asignado_id=str(user.id),
        documentos_adjuntos=[],
        desglose_pagos_dict={},
        items=[{
            "fecha_compra_prov": "", "proveedor": "", "parte_numero": "",
            "parte_numero_alterno": "", "unidad_medida": "UND", "descripcion": "",
            "cantidad": "1", "precio_unitario": "", "cobro_hazmat": "0.00",
            "cobro_certificado": "0.00", "sin_iva": "0", "iva_item": "19",
            "estado_producto": "PENDIENTE DE PAGO"
        }]
    )

    if cot_id:
        cot = db.query(models.Cotizacion).filter(models.Cotizacion.id == cot_id).first()
        if cot:
            d["cliente_final"] = cot.to_company or ""
            d["moneda"] = getattr(cot, "moneda", "COP") or "COP"
            d["plazo_entrega"] = str(cot.expiration_date or "")
            if cot.items:
                d["items"] = [{
                    "fecha_compra_prov": "",
                    "proveedor": it.remarks or "",
                    "parte_numero": it.description or "",
                    "parte_numero_alterno": "",
                    "unidad_medida": it.unit or "UND",
                    "descripcion": it.description or "",
                    "cantidad": str(it.qty or 1),
                    "precio_unitario": f"{float(it.unit_price or 0):.2f}",
                    "cobro_hazmat": f"{float(getattr(it, 'cobro_hazmat_dec', 0)):.2f}",
                    "cobro_certificado": f"{float(getattr(it, 'cobro_certificado_dec', 0)):.2f}",
                    "sin_iva": "1" if getattr(it, "sin_iva", False) else "0",
                    "iva_item": "0" if getattr(it, "sin_iva", False) else "19",
                    "estado_producto": "PENDIENTE DE PAGO",
                } for it in cot.items]

    return _form(request, db, user, d, "nuevo")


@app.post("/pedidos/nuevo", response_class=HTMLResponse)
async def nuevo_post(request: Request, tareas: BackgroundTasks,
                     user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    form_raw = await request.form()
    d = datos_desde_form(form_raw)
    archivos_subidos = (
        list(form_raw.getlist("archivos_documentacion"))
        + list(form_raw.getlist("archivo_documentacion"))
    )
    
    pedido = None
    rutas_docs = []
    
    try:
        v = parsear(d)
        
        exito = False
        for intento in range(6):
            try:
                if intento > 0:
                    v["po"] = sugerir_siguiente_po(db)
                
                existente_po = db.query(models.Pedido).filter(
                    func.lower(models.Pedido.po) == v["po"].lower()
                ).with_for_update().first()
                if existente_po and intento == 0:
                    raise ValueError(
                        f"La orden «{v['po']}» ya se encuentra registrada ({existente_po.tipo}). "
                        f"Consecutivo disponible: {sugerir_siguiente_po(db)}"
                    )

                asesor_id_final = user.id
                if user.rol in ROLES_SUPERIORES and str(d.get("asesor_asignado_id") or "").isdigit():
                    asesor_id_final = int(d["asesor_asignado_id"])

                pedido = models.Pedido(po=v["po"], tipo=v["tipo"], asesor_id=asesor_id_final)
                aplicar(db, pedido, v, user=user)
                db.add(pedido)
                db.flush()

                registrar_historial_pedido(
                    db, pedido, "", pedido.estado,
                    accion="CREACION_PO",
                    motivo=pedido.motivo_cancelacion or "Creación inicial de Orden de Compra",
                    user=user
                )

                rutas_docs = await guardar_archivos_documentacion_multiple(db, pedido, archivos_subidos)
                db.commit()
                exito = True
                break
            except (IntegrityError, OperationalError) as err:
                db.rollback()
                if intento == 5:
                    raise ValueError(f"Alta concurrencia operativa (20 usuarios simultáneos): {err}. Intente guardar de nuevo.")
                time_lib.sleep(0.06 * (intento + 1))
                continue

        if not exito:
            raise ValueError("No fue posible asignar un consecutivo único debido a la alta concurrencia. Intente de nuevo.")

    except ValueError as e:
        db.rollback()
        return _form(request, db, user, d, "nuevo", error=str(e), status=400)
    except Exception as e:
        db.rollback()
        return _form(request, db, user, d, "nuevo", error=f"Error al procesar la orden: {e}", status=400)

    db.refresh(pedido)
    ruta_doc_principal = rutas_docs[0] if rutas_docs else None
    estado = generar_salidas(db, pedido, ruta_doc_principal, rutas_docs)

    if pedido.plazo_entrega:
        destino = (pedido.asesor.email if (pedido.asesor and pedido.asesor.email) else None) or CORREO_OFICIAL
        asunto, cuerpo = construir_cuerpo_alerta_po(pedido)
        tareas.add_task(enviar_correo, destino, asunto, cuerpo)

    request.session["flash"] = {"ok": True, "msg": f"Orden {pedido.po_codigo} registrada con éxito. {estado}."}
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/pedidos/{pid}/editar", response_class=HTMLResponse)
def editar_form(pid: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada o sin permisos de acceso."}
        return RedirectResponse("/pedidos", status_code=303)
    return _form(request, db, user, datos_desde_pedido(pedido), "editar", pedido)


@app.post("/pedidos/{pid}/editar", response_class=HTMLResponse)
async def editar_post(pid: int, request: Request,
                      user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    pedido = pedido_del_usuario(db, pid, user, bloquear=True)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada."}
        return RedirectResponse("/pedidos", status_code=303)

    carpeta_anterior = carpeta_po_local(pedido)
    form_raw = await request.form()
    d = datos_desde_form(form_raw)
    archivos_subidos = (
        list(form_raw.getlist("archivos_documentacion"))
        + list(form_raw.getlist("archivo_documentacion"))
    )

    d["po_numero"] = pedido.po
    if not d.get("tipo_pedido") or d.get("tipo_pedido") not in ("Nacional", "Internacional"):
        d["tipo_pedido"] = pedido.tipo

    try:
        v = parsear(d, tipo_fijo=None)
        aplicar(db, pedido, v, user=user)
        if user.rol in ROLES_SUPERIORES and d.get("asesor_asignado_id", "").isdigit():
            pedido.asesor_id = int(d["asesor_asignado_id"])

        db.flush()
        carpeta_nueva = carpeta_po_local(pedido)
        if carpeta_anterior != carpeta_nueva and os.path.exists(carpeta_anterior):
            for doc_existente in pedido.documentos:
                origen_f = os.path.join(carpeta_anterior, doc_existente.nombre_archivo)
                destino_f = os.path.join(carpeta_nueva, doc_existente.nombre_archivo)
                if os.path.exists(origen_f) and not os.path.exists(destino_f):
                    shutil.copy2(origen_f, destino_f)

        rutas_docs = await guardar_archivos_documentacion_multiple(db, pedido, archivos_subidos)
        for intento_commit in range(5):
            try:
                db.commit()
                break
            except OperationalError:
                db.rollback()
                if intento_commit == 4:
                    raise
                time_lib.sleep(0.06 * (intento_commit + 1))
    except ValueError as e:
        db.rollback()
        d["documentos_adjuntos"] = [
            {"id": doc.id, "nombre_archivo": doc.nombre_archivo, "drive_link": doc.drive_link or ""}
            for doc in pedido.lista_documentos
        ]
        return _form(request, db, user, d, "editar", pedido, error=str(e), status=400)
    except IntegrityError:
        db.rollback()
        return _form(request, db, user, d, "editar", pedido, error="Error de integridad al actualizar la orden.", status=400)

    db.refresh(pedido)
    ruta_doc_principal = rutas_docs[0] if rutas_docs else None
    estado = generar_salidas(db, pedido, ruta_doc_principal, rutas_docs)
    request.session["flash"] = {"ok": True, "msg": f"Orden {pedido.po_codigo} actualizada. {estado}."}
    return RedirectResponse("/pedidos", status_code=303)


@app.post("/pedidos/{pid}/cancelar")
async def cancelar_pedido(pid: int, request: Request,
                          user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    pedido = pedido_del_usuario(db, pid, user, bloquear=True)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada o sin permisos de acceso."}
        return RedirectResponse("/pedidos", status_code=303)

    form_raw = await request.form()
    motivo = (form_raw.get("motivo_cancelacion") or "Orden cancelada con retención histórica de auditoría").strip()
    ahora_dt = datetime.now()
    estado_prev = pedido.estado or "En proceso"
    pedido.estado = "Cancelado"
    pedido.motivo_cancelacion = motivo
    pedido.fecha_cancelacion = ahora_dt
    pedido.actualizado_en = ahora_dt

    registrar_historial_pedido(
        db, pedido, estado_prev, "Cancelado",
        accion="CANCELACION_PO",
        motivo=motivo,
        user=user
    )
    db.commit()
    db.refresh(pedido)
    generar_salidas(db, pedido)
    request.session["flash"] = {
        "ok": True,
        "msg": f"Orden {pedido.po_codigo} cambiada a estado CANCELADO. Historial retenido en auditoría."
    }
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/pedidos/{pid}/documento")
def descargar_documento_adjunto(pid: int, nombre: str = "", user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        return RedirectResponse("/pedidos", status_code=303)
    arch_objetivo = nombre.strip() if nombre.strip() else (
        pedido.documentacion_archivo.split("|")[0].strip() if pedido.documentacion_archivo else ""
    )
    if not arch_objetivo:
        return RedirectResponse("/pedidos", status_code=303)
    ruta_doc = os.path.join(carpeta_po_local(pedido), nombre_seguro(arch_objetivo))
    if not os.path.exists(ruta_doc):
        return RedirectResponse("/pedidos", status_code=303)
    return FileResponse(ruta_doc, filename=os.path.basename(ruta_doc))


@app.get("/pedidos/{pid}/documento/preview")
def previsualizar_documento_adjunto(pid: int, nombre: str = "", user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        return HTMLResponse("<div style='padding:24px;font-family:sans-serif;color:#64748b;'>Orden no encontrada.</div>", status_code=404)
    arch_objetivo = nombre.strip() if nombre.strip() else (
        pedido.documentacion_archivo.split("|")[0].strip() if pedido.documentacion_archivo else ""
    )
    nombre_limpio = nombre_seguro(arch_objetivo) if arch_objetivo else ""
    ruta_doc = os.path.join(carpeta_po_local(pedido), nombre_limpio) if nombre_limpio else ""
    if ruta_doc and os.path.exists(ruta_doc):
        ext = os.path.splitext(nombre_limpio)[1].lower()
        mapa_mime = {
            ".pdf": "application/pdf",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".webp": "image/webp",
            ".gif": "image/gif",
            ".txt": "text/plain; charset=utf-8",
        }
        if ext in mapa_mime:
            return FileResponse(
                ruta_doc,
                media_type=mapa_mime[ext],
                headers={"Content-Disposition": f'inline; filename="{nombre_limpio}"'},
            )
    # Si es un archivo de Office o registro documental sin binario PDF/Imagen local, mostramos ficha técnica interactiva
    html_ficha = f"""<!DOCTYPE html>
    <html lang="es"><head><meta charset="utf-8"><script src="https://cdn.tailwindcss.com"></script></head>
    <body class="bg-slate-900 text-slate-100 p-6 font-sans">
        <div class="max-w-xl mx-auto bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl">
            <div class="flex items-center justify-between border-b border-slate-700 pb-3 mb-4">
                <div>
                    <span class="text-[10px] font-bold uppercase tracking-widest text-sky-400 block">Visor Documental Fly Logistic</span>
                    <h2 class="text-base font-extrabold text-white mt-0.5">{_xml_escape(arch_objetivo or pedido.documentacion or 'Soporte de Orden ' + pedido.po_codigo)}</h2>
                </div>
                <span class="px-2.5 py-1 rounded bg-blue-900/60 text-sky-300 text-xs font-bold">{_xml_escape(pedido.po_codigo)}</span>
            </div>
            <div class="space-y-2 text-xs text-slate-300 mb-5">
                <p><strong>Cliente:</strong> {_xml_escape(pedido.cliente.nombre if pedido.cliente else '')}</p>
                <p><strong>Asesor responsable:</strong> {_xml_escape(pedido.asesor.nombre if pedido.asesor else '')}</p>
                <p><strong>Referencia documental:</strong> {_xml_escape(pedido.documentacion or 'Registrada en carpeta operativa')}</p>
                <p><strong>Estado de la orden:</strong> {_xml_escape(pedido.estado)}</p>
            </div>
            <div class="flex flex-wrap gap-2">
                <a href="/pedidos/{pedido.id}/excel" target="_blank" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-xs font-bold">Descargar Formato COM-FM-08 (.xlsx)</a>
                {f'<a href="/pedidos/{pedido.id}/documento?nombre={_xml_escape(nombre_limpio)}" target="_blank" class="bg-sky-600 hover:bg-sky-700 text-white px-4 py-2 rounded-lg text-xs font-bold">Descargar Archivo Original</a>' if (ruta_doc and os.path.exists(ruta_doc)) else ''}
                {f'<a href="{_xml_escape(pedido.drive_folder_link or pedido.documentacion_drive_link)}" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-xs font-bold">Abrir en Google Drive</a>' if (pedido.drive_folder_link or pedido.documentacion_drive_link) else ''}
            </div>
        </div>
    </body></html>"""
    return HTMLResponse(content=html_ficha)


@app.get("/pedidos/{pid}/documento/{doc_id}")
def descargar_documento_por_id(pid: int, doc_id: int, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        return RedirectResponse("/pedidos", status_code=303)
    doc = db.query(models.PedidoDocumento).filter(
        models.PedidoDocumento.id == doc_id,
        models.PedidoDocumento.pedido_id == pedido.id
    ).first()
    if not doc:
        return RedirectResponse(f"/pedidos/{pid}/documento", status_code=303)
    ruta_doc = os.path.join(carpeta_po_local(pedido), doc.nombre_archivo)
    if not os.path.exists(ruta_doc):
        return RedirectResponse("/pedidos", status_code=303)
    return FileResponse(ruta_doc, filename=doc.nombre_archivo)


@app.api_route("/pedidos/{pid}/documento/{doc_id}/eliminar", methods=["POST", "DELETE"])
def eliminar_documento_individual(pid: int, doc_id: int, request: Request,
                                  user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        if request.method == "DELETE":
            return JSONResponse({"ok": False, "msg": msg_horario}, status_code=403)
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse(f"/pedidos/{pid}/editar", status_code=303)

    pedido = pedido_del_usuario(db, pid, user, bloquear=True)
    if not pedido:
        if request.method == "DELETE":
            return JSONResponse({"ok": False, "msg": "Orden no encontrada."}, status_code=404)
        return RedirectResponse("/pedidos", status_code=303)

    doc = db.query(models.PedidoDocumento).filter(
        models.PedidoDocumento.id == doc_id,
        models.PedidoDocumento.pedido_id == pedido.id
    ).with_for_update().first()

    if doc:
        ruta_local = os.path.join(carpeta_po_local(pedido), doc.nombre_archivo)
        if os.path.exists(ruta_local):
            try:
                os.remove(ruta_local)
            except Exception:
                pass
        if doc.drive_file_id:
            service = obtener_drive()
            if service:
                eliminar_archivo_de_drive(service, doc.drive_file_id)
        db.delete(doc)
        db.flush()
        db.refresh(pedido)
        sincronizar_resumen_documentos_po(pedido)
        db.commit()
        generar_salidas(db, pedido)

    if request.method == "DELETE" or "application/json" in (request.headers.get("accept") or ""):
        return JSONResponse({"ok": True, "msg": "Documento eliminado correctamente."})

    request.session["flash"] = {"ok": True, "msg": "Documento eliminado de la Orden de Compra."}
    return RedirectResponse(f"/pedidos/{pid}/editar", status_code=303)


@app.post("/pedidos/{pid}/eliminar")
def eliminar_pedido(pid: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    request.session["flash"] = {
        "ok": False,
        "msg": "Por políticas de auditoría y trazabilidad aeronáutica, las órdenes no se pueden eliminar. Si requieres anularla, utiliza la opción Cancelar PO manteniendo el historial."
    }
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/pedidos/{pid}/excel")
def descargar_excel(pid: int, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        return RedirectResponse("/pedidos", status_code=303)
    ruta, nombre = ruta_excel(pedido)
    generar_excel_com_fm_08(pedido, ruta)
    return FileResponse(ruta, filename=nombre, media_type=XLSX_MIME)


@app.get("/pedidos/{pid}/sincronizar_drive")
def sincronizar_pedido_drive_manual(pid: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    pedido = pedido_del_usuario(db, pid, user)
    if not pedido:
        request.session["flash"] = {"ok": False, "msg": "Orden no encontrada o no autorizada."}
        return RedirectResponse("/pedidos", status_code=303)
    res = generar_salidas(db, pedido)
    db.refresh(pedido)
    if pedido.drive_link or pedido.drive_folder_link:
        request.session["flash"] = {"ok": True, "msg": f"Orden {pedido.po} sincronizada exitosamente con Google Drive."}
    else:
        request.session["flash"] = {"ok": False, "msg": f"No se pudo sincronizar la orden con Google Drive ({res})."}
    return RedirectResponse("/pedidos", status_code=303)


# =========================================================================================
# NUEVO MÓDULO COTIZADOR (FORMATO COTIZACION VERSION 4)
# =========================================================================================
MAX_ITEMS_COTIZADOR = 13


def sugerir_siguiente_cotizacion(db: Session) -> str:
    todas = [c.numero for c in db.query(models.Cotizacion.numero).all() if c.numero]
    max_num = 0
    for qn in todas:
        nums = re.findall(r"\d+", str(qn))
        if nums:
            val = int(nums[-1])
            if val < 100000 and val > max_num:
                max_num = val
    return f"Q-2026-{(max_num + 1):03d}"


def _escribir_celda_segura(ws, coord: str, valor):
    celda = ws[coord]
    if isinstance(celda, MergedCell):
        for rng in ws.merged_cells.ranges:
            if coord in rng:
                celda = ws.cell(rng.min_row, rng.min_col)
                break
    celda.value = valor
    return celda


# --- Cotizador: llenado directo del XML de la plantilla (conserva logos, firma y recuadros) ---
COT_FILA_PRIMER_ITEM = 25
COT_FILA_ULTIMO_ITEM = 37
COT_IVA = 0.19


# ----------------------------------------------------------------------------------------
# Utilidades de celdas (manipulación de texto, no de árbol XML, para no alterar prefijos)
# ----------------------------------------------------------------------------------------
def _cot_col_a_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _cot_partir_ref(ref: str):
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def _cot_celda_xml(ref: str, estilo, valor=None, formula=None) -> str:
    """Construye el XML de una celda conservando su estilo (atributo s)."""
    s_attr = f' s="{estilo}"' if estilo is not None else ""
    if formula is not None:
        v = "" if valor is None else f"<v>{valor!r}</v>" if isinstance(valor, float) else f"<v>{valor}</v>"
        return f'<c r="{ref}"{s_attr}><f>{_xml_escape(formula)}</f>{v}</c>'
    if valor is None or valor == "":
        return f'<c r="{ref}"{s_attr}/>'
    if isinstance(valor, bool):
        valor = int(valor)
    if isinstance(valor, (int, float)):
        return f'<c r="{ref}"{s_attr}><v>{valor}</v></c>'
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{_xml_escape(str(valor))}</t></is></c>'


def _cot_set_celda(xml: str, ref: str, valor=None, formula=None, mantener_formula=False) -> str:
    col, fila = _cot_partir_ref(ref)

    fila_m = re.search(rf'<row r="{fila}"[^>]*?(?:/>|>.*?</row>)', xml, re.S)
    if not fila_m:
        raise ValueError(f"La fila {fila} no existe en la plantilla")
    fila_xml = fila_m.group(0)

    patron = re.compile(rf'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>(?P<cuerpo>.*?)</c>)', re.S)
    celda_m = patron.search(fila_xml)

    if celda_m:
        estilo_m = re.search(r'\bs="(\d+)"', celda_m.group("attrs"))
        estilo = estilo_m.group(1) if estilo_m else None
        if mantener_formula and celda_m.group("cuerpo") and "<f>" in celda_m.group("cuerpo"):
            formula = re.search(r"<f>(.*?)</f>", celda_m.group("cuerpo"), re.S).group(1)
            formula = formula.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
        nueva = _cot_celda_xml(ref, estilo, valor, formula)
        fila_nueva = fila_xml[:celda_m.start()] + nueva + fila_xml[celda_m.end():]
    else:
        # La celda no existe: se inserta respetando el orden de columnas
        nueva = _cot_celda_xml(ref, None, valor, formula)
        if fila_xml.endswith("/>"):
            fila_nueva = fila_xml[:-2] + ">" + nueva + "</row>"
        else:
            pos = fila_xml.rfind("</row>")
            for m in re.finditer(r'<c r="([A-Z]+)\d+"', fila_xml):
                if _cot_col_a_num(m.group(1)) > _cot_col_a_num(col):
                    pos = m.start()
                    break
            fila_nueva = fila_xml[:pos] + nueva + fila_xml[pos:]

    return xml[:fila_m.start()] + fila_nueva + xml[fila_m.end():]


def _cot_set_altura_fila(xml: str, fila: int, alto: float) -> str:
    m = re.search(rf'<row r="{fila}"([^>]*?)(/?)>', xml)
    if not m:
        return xml
    attrs = m.group(1)
    attrs = re.sub(r'\s(?:ht|customHeight)="[^"]*"', "", attrs)
    nuevo = f'<row r="{fila}"{attrs} ht="{alto}" customHeight="1"{m.group(2)}>'
    return xml[:m.start()] + nuevo + xml[m.end():]


# ----------------------------------------------------------------------------------------
# Función principal
# ----------------------------------------------------------------------------------------
def _generar_cotizacion_desde_plantilla_xml(plantilla: str, destino: str, cab: dict,
                                           items: list, transporte: float = 0.0,
                                           alto_fila_item: float = 66.0) -> str:
    """
    cab    : dict con quote_number, quote_date, expiration_date, prepared_by, to_company,
             ref_no, site, terms, phone, contact, email
    items  : lista de dicts {description, unit, qty, unit_price, remarks}
    """
    with zipfile.ZipFile(plantilla, "r") as zin:
        sheet = zin.read("xl/worksheets/sheet1.xml").decode("utf-8")

        # 1) Cabecera
        sheet = _cot_set_celda(sheet, "K7", cab.get("quote_number", ""))
        sheet = _cot_set_celda(sheet, "K8", cab.get("quote_date", ""))
        sheet = _cot_set_celda(sheet, "K9", cab.get("expiration_date", ""))

        # 2) Cliente / contacto
        sheet = _cot_set_celda(sheet, "D17", cab.get("prepared_by", ""))
        sheet = _cot_set_celda(sheet, "D18", cab.get("to_company", ""))
        sheet = _cot_set_celda(sheet, "D19", cab.get("ref_no", ""))
        sheet = _cot_set_celda(sheet, "D20", cab.get("site", ""))
        sheet = _cot_set_celda(sheet, "D21", cab.get("terms", ""))
        sheet = _cot_set_celda(sheet, "I19", cab.get("phone", ""))
        sheet = _cot_set_celda(sheet, "I20", cab.get("contact", ""))
        sheet = _cot_set_celda(sheet, "I21", cab.get("email", ""))

        # 3) Ítems (filas 25..37). Los estilos/fuentes de cada celda vienen de la plantilla.
        subtotal = 0.0
        for idx, fila in enumerate(range(COT_FILA_PRIMER_ITEM, COT_FILA_ULTIMO_ITEM + 1)):
            if idx < len(items):
                it = items[idx]
                qty = float(it.get("qty") or 0)
                precio = float(it.get("unit_price") or 0)
                total = qty * precio
                subtotal += total
                sheet = _cot_set_celda(sheet, f"B{fila}", idx + 1)
                sheet = _cot_set_celda(sheet, f"C{fila}", it.get("description", ""))
                sheet = _cot_set_celda(sheet, f"G{fila}", it.get("unit", "UND"))
                sheet = _cot_set_celda(sheet, f"H{fila}", qty)
                sheet = _cot_set_celda(sheet, f"I{fila}", precio)
                sheet = _cot_set_celda(sheet, f"J{fila}", total, formula=f"H{fila}*I{fila}")
                sheet = _cot_set_celda(sheet, f"K{fila}", it.get("remarks", ""))
                sheet = _cot_set_altura_fila(sheet, fila, alto_fila_item)
            else:
                for col in ("B", "C", "G", "H", "I", "K"):
                    sheet = _cot_set_celda(sheet, f"{col}{fila}", None)
                sheet = _cot_set_celda(sheet, f"J{fila}", 0, formula=f"H{fila}*I{fila}")

        # 4) Totales (con valor en caché para visores que no recalculan)
        iva = subtotal * COT_IVA
        transporte = float(transporte or 0)
        sheet = _cot_set_celda(sheet, "J42", subtotal, formula="SUM(J24:J38)")
        sheet = _cot_set_celda(sheet, "J43", iva, formula="J42*19%")
        sheet = _cot_set_celda(sheet, "J44", transporte)
        sheet = _cot_set_celda(sheet, "J45", subtotal + iva + transporte, formula="J42+J43+J44")

        # 5) Escribir el nuevo xlsx: todo igual a la plantilla salvo lo que se indica
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                nombre = info.filename
                if nombre == "xl/calcChain.xml":
                    continue  # se elimina: Excel lo regenera y evita el aviso de reparación
                data = zin.read(nombre)
                if nombre == "xl/worksheets/sheet1.xml":
                    data = sheet.encode("utf-8")
                elif nombre == "[Content_Types].xml":
                    t = data.decode("utf-8")
                    t = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", t)
                    data = t.encode("utf-8")
                elif nombre == "xl/_rels/workbook.xml.rels":
                    t = data.decode("utf-8")
                    t = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", t)
                    data = t.encode("utf-8")
                elif nombre == "xl/workbook.xml":
                    t = data.decode("utf-8")
                    if "fullCalcOnLoad" not in t:
                        t = t.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
                    data = t.encode("utf-8")
                zout.writestr(info, data)
    return destino


def sincronizar_cotizacion_en_drive(
    service,
    cot: models.Cotizacion,
    ruta_excel_local: str,
    nombre_excel: str
) -> str:
    raiz = "Cotizaciones a Clientes"
    id_raiz, _ = carpeta_drive(service, raiz)
    nombre_asesor = cot.asesor.nombre if cot.asesor else (cot.prepared_by or "General")
    id_asesor, _ = carpeta_drive(service, nombre_asesor, id_raiz)
    id_cot_folder, link_cot_folder = carpeta_drive(service, f"COT_{cot.quote_number}", id_asesor)

    res_excel = subir_archivo_a_carpeta_drive(service, ruta_excel_local, nombre_excel, id_cot_folder, XLSX_MIME)
    link_excel = res_excel.get("webViewLink", "") or link_cot_folder
    return link_excel


def generar_y_sincronizar_cotizacion(db: Session, cot: models.Cotizacion) -> str:
    plantilla_original = os.path.join(BASE_DIR, "FORMATO COTIZACION VERSION 4.xlsx")
    if not os.path.exists(plantilla_original):
        return "Plantilla FORMATO COTIZACION VERSION 4.xlsx no encontrada."

    if not cot.asesor and cot.asesor_id:
        cot.asesor = db.get(models.Asesor, cot.asesor_id)

    nombre_asesor = cot.asesor.nombre if cot.asesor else (cot.prepared_by or "Fly Logistic")
    cab = {
        "quote_number": cot.quote_number,
        "quote_date": cot.quote_date.strftime("%d/%m/%Y") if cot.quote_date else "",
        "expiration_date": cot.expiration_date.strftime("%d/%m/%Y") if cot.expiration_date else "",
        "to_company": cot.to_company or "",
        "ref_no": cot.ref_no or "",
        "phone": cot.phone or "",
        "site": cot.site or "",
        "contact": cot.contact or "",
        "terms": cot.terms or "",
        "email": cot.email or "",
        "prepared_by": nombre_asesor,
    }
    items = [{
        "item": idx + 1,
        "description": it.description,
        "unit": it.unit or "UND",
        "qty": float(it.qty or 1),
        "unit_price": float(it.unit_price_dec) + (float(it.cobro_hazmat_dec + it.cobro_certificado_dec) / float(it.qty or 1) if float(it.qty or 1) > 0 else 0.0),
        "remarks": it.remarks or "",
    } for idx, it in enumerate(cot.items)]

    nombre_archivo = f"Cotizacion_{nombre_seguro(cot.quote_number)}.xlsx"
    carpeta_local = os.path.join(BASE_DIR, "Cotizaciones", nombre_seguro(nombre_asesor))
    os.makedirs(carpeta_local, exist_ok=True)
    ruta_local = os.path.join(carpeta_local, nombre_archivo)

    try:
        _generar_cotizacion_desde_plantilla_xml(
            plantilla_original, ruta_local, cab, items, float(cot.transport_dec)
        )
        cot.archivo_excel = ruta_local
        db.commit()
    except Exception as exc:
        return f"Error al generar Excel: {exc}"

    try:
        service = obtener_drive()
        if service:
            link = sincronizar_cotizacion_en_drive(service, cot, ruta_local, nombre_archivo)
            if link:
                cot.drive_link = link
                db.commit()
                return "Cotización generada y sincronizada en Google Drive"
        return "Cotización generada localmente"
    except Exception as exc:
        print(f"[GOOGLE DRIVE COTIZACION SYNC]: {exc}")
        return f"Excel generado localmente ({exc})"


def sincronizar_todo_pendiente_drive(db: Session) -> dict:
    service = obtener_drive()
    if not service:
        return {
            "ok": False,
            "msg": "No hay conexión activa con Google Drive. Verifique su conexión de red o la autenticación en token.json."
        }

    pos_pendientes = db.query(models.Pedido).filter(
        ~models.Pedido.po.like("PO-QA-%"),
        (models.Pedido.drive_link == None) | (models.Pedido.drive_link == '')
    ).all()
    cots_pendientes = db.query(models.Cotizacion).filter(
        (models.Cotizacion.drive_link == None) | (models.Cotizacion.drive_link == '')
    ).all()

    exito_po = 0
    exito_cot = 0
    for p in pos_pendientes:
        try:
            res = generar_salidas(db, p)
            if "sincronizados" in res.lower() or p.drive_link:
                exito_po += 1
        except Exception:
            pass

    for c in cots_pendientes:
        try:
            res = generar_y_sincronizar_cotizacion(db, c)
            if "sincronizada" in res.lower() or c.drive_link:
                exito_cot += 1
        except Exception:
            pass

    return {
        "ok": True,
        "pos_sincronizadas": exito_po,
        "cots_sincronizadas": exito_cot,
        "total_pos": len(pos_pendientes),
        "total_cots": len(cots_pendientes),
        "msg": f"Sincronización con Google Drive completada: {exito_po} órdenes y {exito_cot} cotizaciones respaldadas exitosamente en la nube."
    }


def _datos_cotizador_desde_form(f, user: models.Asesor) -> dict:
    partes = f.getlist("part_number") or f.getlist("item_description")
    unidades = f.getlist("unit") or f.getlist("item_unit")
    cantidades = f.getlist("qty") or f.getlist("item_qty")
    precios = f.getlist("unit_price") or f.getlist("item_unit_price")
    hazmats = f.getlist("cobro_hazmat") or f.getlist("item_hazmat")
    certificados = f.getlist("cobro_certificado") or f.getlist("item_certificado")
    sin_ivas = f.getlist("sin_iva") or f.getlist("item_sin_iva")
    iva_pcts = f.getlist("iva_pct") or f.getlist("item_iva_pct")
    remarks_list = f.getlist("remarks") or f.getlist("item_remarks")

    items = []
    n = max(len(partes), len(unidades), len(cantidades), len(precios), len(remarks_list))
    for i in range(min(n, MAX_ITEMS_COTIZADOR)):
        desc = (partes[i] if i < len(partes) else "").strip()
        rem = (remarks_list[i] if i < len(remarks_list) else "").strip()
        pr = (precios[i] if i < len(precios) else "").strip()
        if not (desc or rem or pr):
            continue
        try:
            iva_val = float(str(iva_pcts[i] if i < len(iva_pcts) else "19").replace(",", ""))
        except Exception:
            iva_val = 19.0
        items.append({
            "item": len(items) + 1,
            "description": desc,
            "unit": (unidades[i] if i < len(unidades) else "UND").strip() or "UND",
            "qty": (cantidades[i] if i < len(cantidades) else "1").strip() or "1",
            "unit_price": pr or "0.00",
            "cobro_hazmat": (hazmats[i] if i < len(hazmats) else "0.00").strip() or "0.00",
            "cobro_certificado": (certificados[i] if i < len(certificados) else "0.00").strip() or "0.00",
            "sin_iva": "1" if (i < len(sin_ivas) and str(sin_ivas[i]).strip() in ("1", "true", "True", "on")) else "0",
            "iva_pct": iva_val,
            "remarks": rem,
        })

    if not items:
        items = [{
            "item": 1, "description": "", "unit": "UND", "qty": "1",
            "unit_price": "0.00", "cobro_hazmat": "0.00", "cobro_certificado": "0.00",
            "sin_iva": "0", "iva_pct": 19.0, "remarks": ""
        }]

    return {
        "quote_number": (f.get("quote_number") or "").strip(),
        "quote_date": (f.get("quote_date") or str(date.today())).strip(),
        "expiration_date": (f.get("expiration_date") or "").strip(),
        "prepared_by": (f.get("prepared_by") or user.nombre).strip() if user.rol in ROLES_SUPERIORES else user.nombre,
        "asesor_id": (f.get("asesor_id") or str(user.id)).strip() if user.rol in ROLES_SUPERIORES else str(user.id),
        "to_company": (f.get("to_company") or "").strip(),
        "ref_no": (f.get("ref_no") or "").strip(),
        "phone": (f.get("phone") or "").strip(),
        "site": (f.get("site") or "").strip(),
        "contact": (f.get("contact") or "").strip(),
        "terms": (f.get("terms") or "").strip(),
        "incoterm": (f.get("incoterm") or "").strip(),
        "delivery_time": (f.get("delivery_time") or "").strip(),
        "email": (f.get("email") or "").strip(),
        "tipo_operacion": (f.get("tipo_operacion") or "Nacional").strip(),
        "moneda": (f.get("moneda") or "COP").strip().upper() if (f.get("moneda") or "COP").strip().upper() in ("COP", "USD") else "COP",
        "transport": (f.get("transport") or "0.00").strip(),
        "items": items,
    }


def _guardar_cotizacion_en_bd(db: Session, user: models.Asesor, d: dict) -> models.Cotizacion:
    qn = (d.get("quote_number") or "").strip() or sugerir_siguiente_cotizacion(db)
    existente = db.query(models.Cotizacion).filter(
        func.lower(models.Cotizacion.numero) == qn.lower()
    ).with_for_update().first()

    def _parse_fecha_opc(s: str):
        s = (s or "").strip()
        if not s:
            return None
        try:
            return date.fromisoformat(s)
        except ValueError:
            return None

    try:
        transp_dec = Decimal(str(d.get("transport") or "0").replace(",", "")).quantize(Decimal("0.01"))
    except Exception:
        transp_dec = Decimal("0.00")

    asesor_id_final = user.id
    if user.rol in ROLES_SUPERIORES and str(d.get("asesor_id") or "").isdigit():
        asesor_id_final = int(d["asesor_id"])

    if existente:
        cot = existente
        if user.rol in ROLES_SUPERIORES and str(d.get("asesor_id") or "").isdigit():
            cot.asesor_id = int(d["asesor_id"])
    else:
        cot = models.Cotizacion(numero=qn, asesor_id=asesor_id_final)
        db.add(cot)

    cot.quote_number = qn
    cot.quote_date = _parse_fecha_opc(d.get("quote_date")) or date.today()
    cot.expiration_date = _parse_fecha_opc(d.get("expiration_date"))
    cot.prepared_by = d.get("prepared_by") or user.nombre if user.rol in ROLES_SUPERIORES else user.nombre
    cot.to_company = d.get("to_company") or ""
    cot.ref_no = d.get("ref_no") or ""
    cot.phone = d.get("phone") or ""
    cot.site = d.get("site") or ""
    cot.contact = d.get("contact") or ""
    terms_val = d.get("terms") or ""
    if d.get("incoterm") and d.get("incoterm") not in terms_val:
        terms_val = f"{terms_val} | Incoterm: {d.get('incoterm')}".strip(" |")
    if d.get("delivery_time") and d.get("delivery_time") not in terms_val:
        terms_val = f"{terms_val} | Entrega: {d.get('delivery_time')}".strip(" |")
    cot.terms = terms_val
    cot.email = d.get("email") or ""
    cot.moneda = d.get("moneda") or "COP"
    cot.tipo_operacion = d.get("tipo_operacion") or "Nacional"
    cot.transport = f"{transp_dec:.2f}"

    cot.items.clear()
    db.flush()

    for idx, it in enumerate(d.get("items", [])[:MAX_ITEMS_COTIZADOR]):
        desc = (it.get("description") or "").strip()
        if not desc:
            continue
        try:
            qty_f = float(str(it.get("qty") or "1").replace(",", ""))
        except Exception:
            qty_f = 1.0
        try:
            pr_dec = Decimal(str(it.get("unit_price") or "0").replace(",", "")).quantize(Decimal("0.01"))
        except Exception:
            pr_dec = Decimal("0.00")
        try:
            hz_dec = Decimal(str(it.get("cobro_hazmat") or "0").replace(",", "")).quantize(Decimal("0.01"))
        except Exception:
            hz_dec = Decimal("0.00")
        try:
            ce_dec = Decimal(str(it.get("cobro_certificado") or "0").replace(",", "")).quantize(Decimal("0.01"))
        except Exception:
            ce_dec = Decimal("0.00")

        sin_iva_bool = str(it.get("sin_iva", "0")).strip() in ("1", "true", "True", "on")
        try:
            iva_pct_f = float(str(it.get("iva_pct") or ("19" if d.get("moneda") == "COP" else "7")).replace(",", ""))
        except Exception:
            iva_pct_f = 19.0 if d.get("moneda") == "COP" else 7.0

        cot.items.append(models.CotizacionItem(
            orden=idx + 1,
            description=desc,
            unit=(it.get("unit") or "UND").strip() or "UND",
            qty=qty_f,
            unit_price=f"{pr_dec:.2f}",
            cobro_hazmat=f"{hz_dec:.2f}",
            cobro_certificado=f"{ce_dec:.2f}",
            sin_iva=sin_iva_bool,
            iva_item="0.00" if sin_iva_bool else f"{iva_pct_f:.2f}",
            remarks=(it.get("remarks") or "").strip(),
        ))

    db.flush()
    return cot


@app.get("/api/validar_cotizacion")
def api_validar_cotizacion(q: str = "", numero: str = "", excluir_id: int = 0, db: Session = Depends(get_db), user: models.Asesor = Depends(usuario_actual)):
    q_limpia = (q or numero or "").strip()
    sugerido = sugerir_siguiente_cotizacion(db)
    if not q_limpia:
        return JSONResponse({"existe": False, "sugerido": sugerido})
    query = db.query(models.Cotizacion).filter(func.lower(models.Cotizacion.numero) == q_limpia.lower())
    if excluir_id:
        query = query.filter(models.Cotizacion.id != excluir_id)
    existente = query.first()
    if existente:
        return JSONResponse({
            "existe": True,
            "cliente": existente.to_company or "",
            "asesor": existente.asesor.nombre if existente.asesor else (existente.prepared_by or ""),
            "sugerido": sugerido,
        })
    return JSONResponse({"existe": False, "sugerido": sugerido})


@app.get("/cotizador/nuevo", response_class=HTMLResponse)
def cotizador_nuevo(request: Request, cot_id: int = 0, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    hoy = str(date.today())
    expira_default = str(date.today() + timedelta(days=15))
    sugerida = sugerir_siguiente_cotizacion(db)
    d = {
        "quote_number": sugerida,
        "quote_date": hoy,
        "expiration_date": expira_default,
        "prepared_by": user.nombre,
        "asesor_id": str(user.id),
        "to_company": "",
        "ref_no": "",
        "phone": "",
        "site": "",
        "contact": "",
        "terms": "",
        "incoterm": "",
        "delivery_time": "",
        "email": "",
        "moneda": "COP",
        "tipo_operacion": "Nacional",
        "transport": "0.00",
        "items": [{
            "description": "", "unit": "UND", "qty": "1", "unit_price": "0.00",
            "cobro_hazmat": "0.00", "cobro_certificado": "0.00", "sin_iva": "0", "iva_pct": "19", "remarks": ""
        }]
    }

    if cot_id:
        c_existente = db.query(models.Cotizacion).filter(models.Cotizacion.id == cot_id).first()
        if c_existente:
            d = {
                "id": c_existente.id,
                "quote_number": c_existente.numero,
                "quote_date": str(c_existente.fecha or hoy),
                "expiration_date": str(c_existente.fecha_expiracion or expira_default),
                "prepared_by": c_existente.asesor.nombre if c_existente.asesor else (c_existente.prepared_by or user.nombre),
                "asesor_id": str(c_existente.asesor_id or user.id),
                "to_company": c_existente.empresa or "",
                "ref_no": c_existente.ref_no or "",
                "phone": c_existente.phone or "",
                "site": c_existente.site or "",
                "contact": c_existente.contact or "",
                "terms": c_existente.terms or "",
                "incoterm": "",
                "delivery_time": "",
                "email": c_existente.email or "",
                "moneda": c_existente.moneda or "COP",
                "tipo_operacion": c_existente.tipo_operacion or "Nacional",
                "transport": str(c_existente.transporte or "0.00"),
                "items": [{
                    "description": it.description or "",
                    "unit": it.unit or "UND",
                    "qty": str(it.qty or 1),
                    "unit_price": str(it.unit_price or "0.00"),
                    "cobro_hazmat": str(it.cobro_hazmat or "0.00"),
                    "cobro_certificado": str(it.cobro_certificado or "0.00"),
                    "sin_iva": "1" if it.sin_iva else "0",
                    "iva_pct": str(it.iva_porcentaje or ("19" if c_existente.moneda == "COP" else "7")),
                    "remarks": it.remarks or ""
                } for it in c_existente.items] or [{
                    "description": "", "unit": "UND", "qty": "1", "unit_price": "0.00",
                    "cobro_hazmat": "0.00", "cobro_certificado": "0.00", "sin_iva": "0", "iva_pct": "19", "remarks": ""
                }]
            }

    items_serializados = []
    for it in d.get("items", []):
        try:
            qty_val = float(str(it.get("qty") or "1").replace(",", ""))
        except Exception:
            qty_val = 1.0
        try:
            price_val = float(str(it.get("unit_price") or "0").replace(",", ""))
        except Exception:
            price_val = 0.0
        try:
            hz_val = float(str(it.get("cobro_hazmat") or "0").replace(",", ""))
        except Exception:
            hz_val = 0.0
        try:
            cert_val = float(str(it.get("cobro_certificado") or "0").replace(",", ""))
        except Exception:
            cert_val = 0.0
        sin_iva_val = str(it.get("sin_iva") or "0") in ("1", "true", "True")
        try:
            iva_val = float(str(it.get("iva_pct") or ("19" if d.get("moneda") == "COP" else "7")).replace(",", ""))
        except Exception:
            iva_val = 19.0 if d.get("moneda") == "COP" else 7.0

        items_serializados.append({
            "description": it.get("description") or "",
            "qty": qty_val,
            "unit": it.get("unit") or "UND",
            "unit_price": price_val,
            "hazmat": hz_val,
            "certificado": cert_val,
            "sin_iva": sin_iva_val,
            "iva_pct": iva_val,
            "remarks": it.get("remarks") or "",
        })
    d["items_json"] = json.dumps(items_serializados)

    asesores_disponibles = db.query(models.Asesor).order_by(models.Asesor.nombre).all()
    clientes_cot = [c[0] for c in db.query(models.Cotizacion.empresa).filter(models.Cotizacion.empresa.isnot(None), models.Cotizacion.empresa != '').distinct().all()]
    clientes_ped = [c[0] for c in db.query(models.Cliente.nombre).filter(models.Cliente.nombre.isnot(None), models.Cliente.nombre != '').distinct().all()]
    clientes = sorted(list(set([c.strip() for c in (clientes_cot + clientes_ped) if c and c.strip()])))

    return render(request, "cotizador_form.html", {
        "user": user,
        "d": d,
        "max_items": MAX_ITEMS_COTIZADOR,
        "fecha_hoy": hoy,
        "hoy": hoy,
        "expira_default": expira_default,
        "sugerencia_quote": d["quote_number"],
        "clientes": clientes,
        "unidades_medida": UNIDADES_MEDIDA,
        "asesores_disponibles": asesores_disponibles,
    })


@app.post("/cotizador/generar")
@app.post("/cotizador/guardar")
@app.post("/cotizador/exportar_excel")
async def cotizador_guardar(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/pedidos", status_code=303)

    f = await request.form()
    accion_posterior = (f.get("accion_posterior") or "").strip()
    d = _datos_cotizador_desde_form(f, user)
    if not d["quote_number"]:
        d["quote_number"] = sugerir_siguiente_cotizacion(db)

    items_validos = [it for it in d["items"] if (it.get("description") or "").strip()]
    asesores_disponibles = db.query(models.Asesor).order_by(models.Asesor.nombre).all()
    hoy = str(date.today())
    expira_default = str(date.today() + timedelta(days=15))
    clientes_cot = [c[0] for c in db.query(models.Cotizacion.empresa).filter(models.Cotizacion.empresa.isnot(None), models.Cotizacion.empresa != '').distinct().all()]
    clientes_ped = [c[0] for c in db.query(models.Cliente.nombre).filter(models.Cliente.nombre.isnot(None), models.Cliente.nombre != '').distinct().all()]
    clientes = sorted(list(set([c.strip() for c in (clientes_cot + clientes_ped) if c and c.strip()])))

    items_serializados = []
    for it in d.get("items", []):
        try:
            qty_val = float(str(it.get("qty") or "1").replace(",", ""))
        except Exception:
            qty_val = 1.0
        try:
            price_val = float(str(it.get("unit_price") or "0").replace(",", ""))
        except Exception:
            price_val = 0.0
        try:
            hz_val = float(str(it.get("cobro_hazmat") or "0").replace(",", ""))
        except Exception:
            hz_val = 0.0
        try:
            cert_val = float(str(it.get("cobro_certificado") or "0").replace(",", ""))
        except Exception:
            cert_val = 0.0
        sin_iva_val = str(it.get("sin_iva") or "0") in ("1", "true", "True")
        try:
            iva_val = float(str(it.get("iva_pct") or ("19" if d.get("moneda") == "COP" else "7")).replace(",", ""))
        except Exception:
            iva_val = 19.0 if d.get("moneda") == "COP" else 7.0

        items_serializados.append({
            "description": it.get("description") or "",
            "qty": qty_val,
            "unit": it.get("unit") or "UND",
            "unit_price": price_val,
            "hazmat": hz_val,
            "certificado": cert_val,
            "sin_iva": sin_iva_val,
            "iva_pct": iva_val,
            "remarks": it.get("remarks") or "",
        })
    d["items_json"] = json.dumps(items_serializados)

    if not items_validos:
        return render(request, "cotizador_form.html", {
            "user": user,
            "d": d,
            "max_items": MAX_ITEMS_COTIZADOR,
            "fecha_hoy": hoy,
            "hoy": hoy,
            "expira_default": expira_default,
            "sugerencia_quote": d["quote_number"],
            "clientes": clientes,
            "unidades_medida": UNIDADES_MEDIDA,
            "asesores_disponibles": asesores_disponibles,
            "error": "Debe registrar al menos un ítem con descripción o número de parte.",
        }, status=400)

    cot = None
    exito = False
    for intento in range(6):
        try:
            if intento > 0:
                d["quote_number"] = sugerir_siguiente_cotizacion(db)
            cot = _guardar_cotizacion_en_bd(db, user, d)
            db.commit()
            exito = True
            break
        except (IntegrityError, OperationalError) as exc:
            db.rollback()
            if intento == 5:
                return render(request, "cotizador_form.html", {
                    "user": user,
                    "d": d,
                    "max_items": MAX_ITEMS_COTIZADOR,
                    "fecha_hoy": hoy,
                    "hoy": hoy,
                    "expira_default": expira_default,
                    "sugerencia_quote": d["quote_number"],
                    "clientes": clientes,
                    "unidades_medida": UNIDADES_MEDIDA,
                    "asesores_disponibles": asesores_disponibles,
                    "error": f"Alta concurrencia operativa al guardar la cotización: {exc}. Por favor intente de nuevo.",
                }, status=400)
            time_lib.sleep(0.06 * (intento + 1))
            continue
        except Exception as exc:
            db.rollback()
            return render(request, "cotizador_form.html", {
                "user": user,
                "d": d,
                "max_items": MAX_ITEMS_COTIZADOR,
                "fecha_hoy": hoy,
                "hoy": hoy,
                "expira_default": expira_default,
                "sugerencia_quote": d["quote_number"],
                "clientes": clientes,
                "unidades_medida": UNIDADES_MEDIDA,
                "asesores_disponibles": asesores_disponibles,
                "error": f"No fue posible guardar la cotización: {exc}",
            }, status=400)

    res_drive = generar_y_sincronizar_cotizacion(db, cot)

    if accion_posterior == "descargar_excel" or request.url.path.endswith("/exportar_excel"):
        if not cot.archivo_excel or not os.path.exists(cot.archivo_excel):
            generar_y_sincronizar_cotizacion(db, cot)
        if cot.archivo_excel and os.path.exists(cot.archivo_excel):
            filename_download = f"Cotizacion_{nombre_seguro(cot.quote_number)}.xlsx"
            return FileResponse(
                cot.archivo_excel,
                filename=filename_download,
                media_type=XLSX_MIME
            )

    if accion_posterior == "convertir_po":
        request.session["flash"] = {"ok": True, "msg": f"Cotización {cot.quote_number} guardada exitosamente. {res_drive}."}
        return RedirectResponse(f"/pedidos/nuevo?cot_id={cot.id}", status_code=303)

    request.session["flash"] = {"ok": True, "msg": f"Cotización {cot.quote_number} guardada exitosamente. {res_drive}."}
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/cotizaciones", response_class=HTMLResponse)
def listar_cotizaciones(request: Request, q: str = "", page: int = 1,
                        user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    visibles = ids_asesores_visibles(db, user)
    query = db.query(models.Cotizacion).options(
        joinedload(models.Cotizacion.items),
        joinedload(models.Cotizacion.asesor)
    )
    if visibles is not None:
        query = query.filter(models.Cotizacion.asesor_id.in_(visibles))
    if q.strip():
        like = f"%{q.strip()}%"
        query = query.filter(
            models.Cotizacion.numero.ilike(like) |
            models.Cotizacion.empresa.ilike(like) |
            models.Cotizacion.prepared_by.ilike(like) |
            models.Cotizacion.contact.ilike(like)
        )
    todas = query.order_by(models.Cotizacion.id.desc()).all()
    total_cots = len(todas)
    total_paginas = max(1, (total_cots + REGISTROS_POR_PAGINA - 1) // REGISTROS_POR_PAGINA)
    pagina_actual = max(1, min(page, total_paginas))
    inicio = (pagina_actual - 1) * REGISTROS_POR_PAGINA
    cotizaciones_pag = todas[inicio:inicio + REGISTROS_POR_PAGINA]

    return render(request, "cotizaciones.html", {
        "user": user,
        "cotizaciones": cotizaciones_pag,
        "total_cotizaciones": total_cots,
        "pagina_actual": pagina_actual,
        "total_paginas": total_paginas,
        "q": q,
    })


@app.get("/cotizaciones/{cid}/excel")
def descargar_cotizacion_excel(cid: int, request: Request,
                               user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    cot = db.query(models.Cotizacion).filter(models.Cotizacion.id == cid).first()
    if not cot:
        return RedirectResponse("/pedidos", status_code=303)
    plantilla_original = os.path.join(BASE_DIR, "FORMATO COTIZACION VERSION 4.xlsx")
    if not os.path.exists(plantilla_original):
        return RedirectResponse("/pedidos", status_code=303)

    cab = {
        "quote_number": cot.quote_number,
        "quote_date": str(cot.quote_date or date.today()),
        "expiration_date": str(cot.expiration_date or ""),
        "prepared_by": cot.prepared_by or (cot.asesor.nombre if cot.asesor else user.nombre),
        "to_company": cot.to_company or "",
        "ref_no": cot.ref_no or "",
        "phone": cot.phone or "",
        "site": cot.site or "",
        "contact": cot.contact or "",
        "terms": cot.terms or "",
        "email": cot.email or "",
    }
    items = [{
        "item": idx + 1,
        "description": it.description,
        "unit": it.unit or "UND",
        "qty": float(it.qty or 1),
        "unit_price": float(it.unit_price_dec) + (float(it.cobro_hazmat_dec + it.cobro_certificado_dec) / float(it.qty or 1) if float(it.qty or 1) > 0 else 0.0),
        "remarks": it.remarks or "",
    } for idx, it in enumerate(cot.items)]

    nombre_archivo_final = f"Cotizacion_{nombre_seguro(cot.quote_number)}.xlsx"
    carpeta_cotizaciones = os.path.join(BASE_DIR, "Cotizaciones", nombre_seguro(user.nombre))
    os.makedirs(carpeta_cotizaciones, exist_ok=True)
    ruta_archivo_final = os.path.join(carpeta_cotizaciones, nombre_archivo_final)
    _generar_cotizacion_desde_plantilla_xml(
        plantilla_original, ruta_archivo_final, cab, items, float(cot.transport_dec)
    )
    return FileResponse(ruta_archivo_final, filename=nombre_archivo_final, media_type=XLSX_MIME)


@app.post("/cotizaciones/{cid}/eliminar")
def eliminar_cotizacion(cid: int, request: Request,
                        user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    request.session["flash"] = {
        "ok": False,
        "msg": "Por control y archivo comercial, las cotizaciones no se eliminan físicamente: permanecen preservadas en el histórico del sistema."
    }
    return RedirectResponse("/pedidos", status_code=303)


@app.get("/cotizaciones/{cid}/sincronizar_drive")
def sincronizar_cotizacion_drive_manual(cid: int, request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    cot = db.query(models.Cotizacion).filter(models.Cotizacion.id == cid).first()
    if not cot:
        request.session["flash"] = {"ok": False, "msg": "Cotización no encontrada."}
        return RedirectResponse("/pedidos", status_code=303)
    res = generar_y_sincronizar_cotizacion(db, cot)
    db.refresh(cot)
    if cot.drive_link:
        request.session["flash"] = {"ok": True, "msg": f"Cotización {cot.quote_number} sincronizada exitosamente con Google Drive."}
    else:
        request.session["flash"] = {"ok": False, "msg": f"No se pudo sincronizar la cotización con Google Drive ({res})."}
    return RedirectResponse("/pedidos", status_code=303)


@app.api_route("/drive/sincronizar_todo", methods=["GET", "POST"])
@app.api_route("/drive/sincronizar", methods=["GET", "POST"])
@app.api_route("/pedidos/sincronizar_drive_todo", methods=["GET", "POST"])
@app.api_route("/pedidos/sincronizar_drive", methods=["GET", "POST"])
def sincronizar_todo_drive_endpoint(request: Request, user: models.Asesor = Depends(usuario_actual), db: Session = Depends(get_db)):
    if user.rol == "Almacenista":
        return RedirectResponse("/inventario?warehouse=1", status_code=303)
    resultado = sincronizar_todo_pendiente_drive(db)
    request.session["flash"] = {"ok": resultado["ok"], "msg": resultado["msg"]}
    return RedirectResponse("/pedidos", status_code=303)


@app.api_route("/logo_fly.webp", methods=["GET", "HEAD"])
@app.api_route("/logo_fly.webp.webp", methods=["GET", "HEAD"])
@app.api_route("/logo", methods=["GET", "HEAD"])
@app.api_route("/logo-empresa", methods=["GET", "HEAD"])
@app.api_route("/favicon.ico", methods=["GET", "HEAD"])
@app.api_route("/static/logo.png", methods=["GET", "HEAD"])
@app.api_route("/static/logo_fly.webp", methods=["GET", "HEAD"])
@app.api_route("/static/{archivo:path}", methods=["GET", "HEAD"])
def servir_logo_corporativo(archivo: str = ""):
    candidatos = [
        os.path.join(BASE_DIR, "logo_fly.webp.webp"),
        os.path.join(BASE_DIR, "logo_fly.webp"),
        os.path.join(BASE_DIR, "cropped-Logo-Fly-Logistic-1-1-convertido-de-png-600x233.webp"),
        os.path.join(BASE_DIR, "logo_fly.png"),
        os.path.join(BASE_DIR, "logo.png"),
        "logo_fly.webp.webp",
        "logo_fly.webp",
    ]
    for ruta in candidatos:
        if os.path.exists(ruta):
            mime = "image/webp" if ruta.lower().endswith(".webp") else "image/png"
            return FileResponse(ruta, media_type=mime, headers={"Cache-Control": "public, max-age=86400"})
    svg_fallback = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="180" height="48" viewBox="0 0 180 48">'
        '<rect width="180" height="48" rx="8" fill="#002060"/>'
        '<text x="90" y="30" font-family="Arial,sans-serif" font-size="18" font-weight="bold" fill="#ffffff" text-anchor="middle">FLY LOGISTIC</text>'
        '</svg>'
    )
    return HTMLResponse(content=svg_fallback, media_type="image/svg+xml")


# =========================================================================================
# MÓDULO DE INVENTARIO AERONÁUTICO MULTIBODEGA (3 WAREHOUSES + PISTOLA CÓDIGO DE BARRAS)
# =========================================================================================
def _serializar_item_inventario(it: models.InventarioItem) -> dict:
    info = it.convencion_info
    return {
        "id": it.id,
        "warehouse": it.warehouse,
        "codigo_barras": it.codigo_barras,
        "parte_numero": it.parte_numero,
        "descripcion": it.descripcion,
        "medida": it.medida or "UND",
        "marca": it.marca or "",
        "normativa": it.normativa or "MIL-SPEC / HAZMAT",
        "lote": it.lote or "",
        "dom": it.dom.isoformat() if it.dom else "",
        "doe": it.doe.isoformat() if it.doe else "",
        "vigencia_dias": info["vigencia_dias"],
        "convencion": info["convencion"],
        "color": info["color"],
        "vencido": info["vencido"],
        "existencia": int(it.existencia or 0),
        "ubicacion": it.ubicacion or "",
        "actualizado_por": it.actualizado_por_nombre or "Sistema",
        "actualizado_en": it.actualizado_en.strftime("%Y-%m-%d %H:%M") if it.actualizado_en else "",
    }


def _registrar_movimiento_bodega(
    db: Session,
    item: models.InventarioItem,
    user: models.Asesor,
    tipo_movimiento: str,
    cant_anterior: int,
    cant_movimiento: int,
    cant_nueva: int,
    po_asociada: str = "",
    observacion: str = "",
) -> models.MovimientoInventario:
    ahora_bog = datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d %H:%M:%S")
    mov = models.MovimientoInventario(
        item_id=item.id,
        warehouse=item.warehouse,
        codigo_barras=item.codigo_barras,
        parte_numero=item.parte_numero,
        lote=item.lote or "",
        tipo_movimiento=tipo_movimiento,
        cantidad_anterior=cant_anterior,
        cantidad_movimiento=cant_movimiento,
        cantidad_nueva=cant_nueva,
        po_asociada=po_asociada or "",
        observacion=observacion or "",
        usuario_id=user.id if user else None,
        usuario_nombre=user.nombre if user else "Sistema",
        usuario_rol=user.rol if user else "",
        fecha_hora_bogota=ahora_bog,
        creado_en=datetime.now(),
    )
    db.add(mov)
    return mov


@app.get("/inventario", response_class=HTMLResponse)
def vista_inventario(
    request: Request,
    warehouse: str = "1",
    q: str = "",
    convencion: str = "",
    page: int = 1,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    w_limpio = (warehouse or "1").strip().lower()
    if w_limpio in ("1", "2", "3"):
        warehouse_activo = w_limpio
        w_int = int(w_limpio)
    else:
        warehouse_activo = "maestra"
        w_int = None

    query = db.query(models.InventarioItem)
    if w_int is not None:
        query = query.filter(models.InventarioItem.warehouse == w_int)
    if q.strip():
        like = f"%{q.strip()}%"
        query = query.filter(
            models.InventarioItem.codigo_barras.ilike(like)
            | models.InventarioItem.parte_numero.ilike(like)
            | models.InventarioItem.descripcion.ilike(like)
            | models.InventarioItem.lote.ilike(like)
            | models.InventarioItem.marca.ilike(like)
            | models.InventarioItem.ubicacion.ilike(like)
        )

    items_todos = query.order_by(models.InventarioItem.warehouse.asc(), models.InventarioItem.parte_numero.asc()).all()
    if convencion.strip():
        c_up = convencion.strip().upper()
        items_todos = [it for it in items_todos if c_up in it.convencion.upper()]

    # Métricas globales por cada Warehouse (1, 2 y 3) y Maestra
    todos_global = db.query(models.InventarioItem).all()
    resumen_warehouses = {}
    for wid in (1, 2, 3):
        lote_w = [x for x in todos_global if x.warehouse == wid]
        resumen_warehouses[wid] = {
            "warehouse": wid,
            "skus": len(lote_w),
            "existencia_total": sum(int(x.existencia or 0) for x in lote_w),
            "vencidos": sum(1 for x in lote_w if x.esta_vencido),
            "criticos": sum(1 for x in lote_w if x.convencion_info["color"] == "amber"),
            "vigentes": sum(1 for x in lote_w if x.convencion_info["color"] == "emerald"),
        }

    total_items = len(items_todos)
    total_paginas = max(1, (total_items + REGISTROS_POR_PAGINA - 1) // REGISTROS_POR_PAGINA)
    pagina_actual = max(1, min(page, total_paginas))
    inicio = (pagina_actual - 1) * REGISTROS_POR_PAGINA
    items_paginados = items_todos[inicio:inicio + REGISTROS_POR_PAGINA]

    movimientos_recientes = (
        db.query(models.MovimientoInventario)
        .order_by(models.MovimientoInventario.id.desc())
        .limit(12)
        .all()
    )
    pos_activas = (
        db.query(models.Pedido)
        .filter(models.Pedido.estado == "En proceso")
        .order_by(models.Pedido.id.desc())
        .limit(40)
        .all()
    )

    return render(request, "base.html", {
        "user": user,
        "modo_vista_inventario": True,
        "warehouse_activo": warehouse_activo,
        "q_inv": q,
        "convencion_f": convencion,
        "items_inventario": items_paginados,
        "todos_items_inventario": items_todos,
        "items_inventario_json": json.dumps([_serializar_item_inventario(x) for x in items_todos], ensure_ascii=False),
        "resumen_warehouses": resumen_warehouses,
        "movimientos_recientes": movimientos_recientes,
        "pos_activas": pos_activas,
        "unidades_medida": UNIDADES_MEDIDA,
        "pagina_actual": pagina_actual,
        "total_paginas": total_paginas,
        "total_items": total_items,
    })


@app.post("/inventario/guardar")
async def guardar_item_inventario(
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/inventario", status_code=303)

    f = await request.form()
    item_id_str = (f.get("item_id") or "").strip()
    w_str = (f.get("warehouse") or "1").strip()
    warehouse_num = int(w_str) if w_str in ("1", "2", "3") else 1

    codigo_barras = (f.get("codigo_barras") or "").strip()
    parte_numero = (f.get("parte_numero") or "").strip().upper()
    descripcion = (f.get("descripcion") or "").strip()
    medida = (f.get("medida") or "UND").strip() or "UND"
    marca = (f.get("marca") or "").strip()
    normativa = (f.get("normativa") or "MIL-SPEC / HAZMAT").strip()
    lote = (f.get("lote") or "").strip().upper()
    ubicacion = (f.get("ubicacion") or "").strip().upper()

    try:
        existencia_val = max(0, int(float(str(f.get("existencia") or "0").strip())))
    except Exception:
        existencia_val = 0

    def _parse_f(s):
        s = (s or "").strip()
        if not s:
            return None
        try:
            return date.fromisoformat(s)
        except ValueError:
            return None

    dom_date = _parse_f(f.get("dom"))
    doe_date = _parse_f(f.get("doe"))

    if not codigo_barras or not parte_numero or not descripcion:
        request.session["flash"] = {
            "ok": False,
            "msg": "Complete Código de Barras, Parte Número y Descripción para registrar el repuesto en bodega."
        }
        return RedirectResponse(f"/inventario?warehouse={warehouse_num}", status_code=303)

    if item_id_str.isdigit() and int(item_id_str) > 0:
        item = (
            db.query(models.InventarioItem)
            .with_for_update()
            .filter(models.InventarioItem.id == int(item_id_str))
            .first()
        )
    else:
        item = (
            db.query(models.InventarioItem)
            .with_for_update()
            .filter(
                models.InventarioItem.warehouse == warehouse_num,
                func.lower(models.InventarioItem.codigo_barras) == codigo_barras.lower(),
                func.lower(models.InventarioItem.lote) == lote.lower(),
            )
            .first()
        )

    if item:
        cant_ant = int(item.existencia or 0)
        item.warehouse = warehouse_num
        item.codigo_barras = codigo_barras
        item.parte_numero = parte_numero
        item.descripcion = descripcion
        item.medida = medida
        item.marca = marca
        item.normativa = normativa
        item.lote = lote
        item.dom = dom_date
        item.doe = doe_date
        item.existencia = existencia_val
        item.ubicacion = ubicacion
        item.actualizado_por_id = user.id
        item.actualizado_por_nombre = user.nombre
        db.flush()
        _registrar_movimiento_bodega(
            db, item, user, "ACTUALIZACION_FICHA",
            cant_ant, existencia_val - cant_ant, existencia_val,
            observacion=f"Actualización manual en Warehouse {warehouse_num}"
        )
        msg_ok = f"Repuesto {parte_numero} actualizado en Warehouse {warehouse_num}."
    else:
        item = models.InventarioItem(
            warehouse=warehouse_num,
            codigo_barras=codigo_barras,
            parte_numero=parte_numero,
            descripcion=descripcion,
            medida=medida,
            marca=marca,
            normativa=normativa,
            lote=lote,
            dom=dom_date,
            doe=doe_date,
            existencia=existencia_val,
            ubicacion=ubicacion,
            actualizado_por_id=user.id,
            actualizado_por_nombre=user.nombre,
        )
        db.add(item)
        db.flush()
        _registrar_movimiento_bodega(
            db, item, user, "INGRESO_INICIAL",
            0, existencia_val, existencia_val,
            observacion=f"Registro inicial de lote {lote or 'S/L'} en Warehouse {warehouse_num}"
        )
        msg_ok = f"Repuesto {parte_numero} registrado en Warehouse {warehouse_num}."

    db.commit()
    request.session["flash"] = {"ok": True, "msg": msg_ok}
    return RedirectResponse(f"/inventario?warehouse={warehouse_num}", status_code=303)


@app.post("/api/inventario/escanear")
async def api_escanear_codigo_barras(
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    """
    Endpoint de respuesta en milisegundos para operación con Pistola Lectora de Código de Barras.
    Valida horario RBAC, aislamiento de Warehouse, fecha DOE y actualiza existencia con FOR UPDATE.
    """
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        return JSONResponse({
            "ok": False,
            "alerta_restrictiva": True,
            "codigo_alerta": "HORARIO_BLOQUEADO",
            "titulo": "Operación fuera de horario autorizado",
            "msg": msg_horario,
        }, status_code=403)

    try:
        payload = await request.json()
    except Exception:
        f = await request.form()
        payload = dict(f)

    codigo = str(payload.get("codigo_barras") or "").strip()
    w_raw = str(payload.get("warehouse") or "1").strip()
    modo = str(payload.get("modo") or "SALIDA").strip().upper()
    po_asociada = str(payload.get("po_asociada") or "").strip()

    if not codigo:
        return JSONResponse({
            "ok": False,
            "alerta_restrictiva": False,
            "codigo_alerta": "CODIGO_VACIO",
            "msg": "Escanee o digite un código de barras válido.",
        }, status_code=400)

    warehouse_sel = int(w_raw) if w_raw in ("1", "2", "3") else None
    if warehouse_sel is None:
        return JSONResponse({
            "ok": False,
            "alerta_restrictiva": True,
            "codigo_alerta": "SELECCIONE_WAREHOUSE",
            "titulo": "Seleccione una Bodega Operativa (Warehouse 1, 2 o 3)",
            "msg": "Para usar la pistola lectora debe ubicarse en Warehouse 1, 2 o 3 (no en la vista Maestra consolidada).",
        }, status_code=400)

    # Buscar primero en el Warehouse activo con bloqueo FOR UPDATE
    item = (
        db.query(models.InventarioItem)
        .with_for_update()
        .filter(
            models.InventarioItem.warehouse == warehouse_sel,
            or_(
                func.lower(models.InventarioItem.codigo_barras) == codigo.lower(),
                func.lower(models.InventarioItem.parte_numero) == codigo.lower(),
            ),
        )
        .first()
    )

    if not item:
        # Verificar si pertenece a otro Warehouse para disparar la Alerta Sonora/Visual restrictiva
        en_otra_bodega = (
            db.query(models.InventarioItem)
            .filter(
                models.InventarioItem.warehouse != warehouse_sel,
                or_(
                    func.lower(models.InventarioItem.codigo_barras) == codigo.lower(),
                    func.lower(models.InventarioItem.parte_numero) == codigo.lower(),
                ),
            )
            .first()
        )
        if en_otra_bodega:
            return JSONResponse({
                "ok": False,
                "alerta_restrictiva": True,
                "codigo_alerta": "OTRO_WAREHOUSE",
                "titulo": f"ALERTA RESTRICTIVA: Material ubicado en Warehouse {en_otra_bodega.warehouse}",
                "msg": (
                    f"El código «{codigo}» ({en_otra_bodega.parte_numero} - {en_otra_bodega.descripcion}) "
                    f"pertenece a Warehouse {en_otra_bodega.warehouse} (Ubicación: {en_otra_bodega.ubicacion or 'General'}), "
                    f"no al Warehouse {warehouse_sel} actual."
                ),
                "item": _serializar_item_inventario(en_otra_bodega),
            }, status_code=409)

        return JSONResponse({
            "ok": False,
            "alerta_restrictiva": True,
            "codigo_alerta": "NO_ENCONTRADO",
            "titulo": "Código de Barras no registrado",
            "msg": f"No existe ningún ítem con el código «{codigo}» en el Warehouse {warehouse_sel}.",
        }, status_code=404)

    # Alerta Sonora/Visual restrictiva si el DOE está VENCIDO
    if item.esta_vencido:
        return JSONResponse({
            "ok": False,
            "alerta_restrictiva": True,
            "codigo_alerta": "DOE_VENCIDO",
            "titulo": "BLOQUEO DE CALIDAD AERONÁUTICA: LOTE VENCIDO (DOE EXPIRADO)",
            "msg": (
                f"El ítem {item.parte_numero} (Lote {item.lote or 'S/L'}) expiró el {item.doe} "
                f"({abs(item.vigencia_dias or 0)} días vencido). Normativa aeronáutica prohíbe su despacho."
            ),
            "item": _serializar_item_inventario(item),
        }, status_code=422)

    cant_ant = int(item.existencia or 0)
    delta = 1 if modo == "ENTRADA" else -1

    if delta < 0 and cant_ant < 1:
        return JSONResponse({
            "ok": False,
            "alerta_restrictiva": True,
            "codigo_alerta": "SIN_STOCK",
            "titulo": f"Stock agotado en Warehouse {warehouse_sel}",
            "msg": f"El repuesto {item.parte_numero} tiene existencia 0 en Warehouse {warehouse_sel}. No se puede descontar.",
            "item": _serializar_item_inventario(item),
        }, status_code=409)

    cant_nueva = cant_ant + delta
    ahora_dt = datetime.now()
    item.existencia = cant_nueva
    item.actualizado_por_id = user.id
    item.actualizado_por_nombre = user.nombre
    item.actualizado_en = ahora_dt

    tipo_mov = "PISTOLA_ENTRADA" if delta > 0 else ("DESPACHO_PO" if po_asociada else "PISTOLA_SALIDA")
    mov_creado = _registrar_movimiento_bodega(
        db, item, user, tipo_mov,
        cant_ant, delta, cant_nueva,
        po_asociada=po_asociada,
        observacion=f"Lectura de pistola Tera HW0009 ({modo}) en Warehouse {warehouse_sel}"
    )
    db.commit()
    db.refresh(item)
    db.refresh(mov_creado)

    todos_global = db.query(models.InventarioItem).all()
    resumen_warehouses = {}
    for wid in (1, 2, 3):
        lote_w = [x for x in todos_global if x.warehouse == wid]
        resumen_warehouses[str(wid)] = {
            "warehouse": wid,
            "skus": len(lote_w),
            "existencia_total": sum(int(x.existencia or 0) for x in lote_w),
            "vencidos": sum(1 for x in lote_w if x.esta_vencido),
            "criticos": sum(1 for x in lote_w if x.convencion_info["color"] == "amber"),
            "vigentes": sum(1 for x in lote_w if x.convencion_info["color"] == "emerald"),
        }

    return JSONResponse({
        "ok": True,
        "alerta_restrictiva": False,
        "dispositivo": "Tera HW0009",
        "modo": modo,
        "delta": delta,
        "msg": f"{item.parte_numero}: {'+1 unidad ingresada' if delta > 0 else '-1 unidad descontada'} en Warehouse {warehouse_sel} (Stock disponible: {cant_nueva} {item.medida}).",
        "item": _serializar_item_inventario(item),
        "movimiento": {
            "id": mov_creado.id,
            "fecha_hora_bogota": mov_creado.fecha_hora_bogota,
            "warehouse": mov_creado.warehouse,
            "codigo_barras": mov_creado.codigo_barras,
            "parte_numero": mov_creado.parte_numero,
            "lote": mov_creado.lote or "",
            "tipo_movimiento": mov_creado.tipo_movimiento,
            "cantidad_anterior": mov_creado.cantidad_anterior,
            "cantidad_movimiento": mov_creado.cantidad_movimiento,
            "cantidad_nueva": mov_creado.cantidad_nueva,
            "po_asociada": mov_creado.po_asociada or "",
            "observacion": mov_creado.observacion or "",
            "usuario_nombre": mov_creado.usuario_nombre or user.nombre,
            "usuario_rol": mov_creado.usuario_rol or user.rol,
        },
        "resumen_warehouses": resumen_warehouses,
    })


@app.get("/api/inventario/validar_stock")
def api_validar_stock_bodega(
    parte_numero: str = "",
    codigo_barras: str = "",
    warehouse: int = 1,
    cantidad: int = 1,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    """
    Validación de Stock Aislada: verifica disponibilidad únicamente en el Warehouse seleccionado.
    """
    criterio = (parte_numero or codigo_barras or "").strip()
    if not criterio:
        return JSONResponse({"ok": False, "msg": "Indique Parte Número o Código de Barras."}, status_code=400)

    item = (
        db.query(models.InventarioItem)
        .filter(
            models.InventarioItem.warehouse == warehouse,
            or_(
                func.lower(models.InventarioItem.parte_numero) == criterio.lower(),
                func.lower(models.InventarioItem.codigo_barras) == criterio.lower(),
            ),
        )
        .first()
    )
    if not item:
        otro = (
            db.query(models.InventarioItem)
            .filter(
                models.InventarioItem.warehouse != warehouse,
                or_(
                    func.lower(models.InventarioItem.parte_numero) == criterio.lower(),
                    func.lower(models.InventarioItem.codigo_barras) == criterio.lower(),
                ),
            )
            .first()
        )
        return JSONResponse({
            "ok": False,
            "disponible": False,
            "warehouse_solicitado": warehouse,
            "en_otro_warehouse": otro.warehouse if otro else None,
            "msg": (
                f"Sin existencia en Warehouse {warehouse}."
                + (f" (Disponible en Warehouse {otro.warehouse}: {otro.existencia} {otro.medida})" if otro else "")
            ),
        })

    if item.esta_vencido:
        return JSONResponse({
            "ok": False,
            "disponible": False,
            "vencido": True,
            "warehouse_solicitado": warehouse,
            "existencia": item.existencia,
            "msg": f"Material VENCIDO en Warehouse {warehouse} (DOE: {item.doe}). Despacho bloqueado.",
            "item": _serializar_item_inventario(item),
        })

    alcanza = int(item.existencia or 0) >= max(1, int(cantidad or 1))
    return JSONResponse({
        "ok": alcanza,
        "disponible": alcanza,
        "warehouse_solicitado": warehouse,
        "existencia": int(item.existencia or 0),
        "msg": (
            f"Stock validado en Warehouse {warehouse}: {item.existencia} {item.medida} disponibles."
            if alcanza
            else f"Stock insuficiente en Warehouse {warehouse}: solicita {cantidad}, disponible {item.existencia}."
        ),
        "item": _serializar_item_inventario(item),
    })


@app.post("/inventario/despachar_po")
async def despachar_material_a_po(
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/inventario", status_code=303)

    f = await request.form()
    item_id_str = (f.get("item_id") or "").strip()
    warehouse_str = (f.get("warehouse") or "1").strip()
    po_codigo = (f.get("po_numero") or "").strip()
    try:
        cant_despachar = max(1, int(float(str(f.get("cantidad") or "1").strip())))
    except Exception:
        cant_despachar = 1

    warehouse_sel = int(warehouse_str) if warehouse_str in ("1", "2", "3") else 1
    if not item_id_str.isdigit():
        request.session["flash"] = {"ok": False, "msg": "Seleccione un ítem válido para despachar."}
        return RedirectResponse(f"/inventario?warehouse={warehouse_sel}", status_code=303)

    item = (
        db.query(models.InventarioItem)
        .with_for_update()
        .filter(models.InventarioItem.id == int(item_id_str))
        .first()
    )
    if not item:
        request.session["flash"] = {"ok": False, "msg": "El repuesto seleccionado no existe."}
        return RedirectResponse(f"/inventario?warehouse={warehouse_sel}", status_code=303)

    # Validación de Stock Aislada por Warehouse
    if item.warehouse != warehouse_sel:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Bloqueo de bodega: El repuesto {item.parte_numero} pertenece a Warehouse {item.warehouse}, no a Warehouse {warehouse_sel}."
        }
        return RedirectResponse(f"/inventario?warehouse={warehouse_sel}", status_code=303)

    if item.esta_vencido:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Bloqueo de calidad: El lote {item.lote or 'S/L'} de {item.parte_numero} tiene fecha DOE vencida ({item.doe})."
        }
        return RedirectResponse(f"/inventario?warehouse={warehouse_sel}", status_code=303)

    if int(item.existencia or 0) < cant_despachar:
        request.session["flash"] = {
            "ok": False,
            "msg": f"Stock insuficiente en Warehouse {warehouse_sel}: solicitó {cant_despachar} {item.medida}, pero solo hay {item.existencia} en esta bodega."
        }
        return RedirectResponse(f"/inventario?warehouse={warehouse_sel}", status_code=303)

    cant_ant = int(item.existencia or 0)
    cant_nueva = cant_ant - cant_despachar
    item.existencia = cant_nueva
    item.actualizado_por_id = user.id
    item.actualizado_por_nombre = user.nombre
    item.actualizado_en = datetime.utcnow()

    _registrar_movimiento_bodega(
        db, item, user, "ASIGNACION_PO",
        cant_ant, -cant_despachar, cant_nueva,
        po_asociada=po_codigo,
        observacion=f"Despacho aislado desde Warehouse {warehouse_sel} hacia orden {po_codigo}"
    )
    db.commit()
    request.session["flash"] = {
        "ok": True,
        "msg": f"Se despacharon {cant_despachar} {item.medida} de {item.parte_numero} desde Warehouse {warehouse_sel} para la orden {po_codigo}."
    }
    return RedirectResponse(f"/inventario?warehouse={warehouse_sel}", status_code=303)


@app.post("/inventario/{item_id}/eliminar")
def eliminar_item_inventario(
    item_id: int,
    request: Request,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    if user.rol not in ROLES_SUPERIORES:
        request.session["flash"] = {"ok": False, "msg": "Solo Jefatura, Gerencia o Administración pueden dar de baja registros de bodega."}
        return RedirectResponse("/inventario", status_code=303)
    permitido, msg_horario = validar_horario_asesor(user)
    if not permitido:
        request.session["flash"] = {"ok": False, "msg": msg_horario}
        return RedirectResponse("/inventario", status_code=303)

    item = db.query(models.InventarioItem).with_for_update().filter(models.InventarioItem.id == item_id).first()
    w_ret = item.warehouse if item else 1
    if item:
        pn = item.parte_numero
        db.delete(item)
        db.commit()
        request.session["flash"] = {"ok": True, "msg": f"Registro {pn} eliminado de Warehouse {w_ret}."}
    return RedirectResponse(f"/inventario?warehouse={w_ret}", status_code=303)


def _estilizar_hoja_inventario_excel(ws, titulo: str, items: list[models.InventarioItem]):
    ws.merge_cells("A1:M1")
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="002060")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    encabezados = [
        "WAREHOUSE",
        "CÓDIGO DE BARRAS",
        "PARTE NÚMERO",
        "DESCRIPCIÓN",
        "MEDIDA",
        "MARCA",
        "NORMATIVA",
        "LOTE",
        "DOM (FABRICACIÓN)",
        "DOE (EXPIRACIÓN)",
        "VIGENCIA (DÍAS)",
        "CONVENCIONES",
        "EXISTENCIA / UBICACIÓN",
    ]
    ws.append(encabezados)
    ws.row_dimensions[2].height = 22
    borde_fino = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )
    for col_idx in range(1, len(encabezados) + 1):
        c = ws.cell(row=2, column=col_idx)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1E3A8A")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde_fino

    for fila_idx, it in enumerate(items, start=3):
        info = it.convencion_info
        vig_txt = f"{info['vigencia_dias']} días" if info["vigencia_dias"] is not None else "N/A"
        ws.append([
            f"Warehouse {it.warehouse}",
            it.codigo_barras,
            it.parte_numero,
            it.descripcion,
            it.medida or "UND",
            it.marca or "",
            it.normativa or "MIL-SPEC",
            it.lote or "",
            it.dom.isoformat() if it.dom else "N/A",
            it.doe.isoformat() if it.doe else "N/A",
            vig_txt,
            info["convencion"],
            f"{it.existencia} {it.medida or 'UND'} ({it.ubicacion or 'General'})",
        ])
        for c_i in range(1, len(encabezados) + 1):
            celda = ws.cell(row=fila_idx, column=c_i)
            celda.border = borde_fino
            celda.font = Font(size=10, color="0F172A")
        # Colorear automáticamente la celda de CONVENCIONES según fecha DOE
        celda_conv = ws.cell(row=fila_idx, column=12)
        celda_conv.fill = PatternFill("solid", fgColor=info["hex_bg"])
        celda_conv.font = Font(bold=True, size=10, color=info["hex_fg"])
        celda_conv.alignment = Alignment(horizontal="center", vertical="center")

    anchos = [15, 20, 20, 38, 12, 20, 20, 16, 18, 18, 16, 24, 28]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


@app.get("/inventario/exportar/warehouse/{warehouse_id}")
def exportar_excel_warehouse(
    warehouse_id: int,
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    wid = warehouse_id if warehouse_id in (1, 2, 3) else 1
    fecha_str = datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d")
    nombre_archivo = f"Inventario_Warehouse_{wid}_{fecha_str}.xlsx"

    items = (
        db.query(models.InventarioItem)
        .filter(models.InventarioItem.warehouse == wid)
        .order_by(models.InventarioItem.parte_numero.asc())
        .all()
    )

    wb = Workbook()
    ws = wb.active
    ws.title = f"Warehouse_{wid}"
    _estilizar_hoja_inventario_excel(
        ws,
        f"CONTROL DE INVENTARIO Y TRAZABILIDAD AERONÁUTICA — WAREHOUSE {wid} ({fecha_str})",
        items,
    )

    carpeta_inv = os.path.join(BASE_DIR, "Reportes_Gerenciales")
    os.makedirs(carpeta_inv, exist_ok=True)
    ruta_final = os.path.join(carpeta_inv, nombre_archivo)
    wb.save(ruta_final)
    return FileResponse(ruta_final, filename=nombre_archivo, media_type=XLSX_MIME)


@app.get("/inventario/exportar/maestro")
def exportar_excel_inventario_maestro(
    user: models.Asesor = Depends(usuario_actual),
    db: Session = Depends(get_db),
):
    fecha_str = datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d")
    nombre_archivo = f"Inventario_Maestro_Consolidado_{fecha_str}.xlsx"

    todos = (
        db.query(models.InventarioItem)
        .order_by(models.InventarioItem.parte_numero.asc(), models.InventarioItem.warehouse.asc())
        .all()
    )

    wb = Workbook()
    ws_cons = wb.active
    ws_cons.title = "Consolidado_3_Warehouses"
    ws_cons.merge_cells("A1:J1")
    ws_cons["A1"] = f"REPORTE MAESTRO CONSOLIDADO MULTIBODEGA (WAREHOUSES 1, 2 Y 3) — FLY LOGISTIC S.A.S. ({fecha_str})"
    ws_cons["A1"].font = Font(bold=True, size=13, color="FFFFFF")
    ws_cons["A1"].fill = PatternFill("solid", fgColor="002060")
    ws_cons["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_cons.row_dimensions[1].height = 28

    enc_m = [
        "PARTE NÚMERO",
        "CÓDIGO DE BARRAS",
        "DESCRIPCIÓN",
        "MEDIDA",
        "MARCA",
        "STOCK WAREHOUSE 1",
        "STOCK WAREHOUSE 2",
        "STOCK WAREHOUSE 3",
        "EXISTENCIA TOTAL",
        "DETALLE DE UBICACIONES Y LOTES POR BODEGA",
    ]
    ws_cons.append(enc_m)
    for col_idx in range(1, len(enc_m) + 1):
        c = ws_cons.cell(row=2, column=col_idx)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1E3A8A")
        c.alignment = Alignment(horizontal="center", vertical="center")

    agrupado = {}
    for it in todos:
        clave = (it.parte_numero or "").strip().upper()
        if clave not in agrupado:
            agrupado[clave] = {
                "parte_numero": it.parte_numero,
                "codigo_barras": it.codigo_barras,
                "descripcion": it.descripcion,
                "medida": it.medida or "UND",
                "marca": it.marca or "",
                "w1": 0,
                "w2": 0,
                "w3": 0,
                "total": 0,
                "ubicaciones": [],
            }
        cant = int(it.existencia or 0)
        if it.warehouse == 1:
            agrupado[clave]["w1"] += cant
        elif it.warehouse == 2:
            agrupado[clave]["w2"] += cant
        elif it.warehouse == 3:
            agrupado[clave]["w3"] += cant
        agrupado[clave]["total"] += cant
        agrupado[clave]["ubicaciones"].append(
            f"W{it.warehouse} [{it.ubicacion or 'General'} | Lote: {it.lote or 'S/L'} | DOE: {it.doe or 'N/A'} | {it.convencion}: {cant}]"
        )

    for g in agrupado.values():
        ws_cons.append([
            g["parte_numero"],
            g["codigo_barras"],
            g["descripcion"],
            g["medida"],
            g["marca"],
            g["w1"],
            g["w2"],
            g["w3"],
            g["total"],
            " ; ".join(g["ubicaciones"]),
        ])

    anchos_m = [20, 20, 36, 12, 18, 18, 18, 18, 18, 65]
    for i, w in enumerate(anchos_m, start=1):
        ws_cons.column_dimensions[get_column_letter(i)].width = w

    # Hojas detalladas por cada Warehouse (1, 2 y 3)
    for wid in (1, 2, 3):
        ws_w = wb.create_sheet(f"Detalle_Warehouse_{wid}")
        items_w = [x for x in todos if x.warehouse == wid]
        _estilizar_hoja_inventario_excel(
            ws_w,
            f"DETALLE TRAZABILIDAD AERONÁUTICA — WAREHOUSE {wid} ({fecha_str})",
            items_w,
        )

    carpeta_inv = os.path.join(BASE_DIR, "Reportes_Gerenciales")
    os.makedirs(carpeta_inv, exist_ok=True)
    ruta_final = os.path.join(carpeta_inv, nombre_archivo)
    wb.save(ruta_final)
    return FileResponse(ruta_final, filename=nombre_archivo, media_type=XLSX_MIME)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
