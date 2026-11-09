import re
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal

from sqlalchemy import (Column, Integer, String, Date, DateTime, Text,
                        Numeric, ForeignKey, Boolean)
from sqlalchemy.orm import relationship

from database import Base

TRANSPORTADORAS_OFICIALES = ["FedEx", "Coordinadora", "DHL", "Clickmail", "Inter Rapidísimo"]
ROLES_SISTEMA = ["Administrador", "Gerencia", "Jefe de area", "Asesor", "Almacenista"]
ESTADOS = ["En proceso", "Entregado", "Cancelado"]
ESTADOS_PRODUCTO = ["PENDIENTE", "DESPACHADO", "EN FLY", "LLEGADO FLY", "ENTREGADO"]
ESTADOS_ADUANA = ["No aplica / Pendiente", "En revisión aduanera", "Pago de aranceles", "Liberado / Nacionalizado"]
CONVENCIONES_INVENTARIO = ["VIGENTE", "PRECAUCION", "CRITICO", "VENCIDOS"]
WAREHOUSES = [1, 2, 3]


def normalizar(texto: str) -> str:
    t = unicodedata.normalize("NFKD", texto or "")
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.lower().split())


def contar_dias_habiles(fecha_inicio: date, fecha_fin: date) -> int:
    if not fecha_inicio or not fecha_fin or fecha_inicio == fecha_fin:
        return 0
    signo = 1 if fecha_fin > fecha_inicio else -1
    actual = fecha_inicio
    pasos = 0
    while actual != fecha_fin:
        actual += timedelta(days=signo)
        if actual.weekday() < 5:
            pasos += signo
    return pasos


def detectar_transportadora_tracking(
    tracking: str | None,
    llega_click_mail: str | None = "No",
    transportadora_manual: str | None = "",
    tipo_operacion: str = "Internacional"
) -> dict:
    codigo = (tracking or "").strip()
    manual = (transportadora_manual or "").strip()
    es_cm = str(llega_click_mail or "").strip().lower() in ("si", "sí", "true", "1", "clickmail")

    upper_code = codigo.upper()
    manual_norm = normalizar(manual)

    carrier = ""
    if manual_norm in ("fedex", "fed ex", "fdx"):
        carrier = "FedEx"
    elif manual_norm in ("dhl", "dhl express"):
        carrier = "DHL"
    elif manual_norm in ("coordinadora", "coord"):
        carrier = "Coordinadora"
    elif manual_norm in ("clickmail", "click mail"):
        carrier = "Clickmail"
    elif manual_norm in ("interrapidisimo", "inter rapidisimo", "inter rapidísimo", "inter"):
        carrier = "Inter Rapidísimo"

    if not carrier and upper_code:
        if upper_code.startswith(("FX", "FDX")) or (upper_code.isdigit() and len(upper_code) in (12, 15, 20, 22)):
            carrier = "FedEx"
        elif upper_code.startswith(("DHL", "JD", "JJD", "GM")) or (upper_code.isdigit() and len(upper_code) == 10):
            carrier = "DHL"
        elif upper_code.startswith(("COORD", "CRD", "73", "20")) or (upper_code.isdigit() and len(upper_code) == 11):
            carrier = "Coordinadora"
        elif upper_code.startswith(("CM", "CLK", "CLICK", "MIA")):
            carrier = "Clickmail"
        elif upper_code.startswith(("INT", "IR")) or (upper_code.isdigit() and len(upper_code) in (9, 10, 11, 12) and tipo_operacion == "Nacional" and manual_norm == "inter"):
            carrier = "Inter Rapidísimo"

    if not carrier:
        if es_cm:
            carrier = "Clickmail"
        elif codigo:
            carrier = "Coordinadora" if tipo_operacion == "Nacional" else "FedEx"
        else:
            carrier = "Clickmail" if es_cm else ("Coordinadora" if tipo_operacion == "Nacional" else "Pendiente")

    url_base = {
        "FedEx": "https://www.fedex.com/es-co/tracking.html",
        "DHL": "https://www.dhl.com/co-es/home/rastreo.html",
        "Coordinadora": "https://coordinadora.com/rastreo/rastreo-de-guia/",
        "Clickmail": "https://clickmail.com.co/tracking",
        "Inter Rapidísimo": "https://siguetuenvio.interrapidisimo.com/",
    }
    url = url_base.get(carrier, "")
    if codigo:
        q_code = re.sub(r"[^A-Za-z0-9\-]", "", codigo)
        if carrier == "FedEx":
            url = f"https://www.fedex.com/fedextrack/?trknbr={q_code}"
        elif carrier == "DHL":
            url = f"https://www.dhl.com/co-es/home/rastreo.html?tracking-id={q_code}"
        elif carrier == "Coordinadora":
            url = f"https://coordinadora.com/rastreo/rastreo-de-guia/?guia={q_code}"
        elif carrier == "Clickmail":
            url = f"https://clickmail.com.co/tracking?guia={q_code}"
        elif carrier == "Inter Rapidísimo":
            url = f"https://siguetuenvio.interrapidisimo.com/?guia={q_code}"

    etiqueta_txt = f"{carrier}: {codigo}" if codigo else (f"{carrier} (Warehouse EE. UU.)" if es_cm else (f"Despacho {carrier}" if carrier and carrier != "Pendiente" else "Sin guía registrada"))
    return {
        "carrier": carrier if carrier != "Pendiente" else "",
        "transportadora": carrier,
        "guia": codigo,
        "codigo": codigo or ("Consolidado Clickmail" if es_cm else "Sin asignar"),
        "tiene_codigo": bool(codigo),
        "es_clickmail": es_cm or carrier == "Clickmail",
        "url": url,
        "etiqueta": etiqueta_txt,
        "estado_tracking": etiqueta_txt,
    }


def parsear_desglose_pagos(metodo_pago_str: str | None) -> list[dict]:
    if not metodo_pago_str:
        return []
    partes = [
        p.strip()
        for p in re.split(r",\s*(?=[A-Za-zÁÉÍÓÚáéíóúÑñ])", metodo_pago_str)
        if p.strip()
    ]
    resultado = []
    for p in partes:
        if ":" in p:
            nombre, monto_raw = p.split(":", 1)
            monto_limpio = re.sub(r"[^\d.\-]", "", monto_raw)
            try:
                monto_dec = Decimal(monto_limpio) if monto_limpio else Decimal("0.00")
            except Exception:
                monto_dec = Decimal("0.00")
            resultado.append({
                "metodo": nombre.strip(),
                "monto": monto_dec,
                "monto_str": f"{monto_dec:.2f}" if monto_dec > 0 else ""
            })
        else:
            resultado.append({
                "metodo": p.strip(),
                "monto": Decimal("0.00"),
                "monto_str": ""
            })
    return resultado


class Asesor(Base):
    __tablename__ = "asesores"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(100), nullable=False)
    nombre_norm = Column(String(100), unique=True, nullable=False, index=True)
    username = Column(String(50), unique=True, nullable=True, index=True)
    password_hash = Column(String(300), nullable=True)
    email = Column(String(150), nullable=True)
    telefono = Column(String(50), nullable=True)

    rol = Column(String(50), nullable=False, default="Asesor")
    jefe_id = Column(Integer, ForeignKey("asesores.id"), nullable=True)

    reset_token_hash = Column(String(64), nullable=True, index=True)
    reset_expira = Column(DateTime, nullable=True)
    desbloqueo_horario_hasta = Column(DateTime, nullable=True)
    desbloqueo_clave_id = Column(Integer, nullable=True)
    creado_en = Column(DateTime, default=datetime.utcnow)

    pedidos = relationship("Pedido", back_populates="asesor")
    cotizaciones = relationship("Cotizacion", back_populates="asesor")
    jefe = relationship("Asesor", remote_side=[id], backref="subordinados")


class Cliente(Base):
    __tablename__ = "clientes"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(150), unique=True, nullable=False)

    pedidos = relationship("Pedido", back_populates="cliente")


class Pedido(Base):
    __tablename__ = "pedidos"

    id = Column(Integer, primary_key=True, index=True)
    po = Column(String(50), unique=True, nullable=False, index=True)
    tipo = Column(String(15), nullable=False)
    fecha_inicial = Column(Date, nullable=False)

    cliente_id = Column(Integer, ForeignKey("clientes.id"), nullable=False)
    asesor_id = Column(Integer, ForeignKey("asesores.id"), nullable=False)

    metodo_pago = Column(String(250), nullable=True)
    iva_porcentaje = Column(Numeric(5, 2), nullable=False, default=19)
    valor_nacionalizacion = Column(Numeric(14, 2), nullable=False, default=0)
    shipping = Column(Numeric(14, 2), nullable=False, default=0)

    plazo_entrega = Column(Date, nullable=True)
    fecha_entrega_cliente = Column(Date, nullable=True)
    numero_remision = Column(String(50), nullable=True)

    transportadora_nacional = Column(String(60), nullable=True, default="")
    guia_nacional = Column(String(80), nullable=True, default="")

    documentacion = Column(String(500), nullable=True)
    documentacion_archivo = Column(String(500), nullable=True)
    documentacion_drive_link = Column(String(500), nullable=True)

    estado = Column(String(25), nullable=False, default="En proceso")
    observaciones = Column(Text, nullable=True)
    motivo_cancelacion = Column(Text, nullable=True)
    fecha_cancelacion = Column(DateTime, nullable=True)
    cotizacion_origen_id = Column(Integer, nullable=True)

    drive_link = Column(String(500), nullable=True)
    drive_folder_link = Column(String(500), nullable=True)
    creado_en = Column(DateTime, default=datetime.utcnow)
    actualizado_en = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    cliente = relationship("Cliente", back_populates="pedidos")
    asesor = relationship("Asesor", back_populates="pedidos")
    items = relationship("ItemPedido", back_populates="pedido",
                         cascade="all, delete-orphan", order_by="ItemPedido.id")
    internacional = relationship("PedidoInternacional", back_populates="pedido",
                                 uselist=False, cascade="all, delete-orphan")
    documentos = relationship("PedidoDocumento", back_populates="pedido",
                              cascade="all, delete-orphan", order_by="PedidoDocumento.id")
    historial = relationship("PedidoHistorial", back_populates="pedido",
                             cascade="all, delete-orphan", order_by="PedidoHistorial.id.desc()")

    @property
    def po_codigo(self) -> str:
        val = (self.po or "").strip()
        return val if val.upper().startswith("PO") else f"PO-{val}"

    @property
    def moneda(self) -> str:
        return "COP" if self.tipo == "Nacional" else "USD"

    @property
    def total_hazmat(self) -> Decimal:
        return sum((i.cobro_hazmat_dec for i in self.items), Decimal(0))

    @property
    def total_certificados(self) -> Decimal:
        return sum((i.cobro_certificado_dec for i in self.items), Decimal(0))

    @property
    def subtotal(self) -> Decimal:
        return sum((i.subtotal_item for i in self.items), Decimal(0))

    @property
    def total_impuestos(self) -> Decimal:
        return sum((i.impuesto_item for i in self.items), Decimal(0))

    @property
    def total_productos_con_impuesto(self) -> Decimal:
        return sum((i.total_item for i in self.items), Decimal(0))

    @property
    def valor_nacionalizacion_dec(self) -> Decimal:
        return Decimal(str(self.valor_nacionalizacion or 0))

    @property
    def total(self) -> Decimal:
        return self.total_productos_con_impuesto + self.valor_nacionalizacion_dec

    @property
    def desglose_pagos(self) -> list[dict]:
        return parsear_desglose_pagos(self.metodo_pago)

    @property
    def tracking_info(self) -> dict:
        if self.tipo == "Internacional" and self.internacional:
            return detectar_transportadora_tracking(
                tracking=self.internacional.tracking,
                llega_click_mail=self.internacional.llega_click_mail,
                transportadora_manual=getattr(self.internacional, "transportadora", "") or self.transportadora_nacional,
                tipo_operacion="Internacional"
            )
        return detectar_transportadora_tracking(
            tracking=self.guia_nacional or self.numero_remision,
            llega_click_mail="No",
            transportadora_manual=self.transportadora_nacional,
            tipo_operacion="Nacional"
        )

    @property
    def lista_documentos(self) -> list:
        docs = list(self.documentos or [])
        nombres_registrados = {d.nombre_archivo for d in docs}
        if self.documentacion_archivo:
            for nombre_leg in [x.strip() for x in self.documentacion_archivo.split("|") if x.strip()]:
                if nombre_leg not in nombres_registrados:
                    docs.append(PedidoDocumento(
                        id=0,
                        pedido_id=self.id,
                        nombre_archivo=nombre_leg,
                        drive_link=self.documentacion_drive_link or "",
                        drive_file_id=""
                    ))
        return docs

    @property
    def transportadora(self) -> str:
        return self.transportadora_nacional or ""

    @transportadora.setter
    def transportadora(self, val: str):
        self.transportadora_nacional = val or ""

    @property
    def numero_guia_nacional(self) -> str:
        return self.guia_nacional or ""

    @numero_guia_nacional.setter
    def numero_guia_nacional(self, val: str):
        self.guia_nacional = val or ""

    @property
    def progreso_operativo(self) -> dict:
        nombres = ["Apertura", "Compra", "En tránsito", "Aduana/Bodega", "Entregado"]
        if self.estado == "Entregado" or self.fecha_entrega_cliente:
            idx = 5
        elif self.estado == "Cancelado":
            idx = 1
        else:
            tiene_compra = any(bool(it.fecha_compra_proveedor) for it in (self.items or []))
            tiene_guia = bool(self.guia_nacional or (self.internacional and self.internacional.tracking))
            tiene_aduana = bool(
                self.internacional
                and (self.internacional.fecha_nacionalizacion or self.internacional.fecha_click_mail)
            )
            if tiene_aduana:
                idx = 4
            elif tiene_guia:
                idx = 3
            elif tiene_compra:
                idx = 2
            else:
                idx = 1
        return {
            "indice_actual": idx,
            "etapa_actual": "Cancelada" if self.estado == "Cancelado" else nombres[idx - 1],
            "etapas": [
                {"numero": i + 1, "nombre": n, "completado": (i + 1) <= idx}
                for i, n in enumerate(nombres)
            ],
        }

    @property
    def dias_habiles_info(self) -> dict:
        etiqueta_po = self.po_codigo
        if self.estado == "Cancelado":
            return {
                "dias": 0,
                "orden_urgencia": 999999,
                "en_contra": False,
                "etiqueta": f"{etiqueta_po} | CANCELADO",
                "detalle": "Orden cancelada con historial retenido",
                "color": "amber",
                "estado_plazo": "Cancelado",
            }
        if not self.plazo_entrega:
            return {
                "dias": None,
                "orden_urgencia": 999998,
                "en_contra": False,
                "etiqueta": f"{etiqueta_po} | SIN PLAZO",
                "detalle": "Sin fecha límite",
                "color": "slate",
                "estado_plazo": "Sin plazo",
            }
        ref = self.fecha_entrega_cliente if self.fecha_entrega_cliente else date.today()
        dias = contar_dias_habiles(ref, self.plazo_entrega)
        if dias >= 0:
            palabra = "DÍA HÁBIL" if dias == 1 else "DÍAS HÁBILES"
            return {
                "dias": dias,
                "orden_urgencia": dias,
                "en_contra": False,
                "etiqueta": f"{etiqueta_po} | {dias} {palabra}",
                "detalle": f"{dias} {palabra.lower()} a favor",
                "color": "emerald",
                "estado_plazo": "En plazo",
            }
        else:
            dias_abs = abs(dias)
            palabra = "DÍA EN CONTRA" if dias_abs == 1 else "DÍAS EN CONTRA"
            return {
                "dias": dias,
                "orden_urgencia": dias,
                "en_contra": True,
                "etiqueta": f"{etiqueta_po} | {dias_abs} {palabra}",
                "detalle": f"{dias_abs} {palabra.lower()}",
                "color": "red",
                "estado_plazo": "Vencido",
            }


class ItemPedido(Base):
    __tablename__ = "pedido_items"

    id = Column(Integer, primary_key=True, index=True)
    pedido_id = Column(Integer, ForeignKey("pedidos.id"), nullable=False, index=True)

    proveedor = Column(String(150), nullable=False)
    fecha_compra_proveedor = Column(Date, nullable=True)

    parte_numero = Column(String(80), nullable=False)
    parte_numero_alterno = Column(String(80), nullable=True, default="")
    unidad_medida = Column(String(40), nullable=False, default="UND")
    descripcion = Column(String(255), nullable=False)
    cantidad = Column(Integer, nullable=False)
    precio_unitario = Column(Numeric(14, 2), nullable=False)

    cobro_hazmat = Column(Numeric(14, 2), nullable=False, default=0)
    cobro_certificado = Column(Numeric(14, 2), nullable=False, default=0)
    sin_iva = Column(Integer, nullable=False, default=0)

    iva_porcentaje = Column(Numeric(5, 2), nullable=False, default=19)
    estado_producto = Column(String(30), nullable=False, default="PENDIENTE DE PAGO")

    pedido = relationship("Pedido", back_populates="items")

    @property
    def cobro_hazmat_dec(self) -> Decimal:
        return Decimal(str(self.cobro_hazmat or 0))

    @property
    def cobro_certificado_dec(self) -> Decimal:
        return Decimal(str(self.cobro_certificado or 0))

    @property
    def subtotal_base_item(self) -> Decimal:
        return Decimal(str(self.cantidad or 0)) * Decimal(str(self.precio_unitario or 0))

    @property
    def subtotal_item(self) -> Decimal:
        return self.subtotal_base_item + self.cobro_hazmat_dec + self.cobro_certificado_dec

    @property
    def es_exento_iva(self) -> bool:
        return bool(self.sin_iva) or Decimal(str(self.iva_porcentaje or 0)) == Decimal(0)

    @property
    def impuesto_item(self) -> Decimal:
        if self.es_exento_iva:
            return Decimal("0.00")
        return self.subtotal_item * (Decimal(str(self.iva_porcentaje or 0)) / Decimal(100))

    @property
    def total_item(self) -> Decimal:
        return self.subtotal_item + self.impuesto_item


class PedidoInternacional(Base):
    __tablename__ = "pedidos_internacionales"

    pedido_id = Column(Integer, ForeignKey("pedidos.id"), primary_key=True)
    llega_click_mail = Column(String(10), nullable=True, default="No")
    fecha_click_mail = Column(Date, nullable=True)
    transportadora = Column(String(60), nullable=True, default="FedEx")
    tracking = Column(String(80), nullable=True)
    estado_nacionalizacion = Column(String(50), nullable=True)
    fecha_nacionalizacion = Column(Date, nullable=True)

    pedido = relationship("Pedido", back_populates="internacional")


class PedidoDocumento(Base):
    __tablename__ = "pedido_documentos"

    id = Column(Integer, primary_key=True, index=True)
    pedido_id = Column(Integer, ForeignKey("pedidos.id"), nullable=False, index=True)
    nombre_archivo = Column(String(255), nullable=False)
    drive_file_id = Column(String(120), nullable=True, default="")
    drive_link = Column(String(500), nullable=True, default="")
    creado_en = Column(DateTime, default=datetime.utcnow)

    pedido = relationship("Pedido", back_populates="documentos")


class PedidoHistorial(Base):
    __tablename__ = "pedido_historial"

    id = Column(Integer, primary_key=True, index=True)
    pedido_id = Column(Integer, ForeignKey("pedidos.id"), nullable=True, index=True)
    po_numero = Column(String(50), nullable=False, index=True)
    accion = Column(String(60), nullable=True, default="CAMBIO_ESTADO")
    estado_anterior = Column(String(30), nullable=True)
    estado_nuevo = Column(String(30), nullable=False)
    usuario_id = Column(Integer, nullable=True)
    usuario_nombre = Column(String(100), nullable=True)
    motivo = Column(Text, nullable=True)
    resumen_items = Column(Text, nullable=True)
    total_snapshot = Column(Numeric(14, 2), nullable=False, default=0)
    moneda = Column(String(10), nullable=False, default="COP")
    creado_en = Column(DateTime, default=datetime.utcnow)

    pedido = relationship("Pedido", back_populates="historial")

    @property
    def po_codigo(self) -> str:
        return self.po_numero or ""

    @po_codigo.setter
    def po_codigo(self, val: str):
        self.po_numero = val or ""


class Cotizacion(Base):
    __tablename__ = "cotizaciones"

    id = Column(Integer, primary_key=True, index=True)
    numero = Column(String(50), unique=True, nullable=False, index=True)
    fecha = Column(Date, nullable=False, default=date.today)
    fecha_expiracion = Column(Date, nullable=True)

    asesor_id = Column(Integer, ForeignKey("asesores.id"), nullable=False, index=True)
    prepared_by = Column(String(120), nullable=True)
    empresa = Column(String(180), nullable=False, default="")
    ref_no = Column(String(100), nullable=True, default="")
    site = Column(String(150), nullable=True, default="")
    terms = Column(String(150), nullable=True, default="")
    phone = Column(String(80), nullable=True, default="")
    contact = Column(String(120), nullable=True, default="")
    email = Column(String(150), nullable=True, default="")

    tipo_operacion = Column(String(20), nullable=False, default="Nacional")
    moneda = Column(String(10), nullable=False, default="COP")
    iva_porcentaje = Column(Numeric(5, 2), nullable=False, default=19)
    transporte = Column(Numeric(14, 2), nullable=False, default=0)
    estado = Column(String(30), nullable=False, default="Activa")
    po_convertida = Column(String(50), nullable=True, default="")

    archivo_excel = Column(String(300), nullable=True, default="")
    drive_link = Column(String(500), nullable=True, default="")
    creado_en = Column(DateTime, default=datetime.utcnow)
    actualizado_en = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    asesor = relationship("Asesor", back_populates="cotizaciones")
    items = relationship("CotizacionItem", back_populates="cotizacion",
                         cascade="all, delete-orphan", order_by="CotizacionItem.id")

    @property
    def quote_number(self) -> str:
        return self.numero or ""

    @quote_number.setter
    def quote_number(self, val: str):
        self.numero = val or ""

    @property
    def quote_date(self):
        return self.fecha

    @quote_date.setter
    def quote_date(self, val):
        self.fecha = val

    @property
    def expiration_date(self):
        return self.fecha_expiracion

    @expiration_date.setter
    def expiration_date(self, val):
        self.fecha_expiracion = val

    @property
    def to_company(self) -> str:
        return self.empresa or ""

    @to_company.setter
    def to_company(self, val: str):
        self.empresa = val or ""

    @property
    def transport(self):
        return self.transporte

    @transport.setter
    def transport(self, val):
        self.transporte = val

    @property
    def transport_dec(self) -> Decimal:
        return self.transporte_dec

    @property
    def codigo_cotizacion(self) -> str:
        val = (self.numero or "").strip()
        return val if val.upper().startswith(("Q-", "COT-")) else f"Q-{val}"

    @property
    def total_hazmat(self) -> Decimal:
        return sum((i.cobro_hazmat_dec for i in self.items), Decimal(0))

    @property
    def total_certificados(self) -> Decimal:
        return sum((i.cobro_certificado_dec for i in self.items), Decimal(0))

    @property
    def subtotal(self) -> Decimal:
        return sum((i.subtotal_item for i in self.items), Decimal(0))

    @property
    def total_impuestos(self) -> Decimal:
        return sum((i.impuesto_item for i in self.items), Decimal(0))

    @property
    def transporte_dec(self) -> Decimal:
        return Decimal(str(self.transporte or 0))

    @property
    def total(self) -> Decimal:
        return self.subtotal + self.total_impuestos + self.transporte_dec

    @property
    def progreso_operativo(self) -> dict:
        nombres = ["Apertura", "Compra", "En tránsito", "Aduana/Bodega", "Entregado"]
        if self.estado == "Entregado" or self.fecha_entrega_cliente:
            idx = 5
        elif self.estado == "Cancelado":
            idx = 1
        else:
            tiene_compra = any(bool(it.fecha_compra_proveedor) for it in (self.items or []))
            tiene_guia = bool(self.guia_nacional or (self.internacional and self.internacional.tracking))
            tiene_aduana = bool(
                self.internacional
                and (self.internacional.fecha_nacionalizacion or self.internacional.fecha_click_mail)
            )
            if tiene_aduana:
                idx = 4
            elif tiene_guia:
                idx = 3
            elif tiene_compra:
                idx = 2
            else:
                idx = 1
        return {
            "indice_actual": idx,
            "etapa_actual": "Cancelada" if self.estado == "Cancelado" else nombres[idx - 1],
            "etapas": [
                {"numero": i + 1, "nombre": n, "completado": (i + 1) <= idx}
                for i, n in enumerate(nombres)
            ],
        }

    @property
    def dias_habiles_info(self) -> dict:
        etiqueta_cot = self.codigo_cotizacion
        if self.estado == "Cancelada":
            return {
                "dias": 0,
                "orden_urgencia": 999999,
                "en_contra": False,
                "etiqueta": f"{etiqueta_cot} | CANCELADA",
                "detalle": "Cotización cancelada",
                "color": "slate"
            }
        if not self.fecha_expiracion:
            return {
                "dias": None,
                "orden_urgencia": 999997,
                "en_contra": False,
                "etiqueta": f"{etiqueta_cot} | VIGENTE",
                "detalle": "Sin fecha de expiración",
                "color": "emerald"
            }
        dias = contar_dias_habiles(date.today(), self.fecha_expiracion)
        if dias >= 0:
            palabra = "DÍA VIGENCIA" if dias == 1 else "DÍAS VIGENCIA"
            return {
                "dias": dias,
                "orden_urgencia": dias,
                "en_contra": False,
                "etiqueta": f"{etiqueta_cot} | {dias} {palabra}",
                "detalle": f"{dias} {palabra.lower()}",
                "color": "emerald"
            }
        dias_abs = abs(dias)
        return {
            "dias": dias,
            "orden_urgencia": dias,
            "en_contra": True,
            "etiqueta": f"{etiqueta_cot} | EXPIRADA ({dias_abs}D)",
            "detalle": f"Expirada hace {dias_abs} día(s)",
            "color": "red"
        }


class CotizacionItem(Base):
    __tablename__ = "cotizacion_items"

    id = Column(Integer, primary_key=True, index=True)
    cotizacion_id = Column(Integer, ForeignKey("cotizaciones.id"), nullable=False, index=True)

    item_num = Column(Integer, nullable=False, default=1)
    description = Column(String(255), nullable=False)
    unit = Column(String(40), nullable=False, default="UND")
    qty = Column(Numeric(12, 2), nullable=False, default=1)
    unit_price = Column(Numeric(14, 2), nullable=False, default=0)
    cobro_hazmat = Column(Numeric(14, 2), nullable=False, default=0)
    cobro_certificado = Column(Numeric(14, 2), nullable=False, default=0)
    sin_iva = Column(Integer, nullable=False, default=0)
    iva_porcentaje = Column(Numeric(5, 2), nullable=False, default=19)
    remarks = Column(String(255), nullable=True, default="")

    cotizacion = relationship("Cotizacion", back_populates="items")

    @property
    def orden(self) -> int:
        return self.item_num or 1

    @orden.setter
    def orden(self, val: int):
        self.item_num = val or 1

    @property
    def iva_item(self):
        return self.iva_porcentaje

    @iva_item.setter
    def iva_item(self, val):
        self.iva_porcentaje = val

    @property
    def unit_price_dec(self) -> Decimal:
        return Decimal(str(self.unit_price or 0))

    @property
    def cobro_hazmat_dec(self) -> Decimal:
        return Decimal(str(self.cobro_hazmat or 0))

    @property
    def cobro_certificado_dec(self) -> Decimal:
        return Decimal(str(self.cobro_certificado or 0))

    @property
    def subtotal_base_item(self) -> Decimal:
        return Decimal(str(self.qty or 0)) * Decimal(str(self.unit_price or 0))

    @property
    def subtotal_item(self) -> Decimal:
        return self.subtotal_base_item + self.cobro_hazmat_dec + self.cobro_certificado_dec

    @property
    def es_exento_iva(self) -> bool:
        return bool(self.sin_iva) or Decimal(str(self.iva_porcentaje or 0)) == Decimal(0)

    @property
    def impuesto_item(self) -> Decimal:
        if self.es_exento_iva:
            return Decimal("0.00")
        return self.subtotal_item * (Decimal(str(self.iva_porcentaje or 0)) / Decimal(100))

    @property
    def total_item(self) -> Decimal:
        return self.subtotal_item + self.impuesto_item


WAREHOUSES_OFICIALES = (1, 2, 3)


def calcular_convencion_vigencia(doe: date | None, fecha_ref: date | None = None) -> dict:
    """
    Calcula automáticamente la vigencia en días y el semáforo de convenciones aeronáuticas
    (HAZMAT / MIL-SPEC) basándose en la Fecha de Expiración (DOE).
    """
    hoy = fecha_ref or date.today()
    if not doe:
        return {
            "vigencia_dias": None,
            "convencion": "SIN VENCIMIENTO / N/A",
            "color": "slate",
            "hex_bg": "475569",
            "hex_fg": "FFFFFF",
            "vencido": False,
            "alerta": False,
        }
    dias = (doe - hoy).days
    if dias < 0:
        return {
            "vigencia_dias": dias,
            "convencion": "VENCIDOS",
            "color": "red",
            "hex_bg": "B91C1C",
            "hex_fg": "FFFFFF",
            "vencido": True,
            "alerta": True,
        }
    if dias <= 30:
        return {
            "vigencia_dias": dias,
            "convencion": "CRÍTICO (<= 30 DÍAS)",
            "color": "amber",
            "hex_bg": "D97706",
            "hex_fg": "FFFFFF",
            "vencido": False,
            "alerta": True,
        }
    if dias <= 90:
        return {
            "vigencia_dias": dias,
            "convencion": "PRECAUCIÓN (<= 90 DÍAS)",
            "color": "yellow",
            "hex_bg": "EAB308",
            "hex_fg": "0F172A",
            "vencido": False,
            "alerta": False,
        }
    return {
        "vigencia_dias": dias,
        "convencion": "VIGENTE",
        "color": "emerald",
        "hex_bg": "059669",
        "hex_fg": "FFFFFF",
        "vencido": False,
        "alerta": False,
    }


class InventarioItem(Base):
    __tablename__ = "inventario_items"

    id = Column(Integer, primary_key=True, index=True)
    warehouse = Column(Integer, nullable=False, default=1, index=True)

    # Identificación Rápida
    codigo_barras = Column(String(100), nullable=False, index=True)
    parte_numero = Column(String(120), nullable=False, index=True)

    # Detalle del Producto
    descripcion = Column(String(255), nullable=False)
    medida = Column(String(40), nullable=False, default="UND")
    marca = Column(String(120), nullable=True, default="")
    normativa = Column(String(60), nullable=True, default="MIL-SPEC / HAZMAT")

    # Control de Calidad y Trazabilidad
    lote = Column(String(100), nullable=True, default="", index=True)
    dom = Column(Date, nullable=True)  # Date of Manufacture (Fecha de Fabricación)
    doe = Column(Date, nullable=True, index=True)  # Date of Expiration (Fecha de Expiración)

    # Estado y Localización
    existencia = Column(Integer, nullable=False, default=0)
    ubicacion = Column(String(120), nullable=True, default="")

    actualizado_por_id = Column(Integer, ForeignKey("asesores.id"), nullable=True)
    actualizado_por_nombre = Column(String(120), nullable=True, default="")
    creado_en = Column(DateTime, default=datetime.utcnow)
    actualizado_en = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    actualizado_por = relationship("Asesor")
    movimientos = relationship(
        "MovimientoInventario",
        back_populates="item",
        cascade="all, delete-orphan",
        order_by="MovimientoInventario.id.desc()",
    )

    @property
    def convencion_info(self) -> dict:
        return calcular_convencion_vigencia(self.doe)

    @property
    def vigencia_dias(self) -> int | None:
        return self.convencion_info["vigencia_dias"]

    @property
    def convencion(self) -> str:
        return self.convencion_info["convencion"]

    @property
    def esta_vencido(self) -> bool:
        return self.convencion_info["vencido"]


class MovimientoInventario(Base):
    __tablename__ = "inventario_movimientos"

    id = Column(Integer, primary_key=True, index=True)
    item_id = Column(Integer, ForeignKey("inventario_items.id"), nullable=True, index=True)
    warehouse = Column(Integer, nullable=False, default=1, index=True)
    codigo_barras = Column(String(100), nullable=False, default="")
    parte_numero = Column(String(120), nullable=False, default="")
    lote = Column(String(100), nullable=True, default="")
    tipo_movimiento = Column(String(40), nullable=False, default="ENTRADA")
    cantidad_anterior = Column(Integer, nullable=False, default=0)
    cantidad_movimiento = Column(Integer, nullable=False, default=0)
    cantidad_nueva = Column(Integer, nullable=False, default=0)
    po_asociada = Column(String(60), nullable=True, default="")
    observacion = Column(String(255), nullable=True, default="")
    usuario_id = Column(Integer, ForeignKey("asesores.id"), nullable=True)
    usuario_nombre = Column(String(120), nullable=False, default="Sistema")
    usuario_rol = Column(String(50), nullable=True, default="")
    fecha_hora_bogota = Column(String(40), nullable=False, default="")
    creado_en = Column(DateTime, default=datetime.utcnow)

    item = relationship("InventarioItem", back_populates="movimientos")
    usuario = relationship("Asesor")



class ClaveHorarioExtendido(Base):
    __tablename__ = "claves_horario_extendido"

    id = Column(Integer, primary_key=True, index=True)
    codigo_hash = Column(String(64), nullable=False, unique=True, index=True)
    codigo_sufijo = Column(String(8), nullable=False, default="")
    generado_por_id = Column(Integer, ForeignKey("asesores.id"), nullable=False, index=True)
    asesor_destino_id = Column(Integer, ForeignKey("asesores.id"), nullable=True, index=True)
    motivo = Column(String(255), nullable=True, default="")
    creado_en = Column(DateTime, default=datetime.utcnow)
    expira_en = Column(DateTime, nullable=False)
    revocada = Column(Boolean, nullable=False, default=False)
    usos = Column(Integer, nullable=False, default=0)
    ultimo_uso_en = Column(DateTime, nullable=True)
    ultimo_uso_por = Column(String(120), nullable=True, default="")

    generado_por = relationship("Asesor", foreign_keys=[generado_por_id])
    asesor_destino = relationship("Asesor", foreign_keys=[asesor_destino_id])

