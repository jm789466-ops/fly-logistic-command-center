import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal

from sqlalchemy import (Column, Integer, String, Date, DateTime, Text,
                        Numeric, ForeignKey, Boolean)
from sqlalchemy.orm import relationship

from database import Base


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
    creado_en = Column(DateTime, default=datetime.utcnow)

    pedidos = relationship("Pedido", back_populates="asesor")
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

    metodo_pago = Column(String(150), nullable=True)
    iva_porcentaje = Column(Numeric(5, 2), nullable=False, default=19)
    valor_nacionalizacion = Column(Numeric(14, 2), nullable=False, default=0)
    shipping = Column(Numeric(14, 2), nullable=False, default=0)

    plazo_entrega = Column(Date, nullable=True)
    fecha_entrega_cliente = Column(Date, nullable=True)
    numero_remision = Column(String(50), nullable=True)

    documentacion = Column(String(500), nullable=True)
    documentacion_archivo = Column(String(500), nullable=True)
    documentacion_drive_link = Column(String(500), nullable=True)

    estado = Column(String(25), nullable=False, default="En proceso")
    observaciones = Column(Text, nullable=True)

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
    def dias_habiles_info(self) -> dict:
        etiqueta_po = self.po_codigo
        if not self.plazo_entrega:
            return {
                "dias": None,
                "orden_urgencia": 999999,
                "en_contra": False,
                "etiqueta": f"{etiqueta_po} | SIN PLAZO",
                "detalle": "Sin fecha límite",
                "color": "slate"
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
                "color": "emerald"
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
                "color": "red"
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
    estado_producto = Column(String(30), nullable=False, default="PENDIENTE")

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