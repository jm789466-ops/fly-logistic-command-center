import os
import sys
from datetime import date, timedelta
from sqlalchemy import or_

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from database import SessionLocal, engine, Base, BASE_DIR
import models
from main import (
    parsear, aplicar, ruta_excel, hash_clave,
    generar_excel_com_fm_08,
    generar_excel_consolidado_usuario,
    generar_excel_maestro_productos
)


def crear_o_actualizar_usuario(db, nombre: str, username: str, email: str, rol: str, password: str, jefe_id=None) -> models.Asesor:
    norm = models.normalizar(nombre)
    u = db.query(models.Asesor).filter(
        or_(models.Asesor.username == username, models.Asesor.nombre_norm == norm)
    ).first()
    if not u:
        u = models.Asesor(
            nombre=nombre,
            nombre_norm=norm,
            username=username,
            email=email,
            rol=rol,
            jefe_id=jefe_id,
            password_hash=hash_clave(password)
        )
        db.add(u)
    else:
        u.nombre = nombre
        u.nombre_norm = norm
        u.username = username
        u.email = email
        u.rol = rol
        u.jefe_id = jefe_id
        u.password_hash = hash_clave(password)
    db.commit()
    db.refresh(u)
    return u


def guardar_o_reemplazar_po(db, datos_orden: dict, asesor_id: int):
    po_codigo = datos_orden["po_numero"]
    existente = db.query(models.Pedido).filter(models.Pedido.po == po_codigo).first()
    if existente:
        db.delete(existente)
        db.commit()

    parsed = parsear(datos_orden)
    pedido = models.Pedido(po=parsed["po"], tipo=parsed["tipo"], asesor_id=asesor_id)
    aplicar(db, pedido, parsed)
    db.add(pedido)
    db.commit()
    db.refresh(pedido)

    ruta_po_xlsx, _ = ruta_excel(pedido)
    generar_excel_com_fm_08(pedido, ruta_po_xlsx)
    return pedido, ruta_po_xlsx


def ejecutar_qa_completo():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()

    try:
        admin = crear_o_actualizar_usuario(
            db, "Juan Esteban Martínez", "admin",
            "notificaciones@flylogistic.aero", "Administrador", "Colombia2024*"
        )
        leonardo = crear_o_actualizar_usuario(
            db, "Leonardo Garzón", "lgarzon",
            "leonardo.garzon@flylogistic.aero", "Jefe de area", "Colombia2024*", jefe_id=admin.id
        )
        miguel = crear_o_actualizar_usuario(
            db, "Miguel Mosquera", "mmosquera",
            "miguel.mosquera@flylogistic.aero", "Asesor", "Colombia2024*", jefe_id=leonardo.id
        )
        carlos = crear_o_actualizar_usuario(
            db, "Carlos Mendoza", "cmendoza",
            "carlos.mendoza@flylogistic.aero", "Asesor", "Colombia2024*", jefe_id=leonardo.id
        )

        # Eliminar usuarios de prueba residuales si existieran
        for u_old in db.query(models.Asesor).filter(models.Asesor.username.in_(["qa_tester", "asesor_qa"])).all():
            db.query(models.Pedido).filter(models.Pedido.asesor_id == u_old.id).update(
                {"asesor_id": leonardo.id}, synchronize_session=False
            )
            db.delete(u_old)
        db.commit()

        hoy = date.today()
        hoy_str = hoy.strftime("%Y-%m-%d")
        plazo_favor = (hoy + timedelta(days=10)).strftime("%Y-%m-%d")
        plazo_contra = (hoy - timedelta(days=4)).strftime("%Y-%m-%d")

        escenarios_etapas = [
            (leonardo, {
                "po_numero": "PO-2026-101",
                "cliente": "AeroAndina S.A.S.",
                "tipo_pedido": "Nacional",
                "fecha_inicial": hoy_str,
                "plazo_entrega": plazo_favor,
                "metodo_pago": "Transferencia",
                "iva_porcentaje": "19",
                "valor_nacionalizacion": "0.00",
                "estado_general": "En proceso",
                "observaciones": "Orden inicial registrada con ítem pendiente de despacho.",
                "items": [{
                    "fecha_compra_prov": "",
                    "proveedor": "Distribuidora Aeronáutica Bogotá",
                    "parte_numero": "PN-CAB-001",
                    "parte_numero_alterno": "",
                    "unidad_medida": "UND",
                    "descripcion": "Válvula de alivio de presión de cabina",
                    "cantidad": "2",
                    "precio_unitario": "950.00",
                    "iva_item": "19",
                    "estado_producto": "PENDIENTE"
                }]
            }),
            (miguel, {
                "po_numero": "PO-2026-102",
                "cliente": "Helicópteros del Norte S.A.",
                "tipo_pedido": "Nacional",
                "fecha_inicial": (hoy - timedelta(days=3)).strftime("%Y-%m-%d"),
                "plazo_entrega": plazo_favor,
                "metodo_pago": "Visa: $2,000.00, Transferencia: $1,808.00",
                "iva_porcentaje": "19",
                "valor_nacionalizacion": "0.00",
                "estado_general": "En proceso",
                "observaciones": "Orden en curso con compra efectuada a proveedor.",
                "items": [
                    {
                        "fecha_compra_prov": hoy_str,
                        "proveedor": "Distribuidora Aeronáutica Bogotá",
                        "parte_numero": "PN-HYD-101",
                        "parte_numero_alterno": "ALT-HYD-1",
                        "unidad_medida": "UND",
                        "descripcion": "Actuador hidráulico de tren principal",
                        "cantidad": "1",
                        "precio_unitario": "2200.00",
                        "iva_item": "19",
                        "estado_producto": "DESPACHADO"
                    },
                    {
                        "fecha_compra_prov": hoy_str,
                        "proveedor": "Componentes Aero Medellín",
                        "parte_numero": "PN-EGT-102",
                        "parte_numero_alterno": "",
                        "unidad_medida": "EA",
                        "descripcion": "Sensor de temperatura de gases de escape EGT",
                        "cantidad": "2",
                        "precio_unitario": "500.00",
                        "iva_item": "19",
                        "estado_producto": "EN FLY"
                    }
                ]
            }),
            (carlos, {
                "po_numero": "PO-2026-103",
                "cliente": "Servicios Aéreos Especiales S.A.S.",
                "tipo_pedido": "Nacional",
                "fecha_inicial": (hoy - timedelta(days=8)).strftime("%Y-%m-%d"),
                "plazo_entrega": (hoy - timedelta(days=1)).strftime("%Y-%m-%d"),
                "fecha_entrega_cliente": hoy_str,
                "numero_remision": "REM-2026-099",
                "documentacion": "Factura FAC-9901 / Packing Slip",
                "metodo_pago": "Visa: $1,500.00, American: $1,475.00",
                "iva_porcentaje": "19",
                "valor_nacionalizacion": "0.00",
                "estado_general": "Entregado",
                "observaciones": "Orden entregada conforme con remisión firmada.",
                "items": [{
                    "fecha_compra_prov": (hoy - timedelta(days=7)).strftime("%Y-%m-%d"),
                    "proveedor": "Avionics Colombia S.A.S.",
                    "parte_numero": "PN-FCU-301",
                    "parte_numero_alterno": "ALT-301",
                    "unidad_medida": "UND",
                    "descripcion": "Unidad de control de combustible FCU",
                    "cantidad": "1",
                    "precio_unitario": "2500.00",
                    "iva_item": "19",
                    "estado_producto": "ENTREGADO"
                }]
            }),
            (leonardo, {
                "po_numero": "PO-2026-104",
                "cliente": "Florida Jet Support LLC",
                "tipo_pedido": "Internacional",
                "fecha_inicial": (hoy - timedelta(days=12)).strftime("%Y-%m-%d"),
                "plazo_entrega": plazo_contra,
                "metodo_pago": "PayPal: $4,000.00, Transferencia: $3,295.00",
                "iva_porcentaje": "7",
                "valor_nacionalizacion": "620.00",
                "estado_general": "En proceso",
                "llega_click_mail": "Si",
                "fecha_click_mail": (hoy - timedelta(days=5)).strftime("%Y-%m-%d"),
                "tracking": "1Z999AA10123456784",
                "estado_nacionalizacion": "Pago de aranceles",
                "fecha_nacionaliza": "",
                "observaciones": "Importación en proceso de liquidación arancelaria.",
                "items": [{
                    "fecha_compra_prov": (hoy - timedelta(days=10)).strftime("%Y-%m-%d"),
                    "proveedor": "Miami Aero Parts Inc.",
                    "parte_numero": "PN-CFM-770",
                    "parte_numero_alterno": "ALT-770",
                    "unidad_medida": "KIT",
                    "descripcion": "Juego de sellos de turbina CFM56",
                    "cantidad": "1",
                    "precio_unitario": "6238.32",
                    "iva_item": "7",
                    "estado_producto": "LLEGADO FLY"
                }]
            }),
            (miguel, {
                "po_numero": "PO-2026-105",
                "cliente": "Latam Cargo Operations",
                "tipo_pedido": "Internacional",
                "fecha_inicial": (hoy - timedelta(days=15)).strftime("%Y-%m-%d"),
                "plazo_entrega": plazo_favor,
                "fecha_entrega_cliente": hoy_str,
                "numero_remision": "REM-INT-505",
                "documentacion": "Declaración de Importación DIM-8821",
                "metodo_pago": "Transferencia",
                "iva_porcentaje": "7",
                "valor_nacionalizacion": "980.00",
                "estado_general": "Entregado",
                "llega_click_mail": "Si",
                "fecha_click_mail": (hoy - timedelta(days=9)).strftime("%Y-%m-%d"),
                "tracking": "FX-8844221100",
                "estado_nacionalizacion": "Liberado / Nacionalizado",
                "fecha_nacionaliza": (hoy - timedelta(days=2)).strftime("%Y-%m-%d"),
                "observaciones": "Mercancía nacionalizada y entregada en hangar.",
                "items": [{
                    "fecha_compra_prov": (hoy - timedelta(days=14)).strftime("%Y-%m-%d"),
                    "proveedor": "Dallas Turbine Supply",
                    "parte_numero": "PN-PMP-900",
                    "parte_numero_alterno": "",
                    "unidad_medida": "EA",
                    "descripcion": "Bomba de combustible de alta presión",
                    "cantidad": "1",
                    "precio_unitario": "9500.00",
                    "iva_item": "7",
                    "estado_producto": "ENTREGADO"
                }]
            })
        ]

        for responsable, esc in escenarios_etapas:
            guardar_o_reemplazar_po(db, esc, responsable.id)

        estados_prod_ciclo = ["PENDIENTE", "DESPACHADO", "EN FLY", "LLEGADO FLY", "ENTREGADO"]
        estados_aduana_ciclo = ["No aplica / Pendiente", "En revisión aduanera", "Pago de aranceles", "Liberado / Nacionalizado"]
        unidades_ciclo = ["UND", "EA", "GAL", "LT", "KIT", "LB", "CAJA", "SET"]
        equipo_rotativo = [leonardo, miguel, carlos]

        for cant_items in range(1, 21):
            responsable_actual = equipo_rotativo[(cant_items - 1) % len(equipo_rotativo)]
            es_intl = (cant_items % 2 == 0)
            tipo_op = "Internacional" if es_intl else "Nacional"
            imp_def = "7" if es_intl else "19"
            estado_po = "Entregado" if cant_items in (4, 8, 12, 16, 20) else "En proceso"
            po_code = f"PO-2026-{cant_items:02d}P"

            items_generados = []
            for idx in range(1, cant_items + 1):
                est_item = "ENTREGADO" if estado_po == "Entregado" else estados_prod_ciclo[(idx - 1) % len(estados_prod_ciclo)]
                items_generados.append({
                    "fecha_compra_prov": hoy_str if est_item != "PENDIENTE" else "",
                    "proveedor": f"AeroSupplier Global #{idx}",
                    "parte_numero": f"PN-{cant_items:02d}-{idx:03d}",
                    "parte_numero_alterno": f"ALT-{idx:03d}" if idx % 2 == 0 else "",
                    "unidad_medida": unidades_ciclo[(idx - 1) % len(unidades_ciclo)],
                    "descripcion": f"Componente aeronáutico certificado línea #{idx}",
                    "cantidad": str((idx % 4) + 1),
                    "precio_unitario": f"{125.50 * idx:.2f}",
                    "iva_item": imp_def,
                    "estado_producto": est_item
                })

            est_aduana = "Liberado / Nacionalizado" if (es_intl and estado_po == "Entregado") else (
                estados_aduana_ciclo[cant_items % len(estados_aduana_ciclo)] if es_intl else "No aplica / Pendiente"
            )

            datos_escala = {
                "po_numero": po_code,
                "cliente": f"Corporación Aérea Colombiana ({cant_items} Ítems)",
                "tipo_pedido": tipo_op,
                "fecha_inicial": hoy_str,
                "plazo_entrega": plazo_favor if cant_items % 3 != 0 else plazo_contra,
                "fecha_entrega_cliente": hoy_str if estado_po == "Entregado" else "",
                "numero_remision": f"REM-{cant_items:02d}" if estado_po == "Entregado" else "",
                "documentacion": f"SOPORTE-PO-{cant_items:02d}.pdf",
                "metodo_pago": "Visa: $3,500.00, Transferencia: $6,500.00" if cant_items % 2 == 0 else "Transferencia",
                "iva_porcentaje": imp_def,
                "valor_nacionalizacion": f"{150.00 * cant_items:.2f}" if es_intl else "0.00",
                "estado_general": estado_po,
                "llega_click_mail": "Si" if es_intl else "No",
                "fecha_click_mail": hoy_str if es_intl else "",
                "tracking": f"TRK-INTL-{cant_items:02d}99" if es_intl else "",
                "estado_nacionalizacion": est_aduana,
                "fecha_nacionaliza": hoy_str if (es_intl and estado_po == "Entregado") else "",
                "observaciones": f"Orden corporativa con {cant_items} ítem(s) ({tipo_op} - {estado_po}).",
                "items": items_generados
            }

            guardar_o_reemplazar_po(db, datos_escala, responsable_actual.id)

        for integrante in (admin, leonardo, miguel, carlos):
            pedidos_usr = db.query(models.Pedido).filter_by(asesor_id=integrante.id).all()
            ruta_cons = os.path.join(BASE_DIR, "Consolidados_Usuarios", f"Consolidado_POs_{integrante.username}.xlsx")
            generar_excel_consolidado_usuario(ruta_cons, pedidos_usr, integrante.nombre)

        todos_pedidos = db.query(models.Pedido).all()
        ruta_maestro = os.path.join(BASE_DIR, "Reportes_Gerenciales", "Control_Maestro_Productos_PO_FlyLogistic.xlsx")
        generar_excel_maestro_productos(ruta_maestro, todos_pedidos)

        print("OK - Base de datos y reportes actualizados correctamente.")

    except Exception as e:
        db.rollback()
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        db.close()


if __name__ == "__main__":
    ejecutar_qa_completo()