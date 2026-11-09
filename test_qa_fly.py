import os
import re
import sys
import json
import inspect
import asyncio
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from sqlalchemy import or_

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from database import SessionLocal, engine, Base, BASE_DIR
import models
import main
from main import (
    parsear, aplicar, ruta_excel, hash_clave, migrar_bd,
    generar_excel_com_fm_08,
    generar_excel_consolidado_usuario,
    generar_excel_maestro_productos,
    validar_horario_asesor,
    pedido_del_usuario,
    registrar_historial_pedido,
    sincronizar_resumen_documentos_po,
    _guardar_cotizacion_en_bd,
    _construir_analitica,
    mis_pedidos,
    nuevo_form,
    editar_form,
    cotizador_nuevo,
    listar_cotizaciones,
    vista_informes,
    vista_supervision,
    gestion_usuarios,
    perfil,
    login_form,
    olvide_form,
    registro_form,
    restablecer_form,
    api_tracking,
    api_validar_po,
    api_validar_cotizacion,
    inicializar_admin_y_correo,
    vista_inventario,
    api_escanear_codigo_barras,
    api_validar_stock_bodega,
    despachar_material_a_po,
    exportar_excel_warehouse,
    exportar_excel_inventario_maestro,
    previsualizar_documento_adjunto,
    cancelar_pedido,
    reactivar_pedido_cancelado,
    actualizar_estado_rapido_supervision,
    REGISTROS_POR_PAGINA,
    MAX_ARCHIVOS_PO,
    generar_clave_horario,
    desbloquear_horario,
    revocar_clave_horario,
    estado_horario_usuario,
    desbloqueo_horario_vigente,
    ahora_bogota_naive,
    servir_logo_corporativo,
    ruta_inicio_rol,
    ROLES_SISTEMA,
)


def run_coro(coro):
    stack = [coro]
    val = None
    while stack:
        top = stack[-1]
        try:
            val = top.send(val)
            if inspect.iscoroutine(val):
                stack.append(val)
                val = None
        except StopIteration as e:
            val = e.value
            stack.pop()
    return val


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
    os.environ["OMS_DISABLE_DRIVE"] = "1"
    Base.metadata.create_all(bind=engine)
    migrar_bd()
    inicializar_admin_y_correo()
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
        almacen = crear_o_actualizar_usuario(
            db, "Andrés Almacén", "almacen",
            "almacen@flylogistic.aero", "Almacenista", "Colombia2024*", jefe_id=admin.id
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
                "observaciones": "Orden inicial registrada con ítem pendiente de despacho y cobro certificado.",
                "items": [{
                    "fecha_compra_prov": "",
                    "proveedor": "Distribuidora Aeronáutica Bogotá",
                    "parte_numero": "PN-CAB-001",
                    "parte_numero_alterno": "",
                    "unidad_medida": "UND",
                    "descripcion": "Válvula de alivio de presión de cabina",
                    "cantidad": "2",
                    "precio_unitario": "950.00",
                    "cobro_hazmat": "0.00",
                    "cobro_certificado": "120.00",
                    "sin_iva": "0",
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
                "observaciones": "Orden en curso con compra efectuada a proveedor e ítem exento de IVA.",
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
                        "cobro_hazmat": "150.00",
                        "cobro_certificado": "80.00",
                        "sin_iva": "0",
                        "iva_item": "19",
                        "estado_producto": "DESPACHADO"
                    },
                    {
                        "fecha_compra_prov": hoy_str,
                        "proveedor": "Componentes Aero Medellín",
                        "parte_numero": "PN-EGT-102",
                        "parte_numero_alterno": "",
                        "unidad_medida": "EA",
                        "descripcion": "Sensor de temperatura de gases de escape EGT (Exento)",
                        "cantidad": "2",
                        "precio_unitario": "500.00",
                        "cobro_hazmat": "0.00",
                        "cobro_certificado": "0.00",
                        "sin_iva": "1",
                        "iva_item": "0",
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
                    "cobro_hazmat": "0.00",
                    "cobro_certificado": "95.00",
                    "sin_iva": "0",
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
                "observaciones": "Importación en proceso de liquidación arancelaria con cargo Hazmat.",
                "items": [{
                    "fecha_compra_prov": (hoy - timedelta(days=10)).strftime("%Y-%m-%d"),
                    "proveedor": "Miami Aero Parts Inc.",
                    "parte_numero": "PN-CFM-770",
                    "parte_numero_alterno": "ALT-770",
                    "unidad_medida": "KIT",
                    "descripcion": "Juego de sellos de turbina CFM56",
                    "cantidad": "1",
                    "precio_unitario": "6238.32",
                    "cobro_hazmat": "210.00",
                    "cobro_certificado": "125.00",
                    "sin_iva": "0",
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
                    "cobro_hazmat": "320.00",
                    "cobro_certificado": "150.00",
                    "sin_iva": "0",
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
                es_exento_prueba = (idx % 5 == 0)
                items_generados.append({
                    "fecha_compra_prov": hoy_str if est_item != "PENDIENTE" else "",
                    "proveedor": f"AeroSupplier Global #{idx}",
                    "parte_numero": f"PN-{cant_items:02d}-{idx:03d}",
                    "parte_numero_alterno": f"ALT-{idx:03d}" if idx % 2 == 0 else "",
                    "unidad_medida": unidades_ciclo[(idx - 1) % len(unidades_ciclo)],
                    "descripcion": f"Componente aeronáutico certificado línea #{idx}",
                    "cantidad": str((idx % 4) + 1),
                    "precio_unitario": f"{125.50 * idx:.2f}",
                    "cobro_hazmat": f"{35.00 if idx % 3 == 0 else 0.00:.2f}",
                    "cobro_certificado": f"{25.00 if idx % 2 == 0 else 0.00:.2f}",
                    "sin_iva": "1" if es_exento_prueba else "0",
                    "iva_item": "0" if es_exento_prueba else imp_def,
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
            nombre_limpio = integrante.nombre.replace(" ", "_")
            ruta_cons = os.path.join(BASE_DIR, "Consolidados_Usuarios", f"Consolidado_POs_{nombre_limpio}.xlsx")
            generar_excel_consolidado_usuario(ruta_cons, pedidos_usr, integrante.nombre)

        todos_pedidos = db.query(models.Pedido).all()
        ruta_maestro = os.path.join(BASE_DIR, "Reportes_Gerenciales", "Control_Maestro_Productos_PO_FlyLogistic.xlsx")
        generar_excel_maestro_productos(ruta_maestro, todos_pedidos)

        verificar_reglas_negocio_y_arquitectura(db, admin, leonardo, miguel, carlos, almacen)

        print("OK - Base de datos, cobros Hazmat/Certificados, IVA Exento y los 4 reportes actualizados correctamente.")

    except Exception as e:
        db.rollback()
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        db.close()


def verificar_reglas_negocio_y_arquitectura(db, admin, leonardo, miguel, carlos, almacen):
    # 1. Paginación estricta y límite de archivos
    assert REGISTROS_POR_PAGINA == 10, f"Paginación debe ser máximo 10, actual: {REGISTROS_POR_PAGINA}"
    assert MAX_ARCHIVOS_PO == 10, f"Máximo de archivos por PO debe ser 10, actual: {MAX_ARCHIVOS_PO}"

    # 2. Restricción Horaria RBAC (America/Bogota): Asesor 17:30, Gerencia 21:00
    gerente_tmp = models.Asesor(nombre="Gerente Test", username="ger_tmp", rol="Gerencia", password_hash="x")
    ok_as_1729, _ = validar_horario_asesor(miguel, hora_prueba=time(17, 29))
    ok_as_1730, msg_as = validar_horario_asesor(miguel, hora_prueba=time(17, 30))
    ok_ge_2059, _ = validar_horario_asesor(gerente_tmp, hora_prueba=time(20, 59))
    ok_ge_2100, msg_ge = validar_horario_asesor(gerente_tmp, hora_prueba=time(21, 0))
    ok_ad_2330, _ = validar_horario_asesor(admin, hora_prueba=time(23, 30))
    assert ok_as_1729 is True, "Asesor debe poder operar a las 17:29"
    assert ok_as_1730 is False and "17:30" in msg_as, "Asesor debe bloquearse a las 17:30"
    assert ok_ge_2059 is True, "Gerencia debe tener extensión activa a las 20:59"
    assert ok_ge_2100 is False and "21:00" in msg_ge, "Gerencia debe bloquearse a las 21:00"
    assert ok_ad_2330 is True, "Administrador no debe tener bloqueo horario"

    # 3. Motor de Tracking: FedEx, Coordinadora, DHL y Clickmail
    for carrier_esperado, guia_ejemplo, cm_flag in [
        ("FedEx", "FX-778899001122", False),
        ("Coordinadora", "73000123456", False),
        ("DHL", "JD0146000012345678", False),
        ("Clickmail", "CM-USA-9001", True),
    ]:
        info_trk = models.detectar_transportadora_tracking(
            tracking=guia_ejemplo,
            llega_click_mail="Si" if cm_flag else "No",
            transportadora_manual=carrier_esperado,
        )
        assert info_trk["carrier"] == carrier_esperado, f"Fallo en motor tracking para {carrier_esperado}"
        assert info_trk["url"].startswith("https://"), f"URL inválida para {carrier_esperado}"

    # 4. Lógica Financiera: COP vs USD, Splits, Toggle sin_iva, cobro_hazmat y cobro_certificado
    p_split = db.query(models.Pedido).filter(models.Pedido.po == "PO-2026-102").first()
    assert p_split is not None, "PO-2026-102 debe existir"
    assert p_split.moneda == "COP", "PO Nacional debe estar en COP"
    assert len(p_split.desglose_pagos) == 2, "Debe tener 2 splits de pago registrados"
    assert p_split.items[0].cobro_hazmat_dec == Decimal("150.00")
    assert p_split.items[0].cobro_certificado_dec == Decimal("80.00")
    assert p_split.items[1].es_exento_iva is True and p_split.items[1].impuesto_item == Decimal("0.00")

    p_usd = db.query(models.Pedido).filter(models.Pedido.po == "PO-2026-104").first()
    assert p_usd is not None and p_usd.moneda == "USD", "PO Internacional debe estar en USD"

    analitica = _construir_analitica(db, admin)
    assert analitica["valor_total_cop"] > 0 and analitica["valor_total_usd"] > 0, "Separación de divisas COP y USD activa"

    # 5. Concurrencia FOR UPDATE y Retención de Historial al pasar a CANCELADO
    po_bloqueada = pedido_del_usuario(db, p_split.id, admin, bloquear=True)
    assert po_bloqueada is not None, "Consulta con FOR UPDATE debe retornar el pedido bloqueado"
    estado_previo = po_bloqueada.estado
    po_bloqueada.estado = "Cancelado"
    po_bloqueada.motivo_cancelacion = "Prueba QA de retención histórica en estado CANCELADO"
    registrar_historial_pedido(
        db, po_bloqueada, estado_previo, "Cancelado",
        accion="CANCELACION_QA", motivo=po_bloqueada.motivo_cancelacion, user=admin
    )
    db.commit()
    db.refresh(po_bloqueada)
    assert po_bloqueada.estado == "Cancelado"
    assert po_bloqueada.dias_habiles_info["estado_plazo"] == "Cancelado"
    assert len(po_bloqueada.historial) >= 1, "El historial de cancelación debe retenerse en base de datos"
    # Restaurar a En proceso registrando también la transición en el historial
    po_bloqueada.estado = "En proceso"
    registrar_historial_pedido(
        db, po_bloqueada, "Cancelado", "En proceso",
        accion="REACTIVACION_QA", motivo="Restauración posterior a prueba QA", user=admin
    )
    db.commit()

    # 6. Gestión de Archivos: Lote de hasta 10 documentos y borrado individual sin afectar el resto
    db.query(models.PedidoDocumento).filter(models.PedidoDocumento.pedido_id == p_split.id).delete()
    db.commit()
    for i in range(1, 11):
        db.add(models.PedidoDocumento(
            pedido_id=p_split.id,
            nombre_archivo=f"soporte_qa_{i:02d}.pdf",
            drive_file_id=f"local_id_{i}",
            drive_link=f"/pedidos/{p_split.id}/documento?nombre=soporte_qa_{i:02d}.pdf",
        ))
    db.commit()
    db.refresh(p_split)
    sincronizar_resumen_documentos_po(p_split)
    db.commit()
    assert len(p_split.lista_documentos) == 10, "Debe soportar hasta 10 documentos por PO"
    doc_a_borrar = p_split.lista_documentos[0]
    db.delete(doc_a_borrar)
    db.commit()
    db.refresh(p_split)
    sincronizar_resumen_documentos_po(p_split)
    db.commit()
    assert len(p_split.lista_documentos) == 9, "El borrado individual debe conservar los 9 documentos restantes intactos"

    # 7. Cotizaciones en Dashboard Centralizado (COP y USD)
    cot_cop = _guardar_cotizacion_en_bd(db, leonardo, {
        "quote_number": "Q-2026-001",
        "quote_date": str(date.today()),
        "expiration_date": str(date.today() + timedelta(days=7)),
        "prepared_by": leonardo.nombre,
        "to_company": "Aviación Ejecutiva Colombia S.A.S.",
        "ref_no": "REF-COP-01",
        "moneda": "COP",
        "transport": "150000.00",
        "items": [{
            "description": "Filtro hidráulico de alta presión",
            "unit": "UND",
            "qty": "2",
            "unit_price": "850000.00",
            "cobro_hazmat": "50000.00",
            "cobro_certificado": "30000.00",
            "sin_iva": "0",
            "remarks": "Entrega inmediata"
        }]
    })
    db.commit()
    assert cot_cop.total > Decimal("0.00") and cot_cop.moneda == "COP"

    # 8. Verificación de los 13 HTMLs (Cero emojis, Copyright obligatorio y botón antiguo eliminado)
    html_files = [
        "base.html", "cotizaciones.html", "cotizador_form.html", "Informes.html",
        "login.html", "olvide.html", "pedido_form.html", "pedidos.html",
        "perfil.html", "registro.html", "restablecer.html", "Supervision.html", "Usuarios.html"
    ]
    emoji_pattern = re.compile(
        "["
        "\U0001F300-\U0001F9FF"
        "\U00002600-\U000027BF"
        "\U0001FA00-\U0001FAFF"
        "]+",
        flags=re.UNICODE,
    )
    for hf in html_files:
        ruta_hf = os.path.join(BASE_DIR, "templates", hf)
        assert os.path.exists(ruta_hf), f"Falta plantilla HTML: {hf}"
        with open(ruta_hf, "r", encoding="utf-8") as f_in:
            contenido = f_in.read()
        assert "Copyright © 2026 Fly Logistic." in contenido, f"Falta footer obligatorio en {hf}"
        emojis_hallados = emoji_pattern.findall(contenido)
        assert not emojis_hallados, f"Emojis prohibidos detectados en {hf}: {emojis_hallados}"
        if hf in ("pedidos.html", "cotizador_form.html"):
            assert "Historial de Cotizaciones" not in contenido and "Historial de cotizaciones" not in contenido, (
                f"El botón antiguo de Historial de Cotizaciones debe estar eliminado en {hf}"
            )

    # 9. Renderizado real de las vistas Jinja2 y endpoints API
    class DummyRequest:
        def __init__(self, payload=None):
            self.session = {}
            self.headers = {}
            self.method = "GET"
            self.url = "http://127.0.0.1:8000/pedidos"
            self.scope = {"type": "http", "session": self.session}
            self._payload = payload or {}

        async def json(self):
            return self._payload

        async def form(self):
            return self._payload

    req_qa = DummyRequest()
    for resp_view in [
        mis_pedidos(req_qa, user=admin, db=db),
        mis_pedidos(req_qa, vista="cotizacion", user=admin, db=db),
        mis_pedidos(req_qa, kpi="vencidas", user=admin, db=db),
        nuevo_form(req_qa, cot_id=cot_cop.id, user=admin, db=db),
        editar_form(p_split.id, req_qa, user=admin, db=db),
        cotizador_nuevo(req_qa, user=admin, db=db),
        listar_cotizaciones(req_qa, user=admin, db=db),
        vista_informes(req_qa, user=admin, db=db),
        vista_supervision(req_qa, user=admin, db=db),
        gestion_usuarios(req_qa, user=admin, db=db),
        perfil(req_qa, user=admin, db=db),
        login_form(req_qa),
        olvide_form(req_qa),
        registro_form(req_qa, db=db),
        restablecer_form("token_invalido_qa", req_qa, db=db),
    ]:
        assert resp_view.status_code in (200, 400), f"Error HTTP al renderizar vista: {resp_view.status_code}"
        body_html = resp_view.body.decode("utf-8")
        assert "Copyright © 2026 Fly Logistic." in body_html, "El HTML renderizado debe incluir el footer corporativo"

    assert api_tracking(guia="FX-999", transportadora="FedEx", clickmail="0", user=admin).status_code == 200
    assert api_validar_po(po="PO-2026-102", user=admin, db=db).status_code == 200
    assert api_validar_cotizacion(q="Q-2026-001", db=db, user=admin).status_code == 200

    # 10. Pruebas de Previsualización Integrada de Documentos, Cancelación SQLite DateTime y Supervisión Avanzada
    resp_prev_doc = previsualizar_documento_adjunto(p_split.id, nombre="soporte_qa_02.pdf", user=admin, db=db)
    assert resp_prev_doc.status_code == 200, "Previsualización inline de documento debe responder HTTP 200"

    # Prueba Crítica QA (a): Cancelación de pedido en SQLite persistiendo objetos datetime puros sin TypeError
    req_cancel_endpoint = DummyRequest({"motivo_cancelacion": "Prueba QA endpoint cancelar_pedido con datetime puro SQLite"})
    resp_cancel_ep = run_coro(cancelar_pedido(p_split.id, req_cancel_endpoint, user=admin, db=db))
    assert resp_cancel_ep.status_code == 303, "Endpoint cancelar_pedido debe redirigir con 303 sin lanzar TypeError"
    db.refresh(p_split)
    assert p_split.estado == "Cancelado", "El pedido debe quedar en estado Cancelado"
    assert isinstance(p_split.fecha_cancelacion, datetime), (
        f"fecha_cancelacion debe ser objeto datetime puro en SQLite, recibido: {type(p_split.fecha_cancelacion)}"
    )
    assert isinstance(p_split.actualizado_en, datetime), (
        f"actualizado_en debe ser objeto datetime puro en SQLite, recibido: {type(p_split.actualizado_en)}"
    )

    # Reactivación por Gerencia / Administrador y actualización rápida desde Supervisión
    req_sup = DummyRequest({"pedido_id": str(p_split.id), "nuevo_estado": "En proceso", "nota_supervision": "Validación QA"})
    resp_sup_est = run_coro(actualizar_estado_rapido_supervision(req_sup, user=admin, db=db))
    assert resp_sup_est.status_code == 303
    db.refresh(p_split)
    assert isinstance(p_split.actualizado_en, datetime)

    # 11. Pruebas Integrales del Módulo de Inventario Multibodega (Warehouses 1, 2 y 3), Pistola Tera HW0009 y Reportes Excel
    for wh_ctx in ("maestra", "1", "2", "3"):
        resp_inv = vista_inventario(req_qa, warehouse=wh_ctx, q="", convencion="", page=1, user=admin, db=db)
        assert resp_inv.status_code == 200, f"Error al renderizar vista_inventario para warehouse={wh_ctx}"
        inv_html = resp_inv.body.decode("utf-8")
        assert "Copyright © 2026 Fly Logistic." in inv_html
        assert "Tera HW0009" in inv_html, "La interfaz de inventario debe integrar la pistola Tera HW0009"
        assert "iniciarEscuchaSilenciosaTeraHW0009" in inv_html, "Debe incluir el motor JS de escucha silenciosa Tera HW0009"
        assert "actualizarTablaInventarioSinRecargar" in inv_html, "Debe actualizar el DOM sin recargar la página"
        assert "window.location.reload()" not in inv_html, "El escaneo con Tera HW0009 no debe recargar la página"
        assert not emoji_pattern.findall(inv_html), f"Cero emojis requeridos en vista de inventario ({wh_ctx})"

    # Verificar lotes semilla en las 3 bodegas y semáforo DOE
    lote_wh1 = db.query(models.InventarioItem).filter(
        models.InventarioItem.codigo_barras == "770999000101",
        models.InventarioItem.warehouse == 1,
    ).first()
    assert lote_wh1 is not None, "Debe existir lote semilla en Warehouse 1"
    assert lote_wh1.convencion_info["convencion"] == "VIGENTE"

    lote_vencido_wh2 = db.query(models.InventarioItem).filter(
        models.InventarioItem.codigo_barras == "770999000201",
        models.InventarioItem.warehouse == 2,
    ).first()
    assert lote_vencido_wh2 is not None, "Debe existir lote vencido de prueba en Warehouse 2"
    assert lote_vencido_wh2.esta_vencido is True
    assert lote_vencido_wh2.convencion_info["convencion"] == "VENCIDOS"

    # Prueba Crítica QA (b): Simulación de inyección de datos de la pistola Tera HW0009: Entrada (+1) y Salida (-1)
    stock_antes = lote_wh1.existencia
    req_scan_in = DummyRequest({
        "codigo_barras": "770999000101",
        "warehouse": "1",
        "modo": "ENTRADA",
        "dispositivo": "Tera HW0009",
    })
    resp_scan_in = run_coro(api_escanear_codigo_barras(req_scan_in, user=admin, db=db))
    assert resp_scan_in.status_code == 200
    payload_scan_in = json.loads(resp_scan_in.body.decode("utf-8"))
    assert payload_scan_in["ok"] is True
    assert payload_scan_in["dispositivo"] == "Tera HW0009"
    assert payload_scan_in["item"]["existencia"] == stock_antes + 1
    assert "movimiento" in payload_scan_in and payload_scan_in["movimiento"]["cantidad_movimiento"] == 1
    assert "resumen_warehouses" in payload_scan_in and "1" in payload_scan_in["resumen_warehouses"]
    db.refresh(lote_wh1)
    assert lote_wh1.existencia == stock_antes + 1, "Escaneo ENTRADA Tera HW0009 debe sumar +1 en milisegundos"
    assert isinstance(lote_wh1.actualizado_en, datetime)

    req_scan_out = DummyRequest({
        "codigo_barras": "770999000101",
        "warehouse": "1",
        "modo": "SALIDA",
        "dispositivo": "Tera HW0009",
    })
    resp_scan_out = run_coro(api_escanear_codigo_barras(req_scan_out, user=admin, db=db))
    assert resp_scan_out.status_code == 200
    payload_scan_out = json.loads(resp_scan_out.body.decode("utf-8"))
    assert payload_scan_out["ok"] is True
    assert payload_scan_out["item"]["existencia"] == stock_antes
    assert payload_scan_out["movimiento"]["cantidad_movimiento"] == -1
    db.refresh(lote_wh1)
    assert lote_wh1.existencia == stock_antes, "Escaneo SALIDA Tera HW0009 debe restar -1 en milisegundos"

    # Prueba Alerta Restrictiva Sonora/Visual: Escanear ítem de OTRO Warehouse (Lote de WH 3 escaneado en WH 1)
    req_scan_otro_wh = DummyRequest({"codigo_barras": "770999000301", "warehouse": "1", "modo": "SALIDA"})
    resp_scan_otro_wh = run_coro(api_escanear_codigo_barras(req_scan_otro_wh, user=admin, db=db))
    assert resp_scan_otro_wh.status_code == 409, "Debe bloquear y disparar alerta restrictiva si pertenece a otro Warehouse"

    # Prueba Alerta Restrictiva Sonora/Visual: Escanear ítem con DOE VENCIDO en Warehouse 2
    req_scan_vencido = DummyRequest({"codigo_barras": "770999000201", "warehouse": "2", "modo": "SALIDA"})
    resp_scan_vencido = run_coro(api_escanear_codigo_barras(req_scan_vencido, user=admin, db=db))
    assert resp_scan_vencido.status_code in (400, 422), "Debe bloquear y disparar alerta restrictiva si el DOE está VENCIDO"

    # Prueba Validación Aislada de Stock por Bodega y Despacho a PO
    resp_val_stock = api_validar_stock_bodega(
        warehouse=1, parte_numero=lote_wh1.parte_numero, cantidad=1, user=admin, db=db
    )
    assert resp_val_stock.status_code == 200

    # Prueba Exportación Dinámica de Reportes Excel (Warehouse 1, 2, 3 y Maestro Consolidado)
    for wh_num in (1, 2, 3):
        resp_xlsx_wh = exportar_excel_warehouse(wh_num, user=admin, db=db)
        assert resp_xlsx_wh.status_code == 200
        assert os.path.exists(resp_xlsx_wh.path), f"Debe generarse archivo físico Inventario_Warehouse_{wh_num}"

    resp_xlsx_maestro = exportar_excel_inventario_maestro(user=admin, db=db)
    assert resp_xlsx_maestro.status_code == 200
    assert os.path.exists(resp_xlsx_maestro.path), "Debe generarse el Reporte Maestro Consolidado de las 3 bodegas"

    # 11b. Contraseña maestra de horario extendido: generada solo por Administrador, Gerencia o Jefe de área
    total_claves_antes = db.query(models.ClaveHorarioExtendido).count()
    req_gen_asesor = DummyRequest({"vigencia_horas": "2"})
    run_coro(generar_clave_horario(req_gen_asesor, user=miguel, db=db))
    assert db.query(models.ClaveHorarioExtendido).count() == total_claves_antes, "Un asesor no puede generar contraseñas maestras"
    assert req_gen_asesor.session["flash"]["ok"] is False

    req_gen = DummyRequest({"asesor_destino_id": str(miguel.id), "vigencia_horas": "2", "motivo": "Prueba QA de cierre urgente"})
    resp_gen = run_coro(generar_clave_horario(req_gen, user=leonardo, db=db))
    assert resp_gen.status_code == 303
    codigo_maestro = req_gen.session["clave_horario_nueva"]["codigo"]
    assert re.fullmatch(r"FLY-[A-Z2-9]{4}-[A-Z2-9]{4}", codigo_maestro), "Formato de contraseña maestra inválido"
    clave_bd = db.query(models.ClaveHorarioExtendido).order_by(models.ClaveHorarioExtendido.id.desc()).first()
    assert clave_bd.codigo_hash != codigo_maestro and codigo_maestro not in clave_bd.codigo_hash, "El código debe almacenarse cifrado"
    assert clave_bd.asesor_destino_id == miguel.id and clave_bd.expira_en > ahora_bogota_naive()

    # Un asesor distinto al destinatario no puede usarla
    req_otro = DummyRequest({"clave_horario": codigo_maestro})
    run_coro(desbloquear_horario(req_otro, user=carlos, db=db))
    db.refresh(carlos)
    assert carlos.desbloqueo_horario_hasta is None and req_otro.session["flash"]["ok"] is False

    # Código incorrecto
    req_mal = DummyRequest({"clave_horario": "FLY-ZZZZ-ZZZZ"})
    run_coro(desbloquear_horario(req_mal, user=miguel, db=db))
    db.refresh(miguel)
    assert miguel.desbloqueo_horario_hasta is None and req_mal.session["flash"]["ok"] is False

    # Código correcto del asesor destinatario
    req_ok = DummyRequest({"clave_horario": codigo_maestro.lower(), "next": "/pedidos"})
    resp_ok = run_coro(desbloquear_horario(req_ok, user=miguel, db=db))
    assert resp_ok.status_code == 303 and req_ok.session["flash"]["ok"] is True
    db.refresh(miguel)
    db.refresh(clave_bd)
    assert miguel.desbloqueo_horario_hasta == clave_bd.expira_en and miguel.desbloqueo_clave_id == clave_bd.id
    assert clave_bd.usos == 1 and clave_bd.ultimo_uso_por == miguel.nombre
    assert desbloqueo_horario_vigente(miguel) is True

    # El desbloqueo permite operar después de las 17:30 solo mientras esté vigente
    miguel.desbloqueo_horario_hasta = datetime.combine(ahora_bogota_naive().date(), time(23, 59))
    db.commit()
    ok_post_cierre, _ = validar_horario_asesor(miguel, hora_prueba=time(18, 0))
    assert ok_post_cierre is True, "Con contraseña maestra vigente el asesor puede operar a las 18:00"
    miguel.desbloqueo_horario_hasta = None
    db.commit()
    ok_sin_clave, msg_sin_clave = validar_horario_asesor(miguel, hora_prueba=time(18, 0))
    assert ok_sin_clave is False and "contraseña maestra" in msg_sin_clave
    assert estado_horario_usuario(miguel)["aplica"] is True and estado_horario_usuario(admin)["aplica"] is False

    # Revocación inmediata
    miguel.desbloqueo_horario_hasta = clave_bd.expira_en
    miguel.desbloqueo_clave_id = clave_bd.id
    db.commit()
    req_rev = DummyRequest()
    resp_rev = revocar_clave_horario(clave_bd.id, req_rev, user=leonardo, db=db)
    assert resp_rev.status_code == 303
    db.refresh(miguel)
    db.refresh(clave_bd)
    assert clave_bd.revocada is True and miguel.desbloqueo_horario_hasta is None
    req_reuso = DummyRequest({"clave_horario": codigo_maestro})
    run_coro(desbloquear_horario(req_reuso, user=miguel, db=db))
    db.refresh(miguel)
    assert miguel.desbloqueo_horario_hasta is None, "Una contraseña revocada no debe reactivarse"

    # Protección contra intentos repetidos
    req_gen2 = DummyRequest({"vigencia_horas": "1"})
    run_coro(generar_clave_horario(req_gen2, user=admin, db=db))
    codigo_2 = req_gen2.session["clave_horario_nueva"]["codigo"]
    sesion_fuerza_bruta = DummyRequest({"clave_horario": "FLY-AAAA-AAAA"})
    for _ in range(5):
        run_coro(desbloquear_horario(sesion_fuerza_bruta, user=carlos, db=db))
    sesion_fuerza_bruta._payload = {"clave_horario": codigo_2}
    run_coro(desbloquear_horario(sesion_fuerza_bruta, user=carlos, db=db))
    db.refresh(carlos)
    assert carlos.desbloqueo_horario_hasta is None, "Tras 5 intentos fallidos el ingreso debe quedar pausado"
    for c_qa in db.query(models.ClaveHorarioExtendido).filter(models.ClaveHorarioExtendido.revocada.is_(False)).all():
        c_qa.revocada = True
    db.commit()

    # La interfaz del perfil y el aviso para asesores se renderizan correctamente
    cuerpo_perfil_admin = perfil(req_qa, user=admin, db=db).body.decode("utf-8")
    assert "Contraseñas maestras para trabajar fuera de horario" in cuerpo_perfil_admin
    cuerpo_perfil_asesor = perfil(req_qa, user=miguel, db=db).body.decode("utf-8")
    assert "Contraseñas maestras para trabajar fuera de horario" not in cuerpo_perfil_asesor
    assert "/horario/desbloquear" in cuerpo_perfil_asesor, "El asesor debe contar con el formulario de contraseña maestra"
    html_pedidos_qa = open(os.path.join(BASE_DIR, "templates", "pedidos.html"), "r", encoding="utf-8").read()
    for texto_retirado in (
        "Mostrar Ambas Tablas Separadas",
        "Inventario (3 Bodegas)",
        "Auto / PC / Tablet",
        "Abrir Ficha Documental",
        "Abrir vista dedicada de Cotizaciones",
    ):
        assert texto_retirado not in html_pedidos_qa, f"Texto retirado todavía presente: {texto_retirado}"
    assert "previewData.kind === 'PO' && previewData.documentos && previewData.documentos.length > 0" in html_pedidos_qa, (
        "La sección de archivos adjuntos del modal solo debe mostrarse en órdenes PO con documentos reales"
    )

    # 12. Verificación de líneas mínimas y cero emojis en los 13 HTMLs y 6 archivos Python
    min_lines_html = {
        "base.html": 89,
        "cotizaciones.html": 116,
        "cotizador_form.html": 400,
        "Informes.html": 217,
        "login.html": 69,
        "olvide.html": 48,
        "pedido_form.html": 690,
        "pedidos.html": 578,
        "perfil.html": 93,
        "registro.html": 95,
        "restablecer.html": 65,
        "Supervision.html": 197,
        "Usuarios.html": 223,
    }
    for hf_name, min_hl in min_lines_html.items():
        ruta_hf = os.path.join(BASE_DIR, "templates", hf_name)
        with open(ruta_hf, "r", encoding="utf-8") as f_h:
            lineas_h = f_h.readlines()
        assert len(lineas_h) >= min_hl, f"{hf_name} tiene menos líneas ({len(lineas_h)}) que el mínimo original ({min_hl})"

    # 13. Verificación de líneas mínimas y cero emojis en los 6 archivos Python del backend
    min_lines_backend = {
        "main.py": 2983,
        "models.py": 272,
        "database.py": 18,
        "excel_com_fm_08.py": 351,
        "restablecer_clave.py": 87,
        "test_qa_fly.py": 359,
    }
    for py_name, min_l in min_lines_backend.items():
        ruta_py = os.path.join(BASE_DIR, py_name)
        with open(ruta_py, "r", encoding="utf-8") as f_py:
            lineas_py = f_py.readlines()
        assert len(lineas_py) >= min_l, f"{py_name} tiene menos líneas ({len(lineas_py)}) que el mínimo original ({min_l})"

    # 14. Pruebas Exhaustivas: Rol Almacenista (Aislamiento Estricto a Inventario), Logo Corporativo y Estados de Variables
    assert "Almacenista" in ROLES_SISTEMA, "El rol Almacenista debe estar registrado en ROLES_SISTEMA"
    assert "Almacenista" in models.ROLES_SISTEMA, "El rol Almacenista debe estar registrado en models.ROLES_SISTEMA"
    assert almacen.rol == "Almacenista", "El usuario almacen debe tener el rol Almacenista"
    assert ruta_inicio_rol(almacen) == "/inventario?warehouse=1", "La ruta de inicio para Almacenista debe ser /inventario?warehouse=1"

    # Redirección 303 obligatoria hacia /inventario?warehouse=1 en todos los módulos comerciales y directivos
    resp_alm_ped = mis_pedidos(req_qa, user=almacen, db=db)
    assert resp_alm_ped.status_code == 303 and resp_alm_ped.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /pedidos"
    )
    resp_alm_nuev = nuevo_form(req_qa, user=almacen, db=db)
    assert resp_alm_nuev.status_code == 303 and resp_alm_nuev.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /pedidos/nuevo"
    )
    resp_alm_cot = listar_cotizaciones(req_qa, user=almacen, db=db)
    assert resp_alm_cot.status_code == 303 and resp_alm_cot.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /cotizaciones"
    )
    resp_alm_cot_n = cotizador_nuevo(req_qa, user=almacen, db=db)
    assert resp_alm_cot_n.status_code == 303 and resp_alm_cot_n.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /cotizador/nuevo"
    )
    resp_alm_inf = vista_informes(req_qa, user=almacen, db=db)
    assert resp_alm_inf.status_code == 303 and resp_alm_inf.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /informes"
    )
    resp_alm_sup = vista_supervision(req_qa, user=almacen, db=db)
    assert resp_alm_sup.status_code == 303 and resp_alm_sup.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /supervision"
    )
    resp_alm_usr = gestion_usuarios(req_qa, user=almacen, db=db)
    assert resp_alm_usr.status_code == 303 and resp_alm_usr.headers.get("location") == "/inventario?warehouse=1", (
        "Almacenista debe ser redirigido con 303 al intentar acceder a /gestion_usuarios"
    )

    # Almacenista cuenta con acceso operativo completo al módulo de Inventario y pistola Tera HW0009
    resp_alm_inv = vista_inventario(req_qa, warehouse="1", user=almacen, db=db)
    assert resp_alm_inv.status_code == 200, "Almacenista debe poder visualizar el Centro de Inventario (HTTP 200)"
    html_alm_inv = resp_alm_inv.body.decode("utf-8")
    assert "Copyright © 2026 Fly Logistic." in html_alm_inv
    assert "Tera HW0009" in html_alm_inv

    # Simulación de lectura con pistola Tera HW0009 por parte de Almacenista
    req_scan_alm = DummyRequest({
        "codigo_barras": "770999000101",
        "warehouse": "1",
        "modo": "ENTRADA",
        "dispositivo": "Tera HW0009"
    })
    resp_scan_alm = run_coro(api_escanear_codigo_barras(req_scan_alm, user=almacen, db=db))
    assert resp_scan_alm.status_code == 200, "Almacenista debe poder escanear entradas con Tera HW0009"
    payload_alm_scan = json.loads(resp_scan_alm.body.decode("utf-8"))
    assert payload_alm_scan["ok"] is True

    # Reversión limpia del escaneo de prueba
    req_scan_alm_rev = DummyRequest({
        "codigo_barras": "770999000101",
        "warehouse": "1",
        "modo": "SALIDA",
        "dispositivo": "Tera HW0009"
    })
    resp_scan_alm_rev = run_coro(api_escanear_codigo_barras(req_scan_alm_rev, user=almacen, db=db))
    assert resp_scan_alm_rev.status_code == 200

    # Verificación de endpoints del Logo Corporativo
    resp_logo = servir_logo_corporativo()
    assert resp_logo.status_code == 200, "El endpoint del logo corporativo debe responder HTTP 200"
    assert resp_logo.media_type in ("image/webp", "image/png"), f"MIME type del logo incorrecto: {resp_logo.media_type}"
    assert "Cache-Control" in resp_logo.headers

    # Verificación de todos los estados de cada variable del sistema
    estados_sistema_esperados = ["En proceso", "Entregado", "Cancelado"]
    for est in estados_sistema_esperados:
        assert est in models.ESTADOS, f"Estado {est} debe estar definido en models.ESTADOS"

    estados_producto_esperados = ["PENDIENTE", "DESPACHADO", "EN FLY", "LLEGADO FLY", "ENTREGADO"]
    for est_p in estados_producto_esperados:
        assert est_p in models.ESTADOS_PRODUCTO, f"Estado producto {est_p} debe estar definido en models.ESTADOS_PRODUCTO"

    estados_aduana_esperados = ["No aplica / Pendiente", "En revisión aduanera", "Pago de aranceles", "Liberado / Nacionalizado"]
    for est_a in estados_aduana_esperados:
        assert est_a in models.ESTADOS_ADUANA, f"Estado aduana {est_a} debe estar definido en models.ESTADOS_ADUANA"

    convenciones_esperadas = ["VIGENTE", "PRECAUCION", "CRITICO", "VENCIDOS"]
    for conv in convenciones_esperadas:
        assert conv in models.CONVENCIONES_INVENTARIO, f"Convención {conv} debe estar definida en models.CONVENCIONES_INVENTARIO"

    warehouses_esperados = [1, 2, 3]
    for wid in warehouses_esperados:
        assert wid in models.WAREHOUSES, f"Warehouse {wid} debe estar definido en models.WAREHOUSES"

    # 15. Nuevos Requerimientos: Eliminación Bloqueada, Ocultamiento de Estados Finales, Barra Limpia y SLA Gráfico
    html_pedidos_borrar = open(os.path.join(BASE_DIR, "templates", "pedidos.html"), "r", encoding="utf-8").read()
    html_cot_borrar = open(os.path.join(BASE_DIR, "templates", "cotizaciones.html"), "r", encoding="utf-8").read()
    assert "/pedidos/{{ p.id }}/eliminar" not in html_pedidos_borrar, "El botón Borrar de POs debe estar eliminado"
    assert "/cotizaciones/{{ c.id }}/eliminar" not in html_pedidos_borrar, "El botón Borrar de cotizaciones en pedidos debe estar eliminado"
    assert "/cotizaciones/{{ c.id }}/eliminar" not in html_cot_borrar, "El botón Borrar en vista cotizaciones debe estar eliminado"

    # Verificación de que en base.html el enlace a Cotizaciones fue retirado de la barra principal
    html_base_nav = open(os.path.join(BASE_DIR, "templates", "base.html"), "r", encoding="utf-8").read()
    assert 'href="/cotizaciones"' not in html_base_nav, "El enlace a Cotizaciones debe estar retirado de la barra de navegación principal"

    # Verificación de que el módulo SLA no tiene la tabla gigante y cuenta con analítica por asesor y telemetría gráfica
    html_informes_sla = open(os.path.join(BASE_DIR, "templates", "Informes.html"), "r", encoding="utf-8").read()
    assert "Control Maestro de Productos, Proveedores y Estado de Compra" not in html_informes_sla, (
        "La tabla extensa de productos debe ser retirada de Informes.html"
    )
    assert "Semáforo Operativo y Cumplimiento SLA" in html_informes_sla
    assert "Rendimiento y Cumplimiento por Asesor Comercial" in html_informes_sla

    # Verificación de endpoints de eliminación bloqueados institucionalmente
    resp_bloq_po = main.eliminar_pedido(p_split.id, req_qa, user=admin, db=db)
    assert resp_bloq_po.status_code == 303
    assert db.query(models.Pedido).filter_by(id=p_split.id).first() is not None, "La orden no debe borrarse físicamente"
    resp_bloq_cot = main.eliminar_cotizacion(cot_cop.id, req_qa, user=admin, db=db)
    assert resp_bloq_cot.status_code == 303
    assert db.query(models.Cotizacion).filter_by(id=cot_cop.id).first() is not None, "La cotización no debe borrarse físicamente"

    # Verificación del filtro de estado final en Centro de Operaciones:
    # Vista por defecto: Solo órdenes en proceso (las órdenes entregadas y canceladas no se ven por defecto pero sí están)
    resp_mis_ped_default = mis_pedidos(req_qa, user=admin, db=db)
    assert resp_mis_ped_default.status_code == 200

    # 16. Pruebas de Concurrencia Masiva (20 Usuarios Simultáneos Registrando POs y Cotizaciones)
    import concurrent.futures

    def _limpiar_concurrencia():
        try:
            db.query(models.CotizacionItem).filter(models.CotizacionItem.cotizacion_id.in_(
                db.query(models.Cotizacion.id).filter(models.Cotizacion.numero.like("Q-CONC-%"))
            )).delete(synchronize_session=False)
            db.query(models.Cotizacion).filter(models.Cotizacion.numero.like("Q-CONC-%")).delete(synchronize_session=False)
            db.query(models.PedidoInternacional).filter(models.PedidoInternacional.pedido_id.in_(
                db.query(models.Pedido.id).filter(models.Pedido.po.like("PO-CONC-%"))
            )).delete(synchronize_session=False)
            db.query(models.ItemPedido).filter(models.ItemPedido.pedido_id.in_(
                db.query(models.Pedido.id).filter(models.Pedido.po.like("PO-CONC-%"))
            )).delete(synchronize_session=False)
            db.query(models.Pedido).filter(models.Pedido.po.like("PO-CONC-%")).delete(synchronize_session=False)
            db.commit()
        except Exception:
            db.rollback()

    _limpiar_concurrencia()
    usuarios_ids = [admin.id, leonardo.id, miguel.id, carlos.id]

    def _simular_registro_po_concurrente(idx: int):
        uid = usuarios_ids[idx % len(usuarios_ids)]
        with SessionLocal() as db_hilo:
            asesor_hilo = db_hilo.get(models.Asesor, uid)
            po_num = f"PO-CONC-{idx:03d}"
            existente = db_hilo.query(models.Pedido).filter(models.Pedido.po == po_num).first()
            if existente:
                db_hilo.delete(existente)
                db_hilo.commit()
            
            d_conc = {
                "po_numero": po_num,
                "cliente": f"Cliente Concurrente {idx}",
                "fecha_inicial": "2026-11-09",
                "tipo_pedido": "Nacional" if idx % 2 == 0 else "Internacional",
                "plazo_entrega": "2026-11-30",
                "metodo_pago": "Efectivo",
                "items": [{
                    "proveedor": f"Proveedor Aeronáutico {idx}",
                    "parte_numero": f"PN-CONC-{idx:03d}",
                    "descripcion": f"Repuesto de alta presión concurrente {idx}",
                    "cantidad": "2",
                    "precio_unitario": "150000.00",
                    "unidad_medida": "UND",
                    "estado_producto": "PENDIENTE",
                    "sin_iva": "0",
                    "cobro_hazmat": "0.00",
                    "cobro_certificado": "0.00",
                }]
            }
            v_conc = parsear(d_conc)
            p_conc = models.Pedido(po=po_num, tipo=v_conc["tipo"], asesor_id=uid)
            aplicar(db_hilo, p_conc, v_conc, user=asesor_hilo)
            db_hilo.add(p_conc)
            db_hilo.commit()
            return p_conc.id

    def _simular_registro_cot_concurrente(idx: int):
        uid = usuarios_ids[idx % len(usuarios_ids)]
        with SessionLocal() as db_hilo:
            asesor_hilo = db_hilo.get(models.Asesor, uid)
            cot_num = f"Q-CONC-{idx:03d}"
            d_cot_conc = {
                "quote_number": cot_num,
                "quote_date": "2026-11-09",
                "expiration_date": "2026-11-30",
                "to_company": f"Aerolínea Concurrente {idx}",
                "moneda": "COP" if idx % 2 == 0 else "USD",
                "items": [{
                    "description": f"Válvula neumática concurrente {idx}",
                    "qty": "3",
                    "unit": "UND",
                    "unit_price": "220000.00",
                    "cobro_hazmat": "0.00",
                    "cobro_certificado": "0.00",
                    "sin_iva": "0"
                }]
            }
            cot_obj = _guardar_cotizacion_en_bd(db_hilo, asesor_hilo, d_cot_conc)
            db_hilo.commit()
            return cot_obj.id

    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
        futuros_po = [executor.submit(_simular_registro_po_concurrente, i) for i in range(1, 21)]
        ids_pos_creadas = [f.result() for f in concurrent.futures.as_completed(futuros_po)]

    assert len(ids_pos_creadas) == 20, "Deben haberse registrado 20 POs simultáneamente sin error de base de datos bloqueada"

    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
        futuros_cot = [executor.submit(_simular_registro_cot_concurrente, i) for i in range(1, 21)]
        ids_cots_creadas = [f.result() for f in concurrent.futures.as_completed(futuros_cot)]

    assert len(ids_cots_creadas) == 20, "Deben haberse registrado 20 cotizaciones simultáneamente sin error de base de datos bloqueada"

    # Limpieza de registros de prueba de concurrencia
    _limpiar_concurrencia()

    # 17. Verificación de Transportadoras Oficiales y Enlaces de Rastreo (5 Transportadoras Oficiales)
    assert "Inter Rapidísimo" in models.TRANSPORTADORAS_OFICIALES, "Inter Rapidísimo debe estar en TRANSPORTADORAS_OFICIALES"
    assert "Inter Rapidísimo" in main.TRANSPORTADORAS_OPCIONES, "Inter Rapidísimo debe estar en main.TRANSPORTADORAS_OPCIONES"

    # Pruebas de detección y URLs exactas
    info_cm = models.detectar_transportadora_tracking(tracking="CLK-8899", llega_click_mail="Si")
    assert "clickmail.com.co/tracking" in info_cm["url"], f"URL Clickmail incorrecta: {info_cm['url']}"

    info_fdx = models.detectar_transportadora_tracking(tracking="FX1234567890", transportadora_manual="FedEx")
    assert "fedex.com" in info_fdx["url"] and ("tracking.html" in info_fdx["url"] or "fedextrack" in info_fdx["url"])

    info_dhl = models.detectar_transportadora_tracking(tracking="DHL1234567", transportadora_manual="DHL")
    assert "dhl.com/co-es/home/rastreo.html" in info_dhl["url"]

    info_crd = models.detectar_transportadora_tracking(tracking="COORD998877", transportadora_manual="Coordinadora")
    assert "coordinadora.com/rastreo/rastreo-de-guia/" in info_crd["url"]

    info_inter = models.detectar_transportadora_tracking(tracking="INT-998877", transportadora_manual="Inter Rapidísimo")
    assert "siguetuenvio.interrapidisimo.com" in info_inter["url"]
    assert info_inter["carrier"] == "Inter Rapidísimo"

    # 18. Bloqueo de Asesor Comercial y Responsabilidad Directa
    asesor = miguel  # Rol Asesor comercial
    # Un asesor regular no puede reasignar una cotización a otro asesor
    d_cot_bloqueo = {
        "quote_number": "Q-TEST-BLOQ-01",
        "quote_date": "2026-11-09",
        "to_company": "Cliente Prueba Bloqueo",
        "asesor_id": str(admin.id),  # Intenta asignárselo a admin
        "prepared_by": "Juan Esteban Martínez",
        "items": [{
            "description": "Filtro hidráulico", "qty": "1", "unit": "UND", "unit_price": "50000.00", "sin_iva": "0"
        }]
    }
    cot_bloqueada = _guardar_cotizacion_en_bd(db, asesor, d_cot_bloqueo)
    db.commit()
    assert cot_bloqueada.asesor_id == asesor.id, "El asesor comercial no puede cambiar la responsabilidad de su cotización"
    assert cot_bloqueada.prepared_by == asesor.nombre, "El asesor comercial no puede falsificar el prepared_by"
    db.delete(cot_bloqueada)
    db.commit()

    # 19. Notificaciones por Correo: Restricción a Nivel Directivo y Nuevas Opciones
    # Asesor regular bloqueado
    req_notif_asesor = DummyRequest({
        "tipo_notificacion": "seguimiento_po",
        "destino": "supervisor@flylogistic.aero"
    })
    resp_notif_bloq = run_coro(main.pedidos_enviar_correo_individual(p_split.id, req_notif_asesor, user=asesor, db=db))
    assert resp_notif_bloq.status_code == 303, "El asesor regular debe recibir redirección"
    assert "Acción restringida" in req_notif_asesor.session.get("flash", {}).get("msg", ""), (
        "El asesor debe ser bloqueado con mensaje explicativo"
    )

    # Administrador / Gerencia permitido
    req_notif_admin = DummyRequest({
        "tipo_notificacion": "productos_por_comprar",
        "destino": "compras@flylogistic.aero",
        "adjuntar_po": "1"
    })
    resp_notif_ok = run_coro(main.pedidos_enviar_correo_individual(p_split.id, req_notif_admin, user=admin, db=db))
    assert resp_notif_ok.status_code == 303
    assert req_notif_admin.session.get("flash", {}).get("ok") is not None

    # Verificación de construcción de las 3 opciones de notificación
    as_1, txt_1, html_1 = main.construir_notificacion_po_opciones(p_split, "seguimiento_po")
    assert "Seguimiento" in as_1 and "COM-FM-08" in txt_1

    as_2, txt_2, html_2 = main.construir_notificacion_po_opciones(p_split, "productos_por_comprar")
    assert "Repuestos Pendientes" in as_2 and "REPUESTOS REQUERIDOS" in txt_2

    as_3, txt_3, html_3 = main.construir_notificacion_po_opciones(p_split, "cumplimiento_sla")
    assert "SLA" in as_3 and "SEGUIMIENTO DE PLAZO" in txt_3

    # 20. Verificación de Limpieza de Botones y Centralización de Reportes Excel
    html_inf = open(os.path.join(BASE_DIR, "templates", "Informes.html"), "r", encoding="utf-8").read()
    assert "Exportar Reportes Excel" not in html_inf, "Informes.html NO debe tener botón de Exportar Reportes Excel"
    assert "Volver al Centro de Operaciones" not in html_inf, "Informes.html NO debe tener botón de Volver al Centro de Operaciones"

    html_sup = open(os.path.join(BASE_DIR, "templates", "Supervision.html"), "r", encoding="utf-8").read()
    assert "Exportar Reportes Excel" not in html_sup, "Supervision.html NO debe tener botón de Exportar Reportes Excel"
    assert "Ver Cartera Global en Centro de Operaciones" not in html_sup, "Supervision.html NO debe tener botón de Volver al Centro de Operaciones"

    html_usr = open(os.path.join(BASE_DIR, "templates", "Usuarios.html"), "r", encoding="utf-8").read()
    assert "Ir a Supervisión de Equipo" not in html_usr, "Usuarios.html NO debe tener botón de Ir a Supervisión de Equipo"
    assert "Volver al Centro de Operaciones" not in html_usr, "Usuarios.html NO debe tener botón de Volver al Centro de Operaciones"

    html_base = open(os.path.join(BASE_DIR, "templates", "base.html"), "r", encoding="utf-8").read()
    assert "menuExcelInv" not in html_base, "base.html en vista inventario NO debe tener menú de Exportar Reportes Excel"
    assert "Inventario (3 Bodegas)" not in html_base, "base.html debe decir estrictamente Inventario sin (3 Bodegas)"
    assert "Inventario (3 Warehouses)" not in html_base

    html_pedidos = open(os.path.join(BASE_DIR, "templates", "pedidos.html"), "r", encoding="utf-8").read()
    assert "Exportar Reportes Excel" in html_pedidos, "Centro de Operaciones debe ser el concentrador de Exportar Reportes Excel"
    assert "/cotizaciones/exportar/excel" in html_pedidos, "El menú concentrador debe incluir la exportación de cotizaciones"

    resp_exp_cot = main.exportar_cotizaciones_excel(user=admin, db=db)
    assert resp_exp_cot.status_code == 200, "El endpoint /cotizaciones/exportar/excel debe responder HTTP 200 con el archivo generado"

    # 21. Validación Rigurosa de los 8 Reportes Excel Centralizados (HTTP 200 y Generación de Archivo)
    resp_exp_1 = main.descargar_consolidado_individual(user=admin, db=db)
    assert resp_exp_1.status_code == 200, "1. Consolidado de POs debe responder HTTP 200"

    resp_exp_2 = main.descargar_excel_maestro_productos(user=admin, db=db)
    assert resp_exp_2.status_code == 200, "2. Control Maestro de Productos por PO debe responder HTTP 200"

    resp_exp_3 = main.descargar_informe_gerencial(user=admin, db=db)
    assert resp_exp_3.status_code == 200, "3. Informe Ejecutivo de Gestión y SLA debe responder HTTP 200"

    resp_exp_4 = main.exportar_cotizaciones_excel(user=admin, db=db)
    assert resp_exp_4.status_code == 200, "4. Cartera de Cotizaciones debe responder HTTP 200"

    resp_exp_5 = main.exportar_excel_inventario_maestro(user=admin, db=db)
    assert resp_exp_5.status_code == 200, "5. Inventario Maestro Consolidado debe responder HTTP 200"

    resp_exp_6 = main.exportar_excel_warehouse(1, user=admin, db=db)
    assert resp_exp_6.status_code == 200, "6. Inventario Bodega 1 debe responder HTTP 200"

    resp_exp_7 = main.exportar_excel_warehouse(2, user=admin, db=db)
    assert resp_exp_7.status_code == 200, "7. Inventario Bodega 2 debe responder HTTP 200"

    resp_exp_8 = main.exportar_excel_warehouse(3, user=admin, db=db)
    assert resp_exp_8.status_code == 200, "8. Inventario Bodega 3 debe responder HTTP 200"

    # Validación con rol Asesor Comercial (debe responder 200 filtrado a sus pedidos)
    resp_exp_asesor_sla = main.descargar_informe_gerencial(asesor=str(asesor.id), user=asesor, db=db)
    assert resp_exp_asesor_sla.status_code == 200, "Asesor debe poder descargar su propio Informe de Gestión y SLA"

    resp_exp_asesor_maestro = main.descargar_excel_maestro_productos(asesor=str(asesor.id), user=asesor, db=db)
    assert resp_exp_asesor_maestro.status_code == 200, "Asesor debe poder descargar su propio Control Maestro de Productos"

    # 22. Validación QA Rigurosa del Cotizador Formato V4 y Sincronización Google Drive
    class DummyFormRequest:
        def __init__(self, payload=None, path="/cotizador/generar"):
            self.session = {}
            self.headers = {}
            self.method = "POST"
            self.url = type("URLObj", (), {"path": path})()
            self.scope = {"type": "http", "session": self.session}
            self._payload = payload or {}

        async def form(self):
            class FormDataMock(dict):
                def getlist(self, key):
                    val = self.get(key)
                    if val is None:
                        return []
                    if isinstance(val, list):
                        return val
                    return [val]
            return FormDataMock(self._payload)

    # 22a. Prueba: Guardar y Descargar Excel V4 (.xlsx) directamente en el navegador
    req_cot_excel = DummyFormRequest({
        "quote_number": "Q-QA-TEST-001",
        "quote_date": "2026-11-09",
        "expiration_date": "2026-11-24",
        "to_company": "AeroRepublica QA",
        "ref_no": "RFQ-QA-999",
        "contact": "Ing. Pruebas QA",
        "phone": "+57 311 000 0000",
        "email": "qa@aerorepublica.com",
        "site": "Bogota D.C.",
        "terms": "Contado",
        "incoterm": "EXW",
        "delivery_time": "3 a 5 dias",
        "moneda": "COP",
        "tipo_operacion": "Nacional",
        "part_number": ["EF-TEST-100", "EF-TEST-200"],
        "unit": ["UND", "EA"],
        "qty": ["2", "1"],
        "unit_price": ["150000", "280000"],
        "cobro_hazmat": ["0", "20000"],
        "cobro_certificado": ["5000", "0"],
        "sin_iva": ["0", "0"],
        "iva_pct": ["19", "19"],
        "remarks": ["Repuesto critico", "Certificado FAA 8130"],
        "transport": "15000",
        "accion_posterior": "descargar_excel",
    }, path="/cotizador/generar")

    resp_cot_excel = run_coro(main.cotizador_guardar(req_cot_excel, user=admin, db=db))
    assert resp_cot_excel.status_code == 200, "Guardar y Descargar Excel V4 debe retornar HTTP 200"
    assert "application/vnd.openxmlformats" in (resp_cot_excel.media_type or "").lower() or resp_cot_excel.media_type == main.XLSX_MIME
    assert os.path.exists(resp_cot_excel.path), "El archivo Excel V4 generado debe existir físicamente en el disco"

    # Verificar que la cotización se registró en base de datos
    cot_guardada = db.query(models.Cotizacion).filter(models.Cotizacion.numero == "Q-QA-TEST-001").first()
    assert cot_guardada is not None, "La cotización Q-QA-TEST-001 debe existir en base de datos"
    assert len(cot_guardada.items) == 2, "La cotización debe tener sus 2 ítems registrados"
    assert cot_guardada.archivo_excel and os.path.exists(cot_guardada.archivo_excel), "La ruta al Excel debe estar guardada en la cotización"

    # 22b. Prueba: Descarga de Excel existente vía /cotizaciones/{cid}/excel
    resp_get_excel = main.descargar_cotizacion_excel(cot_guardada.id, req_qa, user=admin, db=db)
    assert resp_get_excel.status_code == 200, "Descarga directa /cotizaciones/{cid}/excel debe responder HTTP 200"
    assert os.path.exists(resp_get_excel.path), "El archivo descargado debe existir"

    # 22c. Prueba: Guardar y Convertir en Orden (PO)
    req_cot_conv = DummyFormRequest({
        "quote_number": "Q-QA-TEST-CONV",
        "to_company": "Avianca QA",
        "part_number": ["CONV-PART-1"],
        "unit": ["UND"],
        "qty": ["1"],
        "unit_price": ["500000"],
        "accion_posterior": "convertir_po",
    }, path="/cotizador/generar")
    resp_cot_conv = run_coro(main.cotizador_guardar(req_cot_conv, user=admin, db=db))
    assert resp_cot_conv.status_code == 303, "Convertir PO debe redireccionar HTTP 303"
    assert "/pedidos/nuevo?cot_id=" in resp_cot_conv.headers["location"]

    # 22d. Prueba: Guardar y Volver al Panel
    req_cot_panel = DummyFormRequest({
        "quote_number": "Q-QA-TEST-PANEL",
        "to_company": "LATAM QA",
        "part_number": ["PANEL-PART-1"],
        "unit": ["UND"],
        "qty": ["1"],
        "unit_price": ["320000"],
        "accion_posterior": "guardar_ir_panel",
    }, path="/cotizador/guardar")
    resp_cot_panel = run_coro(main.cotizador_guardar(req_cot_panel, user=admin, db=db))
    assert resp_cot_panel.status_code == 303, "Guardar y volver debe redireccionar HTTP 303"
    assert resp_cot_panel.headers["location"] == "/pedidos"

    # 22e. Prueba: Endpoint de sincronización masiva Google Drive
    req_drive = DummyFormRequest(path="/drive/sincronizar_todo")
    resp_drive = main.sincronizar_todo_drive_endpoint(req_drive, user=admin, db=db)
    assert resp_drive.status_code == 303, "Sincronizar Drive debe redireccionar a /pedidos"
    assert "flash" in req_drive.session, "Debe reportar el resultado de sincronización en flash session"

    # 22f. Validación de consistencia en plantillas HTML
    html_cot_form = open(os.path.join(BASE_DIR, "templates", "cotizador_form.html"), "r", encoding="utf-8").read()
    assert 'action="/cotizador/generar"' in html_cot_form, "Formulario debe apuntar al endpoint /cotizador/generar"
    assert 'name="accion_posterior"' in html_cot_form and 'value="descargar_excel"' in html_cot_form, "Botón descargar excel debe estar presente"
    assert 'Sincronizar con Drive' in html_pedidos, "pedidos.html debe contar con el botón Sincronizar con Drive"

    print("OK - Todas las pruebas unitarias, financieras, horarias, concurrencia FOR UPDATE, tracking y arquitectura superadas al 100%.")


if __name__ == "__main__":
    ejecutar_qa_completo()