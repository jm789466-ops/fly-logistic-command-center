# Fly Logistic — Centro de Comando y Gestión Operativa Aeronáutica (OMS)

Sistema integral de gestión de órdenes de compra (PO), cotizaciones comerciales, control de inventario aeronáutico multibodega (Warehouses 1, 2 y 3) y analítica de cumplimiento SLA. Desarrollado con arquitectura asíncrona de alto desempeño en FastAPI, motor relacional SQLAlchemy y renderizado responsivo con Tailwind CSS y Alpine.js.

---

## Características Principales

### 1. Centro de Operaciones Logísticas
- **Gestión Integral de Órdenes (PO):** Registro, edición, inspección detallada y semaforización de plazos en días hábiles (calendario bancario y festivos de Colombia).
- **Protección de Datos e Historial:** Las órdenes no se eliminan físicamente; cuentan con flujo de cancelación lógica con retención de motivos y opción de reactivación autorizada por Gerencia/Administración.
- **Trazabilidad de Envíos en Tiempo Real:** Integración directa con 5 operadores logísticos oficiales:
  - Clickmail (`https://clickmail.com.co/tracking`)
  - FedEx (`https://www.fedex.com/es-co/tracking.html`)
  - Coordinadora (`https://coordinadora.com/rastreo/rastreo-de-guia/`)
  - DHL (`https://www.dhl.com/co-es/home/rastreo.html`)
  - Inter Rapidísimo (`https://siguetuenvio.interrapidisimo.com/`)

### 2. Módulo de Cotizaciones y Exportación Excel V4
- **Generador de Cotizaciones Técnicas:** Soporte para hasta 14 ítems con cálculo dinámico de subtotales, recargos Hazmat, cargos por Certificados aeronáuticos (FAA 8130 / EASA Form 1), impuestos y flete de transporte.
- **Exportación Fiel a Plantilla Corporativa:** Llenado directo de la plantilla oficial `FORMATO COTIZACION VERSION 4.xlsx` conservando membretes, fórmulas, estilos y firmas institucionales.
- **Conversión Inmediata a PO:** Creación de orden de compra en un solo clic a partir de una cotización aprobada.

### 3. Bodega e Inventario Aeronáutico Multibodega
- **Control de 3 Almacenes:** Gestión física y contable para Warehouse 1, Warehouse 2 y Warehouse 3, además de vista de Inventario Maestro Consolidado.
- **Compatibilidad con Lectores de Códigos de Barras:** Integración optimizada para escaneo rápido (pistola Tera HW0009 y equivalentes estándar).
- **Despacho Automático a Órdenes:** Trazabilidad de números de parte (PN), números de serie (SN), lotes y reducción automática de existencias al asociar material a una PO.
- **Rol Exclusivo de Almacenista:** Perfil de acceso restringido únicamente al módulo de inventario con bloqueo de navegación administrativa o comercial.

### 4. Analítica SLA y Supervisión Operativa
- **Métricas de Cumplimiento:** Medición de tiempos de entrega, porcentaje de cumplimiento de fechas de compromiso y distribución de carga de trabajo.
- **Centralización de Reportes:** Generador unificado de 8 reportes ejecutivos en formato Excel `.xlsx` ubicado exclusivamente en el Centro de Operaciones.

### 5. Respaldo y Concurrencia Empresarial
- **Alta Concurrencia (20+ Usuarios Simultáneos):** Base de datos SQLite optimizada con modo WAL (`PRAGMA journal_mode=WAL`), bloqueo extendido a 60 segundos (`PRAGMA busy_timeout=60000`), caché en memoria y bucles de reintento automático con retroceso exponencial ante contención de escritura.
- **Sincronización Automática con Google Drive:** Servicio silencioso y desatendido que respalda órdenes y cotizaciones en la nube en segundo plano al iniciar el servidor y al registrar documentos.

---

## Requisitos del Sistema

- **Python:** 3.10 o superior (compatible con Python 3.11, 3.12 y 3.13 en Windows/Linux/macOS).
- **Navegadores Soportados:** Chrome, Edge, Firefox, Safari (últimas versiones estables).

---

## Estructura del Proyecto

```text
OMS_Pedidos/
├── main.py                             # Aplicación principal FastAPI, rutas y controladores
├── models.py                           # Modelos de datos SQLAlchemy y lógica de negocio
├── database.py                         # Configuración del motor de base de datos y PRAGMAs WAL
├── excel_com_fm_08.py                  # Generación y llenado técnico de reportes Excel
├── test_qa_fly.py                      # Suite integral de pruebas automatizadas QA
├── requirements.txt                    # Dependencias oficiales del proyecto
├── FORMATO COTIZACION VERSION 4.xlsx   # Plantilla maestra de cotizaciones aeronáuticas
├── logo_fly.webp                       # Identidad visual corporativa de Fly Logistic
├── correo_config.example.json          # Plantilla para configuración del servidor SMTP
├── .gitignore                          # Exclusiones de seguridad para Git
├── templates/                          # Vistas HTML con Jinja2, Tailwind y Alpine.js
│   ├── base.html                       # Layout maestro y navegación corporativa
│   ├── pedidos.html                    # Centro de Operaciones (tablas de PO y cotizaciones)
│   ├── pedido_form.html                # Formulario de registro y edición de órdenes de compra
│   ├── cotizaciones.html               # Panel dedicado de cotizaciones
│   ├── cotizador_form.html             # Formulario de emisión de cotizaciones V4
│   ├── Informes.html                   # Analítica SLA y gráficos de rendimiento
│   ├── Supervision.html                # Supervisión de equipo y auditoría operativa
│   ├── Usuarios.html                   # Administración de usuarios y roles
│   ├── perfil.html                     # Perfil de usuario y seguridad de contraseña
│   ├── login.html                      # Inicio de sesión corporativo
│   ├── registro.html                   # Solicitud de nueva cuenta
│   ├── olvide.html                     # Solicitud de restablecimiento de contraseña
│   └── restablecer.html                # Ingreso de nueva contraseña mediante token
```

---

## Instalación y Puesta en Marcha

### 1. Clonar el Repositorio
```bash
git clone https://github.com/TU_USUARIO/OMS_Pedidos.git
cd OMS_Pedidos
```

### 2. Configurar el Entorno Virtual
En Windows (PowerShell):
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

En Linux / macOS:
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar Dependencias
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Configurar Variables de Notificación (Opcional)
Copie el archivo de ejemplo para configurar el envío de correos:
```powershell
Copy-Item correo_config.example.json correo_config.json
```
Edite `correo_config.json` con los datos del servidor SMTP corporativo.

### 5. Ejecutar Pruebas Automatizadas
Para verificar la integridad de todos los módulos, finanzas y base de datos:
```powershell
python test_qa_fly.py
```

### 6. Iniciar el Servidor de Operaciones
```powershell
python main.py
```
O ejecutando Uvicorn directamente:
```powershell
uvicorn main:app --host 127.0.0.1 --port 8000 --reload
```

Acceda desde su navegador a:
```text
http://127.0.0.1:8000
```

---

## Usuarios y Roles Iniciales

Al iniciar por primera vez, el sistema genera automáticamente el usuario administrativo predeterminado si no existe:

- **Usuario:** `admin`
- **Contraseña:** `admin1234`
- **Rol:** `Administrador`

Se recomienda cambiar la contraseña desde el módulo **Mi Perfil** tras el primer inicio de sesión.

---

## Licencia y Derechos

Copyright © 2026 Fly Logistic S.A.S. Todos los derechos reservados.
Prohibida su reproducción, distribución o modificación no autorizada.
