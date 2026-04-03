# Automatización de notificaciones de vencimientos por correo

Script en Python que automatiza el envío masivo de correos personalizados desde Outlook, con reportes adjuntos en Excel generados dinámicamente por grupo de cliente.

Desarrollado para reemplazar un proceso 100% manual que requería filtrar datos, crear archivos individuales y enviar correos uno por uno.

---

## Problema que resuelve

En empresas con carteras grandes de clientes, notificar vencimientos de contratos de forma manual implica:

- Filtrar el Excel por cada cliente
- Guardar un archivo por separado
- Redactar y enviar cada correo individualmente

Con este script, todo ese proceso ocurre con **un solo comando**.

---

## Qué hace el script

1. Lee un archivo Excel con todos los contratos y sus datos de contacto
2. Agrupa la información por cliente
3. Genera un archivo Excel personalizado por grupo, con formato automático
4. Envía un correo desde Outlook a cada cliente con su reporte adjunto
5. Incluye los datos del asesor asignado en el cuerpo del correo
6. Registra el avance en consola en tiempo real

---

## Tecnologías

| Librería | Uso |
|---|---|
| `pandas` | Lectura y agrupación del Excel fuente |
| `openpyxl` | Formato automático de los archivos generados |
| `win32com` | Integración con Microsoft Outlook |
| `os` | Manejo de rutas y carpetas |

---

## Estructura del Excel fuente

El archivo de entrada debe tener al menos estas columnas (los nombres son configurables):

| Columna | Descripción |
|---|---|
| `Grupo Cliente` | Nombre del grupo — define cómo se agrupan los envíos |
| `Correo Contacto` | Email del destinatario principal |
| `Asesor` | Nombre del ejecutivo de cuenta |
| `Correo Asesor` | Email del ejecutivo |
| `Teléfono Asesor` | Teléfono del ejecutivo |

---

## Cómo usarlo

### 1. Instalar dependencias

```bash
pip install pandas openpyxl pywin32
```

### 2. Configurar rutas y columnas

Edita la sección `CONFIGURACIÓN` al inicio del archivo `notificaciones_vencimientos.py`:

```python
ARCHIVO_EXCEL  = r'C:\ruta\a\tu\archivo\VENCIMIENTOS.xlsx'
CARPETA_SALIDA = r'C:\ruta\a\tu\carpeta\REPORTES_POR_CLIENTE'
BCC_INTERNOS   = "supervisor@empresa.com; gerencia@empresa.com"
```

### 3. Ejecutar

```bash
python notificaciones_vencimientos.py
```

La consola mostrará el progreso en tiempo real:

```
Se procesarán 12 grupos de clientes...

[1/12] Correo enviado a: Cliente A (contacto@clienteA.com)
[2/12] Correo enviado a: Cliente B (gerencia@clienteB.com)
...
Proceso completado. Todos los correos fueron enviados.
```

---

## Requisitos del sistema

- Windows con Microsoft Outlook instalado y configurado
- Python 3.8 o superior
- Acceso al archivo Excel fuente

---

## Casos de uso similares

Este script puede adaptarse fácilmente para:

- Notificaciones de pagos vencidos o por vencer
- Recordatorios de renovación de pólizas o servicios
- Reportes periódicos personalizados a clientes
- Alertas de inventario o stock mínimo

---

## Autor

Desarrollado como parte de una automatización real en empresa del sector financiero-automotriz.
Disponible para adaptación a otros procesos administrativos.
