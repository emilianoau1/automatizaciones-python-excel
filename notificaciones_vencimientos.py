"""
notificaciones_vencimientos.py
-------------------------------
Automatización de notificaciones de vencimientos de contratos.

Lee un archivo Excel con información de contratos por vencer,
genera un reporte personalizado por grupo de cliente y lo envía
automáticamente por correo Outlook con el archivo adjunto.

Autor: [Tu nombre]
Herramientas: Python, pandas, openpyxl, win32com (Outlook)
"""

import os
import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
from openpyxl.styles import Alignment


# ─────────────────────────────────────────────
# CONFIGURACIÓN — Ajusta estas rutas y columnas
# ─────────────────────────────────────────────

ARCHIVO_EXCEL = r'C:\ruta\a\tu\archivo\VENCIMIENTOS.xlsx'
CARPETA_SALIDA = r'C:\ruta\a\tu\carpeta\REPORTES_POR_CLIENTE'

# Nombres de columnas en el Excel fuente
COL_GRUPO        = 'Grupo Cliente'        # Agrupa los envíos (ej. por empresa o región)
COL_EMAIL_CLIENTE = 'Correo Contacto'     # Destinatario principal
COL_NOMBRE_ASESOR = 'Asesor'             # Nombre del ejecutivo de cuenta
COL_EMAIL_ASESOR  = 'Correo Asesor'      # Correo del ejecutivo
COL_TEL_ASESOR    = 'Teléfono Asesor'    # Teléfono del ejecutivo

# Correos en copia oculta (supervisores, gerencia, etc.)
BCC_INTERNOS = "supervisor@tuempresa.com; gerencia@tuempresa.com"


# ─────────────────────────────────────────────
# FUNCIONES
# ─────────────────────────────────────────────

def crear_carpeta(ruta):
    """Crea la carpeta de salida si no existe."""
    if not os.path.exists(ruta):
        os.makedirs(ruta)


def formatear_excel(ruta_archivo):
    """
    Abre el Excel generado y aplica formato automático:
    - Ancho de columna ajustado al contenido
    - Texto centrado horizontal y verticalmente
    """
    wb = load_workbook(ruta_archivo)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass

        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(ruta_archivo)


def generar_cuerpo_correo(nombre_grupo, nombre_asesor, correo_asesor, tel_asesor):
    """
    Genera el HTML del cuerpo del correo con los datos del grupo
    y del asesor de cuenta asignado.
    """
    return (
        f"<p>Estimado cliente ({nombre_grupo}):</p>"
        f"<p>Por medio del presente, le compartimos la relación de contratos próximos a vencer, "
        f"así como los detalles y montos correspondientes.</p>"
        f"<p>En caso de estar interesado en renovar o gestionar alguno de estos contratos, "
        f"le solicitamos ponerse en contacto con su asesor asignado, quien podrá brindarle "
        f"mayor información y acompañarlo en el proceso.</p>"
        f"<p><b>Asesor de cuenta:</b><br>"
        f"{nombre_asesor} – {correo_asesor} – {tel_asesor}</p>"
        f"<p>Agradecemos su atención y quedamos atentos a cualquier duda o comentario.</p>"
        f"<p>Atentamente,<br>[Nombre de tu empresa]</p>"
    )


def enviar_correo(destinatarios, bcc, asunto, cuerpo_html, ruta_adjunto):
    """
    Envía un correo desde Outlook con un archivo adjunto.

    Parámetros:
        destinatarios  : str  — correos separados por ";"
        bcc            : str  — correos en copia oculta
        asunto         : str  — asunto del correo
        cuerpo_html    : str  — contenido HTML del cuerpo
        ruta_adjunto   : str  — ruta absoluta del archivo a adjuntar
    """
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = destinatarios
    mail.BCC = bcc
    mail.Subject = asunto
    mail.HTMLBody = cuerpo_html
    mail.Attachments.Add(ruta_adjunto)
    mail.Send()


# ─────────────────────────────────────────────
# PROCESO PRINCIPAL
# ─────────────────────────────────────────────

def main():
    crear_carpeta(CARPETA_SALIDA)

    df = pd.read_excel(ARCHIVO_EXCEL)

    grupos = df.groupby(COL_GRUPO)
    total = len(grupos)

    print(f"Se procesarán {total} grupos de clientes...\n")

    for i, (nombre_grupo, datos_grupo) in enumerate(grupos, start=1):

        # Obtener datos del asesor (primer registro del grupo)
        nombre_asesor = datos_grupo[COL_NOMBRE_ASESOR].iloc[0]
        correo_asesor = datos_grupo[COL_EMAIL_ASESOR].iloc[0]
        tel_asesor    = datos_grupo[COL_TEL_ASESOR].iloc[0]

        # Obtener correos únicos del cliente
        correos = datos_grupo[COL_EMAIL_CLIENTE].dropna().unique()
        destinatarios = "; ".join(correos)

        # Generar nombre de archivo seguro (sin caracteres especiales)
        nombre_archivo = f'vencimientos_{nombre_grupo}.xlsx'.replace('/', '_').replace('\\', '_')
        ruta_archivo = os.path.join(CARPETA_SALIDA, nombre_archivo)

        # Exportar datos del grupo a Excel y aplicar formato
        datos_grupo.to_excel(ruta_archivo, index=False)
        formatear_excel(ruta_archivo)

        # Preparar y enviar correo
        asunto = f'Notificación de vencimientos — {nombre_grupo}'
        cuerpo = generar_cuerpo_correo(nombre_grupo, nombre_asesor, correo_asesor, tel_asesor)
        enviar_correo(destinatarios, BCC_INTERNOS, asunto, cuerpo, ruta_archivo)

        print(f"[{i}/{total}] Correo enviado a: {nombre_grupo} ({destinatarios})")

    print("\nProceso completado. Todos los correos fueron enviados.")


if __name__ == '__main__':
    main()
