import streamlit as st
import pandas as pd
import os
import qrcode
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter

# 1. CONFIGURACIÓN DE LA PÁGINA Y COLORES
st.set_page_config(page_title="Generador de Certificados", page_icon="🎓", layout="centered")

# Inyectamos CSS personalizado para el Azul Marino (#000080) y Blanco
st.markdown("""
    <style>
    /* Fondo general de la aplicación */
    .stApp { background-color: #ffffff; }
    
    /* Textos generales en azul marino */
    h1, h2, h3, p, label { color: #000080 !important; }
    
    /* --- CONFIGURACIÓN DEL BOTÓN --- */
    .stButton > button {
        background-color: #000080 !important; /* Azul marino */
        border-radius: 8px; 
        border: none; 
        width: 100%; 
        padding: 5px;
    }
    /* Forzamos que cualquier elemento dentro del botón sea BLANCO y NEGRITA */
    .stButton > button, 
    .stButton > button * p, 
    .stButton > button div, 
    .stButton > button span {
        color: #ffffff !important;
        font-weight: bold !important;
    }
    /* Color al pasar el mouse por encima */
    .stButton > button:hover { background-color: #0000cd !important; }
    
    /* --- CONFIGURACIÓN DE LAS NOTIFICACIONES (TOASTS) --- */
    div[data-testid="stToast"] {
        background-color: #000080 !important; /* Fondo azul marino */
        border: 1px solid #ffffff !important; /* Pequeño borde blanco para que resalte */
    }
    /* Forzamos que el texto de la notificación sea BLANCO y NEGRITA */
    div[data-testid="stToast"] * {
        color: #ffffff !important;
        font-weight: bold !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 2. INTERFAZ DE USUARIO
st.title("Sistema de Generación de Certificados")
st.write("Sube el archivo de Excel con los datos de los participantes.")

archivo_subido = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx", "xls"])

# 3. LÓGICA DE EJECUCIÓN
if archivo_subido is not None:
    df = pd.read_excel(archivo_subido)
    st.write("Vista previa de los datos:")
    st.dataframe(df.head())
    
    if st.button("Generar y Enviar Certificados"):
        with st.spinner("Conectando con Google Drive y generando certificados..."):
            
            try:
                # =======================================================
                # 1. AUTENTICACIÓN FLEXIBLE (Local vs Nube)
                # =======================================================
                SCOPES = ['https://www.googleapis.com/auth/drive']
                
                # Si estamos en Streamlit Cloud, leemos de los secretos
                if "google_token" in st.secrets:
                    token_info = dict(st.secrets["google_token"])
                    creds = Credentials.from_authorized_user_info(token_info, SCOPES)
                # Si estamos en tu laptop, leemos el archivo local
                else:
                    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
                
                drive_service = build('drive', 'v3', credentials=creds)
                
                # =======================================================
                # 2. CONFIGURACIÓN INICIAL LOCAL
                # =======================================================
                ID_CARPETA_DESTINO = "1_85lcSovQ6UDikiljKUzSG_N4PulQZlI"       
                # Ahora la plantilla debe estar en la misma carpeta que este script
                RUTA_PLANTILLA_PDF = "C:\\Users\\ferna\\Desktop\\Certificados_pediatra\\plantilla_SECIP.pdf" 
                
                pdfmetrics.registerFont(TTFont('IBMPlexCondensed', 'IBMPlexSansCondensed-Bold.ttf'))

                COORD_NOMBRE = (75, 640)  
                COORD_FECHA_COMP = (135, 625) 
                COORD_FECHA_EXP = (125, 619)  
                COORD_QR = (465, 600)         
                TAMANO_QR = 50               

                # =======================================================
                # 3. BUCLE SOBRE EL DATAFRAME DEL EXCEL
                # =======================================================
                # Iteramos directamente sobre las filas del Excel que subió el usuario
                for index, fila in df.iterrows():
                    nombre = str(fila.get('nombre', ''))
                    fecha_comp = str(fila.get('fecha_completado', ''))[:10]
                    fecha_exp = str(fila.get('fecha_expiracion', ''))[:10]

                    if not nombre or nombre == 'nan':
                        continue

                    # Mostramos el progreso en la interfaz de Streamlit
                    st.toast(f"Procesando certificado de: {nombre}...")

                    # PASO A: Crear archivo vacío en Drive
                    file_metadata = {
                        'name': f"Certificado_{nombre}.pdf",
                        'parents': [ID_CARPETA_DESTINO],
                        'mimeType': 'application/pdf'
                    }
                    dummy_file = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
                    file_id = dummy_file.get('id')
                    web_link = dummy_file.get('webViewLink')

                    permission = {'type': 'anyone', 'role': 'reader'}
                    drive_service.permissions().create(fileId=file_id, body=permission).execute()

                    # PASO B: Generar QR
                    # Generamos nombres genéricos y seguros para evitar errores en Windows
                    qr_path = "temp_qr.png"
                    capa_path = "temp_capa.pdf"
                    pdf_final_path = "temp_final.pdf"

                    # PASO B: Generar QR
                    qr = qrcode.make(web_link)
                    qr.save(qr_path)

                    # PASO C: Crear capa ReportLab
                    c = canvas.Canvas(capa_path)
                    c.setFont("IBMPlexCondensed", 10)
                    c.drawString(COORD_NOMBRE[0], COORD_NOMBRE[1], nombre)
                    c.setFont("IBMPlexCondensed", 6)
                    c.drawString(COORD_FECHA_COMP[0], COORD_FECHA_COMP[1], fecha_comp)
                    c.drawString(COORD_FECHA_EXP[0], COORD_FECHA_EXP[1], fecha_exp)
                    c.drawImage(qr_path, COORD_QR[0], COORD_QR[1], width=TAMANO_QR, height=TAMANO_QR)
                    c.save()

                    # PASO D: Fusionar PDFs
                    plantilla = PdfReader(RUTA_PLANTILLA_PDF)
                    capa = PdfReader(capa_path)
                    writer = PdfWriter()
                    pagina_base = plantilla.pages[0]
                    pagina_base.merge_page(capa.pages[0])
                    writer.add_page(pagina_base)

                    with open(pdf_final_path, "wb") as fOut:
                        writer.write(fOut)

                    # PASO E: Subir a Drive
                    with open(pdf_final_path, "rb") as f_final:
                        media = MediaIoBaseUpload(f_final, mimetype='application/pdf', resumable=True)
                        drive_service.files().update(fileId=file_id, media_body=media).execute()

                    # Limpieza segura
                    os.remove(qr_path)
                    os.remove(capa_path)
                    os.remove(pdf_final_path)
                    
                st.success("¡Proceso completado! Los certificados están listos en Drive.")
                
            except Exception as e:
                st.error(f"Hubo un error en el proceso: {e}")