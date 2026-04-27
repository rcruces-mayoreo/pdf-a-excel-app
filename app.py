import streamlit as st
import fitz  # PyMuPDF
import io
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

# --- FUNCIÓN DE CONEXIÓN (MÉTODO NUBE) ---
def conectar_google():
    # Esto leerá la clave desde la configuración secreta de Streamlit
    claves = st.secrets["google_cloud"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(claves, scopes=scopes)
    return gspread.authorize(creds)

st.set_page_config(page_title="PDF Tool Pro", page_icon="📄")
st.title("📄 PDF a Excel & Google Sheets")

modo = st.sidebar.radio("Enviar a:", ["Solo Excel", "Google Sheets"])
uploaded_pdfs = st.file_uploader("Sube tus PDFs", type=['pdf'], accept_multiple_files=True)

if uploaded_pdfs:
    nombre_doc = st.text_input("Nombre del archivo:", value="Auditoria_Facturas")

    if modo == "Google Sheets":
        url_sheet = st.text_input("Pega la URL de tu Google Sheet (Recuerda compartirla con el email del robot):")
        
        if st.button("🚀 Sincronizar con Goog
