import streamlit as st
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
import io
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="PDF Tool Pro", page_icon="📄")

st.title("📄 PDF a Excel o Google Sheets")

# --- SIDEBAR / CONFIGURACIÓN ---
st.sidebar.header("Configuración")
modo = st.sidebar.radio("Selecciona destino:", ["Solo Excel", "Google Sheets"])

uploaded_pdfs = st.file_uploader("Sube tus PDFs aquí", type=['pdf'], accept_multiple_files=True)

if uploaded_pdfs:
    custom_name = st.text_input("Nombre del archivo/hoja:", value="Auditoria_Facturas")

    # --- OPCIÓN 1: EXCEL ---
    if modo == "Solo Excel":
        if st.button("Procesar y Descargar Excel"):
            with st.spinner('Creando Excel...'):
                wb = Workbook()
                is_first = True
                for archivo_pdf in uploaded_pdfs:
                    sheet_name = archivo_pdf.name[:30].replace(".pdf", "")
                    ws = wb.active if is_first else wb.create_sheet(title=sheet_name)
                    if is_first: ws.title = sheet_name; is_first = False
                    
                    ws["A1"] = f"FACTURAS - {archivo_pdf.name}"
                    # (Aquí va tu lógica de imágenes que ya tenías)
                    # ... [He omitido el relleno de imágenes para acortar, pero mantén tu lógica aquí]
                
                buf = io.BytesIO()
                wb.save(buf)
                st.download_button("📥 Descargar Excel", buf.getvalue(), file_name=f"{custom_name}.xlsx")

    # --- OPCIÓN 2: GOOGLE SHEETS ---
    elif modo == "Google Sheets":
        st.info("Asegúrate de haber compartido el Google Sheet con el email de tu credentials.json")
        sheet_url = st.text_input("Pega el enlace (URL) de tu Google Sheet:")
        
        if st.button("Sincronizar con Google Sheets"):
            try:
                # Autenticación
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
                client = gspread.authorize(creds)
                
                # Abrir el documento
                sh = client.open_by_url(sheet_url)
                
                with st.spinner('Actualizando Google Sheets...'):
                    for archivo_pdf in uploaded_pdfs:
                        sheet_name = archivo_pdf.name[:30].replace(".pdf", "")
                        
                        # Intentar crear la pestaña o seleccionarla si ya existe
                        try:
                            ws = sh.add_worksheet(title=sheet_name, rows="100", cols="20")
                        except:
                            ws = sh.worksheet(sheet_name)
                        
                        # Escribir algo de datos
                        ws.update('A1', [[f"FACTURA: {archivo_pdf.name}"]])
                        ws.format("A1", {"textFormat": {"bold": True, "fontSize": 14}})
                        
                        st.write(f"✅ Hoja '{sheet_name}' lista.")
                        
                st.success("¡Google Sheets actualizado con éxito! 🎉")
                
            except Exception as e:
                st.error(f"Error: {e}")
                st.info("Recuerda que debes tener el archivo 'credentials.json' en la carpeta.")
