import streamlit as st
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
import io

st.set_page_config(page_title="PDF to Excel", page_icon="📄")

st.title("📄 PDF to Excel (Separate Sheets)")
st.write("Upload multiple PDF files. The tool will convert every page into an image and place each PDF file in its own separate Excel sheet.")

uploaded_pdfs = st.file_uploader("Upload your PDF files here", type=['pdf'], accept_multiple_files=True)

if uploaded_pdfs:
    if st.button("Process and Create Excel"):
        with st.spinner('Processing documents... please wait.'):
            
            # Create the Excel workbook
            wb = Workbook()
            default_sheet = wb.active
            is_first_sheet = True
            
            ancho_deseado = 500  # Desired width for images
            
            # LOOP THROUGH EACH PDF FILE
            for archivo_pdf in uploaded_pdfs:
                
                # --- PREPARE THE SHEET NAME ---
                # 1. Get the file name and remove ".pdf"
                sheet_name = archivo_pdf.name.replace(".pdf", "")
                
                # 2. Remove characters that Excel doesn't allow in sheet names
                invalid_chars =['\\', '/', '?', '*', '[', ']', ':']
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, "")
                    
                # 3. Excel only allows 31 characters max for a sheet name
                sheet_name = sheet_name[:31] 
                
                # --- CREATE OR REUSE THE SHEET ---
                if is_first_sheet:
                    ws = default_sheet
                    ws.title = sheet_name
                    is_first_sheet = False
                else:
                    ws = wb.create_sheet(title=sheet_name)
                
                # --- SET UP THE NEW SHEET ---
                ws["D2"] = "FACTURAS"
                ws["D2"].font = Font(bold=True, size=14)
                ws["D2"].alignment = Alignment(horizontal="center")
                
                # Reset the counters for THIS specific sheet
                fila_actual = 4  
                image_counter = 0 
                
                # Read the PDF
                pdf_bytes = archivo_pdf.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                
                # Loop through each page
                for pagina in doc:
                    pix = pagina.get_pixmap(dpi=150)
                    img_data = pix.tobytes("png")
                    
                    image_stream = io.BytesIO(img_data)
                    img_excel = ExcelImage(image_stream)
                    
                    proporcion = ancho_deseado / pix.width
                    alto_deseado = int(pix.height * proporcion)
                    
                    img_excel.width = ancho_deseado
                    img_excel.height = alto_deseado
                    
                    if image_counter % 2 == 0:
                        celda_destino = f"A{fila_actual}"
                    else:
                        celda_destino = f"H{fila_actual}"
                        
                    ws.add_image(img_excel, celda_destino)
                    
                    if image_counter % 2 != 0:
                        filas_necesarias = int(alto_deseado / 20) + 2
                        fila_actual += filas_necesarias
                    
                    image_counter += 1
            
            # Save Excel in memory
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.success("Process completed successfully! 🎉")
            
            # Download button
            st.download_button(
                label="📥 Download Excel file",
                data=excel_buffer,
                file_name="auditoria_facturas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
