import streamlit as st
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
import io

st.set_page_config(page_title="PDF to Excel", page_icon="📄")

# I translated the text to English for you to practice!
st.title("📄 PDF to Excel ")
st.write("Upload one or multiple PDF files. The tool will convert every page into an image and fit them all into a single Excel file.")

# 1. ALLOW MULTIPLE FILES
uploaded_pdfs = st.file_uploader("Upload your PDF files here", type=['pdf'], accept_multiple_files=True)

# If the list of uploaded PDFs is not empty
if uploaded_pdfs:
    if st.button("Process and Create Excel"):
        with st.spinner('Processing documents... please wait.'):
            
            # Create the Excel workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "PDF Pages"
            
            # Title inside the Excel file
            ws["D2"] = "FACTURAS"
            ws["D2"].font = Font(bold=True, size=14)
            ws["D2"].alignment = Alignment(horizontal="center")
            
            fila_actual = 4  # Starting row in Excel
            ancho_deseado = 500  # Desired width for the images
            
            # 2. GLOBAL IMAGE COUNTER
            image_counter = 0 
            
            # 3. LOOP THROUGH EACH PDF FILE
            for archivo_pdf in uploaded_pdfs:
                pdf_bytes = archivo_pdf.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                
                # Loop through each page of the current PDF
                for pagina in doc:
                    # Extract image
                    pix = pagina.get_pixmap(dpi=150)
                    img_data = pix.tobytes("png")
                    
                    image_stream = io.BytesIO(img_data)
                    img_excel = ExcelImage(image_stream)
                    
                    # Resize logic
                    proporcion = ancho_deseado / pix.width
                    alto_deseado = int(pix.height * proporcion)
                    
                    img_excel.width = ancho_deseado
                    img_excel.height = alto_deseado
                    
                    # 4. COLUMN LOGIC USING THE GLOBAL COUNTER
                    if image_counter % 2 == 0:
                        celda_destino = f"A{fila_actual}"
                    else:
                        celda_destino = f"H{fila_actual}"
                        
                    ws.add_image(img_excel, celda_destino)
                    
                    # Move to the next row only when the right column (H) is filled
                    if image_counter % 2 != 0:
                        filas_necesarias = int(alto_deseado / 20) + 2
                        fila_actual += filas_necesarias
                    
                    # Increase the counter for the next image
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
