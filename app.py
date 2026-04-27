import streamlit as st
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
import io

st.set_page_config(page_title="PDF to Excel", page_icon="📄")

st.title("📄 PDF to Excel (Separate Sheets)")
st.write("Upload multiple PDF files. The tool will convert every page into an image and place each PDF file in its own separate Excel sheet.")

# 1. File uploader box
uploaded_pdfs = st.file_uploader("Upload your PDF files here", type=['pdf'], accept_multiple_files=True)

# 2. Everything inside here is HIDDEN until you upload a file
if uploaded_pdfs:
    
    # --- NEW FEATURE: Custom File Name ---
    custom_name = st.text_input("Enter the name for your Excel file:", value="auditoria_facturas")
    
    # We add .xlsx if you forget to type it
    if not custom_name.endswith(".xlsx"):
        custom_name += ".xlsx"

    # 3. The process button
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
                sheet_name = archivo_pdf.name.replace(".pdf", "")
                
                invalid_chars =['\\', '/', '?', '*', '[', ']', ':']
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, "")
                    
                sheet_name = sheet_name[:31] 
                
                # --- CREATE OR REUSE THE SHEET ---
                if is_first_sheet:
                    ws = default_sheet
                    ws.title = sheet_name
                    is_first_sheet = False
                else:
                    # Prevent crashes if two PDFs have the same name
                    original_sheet_name = sheet_name
                    counter = 1
                    while sheet_name in wb.sheetnames:
                        suffix = f"_{counter}"
                        sheet_name = original_sheet_name[:(31 - len(suffix))] + suffix
                        counter += 1
                        
                    ws = wb.create_sheet(title=sheet_name)
                
                # --- SET UP THE NEW SHEET ---
                ws["D2"] = "FACTURAS"
                ws["D2"].font = Font(bold=True, size=14)
                ws["D2"].alignment = Alignment(horizontal="center")
                
                # Make columns A and H wider for better presentation
                ws.column_dimensions['A'].width = 70
                ws.column_dimensions['H'].width = 70
                
                fila_actual = 4  
                image_counter = 0 
                
                # Read the PDF
                pdf_bytes = archivo_pdf.read()
                
                # Prevent crashes from corrupted or password-protected PDFs
                try:
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    if doc.needs_pass:
                        st.warning(f"File {archivo_pdf.name} is password protected and was skipped.")
                        continue
                except Exception as e:
                    st.error(f"Error processing {archivo_pdf.name}: {e}")
                    continue
                
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
                file_name=custom_name, # USING THE CUSTOM NAME HERE
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
