import streamlit as st
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
import io

st.set_page_config(page_title="PDF a Excel", page_icon="📄")

st.title("📄 PDF a Excel (Imágenes PNG)")
st.write("Sube tu archivo PDF. La herramienta convertirá cada página en una imagen PNG y las pegará en orden en un archivo de Excel.")

# Widget para subir archivo
archivo_pdf = st.file_uploader("Sube tu archivo PDF aquí", type=['pdf'])

if archivo_pdf is not None:
    if st.button("Procesar y Crear Excel"):
        with st.spinner('Procesando documento... esto puede tardar unos segundos.'):
            # 1. Leer el PDF desde la memoria
            pdf_bytes = archivo_pdf.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            # 2. Crear el libro de Excel en memoria
            wb = Workbook()
            ws = wb.active
            ws.title = "Páginas del PDF"
            
            # Fila inicial donde se pegará la primera imagen
            fila_actual = 1
            
            # 3. Recorrer cada página del PDF
            for numero_pagina in range(len(doc)):
                pagina = doc[numero_pagina]
                
                # Renderizar la página como imagen (dpi=150 para buena calidad sin ser muy pesada)
                pix = pagina.get_pixmap(dpi=150)
                img_data = pix.tobytes("png")
                
                # Convertir los bytes de la imagen a un formato que openpyxl acepte
                image_stream = io.BytesIO(img_data)
                img_excel = ExcelImage(image_stream)
                
                # Pegar la imagen en la celda correspondiente (A1, A40, etc.)
                celda_destino = f"A{fila_actual}"
                ws.add_image(img_excel, celda_destino)
                
                # Calcular cuántas filas de Excel ocupa la imagen para que la siguiente no se superponga
                # Una fila de Excel por defecto mide aprox 15-20 píxeles.
                filas_necesarias = int(pix.height / 15) + 2 
                fila_actual += filas_necesarias
            
            # 4. Guardar el Excel en memoria para su descarga
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.success("¡El proceso ha terminado con éxito! 🎉")
            
            # 5. Botón de descarga
            st.download_button(
                label="📥 Descargar archivo Excel",
                data=excel_buffer,
                file_name="pdf_convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
