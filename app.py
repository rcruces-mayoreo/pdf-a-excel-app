import streamlit as st
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
import io

st.set_page_config(page_title="PDF a Excel", page_icon="📄")

st.title("📄 PDF a Excel (Tamaño Ajustado y 2 Columnas)")
st.write("Sube tu archivo PDF. La herramienta convertirá cada página y las ajustará al tamaño perfecto.")

archivo_pdf = st.file_uploader("Sube tu archivo PDF aquí", type=['pdf'])

if archivo_pdf is not None:
    if st.button("Procesar y Crear Excel"):
        with st.spinner('Procesando documento y ajustando imágenes...'):
            pdf_bytes = archivo_pdf.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Páginas del PDF"
            
            # --- TÍTULO OPCIONAL --- 
            # Replicando el título "FACTURAS" de tu Imagen 1
            ws["D2"] = "FACTURAS"
            ws["D2"].font = Font(bold=True, size=14)
            ws["D2"].alignment = Alignment(horizontal="center")
            
            fila_actual = 4  # Empezamos a pegar en la fila 4 (como en tu ejemplo)
            
            # --- TAMAÑO DESEADO ---
            # 500 píxeles de ancho da un tamaño casi idéntico al de tu Imagen 1
            ancho_deseado = 500 
            
            for i in range(len(doc)):
                pagina = doc[i]
                
                # Extraer imagen con buena calidad
                pix = pagina.get_pixmap(dpi=150)
                img_data = pix.tobytes("png")
                
                image_stream = io.BytesIO(img_data)
                img_excel = ExcelImage(image_stream)
                
                # --- LA MAGIA DEL TAMAÑO ---
                # Calculamos la proporción de alto para no deformar la imagen
                proporcion = ancho_deseado / pix.width
                alto_deseado = int(pix.height * proporcion)
                
                # Le asignamos el nuevo tamaño a la imagen de Excel
                img_excel.width = ancho_deseado
                img_excel.height = alto_deseado
                
                # --- LÓGICA DE 2 COLUMNAS (A y H) ---
                # Si es un número par (0, 2, 4...) va en la Columna A
                # Si es impar (1, 3, 5...) va en la Columna H
                if i % 2 == 0:
                    celda_destino = f"A{fila_actual}"
                else:
                    celda_destino = f"H{fila_actual}"
                    
                ws.add_image(img_excel, celda_destino)
                
                # Solo bajamos de fila cuando ya pusimos las 2 imágenes en la misma fila
                if i % 2 != 0:
                    # Una fila de Excel mide aprox 20 píxeles de alto.
                    # Calculamos cuántas filas ocupa la imagen y le sumamos 2 filas de margen.
                    filas_necesarias = int(alto_deseado / 20) + 2
                    fila_actual += filas_necesarias
            
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.success("¡El proceso ha terminado con éxito! 🎉")
            
            st.download_button(
                label="📥 Descargar archivo Excel",
                data=excel_buffer,
                file_name="auditoria_facturas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
