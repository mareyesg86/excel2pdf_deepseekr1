import streamlit as st
import openpyxl
from fpdf import FPDF
import os
from io import BytesIO
import tempfile

# Configuración inicial de la aplicación
st.set_page_config(page_title="Excel a PDF", page_icon="📄")
st.title("📄 Conversor de Excel a PDF")
st.write("Convierte tus archivos Excel a PDF con formato profesional")

# Mapa de colores a niveles de riesgo (se mantiene igual)
COLOR_TO_RISK_LEVEL = {
    "FF00FF00": "ACEPTABLE",
    "00FF00": "ACEPTABLE",
    "FF008000": "ACEPTABLE",
    "FF92D050": "ACEPTABLE",
    "FFFFFF00": "INTERMEDIO",
    "FFFFC000": "INTERMEDIO",
    "FFFFA500": "INTERMEDIO",
    "FFFF0000": "CRÍTICO",
}

# Funciones auxiliares (add_table_to_pdf y add_structured_entry_to_pdf se mantienen igual)
# [Pega aquí las funciones add_table_to_pdf y add_structured_entry_to_pdf del código original]

def process_excel_to_pdf(uploaded_file):
    """Procesa el archivo Excel subido y genera un PDF"""
    try:
        # Crear un archivo temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            # Escribir el contenido del archivo subido al temporal
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        
        # Ahora podemos cargar el workbook desde el archivo temporal
        wb = openpyxl.load_workbook(tmp_path, data_only=True)
        
        # Crear PDF
        pdf = FPDF(orientation='L', unit='mm', format='A3')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=10)
        pdf.add_page()

        # Procesar hojas (ejemplo con Hoja 1)
        try:
            hoja1 = wb["1"]
            st.success("Procesando Hoja 1...")
            
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, txt="Datos de Hoja 1", ln=True, align="L")
            pdf.ln(2)
            
            # [Aquí va tu lógica de procesamiento para Hoja 1]
            
        except KeyError:
            st.warning("Hoja '1' no encontrada, omitiendo...")
        
        # [Repite el patrón para las otras hojas]

        # Guardar PDF en memoria
        pdf_bytes = BytesIO()
        pdf.output(pdf_bytes)
        pdf_bytes.seek(0)
        
        # Eliminar el archivo temporal
        os.unlink(tmp_path)
        
        return pdf_bytes

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        if 'tmp_path' in locals() and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return None

# Interfaz de usuario
uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Generar PDF"):
        with st.spinner("Procesando archivo..."):
            pdf_bytes = process_excel_to_pdf(uploaded_file)
            
            if pdf_bytes:
                st.success("¡PDF generado con éxito!")
                st.download_button(
                    label="⬇️ Descargar PDF",
                    data=pdf_bytes,
                    file_name=f"{os.path.splitext(uploaded_file.name)[0]}_reporte.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("No se pudo generar el PDF")
