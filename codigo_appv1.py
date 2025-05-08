import streamlit as st
import openpyxl
from fpdf import FPDF
from io import BytesIO

# Configuración de la aplicación
st.set_page_config(page_title="Excel a PDF", layout="wide")
st.title("Conversor de Excel a PDF Profesional")

# Función para procesar el Excel y generar PDF
def excel_to_pdf(uploaded_file):
    try:
        wb = openpyxl.load_workbook(filename=BytesIO(uploaded_file.getvalue()), data_only=True)

        # Crear PDF
        pdf = FPDF(orientation='L', unit='mm', format='A3')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        # Procesar hoja 1 (ejemplo básico)
        if "1" in wb.sheetnames:
            sheet = wb["1"]
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, "Datos de la Empresa", ln=1)

            # Extraer datos clave (ajusta según tu estructura)
            datos = {
                "Razón Social": sheet["E15"].value,
                "RUT": sheet["L15"].value,
                "Actividad": sheet["E17"].value
            }

            pdf.set_font("Arial", "", 12)
            for k, v in datos.items():
                if v:  # Solo mostrar si hay valor
                    pdf.cell(0, 10, f"{k}: {v}", ln=1)

        # Convertir PDF a bytes para descarga
        pdf_bytes = BytesIO()
        pdf.output(dest='S')  # Generar PDF en memoria
        pdf_bytes.write(pdf.output(dest='S').encode('latin1'))  # Escribir en BytesIO
        pdf_bytes.seek(0)
        return pdf_bytes

    except Exception as e:
        st.error(f"Error crítico: {str(e)}")
        return None

# Interfaz de usuario mejorada
st.sidebar.header("Configuración")
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    st.success(f"Archivo {uploaded_file.name} cargado correctamente")

    if st.button("Generar Informe PDF", type="primary"):
        with st.spinner("Procesando..."):
            result = excel_to_pdf(uploaded_file)

            if result:
                st.balloons()
                st.success("Informe generado con éxito!")

                # Botón de descarga
                st.download_button(
                    label="Descargar PDF",
                    data=result,
                    file_name=f"informe_{os.path.splitext(uploaded_file.name)[0]}.pdf",
                    mime="application/pdf"
                )

                # Vista previa (primeras páginas)
                with st.expander("Vista previa del PDF"):
                    st.write("El PDF contiene los siguientes datos:")
                    # Aquí podrías mostrar un resumen de los datos procesados
            else:
                st.error("No se pudo generar el PDF. Verifica el formato del archivo.")

# Información adicional
st.sidebar.markdown("""
**Instrucciones:**
1. Sube tu archivo Excel (.xlsx)
2. Haz clic en 'Generar Informe PDF'
3. Descarga el resultado
""")
