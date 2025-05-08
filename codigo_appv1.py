import streamlit as st
import openpyxl
from fpdf import FPDF
from io import BytesIO

# Configuración de página
st.set_page_config(page_title="Excel a PDF - TMERT", layout="centered")
st.title("📄 Conversor Excel a PDF - TMERT")
st.markdown("Sube tu archivo Excel y genera un informe PDF con los datos.")

# Función para limpiar texto
def limpiar_texto(texto):
    if texto is None:
        return ""
    return str(texto).strip()

# Función para generar el PDF
def generar_pdf(datos_empresa):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", "B", 16)

    # Título del informe
    pdf.cell(0, 10, "Informe TMERT - Datos de la Empresa", ln=1, align="C")
    pdf.ln(10)

    # Logo (opcional) - si tienes un logo.png en la misma carpeta
    # pdf.image("logo.png", x=10, y=10, w=40)

    # Fuente normal
    pdf.set_font("Arial", size=12)

    # Agregar datos
    for campo, valor in datos_empresa.items():
        if valor:
            pdf.cell(0, 8, f"{campo}: {valor}", ln=1)

    # Guardar PDF en memoria
    pdf_bytes = BytesIO()
    pdf.output(pdf_bytes)
    pdf_bytes.seek(0)
    return pdf_bytes

# Función para extraer datos del Excel
def procesar_excel(uploaded_file):
    try:
        # Leer el archivo subido
        wb = openpyxl.load_workbook(filename=BytesIO(uploaded_file.getvalue()), data_only=True)

        if "1" not in wb.sheetnames:
            st.error("❌ No se encontró la hoja '1' en el archivo.")
            return None

        ws = wb["1"]

        # Extraer datos clave (ajustar según estructura real)
        datos_empresa = {
            "Razón Social": limpiar_texto(ws["E15"].value),
            "RUT Empresa": limpiar_texto(ws["L15"].value),
            "Actividad Económica": limpiar_texto(ws["E17"].value),
            "Código CIIU": limpiar_texto(ws["L17"].value),
            "Dirección Empresa": limpiar_texto(ws["E19"].value),
            "Comuna Empresa": limpiar_texto(ws["L19"].value),
            "Representante Legal": limpiar_texto(ws["E21"].value),
            "Organismo Administrador": limpiar_texto(ws["E23"].value),
            "Fecha Inicio": limpiar_texto(ws["L23"].value),
        }

        return datos_empresa

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        return None

# Interfaz de usuario
uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success(f"✅ Archivo '{uploaded_file.name}' cargado.")

    if st.button("🛠️ Generar Informe PDF", type="primary"):
        with st.spinner("Procesando datos..."):
            datos_empresa = procesar_excel(uploaded_file)

            if datos_empresa:
                with st.spinner("Generando PDF..."):
                    pdf_data = generar_pdf(datos_empresa)

                    st.success("🎉 Informe generado exitosamente.")
                    st.download_button(
                        label="⬇️ Descargar PDF",
                        data=pdf_data,
                        file_name=f"Informe_TMERT_{datos_empresa['Razón Social']}.pdf",
                        mime="application/pdf"
                    )

                    # Mostrar resumen en pantalla
                    with st.expander("🔍 Ver datos incluidos en el PDF"):
                        st.write(datos_empresa)
