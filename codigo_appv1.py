import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import numpy as np
import io
from datetime import datetime

st.set_page_config(page_title="Procesador de Matriz TMERT", layout="wide")

def extraer_datos_empresa(wb):
    """Extrae los datos de la empresa desde la hoja 1"""
    try:
        sheet = wb["1"]
        datos_empresa = {
            "razon_social": "",
            "rut": "",
            "actividad_economica": "",
            "codigo_ciiu": "",
            "direccion": "",
            "comuna": "",
            "representante_legal": "",
            "organismo_administrador": "",
            "fecha_inicio": "",
            "centro_trabajo": "",
            "trabajadores_hombres": 0,
            "trabajadores_mujeres": 0
        }
        for row in range(1, 40):  # Revisar las primeras 40 filas
            row_data = [sheet.cell(row=row, column=col).value for col in range(1, 20)]
            row_text = ' '.join([str(x) for x in row_data if x is not None])

            # Buscar datos específicos
            if "Razón Social" in row_text:
                for col in range(1, 20):
                    if sheet.cell(row=row, column=col).value == "Razón Social":
                        datos_empresa["razon_social"] = sheet.cell(row=row, column=col+1).value
                    if "RUT" in str(sheet.cell(row=row, column=col).value):
                        datos_empresa["rut"] = sheet.cell(row=row, column=col+1).value
            # (Continúa de manera similar para las demás columnas...)

        return datos_empresa

    except Exception as e:
        st.error(f"Error al extraer datos de la empresa: {str(e)}")
        return {}

def extraer_puestos_trabajo(wb):
    """Extrae los puestos de trabajo desde la hoja 2"""
    try:
        sheet = wb["2"]
        puestos = []
        max_row = sheet.max_row
        for row in range(12, min(100, max_row + 1)):
            cell_value = sheet.cell(row=row, column=1).value
            if cell_value and isinstance(cell_value, (int, float)):
                numero = int(cell_value)
                area = sheet.cell(row=row, column=2).value or ""
                puesto = sheet.cell(row=row, column=3).value or ""
                # (Continúa el código de la misma manera...)
        return puestos
    except Exception as e:
        st.error(f"Error al extraer puestos de trabajo: {str(e)}")
        return []

def procesar_matriz(archivo_excel):
    """Función principal para procesar la matriz"""
    try:
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        datos_empresa = extraer_datos_empresa(wb)
        puestos_trabajo = extraer_puestos_trabajo(wb)
        # Continúa con las funciones de evaluación, resultados, etc.

        # Asegurémonos de que todos los datos estén listos para mostrar
        if datos_empresa and puestos_trabajo:
            # Genera el DataFrame resumen
            df_resumen = pd.DataFrame(puestos_trabajo)  # Ejemplo, ajusta según lo que necesites
            # Asegúrate de definir cómo generar observaciones y bytes para el archivo Excel
            observaciones = []  # Aquí va la lógica para generar observaciones
            excel_bytes = io.BytesIO()  # Genera el archivo Excel en memoria

            return datos_empresa, df_resumen, observaciones, excel_bytes
        else:
            raise ValueError("No se han extraído datos correctamente.")

    except Exception as e:
        st.error(f"Error al procesar la matriz TMERT: {str(e)}")
        return None, None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    st.title("Procesador de Matriz TMERT")
    
    st.write("""Esta aplicación procesa archivos Excel con matrices TMERT (Trastornos Musculoesqueléticos Relacionados al Trabajo) y genera un informe estructurado con los resultados de la evaluación.""")

    archivo = st.file_uploader("Cargar archivo de Matriz TMERT", type=["xlsx"])
    
    if archivo is not None:
        with st.spinner("Procesando archivo..."):
            datos_empresa, df_resumen, observaciones, excel_bytes = procesar_matriz(archivo)
        
        # Verificar si los datos fueron procesados correctamente
        if datos_empresa and not df_resumen.empty and observaciones and excel_bytes:
            st.success("¡Archivo procesado correctamente!")
            # Mostrar datos de la empresa
            st.header("Datos de la empresa")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Razón Social:** {datos_empresa['razon_social']}")
                st.write(f"**RUT:** {datos_empresa['rut']}")
                st.write(f"**Actividad Económica:** {datos_empresa['actividad_economica']}")
                st.write(f"**Dirección:** {datos_empresa['direccion']}, {datos_empresa['comuna']}")
            with col2:
                st.write(f"**Representante Legal:** {datos_empresa['representante_legal']}")
                st.write(f"**Organismo Administrador:** {datos_empresa['organismo_administrador']}")
                st.write(f"**Centro de Trabajo:** {datos_empresa['centro_trabajo']}")
                st.write(f"**Trabajadores:** {datos_empresa['trabajadores_hombres']} hombres, {datos_empresa['trabajadores_mujeres']} mujeres")
            
            # Mostrar tabla de resumen
            st.header("Resumen de evaluación por puesto de trabajo")
            st.dataframe(df_resumen, use_container_width=True)
            
            # Mostrar observaciones
            st.header("Observaciones")
            for obs in observaciones:
                st.write(obs)
            
            # Descargar Excel
            st.download_button(
                label="Descargar informe Excel",
                data=excel_bytes,
                file_name=f"Informe_TMERT_{datos_empresa['razon_social']}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
