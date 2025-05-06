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
        # Acceder a la hoja 1
        sheet = wb["1"]
        
        # Inicializar diccionario para almacenar datos
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
        
        # Buscar datos en la hoja
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
            
            if "Actividad Económica" in row_text:
                for col in range(1, 20):
                    if sheet.cell(row=row, column=col).value == "Actividad Económica":
                        datos_empresa["actividad_economica"] = sheet.cell(row=row, column=col+1).value
                    if "Código CIIU" in str(sheet.cell(row=row, column=col).value):
                        datos_empresa["codigo_ciiu"] = sheet.cell(row=row, column=col+1).value
            
            if "Dirección" in row_text:
                for col in range(1, 20):
                    if sheet.cell(row=row, column=col).value == "Dirección" or sheet.cell(row=row, column=col).value == "Dirección ":
                        datos_empresa["direccion"] = sheet.cell(row=row, column=col+1).value
                    if "Comuna" in str(sheet.cell(row=row, column=col).value):
                        datos_empresa["comuna"] = sheet.cell(row=row, column=col+1).value
            
            if "Representante Legal" in row_text:
                for col in range(1, 20):
                    if "Representante Legal" in str(sheet.cell(row=row, column=col).value):
                        datos_empresa["representante_legal"] = sheet.cell(row=row, column=col+1).value
            
            if "Organismo administrador" in row_text:
                for col in range(1, 20):
                    if "Organismo administrador" in str(sheet.cell(row=row, column=col).value):
                        datos_empresa["organismo_administrador"] = sheet.cell(row=row, column=col+1).value
                    if "Fecha inicio" in str(sheet.cell(row=row, column=col).value):
                        fecha = sheet.cell(row=row, column=col+1).value
                        if isinstance(fecha, datetime):
                            datos_empresa["fecha_inicio"] = fecha.strftime('%d/%m/%Y')
                        else:
                            datos_empresa["fecha_inicio"] = str(fecha)
            
            if "centro de trabajo" in row_text:
                for col in range(1, 20):
                    if "centro de trabajo" in str(sheet.cell(row=row, column=col).value):
                        datos_empresa["centro_trabajo"] = sheet.cell(row=row, column=col+1).value
            
            if "Nº de trabajadores" in row_text:
                for col in range(1, 20):
                    if "Hombres" in str(sheet.cell(row=row, column=col).value):
                        val = sheet.cell(row=row, column=col+1).value
                        datos_empresa["trabajadores_hombres"] = 0 if val is None else int(val)
                    if "Mujeres" in str(sheet.cell(row=row, column=col).value):
                        val = sheet.cell(row=row, column=col+1).value
                        datos_empresa["trabajadores_mujeres"] = 0 if val is None else int(val)
        
        return datos_empresa
    
    except Exception as e:
        st.error(f"Error al extraer datos de la empresa: {str(e)}")
        return {}

def extraer_puestos_trabajo(wb):
    """Extrae los puestos de trabajo desde la hoja 2"""
    try:
        # Acceder a la hoja 2
        sheet = wb["2"]
        
        # Lista para almacenar los puestos de trabajo
        puestos = []
        
        # Buscar filas con información de puestos de trabajo
        max_row = sheet.max_row
        for row in range(12, min(100, max_row + 1)):  # Limitar a 100 filas máximo
            cell_value = sheet.cell(row=row, column=1).value
            
            # Verificar si es una fila de puesto de trabajo (tiene número en la primera columna)
            if cell_value and isinstance(cell_value, (int, float)):
                numero = int(cell_value)
                area = sheet.cell(row=row, column=2).value or ""
                puesto = sheet.cell(row=row, column=3).value or ""
                tareas = sheet.cell(row=row, column=4).value or ""
                descripcion = sheet.cell(row=row, column=5).value or ""
                horario = sheet.cell(row=row, column=6).value or ""
                
                # Trabajadores (pueden estar en distintas columnas)
                trabajadores_hombres = 0
                trabajadores_mujeres = 0
                
                for col in range(7, 15):
                    header = sheet.cell(row=12, column=col).value or ""
                    header_below = sheet.cell(row=13, column=col).value or ""
                    
                    if "Hombre" in str(header_below):
                        val = sheet.cell(row=row, column=col).value
                        if val is not None and isinstance(val, (int, float)):
                            trabajadores_hombres = int(val)
                    
                    if "Mujer" in str(header_below):
                        val = sheet.cell(row=row, column=col).value
                        if val is not None and isinstance(val, (int, float)):
                            trabajadores_mujeres = int(val)
                
                # Añadir el puesto a la lista
                puestos.append({
                    "numero": numero,
                    "area": area,
                    "puesto": puesto,
                    "tareas": tareas,
                    "descripcion": descripcion,
                    "horario": horario,
                    "trabajadores_hombres": trabajadores_hombres,
                    "trabajadores_mujeres": trabajadores_mujeres
                })
        
        return puestos
    
    except Exception as e:
        st.error(f"Error al extraer puestos de trabajo: {str(e)}")
        return []

def extraer_evaluacion_inicial(wb):
    """Extrae la evaluación inicial desde la hoja 3"""
    try:
        # Acceder a la hoja 3
        sheet = wb["3"]
        
        # Lista para almacenar las evaluaciones iniciales
        evaluaciones = []
        
        # Buscar filas con información de evaluación inicial
        max_row = sheet.max_row
        for row in range(12, min(100, max_row + 1)):  # Limitar a 100 filas máximo
            cell_value = sheet.cell(row=row, column=1).value
            
            # Verificar si es una fila de evaluación (tiene número en la primera columna)
            if cell_value and isinstance(cell_value, (int, float)):
                numero = int(cell_value)
                puesto = sheet.cell(row=row, column=2).value or ""
                tareas = sheet.cell(row=row, column=3).value or ""
                
                # Factores de riesgo
                trabajo_repetitivo = sheet.cell(row=row, column=4).value == "SI"
                postura_estatica = sheet.cell(row=row, column=5).value == "SI"
                mmc_ldt = sheet.cell(row=row, column=6).value == "SI"
                mmc_ea = sheet.cell(row=row, column=7).value == "SI"
                mmp = sheet.cell(row=row, column=8).value == "SI"
                vibracion_cc = sheet.cell(row=row, column=9).value == "SI"
                vibracion_mb = sheet.cell(row=row, column=10).value == "SI"
                
                # Resultado
                resultado = sheet.cell(row=row, column=11).value or ""
                
                # Añadir la evaluación a la lista
                evaluaciones.append({
                    "numero": numero,
                    "puesto": puesto,
                    "tareas": tareas,
                    "trabajo_repetitivo": trabajo_repetitivo,
                    "postura_estatica": postura_estatica,
                    "mmc_ldt": mmc_ldt,
                    "mmc_ea": mmc_ea,
                    "mmp": mmp,
                    "vibracion_cc": vibracion_cc,
                    "vibracion_mb": vibracion_mb,
                    "resultado": resultado
                })
        
        return evaluaciones
    
    except Exception as e:
        st.error(f"Error al extraer evaluación inicial: {str(e)}")
        return []

def buscar_resultado_evaluacion(wb, hoja, numero_puesto):
    """Busca el resultado de la evaluación avanzada para un puesto en una hoja específica"""
    try:
        if hoja not in wb.sheetnames:
            return "Hoja no encontrada"
        
        sheet = wb[hoja]
        
        # Buscar fila con el número de puesto
        for row in range(15, min(100, sheet.max_row + 1)):
            cell_value = sheet.cell(row=row, column=1).value
            
            if cell_value and isinstance(cell_value, (int, float)) and int(cell_value) == numero_puesto:
                # Buscar el resultado en la fila
                resultado = None
                celdas_clave_vacias = True
                
                # Verificar si las celdas importantes están vacías
                celdas_importantes = []
                for col in range(5, 30):  # Columnas donde suelen estar los resultados
                    val = sheet.cell(row=row, column=col).value
                    if isinstance(val, str) and len(val) > 0:
                        celdas_importantes.append(val)
                        if ("Aceptable" in val or "aceptable" in val or 
                            "Crítico" in val or "crítico" in val or
                            "Intermedio" in val or "Solicitar evaluación" in val):
                            celdas_clave_vacias = False
                
                # Si no hay celdas importantes con contenido, está incompleta
                if len(celdas_importantes) == 0:
                    return "No contestada por usuario, revisar"
                
                # Si hay contenido pero no hay resultado claro
                if celdas_clave_vacias:
                    return "No contestada por usuario, revisar"
                
                # Buscar texto específico de resultado
                for col in range(5, 40):
                    val = sheet.cell(row=row, column=col).value
                    if isinstance(val, str):
                        if "Intermedio" in val and "Solicitar evaluación" in val:
                            return "Intermedio - Solicitar evaluación OAL"
                        elif "Ausencia total del riesgo" in val:
                            return "Ausencia total del riesgo"
                        elif "Aceptable" == val:
                            return "Aceptable"
                        elif "No aceptable" == val or "No Aceptable" == val:
                            return "No Aceptable"
                        elif "Crítico" in val:
                            return "Crítico"
                
                # Si no encontramos un resultado claro, pero hay contenido
                return "Resultado no determinado, revisar"
        
        return "No se encontró evaluación"
    
    except Exception as e:
        st.error(f"Error al buscar resultado en hoja {hoja}: {str(e)}")
        return "Error en evaluación"

def extraer_resultados_avanzados(wb, puesto, evaluaciones_iniciales):
    """Extrae los resultados de evaluaciones avanzadas para un puesto"""
    # Buscar evaluación inicial del puesto
    eval_inicial = next((e for e in evaluaciones_iniciales if e["numero"] == puesto["numero"]), None)
    
    if not eval_inicial:
        return {
            "trabajo_repetitivo": "No hay datos",
            "postura_estatica": "No hay datos",
            "mmc_ldt": "No hay datos",
            "mmc_ea": "No hay datos",
            "mmp": "No hay datos",
            "vibracion_cc": "No hay datos",
            "vibracion_mb": "No hay datos"
        }
    
    # Obtener resultados para cada factor de riesgo
    resultados = {}
    
    # Trabajo repetitivo (hoja 4)
    if eval_inicial["trabajo_repetitivo"]:
        resultados["trabajo_repetitivo"] = buscar_resultado_evaluacion(wb, "4", puesto["numero"])
    else:
        resultados["trabajo_repetitivo"] = "No aplica (NO)"
    
    # Postura estática (hoja 5)
    if eval_inicial["postura_estatica"]:
        resultados["postura_estatica"] = buscar_resultado_evaluacion(wb, "5", puesto["numero"])
    else:
        resultados["postura_estatica"] = "No aplica (NO)"
    
    # MMC - Levantamiento/Descenso/Transporte (hoja 6)
    if eval_inicial["mmc_ldt"]:
        resultados["mmc_ldt"] = buscar_resultado_evaluacion(wb, "6", puesto["numero"])
    else:
        resultados["mmc_ldt"] = "No aplica (NO)"
    
    # MMC - Empuje/Arrastre (hoja 7)
    if eval_inicial["mmc_ea"]:
        resultados["mmc_ea"] = buscar_resultado_evaluacion(wb, "7", puesto["numero"])
    else:
        resultados["mmc_ea"] = "No aplica (NO)"
    
    # Manejo Manual de Pacientes (hoja 8)
    if eval_inicial["mmp"]:
        resultados["mmp"] = buscar_resultado_evaluacion(wb, "8", puesto["numero"])
    else:
        resultados["mmp"] = "No aplica (NO)"
    
    # Vibración Cuerpo Completo (hoja 10)
    if eval_inicial["vibracion_cc"]:
        resultados["vibracion_cc"] = buscar_resultado_evaluacion(wb, "10", puesto["numero"])
    else:
        resultados["vibracion_cc"] = "No aplica (NO)"
    
    # Vibración Mano-Brazo (hoja 9)
    if eval_inicial["vibracion_mb"]:
        resultados["vibracion_mb"] = buscar_resultado_evaluacion(wb, "9", puesto["numero"])
    else:
        resultados["vibracion_mb"] = "No aplica (NO)"
    
    return resultados

def generar_tabla_resumen(puestos, evaluaciones, resultados):
    """Genera una tabla de resumen con los resultados"""
    # Crear dataframe
    tabla = []
    
    for puesto in puestos:
        num = puesto["numero"]
        res = resultados.get(num, {})
        
        fila = {
            "Número": num,
            "Área de trabajo": puesto["area"],
            "Puesto de trabajo": puesto["puesto"],
            "Tareas del puesto": puesto["tareas"],
            "Trabajo repetitivo": res.get("trabajo_repetitivo", "No hay datos"),
            "Postura estática": res.get("postura_estatica", "No hay datos"),
            "MMC - LDT": res.get("mmc_ldt", "No hay datos"),
            "MMC - EA": res.get("mmc_ea", "No hay datos"),
            "MMP": res.get("mmp", "No hay datos"),
            "Vibración CC": res.get("vibracion_cc", "No hay datos"),
            "Vibración MB": res.get("vibracion_mb", "No hay datos")
        }
        
        tabla.append(fila)
    
    # Convertir a dataframe
    df_tabla = pd.DataFrame(tabla)
    
    return df_tabla

def generar_observaciones(df_resumen):
    """Genera observaciones basadas en el resumen"""
    observaciones = []
    
    # Nota explicativa general
    observaciones.append("NOTAS EXPLICATIVAS:")
    observaciones.append("- \"Intermedio - Solicitar evaluación OAL\" significa que la condición es No Aceptable pero No Crítica, requiriendo evaluación por el Organismo Administrador de la Ley (ACHS).")
    observaciones.append("- \"No aplica (NO)\" significa que este factor de riesgo fue evaluado en la identificación inicial y se determinó que no está presente en esa tarea.")
    observaciones.append("- \"No contestada por usuario, revisar\" indica que aunque el factor de riesgo fue identificado en la evaluación inicial como presente (SI), la evaluación avanzada está incompleta o las celdas de resultado están vacías.")
    observaciones.append("")
    
    # Observaciones específicas
    observaciones.append("OBSERVACIONES ESPECÍFICAS:")
    
    # Contar puestos con cada tipo de evaluación
    total_puestos = len(df_resumen)
    puestos_sin_riesgo = 0
    puestos_incompletos = 0
    puestos_intermedios = 0
    puestos_criticos = 0
    
    for _, row in df_resumen.iterrows():
        has_risk = False
        incomplete = False
        has_intermediate = False
        has_critical = False
        
        for factor in ["Trabajo repetitivo", "Postura estática", "MMC - LDT", "MMC - EA", "MMP", "Vibración CC", "Vibración MB"]:
            if "No aplica" not in row[factor]:
                has_risk = True
                
                if "No contestada" in row[factor]:
                    incomplete = True
                
                if "Intermedio" in row[factor]:
                    has_intermediate = True
                
                if "Crítico" in row[factor]:
                    has_critical = True
        
        if not has_risk:
            puestos_sin_riesgo += 1
        
        if incomplete:
            puestos_incompletos += 1
        
        if has_intermediate:
            puestos_intermedios += 1
        
        if has_critical:
            puestos_criticos += 1
    
    # Añadir observaciones basadas en el análisis
    observaciones.append(f"1. Se evaluaron {total_puestos} puestos de trabajo.")
    observaciones.append(f"2. {puestos_sin_riesgo} puestos no presentan factores de riesgo que requieran evaluación avanzada.")
    
    if puestos_incompletos > 0:
        observaciones.append(f"3. {puestos_incompletos} puestos tienen evaluaciones incompletas que requieren revisión.")
    
    observaciones.append(f"4. {puestos_intermedios} puestos presentan condición Intermedia que requiere evaluación por parte del OAL.")
    
    if puestos_criticos > 0:
        observaciones.append(f"5. {puestos_criticos} puestos presentan condición Crítica que requiere atención inmediata.")
    
    # Factores de riesgo más frecuentes
    factores_presentes = {}
    for factor in ["Trabajo repetitivo", "Postura estática", "MMC - LDT", "MMC - EA", "MMP", "Vibración CC", "Vibración MB"]:
        count = 0
        for _, row in df_resumen.iterrows():
            if "No aplica" not in row[factor]:
                count += 1
        factores_presentes[factor] = count
    
    # Ordenar factores por frecuencia
    factores_ordenados = sorted(factores_presentes.items(), key=lambda x: x[1], reverse=True)
    
    if factores_ordenados[0][1] > 0:
        observaciones.append(f"6. Los factores de riesgo más frecuentes son: {factores_ordenados[0][0]} ({factores_ordenados[0][1]} puestos) y {factores_ordenados[1][0]} ({factores_ordenados[1][1]} puestos).")
    
    return observaciones

def generar_excel_salida(datos_empresa, df_resumen, observaciones):
    """Genera un archivo Excel con los resultados"""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    # Hoja 1: Datos de la empresa
    df_empresa = pd.DataFrame([datos_empresa])
    df_empresa.to_excel(writer, sheet_name='Datos Empresa', index=False)
    
    # Hoja 2: Tabla de resumen
    df_resumen.to_excel(writer, sheet_name='Resumen Evaluación', index=False)
    
    # Hoja 3: Observaciones
    df_observaciones = pd.DataFrame({'Observaciones': observaciones})
    df_observaciones.to_excel(writer, sheet_name='Observaciones', index=False)
    
    # Ajustar anchos de columna
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column].width = min(adjusted_width, 50)
    
    writer.close()
    return output.getvalue()

def procesar_matriz(archivo_excel):
    """Función principal que procesa la matriz TMERT"""
    try:
        # Cargar el archivo Excel
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        
        # Extraer datos
        datos_empresa = extraer_datos_empresa(wb)
        puestos_trabajo = extraer_puestos_trabajo(wb)
        evaluaciones_iniciales = extraer_evaluacion_inicial(wb)
        
        # Extraer resultados avanzados
        resultados = {}
        for puesto in puestos_trabajo:
            resultados[puesto["numero"]] = extraer_resultados_avanzados(wb, puesto, evaluaciones_iniciales)
        
        # Generar tabla de resumen
        df_resumen = generar_tabla_resumen(puestos_trabajo, evaluaciones_iniciales, resultados)
        
        # Generar observaciones
        observaciones = generar_observaciones(df_resumen)
        
        # Generar Excel de salida
        excel_bytes = generar_excel_salida(datos_empresa, df_resumen, observaciones)
        
        return datos_empresa, df_resumen, observaciones, excel_bytes
    
    except Exception as e:
        st.error(f"Error al procesar la matriz TMERT: {str(e)}")
        return None, None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    st.title("Procesador de Matriz TMERT")
    
    st.write("""
    Esta aplicación procesa archivos Excel con matrices TMERT (Trastornos Musculoesqueléticos Relacionados al Trabajo)
    y genera un informe estructurado con los resultados de la evaluación.
    """)
    
    # Cargar archivo
    archivo = st.file_uploader("Cargar archivo de Matriz TMERT", type=["xlsx"])
    
    if archivo is not None:
        with st.spinner("Procesando archivo..."):
            datos_empresa, df_resumen, observaciones, excel_bytes = procesar_matriz(archivo)
        
        if datos_empresa and df_resumen and observaciones and excel_bytes:
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
