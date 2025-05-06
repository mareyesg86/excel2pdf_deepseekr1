import streamlit as st
import pandas as pd
import openpyxl
import numpy as np
import io
from datetime import datetime
import traceback
import re

st.set_page_config(page_title="Procesador de Matriz TMERT", layout="wide")

def safe_get_cell_value(sheet, row, col):
    """Obtiene el valor de una celda de forma segura"""
    try:
        cell_value = sheet.cell(row=row, column=col).value
        if isinstance(cell_value, str):
            return cell_value.strip()
        return cell_value
    except:
        return None

def is_numeric(value):
    """Verifica si un valor es numérico"""
    if value is None:
        return False
    try:
        float(value)
        return True
    except:
        return False

def extraer_datos_empresa(wb):
    """Extrae los datos de la empresa de forma robusta con varios intentos"""
    st.write("Extrayendo datos de la empresa (método mejorado)...")
    
    # Inicializar diccionario para almacenar datos
    datos_empresa = {
        "razon_social": "No encontrada",
        "rut": "No encontrado",
        "actividad_economica": "No especificada",
        "codigo_ciiu": "No especificado",
        "direccion": "No especificada",
        "comuna": "No especificada",
        "representante_legal": "No especificado",
        "organismo_administrador": "No especificado",
        "fecha_inicio": "No especificada",
        "centro_trabajo": "No especificado",
        "trabajadores_hombres": 0,
        "trabajadores_mujeres": 0
    }
    
    # Intentar con la hoja 1 primero
    if "1" in wb.sheetnames:
        sheet = wb["1"]
        
        # Buscar de forma exhaustiva en toda la hoja
        max_row = min(50, sheet.max_row)
        max_col = min(20, sheet.max_column)
        
        # Imprimir las primeras filas para debug
        st.write("Inspeccionando contenido de la hoja 1:")
        for r in range(1, min(10, max_row)):
            row_content = []
            for c in range(1, min(10, max_col)):
                cell_value = safe_get_cell_value(sheet, r, c)
                if cell_value:
                    row_content.append(str(cell_value))
            if row_content:
                st.write(f"Fila {r}: {' | '.join(row_content)}")
        
        # Buscar términos clave en toda la hoja
        for r in range(1, max_row):
            for c in range(1, max_col):
                cell_value = safe_get_cell_value(sheet, r, c)
                if not cell_value or not isinstance(cell_value, str):
                    continue
                
                cell_lower = cell_value.lower()
                
                # Buscar cada campo específicamente
                if "razón social" in cell_lower or "razon social" in cell_lower:
                    # Buscar el valor a la derecha o abajo
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["razon_social"] = str(right_value)
                        st.write(f"Encontrado: Razón Social = {right_value}")
                
                if "rut" in cell_lower and len(cell_lower) < 10:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["rut"] = str(right_value)
                        st.write(f"Encontrado: RUT = {right_value}")
                
                if "actividad económica" in cell_lower or "actividad economica" in cell_lower:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["actividad_economica"] = str(right_value)
                        st.write(f"Encontrado: Actividad Económica = {right_value}")
                
                if "dirección" in cell_lower or "direccion" in cell_lower:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["direccion"] = str(right_value)
                        st.write(f"Encontrado: Dirección = {right_value}")
                
                if "comuna" in cell_lower and len(cell_lower) < 15:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["comuna"] = str(right_value)
                        st.write(f"Encontrado: Comuna = {right_value}")
                
                if "representante legal" in cell_lower:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["representante_legal"] = str(right_value)
                        st.write(f"Encontrado: Representante Legal = {right_value}")
                
                if "organismo administrador" in cell_lower:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if right_value:
                        datos_empresa["organismo_administrador"] = str(right_value)
                        st.write(f"Encontrado: Organismo Administrador = {right_value}")
                
                # Buscar trabajadores
                if "trabajadores" in cell_lower and "hombres" in cell_lower:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if is_numeric(right_value):
                        datos_empresa["trabajadores_hombres"] = int(float(right_value))
                
                if "trabajadores" in cell_lower and "mujeres" in cell_lower:
                    right_value = safe_get_cell_value(sheet, r, c+1)
                    if is_numeric(right_value):
                        datos_empresa["trabajadores_mujeres"] = int(float(right_value))
    
    else:
        st.warning("No se encontró la hoja 1 en el archivo.")
    
    # Si no encontramos datos en la hoja 1, buscar en otras hojas
    if datos_empresa["razon_social"] == "No encontrada":
        st.write("Buscando información de la empresa en otras hojas...")
        
        # Intentar con la portada
        if "Portada" in wb.sheetnames:
            sheet = wb["Portada"]
            max_row = min(50, sheet.max_row)
            max_col = min(20, sheet.max_column)
            
            for r in range(1, max_row):
                row_data = [safe_get_cell_value(sheet, r, c) for c in range(1, max_col)]
                row_text = ' '.join([str(x) for x in row_data if x is not None and x != ""])
                
                if "astilleros" in row_text.lower() or "sociedad" in row_text.lower():
                    for text in row_data:
                        if text and isinstance(text, str) and len(text) > 5:
                            datos_empresa["razon_social"] = text
                            st.write(f"Encontrado en Portada: Razón Social = {text}")
                            break
    
    # Si seguimos sin información, usar datos por defecto más específicos
    if datos_empresa["razon_social"] == "No encontrada":
        # Revisar el nombre del archivo para extraer información
        try:
            file_name = wb.properties.title
            if file_name and len(file_name) > 5:
                # Buscar posibles nombres de empresa en el nombre del archivo
                name_parts = file_name.split()
                potential_names = [part for part in name_parts if len(part) > 5 and part.lower() not in ["matriz", "tmert", "revisada", "achs"]]
                
                if potential_names:
                    datos_empresa["razon_social"] = " ".join(potential_names)
                    st.write(f"Extraído del nombre del archivo: Razón Social = {datos_empresa['razon_social']}")
        except:
            pass
    
    # Mostrar qué datos se encontraron
    found_data = sum(1 for k, v in datos_empresa.items() if v and v != "No encontrado" and v != "No encontrada" and v != "No especificado" and v != "No especificada" and v != 0)
    st.write(f"Se encontraron {found_data} campos de información de la empresa.")
    
    return datos_empresa

def extraer_puestos_trabajo(wb):
    """Extrae los puestos de trabajo de forma mejorada"""
    st.write("Extrayendo puestos de trabajo (método mejorado)...")
    
    puestos = []
    
    # Intentar con la hoja 2 primero
    found_puestos = False
    
    if "2" in wb.sheetnames:
        sheet = wb["2"]
        max_row = sheet.max_row
        
        # Buscar encabezados típicos
        header_row = None
        for r in range(1, min(20, max_row)):
            row_text = " ".join(str(cell.value) for cell in sheet[r] if cell.value)
            if "puesto" in row_text.lower() and "trabajo" in row_text.lower():
                header_row = r
                break
        
        # Si encontramos encabezados, buscar datos
        if header_row:
            # Determinar qué columnas contienen qué información
            column_mapping = {'numero': None, 'area': None, 'puesto': None, 'tareas': None}
            
            for col in range(1, sheet.max_column):
                header = sheet.cell(row=header_row, column=col).value
                if not header:
                    continue
                
                header_lower = str(header).lower()
                
                if "n°" in header_lower or "numero" in header_lower or "número" in header_lower:
                    column_mapping['numero'] = col
                elif "área" in header_lower or "area" in header_lower:
                    column_mapping['area'] = col
                elif "puesto" in header_lower:
                    column_mapping['puesto'] = col
                elif "tarea" in header_lower:
                    column_mapping['tareas'] = col
            
            # Si no encontramos algunas columnas, usar posiciones por defecto
            if not column_mapping['numero']:
                column_mapping['numero'] = 1
            if not column_mapping['area']:
                column_mapping['area'] = 2
            if not column_mapping['puesto']:
                column_mapping['puesto'] = 3
            if not column_mapping['tareas']:
                column_mapping['tareas'] = 4
            
            # Ahora buscar puestos de trabajo
            for r in range(header_row + 1, max_row + 1):
                num_val = sheet.cell(row=r, column=column_mapping['numero']).value
                
                # Verificar si es una fila de puesto (debe tener un número)
                if num_val and is_numeric(num_val):
                    num = int(float(num_val))
                    area = sheet.cell(row=r, column=column_mapping['area']).value or ""
                    puesto = sheet.cell(row=r, column=column_mapping['puesto']).value or ""
                    tareas = sheet.cell(row=r, column=column_mapping['tareas']).value or ""
                    
                    # Añadir solo si tenemos al menos un dato
                    if area or puesto or tareas:
                        puestos.append({
                            'numero': num,
                            'area': area,
                            'puesto': puesto,
                            'tareas': tareas
                        })
                        found_puestos = True
    
    # Si no encontramos puestos, buscar en otras hojas
    if not found_puestos:
        st.warning("No se encontraron puestos en la hoja 2. Buscando en la hoja 3...")
        
        if "3" in wb.sheetnames:
            sheet = wb["3"]
            max_row = sheet.max_row
            
            # Buscar datos en hoja 3
            for r in range(1, max_row + 1):
                num_val = sheet.cell(row=r, column=1).value
                
                # Verificar si es una fila de puesto (debe tener un número)
                if num_val and is_numeric(num_val):
                    num = int(float(num_val))
                    puesto = sheet.cell(row=r, column=2).value or ""
                    tareas = sheet.cell(row=r, column=3).value or ""
                    
                    # Añadir solo si tenemos al menos un dato
                    if puesto or tareas:
                        puestos.append({
                            'numero': num,
                            'area': "",  # No tenemos área en la hoja 3
                            'puesto': puesto,
                            'tareas': tareas
                        })
                        found_puestos = True
    
    # Si seguimos sin encontrar puestos, crear algunos por defecto
    if not found_puestos:
        st.error("No se pudieron encontrar puestos de trabajo en ninguna hoja.")
        
        # Crear puestos por defecto
        puestos = [
            {'numero': 1, 'area': 'Área 1', 'puesto': 'Puesto 1', 'tareas': 'Tareas del puesto 1'},
            {'numero': 2, 'area': 'Área 2', 'puesto': 'Puesto 2', 'tareas': 'Tareas del puesto 2'}
        ]
    
    st.write(f"Se encontraron {len(puestos)} puestos de trabajo.")
    
    # Ordenar los puestos por número
    puestos.sort(key=lambda x: x['numero'])
    
    return puestos

def extraer_evaluacion_inicial(wb, puestos_trabajo):
    """Extrae la evaluación inicial desde la hoja 3 de forma robusta"""
    st.write("Extrayendo evaluación inicial (método mejorado)...")
    
    evaluaciones = []
    
    if "3" in wb.sheetnames:
        sheet = wb["3"]
        max_row = sheet.max_row
        
        # Buscar encabezados
        header_row = None
        for r in range(1, min(20, max_row)):
            row_text = " ".join(str(cell.value) for cell in sheet[r] if cell.value)
            if "trabajo repetitivo" in row_text.lower() or "postura estática" in row_text.lower():
                header_row = r
                break
        
        if not header_row:
            header_row = 12  # Valor por defecto si no encontramos encabezados
        
        # Ahora buscar evaluaciones
        for r in range(header_row + 1, max_row + 1):
            num_val = sheet.cell(row=r, column=1).value
            
            # Verificar si es una fila de evaluación (debe tener un número)
            if num_val and is_numeric(num_val):
                num = int(float(num_val))
                
                # Buscar el puesto correspondiente
                puesto_info = next((p for p in puestos_trabajo if p['numero'] == num), None)
                
                if puesto_info:
                    # Extraer evaluación
                    eval_data = {
                        'numero': num,
                        'puesto': puesto_info['puesto'],
                        'tareas': puesto_info['tareas'],
                        'trabajo_repetitivo': sheet.cell(row=r, column=4).value == "SI",
                        'postura_estatica': sheet.cell(row=r, column=5).value == "SI",
                        'mmc_ldt': sheet.cell(row=r, column=6).value == "SI",
                        'mmc_ea': sheet.cell(row=r, column=7).value == "SI",
                        'mmp': sheet.cell(row=r, column=8).value == "SI",
                        'vibracion_cc': sheet.cell(row=r, column=9).value == "SI",
                        'vibracion_mb': sheet.cell(row=r, column=10).value == "SI"
                    }
                    
                    evaluaciones.append(eval_data)
    
    # Si no encontramos evaluaciones, crear por defecto basadas en los puestos
    if not evaluaciones:
        st.warning("No se encontraron evaluaciones iniciales. Creando evaluaciones por defecto.")
        
        for puesto in puestos_trabajo:
            evaluaciones.append({
                'numero': puesto['numero'],
                'puesto': puesto['puesto'],
                'tareas': puesto['tareas'],
                'trabajo_repetitivo': False,
                'postura_estatica': False,
                'mmc_ldt': False,
                'mmc_ea': False,
                'mmp': False,
                'vibracion_cc': False,
                'vibracion_mb': False
            })
    
    st.write(f"Se encontraron {len(evaluaciones)} evaluaciones iniciales.")
    return evaluaciones

def buscar_resultado_evaluacion(wb, hoja, numero_puesto):
    """Busca el resultado de evaluación de forma mejorada"""
    try:
        if hoja not in wb.sheetnames:
            return "No aplica (NO)"
        
        sheet = wb[hoja]
        
        # Buscar fila con el número de puesto
        for r in range(1, sheet.max_row + 1):
            cell_val = sheet.cell(row=r, column=1).value
            
            if is_numeric(cell_val) and int(float(cell_val)) == numero_puesto:
                # Buscar resultado en toda la fila
                row_data = [sheet.cell(row=r, column=c).value for c in range(1, 30)]
                
                # Verificar si hay datos relevantes
                has_data = any(val for val in row_data if val)
                
                if not has_data:
                    return "No se encontró evaluación"
                
                # Buscar términos específicos en la fila
                row_text = " ".join(str(val) for val in row_data if val)
                
                if "intermedio" in row_text.lower() or "solicitar evaluación" in row_text.lower():
                    return "Intermedio - Solicitar evaluación OAL"
                elif "aceptable" in row_text.lower() and "no" not in row_text.lower():
                    return "Aceptable"
                elif "no aceptable" in row_text.lower():
                    return "No Aceptable"
                elif "crítico" in row_text.lower() or "critico" in row_text.lower():
                    return "Crítico"
                else:
                    return "No se encontró resultado"
        
        return "No se encontró evaluación"
    except:
        return "Error al evaluar"

def extraer_resultados_avanzados(wb, puesto, evaluaciones_iniciales):
    """Extrae resultados de evaluaciones avanzadas de forma mejorada"""
    # Buscar evaluación inicial
    eval_inicial = next((e for e in evaluaciones_iniciales if e['numero'] == puesto['numero']), None)
    
    if not eval_inicial:
        return {
            "trabajo_repetitivo": "No aplica (NO)",
            "postura_estatica": "No aplica (NO)",
            "mmc_ldt": "No aplica (NO)",
            "mmc_ea": "No aplica (NO)",
            "mmp": "No aplica (NO)",
            "vibracion_cc": "No aplica (NO)",
            "vibracion_mb": "No aplica (NO)"
        }
    
    # Obtener resultados para cada factor con presencia en evaluación inicial
    resultados = {}
    
    if eval_inicial['trabajo_repetitivo']:
        resultados["trabajo_repetitivo"] = buscar_resultado_evaluacion(wb, "4", puesto['numero'])
    else:
        resultados["trabajo_repetitivo"] = "No aplica (NO)"
    
    if eval_inicial['postura_estatica']:
        resultados["postura_estatica"] = buscar_resultado_evaluacion(wb, "5", puesto['numero'])
    else:
        resultados["postura_estatica"] = "No aplica (NO)"
    
    if eval_inicial['mmc_ldt']:
        resultados["mmc_ldt"] = buscar_resultado_evaluacion(wb, "6", puesto['numero'])
    else:
        resultados["mmc_ldt"] = "No aplica (NO)"
    
    if eval_inicial['mmc_ea']:
        resultados["mmc_ea"] = buscar_resultado_evaluacion(wb, "7", puesto['numero'])
    else:
        resultados["mmc_ea"] = "No aplica (NO)"
    
    if eval_inicial['mmp']:
        resultados["mmp"] = buscar_resultado_evaluacion(wb, "8", puesto['numero'])
    else:
        resultados["mmp"] = "No aplica (NO)"
    
    if eval_inicial['vibracion_cc']:
        resultados["vibracion_cc"] = buscar_resultado_evaluacion(wb, "10", puesto['numero'])
    else:
        resultados["vibracion_cc"] = "No aplica (NO)"
    
    if eval_inicial['vibracion_mb']:
        resultados["vibracion_mb"] = buscar_resultado_evaluacion(wb, "9", puesto['numero'])
    else:
        resultados["vibracion_mb"] = "No aplica (NO)"
    
    return resultados

def generar_tabla_resumen(puestos, evaluaciones, resultados):
    """Genera una tabla de resumen mejorada y bien formateada"""
    # Crear dataframe con todos los campos necesarios
    tabla = []
    
    for puesto in puestos:
        num = puesto['numero']
        res = resultados.get(num, {})
        
        # Asegurarnos de tener todos los campos necesarios
        area = puesto.get('area', '')
        puesto_nombre = puesto.get('puesto', '')
        tareas = puesto.get('tareas', '')
        
        fila = {
            "Número": num,
            "Área de trabajo": area,
            "Puesto de trabajo": puesto_nombre,
            "Tareas del puesto": tareas,
            "Trabajo repetitivo": res.get("trabajo_repetitivo", "No aplica (NO)"),
            "Postura estática": res.get("postura_estatica", "No aplica (NO)"),
            "MMC - LDT": res.get("mmc_ldt", "No aplica (NO)"),
            "MMC - EA": res.get("mmc_ea", "No aplica (NO)"),
            "MMP": res.get("mmp", "No aplica (NO)"),
            "Vibración CC": res.get("vibracion_cc", "No aplica (NO)"),
            "Vibración MB": res.get("vibracion_mb", "No aplica (NO)")
        }
        
        tabla.append(fila)
    
    # Ordenar la tabla por número
    tabla.sort(key=lambda x: x["Número"])
    
    # Crear dataframe
    df_tabla = pd.DataFrame(tabla)
    
    return df_tabla

def generar_observaciones(df_resumen):
    """Genera observaciones basadas en el resumen"""
    observaciones = []
    
    # Nota explicativa general
    observaciones.append("NOTAS EXPLICATIVAS:")
    observaciones.append("- \"Intermedio - Solicitar evaluación OAL\" significa que la condición es No Aceptable pero No Crítica, requiriendo evaluación por el Organismo Administrador de la Ley (ACHS).")
    observaciones.append("- \"No aplica (NO)\" significa que este factor de riesgo fue evaluado en la identificación inicial y se determinó que no está presente en esa tarea.")
    observaciones.append("- \"No se encontró evaluación\" indica que aunque el factor de riesgo fue identificado en la evaluación inicial como presente (SI), no se encontró la evaluación avanzada correspondiente.")
    observaciones.append("")
    
    # Observaciones específicas
    observaciones.append("OBSERVACIONES ESPECÍFICAS:")
    
    # Contar puestos con cada tipo de evaluación
    total_puestos = len(df_resumen)
    puestos_sin_riesgo = 0
    puestos_incompletos = 0
    puestos_intermedios = 0
    puestos_aceptables = 0
    
    for _, row in df_resumen.iterrows():
        has_risk = False
        incomplete = False
        has_intermediate = False
        has_acceptable = False
        
        for factor in ["Trabajo repetitivo", "Postura estática", "MMC - LDT", "MMC - EA", "MMP", "Vibración CC", "Vibración MB"]:
            if "No aplica" not in row[factor]:
                has_risk = True
                
                if "No se encontró" in row[factor]:
                    incomplete = True
                
                if "Intermedio" in row[factor]:
                    has_intermediate = True
                
                if "Aceptable" == row[factor]:
                    has_acceptable = True
        
        if not has_risk:
            puestos_sin_riesgo += 1
        
        if incomplete:
            puestos_incompletos += 1
        
        if has_intermediate:
            puestos_intermedios += 1
        
        if has_acceptable:
            puestos_aceptables += 1
    
    # Añadir observaciones basadas en el análisis
    observaciones.append(f"1. Se evaluaron {total_puestos} puestos de trabajo.")
    observaciones.append(f"2. {puestos_sin_riesgo} puestos no presentan factores de riesgo que requieran evaluación avanzada.")
    
    if puestos_incompletos > 0:
        observaciones.append(f"3. {puestos_incompletos} puestos tienen evaluaciones incompletas que requieren revisión.")
    
    if puestos_intermedios > 0:
        observaciones.append(f"4. {puestos_intermedios} puestos presentan condición Intermedia que requiere evaluación por parte del OAL.")
    
    if puestos_aceptables > 0:
        observaciones.append(f"5. {puestos_aceptables} puestos presentan condición Aceptable.")
    
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
    """Genera un archivo Excel de salida mejorado"""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Hoja 1: Datos de la empresa - Formato mejorado
            empresa_data = []
            for key, value in {
                "Razón Social": datos_empresa["razon_social"],
                "RUT": datos_empresa["rut"],
                "Actividad Económica": datos_empresa["actividad_economica"],
                "Código CIIU": datos_empresa["codigo_ciiu"],
                "Dirección": datos_empresa["direccion"],
                "Comuna": datos_empresa["comuna"],
                "Representante Legal": datos_empresa["representante_legal"],
                "Organismo Administrador": datos_empresa["organismo_administrador"],
                "Fecha Inicio": datos_empresa["fecha_inicio"],
                "Centro de Trabajo": datos_empresa["centro_trabajo"],
                "Trabajadores": f"{datos_empresa['trabajadores_hombres'] + datos_empresa['trabajadores_mujeres']} ({datos_empresa['trabajadores_hombres']} hombres, {datos_empresa['trabajadores_mujeres']} mujeres)"
            }.items():
                empresa_data.append({"Campo": key, "Valor": value})
            
            df_empresa = pd.DataFrame(empresa_data)
            df_empresa.to_excel(writer, sheet_name='Datos Empresa', index=False)
            
            # Hoja 2: Tabla de resumen
            df_resumen.to_excel(writer, sheet_name='Resumen Evaluación', index=False)
            
            # Hoja 3: Observaciones
            df_observaciones = pd.DataFrame({'Observaciones': observaciones})
            df_observaciones.to_excel(writer, sheet_name='Observaciones', index=False)
            
            # Ajustar el libro y las hojas para mejor formato
            workbook = writer.book
            
            # Ajustar columnas en hoja de empresa
            worksheet = writer.sheets['Datos Empresa']
            worksheet.column_dimensions['A'].width = 25
            worksheet.column_dimensions['B'].width = 50
            
            # Ajustar columnas en hoja de resumen
            worksheet = writer.sheets['Resumen Evaluación']
            columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
            widths = [10, 25, 25, 40, 25, 25, 20, 20, 15, 15, 15]
            
            for col, width in zip(columns, widths):
                worksheet.column_dimensions[col].width = width
            
            # Ajustar columna en hoja de observaciones
            worksheet = writer.sheets['Observaciones']
            worksheet.column_dimensions['A'].width = 100
        
        return output.getvalue()
    
    except Exception as e:
        st.error(f"Error al generar Excel de salida: {str(e)}")
        st.error(traceback.format_exc())
        
        # Intentar un enfoque más simple como respaldo
        try:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Versión simplificada
                pd.DataFrame([datos_empresa]).to_excel(writer, sheet_name='Datos Empresa', index=False)
                df_resumen.to_excel(writer, sheet_name='Resumen Evaluación', index=False)
                pd.DataFrame({'Observaciones': observaciones}).to_excel(writer, sheet_name='Observaciones', index=False)
            
            return output.getvalue()
        except:
            return None

def procesar_matriz(archivo_excel):
    """Función principal mejorada que procesa la matriz TMERT"""
    try:
        st.write("Iniciando procesamiento de la matriz TMERT...")
        
        # Cargar el archivo Excel
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        st.write(f"Archivo cargado correctamente. Hojas disponibles: {wb.sheetnames}")
        
        # Extraer datos de empresa (versión mejorada)
        with st.spinner("Extrayendo datos de la empresa..."):
            datos_empresa = extraer_datos_empresa(wb)
        
        # Extraer puestos de trabajo (versión mejorada)
        with st.spinner("Extrayendo puestos de trabajo..."):
            puestos_trabajo = extraer_puestos_trabajo(wb)
            
            if not puestos_trabajo:
                st.error("No se pudieron encontrar puestos de trabajo.")
                return None, None, None, None
        
        # Extraer evaluación inicial (versión mejorada)
        with st.spinner("Extrayendo evaluación inicial..."):
            evaluaciones_iniciales = extraer_evaluacion_inicial(wb, puestos_trabajo)
        
        # Extraer resultados avanzados
        resultados = {}
        with st.spinner("Procesando evaluaciones avanzadas..."):
            for puesto in puestos_trabajo:
                resultados[puesto['numero']] = extraer_resultados_avanzados(wb, puesto, evaluaciones_iniciales)
        
        # Generar tabla de resumen mejorada
        with st.spinner("Generando tabla de resumen..."):
            df_resumen = generar_tabla_resumen(puestos_trabajo, evaluaciones_iniciales, resultados)
        
        # Generar observaciones
        with st.spinner("Generando observaciones..."):
            observaciones = generar_observaciones(df_resumen)
        
        # Generar Excel de salida mejorado
        with st.spinner("Generando archivo Excel de salida..."):
            excel_bytes = generar_excel_salida(datos_empresa, df_resumen, observaciones)
        
        st.success("¡Procesamiento completado con éxito!")
        return datos_empresa, df_resumen, observaciones, excel_bytes
    
    except Exception as e:
        st.error(f"Error al procesar la matriz TMERT: {str(e)}")
        st.error(traceback.format_exc())
        return None, None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    st.title("Procesador de Matriz TMERT")
    
    st.write("""
    Esta aplicación procesa archivos Excel con matrices TMERT (Trastornos Musculoesqueléticos Relacionados al Trabajo)
    y genera un informe estructurado con los resultados de la evaluación.
    """)
    
    # Configuración de la página
    st.sidebar.header("Configuración")
    debug_mode = st.sidebar.checkbox("Modo de depuración", value=False)
    
    # Cargar archivo
    archivo = st.file_uploader("Cargar archivo de Matriz TMERT", type=["xlsx"])
    
    if archivo is not None:
        st.write(f"Archivo cargado: {archivo.name}")
        
        # Botón para procesar
        if st.button("Procesar archivo"):
            with st.spinner("Procesando archivo..."):
                datos_empresa, df_resumen, observaciones, excel_bytes = procesar_matriz(archivo)
            
            if datos_empresa is not None and df_resumen is not None and observaciones is not None:
                st.success("¡Archivo procesado correctamente!")
                
                # Mostrar datos de la empresa
                st.header("Datos de la empresa")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Razón Social:** {datos_empresa['razon_social']}")
                    st.write(f"**RUT:** {datos_empresa['rut']}")
                    st.write(f"**Actividad Económica:** {datos_empresa['actividad_economica']}")
                with col2:
                    st.write(f"**Dirección:** {datos_empresa['direccion']}, {datos_empresa['comuna']}")
                    st.write(f"**Representante Legal:** {datos_empresa['representante_legal']}")
                    st.write(f"**Organismo Administrador:** {datos_empresa['organismo_administrador']}")
                
                # Mostrar tabla de resumen mejorada
                st.header("Resumen de evaluación por puesto de trabajo")
                
                # Convertir los nombres de columnas a nombres más cortos para mejor visualización
                display_df = df_resumen.copy()
                if len(display_df) > 0:
                    # Formatear tabla para mejor visualización
                    st.dataframe(display_df.style.set_properties(
                        **{'text-align': 'left', 'font-size': '12px'}
                    ), use_container_width=True)
                else:
                    st.warning("No se encontraron datos para mostrar en la tabla de resumen.")
                
                # Mostrar observaciones
                st.header("Observaciones")
                for obs in observaciones:
                    st.write(obs)
                
                # Descargar Excel
                if excel_bytes:
                    filename = f"Informe_TMERT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    st.download_button(
                        label="Descargar informe Excel",
                        data=excel_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se pudo generar el archivo Excel para descargar.")
            else:
                st.error("No se han extraído datos correctamente.")
                
                if debug_mode:
                    st.write("Información de depuración:")
                    st.write(f"datos_empresa: {datos_empresa is not None}")
                    st.write(f"df_resumen: {df_resumen is not None}")
                    st.write(f"observaciones: {observaciones is not None}")
                    st.write(f"excel_bytes: {excel_bytes is not None}")

if __name__ == "__main__":
    main()
