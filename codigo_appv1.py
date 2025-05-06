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
    """Extrae los datos de la empresa desde la hoja 1 con validación mejorada"""
    st.write("Extrayendo datos de la empresa...")
    try:
        # Verificar si la hoja 1 existe
        if "1" not in wb.sheetnames:
            st.warning("La hoja '1' no se encuentra en el archivo. Buscando hojas alternativas...")
            # Intentar con nombres alternativos
            alternate_names = ["Datos Empresa", "Empresa", "Informacion"]
            found = False
            for name in alternate_names:
                if name in wb.sheetnames:
                    sheet = wb[name]
                    found = True
                    break
            if not found:
                st.error("No se encontró una hoja con datos de la empresa.")
                return {}
        else:
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
        
        # Patrones para buscar en toda la hoja
        patterns = {
            "razon_social": ["Razón Social", "Razon Social", "Nombre Empresa"],
            "rut": ["RUT", "Rut"],
            "actividad_economica": ["Actividad Económica", "Actividad Economica", "Giro"],
            "codigo_ciiu": ["Código CIIU", "Codigo CIIU"],
            "direccion": ["Dirección", "Direccion"],
            "comuna": ["Comuna"],
            "representante_legal": ["Representante Legal", "Gerente"],
            "organismo_administrador": ["Organismo administrador", "OAL", "Mutual"],
            "fecha_inicio": ["Fecha inicio", "Fecha Inicio"],
            "centro_trabajo": ["centro de trabajo", "sucursal", "sede"],
            "trabajadores": ["trabajadores", "Trabajadores", "N° trabajadores"]
        }
        
        # Buscar en toda la hoja
        max_row = min(100, sheet.max_row)  # Limitar búsqueda a primeras 100 filas
        max_col = min(20, sheet.max_column)  # Limitar búsqueda a primeras 20 columnas
        
        found_data = False
        
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell_value = safe_get_cell_value(sheet, r, c)
                if not cell_value or not isinstance(cell_value, str):
                    continue
                
                # Buscar patrones en el texto de la celda
                for key, keywords in patterns.items():
                    for keyword in keywords:
                        if keyword.lower() in cell_value.lower():
                            # Determinar dónde está el valor (puede ser a la derecha, abajo, etc.)
                            if key == "trabajadores":
                                # Buscar valores de trabajadores por género
                                for nearby_col in range(max(1, c-2), min(max_col, c+5)):
                                    nearby_value = safe_get_cell_value(sheet, r, nearby_col)
                                    if nearby_value and isinstance(nearby_value, str):
                                        if "hombre" in nearby_value.lower():
                                            num_col = nearby_col + 1
                                            hombres_value = safe_get_cell_value(sheet, r, num_col)
                                            if is_numeric(hombres_value):
                                                datos_empresa["trabajadores_hombres"] = int(float(hombres_value))
                                                found_data = True
                                        
                                        if "mujer" in nearby_value.lower():
                                            num_col = nearby_col + 1
                                            mujeres_value = safe_get_cell_value(sheet, r, num_col)
                                            if is_numeric(mujeres_value):
                                                datos_empresa["trabajadores_mujeres"] = int(float(mujeres_value))
                                                found_data = True
                            else:
                                # Buscar el valor en posiciones cercanas (derecha, abajo)
                                # Probar a la derecha
                                right_value = safe_get_cell_value(sheet, r, c+1)
                                if right_value and right_value != "":
                                    datos_empresa[key] = str(right_value)
                                    found_data = True
                                    continue
                                
                                # Probar abajo
                                below_value = safe_get_cell_value(sheet, r+1, c)
                                if below_value and below_value != "":
                                    datos_empresa[key] = str(below_value)
                                    found_data = True
        
        if not found_data:
            st.warning("No se encontraron datos de la empresa. Se procederá con campos vacíos.")
        
        # Si no se encontraron trabajadores, asignar valores por defecto
        if datos_empresa["trabajadores_hombres"] == 0 and datos_empresa["trabajadores_mujeres"] == 0:
            # Buscar explícitamente en la hoja para encontrar los trabajadores
            for r in range(1, max_row + 1):
                row_values = [safe_get_cell_value(sheet, r, c) for c in range(1, max_col + 1)]
                row_text = ' '.join([str(v) for v in row_values if v is not None])
                
                if "trabajadores" in row_text.lower() and "hombres" in row_text.lower() and "mujeres" in row_text.lower():
                    for c in range(1, max_col + 1):
                        cell_text = safe_get_cell_value(sheet, r, c)
                        if cell_text and isinstance(cell_text, str) and "hombre" in cell_text.lower():
                            val = safe_get_cell_value(sheet, r, c+1)
                            if is_numeric(val):
                                datos_empresa["trabajadores_hombres"] = int(float(val))
                        
                        if cell_text and isinstance(cell_text, str) and "mujer" in cell_text.lower():
                            val = safe_get_cell_value(sheet, r, c+1)
                            if is_numeric(val):
                                datos_empresa["trabajadores_mujeres"] = int(float(val))
        
        # Validar que al menos tenemos algunos datos básicos
        if datos_empresa["razon_social"] == "" and datos_empresa["rut"] == "":
            st.warning("No se encontraron datos básicos de la empresa. Se continuará con campos vacíos.")
        
        return datos_empresa
    
    except Exception as e:
        st.error(f"Error al extraer datos de la empresa: {str(e)}")
        st.error(traceback.format_exc())
        # Retornar un diccionario vacío con estructura correcta para evitar errores posteriores
        return {
            "razon_social": "No encontrada",
            "rut": "No encontrado",
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

def extraer_puestos_trabajo(wb):
    """Extrae los puestos de trabajo desde la hoja 2 de forma más robusta"""
    st.write("Extrayendo puestos de trabajo...")
    try:
        # Verificar si la hoja 2 existe
        if "2" not in wb.sheetnames:
            st.warning("La hoja '2' no se encuentra en el archivo. Buscando hojas alternativas...")
            # Intentar con nombres alternativos
            alternate_names = ["Puestos", "Tareas", "Caracterizacion", "Caracterización"]
            found = False
            for name in alternate_names:
                if name in wb.sheetnames:
                    sheet = wb[name]
                    found = True
                    break
            if not found:
                # Buscar una hoja que parezca contener puestos de trabajo
                for name in wb.sheetnames:
                    sheet = wb[name]
                    # Verificar si la hoja tiene encabezados típicos de puestos de trabajo
                    for row in range(1, 20):
                        row_data = [safe_get_cell_value(sheet, row, col) for col in range(1, 10)]
                        row_text = ' '.join([str(x) for x in row_data if x is not None])
                        if "puesto" in row_text.lower() and "trabajo" in row_text.lower():
                            found = True
                            break
                    if found:
                        break
                
                if not found:
                    st.error("No se encontró una hoja con puestos de trabajo.")
                    return []
        else:
            sheet = wb["2"]
        
        # Lista para almacenar los puestos de trabajo
        puestos = []
        
        # Buscar encabezados
        header_row = None
        for row in range(1, 30):
            row_data = [safe_get_cell_value(sheet, row, col) for col in range(1, 15)]
            row_text = ' '.join([str(x) for x in row_data if x is not None and x != ""])
            
            if ("puesto" in row_text.lower() and "trabajo" in row_text.lower()) or "área" in row_text.lower():
                header_row = row
                break
        
        if header_row is None:
            st.warning("No se encontró fila de encabezados. Asumiendo fila 12 como encabezado.")
            header_row = 12
        
        # Determinar columnas de interés
        col_mapping = {}
        max_col = min(25, sheet.max_column)
        
        for col in range(1, max_col + 1):
            header = safe_get_cell_value(sheet, header_row, col)
            if not header:
                continue
            
            header_lower = str(header).lower()
            
            if "n°" in header_lower or "numero" in header_lower or "número" in header_lower:
                col_mapping["numero"] = col
            elif "área" in header_lower or "area" in header_lower:
                col_mapping["area"] = col
            elif "puesto" in header_lower:
                col_mapping["puesto"] = col
            elif "tarea" in header_lower:
                col_mapping["tareas"] = col
            elif "descrip" in header_lower:
                col_mapping["descripcion"] = col
            elif "horario" in header_lower:
                col_mapping["horario"] = col
            
            # Buscar columnas de trabajadores
            if "hombre" in header_lower:
                col_mapping["trabajadores_hombres"] = col
            if "mujer" in header_lower:
                col_mapping["trabajadores_mujeres"] = col
        
        # Si no encontramos algunas columnas importantes, asumimos posiciones estándar
        if "numero" not in col_mapping:
            col_mapping["numero"] = 1
        if "area" not in col_mapping:
            col_mapping["area"] = 2
        if "puesto" not in col_mapping:
            col_mapping["puesto"] = 3
        if "tareas" not in col_mapping and "descripcion" not in col_mapping:
            col_mapping["tareas"] = 4
        
        # Buscar filas con información de puestos de trabajo
        max_row = min(100, sheet.max_row)
        start_row = header_row + 1
        
        found_puestos = False
        numero_counter = 1  # Contador para asignar números automáticamente si no hay
        
        for row in range(start_row, max_row + 1):
            # Verificar si es una fila con datos de puesto de trabajo
            numero = safe_get_cell_value(sheet, row, col_mapping.get("numero", 1))
            
            # Si no hay número pero hay datos en otras celdas, asignar número automático
            if (not numero or not is_numeric(numero)) and any(safe_get_cell_value(sheet, row, col) for col in range(2, 5)):
                numero = numero_counter
                numero_counter += 1
            elif is_numeric(numero):
                numero = int(float(numero))
            else:
                continue  # No es una fila de puesto de trabajo
            
            area = safe_get_cell_value(sheet, row, col_mapping.get("area", 2)) or ""
            puesto = safe_get_cell_value(sheet, row, col_mapping.get("puesto", 3)) or ""
            tareas = safe_get_cell_value(sheet, row, col_mapping.get("tareas", 4)) or ""
            
            # Si no hay información básica, ignorar la fila
            if not (area or puesto or tareas):
                continue
            
            descripcion = safe_get_cell_value(sheet, row, col_mapping.get("descripcion", 5)) or ""
            horario = safe_get_cell_value(sheet, row, col_mapping.get("horario", 6)) or ""
            
            # Trabajadores
            trabajadores_hombres = 0
            if "trabajadores_hombres" in col_mapping:
                val = safe_get_cell_value(sheet, row, col_mapping["trabajadores_hombres"])
                if is_numeric(val):
                    trabajadores_hombres = int(float(val))
            
            trabajadores_mujeres = 0
            if "trabajadores_mujeres" in col_mapping:
                val = safe_get_cell_value(sheet, row, col_mapping["trabajadores_mujeres"])
                if is_numeric(val):
                    trabajadores_mujeres = int(float(val))
            
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
            
            found_puestos = True
        
        if not found_puestos:
            st.warning("No se encontraron puestos de trabajo en la hoja. Verificando toda la hoja...")
            
            # Intentar una búsqueda más agresiva en toda la hoja
            for row in range(1, max_row + 1):
                cell_value = safe_get_cell_value(sheet, row, 1)
                if is_numeric(cell_value):
                    numero = int(float(cell_value))
                    area = safe_get_cell_value(sheet, row, 2) or ""
                    puesto = safe_get_cell_value(sheet, row, 3) or ""
                    tareas = safe_get_cell_value(sheet, row, 4) or ""
                    
                    if area or puesto or tareas:
                        puestos.append({
                            "numero": numero,
                            "area": area,
                            "puesto": puesto,
                            "tareas": tareas,
                            "descripcion": "",
                            "horario": "",
                            "trabajadores_hombres": 0,
                            "trabajadores_mujeres": 0
                        })
                        found_puestos = True
        
        if not found_puestos:
            st.error("No se pudieron encontrar puestos de trabajo en ninguna hoja.")
        
        return puestos
    
    except Exception as e:
        st.error(f"Error al extraer puestos de trabajo: {str(e)}")
        st.error(traceback.format_exc())
        return []

def extraer_evaluacion_inicial(wb, puestos_trabajo):
    """Extrae la evaluación inicial desde la hoja 3"""
    st.write("Extrayendo evaluación inicial...")
    try:
        # Verificar si la hoja 3 existe
        if "3" not in wb.sheetnames:
            st.warning("La hoja '3' no se encuentra en el archivo. Buscando hojas alternativas...")
            # Intentar con nombres alternativos
            alternate_names = ["Identificación inicial", "Identificacion inicial", "Evaluación inicial"]
            found = False
            for name in alternate_names:
                if name in wb.sheetnames:
                    sheet = wb[name]
                    found = True
                    break
            if not found:
                # Buscar una hoja que parezca contener evaluaciones iniciales
                for name in wb.sheetnames:
                    sheet = wb[name]
                    # Verificar si la hoja tiene encabezados típicos de evaluación inicial
                    for row in range(1, 20):
                        row_data = [safe_get_cell_value(sheet, row, col) for col in range(1, 10)]
                        row_text = ' '.join([str(x) for x in row_data if x is not None])
                        if "trabajo repetitivo" in row_text.lower() or "postura estática" in row_text.lower():
                            found = True
                            break
                    if found:
                        break
                
                if not found:
                    st.error("No se encontró una hoja con evaluación inicial.")
                    # Crear evaluaciones iniciales vacías basadas en los puestos encontrados
                    evaluaciones = []
                    for puesto in puestos_trabajo:
                        evaluaciones.append({
                            "numero": puesto["numero"],
                            "puesto": puesto["puesto"],
                            "tareas": puesto["tareas"],
                            "trabajo_repetitivo": False,
                            "postura_estatica": False,
                            "mmc_ldt": False,
                            "mmc_ea": False,
                            "mmp": False,
                            "vibracion_cc": False,
                            "vibracion_mb": False,
                            "resultado": "Sin evaluación"
                        })
                    return evaluaciones
        else:
            sheet = wb["3"]
        
        # Lista para almacenar las evaluaciones iniciales
        evaluaciones = []
        
        # Buscar encabezados
        header_row = None
        for row in range(1, 30):
            row_data = [safe_get_cell_value(sheet, row, col) for col in range(1, 15)]
            row_text = ' '.join([str(x) for x in row_data if x is not None])
            
            if "trabajo repetitivo" in row_text.lower() or "postura" in row_text.lower():
                header_row = row
                break
        
        if header_row is None:
            st.warning("No se encontró fila de encabezados en evaluación inicial. Asumiendo fila 12.")
            header_row = 12
        
        # Determinar columnas de interés
        col_mapping = {}
        max_col = min(15, sheet.max_column)
        
        for col in range(1, max_col + 1):
            header = safe_get_cell_value(sheet, header_row, col)
            if not header:
                continue
            
            header_lower = str(header).lower()
            
            if "n°" in header_lower or "numero" in header_lower or "número" in header_lower:
                col_mapping["numero"] = col
            elif "puesto" in header_lower:
                col_mapping["puesto"] = col
            elif "tarea" in header_lower:
                col_mapping["tareas"] = col
            elif "trabajo repetitivo" in header_lower:
                col_mapping["trabajo_repetitivo"] = col
            elif "postura" in header_lower and "estática" in header_lower:
                col_mapping["postura_estatica"] = col
            elif "manipulación" in header_lower or "levantamiento" in header_lower or "descenso" in header_lower:
                col_mapping["mmc_ldt"] = col
            elif "empuje" in header_lower or "arrastre" in header_lower:
                col_mapping["mmc_ea"] = col
            elif "paciente" in header_lower or "personas" in header_lower:
                col_mapping["mmp"] = col
            elif "cuerpo completo" in header_lower:
                col_mapping["vibracion_cc"] = col
            elif "mano" in header_lower and "brazo" in header_lower:
                col_mapping["vibracion_mb"] = col
            elif "resultado" in header_lower:
                col_mapping["resultado"] = col
        
        # Si no encontramos algunas columnas importantes, asumimos posiciones estándar
        if "numero" not in col_mapping:
            col_mapping["numero"] = 1
        if "puesto" not in col_mapping:
            col_mapping["puesto"] = 2
        if "tareas" not in col_mapping:
            col_mapping["tareas"] = 3
        
        # Asegurar que tenemos todas las columnas de factores de riesgo
        risk_factors = [
            "trabajo_repetitivo", "postura_estatica", "mmc_ldt", "mmc_ea", 
            "mmp", "vibracion_cc", "vibracion_mb"
        ]
        
        # Asignar columnas secuencialmente para factores faltantes
        next_col = 4
        for factor in risk_factors:
            if factor not in col_mapping:
                while next_col in [col_mapping.get(k) for k in col_mapping]:
                    next_col += 1
                col_mapping[factor] = next_col
                next_col += 1
        
        # Buscar filas con información de evaluación inicial
        max_row = min(100, sheet.max_row)
        start_row = header_row + 1
        
        found_evaluaciones = False
        
        for row in range(start_row, max_row + 1):
            # Verificar si es una fila con datos de evaluación
            numero = safe_get_cell_value(sheet, row, col_mapping.get("numero", 1))
            
            if not numero or not is_numeric(numero):
                continue
            
            numero = int(float(numero))
            
            puesto = safe_get_cell_value(sheet, row, col_mapping.get("puesto", 2)) or ""
            tareas = safe_get_cell_value(sheet, row, col_mapping.get("tareas", 3)) or ""
            
            # Factores de riesgo
            trabajo_repetitivo = safe_get_cell_value(sheet, row, col_mapping.get("trabajo_repetitivo", 4))
            trabajo_repetitivo = str(trabajo_repetitivo).upper() == "SI" if trabajo_repetitivo else False
            
            postura_estatica = safe_get_cell_value(sheet, row, col_mapping.get("postura_estatica", 5))
            postura_estatica = str(postura_estatica).upper() == "SI" if postura_estatica else False
            
            mmc_ldt = safe_get_cell_value(sheet, row, col_mapping.get("mmc_ldt", 6))
            mmc_ldt = str(mmc_ldt).upper() == "SI" if mmc_ldt else False
            
            mmc_ea = safe_get_cell_value(sheet, row, col_mapping.get("mmc_ea", 7))
            mmc_ea = str(mmc_ea).upper() == "SI" if mmc_ea else False
            
            mmp = safe_get_cell_value(sheet, row, col_mapping.get("mmp", 8))
            mmp = str(mmp).upper() == "SI" if mmp else False
            
            vibracion_cc = safe_get_cell_value(sheet, row, col_mapping.get("vibracion_cc", 9))
            vibracion_cc = str(vibracion_cc).upper() == "SI" if vibracion_cc else False
            
            vibracion_mb = safe_get_cell_value(sheet, row, col_mapping.get("vibracion_mb", 10))
            vibracion_mb = str(vibracion_mb).upper() == "SI" if vibracion_mb else False
            
            # Resultado
            resultado = ""
            if "resultado" in col_mapping:
                resultado = safe_get_cell_value(sheet, row, col_mapping["resultado"]) or ""
            
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
            
            found_evaluaciones = True
        
        if not found_evaluaciones:
            st.warning("No se encontraron evaluaciones iniciales. Generando evaluaciones basadas en los puestos...")
            
            # Crear evaluaciones iniciales basadas en los puestos encontrados
            for puesto in puestos_trabajo:
                evaluaciones.append({
                    "numero": puesto["numero"],
                    "puesto": puesto["puesto"],
                    "tareas": puesto["tareas"],
                    "trabajo_repetitivo": False,
                    "postura_estatica": False,
                    "mmc_ldt": False,
                    "mmc_ea": False,
                    "mmp": False,
                    "vibracion_cc": False,
                    "vibracion_mb": False,
                    "resultado": "Sin evaluación"
                })
        
        return evaluaciones
    
    except Exception as e:
        st.error(f"Error al extraer evaluación inicial: {str(e)}")
        st.error(traceback.format_exc())
        
        # Crear evaluaciones iniciales vacías basadas en los puestos encontrados
        evaluaciones = []
        for puesto in puestos_trabajo:
            evaluaciones.append({
                "numero": puesto["numero"],
                "puesto": puesto["puesto"],
                "tareas": puesto["tareas"],
                "trabajo_repetitivo": False,
                "postura_estatica": False,
                "mmc_ldt": False,
                "mmc_ea": False,
                "mmp": False,
                "vibracion_cc": False,
                "vibracion_mb": False,
                "resultado": "Sin evaluación"
            })
        return evaluaciones

def buscar_resultado_evaluacion(wb, hoja, numero_puesto):
    """Busca el resultado de la evaluación avanzada para un puesto en una hoja específica"""
    try:
        if hoja not in wb.sheetnames:
            return "Hoja no encontrada"
        
        sheet = wb[hoja]
        
        # Buscar fila con el número de puesto
        max_row = min(100, sheet.max_row)
        
        # Patrones de resultados
        patrones_resultado = [
            (r"intermedio.*solicitar\s+evaluaci[oó]n", "Intermedio - Solicitar evaluación OAL"),
            (r"ausencia\s+total\s+del\s+riesgo", "Ausencia total del riesgo"),
            (r"aceptable", "Aceptable"),
            (r"no\s+aceptable", "No Aceptable"),
            (r"cr[ií]tico", "Crítico"),
            (r"no\s+cr[ií]tico", "No crítico")
        ]
        
        for row in range(15, max_row + 1):
            cell_value = safe_get_cell_value(sheet, row, 1)
            
            if is_numeric(cell_value) and int(float(cell_value)) == numero_puesto:
                # Buscar el resultado en la fila
                row_values = [safe_get_cell_value(sheet, row, col) for col in range(1, 50)]
                row_text = ' '.join([str(v) for v in row_values if v is not None])
                
                # Verificar si hay datos relevantes en la fila
                has_data = False
                for val in row_values:
                    if val is not None and val != "":
                        has_data = True
                        break
                
                if not has_data:
                    return "No contestada por usuario, revisar"
                
                # Buscar patrones específicos en el texto de la fila
                for patron, resultado in patrones_resultado:
                    if re.search(patron, row_text, re.IGNORECASE):
                        return resultado
                
                # Si no encontramos un patrón específico, buscar celdas con resultados
                for col in range(5, 40):
                    val = safe_get_cell_value(sheet, row, col)
                    if isinstance(val, str):
                        if "Intermedio" in val and "Solicitar evaluación" in val:
                            return "Intermedio - Solicitar evaluación OAL"
                        elif "Ausencia total del riesgo" in val:
                            return "Ausencia total del riesgo"
                        elif val == "Aceptable":
                            return "Aceptable"
                        elif val == "No aceptable" or val == "No Aceptable":
                            return "No Aceptable"
                        elif "Crítico" in val or "Critico" in val:
                            return "Crítico"
                        elif "No crítico" in val or "No Crítico" in val or "No critico" in val:
                            return "No crítico"
                
                # Verificar si las celdas importantes están vacías
                celdas_vacias = True
                for col in range(8, 30):
                    val = safe_get_cell_value(sheet, row, col)
                    if isinstance(val, str) and len(val) > 1:
                        celdas_vacias = False
                        break
                
                if celdas_vacias:
                    return "No contestada por usuario, revisar"
                
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
    try:
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        
        # Formatear el diccionario de empresa para DataFrame
        empresa_data = []
        for key, value in datos_empresa.items():
            if key == "trabajadores_hombres" or key == "trabajadores_mujeres":
                continue
            empresa_data.append({"Campo": key.replace("_", " ").capitalize(), "Valor": value})
        
        # Agregar trabajadores como una fila adicional
        trabajadores_total = datos_empresa.get("trabajadores_hombres", 0) + datos_empresa.get("trabajadores_mujeres", 0)
        empresa_data.append({
            "Campo": "Trabajadores", 
            "Valor": f"Total: {trabajadores_total} ({datos_empresa.get('trabajadores_hombres', 0)} hombres, {datos_empresa.get('trabajadores_mujeres', 0)} mujeres)"
        })
        
        # Hoja 1: Datos de la empresa
        df_empresa = pd.DataFrame(empresa_data)
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
    
    except Exception as e:
        st.error(f"Error al generar Excel de salida: {str(e)}")
        st.error(traceback.format_exc())
        
        # Intentar con un enfoque más simple
        try:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Hoja 1: Datos de la empresa
                pd.DataFrame([datos_empresa]).to_excel(writer, sheet_name='Datos Empresa', index=False)
                
                # Hoja 2: Tabla de resumen
                df_resumen.to_excel(writer, sheet_name='Resumen Evaluación', index=False)
                
                # Hoja 3: Observaciones
                pd.DataFrame({'Observaciones': observaciones}).to_excel(writer, sheet_name='Observaciones', index=False)
            
            return output.getvalue()
            
        except Exception as e2:
            st.error(f"Error al generar Excel simplificado: {str(e2)}")
            return None

def procesar_matriz(archivo_excel):
    """Función principal que procesa la matriz TMERT"""
    try:
        st.write("Iniciando procesamiento de la matriz TMERT...")
        
        # Cargar el archivo Excel con manejo de errores
        try:
            wb = openpyxl.load_workbook(archivo_excel, data_only=True)
            st.write(f"Archivo cargado correctamente. Hojas disponibles: {wb.sheetnames}")
        except Exception as e:
            st.error(f"Error al cargar el archivo Excel: {str(e)}")
            st.error(traceback.format_exc())
            return None, None, None, None
        
        # Extraer datos con manejo de progreso
        with st.spinner("Extrayendo datos de la empresa..."):
            datos_empresa = extraer_datos_empresa(wb)
            if not datos_empresa:
                st.warning("No se pudieron extraer datos de la empresa. Continuando con el proceso...")
        
        with st.spinner("Extrayendo puestos de trabajo..."):
            puestos_trabajo = extraer_puestos_trabajo(wb)
            if not puestos_trabajo:
                st.error("No se pudieron extraer puestos de trabajo. No se puede continuar.")
                return None, None, None, None
            
            st.write(f"Se encontraron {len(puestos_trabajo)} puestos de trabajo.")
        
        with st.spinner("Extrayendo evaluación inicial..."):
            evaluaciones_iniciales = extraer_evaluacion_inicial(wb, puestos_trabajo)
            if not evaluaciones_iniciales:
                st.warning("No se pudo extraer la evaluación inicial. Se usarán evaluaciones vacías.")
                # Crear evaluaciones iniciales basadas en los puestos
                evaluaciones_iniciales = []
                for puesto in puestos_trabajo:
                    evaluaciones_iniciales.append({
                        "numero": puesto["numero"],
                        "puesto": puesto["puesto"],
                        "tareas": puesto["tareas"],
                        "trabajo_repetitivo": False,
                        "postura_estatica": False,
                        "mmc_ldt": False,
                        "mmc_ea": False,
                        "mmp": False,
                        "vibracion_cc": False,
                        "vibracion_mb": False,
                        "resultado": "Sin evaluación"
                    })
        
        # Extraer resultados avanzados con barra de progreso
        resultados = {}
        progress_bar = st.progress(0)
        for i, puesto in enumerate(puestos_trabajo):
            progress_bar.progress((i + 1) / len(puestos_trabajo))
            st.write(f"Procesando puesto {i+1} de {len(puestos_trabajo)}: {puesto['puesto']}")
            resultados[puesto["numero"]] = extraer_resultados_avanzados(wb, puesto, evaluaciones_iniciales)
        
        progress_bar.empty()
        
        # Generar tabla de resumen
        with st.spinner("Generando tabla de resumen..."):
            df_resumen = generar_tabla_resumen(puestos_trabajo, evaluaciones_iniciales, resultados)
            
            if df_resumen.empty:
                st.error("No se pudo generar la tabla de resumen.")
                return None, None, None, None
        
        # Generar observaciones
        with st.spinner("Generando observaciones..."):
            observaciones = generar_observaciones(df_resumen)
        
        # Generar Excel de salida
        with st.spinner("Generando archivo Excel de salida..."):
            excel_bytes = generar_excel_salida(datos_empresa, df_resumen, observaciones)
            
            if excel_bytes is None:
                st.error("No se pudo generar el archivo Excel de salida.")
                return datos_empresa, df_resumen, observaciones, None
        
        st.success("Procesamiento completado con éxito!")
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
    
    # Cargar archivo
    archivo = st.file_uploader("Cargar archivo de Matriz TMERT", type=["xlsx"])
    
    if archivo is not None:
        st.write(f"Archivo cargado: {archivo.name}")
        
        # Botón para procesar
        if st.button("Procesar archivo"):
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
                if excel_bytes:
                    razon_social = datos_empresa['razon_social'].replace(" ", "_") if datos_empresa['razon_social'] else "empresa"
                    st.download_button(
                        label="Descargar informe Excel",
                        data=excel_bytes,
                        file_name=f"Informe_TMERT_{razon_social}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se pudo generar el archivo Excel para descargar.")
            else:
                st.error("No se han extraído datos correctamente.")

if __name__ == "__main__":
    main()
