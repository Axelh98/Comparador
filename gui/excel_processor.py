"""
Módulo para el procesamiento y comparación de archivos Excel
"""

import os
import pandas as pd
from utils.data_utils import DataUtils
from utils.file_utils import FileUtils


class ExcelProcessor:
    """Clase para procesar y comparar archivos Excel"""
    
    def __init__(self):
        self.data_utils = DataUtils()
        self.file_utils = FileUtils()
    
    def cargar_datos(self, archivo, sheet_anterior, sheet_actual):
        """Carga los datos desde las hojas especificadas."""
        try:
            df_anterior = pd.read_excel(archivo, sheet_name=sheet_anterior)
            df_actual = pd.read_excel(archivo, sheet_name=sheet_actual)
            return df_anterior, df_actual
        except Exception as e:
            raise Exception(f"Error al cargar datos: {e}")

    def tabla_comparacion_completa(self, df1, df2):
        """Crea una tabla que muestra los cambios entre dos DataFrames."""
        df1, df2 = self.data_utils.normalizar_df(df1), self.data_utils.normalizar_df(df2)
        comparacion = df1.where(df1 == df2, df1 + " ➔ " + df2)
        return comparacion

    def crear_reemplazo(self, df1, df2):
        """Crea un DataFrame que reemplaza valores antiguos con nuevos."""
        df1, df2 = self.data_utils.normalizar_df(df1), self.data_utils.normalizar_df(df2)
        return df1.where(df1 == df2, df2)

    def detectar_nuevas_filas(self, df1, df2, id_column):
        """Identifica registros que existen en df2 pero no en df1."""
        return df2.loc[~df2[id_column].isin(df1[id_column])]

    def detectar_filas_eliminadas(self, df1, df2, id_column):
        """Identifica registros que existen en df1 pero no en df2."""
        return df1.loc[~df1[id_column].isin(df2[id_column])]

    def crear_historial_de_cambios(self, df1, df2, mes, id_column, log_func=None):
        """
        Genera un historial detallado de todos los cambios entre dos DataFrames,
        teniendo en cuenta modificaciones, añadidos y eliminaciones.
        """
        historial = []
        
        # Nos aseguramos que el id_column exista en ambos DataFrames
        if id_column not in df1.columns:
            if log_func:
                log_func(f"Advertencia: La columna '{id_column}' no existe en el DataFrame anterior.")
            return historial
        if id_column not in df2.columns:
            if log_func:
                log_func(f"Advertencia: La columna '{id_column}' no existe en el DataFrame actual.")
            return historial

        # Detectar modificaciones en registros existentes en ambos DataFrames
        common_indices = set(df1.index).intersection(set(df2.index))
        for idx in common_indices:
            try:
                facility_id = str(df1.at[idx, id_column])
                for col in df1.columns:
                    if col in df2.columns:  # Solo comparamos columnas que existen en ambos
                        val1, val2 = df1.at[idx, col], df2.at[idx, col]
                        # Convertimos a string para comparación segura (evita errores de tipo)
                        str_val1 = str(val1) if not pd.isna(val1) else ""
                        str_val2 = str(val2) if not pd.isna(val2) else ""
                        
                        if str_val1 != str_val2:
                            historial.append({
                                id_column: facility_id,
                                "Fila (índice)": idx,
                                "Columna": col,
                                "Valor Anterior": val1,
                                "Valor Nuevo": val2,
                                "Mes": mes,
                                "Tipo de Cambio": "Modificación"
                            })
            except Exception as e:
                if log_func:
                    log_func(f"Error al procesar modificación en índice {idx}: {e}")

        # Detectar nuevos registros
        nuevas_filas = self.detectar_nuevas_filas(df1, df2, id_column)
        for idx in nuevas_filas.index:
            try:
                facility_id = str(df2.at[idx, id_column])
                historial.append({
                    id_column: facility_id,
                    "Fila (índice)": idx,
                    "Columna": "(Nuevo Registro)",
                    "Valor Anterior": "(Nuevo Registro)",
                    "Valor Nuevo": "(Nuevo Registro)",
                    "Mes": mes,
                    "Tipo de Cambio": "Nuevo Registro"
                })
            except Exception as e:
                if log_func:
                    log_func(f"Error al procesar nuevo registro en índice {idx}: {e}")

        # Detectar registros eliminados
        filas_eliminadas = self.detectar_filas_eliminadas(df1, df2, id_column)
        for idx in filas_eliminadas.index:
            try:
                facility_id = str(df1.at[idx, id_column])
                historial.append({
                    id_column: facility_id,
                    "Fila (índice)": idx,
                    "Columna": "(Registro Eliminado)",
                    "Valor Anterior": "(Registro Eliminado)",
                    "Valor Nuevo": "(Registro Eliminado)",
                    "Mes": mes,
                    "Tipo de Cambio": "Registro Eliminado"
                })
            except Exception as e:
                if log_func:
                    log_func(f"Error al procesar registro eliminado en índice {idx}: {e}")

        return historial

    def crear_resumen_informe(self, historial_df, mes):
        """Crea un resumen de los cambios detectados."""
        resumen = {
            "Mes": [mes],
            "Total de Cambios": [len(historial_df)],
            "Modificaciones": [sum(historial_df["Tipo de Cambio"] == "Modificación")],
            "Nuevos Registros": [sum(historial_df["Tipo de Cambio"] == "Nuevo Registro")],
            "Registros Eliminados": [sum(historial_df["Tipo de Cambio"] == "Registro Eliminado")]
        }
        return pd.DataFrame(resumen)

    def procesar_comparacion(self, config, log_func=None):
        """Función principal que coordina todo el proceso de comparación."""
        try:
            # Extraer configuración
            archivo = config["archivo_base"]
            mes = config["mes_actual"]
            id_column = config["id_column"]
            titulo = config["titulo"]
            
            if log_func:
                log_func(f"Cargando datos de las hojas '{config['sheet_anterior']}' y '{config['sheet_actual']}'...")
            
            # Carga y prepara los datos
            df1, df2 = self.cargar_datos(archivo, config["sheet_anterior"], config["sheet_actual"])
            
            if log_func:
                log_func(f"Alineando DataFrames... {df1.shape} y {df2.shape}")
            df1, df2 = self.data_utils.alinear_dataframes(df1, df2)

            # Normaliza los datos para comparación
            if log_func:
                log_func("Normalizando datos para comparación...")
            df1_norm = self.data_utils.normalizar_df(df1)
            df2_norm = self.data_utils.normalizar_df(df2)

            # Realiza la comparación
            if log_func:
                log_func("Generando tabla de comparación completa...")
            comparacion_completa = self.tabla_comparacion_completa(df1, df2)
            mask_cambios = (df1_norm != df2_norm) & ~(df1_norm.eq("") & df2_norm.eq(""))

            # Genera el historial de cambios
            if log_func:
                log_func("Creando historial de cambios...")
            historial = self.crear_historial_de_cambios(df1, df2, mes, id_column, log_func)
            historial_df = pd.DataFrame(historial)
            
            # Rutas de archivos
            historial_archivo = f"{titulo} - Historial de Cambios.xlsx"
            nombre_salida = f"{titulo}-procedimiento-{mes}.xlsx"
            nombre_datos_actualizados = f"{titulo} - Datos Actualizados.xlsx"

            # Actualiza o crea el historial
            if log_func:
                log_func(f"Guardando historial en '{historial_archivo}'...")
            if os.path.exists(historial_archivo):
                historial_existente = pd.read_excel(historial_archivo)
                historial_total = pd.concat([historial_existente, historial_df], ignore_index=True)
            else:
                historial_total = historial_df

            # Guarda el historial
            historial_total.to_excel(historial_archivo, index=False)

            # Crea el resumen de cambios
            if log_func:
                log_func("Generando resumen de cambios...")
            resumen = self.crear_resumen_informe(historial_df, mes)

            # Agrega el resumen al historial
            with pd.ExcelWriter(historial_archivo, engine='openpyxl', mode='a') as writer:
                resumen.to_excel(writer, sheet_name='Resumen Informe', index=False)

            # Crea el DataFrame con los datos actualizados
            if log_func:
                log_func("Creando DataFrame actualizado...")
            reemplazo = self.crear_reemplazo(df1, df2)

            # Guarda la comparación completa y los cambios
            if log_func:
                log_func(f"Guardando comparación completa en '{nombre_salida}'...")
            with pd.ExcelWriter(nombre_salida, engine="xlsxwriter") as writer:
                self.file_utils.escribir_formateado(writer, "Comparación Completa", comparacion_completa, mask_cambios, resaltar=True)
                self.file_utils.escribir_formateado(writer, "Solo Cambios", comparacion_completa[mask_cambios.any(axis=1)], mask_cambios, resaltar=True)

            # Guarda los datos actualizados
            if log_func:
                log_func(f"Guardando datos actualizados en '{nombre_datos_actualizados}'...")
            reemplazo_df = pd.DataFrame(reemplazo)
            reemplazo_df.to_excel(nombre_datos_actualizados, index=False)
            
            if log_func:
                log_func("Proceso completado exitosamente.")
                log_func(f"Se generaron los siguientes archivos:")
                log_func(f"- {historial_archivo}")
                log_func(f"- {nombre_salida}")
                log_func(f"- {nombre_datos_actualizados}")
            
            return resumen
            
        except Exception as e:
            if log_func:
                log_func(f"Error durante el procesamiento: {e}")
            raise