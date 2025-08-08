"""
Módulo para el procesamiento y comparación de archivos Excel
"""

import os
import pandas as pd
from utils.data_utils import DataUtils
from utils.file_utils import FileUtils
from excel_charts import generar_graficos_historial

# Constantes para tipos de cambio
MODIFICACION = "Modificación"
NUEVO_REGISTRO = "Nuevo Registro"
NUEVO_REGISTRO_CAMPO = "Nuevo Registro - Campo"
REGISTRO_ELIMINADO = "Registro Eliminado"


class ExcelProcessor:
    """Clase para procesar y comparar archivos Excel"""
    
    def __init__(self):
        self.data_utils = DataUtils()
        self.file_utils = FileUtils()
    
    def cargar_datos(self, archivo: str, sheet_anterior: str, sheet_actual: str) -> tuple[pd.DataFrame, pd.DataFrame]:
        """Carga y normaliza datos desde las hojas especificadas en un archivo Excel."""
        try:
            hojas = pd.read_excel(archivo, sheet_name=[sheet_anterior, sheet_actual])
            df_anterior = self.data_utils.normalizar_df(hojas[sheet_anterior])
            df_actual = self.data_utils.normalizar_df(hojas[sheet_actual])
            return df_anterior, df_actual
        except Exception as e:
            raise Exception(f"Error al cargar datos: {e}")

    def tabla_comparacion_completa(self, df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        """Crea una tabla que muestra los cambios entre dos DataFrames."""
        df1_str, df2_str = df1.astype(str), df2.astype(str)
        return df1.where(df1 == df2, df1_str + " ➔ " + df2_str)

    def crear_reemplazo(self, df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        """Crea un DataFrame que reemplaza valores antiguos con nuevos."""
        return df1.where(df1 == df2, df2)

    def _filas_por_diferencia_ids(self, df1: pd.DataFrame, df2: pd.DataFrame, id_column: str, modo: str) -> pd.DataFrame:
        """
        Devuelve filas que están en un DataFrame pero no en el otro.
        modo = "nuevas" -> en df2 pero no en df1
        modo = "eliminadas" -> en df1 pero no en df2
        """
        if id_column not in df1.columns or id_column not in df2.columns:
            return pd.DataFrame()

        df1_clean = df1.dropna(subset=[id_column])
        df2_clean = df2.dropna(subset=[id_column])

        ids_df1 = set(df1_clean[id_column].astype(str))
        ids_df2 = set(df2_clean[id_column].astype(str))

        if modo == "nuevas":
            target_ids = ids_df2 - ids_df1
            return df2_clean[df2_clean[id_column].astype(str).isin(target_ids)]
        elif modo == "eliminadas":
            target_ids = ids_df1 - ids_df2
            return df1_clean[df1_clean[id_column].astype(str).isin(target_ids)]
        else:
            return pd.DataFrame()

    def detectar_nuevas_filas(self, df1: pd.DataFrame, df2: pd.DataFrame, id_column: str) -> pd.DataFrame:
        return self._filas_por_diferencia_ids(df1, df2, id_column, "nuevas")

    def detectar_filas_eliminadas(self, df1: pd.DataFrame, df2: pd.DataFrame, id_column: str) -> pd.DataFrame:
        return self._filas_por_diferencia_ids(df1, df2, id_column, "eliminadas")

    def crear_historial_de_cambios(self, df1: pd.DataFrame, df2: pd.DataFrame, mes: str, id_column: str, log_func=None) -> list[dict]:
        """
        Genera un historial detallado de todos los cambios entre dos DataFrames,
        teniendo en cuenta modificaciones, añadidos y eliminaciones.
        """
        historial = []

        if id_column not in df1.columns or id_column not in df2.columns:
            log_func and log_func(f"Advertencia: La columna '{id_column}' no existe en ambos DataFrames.")
            return historial

        df1_clean = df1.dropna(subset=[id_column]).copy()
        df2_clean = df2.dropna(subset=[id_column]).copy()

        df1_clean[id_column] = df1_clean[id_column].astype(str)
        df2_clean[id_column] = df2_clean[id_column].astype(str)

        dict_df1 = df1_clean.set_index(id_column).to_dict('index')
        dict_df2 = df2_clean.set_index(id_column).to_dict('index')

        ids_df1 = set(dict_df1.keys())
        ids_df2 = set(dict_df2.keys())

        # Modificaciones
        ids_comunes = ids_df1 & ids_df2
        log_func and log_func(f"Analizando {len(ids_comunes)} registros comunes para modificaciones...")

        for facility_id in ids_comunes:
            registro_anterior = dict_df1[facility_id]
            registro_actual = dict_df2[facility_id]
            for col in df1_clean.columns:
                if col in df2_clean.columns and col != id_column:
                    val1, val2 = registro_anterior.get(col, ""), registro_actual.get(col, "")
                    str_val1, str_val2 = ("" if pd.isna(val1) else str(val1)), ("" if pd.isna(val2) else str(val2))
                    if str_val1 != str_val2:
                        historial.append({
                            id_column: facility_id,
                            "Fila (índice)": f"ID: {facility_id}",
                            "Columna": col,
                            "Valor Anterior": val1,
                            "Valor Nuevo": val2,
                            "Mes": mes,
                            "Tipo de Cambio": MODIFICACION
                        })

        # Nuevos registros
        nuevos_ids = ids_df2 - ids_df1
        log_func and log_func(f"Detectados {len(nuevos_ids)} nuevos registros...")
        for facility_id in nuevos_ids:
            registro_nuevo = dict_df2[facility_id]
            historial.append({
                id_column: facility_id,
                "Fila (índice)": f"ID: {facility_id}",
                "Columna": "Nuevo Registro Completo",
                "Valor Anterior": "N/A",
                "Valor Nuevo": f"Nuevo registro con {len(registro_nuevo)} campos",
                "Mes": mes,
                "Tipo de Cambio": NUEVO_REGISTRO
            })
            for col, val in registro_nuevo.items():
                if col != id_column:
                    historial.append({
                        id_column: facility_id,
                        "Fila (índice)": f"ID: {facility_id}",
                        "Columna": col,
                        "Valor Anterior": "N/A",
                        "Valor Nuevo": val,
                        "Mes": mes,
                        "Tipo de Cambio": NUEVO_REGISTRO_CAMPO
                    })

        # Registros eliminados
        eliminados_ids = ids_df1 - ids_df2
        log_func and log_func(f"Detectados {len(eliminados_ids)} registros eliminados...")
        for facility_id in eliminados_ids:
            registro_eliminado = dict_df1[facility_id]
            historial.append({
                id_column: facility_id,
                "Fila (índice)": f"ID: {facility_id}",
                "Columna": "Registro Eliminado Completo",
                "Valor Anterior": f"Registro eliminado con {len(registro_eliminado)} campos",
                "Valor Nuevo": "N/A",
                "Mes": mes,
                "Tipo de Cambio": REGISTRO_ELIMINADO
            })

        log_func and log_func(f"Historial completado: {len(historial)} cambios detectados en total")
        return historial

    def obtener_resumen_nuevas_filas(self, df1: pd.DataFrame, df2: pd.DataFrame, id_column: str, log_func=None) -> pd.DataFrame:
        nuevas_filas = self.detectar_nuevas_filas(df1, df2, id_column)
        if nuevas_filas.empty:
            log_func and log_func("No se detectaron nuevas filas.")
            return pd.DataFrame()
        log_func and log_func(f"Se detectaron {len(nuevas_filas)} nuevas filas:")
        for _, row in nuevas_filas.iterrows():
            log_func and log_func(f"  - Nueva fila con ID: {row[id_column]}")
        return nuevas_filas

    def crear_resumen_informe(self, historial_df: pd.DataFrame, mes: str) -> pd.DataFrame:
        if historial_df.empty:
            return pd.DataFrame([{
                "Mes": mes,
                "Total de Cambios": 0,
                "Modificaciones": 0,
                "Nuevos Registros": 0,
                "Registros Eliminados": 0
            }])
        return pd.DataFrame([{
            "Mes": mes,
            "Total de Cambios": len(historial_df),
            "Modificaciones": (historial_df["Tipo de Cambio"] == MODIFICACION).sum(),
            "Nuevos Registros": historial_df["Tipo de Cambio"].str.contains(NUEVO_REGISTRO, na=False).sum(),
            "Registros Eliminados": (historial_df["Tipo de Cambio"] == REGISTRO_ELIMINADO).sum()
        }])

    def procesar_comparacion(self, config: dict, log_func=None) -> pd.DataFrame:
        try:
            archivo, mes, id_column, titulo = config["archivo_base"], config["mes_actual"], config["id_column"], config["titulo"]

            log_func and log_func(f"Cargando datos de las hojas '{config['sheet_anterior']}' y '{config['sheet_actual']}'...")
            df1, df2 = self.cargar_datos(archivo, config["sheet_anterior"], config["sheet_actual"])

            log_func and log_func(f"Alineando DataFrames... {df1.shape} y {df2.shape}")
            df1, df2 = self.data_utils.alinear_dataframes(df1, df2)

            log_func and log_func("Generando tabla de comparación completa...")
            comparacion_completa = self.tabla_comparacion_completa(df1, df2)
            mask_cambios = (df1 != df2) & ~(df1.eq("") & df2.eq(""))

            self.obtener_resumen_nuevas_filas(df1, df2, id_column, log_func)

            log_func and log_func("Creando historial de cambios...")
            historial = self.crear_historial_de_cambios(df1, df2, mes, id_column, log_func)
            historial_df = pd.DataFrame(historial)

            nombres = {
                "historial": f"{titulo} - Historial de Cambios.xlsx",
                "comparacion": f"{titulo}-procedimiento-{mes}.xlsx",
                "actualizados": f"{titulo} - Datos Actualizados.xlsx"
            }

            if os.path.exists(nombres["historial"]):
                historial_existente = pd.read_excel(nombres["historial"])
                historial_total = pd.concat([historial_existente, historial_df], ignore_index=True)
            else:
                historial_total = historial_df
            historial_total.to_excel(nombres["historial"], index=False)

            log_func and log_func("Generando resumen de cambios...")
            resumen = self.crear_resumen_informe(historial_df, mes)
            with pd.ExcelWriter(nombres["historial"], engine='openpyxl', mode='a') as writer:
                resumen.to_excel(writer, sheet_name='Resumen Informe', index=False)

            log_func and log_func("Creando DataFrame actualizado...")
            reemplazo = self.crear_reemplazo(df1, df2)
            with pd.ExcelWriter(nombres["comparacion"], engine="xlsxwriter") as writer:
                self.file_utils.escribir_formateado(writer, "Comparación Completa", comparacion_completa, mask_cambios, resaltar=True)
                self.file_utils.escribir_formateado(writer, "Solo Cambios", comparacion_completa[mask_cambios.any(axis=1)], mask_cambios, resaltar=True)

            pd.DataFrame(reemplazo).to_excel(nombres["actualizados"], index=False)

            log_func and log_func("Proceso completado exitosamente.")
            log_func and log_func("Archivos generados:")
            for archivo in nombres.values():
                log_func and log_func(f"- {archivo}")

            return resumen

        except Exception as e:
            log_func and log_func(f"Error durante el procesamiento: {e}")
            raise
