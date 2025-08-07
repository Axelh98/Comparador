"""
Utilidades para el manejo y manipulación de datos
"""

import pandas as pd


class DataUtils:
    """Clase con utilidades para el manejo de datos"""
    
    @staticmethod
    def normalizar_df(df):
        """Normaliza el DataFrame para comparación."""
        return df.fillna("").astype(str)

    @staticmethod
    def alinear_dataframes(df1, df2):
        """
        Alinea dos DataFrames para asegurar una comparación adecuada,
        incluso cuando tienen diferentes cantidades de filas o columnas.
        """
        # Aseguramos que ambos DataFrames tengan las mismas columnas
        all_columns = sorted(list(set(df1.columns) | set(df2.columns)))
        
        # Añadimos columnas faltantes con valores NaN
        for col in all_columns:
            if col not in df1.columns:
                df1[col] = pd.NA
            if col not in df2.columns:
                df2[col] = pd.NA
        
        # Alineamos ambos DataFrames, permitiendo filas y columnas nuevas
        df1_aligned, df2_aligned = df1.align(df2, join='outer', axis=None)
        
        return df1_aligned, df2_aligned
    
    @staticmethod
    def validar_columna_existe(df, columna):
        """Valida si una columna existe en el DataFrame"""
        return columna in df.columns
    
    @staticmethod
    def obtener_columnas