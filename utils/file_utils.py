"""
Utilidades para el manejo de archivos Excel y formateo
"""

import os


class FileUtils:
    """Clase con utilidades para el manejo de archivos"""
    
    @staticmethod
    def escribir_formateado(writer, sheet_name, df, mask_cambios=None, resaltar=False):
        """Escribe un DataFrame en Excel con formato condicional para resaltar cambios."""
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet

        # Formatos
        formato_general = workbook.add_format({'border': 1})
        formato_cambio = workbook.add_format({'bg_color': '#FF9999', 'border': 1})
        formato_header = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})

        # Escribir índice
        worksheet.write(0, 0, "Index", formato_header)
        for row_idx, index_val in enumerate(df.index, start=1):
            worksheet.write(row_idx, 0, index_val, formato_general)

        # Escribir encabezados
        for col_idx, col_name in enumerate(df.columns, start=1):
            worksheet.write(0, col_idx, col_name, formato_header)

        # Escribir datos
        for row_idx in range(df.shape[0]):
            for col_idx in range(df.shape[1]):
                valor = df.iat[row_idx, col_idx]
                formato = formato_general
                
                if resaltar:
                    if isinstance(valor, str) and "➔" in valor:
                        formato = formato_cambio
                    elif mask_cambios is not None and mask_cambios.iat[row_idx, col_idx]:
                        formato = formato_cambio
                        
                worksheet.write(row_idx + 1, col_idx + 1, valor, formato)

        # Ajustar ancho de columnas
        worksheet.set_column(0, df.shape[1], 20)
    
    @staticmethod
    def validar_archivo_excel(ruta_archivo):
        """Valida si el archivo es un Excel válido y existe"""
        if not os.path.exists(ruta_archivo):
            return False, "El archivo no existe"
        
        if not ruta_archivo.lower().endswith(('.xlsx', '.xls')):
            return False, "El archivo no es un archivo Excel válido"
        
        return True, "Archivo válido"
    
    @staticmethod
    def crear_directorio_si_no_existe(ruta):
        """Crea un directorio si no existe"""
        if not os.path.exists(ruta):
            os.makedirs(ruta)
            return True
        return False
    
    @staticmethod
    def obtener_nombre_archivo_seguro(nombre_base, extension='.xlsx'):
        """Genera un nombre de archivo único si ya existe"""
        contador = 1
        nombre_original = f"{nombre_base}{extension}"
        nombre_final = nombre_original
        
        while os.path.exists(nombre_final):
            nombre_final = f"{nombre_base}_{contador}{extension}"
            contador += 1
        
        return nombre_final