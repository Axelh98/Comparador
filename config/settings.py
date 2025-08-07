"""
Configuraciones y constantes de la aplicación
"""

# Configuraciones de la aplicación
APP_CONFIG = {
    "TITLE": "Comparador de Tablas Excel",
    "VERSION": "1.0.0",
    "WINDOW_SIZE": "800x650",
    "DEFAULT_PROJECT_TITLE": "Propiedades"
}

# Configuraciones de archivos
FILE_CONFIG = {
    "EXCEL_EXTENSIONS": [".xlsx", ".xls"],
    "MAX_FILE_SIZE_MB": 100,
    "OUTPUT_ENCODING": "utf-8"
}

# Configuraciones de formato Excel
EXCEL_FORMAT = {
    "HEADER_COLOR": "#D9D9D9",
    "CHANGE_COLOR": "#FF9999",
    "BORDER_STYLE": 1,
    "COLUMN_WIDTH": 20
}

# Mensajes de la aplicación
MESSAGES = {
    "FILE_NOT_SELECTED": "Por favor, seleccione un archivo Excel primero",
    "INVALID_FILE": "El archivo seleccionado no es válido",
    "PROCESS_COMPLETE": "Proceso completado exitosamente",
    "ERROR_LOADING": "Error al cargar el archivo",
    "ERROR_PROCESSING": "Error durante el procesamiento"
}

# Nombres de columnas por defecto
DEFAULT_COLUMNS = {
    "CHANGE_TYPE": "Tipo de Cambio",
    "PREVIOUS_VALUE": "Valor Anterior",
    "NEW_VALUE": "Valor Nuevo",
    "COLUMN_NAME": "Columna",
    "ROW_INDEX": "Fila (índice)",
    "MONTH": "Mes"
}

# Tipos de cambios
CHANGE_TYPES = {
    "MODIFICATION": "Modificación",
    "NEW_RECORD": "Nuevo Registro",
    "DELETED_RECORD": "Registro Eliminado"
}