# Comparador de Tablas Excel

Una aplicación GUI para comparar y analizar cambios entre diferentes hojas de cálculo Excel.

## Estructura del Proyecto

```
comparador-excel/
│
├── main.py                    # Archivo principal de ejecución
├── requirements.txt           # Dependencias del proyecto
├── README.md                 # Documentación
│
├── gui/                      # Módulo de interfaz gráfica
│   ├── __init__.py
│   └── main_window.py        # Ventana principal de la aplicación
│
├── processors/               # Módulo de procesamiento de datos
│   ├── __init__.py
│   └── excel_processor.py    # Procesador principal de Excel
│
├── utils/                    # Módulo de utilidades
│   ├── __init__.py
│   ├── data_utils.py         # Utilidades de manipulación de datos
│   └── file_utils.py         # Utilidades de manejo de archivos
│
└── config/                   # Módulo de configuraciones
    ├── __init__.py
    └── settings.py           # Configuraciones y constantes
```

## Características

- **Interfaz gráfica intuitiva**: GUI desarrollada con Tkinter
- **Comparación de hojas Excel**: Detecta modificaciones, registros nuevos y eliminados
- **Exportación de resultados**: Genera reportes detallados en Excel
- **Historial de cambios**: Mantiene un registro histórico de todas las modificaciones
- **Configuración flexible**: Permite seleccionar columnas ID y períodos de comparación

## Instalación

1. Clona o descarga el proyecto
2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

## Uso

1. Ejecuta la aplicación:
   ```bash
   python main.py
   ```

2. Selecciona tu archivo Excel
3. Configura las hojas a comparar
4. Define la columna ID para el seguimiento de registros
5. Ejecuta la comparación

## Salida

La aplicación genera tres tipos de archivos:

- **Historial de Cambios**: Registro detallado de todas las modificaciones
- **Comparación Completa**: Vista lado a lado de los cambios
- **Datos Actualizados**: Versión actualizada de los datos

## Módulos

### GUI (`gui/`)
- `main_window.py`: Interfaz principal de usuario

### Procesadores (`processors/`)
- `excel_processor.py`: Lógica de comparación y procesamiento

### Utilidades (`utils/`)
- `data_utils.py`: Funciones para manipulación de datos
- `file_utils.py`: Funciones para manejo de archivos

### Configuración (`config/`)
- `settings.py`: Constantes y configuraciones de la aplicación

## Requisitos

- Python 3.7+
- pandas
- openpyxl
- xlsxwriter
- tkinter (incluido con Python)

## Contribución

1. Fork el proyecto
2. Crea una rama para tu feature
3. Commit tus cambios
4. Push a la rama
5. Abre un Pull Request