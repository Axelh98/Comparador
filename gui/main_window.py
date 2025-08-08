"""
MAIN WINDOW FOR THE COMPARADOR EXCEL APPLICATION
This module defines the main GUI for the Excel comparison application.
It allows users to select Excel files, configure comparison settings,
and execute the comparison process.
"""

import os
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

from processors.excel_processor import ExcelProcessor


class ComparadorExcelApp:
    """Main class for the GUI application"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Comparator of Excel Databases")
        self.root.geometry("800x650")
        self.root.resizable(True, True)
        
        # Variables to hold user inputs
        self.archivo_excel = tk.StringVar()
        self.directorio_salida = tk.StringVar()
        self.directorio_salida.set(os.getcwd())  # Current directory by default
        self.titulo_proyecto = tk.StringVar(value="Name of the Project")
        self.columna_id = tk.StringVar()
        self.mes_actual = tk.StringVar(value=datetime.datetime.now().strftime("%B"))
        
        # Lists for sheets and columns
        self.hojas_disponibles = []
        self.columnas_disponibles = []
        self.hoja_anterior = tk.StringVar()
        self.hoja_actual = tk.StringVar()
        
        # Initialize the Excel processor
        self.processor = ExcelProcessor()

        # Create the interface
        self.crear_interfaz()
        
    def crear_interfaz(self):
        """Create all user interface elements"""
        # Style
        estilo = ttk.Style()
        estilo.configure("TFrame", padding=10)
        estilo.configure("TButton", padding=5)
        estilo.configure("TLabel", padding=5)
        estilo.configure("Header.TLabel", font=("Arial", 12, "bold"))

        # Main container
        main_frame = ttk.Frame(self.root, style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create interface sections
        self._crear_seccion_titulo(main_frame)
        self._crear_seccion_archivo(main_frame)
        self._crear_seccion_configuracion(main_frame)
        self._crear_seccion_hojas(main_frame)
        self._crear_seccion_columna_id(main_frame)
        self._crear_botones(main_frame)
        self._crear_seccion_log(main_frame)
        
        # Set grid configuration
        main_frame.columnconfigure(1, weight=1)
    
    def _crear_seccion_titulo(self, parent):
        """Create the title section"""
        ttk.Label(parent, text="Comparator of Excel Databases",
                 style="Header.TLabel").grid(row=0, column=0, columnspan=3, pady=10)
    
    def _crear_seccion_archivo(self, parent):
        """Create the file selection section"""
        ttk.Label(parent, text="Excel File:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(parent, textvariable=self.archivo_excel, width=50).grid(
            row=1, column=1, sticky=tk.W+tk.E)
        ttk.Button(parent, text="Browse...", command=self.seleccionar_archivo).grid(
            row=1, column=2, padx=5)
    
    def _crear_seccion_configuracion(self, parent):
        """Create the basic configuration sections"""
        # Project title
        ttk.Label(parent, text="Project Title:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(parent, textvariable=self.titulo_proyecto, width=50).grid(
            row=2, column=1, sticky=tk.W+tk.E)

        # Current month
        ttk.Label(parent, text="Comparison Period:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(parent, textvariable=self.mes_actual, width=50).grid(
            row=3, column=1, sticky=tk.W+tk.E)

        # Output directory
        ttk.Label(parent, text="Output Directory:").grid(row=4, column=0, sticky=tk.W)
        ttk.Entry(parent, textvariable=self.directorio_salida, width=50).grid(
            row=4, column=1, sticky=tk.W+tk.E)
        ttk.Button(parent, text="Browse...", command=self.seleccionar_directorio).grid(
            row=4, column=2, padx=5)
    
    def _crear_seccion_hojas(self, parent):
        """Create the sheet selection section"""
        sheets_frame = ttk.LabelFrame(parent, text="Sheet Selection")
        sheets_frame.grid(row=5, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)

        ttk.Label(sheets_frame, text="Previous Sheet:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.combo_anterior = ttk.Combobox(sheets_frame, textvariable=self.hoja_anterior, 
                                           state="readonly", width=40)
        self.combo_anterior.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)

        ttk.Label(sheets_frame, text="Current Sheet:").grid(
            row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.combo_actual = ttk.Combobox(sheets_frame, textvariable=self.hoja_actual, 
                                         state="readonly", width=40)
        self.combo_actual.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
    
    def _crear_seccion_columna_id(self, parent):
        """Create the ID column selection section"""
        id_frame = ttk.LabelFrame(parent, text="ID Column Configuration")
        id_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)

        ttk.Label(id_frame, text="ID Column:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.combo_id = ttk.Combobox(id_frame, textvariable=self.columna_id, 
                                     state="readonly", width=40)
        self.combo_id.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
    
    def _crear_botones(self, parent):
        """Create the action buttons"""
        # Button to load sheets
        ttk.Button(parent, text="Load Sheets and Columns", 
                   command=self.cargar_hojas).grid(row=7, column=0, columnspan=3, pady=10)

        # Execution buttons
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=8, column=0, columnspan=3, pady=10)

        ttk.Button(button_frame, text="Run Comparison", 
                   command=self.ejecutar_comparacion).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", 
                   command=self.root.destroy).pack(side=tk.LEFT, padx=5)
    
    def _crear_seccion_log(self, parent):
        """Create the log section"""
        log_frame = ttk.LabelFrame(parent, text="Operation Log")
        log_frame.grid(row=9, column=0, columnspan=3, sticky=tk.W+tk.E+tk.N+tk.S, pady=10)

        # Text widget for the log
        self.log_text = tk.Text(log_frame, height=10, width=80, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def seleccionar_archivo(self):
        """Open the file explorer to select the Excel file"""
        archivo = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if archivo:
            self.archivo_excel.set(archivo)
            self.log("Archivo seleccionado: " + archivo)
    
    def seleccionar_directorio(self):
        """Open the file explorer to select the output directory"""
        directorio = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directorio:
            self.directorio_salida.set(directorio)
            self.log("Output directory: " + directorio)

    def cargar_hojas(self):
        """Load the available sheets from the selected Excel file"""
        archivo = self.archivo_excel.get()
        if not archivo:
            messagebox.showerror("Error", "Please select an Excel file first")
            return
        
        try:
            # Load available sheets
            excel = pd.ExcelFile(archivo)
            self.hojas_disponibles = excel.sheet_names
            
            # Update the combo boxes with available sheets
            self.combo_anterior['values'] = self.hojas_disponibles
            self.combo_actual['values'] = self.hojas_disponibles

            # Select the first two sheets if there are at least two
            if len(self.hojas_disponibles) >= 2:
                self.hoja_anterior.set(self.hojas_disponibles[0])
                self.hoja_actual.set(self.hojas_disponibles[1])
            elif len(self.hojas_disponibles) == 1:
                self.hoja_anterior.set(self.hojas_disponibles[0])
                self.hoja_actual.set(self.hojas_disponibles[0])

            # Load the columns from the first sheet
            self.cargar_columnas()

            self.log(f"Loaded {len(self.hojas_disponibles)} sheets from the file")

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {str(e)}")
            self.log(f"ERROR: {str(e)}")
    
    def cargar_columnas(self):
        """Load the available columns in the selected sheet"""
        archivo = self.archivo_excel.get()
        hoja = self.hoja_actual.get()
        
        if not archivo or not hoja:
            return
        
        try:
            # Load the columns
            df = pd.read_excel(archivo, sheet_name=hoja, nrows=0)
            self.columnas_disponibles = df.columns.tolist()

            # Update the ID columns combo box
            self.combo_id['values'] = self.columnas_disponibles

            # Try to automatically select a suitable ID column
            id_columns = [col for col in self.columnas_disponibles if 'id' in col.lower()]
            if id_columns:
                self.columna_id.set(id_columns[0])
            elif self.columnas_disponibles:
                self.columna_id.set(self.columnas_disponibles[0])

            self.log(f"Loaded {len(self.columnas_disponibles)} columns from the sheet '{hoja}'")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading columns: {str(e)}")
            self.log(f"ERROR: {str(e)}")
    
    def log(self, mensaje):
        """Adds a message to the log with date and time"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.log_text.see(tk.END)  # Auto-scroll to the end

    def ejecutar_comparacion(self):
        """Executes the comparison process with the selected parameters"""
        # Validate that all necessary fields are filled
        if not self.validar_campos():
            return
            
        # Configure the processing
        config = {
            "archivo_base": self.archivo_excel.get(),
            "titulo": self.titulo_proyecto.get(),
            "mes_actual": self.mes_actual.get(),
            "sheet_anterior": self.hoja_anterior.get(),
            "sheet_actual": self.hoja_actual.get(),
            "id_column": self.columna_id.get(),
            "directorio_salida": self.directorio_salida.get()
        }
        
        self.log("Iniciando proceso de comparación...")
        
        try:
            # Change the current working directory to the output directory
            os.chdir(config["directorio_salida"])
            
            # Execute the comparison
            resumen = self.processor.procesar_comparacion(config, self.log)

            # Show success message
            mensaje = (f"Process completed successfully.\n\n"
                      f"Detected {resumen['Total de Cambios'].values[0]} changes:\n"
                      f"- {resumen['Modificaciones'].values[0]} modifications\n"
                      f"- {resumen['Nuevos Registros'].values[0]} new records\n"
                      f"- {resumen['Registros Eliminados'].values[0]} deleted records")
                      
            messagebox.showinfo("Process Completed", mensaje)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during processing: {str(e)}")
            self.log(f"ERROR: {str(e)}")
    
    def validar_campos(self):
        """Validates that all necessary fields are complete"""
        campos = [
            (self.archivo_excel.get(), "You must select an Excel file"),
            (self.hoja_anterior.get(), "You must select a previous sheet"),
            (self.hoja_actual.get(), "You must select a current sheet"),
            (self.columna_id.get(), "You must select an ID column"),
            (self.mes_actual.get(), "You must specify a comparison period"),
            (self.titulo_proyecto.get(), "You must specify a project title")
        ]
        
        for valor, mensaje in campos:
            if not valor:
                messagebox.showerror("Error de validación", mensaje)
                return False
                
        return True