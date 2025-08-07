"""
Main entry point for the Comparador Excel application.
This script initializes the main window and starts the Tkinter event loop.
"""

import tkinter as tk
from gui.main_window import ComparadorExcelApp

def main():
    """Main function to start the application"""
    root = tk.Tk()
    app = ComparadorExcelApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()