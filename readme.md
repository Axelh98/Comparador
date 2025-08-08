Excel Data Comparator
This Python project helps you compare data between two Excel sheets. It detects changes, new records, and deleted records, and generates detailed reports with summaries and optional charts.

Features
Load and normalize Excel data.

Compare two sheets based on a key ID column.

Detect modifications, new rows, and deleted rows.

Create detailed change history.

Generate summary reports.

Export results to new Excel files with formatting.

Optional: generate charts to visualize changes.

Simple graphical user interface (GUI) built with Tkinter.

Technologies and Libraries
Python 3.x

pandas — for data handling and comparison.

tkinter — for the GUI.

matplotlib — (optional) for generating charts.

os — for file and folder operations.

How to Use
Run the GUI application.

Select your Excel file and choose the sheets to compare.

Pick the key ID column that identifies each record.

Set the project title and comparison month.

Run the comparison and check the logs for progress.

Find the output Excel files with detailed reports and summaries in the output folder.

File Structure
excel_processor.py — main module for loading, comparing, and saving Excel data.

data_utils.py — helper functions for cleaning and aligning data.

file_utils.py — helper functions for writing Excel files with formatting.

excel_charts.py — functions to create charts from change history.

main_window.py — GUI code.