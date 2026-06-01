# GUI-AMR
Python GUI for LabVIEW experiment setup. Lets the user select experiment parameters from an Excel database, look up AMR properties automatically, and return the final configuration back to LabVIEW.


LabVIEW Experiment Configuration GUI
A Python pop-up panel launched from LabVIEW that allows the user to select experiment parameters from an Excel database, look up AMR thermal properties automatically, and return the final configuration back to LabVIEW.
# Requirements
Python 3.6+  
openpyxl  

# Install the dependency with:
pip install openpyxl  
# Configuration
The path to the Excel database is defined as a constant at the top of the file:  
EXCEL_PATH = r"C:\Users\laura\Desktop\NewFormatDatabase.xlsx"  
Change this to match the location of the file on your machine before running.  
# Excel Database
A mockup of the Excel database is included in this repository for testing purposes. The actual database used in the lab contains sensitive data and cannot be shared.  
# How it works  
The GUI lets the user select a user, material type, and cascade from the database. Once a cascade is selected, the relevant AMR properties (cp, porosity, mass) are retrieved automatically from the spreadsheet. On confirmation, the selected configuration is returned.  
# Usage
Standalone (testing):  
python labview_gui.py  
The selected configuration will be printed to the console.  
# From LabVIEW:  
The script is designed to be called from LabVIEW via the Python Node. The GUI will pop up, and once the user confirms, the values are returned directly to LabVIEW to be used in the running VI.  
