import openpyxl
import pdb
import re

# Open and read the spreadsheet file
Input_File_Path = 'C:\\Users\\sgh4n\\OneDrive\\Escritorio\\Mariser\\Modelado\\Input\\Mayores Contables, Modelado Spreadsheet.xlsx' # For now, this will be something static, later on it will be dynamically assigned
Input_File = None
Input_Worksheet = None

# pdb.set_trace()
try:
        with open(Input_File_Path, 'r') as Input_File:
                 from openpyxl import load_workbook
                 Input_Worksheet = load_workbook(Input_File.name).active
                 pass

except IOError as Error:
        print("Error leyendo el archivo ", Input_File_Path, ": ", Error)
        print("Abortando programa :(")
        quit()
except Exception as Error:
        print("Hubo un error con el archivo de Excel ingresasdo.")
        print("Abortando programa :(")
        quit()

print("Archivo ", Input_File.name, "existe y es legible!")

# Extract the lotes from the file
Lotes = []

def Get_Lotes_Numbers():
        # Extract the lote numbers from the operations column
        Operations_Column = 2
        Operations = []
        Lote_Numbers = []
        
        for i, Row in enumerate(Input_Worksheet):
                if i == 0: continue
                Operations.append(Row[Operations_Column].value)
                
        def Operation_Is_Cierre(Operation): return re.match(r"^Cierre de Lote Nro\.[0-9]+$", Operation)
        def Operation_Is_Cobro(Operation): return re.match(r"^Cob\. Lote Nro\. [0-9]+ s\/Compr\.[0-9]+$", Operation)
        
Get_Lotes_Numbers()

