import openpyxl
import pdb

Input_File_Path = 'C:\\Users\\sgh4n\\OneDrive\\Escritorio\\Mariser\\Modelado\\Input\\Mayores Contables, Modelado Spreadsheet.xlsx' # For now, this will be something static, later on it will be dynamically assigned
Input_File = ''

# pdb.set_trace()
try:
        with open(Input_File_Path, 'r') as Input_File:
                 from openpyxl import load_workbook
                 load_workbook(Input_File.name)
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



