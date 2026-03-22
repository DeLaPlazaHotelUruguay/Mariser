import openpyxl
import pdb

Input_File_Path = 'C:\\Users\\sgh4n\\OneDrive\\Escritorio\\Mariser\\Modelado\\Input\\Mayores Contables, Modelado Spreadsheet.xlsx' # For now, this will be something static, later on it will be dynamically assigned
Input_File = ''

# pdb.set_trace()
try:
        with open(Input_File_Path, 'r') as Input_File:
                 pass

except IOError as Error:
        print("Error leyendo el archivo ", Input_File_Path, ": ", Error)
        F"Abortando programa :("
        quit()

print("Archivo ", Input_File.name, "abierto con exito!")
