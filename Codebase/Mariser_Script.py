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
class Lote:
        def __init__(self, Number, Initial_Date, Operations = []):
                self.Number = Number
                self.Initial_Date = Initial_Date
                self.Operations = Operations
                self.Balance = None
                
        def To_String(self):
                print("Lote Nro" + self.Number, self.Initial_Date)
                
class Cierre:
        def __init__(self, Lote_Number, Date, Amount):
                self.Lote_Number = Lote_Number
                self.Date = Date
                self.Amount = Amount
                
        def Operation_Is_Cierre(Operation): return re.match(r"^Cierre de Lote Nro\.[0-9]+$", Operation)

class Cobro:
        def __init__(self, Lote_Number, Date, Compr, Amount):
                self.Lote_Number = Lote_Number
                self.Date = Date
                self.Compr = Compr
                self.Amount = Amount
                
        def Operation_Is_Cobro(Operation): return re.match(r"^Cob\. Lote Nro\. [0-9]+ s\/Compr\.[0-9]+$", Operation)

        
Lotes = []

def Get_Lotes_Numbers():
        # Extract the lote numbers from the operations column
        Operations_Column = 2
        Operations = []
        Lote_Numbers = []
        
        for Row_Number, Row in enumerate(Input_Worksheet):
                if Row_Number == 0: continue
                if Row_Number == Input_Worksheet.max_row - 1: continue
                Operations.append(Row[Operations_Column].value)
                
        def Operation_Is_Cierre(Operation): return re.match(r"^Cierre de Lote Nro\.[0-9]+$", Operation)
        def Operation_Is_Cobro(Operation): return re.match(r"^Cob\. Lote Nro\. [0-9]+ s\/Compr\.[0-9]+$", Operation)
        
        for Row_Number, Operation in enumerate(Operations):
                if Operation_Is_Cierre(Operation): Lote_Numbers.append(Operation.split('.')[1])
                elif Operation_Is_Cobro(Operation): Lote_Numbers.append(re.search(r'^Cob\. Lote Nro\. ([0-9]+) s\/Compr\.[0-9]+$', Operation).group(1))
                else:
                        print("Operacion no reconocida! La tercera celda de la fila {0} no coincide con el formato de un Cierre ni de un Cobro.".format(Row_Number),
                        "Probablemente seria una buena idea rehacer el archivo de Mayores Contables."
                        )
                        quit()
                        
        Lote_Numbers = set(Lote_Numbers)
        return (Lote_Numbers)
Get_Lotes_Numbers()

def Get_Lotes_Dates():
        Dates_Column = 0
        Lote_Dates = []
        
        for Row_Number, Row in enumerate(Input_Worksheet):
                if Row_Number == 0: continue
                if Row_Number == Input_Worksheet.max_row - 1: continue
                Lote_Dates.append(Row[Dates_Column].value)

Get_Lotes_Dates()

def Get_Lotes():
        def Lote_Exists(Number):
                for Lote in Lotes: 
                        if Number == Lote.Number: return True
                return False
                
        def Date_Is_Earlier_Than_Lote(Lote, Operand_Date): return (Lote.Date > Operand_Date)
        
        Operations_Column = 2
        Dates_Column = 0
        Lotes = []

        for Row_Number, Row in enumerate(Input_Worksheet):
                Operation = None
                if Row_Number == 0: continue
                if Row_Number == Input_Worksheet.max_row - 1: continue
                Operation = Row[Operations_Column].value
                
                Lote_Number = None
                Lote_Initial_Date = None
                
                if Cierre.Operation_Is_Cierre(Operation): Lote_Number = Operation.split('.')[1]
                elif Cobro.Operation_Is_Cobro(Operation): Lote_Number = re.search(r'^Cob\. Lote Nro\. ([0-9]+) s\/Compr\.[0-9]+$', Operation).group(1)
                else:
                        print("Operacion no reconocida! La tercera celda de la fila {0} no coincide con el formato de un Cierre ni de un Cobro.".format(Row_Number),
                        "Probablemente seria una buena idea rehacer el archivo de Mayores Contables."
                        )
                        quit()
                        
                Lote_Initial_Date = Row[Dates_Column].value.date()
                
                if not Lote_Exists(Lote_Number):
                        Lotes.append(Lote(Lote_Number, Lote_Initial_Date))
                        continue
Get_Lotes()


