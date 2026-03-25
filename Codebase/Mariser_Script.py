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
                self.Operations = Operations # Currently doesn't work; it stores Operations from all the Lotes. I won't bother with this rn
                self.Cierres = []
                self.Cobros = []
                
        def Get_Balance(self):
                Balance = 0
                
                for Operation in (self.Cierres + self.Cobros):
                        if type(Operation) is Cierre: Balance -= Operation.Amount
                        elif type(Operation) is Cobro: Balance += Operation.Amount
                return Balance
        
        def To_String(self):
                print("Lote Nro" + self.Number, self.Initial_Date)
                
class Cierre:
        def __init__(self, Lote_Number, Date, Num_Asto, Amount):
                self.Lote_Number = Lote_Number
                self.Date = Date
                self.Asto = Num_Asto
                self.Amount = Amount
                
        def Operation_Is_Cierre(Operation): return re.match(r"^Cierre de Lote Nro\.[0-9]+$", Operation)
        def To_String(self): print("Cierre, Numero_Lote" + self.Lote_Number, self.Asto, self.Date, self.Amount )

class Cobro:
        def __init__(self, Lote_Number, Date, Compr, Num_Asto, Amount):
                self.Lote_Number = Lote_Number
                self.Date = Date
                self.Compr = Compr
                self.Asto = Num_Asto
                self.Amount = Amount
                
        def Operation_Is_Cobro(Operation): return re.match(r"^Cob\. Lote Nro\. [0-9]+ s\/Compr\.[0-9]+$", Operation)
        def To_String(self): print("Cobro, Numero_Lote" + self.Lote_Number, self.Date, self.Compr, self.Asto, self.Amount)
        
def Get_Lotes():
        def Lote_Exists(Number):
                for Lote in Lotes.values(): 
                        if Number == Lote.Number: return True
                return False
                
        def Date_Is_Earlier_Than_Lote(Lote, Operand_Date): return (Lote.Initial_Date > Operand_Date)
        
        Operations_Column = 2
        Dates_Column = 0
        Lotes = {}

        for Row_Number, Row in enumerate(Input_Worksheet):
                Operation = None
                if Row_Number == 0: continue
                if Row_Number == Input_Worksheet.max_row - 1: continue
                Operation = Row[Operations_Column].value
                
                Lote_Number = None
                Operation_Date = None
                
                if Cierre.Operation_Is_Cierre(Operation): Lote_Number = Operation.split('.')[1]
                elif Cobro.Operation_Is_Cobro(Operation): Lote_Number = re.search(r'^Cob\. Lote Nro\. ([0-9]+) s\/Compr\.[0-9]+$', Operation).group(1)
                else:
                        print("Operacion no reconocida! La tercera celda de la fila {0} no coincide con el formato de un Cierre ni de un Cobro.".format(Row_Number),
                        "Probablemente seria una buena idea rehacer el archivo de Mayores Contables."
                        )
                        quit()
                        
                Operation_Date = Row[Dates_Column].value.date()
                
                if not Lote_Exists(Lote_Number): Lotes[Lote_Number] = Lote(Lote_Number, Operation_Date)
                elif Date_Is_Earlier_Than_Lote( Lotes[Lote_Number], Operation_Date ): Lotes[Lote_Number].Initial_Date = Operation_Date
                
                # Construct the operation
                def Get_Compr_Operation(Operation): return re.search(r"^Cob\. Lote Nro\. [0-9]+ s\/Compr\.([0-9]+)$", Operation).group(1)
                
                Current_Lote = Lotes[Lote_Number]
                Asto_Column = 1
                Value_Column = None
                Num_Asto = Row[Asto_Column].value

                if Cierre.Operation_Is_Cierre(Operation):
                        Value_Column = 3
                        Compr_Column = 2
                        Amount = Row[Value_Column].value
                        
                        Operation_Object = Cierre(Lote_Number, Operation_Date, Num_Asto, Amount)
                        Current_Lote.Cierres.append(Operation_Object)
                        Current_Lote.Operations.append(Operation_Object)
                elif Cobro.Operation_Is_Cobro(Operation):
                        Value_Column = 4
                        Compr = Get_Compr_Operation(Operation)
                        Amount = Row[Value_Column].value
                        
                        Operation_Object = Cobro(Lote_Number, Operation_Date, Compr, Num_Asto, Amount)
                        Current_Lote.Cobros.append(Operation_Object)
                        Current_Lote.Operations.append(Operation_Object)
        return Lotes
                        
Lotes = Get_Lotes()

# Build the output file
from openpyxl import Workbook
Output_Workbook = Workbook()
Output_Worksheet = Output_Workbook.active

def Get_Output_File():
        def Get_Lote_Rowspan(Lote): return len(Lote.Cierres) if len(Lote.Cierres) > len(Lote.Cobros) else len(Lote.Cobros)
        def Order_Lotes_Inplace():
                Ordered_Lotes = {}
                
                # Get the Initial Dates of each Lote, from oldest to newest
                Lotes_Tuples = []
                for Lote_Tuple in Lotes.items(): Lotes_Tuples.append( (Lote_Tuple[1].Initial_Date, Lote_Tuple[0]) )
                Lotes_Tuples.sort()
                                
                # Assign the Lotes by their initial date (oldest first)
                for Lote_Tuple in Lotes_Tuples:
                        Lote_Number = Lote_Tuple[1]
                        Ordered_Lotes[str(Lote_Number)] = Lotes[Lote_Number]
        
        Sheet_Rows = []
        
        # Construct the Lote rows
        Lote_Rows = []
        
        Order_Lotes_Inplace()
        for Lote in Lotes.values():
                def Cierre_Num_Exists(Number, Lote): return Number < len(Lote.Cierres)
                def Cobro_Num_Exists(Number, Lote): return Number < len(Lote.Cobros)
                def Get_State_Character_Lote(Balance):
                        if Balance == 0: return '🗸'
                        elif Balance < 0: return '🗴'
                        elif Balance > 0: return '⚠'
                
                # Contruct the rows used by the lote
                Lote_Rowspan = Get_Lote_Rowspan(Lote)
                
                for Row_Number in range(Lote_Rowspan):
                        Row_Cells = [None]
                        if Row_Number == 0: Row_Cells[0] = Lote.Number
                        
                        Current_Cierre = None
                        Current_Cobro = None
                        if Cierre_Num_Exists(Row_Number, Lote): Current_Cierre = Lote.Cierres[Row_Number]
                        if Cobro_Num_Exists(Row_Number, Lote): Current_Cobro = Lote.Cobros[Row_Number]
                        
                        if Current_Cierre is not None: Row_Cells += [Current_Cierre.Asto, Current_Cierre.Date, Current_Cierre.Amount]
                        else: Row_Cells += [None, None, None]
                        
                        if Current_Cobro is not None: Row_Cells += [Current_Cobro.Asto, Current_Cobro.Date, Current_Cobro.Compr, Current_Cobro.Amount]
                        else: Row_Cells += [None, None, None, None]
                        
                        if Row_Number == 0: Row_Cells += [Get_State_Character_Lote(Lote.Get_Balance()), Lote.Get_Balance()]
                            
                        Lote_Rows.append(Row_Cells)
        
        # Construct the header rows
        Header_Rows = []
        Upper_Header_Row = ["Lote", "Cierres", None, None, "Cobros", None, None, "Estado", "Balance"]
        Lower_Header_Row = [None, "Nro. Asto", "Fecha", "Valor", "Nro. Asto", "Fecha", "Compr.", "Valor", None, None]
        
        Header_Rows.append(Upper_Header_Row)
        Header_Rows.append(Lower_Header_Row)
        
        # Load the rows to the sheet
        Sheet_Rows += Header_Rows
        Sheet_Rows += Lote_Rows
        
        # wb = Workbook()
        # ws = wb.active
        # for Row in Sheet_Rows: ws.append(Row)
        # wb.save("Test.xlsx")
Get_Output_File()
