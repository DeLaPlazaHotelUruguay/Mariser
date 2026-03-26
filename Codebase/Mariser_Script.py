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
                print("Lote Nro" + self.Number, self.Initial_Date.date())
                
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
                        
                Operation_Date = Row[Dates_Column].value
                
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
        
        ## Group lotes by their balance
        def Get_Grouped_Lotes():
                Grouped_Lotes = []
                Lotes_Zero = []
                Lotes_Negative = []
                Lotes_Positive = []
                
                for Lote in Lotes.values():
                        Lote_Balance = Lote.Get_Balance()
                        if  Lote_Balance == 0: Lotes_Zero.append(Lote)
                        elif Lote_Balance < 0: Lotes_Negative.append(Lote)
                        elif Lote_Balance > 0: Lotes_Positive.append(Lote)
                
                if Lotes_Negative and Lotes_Positive: Lotes_Positive.append(None)
                elif Lotes_Negative and not Lotes_Positive: Lotes_Negative.append(None)
                elif not Lotes_Negative and Lotes_Positive: Lotes_Positive.append()
                elif not Lotes_Negative and not Lotes_Positive: pass
                Grouped_Lotes += Lotes_Negative + Lotes_Positive + Lotes_Zero
                
                return Grouped_Lotes
        
        Order_Lotes_Inplace()
        
        for Lote in Get_Grouped_Lotes():
                def Cierre_Num_Exists(Number, Lote): return Number < len(Lote.Cierres)
                def Cobro_Num_Exists(Number, Lote): return Number < len(Lote.Cobros)
                def Get_State_Character_Lote(Balance):
                        if Balance == 0: return '🗸'
                        elif Balance < 0: return '🗴'
                        elif Balance > 0: return '⚠'
                
                if Lote is None: Lote_Rows.append([]); continue
                
                # Contruct the rows used by the lote
                Lote_Rowspan = Get_Lote_Rowspan(Lote)
                
                for Row_Number in range(Lote_Rowspan):
                        Row_Cells = [None]
                        if Row_Number == 0 : Row_Cells[0] = Lote.Number
                        
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
        Upper_Header_Row = ["Lote", "Cierres", None, None, "Cobros", None, None, None, "Estado", "Balance"]
        Lower_Header_Row = [None, "Nro. Asto", "Fecha", "Valor", "Nro. Asto", "Fecha", "Compr.", "Valor", None, None]
        
        Header_Rows.append(Upper_Header_Row)
        Header_Rows.append(Lower_Header_Row)
        
        # Load the rows to the sheet
        Sheet_Rows += Header_Rows
        Sheet_Rows += Lote_Rows
        
        # Format the worksheet
        from openpyxl.styles import Alignment, Font
        ## Load the cells into the workshit first
        for Row in Sheet_Rows: Output_Worksheet.append(Row) # Necessary.
        ##Format the dates
        for Date_Cell in (Output_Worksheet['C'] + Output_Worksheet['F']): Date_Cell.number_format = 'dd/mm/yyyy'
        Output_Worksheet.column_dimensions['C'].width = 12
        Output_Worksheet.column_dimensions['F'].width = 12
        
        ## Format the header rows
        Output_Worksheet.merge_cells('A1:A2')
        Output_Worksheet.merge_cells('I1:I2')
        Output_Worksheet.merge_cells('J1:J2')
        Output_Worksheet.merge_cells('B1:D1')
        Output_Worksheet.merge_cells('E1:H1')
        for Header_Cell in (Output_Worksheet[1] + Output_Worksheet[2]):
                Header_Cell.font = Font(bold=True)
                Header_Cell.alignment = Alignment(horizontal='center', vertical='top')

        ## Format the amounts
        for Amount_Cell in (Output_Worksheet['D'] + Output_Worksheet['H']): Amount_Cell.number_format = r'#,##0.00'
        Output_Worksheet.column_dimensions['D'].width = 11
        Output_Worksheet.column_dimensions['H'].width = 11
        
        ## Format the balances
        for Balance_Cell in Output_Worksheet['J']: Balance_Cell.number_format = r'\+#,##0.00;\-#,##0.00;#0'
        Output_Worksheet.column_dimensions['J'].width = 11
        
        ## Merge Lote Nro, State and Balance cells
        for Current_Row in Output_Worksheet.iter_rows(min_row=3, max_row=Output_Worksheet.max_row):
                Current_Lote_Number = Current_Row[0].value
                Current_Row_Number = Current_Row[0].row
                if Current_Lote_Number is None: continue
                
                Current_Lote_Rowspan = Get_Lote_Rowspan(Lotes[Current_Lote_Number])
                if Current_Lote_Rowspan == 1: continue
                
                Additional_Lote_Rows = Current_Lote_Rowspan - 1
                End_Current_Lote_Rowspan = Current_Row_Number + Additional_Lote_Rows
                Output_Worksheet.merge_cells('A{0}:A{1}'.format(Current_Row_Number, End_Current_Lote_Rowspan))
                Output_Worksheet.merge_cells('I{0}:I{1}'.format(Current_Row_Number, End_Current_Lote_Rowspan))
                Output_Worksheet.merge_cells('J{0}:J{1}'.format(Current_Row_Number, End_Current_Lote_Rowspan))

        
        # Save the file
        Output_Worksheet.title = 'Mayores contables'
        Output_Workbook.save("Test.xlsx")
Get_Output_File()