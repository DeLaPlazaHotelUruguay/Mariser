import openpyxl
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import font
from PIL import Image, ImageTk
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
                        
# Lotes = Get_Lotes()

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
                        if Balance == 0: return 'P' # '🗸'
                        elif Balance < 0: return 'O' # '🗴'
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
        from openpyxl.styles import Alignment, Font, NamedStyle
        ## Load the cells into the workshit first # I realized it says "shit", I find it funnilly fitting out of all the struggle I have had with all of this formatting thing hahahah
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

        ## Format Lote Number cells and Balance cells
        Top_Right_Alignment = Alignment(horizontal='right', vertical='top')
        for Lote_Num_Cell in Output_Worksheet['A']:
                if not Lote_Num_Cell.row > 2: continue
                Lote_Num_Cell.alignment = Top_Right_Alignment
        for Balance_Cell in Output_Worksheet['J']:
                if not Balance_Cell.row > 2: continue
                Balance_Cell.alignment = Top_Right_Alignment
                
        ## Format State cells
        for State_Cell in Output_Worksheet['I']:
                if not State_Cell.row > 2: continue
                State_Cell.alignment = Alignment(horizontal='center', vertical='center')
                match State_Cell.value:
                        case 'P': State_Cell.font = Font(name='Wingdings 2', size='12', color='00008000', bold=True) # '🗸'
                        case 'O': State_Cell.font = Font(name='Wingdings 2', size='12', color='00FF0000', bold=True) # '🗴'
                        case '⚠': State_Cell.font = Font(name='Calibri', size='12', color='00FF9900', bold=True)
        
        # Save the file
        Output_Worksheet.title = 'Mayores contables'
        # Output_Workbook.save("Test.xlsx")

class Main_Window_Mariser:
        class Rule_Units:
                """
                I am unsure of exactly what will be the size in pixels of each element of the window...
                Hoewever, I am sure of the proportion each element holds in relation to each other.
                
                So what I will do, is useing the very proportional element I have used for the sketch in
                order to measure things: the rules of the notebook.
                """
                
                def To_Pixels(Rule_Units):
                        Pixel_Equivalency = 45
                        return int( Rule_Units * Pixel_Equivalency )

        def __init__(self):
                Rule_Units = self.Rule_Units

                Root_Window = Tk()
                Root_Window.resizable(width=False, height=False)
                Root_Window.title("Mariser!")
                
                # Tailor the content frame
                Main_Frame = ttk.Frame(Root_Window)
                Main_Frame.grid(sticky=(N, W, E, S))
                
                # Set the Logo
                Logo_Size = ( Rule_Units.To_Pixels(2), Rule_Units.To_Pixels(2)) # (Width, height)
                Logo_Image = Image.open(r'../Assets/Images/Hotel Logo.png').resize(Logo_Size)
                Logo_Image = ImageTk.PhotoImage(Logo_Image)
                Logo_Label = ttk.Label(Main_Frame)
                Logo_Label['image'] = Logo_Image
                Logo_Label.grid(column=1, row=1, sticky='NWES')
                                
                # Set the Title and Subtitle Banner
                Text_Frame = ttk.Frame(Main_Frame, borderwidth=10, relief='ridge')
                Text_Frame.grid(column=3, row=1, sticky='NSWE')
                
                ## Append the Header Banner
                Title_Banner_Size = (Rule_Units.To_Pixels(4), Rule_Units.To_Pixels(0.8))
                Title_Banner_Image = Image.open(r'../Assets/Images/Header, Berlin Sans Demi Bold Banner.png').resize(Title_Banner_Size)
                Title_Banner_Image = ImageTk.PhotoImage(Title_Banner_Image)
                Title_Banner_Label = ttk.Label(Text_Frame, image=Title_Banner_Image)
                Title_Banner_Label.grid(column=0, row=0, sticky='N')
                ## Append the Subtitle
                Subtitle_Label = ttk.Label(
                                                Text_Frame,
                                                text = 'Herramienta de control de mayores contables.',
                                                font=font.Font(size=9),
                                                wraplength=Rule_Units.To_Pixels(4.5),
                                                justify='left'
                                           )
                Subtitle_Label.grid(column=0, row=3, padx=(Rule_Units.To_Pixels(0.3), 0))
                
                # Set the terminal textbox
                Terminal_Widget = Text(Main_Frame, width=1, height=1) # The width and heihgt value is just so it doesn't assign itself a value by default
                Terminal_Widget.configure(font="TkFixedFont")
                Custom_Font = font.nametofont("TkFixedFont")
                Terminal_Widget.grid(column=5, row=1, rowspan=2, sticky='NWES')
                
                ## Add the scrolbar for the terminal
                Terminal_Scrollbar = ttk.Scrollbar(Main_Frame, orient=VERTICAL, command=Terminal_Widget.yview)
                Terminal_Widget['yscrollcommand'] = Terminal_Scrollbar.set
                Terminal_Scrollbar.grid(column=6, row=1, rowspan=2, sticky='NS')
                
                # Set the operation buttons
                Dummy_Image = tkinter.PhotoImage(width=1, height=1) # https://stackoverflow.com/a/46286221
                
                Buttons_Frame = ttk.Frame(Main_Frame)
                File_Button = tkinter.Button(Buttons_Frame, text='Seleccionar Documento Fuente',
                                             wraplength = 125, bg = 'firebrick2', fg = 'floral white', relief='solid',
                                             activebackground = 'firebrick3', activeforeground = 'white smoke',                                                                                                                                                                                              
                                             font = font.Font(weight='bold', size=9),
                                             image=Dummy_Image, compound='c')
                Directory_Button = tkinter.Button(Buttons_Frame, text='Seleccionar Carpeta de Destino',
                                                  wraplength = 130, bg = 'yellow', fg = 'black', relief='solid',
                                                  activebackground = 'gold2', activeforeground = 'black',
                                                  font = font.Font(weight='bold', size=9),
                                                  image=Dummy_Image, compound='c')
                Generate_Button = tkinter.Button(Buttons_Frame, text='Generar!',
                                                 bg = 'DarkOliveGreen3', fg = 'white', relief='solid',
                                                 activebackground = 'DarkOliveGreen4', activeforeground = 'snow',
                                                 font = font.Font(weight='bold', size=12),
                                                 image=Dummy_Image, compound='c')
                                                 
                Buttons_Frame.grid(column=5, row=6, columnspan=2, sticky='NWES')
                File_Button.grid(column=0, row=0, sticky='NWES')
                Directory_Button.grid(column=2, row=0, sticky='NWES')
                Generate_Button.grid(column=0, row=2, columnspan=3, sticky='NWES')
                
                # Add the lower left corner buttons
                Corner_Buttons_Frame = ttk.Frame(Main_Frame, relief='solid')
                Corner_Buttons_Frame.grid(column=1, row=11, columnspan=2, sticky='NWES')
                
                Icons_Size = ( Rule_Units.To_Pixels(1), Rule_Units.To_Pixels(1)) # (Width, height)
                Exit_Image = Image.open(r'../Assets/Images/Exit Icon.png').resize(Icons_Size)
                Info_Image = Image.open(r'../Assets/Images/Info Icon.png').resize(Icons_Size)
                Info_Image = ImageTk.PhotoImage(Info_Image)
                Exit_Image = ImageTk.PhotoImage(Exit_Image)
                
                Info_Button = tkinter.Button(Corner_Buttons_Frame, image=Info_Image, relief='groove', borderwidth=3)
                Exit_Button = tkinter.Button(Corner_Buttons_Frame, image=Exit_Image, relief='groove', borderwidth=3, command=Root_Window.destroy)
                
                Info_Button.grid(column=1, row=1, sticky='NWES')
                Exit_Button.grid(column=3, row=1, sticky='NWES')
                
                # Space management
                Padding_Size = Rule_Units.To_Pixels(1)
                Half_Padding = int(Padding_Size / 2)
                
                ## Frame padding
                Rightmost_Column, Downmost_Row = Main_Frame.grid_size()
                Main_Frame.rowconfigure(0, minsize=Padding_Size)
                Main_Frame.rowconfigure(Downmost_Row, minsize=Padding_Size)
                Main_Frame.columnconfigure(0, minsize=Padding_Size)
                Main_Frame.columnconfigure(Rightmost_Column, minsize=Padding_Size)
                
                # Element padding
                ## Multielement
                Main_Frame.columnconfigure(2, minsize=Half_Padding)
                Main_Frame.columnconfigure(4, minsize=Padding_Size)
                Main_Frame.rowconfigure(5, minsize=Half_Padding)
                Main_Frame.rowconfigure(10, minsize=Half_Padding)
                ## Operation Buttons
                Buttons_Frame.columnconfigure(1, minsize=Half_Padding)
                Buttons_Frame.rowconfigure(1, minsize=Half_Padding)
                Buttons_Frame.columnconfigure(1, weight=0)
                ## Corner Buttons
                Corner_Buttons_Frame.columnconfigure(0, minsize=Rule_Units.To_Pixels(0.16))
                Corner_Buttons_Frame.columnconfigure(2, minsize=Rule_Units.To_Pixels(0.16))
                Corner_Buttons_Frame.columnconfigure(4, minsize=Rule_Units.To_Pixels(0.16))
                Corner_Buttons_Frame.rowconfigure(0, minsize=Rule_Units.To_Pixels(0.25))
                Corner_Buttons_Frame.rowconfigure(2, minsize=Rule_Units.To_Pixels(0.25))

                
                # Element size
                ## Multielement
                Main_Frame.rowconfigure(1, minsize=Rule_Units.To_Pixels(2))
                ## Logo Banner
                Main_Frame.columnconfigure(1, minsize=Rule_Units.To_Pixels(2))
                ## Text Banner
                Main_Frame.columnconfigure(3, minsize=Rule_Units.To_Pixels(5))
                ## Terminal and its scrollbar
                Main_Frame.rowconfigure(2, minsize=Rule_Units.To_Pixels(2))
                Main_Frame.columnconfigure(5, minsize=Rule_Units.To_Pixels(7.7))
                Main_Frame.columnconfigure(6, minsize=Rule_Units.To_Pixels(0.3))
                ## Operation Buttons
                Main_Frame.rowconfigure(6, minsize=Rule_Units.To_Pixels(3))
                Buttons_Frame.columnconfigure(0, minsize=Rule_Units.To_Pixels(3.75))
                Buttons_Frame.columnconfigure(2, minsize=Rule_Units.To_Pixels(3.75))
                Buttons_Frame.rowconfigure(0, minsize=Rule_Units.To_Pixels(1.25))
                Buttons_Frame.rowconfigure(2, minsize=Rule_Units.To_Pixels(1.25))
                ## Corner Buttons
                Main_Frame.rowconfigure(11, minsize=Rule_Units.To_Pixels(1.5))
                Corner_Buttons_Frame.columnconfigure(1, minsize=Rule_Units.To_Pixels(1))
                Corner_Buttons_Frame.columnconfigure(3, minsize=Rule_Units.To_Pixels(1))
                Corner_Buttons_Frame.rowconfigure(1, minsize=Rule_Units.To_Pixels(1))
                
                Root_Window.mainloop()
                
Main_Window_Mariser()