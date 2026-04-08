import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, NamedStyle
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import font
from PIL import Image, ImageTk
import pdb
import re
import traceback

Source_File_Path = ''
Destination_File_Path = ''
def Log(): pass
class Mariser_Exception(Exception): pass

def Generate_Output_File():
        global Source_File_Path
        global Destination_File_Path        
        Input_File_Path = Source_File_Path
        Output_Directory_Path = Destination_File_Path
        Input_File = None
        Input_Worksheet = None
        
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
                
        # Open the spreadsheet file
        try:
                with open(Input_File_Path, 'r') as Input_File:
                         from openpyxl import load_workbook
                         Input_Worksheet = load_workbook(Input_File.name).active

        except IOError as Error:
                Log("No se pudo encontrar o leer el archivo fuente(Ha cambiado de nombre o úbicación desde que fue seleccionado?)", 'error')
                Log('Por favor, selecciona un nuevo archivo fuente')
                Source_File_Path = ''
                return
        except Exception as Error:
                Log("Hubo un error leyendo el archivo de Excel ingresasdo.", 'error')
                Log('Por favor, selecciona un nuevo archivo fuente')
                Source_File_Path = ''
                return

        # Extract the lotes from the file
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
                                Log(f'Operacion no reconocida! La tercera celda de la fila {Row_Number + 1} no coincide con el formato de un Cierre ni de un Cobro.\nProbablemente seria una buena idea rehacer el archivo de Mayores Contables.', 'error')
                                Log('Por favor, selecciona un nuevo documento fuente')
                                global Source_File_Path
                                Source_File_Path = ''
                                raise Mariser_Exception('Fila invalida en documento fuente')
                                
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
        try: Lotes = Get_Lotes()
        except Mariser_Exception: return

        # Build the output file
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
        try: Get_Output_File()
        except Mariser_Exception: return
        
        # Save the file
        Output_File_Path = Output_Directory_Path + '/Mayores Contables.xlsx'
        
        Output_Worksheet.title = 'Mayores contables'
        try: Output_Workbook.save(Output_File_Path)
        except PermissionError:
                Log(f'No se puede puede generar el archivo `{Output_File_Path}` porque el archivo ya existe y está abierto/siendo útilizado por algún otro programa(o hay un error de permisos)', 'error')
                return
        except FileNotFoundError:
                Log(f'No se puede generar el archvio `{Output_File_Path}` porque no se encontró el directorio de destino(cambió de nombre, o se movió desde que fué seleccionado?)', 'error')
                Log('Por favor, selecciona un nuevo directorio de destino.')
                Destination_File_Path = ''
                return
        Log('Se ha generado el archivo con exito!')

class Main_Window_Mariser:
        class Rule_Units:
                """
                I am unsure of exactly what will be the size in pixels of each element of the window...
                Hoewever, I am sure of the proportion each element holds in relation to each other.
                
                So what I will do, is useing the very proportional element I have used for the sketch in
                order to measure things: the rules of the notebook.
                """
                
                def To_Pixels(Rule_Units):
                        Pixel_Equivalency = 55
                        return int( Rule_Units * Pixel_Equivalency )
        
        def __init__(self):
                import tkinter.filedialog
                import os.path
                Rule_Units = self.Rule_Units
                
                #> Program parameters
                global Source_File_Path
                global Destination_File_Path
                
                #> Execution
                # Set DPI awareness(makes the UI look HD lol)
                from sys import platform
                if platform in ('win32', 'darwin'):
                        import ctypes
                        try: # >= win 8.1
                            ctypes.windll.shcore.SetProcessDpiAwareness(2)
                        except: # win 8.0 or less
                            ctypes.windll.user32.SetProcessDPIAware()

                def Center_Toplevel(parent, toplevel):
                        def Get_TitleBar_Height():
                        # Source: Adapted from: https://stackoverflow.com/a/61492953
                                TitleBar_Height = 0
                                
                                if platform not in ('win32', 'darwin'): return 0
                                # All of the dpi thing is done above cuz it makes the whole interface look nicer :)
                                offset_y = int(toplevel.geometry().rsplit('+', 1)[-1])

                                TitleBar_Height = toplevel.winfo_rooty() - offset_y
                                return TitleBar_Height
                        
                        # Source: https://chatgpt.com/s/t_69d0924b0ed4819184d50c851c3a49c0
                        parent.update_idletasks()
                        toplevel.update_idletasks()
                        Title_Bars_Height = ( Get_TitleBar_Height() * 2 )

                        # Parent geometry
                        parent_x = parent.winfo_rootx()
                        parent_y = parent.winfo_rooty()
                        parent_width = parent.winfo_width()
                        parent_height = parent.winfo_height()

                        # Toplevel size
                        width = ( toplevel.winfo_width() + Title_Bars_Height )
                        height = ( toplevel.winfo_height() + Title_Bars_Height )

                        # Compute centered position
                        x = parent_x + ( ( parent_width - width ) // 2 )
                        y = parent_y + ( ( parent_height - height ) // 2 )

                        toplevel.geometry(f"+{x}+{y}")
                                
                def Create_Info_Window():
                        Info_Modal_Window = tkinter.Toplevel(Root_Window)
                        Info_Modal_Window.resizable(width=False, height=False)
                        Info_Modal_Window.title("Info: Mariser!")
                        
                        Info_Frame = ttk.Frame(Info_Modal_Window)
                        Info_Frame.grid(sticky='NEWS')
                        
                        # Add the Scale Image Banner
                        Logo_Size2 = ( Rule_Units.To_Pixels(2), Rule_Units.To_Pixels(2)) # (Width, height)
                        Scale_Image = Image.open(r'../Assets/Images/Window Icon.png').resize(Logo_Size2)
                        Scale_Image = ImageTk.PhotoImage(Scale_Image)
                        
                        Scale_Banner = ttk.Label(Info_Frame, image=Scale_Image) # I have no idea why is it that tkinter is such a bitch with this image thing,
                        Scale_Banner.image = Scale_Image                        # cuz only passing the image on construction doesn't work, setting it with .image
                        Scale_Banner.grid(column=1, row=1, sticky='NEWS')       # alone doesn't work, Banner['image'] = image doesn't seem to work... But I'm done with all of this.
                        
                        # Add the title
                        Title_Frame = ttk.Frame(Info_Frame)
                        Title_Frame.grid(column=3, row=1)
                        
                        Title_Image_Size = (Rule_Units.To_Pixels(1.33), Rule_Units.To_Pixels(0.89))
                        Title_Image = ImageTk.getimage(Title_Banner_Image).resize(Title_Image_Size)
                        Image_Width, Image_Height = Title_Image.size
                        Title_Image = Title_Image.crop(Title_Image.getbbox())
                        Title_Image = ImageTk.PhotoImage(Title_Image)
                        Title_Label1 = ttk.Label(Title_Frame, image=Title_Image)
                        Title_Label2 = ttk.Label(Title_Frame, text = ' es un programa para automatizar el control de los lotes.')
                        Title_Label1.image = Title_Image
                        Title_Label1.grid(column=0, row=0, sticky='WS')
                        Title_Label2.grid(column=1, row=0, sticky='WS')
                        
                        # Add subtitle
                        Subtitle_Label = ttk.Label(Title_Frame, wraplength=Rule_Units.To_Pixels(8), text='Es un regalo🎁 de Lucas para Marisa, ojalá tu vida sea un poco más fácil con esto :)')
                        Subtitle_Label.grid(column=0, columnspan=3, row=1, sticky='W')
                        
                        # Add the acknowldegments
                        Acknowledgements_Frame = ttk.Frame(Info_Frame)
                        Acknowledgements_Frame.grid(column=1, columnspan=3, row=5, sticky='W')
                        
                        Icons_Size = (Rule_Units.To_Pixels(0.75),  Rule_Units.To_Pixels(0.75))
                       
                        Python_Icon_Image = ImageTk.PhotoImage(Image.open(r'../Assets/Images/Python Logo.png').resize(Icons_Size))
                        Github_Icon_Image = ImageTk.PhotoImage(Image.open(r'../Assets/Images/Github Logo.png').resize(Icons_Size))
                        Hotel_Icon_Image = ImageTk.PhotoImage(Image.open(r'../Assets/Images/Hotel Logo.png').resize(Icons_Size))
                        Python_Label1 = ttk.Label(Acknowledgements_Frame, image=Python_Icon_Image)
                        Github_Label1 = ttk.Label(Acknowledgements_Frame, image=Github_Icon_Image)
                        Hotel_Label1 = ttk.Label(Acknowledgements_Frame, image=Hotel_Icon_Image)
                        Python_Label1.image = Python_Icon_Image
                        Github_Label1.image = Github_Icon_Image
                        Hotel_Label1.image = Hotel_Icon_Image
                        
                        Python_Label1.grid(column=1, row=0)
                        Github_Label1.grid(column=1, row=1)
                        Hotel_Label1.grid(column=1, row=2)
                        
                        Python_Label2 = ttk.Label(Acknowledgements_Frame, text='Diseñado en Python con Tkinter')
                        Github_Label2 = ttk.Label(Acknowledgements_Frame, text='Código libre a beneficio del equipo (InnerSource)')
                        Hotel_Label2 = ttk.Label(Acknowledgements_Frame, text='Para el equipo de De La Plaza Hotel')
                        
                        Python_Label2.grid(column=3, row=0, sticky='W')
                        Github_Label2.grid(column=3, row=1, sticky='W')
                        Hotel_Label2.grid(column=3, row=2, sticky='W')
                        
                        # Personal signature
                        Signature_Label = tkinter.Label(Info_Frame, text='Creado por Lucas Da Silva @ 2026', fg='gray40', font=('Helvetica', 7, 'italic'))
                        Signature_Label.grid(row=6, column=3, sticky='E')
                        
                        # Make the window modal
                        Info_Modal_Window.transient(Root_Window)
                        Center_Toplevel(Root_Window, Info_Modal_Window)
                        Info_Modal_Window.grab_set()
                        
                        # Frame Padding
                        Padding_Size = Rule_Units.To_Pixels(0.5)
                        Half_Padding = int(Padding_Size/2)
                        
                        Rightmost_Column, Downmost_Row = Info_Frame.grid_size()
                        Info_Frame.rowconfigure(0, minsize=Padding_Size)
                        Info_Frame.rowconfigure(Downmost_Row, minsize=Padding_Size)
                        Info_Frame.columnconfigure(0, minsize=Padding_Size)
                        Info_Frame.columnconfigure(Rightmost_Column, minsize=Padding_Size)
                        
                        # Element Padding
                        ## Title
                        Info_Frame.columnconfigure(2, minsize=Padding_Size)
                        ## Acknowledgments
                        Info_Frame.rowconfigure(2, minsize=Padding_Size)
                        Acknowledgements_Frame.columnconfigure(0, minsize=Half_Padding)
                        Acknowledgements_Frame.columnconfigure(2, minsize=Half_Padding)
                        
                        # Element Sizing
                        ## Acknowledgments
                        Acknowledgements_Frame.columnconfigure(0, minsize=Half_Padding)
                        ## Signature
                        Info_Frame.rowconfigure(6, minsize=Half_Padding)
                
                def Create_Exception_Dialog(parent, exception): 
                        # Made by ChatGPT, and implementing stuff from https://stackoverflow.com/a/50650817;
                        # I am more than fed up with developing GUI from scratch for this project, but I think
                        # ChatGPT's implementation is actually kinda good by default, and the one from the dude
                        # in stack overflow has features I like, so I will mix the two :)
                                                      
                        # Crear ventana modal
                        win = tkinter.Toplevel(parent)
                        win.title("Ha ocurrido un error inesperado")
                        win.transient(parent)
                        win.grab_set()

                        # Frame principal
                        frame = ttk.Frame(win, padding=10)
                        frame.grid(row=0, column=0, columnspan=2, sticky='NEWS')

                        # --- Parte superior: icono + mensaje ---
                        top_frame = ttk.Frame(frame)
                        top_frame.grid(row=0, column=0, columnspan=3, sticky='WE', pady=(0, 10))
                        ## Icono (X roja)
                        Error_Icon_Image = ImageTk.PhotoImage(Image.open('../Assets/Images/Error Icon.png').resize((40, 40)))
                        icon_label = ttk.Label(top_frame, image=Error_Icon_Image, font=("Segoe UI Emoji", 24))
                        icon_label.image = Error_Icon_Image
                        icon_label.grid(row=0, column=0, sticky='W', padx=(0, 10))
                        ## Text
                        Text_Frame = ttk.Frame(top_frame)
                        Title_Label = ttk.Label(
                            Text_Frame,
                            text="Ha ocurrido un error inesperado",
                            font=("Segoe UI", 10, "bold")
                        )
                        Description_Label = ttk.Label(
                                Text_Frame,
                                text=exception,
                                font=('Segoe UI', 8, 'italic'),
                                foreground='gray'
                        )
                        Text_Frame.grid(row=0, column=1)
                        Description_Label.grid(row=1, column=0, sticky='W')
                        Title_Label.grid(row=0, column=0, sticky='W')

                        # Set the buttons
                        ## Buttons Frame
                        Buttons_Frame = ttk.Frame(frame)
                        Buttons_Frame.grid(row=1, column=0, columnspan=2, sticky='NEWS')
                        ## Toggle button
                        def toggle():
                            nonlocal details_visible
                            if details_visible:
                                details_frame.grid_forget()
                                toggle_button.config(text="▶ Información para el programador")
                            else:
                                details_frame.grid(row=2, column=0, columnspan=2, sticky='NEWS', pady=(5, 10))
                                toggle_button.config(text="▼ Información para el programador")
                            details_visible = not details_visible
                        
                        toggle_button = ttk.Button(Buttons_Frame, text="▶ Información para el programador", command=toggle)
                        toggle_button.grid(row=0, column=0, sticky="we")
                        ## OK button
                        ok_button = ttk.Button(Buttons_Frame, text="Aceptar", command=win.destroy)
                        ok_button.grid(row=0, column=1, sticky='NEWS', padx=(10, 0))
                        
                        # Frame oculto (contenido técnico)
                        details_frame = ttk.Frame(frame)
                        details_visible = False
                        frame.columnconfigure(0, weight=0)
                        frame.columnconfigure(1, weight=0)
                        details_frame.columnconfigure(0, weight=0)
                        details_frame.columnconfigure(1, weight=0)
                        ## Texto de excepción
                        error_text = tkinter.Text(details_frame, height=10, width=59)
                        error_text.grid(row=0, column=0, columnspan=3, sticky='NEWS')
                        ## Scrollbar
                        Details_Scrollbar = ttk.Scrollbar(details_frame, command=error_text.yview)
                        Details_Scrollbar.grid(row=0, column=2, sticky='NSE')
                        error_text.config(yscrollcommand=Details_Scrollbar.set)

                        # Insertar traceback
                        tb = "".join(traceback.format_exception(type(exception), exception, exception.__traceback__))
                        error_text.insert("1.0", tb)
                        error_text.config(state="disabled")
                        
                        # Handle the size and position of the window
                        ## Set its complete width
                        Min_Width = 0
                        Min_Height = 0
                        
                        Min_Height = win.winfo_height()
                        details_visible = False; toggle()
                        details_frame.update_idletasks()
                        win.update_idletasks()
                        Min_Width = win.winfo_reqwidth()
                        win.resizable(False, False)
                        details_visible = True; toggle()
                        win.update_idletasks()

                        win.minsize(Min_Width, Min_Height)
                        frame.configure(width=Min_Width)
                        ## Center the window
                        Center_Toplevel(parent, win)
                
                def Select_Source_File():
                        File_Path = ''
                        Home_Dir = home = os.path.expanduser('~')
                        global Source_File_Path
                        
                        File_Path = tkinter.filedialog.askopenfilename(parent=Root_Window, initialdir=Home_Dir, filetypes=[('Archivo de Excel', '*.xlsx')])
                        if File_Path: Source_File_Path = File_Path ; Log(f'Se ha seleccionado el archivo `{Source_File_Path}` con exito')
                        else: pass
                        
                def Select_Output_Directory():
                        global Destination_File_Path
                        Directory_Path = ''
                        Home_Dir = home = os.path.expanduser('~')

                        Directory_Path = tkinter.filedialog.askdirectory(parent=Root_Window, initialdir=Home_Dir, mustexist=True)
                        if Directory_Path: Destination_File_Path = Directory_Path; Log(f'Se ha seleccionado el directorio `{Destination_File_Path}` como destino')
                        else: pass
                
                Root_Window = Tk()
                Root_Window.resizable(width=False, height=False)
                try:
                        Icon_Image = ImageTk.PhotoImage(Image.open(r'../Assets/Images/Window Icon.png'))
                        Root_Window.wm_iconphoto(True, Icon_Image)
                except FileNotFoundError: pass
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
                global Log
                def Log(Message, Type: Literal['cue', 'error', 'alert'] = 'cue'): 
                        if Type not in ['cue', 'error', 'alert']: raise AttributeError('Log function invoked with unrecognized mode!') # https://stackoverflow.com/a/59874453
                        Error_Font = font.nametofont('TkFixedFont').configure(weight='bold')
                        Alert_Font = Error_Font
                        Normal_Font = font.nametofont('TkFixedFont')
                        
                        Terminal_Widget.tag_configure('Error_Tag', font=Error_Font, foreground='red')
                        Terminal_Widget.tag_configure('Alert_Tag', font=Alert_Font, foreground='orange')
                        Terminal_Widget.tag_configure('Normal_Tag', font=Normal_Font)
                                                
                        match Type:
                                case 'cue': Terminal_Widget.insert(END, '> ', 'Normal_Tag')
                                case 'error': Terminal_Widget.insert(END, 'Error> ', 'Error_Tag')
                                case 'alert': Terminal_Widget.insert(END, 'Alerta> ', 'Alert_Tag')
                        
                        Terminal_Widget.insert(END, Message + '\n', 'Normal_Tag')
                
                Terminal_Widget = Text(Main_Frame, width=1, height=1) # The width and heihgt value is just so it doesn't assign itself a value by default
                Terminal_Widget.configure(font="TkFixedFont")
                Custom_Font = font.nametofont("TkFixedFont")
                Terminal_Widget.bind("<Key>", lambda e: "break") # https://stackoverflow.com/a/34811313
                Terminal_Widget.grid(column=5, row=1, rowspan=2, sticky='NWES')
                
                ## Add the scrolbar for the terminal
                Terminal_Scrollbar = ttk.Scrollbar(Main_Frame, orient=VERTICAL, command=Terminal_Widget.yview)
                Terminal_Widget['yscrollcommand'] = Terminal_Scrollbar.set
                Terminal_Scrollbar.grid(column=6, row=1, rowspan=2, sticky='NS')
                
                # Set the operation buttons
                def Refresh_Button_States():
                        if Source_File_Path: Directory_Button.config(state='normal')
                        else: Directory_Button.config(state='disabled')
                        if (Source_File_Path and Destination_File_Path): Generate_Button.config(state='normal')
                        else: Generate_Button.config(state='disabled')
                # https://www.geeksforgeeks.org/python/how-to-bind-multiple-commands-to-tkinter-button/#:~:text=provide%20horizontal%20padding-,Method%201,-%3A%20By%20using%20the
                
                Dummy_Image = tkinter.PhotoImage(width=1, height=1) # https://stackoverflow.com/a/46286221
                
                Buttons_Frame = ttk.Frame(Main_Frame)
                File_Button = tkinter.Button(Buttons_Frame, text='Seleccionar Documento Fuente', state='normal', command = lambda: [Select_Source_File(), Refresh_Button_States()],
                                             wraplength = Rule_Units.To_Pixels(2.8), bg = 'firebrick2', fg = 'floral white', relief='solid',
                                             activebackground = 'firebrick3', activeforeground = 'white smoke',                                                                                                                                                                                              
                                             font = font.Font(weight='bold', size=9),
                                             image=Dummy_Image, compound='c')
                Directory_Button = tkinter.Button(Buttons_Frame, text='Seleccionar Carpeta de Destino', state='disabled', command = lambda: [Select_Output_Directory(), Refresh_Button_States()],
                                                  wraplength = Rule_Units.To_Pixels(2.9), bg = 'yellow', fg = 'black', relief='solid',
                                                  activebackground = 'gold2', activeforeground = 'black',
                                                  font = font.Font(weight='bold', size=9),
                                                  image=Dummy_Image, compound='c')
                Generate_Button = tkinter.Button(Buttons_Frame, text='Generar!', state='disabled', command = lambda: [Generate_Output_File(), Refresh_Button_States()],
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
                
                Info_Button = tkinter.Button(Corner_Buttons_Frame, image=Info_Image, relief='groove', borderwidth=3, command=Create_Info_Window)
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
                
                # Throw the unexpected exceptions into a modal dialog rather than being silent about it
                def Handle_Unexpected_Exception(Exception_Type, Exception_Object, Exception_Trace):
                        Create_Exception_Dialog(Root_Window, Exception_Object)
                        traceback.print_exception(Exception_Type, Exception_Object, Exception_Trace) # Print exception as well
                        
                Root_Window.report_callback_exception = Handle_Unexpected_Exception
                
                # Run the mainloop
                Root_Window.mainloop()
                
Main_Window_Mariser()