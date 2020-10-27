
from tkinter import *
from tkinter import filedialog,Tk,messagebox
import tkinter.ttk as ttk
from PIL import ImageTk, Image
import datetime, time, os, shutil
import pyautogui
from os import listdir
from sys import platform
from fpdf import FPDF
from pynput.keyboard import Key,Controller
import threading
Simkey = Controller()


if platform == "darwin": #OSX
    OS_Entry_factor = 1
    OS_Text_w_factor = 1
    OS_Text_h_factor = 1
    OS_Sup_key = '\x7f'
elif platform == "win32": #Windows
    OS_Entry_factor = 1.53
    OS_Text_w_factor = 0.84
    OS_Text_h_factor = 0.9375
    OS_Sup_key = '\x08'


Window_w = 850
Window_h = 500
Window_r = 14
Window_c = 10
font=('Arial',-25,'bold')
font_1 = ('Arial',-20)
font_2 = ('Arial',-15)
font_3 = ('Arial',-18)
font_4 = ('Arial',-20,'bold')

Bin_path = "Bin/"   #fixe
Data_path = "Data/"
OUTPUT_path = "OUTPUT/"
PDF_TEMPLATE_path ="Data/PDF_TEMPLATE/" #fixe

PDF_header_name = "PDF_header.jpg"
PDF_footer_name = "PDF_footer.jpg"
PDF_header_img = None
PDF_footer_img = None


fen = Tk()
fen.resizable(False, False)
fen.title('SLG Order Creator')
fen.geometry(str(Window_w)+'x'+str(Window_h))



def Global_vars():
    global go_flag,processed_data,Gad_bg,Gad_btn,Gad_text,Gad_title,G_state,User_config_list,Parts_database_list,Current_Order_list,MB_flag,date,A_state,E_state,S_state,Add_bg,Add_btn,Add_text,Add_title,Export_bg,Export_btn,Export_text,Export_title,Settings_bg,Settings_btn,Settings_text,Settings_title,Menu_X_border,Menu_Y_border,Screen_w,Screen_h,btn_list

    date = str(datetime.datetime.now().day) + '-' + str(datetime.datetime.now().month) + '-' + str(datetime.datetime.now().year)
    Menu_X_border = 2*Window_w/Window_c
    Menu_Y_border = 2*Window_h/Window_r
    Screen_w = Window_w-Menu_X_border
    Screen_h = Window_h-Menu_Y_border
    A_state = True
    E_state = False
    S_state = False
    G_state = False
    Add_bg = None
    Add_btn = None
    Add_text = None
    Add_title = None
    Export_bg = None
    Export_btn = None
    Export_text = None
    Export_title = None
    Settings_bg = None
    Settings_btn = None
    Settings_text = None
    Settings_title = None
    Gad_bg = None
    Gad_btn = None
    Gad_text = None
    Gad_title = None

    btn_list = []
    MB_flag = False

    Current_Order_list = []
    Parts_database_list = []
    User_config_list =[]

    processed_data = []
    go_flag = False

Global_vars()

def Screen_init():
    global Window, Logo_img,Screen
    Window = Canvas(fen, width = Window_w , height = Window_h, bg = 'grey98',bd=0, highlightthickness=0)
    Window.place(x=0,y=0,anchor = "nw")
    Window.create_rectangle(0,0,Menu_X_border,Window_h, fill = 'grey25',width =0)
    Window.create_rectangle(Menu_X_border,0,Window_w,Menu_Y_border, fill = 'grey90',width =0)
    Logo_img = ImageTk.PhotoImage(Image.open(Bin_path+"SLG_logo.jpg"))
    Logo = Label(fen,image = Logo_img,width=168,height=168,bd=0, highlightthickness=0)
    Logo.place(x=0,y=0,anchor="nw")
    Window.create_line(0, 246.25, 2 * Window_w / Window_c, 246.25, width=0.5, fill = 'grey18')
    Window.create_line(0, 322.5, 2 * Window_w / Window_c, 322.5, width=0.5, fill = 'grey18')
    Window.create_line(0, 398.75, 2 * Window_w / Window_c, 398.75, width=0.5, fill='grey18')
    Window.create_line(0, 475, 2 * Window_w / Window_c, 475, width=0.5, fill='grey18')
    Window.create_text(65,487.5, text = '©  M. Michotte', font = ('Arial',-12),fill='grey85',anchor='center')
    Window.create_text(750, Window_h/Window_r, text=date, font=font, fill='grey55', anchor='center')

    Screen = Canvas(fen, width = Window_w-Menu_X_border , height = Window_h-Menu_Y_border, bg = 'grey98',bd=0, highlightthickness=0)
    Screen.place(x=Menu_X_border,y=Menu_Y_border,anchor='nw')


    Add_Btn()
    Export_Btn()
    Gad_Btn()
    Settings_Btn()

def Add_Btn():
    global Add_btn_img, A_state,Add_btn,Add_text,Add_title, Add_bg

    if A_state == False:
        A_state = True
        Add_btn_img = ImageTk.PhotoImage(file=Bin_path + "Add_0.png")
        color = 'white'
        try:
            Window.delete(Add_title)
            Window.delete(Add_bg)
        except: pass
    else:
        A_state = False
        Add_bg = Window.create_rectangle(0, 168, Menu_X_border,246.25, fill='grey18', width=0)
        Add_btn_img = ImageTk.PhotoImage(file=Bin_path + "Add_1.png")
        color = 'chartreuse3'
        Add_title = Window.create_text(Menu_X_border+10, Menu_Y_border/2, text='Add part', font=font, anchor='w')
        Add_Part_screen()
    try:
        Window.delete(Add_btn)
        Window.delete(Add_text)
    except:
        pass
    Add_btn = Window.create_image(25,208.125, image=Add_btn_img)
    Add_text = Window.create_text(52,208.125,text='Add part', font=font,fill=color,anchor='w')

def Export_Btn():
    global Export_btn_img,E_state,Export_btn,Export_text,Export_title,Export_bg

    if E_state == False:
        E_state = True
        Export_btn_img = ImageTk.PhotoImage(file=Bin_path + "Export_0.png")
        color = 'white'
        try:
            Window.delete(Export_title)
            Window.delete(Export_bg)
        except: pass
    else:
        E_state = False
        Export_bg = Window.create_rectangle(0, 246.25, Menu_X_border, 322.5, fill='grey18', width=0)
        Export_btn_img = ImageTk.PhotoImage(file=Bin_path + "Export_1.png")
        color = 'chartreuse3'
        Export_title = Window.create_text(Menu_X_border+10, Menu_Y_border/2, text='Export to PDF', font=font, anchor='w')
        Export_screen()
    try:
        Window.delete(Export_btn)
    except:
        pass
    Export_btn = Window.create_image(25, 284.375, image=Export_btn_img)
    Export_text = Window.create_text(52, 284.375, text='Export', font=font, fill=color,anchor='w')

def Gad_Btn():
    global Gad_btn_img,G_state,Gad_btn,Gad_text,Gad_title,Gad_bg
    if G_state == False:
        G_state = True
        Gad_btn_img = ImageTk.PhotoImage(file=Bin_path + "GAD_0.png")
        color = 'white'
        try:
            Window.delete(Gad_title)
            Window.delete(Gad_bg)
        except: pass
    else:
        G_state = False
        Gad_bg = Window.create_rectangle(0, 322.5, Menu_X_border, 398.75, fill='grey18', width=0)
        Gad_btn_img = ImageTk.PhotoImage(file=Bin_path + "GAD_1.png")
        color = 'chartreuse3'
        Gad_title = Window.create_text(Menu_X_border+10, Menu_Y_border/2, text='Comptabilité', font=font, anchor='w')
        Gad_screen()
    try:
        Window.delete(Gad_btn)
    except:
        pass
    Gad_btn = Window.create_image(25, 360.625, image=Gad_btn_img)
    Gad_text = Window.create_text(52, 360.625, text='Compta', font=font, fill=color,anchor='w')

def Settings_Btn():
    global Settings_btn_img,S_state,Settings_btn,Settings_text,Settings_title,Settings_bg
    if S_state == False:
        S_state = True
        Settings_btn_img = ImageTk.PhotoImage(file=Bin_path + "Settings_0.png")
        color = 'white'
        try:
            Window.delete(Settings_title)
            Window.delete(Settings_bg)
        except: pass
    else:
        S_state = False
        Settings_bg = Window.create_rectangle(0, 398.75, Menu_X_border, 475, fill='grey18', width=0)
        Settings_btn_img = ImageTk.PhotoImage(file=Bin_path + "Settings_1.png")
        color = 'chartreuse3'
        Settings_title = Window.create_text(Menu_X_border+10, Menu_Y_border/2, text='Settings', font=font, anchor='w')
        Settings_screen()
    try:
        Window.delete(Settings_btn)
    except:
        pass
    Settings_btn = Window.create_image(25, 436.875, image=Settings_btn_img)
    Settings_text = Window.create_text(52, 436.875, text='Settings', font=font, fill=color,anchor='w')



def Add_Part_screen():
    global Search_EB,New_part_btn,Delete_btn,Save_btn,Match_box,Current_Order_box
    Clear_Screen()
    Screen.create_text(Screen_w/5-20,40 , text='Search :', font=font_4, anchor='center')
    Search_EB = Entry(Screen, width=int(45*OS_Entry_factor),relief = 'groove',bd=1,highlightthickness=0)
    Search_EB.place(x=Screen_w*3/5-30, y=40, anchor='center')


    Screen.create_text(Screen_w-240, 90, text="Current list:", font=font_4, anchor='center')
    Current_Order_box = ttk.Treeview(Screen,columns=("Quantity","Number","Name"))
    style = ttk.Style()
    style.configure("Treeview", font=font_2,rowheight=23,relief = 'ridge',borderwidth = 0.5)
    style.configure("Treeview.Heading", font=font_3)
    Current_Order_box['show']='headings'
    Current_Order_box.heading('#1',text="#")
    Current_Order_box.column('#1',width=60,anchor='center')
    Current_Order_box.heading("#2",text="P. Number")
    Current_Order_box.column('#2',width=120,anchor='center')
    Current_Order_box.heading("#3",text="P. Name")
    Current_Order_box.column('#3',width=220,anchor='center')
    Current_Order_box.place(x=450,y=235,anchor='center')

    Screen.create_text(125, 90, text='Comments:', font=font_4, anchor='center')
    Comment_txt = Text(Screen, width=int(30*OS_Text_w_factor), height=int(16*OS_Text_h_factor), bd=1, relief="groove",highlightthickness=0,padx=5,pady=5,wrap =WORD)
    Comment_txt.place(x=125, y=Screen_h-194, anchor='center')

    def NewPart():
        NewP = Toplevel()
        NewP.geometry('430x70%+d%+d' % (int(fen.winfo_geometry().split('+')[1]) + Window_w / 2 - 215,int(fen.winfo_geometry().split('+')[2]) + Window_h / 2 - 30))
        NewP.title("New part:")

        Label(NewP, text="Part number:", font=font_4, fg='Chartreuse3').grid(row=0, column=0, sticky='w')
        Number_EB = Entry(NewP, width=int(10*OS_Entry_factor))
        Number_EB.grid(row=0, column=1, sticky='w')
        Number_EB.configure(justify=CENTER)
        Number_EB.focus_set()

        Label(NewP, text="Part name:", font=font_4, fg='Chartreuse3').grid(row=1, column=0, sticky='w')
        Name_EB = Entry(NewP, width=int(30*OS_Entry_factor))
        Name_EB.grid(row=1, column=1, sticky='w')
        Name_EB.configure(justify=CENTER)

        def save(event):
            if messagebox.askyesno("INFO", "Save into database?"):
                Number = Number_EB.get().upper()
                Name = Name_EB.get().upper()
                part = Number + " - " + Name
                Parts_database_list.append(part)
                UPDATE_FILES('Write')
                NewP.destroy()
            else:
                pass

        NewP.bind('<Return>', save)

    def Delete_item():
        del_quest = messagebox.askquestion("WARNING", "DELETE this item from current list?", icon="warning")
        if del_quest == 'yes':
            try:
                part = Current_Order_box.item(Current_Order_box.focus()).get('values')
                part = [str(part[0]), str(part[1]), part[2]]
                if part in Current_Order_list:
                    del (Current_Order_list[Current_Order_list.index(part)])
                Current_Order_box.delete(Current_Order_box.selection()[0])
                UPDATE_FILES('Write')
            except:
                pass
        else:
            pass

    def Save_item():
        UPDATE_FILES('Write')
        Update_Cur_Order()

    def Edit_item(event):
        EditPart = Toplevel()
        EditPart.geometry('440x140%+d%+d' % (int(fen.winfo_geometry().split('+')[1]) + Window_w / 2 - 220,int(fen.winfo_geometry().split('+')[2]) + Window_h / 2 - 70))
        EditPart.title("Edit:")

        part = Current_Order_box.item(Current_Order_box.focus()).get('values')
        part_inlist = [str(part[0]), str(part[1]), part[2]]

        Label(EditPart, text="Quantity:", font=font_4,fg='Chartreuse3').grid(row=0,column=0,sticky='w')
        Quant_EB = Entry(EditPart, width=int(5*OS_Entry_factor))
        Quant_EB.grid(row=0,column=1,sticky='w')
        Quant_EB.configure(justify=CENTER)
        Quant_EB.insert(0, part[0])
        Quant_EB.focus_set()

        Label(EditPart, text="Part number:", font=font_4,fg='Chartreuse3').grid(row=1,column=0,sticky='w')
        Number_EB = Entry(EditPart, width=int(10*OS_Entry_factor))
        Number_EB.grid(row=1,column=1,sticky='w')
        Number_EB.configure(justify=CENTER)
        Number_EB.insert(0, part[1])

        Label(EditPart, text="Part description:", font=font_4,fg='Chartreuse3').grid(row=2,column=0,sticky='w')
        Name_EB = Text(EditPart, width=int(35*OS_Text_w_factor), height=int(4*OS_Text_h_factor), bd=1, relief="groove",highlightthickness=0,padx=5,pady=5,wrap =WORD)
        Name_EB.grid(row=2,column=1,sticky='w')
        Name_EB.configure(font=('Arial',12))
        Name_EB.insert(1.0, part[2])


        def save(event):
            Quantity = Quant_EB.get()
            Number = Number_EB.get().upper()
            Name =Name_EB.get(1.0, 'end-1c')
            Name = Name.replace('\n', "")
            Name = Name.strip()
            for i,part in enumerate(Current_Order_list):
                if part == part_inlist:
                    Current_Order_list[i] = [Quantity, Number, Name]

            UPDATE_FILES('Write')
            Update_Cur_Order()
            EditPart.destroy()

        EditPart.bind('<Return>', save)

    def Save_comment(event):
        if Comment_txt.get(1.0, 'end-1c') == "":
            pass
        else:
            text = Comment_txt.get(1.0, 'end-1c')
            text = text.replace('\n', "")
            text = text.strip()
            text = ['N/A', 'N/A' , text]
            Current_Order_list.append(text)
            UPDATE_FILES('Write')
            Update_Cur_Order()
            Comment_txt.delete(1.0,END)
            Search_EB.focus_set()
            return 'break'


    Current_Order_box.bind('<Double-ButtonRelease-1>', Edit_item)
    Comment_txt.bind('<Return>', Save_comment)

    New_part_btn = Custom_Btn(Screen,Screen_w * 1 / 6 - 50, Screen_h - 30, 'New Part',NewPart)
    New_part_btn.NOTactive()
    Delete_btn = Custom_Btn(Screen,Screen_w * 3 / 6 - 50, Screen_h - 30, 'Delete', Delete_item)
    Delete_btn.NOTactive()
    Save_btn = Custom_Btn(Screen,Screen_w * 5 / 6 - 50, Screen_h - 30, 'Save', Save_item)
    Save_btn.NOTactive()

    Update_Cur_Order()
    Search_EB.focus_set()

def Export_screen():
    global Current_Order_box,OK_img,Current_Order_list
    Clear_Screen()
    to_export_list = []

    Screen.create_text(Screen_w/2,20, text="Current list:", font=font_4, anchor='center')
    Current_Order_box = ttk.Treeview(Screen, columns=("Quantity", "Number", "Name"))
    style = ttk.Style()
    style.configure("Treeview", font=font_2, rowheight=30, relief='ridge', borderwidth=0.5)
    style.configure("Treeview.Heading", font=font_3)
    Current_Order_box.heading('#0', text="Export")
    Current_Order_box.column('#0', width=70, anchor='e')
    Current_Order_box.heading('#1', text="#")
    Current_Order_box.column('#1', width=60, anchor='center')
    Current_Order_box.heading("#2", text="P. Number")
    Current_Order_box.column('#2', width=120, anchor='center')
    Current_Order_box.heading("#3", text="P. Name")
    Current_Order_box.column('#3', width=220, anchor='center')
    Current_Order_box.place(x=Screen_w/2, y=Screen_h/2-15, anchor='center')
    Update_Cur_Order_EXPORT()

    def Delete_item():
        del_quest = messagebox.askquestion("WARNING", "DELETE this item from today's list?", icon="warning")
        if del_quest == 'yes':
            try:
                part = Current_Order_box.item(Current_Order_box.focus()).get('values')
                part = [str(part[0]), str(part[1]), part[2]]
                if part in Current_Order_list:
                    del (Current_Order_list[Current_Order_list.index(part)])
                if part in to_export_list:
                    del (to_export_list[to_export_list.index(part)])
                Current_Order_box.delete(Current_Order_box.selection()[0])
                UPDATE_FILES('Write')
            except:
                pass
        else:
            pass

    def Export_all():
        global Current_Order_list
        nonlocal to_export_list
        to_export_list =[]
        for part in Current_Order_list:
            to_export_list.append(part)
        Export_PDF(to_export_list)
        Current_Order_list =[]
        Update_Cur_Order_EXPORT()
        UPDATE_FILES('Write')

    def Export_selection():
        nonlocal to_export_list
        for e in to_export_list:
            if e in Current_Order_list:
                del(Current_Order_list[Current_Order_list.index(e)])
        Export_PDF(to_export_list)
        to_export_list = []
        Update_Cur_Order_EXPORT()
        UPDATE_FILES('Write')


    OK_img = ImageTk.PhotoImage(file=Bin_path + "OK.png")
    def choose(event):
        nonlocal to_export_list
        global Current_Order_box
        flag = Current_Order_box.item(Current_Order_box.focus()).get('image')
        part = Current_Order_box.item(Current_Order_box.focus()).get('values')
        part = [str(part[0]), str(part[1]), part[2]]
        Current_Order_box.delete(*Current_Order_box.get_children())
        for item in Current_Order_list:
            if item == part:
                if str(OK_img) in flag:
                    del(to_export_list[to_export_list.index(item)])
                    Current_Order_box.insert('', END, image='', values=(part[0], part[1], part[2]))
                else:
                    to_export_list.append(part)
                    Current_Order_box.insert('', END, image=OK_img, values=(part[0], part[1], part[2]),tags='checked')
            elif item in to_export_list:
                Current_Order_box.insert('', END, image=OK_img, values=(item[0], item[1], item[2]), tags='checked')
            else:
                Current_Order_box.insert('', END, image='', values=(item[0], item[1], item[2]))
        Current_Order_box.tag_configure('checked',foreground='Chartreuse3')


    Current_Order_box.bind('<Double-ButtonRelease-1>', choose)

    Delete_btn = Custom_Btn(Screen, Screen_w * 1 / 6 - 50, Screen_h - 30, 'Delete', Delete_item)
    Delete_btn.NOTactive()
    All_btn = Custom_Btn(Screen, Screen_w * 3 / 6 - 50, Screen_h - 30, 'All items', Export_all)
    All_btn.NOTactive()
    SEL_btn = Custom_Btn(Screen, Screen_w * 5 / 6 - 50, Screen_h - 30, 'Selection', Export_selection)
    SEL_btn.NOTactive()

def Settings_screen():
    global PDF_header_name,PDF_footer_name,Data_path,OUTPUT_path

    Clear_Screen()
    def Browse(x):
        if x == 'data':
            dir = fen.directory = filedialog.askdirectory()
            if dir != "":
                Data_EB.delete(0, END)
                Data_EB.insert(0, dir)
        elif x == 'output':
            dir = fen.directory = filedialog.askdirectory()
            if dir != "":
                Output_EB.delete(0,END)
                Output_EB.insert(0, dir)
        elif x == 'pdf_temp_h':
            dir = fen.directory = filedialog.askopenfilename(filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
            if dir != "":
                PDF_temp_H_EB.delete(0,END)
                PDF_temp_H_EB.insert(0, dir)
        elif x == 'pdf_temp_f':
            dir = fen.directory = filedialog.askopenfilename(filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
            if dir != "":
                PDF_temp_F_EB.delete(0,END)
                PDF_temp_F_EB.insert(0, dir)

    Screen.create_text(40,15 , text='Choose Data files folder', font=font_1, anchor='w')
    Data_EB = Entry(Screen, width=int(55*OS_Entry_factor), relief='groove', highlightthickness=0)
    Data_EB.place(x=Screen_w/2- 50, y=45, anchor='center')
    Data_EB.insert(0,Data_path)
    Browse_1 = Custom_Btn(Screen,Screen_w - 115,45,'Browse',command= lambda: Browse('data'))
    Browse_1.NOTactive()

    Screen.create_text(40, 75, text='Cooose PDF output folder:', font=font_1, anchor='w')
    Output_EB = Entry(Screen, width=int(55*OS_Entry_factor), relief='groove', highlightthickness=0)
    Output_EB.place(x=Screen_w/2- 50, y=105, anchor='center')
    Output_EB.insert(0,OUTPUT_path)
    Browse_2 = Custom_Btn(Screen, Screen_w - 115, 105, 'Browse', command=lambda: Browse('output'))
    Browse_2.NOTactive()

    Screen.create_text(40, 135+20, text='Use other PDF template (jpg):', font=font_1, anchor='w')
    Screen.create_text(40, 165+20, text='Header:', font=font_2, anchor='w')
    PDF_temp_H_EB = Entry(Screen, width=int(55*OS_Entry_factor), relief='groove', highlightthickness=0)
    PDF_temp_H_EB.place(x=Screen_w/2- 50, y=190+20, anchor='center')
    PDF_temp_H_EB.insert(0,PDF_header_name)
    Browse_3 = Custom_Btn(Screen, Screen_w - 115, 190+20, 'Browse', command=lambda: Browse('pdf_temp_h'))
    Browse_3.NOTactive()

    Screen.create_text(40, 220+20, text='Footer:', font=font_2, anchor='w')
    PDF_temp_F_EB = Entry(Screen, width=int(55*OS_Entry_factor), relief='groove', highlightthickness=0)
    PDF_temp_F_EB.place(x=Screen_w / 2 - 50, y=245+20, anchor='center')
    PDF_temp_F_EB.insert(0, PDF_footer_name)
    Browse_4 = Custom_Btn(Screen, Screen_w - 115, 245+20, 'Browse', command=lambda: Browse('pdf_temp_f'))
    Browse_4.NOTactive()

    def Save_Config():
        global Data_path,OUTPUT_path,PDF_header_name, PDF_footer_name

        if Data_EB.get() == Data_path:
            New_Data_path = Data_path
        else:
            New_Data_path = Data_EB.get()+"/"
            try:
                shutil.copytree("Data", New_Data_path + "Data")
                New_Data_path = New_Data_path+"Data/"
            except FileExistsError:
                New_Data_path = New_Data_path + "Data/"
        if Output_EB.get() == OUTPUT_path:
            New_OUTPUT_path = OUTPUT_path
        else:
            New_OUTPUT_path = Output_EB.get()+"/"

        try:
            New_PDF_header_name = os.path.split(PDF_temp_H_EB.get())[1]
            New_PDF_footer_name = os.path.split(PDF_temp_F_EB.get())[1]
            shutil.copy(PDF_temp_H_EB.get(), PDF_TEMPLATE_path)
            shutil.copy(PDF_temp_F_EB.get(), PDF_TEMPLATE_path)
        except:
            New_PDF_header_name = PDF_header_name
            New_PDF_footer_name = PDF_footer_name

        Data_path = New_Data_path
        OUTPUT_path = New_OUTPUT_path
        PDF_header_name = New_PDF_header_name
        PDF_footer_name = New_PDF_footer_name

        UPDATE_FILES('Write')
        UPDATE_FILES('Read')
        messagebox.showinfo("Information", "Settings saved")

    Save_btn = Custom_Btn(Screen, Screen_w/2-50, Screen_h - 30, 'Save', Save_Config)
    Save_btn.NOTactive()

def Gad_screen():
    global pyautogui
    from pywinauto.application import Application, win32defines
    from pywinauto.win32functions import SetForegroundWindow, ShowWindow
    Clear_Screen()
    te =[]
    for doc in listdir( Data_path + 'Past_Orders/'):
        te.append(doc)
    if len(te) == 0:
        messagebox.showerror("Error", "You must have placed at least 1 order to acces this menu!")
    else:
        past_files = []
        Shipping_price = 0
        Debited_amount = 0
        data_to_gad_0 = []

        Gad_box = ttk.Treeview(Screen, columns=("Quantity", "Number", "Name", "Price/pc"))
        style = ttk.Style()
        style.configure("Treeview", font=font_2, rowheight=28, relief='ridge', borderwidth=0.5)
        style.configure("Treeview.Heading", font=font_3)
        Gad_box['show'] = 'headings'
        Gad_box.heading('#1', text="#")
        Gad_box.column('#1', width=50, anchor='center')
        Gad_box.heading("#2", text="P. Number")
        Gad_box.column('#2', width=110, anchor='center')
        Gad_box.heading("#3", text="P. Name")
        Gad_box.column('#3', width=140, anchor='center')
        Gad_box.heading("#4", text="Price/pc")
        Gad_box.column('#4', width=80, anchor='center')
        Gad_box.place(x=210, y=235, anchor='center')

        def set_gad_box(*args):
            nonlocal data_to_gad_0
            chosen_file = []
            try:
                with open(Data_path + "Past_Orders/%s"%(v.get()), 'r') as file:
                    Parts_file = file.read().splitlines()
                    for part in Parts_file:
                        part = part.split(';')
                        part = [part[0], part[1], part[2]]
                        chosen_file.append(part)
                    file.close()
                Gad_box.delete(*Gad_box.get_children())
                data_to_gad_0 = []
                for e in chosen_file:
                    e.append("N/C")
                    data_to_gad_0.append(e)
                    Gad_box.insert('', 'end', values=(e[0], e[1], e[2],e[3]))
                FDP_EB.delete(0, END)
                Deb_EB.delete(0, END)
                FDP_EB.insert(0, 0.00)
                Deb_EB.insert(0, 0.00)
            except:
                None


        def sorting(L):
            splitup = L.split('-')
            day = splitup[0]
            month = splitup[1]
            year = splitup[2].split('_')[0]
            art_P1 = splitup[2].split('_')
            art = art_P1[1].split('.')[0]
            return int(year), int(month), int(day), int(art)

        Screen.create_text(Screen_w / 5 - 20, 40, text='Choose order :', font=font_4, anchor='center')
        for doc in listdir( Data_path + 'Past_Orders/'):
            past_files.append(doc)
        past_files.sort(key=sorting, reverse=True)
        v = StringVar()
        v.trace("w", set_gad_box)
        v.set('...')
        om = OptionMenu(Screen, v, *past_files)
        om.place(x=300,y=40,anchor='center')
        om.config(font=font_3, relief='groove', width=12,highlightthickness=0)
        om['menu'].config(font=font_3)


        Screen.create_text(Screen_w * 5 / 6 - 20, 170, text='Shipping costs:', font=font_1, anchor='center')
        FDP_EB = Entry(Screen, width=int(10 * OS_Entry_factor), relief='groove', highlightthickness=0,justify='center')
        FDP_EB.place(x=Screen_w * 5 / 6 - 20, y=200, anchor='center')
        FDP_EB.insert(0, 0.00)

        Screen.create_text(Screen_w * 5 / 6 - 20, 230, text='Debited amount (€):', font=font_1, anchor='center')
        Deb_EB = Entry(Screen, width=int(10 * OS_Entry_factor), relief='groove', highlightthickness=0,justify='center')
        Deb_EB.place(x=Screen_w * 5 / 6 - 20, y=260, anchor='center')
        Deb_EB.insert(0, 0.00)

        price_showlbl = Screen.create_text(Screen_w * 5 / 6 - 20, 300, text='Price (£):', font=font_1, anchor='center')
        price_show = Screen.create_text(Screen_w * 5 / 6 - 20, 330, text='0.0', font=font_1,anchor='center')

        def Edit_it(event):
            EditPart = Toplevel()
            EditPart.geometry('440x170%+d%+d' % (int(fen.winfo_geometry().split('+')[1]) + Window_w / 2 - 220,int(fen.winfo_geometry().split('+')[2]) + Window_h / 2 - 70))
            EditPart.title("Edit:")

            part = Gad_box.item(Gad_box.focus()).get('values')
            part_inlist = [str(part[0]), str(part[1]), part[2],part[3]]

            Label(EditPart, text="Quantity:", font=font_4, fg='Chartreuse3').grid(row=0, column=0, sticky='w')
            Quant_EB = Entry(EditPart, width=int(5 * OS_Entry_factor))
            Quant_EB.grid(row=0, column=1, sticky='w')
            Quant_EB.configure(justify=CENTER)
            Quant_EB.insert(0, part[0])

            Label(EditPart, text="Part number:", font=font_4, fg='Chartreuse3').grid(row=1, column=0, sticky='w')
            Number_EB = Entry(EditPart, width=int(10 * OS_Entry_factor))
            Number_EB.grid(row=1, column=1, sticky='w')
            Number_EB.configure(justify=CENTER)
            Number_EB.insert(0, part[1])

            Label(EditPart, text="Part description:", font=font_4, fg='Chartreuse3').grid(row=2, column=0, sticky='w')
            Name_EB = Text(EditPart, width=int(35 * OS_Text_w_factor), height=int(4 * OS_Text_h_factor), bd=1,relief="groove", highlightthickness=0, padx=5, pady=5, wrap=WORD)
            Name_EB.grid(row=2, column=1, sticky='w')
            Name_EB.configure(font=('Arial', 12))
            Name_EB.insert(1.0, part[2])

            Label(EditPart, text="Price:", font=font_4, fg='Chartreuse3').grid(row=3, column=0, sticky='w')
            Price_EB = Entry(EditPart, width=int(10 * OS_Entry_factor))
            Price_EB.grid(row=3, column=1, sticky='w')
            Price_EB.configure(justify=CENTER)
            Price_EB.insert(0, part[3])
            Price_EB.focus_set()
            Price_EB.selection_range(0, END)



            def save(event):
                nonlocal data_to_gad_0
                Quantity = Quant_EB.get()
                Number = Number_EB.get().upper()
                Name = Name_EB.get(1.0, 'end-1c')
                Name = Name.replace('\n', "")
                Name = Name.strip()
                Price = Price_EB.get()
                for i, part in enumerate(data_to_gad_0):
                    if part[1] == part_inlist[1]:  #TODO replaced 2 by 1
                        data_to_gad_0[i] = [Quantity, Number, Name, Price]
                Gad_box.delete(*Gad_box.get_children())
                for e in data_to_gad_0:
                    Gad_box.insert('', 'end', values=(e[0], e[1], e[2], e[3]))
                EditPart.destroy()
                calc_total()
            EditPart.bind('<Return>', save)

        Gad_box.bind('<Double-ButtonRelease-1>', Edit_it)

        def set_FDP():
            nonlocal Shipping_price
            Shipping_price = float(FDP_EB.get())

        def set_deb():
            nonlocal Debited_amount
            Debited_amount = float(Deb_EB.get())

        def calc_total():
            nonlocal price_show,price_showlbl
            total_pounds = 0
            for e in data_to_gad_0:
                if e[3] != "N/C":
                    total_pounds += float(e[3]) * int(e[0])
            set_FDP()
            total_pounds += Shipping_price
            Screen.delete(price_showlbl)
            Screen.delete(price_show)
            price_showlbl = Screen.create_text(Screen_w * 5 / 6 - 20, 300, text='Price (£):', font=font_1, anchor='center')
            price_show = Screen.create_text(Screen_w * 5 / 6 - 20, 330, text='%0.2f'%total_pounds, font=font_1,anchor='center')


        def calc_price():
            global processed_data,go_flag
            go_flag = False
            set_FDP()
            set_deb()
            if Shipping_price == 0 or Debited_amount == 0:
                messagebox.showerror("Error", "You must enter the shipping costs and the debited amout!")
            else:
                for e in data_to_gad_0:
                    if e[3] != "N/C":
                        go_flag = True
                    else:
                        go_flag = False
                        break

            if go_flag == True:
                processed_data = data_to_gad_0
                total_price = 0
                for e in processed_data:
                    total_price += float(e[3])*int(e[0])
                for e in processed_data:
                    e[3] = (float(e[3])+float(e[3])*Shipping_price/total_price)*(Debited_amount/(total_price+Shipping_price))
            else:
                messagebox.showerror("Error","You must either enter the purchase price of each item or delete the item!")

        def bring_foreground():
            if w.has_style(win32defines.WS_MINIMIZE):  # if minimized
                ShowWindow(w.wrapper_object(), 9)  # restore window state
            else:
                SetForegroundWindow(w.wrapper_object())  # bring to front

        def Data_process():
            global w, conf
            # once data is processed:
            start_flag = False
            try:
                Garage_app = "C:\GAD Garage\\Garage.exe"
                app = Application().connect(path=Garage_app)
                w = app.top_window()
                start_flag = True
            except:
                start_flag = False
                messagebox.showerror("ERROR", "Gad Garage is not open")
            if start_flag:
                if messagebox.askyesno("WARNING",'Avez-vous:\n - créé un nouveau bon de livraison?\n - double-cliqué sur la case "Rechercher" de la première ligne?'):
                    bring_foreground()
                    time.sleep(1)
                    conf = None
                    for d in processed_data:
                        if conf != 'ABORT':
                            Gad_Garage_input(d[1], d[0], d[3], '', '', 0)
                    if conf != 'ABORT':
                        messagebox.showinfo("Done", "Task completed")
                    else:
                        messagebox.showinfo("Aborted", "Task aborted")
                else:
                    pass
            else:
                pass

        def Gad_Garage_input(P_num, P_quant, P_price, P_fournisseur, P_N_commande, P_TVA):
            global pyautogui,conf
            pyautogui.typewrite(str(P_num))  # Rechercher
            pyautogui.keyDown('return')
            conf = pyautogui.confirm(text='Choisissez le bon article et appuyez sur "Enter"', title='CHECK',buttons=['OK', 'ABORT'])
            if conf != 'ABORT':
                time.sleep(1)
                pyautogui.keyDown('tab')
                pyautogui.keyDown('tab')
                pyautogui.typewrite(str(P_quant))  # Quantity
                pyautogui.keyDown('tab')
                pyautogui.typewrite(str(P_price))  # price
                pyautogui.keyDown('tab')
                pyautogui.typewrite(str(P_fournisseur))  # fournisseur
                pyautogui.keyDown('tab')
                pyautogui.typewrite(str(P_N_commande))  # Num commande fournisseur
                pyautogui.keyDown('tab')
                pyautogui.typewrite(str(P_TVA))  # TVA
                pyautogui.keyDown('tab')  # switch to next line

        def GAD():
            calc_price()
            if go_flag ==True:
                Data_process()


        def add():
            NewP = Toplevel()
            NewP.geometry('440x170%+d%+d' % (int(fen.winfo_geometry().split('+')[1]) + Window_w / 2 - 215,int(fen.winfo_geometry().split('+')[2]) + Window_h / 2 - 30))
            NewP.title("New part:")

            Label(NewP, text="Quantity:", font=font_4, fg='Chartreuse3').grid(row=0, column=0, sticky='w')
            Quant_EB = Entry(NewP, width=int(5 * OS_Entry_factor))
            Quant_EB.grid(row=0, column=1, sticky='w')
            Quant_EB.configure(justify=CENTER)
            Quant_EB.focus_set()
            Quant_EB.selection_range(0, END)

            Label(NewP, text="Part number:", font=font_4, fg='Chartreuse3').grid(row=1, column=0, sticky='w')
            Number_EB = Entry(NewP, width=int(10 * OS_Entry_factor))
            Number_EB.grid(row=1, column=1, sticky='w')
            Number_EB.configure(justify=CENTER)

            Label(NewP, text="Part description:", font=font_4, fg='Chartreuse3').grid(row=2, column=0, sticky='w')
            Name_EB = Text(NewP, width=int(35 * OS_Text_w_factor), height=int(4 * OS_Text_h_factor), bd=1,relief="groove", highlightthickness=0, padx=5, pady=5, wrap=WORD)
            Name_EB.grid(row=2, column=1, sticky='w')
            Name_EB.configure(font=('Arial', 12))

            Label(NewP, text="Price:", font=font_4, fg='Chartreuse3').grid(row=3, column=0, sticky='w')
            Price_EB = Entry(NewP, width=int(10 * OS_Entry_factor))
            Price_EB.grid(row=3, column=1, sticky='w')
            Price_EB.configure(justify=CENTER)


            def save(event):
                if messagebox.askyesno("INFO", "Add new part?"):
                    Quantity = Quant_EB.get()
                    Number = Number_EB.get().upper()
                    Name = Name_EB.get(1.0, 'end-1c')
                    Name = Name.replace('\n', "")
                    Name = Name.strip()
                    Price = Price_EB.get()
                    P_ex = False
                    for e in data_to_gad_0:
                        if Number == e[1]:
                            messagebox.showinfo("Information", "Part is already in today's list!")
                            NewP.destroy()
                            P_ex = True
                            break
                    if P_ex == False:
                        data_to_gad_0.append([Quantity, Number, Name, Price])
                        Gad_box.delete(*Gad_box.get_children())
                        for e in data_to_gad_0:
                            Gad_box.insert('', 'end', values=(e[0], e[1], e[2], e[3]))
                        NewP.destroy()
                        try:
                            calc_total()
                        except:
                            pass
                    else:
                        pass
                else:
                    pass

            NewP.bind('<Return>', save)

        def delete():
            del_quest = messagebox.askquestion("WARNING", "DELETE this item?", icon="warning")
            if del_quest == 'yes':
                try:
                    part = Gad_box.item(Gad_box.focus()).get('values')
                    part = [str(part[0]), str(part[1]), part[2],str(part[3])]
                    if part in data_to_gad_0:
                        del (data_to_gad_0[data_to_gad_0.index(part)])
                        Gad_box.delete(Gad_box.selection()[0])
                    calc_total()
                except:
                    pass
            else:
                pass

        Add_P_btn = Custom_Btn(Screen, Screen_w * 5 / 6 - 65, 70, 'Add part', add)
        Add_P_btn.NOTactive()
        Del_P_btn = Custom_Btn(Screen, Screen_w * 5 / 6 - 65, 110, 'Delete', delete)
        Del_P_btn.NOTactive()
        Gad_P_btn = Custom_Btn(Screen,Screen_w * 5 / 6 - 50, Screen_h - 30, 'To GAD', GAD)
        Gad_P_btn.NOTactive()


def Check_searchbar(event):
    global Match_box, MB_flag
    #print(repr(event))
    try:
        typing = Search_EB.get()
        typing = typing.upper()

        if typing != "":
            if MB_flag == False:
                MB_flag = True
                Match_box = Listbox(Screen, width=int(45*OS_Entry_factor), height=6, relief='ridge', bd=1, highlightthickness=0)
                Match_box.place(x=Screen_w*3/5-30, y=50, anchor='n')
                Match_box.bind("<Double-Button-1>", Match_to_today)
            match_list = []

            for i in Parts_database_list:
                match_flag = 0
                if i.split(' - ')[0].startswith(typing):
                    match_list.append(i)
                else:
                    typing_S = typing.split()
                    for word in range(len(typing_S)):
                        if typing_S[word] in i.split(' - ')[1]:
                            match_flag += 1
                        else:
                            pass
                    if match_flag == len(typing_S):
                         match_list.append(i)
            Match_box.delete(0, END)
            for match in match_list:
                Match_box.insert(END, match)
        else:
            Match_box.destroy()
            MB_flag = False
    except Exception as e:
        print(str(e))

def Match_to_today(event):
    global MB_flag,Match_box
    part = Match_box.get(ACTIVE).split(" - ")
    flag = False

    for e in Current_Order_list:
        if part[0] in e[1]:
            messagebox.showinfo("Information", "Part is already in today's list!")
            Match_box.destroy()
            Search_EB.focus_set()
            MB_flag = False
            flag = True
            break
    if flag == False:
        Quantity = Toplevel()
        Quantity.geometry('150x30%+d%+d'%(int(fen.winfo_geometry().split('+')[1])+Window_w/2-75,int(fen.winfo_geometry().split('+')[2])+Window_h/2-15))
        Quantity.title("Quantity:")


        Quant_EB = Entry(Quantity, width=5)
        Quant_EB.pack()
        Quant_EB.configure(justify=CENTER)
        Quant_EB.insert(0, 1)
        Quant_EB.focus_set()
        Quant_EB.selection_range(0, END)

        def save(event):
            nonlocal part
            quantity = Quant_EB.get()
            part = [quantity, part[0], part[1]]
            Current_Order_list.append(part)
            UPDATE_FILES('Write')
            Update_Cur_Order()
            Match_box.destroy()
            Quantity.destroy()
            Search_EB.delete(0,END)
            Search_EB.focus_set()
            Simkey.press(OS_Sup_key)
            Simkey.release(OS_Sup_key)
            MB_flag = False

        Quantity.bind('<Return>', save)


def Update_Cur_Order():
    global Current_Order_box
    Current_Order_box.delete(*Current_Order_box.get_children())
    for part in Current_Order_list:
        Current_Order_box.insert('', 'end', values=(part[0], part[1],part[2]))

def Update_Cur_Order_EXPORT():
    global Current_Order_box,img
    Current_Order_box.delete(*Current_Order_box.get_children())
    for part in Current_Order_list:
        Current_Order_box.insert('', 'end', values=(part[0], part[1],part[2]))


def Export_PDF(list):
    pdf = FPDF()
    to_pdf = []
    page = 0
    max_lines_p0 = 20
    max_lines_pn = 30
    rect_1_w = 35
    rect_2_w = 105
    rect_3_w = 25
    rect_h = 8
    rect_y_ofset = -1.5
    start_x = (210-(rect_1_w+rect_2_w+rect_3_w))/2
    start_y = 80
    NP_ofset = -60
    f = 5
    def generate_txt():
        i = 0
        path = Data_path + 'Past_Orders/%s_%s.txt' % (date, i)
        while os.path.isfile(path):
            i += 1
            path = Data_path + 'Past_Orders/%s_%s.txt' % (date, i)

        with open(path, 'w') as file:
            for part in list:
                file.write(part[0] + ";" + part[1] + ";" + part[2])
                file.write('\n')
            file.close()
    generate_txt()

    def DataProcess(lis):
        nonlocal to_pdf
        def Text_splitter(text):
            splitted_text = []
            max_len = 35
            NL = float(len(text) / max_len)
            i = 0
            while NL > 1:
                if text[max_len] == " ":
                    splitted_text.append(text[0:max_len + 1])
                    text = text[max_len + 1:]
                    NL = float(len(text) / max_len)
                else:
                    while text[max_len - i] != " ":
                        i += 1
                        pos = max_len - i
                    splitted_text.append(text[0:pos + 1])
                    text = text[pos + 1:]
                    NL = float(len(text) / max_len)
            splitted_text.append(text)
            return splitted_text
        def Num_splitter(text):
            splitted_text = []
            max_len = 11
            NL = float(len(text) / max_len)
            i = 0
            while NL > 1:
                splitted_text.append(text[0:max_len])
                text = text[max_len:]
                NL = float(len(text) / max_len)

            splitted_text.append(text)
            return splitted_text

        for val in lis:
            if len(val[2]) > 35:
                val[2] = Text_splitter(val[2])
            if len(val[1]) > 11:
                val[1] = Num_splitter(val[1])
            to_pdf.append(val)
        dataWrite(to_pdf)


    def Structur(page):
        nonlocal pdf
        if page == 0:
            pdf.add_page('p', 'A4')
            pdf.set_font('Arial', 'B', 14.0)
            pdf.set_margins(24, 24)
            pdf.image(PDF_TEMPLATE_path + PDF_header_name, 0, 0, 210, 71)     #header
            pdf.image(PDF_TEMPLATE_path + PDF_footer_name, 0,297-19 , 210, 19)  # footer
            pdf.set_xy(130, 34)
            pdf.write(f, date)
            pdf.set_line_width(1)
            pdf.rect(start_x, start_y+rect_y_ofset, rect_1_w, rect_h)
            pdf.rect(start_x + rect_1_w, start_y+rect_y_ofset, rect_2_w, rect_h)
            pdf.rect(start_x+rect_1_w+rect_2_w, start_y+rect_y_ofset, rect_3_w, rect_h)
            pdf.set_xy(start_x+2, start_y)
            pdf.write(f, "Part number")
            pdf.set_xy(start_x+rect_1_w+35, start_y)
            pdf.write(f, "Part name")
            pdf.set_xy(start_x+rect_1_w+rect_2_w+1, start_y)
            pdf.write(f, "Quantity")
            pdf.set_font('Arial', '', 12.5)
            pdf.set_line_width(0.4)
        else:
            pdf.add_page('p', 'A4')
            pdf.set_font('Arial', '', 12.5)
            pdf.set_margins(24, 24)
            pdf.image(PDF_TEMPLATE_path + PDF_footer_name, 0,297-19 , 210, 19)  # footer

    def dataWrite(data):
        nonlocal page
        Structur(page)
        line = 0
        written_lines = 0
        rect_w_mult = 1
        for p in data:
            line += rect_w_mult
            if type(p[2]) == type([]) and type(p[1]) == type([]):
                if len(p[2]) > len(p[1]):
                    written_lines = written_lines + len(p[2])
                    rect_w_mult = len(p[2])
                else:
                    written_lines = written_lines + len(p[1])
                    rect_w_mult = len(p[1])
            elif type(p[2]) == type([]):
                written_lines = written_lines + len(p[2])
                rect_w_mult = len(p[2])
            elif type(p[1]) == type([]):
                written_lines = written_lines + len(p[1])
                rect_w_mult = len(p[1])
            else:
                written_lines = written_lines + 1
                rect_w_mult = 1
            if page == 0 and written_lines > max_lines_p0:
                page += 1
                if type(p[2]) == type([]) and type(p[1]) == type([]):
                    if len(p[2]) > len(p[1]):
                        written_lines = len(p[2])
                    else:
                        written_lines = len(p[1])
                elif type(p[2]) == type([]):
                    written_lines = len(p[2])
                elif type(p[1]) == type([]):
                    written_lines = len(p[1])
                else:
                    written_lines = 1
                line = 0
                Structur(page)
                pdf.rect(start_x, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_1_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_2_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w + rect_2_w, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_3_w, rect_h*rect_w_mult)
                pdf.set_xy(start_x + 2, start_y + NP_ofset + rect_h * line)
                try:
                    pdf.write(f, p[1])
                except AttributeError:
                    for l in range(len(p[1])):
                        pdf.write(f, p[1][l])
                        pdf.set_xy(start_x + 2, start_y + NP_ofset + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + 2, start_y + NP_ofset + rect_h * line)
                try:
                    pdf.write(f, p[2])
                except AttributeError:
                    for l in range(len(p[2])):
                        pdf.write(f, p[2][l])
                        pdf.set_xy(start_x + rect_1_w + 2, start_y + NP_ofset + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + rect_2_w + 1, start_y + NP_ofset + rect_h * line)
                pdf.write(f, p[0])

            elif page > 0 and written_lines > max_lines_pn:
                page += 1
                if type(p[2]) == type([]) and type(p[1]) == type([]):
                    if len(p[2]) > len(p[1]):
                        written_lines = len(p[2])
                    else:
                        written_lines = len(p[1])
                elif type(p[2]) == type([]):
                    written_lines = len(p[2])
                elif type(p[1]) == type([]):
                    written_lines = len(p[1])
                else:
                    written_lines = 1
                line = 0
                Structur(page)
                pdf.rect(start_x, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_1_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_2_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w + rect_2_w, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_3_w, rect_h*rect_w_mult)
                pdf.set_xy(start_x + 2, start_y + NP_ofset + rect_h * line)
                try:
                    pdf.write(f, p[1])
                except AttributeError:
                    for l in range(len(p[1])):
                        pdf.write(f, p[1][l])
                        pdf.set_xy(start_x + 2, start_y + NP_ofset + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + 2, start_y + NP_ofset + rect_h * line)
                try:
                    pdf.write(f, p[2])
                except AttributeError:
                    for l in range(len(p[2])):
                        pdf.write(f, p[2][l])
                        pdf.set_xy(start_x + rect_1_w + 2, start_y + NP_ofset + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + rect_2_w + 1, start_y + NP_ofset + rect_h * line)
                pdf.write(f, p[0])

            elif page == 0:
                pdf.rect(start_x, start_y+ rect_y_ofset+rect_h*line , rect_1_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w, start_y + rect_y_ofset+rect_h*line , rect_2_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w + rect_2_w, start_y + rect_y_ofset+rect_h*line , rect_3_w, rect_h*rect_w_mult)
                pdf.set_xy(start_x + 2, start_y+rect_h*line)
                try:
                    pdf.write(f, p[1])
                except AttributeError:
                    for l in range(len(p[1])):
                        pdf.write(f, p[1][l])
                        pdf.set_xy(start_x + 2, start_y + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + 2, start_y+rect_h*line )
                try:
                    pdf.write(f, p[2])
                except AttributeError:
                    for l in range(len(p[2])):
                        pdf.write(f, p[2][l])
                        pdf.set_xy(start_x + rect_1_w + 2, start_y + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + rect_2_w + 1, start_y+rect_h*line)
                pdf.write(f, p[0])

            else:
                pdf.rect(start_x, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_1_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_2_w, rect_h*rect_w_mult)
                pdf.rect(start_x + rect_1_w + rect_2_w, start_y + NP_ofset + rect_y_ofset + rect_h * line, rect_3_w,rect_h*rect_w_mult)
                pdf.set_xy(start_x + 2, start_y + NP_ofset + rect_h * line)
                try:
                    pdf.write(f, p[1])
                except AttributeError:
                    for l in range(len(p[1])):
                        pdf.write(f, p[1][l])
                        pdf.set_xy(start_x + 2, start_y + NP_ofset + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + 2, start_y + NP_ofset + rect_h * line)
                try:
                    pdf.write(f, p[2])
                except AttributeError:
                    for l in range(len(p[2])):
                        pdf.write(f, p[2][l])
                        pdf.set_xy(start_x + rect_1_w + 2, start_y + NP_ofset + rect_h * line + rect_h*(l+1))
                pdf.set_xy(start_x + rect_1_w + rect_2_w + 1, start_y + NP_ofset + rect_h * line)
                pdf.write(f, p[0])


        Generate_pdf()

    def Generate_pdf():
        i = 0
        path = OUTPUT_path + '%s_%s.pdf' % (date, i)
        while os.path.isfile(path):
            i += 1
            path = OUTPUT_path + '%s_%s.pdf' % (date, i)
        pdf.output(path, 'F')

    DataProcess(list)

def UPDATE_FILES(how):
    global Current_Order_list, Parts_database_list,User_config_list,OUTPUT_path,Data_path,PDF_header_name,PDF_footer_name,PDF_footer_img,PDF_header_img

    if how == 'Read':
        with open("Data/User_config.txt",'r') as file:
            Parts_file = file.read().splitlines()
            for part in Parts_file:
                part = part.split(';')
                if part[0] == 'OUTPUT_path':
                    OUTPUT_path = part[1]
                elif part[0] == 'Data_path':
                    Data_path = part[1]
                elif part[0] == 'PDF_header':
                    PDF_header_name = part[1]
                    PDF_header_img = ImageTk.PhotoImage(file=PDF_TEMPLATE_path+PDF_header_name)
                elif part[0] == 'PDF_footer':
                    PDF_footer_name = part[1]
                    PDF_footer_img = ImageTk.PhotoImage(file=PDF_TEMPLATE_path+PDF_footer_name)
            file.close()
        try:
            with open(Data_path+"Parts_database.txt", 'r') as file:
                Parts_database_list =[]
                Parts_file = file.read().splitlines()
                for part in Parts_file:
                    part = part.split(";")
                    part = part[0] + " - " + part[1]
                    Parts_database_list.append(part)
                file.close()
            with open(Data_path+"Current_Order.txt",'r') as file:
                Current_Order_list=[]
                Parts_file = file.read().splitlines()
                for part in Parts_file:
                    part = part.split(';')
                    part = [part[0],part[1],part[2]]
                    Current_Order_list.append(part)
                file.close()
        except FileNotFoundError:
            Data_path ="Data/"
            OUTPUT_path = "OUTPUT/"
            with open("Data/Parts_database.txt", 'r') as file:
                Parts_database_list =[]
                Parts_file = file.read().splitlines()
                for part in Parts_file:
                    part = part.split(";")
                    part = part[0] + " - " + part[1]
                    Parts_database_list.append(part)
                file.close()
            with open("Data/Current_Order.txt",'r') as file:
                Current_Order_list=[]
                Parts_file = file.read().splitlines()
                for part in Parts_file:
                    part = part.split(';')
                    part = [part[0],part[1],part[2]]
                    Current_Order_list.append(part)
                file.close()

    elif how == 'Write':
        with open("Data/User_config.txt",'w') as file:
            file.write("OUTPUT_path" + ";" + OUTPUT_path +"\n")
            file.write("Data_path" + ";" + Data_path + "\n")
            file.write("PDF_header" + ";" + PDF_header_name + "\n")
            file.write("PDF_footer" + ";" + PDF_footer_name + "\n")
            file.close()
        with open(Data_path+"Parts_database.txt",'w') as file:
            for part in Parts_database_list:
                part = part.split(" - ")
                file.write(part[0]+";"+part[1])
                file.write('\n')
            file.close()

        with open(Data_path+"Current_Order.txt",'w') as file:
            for part in Current_Order_list:
                file.write(part[0]+";"+part[1]+";"+part[2])
                file.write('\n')
            file.close()


class Custom_Btn:
    """ possible types: Save, Browse, Delete, Export, New_part"""
    global btn_list

    def __init__(self,parent,posx,posy,type,command):
        self.name = self
        self.state = True
        self.parent = parent
        self.posx = posx
        self.posy = posy
        self.type = type
        self.command = command

        btn_list.append([self.posx, self.posy, self.name, self.state])

    def NOTactive(self):
        self.bg = self.parent.create_rectangle(self.posx-20, self.posy-15,self.posx+110 , self.posy+15, fill='grey90', width=0.5, outline='grey70')
        self.img = ImageTk.PhotoImage(file=Bin_path + self.type + "_0.png")
        self.btn = self.parent.create_image(self.posx, self.posy, image=self.img)
        self.text = self.parent.create_text(self.posx+25, self.posy, text=self.type, font=font_3, fill='black', anchor='w')

    def active(self):
        self.bg = self.parent.create_rectangle(self.posx-20, self.posy-15,self.posx+110 , self.posy+15, fill='grey40', width=0.5, outline='grey70')
        self.img = ImageTk.PhotoImage(file=Bin_path + self.type + "_1.png")
        self.btn = self.parent.create_image(self.posx, self.posy, image=self.img)
        self.text = self.parent.create_text(self.posx+25, self.posy, text=self.type, font=font_3, fill='chartreuse3', anchor='w')


    def exec(self):
        self.command()

def Clear_Screen():
    global btn_list
    try:
        Screen.delete('all')
        btn_list=[]
        for label in Screen.place_slaves():
            label.place_forget()
    except:
        pass

def Check_click(event):
    global A_state,E_state,S_state,G_state
    if 0 < event.x < Menu_X_border:
        if 170 < event.y < 246 and A_state == True:
            E_state = False
            S_state = False
            G_state = False
            Add_Btn()
            Export_Btn()
            Gad_Btn()
            Settings_Btn()
        elif 247 < event.y < 322 and E_state == True:
            A_state = False
            S_state = False
            G_state = False
            Add_Btn()
            Export_Btn()
            Gad_Btn()
            Settings_Btn()
        elif 323 < event.y < 398 and G_state == True:
            A_state = False
            E_state = False
            S_state = False
            Add_Btn()
            Export_Btn()
            Gad_Btn()
            Settings_Btn()
        elif 399 < event.y < 475 and S_state == True:
            A_state = False
            E_state = False
            G_state = False
            Add_Btn()
            Export_Btn()
            Gad_Btn()
            Settings_Btn()

def Check_click_hold(event):
    if 'ButtonPress' in str(event):
        for b in btn_list:
            if b[0]-20<event.x<b[0]+110:
                if b[1]-15<event.y<b[1]+15:
                    b[2].active()
    else:
        for b in btn_list:
            if b[0]-20<event.x<b[0]+110:
                if b[1]-15<event.y<b[1]+15:
                    b[2].NOTactive()
                    b[2].exec()


def Server_updater():
    while quit_flag == False:
        previous_Order_list = Current_Order_list
        previous_Database_list = Parts_database_list
        UPDATE_FILES('Read')
        if previous_Order_list == Current_Order_list and previous_Database_list == Parts_database_list:
            pass
        else:
            Update_Cur_Order()
        #print('updating')
        time.sleep(2)

quit_flag = False
def quit():
    global quit_flag
    quit_flag = True
    fen.destroy()

fen.protocol("WM_DELETE_WINDOW", quit)
threading.Thread(target=Server_updater).start()
UPDATE_FILES('Read')
Screen_init()
fen.bind('<Key>', Check_searchbar)
Window.bind('<Button-1>',Check_click)
Screen.bind('<Button-1>',Check_click_hold)
Screen.bind('<ButtonRelease-1>',Check_click_hold)
fen.mainloop()