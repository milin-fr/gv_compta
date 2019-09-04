import tkinter
from tkinter import Label, Button, Entry, Checkbutton, OptionMenu, Canvas, Frame, Toplevel
from tkinter.ttk import Combobox
from tkinter import StringVar, IntVar
from tkinter.messagebox import showinfo
import openpyxl
from openpyxl import Workbook, load_workbook
import datetime

'''Below module allows us to interact with Windows files.'''
import os

'''below 3 lines allows script to check the directory where it is executed, so it knows where to crete the excel file. I copied the whole block from stack overflow'''
abspath = os.path.abspath(__file__)
current_directory = os.path.dirname(abspath)
os.chdir(current_directory)

LIST_TYPE_OF_WORK = ["work1", "work2"]
LIST_OF_COMPANIES = ["company1", "company2"]

def get_file_names_in_script_directory():
    file_names = []
    for root, dirs, files in os.walk(current_directory):
        for filename in files:
            file_names.append(filename)
    return file_names


def get_existing_excel_names():
    '''Each excel file stands for type of work'''
    existing_excel_names = []
    for file_name in get_file_names_in_script_directory():
        if "GV compta" in file_name:
            existing_excel_names.append(file_name)
    return existing_excel_names


def get_date_yyyy_mm_dd():
    return str(datetime.date.today())


def get_current_year():
    #current_month = datetime.datetime.now().strftime('%B') #extracts the name of the current month
    current_year = datetime.datetime.now().strftime('%Y') #extracts the current year
    return str(current_year)


def generate_the_excel_file_name_with_current_year_in_name():
    return (get_current_year() + " sandbox.xlsx")


def check_if_excel_file_is_there():
    return generate_the_excel_file_name_with_current_year_in_name() in get_existing_excel_names()


def create_excel_file_if_it_was_not_there():
    if not check_if_excel_file_is_there():
        wb = Workbook()
        ws = wb.active
        ws.title = "first"
        ws.cell(row=1, column=1, value="name")
        ws.cell(row=1, column=2, value="last date")
        ws.cell(row=1, column=3, value="last payment")
        ws.cell(row=1, column=4, value="total payment")
        wb.save(generate_the_excel_file_name_with_current_year_in_name())


def find_the_next_empty_row(ws):
    row_index = 1
    cell_to_check = ws.cell(row = row_index, column = 1)
    while cell_to_check.value != None:
        row_index = row_index + 1
        cell_to_check = ws.cell(row = row_index, column = 1)
    return row_index


def create_missing_sheet(name, excel_workbook):
    print("creating missing sheet " + name)
    if name not in excel_workbook.sheetnames:
        excel_workbook.create_sheet(name)
        ws = excel_workbook[name]
        ws.cell(row=1, column=1, value="last date")
        ws.cell(row=1, column=2, value="last payment")

def button_pressed():
    pass


def open_work_selection_window():
    window_start_new_work = Toplevel()
    x = main_window_of_gui.winfo_x()
    y = main_window_of_gui.winfo_y()
    w = main_window_of_gui.winfo_width()
    h = main_window_of_gui.winfo_height()  
    
    window_start_new_work.geometry("%dx%d+%d+%d" % (w, h, x, y))
    window_start_new_work.title("Type de traveaux")
    window_start_new_work.wm_attributes("-topmost", 1)
    
    combo_work_selection_window = Combobox(window_start_new_work, values = LIST_TYPE_OF_WORK)
    combo_work_selection_window.grid(column=0, row=0, columnspan=2)
    combo_work_selection_window.current(0)

    button_confirm_work_selection = Button(window_start_new_work, text="Selectioner", width=20, height=3, command=lambda: open_company_selection_window(combo_work_selection_window))
    button_confirm_work_selection.grid(column=0, row=1)

    button_cancel_work_selection = Button(window_start_new_work, text="Annuler", width=20, height=3, command=lambda: cancel_selection_window(window_start_new_work))
    button_cancel_work_selection.grid(column=1, row=1)

def open_company_selection_window(combo_work_selection_window):
    window_select_company = Toplevel()
    x = main_window_of_gui.winfo_x()
    y = main_window_of_gui.winfo_y()
    w = main_window_of_gui.winfo_width()
    h = main_window_of_gui.winfo_height()
    
    window_select_company.geometry("%dx%d+%d+%d" % (w, h, x, y))
    window_select_company.title("Choix de l'entreprise")
    window_select_company.wm_attributes("-topmost", 1)
    
    combo_company_selection_window = Combobox(window_select_company, values = LIST_OF_COMPANIES)
    combo_company_selection_window.grid(column=0, row=0, columnspan=2)
    combo_company_selection_window.current(0)

    button_open_details_entry = Button(window_select_company, text="Selectioner", width=20, height=3, command=lambda: open_details_entry(combo_work_selection_window, combo_company_selection_window))
    button_open_details_entry.grid(column=0, row=1)

    button_cancel_company_selection = Button(window_select_company, text="Annuler", width=20, height=3, command=lambda: cancel_selection_window(window_select_company))
    button_cancel_company_selection.grid(column=1, row=1)


def open_details_entry(combo_work_selection_window, combo_company_selection_window):
    window_details_entry = Toplevel()
    x = main_window_of_gui.winfo_x()
    y = main_window_of_gui.winfo_y()
    w = main_window_of_gui.winfo_width()
    h = main_window_of_gui.winfo_height()
    
    window_details_entry.geometry("%dx%d+%d+%d" % (w, h, x, y))
    window_details_entry.title("Informations additionelles")
    window_details_entry.wm_attributes("-topmost", 1)

    var_work_type = StringVar()
    var_work_type.set(combo_work_selection_window.get())
    label_work_type = Label(window_details_entry, textvariable=var_work_type)
    label_work_type.grid(column=0, row=0)

    var_company_name = StringVar()
    var_company_name.set(combo_company_selection_window.get())
    label_company_name = Label(window_details_entry, textvariable=var_company_name)
    label_company_name.grid(column=0, row=1)

    entry_forecasted_price = Entry(window_details_entry)
    entry_forecasted_price.grid(column=0, row=2)

    entry_forecasted_start_date = Entry(window_details_entry)
    entry_forecasted_start_date.insert(0, get_date_yyyy_mm_dd())
    entry_forecasted_start_date.grid(column=0, row=2)

def cancel_selection_window(window_to_close):
    window_to_close.destroy()


def get_list_of_names_from_first_sheet(excel_workbook):
    list_of_names = []
    ws = excel_workbook["first"]
    empty_line_number = find_the_next_empty_row(ws)
    for index in range(2, empty_line_number):
        saved_name = str(ws.cell(row=index, column=1).value)
        list_of_names.append(saved_name)
    return list_of_names


'''
writing a value to a cell
ws.cell(row=empty_line_number, column=2, value=last_payment)

getting value from a cell
saved_name = str(ws.cell(row=index, column=1).value)

'''

def create_excel_file_if_it_was_not_there():
    if not check_if_excel_file_is_there():
        wb = Workbook()
        ws = wb.active
        ws.title = "first"
        ws.cell(row=1, column=1, value="name")
        ws.cell(row=1, column=2, value="last date")
        ws.cell(row=1, column=3, value="last payment")
        ws.cell(row=1, column=4, value="total payment")
        wb.save(generate_the_excel_file_name_with_current_year_in_name())


main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("sandbox")
main_window_of_gui.wm_attributes("-topmost", 1)

button_start_new_work = Button(main_window_of_gui, text="Nouvelle facture", width=20, height=3, command=open_work_selection_window)
button_start_new_work.grid(row=0, column=0)

main_window_of_gui.mainloop()
