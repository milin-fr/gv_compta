import tkinter
from tkinter import Label, Button, Entry, Text, Checkbutton, OptionMenu, Canvas, Frame, Toplevel
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


class Bill:
    def __init__(self, current_state, work_type, company_name, forecasted_price, forecasted_start_date, forecasted_end_date, initial_comment):
        self.current_state = current_state
        self.work_type = work_type
        self.company_name = company_name
        self.forecasted_price = forecasted_price
        self.forecasted_start_date = forecasted_start_date
        self.forecasted_end_date = forecasted_end_date
        self.initial_comment = initial_comment
        self.work_started_comment = ""
        self.work_ongoing_comment = ""
        self.work_finished_comment = ""
        self.real_start_date = ""
        self.real_price = ""
        self.real_end_date = ""
        self.state_list = ["Not started", "Started", "Ongoing", "Finished", "Canceled"]
        self.excel_file_name = "GV compta " + work_type + ".xlsx"


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


def create_excel_file_if_it_was_not_there(bill_object):
    if bill_object.excel_file_name not in get_existing_excel_names():
        wb = Workbook()
        ws = wb.active
        ws.title = bill_object.company_name
        ws.cell(row=1, column=1, value="work_type")
        ws.cell(row=1, column=2, value="company_name")
        ws.cell(row=1, column=3, value="forecasted_price")
        ws.cell(row=1, column=4, value="forecasted_start_date")
        ws.cell(row=1, column=5, value="forecasted_end_date")
        ws.cell(row=1, column=6, value="initial_comment")
        ws.cell(row=1, column=7, value="current_state")
        ws.cell(row=1, column=8, value="work_started_comment")
        ws.cell(row=1, column=9, value="work_ongoing_comment")
        ws.cell(row=1, column=10, value="work_finished_comment")
        ws.cell(row=1, column=11, value="real_start_date")
        ws.cell(row=1, column=12, value="real_price")
        ws.cell(row=1, column=13, value="real_end_date")
        wb.save(bill_object.excel_file_name)


def find_the_next_empty_row(ws):
    row_index = 1
    cell_to_check = ws.cell(row = row_index, column = 1)
    while cell_to_check.value != None:
        row_index = row_index + 1
        cell_to_check = ws.cell(row = row_index, column = 1)
    return row_index


def create_missing_sheet_if_it_was_not_there(bill_object):
    wb = load_workbook(bill_object.excel_file_name)
    if bill_object.company_name not in wb.sheetnames:
        wb.create_sheet(bill_object.company_name)
        ws = wb[bill_object.company_name]
        ws.cell(row=1, column=1, value="work_type")
        ws.cell(row=1, column=2, value="company_name")
        ws.cell(row=1, column=3, value="forecasted_price")
        ws.cell(row=1, column=4, value="forecasted_start_date")
        ws.cell(row=1, column=5, value="forecasted_end_date")
        ws.cell(row=1, column=6, value="initial_comment")
        ws.cell(row=1, column=7, value="current_state")
        ws.cell(row=1, column=8, value="work_started_comment")
        ws.cell(row=1, column=9, value="work_ongoing_comment")
        ws.cell(row=1, column=10, value="work_finished_comment")
        ws.cell(row=1, column=11, value="real_start_date")
        ws.cell(row=1, column=12, value="real_price")
        ws.cell(row=1, column=13, value="real_end_date")
        wb.save(bill_object.excel_file_name)
    else:
        wb.close()


def open_details_entry():
    today_date_yyyy_mm_dd = get_date_yyyy_mm_dd()
    window_details_entry = Toplevel()
    x = main_window_of_gui.winfo_x()
    y = main_window_of_gui.winfo_y()
    w = main_window_of_gui.winfo_width()
    h = main_window_of_gui.winfo_height()
    
    window_details_entry.geometry("%dx%d+%d+%d" % (w, h, x, y))
    window_details_entry.title("Informations additionelles")
    window_details_entry.wm_attributes("-topmost", 1)

    combo_work_selection_window = Combobox(window_details_entry, values = LIST_TYPE_OF_WORK)
    combo_work_selection_window.grid(column=0, row=0, columnspan=2)
    combo_work_selection_window.current(0)

    combo_company_selection_window = Combobox(window_details_entry, values = LIST_OF_COMPANIES)
    combo_company_selection_window.grid(column=0, row=1, columnspan=2)
    combo_company_selection_window.current(0)

    entry_forecasted_price = Entry(window_details_entry)
    entry_forecasted_price.grid(column=0, row=3)

    entry_forecasted_start_date = Entry(window_details_entry)
    entry_forecasted_start_date.insert(0, today_date_yyyy_mm_dd)
    entry_forecasted_start_date.grid(column=0, row=4)

    entry_forecasted_end_date = Entry(window_details_entry)
    entry_forecasted_end_date.insert(0, today_date_yyyy_mm_dd)
    entry_forecasted_end_date.grid(column=0, row=5)

    text_first_comment = Text(window_details_entry)
    text_first_comment.grid(column=0, row=6, columnspan=2)

    data_entries = [combo_work_selection_window, combo_company_selection_window, entry_forecasted_price, entry_forecasted_start_date, entry_forecasted_end_date, text_first_comment]

    button_confirm_details_entry = Button(window_details_entry, text="Selectioner", width=20, height=3, command=lambda: confirm_details_entry(data_entries))
    button_confirm_details_entry.grid(column=0, row=7)

    button_cancel_details_entry = Button(window_details_entry, text="Annuler", width=20, height=3, command=lambda: cancel_current_window(window_details_entry))
    button_cancel_details_entry.grid(column=1, row=7)


def cancel_current_window(window_to_close):
    window_to_close.destroy()


def confirm_details_entry(data_entries):
    type_of_work = data_entries[0].get()
    company_name = data_entries[1].get()
    forecasted_price = data_entries[2].get()
    forecasted_start_date = data_entries[3].get()
    forecasted_end_date = data_entries[4].get()
    first_comment = data_entries[5].get('1.0', 'end-1c')
    
    bill_object = Bill("Not started", type_of_work, company_name, forecasted_price, forecasted_start_date, forecasted_end_date, first_comment)
    
    create_excel_file_if_it_was_not_there(bill_object)
    create_missing_sheet_if_it_was_not_there(bill_object)



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

main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("sandbox")
main_window_of_gui.wm_attributes("-topmost", 1)

button_start_new_work = Button(main_window_of_gui, text="Nouvelle facture", width=20, height=3, command=open_details_entry)
button_start_new_work.grid(row=0, column=0)

main_window_of_gui.mainloop()
