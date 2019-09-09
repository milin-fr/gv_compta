import tkinter
from tkinter import Label, Button, Entry, Text, Checkbutton, OptionMenu, Canvas, Frame, Toplevel, Scrollbar, Listbox, Frame
from tkinter.ttk import Combobox, Treeview
from tkinter import StringVar, IntVar, RIGHT, LEFT, BOTH, Y, END
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

LIST_TYPE_OF_WORK = []
LIST_OF_COMPANIES = []
LIST_OF_BILLS = []


class Bill:
    row_placement = ""
    current_state = ""
    work_type = ""
    company_name = ""
    forecasted_price = ""
    forecasted_start_date = ""
    forecasted_end_date = ""
    initial_comment = ""
    work_started_comment = ""
    work_ongoing_comment = ""
    work_finished_comment = ""
    real_start_date = ""
    real_price = ""
    real_end_date = ""
    state_list = ["Not started", "Started", "Ongoing", "Finished", "Canceled"]
    excel_file_name = ""
    def set_excel_name(self):
        self.excel_file_name = "GV compta " + self.work_type + ".xlsx"


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
        ws.cell(row=1, column=1, value="current_state")
        ws.cell(row=1, column=2, value="work_type")
        ws.cell(row=1, column=3, value="company_name")
        ws.cell(row=1, column=4, value="forecasted_price")
        ws.cell(row=1, column=5, value="forecasted_start_date")
        ws.cell(row=1, column=6, value="forecasted_end_date")
        ws.cell(row=1, column=7, value="initial_comment")
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
    update_work_type_list()
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

    combo_company_selection_window = Combobox(window_details_entry, values = LIST_OF_COMPANIES)
    combo_company_selection_window.grid(column=0, row=1, columnspan=2)
    
    button_confirm_work_type = Button(window_details_entry, text="Ok", command=lambda: update_company_name_list(combo_work_selection_window, combo_company_selection_window))
    button_confirm_work_type.grid(column=2, row=0)

    entry_forecasted_price = Entry(window_details_entry)
    entry_forecasted_price.grid(column=0, row=3)

    entry_forecasted_start_date = Entry(window_details_entry)
    entry_forecasted_start_date.insert(0, today_date_yyyy_mm_dd)
    entry_forecasted_start_date.grid(column=0, row=4)

    entry_forecasted_end_date = Entry(window_details_entry)
    entry_forecasted_end_date.insert(0, today_date_yyyy_mm_dd)
    entry_forecasted_end_date.grid(column=0, row=5)

    text_first_comment = Text(window_details_entry, width=40, height=20)
    text_first_comment.grid(column=0, row=6, columnspan=2)

    data_entries = [combo_work_selection_window, combo_company_selection_window, entry_forecasted_price, entry_forecasted_start_date, entry_forecasted_end_date, text_first_comment]

    button_confirm_details_entry = Button(window_details_entry, text="Selectioner", width=20, height=3, command=lambda: confirm_details_entry(data_entries, window_details_entry))
    button_confirm_details_entry.grid(column=0, row=7)

    button_cancel_details_entry = Button(window_details_entry, text="Annuler", width=20, height=3, command=lambda: cancel_current_window(window_details_entry))
    button_cancel_details_entry.grid(column=1, row=7)


def cancel_current_window(window_to_close):
    window_to_close.destroy()


def confirm_details_entry(data_entries, window_to_close):
    work_type = data_entries[0].get()
    company_name = data_entries[1].get()
    forecasted_price = data_entries[2].get()
    forecasted_start_date = data_entries[3].get()
    forecasted_end_date = data_entries[4].get()
    first_comment = data_entries[5].get('1.0', 'end-1c')
    
    bill_object = Bill()
    bill_object.current_state = "Not started"
    bill_object.work_type = work_type
    bill_object.company_name = company_name
    bill_object.forecasted_price = forecasted_price
    bill_object.forecasted_start_date = forecasted_start_date
    bill_object.forecasted_end_date = forecasted_end_date
    bill_object.initial_comment = first_comment
    bill_object.set_excel_name()

    create_excel_file_if_it_was_not_there(bill_object)
    create_missing_sheet_if_it_was_not_there(bill_object)
    save_first_entry(bill_object)
    cancel_current_window(window_to_close)


def save_first_entry(bill_object):
    wb = load_workbook(bill_object.excel_file_name)
    ws = wb[bill_object.company_name]
    new_row = find_the_next_empty_row(ws)
    ws.cell(row=new_row, column=1, value=bill_object.current_state)
    ws.cell(row=new_row, column=2, value=bill_object.work_type)
    ws.cell(row=new_row, column=3, value=bill_object.company_name)
    ws.cell(row=new_row, column=4, value=bill_object.forecasted_price)
    ws.cell(row=new_row, column=5, value=bill_object.forecasted_start_date)
    ws.cell(row=new_row, column=6, value=bill_object.forecasted_end_date)
    ws.cell(row=new_row, column=7, value=bill_object.initial_comment)

    wb.save(bill_object.excel_file_name)


def get_list_of_names_from_first_sheet(excel_workbook):
    list_of_names = []
    ws = excel_workbook["first"]
    empty_line_number = find_the_next_empty_row(ws)
    for index in range(2, empty_line_number):
        saved_name = str(ws.cell(row=index, column=1).value)
        list_of_names.append(saved_name)
    return list_of_names


def update_work_type_list():
    global LIST_TYPE_OF_WORK
    LIST_TYPE_OF_WORK = []
    existing_excel_names = get_existing_excel_names()
    for file_name in existing_excel_names:
        work_type = file_name[10:-5]
        LIST_TYPE_OF_WORK.append(work_type)


def update_company_name_list(combo_work_selection_window, combo_company_selection_window):
    global LIST_OF_COMPANIES
    LIST_OF_COMPANIES = []
    excel_file_name = "GV compta " + combo_work_selection_window.get() + ".xlsx"
    if excel_file_name in get_existing_excel_names():
        wb = load_workbook(excel_file_name)
        LIST_OF_COMPANIES = wb.sheetnames
        wb.close()
        combo_company_selection_window['values'] = LIST_OF_COMPANIES
        combo_company_selection_window.current(0)

def onselect(evt):
    global LIST_OF_BILLS
    # Note here that Tkinter passes an event object to onselect()
    w = evt.widget
    row_object = w.selection()[0]
    text = w.item(row_object, 'text')
    print('You selected item ', row_object,text)



def open_ongoing_view():
    
    get_list_of_bills()
    window_details_entry = Toplevel()
    x = main_window_of_gui.winfo_x()
    y = main_window_of_gui.winfo_y()
    w = main_window_of_gui.winfo_width()
    h = main_window_of_gui.winfo_height()
    
    window_details_entry.geometry("%dx%d+%d+%d" % (w, h, x, y))
    window_details_entry.title("Factures en cours")
    window_details_entry.wm_attributes("-topmost", 1)

    frame_for_the_list = Frame(window_details_entry)
    frame_for_the_list.grid(column=0, row=0)
    
    treeview_details_of_ongoing_bills = Treeview(frame_for_the_list)

    scrollbar = Scrollbar(frame_for_the_list, command=treeview_details_of_ongoing_bills.yview)
    scrollbar.pack(side=RIGHT, fill=Y)

    treeview_details_of_ongoing_bills['columns'] = ('1', '2', '3')
    treeview_details_of_ongoing_bills['show'] = 'headings'
    treeview_details_of_ongoing_bills.heading('1', text='my 1')
    treeview_details_of_ongoing_bills.heading('2', text='my 2')
    treeview_details_of_ongoing_bills.heading('3', text='my 3')
    treeview_details_of_ongoing_bills.column('1', anchor='center', width=100)
    treeview_details_of_ongoing_bills.column('2', anchor='center', width=100)
    treeview_details_of_ongoing_bills.column('3', anchor='center', width=100)
    treeview_details_of_ongoing_bills.pack()
    for bill in LIST_OF_BILLS:
        for i in range(100):
            treeview_details_of_ongoing_bills.insert('', 'end', text=str(bill.row_placement), values=('10:00 ' + str(i),
                                 '10:10 ' + str(i), 'Ok ' + str(i)))
    treeview_details_of_ongoing_bills.bind('<Double-1>', onselect)
    treeview_details_of_ongoing_bills.configure(yscrollcommand=scrollbar.set)

    button = Button(window_details_entry, text="OK")
    button.grid(column=0, row=1)



def get_details_out_of_excel(excel_file):
    global LIST_OF_BILLS
    wb = load_workbook(excel_file)
    for company_sheet in wb.sheetnames:
        ws = wb[company_sheet]
        maximum_number_of_rows = find_the_next_empty_row(ws)
        for current_row in range(2, maximum_number_of_rows):
            bill = Bill()
            bill.row_placement = current_row
            bill.current_state = ws.cell(row=current_row, column=1).value
            bill.work_type = ws.cell(row=current_row, column=2).value
            bill.company_name = ws.cell(row=current_row, column=3).value
            bill.forecasted_price = ws.cell(row=current_row, column=4).value
            bill.forecasted_start_date = ws.cell(row=current_row, column=5).value
            bill.forecasted_end_date = ws.cell(row=current_row, column=6).value
            bill.initial_comment = ws.cell(row=current_row, column=7).value
            bill.work_started_comment = ws.cell(row=current_row, column=8).value
            bill.work_ongoing_comment = ws.cell(row=current_row, column=9).value
            bill.work_finished_comment = ws.cell(row=current_row, column=10).value
            bill.real_start_date = ws.cell(row=current_row, column=11).value
            bill.real_price = ws.cell(row=current_row, column=12).value
            bill.real_end_date = ws.cell(row=current_row, column=13).value
            LIST_OF_BILLS.append(bill)
    wb.close()



def get_list_of_bills():
    global LIST_OF_BILLS
    LIST_OF_BILLS = []
    existing_excel_names = get_existing_excel_names()
    for excel_file in existing_excel_names:
        get_details_out_of_excel(excel_file)



'''
writing a value to a cell
ws.cell(row=empty_line_number, column=2, value=last_payment)

getting value from a cell
saved_name = str(ws.cell(row=index, column=1).value)

'''

main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("sandbox")
main_window_of_gui.wm_attributes("-topmost", 1)
main_window_of_gui.geometry("500x500") #You want the size of the app to be 500x500
main_window_of_gui.resizable(0, 0)

button_start_new_work = Button(main_window_of_gui, text="Nouvelle facture", width=20, height=3, command=open_details_entry)
button_start_new_work.grid(row=0, column=0)

button_view_ongoing_work = Button(main_window_of_gui, text="Facture en cours", width=20, height=3, command=open_ongoing_view)
button_view_ongoing_work.grid(row=0, column=1)


main_window_of_gui.mainloop()
