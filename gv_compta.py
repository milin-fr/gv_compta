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
LIST_WORK_STATUS = ["Not started", "Started", "Finished", "Canceled"]

class Bill:
    row_placement = ""
    work_type = ""  #1
    company_name = ""  #2
    comment = ""  #3
    start_date = ""  #4
    end_date = ""  #5
    price = ""  #6
    work_status = ""  #7

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


def get_date_dd_mm_yyyy():
    return str(datetime.date.today().strftime("%d.%m.%Y"))


def create_bill_excel_file_if_it_was_not_there(bill_object):
    if bill_object.excel_file_name not in get_existing_excel_names():
        wb = Workbook()
        ws = wb.active
        ws.title = bill_object.company_name
        ws.cell(row=1, column=1, value="work_type")
        ws.cell(row=1, column=2, value="company_name")
        ws.cell(row=1, column=3, value="comment")
        ws.cell(row=1, column=4, value="start_date")
        ws.cell(row=1, column=5, value="end_date")
        ws.cell(row=1, column=6, value="price")
        ws.cell(row=1, column=7, value="work_status")
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
        ws.cell(row=1, column=3, value="comment")
        ws.cell(row=1, column=4, value="start_date")
        ws.cell(row=1, column=5, value="end_date")
        ws.cell(row=1, column=6, value="price")
        ws.cell(row=1, column=7, value="work_status")
        wb.save(bill_object.excel_file_name)
    else:
        wb.close()

def toplevel_was_closed(evt):
    unblock_root_buttons()

def open_details_entry():
    bloc_root_buttons()
    today_date_yyyy_mm_dd = get_date_dd_mm_yyyy()
    update_work_type_list()
    window_details_entry = Toplevel()
    x = main_window_of_gui.winfo_x()
    y = main_window_of_gui.winfo_y()
    w = main_window_of_gui.winfo_width()
    h = main_window_of_gui.winfo_height()
    
    window_details_entry.geometry("%dx%d+%d+%d" % (w, h, x, y))
    window_details_entry.title("Informations additionelles")
    window_details_entry.wm_attributes("-topmost", 1)
    
    label_work_selection = Label(window_details_entry, text = "Type de traveaux :", width=15)
    label_work_selection.grid(column=0, row=0, pady=5)

    combo_work_selection = Combobox(window_details_entry, values = LIST_TYPE_OF_WORK)
    combo_work_selection.grid(column=1, row=0, columnspan=2, pady=5)

    label_company_selection = Label(window_details_entry, text = "Nom de l'entreprise :", width=15)
    label_company_selection.grid(column=0, row=1, pady=5)

    combo_company_selection = Combobox(window_details_entry, values = LIST_OF_COMPANIES)
    combo_company_selection.grid(column=1, row=1, columnspan=2, pady=5)
    

    label_price = Label(window_details_entry, text = "Prix :", width=15)
    label_price.grid(column=0, row=2, pady=5)

    entry_price = Entry(window_details_entry, width=23)
    entry_price.grid(column=1, row=2, columnspan=2, pady=5)

    label_start_date = Label(window_details_entry, text = "Date de debut :", width=15)
    label_start_date.grid(column=0, row=3, pady=5)

    entry_start_date = Entry(window_details_entry, width=23)
    entry_start_date.insert(0, today_date_yyyy_mm_dd)
    entry_start_date.grid(column=1, row=3, columnspan=2, pady=5)

    label_end_date = Label(window_details_entry, text = "Date de fin :", width=15)
    label_end_date.grid(column=0, row=4, pady=5)

    entry_end_date = Entry(window_details_entry, width=23)
    entry_end_date.insert(0, today_date_yyyy_mm_dd)
    entry_end_date.grid(column=1, row=4, columnspan=2, pady=5)

    label_status = Label(window_details_entry, text = "Etat de traveaux :", width=15)
    label_status.grid(column=0, row=5, pady=5)

    var_work_status = StringVar()
    var_work_status.set(LIST_WORK_STATUS[0])
    dropdown_work_status = OptionMenu(window_details_entry, var_work_status, *LIST_WORK_STATUS)
    dropdown_work_status.grid(column=1, row=5, columnspan=2, pady=5)
    dropdown_work_status.config(width=18)

    label_comment = Label(window_details_entry, text = "Commentaires :", width=15)
    label_comment.grid(column=0, row=6, columnspan=3, pady=5)

    text_comment = Text(window_details_entry, width=60, height=10)
    text_comment.grid(column=0, row=7, columnspan=3, pady=5)

    data_entries = [combo_work_selection, combo_company_selection, text_comment, entry_start_date, entry_end_date, entry_price, var_work_status]

    button_confirm_details_entry = Button(window_details_entry, text="Selectioner", width=10, height=3, command=lambda: confirm_details_entry(data_entries, window_details_entry))
    button_confirm_details_entry.grid(column=0, row=8, pady=5)

    button_cancel_details_entry = Button(window_details_entry, text="Annuler", width=10, height=3, command=lambda: cancel_current_window(window_details_entry))
    button_cancel_details_entry.grid(column=2, row=8, pady=5)

    button_cancel_details_entry.bind("<Destroy>", toplevel_was_closed)  # if bind on toplevel, the destruction of all widgets in toplevel trigers the function


def cancel_current_window(window_to_close):
    window_to_close.destroy()


def confirm_details_entry(data_entries, window_to_close):
    bill_object = Bill()
    bill_object.work_type = data_entries[0].get()
    bill_object.company_name = data_entries[1].get()
    bill_object.comment = data_entries[2].get('1.0', 'end-1c')
    bill_object.start_date = data_entries[3].get()
    bill_object.end_date = data_entries[4].get()
    bill_object.price = data_entries[5].get()
    bill_object.work_status = data_entries[6].get()
    bill_object.set_excel_name()

    create_bill_excel_file_if_it_was_not_there(bill_object)
    create_missing_sheet_if_it_was_not_there(bill_object)
    save_first_entry(bill_object)
    cancel_current_window(window_to_close)


def save_first_entry(bill_object):
    wb = load_workbook(bill_object.excel_file_name)
    ws = wb[bill_object.company_name]
    new_row = find_the_next_empty_row(ws)
    ws.cell(row=new_row, column=1, value=bill_object.work_type)
    ws.cell(row=new_row, column=2, value=bill_object.company_name)
    ws.cell(row=new_row, column=3, value=bill_object.comment)
    ws.cell(row=new_row, column=4, value=bill_object.start_date)
    ws.cell(row=new_row, column=5, value=bill_object.end_date)
    ws.cell(row=new_row, column=6, value=bill_object.price)
    ws.cell(row=new_row, column=7, value=bill_object.work_status)

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


def update_company_name_list(combo_work_selection, combo_company_selection):
    global LIST_OF_COMPANIES
    LIST_OF_COMPANIES = []
    excel_file_name = "GV compta " + combo_work_selection.get() + ".xlsx"
    if excel_file_name in get_existing_excel_names():
        wb = load_workbook(excel_file_name)
        LIST_OF_COMPANIES = wb.sheetnames
        wb.close()
        combo_company_selection['values'] = LIST_OF_COMPANIES
        combo_company_selection.current(0)

def onselect(evt):
    global LIST_OF_BILLS
    # Note here that Tkinter passes an event object to onselect()
    clicked_widger = evt.widget
    row_id = clicked_widger.selection()[0] #particular line that is selected
    print(clicked_widger.item(row_id))
    text = clicked_widger.item(row_id, 'text')
    print('You selected item ', row_id, text)


def treeview_sort_column(tv, col, reverse):
    list_of_lines = tv.get_children('')
    list_of_something = []
    for single_line in list_of_lines:
        list_of_something.append((tv.set(single_line, col), single_line))
    list_of_something.sort(reverse=reverse)

    for index, (val, single_line) in enumerate(list_of_something):
        tv.move(single_line, '', index)

    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))


def open_ongoing_view():
    bloc_root_buttons()
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
    
    tv_columns = ('work_type', 'company_name', 'comment', "start_date", "end_date", "price", "work_status")
    treeview_details_of_ongoing_bills = Treeview(frame_for_the_list, columns=tv_columns, show='headings')
    
    for column in tv_columns:
        treeview_details_of_ongoing_bills.heading(column, text=column, command=lambda col=column: treeview_sort_column(treeview_details_of_ongoing_bills, col, False))
        treeview_details_of_ongoing_bills.column(column, anchor='center', width=100)
    treeview_details_of_ongoing_bills.column('comment', anchor='center', width=300)


    scrollbar = Scrollbar(frame_for_the_list, command=treeview_details_of_ongoing_bills.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
      

    treeview_details_of_ongoing_bills.pack()
    for bill in LIST_OF_BILLS:
        for i in range(10):
            treeview_details_of_ongoing_bills.insert('', 'end', text=str(bill.row_placement), 
                                values=(
                                    bill.work_type + str(i),
                                    bill.company_name + str(i),
                                    bill.comment + str(i),
                                    bill.start_date + str(i),
                                    bill.end_date + str(i),
                                    bill.price + str(i),
                                    bill.work_status + str(i)
                                ))
    treeview_details_of_ongoing_bills.bind('<Double-1>', onselect)
    treeview_details_of_ongoing_bills.configure(yscrollcommand=scrollbar.set)

    button = Button(window_details_entry, text="OK")
    button.grid(column=0, row=1)

    treeview_details_of_ongoing_bills.bind("<Destroy>", toplevel_was_closed)  # if bind on toplevel, the destruction of all widgets in toplevel trigers the function



def get_details_out_of_excel(excel_file):
    global LIST_OF_BILLS
    wb = load_workbook(excel_file)
    for company_sheet in wb.sheetnames:
        ws = wb[company_sheet]
        maximum_number_of_rows = find_the_next_empty_row(ws)
        for current_row in range(2, maximum_number_of_rows):
            bill = Bill()
            bill.row_placement = current_row
            bill.work_type = ws.cell(row=current_row, column=1).value
            bill.company_name = ws.cell(row=current_row, column=2).value
            bill.comment = ws.cell(row=current_row, column=3).value
            bill.start_date = ws.cell(row=current_row, column=4).value
            bill.end_date = ws.cell(row=current_row, column=5).value
            bill.price = ws.cell(row=current_row, column=6).value
            bill.work_status = ws.cell(row=current_row, column=7).value
            bill.set_excel_name()
            LIST_OF_BILLS.append(bill)
    wb.close()



def get_list_of_bills():
    global LIST_OF_BILLS
    LIST_OF_BILLS = []
    existing_excel_names = get_existing_excel_names()
    for excel_file in existing_excel_names:
        get_details_out_of_excel(excel_file)


def bloc_root_buttons():
    button_start_new_work.config(state="disabled")
    button_view_ongoing_work.config(state="disabled")


def unblock_root_buttons():
    button_start_new_work.config(state="normal")
    button_view_ongoing_work.config(state="normal")


'''
writing a value to a cell
ws.cell(row=empty_line_number, column=2, value=last_payment)

getting value from a cell
saved_name = str(ws.cell(row=index, column=1).value)

'''

main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("sandbox")
main_window_of_gui.geometry("500x500")
main_window_of_gui.resizable(0, 0)

button_start_new_work = Button(main_window_of_gui, text="Nouvelle facture", width=20, height=3, command=open_details_entry)
button_start_new_work.grid(row=0, column=0)

button_view_ongoing_work = Button(main_window_of_gui, text="Facture en cours", width=20, height=3, command=open_ongoing_view)
button_view_ongoing_work.grid(row=0, column=1)


main_window_of_gui.mainloop()
