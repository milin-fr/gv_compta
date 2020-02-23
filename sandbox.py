from tkinter import *
from tkinter.ttk import *


root = Tk()


frame_for_the_list = Frame(root)
frame_for_the_list.grid(column=0, row=0)
tv = Treeview(frame_for_the_list)
tv['columns'] = ('starttime', 'endtime', 'status')
tv.heading("#0", text='Sources', anchor='w')
tv.column("#0", anchor="w")
tv.heading('starttime', text='Start Time')
tv.column('starttime', anchor='center', width=100)
tv.heading('endtime', text='End Time')
tv.column('endtime', anchor='center', width=100)
tv.heading('status', text='Status')
tv.column('status', anchor='center', width=100)
tv.pack()
for i in range(1000):
    tv.insert('', 'end', text="First " + str(i), values=('10:00 ' + str(i),
                         '10:10 ' + str(i), 'Ok ' + str(i)))
tv.configure(yscrollcommand=scrollbar.set)

root.mainloop()

