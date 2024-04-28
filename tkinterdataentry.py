#data entry form 
import tkinter as tk
from tkinter import ttk
import openpyxl

def load_data():
    path="PEOPLE.xlsx"
    workbook=openpyxl.load_workbook(path)
    sheet=workbook.active

    list_values=list(sheet.values)
    print(list_values)
    # for col_name in list_values[0]:
    #     treeview.heading(column=col_name, text=col_name)
            
        
    for value_tuple in list_values[1:]:
        treeview.insert('',tk.END,values=value_tuple)

def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Y" if a.get() else "N"

    print(name, age, subscription_status, employment_status)

    # Insert row into Excel sheet
    path = "PEOPLE.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    treeview.heading(row_values[0], text='Name')
    
    # Clear the values
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list['subscribtion'])
    checkbutton.state(["!selected"])



def toggle_mode():
    if mode_switch.instate(['selected']):
        style.theme_use('forest-light')
    else:
        style.theme_use('forest-dark')

combo_list=['Y','N','Other']

root= tk.Tk()

style=ttk.Style(root)
root.tk.call('source','forest-light.tcl')
root.tk.call('source','forest-dark.tcl')
style.theme_use('forest-dark')

frame=ttk.Frame(root)
frame.pack()

widget_frame=ttk.Labelframe(frame,text='Insert Row')
widget_frame.grid(row=0,column=0,padx=20,pady=10)

name_entry=ttk.Entry(widget_frame)
name_entry.insert(0,'Name')
name_entry.bind('<FocusIn>',lambda e: name_entry.delete ('0','end'))
name_entry.grid(row=0,column=0,padx=5,pady=(0,5),sticky='ew')

age_spinbox=ttk.Spinbox(widget_frame,from_=18, to=100)
age_spinbox.insert(0,'Age')
age_spinbox.grid(row=1,column=0,padx=5,pady=5,sticky='ew')

status_combobox=ttk.Combobox(widget_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2,column=0,padx=5,pady=5,sticky='ew')

a=tk.BooleanVar()
checkbutton=ttk.Checkbutton(widget_frame,text='Employed',variable=a)
checkbutton.grid(row=3, column=0,padx=5,pady=5, sticky='nsew')



button=ttk.Button(widget_frame, text='Insert',command=insert_row)
button.grid(row=4,column=0,padx=5,pady=5, sticky='nsew')

seperator=ttk.Separator(widget_frame)
seperator.grid(row=5,column=0,padx=(20,10),pady=10,sticky='ew')

mode_switch=ttk.Checkbutton(widget_frame,text='Mode',style='Switch',command=toggle_mode)
mode_switch.grid(row=6,column=0,padx=5,pady=10,sticky='nsew')

treeframe=ttk.Frame(frame)
treeframe.grid(row=0, column=1, pady=10)
treescroll=ttk.Scrollbar(treeframe)
treescroll.pack(side='right',fill='y')

cols=('Name','Age','Subscribtion','Employment')
treeview=ttk.Treeview(treeframe,show='headings',yscrollcommand=treescroll.set,columns=cols,height=13)
treeview.column('Name',width=100)
treeview.heading('Name',text='Name')
treeview.column('Age',width=50)
treeview.heading('Age',text='Age')
treeview.column('Subscribtion',width=100)
treeview.heading('Subscribtion',text='Subsribtion')
treeview.column('Employment',width=100)
treeview.heading('Employment',text='Employment')
treeview.pack()
treescroll.config(command=treeview.yview)
load_data()



root.mainloop()