import tkinter as tk
from tkinter import ttk
import openpyxl

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

def load_data():
    path = 'C:\\Users\\augus\\Documents\\Python\QT\\people.xlsx'
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    # print(list_values)
    for col_name in list_values[0]:
        treeView.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeView.insert('', tk.END, values=value_tuple)

def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_satus = "Employed" if a.get() else "Unemployed"

    print(name, age, subscription_status, employment_satus)

    path = 'C:\\Users\\augus\\Documents\\Python\QT\\people.xlsx'
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employment_satus]
    sheet.append(row_values)
    workbook.save(path)

    treeView.insert('', tk.END, values=row_values)
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list[0])
    checkButton.state(["!selected"])


root = tk.Tk()
root.title("Subcripciones")
root.option_add("*tearOff", False)

#-------------
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style = ttk.Style(root)
style.theme_use("forest-dark")

#
combo_list = ["Suscrito", "No Suscrito", "Otro"]


#---------
# Frame
#---------
# Toda va a estar dento del frame para que sea mas practico
frame = tk.Frame(root)
frame.pack()

#  El primer widget  que es el contenedor de la entrada
widget_frame = ttk.LabelFrame(frame, text="Insert Row")
widget_frame.grid(row=0, column=0, padx= 10, pady= 10)
# El form para agregar el nombre 
name_entry = ttk.Entry(widget_frame)
name_entry.insert(0, "Nombre")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end')) # Borro el nombre al hacer foco
name_entry.grid(row=0, column=0, padx=5, pady=(0,5), sticky="ew")
# 
age_spinbox = ttk.Spinbox(widget_frame, from_=18, to= 100)
age_spinbox.insert(0, "Edad")
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
# 
status_combobox = ttk.Combobox(widget_frame, values= combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
#
a = tk.BooleanVar()
checkButton = ttk.Checkbutton(widget_frame, text="Employed", variable=a)
checkButton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")
# Boton instertar 
button = ttk.Button(widget_frame, text="Insertar", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")
#----------------------------------------------------------------
separator = ttk.Separator(widget_frame)
separator.grid(row=5, column=0, padx=5, pady=5, sticky="ew")
#
mode_switch = ttk.Checkbutton(
    widget_frame, text="Modo", state="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

#----------------------------------------------------------------
# Tree Frame
#----------------------------------------------------------------
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

#
cols = ("Nombre", "Edad", "Subscrito", "Empleo")
treeView = ttk.Treeview(treeFrame, columns=cols, height=13, show="headings")
treeView.column("Nombre", width=100)
treeView.column("Edad", width=50)
treeView.column("Subscrito", width=100)
treeView.column("Empleo", width=100)
treeView.pack()
treeScroll.config(command=treeView.yview)
load_data()

root.mainloop()