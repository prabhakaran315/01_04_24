import tkinter as tk
from tkinter import ttk
import math
import time

# Sample data for the Treeview
data_1 = [
    [1, "630A ASTRA - B3", 1252],
    [2, "630A ASTRA - B5", 1809],
    [3, "630A ASTRA - B10", 2782],
    [4, "315A ASTRA - B3", 835],
    [5, "315A ASTRA - B5", 1252],
    [6, "315A ASTRA - B10", 1530],
    [7, "420A ASTRA - B10", 1669],
    [8, "420A ASTRA - B5", 1252],
    [9, "210A ASTRA - B10", 1530],
    [10, "210A ASTRA - B5", 696],
    [11, "150A ASTRA - B5", 835],
    [12, "150A ASTRA - B10", 1669]
]

data_2 = [
    [1, "630A ASTRA - B3", 255],
    [2, "630A ASTRA - B5", 325],
    [3, "630A ASTRA - B10", 500],
    [4, "315A ASTRA - B3", 150],
    [5, "315A ASTRA - B5", 225],
    [6, "315A ASTRA - B10", 275],
    [7, "420A ASTRA - B10", 300],
    [8, "420A ASTRA - B5", 225],
    [9, "210A ASTRA - B10", 275],
    [10, "210A ASTRA - B5", 125],
    [11, "150A ASTRA - B5", 150],
    [12, "150A ASTRA - B10", 300]
]

def submit():
    try:
        value_1 = float(primary.get())
        value_2 = float(secondary.get())
        cal_opt_val = round(value_1 / value_2)
        kva = math.floor(1.732 * 415.0 * (value_1 / 1000))
        kvar = (float(kva * 0.03))
        kw = math.ceil(float(kva * 0.05))

        #------Label for  cal_opt_val-----------#
        lab_1 = tk.Label(frame, bg="white")
        lab_1.config(text="Primary / Secondary = " + str(cal_opt_val)+"\t\t\t\t", bg="white")
        lab_1.place(x=20, y=400)

        #---------------LAbel for kVA---------------#
        lab_2 = tk.Label(frame, bg="white")
        lab_2.config(text="kVA = " + str(kva) + "\t\t\t\t", bg="white")
        lab_2.place(x=20, y=420)

        #----------------Label for kVAr-------------#
        lab_3 = tk.Label(frame, bg="white")
        lab_3.config(text="Primary / Secondary = " + str(kvar) + "\t\t\t\t", bg="white")
        lab_3.place(x=20, y=440)

        #---------------Label for kW----------------#
        lab_4 = tk.Label(frame, bg="white")
        lab_4.config(text="kW = " + str(kw) + "\t\t\t\t", bg="white")
        lab_4.place(x=20, y=460)

        if all(row[2] > cal_opt_val for row in data_1):
            switch_condition(data_2, columns_cond_2, 0, kw)
        else:
            switch_condition(data_1, columns_cond_1, cal_opt_val, 0)
    except ValueError:
        lab_5 = tk.Label(frame, text = "Kindly Enter the both two values", bg="red", fg="white")
        lab_5.place(x=370, y= 550)
        time.sleep(5)


def switch_condition(data, columns, cal_opt_val=0, kw=0):
    tree.delete(*tree.get_children())
    populate_treeview(tree, columns, data, cal_opt_val, kw)

def populate_treeview(tree, columns, data, cal_opt_val=0, kw=0):
    tree["columns"] = columns
    tree.heading("#0", text="Index", anchor=tk.CENTER)

    for col in columns:
        tree.heading(col, text=col, anchor=tk.CENTER)
        tree.column(col, anchor=tk.CENTER)

    for i, row in enumerate(data, start=1):
        if columns == columns_cond_1:
            tree.insert("", "end", iid=i, text=str(i), values=(row[0], row[1], row[2], math.ceil(row[2] / cal_opt_val)))
        else:
            tree.insert("", "end", iid=i, text=str(i), values=(row[0], row[1], row[2], math.ceil(row[2] / kw)))

# Create the Tkinter window
root = tk.Tk()
root.title("True_Power_Factor")
root.geometry("800x800")

# Create a frame for the Treeview and scrollbars
frame = tk.Frame(root, bg="Grey")
frame.pack(fill=tk.BOTH, expand=True)

# Create a Treeview widget
tree = ttk.Treeview(frame, show="headings", height=12)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Define columns for Condition 1 and Condition 2
columns_cond_1 = ["S.No", "Panel Rating", "Optimal CT ratio", "No. of Panel required"]
columns_cond_2 = ["S.No", "Panel Rating", "Optimal kW", "No. of Panel required"]

# Labels, entries, and button
tk.Label(root, text="Enter the Primary value").pack()
primary = tk.Entry(root)
primary.pack()
tk.Label(root, text="Enter the Secondary value").pack()
secondary = tk.Entry(root)
secondary.pack()

tk.Button(root, text="Submit", command=submit).pack()

root.mainloop()
