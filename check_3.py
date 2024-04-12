from tkinter import *
from tkinter.ttk import Treeview
import math

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
    global cal_opt_val, value_1, value_2
    value_1 = float(primary.get())
    value_2 = float(secondary.get())
    cal_opt_val = round(value_1 / value_2)
    #------------To find the grid CT Value----------#
    primary_value = float(primary.get())
    kva = math.floor(1.732 * 415.0 * (primary_value / 1000))
    print("KVA:", kva, "kVA")

    #-----------To find the kVAr VAlue-------------------#
    kvar = float(kva*0.03)
    print("kVAr = ",kvar)

    #-----------To find the kW------------------#
    kw = float(kva*0.05)
    print("kW = ", math.ceil(kw))

    if cal_opt_val > data_1[-1][2]:
        condition_1()
    else:
        condition_2()



def condition_1():
    global value_1, value_2
    out_val_lab=Label(root, text="kW : " + str(cal_opt_val)).place(x=300, y=180)
    condi_lab=Label(root, text="This CT ratio comes under Condition - 1").place(x=300, y=210)

    # Create a Treeview widget
    tree = Treeview(root, columns=["S.No", "Panel Rating", "Optimal CT ratio"], show="headings", height=12)

    # Add columns to the Treeview
    for col in ["S.No", "Panel Rating", "Optimal CT ratio"]:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, anchor="center")

    # Insert data into the Treeview
    for item in data_1:
        tree.insert("", "end", values=item)

    # Add Treeview to the window and pack it
    tree.place(x=10, y=280)
    scrollbar = Scrollbar(root, orient=VERTICAL, command=tree.yview)
    scrollbar.place(x=775, y=200, height=500)
    scrollbar_2 = Scrollbar(root, orient=HORIZONTAL, command=tree.xview)
    scrollbar_2.place(x=10, y=750, height=760)
    tree.config(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_2.set)

def condition_2():
    global value_1, value_2
    out_val_lab = Label(root, text="kW : " + str(cal_opt_val)).place(x=300, y=180)
    condi_lab = Label(root, text="This CT ratio comes under Condition - 2").place(x=300, y=210)

    # Create a Treeview widget
    tree = Treeview(root, columns=["S.No", "Panel Rating", "Optimal kW"], show="headings", height=12)

    # Add columns to the Treeview
    for col in ["S.No", "Panel Rating", "Optimal kW"]:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, anchor="center")

    # Insert data into the Treeview
    for item in data_2:
        tree.insert("", "end", values=item)

    # Add Treeview to the window and pack it
    tree.place(x=10, y=280)
    scrollbar = Scrollbar(root, orient=VERTICAL, command=tree.yview)
    scrollbar.place(x=775, y=200, height=500)
    scrollbar_2 = Scrollbar(root, orient=HORIZONTAL, command=tree.xview)
    scrollbar_2.place(x=10, y=750, height=760)
    tree.config(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_2.set)

# Create the Tkinter window
root = Tk()
root.title("True_Power_Factor")
root.geometry("800x800")
root.resizable(False, False)

# Create a frame
frame_1 = Frame(root, height=200, width=800, bg="Grey")
frame_1.place(x=0, y=0)
frame_2 = Frame(root, height=500, width=760, bg="Pink")
frame_2.place(x=10, y=210)

#--------------text variable-------#
primary = StringVar()
secondary = StringVar()

# Create labels, entries, and button
Label(frame_1, text="Enter the Primary value", bg="Grey", fg="White").place(x=195, y=49)
Label(frame_1, text="Enter the Secondary value", bg="Grey", fg="White").place(x=195, y=79)
primary_ent = Entry(frame_1, textvariable=primary)
primary_ent.place(x=350, y=50)
secondary_ent = Entry(frame_1,textvariable =secondary)
secondary_ent.place(x=350, y=80)
Button(frame_1, text="Submit", command=submit).place(x=385, y=110)


root.mainloop()
