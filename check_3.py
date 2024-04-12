from tkinter import *
from tkinter.ttk import Treeview

from tabulate import tabulate

#------Creating  the Window------#
root = Tk()
root.title("True_Power_Factor")
root.geometry("800x800")
root.resizable(False, False)

#-----------Submit Button Function--------#
def submit():
    value_1 = float(primary.get())
    value_2 = float(secondary.get())
    cal_opt_val = (round(value_1 / value_2))
    #out_val_lab = Label(frame_1, text ="kW : "+cal_opt_val).place(x=350, y=140)
   #condi_lab = Label(root, text = "This CT ratio comes under Condition - 2").place(x=300, y=250)
    return cal_opt_val

def condition_2():
    value_1 = float(primary.get())
    value_2 = float(secondary.get())
    cal_opt_val = (round(value_1 / value_2))
    out_val_lab = Label(frame_1, text="kW : " + str(cal_opt_val)).place(x=350, y=140)
    condi_lab = Label(root, text="This CT ratio comes under Condition - 2").place(x=300, y=250)
    # create data
    data = [[1, "630A ASTRA - B3", 255],
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
            [12, "150A ASTRA - B10", 300]]

    # define header names
    col_names = ["S.No", "Panel Rating", "Optimal kW"]
    # Create a Treeview widget
    tree = Treeview(root, columns=col_names, show="headings", height=12)

    # Add columns to the Treeview
    for col in col_names:
        tree.heading(col, text=col, anchor="center")  # Align text in headers to center
        tree.column(col, anchor="center")

    # Insert data into the Treeview
    for item in data:
        tree.insert("", "end", values=item)

    # Add Treeview to the window and pack it
    tree.place(x=100, y= 280)


#--------Create a frame------------#
frame_1 = Frame(root, height=200, width=800, bg="Grey").place(x=0, y=0)

#-------Text Variable-----#
primary=StringVar()
secondary=StringVar()

#----------Creating the Entry, Labels and button----------------#
#-----(Label)------#
primary_Label = Label(frame_1, text = "Enter the Primary value", bg="Grey", fg="White").place(x=195, y=49)
secondary_Label = Label(frame_1, text = "Enter the Secondary value", bg="Grey", fg="White").place(x=195, y=79)
#-----(Entry)------#
primary_entry = Entry(frame_1, textvariable=primary).place(x=350, y=50)
secondary_entry = Entry(frame_1, textvariable=secondary).place(x=350, y=80)
#-----(Submit Button)------#
submit_btn = Button(frame_1, text = "Submit", command=condition_2).place(x=385, y=110)


root.mainloop()