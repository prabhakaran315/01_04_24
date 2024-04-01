from tkinter import *
root = Tk()
root.title("Addition")
root.geometry("800x800")
root.resizable(False, False)

primary = StringVar()
secondary = StringVar()
prim_label = Label(root, text = "Enter the Primary value").place(x=350, y=20)
Secon_label = Label(root, text = "Enter the Secondary value").place(x=350, y=80)
prim_entry = Entry(root, textvariable=primary).place(x=350, y=50)
secon_entry = Entry(root, textvariable=secondary).place(x=350, y=110)
root.mainloop()