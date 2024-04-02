'''import pandas as pd
from tkinter import *
from openpyxl import load_workbook

def update_result(*args):
    try:
        value1 = float(primary.get())
        value2 = float(secondary.get())
        result_var.set(value1 / value2)

        # Load the Excel file into a DataFrame
        excel_file = 'data.xlsx'
        df = pd.read_excel(excel_file)

        # Define the input value
        input_value = float(value1 / value2)  # Use the calculated result as the input value

        # Compare 'OP' column with the input value
        df['Comparison'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['KVA'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['kW'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['KVAr'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        # Save the DataFrame back to Excel with the results in a new column
        output_file = 'output.xlsx'  # Change this to the desired output file path
        df.to_excel(output_file, index=False)

        # Load the updated Excel file
        wb2 = load_workbook(output_file)
        ws = wb2.active

        # Get data from specific columns
        data_a = [str(cell.value) for cell in ws['A']]
        data_b = [str(cell.value) for cell in ws['B']]
        data_c = [str(cell.value) for cell in ws['C']]
        data_d = [str(cell.value) for cell in ws['D']]
        data_e = [str(cell.value) for cell in ws['E']]
        data_f = [str(cell.value) for cell in ws['F']]
        data_g = [str(cell.value) for cell in ws['G']]

        # Set data to label variables
        label__a.set('\n'.join(data_a))
        label__b.set('\n'.join(data_b))
        label__c.set('\n'.join(data_c))
        label__d.set('\n'.join(data_d))
        label__e.set('\n'.join(data_e))
        label__f.set('\n'.join(data_f))
        label__g.set('\n'.join(data_g))

        print("Data loaded successfully.")
    except ValueError:
        result_var.set("Enter both two values!!!")

root = Tk()
root.title("Project")
root.geometry("800x800")
root.resizable(False, False)

primary = StringVar()
secondary = StringVar()
result_var = StringVar()

label__a = StringVar()
label__b = StringVar()
label__c = StringVar()
label__d = StringVar()
label__e = StringVar()
label__f = StringVar()
label__g = StringVar()
label__h = StringVar()

prim_label = Label(root, text="Enter the Primary value")
prim_label.place(x=350, y=20)
Secon_label = Label(root, text="Enter the Secondary value")
Secon_label.place(x=350, y=80)
prim_entry = Entry(root, textvariable=primary)
prim_entry.place(x=350, y=50)
secon_entry = Entry(root, textvariable=secondary)
secon_entry.place(x=350, y=110)
result_label = Label(root, textvariable=result_var)
result_label.place(x=360, y=130)

label_a = Label(root, textvariable=label__a)
label_a.place(x=10, y=160)
label_b = Label(root, textvariable=label__b)
label_b.place(x=150, y=160)
label_c = Label(root, textvariable=label__c)
label_c.place(x=275, y=160)
label_d = Label(root, textvariable=label__d)
label_d.place(x=350, y=160)
label_e = Label(root, textvariable=label__e)
label_e.place(x=425, y=160)
label_f = Label(root, textvariable=label__f)
label_f.place(x=500, y=160)
label_g = Label(root, textvariable=label__g)
label_g.place(x=675, y=160)
label_h = Label(root, textvariable=label__h)
label_h.place(x=775, y=160)

primary.trace_add('write', update_result)
secondary.trace_add('write', update_result)
root.mainloop()'''

'''import pandas as pd
from tkinter import *
from tkinter import ttk  # Import ttk for Treeview widget
from openpyxl import load_workbook

def update_result(*args):
    try:
        value1 = float(primary.get())
        value2 = float(secondary.get())
        result_var.set(value1 / value2)

        # Load the Excel file into a DataFrame
        excel_file = 'data.xlsx'
        df = pd.read_excel(excel_file)

        # Define the input value
        input_value = float(value1 / value2)  # Use the calculated result as the input value

        # Compare 'OP' column with the input value
        df['Comparison'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['KVA'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['kW'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['KVAr'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        # Save the DataFrame back to Excel with the results in a new column
        output_file = 'output.xlsx'  # Change this to the desired output file path
        df.to_excel(output_file, index=False)

        # Load the updated Excel file
        wb2 = load_workbook(output_file)
        ws = wb2.active

        # Get data from specific columns
        data = []
        for row in ws.iter_rows(min_row=1, max_col=8, values_only=True):
            data.append(row)

        # Clear existing treeview data
        for item in treeview.get_children():
            treeview.delete(item)

        # Insert data into treeview
        for item in data[1:]:
            treeview.insert("", "end", values=item)

        print("Data loaded successfully.")
    except ValueError:
        result_var.set("Enter both two values!!!")

root = Tk()
root.title("Project")
root.geometry("800x800")
root.resizable(False, False)

primary = StringVar()
secondary = StringVar()
result_var = StringVar()

prim_label = Label(root, text="Enter the Primary value")
prim_label.place(x=350, y=20)
Secon_label = Label(root, text="Enter the Secondary value")
Secon_label.place(x=350, y=80)
prim_entry = Entry(root, textvariable=primary)
prim_entry.place(x=350, y=50)
secon_entry = Entry(root, textvariable=secondary)
secon_entry.place(x=350, y=110)
result_label = Label(root, textvariable=result_var)
result_label.place(x=360, y=130)

# Create a Treeview widget for displaying data in a table
treeview = ttk.Treeview(root, columns=('S.no', 'Power', 'OP', 'CT/R', 'Above / Below', 'KVA', 'kW', 'KVAr'),
                        show='headings')
treeview.place(x=10, y=200, width=780, height=500)

# Define headings for the columns
treeview.heading('S.no', text='S.no')
treeview.heading('Power', text='Power')
treeview.heading('OP', text='OP')
treeview.heading('CT/R', text='CT/R')
treeview.heading('Above / Below', text='Above / Below')
treeview.heading('KVA', text='KVA')
treeview.heading('kW', text='kW')
treeview.heading('KVAr', text='KVAr')

# Set column widths
treeview.column('S.no', width=50)
treeview.column('Power', width=70)
treeview.column('OP', width=70)
treeview.column('CT/R', width=70)
treeview.column('Above / Below', width=100)
treeview.column('KVA', width=70)
treeview.column('kW', width=70)
treeview.column('KVAr', width=70)

# Bind the update_result function to changes in primary and secondary values
primary.trace_add('write', update_result)
secondary.trace_add('write', update_result)

root.mainloop()'''

import pandas as pd
from tkinter import *
from tkinter import ttk  # Import ttk for Treeview widget
from openpyxl import load_workbook

def update_result(*args):
    try:
        value1 = float(primary.get())
        value2 = float(secondary.get())
        result_var.set(value1 / value2)

        # Load the Excel file into a DataFrame
        excel_file = 'data.xlsx'
        df = pd.read_excel(excel_file)

        # Define the input value
        input_value = float(value1 / value2)  # Use the calculated result as the input value

        # Compare 'OP' column with the input value
        df['Comparison'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['KVA'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['kW'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        df['KVAr'] = df['OP'].apply(
            lambda x: 'Below' if x > input_value else ('Above' if x < input_value else 'Equal'))

        # Save the DataFrame back to Excel with the results in a new column
        output_file = 'output.xlsx'  # Change this to the desired output file path
        df.to_excel(output_file, index=False)

        # Load the updated Excel file
        wb2 = load_workbook(output_file)
        ws = wb2.active

        # Get data from specific columns
        data = []
        for row in ws.iter_rows(min_row=1, max_col=8, values_only=True):
            data.append(row)

        # Clear existing treeview data
        for item in treeview.get_children():
            treeview.delete(item)

        # Insert data into treeview
        for item in data[1:]:
            treeview.insert("", "end", values=item)

        print("Data loaded successfully.")
    except ValueError:
        result_var.set("Enter both two values!!!")

root = Tk()
root.title("Project")
root.geometry("800x800")
root.resizable(False, False)

primary = StringVar()
secondary = StringVar()
result_var = StringVar()

prim_label = Label(root, text="Enter the Primary value")
prim_label.place(x=350, y=20)
Secon_label = Label(root, text="Enter the Secondary value")
Secon_label.place(x=350, y=80)
prim_entry = Entry(root, textvariable=primary)
prim_entry.place(x=350, y=50)
secon_entry = Entry(root, textvariable=secondary)
secon_entry.place(x=350, y=110)
result_label = Label(root, textvariable=result_var)
result_label.place(x=360, y=130)

# Create a Frame for Treeview and Scrollbar
frame = Frame(root)
frame.place(x=10, y=200, width=780, height=500)

# Create a Treeview widget for displaying data in a table
treeview = ttk.Treeview(frame, columns=('S.no', 'Power', 'OP', 'CT/R', 'Above / Below', 'KVA', 'kW', 'KVAr'),
                        show='headings')
treeview.pack(side=LEFT, fill=BOTH, expand=True)

# Create a Scrollbar
scrollbar = Scrollbar(frame, orient=VERTICAL, command=treeview.yview)
scrollbar.pack(side=RIGHT, fill=Y)

# Configure the Treeview to use the Scrollbar
treeview.config(yscrollcommand=scrollbar.set)

# Define headings for the columns
treeview.heading('S.no', text='S.no')
treeview.heading('Power', text='Power')
treeview.heading('OP', text='OP')
treeview.heading('CT/R', text='CT/R')
treeview.heading('Above / Below', text='Above / Below')
treeview.heading('KVA', text='KVA')
treeview.heading('kW', text='kW')
treeview.heading('KVAr', text='KVAr')

# Set column widths
treeview.column('S.no', width=50)
treeview.column('Power', width=70)
treeview.column('OP', width=70)
treeview.column('CT/R', width=70)
treeview.column('Above / Below', width=100)
treeview.column('KVA', width=70)
treeview.column('kW', width=70)
treeview.column('KVAr', width=70)

# Bind the update_result function to changes in primary and secondary values
primary.trace_add('write', update_result)
secondary.trace_add('write', update_result)

root.mainloop()


