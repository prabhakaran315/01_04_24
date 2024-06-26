#----------------Importing the required libraries----------------#
import tkinter as tk
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from tkinter import messagebox
import math
from reportlab.lib.units import inch
import datetime

class TruePowerFactorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Project")
        self.create_widgets()

        #------------Creating the Data-------------------#
        self.data_1 = [
            ["ID", "Panel Rating", "Optimal CT Ratio", "Condition", "Minimum kVAr", "Minimum kW"],
            [1, "630A ASTRA - B 3", 1252, "", "", ""],
            [2, "630A ASTRA - B 5", 1809, "", "", ""],
            [3, "630A ASTRA - B10", 2782, "", "", ""],
            [4, "315A ASTRA - B 3", 835, "", "", ""],
            [5, "315A ASTRA - B 5", 1252, "", "", ""],
            [6, "315A ASTRA - B10", 1530, "", "", ""],
            [7, "420A ASTRA - B10", 1669, "", "", ""],
            [8, "420A ASTRA - B 5", 1252, "", "", ""],
            [9, "210A ASTRA - B10", 1530, "", "", ""],
            [10, "210A ASTRA - B 5", 696, "", "", ""],
            [11, "150A ASTRA - B 5", 835, "", "", ""],
            [12, "150A ASTRA - B10", 1669, "", "", ""]
        ]
        self.data_3 = [
            ["Optimal kW"],
            [225],
            [325],
            [500],
            [150],
            [225],
            [275],
            [300],
            [225],
            [275],
            [125],
            [150],
            [300]
        ]
        self.rows = len(self.data_1)
        self.columns = len(self.data_1[0])  # Number of columns is the length of the first rowz

    def create_widgets(self):
        # Frames
        self.frame_1 = tk.Frame(self.root, bg="Grey")
        self.frame_1.pack(expand=True, fill=tk.BOTH)
        self.frame_2 = tk.Frame(self.root, height=650, bg="White")
        self.frame_2.pack(expand=True, fill=tk.BOTH)
        self.frame_3 = tk.Frame(self.root, height=80, bg="Grey")
        self.frame_3.pack(expand=True, fill=tk.BOTH)

        self.primary = tk.StringVar()
        self.secondary = tk.StringVar()

        # Labels
        tk.Label(self.frame_1, text="Conditions to Achieve True Power Factor Performance", bg="Grey", fg="White", font=("Arial", 20, "bold")).pack(padx=10, pady=10)
        tk.Label(self.frame_1, text="Enter the Primary value of transformer", bg="Grey", fg="White", font=("Arial", 14)).pack()
        self.primary_entry = tk.Entry(self.frame_1, textvariable=self.primary,font=("Arial", 15))
        self.primary_entry.pack()

        self.label_1=tk.Label(self.frame_1, text="Enter the Secondary value of transformer", bg="Grey", fg="White", font=("Arial", 14)).pack()

        self.secondary_entry = tk.Entry(self.frame_1,textvariable=self.secondary, font=("Arial", 15))
        self.secondary_entry.pack()

        self.primary.trace_add('write', self.update_result)
        self.secondary.trace_add('write', self.update_result)

        # Buttons
        tk.Button(self.frame_1, text="Submit", command=self.submit, font=("Arial", 15, "bold")).pack(padx=10, pady=10)


    def submit(self):
        try:
            value_1 = float(self.primary_entry.get())
            value_2 = float(self.secondary_entry.get())

            if value_1 > 0 and value_2 > 0 :  # Corrected condition check
                cal_opt_val = round(value_1 / value_2)
                self.condition_1(value_1, value_2, cal_opt_val)
            else:
                messagebox.showinfo('Error!', "Kindly enter positive integers and greater than zero values!!!")
        except ValueError:
            messagebox.showinfo("Error!", "Kindly enter both values as numbers!!!")

#---------------Below code for without highlighting the table---------------#

    def condition_1(self, value_1, value_2, cal_opt_val):
        kva = math.floor(1.732 * 415.0 * (value_1 / 1000))
        kvar = math.ceil(kva * 0.03)
        kw = math.ceil(kva * 0.05)

        self.data_2 = [row[:] for row in self.data_1]
        # Update the "Condition" column in self.data_2 and clear appropriate columns
        for row_idx, row in enumerate(self.data_2[1:], start=1):
            data_val = row[2]  # Get the value from the third column of data_2
            if data_val < cal_opt_val:
                row[3:5] = "Condition - 1", kvar, kw
            else:
                if row_idx < len(self.data_3):
                    row[3:5] = "Condition - 2", "-", self.data_3[row_idx][0]

        # Clear previous labels in frame_2
        for widget in self.frame_2.winfo_children():
            widget.destroy()

        # Create Table using tkinter Labels using self.data_2 (updated data)
        for x in range(self.rows):
            for y in range(self.columns):
                font_style = ("Arial", 15, "bold") if x == 0 else ("Arial", 15)
                self.lab = tk.Label(self.frame_2, text=self.data_2[x][y], width=20, anchor="center", borderwidth=1, relief="solid", font=font_style).grid(row=x, column=y, padx=0, pady=0, sticky="nsew")
        self.first_three = []  # Initialize an empty list to store the filtered elements

        # Sort self.data_2 based on the 6th element
        sort_column = sorted(self.data_2, key=lambda x: str(x[5]))

        # Get the first three sorted elements
        first = sort_column[:3]
        # Iterate through sort_column and append elements to self.first_three if they meet the condition
        for row in sort_column:
            if str(row[5]) in [str(item[5]) for item in first]:  # Check if the 6th element is in the first three sorted elements
                self.first_three.append((row[0], row[1], row[5]))
        # Print the filtered elements
        print(self.first_three)

        # Display low, medium, high values in separate table format
        table_frame = tk.Frame(self.frame_2, bg="white")
        table_frame.grid(row=self.rows + 1, column=0, columnspan=3, pady=10)

        # Table heading
        tk.Label(table_frame, text="Optimal Panel Rating Based on Optimum kW", font=("Arial", 16, "bold"), bg = "White").grid(row=0,column=0, columnspan=5, padx=9, pady=5)

        # Labels for the table headers
        tk.Label(table_frame, text="Panel ID ", font=("Arial", 14, "bold"), bg = "White").grid(row=1, column=0, padx=5, pady=5)
        tk.Label(table_frame, text="Panel Rating", font=("Arial", 14, "bold"), bg = "White").grid(row=1, column=1, padx=5, pady=5)
        tk.Label(table_frame, text="Minimum kW", font=("Arial", 14, "bold"), bg = "White").grid(row=1, column=2, padx=5, pady=5)

        # Display low, medium, high values in the table format
        row_index = 2

        for i in range(min(3, len(self.first_three))):
            panel_id = self.first_three[i][0]
            panel_rating = self.first_three[i][1]
            panel_kw = self.first_three[i][2]  # Corrected index from 5 to 2
            self.pan_rat = (panel_id, panel_rating, panel_kw)

            #print("Condition -1 :", self.pan_rat)

            # Unpack pan_rat and create labels in the table_frame
            category, data, data_1 = self.pan_rat

            tk.Label(table_frame, text=category, font=("Arial", 14), bg="White").grid(row=row_index, column=0, padx=5, pady=5)
            tk.Label(table_frame, text=str(data), font=("Arial", 14), bg="White").grid(row=row_index, column=1, padx=5, pady=5)
            tk.Label(table_frame, text=str(data_1), font=("Arial", 14), bg="White").grid(row=row_index, column=2, padx=5, pady=5)
            row_index += 1
        self.hello = tk.IntVar()
        tk.Checkbutton(self.frame_3, text="", variable=self.hello, onvalue=1, offvalue=0, bg = "Grey").place(relx=0.42, rely=0.22, anchor=tk.NW)
        tk.Label(self.frame_3, text = ' If you want to print suggestion panel rating!!!', bg="Grey", fg = "White", font = ("Arial", 14)).place(relx=0.43, rely=0.2, anchor=tk.NW)
        tk.Button(self.frame_3, text="Print", command=self.export_pdf_condition_1, font=("Arial", 15, "bold")).place(relx=0.5, rely=0.5, anchor=tk.NW)

    def update_result(self, *args):
        # Clear previous labels in frame_2
        for widget in self.frame_2.winfo_children():
            widget.destroy()
        for widget_1 in self.frame_3.winfo_children():
            widget_1.destroy()

    def header(self, canvas, doc):
        canvas.saveState()
        canvas.drawImage('test.png', 40, 770, width=100, height=50)
        canvas.restoreState()
        # Add date and time to the header
        now = timestamp = datetime.datetime.now().strftime("%d/%m/%Y")
        header_style = ParagraphStyle(name='HeaderStyle', fontSize=12, textColor='black')
        header_text = Paragraph(f"{now}", header_style)
        header_text.wrapOn(canvas, inch, inch)
        header_text.drawOn(canvas, 500, 800)

    def export_pdf_condition_1(self):
        timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
        pdf_filename = f'{timestamp}.pdf'

        # Create a custom header with the image
        doc = BaseDocTemplate(pdf_filename, pagesize=A4)
        header_frame = Frame(doc.leftMargin, doc.topMargin, doc.width, doc.height)
        header_template = PageTemplate(id='header_template', frames=[header_frame], onPage=self.header)
        doc.addPageTemplates([header_template])

        # Define paragraph styles
        title_style = ParagraphStyle(name='TitleStyle', fontSize=14, textColor=colors.black, leading=15)
        body_style = ParagraphStyle(name='BodyStyle', fontSize=12, textColor=colors.black, leading=14)
        elements = []

        # Define paragraph styles
        styles = getSampleStyleSheet()
        title_style = styles['Title']
        body_style = styles['BodyText']

        # Add date and time to the header
        now = datetime.datetime.now().strftime("%d/%m/%Y")
        header_text = Paragraph(f"Date: {now}", body_style)
        elements.append(header_text)
        elements.append(Spacer(1, 12))

        # Add content about True Power Factor
        content = [
            Paragraph("True Power Factor", title_style),
            Spacer(1, 12),
            Paragraph(
                "The true power factor is a measure of how efficiently electrical power is being used. "
                "It represents the ratio of true power (measured in watts) to apparent power (measured in volt-amperes). "
                "A higher true power factor indicates better utilization of electrical power.", body_style),
            Spacer(1, 20)
        ]

        elements.extend(content)

        primary_value_paragraph = Paragraph('Primary value : {}'.format(self.primary_entry.get()), body_style)
        secondary_value_paragraph = Paragraph("Secondary value : {}".format(self.secondary_entry.get()), body_style)
        content.append(primary_value_paragraph)
        content.append(secondary_value_paragraph)

        content.append(Spacer(1, 20))

        data = [["ID", "Panel Rating", "Optimal CT Ratio", "Condition", "Minimum kVAr", "Minimum kW"]]

        data.extend([item[:4] + [item[4], item[5]] for item in self.data_2[1:]])

        t = Table(data, repeatRows=1)
        t.setStyle(
            TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0, -1), (-1, 0), 20), ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)]))

        elements = content + [t]

        elements.append(Spacer(1, 25))

        check_box = self.hello.get()
        # Inside the export_pdf_condition_1 method
        if check_box == 1:
            try:
                content_1 = [
                    Paragraph("Optimal Panel Rating Based on Optimum kW", title_style), Spacer(1, 5)
                    ]
                elements.extend(content_1)

                table_data = [["Panel ID", "Panel Rating", "Minimum kW"]]
                for i in range(3):
                    panel_id = self.first_three[i][0]
                    panel_rating = self.first_three[i][1]
                    panel_kw = self.first_three[i][2]
                    self.pan_rat = (panel_id, panel_rating, panel_kw)

                    for row in [self.pan_rat]:
                        #print("DEBUG: Processing row:", row)  # Print each row to see which rows are being processed
                        table_data.append(list(row))  # Append the row to table_data
                        #print("DEBUG: Updated table_data:", table_data)  # Print table_data after each iteration
                    # Append lists, not tuples

                    # Create the table with the formatted table_data
                table = Table(table_data, repeatRows=1)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, -1), (-1, 0), 20),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)]
                ))
                elements.append(table)  # Append the table to elements only once

            except Exception as e:
                print("Error:", e)

            # Build the PDF document
            doc.build(elements)
            messagebox.showinfo("Done", "PDF file exported Successfully!!!")

        else:
            doc.build(elements)
            messagebox.showinfo("Done", "PDF file exported Successfully!!!")

if __name__ == "__main__":
    root = tk.Tk()
    app = TruePowerFactorApp(root)
    root.mainloop()