#----------------Importing the required libraries----------------#
import tkinter as tk
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame
from reportlab.lib.styles import ParagraphStyle
from tkinter import messagebox, NW
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
            ["ID", "Panel Rating", "Optimal CT Ratio", "Condition", "kVA", "kVAr", "kW", "Optimal kW"],
            [1, "630A ASTRA - B3", 1252, "", "", "", "", 225],
            [2, "630A ASTRA - B5", 1809, "", "", "", "", 325],
            [3, "630A ASTRA - B10", 2782, "", "", "", "", 500],
            [4, "315A ASTRA - B3", 835, "", "", "", "", 150],
            [5, "315A ASTRA - B5", 1252, "", "", "", "", 225],
            [6, "315A ASTRA - B10", 1530, "", "", "", "", 275],
            [7, "420A ASTRA - B10", 1669, "", "", "", "", 300],
            [8, "420A ASTRA - B5", 1252, "", "", "", "", 225],
            [9, "210A ASTRA - B10", 1530, "", "", "", "", 275],
            [10, "210A ASTRA - B5", 696, "", "", "", "", 125],
            [11, "150A ASTRA - B5", 835, "", "", "", "", 150],
            [12, "150A ASTRA - B10", 1669, "", "", "", "", 300]
        ]
        self.rows = len(self.data_1)
        self.columns = len(self.data_1[0])  # Number of columns is the length of the first row

    def create_widgets(self):
        # Frames
        self.frame_1 = tk.Frame(self.root, bg="Grey")
        self.frame_1.pack(expand=True, fill=tk.BOTH)
        self.frame_2 = tk.Frame(self.root, height=650, bg="White")
        self.frame_2.pack(expand=True, fill=tk.BOTH)
        self.frame_3 = tk.Frame(self.root, height=80, bg="Grey")
        self.frame_3.pack(expand=True, fill=tk.BOTH)

        # Labels
        tk.Label(self.frame_1, text="Conditions to Achieve True Power Factor Performance", bg="Grey", fg="White", font=("Arial", 20, "bold")).pack(padx=10, pady=10)
        tk.Label(self.frame_1, text="Enter the Primary value of transformer", bg="Grey", fg="White", font=("Arial", 14)).pack()
        self.primary_entry = tk.Entry(self.frame_1, font=("Arial", 14))
        self.primary_entry.pack()

        tk.Label(self.frame_1, text="Enter the Secondary value of transformer", bg="Grey", fg="White", font=("Arial", 14)).pack()

        self.secondary_entry = tk.Entry(self.frame_1, font=("Arial", 14))
        self.secondary_entry.pack()

        # Buttons
        tk.Button(self.frame_1, text="Submit", command=self.submit, font=("Arial", 15, "bold")).pack(padx=10, pady=10)

    def submit(self):
        try:
            value_1 = float(self.primary_entry.get())
            value_2 = float(self.secondary_entry.get())
            cal_opt_val = round(value_1 / value_2)

            if value_1 >= 0 and value_2 >= 0:  # Corrected condition check
                self.condition_1(value_1, value_2, cal_opt_val)
            else:
                messagebox.showinfo('Error!', "Kindly enter positive integers!!!")
        except ValueError:
            messagebox.showinfo("Error!", "Kindly enter both values as numbers!!!")

    def condition_1(self, value_1, value_2, cal_opt_val):
        kva = math.floor(1.732 * 415.0 * (value_1 / 1000))
        kvar = math.ceil(kva * 0.03)
        kw = math.ceil(kva * 0.05)

        self.data_2 = [row[:] for row in self.data_1]
        # Update the "Condition" column in self.data_2 and clear appropriate columns
        for row in self.data_2[1:]:
            data_val = row[2]  # Get the value from the third column of data_2s
            if data_val < cal_opt_val:
                row[3:7] = "Condition - 1", kva, kvar, kw,"-"
            else:
                row[3:7] = "Condition - 2", "-", "-", "-", row[7]

        # Clear previous labels in frame_2
        for widget in self.frame_2.winfo_children():
            widget.destroy()

        # Create Table using tkinter Labels using self.data_2 (updated data)
        for x in range(self.rows):
            for y in range(self.columns):
                font_style = ("Arial", 15, "bold") if x == 0 else ("Arial", 15)
                tk.Label(self.frame_2, text=self.data_2[x][y], width=20, anchor="center", borderwidth=1, relief="solid", font=font_style).grid(row=x, column=y, padx=0, pady=0, sticky="nsew")
        tk.Button(self.frame_3, text="Print", command=self.export_pdf_condition_1, font=("Arial", 15, "bold")).place(relx=0.5, rely=0.5, anchor=NW)

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

        # Create content about True Power Factor
        content = []
        content.append(Spacer(1, 20))

        content.append(Paragraph("True Power Factor:", title_style))
        content.append(Spacer(1, 10))
        content.append(Paragraph(
            "The true power factor is a measure of how efficiently electrical power is being used. "
            "It represents the ratio of true power (measured in watts) to apparent power (measured in volt-amperes)."
            " A higher true power factor indicates better utilization of electrical power.", body_style))

        content.append(Spacer(1, 20))

        # Create the Paragraph with the bold "Primary value :" text
        primary_value_paragraph = Paragraph('Primary value : {}'.format(self.primary_entry.get()), body_style)
        secondary_value_paragraph = Paragraph("Secondary value : {}".format(self.secondary_entry.get()), body_style)
        content.append(primary_value_paragraph)
        content.append(secondary_value_paragraph)

        content.append(Spacer(1, 20))

        data = [["ID", "Panel Rating", "Optimal CT Ratio", "Condition", "kVA", "kVAr", "kW", "Optimal kW"]]

        data.extend([item[:4] + [item[4], item[5], item[6], item[7]] for item in self.data_2[1:]])

        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('BOTTOMPADDING', (0, -1), (-1, 0), 20), ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))

        elements = content + [t]

        # Build the PDF document
        doc.build(elements)

        messagebox.showinfo("Done", "PDF file exported Successfully!!!")

if __name__ == "__main__":
    root = tk.Tk()
    app = TruePowerFactorApp(root)
    root.mainloop()