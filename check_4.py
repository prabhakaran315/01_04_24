import tkinter as tk
from tkinter import messagebox, ttk, NW
import math
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Spacer
from reportlab.lib import colors
import datetime
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph
from reportlab.lib.styles import ParagraphStyle

class TruePowerFactorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Project")

        self.create_widgets()
        self.data_1 = [
            ["ID", "Panel Rating", "Optimal CT Ratio", "Condition", "kVA", "kVAr", "kW"],
            [1,  "630A ASTRA - B3", 1252, "", "", "", ""],
            [2,  "630A ASTRA - B5", 1809, "", "", "", ""],
            [3,  "630A ASTRA - B10", 2782,"", "", "", ""],
            [4,  "315A ASTRA - B3", 835,"", "", "", ""],
            [5,  "315A ASTRA - B5", 1252,"", "", "", ""],
            [6,  "315A ASTRA - B10", 1530,"", "", "", ""],
            [7,  "420A ASTRA - B10", 1669,"", "", "", ""],
            [8,  "420A ASTRA - B5", 1252,"", "", "", ""],
            [9,  "210A ASTRA - B10", 1530,"", "", "", ""],
            [10, "210A ASTRA - B5", 696,"", "", "", ""],
            [11, "150A ASTRA - B5", 835,"", "", "", ""],
            [12, "150A ASTRA - B10", 1669,"", "", "", ""]
        ]

        self.data_2 = [
            ["ID", "Panel Rating", "Optimal CT Ratio", "Condition","Optimal kW"],
            [1, "630A ASTRA - B3", 1252,"",225],
            [2, "630A ASTRA - B5", 1809,"", 325],
            [3, "630A ASTRA - B10", 2782,"", 500],
            [4, "315A ASTRA - B3", 835, "",150],
            [5, "315A ASTRA - B5", 1252, "",225],
            [6, "315A ASTRA - B10", 1530,"", 275],
            [7, "420A ASTRA - B10", 1669, "",300],
            [8, "420A ASTRA - B5", 1252, "",225],
            [9, "210A ASTRA - B10", 1530, "",275],
            [10, "210A ASTRA - B5", 696, "",125],
            [11, "150A ASTRA - B5", 835, "",150],
            [12, "150A ASTRA - B10", 1669, "",300]
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
        self.primary_entry = tk.Entry(self.frame_1, font=("Arial", 12))
        self.primary_entry.pack()

        tk.Label(self.frame_1, text="Enter the Secondary value of transformer", bg="Grey", fg="White", font=("Arial", 14)).pack()

        self.secondary_entry = tk.Entry(self.frame_1, font=("Arial", 12))
        self.secondary_entry.pack()

        # Buttons
        tk.Button(self.frame_1, text="Submit", command=self.submit, font=("Arial", 12, "bold")).pack(padx=10, pady=10)

    def submit(self):
        try:
            value_1 = float(self.primary_entry.get())
            value_2 = float(self.secondary_entry.get())
            cal_opt_val = round(value_1 / value_2)

            if value_1 and value_2 >= 0:

                third_column = [row[2] for row in self.data_1[1:]]  # Extract the third column values

                for i in range(len(third_column)):
                    data_val = third_column[i]  # Get the value from the third column of data_1

                    if data_val < cal_opt_val:
                        self.condition_1()
                        break  # Stop looping if condition_1 is triggered
                    else:
                        self.condition_2()

            else:
                messagebox.showinfo('Error!', "Kindly enter the Positive integers!!!")

        except ValueError:
            messagebox.showinfo("Error!", "Kindly Enter the Both Values!!!")

    def condition_1(self):
        value_1 = float(self.primary_entry.get())
        value_2 = float(self.secondary_entry.get())
        cal_opt_val = round(value_1 / value_2)

        self.kva = math.floor(1.732 * 415.0 * (value_1 / 1000))
        self.kvar = math.ceil(self.kva * 0.03)
        self.kw = math.ceil(self.kva * 0.05)

        # Update the "Condition" column in self.data_1
        for row in self.data_1[1:]:
            third_column = [row[2] for row in self.data_1[1:]]  # Extract the third column values

            for i in range(len(third_column)):
                data_val = third_column[i]  # Get the value from the third column of data_1

                if data_val < cal_opt_val:
                    row[3] = "Condition - 1"
                    row[4] = self.kva
                    row[5] = self.kvar
                    row[6] = self.kw
                    break  # Stop looping if condition_1 is triggered
                else:
                    self.condition_2()
                    row[3] = "Condition - 2"
                    row[4] = self.kva
                    row[5] = self.kvar
                    row[6] = self.kw

        # Clear previous labels in frame_2
        for widget in self.frame_2.winfo_children():
            widget.destroy()

        # Create Table using ttk.Labels
        for x in range(self.rows):
            for y in range(self.columns):
                font_style = ("Arial", 15, "bold") if x == 0 else ("Arial", 15)
                ttk.Label(self.frame_2, text=self.data_1[x][y], width=20, anchor="center", borderwidth=1,relief="solid",font= font_style).grid(row=x, column=y, padx=0, pady=0, sticky="nsew")
                # Configure grid to center the label
                self.frame_2.grid_columnconfigure(y, weight=1)
                self.frame_2.grid_rowconfigure(x, weight=1)
                # Create and grid the label with the appropriate font style
                #ttk.Label(self.frame_2, text=self.data_1[x][y], width=20, anchor="center", borderwidth=1, relief="solid",font=("Arial", 15, "bold")).grid(row=x, column=y, padx=0, pady=0, sticky="nsew")


        tk.Button(self.frame_3, text="Print", command=self.export_pdf_condition_1, font=("Arial", 12, "bold")).place(relx=0.5, rely=0.5, anchor=NW)

    def condition_2(self):
        value_1 = float(self.primary_entry.get())
        value_2 = float(self.secondary_entry.get())
        cal_opt_val = round(value_1 / value_2)
        self.kva = math.floor(1.732 * 415.0 * (value_1 / 1000))
        self.kw = math.ceil(self.kva * 0.05)
        self.condition = "Condition - 2"

        for row in self.data_2[1:]:
            row[3] = "Condition - 2"

        # Clear previous labels in frame_2
        for widget in self.frame_2.winfo_children():
            widget.destroy()

        # Create Table using ttk.Labels
        for x in range(self.rows):
            for y in range(len(self.data_2[x])):  # Use len(self.data_2[x]) to get the correct number of columns
                font_style = ("Arial", 15, "bold") if x == 0 else ("Arial", 15)
                ttk.Label(self.frame_2, text=self.data_2[x][y], width=20, anchor="center", borderwidth=1,relief="solid", font=font_style).grid(row=x, column=y, padx=0, pady=0, sticky="nsew" )
                # Configure grid to center the label
                self.frame_2.grid_columnconfigure(y, weight=1)
                self.frame_2.grid_rowconfigure(x, weight=1)
        tk.Button(self.frame_3, text="Print", command=self.export_pdf_condition_2, font=("Arial", 12, "bold")).place(relx=0.5, rely=0.5, anchor=NW)

    def header(self, canvas, doc):
        canvas.saveState()
        canvas.drawImage('test.png', 40, 770, width=100, height=50)
        canvas.restoreState()

    import datetime
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import BaseDocTemplate, Frame, Paragraph, PageTemplate, Spacer, Table, TableStyle

    def export_pdf_condition_1(self):
        # Generate a timestamp for the PDF filename
        timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
        pdf_filename = f'{timestamp}.pdf'

        # Create a custom header with the image
        def header(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica-Bold', 12)
            canvas.drawString(40, doc.height + doc.topMargin - 20, "My Custom Header")
            canvas.restoreState()

        doc = BaseDocTemplate(pdf_filename, pagesize=A4)
        header_frame = Frame(doc.leftMargin, doc.height + doc.topMargin - 40, doc.width, 40)
        header_template = PageTemplate(id='header_template', frames=[header_frame],
                                       onPage=lambda canvas, doc, content=None: header(canvas, doc))
        doc.addPageTemplates([header_template])

        # Define paragraph styles
        title_style = ParagraphStyle(name='TitleStyle', fontSize=14, textColor=colors.black, leading=16)
        body_style = ParagraphStyle(name='BodyStyle', fontSize=12, textColor=colors.black, leading=14)

        # Create content about True Power Factor
        content = []
        content.append(Paragraph("True Power Factor:", title_style))
        content.append(Paragraph(
            "The true power factor is a measure of how efficiently electrical power is being used. "
            "It represents the ratio of true power (measured in watts) to apparent power (measured in volt-amperes)."
            " A higher true power factor indicates better utilization of electrical power.", body_style))

        content.append(Spacer(1, 14))

        # Assuming data_1 is defined elsewhere
        data = [["Id", "Panel Rating", "Optimal CT Ratio", "Condition", "kVA", "kVAr", "kW"]]

        data.extend([item[:4] + [item[4], item[5], item[6]] for item in self.data_1[1:]])

        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                               ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))

        # Add the content about True Power Factor and the table to the PDF elements
        elements = content + [t]

        # Build the PDF document
        doc.build(elements)

        messagebox.showinfo("Done", "PDF file exported Successfully!!!")

    def export_pdf_condition_2(self):
        timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
        pdf_filename = f'{timestamp}.pdf'

        # Create a custom header with the image
        doc = BaseDocTemplate(pdf_filename, pagesize=A4)
        header_frame = Frame(doc.leftMargin, doc.topMargin, doc.width, doc.height)
        header_template = PageTemplate(id='header_template', frames=[header_frame], onPage=self.header)
        doc.addPageTemplates([header_template])

        # Define paragraph styles
        title_style = ParagraphStyle(name='TitleStyle', fontSize=14, textColor=colors.black, leading=16)
        body_style = ParagraphStyle(name='BodyStyle', fontSize=12, textColor=colors.black, leading=14)

        # Create content about True Power Factor
        content = []
        content.append(Paragraph("True Power Factor:", title_style))
        content.append(Paragraph(
            "The true power factor is a measure of how efficiently electrical power is being used. "
            "It represents the ratio of true power (measured in watts) to apparent power (measured in volt-amperes)."
            " A higher true power factor indicates better utilization of electrical power.", body_style))

        content.append(Spacer(1, 14))

        data = [["Id", "Panel Rating", "Optimal CT Ratio", "Condition", "Optimal kW"]]
        data.extend([item[:4] + [item[4]] for item in self.data_2[1:]])

        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                               ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))

        elements = content + [t]

        # Build the PDF document
        doc.build(elements)

        messagebox.showinfo("Done", "PDF file exported Successfully!!!")

if __name__ == "__main__":
    root = tk.Tk()
    app = TruePowerFactorApp(root)

    root.mainloop()
