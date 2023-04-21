import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox
from ttkwidgets.autocomplete import AutocompleteCombobox
from datetime import *
import xlsxwriter
import os

# Small app for creating evaluation reports related to heat treatment processes as quenching, carburising, carbonitriding,
# nitriding and nitrocarburising in Bodycote company.

class AskTypeWindow:
    def __init__(self):
        self.asktype = tk.Tk()
        self.asktype.title("HT Report")
        self.asktype.resizable(width=False, height=False)
        self.width = 260
        self.height = 160
        self.sc_width = self.asktype.winfo_screenwidth()
        self.sc_height = self.asktype.winfo_screenheight()
        self.x = (self.sc_width/2) - (self.width/2)
        self.y = (self.sc_height/3) - (self.height/3)
        self.askLabel = tk.Label(self.asktype, text="Choose report type: ", font=("aerial", 10), bg="grey", fg="white")
        self.askLabel.grid(row=0, column=1, columnspan=1, sticky="N", ipadx=70)
        self.reporttype = tk.IntVar()
        self.P1 = tk.Radiobutton(self.asktype, text="Quenching / Annealing", variable=self.reporttype, value=1)
        self.P1.grid(row=1, column=1, ipady=4, ipadx=50, sticky="NE")
        self.P2 = tk.Radiobutton(self.asktype, text="Carburising / Carbonitriding", variable=self.reporttype, value=2)
        self.P2.grid(row=2, column=1, ipady=4, ipadx=25, sticky="NE")
        self.P3 = tk.Radiobutton(self.asktype, text="Nitriding / Nitrocarburising", variable=self.reporttype, value=3)
        self.P3.grid(row=3, column=1, ipady=4, ipadx=30, sticky="NE")

        self.OKbutton = tk.Button(self.asktype, text="OK", command=self.okey)
        self.OKbutton.grid(row=4, column=1, rowspan=4, columnspan=4, padx=4, pady=4, ipadx=40)

        self.asktype.geometry('%dx%d+%d+%d' % (self.width, self.height, self.x, self.y))

    def ask_type_run(self):
        self.asktype.mainloop()

    def okey(self):
        if self.reporttype.get() == 1:
            self.asktype.destroy()
            QuenRep = Quenching("Quenching / Annealing Report", 'header_quenching-annealing.png', 430, 'Ret.austenite:', 'nothing')
            QuenRep.run()
        elif self.reporttype.get() == 2:
            self.asktype.destroy()
            CarbRep = Carburizing("Carburising / Carbonitriding", 'header_carburising-carbonitriding.png', 290, 'Ret.austenite:',
                                       'Int.oxidation:')
            CarbRep.run()
        elif self.reporttype.get() == 3:
            self.asktype.destroy()
            NitRep = Nitriding("Nitriding / Nitrocarburising", 'header_nitriding-nitrocarburising.png', 290, 'CLT:', 'Oxid. layer:')
            NitRep.run()


class BaseView:
    methods = ['HRA', 'HRB', 'HRC', 'HRD', 'HRF', 'HRG', 'HR15N', 'HR30N', 'HR45N', 'HR15T', 'HR15T/0.25mm', 'HR15T/0.51mm', 'HR15T/1.02mm',
               'HR30T', 'HR30T/0.51mm', 'HR30T/1.02mm', 'HR45T', 'HR15W', 'HV 0.001', 'HV 0.002', 'HV 0.005', 'HV 0.01', 'HV 0.02', 'HV 0.025',
               'HV 0.05', 'HV 0.1', 'HV 0.2', 'HV 0.25', 'HV 0.3', 'HV 0.5', 'HV 1', 'HV 2', 'HV 2 (UCI)', 'HV 2.5', 'HV 3', 'HV 3 (UCI)',
               'HV 5', 'HV 5 (UCI)', 'HV 10', 'HV 10 (UCI)', 'HV 20', 'HV 30', 'HV 40', 'HV (k.A.)', 'HV 50', 'HV 60', 'HV 100', 'HV 120',
               'HBW 1/1', 'HBW 1/2.5', 'HBW 2/4', 'HBW 1/5', 'HBW 2.5/6.25', 'HBW 1/10', 'HBW 2/10', 'HBW 2.5/15.625', 'HBW 2/20', 'HBW 5/25',
               'HBW 1/30', 'HBW 2.5/31.25', 'HBW 2/40', 'HBW 2.5/62.5', 'HBW 5/62.5', 'HBW 10/100', 'HBW 2/120', 'HBW 5/125', 'HBW 2.5/187.5',
               'HBW 10/250', 'HBW 5/250', 'HBW 10/500', 'HBW 5/750', 'HBW 10/1000', 'HBW 10/1500', 'HBW (k.A.)', 'HBW 10/3000', 'HLC', 'HLD',
               'HLE', 'HLG', 'HLS', 'HLDC', 'HLDL', 'HLDL+15', 'HK 0.01', 'HK 0.025', 'HK 0.05', 'HK 0.1', 'HK 0.2', 'HK 0.3', 'HK 0.5', 'HK 1',
               'HK 2', 'MPa', 'N/mm²', 'GPa', 'Ib/in²', 'psi', 'ksi', 'tsi', 'tt', 'N/m²', 'kg/mm²', 'kJ']

    other_units = ['nm', 'µm', 'mm', 'cm', '%']

    localaddress = "C:\\Users\\lukas\\python_files\\HT Report\\Archive\\"
    imgs_address = "C:\\Users\\lukas\\python_files\\HT Report\\imgs\\"

    def __init__(self, title, image):
        # Window settings
        self.cellwidth = int(30)
        self.cellheight = int(1)
        self.qcore = tk.Tk()
        self.qcore.geometry("1000x744")
        self.qcore.title(title)
        self.qcore.rowconfigure(1, weight=1)
        self.qcore.columnconfigure(1, weight=1)
        self.qcore.resizable(width=False, height=False)
        self.width = 1000
        self.height = 744
        self.sc_width = self.qcore.winfo_screenwidth()
        self.sc_height = self.qcore.winfo_screenheight()
        self.x = (self.sc_width / 2) - (self.width / 2)
        self.y = (self.sc_height / 3) - (self.height / 3)
        self.qcore.geometry('%dx%d+%d+%d' % (self.width, self.height, self.x, self.y))
        # Header and Banner
        self.header = tk.Frame(self.qcore, width=1000, height=140, bg="white")
        self.header.grid(row=0, sticky="NW")
        self.header.grid_propagate(False)
        self.bg_load = tk.PhotoImage(file=image)
        self.bg_image = tk.Label(self.header, image=self.bg_load)
        self.bg_image.grid()
        # Date and time
        self.date = datetime.today()
        self.reportdate = self.date.strftime("%d-%b-%Y")
        self.year = self.date.strftime("%Y")
        # Workspace and Order info
        self.workspace = tk.Frame(self.qcore, width=994, height=130, bg="light grey")
        self.workspace.grid(row=1, sticky="NW", ipadx=16, ipady=0)
        self.workspace.grid_propagate(False)
        self.label_names = ['Customer:', 'Bodycote No.:', 'Part name:', 'Quantity:', 'Order No.:', 'Order description:', 'Material:',
                            'Requirements:']
        self.column = 0
        for name, row in zip(self.label_names, range(8)):
            if row <= 3:
                self.label = tk.Label(self.workspace, text=name, anchor="w", bg="light grey")
                self.label.grid(row=row, column=self.column, padx=10, ipady=4)
            else:
                self.column = 2
                self.label = tk.Label(self.workspace, text=name, anchor="w", bg="light grey")
                self.label.grid(row=row - 4, column=self.column, padx=10, ipady=4)

        # Temporary containers for autofilling parameters
        self.customers = []
        self.materials = []
        self.controllers = []

        self.unpacking_dat()

        self.customer_entry = AutocompleteCombobox(self.workspace, width=22, completevalues=self.customers)
        self.customer_entry.grid(row=0, column=1)
        self.order_entry = tk.Entry(self.workspace, width=24)
        self.order_entry.grid(row=0, column=3, sticky="W")
        self.dispathnote_entry = tk.Entry(self.workspace, width=24)
        self.dispathnote_entry.grid(row=1, column=1)
        self.other_desc_entry = tk.Entry(self.workspace, width=24)
        self.other_desc_entry.grid(row=1, column=3, sticky="W")
        self.partname_entry = tk.Entry(self.workspace, width=24)
        self.partname_entry.grid(row=2, column=1)
        self.material_entry = AutocompleteCombobox(self.workspace, width=22, completevalues=self.materials)
        self.material_entry.grid(row=2, column=3, sticky="W")
        self.quantity_entry = tk.Entry(self.workspace, width=24)
        self.quantity_entry.grid(row=3, column=1)
        self.requirements_entry = tk.Text(self.workspace, width=74, height=2)
        self.requirements_entry.grid(row=3, column=3, pady=2)

    def unpacking_dat(self):
        with open("dat_customer.txt", "r") as dat_customers:
            for cust in dat_customers.read().split("**"):
                self.customers.append(cust)

        with open("dat_material.txt", "r") as dat_materials:
            for mat in dat_materials.read().split("**"):
                self.materials.append(mat)

        with open("dat_controller.txt", "r") as dat_controllers:
            for cont in dat_controllers.read().split("**"):
                self.controllers.append(cont)


class Quenching(BaseView):
    def __init__(self, title, image, height, add_one_p, add_two_p):
        super().__init__(title, image)
        # Results
        self.results = tk.Frame(self.qcore, width=994, height=height, bg="gray67")
        self.results.grid(row=3, sticky="NW", padx=2)
        self.results.grid_propagate(False)
        self.results_label = tk.Label(self.results, text="Results:", anchor="e", font=("aerial", 16, "bold"), fg="black", bg="gray67")
        self.results_label.grid(row=0, column=0, ipady=4)
        # Surface hardness
        self.sh_container = []
        self.sh_method_container = []
        self.surface_hardness = tk.Label(self.results, text="Surface hardness:", anchor="e", bg="gray67")
        self.surface_hardness.grid(row=1, column=0, padx=5)
        self.surface_hardness_entry = tk.Entry(self.results, width=24)
        self.surface_hardness_entry.grid(row=1, column=1, columnspan=3, padx=10, pady=4, ipady=2, sticky="W")
        self.sh_container.append(self.surface_hardness_entry)
        self.sh_methods = AutocompleteCombobox(self.results, width=15, completevalues=self.methods)
        self.sh_methods.grid(row=1, column=4, ipady=2, pady=4, sticky="W")
        self.sh_method_container.append(self.sh_methods)
        self.sh_plus = tk.Button(self.results, text="+", command=self.plussurface, font=("aerial", 11, "bold"), fg="black")
        self.sh_plus.grid(row=1, column=5, padx=10, ipadx=4, pady=2, sticky="W")
        # Core hardness
        self.ch_container = []
        self.ch_method_container = []
        self.core_hardness = tk.Label(self.results, text="Core hardness:", anchor="e", bg="gray67")
        self.core_hardness.grid(row=2, column=0, padx=10)
        self.core_hardness_entry = tk.Entry(self.results, width=24)
        self.core_hardness_entry.grid(row=2, column=1, columnspan=3, padx=10, pady=4, ipady=2, sticky="W")
        self.ch_container.append(self.core_hardness_entry)
        self.ch_methods = AutocompleteCombobox(self.results, width=15, completevalues=self.methods)
        self.ch_methods.grid(row=2, column=4, ipady=2, pady=4, sticky="W")
        self.ch_method_container.append(self.ch_methods)
        self.ch_plus = tk.Button(self.results, text="+", command=self.pluscore, font=("aerial", 11, "bold"), fg="black")
        self.ch_plus.grid(row=2, column=5, padx=10, ipadx=4, pady=2, sticky="W")
        self.sh_iterator = [0, ]
        self.ch_iterator = [0, ]
        # Additional parameters
        self.add_one = tk.Label(self.results, text=add_one_p, anchor="e", bg="gray67")
        self.add_one.grid(row=3, column=0, padx=10)
        self.add_one_entry = tk.Entry(self.results, width=8)
        self.add_one_entry.grid(row=3, column=1, padx=10, pady=4, ipady=2, sticky="W")
        self.add_one_methods = AutocompleteCombobox(self.results, width=8, completevalues=self.other_units)
        self.add_one_methods.grid(row=3, column=2, ipady=2, pady=4, sticky="W")
        self.add_two = tk.Label(self.results, text=add_two_p, anchor="e", bg="gray67")
        self.add_two.grid(row=4, column=0, padx=10)
        self.add_two_entry = tk.Entry(self.results, width=8)
        self.add_two_entry.grid(row=4, column=1, padx=10, pady=4, ipady=2, sticky="W")
        self.add_two_methods = AutocompleteCombobox(self.results, width=8, completevalues=self.other_units)
        self.add_two_methods.grid(row=4, column=2, ipady=2, pady=4, sticky="W")
        # Notes
        self.notes = tk.Label(self.results, text="Notes:", anchor="e", bg="gray67")
        self.notes.grid(row=5, column=0, padx=10)
        self.notes_entry = tk.Text(self.results, width=33, height=4)
        self.notes_entry.grid(row=5, column=1, padx=10, pady=4, sticky="W", columnspan=4)
        self.report_status = tk.Label(self.results, text="Report status:", anchor="e", bg="gray67")
        self.report_status.grid(row=6, column=0, padx=10, sticky="W")
        self.reportstatus = tk.IntVar()
        self.P1 = tk.Radiobutton(self.results, text="OK", variable=self.reportstatus, value=1, anchor="w", bg="gray67")
        self.P1.grid(row=6, column=1, ipady=4, padx=10, sticky="W")
        self.P2 = tk.Radiobutton(self.results, text="NOK", variable=self.reportstatus, value=2, anchor="w", bg="gray67")
        self.P2.grid(row=6, column=2, ipady=4, padx=10, sticky="W")
        self.controller = tk.Label(self.results, text="Controller:", anchor="e", bg="gray67")
        self.controller.grid(row=7, column=0, padx=10, sticky="W")
        self.controller_entry = AutocompleteCombobox(self.results, width=20, completevalues=self.controllers)
        self.controller_entry.grid(row=7, column=1, columnspan=2, padx=10, pady=4, ipady=2, sticky="W")
        # Buttons
        self.for_buttons = tk.Frame(self.qcore, width=994, height=200, bg="gray60")
        self.for_buttons.grid(row=4, sticky="NW", pady=1, padx=2)
        self.for_buttons.columnconfigure(0, weight=3)
        self.savebutton = tk.Button(self.for_buttons, text="Save report", command=self.get_data, height=2, width=30)
        self.savebutton.grid(row=0, column=1, columnspan=1, sticky="EW", padx=60)
        self.print_button = tk.Button(self.for_buttons, text='Print', command=self.print_it, state=tk.DISABLED, height=2, width=30)
        self.print_button.grid(row=0, column=2, columnspan=1, sticky="EW", padx=48)
        self.delete_button = tk.Button(self.for_buttons, text= 'Reset cells', command=self.clean_cells, height=2, width=30)
        self.delete_button.grid(row=0, column=3, columnspan=1, sticky="EW", padx=60)
        print(self.sh_iterator)
        # Logical test
        self.sh_test = False
        self.ch_test = False

    def updating_database(self):
        if self.customer_entry.get() not in self.customers:
            self.customers.append(self.customer_entry.get())

        if self.material_entry.get() not in self.materials:
            self.materials.append(self.material_entry.get())

        if self.controller_entry.get() not in self.controllers:
            self.controllers.append(self.controller_entry.get())

        with open("dat_customer.txt", "w") as dat_customer:
            dat_customer.write('**'.join(self.customers))

        self.customer_entry.configure(completevalues=self.customers)

        with open("dat_material.txt", "w") as dat_material:
            dat_material.write('**'.join(self.materials))

        self.material_entry.configure(completevalues=self.materials)

        with open("dat_controller.txt", "w") as dat_controller:
            dat_controller.write('**'.join(self.controllers))

        self.controller_entry.configure(completevalues=self.controllers)

    def run(self):
        self.qcore.mainloop()

    def plussurface(self):
        self.sh_plus.destroy()
        self.surface_hardness_entry = tk.Entry(self.results, width=24)
        self.surface_hardness_entry.grid(row=1, column=5, columnspan=3, padx=10, pady=4, ipady=2, sticky="W")
        self.sh_container.append(self.surface_hardness_entry)
        self.sh_methods = AutocompleteCombobox(self.results, width=12, completevalues=self.methods)
        self.sh_methods.grid(row=1, column=8, ipady=2, sticky="W")
        self.sh_method_container.append(self.sh_methods)
        self.sh_iterator.append(2)
        print(self.sh_iterator)
        self.sh_test = True
        return self.sh_iterator, self.sh_container, self.sh_method_container, self.sh_test

    def pluscore(self):
        self.ch_plus.destroy()
        self.core_hardness_entry = tk.Entry(self.results, width=24)
        self.core_hardness_entry.grid(row=2, column=5, columnspan=3, padx=10, pady=4, ipady=2, sticky="W")
        self.ch_container.append(self.core_hardness_entry)
        self.ch_methods = AutocompleteCombobox(self.results, width=12, completevalues=self.methods)
        self.ch_methods.grid(row=2, column=8, ipady=2, sticky="W")
        self.ch_method_container.append(self.ch_methods)
        self.ch_iterator.append(2)
        print(self.ch_container)
        self.ch_test = True
        return self.ch_iterator, self.ch_container, self.ch_method_container

    def print_it(self):
        os.startfile(self.localaddress + self.customer_entry.get().capitalize() + "\\" + str(self.dispathnote_entry.get()) + "-Report.xlsx", 'print')

    def clean_cells(self):
        self.customer_entry.delete(0, 'end')
        self.order_entry.delete(0, 'end')
        self.dispathnote_entry.delete(0, 'end')
        self.other_desc_entry.delete(0, 'end')
        self.partname_entry.delete(0, 'end')
        self.material_entry.delete(0, 'end')
        self.quantity_entry.delete(0, 'end')
        self.requirements_entry.delete('1.0', 'end')
        self.notes_entry.delete('1.0', 'end')
        self.add_one_entry.delete(0, 'end')
        self.add_one_methods.delete(0, 'end')
        self.add_two_entry.delete(0, 'end')
        self.add_two_methods.delete(0, 'end')
        self.controller_entry.delete(0, 'end')

        # Hardness remove
        # Surface
        for entry1, method1, in zip(self.sh_container, self.sh_method_container):
            entry1.delete(0, 'end')
            method1.delete(0, 'end')

        if self.sh_test == True:
            self.sh_iterator.pop()
            for hard, meth in zip(self.sh_container[1:], self.sh_method_container[1:]):
                hard.destroy()
                meth.destroy()

            for ha, me in zip(range(1, len(self.sh_container)), range(1, len(self.sh_method_container))):
                self.sh_container.pop(ha - ha - 1)
                self.sh_method_container.pop(me - me - 1)

            self.sh_plus = tk.Button(self.results, text="+", command=self.plussurface, font=("aerial", 11, "bold"), fg="black")
            self.sh_plus.grid(row=1, column=5, padx=10, ipadx=4, pady=2, sticky="W")
            self.sh_test = False

        # Core
        for entry2, method2, in zip(self.ch_container, self.ch_method_container):
            entry2.delete(0, 'end')
            method2.delete(0, 'end')

        if self.ch_test == True:
            self.ch_iterator.pop()
            for hard, meth in zip(self.ch_container[1:], self.ch_method_container[1:]):
                hard.destroy()
                meth.destroy()

            for har, met in zip(range(1, len(self.ch_container)), range(1, len(self.ch_method_container))):
                self.ch_container.pop(har - har - 1)
                self.ch_method_container.pop(met - met - 1)

            self.ch_plus = tk.Button(self.results, text="+", command=self.pluscore, font=("aerial", 11, "bold"), fg="black")
            self.ch_plus.grid(row=2, column=5, padx=10, ipadx=4, pady=2, sticky="W")
            self.ch_test = False

        self.savebutton['state'] = tk.NORMAL
        self.print_button['state'] = tk.DISABLED
        print(self.sh_iterator)
        return self.reportstatus.set(int(0)), self.sh_iterator, self.sh_container, self.sh_method_container, \
               self.ch_iterator, self.ch_container, self.ch_method_container, self.sh_test, self.ch_test

    def get_data(self):
        # Bodycote ID test
        if len(self.dispathnote_entry.get()) == 0:
            messagebox.showinfo('Warning', 'Field: Bodycote No. is not filled in')
            return

        # Updating databases
        self.updating_database()
        # Creating path
        if os.path.exists(self.localaddress + self.customer_entry.get().capitalize()) == False:
            os.mkdir(self.localaddress + self.customer_entry.get().capitalize())
        # Creating of excel file
        self.workbook = xlsxwriter.Workbook(str(self.localaddress) + self.year + "\\" + self.customer_entry.get().capitalize() + "\\" +
                                            (self.dispathnote_entry.get()) + "-Report.xlsx")
        self.worksheet_report = self.workbook.add_worksheet("Report")
        self.filetitle = self.dispathnote_entry.get() + "-Report.xlsx"
        # Setting the sizes of cells
        self.cellposition = ("A:A", "B:B", "C:C", "D:D", "E:E", "F:F", "G:G", "H:H", "I:I", "J:J", "K:K", "L:L", "M:M", "N:N", "O:O", "P:P", "Q:Q")
        self.colwidth = (17.29, 3.29, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71)

        for (x, y) in zip(self.cellposition, self.colwidth):
            self.worksheet_report.set_column(x, y)

        self.rowheight = (
            56.25, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 28.5, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15,
            15, 15, 15, 15, 15, 15, 8.25, 15, 15, 15, 8.25, 15, 15, 15)

        for (r, h) in zip(range(46), self.rowheight):
            self.worksheet_report.set_row(r, h)

        self.worksheet_report.set_paper(9)
        self.worksheet_report.fit_to_pages(1, 1)
        self.worksheet_report.set_margins(left=0.1, right=0.1, top=0.1, bottom=0.1)
        self.worksheet_report.print_area(0, 0, 46, 16)
        self.worksheet_report.set_print_scale(100)
        self.worksheet_report.center_horizontally()
        self.worksheet_report.center_vertically()
        # Formating
        self.title_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.base_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format2 = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'left', 'valign': 'bottom', 'text_wrap': 1, 'border': 1})
        self.base_format3 = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.base_format_TL = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TL.set_top()
        self.base_format_TL.set_left()
        self.base_format_L = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_L.set_left()
        self.base_format_BL = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_BL.set_bottom()
        self.base_format_BL.set_left()
        self.base_format_TLB = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TLB.set_top()
        self.base_format_TLB.set_left()
        self.base_format_TLB.set_bottom()
        self.base_format_TRB = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TRB.set_top()
        self.base_format_TRB.set_right()
        self.base_format_TRB.set_bottom()

        # Value formats
        self.value_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.value_format2 = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.value_format_TLR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_TLR.set_top()
        self.value_format_TLR.set_left()
        self.value_format_TLR.set_right()
        self.value_format_LR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_LR.set_left()
        self.value_format_LR.set_right()
        self.value_format_BLR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_BLR.set_bottom()
        self.value_format_BLR.set_left()
        self.value_format_BLR.set_right()
        self.ok_nok_format = self.workbook.add_format({'font_name': 'Calibri', 'bold': 1, 'font_size': 32, 'align': 'center', 'valign': 'vcenter',
                                                       'border': 1})
        # Images
        self.worksheet_report.insert_image('A1:I1', 'Bodycote_logo.png', {'x_offset': 4, 'y_offset': 4})
        # Formating cells
        # Blank formating
        # self.worksheet_report.conditional_format("B45:K45", {, 'format': self.border_format})
        # No Blanks
        self.worksheet_report.merge_range("A1:I1", "", self.title_format)
        self.worksheet_report.merge_range("J1:Q1", "Quenching / Annealing REPORT", self.title_format)
        #Header
        self.worksheet_report.merge_range("A3:B3", 'Customer:', self.base_format_TL)
        self.worksheet_report.merge_range("A4:B4", 'Order no.:', self.base_format_L)
        self.worksheet_report.merge_range("A5:B5", 'Bodycote no.:', self.base_format_L)
        self.worksheet_report.merge_range("A6:B6", 'Other description:', self.base_format_L)
        self.worksheet_report.merge_range("A7:B7", 'Part name:', self.base_format_L)
        self.worksheet_report.merge_range("A8:B8", 'Material:', self.base_format_L)
        self.worksheet_report.merge_range("A9:B9", 'Quantity:', self.base_format_L)
        self.worksheet_report.merge_range("A10:B11", 'Requirements:', self.base_format_BL)
        self.worksheet_report.write("A15", 'Surface hardness:', self.base_format)
        self.worksheet_report.write("A17", 'Core hardness:', self.base_format)
        self.worksheet_report.write("A37", 'Notes:', self.base_format)
        self.worksheet_report.merge_range("B41:M43", "The parts are according to the customer's requirements:", self.base_format2)
        self.worksheet_report.merge_range("A45:A47", 'Date: ', self.base_format3)
        self.worksheet_report.merge_range("G45:K47", 'Controller: ', self.base_format_TRB)
        # Writing datas
        self.worksheet_report.merge_range("C3:Q3", self.customer_entry.get(), self.value_format_TLR)
        self.worksheet_report.merge_range("C4:Q4", self.order_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C5:Q5", self.dispathnote_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C6:Q6", self.other_desc_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C7:Q7", self.partname_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C8:Q8", self.material_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C9:Q9", self.quantity_entry.get()+" pc/pcs", self.value_format_LR)
        self.worksheet_report.merge_range("C10:Q11", self.requirements_entry.get(1.0, "end-1c"), self.value_format_BLR)
        self.worksheet_report.merge_range("A13:Q13", "RESULTS:", self.title_format)
        self.worksheet_report.merge_range("B37:P38", self.notes_entry.get(1.0, "end-1c"), self.value_format)
        self.worksheet_report.merge_range("B45:F47", self.reportdate, self.base_format_TLB)
        self.worksheet_report.merge_range("L45:Q47", self.controller_entry.get(), self.value_format2)
        self.value_mover = ['B', 'F', 'J', 'N']
        self.method_mover = ['G', 'I', 'O', 'Q']
        for entry, method, position in zip(self.sh_container, self.sh_method_container, self.sh_iterator):
            self.worksheet_report.merge_range(self.value_mover[0+position]+"15:"+self.value_mover[1+position]+"15", entry.get(),
                                              self.value_format)
            self.worksheet_report.merge_range(self.method_mover[0+position]+"15:"+self.method_mover[1+position]+"15", method.get(),
                                              self.base_format)

        for entry, method, position in zip(self.ch_container, self.ch_method_container, self.ch_iterator):
            self.worksheet_report.merge_range(self.value_mover[0+position]+"17:"+self.value_mover[1+position]+"17", entry.get(),
                                              self.value_format)
            self.worksheet_report.merge_range(self.method_mover[0+position]+"17:"+self.method_mover[1+position]+"17", method.get(),
                                              self.base_format)

        if self.reportstatus.get() == 1:
            self.worksheet_report.merge_range("N41:P43", "OK", self.ok_nok_format)

        elif self.reportstatus.get() == 2:
            self.worksheet_report.merge_range("N41:P43", "NOK", self.ok_nok_format)

        if self.add_one_entry.get():
            self.worksheet_report.write("A20", 'Ret. austenite:', self.base_format)
            self.worksheet_report.merge_range('B20:C20', self.add_one_entry.get(), self.value_format)
            self.worksheet_report.merge_range('D20:E20', self.add_one_methods.get(), self.base_format)

        self.workbook.close()
        self.savebutton['state'] = tk.DISABLED
        self.print_button['state'] = tk.NORMAL


class Carburizing(Quenching):
    def __init__(self, title, image, height, add_one_p, add_two_p):
        super().__init__(title, image, height, add_one_p, add_two_p)
        # Layer
        self.layer_name = "CHD"
        self.layer = tk.Frame(self.qcore, width=994, height=140, bg="gray67")
        self.layer.grid(row=2, sticky="NW", padx=2)
        self.layer.grid_propagate(False)
        self.layer_label = tk.Label(self.layer, text="Results:", anchor="e", font=("aerial", 16, "bold"), fg="black", bg="gray67")
        self.layer_label.grid(row=0, column=0, ipady=4)
        self.chd_iterator = 1
        self.chd_depth_container = []
        self.chd_container = []
        self.layer_depth = tk.Label(self.layer, text="CHD depth:", bg="gray67")
        self.layer_depth.grid(row=1, column=0, padx=5, ipadx=17)
        self.layer_depth_entry = tk.Entry(self.layer, width=5)
        self.layer_depth_entry.grid(row=1, column=self.chd_iterator, padx=6, pady=4, ipady=2, sticky="W")
        self.chd_depth_container.append(self.layer_depth_entry)
        self.layer_hardness = tk.Label(self.layer, text="CHD hardness:", anchor="e", bg="gray67")
        self.layer_hardness.grid(row=2, column=0, padx=5)
        self.layer_hardness_entry = tk.Entry(self.layer, width=5)
        self.layer_hardness_entry.grid(row=2, column=self.chd_iterator, padx=6, pady=4, ipady=2, sticky="W")
        self.chd_container.append(self.layer_hardness_entry)
        self.layer_plus = tk.Button(self.layer, text="+", command=self.pluslayer, font=("aerial", 11, "bold"), fg="black")
        self.layer_plus.grid(row=1, column=self.chd_iterator+1, padx=6, ipadx=4, pady=2, sticky="W")
        self.layerhardness_methods = AutocompleteCombobox(self.layer, width=12, completevalues=self.methods)
        self.layerhardness_methods.grid(row=2, column=self.chd_iterator+1, ipady=2, pady=4, sticky="W")
        self.ultimate_hardness = tk.Label(self.layer, text="UHD:", anchor="e", bg="gray67")
        self.ultimate_hardness.grid(row=3, column=0, padx=5)
        self.ultimate_hardness_entry = tk.Entry(self.layer, width=13)
        self.ultimate_hardness_entry.grid(row=3, column=1, padx=6, pady=4, ipady=2, sticky="W", columnspan=2)
        #Results
        self.results_label.grid_forget()
        # Logical test
        self.test = False

    def pluslayer(self):
        self.chd_iterator += 1
        print(self.chd_iterator)
        if self.chd_iterator <= 14:
            # Depth
            self.layer_depth_entry = tk.Entry(self.layer, width=5)
            self.layer_depth_entry.grid(row=1, column=self.chd_iterator, padx=6, pady=4, ipady=2, sticky="W")
            self.chd_depth_container.append(self.layer_depth_entry)
            # Hardness
            self.layer_hardness_entry = tk.Entry(self.layer, width=5)
            self.layer_hardness_entry.grid(row=2, column=self.chd_iterator, padx=6, pady=4, ipady=2, sticky="W")
            self.chd_container.append(self.layer_hardness_entry)
            # Plus button
            self.layer_plus.destroy()
            self.layer_plus = tk.Button(self.layer, text="+", command=self.pluslayer, font=("aerial", 11, "bold"), fg="black")
            self.layer_plus.grid(row=1, column=self.chd_iterator+1, padx=6, ipadx=4, pady=2, sticky="W")
            # Method
            self.layerhardness_methods.destroy()
            self.layerhardness_methods = AutocompleteCombobox(self.layer, width=12, completevalues=self.methods)
            self.layerhardness_methods.grid(row=2, column=self.chd_iterator+1, ipady=2, sticky="W")
            return self.chd_iterator
        else:
            self.layer_plus.destroy()

    def run(self):
        self.qcore.mainloop()

    def calculate_layer(self, layer_name):
        # Processing
        self.point_hardness = []
        self.depth = []
        for chd_h in self.chd_container:
            if len(chd_h.get()) == 0:
                messagebox.showinfo('Warning', 'Some fields are not filled in')
                return
            else:
                self.point_hardness.append(int(chd_h.get()))

        for chd_d in self.chd_depth_container:
            if len(chd_d.get()) == 0:
                messagebox.showinfo('Warning', 'Some fields are not filled in')
                return
            else:
                self.depth.append(float(chd_d.get().replace(",", ".")))

        print(self.point_hardness)
        print(self.depth)
        self.UHD = []
        if len(self.ultimate_hardness_entry.get()) == 0:
            messagebox.showinfo('Warning', 'UHD missing !')
            return
        else:
            self.UHD = [int(self.ultimate_hardness_entry.get())] * (len(self.point_hardness))

        print(self.UHD)
        self.min_eq_value = 0
        self.min_index = 0
        self.max_eq_value = 0
        self.max_index = 0

        if int(self.ultimate_hardness_entry.get()) > self.point_hardness[0] or int(self.ultimate_hardness_entry.get()) < self.point_hardness[-1]:
            self.min_eq_value = 0
            self.max_eq_value = 0
        else:
            for i in self.point_hardness:
                if i <= int(self.ultimate_hardness_entry.get()) and i != 0:
                    self.min_index = self.point_hardness.index(i)
                    self.min_eq_value = int(self.point_hardness[self.min_index])
                    print(self.min_eq_value)
                    break

        for a in self.point_hardness[::-1]:
            if a >= int(self.ultimate_hardness_entry.get()):
                self.max_index = self.point_hardness.index(a)
                self.max_eq_value = int(self.point_hardness[self.max_index])
                print(self.max_eq_value)
                break

        self.ld_result = []
        self.layer_diff = None

        if self.max_eq_value == 0 or self.min_eq_value == 0:
            plt.plot(self.depth, self.point_hardness)
            plt.plot(self.depth, self.UHD)
            self.test = True
            plt.savefig(str(self.imgs_address) + self.dispathnote_entry.get() + '.png')
        elif self.min_eq_value == self.max_eq_value:
            self.ld_result = round(float(self.depth[self.min_index]), 2)
            plt.plot(self.depth, self.point_hardness)
            plt.plot(self.depth, self.UHD)
            plt.plot(self.ld_result, self.UHD[0], 'ro')
            plt.annotate(f'{self.layer_name}: {self.ld_result} mm / {self.UHD[0]} {self.layerhardness_methods.get()}',
                         xy=(self.ld_result, self.UHD[0]), xytext=(self.depth[0] + 0.14, self.UHD[0] - 40),
                         horizontalalignment='center', fontsize=12, bbox=dict(boxstyle="round", fc="0.8"))
            self.test = True
            plt.savefig(str(self.imgs_address) + self.dispathnote_entry.get() + '.png')
        else:
            self.layer_diff = ((self.max_eq_value - self.UHD[0]) * (self.depth[self.min_index] - self.depth[self.max_index])) / (
                        self.max_eq_value - self.min_eq_value)
            self.ld_result = round(self.depth[self.max_index] + self.layer_diff, 2)
            plt.plot(self.depth, self.point_hardness)
            plt.plot(self.depth, self.UHD)
            plt.plot(self.ld_result, self.UHD[0], 'ro')
            plt.annotate(f'{self.layer_name}: {self.ld_result} mm / {self.UHD[0]} {self.layerhardness_methods.get()}',
                         xy=(self.ld_result, self.UHD[0]), xytext=(self.depth[len(self.point_hardness) // 2], self.UHD[0] - 40),
                         horizontalalignment='center', fontsize=12, bbox=dict(boxstyle="round", fc="0.8"))
            self.test = True
            plt.savefig(str(self.imgs_address) + self.dispathnote_entry.get() + '.png')

    def range_char(self, first, last):
        return (chr(n) for n in range(ord(first), ord(last) + 1))

    def clean_cells(self):
        self.customer_entry.delete(0, 'end')
        self.order_entry.delete(0, 'end')
        self.dispathnote_entry.delete(0, 'end')
        self.other_desc_entry.delete(0, 'end')
        self.partname_entry.delete(0, 'end')
        self.material_entry.delete(0, 'end')
        self.quantity_entry.delete(0, 'end')
        self.requirements_entry.delete('1.0', 'end')
        self.notes_entry.delete('1.0', 'end')
        self.ultimate_hardness_entry.delete(0, 'end')
        self.layerhardness_methods.delete(0, 'end')
        self.add_one_entry.delete(0, 'end')
        self.add_one_methods.delete(0, 'end')
        self.add_two_entry.delete(0, 'end')
        self.add_two_methods.delete(0, 'end')
        self.controller_entry.delete(0, 'end')

        # Hardness remove
        # Surface
        for entry1, method1, in zip(self.sh_container, self.sh_method_container):
            entry1.delete(0, 'end')
            method1.delete(0, 'end')

        if self.sh_test == True:
            self.sh_iterator.pop()
            for hard, meth in zip(self.sh_container[1:], self.sh_method_container[1:]):
                hard.destroy()
                meth.destroy()

            for ha, me in zip(range(1, len(self.sh_container)), range(1, len(self.sh_method_container))):
                self.sh_container.pop(ha - ha - 1)
                self.sh_method_container.pop(me - me - 1)

            self.sh_plus = tk.Button(self.results, text="+", command=self.plussurface, font=("aerial", 11, "bold"), fg="black")
            self.sh_plus.grid(row=1, column=5, padx=10, ipadx=4, pady=2, sticky="W")
            self.sh_test = False

        # Core
        for entry2, method2, in zip(self.ch_container, self.ch_method_container):
            entry2.delete(0, 'end')
            method2.delete(0, 'end')

        if self.ch_test == True:
            self.ch_iterator.pop()
            for hard, meth in zip(self.ch_container[1:], self.ch_method_container[1:]):
                hard.destroy()
                meth.destroy()

            for har, met in zip(range(1, len(self.ch_container)), range(1, len(self.ch_method_container))):
                self.ch_container.pop(har - har - 1)
                self.ch_method_container.pop(met - met - 1)

            self.ch_plus = tk.Button(self.results, text="+", command=self.pluscore, font=("aerial", 11, "bold"), fg="black")
            self.ch_plus.grid(row=2, column=5, padx=10, ipadx=4, pady=2, sticky="W")
            self.ch_test = False

        # Layer remove
        for chd_h, chd_d in zip(self.chd_container[1:], self.chd_depth_container[1:]):
            chd_h.destroy()
            chd_d.destroy()

        for h, d in zip(range(1, len(self.chd_container)), range(1, len(self.chd_depth_container))):
            self.chd_container.pop(h-h-1)
            self.chd_depth_container.pop(d-d-1)

        self.layer_plus.destroy()
        self.layerhardness_methods.destroy()
        self.chd_iterator = 1
        self.layer_plus = tk.Button(self.layer, text="+", command=self.pluslayer, font=("aerial", 11, "bold"), fg="black")
        self.layer_plus.grid(row=1, column=self.chd_iterator + 1, padx=6, ipadx=4, pady=2, sticky="W")
        self.layerhardness_methods = AutocompleteCombobox(self.layer, width=12, completevalues=self.methods)
        self.layerhardness_methods.grid(row=2, column=self.chd_iterator + 1, ipady=2, pady=4, sticky="W")
        plt.cla()

        for chd_h, chd_d in zip(self.chd_container, self.chd_depth_container):
            chd_h.delete(0, 'end')
            chd_d.delete(0, 'end')

        self.savebutton['state'] = tk.NORMAL
        self.print_button['state'] = tk.DISABLED
        return self.reportstatus.set(int(0)), self.sh_iterator, self.sh_container, self.sh_method_container, \
               self.ch_iterator, self.ch_container, self.ch_method_container, self.chd_iterator, self.chd_container, self.chd_depth_container, \
               self.sh_test, self.ch_test

    def get_data(self):
        # Bodycote ID test
        if len(self.dispathnote_entry.get()) == 0:
            messagebox.showinfo('Warning', 'Field: Bodycote No. is not filled in')
            return

        # Updating databases
        self.updating_database()
        # Creating path
        if os.path.exists(self.localaddress + self.customer_entry.get().capitalize()) == False:
            os.mkdir(self.localaddress + self.customer_entry.get().capitalize())
        # Calculating of CHD layer
        self.calculate_layer(self.layer_name)
        # Creating of excel file
        self.workbook = xlsxwriter.Workbook(str(self.localaddress) + self.customer_entry.get().capitalize() + "\\" +
                                            (self.dispathnote_entry.get()) + "-Report.xlsx")
        self.worksheet_report = self.workbook.add_worksheet("Report")
        self.filetitle = self.dispathnote_entry.get() + "-Report.xlsx"
        # Setting the sizes of cells
        self.cellposition = ("A:A", "B:B", "C:C", "D:D", "E:E", "F:F", "G:G", "H:H", "I:I", "J:J", "K:K", "L:L", "M:M", "N:N", "O:O", "P:P", "Q:Q")
        self.colwidth = (17.29, 3.29, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71)

        for (x, y) in zip(self.cellposition, self.colwidth):
            self.worksheet_report.set_column(x, y)

        self.rowheight = (
            56.25, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 28.5, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15,
            15, 15, 15, 15, 15, 15, 8.25, 15, 15, 15, 8.25, 15, 15, 15)

        for (r, h) in zip(range(46), self.rowheight):
            self.worksheet_report.set_row(r, h)

        self.worksheet_report.set_paper(9)
        self.worksheet_report.fit_to_pages(1, 1)
        self.worksheet_report.set_margins(left=0.1, right=0.1, top=0.1, bottom=0.1)
        self.worksheet_report.print_area(0, 0, 46, 16)
        self.worksheet_report.set_print_scale(100)
        self.worksheet_report.center_horizontally()
        self.worksheet_report.center_vertically()
        # Formating
        self.title_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.base_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format2 = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'left', 'valign': 'bottom', 'text_wrap': 1, 'border': 1})
        self.base_format3 = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.base_format_TL = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TL.set_top()
        self.base_format_TL.set_left()
        self.base_format_L = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_L.set_left()
        self.base_format_BL = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_BL.set_bottom()
        self.base_format_BL.set_left()
        self.base_format_TLB = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TLB.set_top()
        self.base_format_TLB.set_left()
        self.base_format_TLB.set_bottom()
        self.base_format_TRB = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TRB.set_top()
        self.base_format_TRB.set_right()
        self.base_format_TRB.set_bottom()

        # Value formats
        self.value_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.value_format2 = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.value_format_TLR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_TLR.set_top()
        self.value_format_TLR.set_left()
        self.value_format_TLR.set_right()
        self.value_format_LR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_LR.set_left()
        self.value_format_LR.set_right()
        self.value_format_BLR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_BLR.set_bottom()
        self.value_format_BLR.set_left()
        self.value_format_BLR.set_right()
        self.ok_nok_format = self.workbook.add_format({'font_name': 'Calibri', 'bold': 1, 'font_size': 32, 'align': 'center', 'valign': 'vcenter',
                                                       'border': 1})
        # Images
        self.worksheet_report.insert_image('A1:I1', 'Bodycote_logo.png', {'x_offset': 4, 'y_offset': 4})

        if self.test:
            self.worksheet_report.insert_image('F23:Q36', str(self.imgs_address) + self.dispathnote_entry.get() + '.png',
                                               {'x_scale': 0.6, 'y_scale': 0.6, 'x_offset': 4})
        else:
            return

        # Formatting cells
        # No Blanks
        self.worksheet_report.merge_range("A1:I1", "", self.title_format)
        self.worksheet_report.merge_range("J1:Q1", "Carburising / Carbonitriding", self.title_format)
        # Header
        self.worksheet_report.merge_range("A3:B3", 'Customer:', self.base_format_TL)
        self.worksheet_report.merge_range("A4:B4", 'Order no.:', self.base_format_L)
        self.worksheet_report.merge_range("A5:B5", 'Bodycote no.:', self.base_format_L)
        self.worksheet_report.merge_range("A6:B6", 'Other description:', self.base_format_L)
        self.worksheet_report.merge_range("A7:B7", 'Part name:', self.base_format_L)
        self.worksheet_report.merge_range("A8:B8", 'Material:', self.base_format_L)
        self.worksheet_report.merge_range("A9:B9", 'Quantity:', self.base_format_L)
        self.worksheet_report.merge_range("A10:B11", 'Requirements:', self.base_format_BL)
        self.worksheet_report.write("A15", 'Surface hardness:', self.base_format)
        self.worksheet_report.write("A17", 'Core hardness:', self.base_format)
        self.worksheet_report.write("A19", 'CHD depth:', self.base_format)
        self.worksheet_report.write("A20", 'CHD hardness:', self.base_format)
        self.worksheet_report.write("A21", 'UHD:', self.base_format)
        self.worksheet_report.write("A37", 'Notes:', self.base_format)
        self.worksheet_report.merge_range("B41:M43", "The parts are according to the customer's requirements:", self.base_format2)
        self.worksheet_report.merge_range("A45:A47", 'Date: ', self.base_format3)
        self.worksheet_report.merge_range("G45:K47", 'Controller: ', self.base_format_TRB)
        # Writing datas
        self.worksheet_report.merge_range("C3:Q3", self.customer_entry.get(), self.value_format_TLR)
        self.worksheet_report.merge_range("C4:Q4", self.order_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C5:Q5", self.dispathnote_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C6:Q6", self.other_desc_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C7:Q7", self.partname_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C8:Q8", self.material_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C9:Q9", self.quantity_entry.get()+" pc/pcs", self.value_format_LR)
        self.worksheet_report.merge_range("C10:Q11", self.requirements_entry.get(1.0, "end-1c"), self.value_format_BLR)
        self.worksheet_report.merge_range("A13:Q13", "RESULTS:", self.title_format)
        self.worksheet_report.merge_range("B37:P38", self.notes_entry.get(1.0, "end-1c"), self.value_format)
        self.worksheet_report.merge_range("B45:F47", self.reportdate, self.base_format_TLB)
        self.worksheet_report.merge_range("L45:Q47", self.controller_entry.get(), self.value_format2)
        self.worksheet_report.merge_range("B21:C21", self.ultimate_hardness_entry.get(), self.value_format)
        self.worksheet_report.merge_range("D21:E21", self.layerhardness_methods.get(), self.base_format)
        self.value_mover = ['B', 'F', 'J', 'N']
        self.method_mover = ['G', 'I', 'O', 'Q']
        for entry, method, position in zip(self.sh_container, self.sh_method_container, self.sh_iterator):
            self.worksheet_report.merge_range(self.value_mover[0+position]+"15:"+self.value_mover[1+position]+"15", entry.get(),
                                              self.value_format)
            self.worksheet_report.merge_range(self.method_mover[0+position]+"15:"+self.method_mover[1+position]+"15", method.get(),
                                              self.base_format)

        for entry, method, position in zip(self.ch_container, self.ch_method_container, self.ch_iterator):
            self.worksheet_report.merge_range(self.value_mover[0+position]+"17:"+self.value_mover[1+position]+"17", entry.get(),
                                              self.value_format)
            self.worksheet_report.merge_range(self.method_mover[0+position]+"17:"+self.method_mover[1+position]+"17", method.get(),
                                              self.base_format)

        for depth, hardness, letter in zip(self.chd_depth_container, self.chd_container, self.range_char("B", "P")):
            self.worksheet_report.write(letter + "19", depth.get(), self.value_format)
            self.worksheet_report.write(letter + "20", hardness.get(), self.value_format)

        if self.reportstatus.get() == 1:
            self.worksheet_report.merge_range("N41:P43", "OK", self.ok_nok_format)

        elif self.reportstatus.get() == 2:
            self.worksheet_report.merge_range("N41:P43", "NOK", self.ok_nok_format)

        if self.add_one_entry.get():
            self.worksheet_report.write("A23", 'Ret. austenite:', self.base_format)
            self.worksheet_report.merge_range('B23:C23', self.add_one_entry.get(), self.value_format)
            self.worksheet_report.merge_range('D23:E23', self.add_one_methods.get(), self.base_format)

        if self.add_two_entry.get():
            self.worksheet_report.write("A25", 'Int. oxidation:', self.base_format)
            self.worksheet_report.merge_range('B25:C25', self.add_two_entry.get(), self.value_format)
            self.worksheet_report.merge_range('D25:E25', self.add_two_methods.get(), self.base_format)

        self.workbook.close()
        self.savebutton['state'] = tk.DISABLED
        self.print_button['state'] = tk.NORMAL


class Nitriding(Carburizing):
    def __init__(self, title, image, height, add_one_p, add_two_p):
        super().__init__(title, image, height, add_one_p, add_two_p)
        # Layer
        self.layer_name = "NHT"
        self.layer = tk.Frame(self.qcore, width=994, height=140, bg="gray67")
        self.layer.grid(row=2, sticky="NW", padx=4, pady=0)
        self.layer.grid_propagate(False)
        self.layer_label = tk.Label(self.layer, text="Results:", anchor="e", font=("aerial", 16, "bold"), fg="black", bg="gray67")
        self.layer_label.grid(row=0, column=0, ipady=4)
        self.chd_iterator = 1
        self.chd_depth_container = []
        self.chd_container = []
        self.layer_depth = tk.Label(self.layer, text="NHT depth:", bg="gray67")
        self.layer_depth.grid(row=1, column=0, padx=5, ipadx=17)
        self.layer_depth_entry = tk.Entry(self.layer, width=5)
        self.layer_depth_entry.grid(row=1, column=self.chd_iterator, padx=6, pady=4, ipady=2, sticky="W")
        self.chd_depth_container.append(self.layer_depth_entry)
        self.layer_hardness = tk.Label(self.layer, text="NHT hardness:", anchor="e", bg="gray67")
        self.layer_hardness.grid(row=2, column=0, padx=5)
        self.layer_hardness_entry = tk.Entry(self.layer, width=5)
        self.layer_hardness_entry.grid(row=2, column=self.chd_iterator, padx=6, pady=4, ipady=2, sticky="W")
        self.chd_container.append(self.layer_hardness_entry)
        self.layer_plus = tk.Button(self.layer, text="+", command=self.pluslayer, font=("aerial", 11, "bold"), fg="black")
        self.layer_plus.grid(row=1, column=self.chd_iterator + 1, padx=6, ipadx=4, pady=2, sticky="W")
        self.layerhardness_methods = AutocompleteCombobox(self.layer, width=12, completevalues=self.methods)
        self.layerhardness_methods.grid(row=2, column=self.chd_iterator + 1, ipady=2, pady=4, sticky="W")
        self.ultimate_hardness = tk.Label(self.layer, text="UHD:", anchor="e", bg="gray67")
        self.ultimate_hardness.grid(row=3, column=0, padx=5)
        self.ultimate_hardness_entry = tk.Entry(self.layer, width=13)
        self.ultimate_hardness_entry.grid(row=3, column=1, padx=6, pady=4, ipady=2, sticky="W", columnspan=2)
        self.test = False

    def get_data(self):
        # Bodycote ID test
        if len(self.dispathnote_entry.get()) == 0:
            messagebox.showinfo('Warning', 'Field: Bodycote No. is not filled in')
            return

        # Updating databases
        self.updating_database()
        # Creating path
        if os.path.exists(self.localaddress + self.customer_entry.get().capitalize()) == False:
            os.mkdir(self.localaddress + self.customer_entry.get().capitalize())
        # Calculating of NHT layer
        self.calculate_layer(self.layer_name)
        # Creating of excel file
        self.workbook = xlsxwriter.Workbook(str(self.localaddress) + self.customer_entry.get().capitalize() + "\\" +
                                            (self.dispathnote_entry.get()) + "-Report.xlsx")
        self.worksheet_report = self.workbook.add_worksheet("Report")
        self.filetitle = self.dispathnote_entry.get() + "-Report.xlsx"
        # Setting the sizes of cells
        self.cellposition = ("A:A", "B:B", "C:C", "D:D", "E:E", "F:F", "G:G", "H:H", "I:I", "J:J", "K:K", "L:L", "M:M", "N:N", "O:O", "P:P", "Q:Q")
        self.colwidth = (17.29, 3.29, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71, 3.71)

        for (x, y) in zip(self.cellposition, self.colwidth):
            self.worksheet_report.set_column(x, y)

        self.rowheight = (
            56.25, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 28.5, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15,
            15, 15, 15, 15, 15, 15, 8.25, 15, 15, 15, 8.25, 15, 15, 15)

        for (r, h) in zip(range(46), self.rowheight):
            self.worksheet_report.set_row(r, h)

        self.worksheet_report.set_paper(9)
        self.worksheet_report.fit_to_pages(1, 1)
        self.worksheet_report.set_margins(left=0.1, right=0.1, top=0.1, bottom=0.1)
        self.worksheet_report.print_area(0, 0, 46, 16)
        self.worksheet_report.set_print_scale(100)
        self.worksheet_report.center_horizontally()
        self.worksheet_report.center_vertically()
        # Formating
        self.title_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.base_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format2 = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'left', 'valign': 'bottom', 'text_wrap': 1, 'border': 1})
        self.base_format3 = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.base_format_TL = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TL.set_top()
        self.base_format_TL.set_left()
        self.base_format_L = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_L.set_left()
        self.base_format_BL = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_BL.set_bottom()
        self.base_format_BL.set_left()
        self.base_format_TLB = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TLB.set_top()
        self.base_format_TLB.set_left()
        self.base_format_TLB.set_bottom()
        self.base_format_TRB = self.workbook.add_format(
            {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1})
        self.base_format_TRB.set_top()
        self.base_format_TRB.set_right()
        self.base_format_TRB.set_bottom()

        # Value formats
        self.value_format = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.value_format2 = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
        self.value_format_TLR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_TLR.set_top()
        self.value_format_TLR.set_left()
        self.value_format_TLR.set_right()
        self.value_format_LR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_LR.set_left()
        self.value_format_LR.set_right()
        self.value_format_BLR = self.workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1})
        self.value_format_BLR.set_bottom()
        self.value_format_BLR.set_left()
        self.value_format_BLR.set_right()
        self.ok_nok_format = self.workbook.add_format({'font_name': 'Calibri', 'bold': 1, 'font_size': 32, 'align': 'center', 'valign': 'vcenter',
                                                       'border': 1})
        # Images
        self.worksheet_report.insert_image('A1:I1', 'Bodycote_logo.png', {'x_offset': 4, 'y_offset': 4})

        if self.test:
            self.worksheet_report.insert_image('F23:Q36', str(self.imgs_address) + self.dispathnote_entry.get() + '.png',
                                               {'x_scale': 0.6, 'y_scale': 0.6, 'x_offset': 4})
        else:
            return

        # Formatting cells
        # No Blanks
        self.worksheet_report.merge_range("A1:I1", "", self.title_format)
        self.worksheet_report.merge_range("J1:Q1", "Nitriding / Nitrocarburising", self.title_format)
        # Header
        self.worksheet_report.merge_range("A3:B3", 'Customer:', self.base_format_TL)
        self.worksheet_report.merge_range("A4:B4", 'Order no.:', self.base_format_L)
        self.worksheet_report.merge_range("A5:B5", 'Bodycote no.:', self.base_format_L)
        self.worksheet_report.merge_range("A6:B6", 'Other description:', self.base_format_L)
        self.worksheet_report.merge_range("A7:B7", 'Part name:', self.base_format_L)
        self.worksheet_report.merge_range("A8:B8", 'Material:', self.base_format_L)
        self.worksheet_report.merge_range("A9:B9", 'Quantity:', self.base_format_L)
        self.worksheet_report.merge_range("A10:B11", 'Requirements:', self.base_format_BL)
        self.worksheet_report.write("A15", 'Surface hardness:', self.base_format)
        self.worksheet_report.write("A17", 'Core hardness:', self.base_format)
        self.worksheet_report.write("A19", 'NHT depth:', self.base_format)
        self.worksheet_report.write("A20", 'NHT hardness:', self.base_format)
        self.worksheet_report.write("A21", 'UHD:', self.base_format)
        self.worksheet_report.write("A37", 'Notes:', self.base_format)
        self.worksheet_report.merge_range("B41:M43", "The parts are according to the customer's requirements:", self.base_format2)
        self.worksheet_report.merge_range("A45:A47", 'Date: ', self.base_format3)
        self.worksheet_report.merge_range("G45:K47", 'Controller: ', self.base_format_TRB)
        # Writing datas
        self.worksheet_report.merge_range("C3:Q3", self.customer_entry.get(), self.value_format_TLR)
        self.worksheet_report.merge_range("C4:Q4", self.order_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C5:Q5", self.dispathnote_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C6:Q6", self.other_desc_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C7:Q7", self.partname_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C8:Q8", self.material_entry.get(), self.value_format_LR)
        self.worksheet_report.merge_range("C9:Q9", self.quantity_entry.get()+" pc/pcs", self.value_format_LR)
        self.worksheet_report.merge_range("C10:Q11", self.requirements_entry.get(1.0, "end-1c"), self.value_format_BLR)
        self.worksheet_report.merge_range("A13:Q13", "RESULTS:", self.title_format)
        self.worksheet_report.merge_range("B37:P38", self.notes_entry.get(1.0, "end-1c"), self.value_format)
        self.worksheet_report.merge_range("B45:F47", self.reportdate, self.base_format_TLB)
        self.worksheet_report.merge_range("L45:Q47", self.controller_entry.get(), self.value_format2)
        self.worksheet_report.merge_range("B21:C21", self.ultimate_hardness_entry.get(), self.value_format)
        self.worksheet_report.merge_range("D21:E21", self.layerhardness_methods.get(), self.base_format)
        self.value_mover = ['B', 'F', 'J', 'N']
        self.method_mover = ['G', 'I', 'O', 'Q']
        for entry, method, position in zip(self.sh_container, self.sh_method_container, self.sh_iterator):
            self.worksheet_report.merge_range(self.value_mover[0+position]+"15:"+self.value_mover[1+position]+"15", entry.get(),
                                              self.value_format)
            self.worksheet_report.merge_range(self.method_mover[0+position]+"15:"+self.method_mover[1+position]+"15", method.get(),
                                              self.base_format)

        for entry, method, position in zip(self.ch_container, self.ch_method_container, self.ch_iterator):
            self.worksheet_report.merge_range(self.value_mover[0+position]+"17:"+self.value_mover[1+position]+"17", entry.get(),
                                              self.value_format)
            self.worksheet_report.merge_range(self.method_mover[0+position]+"17:"+self.method_mover[1+position]+"17", method.get(),
                                              self.base_format)

        for depth, hardness, letter in zip(self.chd_depth_container, self.chd_container, self.range_char("B", "P")):
            self.worksheet_report.write(letter + "19", depth.get(), self.value_format)
            self.worksheet_report.write(letter + "20", hardness.get(), self.value_format)

        if self.reportstatus.get() == 1:
            self.worksheet_report.merge_range("N41:P43", "OK", self.ok_nok_format)

        elif self.reportstatus.get() == 2:
            self.worksheet_report.merge_range("N41:P43", "NOK", self.ok_nok_format)

        if self.add_one_entry.get():
            self.worksheet_report.write("A23", 'CLT:', self.base_format)
            self.worksheet_report.merge_range('B23:C23', self.add_one_entry.get(), self.value_format)
            self.worksheet_report.merge_range('D23:E23', self.add_one_methods.get(), self.base_format)

        if self.add_two_entry.get():
            self.worksheet_report.write("A25", 'Oxid. layer:', self.base_format)
            self.worksheet_report.merge_range('B25:C25', self.add_two_entry.get(), self.value_format)
            self.worksheet_report.merge_range('D25:E25', self.add_two_methods.get(), self.base_format)

        self.workbook.close()
        self.savebutton['state'] = tk.DISABLED
        self.print_button['state'] = tk.NORMAL

    def run(self):
        self.qcore.mainloop()


gui = AskTypeWindow()
gui.ask_type_run()
