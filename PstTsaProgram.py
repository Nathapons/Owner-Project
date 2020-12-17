from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msb
import xlrd
import openpyxl as xl
import csv
import os

class PsaTsaProgram:
    def __init__(self):
        self.raw_data_location = '\\\\Ta1d181222\\01.result test peel strength\\DATA IPQC'
        self.master_list_location = "\\\\Ta1d180613\\2.MASTER LIST RUNNING NUMBER\\2. Master list IPQC"
        self.result_peel_location = "\\\\Ta1d180613\\3.RESULT DATA PSA\IPQC (SMT)\\2. Result peel strength (ผลการตรวจสอบจาก ANC)"
        self.spc_master_location = "\\\\Ta1d180613\\3.RESULT DATA PSA\IPQC (SMT)\\99.SPC Master(ห้ามลบ)\\QF-A1-QUA-2036-6.xlsx"
        
        # Widget Properties
        self.topic_font = ('Arial', 24, 'bold')
        self.detail_font = ('Arial', 16)

    def main_program(self):
        self.spc_window = Tk()
        WIDTH = 550
        HEIGHT = 580
        screen_width = self.spc_window.winfo_screenwidth()
        screen_height = self.spc_window.winfo_screenheight()
        x = int( (screen_width/2) - (WIDTH/2) )
        y = int( (screen_height/2) - (HEIGHT/1.8))

        # GUI Properties
        self.spc_window.title('IPQC-PSATSA Program')
        self.spc_window.resizable(0, 0)
        self.spc_window.geometry(f'{WIDTH}x{HEIGHT}+{x}+{y}')
        self.spc_window.iconbitmap('fujikura_logo.ico')

        # WIDGET Properties
        big_frame = Frame(self.spc_window)
        topic_label = Label(big_frame, text='SPC Report Program', font=self.topic_font, fg='white', bg='blue')
        select_frame = Frame(big_frame)
        year_label = Label(select_frame, text='Year Report:', font=self.detail_font)
        self.year_cb = ttk.Combobox(select_frame, justify='center', font=self.detail_font, width=18)
        month_label = Label(select_frame, text='Month Report:', font=self.detail_font)
        self.month_cb = ttk.Combobox(select_frame, justify='center', font=self.detail_font, width=18)
        product_label = Label(select_frame, text='Product Report:', font=self.detail_font)
        self.product_cb = ttk.Combobox(select_frame, justify='center', font=self.detail_font, width=18)
        machine_label = Label(select_frame, text='Machine No:', font=self.detail_font)
        self.machine_cb = ttk.Combobox(select_frame, justify='center', font=self.detail_font, width=18)
        doc_req_label = Label(select_frame, text='Report Name:', font=self.detail_font)
        check_box_frame = Frame(select_frame)
        self.peel_strength_name = StringVar()
        flex_check_box = Radiobutton(check_box_frame, text='Flex Report', font=self.detail_font, variable=self.peel_strength_name, value='flex')
        liner_check_box = Radiobutton(check_box_frame, text='Liner Report', font=self.detail_font, variable=self.peel_strength_name, value='liner')
        run_program_button = Button(big_frame, text='Run Program', font=('Arial', 18, 'bold'), command=self.spc_program_overview)
        treeview_label = Label(big_frame, text='Click Here to Open Excel:', font=self.detail_font, fg='#00802b')

        # Treeview Widget
        headers = ['File Update', 'Report status']
        self.psa_tas_treeview = ttk.Treeview(big_frame, column=headers, show='headings',height=7)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 14, 'bold'), foreground="#bf00ff")
        style.configure("Treeview", font=('Arial', 10))
        for header in headers:
            self.psa_tas_treeview.heading(header, text=header)
            self.psa_tas_treeview.column(header, anchor='center', width=240, minwidth=0)

        # Widget Event
        self.year_cb.bind('<Button-1>', self.year_click)
        self.month_cb.bind('<Button-1>', self.month_click)
        self.product_cb.bind('<Button-1>', self.product_click)
        self.machine_cb.bind('<Button-1>', self.machine_click)
        flex_check_box.invoke()

        # Widget Position
        big_frame.pack()
        topic_label.grid(row=0, column=0, pady=10, ipadx=10)
        select_frame.grid(row=1, column=0, pady=5)
        year_label.grid(row=0, column=0, padx=5, pady=5)
        self.year_cb.grid(row=0, column=1, padx=5, pady=5, ipady=2)
        month_label.grid(row=1, column=0, padx=5, pady=5)
        self.month_cb.grid(row=1, column=1, padx=5, pady=5, ipady=2)
        product_label.grid(row=2, column=0, padx=5, pady=5)
        self.product_cb.grid(row=2, column=1, padx=5, pady=5, ipady=2)
        machine_label.grid(row=3, column=0, padx=5, pady=5)
        self.machine_cb.grid(row=3, column=1, padx=5, pady=5, ipady=2)
        doc_req_label.grid(row=4, column=0, padx=5, pady=5)
        check_box_frame.grid(row=4, column=1, padx=5, pady=5)
        flex_check_box.grid(row=0, column=0, padx=5)
        liner_check_box.grid(row=0, column=1, padx=5)
        run_program_button.grid(row=2, column=0, pady=2, ipadx=20)
        treeview_label.grid(row=3, column=0, pady=5)
        self.psa_tas_treeview.grid(row=4, column=0, pady=2)

        # GUI Activate
        self.spc_window.mainloop()

    def year_click(self, event):
        folder_name_list = []

        for folder_name in os.listdir(self.raw_data_location):
            if str(folder_name).startswith("20"):
                folder_name_list.append(folder_name)

        self.year_cb['values'] = folder_name_list
        self.month_cb.set("")
        self.product_cb.set("")
    
    def month_click(self, event):
        year_input = self.year_cb.get()

        if year_input != "":
            month_all_folders_location = os.path.join(self.raw_data_location, year_input)
            month_folders_list = os.listdir(month_all_folders_location)

            self.month_cb['values'] = month_folders_list
            self.product_cb.set("")
            self.machine_cb.set("")
        else:
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Year Report ก่อน')
            self.month_cb['values'] = []
    
    def product_click(self, event):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()

        if year_input != "" and month_input != "": 
            month_input_list = month_input.split("'")
            month_name = month_input_list[1][0:3]
            new_master_file_location = os.path.join(self.master_list_location, year_input)
            product_name_list = []

            for excel_file_name in os.listdir(new_master_file_location):
                if month_name in excel_file_name and not excel_file_name.startswith("~$") and excel_file_name.endswith(".xlsx"):
                    master_file_location = os.path.join(new_master_file_location, excel_file_name)
                    # Open master file
                    master_book = xlrd.open_workbook(filename=master_file_location)
                    master_sheet = master_book.sheet_by_index(0)
                    start_row = 8
                    product_name_cell = master_sheet.cell(rowx=start_row, colx=3).value

                    while product_name_cell != "":
                        product_name = str(product_name_cell)[0:10]

                        # Add data to list when product startwith RG and don't have add to list
                        if product_name.upper() not in product_name_list and product_name.upper().startswith('RG'):
                            product_name_list.append(product_name.upper())

                        start_row += 5
                        product_name_cell = master_sheet.cell(rowx=start_row, colx=3).value

                    # Close Master file
                    master_book.release_resources()
            
            self.product_cb['values'] = product_name_list
            self.machine_cb.set("")
        else:
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Year Report และ Month Report ก่อน')
            self.product_cb['values'] = []
            self.product_cb.set("")

    def machine_click(self, event):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()
        product_input = self.product_cb.get()

        if year_input != "" and month_input != "" and product_input != "": 
            month_input_list = month_input.split("'")
            month_name = month_input_list[1][0:3]
            new_master_file_location = os.path.join(self.master_list_location, year_input)
            machine_number_list = []
            
            for excel_file_name in os.listdir(new_master_file_location):
                if month_name in excel_file_name and not excel_file_name.startswith("~$") and excel_file_name.endswith(".xlsx"):
                    master_file_location = os.path.join(new_master_file_location, excel_file_name)
                    # Open master file
                    master_book = xlrd.open_workbook(filename=master_file_location)
                    master_sheet = master_book.sheet_by_index(0)
                    start_row = 8
                    product_name_cell = master_sheet.cell(rowx=start_row, colx=3).value
                    machine_number_cell = master_sheet.cell(rowx=start_row, colx=5).value

                    while product_name_cell != "":
                        product_name = str(product_name_cell).upper()
                        
                        if product_name.startswith(product_input) and machine_number_cell not in machine_number_list:
                            machine_number_list.append(str(machine_number_cell).strip())

                        # Loop command
                        start_row += 5
                        product_name_cell = master_sheet.cell(rowx=start_row, colx=3).value
                        machine_number_cell = master_sheet.cell(rowx=start_row, colx=5).value

                    # Close master file
                    master_book.release_resources()

            self.machine_cb['values'] = machine_number_list
        else:
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Year Report Month Report และ Product Report ก่อน')
            self.machine_cb['values'] = []
            self.machine_cb.set("")

    def spc_program_overview(self):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()
        product_input = self.product_cb.get()
        machine_input = self.machine_cb.get()

        if year_input == "" and month_input == "" and product_input == "" and machine_input == "":
            msb.showinfo(title='แจ้งเตือนไปยังผู้ใช้', message='คุณกรอกข้อมูลข้างบนไม่ครบถ้วน')
        else:
            spc_file_location = self.search_spc_file_location()
            self.check_spc_file_exist(spc_file_location)

    def search_spc_file_location(self):
        spc_file_location = ""
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()
        product_input = self.product_cb.get()
        machine_input = self.machine_cb.get()
        peel_strength_input = self.peel_strength_name.get()

        # Transform to folder name and path name
        year_folder_name = 'Data' + year_input
        result_peel_location1 = os.path.join(self.result_peel_location, year_folder_name)
        month_folder_list = month_input.split("'")
        month_folder_name = month_folder_list[0] + "." + month_folder_list[1]
        result_peel_location2 = os.path.join(result_peel_location1, month_folder_name)

        for spc_file in os.listdir(result_peel_location2):
            if not spc_file.startswith('~$') and spc_file.startswith('FLEX') and peel_strength_input == 'flex' and product_input in spc_file and machine_input in spc_file:
                spc_file_location = os.path.join(result_peel_location2, spc_file)
            elif not spc_file.startswith('~$') and spc_file.startswith('LINER') and peel_strength_input == 'liner' and product_input in spc_file and machine_input in spc_file:
                spc_file_location = os.path.join(result_peel_location2, spc_file)

        return spc_file_location

    def check_spc_file_exist(self, spc_file_location):
        if spc_file_location == "":
            msb.showinfo(title='แจ้งเตือนไปยังผู้ใช้', message='Create File')
        else:
            msb.showinfo(title='แจ้งเตือนไปยังผู้ใช้', message=f'Adjust from {spc_file_location}')

    def get_machine_no_list_from_master_file(self):
        pass

    def get_result_file_location(self):
        pass
    
    def get_testing_no_in_master(self):
        pass


app = PsaTsaProgram()
app.main_program()