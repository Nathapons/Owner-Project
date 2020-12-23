from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msb
import xlrd
import openpyxl as xl
import csv
import os
import webbrowser
import datetime

class PsaTsaProgram:
    def __init__(self):
        self.raw_data_location = '\\\\Ta1d181222\\01.result test peel strength\\DATA IPQC'
        self.master_list_location = "\\\\Ta1d180613\\2.MASTER LIST RUNNING NUMBER\\2. Master list IPQC"
        self.result_peel_location = "\\\\Ta1d180613\\3.RESULT DATA PSA\IPQC (SMT)\\2. Result peel strength (ผลการตรวจสอบจาก ANC)"
        self.spc_master_location = "\\\\Ta1d180613\\3.RESULT DATA PSA\IPQC (SMT)\\99.SPC Master(ห้ามลบ)\\QF-A1-QUA-2209-1.xlsx"

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

        # Treeview Widget
        headers = ['Testing No.', 'Lot No.', 'Report status']
        self.psa_tas_treeview = ttk.Treeview(big_frame, column=headers, show='headings', height=9)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 14, 'bold'), foreground="green")
        style.configure("Treeview", font=('Arial', 10))
        for header in headers:
            self.psa_tas_treeview.heading(header, text=header)
            self.psa_tas_treeview.column(header, anchor='center', width=170, minwidth=0)

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
        self.psa_tas_treeview.grid(row=4, column=0, pady=5)

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
        self.machine_cb.set("")
    
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
                        
                        if product_name.startswith(product_input) and machine_number_cell.strip() not in machine_number_list:
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
            # Function Overview
            master_list = self.search_and_open_master_file()
            spc_file_location, result_peel_location2 = self.search_spc_file_location()
            spc_book, spc_sheet1, spc_sheet2, filename_open= self.open_spc_file(spc_file_location)
            self.clear_psa_tsa_treeview()
            self.check_iqic_has_update(spc_book, spc_sheet1, spc_sheet2, master_list)
            self.close_and_save_spc_file(spc_book, spc_file_location, result_peel_location2, filename_open)
            
    def search_and_open_master_file(self):
        product_input = self.product_cb.get()
        machine_input = self.machine_cb.get()
        month_input = self.month_cb.get()
        year_input = self.year_cb.get()
        peel_strength_input = self.peel_strength_name.get()

        month_input_list = month_input.split("'")
        month_name = month_input_list[1][0:3]
        new_master_file_location = os.path.join(self.master_list_location, year_input)

        for excel_file_name in os.listdir(new_master_file_location):
            if month_name in excel_file_name and not excel_file_name.startswith("~$") and excel_file_name.endswith(".xlsx"):
                # Open Excel file
                master_file_location = os.path.join(new_master_file_location, excel_file_name)
                master_book = xlrd.open_workbook(filename=master_file_location)
                master_sheet = master_book.sheet_by_index(0)

                # Run function
                master_list = self.get_information_in_master(master_book, master_sheet)
                
                master_book.release_resources()
                break

        return master_list

    def get_information_in_master(self, master_book, master_sheet):
        product_input = self.product_cb.get()
        machine_input = self.machine_cb.get()
        peel_strength_input = self.peel_strength_name.get()
        master_list = []

        start_row = 8
        end_row = master_sheet.nrows
        product_name = master_sheet.cell(rowx=start_row, colx=3).value

        while str(product_name) != "":
            testing_no = str(master_sheet.cell(rowx=start_row, colx=1).value).strip()
            request_date_cell = master_sheet.cell(rowx=start_row, colx=2).value
            request_date = datetime.datetime(*xlrd.xldate_as_tuple(request_date_cell, master_book.datemode)).strftime('%x')
            lot_no = int(master_sheet.cell(rowx=start_row, colx=4).value)
            auto_press_machine = master_sheet.cell(rowx=start_row, colx=5).value

            if machine_input in auto_press_machine and product_input in product_name:
                # Get Serial Bar Code list
                serial_barcode_list = []
                if peel_strength_input == 'flex':
                    peel_column = 8
                else:
                    peel_column = 10
                for peel_row in range(start_row, start_row+5):
                    serial_barcode = master_sheet.cell(rowx=peel_row, colx=peel_column).value
                    serial_barcode_list.append(serial_barcode)

                # Run function
                peel_result_list = self.search_raw_file_location(testing_no, lot_no)
                if len(peel_result_list) == 5:
                    data_list = [testing_no, request_date, product_name, lot_no, auto_press_machine, serial_barcode_list, peel_result_list]
                    master_list.append(data_list)

            # Loop command
            start_row += 5
            product_name = master_sheet.cell(rowx=start_row, colx=3).value
        
        return master_list

    def search_raw_file_location(self, testing_no, lot_no):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()
        peel_strength_input = self.peel_strength_name.get()
        product_input = self.product_cb.get()
        peel_result_list = []
        
        iqic_folder_location = self.raw_data_location + "\\" + year_input + "\\" + month_input + "\\" + testing_no
        if os.path.isdir(iqic_folder_location):
            for file in os.listdir(iqic_folder_location):
                # Check case file name
                if "SP" in file.upper():
                    continue
                elif peel_strength_input == "flex" and file.upper().endswith("CSV") and "F" in file:
                    csv_file_location = iqic_folder_location + "\\" + file
                    peel_result_list = self.get_list_when_open_csv(csv_file_location)
                    break
                elif peel_strength_input == "flex" and file.upper().endswith("XLS") and "F" in file:
                    excel_file_location = iqic_folder_location + "\\" + file
                    peel_result_list = self.get_list_when_open_excel(excel_file_location)
                    break
                elif peel_strength_input == "liner" and file.upper().endswith("CSV") and "L" in file:
                    csv_file_location = iqic_folder_location + "\\" + file
                    peel_result_list = self.get_list_when_open_csv(csv_file_location)
                    break
                elif peel_strength_input == "liner" and file.upper().endswith("XLS") and "L" in file:
                    excel_file_location = iqic_folder_location + "\\" + file
                    peel_result_list = self.get_list_when_open_excel(excel_file_location)
                    break

        return peel_result_list

    def get_list_when_open_excel(self, excel_file_location):
        peel_book = xlrd.open_workbook(filename=excel_file_location)
        peel_sheet = peel_book.sheet_by_index(0)
        start_row = 0
        end_row = peel_sheet.nrows - 1
        peel_result_list = []

        while start_row <= end_row:
            title_name = peel_sheet.cell(rowx=start_row, colx=0).value

            if isinstance(title_name, float):
                peel_result = peel_sheet.cell(rowx=start_row, colx=1).value
                peel_result_list.append(peel_result)

            start_row += 1

        peel_book.release_resources()
        
        return peel_result_list

    def get_list_when_open_csv(self, csv_file_location):
        item_no_list = ['1', '2', '3', '4', '5']
        peel_result_list = []

        with open(csv_file_location) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')

            for row in csv_reader:
                item_no = row[0]
                if item_no in item_no_list:
                    peel_result = row[1]
                    peel_result_list.append(peel_result)
                    if item_no == '5':
                        break

        csv_file.close()
        
        return peel_result_list

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

        return spc_file_location, result_peel_location2

    def open_spc_file(self, spc_file_location):
        if spc_file_location == "":
            filename_open= self.spc_master_location
        else:
            filename_open = spc_file_location

        spc_book = xl.open(filename=filename_open, data_only=True)
        spc_sheet_list = spc_book.sheetnames
        spc_sheet1 = spc_book[spc_sheet_list[0]]
        spc_sheet2 = spc_book[spc_sheet_list[1]]

        return spc_book, spc_sheet1, spc_sheet2, filename_open

    def clear_psa_tsa_treeview(self):
        for member in self.psa_tas_treeview.get_children():
            self.psa_tas_treeview.delete(member)

    def check_iqic_has_update(self, spc_book, spc_sheet1, spc_sheet2, master_list):
        col_spc_sheet2 = 2

        for row in master_list:
            iqic_no = row[0]
            lot_no = row[3]

            has_upate_at_spc_sheet1, column_update_in_spc_sheet1 = self.iqic_no_has_update_in_spc_sheet1(spc_sheet1 , master_list, lot_no)
            # has_upate_at_spc_sheet2, column_update_in_spc_sheet2 = self.iqic_no_has_update_in_spc_sheet2(spc_sheet2 , master_list, iqic_no, lot_no)

            if has_upate_at_spc_sheet1:
                data = [iqic_no, lot_no, "อัพเดทไปแล้ว"]
            else:
                data = [iqic_no, lot_no, "โปรแกรมอัพเดท"]
            
            self.psa_tas_treeview.insert('', 'end', value=data)

    def iqic_no_has_update_in_spc_sheet1(self, spc_sheet1 , master_list, lot_no):
        has_upate_at_spc_sheet1 = False
        spc_column = 3
        column_update_in_spc_sheet1 = spc_column
        lot_no_in_spc_sheet1 = spc_sheet1.cell(row=54, column=spc_column).value
        
        while str(lot_no_in_spc_sheet1) != "":
            if lot_no == lot_no_in_spc_sheet1:
                has_upate_at_spc_sheet1 = True
                return has_upate_at_spc_sheet1, column_update_in_spc_sheet1

            # Loop command
            spc_column += 1
            lot_no_in_spc_sheet1 = spc_sheet1.cell(row=54, column=spc_column).value
            column_update_in_spc_sheet1 = spc_column

        return has_upate_at_spc_sheet1, column_update_in_spc_sheet1

    def iqic_no_has_update_in_spc_sheet2(self, spc_sheet2 , master_list, iqic_no, lot_no):
        has_upate_at_spc_sheet2 = False
        spc_column = 3
        column_update_in_spc_sheet2 = spc_column
        lot_no_in_spc_sheet2 = spc_sheet2.cell(row=4, column=spc_column).value
        
        while lot_no_in_spc_sheet2 != "":
            if lot_no == lot_no_in_spc_sheet2:
                has_upate_at_spc_sheet1 = True
                return has_upate_at_spc_sheet2, column_update_in_spc_sheet2

            # Loop command
            spc_column += 1
            lot_no_in_spc_sheet2 = spc_sheet2.cell(row=4, column=spc_column).value
            column_update_in_spc_sheet2 = spc_column

        return has_upate_at_spc_sheet2, column_update_in_spc_sheet2

    def close_and_save_spc_file(self, spc_book, spc_file_location, result_peel_location2, filename_open):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()
        product_input = self.product_cb.get()
        machine_input = self.machine_cb.get()
        peel_strength_input = self.peel_strength_name.get()
        new_date_format = month_input.split("'")[1] + "'" + year_input[2:4]

        # Set Default location and name to save file
        if spc_file_location == "" and peel_strength_input == 'flex':
            new_save_file_name = 'FLEX PEEL STRENGTH_' + product_input + "_" + machine_input + "_" + new_date_format + ".xlsx"
            new_save_file_location = os.path.join(result_peel_location2, new_save_file_name)
        elif spc_file_location == "" and peel_strength_input == 'liner':
            new_save_file_name = 'LINER PEEL STRENGTH_' + product_input + "_" + machine_input + "_" + new_date_format + ".xlsx"
            new_save_file_location = os.path.join(result_peel_location2, new_save_file_name)
        else:
            new_save_file_name = spc_file_location

        # Save file and Close Workbook
        # spc_book.save(new_save_file_location)
        spc_book.close()

        # Run function for open SPC file
        # self.ask_open_file_name(new_save_file_location)

    def ask_open_file_name(self, new_save_file_location):
        need_open_spc_file = msb.askyesno(title='แจ้งเตือนไปยังผู้ใช้', message='คุณต้องการจะเปิดไฟล์ \n' + os.path.basename(new_save_file_location))

        if need_open_spc_file:
            webbrowser.open(url=new_save_file_location)


app = PsaTsaProgram()
app.main_program()