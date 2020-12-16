from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msb
import xlrd
import openpyxl as xl
import os

class PsaTsaProgram:
    def __init__(self):
        self.raw_data_location = '\\\\Ta1d181222\\01.result test peel strength\\DATA IPQC'
        self.master_data_location = "\\\\Ta1d180613\\2.MASTER LIST RUNNING NUMBER\\2. Master list IPQC"
        
        # Widget Properties
        self.topic_font = ('Arial', 24, 'bold')
        self.detail_font = ('Arial', 16)

    def main_program(self):
        self.main_window = Tk()
        WIDTH = 400
        HEIGHT = 230
        screen_width = self.main_window.winfo_screenwidth()
        screen_height = self.main_window.winfo_screenheight()
        x = int( (screen_width/2) - (WIDTH/2) )
        y = int( (screen_height/2) - (HEIGHT/1.6) )

        # GUI Properties
        self.main_window.title('IPQC-PSATSA Program')
        self.main_window.geometry(f'{WIDTH}x{HEIGHT}+{x}+{y}')

        # Widget Properties
        big_frame = Frame(self.main_window)
        topic_label = Label(big_frame, text='IPQC-PSATSA Program', font=self.topic_font, fg='white', bg='blue')
        creator_label = Label(big_frame, text='Create: Nathapon.S IoT Division', font=self.detail_font)
        spc_program_button = Button(big_frame, text='SPC Report Program', font=('Arial', 18), command=self.spc_userinterface, width=22)
        control_limit_button = Button(big_frame, text='Contorl Limit Program', font=('Arial', 18), command=self.control_limit_report_window, width=22)

        # Widget Position
        big_frame.pack()
        topic_label.grid(row=0, column=0, pady=10)
        creator_label.grid(row=1, column=0, pady=2)
        spc_program_button.grid(row=2, column=0, pady=5)
        control_limit_button.grid(row=3, column=0, pady=5)
        
        self.main_window.mainloop()

    def hide_main_window(self):
        self.main_window.withdraw()

    def spc_userinterface(self):
        self.hide_main_window()

        self.spc_window = Toplevel()
        WIDTH = 550
        HEIGHT = 330
        screen_width = self.spc_window.winfo_screenwidth()
        screen_height = self.spc_window.winfo_screenheight()
        x = int( (screen_width/2) - (WIDTH/2) )
        y = int( (screen_height/2) - (HEIGHT/1.6))

        # GUI Properties
        self.spc_window.title('IPQC-PSATSA Program')
        self.spc_window.resizable(0, 0)
        self.spc_window.geometry(f'{WIDTH}x{HEIGHT}+{x}+{y}')
        self.spc_window.protocol("WM_DELETE_WINDOW", self.close_spc_window)

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
        doc_req_label = Label(select_frame, text='Report Name:', font=self.detail_font)
        check_box_frame = Frame(select_frame)
        self.doc_radio = StringVar()
        psa_check_box = Radiobutton(check_box_frame, text='PSA Report', font=self.detail_font, variable=self.doc_radio, value='psa')
        tsa_check_box = Radiobutton(check_box_frame, text='TSA Report', font=self.detail_font, variable=self.doc_radio, value='tsa')
        run_program_button = Button(big_frame, text='Run Program', font=('Arial', 18, 'bold'), command=self.spc_program_overview)

        # Widget Event
        self.year_cb.bind('<Button-1>', self.year_click)
        self.month_cb.bind('<Button-1>', self.month_click)
        self.product_cb.bind('<Button-1>', self.product_click)
        psa_check_box.invoke()

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
        doc_req_label.grid(row=3, column=0, padx=5, pady=5)
        check_box_frame.grid(row=3, column=1, padx=5, pady=5)
        psa_check_box.grid(row=0, column=0, padx=5)
        tsa_check_box.grid(row=0, column=1, padx=5)
        run_program_button.grid(row=2, column=0, pady=5, ipadx=20)

        # GUI Activate
        self.spc_window.mainloop()

    def close_spc_window(self):
        self.spc_window.destroy()
        self.main_program()

    def year_click(self, event):
        self.year_cb['values'] = os.listdir(self.raw_data_location)
    
    def month_click(self, event):
        year_input = self.year_cb.get()

        if year_input != "":
            month_all_folders_location = os.path.join(self.raw_data_location, year_input)
            month_folders_list = os.listdir(month_all_folders_location)

            self.month_cb['values'] = month_folders_list
        else:
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Year Report ก่อน')
            self.month_cb['values'] = []
            self.month_cb.set("")
    
    def product_click(self, event):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()

        if year_input != "" and month_input != "": 
            month_input_list = month_input.split("'")
            month_name = month_input_list[1][0:3]
            new_master_file_location = os.path.join(self.master_data_location, year_input)
            product_name_list = []

            for excel_file_name in os.listdir(new_master_file_location):
                if month_name in excel_file_name and not excel_file_name.startswith("~$") and excel_file_name.endswith(".xlsx"):
                    master_file_location = os.path.join(new_master_file_location, excel_file_name)
                    print(f'open {excel_file_name}')
                    # Open master file
                    master_book = xlrd.open_workbook(filename=master_file_location)
                    master_sheet = master_book.sheet_by_index(0)
                    start_row = 8
                    product_name_cell = master_sheet.cell(rowx=start_row, colx=3).value

                    while product_name_cell != "":
                        product_name = str(product_name_cell)[0:11]
                        product_version = str(product_name_cell)[11:].replace("O", "0")
                        product_full_name = product_name + product_version

                        # Add data to list when product startwith RG and don't have add to list
                        if product_full_name.upper() not in product_name_list and product_full_name.upper().startswith('RG'):
                            product_name_list.append(product_full_name.upper())

                        start_row += 5
                        product_name_cell = master_sheet.cell(rowx=start_row, colx=3).value

                    master_book.release_resources()
            
            self.product_cb['values'] = product_name_list

        elif year_input != "" and month_input == "":
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Month Report ก่อน')
            self.product_cb['values'] = []
            self.product_cb.set("")
        elif year_input == "" and month_input != "":
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Year Report ก่อน')
            self.product_cb['values'] = []
            self.product_cb.set("")
        else:
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'กรุณากรอกข้อมูลที่ช่อง Year Report และ Month Report ก่อน')
            self.product_cb['values'] = []
            self.product_cb.set("")
        
    def spc_program_overview(self):
        year_input = self.year_cb.get()
        month_input = self.month_cb.get()
        product_input = self.product_cb.get()

        if year_input != "" and month_input != "" and product_input != "":
            pass
        else:
            msb.showwarning('แจ้งเตือนไปยังผู้ใช้', 'คุณยังกรอกข้อมูลไม่ครบถ้วน')
    
    def control_limit_report_window(self):
        self.hide_main_window()

        self.control_window = Toplevel()

        self.control_window.mainloop()


app = PsaTsaProgram()
app.main_program()