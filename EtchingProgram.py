from tkinter import *
import openpyxl
import xlrd
import os
import time
import schedule
import datetime


class EtchingRate():
    def etching_window(self):
        self.window = Tk()
        WIDTH = 400
        HEIGHT = 100
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = int((screen_width/2) - (WIDTH/2))
        y = int((screen_height/2) - (HEIGHT/1.6))

        # UI Properties
        self.window.title("Etching Rate Import Program")
        self.window.geometry(f'{WIDTH}x{HEIGHT}+{x}+{y}')
        # self.window.attributes('-disabled', True)
        self.window.iconbitmap('fujikura_logo.ico')
        self.window.resizable(0, 0)

        # Widget
        frame = Frame(self.window)
        topics_label = Label(frame, text='Etching Rate Program', font=('Arial', 18, 'bold'))
        self.status_label = Label(frame, fg='white', bg='green', font=('Arial', 16, 'bold'))

        # Widget Position
        frame.pack()
        topics_label.grid(row=0, column=0, pady=3, ipadx=5)
        self.status_label.grid(row=1, column=0, pady=2, ipadx=5)

        # Run program overview after activate self.window at 1 second
        self.window.after(1000, self.etching_overview)

        # Activate GUI
        self.window.mainloop()

    def etching_overview(self):
        time_now = datetime.datetime.now()
        self.status_label['text'] = "Etching Import on " + time_now.strftime('%x') + " " + time_now.strftime('%X')

        # Run Program
        self.search_etching_record_excel()

        self.window.after(1000, self.etching_overview)
        # Run after 6hrs
        # self.window.after(14400, self.etching_overview)

    
    def search_etching_record_excel(self):
        etching_record_path = "\\\\10.17.164.209\\iot"
        month_list = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

        for factory_name in os.listdir(etching_record_path):
            factory_path = os.path.join(etching_record_path, factory_name)

            for year in os.listdir(factory_path):
                year_path = os.path.join(factory_path, year)

                for record_file in os.listdir(year_path):
                    month_year_record = record_file.split(" ")[1]
                    month_name = month_year_record[0:3]

                    if month_name in month_list and not record_file.startswith("~$"):
                        month_number = month_list.index(month_name) + 1
                        etching_record_location = os.path.join(year_path, record_file)
                        self.open_close_etching_record_excel(etching_record_location, year, month_name, month_number)

    def open_close_etching_record_excel(self, etching_record_location, year, month_name, month_number):
        etching_record_wb = xlrd.open_workbook(filename=etching_record_location)
        machine_no_sheet_list = etching_record_wb.sheet_names()

        for machine_no in machine_no_sheet_list:
            if len(machine_no) == 7:
                self.check_etching_format_exits(year, month_name, month_number, machine_no)
        
        etching_record_wb.release_resources()

    
    def check_etching_format_exits(self, year, month_name, month_number, machine_no):
        spc_path = "\\\\ta1d171009\\Users\\wissanu.t\\Desktop\\SPC C2R"
        etching_file_location = f'{spc_path}\ค่า Etching rate  -{year}\{month_number}.{month_name}\{machine_no}'
        print(etching_file_location)


app = EtchingRate()
app.etching_window()