from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msb
import openpyxl as xl
import xlrd
import os


class CrossSection():
    def __init__(self):
        # self.link = "\\\\10.17.73.53\\ORT_Result\\03.Data Cross Section\\1.Cross section\\1.For Per Day"
        self.link = 'D:\\Nathapon\\0.My work\\01.IoT\\06.VBA\\06.CROSS_SECTION\\1.For Per Day'
        self.topics_font = ('Arial', 22, 'bold')
        self.detail_font = ('Arial', 16)
        
    def main_window(self):
        root = Tk()
        WIDTH = 500
        HEIGHT = 400
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_width = int((screen_width/2) - (WIDTH/2))
        center_height = int((screen_height/2) - (HEIGHT/2))

        # PROGRAM PROPERTIES
        root.title("CROSS SECTION PROGRAM")
        root.iconbitmap('fujikura_logo.ico')
        root.geometry(f'{WIDTH}x{HEIGHT}+{center_width}+{center_height}')
        root.config(bg='white')
        root.resizable(0, 0)

        # combostyle = ttk.Style()
        # combostyle.theme_create('combostyle', parent='alt',
        #                  settings = {'TCombobox':
        #                              {'configure':
        #                               {'fieldbackground': '#ffffb3',
        #                                'selectbackground': 'blue',
        #                                'background': 'white'
        #                                }}})
        # combostyle.theme_use('combostyle')

        # WIDGET PROPERTIES
        big_frame = Frame(root, bg='white')
        topics_label = Label(big_frame, text='CROSS SECTION PROGRAM', font=self.topics_font, fg='white', bg='blue')
        creator_label = Label(big_frame, text='Create: Nathapon.S Tel: 4308', font=self.detail_font, bg='white')
        product_frame = Frame(big_frame, bg='white')
        product_label = Label(product_frame, text='Product:', font=self.detail_font, width=8, bg='white')
        self.product_box = ttk.Combobox(product_frame, justify='center', font=self.detail_font, width=20)
        lot_frame = Frame(big_frame, bg='white')
        lot_label = Label(lot_frame, text='Lot no:', font=self.detail_font, width=8, bg='white')
        self.lot_box = ttk.Combobox(lot_frame, justify='center', font=self.detail_font, width=20)
        run_button = Button(big_frame, text='Run Program', command=self.checkUserFilled, width=15, font=('Arial', 16, 'bold'), bg='#33cc00', fg='white')
        
        # Status Report
        headers = ['TOPICS', 'STATUS']
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        self.status_tree = ttk.Treeview(big_frame, column=headers, show='headings', height=6)
        for header in headers:
            self.status_tree.heading(header, text=header)
            col_width = 200
            self.status_tree.column(header, anchor='center', width=col_width, minwidth=0)
        self.status_tree.tag_configure('odd', background='#E8E8E8')
        self.status_tree.tag_configure('even', background='#DFDFDF')


        # Add Event in Combobox
        self.product_box.bind('<Button-1>', self.product_box_click)
        self.lot_box.bind('<Button-1>', self.lotno_click)

        # WIDGET POSITION
        big_frame.pack()
        topics_label.grid(row=0, column=0, ipadx=5, pady=5)
        creator_label.grid(row=1, column=0, ipadx=5)
        product_frame.grid(row=2, column=0, pady=6)
        product_label.grid(row=0, column=0, padx=5)
        self.product_box.grid(row=0, column=1, padx=5, ipady=2)
        lot_frame.grid(row=3, column=0, pady=6)
        lot_label.grid(row=0, column=0, padx=5)
        self.lot_box.grid(row=0, column=1, padx=5, ipady=2)
        run_button.grid(row=4, column=0, pady=10)
        self.status_tree.grid(row=5, column=0, pady=2)

        # Activate GUI
        root.mainloop()

    def product_box_click(self, event):
        self.lot_box.set("")
        self.product_box['values'] = os.listdir(self.link)

    def lotno_click(self, event):
        product_select = self.product_box.get()

        if not product_select:
            msb.showwarning(title='Alarm to User', message='กรุณากรอกข้อมูลที่ช่อง Product')
        else:
            new_path = f'{os.path.join(self.link, product_select)}\วัดแล้ว'
            folder_list = os.listdir(new_path)
            self.lot_box['values'] = folder_list

    def checkUserFilled(self):
        # print(self.product_box.get(), self.lot_box.get())
        if (not self.product_box.get()) and (not self.lot_box.get()):
            msb.showwarning(title='Alarm to User', message='คุณยังกรอกข้อมูลไม่ครบถ้วน')
        else:
            self.program_overview()
            msb.showinfo(title='Information', message='Complete')
    
    def program_overview(self):
        result_path = f'{self.link}\{self.product_box.get()}\วัดแล้ว\{self.lot_box.get()}\{self.lot_box.get()}.xlsx'
        report_path = f'{self.link}\{self.product_box.get()}\วัดแล้ว\{self.lot_box.get()}\Report_{self.lot_box.get()}.xlsx'

        try:
            # Get Parameter Return Form open_excel_file function
            result_wb = self.open_excel_file_read_state(excel_path=result_path)
            report_wb = self.oepn_excel_file(excel_path=report_path)
            
            # Clear member in status_treeview
            self.clear_treeview()

            # Run function each sheetname
            report_sheets = report_wb.sheetnames
            for sheetname in report_sheets:
                report_ws = report_wb[sheetname]
                result = 'Update'
                
                if sheetname == "OQC":
                    pass
                elif sheetname == 'Solder Mask':
                    # print('Run')
                    result_ws = result_wb.sheet_by_name('Solder Mask Coverage')
                    self.soldermask_document(report_ws, result_ws)
                    pass
                elif sheetname == 'Hot bar':
                    result_ws = result_wb.sheet_by_name('hotbar ')
                    self.hotbar_document(report_ws=report_ws, result_ws=result_ws)
                    result = 'COMPLETE'
                elif sheetname == 'Addtion X-section and Thickness':
                    pass
                elif sheetname == 'FAI':
                    pass

                data = [sheetname, result]
                self.status_tree.insert('', 'end', value=data)

            result_wb.release_resources()
            report_wb.save(filename=f'Report_{self.lot_box.get()}.xlsx')
            report_wb.close()
        except Exception:
            print("Error")

    def oepn_excel_file(self, excel_path):
        wb = xl.open(filename=excel_path)
        return wb

    def open_excel_file_read_state(self, excel_path):
        wb = xlrd.open_workbook(filename=excel_path)
        return wb

    def clear_treeview(self):
        for member in self.status_tree.get_children():
            self.status_tree.delete(member)

    def oqc_document(self, report_ws, result_ws):
        pass

    # ------------------------------------------------ Solder Mask Page ---------------------------------------------------------
    def soldermask_document(self, report_ws, result_ws):
        flex_row, bga_pic_row, con_pic_row, hotbar_col, bga_col, thick_col, off_col = self.solder_get_pasteposition(report_ws=report_ws)

    def solder_get_pasteposition(self, report_ws):
        report_row = 1
        hotbar_col = 1
        bga_col = 1
        thick_col = 1
        off_col = 1

        while report_row != report_ws.max_row:
            detail = str(report_ws.cell(row=report_row, column=1).value).upper()

            if detail == 'FLEX NO.':
                for column in range(1, 14):
                    subdetail = str(report_ws.cell(row=report_row, column=column).value)
                    if subdetail == 'Coverage (Hot bar area)':
                        hotbar_col = column
                    elif subdetail == 'Coverage (BGA area or the nearest trace)':
                        bga_col = column
                    elif subdetail == 'solder mask thickness':
                        thick_col = column
                    elif subdetail == 'Solder mask Offset (clear from Cu)':
                        off_col = column

            elif detail == '1':
                flex_row = report_row
            elif detail == 'BGA':
                bga_pic_row = report_row
            elif 'CONNECTOR' in detail:
                con_pic_row = report_row

            # Loop command
            report_row += 1

        return flex_row, bga_pic_row, con_pic_row, hotbar_col, bga_col, thick_col, off_col

    def hotbar_import_data(self, flex_row, hotbar_col, bga_col, thick_col, off_col):
        pass

    # ---------------------------------------------------- Hot Bar Page ---------------------------------------------------------
    def hotbar_document(self, report_ws, result_ws):
        report_row, soldera_col, solderb_col, solderbstar_col = self.get_horbar_row(report_ws=report_ws)
        report_row += 1
        result_row = 18

        while result_ws.cell(rowx=result_row, colx=2).value != "":
            # Receive Value in cell
            solder_a = round(result_ws.cell(rowx=result_row, colx=2).value, 3)
            solder_b = round(result_ws.cell(rowx=result_row, colx=3).value, 3)
            solder_bstar = round(result_ws.cell(rowx=result_row, colx=4).value, 3)
            # Send value to cell
            report_ws.cell(row=report_row, column=soldera_col).value = solder_a
            report_ws.cell(row=report_row, column=solderb_col).value = solder_b
            report_ws.cell(row=report_row, column=solderbstar_col).value = solder_bstar

            # Loop command
            result_row += 1
            report_row += 1


    def get_horbar_row(self, report_ws):
        report_row = 1
        column = 2

        while 'Supplier Name' not in str(report_ws.cell(row=report_row, column=column).value):
            report_row += 1

        while column != 14:
            measure_name = report_ws.cell(row=report_row, column=column).value
            if 'Solder Mask Thickness - A' in measure_name:
                soldera_col = column
                solderb_col = column + 1
                solderbstar_col = column + 2  
            # Loop command
            column += 1

        return report_row, soldera_col, solderb_col, solderbstar_col


# Setting to Run Program
app = CrossSection()
app.main_window()