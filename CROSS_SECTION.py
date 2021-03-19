from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msb
import openpyxl as xl
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
import xlrd
from pathlib import Path
import os
import webbrowser

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
    
    def program_overview(self):
        result_path = f'{self.link}\{self.product_box.get()}\วัดแล้ว\{self.lot_box.get()}\{self.lot_box.get()}.xlsx'
        report_path = f'D:\\Nathapon\\0.My work\\01.IoT\\06.VBA\\06.CROSS_SECTION\\Master Report\Report_{self.product_box.get()}.xlsx'

        # Check file Exist in path?
        if os.path.isfile(result_path) and os.path.isfile(report_path):
            # Get Parameter Return Form open_excel_file function
            result_wb = self.open_excel_file_read_state(excel_path=result_path)
            report_wb = self.oepn_excel_file(excel_path=report_path)
            
            # Clear member in status_treeview
            self.clear_treeview()

            # Run function each sheetname
            report_sheets = report_wb.sheetnames
            for sheetname in report_sheets:
                report_ws = report_wb[sheetname]
                status = 'Not Import'
                
                if sheetname == "OQC":
                    pass
                elif sheetname == 'Solder Mask':
                    result_ws = result_wb.sheet_by_name('Solder Mask Coverage')
                    status = self.soldermask_document(report_ws, result_ws)
                elif sheetname == 'Hot bar':
                    result_ws = result_wb.sheet_by_name('hotbar ')
                    status = self.hotbar_document(report_ws=report_ws, result_ws=result_ws)
                elif sheetname == 'Addtion X-section and Thickness':
                    result_ws = result_wb.sheet_by_name('X-section')
                    self.addition_program(report_ws, result_ws)
                elif sheetname == 'FAI':
                    result_ws = result_wb.sheet_by_name('FAI')
                    status = self.fai_program(report_ws, result_ws)

                data = [sheetname, status]
                self.status_tree.insert('', 'end', value=data)

            file_saveas = f'{self.link}\{self.product_box.get()}\วัดแล้ว\{self.lot_box.get()}\Report_{self.lot_box.get()}.xlsx'
            try:
                result_wb.release_resources()
                report_wb.save(filename=file_saveas)
                report_wb.close()

                # Ask user to open excel file
                ask = msb.askyesno(title="Ask to user", message=f'คุณต้องการเปิดไฟล์ {os.path.basename(file_saveas)} หรือไม่?')
                if ask:
                    webbrowser.open(file_saveas)
                else:
                    msb.showinfo(title="Ask to User", message='Import Complete')
            except Exception:
                msb.showwarning(title='Alarm to user', message=f'กรุณาปิดไฟล์ {os.path.basename(file_saveas)} แล้วรันโปรแกรมอีกครั้ง!')

        else:
            msb.showwarning(message=f'ไม่มีไฟล์ที่ชื่อ \n {result_path}', title='Alarm To User')

    def oepn_excel_file(self, excel_path):
        wb = xl.open(filename=excel_path)
        return wb

    def open_excel_file_read_state(self, excel_path):
        wb = xlrd.open_workbook(filename=excel_path)
        return wb

    def clear_treeview(self):
        for member in self.status_tree.get_children():
            self.status_tree.delete(member)

    # ------------------------------------------------ Cross Section Page ---------------------------------------------------------
    def oqc_document(self, report_ws, result_ws):
        try:
            status = 'Complete'
        except Exception:
            status = 'Program Error!'

        return status

    # ------------------------------------------------ Solder Mask Page -----------------------------------------------------------
    def soldermask_document(self, report_ws, result_ws):
        try:
            flex_row, bga_pic_row, hotbar_pic_row, con_pic_row, hotbar_col, bga_col, thick_col, off_col = self.solder_get_pasteposition(report_ws=report_ws)
            # Import data to Document
            import_name = self.hotbar_import_data(report_ws, result_ws, flex_row, hotbar_col, bga_col, thick_col, off_col)
            # Import picture to ducument
            self.hotbar_import_picture(report_ws, bga_pic_row, hotbar_pic_row, con_pic_row, import_name)
            status = 'COMPLETE'
        except Exception:
            status = 'Program Error!!'

        return status

    def solder_get_pasteposition(self, report_ws):
        report_row = 1
        hotbar_col = 1
        bga_col = 1
        thick_col = 1
        off_col = 1
        horbar_pic_row = 1

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
            elif 'BGA' in detail:
                bga_pic_row = report_row - 1
            elif 'HOT' in detail:
                hotbar_pic_row = report_row - 1
            elif 'CONNECTOR' in detail:
                con_pic_row = report_row - 1

            # Loop command
            report_row += 1

        return flex_row, bga_pic_row, hotbar_pic_row, con_pic_row, hotbar_col, bga_col, thick_col, off_col

    def hotbar_import_data(self, report_ws, result_ws, flex_row, hotbar_col, bga_col, thick_col, off_col):
        b2b_offset = result_ws.cell(rowx=6, colx=2).value
        b2b_coverage = result_ws.cell(rowx=7, colx=2).value
        b2b_thickness = result_ws.cell(rowx=8, colx=2).value

        if (not b2b_offset) == False and (not b2b_coverage) == False and (not b2b_thickness) == False:
            # import B2B/BGA to format
            import_name = 'B2B/BGA'
            result_offset_row = 6
            result_coverage_row = 7
            result_thickness_row = 8
            report_cover_column = bga_col
        else:
            # import hotbar to format
            import_name = 'HOT BAR'
            result_offset_row = 13
            result_coverage_row = 14
            result_thickness_row = 15
            report_cover_column = hotbar_col

        result_col = 2
        while result_col <= result_ws.ncols-1:
            # Get value from Result.xlsx
            left_coverage = result_ws.cell(rowx=result_coverage_row, colx=result_col).value
            right_coverage = result_ws.cell(rowx=result_coverage_row, colx=result_col+1).value
            offset = result_ws.cell(rowx=result_offset_row, colx=result_col).value
            thickness = result_ws.cell(rowx=result_thickness_row, colx=result_col).value
            # Send value to report.xlsx
            report_ws.cell(row=flex_row, column=report_cover_column).value = left_coverage
            report_ws.cell(row=flex_row, column=report_cover_column+1).value = right_coverage
            report_ws.cell(row=flex_row, column=thick_col).value = thickness
            report_ws.cell(row=flex_row, column=off_col).value = offset
            
            # Loop command
            result_col += 2
            flex_row += 1

        return import_name

    def get_b2b_row_result(self, result_ws):
        max_row = result_ws.nrows

    def hotbar_import_picture(self, report_ws, bga_pic_row, hotbar_pic_row, con_pic_row, import_name):
        b2b_pic_path, hotbar_pic_path = self.get_pic_folder()
        b2b_pics = sorted(Path(b2b_pic_path).iterdir(), key=os.path.getmtime)
        hotbar_pics = sorted(Path(hotbar_pic_path).iterdir(), key=os.path.getmtime)

        if import_name == 'B2B/BGA':
            row_export = bga_pic_row
        else:
            row_export = hotbar_pic_row

        # Import Hot Bar& B2B Picture
        COLUMN_INSERT = 1
        for i in range(1, len(b2b_pics), 2):
            if str(b2b_pics[i]).upper().endswith('JPG'):
                # Get Picture
                first_pic = b2b_pics[i-1]
                second_pic = b2b_pics[i]
                # Call add image
                first_img = Image(first_pic)
                second_img = Image(second_pic)

                p2e = pixels_to_EMU
                c2e = cm_to_EMU
                # Assign picture size
                HEIGHT = 50
                WIDTH = 50
                # Function calculate offset
                cellh = lambda x: c2e((x * 49.77)/99)
                cellw = lambda x: c2e((x * (18.65-1.71))/10)
                # Set Size and Postion
                colloff1 = cellw(0.1)
                colloff2 = cellw(1)
                rowoffset = cellh(0.5)
                first_marker = AnchorMarker(col=COLUMN_INSERT, colOff=colloff1, row=row_export, rowOff=rowoffset)
                second_marker = AnchorMarker(col=COLUMN_INSERT, colOff=colloff2, row=row_export, rowOff=rowoffset)
                size = XDRPositiveSize2D(p2e(HEIGHT), p2e(WIDTH))
                # Paste Image to cell
                first_img.anchor = OneCellAnchor(_from=first_marker, ext=size)
                report_ws.add_image(first_img)
                second_img.anchor = OneCellAnchor(_from=second_marker, ext=size)
                report_ws.add_image(second_img)
                
                # Prepare for next
                COLUMN_INSERT += 1

        # Import Connector Picture
        COLUMN_INSERT = 1
        for i in range(1, len(hotbar_pics), 2):
            if str(hotbar_pics[i]).upper().endswith('JPG'):
                # Get Picture
                first_pic = hotbar_pics[i-1]
                second_pic = hotbar_pics[i]
                # Call add image
                first_img = Image(first_pic)
                second_img = Image(second_pic)

                p2e = pixels_to_EMU
                c2e = cm_to_EMU
                # Assign picture size
                HEIGHT = 50
                WIDTH = 50
                # Function calculate offset
                cellh = lambda x: c2e((x * 49.77)/99)
                cellw = lambda x: c2e((x * (18.65-1.71))/10)
                # Set Size and Postion
                colloff1 = cellw(0.1)
                colloff2 = cellw(1)
                rowoffset = cellh(0.5)
                first_marker = AnchorMarker(col=COLUMN_INSERT, colOff=colloff1, row=con_pic_row, rowOff=rowoffset)
                second_marker = AnchorMarker(col=COLUMN_INSERT, colOff=colloff2, row=con_pic_row, rowOff=rowoffset)
                size = XDRPositiveSize2D(p2e(HEIGHT), p2e(WIDTH))
                # Paste Image to cell
                first_img.anchor = OneCellAnchor(_from=first_marker, ext=size)
                report_ws.add_image(first_img)
                second_img.anchor = OneCellAnchor(_from=second_marker, ext=size)
                report_ws.add_image(second_img)
                
                # Prepare for next
                COLUMN_INSERT += 1

    def get_pic_folder(self):
        soldermask_folder = self.get_link_soldermask_folder()

        for pic_folder in os.listdir(soldermask_folder):
            if 'B2B' in pic_folder.upper():
                b2b_pic_path = os.path.join(soldermask_folder, pic_folder)
            elif 'HOT' in pic_folder.upper() or 'BGA' in pic_folder.upper():
                hotbar_pic_path = os.path.join(soldermask_folder, pic_folder)

        return b2b_pic_path, hotbar_pic_path

    def get_link_soldermask_folder(self):
        folder_path = f'{self.link}\{self.product_box.get()}\วัดแล้ว\{self.lot_box.get()}'
        folder_list = os.listdir(folder_path)

        for folder_name in folder_list:
            if 'SOLDER' in folder_name.upper():
                return os.path.join(folder_path, folder_name)

    # ---------------------------------------------------- Hot Bar Page ---------------------------------------------------------
    def hotbar_document(self, report_ws, result_ws):
        try:
            report_row, soldera_col, solderb_col, solderbstar_col = self.get_hotbar_row(report_ws=report_ws)
            report_row += 1
            result_row = 17

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

            status = 'COMPLETE'
        except Exception:
            status = 'Program Error!'

        return status

    def get_hotbar_row(self, report_ws):
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
    # ---------------------------------------------------- Addtion X-section and Thickness ---------------------------------------------------------
    def addition_program(self, report_ws, result_ws):
        result_col = 1
        max_col = result_ws.ncols - 1

        while result_col < max_col:
            region_no = result_ws.cell(rowx=1, colx=result_col).value

            if isinstance(region_no, float):
                region_no = int(region_no)
                row, col = self.get_addition_row(report_ws, region_no)
                layer_dict = self.get_layer_result(result_ws, result_col)
                self.import_addition_to_report(report_ws, layer_dict, row, col)

            # Prepare for next column
            result_col += 5

    def get_addition_row(self, report_ws, region_no):
        row = 1
        col = 1
        max_row = report_ws.max_row
        max_col = report_ws.max_column + 1

        while row <= max_row:
            name = str(report_ws.cell(row=row, column=1).value).upper()
            number = str(report_ws.cell(row=row+1, column=1).value)
            
            if name == 'REGION' and number == str(region_no):
                for col in range(1, max_col):
                    subdetail = report_ws.cell(row=row, column=col).value
                    fill_row = row + 1
                    if '#1' in subdetail:
                        return fill_row, col

            # Loop command
            row += 1

    def get_layer_result(self, result_ws, result_col):
        layer_dict = {}
        end_layer_col = result_col + 5
        layer_row = 4
        
        for col in range(result_col, end_layer_col):
            layer_no = result_ws.cell(rowx=2, colx=col).value

            layer_result_list = []
            while layer_row <= 12:
                layer_result = result_ws.cell(rowx=layer_row, colx=col).value
                layer_result_list.append(layer_result)

                # Loop command
                layer_row += 1
            
            # Prepare for next column
            layer_dict[layer_no] = layer_result_list
            layer_row = 4

        return layer_dict

    def import_addition_to_report(self, report_ws, layer_dict, start_row, col):
        row = start_row

        for layer_no in layer_dict:
            layer_results = layer_dict[layer_no]

            for layer_result in layer_results:
                report_ws.cell(row=row, column=col).value = layer_result
                row += 1
            
            # Reset row col to prepare paste next layer
            row = start_row
            col += 1

    # ------------------------------------------------------------------ FAI -----------------------------------------------------------------------
    def fai_program(self, report_ws, result_ws):
        fai_length_list = self.get_data_from_result_ws(result_ws)
        report_row = self.get_fai_row(report_ws)
        self.import_fai_to_report(report_ws, fai_length_list, report_row)
        status = 'COMPLETE'

        # try:
        #     fai_length_list = self.get_data_from_result_ws(result_ws)
        #     report_row = self.get_fai_row(report_ws)
        #     self.import_fai_to_report(fai_lenght_list, report_row)
        # except Exception:
        #     status = 'Program Error!'

        return status

    def get_data_from_result_ws(self, result_ws):
        row = 24
        fai_length_list = []

        for col in range(1, 7):
            fai_length = round(result_ws.cell(rowx=row, colx=col).value, 2)
            fai_length_list.append(fai_length)

        return fai_length_list

    def get_fai_row(self, report_ws):
        row = 1
        max_row = report_ws.max_row

        while row <= max_row:
            detail = report_ws.cell(row=row, column=1).value
            # Get Row Filled
            if str(detail) == '1':
                break

            # Loop command
            row += 1

        return row

    def import_fai_to_report(self, report_ws, fai_length_list, report_row):
        for fai_length in fai_length_list:
            report_ws.cell(row=report_row, column=2).value = fai_length
            # Prepare for next row
            report_row += 1


# Setting to Run Program
app = CrossSection()
app.main_window()