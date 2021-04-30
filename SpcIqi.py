from tkinter import *
from tkinter import ttk
from tkinter.tix import *
from tkinter import messagebox as msb
import sqlite3
import openpyxl as xl
import os

class SpcIqi():
    def __init__(self):
        try:
            self.spec_datatable()
        except Exception:
            pass

        self.topics_font = ('Arial', 22, 'bold')
        self.detail_font = ('Arial', 16)
        self.entry_width = 15
        self.combobox_width = 13
        

    def spec_datatable(self):
        with sqlite3.connect('\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\Database\IQI.DB') as con:
            query = """
                CREATE TABLE control_limits (
                    ID integer PRIMARY KEY AUTOINCREMENT,
                    MATERIAL varchar(255),
                    METHOD varchar(255),
                    USL FLOAT,
                    TARGET FLOAT,
                    LSL FLOAT,
                    UCLX FLOAT,
                    CLX FLOAT,
                    LCLX FLOAT,
                    CREATED_ON  DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """
            con.execute(query)

    def main_window(self):
        self.root = Tk()
        WIDTH = 330
        HEIGHT = 210
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_width = int((screen_width/2) - (WIDTH/2))
        center_height = int((screen_height/2) - (HEIGHT/1.8))

        self.root.title('Main Menu')
        self.root.geometry(f'{WIDTH}x{HEIGHT}+{center_width}+{center_height}')
        self.root.config(bg='white')
        self.root.resizable(0, 0)

        big_frame = Frame(self.root, bg='white')
        topic_label = Label(big_frame, text='SPC PROGRAM', font=self.topics_font, fg='white', bg='blue')
        creator_label = Label(big_frame, text='Create: Nathapon.S Tel: 4308', font=self.detail_font, bg='white')
        control_limit = Button(big_frame, text='CONTROL LIMIT MENU', font=self.detail_font, width=20, bg='white', command=self.control_limit_window)
        import_spc = Button(big_frame, text='SPC REPORT MENU', font=self.detail_font, width=20, bg='white', command=self.import_spc_report)

        big_frame.pack()
        topic_label.grid(row=0, column=0, pady=5, ipadx=20)
        creator_label.grid(row=1, column=0)
        control_limit.grid(row=2, column=0, pady=5, ipady=2)
        import_spc.grid(row=3, column=0, pady=5, ipady=2)
        self.root.mainloop()

    def control_limit_window(self):
        self.root.withdraw()

        self.control_window = Toplevel()
        tip = Balloon(self.control_window)
        WIDTH = 1100
        HEIGHT = 500
        screen_width = self.control_window.winfo_screenwidth()
        screen_height = self.control_window.winfo_screenheight()
        center_width = int((screen_width/2) - (WIDTH/2))
        center_height = int((screen_height/2) - (HEIGHT/1.8))

        self.control_window.title('Control Limit Window')
        self.control_window.geometry(f'{WIDTH}x{HEIGHT}+{center_width}+{center_height}')
        self.control_window.resizable(0, 0)
        self.control_window.protocol('WM_DELETE_WINDOW', self.close_control_limit_window)

        big_frame = Frame(self.control_window)
        topics_label = Label(big_frame, text='IQI Admin System', font=self.topics_font, fg='white', bg='blue', width=15)
        second_frame = Frame(big_frame)
        material_label = Label(second_frame, text='Material:', font=self.detail_font)
        self.material = ttk.Combobox(second_frame, font=self.detail_font, width=self.combobox_width)
        method_label = Label(second_frame, text='Method:', font=self.detail_font)
        self.method= ttk.Combobox(second_frame, font=self.detail_font, width=self.combobox_width)
        usl_label = Label(second_frame, text='USL:', font=self.detail_font)
        self.usl= ttk.Entry(second_frame, font=self.detail_font, width=self.entry_width)
        target_label = Label(second_frame, text='Target:', font=self.detail_font)
        self.target= ttk.Entry(second_frame, font=self.detail_font, width=self.entry_width)
        lsl_label = Label(second_frame, text='LSL:', font=self.detail_font)
        self.lsl = ttk.Entry(second_frame, font=self.detail_font, width=self.entry_width)
        uclx_label = Label(second_frame, text='UCL:', font=self.detail_font)
        self.uclx = ttk.Entry(second_frame, font=self.detail_font, width=self.entry_width)
        clx_label = Label(second_frame, text='CL:', font=self.detail_font)
        self.clx = ttk.Entry(second_frame, font=self.detail_font, width=self.entry_width)
        lclx_label = Label(second_frame, text='LCL:', font=self.detail_font)
        self.lclx = ttk.Entry(second_frame, font=self.detail_font, width=self.entry_width)

        # ToolTip
        tip.bind_widget(self.material, balloonmsg='กรุณาเลือกข้อมูลในลิสต์')
        tip.bind_widget(self.method, balloonmsg='กรุณาเลือกข้อมูลในลิสต์')
        tip.bind_widget(self.uclx, balloonmsg='กรุณากรอกเป็นตัวเลข')
        tip.bind_widget(self.clx, balloonmsg='กรุณากรอกเป็นตัวเลข')
        tip.bind_widget(self.lclx, balloonmsg='กรุณากรอกเป็นตัวเลข')
        tip.bind_widget(self.usl, balloonmsg='กรุณากรอกเป็นตัวเลข')
        tip.bind_widget(self.target, balloonmsg='กรุณากรอกเป็นตัวเลข')
        tip.bind_widget(self.lsl, balloonmsg='กรุณากรอกเป็นตัวเลข')
        
        headers = ['ID', 'MATERIAL', 'METHOD', 'USL', 'TARGET', 'LSL', 'UCL', 'CL', 'LCL', 'CREATED_ON']
        self.table_tree = ttk.Treeview(big_frame, column=headers, show='headings', height=17)
        for header in headers:
            self.table_tree.heading(header, text=header)
            col_width = 60
            if header == 'ID':
                col_width = 20

            if header in ['MATERIAL', 'METHOD', 'CREATED_ON']:
                col_width = 120
            self.table_tree.column(header, anchor='center', width=col_width, minwidth=0)

        self.table_tree.bind('<Double-1>', self.delete_table)
        insert_button = Button(big_frame, text='Add Data', font=self.detail_font, width=10, bg='#ff6699', command=self.put_control_limit)
        display_frame = Frame(big_frame)
        print_button = Button(display_frame, text='Print Report', font=self.detail_font, width=20, bg='#ffcc66', command='')
        get_button = Button(display_frame, text='Display Control Limit', font=self.detail_font, width=20, bg='#ffcc66', command=self.get_control_limit)
        self.check_bool = IntVar()
        delete_cb = Checkbutton(display_frame, text='Delete Data', font=self.detail_font, variable=self.check_bool, onvalue=1, offvalue=0)

        big_frame.pack()
        topics_label.grid(row=0, column=0)
        second_frame.grid(row=1, column=0)
        material_label.grid(row=0, column=0, pady=5)
        self.material.grid(row=0, column=1, pady=5, ipadx=1)
        method_label.grid(row=1, column=0, pady=5)
        self.method.grid(row=1, column=1, pady=5, ipadx=1)
        usl_label.grid(row=2, column=0, pady=5)
        self.usl.grid(row=2, column=1, pady=5)
        target_label.grid(row=3, column=0, pady=5)
        self.target.grid(row=3, column=1, pady=5)
        lsl_label.grid(row=4, column=0, pady=5)
        self.lsl.grid(row=4, column=1, pady=5)
        uclx_label.grid(row=5, column=0, pady=5)
        self.uclx.grid(row=5, column=1, pady=5)
        clx_label.grid(row=6, column=0, pady=5)
        self.clx.grid(row=6, column=1, pady=5)
        lclx_label.grid(row=7, column=0, pady=5)
        self.lclx.grid(row=7, column=1, pady=5)
        
        
        self.table_tree.grid(row=1, column=1, padx=5)
        insert_button.grid(row=2, column=0, pady=10)
        display_frame.grid(row=2, column=1, pady=10)
        get_button.grid(row=0, column=0, padx=5)
        delete_cb.grid(row=0, column=1, padx=5)

    def close_control_limit_window(self):
        self.control_window.destroy()
        self.main_window()

    def put_control_limit(self):
        fill_dict = {'material':self.material.get(), 
                     'method': self.method.get(), 
                     'USL': self.usl.get(),
                     'Target': self.target.get(),
                     'LSL': self.lsl.get(),
                     'UCL': self.uclx.get(),
                     'CL': self.clx.get(),
                     'LCL': self.lclx.get()}
        count_blank = 0
        message = ''
        for fill in fill_dict:
            if not fill_dict[fill]:
                message += (fill + ' ')
                count_blank += 1

        if count_blank == 0:
            try:
                with sqlite3.connect('\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\Database\IQI.DB') as con:
                    query = f"""
                            INSERT INTO control_limits (MATERIAL, METHOD, USL, TARGET, LSL, UCLX, CLX, LCLX) 
                            VALUES ('{self.material.get()}', 
                                    '{self.method.get()}', 
                                    {self.usl.get()}, 
                                    {self.target.get()}, 
                                    {self.lsl.get()}, 
                                    {self.uclx.get()}, 
                                    {self.clx.get()}, 
                                    {self.lclx.get()})
                            """
                    con.execute(query)

                msb.showinfo(title='Information', message='เพิ่มข้อมูลลงในฐานข้อมูลเรียบร้อยแล้ว')
            except Exception:
                msb.showwarning(title='Alarm', message='กรอกข้อมูลไม่ถูกต้อง')
        else:
            msb.showwarning(title='Alarm', message=f'คุณยังไม่กรอกที่ช่อง {message}' )

    def get_control_limit(self):
        try:
            for member in self.table_tree.get_children():
                self.table_tree.delete(member)

            with sqlite3.connect('\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\Database\IQI.DB') as con:
                query = "SELECT * FROM control_limits"
                cursor = con.execute(query)

                for row in con.execute(query):
                    self.table_tree.insert('', 'end', value=row)

        except Exception:
            msb.showinfo(title='Information', message='ไม่มีข้อมูลในฐานข้อมูล')

    def delete_table(self, event):
        is_check = self.check_bool.get()

        if bool(is_check):
            row_select = self.table_tree.selection()
            control_limits_id = str(self.table_tree.item(row_select)['values'][0])
            need_delete = msb.askyesno(title='Ask to User', message=f'คุณต้องการลบข้อมูล ID = {control_limits_id} หรือไม่?')

            if need_delete:
                with sqlite3.connect('\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\Database\IQI.DB') as con:
                    query = f'DELETE FROM control_limits WHERE Id={control_limits_id}'
                    con.execute(query)
                    con.commit()

                msb.showinfo(title='Information', message=f'ทำการลบข้อมูลเรียบร้อยที่ id={control_limits_id}')

    def print_report(self):
        pass

    def import_spc_report(self):
        self.root.withdraw()

        self.spc_window = Toplevel()
        WIDTH = 330
        HEIGHT = 210
        screen_width = self.spc_window.winfo_screenwidth()
        screen_height = self.spc_window.winfo_screenheight()
        center_width = int((screen_width/2) - (WIDTH/2))
        center_height = int((screen_height/2) - (HEIGHT/1.8))

        self.spc_window.title('Control Limit Window')
        self.spc_window.geometry(f'{WIDTH}x{HEIGHT}+{center_width}+{center_height}')
        self.spc_window.protocol('WM_DELETE_WINDOW', self.close_import_spc_report)

    def close_import_spc_report(self):
        self.spc_window.destroy()
        self.main_window()
    


if __name__ == '__main__':
    app = SpcIqi()
    app.main_window()