from tkinter import *
from tkinter import ttk
import os
import openpyxl as xl
import sqlite3


class SpcProgram:
    def __init__(self):
        self.create_spec_database()
        self.main_ui()

    def create_spec_database(self):
        error_text = ""
        try:
            database_path = "\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\\Database\\IQI.DB"
            with sqlite3.connect(database_path) as conn:
                sql_cmd = """
                    CREATE TABLE MATERIAL (
                           MATERIAL_ID INTEGER PRIMARY KEY AUTOINCREMENT
                        ,  NAME TEXT
                        ,  CREATED_ON DATE_TIME DEFAULT CURRENT_TIMESTAMP
                    )
                """
                conn.execute(sql_cmd)

                sql_cmd = """
                    CREATE TABLE METHOD (
                           ID INTEGER PRIMARY KEY AUTOINCREMENT
                        ,  MATERIAL_ID INTEGER NOT NULL
                        ,  NAME TEXT
                        ,  CREATED_ON DATE_TIME DEFAULT CURRENT_TIMESTAMP
                        ,  FOREIGN KEY (MATERIAL_ID) REFERENCES MATERIAL(MATERIAL_ID)
                    )
                """
                conn.execute(sql_cmd)

                sql_cmd = """
                    CREATE TABLE SPC_LINE (
                           ID INTEGER PRIMARY KEY AUTOINCREMENT
                        ,  MATERIAL_ID INTEGER NOT NULL
                        ,  METHOD_ID INTEGER PRIMARY KEY
                        ,  USL FLOAT
                        ,  TARGET FLOAT
                        ,  LSL FLOAT
                        ,  UCL1 FLOAT
                        ,  CL1 FLOAT
                        ,  LCL1 FLOAT
                        ,  UCL2 FLOAT
                        ,  CL2 FLOAT
                        ,  LCL2 FLOAT
                        ,  CREATED_ON DATE_TIME DEFAULT CURRENT_TIMESTAMP
                        ,  FOREIGN KEY (MATERIAL_ID) REFERENCES MATERIAL(MATERIAL_ID)
                        ,  FOREIGN KEY (METHOD_ID) REFERENCES MATERIAL(METHOD_ID)
                    )
                """
                conn.execute(sql_cmd)

                conn.close()
        except Exception:
            error_text = "DB has been created!!"
        return error_text

    def main_ui(self):
        self.root = Tk()
        WIDTH = 400
        HEIGHT = 250
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        top = int((screen_height/2) - (HEIGHT/2))
        left = int((screen_width/2) - (WIDTH/2))
        self.root.geometry(f'{WIDTH}x{HEIGHT}+{left}+{top}')
        self.root.title('MAIN PROGRAM')
        self.root.resizable(0, 0)
        self.topic_font = ('Arial', 28, 'bold')
        self.detail_font = ('Arial', 20)
        big_frame = Frame(self.root)
        big_frame.pack()
        topic = Label(big_frame, text='SPC PROGRAM', font=self.topic_font, fg='white', bg='blue')
        topic.grid(row=0, column=0, ipadx=10)
        inform = Label(big_frame, text='Please select report', font=self.detail_font)
        inform.grid(row=1, column=0, pady=5)
        spec_btn = Button(big_frame, text='Material Spec Record', font=self.detail_font, width=20, command="")
        spec_btn.grid(row=2, column=0, pady=5)
        report_btn = Button(big_frame, text='Spc Report', font=self.detail_font, width=20, command="")
        report_btn.grid(row=3, column=0, pady=5)
        self.root.mainloop()

    def spec_ui(self):
        self.root.withdraw()
        self.spec_window = Toplevel()
    
    def spec_program(self):
        pass

    def report_ui(self):
        pass

    def report_program(self):
        pass


if __name__ == '__main__':
    app = SpcProgram()