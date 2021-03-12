from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msb
import openpyxl as xl
import os


class CrossSection():
    def __int__(self):
        self.link = "\\\\10.17.73.53\\ORT_Result\\03.Data Cross Section\\1.Cross section\\1.For Per Day"
        self.topic_font = ('Arial', 24, 'bold')
        
    def main_window(self):
        root = Tk()
        WIDTH = 500
        HEIGHT = 350
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_width = int((screen_width/2) - (WIDTH/2))
        center_height = int((screen_height/2) - (HEIGHT/2))

        # PROGRAM PROPERTIES
        root.title("CROSS SECTION PROGRAM")
        root.geometry(f'{WIDTH}x{HEIGHT}+{center_width}+{center_height}')
        root.resizable(0, 0)

        # WIDGET PROPERTIES
        big_frame = Frame(root)
        topicsLabel = Label(big_frame, text='CROSS SECTION PROGRAM', font=('Arial', 22), fg='white', bg='blue')
        # topics_label = Label(big_frame, text='Please Select Menu', font=self.topic_font, fg='white', bg='blue')

        # WIDGET POSITION
        big_frame.pack()
        topicsLabel.grid(row=0, column=0, ipadx=5, pady=5)

        # Activate GUI
        root.mainloop()

    def productNameClick(self, event):
        pass


    def lotNoClick(self, event):
        pass
    


app = CrossSection()
app.main_window()