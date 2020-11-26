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
        HEIGHT = 80
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = int((screen_width/2) - (WIDTH/2))
        y = int((screen_height/2) - (HEIGHT/1.6))

        # UI Properties
        self.window.title("Etching Rate Import Program")
        self.window.geometry(f'{WIDTH}x{HEIGHT}+{x}+{y}')
        self.window.attributes('-disabled', True)
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

        self.window.after(1000, self.etching_overview)
        # Run after 6hrs
        # self.window.after(14400, self.etching_overview)


app = EtchingRate()
app.etching_window()