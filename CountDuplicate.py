from collections import Counter
from tkinter import *
from tkinter import messagebox as msb
from tkinter import ttk
import os


class CountDuplicate:
    def ui(self):
        root = Tk()
        WIDTH = 700
        HEIGHT = 500
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = int( (screen_width/2) - (WIDTH/2) )
        y = int( (screen_height/2) - (HEIGHT/1.8))
        root.title('Count Duplicate')
        root.resizable(0, 0)
        root.geometry(f'{WIDTH}x{HEIGHT}+{x}+{y}')

        big_frame = Frame(root)
        topic_label = Label(big_frame, text='Count Duplicate', font=('Arial', 24, 'bold'), fg='white', bg='blue')
        select_frame = Frame(big_frame)
        path_label = Label(select_frame, text='Path:', font=('Arial', 16))
        self.path_box = Entry(select_frame, font=('Arial', 16), justify='center')
        year_label = Label(select_frame, text='Year Report:', font=('Arial', 16))
        self.year_cb = ttk.Combobox(select_frame, justify='center', font=('Arial', 16), width=18)
        month_label = Label(select_frame, text='Month Report:', font=('Arial', 16))
        self.month_cb = ttk.Combobox(select_frame, justify='center', font=('Arial', 16), width=18)
        date_label = Label(select_frame, text='Date Report:', font=('Arial', 16))
        self.date_cb = ttk.Combobox(select_frame, justify='center', font=('Arial', 16), width=18)
        run_button = Button(big_frame, text='Run Program', command=self.programs_work,font=('Arial', 16), width=10)

        treeview_frame = Frame(big_frame)
        headers = ['CSV FILE', 'PIN No', 'RETEST']
        self.count_duplicate_tree = ttk.Treeview(treeview_frame, column=headers, show='headings', height=8)
        vertical_scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.count_duplicate_tree.yview)
        self.count_duplicate_tree.configure(yscrollcommand=vertical_scrollbar.set)

        for header in headers:
            self.count_duplicate_tree.heading(header, text=header)

            col_width = 150
            if header == 'CSV FILE':
                col_width = 300
            elif header == 'RETEST':
                col_width = 80
            self.count_duplicate_tree.column(header, anchor='center', width=col_width, minwidth=0)

        #Create Event Button
        self.year_cb.bind('<Button-1>', self.year_click)
        self.month_cb.bind('<Button-1>', self.month_click)
        self.date_cb.bind('<Button-1>', self.date_click)

        big_frame.pack()
        topic_label.grid(row=0, column=0, pady=10, ipadx=10)
        select_frame.grid(row=1, column=0, ipadx=10)
        path_label.grid(row=0, column=0, pady=5, ipadx=20)
        self.path_box.grid(row=0, column=1, pady=5, ipadx=5)
        year_label.grid(row=1, column=0, pady=5, ipadx=20)
        self.year_cb.grid(row=1, column=1, pady=5, ipadx=10)
        month_label.grid(row=2, column=0, pady=5, ipadx=20)
        self.month_cb.grid(row=2, column=1, pady=5, ipadx=10)
        date_label.grid(row=3, column=0, pady=5, ipadx=20)
        self.date_cb.grid(row=3, column=1, pady=5, ipadx=10)
        run_button.grid(row=2, column=0, ipadx=10)
        treeview_frame.grid(row=3, column=0, ipadx=10)
        self.count_duplicate_tree.grid(row=0, column=0, ipadx=10, pady=10)
        vertical_scrollbar.grid(row=0, column=1, ipady=70)
        root.mainloop()

    def year_click(self, event):
        try:
            link = self.path_box.get()
            self.year_cb['values'] = os.listdir(link)
        except:
            msb.showinfo(title="Information", message="กรุณาใส่ลิงค์ใหม่อีกครั้ง")

    def month_click(self, event):
        link = self.path_box.get()
        year_value = self.year_cb.get()

        if year_value == "":
            msb.showinfo(title="Information", message="กรุณากรอกที่ช่อง Year Report")
        else:
            year_link = os.path.join(link, year_value)
            self.month_cb['values'] = os.listdir(year_link)


    def date_click(self, event):
        link = self.path_box.get()
        year_value = self.year_cb.get()
        month_value = self.month_cb.get()

        if year_value == "" or month_value == "":
            msb.showinfo(title="Information", message="กรุณากรอกที่ช่อง Year Report และ Month Report")
        else:
            month_link = os.path.join(link, year_value) + "\\" + month_value
            self.date_cb['values'] = os.listdir(month_link)

    def programs_work(self):
        link = self.path_box.get()
        year_value = self.year_cb.get()
        month_value = self.month_cb.get()
        date_value = self.date_cb.get()

        if year_value != "" and month_value != "" and date_value != "":
            csv_link = os.path.join(link, year_value) + "\\" + month_value + "\\" + date_value
            csv_file = os.listdir(csv_link)
            new_csv_file = []
            csv_list = []

            for item in csv_file:
                if item.upper().endswith('.CSV'):
                    csv_list.append(item.split("_")[2])
                    new_csv_file.append(item)

            for member in self.count_duplicate_tree.get_children():
                self.count_duplicate_tree.delete(member)

            index = 0
            my_dict = {i:csv_list.count(i) for i in csv_list}
            for key in my_dict:
                pin_no = key.split('.')[0]
                if int(my_dict[key]) > 1:
                    data = [new_csv_file[index], pin_no, my_dict[key]]
                    self.count_duplicate_tree.insert('', 'end', value=data)
                index += 1
        else:
            msb.showinfo(title="Information", message="กรุณากรอกข้อมูลให้ครบ")

app = CountDuplicate()
app.ui()
