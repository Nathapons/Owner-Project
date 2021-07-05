import os
import csv

path = "\\\\TA1D180920\\Share file data\\DATA 146,147 11-13 June 21"
complete = "\\\\TA1D180920\\Share file data\\DATA 146,147 11-13 June 21\\COMPLETE"
inspect_machine = input('Please Enter Machine name: ')

for filename in os.listdir(path):
    if filename.upper().endswith('CSV'):
        # เปิดไฟล์
        filename_path = os.path.join(path, filename)
        r = csv.reader(open(filename_path))

        # แก้ไขข้อมูล
        li = list(r)
        li[0] = ['Inspection Machine : {}'.format(inspect_machine)]

        filename_paste = os.path.join(complete, filename)
        if os.path.isfile(filename_paste) == False:
            with open(filename_paste, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(li)


print('Program Complete')