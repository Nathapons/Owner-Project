import pandas as pd
from os import path, listdir, mkdir
from shutil import copy2
from datetime import datetime

class SortAncFileSystem():
    def __init__(self):
        self.master_file_path = "\\\\10.17.78.169\\Main_DATA\\00.Record status ORT"
        self.barcode_error_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\แยก data แล้ว\\BARCODE ERROR'
        
    def program_overview(self):
        table_list = []

        for excel_file in listdir(self.master_file_path):
            if (excel_file.upper().startswith('ORT')) and excel_file.upper().endswith('XLSX'):
                excel_file_location = path.join(self.master_file_path, excel_file)
                sheetnames = self.get_all_sheet(path=excel_file_location)
                table_list = self.get_table_list(sheetnames=sheetnames, table_list=table_list, filename=excel_file_location)

        self.get_filenames(table_list=table_list)
        

    def get_all_sheet(self, path):
        with pd.ExcelFile(io=path) as xl:
            sheetnames = [sheetname for sheetname in xl.sheet_names if sheetname != 'ห้ามลบ']
            xl.close()
        return sheetnames

    def get_table_list(self, filename, sheetnames, table_list):
        found_time = 0
        sheet_found = ''

        for sheetname in sheetnames:
            df = pd.read_excel(filename, sheet_name=sheetname, header=1, dtype=str)
            max_row = list(df.shape)[0]

            for row in range(max_row):
                barcode = str(df['S/N'][row])
                product_name = str(df['P/D name'][row])
                lotno = str(df['lot no. for OST'][row])
                item_test = str(df['Item test'][row])

                row = [barcode, product_name, lotno, item_test]
                if 'nan' not in row:
                    table_list.append(row)

        return table_list

    def get_filenames(self, table_list):
        file_obj = self.create_text_file()

        file_obj.write('Start move file machine: R2-40-131 \n')
        nidech_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\R2-40-131'
        for csv_file in listdir(nidech_path):
            barcode = csv_file.split('_')[0]
            if barcode.startswith('A'):
                csv_file_location = path.join(nidech_path, csv_file)
                self.exist_in_table_list(barcode=barcode, table_list=table_list, csv_file=csv_file_location, file_obj=file_obj)
        
        file_obj.write('End move file machine: R2-40-131 \n\n')

        file_obj.write('Start move file machine: W-40-112 \n')
        nidech_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\W-40-112'
        for csv_file in listdir(nidech_path):
            barcode = csv_file.split('_')[0]
            if barcode.startswith('A'):
                csv_file_location = path.join(nidech_path, csv_file)
                self.exist_in_table_list(barcode=barcode, table_list=table_list, csv_file=csv_file_location, file_obj=file_obj)
        file_obj.write('End move file machine: W-40-112 \n\n')

    def create_text_file(self):
        now = datetime.now()
        now_filename = now.strftime("%Y_%m_%d_%H_%M") + ".txt"
        log_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\แยก data แล้ว\\ERROR REPORT'
        now_filename = path.join(log_path, now_filename)
        file_obj = open(now_filename, 'a')
        
        return file_obj

    def get_elt_barcode(self, excel_file):
        excel_file_list = excel_file.split('-')

        for barcode in excel_file_list:
            if barcode.startswith('THA'):
                return barcode

    def exist_in_table_list(self, barcode, table_list, csv_file, file_obj):
        items = []

        for row in table_list:
            if (barcode in row) and (row not in items):
                items.append(row)

        if len(items) >= 2:
            print(items)
        # if len(items) >= 2:
        #     link = '\\\\10.17.73.53\\ORT_Result\\37.OST\\แยก data แล้ว\\ITEM TEST ERROR'
        #     self.copy_to_error(csv_file=csv_file, link=link, file_obj=file_obj)
        # elif len(items) == 1:
        #     barcode_details = items[0]
        #     self.copy_to_folder(barcode_details=barcode_details, csv_file=csv_file)

    def copy_to_error(self, link, csv_file, file_obj):
        csv_file_name = path.basename(csv_file)
        error_file_list = listdir(link)
        destination = path.join(link, csv_file_name)

        if csv_file not in error_file_list:
            file_obj.write(f'   Move to ITEM TEST ERROR: {csv_file_name}\n')
            copy2(src=csv_file,dst=destination)

    def copy_to_folder(self, barcode_details, csv_file):
        master_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\แยก data แล้ว'
        product_name = barcode_details[1]
        if '-' not in product_name:
            if 'Z' in product_name.upper():
                product = product_name[0:4] + '-' + product_name[4:]
            else:
                product = product_name[0:3] + '-' + product_name[3:]
        else:
            product = product_name
        lotno = barcode_details[2]
        item_test = barcode_details[3]

        # create folder
        product_path = self.create_folder(link=master_path, folder_name=product)
        item_path = self.create_folder(link=product_path, folder_name=item_test)
        lotno_path = self.create_folder(link=item_path, folder_name=lotno)

        csv_file_list = listdir(lotno_path)
        if csv_file not in csv_file_list:
            csv_filename = path.basename(csv_file)
            destination = path.join(lotno_path, csv_filename)
            copy2(src=csv_file, dst=destination)


    def create_folder(self, link, folder_name):
        folder_list = listdir(link)
        folder_path = path.join(link, folder_name)

        if folder_name not in folder_list:
            mkdir(folder_path)

        return folder_path


app = SortAncFileSystem()
app.program_overview()