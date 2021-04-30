import pandas as pd
from os import path, listdir
from shutil import copy2


class SortAncFileSystem():
    def __init__(self):
        self.machine_path = "\\\\10.17.73.53\\ORT_Result\\37.OST"
        self.master_file_path = "\\\\10.17.78.169\\Main_DATA\\00.Record status ORT"
        
    def program_overview(self):
        table_list = []

        for excel_file in listdir(self.master_file_path):
            if (excel_file.upper().startswith('ORT')) and excel_file.upper().endswith('XLSX'):
                excel_file_location = path.join(self.master_file_path, excel_file)
                sheetnames = self.get_all_sheet(path=excel_file_location)
                table_list = self.get_table_list(sheetnames=sheetnames, table_list=table_list, filename=excel_file_location)

                # print(table_list)
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

                row = [barcode, sheetname, product_name, lotno, item_test]
                if 'nan' not in row:
                    table_list.append(row)

        return table_list

    def get_filenames(self, table_list):
        folder_names = listdir(self.machine_path)

        for folder_name in folder_names:
            if folder_name == 'R2-40-131':
                folder_path = path.join(self.machine_path, folder_name)
                for excel_file in listdir(folder_path):
                    barcode = excel_file.split('_')[0]
                    if len(barcode) == 20:
                        excel_file = path.join(self.machine_path, excel_file)
                        self.exist_in_table_list(barcode, table_list, excel_file)

            if folder_name == 'W-40-112':
                folder_path = path.join(self.machine_path, folder_name)
                for excel_file in listdir(folder_path):
                    barcode = excel_file.split('_')[0]
                    if len(barcode) == 20:
                        excel_file = path.join(self.machine_path, excel_file)

            if folder_name == 'ELT':
                folder_path = path.join(self.machine_path, folder_name)
                for excel_file in listdir(folder_path):
                    if 'THA' in excel_file.upper():
                        barcode = self.get_elt_barcode(excel_file)
                        excel_file = path.join(self.machine_path, excel_file)
                        self.exist_in_table_list(barcode, table_list, excel_file)

    def get_elt_barcode(self, excel_file):
        excel_file_list = excel_file.split('-')

        for barcode in excel_file_list:
            if barcode.startswith('THA'):
                return barcode

    def exist_in_table_list(self, barcode, table_list, excel_file):
        items = []

        for row in table_list:
            if (barcode in row) and (row not in table_list):
                print(f'{row}')
                items.append(row)


app = SortAncFileSystem()
app.program_overview()
