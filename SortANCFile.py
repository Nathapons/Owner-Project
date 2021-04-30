import pandas as pd
from os import path, listdir
from shutil import copy2


class UpdateTestReport():
    def __init__(self):
        self.nidech_file_path = "\\\\10.17.73.53\\ORT_Result\\37.OST\\R2-40-131"
        self.yamaha_file_path = "\\\\10.17.73.53\\ORT_Result\\37.OST\\W-40-112"
        self.master_file_path = "\\\\10.17.78.169\\Main_DATA\\00.Record status ORT"
        self.item_test_error_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\แยก data แล้ว\\ITEM TEST ERROR'
        self.barcode_error_path = '\\\\10.17.73.53\\ORT_Result\\37.OST\\แยก data แล้ว\\BARCODE ERROR'

    def sort_anc_program(self):
        for excel_file in listdir(self.master_file_path):
            if 'ORT' in excel_file.upper():
                master_file_location = path.join(self.master_file_path, excel_file)
                sheet_names = self.get_all_sheets()
                table_list = self.get_table_list(sheet_names=sheet_names)
        self.get_filenames(table_list=table_list)

    def get_all_sheets(self, master_file_location):
        with pd.ExcelFile(master_file_location) as xl:
            sheet_names = xl.sheet_names
            del sheet_names[len(sheet_names) - 1]
            xl.close()
        return sheet_names

    def get_table_list(self, sheet_names):
        table_list = []
        found_time = 0
        sheet_found = ''

        for sheet_name in sheet_names:
            df = pd.read_excel(self.master_file_location, sheet_name=sheet_name, header=1, dtype=str)
            max_row = list(df.shape)[0]

            for row in range(max_row):
                barcode = str(df['S/N'][row])
                product_name = str(df['P/D name'][row])
                lotno = str(df['lot no. for OST'][row])
                item_test = str(df['Item test'][row])

                row = [barcode, sheet_name, product_name, lotno, item_test]
                if 'nan' not in row:
                    table_list.append(row)

        return table_list

    def get_filenames(self, table_list):
        print('R2-40-131')
        for excel_file in listdir(self.nidech_file_path):
            if "_" in excel_file:
                barcode = excel_file.split("_")[0]
                if len(barcode) == 20:
                    self.exist_in_table_list(barcode, table_list, excel_file)

        # print('W-40-112')
        # for excel_file in listdir(self.yamaha_file_path):
        #     if "_" in excel_file:
        #         barcode = excel_file.split("_")[0]
        #         self.exist_in_table_list(barcode, table_list, excel_file)
        
    def exist_in_table_list(self, barcode, table_list, excel_file):
        item_list = []

        for row in table_list:
            barcode_in_row = row[0]
            item_test = row[4]
            if (barcode in barcode_in_row) and (row not in item_list):
                item_list.append(row)

        if len(item_list) != 0:
            for item in item_list:
                print(f'{excel_file}/ {barcode}/ {item[0]}/ {item[1]} / {item[2]} / {item[3]} / {item[4]}')

    
if __name__ == '__main__':
    app = UpdateTestReport()
    app.sort_anc_program()