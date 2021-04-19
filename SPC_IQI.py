from openpyxl import *
from os import path, listdir
from datetime import datetime


class SpcIqi():
    def __init__(self):
        self.iqi_server = '\\\\ta1d170506\\IQI ONLY\\21.Upload to system'
        self.spc_server = '\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\\Format'
        self.current_year = str(now.year)

    def get_each_result_location(self):
        now = datetime.now()
        material_folders = listdir(self.iqi_server)

        for material_folder in material_folders:
            if 'SPC' not in material_folder:
                material_folder_path = path.join(self.iqi_server, material_folder)
                if path.isdir(material_folder_path):
                    year_folders = listdir(material_folder_path)

                    for year_folder in year_folders:
                        if year_folder == self.current_year:
                            year_folder_path = path.join(material_folder_path, year_folder)
                            item_code_files = listdir(year_folder_path)

                            for item_code_file in item_code_files:
                                if item_code_file.endswith('.xlsx'):
                                    item_code_path = path.join(year_folder, item_code_file)
                                    # Run function to open Excel
                                    item_code_wb = self.open_item_code_path(excel_path=item_code_path)
                                    self.close_item_code_wb(wb=item_code_wb)


    def open_item_code_path(self, excel_path):
        wb = load_workbook(filename=excel_path, read_only=True)
        return wb

    def close_item_code_wb(self, wb):
        wb.close()





if __name__ == '__main__':
    app = SpcIqi()
    app.get_each_result_location()