from openpyxl import *
from os import path, listdir
from datetime import datetime


class SpcIqi():
    def __init__(self):
        self.iqi_server = '\\\\ta1d170506\\IQI ONLY\\21.Upload to system'
        self.spc_server = '\\\\ta1d170506\\IQI ONLY\\21.Upload to system\\0.SPC\\Format'

    def get_each_result_location(self):
        now = datetime.now()
        current_year = str(now.year)
        material_folders = listdir(self.iqi_server)

        for material_folder in material_folders:
            if 'SPC' not in material_folder:
                material_folder_path = path.join(self.iqi_server, material_folder)
                if path.isdir(material_folder_path):
                    year_folders = listdir(material_folder_path)

                    for year_folder in year_folders:
                        if year_folder == current_year:
                            print(year_folder, type(year_folder))
                            year_folder_path = path.join(material_folder_path, year_folder)
                            item_code_folders = listdir(year_folder_path)





if __name__ == '__main__':
    app = SpcIqi()
    app.get_each_result_location()