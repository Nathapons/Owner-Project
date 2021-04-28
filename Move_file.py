from os import path, listdir, mkdir
from shutil import copy2
from time import ctime

class MoveFile():
    def __init__(self):
        self.source = '\\\\10.17.78.169\\ost'
        self.destination = '\\\\10.17.73.53\\ORT_Result\\37.OST'

    def get_all_file(self):
        try:
            machine_folders = listdir(self.source)
            
            for machine_folder in machine_folders:
                machine_source_path = path.join(self.source, machine_folder)
                machine_destina_path = path.join(self.destination, machine_folder)
                machine_destina_exist = path.isdir(machine_destina_path)
                
                if machine_destina_exist == False:
                    mkdir(machine_destina_path)
                
                csv_source_lists = set(csv_file for csv_file in listdir(machine_source_path) if csv_file.upper().endswith('CSV'))
                csv_destination_lists = set(csv_file for csv_file in listdir(machine_destina_path) if csv_file.upper().endswith('CSV'))
                csv_files = list(csv_source_lists - csv_destination_lists)

                print(f'------- Machine Import: {machine_folder} -------')
                if len(csv_files) > 0:
                    for csv_file in csv_files:
                        print(f'Move file: {csv_file}')
                        csv_file_source = path.join(machine_source_path, csv_file)
                        csv_file_destination = path.join(machine_destina_path, csv_file)
                        # Copy to ANC Server
                        copy2(csv_file_source, csv_file_destination)
        
        except Exception:
            print('Move_file program is error!!')

            
            
if __name__ == '__main__':
    app = MoveFile()
    app.get_all_file()