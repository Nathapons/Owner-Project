import csv
import datetime
import os


class DataMerge():
    def __init__(self):
        # server = "D:\\Nathapon_KeepFolder\\0.My work\\01.IoT\\08.P'Muse\\SERVER"
        server = "D:\\DB_STATUS_CHECK"
        previous_date = datetime.datetime.today() - datetime.timedelta(days=1)
        previous_format = previous_date.strftime("%Y%m%d")
        previous_normal = previous_date.strftime("%d/%m/%Y")
        previous_location = os.path.join(server, previous_format)

        if os.path.isdir(previous_location):
            result_location = self.create_result_folder(previous_location=previous_location)

            # For File has ACT in name
            act_lists = [file for file in os.listdir(previous_location) if "ACT" in file]
            data_table = self.open_log_file(location=previous_location, file_list=act_lists, previous_normal=previous_normal)
            self.save_csv(location=result_location, csv_name="ACT_RESULT.CSV", data_table=data_table)

            # For File has SBY in name
            sby_lists = [file for file in os.listdir(previous_location) if "SBY" in file]
            data_table = self.open_log_file(location=previous_location, file_list=sby_lists, previous_normal=previous_normal)
            self.save_csv(location=result_location, csv_name="SBY_RESULT.CSV", data_table=data_table)

    def create_result_folder(self, previous_location):
        result_location = os.path.join(previous_location, "RESULT")
        if os.path.isdir(result_location) == False:
            os.mkdir(result_location)

        return result_location

    def open_log_file(self, location, file_list, previous_normal):
        data_table = []
        headers = ['Date', 
                'Time', 
                'Load Average 1m', 
                'Load Average 5m',
                'Load Average 15m',
                "CPU Used",
                "CPU System",
                "Memory Total",
                "Memory Free",
                "Memory Used",
                "SSD Total",
                "SSD Used",
                "SSD Free",
                "SSD %Used",
                ]
        data_table.append(headers)

        for filename in file_list:
            file_location = os.path.join(location, filename)
            log = open(file_location, "r")

            table_row = self.get_log_data(log=log, previous_normal=previous_normal)
            data_table.append(table_row)
            log.close()

        return data_table
        
    def get_log_data(self, log, previous_normal):
        table_row = []
        time = ""
        load_1m = ""
        load_5m = ""
        load_15m = ""
        cpu_used = ""
        cpu_system = ""
        memory_total = ""
        memory_free = ""
        memory_used = ""
        ssd_lists = []

        table_row.append(previous_normal)
        for line in log:
            # For row top-
            if line.startswith("top -"):
                line_split_space = line.split(' ')
                line_split_comma = line.split(',')
                load_1m_text = line_split_comma[3]
                # Get Result from top row
                time = line_split_space[2]
                load_1m = load_1m_text.split(" ")[4]
                load_5m = line_split_comma[4]
                load_15m_text = line_split_comma[5]
                load_15m_list = load_15m_text.split('\n')
                load_15m = str(load_15m_list[0]).strip()

            # For row %CPU
            elif line.startswith("%Cpu(s):"):
                line_split = line.split(',')

                for col in line_split:
                    # Get Data from CPU Used
                    if col.startswith('%Cpu(s):'):
                        cpu_used = col.split(' ')[1]
                        if cpu_used == "":
                            cpu_used = col.split(' ')[2]
                    
                    # Get Data from CPU System
                    if col.endswith('sy'):
                        cpu_system = col.split(' ')[2]

            # For row KiB Mem
            elif line.startswith('KiB Mem'):
                line_split = line.split(',')

                for col in line_split:
                    if col.endswith('total'):
                        memory_total = col.split(' ')[3]

                    if col.endswith('free'):
                        memory_free = col.split(' ')[2]
                        if memory_free == "":
                            memory_free = col.split(' ')[3]

                    if col.endswith('used'):
                        memory_used = col.split(' ')[2]
                        if memory_used == "":
                            memory_used = col.split(' ')[3]

            # For row = /dev/mapper/centos-root xfs
            elif line.startswith('/dev/mapper/centos-root xfs'):
                line_split = line.split(' ')

                for col in line_split:
                    if col.endswith('T'):
                        ssd_lists.append(col)
                    elif col.endswith('%'):
                        ssd_lists.append(col)

        details = [time, load_1m, load_5m, load_15m, cpu_used, cpu_system, memory_total, memory_free, memory_used]
        table_row.extend(details)
        table_row.extend(ssd_lists)
        return table_row

    def save_csv(self, location, csv_name, data_table):
        csv_location = os.path.join(location, csv_name)
        if os.path.isfile(csv_location) == False:
            with open(csv_location, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(data_table)
        else:
          print(f'Folder has {csv_name}')


if __name__ == "__main__":
    app = DataMerge()