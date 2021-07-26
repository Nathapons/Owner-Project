import cx_Oracle
import os
import openpyxl as xl


class MoveProgram:
    def __init__(self):
        file_config = ""
        os.environ['PATH'] = 'C:\Program Files (x86)\Oracle\instantclient_19_10'
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        self.conn = cx_Oracle.connect(user='fpc', password='fpc', dsn=dsn_tns)

        self.conn.close()

    def get_path(self):
        pass


if __name__ == "__main__":
    app = MoveProgram