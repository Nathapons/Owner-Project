import cx_Oracle
import os

os.environ['PATH'] = 'C:\Program Files (x86)\Oracle\instantclient_19_8'
# Establish connection to database
dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
conn = cx_Oracle.connect(user='fpc', password='fpc', dsn=dsn_tns)
print('Connected')

cur = conn.cursor()
query = ("SELECT T.* FROM FPCC_WORKING_RECORD_TYPE T WHERE T.WRT_TYPE = '00003'")        
cur.execute(query)

row_no = 0
for row in cur:
    print(row)

conn.close()
print('Complete')