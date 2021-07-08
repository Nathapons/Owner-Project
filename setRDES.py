import cx_Oracle
import os

os.environ['PATH'] = 'C:\Program Files (x86)\Oracle\instantclient_19_10'
# Establish connection to database
dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
conn = cx_Oracle.connect(user='fpc', password='fpc', dsn=dsn_tns)
print('Connected')

cur = conn.cursor()
query = ("query")        
cur.execute(query)

row_no = 0
for row in cur:
    print(row)

conn.close()
print('Complete')