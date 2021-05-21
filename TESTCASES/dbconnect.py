import cx_Oracle
import os

os.environ['PATH'] = 'C:\\Users\\pranav4013\\PycharmProjects\\pythonProject1\\venv\\instantclient_19_11'
ip = '192.168.83.135'
port = 1521
service_name = 'kotakdb'
dsn = cx_Oracle.makedsn(ip, port, service_name=service_name)
con = cx_Oracle.connect('kotak_ipg', 'K1o2t3ak_ipg', dsn)
print("connected")

cur = con.cursor()
query = "select * from KPG_MV_IPG_REQ_TXN"
cur.execute(query)
TXN_REF_NO = '160317190258'
for columns in cur:
    if columns[0] == TXN_REF_NO:
        print(columns[0],"  ",columns[1],"  ",columns[2],"  ",columns[3],"  ",columns[4],"  ",columns[5])
        break
cur.close()
con.close()
print("completed")