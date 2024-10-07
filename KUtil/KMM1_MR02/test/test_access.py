import win32com.client
import pyodbc
import adodbapi
from KUtil.KMM1_MR02.Class import CAccessDB as accdb

db = r"C:\zDsk\github\Koin\KUtil\KMM1_MR02\_Data\koinWDB1.accdb;"

#conn = win32com.client.Dispatch(r'ADODB.Connection')
#DSN = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + db
#mdb = conn.Open(DSN)

conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r"C:\zDsk\github\Koin\KUtil\KMM1_MR02\_Data\koinWDB1.accdb;")
conn = pyodbc.connect(conn_str, autocommit=True)

tab_names = mdb.getTableNames()
tabs = {}
for tab_name in tab_names:
    tabs[tab_name] = mdb.Table(conn, tab_name)
    print(tab_name)


adb = accdb.AccessDb
conn = adb.getConn

tnames =  adb.getTables

print(tnames)
#tb1 = accdb.Table(conn, tab_name="tcPrjMat1")