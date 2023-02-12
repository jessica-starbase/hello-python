
import snowflake.connector
import pandas as pdpath = "C:\\Users\\customers_excel.xlsx"
import pandas as pd
import pypyodbcscript = """SELECT * from DB.dbo.orders"""

file = pd.ExcelFile(path)
df = pd.read_excel(file,  sheet_name = 'Sheet1')

connection = pypyodbc.connect("Driver={SQL Server Native Client 11.0}; server=Server_Name; database=DB; uid=MSSQL_UserID; pwd=MSSQL_Passcode")cursor = connection.cursor()
df_sql = pd.read_sql(script, connection )


snowflake.connector.paramstyle='qmark'
ctx = snowflake.connector.connect(
user='snowflake_passcode',
password='snowflake_passcode',
account='snowflake_account'
)
cs = ctx.cursor()
cs.execute(" DELETE FROM schema.python_import; ")

for row in df.to_records(index=False):
cs.execute(" Insert Into schema.python_import (cust_id , state)"
    "VALUES (?, ? ) ", (str(row[0]),str(row[1])) )
print("uploaded")

for row in df_sql.to_records(index=False):
    cs.execute(" Insert Into schema.python_import (order_id, cust_id, type, date, price, qty)"
   "VALUES (?, ?, ?, ?, ?, ?) ", (str(row[0]),str(row[1]),str(row[2]),str(row[3]),str(row[4]),str(row[5])) )
print("uploaded")

