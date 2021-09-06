import pygsheets
import  pymssql
import json
# import pandas as pd
import configparser
import  re
from google_spreadsheet import google_sheet_title

# 檔案：google_spreadsheet
# 頁籤：google_sheets
# ---------------------------------------------------------------------------------
config = configparser.ConfigParser()
config.read('./william_code/config.ini')
survey_url = json.loads(config.get("google","survey_url"))
google_key = config['google']['google_key']
mssql_table = config['mssql']['table']
mssql_host = config['mssql']['host']
mssql_user = config['mssql']['user']
mssql_password = config['mssql']['password']
mssql_database = config['mssql']['database']
# ----------------------------------------------------------------------------------
def mssql_insert(insert_text):
    conn = pymssql.connect(
        host= mssql_host ,
        user= mssql_user ,
        password= mssql_password ,
        database= mssql_database
        )
    cursor = conn.cursor(as_dict=True)

    # Clean Data : google_forms_row_data
    cursor.execute(insert_text)
    print(cursor.fetchall())
    # 關閉資料庫連線
    conn.commit()
    conn.close()

# 資料表欄位清空
def mssql_turncate(db_table):
    conn = pymssql.connect(
        host='122.116.168.75',
        user='william_tien',
        password='william_tien',
        database='william_test_db'
        )
    cursor = conn.cursor(as_dict=True)

    # Clean Data : google_forms_row_data
    sql = f'TRUNCATE TABLE [dbo].[{db_table}]'
    cursor.execute(sql)
    # 關閉資料庫連線
    conn.commit()
    conn.close()

def sql_create_table(new_table_name,url):
    sql_create_table_title = f'''
        DROP TABLE IF EXISTS [dbo].[{new_table_name}]
        CREATE TABLE [dbo].[{new_table_name}] (
        [id] INT IDENTITY NOT NULL,
        [場次] varchar(50),
        '''
    sql_create_table_tail = '''
        PRIMARY KEY ([id])
        )'''

    for table_header_list in google_sheet_title(url):
        sql_create_table_title=sql_create_table_title+f'  [{table_header_list}] varchar(50),'

    return sql_create_table_title+sql_create_table_tail






# print(search_mssql_table1('rrrrr'))