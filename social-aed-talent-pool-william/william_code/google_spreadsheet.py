import pygsheets
import  pymssql
import pandas as pd
import configparser
import json
import  re
# from mssql import mssql_insert, mssql_user, mssql_host, mssql_password, mssql_database

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
# 抓檔案每個分頁名稱
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
    # print(cursor.fetchall())
    # 關閉資料庫連線
    conn.commit()
    conn.close()

# 資料表欄位清空
def mssql_turncate(db_table):
    conn = pymssql.connect(
        host='122.116.168.75',
        user='socialaed_tp_user',
        password='VGec6ddCP6PtNXeO',
        database='socialaed_talent_pool'
        )
    cursor = conn.cursor(as_dict=True)

    # Clean Data : google_forms_row_data
    sql = f'TRUNCATE TABLE [dbo].[{db_table}]'
    cursor.execute(sql)
    # 關閉資料庫連線
    conn.commit()
    conn.close()


def google_sheet(sheet_url):
    # google_sheet_key
    gc = pygsheets.authorize(service_account_file=google_key)
    # google_sheet_url
    sh = gc.open_by_url(sheet_url)
    return sh.worksheets()

# 抓檔案名稱
def google_spreadsheet_name(sheet_url):
    # google_spreadsheet_key
    gc = pygsheets.authorize(service_account_file= google_key)
    # google_spreadsheet_url
    sh = gc.open_by_url(sheet_url)
    return sh.title

# 抓每個column的標題
def google_sheet_title(sheet_url):
    gc = pygsheets.authorize(service_account_file= google_key)
    sh = gc.open_by_url(sheet_url)
    sheet1_title_all = sh.sheet1.get_all_values(returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False)[0]
    return [sheet1_title_list for sheet1_title_list in sheet1_title_all]


# 正則sheet分頁名稱
def regex_sheet_name(google_sheets_name):
    regex = re.compile(r"'(.+)'")   
    match = regex.search(str(google_sheets_name))
    google_sheets_name = match.group(1)
    return google_sheets_name


# create new table table_header=新的google_spreadsheet裡面的第一分頁cloumn標題 new_table_name=google_spreadsheet檔案名稱
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
    x=0
    for table_header_list in google_sheet_title(url):
        # print(table_header_list)
        
        if table_header_list != "":
            sql_create_table_title=sql_create_table_title+f'  [{table_header_list}] varchar(200),'
        else:
            x=x+1
            text ='Null'+str(x)
            # print(text)
            sql_create_table_title=sql_create_table_title+f'  [{text}] varchar(10),'
    # print(sql_create_table_title+sql_create_table_tail)

    return sql_create_table_title+sql_create_table_tail


def insert_text(table_name,url,sheet_name,index):
    google_sheet_name = google_sheet(url)[index]
    google_sheet_header = google_sheet_name.get_all_values(returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False)
    gg = str(google_sheet_header[0]).replace('[','').replace(']','')#標題
    kk =str(google_sheet_header[1]).replace('[','').replace(']','')
    oo = [f"INSERT INTO [dbo].[{table_name}]  VALUES('{sheet_name}',{str(google_sheet_header[i+1]).replace('[','').replace(']','')})" for i in range(len(google_sheet_header)-1)]
    return oo

def search_mssql_table1(mssql_name,url):
    conn = pymssql.connect(
        host= mssql_host ,
        user= mssql_user ,
        password= mssql_password ,
        database= mssql_database
        )
    cursor = conn.cursor(as_dict=True)

    # Clean Data : google_forms_row_data
    cursor.execute('SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME')
    all_table_list = cursor.fetchall()
    # print(type(all_table_list))
    
    if mssql_name in [all_table["TABLE_NAME"] for all_table in all_table_list]:
        print(mssql_name+"已存在!")

        pass
    else:
        print("開始建表")
        mssql_insert(sql_create_table(mssql_name,url))
        print("建表完成")
    # print(next((data for data in all_table_list if data.get('TABLE_NAME')==mssql_name), "找不到table"))
    # 關閉資料庫連線
    conn.commit()
    conn.close()
