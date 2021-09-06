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


def google_sheet(sheet_url):
    # google_sheet_key
    gc = pygsheets.authorize(service_account_file=google_key)
    # google_sheet_url
    sh = gc.open_by_url(sheet_url)
    return sh.worksheets()

# 抓檔案第一個分頁名稱
def google_sheets_name_1(sheet_url):
    # google_sheet_key
    gc = pygsheets.authorize(service_account_file=google_key)
    # google_sheet_url
    sh = gc.open_by_url(sheet_url)
    return sh.sheet1

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

# 抓出每個cloumn資料更換成 mssql insert into 的語法，寫成list
def google_sheet_value_list(worksheets,url):
    insert_data_list = []
    for google_sheets_name_list in worksheets:
        google_spreadsheet_name1 = google_spreadsheet_name(url)
        google_sheet_name = regex_sheet_name(google_sheets_name_list)
        google_sheets_df = google_sheets_name_list.get_as_df(start='A1', index_colum=0, empty_value='', include_tailing_empty=False)
        for order_id,ticket_number,payment_status,order_name,order_email,order_cellphone,attendence_name,attendence_email,attendence_cellphone,order_time,ticket_use_time,ticket_group,ticket_name,ticket_detail,ticket_price,payment_time,payment_method,credit_card_last_4_number,first_ticket_check_time,first_ticket_check_note,last_ticket_check_time,last_ticket_check_note,ticket_check_times,ticket_check_notify,note,cancel_reason,from_where,invite_ticket_organization,sex,bdate,industry,title,line_id,any_comment in zip(google_sheets_df['訂單編號'],google_sheets_df['票號'],google_sheets_df['狀態'],google_sheets_df['訂購人姓名'],google_sheets_df['訂購人Email'],google_sheets_df['訂購人電話'],google_sheets_df['參加人姓名'],google_sheets_df['參加人Email'],google_sheets_df['參加人電話'],google_sheets_df['報名時間(GTM+8)'],google_sheets_df['有效時間(GTM+8)'],google_sheets_df['票種分組'],google_sheets_df['票券名稱'],google_sheets_df['票券細節'],google_sheets_df['票價(NT)'],google_sheets_df['付款時間(GTM+8)'],google_sheets_df['付款方式'],google_sheets_df['信用卡末四碼'],google_sheets_df['首次驗票時間(GTM+8)'],google_sheets_df['首次驗票備註'],google_sheets_df['最後驗票時間(GTM+8)'],google_sheets_df['最後驗票備註'],google_sheets_df['驗票次數'],google_sheets_df['驗票通知'],google_sheets_df['備註'],google_sheets_df['取消原因'],google_sheets_df['你是從哪裡得知圓桌早餐的活動呢？'],google_sheets_df['若您選擇公關票，請問配票單位為？'],google_sheets_df['性別'],google_sheets_df['出生年月日'],google_sheets_df['你的行業別是？'],google_sheets_df['你的職位是？'],google_sheets_df['Line帳號'],google_sheets_df['有什麼話想跟我們說？']):
            data_list= (google_spreadsheet_name1,google_sheet_name,order_id,ticket_number,payment_status,order_name,order_email,order_cellphone,attendence_name,attendence_email,attendence_cellphone,order_time,ticket_use_time,ticket_group,ticket_name,ticket_detail,ticket_price,payment_time,payment_method,credit_card_last_4_number,first_ticket_check_time,first_ticket_check_note,last_ticket_check_time,last_ticket_check_note,ticket_check_times,ticket_check_notify,note,cancel_reason,from_where,invite_ticket_organization,sex,bdate,industry,title,line_id,any_comment)
            insert_mssql = f"INSERT INTO [dbo].[{mssql_table}]  VALUES('{google_spreadsheet_name1}','{google_sheet_name}','{order_id}','{ticket_number}','{payment_status}','{order_name}','{order_email}','{order_cellphone}','{attendence_name}','{attendence_email}','{attendence_cellphone}','{order_time}','{ticket_use_time}','{ticket_group}','{ticket_name}','{ticket_detail}','{ticket_price}','{payment_time}','{payment_method}','{credit_card_last_4_number}','{first_ticket_check_time}','{first_ticket_check_note}','{last_ticket_check_time}','{last_ticket_check_note}','{ticket_check_times}','{ticket_check_notify}','{note}','{cancel_reason}','{from_where}','{invite_ticket_organization}','{sex}','{bdate}','{industry}','{title}','{line_id}','{any_comment}')"
            insert_data_list.append(insert_mssql)
            # mssql_insert(insert_mssql)
            # print(data_list)
    return insert_data_list

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

    for table_header_list in google_sheet_title(url):
        # print(table_header_list)
        if table_header_list != "":
            sql_create_table_title=sql_create_table_title+f'  [{table_header_list}] varchar(100),'
        else:
            print('no')
            sql_create_table_title=sql_create_table_title+f'  [Null] varchar(50),'


    return sql_create_table_title+sql_create_table_tail


def google_sheet_value_list1(workshtable_name,url):
    google_sheet_name = google_sheet(url)[1]
    google_sheet_header = google_sheet_name.get_all_values(returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False)
    gg = str(google_sheet_header[0]).replace('[','').replace(']','')#標題

# print(google_sheet_value_list1(google_sheet(survey_url[0])))

def insert_text(table_name,url):
    google_sheet_name = google_sheet(url)[1]
    google_sheet_header = google_sheet_name.get_all_values(returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=False)
    gg = str(google_sheet_header[0]).replace('[','').replace(']','')#標題
    kk =str(google_sheet_header[1]).replace('[','').replace(']','')
    oo = [f"INSERT INTO [dbo].[{table_name}]  VALUES('第七場',{str(google_sheet_header[i+1]).replace('[','').replace(']','')})" for i in range(len(google_sheet_header)-1)]
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
        print(mssql_name+"已存在")

        return
    else:
        print("no")
        mssql_insert(sql_create_table(mssql_name,url))
        print("建表完成")
    # print(next((data for data in all_table_list if data.get('TABLE_NAME')==mssql_name), "找不到table"))
    # 關閉資料庫連線
    conn.commit()
    conn.close()

# google_sheet_name = google_sheet(survey_url[0])[0]
# google_sheet_header = google_sheet_name.get_all_values(returnas='matrix', majdim='ROWS', include_tailing_empty=True, include_tailing_empty_rows=True)
# gg = str(google_sheet_header[0]).replace('[','').replace(']','')#標題
# kk =str(google_sheet_header[1]).replace('[','').replace(']','')
# print(len(google_sheet_header[0]))
# print(len(google_sheet_header[1]))
# print(google_sheet_header[1])
# print(google_sheet_header[0])
# print(sql_create_table('fffff',survey_url[0]))

