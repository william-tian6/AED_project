import pygsheets
import  pymssql
import pandas as pd
import configparser
import json
import time
from google_spreadsheet import search_mssql_table1, mssql_turncate, mssql_insert, google_sheet, insert_text
# 檔案：google_spreadsheet
# 頁籤：google_sheets
# ---------------------------------------------------------------------------------
config = configparser.ConfigParser()
config.read('./william_code/config.ini')
survey_url = json.loads(config.get("google","survey_url"))
sheet_content_list = json.loads(config.get("table","sheet_table_1"))
google_key = config['google']['google_key']
mssql_table = config['mssql']['table']
# ----------------------------------------------------------------------------------

sql_create_table_title = '''
    CREATE TABLE [dbo].[table_name1] (
    [id] INT IDENTITY NOT NULL,
    [活動]varchar(50),
    [場次]varchar(50),
    [訂單編號]varchar(50),
    [票號]varchar(50),
    [狀態]varchar(50),
    [訂購人姓名]varchar(50),
    [訂購人Email]varchar(50),
    [訂購人電話]varchar(50),
    [參加人姓名]varchar(50),
    [參加人Email]varchar(50),
    [參加人電話]varchar(50),
    [報名時間(GTM+8)]varchar(50),
    [有效時間(GTM+8)]varchar(50),
    [票種分組]varchar(50),
    [票券名稱]varchar(50),
    [票券細節]varchar(50),
    [票價(NT)]varchar(50),
    [付款時間(GTM+8)]varchar(50),
    [付款方式]varchar(50),
    [信用卡末四碼]varchar(50),
    [首次驗票時間(GTM+8)]varchar(50),
    [首次驗票備註]varchar(50),
    [最後驗票時間(GTM+8)]varchar(50),
    [最後驗票備註]varchar(50),
    [驗票次數]varchar(50),
    [驗票通知]varchar(50),
    [備註]varchar(50),
    [取消原因]varchar(50),
    [你是從哪裡得知圓桌早餐的活動呢？]varchar(50),
    [若您選擇公關票，請問配票單位為？]varchar(50),
    [性別]varchar(50),
    [出生年月日]varchar(50),
    [你的行業別是？]varchar(50),
    [你的職位是？]varchar(50),
    [Line帳號]varchar(50),
    [有什麼話想跟我們說？]varchar(50)),
    PRIMARY KEY ([id])
    )'''

if __name__== '__main__':
    # for sheet_content in sheet_content_list:
    #     spreadsheet_name = sheet_content["name"]
    #     spreadsheet_url = sheet_content["url"]
    #     mssql_name = sheet_content["mssql"]
    #     print(mssql_name)
    #     search_mssql_table1('taipei_round_table',spreadsheet_url)
    #     print('清除table資料開始~')
    #     mssql_turncate('fff')
    #     print('清除table資料完成~')
    #     time.sleep(3)
    #     print('寫入資料開始~')
    #     print(mssql_name)
    #     print(insert_text(mssql_name,spreadsheet_url))
    #     oo =[]
    #     # mssql_insert('SET IDENTITY_INSERT fff ON')
    #     for i in insert_text('taipei_round_table',spreadsheet_url):
    #         # oo.append(i)
    #         mssql_insert(i)
    #     print(type(oo))

    mssql_insert("INSERT INTO [dbo].[taipei_round_table]  VALUES('第七場','TRUE', '2106290631131012105970', '2106290631132664958360', '已付款', '鄭寧萱', 'lyseop0105@gmail.com', '886988242438', '鄭寧萱', 'lyseop0105@gmail.com', '886988242438', '2021-06-29 14:31', '2021-07-24 10:00 ~ 2021-07-24 12:00', '', '幹部福利票', '', '0', '', '未選擇', '', '', '', '', '', '0', '', '', '', 'Social AED幹部', '', '生理女', '1994/09/09', '其它', '經營者', 'lydiaballou', '', '')")
