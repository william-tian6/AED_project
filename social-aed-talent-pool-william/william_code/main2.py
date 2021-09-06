from numpy.lib.function_base import angle
import pygsheets
import  pymssql
import pandas as pd
import configparser
import json
import time
import  re
from mssql import mssql_insert, mssql_turncate
from google_spreadsheet import google_sheet, google_sheets_name_1, google_spreadsheet_name ,google_sheet_title,google_sheet_value_list

# 檔案：google_spreadsheet
# 頁籤：google_sheets
# ---------------------------------------------------------------------------------
config = configparser.ConfigParser()
config.read('./william_code/config.ini')
survey_url = json.loads(config.get("google","survey_url"))
google_key = config['google']['google_key']
mssql_table = config['mssql']['table']

# config_1 = configparser.ConfigParser()
# config_1.read('./william_code/config_1.ini')
# table_name = config_1[]
# ----------------------------------------------------------------------------------
sheet_table = {'table_1':{'name':'台北圓桌參加名單','url':'https://docs.google.com/spreadsheets/d/1QM6T8KPcmjOPJx8xCffNqObWO0gprJanozaKXBClX3U/edit#gid=133511107','mssql_name':''},
            'table_2':{'name':'桃園圓桌參加名單','url':'https://docs.google.com/spreadsheets/d/1Wnmb_Dr_C1cjYrRoJdsQLncnAVQTO4c50f_EMNf6INM/edit#gid=319926160','mssql_name':''}, 
            'table_3':{'name':'巧思洞察局參加名單','url':'https://docs.google.com/spreadsheets/d/1RQV4bUlMDOJNa1jwIpLxoysW38IAxyJXuLzOFgDZ3Jo/edit#gid=570843625','mssql_name':''},
            'table_4':{'name':'生命探索工作坊參加名單','url':'https://docs.google.com/spreadsheets/d/1SVt8-uBWa1-2G8XV7Tw2KyGD0f3zy8Bue5UPYgtXNtE/edit#gid=0','mssql_name':''},
            'table_5':{'name':'創新思維工作坊參加名單','url':'https://docs.google.com/spreadsheets/d/1UWSuwJ_gewXbljM1XKumFJhg1k8vKLFnGcWiP72mDjg/edit#gid=0','mssql_name':''}, 
            'table_5':{'name':'點子微醺夜參加名單','url':'https://docs.google.com/spreadsheets/d/18MFRSCYm1yQrR9YusGYAFHB-X7_cqVeWKS0wbAQM2iY/edit#gid=1906600855','mssql_name':''}, 
}
data = [{'name':'台北圓桌參加名單','url':'https://docs.google.com/spreadsheets/d/1QM6T8KPcmjOPJx8xCffNqObWO0gprJanozaKXBClX3U/edit#gid=133511107','mssql_name':''},{'name':'桃園圓桌參加名單','url':'https://docs.google.com/spreadsheets/d/1Wnmb_Dr_C1cjYrRoJdsQLncnAVQTO4c50f_EMNf6INM/edit#gid=319926160','mssql_name':''}]
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
def test(g,h,j):
    print(g)
    print(h)
    print(j)

# if __name__== '__main__':
#     # print(sheet_table.keys())#取表順序
#     # print((sheet_table.values()))
#     # print('清除table資料開始~')
#     # mssql_turncate(mssql_table)
#     # print('清除table資料完成~')
#     # time.sleep(3)
#     # print('寫入資料開始~')
#     # for i in survey_url:
#     #     for data_list in google_sheet_value_list(google_sheet(i),i):
#     #         mssql_insert(data_list)
#    f # print('完成!')



otter = [

]

test_url = json.loads(config.get("table","sheet_table_1"))

sheet_table_1 = [('台北圓桌參加名單','https://docs.google.com/spreadsheets/d/1QM6T8KPcmjOPJx8xCffNqObWO0gprJanozaKXBClX3U/edit#gid=133511107','taipei_round_table'),('桃園圓桌參加名單','https://docs.google.com/spreadsheets/d/1Wnmb_Dr_C1cjYrRoJdsQLncnAVQTO4c50f_EMNf6INM/edit#gid=319926160','taoyuan_round_table')]
for ii in test_url:
    print(ii["name"])
    print(ii['mssql'])
print('-----------------------------')
print(test_url)
print('-----------------------------')
print(test_url[1]["name"])
print(test_url[0]["name"])
print(test_url[1]["url"])
print('-----------------------------')
print(type(test_url[1]["name"]))
print('-----------------------------')