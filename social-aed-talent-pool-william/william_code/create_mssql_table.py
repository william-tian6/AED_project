import pygsheets
import  pymssql
import pandas as pd
import configparser
import json
import time
import  re
from mssql import mssql_insert, mssql_turncate
from google_spreadsheet import google_sheets_name_1, google_spreadsheet_name , google_sheet_title, google_sheet_value_list, sql_create_table

# 檔案：google_spreadsheet
# 頁籤：google_sheets
# ---------------------------------------------------------------------------------
config = configparser.ConfigParser()
config.read('./william_code/config.ini')
survey_url = json.loads(config.get("google","survey_url"))
google_key = config['google']['google_key']
mssql_table = config['mssql']['table']
new_mssql_table = config['mssql']['new_table']
# ----------------------------------------------------------------------------------


if __name__== '__main__':
    print(f'新增{new_mssql_table}開始!')
    create_table_text =sql_create_table(new_mssql_table,google_sheet_title(survey_url[0]))
    print(f'新增{new_mssql_table}完成!')