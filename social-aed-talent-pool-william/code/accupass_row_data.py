# -*- coding:utf-8 -*-

# MySQL USER
'''
CREATE LOGIN [socialaed_tp_user] WITH PASSWORD=N'CZIvyP8FWoNwmZ3rM2De5ikx4oW05GWgw9SfxFDGZ6c=', DEFAULT_DATABASE=[socialaed_talent_pool], DEFAULT_LANGUAGE=[繁體中文], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF
GO
'''

# TABLE
'''
USE [socialaed_talent_pool]
GO

DROP TABLE IF EXISTS [dbo].[accupass_row_data]
 CREATE TABLE [dbo].[accupass_row_data] (
  [id] INT IDENTITY NOT NULL,
  [order_id] VARCHAR(50),
  [ticket_number] varchar(50),
  [payment_status] varchar(50),
  [order_name] varchar(50),
  [order_email] varchar(50),
  [order_cellphone] varchar(50),
  [attendence_name] varchar(50),
  [attendence_email] varchar(50),
  [attendence_cellphone] varchar(50),
  [order_time] varchar(50),
  [ticket_use_time] varchar(50),
  [ticket_group] varchar(50),
  [ticket_name] varchar(100),
  [ticket_detail] varchar(50),
  [ticket_price] varchar(50),
  [payment_time] varchar(50),
  [payment_method] varchar(50),
  [credit_card_last_4_number] varchar(10),
  [first_ticket_check_time] varchar(50),
  [first_ticket_check_note] varchar(50),
  [last_ticket_check_time] varchar(50),
  [last_ticket_check_note] varchar(50),
  [ticket_check_times] varchar(50),
  [note] varchar(50),
  [cancel_reason] varchar(50),
  [from_where] varchar(50),
  [invite_ticket_organization] varchar(50),
  [sex] varchar(10),
  [cellphone] varchar(50),
  [email] varchar(50),
  [industry] varchar(50),
  [title] varchar(50),
  [line_id] varchar(50),
  [any_comment] varchar(50),
  PRIMARY KEY ([id])
)
GO
'''

import pandas as pd
import pymssql
import time, datetime

def get_data(file_name):
    # 用pandas讀取csv
    # data = pd.read_csv(file_name,engine='python',encoding='gbk')
    data = pd.read_csv(file_name,engine='python')
    print (data.head(5)) #列印前5行
    
    # 資料庫連線
    conn = pymssql.connect(
        host='1.34.190.85',
        user='socialaed_tp_user',
        password='VGec6ddCP6PtNXeO',
        database='socialaed_talent_pool'
    )

    # 使用cursor()方法獲取操作遊標
    cursor = conn.cursor(as_dict=True)

    # Clean Data : accupass_row_data
    sql = 'TRUNCATE TABLE [dbo].[accupass_row_data]'
    cursor.execute(sql)
    
    # 資料過濾，替換 nan 值為 None
    data = data.astype(object).where(pd.notnull(data), None) 

    for order_id,ticket_number,payment_status,order_name,order_email,order_cellphone,attendence_name,attendence_email,attendence_cellphone,order_time,ticket_use_time,ticket_group,ticket_name,ticket_detail,ticket_price,payment_time,payment_method,credit_card_last_4_number,first_ticket_check_time,first_ticket_check_note,last_ticket_check_time,last_ticket_check_note,ticket_check_times,note,cancel_reason,from_where,invite_ticket_organization,sex,cellphone,email,industry,title,line_id,any_comment in zip(data['訂單編號'],data['票號'],data['狀態'],data['訂購人姓名'],data['訂購人Email'],data['訂購人電話'],data['參加人姓名'],data['參加人Email'],data['參加人電話'],data['報名時間(GTM+8)'],data['有效時間(GTM+8)'],data['票種分組'],data['票券名稱'],data['票券細節'],data['票價(NT)'],data['付款時間(GTM+8)'],data['付款方式'],data['信用卡末四碼'],data['首次驗票時間(GTM+8)'],data['首次驗票備註'],data['最後驗票時間(GTM+8)'],data['最後驗票備註'],data['驗票次數'],data['備註'],data['取消原因'],data['你是從哪裡得知圓桌早餐的活動呢？'],data['若您選擇公關票，請問配票單位為？'],data['性別'],data['聯絡電話'],data['電子信箱（仔細檢查不要填錯唷！）'],data['你的行業別是？'],data['你的職位是？'],data['Line帳號'],data['有什麼話想跟我們說？']):
    
        # CREATE_DATE = format_date(CREATE_DATE) # 這裡由於對日期有特殊需求，自己處理了一下，程式碼就不貼了，如無需要可略過。
        
        dataList = (order_id,ticket_number,payment_status,order_name,order_email,order_cellphone,attendence_name,attendence_email,attendence_cellphone,order_time,ticket_use_time,ticket_group,ticket_name,ticket_detail,ticket_price,payment_time,payment_method,credit_card_last_4_number,first_ticket_check_time,first_ticket_check_note,last_ticket_check_time,last_ticket_check_note,ticket_check_times,note,cancel_reason,from_where,invite_ticket_organization,sex,cellphone,email,industry,title,line_id,any_comment)

        print (dataList) # 插入的值

        try:
            insertsql = "INSERT INTO [dbo].[accupass_row_data] (order_id,ticket_number,payment_status,order_name,order_email,order_cellphone,attendence_name,attendence_email,attendence_cellphone,order_time,ticket_use_time,ticket_group,ticket_name,ticket_detail,ticket_price,payment_time,payment_method,credit_card_last_4_number,first_ticket_check_time,first_ticket_check_note,last_ticket_check_time,last_ticket_check_note,ticket_check_times,note,cancel_reason,from_where,invite_ticket_organization,sex,cellphone,email,industry,title,line_id,any_comment) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            cursor.execute(insertsql, dataList)
            conn.commit()
        except Exception as e:
            print ("Exception")
            print (e)
            conn.rollback()
    cursor.close()
    # 關閉資料庫連線
    conn.commit()
    conn.close()

def main():
    # 讀取資料
    get_data('/hdd/hdd2/program/socialtp_mssql/data/accupass.csv')


if __name__ == '__main__':
    main()
