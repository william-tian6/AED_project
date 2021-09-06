# -*- coding:utf-8 -*-

# MSSQL USER
'''
CREATE LOGIN [socialaed_tp_user] WITH PASSWORD=N'CZIvyP8FWoNwmZ3rM2De5ikx4oW05GWgw9SfxFDGZ6c=', DEFAULT_DATABASE=[socialaed_talent_pool], DEFAULT_LANGUAGE=[繁體中文], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF
GO
'''

# TABLE
'''
USE [socialaed_talent_pool]
GO

DROP TABLE IF EXISTS [dbo].[google_row_data]
CREATE TABLE [dbo].[google_row_data] (
  [id] INT IDENTITY NOT NULL,
  [name] varchar(50),
  [sex] varchar(10),
  [bdate] varchar(50),
  [cellphone] varchar(50),
  [email] varchar(50),
  [industry] varchar(50),
  [organization] varchar(50),
  [title] varchar(50),
  [activities] varchar(50),
  [note] varchar(50),
  [care_issue] varchar(50),
  [specialty] varchar(50),
  [project_excute] varchar(50),
  [wanna_get_social_aed_info] varchar(50),
  [social_aed_mailer_invite_for_project] varchar(50),
  PRIMARY KEY ([id])
)
GO
'''

import pandas as pd
import pymysql
import pymssql
import time, datetime

def get_data(file_name):
    # 用pandas讀取csv
    # data = pd.read_csv(file_name,engine='python',encoding='gbk')
    data = pd.read_csv(file_name,engine='python')
    print (data.head(5)) #列印前5行
    
    # MySQL 資料庫連線
    '''
    conn = pymysql.connect(
        host="localhost",
        port=3306,
        user="socialaed_hr_program",
                passwd="qPC20ew7JQ6Owpr6",
                db="social_aed_hr_source",
                charset = 'utf8mb4'
    )
    '''
    # MSSQL 資料庫連線
    conn = pymssql.connect(
        host='1.34.190.85',
        user='socialaed_tp_user',
        password='VGec6ddCP6PtNXeO',
        database='socialaed_talent_pool'
    )

    # 使用cursor()方法獲取操作遊標
    cursor = conn.cursor(as_dict=True)

    # Clean Data : google_forms_row_data
    sql = 'TRUNCATE TABLE [dbo].[google_row_data]'
    cursor.execute(sql)

    # 資料過濾，替換 nan 值為 None
    data = data.astype(object).where(pd.notnull(data), None) 

    for name,sex,bdate,cellphone,email,industry,organization,title,activities,note,care_issue,specialty,project_excute,wanna_get_social_aed_info,social_aed_mailer_invite_for_project in zip(data['name'],data['sex'],data['bdate'],data['cellphone'],data['email'],data['industry'],data['organization'],data['title'],data['activities'],data['note'],data['care_issue'],data['specialty'],data['project_excute'],data['wanna_get_social_aed_info'],data['social_aed_mailer_invite_for_project']):
    
        # CREATE_DATE = format_date(CREATE_DATE) # 這裡由於對日期有特殊需求，自己處理了一下，程式碼就不貼了，如無需要可略過。
        
        dataList = (name,sex,bdate,cellphone,email,industry,organization,title,activities,note,care_issue,specialty,project_excute,wanna_get_social_aed_info,social_aed_mailer_invite_for_project)

        print (dataList) # 插入的值

        try:
            insertsql = "INSERT INTO [dbo].[google_row_data] (name,sex,bdate,cellphone,email,industry,organization,title,activities,note,care_issue,specialty,project_excute,wanna_get_social_aed_info,social_aed_mailer_invite_for_project) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            cursor.execute(insertsql,dataList)
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
    get_data('/hdd/hdd2/program/socialtp/data/1775946841.csv')


if __name__ == '__main__':
    main()
