import pygsheets
gc = pygsheets.authorize(service_file='/hdd/hdd2/program/socialtp/key/credentials.json')

sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1opLKZINvBlTC4PbaJgythblU4IMTZGGTF3AU4jRDT9Y/')

wks_list = sht.worksheets()
print(wks_list)

#選取by順序
wks = sht[0]

#選取by名稱
# wks2 = sht.worksheet_by_title("備忘錄")

#更新名稱
# wks.title = "Google-Human-Source-Data"

#隱藏清單
# wks2.hidden = False

#讀取
A1 = wks.cell('A1')
A1.value
#匯出CSV
wks.export(pygsheets.ExportType.CSV)

# Update
#wks.update_cell('A1', "Hey yank this numpy array")
#wks3.update_cells('A2:A5',[['name1'],['name2'],['name3'],['name4']])

# 清除sheet內所有值 
#wks.clear()
 
# 刪除sheet
#sht.del_worksheet(wks)
