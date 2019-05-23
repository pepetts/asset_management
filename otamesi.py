import datetime,openpyxl

today = datetime.date.today()

#ワークブック、シートを取得
wb = openpyxl.load_workbook('analysis.xlsx')
sheet = wb.get_sheet_by_name('monthly_data')


while today.day >= 2:     #日付をエクセルに合わせて1日にする
  today += datetime.timedelta(days=-1)

old_year = today + datetime.timedelta(days=-365)  #1年前の日付を取得
old_str = old_year.strftime('%Y-%m-%d')


#1年前のデータを取得
val = ''
for i in range(2,sheet.max_row):
  cell = sheet.cell(row=i,column=1).value
  cell_str = cell.strftime('%Y-%m-%d')
  if cell_str == old_str:
    val = sheet.cell(row=i,column=2).value

print(val)

