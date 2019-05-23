import requests,bs4,openpyxl,datetime
import subprocess


today = datetime.date.today()
year = today.year
manth = today.month


ff_rate_url = 'https://fred.stlouisfed.org/series/FEDFUNDS'   #政策金利
baa_corporate_bond_url = 'https://fred.stlouisfed.org/series/BAA10Y'    #社債スプレッド
dollar_index_url = 'https://fred.stlouisfed.org/series/TWEXMMTH'    #ドル指数
rate_10year_url = 'https://fred.stlouisfed.org/series/DGS10'    #10年債利回り

urls = [ff_rate_url,baa_corporate_bond_url,dollar_index_url,rate_10year_url]
val_f = []

#対象のワークブックとシートを設定
wb = openpyxl.load_workbook('analysis_sub.xlsx')
sheet_score = wb.get_sheet_by_name('score')
sheet_monthly = wb.get_sheet_by_name('monthly_data')

sheet_score.cell(row=2,column=2).value = today

print("指標の値を取得中...")
#各指標の値を取得してくる
n = 0
for url in urls:
  res = requests.get(url)
  html = bs4.BeautifulSoup(res.text,'html.parser')
  vals = html.select('#meta-left-col > div:nth-child(2) > span.series-meta-observation-value')
  val = vals[0].getText()
  val_f.append(float(val))
  
  print(val_f[n])
  n += 1

#直近と1年前の指標データ　宣言と代入
ff_rate = val_f[0]
baa_corporate = val_f[1]
dollar_index = val_f[2]
rate_10year = val_f[3]

ff_rate_old = ""
baa_corporate_old = ""
dollar_index_old = ""
rate_10year_old = ""

#1年前のデータを取得






#sheet_monthlyを更新    直近データを入力
print("直近のデータを書き込み中...")
new_row = sheet_monthly.max_row + 1

sheet_monthly.cell(row=new_row, column=1).value = today
sheet_monthly.cell(row=new_row, column=2).value = ff_rate
sheet_monthly.cell(row=new_row, column=6).value = baa_corporate
sheet_monthly.cell(row=new_row, column=7).value = dollar_index
sheet_monthly.cell(row=new_row, column=4).value = rate_10year


#sheet_scoreを更新
sheet_score.cell(row=5, column=3).value = ff_rate
sheet_score.cell(row=10, column=3).value = baa_corporate
sheet_score.cell(row=14, column=3).value = dollar_index
sheet_score.cell(row=18, column=3).value = rate_10year

sheet_score.cell(row=5, column=4).value = ff_rate_old
sheet_score.cell(row=10, column=4).value = baa_corporate_old
sheet_score.cell(row=14, column=4).value = dollar_index_old
sheet_score.cell(row=18, column=4).value = rate_10year_old


wb.save('analysis.xlsx')
print("完了")
subprocess.call(['open','analysis.xlsx'])
