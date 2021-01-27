from tableauscraper import TableauScraper as TS

url = "https://datavizpublic.in.gov/views/Dashboard/Vaccination"

ts = TS()
ts.loads(url)
dashboard = ts.getDashboard()

for t in dashboard.worksheets:
	#show worksheet name
#	print(f"WORKSHEET NAME : {t.name}")
	#show dataframe for this worksheet
#	print(t.data)
	pass

#print(dashboard.worksheets)

dashboard = ts.getWorksheet("Map (1st Dose)")
dataframe = dashboard.data
#
# #show selectable columns
# columns = ts.getWorksheet("Map (1st Dose)").getSelectableColumns()
# print(columns)
#
# #show values by column name
# values = ts.getWorksheet("Map (1st Dose)").getValues("County")
# print(values)
#
# dashboard = ts.getWorksheet("Map (1st Dose)").select("County", "Marion")
# for t in dashboard.worksheets:
# # 	#show worksheet name
# 	print(f"WORKSHEET NAME : {t.name}")
# # 	#show dataframe for this worksheet
# # 	print(t.data)
import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from  datetime import datetime

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


gc = gspread.service_account(filename='./google-key.json')

spreadsheet_key = '14i8B3UeUkE1XbbYO_CtLuiQF6R3cwQxKbYTrT6UlPck'
wks_name = 'Master'

sh = gc.open_by_key(spreadsheet_key)

worksheet = sh.worksheet("First Dose Count by County")

def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
	# From https://stackoverflow.com/a/45312360/2731298 - Thanks devon.
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

import string
for i in range(1,20):
	value = worksheet.cell(1,i).value
#	print(i)
#	print('[{}]'.format(value,) )
	#row = string.ascii_uppercase[i-1]
	row = excel_column_name(i)
#	print(row)
	if value is None or value == '':
		print('Updating')
		worksheet.update('{}1:{}1'.format(row,row), datetime.now().strftime('%x %X'))
		worksheet.update('{}2:{}93'.format(row,row), dataframe[['SUM(Map Layer First Dose)-alias']].values.tolist())
		break
