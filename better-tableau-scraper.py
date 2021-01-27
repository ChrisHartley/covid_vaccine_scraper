from tableauscraper import TableauScraper as TS
import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from datetime import datetime

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Service key is created per https://gspread.readthedocs.io/en/latest/oauth2.html#for-bots-using-service-account
gc = gspread.service_account(filename='./google-key.json')

spreadsheet_key = '14i8B3UeUkE1XbbYO_CtLuiQF6R3cwQxKbYTrT6UlPck'
wks_name = 'Master'

sh = gc.open_by_key(spreadsheet_key)

url = "https://datavizpublic.in.gov/views/Dashboard/Vaccination"

ts = TS()
ts.loads(url)
dashboard = ts.getDashboard()

def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
	# From https://stackoverflow.com/a/45312360/2731298 - Thanks devon.
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

# use this to see all the dataframes that are available
# for t in dashboard.worksheets:
# 	#show worksheet name
# 	print(f"WORKSHEET NAME : {t.name}")
# 	#show dataframe for this worksheet
# 	print(t.data)

data_to_update = [
{"source":"Map (1st Dose)", "sheet": "First Dose Count by County", "dataframe":'SUM(Map Layer First Dose)-alias'},
{"source":"Age", "sheet":"First Dose by Age", "dataframe": 'SUM(Demographics Measure)-alias' },
{"source":"Ethnicity", "sheet":"First Dose by Ethnicity", "dataframe": 'Measure Values-alias'},
{"source":"Gender", "sheet":"First Dose by Gender", "dataframe": 'SUM(Demographics Measure)-value'},
{"source":"Race", "sheet":"First Dose by Race", "dataframe": 'Measure Values-alias'},

]

def main():
	for data_set in data_to_update:

		dashboard = ts.getWorksheet(data_set['source'])
		dataframe = dashboard.data


		worksheet = sh.worksheet(data_set['sheet'])

		# step through the columns finding the first blank one and stick the data there.
		for i in range(1,20):
			value = worksheet.cell(1,i).value
			row = excel_column_name(i)
			if value is None or value == '':
				print('Updating')
				worksheet.update('{}1:{}1'.format(row,row), datetime.now().strftime('%x %X'))
				worksheet.update('{}2:{}93'.format(row,row), dataframe[[data_set['dataframe']]].values.tolist())
				break

if __name__ == "__main__":
    main()
