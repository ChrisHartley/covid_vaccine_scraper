from tableauscraper import TableauScraper as TS
import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from datetime import datetime

import argparse

parser = argparse.ArgumentParser(description='Scrape Indiana COVID Vaccine Dashboard and Update Google Sheet')
parser.add_argument('-d', '--dry-run', default=False, action='store_true')
parser.add_argument('-u', '--url', default="https://datavizpublic.in.gov/views/Dashboard/Vaccination", help="URL for show-tables option")
parser.add_argument('-t', '--show-tables', default=False, action="store_true", help="Show available datatables at URL specified")
args = parser.parse_args()


scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Service key is created per https://gspread.readthedocs.io/en/latest/oauth2.html#for-bots-using-service-account
gc = gspread.service_account(filename='./google-key.json')

# You can copy my spreadsheet and then get the new ID from the URL if you want to make your own
spreadsheet_key = '14i8B3UeUkE1XbbYO_CtLuiQF6R3cwQxKbYTrT6UlPck'
wks_name = 'Master'

sh = gc.open_by_key(spreadsheet_key)


ts = TS()


def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
	# From https://stackoverflow.com/a/45312360/2731298 - Thanks devon.
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name


data_to_update = [
# {"url": "https://datavizpublic.in.gov/views/Dashboard/Vaccination", "source":"Map (1st Dose)", "sheet": "First Dose Count by County", "dataframe":'SUM(Map Layer First Dose)-alias'},
# {"url": "https://datavizpublic.in.gov/views/Dashboard/Vaccination", "source":"Age", "sheet":"First Dose by Age", "dataframe": 'SUM(Demographics Measure)-alias' },
# {"url": "https://datavizpublic.in.gov/views/Dashboard/Vaccination", "source":"Ethnicity", "sheet":"First Dose by Ethnicity", "dataframe": 'Measure Values-alias'},
# {"url": "https://datavizpublic.in.gov/views/Dashboard/Vaccination", "source":"Gender", "sheet":"First Dose by Gender", "dataframe": 'SUM(Demographics Measure)-value'},
# {"url": "https://datavizpublic.in.gov/views/Dashboard/Vaccination", "source":"Race", "sheet":"First Dose by Race", "dataframe": 'Measure Values-alias'},
# {
# "url": "https://tableau.bi.iu.edu/t/prd/views/RegenstriefInstituteCOVID-19HospitalizationsandTestsPublicDashboard/RICOVID-19HospitalizationsandTests",
# "source": "Pyramid admissions age value",
# "sheet": "Hospital Admissions by Age",
# "dataframe": "AGG(Headcount Gender Total)-alias"},
# {
# "url": "https://tableau.bi.iu.edu/t/prd/views/RegenstriefInstituteCOVID-19HospitalizationsandTestsPublicDashboard/RICOVID-19HospitalizationsandTests",
# "source": "Pyramid admissions female",
# "sheet": "Hospital Admissions by Age - female",
# "dataframe":"AGG(Headcount Admissions Female (copy))-value",
# },
# {
# "url": "https://tableau.bi.iu.edu/t/prd/views/RegenstriefInstituteCOVID-19HospitalizationsandTestsPublicDashboard/RICOVID-19HospitalizationsandTests",
# "source": "Pyramid admissions male",
# "sheet": "Hospital Admissions by Age - male",
# "dataframe":"AGG(Headcount Admissions Male (copy))-value",
# },

]

def main():

    if args.show_tables:
        # use this to see all the dataframes that are available
        ts.loads(args.url)
        dashboard = ts.getDashboard()
        for t in dashboard.worksheets:
        	#show worksheet name
        	print(f"WORKSHEET NAME : {t.name}")
        	#show dataframe for this worksheet
        	print(t.data)
        exit()

    for data_set in data_to_update:

        ts.loads(data_set["url"])
        dashboard = ts.getDashboard()
        dashboard = ts.getWorksheet(data_set['source'])
        dataframe = dashboard.data

        worksheet = sh.worksheet(data_set['sheet'])
        print('Processing {}'.format(data_set['sheet'],))
        last_update = None
		# step through the columns finding the first blank one and stick the data there.
        for i in range(1,20):
            value = worksheet.cell(1,i).value
            row = excel_column_name(i)

            if value is None or value == '':
                try:
                    # if last_update is not None:
                    #     difference = datetime.now - datetime.strptime(value, '%x %X')
                    #     if difference.days >= 1:
                    #         print('Updated in last 24 hours, skipping')
                    #         break

                    if args.dry_run == False:
                        print('Updating google sheet')
                        worksheet.update('{}1:{}1'.format(row,row), datetime.now().strftime('%x %X'))
                        worksheet.update('{}2:{}93'.format(row,row), dataframe[[data_set['dataframe']]].values.tolist())
                    else:
                        print('Dry run requested, not updating..')
                        print(data_set['sheet'])
                        print(dataframe[[data_set['dataframe']]].values.tolist())
                    break
                except KeyError:
                    print('KeyError on {}, not updated.'.format(data_set['source'],))
            # else:
            #     print('value {}'.format(value,))
            #     try:
            #         last_update = datetime.strptime(value, '%x %X')
            #     except:
            #         print('Unable to parse last update')
if __name__ == "__main__":
    main()
