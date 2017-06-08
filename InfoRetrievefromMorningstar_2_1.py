import re, time, DMFlib, argparse #import regular expresions, time and DMF library
from yahoo_finance import Share #we are going to need yahoo finance module

date=re.findall('\S*? (\S{3,}) (\d{1,2}) .+?:.+?:.+? (\d{4})',time.ctime()) #get day (0-31), month and year
month,day,year=date[0][0],date[0][1],date[0][2]

print month, year

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Choose whether to input tickers manually or use a spreadsheet')
    parser.add_argument('-m', '--manual', help='Use this command if you want to input tickers one by one')
    parser.add_argument('-i', '--input', help="Enter this command followed by the excel spreadsheet of tickers")
    args = parser.parse_args()
    if args.input:
        file_path = args.input
        companies = DMFlib.ReadFromExcel(file_path)
        for company in companies:
            companyinfo = DMFlib.GetInfo(company)
            DMFlib.WriteToExcel(company, companyinfo, month, year)
    if args.manual == 'Yes':
        while True:
            company = raw_input("Enter company's ticker: ")
            if len(company) < 1: break
            companyinfo = companyinfo = DMFlib.GetInfo(company)
            DMFlib.WriteToExcel(company, companyinfo, month, year)

