import re, time, DMFlib #import regular expresions, time and DMF library
from yahoo_finance import Share #we are going to need yahoo finance module

date=re.findall('\S*? (\S{3,}) (\d{1,2}) .+?:.+?:.+? (\d{4})',time.ctime()) #get day (0-31), month and year
month,day,year=date[0][0],date[0][1],date[0][2]

print month, year
while True:
    company=raw_input('Enter company ID -> ')
    if len(company)<1:break

    companyinfo={}
    dictIS=DMFlib.IncomeStatement(company)
    dictBS=DMFlib.BalanceSheet(company)
    dictKR=DMFlib.KeyRatios(company)
    dictionaries = [dictIS, dictBS, dictKR]
    companyinfo['EXR']=DMFlib.ExchangeRate(dictIS['currency'])

    #price and market capitalization
    share=Share(company)
    companyinfo['quote']=share.get_price()
    companyinfo['MC']=float(share.get_market_cap()[:-1])
    companyinfo['name']=share.get_name()

    for dictionary in dictionaries: #merge dictionaries
        companyinfo.update(dictionary)

    #write to DMF excel
    DMFlib.WriteToExcel(company,companyinfo,month,year)
