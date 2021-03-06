import re, time, DMFlib #import regular expresions, time and DMF library
from yahoo_finance import Share #we are going to need yahoo finance module

date=re.findall('\S*? (\S{3,}) (\d{1,2}) .+?:.+?:.+? (\d{4})',time.ctime()) #get day (0-31), month and year
month,day,year=date[0][0],date[0][1],date[0][2]

print day, month, year
errorlist=[]
while True:
    company=raw_input('Enter company ID -> ')
    if len(company)<1:break

    list=[]
    if company=='lista': #hidden egg command to analyze several companies from a list
        list=DMFlib.GetListFromFile('lista')
    elif company=='52 week lows': #hidden egg command to get 52 week lows from 24/7 wall street
        list=DMFlib.Retrieve52WeekLows(day,DMFlib.WMonthtoNMonth(month))
    elif company=='52 week lows yesterday': #hidden egg command to get 52 week lows from yesterday's 24/7 wall street
        if int(day)>9:list=DMFlib.Retrieve52WeekLows(str(int(day)-1),DMFlib.WMonthtoNMonth(month))
        else:list=DMFlib.Retrieve52WeekLows('0'+str(int(day)-1),DMFlib.WMonthtoNMonth(month))
    else: list.append(company)

    if list==None or len(list)==0:continue #if a list or the 52 week low article was not found... continue

    for company in list:
        try: #some companies fail
            print "Getting",company+'...'
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

            DMFlib.WriteToExcel(company,companyinfo,month,year) #write to DMF excel
        except: #if company fail print error and add it to a list for further documentation
            print 'Error: Could not add',company
            errorlist.append(company)

errorfile=open('ErrorList.txt','w') #write the companies that couldn't be analyzed to a file in plain text
print 'Generating error file...'
for company in errorlist:
    errorfile.write(company+'\n')
print 'All Done.'