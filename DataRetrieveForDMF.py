import urllib, re, json, openpyxl, time #import urllib and regular expressions and json, and openpyxl and time modules
from yahoo_finance import Share #we are going to need yahoo finance module

urlbase='http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t=' #base url to talk with morningstar
ISparameters={'reportType':'is','period':'12','dataType':'A','order':'desc','columnYear':'5','number':'3'} #parameters to build URL for Income Statement
BSparameters={'reportType':'bs','period':'12','dataType':'A','order':'desc','columnYear':'5','number':'3'} #parameters to build URL for Balance Sheet
tagsIS=[('revenue','Revenue.*?,(\S*?),')] #tags to be extracted from income statement with is matching regex (list of tuples)
tagsBS=[('TCA','Total current assets.*?,(\S*?),'),('TNCA','Total non-current assets.*?,(\S*?),'),('TCL','Total current liabilities.*?,(\S*?),'),('TNCL','Total non-current liabilities.*?,(\S*?),'),
('GW','Goodwill.*?,(\S*?),'),('IA','Intangible assets.*?,(\S*?),'),('LTD','Long-term debt.*?,(\S*?),'),('OLTL','Other long-term liabilities.*?,(\S*?),')] #tags to be extracted from balance sheet
APIID='fe39ef6b876a482e874dba7bb29af7de' #ID for url for exchange rate
urlEXR='https://openexchangerates.org/api/latest.json?app_id='+APIID #build exchangerate DB URL

date=re.findall('\S*? (\S{3,}) (\d{1,2}) .+?:.+?:.+? (\d{4})',time.ctime()) #get day (0-31), month and year
month,day,year=date[0][0],date[0][1],date[0][2]

number=0
cell=0
try: #check if there is already a file for the current month
    wb = openpyxl.load_workbook(month + ' ' + year + '.xlsx')
    ws = wb.get_sheet_by_name('sheet1')
except: #if there was no file, generate a new one starting from the blankDMF.xlsx file
    cell,number=None,1
    print 'no existing file'
    wb = openpyxl.load_workbook('blankDMF.xlsx')
    ws = wb.get_sheet_by_name('sheet1')
while cell is not None: #determine where should we append the companies about to be entered.
    number+=1
    cell=ws['C'+str(number*3+1)].value

print month, year
while True:
    company=raw_input('Enter company ID -> ')
    if len(company)<1:break
    urlIS=urlbase+company+'&&'+urllib.urlencode(ISparameters) #generate income statement url
    urlBS=urlbase+company+'&&'+urllib.urlencode(BSparameters) #generate balance sheet url
    urlKR='http://financials.morningstar.com/ajax/exportKR2CSV.html?t='+company+'&' #generate key ratios url

    dataIS=urllib.urlopen(urlIS).read() #get all necessary data
    dataBS=urllib.urlopen(urlBS).read()
    dataKR=urllib.urlopen(urlKR).read()
    filee=open('testData.txt','w')
    filee.write(dataIS) #print info for later diagnostic if necessary
    filee.write('\n')
    filee.write(dataBS)
    filee.write('\n')
    filee.write(dataKR)
    filee.close()

    #file2=open('testData.txt') #temporary access of data to prevent too many url requests
    #dataIS=file2.read()
    companyinfo={}
    for tag,regex in tagsIS:
        end=regex.find('.')
        if regex[:end] in dataIS:
            companyinfo[tag]=float(re.findall(regex,dataIS)[0])
        else:
            companyinfo[tag]=0

    for tag,regex in tagsBS:
        end=regex.find('.')
        if regex[:end] in dataBS:
            try: companyinfo[tag]=float(re.findall(regex,dataBS)[0])
            except:
                companyinfo[tag]=0
                print 'warning: ', tag, ' was not a number.'
        else:
            companyinfo[tag]=0

    epsall=re.findall('Earnings Per Share.*?,(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*)',dataKR) #get 10 years of eps
    i,addES,ES=1,True,0
    while i<11: #match eps1-10 with corresponding values and insert into companyinfo
        try:companyinfo['eps'+str(i)]=float(epsall[0][11-i])
        except:companyinfo['eps'+str(i)]=0
        if addES and companyinfo['eps'+str(i)]>0:ES+=1
        else:addES=False
        i+=1
    companyinfo['ES']=ES

    dividends=re.findall('Dividends.*?,(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?)',dataKR) #get all the dividends paid
    DR,i=0,9
    while i>0: #calculate dividend record
        try:
            if float(dividends[0][i])>0:
                DR+=1
                i-=1
            else:break
        except:break
    companyinfo['DR']=DR

    companyinfo['currency']=re.findall('.*\. (\S{3}) in millions except per share data',dataIS)[0] #get what currency the data is in
    if companyinfo['currency']=='USD':
        companyinfo['EXR']=1
    else:
        dataEXR=urllib.urlopen(urlEXR).read()
        jsonEXR=json.loads(dataEXR)
        rates=jsonEXR.get('rates',None)
        companyinfo['EXR']=rates.get(companyinfo['currency'],None) #get the exchange rate

    #price and market capitalization
    share=Share(company)
    companyinfo['quote']=share.get_price()
    companyinfo['MC']=float(share.get_market_cap()[:-1])
    companyinfo['name']=share.get_name()

    #write to DMF excel
#    if number==1:wb=openpyxl.load_workbook('blankDMF.xlsx')
#    else: wb=openpyxl.load_workbook(month+' '+year+'.xlsx')
#    ws=wb.get_sheet_by_name('sheet1')
    ws['B' + str(number * 3 + 1)] = companyinfo['name']
    ws['C' + str(number * 3 + 1)] = company
    ws['F' + str(number * 3 + 1)] = companyinfo['EXR']
    ws['G' + str(number * 3 + 1)] = companyinfo['quote']
    ws['H' + str(number * 3 + 1)] = companyinfo['MC']
    ws['I' + str(number * 3 + 1)] = companyinfo['eps1']
    ws['I' + str(number * 3 + 2)] = companyinfo['eps2']
    ws['I' + str(number * 3 + 3)] = companyinfo['eps3']
    ws['J' + str(number * 3 + 1)] = companyinfo['revenue']
    ws['L' + str(number * 3 + 1)] = companyinfo['TCA']
    ws['L' + str(number * 3 + 2)] = companyinfo['TNCA']
    ws['N' + str(number * 3 + 1)] = companyinfo['TCL']
    ws['N' + str(number * 3 + 2)] = companyinfo['TNCL']
    ws['P' + str(number * 3 + 1)] = companyinfo['GW']
    ws['P' + str(number * 3 + 2)] = companyinfo['IA']
    ws['R' + str(number * 3 + 1)] = companyinfo['LTD']
    ws['R' + str(number * 3 + 2)] = companyinfo['OLTL']
    ws['U' + str(number * 3 + 1)] = companyinfo['eps8']
    ws['U' + str(number * 3 + 2)] = companyinfo['eps9']
    ws['U' + str(number * 3 + 3)] = companyinfo['eps10']
    ws['S' + str(number * 3 + 1)] = companyinfo['ES']
    ws['V' + str(number * 3 + 1)] = companyinfo['DR']
    wb.save(month+' '+year+'.xlsx')

    number+=1

    print companyinfo
