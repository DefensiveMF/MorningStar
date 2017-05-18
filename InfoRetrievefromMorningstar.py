import urllib, re, json #import urllib and regular expressions and json
from yahoo_finance import Share

urlbase='http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t='
ISparameters={'reportType':'is','period':'12','dataType':'A','order':'desc','columnYear':'5','number':'3'}
BSparameters={'reportType':'bs','period':'12','dataType':'A','order':'desc','columnYear':'5','number':'3'}
tagsIS=[('revenue','Revenue.*?,(\S*?),')]
tagsBS=[('TCA','Total current assets.*?,(\S*?),'),('TNCA','Total non-current assets.*?,(\S*?),'),('TCL','Total current liabilities.*?,(\S*?),'),('TNCL','Total non-current liabilities.*?,(\S*?),'),
('GW','Goodwill.*?,(\S*?),'),('IA','Intangible assets.*?,(\S*?),'),('LTD','Long-term debt.*?,(\S*?),'),('OLTL','Other long-term liabilities.*?,(\S*?),')]
APIID='fe39ef6b876a482e874dba7bb29af7de'
urlEXR='https://openexchangerates.org/api/latest.json?app_id='+APIID
order =['']

def mean(numbers):
    return float(sum(numbers)) / max(len(numbers), 1)

while True:
    company=raw_input('Enter company ID -> ')+'&'
    ticker = company[:-1]
    if len(company)<2:break
    urlIS=urlbase+company+'&'+urllib.urlencode(ISparameters) #generate income statement url
    urlBS=urlbase+company+'&'+urllib.urlencode(BSparameters) #generate balance sheet url
    urlKR='http://financials.morningstar.com/ajax/exportKR2CSV.html?t='+company #generate key ratios url

    dataIS=urllib.urlopen(urlIS).read() #get all necessary data
    dataBS=urllib.urlopen(urlBS).read()
    dataKR=urllib.urlopen(urlKR).read()
    filee=open('testData.txt','w')
    filee.write(dataIS)
    filee.close()

    #file2=open('testData.txt') #temporary access of data to prevent too many url requests
    #dataIS=file2.read()
    companyinfo={}
    share = Share(ticker)
    companyinfo['quote'] = float(share.get_price())
    companyinfo['MC'] = float(share.get_market_cap()[:-1])

    for tag,regex in tagsIS:
        end=regex.find('.')
        if regex[:end] in dataIS:
            companyinfo[tag]=float(re.findall(regex,dataIS)[0])
        else:
            companyinfo[tag]=0

    for tag,regex in tagsBS:
        end=regex.find('.')
        if regex[:end] in dataBS:
            companyinfo[tag]=float(re.findall(regex,dataBS)[0])
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
    filters = []
    filters.append(companyinfo['revenue'])
    filters.append(companyinfo['TCA'] + companyinfo['TNCA'])
    filters.append(companyinfo['TCA']/companyinfo['TCL'])
    filters.append(companyinfo['LTD'] + companyinfo['OLTL'])
    filters.append(companyinfo['TCA'] - companyinfo['TCL'])
    book_value = companyinfo['TCA'] - (companyinfo['IA'] + companyinfo['GW']+ companyinfo['TCL'])
    filters.append(book_value)
    #Calculate Earnings Stability
    eps = 0
    value = 0
    year = 1
    while value >= 0:
        value = companyinfo['eps'+str(year)]
        if value > 0:
            eps += 1
        year += 1
        if year == 10:
            break
    filters.append(eps)
    filters.append(companyinfo['DR'])
    growth = mean([companyinfo['eps10'],companyinfo['eps9'],companyinfo['eps8']]) / mean([companyinfo['eps3'],companyinfo['eps2'],companyinfo['eps1']])
    filters.append(growth)
    p_e = companyinfo['quote']/ mean([companyinfo['eps10'],companyinfo['eps9'],companyinfo['eps8']])
    filters.append(p_e)
    p_b = (companyinfo['MC']*1000*companyinfo['EXR'])/book_value
    filters.append(p_b)

    print filters
    print companyinfo
