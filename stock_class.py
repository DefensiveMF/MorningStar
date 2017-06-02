from yahoo_finance import Share
import urllib, re, json, csv, argparse

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



class Company:
    def __init__(self,ticker):
        self.ticker = str(ticker)
        self.yahoo = Share(ticker)
        self.yahoo_name = self.yahoo.get_name()
        if self.yahoo.get_price() is None:
            self.quote = False
        else:
            self.quote = float(self.yahoo.get_price())
        if self.yahoo.get_market_cap() is None:
            self.MC = False
        else:
            self.MC = float(self.yahoo.get_market_cap()[:-1])
        self.companyinfo = self.get_company_info()
        self.filters = self.calculate_filters()
    def generate_reports(self):
        urlIS = urlbase + self.ticker + '&&' + urllib.urlencode(ISparameters)  # generate income statement url
        urlBS = urlbase + self.ticker + '&&' + urllib.urlencode(BSparameters)  # generate balance sheet url
        urlKR = 'http://financials.morningstar.com/ajax/exportKR2CSV.html?t=' + self.ticker +'&'  # generate key ratios url

        self.dataIS = urllib.urlopen(urlIS).read()  # get all necessary data
        self.dataBS = urllib.urlopen(urlBS).read()
        self.dataKR = urllib.urlopen(urlKR).read()

    def get_company_info(self):
        companyinfo = {}
        self.generate_reports()
        for tag,regex in tagsIS:
            end=regex.find('.')
            if regex[:end] in self.dataIS:
                try:
                    companyinfo[tag]=float(re.findall(regex,self.dataIS)[0])
                except:
                    companyinfo[tag] = 0
            else:
                companyinfo[tag]=0
        for tag,regex in tagsBS:
            end=regex.find('.')
            if regex[:end] in self.dataBS:
                try:
                    companyinfo[tag]=float(re.findall(regex,self.dataBS)[0])
                except:
                    companyinfo[tag] = 0
            else:
                companyinfo[tag]=0

        epsall=re.findall('Earnings Per Share.*?,(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*)',self.dataKR) #get 10 years of eps
        i,addES,ES=1,True,0
        while i<11: #match eps1-10 with corresponding values and insert into companyinfo
            try:companyinfo['eps'+str(i)]=float(epsall[0][11-i])
            except:companyinfo['eps'+str(i)]=0
            if addES and companyinfo['eps'+str(i)]>0:ES+=1
            else:addES=False
            i+=1
        companyinfo['ES']=ES

        dividends=re.findall('Dividends.*?,(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?),(\S*?)',self.dataKR) #get all the dividends paid
        DR,i=0,9
        while i>0: #calculate dividend record
            try:
                if float(dividends[0][i])>0:
                    DR+=1
                    i-=1
                else:break
            except:break
        companyinfo['DR']=DR
        try:
            companyinfo['currency']=re.findall('.*\. (\S{3}) in millions except per share data',self.dataIS)[0] #get what currency the data is in
        except:
		    companyinfo['currency'] = 'NA'
        if companyinfo['currency']=='USD' or companyinfo['currency'] == 'NA':
            companyinfo['EXR']=1
        else:
            dataEXR=urllib.urlopen(urlEXR).read()
            jsonEXR=json.loads(dataEXR)
            rates=jsonEXR.get('rates',None)
            companyinfo['EXR']=rates.get(companyinfo['currency'],None) #get the exchange rate
        return companyinfo
    def calculate_filters(self):
        filters = []
        companyinfo = self.companyinfo
        filters.append(companyinfo['revenue'])
        calculations = ["companyinfo['TCA'] + companyinfo['TNCA']","companyinfo['TCA'] / companyinfo['TCL']","companyinfo['LTD'] + companyinfo['OLTL']",
        "companyinfo['TCA'] - companyinfo['TCL']","companyinfo['TCA'] - (companyinfo['IA'] + companyinfo['GW'] + companyinfo['TCL'])",
        "mean([companyinfo['eps10'], companyinfo['eps9'], companyinfo['eps8']]) / mean([companyinfo['eps3'], companyinfo['eps2'], companyinfo['eps1']])",
        "(self.MC * 1000 * companyinfo['EXR']) / book_value"]
        try:
            book_value = eval(calculations[-2])
        except:
		    book_value = 0
        for c in calculations:
            try:
                x = eval(c)
                filters.append(x)
            except:
                print c
                filters.append('Unable to calculate')			
        # Calculate Earnings Stability
        eps = 0
        value = 0
        year = 1
        while value >= 0:
            value = companyinfo['eps' + str(year)]
            if value > 0:
                eps += 1
            year += 1
            if year == 10:
                break
        filters.append(eps)
        filters.append(companyinfo['DR'])
        try:
            p_e = companyinfo['quote'] / mean([companyinfo['eps10'], companyinfo['eps9'], companyinfo['eps8']])
        except:
            p_e = 0
        filters.append(p_e)
        return filters



def read_tickers(tickers_file,output_file):
    companies = []
    for line in tickers_file:
        print line
        line = line.split(',')
        ticker = line[0]
        name1 = line[1].replace('*', ',')[:-1]
        obj = Company(ticker)
        if obj.quote is False:
            continue
        companies.append(obj)
    output_companies(output_file,companies)
def output_companies(output_file,companies):
    with open(output_file,"w") as of:
        wr = csv.writer(of)
        for i in companies:
            print i
            wr.writerow([i.yahoo_name, i.ticker, i.quote, i.MC,
                     i.companyinfo['currency'], i.companyinfo['EXR'],
                     i.companyinfo['eps1'], i.companyinfo['eps2'], i.companyinfo['eps3'],
                     i.companyinfo['eps8'],i.companyinfo['eps9'], i.companyinfo['eps10'], i.filters[0],
                     i.filters[1], i.filters[2], i.filters[3],i.filters[4], i.filters[5],
                     i.filters[6], i.filters[7], i.filters[8], i.filters[9]])
        


if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Get company information from Morningstar and Yahoo Finance')
    parser.add_argument('-i', '--input', help="csv of company tickers")
    parser.add_argument('-o', '--output', help="csv with company information and Graham's filters.")
    args = parser.parse_args()

    tickers_file = open(args.input,"r")
    output_file = args.output
    read_tickers(tickers_file,output_file)