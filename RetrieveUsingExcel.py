import argparse
def ReadFromExcel(file):
    ticker_file = xlrd.open_workbook(file)
    sheet1 = ticker_file.sheet_by_index(0)
    tickers = sheet1.col_values(0)
    return tickers

def GetInfo(company):
    companyinfo = {}
    dictIS = IncomeStatement(company)
    dictBS = BalanceSheet(company)
    dictKR = KeyRatios(company)
    dictionaries = [dictIS, dictBS, dictKR]
    companyinfo['EXR'] = ExchangeRate(dictIS['currency'])

    # price and market capitalization
    share = Share(company)
    companyinfo['quote'] = share.get_price()
    companyinfo['MC'] = float(share.get_market_cap()[:-1])
    companyinfo['name'] = share.get_name()

    for dictionary in dictionaries:  # merge dictionaries
        companyinfo.update(dictionary)

    return companyinfo


if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Choose whether to input tickers manually or use a spreadsheet')
    parser.add_argument('-i', '--input', help="Enter this command followed by the excel spreadsheet of tickers")
    args = parser.parse_args()

    file = args.input
    companies = ReadFromExcel(file)
    for company in companies:
        print GetInfo(company)

