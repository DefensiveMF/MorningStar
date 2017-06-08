import argparse,DMFlib



if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Choose whether to input tickers manually or use a spreadsheet')
    parser.add_argument('-i', '--input', help="Enter this command followed by the excel spreadsheet of tickers")
    args = parser.parse_args()

    file = args.input
    companies = DMFlib.ReadFromExcel(file)
    for company in companies:
        print DMFlib.GetInfo(company)

