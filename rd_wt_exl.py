import xlrd


def read_excel(path):
    symbols = []
    fliter = 'OKEx'
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    lastSymbol =''
    #print(sheet.cell(2,3))
    #symbol,exchange = sheet.cell(2,3),sheet.cell(2,4)
    for i in range(2,2576):
        symbol,exchange = sheet.cell_value(i,3),sheet.cell_value(i,4)
        
        #有合并单元格的情况
        if not len(symbol):
            symbol = lastSymbol
        else:
            lastSymbol = symbol
        
        #print(symbol,exchange)
        if fliter in exchange:
            if not len(symbol):
                print(i,exchange)
                continue
            symbols.append(symbol.strip())
    return symbols


if __name__=="__main__":
    symbols = read_excel("./货币对03.xlsx")
    print(len(symbols))
    print(symbols)
    
