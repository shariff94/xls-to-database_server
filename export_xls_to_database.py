import pyodbc
import datetime
import xlrd
import sys

def sqlit(s):
    s = str(s)
    if s == '':
        return None
    else:
        return s

log = open("Output.log", "w")
sys.stdout = log

try:
    #username = 'username'
    #password = 'password'
    server = 'INBLR03L076\SQLEXPRESS'
    db = 'test'
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + db + ';Trusted_Connection=yes;')
    #conn = pyodbc.connect ('DRIVER={SQL Server};SERVER='+server+';DATABASE='+db+';UID='+username+';PWD='+password+';')
    cursor = conn.cursor()
    print(datetime.datetime.now(),"-","Connection to DB established")

except:
    print("Connection to the DB failed!!!!")
    raise SystemError

try:
    book = xlrd.open_workbook('Currency_Exchange_Rates.xlsx')
    sheet = book.sheet_by_index(0)
    print(datetime.datetime.now(), "-", "Excel File 'Currency_Exchange_Rates.xlsx' opened")
except:
    print(datetime.datetime.now(), "-", "Excel File 'Currency_Exchange_Rates.xlsx' was not found")
    raise FileNotFoundError

query1 = """
CREATE TABLE [upload].[ExchangeRates](
    [Source Currency] varchar(5),
    [Target Currency] varchar(5),
    [Currency Rate Type] varchar(20),
    [Currency Rate Timestamp] datetime,
    [Currency Rate] decimal(10,6)
    )
"""

query2 = """
TRUNCATE TABLE [upload].[ExchangeRates]
"""


query3 = """
INSERT INTO [upload].[ExchangeRates](
    [Source Currency],
    [Target Currency],
    [Currency Rate Type],
    [Currency Rate Timestamp],
    [Currency Rate]
)VALUES (?, ?, ?, ?,?)"""

try:
    cursor.execute(query1)
    print(datetime.datetime.now(), "-", "Table was created.")
except:
    cursor.execute(query2)
    print(datetime.datetime.now(), "-", "Table was already present, which is now truncated")

print(datetime.datetime.now(), "-", "Exporting Data from excel to DB started")
for k in range (1,sheet.nrows):
    try:
        SourceCurrency = sqlit(sheet.cell(k,0).value)
        TargetCurrency = sqlit(sheet.cell(k,1).value)
        CurrencyRateType = sqlit(sheet.cell(k,2).value)
        CurrencyRateTimestamp = sqlit(str(datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(k, 3).value, book.datemode))))
        CurrencyRate = sqlit(sheet.cell(k, 4).value)
        value_tuple = (SourceCurrency,TargetCurrency,CurrencyRateType,CurrencyRateTimestamp,CurrencyRate)
        cursor.execute(query2, value_tuple)
        print(datetime.datetime.now(), "-", "Row", k, "succesfully inserted")
    except:
        print(datetime.datetime.now(), "-", "Row", k, "was not inserted, please check row", k)



cursor.commit()
print(datetime.datetime.now(), "-", "Exporting Data from excel to DB finished")

log.close()
cursor.close()
conn.close()