import Quandl as q
import matplotlib.pyplot as plt
from openpyxl.reader.excel import load_workbook
import os
import sys
import re
import pandas as pd
from datetime import datetime

def writeDataToFile(filename, spread):
    cache = open("cache-data/" + filename + ".cache", "w+")
    cache.write(str(spread))
    cache.close()

def loadSpreadMatrix(filename):
    wb = load_workbook(filename)
    table = []
    for sheet in wb.worksheets:
        for row in sheet.rows:
            table_row = []
            for cell in row:
                value = cell.value
                if value is None:
                    value = ''
                if not isinstance(value, unicode):
                    value = unicode(value)
                value = value.encode('utf8')
                table_row.insert(len(table_row), value)
            table.insert(len(table), table_row)
    return table

def getSpreadDelta(row):
    spread = loadQuandlSpread(row[0], row[1], row[2], int(row[5][:4]), int(row[6][:4]), int(row[3]), int(row[4]), row[5], row[6], int(row[7]), True)  
    return convertSpreadSeriesToDelta(spread)

def loadQuandlSpread(CONTRACT, M1, M2, ST_YEAR, END_YEAR, CONT_YEAR1, CONT_YEAR2, ST_DATE, END_DATE, BUCK_PRICE, STARTFROMZERO):
    year = str(ST_YEAR)
    price = str(BUCK_PRICE)
    filename = CONTRACT + M1 + M2 + year + ST_DATE + END_DATE + price 
    filename = re.sub('[/ ]', '_', filename)

    years = [sys.argv[2], sys.argv[3]]
    cont1 = CONTRACT + M1 + str(ST_YEAR + CONT_YEAR1)
    cont2 = CONTRACT + M2 + str(ST_YEAR + CONT_YEAR2)
    print ("contract1: " + cont1)
    print ("contract2: " + cont2)

    startdate = datetime.strptime(ST_DATE, '%Y-%m-%d %H:%M:%S')
    enddate = datetime.strptime(END_DATE, '%Y-%m-%d %H:%M:%S')

    for i in years:
        startDate = startdate.replace(year=ST_YEAR - 2000 + int(i))
        endDate = enddate.replace(year=END_YEAR - 2000 + int(i))

        data1 = q.get(cont1, authtoken="gmgPdqznEbntQRCrt3Wu", trim_start=startDate, trim_end=endDate)
        data2 = q.get(cont2, authtoken="gmgPdqznEbntQRCrt3Wu", trim_start=startDate, trim_end=endDate)
        spread = (data1-data2).Settle*BUCK_PRICE
    if STARTFROMZERO:
        spread = spread - spread[0]

    writeDataToFile(filename, spread)
    return spread

def convertSpreadSeriesToDelta(DATA):
    DATADELTA = DATA.copy(True)
    previ = DATA.index[0]
    for i in DATA.index:
        if DATA.index[0] != i:
            DATADELTA[i] = DATA.ix[i] - DATA.ix[previ]
            previ = i
    return DATADELTA

def convertDeltaSeriesToCumulativeGraph(DATA):
    GRAPHDATA = DATA.copy(True)
    previ = DATA.index[0]
    for i in DATA.index:
        if DATA.index[0] == i:
            GRAPHDATA.ix[i] = DATA.ix[i]
            previ = i
        else:
            GRAPHDATA.ix[i] = GRAPHDATA.ix[previ] + DATA.ix[i]
            previ = i
    return GRAPHDATA

##############################

table = loadSpreadMatrix(sys.argv[1])
spread1Delta = getSpreadDelta(table[1])
spread2Delta = getSpreadDelta(table[2])
totalSpreadDelta = spread1Delta.add(spread2Delta, fill_value = 0)
for i in range(0, 3):
    del table[0]
for row in table:
    totalSpreadDelta = totalSpreadDelta.add(getSpreadDelta(row), fill_value = 0)
totalCumulativeChart = convertDeltaSeriesToCumulativeGraph(totalSpreadDelta)

print("Total Cumulative Chart:")
print (totalCumulativeChart.astype(int))

plt.plot(totalCumulativeChart.index, totalCumulativeChart)
plt.show()
