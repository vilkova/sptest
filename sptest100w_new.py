import Quandl as q
import matplotlib.pyplot as plt
from openpyxl.reader.excel import load_workbook
import os
import sys
import re
import pickle
from datetime import datetime
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd

AUTH_TOKEN = 'e6FuWkfWH9qypKzJz6sR'
CACHE_DIR = "cache-data/"

def main():
    if not os.path.exists(CACHE_DIR):
        os.makedirs(CACHE_DIR)
    table = loadSpreadMatrix(sys.argv[1])
    spread1Delta = getSpreadDelta(table[1])

    if len(table) == 2:
        totalSpreadDelta = spread1Delta.add(spread1Delta)
        totalCumulativeChart = convertDeltaSeriesToCumulativeGraph(totalSpreadDelta)
        print("Total Cumulative Chart:")
        print(totalCumulativeChart.astype(int))
        showPlot(totalCumulativeChart)

    else:
        spread2Delta = getSpreadDelta(table[2])
        totalSpreadDelta = spread1Delta.add(spread2Delta, fill_value = 0)
        for i in range(0, 3):
            del table[0]
        for row in table:
            totalSpreadDelta = totalSpreadDelta.add(getSpreadDelta(row), fill_value = 0)
        totalCumulativeChart = convertDeltaSeriesToCumulativeGraph(totalSpreadDelta)
        print("Total Cumulative Chart:")
        print(totalCumulativeChart.astype(int))
        showPlot(totalCumulativeChart)

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
                if not isinstance(value, str):
                    value = str(value)
                value = value.encode('utf8')
                table_row.insert(len(table_row), value)
            table.insert(len(table), table_row)
    return table

def getSpreadDelta(row):
    if len(sys.argv) == 2:
        years = [2000]
    else:
        years = range(int(sys.argv[2]), int(sys.argv[3]) + 1)

    spread = fetchSpread(row[0].decode("utf-8"), row[1].decode("utf-8"), row[2].decode("utf-8"), int(row[5][:4].decode("utf-8")),
                              int(row[6][:4].decode("utf-8")), int(row[3].decode("utf-8")), int(row[4].decode("utf-8")), row[5].decode("utf-8"),
                              row[6].decode("utf-8"), int(row[7].decode("utf-8")), True, years)
    return convertSpreadSeriesToDelta(spread)

def checkIfCached(filename):
    fileNames = os.listdir(CACHE_DIR)
    for fileName in fileNames:
        if fileName == filename:
            return True
    return False

def readCacheFromFile(filename):
    cacheFile = open(CACHE_DIR + filename, "rb")
    cache = pickle.load(cacheFile)
    cacheFile.close()
    return cache

def fetchSpread(CONTRACT, M1, M2, ST_YEAR, END_YEAR, CONT_YEAR1, CONT_YEAR2, ST_DATE, END_DATE, BUCK_PRICE, STARTFROMZERO, years):
    startdate = datetime.strptime(ST_DATE, '%Y-%m-%d %H:%M:%S')
    enddate = datetime.strptime(END_DATE, '%Y-%m-%d %H:%M:%S')
    totalSpread = pd.Series()
    lastValue = 0
    spread = pd.Series()
    for i in years:
        year = str(i)
        price = str(BUCK_PRICE)
        filename = CONTRACT + M1 + M2 + year + ST_DATE + END_DATE + price
        filename = re.sub('[/ ]', '_', filename)
        filename = re.sub('[:]', '.', filename)
        cont1 = str(CONTRACT) + str(M1) + str(i + CONT_YEAR1)
        cont2 = str(CONTRACT) + str(M2) + str(i + CONT_YEAR2)
        print ("contract1: " + cont1)
        print ("contract2: " + cont2)

        startDate = startdate.replace(year = ST_YEAR - 2000 + i)
        endDate = enddate.replace(year = END_YEAR - 2000 + i)

        print('==============')
        print('Trim start:')
        print(startDate.strftime('%Y-%m-%d'))
        print('Trim end:')
        print(endDate.strftime('%Y-%m-%d'))
        print('==============')

        if not checkIfCached(filename):
            data1 = q.get(cont1, authtoken = AUTH_TOKEN, trim_start = startDate, trim_end = endDate)
            data2 = q.get(cont2, authtoken = AUTH_TOKEN, trim_start = startDate, trim_end = endDate)
            spread = (data1 - data2).Settle * BUCK_PRICE


        else:
            print("Loading cached data from file: %s !" %filename)
            cache = readCacheFromFile(filename)
            if years == cache['years']:
                spread = cache['spread']

        if STARTFROMZERO:
            if spread.size > 0:
                delta = lastValue - spread[0]
                spread = spread + delta
                totalSpread = totalSpread.append(spread)
                lastValue = totalSpread[-1]
                
                print('11111111§1111')
                print(spread[0])        
                print(lastValue)
                print('11111111§1111')

                writeCacheToFile(filename, totalSpread, years)
            else:
                print('There is no data for %s' %startdate)
                sys.exit(-1)
        # totalSpread = totalSpread.append(spread)
                    
    return totalSpread


def writeCacheToFile(filename, spread, years):
    cacheFile = open(CACHE_DIR + filename, 'wb')
    pickle.dump({
        'years': years,
        'spread': spread
    }, cacheFile)
    cacheFile.close()

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
    prev_date = DATA.index[0]
    for i in range(1, len(DATA.index)):
        date = DATA.index[i]
        GRAPHDATA.ix[date] = GRAPHDATA.ix[prev_date] + DATA.ix[date]
        prev_date = date
    return GRAPHDATA

def showPlot(totalCumulativeChart):
    def format_date(x, pos=None):
        thisind = np.clip(int(x+0.5), 0, N-1)
        return totalCumulativeChart.index[thisind].strftime('%b %Y')

    N = len(totalCumulativeChart)
    ind = np.arange(N)

    #shows plot with empty data intervals    
    # fig, ax = plt.subplots()
    # ax.plot(totalCumulativeChart.index, totalCumulativeChart)
    # fig.autofmt_xdate()

    #shows plot without empty data intervals
    fig, ax = plt.subplots()
    ax.plot(ind, totalCumulativeChart)
    ax.xaxis.set_major_formatter(ticker.FuncFormatter(format_date))
    fig.autofmt_xdate()

    plt.show()


main()
