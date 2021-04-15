# RSI-Backtester
# This is program that backtests a simple RSI based trading strategy on a given stock. It then uploads a report of the backtest onto a excel workbook.

import pandas as pd

from pandas_datareader import data as wb

from datetime import date

import xlwings as xw


#function to find rsi for a stock between 01/01/2016 and present day
def RSI(Ticker):
    # start date before 2016 as need some data to calculate rsi for 01/01/2016
    start = '2015-12-10'
    end = date.today()
    # Get data
    data = wb.DataReader(Ticker, 'yahoo', start, end)
    # Get just the adjusted close
    close = data['Adj Close']
    # Get the difference in price from previous step
    delta = close.diff()
    # Get rid of the first row, which is NaN since it did not have a previous
    # row to calculate the differences
    delta = delta[1:]
    # Make the positive gains (up) and negative gains (down) Series
    up, down = delta.copy(), delta.copy()
    up[up < 0] = 0
    down[down > 0] = 0

    # Calculate the EWMA
    roll_up = up.ewm(span=14).mean()
    roll_down = down.abs().ewm(span=14).mean()

    # Calculate the RSI based on EWMA
    RS = roll_up / roll_down
    RSI = 100.0 - (100.0 / (1.0 + RS))
    return RSI


# function to back test strategy just using rsi
def BT_RSI(Ticker, RSIBuy, RSISell):
    # start date before 2016 as need some data to calculate rsi for 01/01/2016
    start = '2015-12-10'
    # sets end date as present day
    end = date.today()
    # gets data from yahoo finance
    d = wb.DataReader(Ticker, 'yahoo', start, end)
    data2 = d['Adj Close']
    # Set as empty list to store buy and sell dates in
    Bdate = []
    Sdate = []
    # Calls RSI func
    rsi = RSI(Ticker)
    # Gets dates of RSI and stores in a list
    keys = list(rsi.keys())
    # Find all buy and sell dates (as index of dataframe) when rsi criteria met
    for i in range((len(rsi)-14)):
        if rsi[i+14] < RSIBuy and rsi[i:i+14].min() > RSIBuy:
            Bdate.append(i+14)
        elif rsi[i+14] > RSISell and rsi[i:i+14].max() < RSISell:
            Sdate.append(i+14)

    # combine buy and sell dates into list 'total'
    total = Bdate.copy()
    total.extend(Sdate)
    total.sort()
    # gets rid of useless dates
    # EG 2 sells in a row, 2 buys in a row
    j = 0
    for i in range(len(total)):
        if j % 2 == 0:
            if total[j] in Bdate:
                j += 1
            else:
                total.pop(j)
        else:
            if total[j] in Sdate:
                j += 1
            else:
                total.pop(j)
    # removes dates if first date is a sell or last is a buy
    if len(total) >= 1:
        if total[-1] in Bdate:
            total.pop(-1)
        elif total[0] in Sdate:
            total.pop(0)
        else:
            pass

    # Dataframe to show results
    Res = pd.DataFrame(columns=['Buy Date','Buy Price','Sell Date','Sell Price','Days Held','Profit'])

    # Fills dataframe
    for w in range(0, len(total), 2):
         q = keys[total[w]]
         e = keys[total[w + 1]]
         r = e - q
         p1 = data2[keys[total[w]]]
         p2 = data2[keys[total[w+1]]]
         profit = round(((p2-p1)/p1)*100, 2)
         Res.loc[w/2] = [q, round(p1, 2), e, round(p2, 2), r, profit]
    # returns dataframe with full back test
    return Res






# Calls RSI backtest function (See below)
# BT_RSI("Stock Ticker", lower rsi bound, upper rsi bound)
Test = BT_RSI("D", 30, 70)
# find length of pandas dataframe test
lenTest = len(Test)
# prints pandas dataframe test
print(Test)
# opens excel workbook and selects first sheet
wb = xw.Book()
sheet = wb.sheets[0]

# sets row titles for worksheet
sheet.range((1, 1)).value = "Buy Date"
sheet.range((1, 2)).value = "Buy Price"
sheet.range((1, 3)).value = "Sell Date"
sheet.range((1, 4)).value = "Sell Price"
sheet.range((1, 5)).value = "Days Held"
sheet.range((1, 6)).value = "Profit"

sheet.range((1, 8)).value = "Summary:"
sheet.range((1, 9)).value = "Mean Profit"
sheet.range((1, 10)).value = "Mean Days Held"


# Takes Values from database and puts on workbook
for i in range(0, 6):
    for j in range(0, lenTest):
        if i == 4:
            sheet.range((j+2, i+1)).value = float(str(Test.iloc[j,i]).replace(' days 00:00:00', ''))
        elif i == 5:
            sheet.range((j + 2, i + 1)).value = str(Test.iloc[j, i]) + '%'
        else:
            sheet.range((j+2, i+1)).value = Test.iloc[j, i]

# puts mean days and profit on book
sheet.range((2, 9)).formula = "=Average(F2:F" + str(lenTest+1) + ")"
sheet.range((2, 10)).formula = "=Average(E2:E" + str(lenTest+1) + ")"

