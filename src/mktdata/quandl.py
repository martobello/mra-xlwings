import datetime as dt
import quandl
import xlwings as xw

#
# xwQuandlQuote - Get stock quote data from Quandl
#
@xw.func
@xw.ret(expand='table')
def xwQuandlQoute(symbol,startdate=None,enddate=None,apikey=None):

    if isinstance(startdate, dt.datetime):
        startdate = startdate.strftime('%Y-%m-%d')
    if isinstance(enddate, dt.datetime):
        enddate = enddate.strftime('%Y-%m-%d')

    df = quandl.get(symbol,start_date=startdate,end_date=enddate,api_key=apikey)
    
    return df


#
# In PyDev for Eclipse, Debug as Python Run starts the Debug Server
#
if __name__ == '__main__':
    xw.serve()
