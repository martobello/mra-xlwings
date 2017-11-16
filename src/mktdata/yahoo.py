import requests
import xlwings as xw

from pandas_datareader.data import Options

#
# xwYahooOptionChain - Get option chain data from Yahoo
#
@xw.func
@xw.ret(expand='table')
def xwYahooOptionChain(symbol,optype=None,bid=None,minstrike=None,maxstrike=None,expiry=None):
    
    stk = Options(symbol, 'yahoo')

    df = stk.get_all_data()
    
    df.drop(['Chg','PctChg','Vol','Open_Int','IV','Root','IsNonstandard','Underlying','Underlying_Price','Quote_Time','Last_Trade_Date','JSON'], axis=1,inplace=True)
    df.reset_index(inplace=True)
   
    if optype is not None:
        if optype.lower() == 'call': df = df.loc[df['Type'] == 'call']
        if optype.lower() == 'put':  df = df.loc[df['Type'] == 'put']

    if bid is not None:
        if bid > 0: df = df.loc[df['Bid'] >= bid]
     
    if minstrike is not None and maxstrike is not None:
        if minstrike != maxstrike:
            df = df.loc[(df['Strike'] >= minstrike) & (df['Strike'] <= maxstrike)]
         
    df.sort_values(by=['Expiry','Strike','Type'],ascending=[True,True,True],inplace=True)
    df.set_index('Expiry',inplace=True)

    return df    

#
# xwYahooQuote - Get stock quote data from Yahoo
#
@xw.func
def xwYahooQuote(symbol,data):
    
    url = "https://query1.finance.yahoo.com/v7/finance/quote"
    
    if   data.lower() == 'ask':         field = 'ask'
    elif data.lower() == 'bid':         field = 'bid'
    elif data.lower() == 'fullname':    field = 'longName'
    elif data.lower() == 'high':        field = 'regularMarketDayHigh'
    elif data.lower() == 'last':        field = 'regularMarketPrice'
    elif data.lower() == 'low':         field = 'regularMarketDayLow'
    elif data.lower() == 'name':        field = 'shortName'
    elif data.lower() == 'prev':        field = 'regularMarketPreviousClose'
    elif data.lower() == 'open':        field = 'regularMarketOpen'
    else: pass                              
    
    parms = {
        'fields': field,
        'symbols': symbol}
    
    r = requests.get(url, params=parms)
    
    for i in r.json()['quoteResponse']['result']:
        if field in i:
            return i[field]
