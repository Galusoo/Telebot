import yfinance as yf

class Stocks:
    def __init__(self, currentPrice, volume, _52WeekChange ,beta, trailingPE):
        self.price = currentPrice
        self.volume = volume
        self._52WeekChange = _52WeekChange
        self.beta = beta
        self.trailingPE = trailingPE
        
    def __str__(self):
        return f"Current Price : {self.price}\nVolume : {self.volume}\n52 Week Change : {self._52WeekChange}\nBeta : {self.beta}\nTrailing PE : {self.trailingPE}"

# Get the stock data with API
def fetch_stock_data(stock_name):
    dat = yf.Ticker(stock_name)
    stock = Stocks(dat.info['currentPrice'], dat.info['volume'], dat.info['52WeekChange'], dat.info['beta'], dat.info['trailingPE'])
    return stock if stock else "No data found"

