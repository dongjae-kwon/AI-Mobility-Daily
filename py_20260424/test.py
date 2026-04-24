import xlwings as xw
import pandas as pd
import requests
import datetime

def collect_Ratingspread():
    wb = xw.Book.caller()
    ws = wb.sheets['금리스프레드']

    url = 'http://kisrating.com/ratingsStatistics/statics_spread.do#'
    r = requests.get(url)
    df = pd.read_html(r.text)  
    df = df[0]
    df = df.set_index('구분')

    ws.range("A5").value = df



if __name__ == "__main__":
    xw.Book("C:\work\AI-Mobility-Daily\py_20260424\test.xlsm").set_mock_caller()