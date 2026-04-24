import xlwings as xw
import pandas as pd
import requests
import datetime

def collect_Ratingspread():
    wb = xw.Book.caller()
    ws = wb.sheets['금리스프레드']

    url = 'https://kisrating.com/ratingsStatistics/statics_spread.do#'
    r = requests.get(url)
    df = pd.read_html(r.text)
    df = df[0]
    df = df.set_index('구분')

    ws.range("I4").value = "조회일자:" + datetime.datetime.now().strftime("%Y-%m-%d")
    ws.range("A5").value = df

    if __name__ == "__main__" :
        xw.Book("testwings.xlsm").set_mock_caller()
        collect_Ratingspread()