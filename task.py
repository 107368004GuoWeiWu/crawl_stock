#!/usr/bin/env python
# coding: utf-8

from pandas_datareader import data as web
import datetime as dt
from datetime import datetime
import yfinance as yf
import warnings
import pandas as pd
import time 
import warnings
import schedule
import time

warnings.filterwarnings("ignore")
yf.pdr_override()

# set up limit to show every row
pd.set_option('display.max_rows', 10000)
pd.options.display.float_format = '{:,.2f}'.format

timestamps_of_a_day=86400

def job():
    # ST=the stock day, ET=the stock day+1
    ts = dt.datetime.now().timestamp()
    ST=datetime.fromtimestamp(ts+timestamps_of_a_day*0).strftime('%Y-%m-%d')
    ET=datetime.fromtimestamp(ts+timestamps_of_a_day*1).strftime('%Y-%m-%d')

    # read_excel open .xlsx to load data
    df_sample_excel = pd.read_excel("./result/"+ST[0:4]+".xlsx")

    # use sample data to acquire stock's No.
    # read_csv open .csv to load data
    df_stock_name_list = []
    for i in range(len(df_sample_excel['編號'])):
        df_csv_name = str(df_sample_excel['編號'][i])
        df_stock_name_list.append(df_csv_name)
    # df_stock_name_list is list of stock

    start = dt.datetime(int(ST[0:4]), int(ST[5:7]), int(ST[8:10]))
    end = dt.datetime(int(ET[0:4]), int(ET[5:7]), int(ET[8:10]))
    stock_close_list = []
    cnt=0
    for i in range(len(df_stock_name_list)):
        try:
            df = web.get_data_yahoo([df_stock_name_list[i]+'.TW'], start, end)
            stock_close_list.append(df['Close'][0])
        except IndexError:
            stock_close_list.append(None)
            cnt=cnt+1
    # stock_close_list is list of stock close data

    # cnt!=len(stock_close_list) represent data existed
    if cnt!=len(stock_close_list):
        # read 
        new_df = pd.read_excel("./result/"+ST[0:4]+".xlsx")

        for i in range(len(stock_close_list)):
            stock_close_list[i]
            df_temp = pd.DataFrame(stock_close_list,columns=[ST]) 
            new_df[ST]=df_temp[ST]
        print(ST)

        # store data
        new_df.to_excel("./result/"+ST[0:4]+".xlsx", index=False)  

# choose a time to do job()
schedule.every().day.at("10:30").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
