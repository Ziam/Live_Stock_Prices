"""
Created on Thu May 28 13:53:34 2020

@author: Ziam Ghaznavi

This script first pulls tickers from an excel sheet in the 
same directory, and then it web scraps yahoo finances to retrieve
live prices. Finally, it updates the excel sheet with the live prices
"""

import pandas as pd
from yahoo_fin import stock_info
from openpyxl import *
from datetime import datetime
import os

# Handy constants
file = "myTickers.xlsx"
sheet = "Prices"

def get_tickers():
    # Retrieves tickers from "Prices" sheet of excel workbook
    data = pd.read_excel(file,
                         sheet_name=sheet,
                         usecols='A',
                         skiprows=1,
                         header=None).rename(columns={0: "tickers"})
    
    # Formats df so it can be feed it directly to yahoo_fin
    data = data.dropna().drop(index=data[
        (data['tickers'] == '-') | (data['tickers'] == 'Total')
        ].index)    
    
    data['tickers'] = data['tickers'].str.split(' ').str.get(0)
    
    print('Tickers collected\n')
    return data

def get_prices(data):
    # Getting live prices for each ticker using yahoo_fin package
    print('Getting live prices\n')
    prices = [stock_info.get_live_price(stocks) for stocks in data['tickers']]
    return prices

def update_excel(data):
    # update excel with live prices then close 
    wb = load_workbook(file)
    for index, price in zip(data.index, data['prices']):
        wb[sheet].cell(2+index, 2).value = price
    wb.save(file)
    wb.close()
    print('Price update complete')

def main():
    data_frame = get_tickers()
    data_frame['prices'] = get_prices(data_frame)
    print('Live Prices Report: ',
          datetime.now().strftime("%m/%d/%Y %H:%M:%S"), '\n',
          data_frame)
    # update_excel(data_frame)
    
    
if __name__ == '__main__':
    main()
    
    
    


