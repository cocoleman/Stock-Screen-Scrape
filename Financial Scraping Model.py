import yfinance as yf
import pandas as pd
import openpyxl
import requests
from bs4 import BeautifulSoup

def fetch_financial_data(symbol):
    stock = yf.Ticker(symbol)
    data = stock.history(period='1d')
    info = stock.info

    # Extract data through yfinance database
    stock_name = info.get('longName', 'N/A')
    name = {stock_name}
    current_price = data['Close'].iloc[-1] if not data.empty else 'N/A'
    eps = info.get('trailingEps', 'N/A')
    sps = get_revenue_per_share(symbol)
    dividend = info.get('dividendRate', '0')
    beta = info.get('beta', 'N/A')
    growth_estimate = get_next_five_years_growth_estimate(symbol)

    return current_price, eps, sps, dividend, beta, growth_estimate

# Webscrape for values not in yfinance

def get_next_five_years_growth_estimate(symbol):
    url = f"https://finance.yahoo.com/quote/{symbol}/analysis"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')

        for table in soup.select("table"):
            th_row = [th.text for th in table.find_all("th")]
            if 'Growth Estimates' in th_row:
                for tr in table.select("tr:has(td)"):
                    td_row = [td.text for td in tr.find_all("td")]
                    if td_row[0] == 'Next 5 Years (per annum)':
                        return td_row[1]  # The growth estimate value is in the second column

    return "Data not found."

def get_revenue_per_share(symbol):
    url = f"https://finance.yahoo.com/quote/{symbol}/key-statistics"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        text_elements = soup.find_all(text=lambda text: "Revenue Per Share" in text if text else False)
        for element in text_elements:
            # Attempt to find the the value manually
            parent = element.parent
            while parent and parent.name != 'td':
                parent = parent.parent
            if parent:
                value_td = parent.find_next_sibling("td")
                if value_td:
                    return value_td.text

    return "Data not found."

    


# Load the workbook and the specific sheet you want to work with (UPDATE THIS!!!)
wb = openpyxl.load_workbook(r"C:\Users\chc53\Downloads\Excel Scrape\Coleman Screener Scrape.xlsm")
sheet = wb['Sheet1']  # Replace 'SheetName' with the name of your sheet

# Input columns and row that tickers are in (numerical)
start_column = 4
end_column = 8  
row_number = 9

# Open a new file to write the output in CSV format (this will create a new file if not already made)
with open("output.csv", "w") as file:
    # Assuming your tickers are in the range defined above
    start_column = 4
    end_column = 8
    row_number = 9
    for col in range(start_column, end_column + 1):
        symbol = sheet.cell(row=row_number, column=col).value
        if symbol:
            # Fetch the financial data for the symbol
            data = fetch_financial_data(symbol)
            # Write the data in CSV format
            file.write(','.join(map(str, data)) + '\n')

wb.close()