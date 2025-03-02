import requests
import pandas as pd
import time
from openpyxl import Workbook

# Function to fetch top 50 cryptocurrencies data from CoinGecko API
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return None

# Function to perform data analysis
def analyze_data(df):
    top_5_by_market_cap = df.nlargest(5, 'Market Cap')[['Cryptocurrency', 'Market Cap']]
    avg_price = df['Current Price'].mean()
    highest_price_change = df.loc[df['24h Change %'].idxmax(), ['Cryptocurrency', '24h Change %']]
    lowest_price_change = df.loc[df['24h Change %'].idxmin(), ['Cryptocurrency', '24h Change %']]
    
    return top_5_by_market_cap, avg_price, highest_price_change, lowest_price_change

# Function to update Excel file
def update_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Crypto Live Data"

    # Writing headers
    headers = ["Cryptocurrency", "Symbol", "Current Price (USD)", "Market Cap", "24h Volume", "24h Change %"]
    ws.append(headers)

    while True:
        data = fetch_crypto_data()
        if data:
            df = pd.DataFrame(data)
            df = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
            df.columns = ["Cryptocurrency", "Symbol", "Current Price", "Market Cap", "24h Volume", "24h Change %"]

            # Clear previous data
            ws.delete_rows(2, ws.max_row)

            # Append new data
            for row in df.itertuples(index=False, name=None):
                ws.append(row)

            # Save the file
            wb.save("crypto_live_data.xlsx")

            # Perform data analysis
            top_5, avg_price, highest, lowest = analyze_data(df)
            print("\nTop 5 Cryptos by Market Cap:\n", top_5)
            print(f"\nAverage Price of Top 50 Cryptos: ${avg_price:.2f}")
            print(f"\nHighest 24h Change: {highest.Cryptocurrency} ({highest['24h Change %']}%)")
            print(f"Lowest 24h Change: {lowest.Cryptocurrency} ({lowest['24h Change %']}%)")

        # Wait for 5 minutes before updating again
        time.sleep(300)

# Run the update function
update_excel()
