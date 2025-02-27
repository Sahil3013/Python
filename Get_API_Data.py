import requests
import pandas as pd
import time
from datetime import datetime
import os

# Replace with your actual API Key
API_KEY = "93NBONB3SDGDEDPP"

# API URL
url = f"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=MSFT&apikey={API_KEY}"

# Make API request with error handling
response = requests.get(url)
data = response.json()

# Check if the response contains data
if "Time Series (Daily)" in data:
    # Convert JSON to DataFrame
    df = pd.DataFrame.from_dict(data["Time Series (Daily)"], orient="index")

    # Rename columns for better readability
    df = df.rename(columns={
        "1. open": "Open",
        "2. high": "High",
        "3. low": "Low",
        "4. close": "Close",
        "5. volume": "Volume"
    })

    # Convert index to datetime
    df.index = pd.to_datetime(df.index)
    
    # Rename index to 'Date'
    df.index.name = "Date"

    # Convert numeric columns to float
    df = df.astype(float)

    # Sort by date (newest first)
    df = df.sort_index(ascending=False)

    # Get current date
    current_date = datetime.now().strftime("%Y-%m-%d")
    file_name = f"stock_data_{current_date}.xlsx"

    # Check if file exists
    if os.path.exists(file_name):
        existing_df = pd.read_excel(file_name, index_col=0, parse_dates=True)
        combined_df = pd.concat([df, existing_df]).drop_duplicates().sort_index(ascending=False)
        combined_df.to_excel(file_name)
        print(f"Stock data updated successfully in {file_name}!")
    else:
        df.to_excel(file_name)
        print(f"Stock data saved successfully as {file_name}!")
else:
    print("Error fetching data:", data)
