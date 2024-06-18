# import requests
# import pandas as pd
# from datetime import datetime, timedelta
# from requests.auth import HTTPBasicAuth
# from dateutil import parser

# # API URL and credentials
# api_url = "https://sheet2api.com/v1/QnhwtLLOQVev/test-python"
# username = "collexe-test"
# password = "123456"

# # Fetch data from the API
# response = requests.get(api_url, auth=HTTPBasicAuth(username, password))
# data = response.json()

# # Convert the data to a DataFrame
# df = pd.DataFrame(data)

# # Parse 'Date' to datetime using dateutil's parser
# df['Date'] = df['Date'].apply(parser.parse)

# # Sort the DataFrame by 'Date'
# df = df.sort_values(by='Date')

# # Calculate the rolling average for Qty and Amount for the last 3 days
# df['Qty Avg'] = df['Qty'].rolling(window=3).mean()
# df['Amount Avg'] = df['Amount'].rolling(window=3).mean()

# # Fill NaN values in averages with 0 (for the first two days where rolling window is not complete)
# df['Qty Avg'].fillna(0, inplace=True)
# df['Amount Avg'].fillna(0, inplace=True)

# # Add 'API Access Date' column
# df['API Access Date'] = datetime.now().strftime('%d/%m/%Y')

# # Reformat 'Date' column back to string in the desired format
# df['Date'] = df['Date'].dt.strftime('%Y%m%d')

# # Define the order of columns
# columns_order = ['Date', 'API Access Date', 'Qty', 'Amount', 'Qty Avg', 'Amount Avg']

# # Reorder the DataFrame
# df = df[columns_order]

# # Save the DataFrame to an Excel file
# df.to_excel("api_data.xlsx", index=False)
# print(df)

#=============================== part2====================================

# import requests
# import pandas as pd
# from datetime import datetime, timedelta
# from requests.auth import HTTPBasicAuth
# from dateutil import parser

# # API URL and credentials
# api_url = "https://sheet2api.com/v1/QnhwtLLOQVev/test-python"
# username = "collexe-test"
# password = "123456"

# # Fetch data from the API
# response = requests.get(api_url, auth=HTTPBasicAuth(username, password))
# data = response.json()

# # Convert the data to a DataFrame
# df = pd.DataFrame(data)

# # Parse 'Date' to datetime using dateutil's parser
# df['Date'] = df['Date'].apply(parser.parse)

# # Sort the DataFrame by 'Date'
# df = df.sort_values(by='Date')

# # Calculate the rolling values for the last 3 days for detailed output
# df['Qty-2'] = df['Qty'].shift(2)
# df['Qty-1'] = df['Qty'].shift(1)
# df['Qty-0'] = df['Qty']
# df['Amount-2'] = df['Amount'].shift(2)
# df['Amount-1'] = df['Amount'].shift(1)
# df['Amount-0'] = df['Amount']

# # Calculate the rolling average for Qty and Amount for the last 3 days
# df['Qty Avg'] = df[['Qty-2', 'Qty-1', 'Qty-0']].mean(axis=1)
# df['Amount Avg'] = df[['Amount-2', 'Amount-1', 'Amount-0']].mean(axis=1)

# # Fill NaN values in averages and detailed columns with 0 (for the first two days where rolling window is not complete)
# df[['Qty-2', 'Qty-1', 'Qty-0', 'Qty Avg', 'Amount-2', 'Amount-1', 'Amount-0', 'Amount Avg']] = df[['Qty-2', 'Qty-1', 'Qty-0', 'Qty Avg', 'Amount-2', 'Amount-1', 'Amount-0', 'Amount Avg']].fillna(0)

# # Add 'API Access Date' column
# df['API Access Date'] = datetime.now().strftime('%d/%m/%Y')

# # Reformat 'Date' column back to string in the desired format
# df['Date'] = df['Date'].dt.strftime('%Y%m%d')

# # Define the order of columns
# columns_order = ['Date', 'API Access Date', 'Qty', 'Amount', 'Qty-2', 'Qty-1', 'Qty-0', 'Qty Avg', 'Amount-2', 'Amount-1', 'Amount-0', 'Amount Avg']

# # Reorder the DataFrame
# df = df[columns_order]

# # Save the DataFrame to an Excel file
# df.to_excel("api_data_with_details.xlsx", index=False)

# =================================== part 2 ================================

# =================================== part 3 ============================

import requests
import pandas as pd
from datetime import datetime, timedelta
from requests.auth import HTTPBasicAuth
from dateutil import parser

# API URL and credentials
api_url = "https://sheet2api.com/v1/QnhwtLLOQVev/test-python"
username = "collexe-test"
password = "123456"

# Fetch data from the API
response = requests.get(api_url, auth=HTTPBasicAuth(username, password))
data = response.json()

# Convert the data to a DataFrame
df = pd.DataFrame(data)

# Parse 'Date' to datetime using dateutil's parser
df['Date'] = df['Date'].apply(parser.parse)

# Sort the DataFrame by 'Date'
df = df.sort_values(by='Date')

# Calculate the rolling values for the last 3 days for detailed output
df['Qty-2'] = df['Qty'].shift(2)
df['Qty-1'] = df['Qty'].shift(1)

df['Amount-2'] = df['Amount'].shift(2)
df['Amount-1'] = df['Amount'].shift(1)

# Calculate the rolling average for Qty and Amount for the last 3 days
df['Qty Avg'] = df[['Qty-2', 'Qty-1', 'Qty']].mean(axis=1)
df['Amount Avg'] = df[['Amount-2', 'Amount-1', 'Amount']].mean(axis=1)

# Fill NaN values in averages and detailed columns with 0 (for the first two days where rolling window is not complete)
df[['Qty-2', 'Qty-1', 'Qty Avg', 'Amount-2', 'Amount-1', 'Amount Avg']] = df[['Qty-2', 'Qty-1', 'Qty Avg', 'Amount-2', 'Amount-1', 'Amount Avg']].fillna(0)

# Add 'API Access Date' column
df['API Access Date'] = datetime.now().strftime('%d/%m/%Y')

# Reformat 'Date' column back to string in the desired format
df['Date'] = df['Date'].dt.strftime('%Y%m%d')

# Calculate the sum and average of Qty for the entire dataset
total_qty = df['Qty'].sum()
average_qty = df['Qty'].mean()

# Append the sum and average to the DataFrame
sum_avg_df = pd.DataFrame({
    'Date': [''],
    'Qty-2': [0],
    'Qty-1': [0],
    'Qty Avg': [average_qty],
    'Qty': [total_qty],
    'Amount-2': [0],
    'Amount-1': [0],
    'Amount': [0],
    'Amount Avg': [0],
    'API Access Date': ['']
})

df = pd.concat([df, sum_avg_df], ignore_index=True)

# Define the order of columns
columns_order = ['Date', 'Qty-2', 'Qty-1', 'Qty Avg', 'Qty', 'Amount-2', 'Amount-1', 'Amount', 'Amount Avg', 'API Access Date']

# Reorder the DataFrame
df = df[columns_order]

# Save the DataFrame to an Excel file
df.to_excel("api_data_with_sum_avg.xlsx", index=False)
