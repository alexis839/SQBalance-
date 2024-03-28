import requests
import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# The API token
token = "Token goes here"

# Prepare the headers with the API token
headers = {'Authorization': f'Bearer {token}'}

# The URL for getting balance information
Balance_url = "https://portal.squaretalk.com/api/balance"

# Make a GET request to the API with the headers
Balance_response = requests.get(Balance_url, headers=headers)

# Check if the request was successful
if Balance_response.status_code == 200:
    # Convert the response to JSON
    Balance_response_json = Balance_response.json()
    
    # Extract the 'data' list from the JSON response then Convert the list of dictionaries to a pandas DataFrame
    data_list = Balance_response_json['data']
    df = pd.DataFrame(data_list)
    
    
    # Rename the columns to 'Accounts' and 'Balance'
    df = df.rename(columns={'account_name': 'Accounts', 'balance': 'Balance'})
    
    # Specify the directory where you want to save the Excel file
    directory = 'SquareTalk_Balance_Report'
    
    # Ensure the directory exists
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    # Get the current date and time prepair the date and then set date on Excel file Name 
    now = datetime.now()
    date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f'{date_time_str}_Balance.xlsx'
    
    # Write the DataFrame to an Excel file in the specified directory with the new filename
    df.to_excel(f'{directory}/{filename}', index=False)
    
    # Display a message box with the success message
    root = tk.Tk()
    root.withdraw() # Hide the main window
    messagebox.showinfo("Success", "Excel file created successfully")
    root.destroy() # Close the Tkinter window
    print("Success", f"Excel file created successfully: {directory}/{filename}")
else:
    root = tk.Tk()
    root.withdraw() # Hide the main window
    messagebox.showerror("Error", "Failed to fetch data")
    root.destroy() # Close the Tkinter window
    print("Error", f"Failed to fetch data. Status code: {Balance_response.status_code}")
