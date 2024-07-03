import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# Function to select input files
def select_input_files():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="Select Input Files", filetypes=[("Excel files", "*.xlsx")])
    return file_paths

# Function to select output file
def select_output_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(title="Select Output File", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    return file_path

# Main code execution
input_files = select_input_files()
if input_files:
    output_file = select_output_file()
    
    if output_file:
        # List to store data from all files
        data_list = []

        # Loop through all selected files
        for file_path in input_files:
            if file_path.endswith('.xlsx'):
                # Read the Adjustments sheet
                df = pd.read_excel(file_path, sheet_name='Adjustments')
                
                # Convert DATE column to dd.mm.yyyy format
                df['DATE'] = pd.to_datetime(df['DATE']).dt.strftime('%d.%m.%Y')
                
                # Take the absolute value of the AMOUNT column
                df['AMOUNT'] = df['AMOUNT'].abs()
                
                # Append necessary columns to the data_list
                data_list.append(df[['DATE', 'MERCHANT ID', 'AMOUNT']])

        # Concatenate all data into a single DataFrame
        all_data = pd.concat(data_list)

        # Group by DATE and MERCHANT ID, and sum the AMOUNT
        summary = all_data.groupby(['DATE', 'MERCHANT ID'], as_index=False)['AMOUNT'].sum()

        # Rename columns to the desired format and reorder columns
        summary.columns = ['Dato', 'Merchant ID', 'Amount']
        summary = summary[['Dato', 'Amount', 'Merchant ID']]

        # Save the summary DataFrame to an Excel file
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        summary.to_excel(output_file, index=False)

        print(f"Summary file saved to {output_file}")
    else:
        print("Output file not selected.")
else:
    print("Input files not selected.")
