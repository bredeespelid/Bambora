import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# Function to select input files
def select_input_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Input Folder")
    return folder_path

# Function to select output file
def select_output_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(title="Select Output File", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    return file_path

# Main code execution
input_folder = select_input_folder()
if input_folder:
    output_file = select_output_file()
    
    if output_file:
        # List to store data from all files
        data_list = []

        # Loop through all files in the input folder
        for file_name in os.listdir(input_folder):
            if file_name.endswith('.xlsx'):
                file_path = os.path.join(input_folder, file_name)
                
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
    print("Input folder not selected.")
