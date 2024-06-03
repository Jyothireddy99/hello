import pandas as pd
import os
from concurrent.futures import ThreadPoolExecutor

# Folder path containing the XLSX files
folder_path = r'10.10.10.11'

def convert_xlsx_to_csv(filename):
    try:
        # Full file path
        file_path = os.path.join(folder_path, filename)
        
        # Read the Excel file
        excel_data = pd.read_excel(file_path, engine='openpyxl')
        
        # Define the output CSV file path
        csv_file_path = os.path.join(folder_path, filename.replace('.xlsx', '.csv'))
        
        # Save the DataFrame to a CSV file
        excel_data.to_csv(csv_file_path, index=False)
        
        return f'Success: Converted {filename} to CSV format.'
    except Exception as e:
        return f'Error: Failed to convert {filename}. Reason: {str(e)}'

def main():
    # List all XLSX files in the folder
    xlsx_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    # Use ThreadPoolExecutor to process files concurrently
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(convert_xlsx_to_csv, xlsx_files))
    
    # Print results
    for result in results:
        print(result)

if __name__ == "__main__":
    main()

print('All XLSX files have been converted to CSV format.')
