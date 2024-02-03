import os
import pandas as pd

def process_excel_sheets_to_csv(base_folder_path, row_indices, output_csv):
    all_data = []
    
    for root, dirs, files in os.walk(base_folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                try:
                    xls = pd.ExcelFile(file_path)
                    for sheet_name in xls.sheet_names:
                        try:
                            df = pd.read_excel(xls, sheet_name=sheet_name)
                            for row_index in row_indices:
                                if row_index - 1 < len(df):
                                    row_data = df.iloc[row_index - 1].astype(str).tolist()  # Convert the row to a list of strings
                                    row_str = ', '.join(row_data)  # Join all row items into a single string
                                    all_data.append([sheet_name, f"Row {row_index}", row_str])
                                    print(all_data)
                        except ValueError as e:
                            print(f"Error reading sheet '{sheet_name}' in file '{file_path}': {e}")
                except Exception as e:
                    print(f"Error processing file '{file_path}': {e}")
    
    # Specify columns for the CSV. Adjust according to your needs.
    columns = ['Sheet Name', 'Row Index', 'Row Data']
    result_df = pd.DataFrame(all_data, columns=columns)
    result_df.to_csv(output_csv, index=False)

base_folder_path = '/home/avf/Documents/FAMC/Balance Sheet 22-23'  # Update with your base folder path
row_indices = [8, 54, 67]
output_csv = 'output_data.csv'

process_excel_sheets_to_csv(base_folder_path, row_indices, output_csv)

print(f"Data extracted and saved to {output_csv}")
