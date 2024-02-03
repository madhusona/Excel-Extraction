import os
import pandas as pd

def process_excel_sheets_to_csv(base_folder_path, row_indices, column_indices_map, output_csv):
    # Prepare the header based on provided specifications
    headers = ['Sheet Name', 'Capital Fund 23', 'Capital Fund 22', 'Income 23', 'Income 22', 'Expenses 23', 'Expenses 22']
    all_data = []
    
    for root, dirs, files in os.walk(base_folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                try:
                    xls = pd.ExcelFile(file_path)
                    for sheet_name in xls.sheet_names:
                        sheet_data = [sheet_name]  # Start with the sheet name
                        try:
                            df = pd.read_excel(xls, sheet_name=sheet_name)
                            
                            for row_index in row_indices:
                                # Extract the values for the specified columns for this row
                                cols = column_indices_map.get(row_index, [])
                                if row_index - 1 < len(df):
                                    values = df.iloc[row_index - 1, cols].values
                                    sheet_data.extend(values)  # Add the extracted values to the sheet's data
                        except ValueError as e:
                            print(f"Error reading sheet '{sheet_name}' in file '{file_path}': {e}")
                        all_data.append(sheet_data)  # Append the sheet's data to the overall data list
                except Exception as e:
                    print(f"Error processing file '{file_path}': {e}")
    
    # Convert the collected data into a DataFrame and save it to a CSV file
    result_df = pd.DataFrame(all_data, columns=headers)
    result_df.to_csv(output_csv, index=False)

base_folder_path = '/home/avf/Documents/FAMC/Balance Sheet 22-23/Reviesed Balance Sheet_2022-23_CAG'  # Update with your base folder path
row_indices = [8, 54, 67]  # Rows for "Capital Fund", "Income", and "Expenses"
column_indices_map = {
    8: [5, 6],  # Columns to extract for "Capital Fund"
    54: [5, 6],  # Columns to extract for "Income"
    67: [5, 6],  # Columns to extract for "Expenses"
}
output_csv = 'output_data.csv'

process_excel_sheets_to_csv(base_folder_path, row_indices, column_indices_map, output_csv)

print(f"Data extracted and saved to {output_csv}")
