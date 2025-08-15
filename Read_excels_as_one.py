import os
import pandas as pd
import utils.read_data as read_data
from utils.output_excel import output_as 


parameters = {
    'read_all_sheets': True,  # Set to True if you want to read all sheets in each Excel file
    'path_data': 'Data',  # Specify the path to your Excel files
    'pattern': '苗栗縣',  # You can specify a pattern if needed, ['default', '苗栗縣']
    'output_path': 'Output',  # Specify the output path
    'file_name': 'Aggregated_data.xlsx'  # Specify the output file name
}

def main():
    # Create an instance of the read_data class
    data_reader = read_data.read_data(parameters=parameters)
    
    # Access the combined DataFrame
    path = data_reader.get_path()
    combined_data = data_reader.read_excel_files()
    output_as(combined_data, data_reader.parameters)
    
    # for file, df in combined_data.items():
    #     print(f"Data from {file}:")
    #     print(df.head(5))  # Display the first few rows of each sheet
    #     print(df.keys())

if __name__=="__main__":
    main()  