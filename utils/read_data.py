import pandas as pd
import os   
import numpy as np
class read_data: 
    """
    This class allows you to specify a directory containing Excel files and read them into a pandas DataFrame.
    """

    def __init__(self, parameters):
        """
        Initializes the read_data class with the given parameters.   
        Args:
        """
        ##.get("param", "DefaultParam")
        self.parameters = parameters
        self.parameters['folder_path'] = self.get_path()
        self.parameters['read_all_sheets'] = self.parameters.get('read_all_sheets', False)
        self.parameters['pattern'] = self.parameters.get('pattern', 'default')
        
        self.data = self.read_excel_files()


    def get_path(self):
        """        Returns the path where the Excel files are located.        
        Returns:
            str: The path to the directory containing the Excel files.
        """
        try:
            import sys
            # This works when the code is executed as a script
            # path =  os.path.dirname(__file__)#os.getcwd() #
            path =  os.path.abspath(sys.modules['__main__'].__file__).replace('/Read_excels_as_one.py',"")
            print('Current working directory:', path)
        except NameError:
            # This works in Jupyter Notebooks
            from pathlib import Path
            path = Path(globals()['_dh'][0])    
            path = str(path)
        return path + '/Data' if 'path_data' not in self.parameters else path + '/' + self.parameters['path_data']
    
    def stack_tables(self, keys, values):
        """        Stacks the tables from the keys and values into a single DataFrame.        
        Args:
            keys (list): List of keys from the DataFrame.
            values (list): List of values corresponding to the keys.
        Returns:
            pd.DataFrame: A DataFrame containing the stacked data.
        """
        value = []
        for k, v in zip(keys,values):                
            i, j, k= np.array(v.columns.values), np.array(v.values), np.array(k)
            value.extend(k.reshape(1, 1).tolist())
            value.extend(i.reshape(1, len(i)).tolist())
            value.extend(j.tolist())
        df = pd.DataFrame(value)   
        df = df.dropna(axis=0, how='all')  # Drop rows where all elements are NaN
        df = df.reset_index(drop=True)  # Reset the index 
        return df
    

    def read_with_pattern(self, keys, values, pattern):
        """        Reads the pattern from the parameters and returns the corresponding value.        
        Args:
            keys (list): List of keys from the DataFrame.
            values (list): List of values corresponding to the keys.
            pattern (str): The pattern to use for reading the data.
        """
        pattern = self.parameters['pattern']
        if pattern == 'default':
            print('Pattern is default')
            # If the pattern is default, stack all sheet-tables to one DataFrame
            df = self.stack_tables(keys, values)
            return df
        
        elif pattern == '苗栗縣':
            print('Pattern is 苗栗縣')
            df_keys = ['基本資料', '國內證書', '國外證書', '消防車輛設備', '火災搶救設備', 
            '個人防護設備', '化災搶救設備','偵檢警報設備', '其他救災設備']
            df_values0 = values[0]
            df_values1 = values[1][:25]
            df_values2 = values[1][25:]
            df_values3 = pd.DataFrame(np.array(values[2])[:23,:4])
            df_values4 = pd.DataFrame(np.array(values[2])[:10,5:8])
            df_values5 = pd.DataFrame(np.array(values[2])[12:22,5:8])
            df_values6 = pd.DataFrame(np.array(values[2])[:12,9:12])
            df_values7 = pd.DataFrame(np.array(values[2])[12:20,9:12])
            df_values8 = pd.DataFrame(np.array(values[2])[:27,13:17])
            df_values = [df_values0, df_values1, df_values2, df_values3, df_values4, 
                        df_values5, df_values6, df_values7, df_values8]
            df = self.stack_tables(df_keys, df_values)
            return df



    def read_excel_files(self):
        """        Reads all Excel files in the specified directory and returns a DataFrame.        Returns:
            pd.DataFrame: A DataFrame containing the combined data from all Excel files.
        """
        folder_path = self.parameters['folder_path']
        read_all_sheets = self.parameters['read_all_sheets']
        pattern = self.parameters['pattern']
        files = os.listdir(folder_path)
        excel_files = [f for f in files if f.endswith(('.xlsx', '.xls','ods'))]
        dfs = {}
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path,sheet_name = None if read_all_sheets else 0)  # Read all sheets if specified
            df_keys = [i for i,j in df.items()]
            df_values = [j.dropna(axis=0, how='all') for i,j in df.items()]
            
            dfs[file] = self.read_with_pattern(df_keys, df_values, pattern)
        # dfs = pd.DataFrame(dfs) # Print sheet names if reading all sheets
        return dfs  # Return the dictionary of DataFrames

