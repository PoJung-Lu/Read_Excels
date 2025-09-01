import pandas as pd
import os
import numpy as np



class read_data:
    """
    This class allows you to specify a directory containing Excel files and read them into a pandas DataFrame.
    """

    from .patterns import other_pattern

    def __init__(self, parameters):
        """
        Initializes the read_data class with the given parameters.
        Args:
        """
        import copy

        ##.get("param", "DefaultParam")
        self.parameters = copy.deepcopy(parameters)
        self.parameters["folder_path"] = self.get_path()
        self.parameters["read_all_sheets"] = self.parameters.get(
            "read_all_sheets", False
        )
        self.parameters["pattern"] = self.parameters.get("pattern", "default")

        # self.data = self.read_excel_files()

    def get_path(self):
        """Returns the path where the Excel files are located.
        Returns:
            str: The path to the directory containing the Excel files.
        """
        try:
            import sys

            # This works when the code is executed as a script
            # path =  os.path.dirname(__file__)#os.getcwd() #
            path = os.path.abspath(sys.modules["__main__"].__file__).replace(
                "/Read_excels_as_one.py", ""
            )
            print("Current working directory:", path)
        except NameError:
            # This works in Jupyter Notebooks
            from pathlib import Path

            path = Path(globals()["_dh"][0])
            path = str(path)
        # except AttributeError:
        return (
            path + "/Data"
            if not ("path_data" in self.parameters)
            else path + "/" + self.parameters["path_data"]
        )

    def stack_tables(self, keys, values):
        """Stacks the tables from the keys and values into a single DataFrame.
        Args:
            keys (list): List of keys from the DataFrame.
            values (list): List of values corresponding to the keys.
        Returns:
            pd.DataFrame: A DataFrame containing the stacked data.
        """
        value = []
        print(f"Keys: {keys}")
        for k, v in zip(keys, values):
            i, j, k = np.array(v.columns.values), np.array(v.values), np.array(k)
            value.extend(k.reshape(1, 1).tolist())
            value.extend(i.reshape(1, len(i)).tolist())
            value.extend(j.tolist())
        df = pd.DataFrame(value)
        df = df.dropna(axis=0, how="all")  # Drop rows where all elements are NaN
        df = df.reset_index(drop=True)  # Reset the index
        return df

    def read_with_pattern(self, keys, values, pattern):
        """Reads the pattern from the parameters and returns the corresponding value.
        Args:
            keys (list): List of keys from the DataFrame.
            values (list): List of values corresponding to the keys.
            pattern (str): The pattern to use for reading the data.
        """
        pattern = self.parameters["pattern"]
        if pattern == "default":
            # If the pattern is default, stack all sheet-tables to one DataFrame
            df = self.stack_tables(keys, values)
            return df
        else:
            return self.other_pattern(keys, values, pattern)

    def read_one_excel(self, file_path):
        read_all_sheets = self.parameters["read_all_sheets"]
        sheet_names = self.parameters.get("sheet_names", None)
        sheet_name = (
            None if read_all_sheets else sheet_names
        )  # Read all sheets if not specified
        df = pd.read_excel(file_path, sheet_name=sheet_name, thousands=",")
        df_keys = [i for i, j in df.items()]
        # df_values = [j.dropna(axis=0, how="all") for i, j in df.items()]
        df_values = [j for i, j in df.items()]
        return df_keys, df_values

    def read_excel_files(self):
        """Reads all Excel files in the specified directory and returns a DataFrame.
        Returns:
        file: file name string
        pd.DataFrame: A DataFrame containing the combined data from all Excel files, each sheet corresponds to one file.
        """
        folder_path = self.parameters["folder_path"]
        pattern = self.parameters["pattern"]
        files = os.listdir(folder_path)
        print(f"Folder path: {folder_path}")

        if self.parameters["file_name"] in files:
            files.remove(
                self.parameters["file_name"]
            )  # Remove the Output file if it exists

        excel_files = [
            f for f in files if f.endswith((".xlsx", ".xls", ".xlsm", "ods"))
        ]
        dfs = {}
        for file in excel_files:
            print(f"Reading file: {file}")
            file_path = os.path.join(folder_path, file)
            df_keys, df_values = self.read_one_excel(file_path)
            # dfs[file] = self.read_with_pattern(df_keys, df_values, pattern)
            yield file, self.read_with_pattern(df_keys, df_values, pattern)
        # dfs = pd.DataFrame(dfs) # Print sheet names if reading all sheets
        # return dfs  # Return the dictionary of DataFrames
