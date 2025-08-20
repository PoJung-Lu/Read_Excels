import pandas as pd
import os
import numpy as np


def other_pattern(self, keys, values, pattern):
    """Collect and analysis data from keys and values based on a pattern.

    Args:


    Returns:
        pd.DataFrame: Filtered DataFrame.
    """
    # if 'pattern' in df.columns:
    #     return df[df['pattern'].str.contains(pattern, na=False)]
    # else:
    #     return df[df.apply(lambda row: row.astype(str).str.contains(pattern, na=False).any(), axis=1)]
    if pattern == "苗栗縣":
        print("Pattern is 苗栗縣")
        df_keys = [
            "基本資料",
            "國內證書",
            "國外證書",
            "消防車輛設備",
            "火災搶救設備",
            "個人防護設備",
            "化災搶救設備",
            "偵檢警報設備",
            "其他救災設備",
        ]
        df_values0 = values[0]
        df_values1 = values[1][:25]
        df_values2 = values[1][25:]
        df_values3 = pd.DataFrame(np.array(values[2])[:23, :4])
        df_values4 = pd.DataFrame(np.array(values[2])[:10, 5:8])
        df_values5 = pd.DataFrame(np.array(values[2])[12:22, 5:8])
        df_values6 = pd.DataFrame(np.array(values[2])[:12, 9:12])
        df_values7 = pd.DataFrame(np.array(values[2])[12:20, 9:12])
        df_values8 = pd.DataFrame(np.array(values[2])[:27, 13:17])
        df_values = [
            df_values0,
            df_values1,
            df_values2,
            df_values3,
            df_values4,
            df_values5,
            df_values6,
            df_values7,
            df_values8,
        ]
        return df_keys, df_values
    elif pattern == "top_ten_operating_chemicals":

        folder_path = self.parameters["folder_path"]
        read_all_sheets = self.parameters["read_all_sheets"]
        pattern = self.parameters["pattern"]
        files = os.listdir(folder_path)
        print(f"Folder path: {folder_path}")
        excel_files = [f for f in files if f.endswith((".xlsx", ".xls", "ods"))]
        dfs = {}
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(
                file_path,
                sheet_name=None if read_all_sheets else self.parameters["sheet_names"],
            )  # Read all sheets if specified
            print(f"Reading file: {file}")
            df_keys = []
            df_values = []

            for i, j in df.items():
                if ("廠場達管制量30倍" in i) & (len(i) < 31):
                    sheet_name = j.iloc[1, 1]
                    if sheet_name and str(sheet_name) == "nan":
                        sheet_name = j.iloc[1, 2]
                    df_keys.append(sheet_name)

                if ("公共危險物品運作資料" in i) & (len(i) < 31):

                    value = j.dropna(axis=0, how="all")
                    column_names = [
                        "No.",
                        "危險品分類",
                        "化學物質名稱",
                        "濃度 (w/w %)",
                        "物質儲存型態",
                        "容器材質",
                        "管制量倍數",
                        "長(公分)",
                        "寬(公分)",
                        "高(公分)",
                        "直徑(公分)",
                        "單一容器最大存量(公斤)",
                        "單一容器最大存量(公升)",
                        "廠內最大儲存量(公斤)",
                        "廠內最大儲存量(公升)",
                        "壓力容器罐裝錶壓(kg/cm2)",
                        "存放位置",
                    ]
                    len_old_column = len(value.iloc[0])
                    if len_old_column > 17:
                        value.columns = column_names + list(value.iloc[0][17:])
                    elif len_old_column == 17:
                        value.columns = column_names
                    else:
                        value.columns = column_names[:16]
                    df_values.append(value)
                else:
                    pass

            if len(df_keys) == 0:
                print(f"No data found in {file}. Skipping.")
            else:
                stacked_df = self.stack_tables(df_keys, df_values)
                stacked_df.columns = value.columns
                drop_if_contains = ["No.", "範例1", "範例2", "公共危險物品運作資料"]
                a = [
                    True if i in drop_if_contains else False for i in stacked_df["No."]
                ]
                stacked_df = stacked_df.drop(stacked_df[a].index)
                stacked_df = stacked_df.drop(columns="No.")
                stacked_df = stacked_df.dropna(axis=0, how="all")

                dfs[df_keys[0]] = stacked_df
        return dfs  # Return the dictionary of DataFrames

    elif pattern == "sort_by_location":
        """Sorts the data by location.
        Args:
            keys (list): location name must be included in sheet name i.e. key here.
            values (list): table of specified location.
        """
        print("Pattern is sort_by_location")
        north_tech = ["新竹", "龍潭", "竹南", "銅鑼"]
        mid_tech = ["中", "台中", "后里", "虎尾", "二林"]
        south_tech = ["南", "高雄", "楠梓", "嘉義", "樹谷", "路竹"]
        print(keys)
        if any(location in keys[0] for location in north_tech):
            print("Location is in 北部園區")
            region = "北部園區"
        elif any(location in keys[0] for location in mid_tech):
            print("Location is in 中部園區")
            region = "中部園區"

        elif any(location in keys[0] for location in south_tech):
            print("Location is in 南部園區")
            region = "南部園區"
        else:
            print("Location not found in any region")
            region = "其他"
        return region, values


def analysis_pattern(self, keys, values, pattern):
    """Reads the pattern from the parameters and returns the corresponding value.
    Args:
        keys (list): List of keys from the DataFrame.
        values (list): List of values corresponding to the keys.
        pattern (str): The pattern to use for reading the data.
    """


# elif pattern == "analysis": #read with pattern analysis
