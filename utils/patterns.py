import pandas as pd
import os
import numpy as np
from collections import defaultdict
from pathlib import Path
from typing import Union
from typing import Optional


def process_basic_data_sheet(sheet_name, dataframe, dfs_dict, required_key):
    """
    Process basic data sheet and return a DataFrame.

    Args:
        sheet_name (str): Name of the sheet
        dataframe (pd.DataFrame): Input dataframe to process
        dfs_dict (dict): Dictionary to store processed data
        required_key (str): Required key for processing

    Returns:
        pd.DataFrame: Processed dataframe or None if skip condition met
    """
    if "救災能量" in sheet_name:
        return dfs_dict, None

    j = drop_cells_with_string(
        dataframe,
        ["麥寮", "郭ＸＸ", "（05）693－3143", "範例", "-", "範例：１"],
        except_str=["E-mail"],
    )

    index_staffs = j.index[j.iloc[:, 0].astype(str).str.contains("救災能量")].astype(
        int
    )[0]

    dfs_dict["基本資料"].extend(j.values[0])
    dfs_dict["基本資料內容"].extend(
        j[1:index_staffs].dropna(axis=0, how="all").values.T.tolist()
    )
    dfs_dict["人員編制"].extend(j.values[index_staffs + 2 :, 0])
    dfs_dict["編制數量"].extend(j.values[index_staffs + 2 :, 2].astype(float))

    df = pd.DataFrame.from_dict(dfs_dict, orient="index")
    df = df.transpose()

    return (dfs_dict, df)


def drop_cells_with_string(df, string, except_str=None):
    """
    Replace any cell containing a specific string with NaN (drop cell only).
    """
    # 建立布林遮罩：True 表示該 cell 包含該字串
    contains_A = df.astype(str).apply(
        lambda x: x.str.contains("|".join(string), na=False)
    )
    if except_str:
        contains_B = df.astype(str).apply(
            lambda x: x.str.contains("|".join(except_str), na=False)
        )
        mask = contains_A & ~contains_B
    else:
        mask = contains_A
    # 用 NaN 替換那些 cell
    df_clean = df.mask(mask, None)
    return df_clean


def merge_sheets_by_group(
    dfs: dict,
    required_keys: list = [
        "基本資料",
        "消防車輛設備",
        "其他救災設備",
        "國內證書",
        "國外證書",
        "火災搶救設備",
        "個人防護設備",
        "化災搶救設備",
        "偵檢警報設備",
    ],
) -> dict:
    df_dict = {}
    for k, v in dfs.items():
        if k in required_keys[0]:
            if "救災能量" in k:
                continue
            print("##################################", k)
            columns = v.columns.tolist()
            # v[columns[1]] = v[columns[1]].astype(str)
            # Group by column 0 and column 2 separately
            g1 = v.groupby(columns[0], as_index=False, dropna=False, sort=False).agg(
                {columns[1]: list}
            )
            g2 = v.groupby(columns[2], as_index=False, dropna=False, sort=False).agg(
                {columns[3]: "sum"}
            )
            # Combine results maintaining original column order
            g = pd.concat([g1, g2], axis=1)
            df_dict[required_keys[0]] = g

        elif k in required_keys[1:3]:
            print("##################################", k)
            group_keys = [v.columns[0], v.columns[1], v.columns[2]]
            df_dict[k] = v.groupby(
                group_keys, as_index=False, dropna=False, sort=False
            ).sum()

        elif k in required_keys[3:]:
            print("##################################", k)
            group_keys = [v.columns[0], v.columns[1]]
            df_dict[k] = v.groupby(
                group_keys, as_index=False, dropna=False, sort=False
            ).sum()
        else:
            df_dict[k] = v
    return df_dict


def other_pattern(self, keys, values, pattern):
    """Collect and analysis data from keys and values based on a pattern.
    Args:
        keys (list): List of keys from the DataFrame.
        values (list): List of values corresponding to the keys.
        pattern (str): The pattern to use for reading the data.

    Returns:
        pd.DataFrame: Filtered DataFrame.
    """
    # if 'pattern' in df.columns:
    #     return df[df['pattern'].str.contains(pattern, na=False)]
    # else:
    #     return df[df.apply(lambda row: row.astype(str).str.contains(pattern, na=False).any(), axis=1)]
    if pattern == "stack_all":
        # If the pattern is default, stack all sheet-tables to one DataFrame
        df = self.stack_tables(keys, values)
        return df
    if pattern == "苗栗縣":
        print("Pattern is 苗栗縣")
        drop = values[1].columns.tolist()[1]
        values[1] = values[1].drop(drop, axis=1)
        index_certification = (
            values[1]
            .index[values[1].iloc[:, 0].astype(str).str.contains("其它證書")]
            .astype(int)[0]
            + 1
        )
        df_keys = [
            "基本資料",
            "國內證書",
            "國外證書",
            "國內證照",
            "國外證照",
            "消防車輛設備",
            "火災搶救設備",
            "個人防護設備",
            "化災搶救設備",
            "偵檢警報設備",
            "其他救災設備",
        ]
        n_columns = ["項次", "設備名稱", "數量"]
        df_values0 = process_basic_data_sheet(
            keys[0], values[0], defaultdict(list), []
        )[1]
        title_name = ["國內專業訓練證書(證照類型)"] + values[1].iloc[0].values.tolist()[
            1:
        ]
        df_values1 = pd.DataFrame(values[1][1:index_certification])
        df_values1.columns = title_name
        df_values1 = pd.DataFrame(df_values1.T.groupby(level=0, sort=False).sum().T)
        df_values1.insert(1, "統計人數", np.nan)

        df_values2 = pd.DataFrame(values[1][index_certification:])
        df_values2.columns = ["國外專業訓練證書(證照類型)"] + title_name[1:]
        df_values2 = pd.DataFrame(df_values2.T.groupby(level=0, sort=False).sum()).T
        df_values2.insert(1, "統計人數", np.nan)
        df_values3 = pd.DataFrame(
            np.array(values[2])[:23, :4], columns=["項次", "車輛名稱", "", "數量"]
        )
        df_values4 = pd.DataFrame(np.array(values[2])[:10, 5:8], columns=n_columns)[1:]
        df_values5 = pd.DataFrame(np.array(values[2])[12:22, 5:8], columns=n_columns)[
            1:
        ]
        df_values6 = pd.DataFrame(np.array(values[2])[:12, 9:12], columns=n_columns)[1:]
        df_values7 = pd.DataFrame(np.array(values[2])[12:20, 9:12], columns=n_columns)[
            1:
        ]
        df_values8 = pd.DataFrame(
            np.array(values[2])[:27, 13:17], columns=["項次", "設備名稱", "", "數量"]
        )[1:]
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
        c = 0
        NUM_RE = r"(\d+)"
        for i in df_values:
            for col in i.columns:
                if "數量" in col:
                    i[col] = (
                        pd.DataFrame(i)[col]
                        .astype(str)
                        .str.extract(NUM_RE)[0]
                        .str.replace(",", "", regex=False)
                        .astype(float)
                    )
            print("########", c)
            c = c + 1

        return df_keys, df_values

    elif pattern == "top_ten_operating_chemicals":
        df_keys = []
        df_values = []
        for i, j in zip(keys, values):
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
            return df_keys, df_values
        else:
            stacked_df = pd.concat(df_values)
            stacked_df.columns = value.columns
            drop_if_contains = ["No.", "範例1", "範例2", "公共危險物品運作資料"]
            a = [True if i in drop_if_contains else False for i in stacked_df["No."]]
            stacked_df = stacked_df.drop(stacked_df[a].index)
            stacked_df = stacked_df.drop(columns="No.")
            stacked_df = stacked_df.dropna(axis=0, how="all")

            return df_keys, stacked_df

    elif pattern == "sort_by_location":
        """Sorts the data by location.
        Args:
            keys (list): location name must be included in sheet name i.e. key here.
            values (list): table of specified location.
        """
        print("Pattern is sort_by_location")
        north_tech = ["竹科", "新竹", "龍潭", "竹南", "銅鑼"]
        mid_tech = ["中", "台中", "后里", "虎尾", "二林"]
        south_tech = ["南", "高雄", "楠梓", "嘉義", "樹谷", "路竹"]
        # North -> Mid -> south
        # Although 南 is includede in 竹南, it is sorted already, not affected.
        print(keys)
        if any(
            (location in keys[0]) & (not "路" in keys[0]) for location in north_tech
        ):
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

    elif pattern == "industry_rescue_equipment":
        """
        Args:
            keys (list):  sheet names of excel that is read in.
            values (list):  list of dfs correspond to the sheet.
        """
        dfs = defaultdict(
            list
        )  # all keys:['基本資料', '基本資料內容', '證照','證照數量','演練','演練數量','應變設備','應變設備數量']
        df_keys = []
        df_values = []
        for i, j in zip(keys, values):
            if ("基本資料" in i) & (len(i) < 31):
                sheet_name = i  # j.iloc[1, 6]
                if sheet_name and str(sheet_name) == "nan":
                    sheet_name = j.iloc[1, 6]
                df_keys.append(sheet_name)
                index_title = [[1, 1, 2, 3, 4, 3, 4], [0, 4, 0, 0, 0, 4, 4]]
                index_value = [[1, 1, 2, 3, 4, 3, 4, 5, 6], [1, 6, 1, 1, 1, 6, 6, 2, 2]]

                dfs["基本資料"].extend(j.values[*index_title])
                dfs["基本資料"].append("經度")
                dfs["基本資料"].append("緯度")
                dfs["基本資料內容"].extend(j.values[*index_value])

            elif ("證照及演練" in i) & (len(i) < 31):
                index_training = j.index[
                    j.iloc[:, 1].astype(str).str.contains("消防法演練")
                ].astype(int)[0]
                print("#################", index_training, j.values[index_training, 1])
                index_title = [[i for i in range(1, index_training - 1)], [4]]
                index_value = [[i for i in range(1, index_training - 1)], [7]]
                index_title_train = [
                    [
                        index_training,
                        index_training,
                        index_training + 1,
                        index_training + 1,
                        index_training + 2,
                    ],
                    [1, 5, 1, 5, 1],
                ]
                index_value_train = [
                    [
                        index_training,
                        index_training,
                        index_training + 1,
                        index_training + 1,
                        index_training + 2,
                    ],
                    [4, 7, 4, 7, 4],
                ]
                dfs["證照"].extend(j.values[*index_title])
                dfs["證照數量"].extend(j.values[*index_value])
                dfs["演練"].extend(j.values[*index_title_train])
                dfs["演練數量"].extend(j.values[*index_value_train])

            elif ("應變設備" in i) & (len(i) < 31):
                drop_rows = [
                    (
                        True
                        if ("如欄位不敷使用" in str(i) or ("設備名稱" in str(i)))
                        else False
                    )
                    for i in j.iloc[:, 1]
                ]
                df = j.iloc[:, [1, 6, 7]]
                df = df.drop(df[drop_rows].index)
                df = df.dropna(axis=0, how="all")

                dfs["應變設備"] = df.iloc[:, 0].values
                dfs["應變設備數量"] = df.iloc[:, 1].values
                dfs["應變設備可支援數量"] = df.iloc[:, 2].values
            else:
                pass

        if len(df_keys) != 0:
            df = pd.DataFrame.from_dict(dfs, orient="index")
            df = df.transpose()
            df_values = df.dropna(axis=0, how="all")

            # stacked_df = stacked_df.drop(stacked_df[a].index)
            # print("df_keys", df_keys)
            # print("df_values", df_values)

        else:
            pass
        return df_keys, df_values

    elif pattern == "firefighter_rescue_survey":
        """
        Args:
            keys (list):  sheet names of excel that is read in.
            values (list):  list of dfs correspond to the sheet.

        Returns:
            df_keys (list): list of keys extracted from the sheets.
            df_values (list): list of DataFrames extracted from the sheets.
        """
        NUM_RE = r"(\d+)"
        dfs = defaultdict(list)
        required_keys = [
            "基本資料",
            "消防車輛設備",
            "其他救災設備",
            "國內證書",
            "國外證書",
            "火災搶救設備",
            "個人防護設備",
            "化災搶救設備",
            "偵檢警報設備",
        ]
        typo_map = {
            "國內證照": "國內證書",
            "國外證照": "國外證書",
        }
        df_keys = []
        df_values = []
        for i, j in zip(keys, values):
            i = typo_map.get(i, i)
            if (required_keys[0] in i) & (len(i) < 31):
                dfs, df = process_basic_data_sheet(i, j, dfs, required_keys[0])
                if df is not None:
                    df_keys.append(required_keys[0])
                    df_values.append(df)
            elif any(r_k in i for r_k in required_keys):
                df_keys.append(*[r_k for r_k in required_keys if r_k in i])
                # dfs = {k: v for k, v in j.items() if len(v) > 0}
                j = drop_cells_with_string(j, ["範例", "-", "範例：１"])
                equipment = "設備" in i
                for col in j.columns:
                    if any(
                        k in col
                        for k in ["項次", "訓練證書", "統計人數", "設備", "車輛"]
                    ):
                        j[col] = j[col].ffill()
                    elif equipment:
                        if "數量" in col:
                            j[col] = (
                                pd.DataFrame(j)[col]
                                .astype(str)
                                .str.extract(NUM_RE)[0]
                                .str.replace(",", "", regex=False)
                                .astype(float)
                            )
                        else:
                            j[col] = j[col].astype(str)
                            pass
                    else:
                        # j[col] = j[col].astype(float)
                        j[col] = (
                            pd.DataFrame(j)[col]
                            .astype(str)
                            .str.extract(NUM_RE)[0]
                            .str.replace(",", "", regex=False)
                            .astype(float)
                        )
                df_values.append(j)
            else:
                continue
        return df_keys, df_values
