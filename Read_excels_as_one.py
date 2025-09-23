import os
import pandas as pd
import numpy as np
from utils.patterns import merge_sheets_by_group
import utils.read_data as read_data
from utils.output_excel import output_as
from collections import defaultdict
from pathlib import Path
import logging
from typing import Union
from typing import Optional

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# -------------------- 通用工具 --------------------
# NUM_RE = r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][-+]?\d+)?"
NUM_RE = r"(\d+\.?\d*)"
exclude_files = ("Output", "Distribution_by_city")


def extract_first_number(s: pd.Series) -> pd.Series:
    out = (
        s.astype(str).str.extract(NUM_RE)[0].str.replace(",", "", regex=False)
    )  # DataFrame series
    # out = s.astype(str).str.extract(NUM_RE) #.str.replace(",", "", regex=False)#DataFrame
    return pd.to_numeric(out, errors="coerce")


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def list_subfolders(
    root: Path, exclude=exclude_files, specify_folders: Union[list[str], None] = None
) -> list[Path]:
    excl = set(exclude)
    if specify_folders:
        return [
            p
            for p in root.iterdir()
            if p.is_dir()
            and (p.name not in excl)
            and (i in p.name for i in specify_folders)
        ]
    else:
        return [p for p in root.iterdir() if p.is_dir() and (p.name not in excl)]


def concat_list_dict(d: dict[str, list[pd.DataFrame]]) -> dict[str, pd.DataFrame]:
    return {
        k: (pd.concat(v, ignore_index=True) if v else pd.DataFrame())
        for k, v in d.items()
    }


# -------------------- 共用流程 --------------------
def process_folder_tree(
    base_path: Path, out_root: Path, pattern: str, filename: Optional[str] = None
) -> None:
    """
    走訪 base_path 下各子資料夾，讀檔並各自輸出一份合併檔。
    依 read_data API：reader.read_excel_files() / read_with_pattern()
    回傳    (file, (keys, df))

    Args:
        base_path: Root directory containing subdirectories with Excel files
        out_root: Output directory path for processed files
        pattern: Processing pattern to apply (e.g., 'top_ten_operating_chemicals')
        filename: Optional prefix for output filenames

    Process:
        1. Iterates through each subdirectory in base_path
        2. Reads Excel files using specified pattern
        3. Combines data from multiple sheets/files
        4. Outputs consolidated Excel file for each subdirectory
    """
    # ensure_dir(out_root)
    root_reader = read_data.read_data(
        {"path_data": str(base_path), "path_output": str(out_root), "pattern": pattern}
    )
    base = Path(root_reader.get_path())

    for folder in list_subfolders(base):
        logging.info(f"Processing folder: {folder.name}")
        file_name = (
            f"{folder.name}.xlsx"
            if not filename
            else filename + "_" + f"{folder.name}.xlsx"
        )
        params = {
            "path_data": str(base_path),
            "path_output": str(out_root),
            "folder_path": str(folder),
            "file_name": file_name,
            "pattern": pattern,
        }
        reader = read_data.read_data(params)
        combined = defaultdict(list)
        iterator = reader.read_excel_files()

        for f, (k, v) in iterator:
            if isinstance(k, (list, tuple)) and len(k) > 1:
                [combined[i].append(j.dropna(axis=0, how="all")) for i, j in zip(k, v)]
            elif isinstance(k, (list, tuple)) and len(k) == 1:
                combined[k[0]].append(v[0])
            elif isinstance(k, str) and k:
                combined[k].append(v)
            else:
                logging.info(f"No data in {f}; skip.")
        combined = {k: pd.concat(v, ignore_index=True) for k, v in combined.items()}

        if len(combined.keys()) >= 1:
            combined = merge_sheets_by_group(combined)
        params["output_path"] = str(base / out_root)
        output_as(combined, params)
    logging.info("All folders processed successfully.")


def sort_by_location(
    sorted_out_name: str, base_for_sorted: Path, pattern="sort_by_location"
) -> Path:
    """
    讀取process_folder_tree的輸出檔，依園區（北/中/南/其他）彙整成多工作表。

    Args:
        sorted_out_name: Name for the consolidated output file
        base_for_sorted: Directory containing files to be sorted by location
        pattern: Processing pattern (default: 'sort_by_location')

    Returns:
        Path to the created sorted output file

    Process:
        1. Reads previously processed files from subdirectories
        2. Categorizes data by Taiwan regions (北部園區/中部園區/南部園區/其他)
        3. Combines data for each region into separate sheets
        4. Outputs multi-sheet Excel file organized by region
    """
    root_reader = read_data.read_data(
        {
            "path_data": str(base_for_sorted),
            "path_output": str(base_for_sorted),
            "pattern": pattern,
        }
    )
    base = Path(root_reader.get_path())
    params = {
        "path_data": str(base_for_sorted),
        "path_output": str(base_for_sorted),
        "pattern": pattern,
        "file_name": sorted_out_name,
        "folder_path": str(base),
        "output_path": str(base),
    }
    reader = read_data.read_data(params)
    buckets = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}
    factories = {k: [] for k in buckets}

    for f, (region, dfs) in reader.read_excel_files():
        logging.info(f"Sorting file: {f} -> region = {region}")
        buckets[region].extend(dfs)
        factories[region].append(f)

    merged = concat_list_dict(buckets)
    output_as(merged, params)
    return Path(params["output_path"]) / sorted_out_name


def analyze_grouped(
    sorted_path: Path,
    group_specs: list[tuple[str, list[str], str]],
    cleaner: Union[callable, None],
    path_output: Union[Path, None],
) -> None:
    """
    從 Sorted_data.xlsx 讀出各工作表，依 group_specs 做 groupby-sum-sort，分別輸出。
    Args:
        sorted_path: Path to the sorted data Excel file
        group_specs: List of tuples specifying grouping operations:
                    [(group_column, columns_to_sum, output_filename), ...]
        cleaner: Optional cleaning function to apply to data before grouping
        path_output: Output directory path (defaults to sorted_path parent)

    Process:
        1. Reads all sheets from sorted Excel file
        2. Applies cleaning function if provided
        3. For each grouping specification:
           - Groups data by specified column
           - Sums the specified numeric columns
           - Sorts by summed values in descending order
           - Outputs to separate Excel file
    """
    base_params = {
        "path_data": str(sorted_path.parent.parent),
        "path_output": str(path_output or sorted_path.parent),
        "read_all_sheets": True,
        "folder_path": str(sorted_path.parent),
        "output_path": str(sorted_path.parent),
    }
    reader = read_data.read_data(base_params)
    keys, values = reader.read_one_excel(str(sorted_path))

    for group_col, sum_cols, out_file in group_specs:
        result = {}
        for k, df in zip(keys, values):
            if k == "其他":
                pass
            if cleaner:
                df = cleaner(df.copy())
            g = df.groupby([group_col], dropna=False)[sum_cols].sum().reset_index()
            # 依 sum_cols 的逆序排序（和原本邏輯一致）
            g = g.sort_values(by=sum_cols[::-1], ascending=[False] * len(sum_cols))
            result[k] = g
        params = {**base_params, "file_name": out_file}
        output_as(result, params)


# -------------------- 資料類型清理 --------------------


def clean_chems(df: pd.DataFrame) -> pd.DataFrame:
    for c in ("化學物質名稱", "容器材質", "物質儲存型態"):
        if c in df:
            df[c] = df[c].astype(str)
    for c in ("廠內最大儲存量(公斤)", "廠內最大儲存量(公升)"):
        if c in df:
            df[c] = extract_first_number(df[c])
    return df


def clean_equipment(df: pd.DataFrame) -> pd.DataFrame:
    for c in ("證照", "演練", "應變設備"):
        if c in df:
            df[c] = df[c].astype(str)
    for c in ("證照數量", "演練數量", "應變設備數量", "應變設備可支援數量"):
        if c in df:
            df[c] = extract_first_number(df[c])
    return df


# -------------------- 主流程 --------------------
def main():
    # Create an instance of the read_data class
    parameters = {
        "path_data": "../Data/科技廠救災能量",  # Specify the path to your Excel files
        "pattern": "default",  # You can specify a pattern if needed, ['default', '苗栗縣', top_ten_operating_chemicals, sort_by_location, industry_rescue_equipment]
        "path_output": "/Output",  # Specify the output folder
        "file_name": "Aggregated_data.xlsx",  # Specify the output file name
    }
    data_reader = read_data.read_data(parameters)
    # Access the combined DataFrame
    path = data_reader.get_path()
    combined_data = data_reader.read_excel_files()
    output_as(combined_data, data_reader.parameters)


def high_tech_industry_chems_main(base="../Data/科技廠救災能量", out_rel="/Output"):
    """
    Function to handle high-tech industry chemical storage data processing.

    Args:
        base: Base directory containing industrial chemical data
        out_rel: Relative output directory path

    Workflow:
        1. Process each company folder using 'top_ten_operating_chemicals' pattern
        2. Sort and aggregate data by Taiwan industrial park regions
        3. Generate analysis reports grouped by:
           - Chemical names with storage quantities
           - Container materials with storage volumes
           - Storage states (solid/liquid/gas) with quantities

    Output:
        Creates multiple Excel files with aggregated chemical inventory data
    """
    base_path = Path(base)
    out_root = out_rel.strip("/")
    # 1) 逐資料夾處理
    process_folder_tree(
        base_path=base_path, out_root=out_root, pattern="top_ten_operating_chemicals"
    )
    # 2) 依園區彙整
    base_for_sorted = base_path / out_root
    sorted_path = sort_by_location(
        "Sorted_data.xlsx", base_for_sorted, pattern="sort_by_location"
    )
    # 3) 分析輸出
    specs = [
        (
            "化學物質名稱",
            ["廠內最大儲存量(公斤)", "廠內最大儲存量(公升)"],
            "sort_by_hazmat.xlsx",
        ),
        (
            "容器材質",
            ["廠內最大儲存量(公斤)", "廠內最大儲存量(公升)"],
            "sort_by_container.xlsx",
        ),
        (
            "物質儲存型態",
            ["廠內最大儲存量(公斤)", "廠內最大儲存量(公升)"],
            "sort_by_state.xlsx",
        ),
    ]
    path_output = base_path / out_root
    analyze_grouped(sorted_path, specs, cleaner=clean_chems, path_output=path_output)


def high_tech_industry_rescue_equipment_main(
    base="../Data/科技廠救災能量", out_rel="/Output/Rescue_equipment"
):
    """
    Function to handle high-tech industry rescue equipment data processing.

    Args:
        base: Base directory containing industrial rescue equipment data
        out_rel: Relative output directory path for equipment reports

    Workflow:
        1. Process each facility folder using 'industry_rescue_equipment' pattern
        2. Sort and aggregate data by Taiwan industrial park regions
        3. Generate analysis reports grouped by:
           - Certifications and license counts
           - Training exercises and participation numbers
           - Emergency response equipment and support capacities

    Output:
        Creates categorized Excel files for equipment inventory analysis
    """
    base_path = Path(base)
    out_root = out_rel.strip("/")

    # 1) 逐資料夾處理（讀取模式不同）
    process_folder_tree(
        base_path, out_root=out_root, pattern="industry_rescue_equipment"
    )

    # 2) 依園區彙整
    base_for_sorted = base_path / out_root
    sorted_path = sort_by_location(
        "Sorted_data.xlsx", base_for_sorted, pattern="sort_by_location"
    )

    # 3) 分析輸出
    specs = [
        ("證照", ["證照數量"], "sort_by_certificate.xlsx"),
        ("演練", ["演練數量"], "sort_by_training.xlsx"),
        ("應變設備", ["應變設備數量", "應變設備可支援數量"], "sort_by_equipment.xlsx"),
    ]
    path_output = base_path / out_root
    analyze_grouped(
        sorted_path, specs, cleaner=clean_equipment, path_output=path_output
    )


def analyze_ff_survey_files(
    base_path: Path,
    group_specs: list,
    out_root: Union[Path, None],
    pattern: str = "default",
    filename: Optional[str] = None,
) -> None:
    """ """
    file_name = filename or "Grouped_data.xlsx"
    combined = {}
    root_data = Path(str(base_path) + str(out_root.parent))
    print("root_data", root_data)
    valid_column = [
        "大隊長",
        "副大隊長",
        "組長",
        "主任",
        "中隊長",
        "副中隊長",
        "分隊長",
        "組員",
        "小隊長",
        "隊員",
        "科員",
    ]
    for folder in list_subfolders(root_data):

        logging.info(f"Processing folder: {folder.name}")
        process_file = f"/{folder.name}.xlsx"
        params = {
            "path_data": str(base_path),
            "folder_path": str(folder),
            "file_name": file_name,
            "pattern": pattern,
        }
        reader = read_data.read_data(params)

        iterator = reader.read_excel_files()

        agg_k, agg_v = reader.read_one_excel(
            str(base_path) + str(out_root) + process_file
        )
        # city_files = [i for i in iterator]+[(process_file, (agg_k, agg_v))]
        cert_dict_division = []  # defaultdict(list)
        for f, (k, v) in iterator:  # division, sheet_names, dfs
            f = Path(f).name.replace(".xlsx", "")
            if isinstance(k, (list, tuple)) and len(k) > 0:
                df_list = []
                for i, j in zip(k, v):
                    if ("基本資料" in i) and ("救災能量" not in i):
                        dd = {i: [j] for i, j in zip(j["人員編制"], j["編制數量"])}
                        df_dict = pd.DataFrame(dd)
                        df_dict.index = pd.MultiIndex.from_tuples([(f, "編制數量")])
                        # (f, "編制數量")
                        df_list.append(df_dict)
                        # print(df_dict)
                    elif "證書" in i:
                        for spec_i in group_specs:
                            first_col = j.columns[0]
                            mask = j[first_col].str.contains(
                                spec_i, case=False, na=False
                            )
                            df = pd.DataFrame(j.loc[mask].iloc[:, 2:].sum()).T
                            df.index = pd.MultiIndex.from_tuples([(f, spec_i)])
                            df_list.append(df)
                    else:
                        continue

                cert_dict_division.append(pd.concat(df_list))
            else:
                logging.info(f"No data in {f}; skip.")
        dfs = pd.concat(cert_dict_division).reset_index().fillna(0)
        print()
        print(dfs.columns[:2])
        print(dfs)

        dfs = (
            dfs.groupby(dfs.columns[:2].to_list(), dropna=False, sort=False)
            .sum()
            .reset_index()
        )
        columns = [i for i in dfs.columns.to_list() if i in valid_column]
        dfs_sum = dfs.groupby(dfs.columns[1], dropna=False, sort=False)[columns].sum()
        dfs_sum.index = pd.MultiIndex.from_tuples(
            [("彙整", i) for i in (["編制數量"] + group_specs)]
        )
        dfs_sum.loc[("彙整", "未受訓"), :] = (
            dfs_sum.loc[("彙整", "編制數量"), :]
            - dfs_sum.loc[("彙整", "化災搶救基礎班"), :]
        )
        df = pd.concat([dfs, dfs_sum.reset_index()])
        df["總計"] = df.iloc[:, 2:].sum(axis=1)
        combined[folder.name] = df
        print(folder.name, combined[folder.name])
    params = {
        "path_data": str(base_path),
        "file_name": file_name,
        "pattern": pattern,
        "output_path": str(base_path) + str(out_root),
        "folder_path": str(root_data),
    }
    output_as(combined, params)
    logging.info("All folders processed successfully.")


def firefighter_training_survey_main(
    base="../Data/消防機關救災能量", out_rel="../Output"
):
    """
    Function to handle firefighters' rescue capability data processing.

    Args:
        base: Base directory containing firefighter survey data
        out_rel: Relative output directory path

    Expected folder structure:
        base_path/cities/Division/files.xlsx

    Workflow:
        1. Process firefighter data by city and division
        2. Aggregate training levels (基礎班/進階班/指揮官班/教官班)
        3. Compile personnel counts and equipment inventories
        4. Generate city-level distribution reports

    Output:
        Creates consolidated reports showing firefighter capabilities by region
    """
    base = Path(base)
    out_root = out_rel.strip("/")
    root_reader = read_data.read_data(
        {"path_data": str(base), "path_output": str(out_root), "pattern": ""}
    )
    base_path = Path(root_reader.get_path())
    # 1) 逐大隊資料夾處理
    for cities in list_subfolders(base_path, exclude=exclude_files + ("Raw_data",)):

        if "苗栗縣" in str(cities):
            pattern = "苗栗縣"
        else:
            pattern = "firefighter_rescue_survey"
        logging.info(f"Processing city folder: {cities.name}")
        city_path = base / cities.name
        city_out_root = out_root + f"/{cities.name}"
        process_folder_tree(
            base_path=city_path,
            out_root=city_out_root,
            pattern=pattern,
            filename=cities.name,
        )
    # 2) 逐縣市資料夾處理
    out_root = "Distribution_by_city"
    path_output = out_root
    base_path = base / Path("Output")
    process_folder_tree(
        base_path=base_path,
        out_root=out_root,
        pattern="default",
    )
    # 3) 依縣市彙整
    specs = [
        "化災搶救基礎班",
        "化災搶救進階班",
        "化災搶救指揮官班",
        "化災搶救教官班",
        # "通識級",
        # "操作級",
        # "技術級",
        # "指揮級",
        # "專家級",
    ]

    out_root = Path("/Output/Distribution_by_city")
    base_path = Path(root_reader.get_path())
    # base_path = Path(str(base_path) + str(out_root.parent))
    analyze_ff_survey_files(
        base_path,
        specs,
        out_root=out_root,
        pattern="default",
        filename="Grouped_data.xlsx",
    )


if __name__ == "__main__":

    # main()
    # base = "../Data/消防機關救災能量"  # "../Data/科技廠救災能量"  #"../Test"  #
    # root_out = "/../Output"
    base = "../Test"  # "../Data/消防機關救災能量"  # "../Data/科技廠救災能量"  #
    root_out = "/../Output"  #
    # high_tech_industry_chems_main(base, out_rel=root_out)
    # high_tech_industry_rescue_equipment_main(
    #     base, out_rel=root_out + "/Rescue_equipment"
    # )
    firefighter_training_survey_main(base, root_out)
