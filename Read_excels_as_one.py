import os
import pandas as pd
import utils.read_data as read_data
from utils.output_excel import output_as
from collections import defaultdict
from pathlib import Path
import logging
from typing import Union

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# -------------------- 通用工具 --------------------

# NUM_RE = r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][-+]?\d+)?"
NUM_RE = r"(\d+\.?\d*)"
exclude_files = ("Output",)


def extract_first_number(s: pd.Series) -> pd.Series:
    out = (
        s.astype(str).str.extract(NUM_RE)[0].str.replace(",", "", regex=False)
    )  # DataFrame series
    # out = s.astype(str).str.extract(NUM_RE) #.str.replace(",", "", regex=False)#DataFrame
    return pd.to_numeric(out, errors="coerce")


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def list_subfolders(root: Path, exclude=exclude_files) -> list[Path]:
    excl = set(exclude)
    return [p for p in root.iterdir() if p.is_dir() and p.name not in excl]


def concat_list_dict(d: dict[str, list[pd.DataFrame]]) -> dict[str, pd.DataFrame]:
    return {
        k: (pd.concat(v, ignore_index=True) if v else pd.DataFrame())
        for k, v in d.items()
    }


# -------------------- 共用流程 --------------------
def process_folder_tree(base_path: Path, out_root: Path, pattern: str) -> None:
    """
    走訪 base_path 下各子資料夾，讀檔並各自輸出一份合併檔。
    依 read_data API：reader.read_excel_files() / read_with_pattern()
    回傳    (file, (keys, df))
    """
    # ensure_dir(out_root)
    root_reader = read_data.read_data(
        {"path_data": str(base_path), "path_output": str(out_root), "pattern": pattern}
    )
    base = Path(root_reader.get_path())

    for folder in list_subfolders(base):
        logging.info(f"Processing folder: {folder.name}")
        params = {
            "path_data": str(base_path),
            "path_output": str(out_root),
            "folder_path": str(folder),
            "file_name": f"{folder.name}.xlsx",
            "pattern": pattern,
        }
        reader = read_data.read_data(params)
        combined = defaultdict(list)
        iterator = reader.read_excel_files()

        for f, (k, v) in iterator:
            if isinstance(k, (list, tuple)) and len(k) > 0:
                combined[k[0]].append(v)
            elif isinstance(k, str) and k:
                combined[k].append(v)
            else:
                logging.info(f"No data in {f}; skip.")

        combined = {k: pd.concat(v, ignore_index=True) for k, v in combined.items()}

        params["output_path"] = str(base / out_root)
        output_as(combined, params)

    logging.info("All folders processed successfully.")


def sort_by_location(
    sorted_out_name: str, base_for_sorted: Path, pattern="sort_by_location"
) -> Path:
    """
    讀取前一步各資料夾的輸出檔，依園區（北/中/南/其他）彙整成多工作表。
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
    group_specs: [(group_col, sum_cols, out_filename), ...]
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
                continue
            if cleaner:
                df = cleaner(df.copy())
            g = df.groupby([group_col], dropna=False)[sum_cols].sum().reset_index()
            # 依 sum_cols 的逆序排序（和你原本邏輯一致）
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
    """Function to handle high-tech industry data processing."""
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


if __name__ == "__main__":
    # main()
    base = "../Data/科技廠救災能量"
    high_tech_industry_chems_main(base, out_rel="/Output")
    high_tech_industry_rescue_equipment_main(base, out_rel="/Output/Rescue_equipment")
