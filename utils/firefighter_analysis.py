"""Analysis functions for firefighter rescue capability data processing."""

from pathlib import Path
from typing import Optional, Union
import logging
import pandas as pd
import utils.read_data as read_data
from utils.output_excel import output_as

exclude_files = ("Output", "Distribution_by_city")


def list_subfolders(root: Path, exclude: tuple = exclude_files) -> list[Path]:
    """Lists subdirectories excluding specified folder names."""
    excl = set(exclude)
    return [p for p in root.iterdir() if p.is_dir() and p.name not in excl]


def analyze_ff_survey_files(
    base_path: Path,
    group_specs: list,
    out_root: Optional[Path],
    pattern: str = "default",
    filename: Optional[str] = None,
) -> None:
    """
    Analyzes firefighter survey files by aggregating personnel composition and training certification data across divisions.

    Processes division-level Excel files, extracts personnel structure (編制數量) and training certifications,
    calculates regional summaries with training percentages, and outputs multi-sheet Excel files organized by city/region.

    Args:
        base_path: Root directory containing firefighter survey data
        group_specs: List of training course names to aggregate (e.g., ['化災搶救基礎班', '化災搶救進階班'])
        out_root: Output directory path relative to base_path
        pattern: Processing pattern for reading Excel files (default: 'default')
        filename: Output filename (default: 'Grouped_data.xlsx')
    """
    file_name = filename or "Grouped_data.xlsx"
    combined = {}
    skipped_folders = []
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
        "科長",
        "科員",
        "",
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
        cert_dict_division = []
        for f, (k, v) in iterator:
            f = Path(f).name.replace(".xlsx", "")
            if isinstance(k, (list, tuple)) and len(k) > 0:
                df_list = []
                for i, j in zip(k, v):
                    if ("基本資料" in i) and ("救災能量" not in i):
                        valid_set = set(valid_column)
                        dd = {role: [count] for role, count in zip(j["人員編制"], j["編制數量"])}
                        dd = {k: dd.get(k, [0]) for k in valid_column} | {k: v for k, v in dd.items() if k not in valid_set}
                        df_dict = pd.DataFrame(dd)
                        df_dict.index = pd.MultiIndex.from_tuples([(f, "編制數量")])
                        df_list.append(df_dict)
                    elif "證書" in i:
                        for spec_i in group_specs:
                            first_col = j.columns[0]
                            val_columns = [
                                col for col in j.columns if col in valid_column
                            ]
                            mask = j[first_col].str.contains(
                                spec_i, case=False, na=False
                            )
                            df = pd.DataFrame(
                                j.loc[mask].iloc[:, 2:][val_columns].sum()
                            ).T
                            df.index = pd.MultiIndex.from_tuples(
                                [(f, spec_i)], names=["單位", "課程"]
                            )
                            df_list.append(df)
                if df_list:
                    cert_dict_division.append(pd.concat(df_list))
                else:
                    logging.warning(f"No matching data found in {f}; skipping.")
            else:
                logging.info(f"No data in {f}; skip.")

        if not cert_dict_division:
            skipped_folders.append(f"{folder.name} (no valid data found)")
            continue

        dfs = pd.concat(cert_dict_division).reset_index().fillna(0)
        dfs = dfs.groupby(dfs.columns[:2].to_list(), dropna=False, sort=False).sum()
        columns = [i for i in dfs.columns.to_list() if i in valid_column]
        dfs_sum = dfs.groupby(level="level_1", dropna=False, sort=False)[columns].sum()
        dfs_sum.index = pd.MultiIndex.from_tuples(
            [("彙整", i) for i in (["編制數量"] + group_specs)], names=["單位", "課程"]
        )
        training_classes = [
            "化災搶救基礎班",
            "化災搶救進階班",
            "化災搶救教官班",
            "化災搶救指揮官班",
        ]
        dfs_sum.loc[("彙整", "未受訓"), :] = dfs_sum.loc[("彙整", "編制數量"), :] - sum(
            dfs_sum.loc[("彙整", cls), :] for cls in training_classes
        )
        dfs_sum.loc[("彙整", "未受訓"), :] = dfs_sum.loc[("彙整", "未受訓"), :].clip(lower=0)
        df = pd.concat([dfs, dfs_sum])
        df["總計"] = df.sum(axis=1)
        df["比例"] = (
            df["總計"]
            .div(df.loc[("彙整", "編制數量"), "總計"])
            .round(3)
            .map("{:.2%}".format)
        )
        combined[folder.name] = df.reset_index()

    # Report skipped folders
    if skipped_folders:
        logging.warning(f"⚠️  Skipped folders in firefighter analysis:")
        for folder in skipped_folders:
            logging.warning(f"   - {folder}")

    # Only write output if we have data
    if combined:
        params = {
            "path_data": str(base_path),
            "file_name": file_name,
            "pattern": pattern,
            "output_path": str(base_path) + str(out_root),
            "folder_path": str(root_data),
        }
        output_as(combined, params)
        logging.info("All folders processed successfully.")
    else:
        logging.warning("⚠️  No data available to write - all folders were skipped or empty")
