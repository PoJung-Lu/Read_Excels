"""Analysis functions for high-tech industry data processing."""
from __future__ import annotations

from pathlib import Path
from typing import Optional, Callable
import logging
import pandas as pd
import utils.read_data as read_data
from utils.output_excel import output_as


def analyze_grouped(
    sorted_path: Path,
    group_specs: list[tuple[str, list[str], str]],
    cleaner: Optional[Callable],
    path_output: Optional[Path],
) -> None:
    """
    Reads sorted data Excel file and generates grouped analysis reports.

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
        skipped_sheets = []

        for k, df in zip(keys, values):
            if k == "其他":
                continue
            if cleaner:
                df = cleaner(df.copy())
            # Check if required columns exist in the DataFrame
            if group_col not in df.columns:
                skipped_sheets.append(f"{k} (missing column: {group_col})")
                continue
            missing_cols = [col for col in sum_cols if col not in df.columns]
            if missing_cols:
                skipped_sheets.append(f"{k} (missing columns: {', '.join(missing_cols)})")
                continue
            g = df.groupby([group_col], dropna=False)[sum_cols].sum().reset_index()
            g = g.sort_values(by=sum_cols[::-1], ascending=[False] * len(sum_cols))
            result[k] = g

        # Report skipped sheets
        if skipped_sheets:
            logging.warning(f"⚠️  Skipped sheets for {out_file}:")
            for sheet in skipped_sheets:
                logging.warning(f"   - {sheet}")

        # Only write output if we have results
        if result:
            params = {**base_params, "file_name": out_file}
            output_as(result, params)
        elif not result and not skipped_sheets:
            logging.warning(f"⚠️  No data available for {out_file}")