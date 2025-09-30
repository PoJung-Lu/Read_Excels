"""Analysis functions for high-tech industry data processing."""

from pathlib import Path
from typing import Optional, Callable
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
        for k, df in zip(keys, values):
            if k == "其他":
                pass
            if cleaner:
                df = cleaner(df.copy())
            g = df.groupby([group_col], dropna=False)[sum_cols].sum().reset_index()
            g = g.sort_values(by=sum_cols[::-1], ascending=[False] * len(sum_cols))
            result[k] = g
        params = {**base_params, "file_name": out_file}
        output_as(result, params)