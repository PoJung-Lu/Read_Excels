"""Data cleaning functions for different data types."""

import pandas as pd

NUM_RE = r"(\d+\.?\d*)"


def extract_first_number(s: pd.Series) -> pd.Series:
    """Extracts the first numeric value from a Series of mixed text/numbers."""
    out = s.astype(str).str.extract(NUM_RE)[0].str.replace(",", "", regex=False)
    return pd.to_numeric(out, errors="coerce")


def clean_chems(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans chemical storage data by standardizing data types.

    Args:
        df: DataFrame containing chemical storage information

    Returns:
        Cleaned DataFrame with proper data types for chemical columns
    """
    for c in ("化學物質名稱", "容器材質", "物質儲存型態"):
        if c in df:
            df[c] = df[c].astype(str)
    for c in ("廠內最大儲存量(公斤)", "廠內最大儲存量(公升)"):
        if c in df:
            df[c] = extract_first_number(df[c])
    return df


def clean_equipment(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans rescue equipment data by standardizing data types.

    Args:
        df: DataFrame containing rescue equipment information

    Returns:
        Cleaned DataFrame with proper data types for equipment columns
    """
    for c in ("證照", "演練", "應變設備"):
        if c in df:
            df[c] = df[c].astype(str)
    for c in ("證照數量", "演練數量", "應變設備數量", "應變設備可支援數量"):
        if c in df:
            df[c] = extract_first_number(df[c])
    return df