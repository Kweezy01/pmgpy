#!/usr/bin/env python3
"""
utilities.py

Helper functions for:
 - reading CSV with 'sep=,' checks
 - cleaning DataFrames
 - converting column index to Excel letters
"""

import os
import pandas as pd


def read_csv_with_sep_check(csv_path: str) -> pd.DataFrame:
    """
    Reads a CSV file, skipping the first line if it starts with 'sep='.
    Returns an empty DataFrame if the file is missing.
    """
    if not os.path.isfile(csv_path):
        return pd.DataFrame()

    skip_rows = 0
    with open(csv_path, 'r', encoding='utf-8') as f:
        first_line = f.readline()
        if "sep=" in first_line.lower():
            skip_rows = 1

    return pd.read_csv(csv_path, skiprows=skip_rows)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    1) Drops columns/rows that are entirely NaN
    2) Fills NaN in object (string) columns with ""
    """
    df = df.dropna(axis=1, how='all').dropna(axis=0, how='all')
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].fillna("")
    return df


def col_index_to_excel_col_name(col_index: int) -> str:
    """
    Converts a 0-based column index to Excel column letters:
      0->"A", 1->"B", 25->"Z", 26->"AA", etc.
    """
    col_index += 1  # switch to 1-based
    letters = ""
    while col_index > 0:
        remainder = (col_index - 1) % 26
        letters = chr(65 + remainder) + letters
        col_index = (col_index - 1) // 26
    return letters
