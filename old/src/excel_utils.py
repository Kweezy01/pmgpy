# src/excel_utils.py
import pandas as pd

def read_excel_to_dataframe(file_path: str) -> pd.DataFrame:
    return pd.read_excel(file_path, engine='openpyxl')

def write_dataframe_to_excel(df: pd.DataFrame, file_path: str) -> None:
    df.to_excel(file_path, index=False)
