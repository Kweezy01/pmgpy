# src/transformations.py
import pandas as pd

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # Example transformation
    df = df.dropna()
    df = df.drop_duplicates()
    return df

def add_calculated_column(df: pd.DataFrame, col_name: str, formula) -> pd.DataFrame:
    df[col_name] = df.apply(formula, axis=1)
    return df
