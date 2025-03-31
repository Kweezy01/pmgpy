import pandas as pd

def read_csv_with_sep_check(filepath):
    try:
        df = pd.read_csv(filepath)
        if df.columns[0].startswith("sep="):
            df = pd.read_csv(filepath, sep=";")
    except Exception:
        df = pd.read_csv(filepath, sep=";")
    return df

def clean_dataframe(df):
    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    return df.reset_index(drop=True)
