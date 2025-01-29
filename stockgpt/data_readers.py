# data_readers.py

import os
import pandas as pd

from utilities import read_csv_with_sep_check, clean_dataframe

PINNACLE_SCHEMA = [
    "Stock Number",
    "Make",
    "Model",
    "Specification",
    "Colour",
    "Registration Date",
    "VIN",
    "Odometer",
    "Photo Count",
    "Selling Price",
    "Stand In Value",
    "Internet Price",
    "Date In Stock",
    "Original Group Date In Stock",
    "Stock Days",
    "Branch",
    "Location",
    "Body Style",
    "Fuel Type",
    "Transmission",
    "Customer Ordered",
    "Profiles",
]

def read_dms_dict(dms_file: str) -> dict:
    """Reads pmg_dms_data.csv -> {stockNum -> rowData} with only relevant columns."""
    if not os.path.isfile(dms_file):
        print(f"[WARN] Missing DMS file: {dms_file}")
        return {}

    # Try reading first few rows and check structure
    print(f"[DEBUG] Checking DMS file structure: {dms_file}")
    df_preview = pd.read_csv(dms_file, nrows=5)  # Read first 5 rows only
    print(df_preview)

    # Actual reading logic
    df = pd.read_csv(dms_file)
    print(f"[DEBUG] DMS Data Loaded: {df.shape}")  # Print shape

    if "Stock Number" not in df.columns:
        print("[ERROR] 'Stock Number' column is missing in DMS file! Check column names.")
        print(f"Available columns: {df.columns.tolist()}")
        return {}

    # Convert to dictionary (Stock Number as key)
    dms_map = {}
    for _, row in df.iterrows():
        sn = row["Stock Number"]
        if sn:
            dms_map[sn] = dict(row)

    return dms_map



def read_autotrader_data(src_folder: str):
    """
    For each subfolder in src_folder => 'autotrader.csv'
    builds set_of_stocknums, dict_of_prices
    """
    import pandas as pd

    stks, prices = set(), {}
    for sub in os.listdir(src_folder):
        sub_path = os.path.join(src_folder, sub)
        if not os.path.isdir(sub_path):
            continue
        csv_file = os.path.join(sub_path, "autotrader.csv")
        if not os.path.isfile(csv_file):
            continue

        df = read_csv_with_sep_check(csv_file)
        df.columns = df.columns.str.strip()
        if "StockNumber" in df.columns:
            df.rename(columns={"StockNumber": "Stock Number"}, inplace=True)

        price_col = None
        for c in ["PriceFormatted","Price","price"]:
            if c in df.columns:
                price_col = c
                break

        if "Stock Number" in df.columns:
            for _, row in df.iterrows():
                sn = str(row["Stock Number"]).strip()
                if sn:
                    stks.add(sn)
                    p = str(row[price_col]).strip() if price_col else ""
                    prices[sn] = p
    return stks, prices


def read_cars_data(src_folder: str):
    stks, prices = set(), {}
    for sub in os.listdir(src_folder):
        sub_path = os.path.join(src_folder, sub)
        if not os.path.isdir(sub_path):
            continue

        xlsx_path = os.path.join(sub_path, "cars.xlsx")
        if not os.path.isfile(xlsx_path):
            continue

        df = pd.read_excel(xlsx_path)
        df.columns = df.columns.str.strip()
        if "Reference" in df.columns:
            df.rename(columns={"Reference": "Stock Number"}, inplace=True)

        price_col = None
        for c in ["Price","price"]:
            if c in df.columns:
                price_col = c
                break

        if "Stock Number" in df.columns:
            for _, row in df.iterrows():
                sn = str(row["Stock Number"]).strip()
                if sn:
                    stks.add(sn)
                    prices[sn] = str(row[price_col]).strip() if price_col else ""
    return stks, prices


def read_pmg_web_data(src_folder: str):
    path = os.path.join(src_folder, "pmg_web_data.csv")
    if not os.path.isfile(path):
        return set(), {}

    df = read_csv_with_sep_check(path)
    df.columns = df.columns.str.strip()
    if "SKU" in df.columns:
        df.rename(columns={"SKU":"Stock Number"}, inplace=True)

    price_col = None
    for c in ["Regular price","Regular Price","price","Price"]:
        if c in df.columns:
            price_col = c
            break

    stks, prices = set(), {}
    if "Stock Number" in df.columns:
        for _, row in df.iterrows():
            sn = str(row["Stock Number"]).strip()
            if sn:
                stks.add(sn)
                p = str(row[price_col]).strip() if price_col else ""
                prices[sn] = p

    return stks, prices
