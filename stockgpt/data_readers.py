import os
import pandas as pd
from .utilities import read_csv_with_sep_check

def read_autotrader_data(folder):
    stks, prices = set(), {}
    for sub in os.listdir(folder):
        subdir = os.path.join(folder, sub)
        if not os.path.isdir(subdir): continue
        file = os.path.join(subdir, "autotrader.csv")
        if not os.path.isfile(file): continue
        df = read_csv_with_sep_check(file)
        df.columns = df.columns.str.strip()
        if "StockNumber" in df.columns:
            df.rename(columns={"StockNumber": "Stock Number"}, inplace=True)
        price_col = next((c for c in df.columns if "price" in c.lower()), None)
        for _, row in df.iterrows():
            sn = str(row.get("Stock Number", "")).strip()
            if sn:
                stks.add(sn)
                prices[sn] = str(row.get(price_col, "")).strip() if price_col else ""
    return stks, prices

def read_cars_data(folder):
    stks, prices = set(), {}
    for sub in os.listdir(folder):
        subdir = os.path.join(folder, sub)
        if not os.path.isdir(subdir): continue
        file = os.path.join(subdir, "cars.xlsx")
        if not os.path.isfile(file): continue
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        if "Reference" in df.columns:
            df.rename(columns={"Reference": "Stock Number"}, inplace=True)
        price_col = next((c for c in df.columns if "price" in c.lower()), None)
        for _, row in df.iterrows():
            sn = str(row.get("Stock Number", "")).strip()
            if sn:
                stks.add(sn)
                prices[sn] = str(row.get(price_col, "")).strip() if price_col else ""
    return stks, prices

def read_pmg_web_data(folder):
    file = os.path.join(folder, "pmg_web_data.csv")
    if not os.path.isfile(file): return set(), {}
    df = read_csv_with_sep_check(file)
    df.columns = df.columns.str.strip()
    if "SKU" in df.columns:
        df.rename(columns={"SKU": "Stock Number"}, inplace=True)
    price_col = next((c for c in df.columns if "price" in c.lower()), None)
    stks, prices = set(), {}
    for _, row in df.iterrows():
        sn = str(row.get("Stock Number", "")).strip()
        if sn:
            stks.add(sn)
            prices[sn] = str(row.get(price_col, "")).strip() if price_col else ""
    return stks, prices

def read_dms_csv(file):
    if not os.path.isfile(file):
        print(f"[WARN] DMS file not found: {file}")
        return {}
    df = pd.read_csv(file)
    if "Customer Order" in df.columns:
        df.rename(columns={"Customer Order": "Customer Ordered"}, inplace=True)
    df = df.fillna("").astype(str)
    for col in df.columns:
        df[col] = df[col].str.strip()
    dms_map = {}
    for _, row in df.iterrows():
        sn = row.get("Stock Number", "")
        if sn:
            dms_map[sn] = row.to_dict()
    return dms_map


def read_all_sources(src="src"):
    print("[INFO] Reading DMS data...")
    dms_map = read_dms_csv(os.path.join(src, "pmg_dms_data.csv"))
    print(f"[INFO] DMS cars loaded: {len(dms_map)}")
    print("[INFO] Reading website data...")
    at, at_prices = read_autotrader_data(src)
    cars, cars_prices = read_cars_data(src)
    pmg, pmg_prices = read_pmg_web_data(src)
    return dms_map, at, at_prices, cars, cars_prices, pmg, pmg_prices
