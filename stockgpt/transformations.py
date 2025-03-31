import pandas as pd

def build_master_df(dms_map, at_set, at_prices, cars_set, cars_prices, pmg_set, pmg_prices):
    all_keys = sorted(set(dms_map) | at_set | cars_set | pmg_set)
    rows = []
    for sn in all_keys:
        row = dict(dms_map.get(sn, {}))
        row.setdefault("Stock Number", sn)
        row["in_dms"] = sn in dms_map
        row["is_on_autotrader"] = "Yes" if sn in at_set else "No"
        row["autotrader_price"] = at_prices.get(sn, "")
        row["is_on_cars"] = "Yes" if sn in cars_set else "No"
        row["cars_price"] = cars_prices.get(sn, "")
        row["is_on_pmgWeb"] = "Yes" if sn in pmg_set else "No"
        row["pmg_web_price"] = pmg_prices.get(sn, "")
        rows.append(row)
    return pd.DataFrame(rows)

def reorder_columns(df):
    front = [
        "Stock Number", "Make", "Model", "Specification", "Colour", "Registration Date", "VIN", "Odometer",
        "Selling Price", "Stand In Value", "Internet Price", "Photo Count", "Stock Days",
        "Location", "Customer Ordered", "Profiles", "in_dms"
    ]
    end = [
        "is_on_cars", "cars_price",
        "is_on_autotrader", "autotrader_price",
        "is_on_pmgWeb", "pmg_web_price"
    ]
    extra = [c for c in df.columns if c not in front + end]
    ordered = front + extra + end
    return df.reindex(columns=[c for c in ordered if c in df.columns])

def split_dms_by_prefix(df, prefix):
    return df[(df["in_dms"]) & (df["Stock Number"].str.startswith(prefix, na=False))].copy()

def generate_site_sheets(df):
    at = df[df["is_on_autotrader"] == "Yes"].copy()
    cars = df[df["is_on_cars"] == "Yes"].copy()
    return at, cars

def generate_to_upload(df):
    mask = (df["in_dms"]) & (
        (df["is_on_autotrader"] == "No") |
        (df["is_on_cars"] == "No") |
        (df["is_on_pmgWeb"] == "No")
    )
    subset = df[mask].copy()
    subset["Note"] = subset.apply(lambda row: describe_missing_sites(row), axis=1)
    subset["Done?"] = ""
    return subset[["Stock Number", "Make", "Model", "Note", "Done?"]]

def describe_missing_sites(row):
    sites = []
    if row.get("is_on_autotrader") == "No":
        sites.append("AutoTrader")
    if row.get("is_on_cars") == "No":
        sites.append("Cars.co.za")
    if row.get("is_on_pmgWeb") == "No":
        sites.append("PMG Web")
    return "Add to " + ", ".join(sites)

def generate_to_remove(df):
    remove_df = df[~df["in_dms"]].copy()
    keep = remove_df[remove_df["Stock Number"].str.startswith("U", na=False)].copy()
    others = remove_df[~remove_df["Stock Number"].str.startswith("U", na=False)].copy()

    keep["Done?"] = ""
    others["Done?"] = ""

    columns = ["Stock Number", "is_on_cars", "cars_price", "is_on_autotrader", "autotrader_price", "is_on_pmgWeb", "Done?"]
    return keep[columns], others[columns]
