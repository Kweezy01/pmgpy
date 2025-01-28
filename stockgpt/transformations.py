# transformations.py

import pandas as pd
import os

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

DEALER_PREFIXES = {
    "Ford_Nelspruit":   "UF",
    "Ford_Mazda":       "UG",
    "Produkta_Nissan":  "UA",
    "Suzuki_Nelspruit": "UE",  # forced is_on_cars="Yes"
    "Ford_Malalane":    "US",
}

def build_master_df(dms_map: dict,
                    at_set, at_prices,
                    cars_set, cars_prices,
                    pmg_set, pmg_prices):
    """
    Union of DMS + website sets => DataFrame
    If prefix=UE => force is_on_cars="Yes"
    """
    all_sn = set(dms_map.keys()) | at_set | cars_set | pmg_set
    rows = []
    for sn in sorted(all_sn):
        row = {}
        if sn in dms_map:
            row.update(dms_map[sn])
            row["in_dms"] = True
        else:
            for c in PINNACLE_SCHEMA:
                row[c] = ""
            row["in_dms"] = False

        # Cars
        row["is_on_cars"] = "Yes" if sn in cars_set else "No"
        row["cars_price"] = cars_prices.get(sn,"")

        # AutoTrader
        row["is_on_autotrader"] = "Yes" if sn in at_set else "No"
        row["autotrader_price"] = at_prices.get(sn,"")

        # pmgWeb
        row["is_on_pmgWeb"] = "Yes" if sn in pmg_set else "No"
        row["pmg_web_price"] = pmg_prices.get(sn,"")

        # Force is_on_cars if prefix=UE
        if sn.startswith("UE"):
            row["is_on_cars"] = "Yes"

        row["Stock Number"] = sn
        rows.append(row)

    df = pd.DataFrame(rows)
    return df


def reorder_final_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Put Pinnacle columns up front, then site columns.
    """
    front_cols = [
        "Stock Number","Make","Model","Specification","Colour",
        "Registration Date","VIN","Odometer","Selling Price",
        "Stand In Value","Date In Stock","Original Group Date In Stock",
        "Stock Days","Branch","Location","Body Style","Fuel Type",
        "Transmission","Customer Ordered","Profiles","in_dms"
    ]
    end_cols = [
        "Photo Count","Internet Price",
        "is_on_cars","cars_price",
        "is_on_autotrader","autotrader_price",
        "is_on_pmgWeb","pmg_web_price"
    ]
    existing = df.columns.tolist()
    final = [c for c in front_cols if c in existing] + [c for c in end_cols if c in existing]
    return df.reindex(columns=final)


def generate_todos(df_in, df_removed):
    """
    Auto-gen tasks:
     - If PhotoCount>1 & any site=No => "Need to fix listing"
     - If website-only => "Remove from site"
    """
    tasks = []
    if not df_in.empty and "Photo Count" in df_in.columns:
        for _, row in df_in.iterrows():
            pc = row.get("Photo Count","0")
            try:
                pc_val = float(pc)
            except:
                pc_val = 0
            if pc_val>1:
                # check sites
                for sitecol in ["is_on_cars","is_on_autotrader","is_on_pmgWeb"]:
                    if row.get(sitecol,"No")=="No":
                        tasks.append({
                            "Task":"Need to fix listing",
                            "Stock Number": row["Stock Number"],
                            "Notes":f"PhotoCount>1 but missing on {sitecol}"
                        })
                        break

    # website-only => remove
    for _, row in df_removed.iterrows():
        tasks.append({
            "Task":"Remove from site",
            "Stock Number": row["Stock Number"],
            "Notes":"Website-only"
        })

    df_todo = pd.DataFrame(tasks, columns=["Task","Stock Number","Notes"])
    return df_todo
