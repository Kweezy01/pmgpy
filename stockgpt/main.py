#!/usr/bin/env python3
import os
import pandas as pd
import warnings
from .utilities import clean_dataframe
from .data_readers import read_all_sources
from .transformations import (
    build_master_df, reorder_columns, split_dms_by_prefix, generate_site_sheets,
    generate_to_upload, generate_to_remove
)
from .formatting import style_sheet, auto_size_columns, generate_corporate_report

warnings.simplefilter("ignore", UserWarning)

OUTPUT_FILE = "output/stockgpt.xlsx"
DEALER_PREFIXES = {
    "Ford_Nelspruit":   "UF",
    "Ford_Mazda":       "UG",
    "Produkta_Nissan":  "UA",
    "Suzuki_Nelspruit": "UE",
    "Ford_Malalane":    "US",
}

def main():
    os.makedirs("output", exist_ok=True)

    dms_map, at_set, at_prices, cars_set, cars_prices, pmg_set, pmg_prices = read_all_sources()
    df_master = reorder_columns(build_master_df(dms_map, at_set, at_prices, cars_set, cars_prices, pmg_set, pmg_prices))

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for name, prefix in DEALER_PREFIXES.items():
            df = split_dms_by_prefix(df_master, prefix)
            if not df.empty:
                df.drop(columns=["Date In Stock", "Branch", "Body Style", "Transmission", "Fuel Type"], errors="ignore", inplace=True)
                sheet_name = f"DMS_{name}"
                df.to_excel(excel_writer=writer, sheet_name=sheet_name, index=False)
                style_sheet(writer, sheet_name, df)

        at_df, cars_df = generate_site_sheets(df_master)
        at_df.to_excel(excel_writer=writer, sheet_name="AutoTrader_Listings", index=False)
        cars_df.to_excel(excel_writer=writer, sheet_name="Cars_Listings", index=False)
        style_sheet(writer, "AutoTrader_Listings", at_df)
        style_sheet(writer, "Cars_Listings", cars_df)

        df_upload = generate_to_upload(df_master)
        df_remove, df_remove_others = generate_to_remove(df_master)

        df_upload.to_excel(excel_writer=writer, sheet_name="To_Upload", index=False)
        df_remove.to_excel(excel_writer=writer, sheet_name="To_Remove", index=False)
        df_remove_others.to_excel(excel_writer=writer, sheet_name="to_remove_others", index=False)

        style_sheet(writer, "To_Upload", df_upload)
        style_sheet(writer, "To_Remove", df_remove)
        style_sheet(writer, "to_remove_others", df_remove_others)

        generate_corporate_report(writer, df_master)

    print("[✔] Excel workbook generated:", OUTPUT_FILE)

if __name__ == "__main__":
    main()
