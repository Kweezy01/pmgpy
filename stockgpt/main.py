#!/usr/bin/env python3
"""
main.py

Fixes:  
 - Prevents KeyError for 'corporate_report'.  
 - Ensures all sheets retain table formatting.  
 - Business-ready charts for director meetings.  

Usage:
  python -m stockgpt.main
"""

import os
import pandas as pd
import warnings
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.worksheet.table import Table, TableStyleInfo

from utilities import clean_dataframe
from .data_readers import read_dms_dict, read_autotrader_data, read_cars_data, read_pmg_web_data
from .transformations import build_master_df, reorder_final_columns, generate_todos
from .formatting import style_sheet, create_excel_table, auto_size_columns

warnings.simplefilter("ignore", UserWarning)  # Suppress OpenPyXL warnings

DEALER_PREFIXES = {
    "Ford_Nelspruit":   "UF",
    "Ford_Mazda":       "UG",
    "Produkta_Nissan":  "UA",
    "Suzuki_Nelspruit": "UE",
    "Ford_Malalane":    "US",
}

def generate_corporate_report(writer, df_master):
    """
    Creates a 'corporate_report' sheet with key statistics and professional graphs.
    """
    ws = writer.book.create_sheet("corporate_report")

    # --- 📊 Key Metrics ---
    total_dms = df_master[df_master["in_dms"]].shape[0]
    total_cars = df_master[df_master["is_on_cars"] == "Yes"].shape[0]
    total_autotrader = df_master[df_master["is_on_autotrader"] == "Yes"].shape[0]
    total_pmgweb = df_master[df_master["is_on_pmgWeb"] == "Yes"].shape[0]
    total_to_remove = df_master[df_master["in_dms"] == False].shape[0]

    ws.append(["Corporate Vehicle Report"])
    ws.append([""])
    ws.append(["Total Vehicles in DMS:", total_dms])
    ws.append(["Total Vehicles on Cars.co.za:", total_cars])
    ws.append(["Total Vehicles on AutoTrader:", total_autotrader])
    ws.append(["Total Vehicles on PMG Web:", total_pmgweb])
    ws.append(["Total Vehicles to Remove:", total_to_remove])
    ws.append([""])

    # --- 📌 Breakdown by Dealership ---
    ws.append(["Dealership", "Total Vehicles"])
    row_start = ws.max_row + 1

    for dealer, prefix in DEALER_PREFIXES.items():
        count = df_master[df_master["Stock Number"].str.startswith(prefix)].shape[0]
        ws.append([dealer, count])

    row_end = ws.max_row

    # --- 📈 Bar Chart (Dealership Breakdown) ---
    chart = BarChart()
    chart.title = "Vehicle Count Per Dealership"
    chart.x_axis.title = "Dealerships"
    chart.y_axis.title = "Total Vehicles"
    chart.width = 12
    chart.height = 6

    data = Reference(ws, min_col=2, min_row=row_start, max_row=row_end)
    categories = Reference(ws, min_col=1, min_row=row_start + 1, max_row=row_end)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "E5")

    # --- 📊 Pie Chart (Stock Distribution Across Websites) ---
    pie = PieChart()
    pie.title = "Stock Distribution on Websites"
    pie_data = Reference(ws, min_col=2, min_row=3, max_row=6)  # Total stock numbers
    pie_categories = Reference(ws, min_col=1, min_row=3, max_row=6)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_categories)
    ws.add_chart(pie, "E15")

    auto_size_columns(ws, df_master)  # ✅ Fixed: Passing df_master correctly

def main():
    src_folder, output_folder = "src", "output"
    os.makedirs(output_folder, exist_ok=True)

    print("[INFO] Reading DMS data...")
    dms_map = read_dms_dict(os.path.join(src_folder, "pmg_dms_data.csv"))
    print(f"[INFO] DMS cars loaded: {len(dms_map)}")

    print("[INFO] Reading website data...")
    at_set, at_prices = read_autotrader_data(src_folder)
    cars_set, cars_prices = read_cars_data(src_folder)
    pmg_set, pmg_prices = read_pmg_web_data(src_folder)

    print("[INFO] Building master dataset...")
    df_master = reorder_final_columns(build_master_df(
        dms_map, at_set, at_prices, cars_set, cars_prices, pmg_set, pmg_prices
    ))

    out_file = os.path.join(output_folder, "stockgpt.xlsx")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_count = 0

        # Write dealership sheets with table formatting
        for dealer, prefix in DEALER_PREFIXES.items():
            df_sub = df_master[df_master["Stock Number"].str.startswith(prefix)].copy()
            if not df_sub.empty:
                df_sub.to_excel(writer, sheet_name=dealer, index=False)
                sheet_count += 1

        # Write to_be_removed (Minimal: Only website info + "Done?" Column)
        df_removed = df_master[df_master["in_dms"] == False][
            ["Stock Number", "is_on_cars", "is_on_autotrader", "is_on_pmgWeb"]
        ].copy()
        df_removed["Done?"] = ""

        if not df_removed.empty:
            df_removed.to_excel(writer, sheet_name="to_be_removed", index=False)
            sheet_count += 1

        # Write to_dos (with "Done?" Column)
        df_todos = generate_todos(df_master, df_removed)
        if not df_todos.empty:
            df_todos["Done?"] = ""
            df_todos.to_excel(writer, sheet_name="to_dos", index=False)
            sheet_count += 1

        # Add Corporate Report with Graphs
        generate_corporate_report(writer, df_master)

        # Apply table formatting to all sheets (SKIP "corporate_report")
        for sheet_name in writer.sheets.keys():
            if sheet_name == "corporate_report":
                continue  # ✅ FIX: Skip formatting for corporate_report

            sheet = writer.sheets[sheet_name]
            if sheet_name in ["to_be_removed", "to_dos"]:
                style_sheet(writer, sheet_name, df_removed if sheet_name == "to_be_removed" else df_todos)
            else:
                df_sub = df_master[df_master["Stock Number"].str.startswith(DEALER_PREFIXES[sheet_name])]
                style_sheet(writer, sheet_name, df_sub)

    print(f"[stockgpt] Wrote {sheet_count + 1} sheets => {out_file}")
    print("Includes 'corporate_report' with charts & retains table formatting.")

if __name__ == "__main__":
    main()
