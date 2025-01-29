#!/usr/bin/env python3
"""
main.py

Generates 'stockgpt.xlsx' with:
 - One sheet per dealership prefix (UF, UG, UA, UE, US)
 - 'to_be_removed' for website-only vehicles
 - 'to_dos' with a 'Completed?' column (choosing "Yes" turns row green).

Suppresses OpenPyXL warnings and improves "Notes" column in `to_dos`
to list all missing sites.

Usage:
  python -m stockgpt.main
"""

import os
import pandas as pd
import warnings

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

from openpyxl.worksheet.table import Table, TableStyleInfo
from utilities import clean_dataframe
from .data_readers import (
    read_dms_dict,
    read_autotrader_data,
    read_cars_data,
    read_pmg_web_data,
)
from .transformations import (
    build_master_df,
    reorder_final_columns,
    generate_todos
)
from .formatting import style_sheet, create_excel_table, auto_size_columns

# Suppress OpenPyXL warnings about missing default styles
warnings.simplefilter("ignore", UserWarning)

# Prefixes mapped to dealer names
DEALER_PREFIXES = {
    "Ford_Nelspruit":   "UF",
    "Ford_Mazda":       "UG",
    "Produkta_Nissan":  "UA",
    "Suzuki_Nelspruit": "UE",  # forced is_on_cars="Yes"
    "Ford_Malalane":    "US",
}

def style_to_dos_sheet(writer, sheet_name: str, df_todos: pd.DataFrame):
    """
    Creates a table, auto-sizes columns, adds data validation for 'Completed?' column,
    and color-codes row green if 'Completed?' == 'Yes'.
    """
    ws = writer.sheets[sheet_name]

    # 1) Turn data region into a table
    create_excel_table(ws, df_todos, table_name="ToDosTable")
    # 2) Auto-size columns
    auto_size_columns(ws, df_todos)

    # 3) Data validation => "Yes"/"No" for 'Completed?' column
    if "Completed?" in df_todos.columns:
        rows_count = len(df_todos)
        cols_count = len(df_todos.columns)
        completed_ix = df_todos.columns.get_loc("Completed?") + 1  # 1-based

        dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
        ws.add_data_validation(dv)

        # Apply to data rows => row 2.. row_count+1
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(completed_ix)
        for row in range(2, rows_count + 2):
            cell_coord = f"{col_letter}{row}"
            dv.add(cell_coord)

        # 4) If 'Completed?' == 'Yes', entire row => green
        last_col_letter = get_column_letter(cols_count)
        rng = f"A2:{last_col_letter}{rows_count+1}"
        formula_str = f'${col_letter}2="Yes"'
        fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        rule_green = FormulaRule(formula=[formula_str], fill=fill_green, stopIfTrue=False)
        ws.conditional_formatting.add(rng, rule_green)


def main():
    src_folder = "src"
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    print("[INFO] Reading DMS data...")
    dms_map = read_dms_dict(os.path.join(src_folder, "pmg_dms_data.csv"))
    print(f"[INFO] DMS cars loaded: {len(dms_map)}")

    print("[INFO] Reading website data...")
    at_set, at_prices = read_autotrader_data(src_folder)
    cars_set, cars_prices = read_cars_data(src_folder)
    pmg_set, pmg_prices = read_pmg_web_data(src_folder)

    print("[INFO] Building master dataset...")
    df_master = build_master_df(
        dms_map,
        at_set, at_prices,
        cars_set, cars_prices,
        pmg_set, pmg_prices
    )
    df_master = reorder_final_columns(df_master)

    print(f"[INFO] Total cars in master dataset: {df_master.shape}")

    recognized_prefixes = tuple(DEALER_PREFIXES.values())

    df_master["Stock Number"] = df_master["Stock Number"].astype(str)
    df_in_dms = df_master[
        (df_master["in_dms"]==True) &
        df_master["Stock Number"].str.startswith(recognized_prefixes, na=False)
    ].copy()

    print(f"[INFO] Cars in DMS with recognized prefixes: {df_in_dms.shape}")

    df_removed = df_master[
        (df_master["in_dms"]==False) &
        df_master["Stock Number"].str.startswith(recognized_prefixes, na=False)
    ].copy()
    remove_cols = [
        "Stock Number",
        "is_on_cars","cars_price",
        "is_on_autotrader","autotrader_price",
        "is_on_pmgWeb"
    ]
    df_removed = df_removed.reindex(columns=remove_cols)
    df_removed = clean_dataframe(df_removed)

    df_todos = generate_todos(df_in_dms, df_removed)
    if not df_todos.empty and "Completed?" not in df_todos.columns:
        df_todos["Completed?"] = ""

    out_file = os.path.join(output_folder, "stockgpt.xlsx")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_count = 0

        for dealer_name, prefix_val in DEALER_PREFIXES.items():
            df_sub = df_in_dms[df_in_dms["Stock Number"].str.startswith(prefix_val)].copy()
            if not df_sub.empty:
                df_sub.to_excel(writer, sheet_name=dealer_name, index=False)
                sheet_count += 1

        if not df_removed.empty:
            df_removed.to_excel(writer, sheet_name="to_be_removed", index=False)
            sheet_count += 1

        if not df_todos.empty:
            df_todos.to_excel(writer, sheet_name="to_dos", index=False)
            sheet_count += 1

        for sheet_name in writer.sheets.keys():
            if sheet_name == "to_be_removed":
                style_sheet(writer, sheet_name, df_removed)
            elif sheet_name == "to_dos":
                style_to_dos_sheet(writer, sheet_name, df_todos)
            else:
                df_sub = df_in_dms[df_in_dms["Stock Number"].str.startswith(DEALER_PREFIXES[sheet_name])]
                style_sheet(writer, sheet_name, df_sub)

    print(f"[stockgpt] Wrote {sheet_count} sheets => {out_file}")
    print("Dealership sheets, 'to_be_removed' & 'to_dos' included.")


if __name__ == "__main__":
    main()
