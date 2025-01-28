# main.py

import os
import pandas as pd

from .data_readers import (
    read_dms_dict,
    read_autotrader_data,
    read_cars_data,
    read_pmg_web_data
)
from .transformations import (
    build_master_df,
    reorder_final_columns,
    generate_todos,
    DEALER_PREFIXES
)
from .formatting import style_sheet

from utilities import clean_dataframe

def main():
    src_folder = "src"
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    # 1) read DMS -> dict
    dms_map = read_dms_dict(os.path.join(src_folder))

    # 2) read website sets/dicts
    at_stk, at_prices   = read_autotrader_data(src_folder)
    cars_stk, cars_prices = read_cars_data(src_folder)
    pmg_stk, pmg_prices   = read_pmg_web_data(src_folder)

    # 3) build master df
    df_master = build_master_df(dms_map,
                                at_stk, at_prices,
                                cars_stk, cars_prices,
                                pmg_stk, pmg_prices)
    df_master = reorder_final_columns(df_master)

    # recognized prefixes
    recognized_prefixes = tuple(DEALER_PREFIXES.values())

    # in_dms => prefix => multiple sheets
    df_in_dms = df_master[
        (df_master["in_dms"]==True) &
        df_master["Stock Number"].str.startswith(recognized_prefixes, na=False)
    ].copy()

    # to_be_removed => website-only => prefix recognized
    df_removed = df_master[
        (df_master["in_dms"]==False) &
        df_master["Stock Number"].str.startswith(recognized_prefixes, na=False)
    ].copy()

    # keep only 6 columns in removed
    rm_cols = [
        "Stock Number",
        "is_on_cars","cars_price",
        "is_on_autotrader","autotrader_price",
        "is_on_pmgWeb"
    ]
    df_removed = df_removed.reindex(columns=rm_cols)
    df_removed = clean_dataframe(df_removed)

    # generate to_dos
    df_todos = generate_todos(df_in_dms, df_removed)

    # 4) write everything
    out_file = os.path.join(output_folder, "stockgpt.xlsx")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        sheet_count = 0

        # a) one sheet per prefix
        for dealer_name, prefix_val in DEALER_PREFIXES.items():
            subset = df_in_dms[df_in_dms["Stock Number"].str.startswith(prefix_val)]
            if subset.empty:
                continue
            subset.to_excel(writer, sheet_name=dealer_name, index=False)
            sheet_count += 1

        # b) to_be_removed
        if not df_removed.empty:
            df_removed.to_excel(writer, sheet_name="to_be_removed", index=False)
            sheet_count += 1

        # c) to_dos
        if not df_todos.empty:
            df_todos.to_excel(writer, sheet_name="to_dos", index=False)
            sheet_count += 1

        # styling each sheet
        for sheet_name in writer.sheets.keys():
            if sheet_name == "to_be_removed":
                style_sheet(writer, sheet_name, df_removed)
            elif sheet_name == "to_dos":
                # no color code but we do table + auto-size
                from .formatting import create_excel_table, auto_size_columns
                ws = writer.sheets[sheet_name]
                create_excel_table(ws, df_todos, table_name="ToDosTable")
                auto_size_columns(ws, df_todos)
            else:
                # must be a dealer prefix
                prefix_val = None
                for nm,val in DEALER_PREFIXES.items():
                    if nm == sheet_name:
                        prefix_val = val
                if prefix_val:
                    sub = df_in_dms[df_in_dms["Stock Number"].str.startswith(prefix_val)]
                    style_sheet(writer, sheet_name, sub)

    print(f"[stockgpt] Wrote {sheet_count} sheets => {out_file}\n"
          f"One sheet per prefix, plus to_be_removed, plus to_dos.")

if __name__=="__main__":
    main()
