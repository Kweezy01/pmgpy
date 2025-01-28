# formatting.py

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from utilities import col_index_to_excel_col_name

def auto_size_columns(sheet, df: pd.DataFrame):
    """
    Approx 'auto-size' each column by checking header + sample data.
    """
    for i, col_name in enumerate(df.columns, start=1):
        hdr_len = len(col_name)
        sample_vals = df[col_name].astype(str).head(50).tolist()
        avg_len = sum(len(x) for x in sample_vals)/max(len(sample_vals),1)
        best = int(max(hdr_len, avg_len)) + 2
        sheet.column_dimensions[get_column_letter(i)].width = best


def create_excel_table(sheet, df: pd.DataFrame, table_name="DataTable"):
    """
    Turn the data region => an Excel Table with style Medium9
    """
    rows = df.shape[0]
    cols = df.shape[1]
    if rows<1 or cols<1:
        return

    last_col_letter = col_index_to_excel_col_name(cols-1)
    ref = f"A1:{last_col_letter}{rows+1}"  # +1 for header
    tab = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    tab.tableStyleInfo = style
    sheet.add_table(tab)


def apply_conditional_formatting(sheet, df: pd.DataFrame):
    """
    Green => all site columns == "Yes"
    Red => PhotoCount>1 & any site col == "No"
    """
    all_cols = df.columns.tolist()
    # green
    try:
        c_cars = all_cols.index("is_on_cars")
        c_auto = all_cols.index("is_on_autotrader")
        c_pmg  = all_cols.index("is_on_pmgWeb")
    except ValueError:
        pass
    else:
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        col_cars = col_index_to_excel_col_name(c_cars)
        col_auto = col_index_to_excel_col_name(c_auto)
        col_pmg  = col_index_to_excel_col_name(c_pmg)
        max_row = len(df)+1
        max_col = len(all_cols)
        last_col = col_index_to_excel_col_name(max_col-1)
        rng = f"A2:{last_col}{max_row}"
        formula_g = f'AND(${col_cars}2="Yes", ${col_auto}2="Yes", ${col_pmg}2="Yes")'
        rule_g = FormulaRule(formula=[formula_g], fill=green_fill, stopIfTrue=False)
        sheet.conditional_formatting.add(rng, rule_g)

    # red
    if "Photo Count" in all_cols and "is_on_cars" in all_cols and \
       "is_on_autotrader" in all_cols and "is_on_pmgWeb" in all_cols:
        photo_ix = all_cols.index("Photo Count")
        col_photo = col_index_to_excel_col_name(photo_ix)
        col_cars  = col_index_to_excel_col_name(all_cols.index("is_on_cars"))
        col_auto  = col_index_to_excel_col_name(all_cols.index("is_on_autotrader"))
        col_pmg   = col_index_to_excel_col_name(all_cols.index("is_on_pmgWeb"))
        max_row = len(df)+1
        last_col = col_index_to_excel_col_name(len(all_cols)-1)
        rng = f"A2:{last_col}{max_row}"

        formula_r = f'AND(${col_photo}2>1, OR(${col_cars}2="No", ${col_auto}2="No", ${col_pmg}2="No"))'
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        rule_r = FormulaRule(formula=[formula_r], fill=red_fill, stopIfTrue=False)
        sheet.conditional_formatting.add(rng, rule_r)


def style_sheet(writer, sheet_name: str, df: pd.DataFrame):
    """
    - Insert Excel Table
    - Auto-size columns
    - Apply color-coded rules
    """
    sheet = writer.sheets[sheet_name]
    create_excel_table(sheet, df, table_name=f"{sheet_name}Table")
    auto_size_columns(sheet, df)
    apply_conditional_formatting(sheet, df)
