# main.py
from pathlib import Path
from .data_loader.dms_loader import load_dms_data
from .data_loader.website_loader import load_pmg_web_data, load_dealership_data
from .exporter.excel_report import write_master_excel

def main():
    base_path = Path(__file__).parent / "src"

    # Load core DMS and PMG web data
    dms_data = load_dms_data(base_path / "pmg_dms_data.csv")
    pmg_web_data = load_pmg_web_data(base_path / "pmg_web_data.csv")

    print("DMS Data Loaded:", dms_data.shape)
    print("PMG Web Data Loaded:", pmg_web_data.shape)

    # Dealerships to process
    dealership_dirs = [
        "fordMalalane",
        "fordNelspruit",
        "mazdaNelspruit",
        "produktaNissan",
        "suzukiNissan",
    ]

    dealership_data = {}
    for name in dealership_dirs:
        folder_path = base_path / name
        dealership_data[name] = load_dealership_data(folder_path)
        print(f"Loaded {name}:",
              "AutoTrader:", dealership_data[name]['autotrader'].shape,
              "Cars:", dealership_data[name]['cars'].shape)

    # Output Excel workbook
    output_path = base_path.parent / "master_vehicle_report.xlsx"
    write_master_excel(dms_data, pmg_web_data, dealership_data, output_path)
    print(f"Master Excel workbook saved to {output_path}")

if __name__ == "__main__":
    main()
