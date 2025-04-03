# website_loader.py
import pandas as pd
from pathlib import Path
import warnings

def load_pmg_web_data(web_path: Path) -> pd.DataFrame:
    """Load the PMG web CSV export."""
    try:
        return pd.read_csv(web_path)
    except Exception as e:
        print(f"Error loading PMG web data: {e}")
        return pd.DataFrame()

def load_dealership_data(dealership_dir: Path) -> dict:
    """Load AutoTrader CSV and Cars Excel files for a given dealership folder."""
    data = {}

    autotrader_path = dealership_dir / "autotrader.csv"
    cars_path = dealership_dir / "cars.xlsx"

    try:
        data['autotrader'] = pd.read_csv(autotrader_path, skiprows=1, encoding='utf-8')  # skiprows for sep=,
    except Exception as e:
        print(f"Error loading {autotrader_path}: {e}")
        data['autotrader'] = pd.DataFrame()

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            data['cars'] = pd.read_excel(cars_path, engine='openpyxl')
    except Exception as e:
        print(f"Error loading {cars_path}: {e}")
        data['cars'] = pd.DataFrame()

    return data
