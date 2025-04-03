# dms_loader.py
import pandas as pd
from pathlib import Path

def load_dms_data(dms_path: Path) -> pd.DataFrame:
    """Load the main DMS CSV export."""
    try:
        df = pd.read_csv(dms_path)
        return df
    except Exception as e:
        print(f"Error loading DMS data: {e}")
        return pd.DataFrame()
