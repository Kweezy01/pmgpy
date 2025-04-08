import pandas as pd
import json

# === Config ===
stock_number_to_find = input("Enter stock number to load: ").strip().upper()

# Load vehicle data from CSV
try:
    df = pd.read_csv("src/pmg_dms_data.csv")
except FileNotFoundError:
    print("❌ src/pmg_dms_data.csv not found. Please make sure the file exists.")
    exit()

# Find the vehicle with the matching stock number
row = df[df["Stock Number"].str.upper() == stock_number_to_find]

if row.empty:
    print(f"❌ No vehicle found with stock number '{stock_number_to_find}'")
    exit()

row = row.iloc[0]  # Get the first match

# Extract year from registration date
reg_date = str(row.get("Registration Date", ""))
year_only = reg_date[-4:] if len(reg_date) >= 4 else ""

# Get and clean mileage
mileage_val = row.get("Odometer", "")
if pd.notnull(mileage_val):
    mileage_val = str(int(float(mileage_val)))  # remove .0 and cast to int then str
else:
    mileage_val = ""

# Get and clean vehicle code
vehicle_code = row.get("Vehicle Code")
vehicle_code_str = str(int(vehicle_code)) if pd.notnull(vehicle_code) else ""

# Extract transmission keyword
transmission_val = str(row.get("Transmission", ""))
transmission_keyword = transmission_val.split()[0] if transmission_val else ""

# Extract drive type from Transmission or Drive
drive_val = str(row.get("Drive", ""))
drive_type = "4x4" if "4x4" in transmission_val or "4x4" in drive_val else "4x2"

# Get engine size
engine_size = str(row.get("Engine Size", "") or "")

# Shared base data
base_data = {
    "StockNum": str(row.get("Stock Number", "") or ""),
    "VIN": str(row.get("VIN", "") or ""),
    "VehicleCode": vehicle_code_str,
    "Year": year_only
}

# Cars JSON
cars_data = base_data.copy()
cars_data.update({
    "Mileage": mileage_val,
    "Color": str(row.get("Colour", "") or "")
})

# AutoTrader JSON
autotrader_data = base_data.copy()
autotrader_data.update({
    "Mileage": mileage_val,
    "Color": str(row.get("Colour", "") or "")
})

# PMG Web JSON
pmg_data = {
    "StockNum": base_data["StockNum"],
    "Make": str(row.get("Make", "") or ""),
        "Variant": str(row.get("Specification", "") or ""),
    "Year": base_data["Year"],
    "Mileage": mileage_val,
    "Color": str(row.get("Colour", "") or ""),
    "FuelType": str(row.get("Fuel Type", "") or ""),
    "Transmission": transmission_keyword,
    "ServiceHistory": "Full History",
    "EngineSize": engine_size,
    "DriveType": drive_type,
    "BodyStyle": str(row.get("Body Style", "") or ""),
    "Interior": str(row.get("Interior", "") or ""),
    "Dealership": str(row.get("Branch", "") or "")
}

# Output JSONs
print("\n📋 Cars.co.za JSON:")
print(json.dumps(cars_data, indent=2))

print("\n📋 AutoTrader JSON:")
print(json.dumps(autotrader_data, indent=2))

print("\n📋 PMG Web JSON:")
print(json.dumps(pmg_data, indent=2))
