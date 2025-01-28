import os

csv_path = os.path.join("/", folder_name, "autotrader.csv")
if not os.path.isfile(csv_path):
    print("UP")
print("Down")