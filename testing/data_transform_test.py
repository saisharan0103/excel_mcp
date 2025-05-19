# data_transform_test.py

import sys
import os
sys.path.insert(0, os.path.abspath("./src"))

from excel_mcp.data import read_excel_range, write_data

def transform_data(records):
    """Add 'Seniority' column based on Age"""
    transformed = []

    for row in records:
        age = row.get("Age", 0)
        seniority = "Junior" if age < 35 else "Senior"
        new_row = list(row.values()) + [seniority]
        transformed.append(new_row)

    return transformed

# Step 1: Read data from Sheet1
raw_data = read_excel_range(
    filepath="sample_data.xlsx",
    sheet_name="Sheet1",
    start_cell="A1"
)

# Convert rows (list of dicts) to structured headers + transformed rows
headers = list(raw_data[0].keys()) + ["Seniority"]
records = transform_data(raw_data)

# Step 2: Write to new sheet
write_result = write_data(
    filepath="sample_data.xlsx",
    sheet_name="Modified_Data",   # ✅ NEW sheet
    data=[headers] + records,     # Include headers + rows
    start_cell="A1"
)

print("✅ Data transformation complete:")
print(write_result)
