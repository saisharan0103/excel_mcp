# run_test.py

import sys
import os
sys.path.insert(0, os.path.abspath("./src"))  # So Python can see 'excel_mcp'

from excel_mcp.data import read_excel_range, write_data

# === READ EXAMPLE ===
read_result = read_excel_range(
    filepath="sample_data.xlsx",     # Make sure this file exists
    sheet_name="Sheet1",             # Adjust to your actual sheet name
    start_cell="A1"
)

print("📥 Read Result:")
for row in read_result:
    print(row)

# === WRITE EXAMPLE ===
sample_write_data = [
    ["Name", "Score"],
    ["Sai", 95],
    ["Rhea", 88],
    ["AI", 100]
]

write_result = write_data(
    filepath="sample_data.xlsx",
    sheet_name="Sheet2",            
    data=sample_write_data,
    start_cell="A7"                 
)

print("📤 Write Result:")
print(write_result)
