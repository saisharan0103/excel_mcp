# pivot_test.py

import sys
import os
sys.path.insert(0, os.path.abspath("./src"))

from excel_mcp.pivot import create_pivot_table

result = create_pivot_table(
    filepath="sample_data.xlsx",
    sheet_name="Sheet1",
    data_range="A1:D6",              # 👈 Must include headers!
    rows=["Department"],             # 👈 Group by Department
    values=["Salary"],               # 👈 Aggregate on Salary
    agg_func="average"               # 👈 Try "sum", "count", "min", "max", etc.
)

print("✅ Pivot Table Created:")
print(result)
