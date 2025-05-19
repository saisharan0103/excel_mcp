# chart_test.py

import sys
import os
sys.path.insert(0, os.path.abspath("./src"))

from excel_mcp.data import write_data
from excel_mcp.chart import create_chart_in_sheet


write_result = write_data(
    filepath="sample_data.xlsx",
    sheet_name="Sheet2",
    
)

# ✅ 2. Create a chart in Sheet3 from that data
result = create_chart_in_sheet(
    filepath="sample_data.xlsx",
    sheet_name="Sheet2",            # Data is in Sheet2 now
    data_range="A1:D6",             # Covers headers + rows
    chart_type="bar",
    target_cell="F5",
    title="Salary by Department",
    x_axis="Department",
    y_axis="Salary",
    style={
        "show_legend": True,
        "legend_position": "r",
        "show_data_labels": True,
        "grid_lines": True
    }
)

print("✅ Data written and chart generated successfully:")
print(result)
