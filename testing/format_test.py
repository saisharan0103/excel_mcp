# format_existing_data_test.py

import sys
import os
sys.path.insert(0, os.path.abspath("./src"))

from excel_mcp.formatting import format_range

filepath = "sample_data.xlsx"
sheet = "Sheet1"

# 1. Format the header row
format_range(
    filepath=filepath,
    sheet_name=sheet,
    start_cell="A1",
    end_cell="C1",
    bold=True,
    font_color="FFFFFF",
    bg_color="4F81BD",   # Blue-gray
    alignment="center",
    wrap_text=True,
    border_style="thin",
    border_color="000000"
)

# 2. Format data rows
format_range(
    filepath=filepath,
    sheet_name=sheet,
    start_cell="A2",
    end_cell="C5",
    alignment="center",
    border_style="thin",
    border_color="000000"
)

# 3. Highlight Score < 70
format_range(
    filepath=filepath,
    sheet_name=sheet,
    start_cell="B2",
    end_cell="B5",
    conditional_format={
        "type": "cell_is",
        "params": {
            "operator": "lessThan",
            "formula": ["70"],
            "stopIfTrue": False,
            "fill": {"fgColor": "FFC7CE"}
        }
    }
)

# 4. Highlight Status == "Fail"
format_range(
    filepath=filepath,
    sheet_name=sheet,
    start_cell="C2",
    end_cell="C5",
    conditional_format={
        "type": "cell_is",
        "params": {
            "operator": "equal",
            "formula": ['"Fail"'],
            "stopIfTrue": False,
            "fill": {"fgColor": "FFC7CE"}
        }
    }
)

print("✅ Existing data formatted successfully!")
