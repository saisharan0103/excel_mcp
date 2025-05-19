import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "src")))

from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from excel_mcp.data import write_data
from excel_mcp.formatting import format_range

app = FastAPI()

# ✅ Input model for writing Excel data
class WriteDataRequest(BaseModel):
    filepath: str
    sheet_name: str
    data: List[List[Any]]
    start_cell: Optional[str] = "A1"

@app.post("/write-data")
def write_to_excel(req: WriteDataRequest):
    
    if not os.path.exists(req.filepath):
        wb = Workbook()
        wb.save(req.filepath)
        
    result = write_data(
        filepath=req.filepath,
        sheet_name=req.sheet_name,
        data=req.data,
        start_cell=req.start_cell
    )
    return {"status": "success", "details": result}


# ✅ Input model for formatting Excel ranges
class FormatRequest(BaseModel):
    filepath: str
    sheet_name: str
    start_cell: str
    end_cell: Optional[str] = None
    bold: Optional[bool] = False
    italic: Optional[bool] = False
    underline: Optional[bool] = False
    font_size: Optional[int] = None
    font_color: Optional[str] = None
    bg_color: Optional[str] = None
    border_style: Optional[str] = None
    border_color: Optional[str] = None
    number_format: Optional[str] = None
    alignment: Optional[str] = None
    wrap_text: Optional[bool] = False
    merge_cells: Optional[bool] = False
    protection: Optional[Dict[str, Any]] = None
    conditional_format: Optional[Dict[str, Any]] = None

@app.post("/format-range")
def format_excel_range(req: FormatRequest):
    result = format_range(
        filepath=req.filepath,
        sheet_name=req.sheet_name,
        start_cell=req.start_cell,
        end_cell=req.end_cell,
        bold=req.bold,
        italic=req.italic,
        underline=req.underline,
        font_size=req.font_size,
        font_color=req.font_color,
        bg_color=req.bg_color,
        border_style=req.border_style,
        border_color=req.border_color,
        number_format=req.number_format,
        alignment=req.alignment,
        wrap_text=req.wrap_text,
        merge_cells=req.merge_cells,
        protection=req.protection,
        conditional_format=req.conditional_format
    )
    return {"status": "success", "details": result}
