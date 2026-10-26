import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from typing import List, Dict, Any, Optional

DEFAULT_TEMPLATE_PATH = r"D:\projects\boq-services\Blank BOQ_PLC HOUSE .xlsx"

def load_template_workbook(template_source: Optional[Any] = None) -> openpyxl.Workbook:
    """Loads openpyxl Workbook from path, buffer, or defaults to master file."""
    if template_source is None:
        return openpyxl.load_workbook(DEFAULT_TEMPLATE_PATH)
    elif isinstance(template_source, str) and os.path.exists(template_source):
        return openpyxl.load_workbook(template_source)
    elif hasattr(template_source, "read"):
        data = template_source.read()
        if hasattr(template_source, "seek"):
            template_source.seek(0)
        return openpyxl.load_workbook(io.BytesIO(data))
    elif isinstance(template_source, bytes):
        return openpyxl.load_workbook(io.BytesIO(template_source))
    else:
        return openpyxl.load_workbook(DEFAULT_TEMPLATE_PATH)

def get_template_items(wb: openpyxl.Workbook) -> List[Dict[str, Any]]:
    """Reads existing line items from the 'BOQ' sheet in template."""
    ws = wb["BOQ"]
    items = []
    current_category = "งานทั่วไป"
    
    for r in range(6, ws.max_row + 1):
        item_no = ws.cell(r, 1).value
        code = ws.cell(r, 2).value
        desc = ws.cell(r, 7).value
        qty = ws.cell(r, 17).value
        unit = ws.cell(r, 19).value
        mat_rate = ws.cell(r, 20).value
        lab_rate = ws.cell(r, 22).value
        
        # Check if it's a section header (e.g. ชั้น 1, ชั้น 2)
        if item_no is not None and code is not None and desc is None:
            current_category = str(code)
            continue
            
        # Check if it's a subtotal row
        if code and "รวม" in str(code):
            continue
            
        # If it's a valid item row
        if (code or desc) and (qty is not None or unit is not None or desc is not None):
            items.append({
                "row_idx": r,
                "category": current_category,
                "code": str(code or "").strip(),
                "description": str(desc or code or "").strip(),
                "qty": float(qty) if isinstance(qty, (int, float)) else 1.0,
                "unit": str(unit or "SET").strip(),
                "material_rate": float(mat_rate) if isinstance(mat_rate, (int, float)) else 0.0,
                "labour_rate": float(lab_rate) if isinstance(lab_rate, (int, float)) else 0.0,
                "remark": ws.cell(r, 27).value or ""
            })
            
    return items

def populate_boq_excel(
    template_source: Optional[Any],
    items_data: List[Dict[str, Any]],
    overhead_pct: float = 10.0,
    project_name: str = "PLC House (Reignwood Park)",
    customer_name: str = ""
) -> io.BytesIO:
    """
    Fills items_data into the Excel template with exact formatting and formulas.
    Returns BytesIO object containing the final .xlsx file.
    """
    wb = load_template_workbook(template_source)
    ws = wb["BOQ"]
    
    # Map input items by code or description
    item_map = {}
    for item in items_data:
        c = str(item.get("code") or "").strip().upper()
        if c:
            item_map[c] = item
            
    # Iterate through BOQ rows and update matching items
    for r in range(6, ws.max_row + 1):
        code_cell = ws.cell(r, 2).value
        c_str = str(code_cell or "").strip().upper()
        
        # Check if this row is an item row
        if c_str and "รวม" not in c_str and ws.cell(r, 1).value is None:
            # We have an item row
            matched = item_map.get(c_str)
            
            # Extract values
            qty = matched.get("qty") if matched else ws.cell(r, 17).value
            unit = matched.get("unit") if matched else ws.cell(r, 19).value
            mat_rate = matched.get("material_rate") if matched else ws.cell(r, 20).value
            lab_rate = matched.get("labour_rate") if matched else ws.cell(r, 22).value
            
            # Format numbers safely
            try:
                qty_val = float(qty) if qty is not None else 1.0
            except:
                qty_val = 1.0
                
            try:
                mat_val = float(mat_rate) if mat_rate is not None else 0.0
            except:
                mat_val = 0.0
                
            try:
                lab_val = float(lab_rate) if lab_rate is not None else 0.0
            except:
                lab_val = 0.0
                
            # Assign values
            ws.cell(r, 17).value = qty_val
            ws.cell(r, 19).value = unit or "SET"
            ws.cell(r, 20).value = mat_val
            ws.cell(r, 22).value = lab_val
            
            # Col 24 (X) formula: =(T{r} + V{r}) * Q{r}
            ws.cell(r, 24).value = f"=(T{r}+V{r})*Q{r}"
            
    # Subtotal formulas in 'BOQ'
    ws.cell(21, 24).value = "=SUM(X7:X20)"        # รวมชั้น 1
    ws.cell(41, 24).value = "=SUM(X24:X40)"       # รวมชั้น 2
    ws.cell(47, 24).value = "=SUM(X43:X46)"       # รวมงานระบบ
    ws.cell(52, 24).value = "=SUM(X49:X51)"       # รวม Preliminary
    ws.cell(54, 24).value = "=X21+X41+X47"        # รวม 1+2+3
    ws.cell(55, 24).value = f"=X54*{overhead_pct / 100.0}" # ค่าดำเนินการ
    ws.cell(56, 24).value = "=X54+X55+X52"        # รวมทั้งหมด
    
    # Update 'Summary' sheet
    if "Summary" in wb.sheetnames:
        s_ws = wb["Summary"]
        if project_name:
            s_ws.cell(3, 8).value = project_name
        if customer_name:
            s_ws.cell(4, 8).value = customer_name
            
        # Link Summary sheet rows directly to BOQ sheet cells
        s_ws.cell(8, 26).value = "=BOQ!X21"   # ชั้น 1 Amount
        s_ws.cell(9, 26).value = "=BOQ!X41"   # ชั้น 2 Amount
        s_ws.cell(10, 26).value = "=BOQ!X47"  # งานระบบ
        s_ws.cell(11, 26).value = "=BOQ!X52"  # งานเตรียมการ (Preliminary)
        s_ws.cell(12, 26).value = "=SUM(Z8:Z11)" # Total
        s_ws.cell(13, 26).value = "=BOQ!X55"  # ค่าดำเนินการ
        s_ws.cell(14, 26).value = "=Z12+Z13"  # Sum Total

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
