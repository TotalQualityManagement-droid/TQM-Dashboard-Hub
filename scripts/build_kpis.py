"""Generate website KPI data from the maintained Excel workbook."""
import json
from pathlib import Path
from openpyxl import load_workbook

def read_kpis():
    workbook = load_workbook("TQM_Analytics_Hub_KPIs.xlsx", data_only=True)
    sheet = workbook["KPI Data"]
    headers = [str(cell.value or "").strip() for cell in sheet[1]]
    required = ["Dashboard Name", "KPI", "Value"]
    if any(key not in headers for key in required):
        raise ValueError("KPI Data must contain Dashboard Name, KPI, and Value columns")
    rows = []
    for cells in sheet.iter_rows(min_row=2):
        row = {key: cells[headers.index(key)].value for key in required}
        if not any(value is not None for value in row.values()):
            continue
        if not row["Dashboard Name"] or not row["KPI"]:
            raise ValueError(f"Missing dashboard or KPI at row {cells[0].row}")
        value_cell = cells[headers.index("Value")]
        value = row["Value"]
        if value_cell.data_type == "e":
            raise ValueError(f"Excel error at {value_cell.coordinate}")
        if value is None:
            row["Value"] = "Not reported"
        elif isinstance(value, (int, float)) and "%" in value_cell.number_format:
            row["Value"] = f"{value * 100:.2f}".rstrip("0").rstrip(".") + "%"
        else:
            row["Value"] = str(value)
        rows.append(row)
    if not rows:
        raise ValueError("Workbook contains no KPI rows")
    return rows

if __name__ == "__main__":
    rows = read_kpis()
    Path("tqm-kpis.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generated {len(rows)} KPI rows from Excel")
