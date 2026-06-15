"""
Print column indices → values for every non-blank row in each XLS file.
Run once to map exact positions before writing build_dashboard.py.
"""
import xlrd
from pathlib import Path

reports_dir = Path(__file__).parent.parent / "reports"

for period in ["prior_week"]:
    for xlsx in sorted((reports_dir / period).glob("*.xlsx")):
        print(f"\n{'='*60}")
        print(f"FILE: {xlsx.name}")
        print(f"{'='*60}")
        wb = xlrd.open_workbook(str(xlsx))
        ws = wb.sheets()[0]
        print(f"Size: {ws.nrows} rows × {ws.ncols} cols")
        for r in range(ws.nrows):
            row = ws.row_values(r)
            # Only print rows that have at least one non-empty cell
            cells = {c: v for c, v in enumerate(row) if v != '' and v is not None}
            if cells:
                print(f"  row{r:02d}: {cells}")
