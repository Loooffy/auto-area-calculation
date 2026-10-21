import openpyxl
from pathlib import Path
from openpyxl.utils import get_column_letter

for p in [Path("template/試算面積-混公.xlsx"), Path("試算面積-混公.xlsx")]:
    if p.exists():
        path = p
        break
else:
    raise SystemExit("no source")

wb = openpyxl.load_workbook(path, data_only=False)
ws = wb.active

print("=== HEADER 1-4 ===")
for r in range(1, 5):
    for c in range(1, 24):
        cell = ws.cell(r, c)
        if cell.value is not None:
            print(f"({r},{c}): {cell.value!r} f={cell.data_type == 'f'}")

print("=== TOTAL 54 ===")
for c in range(1, 24):
    cell = ws.cell(54, c)
    if cell.value is not None:
        print(f"(54,{c}): {cell.value!r} f={cell.data_type == 'f'}")

print("=== PARKING 55-84 ===")
for r in range(55, 85):
    for c in range(1, 24):
        cell = ws.cell(r, c)
        if cell.value is not None:
            print(f"({r},{c}): {cell.value!r} f={cell.data_type == 'f'}")

print("=== MERGED ===")
for m in sorted(ws.merged_cells.ranges, key=lambda x: (x.min_row, x.min_col)):
    print(m.coord)

print("=== COL WIDTHS ===")
for c in range(1, 24):
    w = ws.column_dimensions[get_column_letter(c)].width
    print(f"{get_column_letter(c)}: {w}")
