import csv
from openpyxl import Workbook
from pathlib import Path

A3 = Path(r"A:\TTED731\03_Assignment3_Instructional_Design")
csv_file = A3 / "Reference Verification" / "crossref_verification_results.csv"
xlsx_file = A3 / "Reference Verification" / "Assignment3_Verified_References.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Verified References"

with csv_file.open(newline="", encoding="utf-8") as f:
    reader = csv.reader(f)
    for row in reader:
        ws.append(row)

wb.save(xlsx_file)

print(f"Created: {xlsx_file}")
