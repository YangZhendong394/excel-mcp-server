#!/usr/bin/env python3
"""Build a formula workbook whose formula cells intentionally have no cached results."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output", required=True, type=Path)
parser.add_argument("--vendor-dir", required=True, type=Path)
args = parser.parse_args()

sys.path.insert(0, str(args.vendor_dir.resolve(strict=True)))
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

output = args.output.resolve()
output.parent.mkdir(parents=True, exist_ok=True)

workbook = Workbook()
sheet = workbook.active
sheet.title = "计算测试"
sheet.append(["项目", "公式"])
sheet.append(["10 + 20", "=SUM(10,20)"])
sheet.append(["5 × 6", "=5*6"])
for cell in sheet[1]:
    cell.font = Font(bold=True, color="FFFFFFFF")
    cell.fill = PatternFill(fill_type="solid", fgColor="FF1F4E78")
workbook.save(output)
workbook.close()
print(f"FORMULA_FIXTURE={output}")
print(f"FORMULA_FIXTURE_BYTES={output.stat().st_size}")
