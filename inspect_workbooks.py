
from pathlib import Path

import pandas as pd


DATA_DIR = Path("data/raw")

workbooks = [
    DATA_DIR / "ABB_Price_List_Feb2026.xlsx",
    DATA_DIR / "Schneider LV Price list Apr 2026.xlsx",
]

for workbook in workbooks:
    print(f"\n{'=' * 70}")
    print(f"Workbook: {workbook.name}")
    print("=" * 70)

    if not workbook.exists():
        print("File not found. Check the filename and data/raw folder.")
        continue

    excel_file = pd.ExcelFile(workbook)

    print("Sheets:", excel_file.sheet_names)

    for sheet_name in excel_file.sheet_names:
        preview = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
            header=None,
            nrows=6,
        )

        print(f"\nSheet: {sheet_name}")
        print(f"Preview shape: {preview.shape}")
        print(preview.to_string(index=False, header=False))
