
from pathlib import Path

import pandas as pd


DATA_DIR = Path("data/raw")

workbooks = [
    DATA_DIR / "ABB_Price_List_Feb2026.xlsx",
    DATA_DIR / "Schneider LV Price list Apr 2026.xlsx",
]

# Words that may appear in a product table's column headings.
header_keywords = {
    "manufacturer",
    "category",
    "description",
    "reference",
    "unit_mrp",
    "part code",
    "price",
    "net price",
    "group",
    "rated_a",
    "poles",
}


def normalize_cell(value):
    """Convert a cell to comparable lowercase text."""
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


for workbook in workbooks:
    print(f"\n{'=' * 70}")
    print(f"Workbook: {workbook.name}")
    print("=" * 70)

    if not workbook.exists():
        print(f"File not found: {workbook}")
        continue

    excel_file = pd.ExcelFile(workbook)

    for sheet_name in excel_file.sheet_names:
        # Read the first 30 rows without assuming a header row.
        raw = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
            header=None,
            nrows=30,
        )

        candidates = []

        for row_index, row in raw.iterrows():
            cells = {normalize_cell(value) for value in row}
            matches = cells.intersection(header_keywords)

            # A real heading row usually contains several known labels.
            if len(matches) >= 3:
                candidates.append((row_index, matches))

        if not candidates:
            continue

        print(f"\nSheet: {sheet_name}")

        for header_index, matches in candidates:
            print(f"Possible header row: Excel row {header_index + 1}")
            print("Recognized labels:", ", ".join(sorted(matches)))

            # Read again using this candidate row as the header.
            table = pd.read_excel(
                workbook,
                sheet_name=sheet_name,
                header=header_index,
                nrows=3,
            )

            print("Detected columns:")
            print(table.columns.tolist())

            print("First three data rows:")
            print(table.to_string(index=False))
            print("-" * 50)
