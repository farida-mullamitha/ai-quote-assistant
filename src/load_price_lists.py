"""Load the ABB and Schneider price lists into ONE clean table.

Run from the project root:
    python src/load_price_lists.py

What it does
------------
1. Reads the ABB workbook (one big sheet) and every Schneider product sheet.
2. Renames the different column layouts to one common set of columns.
3. Cleans the part codes and prices so they can be searched reliably.
4. Saves everything to data/processed/price_catalogue.csv and prints a report.
"""

import argparse
import re
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------- settings
RAW_DIR = Path("data/raw")
DEFAULT_ABB = RAW_DIR / "ABB_Price_List_Feb2026.xlsx"
DEFAULT_SCHNEIDER = RAW_DIR / "Schneider LV Price list Apr 2026.xlsx"
DEFAULT_OUT = Path("data/processed/price_catalogue.csv")

# Every sheet has: row 1 title, row 2 source note, row 3 blank, row 4 headers.
# pandas counts from 0, so the header row is number 3.
HEADER_ROW = 3

ABB_SHEET = "ABB Price List"

# Schneider sheets that are NOT product tables (no part codes / prices).
SCHNEIDER_SKIP = {"Read Me & Search Guide", "Category Technical Data"}

# The final, common columns. Every row from both suppliers has all of these.
COLUMNS = [
    "supplier",           # ABB or Schneider Electric
    "source_sheet",       # which sheet the row came from (for checking)
    "category",           # MCCB, ACB, Contactor, ...
    "part_code",          # exactly as printed in the price list
    "part_code_key",      # cleaned version used for searching
    "description",        # human-readable description
    "rating_a",           # rated current in amps (number, if known)
    "poles",              # number of poles (number, if known)
    "breaking_capacity",  # ABB: class letter (e.g. N). Schneider: kA number.
    "specs",              # other specs / notes, kept as text so nothing is lost
    "price_inr",          # list price (MRP) in INR, number. Blank if no price.
    "price_status",        # ok / on request / not listed / missing
    "stocked",            # ABB only: Stocked / Not stocked
    "page",               # page in the supplier's PDF, for checking
]


# ---------------------------------------------------------------- helpers
def make_key(code) -> str:
    """Clean a part code for searching: upper case, letters and digits only.

    '1SYS 271012 R0164' and '1sys271012r0164' both become '1SYS271012R0164'.
    """
    if code is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(code).upper())


def first_number(text) -> str:
    """Return the first number inside a text like 'In (A): 125' -> '125'."""
    match = re.search(r"\d+(?:\.\d+)?", str(text))
    return match.group(0) if match else ""


def read_sheet(path: Path, sheet) -> pd.DataFrame:
    """Read one sheet as text (so '29450' stays a code, not a number)."""
    df = pd.read_excel(path, sheet_name=sheet, header=HEADER_ROW, dtype=str)
    df = df.fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df.apply(lambda col: col.str.strip())


def clean_prices(raw: pd.Series):
    """Turn price text into numbers. Keep a status for the non-numeric ones.

    '19,900'      -> price 19900, status 'ok'
    'On Request'  -> price blank, status 'on request'
    ''            -> price blank, status 'missing'
    """
    price = pd.to_numeric(raw.str.replace(",", "", regex=False), errors="coerce")
    status = raw.str.lower().where(price.isna(), "ok")
    status = status.mask(raw == "", "missing")
    return price, status


# ---------------------------------------------------------------- ABB
def parse_abb_ratings(row: pd.Series):
    """Pull rating, poles and breaking capacity out of 'Rating 1..5' text.

    Each Rating cell looks like 'Header: value', for example
    'In (A): 125', 'Poles: 3', 'Breaking capacity: N'.
    The headers differ by product type, so we match on keywords.
    """
    rating_a = poles = breaking = ""
    leftovers = []
    for col in ["Rating 1", "Rating 2", "Rating 3", "Rating 4", "Rating 5"]:
        cell = row.get(col, "")
        if not cell:
            continue
        header, _, value = cell.partition(":")
        h = header.lower()
        if "pole" in h and not poles:
            poles = first_number(value)
        elif (
            re.search(r"\bin\s*\(a\)|^in$|current rating|rated current", h)
            and not rating_a
        ):
            rating_a = first_number(value)
        elif "breaking capacity" in h and not breaking:
            breaking = value.strip()
        else:
            leftovers.append(cell)
    if row.get("Other Ratings / Specs"):
        leftovers.append(row["Other Ratings / Specs"])
    # Some rows only state the poles in the type code, e.g. 'A1 TMF N 3P'.
    if not poles:
        match = re.search(r"\b([1-4])P\b", row.get("Key word", "") + " " + row.get("Description", ""))
        poles = match.group(1) if match else ""
    return rating_a, poles, breaking, "; ".join(leftovers)


def load_abb(path: Path) -> pd.DataFrame:
    raw = read_sheet(path, ABB_SHEET)
    parsed = raw.apply(parse_abb_ratings, axis=1, result_type="expand")
    parsed.columns = ["rating_a", "poles", "breaking_capacity", "specs"]
    price, status = clean_prices(raw["Price"])

    out = pd.DataFrame(
        {
            "supplier": "ABB",
            "source_sheet": ABB_SHEET,
            "category": raw["Category"],
            "part_code": raw["Part code"],
            "part_code_key": raw["Part code"].map(make_key),
            "description": (raw["Description"] + " | " + raw["Key word"]).str.strip(" |"),
            "rating_a": parsed["rating_a"],
            "poles": parsed["poles"],
            "breaking_capacity": parsed["breaking_capacity"],
            "specs": parsed["specs"],
            "price_inr": price,
            "price_status": status,
            "stocked": raw["Stocked or Not stocked"],
            "page": raw["Price List Page"],
        }
    )
    # Drop completely empty spacer rows (no description and no code).
    return out[(out["part_code"] != "") | (out["description"] != "")]


# ---------------------------------------------------------------- Schneider
# Columns that are NOT part of the description text.
SCHNEIDER_NON_TEXT = {
    "Reference", "Unit_MRP", "Page", "Rated_A", "Rating_A",
    "Poles", "Poles/Mount", "Breaking_kA", "BIS", "Notes",
}


def load_schneider_sheet(path: Path, sheet: str) -> pd.DataFrame:
    raw = read_sheet(path, sheet)
    if "Reference" not in raw.columns or "Unit_MRP" not in raw.columns:
        return pd.DataFrame(columns=COLUMNS)  # not a product sheet
    raw = raw[raw["Reference"] != ""]

    # Description = every descriptive column joined, e.g.
    # 'EasyPact CVS | 16-630A | MCCB - Distribution | TMD Fixed ...'
    text_cols = [c for c in raw.columns if c not in SCHNEIDER_NON_TEXT]
    description = raw[text_cols].apply(
        lambda r: " | ".join(v for v in r if v), axis=1
    )

    def pick(*names):
        for name in names:
            if name in raw.columns:
                return raw[name]
        return pd.Series("", index=raw.index)

    price, status = clean_prices(raw["Unit_MRP"])
    notes = pick("Notes")
    bis = pick("BIS").map(lambda v: f"BIS: {v}" if v else "")
    specs = pd.concat([bis, notes], axis=1).apply(
        lambda r: "; ".join(v for v in r if v), axis=1
    )
    poles = pick("Poles", "Poles/Mount")

    return pd.DataFrame(
        {
            "supplier": "Schneider Electric",
            "source_sheet": sheet,
            "category": pick("Category", "Group", "Family"),
            "part_code": raw["Reference"],
            "part_code_key": raw["Reference"].map(make_key),
            "description": description,
            "rating_a": pick("Rated_A", "Rating_A"),
            # keep only a clean number if the cell is like '3P' or '3Ph+3W'
            "poles": poles.map(lambda v: first_number(v) if re.match(r"^\d\s*P", v) else ""),
            "breaking_capacity": pick("Breaking_kA").replace("-", ""),
            "specs": specs,
            "price_inr": price,
            "price_status": status,
            "stocked": "",
            "page": pick("Page"),
        }
    )


def load_schneider(path: Path) -> pd.DataFrame:
    sheets = pd.ExcelFile(path).sheet_names
    parts = [
        load_schneider_sheet(path, s) for s in sheets if s not in SCHNEIDER_SKIP
    ]
    return pd.concat(parts, ignore_index=True)


# ---------------------------------------------------------------- main
def load_catalogue(abb_path=DEFAULT_ABB, schneider_path=DEFAULT_SCHNEIDER) -> pd.DataFrame:
    """Return one table with every ABB and Schneider product row."""
    abb = load_abb(Path(abb_path))
    schneider = load_schneider(Path(schneider_path))
    return pd.concat([abb, schneider], ignore_index=True)[COLUMNS]


def print_report(df: pd.DataFrame) -> None:
    print(f"Total rows: {len(df)}")
    print("\nRows per supplier / sheet:")
    print(df.groupby(["supplier", "source_sheet"]).size().to_string())

    print("\nPrice status:")
    print(df["price_status"].value_counts().to_string())

    no_code = df[df["part_code_key"] == ""]
    print(f"\nRows with NO part code: {len(no_code)}")

    coded = df[df["part_code_key"] != ""]
    dup = coded[coded.duplicated("part_code_key", keep=False)]
    print(f"Rows sharing a part code with another row: {len(dup)}")
    if len(dup):
        print(dup.sort_values("part_code_key")[["supplier", "source_sheet", "part_code", "price_inr"]].head(10).to_string(index=False))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--abb", default=DEFAULT_ABB, type=Path)
    parser.add_argument("--schneider", default=DEFAULT_SCHNEIDER, type=Path)
    parser.add_argument("--out", default=DEFAULT_OUT, type=Path)
    args = parser.parse_args()

    catalogue = load_catalogue(args.abb, args.schneider)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    catalogue.to_csv(args.out, index=False)
    print_report(catalogue)
    print(f"\nSaved to {args.out}")
