
"""Load ABB and Schneider price lists into one clean table."""

import argparse
import re
from pathlib import Path

import pandas as pd

# File locations
RAW_DIR = Path("data/raw")
DEFAULT_ABB = RAW_DIR / "ABB_Price_List_Feb2026.xlsx"
DEFAULT_SCHNEIDER = RAW_DIR / "Schneider LV Price list Apr 2026.xlsx"
DEFAULT_OUT = Path("data/processed/price_catalogue.csv")

HEADER_ROW = 3
ABB_SHEET = "ABB Price List"

SCHNEIDER_SKIP = {
    "Read Me & Search Guide",
    "Category Technical Data",
}

COLUMNS = [
    "supplier",
    "source_sheet",
    "category",
    "part_code",
    "part_code_key",
    "description",
    "rating_a",
    "poles",
    "breaking_capacity",
    "specs",
    "price_inr",
    "price_status",
    "stocked",
    "page",
]


def make_key(code) -> str:
    """Make part codes consistent for searching."""
    if code is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(code).upper())


def first_number(text) -> str:
    """Extract the first number from text."""
    match = re.search(r"\d+(?:\.\d+)?", str(text))
    return match.group(0) if match else ""


def read_sheet(path: Path, sheet) -> pd.DataFrame:
    """Read a spreadsheet sheet and clean its text."""
    df = pd.read_excel(
        path,
        sheet_name=sheet,
        header=HEADER_ROW,
        dtype=str,
    )
    df = df.fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df.apply(lambda col: col.str.strip())


def clean_prices(raw: pd.Series):
    """Convert price text into numbers and label missing prices."""
    price = pd.to_numeric(
        raw.str.replace(",", "", regex=False),
        errors="coerce",
    )
    status = raw.str.lower().where(price.isna(), "ok")
    status = status.mask(raw == "", "missing")
    return price, status


def parse_abb_ratings(row: pd.Series):
    """Extract product ratings from ABB specification columns."""
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

    if not poles:
        text = row.get("Key word", "") + " " + row.get("Description", "")
        match = re.search(r"\b([1-4])P\b", text)
        poles = match.group(1) if match else ""

    return rating_a, poles, breaking, "; ".join(leftovers)


def load_abb(path: Path) -> pd.DataFrame:
    """Load ABB products into the common catalogue format."""
    raw = read_sheet(path, ABB_SHEET)

    parsed = raw.apply(
        parse_abb_ratings,
        axis=1,
        result_type="expand",
    )
    parsed.columns = [
        "rating_a",
        "poles",
        "breaking_capacity",
        "specs",
    ]

    price, status = clean_prices(raw["Price"])

    out = pd.DataFrame({
        "supplier": "ABB",
        "source_sheet": ABB_SHEET,
        "category": raw["Category"],
        "part_code": raw["Part code"],
        "part_code_key": raw["Part code"].map(make_key),
        "description": (
            raw["Description"] + " | " + raw["Key word"]
        ).str.strip(" |"),
        "rating_a": parsed["rating_a"],
        "poles": parsed["poles"],
        "breaking_capacity": parsed["breaking_capacity"],
        "specs": parsed["specs"],
        "price_inr": price,
        "price_status": status,
        "stocked": raw["Stocked or Not stocked"],
        "page": raw["Price List Page"],
    })

    return out[
        (out["part_code"] != "") | (out["description"] != "")
    ]


SCHNEIDER_NON_TEXT = {
    "Reference", "Unit_MRP", "Page", "Rated_A", "Rating_A",
    "Poles", "Poles/Mount", "Breaking_kA", "BIS", "Notes",
}


def load_schneider_sheet(path: Path, sheet: str) -> pd.DataFrame:
    """Load one Schneider product sheet."""
    raw = read_sheet(path, sheet)

    if "Reference" not in raw.columns or "Unit_MRP" not in raw.columns:
        return pd.DataFrame(columns=COLUMNS)

    raw = raw[raw["Reference"] != ""]

    text_cols = [
        c for c in raw.columns
        if c not in SCHNEIDER_NON_TEXT
    ]
    description = raw[text_cols].apply(
        lambda r: " | ".join(v for v in r if v),
        axis=1,
    )

    def pick(*names):
        for name in names:
            if name in raw.columns:
                return raw[name]
        return pd.Series("", index=raw.index)

    price, status = clean_prices(raw["Unit_MRP"])

    notes = pick("Notes")
    bis = pick("BIS").map(
        lambda v: f"BIS: {v}" if v else ""
    )
    specs = pd.concat([bis, notes], axis=1).apply(
        lambda r: "; ".join(v for v in r if v),
        axis=1,
    )

    poles = pick("Poles", "Poles/Mount")

    return pd.DataFrame({
        "supplier": "Schneider Electric",
        "source_sheet": sheet,
        "category": pick("Category", "Group", "Family"),
        "part_code": raw["Reference"],
        "part_code_key": raw["Reference"].map(make_key),
        "description": description,
        "rating_a": pick("Rated_A", "Rating_A"),
        "poles": poles.map(
            lambda v: first_number(v) if re.match(r"^\d\s*P", v) else ""
        ),
        "breaking_capacity": pick("Breaking_kA").replace("-", ""),
        "specs": specs,
        "price_inr": price,
        "price_status": status,
        "stocked": "",
        "page": pick("Page"),
    })


def load_schneider(path: Path) -> pd.DataFrame:
    """Load all Schneider product sheets."""
    sheets = pd.ExcelFile(path).sheet_names

    parts = [
        load_schneider_sheet(path, sheet)
        for sheet in sheets
        if sheet not in SCHNEIDER_SKIP
    ]

    return pd.concat(parts, ignore_index=True)


def load_catalogue(
    abb_path=DEFAULT_ABB,
    schneider_path=DEFAULT_SCHNEIDER,
) -> pd.DataFrame:
    """Combine both supplier price lists."""
    abb = load_abb(Path(abb_path))
    schneider = load_schneider(Path(schneider_path))

    return pd.concat(
        [abb, schneider],
        ignore_index=True,
    )[COLUMNS]


def add_supplier_prices(
    lines: pd.DataFrame,
    catalogue: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Add supplier list prices to RFQ lines when the match is safe."""
    if catalogue is None:
        catalogue = load_catalogue()

    result = lines.copy()
    result["supplier_price_inr"] = None
    result["supplier_price_status"] = "REVIEW - no exact code match"
    result["supplier_price_source"] = ""

    for index, line in result.iterrows():
        code = str(line.get("part_code", "") or "").strip()
        supplier = str(line.get("make", "") or "").strip()
        code_key = make_key(code)

        if not code_key:
            result.at[index, "supplier_price_status"] = (
                "REVIEW - missing part code"
            )
            continue

        if "XX" in code.upper():
            result.at[index, "supplier_price_status"] = (
                "REVIEW - incomplete part code"
            )
            continue

        hits = catalogue[
            (
                catalogue["supplier"].astype(str).str.casefold()
                == supplier.casefold()
            )
            & (catalogue["part_code_key"] == code_key)
        ]

        if hits.empty:
            result.at[index, "supplier_price_status"] = (
                "REVIEW - code not found"
            )
            continue

        prices = hits["price_inr"].dropna().unique()

        if len(prices) == 0:
            result.at[index, "supplier_price_status"] = (
                "REVIEW - supplier price unavailable"
            )
            continue

        if len(prices) > 1:
            result.at[index, "supplier_price_status"] = (
                "REVIEW - conflicting supplier prices"
            )
            continue

        result.at[index, "supplier_price_inr"] = float(prices[0])
        result.at[index, "supplier_price_source"] = ", ".join(
            hits["source_sheet"].dropna().astype(str).unique()
        )

        if len(hits) > 1:
            result.at[index, "supplier_price_status"] = (
                "LIST PRICE FOUND - duplicate catalogue entries; verify specs"
            )
        else:
            result.at[index, "supplier_price_status"] = (
                "LIST PRICE FOUND - verify specs"
            )

    return result


def print_report(df: pd.DataFrame) -> None:
    """Print a summary of the loaded catalogue."""
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
        print(
            dup.sort_values("part_code_key")[
                ["supplier", "source_sheet", "part_code", "price_inr"]
            ].head(10).to_string(index=False)
        )


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
