
"""Check RFQ part codes against the supplier price catalogue."""

import argparse
from pathlib import Path

import pandas as pd

from src.extract_rfq import extract_rfq_lines
from src.load_price_lists import load_catalogue, make_key


def check_rfq_codes(rfq_path, catalogue=None):
    """Report whether each RFQ part code has an exact supplier match."""

    if catalogue is None:
        catalogue = load_catalogue()

    rfq = extract_rfq_lines(str(rfq_path))
    results = []

    for _, item in rfq.iterrows():
        code = str(item.get("part_code", "") or "").strip()
        make = str(item.get("make", "") or "").strip()
        code_key = make_key(code)

        matching_rows = catalogue[
            (
                catalogue["supplier"].astype(str).str.casefold()
                == make.casefold()
            )
            & (
                catalogue["part_code_key"].astype(str)
                == code_key
            )
        ] if code_key else catalogue.iloc[0:0]

        if not code_key:
            status = "MISSING PART CODE"
        elif "XX" in code.upper():
            status = "INCOMPLETE PART CODE"
        elif matching_rows.empty:
            status = "CODE NOT FOUND"
        elif matching_rows["price_inr"].dropna().empty:
            status = "CODE FOUND, PRICE MISSING"
        else:
            prices = matching_rows["price_inr"].dropna().unique()
            if len(prices) > 1:
                status = "MULTIPLE DIFFERENT PRICES"
            else:
                status = "CODE AND PRICE FOUND"

        results.append({
            "line_no": item.get("line_no", ""),
            "make": make,
            "part_code": code,
            "description": item.get("description", ""),
            "qty": item.get("qty", ""),
            "status": status,
            "matched_price_inr": (
                matching_rows["price_inr"].dropna().iloc[0]
                if status == "CODE AND PRICE FOUND"
                else None
            ),
            "source_sheet": (
                ", ".join(
                    matching_rows["source_sheet"]
                    .dropna().astype(str).unique()
                )
                if not matching_rows.empty else ""
            ),
        })

    return pd.DataFrame(results)


def main():
    parser = argparse.ArgumentParser(
        description="Check RFQ codes against supplier price lists."
    )
    parser.add_argument(
        "rfq",
        nargs="?",
        default="data/raw/RFQ_KBS-0147.pdf",
        help="Path to the RFQ file",
    )
    parser.add_argument(
        "--out",
        default="data/processed/rfq_code_check.csv",
        help="Path for the report CSV",
    )
    args = parser.parse_args()

    report = check_rfq_codes(Path(args.rfq))

    output = Path(args.out)
    output.parent.mkdir(parents=True, exist_ok=True)
    report.to_csv(output, index=False)

    print(report.to_string(index=False))
    print(f"\nReport saved to: {output}")
    print("\nStatus summary:")
    print(report["status"].value_counts().to_string())


if __name__ == "__main__":
    main()
