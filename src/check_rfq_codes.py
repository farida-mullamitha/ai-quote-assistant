"""Read an RFQ file and check which of its part codes exist in the price lists.

Run from the project root:
    python src/check_rfq_codes.py data/raw/RFQ_KBS-0147.pdf

The RFQ can be a PDF, Word (.docx) or Excel (.xlsx) file. Nothing is typed
in by hand: the lines come from extract_rfq.py, the prices from
load_price_lists.py.

This is the "coverage check": it shows how many RFQ lines a plain
part-code lookup can answer, before any AI or fuzzy matching.
"""

import sys

import pandas as pd

from extract_rfq import extract_rfq_lines
from load_price_lists import load_catalogue, make_key



def check_lines(lines: pd.DataFrame, catalogue: pd.DataFrame) -> pd.DataFrame:
    results = []

    for line in lines.itertuples():
        base = {
            "line": line.line_no,
            "make": line.make,
            "rfq_code": line.part_code,
        }

        # No code was supplied by the customer.
        if not line.part_code:
            note = f" ({line.code_note})" if line.code_note else ""
            results.append({
                **base,
                "result": "HUMAN REVIEW NEEDED - MISSING CODE" + note,
                "sheet": "",
                "price_inr": "",
            })
            continue

        # Find matching products for the requested supplier and code.
        hits = catalogue[
            (catalogue["supplier"] == line.make)
            & (
                catalogue["part_code_key"]
                == make_key(line.part_code)
            )
        ]

        if hits.empty:
            results.append({
                **base,
                "result": "HUMAN REVIEW NEEDED - NOT FOUND",
                "sheet": "",
                "price_inr": "",
            })
            continue

        # Compare all matching prices before deciding the status.
        different_prices = (
            hits["price_inr"].nunique(dropna=False) > 1
        )

        incomplete_code = "XX" in line.part_code.upper()

        for _, row in hits.iterrows():
            reasons = []

            if len(hits) > 1:
                if different_prices:
                    reasons.append("PRICE CONFLICT")
                else:
                    reasons.append("DUPLICATE MATCH")

            if incomplete_code:
                reasons.append("INCOMPLETE CODE")

            if pd.isna(row["price_inr"]):
                reasons.append("PRICE UNAVAILABLE")

            if reasons:
                status = "HUMAN REVIEW NEEDED - " + ", ".join(reasons)
            else:
                status = "MATCHED - VERIFY SPECS"

            if pd.notna(row["price_inr"]):
                price = f"{row['price_inr']:,.0f}"
            else:
                price = row["price_status"]

            results.append({
                **base,
                "result": status,
                "sheet": row["source_sheet"],
                "price_inr": price,
            })

    return pd.DataFrame(results)

def main() -> None:
    if len(sys.argv) != 2:
        sys.exit("Usage: python src/check_rfq_codes.py <rfq file>")

    lines = extract_rfq_lines(sys.argv[1])
    report = check_lines(lines, load_catalogue())

    pd.set_option("display.width", 200, "display.max_colwidth", 60)
    print(report.to_string(index=False))
    print(f"\n{report['line'].nunique()} RFQ lines checked.")
    print(f"{len(report)} report rows, including duplicate matches.")

    review_lines = report.loc[
        report["result"].str.startswith("HUMAN REVIEW NEEDED"),
        "line",
    ].nunique()

    print(f"{review_lines} RFQ lines need human review.")


if __name__ == "__main__":
    main()