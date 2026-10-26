import re
import pandas as pd
from src.extract_rfq import extract_rfq_lines
from src.load_price_lists import load_catalogue


def is_missing_or_incomplete(code):
    """Identify blank codes and common placeholder codes."""
    if code is None or pd.isna(code):
        return True

    code = str(code).strip().upper()

    if not code:
        return True

    placeholders = {
        "XX", "XXX", "TBC", "TBD", "N/A", "NA",
        "UNKNOWN", "NOT PROVIDED", "NOT AVAILABLE",
    }

    return (
        code in placeholders
        or "XX" in code
        or "TBC" in code
    )


def extract_specs(text):
    """Extract basic technical specifications from a description."""
    text = str(text or "").upper()

    current = re.search(r"\b(\d+(?:\.\d+)?)\s*A\b", text)
    poles = re.search(r"\b([1-4])\s*(?:POLES?|P)\b", text)
    breaking = re.search(r"\b(\d+(?:\.\d+)?)\s*KA\b", text)

    return {
        "current_a": float(current.group(1)) if current else None,
        "poles": int(poles.group(1)) if poles else None,
        "breaking_ka": float(breaking.group(1)) if breaking else None,
        "draw_out": "DRAW-OUT" in text or "DRAWOUT" in text,
        "lsig": "LSIG" in text,
        "manual": "MANUAL" in text,
        "electrical": (
            "ELECTRICALLY OPERATED" in text
            or "ELECTRIC OPERATION" in text
        ),
    }


def get_text(row):
    """Combine catalogue fields for searching."""
    fields = [
        "supplier",
        "source_sheet",
        "category",
        "description",
        "specs",
        "rating_a",
        "poles",
        "breaking_capacity",
    ]

    values = []

    for field in fields:
        value = row.get(field)

        if value is not None and not pd.isna(value):
            values.append(str(value))

    return " ".join(values).upper()


def number_from_row(row, field, pattern):
    """Extract a numeric specification from a catalogue field."""
    value = row.get(field)

    if value is None or pd.isna(value):
        return None

    match = re.search(pattern, str(value).upper())

    return float(match.group(1)) if match else None


def score_candidate(rfq_description, make, row):
    """Score a possible product match and record specification conflicts."""
    candidate_text = get_text(row)
    requested_specs = extract_specs(rfq_description)

    score = 0
    reasons = []
    conflicts = []

    supplier = str(row.get("supplier", "") or "").upper()

    if make and make.upper() in supplier:
        score += 3
        reasons.append("manufacturer matches")

    # This version has detailed matching rules for air circuit breakers.
    if "AIR CIRCUIT BREAKER" in rfq_description.upper():
        is_acb = (
            "AIR CIRCUIT BREAKER" in candidate_text
            or " ACB " in f" {candidate_text} "
        )

        if is_acb:
            score += 5
            reasons.append("air circuit breaker wording matches")
        else:
            conflicts.append("product family not confirmed")

    requested_a = requested_specs["current_a"]
    candidate_a = number_from_row(
        row, "rating_a", r"(\d+(?:\.\d+)?)"
    )

    if requested_a is not None and candidate_a is not None:
        if requested_a == candidate_a:
            score += 4
            reasons.append(f"current matches ({requested_a:g}A)")
        else:
            score -= 6
            conflicts.append(
                f"current differs: requested {requested_a:g}A, "
                f"catalogue {candidate_a:g}A"
            )

    requested_poles = requested_specs["poles"]
    candidate_poles = number_from_row(
        row, "poles", r"\b([1-4])\b"
    )

    if requested_poles is not None and candidate_poles is not None:
        if requested_poles == candidate_poles:
            score += 3
            reasons.append(f"poles match ({requested_poles}P)")
        else:
            score -= 6
            conflicts.append(
                f"pole count differs: requested {requested_poles}, "
                f"catalogue {candidate_poles:g}"
            )

    requested_ka = requested_specs["breaking_ka"]
    candidate_ka = number_from_row(
        row, "breaking_capacity", r"(\d+(?:\.\d+)?)"
    )

    if requested_ka is not None and candidate_ka is not None:
        if requested_ka == candidate_ka:
            score += 4
            reasons.append(
                f"breaking capacity matches ({requested_ka:g}kA)"
            )
        else:
            score -= 6
            conflicts.append(
                f"breaking capacity differs: requested {requested_ka:g}kA, "
                f"catalogue {candidate_ka:g}kA"
            )

    checks = [
        ("DRAW-OUT", "DRAWOUT", "draw-out mounting"),
        ("LSIG", "LSIG", "LSIG protection"),
        ("MANUAL", "MANUAL", "manual operation"),
    ]

    for phrase, alternative, label in checks:
        if phrase in rfq_description.upper():
            if phrase in candidate_text or alternative in candidate_text:
                score += 2
                reasons.append(f"{label} found")
            else:
                conflicts.append(f"{label} not confirmed")

    if "ELECTRICALLY OPERATED" in rfq_description.upper():
        if (
            "ELECTRICALLY OPERATED" in candidate_text
            or "ELECTRIC OPERATION" in candidate_text
        ):
            score += 2
            reasons.append("electrical operation found")
        else:
            conflicts.append("electrical operation not confirmed")

    return score, reasons, conflicts


def find_missing_code_candidates(rfq_lines, catalogue=None):
    """Rank possible catalogue matches for RFQ items with unusable codes."""
    if isinstance(rfq_lines, pd.DataFrame):
        rfq_lines = rfq_lines.to_dict("records")

    if catalogue is None:
        catalogue = load_catalogue()

    if isinstance(catalogue, pd.DataFrame):
        catalogue_rows = catalogue.to_dict("records")
    else:
        catalogue_rows = catalogue

    report = []

    for line in rfq_lines:
        code = line.get("part_code", "")

        if not is_missing_or_incomplete(code):
            continue

        description = str(line.get("description", "") or "")
        make = str(line.get("make", "") or "")
        line_no = line.get("line_no", "?")

        # Detailed candidate scoring currently supports ACB products.
        if "AIR CIRCUIT BREAKER" not in description.upper():
            report.append({
                "line_no": line_no,
                "rfq_description": description,
                "requested_make": make,
                "rank": None,
                "score": None,
                "part_code": "",
                "candidate_description": "",
                "source_sheet": "",
                "price": None,
                "supporting_details": "",
                "checks_needed": (
                    "Product family not supported by current matching rules"
                ),
                "decision": "HUMAN REVIEW REQUIRED",
            })
            continue

        results = []

        for row in catalogue_rows:
            supplier = str(row.get("supplier", "") or "").upper()

            if make and make.upper() not in supplier:
                continue

            candidate_text = get_text(row)

            if (
                "AIR CIRCUIT BREAKER" not in candidate_text
                and " ACB " not in f" {candidate_text} "
            ):
                continue

            score, reasons, conflicts = score_candidate(
                description, make, row
            )

            results.append((score, reasons, conflicts, row))

        results.sort(key=lambda result: result[0], reverse=True)

        if not results:
            report.append({
                "line_no": line_no,
                "rfq_description": description,
                "requested_make": make,
                "rank": None,
                "score": None,
                "part_code": "",
                "candidate_description": "",
                "source_sheet": "",
                "price": None,
                "supporting_details": "",
                "checks_needed": "No matching candidates found in catalogue",
                "decision": "HUMAN REVIEW REQUIRED",
            })
            continue

        for rank, (score, reasons, conflicts, row) in enumerate(
            results[:5], start=1
        ):
            price = row.get("price_inr")

            if price is not None and pd.isna(price):
                price = None

            report.append({
                "line_no": line_no,
                "rfq_description": description,
                "requested_make": make,
                "rank": rank,
                "score": score,
                "part_code": row.get("part_code", ""),
                "candidate_description": row.get("description", ""),
                "source_sheet": row.get(
                    "source_sheet", row.get("category", "")
                ),
                "price": price,
                "supporting_details": "; ".join(reasons),
                "checks_needed": (
                    "; ".join(conflicts)
                    if conflicts
                    else "Verify full specifications and exact product"
                ),
                "decision": "HUMAN REVIEW REQUIRED",
            })

    return report


def main():
    rfq_path = "data/raw/RFQ_KBS-0147.pdf"

    rfq_lines = extract_rfq_lines(rfq_path)
    candidates = find_missing_code_candidates(rfq_lines)

    if not candidates:
        print("No missing or incomplete part codes found.")
        return

    report = pd.DataFrame(candidates)

    output_path = "data/processed/missing_code_candidates.csv"
    report.to_csv(output_path, index=False)

    print(f"Candidate report saved to: {output_path}")
    print(report.to_string(index=False))


if __name__ == "__main__":
    main()
