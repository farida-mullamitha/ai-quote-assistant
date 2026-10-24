import re

import pandas as pd

from src.extract_rfq import extract_rfq_lines
from src.load_price_lists import load_catalogue


# Specifications we can look for in a product description.
def extract_specs(text):
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
        "electrical": "ELECTRICALLY OPERATED" in text
        or "ELECTRIC OPERATION" in text,
    }


def get_text(row):
    """Combine catalogue fields so we can search descriptions and specs."""
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

    return " ".join(
        str(row.get(field, ""))
        for field in fields
        if pd.notna(row.get(field, ""))
    ).upper()


def number_from_row(row, field, pattern):
    """Read a number from a catalogue field, if possible."""
    value = row.get(field)

    if value is None or pd.isna(value):
        return None

    match = re.search(pattern, str(value).upper())
    return float(match.group(1)) if match else None


def score_candidate(rfq_description, make, row):
    candidate_text = get_text(row)
    rfq_specs = extract_specs(rfq_description)
    score = 0
    reasons = []
    conflicts = []

    # Prefer the requested manufacturer.
    supplier = str(row.get("supplier", "")).upper()
    if make and make.upper() in supplier:
        score += 3
        reasons.append("manufacturer matches")

    # These two RFQ lines specifically request air circuit breakers.
    if "AIR CIRCUIT BREAKER" in rfq_description.upper():
        if "AIR CIRCUIT BREAKER" in candidate_text or " ACB " in (
            f" {candidate_text} "
        ):
            score += 5
            reasons.append("air circuit breaker wording matches")
        else:
            conflicts.append("product family not confirmed")

    # Compare current rating.
    requested_a = rfq_specs["current_a"]
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

    # Compare pole count.
    requested_poles = rfq_specs["poles"]
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

    # Compare breaking capacity.
    requested_ka = rfq_specs["breaking_ka"]
    candidate_ka = number_from_row(
        row, "breaking_capacity", r"(\d+(?:\.\d+)?)"
    )

    if requested_ka is not None and candidate_ka is not None:
        if requested_ka == candidate_ka:
            score += 4
            reasons.append(f"breaking capacity matches ({requested_ka:g}kA)")
        else:
            score -= 6
            conflicts.append(
                f"breaking capacity differs: requested {requested_ka:g}kA, "
                f"catalogue {candidate_ka:g}kA"
            )

    # Compare important description details.
    for label, phrase in [
        ("draw-out mounting", "DRAW-OUT"),
        ("LSIG protection", "LSIG"),
        ("manual operation", "MANUAL"),
    ]:
        if phrase in rfq_description.upper():
            if phrase in candidate_text or (
                phrase == "DRAW-OUT" and "DRAWOUT" in candidate_text
            ):
                score += 2
                reasons.append(f"{label} found")
            else:
                conflicts.append(f"{label} not confirmed")

    if "ELECTRICALLY OPERATED" in rfq_description.upper():
        if "ELECTRICALLY OPERATED" in candidate_text:
            score += 2
            reasons.append("electrical operation found")
        else:
            conflicts.append("electrical operation not confirmed")

    return score, reasons, conflicts



def find_missing_code_candidates(rfq_lines, catalogue=None):
    """Return ranked catalogue candidates for RFQ lines with missing codes."""

    if isinstance(rfq_lines, pd.DataFrame):
        rfq_lines = rfq_lines.to_dict("records")

    if catalogue is None:
        catalogue = load_catalogue()

    if isinstance(catalogue, pd.DataFrame):
        catalogue_rows = catalogue.to_dict("records")
    else:
        catalogue_rows = catalogue

    missing_lines = [
        line for line in rfq_lines
        if not str(line.get("part_code", "") or "").strip()
    ]

    report = []

    for line in missing_lines:
        description = str(line.get("description", "") or "")
        make = str(line.get("make", "") or "")
        line_no = line.get("line_no", "?")

        # This prototype currently supports air circuit breakers only.
        if "AIR CIRCUIT BREAKER" not in description.upper():
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

        results.sort(key=lambda item: item[0], reverse=True)

        if not results:
            report.append({
                "line_no": line_no,
                "rfq_description": description,
                "requested_make": make,
                "rank": None,
                "score": None,
                "part_code": "",
                "candidate_description": "",
                "price": "",
                "supporting_details": "",
                "checks_needed": "No candidates found in catalogue",
                "decision": "HUMAN REVIEW REQUIRED",
            })
            continue

        for rank, (score, reasons, conflicts, row) in enumerate(
            results[:5], start=1
        ):
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
                "price": row.get(
                    "price_inr", row.get("price_status", "Unknown")
                ),
                "supporting_details": "; ".join(reasons),
                "checks_needed": "; ".join(conflicts)
                or "Verify full specifications",
                "decision": "HUMAN REVIEW REQUIRED",
            })

    return report


def main():
    rfq_path = "data/raw/RFQ_KBS-0147.pdf"
    rfq_lines = extract_rfq_lines(rfq_path)

    candidates = find_missing_code_candidates(rfq_lines)

    if not candidates:
        print("No candidates found for missing-code RFQ lines.")
        return

    for candidate in candidates:
        print(candidate)


if __name__ == "__main__":
    main()

