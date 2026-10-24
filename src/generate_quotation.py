
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


PROJECT_ROOT = Path(__file__).resolve().parent.parent

TEMPLATE_PATH = PROJECT_ROOT / "data" / "raw" / "IEPCI Quotation .xlsx"
RFQ_PATH = PROJECT_ROOT / "data" / "raw" / "RFQ_KBS-0147.pdf"
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"

FIRST_ITEM_ROW = 13
LAST_TEMPLATE_ITEM_ROW = 29


def copy_row_format(ws, source_row, target_row):
    """Copy cell formatting and row height to a new item row."""
    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)

        if source.has_style:
            target._style = copy(source._style)

        if source.number_format:
            target.number_format = source.number_format

        if source.alignment:
            target.alignment = copy(source.alignment)

        if source.protection:
            target.protection = copy(source.protection)

    height = ws.row_dimensions[source_row].height
    if height is not None:
        ws.row_dimensions[target_row].height = height


def expand_item_table(ws, item_count):
    """Insert item rows and correctly restore merged cells below the table."""
    capacity = LAST_TEMPLATE_ITEM_ROW - FIRST_ITEM_ROW + 1

    if item_count <= capacity:
        return None

    extra_rows = item_count - capacity
    insert_at = LAST_TEMPLATE_ITEM_ROW + 1

    # Save and temporarily remove merges that could be affected by insertion.
    affected_merges = [
        str(merged_range)
        for merged_range in list(ws.merged_cells.ranges)
        if merged_range.max_row >= insert_at
    ]

    for merged_range in affected_merges:
        ws.unmerge_cells(merged_range)

    # openpyxl does not automatically move row heights when inserting rows.
    saved_heights = {
        row: ws.row_dimensions[row].height
        for row in list(ws.row_dimensions)
        if row >= insert_at
    }

    for row in saved_heights:
        del ws.row_dimensions[row]

    ws.insert_rows(insert_at, extra_rows)

    # Restore row heights at their new positions.
    for row, height in saved_heights.items():
        if height is not None:
            ws.row_dimensions[row + extra_rows].height = height

    # Format the additional item rows using the final template item row.
    for row in range(insert_at, insert_at + extra_rows):
        copy_row_format(ws, LAST_TEMPLATE_ITEM_ROW, row)

    # Restore the original merged cells at their shifted positions.
    for merged_range in affected_merges:
        min_col, min_row, max_col, max_row = range_boundaries(
            merged_range
        )

        if min_row >= insert_at:
            min_row += extra_rows
            max_row += extra_rows
        elif max_row >= insert_at:
            max_row += extra_rows

        ws.merge_cells(
            start_row=min_row,
            start_column=min_col,
            end_row=max_row,
            end_column=max_col,
        )

    # Create the usual item-row merges for each additional row.
    for row in range(insert_at, insert_at + extra_rows):
        for start_col, end_col in [
            (3, 8),    # C:H — Description
            (11, 12),  # K:L — Unit rate
            (13, 14),  # M:N — Line total
        ]:
            ws.merge_cells(
                start_row=row,
                start_column=start_col,
                end_row=row,
                end_column=end_col,
            )

    last_item_row = FIRST_ITEM_ROW + item_count - 1
    subtotal_row = last_item_row + 1
    tax_row = subtotal_row + 1
    packing_row = subtotal_row + 2
    freight_row = subtotal_row + 3
    inspection_row = subtotal_row + 4
    total_row = subtotal_row + 5

    return {
        "last_item_row": last_item_row,
        "subtotal_row": subtotal_row,
        "tax_row": tax_row,
        "packing_row": packing_row,
        "freight_row": freight_row,
        "inspection_row": inspection_row,
        "total_row": total_row,
    }


def generate_quotation(rfq_path=RFQ_PATH, output_path=None, rfq_df=None):
    """Create a draft Excel quotation from an RFQ and the IEPCI template."""
    from src.extract_rfq import extract_rfq_lines

    if rfq_df is None:
        rfq_df = extract_rfq_lines(str(rfq_path))

    if rfq_df.empty:
        raise ValueError("The RFQ contains no extracted line items.")

    if len(rfq_df) > 100:
        raise ValueError(
            "This RFQ has more than 100 items; review it before generating."
        )

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Quotation template not found: {TEMPLATE_PATH}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if output_path is None:
        output_path = OUTPUT_DIR / "IEPCI_Quotation_DRAFT.xlsx"
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = load_workbook(TEMPLATE_PATH)
    ws = workbook["Format"]

    layout = expand_item_table(ws, len(rfq_df))

    if layout is None:
        last_item_row = FIRST_ITEM_ROW + len(rfq_df) - 1
        subtotal_row = LAST_TEMPLATE_ITEM_ROW + 1
        tax_row = subtotal_row + 1
        packing_row = subtotal_row + 2
        freight_row = subtotal_row + 3
        inspection_row = subtotal_row + 4
        total_row = subtotal_row + 5
    else:
        last_item_row = layout["last_item_row"]
        subtotal_row = layout["subtotal_row"]
        tax_row = layout["tax_row"]
        packing_row = layout["packing_row"]
        freight_row = layout["freight_row"]
        inspection_row = layout["inspection_row"]
        total_row = layout["total_row"]

    # Fill only the RFQ details that are known.
    ws["M6"] = "IEPCI/QTN/2026/DRAFT-01"
    ws["M8"] = "KBS/RFQ/2026/0147"
    ws["A5"] = "QUOTATION - DRAFT (PRICES TO BE VERIFIED)"

    # Populate all extracted RFQ items.
    for index, (_, item) in enumerate(rfq_df.iterrows()):
        row = FIRST_ITEM_ROW + index

        line_no = item.get("line_no", index + 1)
        make = str(item.get("make", "") or "").strip()
        part_code = str(item.get("part_code", "") or "").strip()
        description = str(item.get("description", "") or "").strip()
        code_note = str(item.get("code_note", "") or "").strip()
        qty = item.get("qty", None)
        unit = str(item.get("unit", "") or "").strip()

        ws.cell(row, 1).value = (
            int(line_no) if str(line_no).isdigit() else index + 1
        )
        ws.cell(row, 2).value = part_code or None

        description_parts = []
        if make:
            description_parts.append(make)
        if description:
            description_parts.append(description)
        if code_note:
            description_parts.append(f"REVIEW: {code_note}")

        ws.cell(row, 3).value = " — ".join(description_parts)
        ws.cell(row, 9).value = unit
        ws.cell(row, 10).value = int(qty) if str(qty).isdigit() else qty

        # No unapproved selling price is entered.
        unit_rate = item.get("unit_rate")
        approved = item.get("price_approved", False)

        if approved and unit_rate is not None and str(unit_rate).strip().lower() not in ("", "nan"):
            ws.cell(row, 11).value = float(unit_rate)
        else:
            ws.cell(row, 11).value = None
        ws.cell(row, 13).value = (
            f'=IF(OR(K{row}="",J{row}=""),"",K{row}*J{row})'
        )

    # Keep labels aligned with the template's commercial section.
    ws.cell(subtotal_row, 3).value = "Subtotal"
    ws.cell(tax_row, 3).value = "Taxes"
    ws.cell(packing_row, 3).value = "Packing and Forwarding"
    ws.cell(freight_row, 3).value = "Freight and Insurance"
    ws.cell(inspection_row, 3).value = "Inspection"
    ws.cell(total_row, 3).value = "Total"

    # Calculate totals only when rates have been entered.
    ws.cell(subtotal_row, 13).value = (
        f'=IF(COUNT(K{FIRST_ITEM_ROW}:K{last_item_row})=0,"",'
        f'SUM(M{FIRST_ITEM_ROW}:M{last_item_row}))'
    )

    # The template uses 18% GST; applicability must be confirmed before issue.
    ws.cell(tax_row, 10).value = 0.18
    ws.cell(tax_row, 13).value = (
        f'=IF(M{subtotal_row}="","",M{subtotal_row}*18%)'
    )

    # Packing is a direct amount; leave it blank until approved.
    ws.cell(packing_row, 13).value = None

    ws.cell(freight_row, 13).value = (
        f'=IF(OR(K{freight_row}="",J{freight_row}=""),"",'
        f'K{freight_row}*J{freight_row})'
    )

    ws.cell(inspection_row, 13).value = (
        f'=IF(OR(K{inspection_row}="",J{inspection_row}=""),"",'
        f'K{inspection_row}*J{inspection_row})'
    )

    ws.cell(total_row, 13).value = (
        f'=IF(M{subtotal_row}="","",'
        f'M{subtotal_row}+M{tax_row}+N(M{packing_row})'
        f'+N(M{freight_row})+N(M{inspection_row}))'
    )

    # Keep the draft warning in the title instead of overwriting row 36.
    # Row 36 and the following rows remain part of the quotation layout.

    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass

    workbook.save(output_path)
    return output_path


if __name__ == "__main__":
    path = generate_quotation()
    print(f"Draft quotation created: {path}")
    print("Unit rates and final totals have not been approved.")
