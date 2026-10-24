# IEPCI Quote Assistant

Reads a customer RFQ (PDF, Word or Excel), finds each item in the ABB and
Schneider price lists, and produces a DRAFT Excel quotation from the IEPCI
template. A person reviews everything. Selling prices are never filled in
automatically: a rate only appears in the quotation after someone types it and
ticks "Rate approved".

## Setup (once)

```
python -m venv .venv
.venv\Scripts\activate          # Windows PowerShell
pip install -r requirements.txt
```

Put these files in `data/raw/` with exactly these names:

| File | What it is |
|---|---|
| `ABB_Price_List_Feb2026.xlsx` | ABB price list |
| `Schneider LV Price list Apr 2026.xlsx` | Schneider price list |
| `IEPCI Quotation .xlsx` | quotation template (note the space before `.xlsx`) |
| `RFQ_KBS-0147.pdf` | the Kestrel test RFQ |

## Run

```
streamlit run app.py                                  # the app
python -m pytest                                      # automated tests
python -m src.check_rfq_codes data/raw/RFQ_KBS-0147.pdf   # part-code check
python -m src.match_missing_codes data/raw/RFQ_KBS-0147.pdf
python -m src.generate_quotation                      # draft quote from the test RFQ
python src/load_price_lists.py                        # rebuild data/processed/price_catalogue.csv
python src/extract_rfq.py data/raw/RFQ_KBS-0147.pdf   # show extracted lines + details
```

Always run from the project root.

## How the pieces fit

```
RFQ file --extract_rfq.py--> line items + RFQ details (customer, RFQ no., ...)
                                   |
price lists --load_price_lists.py--+--> one catalogue table (6,366 rows)
                                   |
        check_rfq_codes.py ........+--> per line: found / price / needs review
        match_missing_codes.py ....+--> ranked candidates for lines with no code
                                   |
         person reviews and approves rates in app.py
                                   |
        generate_quotation.py -----+--> Excel draft from the IEPCI template
```

| File | Job |
|---|---|
| `app.py` | Streamlit app: upload, review, matching, catalogue, pricing, download |
| `src/extract_rfq.py` | Tables out of PDF/Word/Excel; header details; code-vs-note split |
| `src/load_price_lists.py` | ABB and Schneider layouts into one common table |
| `src/check_rfq_codes.py` | Part-code lookup with review flags |
| `src/match_missing_codes.py` | Spec-based candidates for lines without a code (air circuit breakers only) |
| `src/generate_quotation.py` | Fills the template; adds rows beyond 17 lines |
| `tests/test_kestrel_rfq.py` | Golden test for the Kestrel RFQ (26 checks) |

## Known limits

- RFQs must contain a table. Scanned PDFs (need OCR) and RFQs written as
  email text (need AI extraction) are not supported yet.
- Candidate matching only handles air circuit breakers, and its scoring
  weights have only been tuned on two RFQ lines.
- "MATCHED - VERIFY SPECS" means the part code exists. It does not yet compare
  the catalogue specs with the RFQ description.
- `breaking_capacity` is a class letter for ABB (for example N) and kA for
  Schneider. Do not compare them directly.
- Prices are INR list price (MRP), before any discount. The template is INR with
  18% GST, but the Kestrel RFQ asks for GBP: the draft banner says so.
- The tests check exact prices from the current price lists. If the price lists
  are updated, expect those tests to need new expected values.

## Questions still open for the senior

1. How is the selling price calculated from the list price (discount, margin,
   freight, packing)?
2. Currency, exchange-rate source and tax treatment for UK customers.
3. Numbering rule for the quote reference (currently `DRAFT-01`).
4. What does ABB's "MDO" mean in the Formula Air descriptions?
5. A finished quote for the Kestrel RFQ, to use as the expected answer.
