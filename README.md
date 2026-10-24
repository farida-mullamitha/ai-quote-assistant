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