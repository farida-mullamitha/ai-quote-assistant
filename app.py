from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from src.extract_rfq import extract_rfq_lines
from src.load_price_lists import load_catalogue, add_supplier_prices
from src.match_missing_codes import find_missing_code_candidates
from src.generate_quotation import generate_quotation


PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


st.set_page_config(
    page_title="IEPCI Quote Assistant",
    page_icon="📄",
    layout="wide",
)

st.title("IEPCI Quote Assistant")
st.caption("RFQ extraction • Supplier matching • Missing-code review • Quotations")

st.warning(
    "Draft only. Verify exact product specifications, currency, tax, "
    "and commercial terms before sending a quotation."
)

uploaded_file = st.file_uploader(
    "Upload an RFQ document",
    type=["pdf", "docx", "xlsx"],
)

if uploaded_file is None:
    st.info("Upload an RFQ to begin.")
    st.stop()


# Save the uploaded file temporarily so existing modules can process it.
with tempfile.TemporaryDirectory() as temp_dir:
    rfq_path = Path(temp_dir) / f"uploaded_rfq{Path(uploaded_file.name).suffix}"

    rfq_path.write_bytes(uploaded_file.getvalue())

    try:
        with st.spinner("Extracting RFQ items..."):
            rfq_df = extract_rfq_lines(rfq_path)

        if rfq_df is None or rfq_df.empty:
            st.error("No RFQ items were extracted.")
            st.stop()

        with st.spinner("Loading supplier prices and matching part codes..."):
            catalogue = load_catalogue()
            priced_df = add_supplier_prices(rfq_df, catalogue)

        # Run missing and incomplete code matching.
        with st.spinner("Searching for possible missing-code matches..."):
            candidate_records = find_missing_code_candidates(
                priced_df,
                catalogue,
            )

        candidate_df = pd.DataFrame(candidate_records)

        # Navigation for the workflow.
        page = st.sidebar.radio(
            "Navigation",
            [
                "Dashboard",
                "RFQ & Price Matching",
                "Missing-Code Review",
                "Generate Quotation",
                "Price Catalogue",
            ],
        )

        st.sidebar.divider()
        st.sidebar.write(f"**Uploaded file:** {uploaded_file.name}")
        st.sidebar.write(f"**RFQ items:** {len(priced_df)}")

        # ---------------- Dashboard ----------------
        if page == "Dashboard":
            st.header("Dashboard")
            st.write("Overview of the uploaded RFQ and matching results.")

            col1, col2, col3 = st.columns(3)

            col1.metric("RFQ Items", len(priced_df))

            if "supplier_price_status" in priced_df.columns:
                statuses = priced_df["supplier_price_status"].fillna("").astype(str)
                matched_count = statuses.str.contains(
                    "LIST PRICE FOUND|DUPLICATE CODE - same price",
                    case=False,
                    regex=True,
                ).sum()
            else:
                matched_count = 0

            col2.metric("Items with List Prices", int(matched_count))
            col3.metric("Missing-Code Suggestions", len(candidate_df))

            st.subheader("Workflow")
            st.write("1. Upload and extract the RFQ.")
            st.write("2. Review supplier price matches.")
            st.write("3. Investigate missing or incomplete part codes.")
            st.write("4. Generate a draft quotation.")

            st.info(
                "A candidate suggestion is not an approved product match. "
                "Review the product code and specifications before use."
            )

        # ---------------- RFQ and Price Matching ----------------
        elif page == "RFQ & Price Matching":
            st.header("RFQ & Price Matching")
            st.write("Review the extracted items and supplier matching status.")

            st.dataframe(priced_df, use_container_width=True)

            if "supplier_price_status" in priced_df.columns:
                st.subheader("Matching Summary")

                summary = (
                    priced_df["supplier_price_status"]
                    .fillna("STATUS NOT AVAILABLE")
                    .value_counts()
                    .rename_axis("Status")
                    .reset_index(name="Items")
                )

                st.dataframe(summary, use_container_width=True)

            csv_data = priced_df.to_csv(index=False).encode("utf-8")

            st.download_button(
                "Download price-matching report",
                data=csv_data,
                file_name="rfq_price_matching.csv",
                mime="text/csv",
            )

        # ---------------- Missing-Code Review ----------------
        elif page == "Missing-Code Review":
            st.header("Missing-Code Review")

            st.write(
                "Review suggestions for missing or incomplete part codes. "
                "Suggestions never automatically replace the original RFQ code."
            )

            if candidate_df.empty:
                st.success(
                    "No missing-code candidate records were generated "
                    "for this RFQ."
                )
            else:
                review_lines = candidate_df["line_no"].nunique()
                st.metric("RFQ Lines Requiring Review", review_lines)

                if "decision" in candidate_df.columns:
                    st.warning(
                        "All candidates require human verification before "
                        "a product or price can be approved."
                    )

                st.dataframe(candidate_df, use_container_width=True)

                candidates_csv = candidate_df.to_csv(
                    index=False
                ).encode("utf-8")

                st.download_button(
                    "Download missing-code candidates",
                    data=candidates_csv,
                    file_name="missing_code_candidates.csv",
                    mime="text/csv",
                )

                st.caption(
                    "A high score indicates similarity based on the current "
                    "rules; it does not prove that the product is equivalent."
                )

        # ---------------- Generate Quotation ----------------
        elif page == "Generate Quotation":
            st.header("Generate Quotation")

            st.write(
                f"Ready to generate a draft using {len(priced_df)} extracted items."
            )

            st.warning(
                "Supplier prices may be in INR while the RFQ requests GBP. "
                "Confirm currency conversion, applicable tax, and commercial "
                "terms before issuing a customer quotation."
            )

            if st.button("Generate Draft Quotation", type="primary"):
                with tempfile.TemporaryDirectory() as output_dir:
                    output_path = (
                        Path(output_dir) / "IEPCI_Quotation_DRAFT.xlsx"
                    )

                    try:
                        with st.spinner("Generating Excel quotation..."):
                            generate_quotation(
                                rfq_path=rfq_path,
                                output_path=output_path,
                            )

                        if output_path.exists():
                            quotation_bytes = output_path.read_bytes()

                            st.success("Draft quotation generated successfully.")

                            st.download_button(
                                "Download Draft Quotation",
                                data=quotation_bytes,
                                file_name="IEPCI_Quotation_DRAFT.xlsx",
                                mime=(
                                    "application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet"
                                ),
                            )
                        else:
                            st.error("The quotation file was not created.")

                    except Exception as error:
                        st.error(f"Quotation generation failed: {error}")

        # ---------------- Price Catalogue ----------------
        elif page == "Price Catalogue":
            st.header("Supplier Price Catalogue")

            st.write(
                "Browse the supplier records loaded by the existing "
                "price-list module."
            )

            st.dataframe(catalogue, use_container_width=True)

            catalogue_csv = catalogue.to_csv(index=False).encode("utf-8")

            st.download_button(
                "Download Price Catalogue",
                data=catalogue_csv,
                file_name="price_catalogue.csv",
                mime="text/csv",
            )

    except Exception as error:
        st.error(f"Could not process this RFQ: {error}")