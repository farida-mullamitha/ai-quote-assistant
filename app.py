
import streamlit as st
import pandas as pd
from pathlib import Path
from tempfile import TemporaryDirectory

from src.extract_rfq import extract_rfq_lines
from src.match_missing_codes import find_missing_code_candidates
from src.load_price_lists import load_catalogue
from src.generate_quotation import generate_quotation


st.set_page_config(
    page_title="IEPCI Quote Assistant",
    page_icon="📋",
    layout="wide",
)

st.title("IEPCI Quote Assistant")
st.caption("RFQ extraction · Product matching · Pricing · Excel quotations")


@st.cache_data
def get_catalogue():
    return load_catalogue()


def enrich_rfq_with_catalogue(rfq_df, catalogue):
    """Attach exact part-code catalogue results to each RFQ item."""
    from re import sub

    df = rfq_df.copy()

    def normalise_code(value):
        if pd.isna(value):
            return ""
        return sub(r"[^A-Z0-9]", "", str(value).upper())

    catalogue = catalogue.copy()
    catalogue["_lookup_code"] = catalogue["part_code"].map(normalise_code)

    # Keep only non-empty codes that identify one catalogue record.
    catalogue = catalogue[
        catalogue["_lookup_code"].ne("")
    ].drop_duplicates("_lookup_code", keep=False)

    lookup = catalogue.set_index("_lookup_code", drop=False)

    supplier_prices = []
    catalogue_descriptions = []
    match_statuses = []

    for code in df["part_code"]:
        key = normalise_code(code)

        if key and key in lookup.index:
            match = lookup.loc[key]

            supplier_prices.append(match["price_inr"])
            catalogue_descriptions.append(match["description"])
            match_statuses.append(
                "Exact code — price available"
                if pd.notna(match["price_inr"])
                else f"Exact code — {match['price_status']}"
            )
        else:
            supplier_prices.append(None)
            catalogue_descriptions.append("")
            match_statuses.append("No unique exact-code match")

    df["catalogue_match_status"] = match_statuses
    df["catalogue_description"] = catalogue_descriptions
    df["supplier_list_price"] = supplier_prices

    if "unit_rate" not in df.columns:
        df["unit_rate"] = None

    if "price_approved" not in df.columns:
        df["price_approved"] = False

    return df


def get_rfq():
    return st.session_state.get("rfq_lines")


def prepare_rfq(df):
    df = df.copy()

    if "unit_rate" not in df.columns:
        df["unit_rate"] = None

    if "price_approved" not in df.columns:
        df["price_approved"] = False

    return df


if "rfq_lines" not in st.session_state:
    st.session_state["rfq_lines"] = None

if "match_candidates" not in st.session_state:
    st.session_state["match_candidates"] = None


pages = [
    "Dashboard",
    "RFQ Analysis",
    "Product Matching",
    "Supplier Catalogue",
    "Pricing & Quotation",
]

page = st.sidebar.radio("Navigation", pages)

rfq = get_rfq()


# ---------------- DASHBOARD ----------------

if page == "Dashboard":
    st.subheader("Dashboard")

    total_items = len(rfq) if rfq is not None else 0

    missing_codes = 0
    unapproved_prices = 0

    if rfq is not None:
        if "part_code" in rfq.columns:
            missing_codes = (
                rfq["part_code"]
                .fillna("")
                .astype(str)
                .str.strip()
                .eq("")
                .sum()
            )

        if "price_approved" in rfq.columns:
            unapproved_prices = (
                ~rfq["price_approved"].fillna(False).astype(bool)
            ).sum()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("RFQ items", total_items)
    col2.metric("Missing part codes", int(missing_codes))
    col3.metric("Prices awaiting approval", int(unapproved_prices))
    col4.metric(
        "Candidate records",
        len(st.session_state["match_candidates"] or []),
    )

    st.divider()

    if rfq is None:
        st.info("Start by opening RFQ Analysis and uploading a customer RFQ.")
    else:
        st.success(
            f"Loaded RFQ: {st.session_state.get('rfq_filename', 'RFQ')}"
        )
        st.dataframe(rfq, use_container_width=True, hide_index=True)


# ---------------- RFQ ANALYSIS ----------------

elif page == "RFQ Analysis":
    st.subheader("Upload and review RFQ")

    uploaded_file = st.file_uploader(
        "Choose a customer RFQ",
        type=["pdf", "xlsx", "xls", "docx"],
    )

    if uploaded_file and st.button("Extract RFQ items", type="primary"):
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir) / uploaded_file.name
            temp_path.write_bytes(uploaded_file.getvalue())

            try:
                extracted = extract_rfq_lines(str(temp_path))
                extracted = pd.DataFrame(extracted)

                if extracted.empty:
                    st.warning("No line items were extracted.")
                else:
                    st.session_state["rfq_lines"] = prepare_rfq(extracted)
                    st.session_state["rfq_filename"] = uploaded_file.name
                    st.session_state["match_candidates"] = None
                    st.session_state.pop("rfq_editor", None)
                    st.session_state.pop("pricing_editor", None)
                    st.session_state.pop("quotation_bytes", None)

                    st.success(f"Extracted {len(extracted)} items.")

            except Exception as error:
                st.error(f"Extraction failed: {error}")

    rfq = get_rfq()

    if rfq is not None:
        st.markdown("### Extracted items")
        st.caption("Review the extracted information and correct mistakes.")

        edited = st.data_editor(
            rfq,
            key="rfq_editor",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
        )

        if st.button("Save RFQ corrections"):
            st.session_state["rfq_lines"] = prepare_rfq(edited)
            st.session_state.pop("match_candidates", None)
            st.session_state.pop("quotation_bytes", None)
            st.success("RFQ corrections saved.")


# ---------------- PRODUCT MATCHING ----------------

elif page == "Product Matching":
    st.subheader("Product and part-code matching")

    rfq = get_rfq()

    if rfq is None:
        st.info("Upload and extract an RFQ first.")
    else:
        st.write("Items requiring part-code verification")

        if "part_code" in rfq.columns:
            missing = rfq[
                rfq["part_code"].fillna("").astype(str).str.strip().eq("")
            ]

            st.metric("Items with missing codes", len(missing))
            if not missing.empty:
                st.dataframe(missing, use_container_width=True, hide_index=True)

        if st.button("Find candidate matches", type="primary"):
            try:
                with st.spinner("Searching supplier catalogues..."):
                    candidates = find_missing_code_candidates(rfq)

                st.session_state["match_candidates"] = candidates
                st.success(f"Found {len(candidates)} candidate records.")

            except Exception as error:
                st.error(f"Matching failed: {error}")

        candidates = st.session_state.get("match_candidates")

        if candidates is not None:
            results = pd.DataFrame(candidates)

            if results.empty:
                st.warning("No candidates found. Supplier confirmation may be needed.")
            else:
                for line_no, group in results.groupby("line_no", sort=False):
                    st.markdown(f"#### RFQ line {line_no}")

                    if "rfq_description" in group.columns:
                        st.write(group.iloc[0]["rfq_description"])

                    columns = [
                        "rank",
                        "score",
                        "part_code",
                        "candidate_description",
                        "price",
                        "supporting_details",
                        "checks_needed",
                        "decision",
                    ]

                    st.dataframe(
                        group[[c for c in columns if c in group.columns]],
                        use_container_width=True,
                        hide_index=True,
                    )

                st.warning(
                    "Candidate matches are suggestions, not confirmed "
                    "equivalent products. Verify specifications with the supplier."
                )


# ---------------- SUPPLIER CATALOGUE ----------------

elif page == "Supplier Catalogue":
    st.subheader("Supplier Catalogue Explorer")

    try:
        catalogue = get_catalogue()

        st.metric("Catalogue records", len(catalogue))

        if "supplier" in catalogue.columns:
            suppliers = sorted(
                catalogue["supplier"].dropna().astype(str).unique()
            )
            selected_supplier = st.selectbox(
                "Supplier",
                ["All suppliers"] + suppliers,
            )

            if selected_supplier != "All suppliers":
                catalogue = catalogue[
                    catalogue["supplier"].astype(str) == selected_supplier
                ]

        search = st.text_input(
            "Search by part code or description",
            placeholder="e.g. AF26-30-00",
        )

        if search.strip():
            searchable = catalogue.fillna("").astype(str)
            mask = searchable.apply(
                lambda col: col.str.contains(
                    search.strip(), case=False, regex=False
                )
            ).any(axis=1)
            catalogue = catalogue[mask]

        display_columns = [
            "supplier",
            "part_code",
            "description",
            "price_inr",
            "price_status",
            "rating_a",
            "poles",
            "breaking_capacity",
        ]

        display_columns = [
            c for c in display_columns if c in catalogue.columns
        ]

        st.dataframe(
            catalogue[display_columns].head(500),
            use_container_width=True,
            hide_index=True,
        )

        if len(catalogue) > 500:
            st.caption("Showing the first 500 matching records.")

    except Exception as error:
        st.error(f"Could not load the supplier catalogue: {error}")


# ---------------- PRICING & QUOTATION ----------------
elif page == "Pricing & Quotation":
    st.subheader("Pricing Review and Excel Quotation")

    rfq = get_rfq()

    if rfq is None:
        st.info("Upload and extract an RFQ first.")
    else:
        try:
            catalogue = get_catalogue()
            pricing_df = enrich_rfq_with_catalogue(rfq, catalogue)

            st.caption(
                "Supplier list prices are reference prices, not approved "
                "IEPCI selling rates."
            )

            matched = pricing_df[
                pricing_df["catalogue_match_status"].str.startswith(
                    "Exact code — price available"
                )
            ]

            col1, col2, col3 = st.columns(3)
            col1.metric("RFQ items", len(pricing_df))
            col2.metric("Exact codes with prices", len(matched))
            col3.metric(
                "Items needing review",
                len(pricing_df) - len(matched),
            )

            st.markdown("### Review each item's price")

            editable_columns = [
                "line_no",
                "make",
                "part_code",
                "description",
                "qty",
                "catalogue_match_status",
                "catalogue_description",
                "supplier_list_price",
                "unit_rate",
                "price_approved",
            ]

            edited = st.data_editor(
                pricing_df[editable_columns],
                key="pricing_review_editor",
                use_container_width=True,
                hide_index=True,
                disabled=[
                    "line_no",
                    "make",
                    "part_code",
                    "description",
                    "qty",
                    "catalogue_match_status",
                    "catalogue_description",
                    "supplier_list_price",
                ],
                column_config={
                    "supplier_list_price": st.column_config.NumberColumn(
                        "Supplier list price",
                        format="₹%.2f",
                    ),
                    "unit_rate": st.column_config.NumberColumn(
                        "IEPCI selling rate",
                        min_value=0.0,
                        format="₹%.2f",
                    ),
                    "price_approved": st.column_config.CheckboxColumn(
                        "Rate approved",
                    ),
                },
            )

            st.warning(
                "Enter the approved IEPCI selling rate separately. "
                "Only tick Rate approved after the product, price and "
                "commercial rules have been verified."
            )

            if st.button("Save pricing review", type="primary"):
                updated = pricing_df.copy()

                updated["unit_rate"] = edited["unit_rate"].tolist()
                updated["price_approved"] = (
                    edited["price_approved"].fillna(False).tolist()
                )

                st.session_state["rfq_lines"] = updated
                st.session_state.pop("quotation_bytes", None)

                st.success("Pricing review saved.")

        except Exception as error:
            st.error(f"Pricing review failed: {error}")

