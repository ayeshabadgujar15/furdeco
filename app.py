import io
import os
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment

from extractor import extract_from_pdf

# --- Palette ---
DARK_GREEN = "#0A3323"
MOSS_GREEN = "#839958"
BEIGE = "#F7F4D5"
ROSY_BROWN = "#D3968C"
MIDNIGHT_GREEN = "#105666"

st.set_page_config(page_title="Furdeco Invoice Extractor", page_icon="📄", layout="wide")

OUTPUT_COLUMNS = ["DATE", "RECIPIENT", "ORDER #", "POSTCODE", "DETAILS", "AMOUNT", "Price", "Invoice No"]

st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: {BEIGE};
    }}
    [data-testid="stHeader"] {{
        background-color: {BEIGE};
    }}
    .furdeco-banner {{
        background-color: {DARK_GREEN};
        color: {BEIGE};
        padding: 1.4rem 1.8rem;
        border-radius: 10px;
        margin-bottom: 1.4rem;
    }}
    .furdeco-banner h1 {{
        color: {BEIGE};
        margin: 0;
        font-size: 1.7rem;
    }}
    .furdeco-banner p {{
        color: {MOSS_GREEN};
        margin: 0.3rem 0 0 0;
        font-size: 0.95rem;
    }}
    section[data-testid="stFileUploaderDropzone"] {{
        background-color: white;
        border: 1.5px dashed {MOSS_GREEN};
        border-radius: 8px;
    }}
    .stButton > button, .stDownloadButton > button {{
        background-color: {MIDNIGHT_GREEN};
        color: {BEIGE};
        border: none;
        border-radius: 6px;
        font-weight: 600;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover {{
        background-color: {DARK_GREEN};
        color: {BEIGE};
    }}
    div[data-testid="stExpander"] {{
        background-color: white;
        border: 1px solid {MOSS_GREEN};
        border-radius: 8px;
    }}
    div[data-testid="stMetric"] {{
        background-color: white;
        border-radius: 8px;
        padding: 0.6rem 0.8rem;
        border: 1px solid {MOSS_GREEN};
    }}
    div[data-baseweb="notification"] {{
        border-radius: 8px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="furdeco-banner">
        <h1>Furdeco Invoice Extractor</h1>
        <p>Upload one or more Furdeco invoice PDFs and get a single Excel file with every line item.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Upload invoice PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_rows = []
    errors = []

    with st.spinner("Extracting..."):
        for f in uploaded_files:
            try:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp.write(f.getvalue())
                    tmp_path = tmp.name
                try:
                    invoice_no, rows = extract_from_pdf(tmp_path)
                finally:
                    os.unlink(tmp_path)

                if not rows:
                    errors.append(f"{f.name}: no line items found — check the file is a Furdeco invoice.")
                    continue

                for r in rows:
                    all_rows.append(
                        {
                            "DATE": r["DATE"],
                            "RECIPIENT": r["RECIPIENT"],
                            "ORDER #": r["ORDER #"],
                            "POSTCODE": r["POSTCODE"],
                            "DETAILS": r["DETAILS"],
                            "AMOUNT": r["AMOUNT"],
                            "Price": "",
                            "Invoice No": invoice_no,
                            "Source File": f.name,
                        }
                    )
            except Exception as e:
                errors.append(f"{f.name}: {e}")

    if errors:
        for e in errors:
            st.warning(e)

    if all_rows:
        df = pd.DataFrame(all_rows)

        total_amount = sum(
            float(str(v).replace("£", "").replace(",", "") or 0) for v in df["AMOUNT"]
        )
        stat_cols = st.columns(3)
        stat_cols[0].metric("Line items", len(df))
        stat_cols[1].metric("Invoices", df["Invoice No"].nunique())
        stat_cols[2].metric("Total amount", f"£{total_amount:,.2f}")

        st.success(f"Extracted {len(df)} line item(s) from {len(uploaded_files) - len(errors)} PDF(s).")

        # --- On-screen filters (narrow down the preview + the exported Excel) ---
        filterable_cols = ["RECIPIENT", "POSTCODE", "DETAILS", "Invoice No", "Source File"]
        with st.expander("Filter", expanded=False):
            filter_cols = st.columns(len(filterable_cols))
            selections = {}
            for col_widget, col_name in zip(filter_cols, filterable_cols):
                options = sorted(df[col_name].dropna().unique().tolist())
                selections[col_name] = col_widget.multiselect(col_name, options)

        filtered_df = df
        for col_name, chosen in selections.items():
            if chosen:
                filtered_df = filtered_df[filtered_df[col_name].isin(chosen)]

        st.dataframe(filtered_df[OUTPUT_COLUMNS + ["Source File"]], use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(filtered_df)} of {len(df)} row(s).")

        # Build the Excel file (only the requested columns, in the requested order),
        # with a filter dropdown on every column header so it can be filtered in Excel too.
        export_df = filtered_df[OUTPUT_COLUMNS]
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Line Items")
            ws = writer.sheets["Line Items"]
            # Generous widths, and text/number formats and no-wrap alignment set
            # explicitly on every cell — otherwise some viewers (older Excel,
            # LibreOffice with certain locale/default settings) can auto-wrap
            # long values in a narrow column and make rows look "broken" across
            # two visual lines even though the underlying data is a single value.
            widths = [12, 26, 26, 12, 34, 10, 10, 12]
            for col_idx, width in enumerate(widths, start=1):
                ws.column_dimensions[chr(64 + col_idx)].width = width

            last_col_letter = chr(64 + len(OUTPUT_COLUMNS))
            last_row = len(export_df) + 1
            no_wrap = Alignment(wrap_text=False, vertical="center")
            for row in ws.iter_rows(min_row=1, max_row=last_row, max_col=len(OUTPUT_COLUMNS)):
                for cell in row:
                    cell.alignment = no_wrap

            ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"
            ws.freeze_panes = "A2"
        buffer.seek(0)

        default_name = f"furdeco_extract_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "Download Excel",
            data=buffer,
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif not errors:
        st.info("No line items found in the uploaded file(s).")
else:
    st.info("Upload one or more PDFs to get started.")
