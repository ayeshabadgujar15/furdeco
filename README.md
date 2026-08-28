# Furdeco Invoice Extractor

A simple Streamlit app: upload one or more Furdeco invoice PDFs, get back a
single Excel file with every line item.

## Output columns
DATE, RECIPIENT, ORDER #, POSTCODE, DETAILS, AMOUNT, Price, Invoice No

- The first six columns are read straight off each invoice's line-item table.
- **Price** is left blank — that's a later step, not part of this extractor.
- **Invoice No** is read from each PDF's own "INVOICE #:" line and applied to
  every row that came from that file (so if you upload several invoices at
  once, each keeps its own number).
- The invoice's own rollup/summary line (e.g. "37 Charges @ £985.31") is
  automatically excluded — only the individual order line items are kept.

The app also includes a "Source File" column in the on-screen preview (not
in the Excel export) so you can see which PDF each row came from when you
upload several at once.

## Running it

Requires Python 3 with the packages in `requirements.txt`, plus the
`pdftotext` command-line tool (from the `poppler-utils` package), which is
used to pull text out of the PDF while preserving its column layout.

```
pip install -r requirements.txt
# Debian/Ubuntu: sudo apt-get install poppler-utils
# macOS: brew install poppler

streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## Files

- `app.py` — the Streamlit UI (upload → extract → preview → download Excel).
- `extractor.py` — the PDF parsing logic, usable on its own if needed.
- `requirements.txt` / `packages.txt` — Python and system dependencies
  (the `packages.txt` file also lets this deploy directly on Streamlit
  Community Cloud, which reads it automatically to install poppler-utils).

## Notes / next steps

This is just the extraction step, as requested — the UI styling and any
further steps (like filling in Price) can be layered on next.
