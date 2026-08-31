"""
Core extraction logic for Furdeco invoice PDFs.

Parses the line-item table (DATE, RECIPIENT, ORDER #, POSTCODE, DETAILS, AMOUNT)
directly from each word's true PDF coordinates (via pdfplumber), rather than
from reflowed/re-justified text. An earlier version of this parser used
`pdftotext -layout` and sliced each line by fixed character offsets computed
from the table header. That works until a row's content is wide enough that
pdftotext re-justifies the line to avoid overlapping columns — at which point
the fixed character offsets no longer line up with that row's actual fields,
and text quietly slides into the wrong column. Reading each word's real (x, y)
position straight from the PDF sidesteps that failure mode entirely: a word's
horizontal position never changes just because a neighboring field is long.
"""

import re

import pdfplumber

INVOICE_NO_RE = re.compile(r"INVOICE\s*#:?\s*(\d+)", re.IGNORECASE)
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
AMOUNT_RE = re.compile(r"^£[\d,]+\.\d{2}$")
SUMMARY_ROW_RE = re.compile(r"^\d+\s+Charges\s+@", re.IGNORECASE)

COLUMNS = ["DATE", "RECIPIENT", "ORDER #", "POSTCODE", "DETAILS", "AMOUNT"]
HEADER_LABELS = {"DATE", "RECIPIENT", "ORDER", "#", "POSTCODE", "DETAILS", "AMOUNT"}

ROW_TOLERANCE = 2.5  # points; words within this many points of top are "the same line"


def clean_mojibake(value: str) -> str:
    """Repair text that was originally valid UTF-8 but got decoded as Latin-1
    somewhere upstream (the classic '£' -> 'Â£' mangling). Generic round-trip
    repair rather than a hardcoded swap, so it self-corrects regardless of
    which tool/build produced the mis-decoding, and is a no-op on text that
    was never mis-decoded in the first place (the round-trip simply fails and
    the original is returned unchanged)."""
    if not value:
        return value
    try:
        return value.encode("latin-1").decode("utf-8")
    except (UnicodeDecodeError, UnicodeEncodeError):
        return value


def smart_join(parts):
    """Join wrapped text fragments, avoiding a stray space after a trailing hyphen
    that is itself a mid-word/line-break hyphen (no space before it), while still
    inserting a space after a hyphen that was originally used as a standalone
    separator (i.e. preceded by a space, like "CFS321432 -")."""
    parts = [p for p in parts if p and p.strip()]
    if not parts:
        return ""
    out = parts[0].strip()
    for p in parts[1:]:
        p = p.strip()
        if out.endswith("-") and not out.endswith(" -"):
            out += p
        else:
            out += " " + p
    return out.strip()


def _group_into_lines(words):
    """Cluster words (each a dict with 'top', 'x0', 'text', ...) into visual
    lines using their vertical position, tolerant of tiny sub-pixel jitter."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        placed = False
        for line in lines:
            if abs(line["top"] - w["top"]) <= ROW_TOLERANCE:
                line["words"].append(w)
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})
    lines.sort(key=lambda l: l["top"])
    for line in lines:
        line["words"].sort(key=lambda w: w["x0"])
    return lines


def _find_header_columns(line_words):
    """If this line is the table header, return {column_name: x0_start}."""
    texts = {w["text"].upper() for w in line_words}
    if not HEADER_LABELS.issubset(texts):
        return None

    by_text = {}
    for w in line_words:
        by_text.setdefault(w["text"].upper(), w["x0"])

    order_x0 = min(by_text.get("ORDER", by_text.get("#", 0)), by_text.get("#", by_text.get("ORDER", 0)))
    return {
        "DATE": by_text["DATE"],
        "RECIPIENT": by_text["RECIPIENT"],
        "ORDER #": order_x0,
        "POSTCODE": by_text["POSTCODE"],
        "DETAILS": by_text["DETAILS"],
        "AMOUNT": by_text["AMOUNT"],
    }


def _column_for_word(word, col_bounds):
    """Classify a word into one of COLUMNS. AMOUNT is identified by pattern
    (it's right-aligned, so its x0 shifts with digit count) rather than by
    a fixed boundary; everything else is bucketed by x-position against the
    boundaries taken from the header row."""
    text = word["text"]
    if AMOUNT_RE.match(text):
        return "AMOUNT"

    # A small tolerance: a word that visually starts exactly at a column's
    # boundary can, after PDF glyph-metric rounding, land a hair to the left
    # of the header word's x0 (e.g. 122.279 vs a header boundary of 122.280).
    # Without slack that pushes it into the previous column entirely.
    EPSILON = 3.0
    x0 = word["x0"]
    best_col = COLUMNS[0]
    best_start = -1
    for col in COLUMNS:
        start = col_bounds[col] - EPSILON
        if start <= x0 and start > best_start:
            best_col = col
            best_start = start
    return best_col


def extract_invoice_number(full_text: str) -> str:
    m = INVOICE_NO_RE.search(full_text)
    return m.group(1) if m else ""


def extract_from_pdf(pdf_path: str):
    """Return (invoice_number, list_of_row_dicts) for one PDF file."""
    all_rows = []
    full_text_parts = []

    # This state deliberately lives OUTSIDE the page loop, not inside it — a
    # multi-page invoice only prints the table header once (on the first
    # page); later pages just continue the same table with no header line of
    # their own. If col_bounds/table_active were reset to "no table seen yet"
    # at the top of every page (as an earlier version of this function did),
    # every row after page 1 would silently fail the "no header seen yet"
    # check below and get dropped — the table would only ever be as long as
    # whatever fit on the first page. Carrying col_bounds forward lets a page
    # with no header of its own still be read using the columns from the
    # last header we saw; encountering an actual header line (on any page)
    # still refreshes col_bounds, which also handles invoices that *do*
    # reprint the header on every page.
    col_bounds = None
    pending = []  # buffered (col, x0, text) fragments for the next row
    current = None  # {col: [(x0, text), ...]}
    main_line_had_content = {}
    awaiting_suffix = False
    # table_finished is permanent for the rest of the document once the
    # invoice's own rollup/summary row is seen — that row only ever appears
    # once, right after the true last line item, so nothing after it (on
    # this page or any later one) is ever a line item.
    table_finished = False

    def flush(row):
        """Returns True if the table is done after this row (i.e. it was
        the invoice-level rollup row, which is always the only row in
        its table)."""
        if row is None:
            return False
        joined = {c: smart_join([t for _, _, t in sorted(row[c])]) for c in COLUMNS}
        if SUMMARY_ROW_RE.match(joined["DETAILS"]):
            return True
        if not any(joined.values()):
            return False
        all_rows.append(joined)
        return False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text_parts.append(page.extract_text() or "")
            words = page.extract_words(x_tolerance=1.5, keep_blank_chars=False)
            if not words:
                continue
            lines = _group_into_lines(words)

            # Resume capturing at the top of every page as long as we already
            # know the columns and haven't hit the real end of the table yet
            # — this is what lets a continuation page with no header of its
            # own still get read. table_active can still be paused for the
            # rest of THIS page below (by the prose/footer heuristic) without
            # that pause carrying over to the next page.
            table_active = col_bounds is not None and not table_finished

            for line in lines:
                header = _find_header_columns(line["words"])
                if header is not None:
                    flush(current)
                    current = None
                    pending = []
                    col_bounds = header
                    awaiting_suffix = False
                    table_active = True
                    # A fresh header line is unambiguous proof this is a new,
                    # active table section — reset table_finished too, not
                    # just table_active. Some invoices put a one-row summary
                    # ("N Charges @ £X") on its own cover page, in the same
                    # header+row layout as the real per-order table that then
                    # starts fresh on the next page. Without this reset, that
                    # cover-page summary would permanently mark the document
                    # "finished", and a later page that continues the *real*
                    # table without repeating the header would never re-arm
                    # (nothing else resets table_finished), silently dropping
                    # every row on it.
                    table_finished = False
                    continue

                if col_bounds is None or not table_active:
                    continue  # no table header seen yet, or this table has ended

                by_col = {c: [] for c in COLUMNS}
                for w in line["words"]:
                    col = _column_for_word(w, col_bounds)
                    by_col[col].append((line["top"], w["x0"], w["text"]))

                has_date = any(DATE_RE.match(t) for _, _, t in by_col["DATE"])
                has_amount = bool(by_col["AMOUNT"])
                is_main_line = has_date or has_amount
                columns_touched = sum(1 for c in COLUMNS if by_col[c])

                if is_main_line:
                    if flush(current):
                        table_finished = True
                        table_active = False
                        current = None
                        continue
                    current = {c: [] for c in COLUMNS}
                    for c, vals in pending:
                        current[c].append(vals)
                    pending = []
                    main_line_had_content = {c: bool(by_col[c]) for c in COLUMNS}
                    for c in COLUMNS:
                        current[c].extend(by_col[c])
                    awaiting_suffix = True
                elif columns_touched > 2:
                    # A non-main-line touching several columns at once isn't a
                    # genuine wrapped fragment (those only ever spill into one
                    # column) — it's prose below the table (bank details, T&Cs,
                    # etc.). Only pause capturing for the rest of *this* page —
                    # boilerplate reprinted at the bottom of every page (e.g. a
                    # "continued" notice or company details) shouldn't stop a
                    # later page's genuine rows from being read; only the real
                    # summary row (handled above) permanently ends the table.
                    if flush(current):
                        table_finished = True
                    current = None
                    table_active = False
                else:
                    if awaiting_suffix and current is not None:
                        for c in COLUMNS:
                            if not by_col[c]:
                                continue
                            if not main_line_had_content.get(c, False):
                                current[c].extend(by_col[c])
                            else:
                                for v in by_col[c]:
                                    pending.append((c, v))
                        flush(current)
                        current = None
                        awaiting_suffix = False
                    else:
                        for c in COLUMNS:
                            for v in by_col[c]:
                                pending.append((c, v))

            flush(current)
            current = None

    invoice_no = extract_invoice_number("\n".join(full_text_parts))
    all_rows = [{k: clean_mojibake(v) for k, v in row.items()} for row in all_rows]
    return invoice_no, all_rows
