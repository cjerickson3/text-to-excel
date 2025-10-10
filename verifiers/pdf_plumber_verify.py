"""
Lightweight PDF verification for Chase statements using pdfplumber.

Usage from your main script:
    from verifiers.pdf_plumber_verify import verify_statement_pdf, parse_pdf_totals

    report = verify_statement_pdf(args.pdf, df, begin_bal, computed_end, debug=args.debug)
    # report = {"summary": {...}, "issues": [...], "by_page": {...}}

This module does not write images or files. It just returns a structured report.
"""

from __future__ import annotations

import re
from typing import Dict, Optional, Any, List

# Optional import: your program should continue even if pdfplumber isn't installed
try:
    import pdfplumber  # type: ignore
except Exception:  # pragma: no cover
    pdfplumber = None

# ---------- Regex helpers ----------
# Tolerant “row-level” predicates
TOP_PATTS = [
    re.compile(r'\bCHECK(ING)?\b.*\bSUMMARY\b', re.I),           # CHECKING SUMMARY
    re.compile(r'\bCONSOLIDATED\b.*\bSUMMARY\b', re.I),          # CONSOLIDATED BALANCE SUMMARY
    re.compile(r'\bDATE\b.*\bAMOUNT\b', re.I),                   # DATE ... AMOUNT (table header)
]
BOTTOM_PATTS = [
    re.compile(r'\bSAVINGS\b.*\bSUMMARY\b', re.I),               # SAVINGS SUMMARY
    re.compile(r'\bCHASE\b.*\bSAVINGS\b', re.I),                 # CHASE SAVINGS
]
# Accept multiple ways to recognize CHECKS
_SEC_PATTERNS = {
    "DEPOSITS": re.compile(r'^\s*DEPOSITS?\b.*\b(ADDITIONS?|CREDITS?)\b', re.I),
    "CHECKS":   re.compile(r'^\s*(CHECKS?\s+PAID|CHECK\s+NO\.\s+DESCRIPTION)\b', re.I),
    "ATM":      re.compile(r'^\s*ATM\s*&?\s*DEBIT\s*CARD\s*WITHDRAWALS?\b', re.I),
    "ELECTRONIC": re.compile(r'^\s*ELECTRONIC\s+WITHDRAWALS?\b', re.I),
}
_CHECK_ROW = re.compile(r'^\s*\d{4,6}\s+(?:[\^\*]\s+)?\d{1,2}[/-]\d{1,2}\b')  # a check row

AMOUNT_RE = re.compile(r"\$?\s*\d{1,3}(?:,\d{3})*\.\d{2}")

BEGIN_BAL_RE = re.compile(
    r"Beginning\s+Balance\s+(\$?\s*\d{1,3}(?:,\d{3})*\.\d{2})",
    re.I,
)
END_BAL_RE = re.compile(
    r"(Ending|Closing)\s+Balance\s+(\$?\s*\d{1,3}(?:,\d{3})*\.\d{2})",
    re.I,
)
TOT_DEP_RE = re.compile(
    r"Total\s+Deposits\s+and\s+Additions\s+(\$?\s*\d{1,3}(?:,\d{3})*\.\d{2})",
    re.I,
)
# Allow optional "ATM & Debit Card " prefix and optional "and Debits"
TOT_WD_RE = re.compile(
    r"Total\s+(?:ATM\s*&\s*Debit\s*Card\s+)?Withdrawals(?:\s+and\s+Debits)?\s+(\$?\s*\d{1,3}(?:,\d{3})*\.\d{2})",
    re.I,
)

PADDING_Y = 2.0  # avoid cropping through baselines

# ---------- Small utilities ----------

def _to_float(txt: str) -> float:
    s = txt.strip().replace("$", "").replace(",", "")
    neg = False
    if s.endswith("-"):
        neg, s = True, s[:-1]
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    v = float(s)
    return -v if neg else v


def _fmt_amt(v: float) -> str:
    return f"{abs(v):,.2f}"


def _mk_date_tokens(date_str: str) -> List[str]:
    """
    Input: 'YYYY-MM-DD'
    Output tokens that commonly appear in the PDF text.
    """
    mm = date_str[5:7]
    dd = date_str[8:10]
    return [f"{mm}/{dd}", f"{int(mm)}/{int(dd)}", f"{mm}-{dd}", f"{int(mm)}-{int(dd)}"]

def _extract_cropped_text(page, bbox, *, x_tol=2, y_tol=2):
    """
    bbox = (x0, y0, x1, y1) in PDF coordinate space (origin top-left).
    Returns '' on any issue; never raises inside verify path.
    """
    x0, y0, x1, y1 = bbox
    # Guard weird coords
    if x0 >= x1 or y0 >= y1:
        return ""
    try:
        band = page.crop(bbox)
        return band.extract_text(x_tolerance=x_tol, y_tolerance=y_tol) or ""
    except Exception:
        return ""
def _extract_cropped_lines(page, bbox, *, x_tol=2, y_tol=1.0):
    """
    Convert a cropped region into stable text rows using pdfplumber's words.
    Falls back to extract_text() if needed.
    """
    x0, y0, x1, y1 = bbox
    # pad vertical bounds slightly
    y0p = max(0, y0 - PADDING_Y)
    y1p = min(page.height, y1 + PADDING_Y)
    if x0 >= x1 or y0p >= y1p:
        return []

    try:
        band = page.crop((x0, y0p, x1, y1p))
        words = band.extract_words(
            x_tolerance=x_tol,
            y_tolerance=y_tol,
            use_text_flow=False,
            keep_blank_chars=False,
        ) or []
        if not words:
            txt = band.extract_text(x_tolerance=x_tol, y_tolerance=y_tol) or ""
            return txt.splitlines()

        # group words into rows by quantized 'top'
        buckets = {}
        q = max(0.1, y_tol)
        for w in words:
            key = int(round(w["top"] / q))
            buckets.setdefault(key, []).append(w)

        lines = []
        for key in sorted(buckets):
            row = sorted(buckets[key], key=lambda w: w["x0"])
            lines.append(" ".join(w["text"] for w in row))
        return lines
    except Exception:
        return []   
_hdr_end_then_date = re.compile(r'^(\*end\*[\s\S]*?)\s+(\d{1,2}/\d{2}\b)', re.IGNORECASE) 
def _split_header_joined_lines(lines):
    """
    If a line starts with '*end*... <date>' because two lines merged,
    split into two lines so parsing sees the date at line start.
    """
    fixed = []
    for ln in lines:
        m = _hdr_end_then_date.match(ln)
        if m:
            fixed.append(m.group(1))      # header end on its own line
            fixed.append(m.group(2) + ln[m.end(2):].rstrip())  # date line
        else:
            fixed.append(ln)
    return fixed
# --- top of file ---
def _page_rows(page, *, x_tol=2, y_tol=1.0):
    """
    Return list of (y_top, row_text) by grouping words that share a similar top.
    """
    words = page.extract_words(
        x_tolerance=x_tol, y_tolerance=y_tol,
        use_text_flow=False, keep_blank_chars=False
    ) or []
    if not words:
        return []

    buckets = {}
    q = max(0.1, y_tol)
    for w in words:
        key = int(round(w["top"] / q))
        buckets.setdefault(key, []).append(w)

    rows = []
    for key in sorted(buckets):
        row = sorted(buckets[key], key=lambda w: w["x0"])
        txt = " ".join(w["text"] for w in row)
        y_top = min(w["top"] for w in row)
        rows.append((y_top, txt))
    return rows

def _find_row_y(rows, patterns, *, min_y=0.0):
    for y, text in rows:
        if y <= min_y:
            continue
        if any(p.search(text) for p in patterns):
            return y
    return None

def _compute_checking_band_y(page):
    """
    Discover [y0,y1] using row-level text:
      y0 = first row that looks like a Checking header or table header
      y1 = first Savings banner row *below y0*, else page bottom
    """
    rows = _page_rows(page, x_tol=2, y_tol=1.0)
    if not rows:
        return (None, None)

    y0 = _find_row_y(rows, TOP_PATTS)
    if y0 is None:
        # NEW: lightweight debug—peek first 8 rows to see what we missed
        # (guard so verify() doesn't spam unless debug=True in caller)
        return (None, None)

    y1 = _find_row_y(rows, BOTTOM_PATTS, min_y=y0)
    if y1 is None:
        y1 = float(page.height)

    return (y0, y1)

def _page_texts(pdf, *, x_tol=2, y_tol=1.0) -> list[str]:
    """
    Return a list of strings, one per PDF page (full page text).
    Prefers word-rows for stability; falls back to extract_text().
    """
    texts: list[str] = []
    for page in pdf.pages:
        rows = _page_rows(page, x_tol=x_tol, y_tol=y_tol)  # uses your existing _page_rows
        if rows:
            texts.append("\n".join(t for _, t in rows))
        else:
            txt = page.extract_text(x_tolerance=x_tol, y_tolerance=y_tol) or ""
            texts.append(txt)
    return texts
def _page_texts_checking_band(pdf, *, debug: bool = False) -> list[str]:
    """
    Return a list of strings, one per page, containing only the 'Checking' band text.
    Empty string for pages with no Checking band.
    """
    page_texts: list[str] = []
    for i, page in enumerate(pdf.pages, start=1):
        y0, y1 = _compute_checking_band_y(page)
        if y0 is None or y1 is None or y1 <= y0:
            if debug:
                print(f"[pdf-band] page {i}: no checking band detected; skipped")
            page_texts.append("")  # keep page indexing
            continue

        lines = _extract_cropped_lines(page, (0, y0, float(page.width), y1), x_tol=2, y_tol=1.0)
        if debug:
            print(f"[pdf-band] page {i}: band {y0:.1f}..{y1:.1f}, lines={len(lines)}")
        lines = _split_header_joined_lines(lines)
        page_texts.append("\n".join(lines))

    if debug:
        print(f"[pdf-band] built texts for {len(page_texts)} pages")
    return page_texts
# --- Section span finder used to keep only transaction blocks (optional) ---
_SEC_PATTERNS = {
    "DEPOSITS": re.compile(r'^\s*DEPOSITS?\b.*\b(ADDITIONS?|CREDITS?)\b', re.I),
    "CHECKS":   re.compile(r'^\s*(CHECKS?\s+PAID|CHECK\s+NO\.\s+DESCRIPTION)\b', re.I),
    "ATM":      re.compile(r'^\s*ATM\s*&?\s*DEBIT\s*CARD\s*WITHDRAWALS?\b', re.I),
    "ELECTRONIC": re.compile(r'^\s*ELECTRONIC\s+WITHDRAWALS?\b', re.I),
}
# also accept a row that *is itself* a check line
_CHECK_ROW = re.compile(r'^\s*\d{4,6}\s+(?:[\^\*]\s+)?\d{1,2}[/-]\d{1,2}\b')

def _find_section_spans(lines: list[str]):
    labels = []
    for i, ln in enumerate(lines):
        s = (ln or "").strip()
        # Label by headers first
        for name, rx in _SEC_PATTERNS.items():
            if rx.search(s):
                labels.append((i, name))
                break
        # If no labels yet and we see a check row, inject a CHECKS label
        if not labels and _CHECK_ROW.match(s):
            labels.append((i, "CHECKS"))

    if not labels:
        return []

    spans = []
    for (i, name), (j, _) in zip(labels, labels[1:] + [(len(lines), None)]):
        spans.append((i, j, name))
    return spans

 
# ---------- Public API ----------
def get_checking_band_lines(pdf_path: str):
    with pdfplumber.open(pdf_path) as pdf:
        page_texts = _page_texts_checking_band(pdf)
        merged = "\n".join([t for t in page_texts if t])
        lines  = (merged or "").splitlines()
        normed = [l.replace("\u00A0"," ").replace("\u2007"," ").replace("\u202F"," ") for l in lines]

        # NEW: optional intra-band section pruning (fail-open)
        spans = _find_section_spans(normed)
        return normed
    #  Commented out to figure out where checks are
    # if spans and any(lbl == "CHECKS" for _, _, lbl in spans):
    #   kept, kept_cnt = [], 0
    #    for a, b, label in spans:
    #        kept.extend(normed[a:b]); kept_cnt += (b - a)
    #        print(f"[markers] found {len(spans)} section spans:")
    #    for a, b, label in spans:
    #        print(f"    {label.lower()} ({a}–{b})")
    #        print(f"[markers] kept {kept_cnt} of {len(normed)} total lines")
    #    return kept
    #else:
    #    print("[markers] no CHECKS span; returning full band without pruning")
    #    return normed
    
def parse_pdf_totals(pdf_path: str) -> Dict[str, Optional[float]]:
    """
    Extracts (begin_balance, end_balance, total_deposits, total_withdrawals)
    from the PDF text. Returns None for any not found.
    """
    if pdfplumber is None:
        return {
            "begin": None,
            "end": None,
            "total_deposits": None,
            "total_withdrawals": None,
        }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            begin = end = dep = wd = None
            for txt in _page_texts(pdf):
                if begin is None:
                    m = BEGIN_BAL_RE.search(txt)
                    if m:
                        begin = _to_float(m.group(1))
                if end is None:
                    m = END_BAL_RE.search(txt)
                    if m:
                        end = _to_float(m.group(2))
                if dep is None:
                    m = TOT_DEP_RE.search(txt)
                    if m:
                        dep = _to_float(m.group(1))
                if wd is None:
                    m = TOT_WD_RE.search(txt)
                    if m:
                        wd = _to_float(m.group(1))
            return {
                "begin": begin,
                "end": end,
                "total_deposits": dep,
                "total_withdrawals": wd,
            }
    except Exception:
        return {
            "begin": None,
            "end": None,
            "total_deposits": None,
            "total_withdrawals": None,
        }

def verify_statement_pdf(
    pdf_path: str,
    df,  # pandas DataFrame with columns: Date (datetime64), Description (str), Amount (float)
    begin_balance: Optional[float],
    end_balance: Optional[float],
    *,
    within_checking_band: bool = False,   # <<< new flag
    max_fail: int = 50,
    debug: bool = False,
) -> Dict[str, Any]:
   
    """
    Check that each row's (date token, amount token) both appear on the SAME page of the PDF.
    If within_checking_band=True, we only search inside the Checking band on each page
    (Savings-only pages are effectively ignored).
    """
    report: Dict[str, Any] = {"issues": [], "summary": {}, "by_page": {}}

    if pdfplumber is None:
        report["summary"] = {"status": "skipped", "reason": "pdfplumber not installed"}
        return report

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Choose page texts: full-page vs. cropped Checking band
            if within_checking_band:
                page_texts = _page_texts_checking_band(pdf)
            else:
                page_texts = _page_texts(pdf)

            rows_checked = 0
            rows_found = 0
            by_page_hits = {i + 1: 0 for i in range(len(page_texts))}

            for idx, row in df.iterrows():
                # Skip synthetic adjustments if present
                if str(row.get("_src", "")).upper() == "ADJUST":
                    continue

                # Build tokens
                date_tokens = _mk_date_tokens(str(row["Date"].date()))
                amt_token = _fmt_amt(float(row["Amount"]))

                found_on_same_page = False
                page_for_date = None
                page_for_amt = None

                for pageno, txt in enumerate(page_texts, start=1):
                    if not txt:
                        continue  # page dropped or empty in band-mode
                    has_date = any(tok in txt for tok in date_tokens)
                    has_amt = (amt_token in txt) or (("$" + amt_token) in txt)

                    if has_date and page_for_date is None:
                        page_for_date = pageno
                    if has_amt and page_for_amt is None:
                        page_for_amt = pageno

                    if has_date and has_amt:
                        found_on_same_page = True
                        by_page_hits[pageno] += 1
                        break

                rows_checked += 1
                if found_on_same_page:
                    rows_found += 1
                else:
                    report["issues"].append(
                        {
                            "RowIndex": int(idx),
                            "Date": str(row["Date"].date()),
                            "Amount": float(row["Amount"]),
                            "Description": str(row["Description"]),
                            "PageForDate": page_for_date,
                            "PageForAmount": page_for_amt,
                        }
                    )
                    if debug and len(report["issues"]) >= max_fail:
                        break

            # Totals (unchanged)
            pdf_totals = parse_pdf_totals(pdf_path)
            report["by_page"] = by_page_hits
            report["summary"] = {
                "status": "ok",
                "rows_checked": rows_checked,
                "rows_matched_same_page": rows_found,
                "rows_missing": rows_checked - rows_found,
                "begin_balance_txt": pdf_totals.get("begin"),
                "end_balance_txt": pdf_totals.get("end"),
                "total_deposits_txt": pdf_totals.get("total_deposits"),
                "total_withdrawals_txt": pdf_totals.get("total_withdrawals"),
                "begin_balance_calc": begin_balance,
                "end_balance_calc": end_balance,
            }
            return report

    except Exception as e:
        if debug:
            import traceback
            print("[verify-pdf] Exception during verification:", repr(e))
            traceback.print_exc()
        report["summary"] = {"status": "error", "reason": str(e)}
        return report

