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


def _page_texts(pdf) -> List[str]:
    out: List[str] = []
    for p in pdf.pages:
        out.append(p.extract_text() or "")
    return out


# ---------- Public API ----------

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
# --- Checking-band text extractor (per page) ---
# Returns a list (same length as pdf.pages) with the text cropped to the Checking band.
# If a page has no Checking header, we return "" for that page (so Savings-only pages won't match).
def _page_texts_checking_band(pdf):
    # Checking headers (also matches "(continued)")
    _CHECKING_HDRS = [
        r"DEPOSITS\s+AND\s+ADDITIONS",
        r"CHECKS?\s+PAID",
        r"ATM\s*&\s*DEBIT\s*CARD\s+WITHDRAWALS",
        r"ELECTRONIC\s+WITHDRAWALS",
    ]
    # Savings banners
    _SAVINGS_BANNERS = [
        r"SAVINGS\s+SUMMARY",
        r"CHASE\s+SAVINGS",
        r"^SAVINGS\b",
    ]
    _HDR_RE = [re.compile(p, re.I) for p in _CHECKING_HDRS]
    _SAV_RE = [re.compile(p, re.I) for p in _SAVINGS_BANNERS]

    def _find_y(page, pats):
        words = page.extract_words() or []
        # group by line (rounded 'top'), then test each line string against any pattern
        by_top = {}
        for w in words:
            top = round(float(w.get("top", 0.0)), 1)
            row = by_top.setdefault(top, {"text": [], "top": top})
            row["text"].append(w.get("text", ""))
        hits = []
        for row in sorted(by_top.values(), key=lambda r: r["top"]):
            line = " ".join(row["text"])
            if any(rx.search(line) for rx in pats):
                hits.append(row["top"])
        return hits

    out = []
    for page in pdf.pages:
        ys_check = _find_y(page, _HDR_RE)
        if not ys_check:
            out.append("")  # drop page from consideration
            continue
        y_start = min(ys_check)
        ys_sav  = _find_y(page, _SAV_RE)
        y_end   = min(ys_sav) if ys_sav else page.height
        if y_end <= y_start:
            y_end = page.height
        band = (0, y_start, page.width, y_end)
        txt  = (page.crop(band).extract_text() or "")
        out.append(txt)
    return out
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
        report["summary"] = {"status": "error", "reason": str(e)}
        return report



