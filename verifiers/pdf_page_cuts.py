
"""
pdf_page_cuts.py
Crop each PDF page to the CHECKING transaction region using pdfplumber,
then return reconstructed text lines to feed the parser.

Strategy per page:
- Find the highest Y (lowest on the page) among these CHECKING headers:
    "DEPOSITS AND ADDITIONS", "CHECKS PAID",
    "ATM & DEBIT CARD WITHDRAWALS", "ELECTRONIC WITHDRAWALS"
  Use that as y_start.
- Find the first SAVINGS banner on the same page:
    "SAVINGS SUMMARY", "CHASE SAVINGS", or a standalone "SAVINGS" banner
  Use its top Y as y_end. If none, y_end = page.height.
- Crop to [0, y_start, page.width, y_end] and extract text.
- If no checking header found on the page, return the page's original text.

We keep page order and rejoin with form-feed separators so downstream page indexing is preserved.
"""
from __future__ import annotations
import re
from typing import List

try:
    import pdfplumber  # type: ignore
except Exception:
    pdfplumber = None

CHECKING_HDRS = [
    r"DEPOSITS\s+AND\s+ADDITIONS",
    r"CHECKS?\s+PAID",
    r"ATM\s*&\s*DEBIT\s*CARD\s+WITHDRAWALS",
    r"ELECTRONIC\s+WITHDRAWALS"
]
SAVINGS_BANNERS = [
    r"SAVINGS\s+SUMMARY",
    r"CHASE\s+SAVINGS",
    r"^SAVINGS\b"
]
HDR_RE_LIST = [re.compile(h, re.I) for h in CHECKING_HDRS]
SAV_RE_LIST = [re.compile(s, re.I) for s in SAVINGS_BANNERS]

def _find_y_of_phrase(page, regexes) -> list:
    """Return a list of top-Y positions for any word-run that matches regex on page.extract_words()."""
    words = page.extract_words() or []
    # Build a full-line strings with bounding box union per line (group by rounded 'top')
    # Then search regexes in that line text, track the min top of a matching segment.
    by_top = {}
    for w in words:
        top = round(float(w.get("top", 0)), 1)
        row = by_top.setdefault(top, {"text": [], "x0": float("inf"), "x1": 0.0, "top": top, "bottom": top})
        row["text"].append(w["text"])
        row["x0"] = min(row["x0"], float(w.get("x0", 0)))
        row["x1"] = max(row["x1"], float(w.get("x1", 0)))
        row["bottom"] = max(row["bottom"], float(w.get("bottom", top)))
    hits = []
    for row in sorted(by_top.values(), key=lambda r: r["top"]):
        line = " ".join(row["text"])
        for rx in regexes:
            if rx.search(line):
                hits.append(row["top"])
                break
    return hits

CHECKING_HDRS = [
    r"DEPOSITS\s+AND\s+ADDITIONS",
    r"CHECKS?\s+PAID",
    r"ATM\s*&\s*DEBIT\s*CARD\s+WITHDRAWALS",     # matches "(continued)" too
    r"ELECTRONIC\s+WITHDRAWALS"                  # matches "(continued)" too
]
SAVINGS_BANNERS = [
    r"SAVINGS\s+SUMMARY",
    r"CHASE\s+SAVINGS",
    r"^SAVINGS\b"
]
HDR_RE_LIST = [re.compile(h, re.I) for h in CHECKING_HDRS]
SAV_RE_LIST = [re.compile(s, re.I) for s in SAVINGS_BANNERS]

def pdf_clip_checking_pages(pdf_path: str, raw_lines: List[str], debug: bool=False) -> List[str]:
    """
    Given the statement PDF and pdftotext raw_lines (with form feeds), return a new list of lines where
    each page is cropped to the CHECKING transaction band. Pages with no checking header are dropped.
    """
    if pdfplumber is None:
        return raw_lines

    # keep page boundaries
    pages_txt = []
    current = []
    for ln in raw_lines:
        if "\f" in ln:
            pre, post = ln.split("\f", 1)
            current.append(pre)
            pages_txt.append("\n".join(current))
            current = [post] if post else []
        else:
            current.append(ln)
    if current:
        pages_txt.append("\n".join(current))

    try:
        with pdfplumber.open(pdf_path) as pdf:
            new_pages: List[str] = []
            for i, page in enumerate(pdf.pages):
                orig_txt = pages_txt[i] if i < len(pages_txt) else ""

                # find checking header line(s)
                ys_checking = _find_y_of_phrase(page, HDR_RE_LIST)
                if not ys_checking:
                    # No checking header on this page -> DROP page (it's Savings or non-detail)
                    if debug:
                        print(f"[pdf-cuts] Page {i+1}: no checking header; DROPPED")
                    continue

                # find first savings banner position
                ys_savings = _find_y_of_phrase(page, SAV_RE_LIST)
                y_start = min(ys_checking)                   # top of checking area
                y_end   = min(ys_savings) if ys_savings else page.height

                if y_end <= y_start:
                    # Defensive: if savings appears above, keep only from header down to page end
                    y_end = page.height

                band = (0, y_start, page.width, y_end)
                sub  = page.crop(band)
                band_txt = sub.extract_text() or ""
                if debug:
                    print(f"[pdf-cuts] Page {i+1}: cropped {y_start:.1f}..{y_end:.1f}, kept_len={len(band_txt)}")
                new_pages.append(band_txt)

            # Rebuild into lines with form feeds preserved
            rebuilt = []
            for idx, txt in enumerate(new_pages):
                rebuilt.extend((txt or "").splitlines())
                if idx < len(new_pages) - 1:
                    rebuilt.append("\f")
            return rebuilt
    except Exception:
        return raw_lines

