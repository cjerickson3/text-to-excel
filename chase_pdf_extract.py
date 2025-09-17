#!/usr/bin/env python3
"""
chase_pdf_extract.py

Extracts transactions from Chase checking statements into CSV with optional
categorization via a simple keyword rules CSV.

✅ Works on Python 3.10+.
✅ Includes a pure-Python fallback (pypdf) so it can run on 3.13 even if
   pdfplumber / PyMuPDF wheels aren't available for your platform yet.

Usage examples:
  # January statement (spans Dec + Jan)
  python chase_pdf_extract.py "20190107-statements-5263-.pdf" --jan-year 2019 --out 2019-01.csv

  # Restrict to a section by header text (regex allowed)
  python chase_pdf_extract.py "20190107-statements-5263-.pdf" --jan-year 2019 \
    --section-start "ATM\s*&\s*CARD\s*WITHDRAWALS" --section-stop "ELECTRONIC\s*WITHDRAWALS" \
    --out 2019-01_ATM.csv

  # Provide category rules
  python chase_pdf_extract.py "file.pdf" --jan-year 2019 --category-rules category_rules.csv

  # Force engine (auto|tables|text|pypdf)
  python chase_pdf_extract.py "file.pdf" --jan-year 2019 --engine pypdf
"""

from __future__ import annotations
import argparse
import json
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional

import pandas as pd

# Optional engines
try:
    import pdfplumber
except Exception:  # pragma: no cover
    pdfplumber = None

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None

DATE_MMDD = re.compile(r"^\s*(\d{2})/(\d{2})\b")
AMOUNT = re.compile(r"(-?\$?\s*\d{1,3}(?:,\d{3})*\.\d{2})\s*$")

DEFAULT_INCOME_KEYS = [
    "ONLINE TRANSFER FROM", "TRANSFER FROM", "DEPOSIT", "PAYROLL",
    "CHECK DEPOSIT", "ATM CHECK DEPOSIT", "INTEREST PAYMENT",
    "REFUND", "REVERSAL", "ACH CREDIT"
]

def clean_amount(a: str) -> Optional[float]:
    if a is None:
        return None
    a = a.replace("$", "").replace(",", "").replace(" ", "")
    try:
        return float(a)
    except Exception:
        return None

def infer_year_for_january_statement(mm: int, jan_year: int) -> int:
    # Chase January statement includes Dec (previous year) + Jan (jan_year)
    return (jan_year - 1) if mm == 12 else jan_year

def infer_sign(description: str, amt: float, income_keys=None) -> float:
    if income_keys is None:
        income_keys = DEFAULT_INCOME_KEYS
    d = (description or "").upper()
    return abs(amt) if any(k in d for k in income_keys) else -abs(amt)

def load_category_rules(csv_path: Optional[str]):
    rules = []
    if not csv_path:
        return rules
    p = Path(csv_path)
    if not p.exists():
        return rules
    df = pd.read_csv(p)
    # expected columns: keyword, category, match_type (contains|startswith|regex), case_sensitive (0/1)
    for _, r in df.iterrows():
        kw = str(r.get("keyword", "")).strip()
        cat = str(r.get("category", "")).strip() or "Other"
        mt = str(r.get("match_type", "contains")).strip().lower()
        cs = bool(r.get("case_sensitive", False))
        if kw:
            rules.append({"keyword": kw, "category": cat, "match_type": mt, "case_sensitive": cs})
    return rules

def apply_category_rules(description: str, rules):
    if not rules:
        return None
    text = description or ""
    for rule in rules:
        kw = rule["keyword"]
        cat = rule["category"]
        mt = rule["match_type"]
        cs = rule["case_sensitive"]
        hay = text if cs else text.upper()
        needle = kw if cs else kw.upper()
        try:
            if mt == "contains" and needle in hay:
                return cat
            if mt == "startswith" and hay.startswith(needle):
                return cat
            if mt == "regex" and re.search(kw, text):
                return cat
        except Exception:
            continue
    return None

def categorize_default(description: str) -> str:
    d = (description or "").upper()
    if "WAL-MART" in d or "WALMART" in d: return "Groceries"
    if d.startswith("TST*"): return "Food & drink"
    if "PERSHING" in d: return "401K transfer"
    if "ONLINE PAYMENT" in d: return "Online payment"
    if "HOME DEPOT" in d or "LOWES" in d: return "Home Repair"
    if "GOLF" in d: return "Golf"
    if "WELLCARE" in d or "CHIROPRACT" in d: return "Health & wellness"
    if "KERA" in d: return "Donations"
    return "Other"

def extract_with_pdfplumber(pdf_path: str, jan_year: Optional[int]) -> List[Tuple[str,str,float]]:
    rows = []
    if not pdfplumber:
        return rows
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for ts in [
                {"vertical_strategy": "lines", "horizontal_strategy": "lines", "intersection_tolerance": 5},
                {"vertical_strategy": "text", "horizontal_strategy": "text"},
            ]:
                try:
                    tables = page.extract_tables(table_settings=ts) or []
                except Exception:
                    tables = []
                for t in tables:
                    for r in t:
                        if not r: 
                            continue
                        r = [("" if c is None else str(c)).strip() for c in r]
                        if r and DATE_MMDD.match(r[0] or ""):
                            mm, dd = map(int, DATE_MMDD.match(r[0]).groups())
                            amt_val = None; amt_cell = None
                            for c in reversed(r):
                                if not c:
                                    continue
                                m_amt = AMOUNT.search(c.replace(" ", ""))
                                if m_amt:
                                    amt_cell = c
                                    amt_val = clean_amount(m_amt.group(1))
                                    break
                            if amt_val is None:
                                continue
                            mid = [c for c in r[1:] if c != amt_cell and c is not None]
                            desc = " ".join([m for m in mid if m]).strip()
                            if not desc:
                                continue
                            year = infer_year_for_january_statement(mm, jan_year) if jan_year else 1900
                            rows.append((f"{year:04d}-{mm:02d}-{dd:02d}", desc, amt_val))
    return rows

def _lines_from_pymupdf(pdf_path: str) -> List[str]:
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        txt = page.get_text("text")
        lines.extend([l.strip() for l in txt.splitlines() if l.strip()])
    return lines

def _lines_from_pypdf(pdf_path: str) -> List[str]:
    reader = PdfReader(pdf_path)
    lines = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        lines.extend([l.strip() for l in txt.splitlines() if l.strip()])
    return lines

def extract_textwise_lines(lines: List[str], jan_year: Optional[int]) -> List[Tuple[str,str,float]]:
    rows = []
    i = 0
    n = len(lines)
    while i < n:
        m_date = DATE_MMDD.match(lines[i])
        if m_date:
            mm, dd = map(int, m_date.groups())
            # Collect description lines until amount or next date
            j = i + 1
            desc_parts = []
            amt_val = None
            while j < n:
                s = lines[j]
                m_amt = AMOUNT.search(s.replace(" ", ""))
                if m_amt:
                    amt_val = clean_amount(m_amt.group(1))
                    break
                if DATE_MMDD.match(s):
                    break
                desc_parts.append(s.strip())
                j += 1
            if amt_val is not None and desc_parts:
                year = infer_year_for_january_statement(mm, jan_year) if jan_year else 1900
                date_iso = f"{year:04d}-{mm:02d}-{dd:02d}"
                desc = " ".join(desc_parts).strip()
                rows.append((date_iso, desc, amt_val))
                i = j + 1
                continue
        i += 1
    return rows

def extract_text_engine(pdf_path: str, jan_year: Optional[int]) -> List[Tuple[str,str,float]]:
    if fitz:
        lines = _lines_from_pymupdf(pdf_path)
        return extract_textwise_lines(lines, jan_year)
    return []

def extract_pypdf_engine(pdf_path: str, jan_year: Optional[int]) -> List[Tuple[str,str,float]]:
    if PdfReader:
        lines = _lines_from_pypdf(pdf_path)
        return extract_textwise_lines(lines, jan_year)
    return []

def filter_section(lines: List[str], start_re: Optional[str], stop_re: Optional[str]) -> List[str]:
    if not (start_re or stop_re):
        return lines
    start_pat = re.compile(start_re, re.I) if start_re else None
    stop_pat = re.compile(stop_re, re.I) if stop_re else None
    out = []
    capturing = start_pat is None
    for line in lines:
        if start_pat and start_pat.search(line):
            capturing = True
        if capturing:
            out.append(line)
        if stop_pat and stop_pat.search(line):
            capturing = False
    return out

def run_extract(pdf_path: str, jan_year: Optional[int], engine: str, section_start: Optional[str], section_stop: Optional[str]) -> Tuple[pd.DataFrame, str]:
    rows = []
    method = "auto"
    if engine in ("auto", "tables"):
        rows = extract_with_pdfplumber(pdf_path, jan_year)
        method = "tables"
    if engine in ("auto", "text"):
        if len(rows) < 10:  # fallback if few rows
            rows2 = extract_text_engine(pdf_path, jan_year)
            if len(rows2) > len(rows):
                rows = rows2
                method = "text"
    if engine in ("auto", "pypdf"):
        if len(rows) < 10:
            rows3 = extract_pypdf_engine(pdf_path, jan_year)
            if len(rows3) > len(rows):
                rows = rows3
                method = "pypdf"

    # If section filtering requested, re-run on lines for best granularity
    if section_start or section_stop:
        # Use whichever line engine we have
        lines = []
        if fitz:
            lines = _lines_from_pymupdf(pdf_path)
        elif PdfReader:
            lines = _lines_from_pypdf(pdf_path)
        if lines:
            lines = filter_section(lines, section_start, section_stop)
            sec_rows = extract_textwise_lines(lines, jan_year)
            if len(sec_rows) > 0:
                rows = sec_rows
                method = f"{method}+section"

    df = pd.DataFrame(rows, columns=["Date","Description","Amount"])
    return df, method

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Path to Chase statement PDF")
    ap.add_argument("--jan-year", type=int, default=None, help="If this is a January statement, provide the January year (e.g., 2019).")
    ap.add_argument("--engine", choices=["auto","tables","text","pypdf"], default="auto", help="Extraction engine to use.")
    ap.add_argument("--section-start", type=str, default=None, help="Regex marking the start of a section to capture")
    ap.add_argument("--section-stop", type=str, default=None, help="Regex marking the end of a section")
    ap.add_argument("--category-rules", type=str, default=None, help="CSV with columns: keyword,category,match_type,case_sensitive")
    ap.add_argument("--income-keys", type=str, default=None, help="JSON array override for income keywords")
    ap.add_argument("--out", type=str, default="out.csv", help="Output CSV path")
    args = ap.parse_args()

    rules = load_category_rules(args.category_rules)
    if args.income_keys:
        try:
            income_keys = json.loads(args.income_keys)
        except Exception:
            income_keys = DEFAULT_INCOME_KEYS
    else:
        income_keys = DEFAULT_INCOME_KEYS

    df, method = run_extract(args.pdf, args.jan_year, args.engine, args.section_start, args.section_stop)
    if df.empty:
        print("No transactions found.")
        Path(args.out).write_text("Date,Description,Category,Amount\n")
        return

    # Signs & categories
    df["Amount"] = df.apply(lambda r: infer_sign(r["Description"], float(r["Amount"]), income_keys), axis=1)
    cats = []
    for desc in df["Description"].astype(str):
        c = apply_category_rules(desc, rules) or categorize_default(desc)
        cats.append(c)
    df["Category"] = cats

    df = df[["Date","Description","Category","Amount"]]
    df.to_csv(args.out, index=False)
    print(f"Extraction engine: {method}. Wrote {len(df)} rows to {args.out}")

if __name__ == "__main__":
    main()
