# statement_to_excel

[![Version](https://img.shields.io/github/v/tag/cjerickson3/statement_to_excel?label=version&color=3E7EBB)](https://github.com/cjerickson3/statement_to_excel/tags)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](./LICENSE)

A Python utility that parses **Chase Bank statement text exports** into a structured **Excel dashboard** for personal budgeting and long-term financial analysis.

This project evolved from **text-to-excel** and introduces a refined parsing engine, PDF band verification, and fully automated workbook integration — ideal for long-term, reproducible budget tracking.

---

## 🧭 Overview

`statement_to_excel.py` ingests bank statement text files (converted from PDF) and populates **Chase_Budget_Dashboard.xlsx** with categorized transactions, monthly summaries, and optional verification sheets.

It automatically detects *Deposits*, *Checks*, *ATM*, and *Electronic Withdrawals*, applies geometry-aware clipping to exclude *Savings* transactions, and maintains consistent versioning using Git tags.

---

## ⚙️ Key Features

- 🧾 Converts Chase statement text to Excel dashboards  
- ✂️ Automatically excludes Savings transactions  
- 🧩 Handles multi-page statements and inter-account transfers  
- 🔍 Optional PDF band verification (`--verify-pdf`)  
- 🧠 Detects start/end of sections without page headers  
- 🪶 Lightweight, dependency-minimal core (pandas, openpyxl, pdfplumber)  
- 🧰 Includes Git-based version stamping in Excel exports  
- 🧪 Generates debug `.tsv` traces for QA (ignored by Git)

---

## 🚀 Getting Started

### 1️⃣ Clone the repository
```bash
git clone https://github.com/cjerickson3/statement_to_excel.git
cd statement_to_excel
### Set up a Python virtual environment
python -m venv .venv
# Activate it
#   On Windows (PowerShell):
.venv\Scripts\Activate.ps1
#   On macOS/Linux:
source .venv/bin/activate
2️⃣ Set up a Python virtual environment
python -m venv .venv
# Activate it
#   On Windows (PowerShell):
.venv\Scripts\Activate.ps1
#   On macOS/Linux:
source .venv/bin/activate

3️⃣ Install dependencies
pip install -r requirements.txt


(If no requirements.txt yet, you can install manually:)

pip install pandas openpyxl pdfplumber

🧩 Example Usage
python statement_to_excel.py ^
    --input History_text/20181106-raw.txt ^
    --dashboard Chase_Budget_Dashboard.xlsx ^
    --pdf Chase_history/20181106-statements-5263-.pdf ^
    --verify-pdf --auto-adjust --debug --force

Flag	Description
--input	Path to raw text file (from pdftotext)
--dashboard	Target Excel workbook
--pdf	Original PDF for verification
--verify-pdf	Run pdfplumber-based band verification
--auto-adjust	Trim text based on detected section boundaries
--force	Overwrite existing month data in Excel
--debug	Generate .tsv debug output
📁 Project Structure
statement_to_excel/
├── statement_to_excel.py
├── verifiers/
│   ├── pdf_plumber_verify.py
│   └── pdf_page_cuts.py
├── Chase_Budget_Dashboard.xlsx
├── History_text/
│   ├── 20181106-raw.txt
│   └── ...
└── debug_band_lines.txt

🧮 Versioning

Git tags serve as traceable checkpoints for each stable release.
Current baseline: v0.10.6 (cleanup + verification improvements).

Example local output:

Version # v0.10.6-1-gbea0d6e-dirty


Tagging a new version:

git tag -a v0.10.7 -m "Stable post-cleanup version"
git push origin v0.10.7

🧰 Development Notes

.gitignore excludes .tsv debug outputs

PDF clipping handled by verifiers/pdf_plumber_verify.py

Automatic Savings detection uses both geometry and section context

git describe used for version embedding into Excel output

🏁 Roadmap

 Add separate Savings worksheet

 Integrate automatic Ingest Log updates

 Build yearly spending rollup dashboard

 Add CLI summary flag (--summary-only)

📜 License

MIT License © 2025 Chris Erickson

Maintainer: @cjerickson3

Budget Building Project — Continuous refinement of Chase statement ingestion and budget analysis.


---

Would you like me to also generate a matching **`requirements.txt`** and a short **`setup_instructions.md`** (for collaborators or future automation)?  
They’d fit naturally alongside this README.
