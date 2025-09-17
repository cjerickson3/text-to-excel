# text_to_excel

Parses bank/statement text files and produces an Excel dashboard with:
- Transaction detail
- Monthly & yearly summaries
- Balance reconciliation (v0.10 “fix sign” working baseline)

This repo originated from a working script previously named `ingest_statement_text_balrecon_fixsign_v10.py`.  
The canonical entrypoint is now **`text_to_excel.py`**.

---

## Quick start

### 1) Create and activate a virtual environment (Windows)

**PowerShell**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

**Git Bash**
```bash
python -m venv .venv
source .venv/Scripts/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 2) Run
```bash
python text_to_excel.py --debug --input "History_text\20190107-raw.txt" --dashboard "Chase_Budget_Dashboard.xlsx"
```

> Replace the input file with any statement text you’ve exported. The script will populate the dashboard with detail, monthly/yearly summaries, and balance reconciliation.

---

## Command-line options (current)
- `--input <path>`: raw text statement file (e.g., from bank statement export)
- `--dashboard <path>`: output Excel workbook to create/update
- `--debug` (optional): verbose logs while parsing

*(We’ll keep this section in sync with the script’s `argparse` as we evolve.)*

---

## Features

- Robust text->structured ingest for bank statements (Chase first; others can be added)
- Writes a single Excel dashboard with:
  - **Detail** sheet (canonical transaction table)
  - **Monthly summary** (pivot)
  - **Yearly summary** (pivot)
  - **Balance reconciliation** (validated in v0.10 with corrected sign handling)

### Roadmap / Open items
- Split Deposits into subcategories: **Payroll**, **Transfer In**, **Check Deposit**, **Return**
- Replace post-load dedupe with a **load-once guard** (avoid accidental double-ingest)
- Add unit tests for parsers and reconciliation math
- Optional parsers for additional institutions (Chase vs others)
- Configurable category rules

---

## Development

### Conventions
- Use `dev` branch for WIP; keep `main` green.
- Tag releases: `v0.10`, `v0.11`, …
- Prefer LF line endings across the repo. Windows-compatible scripts can use CRLF—see `.gitattributes` below.

### Line endings (Windows vs. LF)
This project includes a `.gitattributes` that enforces LF by default and CRLF for Windows-only scripts like `.ps1`. VS Code handles LF fine on Windows.

If you prefer Windows-style locally while keeping LF in commits:
```bash
git config core.autocrlf true
```

### VS Code debug
A sample `launch.json` is provided to run the script with your usual arguments.

---

## Troubleshooting

**Curl/Download TLS errors on Windows (Schannel/CRL/OCSP):**  
Use PowerShell’s `Invoke-WebRequest`, or paste the file manually. For one-off trusted fetches, `curl -k` works but skips verification (not recommended).

**Git Bash vs PowerShell:**  
Git commands are the same. Only venv activation differs:
- PowerShell: `.\.venv\Scripts\Activate.ps1`
- Git Bash: `source .venv/Scripts/activate`

**LF→CRLF warnings:**  
Harmless. The included `.gitattributes` establishes a consistent policy; you can also run:
```bash
git add --renormalize .
```

---

## Example Git workflow

```bash
# First commit
git add .
git commit -m "chore: initial project scaffold"
git tag -a v0.10 -m "Working balance recon (fix sign) v10"
git push -u origin main --tags

# Work on dev
git checkout -b dev
# ... make changes ...
git commit -m "feat: deposit subcategories (payroll, transfer, check, return)"
git push -u origin dev
```

---

## License
Proprietary (default). If you want this open source, add an OSI-approved LICENSE (MIT/Apache-2.0/BSD-3-Clause).
