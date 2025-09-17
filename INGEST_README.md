# Ingest Chase Statement Text → Dashboard

This script ingests a `pdftotext -raw` file (with our Chase markers) and appends
transactions to the correct year tab in your Dashboard, rebuilds summaries,
and formats dates as `yyyy/mm/dd`.

## Install

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install pandas openpyxl
Usage
powershell
Copy code
# Example for Nov–Dec 2018 statement
python ingest_statement_text.py --input 20181206.raw.txt --dashboard Chase_Budget_Dashboard.xlsx --rules category_rules.csv
How it decides the year
It pulls the year from the filename (e.g., 20181206.raw.txt → 2018).

For December statements that include November dates, it still assigns those lines to 2018 (per your preference).

What it extracts
Everything between *start*checks paid and *end*electronic withdrawal:

Checks Paid

ATM & Debit Card Withdrawals

Electronic Withdrawals

Skips headers and lines like Total ... Withdrawals.

Category rules (optional)
Provide a category_rules.csv with columns:

keyword

category

match_type (contains|startswith|regex)

case_sensitive (0|1)

Rules are applied first; if no rule matches, the script uses sensible defaults.

De-duplication
Dedupe key is (Date, Description, Amount) so re-running the same file won't double-count.

Output
Appends to the proper year sheet (e.g., 2018).

Rebuilds Monthly Summary, Yearly Summary, YOY Comparison.

Leaves your workbook saved in place.

pgsql
Copy code

---

👉 Save that into a file called `INGEST_README.md`.  
- If you open it in **Notepad**, it will look like plain text.  
- If you open it in **VS Code** or on **GitHub**, it will render with nice formatting.  

Do you want me to also give you a **plain `.txt` version** of the README so you don’t have to deal with Markdown at all?





Ask ChatGPT
