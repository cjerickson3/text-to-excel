Budget Building – Statement Ingestor

Small Python utility to ingest Chase statement text (from pdftotext -raw) into an Excel Dashboard with per-year sheets and monthly/yearly summaries.

What it does

Parses transactions from Checks Paid and Deposits & Additions.

Correctly assigns year/month for statements that span two months.

Writes to per-year sheets (e.g., 2018, 2019) with merge + de-dupe on (Date, Description, Amount).

Builds:

Monthly Summary (Category totals + TOTAL row per month)

Yearly Summary (Category totals + TOTAL row per year)

YOY Comparison (pivot across years)

Year/Month assignment (final rule)

We derive the statement’s end date from the filename YYYYMMDD (e.g., 20190107-raw.txt → end_year=2019, end_month=1).

For each transaction with month m:

if m == end_month:                    year = end_year
elif m == end_month - 1:              year = end_year
elif end_month == 1 and m == 12:      year = end_year - 1   # Dec→Jan special
else:                                 year = end_year       # safe fallback


Dates are stored as real datetimes; grouping by df["Date"].dt.year selects the correct year sheet.

Parsing notes (how we grab the right lines)
Deposits & Additions

We look for a header like Deposits and Additions and then capture the contiguous run of lines that start with a date (MM/DD, M-D, or MM/DD/YYYY).

We skip a table header row (DATE DESCRIPTION AMOUNT) if present.

We ignore summary lines like Checks Paid -1,940.66 (those include amounts) and only stop on true major section headers (e.g., CHECKING SUMMARY, CHECK NO. DESCRIPTION, CHECKS PAID, ATM & DEBIT CARD WITHDRAWALS, ELECTRONIC WITHDRAWALS).

If multiple deposit-ish headers exist, we choose the best run (deposit-keyword score → length → proximity to header).

Amounts from this section are forced positive (source tag _src="DEP_ADD").

Checks

A check row is recognized only if it has a real check number (or the Checks Paid header). We do not stop on the word “check” inside deposit descriptions (e.g., ATM Check Deposit).

Safer date matching

Date regexes prefer two-digit day matches first (fixes 12/10 mis-capturing as 12/1).

Range-checked months/days prevent false hits like 1-800-….

Amount sign logic

Deposits & Additions: positive.

Everything else uses a conservative sign heuristic:

If description contains a known income keyword (e.g., “PAYROLL”, “DEPOSIT”, “TRANSFER FROM”, “ACH CREDIT”, “INTEREST PAYMENT”, “REFUND”, “REVERSAL”), treat as positive.

Otherwise negative (spend).

You can expand the keyword lists in the code or keep using your rules CSV.

Merge & de-dupe (idempotent writes)

When writing a year sheet:

Read existing rows (if any).

concat with new rows.

De-dupe on Date + Description + Amount.

Sort by Date; write back with formatted columns.

This lets you re-run the same statement without duplicating transactions.

Summaries

Monthly Summary
Group by Month (YYYY-MM) × Category, sum Amount, append a per-month TOTAL row, sort month-chronologically and pin TOTAL last within each month for readability.

Yearly Summary
Group by Year × Category, sum Amount, append a per-year TOTAL row, sort by year and pin TOTAL.

YOY Comparison
Pivot Yearly Summary to columns by Year, index Category, fill_value=0.

CLI
# Single file
python ingest_statement_text.py --input History_text\20190107-raw.txt --dashboard Chase_Budget_Dashboard.xlsx

# With debug traces
python ingest_statement_text.py --debug --input History_text\20190107-raw.txt --dashboard Chase_Budget_Dashboard.xlsx


PowerShell helpers

# All 2019 statements
Get-ChildItem .\History_text\2019*-raw.txt | ForEach-Object {
  python ingest_statement_text.py --input $_.FullName --dashboard .\Chase_Budget_Dashboard.xlsx
}

# All years
Get-ChildItem .\History_text\*-raw.txt | ForEach-Object {
  python ingest_statement_text.py --input $_.FullName --dashboard .\Chase_Budget_Dashboard.xlsx
}


VS Code launch.json arg example

"args": ["--debug", "--input", "History_text\\20190107-raw.txt", "--dashboard", "Chase_Budget_Dashboard.xlsx"]

Known gotchas / tips


Deposits not found
Use --debug to verify the chosen window; deposits should show 5–20 lines, not card-purchase/ATM rows.

Extending categorization

Optionally supply a CSV of rules:

keyword,category,match_type,case_sensitive
PAYROLL,Income,contains,false
HOME DEPOT,Home Repair,contains,false
^TST\*,Food & drink,regex,false


match_type: contains | startswith | regex.

Mini changelog (what got fixed)

Correct year assignment for straddling statements (Dec→Jan edge included).

Stopped mis-parsing 12/10 as 12/1.

Prevented “Checks Paid -$X” summary from truncating the deposits window.

Deposits block chosen by header + scoring (won’t drift into debit card section).

Idempotent per-year writes with cross-sheet de-dupe.

Added --debug traces for window selection and row assembly.