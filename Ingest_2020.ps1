# Set your paths
$root      = "C:\Users\chris\Documents\2425_Delmar\Budget"
$script    = Join-Path $root "ingest_statement_text.py"
$inDir     = Join-Path $root "History_text"
$dashboard = Join-Path $root "Chase_Budget_Dashboard.xlsx"   # change if yours is named differently
# $rules     = Join-Path $root "rules.csv"        # optional

# Ingest all 2020*-prefixed .txt files
Get-ChildItem -LiteralPath $inDir -File -Filter "2020*.txt" |
  Sort-Object Name |
  ForEach-Object {
    python $script --input $_.FullName --dashboard $dashboard 
    # If you don't have a rules file, drop:  --rules $rules
  }
