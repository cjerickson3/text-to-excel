# Set your paths
$root      = "C:\Users\chris\Documents\2425_Delmar\Budget"
$script    = Join-Path $root "text-to-excel.py"
$inDir     = Join-Path $root "History_text"
$dashboard = Join-Path $root "Chase_Budget_Dashboard.xlsx"   # change if yours is named differently
$rules     = Join-Path $root "category_rules.csv"        # optional

# Ingest all 2019*-prefixed .txt files
Get-ChildItem -LiteralPath $inDir -File -Filter "2019*.txt" |
  Sort-Object Name |
  ForEach-Object {
    python $script --input $_.FullName --dashboard $dashboard --rules $rules
    # If you don't have a rules file, drop:  --rules $rules --auto-adjust
  }
