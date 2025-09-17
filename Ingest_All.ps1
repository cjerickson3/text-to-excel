# --- Config ---
$root      = "C:\Users\chris\Documents\2425_Delmar\Budget"
$script    = Join-Path $root "ingest_statement_text.py"
$inDir     = Join-Path $root "History_text"
$dashboard = Join-Path $root "Chase_Budget_Dashboard.xlsx"
$rules     = Join-Path $root "rules.csv"       # optional

# --- Discover available years from filenames like 20190107-raw.txt ---
$years = Get-ChildItem -LiteralPath $inDir -File -Filter '20*.txt' |
  ForEach-Object { if ($_.BaseName -match '^(?<y>\d{4})') { $Matches['y'] } } |
  Sort-Object -Unique

if (-not $years) {
  Write-Host "No year-prefixed files found in $inDir." -ForegroundColor Yellow
  return
}

# Build optional rules arg only if the file exists
$rulesArg = @()
if (Test-Path $rules) { $rulesArg = @('--rules', $rules) }

Write-Host "Years detected: $($years -join ', ')" -ForegroundColor Cyan

foreach ($y in $years) {
  Write-Host "==== Processing year $y ====" -ForegroundColor Cyan
  Get-ChildItem -LiteralPath $inDir -File -Filter "$y*.txt" |
    Sort-Object Name |
    ForEach-Object {
      Write-Host (" -> " + $_.Name)
      & python $script --input $_.FullName --dashboard $dashboard @rulesArg
      # If python isn't on PATH, use:  & py -3 $script --input $_.FullName --dashboard $dashboard @rulesArg
    }
}
