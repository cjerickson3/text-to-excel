# --- Config ---
$root      = "C:\Users\chris\Documents\2425_Delmar\Budget"
$script    = Join-Path $root "text-to-excel.py"
$inDir     = Join-Path $root "History_text"
$dashboard = Join-Path $root "Chase_Budget_Dashboard.xlsx"
$rules     = Join-Path $root "category_rules.csv"       # optional

# --- Just use 20190107-raw.txt ---
$years = Get-ChildItem -LiteralPath $inDir -File -Filter '2018*.txt' |
  ForEach-Object { if ($_.BaseName -match '^(?<y>\d{4})') { $Matches['y'] } } |
  Sort-Object -Unique

if (-not $years) {
  Write-Host "No year-prefixed files found in $inDir." -ForegroundColor Yellow
  return
}
# Refresh Dashboard file before doing 2018.
Copy-Item .\templates\Chase_Budget_Dashboard.xlsx ..\
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
      & python $script --input $_.FullName --audit --auto-adjust --dashboard $dashboard @rulesArg
      
    }
}