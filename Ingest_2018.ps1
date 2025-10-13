# Ingest_2018.ps1
# Runs statement_to_excel.py for all 2018 text files, pairing the correct PDF and enabling --verify-pdf.
# Safe path handling, clear logging, and optional rules file.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Config (edit these if your folders move) ---
$root      = "C:\Users\chris\Documents\2425_Delmar\Budget"
$script    = Join-Path $root "statement_to_excel.py"
$inDir     = Join-Path $root "History_text"
$pdfDir    = Join-Path $root "Chase_history"
$dashboard = Join-Path $root "Chase_Budget_Dashboard.xlsx"
$rules     = Join-Path $root "category_rules.csv"   # optional
$yearPrefix = "2018"

# --- Guard checks ---
if (-not (Test-Path -LiteralPath $script))    { throw "Missing script: $script" }
if (-not (Test-Path -LiteralPath $inDir))     { throw "Missing input dir: $inDir" }
if (-not (Test-Path -LiteralPath $dashboard)) {
  Write-Host "Dashboard not found; the script will create it at: $dashboard" -ForegroundColor Yellow
}
if (-not (Test-Path -LiteralPath $pdfDir)) {
  Write-Host "PDF dir not found ($pdfDir). I’ll run without --pdf/--verify-pdf." -ForegroundColor Yellow
}

# Optional args as arrays (PowerShell-friendly splatting later)
$rulesArg = @()
if (Test-Path -LiteralPath $rules) { $rulesArg = @('--rules', (Resolve-Path -LiteralPath $rules).Path) }

# Helper: try to find the PDF that matches an input *.txt based on the leading yyyymmdd
function Get-MatchingPdfForTxt {
  param(
    [Parameter(Mandatory=$true)] [IO.FileInfo] $TextFile,
    [Parameter(Mandatory=$true)] [string] $PdfRoot
  )
  # Extract leading 8 digits (yyyymmdd) from basename like "20181106-raw"
  if ($TextFile.BaseName -notmatch '^(?<date>\d{8})') { return $null }
  $date = $Matches['date']

  if (-not (Test-Path -LiteralPath $PdfRoot)) { return $null }

  # Prefer files beginning with that date; if multiple exist, bias toward names containing "statement" or the last 4 digits of acct if present
  $candidates = Get-ChildItem -LiteralPath $PdfRoot -File -Filter "$date*statements*.pdf"
  if (-not $candidates) { return $null }

  # Heuristic ranking: contains "statement" > contains "statements" > everything else
  $ranked = $candidates | Sort-Object {
    $n = $_.Name.ToLowerInvariant()
    if ($n -match 'statement') { 0 }
    elseif ($n -match 'statements') { 1 }
    else { 2 }
  }, Name

  return $ranked[0]
}

Write-Host "==== Ingesting $yearPrefix files ====" -ForegroundColor Cyan
$files = Get-ChildItem -LiteralPath $inDir -File -Filter "$yearPrefix*.txt" | Sort-Object Name
if (-not $files) {
  Write-Host "No files like $yearPrefix*.txt in $inDir" -ForegroundColor Yellow
  exit 0
}

foreach ($f in $files) {
  $inPath  = (Resolve-Path -LiteralPath $f.FullName).Path
  $pdfPath = $null

  if (Test-Path -LiteralPath $pdfDir) {
    $match = Get-MatchingPdfForTxt -TextFile $f -PdfRoot $pdfDir
    if ($match) { $pdfPath = (Resolve-Path -LiteralPath $match.FullName).Path }
  }

  # Build the argument list as separate tokens: no manual quoting needed
  $argsList = @(
    $script,
    '--input',     $inPath,
    '--dashboard', $dashboard,
    '--auto-adjust',
    '--audit'
  ) + $rulesArg

  if ($pdfPath) {
    $argsList += @('--pdf', $pdfPath, '--verify-pdf')
  } else {
    Write-Host ("[warn] No matching PDF for {0} → running without --pdf/--verify-pdf" -f $f.Name) -ForegroundColor Yellow
  }

  Write-Host (" -> {0}" -f $f.Name) -ForegroundColor Gray
  Write-Host ("    python {0}" -f ($argsList -join ' ')) -ForegroundColor DarkGray

  # Invoke python with the args
  & python @argsList
  if ($LASTEXITCODE -ne 0) {
    Write-Host ("[error] python exited with code $LASTEXITCODE for {0}" -f $f.Name) -ForegroundColor Red
    break
  }
}
Write-Host "Done." -ForegroundColor Green
