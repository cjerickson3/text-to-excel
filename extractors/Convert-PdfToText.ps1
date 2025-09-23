<#
.SYNOPSIS
  Convert PDF(s) from "Chase_history" to text files in "texts" using pdftotext.

.DESCRIPTION
  Runs from a workspace folder (defaults to C:\Users\chris\Documents\2425_Delmar\Budget).
  You may pass a specific file, a wildcard pattern, or a folder. By default it converts
  all PDFs under ".\Chase_history" and writes outputs to ".\texts\<base>-raw.txt".

  Requires "pdftotext.exe" (Poppler/Xpdf). Put it on PATH or provide -PdfToText.

.PARAMETER Workspace
  Root folder to operate from. Defaults to "C:\Users\chris\Documents\2425_Delmar\Budget".

.PARAMETER Input
  A file path, wildcard, or folder relative to Workspace (or absolute).
  Defaults to "Chase_history\*.pdf".

.PARAMETER Output
  Output folder relative to Workspace (or absolute). Defaults to "texts".

.PARAMETER Recurse
  When Input is a folder (or wildcard), recurse into subfolders.

.PARAMETER Overwrite
  Overwrite existing output files if they exist.

.PARAMETER Suffix
  Suffix to append to the base file name before ".txt". Defaults to "-raw.txt".

.PARAMETER PdfToText
  Path to pdftotext executable (if not on PATH). Defaults to "pdftotext.exe".

.PARAMETER Layout
  Preserve original physical layout (-layout). Enabled by default.

.PARAMETER NoPageBreaks
  Remove page breaks (-nopgbrk). Enabled by default.

.EXAMPLE
  # Convert everything under the default folders:
  .\Convert-PdfToText.ps1

.EXAMPLE
  # Convert a single file:
  .\Convert-PdfToText.ps1 -Input "Chase_history\2025-07-12.pdf"

.EXAMPLE
  # Convert matching PDFs, recursing into subfolders, overwriting any existing outputs:
  .\Convert-PdfToText.ps1 -Input "Chase_history\*.pdf" -Recurse -Overwrite

.EXAMPLE
  # Specify a custom pdftotext path:
  .\Convert-PdfToText.ps1 -PdfToText "C:\Tools\poppler\bin\pdftotext.exe"
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$Workspace   = "C:\Users\chris\Documents\2425_Delmar\Budget",
  [string]$Input       = "Chase_history\*.pdf",
  [string]$Output      = "texts",
  [switch]$Recurse,
  [switch]$Overwrite,
  [string]$Suffix      = "-raw.txt",
  [string]$PdfToText   = "pdftotext.exe",
  [switch]$Layout      = $true,
  [switch]$NoPageBreaks= $true
)

function Resolve-RootedPath {
  param([string]$Path, [string]$Root)
  if ([System.IO.Path]::IsPathRooted($Path)) { return (Resolve-Path -LiteralPath $Path).Path }
  $combined = Join-Path -Path $Root -ChildPath $Path
  if (Test-Path -LiteralPath $combined) { return (Resolve-Path -LiteralPath $combined).Path }
  # Might be a wildcard that doesn't resolve yet; return combined as-is
  return $combined
}

# 1) Resolve workspace
try {
  $ws = Resolve-Path -LiteralPath $Workspace -ErrorAction Stop
} catch {
  throw "Workspace not found: $Workspace"
}

# 2) Resolve input and output roots
$inArg  = Resolve-RootedPath -Path $Input  -Root $ws
$outDir = Resolve-RootedPath -Path $Output -Root $ws

# Create output folder if needed
if (-not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# 3) Collect PDF files
$items = @()

# If $inArg is a directory, collect PDFs inside it
if (Test-Path -LiteralPath $inArg -PathType Container) {
  $items = Get-ChildItem -LiteralPath $inArg -Filter *.pdf -File -Recurse:$Recurse
}
# If $inArg is a single file that exists, use it
elseif (Test-Path -LiteralPath $inArg -PathType Leaf) {
  $fi = Get-Item -LiteralPath $inArg
  if ($fi.Extension -ieq ".pdf") { $items = @($fi) }
  else { throw "Input file is not a PDF: $inArg" }
}
# Otherwise treat as wildcard (e.g., "Chase_history\*.pdf")
else {
  $dirOfWildcard = Split-Path -Path $inArg -Parent
  if (-not $dirOfWildcard) { $dirOfWildcard = $ws }
  $leafWildcard  = Split-Path -Path $inArg -Leaf
  if (-not $leafWildcard) { $leafWildcard = "*.pdf" }
  if (-not (Test-Path -LiteralPath $dirOfWildcard)) {
    throw "Input path does not exist: $dirOfWildcard"
  }
  $items = Get-ChildItem -Path $dirOfWildcard -Filter $leafWildcard -File -Recurse:$Recurse
}

if ($items.Count -eq 0) {
  Write-Warning "No PDFs found for input: $Input"
  return
}

# 4) Verify pdftotext is available
function Resolve-Exe {
  param([string]$Candidate)
  if ([System.IO.Path]::IsPathRooted($Candidate)) {
    if (Test-Path -LiteralPath $Candidate -PathType Leaf) { return $Candidate }
    throw "pdftotext not found at: $Candidate"
  }
  # Search in PATH
  $pathExts = (Get-ChildItem Env:Path).Value -split ';' | Where-Object { $_ -ne '' }
  foreach ($p in $pathExts) {
    $probe = Join-Path $p $Candidate
    if (Test-Path -LiteralPath $probe -PathType Leaf) { return $probe }
  }
  # Try with .exe appended if not provided
  if (-not $Candidate.EndsWith(".exe")) {
    foreach ($p in $pathExts) {
      $probe = Join-Path $p ($Candidate + ".exe")
      if (Test-Path -LiteralPath $probe -PathType Leaf) { return $probe }
    }
  }
  throw "pdftotext not found on PATH. Install Poppler/Xpdf utilities or pass -PdfToText with a full path."
}

$pdftotextExe = Resolve-Exe -Candidate $PdfToText

# 5) Process files
$converted = 0
$skipped   = 0
$failed    = 0

foreach ($pdf in $items) {
  try {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pdf.Name)
    $outPath  = Join-Path -Path $outDir -ChildPath ($baseName + $Suffix)

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $outPath)) {
      Write-Verbose "Skipping (exists): $outPath"
      $skipped++
      continue
    }

    $args = @()

    if ($Layout) { $args += "-layout" }
    if ($NoPageBreaks) { $args += "-nopgbrk" }
    # Force UTF-8 encoding from pdftotext
    $args += @("-enc", "UTF-8")

    # input and output
    $args += @($pdf.FullName, $outPath)

    if ($PSCmdlet.ShouldProcess($pdf.FullName, "pdftotext -> $outPath")) {
      $proc = Start-Process -FilePath $pdftotextExe -ArgumentList $args -NoNewWindow -PassThru -Wait
      if ($proc.ExitCode -ne 0) {
        throw "pdftotext exited with code $($proc.ExitCode)"
      }
      # Basic sanity-check: ensure file exists and is non-empty
      if (-not (Test-Path -LiteralPath $outPath)) {
        throw "pdftotext reported success but no output found: $outPath"
      }
      $len = (Get-Item -LiteralPath $outPath).Length
      if ($len -lt 3) {
        Write-Warning "Very small output ($len bytes): $outPath"
      }
      Write-Host "OK: $($pdf.Name) -> $([System.IO.Path]::GetFileName($outPath))"
      $converted++
    }
  } catch {
    Write-Error "FAILED: $($pdf.FullName) `n$($_.Exception.Message)"
    $failed++
  }
}

Write-Host "Done. Converted: $converted  Skipped: $skipped  Failed: $failed"
exit 0
