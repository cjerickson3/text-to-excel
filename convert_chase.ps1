Get-ChildItem .\Chase_history\*-statements-5263-.pdf | ForEach-Object {
    $stem = $_.BaseName.Split("-")[0]
    $out = ".\History_text\$stem-raw.txt"
    Write-Host "Converting $($_.Name) -> $(Split-Path $out -Leaf)"
    pdftotext -raw $_.FullName $out
}
