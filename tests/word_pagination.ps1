# Run on an interactive desktop with Microsoft Word installed.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$sample = Join-Path $root 'samples/qualification.pdf'
$output = Join-Path ([IO.Path]::GetTempPath()) ('pdfword-' + [Guid]::NewGuid() + '.docx')
$word = $null
$doc = $null
function Assert-OnPage([string]$phrase, [int]$expectedPage) {
    $range = $doc.Content.Duplicate
    if (!$range.Find.Execute($phrase)) { throw "Missing text: $phrase" }
    if ($range.Information(3) -ne $expectedPage) { throw "Unexpected page for $phrase." }
}
try {
    & (Join-Path $root 'bin/PDFtoWord.exe') edit $sample $output | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Conversion failed.' }
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($output,$false,$true,$false)
    $doc.Repaginate()
    if ($doc.ComputeStatistics(2) -ne 2) { throw 'Expected two pages.' }
    Assert-OnPage '备注' 1
    Assert-OnPage '3.5.1' 1
    Assert-OnPage '银行名称' 2
    Assert-OnPage '银行地址' 2
    Assert-OnPage '3.5.2' 2
    $counts = @($doc.Tables | ForEach-Object { $_.Range.Cells.Count })
    if (($counts -join ',') -ne '59,17,43') { throw "Incorrect table cell counts: $counts" }
    Write-Output 'Word verified page breaks, bank field labels and native table cell counts.'
} finally {
    try {
        if ($null -ne $doc) { $doc.Close([ref]0) }
    } finally {
        if ($null -ne $word) { $word.Quit([ref]0) }
        if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output }
    }
}
