$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $root 'bin/PDFtoWord.exe'
$pdf = Join-Path $root 'samples/inspection.pdf'
$target = Join-Path ([IO.Path]::GetTempPath()) ("pdfword-smoke-" + [guid]::NewGuid() + '.docx')
try {
    & $exe fixed $pdf $target 300 | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Fixed-page conversion failed.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($target)
    try {
        if (@($zip.Entries | Where-Object FullName -like 'word/media/*.png').Count -ne 3) { throw 'Expected three page images.' }
        $stream = $zip.GetEntry('word/document.xml').Open()
        try {
            $doc = New-Object Xml.XmlDocument
            $doc.Load($stream)
        } finally { $stream.Dispose() }
        $sizes = $doc.GetElementsByTagName('pgSz','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $expected = @('11900x16840','13200x8600','8400x11900')
        $actual = @($sizes | ForEach-Object { $_.GetAttribute('w:w') + 'x' + $_.GetAttribute('w:h') })
        if (($actual -join ',') -ne ($expected -join ',')) { throw "Unexpected page sizes: $actual" }
    } finally { $zip.Dispose() }
    $hash = (Get-FileHash -LiteralPath $target).Hash
    $job = Start-Process -WindowStyle Hidden -PassThru -Wait -FilePath $exe -ArgumentList ('fixed "'+$pdf+'" "'+$target+'"')
    if ($job.ExitCode -eq 0 -or $hash -ne (Get-FileHash -LiteralPath $target).Hash) { throw 'Existing file protection failed.' }
    Write-Output 'Passed page dimensions, page count and overwrite protection.'
} finally { if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target } }
