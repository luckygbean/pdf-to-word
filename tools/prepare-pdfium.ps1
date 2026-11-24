param(
    [Parameter(Mandatory = $true)] [string]$OutputDirectory,
    [string]$WheelPath
)
$ErrorActionPreference = 'Stop'
$url = 'https://files.pythonhosted.org/packages/fe/7f/8aa446529fa14523e5c18f89e823afc95526689f76a75efe68a82a10d850/pypdfium2-5.15.2-py3-none-win_amd64.whl'
$wheelHash = '74D1BACE34FB01C660D0B64344D5A4B7D276942B7C6623C88AB59EA21C5ACFBD'
$dllHash = '42D87F68957D745CE9377DFB1D8D63691E19082A7EC29FF421D22779D3142B29'
$null = New-Item -ItemType Directory -Force -Path $OutputDirectory
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$dllPath = Join-Path $directory 'pdfium.dll'
$noticesPath = Join-Path $directory 'PDFium-NOTICES.txt'
if ((Test-Path -LiteralPath $dllPath) -and (Test-Path -LiteralPath $noticesPath) -and
    ((Get-FileHash -LiteralPath $dllPath).Hash -eq $dllHash) -and (Get-Item -LiteralPath $noticesPath).Length -gt 1000) { return }
if (!$WheelPath) {
    $WheelPath = Join-Path $directory 'pypdfium2-5.15.2-py3-none-win_amd64.whl'
    if (!(Test-Path -LiteralPath $WheelPath) -or (Get-FileHash -LiteralPath $WheelPath).Hash -ne $wheelHash) {
        Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $WheelPath
    }
}
if ((Get-FileHash -LiteralPath $WheelPath).Hash -ne $wheelHash) { throw 'PDFium distribution checksum mismatch.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $WheelPath).Path)
try {
    $entry = $zip.GetEntry('pypdfium2_raw/pdfium.dll')
    [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dllPath, $true)
    if ((Get-FileHash -LiteralPath $dllPath).Hash -ne $dllHash) { throw 'Unexpected PDFium DLL.' }
    $notes = New-Object Text.StringBuilder
    $null = $notes.AppendLine('PDFium 159.0.8149.0, from pypdfium2 5.15.2 for Windows x64.')
    $null = $notes.AppendLine('https://pdfium.googlesource.com/pdfium/  https://github.com/pypdfium2-team/pypdfium2')
    $null = $notes.AppendLine("Distribution: $url")
    foreach ($license in ($zip.Entries | Where-Object { $_.FullName -like '*.dist-info/licenses/*' -and $_.Length -gt 0 } | Sort-Object FullName)) {
        $null = $notes.AppendLine()
        $null = $notes.AppendLine($license.FullName)
        $reader = New-Object IO.StreamReader($license.Open())
        try { $null = $notes.AppendLine($reader.ReadToEnd()) } finally { $reader.Dispose() }
    }
    [IO.File]::WriteAllText($noticesPath, $notes.ToString(), [Text.UTF8Encoding]::new($false))
} finally { $zip.Dispose() }
