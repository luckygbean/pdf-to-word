param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'bin'), [string]$WindowsMetadata, [switch]$EngineOnly)
$ErrorActionPreference = 'Stop'
$framework = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
$sdk = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10/UnionMetadata'
if (!$WindowsMetadata -and (Test-Path $sdk)) {
    $WindowsMetadata = Get-ChildItem $sdk -Directory | Where-Object Name -Match '^10\.' |
        Sort-Object { [version]($_.Name) } -Descending | ForEach-Object { Join-Path $_.FullName 'Windows.winmd' } |
        Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (!$WindowsMetadata) { $WindowsMetadata = Join-Path $PSScriptRoot 'obj/Windows.winmd' }
if (!(Test-Path -LiteralPath $WindowsMetadata)) { throw 'Install the Windows 10/11 SDK, or pass -WindowsMetadata with its UnionMetadata Windows.winmd path.' }
$refs = @('System.Runtime.dll','System.Runtime.InteropServices.WindowsRuntime.dll','System.Runtime.WindowsRuntime.dll',
    'System.Collections.dll','System.ObjectModel.dll','System.Threading.Tasks.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Windows.Forms.dll','System.Drawing.dll','System.Xml.Linq.dll') |
    ForEach-Object { '/reference:' + (Join-Path $framework $_) }
$refs += '/reference:' + (Resolve-Path -LiteralPath $WindowsMetadata).Path
$refs += @('PresentationCore.dll','WindowsBase.dll') | ForEach-Object { '/reference:' + (Join-Path $framework "WPF/$_") }
$null = New-Item -ItemType Directory -Force -Path $OutputDirectory
$sources = Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs'
if ($EngineOnly) {
    $binary = Join-Path $OutputDirectory 'PDFtoWord.Engine.dll'
    $target = @('/target:library')
    $sources = $sources | Where-Object { $_.Name -notin @('Program.cs','MainForm.cs') }
} else {
    $binary = Join-Path $OutputDirectory 'PDFtoWord.exe'
    $target = @('/target:winexe',"/win32manifest:$PSScriptRoot/app.manifest")
    $sources = $sources | Where-Object { $_.Name -ne 'PythonApi.cs' }
}
$sourcePaths = $sources | ForEach-Object FullName
$scriptResource = Join-Path $PSScriptRoot 'src/editable.ps1'
$pdfiumDirectory = Join-Path $PSScriptRoot 'obj/pdfium'
& (Join-Path $PSScriptRoot 'tools/prepare-pdfium.ps1') -OutputDirectory $pdfiumDirectory
$resources = @("/resource:$scriptResource,editable.ps1",
    "/resource:$pdfiumDirectory/pdfium.dll,pdfium.dll",
    "/resource:$pdfiumDirectory/PDFium-NOTICES.txt,pdfium.NOTICES.txt")
& (Join-Path $framework 'csc.exe') /nologo /codepage:65001 /optimize+ /platform:x64 $target "/out:$binary" $resources $refs $sourcePaths
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Write-Output $binary
