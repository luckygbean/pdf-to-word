param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'bin'), [string]$WindowsMetadata)
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
$null = New-Item -ItemType Directory -Force -Path $OutputDirectory
$exe = Join-Path $OutputDirectory 'PDFtoWord.exe'
$sources = Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | ForEach-Object FullName
$scriptResource = Join-Path $PSScriptRoot 'src/editable.ps1'
& (Join-Path $framework 'csc.exe') /nologo /codepage:65001 /optimize+ /target:winexe /platform:x64 "/win32manifest:$PSScriptRoot/app.manifest" "/out:$exe" "/resource:$scriptResource,editable.ps1" $refs $sources
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Write-Output $exe
