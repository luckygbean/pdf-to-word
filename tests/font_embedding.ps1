# Run in Windows PowerShell with the Chinese supplemental fonts installed.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Split-Path $PSScriptRoot -Parent
$scratch = New-Item -Type Directory -Path (Join-Path ([IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString('N')))
$wordFile = Join-Path $scratch.FullName 'form.docx'
$fontFile = Join-Path $scratch.FullName 'embedded.ttf'
try {
    & (Join-Path $root 'bin/PDFtoWord.exe') edit (Join-Path $root 'samples/qualification.pdf') $wordFile | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Editable conversion failed.' }
    $zip = [IO.Compression.ZipFile]::OpenRead($wordFile)
    try {
        function Read-XmlPart([string]$name) {
            $stream = $zip.GetEntry($name).Open()
            try {
                $xml = New-Object Xml.XmlDocument
                $xml.Load($stream)
                return $xml
            } finally { $stream.Dispose() }
        }
        $fonts = Read-XmlPart 'word/fontTable.xml'
        $relationships = Read-XmlPart 'word/_rels/fontTable.xml.rels'
        $w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $simsun = $fonts.GetElementsByTagName('font',$w) | Where-Object { $_.GetAttribute('name',$w) -eq 'SimSun' }
        $regular = $simsun.GetElementsByTagName('embedRegular',$w).Item(0)
        $id = $regular.GetAttribute('id',$r)
        $rel = $relationships.DocumentElement.ChildNodes | Where-Object { $_.GetAttribute('Id') -eq $id }
        if ($rel.Type -ne "$r/font") { throw 'Missing embedded font relationship.' }
        $key = $regular.GetAttribute('fontKey',$w).Trim('{}').Replace('-','')
        $blob = $zip.GetEntry("word/$($rel.Target)")
        [IO.Compression.ZipFileExtensions]::ExtractToFile($blob,$fontFile)
        $bytes = [IO.File]::ReadAllBytes($fontFile)
        for ($i=0; $i -lt 32; $i++) {
            $bytes[$i] = $bytes[$i] -bxor [Convert]::ToByte($key.Substring((15-($i%16))*2,2),16)
        }
        [IO.File]::WriteAllBytes($fontFile,$bytes)
        $face = [Windows.Media.GlyphTypeface]::new([uri]$fontFile)
        foreach ($code in @(0x6295,0x6807,0x4eba,0x540d,0x79f0)) {
            $glyph = $face.CharacterToGlyphMap[$code]
            if (!$glyph -or $face.GetGlyphOutline($glyph,20,20).Bounds.IsEmpty) {
                throw ('Missing printable embedded glyph U+{0:X4}.' -f $code)
            }
        }
        'Verified embedded SimSun outlines and font relationship.'
    } finally { $zip.Dispose() }
} finally {
    foreach ($file in @($fontFile,$wordFile)) { if ([IO.File]::Exists($file)) { [IO.File]::Delete($file) } }
    [IO.Directory]::Delete($scratch.FullName)
}
