Param(
    [Parameter(Mandatory=$true)][string]$DocxPath,
    [Parameter(Mandatory=$true)][string]$OutPath
)

$ErrorActionPreference = 'Stop'

if (!(Test-Path -LiteralPath $DocxPath)) {
    throw "DOCX file not found: $DocxPath"
}

$outDir = [System.IO.Path]::GetDirectoryName($OutPath)
if (![string]::IsNullOrEmpty($outDir) -and !(Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$tmp = Join-Path $outDir ("docx_tmp_" + [IO.Path]::GetFileNameWithoutExtension($OutPath) + '_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null

try {
    # Copy to .zip since Expand-Archive requires .zip extension
    $zipPath = Join-Path $tmp 'payload.zip'
    Copy-Item -LiteralPath $DocxPath -Destination $zipPath -Force
    Expand-Archive -LiteralPath $zipPath -DestinationPath $tmp -Force
    $docXmlPath = Join-Path $tmp 'word\document.xml'
    if (!(Test-Path -LiteralPath $docXmlPath)) { throw "document.xml not found in DOCX" }

    [xml]$xml = Get-Content -LiteralPath $docXmlPath -Raw
    $nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsm.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $paras = $xml.SelectNodes('//w:body/w:p',$nsm)

    $out = New-Object System.Collections.Generic.List[string]
    foreach($p in $paras){
        $runs = $p.SelectNodes('.//w:t',$nsm)
        $text = ($runs | ForEach-Object { $_.'#text' }) -join ''
        if ($text -ne $null) { $out.Add($text) } else { $out.Add('') }
    }
    $sep = [Environment]::NewLine + [Environment]::NewLine
    $txt = [string]::Join($sep, $out)
    Set-Content -LiteralPath $OutPath -Value $txt -Encoding UTF8
}
finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item $tmp -Recurse -Force }
}

Write-Host "OK"
