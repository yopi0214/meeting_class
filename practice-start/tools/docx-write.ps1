# docx-write.ps1 -- JSON spec -> .docx (Word). Needs no Office, Python or Node. Keep this file ASCII-only.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools/docx-write.ps1 <spec.json> <out.docx> [-Force] [-CountPages]
#
# The spec is a UTF-8 JSON file (write Korean text in the JSON, never in the command line):
# { "blocks": [
#   {"type":"title","text":"..."},
#   {"type":"h1","text":"..."}, {"type":"h2","text":"..."}, {"type":"h3","text":"..."},
#   {"type":"p","text":"plain text with **bold** parts"},
#   {"type":"note","text":"small grey italic text"},
#   {"type":"bullets","items":["a", {"t":"sub item","l":1}]},
#   {"type":"numbered","items":["first","second"]},
#   {"type":"table","header":["col1","col2"],"rows":[["a","b"],["c","d"]],"widths":[2,1]},
#   {"type":"spacer"}
# ] }
# Refuses to overwrite an existing file unless -Force is given. -CountPages opens the result in Word (if installed) and prints "pages: N".
param(
  [Parameter(Mandatory = $true)][string]$Spec,
  [Parameter(Mandatory = $true)][string]$Out,
  [switch]$Force,
  [switch]$CountPages
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Out = [System.IO.Path]::GetFullPath($Out)
if ((Test-Path -LiteralPath $Out) -and -not $Force) { throw "Output exists (use -Force to overwrite): $Out" }
$dir = Split-Path -Parent $Out
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$json = Get-Content -Raw -Encoding UTF8 -LiteralPath $Spec | ConvertFrom-Json
if ($null -eq $json.blocks) { throw 'Spec needs a "blocks" array.' }

function X([object]$s) { if ($null -eq $s) { return '' }; return [System.Security.SecurityElement]::Escape([string]$s) }

function Runs([string]$text, [bool]$bold, [bool]$italic, [string]$color, [int]$size) {
  $sb = New-Object System.Text.StringBuilder
  $parts = $text -split '\*\*'
  for ($i = 0; $i -lt $parts.Count; $i++) {
    if ($parts[$i] -eq '') { continue }
    $b = $bold -or (($i % 2) -eq 1)
    $rpr = ''
    if ($b) { $rpr += '<w:b/>' }
    if ($italic) { $rpr += '<w:i/>' }
    if ($color) { $rpr += '<w:color w:val="' + $color + '"/>' }
    if ($size -gt 0) { $rpr += '<w:sz w:val="' + $size + '"/><w:szCs w:val="' + $size + '"/>' }
    if ($rpr) { $rpr = '<w:rPr>' + $rpr + '</w:rPr>' }
    [void]$sb.Append('<w:r>' + $rpr + '<w:t xml:space="preserve">' + (X $parts[$i]) + '</w:t></w:r>')
  }
  return $sb.ToString()
}

function Para([string]$text, [string]$style, [bool]$bold, [bool]$italic, [string]$color, [int]$size, [int]$numId, [int]$lvl) {
  $ppr = ''
  if ($style) { $ppr += '<w:pStyle w:val="' + $style + '"/>' }
  if ($numId -gt 0) { $ppr += '<w:numPr><w:ilvl w:val="' + $lvl + '"/><w:numId w:val="' + $numId + '"/></w:numPr>' }
  if ($ppr) { $ppr = '<w:pPr>' + $ppr + '</w:pPr>' }
  return '<w:p>' + $ppr + (Runs $text $bold $italic $color $size) + '</w:p>'
}

function DispLen([string]$s) {
  $n = 0
  foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 0x2E80) { $n += 2 } else { $n += 1 } }
  return $n
}

function Tbl($header, $rows, $widths) {
  $hdr = @(); if ($null -ne $header) { $hdr = @($header) }
  $rws = @(); if ($null -ne $rows) { $rws = @($rows) }
  $ncol = $hdr.Count
  foreach ($r in $rws) { if (@($r).Count -gt $ncol) { $ncol = @($r).Count } }
  if ($ncol -eq 0) { return '' }
  $weights = New-Object 'double[]' $ncol
  if ($null -ne $widths -and @($widths).Count -eq $ncol) {
    for ($c = 0; $c -lt $ncol; $c++) { $weights[$c] = [double]@($widths)[$c] }
  } else {
    for ($c = 0; $c -lt $ncol; $c++) {
      $m = 0
      if ($c -lt $hdr.Count) { $m = [Math]::Max($m, (DispLen ([string]$hdr[$c]))) }
      foreach ($r in $rws) { $rr = @($r); if ($c -lt $rr.Count) { $m = [Math]::Max($m, (DispLen (([string]$rr[$c]) -replace '\*\*', ''))) } }
      $weights[$c] = [Math]::Min(34, [Math]::Max(8, $m))
    }
  }
  $sum = 0.0; foreach ($w in $weights) { $sum += $w }
  $total = 9638
  $tw = New-Object 'int[]' $ncol
  for ($c = 0; $c -lt $ncol; $c++) { $tw[$c] = [int][Math]::Floor($total * $weights[$c] / $sum) }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<w:tbl><w:tblPr><w:tblW w:w="' + $total + '" w:type="dxa"/><w:tblBorders>')
  foreach ($side in 'top', 'left', 'bottom', 'right', 'insideH', 'insideV') {
    [void]$sb.Append('<w:' + $side + ' w:val="single" w:sz="4" w:space="0" w:color="8EA0B8"/>')
  }
  [void]$sb.Append('</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>')
  for ($c = 0; $c -lt $ncol; $c++) { [void]$sb.Append('<w:gridCol w:w="' + $tw[$c] + '"/>') }
  [void]$sb.Append('</w:tblGrid>')
  $allRows = @()
  if ($hdr.Count -gt 0) { $allRows += , @{ cells = $hdr; head = $true } }
  foreach ($r in $rws) { $allRows += , @{ cells = @($r); head = $false } }
  foreach ($row in $allRows) {
    [void]$sb.Append('<w:tr>')
    if ($row.head) { [void]$sb.Append('<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>') } else { [void]$sb.Append('<w:trPr><w:cantSplit/></w:trPr>') }
    for ($c = 0; $c -lt $ncol; $c++) {
      $txt = ''
      if ($c -lt $row.cells.Count -and $null -ne $row.cells[$c]) { $txt = [string]$row.cells[$c] }
      [void]$sb.Append('<w:tc><w:tcPr><w:tcW w:w="' + $tw[$c] + '" w:type="dxa"/>')
      if ($row.head) { [void]$sb.Append('<w:shd w:val="clear" w:color="auto" w:fill="DCE6F2"/>') }
      [void]$sb.Append('</w:tcPr>')
      foreach ($line in ([regex]::Split($txt, '\r?\n'))) {
        [void]$sb.Append('<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>' + (Runs $line $row.head $false '' 19) + '</w:p>')
      }
      [void]$sb.Append('</w:tc>')
    }
    [void]$sb.Append('</w:tr>')
  }
  [void]$sb.Append('</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="60"/></w:pPr></w:p>')
  return $sb.ToString()
}

$body = New-Object System.Text.StringBuilder
$numberedCount = 0
foreach ($b in @($json.blocks)) {
  switch ([string]$b.type) {
    'title'    { [void]$body.Append((Para ([string]$b.text) 'Title' $false $false '' 0 0 0)) }
    'h1'       { [void]$body.Append((Para ([string]$b.text) 'Heading1' $false $false '' 0 0 0)) }
    'h2'       { [void]$body.Append((Para ([string]$b.text) 'Heading2' $false $false '' 0 0 0)) }
    'h3'       { [void]$body.Append((Para ([string]$b.text) 'Heading3' $false $false '' 0 0 0)) }
    'p'        { [void]$body.Append((Para ([string]$b.text) '' $false $false '' 0 0 0)) }
    'note'     { [void]$body.Append((Para ([string]$b.text) '' $false $true '666666' 19 0 0)) }
    'spacer'   { [void]$body.Append('<w:p/>') }
    'bullets'  {
      foreach ($it in @($b.items)) {
        $t = ''; $l = 0
        if ($it -is [string]) { $t = $it } else { $t = [string]$it.t; if ($null -ne $it.l) { $l = [int]$it.l } }
        [void]$body.Append((Para $t 'ListParagraph' $false $false '' 0 1 $l))
      }
    }
    'numbered' {
      $numberedCount++
      $nid = 1 + $numberedCount
      if ($nid -gt 21) { $nid = 21 }
      foreach ($it in @($b.items)) {
        $t = ''; $l = 0
        if ($it -is [string]) { $t = $it } else { $t = [string]$it.t; if ($null -ne $it.l) { $l = [int]$it.l } }
        [void]$body.Append((Para $t 'ListParagraph' $false $false '' 0 $nid $l))
      }
    }
    'table'    { [void]$body.Append((Tbl $b.header $b.rows $b.widths)) }
    default    { throw ('Unknown block type: ' + [string]$b.type) }
  }
}

$ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + $ns + '><w:body>' + $body.ToString() +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'

$fontAttr = 'w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic" w:cs="Malgun Gothic"'
$stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + $ns + '>' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts ' + $fontAttr + '/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="80" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="200"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="100"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="220" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="25"/><w:szCs w:val="25"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="404040"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="40"/><w:ind w:left="720"/></w:pPr></w:style>' +
  '</w:styles>'

$numSb = New-Object System.Text.StringBuilder
[void]$numSb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + $ns + '>')
[void]$numSb.Append('<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="500" w:hanging="260"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8211;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="900" w:hanging="260"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#183;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1300" w:hanging="260"/></w:pPr></w:lvl></w:abstractNum>')
[void]$numSb.Append('<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="500" w:hanging="340"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="900" w:hanging="340"/></w:pPr></w:lvl></w:abstractNum>')
[void]$numSb.Append('<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>')
for ($i = 2; $i -le 21; $i++) {
  [void]$numSb.Append('<w:num w:numId="' + $i + '"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>')
}
[void]$numSb.Append('</w:numbering>')

$contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
$rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
$docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'

$utf8 = New-Object System.Text.UTF8Encoding($false)
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
$zip = [System.IO.Compression.ZipFile]::Open($Out, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  foreach ($e in @(
      @('[Content_Types].xml', $contentTypes), @('_rels/.rels', $rootRels), @('word/document.xml', $documentXml),
      @('word/styles.xml', $stylesXml), @('word/numbering.xml', $numSb.ToString()), @('word/_rels/document.xml.rels', $docRels))) {
    $entry = $zip.CreateEntry($e[0])
    $sw = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
    $sw.Write($e[1]); $sw.Dispose()
  }
} finally { $zip.Dispose() }
Write-Output ('OK: ' + $Out + ' (' + @($json.blocks).Count + ' blocks)')
if ($CountPages) {
  try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false; $w.DisplayAlerts = 0
    try {
      $d = $w.Documents.Open($Out, $false, $true)
      $n = $d.ComputeStatistics(2)
      $d.Close($false)
      Write-Output ('pages: ' + $n)
    } finally { $w.Quit() }
  } catch { Write-Output 'pages: unknown (Word is not available)' }
}
