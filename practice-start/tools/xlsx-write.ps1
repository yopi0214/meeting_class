# xlsx-write.ps1 -- JSON spec -> .xlsx (Excel). Needs no Office, Python or Node. Keep this file ASCII-only.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools/xlsx-write.ps1 <spec.json> <out.xlsx> [-Force]
#
# { "sheets": [ { "name": "rep1",
#     "noteRows": [1],        # 1-based rows shown as small grey italic notes (optional)
#     "headerRow": 2,         # 1-based row shown bold with a fill (optional)
#     "freeze": true,         # freeze the rows up to headerRow (optional)
#     "widths": [12, 10],     # column widths (optional; guessed from content when omitted)
#     "rows": [ ["note text"], ["sample","target","Ct"], ["WT","TP53",24.1] ] } ] }
# JSON numbers become number cells, strings become text cells, null leaves a cell empty.
# Refuses to overwrite an existing file unless -Force is given.
param(
  [Parameter(Mandatory = $true)][string]$Spec,
  [Parameter(Mandatory = $true)][string]$Out,
  [switch]$Force
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Out = [System.IO.Path]::GetFullPath($Out)
if ((Test-Path -LiteralPath $Out) -and -not $Force) { throw "Output exists (use -Force to overwrite): $Out" }
$dir = Split-Path -Parent $Out
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$json = Get-Content -Raw -Encoding UTF8 -LiteralPath $Spec | ConvertFrom-Json
if ($null -eq $json.sheets) { throw 'Spec needs a "sheets" array.' }

function X([object]$s) { if ($null -eq $s) { return '' }; return [System.Security.SecurityElement]::Escape([string]$s) }
function ColName([int]$i) { $s = ''; $i++; while ($i -gt 0) { $m = ($i - 1) % 26; $s = [string][char](65 + $m) + $s; $i = [int][Math]::Floor(($i - 1) / 26) }; return $s }
function DispLen([string]$s) { $n = 0; foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 0x2E80) { $n += 2 } else { $n += 1 } }; return $n }
function IsNum($v) { return ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal] -or $v -is [single]) }

$ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
$sheetXmls = @(); $sheetNames = @()
foreach ($sh in @($json.sheets)) {
  $name = [string]$sh.name
  if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
  $name = $name -replace '[\\/\?\*\[\]:]', '_'
  $sheetNames += $name
  $rows = @($sh.rows)
  $noteRows = @(); if ($null -ne $sh.noteRows) { $noteRows = @($sh.noteRows | ForEach-Object { [int]$_ }) }
  $headerRow = 0; if ($null -ne $sh.headerRow) { $headerRow = [int]$sh.headerRow }
  $ncol = 0
  foreach ($r in $rows) { $rr = @($r); if ($rr.Count -gt $ncol) { $ncol = $rr.Count } }
  # column widths: ignore note rows when guessing
  $wd = New-Object 'double[]' $ncol
  if ($null -ne $sh.widths -and @($sh.widths).Count -gt 0) {
    for ($c = 0; $c -lt $ncol; $c++) { if ($c -lt @($sh.widths).Count) { $wd[$c] = [double]@($sh.widths)[$c] } else { $wd[$c] = 12 } }
  } else {
    for ($c = 0; $c -lt $ncol; $c++) {
      $m = 6; $ri = 0
      foreach ($r in $rows) { $ri++; if ($noteRows -contains $ri) { continue }; $rr = @($r); if ($c -lt $rr.Count -and $null -ne $rr[$c]) { $m = [Math]::Max($m, (DispLen ([string]$rr[$c])) + 2) } }
      $wd[$c] = [Math]::Min(60, $m)
    }
  }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ' + $ns + '>')
  if ($sh.freeze -and $headerRow -gt 0) {
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="' + $headerRow + '" topLeftCell="A' + ($headerRow + 1) + '" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
  }
  if ($ncol -gt 0) {
    [void]$sb.Append('<cols>')
    for ($c = 0; $c -lt $ncol; $c++) { [void]$sb.Append('<col min="' + ($c + 1) + '" max="' + ($c + 1) + '" width="' + $wd[$c].ToString([Globalization.CultureInfo]::InvariantCulture) + '" customWidth="1"/>') }
    [void]$sb.Append('</cols>')
  }
  [void]$sb.Append('<sheetData>')
  $ri = 0
  foreach ($r in $rows) {
    $ri++
    $rr = @($r)
    [void]$sb.Append('<row r="' + $ri + '">')
    $style = 0; if ($ri -eq $headerRow) { $style = 1 } elseif ($noteRows -contains $ri) { $style = 2 }
    for ($c = 0; $c -lt $rr.Count; $c++) {
      $v = $rr[$c]
      if ($null -eq $v) { continue }
      $ref = (ColName $c) + $ri
      $sattr = ''; if ($style -gt 0) { $sattr = ' s="' + $style + '"' }
      if (IsNum $v) {
        [void]$sb.Append('<c r="' + $ref + '"' + $sattr + '><v>' + ([string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0}', $v)) + '</v></c>')
      } elseif ($v -is [bool]) {
        [void]$sb.Append('<c r="' + $ref + '" t="b"' + $sattr + '><v>' + $(if ($v) { 1 } else { 0 }) + '</v></c>')
      } else {
        [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"' + $sattr + '><is><t xml:space="preserve">' + (X $v) + '</t></is></c>')
      }
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData></worksheet>')
  $sheetXmls += $sb.ToString()
}

$font = '<name val="Malgun Gothic"/>'
$stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ' + $ns + '>' +
  '<fonts count="3"><font><sz val="10"/>' + $font + '</font><font><b/><sz val="10"/>' + $font + '</font><font><i/><sz val="9"/><color rgb="FF666666"/>' + $font + '</font></fonts>' +
  '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

$n = $sheetNames.Count
$wbSheets = ''; $wbRels = ''; $ctSheets = ''
for ($i = 1; $i -le $n; $i++) {
  $wbSheets += '<sheet name="' + (X $sheetNames[$i - 1]) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>'
  $wbRels += '<Relationship Id="rId' + $i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $i + '.xml"/>'
  $ctSheets += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
}
$wbRels += '<Relationship Id="rId' + ($n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
$workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ' + $ns + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbSheets + '</sheets></workbook>'
$wbRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $wbRels + '</Relationships>'
$contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + $ctSheets +
  '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
$rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

$utf8 = New-Object System.Text.UTF8Encoding($false)
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
$zip = [System.IO.Compression.ZipFile]::Open($Out, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  $entries = @(@('[Content_Types].xml', $contentTypes), @('_rels/.rels', $rootRels), @('xl/workbook.xml', $workbookXml), @('xl/_rels/workbook.xml.rels', $wbRelsXml), @('xl/styles.xml', $stylesXml))
  for ($i = 1; $i -le $n; $i++) { $entries += , @(('xl/worksheets/sheet' + $i + '.xml'), $sheetXmls[$i - 1]) }
  foreach ($e in $entries) {
    $entry = $zip.CreateEntry($e[0])
    $sw = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
    $sw.Write($e[1]); $sw.Dispose()
  }
} finally { $zip.Dispose() }
Write-Output ('OK: ' + $Out + ' (' + $n + ' sheets)')
