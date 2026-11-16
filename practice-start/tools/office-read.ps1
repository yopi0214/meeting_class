# office-read.ps1 -- print the text of .docx / .pptx / .xlsx as plain text. Needs no Office, Python or Node. Keep this file ASCII-only.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools/office-read.ps1 <file> [-MaxRows 300]
#
# docx : headings as "#", bullets as "-", tables as "| a | b |" rows, Word comments at the end
# pptx : "=== Slide n ===", TITLE, body lines, tables, [picture: name], [chart], [notes] text
# xlsx : "=== Sheet: name ===", rows as "R<row>: a | b | c" (dates shown as yyyy-MM-dd)
# Images and charts are only reported by name; their content is not read.
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [int]$MaxRows = 300
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Path = (Resolve-Path -LiteralPath $Path).Path
$ext = [System.IO.Path]::GetExtension($Path).ToLower()
if ($ext -in '.doc', '.xls', '.ppt') { Write-Output ('UNSUPPORTED: old binary format ' + $ext + ' - save it as docx/xlsx/pptx first'); exit 2 }
if ($ext -notin '.docx', '.pptx', '.xlsx') { Write-Output ('UNSUPPORTED: ' + $ext); exit 2 }

$zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
$out = New-Object System.Collections.Generic.List[string]

function LoadXml([string]$name) {
  $e = $zip.GetEntry($name)
  if ($null -eq $e) { return $null }
  $st = $e.Open()
  $s = New-Object System.Xml.XmlReaderSettings
  $s.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
  $s.XmlResolver = $null
  $rd = [System.Xml.XmlReader]::Create($st, $s)
  $d = New-Object System.Xml.XmlDocument
  $d.XmlResolver = $null
  $d.Load($rd); $rd.Close(); $st.Close()
  return $d
}
function NsMgr($doc) {
  $m = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $m.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $m.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
  $m.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
  $m.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $m.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
  $m.AddNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart')
  $m.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
  return , $m
}
function ResolveRel([string]$baseDir, [string]$target) {
  if ($target.StartsWith('/')) { return $target.TrimStart('/') }
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($seg in (($baseDir + '/' + $target) -split '/')) {
    if ($seg -eq '' -or $seg -eq '.') { continue }
    if ($seg -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }; continue }
    $parts.Add($seg)
  }
  return ($parts -join '/')
}
function RelMap([string]$relPath) {
  $map = @{}
  $d = LoadXml $relPath
  if ($null -eq $d) { return $map }
  $m = NsMgr $d
  foreach ($r in $d.SelectNodes('//rel:Relationship', $m)) { $map[$r.GetAttribute('Id')] = @{ Target = $r.GetAttribute('Target'); Type = $r.GetAttribute('Type') } }
  return $map
}

# ---------------- docx ----------------
function WText($node, $ns) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($t in $node.SelectNodes('.//w:t | .//w:tab | .//w:br', $ns)) {
    switch ($t.LocalName) { 't' { [void]$sb.Append($t.InnerText) } 'tab' { [void]$sb.Append("`t") } 'br' { [void]$sb.Append(' ') } }
  }
  return $sb.ToString()
}
function WalkBody($parent, $ns) {
  foreach ($n in $parent.ChildNodes) {
    switch ($n.LocalName) {
      'p' {
        $txt = (WText $n $ns)
        if ($txt.Trim() -eq '') { continue }
        $sty = $n.SelectSingleNode('w:pPr/w:pStyle/@w:val', $ns)
        $num = $n.SelectSingleNode('w:pPr/w:numPr', $ns)
        $prefix = ''
        if ($sty -and $sty.Value -match '^(?i)heading\s*(\d)') { $prefix = ('#' * [int]$Matches[1]) + ' ' }
        elseif ($sty -and $sty.Value -match '^(?i)title') { $prefix = '# ' }
        elseif ($num) {
          $lv = $n.SelectSingleNode('w:pPr/w:numPr/w:ilvl/@w:val', $ns)
          $indent = ''; if ($lv) { $indent = '  ' * [int]$lv.Value }
          $prefix = $indent + '- '
        }
        $out.Add($prefix + $txt.Trim())
      }
      'tbl' {
        foreach ($tr in $n.SelectNodes('w:tr', $ns)) {
          $cells = @()
          foreach ($tc in $tr.SelectNodes('w:tc', $ns)) {
            $parts = @()
            foreach ($cp in $tc.SelectNodes('.//w:p', $ns)) { $tx = (WText $cp $ns).Trim(); if ($tx) { $parts += $tx } }
            $cells += ($parts -join ' / ')
          }
          $out.Add('| ' + ($cells -join ' | ') + ' |')
        }
        $out.Add('')
      }
      'sdt' { $c = $n.SelectSingleNode('w:sdtContent', $ns); if ($c) { WalkBody $c $ns } }
    }
  }
}
function ReadDocx {
  $d = LoadXml 'word/document.xml'
  if ($null -eq $d) { throw 'word/document.xml not found' }
  $ns = NsMgr $d
  $body = $d.SelectSingleNode('//w:body', $ns)
  WalkBody $body $ns
  $cm = LoadXml 'word/comments.xml'
  if ($cm) {
    $cns = NsMgr $cm
    $list = $cm.SelectNodes('//w:comment', $cns)
    if ($list.Count -gt 0) {
      $out.Add(''); $out.Add('=== comments ===')
      foreach ($c in $list) { $out.Add('[comment by ' + $c.GetAttribute('author', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') + '] ' + (WText $c $cns).Trim()) }
    }
  }
}

# ---------------- pptx ----------------
function SpLines($sp, $ns) {
  $ph = $sp.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $ns)
  $phType = ''
  if ($ph) { $phType = $ph.GetAttribute('type') }
  if ($phType -in 'sldNum', 'dt', 'ftr') { return }
  foreach ($para in $sp.SelectNodes('p:txBody/a:p', $ns)) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $para.SelectNodes('a:r/a:t | a:br | a:fld/a:t', $ns)) {
      if ($t.LocalName -eq 'br') { [void]$sb.Append(' ') } else { [void]$sb.Append($t.InnerText) }
    }
    $tx = $sb.ToString().Trim()
    if ($tx -eq '') { continue }
    if ($phType -in 'title', 'ctrTitle') { $script:slideOut.Add('TITLE: ' + $tx) }
    else {
      $lv = $para.SelectSingleNode('a:pPr/@lvl', $ns)
      $indent = ''; if ($lv) { $indent = '  ' * [int]$lv.Value }
      $script:slideOut.Add($indent + $tx)
    }
  }
}
function WalkShapes($tree, $ns) {
  foreach ($n in $tree.ChildNodes) {
    switch ($n.LocalName) {
      'sp' { SpLines $n $ns }
      'grpSp' { WalkShapes $n $ns }
      'pic' {
        $nv = $n.SelectSingleNode('p:nvPicPr/p:cNvPr', $ns)
        $script:slideOut.Add('[picture: ' + $nv.GetAttribute('name') + $(if ($nv.GetAttribute('descr')) { ' / alt: ' + $nv.GetAttribute('descr') } else { '' }) + ']')
      }
      'graphicFrame' {
        $tbl = $n.SelectSingleNode('.//a:tbl', $ns)
        if ($tbl) {
          foreach ($tr in $tbl.SelectNodes('a:tr', $ns)) {
            $cells = @()
            foreach ($tc in $tr.SelectNodes('a:tc', $ns)) {
              $ps = @(); foreach ($pp in $tc.SelectNodes('.//a:p', $ns)) { $tx = (($pp.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim(); if ($tx) { $ps += $tx } }
              $cells += ($ps -join ' / ')
            }
            $script:slideOut.Add('| ' + ($cells -join ' | ') + ' |')
          }
        }
        elseif ($n.SelectSingleNode('.//c:chart', $ns)) { $script:slideOut.Add('[chart]') }
      }
    }
  }
}
function ReadPptx {
  $pres = LoadXml 'ppt/presentation.xml'
  $pns = NsMgr $pres
  $rels = RelMap 'ppt/_rels/presentation.xml.rels'
  $i = 0
  foreach ($sid in $pres.SelectNodes('//p:sldIdLst/p:sldId', $pns)) {
    $i++
    $rid = $sid.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $slidePath = ResolveRel 'ppt' $rels[$rid].Target
    $sd = LoadXml $slidePath
    $ns = NsMgr $sd
    $script:slideOut = New-Object System.Collections.Generic.List[string]
    WalkShapes ($sd.SelectSingleNode('//p:cSld/p:spTree', $ns)) $ns
    $out.Add('=== Slide ' + $i + ' ===')
    foreach ($l in $script:slideOut) { $out.Add($l) }
    # notes
    $slideRels = RelMap ((Split-Path $slidePath -Parent).Replace('\', '/') + '/_rels/' + (Split-Path $slidePath -Leaf) + '.rels')
    foreach ($k in $slideRels.Keys) {
      if ($slideRels[$k].Type -like '*/notesSlide') {
        $np = ResolveRel ((Split-Path $slidePath -Parent).Replace('\', '/')) $slideRels[$k].Target
        $nd = LoadXml $np
        if ($nd) {
          $nns = NsMgr $nd
          $lines = @()
          foreach ($sp in $nd.SelectNodes('//p:sp', $nns)) {
            $ph = $sp.SelectSingleNode('p:nvSpPr/p:nvPr/p:ph', $nns)
            if ($ph -and $ph.GetAttribute('type') -eq 'body') {
              foreach ($para in $sp.SelectNodes('p:txBody/a:p', $nns)) { $tx = (($para.SelectNodes('.//a:t', $nns) | ForEach-Object { $_.InnerText }) -join '').Trim(); if ($tx) { $lines += $tx } }
            }
          }
          if ($lines.Count -gt 0) { $out.Add('[notes] ' + ($lines -join ' / ')) }
        }
      }
    }
    $out.Add('')
  }
}

# ---------------- xlsx ----------------
function ColIndex([string]$ref) {
  $letters = ($ref -replace '[0-9]', '').ToUpper(); $n = 0
  foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
  return $n - 1
}
function ReadXlsx {
  $wb = LoadXml 'xl/workbook.xml'
  $wns = NsMgr $wb
  $rels = RelMap 'xl/_rels/workbook.xml.rels'
  $shared = New-Object System.Collections.Generic.List[string]
  $ss = LoadXml 'xl/sharedStrings.xml'
  if ($ss) { $sns = NsMgr $ss; foreach ($si in $ss.SelectNodes('//s:si', $sns)) { $shared.Add((($si.SelectNodes('.//s:t', $sns) | ForEach-Object { $_.InnerText }) -join '')) } }
  $dateXf = @{}
  $st = LoadXml 'xl/styles.xml'
  if ($st) {
    $tns = NsMgr $st
    $customDate = @{}
    foreach ($nf in $st.SelectNodes('//s:numFmts/s:numFmt', $tns)) { if ($nf.GetAttribute('formatCode') -match '[dmyDMY]' -and $nf.GetAttribute('formatCode') -notmatch '^\[?[Gg]eneral') { $customDate[$nf.GetAttribute('numFmtId')] = $true } }
    $xi = 0
    foreach ($xf in $st.SelectNodes('//s:cellXfs/s:xf', $tns)) {
      $id = $xf.GetAttribute('numFmtId')
      if (($id -match '^\d+$' -and (([int]$id -ge 14 -and [int]$id -le 22) -or ([int]$id -ge 45 -and [int]$id -le 47))) -or $customDate.ContainsKey($id)) { $dateXf[$xi] = $true }
      $xi++
    }
  }
  foreach ($sh in $wb.SelectNodes('//s:sheets/s:sheet', $wns)) {
    $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $hidden = ''; if ($sh.GetAttribute('state') -in 'hidden', 'veryHidden') { $hidden = ' (hidden)' }
    $out.Add('=== Sheet: ' + $sh.GetAttribute('name') + $hidden + ' ===')
    $sd = LoadXml (ResolveRel 'xl' $rels[$rid].Target)
    if ($null -eq $sd) { $out.Add('(unreadable)'); continue }
    $ns = NsMgr $sd
    $shown = 0; $total = 0
    foreach ($row in $sd.SelectNodes('//s:sheetData/s:row', $ns)) {
      $total++
      if ($shown -ge $MaxRows) { continue }
      $vals = New-Object System.Collections.Generic.List[string]
      foreach ($c in $row.SelectNodes('s:c', $ns)) {
        $idx = ColIndex $c.GetAttribute('r')
        while ($vals.Count -le $idx) { $vals.Add('') }
        $t = $c.GetAttribute('t'); $v = $c.SelectSingleNode('s:v', $ns); $val = ''
        if ($t -eq 's' -and $v) { $val = $shared[[int]$v.InnerText] }
        elseif ($t -eq 'inlineStr') { $val = (($c.SelectNodes('.//s:t', $ns) | ForEach-Object { $_.InnerText }) -join '') }
        elseif ($t -eq 'b' -and $v) { $val = $(if ($v.InnerText -eq '1') { 'TRUE' } else { 'FALSE' }) }
        elseif ($v) {
          $val = $v.InnerText
          $sidx = $c.GetAttribute('s')
          if ($sidx -match '^\d+$' -and $dateXf.ContainsKey([int]$sidx) -and $val -match '^[0-9.]+$') { $val = ([datetime]'1899-12-30').AddDays([double]::Parse($val, [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd') }
        }
        else { $f = $c.SelectSingleNode('s:f', $ns); if ($f) { $val = '=' + $f.InnerText } }
        $vals[$idx] = ($val -replace "[\r\n]+", ' ')
      }
      $last = $vals.Count - 1
      while ($last -ge 0 -and $vals[$last] -eq '') { $last-- }
      if ($last -lt 0) { continue }
      $out.Add('R' + $row.GetAttribute('r') + ': ' + (($vals.GetRange(0, $last + 1)) -join ' | '))
      $shown++
    }
    if ($total -gt $shown) { $out.Add('(... ' + ($total - $shown) + ' more rows not shown; use -MaxRows)') }
    $out.Add('')
  }
}

try {
  switch ($ext) { '.docx' { ReadDocx } '.pptx' { ReadPptx } '.xlsx' { ReadXlsx } }
} finally { $zip.Dispose() }
$out | ForEach-Object { Write-Output $_ }
