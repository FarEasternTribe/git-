param(
  [string]$NotebookName  = 'FarEasternTribe',
  [string]$SectionName   = '論文要約',
  [string]$LibraryDir    = 'C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent\OpenAI-Agent\papers\library',
  [switch]$Apply
)
$ErrorActionPreference = 'Stop'
$ONE_NS = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$LibraryPdfDir = Join-Path $LibraryDir 'PDFs'
$IndexDir      = Join-Path $LibraryDir 'index'

function Norm([string]$s){
  if(-not $s){ return '' }
  $s = $s.ToLowerInvariant()
  # drop a trailing parenthetical annotation group like " (jacs 2026, ...)"
  $s = [regex]::Replace($s, '[\(（].*$', '')
  # keep only ascii alphanumerics for a stable comparison key
  return ([regex]::Replace($s, '[^a-z0-9]', ''))
}

$one = New-Object -ComObject OneNote.Application
[string]$h=''; $one.GetHierarchy('',4,[ref]$h); [xml]$hi=$h
$ns=New-Object System.Xml.XmlNamespaceManager($hi.NameTable); $ns.AddNamespace('one',$ONE_NS)
$nb=@($hi.SelectNodes('//one:Notebook',$ns))|?{$_.name -eq $NotebookName}|Select -First 1
if(-not $nb){ throw "Notebook not found: $NotebookName" }
$sec=@($nb.SelectNodes('.//one:Section',$ns))|?{$_.name -eq $SectionName}|Select -First 1
if(-not $sec){ throw "Section not found: $SectionName" }
[string]$sx=''; $one.GetHierarchy($sec.ID,4,[ref]$sx); [xml]$sd=$sx
$pages=@($sd.SelectNodes('//one:Page',$ns))

# library filename -> fullpath (fallback for stale Local PDF paths)
$libMap=@{}
if(Test-Path -LiteralPath $LibraryPdfDir){
  Get-ChildItem -LiteralPath $LibraryPdfDir -Filter *.pdf | %{ $libMap[$_.Name.ToLowerInvariant()] = $_.FullName }
}
# normalized index Title -> pdf fullpath (fallback for pages without a "Local PDF:" line)
$idxMap=@{}
if(Test-Path -LiteralPath $IndexDir){
  foreach($f in Get-ChildItem -LiteralPath $IndexDir -Filter *.md){
    if($f.Name -like 'MASTER_INDEX*'){ continue }
    $txt=[System.IO.File]::ReadAllText($f.FullName,[System.Text.UTF8Encoding]::new($false))
    $tm=[regex]::Match($txt,'(?m)^-\s*Title:\s*(.+?)\s*$')
    $pm=[regex]::Match($txt,'(?m)^-\s*PDF:\s*(.+?)\s*$')
    if($tm.Success -and $pm.Success){
      $key=Norm $tm.Groups[1].Value
      $pdf=Join-Path $LibraryDir ($pm.Groups[1].Value.Trim() -replace '/','\')
      if($key.Length -ge 12 -and (Test-Path -LiteralPath $pdf)){ $idxMap[$key]=$pdf }
    }
  }
}

function Resolve-Pdf($page,[string]$pageXml){
  # 1) explicit "Local PDF:" line in the note body
  [xml]$pd=$pageXml
  $pns=New-Object System.Xml.XmlNamespaceManager($pd.NameTable); $pns.AddNamespace('one',$ONE_NS)
  foreach($t in @($pd.SelectNodes('//one:T',$pns))){
    $txt=[string]$t.InnerText
    if($txt -match '^\s*Local PDF:\s*(.+?)\s*$'){
      $p=$Matches[1].Trim()
      if(Test-Path -LiteralPath $p){ return $p }
      $bn=Split-Path $p -Leaf
      if($bn -and $libMap.ContainsKey($bn.ToLowerInvariant())){ return $libMap[$bn.ToLowerInvariant()] }
      return ''  # had a path but unresolved
    }
  }
  # 2) fallback: match page title against the library index Title
  $pk=Norm $page.name
  if($pk.Length -ge 12){
    if($idxMap.ContainsKey($pk)){ return $idxMap[$pk] }
    $hits=@($idxMap.Keys | ?{ $pk.StartsWith($_) -or $_.StartsWith($pk) })
    if($hits.Count -eq 1){ return $idxMap[$hits[0]] }
  }
  return ''
}

$attached=0; $already=0; $skipped=0; $rows=@()
foreach($p in $pages){
  [string]$px=''
  try{ $one.GetPageContent($p.ID,[ref]$px) }catch{ $px='' }
  if(-not $px){ $rows+="SKIP(no-content)   '$($p.name)'"; $skipped++; continue }
  [xml]$pd=$px
  $pns=New-Object System.Xml.XmlNamespaceManager($pd.NameTable); $pns.AddNamespace('one',$ONE_NS)
  if(@($pd.SelectNodes('//one:InsertedFile',$pns)).Count -gt 0){ $rows+="ALREADY            '$($p.name)'"; $already++; continue }

  $pdfPath = Resolve-Pdf $p $px
  if(-not $pdfPath){ $rows+="SKIP(no-pdf-match) '$($p.name)'"; $skipped++; continue }

  if(-not $Apply){ $rows+="WILL-ATTACH        '$($p.name)'  -> $(Split-Path $pdfPath -Leaf)"; $attached++; continue }

  $outline=@($pd.SelectNodes('//one:Outline',$pns))|Select -Last 1
  if(-not $outline){ $outline=$pd.CreateElement('one','Outline',$ONE_NS); [void]$pd.DocumentElement.AppendChild($outline) }
  $oec=@($outline.SelectNodes('./one:OEChildren',$pns))|Select -Last 1
  if(-not $oec){ $oec=$pd.CreateElement('one','OEChildren',$ONE_NS); [void]$outline.AppendChild($oec) }
  $oe=$pd.CreateElement('one','OE',$ONE_NS)
  $ins=$pd.CreateElement('one','InsertedFile',$ONE_NS)
  $ins.SetAttribute('pathCache',$pdfPath); $ins.SetAttribute('preferredName',(Split-Path $pdfPath -Leaf))
  [void]$oe.AppendChild($ins); [void]$oec.AppendChild($oe)
  try{
    $one.UpdatePageContent($pd.OuterXml); Start-Sleep -Milliseconds 400
    [string]$rd=''; $one.GetPageContent($p.ID,[ref]$rd)
    if($rd -match 'InsertedFile'){ $rows+="ATTACHED           '$($p.name)'  -> $(Split-Path $pdfPath -Leaf)"; $attached++ }
    else{ $rows+="FAIL(verify)       '$($p.name)'"; $skipped++ }
  }catch{ $rows+="FAIL(update)       '$($p.name)' : $($_.Exception.Message)"; $skipped++ }
}

"MODE: " + ($(if($Apply){'APPLY'}else{'DRY-RUN'}))
"Section: $NotebookName / $SectionName  pages=$($pages.Count)  indexTitles=$($idxMap.Count)"
$rows | %{ $_ }
""
"SUMMARY: attach=$attached already=$already skipped=$skipped total=$($pages.Count)"
