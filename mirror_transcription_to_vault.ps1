param(
  [string]$Date = 'today',   # 'today' | 'all' | YYYYMMDD | YYYY-MM-DD
  [switch]$All,
  [string]$NotebookName = 'FarEasternTribe',
  [string]$SectionName  = '手書き認識テスト',
  [string]$VaultLogDir  = ''
)
$ErrorActionPreference = 'Stop'
# 端末非依存: 未指定なら スクリプト位置(…\Claude\OpenAI-Agent) から …\外部脳\日常ログ を導出(Desktop/Lenovo共通)
if([string]::IsNullOrWhiteSpace($VaultLogDir)){ $VaultLogDir=[System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\外部脳\日常ログ')) }
$ONE_NS = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
if ($All) { $Date = 'all' }
$targetKey = ''
if ($Date -notin @('today','all')) { $targetKey = ($Date -replace '[-/]','') }
elseif ($Date -eq 'today') { $targetKey = (Get-Date).ToString('yyyyMMdd') }

New-Item -ItemType Directory -Force -Path $VaultLogDir | Out-Null
$wjp = @('月','火','水','木','金','土','日')
function Weekday([string]$k){ try { $d=[datetime]::ParseExact($k,'yyyyMMdd',$null); $dow=[int]$d.DayOfWeek; return $wjp[($dow+6)%7] } catch { return '?' } }
function Ymd([string]$k){ "$($k.Substring(0,4))-$($k.Substring(4,2))-$($k.Substring(6,2))" }
function CleanText([string]$title,[string]$text){
  $lines = ($text -replace "`r`n","`n").Split("`n")
  $i=0
  while($i -lt $lines.Count -and $lines[$i].Trim() -eq $title.Trim()){ $i++ }
  return (($lines[$i..($lines.Count-1)]) -join "`n").Trim()
}

$one=New-Object -ComObject OneNote.Application
[string]$h=''; $one.GetHierarchy('',4,[ref]$h); [xml]$hi=$h
$ns=New-Object System.Xml.XmlNamespaceManager($hi.NameTable); $ns.AddNamespace('one',$ONE_NS)
$nb=@($hi.SelectNodes('//one:Notebook',$ns))|?{$_.name -eq $NotebookName}|Select -First 1
if(-not $nb){ throw "Notebook not found: $NotebookName" }
$sec=@($nb.SelectNodes('.//one:Section',$ns))|?{$_.name -eq $SectionName}|Select -First 1
if(-not $sec){ throw "Section not found: $SectionName" }
$pages=@($sec.SelectNodes('./one:Page',$ns))

# Collect page text; group by leading YYYYMMDD; keep max-chars (fullest) per date.
$best=@{}   # dateKey -> [pscustomobject]{title,lm,text,chars}
$weekly=@()
foreach($p in $pages){
  $t=[string]$p.name
  if($t -match '週次'){ $weeklyFlag=$true } else { $weeklyFlag=$false }
  $m=[regex]::Match($t,'^(\d{8})')
  if(-not $m.Success){ continue }
  $dk=$m.Groups[1].Value
  if($dk.StartsWith('2027')){ continue }   # year typo
  [string]$px=''
  try{ $one.GetPageContent($p.ID,[ref]$px) }catch{ $px='' }
  if(-not $px){ continue }
  [xml]$pd=$px; $pns=New-Object System.Xml.XmlNamespaceManager($pd.NameTable); $pns.AddNamespace('one',$ONE_NS)
  $txt=((@($pd.SelectNodes('//one:T',$pns))|%{ [System.Net.WebUtility]::HtmlDecode([string]$_.InnerText) }) -join "`n")
  $chars=$txt.Length
  $obj=[pscustomobject]@{ title=$t; lm=[string]$p.lastModifiedTime; text=$txt; chars=$chars; dateKey=$dk }
  if($weeklyFlag){ $weekly += $obj; continue }
  if($chars -lt 50){ continue }
  if(-not $best.ContainsKey($dk) -or $chars -gt $best[$dk].chars){ $best[$dk]=$obj }
}

$targets=@()
if($Date -eq 'all'){ $targets=@($best.Keys) }
elseif($best.ContainsKey($targetKey)){ $targets=@($targetKey) }
else { Write-Output "NO-PAGE for $targetKey (最新の文字起こしがまだOneNoteに無い可能性)"; }

$written=@()
foreach($dk in ($targets|Sort-Object)){
  $o=$best[$dk]; $ymd=Ymd $dk; $wd=Weekday $dk
  $body=CleanText $o.title $o.text
  $fm=@"
---
date: $ymd
曜日: $wd
type: 日常ログ
source: OneNote / $NotebookName / $SectionName
source_page: "$($o.title)"
source_last_modified: $($o.lm)
tags: [日常ログ, 手書き文字起こし]
---

"@
  $path=Join-Path $VaultLogDir "$ymd.md"
  [System.IO.File]::WriteAllText($path, ($fm + $body + "`n"), [System.Text.UTF8Encoding]::new($false))
  $written += "$ymd.md ($($body.Length)字)"
}

# weekly (only on -All)
if($Date -eq 'all'){
  foreach($o in $weekly){
    $mm=[regex]::Match($o.title,'^(\d{8})-(\d{8})')
    if($mm.Success){ $label="$(Ymd $mm.Groups[1].Value)〜$($mm.Groups[2].Value.Substring(4,2))-$($mm.Groups[2].Value.Substring(6,2))" } else { $label=$o.title }
    $body=CleanText $o.title $o.text
    $fm=@"
---
type: 週次まとめ
source_page: "$($o.title)"
tags: [週次まとめ, Todo, 手書き文字起こし]
---

"@
    $path=Join-Path $VaultLogDir "週次_$label.md"
    [System.IO.File]::WriteAllText($path, ($fm + $body + "`n"), [System.Text.UTF8Encoding]::new($false))
    $written += "週次_$label.md"
  }
}

# Rebuild MOC from files present
$files=Get-ChildItem -LiteralPath $VaultLogDir -Filter *.md | ?{ $_.Name -ne '_日常ログMOC.md' }
$daily=$files|?{ $_.Name -match '^\d{4}-\d{2}-\d{2}\.md$' }|Sort-Object Name
$wk=$files|?{ $_.Name -like '週次_*' }|Sort-Object Name
$moc=@('# 日常ログ MOC（手書きノート文字起こし）','',
  "OneNote ``$NotebookName / $SectionName`` から取り込んだ日々の生活・日誌ログ。",'','## 日次','')
foreach($f in $daily){ $nm=$f.BaseName; $moc+="- [[$nm]]" }
$moc+=@('','## 週次まとめ','')
foreach($f in $wk){ $nm=$f.BaseName; $moc+="- [[$nm]]" }
$moc+=''
[System.IO.File]::WriteAllText((Join-Path $VaultLogDir '_日常ログMOC.md'), ($moc -join "`n"), [System.Text.UTF8Encoding]::new($false))

"MIRRORED: " + ($(if($written.Count){$written -join ', '}else{'(none)'}))
"MOC updated: $($daily.Count) daily, $($wk.Count) weekly"
