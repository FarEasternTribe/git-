param(
  [Parameter(Mandatory = $true)]
  [string]$Query
)

$ErrorActionPreference = 'Stop'

$terms = @(
  $Query,
  '有機合成',
  '合成',
  '合成ルート',
  '反応',
  '条件',
  '試薬',
  '文献',
  '論文',
  'DOI',
  'Supplementary',
  'Supporting Information',
  'SI',
  'ESI',
  'supplemental',
  '補足',
  'Supporting',
  '特許',
  'patent',
  'PDF',
  'reference',
  'literature',
  'route',
  '前駆体',
  'モノマー',
  'ルート',
  'synthesis',
  'reaction',
  'precursor',
  'monomer',
  'Suzuki',
  'Sonogashira',
  'Buchwald',
  'Ullmann',
  'Stille',
  'Negishi',
  'GNR',
  'グラフェン',
  'Br化',
  'ヨウ素',
  '臭素',
  'カップリング',
  '精製',
  '再結晶',
  'カラム',
  'NMR',
  'MS',
  'TLC'
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

Write-Host '有機合成Agent'
Write-Host '役割: 合成ルートを考える / 合成に必要な文献を探す / 必要な試薬・条件を検討する'
Write-Host '検索方針: 本文だけでなく、Supplementary Information / Supporting Information / ESI / SI / 特許 / PDF本文まで掘り下げる'
Write-Host '必須出力: ソースとなる文献のタイトル、URL、DOIを必ず添付する。DOIが見つからない場合は探索範囲と代替一次ソースを明示する'
Write-Host ('Query: {0}' -f $Query)
Write-Host ''

function Get-SectionPath($Section) {
  $parts = New-Object System.Collections.ArrayList
  $node = $Section
  while ($null -ne $node -and $node.LocalName -ne 'Notebook') {
    if (($node.LocalName -eq 'Section' -or $node.LocalName -eq 'SectionGroup') -and
        -not [string]::IsNullOrWhiteSpace($node.name)) {
      [void]$parts.Insert(0, [string]$node.name)
    }
    $node = $node.ParentNode
  }
  return ($parts -join ' / ')
}

function Get-Snippet([string]$Text, [string[]]$Terms) {
  $normalized = (($Text -replace '\s+', ' ').Trim())
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return ''
  }

  $lower = $normalized.ToLowerInvariant()
  $hitIndex = -1
  foreach ($term in $Terms) {
    if ([string]::IsNullOrWhiteSpace($term)) { continue }
    $idx = $lower.IndexOf($term.ToLowerInvariant())
    if ($idx -ge 0 -and ($hitIndex -lt 0 -or $idx -lt $hitIndex)) {
      $hitIndex = $idx
    }
  }

  if ($hitIndex -lt 0) {
    $hitIndex = 0
  }
  $start = [Math]::Max(0, $hitIndex - 80)
  $length = [Math]::Min(360, $normalized.Length - $start)
  $snippet = $normalized.Substring($start, $length)
  if ($start -gt 0) { $snippet = '...' + $snippet }
  if ($start + $length -lt $normalized.Length) { $snippet += '...' }
  return $snippet
}

function Get-DoiList([string]$Text) {
  $matches = [regex]::Matches($Text, '\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b', 'IgnoreCase')
  return @($matches | ForEach-Object { $_.Value.TrimEnd('.', ',', ';', ')') } | Select-Object -Unique)
}

$one = New-Object -ComObject OneNote.Application
[xml]$hierarchy = ''
$one.GetHierarchy('', 4, [ref]$hierarchy)

$results = New-Object System.Collections.ArrayList
$sections = @($hierarchy.DocumentElement.SelectNodes('//*[local-name()="Section"]') | Where-Object {
  $_.name -notlike '分類_*' -and $_.name -ne '削除されたページ'
})

foreach ($section in $sections) {
  [xml]$sectionXml = ''
  try {
    $one.GetHierarchy($section.ID, 4, [ref]$sectionXml)
  } catch {
    continue
  }

  foreach ($page in @($sectionXml.DocumentElement.SelectNodes('.//*[local-name()="Page"]'))) {
    $pageXml = ''
    try {
      $one.GetPageContent($page.ID, [ref]$pageXml, 0)
    } catch {
      continue
    }

    $plain = ''
    try {
      [xml]$px = $pageXml
      $texts = @($px.SelectNodes('//*[local-name()="T"]') | ForEach-Object { $_.InnerText })
      $plain = (($texts -join ' ') -replace '\s+', ' ').Trim()
    } catch {
      $plain = ''
    }

    $hay = (($page.name + ' ' + (Get-SectionPath $section) + ' ' + $plain)).ToLowerInvariant()
    $matched = @($terms | Where-Object { $hay.Contains($_.ToLowerInvariant()) })
    if ($matched.Count -eq 0) {
      continue
    }

    $doiList = Get-DoiList $plain
    $link = ''
    try {
      $one.GetHyperlinkToObject($page.ID, '', [ref]$link)
    } catch {
      $link = $page.ID
    }

    $score = $matched.Count
    if ($doiList.Count -gt 0) { $score += 5 }
    if ($page.name -match '合成|synthesis|precursor|前駆体|モノマー') { $score += 3 }
    if ((Get-SectionPath $section) -match '合成|研究|実験|論文') { $score += 2 }
    if ($plain -match 'DOI|doi|参考文献|文献|Chem|J\.|ACS|Wiley|RSC|Nature|Science') { $score += 2 }
    if ($plain -match '試薬|条件|温度|時間|溶媒|収率|yield|equiv|mol|mmol|catalyst') { $score += 2 }

    [void]$results.Add([pscustomobject]@{
      Score = $score
      Section = Get-SectionPath $section
      Page = $page.name
      DOI = (($doiList | Select-Object -First 5) -join ', ')
      Terms = ($matched -join ', ')
      LastModified = $page.lastModifiedTime
      Link = $link
      Snippet = Get-Snippet $plain $matched
    })
  }
}

$topResults = @($results |
  Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'LastModified'; Descending = $true } |
  Select-Object -First 30)

Write-Host '## 関連OneNote候補'
$topResults | Format-List

Write-Host ''
Write-Host '## 検討観点'
Write-Host '- 合成ルート: 逆合成の切断点、入手可能な前駆体、保護基の要否、最後段階で避けたい高リスク反応を確認してください。'
Write-Host '- 文献探索: ページ内の DOI、反応名、著者名、分子名を起点に一次文献を確認してください。本文で見つからない場合は、Supplementary Information / Supporting Information / ESI / SI / 特許 / PDF本文まで確認してください。'
Write-Host '- 試薬・条件: 溶媒、温度、時間、触媒、塩基、濃度、精製法、収率、スケールを表にして比較してください。'
Write-Host '- ソース必須: 根拠文献のタイトル、URL、DOIを必ず記録してください。ソースが示せない候補は未確定として扱ってください。'
Write-Host '- 直接条件が見つかった場合: 試薬量、mmol、当量、溶媒、温度、時間、後処理、精製、収率、NMR/MS、DOI、SIリンクを必ず記録してください。'
Write-Host '- 安全確認: ハロゲン化、強塩基、有機金属、加圧/高温、毒性溶媒は別途リスク確認してください。'
