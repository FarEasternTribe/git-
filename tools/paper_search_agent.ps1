param(
  [Parameter(Mandatory = $true)]
  [string]$Query
)

$ErrorActionPreference = 'Stop'

$terms = @(
  $Query,
  '論文',
  '文献',
  'paper',
  'literature',
  'reference',
  'source',
  'DOI',
  'doi.org',
  'journal',
  'authors',
  'title',
  'Supplementary',
  'Supporting Information',
  'SI',
  'ESI',
  'PDF'
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

Write-Host '論文検索Agent'
Write-Host '役割: 論文・文献・根拠ソースを探索し、候補ごとにタイトル、著者、掲載誌、年、URL、DOIを提示する'
Write-Host '必須出力: ソースとなる文献のDOIを必ず添付する。DOIが見つからない場合は探索範囲と代替一次ソースを明示する'
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

function Get-DoiList([string]$Text) {
  $matches = [regex]::Matches($Text, '\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b', 'IgnoreCase')
  return @($matches | ForEach-Object { $_.Value.TrimEnd('.', ',', ';', ')') } | Select-Object -Unique)
}

function Get-Snippet([string]$Text, [string[]]$Terms) {
  $normalized = (($Text -replace '\s+', ' ').Trim())
  if ([string]::IsNullOrWhiteSpace($normalized)) { return '' }
  $lower = $normalized.ToLowerInvariant()
  $hitIndex = -1
  foreach ($term in $Terms) {
    if ([string]::IsNullOrWhiteSpace($term)) { continue }
    $idx = $lower.IndexOf($term.ToLowerInvariant())
    if ($idx -ge 0 -and ($hitIndex -lt 0 -or $idx -lt $hitIndex)) { $hitIndex = $idx }
  }
  if ($hitIndex -lt 0) { $hitIndex = 0 }
  $start = [Math]::Max(0, $hitIndex - 80)
  $length = [Math]::Min(420, $normalized.Length - $start)
  $snippet = $normalized.Substring($start, $length)
  if ($start -gt 0) { $snippet = '...' + $snippet }
  if ($start + $length -lt $normalized.Length) { $snippet += '...' }
  return $snippet
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

    $sectionPath = Get-SectionPath $section
    $hay = (($page.name + ' ' + $sectionPath + ' ' + $plain)).ToLowerInvariant()
    $matched = @($terms | Where-Object { $hay.Contains($_.ToLowerInvariant()) })
    if ($matched.Count -eq 0) { continue }

    $link = ''
    try {
      $one.GetHyperlinkToObject($page.ID, '', [ref]$link)
    } catch {
      $link = $page.ID
    }

    $doiList = Get-DoiList $plain
    $score = $matched.Count
    if ($doiList.Count -gt 0) { $score += 5 }
    if ($plain -match 'Journal|Nature|Science|ACS|Wiley|RSC|PubMed|arXiv|掲載誌|著者|Authors') { $score += 3 }
    if ($page.name -match '論文|文献|paper|literature|DOI') { $score += 2 }

    [void]$results.Add([pscustomobject]@{
      Score = $score
      Section = $sectionPath
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

Write-Host '## 関連文献候補'
$topResults | Format-List

Write-Host ''
Write-Host '## 検討観点'
Write-Host '- DOI必須: 候補ごとにタイトル、著者、掲載誌、年、URL、DOIを確認してください。'
Write-Host '- DOIなし候補: 未確定として扱い、Crossref、出版社ページ、PubMed、Google Scholar、PDF本文まで検索範囲を広げてください。'
Write-Host '- ソース確認: レビューや二次情報だけでなく、一次文献または出版社ページを優先してください。'

