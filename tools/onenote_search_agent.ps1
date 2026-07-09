param(
  [Parameter(Mandatory = $true)]
  [string]$Query
)

$ErrorActionPreference = 'Stop'
$terms = @($Query)

if ($Query -match '太陽電池') {
  $terms += @('ソーラー', 'solar', 'photovoltaic', 'PV', 'ペロブスカイト', 'perovskite', '光起電', '光電変換')
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

    $hay = (($page.name + ' ' + $section.name + ' ' + $plain)).ToLowerInvariant()
    $matched = @($terms | Where-Object { $hay.Contains($_.ToLowerInvariant()) })
    if ($matched.Count -eq 0) {
      continue
    }

    $link = ''
    try {
      $one.GetHyperlinkToObject($page.ID, '', [ref]$link)
    } catch {
      $link = $page.ID
    }

    $snippet = $plain
    if ($snippet.Length -gt 220) {
      $snippet = $snippet.Substring(0, 220) + '...'
    }

    [void]$results.Add([pscustomobject]@{
      Section = $section.name
      Page = $page.name
      Terms = ($matched -join ', ')
      LastModified = $page.lastModifiedTime
      Link = $link
      Snippet = $snippet
    })
  }
}

$results | Sort-Object LastModified -Descending | Format-List
