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
