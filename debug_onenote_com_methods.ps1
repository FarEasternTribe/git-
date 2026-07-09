$ErrorActionPreference = 'Stop'
$one = New-Object -ComObject OneNote.Application
$one | Get-Member -Name GetBinaryPageContent,GetPageContent,UpdatePageContent,Publish | Format-List *
