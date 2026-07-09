$ErrorActionPreference = 'Stop'

$links = @(
  [pscustomobject]@{
    Title = '太陽電池'
    Url = 'onenote:https://d.docs.live.net/a7339f0ae206a171/ドキュメント/2026実験/仕事/授業.one#太陽電池&section-id={41BD0723-FE0A-4F24-A48B-126C5CE99B38}&page-id={D517CBE7-3EDA-DE4F-9733-D037CE4F4261}&end'
  },
  [pscustomobject]@{
    Title = 'ドリフトの拡散方程式'
    Url = 'onenote:https://d.docs.live.net/a7339f0ae206a171/ドキュメント/2026実験/仕事/授業.one#ドリフトの拡散方程式&section-id={41BD0723-FE0A-4F24-A48B-126C5CE99B38}&page-id={DBFAA935-672E-4328-B1DA-9A3DA7FEAEB6}&end'
  },
  [pscustomobject]@{
    Title = 'GeSe・シフト電流の修士論文メモ'
    Url = 'onenote:https://d.docs.live.net/A7339F0AE206A171/ドキュメント/2025年書込テスト/新しいセクション%201.one#この文書は、バルクGeSeにおける反強誘電性から強誘電性への相転移中のシフト電流の観測に関する修士論文です。&section-id={1DB30EC1-9C11-F546-9C52-33B995B99966}&page-id={DE0DC816-C680-DF4A-90CD-2FC05092E396}&end'
  },
  [pscustomobject]@{
    Title = 'Photovoltaic 論文メモ'
    Url = 'onenote:https://d.docs.live.net/A7339F0AE206A171/ドキュメント/2025年4月日誌/SolubleGNR合成.one#Macromolecular%20Rapid%20Communications%20-%202023%20-%20Ji%20-%20Deciphering%20the&section-id={EE749569-F14D-477B-89AB-15C1B550D1AF}&page-id={D7FB07F4-F932-488F-ACF3-7CC45540078E}&end'
  },
  [pscustomobject]@{
    Title = 'Annual Report作成'
    Url = 'onenote:https://d.docs.live.net/a7339f0ae206a171/ドキュメント/2026実験/仕事/アーカイブ/メモアーカイブ.one#Annual%20Report作成&section-id={DE7D454E-55DB-44A4-B69D-25EAE1F81B7D}&page-id={3E7C4386-F9BA-4B3A-81AA-DF08E5AA6A3A}&end'
  },
  [pscustomobject]@{
    Title = 'Ladderpolymer'
    Url = 'onenote:https://d.docs.live.net/A7339F0AE206A171/ドキュメント/2025年7月日誌/新しいセクション%201.one#Ladderpolymer&section-id={1F052DA1-0F8C-4A3E-A850-9E0C48D463FD}&page-id={F391F3E2-A225-491E-8447-07F9457B9EF3}&end'
  },
  [pscustomobject]@{
    Title = '量子ドット合成法メモ'
    Url = 'onenote:https://d.docs.live.net/A7339F0AE206A171/ドキュメント/2025年書込テスト/とりあえず投げ込み９月_.one#https//www.sigmaaldrich.com/JP/ja/technical-documents&section-id={6B057F8E-13F3-4FD9-82D1-BA0733E30F17}&page-id={7CA53EA4-C845-4FB5-B797-0E47760B9092}&end'
  }
)

Write-Host ''
Write-Host '太陽電池関連 OneNote リンク'
Write-Host '---------------------------'
for ($i = 0; $i -lt $links.Count; $i++) {
  Write-Host ("{0}. {1}" -f ($i + 1), $links[$i].Title)
}
Write-Host ''
$answer = Read-Host '開く番号を入力してください'

$index = 0
if (-not [int]::TryParse($answer, [ref]$index) -or $index -lt 1 -or $index -gt $links.Count) {
  Write-Host '番号が正しくありません。'
  exit 1
}

Start-Process $links[$index - 1].Url
