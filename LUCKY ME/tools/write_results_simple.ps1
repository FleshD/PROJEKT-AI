param(
  [string]$WorkbookPath = 'C:\Users\jan.jedlicka\Desktop\Fotbal to je hra.xlsx'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-CellText {
  param(
    [System.Xml.XmlElement]$Cell,
    [string[]]$Shared
  )

  if (-not $Cell) { return '' }

  $type = $Cell.GetAttribute('t')
  if ($type -eq 's') {
    $index = [int]$Cell.SelectSingleNode('*[local-name()="v"]').InnerText
    if ($index -ge 0 -and $index -lt $Shared.Count) {
      return $Shared[$index]
    }
    return ''
  }

  if ($type -eq 'inlineStr') {
    $tNode = $Cell.SelectSingleNode('*[local-name()="is"]/*[local-name()="t"]')
    if ($tNode) { return [string]$tNode.InnerText }
    return ''
  }

  $vNode = $Cell.SelectSingleNode('*[local-name()="v"]')
  if ($vNode) { return [string]$vNode.InnerText }
  return ''
}

function Set-InlineCell {
  param(
    [System.Xml.XmlElement]$Row,
    [string]$Ref,
    [string]$Text,
    [string]$StyleId = ''
  )

  $doc = $Row.OwnerDocument
  $nsUri = $doc.DocumentElement.NamespaceURI
  $cell = $null

  foreach ($node in $Row.ChildNodes) {
    if ($node.LocalName -eq 'c' -and $node.GetAttribute('r') -eq $Ref) {
      $cell = $node
      break
    }
  }

  if (-not $cell) {
    $cell = $doc.CreateElement('c', $nsUri)
    [void]$cell.SetAttribute('r', $Ref)
    [void]$Row.AppendChild($cell)
  } else {
    $cell.RemoveAll()
    [void]$cell.SetAttribute('r', $Ref)
  }

  [void]$cell.SetAttribute('t', 'inlineStr')
  if ($StyleId) {
    [void]$cell.SetAttribute('s', $StyleId)
  }

  $isNode = $doc.CreateElement('is', $nsUri)
  $tNode = $doc.CreateElement('t', $nsUri)
  $tNode.InnerText = $Text
  [void]$isNode.AppendChild($tNode)
  [void]$cell.AppendChild($isNode)
}

$scoreMap = @{
  'Bellgrano - Talleres' = '0:0'
  'Genk - St. Truidense' = '1:0'
  'Lommel - Francs Boais' = '3:1'
  'Palmeiras - Mirosol' = '1:0'
  'Plzeň - Boheminas' = '2:1'
  'Sparta - Slovácko' = '3:2'
  'Zlín - Slavia' = '1:3'
  'Slavia B - Prostějov' = '1:1'
  'Vlašim - Č. Budějovice' = '4:2'
  'Indepedente - Aucas' = '2:2'
  'Arsenal - Everton' = '2:0'
  'Sunderland - Brighton' = '0:1'
  'Middlesborough - Bristol City' = '1:1'
  'Millwal - Blackburn' = '1:2'
  'Norwich - Preston' = '2:0'
  'Lincoln - Stockport' = '3:1'
  'Flatwood - Tranmere' = '0:0'
  'Notts county - Chesterfield' = '2:3'
  'Swindon - Milton Keynes' = '1:2'
  'Aldershort - York' = '0:3'
  'Tamworth - Carlislie' = '2:1'
  'Bath City - Hornchurch' = '2:2'
  'Kidderminster - Chester' = '0:0'
  'Atletico M - Getafe' = '1:0'
  'Real M - Elche' = '4:1'
  'Leonessa - Santander' = '1:2'
  'Malaga - Huesca' = '5:3'
  'Zaragoza - Almeira' = '2:0'
  'Le Harve - Lyon' = '0:0'
  'Lorient - Lens' = '2:1'
  'Monaco - Brest' = '2:0'
  'Strasbourgh - Paris FC' = '0:0'
  'Amiens - Le Mans' = '3:4'
  'Annecy - Troyes' = '1:2'
  'Dortmund - Augsburg' = '2:0'
  'Frankfurt - Hildenheim' = '1:0'
  'Hoffeinheim - Wolfsburg' = '1:1'
  'Bielfeld - Paderborn' = '2:2'
  'Schalke - Hannover' = '2:2'
  'Heerenven - Telstar' = '3:0'
  'Sambdoria - Venezia' = '0:0'
}

$backupPath = Join-Path ([IO.Path]::GetDirectoryName($WorkbookPath)) (([IO.Path]::GetFileNameWithoutExtension($WorkbookPath)) + ' - zaloha pred vysledky.xlsx')
$sourceZip = Join-Path ([IO.Path]::GetTempPath()) ('xlsx-source-' + [guid]::NewGuid().ToString('N') + '.zip')
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('xlsx-work-' + [guid]::NewGuid().ToString('N'))
$outZip = Join-Path ([IO.Path]::GetTempPath()) ('xlsx-out-' + [guid]::NewGuid().ToString('N') + '.zip')

New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
  if (-not (Test-Path $backupPath)) {
    Copy-Item -Path $WorkbookPath -Destination $backupPath -Force
  }

  Copy-Item -Path $WorkbookPath -Destination $sourceZip -Force
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZip, $tempRoot)

  [xml]$sharedXml = Get-Content -Raw -Path (Join-Path $tempRoot 'xl\sharedStrings.xml')
  [xml]$sheetXml = Get-Content -Raw -Path (Join-Path $tempRoot 'xl\worksheets\sheet1.xml')

  $shared = @()
  foreach ($si in $sharedXml.sst.si) {
    if ($si.t) {
      $shared += [string]$si.t
    } elseif ($si.r) {
      $parts = @()
      foreach ($run in $si.r) {
        if ($run.t) { $parts += [string]$run.t }
      }
      $shared += ($parts -join '')
    } else {
      $shared += ''
    }
  }

  $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
  $ns.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

  $headerStyle = ''
  $i1 = $sheetXml.SelectSingleNode('//d:row[@r="1"]/d:c[@r="I1"]', $ns)
  if ($i1 -and $i1.HasAttribute('s')) {
    $headerStyle = $i1.GetAttribute('s')
  }

  $filled = 0
  $rows = $sheetXml.SelectNodes('//d:sheetData/d:row', $ns)
  foreach ($row in $rows) {
    $rowNum = [int]$row.GetAttribute('r')
    if ($rowNum -eq 1) {
      Set-InlineCell -Row $row -Ref 'J1' -Text 'VYSLEDEK' -StyleId $headerStyle
      continue
    }

    $matchCell = $sheetXml.SelectSingleNode("//d:row[@r='$rowNum']/d:c[@r='C$rowNum']", $ns)
    $matchText = Get-CellText -Cell $matchCell -Shared $shared
    if ($scoreMap.ContainsKey($matchText)) {
      Set-InlineCell -Row $row -Ref "J$rowNum" -Text $scoreMap[$matchText]
      $filled++
    }
  }

  $sheetXml.Save((Join-Path $tempRoot 'xl\worksheets\sheet1.xml'))
  [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $outZip)
  Copy-Item -Path $outZip -Destination $WorkbookPath -Force

  [pscustomobject]@{
    Workbook = $WorkbookPath
    Backup = $backupPath
    FilledRows = $filled
    Unresolved = @('Flamutari Vore - AF Abassani')
  } | ConvertTo-Json -Depth 4
}
finally {
  if (Test-Path $sourceZip) { Remove-Item $sourceZip -Force }
  if (Test-Path $outZip) { Remove-Item $outZip -Force }
  if (Test-Path $tempRoot) { Remove-Item $tempRoot -Recurse -Force }
}
