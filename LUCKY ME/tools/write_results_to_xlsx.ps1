param(
  [string]$WorkbookPath = 'C:\Users\jan.jedlicka\Desktop\Fotbal to je hra.xlsx'
)

$ErrorActionPreference = 'Stop'

function Get-ColumnNumber([string]$letters) {
  $sum = 0
  foreach ($ch in $letters.ToCharArray()) {
    $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $sum
}

function Ensure-InlineStringCell {
  param(
    [System.Xml.XmlElement]$RowNode,
    [string]$CellRef,
    [string]$Text,
    [string]$StyleId
  )

  $worksheet = $RowNode.OwnerDocument.DocumentElement
  $nsUri = $worksheet.NamespaceURI

  $cell = $null
  foreach ($existing in $RowNode.ChildNodes) {
    if ($existing.LocalName -eq 'c' -and $existing.GetAttribute('r') -eq $CellRef) {
      $cell = $existing
      break
    }
  }

  if (-not $cell) {
    $cell = $RowNode.OwnerDocument.CreateElement('c', $nsUri)
    [void]$cell.SetAttribute('r', $CellRef)

    $targetCol = ($CellRef -replace '\d', '')
    $targetColNum = Get-ColumnNumber $targetCol
    $insertBefore = $null

    foreach ($existing in $RowNode.ChildNodes) {
      if ($existing.LocalName -ne 'c') { continue }
      $existingCol = ($existing.GetAttribute('r') -replace '\d', '')
      if ((Get-ColumnNumber $existingCol) -gt $targetColNum) {
        $insertBefore = $existing
        break
      }
    }

    if ($insertBefore) {
      [void]$RowNode.InsertBefore($cell, $insertBefore)
    } else {
      [void]$RowNode.AppendChild($cell)
    }
  } else {
    $cell.RemoveAll()
    [void]$cell.SetAttribute('r', $CellRef)
  }

  [void]$cell.SetAttribute('t', 'inlineStr')
  if ($StyleId) {
    [void]$cell.SetAttribute('s', $StyleId)
  }

  $isNode = $RowNode.OwnerDocument.CreateElement('is', $nsUri)
  $tNode = $RowNode.OwnerDocument.CreateElement('t', $nsUri)
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

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('xlsx-edit-' + [guid]::NewGuid().ToString('N'))
$zipPath = Join-Path ([System.IO.Path]::GetTempPath()) ('xlsx-edit-' + [guid]::NewGuid().ToString('N') + '.zip')
$sourceZipPath = Join-Path ([System.IO.Path]::GetTempPath()) ('xlsx-source-' + [guid]::NewGuid().ToString('N') + '.zip')
$backupPath = Join-Path ([System.IO.Path]::GetDirectoryName($WorkbookPath)) (([System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath)) + ' - zaloha před výsledky.xlsx')

New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
  Copy-Item -Path $WorkbookPath -Destination $backupPath -Force
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  Copy-Item -Path $WorkbookPath -Destination $sourceZipPath -Force
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZipPath, $tempRoot)

  $sheetPath = Join-Path $tempRoot 'xl\worksheets\sheet1.xml'
  [xml]$sheetXml = Get-Content -Raw -Path $sheetPath

  $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
  $ns.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

  $rows = $sheetXml.SelectNodes('//d:worksheet/d:sheetData/d:row', $ns)
  $headerStyle = ''
  $bodyStyle = ''

  $headerCell = $sheetXml.SelectSingleNode('//d:row[@r="1"]/d:c[@r="I1"]', $ns)
  if ($headerCell -and $headerCell.HasAttribute('s')) {
    $headerStyle = $headerCell.GetAttribute('s')
  } else {
    $fallbackHeader = $sheetXml.SelectSingleNode('//d:row[@r="1"]/d:c[@r="H1"]', $ns)
    if ($fallbackHeader -and $fallbackHeader.HasAttribute('s')) {
      $headerStyle = $fallbackHeader.GetAttribute('s')
    }
  }

  $bodyTemplate = $sheetXml.SelectSingleNode('//d:row[@r="2"]/d:c[@r="I2"]', $ns)
  if ($bodyTemplate -and $bodyTemplate.HasAttribute('s')) {
    $bodyStyle = $bodyTemplate.GetAttribute('s')
  }

  foreach ($row in $rows) {
    $rowNumber = [int]$row.GetAttribute('r')
    $cellRef = "J$rowNumber"

    if ($rowNumber -eq 1) {
      Ensure-InlineStringCell -RowNode $row -CellRef $cellRef -Text 'VÝSLEDEK' -StyleId $headerStyle
      continue
    }

    $matchCell = $sheetXml.SelectSingleNode("//d:row[@r='$rowNumber']/d:c[@r='C$rowNumber']", $ns)
    if (-not $matchCell) { continue }

    $matchText = ''
    if ($matchCell.GetAttribute('t') -eq 'inlineStr') {
      $matchText = $matchCell.is.t.InnerText
    } elseif ($matchCell.GetAttribute('t') -eq 's') {
      continue
    } else {
      $matchText = [string]$matchCell.v
    }
  }

  $zipRead = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  try {
    $sharedEntry = $zipRead.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }
    [xml]$sharedXml = (New-Object IO.StreamReader($sharedEntry.Open())).ReadToEnd()
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
  } finally {
    $zipRead.Dispose()
  }

  foreach ($row in $rows) {
    $rowNumber = [int]$row.GetAttribute('r')
    if ($rowNumber -eq 1) { continue }

    $matchCell = $sheetXml.SelectSingleNode("//d:row[@r='$rowNumber']/d:c[@r='C$rowNumber']", $ns)
    if (-not $matchCell) { continue }

    $matchText = ''
    if ($matchCell.GetAttribute('t') -eq 's') {
      $sharedIndex = [int]$matchCell.v.InnerText
      if ($sharedIndex -ge 0 -and $sharedIndex -lt $shared.Count) {
        $matchText = $shared[$sharedIndex]
      }
    } elseif ($matchCell.GetAttribute('t') -eq 'inlineStr') {
      $matchText = $matchCell.is.t.InnerText
    } else {
      $matchText = [string]$matchCell.v
    }

    if ($scoreMap.ContainsKey($matchText)) {
      Ensure-InlineStringCell -RowNode $row -CellRef "J$rowNumber" -Text $scoreMap[$matchText] -StyleId $bodyStyle
    }
  }

  $sheetXml.Save($sheetPath)

  if (Test-Path $zipPath) { Remove-Item -Path $zipPath -Force }
  [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $zipPath)
  Copy-Item -Path $zipPath -Destination $WorkbookPath -Force

  [pscustomobject]@{
    Workbook = $WorkbookPath
    Backup = $backupPath
    FilledMatches = $scoreMap.Count
    Unresolved = @('Flamutari Vore - AF Abassani')
  } | ConvertTo-Json -Depth 4
}
finally {
  if (Test-Path $tempRoot) { Remove-Item -Path $tempRoot -Recurse -Force }
  if (Test-Path $zipPath) { Remove-Item -Path $zipPath -Force }
  if (Test-Path $sourceZipPath) { Remove-Item -Path $sourceZipPath -Force }
}
