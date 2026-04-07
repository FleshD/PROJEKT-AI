param(
  [string]$InputPath = '',
  [string]$OutputPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO\sablona_technicke_zpravy_v1.1_navrh.docx'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not $InputPath) {
  $template = Get-ChildItem 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO' -Filter '*.docx' |
    Where-Object { $_.Name -like 'šablona technické zprávy_v1.0*' -or $_.Name -eq 'sablona_technicke_zpravy_v1.0.docx' } |
    Select-Object -First 1

  if (-not $template) {
    $template = Get-ChildItem 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO' -Filter '*.docx' |
      Where-Object { $_.Name -like 'šablona technické zprávy*' -or $_.Name -like '*technick*' } |
      Sort-Object LastWriteTime |
      Select-Object -First 1
  }

  if (-not $template) {
    throw 'Nepodařilo se najít zdrojovou šablonu technické zprávy ve složce UDIMO.'
  }

  $InputPath = $template.FullName
}

function Get-XmlDocument {
  param(
    [string]$Path
  )

  $doc = New-Object System.Xml.XmlDocument
  $doc.PreserveWhitespace = $true
  $doc.Load($Path)
  return $doc
}

function Get-NamespaceManager {
  param(
    [System.Xml.XmlDocument]$Document
  )

  $ns = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
  $ns.AddNamespace('w',  'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
  $ns.AddNamespace('a',  'http://schemas.openxmlformats.org/drawingml/2006/main')
  $ns.AddNamespace('r',  'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  return $ns
}

function Update-AnchorByRid {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Rid,
    [string]$HOffset,
    [string]$VOffset,
    [string]$Cx,
    [string]$Cy
  )

  $anchor = $Document.SelectSingleNode("//*[local-name()='blip' and @*[local-name()='embed' and .='$Rid']]/ancestor::*[local-name()='anchor'][1]")
  if (-not $anchor) {
    throw "Nepodařilo se najít anchor pro $Rid."
  }

  $hNode = $anchor.SelectSingleNode("./*[local-name()='positionH']/*[local-name()='posOffset']")
  $vNode = $anchor.SelectSingleNode("./*[local-name()='positionV']/*[local-name()='posOffset']")
  $extent = $anchor.SelectSingleNode("./*[local-name()='extent']")

  if (-not $hNode -or -not $vNode -or -not $extent) {
    throw "Anchor pro $Rid nemá očekávanou strukturu."
  }

  $hNode.InnerText = $HOffset
  $vNode.InnerText = $VOffset
  $extent.SetAttribute('cx', $Cx)
  $extent.SetAttribute('cy', $Cy)
}

function Update-LineAnchorByName {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Name,
    [string]$HOffset,
    [string]$VOffset,
    [string]$Cx
  )

  $anchor = $Document.SelectSingleNode("//*[local-name()='anchor'][./*[local-name()='docPr' and @name=`"$Name`"]]")
  if (-not $anchor) {
    return
  }

  $hNode = $anchor.SelectSingleNode("./*[local-name()='positionH']/*[local-name()='posOffset']")
  $vNode = $anchor.SelectSingleNode("./*[local-name()='positionV']/*[local-name()='posOffset']")
  $extent = $anchor.SelectSingleNode("./*[local-name()='extent']")

  if (-not $hNode -or -not $vNode -or -not $extent) {
    throw "Line anchor pro $Name nemá očekávanou strukturu."
  }

  $hNode.InnerText = $HOffset
  $vNode.InnerText = $VOffset
  $extent.SetAttribute('cx', $Cx)
}

function Remove-HeaderImageRun {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Rid
  )

  $run = $Document.SelectSingleNode("//*[local-name()='blip' and @*[local-name()='embed' and .='$Rid']]/ancestor::*[local-name()='r'][1]")
  if ($run -and $run.ParentNode) {
    [void]$run.ParentNode.RemoveChild($run)
  }
}

function Write-ZipFromDirectory {
  param(
    [string]$SourceDir,
    [string]$DestinationZip
  )

  if (Test-Path -LiteralPath $DestinationZip) {
    Remove-Item -LiteralPath $DestinationZip -Force
  }

  $archive = [System.IO.Compression.ZipFile]::Open($DestinationZip, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    Get-ChildItem -LiteralPath $SourceDir -Recurse -File | ForEach-Object {
      $relative = $_.FullName.Substring($SourceDir.Length).TrimStart('\')
      $entryName = $relative -replace '\\', '/'
      [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
  }
  finally {
    $archive.Dispose()
  }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('tech-template-' + [guid]::NewGuid().ToString('N'))
$sourceZip = Join-Path ([IO.Path]::GetTempPath()) ('tech-template-src-' + [guid]::NewGuid().ToString('N') + '.zip')
$outZip = Join-Path ([IO.Path]::GetTempPath()) ('tech-template-out-' + [guid]::NewGuid().ToString('N') + '.zip')

New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
  Copy-Item -LiteralPath $InputPath -Destination $sourceZip -Force
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZip, $tempRoot)

  $documentPath = Join-Path $tempRoot 'word\document.xml'
  $document = Get-XmlDocument -Path $documentPath

  # Title page: split logos into a primary publicity row and a secondary partner row.
  Update-AnchorByRid -Document $document -Rid 'rId8'  -HOffset '860000'  -VOffset '5950000' -Cx '2012950' -Cy '295910'
  Update-AnchorByRid -Document $document -Rid 'rId9'  -HOffset '3190000' -VOffset '5890000' -Cx '1500000' -Cy '387931'
  Update-AnchorByRid -Document $document -Rid 'rId10' -HOffset '1320000' -VOffset '6990000' -Cx '1500000' -Cy '840625'
  Update-AnchorByRid -Document $document -Rid 'rId11' -HOffset '3410000' -VOffset '7120000' -Cx '950000'  -Cy '318292'
  Update-LineAnchorByName -Document $document -Name 'Přímá spojnice 18' -HOffset '870000' -VOffset '6700000' -Cx '3900000'
  $document.Save($documentPath)

  foreach ($headerName in @('header4.xml', 'header5.xml')) {
    $headerPath = Join-Path $tempRoot ('word\' + $headerName)
    $headerDoc = Get-XmlDocument -Path $headerPath
    Remove-HeaderImageRun -Document $headerDoc -Rid 'rId1'
    $headerDoc.Save($headerPath)
  }

  Write-ZipFromDirectory -SourceDir $tempRoot -DestinationZip $outZip
  Copy-Item -LiteralPath $outZip -Destination $OutputPath -Force

  Write-Output $OutputPath
}
finally {
  if (Test-Path -LiteralPath $sourceZip) { Remove-Item -LiteralPath $sourceZip -Force }
  if (Test-Path -LiteralPath $outZip) { Remove-Item -LiteralPath $outZip -Force }
  if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
}
