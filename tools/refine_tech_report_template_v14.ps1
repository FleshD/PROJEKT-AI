param(
  [string]$InputPath = '',
  [string]$OutputPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO\sablona_technicke_zpravy_v1.4_navrh.docx'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$baseDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO'

if (-not $InputPath) {
  $template = Get-ChildItem -LiteralPath $baseDir -Filter '*.docx' |
    Where-Object { $_.Name -like '*technick*' -and $_.Name -like '*v1.0*' } |
    Select-Object -First 1

  if (-not $template) {
    throw 'Could not find the v1.0 technical report template in the UDIMO folder.'
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

function New-NamespaceManager {
  param(
    [System.Xml.XmlDocument]$Document
  )

  $ns = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
  $ns.AddNamespace('w',    'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $ns.AddNamespace('wp',   'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
  $ns.AddNamespace('a',    'http://schemas.openxmlformats.org/drawingml/2006/main')
  $ns.AddNamespace('r',    'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $ns.AddNamespace('pic',  'http://schemas.openxmlformats.org/drawingml/2006/picture')
  $ns.AddNamespace('wps',  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
  $ns.AddNamespace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')
  return $ns
}

function Set-OrCreateAttribute {
  param(
    [System.Xml.XmlElement]$Element,
    [string]$Prefix,
    [string]$LocalName,
    [string]$NamespaceUri,
    [string]$Value
  )

  if ($NamespaceUri) {
    $existing = $Element.Attributes.GetNamedItem($LocalName, $NamespaceUri)
    if ($existing) {
      [void]$Element.Attributes.RemoveNamedItem($LocalName, $NamespaceUri)
    }

    $attr = $Element.OwnerDocument.CreateAttribute($Prefix, $LocalName, $NamespaceUri)
    $attr.Value = $Value
    [void]$Element.Attributes.Append($attr)
  }
  else {
    $Element.SetAttribute($LocalName, $Value)
  }
}

function Set-AnchorPosition {
  param(
    [System.Xml.XmlDocument]$Document,
    [System.Xml.XmlElement]$Anchor,
    [string]$HOffset,
    [string]$VOffset
  )

  $nsWp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

  $positionH = $Anchor.SelectSingleNode("./*[local-name()='positionH']")
  $positionV = $Anchor.SelectSingleNode("./*[local-name()='positionV']")

  if (-not $positionH -or -not $positionV) {
    throw 'Anchor is missing positionH or positionV.'
  }

  $positionH.RemoveAll()
  $positionV.RemoveAll()

  Set-OrCreateAttribute -Element $positionH -Prefix '' -LocalName 'relativeFrom' -NamespaceUri '' -Value 'margin'
  Set-OrCreateAttribute -Element $positionV -Prefix '' -LocalName 'relativeFrom' -NamespaceUri '' -Value 'paragraph'

  $hNode = $Document.CreateElement('wp', 'posOffset', $nsWp)
  $hNode.InnerText = $HOffset
  [void]$positionH.AppendChild($hNode)

  $vNode = $Document.CreateElement('wp', 'posOffset', $nsWp)
  $vNode.InnerText = $VOffset
  [void]$positionV.AppendChild($vNode)
}

function Set-AnchorSize {
  param(
    [System.Xml.XmlElement]$Anchor,
    [string]$Cx,
    [string]$Cy
  )

  $extent = $Anchor.SelectSingleNode("./*[local-name()='extent']")
  if (-not $extent) {
    throw 'Anchor is missing wp:extent.'
  }
  $extent.SetAttribute('cx', $Cx)
  $extent.SetAttribute('cy', $Cy)

  foreach ($innerExtent in $Anchor.SelectNodes(".//*[local-name()='xfrm']/*[local-name()='ext']")) {
    $innerExtent.SetAttribute('cx', $Cx)
    $innerExtent.SetAttribute('cy', $Cy)
  }
}

function Set-AnchorVisibility {
  param(
    [System.Xml.XmlElement]$Anchor,
    [bool]$BehindDoc
  )

  $Anchor.SetAttribute('behindDoc', $(if ($BehindDoc) { '1' } else { '0' }))
}

function New-HexId {
  return ([guid]::NewGuid().ToString('N').Substring(0, 8).ToUpperInvariant())
}

function Get-MaxDocPrId {
  param(
    [System.Xml.XmlDocument]$Document
  )

  $maxId = 0
  foreach ($docPr in $Document.SelectNodes("//*[local-name()='docPr']")) {
    $value = 0
    if ([int]::TryParse($docPr.GetAttribute('id'), [ref]$value) -and $value -gt $maxId) {
      $maxId = $value
    }
  }

  foreach ($cNvPr in $Document.SelectNodes("//*[local-name()='cNvPr']")) {
    $value = 0
    if ([int]::TryParse($cNvPr.GetAttribute('id'), [ref]$value) -and $value -gt $maxId) {
      $maxId = $value
    }
  }

  return $maxId
}

function Get-MaxRelativeHeight {
  param(
    [System.Xml.XmlDocument]$Document
  )

  $maxValue = [int64]0
  foreach ($anchor in $Document.SelectNodes("//*[local-name()='anchor']")) {
    $value = [int64]0
    if ([int64]::TryParse($anchor.GetAttribute('relativeHeight'), [ref]$value) -and $value -gt $maxValue) {
      $maxValue = $value
    }
  }

  return $maxValue
}

function Set-DrawingIds {
  param(
    [System.Xml.XmlElement]$Anchor,
    [int]$DocPrId,
    [string]$Name
  )

  $docPr = $Anchor.SelectSingleNode("./*[local-name()='docPr']")
  if (-not $docPr) {
    throw 'Anchor is missing wp:docPr.'
  }

  $docPr.SetAttribute('id', [string]$DocPrId)
  $docPr.SetAttribute('name', $Name)

  $picDocPr = $Anchor.SelectSingleNode(".//*[local-name()='cNvPr']")
  if ($picDocPr) {
    $picDocPr.SetAttribute('id', [string]$DocPrId)
    $picDocPr.SetAttribute('name', $Name)
  }
}

function Prepare-AnchorClone {
  param(
    [System.Xml.XmlDocument]$Document,
    [System.Xml.XmlElement]$Anchor,
    [string]$Name,
    [string]$HOffset,
    [string]$VOffset,
    [string]$Cx,
    [string]$Cy
  )

  $script:DocPrIdSeed++
  $script:RelativeHeightSeed += 1024

  Set-DrawingIds -Anchor $Anchor -DocPrId $script:DocPrIdSeed -Name $Name
  Set-OrCreateAttribute -Element $Anchor -Prefix '' -LocalName 'relativeHeight' -NamespaceUri '' -Value ([string]$script:RelativeHeightSeed)
  Set-OrCreateAttribute -Element $Anchor -Prefix 'wp14' -LocalName 'anchorId' -NamespaceUri 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing' -Value (New-HexId)
  Set-OrCreateAttribute -Element $Anchor -Prefix 'wp14' -LocalName 'editId' -NamespaceUri 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing' -Value (New-HexId)

  Set-AnchorVisibility -Anchor $Anchor -BehindDoc $false
  Set-AnchorPosition -Document $Document -Anchor $Anchor -HOffset $HOffset -VOffset $VOffset
  Set-AnchorSize -Anchor $Anchor -Cx $Cx -Cy $Cy
}

function Clone-RunToParagraph {
  param(
    [System.Xml.XmlDocument]$Document,
    [System.Xml.XmlElement]$SourceRun,
    [System.Xml.XmlElement]$TargetParagraph
  )

  $clone = $Document.ImportNode($SourceRun, $true)
  [void]$TargetParagraph.AppendChild($clone)
  return $clone
}

function Get-AnchorByRid {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Rid
  )

  $node = $Document.SelectSingleNode("//*[local-name()='blip' and @*[local-name()='embed' and .='$Rid']]/ancestor::*[local-name()='anchor'][1]")
  if (-not $node) {
    throw "Could not find anchor for $Rid."
  }
  return [System.Xml.XmlElement]$node
}

function Get-AnchorByDocPrId {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Id
  )

  $node = $Document.SelectSingleNode("//*[local-name()='anchor'][./*[local-name()='docPr' and @id='$Id']]")
  if (-not $node) {
    throw "Could not find anchor with docPr id $Id."
  }
  return [System.Xml.XmlElement]$node
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

  $anchor = Get-AnchorByRid -Document $Document -Rid $Rid
  Set-AnchorPosition -Document $Document -Anchor $anchor -HOffset $HOffset -VOffset $VOffset
  Set-AnchorSize -Anchor $anchor -Cx $Cx -Cy $Cy
  Set-AnchorVisibility -Anchor $anchor -BehindDoc $false
}

function Update-AnchorByDocPrId {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Id,
    [string]$HOffset,
    [string]$VOffset,
    [string]$Cx
  )

  $anchor = Get-AnchorByDocPrId -Document $Document -Id $Id
  Set-AnchorPosition -Document $Document -Anchor $anchor -HOffset $HOffset -VOffset $VOffset
  $extent = $anchor.SelectSingleNode("./*[local-name()='extent']")
  if (-not $extent) {
    throw "Anchor $Id is missing wp:extent."
  }
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

function Set-TextBoxLabel {
  param(
    [System.Xml.XmlElement]$Anchor,
    [string]$Text,
    [string]$HexColor,
    [string]$HalfPointSize
  )

  $paragraph = $Anchor.SelectSingleNode(".//*[local-name()='txbxContent']/*[local-name()='p'][1]")
  if (-not $paragraph) {
    throw 'Text box anchor is missing its paragraph.'
  }

  $paragraph.InnerXml = @"
<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:jc w:val="center" />
  <w:rPr>
    <w:rFonts w:ascii="Montserrat ExtraBold" w:hAnsi="Montserrat ExtraBold" />
    <w:color w:val="$HexColor" />
    <w:sz w:val="$HalfPointSize" />
    <w:szCs w:val="$HalfPointSize" />
  </w:rPr>
</w:pPr>
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr>
    <w:rFonts w:ascii="Montserrat ExtraBold" w:hAnsi="Montserrat ExtraBold" />
    <w:color w:val="$HexColor" />
    <w:sz w:val="$HalfPointSize" />
    <w:szCs w:val="$HalfPointSize" />
  </w:rPr>
  <w:t xml:space="preserve">$Text</w:t>
</w:r>
"@
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

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('tech-template-v14-' + [guid]::NewGuid().ToString('N'))
$sourceZip = Join-Path ([IO.Path]::GetTempPath()) ('tech-template-v14-src-' + [guid]::NewGuid().ToString('N') + '.zip')
$outZip = Join-Path ([IO.Path]::GetTempPath()) ('tech-template-v14-out-' + [guid]::NewGuid().ToString('N') + '.zip')

New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
  Copy-Item -LiteralPath $InputPath -Destination $sourceZip -Force
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZip, $tempRoot)

  $documentPath = Join-Path $tempRoot 'word\document.xml'
  $document = Get-XmlDocument -Path $documentPath
  $ns = New-NamespaceManager -Document $document

  $script:DocPrIdSeed = Get-MaxDocPrId -Document $document
  $script:RelativeHeightSeed = Get-MaxRelativeHeight -Document $document

  $titleParagraph = $document.SelectSingleNode("/*[local-name()='document']/*[local-name()='body']/*[local-name()='p'][1]")
  if (-not $titleParagraph) {
    throw 'Could not find the title paragraph on page 1.'
  }

  # Main publicity row: EU first, MZP second, both with the same visual weight.
  Update-AnchorByRid -Document $document -Rid 'rId8' -HOffset '760000'  -VOffset '5900000' -Cx '2642361' -Cy '387931'
  Update-AnchorByRid -Document $document -Rid 'rId9' -HOffset '3750000' -VOffset '5900000' -Cx '1500000' -Cy '387931'

  # Visual separator below the main funding logos.
  Update-AnchorByDocPrId -Document $document -Id '1214179664' -HOffset '980000' -VOffset '6550000' -Cx '3840000'

  # Cities placed in a dedicated block.
  Update-AnchorByRid -Document $document -Rid 'rId10' -HOffset '1080000' -VOffset '7820000' -Cx '461165'  -Cy '260000'
  Update-AnchorByRid -Document $document -Rid 'rId11' -HOffset '1690000' -VOffset '7820000' -Cx '780838'  -Cy '260000'

  # Use the title text box as a styling source for smaller helper labels.
  $supportLabelSourceRun = $document.SelectSingleNode("//*[local-name()='anchor'][./*[local-name()='docPr' and @id='178940851']]/ancestor::*[local-name()='r'][1]")
  if (-not $supportLabelSourceRun) {
    throw 'Could not find the source text box for the support label.'
  }
  $cityLabelRun = Clone-RunToParagraph -Document $document -SourceRun $supportLabelSourceRun -TargetParagraph $titleParagraph
  $cityLabelAnchor = [System.Xml.XmlElement]$cityLabelRun.SelectSingleNode(".//*[local-name()='anchor']")
  Prepare-AnchorClone -Document $document -Anchor $cityLabelAnchor -Name 'City label' -HOffset '970000' -VOffset '7440000' -Cx '1650000' -Cy '240000'
  Set-TextBoxLabel -Anchor $cityLabelAnchor -Text 'Města projektu' -HexColor '233848' -HalfPointSize '22'

  $supportLabelRun = Clone-RunToParagraph -Document $document -SourceRun $supportLabelSourceRun -TargetParagraph $titleParagraph
  $supportLabelAnchor = [System.Xml.XmlElement]$supportLabelRun.SelectSingleNode(".//*[local-name()='anchor']")
  Prepare-AnchorClone -Document $document -Anchor $supportLabelAnchor -Name 'Support label' -HOffset '3330000' -VOffset '7440000' -Cx '1650000' -Cy '240000'
  Set-TextBoxLabel -Anchor $supportLabelAnchor -Text 'Za podpory' -HexColor '233848' -HalfPointSize '22'

  # Support logos: Dopravni podnik + UDIMO in a clearly separated block.
  $supportLogoSpecs = @(
    @{
      Rid = 'rId19'
      Name = 'Transport company support logo'
      H = '3360000'
      V = '7820000'
      Cx = '902000'
      Cy = '260000'
    },
    @{
      Rid = 'rId20'
      Name = 'UDIMO support logo'
      H = '4350000'
      V = '7820000'
      Cx = '814000'
      Cy = '260000'
    }
  )

  foreach ($spec in $supportLogoSpecs) {
    $sourceRun = $document.SelectSingleNode("//*[local-name()='blip' and @*[local-name()='embed' and .='$($spec.Rid)']]/ancestor::*[local-name()='r'][1]")
    if (-not $sourceRun) {
      throw "Could not find the source run for $($spec.Rid)."
    }

    $clonedRun = Clone-RunToParagraph -Document $document -SourceRun $sourceRun -TargetParagraph $titleParagraph
    $clonedAnchor = [System.Xml.XmlElement]$clonedRun.SelectSingleNode(".//*[local-name()='anchor']")
    Prepare-AnchorClone -Document $document -Anchor $clonedAnchor -Name $spec.Name -HOffset $spec.H -VOffset $spec.V -Cx $spec.Cx -Cy $spec.Cy
  }

  $document.Save($documentPath)

  foreach ($headerName in @('header3.xml', 'header4.xml', 'header5.xml')) {
    $headerPath = Join-Path $tempRoot ('word\' + $headerName)
    if (Test-Path -LiteralPath $headerPath) {
      $headerDoc = Get-XmlDocument -Path $headerPath
      Remove-HeaderImageRun -Document $headerDoc -Rid 'rId1'
      $headerDoc.Save($headerPath)
    }
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
