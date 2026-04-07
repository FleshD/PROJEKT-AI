param(
  [string]$OutPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\medikit_cr_tiskove_zpravy.xlsx'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ExcelColumnName([int]$Index) {
  $name = ''
  while ($Index -gt 0) {
    $Index--
    $name = [char](65 + ($Index % 26)) + $name
    $Index = [math]::Floor($Index / 26)
  }
  return $name
}

function Escape-XmlText([string]$Text) {
  if ($null -eq $Text) { return '' }
  $escaped = $Text
  $escaped = $escaped.Replace('&', '&amp;')
  $escaped = $escaped.Replace('<', '&lt;')
  $escaped = $escaped.Replace('>', '&gt;')
  $escaped = $escaped.Replace('"', '&quot;')
  $escaped = $escaped.Replace("`r", '')
  $escaped = $escaped.Replace("`n", '&#10;')
  return $escaped
}

function Repair-Mojibake([string]$Text) {
  if ($null -eq $Text) { return '' }
  $cp1250 = [System.Text.Encoding]::GetEncoding(1250)
  $utf8 = [System.Text.Encoding]::UTF8
  return $utf8.GetString($cp1250.GetBytes($Text))
}

$headers = @(
  'Vydavatel',
  'Titul',
  'Typ',
  'Web',
  'Kam poslat TZ / redakční vstup',
  'Placená spolupráce / obchod',
  'Adresa / sídlo',
  'Veřejná cena inzerce',
  'Veřejná cena PR / native',
  'Poznámka',
  'Zdroj'
)

$rows = @(
  @(
    'Economia',
    'Hospodářské noviny',
    'Deník, print',
    'https://hn.cz',
    'Centrální redakční e-mail není v ceníku veřejně uveden; pro komerční výstupy používej obchodní kontakt.',
    "inzerce@economia.cz`n+420 233 073 169",
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    "1/1 438 900 Kč`n1/2A 251 900 Kč`n1/4A 146 650 Kč`nJunior page A 328 900 Kč",
    'na dotaz; kombinovaná nabídka print + digitál se připravuje individuálně',
    'Uzávěrka 5 pracovních dní před zveřejněním. Sponzorství tématu: 329 000 Kč.',
    "https://www.economia.cz/wp-content/uploads/2026/03/Economia_Cenik_PRINT_01-2026_CZ_2.pdf`nhttps://ad.economia.cz/`nhttps://www.economia.cz/"
  ),
  @(
    'Economia',
    'HN.cz / HN BeNative',
    'Web, native studio',
    'https://hn.cz',
    'Redakční tipy podle rubrik; pro placené výstupy přímo obchod.',
    'inzerce@economia.cz',
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    'display ceník veřejně neuveden',
    "Textový výstup 80 000 Kč`nPrint + online 150 000 Kč`nText v printu 120 000 Kč`nAdvertorial 100 000 Kč`nPodcast 90 000 Kč`nMicrosite 350 000 Kč`nTematický speciál od 500 000 Kč",
    'Silná varianta pro garantované uveřejnění branded obsahu v HN.',
    'https://www.economia.cz/wp-content/uploads/2026/02/Economia_Cenik_SPEC-online_01-2026_CZ-1.pdf'
  ),
  @(
    'Economia',
    'Aktuálně.cz',
    'Web, obsahové projekty',
    'https://www.aktualne.cz',
    'Centrální redakční e-mail na ceníku není; pro placenou publikaci obchodní kontakt.',
    'inzerce@economia.cz',
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    'na dotaz',
    "Branded video / podcast / text 150 000 Kč`nObsahové speciály od 150 000 Kč / rok",
    'Vhodné pro tematická partnerství a dlouhodobé branded content projekty.',
    'https://www.economia.cz/wp-content/uploads/2026/02/Economia_Cenik_SPEC-online_01-2026_CZ-1.pdf'
  ),
  @(
    'Economia',
    'Ekonom',
    'Týdeník, print',
    'https://www.ekonom.cz',
    'Redakční kontakt centrálně neuveden; pro PR/native využij obchodní kontakt.',
    'inzerce@economia.cz',
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    "1/1 240 900 Kč`n2/1 328 900 Kč`n1/2A 144 837 Kč`n1/4A 89 536 Kč",
    'na dotaz pro čistě redakční PR; BeNative online + print: 100 000 Kč',
    'Vedle printu má i samostatnou BeNative nabídku pro online / kombinované výstupy.',
    "https://www.economia.cz/wp-content/uploads/2026/03/Economia_Cenik_PRINT_01-2026_CZ_2.pdf`nhttps://www.economia.cz/wp-content/uploads/2026/02/Economia_Cenik_SPEC-online_01-2026_CZ-1.pdf"
  ),
  @(
    'Economia',
    'Ekonom.cz / BeNative',
    'Web, native studio',
    'https://www.ekonom.cz',
    'Redakční / placené formáty řeší obchodní kontakt vydavatelství.',
    'inzerce@economia.cz',
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    'display veřejně neuveden',
    "Text + print 100 000 Kč`nRedakční speciál 250 000 Kč`nDebata 150 000 Kč`nSponzoring podcastu 80 000 Kč`nPodcast Business 90 000 Kč",
    'Dobrá volba pro B2B, management, investice, corporate leadership.',
    'https://www.economia.cz/wp-content/uploads/2026/02/Economia_Cenik_SPEC-online_01-2026_CZ-1.pdf'
  ),
  @(
    'Economia',
    'Právní rádce',
    'Měsíční příloha, print',
    'https://www.economia.cz/ceniky-inzerce/',
    'Redakční kontakt centrálně neuveden; placené výstupy přes Economia.',
    'inzerce@economia.cz',
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    "1/1 119 166 Kč`n1/2 62 700 Kč`n1/3 46 200 Kč`n1/4 34 050 Kč",
    'na dotaz',
    'Příloha míří na management a právní publikum. Hodí se pro legal, compliance, HR a poradenství.',
    'https://www.economia.cz/wp-content/uploads/2026/03/Economia_Cenik_PRINT_01-2026_CZ_2.pdf'
  ),
  @(
    'Economia',
    'Logistika',
    'Oborový časopis + newsletter',
    'https://www.economia.cz/ceniky-inzerce/',
    'Redakční kontakt centrálně neuveden; placené výstupy přes obchod Economia.',
    'inzerce@economia.cz',
    'Economia, a.s., PORT 7, Pod Dráhou 1637/2, 170 00 Praha 7 - Holešovice',
    "1/1 155 760 Kč`n1/2A 45 936 Kč`n1/3A 31 416 Kč`n1/4A 26 268 Kč`n2. strana obálky 89 760 Kč",
    'na dotaz',
    'Veřejně je uveden i newsletter Logistika; termíny pro dodání podkladů jsou zvlášť.',
    'https://www.economia.cz/wp-content/uploads/2026/03/Economia_Cenik_PRINT_01-2026_CZ_2.pdf'
  ),
  @(
    'Czech News Center',
    'Blesk / Blesk.cz',
    'Web + print brand',
    'https://www.blesk.cz',
    'tip@blesk.cz',
    "info@cncenter.cz`npodpora@cncenter.cz",
    'CZECH NEWS CENTER a. s., náměstí Marie Schmolkové 3493/1, 100 00 Praha 10',
    "Blesk homepage branding 560-1 010 Kč CPT`nBlesk floating branding 280-510 Kč CPT",
    "PR článek 30 000 Kč`nSmarticle 200 Kč CPT`nNativní rectangle 180-330 Kč CPT",
    'Silný reach. U cílených webů CNC se native/display často prodává podle konkrétního webu nebo packu.',
    "https://www.blesk.cz/kontakty`nhttps://www.cncenter.cz/`nhttps://www.cncenter.cz/inzerce`nhttps://mxfbijamra.eu-central-1.awsapprunner.com/assets/2a23e4a6-54f4-4a54-a7a9-9ac99a37132e.pdf"
  ),
  @(
    'Czech News Center',
    'Reflex / Reflex.cz',
    'Týdeník + web',
    'https://www.reflex.cz',
    "rxonline@reflex.cz`n+420 253 253 553",
    'info@cncenter.cz',
    'CZECH NEWS CENTER a. s., náměstí Marie Schmolkové 3493/1, 100 00 Praha 10',
    "Cílený branding dle webu 300-540 Kč CPT`nNativní PR premium 100-180 Kč CPT",
    'PR článek 30 000 Kč',
    'Reflex je v CNC i v mužském a zpravodajském packu, lze ho koupit samostatně i v kombinacích.',
    "https://www.reflex.cz/kontakty`nhttps://www.cncenter.cz/inzerce`nhttps://mxfbijamra.eu-central-1.awsapprunner.com/assets/2a23e4a6-54f4-4a54-a7a9-9ac99a37132e.pdf"
  ),
  @(
    'Czech News Center',
    'e15.cz',
    'Web',
    'https://www.e15.cz',
    'Veřejný centrální redakční e-mail se mi nepodařilo na oficiální kontaktní stránce strojově potvrdit; pro jistou placenou publikaci použij obchodní kontakt.',
    'info@cncenter.cz',
    'CZECH NEWS CENTER a. s., náměstí Marie Schmolkové 3493/1, 100 00 Praha 10',
    "Branding 400-720 Kč CPT`nBillboard bottom 340-620 Kč CPT`nNativní rectangle 260-470 Kč CPT",
    'PR článek 50 000 Kč',
    'Prakticky jeden z nejčitelnějších business-news online titulů s veřejným online ceníkem v CNC.',
    "https://www.cncenter.cz/inzerce`nhttps://mxfbijamra.eu-central-1.awsapprunner.com/assets/2a23e4a6-54f4-4a54-a7a9-9ac99a37132e.pdf"
  ),
  @(
    'Czech News Center',
    'iSport.cz',
    'Web',
    'https://isport.blesk.cz',
    'Veřejný centrální redakční e-mail nebyl v oficiálním ceníku / kontaktech jednoduše potvrzen; pro placenou publikaci použij obchod.',
    'info@cncenter.cz',
    'CZECH NEWS CENTER a. s., náměstí Marie Schmolkové 3493/1, 100 00 Praha 10',
    "Branding 470-850 Kč CPT`nBillboard bottom 330-600 Kč CPT`nNativní rectangle 280-510 Kč CPT",
    'PR článek 30 000 Kč',
    'Dobré pro sport, betting, fan engagement a eventy.',
    "https://www.cncenter.cz/inzerce`nhttps://mxfbijamra.eu-central-1.awsapprunner.com/assets/2a23e4a6-54f4-4a54-a7a9-9ac99a37132e.pdf"
  ),
  @(
    'N Media',
    'Deník N',
    'Web, podcasty, newslettery',
    'https://www.denikn.cz',
    'info@denikn.cz',
    "tamara.lozincakova@denikn.cz`nkarolina.planickova@denikn.cz`n+420 777 287 710",
    'N Media, a.s., Spálená 29, Praha 1',
    "Display bannery cca 24 500-50 000 Kč`nSponzoring podcastu 44 000 Kč`nAplikace Minuta po 99 000-129 000 Kč",
    "PR článek 40 000 Kč`nOdemknutí článku od 25 000 Kč",
    'Na inzerci mají rovnou veřejný online ceník a kontakty. Cenová hladina je z PDF platného od 1. 1. 2025, které je na stránce stále odkazované k 26. 3. 2026.',
    "https://www.denikn.cz/inzerce/`nhttps://static.novydenik.com/2024/12/cenik25.pdf"
  ),
  @(
    'MediaRey',
    'Forbes.cz',
    'Web, newslettery, branded studio',
    'https://www.forbes.cz',
    'redakce@forbes.cz',
    "obchod@forbes.cz`n+420 228 225 093",
    'MediaRey, SE (právní subjekt uvedený na stránce inzerce; poštovní adresa na téže stránce není výslovně rozepsaná)',
    "Leaderboard 900 Kč CPT`nBranding 900 Kč CPT`nSkyscraper sticky 750 Kč CPT`nNative Ad 250 Kč CPT",
    "Advoice 90 000 Kč`nAdvoice Plus 140 000 Kč`nBrandvoice 200 000 Kč`nBrandvoice Podcast 180 000 Kč`nSpeciální projekt od 500 000 Kč",
    'Forbes má jeden z nejčitelnějších veřejných webových mediakitů v CZ včetně native / brandvoice cen.',
    "https://www.forbes.cz/inzerce/`nhttps://cdn.forbes.cz/uploads/2026/03/WEB_kit_2026-CZ-DEF-09032026.pdf"
  ),
  @(
    'ČTK / Protext',
    'Protext ČTK',
    'Newswire / distribuce TZ',
    'https://www.protext.cz/',
    "protext@ctk.cz`n+420 222 098 175",
    'protext@ctk.cz',
    'ČTK; na webu ČTK je u nemovitostí zmiňována budova Opletalova, Praha 1 (přesná adresa Protextu není na homepage rozepsaná)',
    'ceník veřejně nezobrazen',
    'na dotaz',
    'Nejde o redakci, ale o nejpraktičtější legální distribuční kanál pro zveřejnění a rozeslání tiskové zprávy do médií.',
    "https://www.protext.cz/`nhttps://www.ctk.cz/"
  ),
  @(
    'Mediář',
    'Mediář',
    'Oborový web',
    'https://www.mediar.cz',
    'potucek@mediar.cz',
    "ondrej@aust.cz`npotucek@mediar.cz",
    'Na stránce inzerce není poštovní adresa veřejně rozepsaná.',
    'na dotaz',
    'na dotaz',
    'Hodí se pro B2B komunikaci v médiích, reklamě, marketingu a publishingu. Veřejně uvádí možnosti inzerce a komerční prezentace.',
    'https://www.mediar.cz/inzerce'
  )
)

$headers = @($headers | ForEach-Object { Repair-Mojibake $_ })
$fixedRows = @()
foreach ($row in $rows) {
  $fixedRows += ,@($row | ForEach-Object { Repair-Mojibake $_ })
}
$rows = $fixedRows

$tempDir = Join-Path ([IO.Path]::GetTempPath()) ('medikit-xlsx-' + [guid]::NewGuid().ToString('N'))
$xlDir = Join-Path $tempDir 'xl'
$relsDir = Join-Path $tempDir '_rels'
$xlRelsDir = Join-Path $xlDir '_rels'
$wsDir = Join-Path $xlDir 'worksheets'

New-Item -ItemType Directory -Path $tempDir | Out-Null
New-Item -ItemType Directory -Path $relsDir | Out-Null
New-Item -ItemType Directory -Path $xlDir | Out-Null
New-Item -ItemType Directory -Path $xlRelsDir | Out-Null
New-Item -ItemType Directory -Path $wsDir | Out-Null

try {
  $lastCol = Get-ExcelColumnName $headers.Count
  $lastRow = $rows.Count + 1
  $dimension = "A1:$lastCol$lastRow"
  $autoFilter = $dimension
  $widths = @(18, 26, 22, 22, 34, 30, 34, 28, 28, 34, 30)

  $sheetBuilder = New-Object System.Text.StringBuilder
  [void]$sheetBuilder.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sheetBuilder.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
  [void]$sheetBuilder.AppendLine("  <dimension ref=""$dimension""/>")
  [void]$sheetBuilder.AppendLine('  <sheetViews>')
  [void]$sheetBuilder.AppendLine('    <sheetView workbookViewId="0">')
  [void]$sheetBuilder.AppendLine('      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
  [void]$sheetBuilder.AppendLine('      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
  [void]$sheetBuilder.AppendLine('    </sheetView>')
  [void]$sheetBuilder.AppendLine('  </sheetViews>')
  [void]$sheetBuilder.AppendLine('  <sheetFormatPr defaultRowHeight="18"/>')
  [void]$sheetBuilder.AppendLine('  <cols>')
  for ($i = 0; $i -lt $widths.Count; $i++) {
    $colIdx = $i + 1
    [void]$sheetBuilder.AppendLine("    <col min=""$colIdx"" max=""$colIdx"" width=""$($widths[$i])"" customWidth=""1""/>")
  }
  [void]$sheetBuilder.AppendLine('  </cols>')
  [void]$sheetBuilder.AppendLine('  <sheetData>')

  for ($rowIndex = 1; $rowIndex -le $lastRow; $rowIndex++) {
    if ($rowIndex -eq 1) {
      [void]$sheetBuilder.AppendLine('    <row r="1" ht="34" customHeight="1">')
      for ($colIndex = 1; $colIndex -le $headers.Count; $colIndex++) {
        $cellRef = "$(Get-ExcelColumnName $colIndex)$rowIndex"
        $value = Escape-XmlText $headers[$colIndex - 1]
        [void]$sheetBuilder.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""1""><is><t xml:space=""preserve"">$value</t></is></c>")
      }
      [void]$sheetBuilder.AppendLine('    </row>')
      continue
    }

    [void]$sheetBuilder.AppendLine("    <row r=""$rowIndex"">")
    $rowData = $rows[$rowIndex - 2]
    for ($colIndex = 1; $colIndex -le $headers.Count; $colIndex++) {
      $cellRef = "$(Get-ExcelColumnName $colIndex)$rowIndex"
      $value = Escape-XmlText $rowData[$colIndex - 1]
      [void]$sheetBuilder.AppendLine("      <c r=""$cellRef"" t=""inlineStr"" s=""2""><is><t xml:space=""preserve"">$value</t></is></c>")
    }
    [void]$sheetBuilder.AppendLine('    </row>')
  }

  [void]$sheetBuilder.AppendLine('  </sheetData>')
  [void]$sheetBuilder.AppendLine("  <autoFilter ref=""$autoFilter""/>")
  [void]$sheetBuilder.AppendLine('</worksheet>')

  $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF36393C"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'@

  $workbookXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Medikit CZ" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@

  $rootRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

  $workbookRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

  $contentTypesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@

  $utf8 = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText((Join-Path $tempDir '[Content_Types].xml'), $contentTypesXml, $utf8)
  [System.IO.File]::WriteAllText((Join-Path $relsDir '.rels'), $rootRelsXml, $utf8)
  [System.IO.File]::WriteAllText((Join-Path $xlDir 'workbook.xml'), $workbookXml, $utf8)
  [System.IO.File]::WriteAllText((Join-Path $xlRelsDir 'workbook.xml.rels'), $workbookRelsXml, $utf8)
  [System.IO.File]::WriteAllText((Join-Path $xlDir 'styles.xml'), $stylesXml, $utf8)
  [System.IO.File]::WriteAllText((Join-Path $wsDir 'sheet1.xml'), $sheetBuilder.ToString(), $utf8)

  if (Test-Path $OutPath) {
    Remove-Item -Path $OutPath -Force
  }

  $zip = [System.IO.Compression.ZipFile]::Open($OutPath, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($file in Get-ChildItem -Path $tempDir -Recurse -File) {
      $relative = $file.FullName.Substring($tempDir.Length).TrimStart('\')
      $entryName = $relative.Replace('\', '/')
      [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $entryName)
    }
  } finally {
    $zip.Dispose()
  }

  [pscustomobject]@{
    Path = $OutPath
    Rows = $rows.Count
    Columns = $headers.Count
  } | ConvertTo-Json -Depth 3
}
finally {
  if (Test-Path $tempDir) {
    Remove-Item -Path $tempDir -Recurse -Force
  }
}
