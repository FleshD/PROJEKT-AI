param(
  [string]$AnalysisDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\analysis_josephine_sk',
  [string]$XlsxPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\analysis_josephine_sk\JOSEPHINE_konkurence_SK.xlsx',
  [string]$PdfPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\analysis_josephine_sk\JOSEPHINE_konkurence_SK.pdf'
)

$ErrorActionPreference = 'Stop'

function Escape-Xml {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  return [Security.SecurityElement]::Escape([string]$Text)
}

function Get-ColName {
  param([int]$Index)
  $name = ''
  while ($Index -gt 0) {
    $remainder = ($Index - 1) % 26
    $name = [char](65 + $remainder) + $name
    $Index = [math]::Floor(($Index - 1) / 26)
  }
  return $name
}

function New-CellSpec {
  param(
    $Value,
    [string]$Kind = 's',
    [int]$Style = 0
  )
  return [pscustomobject]@{
    Value = $Value
    Kind = $Kind
    Style = $Style
  }
}

function Normalize-NumberText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
  $clean = ([string]$Text).Replace(([string][char]0xA0), '').Trim() -replace '\s', ''
  if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
  return $clean.Replace(',', '.')
}

function New-CellXml {
  param(
    [int]$ColumnIndex,
    [int]$RowIndex,
    [pscustomobject]$Cell
  )

  $ref = (Get-ColName $ColumnIndex) + $RowIndex
  $styleAttr = ' s="' + $Cell.Style + '"'

  if ($null -eq $Cell.Value -or [string]$Cell.Value -eq '') {
    return '<c r="' + $ref + '"' + $styleAttr + ' />'
  }

  if ($Cell.Kind -eq 'n') {
    return '<c r="' + $ref + '"' + $styleAttr + '><v>' + [string]$Cell.Value + '</v></c>'
  }

  $escaped = Escape-Xml ([string]$Cell.Value)
  return '<c r="' + $ref + '"' + $styleAttr + ' t="inlineStr"><is><t xml:space="preserve">' + $escaped + '</t></is></c>'
}

function New-WorksheetXml {
  param(
    [object[]]$Rows,
    [double[]]$Widths,
    [int]$FreezeRows = 0,
    [string]$AutoFilterRef = ''
  )

  $rowCount = [math]::Max(1, $Rows.Count)
  $maxCols = 1
  foreach ($row in $Rows) {
    if ($row.Count -gt $maxCols) { $maxCols = $row.Count }
  }

  $dimensionRef = 'A1:' + (Get-ColName $maxCols) + $rowCount
  $colsSb = New-Object System.Text.StringBuilder
  if ($Widths -and $Widths.Count -gt 0) {
    [void]$colsSb.Append('<cols>')
    for ($i = 0; $i -lt $Widths.Count; $i++) {
      [void]$colsSb.Append('<col min="' + ($i + 1) + '" max="' + ($i + 1) + '" width="' + ([string]$Widths[$i]).Replace(',', '.') + '" customWidth="1"/>')
    }
    [void]$colsSb.Append('</cols>')
  }

  $sheetViewsXml = if ($FreezeRows -gt 0) {
    $topLeftRow = $FreezeRows + 1
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="' + $FreezeRows + '" topLeftCell="A' + $topLeftRow + '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A' + $topLeftRow + '" sqref="A' + $topLeftRow + '"/></sheetView></sheetViews>'
  } else {
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
  }

  $rowsSb = New-Object System.Text.StringBuilder
  if ($Rows.Count -eq 0) {
    [void]$rowsSb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve"></t></is></c></row>')
  } else {
    for ($r = 0; $r -lt $Rows.Count; $r++) {
      $rowIndex = $r + 1
      [void]$rowsSb.Append('<row r="' + $rowIndex + '">')
      $cells = @($Rows[$r])
      for ($c = 0; $c -lt $cells.Count; $c++) {
        [void]$rowsSb.Append((New-CellXml -ColumnIndex ($c + 1) -RowIndex $rowIndex -Cell $cells[$c]))
      }
      [void]$rowsSb.Append('</row>')
    }
  }

  $autoFilterXml = if ($AutoFilterRef) { '<autoFilter ref="' + $AutoFilterRef + '"/>' } else { '' }

  return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
__DIMENSION__
__SHEETVIEWS__
<sheetFormatPr defaultRowHeight="18"/>
__COLS__
<sheetData>
__ROWS__
</sheetData>
__AUTOFILTER__
<pageMargins left="0.3" right="0.3" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>
</worksheet>
'@.Replace('__DIMENSION__', '<dimension ref="' + $dimensionRef + '"/>').
    Replace('__SHEETVIEWS__', $sheetViewsXml).
    Replace('__COLS__', $colsSb.ToString()).
    Replace('__ROWS__', $rowsSb.ToString()).
    Replace('__AUTOFILTER__', $autoFilterXml)
}

function Write-Utf8NoBom {
  param(
    [string]$Path,
    [string]$Content
  )
  $enc = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Get-BrowserPath {
  $candidates = @(
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
  )
  foreach ($path in $candidates) {
    if (Test-Path $path) { return $path }
  }
  throw 'No compatible browser found for PDF export.'
}

function Get-OpenXmlDllPath {
  $candidates = @(
    'C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\DocumentFormat.OpenXml.dll',
    'C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\Filters\Documentformat.OpenXml.dll',
    'C:\Program Files\Windows Defender Advanced Threat Protection\Classification\Dprt\DocumentFormat.OpenXml.dll'
  )
  foreach ($path in $candidates) {
    if (Test-Path $path) { return $path }
  }
  return $null
}

$summaryCsv = Join-Path $AnalysisDir 'summary.csv'
$contractsCsv = Join-Path $AnalysisDir 'contracts_raw.csv'
$reportHtml = Join-Path $AnalysisDir 'report.html'

if (-not (Test-Path $summaryCsv)) { throw 'summary.csv not found.' }
if (-not (Test-Path $contractsCsv)) { throw 'contracts_raw.csv not found.' }
if (-not (Test-Path $reportHtml)) { throw 'report.html not found.' }

$summary = @(Import-Csv -Delimiter ';' $summaryCsv)
$raw = @(Import-Csv -Delimiter ';' $contractsCsv)

$josephine = $summary | Where-Object system -eq 'JOSEPHINE' | Select-Object -First 1
$topByCount = @($summary | Select-Object -First 3)
$totalContracts = (($summary | ForEach-Object { [int](Normalize-NumberText $_.contract_count) }) | Measure-Object -Sum).Sum

$summaryRows = @()
$summaryRows += ,@((New-CellSpec 'Analyza konkurence JOSEPHINE na Slovensku' 's' 1))
$summaryRows += ,@((New-CellSpec 'Vygenerovano' 's' 1), (New-CellSpec (Get-Date -Format 'yyyy-MM-dd HH:mm') 's' 0))
$summaryRows += ,@((New-CellSpec 'Primarni zdroj' 's' 1), (New-CellSpec 'CRZ Slovakia + UVO zoznam zapisanych elektronickych prostriedkov' 's' 4))
$summaryRows += ,@((New-CellSpec 'Smluv v datasetu' 's' 1), (New-CellSpec (Normalize-NumberText $totalContracts) 'n' 2))
$summaryRows += ,@()
$summaryRows += ,@(
  (New-CellSpec 'System' 's' 1),
  (New-CellSpec 'Pocet smluv' 's' 1),
  (New-CellSpec 'Zverejnene hodnoty' 's' 1),
  (New-CellSpec 'Nulove hodnoty' 's' 1),
  (New-CellSpec 'Soucet v EUR s DPH' 's' 1),
  (New-CellSpec 'Prumer v EUR s DPH' 's' 1),
  (New-CellSpec 'Unikatni zadavatele' 's' 1),
  (New-CellSpec 'Licence / pristup' 's' 1),
  (New-CellSpec 'Podpora / servis' 's' 1)
)

foreach ($row in $summary) {
  $summaryRows += ,@(
    (New-CellSpec $row.system 's' 0),
    (New-CellSpec (Normalize-NumberText $row.contract_count) 'n' 2),
    (New-CellSpec (Normalize-NumberText $row.disclosed_value_count) 'n' 2),
    (New-CellSpec (Normalize-NumberText $row.zero_value_count) 'n' 2),
    (New-CellSpec (Normalize-NumberText $row.total_disclosed_value_eur_vat) 'n' 3),
    (New-CellSpec (Normalize-NumberText $row.avg_disclosed_value_eur_vat) 'n' 5),
    (New-CellSpec (Normalize-NumberText $row.unique_public_buyers) 'n' 2),
    (New-CellSpec (Normalize-NumberText $row.license_like_contracts) 'n' 2),
    (New-CellSpec (Normalize-NumberText $row.support_contracts) 'n' 2)
  )
}

$summaryRows += ,@()
$summaryRows += ,@((New-CellSpec 'Klicove zavery' 's' 1))
$summaryRows += ,@((New-CellSpec ('JOSEPHINE: ' + $josephine.contract_count + ' smluv, ' + $josephine.unique_public_buyers + ' unikatnich zadavatelu, ' + $josephine.total_disclosed_value_eur_vat + ' EUR ve zverejnenych hodnotach CRZ.') 's' 4))
$summaryRows += ,@((New-CellSpec ('Nejsilnejsi stopa dle poctu smluv: ' + $topByCount[0].system + ' (' + $topByCount[0].contract_count + '), ' + $topByCount[1].system + ' (' + $topByCount[1].contract_count + '), ' + $topByCount[2].system + ' (' + $topByCount[2].contract_count + ').') 's' 4))
$summaryRows += ,@((New-CellSpec ('JOSEPHINE ma ' + $josephine.license_like_contracts + ' smluv licencniho / pristupoveho typu a ' + $josephine.support_contracts + ' smluv se servisnim nebo podpurnym charakterem.') 's' 4))
$summaryRows += ,@((New-CellSpec 'Hodnotove soucty vychazeji z tabulkovych castok CRZ, ktere jsou v tomto registru publikovane v EUR s DPH.' 's' 4))

$rawHeaders = @(
  'system',
  'vendor',
  'market_class',
  'buyer',
  'supplier',
  'description',
  'contract_number',
  'published_date_text',
  'contract_date',
  'year',
  'value_text',
  'value_eur_vat',
  'detail_url',
  'detail_id',
  'mentions_license',
  'mentions_support',
  'mentions_archive',
  'mentions_auction',
  'mentions_tenderbox',
  'mentions_workflow',
  'mentions_wendy',
  'explicit_license_count'
)

$rawHeaderLabels = @(
  'System',
  'Dodavatel',
  'Trzni_trida',
  'Zadavatel',
  'Dodavatel_ve_smlouve',
  'Popis_smlouvy',
  'Cislo_smlouvy',
  'Datum_publikace_text',
  'Datum',
  'Rok',
  'Hodnota_text',
  'Hodnota_EUR_s_DPH',
  'Detail_URL',
  'Detail_ID',
  'Licence_typ',
  'Podpora',
  'Archiv',
  'Aukce',
  'Tenderbox',
  'Workflow',
  'Wendy',
  'Explicitni_licence'
)

$rawRows = @()
$rawRows += ,@($rawHeaderLabels | ForEach-Object { New-CellSpec $_ 's' 1 })
foreach ($row in $raw) {
  $rawRows += ,@(
    (New-CellSpec $row.system 's' 0),
    (New-CellSpec $row.vendor 's' 0),
    (New-CellSpec $row.market_class 's' 0),
    (New-CellSpec $row.buyer 's' 4),
    (New-CellSpec $row.supplier 's' 4),
    (New-CellSpec $row.description 's' 4),
    (New-CellSpec $row.contract_number 's' 0),
    (New-CellSpec $row.published_date_text 's' 0),
    (New-CellSpec $row.contract_date 's' 0),
    (New-CellSpec (Normalize-NumberText $row.year) 'n' 2),
    (New-CellSpec $row.value_text 's' 4),
    (New-CellSpec (Normalize-NumberText $row.value_eur_vat) 'n' 3),
    (New-CellSpec $row.detail_url 's' 4),
    (New-CellSpec $row.detail_id 's' 0),
    (New-CellSpec $row.mentions_license 's' 0),
    (New-CellSpec $row.mentions_support 's' 0),
    (New-CellSpec $row.mentions_archive 's' 0),
    (New-CellSpec $row.mentions_auction 's' 0),
    (New-CellSpec $row.mentions_tenderbox 's' 0),
    (New-CellSpec $row.mentions_workflow 's' 0),
    (New-CellSpec $row.mentions_wendy 's' 0),
    (New-CellSpec (Normalize-NumberText $row.explicit_license_count) 'n' 2)
  )
}

$noteRows = @()
$noteRows += ,@((New-CellSpec 'Metodika a omezeni' 's' 1))
$noteRows += ,@((New-CellSpec 'Ramec trhu vychazi z oficialniho seznamu zapisanych elektronickych prostriedkov na webe UVO.' 's' 4))
$noteRows += ,@((New-CellSpec 'Porovnani zahrnuje slovenske systemy JOSEPHINE, TENDERnet, ERANET, eBIZ platform, EVOSERVIS a ActiveProcurement.' 's' 4))
$noteRows += ,@((New-CellSpec 'Pocet smluv vychazi z verejne dohledatelnych zaznamov v slovenskem CRZ pri kombinaci dodavatele a nazvu systemu.' 's' 4))
$noteRows += ,@((New-CellSpec 'Presne pocty seatovych licenci nelze bez rucniho cteni priloh PDF spolehlive ziskat, proto analyza pracuje hlavne s poctem smluv, objemem hodnot a typem nasazeni.' 's' 4))
$noteRows += ,@((New-CellSpec 'Zdroj trhu: https://www.uvo.gov.sk/zoznam-zapisanych-elektronickych-prostriedkov' 's' 4))
$noteRows += ,@((New-CellSpec 'Zdroj dat: https://www.crz.gov.sk/central-register-of-contracts/' 's' 4))

$sheet1Xml = New-WorksheetXml -Rows $summaryRows -Widths @(28, 22, 24, 22, 18, 18, 18, 18, 18) -FreezeRows 5 -AutoFilterRef ('A6:I' + (6 + $summary.Count))
$sheet2Xml = New-WorksheetXml -Rows $rawRows -Widths @(14, 16, 16, 24, 24, 60, 16, 18, 12, 8, 18, 18, 34, 12, 12, 12, 12, 10, 10, 12, 10, 12) -FreezeRows 1 -AutoFilterRef ('A1:V' + $rawRows.Count)
$sheet3Xml = New-WorksheetXml -Rows $noteRows -Widths @(120) -FreezeRows 0

$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00 &quot;EUR&quot;"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF322F2D"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'@

$contentTypesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@

$rootRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

$workbookXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/>
  </bookViews>
  <sheets>
    <sheet name="Souhrn" sheetId="1" r:id="rId1"/>
    <sheet name="Smlouvy_raw" sheetId="2" r:id="rId2"/>
    <sheet name="Metodika" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>
'@

$workbookRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

$coreXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Analyza konkurence JOSEPHINE na Slovensku</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">__CREATED__</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">__MODIFIED__</dcterms:modified>
</cp:coreProperties>
'@

$appXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <TitlesOfParts>
    <vt:vector size="3" baseType="lpstr">
      <vt:lpstr>Souhrn</vt:lpstr>
      <vt:lpstr>Smlouvy_raw</vt:lpstr>
      <vt:lpstr>Metodika</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>3</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
</Properties>
'@

$created = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$coreXml = $coreXml.Replace('__CREATED__', $created).Replace('__MODIFIED__', $created)

$tempRoot = Join-Path $env:TEMP ('josephine_xlsx_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'docProps') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\worksheets') | Out-Null

Write-Utf8NoBom -Path (Join-Path $tempRoot '[Content_Types].xml') -Content $contentTypesXml
Write-Utf8NoBom -Path (Join-Path $tempRoot '_rels\.rels') -Content $rootRelsXml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'docProps\app.xml') -Content $appXml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'docProps\core.xml') -Content $coreXml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\workbook.xml') -Content $workbookXml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') -Content $workbookRelsXml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\styles.xml') -Content $stylesXml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\worksheets\sheet1.xml') -Content $sheet1Xml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\worksheets\sheet2.xml') -Content $sheet2Xml
Write-Utf8NoBom -Path (Join-Path $tempRoot 'xl\worksheets\sheet3.xml') -Content $sheet3Xml

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path $XlsxPath) { Remove-Item $XlsxPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($XlsxPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  $files = Get-ChildItem -Path $tempRoot -Recurse -File
  $basePath = [System.IO.Path]::GetFullPath($tempRoot).TrimEnd('\')
  foreach ($file in $files) {
    $fullPath = [System.IO.Path]::GetFullPath($file.FullName)
    $relativePath = $fullPath.Substring($basePath.Length).TrimStart('\')
    $entryName = $relativePath -replace '\\', '/'
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
  }
} finally {
  $zip.Dispose()
}
Remove-Item -Path $tempRoot -Recurse -Force

$openXmlDll = Get-OpenXmlDllPath
$xlsxValidated = $false
if ($openXmlDll) {
  try {
    Add-Type -Path $openXmlDll
    $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($XlsxPath, $false)
    $null = $doc.WorkbookPart.Workbook.Sheets.Count
    $doc.Dispose()
    $xlsxValidated = $true
  } catch {
    $xlsxValidated = $false
  }
}

$browserPath = Get-BrowserPath
$reportUri = ([System.Uri]$reportHtml).AbsoluteUri
if (Test-Path $PdfPath) { Remove-Item $PdfPath -Force }

$browserArgs = @(
  '--headless=new',
  '--disable-gpu',
  '--allow-file-access-from-files',
  '--run-all-compositor-stages-before-draw',
  '--virtual-time-budget=12000',
  '--no-pdf-header-footer',
  '--print-to-pdf-no-header',
  ('--print-to-pdf=' + $PdfPath),
  $reportUri
)

$process = Start-Process -FilePath $browserPath -ArgumentList $browserArgs -Wait -PassThru -WindowStyle Hidden
if (-not (Test-Path $PdfPath)) {
  throw 'PDF export failed.'
}

[pscustomobject]@{
  xlsx = $XlsxPath
  pdf = $PdfPath
  xlsx_validated = $xlsxValidated
  browser = $browserPath
  pdf_size = (Get-Item $PdfPath).Length
  xlsx_size = (Get-Item $XlsxPath).Length
} | ConvertTo-Json -Depth 4
