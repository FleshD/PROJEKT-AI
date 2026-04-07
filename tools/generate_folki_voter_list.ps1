param(
    [string]$InputPath = "",
    [string]$OutputPath = "C:\Users\jan.jedlicka\Documents\PROJEKT_AI\FOLKI\Seznam_hlasujicich_FOLKI_2023_2026.xlsx"
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

$SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
$PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"

if (-not $InputPath) {
    $candidate = Get-ChildItem -LiteralPath "C:\Users\jan.jedlicka\Documents\PROJEKT_AI\FOLKI" -Filter "*.xlsx" -File |
        Where-Object { $_.Name -ne [System.IO.Path]::GetFileName($OutputPath) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $candidate) {
        throw "V adresáři FOLKI nebyl nalezen žádný vstupní XLSX soubor."
    }

    $InputPath = $candidate.FullName
}

function Get-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name
    )

    $entry = $Zip.GetEntry($Name)
    if (-not $entry) {
        throw "Missing XLSX entry: $Name"
    }

    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Close()
        $stream.Close()
    }
}

function Get-XlsxSharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)

    $entry = $Zip.GetEntry("xl/sharedStrings.xml")
    if (-not $entry) {
        return @( )
    }

    $xml = [xml](Get-ZipText -Zip $Zip -Name "xl/sharedStrings.xml")
    $shared = New-Object System.Collections.Generic.List[string]

    foreach ($si in $xml.GetElementsByTagName("si", $SpreadsheetNamespace)) {
        $text = ""
        foreach ($node in $si.GetElementsByTagName("t", $SpreadsheetNamespace)) {
            $text += $node.InnerText
        }
        [void]$shared.Add($text)
    }

    return $shared.ToArray()
}

function Get-ColumnName {
    param([string]$CellReference)

    return ($CellReference -replace "\d", "")
}

function Get-CellValue {
    param(
        [System.Xml.XmlElement]$Cell,
        [string[]]$SharedStrings
    )

    $type = [string]$Cell.t

    if ($type -eq "s") {
        $vNode = $Cell.GetElementsByTagName("v", $SpreadsheetNamespace) | Select-Object -First 1
        if ($vNode) {
            return $SharedStrings[[int]$vNode.InnerText]
        }
        return ""
    }

    if ($type -eq "inlineStr") {
        $text = ""
        foreach ($node in $Cell.GetElementsByTagName("t", $SpreadsheetNamespace)) {
            $text += $node.InnerText
        }
        return $text
    }

    $valueNode = $Cell.GetElementsByTagName("v", $SpreadsheetNamespace) | Select-Object -First 1
    if ($valueNode) {
        return [string]$valueNode.InnerText
    }

    return ""
}

function Get-WorkbookSheetMap {
    param([System.IO.Compression.ZipArchive]$Zip)

    $workbookXml = [xml](Get-ZipText -Zip $Zip -Name "xl/workbook.xml")
    $relsXml = [xml](Get-ZipText -Zip $Zip -Name "xl/_rels/workbook.xml.rels")

    $relMap = @{}
    foreach ($rel in $relsXml.GetElementsByTagName("Relationship", $PackageRelationshipNamespace)) {
        $relMap[[string]$rel.Id] = [string]$rel.Target
    }

    $sheets = @()
    foreach ($sheet in $workbookXml.GetElementsByTagName("sheet", $SpreadsheetNamespace)) {
        $target = $relMap[[string]$sheet.GetAttribute("id", $OfficeRelationshipNamespace)]
        if (-not $target) {
            continue
        }

        if ($target -notmatch "^xl/") {
            $target = "xl/$target"
        }

        $sheets += [pscustomobject]@{
            Name = [string]$sheet.name
            Path = $target
        }
    }

    return $sheets
}

function Read-FolkiVotes {
    param([string]$Path)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sharedStrings = Get-XlsxSharedStrings -Zip $zip
        $sheetMap = Get-WorkbookSheetMap -Zip $zip
        $rows = New-Object System.Collections.Generic.List[object]

        foreach ($sheet in $sheetMap) {
            $sheetXml = [xml](Get-ZipText -Zip $zip -Name $sheet.Path)
            $sheetRows = $sheetXml.GetElementsByTagName("row", $SpreadsheetNamespace)

            foreach ($row in ($sheetRows | Select-Object -Skip 1)) {
                $email = ""
                $name = ""

                foreach ($cell in ($row.ChildNodes | Where-Object { $_.LocalName -eq "c" })) {
                    $column = Get-ColumnName -CellReference ([string]$cell.r)
                    $value = Get-CellValue -Cell $cell -SharedStrings $sharedStrings

                    if ($column -eq "B") {
                        $email = $value
                    }
                    elseif ($column -eq "C") {
                        $name = $value
                    }
                }

                if ($email -or $name) {
                    [void]$rows.Add([pscustomobject]@{
                        Year      = [string]$sheet.Name
                        EmailRaw  = [string]$email
                        NameRaw   = [string]$name
                    })
                }
            }
        }

        return $rows.ToArray()
    }
    finally {
        $zip.Dispose()
    }
}

function Normalize-Whitespace {
    param([string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    $normalized = $Value -replace "\s+", " "
    return $normalized.Trim()
}

function Normalize-Email {
    param([string]$Email)

    $email = (Normalize-Whitespace -Value $Email).ToLowerInvariant()
    $email = $email -replace "centrum\.czw$", "centrum.cz"
    return $email
}

function Remove-NameDecorations {
    param([string]$Name)

    $value = Normalize-Whitespace -Value $Name
    $value = $value -replace "^Mgr\.\s*", ""
    $value = $value -replace ",\s*Ph\.D\.\s*$", ""
    $value = $value -replace "\s*\([^)]*\)", ""
    $value = $value -replace "\s+alias\s+.*$", ""
    $value = Normalize-Whitespace -Value $value
    return $value
}

function Convert-ToTitleCase {
    param([string]$Value)

    if (-not $Value) {
        return ""
    }

    $culture = [System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")
    $parts = $Value.Split(" ")
    $result = @()

    foreach ($part in $parts) {
        if (-not $part) {
            continue
        }

        if ($part -match '^[\p{L}-]+$') {
            $lower = $part.ToLower($culture)
            if ($lower -eq 'a') {
                $result += 'a'
                continue
            }

            $result += $culture.TextInfo.ToTitleCase($lower)
            continue
        }

        $result += $part
    }

    return ($result -join " ")
}

function Normalize-Name {
    param([string]$Name)

    $value = Remove-NameDecorations -Name $Name
    $value = $value.Trim(" ", "`"", "'")
    $value = Convert-ToTitleCase -Value $value
    return $value
}

function Get-NameScore {
    param([string]$Name)

    if (-not $Name) {
        return 0
    }

    $score = 0
    $parts = $Name.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)

    if ($parts.Count -ge 2) {
        $score += 10
    }

    $score += [Math]::Min($Name.Length, 40)

    if ($Name -match "alias|Mgr\.|Ph\.D\.|\(") {
        $score -= 5
    }

    return $score
}

function Get-PersonId {
    param(
        [string]$NormalizedEmail,
        [string]$NormalizedName
    )

    $manualPersonMap = @{
        "amarouny@centrum.cz"         = "eva-sukova"
        "evaochrymcukova@seznam.cz"   = "eva-sukova"
        "ivan.kurtev@supergram.cz"    = "ivan-kurtev"
        "ivan.kurtev@folkzije.cz"     = "ivan-kurtev"
        "sansonikamuzika@gmail.com"   = "lucie-bublava"
        "sansonika@email.cz"          = "lucie-bublava"
        "evlnyk@seznam.cz"            = "lucie-vlasakova"
        "epydemye@seznam.cz"          = "lucie-vlasakova"
        "vaclav.fajfr@gmail.com"      = "vaclav-fajfr"
        "vaclav.fajfr@seznam.cz"      = "vaclav-fajfr"
        "vondrak.jirka@gmail.com"     = "jiri-vondrak"
        "agenturajv@gmail.com"        = "jiri-vondrak"
    }

    if ($manualPersonMap.ContainsKey($NormalizedEmail)) {
        return $manualPersonMap[$NormalizedEmail]
    }

    if ($NormalizedEmail) {
        return "email:$NormalizedEmail"
    }

    return "name:$NormalizedName"
}

function Get-ComparisonKey {
    param([string]$Value)

    if (-not $Value) {
        return ""
    }

    $formD = $Value.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $formD.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    $ascii = $builder.ToString().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
    $ascii = $ascii -replace "[^a-z0-9]+", " "
    return ($ascii -replace "\s+", " ").Trim()
}

function Convert-ToExcelColumnName {
    param([int]$Number)

    $name = ""
    $current = $Number

    while ($current -gt 0) {
        $remainder = ($current - 1) % 26
        $name = [char](65 + $remainder) + $name
        $current = [math]::Floor(($current - 1) / 26)
    }

    return $name
}

function Escape-XmlText {
    param([string]$Value)

    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function New-InlineStringCell {
    param(
        [string]$Column,
        [int]$RowNumber,
        [string]$Value,
        [int]$StyleId = 0
    )

    $escaped = Escape-XmlText -Value $Value
    return "<c r=""$Column$RowNumber"" t=""inlineStr"" s=""$StyleId""><is><t xml:space=""preserve"">$escaped</t></is></c>"
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8)
}

function Add-ZipStringEntry {
    param(
        [System.IO.Compression.ZipArchive]$ZipArchive,
        [string]$EntryName,
        [string]$Content
    )

    $entry = $ZipArchive.CreateEntry($EntryName)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function New-MinimalXlsx {
    param(
        [object[]]$Rows,
        [string]$OutputPath
    )

    $headers = @(
        "Jméno a příjmení / přezdívka",
        "E-mailová adresa(y)",
        "Hlasoval",
        "Poslední hlasování"
    )

    $sheetRows = New-Object System.Collections.Generic.List[string]
    $headerCells = @()
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $headerCells += New-InlineStringCell -Column (Convert-ToExcelColumnName ($i + 1)) -RowNumber 1 -Value $headers[$i] -StyleId 1
    }
    [void]$sheetRows.Add("<row r=""1"">" + ($headerCells -join "") + "</row>")

    $rowIndex = 2
    foreach ($row in $Rows) {
        $cells = @()
        $cells += New-InlineStringCell -Column "A" -RowNumber $rowIndex -Value $row.Name
        $cells += New-InlineStringCell -Column "B" -RowNumber $rowIndex -Value $row.Emails
        $cells += New-InlineStringCell -Column "C" -RowNumber $rowIndex -Value $row.CountLabel
        $cells += New-InlineStringCell -Column "D" -RowNumber $rowIndex -Value $row.LastYear
        [void]$sheetRows.Add("<row r=""$rowIndex"">" + ($cells -join "") + "</row>")
        $rowIndex++
    }

    $dimension = "A1:D$($rowIndex - 1)"
    $sheetXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="$dimension"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="28" customWidth="1"/>
    <col min="2" max="2" width="42" customWidth="1"/>
    <col min="3" max="3" width="12" customWidth="1"/>
    <col min="4" max="4" width="18" customWidth="1"/>
  </cols>
  <sheetData>
    $($sheetRows -join [Environment]::NewLine)
  </sheetData>
</worksheet>
"@

    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"@

    $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="18000" windowHeight="9000"/>
  </bookViews>
  <sheets>
    <sheet name="Seznam" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

    $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@

    $rootRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

    $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

    $now = [DateTime]::UtcNow.ToString("s") + "Z"
    $coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@

    $appXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Seznam</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>
"@

    if (Test-Path $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    $fileStream = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::CreateNew)
    try {
        $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "[Content_Types].xml" -Content $contentTypesXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "_rels/.rels" -Content $rootRelsXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "docProps/app.xml" -Content $appXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "docProps/core.xml" -Content $coreXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "xl/_rels/workbook.xml.rels" -Content $workbookRelsXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "xl/styles.xml" -Content $stylesXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "xl/workbook.xml" -Content $workbookXml
            Add-ZipStringEntry -ZipArchive $zipArchive -EntryName "xl/worksheets/sheet1.xml" -Content $sheetXml
        }
        finally {
            $zipArchive.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

$rawVotes = Read-FolkiVotes -Path $InputPath
$preparedVotes = foreach ($vote in $rawVotes) {
    $normalizedEmail = Normalize-Email -Email $vote.EmailRaw
    $normalizedName = Normalize-Name -Name $vote.NameRaw
    [pscustomobject]@{
        Year            = [string]$vote.Year
        Email           = $normalizedEmail
        Name            = $normalizedName
        PersonId        = Get-PersonId -NormalizedEmail $normalizedEmail -NormalizedName $normalizedName
    }
}

$people = foreach ($group in ($preparedVotes | Group-Object PersonId)) {
    $records = $group.Group
    $years = $records | Select-Object -ExpandProperty Year -Unique | Sort-Object
    $emails = $records | Select-Object -ExpandProperty Email -Unique | Sort-Object
    $names = $records | Select-Object -ExpandProperty Name -Unique | Where-Object { $_ } | Sort-Object

    $preferredNameKeys = @{
        "jiri-vondrak"                    = "jiri vondrak"
        "email:michal.simicek@volny.cz"   = "michal kosmonaut simicek"
        "email:petra.palascakova@tul.cz"  = "petra palascakova"
        "email:kaplan@folki.cz"           = "milan kaplan"
        "email:renesoucek@gmail.com"      = "rene soucek"
        "email:samson@samsonlenk.cz"      = "jaroslav samson lenk"
        "email:divmat@seznam.cz"          = "michal vanek"
        "email:mitig@centrum.cz"          = "petr peuker"
        "email:osanec@volny.cz"           = "miroslav osanec"
        "email:slavek.jan@volny.cz"       = "slavek janousek"
        "email:doug@square.cz"            = "tomas machalik"
        "email:mila.bator@gmail.com"      = "milan bator"
    }

    $bestName = $null
    if ($preferredNameKeys.ContainsKey($group.Name)) {
        $preferredKey = $preferredNameKeys[$group.Name]
        $bestName = $names | Where-Object { (Get-ComparisonKey -Value $_) -eq $preferredKey } | Select-Object -First 1
    }

    if (-not $bestName) {
        $bestName = $names |
            Sort-Object -Property @{ Expression = { Get-NameScore -Name $_ }; Descending = $true }, @{ Expression = { $_.Length }; Descending = $true } |
            Select-Object -First 1
    }

    [pscustomobject]@{
        Name       = $bestName
        Emails     = ($emails -join "; ")
        CountLabel = "$($years.Count)x"
        LastYear   = ($years | Sort-Object | Select-Object -Last 1)
        _SortKey   = $bestName
    }
}

$sortedPeople = $people | Sort-Object _SortKey | Select-Object Name, Emails, CountLabel, LastYear

New-MinimalXlsx -Rows $sortedPeople -OutputPath $OutputPath

Write-Output "Vytvořeno: $OutputPath"
Write-Output "Počet konsolidovaných hlasujících: $($sortedPeople.Count)"
