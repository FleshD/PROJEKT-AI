param(
  [string]$OutDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\analysis_josephine_sk',
  [switch]$RefreshCache
)

$ErrorActionPreference = 'Stop'

function Invoke-WebSafe {
  param(
    [string]$Url,
    [int]$MaxAttempts = 4
  )

  for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    try {
      return (Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 120).Content
    } catch {
      if ($attempt -eq $MaxAttempts) { throw }
      Start-Sleep -Seconds ([math]::Min(12, 2 * $attempt))
    }
  }
}

function Strip-Html {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  $t = $Text -replace '<.*?>', ' '
  $t = $t -replace '&nbsp;', ' '
  $t = $t -replace '\s+', ' '
  return [System.Net.WebUtility]::HtmlDecode($t.Trim())
}

function Normalize-SearchText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $normalized.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($ch)
    }
  }
  return $sb.ToString().ToLowerInvariant()
}

function Parse-ValueEur {
  param([string]$ValueText)
  if ([string]::IsNullOrWhiteSpace($ValueText)) { return $null }
  if ($ValueText -match '([\d\s]+(?:,\d+)?)\s*€') {
    $numberText = ($matches[1] -replace '\s', '').Replace(',', '.')
    return [decimal]::Parse($numberText, [Globalization.CultureInfo]::InvariantCulture)
  }
  return $null
}

function Parse-DateSafeSk {
  param([string]$DateText)
  if ([string]::IsNullOrWhiteSpace($DateText)) { return $null }

  $norm = Normalize-SearchText $DateText
  $match = [regex]::Match($norm, '(\d{1,2})\.\s*([a-z]+)\s*(\d{4})')
  if (-not $match.Success) { return $null }

  $monthMap = @{
    'januar' = 1
    'februar' = 2
    'marec' = 3
    'april' = 4
    'maj' = 5
    'jun' = 6
    'jul' = 7
    'august' = 8
    'september' = 9
    'oktober' = 10
    'november' = 11
    'december' = 12
  }

  $day = [int]$match.Groups[1].Value
  $monthName = $match.Groups[2].Value
  $year = [int]$match.Groups[3].Value
  if (-not $monthMap.ContainsKey($monthName)) { return $null }

  return [datetime]::new($year, $monthMap[$monthName], $day)
}

function Get-ContractsForSystem {
  param(
    [string]$SystemName,
    [string]$Vendor,
    [string]$MarketClass,
    [string]$VendorQuery,
    [string[]]$SubjectQueries
  )

  $all = @()
  $seen = @{}

  foreach ($subjectQuery in $SubjectQueries) {
    for ($page = 0; $page -lt 250; $page++) {
      $url = 'https://www.crz.gov.sk/2171273-sk/centralny-register-zmluv/?art_zs2=' +
        [uri]::EscapeDataString($VendorQuery) +
        '&art_predmet=' +
        [uri]::EscapeDataString($subjectQuery) +
        '&odoslat=Search'

      if ($page -gt 0) {
        $url += '&page=' + $page
      }

      $html = Invoke-WebSafe -Url $url
      $tbodyMatch = [regex]::Match($html, '(?is)<tbody>(.*?)</tbody>')
      if (-not $tbodyMatch.Success) { break }

      $rowMatches = [regex]::Matches($tbodyMatch.Groups[1].Value, '(?is)<tr>(.*?)</tr>')
      if ($rowMatches.Count -eq 0) { break }

      $newOnPage = 0
      foreach ($rowMatch in $rowMatches) {
        $inner = $rowMatch.Groups[1].Value
        $cells = [regex]::Matches($inner, '(?is)<td[^>]*class="([^"]+)"[^>]*>(.*?)</td>')
        if ($cells.Count -lt 5) { continue }

        $dateText = Strip-Html $cells[0].Groups[2].Value
        $nameCellHtml = $cells[1].Groups[2].Value
        $description = ''
        $contractNumber = ''

        $anchorMatch = [regex]::Match($nameCellHtml, '(?is)<a [^>]*>(.*?)</a>')
        if ($anchorMatch.Success) {
          $description = Strip-Html $anchorMatch.Groups[1].Value
        } else {
          $description = Strip-Html $nameCellHtml
        }

        $spanMatch = [regex]::Match($nameCellHtml, '(?is)<span>(.*?)</span>')
        if ($spanMatch.Success) {
          $contractNumber = Strip-Html $spanMatch.Groups[1].Value
        }

        $detailMatch = [regex]::Match($nameCellHtml, 'href="([^"]*/zmluva/(\d+)/[^"]*)"')
        $detailUrl = ''
        $detailId = ''
        if ($detailMatch.Success) {
          $detailUrl = $detailMatch.Groups[1].Value
          if ($detailUrl.StartsWith('/')) {
            $detailUrl = 'https://www.crz.gov.sk' + $detailUrl
          }
          $detailId = $detailMatch.Groups[2].Value
        }

        if ([string]::IsNullOrWhiteSpace($detailId)) {
          continue
        }
        if ($seen.ContainsKey($detailId)) {
          continue
        }
        $seen[$detailId] = $true
        $newOnPage++

        $dateValue = Parse-DateSafeSk $dateText
        $descNorm = Normalize-SearchText $description

        $all += [pscustomobject]@{
          system = $SystemName
          vendor = $Vendor
          market_class = $MarketClass
          buyer = Strip-Html $cells[4].Groups[2].Value
          supplier = Strip-Html $cells[3].Groups[2].Value
          description = $description
          contract_number = $contractNumber
          published_date_text = $dateText
          contract_date = if ($dateValue) { $dateValue.ToString('yyyy-MM-dd') } else { '' }
          year = if ($dateValue) { $dateValue.Year } else { $null }
          value_text = Strip-Html $cells[2].Groups[2].Value
          value_eur_vat = Parse-ValueEur (Strip-Html $cells[2].Groups[2].Value)
          detail_url = $detailUrl
          detail_id = $detailId
          mentions_license = [bool]($descNorm -match 'licenc|pristup|vyuziv|uziv')
          mentions_support = [bool]($descNorm -match 'podpor|servis|udrzb|sluz|prevadzk|suvisiac')
          mentions_archive = [bool]($descNorm -match 'archiv|uchovan')
          mentions_auction = [bool]($descNorm -match 'aukci')
          mentions_tenderbox = [bool]($descNorm -match 'tenderbox')
          mentions_workflow = [bool]($descNorm -match 'workflow')
          mentions_wendy = [bool]($descNorm -match 'wendy')
          explicit_license_count = if ($descNorm -match '(\d+)\s*(licenc|pristup)') { [int]$matches[1] } else { $null }
        }
      }

      if ($rowMatches.Count -lt 20) { break }
      if ($newOnPage -eq 0) { break }
    }
  }

  Write-Host ('Collected ' + $all.Count + ' contracts for ' + $SystemName)
  return $all | Sort-Object detail_id -Unique
}

function Format-Int([double]$Value) {
  return ('{0:N0}' -f $Value)
}

function Format-Eur([double]$Value) {
  return (('{0:N0}' -f $Value) + ' EUR')
}

function New-BarSection {
  param(
    [string]$Title,
    [array]$Items,
    [string]$LabelProperty,
    [string]$ValueProperty,
    [string]$Color = '#8b3d2e',
    [scriptblock]$Formatter = { param($v) $v }
  )

  if (-not $Items -or $Items.Count -eq 0) { return '' }
  $max = ($Items | Measure-Object -Property $ValueProperty -Maximum).Maximum
  if (-not $max) { $max = 1 }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("<section class='card'><h2>$Title</h2>")
  foreach ($item in $Items) {
    $label = [System.Net.WebUtility]::HtmlEncode([string]$item.$LabelProperty)
    $value = [double]$item.$ValueProperty
    $width = [math]::Round(($value / $max) * 100, 2)
    $display = & $Formatter $value
    [void]$sb.AppendLine("<div class='bar-row'><div class='bar-label'>$label</div><div class='bar-wrap'><div class='bar' style='width:${width}%; background:$Color'></div></div><div class='bar-value'>$display</div></div>")
  }
  [void]$sb.AppendLine('</section>')
  return $sb.ToString()
}

function New-HeatmapTable {
  param(
    [string]$Title,
    [array]$Rows,
    [array]$Years
  )

  if (-not $Rows -or $Rows.Count -eq 0) { return '' }
  $max = 0
  foreach ($row in $Rows) {
    foreach ($year in $Years) {
      $propName = "y$year"
      $v = [int]$row.$propName
      if ($v -gt $max) { $max = $v }
    }
  }
  if ($max -eq 0) { $max = 1 }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("<section class='card'><h2>$Title</h2><div class='table-wrap'><table><thead><tr><th>Systém</th>")
  foreach ($year in $Years) {
    [void]$sb.AppendLine("<th>$year</th>")
  }
  [void]$sb.AppendLine('</tr></thead><tbody>')
  foreach ($row in $Rows) {
    [void]$sb.AppendLine("<tr><td><strong>$([System.Net.WebUtility]::HtmlEncode($row.system))</strong></td>")
    foreach ($year in $Years) {
      $propName = "y$year"
      $v = [int]$row.$propName
      $alpha = [math]::Round(($v / $max) * 0.88 + 0.08, 2)
      $style = "background: rgba(139,61,46,$alpha); color: " + ($(if ($alpha -gt 0.45) { '#fff' } else { '#222' })) + "; text-align:center;"
      [void]$sb.AppendLine("<td style='$style'>$v</td>")
    }
    [void]$sb.AppendLine('</tr>')
  }
  [void]$sb.AppendLine('</tbody></table></div></section>')
  return $sb.ToString()
}

function Get-TopContractsTable {
  param([array]$Contracts)

  $top = $Contracts |
    Where-Object { $_.value_eur_vat -ne $null } |
    Sort-Object value_eur_vat -Descending |
    Select-Object -First 20

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("<section class='card'><h2>Top 20 zveřejněných hodnot smluv</h2><div class='table-wrap'><table><thead><tr><th>Systém</th><th>Zadavatel</th><th>Popis</th><th>Datum</th><th>Hodnota v EUR s DPH</th></tr></thead><tbody>")
  foreach ($row in $top) {
    [void]$sb.AppendLine('<tr>')
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.system))</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.buyer))</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.description))</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.contract_date))</td>")
    [void]$sb.AppendLine("<td>" + (Format-Eur $row.value_eur_vat) + '</td>')
    [void]$sb.AppendLine('</tr>')
  }
  [void]$sb.AppendLine('</tbody></table></div></section>')
  return $sb.ToString()
}

$systems = @(
  [pscustomobject]@{ system = 'JOSEPHINE'; vendor = 'PROEBIZ'; market_class = 'commercial-core'; vendor_query = 'PROEBIZ'; subject_queries = @('JOSEPHINE') },
  [pscustomobject]@{ system = 'TENDERnet'; vendor = 'eSYST'; market_class = 'commercial-core'; vendor_query = 'eSYST'; subject_queries = @('TENDERnet') },
  [pscustomobject]@{ system = 'ERANET'; vendor = 'innovis'; market_class = 'commercial-core'; vendor_query = 'innovis'; subject_queries = @('ERANET') },
  [pscustomobject]@{ system = 'eBIZ platform'; vendor = 'eBIZ Corp'; market_class = 'commercial-core'; vendor_query = 'eBIZ'; subject_queries = @('eZakazky', 'eAukcie', 'eProcurement') },
  [pscustomobject]@{ system = 'EVOSERVIS'; vendor = 'EVOSERVIS'; market_class = 'commercial-adjacent'; vendor_query = 'EVOSERVIS'; subject_queries = @('EVOSERVIS') },
  [pscustomobject]@{ system = 'ActiveProcurement'; vendor = 'Nuaktív'; market_class = 'commercial-adjacent'; vendor_query = 'Nuaktiv'; subject_queries = @('ActiveProcurement') }
)

if (-not (Test-Path $OutDir)) {
  New-Item -ItemType Directory -Path $OutDir | Out-Null
}

$cachePath = Join-Path $OutDir 'contracts_cache.json'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

if ((-not $RefreshCache) -and (Test-Path $cachePath)) {
  $allContracts = @((Get-Content -Raw $cachePath | ConvertFrom-Json))
} else {
  $allContracts = @()
}

$completedSystems = @($allContracts | Group-Object system | Select-Object -ExpandProperty Name)
foreach ($system in $systems) {
  if ($completedSystems -contains $system.system) {
    Write-Host ("Skipping " + $system.system + " (already in cache)")
    continue
  }

  Write-Host ("Collecting " + $system.system + " ...")
  $items = Get-ContractsForSystem -SystemName $system.system -Vendor $system.vendor -MarketClass $system.market_class -VendorQuery $system.vendor_query -SubjectQueries $system.subject_queries
  $allContracts += $items
  [System.IO.File]::WriteAllText($cachePath, (($allContracts | ConvertTo-Json -Depth 6)), $utf8Bom)
}

$allContracts = foreach ($row in $allContracts) {
  $descNorm = Normalize-SearchText $row.description
  $dateValue = $null
  if ($row.contract_date) {
    try {
      $dateValue = [datetime]::ParseExact([string]$row.contract_date, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    } catch {
      $dateValue = $null
    }
  }

  $row.value_eur_vat = Parse-ValueEur $row.value_text
  $row.year = if ($dateValue) { $dateValue.Year } else { $row.year }
  $row.mentions_license = [bool]($descNorm -match 'licenc|pristup|vyuziv|uziv')
  $row.mentions_support = [bool]($descNorm -match 'podpor|servis|udrzb|sluz|prevadzk|suvisiac')
  $row.mentions_archive = [bool]($descNorm -match 'archiv|uchovan')
  $row.mentions_auction = [bool]($descNorm -match 'aukci')
  $row.mentions_tenderbox = [bool]($descNorm -match 'tenderbox')
  $row.mentions_workflow = [bool]($descNorm -match 'workflow')
  $row.mentions_wendy = [bool]($descNorm -match 'wendy')
  $row.explicit_license_count = if ($descNorm -match '(\d+)\s*(licenc|pristup)') { [int]$matches[1] } else { $null }
  $row
}

$allContracts = $allContracts | Sort-Object system, contract_date, detail_id

$summary = foreach ($group in ($allContracts | Group-Object system)) {
  $rows = @($group.Group)
  $disclosed = @($rows | Where-Object { $_.value_eur_vat -ne $null })

  [pscustomobject]@{
    system = $group.Name
    contract_count = $rows.Count
    disclosed_value_count = $disclosed.Count
    zero_value_count = (@($disclosed | Where-Object { [decimal]$_.value_eur_vat -eq 0 })).Count
    total_disclosed_value_eur_vat = [math]::Round((($disclosed | Measure-Object -Property value_eur_vat -Sum).Sum), 2)
    avg_disclosed_value_eur_vat = if ($disclosed.Count -gt 0) { [math]::Round((($disclosed | Measure-Object -Property value_eur_vat -Average).Average), 2) } else { 0 }
    unique_public_buyers = ($rows | Select-Object -ExpandProperty buyer -Unique).Count
    license_like_contracts = (@($rows | Where-Object { $_.mentions_license })).Count
    support_contracts = (@($rows | Where-Object { $_.mentions_support })).Count
    archive_mentions = (@($rows | Where-Object { $_.mentions_archive })).Count
    auction_mentions = (@($rows | Where-Object { $_.mentions_auction })).Count
    tenderbox_mentions = (@($rows | Where-Object { $_.mentions_tenderbox })).Count
    workflow_mentions = (@($rows | Where-Object { $_.mentions_workflow })).Count
    wendy_mentions = (@($rows | Where-Object { $_.mentions_wendy })).Count
    explicit_license_units = (($rows | Where-Object { $_.explicit_license_count } | Measure-Object -Property explicit_license_count -Sum).Sum)
  }
}

$summary = $summary | Sort-Object contract_count -Descending
$years = $allContracts | Where-Object { $_.year } | Select-Object -ExpandProperty year -Unique | Sort-Object

$topSystemsForYear = @(
  ($summary | Where-Object { $_.system -eq 'JOSEPHINE' }),
  ($summary | Where-Object { $_.system -eq 'TENDERnet' }),
  ($summary | Where-Object { $_.system -eq 'ERANET' }),
  ($summary | Where-Object { $_.system -eq 'eBIZ platform' })
) | Where-Object { $_ }

$yearRows = @()
foreach ($sys in $topSystemsForYear) {
  $row = [ordered]@{ system = $sys.system }
  foreach ($year in $years) {
    $row["y$year"] = (@($allContracts | Where-Object { $_.system -eq $sys.system -and $_.year -eq $year })).Count
  }
  $yearRows += [pscustomobject]$row
}

$josephineRows = @($allContracts | Where-Object { $_.system -eq 'JOSEPHINE' })
$josephineFeatures = @(
  [pscustomobject]@{ label = 'Licence / přístup / užívání'; value = (@($josephineRows | Where-Object { $_.mentions_license })).Count },
  [pscustomobject]@{ label = 'Podpora / servis / provoz'; value = (@($josephineRows | Where-Object { $_.mentions_support })).Count },
  [pscustomobject]@{ label = 'TENDERBOX'; value = (@($josephineRows | Where-Object { $_.mentions_tenderbox })).Count },
  [pscustomobject]@{ label = 'WORKFLOW'; value = (@($josephineRows | Where-Object { $_.mentions_workflow })).Count },
  [pscustomobject]@{ label = 'WENDY'; value = (@($josephineRows | Where-Object { $_.mentions_wendy })).Count },
  [pscustomobject]@{ label = 'Aukce'; value = (@($josephineRows | Where-Object { $_.mentions_auction })).Count }
)

$contractsCsv = Join-Path $OutDir 'contracts_raw.csv'
$summaryCsv = Join-Path $OutDir 'summary.csv'
$reportHtml = Join-Path $OutDir 'report.html'

function Write-CsvUtf8 {
  param(
    [array]$Rows,
    [string]$Path
  )
  $csv = $Rows | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
  [System.IO.File]::WriteAllLines($Path, $csv, $utf8Bom)
}

Write-CsvUtf8 -Rows $allContracts -Path $contractsCsv
Write-CsvUtf8 -Rows $summary -Path $summaryCsv

$summaryTableRows = New-Object System.Text.StringBuilder
foreach ($row in $summary) {
  [void]$summaryTableRows.AppendLine('<tr>')
  [void]$summaryTableRows.AppendLine("<td><strong>$([System.Net.WebUtility]::HtmlEncode($row.system))</strong></td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.contract_count)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.disclosed_value_count)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.zero_value_count)</td>")
  [void]$summaryTableRows.AppendLine("<td>" + (Format-Eur $row.total_disclosed_value_eur_vat) + '</td>')
  [void]$summaryTableRows.AppendLine("<td>" + (Format-Eur $row.avg_disclosed_value_eur_vat) + '</td>')
  [void]$summaryTableRows.AppendLine("<td>$($row.unique_public_buyers)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.license_like_contracts)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.support_contracts)</td>")
  [void]$summaryTableRows.AppendLine('</tr>')
}

$countChart = New-BarSection -Title '1. Počet smluv v CRZ' -Items $summary -LabelProperty 'system' -ValueProperty 'contract_count' -Color '#8b3d2e' -Formatter { param($v) Format-Int $v }
$valueChart = New-BarSection -Title '2. Součet zveřejněných hodnot v EUR s DPH' -Items ($summary | Sort-Object total_disclosed_value_eur_vat -Descending) -LabelProperty 'system' -ValueProperty 'total_disclosed_value_eur_vat' -Color '#275d38' -Formatter { param($v) Format-Eur $v }
$buyerChart = New-BarSection -Title '3. Počet unikátních veřejných zadavatelů' -Items ($summary | Sort-Object unique_public_buyers -Descending) -LabelProperty 'system' -ValueProperty 'unique_public_buyers' -Color '#2b5d8b' -Formatter { param($v) Format-Int $v }
$licenseChart = New-BarSection -Title '4. Smlouvy licenčního / přístupového typu' -Items ($summary | Sort-Object license_like_contracts -Descending) -LabelProperty 'system' -ValueProperty 'license_like_contracts' -Color '#8a5a00' -Formatter { param($v) Format-Int $v }
$josephineChart = New-BarSection -Title '5. JOSEPHINE - nejčastější témata v popisech smluv' -Items $josephineFeatures -LabelProperty 'label' -ValueProperty 'value' -Color '#5b3f82' -Formatter { param($v) Format-Int $v }
$yearHeatmap = New-HeatmapTable -Title '6. Smlouvy podle roku (top 4 systémy)' -Rows $yearRows -Years $years
$topContractsTable = Get-TopContractsTable -Contracts $allContracts

$html = @"
<!doctype html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Analýza konkurence JOSEPHINE na Slovensku</title>
  <style>
    @page { size: A4; margin: 12mm; }
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; color: #1f2328; background: #f4f1eb; }
    .wrap { max-width: 1500px; margin: 0 auto; padding: 28px 22px 50px; }
    h1 { margin: 0 0 8px; font-size: 34px; }
    .sub { margin: 0 0 20px; color: #5f6b76; }
    .card { background: #fff; border: 1px solid #d7d1c7; border-radius: 14px; padding: 18px; margin-bottom: 18px; box-shadow: 0 8px 22px rgba(0,0,0,.05); }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
    .meta { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 18px; }
    .meta .box { background: #fff; border: 1px solid #d7d1c7; border-radius: 14px; padding: 14px; }
    .meta .label { color: #6a7680; font-size: 13px; text-transform: uppercase; letter-spacing: .04em; }
    .meta .value { font-size: 28px; font-weight: 700; margin-top: 4px; }
    .bar-row { display: grid; grid-template-columns: 240px 1fr 140px; gap: 12px; align-items: center; margin: 8px 0; }
    .bar-label { font-weight: 600; }
    .bar-wrap { height: 16px; background: #ece6dc; border-radius: 999px; overflow: hidden; }
    .bar { height: 100%; border-radius: 999px; }
    .bar-value { text-align: right; font-variant-numeric: tabular-nums; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border-top: 1px solid #e6dfd4; padding: 10px 8px; vertical-align: top; text-align: left; }
    thead th { border-top: 0; background: #322f2d; color: #fff; position: sticky; top: 0; }
    .table-wrap { overflow: auto; }
    ul { margin: 0; padding-left: 18px; }
    .note { color: #5f6b76; }
    .small { font-size: 13px; }
    a { color: #0d4f7a; }
    code { background: #f0ece4; padding: 1px 6px; border-radius: 6px; }
    @media print {
      body { background: #fff; font-size: 10pt; }
      .wrap { max-width: none; padding: 0; }
      .grid { grid-template-columns: 1fr; }
      .meta { grid-template-columns: repeat(2, 1fr); }
      .card, .meta .box { box-shadow: none; break-inside: avoid; page-break-inside: avoid; }
      thead th { position: static; }
      .bar-row { grid-template-columns: 180px 1fr 120px; }
      a { color: #1f2328; text-decoration: none; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>Analýza konkurence JOSEPHINE na Slovensku</h1>
    <p class="sub">Primární zdroje: slovenské CRZ a oficiální seznam zapsaných elektronických prostředků ÚVO. Vygenerováno $(Get-Date -Format 'yyyy-MM-dd HH:mm').</p>

    <div class="meta">
      <div class="box"><div class="label">Sledovaných systémů</div><div class="value">$($systems.Count)</div></div>
      <div class="box"><div class="label">Smluv v datasetu</div><div class="value">$((Format-Int $allContracts.Count))</div></div>
      <div class="box"><div class="label">Smluv JOSEPHINE</div><div class="value">$((Format-Int (($summary | Where-Object { $_.system -eq 'JOSEPHINE' }).contract_count)))</div></div>
      <div class="box"><div class="label">Součet hodnot v EUR s DPH</div><div class="value">$((Format-Eur (($summary | Measure-Object -Property total_disclosed_value_eur_vat -Sum).Sum)))</div></div>
    </div>

    <section class="card">
      <h2>Metodika a rozsah</h2>
      <ul>
        <li>Rámec trhu vychází z oficiálního seznamu zapsaných elektronických prostředků na webu Úradu pre verejné obstarávanie.</li>
        <li>Do kvantitativního srovnání jsou zahrnuty slovenské nástroje s reálně dohledatelnou smluvní stopou v CRZ: JOSEPHINE, TENDERnet, ERANET, eBIZ platform (eZakazky / eAukcie / eProcurement), EVOSERVIS a ActiveProcurement.</li>
        <li>Dotazy do CRZ kombinují název dodavatele a název systému, aby se omezily falešné shody a zůstaly smlouvy spojené s konkrétním elektronickým nástrojem.</li>
        <li>Slovenské CRZ uvádí částky v tabulkách jako EUR s DPH; hodnotové grafy proto pracují právě s touto metrikou.</li>
        <li>Stejně jako v české verzi nelze bez plošného čtení příloh PDF spolehlivě vytěžit přesné seatové počty licencí, proto report staví hlavně na počtu smluv, hodnotě, počtu zadavatelů a licenčně podobných kontraktech.</li>
        <li>Státní platformy typu IS EVO nebo EKS nejsou součástí tohoto licenčně orientovaného srovnání, protože fungují v odlišném provozním modelu než komerční nástroje.</li>
      </ul>
      <p class="note small">Zdroj tržního rámce: <a href="https://www.uvo.gov.sk/zoznam-zapisanych-elektronickych-prostriedkov">ÚVO - Zoznam zapísaných elektronických prostriedkov</a><br>Zdroj smluvních dat: <a href="https://www.crz.gov.sk/central-register-of-contracts/">CRZ Slovakia</a></p>
    </section>

    <section class="card">
      <h2>Souhrn podle systému</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Systém</th>
              <th>Smluv</th>
              <th>Zveřejněných hodnot</th>
              <th>Nulových hodnot</th>
              <th>Součet v EUR s DPH</th>
              <th>Průměr v EUR s DPH</th>
              <th>Unikátních zadavatelů</th>
              <th>Licence / přístup</th>
              <th>Podpora / servis</th>
            </tr>
          </thead>
          <tbody>
            $($summaryTableRows.ToString())
          </tbody>
        </table>
      </div>
    </section>

    <div class="grid">
      $countChart
      $valueChart
      $buyerChart
      $licenseChart
      $josephineChart
      $yearHeatmap
    </div>

    $topContractsTable

    <section class="card">
      <h2>Hlavní závěry</h2>
      <ul>
        <li>Na Slovensku má JOSEPHINE viditelnou veřejnou smluvní stopu, ale podle počtu dohledaných smluv není největším komerčním footprintem v CRZ.</li>
        <li>Silnější nebo srovnatelnou stopu v CRZ vykazují zejména TENDERnet, ERANET a eBIZ platform.</li>
        <li>U JOSEPHINE se v popisech smluv vedle licence a podpory často objevují i další moduly stejného vendor stacku, zejména TENDERBOX a WORKFLOW.</li>
        <li>CRZ je vhodné pro srovnání veřejné obchodní stopy a zveřejněných hodnot, ale nikoli pro precizní seat-based srovnání licencí bez ruční práce nad přílohami.</li>
      </ul>
    </section>

    <section class="card">
      <h2>Výstupní soubory</h2>
      <ul>
        <li><code>contracts_raw.csv</code> - surová data použitá v analýze</li>
        <li><code>summary.csv</code> - souhrn po systémech</li>
        <li><code>report.html</code> - report s grafy</li>
      </ul>
    </section>
  </div>
</body>
</html>
"@

[System.IO.File]::WriteAllText($reportHtml, $html, $utf8Bom)

[pscustomobject]@{
  out_dir = $OutDir
  systems = $systems.Count
  contracts = $allContracts.Count
  report = $reportHtml
  contracts_csv = $contractsCsv
  summary_csv = $summaryCsv
} | ConvertTo-Json -Depth 4
