param(
  [string]$OutDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\analysis_josephine_cr',
  [switch]$RefreshCache
)

$ErrorActionPreference = 'Stop'

function Invoke-WebSafe {
  param(
    [string]$Url,
    [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
    [int]$MaxAttempts = 4
  )

  for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    try {
      return (Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 120 -WebSession $WebSession).Content
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

function Parse-ValueCzk {
  param([string]$ValueText)
  if ([string]::IsNullOrWhiteSpace($ValueText)) { return $null }
  if ($ValueText -match 'Neuvedeno') { return $null }
  if ($ValueText -match '([\d\s]+(?:,\d+)?)\s*CZK\s*bez DPH') {
    $numberText = ($matches[1] -replace '\s', '').Replace(',', '.')
    return [decimal]::Parse($numberText, [Globalization.CultureInfo]::InvariantCulture)
  }
  return $null
}

function Parse-DateSafe {
  param([string]$DateText)
  if ([string]::IsNullOrWhiteSpace($DateText)) { return $null }
  return [datetime]::ParseExact($DateText, 'dd.MM.yyyy', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ContractsForSystem {
  param(
    [string]$SystemName,
    [string]$Vendor,
    [string]$MarketClass,
    [string]$QueryString
  )

  $baseUrl = "https://smlouvy.gov.cz/vyhledavani?$QueryString&search=Vyhledat"
  $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
  $firstHtml = Invoke-WebSafe -Url $baseUrl -WebSession $session
  $countMatch = [regex]::Match((Normalize-SearchText $firstHtml), 'pocet naleznych zaznamu\s*(\d+)')
  if (-not $countMatch.Success) {
    throw "Unable to parse count for $SystemName"
  }

  $expectedCount = [int]$countMatch.Groups[1].Value
  $all = @()
  $currentOffset = 0
  $currentHtml = $firstHtml
  $visitedOffsets = @{}

  while ($true) {
    $visitedOffsets[$currentOffset] = $true
    $tbodyMatch = [regex]::Match($currentHtml, '(?is)<tbody class="list">(.*?)</tbody>')
    if (-not $tbodyMatch.Success) {
      throw "Unable to parse result rows for $SystemName at offset $currentOffset"
    }

    $rowMatches = [regex]::Matches($tbodyMatch.Groups[1].Value, '(?is)<tr[^>]*>(.*?)</tr>')
    foreach ($rowMatch in $rowMatches) {
      $inner = $rowMatch.Groups[1].Value
      $cells = [regex]::Matches($inner, '(?is)<td[^>]*class="([^"]+)"[^>]*>(.*?)</td>')
      if ($cells.Count -lt 7) { continue }

      $vals = @()
      foreach ($cell in $cells) {
        $vals += Strip-Html $cell.Groups[2].Value
      }

      $detailMatch = [regex]::Match($inner, 'href="([^"]*/smlouva/\d+[^"]*)"')
      $detailUrl = ''
      $detailId = ''
      if ($detailMatch.Success) {
        $detailUrl = $detailMatch.Groups[1].Value
        if ($detailUrl.StartsWith('/')) {
          $detailUrl = 'https://smlouvy.gov.cz' + $detailUrl
        }
        $idMatch = [regex]::Match($detailUrl, '/smlouva/(\d+)')
        if ($idMatch.Success) { $detailId = $idMatch.Groups[1].Value }
      }

      $dateValue = Parse-DateSafe $vals[3]
      $desc = $vals[1]
      $descNorm = Normalize-SearchText $desc

      $all += [pscustomobject]@{
        system = $SystemName
        vendor = $Vendor
        market_class = $MarketClass
        subject = $vals[0]
        description = $desc
        published_flag = $vals[2]
        contract_date = if ($dateValue) { $dateValue.ToString('yyyy-MM-dd') } else { '' }
        year = if ($dateValue) { $dateValue.Year } else { $null }
        value_text = $vals[4]
        value_czk_no_vat = Parse-ValueCzk $vals[4]
        counterparty = $vals[5]
        detail_url = $detailUrl
        detail_id = $detailId
        mentions_license = [bool]($descNorm -match 'licenc|uzivani|pronajem')
        mentions_support = [bool]($descNorm -match 'podpor|servis|udrzb|sluzeb s nimi? spojenych')
        mentions_dns = [bool]($descNorm -match '\bdns\b|dynamick')
        mentions_profile = [bool]($descNorm -match 'profil zadavatele')
        mentions_catalog = [bool]($descNorm -match 'katalog')
        mentions_auction = [bool]($descNorm -match 'aukc')
        explicit_license_count = if ($descNorm -match '(\d+)\s*(licenc|uzivatel|pristup)') { [int]$matches[1] } else { $null }
      }
    }

    $pageLinks = @()
    $linkMatches = [regex]::Matches($currentHtml, 'href="([^"]*searchResultList-offset=[^"]*)"')
    foreach ($linkMatch in $linkMatches) {
      $href = [System.Net.WebUtility]::HtmlDecode($linkMatch.Groups[1].Value)
      $offsetMatch = [regex]::Match($href, 'searchResultList-offset=(\d+)')
      if (-not $offsetMatch.Success) { continue }
      $offset = [int]$offsetMatch.Groups[1].Value
      $absoluteUrl = if ($href.StartsWith('/')) { 'https://smlouvy.gov.cz' + $href } else { $href }
      $pageLinks += [pscustomobject]@{
        offset = $offset
        url = $absoluteUrl
      }
    }

    $nextPage = $pageLinks |
      Where-Object { $_.offset -gt $currentOffset -and -not $visitedOffsets.ContainsKey($_.offset) } |
      Sort-Object offset |
      Select-Object -First 1

    if (-not $nextPage) { break }
    $currentOffset = $nextPage.offset
    $currentHtml = Invoke-WebSafe -Url $nextPage.url -WebSession $session
  }

  $deduped = @($all | Sort-Object detail_id -Unique)
  if ($deduped.Count -ne $expectedCount) {
    Write-Warning ("Count mismatch for " + $SystemName + ": expected " + $expectedCount + ", collected " + $deduped.Count)
  } else {
    Write-Host ("Collected " + $deduped.Count + " contracts for " + $SystemName)
  }

  $deduped
}

function Format-Int {
  param($Value)
  if ($null -eq $Value -or [string]$Value -eq '') { $Value = 0 }
  return ('{0:N0}' -f [double]$Value)
}

function Format-Czk {
  param($Value)
  if ($null -eq $Value -or [string]$Value -eq '') { $Value = 0 }
  return (('{0:N0}' -f [double]$Value) + ' Kč')
}

function Format-Percent {
  param($Value)
  if ($null -eq $Value -or [string]$Value -eq '') { $Value = 0 }
  return (('{0:N1}' -f [double]$Value) + ' %')
}

function Get-DisclosedValueSum {
  param([array]$Rows)
  $disclosed = @($Rows | Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' })
  if ($disclosed.Count -eq 0) { return 0 }
  return [math]::Round((($disclosed | Measure-Object -Property value_czk_no_vat -Sum).Sum), 2)
}

function Get-ContractClassification {
  param(
    [string]$Subject,
    [string]$Description
  )

  $subjectNorm = Normalize-SearchText $Subject
  $descNorm = Normalize-SearchText $Description
  $text = ($subjectNorm + ' ' + $descNorm).Trim()

  $hasAdminSignals = [bool](
    $text -match 'administrac|zadavatelsk|poradenst|konzultac|pravn|advokat|zastoupeni zadavatele|smluvni zastoupeni|provedeni ukonu|organizace vz|prikazni smlouva|prikaznik se zavazuje|administrator|minitendr|otevrene rizeni|nadlimit|podlimit|bozp|tdi|technicky dozor'
  )

  $hasToolSignals = [bool](
    $text -match 'josephine|e-zak|ezak|tender arena|tenderarena|zadavatel\.cz|zadavatel cz|firstbuysale|eveza|\bcent\b|elektronickeho nastroje|elektronicky nastroj|profil zadavatele|dynamickeho nakupniho systemu|\bdns\b|katalog|aukc|speed katalog|sw nastroj|portal fbs|proebiz tenderbox'
  )

  $hasLicenseSignals = [bool](
    $text -match 'licenc|uzivani|pronajem|poskytnuti systemu|provozovani elektronickeho nastroje|poskytnuti profilu zadavatele|sw josephine|systemu josephine|systemu proebiz|provozu aplikace e-zak'
  )

  $hasToolServiceSignals = [bool](
    $text -match 'sluzeb s nimi? spojenych|podpor|servis|udrzb|implementac|skolen|aktualizac|rozvoj|rozsiren|konzultac|technicke podpory|sprav|hosting'
  )

  $hasTrainingSignals = [bool](
    $text -match 'skoleni|obsluze|vyuzitelnosti'
  )

  if ($hasAdminSignals) {
    return [pscustomobject]@{
      type = 'full_procurement_administration'
      label = 'Konzultace / administrace / právní podpora'
      interpretation = 'Poradenské, zadavatelské, právní nebo administrativní plnění; nevykládat jako přímé SW revenue.'
      reason = 'Popis ukazuje na administraci, poradenství, zastupování zadavatele nebo jiné ne-softwarové plnění.'
    }
  }

  if ($hasToolSignals -or $hasLicenseSignals) {
    if ($hasToolServiceSignals -or $hasTrainingSignals) {
      return [pscustomobject]@{
        type = 'tool_plus_services'
        label = 'SW + rozšířené služby'
        interpretation = 'SW doplněný o servis, podporu, školení, implementaci, rozvoj nebo jiné provozní služby.'
        reason = 'Popis kombinuje elektronický nástroj s podporou, servisem, školením, rozvojem nebo obdobnou službou.'
      }
    }

    return [pscustomobject]@{
      type = 'tool_only'
      label = 'Přímé poskytování SW'
      interpretation = 'Licence, pronájem, přístup nebo samotné poskytnutí elektronického nástroje bez širší poradenské složky.'
      reason = 'Popis se soustředí na licenci, pronájem, užívání nebo samotné poskytnutí nástroje.'
    }
  }

  if ($hasToolServiceSignals -and $text -match 'software|sw|aplikac|system|elektronick|profil zadavatele') {
    return [pscustomobject]@{
      type = 'tool_plus_services'
      label = 'SW + rozšířené služby'
      interpretation = 'SW doplněný o servis, podporu, školení, implementaci, rozvoj nebo jiné provozní služby.'
      reason = 'V popisu převažují servisní nebo rozvojové práce navázané na software či elektronický nástroj.'
    }
  }

  return [pscustomobject]@{
    type = 'unclear'
    label = 'Nejasné'
    interpretation = 'Z názvu ani stručného popisu nelze spolehlivě odlišit software od širší služby.'
    reason = 'Popis je příliš obecný nebo neobsahuje jednoznačný signál o struktuře plnění.'
  }
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
  [void]$sb.AppendLine("<section class='card'><h2>$Title</h2><div class='table-wrap'><table><thead><tr><th>System</th>")
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

function Get-TypeBreakdownTable {
  param(
    [string]$Title,
    [array]$Rows
  )

  if (-not $Rows -or $Rows.Count -eq 0) { return '' }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("<section class='card'><h2>$Title</h2><div class='table-wrap'><table class='fixed compact-table'><thead><tr><th>Typ kontraktu</th><th>Smluv</th><th>Zveřejněných hodnot bez DPH</th><th>Součet bez DPH</th><th>Interpretace</th></tr></thead><tbody>")
  foreach ($row in $Rows) {
    [void]$sb.AppendLine('<tr>')
    [void]$sb.AppendLine("<td><span class='pill'>$([System.Net.WebUtility]::HtmlEncode($row.label))</span></td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Int $row.contract_count)))</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Int $row.disclosed_value_count)))</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Czk $row.total_disclosed_value_czk)))</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.interpretation))</td>")
    [void]$sb.AppendLine('</tr>')
  }
  [void]$sb.AppendLine('</tbody></table></div></section>')
  return $sb.ToString()
}

function Get-SystemBreakdownGrid {
  param([array]$Systems)

  if (-not $Systems -or $Systems.Count -eq 0) { return '' }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("<section class='card'><h2>Rozpad podle typu kontraktu - všichni konkurenti</h2><div class='system-grid'>")
  foreach ($system in $Systems) {
    [void]$sb.AppendLine("<article class='mini-card'><h3>$([System.Net.WebUtility]::HtmlEncode($system.system))</h3><div class='table-wrap'><table class='fixed compact-table mini-breakdown'><thead><tr><th>Typ</th><th>Smluv</th><th>Zveř.</th><th>Součet</th></tr></thead><tbody>")
    foreach ($row in $system.rows) {
      [void]$sb.AppendLine('<tr>')
      [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.label))</td>")
      [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Int $row.contract_count)))</td>")
      [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Int $row.disclosed_value_count)))</td>")
      [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Czk $row.total_disclosed_value_czk)))</td>")
      [void]$sb.AppendLine('</tr>')
    }
    [void]$sb.AppendLine('</tbody></table></div></article>')
  }
  [void]$sb.AppendLine('</div></section>')
  return $sb.ToString()
}

function Get-ContractsTable {
  param(
    [string]$Title,
    [array]$Contracts,
    [int]$Limit = 20
  )

  $top = @(
    $Contracts |
      Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' } |
      Sort-Object value_czk_no_vat -Descending |
      Select-Object -First $Limit
  )

  if ($top.Count -eq 0) { return '' }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("<section class='card'><h2>$Title</h2><div class='table-wrap'><table class='fixed compact-table contracts-table'><thead><tr><th>Systém</th><th>Typ plnění</th><th>Zadavatel</th><th>Popis smlouvy</th><th>Datum</th><th>Hodnota bez DPH</th></tr></thead><tbody>")
  foreach ($row in $top) {
    $detailUrl = [string]$row.detail_url
    $description = [System.Net.WebUtility]::HtmlEncode([string]$row.description)
    $detailLink = if ($detailUrl) {
      "<div class='small note'><a href='$([System.Net.WebUtility]::HtmlEncode($detailUrl))'>detail smlouvy</a></div>"
    } else {
      ''
    }

    [void]$sb.AppendLine('<tr>')
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.system))</td>")
    [void]$sb.AppendLine("<td><span class='pill'>$([System.Net.WebUtility]::HtmlEncode($row.contract_type_label))</span><div class='small note'>$([System.Net.WebUtility]::HtmlEncode($row.contract_value_interpretation))</div></td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.subject))</td>")
    [void]$sb.AppendLine("<td>$description$detailLink</td>")
    [void]$sb.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode($row.contract_date))</td>")
    [void]$sb.AppendLine("<td>" + (Format-Czk $row.value_czk_no_vat) + '</td>')
    [void]$sb.AppendLine('</tr>')
  }
  [void]$sb.AppendLine('</tbody></table></div></section>')
  return $sb.ToString()
}

$systems = @(
  [pscustomobject]@{ system = 'JOSEPHINE'; vendor = 'PROEBIZ'; market_class = 'commercial-core'; query = 'party_name=PROEBIZ&contract_descr=JOSEPHINE' },
  [pscustomobject]@{ system = 'E-ZAK'; vendor = 'QCM'; market_class = 'commercial-core'; query = 'party_name=QCM&contract_descr=E-ZAK' },
  [pscustomobject]@{ system = 'Tender arena'; vendor = 'Tender systems'; market_class = 'commercial-core'; query = 'party_name=Tender%20systems&file_text=Tender%20arena' },
  [pscustomobject]@{ system = 'ZADAVATEL.CZ'; vendor = 'OTIDEA CZ'; market_class = 'commercial-core'; query = 'party_name=OTIDEA&file_text=ZADAVATEL.CZ' },
  [pscustomobject]@{ system = 'FirstBuySale'; vendor = 'ANETE'; market_class = 'commercial-adjacent'; query = 'party_name=ANETE&file_text=FirstBuySale' },
  [pscustomobject]@{ system = 'CENT'; vendor = 'Osigeno'; market_class = 'commercial-adjacent'; query = 'party_name=Osigeno&file_text=CENT' },
  [pscustomobject]@{ system = 'Eveza'; vendor = 'ICT NWT'; market_class = 'commercial-adjacent'; query = 'party_name=ICT%20NWT&file_text=Eveza' }
)

$contractTypeDefinitions = @(
  [pscustomobject]@{
    type = 'tool_only'
    label = 'Přímé poskytování SW'
    interpretation = 'Licence, pronájem, přístup nebo samotné poskytnutí elektronického nástroje bez širší poradenské složky.'
  },
  [pscustomobject]@{
    type = 'tool_plus_services'
    label = 'SW + rozšířené služby'
    interpretation = 'SW doplněný o servis, podporu, školení, implementaci, rozvoj nebo jiné provozní služby.'
  },
  [pscustomobject]@{
    type = 'full_procurement_administration'
    label = 'Konzultace / administrace / právní podpora'
    interpretation = 'Poradenské, zadavatelské, právní nebo administrativní plnění; nevykládat jako přímé SW revenue.'
  },
  [pscustomobject]@{
    type = 'unclear'
    label = 'Nejasné'
    interpretation = 'Z názvu ani stručného popisu nelze bezpečně určit, zda převažuje nástroj nebo širší služba.'
  }
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
  $items = Get-ContractsForSystem -SystemName $system.system -Vendor $system.vendor -MarketClass $system.market_class -QueryString $system.query
  $allContracts += $items
  [System.IO.File]::WriteAllText($cachePath, (($allContracts | ConvertTo-Json -Depth 6)), $utf8Bom)
}

$allContracts = foreach ($row in $allContracts) {
  $dateValue = $null
  if ($row.contract_date) {
    try {
      $dateValue = [datetime]::ParseExact([string]$row.contract_date, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    } catch {
      $dateValue = $null
    }
  }

  $description = [string]$row.description
  $descNorm = Normalize-SearchText $description
  $classification = Get-ContractClassification -Subject ([string]$row.subject) -Description $description
  $valueCzk = Parse-ValueCzk ([string]$row.value_text)

  [pscustomobject]@{
    system = [string]$row.system
    vendor = [string]$row.vendor
    market_class = [string]$row.market_class
    subject = [string]$row.subject
    description = $description
    published_flag = [string]$row.published_flag
    contract_date = [string]$row.contract_date
    year = if ($dateValue) { $dateValue.Year } else { $row.year }
    value_text = [string]$row.value_text
    value_czk_no_vat = $valueCzk
    counterparty = [string]$row.counterparty
    detail_url = [string]$row.detail_url
    detail_id = [string]$row.detail_id
    contract_type = $classification.type
    contract_type_label = $classification.label
    contract_type_reason = $classification.reason
    contract_value_interpretation = $classification.interpretation
    mentions_license = [bool]($descNorm -match 'licenc|uzivani|pronajem')
    mentions_support = [bool]($descNorm -match 'podpor|servis|udrzb|sluzeb s nimi? spojenych')
    mentions_dns = [bool]($descNorm -match '\bdns\b|dynamick')
    mentions_profile = [bool]($descNorm -match 'profil zadavatele')
    mentions_catalog = [bool]($descNorm -match 'katalog')
    mentions_auction = [bool]($descNorm -match 'aukc')
    explicit_license_count = if ($descNorm -match '(\d+)\s*(licenc|uzivatel|pristup)') { [int]$matches[1] } else { $null }
  }
}

$allContracts = @($allContracts | Sort-Object system, contract_date, detail_id)

$summary = foreach ($group in ($allContracts | Group-Object system)) {
  $rows = @($group.Group)
  $disclosed = @($rows | Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' })
  $toolOnly = @($rows | Where-Object { $_.contract_type -eq 'tool_only' })
  $toolPlusServices = @($rows | Where-Object { $_.contract_type -eq 'tool_plus_services' })
  $fullAdmin = @($rows | Where-Object { $_.contract_type -eq 'full_procurement_administration' })
  $unclear = @($rows | Where-Object { $_.contract_type -eq 'unclear' })
  $unusableCount = $rows.Count - $disclosed.Count

  [pscustomobject]@{
    system = $group.Name
    contract_count = $rows.Count
    disclosed_value_count = $disclosed.Count
    undisclosed_value_count = $unusableCount
    undisclosed_value_share_pct = if ($rows.Count -gt 0) { [math]::Round(($unusableCount / $rows.Count) * 100, 1) } else { 0 }
    total_disclosed_value_czk = [math]::Round((($disclosed | Measure-Object -Property value_czk_no_vat -Sum).Sum), 2)
    avg_disclosed_value_czk = if ($disclosed.Count -gt 0) { [math]::Round((($disclosed | Measure-Object -Property value_czk_no_vat -Average).Average), 2) } else { 0 }
    unique_public_buyers = ($rows | Select-Object -ExpandProperty subject -Unique).Count
    tool_only_contracts = $toolOnly.Count
    tool_only_disclosed_value_czk = Get-DisclosedValueSum -Rows $toolOnly
    tool_plus_services_contracts = $toolPlusServices.Count
    tool_plus_services_disclosed_value_czk = Get-DisclosedValueSum -Rows $toolPlusServices
    full_procurement_administration_contracts = $fullAdmin.Count
    full_procurement_administration_disclosed_value_czk = Get-DisclosedValueSum -Rows $fullAdmin
    unclear_contracts = $unclear.Count
    unclear_disclosed_value_czk = Get-DisclosedValueSum -Rows $unclear
    license_like_contracts = (@($rows | Where-Object { $_.mentions_license })).Count
    support_contracts = (@($rows | Where-Object { $_.mentions_support })).Count
    dns_mentions = (@($rows | Where-Object { $_.mentions_dns })).Count
    profile_mentions = (@($rows | Where-Object { $_.mentions_profile })).Count
    catalog_mentions = (@($rows | Where-Object { $_.mentions_catalog })).Count
    auction_mentions = (@($rows | Where-Object { $_.mentions_auction })).Count
    explicit_license_units = (($rows | Where-Object { $_.explicit_license_count } | Measure-Object -Property explicit_license_count -Sum).Sum)
  }
}

$summary = @($summary | Sort-Object contract_count -Descending)
$years = $allContracts | Where-Object { $_.year } | Select-Object -ExpandProperty year -Unique | Sort-Object

$topSystemsForYear = @(
  ($summary | Where-Object { $_.system -eq 'JOSEPHINE' }),
  ($summary | Where-Object { $_.system -eq 'E-ZAK' }),
  ($summary | Where-Object { $_.system -eq 'Tender arena' }),
  ($summary | Where-Object { $_.system -eq 'ZADAVATEL.CZ' })
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
$josephineSummary = $summary | Where-Object { $_.system -eq 'JOSEPHINE' } | Select-Object -First 1

$typeSummaryAll = foreach ($typeDef in $contractTypeDefinitions) {
  $rows = @($allContracts | Where-Object { $_.contract_type -eq $typeDef.type })
  $disclosed = @($rows | Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' })
  [pscustomobject]@{
    type = $typeDef.type
    label = $typeDef.label
    interpretation = $typeDef.interpretation
    contract_count = $rows.Count
    disclosed_value_count = $disclosed.Count
    total_disclosed_value_czk = Get-DisclosedValueSum -Rows $rows
  }
}

$typeSummaryJosephine = foreach ($typeDef in $contractTypeDefinitions) {
  $rows = @($josephineRows | Where-Object { $_.contract_type -eq $typeDef.type })
  $disclosed = @($rows | Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' })
  [pscustomobject]@{
    type = $typeDef.type
    label = $typeDef.label
    interpretation = $typeDef.interpretation
    contract_count = $rows.Count
    disclosed_value_count = $disclosed.Count
    total_disclosed_value_czk = Get-DisclosedValueSum -Rows $rows
  }
}

$systemBreakdownRows = foreach ($systemSummary in $summary) {
  $rows = foreach ($typeDef in $contractTypeDefinitions) {
    $systemTypeRows = @($allContracts | Where-Object { $_.system -eq $systemSummary.system -and $_.contract_type -eq $typeDef.type })
    $disclosed = @($systemTypeRows | Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' })
    [pscustomobject]@{
      label = $typeDef.label
      contract_count = $systemTypeRows.Count
      disclosed_value_count = $disclosed.Count
      total_disclosed_value_czk = Get-DisclosedValueSum -Rows $systemTypeRows
    }
  }

  [pscustomobject]@{
    system = $systemSummary.system
    rows = $rows
  }
}

$summaryByDeliveryModel = foreach ($row in $summary) {
  [pscustomobject]@{
    system = $row.system
    contract_count = $row.contract_count
    direct_sw_contracts = $row.tool_only_contracts
    expanded_services_contracts = $row.tool_plus_services_contracts
    advisory_contracts = $row.full_procurement_administration_contracts
    other_activity_contracts = ([int]$row.tool_plus_services_contracts + [int]$row.full_procurement_administration_contracts)
    direct_sw_value_czk = $row.tool_only_disclosed_value_czk
    expanded_services_value_czk = $row.tool_plus_services_disclosed_value_czk
    advisory_value_czk = $row.full_procurement_administration_disclosed_value_czk
    other_activity_value_czk = ([decimal]$row.tool_plus_services_disclosed_value_czk + [decimal]$row.full_procurement_administration_disclosed_value_czk)
    unique_public_buyers = $row.unique_public_buyers
    undisclosed_value_count = $row.undisclosed_value_count
  }
}

$totalDirectSwContracts = (($summaryByDeliveryModel | Measure-Object -Property direct_sw_contracts -Sum).Sum)
$totalOtherActivityContracts = (($summaryByDeliveryModel | Measure-Object -Property other_activity_contracts -Sum).Sum)
$totalDirectSwValue = (($summaryByDeliveryModel | Measure-Object -Property direct_sw_value_czk -Sum).Sum)
$totalOtherActivityValue = (($summaryByDeliveryModel | Measure-Object -Property other_activity_value_czk -Sum).Sum)

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
  [void]$summaryTableRows.AppendLine("<td>$([System.Net.WebUtility]::HtmlEncode((Format-Int $row.undisclosed_value_count)))<div class='small note'>$([System.Net.WebUtility]::HtmlEncode((Format-Percent $row.undisclosed_value_share_pct)))</div></td>")
  [void]$summaryTableRows.AppendLine("<td>" + (Format-Czk $row.total_disclosed_value_czk) + '</td>')
  [void]$summaryTableRows.AppendLine("<td>$($row.unique_public_buyers)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.tool_only_contracts)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.tool_plus_services_contracts)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.full_procurement_administration_contracts)</td>")
  [void]$summaryTableRows.AppendLine("<td>$($row.unclear_contracts)</td>")
  [void]$summaryTableRows.AppendLine('</tr>')
}

$countChart = New-BarSection -Title '1. Počet smluv v registru' -Items $summary -LabelProperty 'system' -ValueProperty 'contract_count' -Color '#8b3d2e' -Formatter { param($v) Format-Int $v }
$directSwCountChart = New-BarSection -Title '2. Přímé poskytování SW - počet smluv' -Items ($summaryByDeliveryModel | Sort-Object direct_sw_contracts -Descending) -LabelProperty 'system' -ValueProperty 'direct_sw_contracts' -Color '#275d38' -Formatter { param($v) Format-Int $v }
$otherActivitiesCountChart = New-BarSection -Title '3. Další činnosti mimo přímý SW - počet smluv' -Items ($summaryByDeliveryModel | Sort-Object other_activity_contracts -Descending) -LabelProperty 'system' -ValueProperty 'other_activity_contracts' -Color '#8a5a00' -Formatter { param($v) Format-Int $v }
$directSwValueChart = New-BarSection -Title '4. Přímé poskytování SW - součet zveřejněných hodnot' -Items ($summaryByDeliveryModel | Sort-Object direct_sw_value_czk -Descending) -LabelProperty 'system' -ValueProperty 'direct_sw_value_czk' -Color '#2b5d8b' -Formatter { param($v) Format-Czk $v }
$expandedServicesValueChart = New-BarSection -Title '5. SW + rozšířené služby - součet zveřejněných hodnot' -Items ($summaryByDeliveryModel | Sort-Object expanded_services_value_czk -Descending) -LabelProperty 'system' -ValueProperty 'expanded_services_value_czk' -Color '#6e4b7c' -Formatter { param($v) Format-Czk $v }
$advisoryValueChart = New-BarSection -Title '6. Konzultace / administrace / právní podpora - součet zveřejněných hodnot' -Items ($summaryByDeliveryModel | Sort-Object advisory_value_czk -Descending) -LabelProperty 'system' -ValueProperty 'advisory_value_czk' -Color '#7b4b2a' -Formatter { param($v) Format-Czk $v }
$unusableValueChart = New-BarSection -Title '7. Smlouvy bez použitelné bez-DPH hodnoty' -Items ($summaryByDeliveryModel | Sort-Object undisclosed_value_count -Descending) -LabelProperty 'system' -ValueProperty 'undisclosed_value_count' -Color '#946f14' -Formatter { param($v) Format-Int $v }
$yearHeatmap = New-HeatmapTable -Title '8. Smlouvy podle roku (top 4 systémy)' -Rows $yearRows -Years $years
$allCompetitorBreakdownGrid = Get-SystemBreakdownGrid -Systems $systemBreakdownRows
$josephineTypeTable = Get-TypeBreakdownTable -Title 'JOSEPHINE - rozpad podle typu kontraktu' -Rows $typeSummaryJosephine
$extendedServicesTable = Get-ContractsTable -Title 'Smlouvy s rozšířenými službami' -Contracts ($allContracts | Where-Object { $_.contract_type -eq 'tool_plus_services' }) -Limit 20
$topToolOnlyContractsTable = Get-ContractsTable -Title 'Seznam 1 - pouze pronájem / licence / přímé SW' -Contracts ($allContracts | Where-Object { $_.contract_type -eq 'tool_only' }) -Limit 20
$topAdminContractsTable = Get-ContractsTable -Title 'Seznam 2 - konzultace, administrace VZ a další poradenské činnosti' -Contracts ($allContracts | Where-Object { $_.contract_type -eq 'full_procurement_administration' }) -Limit 20

$fzuExample = $allContracts |
  Where-Object {
    $_.contract_type -eq 'full_procurement_administration' -and
    $_.subject -match 'Fyzikální ústav AV ČR' -and
    $null -ne $_.value_czk_no_vat -and
    [string]$_.value_czk_no_vat -ne ''
  } |
  Sort-Object value_czk_no_vat -Descending |
  Select-Object -First 1

$top20All = @(
  $allContracts |
    Where-Object { $null -ne $_.value_czk_no_vat -and [string]$_.value_czk_no_vat -ne '' } |
    Sort-Object value_czk_no_vat -Descending |
    Select-Object -First 20
)

$top20AdminCount = (@($top20All | Where-Object { $_.contract_type -eq 'full_procurement_administration' })).Count
$totalUndisclosed = (($summary | Measure-Object -Property undisclosed_value_count -Sum).Sum)
$totalUndisclosedPct = if ($allContracts.Count -gt 0) { [math]::Round(($totalUndisclosed / $allContracts.Count) * 100, 1) } else { 0 }
$topByCount = @($summary | Select-Object -First 3)
$fzuExampleText = if ($fzuExample) {
  "Příklad z validovaných dat: $($fzuExample.system) / $($fzuExample.subject) / $([string](Format-Czk $fzuExample.value_czk_no_vat)) - $($fzuExample.description)."
} else {
  ''
}

$html = @"
<!doctype html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Analýza konkurence JOSEPHINE v České republice</title>
  <style>
    @page { size: A4 landscape; margin: 10mm; }
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; color: #1f2328; background: #f5f1ea; }
    .wrap { max-width: 1640px; margin: 0 auto; padding: 24px 20px 44px; }
    h1 { margin: 0 0 8px; font-size: 32px; line-height: 1.15; }
    h2 { margin: 0 0 12px; font-size: 20px; }
    h3 { margin: 0 0 10px; font-size: 18px; }
    .sub { margin: 0 0 18px; color: #5f6b76; }
    .card { background: #fff; border: 1px solid #d9d1c4; border-radius: 14px; padding: 16px 16px 14px; margin-bottom: 16px; box-shadow: 0 6px 18px rgba(34, 24, 11, .06); }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .meta { display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; margin-bottom: 16px; }
    .meta .box { background: #fff; border: 1px solid #d9d1c4; border-radius: 14px; padding: 12px 14px; }
    .meta .label { color: #6a7680; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
    .meta .value { font-size: 24px; font-weight: 700; margin-top: 6px; }
    .bar-row { display: grid; grid-template-columns: 220px 1fr 128px; gap: 10px; align-items: center; margin: 8px 0; }
    .bar-label { font-weight: 600; font-size: 13px; }
    .bar-wrap { height: 14px; background: #ece6dc; border-radius: 999px; overflow: hidden; }
    .bar { height: 100%; border-radius: 999px; }
    .bar-value { text-align: right; font-size: 13px; font-variant-numeric: tabular-nums; }
    table { width: 100%; border-collapse: collapse; }
    table.fixed { table-layout: fixed; }
    th, td { border-top: 1px solid #e6dfd4; padding: 8px 6px; vertical-align: top; text-align: left; }
    thead th { border-top: 0; background: #36312d; color: #fff; position: static; }
    .table-wrap { overflow: visible; }
    .compact-table th, .compact-table td { font-size: 12px; line-height: 1.35; word-break: break-word; overflow-wrap: anywhere; }
    .summary-table th:nth-child(1), .summary-table td:nth-child(1) { width: 16%; }
    .summary-table th:nth-child(2), .summary-table td:nth-child(2) { width: 7%; }
    .summary-table th:nth-child(3), .summary-table td:nth-child(3) { width: 12%; }
    .summary-table th:nth-child(4), .summary-table td:nth-child(4) { width: 11%; }
    .summary-table th:nth-child(5), .summary-table td:nth-child(5) { width: 10%; }
    .summary-table th:nth-child(6), .summary-table td:nth-child(6) { width: 10%; }
    .summary-table th:nth-child(7), .summary-table td:nth-child(7) { width: 11%; }
    .summary-table th:nth-child(8), .summary-table td:nth-child(8) { width: 13%; }
    .summary-table th:nth-child(9), .summary-table td:nth-child(9) { width: 10%; }
    .contracts-table th:nth-child(1), .contracts-table td:nth-child(1) { width: 9%; }
    .contracts-table th:nth-child(2), .contracts-table td:nth-child(2) { width: 20%; }
    .contracts-table th:nth-child(3), .contracts-table td:nth-child(3) { width: 17%; }
    .contracts-table th:nth-child(4), .contracts-table td:nth-child(4) { width: 35%; }
    .contracts-table th:nth-child(5), .contracts-table td:nth-child(5) { width: 8%; }
    .contracts-table th:nth-child(6), .contracts-table td:nth-child(6) { width: 11%; }
    .system-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .mini-card { border: 1px solid #e0d8ca; border-radius: 12px; padding: 12px; background: #fcfbf8; break-inside: avoid; }
    .mini-breakdown th:nth-child(1), .mini-breakdown td:nth-child(1) { width: 44%; }
    .mini-breakdown th:nth-child(2), .mini-breakdown td:nth-child(2) { width: 14%; }
    .mini-breakdown th:nth-child(3), .mini-breakdown td:nth-child(3) { width: 14%; }
    .mini-breakdown th:nth-child(4), .mini-breakdown td:nth-child(4) { width: 28%; }
    ul { margin: 0; padding-left: 18px; }
    .note { color: #5f6b76; }
    .small { font-size: 12px; }
    .pill { display: inline-block; padding: 4px 8px; border-radius: 999px; background: #f0ece4; border: 1px solid #d7d1c7; font-weight: 600; }
    a { color: #0d4f7a; }
    code { background: #f0ece4; padding: 1px 6px; border-radius: 6px; }
    @media print {
      body { background: #fff; font-size: 9pt; }
      .wrap { max-width: none; padding: 0; }
      .grid { grid-template-columns: 1fr 1fr; gap: 10px; }
      .meta { grid-template-columns: repeat(6, 1fr); gap: 8px; }
      .card, .meta .box, .mini-card { box-shadow: none; break-inside: avoid; page-break-inside: avoid; }
      .bar-row { grid-template-columns: 170px 1fr 104px; gap: 8px; }
      .compact-table th, .compact-table td { font-size: 10.5px; padding: 6px 5px; }
      a { color: #1f2328; text-decoration: none; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>Analýza konkurence JOSEPHINE v České republice</h1>
    <p class="sub">Primární zdroj: registr smluv. Výstup odděluje přímé poskytování SW od rozšířených služeb a poradensko-administrativních činností. Vygenerováno $(Get-Date -Format 'yyyy-MM-dd HH:mm').</p>

    <div class="meta">
      <div class="box"><div class="label">Sledovaných systémů</div><div class="value">$($systems.Count)</div></div>
      <div class="box"><div class="label">Smluv v datasetu</div><div class="value">$((Format-Int $allContracts.Count))</div></div>
      <div class="box"><div class="label">Přímé SW kontrakty</div><div class="value">$((Format-Int $totalDirectSwContracts))</div></div>
      <div class="box"><div class="label">Další činnosti</div><div class="value">$((Format-Int $totalOtherActivityContracts))</div></div>
      <div class="box"><div class="label">Bez použitelné bez-DPH ceny</div><div class="value">$((Format-Int $totalUndisclosed))</div></div>
      <div class="box"><div class="label">Součet zveřejněných hodnot</div><div class="value">$((Format-Czk (($summary | Measure-Object -Property total_disclosed_value_czk -Sum).Sum)))</div></div>
    </div>

    <section class="card">
      <h2>Metodika a omezení</h2>
      <ul>
        <li>Rámec trhu vychází z oficiálního seznamu certifikovaných elektronických nástrojů na portálu veřejných zakázek.</li>
        <li>Jádro komerčního srovnání tvoří JOSEPHINE, E-ZAK, Tender arena a ZADAVATEL.CZ. Doplňkově jsou zahrnuty menší systémy s viditelnou stopou v registru: FirstBuySale, CENT a Eveza.</li>
        <li>Pro každý systém dotaz do registru kombinuje jméno dodavatele a název systému v předmětu smlouvy nebo v prohledávatelném textu.</li>
        <li>Každá smlouva je nově klasifikována do čtyř typů: <strong>přímé poskytování SW</strong>, <strong>SW + rozšířené služby</strong>, <strong>konzultace / administrace / právní podpora</strong> a <strong>nejasné</strong>.</li>
        <li>Hodnotové grafy používají pouze částky zveřejněné explicitně jako CZK bez DPH. Záznamy se zaslepenou cenou, s cenou pouze s DPH nebo bez jasného režimu DPH jsou z těchto součtů vyloučeny.</li>
        <li>Registr smluv je vhodný pro srovnání veřejné obchodní stopy, nikoli pro přesný výpočet tržeb dodavatelů SW. V tomto datasetu nemá použitelnou bez-DPH hodnotu $((Format-Int $totalUndisclosed)) z $((Format-Int $allContracts.Count)) smluv, tedy $((Format-Percent $totalUndisclosedPct)).</li>
        <li>Vysoké hodnoty v registru nelze automaticky číst jako tržbu za software. Část TOP smluv představuje konzultace, administraci, zastupování zadavatele nebo jiné ne-softwarové činnosti.</li>
        <li>Přesné seatové počty licencí nejsou bez plošného čtení příloh PDF a OCR spolehlivě zjistitelné. Report proto stojí hlavně na počtu smluv, počtu zadavatelů, objemu zveřejněných hodnot a typu plnění.</li>
        <li>NEN není v kvantitativním srovnání zahrnut, protože funguje v odlišném veřejném provozním modelu a není přímo srovnatelný s komerčními licenčními stopami v registru.</li>
      </ul>
      <p class="note small">Zdroj pro vymezení trhu: <a href="https://portal-vz.cz/elektronicke-zadavani-verejnych-zakazek/seznam-certifikovanych-el-nastroju-dle-zakona-c-134-2016-sb/">portal-vz.cz - seznam certifikovaných elektronických nástrojů</a></p>
      $(if ($fzuExampleText) { "<p class='note small'>$([System.Net.WebUtility]::HtmlEncode($fzuExampleText))</p>" } else { '' })
    </section>

    <section class="card">
      <h2>Souhrn podle systému</h2>
      <div class="table-wrap">
        <table class="fixed compact-table summary-table">
          <thead>
            <tr>
              <th>Systém</th>
              <th>Smluv</th>
              <th>Bez použitelné bez-DPH hodnoty</th>
              <th>Součet bez DPH</th>
              <th>Unikátních zadavatelů</th>
              <th>Přímé SW</th>
              <th>SW + rozšířené služby</th>
              <th>Konzultace / admin / právní podpora</th>
              <th>Nejasné</th>
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
      $directSwCountChart
      $otherActivitiesCountChart
      $directSwValueChart
      $expandedServicesValueChart
      $advisoryValueChart
      $unusableValueChart
      $yearHeatmap
    </div>

    $allCompetitorBreakdownGrid

    $josephineTypeTable

    $extendedServicesTable

    $topToolOnlyContractsTable

    $topAdminContractsTable

    <section class="card">
      <h2>Hlavní závěry</h2>
      <ul>
        <li>JOSEPHINE má v registru smluv $((Format-Int $josephineSummary.contract_count)) dohledaných smluv a $((Format-Int $josephineSummary.unique_public_buyers)) unikátních zadavatelů. Z hlediska počtu smluv zůstává pod nejsilnější trojicí $($topByCount[0].system), $($topByCount[1].system) a $($topByCount[2].system).</li>
        <li>U JOSEPHINE převažují kontrakty typu <strong>SW + rozšířené služby</strong> ($((Format-Int (($typeSummaryJosephine | Where-Object { $_.type -eq 'tool_plus_services' }).contract_count)))) a <strong>přímé poskytování SW</strong> ($((Format-Int (($typeSummaryJosephine | Where-Object { $_.type -eq 'tool_only' }).contract_count)))). To je obchodně podstatnější než prosté čtení součtu všech částek.</li>
        <li>ZADAVATEL.CZ má v této metodice výrazně nejsilnější stopu v kategorii <strong>konzultace / administrace / právní podpora</strong>, zatímco E-ZAK a Tender arena mají výraznější mix mezi přímým SW a rozšířenými službami.</li>
        <li>V TOP 20 zveřejněných hodnotách napříč trhem je $((Format-Int $top20AdminCount)) smluv klasifikovaných jako <strong>konzultace / administrace / právní podpora</strong>. Tyto položky není vhodné prezentovat jako čisté softwarové revenue.</li>
        <li>Příklad Fyzikálního ústavu AV ČR potvrzuje obchodní připomínku: vysoké částky u ZADAVATEL.CZ odpovídají rámcovým dohodám a minitendrům na administraci zadávacích řízení, nikoli pouze licenci elektronického nástroje.</li>
        <li>Registr smluv zůstává silným zdrojem pro srovnání veřejné obchodní stopy a klientské báze, ale kvůli zaslepeným nebo nesrovnatelně zveřejněným cenám je slabší pro přesný odhad dosahovaných tržeb dodavatelů.</li>
      </ul>
    </section>

    <section class="card">
      <h2>Výstupní soubory</h2>
      <ul>
        <li><code>contracts_raw.csv</code> - surová data včetně klasifikace typu kontraktu</li>
        <li><code>summary.csv</code> - souhrn po systémech a typech plnění</li>
        <li><code>report.html</code> - report s grafy a interpretací</li>
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
