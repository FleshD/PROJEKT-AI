param(
    [string]$OutputDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\output\pdf',
    [string]$TempDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\tmp\pdfs',
    [string]$FarmsPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\tools\island_farms.tsv',
    [string]$HotelsPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\tools\island_hotels.tsv'
)

$ErrorActionPreference = 'Stop'

function HtmlEscape {
    param([string]$Value)
    return [System.Net.WebUtility]::HtmlEncode($Value)
}

function Get-HostLabel {
    param([string]$Url)
    try {
        $uri = [System.Uri]$Url
        return $uri.Host.Replace('www.', '')
    }
    catch {
        return $Url
    }
}

function New-LinkHtml {
    param([string]$Url)
    $label = HtmlEscape (Get-HostLabel -Url $Url)
    $safeUrl = HtmlEscape $Url
    return "<a href=""$safeUrl"">$label</a>"
}

function New-ShortlistRowsHtml {
    param([object[]]$Items)

    $index = 1
    $rows = foreach ($item in $Items) {
        $num = '{0:D2}' -f $index
        $priorityClass = switch ($item.Priority) {
            'A' { 'pill pill-a' }
            'B' { 'pill pill-b' }
            default { 'pill pill-c' }
        }
        $index++
        "<tr><td class=""num"">$num</td><td><strong>$(HtmlEscape $item.Name)</strong></td><td>$(HtmlEscape $item.Region)</td><td>$(HtmlEscape $item.Roles)</td><td><span class=""$priorityClass"">$(HtmlEscape $item.Priority)</span></td><td>$(New-LinkHtml -Url $item.Url)</td></tr>"
    }

    return ($rows -join "`n")
}

function New-SignalRowsHtml {
    param([object[]]$Items)

    $rows = foreach ($item in $Items) {
        "<tr><td><strong>$(HtmlEscape $item.Employer)</strong></td><td>$(HtmlEscape $item.Role)</td><td>$(HtmlEscape $item.Note)</td><td>$(HtmlEscape $item.Published)</td><td>$(New-LinkHtml -Url $item.Url)</td></tr>"
    }

    return ($rows -join "`n")
}

New-Item -ItemType Directory -Force -Path $OutputDir, $TempDir | Out-Null

$farms = Get-Content -LiteralPath $FarmsPath -Encoding UTF8 | ConvertFrom-Csv -Delimiter "`t"
$hotels = Get-Content -LiteralPath $HotelsPath -Encoding UTF8 | ConvertFrom-Csv -Delimiter "`t"

$signals = @(
    [pscustomobject]@{
        Employer = 'Maddis'
        Role = 'housekeeping + optional farm help'
        Note = 'free accommodation on site; South Iceland countryside; English required'
        Published = '13 March 2026'
        Url = 'https://alfred.is/en/starf/housekepping-job-available-at-maddis'
    }
    [pscustomobject]@{
        Employer = 'Exeter Hotel'
        Role = 'housekeeping / room attendant'
        Note = 'good English explicitly accepted; Reykjavik; no housing'
        Published = '19 March 2026'
        Url = 'https://alfred.is/en/starf/housekeeping-room-attendant-28'
    }
    [pscustomobject]@{
        Employer = 'West Rangá Lodge'
        Role = 'housekeeping + service / kitchen + breakfast'
        Note = 'seasonal summer 2026 signal; accommodation mentioned; asks for age 22+'
        Published = '6 March 2026'
        Url = 'https://alfred.is/en/starf/season-work-at-west-ranga-lodge'
    }
    [pscustomobject]@{
        Employer = 'YLJA at Laugarás Lagoon'
        Role = 'kitchen porter / kitchen assistant'
        Note = 'staff accommodation possible; strong South Iceland hospitality signal'
        Published = '17 March 2026'
        Url = 'https://alfred.is/en/starf/uppvask-kitchen-porter-1'
    }
    [pscustomobject]@{
        Employer = 'Hótel Ísland'
        Role = 'housekeeping'
        Note = 'communicative English essential; shows recurring housekeeping demand'
        Published = '5 January 2026'
        Url = 'https://alfred.is/starf/housekeeping-56'
    }
)

$topTargets = @(
    'Friðheimar - kombinace skleníku, restaurace a turistického provozu dává šanci na kitchen/help/service roli bez vyšší kvalifikace.',
    'Efstidalur II - farma s restaurací a ubytováním; velmi praktický mix pro entry-level kandidáta.',
    'Hotel Eldhestar - funguje jako hotel i koňský ranč, takže nabízí více typů pomocných rolí než běžný hotel.',
    'Maddis - nejčerstvější veřejný signál z března 2026 a navíc se zmiňuje ubytování i zapojení do farmy.',
    'Saltvík Farm Guesthouse - typický severoislandský venkovský provoz, kde se často hodí univerzální pracovník.',
    'Heydalur - odlehlý provoz ve Westfjords obvykle znamená vyšší potřebu spolehlivých lidí na více činností.',
    'Fosshotel Glacier Lagoon - silný turistický tah na jihu Islandu a vysoká pravděpodobnost poptávky po operativních rolích.',
    'Fosshotel Mývatn - výrazná sezonnost a severní poloha dělají z housekeeping a breakfast support velmi realistickou cestu.',
    'Hotel Katla - jižní venkovský hotel, kde má entry-level kandidát obvykle lepší šanci než v centru Reykjaviku.',
    'Berjaya Mývatn Hotel - severní lokalita mimo hlavní město zvyšuje relevanci přímého oslovení pro pomocné pozice.',
    'Berjaya Hérað Hotel - regionální hotel ve východní části země, vhodný pro service a housekeeping shortlist.',
    'Fosshotel Vatnajökull - praktický cíl pro kombinaci housekeeping, dishwashing a breakfast support.'
)

$outreachSteps = @(
    'Oslovit prvních 25 subjektů do 48 hodin a další vlnu 25 subjektů během následujících 2 dnů.',
    'Primárně žádat o housekeeping, kitchen porter, breakfast shift, laundry, stable hand, greenhouse help a basic maintenance assistant.',
    'V e-mailu rovnou uvést dostupnost, fyzickou připravenost, ochotu pracovat o víkendech a zda má kandidát řidičský průkaz.',
    'U Reykjaviku filtrovat inzeráty, které výslovně požadují current residency in Iceland; pro první nástup je lepší venkovský shortlist.',
    'Po 5 až 7 dnech udělat krátký follow-up na všechny neodpovězené kontakty.'
)

$legalPoints = @(
    'Občan ČR jako EEA/EFTA národ nepotřebuje pracovní povolení pro práci na Islandu: https://island.is/en/general-info-on-work-permits',
    'Při pobytu delším než 3 měsíce je potřeba registrace: https://island.is/en/registration-of-eea-efta-foreign-nationals',
    'Při krátké placené práci do 3 měsíců může být potřeba system ID number: https://island.is/en/system-national-id-number',
    'U jednotlivých provozů je nutné hlídat podmínky typu current residency, řidičský průkaz, staff accommodation nebo věk 22+.'
)

$sourceBase = @(
    'Ísland.is - General Info on Work Permits: https://island.is/en/general-info-on-work-permits',
    'Ísland.is - Registration of EEA/EFTA citizens: https://island.is/en/registration-of-eea-efta-foreign-nationals',
    'Ísland.is - System ID number for EEA/EFTA nationals: https://island.is/en/system-national-id-number',
    'Alfreð job market examples from March 2026: Maddis, Exeter Hotel, West Rangá Lodge, Laugarás Lagoon / YLJA, Hótel Ísland',
    'Hey Iceland accommodation index and farm-stay detail pages: https://www.heyiceland.is/accommodation',
    'Center Hotels official hotel list and careers: https://www.centerhotels.com/en and https://www.centerhotels.com/en/careers',
    'Keahotels official hotel list and jobs: https://www.keahotels.is/about-us',
    'Íslandshótel / Fosshotel official contact and careers: https://www.islandshotel.is/about/contact-us/ and https://www.islandshotel.is/about-us/careers/',
    'Iceland Hotel Collection by Berjaya official hotel list and jobs: https://www.icelandhotelcollectionbyberjaya.com/en/about-us/contact-us',
    'Selected official farm / horse farm sites: fridheimar.is, efstidalur.is, laxnes.is, polarhestar.is, lythorse.is, riding.is, saltvik.is, coras.is, icelandonhorseback.com, horsesandtours.com, hestarogfjoll.com'
)

$farmsRows = New-ShortlistRowsHtml -Items $farms
$hotelsRows = New-ShortlistRowsHtml -Items $hotels
$signalsRows = New-SignalRowsHtml -Items $signals
$topTargetsHtml = ($topTargets | ForEach-Object { "<li>$(HtmlEscape $_)</li>" }) -join "`n"
$outreachStepsHtml = ($outreachSteps | ForEach-Object { "<li>$(HtmlEscape $_)</li>" }) -join "`n"
$legalPointsHtml = ($legalPoints | ForEach-Object {
        $text = HtmlEscape $_
        if ($_ -match '(https://\S+)$') {
            $url = $Matches[1]
            $escapedUrl = HtmlEscape $url
            $prefix = HtmlEscape ($_.Substring(0, $_.LastIndexOf($url)).TrimEnd())
            "<li>$prefix <a href=""$escapedUrl"">$escapedUrl</a></li>"
        }
        else {
            "<li>$text</li>"
        }
    }) -join "`n"
$sourceBaseHtml = ($sourceBase | ForEach-Object {
        $matches = [regex]::Matches($_, 'https://\S+')
        $item = HtmlEscape $_
        foreach ($match in $matches) {
            $escapedUrl = HtmlEscape $match.Value
            $item = $item.Replace($escapedUrl, "<a href=""$escapedUrl"">$escapedUrl</a>")
        }
        "<li>$item</li>"
    }) -join "`n"

$html = @"
<!DOCTYPE html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <title>Island 2026 - shortlist zaměstnavatelů</title>
  <style>
    @page { size: A4 landscape; margin: 12mm; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; color: #163047; background: #eef4f8; }
    .page { background: #ffffff; min-height: calc(210mm - 24mm); padding: 18mm 18mm 14mm 18mm; page-break-after: always; position: relative; overflow: hidden; }
    .cover { background: radial-gradient(circle at 85% 15%, rgba(41,118,187,0.20), transparent 24%), radial-gradient(circle at 10% 80%, rgba(10,70,120,0.14), transparent 28%), linear-gradient(145deg, #0f3c63 0%, #185785 52%, #1c6aa1 100%); color: #ffffff; }
    .cover::after { content: ""; position: absolute; right: -45mm; top: -35mm; width: 120mm; height: 120mm; border-radius: 50%; background: rgba(255,255,255,0.08); }
    .eyebrow { text-transform: uppercase; letter-spacing: 0.18em; font-size: 11px; font-weight: 700; opacity: 0.84; margin-bottom: 16px; }
    h1, h2, h3 { font-family: Georgia, "Times New Roman", serif; margin: 0; }
    h1 { font-size: 34px; line-height: 1.1; max-width: 260mm; margin-bottom: 16px; }
    .subtitle { max-width: 225mm; font-size: 14px; line-height: 1.6; opacity: 0.96; }
    .cover-grid { display: grid; grid-template-columns: 1.35fr 0.95fr; gap: 18px; margin-top: 26px; align-items: end; }
    .panel { background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.14); border-radius: 16px; padding: 18px 20px; }
    .panel h3 { font-size: 16px; margin-bottom: 10px; }
    .panel p, .panel li { font-size: 13px; line-height: 1.6; margin: 0; }
    .meta-list { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 18px; margin-top: 8px; font-size: 13px; line-height: 1.5; }
    .stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-top: 18px; }
    .stat { background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.14); border-radius: 16px; padding: 14px 16px; }
    .stat .num { display: block; font-size: 28px; font-weight: 800; margin-bottom: 4px; }
    .section-title { font-size: 24px; color: #0e3554; margin-bottom: 12px; }
    .lede { font-size: 14px; line-height: 1.7; color: #274760; margin-bottom: 16px; max-width: 300mm; }
    .summary-grid { display: grid; grid-template-columns: 1.1fr 0.9fr; gap: 18px; margin-bottom: 18px; }
    .card { border: 1px solid #d7e1e8; border-radius: 14px; padding: 16px 18px; background: linear-gradient(180deg, #ffffff 0%, #f9fbfc 100%); }
    .card h3 { font-size: 16px; color: #0d3a61; margin-bottom: 10px; }
    p { font-size: 13px; line-height: 1.65; margin: 0 0 10px 0; }
    ul, ol { margin: 0; padding-left: 20px; }
    li { font-size: 13px; line-height: 1.6; margin-bottom: 7px; }
    .callout { background: #f1f7fb; border-left: 5px solid #1d6ea5; padding: 12px 14px; border-radius: 8px; margin: 12px 0 0 0; font-size: 13px; line-height: 1.6; }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    thead { display: table-header-group; }
    tr { page-break-inside: avoid; }
    th, td { border: 1px solid #d7e1e8; padding: 8px 9px; vertical-align: top; font-size: 11px; line-height: 1.45; word-wrap: break-word; }
    th { background: #eaf2f7; color: #0b3658; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.03em; }
    tbody tr:nth-child(even) td { background: #fbfdfe; }
    .num { width: 9mm; font-weight: 700; color: #234761; white-space: nowrap; }
    .shortlist th:nth-child(1) { width: 10mm; } .shortlist th:nth-child(2) { width: 57mm; } .shortlist th:nth-child(3) { width: 36mm; } .shortlist th:nth-child(4) { width: 86mm; } .shortlist th:nth-child(5) { width: 19mm; } .shortlist th:nth-child(6) { width: 42mm; }
    .signals th:nth-child(1) { width: 42mm; } .signals th:nth-child(2) { width: 55mm; } .signals th:nth-child(3) { width: 110mm; } .signals th:nth-child(4) { width: 26mm; } .signals th:nth-child(5) { width: 44mm; }
    .pill { display: inline-block; min-width: 18px; text-align: center; padding: 3px 7px; border-radius: 999px; font-size: 10px; font-weight: 800; letter-spacing: 0.04em; }
    .pill-a { background: #dcefdc; color: #1f5f2e; } .pill-b { background: #e7eef5; color: #224864; } .pill-c { background: #f3e8d4; color: #7c4f00; }
    .table-note { font-size: 12px; color: #466177; margin-bottom: 10px; }
    a { color: #0d65a0; text-decoration: none; }
    .footer-note { font-size: 11px; color: #537085; margin-top: 12px; }
  </style>
</head>
<body>
  <section class="page cover">
    <div class="eyebrow">Agenturní shortlist | Island | k 29. březnu 2026</div>
    <h1>SHORTLIST 84 ZAMĚSTNAVATELŮ PRO ENTRY-LEVEL PRÁCI NA ISLANDU</h1>
    <p class="subtitle">Výstup připravený pro kandidáta z ČR: muž, 20 let, komunikativní angličtina, vysoká ochota pracovat, bez vyšší odborné kvalifikace. Fokus: housekeeping, kitchen porter, breakfast support, service support, farm stay / horse farm help, greenhouse help a základní provozní práce.</p>
    <div class="stats">
      <div class="stat"><span class="num">42</span>farem a farm stay provozů</div>
      <div class="stat"><span class="num">42</span>hotelů a hotelových skupin</div>
      <div class="stat"><span class="num">84</span>reálně oslovitelných cílů</div>
    </div>
    <div class="cover-grid">
      <div class="panel">
        <h3>Co tento report řeší</h3>
        <p>Cílem nebylo sepsat obecný seznam ubytování na Islandu, ale vybrat takové provozy, kde má mladý kandidát bez vysoké kvalifikace nejvyšší šanci získat první práci pouze na základě angličtiny, pracovitosti a ochoty vzít fyzicky i provozně náročné směny.</p>
        <div class="callout">Na Islandu je pro takový profil výrazně realističtější mířit na farm stay, horse farm, guesthouse, greenhouse a hotel operations než na čistě odborné pozice nebo městské recepční role.</div>
      </div>
      <div class="panel">
        <h3>Profil kandidáta</h3>
        <div class="meta-list">
          <div><strong>Věk:</strong> 20 let</div><div><strong>Občanství:</strong> Česká republika</div>
          <div><strong>Jazyk:</strong> angličtina</div><div><strong>Typ práce:</strong> pomocné a sezónní role</div>
          <div><strong>Silná stránka:</strong> pracovitost</div><div><strong>Slabší fit:</strong> recepce, manažerské role</div>
        </div>
      </div>
    </div>
  </section>
  <section class="page">
    <h2 class="section-title">Executive Summary</h2>
    <p class="lede">Největší šanci na získání práce má kandidát mimo centrum Reykjavíku, v provozech, kde se kombinuje ubytování, gastronomie a každodenní operativa. Proto shortlist záměrně upřednostňuje venkovské hotely, farm stay provozy, koňské farmy, greenhouse hospitality a turisticky silné oblasti jižního, severního a východního Islandu.</p>
    <div class="summary-grid">
      <div class="card">
        <h3>Hlavní závěry</h3>
        <ol>
          <li>Nejlepší fit jsou role: housekeeping, kitchen porter, breakfast support, laundry, stable hand, greenhouse help a basic maintenance assistant.</li>
          <li>Čisté "sady" nejsou pro Island hlavní cesta; vyšší šanci dávají horse farms, guesthouses a farm stay provozy s turistickým zázemím.</li>
          <li>Venkovské provozy mají často vyšší potřebu univerzálních pracovníků než městské hotely v Reykjaviku.</li>
          <li>Současné inzeráty z března 2026 potvrzují, že poptávka po housekeeping, kitchen porter a service support rolích je stále aktivní.</li>
          <li>Pro vysokou šanci na odpověď je potřeba paralelně oslovit větší počet subjektů, ideálně 30 až 50 během několika dnů.</li>
        </ol>
      </div>
      <div class="card">
        <h3>Praktické čtení priorit</h3>
        <p><span class="pill pill-a">A</span> Nejvyšší priorita: nejvýraznější fit pro angličtinu, entry-level roli a přímé oslovení.</p>
        <p><span class="pill pill-b">B</span> Silný backup: dobrý cíl pro druhou vlnu oslovení nebo jako doplnění kampaně.</p>
        <div class="callout">Priorita vychází z kombinace lokality, typu provozu, pravděpodobnosti univerzální pomocné role, veřejně dostupných kariérních stránek a aktuálních signálů z trhu práce.</div>
      </div>
    </div>
    <h2 class="section-title">Aktuální tržní signály</h2>
    <p class="table-note">Níže jsou veřejné příklady z trhu práce, které potvrzují, že i bez vyšší kvalifikace se na Islandu běžně hledají lidé do provozních a pomocných rolí. Ne všechny tyto konkrétní nabídky jsou ideální pro daný profil, ale všechny potvrzují reálnou poptávku v požadovaném segmentu.</p>
    <table class="signals"><thead><tr><th>Zaměstnavatel</th><th>Role</th><th>Co je důležité</th><th>Publikováno</th><th>Zdroj</th></tr></thead><tbody>
$signalsRows
    </tbody></table>
    <h2 class="section-title" style="margin-top:18px;">Nejsilnější cíle k oslovení jako první</h2>
    <ol>
$topTargetsHtml
    </ol>
    <div class="summary-grid" style="margin-top:18px;">
      <div class="card"><h3>Doporučená strategie oslovení</h3><ol>
$outreachStepsHtml
      </ol></div>
      <div class="card"><h3>Právní minimum pro občana ČR</h3><ul>
$legalPointsHtml
      </ul></div>
    </div>
  </section>
  <section class="page">
    <h2 class="section-title">Shortlist 42 farem a farm-stay provozů</h2>
    <p class="table-note">Do této sekce jsou záměrně zahrnuté working farms, horse farms, greenhouse hospitality, farm stay provozy a countryside guesthouses s farmářským nebo jezdeckým kontextem. Právě zde je pro entry-level kandidáta s angličtinou nejvyšší zásah.</p>
    <table class="shortlist"><thead><tr><th>#</th><th>Subjekt</th><th>Region</th><th>Nejreálnější role</th><th>Priorita</th><th>Web</th></tr></thead><tbody>
$farmsRows
    </tbody></table>
    <p class="footer-note">Poznámka: u části těchto subjektů se práce neobjevuje vždy veřejně na job boardech. U tohoto typu provozu dává smysl přímé oslovení e-mailem i v době, kdy zrovna neběží veřejný inzerát.</p>
  </section>
  <section class="page">
    <h2 class="section-title">Shortlist 42 hotelů</h2>
    <p class="table-note">Hotelový shortlist kombinuje velké islandské skupiny a venkovské provozy mimo hlavní město. Právě tyto hotely nejčastěji generují potřebu rolí jako room attendant, breakfast support, kitchen porter, stewarding nebo service support.</p>
    <table class="shortlist"><thead><tr><th>#</th><th>Subjekt</th><th>Region</th><th>Nejreálnější role</th><th>Priorita</th><th>Web</th></tr></thead><tbody>
$hotelsRows
    </tbody></table>
    <p class="footer-note">Poznámka: městské hotely v Reykjaviku jsou pořád relevantní, ale pro první nástup mají často nižší hit-rate kvůli častějším požadavkům na current residency in Iceland, vyšší konkurenci a nižší šanci na staff accommodation.</p>
  </section>
  <section class="page">
    <h2 class="section-title">Zdrojová báze a metodika</h2>
    <div class="summary-grid">
      <div class="card"><h3>Primární zdroje</h3><ul>
$sourceBaseHtml
      </ul></div>
      <div class="card">
        <h3>Metodika shortlistu</h3>
        <p>Shortlist je kurátorovaný, nikoli náhodný. Každý subjekt byl zařazen proto, že odpovídá alespoň jednomu z těchto kritérií: venkovský provoz s vyšší potřebou univerzálních pracovníků, hotelová skupina s pravidelnou náborovou stránkou, working farm / horse farm s turistickým zázemím nebo čerstvý veřejný signál poptávky po entry-level roli.</p>
        <p>Výstup neznamená garanci přijetí. Znamená však výrazně lepší startovní mapu, než kdyby kandidát oslovoval náhodné podniky bez ohledu na reálnou pravděpodobnost náboru a bez ohledu na islandská právní a provozní specifika.</p>
        <div class="callout">Doporučení pro další krok: na základě tohoto shortlistu připravit jednotný anglický outreach e-mail + CV a rozeslat první vlnu 25 až 30 přihlášek.</div>
      </div>
    </div>
  </section>
</body>
</html>
"@

$htmlPath = Join-Path $TempDir 'island_2026_shortlist_report.html'
$pdfPath = Join-Path $OutputDir 'Island_2026_shortlist_zamestnavatelu_entry_level.pdf'
$htmlPublicPath = Join-Path $OutputDir 'Island_2026_shortlist_zamestnavatelu_entry_level.html'
$screenshotPath = Join-Path $TempDir 'island_2026_shortlist_preview.png'
$browserProfileDir = Join-Path $TempDir 'chromium-profile'
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($false))
Copy-Item -LiteralPath $htmlPath -Destination $htmlPublicPath -Force
New-Item -ItemType Directory -Force -Path $browserProfileDir | Out-Null

$browserCandidates = @(
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
)
$browserPath = $browserCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $browserPath) { throw 'No Chromium-based browser found for PDF export.' }

$htmlUri = [System.Uri]::new((Resolve-Path -LiteralPath $htmlPath)).AbsoluteUri
if (Test-Path -LiteralPath $pdfPath) { Remove-Item -LiteralPath $pdfPath -Force }

& $browserPath --headless=new --disable-gpu --disable-crash-reporter --allow-file-access-from-files "--user-data-dir=$browserProfileDir" "--print-to-pdf=$pdfPath" --no-pdf-header-footer $htmlUri | Out-Null
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $pdfPath)) { throw 'PDF export failed.' }

& $browserPath --headless=new --disable-gpu --disable-crash-reporter --allow-file-access-from-files "--user-data-dir=$browserProfileDir" "--window-size=1600,2200" "--screenshot=$screenshotPath" $htmlUri | Out-Null

Write-Output "HTML: $htmlPublicPath"
Write-Output "PDF: $pdfPath"
if (Test-Path -LiteralPath $screenshotPath) { Write-Output "SCREENSHOT: $screenshotPath" }

