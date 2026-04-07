function Normalize-Text([string]$s) {
  if (-not $s) { return '' }
  $s = $s.ToLowerInvariant()
  $normalized = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $normalized.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($ch)
    }
  }
  $s = $sb.ToString()
  $s = $s.Replace('–', '-').Replace('—', '-').Replace('.', ' ').Replace("'", '')
  $s = $s -replace '&', ' and '
  $s = $s -replace '[^a-z0-9 -]', ' '
  $s = $s -replace '\bfc\b', ' '
  $s = $s -replace '\bafc\b', ' '
  $s = $s -replace '\bclub\b', ' '
  $s = $s -replace '\bsc\b', ' '
  $s = $s -replace '\bac\b', ' '
  $s = $s -replace '\s+', ' '
  return $s.Trim()
}

function TeamKey([string]$s) {
  $s = Normalize-Text $s
  $map = @{
    'aldershort'='aldershot'
    'leonessa'='cultural leonesa'
    'santander'='racing santander'
    'real m'='real madrid'
    'almeira'='almeria'
    'mirosol'='mirassol'
    'hildenheim'='heidenheim'
    'middlesborough'='middlesbrough'
    'carlislie'='carlisle'
    'flatwood'='fleetwood'
    'millwal'='millwall'
    'atletico m'='atletico madrid'
    'le harve'='le havre'
    'strasbourgh'='strasbourg'
    'bielfeld'='bielefeld'
    'heerenven'='heerenveen'
    'sambdoria'='sampdoria'
    'boheminas'='bohemians'
    'indepedente'='independiente del valle'
    'c budejovice'='ceske budejovice'
    'francs boais'='francs borains'
    'bellgrano'='belgrano'
    'st truidense'='st truiden'
    'st truiden'='st truiden'
    'genk krc'='genk'
    'belgrano de cordoba'='belgrano'
    'talleres de cordoba'='talleres'
    'hoffeinheim'='hoffenheim'
    'genclerbirligi'='genclerbirligi'
    'caykur rizespor'='rizespor'
    'milton keynes'='mk dons'
    'milton keynes dons'='mk dons'
    'fleetwood town'='fleetwood'
    'aldershot town'='aldershot'
    'york city'='york'
    'tamworth'='tamworth'
    'carlisle united'='carlisle'
    'brighton and hove albion'='brighton'
    'le havre ac'='le havre'
    'olympique lyonnais'='lyon'
    'as monaco'='monaco'
    'stade brestois 29'='brest'
    'amiens sc'='amiens'
    'eintracht frankfurt'='frankfurt'
    '1 fc heidenheim 1846'='heidenheim'
    'tsg hoffenheim'='hoffenheim'
    'vfl wolfsburg'='wolfsburg'
    'arminia bielefeld'='bielefeld'
    'sc paderborn 07'='paderborn'
    'real zaragoza'='zaragoza'
    'ud almeria'='almeria'
    'bohemians prague 1905'='bohemians'
    'bohemians 1905'='bohemians'
    '1 fc slovacko'='slovacko'
    'viktoria plzen'='plzen'
    'slavia prague'='slavia'
    'sparta prague'='sparta'
    'zlin'='zlin'
    'fc twente'='twente'
    'fc utrecht'='utrecht'
    'heracles'='heracles almelo'
    'schalke'='schalke 04'
    'hannover'='hannover 96'
    'dortmund'='borussia dortmund'
    'frankfurt'='eintracht frankfurt'
    'sparta'='sparta prague'
    'slavia'='slavia prague'
    'slavia b'='slavia prague b'
    'plzen'='viktoria plzen'
    'york'='york city'
    'norwich'='norwich city'
    'preston'='preston north end'
    'blackburn'='blackburn rovers'
    'lincoln'='lincoln city'
    'stockport'='stockport county'
    'swindon'='swindon town'
    'st gallen'='st gallen'
  }
  if ($map.ContainsKey($s)) {
    return $map[$s]
  }
  return $s
}

function Read-XlsxRows([string]$path) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
  try {
    $sharedEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }
    $sharedReader = New-Object System.IO.StreamReader($sharedEntry.Open())
    try {
      [xml]$sst = $sharedReader.ReadToEnd()
    } finally {
      $sharedReader.Dispose()
    }

    $shared = @()
    foreach ($si in $sst.sst.si) {
      if ($si.t) {
        $shared += [string]$si.t
      } elseif ($si.r) {
        $parts = @()
        foreach ($run in $si.r) {
          if ($run.t -is [string]) {
            $parts += $run.t
          } elseif ($run.t.'#text') {
            $parts += [string]$run.t.'#text'
          } else {
            $parts += [string]$run.t.InnerText
          }
        }
        $shared += ($parts -join '')
      } else {
        $shared += ''
      }
    }

    $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/worksheets/sheet1.xml' }
    $sheetReader = New-Object System.IO.StreamReader($sheetEntry.Open())
    try {
      [xml]$sheet = $sheetReader.ReadToEnd()
    } finally {
      $sheetReader.Dispose()
    }

    function Get-CellValue($cell, $sharedStrings) {
      if (-not $cell) { return '' }
      if ($cell.t -eq 's') {
        return $sharedStrings[[int]$cell.v]
      }
      return [string]$cell.v
    }

    $rows = @()
    foreach ($row in $sheet.worksheet.sheetData.row) {
      $r = [int]$row.r
      if ($r -le 1) { continue }
      $cells = @{}
      foreach ($cell in $row.c) {
        $cells[$cell.r] = Get-CellValue $cell $shared
      }
      if ($cells["C$r"]) {
        $rows += [pscustomobject]@{
          Row = $r
          Liga = $cells["B$r"]
          Zapas = $cells["C$r"]
        }
      }
    }
    return $rows
  } finally {
    $zip.Dispose()
  }
}

function Get-EspnEvents {
  $leagueMap = @{
    'ENG1'='eng.1';'ENG2'='eng.2';'ENG3'='eng.3';'ENG4'='eng.4';'ENG5'='eng.5'
    'ESP1'='esp.1';'ESP2'='esp.2';'GER1'='ger.1';'GER2'='ger.2';'FRA1'='fra.1';'FRA2'='fra.2'
    'ITA2'='ita.2';'HOL1'='ned.1';'POR1'='por.1';'BEL1'='bel.1';'SCO1'='sco.1';'TUR1'='tur.1'
    'SUI1'='sui.1';'CYP1'='cyp.1';'BRA1'='bra.1';'ARG1'='arg.1';'CHIL1'='chi.1';'EKV1'='ecu.1'
    'ROM1'='rou.1';'ČR1'='cze.1'
  }
  $dates = '20260314', '20260315', '20260316', '20260317', '20260318', '20260319', '20260320'
  $events = @()

  foreach ($code in $leagueMap.Keys) {
    $slug = $leagueMap[$code]
    foreach ($date in $dates) {
      try {
        $url = "https://site.api.espn.com/apis/site/v2/sports/soccer/$slug/scoreboard?dates=$date"
        $json = Invoke-WebRequest -Uri $url -UseBasicParsing | Select-Object -ExpandProperty Content | ConvertFrom-Json
        foreach ($event in @($json.events)) {
          $comp = $event.competitions[0]
          if (-not $comp) { continue }
          $homeTeam = $comp.competitors | Where-Object { $_.homeAway -eq 'home' } | Select-Object -First 1
          $awayTeam = $comp.competitors | Where-Object { $_.homeAway -eq 'away' } | Select-Object -First 1
          if (-not $homeTeam -or -not $awayTeam) { continue }
          $events += [pscustomobject]@{
            Liga = $code
            Date = $date
            Home = $homeTeam.team.displayName
            Away = $awayTeam.team.displayName
            HomeKey = TeamKey $homeTeam.team.displayName
            AwayKey = TeamKey $awayTeam.team.displayName
            Score = "$($homeTeam.score):$($awayTeam.score)"
          }
        }
      } catch {
      }
    }
  }

  return $events
}

$rows = Read-XlsxRows 'C:\Users\jan.jedlicka\Desktop\Fotbal to je hra.xlsx'
$events = Get-EspnEvents

$result = foreach ($row in ($rows | Sort-Object Liga, Zapas -Unique)) {
  $parts = ($row.Zapas -replace '–', '-' -replace '—', '-').Split('-')
  if ($parts.Count -lt 2) { continue }
  $homeKey = TeamKey($parts[0].Trim())
  $awayKey = TeamKey($parts[1].Trim())
  $candidates = @($events | Where-Object { $_.Liga -eq $row.Liga -and $_.HomeKey -eq $homeKey -and $_.AwayKey -eq $awayKey })
  if (-not $candidates) {
    $candidates = @($events | Where-Object { $_.HomeKey -eq $homeKey -and $_.AwayKey -eq $awayKey })
  }
  $first = $candidates | Select-Object -First 1
  [pscustomobject]@{
    Liga = $row.Liga
    Zapas = $row.Zapas
    Matches = $candidates.Count
    Score = $first.Score
    Date = $first.Date
    Home = $first.Home
    Away = $first.Away
  }
}

$result | Sort-Object Liga, Zapas | ConvertTo-Json -Depth 4
