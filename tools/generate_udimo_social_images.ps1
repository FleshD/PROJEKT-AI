param(
  [string]$OutputDir = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO\PPT MATERIALY WORKSHOP\socialni_podklady_verejne_zasedani'
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

function Get-Color([string]$Hex) {
  return [System.Drawing.ColorTranslator]::FromHtml($Hex)
}

function New-RoundedRectPath {
  param(
    [float]$X,
    [float]$Y,
    [float]$Width,
    [float]$Height,
    [float]$Radius
  )

  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $diameter = $Radius * 2
  $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
  $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
  $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
  $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
  $path.CloseFigure()
  return $path
}

function Add-WrappedText {
  param(
    [System.Drawing.Graphics]$Graphics,
    [string]$Text,
    [System.Drawing.Font]$Font,
    [System.Drawing.Brush]$Brush,
    [float]$X,
    [float]$Y,
    [float]$Width,
    [float]$Height,
    [System.Drawing.StringAlignment]$Alignment = [System.Drawing.StringAlignment]::Near
  )

  $format = New-Object System.Drawing.StringFormat
  $format.Alignment = $Alignment
  $format.LineAlignment = [System.Drawing.StringAlignment]::Near
  $format.Trimming = [System.Drawing.StringTrimming]::EllipsisWord
  $format.FormatFlags = [System.Drawing.StringFormatFlags]::LineLimit
  $rect = New-Object System.Drawing.RectangleF($X, $Y, $Width, $Height)
  $Graphics.DrawString($Text, $Font, $Brush, $rect, $format)
  $format.Dispose()
}

function Draw-RouteMotif {
  param(
    [System.Drawing.Graphics]$Graphics,
    [int]$Width,
    [int]$Height,
    [System.Drawing.Color]$Accent,
    [System.Drawing.Color]$Accent2
  )

  $gridPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(18, 255, 255, 255), 2)
  foreach ($x in 90, 240, 390, 540, 690, 840, 990) {
    $Graphics.DrawLine($gridPen, $x, 0, $x, $Height)
  }
  foreach ($y in 110, 260, 410, 560, 710, 860, 1010, 1160, 1310) {
    $Graphics.DrawLine($gridPen, 0, $y, $Width, $y)
  }
  $gridPen.Dispose()

  $pen1 = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(80, $Accent), 18)
  $pen1.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
  $pen1.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
  $Graphics.DrawBezier($pen1, 50, $Height - 210, 220, $Height - 370, 650, 980, $Width - 40, 470)
  $pen1.Dispose()

  $pen2 = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(70, $Accent2), 12)
  $pen2.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
  $pen2.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
  $Graphics.DrawBezier($pen2, 130, 240, 330, 150, 740, 290, 930, 140)
  $Graphics.DrawBezier($pen2, 760, $Height - 60, 640, 1020, 830, 700, $Width - 80, 510)
  $pen2.Dispose()

  foreach ($node in @(
      @{ X = 220; Y = $Height - 360; Size = 26; Fill = $Accent },
      @{ X = 548; Y = 1012; Size = 22; Fill = $Accent2 },
      @{ X = $Width - 120; Y = 540; Size = 26; Fill = $Accent },
      @{ X = 300; Y = 190; Size = 18; Fill = $Accent2 },
      @{ X = 895; Y = 175; Size = 18; Fill = $Accent2 }
    )) {
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, $node.Fill))
    $outline = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180, 255, 255, 255), 3)
    $Graphics.FillEllipse($brush, $node.X, $node.Y, $node.Size, $node.Size)
    $Graphics.DrawEllipse($outline, $node.X, $node.Y, $node.Size, $node.Size)
    $brush.Dispose()
    $outline.Dispose()
  }
}

function Get-FontFamilies {
  $collection = New-Object System.Drawing.Text.PrivateFontCollection
  foreach ($path in @(
      'C:\Windows\Fonts\framd.ttf',
      'C:\Windows\Fonts\bahnschrift.ttf',
      'C:\Windows\Fonts\segoeuib.ttf',
      'C:\Windows\Fonts\segoeui.ttf'
    )) {
    if (Test-Path -LiteralPath $path) {
      $collection.AddFontFile($path)
    }
  }

  $heading = $collection.Families | Where-Object { $_.Name -eq 'Franklin Gothic Medium' } | Select-Object -First 1
  if (-not $heading) { $heading = New-Object System.Drawing.FontFamily('Segoe UI') }
  $body = $collection.Families | Where-Object { $_.Name -eq 'Bahnschrift' } | Select-Object -First 1
  if (-not $body) { $body = New-Object System.Drawing.FontFamily('Segoe UI') }

  return [pscustomobject]@{
    Collection = $collection
    Heading = $heading
    Body = $body
  }
}

function New-Poster {
  param(
    [hashtable]$Spec,
    [string]$TargetPath,
    [System.Drawing.FontFamily]$HeadingFamily,
    [System.Drawing.FontFamily]$BodyFamily
  )

  $width = 1080
  $height = 1350
  $bitmap = New-Object System.Drawing.Bitmap($width, $height)
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  try {
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    $bg1 = Get-Color $Spec.Background1
    $bg2 = Get-Color $Spec.Background2
    $accent = Get-Color $Spec.Accent
    $accent2 = Get-Color $Spec.Accent2
    $cream = Get-Color '#F5F0E8'
    $ink = Get-Color '#13202C'

    $rect = New-Object System.Drawing.Rectangle(0, 0, $width, $height)
    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $bg1, $bg2, 45)
    $graphics.FillRectangle($bgBrush, $rect)
    $bgBrush.Dispose()

    $glow1 = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(45, $accent))
    $graphics.FillEllipse($glow1, -160, -120, 520, 520)
    $glow1.Dispose()
    $glow2 = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(38, $accent2))
    $graphics.FillEllipse($glow2, 700, 910, 460, 460)
    $glow2.Dispose()

    Draw-RouteMotif -Graphics $graphics -Width $width -Height $height -Accent $accent -Accent2 $accent2

    $labelPath = New-RoundedRectPath -X 72 -Y 68 -Width 465 -Height 56 -Radius 18
    $labelBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(36, 255, 255, 255))
    $graphics.FillPath($labelBrush, $labelPath)
    $labelBrush.Dispose()
    $labelPath.Dispose()

    $kickerFont = New-Object System.Drawing.Font($BodyFamily, 18, [System.Drawing.FontStyle]::Bold)
    $kickerBrush = New-Object System.Drawing.SolidBrush($cream)
    Add-WrappedText -Graphics $graphics -Text 'MĚSTSKÁ MOBILITA CHOMUTOVA A JIRKOVA' -Font $kickerFont -Brush $kickerBrush -X 92 -Y 84 -Width 430 -Height 28

    $badgePath = New-RoundedRectPath -X 876 -Y 66 -Width 120 -Height 122 -Radius 34
    $badgeFill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(225, $cream))
    $badgeOutline = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(95, $accent), 4)
    $graphics.FillPath($badgeFill, $badgePath)
    $graphics.DrawPath($badgeOutline, $badgePath)
    $badgeFill.Dispose()
    $badgeOutline.Dispose()
    $badgePath.Dispose()

    $numberFont = New-Object System.Drawing.Font($HeadingFamily, 40, [System.Drawing.FontStyle]::Bold)
    $numberBrush = New-Object System.Drawing.SolidBrush($ink)
    Add-WrappedText -Graphics $graphics -Text $Spec.Number -Font $numberFont -Brush $numberBrush -X 904 -Y 94 -Width 70 -Height 50 -Alignment Center
    $ofFont = New-Object System.Drawing.Font($BodyFamily, 16, [System.Drawing.FontStyle]::Bold)
    Add-WrappedText -Graphics $graphics -Text 'z 6' -Font $ofFont -Brush $numberBrush -X 904 -Y 138 -Width 70 -Height 28 -Alignment Center

    $topicFont = New-Object System.Drawing.Font($BodyFamily, 24, [System.Drawing.FontStyle]::Bold)
    $topicBrush = New-Object System.Drawing.SolidBrush($accent2)
    Add-WrappedText -Graphics $graphics -Text $Spec.Topic.ToUpperInvariant() -Font $topicFont -Brush $topicBrush -X 78 -Y 188 -Width 800 -Height 40

    $titleFont = New-Object System.Drawing.Font($HeadingFamily, 54, [System.Drawing.FontStyle]::Bold)
    $titleBrush = New-Object System.Drawing.SolidBrush($cream)
    Add-WrappedText -Graphics $graphics -Text $Spec.Title -Font $titleFont -Brush $titleBrush -X 72 -Y 242 -Width 860 -Height 250

    $calloutPath = New-RoundedRectPath -X 72 -Y 540 -Width 936 -Height 250 -Radius 34
    $calloutFill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(204, 247, 240, 232))
    $calloutOutline = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(170, $accent), 5)
    $graphics.FillPath($calloutFill, $calloutPath)
    $graphics.DrawPath($calloutOutline, $calloutPath)
    $calloutFill.Dispose()
    $calloutOutline.Dispose()
    $calloutPath.Dispose()

    $signalPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220, $accent2), 8)
    $signalPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $signalPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($signalPen, 106, 582, 106, 742)
    $graphics.DrawLine($signalPen, 106, 582, 186, 582)
    $graphics.DrawLine($signalPen, 106, 662, 222, 662)
    $graphics.DrawLine($signalPen, 106, 742, 166, 742)
    $signalPen.Dispose()

    $headlineFont = New-Object System.Drawing.Font($HeadingFamily, 34, [System.Drawing.FontStyle]::Bold)
    $headlineBrush = New-Object System.Drawing.SolidBrush($ink)
    Add-WrappedText -Graphics $graphics -Text $Spec.Headline -Font $headlineFont -Brush $headlineBrush -X 234 -Y 584 -Width 724 -Height 82

    $bodyFont = New-Object System.Drawing.Font($BodyFamily, 24, [System.Drawing.FontStyle]::Regular)
    $bodyBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(232, 27, 41, 55))
    Add-WrappedText -Graphics $graphics -Text $Spec.Body -Font $bodyFont -Brush $bodyBrush -X 234 -Y 670 -Width 714 -Height 95

    $ctaPath = New-RoundedRectPath -X 72 -Y 1008 -Width 936 -Height 196 -Radius 30
    $ctaFill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(42, 255, 255, 255))
    $graphics.FillPath($ctaFill, $ctaPath)
    $ctaFill.Dispose()
    $ctaPath.Dispose()

    $ctaFont = New-Object System.Drawing.Font($HeadingFamily, 38, [System.Drawing.FontStyle]::Bold)
    $ctaBrush = New-Object System.Drawing.SolidBrush($cream)
    Add-WrappedText -Graphics $graphics -Text $Spec.CTA -Font $ctaFont -Brush $ctaBrush -X 94 -Y 1048 -Width 700 -Height 54

    $subCtaFont = New-Object System.Drawing.Font($BodyFamily, 22, [System.Drawing.FontStyle]::Regular)
    $subCtaBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(212, 245, 240, 232))
    Add-WrappedText -Graphics $graphics -Text $Spec.SubCTA -Font $subCtaFont -Brush $subCtaBrush -X 94 -Y 1112 -Width 676 -Height 72

    $arrowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(230, $accent2), 10)
    $arrowPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $arrowPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($arrowPen, 850, 1080, 960, 1080)
    $graphics.DrawLine($arrowPen, 960, 1080, 922, 1042)
    $graphics.DrawLine($arrowPen, 960, 1080, 922, 1118)
    $arrowPen.Dispose()

    $footerFont = New-Object System.Drawing.Font($BodyFamily, 18, [System.Drawing.FontStyle]::Bold)
    $footerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(195, 245, 240, 232))
    Add-WrappedText -Graphics $graphics -Text $Spec.Footer -Font $footerFont -Brush $footerBrush -X 72 -Y 1262 -Width 920 -Height 28

    $bitmap.Save($TargetPath, [System.Drawing.Imaging.ImageFormat]::Png)

    $kickerFont.Dispose()
    $kickerBrush.Dispose()
    $numberFont.Dispose()
    $numberBrush.Dispose()
    $ofFont.Dispose()
    $topicFont.Dispose()
    $topicBrush.Dispose()
    $titleFont.Dispose()
    $titleBrush.Dispose()
    $headlineFont.Dispose()
    $headlineBrush.Dispose()
    $bodyFont.Dispose()
    $bodyBrush.Dispose()
    $ctaFont.Dispose()
    $ctaBrush.Dispose()
    $subCtaFont.Dispose()
    $subCtaBrush.Dispose()
    $footerFont.Dispose()
    $footerBrush.Dispose()
  }
  finally {
    $graphics.Dispose()
    $bitmap.Dispose()
  }
}

function New-ContactSheet {
  param(
    [string[]]$Images,
    [string]$OutputPath,
    [System.Drawing.FontFamily]$HeadingFamily,
    [System.Drawing.FontFamily]$BodyFamily
  )

  $width = 1500
  $height = 2120
  $sheet = New-Object System.Drawing.Bitmap($width, $height)
  $graphics = [System.Drawing.Graphics]::FromImage($sheet)
  try {
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $bg = New-Object System.Drawing.SolidBrush((Get-Color '#F3EFE8'))
    $graphics.FillRectangle($bg, 0, 0, $width, $height)
    $bg.Dispose()

    $titleFont = New-Object System.Drawing.Font($HeadingFamily, 34, [System.Drawing.FontStyle]::Bold)
    $titleBrush = New-Object System.Drawing.SolidBrush((Get-Color '#16212B'))
    Add-WrappedText -Graphics $graphics -Text 'Náhled | sociální podklady pro veřejné zasedání' -Font $titleFont -Brush $titleBrush -X 70 -Y 52 -Width 1100 -Height 48
    $subFont = New-Object System.Drawing.Font($BodyFamily, 18, [System.Drawing.FontStyle]::Regular)
    $subBrush = New-Object System.Drawing.SolidBrush((Get-Color '#44546A'))
    Add-WrappedText -Graphics $graphics -Text 'Městská mobilita Chomutova a Jirkova' -Font $subFont -Brush $subBrush -X 72 -Y 102 -Width 600 -Height 28

    $thumbWidth = 620
    $thumbHeight = 775
    $marginX = 80
    $marginY = 170
    $gapX = 80
    $gapY = 70

    for ($i = 0; $i -lt $Images.Count; $i++) {
      $row = [math]::Floor($i / 2)
      $col = $i % 2
      $x = $marginX + ($thumbWidth + $gapX) * $col
      $y = $marginY + ($thumbHeight + $gapY) * $row

      $shadow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(35, 0, 0, 0))
      $graphics.FillRectangle($shadow, $x + 14, $y + 16, $thumbWidth, $thumbHeight)
      $shadow.Dispose()

      $img = [System.Drawing.Image]::FromFile($Images[$i])
      try {
        $graphics.DrawImage($img, $x, $y, $thumbWidth, $thumbHeight)
      }
      finally {
        $img.Dispose()
      }
    }

    $sheet.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $titleFont.Dispose()
    $titleBrush.Dispose()
    $subFont.Dispose()
    $subBrush.Dispose()
  }
  finally {
    $graphics.Dispose()
    $sheet.Dispose()
  }
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$fonts = Get-FontFamilies

$specs = @(
  @{
    Number = '1'
    FileName = '01_verejna_doprava.png'
    Topic = 'Stůl 1 | Veřejná doprava'
    Title = "Pořadí problémů se`npo diskusi změnilo"
    Headline = 'MHD, kolony, přestupy a parciální trolejbusy'
    Body = 'Workshop potvrdil hlavní témata, ale místní debata posunula výš i konkrétní provozní a síťové souvislosti veřejné dopravy.'
    CTA = 'Přijďte na veřejné zasedání'
    SubCTA = 'Ověříme, jak tyto priority vnímají obyvatelé Chomutova a Jirkova.'
    Footer = 'Podklad pro sociální sítě | Městská mobilita Chomutova a Jirkova'
    Background1 = '#16212B'
    Background2 = '#223847'
    Accent = '#4C78A8'
    Accent2 = '#ED7D31'
  },
  @{
    Number = '2'
    FileName = '02_krizovatky_a_doprava.png'
    Topic = 'Stůl 2 | Automobilová doprava'
    Title = "Do popředí se dostaly`nkřižovatky a školní špičky"
    Headline = 'Kolony nejsou jen na hlavních tazích'
    Body = 'Diskuse ukázala, že každodenní provoz zásadně ovlivňují i strategické křižovatky a dopravní zátěž v okolí škol.'
    CTA = 'Veřejné zasedání naváže na workshop'
    SubCTA = 'Další krok je otevřená debata o tom, která místa pálí veřejnost nejvíc.'
    Footer = 'Prioritizace problémů po odborné diskusi'
    Background1 = '#17232D'
    Background2 = '#2E3A45'
    Accent = '#ED7D31'
    Accent2 = '#FFC857'
  },
  @{
    Number = '3'
    FileName = '03_parkovani.png'
    Topic = 'Stůl 3 | Doprava v klidu'
    Title = "Parkování není jen`no počtu míst"
    Headline = 'Nelegální parkování a dlouhodobé stání'
    Body = 'Po workshopu se do čela dostaly problémy, které jsou nejviditelnější přímo v ulicích a nejvíc dopadají na každodenní život.'
    CTA = 'Přidejte svůj pohled veřejnosti'
    SubCTA = 'Veřejné zasedání otevře další debatu o parkování, rezidentech i záchytných kapacitách.'
    Footer = 'Chomutov a Jirkov | doprava v klidu'
    Background1 = '#1D2831'
    Background2 = '#3A4650'
    Accent = '#FFC857'
    Accent2 = '#5B9BD5'
  },
  @{
    Number = '4'
    FileName = '04_pesi_a_cyklo.png'
    Topic = 'Stůl 4 | Aktivní mobilita'
    Title = "Bezpečné přechody`na souvislé trasy"
    Headline = 'Pěší a cyklo témata se výrazně proměnila'
    Body = 'Vysoko se dostala bezpečnost chodců i potřeba celistvé cyklosítě. Zazněla také nová témata, která původní návrh neobsáhl.'
    CTA = 'Teď dostane prostor i veřejnost'
    SubCTA = 'Přijďte říct, co je pro pěší a cyklisty v území nejdůležitější.'
    Footer = 'Bezpečnost, dostupnost, každodenní pohyb po městě'
    Background1 = '#11252B'
    Background2 = '#23424B'
    Accent = '#7FC236'
    Accent2 = '#5B9BD5'
  },
  @{
    Number = '5'
    FileName = '05_rizeni_dopravy_a_sluzby.png'
    Topic = 'Stůl 5 | Řízení dopravy a služby'
    Title = "Doprava je i o kvalitě`nživota a dostupnosti služeb"
    Headline = 'Semafory, preference MHD, senior taxi a prostředí'
    Body = 'Odborná debata ukázala, že vedle infrastruktury řeší město i služby a dopady dopravy na každodenní fungování obyvatel.'
    CTA = 'Zapojte se do veřejného zasedání'
    SubCTA = 'Další diskuse spojí technický pohled s reálnou zkušeností z území.'
    Footer = 'Životní prostředí | organizace dopravy | služby'
    Background1 = '#152633'
    Background2 = '#254154'
    Accent = '#5B9BD5'
    Accent2 = '#70AD47'
  },
  @{
    Number = '6'
    FileName = '06_verejne_zasedani.png'
    Topic = 'Shrnutí | Veřejné zasedání'
    Title = "Odborná debata změnila`npořadí priorit"
    Headline = 'Teď je řada na veřejnosti'
    Body = 'Analytické podklady byly začátek. Workshop přinesl místní specifika a nové souvislosti. Veřejné zasedání je další důležitý krok.'
    CTA = 'Pozvánka na veřejné zasedání'
    SubCTA = 'Podělte se o zkušenosti z každodenní dopravy v Chomutově a Jirkově.'
    Footer = 'Městská mobilita Chomutova a Jirkova | společné plánování'
    Background1 = '#161F29'
    Background2 = '#2D3947'
    Accent = '#ED7D31'
    Accent2 = '#7FC236'
  }
)

$generated = New-Object System.Collections.Generic.List[string]

foreach ($spec in $specs) {
  $targetPath = Join-Path $OutputDir $spec.FileName
  New-Poster -Spec $spec -TargetPath $targetPath -HeadingFamily $fonts.Heading -BodyFamily $fonts.Body
  $generated.Add($targetPath)
}

New-ContactSheet -Images $generated.ToArray() -OutputPath (Join-Path $OutputDir '00_nahled_vsech_podkladu.png') -HeadingFamily $fonts.Heading -BodyFamily $fonts.Body

$fonts.Collection.Dispose()
$generated
