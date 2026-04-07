param(
  [string]$TemplatePath = '',
  [string]$OutputPath = 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO\Pozvanka_na_verejne_zasedani.docx'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not $TemplatePath) {
  $template = Get-ChildItem 'C:\Users\jan.jedlicka\Documents\PROJEKT_AI\UDIMO' -Filter '*.docx' |
    Where-Object { $_.Name -like 'P*workshop*pozv*' } |
    Select-Object -First 1

  if (-not $template) {
    throw 'Nepodařilo se najít šablonu pozvánky v adresáři UDIMO.'
  }

  $TemplatePath = $template.FullName
}

function New-ParagraphXml {
  param(
    [string]$Text,
    [switch]$Bold
  )

  $escaped = [System.Security.SecurityElement]::Escape($Text)
  $runProps = if ($Bold) { '<w:rPr><w:b/><w:bCs/></w:rPr>' } else { '' }
  return "<w:p><w:r>$runProps<w:t xml:space=""preserve"">$escaped</w:t></w:r></w:p>"
}

$paragraphs = @(
  @{ Text = 'Pozvánka na veřejné zasedání'; Bold = $true },
  @{ Text = '' },
  @{ Text = 'Dobrý den,' },
  @{ Text = '' },
  @{ Text = 'dovolujeme si Vás pozvat na veřejné zasedání k projektu Plán udržitelné městské mobility Chomutova a Jirkova.' },
  @{ Text = '' },
  @{ Text = 'Veřejné zasedání je součástí přípravy strategického dokumentu, jehož cílem je nastavit budoucí směřování dopravy ve městech Chomutov a Jirkov. Projekt se věnuje veřejné dopravě, automobilové dopravě, parkování, cyklistické a pěší dopravě i celkové organizaci dopravy ve městě.' },
  @{ Text = '' },
  @{ Text = 'Setkání nabídne prostor pro seznámení s dosavadními výstupy projektu, představení hlavních zjištění a návrhů a také pro veřejnou diskusi nad dalšími kroky. Vaše podněty a zkušenosti z každodenního pohybu po městě jsou pro další zpracování Plánu mobility velmi důležité.' },
  @{ Text = '' },
  @{ Text = 'Termín a místo'; Bold = $true },
  @{ Text = 'Datum: [doplnit datum]' },
  @{ Text = 'Čas: [doplnit čas]' },
  @{ Text = 'Místo: [doplnit místo konání]' },
  @{ Text = '' },
  @{ Text = 'Program setkání'; Bold = $true },
  @{ Text = '1. Představení projektu Plánu udržitelné městské mobility Chomutova a Jirkova' },
  @{ Text = '2. Shrnutí dosavadních zjištění a výstupů' },
  @{ Text = '3. Představení hlavních problémových témat a navrhovaných směrů řešení' },
  @{ Text = '4. Diskuse s veřejností' },
  @{ Text = '5. Další postup projektu' },
  @{ Text = '' },
  @{ Text = 'Prosíme zájemce o potvrzení účasti na e-mailovou adresu [doplnit e-mail] nejpozději do [doplnit termín pro potvrzení].' },
  @{ Text = '' },
  @{ Text = 'Aktuální informace k projektu budou zveřejňovány na [doplnit web nebo odkaz projektu].' },
  @{ Text = '' },
  @{ Text = 'Těšíme se na setkání s Vámi.' },
  @{ Text = '' },
  @{ Text = 'S pozdravem' },
  @{ Text = '' },
  @{ Text = '[doplnit jméno a funkci]' }
)

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('udimo-docx-' + [guid]::NewGuid().ToString('N'))
$sourceZip = Join-Path ([IO.Path]::GetTempPath()) ('udimo-src-' + [guid]::NewGuid().ToString('N') + '.zip')
$outZip = Join-Path ([IO.Path]::GetTempPath()) ('udimo-out-' + [guid]::NewGuid().ToString('N') + '.zip')

New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
  Copy-Item -Path $TemplatePath -Destination $sourceZip -Force
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZip, $tempRoot)

  $bodyXml = ($paragraphs | ForEach-Object { New-ParagraphXml -Text $_.Text -Bold:([bool]$_.Bold) }) -join ''
  $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
  <w:body>
$bodyXml
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

  $documentPath = Join-Path $tempRoot 'word\document.xml'
  [System.IO.File]::WriteAllText($documentPath, $documentXml, [System.Text.UTF8Encoding]::new($false))

  if (Test-Path $outZip) {
    Remove-Item -Path $outZip -Force
  }
  [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $outZip)
  Copy-Item -Path $outZip -Destination $OutputPath -Force

  Write-Output $OutputPath
}
finally {
  if (Test-Path $sourceZip) { Remove-Item -Path $sourceZip -Force }
  if (Test-Path $outZip) { Remove-Item -Path $outZip -Force }
  if (Test-Path $tempRoot) { Remove-Item -Path $tempRoot -Recurse -Force }
}
