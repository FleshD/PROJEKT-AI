param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string]$SourceMarkdownPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDocxPath
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ParagraphStyle {
    param(
        [string]$Text,
        [bool]$IsTitle
    )

    if ($IsTitle) {
        return 'Nzev'
    }

    if ($Text -match '^(Zahájení|[0-9]+\.)') {
        return $null
    }

    if ($Text -match '^(Výsledky hlasování:|Za Výbor|Pavla Barabášová|Renata Kocurová|…|\.{6,})') {
        return $null
    }

    return 'Odstavecseseznamem'
}

function New-ParagraphXml {
    param(
        [string]$Text,
        [string]$Style
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return '<w:p/>'
    }

    $escaped = [System.Security.SecurityElement]::Escape($Text)
    $styleXml = ''
    if ($Style) {
        $styleXml = "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>"
    }

    return "<w:p>$styleXml<w:r><w:t xml:space=`"preserve`">$escaped</w:t></w:r></w:p>"
}

$lines = Get-Content -LiteralPath $SourceMarkdownPath -Encoding UTF8
$paragraphs = New-Object System.Collections.Generic.List[string]
$isFirstParagraph = $true

foreach ($line in $lines) {
    $text = $line
    if ($isFirstParagraph -and $text.StartsWith('# ')) {
        $text = $text.Substring(2)
    }

    $style = Get-ParagraphStyle -Text $text -IsTitle:$isFirstParagraph
    $paragraphs.Add((New-ParagraphXml -Text $text -Style $style))
    $isFirstParagraph = $false
}

$tempRoot = Join-Path ([System.IO.Path]::GetDirectoryName($OutputDocxPath)) ('docx_build_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($TemplatePath, $tempRoot)

    $documentXmlPath = Join-Path $tempRoot 'word\document.xml'
    $documentXml = New-Object xml
    $documentXml.PreserveWhitespace = $true
    $documentXml.Load($documentXmlPath)

    $nsmgr = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
    $nsmgr.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

    $body = $documentXml.SelectSingleNode('//w:body', $nsmgr)
    $sectPr = $body.SelectSingleNode('./w:sectPr', $nsmgr)

    $nodesToRemove = @()
    foreach ($child in $body.ChildNodes) {
        if ($child -ne $sectPr) {
            $nodesToRemove += $child
        }
    }

    foreach ($node in $nodesToRemove) {
        [void]$body.RemoveChild($node)
    }

    foreach ($paragraphXml in $paragraphs) {
        $paragraphDoc = New-Object xml
        $paragraphDoc.LoadXml("<root xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>$paragraphXml</root>")
        $paragraphNode = $documentXml.ImportNode($paragraphDoc.DocumentElement.FirstChild, $true)
        [void]$body.InsertBefore($paragraphNode, $sectPr)
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = $utf8NoBom
    $settings.Indent = $false

    $writer = [System.Xml.XmlWriter]::Create($documentXmlPath, $settings)
    $documentXml.Save($writer)
    $writer.Close()

    if (Test-Path -LiteralPath $OutputDocxPath) {
        Remove-Item -LiteralPath $OutputDocxPath -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $OutputDocxPath)
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}
