param(
    [string]$OutputPath = "C:\Users\jan.jedlicka\Documents\PROJEKT_AI\SVATBA\svatebni_checklist_ostravice.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertFrom-PipeBlock {
    param(
        [string]$Block,
        [string[]]$Columns,
        [int[]]$NumericColumns = @()
    )

    $records = @()
    foreach ($rawLine in ($Block -split "`r?`n")) {
        $line = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith("#")) {
            continue
        }

        $parts = $line.Split("|")
        if ($parts.Count -ne $Columns.Count) {
            throw "Line has $($parts.Count) parts, expected $($Columns.Count): $line"
        }

        $record = [ordered]@{}
        for ($i = 0; $i -lt $Columns.Count; $i++) {
            $value = $parts[$i].Trim()
            if ($NumericColumns -contains $i -and $value -ne "") {
                $record[$Columns[$i]] = [double]$value
            }
            else {
                $record[$Columns[$i]] = $value
            }
        }
        $records += [pscustomobject]$record
    }

    return $records
}

function New-TextCell {
    param(
        [AllowNull()][string]$Value,
        [int]$Style = 1
    )

    return [pscustomobject]@{
        Kind  = "Text"
        Value = if ($null -eq $Value) { "" } else { $Value }
        Style = $Style
    }
}

function New-NumberCell {
    param(
        [double]$Value,
        [int]$Style = 4
    )

    return [pscustomobject]@{
        Kind  = "Number"
        Value = $Value
        Style = $Style
    }
}

function New-FormulaCell {
    param(
        [string]$Formula,
        [int]$Style = 1
    )

    return [pscustomobject]@{
        Kind    = "Formula"
        Formula = $Formula
        Style   = $Style
    }
}

function Escape-XmlText {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) {
        return ""
    }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Get-ColumnName {
    param([int]$Index)

    $name = ""
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $name = [char](65 + $remainder) + $name
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $name
}

function ConvertTo-CellXml {
    param(
        [int]$ColumnIndex,
        [int]$RowIndex,
        [object]$Cell
    )

    $reference = "{0}{1}" -f (Get-ColumnName $ColumnIndex), $RowIndex
    $styleAttribute = ""
    if ($null -ne $Cell.Style) {
        $styleAttribute = " s=""$($Cell.Style)"""
    }

    switch ($Cell.Kind) {
        "Formula" {
            $formula = Escape-XmlText $Cell.Formula
            return "<c r=""$reference""$styleAttribute><f>$formula</f></c>"
        }
        "Number" {
            return "<c r=""$reference""$styleAttribute><v>$($Cell.Value)</v></c>"
        }
        default {
            $text = Escape-XmlText ([string]$Cell.Value)
            return "<c r=""$reference""$styleAttribute t=""inlineStr""><is><t xml:space=""preserve"">$text</t></is></c>"
        }
    }
}

function ConvertTo-RowXml {
    param(
        [int]$RowIndex,
        [object[]]$Row
    )

    if ($null -eq $Row -or $Row.Count -eq 0) {
        return "<row r=""$RowIndex"" />"
    }

    $cellXml = for ($i = 0; $i -lt $Row.Count; $i++) {
        ConvertTo-CellXml -ColumnIndex ($i + 1) -RowIndex $RowIndex -Cell $Row[$i]
    }

    return "<row r=""$RowIndex"">$($cellXml -join '')</row>"
}

function Build-ColumnsXml {
    param([double[]]$Widths)

    if ($null -eq $Widths -or $Widths.Count -eq 0) {
        return ""
    }

    $items = for ($i = 0; $i -lt $Widths.Count; $i++) {
        "<col min=""$($i + 1)"" max=""$($i + 1)"" width=""$($Widths[$i])"" customWidth=""1"" />"
    }

    return "<cols>$($items -join '')</cols>"
}

function Build-DataValidationsXml {
    param([object[]]$Definitions)

    if ($null -eq $Definitions -or $Definitions.Count -eq 0) {
        return ""
    }

    $items = foreach ($definition in $Definitions) {
        "<dataValidation type=""list"" allowBlank=""1"" showErrorMessage=""1"" showInputMessage=""1"" sqref=""$($definition.Sqref)""><formula1>&quot;$($definition.Formula)&quot;</formula1></dataValidation>"
    }

    return "<dataValidations count=""$($items.Count)"">$($items -join '')</dataValidations>"
}

function New-WorksheetXml {
    param(
        [object[]]$Rows,
        [double[]]$ColumnWidths,
        [int]$FreezeRows = 0,
        [string]$AutoFilterRef = "",
        [object[]]$DataValidations = @()
    )

    $rowCount = $Rows.Count
    $maxColumnCount = 1
    foreach ($row in $Rows) {
        if ($row.Count -gt $maxColumnCount) {
            $maxColumnCount = $row.Count
        }
    }

    $lastColumn = Get-ColumnName $maxColumnCount
    $dimension = "A1:{0}{1}" -f $lastColumn, $rowCount
    $sheetRows = for ($i = 0; $i -lt $Rows.Count; $i++) {
        ConvertTo-RowXml -RowIndex ($i + 1) -Row $Rows[$i]
    }

    $sheetViews = if ($FreezeRows -gt 0) {
        $nextRow = $FreezeRows + 1
        "<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""$FreezeRows"" topLeftCell=""A$nextRow"" activePane=""bottomLeft"" state=""frozen"" /><selection pane=""bottomLeft"" activeCell=""A$nextRow"" sqref=""A$nextRow"" /></sheetView></sheetViews>"
    }
    else {
        "<sheetViews><sheetView workbookViewId=""0"" /></sheetViews>"
    }

    $autoFilterXml = if ($AutoFilterRef) { "<autoFilter ref=""$AutoFilterRef"" />" } else { "" }
    $dataValidationsXml = Build-DataValidationsXml -Definitions $DataValidations
    $columnsXml = Build-ColumnsXml -Widths $ColumnWidths

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="$dimension" />
  $sheetViews
  <sheetFormatPr defaultRowHeight="20" />
  $columnsXml
  <sheetData>
    $($sheetRows -join "`n    ")
  </sheetData>
  $autoFilterXml
  $dataValidationsXml
  <pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
</worksheet>
"@
}

function Convert-TableToRows {
    param(
        [string[]]$Headers,
        [string[]]$Keys,
        [object[]]$Records,
        [int[]]$CurrencyColumns = @()
    )

    $rows = @()
    $rows += ,(@($Headers | ForEach-Object { New-TextCell -Value $_ -Style 2 }))

    foreach ($record in $Records) {
        $cells = @()
        for ($i = 0; $i -lt $Keys.Count; $i++) {
            $key = $Keys[$i]
            $value = $record.$key
            if ($CurrencyColumns -contains $i -and $value -ne "") {
                $cells += New-NumberCell -Value $value -Style 4
            }
            else {
                $cells += New-TextCell -Value ([string]$value) -Style 1
            }
        }
        $rows += ,$cells
    }

    return $rows
}

function Build-ChecklistSheetRows {
    param(
        [string]$NoteText,
        [object[]]$Records
    )

    $rows = @()
    $rows += ,(@((New-TextCell -Value "Poznamka" -Style 2), (New-TextCell -Value $NoteText -Style 3)))
    $rows += ,@()
    $rows += Convert-TableToRows -Headers $taskHeaders -Keys $taskColumns -Records $Records
    return $rows
}

$taskColumns = @("ID", "Faze", "Oblast", "Ukol", "Detail", "Odpovednost", "Pomocnik", "Priorita", "Stav", "MaterialPoznamka")
$taskHeaders = @("ID", "Faze", "Oblast", "Ukol", "Detail", "Odpovednost", "Pomocnik", "Priorita", "Stav", "Material / poznamka")
$scheduleColumns = @("Datum", "Od", "Do", "Blok", "Co", "Odpovednost", "Poznamka")
$scheduleHeaders = @("Datum", "Od", "Do", "Blok", "Co se deje", "Odpovednost", "Poznamka")
$budgetColumns = @("Kategorie", "Polozka", "Odhad", "Skutecnost", "Plati", "Stav", "Poznamka")
$budgetHeaders = @("Kategorie", "Polozka", "Odhad CZK", "Skutecnost CZK", "Plati", "Stav", "Poznamka")
$materialColumns = @("Skupina", "Polozka", "Mnozstvi", "KdeJe", "KdoZajisti", "KdoVrati", "Stav", "Poznamka")
$materialHeaders = @("Skupina", "Polozka", "Mnozstvi", "Kde je / od koho", "Kdo zajisti", "Kdo vrati / odveze", "Stav", "Poznamka")
$contactColumns = @("Typ", "Kontakt", "Role", "Telefon", "Komunikuje", "Stav", "Poznamka")
$contactHeaders = @("Typ", "Kontakt", "Role / firma", "Telefon", "Kdo komunikuje", "Stav", "Poznamka")
$packingColumns = @("Komu", "Polozka", "Kriticka", "Stav", "Poznamka")
$packingHeaders = @("Komu", "Polozka", "Kriticka", "Stav", "Poznamka")

$prepTasks = ConvertFrom-PipeBlock -Columns $taskColumns -Block @"
P-001|Priprava|Administrativa|Potvrdit presny cas obradu na matrice|Uzamknout cas 24.6. 14:42 a overit cas prijezdu starostky|Jan|Walter|KRITICKA|ROZPRACOVANO|Matrika a kontakt na starostku
P-002|Priprava|Administrativa|Ulozit vsechny dulezite kontakty do jednoho listu|Starostka, chata, kameraman, ridic, gastro, pujcene veci|Jan|Walter|VYSOKA|CEKA|Vyuzij list 08_Kontakty
P-003|Priprava|Administrativa|Overit pravidla chaty a louky|Ohen, hluk, parkovani, uklid, predani objektu v patek do 11:00|Jan||KRITICKA|CEKA|Majitel objektu
P-004|Priprava|Administrativa|Overit co musi byt fyzicky u obradu|Doklady, obcanky, prstynky, pripadne protokol|Jan|Walter|KRITICKA|CEKA|Sepsat minimum do jednoho pytliku
P-005|Priprava|Administrativa|Udelat mikrotym briefing se svedky|Projit role, casy, krizove body a rozhodovaci pravomoci|Jan|Michaela|KRITICKA|ROZPRACOVANO|Idealne dnes vecer
P-006|Priprava|Administrativa|Uzamknout rozdeleni odpovednosti v mikrotymu|Jan komunikace a misto, Walter krizovky, Michaela atmosfera, Veronika schvaluje citlive veci|Jan|Walter|VYSOKA|ROZPRACOVANO|Zapsat do kontaktu nebo poznamky
P-007|Priprava|Hoste|Dokoncit seznam hostu na den 1|Rodina, starsi hoste, sef z firmy, potvrdit plus minus 42 lidi|Jan|Veronika|KRITICKA|ROZPRACOVANO|Bez toho nejde jidlo ani spani
P-008|Priprava|Hoste|Dokoncit seznam hostu na den 2|Pratele a punk cast, kdo zustava a kdo prijede nove|Jan|Walter|VYSOKA|ROZPRACOVANO|Muze se menit do posledni chvile
P-009|Priprava|Hoste|Rozhodnout varianty pozvani|Den 1, den 2, oba dny, spanek nebo bez spani|Veronika|Michaela|KRITICKA|CEKA|At kazdy vi na co jede
P-010|Priprava|Hoste|Rozeslat pozvani a nastavit RSVP termin|Pridat adresu, parkovani, dress code smart casual a info o punk formatu|Jan|Veronika|KRITICKA|CEKA|Lze resit jednoduse zprava plus RSVP
P-011|Priprava|Hoste|Prubezne zapisovat RSVP|Kdo prijede kdy, kdo spi, kdo odjizdi, kdo potrebuje odvoz|Walter|Jan|KRITICKA|CEKA|Napojit na sdilenou tabulku
P-012|Priprava|Hoste|Rozdelit spani mezi chatu, stany a odjezdy domu|Chata ma 20 mist, zbytek stany nebo odjezd|Michaela|Walter|VYSOKA|CEKA|Zohlednit starsi hosty
P-013|Priprava|Hoste|Vybrat jednoho ridice a zalozni variantu odvozu|Minimalne jeden clovek ma byt dostupny pro krizovy odvoz|Walter|Jan|VYSOKA|CEKA|Napsat cislo do kontaktu
P-014|Priprava|Hoste|Poslat hostum info pred akci|Co si vzit, jak se oblect, parkovani, ze vse bude hlavne venku|Jan|Michaela|VYSOKA|CEKA|Poslat po potvrzeni ucasti
P-015|Priprava|Hoste|Pripravit visacky jako na konferenci|Jmeno plus zda host zustava pres noc nebo je jen na den|Michaela|Veronika|STREDNI|CEKA|Muze pomoct s orientaci
P-016|Priprava|Obrad|Potvrdit presne misto obradu na louce|Kde budou lidi stat nebo sedet, odkud pujde starostka a odkud vy|Michaela|Veronika|KRITICKA|HOTOVO|Louka pred chatou
P-017|Priprava|Obrad|Rozhodnout prstynky a kdo je podrzi|Vybrat pouzdro a konkretniho cloveka z mikrotymu|Jan|Walter|KRITICKA|CEKA|Nenechavat na posledni hodinu
P-018|Priprava|Obrad|Napsat prvni draft slibu Jan|Kratsi, osobni, max 1 az 2 minuty|Jan||VYSOKA|CEKA|Vytisknout pozdeji
P-019|Priprava|Obrad|Napsat prvni draft slibu Veronika|Kratsi, osobni, max 1 az 2 minuty|Veronika|Michaela|VYSOKA|CEKA|Vytisknout pozdeji
P-020|Priprava|Obrad|Vytisknout sliby na karticky a zalohu|Jedna kopie pro kazdeho a jedna nouzova kopie u svedku|Michaela|Jan|VYSOKA|CEKA|Idealne pevne karticky
P-021|Priprava|Obrad|Vybrat hudbu pro nastup a konec obradu|Minimalne nastup, podpis nebo prstynky, odchod|Jan|Veronika|VYSOKA|CEKA|Napojit na playlist
P-022|Priprava|Obrad|Sepsat jednoduchy scenar obradu|Nastup, uvod starostky, sliby, prstynky, polibek, gratulace|Michaela|Walter|KRITICKA|CEKA|Jedna stranka pro vsechny
P-023|Priprava|Obrad|Pripravit tahak pro svedky k obradu|Kdo kde stoji, kdo pusti hudbu, kdo hlida cas|Walter|Michaela|VYSOKA|CEKA|At neni chaos tesne pred 14:42
P-024|Priprava|Misto|Nakreslit mapu louky a chaty|Zona obrad, jidlo, piti, ohen, tanec, odpad, parkovani, stany|Jan|Walter|KRITICKA|CEKA|Staci jednoduchy nacrt
P-025|Priprava|Misto|Spocitat sezeni pro den 1 a den 2|Kde budou pivni sety, kde lavicky, kde volne stani|Walter|Michaela|VYSOKA|CEKA|Den 1 potrebuje vic klidu
P-026|Priprava|Misto|Zajistit party stany|Zjistit kolik kusu a kdo je priveze na chatu|Jan|parta na stavbu|KRITICKA|CEKA|Musi byt pripraveno nejpozdeji ve stredu dopoledne
P-027|Priprava|Misto|Zajistit pivni sety, lavicky a zidli navic|Hlavne pro den 1 a starsi hosty|Walter|ridic|KRITICKA|CEKA|Pripsat do listu 07_Material
P-028|Priprava|Misto|Rozhodnout destovou variantu|Co se presune dovnitr a co se jen prikryje|Jan|Michaela|KRITICKA|CEKA|Mit hotove driv nez prijede rodina
P-029|Priprava|Misto|Naplanovat osvetleni louky a cesty k chate|Svetla vecer, orientace v noci a cesta na zachod|Michaela|Walter|VYSOKA|CEKA|Lampicky, reflektory, baterky
P-030|Priprava|Misto|Vytisknout orientacni cedule|Parkovani, zachod, darky, ohen, odpad, voda|Michaela|Veronika|STREDNI|CEKA|Jednoduche a citelne
P-031|Priprava|Misto|Rozhodnout fallback pro hlasitost po setmeni|Kdy pripadne ztlumit, jak prepnout na klidnejsi rezim|Walter|Jan|STREDNI|CEKA|At je plan i kdyz sousedi budou citlivi
P-032|Priprava|Hudba a elektro|Overit zdroje elektriny a jistic|Kde je nejblizsi zasuvka, co utahne hudbu a svetla|Jan|Walter|KRITICKA|CEKA|Na louce muze byt dlouha trasa
P-033|Priprava|Hudba a elektro|Sesbirat prodluzovacky vcetne dlouhe trasy|Pocitat i se 100 m od chaty na louku|Walter|parta na stavbu|KRITICKA|CEKA|Oznacit komu patri
P-034|Priprava|Hudba a elektro|Sesbirat rozdvojky, listy, pasky a stahovaci pasky|At kabely nelezi nebezpecne v trave|Walter|Jan|VYSOKA|CEKA|Hodi se i gaffa
P-035|Priprava|Hudba a elektro|Zajistit hlavni bednu nebo aparaturu|Velka JBL nebo podobne reseni s dostatecnym vykonem|Jan|Walter|KRITICKA|CEKA|Bez hudby padne atmosfera
P-036|Priprava|Hudba a elektro|Zajistit zalohu pro hudbu|Druha bedna, druhe zarizeni nebo kabelove pripojeni|Walter|Jan|VYSOKA|CEKA|Kdyby spadl Bluetooth
P-037|Priprava|Hudba a elektro|Pripravit playlisty po hodinach pro 24.6.|Ceremony, po obrade, jidlo, vecer, pozdni noc|Jan|Veronika|KRITICKA|CEKA|Vytisknout i jako plakat
P-038|Priprava|Hudba a elektro|Pripravit playlisty po hodinach pro 25.6.|Volnejsi punk format s ohnem a chill bloky|Jan|Walter|VYSOKA|CEKA|Muze byt hravejsi nez den 1
P-039|Priprava|Hudba a elektro|Stahnout playlisty offline a otestovat Spotify|At hudba jede i bez signalu|Jan||KRITICKA|CEKA|Telefon plus zalozni telefon
P-040|Priprava|Hudba a elektro|Vytisknout hudebni plakat po hodinach|Na styl festivaloveho timetable|Michaela|Jan|STREDNI|CEKA|Vizualne muze byt super prvek
P-041|Priprava|Hudba a elektro|Udelat plny zvukovy test primo na miste|Hlasitost, dosah, vypinani, nabijeni, umisteni bedny|Walter|Jan|KRITICKA|CEKA|Idealne uz v utery vecer
P-042|Priprava|Jidlo a piti|Potvrdit kdo vari co a v jakem mnozstvi|Sele, hovezi gulas, hlivovy gulas, korytko|Jan|otec zenicha|KRITICKA|ROZPRACOVANO|Sepsat porce a cas pripravy
P-043|Priprava|Jidlo a piti|Sepsat nakupni seznam pro den 1|Maso, prilohy, pecivo, omacky, zelenina, snacky|Jan|Walter|KRITICKA|CEKA|Rozdelit na nakup dopredu a posledni den
P-044|Priprava|Jidlo a piti|Sepsat nakupni seznam pro den 2|Uzeniny, syry, pecivo, gril veci, dojezd a zbytky|Walter|Jan|VYSOKA|CEKA|Nakoupit tak akorat
P-045|Priprava|Jidlo a piti|Zajistit sele gril, plynovy gril a plyn|Overit funkcnost, rosty, naradi a zapaleni|Jan|otec zenicha|KRITICKA|CEKA|Pridat do materialu
P-046|Priprava|Jidlo a piti|Zajistit drevo, podpalovac a zapalovace|Na ohen i na grill rezim|Walter|parta na stavbu|VYSOKA|CEKA|Myslet i na mokre drevo
P-047|Priprava|Jidlo a piti|Zajistit becky, vycep a chlazeni|Zjistit co presne je potreba k fungovani samoobsluhy|Jan|Walter|KRITICKA|CEKA|Vymena kdyz dojde
P-048|Priprava|Jidlo a piti|Navrhnout samoobsluzny alkoholovy kout|Tvrdsi alkohol, vino, otviraky, odkladaci misto, voda vedle|Walter|Jan|STREDNI|CEKA|At to funguje bez obsluhy
P-049|Priprava|Jidlo a piti|Zajistit nealko, vodu, kavu, caj a led|Dostatek vody je kriticky hlavne u ohnu a alkoholu|Jan|Walter|KRITICKA|CEKA|Voda musi byt viditelna
P-050|Priprava|Jidlo a piti|Zajistit jednorazove talire, misky, pribory a kelimky|Low cost format potrebuje dostatecnou rezervu|Walter|Michaela|KRITICKA|CEKA|Myslet i na kavu a dezert
P-051|Priprava|Jidlo a piti|Zajistit ubrousky, omacky, chleba a zeleninu|Drobnosti delaji dojem a chybi nejdriv|Michaela|Jan|VYSOKA|CEKA|Pridat i sul, pep, kecup, horcici
P-052|Priprava|Jidlo a piti|Rozdelit sluzby na prubezne doplnovani|Kdo hlida pivo, vodu, led, talire a odpad|Walter|Michaela|KRITICKA|CEKA|At neni vse na jednom cloveku
P-053|Priprava|Atmosfera|Potvrdit styl vyzdoby louky a obradu|Neformalni festivalova opekacka, smart casual, bez zbytecneho prezdobeni|Michaela|Veronika|VYSOKA|ROZPRACOVANO|Drzet jednoduchy smer
P-054|Priprava|Atmosfera|Sesbirat kvetiny a dekor material|Co se vyrabi svepomoci a co se pujcije|Michaela|sestra zenicha|STREDNI|CEKA|Rozdelit na utery a stredu
P-055|Priprava|Atmosfera|Rozhodnout hostbook, polaroid nebo audio vzkazy|Vybrat jen to co opravdu budete pouzivat|Michaela|Veronika|STREDNI|CEKA|Nekombinovat moc veci
P-056|Priprava|Atmosfera|Vyrobit misto na penize jako piratskou truhlu|Jednoduche, vkusne, viditelne, ale ne trapne|Michaela|Jan|STREDNI|CEKA|Pridat kratkou cedulku
P-057|Priprava|Osobni veci|Finalizovat saty, oblek, boty a doplnky|Co chybi dokoupit nebo vyzvednout|Veronika|Jan|KRITICKA|ROZPRACOVANO|At nic nechybi v pondeli
P-058|Priprava|Osobni veci|Pripravit nouzovy kit pro par|Naplasti, leky, siticko, deodorant, powerbanka, kapesniky|Michaela|Walter|VYSOKA|CEKA|Dat do jedne tasky
P-059|Priprava|Osobni veci|Pripravit osobni checklist pro Jana|Obleceni, hygiena, nabijecky, doklady, boty, bunda|Jan||VYSOKA|CEKA|Vyuzij list 09_Baleni
P-060|Priprava|Osobni veci|Pripravit osobni checklist pro Veroniku|Saty, druhe boty, hygiena, kosmetika, karticky se sliby|Veronika|Michaela|VYSOKA|CEKA|Vyuzij list 09_Baleni
P-061|Priprava|Foto a video|Potvrdit kameramana a jeho casy|Kdy prijede, co toci, kdy potrebuje par a rodinu|Jan|Walter|VYSOKA|ROZPRACOVANO|Aspon hruby plan
P-062|Priprava|Foto a video|Rozhodnout zda bude dobrovolny fotograf|Kdyz nebude profi foto, nekdo musi hlidat skupinove fotky|Michaela|Jan|STREDNI|CEKA|Muze byt jeden kamarad
P-063|Priprava|Foto a video|Sepsat must have fotky a videa|Par, rodice, svedci, sourozenci, skupina pratel, detail louky|Michaela|kameraman|STREDNI|CEKA|Jeden list staci
P-064|Priprava|Hygiena a bezpeci|Nakoupit zasobu pro zachody a sprchu|Toaletak, mydlo, papir, rucniky, pytel na pradlo|Walter|Jan|VYSOKA|CEKA|Ve vnitrku bude provoz
P-065|Priprava|Hygiena a bezpeci|Postavit system na trideni odpadu|Smes, plast, sklo, plechovky a bio podle reality mista|Walter|parta na uklid|KRITICKA|CEKA|Viditelne pytle a popisky
P-066|Priprava|Hygiena a bezpeci|Pripravit uklidovy kit|Rukavice, pytle, jar, houby, hadry, smetak, lopatka|Walter|Michaela|VYSOKA|CEKA|Hodi se uz behem akce
P-067|Priprava|Hygiena a bezpeci|Pripravit comfort a safety kit|Lekarnicka, repelent, krem, deky, baterky, zapalovac|Michaela|Walter|VYSOKA|CEKA|Dat na jedno dostupne misto
P-068|Priprava|Logistika|Vytvorit master seznam vseho co se musi dovest|Pujcene veci, jidlo, piti, elektro, dekor, osobni tasky|Jan|Walter|KRITICKA|CEKA|Provazat s listem 07_Material
P-069|Priprava|Logistika|Rozdelit kdo priveze ktere veci|Auto, vlecny vozik, kolikrat se jede, kdo co nalozi|Walter|Jan|KRITICKA|CEKA|Idealne s casem odjezdu
P-070|Priprava|Logistika|Naplanovat uterni vykladku a stavbu|Po prijezdu v 17:00 jit systematicky po zonach|Walter|parta na stavbu|KRITICKA|CEKA|Neimprovizovat na miste
P-071|Priprava|Logistika|Naplanovat patecni vratky pujceneho vybaveni|Kdo co odvazi, komu se to vraci a do kdy|Michaela|Walter|VYSOKA|CEKA|Usetri stres pri predani
P-072|Priprava|Finance|Pripravit hotovost a drobne|Na nenadale nakupy, benzinku, led a drobnosti|Jan||STREDNI|CEKA|Dat do jedne obalky
P-073|Priprava|Finance|Vyplnit rozpocet a hlidat strop 125000 CZK|Vsechny hlavni polozky nahodit aspon odhadem|Jan|Veronika|VYSOKA|CEKA|Vyuzij list 06_Rozpocet
P-074|Priprava|Program|Rozhodnout jestli budou kratke proslovy|Doporuceni je max 2 az 3 vstupy po 2 minutach|Michaela|Jan|STREDNI|CEKA|Nepretahovat to
P-075|Priprava|Program|Rozhodnout jestli bude prvni tanec|Kdyz ano tak jeden song bez zbytecne oficiality|Veronika|Jan|STREDNI|CEKA|Muze byt jen symbolicky start party
P-076|Priprava|Program|Vybrat 2 az 3 nenasilne hry na den 2|Napriklad hudebni bingo, karty s ukoly, ring toss u ohniste|Walter|Michaela|NIZKA|CEKA|Jen kdyz na to bude chut
P-077|Priprava|Program|Uzamknout cil pripravenosti do stredy 12:00|Co musi byt hotove nejpozdeji pred prijezdem hostu|Walter|Jan|KRITICKA|CEKA|Sepsat jako mini plan
P-078|Priprava|Program|Udelat finalni mikrotym check v utery vecer|Projit co zbyva na stredu rano a kdo co drzi|Jan|Walter|KRITICKA|CEKA|10 minut staci
"@

$day1Tasks = ConvertFrom-PipeBlock -Columns $taskColumns -Block @"
D1-001|Den 1|Koordinace|Ranni mikrotym briefing|V 08:30 si rict kdo drzi ktere bloky a co je kriticke|Walter|Jan|KRITICKA|NENI HOTOVO|Max 10 minut
D1-002|Den 1|Koordinace|Rozhodnout rezim podle pocasi|Potvrdit venek versus destovy fallback uvnitr|Jan|Michaela|KRITICKA|NENI HOTOVO|Nejpozdeji rano
D1-003|Den 1|Misto|Ranni sweep louky|Posbirat drobnosti, srovnat travu a zkontrolovat cesty|Walter|parta na stavbu|VYSOKA|NENI HOTOVO|At prvni dojem neni bordel
D1-004|Den 1|Misto|Finalni kontrola stanu a sezeni|Napeti, stabilita, rozmisteni a stin|Walter|Jan|KRITICKA|NENI HOTOVO|Hlavne obrad a jidlo
D1-005|Den 1|Misto|Vyznacit parkovani a prijezdovou cestu|At hoste nebloudi a neblokujou provoz|Jan|ridic|VYSOKA|NENI HOTOVO|Cedule nebo clovek na miste
D1-006|Den 1|Misto|Vyvesit cedule a pripravit visacky|Zachod, voda, darky, odpad, ohen a jmenovky|Michaela|Veronika|STREDNI|NENI HOTOVO|Hotove pred 13:00
D1-007|Den 1|Obrad|Doladit obradni zonu|Misto pro starostku, vas dva, svedky a hosty|Michaela|sestra zenicha|KRITICKA|NENI HOTOVO|Nic moc oficialniho
D1-008|Den 1|Obrad|Pripravit tasku obradu|Prstynky, sliby, kapesniky, voda, kontakt na starostku|Walter|Jan|KRITICKA|NENI HOTOVO|Jedna taska na jednom miste
D1-009|Den 1|Hudba a elektro|Posledni test hudby a elektriny|Spusteni playlistu, hlasitost, nabijeni a dosah bedny|Jan|Walter|KRITICKA|NENI HOTOVO|Pred prijezdem hostu
D1-010|Den 1|Jidlo a piti|Zkontrolovat grily a gastro flow|Co se rozehreva, kdy se servira a kdo doplnuje|Jan|otec zenicha|KRITICKA|NENI HOTOVO|At je jasno co se deje
D1-011|Den 1|Jidlo a piti|Otevrit pitny bod a samoobsluzny bar|Voda musi byt hned viditelna, nealko oddelene od tvrdeho|Walter|Michaela|KRITICKA|NENI HOTOVO|Voda a kelimky na jednom miste
D1-012|Den 1|Hygiena a bezpeci|Pripravit zachody, mydlo a odpad|Zkontrolovat toaletak, pytle, rucniky a pristup|Michaela|Walter|VYSOKA|NENI HOTOVO|Prvni vlna hostu to proveri
D1-013|Den 1|Koordinace|Udrzet Veroniku mimo provozni stres|V den svatby nic netahat a nic nehasit|Michaela|Walter|KRITICKA|NENI HOTOVO|Mikrotym filtruje problemy
D1-014|Den 1|Koordinace|Dat Janovi 30 minut klid na pripravu|Pred obradni casti bez logistickych dotazu|Walter|Michaela|VYSOKA|NENI HOTOVO|Minimalni narok
D1-015|Den 1|Hoste|Vitat hosty od 13:00|Ukazat parkovani, louku, zachod, pitny rezim a badge|Jan|Michaela|VYSOKA|NENI HOTOVO|Prvni dojem dela pulku akce
D1-016|Den 1|Hoste|Navadet parkovani|At auta neprekazi ani nezablokuji vyjezd|ridic|Walter|STREDNI|NENI HOTOVO|Muze byt i operativni
D1-017|Den 1|Hoste|Rozdat visacky a zakladni info|Kdo je kdo a kde co je|Michaela|Veronika|STREDNI|NENI HOTOVO|Zvlast pomuze den 2
D1-018|Den 1|Hoste|Posadit starsi rodinu k obradu pohodlneji|Lavice nebo zidlicky ve stinu|Michaela|parta na stavbu|VYSOKA|NENI HOTOVO|At nikdo nestoji zbytecne
D1-019|Den 1|Obrad|Srovnat nastup a start 14:42|Kdo dava signal starostce a kdo pousti hudbu|Walter|Jan|KRITICKA|NENI HOTOVO|Musite vedet posledni 2 minuty
D1-020|Den 1|Obrad|Poustet obradni hudbu v pravy cas|Nastup, mezibloky a odchod|Jan|Walter|KRITICKA|NENI HOTOVO|Mit pripravenou frontu skladeb
D1-021|Den 1|Obrad|Pohlidat prstynky a sliby|At jsou dostupne a neztrati se v chaosu|Walter|Michaela|KRITICKA|NENI HOTOVO|Jedna osoba za to ruci
D1-022|Den 1|Obrad|Usmernit gratulace po obradu|At neni zmatek a hned navazuje foto a jidlo|Michaela|Walter|STREDNI|NENI HOTOVO|Staci slovni navod
D1-023|Den 1|Foto a video|Odchytat klicove skupinove fotky|Rodice, svedci, sourozenci, parta, cela svatba|Michaela|kameraman|VYSOKA|NENI HOTOVO|Bez seznamu to utece
D1-024|Den 1|Jidlo a piti|Otevrit pohoosteni hned po fotkach|Oznamit ze jidlo je prubezne a neceka se na slavnostni obed|Jan|Walter|KRITICKA|NENI HOTOVO|At je jasne jak format funguje
D1-025|Den 1|Jidlo a piti|Prubezne doplnovat jidlo a pecivo|Sele, gulase, korytko, chleba, omacky, zelenina|otec zenicha|Walter|KRITICKA|NENI HOTOVO|Nenechat stoly prazdne
D1-026|Den 1|Jidlo a piti|Prubezne doplnovat piti a led|Pivo, voda, nealko, tvrde, kava|Walter|Jan|KRITICKA|NENI HOTOVO|Nekdo to musi hlidat cele odpoledne
D1-027|Den 1|Program|Kratky blok pripitku a proslovu|Doporuceny blok kolem 17:30, max 10 az 15 minut celkem|Michaela|Jan|STREDNI|NENI HOTOVO|Klidne preskocit kdyz nebude chut
D1-028|Den 1|Program|Spustit prvni tanec nebo startovaci song|Jedna pisnicka jako symbolicky prechod do vecerni casti|Jan|Veronika|NIZKA|NENI HOTOVO|Volitelne
D1-029|Den 1|Hudba a elektro|Prepinat playlisty po hodinach|Drzet festivalovy format jak na plakatu|Jan|Walter|VYSOKA|NENI HOTOVO|Muze se menit podle nalady
D1-030|Den 1|Foto a video|Udelat golden hour slot pro par|Kratsi vybehnuti kolem zapadu slunce|kameraman|Michaela|STREDNI|NENI HOTOVO|15 minut bohate staci
D1-031|Den 1|Atmosfera|Pripominat guestbook, polaroid a darky|Lidi to sami od sebe casto minou|Michaela|parta pratel|NIZKA|NENI HOTOVO|Staci nenapadne
D1-032|Den 1|Hygiena a bezpeci|Hlidal bezpecny rezim u ohne|Drevo, rozestupy, deti, voda pobliz|Walter|parta u ohne|VYSOKA|NENI HOTOVO|Nejen kvuli detem
D1-033|Den 1|Jidlo a piti|Doplnit pozde vecer vodu a neco k jidlu|At neni jen alkohol a uhel|Walter|Jan|STREDNI|NENI HOTOVO|Pomuze druhemu dni
D1-034|Den 1|Hoste|Roztridit kdo kde spi a kdo potrebuje klice|Chata, stany, odjezd domu|Michaela|Walter|VYSOKA|NENI HOTOVO|Ve dve rano je pozde to resit
D1-035|Den 1|Koordinace|Udelat kratky nocni sweep bordelu|Posbirat nejhorsi kelimky a odpad kolem centra akce|Walter|parta na uklid|STREDNI|NENI HOTOVO|Ulehci ctvrtek
D1-036|Den 1|Jidlo a piti|Oznacit a ulozit zbytky pro den 2|Co se muze jist zitra, co patri do chladu a co je na vyhozeni|Jan|Walter|VYSOKA|NENI HOTOVO|Nehadat se rano co je jeste dobre
D1-037|Den 1|Hudba a elektro|Ochranit techniku pred rosou nebo destem|Bedna, kabely, mobil, nabijeni|Walter|Jan|VYSOKA|NENI HOTOVO|Na noc pod strechu
D1-038|Den 1|Hudba a elektro|Vypnout co nemusi jet pres noc|Setrit elektrinu a zmensit riziko|Jan|Walter|STREDNI|NENI HOTOVO|Nechat jen co je potreba
D1-039|Den 1|Koordinace|Zapsat co rano prvni resit na den 2|Doplneni piti, uklid, novy nakup, program|Walter|Michaela|VYSOKA|NENI HOTOVO|Dve minuty pred spankem
"@

$day2Tasks = ConvertFrom-PipeBlock -Columns $taskColumns -Block @"
D2-001|Den 2|Koordinace|Nastartovat rano vodu, kavu a klidny rozjezd|At druhy den neni hned stres a vsichni vedi kde je voda|Walter|Jan|VYSOKA|NENI HOTOVO|Pomaha i po vecernim piti
D2-002|Den 2|Jidlo a piti|Projit zbytky z dne 1 a rozhodnout co se ji|Roztridit bezpecne versus vyhodit|Jan|Walter|KRITICKA|NENI HOTOVO|Nenechat na stole cely den
D2-003|Den 2|Hygiena a bezpeci|Zkontrolovat ohen, grill a popel po noci|Bezpecny restart druheho dne|Walter|parta u ohne|VYSOKA|NENI HOTOVO|Pred dalsim rozpalovanim
D2-004|Den 2|Misto|Preskladat chill zony a sezeni|Druhy den bude volnejsi a s jinymi lidmi|Michaela|parta pratel|STREDNI|NENI HOTOVO|Vic prostoru kolem ohne
D2-005|Den 2|Hygiena a bezpeci|Doplnit zachody, mydlo a pytle|Druhy den byva nejvetsi propad v hygiene|Walter|Michaela|VYSOKA|NENI HOTOVO|Rano prvni kontrola
D2-006|Den 2|Jidlo a piti|Doplnit piti, vodu a led|Pred prijezdem dalsich lidi musi byt vse ready|Jan|Walter|KRITICKA|NENI HOTOVO|Hlavne voda
D2-007|Den 2|Jidlo a piti|Nakoupit uzeniny, syry a pecivo na grill|Co nebylo nakoupeno dopredu doresit rano|Walter|ridic|VYSOKA|NENI HOTOVO|Muze byt i dopoledni vyjezd
D2-008|Den 2|Hoste|Privitat nove prichozi hosty|Vysvetlit jak funguje samoobsluha a kde co je|Michaela|Jan|STREDNI|NENI HOTOVO|Druhy den je vic fluidni
D2-009|Den 2|Hoste|Rict hostum ze den 2 je volnejsi|Bez pevne hostiny, grill, zbytky a dobra nalada|Jan|Walter|STREDNI|NENI HOTOVO|Zabrani zbytecnym dotazum
D2-010|Den 2|Jidlo a piti|Rozbehnout open fire a grill rezim|Bezpecne rozdelat, pripravit rosty a naradi|Walter|parta u ohne|VYSOKA|NENI HOTOVO|At to jede bez kolony
D2-011|Den 2|Hudba a elektro|Pokracovat v hudebnim timetable|Druhy den muze byt punkovejsi a uvolnenejsi|Jan|parta pratel|STREDNI|NENI HOTOVO|Muze se vic improvizovat
D2-012|Den 2|Program|Vybrat jednu jednoduchou skupinovou hru|Jen kdyz bude dobra nalada a chut|Walter|Michaela|NIZKA|NENI HOTOVO|Neni povinne
D2-013|Den 2|Program|Odehrat kratky hravy blok|Napriklad hudebni bingo nebo karty s ukoly|Michaela|Walter|NIZKA|NENI HOTOVO|Spis pro rozhybani
D2-014|Den 2|Foto a video|Stahnout nebo sesbirat fotky od hostu|QR slozka, AirDrop, sdileny album nebo messenger|Jan|Michaela|NIZKA|NENI HOTOVO|At se nic neztrati
D2-015|Den 2|Atmosfera|Pripomenout guestbook a vzkazy|Druhy den jsou lidi casto vic otevreni|Michaela|parta pratel|NIZKA|NENI HOTOVO|Staci jemne
D2-016|Den 2|Jidlo a piti|Rozdelit co zustava na patek rano|Voda, kava, drobne jidlo pro uklidovou partu|Jan|Walter|STREDNI|NENI HOTOVO|Nenechat vse sezrat vecer
D2-017|Den 2|Uklid prubezne|Sbalit dekor co uz neni potreba|Cokoliv co nebude potreba na ctvrtecni vecer|Michaela|sestra zenicha|STREDNI|NENI HOTOVO|Usetri patek rano
D2-018|Den 2|Uklid prubezne|Umyt a uklidit znovupouzitelne veci uz ve ctvrtek|Noze, prkenka, misky, naradi, svetla|Walter|parta na uklid|VYSOKA|NENI HOTOVO|Patek bude kratky
D2-019|Den 2|Uklid prubezne|Delat prubezny sweep louky|Kelimky, lahve a bordel sbirat bez cekani na patek|Walter|parta na uklid|VYSOKA|NENI HOTOVO|Kazde dve hodiny
D2-020|Den 2|Uklid prubezne|Prubezne menit pytle a tridit odpad|At vas patek nezabije logistika odpadu|Walter|parta na uklid|VYSOKA|NENI HOTOVO|Oznacit pytle
D2-021|Den 2|Koordinace|Rozdelit patecni uklidove zony mezi lidi|Louka, gastro, hygiena, spani, technika, pujcene veci|Michaela|Walter|KRITICKA|NENI HOTOVO|Idealne ve ctvrtek vecer
D2-022|Den 2|Hudba a elektro|Schovat techniku na noc do sucha|Bedna, kabely, nabijeni a svetla|Jan|Walter|VYSOKA|NENI HOTOVO|Ne nechat venku
D2-023|Den 2|Koordinace|Pripravit vodu a male jidlo na patek rano|At uklidova parta funguje hned od rana|Jan|Michaela|STREDNI|NENI HOTOVO|Kava plus neco do ruky
D2-024|Den 2|Koordinace|Ctvrtecni nocni debrief mikrotymu|Co je hotove, co zbyva, kdo vstava v kolik|Walter|Jan|KRITICKA|NENI HOTOVO|Bez toho bude patek chaos
"@

$cleanupTasks = ConvertFrom-PipeBlock -Columns $taskColumns -Block @"
U-001|Uklid|Koordinace|Zacit v patek v 07:30 ostre|Nevyspat se do deviti, jinak nestihnete predani|Walter|Jan|KRITICKA|NENI HOTOVO|Budik pro mikrotym
U-002|Uklid|Koordinace|Rozdelit lidi po zonach|Louka, gastro, hygiena, spani, technika, pujcene veci|Michaela|Walter|KRITICKA|NENI HOTOVO|Nenechat vsechny delat vsechno
U-003|Uklid|Louka|Posbirat vsechny kelimky, lahve a plechovky|Hlavni plocha, okoli ohne, kolem stanu a pristupove cesty|parta na uklid|Walter|KRITICKA|NENI HOTOVO|Prvni velka vlna
U-004|Uklid|Louka|Vysbirat drobny bordel a vajgly|To je to co majitel vidi nejvic|parta na uklid|Michaela|KRITICKA|NENI HOTOVO|Vzit rukavice
U-005|Uklid|Louka|Sundat dekor a cedule|Nic nesmi zustat na stromech, stanech ani ceste|Michaela|sestra zenicha|VYSOKA|NENI HOTOVO|Rovnou tridit co vratit
U-006|Uklid|Louka|Sbalit guestbook, polaroid, truhlu a visacky|Drobnosti se ztraceji nejrychleji|Michaela|Jan|STREDNI|NENI HOTOVO|Dat do jedne bedny
U-007|Uklid|Louka|Uhasit a vycistit ohniste|Popel, rosty, zbytky dreva a okolni plocha|Walter|parta u ohne|KRITICKA|NENI HOTOVO|Bezpecnost i dojem
U-008|Uklid|Gastro|Vycistit sele gril, plynovy gril a naradi|Roasty, klece, noze, prkenka, mastnota|otec zenicha|Walter|VYSOKA|NENI HOTOVO|Nechat vychladnout vcas
U-009|Uklid|Gastro|Vypustit a uklidit pitny kout a vycep|Odpojit, ocistit, zabalit, oddelit vratne|Jan|Walter|KRITICKA|NENI HOTOVO|Myslet na becky a hadice
U-010|Uklid|Gastro|Roztridit zbytky jidla|Co se bere domu, co se rozdeli lidem, co se vyhodi|Jan|Michaela|VYSOKA|NENI HOTOVO|At nic netece v aute
U-011|Uklid|Odpad|Rozhodnout keep versus vyhodit|Nevozit domu zbytecnosti jen z lhostejnosti|Walter|Jan|STREDNI|NENI HOTOVO|Drzet se reality
U-012|Uklid|Odpad|Vynest a uzavrit vsechny pytle s odpadem|Kazda frakce zvlast a pevne zavazana|parta na uklid|Walter|KRITICKA|NENI HOTOVO|At se nerozsype v aute
U-013|Uklid|Odpad|Nalozit odpad k odvozu nebo do kontejneru|Podle domluvy s chatou a mistniho rezimu|Walter|ridic|KRITICKA|NENI HOTOVO|Overit predem kam s tim
U-014|Uklid|Louka|Otrit a slozit stoly a lavice|Vse vratit v rozumnem stavu|parta na uklid|Walter|VYSOKA|NENI HOTOVO|Nejen odnest ale i ocistit
U-015|Uklid|Louka|Slozit party stany a nechat je suche|Kdyz budou mokre aspon je oznacit a resit hned|parta na stavbu|Walter|KRITICKA|NENI HOTOVO|Pujcene veci musi byt pod kontrolou
U-016|Uklid|Pujcene veci|Porovnat vratky proti listu materialu|Co jste si pujcili, to musi byt zpet a kompletni|Michaela|Jan|KRITICKA|NENI HOTOVO|Vyuzij list 07_Material
U-017|Uklid|Pujcene veci|Nalozit pujcene veci podle tras a majitelu|At nevznikne jedna mega hromada bez adresy|Walter|ridic|KRITICKA|NENI HOTOVO|Skupit podle komu vracite
U-018|Uklid|Louka|Uklidit parkovani a prijezdovou trasu|Cedule, pasky, zapomenute veci a bordel|parta na uklid|Jan|STREDNI|NENI HOTOVO|At misto vypada normalne
U-019|Uklid|Chata a hygiena|Vycistit zachody, sprchu a koupelnu|Toaletak, kos, spinavy textil a umyvadlo|parta na uklid|Michaela|KRITICKA|NENI HOTOVO|Majitel si vsimne hned
U-020|Uklid|Chata a hygiena|Zkontrolovat pokoje a lozni casti|Zda nezustalo pradlo, lahve nebo drobnosti|Michaela|parta na uklid|VYSOKA|NENI HOTOVO|Hlavne po hostech
U-021|Uklid|Chata a hygiena|Projit stany a najit zapomenute veci|Powerbanky, bundy, boty, kosmetika|Walter|parta pratel|STREDNI|NENI HOTOVO|Udelat fotku nalezu
U-022|Uklid|Technika|Sesbirat prodluzovacky, svetla a audio|Sbalit systematicky a po majitelich|Jan|Walter|KRITICKA|NENI HOTOVO|Nesmotat do jednoho uzlu
U-023|Uklid|Technika|Sesbirat kuchynske a servirovaci vybaveni|Misky, noze, prkna, otviraky, kleste, naberacky|Jan|otec zenicha|VYSOKA|NENI HOTOVO|At nic nezustane pod stanem
U-024|Uklid|Chata a hygiena|Vratit nabytek a vnitrni prostory do puvodniho stavu|Co se presunulo, vratit tam kde to bylo|parta na uklid|Jan|VYSOKA|NENI HOTOVO|Fotka pred by pomohla
U-025|Uklid|Chata a hygiena|Udelat finalni sweep uvnitr chaty|Kuchyn, podlaha, rohy, kos, lednice|Michaela|Walter|KRITICKA|NENI HOTOVO|Pred posledni prohlidkou
U-026|Uklid|Louka|Udelat finalni sweep venku|Pohled od prijezdu, okoli louky i okoli chaty|Walter|Jan|KRITICKA|NENI HOTOVO|Tady se lomi dojem
U-027|Uklid|Predani|Udelat predavaci kolecko s majitelem do 10:45|Mit rezervu do 11:00 a nic neresit na posledni minutu|Jan|Michaela|KRITICKA|NENI HOTOVO|Pripadne skody resit hned
U-028|Uklid|Uzavreni|Uzavrit checklist, rozpocet a rozeslat podekovani pomocnikum|Kratke diky po akci a poznamky pro vas dva|Jan|Veronika|NIZKA|NENI HOTOVO|Hodi se i pro vzpominku
"@

$scheduleRecords = ConvertFrom-PipeBlock -Columns $scheduleColumns -Block @"
23.6.|17:00|17:30|Prijezd mikrotymu|Prijezd, rychla obhlidka mista a rozdeleni prvni prace|Jan + Walter|Bez dlouheho vymysleni
23.6.|17:30|18:30|Vykladka|Sundat z aut vse k elektrine, stanum, sezeni a oznacit hromady|Walter|Zacit od kritickych veci
23.6.|18:30|20:00|Stavba zakladu|Party stany, lavicky, hlavni zony louky, cesta k chate|Walter + parta|Co jde mit hotove uz v utery
23.6.|20:00|21:00|Elektro a svetla|Trasa kabelu, prvni test zasuvek a svetel|Jan|At se rano jen doladuje
23.6.|21:00|21:30|Hudba test|Kratky test bedny, telefonu a playlistu|Jan|Neplest s finalnim testem
23.6.|21:30|22:00|Mikrotym check|Co zbyva na stredu rano a kdo co drzi|Jan + Walter + Michaela|Uzamknout stredocni plan
24.6.|08:30|08:45|Briefing|Ranni mikrotym, pocasi, krizove body|Walter|10 minut a dost
24.6.|08:45|10:00|Finalni setup|Louka, cedule, sezeni, zachody, odpad, pitny bod|Walter + Michaela|Hoste prijedou od 13:00
24.6.|10:00|11:00|Hudba a obrad|Posledni test hudby, obradni taska, sliby, prstynky|Jan + Walter|Uz nic nenechavat bez pana
24.6.|11:00|12:00|Gastro|Spustit grill flow, pitny kout a naskladnit zaklad|Jan|Voda musi byt ready
24.6.|12:00|12:30|Klid pro par|Jan jen rychla priprava, Veronika bez provoznich ukolu|Michaela + Walter|Nikdo nic neresi pres nevestu
24.6.|13:00|14:15|Prijezd hostu|Parkovani, vitani, visacky, welcome rezim|Jan + Michaela|Volne a pratelske
24.6.|14:20|14:40|Nastup k obradu|Srovnat hosty, starostku a hudbu|Walter|Drzet cas
24.6.|14:42|15:05|Obrad|Symbolicky uredni obrad na louce|Starostka + mikrotym|Hlavni bod dne
24.6.|15:05|15:30|Gratulace|Volny prostor na obejmuti a pripitky|Michaela|Bez tlaku
24.6.|15:30|16:15|Foto|Skupinovky plus kratky slot pro vas dva|Michaela + kameraman|Neprotahovat
24.6.|16:15|17:30|Pohoosteni|Otevrene jidlo bez formalni hostiny|Jan + gastro|Lidi jedi kdy chteji
24.6.|17:30|17:45|Kratke proslovy|Doporuceny blok max 10 az 15 minut|Michaela|Volitelne
24.6.|18:30|18:35|Start party song|Prvni tanec nebo jen jeden symbolicky track|Jan + Veronika|Volitelne
24.6.|19:00|22:00|Vecerni volny blok|Hudba po hodinach, jidlo, ohen, mluveni, tanec|Jan + Walter|Festivalovy format
24.6.|20:30|20:50|Golden hour|Kratky slot na video nebo fotku pri zapadu|kameraman|Jen pokud bude chut
24.6.|22:00|00:30|Ohen a noc|Pozdni vecer, voda, zbytky jidla, chill|Walter|Hlidat bezpeci
24.6.|00:30|01:00|Nocni sweep|Nejhorsi bordel a priprava na druhy den|Walter + parta|Zachrani ctvrtek
25.6.|09:30|10:30|Pomaly start|Kava, voda, recovery a lehky uklid|Walter|Bez stresu
25.6.|10:30|11:30|Restart mista|Doplnit zachody, odpad, piti a zbytky|Jan + Walter|Pripravit den 2
25.6.|12:00|13:00|Prijezdy den 2|Volne vitani dalsich lidi|Michaela|Jina skupina hostu
25.6.|13:00|15:00|Grill a chill|Ohen, zbytky, uzeniny, syry a hudba|Walter + parta|Volnejsi format
25.6.|15:00|16:00|Hravy blok|Jedna jednoducha hra jen kdyz bude nalada|Michaela|Volitelne
25.6.|16:00|18:00|Volny program|Hudba, fotky, guestbook, kecani, odpocinek|Jan + Michaela|Nechat dychat
25.6.|18:00|21:00|Punk vecer|Druha vecerni vlna u ohne a s hudbou|Jan + Walter|Klidne ostrejsi
25.6.|21:00|22:00|Predbaleni|Sbalit co uz nebude potreba na patek rano|Walter + Michaela|Usnadnit uklid
25.6.|22:00|23:30|Prubezny uklid|Pytle, kelimky, technika do sucha, rozdeleni zon|Walter|Patek bude kratky
25.6.|23:30|23:45|Debrief|Kdo vstava kdy a co prvni resi|Walter + Jan|Povinne
26.6.|07:30|07:45|Start uklidu|Rozdeleni lidi po zonach a tempo|Walter|Bez velke debaty
26.6.|07:45|09:15|Velky uklid venku|Louka, ohen, stoly, odpad, technika, pujcene veci|Vsechny skupiny|Jit po zonach
26.6.|09:15|10:00|Velky uklid uvnitr|Zachody, sprcha, pokoje, kuchyn|Michaela + parta|Nezapomenout rohy
26.6.|10:00|10:30|Nakladka a vratky|Pujcene veci, odpad, posledni auta|Jan + Walter|Vse podle seznamu
26.6.|10:30|10:45|Finalni kontrola|Projit s majitelem venek i chatu|Jan|Rezerva pred 11:00
26.6.|11:00|11:00|Predani|Objekt predan a muzete odjet|Jan + Michaela|Konec mise
"@

$budgetRecords = ConvertFrom-PipeBlock -Columns $budgetColumns -NumericColumns @(2, 3) -Block @"
Misto|Chata a energie|20000||Jan|ROZPRACOVANO|Chata uz je zajistena, doplnit realnou castku
Administrativa|Matrika a poplatky|3000||Jan|ROZPRACOVANO|Vcetne drobnych uradnich vydaju
Pujcene veci|Party stany, lavice a sezeni|10000||Jan|CEKA|Low cost ale klicove pro komfort
Technika|Hudba, elektro, svetla a kabely|5000||Jan|CEKA|I kdyz si hodne pujcite
Jidlo|Den 1 maso, gulase, korytko|20000||Jan|ROZPRACOVANO|Hlavni jidlo dne
Jidlo|Den 2 grill, uzeniny, syry, pecivo|6000||Walter|CEKA|Volnejsi den
Piti|Pivo, nealko, voda, led, kava, caj|14000||Jan|CEKA|Nejspis poroste
Ohen a grill|Drevo, plyn, podpalovac, drobne gastro|4000||Walter|CEKA|Nezapomenout na rezervu
Atmosfera|Vyzdoba, tisk, visacky, cedule, truhla|4000||Michaela|CEKA|Drzet jednoduchy styl
Osobni veci|Saty, oblek, boty, beauty a doplnky|15000||Veronika|ROZPRACOVANO|Doplnit co jeste chybi
Foto a video|Kameraman a foto drobnosti|6000||Jan|ROZPRACOVANO|Pokud nebude fotograf
Hygiena a odpad|Pytle, toaletak, mydlo, uklid, comfort kit|3000||Walter|CEKA|Podcenene ale nutne
Logistika|Palivo, nakupy, drobne prevozy|5000||Jan|CEKA|Hlavne pred akci
Rezerva|Necekane vydaje|10000||Jan|CEKA|Nesahat na ni bez duvodu
"@

$materialRecords = ConvertFrom-PipeBlock -Columns $materialColumns -Block @"
Misto|Party stan 3x3 nebo podobny|2 az 3 ks||Jan|Walter|CEKA|Podle poctu sezeni a deste
Misto|Pivni sety|4 az 6 sad||Walter|Walter|CEKA|Podle finalniho poctu hostu
Misto|Lavicky nebo zidlicky navic|8 az 12 mist||Walter|Walter|CEKA|Hlavne pro rodinu a starsi
Misto|Skladaci stoly navic|2 ks||Walter|Walter|CEKA|Na pitny kout a jidlo
Misto|Osvetleni na louku|1 sada||Michaela|Michaela|CEKA|Lampicky nebo reflektory
Misto|Cedule na parkovani a zachody|6 az 10 ks||Michaela|Michaela|CEKA|Jednoduchy tisk staci
Hudba a elektro|Hlavni JBL nebo jina bedna|1 ks||Jan|Jan|CEKA|Kriticka polozka
Hudba a elektro|Zalozni bedna nebo reproduktor|1 ks||Walter|Walter|CEKA|Pro pripad vypadku
Hudba a elektro|Telefon nebo notebook na hudbu|2 ks||Jan|Jan|CEKA|Hlavni plus zaloha
Hudba a elektro|Nabijecky a powerbanky|4 ks||Jan|Jan|CEKA|At hudba nepadne
Hudba a elektro|Prodluzovacka dlouha trasa|1 ks 100 m||Walter|Walter|CEKA|Na louku od chaty
Hudba a elektro|Prodluzovacky kratke|4 az 6 ks||Walter|Walter|CEKA|Ruzne zony
Hudba a elektro|Rozdvojky a listy|6 az 8 ks||Walter|Walter|CEKA|K hudbe a svetlum
Hudba a elektro|Gaffa a stahovaci pasky|1 sada||Walter|Walter|CEKA|Na kabely
Jidlo a piti|Sele gril|1 ks||Jan|Jan|CEKA|Domluvit s kucharem
Jidlo a piti|Plynovy gril|1 ks||Jan|Jan|CEKA|Den 2 klicove
Jidlo a piti|Plynova bomba|1 az 2 ks||Jan|Jan|CEKA|Overit stav
Jidlo a piti|Drevo k ohni|dostatecna hromada||Walter|Walter|CEKA|I na druhy den
Jidlo a piti|Podpalovac a zapalovace|1 sada||Walter|Walter|CEKA|Vice kusu
Jidlo a piti|Becky piva|dle poctu||Jan|Jan|CEKA|Resit s vycepem
Jidlo a piti|Vycep a hadice|1 sada||Jan|Jan|CEKA|Bez toho becky nejedou
Jidlo a piti|Chlazeni nebo lednice na piti|1 reseni||Jan|Jan|CEKA|At je piti pitelne
Jidlo a piti|Barely nebo kanystry s vodou|2 az 4 ks||Walter|Walter|CEKA|Voda musi byt videt
Jidlo a piti|Led do piti|dle teploty||Walter|Walter|CEKA|Kupovat i prubezne
Jidlo a piti|Jednorazove talire|80 az 100 ks||Walter|Walter|CEKA|Radsi vic
Jidlo a piti|Jednorazove misky|60 az 80 ks||Walter|Walter|CEKA|Na gulas
Jidlo a piti|Jednorazove pribory|100 sad||Walter|Walter|CEKA|Mix vidlicka nuz lzice
Jidlo a piti|Kelimky na piti|150 az 200 ks||Walter|Walter|CEKA|Pivo i nealko
Jidlo a piti|Ubrousky a kuchynske role|vice baleni||Michaela|Michaela|CEKA|Mizi rychle
Jidlo a piti|Kecup, horcice, sul, pep, omacky|1 sada||Jan|Jan|CEKA|Nepodcenit
Jidlo a piti|Prkenka, noze, otviraky a kleste|1 sada||Jan|Jan|CEKA|Do gastro boxu
Atmosfera|Kytky a zelena dekorace|dle planu||Michaela|Michaela|CEKA|Jen to co zvladnete
Atmosfera|Visacky a klipy nebo snurky|50 ks||Michaela|Michaela|CEKA|Na hosty a mikrotym
Atmosfera|Plakat s hudebnim timetable|1 ks||Michaela|Michaela|CEKA|Festivalovy detail
Atmosfera|Guestbook a fixy|1 sada||Michaela|Michaela|CEKA|Volitelne ale mile
Atmosfera|Polaroid nebo fotokoutek mini|1 reseni||Michaela|Michaela|CEKA|Jen pokud to nekdo doveze
Atmosfera|Truhla nebo box na penize|1 ks||Michaela|Michaela|CEKA|Vkusny detail
Hygiena a odpad|Pytle na smes|10 az 15 ks||Walter|Walter|CEKA|Velke
Hygiena a odpad|Pytle na plast a plech|10 ks||Walter|Walter|CEKA|Oznacit
Hygiena a odpad|Rukavice na uklid|1 baleni||Walter|Walter|CEKA|Patek rano
Hygiena a odpad|Toaletni papir|vice baleni||Walter|Walter|CEKA|Nepritupna klasika
Hygiena a odpad|Mydlo a papir do koupelny|1 sada||Michaela|Michaela|CEKA|Pro chatu
Hygiena a odpad|Jar, houby a hadry|1 sada||Walter|Walter|CEKA|Uklid gastro
Bezpeci|Lekarnicka|1 ks||Michaela|Michaela|CEKA|Mit po ruce
Bezpeci|Repelent a krem na slunce|1 sada||Michaela|Michaela|CEKA|Deti a louka
Bezpeci|Baterky nebo celovky|3 az 5 ks||Walter|Walter|CEKA|Nocni orientace
Bezpeci|Deky a mikiny navic|par kusu||Michaela|Michaela|CEKA|Na vecer
"@

$contactRecords = ConvertFrom-PipeBlock -Columns $contactColumns -Block @"
Mikrotym|Jan|Mistni koordinace a komunikace s chatou||Jan|ROZPRACOVANO|Dopsat cislo
Mikrotym|Veronika|Schvaluje citlive veci a drzi vizi||Michaela|ROZPRACOVANO|V den 1 nepretizovat
Mikrotym|Walter|Svedek a krizovy koordinator||Jan|ROZPRACOVANO|Dopsat cislo
Mikrotym|Michaela|Svedkyne a atmosfera, orientace hostu||Veronika|ROZPRACOVANO|Dopsat cislo
Urad|Starostka|Oddavajici na louce||Jan|CEKA|Dopsat kontakt a potvrdit cas
Urad|Matrika|Potvrzeni casu a formalit||Jan|ROZPRACOVANO|Dopsat telefon nebo mail
Misto|Majitel chaty|Predani objektu a provozni pravidla||Jan|CEKA|Dopsat kontakt
Foto a video|Kameraman|Video zaznam a golden hour slot||Jan|ROZPRACOVANO|Dopsat kontakt
Gastro|Otec zenicha|Sele a gulas||Jan|ROZPRACOVANO|Dopsat co presne vari
Gastro|Dodavatel korytka|Korytko a pripadna dovozni logistika||Jan|CEKA|Pokud se potvrdi
Logistika|Ridic|Krize a operativni odvoz||Walter|CEKA|Vybrat cloveka
Pujcene veci|Pujcovna nebo kamarad na stany a sezeni|Party stany, lavice, sety||Walter|CEKA|Dopsat az bude jasno
"@

$packingRecords = ConvertFrom-PipeBlock -Columns $packingColumns -Block @"
Jan|Doklady a penezenska|ANO|CEKA|Mit v jedne tasce
Jan|Oblek nebo outfit na obrad|ANO|CEKA|Vyskladat den predem
Jan|Kosile nebo triko navic|NE|CEKA|Na vecerni prezleceni
Jan|Spodni pradlo a ponozky|ANO|CEKA|Na oba dny
Jan|Boty na obrad|ANO|CEKA|Rozchodit predem
Jan|Pohodne boty k ohne|NE|CEKA|Na vecer a den 2
Jan|Hygiena a deodorant|ANO|CEKA|Kartacek, pasta, sprcha
Jan|Nabijecka a powerbanka|ANO|CEKA|Kvuli hudbe i fotkam
Jan|Bunda nebo mikina|NE|CEKA|Na vecer na louce
Jan|Leciva a drobna osobni nouzovka|NE|CEKA|Co bezne beres
Jan|Taska na obradni veci|ANO|CEKA|Doklady, sliby, prstynky podle dohody
Veronika|Saty nebo hlavni outfit|ANO|CEKA|Hotove a pripravene
Veronika|Boty na obrad|ANO|CEKA|Aby nebylo prekvapeni na louce
Veronika|Pohodne druhe boty|NE|CEKA|Na vecer nebo den 2
Veronika|Spodni pradlo a nocni veci|ANO|CEKA|Na oba dny
Veronika|Kosmetika a hygiena|ANO|CEKA|Nejnutnejsi minimum po ruce
Veronika|Karticky se sliby|ANO|CEKA|Jedna kopie u Michaely
Veronika|Mikina nebo satek na vecer|NE|CEKA|Louka a noc umi prekvapit
Veronika|Nabijecka a powerbanka|ANO|CEKA|At telefon nezdechne
Veronika|Nouzovy mini kit|ANO|CEKA|Kapesniky, naplasti, siticko, rtnecka
Oba|Prstynky a pouzdro|ANO|CEKA|Rozhodnout kdo drzi
Oba|Telefonni cisla mikrotymu offline|ANO|CEKA|Na papirek i do mobilu
Oba|Plavky nebo obleceni do chaty|NE|CEKA|Jen pokud dava smysl
Oba|Pytel na spinave veci|NE|CEKA|Prakticka drobnost
"@

$checklistWidths = @(10, 12, 18, 28, 54, 16, 18, 12, 16, 36)
$scheduleWidths = @(12, 8, 8, 18, 48, 18, 30)
$budgetWidths = @(16, 28, 14, 16, 14, 16, 30)
$materialWidths = @(16, 24, 12, 22, 16, 18, 14, 28)
$contactWidths = @(14, 22, 26, 16, 16, 14, 28)
$packingWidths = @(12, 28, 10, 14, 28)

$prepRows = Build-ChecklistSheetRows -NoteText "Edituj tento list. Hlavni prehled v listu 00_Prehled se aktualizuje pres vzorce z detailnich listu." -Records $prepTasks
$day1Rows = Build-ChecklistSheetRows -NoteText "List pro operativni ukoly 24.6. Stav men hlavne tady, ne v hlavnim prehledu." -Records $day1Tasks
$day2Rows = Build-ChecklistSheetRows -NoteText "List pro operativni ukoly 25.6. Volnejsi den, ale hlidat piti, odpad a predbaleni." -Records $day2Tasks
$cleanupRows = Build-ChecklistSheetRows -NoteText "List pro patecni uklid. Vse je navazane na predani objektu do 11:00." -Records $cleanupTasks

$detailStartRow = 4
$largestChecklist = @($prepTasks.Count, $day1Tasks.Count, $day2Tasks.Count, $cleanupTasks.Count) | Measure-Object -Maximum
$detailEndRow = $detailStartRow + $largestChecklist.Maximum + 120
$detailValidations = @(
    [pscustomobject]@{ Sqref = "H4:H$detailEndRow"; Formula = "KRITICKA,VYSOKA,STREDNI,NIZKA" },
    [pscustomobject]@{ Sqref = "I4:I$detailEndRow"; Formula = "NENI HOTOVO,ROZPRACOVANO,CEKA,HOTOVO" }
)

$masterRows = @()
$masterRows += ,(@((New-TextCell -Value "Pouziti" -Style 2), (New-TextCell -Value "Tento list je propojeny z detailnich listu 01 az 04. Stav a prioritu men hlavne tam." -Style 3)))
$masterRows += ,@()
$masterRows += ,(@(
    (New-TextCell -Value "Faze" -Style 2),
    (New-TextCell -Value "Celkem" -Style 2),
    (New-TextCell -Value "Hotovo" -Style 2),
    (New-TextCell -Value "Rozpracovano" -Style 2),
    (New-TextCell -Value "Ceka" -Style 2),
    (New-TextCell -Value "Neni hotovo" -Style 2)
))

$masterDataStartRow = 11
$masterDataEndRow = $masterDataStartRow + $prepTasks.Count + $day1Tasks.Count + $day2Tasks.Count + $cleanupTasks.Count - 1
foreach ($phase in @("Priprava", "Den 1", "Den 2", "Uklid")) {
    $masterRows += ,(@(
        (New-TextCell -Value $phase -Style 1),
        (New-FormulaCell -Formula ('COUNTIF($B${0}:$B${1},"{2}")' -f $masterDataStartRow, $masterDataEndRow, $phase) -Style 1),
        (New-FormulaCell -Formula ('COUNTIFS($B${0}:$B${1},"{2}",$I${0}:$I${1},"HOTOVO")' -f $masterDataStartRow, $masterDataEndRow, $phase) -Style 1),
        (New-FormulaCell -Formula ('COUNTIFS($B${0}:$B${1},"{2}",$I${0}:$I${1},"ROZPRACOVANO")' -f $masterDataStartRow, $masterDataEndRow, $phase) -Style 1),
        (New-FormulaCell -Formula ('COUNTIFS($B${0}:$B${1},"{2}",$I${0}:$I${1},"CEKA")' -f $masterDataStartRow, $masterDataEndRow, $phase) -Style 1),
        (New-FormulaCell -Formula ('COUNTIFS($B${0}:$B${1},"{2}",$I${0}:$I${1},"NENI HOTOVO")' -f $masterDataStartRow, $masterDataEndRow, $phase) -Style 1)
    ))
}
$masterRows += ,(@(
    (New-TextCell -Value "CELKEM" -Style 2),
    (New-FormulaCell -Formula "SUM(B4:B7)" -Style 1),
    (New-FormulaCell -Formula "SUM(C4:C7)" -Style 1),
    (New-FormulaCell -Formula "SUM(D4:D7)" -Style 1),
    (New-FormulaCell -Formula "SUM(E4:E7)" -Style 1),
    (New-FormulaCell -Formula "SUM(F4:F7)" -Style 1)
))
$masterRows += ,@()
$masterRows += ,(@($taskHeaders | ForEach-Object { New-TextCell -Value $_ -Style 2 }))

$sourceSheetMap = [ordered]@{
    "01_Priprava" = $prepTasks
    "02_Den1"     = $day1Tasks
    "03_Den2"     = $day2Tasks
    "04_Uklid"    = $cleanupTasks
}
foreach ($sheetName in $sourceSheetMap.Keys) {
    $records = $sourceSheetMap[$sheetName]
    for ($rowIndex = 0; $rowIndex -lt $records.Count; $rowIndex++) {
        $sourceRow = $detailStartRow + $rowIndex
        $formulaRow = @()
        for ($columnIndex = 1; $columnIndex -le $taskColumns.Count; $columnIndex++) {
            $columnName = Get-ColumnName $columnIndex
            $formulaRow += New-FormulaCell -Formula "'$sheetName'!$columnName$sourceRow" -Style 1
        }
        $masterRows += ,$formulaRow
    }
}

$scheduleRows = @()
$scheduleRows += ,(@((New-TextCell -Value "Poznamka" -Style 2), (New-TextCell -Value "Toto je navrzeny hodinovy plan. Upravujte podle reality, ale drzte orientacni kostru." -Style 3)))
$scheduleRows += ,@()
$scheduleRows += Convert-TableToRows -Headers $scheduleHeaders -Keys $scheduleColumns -Records $scheduleRecords

$budgetRows = @()
$budgetRows += ,(@((New-TextCell -Value "Poznamka" -Style 2), (New-TextCell -Value "Rozpoctovy strop je 125000 CZK. Odhady jsou navrh, skutecnost doplnujte prubezne." -Style 3)))
$budgetRows += ,@()
$budgetRows += ,(@((New-TextCell -Value "Rozpoctovy strop" -Style 2), (New-NumberCell -Value 125000 -Style 4)))
$budgetRows += ,(@((New-TextCell -Value "Aktualni odhad" -Style 2), (New-FormulaCell -Formula "SUM(C9:C22)" -Style 4)))
$budgetRows += ,(@((New-TextCell -Value "Aktualni skutecnost" -Style 2), (New-FormulaCell -Formula "SUM(D9:D22)" -Style 4)))
$budgetRows += ,(@((New-TextCell -Value "Rezerva do stropu" -Style 2), (New-FormulaCell -Formula "B3-B4" -Style 4)))
$budgetRows += ,@()
$budgetRows += Convert-TableToRows -Headers $budgetHeaders -Keys $budgetColumns -Records $budgetRecords -CurrencyColumns @(2, 3)

$materialRows = @()
$materialRows += ,(@((New-TextCell -Value "Poznamka" -Style 2), (New-TextCell -Value "Tady drzte fyzicke veci, pujcene vybaveni a kdo je vraci. Je to zaklad pro patek." -Style 3)))
$materialRows += ,@()
$materialRows += Convert-TableToRows -Headers $materialHeaders -Keys $materialColumns -Records $materialRecords

$contactRows = @()
$contactRows += ,(@((New-TextCell -Value "Poznamka" -Style 2), (New-TextCell -Value "Dopln telefonni cisla, mail nebo poznamky. Jeden list pro cely mikrotym." -Style 3)))
$contactRows += ,@()
$contactRows += Convert-TableToRows -Headers $contactHeaders -Keys $contactColumns -Records $contactRecords

$packingRows = @()
$packingRows += ,(@((New-TextCell -Value "Poznamka" -Style 2), (New-TextCell -Value "Osobni baleni je oddelene, aby se neztratily male ale kriticke veci." -Style 3)))
$packingRows += ,@()
$packingRows += Convert-TableToRows -Headers $packingHeaders -Keys $packingColumns -Records $packingRecords

$sheets = @(
    [pscustomobject]@{ Name = "00_Prehled"; Path = "xl/worksheets/sheet1.xml"; Rows = $masterRows; Widths = $checklistWidths; FreezeRows = 10; AutoFilterRef = "A10:J$($masterRows.Count)"; DataValidations = @() },
    [pscustomobject]@{ Name = "01_Priprava"; Path = "xl/worksheets/sheet2.xml"; Rows = $prepRows; Widths = $checklistWidths; FreezeRows = 3; AutoFilterRef = "A3:J$($prepRows.Count)"; DataValidations = $detailValidations },
    [pscustomobject]@{ Name = "02_Den1"; Path = "xl/worksheets/sheet3.xml"; Rows = $day1Rows; Widths = $checklistWidths; FreezeRows = 3; AutoFilterRef = "A3:J$($day1Rows.Count)"; DataValidations = $detailValidations },
    [pscustomobject]@{ Name = "03_Den2"; Path = "xl/worksheets/sheet4.xml"; Rows = $day2Rows; Widths = $checklistWidths; FreezeRows = 3; AutoFilterRef = "A3:J$($day2Rows.Count)"; DataValidations = $detailValidations },
    [pscustomobject]@{ Name = "04_Uklid"; Path = "xl/worksheets/sheet5.xml"; Rows = $cleanupRows; Widths = $checklistWidths; FreezeRows = 3; AutoFilterRef = "A3:J$($cleanupRows.Count)"; DataValidations = $detailValidations },
    [pscustomobject]@{ Name = "05_Harmonogram"; Path = "xl/worksheets/sheet6.xml"; Rows = $scheduleRows; Widths = $scheduleWidths; FreezeRows = 3; AutoFilterRef = "A3:G$($scheduleRows.Count)"; DataValidations = @() },
    [pscustomobject]@{ Name = "06_Rozpocet"; Path = "xl/worksheets/sheet7.xml"; Rows = $budgetRows; Widths = $budgetWidths; FreezeRows = 7; AutoFilterRef = "A8:G$($budgetRows.Count)"; DataValidations = @() },
    [pscustomobject]@{ Name = "07_Material"; Path = "xl/worksheets/sheet8.xml"; Rows = $materialRows; Widths = $materialWidths; FreezeRows = 3; AutoFilterRef = "A3:H$($materialRows.Count)"; DataValidations = @() },
    [pscustomobject]@{ Name = "08_Kontakty"; Path = "xl/worksheets/sheet9.xml"; Rows = $contactRows; Widths = $contactWidths; FreezeRows = 3; AutoFilterRef = "A3:G$($contactRows.Count)"; DataValidations = @() },
    [pscustomobject]@{ Name = "09_Baleni"; Path = "xl/worksheets/sheet10.xml"; Rows = $packingRows; Widths = $packingWidths; FreezeRows = 3; AutoFilterRef = "A3:E$($packingRows.Count)"; DataValidations = @() }
)

$sheetFiles = @{}
foreach ($sheet in $sheets) {
    $sheetFiles[$sheet.Path] = New-WorksheetXml -Rows $sheet.Rows -ColumnWidths $sheet.Widths -FreezeRows $sheet.FreezeRows -AutoFilterRef $sheet.AutoFilterRef -DataValidations $sheet.DataValidations
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("svatba_xlsx_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "docProps") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\worksheets") | Out-Null

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Default Extension="xml" ContentType="application/xml" />
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet8.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet9.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/worksheets/sheet10.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
</Types>
"@

$rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
</Relationships>
"@

$sheetNodes = for ($i = 0; $i -lt $sheets.Count; $i++) {
    "<sheet name=""$($sheets[$i].Name)"" sheetId=""$($i + 1)"" r:id=""rId$($i + 1)"" />"
}
$workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000" />
  </bookViews>
  <sheets>
    $($sheetNodes -join "`n    ")
  </sheets>
  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" />
</workbook>
"@

$workbookRelsNodes = for ($i = 0; $i -lt $sheets.Count; $i++) {
    "<Relationship Id=""rId$($i + 1)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$($i + 1).xml"" />"
}
$workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $($workbookRelsNodes -join "`n  ")
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
</Relationships>
"@

$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0 &quot;CZK&quot;" />
  </numFmts>
  <fonts count="2">
    <font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /></font>
    <font><b /><sz val="11" /><color rgb="FFFFFFFF" /><name val="Calibri" /><family val="2" /></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none" /></fill>
    <fill><patternFill patternType="gray125" /></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78" /><bgColor indexed="64" /></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2" /><bgColor indexed="64" /></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left /><right /><top /><bottom /><diagonal /></border>
    <border><left style="thin"><color auto="1" /></left><right style="thin"><color auto="1" /></right><top style="thin"><color auto="1" /></top><bottom style="thin"><color auto="1" /></bottom><diagonal /></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1" /></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" /></xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1" /></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1" /></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0" /></cellStyles>
</styleSheet>
"@

$titlesOfParts = @()
for ($i = 0; $i -lt $sheets.Count; $i++) {
    $titlesOfParts += "<vt:lpstr>$($sheets[$i].Name)</vt:lpstr>"
}
$titlesOfPartsXml = $titlesOfParts -join ""

$appXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>$($sheets.Count)</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="$($sheets.Count)" baseType="lpstr">$titlesOfPartsXml</vt:vector></TitlesOfParts>
  <AppVersion>16.0000</AppVersion>
</Properties>
"@

$createdAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$createdAt</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$createdAt</dcterms:modified>
  <dc:title>Svatba Ostravice Checklist</dc:title>
</cp:coreProperties>
"@

[System.IO.File]::WriteAllText((Join-Path $tempRoot "[Content_Types].xml"), $contentTypes, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tempRoot "_rels\.rels"), $rootRels, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps\app.xml"), $appXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps\core.xml"), $coreXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\workbook.xml"), $workbookXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\_rels\workbook.xml.rels"), $workbookRelsXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\styles.xml"), $stylesXml, [System.Text.UTF8Encoding]::new($false))
foreach ($relativePath in $sheetFiles.Keys) {
    [System.IO.File]::WriteAllText((Join-Path $tempRoot $relativePath), $sheetFiles[$relativePath], [System.Text.UTF8Encoding]::new($false))
}

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$fileStream = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::CreateNew)
$archive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
foreach ($file in Get-ChildItem -Path $tempRoot -Recurse -File) {
    $relativePath = $file.FullName.Substring($tempRoot.Length + 1).Replace("\", "/")
    $entry = $archive.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Optimal)
    $entryStream = $entry.Open()
    $sourceStream = [System.IO.File]::OpenRead($file.FullName)
    $sourceStream.CopyTo($entryStream)
    $sourceStream.Dispose()
    $entryStream.Dispose()
}
$archive.Dispose()
$fileStream.Dispose()
Remove-Item -LiteralPath $tempRoot -Recurse -Force
Write-Output "Workbook created: $OutputPath"
