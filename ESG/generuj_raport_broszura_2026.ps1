#Requires -Version 5.1
# Audyt broszury handlowej ESG Institute 2026 zgodnie z INSTRUKCJA_GENEROWANIA_RAPORTOW.md
$ErrorActionPreference = 'Stop'
$finalPath = Join-Path $PSScriptRoot 'Raport_Audytu_Greenwashingu_ESG_Institute_Broszura_2026.docx'
$outPath = Join-Path $env:TEMP 'Raport_Audytu_Greenwashingu_ESG_Institute_Broszura_2026.docx'

function Rgb([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }

$NAVY    = Rgb 8 24 31
$ACCENT  = Rgb 42 138 176
$DANGER  = Rgb 190 18 60
$WARNING = Rgb 180 83 9
$SUCCESS = Rgb 4 120 87
$DARK    = Rgb 22 36 44
$MUTED   = Rgb 90 107 119
$WHITE   = Rgb 255 255 255
$REDBG   = Rgb 255 241 242

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $word.Options.CheckSpellingAsYouType = $false
    $word.Options.CheckGrammarAsYouType = $false
    $doc = $word.Documents.Add()
    $ps = $doc.PageSetup
    $ps.PageWidth = $word.CentimetersToPoints(21.0)
    $ps.PageHeight = $word.CentimetersToPoints(29.7)
    $ps.TopMargin = $word.CentimetersToPoints(2.5)
    $ps.BottomMargin = $word.CentimetersToPoints(2.0)
    $ps.LeftMargin = $word.CentimetersToPoints(1.7)
    $ps.RightMargin = $word.CentimetersToPoints(1.7)
    $ps.HeaderDistance = $word.CentimetersToPoints(0.45)
    $ps.FooterDistance = $word.CentimetersToPoints(0.4)
    $contentW = $word.CentimetersToPoints(17.6)
    $sel = $word.Selection

    function Add-P([string]$text, [double]$size, [bool]$bold, [bool]$italic, [int]$color, [double]$before, [double]$after) {
        $sel.Font.Name = 'Arial'
        $sel.Font.Size = $size
        $sel.Font.Bold = $(if ($bold) { 1 } else { 0 })
        $sel.Font.Italic = $(if ($italic) { 1 } else { 0 })
        $sel.Font.Color = $color
        $sel.ParagraphFormat.SpaceBefore = $before
        $sel.ParagraphFormat.SpaceAfter = $after
        $sel.ParagraphFormat.LineSpacingRule = 0
        $sel.ParagraphFormat.Alignment = 0
        $sel.TypeText($text)
        $sel.TypeParagraph()
    }

    function Add-H([string]$text) {
        $sel.Font.Name = 'Arial'
        $sel.Font.Size = 14
        $sel.Font.Bold = 1
        $sel.Font.Italic = 0
        $sel.Font.Color = $NAVY
        $sel.ParagraphFormat.SpaceBefore = 16
        $sel.ParagraphFormat.SpaceAfter = 6
        $sel.ParagraphFormat.LineSpacingRule = 0
        $sel.ParagraphFormat.KeepWithNext = $true
        $sel.ParagraphFormat.Alignment = 0
        $sel.TypeText($text)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.KeepWithNext = $false
    }

    function Shade-Cell($cell, [int]$color) { $cell.Shading.BackgroundPatternColor = $color }

    function Pad-Cell($cell, [double]$top, [double]$bottom, [double]$left, [double]$right) {
        $cell.TopPadding = $top
        $cell.BottomPadding = $bottom
        $cell.LeftPadding = $left
        $cell.RightPadding = $right
        $cell.VerticalAlignment = 1
    }

    function Write-Cell($cell, [string]$text, [double]$size, [bool]$bold, [int]$color) {
        $cell.Range.Text = $text
        $r = $cell.Range
        $r.Font.Name = 'Arial'
        $r.Font.Size = $size
        $r.Font.Bold = $(if ($bold) { 1 } else { 0 })
        $r.Font.Italic = 0
        $r.Font.Color = $color
        $r.ParagraphFormat.SpaceBefore = 0
        $r.ParagraphFormat.SpaceAfter = 0
        $r.ParagraphFormat.LineSpacingRule = 0
    }

    function Move-AfterTable($table) {
        $end = $table.Range.End
        $sel.SetRange($end, $end)
        $sel.TypeParagraph()
    }

    function Add-Callout([string]$title, [string[]]$lines, [int]$bg, [int]$titleColor) {
        $insertAt = $doc.Content.End - 1
        if ($insertAt -lt 0) { $insertAt = 0 }
        $range = $doc.Range($insertAt, $insertAt)
        $t = $doc.Tables.Add($range, 1, 1)
        $t.PreferredWidthType = 3
        $t.PreferredWidth = $contentW
        $t.AllowAutoFit = $false
        $c = $t.Cell(1, 1)
        Shade-Cell $c $bg
        Pad-Cell $c 8 8 10 10
        $body = ([char]0x25A0).ToString() + ' ' + $title
        foreach ($line in $lines) { $body += "`r" + $line }
        $c.Range.Text = $body
        $c.Range.Font.Name = 'Arial'
        $c.Range.Font.Size = 10
        $c.Range.Font.Bold = 0
        $c.Range.Font.Italic = 0
        $c.Range.Font.Color = $DARK
        $c.Range.ParagraphFormat.SpaceBefore = 2
        $c.Range.ParagraphFormat.SpaceAfter = 2
        $c.Range.Paragraphs.Item(1).Range.Font.Bold = 1
        $c.Range.Paragraphs.Item(1).Range.Font.Color = $titleColor
        $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.SpaceBefore = 0
        $sel.ParagraphFormat.SpaceAfter = 6
    }

    $header = $doc.Sections.Item(1).Headers.Item(1)
    $header.Range.Text = ''
    $ht = $header.Range.Tables.Add($header.Range, 1, 2)
    $ht.PreferredWidthType = 3
    $ht.PreferredWidth = $contentW
    $ht.AllowAutoFit = $false
    $ht.Columns.Item(1).Width = $word.CentimetersToPoints(8.2)
    $ht.Columns.Item(2).Width = $word.CentimetersToPoints(9.4)
    foreach ($i in 1, 2) {
        $hc = $ht.Cell(1, $i)
        Shade-Cell $hc $NAVY
        Pad-Cell $hc 4 4 8 8
        $hc.Borders.Enable = $false
    }
    $logo = Join-Path $PSScriptRoot 'esg_logo_white.png'
    $logoOk = $false
    if (Test-Path -LiteralPath $logo) {
        try {
            $ht.Cell(1, 1).Range.Text = ''
            $null = $ht.Cell(1, 1).Range.InlineShapes.AddPicture($logo)
            $logoOk = $true
        } catch {
            $logoOk = $false
        }
    }
    if (-not $logoOk) {
        Write-Cell $ht.Cell(1, 1) 'ESG INSTITUTE' 11 $true $WHITE
    }
    $right = $ht.Cell(1, 2)
    $right.Range.Text = "AUDYT RYZYK GREENWASHINGU`rESG Institute"
    $right.Range.Font.Name = 'Arial'
    $right.Range.Font.Color = $WHITE
    $right.Range.ParagraphFormat.Alignment = 2
    $right.Range.Paragraphs.Item(1).Range.Font.Size = 8
    $right.Range.Paragraphs.Item(1).Range.Font.Bold = 1
    $right.Range.Paragraphs.Item(2).Range.Font.Size = 8
    $right.Range.Paragraphs.Item(2).Range.Font.Bold = 0
    $right.Range.Paragraphs.Item(2).Range.Font.Color = (Rgb 129 192 223)

    $footer = $doc.Sections.Item(1).Footers.Item(1)
    $footer.Range.Text = 'ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Broszura handlowa 2026  ·  1 listopada 2026 r.'
    $footer.Range.Font.Name = 'Arial'
    $footer.Range.Font.Size = 8
    $footer.Range.Font.Color = $MUTED

    $word.ActiveWindow.ActivePane.View.SeekView = 0
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)

    Add-P 'ESG INSTITUTE · RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO' 9 $true $false $ACCENT 0 2
    Add-P 'Audyt Ryzyk Greenwashingu: ESG INSTITUTE' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja broszury handlowej 2026, plik PDF, 23 strony, data utworzenia pliku 31 października 2026 r. Adres na okładce: www.esginstitute.eu. Ocena dotyczy wyłącznie tego pliku. Data audytu: 1 listopada 2026 r. Ocena wobec ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, ustawy o zwalczaniu nieuczciwej konkurencji, ustawy o ochronie konkurencji i konsumentów, dyrektywy 2005/29/WE, dyrektywy (UE) 2024/825, rozporządzenia (UE) 2019/2088, rozporządzenia (UE) 2024/3110 oraz norm ISO 14040, ISO 14044, ISO 14067 i EN 15804.' 10.5 $false $true $MUTED 0 10

    $meta = @(
        @('Badany podmiot', 'ESG Institute, Rondo ONZ 1, 00-124 Warszawa, tel. +48 722 230 101, info@esginstitute.eu, www.esginstitute.eu. Prezeska i współzałożycielka: Agnieszka Orłowska, orlowska@esginstitute.eu. Forma prawna i numer KRS nie są podane w broszurze.'),
        @('Przedmiot audytu', 'Broszura handlowa 2026, 23 strony, plik utworzony 31 października 2026 r. Okładka, spis, wstęp, opisy usług, słownik i kontakt. Witryna esginstitute.eu nie wchodzi do ustaleń i nie była porównywana z tym plikiem.'),
        @('Sektor działalności', 'Doradztwo dla firm: raportowanie zrównoważonego rozwoju, ślad gazów cieplarnianych, cykl życia i deklaracja środowiskowa wyrobu, przegląd twierdzeń, ciągłość działania, luka płacowa, platforma Green0meter, szkolenia, społeczność Collaboration Hub i przygotowanie do EcoVadis. Odbiorca: organizacja. Collaboration Hub sprzedaje też członkostwo osoby indywidualnej za 499 zł netto za rok.'),
        @('Reżim prawny', 'UPNPR art. 5 ust. 2 pkt 1 i 2, ust. 3 pkt 2 oraz art. 6. UZNK art. 14 ust. 2 pkt 2 i 4, art. 16 ust. 1 pkt 2 i ust. 2 oraz art. 18. Ustawa o ochronie konkurencji i konsumentów art. 24 i 106. Dyrektywa 2005/29/WE. Dyrektywa (UE) 2024/825, art. 4 i załącznik I pkt 4a i 4b. Rozporządzenie (UE) 2019/2088. Rozporządzenie (UE) 2024/3110. ISO 14040, ISO 14044, ISO 14067, EN 15804. Wyroki C-611/14 i C-281/12 w granicach wskazanych przy przypadkach.'),
        @('Data audytu', '1 listopada 2026 r.'),
        @('Werdykt operacyjny', 'NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)')
    )

    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $mt = $doc.Tables.Add($range, $meta.Count, 2)
    $mt.PreferredWidthType = 3
    $mt.PreferredWidth = $contentW
    $mt.AllowAutoFit = $false
    $mt.Columns.Item(1).Width = $word.CentimetersToPoints(4.2)
    $mt.Columns.Item(2).Width = $word.CentimetersToPoints(13.4)
    for ($i = 0; $i -lt $meta.Count; $i++) {
        $ck = $mt.Cell($i + 1, 1)
        $cv = $mt.Cell($i + 1, 2)
        Shade-Cell $ck (Rgb 248 250 252)
        $last = ($i -eq ($meta.Count - 1))
        if ($last) { Shade-Cell $cv $REDBG } else { Shade-Cell $cv $WHITE }
        Pad-Cell $ck 3 3 6 6
        Pad-Cell $cv 3 3 6 6
        Write-Cell $ck $meta[$i][0] 8.5 $true $DARK
        Write-Cell $cv $meta[$i][1] 8.5 $last $(if ($last) { $DANGER } else { $DARK })
    }
    Move-AfterTable $mt
    Add-P '' 6 $false $false $DARK 0 6

    Add-Callout 'WERDYKT: NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)' @(
        'Okładka „Tworzyć w pełni zrównoważoną przyszłość, już dziś” jest ogólnym hasłem środowiskowym o całości. Nie ma roku, wskaźnika ani wyrobu i nie ma w tym miejscu EU Ecolabel ani urzędowego znaku ISO 14024 typu I. To punkt 4a i 4b załącznika I dyrektywy (UE) 2024/825.',
        'Ta sama broszura na stronie audytu przesuwa te zakazy na 27 września 2027 r. i nazywa ustawą tekst, który jest drukiem Senatu 825. Ślad produktu, wydarzenia i usługi jest przypisany do metody organizacji. Cztery wskaźniki klimatu są opisane jako już policzone do przyszłej deklaracji właściwości użytkowych.',
        'Na stronie 07 zostaje zdanie, że samo wyliczenie nie jest deklaracją neutralności klimatycznej. Plik wraca do wysyłki po zdjęciu okładki i po korekcie zdań z rozdziału 3.'
    ) $REDBG $DANGER

    Add-H '1. Analiza Wstępna'
    Add-P 'ESG Institute sprzedaje w tej broszurze doradztwo dla organizacji: raport według dyrektywy CSRD i według standardu VSME, obliczenie śladu, analizę cyklu życia, dokumentację deklaracji środowiskowej, przegląd twierdzeń, plan ciągłości, lukę płacową, platformę Green0meter, szkolenia, społeczność Collaboration Hub i przygotowanie do EcoVadis. Nadawcą listu jest Agnieszka Orłowska, prezeska. Odbiorcą jest firma. Osoba indywidualna może kupić członkostwo Hubu za 499 zł netto za rok, więc broszura nie jest wyłącznie pismem między przedsiębiorcami.' 10.5 $false $false $DARK 0 8
    Add-P 'Narracja otwiera się hasłem o w pełni zrównoważonej przyszłości, tworzonej już dziś. Dalej usługi są opisane nazwami metod: GHG Protocol, ISO 14040, ISO 14044, ISO 14025, EN 15804, VSME, EcoVadis. Kilka zdań jest zawężonych i zostaje. Samo wyliczenie śladu nie jest deklaracją neutralności. Deklarację publikuje operator programu, a opłatę za wpis ponosi producent. Przy finansowaniu jest „część banków i funduszy”, nie obietnica warunków kredytu. Rozdział o dyrektywie CSRD oddziela ustawę od projektu. Rozdział o luce płacowej oddziela to, co broszura nazywa już obowiązującym, od projektu ustawy UC127.' 10.5 $false $false $DARK 0 8
    Add-P 'Naruszenia siedzą w haśle, w dacie zakazu i w pomieszaniu obiektu z metodą. Okładka jest ogólnym hasłem o całości. Strona usługi, która takie hasła ocenia, podaje ich start na 27 września 2027 r. Jedna metoda organizacji obejmuje produkt, wydarzenie i usługę. Pierwsze zdanie analizy cyklu życia sięga końca życia, a krok pierwszy dopuszcza bramę zakładu. Ramka przy deklaracji mówi, że wskaźniki do przyszłej deklaracji właściwości użytkowych są już policzone. Platforma stawia raport według SFDR obok raportu dla firmy. Pierwsze zdanie o EcoVadis robi z medalu warunek kluczowych odbiorców, a ramka na tej samej stronie mówi o części odbiorców.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'
    $findings = @(
        @('01','„W pełni zrównoważoną przyszłość, już dziś”','ECGT zał. I pkt 4a i 4b · UPNPR art. 5','WYSOKIE','Ogólne hasło o całości, bez znaku typu I'),
        @('02','Zakazy od 27 września 2027 r., ustawa poz. 1353','dyrektywa 2024/825 art. 4 · UPNPR art. 5 i 6','WYSOKIE','Data zakazu i numer druku podane jako ustawa'),
        @('03','GHG Protocol dla organizacji, produktu, wydarzenia i usługi','UPNPR art. 5 i 6 · ISO 14067','WYSOKIE','Metoda organizacji na czterech obiektach'),
        @('04','Analiza „od surowca do końca życia” i „cały zestaw”','UPNPR art. 5 i 6 · ISO 14040/14044','WYSOKIE','Dwie granice i pełny zestaw skutków'),
        @('05','Cztery wskaźniki klimatu „już policzone” do deklaracji','UPNPR art. 5 i 6 · rozp. 2024/3110','WYSOKIE','Podział z EPD jako treść przyszłej deklaracji'),
        @('06','Raport według SFDR obok CSRD i VSME','UPNPR art. 5 i 6 · rozp. 2019/2088','WYSOKIE','Ujawnienie finansowe jako raport firmy'),
        @('07','Medal EcoVadis warunkiem kluczowych odbiorców','UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14','WYSOKIE','Medal i próg wyniku jako cecha przygotowania'),
        @('08','Spadek emisji obniża rachunki i koszt wytworzenia','UPNPR art. 5 ust. 3 pkt 2','ŚREDNIE','Korzyść finansowa bez pozycji i bez roku'),
        @('09','Kluczowi odbiorcy wymagają śladu, brak ogranicza przetarg','UPNPR art. 5 ust. 3 pkt 2 · art. 6','ŚREDNIE','Wymóg całej grupy odbiorców bez postępowania'),
        @('10','Liczby EPD wchodzą do BREEAM i LEED i służą porównaniu','UPNPR art. 5 i 6 · EN 15804','ŚREDNIE','Deklaracja wyrobu jako ocena budynku'),
        @('11','„Jeden komplet” i ważność deklaracji pięć lat','UPNPR art. 5 ust. 2 pkt 2','ŚREDNIE','Publikacja i okres ważności bez programu'),
        @('12','Procedura ciągłości spełnia NIS2, DORA i normy','UPNPR art. 5 i 6 · UZNK art. 14','ŚREDNIE','Reżim bez podmiotu objętego obowiązkiem'),
        @('13','Udział w Hubie wyróżnia organizację na rynku','UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 16','ŚREDNIE','Wyróżnienie członkostwa bez kryterium')
    )

    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $st = $doc.Tables.Add($range, $findings.Count + 1, 5)
    $st.PreferredWidthType = 3
    $st.PreferredWidth = $contentW
    $st.AllowAutoFit = $false
    $widths = @(1.1, 5.6, 4.4, 2.2, 4.3)
    for ($c = 0; $c -lt 5; $c++) { $st.Columns.Item($c + 1).Width = $word.CentimetersToPoints($widths[$c]) }
    $headers = @('ID','Analizowane twierdzenie','Podstawa prawna','Ryzyko','Istota naruszenia')
    for ($c = 0; $c -lt 5; $c++) {
        $cell = $st.Cell(1, $c + 1)
        Shade-Cell $cell $NAVY
        Pad-Cell $cell 3 3 4 4
        Write-Cell $cell $headers[$c] 8 $true $WHITE
    }
    $st.Rows.Item(1).HeadingFormat = $true
    for ($r = 0; $r -lt $findings.Count; $r++) {
        $bg = $(if ($r % 2 -eq 0) { $WHITE } else { (Rgb 248 250 252) })
        for ($c = 0; $c -lt 5; $c++) {
            $cell = $st.Cell($r + 2, $c + 1)
            Shade-Cell $cell $bg
            Pad-Cell $cell 3 3 4 4
            $val = $findings[$r][$c]
            $bold = ($c -eq 0 -or $c -eq 3)
            $colr = $DARK
            if ($c -eq 0) { $colr = $NAVY }
            if ($c -eq 3) {
                if ($val -eq 'ŚREDNIE') { $colr = $WARNING } else { $colr = $DANGER }
            }
            Write-Cell $cell $val 7.5 $bold $colr
        }
    }
    Move-AfterTable $st

    Add-H '3. Szczegółowa analiza przypadków'

    function Add-Case([string]$title, [string]$basis, [string[]]$quotes, [string]$assessment, [string]$remediation, [string]$claimTitle, [string]$claim) {
        $sel.Font.Name = 'Arial'
        $sel.Font.Size = 12
        $sel.Font.Bold = 1
        $sel.Font.Italic = 0
        $sel.Font.Color = $NAVY
        $sel.ParagraphFormat.SpaceBefore = 12
        $sel.ParagraphFormat.SpaceAfter = 2
        $sel.ParagraphFormat.KeepWithNext = $true
        $sel.TypeText($title)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.KeepWithNext = $false
        Add-P $basis 10 $false $false $DARK 0 4
        Add-P 'Cytat źródłowy:' 10.5 $true $false $DARK 4 2
        foreach ($q in $quotes) { Add-P $q 10.5 $false $true $DARK 0 4 }
        Add-P 'Ocena prawna:' 10.5 $true $false $DARK 4 2
        Add-P $assessment 10.5 $false $false $DARK 0 6
        Add-P 'Kierunek naprawy:' 10.5 $true $false $DARK 2 2
        Add-P $remediation 10.5 $false $false $DARK 0 6
        $bg = (Rgb 240 253 250)
        $tc = $SUCCESS
        if ($claimTitle -like '*NIE DO URATOWANIA*') { $bg = (Rgb 255 241 242); $tc = $DANGER }
        Add-Callout $claimTitle @($claim) $bg $tc
    }

    Add-Case 'PRZYPADEK 01: Ogólne hasło o w pełni zrównoważonej przyszłości' `
        'Podstawa prawna: dyrektywa (UE) 2024/825, załącznik I pkt 4a i 4b · UPNPR art. 5 · UZNK art. 16 ust. 1 pkt 2 · C-611/14 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Okładka: „Tworzyć w pełni zrównoważoną przyszłość, już dziś.”',
            'Spis treści, strona 02: „Oferujemy usługi odpowiadające potrzebom organizacji, które rozwijają zrównoważony model działania.”'
        ) `
        @'
Z okładki czytelnik wynosi, że ESG Institute tworzy w pełni zrównoważoną przyszłość i że jest to stan na dziś. Hasło nie wskazuje roku, wskaźnika, wyrobu ani usługi. Spis treści używa słowa „zrównoważony” przy modelu działania organizacji, też bez tematu i bez roku. Takie ogólne hasło środowiskowe wolno zostawić tylko wtedy, gdy w tym samym miejscu jest EU Ecolabel albo urzędowy znak ISO 14024 typu I. Na okładce i przy spisie tego znaku nie ma. „W pełni” obejmuje przyszłość jako całość, a dalsze strony opisują osobne usługi: raport, ślad, deklarację wyrobu. Słownik i zdanie o neutralności stoją wiele stron dalej. Hasło, które samo tworzy to wrażenie, nie jest poprawiane odesłaniem (C-611/14). Punkt 4a i tak wymaga znaku w tym samym miejscu.
'@ `
        'Z okładki zdjąć całe zdanie „Tworzyć w pełni zrównoważoną przyszłość, już dziś”. Ze spisu zdjąć „zrównoważony model działania”. Nie zastępować go słowem „eko”, „zielony”, „przyjazny środowisku” ani „zrównoważony”. Na okładce nie stawiać liścia, kropli ani innego znaku, który czyta się jak certyfikat. Słownik nie poprawia okładki.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Broszura usług 2026: raportowanie, inwentaryzacja gazów cieplarnianych organizacji, analiza cyklu życia wyrobu i deklaracja środowiskowa wyrobu. Zakres ustalamy po identyfikacji progu, który firmę dotyczy.”'

    Add-Case 'PRZYPADEK 02: Zakazy środowiskowe od 27 września 2027 r. i numer ustawy' `
        'Podstawa prawna: dyrektywa (UE) 2024/825 art. 4 · UPNPR art. 5 ust. 2 pkt 1 i art. 6 · UZNK art. 14 ust. 2 pkt 4 | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona 11, ramka „Od 27 września 2027 r.”: „Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) uznaje za zakazane w każdych okolicznościach ogólne hasło środowiskowe, twierdzenie o całym produkcie oparte na jednym aspekcie, neutralność klimatu z kompensacji, własny znak zrównoważoności oraz wymóg prawa pokazany jako cecha oferty.”',
            'Ta sama strona, ramka „Obowiązuje już dziś”: „Wprowadzenie w błąd i pominięcie istotnej informacji obowiązują na podstawie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym.”',
            'Słownik, strona 22: „Nowe punkty listy praktyk zawsze zakazanych wchodzą w Polsce 27 września 2027 r.”'
        ) `
        @'
Czytelnik ramki wynosi, że pięć nazwanych zakazów zaczyna działać 27 września 2027 r. i że podstawą jest ustawa ogłoszona pod pozycją 1353. Obok jest druga ramka: wprowadzenie w błąd i pominięcie obowiązują już dziś. Ta druga ramka nie wymienia ogólnego hasła, hasła o całym produkcie, kompensacji ani własnego znaku. Słownik powtarza tylko datę 2027 r. Dyrektywa (UE) 2024/825 stosuje się od 27 września 2026 r. Te same wzorce już dziś ocenia się jako wprowadzenie w błąd albo pominięcie. Tekst przyjęty przez Sejm 4 września 2026 r. jest drukiem Senatu 825. Do ogłoszenia nie jest ustawą, a data 27 września 2027 r. stoi w tym druku. Okładka tej broszury jest właśnie ogólnym hasłem. Czytelnik usługi audytu może zostawić takie hasło do września 2027 r.
'@ `
        'Z ramki i ze słownika usunąć „Dz.U. 2026 poz. 1353”, słowo „ustawa” przy tekście z 4 września 2026 r. oraz nagłówek, który robi z 27 września 2027 r. początek zakazu. Zostawić ramkę, że wprowadzenie w błąd i pominięcie obowiązują już dziś, i dopisać w tej samej ramce datę 27 września 2026 r. oraz pięć wzorców. Nie dopisywać zdania, że polska ustawa już weszła w życie.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Od 27 września 2026 r. komunikację środowiskową ocenia się według dyrektywy (UE) 2024/825 oraz ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. Wzorcami są: ogólne hasło, twierdzenie o całym produkcie oparte na jednym aspekcie, neutralność klimatu z kompensacji i własny znak zrównoważoności. Druk Senatu 825, po uchwale Sejmu z 4 września 2026 r., przewiduje krajowe wejście tych punktów 27 września 2027 r.”'

    Add-Case 'PRZYPADEK 03: Jedna metoda GHG Protocol dla organizacji, produktu, wydarzenia i usługi' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · UZNK art. 14 ust. 2 pkt 2 i art. 16 ust. 1 pkt 2 · ISO 14067 · C-611/14 | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona 07: „Ślad węglowy to suma emisji gazów cieplarnianych związanych z organizacją, produktem, wydarzeniem albo usługą. Liczymy go według GHG Protocol na kalkulatorze platformy Green0meter: osobno emisje ze spalania własnych paliw, z kupionej energii i z działań poza firmą, takich jak zakupy, transport i odpady.”',
            'Strona 08: „Wskaźnik klimatyczny produktu liczymy w usłudze śladu węglowego.”',
            'Słownik, strona 22: „GHG Protocol. Standard liczenia emisji gazów cieplarnianych. Zakres 1: spalanie własnych paliw. Zakres 2: energia kupiona od dostawcy. Zakres 3: emisje poza firmą, między innymi zakupy, transport, podróże służbowe i odpady.”',
            'Słownik, hasło CO2 Hub: „Wynik wchodzi do raportu według Dyrektywy CSRD, według VSME albo według SFDR.”'
        ) `
        @'
Czytelnik strony 07 wynosi, że jedna kalkulacja według GHG Protocol obejmuje organizację, produkt, wydarzenie albo usługę. Słownik rozpisuje ten standard na zakres 1, 2 i 3 organizacji. Ślad produktu jest innym obiektem. Liczba przy wyrobie wymaga metody ISO 14067, granicy i jednostki. Strona cyklu życia odsyła wskaźnik klimatyczny produktu do usługi śladu, czyli do metody organizacji. Green0meter jest nazwany bez bazy współczynników w zdaniu, które stawia cztery obiekty obok siebie. Rok porównawczy i rozróżnienie danych firmy oraz tabel emisji stoją niżej, w opisie przebiegu. Nie zmieniają pierwszego zdania. Usługa transportowa, gdy wynik ma dotyczyć przewozu, wymaga osobnej normy EN ISO 14083. Tego zdania w pierwszym akapicie nie ma.
'@ `
        'Z pierwszego akapitu strony 07 usunąć „produktem, wydarzeniem albo usługą” przy GHG Protocol. Ze strony 08 usunąć odesłanie wskaźnika klimatycznego produktu do usługi śladu. Ze słownika usunąć zdanie, że wynik CO2 Hub wchodzi do raportu według SFDR. Wskaźnik emisji w ujawnieniu uczestnika rynku finansowego wpisuje się dopiero w zdaniu z przypadku 06, z numerem artykułu i nazwą produktu. Zostawić: rok porównawczy, rozdział spalania własnego, energii kupionej i zakresu 3 oraz informację, które liczby są z danych firmy, a które z tabel. Nazwę bazy wskaźników wpisać w tym samym akapicie co obiekt. Dopisek na stronie 22 nie poprawia pierwszego akapitu strony 07.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Inwentaryzacja gazów cieplarnianych organizacji za rok [RRRR]: zakres 1, zakres 2 i kategorie zakresu 3 uznane za istotne w tym roku. Metoda: GHG Protocol. Wskaźniki: [nazwa bazy]. W raporcie widać, które liczby są z danych firmy, a które z tej bazy. Ślad wyrobu [nazwa] liczymy osobno według ISO 14067, granica [od–do], jednostka [jednostka], rok danych [RRRR].”'

    Add-Case 'PRZYPADEK 04: Analiza cyklu życia od surowca do końca życia i cały zestaw skutków' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · ISO 14040 i ISO 14044 · C-611/14 | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona 08, pierwsze zdanie: „Analiza zgodna z ISO 14040 i ISO 14044 ocenia ten wpływ od surowca do końca życia.”',
            'Ta sama strona, krok 1: „które etapy wchodzą do analizy: od surowca do bramy zakładu albo przez cały okres użytkowania.”',
            'Ta sama strona, krok 3: „Oceniamy cały zestaw skutków dla środowiska objęty normami ISO 14040 i ISO 14044: klimat, wodę, surowce i pozostałe oddziaływania.”',
            'Akademia, strona 17: „Tłumaczymy wpływ produktu od surowca do końca życia, deklarację środowiskową oraz cyfrowy paszport produktu.”'
        ) `
        @'
Czytelnik pierwszego zdania wynosi, że analiza sięga od surowca do końca życia i że obejmuje cały zestaw skutków z ISO 14040 i ISO 14044. Krok pierwszy tej samej strony dopuszcza granicę do bramy zakładu albo przez okres użytkowania. Granica do bramy nie obejmuje użytkowania ani końca życia. Okres użytkowania nie jest końcem życia. Normy wymagają jednej, nazwanej granicy badania. „Cały zestaw” nie jest zamkniętą listą z normy: kategorie wpływu wybiera się w badaniu i wpisuje do raportu. Akademia powtarza formułę „od surowca do końca życia”. Dwie granice w jednym opisie sprawiają, że badanie do bramy czyta się jak pełny cykl. Pierwsze zdanie samo tworzy to wrażenie. Krok pierwszy go nie poprawia.
'@ `
        'Z pierwszego zdania strony 08 i z tematu Akademii usunąć „od surowca do końca życia” oraz „cały zestaw”. W pierwszym zdaniu zostawić jedną granicę, słowami, przy nazwie wyrobu. Drugiej granicy nie stawiać w tym samym zdaniu. Kategoriom wpływu dać listę w tym samym zdaniu, nie słowo „pozostałe”. Zdanie o wyborze surowca może zostać, jeśli granica tego porównania jest ta sama dla obu wariantów i stoi przy zestawieniu. Porównania wariantów nie ogłaszać hasłem o niższym śladzie, dopóki w tym zestawieniu nie ma jednostki, roku i granicy.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Analiza cyklu życia wyrobu [nazwa] według ISO 14040 i ISO 14044. Granica tego zlecenia: od pozyskania surowców do bramy zakładu [nazwa]. Jednostka: [jednostka]. Rok danych: [RRRR]. Kategorie wpływu w raporcie z tej analizy: [lista].”'

    Add-Case 'PRZYPADEK 05: Cztery wskaźniki klimatu już policzone do deklaracji właściwości użytkowych' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · rozporządzenie (UE) 2024/3110 · EN 15804 | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona 10, ramka „Dane pod przyszłą deklarację”: „Cztery wskaźniki klimatu do przyszłej deklaracji właściwości użytkowych są już policzone.”',
            'Strona 09: „Rozporządzenie (UE) 2024/3110 stosuje się od tego dnia. Charakterystyki środowiskowe w deklaracji właściwości użytkowych wchodzą z nową specyfikacją zharmonizowaną danej rodziny wyrobów, etapami do 2032 r.”',
            'Strona 10, krok 3: „Klimat podajemy w czterech liczbach: łącznie, z paliw kopalnych, z materiałów biologicznych i ze zmiany użytkowania gruntów.”'
        ) `
        @'
Czytelnik ramki wynosi, że cztery wskaźniki klimatu do deklaracji właściwości użytkowych są już policzone. Strona 09 mówi, że charakterystyki środowiskowe w tej deklaracji wchodzą ze specyfikacją zharmonizowaną danej rodziny, etapami do 2032 r. Cztery liczby z kroku 3 są podziałem wskaźnika klimatu w obliczeniu do deklaracji środowiskowej. Nie są z tego powodu treścią deklaracji właściwości użytkowych rodziny, której specyfikacja jeszcze nie weszła. Zdanie „są już policzone” przenosi wynik obliczenia do dokumentu, którego zakres ta sama broszura odkłada na specyfikację rodziny. Krok 3, czytany sam, podział opisuje. Ramka przenosi te liczby do deklaracji właściwości użytkowych.
'@ `
        'Z ramki 03 na stronie 10 usunąć „są już policzone” i słowa „do przyszłej deklaracji właściwości użytkowych”. Zostawić cztery liczby jako podział klimatu w obliczeniu do deklaracji środowiskowej, z nazwą wyrobu, programu i rokiem w tym samym zdaniu. Zdania o deklaracji właściwości użytkowych nie publikować, dopóki specyfikacja zharmonizowana tej rodziny nie wskazuje, które wskaźniki do niej wchodzą. Wtedy osobne zdanie podaje numer aktu i rodzinę. Strona 09 o etapach do 2032 r. zostaje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Wskaźnik klimatu wyrobu [nazwa] w obliczeniu do deklaracji środowiskowej, program [nazwa], rok danych [RRRR]: łącznie, z paliw kopalnych, z materiałów biologicznych i ze zmiany użytkowania gruntów.”'

    Add-Case 'PRZYPADEK 06: Raport według SFDR na równi z CSRD i VSME' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · UZNK art. 14 ust. 2 pkt 2 · rozporządzenie (UE) 2019/2088 · C-611/14 | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona 15, nadtytuł: „Zgodność w jednym miejscu.”',
            'Strona 15: „Green0meter to kompleksowa platforma online wspierająca firmy w raportowaniu ESG i zgodności regulacyjnej.”',
            'Strona 15: „Raport według Dyrektywy CSRD, według VSME albo według SFDR, z liczeniem śladu węglowego.”',
            'Słownik, strona 22: „SFDR. Rozporządzenie (UE) 2019/2088 o ujawnianiu informacji związanych ze zrównoważonym rozwojem w sektorze usług finansowych. Dotyczy uczestników rynku finansowego i doradców.”'
        ) `
        @'
Czytelnik strony 15 wynosi, że platforma daje zgodność regulacyjną, a raport według SFDR stoi na równi z raportem CSRD i VSME, razem z liczeniem śladu. Słownik, siedem stron dalej, ogranicza SFDR do uczestników rynku finansowego i doradców. Na stronie platformy nie ma art. 8 ani art. 9, nazwy produktu finansowego ani listy wskaźników. Samo hasło „zgodność” czyta się jak wynik usługi. Ślad organizacji nie zastępuje ujawnienia produktu finansowego. Słownik tego układu nie poprawia, bo pierwsze zdanie modułu już zestawiło trzy raporty jako ten sam rodzaj usługi dla firm.
'@ `
        'Z nadtytułu strony 15 usunąć „Zgodność w jednym miejscu”. Z listy modułów usunąć „według SFDR”, dopóki w tym samym zdaniu nie ma art. 8 albo art. 9 i nazwy produktu. Słowo „kompleksowa” przy „zgodności regulacyjnej” zdjąć. Zostawić moduły nazwane z osobna: raport CSRD, raport VSME, CO2 Hub, Paygapmeter, Regumeter, Cyberometer. Definicja SFDR w słowniku zostaje. Nie dopisywać na stronie 15 lekcji, czym ujawnienie nie jest.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Dla uczestnika rynku finansowego: projekt ujawnień produktu [nazwa] według rozporządzenia (UE) 2019/2088, art. [8 albo 9], za rok [RRRR], ze wskaźnikami wymienionymi w tych ujawnieniach.”'

    Add-Case 'PRZYPADEK 07: Medal EcoVadis jako warunek kluczowych odbiorców' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 2 · C-281/12 | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona 21, pierwsze zdanie: „EcoVadis ocenia polityki, działania i wyniki firmy oraz przyznaje medal, który kluczowi odbiorcy traktują jako warunek współpracy.”',
            'Ta sama strona, ramka: „Dla części odbiorców medal jest warunkiem zaproszenia do współpracy.”',
            'Ta sama strona: „Przegląd przed wysłaniem ankiety ogranicza ryzyko wyniku poniżej oczekiwań klienta.”',
            'Słownik: „Wynik, często w formie medalu, bywa kryterium wyboru w relacjach B2B.”'
        ) `
        @'
Czytelnik pierwszego zdania wynosi, że ocena kończy się medalem i że kluczowi odbiorcy traktują medal jako warunek współpracy. Ramka na tej samej stronie mówi o części odbiorców. Słownik mówi, że wynik często ma formę medalu i bywa kryterium. Trzy zdania nie opisują tej samej grupy ani tego samego wyniku. Medal przyznaje EcoVadis. Pierwsze zdanie nie podaje skali ani progu. Zdanie o ograniczeniu ryzyka wyniku poniżej oczekiwań klienta nie podaje progu ani roku. Zapowiedź medalu i progu, których warunków nie ma przy zdaniu, jest wprowadzeniem w błąd co do wyniku usługi. Wejście w rozmowę o przygotowaniu jest decyzją dotyczącą umowy (C-281/12). Punkty o przeglądzie polityk i o spójności odpowiedzi z raportem mogą zostać: opisują dokument, nie medal.
'@ `
        'Z pierwszego zdania usunąć człon o medalu jako warunku kluczowych odbiorców. Z ramki usunąć medal jako warunek zaproszenia. Usunąć zdanie o ograniczeniu ryzyka wyniku poniżej oczekiwań. Jeśli zdanie o odbiorcy zostaje, w tym samym zdaniu jest nazwa odbiorcy, rok i to, czy wymaga wyniku, czy medalu. Progu medalu nie publikować, jeśli broszura go nie podaje. Na stronie nie dopisywać wyjaśnienia, kto medalu nie gwarantuje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Przegląd dokumentacji przed złożeniem kwestionariusza EcoVadis za rok [RRRR], w filarach środowisko, praca i prawa człowieka, etyka oraz zrównoważone zakupy. Każda odpowiedź w ankiecie ma pokrycie w załączonym pliku.”'

    Add-Case 'PRZYPADEK 08: Spadek emisji jako niższy rachunek i niższy koszt wytworzenia' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 2 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona 07, ramka „Optymalizacja kosztów”: „Wyliczenie pokazuje Państwu, gdzie spadek emisji obniża też rachunki za energię i transport. Redukcję zaczynamy od tych miejsc.”',
            'Strona 08, ramka „Niższy koszt wytworzenia”: „Wskazujemy etap, na którym mniejszy wpływ na środowisko ogranicza też straty materiału i energii. Zmianę wyrobu zaczynamy od tego etapu.”'
        ) `
        @'
Czytelnik ramki na stronie śladu wynosi, że spadek emisji obniża rachunki za energię i transport. Czytelnik ramki na stronie cyklu życia wynosi, że mniejszy wpływ ogranicza straty i daje niższy koszt wytworzenia. Żadna ramka nie podaje roku, pozycji kosztu ani warunku, przy którym rachunek spada. Spadek emisji i spadek kosztu to dwa wyniki. Złączenie ich bez pozycji i bez roku przedstawia korzyść finansową jako cechę wyliczenia. Plan redukcji opisany obok, jako cele i działania przy energii, transporcie albo zakupach, tego połączenia nie potrzebuje.
'@ `
        'Z obu ramek usunąć spadek rachunku, niższy koszt i ograniczenie strat. Zdanie o miejscach o największym udziale w emisji może zostać z rokiem i z zakresem w tym samym zdaniu. Koszt publikować tylko przy nazwanej pozycji, roku i kwocie, w osobnym zdaniu, nie przy emisji. Plan działań przy energii, transporcie albo zakupach zostaje bez słowa o rachunku.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Wyliczenie za rok [RRRR] wskazuje pozycje energii i transportu o największym udziale w emisjach zakresu 1 i 2. Działania redukcji przypisujemy do tych pozycji, z rokiem bazowym [RRRR].”'

    Add-Case 'PRZYPADEK 09: Kluczowi odbiorcy wymagają śladu węglowego' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 i art. 6 · UZNK art. 16 ust. 1 pkt 2 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona 07, ramka „Oczekiwanie odbiorcy”: „Kluczowi odbiorcy wymagają śladu węglowego od dostawców. Brak wyliczenia ogranicza udział w postępowaniu.”'
        ) `
        @'
Czytelnik ramki wynosi, że kluczowi odbiorcy wymagają śladu i że brak wyliczenia ogranicza udział w postępowaniu. Nie ma roku, branży ani nazwy postępowania. „Kluczowi odbiorcy” czyta się jak całą grupę. Na stronie EcoVadis podobna korzyść jest węższa: medal jest warunkiem dla części odbiorców. Przy śladzie tego zawężenia nie ma. Zapowiedź utraty postępowania bez nazwy i bez roku zmienia znaczenie usługi przy decyzji o zleceniu.
'@ `
        'Ramkę „Oczekiwanie odbiorcy” zdjąć. Zdania o wymogu śladu nie publikować, dopóki w tym samym zdaniu nie ma nazwy zamawiającego, roku postępowania i roku inwentaryzacji. Nie zastępować „kluczowych odbiorców” słowem „rynek” ani „wszyscy odbiorcy”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W postępowaniu [nazwa zamawiającego] z [RRRR] r. warunkiem udziału była inwentaryzacja emisji organizacji za rok [RRRR], zakres 1 i 2.”'

    Add-Case 'PRZYPADEK 10: Liczby z deklaracji EPD w ocenie BREEAM i LEED oraz porównanie wyrobów' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · EN 15804 · rozporządzenie (UE) 2024/3110 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona 09: „Liczby wchodzą do specyfikacji i do oceny budynku, na przykład BREEAM i LEED.”',
            'Strona 10: „Wartości z deklaracji środowiskowej EPD, wymagane w specyfikacjach oraz w systemach oceny budynków, takich jak BREEAM i LEED.”',
            'Strona 10: „Etapy cyklu życia i wskaźniki, według których odbiorca w budownictwie porównuje wyroby między sobą.”'
        ) `
        @'
Czytelnik tych zdań wynosi, że liczby z deklaracji wchodzą do oceny budynku BREEAM i LEED, że są w tych systemach wymagane i że odbiorca porównuje wyroby według etapów EN 15804. Nie ma wersji schematu, poziomu, numeru kredytu ani modułów. BREEAM i LEED są zestawami kredytów. Deklaracja wyrobu nie jest oceną budynku i nie daje poziomu certyfikatu. Porównanie wyrobów według EN 15804 wymaga tych samych modułów, tej samej jednostki i reguł tej samej grupy wyrobów. Tych trzech rzeczy w zdaniu o porównaniu nie ma.
'@ `
        'Z obu stron usunąć „wchodzą do oceny budynku” i słowo „wymagane” przy BREEAM i LEED. Usunąć zdanie, że odbiorca porównuje wyroby między sobą, dopóki przy zestawieniu nie ma modułów, jednostki, grupy wyrobów i programu. Nazwę schematu publikować z wersją i z nazwą kredytu, który dane z nazwanych modułów przyjmuje. Nie publikować poziomu budynku ani słowa „zgodny z BREEAM” przy deklaracji wyrobu.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Deklaracja środowiskowa wyrobu [nazwa], program [operator], moduły [lista], jednostka [jednostka], rok danych [RRRR]. Dane z tych modułów można zgłosić w kredycie [nazwa] schematu [BREEAM albo LEED], wersja [wersja], jeśli regulamin tego kredytu taką deklarację przyjmuje.”'

    Add-Case 'PRZYPADEK 11: Jeden komplet deklaracji i ważność pięć lat' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 · UZNK art. 14 ust. 2 pkt 2 · C-611/14 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona 08, ramka „Deklaracja EPD w jednej usłudze”: „Gdy celem jest deklaracja środowiskowa, obliczenie cyklu życia wchodzi w przygotowanie dokumentacji. Odbierają Państwo jeden komplet.”',
            'Strona 09: „Ważność zwykle pięć lat.”',
            'Strona 10: „Deklaracja jest ważna pięć lat. Wcześniejsza zmiana danych oznacza, że przygotowujemy ją od nowa.”',
            'Strona 10: „Wynagrodzenie osoby sprawdzającej i opłatę za wpis do rejestru producent płaci bezpośrednio organizacji, która ten rejestr prowadzi.”'
        ) `
        @'
Czytelnik ramki na stronie 08 wynosi, że w jednej usłudze odbiera komplet deklaracji. Strona 10 mówi, że wpis i opłatę producent uzgadnia z operatorem, a walidację prowadzi osoba od niego niezależna. „Jeden komplet” bez słowa „do złożenia” czyta się jak deklarację już opublikowaną. Na stronie 09 ważność jest zwykle pięcioletnia. Na stronie 10 deklaracja jest ważna pięć lat, bez nazwy programu. Okres zależy od regulaminu programu. Dwa zdania o tym samym okresie stoją w jednym rozdziale i nie mówią tego samego. Strona 10, czytana sama przy opłacie, zakres zlecenia oddziela. Ramka ze strony 08 tego rozdziału nie powtarza.
'@ `
        'Z ramki na stronie 08 usunąć „jeden komplet”. Zostawić, że obliczenie wchodzi w dokumenty składane operatorowi. Na stronie 10 zastąpić „jest ważna pięć lat” okresem z regulaminu nazwanego programu. Zdanie „zwykle pięć lat” zdjąć albo zastąpić tym samym okresem. Zostawić zdanie, że opłatę za wpis producent płaci operatorowi. Nie dopisywać lekcji, że instytut nie jest operatorem: nazwa programu i słowa „do złożenia” ten zakres niosą.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Dokumenty do złożenia u operatora [nazwa programu]: deklaracja, raport dla osoby sprawdzającej i zestawienie danych z roku [RRRR]. Okres ważności w regulaminie tego programu: [liczba] lat od publikacji. Wpis i opłatę producent uzgadnia z tym operatorem.”'

    Add-Case 'PRZYPADEK 12: Procedura ciągłości jako spełnienie NIS2, DORA i norm' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · UZNK art. 14 ust. 2 pkt 2 i 4 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona 12, ramka „NIS2, DORA, normy”: „Przepisy o cyberbezpieczeństwie, odporności cyfrowej finansów i normy ciągłości działania wymagają udokumentowanego planu. Przygotowujemy go w formie procedury.”'
        ) `
        @'
Czytelnik ramki wynosi, że przepisy cyberbezpieczeństwa, odporności cyfrowej finansów i normy ciągłości wymagają planu oraz że procedura przygotowana w usłudze ten wymóg zamyka. Nie ma zdania, którego podmiotu przepis dotyczy i której normy. Plan ciągłości jest dokumentem o procesach, rolach i próbie. Nie jest z tego zdania zgodnością z całym reżimem. Brak podmiotu i brak nazwy obowiązku sprawia, że usługa czyta się jak spełnienie prawa, którego broszura nie nazywa. Opis przebiegu, od analizy wpływu do próby, tego skrótu nie potrzebuje.
'@ `
        'Z ramki usunąć NIS2, DORA i „normy ciągłości” jako spełnione procedurą. Zostawić przebieg: procesy krytyczne, role, przyjęty czas przestoju, próba z datą. Nazwę reżimu wstawić z powrotem tylko wtedy, gdy w tym samym zdaniu jest podmiot objęty tym reżimem, nazwa obowiązku i dzień, na który sprawdzono ten obowiązek. Nie publikować zdania, że procedura jest zgodna z całym aktem.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Plan ciągłości procesów [nazwa procesów] dla [nazwa firmy]: role, czas przestoju przyjęty w analizie i próba z dnia [data].”'

    Add-Case 'PRZYPADEK 13: Udział w Collaboration Hub wyróżnia organizację' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 16 ust. 1 pkt 2 i ust. 2 · C-281/12 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona 18, ramka „Zaufanie do marki”: „Udział w społeczności i w materiałach Hubu wzmacnia zaufanie klientów, pracowników i partnerów oraz wyróżnia organizację na rynku.”',
            'Strona 18: „Uczestnictwo osoby indywidualnej wynosi 499 PLN netto za rok, a VAT 23% doliczamy osobno.”'
        ) `
        @'
Czytelnik ramki wynosi, że udział w Hubie wzmacnia zaufanie trzech grup i wyróżnia organizację na rynku. Nie ma roku, wskaźnika ani grupy, wobec której organizacja miałaby się wyróżniać. Członkostwo osoby jest wycenione: 499 zł netto za rok. Przy tej cenie zdanie o wyróżnieniu jest zapowiedzią wyniku członkostwa. Wyniku nie da się sprawdzić z broszury. Reklama, która może wpłynąć na zakup, a nie podaje cechy, którą da się sprawdzić, wprowadza w błąd. Zapis osoby indywidualnej jest decyzją dotyczącą umowy (C-281/12). Lista tematów i cena tego zdania nie potrzebują.
'@ `
        'Z ramki usunąć „wzmacnia zaufanie” i „wyróżnia organizację na rynku”. Zostawić to, co członek dostaje: spotkania, bazę, newsletter, cenę i stawkę VAT. Nie publikować porównania z firmami spoza Hubu. Liczba „ponad 100 osób” może zostać tylko z dniem, na który ją policzono, w tym samym zdaniu.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Członkostwo Member: przez rok dostęp do spotkań Tea Time i Coffee Call, bazy materiałów Hubu i newslettera. Cena osoby indywidualnej: 499 zł netto za rok, VAT 23%.”'

    Add-H '4. Ryzyka prawne i finansowe'
    Add-P '• 1. Prezes UOKiK i kara od obrotu' 11 $true $false $NAVY 10 2
    Add-P 'Broszura jest materiałem handlowym, a członkostwo Hubu jest ofertą dla osoby fizycznej za oznaczoną cenę. Praktyka naruszająca zbiorowe interesy konsumentów jest zagrożona karą do 10% obrotu za każdą praktykę, także nieumyślną (art. 24 i art. 106 ustawy o ochronie konkurencji i konsumentów). Okładka jest hasłem z czarnej listy. Przy takim haśle nie bada się osobno, czy czytelnik zdążył kupić. Przy pozostałych zdaniach wejście w rozmowę albo zakup członkostwa jest decyzją dotyczącą umowy (C-281/12, art. 2 pkt 7 UPNPR). Hasło na okładce i data zakazu na stronie 11 są na tyle samodzielne, że pracownik, odbiorca albo konkurent może zgłosić je Prezesowi UOKiK.' 10.5 $false $false $DARK 0 8
    Add-P '• 2. Roszczenie konkurenta i kontrahenta' 11 $true $false $NAVY 8 2
    Add-P 'UZNK art. 14 zakazuje rozpowszechniania nieprawdziwej albo wprowadzającej w błąd wiadomości o świadczonej usłudze i o sytuacji prawnej. Art. 16 ust. 1 pkt 2 dotyczy reklamy, która wprowadza klienta w błąd i może wpłynąć na zakup. Art. 18 daje żądanie zaniechania, usunięcia skutków, oświadczenia, naprawienia szkody i wydania korzyści. Kontrahent, który zlecił audyt twierdzeń na podstawie strony 11, czyta datę 27 września 2027 r. Konkurent, który tej daty i tego hasła na okładce nie używa, ma interes w żądaniu zaniechania.' 10.5 $false $false $DARK 0 8
    Add-P '• 3. Raport i deklaracja klienta' 11 $true $false $NAVY 8 2
    Add-P 'Klient, który powtórzy hasło z okładki, wpisze ślad wyrobu metodą organizacji albo przeniesie cztery wskaźniki klimatu do deklaracji właściwości użytkowych, przenosi wadę broszury do własnego raportu, do karty wyrobu i do pytań biegłego. Rozbieżność między tym plikiem a obliczeniem jest zgłoszeniem z zakresu ochrony konsumenta i ochrony środowiska. Od 25 grudnia 2024 r. takie zgłoszenie można złożyć Prezesowi UOKiK z pominięciem drogi wewnętrznej. Klauzula poufności tego zgłoszenia nie zamyka.' 10.5 $false $false $DARK 0 8
    Add-P '• 4. Data, według której klient odkłada korektę' 11 $true $false $NAVY 8 2
    Add-P 'Strona 11 jest opisem usługi, która ma wskazać, co wycofać od razu. Ramka o 27 września 2027 r. robi z tej daty początek zakazu haseł, które okładka tej samej broszury już używa. Klient, który według tej ramki zostawi ogólne hasło przez kolejny rok, zbiera ryzyko kary i ryzyko powtórzenia hasła u własnego odbiorcy. Korekta strony 11 jest częścią korekty okładki, nie osobnym projektem redakcyjnym.' 10.5 $false $false $DARK 0 8

    Add-H '5. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIASTOWE ZDJĘCIE ZDAŃ (1 listopada 2026 r.)' @(
        'Z okładki zdjąć „Tworzyć w pełni zrównoważoną przyszłość, już dziś”. Ze spisu zdjąć „zrównoważony model działania”.',
        'Ze strony 11 i ze słownika zdjąć „Dz.U. 2026 poz. 1353” oraz początek zakazu w dacie 27 września 2027 r.',
        'Ze strony 07 zdjąć produkt, wydarzenie i usługę przy GHG Protocol, spadek rachunków oraz wymóg kluczowych odbiorców. Ze strony 08 zdjąć „od surowca do końca życia”, „cały zestaw”, niższy koszt i „jeden komplet”. To samo zdjęcie granicy w temacie Akademii.',
        'Ze stron 09 i 10 zdjąć „są już policzone”, „wchodzą do oceny budynku”, „wymagane” przy BREEAM i LEED oraz „jest ważna pięć lat”. Ze strony 15 zdjąć „zgodność w jednym miejscu” i SFDR z listy raportów dla firm. Ze strony 21 zdjąć medal jako warunek i zdanie o ryzyku wyniku. Ze strony 12 zdjąć NIS2 i DORA jako spełnione procedurą. Ze strony 18 zdjąć „wyróżnia organizację”.'
    ) $REDBG $DANGER

    Add-Callout 'ETAP 2: NOWE ZDANIA W TYCH SAMYCH MIEJSCACH (do 8 listopada 2026 r.)' @(
        'Wstawić wzory z przypadków 01–13. Rok, obiekt, granica, program, art. 8 albo 9 i nazwa postępowania stoją w tym samym zdaniu co czynność.',
        'Zostawić bez zmiany: zdanie, że samo wyliczenie nie jest deklaracją neutralności; rozdział danych firmy i tabel emisji; opłatę za wpis u operatora; „część banków i funduszy” przy VSME; rozdział ustawy i projektu w rozdziale o dyrektywie CSRD; rozdział projektu UC127 przy luce płacowej; przebieg planu ciągłości bez nazw reżimów.',
        'Cena 499 zł netto i VAT 23% zostają przy członkostwie osoby, bez zdania o wyróżnieniu.'
    ) (Rgb 255 247 237) $WARNING

    Add-Callout 'ETAP 3: METODYKA W TEJ SAMEJ BROSZURZE (do 1 grudnia 2026 r.)' @(
        'Przy śladzie organizacji, w broszurze, nie w osobnym pliku: rok, nazwa bazy wskaźników, udział danych firmy i lista kategorii zakresu 3 uznanych za istotne. Ślad wyrobu ma osobne zdanie z ISO 14067. Przewóz jako osobny wynik ma osobne zdanie z EN ISO 14083.',
        'Przy deklaracji: program, moduły, jednostka, rok i okres ważności z regulaminu. Przy BREEAM albo LEED: wersja i kredyt, nie poziom budynku. Przy SFDR: art. 8 albo art. 9 i nazwa produktu.',
        'Słownik i strona 11 podają tę samą datę stosowania dyrektywy: 27 września 2026 r. Dopisek w słowniku nie poprawia okładki ani pierwszego zdania strony 07, które łączy cztery obiekty jedną metodą (UPNPR art. 5, C-611/14).'
    ) (Rgb 239 246 255) $ACCENT

    Add-Callout 'ETAP 4: SŁOWNIK ZDAŃ DO BROSZURY (od 1 grudnia 2026 r.)' @(
        'Każde nowe zdanie przechodzi pięć pytań: czy jest ogólnym hasłem bez znaku typu I w tym miejscu, czy metoda pasuje do obiektu, czy liczba ma rok, czy granica jest jedna, czy data prawa jest datą stosowania dyrektywy, a nie datą z druku Senatu.',
        'Cel redukcji na przyszłość, jeśli broszura go ogłasza, ma w tym samym miejscu termin, zasób i regularną weryfikację niezależnego eksperta, z wynikiem dostępnym odbiorcy. Samo wyliczenie nie jest takim celem. Zdanie o braku deklaracji neutralności zostaje.',
        'Zdanie zdjęte w etapie 1 nie wraca w tej broszurze, na stronie, w ofercie ani w zaproszeniu na rozmowę. Usługa audytu ze strony 11 używa tych samych pytań wobec własnego pliku.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '6. Wnioski'
    Add-P 'Broszura handlowa 2026 w stanie pliku z 31 października 2026 r. nie nadaje się do wysyłki. Okładka jest ogólnym hasłem środowiskowym o całości, a strona audytu przesuwa start tych zakazów na 27 września 2027 r. i podaje druk Senatu jako ustawę. Do wysyłki plik wraca po zdjęciu tego hasła i po rozdzieleniu metody od obiektu, którego metoda dotyczy.' 10.5 $false $false $DARK 0 8

    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $doc.SaveAs2([string]$outPath)
    $doc.Close($false)
    $doc = $null
    Copy-Item -LiteralPath $outPath -Destination $finalPath -Force
    Write-Output "SAVED $finalPath"
}
finally {
    if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
    if ($word -ne $null) { try { $word.Quit() } catch {} }
}
