#Requires -Version 5.1
# Audyt sekcji http://localhost:8085/#uslugi zgodnie z INSTRUKCJA_GENEROWANIA_RAPORTOW.md
$ErrorActionPreference = 'Stop'
$finalPath = Join-Path $PSScriptRoot 'Raport_Audytu_Greenwashingu_ESG_Institute_Uslugi.docx'
$outPath = Join-Path $env:TEMP 'Raport_Audytu_Greenwashingu_ESG_Institute_Uslugi.docx'

function Rgb([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }

$NAVY    = Rgb 8 24 31
$ACCENT  = Rgb 42 138 176
$DANGER  = Rgb 190 18 60
$WARNING = Rgb 180 83 9
$SUCCESS = Rgb 4 120 87
$DARK    = Rgb 22 36 44
$MUTED   = Rgb 90 107 119
$WHITE   = Rgb 255 255 255
$AMBERBG = Rgb 255 247 237

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
            [System.IO.File]::WriteAllText((Join-Path $env:TEMP 'esg_logo_err.txt'), ($_ | Out-String))
            $logoOk = $false
        }
    } else {
        [System.IO.File]::WriteAllText((Join-Path $env:TEMP 'esg_logo_err.txt'), "missing $logo")
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
    $footer.Range.Text = 'ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  sekcja Usługi  ·  21 października 2026 r.'
    $footer.Range.Font.Name = 'Arial'
    $footer.Range.Font.Size = 8
    $footer.Range.Font.Color = $MUTED

    $word.ActiveWindow.ActivePane.View.SeekView = 0
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)

    Add-P 'ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO' 9 $true $false $ACCENT 0 2
    Add-P 'Audyt Ryzyk Greenwashingu: ESG INSTITUTE' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja sekcji „Nasze usługi” pod adresem http://localhost:8085/#uslugi: osiem kart oraz okno „Dowiedz się więcej”, które powtarza te same zdania. Stan zebrania i data audytu: 21 października 2026 r. Ocena wobec ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, ustawy o zwalczaniu nieuczciwej konkurencji, ustawy o ochronie konkurencji i konsumentów, dyrektywy 2005/29/WE oraz dyrektywy (UE) 2024/825 w zakresie twierdzeń o wyniku usługi i o metodzie liczenia.' 10.5 $false $true $MUTED 0 8

    $shot = Join-Path $PSScriptRoot 'zrzut_uslugi.png'
    if (Test-Path -LiteralPath $shot) {
        $sel.ParagraphFormat.Alignment = 1
        $shape = $sel.InlineShapes.AddPicture($shot)
        $shape.LockAspectRatio = $true
        $shape.Width = $word.CentimetersToPoints(15.4)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.Alignment = 0
    }
    Add-P 'Zrzut górnej części sekcji „Nasze usługi”, http://localhost:8085/#uslugi, 21 października 2026 r. W pierwszym rzędzie karty CSRD, VSME i śladu węglowego.' 8.5 $false $true $MUTED 2 10

    $meta = @(
        @('Badany podmiot', 'ESG Institute, Rondo ONZ 1, 00-124 Warszawa, tel. +48 797 301 987, info@esginstitute.eu. Oddziały Hubu podane w stopce: Warszawa, Poznań, Kraków, Wrocław, Gdańsk, Rzeszów. Forma prawna i numer KRS nie są podane w badanej sekcji ani w stopce.'),
        @('Przedmiot audytu', 'Sekcja „Nasze usługi” serwowana 21 października 2026 r. pod http://localhost:8085/#uslugi. Osiem kart i okno „Dowiedz się więcej”. Okno nie dodaje innych zdań niż karta. Kalkulator pod sekcją oraz pozostałe działy witryny nie wchodzą do ustaleń. Tej kopii nie porównywano ze stroną esginstitute.eu.'),
        @('Sektor działalności', 'Doradztwo dla firm: raportowanie zrównoważonego rozwoju, ślad gazów cieplarnianych, ocena cyklu życia, rating EcoVadis, prawo AI. Odbiorca karty: zarząd, compliance i zespół sustainability. Strona jest publiczna.'),
        @('Reżim prawny', 'UPNPR art. 5 ust. 2 pkt 1 i 2, ust. 3 pkt 2 i 6 oraz art. 6. UZNK art. 14 ust. 2 pkt 2 i 4, ust. 3 pkt 1 oraz art. 18. Ustawa o ochronie konkurencji i konsumentów art. 24 i 106. Dyrektywa 2005/29/WE. Dyrektywa (UE) 2024/825: załącznik I sprawdzony; w tej sekcji nie ma własnego znaku zrównoważoności, ogólnego hasła „eko/zielony/zrównoważony”, neutralności opartej na kompensacji ani obowiązku prawnego pokazanego jako cecha środowiskowa. Metodyka liczb: GHG Protocol Corporate Standard, ISO 14064-1, ISO 14067, ISO 14040, ISO 14044, EN 15804. Paszport produktu: rozporządzenie (UE) 2024/1781.'),
        @('Data audytu', '21 października 2026 r.'),
        @('Werdykt operacyjny', 'PUBLIKACJA WARUNKOWA PO KOREKCIE (KOD ŻÓŁTY)')
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
        if ($last) { Shade-Cell $cv $AMBERBG } else { Shade-Cell $cv $WHITE }
        Pad-Cell $ck 3 3 6 6
        Pad-Cell $cv 3 3 6 6
        Write-Cell $ck $meta[$i][0] 8.5 $true $DARK
        Write-Cell $cv $meta[$i][1] 8.5 $last $(if ($last) { $WARNING } else { $DARK })
    }
    Move-AfterTable $mt
    Add-P '' 6 $false $false $DARK 0 6

    Add-Callout 'WERDYKT: PUBLIKACJA WARUNKOWA PO KOREKCIE (KOD ŻÓŁTY)' @(
        'Sekcja usług nie używa ogólnego hasła „eko”, „zielony” ani „zrównoważony” jako cechy produktu. Nie obiecuje neutralności opartej na kompensacji i nie sprzedaje własnego znaku zrównoważoności. Załącznik I dyrektywy (UE) 2024/825 nie daje w tym materiale podstawy do kodu czerwonego.',
        'Karty obiecują wynik, którego doradca nie wydaje: pozytywną atestację niezależnego biegłego rewidenta, medal EcoVadis oraz ochronę kontraktów i preferencyjne finansowanie. „Precyzyjny” i „pełny” bilans organizacji i produktów jest przypisany do standardu korporacyjnego. Ocena nazwana kompleksową stoi przy granicy do bramy zakładu.',
        'To działanie wprowadzające w błąd (UPNPR art. 5) i wiadomość o usłudze oraz o kwalifikacjach (UZNK art. 14). Wejście w kartę i w „Dowiedz się więcej” jest decyzją dotyczącą umowy (C-281/12). Do publikacji karty wracają po zawężeniu leadu do czynności, którą instytut wykonuje, z rokiem, granicą i nazwą standardu w tym samym zdaniu.'
    ) $AMBERBG $WARNING

    Add-H '1. Analiza Wstępna'
    Add-P 'Sekcja „Nasze usługi” sprzedaje doradztwo w trzech obszarach: środowisko, społeczeństwo i ład korporacyjny. Osiem kart obejmuje raportowanie CSRD i ESRS, dobrowolny standard VSME, bilans gazów cieplarnianych, ocenę cyklu życia i paszport produktu, ciągłość działania, przygotowanie do EcoVadis, sztuczną inteligencję oraz abonament. Narracja jest ekspercka. Odznaki nazywają standard albo grupę odbiorców: „Kluczowa usługa”, „Nowość Omnibus 2026”, „GHG Protocol”, „ESPR & cyrkularność”, „Globalny rating”.' 10.5 $false $false $DARK 0 8
    Add-P 'Wzorzec ryzyka jest inny niż ogólne hasło na opakowaniu. Lead czterech kart obiecuje rezultat leżący poza instytutem: opinię biegłego rewidenta, medal Bronze, Silver lub Gold, ochronę kontraktu i tańsze finansowanie, pełną zgodność z AI Act. Dwie karty środowiskowe mieszają obiekt liczenia. Bilans „organizacji i produktów” jest przypisany do GHG Protocol Corporate Standard oraz ISO 14064. Ocena „kompleksowa” stoi przy ukośniku „Cradle-to-Grave / Cradle-to-Gate”. Abonament, na którym jest „audyt anty-greenwashingowy”, przedstawia zespół audytorów i inżynierów środowiska bez imion.' 10.5 $false $false $DARK 0 8
    Add-P 'Karta ciągłości działania opisuje analizę wpływu na biznes, plany ciągłości, odtwarzanie po awarii i ćwiczenia sztabu. Nie przypisuje firmie wyniku środowiskowego. Zdanie o zakresie 2 w ujęciu market-based i location-based zostaje: tak właśnie rozdziela się energię kupioną. Przycisk „Dowiedz się więcej” otwiera te same zdania, które już stoją na karcie. Nie dopisuje roku, granicy ani progu medalu. Kalkulator pod sekcją nie był przedmiotem ustaleń.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'
    $findings = @(
        @('01','Gwarancja pozytywnej atestacji biegłego rewidenta','UPNPR art. 5 ust. 2 pkt 1 · UZNK art. 14','WYSOKIE','Opinia niezależnego biegłego jako wynik usługi'),
        @('02','„Ponad 1000 punktów danych ESRS”','UPNPR art. 5 ust. 2 pkt 2 i art. 6','WYSOKIE','Liczba bez roku i bez relacji do podwójnej istotności'),
        @('03','Medale EcoVadis Bronze, Silver lub Gold','UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14','WYSOKIE','Medal i „optymalizacja punktacji” jako cel usługi'),
        @('04','Ochrona kontraktów i preferencyjne finansowanie','UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14','WYSOKIE','Korzyść umowna i finansowa z dobrowolnego VSME'),
        @('05','„Precyzyjny” bilans organizacji i produktów, „pełne” 15 kategorii','UPNPR art. 5 i 6 · ISO 14064-1 / 14067','WYSOKIE','Standard korporacyjny na organizację i na produkt'),
        @('06','„Kompleksowa” LCA i granica grave/gate','UPNPR art. 5 ust. 2 pkt 2 i art. 6','WYSOKIE','Dwie granice cyklu życia za jednym ukośnikiem'),
        @('07','„W pełnej zgodzie z unijnym AI Act”','UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14','WYSOKIE','Zgodność z całym aktem przy audycie klasyfikacji ryzyka'),
        @('08','EPD bez programu, modułów i wyrobu','UPNPR art. 6 · EN 15804','ŚREDNIE','Deklaracja nazwana bez zakresu, który EPD przenosi'),
        @('09','„Zgodnie z normami ESPR”','UPNPR art. 5 ust. 2 pkt 2 · rozp. 2024/1781','ŚREDNIE','Rozporządzenie nazwane normą, bez grupy wyrobów'),
        @('10','„Minimalizacja śladu węglowego modeli”','UPNPR art. 5 i 6','ŚREDNIE','Mniejszy ślad bez roku, kWh i lokalizacji'),
        @('11','Dekarbonizacja „i redukcji kosztów operacyjnych”','UPNPR art. 5 ust. 3 pkt 2','ŚREDNIE','Spadek kosztu jako cecha strategii emisji'),
        @('12','„Zespół prawników, audytorów i inżynierów środowiska”','UZNK art. 14 ust. 3 pkt 1 · UPNPR art. 5 ust. 3 pkt 6','ŚREDNIE','Tytuły bez imion; zespół witryny ich nie pokrywa')
    )

    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $st = $doc.Tables.Add($range, $findings.Count + 1, 5)
    $st.PreferredWidthType = 3
    $st.PreferredWidth = $contentW
    $st.AllowAutoFit = $false
    $widths = @(1.1, 5.5, 4.3, 2.3, 4.4)
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

    Add-Case 'PRZYPADEK 01: Gwarancja pozytywnej atestacji biegłego rewidenta' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 1 i ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 2 i 4 · C-281/12 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta „Raportowanie ESG zgodnie z CSRD & ESRS”, lead: „Kompleksowe wdrożenie unijnej dyrektywy CSRD oraz standardów ESRS (E, S, G) z gwarancją pozytywnej atestacji biegłego rewidenta.”',
            'Ta sama karta, ostatni punkt: „Wsparcie podczas audytu niezależnego biegłego rewidenta.”'
        ) `
        @'
Czytelnik karty CSRD wynosi, że zakup usługi kończy się pozytywną atestacją biegłego rewidenta. Ta sama karta nazywa biegłego niezależnym i rolę instytutu opisuje jako wsparcie podczas jego audytu. W leadzie nie ma roku obrotowego, zakresu atestacji ani zdania, że opinię wydaje biegły, a nie autor raportu. Nieprawdziwa informacja o spodziewanym wyniku usługi jest wprowadzeniem w błąd. Gwarancja opinii i słowo „niezależnego” stoją na jednej karcie, więc wynik nie jest w dyspozycji ESG Institute. Okno „Dowiedz się więcej” powtarza oba zdania i nie dopisuje zakresu.
'@ `
        'Z leadu karty CSRD usunąć człon „z gwarancją pozytywnej atestacji biegłego rewidenta”. Zostawić przygotowanie danych i odpowiedzi na pytania biegłego. Na stronie nie dopisywać lekcji o tym, czego instytut nie gwarantuje. Słowo „audyt luk” przy analizie luki może zostać: dotyczy przeglądu danych, nie opinii biegłego.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Wdrożenie raportowania zrównoważonego rozwoju według dyrektywy CSRD i standardów ESRS: ocena podwójnej istotności, luka w danych i procesy zbierania danych. Przy atestacji przygotowujemy odpowiedzi na pytania niezależnego biegłego rewidenta.”'

    Add-Case 'PRZYPADEK 02: „Ponad 1000 punktów danych ESRS”' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · UZNK art. 14 ust. 2 pkt 2 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta CSRD: „Identyfikacja i mapowanie ponad 1000 punktów danych ESRS (Data Points).”',
            'Ta sama karta, punkt wyżej: „Podwójna istotność (Double Materiality Assessment) z udziałem interesariuszy.”',
            'Sąsiednia karta VSME, odznaka „Nowość Omnibus 2026”: „Dopasowanie do wymogów pakietu upraszczającego UE Omnibus I.”'
        ) `
        @'
Czytelnik wynosi, że usługa CSRD mapuje ponad 1000 punktów danych. Na tej samej karcie punkt o podwójnej istotności mówi o odsiewie tematów, a nie o wpisaniu pełnej listy do raportu. Liczba nie ma roku, nazwy zestawu ESRS ani informacji, czy jest katalogiem do przesiewu, czy liczbą punktów w raporcie tej firmy. Prawdziwy rozmiar katalogu wolno podać tylko tak, by nie czytał się jako zakres raportu każdego klienta. Odznaka sąsiedniej karty mówi o pakiecie upraszczającym w 2026 r. Karta CSRD nie wskazuje, którego zestawu dotyczy „ponad 1000” i czy po tym pakiecie ta liczba nadal opisuje katalog.
'@ `
        'Z karty CSRD zdjąć absolut „ponad 1000”, dopóki w tym samym punkcie nie ma roku, nazwy zestawu i słowa, że do raportu wchodzą punkty zostawione po podwójnej istotności. Nie zastępować tej liczby inną liczbą, której sekcja nie podaje. Punkt o podwójnej istotności zostaje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Po ocenie podwójnej istotności mapujemy punkty danych ESRS, które wchodzą do raportu tej jednostki za rok [RRRR], według zestawu ESRS obowiązującego tę jednostkę.”'

    Add-Case 'PRZYPADEK 03: Medale EcoVadis Bronze, Silver lub Gold' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 2 · C-281/12 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta „Przygotowanie do EcoVadis & Audyt Dojrzałości”, lead: „Eksperckie przygotowanie organizacji do oceny w międzynarodowym ratingu EcoVadis w celu uzyskania medali Bronze, Silver lub Gold.”',
            'Ta sama karta: „Zarządzanie procesem zgłoszenia i optymalizacja punktacji.”',
            'Ta sama karta: „Benchmarking na tle międzynarodowych konkurentów w branży.”'
        ) `
        @'
Czytelnik karty wynosi, że celem usługi jest medal Bronze, Silver albo Gold. Medalu nie przyznaje ESG Institute. Karta nie podaje skali, progu medalu ani roku, w którym próg obowiązywał. „Optymalizacja punktacji” mówi o wyniku kwestionariusza, a nie o tym, który wskaźnik firmy ma się zmienić. Zapowiedź nagrody, której warunków nie ma przy zdaniu, jest wprowadzeniem w błąd co do wyniku usługi. Punkt o procedurach, politykach i KPI może zostać: opisuje dokument, nie medal. Benchmarking jest nazwany jako czynność, bez kryterium i bez grupy porównawczej. To nie jest reklama porównawcza, bo karta nie zestawia usługi instytutu z usługą innego doradcy.
'@ `
        'Z leadu usunąć „w celu uzyskania medali Bronze, Silver lub Gold”. Z listy usunąć „optymalizacja punktacji”. Jeśli benchmarking zostaje, w tym samym punkcie wpisać rok, kryterium i grupę firm. Liczby wobec konkurentów nie publikować, dopóki te trzy elementy nie stoją w tym samym zdaniu. Na karcie nie dopisywać wyjaśnienia, kto przyznaje medal.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Przegląd dokumentacji przed złożeniem kwestionariusza EcoVadis za rok [RRRR], w filarach środowisko, praca i prawa człowieka, etyka oraz zrównoważone zakupy.”'

    Add-Case 'PRZYPADEK 04: Ochrona kontraktów i preferencyjne finansowanie' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 2 · C-281/12 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta „Raportowanie ESG dla MŚP (Standard VSME)”, lead: „Pragmatyczne wdrożenie dobrowolnego standardu VSME dla małych i średnich firm będących kluczowymi ogniwami w łańcuchach dostaw.”',
            'Ta sama karta: „Ochrona kontraktów i preferencyjnych warunków finansowania.”',
            'Ta sama karta: „Gotowe szablony odpowiedzi na zapytania i ankiety od korporacji i banków.”'
        ) `
        @'
Czytelnik wynosi, że dobrowolny standard VSME zabezpieczy kontrakty i da preferencyjne warunki finansowania. Karta sama nazywa standard dobrowolnym. Nie wskazuje banku, produktu finansowania, roku ani klauzuli, z której miałaby wynikać ochrona umowy. Korzyść umowna i finansowa bez tych danych jest zapowiedzią wyniku, którego standard nie daje. Szablon odpowiedzi stoi obok tej obietnicy. Karta nie mówi, że każda odpowiedź ma pochodzić z danych tej firmy za nazwany rok. Okno „Dowiedz się więcej” tego nie uzupełnia.
'@ `
        'Z karty VSME usunąć punkt „Ochrona kontraktów i preferencyjnych warunków finansowania”. Przy szablonach dopisać w tym samym punkcie, że odpowiedź spisuje się z danych tej firmy za nazwany rok. Nie publikować zdania o preferencyjnej cenie kredytu, marży ani o utrzymaniu kontraktu. Punkty o Omnibusie i o wskaźnikach bazowych zostają, o ile przy wskaźniku jest rok.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Dobrowolny raport VSME tej firmy za rok [RRRR]: wskaźniki bazowe i odpowiedzi na pytania kontrahenta spisane z danych tej firmy za ten rok.”'

    Add-Case 'PRZYPADEK 05: „Precyzyjny” bilans organizacji i produktów oraz „pełne” 15 kategorii' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 1 i 2 oraz art. 6 · UZNK art. 14 ust. 2 pkt 2 · C-281/12 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta „Kalkulacja Śladu Węglowego (Scope 1, 2, 3)”, lead: „Precyzyjny bilans gazów cieplarnianych organizacji i produktów zgodny z GHG Protocol Corporate Standard oraz ISO 14064.”',
            'Ta sama karta: „Scope 3: Pełny łańcuch wartości – 15 kategorii upstream i downstream.”',
            'Ta sama karta: „Zautomatyzowane kalkulacje w partnerstwie z platformą Green0meter.”',
            'Ta sama karta: „Scope 2: Emisje pośrednie z zakupionej energii elektrycznej i ciepła (market i location-based).”'
        ) `
        @'
Czytelnik wynosi, że jeden precyzyjny bilans obejmuje organizację i produkty oraz cały łańcuch w piętnastu kategoriach. GHG Protocol Corporate Standard i ISO 14064-1 służą inwentaryzacji organizacji. Ślad produktu jest innym obiektem. Sąsiednia karta liczy cykl życia według ISO 14040 i ISO 14044, a liczba śladu produktu wymaga metody ISO 14067. Słowa „precyzyjny” i „pełny” nie mają roku, granic, udziału danych pierwotnych ani źródła wskaźników. Kategoria zakresu 3 wchodzi do bilansu, gdy jest istotna, a nie dlatego, że lista ma piętnaście pozycji. Nazwa Green0meter nie podaje bazy wskaźników ani tego, czy wynik ktoś zweryfikował. Zdanie o zakresie 2 w ujęciu market-based i location-based zostaje: rozdziela kupioną energię tak, jak wymaga tego standard organizacji. Metoda nazwana na karcie nie obejmuje obu obiektów, które lead stawia obok siebie.
'@ `
        'Z leadu usunąć „precyzyjny” oraz połączenie „organizacji i produktów” pod Corporate Standard i ISO 14064. Z punktu o zakresie 3 usunąć słowo „pełny”. Green0meter może zostać jako nazwa narzędzia tylko obok źródła wskaźników i udziału danych pierwotnych, w tym samym punkcie. Ślad produktu przenieść na kartę LCA albo opisać osobnym zdaniem z ISO 14067, granicą i jednostką. Notatka pod sekcją nie poprawia leadu, który sam mówi o precyzji i o pełnym łańcuchu. Punkt o zakresie 2 zostawić i przy przyszłej liczbie dopisać rok.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Inwentaryzacja gazów cieplarnianych organizacji za rok [RRRR]: zakres 1 i 2, metoda GHG Protocol Corporate Standard i ISO 14064-1, ze źródłem wskaźników i udziałem danych pierwotnych. Zakres 3 obejmuje kategorie uznane za istotne w tym roku. Ślad produktu [nazwa] liczymy osobno metodą ISO 14067, z granicą [od–do] i jednostką [jednostka].”'

    Add-Case 'PRZYPADEK 06: „Kompleksowa” ocena cyklu życia i granica do bramy' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · ISO 14040 i ISO 14044 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta „Cykl Życia Produktu (LCA) & Cyfrowe Paszporty (DPP)”, lead: „Kompleksowa ocena środowiskowa wyrobów (LCA)”.',
            'Ta sama karta: „Ocena cyklu życia Cradle-to-Grave / Cradle-to-Gate wg ISO 14040/14044.”'
        ) `
        @'
Czytelnik wynosi, że ocena wyrobu jest kompleksowa. Drugie zdanie stawia za ukośnikiem granicę do końca życia i granicę do bramy zakładu. Granica do bramy nie obejmuje użytkowania ani końca życia. ISO 14040 i ISO 14044 wymagają jednej, nazwanej granicy badania. Ukośnik czyta się jak jeden zakres. Słowo „kompleksowa” przy tym ukośniku obejmuje także badanie, które kończy się na bramie. Prawdziwa informacja o dwóch wariantach, podana tak, że różnica znika, wprowadza w błąd. Karta nie podaje wyrobu, jednostki ani roku danych.
'@ `
        'Z leadu usunąć słowo „kompleksowa”. Ukośnik „Cradle-to-Grave / Cradle-to-Gate” zastąpić jedną granicą, słowami, przy nazwie wyrobu. Drugiej granicy nie stawiać w tym samym zdaniu. Jeśli badanie dochodzi do końca życia, lead ma tak powiedzieć i nie dokładać słowa „kompleksowa”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Ocena cyklu życia wyrobu [nazwa] według ISO 14040 i ISO 14044. Granica: od pozyskania surowców do bramy zakładu [nazwa]. Jednostka: [jednostka]. Rok danych: [RRRR].”'

    Add-Case 'PRZYPADEK 07: „W pełnej zgodzie z unijnym AI Act”' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 4 | Poziom ryzyka: WYSOKIE' `
        @(
            'Karta „Sztuczna Inteligencja w ESG & Unijny AI Act”, lead: „Wdrażanie agentów AI i zautomatyzowanego przetwarzania danych niefinansowych w pełnej zgodzie z unijnym AI Act.”',
            'Ta sama karta: „Audyt systemów AI pod kątem klasyfikacji ryzyka (AI Act z 2 sierpnia 2026 r.).”'
        ) `
        @'
Czytelnik leadu wynosi, że agenci i przetwarzanie danych są w pełnej zgodzie z AI Act. Punkt niżej mówi o audycie klasyfikacji ryzyka i o dacie 2 sierpnia 2026 r. Pełna zgodność z całym aktem jest szersza niż klasyfikacja ryzyka. Data stoi w nawiasie. W dniu audytu, 21 października 2026 r., ta data już minęła, a karta nie wymienia obowiązków, które już działają, ani tych, które dotyczą danego systemu. Zapowiedź wyniku prawnego szerszego niż czynność z tej samej karty wprowadza w błąd. Punkt o szkoleniu z promptowania i o RODO tego przypadku nie dotyczy.
'@ `
        'Z leadu usunąć „w pełnej zgodzie”. Zostawić przegląd klasy ryzyka albo szkolenie, z datą przeglądu w tym samym zdaniu. Nawiasu „z 2 sierpnia 2026 r.” nie używać jako skrótu całego aktu. Nie publikować zdania, że system klienta jest zgodny, dopóki przegląd nie nazywa systemu i klasy.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Przegląd systemów AI klienta pod kątem klasy ryzyka w AI Act. Stan obowiązków: dzień przeglądu [data]. System objęty przeglądem: [nazwa].”'

    Add-Case 'PRZYPADEK 08: Deklaracja EPD bez programu i modułów' `
        'Podstawa prawna: UPNPR art. 6 · EN 15804 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Karta LCA, punkt listy: „Deklaracje Środowiskowe Produktu (EPD).”'
        ) `
        @'
Czytelnik listy wynosi, że efektem usługi jest deklaracja środowiskowa produktu. Karta nie podaje wyrobu, modułów, jednostki, roku ani programu, w którym deklaracja jest weryfikowana. Deklaracja EPD według EN 15804 jest nośnikiem danych cyklu życia i nie sięga dalej niż jej moduły. Samo słowo „EPD” zostawia wrażenie gotowej deklaracji. Brak programu i modułów zmienia znaczenie tej pozycji przy decyzji o zakupie.
'@ `
        'Punkt „Deklaracje Środowiskowe Produktu (EPD)” zdjąć albo rozwinąć w tym samym punkcie o wyrób, program operatora, moduły, jednostkę i rok. Nie stawiać słowa EPD obok hasła o całym cyklu życia, jeśli moduły kończą się na etapie A1–A3. Na karcie nie dopisywać zdania, że instytut nie jest operatorem programu. Nazwa programu w tym samym punkcie ten zakres już niesie.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Deklaracja środowiskowa wyrobu [nazwa], program [nazwa operatora], moduły [lista], jednostka [jednostka], rok danych [RRRR], po weryfikacji w tym programie.”'

    Add-Case 'PRZYPADEK 09: „Zgodnie z normami ESPR”' `
        'Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 · rozporządzenie (UE) 2024/1781 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Karta LCA, lead: „przygotowanie do wdrożenia Cyfrowych Paszportów Produktu zgodnie z normami ESPR.”',
            'Ta sama karta: „Projektowanie architektury danych dla Cyfrowych Paszportów Produktu (DPP).”'
        ) `
        @'
Czytelnik wynosi, że paszport powstaje według „norm ESPR”, jak według gotowej normy dla wyrobów w ogóle. ESPR to rozporządzenie (UE) 2024/1781. Obowiązek paszportu dochodzi grupą wyrobów, aktem dla tej grupy, a nie jedną normą dla całego katalogu. Karta nie nazywa grupy. Słowo „przygotowanie” uczciwie mówi o czasie. Słowo „normy” i brak grupy sprawiają, że przygotowanie wygląda jak zgodność już wymagana od wyrobu klienta.
'@ `
        'Z leadu usunąć „zgodnie z normami ESPR”. Jeśli dla grupy wyrobu nie ma aktu, który wprowadza paszport, nie publikować zdania o zgodności. Zostawić projekt struktury danych i wpisać grupę oraz numer aktu dopiero wtedy, gdy akt tę grupę obejmuje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Architektura danych cyfrowego paszportu dla grupy wyrobów [nazwa], według rozporządzenia (UE) 2024/1781 i aktu [numer], który tę grupę obejmuje.”'

    Add-Case 'PRZYPADEK 10: „Minimalizacja śladu węglowego modeli”' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Karta AI: „Zabezpieczenie prywatności danych (RODO) i minimalizacja śladu węglowego modeli.”'
        ) `
        @'
Czytelnik wynosi, że usługa zmniejsza ślad węglowy modeli. Zdanie nie ma roku bazowego, zużycia energii, lokalizacji serwerów ani wskaźnika. RODO i ślad stoją w jednym punkcie, choć to dwa różne wyniki. Twierdzenie o mniejszym śladzie wymaga liczby, roku i granicy w tym samym miejscu. Rzeczownik „minimalizacja” tej liczby nie zastępuje. Zdanie o prywatności danych nie jest twierdzeniem środowiskowym. Zostaje, o ile nie służy za dowód śladu.
'@ `
        'Rozdzielić punkt na dwa. Połowę o śladzie zdjąć, dopóki w tym samym zdaniu nie ma roku, kWh, lokalizacji serwerów i nazwy wskaźnika. Połowę o RODO zostawić osobno, bez słowa o śladzie.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Zużycie energii modeli użytych w projekcie w [RRRR] r.: [X] kWh, lokalizacja serwerów [kraj lub region], wskaźnik location-based [nazwa źródła].”'

    Add-Case 'PRZYPADEK 11: Dekarbonizacja i redukcja kosztów operacyjnych' `
        'Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 · UZNK art. 14 ust. 2 pkt 2 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Karta śladu węglowego, ostatni punkt: „Strategia dekarbonizacji i redukcji kosztów operacyjnych.”'
        ) `
        @'
Czytelnik wynosi, że strategia dekarbonizacji obniża koszty operacyjne. Karta nie wskazuje pozycji kosztu, roku bazowego ani warunku, przy którym koszt spada. Spadek emisji i spadek kosztu to dwa wyniki. Złączenie ich w jednym punkcie, bez liczby, przedstawia korzyść finansową jako cechę usługi śladu. Zapowiedź korzyści bez pozycji kosztu i bez roku jest wprowadzeniem w błąd co do rezultatu.
'@ `
        'Punkt rozbić. Strategia redukcji emisji może zostać jako dokument, z zakresem i rokiem bazowym w tym samym zdaniu. Zdanie o niższym koszcie publikować tylko przy nazwanej pozycji, roku bazowym i kwocie. Nie doklejać kosztu do zdania o emisji.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Przygotowanie strategii redukcji emisji organizacji w zakresie 1 i 2, z rokiem bazowym [RRRR] i z zasobem przypisanym do każdego celu w dokumencie strategii.”'

    Add-Case 'PRZYPADEK 12: „Zespół prawników, audytorów i inżynierów środowiska”' `
        'Podstawa prawna: UZNK art. 14 ust. 3 pkt 1 · UPNPR art. 5 ust. 3 pkt 6 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Karta „Abonament Eksperckiego Wsparcia ESG”: „Dedykowany zespół prawników, audytorów i inżynierów środowiska.”'
        ) `
        @'
Czytelnik abonamentu wynosi, że dyżur prowadzi zespół prawników, audytorów i inżynierów środowiska. Karta nie podaje imion ani uprawnień. Na karcie CSRD biegły rewident jest osobą niezależną od instytutu, więc słowo „audytorów” przy tym zestawieniu czyta się jak uprawnienie do atestacji. W sekcji Zespół tej samej witryny są radca prawny Rafał Wiatr i prawniczka Paula Guła. Nie ma tam osoby opisanej jako biegły rewident, audytor albo inżynier środowiska. Dr Beata Latos jest opisana jako doktor nauk przyrodniczych. Tytuły na karcie abonamentu nie pokrywają się z listą zespołu. Nieścisła informacja o kwalifikacjach osób wykonujących usługę wprowadza w błąd co do statusu usługodawcy.
'@ `
        'Z karty abonamentu usunąć „audytorów i inżynierów środowiska”. Na karcie zostawić imiona i role osób, które realnie prowadzą dyżur, w brzmieniu już używanym w sekcji Zespół. Radca prawny Rafał Wiatr może zostać, jeśli jest w dyżurze. Nie wpisywać tytułu, którego sekcja Zespół tej osobie nie daje. Punkt o przeglądzie kampanii może zostać, bez słowa „audyt”, albo ze słowem „przegląd”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Abonament: konsultacje osób wskazanych w umowie. Na tej karcie: Rafał Wiatr, radca prawny; [imię i rola kolejnej osoby, która prowadzi dyżur].”'

    Add-H '4. Ryzyka prawne i finansowe'
    Add-P '• 1. Prezes UOKiK i decyzja o zapytaniu' 11 $true $false $NAVY 10 2
    Add-P 'Strona jest publiczna. Wejście w kartę i w „Dowiedz się więcej” jest decyzją dotyczącą umowy, także gdy rozmowa jeszcze się nie zaczęła (C-281/12, art. 2 pkt 7 UPNPR). Praktyka naruszająca zbiorowe interesy konsumentów jest zagrożona karą do 10% obrotu za każdą praktykę (art. 24 i art. 106 ustawy o ochronie konkurencji i konsumentów). Odbiorcą tych kart jest przede wszystkim przedsiębiorca. Wobec niego bezpośrednia ścieżka to czyn nieuczciwej konkurencji, opisany w punkcie 2. Ścieżka konsumencka zostaje otwarta tam, gdzie przekaz czyta osoba działająca poza działalnością gospodarczą. Zdania o gwarancji atestacji, medalu i „precyzyjnym” bilansie są na tyle samodzielne, że pracownik albo agencja może zgłosić je do Prezesa UOKiK.' 10.5 $false $false $DARK 0 8
    Add-P '• 2. Roszczenie kontrahenta i konkurenta' 11 $true $false $NAVY 8 2
    Add-P 'UZNK art. 14 zakazuje nieprawdziwej albo wprowadzającej w błąd wiadomości o świadczonej usłudze, o sytuacji prawnej przedsiębiorcy i o kwalifikacjach pracowników. Art. 18 daje roszczenie o zaniechanie, usunięcie skutków, naprawienie szkody, wydanie korzyści i oświadczenie. Klient, który kupił kartę CSRD, czyta „gwarancję pozytywnej atestacji”. Klient karty EcoVadis czyta medal jako cel. Konkurent, który takiej gwarancji nie składa, ma interes w żądaniu zaniechania tych zdań.' 10.5 $false $false $DARK 0 8
    Add-P '• 3. Raport klienta i atestacja' 11 $true $false $NAVY 8 2
    Add-P 'Klient, który powtórzy w swoim raporcie „ponad 1000 punktów”, „pełny” zakres 3 albo deklarację EPD bez modułów, przenosi wadę sekcji do sprawozdania i do pytań biegłego rewidenta. Ślad produktu policzony standardem organizacji nie jest liczbą, którą da się wstawić przy wyrobie. Karta abonamentu sprzedaje w tym samym miejscu przegląd oświadczeń marketingowych. Odbiorca tego przeglądu porówna jego wynik ze zdaniami na kartach usług.' 10.5 $false $false $DARK 0 8
    Add-P '• 4. Wynik, którego umowa nie przenosi' 11 $true $false $NAVY 8 2
    Add-P 'Medal, preferencyjne finansowanie i pozytywna opinia biegłego nie powstają z mocy umowy z doradcą. Lead, który te wyniki nazywa, może być czytany jako ich przyrzeczenie. Korekta leadu przed ofertą i przed wzorem umowy ogranicza spór o to, co usługa miała dostarczyć. Karta ciągłości działania i rozróżnienie market-based oraz location-based tego ryzyka nie tworzą.' 10.5 $false $false $DARK 0 8

    Add-H '5. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIASTOWE ZDJĘCIE ZDAŃ (21 października 2026 r.)' @(
        'Z karty CSRD zdjąć „z gwarancją pozytywnej atestacji biegłego rewidenta” oraz „ponad 1000 punktów danych”.',
        'Z karty EcoVadis zdjąć medale Bronze, Silver i Gold oraz „optymalizacja punktacji”. Z karty VSME zdjąć „ochrona kontraktów i preferencyjnych warunków finansowania”.',
        'Z karty śladu zdjąć „precyzyjny”, „organizacji i produktów” pod jednym standardem, „pełny” przy piętnastu kategoriach oraz „redukcji kosztów operacyjnych”. Z karty LCA zdjąć „kompleksowa”, ukośnik grave/gate, „normami ESPR” i samo słowo „EPD”.',
        'Z karty AI zdjąć „w pełnej zgodzie” oraz „minimalizacja śladu węglowego modeli”. Z abonamentu zdjąć „audytorów i inżynierów środowiska”.'
    ) (Rgb 255 241 242) $DANGER

    Add-Callout 'ETAP 2: NOWE LEADY KART (do 28 października 2026 r.)' @(
        'Wstawić wzory z przypadków 01–12. Rok, granica, zestaw ESRS, grupa wyrobów i imię osoby dyżuru stoją w tym samym zdaniu co czynność.',
        'Zostawić bez zmiany: podwójną istotność, wsparcie przy pytaniach biegłego, zakres 2 market-based i location-based, procedury i KPI przed kwestionariuszem EcoVadis, kartę ciągłości działania oraz punkt o RODO po odłączeniu go od śladu modeli.',
        'Przy szablonie odpowiedzi VSME dopisać, że treść pochodzi z danych tej firmy za nazwany rok.'
    ) $AMBERBG $WARNING

    Add-Callout 'ETAP 3: METODYKA W TEJ SAMEJ SEKCJI (do 20 listopada 2026 r.)' @(
        'Przy usłudze śladu opublikować w sekcji usług, nie tylko w pliku do pobrania: rok, źródło wskaźników, udział danych pierwotnych i listę kategorii zakresu 3 uznanych za istotne. Ślad produktu ma osobne zdanie z ISO 14067.',
        'Przy EPD nazwać program operatora i moduły. Przy paszporcie nazwać grupę wyrobów i akt. Przy EcoVadis nie publikować progu medalu, którego karta nie ma.',
        'Notatka pod kartami nie poprawia leadu, który sam obiecuje medal, opinię biegłego albo precyzyjny pełny bilans (UPNPR art. 5, C-611/14).'
    ) (Rgb 239 246 255) $ACCENT

    Add-Callout 'ETAP 4: SŁOWNIK USŁUG (od 20 listopada 2026 r.)' @(
        'Każda nowa karta przechodzi cztery pytania: czy wynik wydaje ktoś inny niż instytut, czy standard pasuje do obiektu (organizacja albo produkt), czy liczba ma rok, czy granica jest jedna.',
        'Słownik prowadzi zespół, który na karcie abonamentu oferuje przegląd oświadczeń marketingowych. Zdanie zdjęte w etapie 1 nie wraca w ofercie, w ofercie PDF ani w zaproszeniu na rozmowę.',
        'Sekcja Zespół i karta abonamentu używają tych samych tytułów zawodowych. Nowy tytuł na karcie pojawia się razem z imieniem.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '6. Wnioski'
    Add-P 'Sekcja „Nasze usługi” w stanie na 21 października 2026 r. może zostać na stronie po zdjęciu gwarancji wyniku i po zawężeniu metody do obiektu, który nazwany standard obejmuje. W obecnym brzmieniu leady kart CSRD, EcoVadis, VSME, śladu węglowego, LCA i AI Act nie nadają się do publikacji. W tej sekcji nie ma ogólnego hasła „eko” ani zdania o neutralności opartej na kompensacji.' 10.5 $false $false $DARK 0 8

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
