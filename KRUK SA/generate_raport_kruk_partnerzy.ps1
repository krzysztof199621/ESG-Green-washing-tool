#Requires -Version 5.1
# Raport audytu strony partnerów biznesowych KRUK.
# Forma: INSTRUKCJA_GENEROWANIA_RAPORTOW.md. Dowód: Sprawozdanie ZR 2025.
$ErrorActionPreference = 'Stop'
$finalPath = Join-Path $PSScriptRoot 'Raport_Audytu_Greenwashingu_KRUK_partnerzy_biznesowi.docx'
$outPath = Join-Path $env:TEMP 'Raport_Audytu_Greenwashingu_KRUK_partnerzy_biznesowi.docx'

function Rgb([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }

$NAVY    = Rgb 8 24 31
$ACCENT  = Rgb 42 138 176
$DANGER  = Rgb 190 18 60
$WARNING = Rgb 180 83 9
$SUCCESS = Rgb 4 120 87
$DARK    = Rgb 22 36 44
$MUTED   = Rgb 90 107 119
$WHITE   = Rgb 255 255 255

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

    function Set-RunFont($range, [double]$size, [bool]$bold, [bool]$italic, [int]$color) {
        $range.Font.Name = 'Arial'
        $range.Font.Size = $size
        $range.Font.Bold = $(if ($bold) { 1 } else { 0 })
        $range.Font.Italic = $(if ($italic) { 1 } else { 0 })
        $range.Font.Color = $color
        $range.ParagraphFormat.LineSpacingRule = 0
    }

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
        [System.IO.File]::AppendAllText((Join-Path $PSScriptRoot 'build.log'), ((Get-Date).ToString('HH:mm:ss') + ' ' + $text + "`r`n"))
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

    function Shade-Cell($cell, [int]$color) {
        $cell.Shading.BackgroundPatternColor = $color
    }

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

    # Header
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
    Write-Cell $ht.Cell(1, 1) 'ESG INSTITUTE' 11 $true $WHITE
    $ht.Cell(1, 1).Range.ParagraphFormat.Alignment = 0
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

    # Footer
    $footer = $doc.Sections.Item(1).Footers.Item(1)
    $footer.Range.Text = 'ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Grupa Kapitałowa KRUK'
    $footer.Range.Font.Name = 'Arial'
    $footer.Range.Font.Size = 8
    $footer.Range.Font.Color = $MUTED
    $footer.Range.ParagraphFormat.Alignment = 0

    $word.ActiveWindow.ActivePane.View.SeekView = 0
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    Add-P 'ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO' 9 $true $false $ACCENT 0 2
    Add-P 'Audyt Ryzyk Greenwashingu: KRUK S.A.' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja twierdzeń ze strony https://pl.kruk.eu/partnerzy-biznesowi (stan na 13 października 2026 r.) oraz trzech stron oferty otwieranych z jej kafelków, zestawionych ze Skonsolidowanym sprawozdaniem zrównoważonego rozwoju Grupy Kapitałowej KRUK za 2025 r. Ocena wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE, UPNPR, ustawy o ochronie konkurencji i konsumentów oraz ustawy o zwalczaniu nieuczciwej konkurencji.' 10.5 $false $true $MUTED 0 8

    $shot = Join-Path $PSScriptRoot 'zrzut_partnerzy_biznesowi.png'
    if (Test-Path -LiteralPath $shot) {
        $sel.ParagraphFormat.Alignment = 1
        $shape = $sel.InlineShapes.AddPicture($shot)
        $shape.LockAspectRatio = $true
        $shape.Width = $word.CentimetersToPoints(15.4)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.Alignment = 0
    }
    Add-P 'Zrzut strony pl.kruk.eu/partnerzy-biznesowi z 13 października 2026 r. Widoczne hasło banera: „Najlepszy partner dla Twojej firmy” oraz kafelek „Serwis wierzytelności”.' 8.5 $false $true $MUTED 2 10

    $meta = @(
        @('Badany podmiot', 'KRUK S.A., Grupa Kapitałowa KRUK, ul. Bolkowska 3, 53-612 Wrocław (adres ze stopki badanej strony). NIP 894-23-89-605, KRS 0000240829.'),
        @('Przedmiot audytu', 'Strona „Partnerzy biznesowi” oraz trzy strony oferty, do których prowadzą kafelki: serwis wierzytelności, zakupy wierzytelności, sprzedaż wierzytelności i nieruchomości. Stan zebrania: 13 października 2026 r. Sprawozdanie zrównoważonego rozwoju za 2025 r. jest źródłem dowodowym dla twierdzeń tej strony, nie osobnym przedmiotem hasłowym.'),
        @('Sektor działalności', 'Zarządzanie wierzytelnościami: serwis na zlecenie (monitoring, windykacja polubowna i sądowo-egzekucyjna), zakup portfeli oraz sprzedaż wierzytelności i nieruchomości. Odbiorca strony: partner biznesowy (bank, ubezpieczyciel, inny wierzyciel). Strona jest publiczna.'),
        @('Reżim prawny', 'ECGT (dyrektywa (UE) 2024/825), stosowanie od 27.09.2026; druk Senatu 825 przewiduje transpozycję 27.09.2027 i nie jest ogłoszony. Do tego czasu te same wzorce ocenia się na art. 5 i 6 UPNPR. Dalej: dyrektywa 2005/29/WE, art. 24 i 106 ustawy o ochronie konkurencji i konsumentów, ustawa o zwalczaniu nieuczciwej konkurencji art. 10 i 16 (strona jest kierowana do przedsiębiorców i zawiera hasła wyższości), Kodeks pracy przy social washingu nagród pracowniczych, ustawa o ochronie sygnalistów przy twierdzeniu o transparentności. Nie stosowano SFDR ani progu nazw funduszy ESMA: strona nie nazywa produktu inwestycyjnego funduszem ESG.'),
        @('Data audytu', '13 października 2026 r.'),
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
        if ($last) { Shade-Cell $cv (Rgb 255 241 242) } else { Shade-Cell $cv $WHITE }
        Pad-Cell $ck 3 3 6 6
        Pad-Cell $cv 3 3 6 6
        Write-Cell $ck $meta[$i][0] 8.5 $true $DARK
        Write-Cell $cv $meta[$i][1] 8.5 $last $(if ($last) { $DANGER } else { $DARK })
    }
    Move-AfterTable $mt
    Add-P '' 6 $false $false $DARK 0 6

    Add-Callout 'WERDYKT: NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)' @(
        'Strona spełnia przesłanki czarnej listy załącznika I do dyrektywy 2005/29/WE w brzmieniu ECGT: pkt 10a (etyczna obsługa i „pozytywne relacje” z klientem partnera jako cecha wyróżniająca, choć rzetelne traktowanie konsumenta-dłużnika jest obowiązkiem prawa) oraz pkt 2a (grafika „Srebrny listek” bez wystawcy, roku i regulaminu w tym samym kanale; liść jest dorozumianym znakiem zrównoważoności).',
        'Osobno banery, których tekst nie ma atrybutu alt, głoszą „Lider rynku zarządzania wierzytelnościami w Europie Środkowej” i „Najlepszy partner dla Twojej firmy”. To twierdzenia wyższości bez kryterium, roku i źródła (UZNK art. 16, UPNPR art. 5). Sprawozdanie za 2025 r. nie podaje udziału w rynku. Wymienia cztery rynki kluczowe (Polska, Rumunia, Włochy, Hiszpania) i stwierdza, że spółka czeska była w 2025 r. w trakcie wygaszania działalności.',
        'Lista „Zaufali nam” obejmuje Wonga.pl sp. z o.o., czyli spółkę Grupy oferującą pożyczki konsumenckie, postawioną w jednym rzędzie z bankami zewnętrznymi. Zakazy ECGT stosuje się od 27 września 2026 r. Opóźnienie transpozycji nie zawiesza art. 5 i 6 UPNPR ani art. 24 i 106 ustawy o ochronie konkurencji i konsumentów.'
    ) (Rgb 255 241 242) $DANGER

    Add-H '1. Analiza Wstępna'
    Add-P 'KRUK S.A. prowadzi pod adresem pl.kruk.eu/partnerzy-biznesowi ofertę dla wierzycieli instytucjonalnych: serwis portfeli masowych, zakup wierzytelności oraz sprzedaż wierzytelności i nieruchomości. Stopka podaje siedzibę przy ul. Bolkowskiej 3 we Wrocławiu. Strona jest publiczna. Jej odbiorcą deklarowanym jest partner biznesowy, ale te same zdania czyta także konsument. Banery nie mają tekstu alternatywnego. Hasła „Lider rynku” i „Najlepszy partner” są wpisane w pliki graficzne.' 10.5 $false $false $DARK 0 8
    Add-P 'Narracja strony ustawia Grupę jako wyróżniającego się, etycznego i bezpiecznego partnera, któremu „zaufali” wymienieni przedsiębiorcy i który pokazuje pasek nagród, w tym grafikę liścia. Trzy kafelki prowadzą do stron, na których ta sama obietnica się zawęża albo rozszerza: serwis obejmuje pozew i komornika, zakup wymienia rynki czeski, słowacki i niemiecki, a sprzedaż „gwarantuje najwyższy poziom bezpieczeństwa i transparentności na każdym etapie”.' 10.5 $false $false $DARK 0 8
    Add-P 'Sprawozdanie zrównoważonego rozwoju za 2025 r. nie potwierdza tych skrótów. Opisuje cztery rynki kluczowe i wygaszanie spółki czeskiej, podaje Wonga jako spółkę Grupy, a nie jako zewnętrznego klienta, oraz ogranicza Certyfikat Etyczny ZPF do KRUK S.A. za 2024 r. W środowisku wykazuje 0% zgodności taksonomicznej, 81,67% energii ze źródeł kopalnych i brak planu przejścia. W społeczeństwie podaje niewykonanie celu zatrudnienia osób z niepełnosprawnościami (2,3% wobec co najmniej 4%) i nie ujawnia luki nieskorygowanej, choć na stronie nagród stoją znaki pracodawcy i jakości. Atestacja sprawozdania nie przenosi się na baner.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'
    $findings = @(
        @('01','„Lider rynku… w Europie Środkowej”; „Najlepszy partner dla Twojej firmy”','UZNK art. 16 · UPNPR art. 5','BARDZO WYSOKIE','Wyższość bez kryterium, roku i udziału w rynku'),
        @('02','„Wyróżnia nas” etyczna obsługa i pewność pozytywnych relacji z klientem','ECGT zał. I pkt 10a · UPNPR art. 5 i 6','WYSOKIE','Obowiązek wobec konsumenta jako cecha wyróżniająca; pominięty etap komorniczy'),
        @('03','„Zaufali nam”, w tym Wonga.pl sp. z o.o.','UPNPR art. 5 i 7 · UZNK art. 10','WYSOKIE','Spółka Grupy w rzędzie referencji zewnętrznych'),
        @('04','Grafika „Srebrny listek” w pasku nagród','ECGT zał. I pkt 2a · UPNPR art. 5','WYSOKIE','Liść bez wystawcy i regulaminu; sprawozdanie nie zna tego znaku'),
        @('05','„Gwarantują bezpieczeństwo oraz transparentność”, „najwyższy poziom… na każdym etapie”','UPNPR art. 5 i 6 · ustawa o sygnalistach art. 3','WYSOKIE','Gwarancja bez zakresu; sprawozdanie nie gwarantuje i zawęża kanał zgłoszeń'),
        @('06','Dream Employer, LGBT+ Diamond Awards, Gwiazda Jakości, Laur Klienta','UPNPR art. 5 i 6','ŚREDNIE','Nagrody bez roku i populacji, przeniesione na ofertę windykacji'),
        @('07','Pozycja budowana na rynkach czeskim, słowackim i od 2015 r. niemieckim','UPNPR art. 5 i 6','ŚREDNIE','Sprawozdanie: cztery rynki kluczowe i wygaszanie spółki czeskiej')
    )
    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $st = $doc.Tables.Add($range, (1 + $findings.Count), 5)
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
                if ($val -eq 'ŚREDNIE') { $colr = $WARNING }
                elseif ($val -eq 'NISKIE') { $colr = $SUCCESS }
                else { $colr = $DANGER }
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

    Add-Case 'PRZYPADEK 01: „Lider rynku” i „Najlepszy partner” bez kryterium' `
        'Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji art. 16 · UPNPR art. 5 | Poziom ryzyka: BARDZO WYSOKIE' `
        @(
            'Baner lider_rynku_baner_1600x500_1.png, bez tekstu alternatywnego: „Lider rynku zarządzania wierzytelnościami w Europie Środkowej”.',
            'Baner partnerzy_biznes_baner_1600x500_3.png, bez tekstu alternatywnego: „Najlepszy partner dla Twojej firmy”.',
            'Kafelek zakupów: „Ponad 17-letnie doświadczenie i wyjątkowy profesjonalizm w tym zakresie, stanowią znak rozpoznawczy Grupy KRUK na rynku europejskim.”'
        ) `
        'Artykuł 16 ustawy o zwalczaniu nieuczciwej konkurencji dopuszcza reklamę porównawczą, gdy jest obiektywna, sprawdzalna i dotyczy istotnych, typowych i weryfikowalnych cech. „Lider” i „najlepszy” są twierdzeniami wyższości wobec reszty rynku. Strona nie podaje kryterium (przychód, wartość portfeli, liczba spraw, kraj), roku pomiaru, źródła ani mianownika. Pusty atrybut alt nie usuwa twierdzenia wobec odbiorcy, który widzi baner. Sprawozdanie za 2025 r. nie zawiera udziału w rynku ani rankingu. Podaje cztery rynki kluczowe: Polska, Rumunia, Włochy i Hiszpania, oraz aktywa we Francji i w Czechach. Włochy i Hiszpania nie są w tym zdaniu „Europą Środkową”, a spółka czeska była w 2025 r. w trakcie wygaszania działalności operacyjnej. „Znak rozpoznawczy” i „wyjątkowy profesjonalizm” powtarzają wyższość bez cechy, którą da się sprawdzić. Siedemnaście lat doświadczenia, jeśli jest faktem historycznym, nie dowodzi pozycji lidera.' `
        'Zdjąć oba banery albo zastąpić je zdaniem bez stopnia wyższego. Nie używać „lider”, „najlepszy”, „wyjątkowy” i „znak rozpoznawczy”, dopóki w tym samym widoku nie ma kryterium, daty, terytorium i źródła. Tekst hasła umieścić w HTML, nie tylko w pikselach pliku graficznego.' `
        'SFORMUŁOWANIE NIE DO URATOWANIA' `
        '„Lider rynku” i „Najlepszy partner” nie mają bezpiecznej wersji skrótowej. Po korekcie zostaje opis usługi i rynków ze sprawozdania: Polska, Rumunia, Włochy i Hiszpania jako rynki kluczowe w 2025 r. Bez przymiotnika stopnia wyższego.'

    Add-Case 'PRZYPADEK 02: Etyczna obsługa jako cecha wyróżniająca' `
        'Podstawa prawna: ECGT, załącznik I pkt 10a · UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            '„Wyróżnia nas podejście pro-klienckie, możesz więc mieć pewność, że podjęte przez nas działania będą etyczne i przełożą się na utrzymanie pozytywnych relacji z Twoimi Klientami.” (kafelek Serwis wierzytelności).',
            'Strona serwisu, do której kafelek prowadzi: „w przypadku braku efektów - złożenia pozwu w imieniu Partnera Biznesowego i uzyskania tytułu wykonawczego w celu skierowania wniosku do Komornika Sądowego”. Oraz: „Polubowne metody odzyskania należności zawiodły? Pomożemy Ci przejść przez trudny etap sądowo-egzekucyjny z poszanowaniem Twojego Klienta.”',
            'Sprawozdanie za 2025 r.: „KRUK S.A. uzyskał Certyfikat Etyczny ZPF za rok 2024.” W 2025 r. „nie odnotowała poważnych incydentów dotyczących poszanowania praw człowieka związanych z klientami”. „W 2026 roku Grupa otrzymała zgłoszenie 1 przypadku naruszenia praw człowieka - w jego sprawie trwa postępowanie wyjaśniające.”'
        ) `
        'Punkt 10a zakazuje przedstawiania uprawnienia konsumenta wynikającego z prawa jako cechy wyróżniającej ofertę. Rzetelne traktowanie dłużnika-konsumenta, w tym zakaz praktyk agresywnych i wprowadzających w błąd, wynika z UPNPR i z nadzoru Prezesa UOKiK. Zdanie „wyróżnia nas” sprzedaje ten obowiązek jako przewagę KRUK. „Możesz więc mieć pewność” jest gwarancją skutku (etyczność i pozytywna relacja), której sprawozdanie nie składa: opisuje certyfikat ZPF za 2024 r. dla KRUK S.A., nie dla całej Grupy i nie za 2026 r., oraz odnotowuje zgłoszenie naruszenia praw człowieka z 2026 r. w toku. Kafelek na stronie zbiorczej nie mówi, że serwis obejmuje pozew i komornika. Ta informacja jest o jedno kliknięcie dalej. Dopisek na podstronie nie sanuje hasła na stronie, z której partner wchodzi. To art. 6: pominięcie etapu, który zmienia znaczenie „pozytywnych relacji”.' `
        'Na kafelku skreślić „wyróżnia nas”, „pewność” i „pozytywnych relacji”. Jeśli zostaje zdanie o etyce, w tym samym akapicie podać nazwę certyfikatu, rok, spółkę i to, że serwis obejmuje także etap sądowo-egzekucyjny. Nie przenosić certyfikatu ZPF za 2024 r. na całą Grupę i na rok 2026.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Serwis portfeli masowych KRUK S.A. obejmuje monitoring, windykację polubowną oraz etap sądowo-egzekucyjny, w tym pozew i wniosek do komornika. KRUK S.A. posiada Certyfikat Etyczny ZPF za 2024 r. Certyfikat nie dotyczy pozostałych spółek Grupy i nie jest gwarancją pozytywnej relacji z każdą osobą zadłużoną.”'

    Add-Case 'PRZYPADEK 03: Wonga.pl na liście „Zaufali nam”' `
        'Podstawa prawna: UPNPR art. 5 i art. 7 · ustawa o zwalczaniu nieuczciwej konkurencji art. 10 | Poziom ryzyka: WYSOKIE' `
        @(
            'Nagłówek „ZAUFALI NAM” oraz, w tym samym bloku, „Wonga.pl sp. z o.o.” obok m.in. Santander Bank Polska S.A., Bank Pekao S.A., PKO Leasing S.A. i BNP Paribas Bank Polska S.A.',
            'Sprawozdanie za 2025 r.: „Grupa KRUK świadczy również usługi pożyczek konsumenckich. Spółka Wonga działa na otwartym rynku pożyczek konsumenckich w Polsce.”'
        ) `
        'Lista pod nagłówkiem „Zaufali nam” jest zestawem referencji. Odbiorca czyta ją jako podmioty spoza Grupy, które powierzyły KRUK wierzytelności. Sprawozdanie zalicza Wonga do Grupy i opisuje ją jako własną działalność pożyczkową, nie jako klienta zewnętrznego. Postawienie spółki zależnej w jednym rzędzie z bankami przypisuje ofercie cudze zaufanie tam, gdzie jest relacja wewnątrz grupy. Artykuł 5 UPNPR obejmuje informację prawdziwą (Wonga istnieje i może korzystać z procesów Grupy) podaną tak, że wprowadza w błąd co do niezależności referencji. Artykuł 7 obejmuje posługiwanie się oznaczeniem zaufania w sposób, który sugeruje rekomendację osoby trzeciej. Artykuł 10 ustawy o zwalczaniu nieuczciwej konkurencji obejmuje oznaczenie usługi mogące wprowadzić klienta w błąd co do tego, kto za usługą stoi. Audyt nie kwestionuje pozostałych nazw na liście. Kwestionuje brak informacji, że jedna z nich jest spółką Grupy.' `
        'Usunąć Wonga.pl z listy „Zaufali nam” albo oznaczyć ją w tym samym wierszu jako spółkę Grupy Kapitałowej KRUK. Nie liczyć jej do liczby podmiotów zewnętrznych, jeśli taka liczba pojawi się w materiale.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Na liście odniesień są podmioty, które sprzedały portfel albo zleciły serwis, w zakresie wynikającym z umowy z danym podmiotem. Wonga.pl sp. z o.o. jest spółką Grupy Kapitałowej KRUK i nie jest referencją zewnętrzną.”'

    Add-Case 'PRZYPADEK 04: „Srebrny listek” bez systemu certyfikacji' `
        'Podstawa prawna: ECGT, załącznik I pkt 2a · UPNPR art. 5 · praktyka Prezesa UOKiK (Tchibo, 19.01.2026): liść przy haśle może być dorozumianym twierdzeniem albo znakiem zrównoważoności | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'W pasku „NAGRODY” grafika listek_srebrny.jpg z jedynym opisem „Srebrny listek”. Na stronie nie ma wystawcy, roku, regulaminu ani zdania, czego znak dotyczy.',
            'Sprawozdanie za 2025 r. nie wymienia tego znaku. Podaje natomiast: brak planu przejścia na 31.12.2025, udział źródeł kopalnych w energii 81,67% oraz wskaźniki obrotu, CapEx i OpEx zgodne z taksonomią równe 0%.'
        ) `
        'Punkt 2a zakazuje znaku zrównoważoności, który nie pochodzi z systemu certyfikacji strony trzeciej albo od organu publicznego, jeżeli system nie jest pokazany. Liść jest w bazie wiedzy przykładem znaku, który UOKiK czyta jako claim albo oznaczenie zrównoważoności, gdy stoi przy przekazie o firmie. Tutaj stoi w pasku nagród oferty windykacyjnej, bez nazwy systemu. Odbiorca nie jest w stanie odróżnić rankingu, nagrody prasowej i certyfikatu środowiskowego. Sprawozdanie, które miałoby być dowodem profilu, tego znaku nie objaśnia, a profil środowiskowy, który objaśnia, to 0% taksonomii i energia w większości kopalna. Znak nie może być czytany jako efektywność ekologiczna serwisu wierzytelności. Nie przypisuję liściowi konkretnego rankingu, którego strona nie nazywa.' `
        'Usunąć grafikę z paska, dopóki w tym samym widoku nie ma wystawcy, roku, podstawy prawnej albo regulaminu oraz zdania, że znak nie opisuje środowiskowej cechy usługi windykacyjnej ani zgodności z taksonomią. Nie stawiać liścia obok „etyczne” i „zaufali nam”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„[Nazwa wyróżnienia], przyznane przez [wystawca] w [rok], według regulaminu [adres]. Wyróżnienie nie jest certyfikatem środowiskowym usługi zarządzania wierzytelnościami i nie oznacza zgodności z taksonomią UE. W 2025 r. wskaźniki taksonomiczne Grupy wyniosły 0%.”'

    Add-Case 'PRZYPADEK 05: Gwarancja bezpieczeństwa i transparentności na każdym etapie' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · ustawa z 14.06.2024 r. o ochronie sygnalistów, art. 3 | Poziom ryzyka: WYSOKIE' `
        @(
            'Kafelek sprzedaży: „Doświadczenie i wiedza ekspertów gwarantują bezpieczeństwo oraz transparentność w procesie sprzedaży aktywów.”',
            'Strona sprzedaży: „Nasze wieloletnie doświadczenie oraz specjalistyczna wiedza ekspertów gwarantują najwyższy poziom bezpieczeństwa i transparentności na każdym etapie procesu inwestycyjnego.”',
            'Sprawozdanie, G1-1: procedura zgłoszeń wewnętrznych kanałem whistleblowing obejmuje w tabeli KRUK S.A., KRUK Romania, KRUK Italia, KRUK España i InvestCapital. Polityka praw człowieka zapowiada „pełne bezpieczeństwo” sygnalistów dla Grupy.'
        ) `
        '„Gwarantują” i „najwyższy poziom… na każdym etapie” są obietnicą skutku, nie opisem procedury. Strona nie wskazuje, jakie ryzyko jest objęte, jaki etap, jaka spółka i jaki dokument to mierzy. Sprawozdanie za ten sam okres nie używa słowa gwarancja wobec sprzedaży aktywów. Przy zgłoszeniach, które są narzędziem transparentności, tabela procedury jest węższa niż zdanie o Grupie: nie wymienia w tym wierszu Wonga, Novum, KRUK TFI ani Kancelarii Prawnej RAVEN. Ustawa o sygnalistach chroni zgłoszenie naruszenia prawa z katalogu art. 3 ust. 1, a nie każdą wątpliwość etyczną (art. 3 ust. 2). Hasło „transparentność na każdym etapie” nie podaje tego rozróżnienia. To art. 5 (obietnica szersza niż dowód) i art. 6 (pominięcie zakresu). Sprzeczności liczbowe w części klimatycznej sprawozdania, opisane w osobnym audycie tego dokumentu, pokazują, że sam raport nie jest jeszcze tekstem, na który da się przenieść słowo „gwarantujemy”.' `
        'Na kafelku i na stronie sprzedaży zamienić „gwarantują” na opis czynności: jakie dokumenty dostaje nabywca, na jakim etapie, która spółka odpowiada. Listę spółek objętych procedurą zgłoszeń podać przy słowie transparentność albo tego słowa nie używać. Zdjąć „najwyższy” i „każdym etapie”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Przy sprzedaży wierzytelności lub nieruchomości przekazujemy nabywcy [lista dokumentów] przed złożeniem oferty. Procedura zgłoszeń wewnętrznych obowiązuje w KRUK S.A., KRUK Romania, KRUK Italia, KRUK España i InvestCapital. Nie jest gwarancją wyniku transakcji ani objęciem każdej spółki Grupy.”'

    Add-Case 'PRZYPADEK 06: Pasek nagród bez roku i bez związku z usługą' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE (transfer atrybutów)' `
        @(
            'Pasek „NAGRODY”, opisy grafik: „Dream Employer”, „LGBT+ Diamond Awards”, „Konsumencki Lider Jakości”, „Gwiazda Jakości 2024”, „Laur Klienta”, „Gwiazda Jakości Obsługi 2023”, „Złoty Laur Konsumenta 2017”, „Marka Godna Zaufania 2017”, „Giełdowa Spółka Roku 2016”. Część grafik w tym samym pasku nie ma tekstu alternatywnego.',
            'Sprawozdanie, cele pracownicze na koniec 2025 r.: zatrudnienie osób z niepełnosprawnościami 2,3% wobec celu co najmniej 4%. Cel „100% menedżerów” przeszkolonych z różnorodności rozliczono zdaniem o 76% pracowników. Cel „0% różnicy” w wynagrodzeniach: skorygowana luka 0,1%, a miernika S1-16 nie ujawniono, bo równość wynagrodzeń uznano za nieistotną.'
        ) `
        'Artykuł 5 obejmuje wprowadzenie w błąd co do nagród i wyróżnień. Pasek stoi na stronie usługi windykacyjnej, między „zaufali nam” a ofertą serwisu. Nie podaje, która spółka dostała nagrodę, za jaki rok poza tymi nazwami, które rok mają w nazwie, ani czy nagroda dotyczy pracodawcy, marki konsumenckiej czy obsługi portfela. „Giełdowa Spółka Roku 2016” ma dekadę. Dream Employer i LGBT+ Diamond Awards są znakami pracodawcy. Sprawozdanie, które jest aktualnym opisem tego pracodawcy, pokazuje niewykonanie celu 4% zatrudnienia osób z niepełnosprawnościami i nie odpowiada na cel 100% menedżerów. Przeniesienie znaku pracodawcy i znaku jakości konsumenckiej na kafelek „działania będą etyczne” jest transferem atrybutu. Audyt nie stwierdza, że nagród nie przyznano. Stwierdza, że strona nie daje danych, bez których pasek zmienia znaczenie oferty.' `
        'Przy każdym znaku podać spółkę, rok, organizatora i przedmiot (pracodawca, marka, obsługa). Znaki starsze niż trzy lata przenieść do archiwum albo opatrzyć rokiem w tym samym widoku. Nie stawiać Dream Employer obok kafelka etycznej windykacji bez zdania, że znak dotyczy pracodawcy, nie usługi dla partnera. Puste alt uzupełnić albo grafikę zdjąć.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„[Nazwa nagrody], [organizator], [rok], przedmiot: [pracodawca / marka konsumencka / obsługa]. Nagroda nie opisuje serwisu wierzytelności ani zakupu portfela. W 2025 r. udział pracowników z orzeczeniem o niepełnosprawności wyniósł 2,3% wobec celu co najmniej 4%.”'

    Add-Case 'PRZYPADEK 07: Rynki czeski, słowacki i niemiecki jako aktualna pozycja' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona zakupów: „Swoją pozycję w tym zakresie Grupa konsekwentnie buduje już od kilkunastu lat - najpierw na rynku polskim, następnie rumuńskim, czeskim, słowackim, a od 2015r. również niemieckim, włoskim i hiszpańskim.”',
            'Sprawozdanie za 2025 r.: „Działa na czterech kluczowych dla Grupy rynkach w Europie – w Polsce, Rumunii, Włoszech i Hiszpanii, posiada również aktywa we Francji i Czechach.” „W 2025 roku spółka KRUK Česká a Slovenská republika w Czechach była w trakcie procesu wygaszania działalności operacyjnej.”'
        ) `
        'Zdanie o zakupach jest w czasie teraźniejszym: pozycja „buduje się” na liście rynków, która obejmuje Czechy, Słowację i Niemcy. Sprawozdanie za ten sam okres wymienia cztery rynki kluczowe i nie wymienia wśród nich Niemiec ani Słowacji. O Czechach mówi wprost, że spółka operacyjna wygaszała działalność. Historia wejścia na rynek może być prawdziwa i jednocześnie wprowadzać w błąd co do stanu na 2025 r., jeżeli nie ma zdania o wygaszeniu. To art. 5 (prawdziwy rys historyczny w trybie, który czytelnik bierze za mapę obecnej działalności) i art. 6 (pominięcie wygaszenia). Nie kwestionuję obecności we Włoszech i Hiszpanii: sprawozdanie je potwierdza jako rynki kluczowe.' `
        'Na stronie zakupów rozdzielić historię i stan obecny. Przy Czechach dopisać, że w 2025 r. działalność operacyjna spółki czeskiej była wygaszana. Niemcy i Słowację zostawić w zdaniu historycznym tylko z datą, albo usunąć, jeśli nie ma aktualnej działalności zakupowej. Nie łączyć tej listy z banerem „Europa Środkowa”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2025 r. rynki kluczowe Grupy to Polska, Rumunia, Włochy i Hiszpania. Spółka KRUK Česká a Slovenská republika w Czechach była w 2025 r. w trakcie wygaszania działalności operacyjnej. Wcześniejsze wejścia na inne rynki opisujemy z datą i nie oznaczają one bieżącej działalności zakupowej.”'

    Add-H '4. Porównanie twierdzeń strony ze sprawozdaniem za 2025 r.'
    Add-P 'Strona nie wymienia konkurenta z nazwy. Tabeli reklamy porównawczej z innym przedsiębiorcą nie sporządza się. Poniższe zestawienie konfrontuje twierdzenie strony z tym, co ta sama Grupa podała w sprawozdaniu zrównoważonego rozwoju za 2025 r. Tego żądał zakres audytu: sprawozdanie jest dowodem dla postulatów strony.' 10.5 $false $false $DARK 0 8

    $comp = @(
        @('„Lider rynku zarządzania wierzytelnościami w Europie Środkowej” i „Najlepszy partner dla Twojej firmy”.', 'Sprawozdanie nie podaje udziału w rynku ani rankingu. Rynki kluczowe 2025: Polska, Rumunia, Włochy, Hiszpania. Spółka czeska wygaszała działalność operacyjną.'),
        @('„Wyróżnia nas” etyczna obsługa i pewność pozytywnych relacji z klientem partnera.', 'Certyfikat Etyczny ZPF za 2024 r. dotyczy KRUK S.A. Serwis obejmuje pozew i komornika. W 2026 r. wpłynęło zgłoszenie naruszenia praw człowieka, w toku.'),
        @('„Zaufali nam”: Wonga.pl sp. z o.o. wśród banków.', 'Wonga jest spółką Grupy i prowadzi pożyczki konsumenckie na rynku otwartym. Nie jest referencją zewnętrzną.'),
        @('„Srebrny listek” w pasku nagród.', 'Sprawozdanie nie opisuje tego znaku. Taksonomia: 0% obrotu, CapEx i OpEx. Energia kopalna: 81,67%. Planu przejścia na 31.12.2025 nie było.'),
        @('„Gwarantują” najwyższą transparentność na każdym etapie sprzedaży.', 'Procedura whistleblowing w tabeli G1-1: KRUK S.A., Rumunia, Włochy, Hiszpania, InvestCapital. Sprawozdanie nie gwarantuje wyniku transakcji.'),
        @('Dream Employer i nagrody jakości przy ofercie windykacji.', 'Cel niepełnosprawności: 2,3% wobec co najmniej 4%. Cel 100% menedżerów DEI rozliczono jako 76% pracowników. S1-16 nieujawnione.')
    )
    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $ct = $doc.Tables.Add($range, (1 + $comp.Count), 2)
    $ct.PreferredWidthType = 3
    $ct.PreferredWidth = $contentW
    $ct.AllowAutoFit = $false
    $ct.Columns.Item(1).Width = $word.CentimetersToPoints(7.2)
    $ct.Columns.Item(2).Width = $word.CentimetersToPoints(10.4)
    Write-Cell $ct.Cell(1, 1) 'Twierdzenie na stronie' 8 $true $WHITE
    Write-Cell $ct.Cell(1, 2) 'Sprawozdanie zrównoważonego rozwoju za 2025 r.' 8 $true $WHITE
    Shade-Cell $ct.Cell(1, 1) $NAVY
    Shade-Cell $ct.Cell(1, 2) $NAVY
    Pad-Cell $ct.Cell(1, 1) 3 3 4 4
    Pad-Cell $ct.Cell(1, 2) 3 3 4 4
    for ($r = 0; $r -lt $comp.Count; $r++) {
        $bg = $(if ($r % 2 -eq 0) { $WHITE } else { (Rgb 248 250 252) })
        for ($c = 0; $c -lt 2; $c++) {
            $cell = $ct.Cell($r + 2, $c + 1)
            Shade-Cell $cell $bg
            Pad-Cell $cell 3 3 4 4
            Write-Cell $cell $comp[$r][$c] 8 $false $DARK
        }
    }
    Move-AfterTable $ct

    Add-H '5. Ryzyka prawne i finansowe'
    Add-P '• 1. Prezes UOKiK' 12 $true $false $NAVY 8 2
    Add-P 'Strona jest publiczna, a kafelek serwisu mówi o klientach partnera, czyli o konsumentach zadłużonych. Artykuł 24 ustawy o ochronie konkurencji i konsumentów zakazuje praktyk naruszających zbiorowe interesy konsumentów, w tym informacji, która nie jest rzetelna, prawdziwa i pełna. Artykuł 106 ust. 1 pkt 4 pozwala nałożyć karę do 10% obrotu, także nieumyślnie. Od 27 września 2026 r. pkt 10a i 2a są zakazane w każdych okolicznościach. Hasło „pewność, że działania będą etyczne” i liść bez regulaminu są tym wzorcem.' 10.5 $false $false $DARK 0 8
    Add-P '• 2. Konkurent i kontrahent' 12 $true $false $NAVY 8 2
    Add-P '„Lider rynku” i „Najlepszy partner” są reklamą wyższości wobec pozostałych firm zarządzających wierzytelnościami. Artykuł 16 ustawy o zwalczaniu nieuczciwej konkurencji wymaga, by takie porównanie dało się sprawdzić. Artykuł 18 daje konkurentowi roszczenie o zaniechanie i o naprawienie szkody. Bank, który czyta „Zaufali nam” łącznie z Wonga, dostaje referencję grupy jako referencję rynku.' 10.5 $false $false $DARK 0 8
    Add-P '• 3. Sprawozdanie, które partner wpisze do własnego ESRS' 12 $true $false $NAVY 8 2
    Add-P 'Partner raportujący łańcuch wartości może przenieść „lidera”, „etyczną pewność” albo liść do własnego opisu dostawcy windykacji. Sprawozdanie KRUK za 2025 r. tych skrótów nie niesie: ma 0% taksonomii, wygaszenie w Czechach i certyfikat ZPF ograniczony do jednej spółki i jednego roku. Rozbieżność między stroną a sprawozdaniem jest ryzykiem dla obu stron, nie potwierdzeniem strony.' 10.5 $false $false $DARK 0 8
    Add-P '• 4. Sygnalista' 12 $true $false $NAVY 8 2
    Add-P 'Ustawa o ochronie sygnalistów obejmuje naruszenia prawa z zakresu ochrony konsumentów i ochrony środowiska. Pracownik, który zestawi baner „lider” albo liść z tabelą rynków i taksonomii w sprawozdaniu, ma podstawę do zgłoszenia. Obietnica transparentności na stronie nie zastępuje kanału ustawowego i nie może być węższa niż art. 3.' 10.5 $false $false $DARK 0 8

    Add-H '6. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIAST (13 października 2026 r.)' @(
        'Wyłączyć oba banery („Lider rynku”, „Najlepszy partner”) oraz grafikę „Srebrny listek”. Na kafelku serwisu skreślić „wyróżnia nas”, „pewność” i „pozytywnych relacji”. Z listy „Zaufali nam” wyjąć Wonga.pl albo oznaczyć ją jako spółkę Grupy. Nie używać tych skrótów w prezentacji dla banku do czasu korekty.'
    ) (Rgb 255 241 242) $DANGER
    Add-Callout 'ETAP 2: KOREKTA STRONY (do 20 października 2026 r.)' @(
        'Na kafelku serwisu dopisać, że obsługa obejmuje etap sądowo-egzekucyjny. Przy etyce podać Certyfikat Etyczny ZPF: spółka KRUK S.A., rok 2024. Na stronie sprzedaży zamienić „gwarantują” i „najwyższy poziom na każdym etapie” na listę dokumentów i listę spółek objętych procedurą zgłoszeń. Na stronie zakupów rozdzielić historię rynków i stan 2025, w tym wygaszenie w Czechach.'
    ) (Rgb 255 251 235) $WARNING
    Add-Callout 'ETAP 3: NAGRODY I ZNAKI (do 12 listopada 2026 r.)' @(
        'Każdej grafice w pasku nagród nadać tekst alternatywny oraz podpis: organizator, rok, spółka, przedmiot. Znaki sprzed więcej niż trzech lat przenieść do archiwum. Liść przywrócić tylko z wystawcą, rokiem i zdaniem, że nie jest cechą środowiskową usługi. Nie łączyć Dream Employer z kafelkiem windykacji.'
    ) (Rgb 240 249 255) $ACCENT
    Add-Callout 'ETAP 4: SAFE HARBOR (od 13 października 2026 r., tryb ciągły)' @(
        'Hasło na tej stronie może wejść tylko wtedy, gdy w tym samym widoku jest spółka, rok i źródło. Zakazane skróty: „lider”, „najlepszy”, „wyjątkowy”, „gwarantujemy”, „pewność”, „wyróżnia nas” przy obowiązku prawnym, liść bez regulaminu, spółka Grupy na liście zewnętrznego zaufania. Tekst banera ma być w HTML, nie tylko w pliku PNG. Raz na kwartał zestawić stronę partnerów ze sprawozdaniem albo z komentarzem kwartalnym i zdjąć zdanie, którego sprawozdanie nie niesie.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '7. Wnioski'
    Add-P 'Strona partnerów biznesowych KRUK w stanie na 13 października 2026 r. nie może zostać w obecnej formie: banery głoszą lidera i najlepszego partnera bez kryterium, kafelek serwisu sprzedaje obowiązek etyczny jako wyróżnik, a lista zaufania miesza banki ze spółką Grupy. Sprawozdanie za 2025 r. tych skrótów nie potwierdza. Do zdjęcia banerów, liścia i zdania o pewności werdykt pozostaje czerwony.' 10.5 $false $false $DARK 2 8

    try { $doc.BuiltInDocumentProperties.Item(1).Value = 'Audyt Ryzyk Greenwashingu: KRUK partnerzy biznesowi' } catch {}
    [System.IO.File]::AppendAllText((Join-Path $PSScriptRoot 'build.log'), ((Get-Date).ToString('HH:mm:ss') + " SAVE`r`n"))
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $name = [string]$outPath
    $doc.SaveAs2($name)
    $doc.Close($false)
    $doc = $null
    Copy-Item -LiteralPath $outPath -Destination $finalPath -Force
    Write-Output "SAVED $finalPath"
}
finally {
    if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
