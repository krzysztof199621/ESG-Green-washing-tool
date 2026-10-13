#Requires -Version 5.1
# Raport audytu greenwashingu: Grupa Kapitałowa KRUK (sprawozdanie ZR 2025).
# Forma: INSTRUKCJA_GENEROWANIA_RAPORTOW.md. Material: Sprawozdanie_ZR_Grupa_KRUK_2025.pdf.
$ErrorActionPreference = 'Stop'
$finalPath = Join-Path $PSScriptRoot 'Raport_Audytu_Greenwashingu_Grupa_KRUK_2025.docx'
$outPath = Join-Path $env:TEMP 'Raport_Audytu_Greenwashingu_Grupa_KRUK_2025.docx'

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
    Add-P 'Audyt Ryzyk Greenwashingu: GRUPA KAPITAŁOWA KRUK' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja twierdzeń środowiskowych i społecznych w Skonsolidowanym sprawozdaniu zrównoważonego rozwoju Grupy Kapitałowej KRUK za okres od 1 stycznia 2025 r. do 31 grudnia 2025 r. (plik Sprawozdanie_ZR_Grupa_KRUK_2025.pdf) wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE, ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, ustawy o ochronie konkurencji i konsumentów, ustawy o ochronie sygnalistów oraz Kodeksu pracy w zakresie social washingu. Ujawnienie taksonomiczne oceniono wobec rozporządzenia (UE) 2020/852. Certyfikat budynku oceniono wobec dyrektywy (UE) 2024/1275.' 10.5 $false $true $MUTED 0 8

    $meta = @(
        @('Badany podmiot', 'KRUK S.A., jednostka dominująca Grupy Kapitałowej KRUK. Sprawozdanie nie podaje adresu siedziby. Wskazuje biura KRUK S.A. we Wrocławiu, Szczawnie-Zdroju i Pile oraz rynki: Polska, Rumunia, Włochy i Hiszpania, aktywa we Francji i w Czechach, a zatrudnienie także na Malcie, w Czechach i w Luksemburgu (3 605 pracowników i 316 osób na innych umowach, stan na koniec 2025 r.).'),
        @('Przedmiot audytu', 'Skonsolidowane sprawozdanie zrównoważonego rozwoju za 2025 r., sporządzone według ustawy o rachunkowości i ESRS oraz art. 8 rozporządzenia (UE) 2020/852, z atestacją biegłego rewidenta w zakresie zgodności ujawnień ze standardami. Audyt obejmuje wyłącznie ten plik. Treści włączone tylko przez odesłanie do sprawozdania zarządu i sprawozdania finansowego nie były badane.'),
        @('Sektor działalności', 'Zarządzanie wierzytelnościami nabytymi na własny rachunek i na zlecenie oraz pożyczki konsumenckie (Wonga w Polsce, Novum w Polsce, RoCapital w Rumunii). Kluczowa grupa klientów: konsumenci – osoby zadłużone. Grupa oświadcza, że nie prowadzi działalności w sektorze paliw kopalnych, chemikaliów, kontrowersyjnych broni ani tytoniu.'),
        @('Reżim prawny', 'ECGT (dyrektywa (UE) 2024/825), stosowanie od 27.09.2026; krajowa nowela (druk Senatu 825) przewiduje 27.09.2027 i nie jest ogłoszona. Do tego czasu te same wzorce ocenia się na art. 5–6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. Dalej: dyrektywa 2005/29/WE, art. 24 i art. 106 ustawy o ochronie konkurencji i konsumentów, ustawa z 14.06.2024 r. o ochronie sygnalistów, Kodeks pracy przy social washingu (art. 18³c, art. 94 pkt 2b, art. 94³, art. 207), definicja greenwashingu ESA z 4.06.2024 r. dla komunikacji podmiotu z sektora finansowego. Taksonomia 2020/852 – jako sprawdzenie, czy ujawnienie myli kwalifikowalność ze zgodnością. EPBD (UE) 2024/1275 art. 11 – wyłącznie przy LEED. Nie stosowano SFDR ani progu nazw funduszy ESMA: plik nie kwalifikuje produktu inwestycyjnego. Nie stosowano CountEmissionsEU ani EN ISO 14083: nie ma kalkulatora emisji usługi transportowej. Nie stosowano art. 16 ustawy o zwalczaniu nieuczciwej konkurencji: nie ma reklamy porównawczej.'),
        @('Data audytu', '12 października 2026 r.'),
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
        'Sprawozdanie spełnia przesłanki czarnej listy załącznika I do dyrektywy 2005/29/WE w brzmieniu nadanym dyrektywą (UE) 2024/825: pkt 4a (ogólne „zrównoważone praktyki” i „działania proekologiczne” oraz „zrównoważony rozwój” jako etykieta compliance, etyki i cyberbezpieczeństwa, bez uznanej wysokiej efektywności ekologicznej), pkt 4b (minimalizacja wpływu całej działalności oraz „dostawy zielonej energii” i flota hybrydowa jako obraz klimatu Grupy, gdy dotyczą wycinka energii i floty) oraz pkt 10a (bezpieczne warunki pracy, zwalczanie mobbingu i „uczciwe wynagrodzenie” przedstawione jako cecha relacji z pracownikami, choć są obowiązkiem Kodeksu pracy).',
        'Te zakazy stosuje się od 27 września 2026 r. Opóźnienie transpozycji w Polsce nie zawiesza oceny według ECGT ani odpowiedzialności z art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. Osobno warstwa liczbowa jest wewnętrznie sprzeczna: narracja „70% redukcji do 2040 r.” nie zgadza się z tabelą celu (70% wartości bazowej, czyli redukcja o 30%), a „osiągnięta redukcja 2 283,70 t CO2” z certyfikatów pochodzenia nie zgadza się z inwentaryzacją zakresu 2 (1 244 tCO2e location-based i 374 tCO2e market-based). To nie jest twierdzenie o neutralności oparte na offsetach; pkt 4c nie został zastosowany, bo sprawozdanie nie twierdzi, że Grupa jest neutralna klimatycznie, i ujawnia brak usuwania gazów cieplarnianych.',
        'Część dokumentu jest ujawnieniem, nie claimem: brak planu przejścia (planowany na 2026 r.), 0% obrotu, CapEx i OpEx zgodnych z taksonomią, 81,67% energii ze źródeł kopalnych, wynagrodzenie zarządu bez kryteriów ESG oraz zmiana wpływu klimatycznego z pozytywnego na negatywny. Werdykt dotyczy narracji, która te ustalenia przykrywa hasłem redukcji, zielonej energii i zrównoważenia. Utrzymanie tych zdań w sprawozdaniu publikowanym i w materiałach, które je powielają, otwiera zarzut z art. 24 i 106 ustawy o ochronie konkurencji i konsumentów (kara do 10% obrotu) oraz drogę zgłoszenia sygnalisty.'
    ) (Rgb 255 241 242) $DANGER

    Add-H '1. Analiza Wstępna'
    Add-P 'KRUK S.A. jest jednostką dominującą Grupy Kapitałowej KRUK. Grupa zarządza wierzytelnościami nabytymi na własny rachunek i na zlecenie oraz udziela pożyczek konsumenckich pod markami Wonga, Novum i RoCapital. Na koniec 2025 r. zatrudniała 3 605 pracowników. Kluczową grupą klientów są konsumenci – osoby zadłużone. Sprawozdanie jest oświadczeniem zrównoważonego rozwoju za 2025 r., sporządzonym według ESRS i art. 8 rozporządzenia (UE) 2020/852, i podlega atestacji biegłego rewidenta co do zgodności ujawnień ze standardami. W styczniu 2025 r. opublikowano strategię biznesową na lata 2025–2029. Strategia ESG z 2023 r. była realizowana po raz ostatni. Nowa strategia zrównoważonego rozwoju ma zostać opublikowana w 2026 r. Za obszar ESG w zarządzie odpowiada członkini zarządu Urszula Okarma.' 10.5 $false $false $DARK 0 8
    Add-P 'Część tekstu da się obronić jako ujawnienie, nie jako obietnicę środowiskową. Rewizja oceny podwójnej istotności zmieniła wpływ klimatyczny z pozytywnego „Korporacyjne działania Grupy KRUK na rzecz klimatu” na negatywny „Wpływ na środowisko poprzez emisję gazów cieplarnianych” i zmniejszyła liczbę kluczowych wpływów z 20 do 14 oraz ryzyk z 6 do 2. Na 31 grudnia 2025 r. nie było planu przejścia; sporządzenie zaplanowano na 2026 r. Grupa nie usuwa gazów cieplarnianych we własnych operacjach, nie stosuje wewnętrznej ceny emisji i nie identyfikuje istotnych skutków finansowych ryzyk klimatycznych. Wskaźniki taksonomiczne obrotu, nakładów inwestycyjnych i wydatków operacyjnych zgodnych z systematyką wynoszą 0%. Koszyk energii jest podany wprost: 13 264,2187 MWh, w tym 81,67% ze źródeł kopalnych. Wynagrodzenie członków zarządu i rady nadzorczej nie zależy od wyników zrównoważonego rozwoju, a program motywacyjny opiera się wyłącznie na wskaźnikach finansowych. W celach społecznych sprawozdanie pokazuje także niewykonanie: zatrudnienie osób z niepełnosprawnościami 2,3% wobec celu co najmniej 4%.' 10.5 $false $false $DARK 0 8
    Add-P 'Narracja wychodzi poza ten zakres. Cel „70% redukcji” do 2040 r. nie ma planu przejścia ani planu działań i nie zgadza się z własną tabelą liczbową, a emisja zakresu 1 w 2025 r. jest równa bazie z 2021 r. Zakup gwarancji pochodzenia jest nazwany dostawą zielonej energii i „osiągniętą redukcją” 2 283,70 t, choć metoda rynkowa zakresu 2 wzrosła rok do roku o 98,93%. Kolumna zmiany całkowitych emisji pokazuje spadek o 20,76%, złożony niemal w całości z kategorii 2 i z załamania dojazdów do pracy z 1 881 tCO2e do 5 tCO2e. Hasła o zrównoważonych praktykach, działaniach proekologicznych i minimalizacji wpływu całej działalności nie mają pokrycia w EU Ecolabel ani w oznakowaniu ISO 14024 typu I. Obok tego sprawozdanie przedstawia obowiązki Kodeksu pracy – bezpieczne warunki pracy, przeciwdziałanie mobbingowi, uczciwe wynagrodzenie – jako element oferty wobec pracowników i polityki praw człowieka. Atestacja zgodności ze standardami sprawozdawczymi nie wyrównuje sprzeczności między tymi zdaniami a tabelami w tym samym pliku.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'

    $findings = @(
        @('01','Cel „70% redukcji CO2 do 2040 r.” wobec bazy 2021; tabela: 70% wartości bazowej; zakres 1 w 2025 r. = 2 262 tCO2e','UPNPR art. 5 i 6 · ECGT, przyszła efektywność (art. 6 dyrektywy 2005/29/WE)','BARDZO WYSOKIE','Trzy niezgodne odczyty tego samego celu; brak planu przejścia i planu działań'),
        @('02','„Dostawy zielonej energii” i osiągnięta redukcja 2 283,70 t z certyfikatów pochodzenia','ECGT zał. I pkt 4b · UPNPR art. 5 i 6 · wzorzec UOKiK (Bolt, EAC)','BARDZO WYSOKIE','Gwarancja pochodzenia podana jako redukcja większa niż cały zakres 2'),
        @('03','Całkowite emisje GHG −20,76% r/r (location-based)','UPNPR art. 5 i 6','WYSOKIE','Spadek prawie w całości z kategorii 2 i z dojazdów 1 881 → 5 tCO2e'),
        @('04','„Minimalizuje negatywny wpływ swojej działalności” na środowisko','ECGT zał. I pkt 4b · UPNPR art. 5 i 6','WYSOKIE','Claim o całej działalności przy śladzie zdominowanym przez zakres 3'),
        @('05','Flota: 5 hybryd i 1 hybryda plug-in w Polsce oraz 79 hybryd we Włoszech jako najważniejsze działanie','ECGT zał. I pkt 4b · UPNPR art. 6','WYSOKIE','42,86 t przy 9 405 MWh paliw ropopochodnych i bez udziału we flocie'),
        @('06','„Zrównoważone praktyki”, „działania proekologiczne”, kampania o „zrównoważonym rozwoju”','ECGT zał. I pkt 4a · UPNPR art. 5','WYSOKIE','Hasło ogólne bez EU Ecolabel i bez ISO 14024 typu I'),
        @('07','Uczciwe wynagrodzenie, bezpieczne warunki pracy, przeciwdziałanie mobbingowi','ECGT zał. I pkt 10a · UPNPR art. 5 · Kodeks pracy','WYSOKIE','Obowiązek ustawowy przedstawiony jako cecha relacji z pracownikami'),
        @('08','Cel „0% różnicy” w wynagrodzeniach; wykonanie: skorygowana luka 0,1%. Szkolenie DEI: cel 100% menedżerów, wynik 76% pracowników','UPNPR art. 5 i 6','ŚREDNIE','Metryka zawężona; ESRS S1-16 pominięty jako nieistotny'),
        @('09','Nowa siedziba „z myślą o zrównoważonym rozwoju”, certyfikat LEED v4 Platinum','UPNPR art. 5 · EPBD art. 11 jako granica znaczenia LEED','ŚREDNIE','Certyfikat budynku przeniesiony na „zrównoważony rozwój” siedziby'),
        @('10','„Pełne bezpieczeństwo” sygnalistów w całej Grupie','Ustawa o ochronie sygnalistów art. 3 · UPNPR art. 6','ŚREDNIE','Procedura węższa podmiotowo; ochrona ustawowa nie obejmuje samego standardu etycznego'),
        @('11','Kodeks: „zrównoważona, niskoemisyjna infrastruktura i transport”','UPNPR art. 5 · ECGT, przyszła efektywność','ŚREDNIE','Aspiracja kodeksu bez planu, przy 81,67% energii kopalnej')
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

    Add-Case 'PRZYPADEK 01: Cel „70% redukcji” do 2040 r. bez planu i w trzech niezgodnych odczytach' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · dyrektywa 2005/29/WE art. 6 w brzmieniu ECGT (przyszła efektywność środowiskowa) | Poziom ryzyka: BARDZO WYSOKIE' `
        @(
            '„Grupa dąży do osiągnięcia 70% redukcji emisji dwutlenku węgla do 2040 roku w odniesieniu do roku bazowego 2021.” (E1-4, s. 218).',
            'Tabela E1-4: zakres 1, wartość bazowa 2021 = 2 262, docelowy poziom 2040 = 1 583,4, „docelowy poziom emisji w % wartości bazowej” = 70%, „roczny cel redukcji wartości bazowej” = 3,68%, rzeczywiste emisje 2025 = 2 262. Dla zakresu 2 location-based: baza 1 326, cel 928,2, ten sam odsetek 70% i ta sama stopa 3,68%, emisja 2025 = 1 244.',
            '„Na dzień 31 grudnia 2025 roku Grupa KRUK nie opracowała Planu przejścia na potrzeby łagodzenia zmian klimatu. Sporządzenie planu przejścia planowane jest w roku 2026.” (E1-1, s. 217). „Grupa nie opracowała szczegółowych planów działań” (E1-3, s. 218).',
            '„Wynagrodzenie te nie jest uzależnione od wyników Spółki, w tym w zakresie zrównoważonego rozwoju.” Program motywacyjny „bazuje wyłącznie na wskaźnikach finansowych i nie odnosi się do kryteriów ESG.” (GOV-3, s. 164–165).'
        ) `
        'Twierdzenie o przyszłej efektywności środowiskowej jest dopuszczalne tylko z jasnym, obiektywnym, publicznym i weryfikowalnym planem: mierzalne cele, terminy, zasoby i regularna weryfikacja. Tego planu w sprawozdaniu nie ma. E1-1 mówi wprost, że planu przejścia nie opracowano, a E1-3 – że nie opracowano szczegółowych planów działań i dlatego nie ujawniono wydatków ani nakładów. Cel na 2040 r. nie jest więc ścieżką, tylko hasłem procentowym. Ten sam dokument podaje trzy niezgodne znaczenia procentu. Zdanie „70% redukcji” zostawia 30% bazy, czyli dla zakresu 1 około 679 tCO2e, nie 1 583,4. Liczba 1 583,4 to dokładnie 70% z 2 262, a więc redukcja o 30%, i tak samo 928,2 to 70% z 1 326. Stopa 3,68% wartości bazowej rocznie przez 19 lat od 2021 do 2040 daje około 70% redukcji i przeczy kolumnie poziomu docelowego. Czytelnik nie jest w stanie powiedzieć, czy Grupa zobowiązała się ściąć emisje o 70%, czy zejść do 70% bazy. Emisja zakresu 1 w 2025 r. jest równa bazie z 2021 r. (2 262 tCO2e). Cztery lata po roku bazowym nie ma redukcji zakresu 1 wobec bazy, niezależnie od tego, który z dwóch odczytów procentu przyjąć. Cel nie obejmuje zakresu 3. Wynagrodzenie organów nie jest powiązane z tym celem, a nakładów na działania nie ujawniono, bo planu działań nie ma. Prawdziwe elementy tabeli, pokazane obok zdania o 70% redukcji, wprowadzają w błąd co do skali zobowiązania (art. 5 – informacja prawdziwa w sposobie podania; art. 6 – pominięcie tego, który odczyt obowiązuje).' `
        'W E1-4 usunąć zdanie o „70% redukcji”, dopóki jedna liczba docelowa, jeden rok bazowy, jeden zakres (1, 2 location-based albo 2 market-based, osobno) i plan przejścia nie stoją w tym samym miejscu. Do publikacji planu w 2026 r. nie używać procentu 70% w prezentacjach dla inwestorów, w serwisie relacji inwestorskich ani w komentarzu kwartalnym. Tabelę albo przeliczyć do jednego znaczenia, albo opatrzyć zdaniem, który wiersz jest celem, a który jest błędem opisu. Nie łączyć tego celu ze zdaniem o odnawialnych źródłach energii, skoro działania opisane w E1-3 to gwarancje pochodzenia i dopisanie hybryd (przypadki 02 i 05).' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Na 31.12.2025 Grupa nie ma planu przejścia; plan jest zapowiedziany na 2026 r. i do jego publikacji Grupa nie podaje procentowego celu redukcji. Emisje zakresu 1: rok bazowy 2021 = 2 262 tCO2e, rok 2025 = 2 262 tCO2e. Zakres 2 według metody opartej na lokalizacji: 2021 = 1 326 tCO2e, 2025 = 1 244 tCO2e. Wynagrodzenie zarządu nie zależy od tych wielkości. Cel, jeśli zostanie przyjęty w 2026 r., wskaże w tym samym akapicie zakres, rok bazowy, poziom w tCO2e i nakłady.”'

    Add-Case 'PRZYPADEK 02: „Zielona energia” i 2 283,70 t redukcji z gwarancji pochodzenia' `
        'Podstawa prawna: ECGT, załącznik I pkt 4b · UPNPR art. 5 i art. 6 · praktyka Prezesa UOKiK wobec „100% energii odnawialnej” opartej na certyfikatach (sprawa Bolt, zarzuty z 19.01.2026) | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            '„Grupa KRUK zabezpieczyła dostawy zielonej energii dla swoich lokalizacji, finalizując zakup certyfikatów pochodzenia dla biur KRUK S.A. we Wrocławiu, Szczawnie-Zdroju i Pile, a także uzyskując analogiczne gwarancje dla spółek KRUK Romania (wraz z RoCapital, KRUK TECH i Corbul), KRUK España oraz KRUK Italia.” W tabeli E1-3 przy tej pozycji: osiągnięta redukcja emisji CO2 w 2025 r. = 2 283,70 t, spodziewana redukcja = 2 283,70 t (s. 218).',
            '„Udział źródeł kopalnych w całkowitym zużyciu energii (%)” 2025 = 81,67%; zużycie kopalne 10 833,2855 MWh; zużycie odnawialne 2 416,6907 MWh; zużycie całkowite 13 264,2187 MWh (E1-5, s. 220). Zakupiona energia objęta gwarancjami pochodzenia: 2 267,36 MWh. „Udział emisji Zakres 2 objętych instrumentami kontraktowymi” = 94,97% (E1-6, s. 224–225).',
            'Zakres 2 brutto, metoda oparta na lokalizacji: 1 305 tCO2e (2024) i 1 244 tCO2e (2025). Metoda rynkowa: 188 tCO2e (2024) i 374 tCO2e (2025), zmiana +98,93% (s. 223). „Metoda rynkowa (market-based) – odzwierciedla rzeczywiste decyzje zakupowe organizacji” (s. 222).'
        ) `
        'Gwarancja pochodzenia i inny certyfikat atrybutu energii nie są fizyczną dostawą prądu ze źródła odnawialnego. Prezes UOKiK w zarzutach wobec Bolt zakwestionował „100% odnawialnej energii”, gdy podstawą były certyfikaty EAC, a nie bezpośredni zakup energii. Sprawozdanie w jednym wierszu nazywa ten zakup „dostawami zielonej energii” i wpisuje 2 283,70 t jako redukcję osiągniętą w 2025 r., równą redukcji spodziewanej. Z inwentaryzacji tej tony nie da się odczytać. Cały zakres 2 według lokalizacji to 1 244 tCO2e, a więc liczba 2 283,70 jest od niego wyższa. Różnica między metodą lokalizacyjną a rynkową w 2025 r. wynosi 870 tCO2e, nie 2 283,70. Rok do roku zakres 2 location-based spadł o 61 tCO2e, a market-based wzrósł o 186 tCO2e. Zdanie, że metoda rynkowa odzwierciedla rzeczywiste decyzje zakupowe, stoi obok wzrostu tej emisji, nie obok redukcji 2 283,70 t. Udział 94,97% emisji zakresu 2 objętych instrumentami też nie schodzi się z relacją 374 do 1 244 (zostaje około 30% zakresu 2 liczonego lokalizacyjnie). Koszyk energii pozostaje w 81,67% kopalny, z czego same paliwa z ropy i produktów naftowych to 9 405,0452 MWh. „Zielona energia dla lokalizacji” opisuje atrybut umowny części zakupionej energii elektrycznej tak, jakby dotyczył dostaw energii Grupy. To pkt 4b: twierdzenie o energii działalności, podczas gdy dowód dotyczy instrumentu kontraktowego dla wskazanych biur i spółek. Pkt 4c nie zachodzi: nie ma tu zdania o neutralności klimatycznej ani o kompensacji offsetem, a E1-7 wyłącza usuwanie gazów cieplarnianych we własnych operacjach. Naruszeniem jest zrównanie umorzenia gwarancji z redukcją ton, której inwentaryzacja nie pokazuje.' `
        'W tabeli E1-3 wykreślić „dostawy zielonej energii” i kolumnę redukcji 2 283,70 t. W tym samym akapicie, w którym pada informacja o gwarancjach, podać wolumen w MWh, listę spółek, obie metody zakresu 2 i udział źródeł kopalnych w koszyku. Nie używać słowa „redukcja” dla różnicy między współczynnikiem residualnym a zerem przypisanym do gwarancji, dopóki ta różnica nie jest policzona i równa liczbie w tabeli emisji. Zdanie o tym, że metoda rynkowa odzwierciedla decyzje zakupowe, uzupełnić o zmianę +98,93% r/r.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2025 r. Grupa umorzyła gwarancje pochodzenia dla 2 267,36 MWh zakupionej energii (biura KRUK S.A. we Wrocławiu, Szczawnie-Zdroju i Pile oraz KRUK Romania z RoCapital, KRUK TECH i Corbul, KRUK España i KRUK Italia). Nie jest to pomiar energii fizycznej ze źródeł odnawialnych w tych lokalizacjach. Emisje zakresu 2: 1 244 tCO2e (lokalizacja) i 374 tCO2e (rynek, +98,93% r/r). Udział energii kopalnej w koszyku 13 264 MWh wyniósł 81,67%. Sprawozdanie nie przypisuje temu zakupowi redukcji w tonach CO2e.”'

    Add-Case 'PRZYPADEK 03: Spadek całkowitych emisji o 20,76% bez mostu zmian' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE' `
        @(
            'Całkowite emisje GHG, metoda oparta na lokalizacji: 48 253 tCO2e (2024) i 38 235 tCO2e (2025), −20,76%. Metoda rynkowa: 47 136 i 37 365 tCO2e, −20,72% (E1-6, s. 224).',
            'Zakres 3, kategoria 2 (dobra inwestycyjne): 10 117 tCO2e (2024) i 2 191 tCO2e (2025), −78,34%. Kategoria 7 (dojazdy pracowników): 1 881 tCO2e i 5 tCO2e, −99,73%. Kategoria 1: 30 873 i 30 709 tCO2e, −0,53%.',
            'Przy kategorii 7 poziom dokładności danych: „Niski”; komentarz: „Ankieta wśród pracowników” (s. 221). „Procent emisji GHG Zakres 3 obliczony przy użyciu danych podstawowych od dostawców lub innych partnerów w łańcuchu wartości” = 0% (s. 224). „Obliczenia śladu węglowego oraz dane dotyczące koszyka energetycznego nie były weryfikowane przez organ zewnętrzny” poza atestacją sprawozdania (s. 222 i BP-2).'
        ) `
        'Kolumna „% 2025 / 2024” jest twierdzeniem o zmianie śladu. Spadek metody lokalizacyjnej wynosi 10 018 tCO2e. Z tego kategoria 2 tłumaczy 7 926 tCO2e, a kategoria 7 – 1 876 tCO2e, łącznie około 98% spadku. Kategoria 1, największa pozycja zakresu 3 (30 709 tCO2e), prawie nie drgnęła. Zakres 1 spadł o 4,07% wobec 2024 r. (2 358 do 2 262) i wrócił do bazy z 2021 r. Sprawozdanie nie podaje mostu: ile z −20,76% jest zmianą działalności, ile zmianą nakładów inwestycyjnych, a ile zmianą ankiety dojazdowej. Pięć ton CO2e dojazdów przy 3 605 pracownikach to około 1,4 kg CO2e na osobę rocznie. Przy deklarowanej niskiej jakości danych ankietowych taki wynik wymaga zdania objaśniającego. Bez niego odbiorca bierze −20,76% za dekarbonizację operacji. Intensywność emisji jest podana z kilkunastoma miejscami po przecinku (0,0119834478919819 na przychód netto), przy zakresie 3 opartym w 0% na danych pierwotnych od dostawców i przy metodzie wydatkowej. Sama atestacja zgodności sprawozdania ze standardami, opisana w BP-1 i BP-2, nie zastępuje mostu zmian. To informacja prawdziwa w tabeli, podana tak, że pomija okoliczność decydującą o znaczeniu procentu (art. 5 i 6).' `
        'Przy tabeli E1-6 wstawić most: zmiana kategorii 2, zmiana kategorii 7 ze wskazaniem, że jakość danych jest niska i że 5 tCO2e nie jest porównywalne z 1 881 tCO2e bez opisu ankiety, oraz zmiana zakresu 1 i 2. Procentu −20,76% nie wynosić do streszczenia, slajdu ani listu zarządu. Intensywność zaokrąglić do poziomu, który odpowiada jakości danych zakresu 3. Jeśli kategoria 7 za 2025 r. jest błędem zasilenia, skorygować liczbę, a nie zostawiać −99,73% jako wynik.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Emisje GHG location-based: 48 253 tCO2e w 2024 r. i 38 235 tCO2e w 2025 r. Z różnicy 10 018 tCO2e około 7 926 tCO2e przypada na kategorię 2 (dobra inwestycyjne: 10 117 → 2 191), a około 1 876 tCO2e na kategorię 7 (dojazdy: 1 881 → 5, dane ankietowe, dokładność niska). Kategoria 1 wyniosła 30 709 tCO2e (−0,53%). Zakres 1: 2 262 tCO2e. Zakres 3 policzono w 0% z danych pierwotnych dostawców. Procent −20,76% nie jest miarą redukcji operacyjnej.”'

    Add-Case 'PRZYPADEK 04: Minimalizacja wpływu całej działalności przy śladzie zakresu 3' `
        'Podstawa prawna: ECGT, załącznik I pkt 4b · UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            '„Prawo do czystego i zdrowego środowiska – Grupa KRUK minimalizuje negatywny wpływ swojej działalności na otoczenie, respektując prawo społeczności do życia w bezpiecznym środowisku naturalnym.” (S3, zobowiązania Polityki Praw Człowieka, s. 252).',
            '„Ograniczenie negatywnego wpływu na środowisko jest dla Grupy ważnym aspektem działalności.” (E1-4, s. 218).',
            'Wpływ istotny E1: „Wpływ na środowisko poprzez emisję gazów cieplarnianych”, charakter negatywny, miejsce powstania: wyższy szczebel łańcucha wartości i operacje własne (SBM-3, s. 180). Emisje 2025: zakres 1 = 2 262 tCO2e, zakres 2 market-based = 374 tCO2e, zakres 3 = 34 729 tCO2e. Całość metodą rynkową = 37 365 tCO2e.'
        ) `
        'Punkt 4b zakazuje twierdzenia o całej działalności, gdy podstawa dotyczy tylko jej aspektu. Zdanie o minimalizowaniu negatywnego wpływu „swojej działalności” jest zdaniem o całości. Inwentaryzacja pokazuje co innego. Zakres 3 to 34 729 z 37 365 tCO2e metody rynkowej, czyli 92,9% śladu, a w tym kategoria 1 (zakupione towary i usługi, głównie metoda wydatkowa) to 30 709 tCO2e. Cel z E1-4 obejmuje tylko zakresy 1 i 2. Działania z E1-3, które sprawozdanie nazywa najważniejszymi, to dopisanie hybryd (42,86 t) i gwarancje pochodzenia. Grupa sama przeniosła opis wpływu klimatycznego z pozytywnego działania „na rzecz klimatu” na wpływ negatywny i umieściła go także w wyższym szczeblu łańcucha. Zdanie o minimalizacji tego nie uwzględnia. Czytelnik polityki praw człowieka dostaje cechę całej działalności, a czytelnik E1-6 – ślad, którego Grupa celem redukcyjnym nie obejmuje w części rozstrzygającej. Ukryty kompromis polega na chwaleniu operacji własnych przy pominięciu tego, że istotny wpływ, według własnej macierzy, powstaje w łańcuchu. To pkt 4b i art. 6.' `
        'W Polityce Praw Człowieka i w miejscu, w którym sprawozdanie ją streszcza, usunąć zdanie o minimalizowaniu wpływu całej działalności. Jeśli zdanie ma zostać, w tym samym akapicie podać udział zakresu 3, brak celu dla zakresu 3 i brak planu przejścia. Nie łączyć „ważnego aspektu działalności” z procentem redukcji z przypadku 01.' `
        'SFORMUŁOWANIE NIE DO URATOWANIA' `
        'Zdanie, że Grupa minimalizuje negatywny wpływ swojej działalności na otoczenie, nie ma bezpiecznej wersji skrótowej. Po korekcie zostaje opis wpływu z macierzy: wpływ negatywny przez emisję, w wyższym szczeblu łańcucha i w operacjach własnych, oraz liczby z E1-6. Nie składa się claimu o minimalizacji.'

    Add-Case 'PRZYPADEK 05: Hybrydy jako najważniejsze działanie klimatyczne floty' `
        'Podstawa prawna: ECGT, załącznik I pkt 4b · UPNPR art. 6 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            '„Najważniejsze z podjętych działań w roku sprawozdawczym dotyczyły obszaru floty samochodowej i gwarancji pochodzenia dla energii.” (E1-3, s. 218).',
            '„W Polsce uzupełniono flotę o 5 aut hybrydowych oraz 1 hybrydę plug-in, natomiast we Włoszech o 79 aut hybrydowych.” Osiągnięta redukcja: 42,86 t CO2; spodziewana: 42,86 t (ta sama tabela).',
            '„Zużycie paliwa z ropy naftowej i produktów naftowych (MWh)” 2025 = 9 405,0452 (E1-5, s. 219). Zakres 1 = 2 262 tCO2e.'
        ) `
        'Punkt 4b obejmuje sytuację, w której wycinek – tu: dopisane hybrydy – jest podany jako obraz działania klimatycznego całości. Sprawozdanie samo nazywa flotę i gwarancje „najważniejszymi” działaniami łagodzącymi. Redukcja przypisana flocie to 42,86 t, czyli około 1,9% emisji zakresu 1, i jest równa wartości spodziewanej, a więc nie wynika z porównania dwóch inwentaryzacji, tylko z przypisania tej samej liczby do roku 2025. Nie podano liczebności floty, udziału hybryd, ani ile z zakresu 1 pochodzi ze spalania mobilnego. W koszyku energii paliwa ropopochodne to 9 405 MWh, około 71% całkowitego zużycia 13 264 MWh. Czytelnik nie dostaje mianownika, bez którego „uzupełniono flotę” nie mówi, czy hybryda jest regułą, czy dopiskiem. To pominięcie z art. 6 i greenlighting z pkt 4b. Sama informacja, że w Polsce doszło pięć hybryd i jedna hybryda plug-in, a we Włoszech 79 hybryd, jest do utrzymania, gdy stoi przy niej mianownik albo gdy wprost napisano, że udziału nie podano i że 42,86 t nie jest redukcją zakresu 1 wobec roku bazowego.' `
        'Zdjąć „najważniejsze z podjętych działań” jako ramę dla 42,86 t. Przy wierszu floty podać liczbę pojazdów służbowych albo napisać, że sprawozdanie jej nie zawiera, oraz że zakres 1 wobec bazy 2021 nie spadł. Nie sumować 42,86 t z 2 283,70 t z przypadku 02 w jedną „redukcję Grupy”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2025 r. do floty służbowej dopisano w Polsce 5 samochodów hybrydowych i 1 hybrydę plug-in oraz 79 hybryd we Włoszech. Przypisana tym pojazdom wielkość wynosi 42,86 t CO2. Sprawozdanie nie podaje liczebności całej floty. Zużycie paliw z ropy naftowej i produktów naftowych wyniosło 9 405 MWh. Emisja zakresu 1 wyniosła 2 262 tCO2e i jest równa bazie z 2021 r. Dopisanie tych pojazdów nie jest twierdzeniem o emisyjności floty.”'

    Add-Case 'PRZYPADEK 06: Ogólne „zrównoważone praktyki”, „działania proekologiczne” i „zrównoważony rozwój”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            '„W 2025 roku Grupa KRUK kontynuowała działania na rzecz łagodzenia zmian klimatu jednocześnie promując zrównoważone praktyki i dzieląc się wiedzą na temat działań proekologicznych wśród pracowników.” (E1-3, s. 218).',
            '„Nowa siedziba została zaprojektowana z myślą o zrównoważonym rozwoju i oszczędzaniu energii.” (E1-3, s. 218; ocena certyfikatu – przypadek 09).',
            '„kampanię „Odpowiedzialność na co dzień”, której celem było wzmocnienie wiedzy o zrównoważonym rozwoju w Grupie, w tym zasad compliance i etyki, cyberbezpieczeństwa czy procesu współpracy z klientami.” (S1-2, s. 233).'
        ) `
        'Ogólne twierdzenie środowiskowe – w tym „zrównoważony” i określenie typu „proekologiczny” – jest zakazane, jeżeli w tym samym kanale nie wykazano uznanej wysokiej efektywności ekologicznej (EU Ecolabel albo oznakowanie ISO 14024 typu I) istotnej dla hasła. Sprawozdanie takiej efektywności nie wykazuje. Zdanie o „zrównoważonych praktykach” i „działaniach proekologicznych” nie mówi, jaki próg, jaki zakres i jaki sprawdzalny wynik czyni praktykę zrównoważoną. Dalszy akapit opisuje relokację biura, LED, wentylację, LEED oraz tabelę hybryd i gwarancji. To są konkretne czynności, ale nie są uznaną wysoką efektywnością w rozumieniu pkt 4a, więc nie „zdają” hasła ogólnego. Osobno kampania „Odpowiedzialność na co dzień” używa „zrównoważonego rozwoju” jako nazwy zbiorczej dla compliance, etyki, cyberbezpieczeństwa i obsługi klienta. Hasło środowiskowe obejmuje wtedy tematy, które środowiskowymi nie są. Motyw ogólnego claimu jest właśnie taki: słowo bez specyfikacji zbiera cechy, których nadawca nie umie ograniczyć. Nazwa „sprawozdanie zrównoważonego rozwoju” jest nazwą dokumentu wymaganą przez ramy sprawozdawcze i nie jest tu traktowana jako claim produktowy. Claimem są zdania o praktykach, działaniach proekologicznych i wiedzy o zrównoważonym rozwoju w kampanii wewnętrznej, która – jak sprawozdanie samo podaje – ma także kanał zewnętrzny, bo informacje pracownicze idą m.in. na LinkedIn.' `
        'W E1-3 skreślić „zrównoważone praktyki” i „działania proekologiczne”. Zostawić opis czynności, które da się sprawdzić (przypadki 02, 05 i 09), bez przymiotnika zbiorczego. W opisie kampanii napisać, czego dotyczyły moduły: compliance, etyka, cyberbezpieczeństwo, obsługa klienta. Nie nazywać tego zestawu zrównoważonym rozwojem środowiskowym.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2025 r. działania opisane w E1-3 to: relokacja siedziby do budynku z certyfikatem LEED v4 Platinum, dopisanie do floty pojazdów wymienionych w tabeli oraz umorzenie gwarancji pochodzenia w wolumenie z E1-6. Kampania „Odpowiedzialność na co dzień” dotyczyła compliance, etyki, cyberbezpieczeństwa i współpracy z klientem. Nie jest twierdzeniem o efektywności środowiskowej Grupy.”'

    Add-Case 'PRZYPADEK 07: Obowiązki Kodeksu pracy jako cecha oferty wobec pracowników' `
        'Podstawa prawna: ECGT, załącznik I pkt 10a · UPNPR art. 5 · Kodeks pracy: art. 18³c, art. 94 pkt 2b, art. 94³, art. 207 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            '„Stawiamy na rozwój swoich pracowników, oferując szeroką gamę szkoleń, zapewniając uczciwe wynagrodzenie, bezpieczne warunki pracy i podejmując działania związane z dobrostanem i zwiększaniem ich zaangażowania.” (SBM-1, nakłady ludzkie, s. 170).',
            '„zapewnienie bezpiecznych i zdrowych warunków pracy - Grupa KRUK zapewnia bezpieczne i higieniczne środowisko pracy, stale minimalizując ryzyka i wspierając dobrostan psychiczny oraz fizyczny pracowników.” (zobowiązania Polityki Praw Człowieka, s. 231).',
            '„zapobieganie dyskryminacji, mobbingowi i molestowaniu - Grupa KRUK aktywnie przeciwdziała wszelkim zachowaniom naruszającym godność”. „zapewnienie uczciwych warunków zatrudnienia - Grupa KRUK przestrzega przepisów prawa pracy, gwarantując równe szanse, transparentne zasady wynagradzania” (s. 231).',
            '„W związku ze zmianą w/w wyników, nie ujęto także w niniejszym Sprawozdaniu ZR wskaźników S1-14 – Mierniki Bezpieczeństwa i Higiena Pracy.” (BP-2, s. 158). W Rumunii przeprowadzono „dodatkowe szkolenie dotyczące przeciwdziałania mobbingowi i dyskryminacji” jako element kultury DEI (s. 239).'
        ) `
        'Punkt 10a zakazuje przedstawiania uprawnień przysługujących pracownikom z mocy prawa jako cechy wyróżniającej ofertę przedsiębiorcy. Równe wynagrodzenie za równą pracę lub pracę równej wartości wynika z art. 18³c Kodeksu pracy. Bezpieczeństwo i higiena pracy, w tym koszty BHP po stronie pracodawcy, wynikają z art. 207. Przeciwdziałanie mobbingowi i dyskryminacji wynika z art. 94 pkt 2b i art. 94³. Sprawozdanie stawia „uczciwe wynagrodzenie” i „bezpieczne warunki pracy” w jednym rzędzie z tym, co Grupa „oferuje”, obok szkoleń i dobrostanu. Polityka praw człowieka powtarza bezpieczne i higieniczne środowisko oraz aktywne przeciwdziałanie mobbingowi jako zobowiązanie Grupy, a w kolejnym punkcie sama przyznaje, że chodzi o przestrzeganie przepisów prawa pracy. Szkolenie antymobbingowe jest opisane jako element wzmacniania kultury DEI, nie jako wykonanie obowiązku ustawowego. Jednocześnie BP-2 wyłącza miernik S1-14, bo bezpieczeństwo i higiena pracy nie przeszły progu istotności. Czytelnik dostaje zdanie o zapewnianiu bezpiecznego środowiska i nie dostaje wskaźnika wypadków, który standard przy tym temacie przewiduje. To pkt 10a oraz pominięcie z art. 6 przy twierdzeniu o bezpieczeństwie pracy. Nie jest ustaleniem tego audytu, że Grupa łamie Kodeks pracy. Ustalenie dotyczy sposobu opisania obowiązku. Szkolenia zawodowe, dobrostan i budżet rozwojowy powyżej ustawowego minimum mogą zostać, jeśli są oddzielone od zdań o BHP, mobbingu i równym wynagrodzeniu i jeśli mają własną miarę.' `
        'W SBM-1 rozdzielić zdanie. Szkolenia i dobrostan zostawić z liczbą (godziny, budżet, odsetek), jeśli jest w sprawozdaniu. „Uczciwe wynagrodzenie” i „bezpieczne warunki pracy” przenieść do zdania „Grupa wykonuje obowiązki Kodeksu pracy”, bez czasownika „oferując”. Przy przeciwdziałaniu mobbingowi dopisać podstawę ustawową. Albo ujawnić S1-14, albo nie twierdzić, że środowisko pracy jest bezpieczne i higieniczne jako cecha wyróżniająca, skoro miernik uznano za nieistotny.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Grupa wykonuje obowiązki pracodawcy z Kodeksu pracy, w tym art. 207 (BHP), art. 94 pkt 2b i art. 94³ (dyskryminacja i mobbing) oraz art. 18³c (równe wynagrodzenie za równą pracę). Nie są to świadczenia powyżej ustawy. Miernika S1-14 w tym sprawozdaniu nie ujawniono, ponieważ temat uznano za nieistotny w ocenie podwójnej istotności. Szkolenia inne niż obowiązkowe BHP opisujemy osobno, z liczbą godzin albo odsetkiem ukończenia.”'

    Add-Case 'PRZYPADEK 08: Skorygowana luka 0,1% i podmiana mianownika szkolenia DEI' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Cel: „0% różnicy w wynagrodzeniach kobiet i mężczyzn do 2025 roku”. Wykonanie: „Skorygowana luka płacowa w wynagrodzeniach kobiet i mężczyzn w Grupie KRUK wyniosła 0,1%” (SBM-1, s. 171). Przypis definiuje ją jako różnicę wynagrodzenia całkowitego na tym samym poziomie, w tej samej rodzinie zawodów i na tym samym stanowisku, jako średnią ważoną zatrudnieniem.',
            '„Grupa nie ujawnia wskaźników (…) S1-16 – Mierniki wynagrodzeń (…) ze względu na to, że (…) równouprawnienie płci i równość wynagrodzeń nie zostały zidentyfikowane jako istotne.” (BP-2, s. 158). W indeksie ESRS pozycja S1-16 (nieskorygowana luka) jest oznaczona „Nieistotne”.',
            'Cel DEI: „Do końca 2025 roku 100% menadżerów zostanie przeszkolonych w zakresie różnorodności i integracji”, a co najmniej 90% pracowników będzie uważało Grupę za miejsce różnorodne i inkluzywne. Wykonanie: „szkolenie w zakresie różnorodności i integracji przeszło 76% pracowników”. W badaniu zaangażowania 92% oceniło Grupę jako takie miejsce (s. 171).',
            'Cel „Co najmniej 40% kobiet w zarządach spółek”. Wykonanie: „Łącznie w Zarządzie i Radzie Nadzorczej Kruk S.A. zasiadało 42% kobiet.” (s. 171).'
        ) `
        'Skorygowana luka 0,1% jest opisana metodycznie i nie jest przedstawiona jako zero. Naruszeniem nie jest sam odsetek, tylko zestawienie go z celem „0% różnicy” oraz wyłączenie S1-16. Czytelnik strategii widzi wynik bliski celu. Czytelnik BP-2 dowiaduje się, że równość wynagrodzeń uznano za nieistotną i że miernika wynagrodzeń ESRS, obejmującego lukę nieskorygowaną, nie podano. Bez luki nieskorygowanej nie wiadomo, czy 0,1% na tym samym stanowisku opisuje rozkład płac w Grupie. To pominięcie istotne przy zdaniu o wykonaniu celu równości wynagrodzeń (art. 6). Osobno cel szkoleniowy ma mianownik „100% menedżerów”, a wykonanie – „76% pracowników”. Odsetek pracowników nie odpowiada na pytanie o menedżerów. Drugi element celu (90% pozytywnych ocen inkluzywności) jest wykonany: 92%. Pierwszy nie jest wykazany. Podanie 76% w rubryce wykonania celu menedżerskiego może zostać odczytane jako realizacja tego celu (art. 5). Cel „40% kobiet w zarządach spółek” jest rozliczony łącznie dla zarządu i rady nadzorczej KRUK S.A., nie dla zarządów spółek Grupy. Mianownik jest węższy i inny niż w celu. W tym sprawozdaniu nie ma podstawy, by twierdzić, że luka płacowa jest niezerowa w innym ujęciu. Podstawa jest taka, że ujęcie, które pozwoliłoby to sprawdzić, zostało pominięte przy jednoczesnym pokazaniu ujęcia korzystnego.' `
        'Przy 0,1% dopisać, że jest to luka skorygowana według przypisu oraz że luki nieskorygowanej (S1-16) nie ujawniono. Nie nazywać 0,1% wykonaniem celu „0%”, albo napisać wprost odchylenie 0,1 punktu procentowego. Przy DEI podać odsetek menedżerów albo napisać, że odsetka menedżerów sprawozdanie nie zawiera; 76% pracowników zostawić jako osobny wskaźnik. Przy celu 40% albo podać udział w zarządach poszczególnych spółek, albo zmienić opis celu tak, by dotyczył zarządu i rady KRUK S.A.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Skorygowana luka płacowa w 2025 r. wyniosła 0,1% (to samo stanowisko, ta sama rodzina zawodów, ten sam poziom, średnia ważona zatrudnieniem). Cel strategii brzmiał „0%”. Nieskorygowanej luki S1-16 nie ujawniono, bo temat uznano za nieistotny. Szkolenie DEI ukończyło 76% pracowników; odsetka menedżerów, wobec celu 100% menedżerów, sprawozdanie nie podaje. W badaniu zaangażowania 92% pracowników oceniło Grupę jako miejsce różnorodne i inkluzywne. Udział 42% kobiet dotyczy łącznie zarządu i rady nadzorczej KRUK S.A., nie zarządów wszystkich spółek.”'

    Add-Case 'PRZYPADEK 09: LEED v4 Platinum jako „zrównoważony rozwój” siedziby' `
        'Podstawa prawna: UPNPR art. 5 · dyrektywa (UE) 2024/1275 art. 11 jako granica, której LEED nie zastępuje | Poziom ryzyka: ŚREDNIE (transfer atrybutów)' `
        @(
            '„W Polsce nastąpiła relokacja siedziby KRUK S.A., Kancelarii Prawnej RAVEN, KRUK TFI i Novum do nowego certyfikowanego budynku” we Wrocławiu. „Nowa siedziba została zaprojektowana z myślą o zrównoważonym rozwoju i oszczędzaniu energii. Wykorzystano w niej nowoczesne technologie, takie jak systemy zarządzania energią, energooszczędne oświetlenie LED oraz zaawansowane systemy wentylacji i klimatyzacji. Obiekt posiada certyfikat LEED v4 na poziomie Platinum.” (E1-3, s. 218).',
            'W tabeli taksonomicznej wydatków operacyjnych działalność „Nabywanie i prawo własności budynków”, kod CCM 7.7, wykazana jest z wartością 0. „W ramach KPI OpEx nie zidentyfikowano wydatków związanych z działalnością zgodną z Taksonomią.” (s. 216).'
        ) `
        'LEED v4 jest dobrowolnym katalogiem kredytów i poziomem certyfikacji budynku. Nie jest definicją budynku bezemisyjnego z art. 11 dyrektywy (UE) 2024/1275 i nie jest zgodnością z taksonomią. Sprawozdanie podaje schemat i poziom (LEED v4 Platinum) oraz wymienia LED, zarządzanie energią i wentylację, czego przy samym haśle by brakowało. Nie podaje numeru certyfikatu, daty, podmiotu certyfikującego ani tego, czy certyfikat obejmuje budynek wynajmowany, czy także sposób, w jaki Grupa ten budynek użytkuje. Zdanie „zaprojektowana z myślą o zrównoważonym rozwoju” dokłada do certyfikatu ocenę ogólną, której LEED nie jest synonimem. W tym samym sprawozdaniu nabywanie budynków (7.7) ma w liczniku zgodności zero. To nie musi być sprzeczność rachunkowa – najem i certyfikat właściciela nie muszą wejść do KPI najemcy – ale czytelnik E1-3 nie dostaje zdania, że Platinum nie oznacza zgodności taksonomicznej ani braku emisji na miejscu z paliw kopalnych. Artykuł 5 obejmuje wprowadzenie w błąd co do oznaczeń i korzyści. Ryzyko jest średnie, bo zakres (jeden budynek we Wrocławiu, cztery podmioty, które się do niego przeniosły) jest węższy niż claim o całej Grupie. Nie wolno tego akapitu czytać łącznie z przypadkami 01, 02 i 04 jako dowodu, że Grupa jest niskoemisyjna.' `
        'Zostawić LEED v4 Platinum przy nazwie budynku i liście podmiotów, które się przeniosły. Dopisać, czego certyfikat nie stwierdza: nie jest to art. 11 EPBD i nie jest to KPI taksonomiczne Grupy (zgodność 0%). Skreślić „zaprojektowana z myślą o zrównoważonym rozwoju” albo zastąpić listą cech, które certyfikat albo projekt faktycznie potwierdza, z numerem i datą dokumentu. Nie wiązać akapitu z celem redukcji z E1-4.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2025 r. KRUK S.A., Kancelaria Prawna RAVEN, KRUK TFI i Novum przeniosły siedzibę do budynku we Wrocławiu z certyfikatem LEED v4 na poziomie Platinum (numer i data: [uzupełnić]). Certyfikat dotyczy tego budynku w zakresie systemu LEED. Nie oznacza, że budynek jest bezemisyjny według art. 11 dyrektywy (UE) 2024/1275, i nie jest ujawnieniem zgodności Grupy z taksonomią. KPI obrotu, CapEx i OpEx zgodne z taksonomią wynoszą 0%.”'

    Add-Case 'PRZYPADEK 10: „Pełne bezpieczeństwo” sygnalistów przy węższej procedurze' `
        'Podstawa prawna: ustawa z 14.06.2024 r. o ochronie sygnalistów, art. 3 ust. 1 i ust. 2 · UPNPR art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            '„zapewnienie bezpiecznych kanałów zgłaszania naruszeń - Grupa KRUK udostępnia anonimowe i chronione kanały zgłaszania naruszeń, zapewniając pełne bezpieczeństwo i brak działań odwetowych wobec sygnalistów.” (Polityka Praw Człowieka, s. 231).',
            'Procedura zgłoszeń wewnętrznych kanałem whistleblowing, zakres podmiotowy w tabeli G1-1: „KRUK S.A., KRUK Romania, KRUK Italia, KRUK España, InvestCapital.” Procedura obejmuje działanie „zgodnie z obowiązującym prawem, przyjętymi regulacjami wewnętrznymi, a także standardami etycznymi” (s. 287–288).'
        ) `
        'Ustawa o ochronie sygnalistów chroni zgłoszenie naruszenia prawa z katalogu art. 3 ust. 1, w tym ochrony środowiska, ochrony konsumentów oraz usług i rynków finansowych. Jeżeli zgłoszenie dotyczy wyłącznie standardu etycznego, a nie naruszenia prawa, nie stosuje się zgłoszenia zewnętrznego do organu ani ochrony przy ujawnieniu publicznym (art. 3 ust. 2). Sprawozdanie opisuje procedurę jako narzędzie prawa, regulacji wewnętrznych i standardów etycznych, a w polityce praw człowieka obiecuje „pełne bezpieczeństwo i brak działań odwetowych” bez tego rozróżnienia. Czytelnik może przyjąć, że każde zgłoszenie etyczne ma tarczę ustawową. Tego ustawa nie daje. Osobno tabela ogranicza procedurę do pięciu podmiotów. W Grupie są także Wonga, Novum, KRUK TFI i Kancelaria Prawna RAVEN oraz zatrudnienie na Malcie, w Czechach i w Luksemburgu. Zdanie o kanałach Grupy tego wyłączenia nie powtarza. To pominięcie zakresu podmiotowego i przedmiotowego (art. 6). Audyt nie stwierdza, że kanałów w pominiętych spółkach nie ma. Stwierdza, że sprawozdanie obiecuje je dla Grupy, a tabelę procedury zamyka węższą listą. Kanał, który realnie przyjmuje zgłoszenie nieuczciwej praktyki środowiskowej lub konsumenckiej, jest obowiązkiem i zabezpieczeniem, nie ozdobnikiem. Opis ma być równy ustawie, nie szerszy od niej.' `
        'Przy zdaniu o „pełnym bezpieczeństwie” dopisać art. 3: ochrona ustawowa dotyczy naruszenia prawa z katalogu, a zgłoszenie wyłącznie etyczne idzie trybem wewnętrznym bez ścieżki zewnętrznej i bez ochrony ujawnienia publicznego. Listę spółek z tabeli G1-1 powtórzyć przy zdaniu o Grupie albo uzupełnić procedurę i listę o pozostałe spółki, które zatrudniają ludzi. Nie obiecywać anonimowości i braku odwetu w szerszym zakresie niż procedura i ustawa.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Procedura zgłoszeń wewnętrznych obowiązuje w KRUK S.A., KRUK Romania, KRUK Italia, KRUK España i InvestCapital. Ochrona z ustawy o ochronie sygnalistów dotyczy zgłoszenia naruszenia prawa wymienionego w art. 3 ust. 1. Zgłoszenie dotyczące wyłącznie standardu etycznego nie uruchamia zgłoszenia zewnętrznego ani ochrony przy ujawnieniu publicznym (art. 3 ust. 2). Zakaz działań odwetowych za zgłoszenie objęte ustawą wynika z tej ustawy, nie z przywileju przyznanego przez Grupę.”'

    Add-Case 'PRZYPADEK 11: „Niskoemisyjna infrastruktura i transport” w Kodeksie Etycznym' `
        'Podstawa prawna: UPNPR art. 5 · dyrektywa 2005/29/WE art. 6 w brzmieniu ECGT (przyszła efektywność) | Poziom ryzyka: ŚREDNIE' `
        @(
            '„W odniesieniu do środowiska, Kodeks Etyczny akcentuje potrzebę ograniczania wpływu organizacji na klimat poprzez monitorowanie oddziaływania, redukcję śladu węglowego oraz rozwój zrównoważonej, niskoemisyjnej infrastruktury i transportu.” (G1-1, s. 285).',
            'E1-1: brak planu przejścia na 31.12.2025. E1-5: udział źródeł kopalnych 81,67%. E1-3: brak szczegółowych planów działań.'
        ) `
        'Zdanie relacjonuje, co akcentuje kodeks, a nie że infrastruktura już jest niskoemisyjna. W sprawozdaniu publikowanym na zewnątrz takie sformułowanie czyta się jednak jako cechę organizacji: „zrównoważona, niskoemisyjna infrastruktura i transport”. Przyszła efektywność wymaga planu, którego E1-1 i E1-3 nie zawierają. Koszyk energii jest w 81,67% kopalny, a transport służbowy w działaniach 2025 r. to dopisanie hybryd przy 9 405 MWh paliw ropopochodnych (przypadek 05). Zestawienie akapitu kodeksu z tymi liczbami, bez odesłania, pozwala wziąć kierunek kodeksu za opis stanu. To art. 5: informacja o treści kodeksu, prawdziwa jako cytat wewnętrznej regulacji, podana w miejscu, w którym brakuje stanu faktów. Nie jest to pkt 4c. Kodeks nie twierdzi, że transport jest neutralny dzięki kompensacji.' `
        'Zostawić relację z kodeksu i w tym samym zdaniu odesłać do E1-1 (brak planu), E1-5 (81,67% energii kopalnej) oraz do tabeli floty. Nie używać przymiotników „zrównoważona, niskoemisyjna” bez tego odesłania w żadnym wypisie kodeksu na stronie esg.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Kodeks Etyczny wskazuje kierunek: monitorowanie wpływu na klimat i ograniczanie śladu. Na 31.12.2025 Grupa nie ma planu przejścia ani szczegółowego planu działań. W 2025 r. 81,67% zużytej energii pochodziło ze źródeł kopalnych, w tym 9 405 MWh z paliw ropopochodnych. Opis floty i gwarancji pochodzenia jest w E1-3 i nie stanowi twierdzenia, że infrastruktura lub transport są niskoemisyjne.”'

    Add-H '4. Porównanie twierdzeń Grupy KRUK z rzeczywistym stanem prawno-środowiskowym'
    Add-P 'Sprawozdanie nie zestawia Grupy, jej usługi ani jej produktu z konkurentem, z innym materiałem ani z cudzym wskaźnikiem w sposób, który byłby reklamą porównawczą. Tabeli z art. 16 ustawy o zwalczaniu nieuczciwej konkurencji nie sporządza się. Poniższe zestawienie nie jest porównaniem z rynkiem. Jest zestawieniem twierdzenia z tego pliku z liczbami z tego samego pliku, bo sprzeczność jest wewnętrzna.' 10.5 $false $false $DARK 0 8

    $comp = @(
        @('„70% redukcji emisji dwutlenku węgla do 2040 roku” wobec bazy 2021.', 'Tabela E1-4 ustawia cel zakresu 1 na 1 583,4 t przy bazie 2 262 t, czyli na 70% bazy (redukcja o 30%), a stopę roczną 3,68% bazy, która przez 19 lat daje redukcję około 70%. Emisja zakresu 1 w 2025 r. = 2 262 tCO2e, jak w 2021 r. Planu przejścia nie ma.'),
        @('„Dostawy zielonej energii” i osiągnięta redukcja 2 283,70 t CO2.', 'Gwarancje pochodzenia: 2 267,36 MWh. Zakres 2: 1 244 tCO2e (lokalizacja) i 374 tCO2e (rynek, +98,93% r/r). Energia kopalna: 81,67% koszyka. Liczba 2 283,70 t nie występuje w inwentaryzacji jako zmiana emisji.'),
        @('Całkowite emisje GHG −20,76% (lokalizacja).', 'Z różnicy 10 018 tCO2e około 7 926 t przypada na dobra inwestycyjne, a około 1 876 t na dojazdy (1 881 → 5 tCO2e, dane ankietowe o niskiej dokładności). Kategoria 1: −0,53%.'),
        @('„Minimalizuje negatywny wpływ swojej działalności” na środowisko.', 'Wpływ E1 po rewizji: negatywny, także w wyższym szczeblu łańcucha. Zakres 3 = 34 729 tCO2e, czyli 92,9% śladu metody rynkowej (37 365 tCO2e). Cel redukcyjny tego zakresu nie obejmuje.'),
        @('Flota hybrydowa jako najważniejsze działanie łagodzące; 42,86 t.', 'Paliwa ropopochodne: 9 405 MWh. Udział hybryd we flocie: niepodany. Zakres 1 wobec bazy 2021: bez redukcji.'),
        @('LEED v4 Platinum i „zrównoważony rozwój” siedziby.', 'Certyfikat dotyczy wskazanego budynku. KPI taksonomiczne obrotu, CapEx i OpEx zgodne z systematyką: 0%. LEED nie jest budynkiem bezemisyjnym z art. 11 dyrektywy (UE) 2024/1275.'),
        @('Cel „0% różnicy” w wynagrodzeniach; wykonanie 0,1%.', 'Podano lukę skorygowaną. Luki nieskorygowanej S1-16 nie ujawniono, bo równość wynagrodzeń uznano za nieistotną. Równe wynagrodzenie za równą pracę jest obowiązkiem z art. 18³c Kodeksu pracy.')
    )
    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $ct = $doc.Tables.Add($range, (1 + $comp.Count), 2)
    $ct.PreferredWidthType = 3
    $ct.PreferredWidth = $contentW
    $ct.AllowAutoFit = $false
    $ct.Columns.Item(1).Width = $word.CentimetersToPoints(7.2)
    $ct.Columns.Item(2).Width = $word.CentimetersToPoints(10.4)
    Write-Cell $ct.Cell(1, 1) 'Twierdzenie w sprawozdaniu' 8 $true $WHITE
    Write-Cell $ct.Cell(1, 2) 'Stan wynikający z tego samego sprawozdania' 8 $true $WHITE
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
    Add-P '• 1. Kary Prezesa UOKiK' 12 $true $false $NAVY 8 2
    Add-P 'Artykuł 24 ustawy o ochronie konkurencji i konsumentów zakazuje praktyk naruszających zbiorowe interesy konsumentów, w tym informacji, która nie jest rzetelna, prawdziwa i pełna, oraz nieuczciwych praktyk rynkowych. Artykuł 106 ust. 1 pkt 4 pozwala nałożyć karę do 10% obrotu za naruszenie art. 24, także nieumyślne. Kluczową grupą klientów Grupy są konsumenci – osoby zadłużone. Sprawozdanie opisuje komunikację zewnętrzną (kampanie, LinkedIn, strona z politykami ESG). Zdania o zielonej energii, o 70% redukcji i o zrównoważonych praktykach, jeśli zostaną powtórzone w serwisie, w kampanii albo w materiale dla klienta, są praktyką podlegającą temu reżimowi już na art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. Od 27 września 2026 r. pkt 4a, 4b i 10a są zakazane w każdych okolicznościach. Termin transpozycji w druku Senatu 825 (27 września 2027 r.) tego nie zawiesza.' 10.5 $false $false $DARK 0 8
    Add-P '• 2. Zgłoszenie sygnalisty' 12 $true $false $NAVY 8 2
    Add-P 'Ustawa o ochronie sygnalistów obejmuje naruszenia prawa z zakresu ochrony środowiska i ochrony konsumentów. Pracownik, który zestawi zdanie o redukcji 2 283,70 t albo o spadku śladu o 20,76% z tabelą E1-6, ma podstawę do zgłoszenia wewnętrznego, a przy naruszeniu prawa – także zewnętrznego do Prezesa UOKiK z pominięciem ścieżki wewnętrznej. Ujawnienie rozbieżności między komunikatem a danymi źródłowymi nie jest objęte odpowiedzialnością za tajemnicę przedsiębiorstwa, jeżeli sygnalista miał uzasadnione podstawy sądzić, że ujawnienie jest niezbędne. Działanie odwetowe jest zagrożone karą. Obietnica „pełnego bezpieczeństwa” z przypadku 10 nie zastępuje tego reżimu i nie może być węższa niż ustawa.' 10.5 $false $false $DARK 0 8
    Add-P '• 3. Odbiorcy sprawozdania, którzy sami raportują ESRS' 12 $true $false $NAVY 8 2
    Add-P 'Banki, które finansują Grupę, oraz inwestorzy mogą przenieść cel „70% do 2040 r.”, „zieloną energię” i −20,76% do własnego łańcucha wartości. Atestacja zgodności ujawnień ze standardami, opisana w BP-1, nie uzgadnia sprzecznych liczb. Odbiorca, który wpisze 2 283,70 t jako redukcję zakresu 2, wpisze wielkość wyższą niż cały zakres 2 Grupy według metody lokalizacyjnej (1 244 tCO2e). Definicja greenwashingu przyjęta przez ESA 4 czerwca 2024 r. obejmuje komunikację zrównoważoną, która nie oddaje rzetelnie profilu podmiotu i może wprowadzać w błąd, także bez zamiaru. Dla podmiotu z sektora finansowego obowiązuje ten sam wymóg: komunikat ma być dokładny, rzetelny i niewprowadzający w błąd. Sprawozdanie nie kwalifikuje funduszu według SFDR i nie nazywa produktu nazwą ESG w rozumieniu wytycznych ESMA; tego reżimu tu nie stosuję. Stosuję standard rzetelności komunikacji.' 10.5 $false $false $DARK 0 8
    Add-P '• 4. Social washing i data 27 września 2026 r.' 12 $true $false $NAVY 8 2
    Add-P 'Punkt 10a dotyczy tego, jak Grupa opisuje obowiązki wobec pracowników, nie tego, czy te obowiązki wykonuje. Audyt nie stwierdza naruszenia art. 18³c, art. 94 ani art. 207 Kodeksu pracy i nie przenosi na to sprawozdanie grzywien Państwowej Inspekcji Pracy. Ryzyko komunikacyjne jest inne: od 27 września 2026 r. przedstawienie ustawowego uprawnienia pracownika jako cechy oferty jest praktyką zakazaną w każdych okolicznościach. Cel „0% różnicy” przy nieujawnionej luce nieskorygowanej utrudnia pracownikowi i inwestorowi sprawdzenie, co zostało zmierzone. Roszczenie z art. 18³d Kodeksu pracy wymaga stanu faktycznego, którego w tym pliku nie ma. Do czasu ujawnienia S1-16 albo rezygnacji z celu „0%” zdanie o wykonaniu równości wynagrodzeń nie powinno wychodzić poza przypis metody skorygowanej.' 10.5 $false $false $DARK 0 8

    Add-H '6. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIASTOWY PRZEGLĄD KOMUNIKACJI (12 października 2026 r.)' @(
        'Wstrzymać powielanie, w serwisie relacji inwestorskich, w wypisach ESG, w prezentacjach i w komentarzu kwartalnym, zdań: „70% redukcji do 2040 r.”, „dostawy zielonej energii”, „osiągnięta redukcja 2 283,70 t”, „minimalizuje negatywny wpływ swojej działalności”, „zrównoważone praktyki” i „działania proekologiczne” oraz procentu −20,76% bez mostu z przypadku 03. Sprawozdania nie wycofywać z obiegu prawnego, jeżeli zostało złożone jako oświadczenie. Do każdej kopii używanej wobec inwestora, banku lub klienta dołączyć sprostowanie przypadków 01–03. Nie sumować 2 283,70 t i 42,86 t w jedną redukcję Grupy.'
    ) (Rgb 255 241 242) $DANGER
    Add-Callout 'ETAP 2: KOREKTA NARRACJI SPRAWOZDANIA (do 19 października 2026 r.)' @(
        'W E1-4 zostawić jedną parę liczb: baza, rok 2025, zakres. Procent 70% usunąć do czasu planu przejścia. W E1-3 wykreślić „zieloną energię” i tony 2 283,70 oraz ramę „najważniejsze działania” przy hybrydach. Przy tabeli E1-6 wstawić most kategorii 2 i 7. W S3 skreślić minimalizację wpływu całej działalności. W SBM-1 rozdzielić obowiązki Kodeksu pracy od szkoleń. Przy luce 0,1% dopisać brak S1-16. Przy DEI podać mianownik menedżerów albo jego brak. Przy LEED dopisać, czym certyfikat nie jest. Przy sygnalistach powtórzyć listę spółek i art. 3 ust. 2.'
    ) (Rgb 255 251 235) $WARNING
    Add-Callout 'ETAP 3: UZGODNIENIE LICZB I ZAKRESU (do 11 listopada 2026 r.)' @(
        'Uzgodnić 2 283,70 t z E1-6 albo usunąć tę liczbę na stałe. Opisać, skąd 94,97% udziału instrumentów przy zakresie 2 rynkowym 374 tCO2e i lokalizacyjnym 1 244 tCO2e. Sprawdzić kategorię 7 (5 tCO2e) i albo ją skorygować, albo jawnie oznaczyć jako nieporównywalną. Nie podnosić udziału danych pierwotnych zakresu 3 ponad 0%, dopóki dostawcy tych danych nie dali. Do celu na 2040 r. nie wracać bez planu przejścia: zakres, rok bazowy, poziom w tCO2e, nakłady, termin weryfikacji. Decyzja o S1-16 i S1-14: albo ujawnienie, albo rezygnacja ze zdań, które te mierniki miałyby potwierdzać. Uzupełnić procedurę whistleblowing o spółki zatrudniające poza listą z G1-1 albo wprost tę listę utrzymać przy każdym zdaniu o „Grupie”.'
    ) (Rgb 240 249 255) $ACCENT
    Add-Callout 'ETAP 4: SAFE HARBOR I STAŁY NADZÓR (od 12 października 2026 r., tryb ciągły)' @(
        'Każde zdanie środowiskowe wychodzące z Departamentu ESG i z komunikacji zewnętrznej ma mieć: zakres (który podmiot, który gaz, która metoda zakresu 2), rok, jednostkę i miejsce w sprawozdaniu, z którego liczba pochodzi. Zakazane skróty: „zielona energia” bez słów „gwarancja pochodzenia”, „redukcja” bez dwóch lat i tego samego zakresu, „zrównoważony” i „proekologiczny” bez EU Ecolabel albo ISO 14024 typu I, „minimalizujemy wpływ” bez udziału zakresu 3, LEED jako cecha Grupy, obowiązek Kodeksu pracy jako benefit. Szkolenie osób, które piszą slajd inwestorski i wpis na LinkedIn, na tych pięciu zakazach. Plan przejścia zapowiedziany na 2026 r. jest warunkiem jakiegokolwiek nowego procentu, nie uzasadnieniem procentu już opublikowanego.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '7. Wnioski'
    Add-P 'Sprawozdanie za 2025 r. w tabelach ESRS ujawnia brak planu przejścia, zerową zgodność taksonomiczną i 81,67% energii ze źródeł kopalnych, ale narracja – „70% redukcji”, „zielona energia” z redukcją 2 283,70 t, spadek śladu o 20,76% oraz ogólne zrównoważenie i minimalizacja wpływu całej działalności – nie może być publikowana ani powielana bez zawężenia do liczb, które w tym samym pliku da się ze sobą uzgodnić. Do tego zawężenia werdykt pozostaje czerwony.' 10.5 $false $false $DARK 2 8

    try {
        $props = $doc.BuiltInDocumentProperties
        $props.Item(1).Value = 'Audyt Ryzyk Greenwashingu: Grupa Kapitalowa KRUK'
    } catch {}

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
