#Requires -Version 5.1
# Raport audytu witryny https://mowi.com/pl/ zgodnie z INSTRUKCJA_GENEROWANIA_RAPORTOW.md
$ErrorActionPreference = 'Stop'
$finalPath = Join-Path $PSScriptRoot 'Raport_Audytu_Greenwashingu_MOWI_POLAND.docx'
$outPath = Join-Path $env:TEMP 'Raport_Audytu_Greenwashingu_MOWI_POLAND.docx'

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
            $pic = $ht.Cell(1, 1).Range.InlineShapes.AddPicture($logo)
            $pic.LockAspectRatio = $true
            $pic.Width = $word.CentimetersToPoints(4.2)
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
    $footer.Range.Text = 'ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Mowi Poland  ·  21 października 2026 r.'
    $footer.Range.Font.Name = 'Arial'
    $footer.Range.Font.Size = 8
    $footer.Range.Font.Color = $MUTED

    $word.ActiveWindow.ActivePane.View.SeekView = 0
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)

    Add-P 'ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO' 9 $true $false $ACCENT 0 2
    Add-P 'Audyt Ryzyk Greenwashingu: MOWI POLAND' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja komunikacji opublikowanej na https://mowi.com/pl/ (strona główna, O nas, strategia, struktura wraz z Mowi Feed, Produkty, hodowla łososia, identyfikowalność, smak i zdrowie, jakość, bezpieczeństwo żywności, Zrównoważony rozwój wraz ze zobowiązaniami, certyfikatami i partnerstwami, kariera oraz artykuł „Białko przyjazne dla klimatu” z 11 marca 2019 r., nadal dostępny pod tym adresem). Stan zebrania i data audytu: 21 października 2026 r. Ocena wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE (UCPD), ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (UPNPR), ustawy o ochronie konkurencji i konsumentów, ustawy o zwalczaniu nieuczciwej konkurencji (UZNK), rozporządzenia (UE) nr 1169/2011 oraz rozporządzenia (UE) 2018/848.' 10.5 $false $true $MUTED 0 8

    $shot = Join-Path $PSScriptRoot 'zrzut_mowi_pl.png'
    if (Test-Path -LiteralPath $shot) {
        $sel.ParagraphFormat.Alignment = 1
        $shape = $sel.InlineShapes.AddPicture($shot)
        $shape.LockAspectRatio = $true
        $shape.Width = $word.CentimetersToPoints(15.4)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.Alignment = 0
    }
    Add-P 'Zrzut górnej części strony https://mowi.com/pl/ z 21 października 2026 r. Hasło wiodące: „W Mowi zapewniamy naszym klientom 8 milionów posiłków dziennie”.' 8.5 $false $true $MUTED 2 10

    $meta = @(
        @('Badany podmiot', 'Mowi Poland S.A., Duninowo 39, 76-270 Ustka (adres na stronie Kontakt), telefon +48 59 810 07 21. Spółka należy do grupy Mowi ASA, adres korespondencyjny: PO Box 4102 Sandviken, 5835 Bergen, Norwegia; adres do odwiedzin: Sandviksboder 77AB, 5035 Bergen. Witryna /pl/ jest polskojęzycznym kanałem grupy.'),
        @('Przedmiot audytu', 'Treść publiczna witryny mowi.com/pl/ zebrana 21 października 2026 r.: działy wymienione w podtytule oraz artykuł o śladzie łososia z 11 marca 2019 r., bo pozostaje opublikowany. Oceniono zdania widoczne dla odbiorcy. Nie oceniano raportu rocznego ani kwestionariusza CDP, jeśli na stronie jest tylko odesłanie, bez cytatu liczby.'),
        @('Sektor działalności', 'Hodowla, pasza, przetwórstwo i sprzedaż łososia oraz innych owoców morza (żywność). Kanał detaliczny, HoReCa i przemysłowy. Odbiorca witryny: konsument i klient biznesowy.'),
        @('Reżim prawny', 'Dyrektywa (UE) 2024/825 (ECGT), stosowanie od 27.09.2026, załącznik I pkt 4a, 4b i 10a. Na badanych stronach nie ma zdania o neutralności opartego na kompensacji (pkt 4c) ani własnego znaku zrównoważoności grupy (pkt 2a). Dalej: dyrektywa 2005/29/WE, UPNPR art. 5 i 6, ustawa o ochronie konkurencji i konsumentów art. 24 i 106, UZNK art. 16 ust. 3 i art. 18, rozporządzenie (UE) nr 1169/2011 art. 7, rozporządzenie (UE) 2018/848 art. 30, ustawa o rolnictwie ekologicznym (Dz.U. 2022 poz. 1370) oraz ISO 14067 jako wzorzec liczby śladu produktu. Badane strony nie używają zwrotu „wolne od wylesiania”, więc EUDR nie jest podstawą ustalenia.'),
        @('Data audytu', '21 października 2026 r.'),
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
        'Witryna zostawia ogólne hasła „przyjazny dla klimatu”, „zrównoważony” i „niski ślad” bez EU Ecolabel i bez urzędowego znaku ISO 14024 typu I w tym samym miejscu. Kafelek „100% zbiorów certyfikowanych w sposób zrównoważony” mówi o całym zbiorze, a jedyny udział ASC podany na tej witrynie to 50% hodowli na koniec 2021 r. Limit krajowy wszy i „minimalny wpływ na bentos zgodnie z przepisami krajowymi” stoją w dziale zrównoważenia jako cecha oferty.',
        'Osobno witryna porównuje żywność z oceanu z żywnością i mięsem lądowym. Artykuł z 11 marca 2019 r., nadal dostępny, podaje ślad łososia hodowlanego 2,9 kg CO2e na kg wobec 30 kg dla wołowiny i 5,9 kg dla wieprzowiny, bez metody, roku badania i granic liczby. Dwie podstrony klimatu podają inne cele na 2030 r.',
        'Zakazy ECGT stosuje się od 27 września 2026 r. Treść strony internetowej nie korzysta z taryfy dla starych zapasów. UPNPR art. 5 i 6 obowiązuje niezależnie od daty polskiej ustawy. Utrzymanie tych zdań naraża wydawcę na postępowanie Prezesa UOKiK, z karą do 10% obrotu za praktykę, oraz na roszczenie z reklamy porównawczej. Na badanych stronach nie ma zdania o neutralności opartego na kompensacji.'
    ) (Rgb 255 241 242) $DANGER

    Add-H '1. Analiza Wstępna'
    Add-P 'Mowi Poland S.A. prowadzi w Ustce zakład przetwórstwa i jest częścią Mowi ASA. Witryna opisuje grupę jako największego producenta łososia atlantyckiego, ze zbiorem 475 000 ton w 2023 r. i udziałem około 20% rynku, a na tej samej stronie podaje też 501 500 ton zbioru w 2023 r. Strona jest kanałem wobec konsumenta oraz klienta handlu i HoReCa. Narracja łączy skalę („8 milionów posiłków dziennie”), wizję „Leading the Blue Revolution” i obietnicę żywności zdrowej dla ludzi oraz dobrej dla planety.' 10.5 $false $false $DARK 0 8
    Add-P 'Hasło „przyjazny dla klimatu” stoi na stronie Produktów, na stronie hodowli, w cytacie prezesa na stronie O nas i w dziale kariery. Obok niego witryna stawia porównanie „niebieskiej żywności” z żywnością lądową, kafelek „100% wielkości zbiorów certyfikowanych w sposób zrównoważony” oraz stopień najwyższy „najbardziej zrównoważony producent białka”. Certyfikaty ASC, GLOBALG.A.P. i BAP są znakami stron trzecich. Nie są EU Ecolabel ani urzędowym znakiem ISO 14024 typu I i nie uprawniają do ogólnego przymiotnika „zrównoważony” wobec łososia albo całej działalności.' 10.5 $false $false $DARK 0 8
    Add-P 'Liczby na tej samej witrynie nie tworzą jednego obrazu. Kafelek klimatu mówi o 35% redukcji zakresów 1, 2 i 3 do 2030 r., a strona zobowiązań o 50,6%, 27,5% i 33,33% wobec bazy 2019. Udział ASC jest podany na koniec 2021 r. Tabela certyfikacji nosi rok 2023. Cele opakowaniowe i sprzętowe mają terminy 2023 i 2025, które w dniu audytu już minęły, a strona nadal używa czasu przyszłego. To wybiórcza informacja nałożona na ogólne hasła środowiskowe.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'
    $findings = @(
        @('01','„Przyjazny dla klimatu”, „zrównoważony”, „niski ślad”','ECGT zał. I pkt 4a · UPNPR art. 5','BARDZO WYSOKIE','Ogólne hasło środowiskowe bez uznanej wysokiej efektywności'),
        @('02','Niebieska żywność i 2,9 kg CO2e wobec wołowiny i wieprzowiny','ECGT zał. I pkt 4b · UZNK art. 16 ust. 3','BARDZO WYSOKIE','Porównanie bez wspólnej jednostki, granic i metody'),
        @('03','„Najbardziej zrównoważony producent białka” (Coller FAIRR)','ECGT zał. I pkt 4a · UPNPR art. 5','BARDZO WYSOKIE','Superlativ całego producenta białka z indeksu ESG'),
        @('04','„100% zbiorów certyfikowanych w sposób zrównoważony” (GSSI)','ECGT zał. I pkt 4a i 4b · UPNPR art. 5 i 6','WYSOKIE','Całość zbioru; ASC podane jako 50% hodowli na koniec 2021 r.'),
        @('05','„0% witryn powyżej limitu”; „minimalny wpływ na bentos”','ECGT zał. I pkt 10a · UPNPR art. 5','WYSOKIE','Zgodność z limitem krajowym jako cecha zrównoważenia'),
        @('06','35% do 2030 r. wobec 50,6% / 27,5% / 33,33%; energia odnawialna','UPNPR art. 5 i 6 · przyszła efektywność','WYSOKIE','Dwa zestawy celów i brak udziału energii'),
        @('07','Opakowania i sprzęt „do 2025” i „do 2023”; kompostowanie','UPNPR art. 5 i 6 · ECGT zał. I pkt 4b','WYSOKIE','Terminy minęły, wyniku brak; kafelek wody szerszy niż cel'),
        @('08','Pasza „ekologiczna” i „organiczna”; fabryka „przyjazna dla środowiska”','Rozp. 2018/848 art. 30 · ECGT zał. I pkt 4a','WYSOKIE','Termin zastrzeżony bez numeru; ogólne hasło o zakładzie'),
        @('09','„100% soi” i „100% surowców morskich” oraz „lub równoważny”','UPNPR art. 5 i 6','WYSOKIE','Absoluty bez roku, masy i nazwy drugiego standardu'),
        @('10','GFSI przy zdaniu o klimacie; zły nagłówek GSSI','UPNPR art. 5','ŚREDNIE','Bezpieczeństwo żywności przy celu klimatycznym'),
        @('11','Astaksantyna „naturalna”; omega-3 a zawał, udar i nagła śmierć','Rozp. 1169/2011 art. 7 · UPNPR art. 5','ŚREDNIE','Pochodzenie i skutek chorobowy bez dawki w porcji'),
        @('12','Dwie tony zbioru 2023, dwa zatrudnienia, 8 milionów posiłków','UPNPR art. 5 i 6','ŚREDNIE','Skala grupy bez roku, zakresu i definicji porcji'),
        @('13','„Oficjalna płaca wystarczająca”; „brak znaczących różnic”','UPNPR art. 5 i 6','ŚREDNIE','Wynik społeczny bez roku, źródła stawki i luki')
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

    Add-Case 'PRZYPADEK 01: Ogólne hasła „przyjazny dla klimatu” i „zrównoważony”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'Produkty: „Naszym celem jest dostarczanie najwyższej jakości łososia i inspirowanie do zdrowego i przyjaznego dla klimatu stylu życia”.',
            'Hodowla łososia: „Hodowla łososia to produkcja żywności przyjazna dla klimatu.” Oraz: „Owoce morza oferują niski ślad węglowy, wysoką wydajność zatrzymywania energii i białka oraz niski ślad wodny.”',
            'O nas, cytat prezesa: „Mowi uwalnia potencjał oceanu do produkcji zdrowej i przyjaznej dla klimatu żywności dla rosnącej populacji świata”. Ta sama strona: „minimalny wpływ na środowisko”.',
            'Strategia: „pożywna i zrównoważona żywność” oraz „wywierać pozytywny wpływ na naszą planetę”. Kariera: „zdrowych i przyjaznych dla klimatu owoców morza”. Strona główna: „ochronie planety dla przyszłych pokoleń”. Strona „Czy hodowla łososia jest zrównoważona?”: farmy „działają w sposób przyjazny dla środowiska”.'
        ) `
        @'
Hasła „przyjazny dla klimatu”, „zrównoważony” i „niski ślad” stoją na stronie Produktów, na stronie hodowli, w cytacie prezesa i w karierze. Nie wskazują partii, roku ani wskaźnika. Takie ogólne hasło środowiskowe wolno zostawić tylko wtedy, gdy w tym samym miejscu jest EU Ecolabel albo urzędowy znak ISO 14024 typu I. Certyfikat ASC, BAP albo GLOBALG.A.P. tego warunku nie spełnia. Dopisek na podstronie certyfikatów nie poprawia hasła na stronie Produktów. Zdania o „pozytywnym wpływie na planetę” i o „ochronie planety” nie mówią o kompensacji. Nie ma w nich offsetu ani neutralności klimatycznej.
'@ `
        'Zdjąć z witryny /pl/ przymiotniki „przyjazny dla klimatu”, „zrównoważony” jako określenie łososia, żywności, diety albo stylu życia, „niski ślad węglowy”, „niski ślad wodny”, „minimalny wpływ na środowisko”, „przyjazny dla środowiska” i „pozytywny wpływ na planetę”, dopóki w tym samym widoku nie ma partii albo zakładu, wskaźnika, roku i granic liczby. Wizji „Leading the Blue Revolution” nie używać jako dowodu tej efektywności.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Łosoś wędzony MOWI, partia [kod], zakład [nazwa], zbiór [rok]. Ślad tej partii: [X] kg CO2e na 1 kg produktu, metoda ISO 14067, granice od paszy do bramy zakładu, dane za [rok].”'

    Add-Case 'PRZYPADEK 02: Porównanie żywności z oceanu z żywnością lądową i liczba 2,9 kg' `
        'Podstawa prawna: ECGT, załącznik I pkt 4b · UZNK art. 16 ust. 3 i art. 18 · UPNPR art. 5 i 6 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'Zrównoważony rozwój: „Niebieska żywność ma niższą emisję gazów cieplarnianych niż żywność pochodzenia lądowego (BFA, 2021). Zmiana diety w kierunku zwiększonej konsumpcji owoców morza jest uznawana za część rozwiązania problemu zmian klimatu.” Oraz: „Niebieska żywność charakteryzuje się niższym zużyciem wody słodkiej i gruntów w porównaniu z mięsem lądowym.”',
            'Ta sama strona: „Zwiększone spożycie żywności z oceanu może zmniejszyć spożycie mięsa lądowego, w konsekwencji zmniejszając liczbę chorób przewlekłych związanych z dietą, takich jak nadciśnienie, otyłość i niektóre rodzaje raka.” Oraz: „Liczba ludności na świecie ma wzrosnąć z 7,7 miliarda obecnie do 9,7 miliarda w 2050 roku.”',
            'Hodowla łososia: „Żywność pozyskiwana z oceanu ma niższą emisję gazów cieplarnianych (GHG) niż żywność pochodzenia lądowego, a zatem hodowla łososia może pomóc wypełnić lukę między potencjałem oceanu a produkcją zrównoważonej żywności”. Certyfikaty: „Żywność z oceanu jest kluczem do zapewnienia pożywnej żywności o mniejszym wpływie na klimat niż produkcja żywności na lądzie. Jest to potrójna wygrana: dla ludzi, dla planety i dla gospodarki.”',
            'Produkty: „niebieska żywność pochodząca z oceanu, taka jak łosoś, ma średnio znacznie większe wartości odżywcze niż żywność pochodząca od zwierząt lądowych, a wiele z nich ma również mniejszy wpływ na środowisko.” Artykuł „Białko przyjazne dla klimatu”, 11 marca 2019 r.: „łosoś hodowlany ma ślad węglowy na poziomie 2,9 kg / CO2e na kg, co stanowi mniej niż jedną dziesiątą produkcji wołowiny (30 kg) i około połowę wieprzowiny (5,9 kg).”'
        ) `
        @'
Na stronie Zrównoważony rozwój czytelnik wynosi, że żywność z oceanu ma niższą emisję, mniejsze zużycie wody i gruntów oraz że jest częścią rozwiązania klimatu. Strona hodowli przenosi to zdanie na hodowlę łososia Mowi. Artykuł z 11 marca 2019 r. nadal podaje 2,9 kg CO2e na kg wobec 30 kg wołowiny i 5,9 kg wieprzowiny. Brakuje wspólnej jednostki (kilogram produktu, białko albo kilokaloria), granic liczby, roku badania i udziału paszy. Porównanie dwóch żywności wolno zostawić tylko wtedy, gdy w tym samym miejscu da się sprawdzić te same cechy, tą samą metodą, dla produktów o tym samym przeznaczeniu. Kategoria „żywność lądowa” i „mięso lądowe” łączy wołowinę z drobiem. Jedno zdanie na stronie Produktów mówi, że „wiele” produktów oceanicznych ma mniejszy wpływ. Zdania obok mówią już o całej żywności z oceanu. „7,7 miliarda obecnie” stoi przy źródle BFA 2021, a strona jest czytana 21 października 2026 r.
'@ `
        'Zdjąć porównanie „ocean kontra ląd”, w tym liczby 2,9 kg, 30 kg i 5,9 kg, ze stron Zrównoważony rozwój, hodowli, certyfikatów, Produktów i z artykułu z 11 marca 2019 r. Artykuł albo usunąć z publicznego kanału, albo zostawić z datą i bez tych liczb. Jeśli porównanie ma wrócić, obie strony muszą mieć w jednym akapicie tę samą jednostkę, ten sam rok, te same granice i tę samą metodę, z nazwanym produktem po drugiej stronie, nie z całą kategorią „mięso lądowe”. Liczbę ludności opatrzyć rokiem źródła i nie pisać „obecnie”, gdy dana nie jest z roku publikacji strony.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Filet z łososia hodowlanego MOWI, partia [kod]. Ślad węglowy tej partii w [roku] wynosi [X] kg CO2e na 1 kg produktu jadalnego. Metoda ISO 14067. W liczbie jest karma, chów i przetworzenie w zakładzie [nazwa].”'

    Add-Case 'PRZYPADEK 03: „Najbardziej zrównoważony producent białka”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'O nas, nagłówek: „Mowi, największy na świecie dostawca łososia hodowlanego i najbardziej zrównoważony producent białka”.',
            'O nas: „Mowi został uznany za najbardziej zrównoważonego producenta białka przez Coller FAIRR Protein Producer Index piąty rok z rzędu.” Przypis: „MOWI zajęło pierwsze miejsce w rankingu Coller FAIRR Protein Producer Index 2019–2024, kompleksowej ocenie największych producentów białka zwierzęcego pod kątem krytycznych kwestii środowiskowych, społecznych i związanych z zarządzaniem (ESG).”'
        ) `
        @'
Nagłówek strony O nas nazywa Mowi najbardziej zrównoważonym producentem białka. Przypis pod spodem mówi o pierwszym miejscu w Coller FAIRR Protein Producer Index za lata 2019–2024, wśród największych producentów białka zwierzęcego, w ocenie spraw środowiskowych, społecznych i zarządczych. Ogólne hasło „najbardziej zrównoważony” wolno zostawić tylko z EU Ecolabel albo urzędowym znakiem ISO 14024 typu I w tym samym miejscu. Pierwsze miejsce w prywatnym indeksie tego warunku nie spełnia. Nagłówek nie powtarza grupy objętej rankingiem. Przypis nie podaje wagi kryteriów ani tego, czego indeks nie mierzy. Zdanie o wolumenie „największy dostawca”, z rokiem 2023 i udziałem około 20%, jest osobne i tego przypadku nie dotyczy, o ile przy liczbie zostaje rok i jeden uzgodniony wolumen.
'@ `
        'Zdjąć superlatyw „najbardziej zrównoważony” z nagłówka O nas. Jeśli indeks zostaje, opublikować tylko miejsce, lata i grupę objętą rankingiem. Nie stawiać tego zdania obok hasła o łososiu, klimacie ani emisji.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W Coller FAIRR Protein Producer Index za lata [RRRR–RRRR] Mowi zajęło miejsce [n] wśród producentów białka zwierzęcego objętych tym indeksem.”'

    Add-Case 'PRZYPADEK 04: „100% zbiorów certyfikowanych w sposób zrównoważony”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5 i 6 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Zrównoważony rozwój, kafelek: „100% wielkości zbiorów certyfikowanych w sposób zrównoważony zgodnie ze standardem uznanym przez GSSI”. To samo zdanie, bez daty, jest na stronie Zobowiązania i działania.',
            'Nasze certyfikaty: „Polityka certyfikacji Mowi (link) stanowi, że 100% naszego rocznego odłowu łososia, tj. wszystkie jednostki hodowlane Mowi – farmy słodkowodne i morskie – muszą być certyfikowane zgodnie ze standardami zrównoważonego rozwoju uznanymi przez Global Seafood Sustainability Initiative (GSSI): ASC, GSA BAP i/lub GLOBALG.A.P.” Zwrot „(link)” nie jest odnośnikiem.',
            'Ta sama strona: „Mowi pozostaje liderem w zakresie certyfikacji ASC, posiadając do końca 2021 r. łącznie 133 miejsca hodowli certyfikowane na całym świecie. Stanowi to 50% wszystkich naszych hodowli i 33% wszystkich hodowli łososia atlantyckiego z certyfikatem ASC na całym świecie”. Oraz: „cały nasz łosoś MOWI Pure jest i będzie certyfikowany ASC”. Odesłanie: „Zobacz tabelę certyfikacji 2023”.',
            'Jakość: „Minimalnym wymogiem dla operacji hodowlanych jest posiadanie certyfikatu Global Sustainable Seafood Initiative (GSSI).”'
        ) `
        @'
Kafelek „100%” stoi w rzędzie wskaźników i czyta się jak wynik. Strona certyfikatów mówi coś węższego: polityka wymaga, aby odłów był certyfikowany według jednego ze standardów uznanych przez GSSI. „Muszą być” nie jest tym samym co „są”, a polityki nie da się otworzyć spod słowa „(link)”. GSSI porównuje standardy. Nie jest oznakowaniem produktu i nie jest urzędowym znakiem ISO 14024 typu I. Hasło o całym zbiorze wolno oprzeć na udziale, który dotyczy tego zbioru, z rokiem i z nazwą standardu. Jedyny udział ASC na tej witrynie to 50% hodowli na koniec 2021 r., liczony miejscami, nie wolumenem. Linia MOWI Pure jest węższa niż cały zbiór. Strona Jakość nazywa GSSI certyfikatem operacji hodowlanej. Tabela, do której strona odsyła, ma w nazwie rok 2023.
'@ `
        'Na kafelku rozdzielić cel polityki od wyniku. Wynik podać z rokiem, udziałem wolumenu i rozbiciem na ASC, BAP i GLOBALG.A.P. Zdjąć przymiotnik „w sposób zrównoważony”. Podlinkować politykę albo usunąć słowo „(link)”. Zdanie o MOWI Pure zostawić przy tej linii, z numerem certyfikatu i datą ważności. Na stronie Jakość nie pisać „certyfikat GSSI”. Napisać nazwę standardu, który farma ma.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] [X]% wolumenu zbioru łososia Mowi pochodziło z jednostek z certyfikatem ASC nr [numer], [Y]% z BAP i [Z]% z GLOBALG.A.P. Tabela z dnia [data].”'

    Add-Case 'PRZYPADEK 05: Limit krajowy wszy i wpływ na bentos' `
        'Podstawa prawna: ECGT, załącznik I pkt 10a · UPNPR art. 5 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Zrównoważony rozwój, kafelek „Wszy morskie”: „0% witryn powyżej limitu krajowego (średnia miesięczna)”. Na stronie zobowiązań to samo zdanie, bez roku.',
            'Zobowiązania: „100% terenów uprawnych o minimalnym wpływie na bentos (zgodnie z przepisami krajowymi)”.'
        ) `
        @'
Kafelek wszy podaje zero miejsc powyżej limitu krajowego i nie podaje miesiąca ani roku. Zdanie o bentosie samo dopisuje „zgodnie z przepisami krajowymi” i nazywa to minimalnym wpływem. Utrzymanie się w krajowym limicie jest obowiązkiem prowadzenia hodowli. Obowiązku prawnego nie wolno pokazywać jako cechy środowiskowej oferty. Sformułowanie „terenów uprawnych” nie mówi, których miejsc hodowli limit dotyczy. Oba zdania stoją w dziale zrównoważenia, obok klimatu i certyfikatów.
'@ `
        'Z kafelka „Wszy morskie” zdjąć „0%” bez miesiąca i roku. Jeśli liczba zostaje, podać miesiąc, rok i udział miejsc hodowli morskiej. Zdanie o „minimalnym wpływie na bentos” zdjąć albo zastąpić metodą, progiem i rokiem. Nie dopisywać na stronie, że limit jest obowiązkiem prawa.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [miesiąc rok] [X]% miejsc hodowli morskiej Mowi nie przekroczyło krajowego limitu zagęszczenia wszy.”'

    Add-Case 'PRZYPADEK 06: Dwa cele emisji i energia odnawialna bez udziału' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · dyrektywa 2005/29/WE art. 6 (przyszła efektywność środowiskowa) | Poziom ryzyka: WYSOKIE' `
        @(
            'Zrównoważony rozwój, kafelek „Zmiana klimatu”: „35% redukcja emisji gazów cieplarnianych z zakresu 1, 2 i 3 do 2030 r. i o 72% do 2050 r.”',
            'Zobowiązania i działania oraz strona Zmiana klimatu: „Mowi zobowiązuje się do zmniejszenia bezwzględnych emisji gazów cieplarnianych zakresu 1 i 2 o 50,6% do 2030 r. w stosunku do roku bazowego 2019. Mowi zobowiązuje się do zmniejszenia bezwzględnych emisji gazów cieplarnianych zakresu 3 o 27,5% w tych samych ramach czasowych. Mowi zobowiązuje się do zmniejszenia bezwzględnych emisji gazów cieplarnianych z zakresu 3 FLAG o 33,33% w tym samym okresie.”',
            'Strategia: „pozyskiwanie coraz większej ilości energii elektrycznej ze źródeł odnawialnych”. Strona Zmiana klimatu, wśród wskaźników: „% energii elektrycznej ze źródeł odnawialnych”. Ta sama strona odsyła do raportu rocznego, polityki klimatycznej i kwestionariusza CDP za 2023 r.'
        ) `
        @'
Kafelek klimatu mówi o spadku o 35% dla zakresów 1, 2 i 3 łącznie do 2030 r. i o 72% do 2050 r. Strona zobowiązań podaje inne liczby na ten sam rok: 50,6%, 27,5% i 33,33%, wobec 2019 r. Cel na przyszłość wolno opublikować, gdy w tym samym miejscu są mierzalny wskaźnik, termin, zasoby i regularna weryfikacja niezależnego eksperta, a wynik jest dostępny odbiorcy. Dwie podstrony nie dają jednego celu. 72% do 2050 r. nie ma zakresu ani roku bazowego. Zdanie o „coraz większej ilości” energii elektrycznej ze źródeł odnawialnych nie podaje udziału, roku, spółki ani tego, czy chodzi o umowę na energię, czy o gwarancję pochodzenia. Odesłanie do CDP 2023 i do raportu rocznego nie uzgadnia 35% z 50,6%.
'@ `
        'Zostawić jeden zestaw celów, z rokiem bazowym, zakresem i datą przyjęcia. Przy każdym celu podać wynik za ostatni rok, metodykę i nazwę weryfikatora. Drugi zestaw zdjąć z kafelka. Przy energii podać udział procentowy zużycia, rok, spółkę i kraj. Gwarancje pochodzenia, jeśli są, wykazuje się osobno od umowy na dostawę. Tego rozróżnienia nie drukować jako lekcji. Ma wynikać z tego, którą liczbę strona podaje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Cel grupowy przyjęty [data]: spadek emisji bezwzględnych zakresu 1 i 2 o 50,6% do 2030 r. wobec 2019 r.; zakresu 3 o 27,5%; zakresu 3 FLAG o 33,33%. Wynik za [rok]: zakres 1 i 2 [A]%, zakres 3 [B]%, zakres 3 FLAG [C]%. Metodyka [nazwa], weryfikator [nazwa]. W [roku] energia elektryczna z umowy na dostawę ze źródeł odnawialnych stanowiła [X]% zużycia [spółka, kraj].”'

    Add-Case 'PRZYPADEK 07: Cele opakowań, sprzętu i wody po upływie terminu' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · ECGT, załącznik I pkt 4b | Poziom ryzyka: WYSOKIE' `
        @(
            'Zobowiązania: „Do 2025 r. 100% naszych plastikowych opakowań będzie nadawało się do ponownego użycia, recyklingu lub kompostowania. Do 2025 r. co najmniej 25% opakowań z tworzyw sztucznych będzie pochodzić z recyklingu. Do 2023 r. cały plastikowy sprzęt rolniczy (sieci, liny, rury zasilające) będzie ponownie wykorzystywany lub poddawany recyklingowi.”',
            'Ta sama strona: „Osiągnięcie do 2025 r. redukcji zużycia wody o 10% w naszych zakładach przetwórczych zlokalizowanych w regionach o średnim i wysokim ryzyku niedoboru wody”. Przypis: rok odniesienia 2018. „Do 2025 r. zero odpadów składowanych na wysypiskach w naszych zakładach przetwórczych. Do 2023 r. minimum 50% naszego stada z monitorowaniem dobrostanu w czasie rzeczywistym.”',
            'Kafelek na stronie Zrównoważony rozwój: „10% zmniejszenie zużycia wody w naszych zakładach przetwórczych do 2025 r.”, bez regionów niedoboru i bez roku 2018. Ten sam rząd kafelków: „100% naszych plastikowych opakowań będzie nadawało się do ponownego użycia, recyklingu lub kompostowania do 2025 r.”'
        ) `
        @'
Dzień audytu jest 21 października 2026 r. Terminy 2023 i 2025 minęły. Strona nadal używa czasu przyszłego i nie podaje, czy cel został wykonany, w jakim procencie i dla których zakładów. Cel po terminie, zapisany jako cecha przyszła, czyta się jak cecha aktualna. Spójnik „lub” przy opakowaniu pozwala uznać cel za spełniony przez samą teoretyczną przydatność do recyklingu albo przez kompostowanie. Strona nie podaje normy, kraju zbiórki ani tego, czy chodzi o kompostownię przemysłową. Kafelek wody mówi o zakładach przetwórczych w ogóle. Strona zobowiązań zawęża ten sam cel do regionów o średnim i wysokim ryzyku niedoboru wody, wobec 2018 r. Węższy cel jest na kafelku pokazany jako szerszy.
'@ `
        'Przy każdym celu z lat 2023 i 2025 wstawić wynik albo zdjąć zdanie. Przy opakowaniu rozdzielić ponowne użycie, recykling i kompostowanie. Przy recyklingu podać kraj i warunek opróżnienia. Przy kompostowaniu podać normę i rodzaj instalacji. Na kafelku wody powtórzyć regiony niedoboru i rok bazowy 2018. Nie zostawiać czasu przyszłego po dacie celu.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] [X]% masy opakowań z tworzyw wprowadzonych przez [spółka] nadawało się do recyklingu w [kraj], po opróżnieniu, w strumieniu [nazwa]. Udział recyklatu w tych opakowaniach wyniósł [Y]% masy. Opakowań kompostowalnych według [norma], w kompostowni przemysłowej, było [Z]% masy. Zużycie wody w zakładach przetwórczych w regionach o średnim i wysokim ryzyku niedoboru wody było w [roku] o [W]% niższe niż w 2018 r.”'

    Add-Case 'PRZYPADEK 08: Pasza „ekologiczna” i „organiczna” oraz fabryka „przyjazna dla środowiska”' `
        'Podstawa prawna: rozporządzenie (UE) 2018/848 art. 30 · ustawa o rolnictwie ekologicznym (Dz.U. 2022 poz. 1370) · ECGT, załącznik I pkt 4a · UPNPR art. 5 | Poziom ryzyka: WYSOKIE (termin zastrzeżony oraz czarna lista)' `
        @(
            'Mowi Feed: „Sprzedajemy karmę dla ryb w Norwegii, Szkocji, Irlandii, Islandii i na Wyspach Owczych – dla produkcji zgodnie z konwencjonalnymi i ekologicznymi standardami.”',
            'Ta sama strona: „Sprzedajemy karmy dla wszystkich etapów cyklu życia łososia większego niż 2-3 g, rosnącego w wodzie słodkiej i morskiej, do produkcji zgodnej ze standardami konwencjonalnymi i organicznymi.”',
            '„Te specjalnie zbudowane fabryki zostały zaprojektowane specjalnie do produkcji karm o wysokiej gęstości odżywczej dla ryb hodowlanych, w tym łososia i pstrąga, i są wyposażone we wszystkie urządzenia, które zwiększają wydajność i efektywność kosztową w sposób przyjazny dla środowiska.”'
        ) `
        @'
Strona Mowi Feed mówi, że karma idzie do produkcji „zgodnie z konwencjonalnymi i ekologicznymi standardami”, a niżej „ze standardami konwencjonalnymi i organicznymi”. W tym akapicie nie ma nazwy paszy, jednostki certyfikującej, numeru ani kodu. Wyrazów „ekologiczny” i „organic” przy paszy wolno użyć tylko dla paszy zgodnej z rozporządzeniem 2018/848, z tymi danymi w tym samym miejscu. Zdanie o fabrykach mówi, że urządzenia zwiększają wydajność „w sposób przyjazny dla środowiska”. To ogólne hasło o zakładzie, bez wskaźnika i roku. Sprawność linii nie opisuje surowców, z których karma powstaje.
'@ `
        'Przy paszy ekologicznej wstawić w tym samym akapicie nazwę, jednostkę, numer i kod albo usunąć „ekologicznymi” i „organicznymi” i zostawić pasze konwencjonalne. Zdania, że pozostałe pasze ekologiczne nie są, nie dopisywać: pierwsze zdanie już rozdziela konwencjonalne od ekologicznych, a „ekologiczna” zostaje tylko przy nazwanej paszy. Zdjąć „w sposób przyjazny dla środowiska”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Mowi Feed sprzedaje pasze konwencjonalne w Norwegii, Szkocji, Irlandii, Islandii i na Wyspach Owczych. Pasza [nazwa handlowa] jest paszą ekologiczną: jednostka [nazwa], numer [numer], kod [kod].”'

    Add-Case 'PRZYPADEK 09: „100% soi” i „100% surowców morskich” z dopiskiem „lub równoważny”' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE' `
        @(
            'Zobowiązania: „100% identyfikowalność surowców paszowych. 100% surowców morskich posiada certyfikat (Marin Trust lub równoważny). 100% soi jest certyfikowane (ProTerra lub odpowiednik). W kierunku niższego śladu węglowego surowców paszowych.”',
            'Kafelek „Zrównoważona pasza” skraca to do: „100% identyfikowalność surowców paszowych”.'
        ) `
        @'
Cztery zdania stoją bez roku, bez udziału masy i bez spółki, której dotyczą. „Lub równoważny” i „lub odpowiednik” nie nazywają drugiego standardu, więc nie da się sprawdzić, co poza ProTerra i Marin Trust wchodzi do licznika 100%. „100% identyfikowalność” nie podaje systemu ani głębokości, do regionu, plantacji albo partii. Kierunek „niższego śladu węglowego surowców” nie ma liczby, jednostki, metody ani roku. W tym miejscu witryna nie pisze „wolne od wylesiania”.
'@ `
        'Przy soi i surowcach morskich podać rok, udział masy, spółkę i nazwę standardu. Zdjąć „lub równoważny” i „lub odpowiednik”. Jeśli w udziale jest drugi standard, wpisać jego nazwę w tym samym zdaniu. Hasło „100% identyfikowalność” albo zdjąć, albo w tym samym miejscu podać poziom i system. Hasło o niższym śladzie surowców albo uzupełnić liczbą, jednostką, metodą i rokiem, albo usunąć.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] [X]% masy soi kupionej do pasz Mowi Feed dla [spółka] miało certyfikat ProTerra nr [numer]. [Y]% masy surowców morskich miało certyfikat Marin Trust nr [numer].”'

    Add-Case 'PRZYPADEK 10: Bezpieczeństwo żywności przy zdaniu o klimacie' `
        'Podstawa prawna: UPNPR art. 5 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Nasze certyfikaty, zdanie po „potrójnej wygranej”: „Kluczem do osiągnięcia naszego celu jest niezależna certyfikacja naszych działań w zakresie zrównoważonego rozwoju i bezpieczeństwa żywności”.',
            'Kluczowe partnerstwa, nagłówek: „Globalna inicjatywa na rzecz bezpieczeństwa żywności”. Akapit: „Nadal wspieramy inicjatywę Global Sustainable Seafood Initiative, która odgrywa ważną rolę w zapewnianiu jasności w zakresie certyfikacji owoców morza.” Adres podany na stronie: ourgssi.org.',
            'Jakość: „100% naszych zakładów przetwórczych posiada uznany standard Global Food Safety Initiative (GFSI).”'
        ) `
        @'
Strona certyfikatów stawia certyfikację zrównoważenia i bezpieczeństwa żywności jako klucz do celu, który zdanie wyżej nazywa mniejszym wpływem na klimat. FSSC, BRC, IFS i GFSI potwierdzają system bezpieczeństwa żywności zakładu. Nie potwierdzają śladu ani paszy. Na stronie partnerstw nagłówek mówi o inicjatywie bezpieczeństwa żywności, a akapit i adres ourgssi.org opisują Global Sustainable Seafood Initiative. To dwa różne ciała. Strona certyfikatów rozdziela je poprawnie. Strona partnerstw je miesza. Sam fakt certyfikatu GFSI na stronie Jakość da się zostawić przy zakładzie, z numerem i datą.
'@ `
        'Rozdzielić na witrynie blok bezpieczeństwa żywności i blok standardów hodowli. Na stronie partnerstw poprawić nagłówek na Global Sustainable Seafood Initiative albo przenieść GFSI do bloku bezpieczeństwa. Zdania, że GSSI i GFSI są dwoma inicjatywami, nie drukować. Przy certyfikacie zakładu podać nazwę, numer i datę. Nie stawiać go obok hasła o klimacie ani „dla planety”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Zakład [nazwa] ma certyfikat FSSC 22000 nr [numer], ważny do [data]. Certyfikat dotyczy systemu bezpieczeństwa żywności tego zakładu.”'

    Add-Case 'PRZYPADEK 11: „Naturalna” astaksantyna i skutki chorobowe omega-3' `
        'Podstawa prawna: rozporządzenie (UE) nr 1169/2011 art. 7 · UPNPR art. 5 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Hodowla łososia i Produkty: „Karma dla łososia zawiera naturalne związki, takie jak astaksantyna, które nadają łososiowi różowy odcień.”',
            'Smak i zdrowie: „Kwasy tłuszczowe omega-3 pomagają utrzymać zdrowe serce poprzez obniżenie ciśnienia krwi, a także zmniejszają ryzyko nagłej śmierci, zawału serca, udaru mózgu i choroby wieńcowej serca.” Oraz, przy kafelku EPA i DHA: „Morskie kwasy Omega-3 zapobiegają rozwojowi chorób sercowo-naczyniowych.”',
            'Ta sama strona, kafelek: „EPA I DHA. 480% RDI”. Nagłówek: „Prawdziwa superżywność”. Produkty: „Łosoś jest ulubieńcem konsumentów, superżywnością”.'
        ) `
        @'
Odpowiedź na pytanie o barwę nazywa astaksantynę w karmie związkiem naturalnym i nie podaje, czy pochodzi z glonów, z drożdży, czy z syntezy. Na stronie Smak i zdrowie kwasy omega-3 zmniejszają ryzyko nagłej śmierci, zawału, udaru i choroby wieńcowej, a niżej zapobiegają chorobom sercowo-naczyniowym. W tym samym miejscu nie ma dawki EPA i DHA w nazwanej porcji produktu Mowi. Kafelek „480% RDI” nie podaje gramatury porcji ani nazwy spożycia referencyjnego. Przypisy, w tym materiał amerykańskiej agencji, nie zastępują dawki w akapicie. „Superżywność” nie jest nazwą składnika ani dawką. Informacja o żywności nie może przypisywać produktowi działania, którego odbiorca nie sprawdzi na porcji.
'@ `
        'Przy astaksantynie podać pochodzenie albo zdjąć „naturalne”. Przy omega-3 zostawić dawkę EPA i DHA w nazwanej porcji. Listę chorób usunąć. Zostaje tylko wtedy, gdy w tym samym miejscu jest oświadczenie zdrowotne dopuszczone dla tego produktu, z dawką i warunkami użycia. Zdjąć „superżywność” i „480% RDI” bez gramatury porcji.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Różowa barwa tego łososia pochodzi z astaksantyny w karmie. W partii [kod] astaksantyna pochodzi z [źródło]. Porcja [gramatura] dostarcza [mg] EPA i DHA.”'

    Add-Case 'PRZYPADEK 12: Dwie wielkości zbioru, dwa zatrudnienia i posiłki bez definicji' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona główna i O nas: „W Mowi zapewniamy naszym klientom 8 milionów posiłków dziennie”.',
            'O nas: „największym na świecie producentem łososia atlantyckiego z wolumenem zbiorów wynoszącym 475 000 ton w 2023 r., co odpowiada globalnemu udziałowi w rynku wynoszącemu około 20%”. Kafelek na tej samej stronie: „501 500 ton. Wielkość zbiorów w 2023 r.”',
            'O nas, kafelek: „13 800. Pracownicy”. Akapit niżej: „zatrudniamy ponad 11 000 osób”. Strona główna: „zatrudnia ponad 11 000 osób na całym świecie”.'
        ) `
        @'
Na stronie O nas zbiór za 2023 r. jest podany dwa razy: 475 000 ton przy udziale około 20% oraz 501 500 ton jako wielkość zbiorów. Strona nie mówi, która liczba jest łososiem atlantyckim, a która obejmuje szerszy zakres. Zatrudnienie w tym samym widoku to 13 800 oraz „ponad 11 000”. „8 milionów posiłków dziennie” nie ma roku ani gramatury porcji. To dane ilościowe o skali tej samej grupy. Odbiorca nie wie, która liczba jest aktualna i czego dotyczy. Na tych liczbach opiera się reszta narracji, w tym zdania o wyżywieniu populacji.
'@ `
        'Zostawić jedną wielkość zbioru za 2023 r., z zakresem w tym samym zdaniu (łosoś atlantycki albo cała grupa). Zostawić jedną liczbę zatrudnienia, z datą. Przy posiłkach podać rok i gramaturę porcji. Drugiej liczby nie zostawiać jako stanu bieżącego.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2023 r. zbiór łososia atlantyckiego Mowi wyniósł [N] ton. Zatrudnienie na koniec [roku]: [K] osób. W [roku] produkcja odpowiadała [P] porcjom dziennie, przy porcji [g] produktu jadalnego.”'

    Add-Case 'PRZYPADEK 13: Płaca wystarczająca i „brak znaczących różnic” bez metody' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Bezpieczne i znaczące miejsca pracy: „Żaden pracownik nie otrzymuje wynagrodzenia niższego niż oficjalna krajowa płaca wystarczająca na utrzymanie wskazana dla danej lokalizacji.”',
            'Ta sama strona: „Regularnie mapujemy również wynagrodzenia ze względu na płeć. Nasza ogólna analiza nie wykazuje znaczących różnic w wynagrodzeniach kobiet i mężczyzn.”',
            'Zobowiązania: „30% kobiet na stanowiskach kierowniczych do 2030 r. Stosunek płci pracowników 50/50 do 2030 r. Brak znaczących różnic między płciami w globalnym badaniu zaangażowania pracowników.”'
        ) `
        @'
Zdanie o płacy wystarczającej jest mocniejsze niż płaca minimalna, ale witryna nie podaje, kto publikuje „oficjalną krajową płacę wystarczającą” w każdym kraju działania, ani roku porównania. „Nie wykazuje znaczących różnic” nie podaje luki, progu istotności, roku ani tego, czy analiza obejmuje zakład w Ustce. Cele 30% i 50/50 są zapisane do 2030 r. Stoją obok zdania o braku różnic, które czyta się jak wynik już osiągnięty. Brakuje metody i roku.
'@ `
        'Przy płacy podać źródło stawki, rok i zakres pracowników. Przy płci podać dwie luki, rok i grupę pracowników albo usunąć „nie wykazuje”. Cele 2030 zostawić jako cele, z wynikiem za ostatni rok obok.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] porównano wynagrodzenie pracowników [spółki] ze stawką [nazwa źródła] dla [lokalizacja]. Luka płacowa kobiet i mężczyzn, nieskorygowana, wyniosła [X]%. Luka skorygowana o [czynniki] wyniosła [Y]%. Cel 30% kobiet na stanowiskach kierowniczych dotyczy 2030 r. Wynik za [rok]: [Z]%.”'

    Add-H '4. Porównanie z żywnością lądową: osobny test porównywalności i sprawdzalności'
    Add-P 'Rozdział 3 opisuje twierdzenia. Ten rozdział sprawdza tylko porównywalność i sprawdzalność zestawienia, które witryna sama robi między żywnością z oceanu, łososiem hodowlanym a żywnością i mięsem lądowym.' 10.5 $false $false $DARK 0 6

    $comp = @(
        @('Emisja. „Niebieska żywność ma niższą emisję gazów cieplarnianych niż żywność pochodzenia lądowego (BFA, 2021).” Hodowla łososia przenosi to zdanie na własną hodowlę. Artykuł z 11 marca 2019 r.: łosoś hodowlany 2,9 kg CO2e na kg, wołowina 30 kg, wieprzowina 5,9 kg.', 'BFA 2021 jest badaniem kategorii. Kwalifikator „wiele z nich” jest tylko w jednym zdaniu strony Produktów. Artykuł nie podaje, czyja jest liczba 2,9 kg, z którego roku jest badanie, jaka jest jednostka funkcjonalna i czy karma jest w granicach. Wołowina i wieprzowina nie są „żywnością lądową” w ogóle. Drób, o skrajnie innym śladzie, w tych liczbach nie występuje. Porównanie nie spełnia art. 16 ust. 3 UZNK: nie da się go sprawdzić na stronie i nie zestawia porównywalnych produktów.'),
        @('Woda i grunty. „Niebieska żywność charakteryzuje się niższym zużyciem wody słodkiej i gruntów w porównaniu z mięsem lądowym.” Hodowla: „niski ślad wodny”.', 'Brak liczby, roku i jednostki. „Mięso lądowe” nie nazywa wołowiny, wieprzowiny ani drobiu. Ślad wody łososia hodowlanego zależy od paszy roślinnej. Strona w tym bloku paszy nie wlicza. Cel wody na stronie zobowiązań (minus 10% do 2025 r. wobec 2018 r., tylko regiony niedoboru) jest innym przedmiotem niż ślad wody produktu.'),
        @('Wartość odżywcza i choroby. „Niebieska żywność, taka jak łosoś, ma średnio znacznie większe wartości odżywcze niż żywność pochodząca od zwierząt lądowych”. „Zwiększone spożycie żywności z oceanu może zmniejszyć spożycie mięsa lądowego, w konsekwencji zmniejszając liczbę chorób przewlekłych”, w tym nadciśnienia, otyłości i niektórych nowotworów.', 'Porównanie wartości odżywczych nie podaje składnika, porcji ani produktu po drugiej stronie. Zdanie o chorobach przenosi skutek zmiany diety na „żywność z oceanu” jako kategorię, bez dawki EPA i DHA w porcji Mowi. Źródło BFA 2021 jest badaniem kategorii, nie kartą produktu.')
    )
    $range = $doc.Range($sel.Range.End, $sel.Range.End)
    $ct = $doc.Tables.Add($range, $comp.Count + 1, 2)
    $ct.PreferredWidthType = 3
    $ct.PreferredWidth = $contentW
    $ct.AllowAutoFit = $false
    $ct.Columns.Item(1).Width = $word.CentimetersToPoints(8.6)
    $ct.Columns.Item(2).Width = $word.CentimetersToPoints(9.0)
    Write-Cell $ct.Cell(1, 1) 'Twierdzenie podmiotu' 8 $true $WHITE
    Write-Cell $ct.Cell(1, 2) 'Rzeczywisty stan prawno-środowiskowy' 8 $true $WHITE
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
    Add-P '• 1. Prezes UOKiK' 11 $true $false $NAVY 10 2
    Add-P 'Praktyka naruszająca zbiorowe interesy konsumentów (art. 24 ustawy o ochronie konkurencji i konsumentów) jest zagrożona karą do 10% obrotu z roku poprzedzającego nałożenie kary, za każdą praktykę (art. 106). Zarzuty z 19.01.2026 r. wobec Tchibo i Zary dotyczą ogólnego „eko” i „zrównoważone” oraz certyfikatu czytanego szerzej niż jego zakres. Zarzut wobec Bolta dotyczy energii odnawialnej opartej na certyfikacie atrybutu. Wzorce z witryny /pl/ są z tej samej grupy: ogólne hasło, certyfikat czytany na całość i energia odnawialna bez udziału oraz bez rozróżnienia umowy i gwarancji pochodzenia. ECGT stosuje się od 27.09.2026 r. do treści online. Opóźnienie polskiej ustawy (druk Senatu 825, projektowana data 27.09.2027, akt nieogłoszony) nie wstrzymuje art. 5 i 6 UPNPR. Rozbieżność kafelka „100%” z udziałem ASC z końca 2021 r. jest daną, którą pracownik albo agencja może zgłosić do Prezesa UOKiK z pominięciem drogi wewnętrznej.' 10.5 $false $false $DARK 0 8
    Add-P '• 2. Reklama porównawcza' 11 $true $false $NAVY 8 2
    Add-P 'Zdania o niższej emisji, niższym zużyciu wody i gruntów oraz o mniejszej liczbie chorób wobec „mięsa lądowego” i „żywności pochodzenia lądowego” pozwalają rozpoznać cudzą kategorię żywności. Artykuł 16 ust. 3 UZNK wymaga, aby takie porównanie było rzetelne, sprawdzalne i dotyczyło produktów o tym samym przeznaczeniu oraz cech istotnych i typowych. Artykuł 18 daje konkurentowi roszczenie o zaniechanie i naprawienie szkody. Nazwa jednej spółki po drugiej stronie nie jest potrzebna, gdy kategoria jest wskazana, a liczby 30 kg i 5,9 kg nazywają wołowinę i wieprzowinę.' 10.5 $false $false $DARK 0 8
    Add-P '• 3. Łańcuch klientów i ESRS' 11 $true $false $NAVY 8 2
    Add-P 'Odbiorcą witryny jest także sieć detaliczna i HoReCa. Hasło „przyjazny dla klimatu”, „100% zbiorów certyfikowanych w sposób zrównoważony” i „najbardziej zrównoważony producent białka” może wejść do danych o łańcuchu wartości klienta, który raportuje ESRS E1. Klient, który powtórzy te skróty, przenosi wadę na własny raport. Korekta na mowi.com/pl/ ogranicza to przeniesienie. Nie zastępuje danych pierwotnych, których ta witryna przy haśle nie publikuje.' 10.5 $false $false $DARK 0 8
    Add-P '• 4. Pasza i termin „ekologiczny”' 11 $true $false $NAVY 8 2
    Add-P 'Użycie „ekologicznymi” i „organicznymi” na stronie Mowi Feed, bez jednostki, numeru i kodu w tym samym akapicie, jest ryzykiem z rozporządzenia 2018/848 i z nadzoru nad rolnictwem ekologicznym, osobnym od kary za praktykę rynkową. Certyfikat, jeśli obejmuje tylko część wolumenu, zostaje przy nazwanej paszy. Reszta oferty w tym zdaniu zostaje przy paszy konwencjonalnej.' 10.5 $false $false $DARK 0 8

    Add-H '6. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIASTOWY PULL-OFF (21 października 2026 r.)' @(
        'Zdjąć z witryny /pl/ hasła: „przyjazny dla klimatu”, „zrównoważony” jako określenie łososia, żywności albo stylu życia, „niski ślad węglowy”, „niski ślad wodny”, „minimalny wpływ na środowisko”, „pozytywny wpływ na planetę”, „najbardziej zrównoważony producent białka” oraz porównanie „niższa emisja niż żywność lądowa”, w tym liczby 2,9 kg, 30 kg i 5,9 kg.',
        'Na kafelkach zrównoważenia zdjąć „100% certyfikowanych w sposób zrównoważony”, „0% witryn powyżej limitu” i „minimalny wpływ na bentos” jako cechy środowiskowe. Cele z terminem 2023 i 2025 przełączyć na wynik albo ukryć do czasu wstawienia wyniku.',
        'Na stronie Mowi Feed zdjąć „ekologicznymi”, „organicznymi” i „przyjazny dla środowiska”, dopóki w akapicie nie ma numeru certyfikatu przy nazwanej paszy.'
    ) (Rgb 255 241 242) $DANGER

    Add-Callout 'ETAP 2: KOREKTA STRON (do 28 października 2026 r.)' @(
        'Uzgodnić jedną wielkość zbioru za 2023 r., jedną liczbę zatrudnienia i jedną liczbę porcji, z rokiem i z gramaturą. Zostawić jeden zestaw celów emisji, z rokiem bazowym 2019 tam, gdzie strona zobowiązań go podaje.',
        'Na stronie certyfikatów podlinkować politykę albo usunąć „(link)”. Na stronie partnerstw poprawić nagłówek GSSI. Przy ASC napisać, że 50% miejsc jest stanem na koniec 2021 r.',
        'Na kafelku wody przywrócić regiony niedoboru i rok 2018. Przy opakowaniach rozdzielić recykling, recyklat i kompostowanie. Przy astaksantynie, omega-3 i płacach zastosować wzory z przypadków 11 i 13.'
    ) (Rgb 255 247 237) $WARNING

    Add-Callout 'ETAP 3: METODYKA I DOWÓD W TYM SAMYM KANALE (do 20 listopada 2026 r.)' @(
        'Przy liczbie śladu opublikować rok, jednostkę, granice, udział paszy i metodę ISO 14067 oraz dokument, który odbiorca otwiera bez rejestracji.',
        'Opublikować aktualny udział wolumenu z ASC, BAP i GLOBALG.A.P. oraz wynik celów 2023 i 2025: opakowania, sprzęt, woda, składowanie, monitoring dobrostanu.',
        'Dla soi i surowców morskich podać masę, rok i nazwany standard. Energię elektryczną ze źródeł odnawialnych podać jako udział zużycia z umowy na dostawę. Gwarancje pochodzenia wykazuje się osobną liczbą.'
    ) (Rgb 239 246 255) $ACCENT

    Add-Callout 'ETAP 4: SŁOWNIK I STAŁY PRZEGLĄD (od 20 listopada 2026 r.)' @(
        'Zatwierdzić słownik zdań dopuszczonych do publikacji: partia, numer certyfikatu, rok, granice. Hasło nie przechodzi na „planetę”, „styl życia” ani na „całą żywność z oceanu”.',
        'Każdy nowy kafelek z liczbą przechodzi sprawdzenie: czy termin nie minął, czy ta sama liczba na drugiej podstronie jest inna i czy zgodność z prawem nie jest pokazana jako przewaga.',
        'Wersja polska nie zostaje w tyle za korektą wersji globalnej. Ten kanał czyta odbiorca w Polsce, w tym sieć, która może powtórzyć zdanie w swoim raporcie.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '7. Wnioski'
    Add-P 'Witryna mowi.com/pl/ w stanie na 21 października 2026 r. nie nadaje się do dalszej publikacji w obecnych hasłach środowiskowych. Ogólne „przyjazny dla klimatu” i „zrównoważony”, porównanie z żywnością lądową wraz z liczbą 2,9 kg CO2e, superlatyw indeksu FAIRR, kafelek 100% zbiorów oraz limity prawne pokazane jako zrównoważenie wyczerpują punkty 4a, 4b i 10a załącznika I. Korekta polega na zawężeniu każdego zdania do partii, roku, standardu i metody albo na jego usunięciu.' 10.5 $false $false $DARK 0 8

    [System.IO.File]::WriteAllText("$env:TEMP\mowi_build_stage.txt", "before-save")
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $name = [string]$outPath
    $doc.SaveAs2($name)
    [System.IO.File]::WriteAllText("$env:TEMP\mowi_build_stage.txt", "saved")
    $doc.Close($false)
    $doc = $null
    Copy-Item -LiteralPath $outPath -Destination $finalPath -Force
    Write-Output "SAVED $finalPath"
}
finally {
    if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
    if ($word -ne $null) { try { $word.Quit() } catch {} }
}
