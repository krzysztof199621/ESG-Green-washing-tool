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
    $footer.Range.Text = 'ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Mowi Poland  ·  15 października 2026 r.'
    $footer.Range.Font.Name = 'Arial'
    $footer.Range.Font.Size = 8
    $footer.Range.Font.Color = $MUTED

    $word.ActiveWindow.ActivePane.View.SeekView = 0
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)

    Add-P 'ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO' 9 $true $false $ACCENT 0 2
    Add-P 'Audyt Ryzyk Greenwashingu: MOWI POLAND' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja komunikacji opublikowanej na https://mowi.com/pl/ (strona główna, O nas, strategia, struktura w tym Mowi Feed i Mowi Farming, Produkty, hodowla łososia, identyfikowalność, smak i zdrowie, jakość, bezpieczeństwo żywności, Mowi 4.0, kariera oraz dział Zrównoważony rozwój: zobowiązania, certyfikaty, partnerstwa, cele SDG i zasady). Stan zebrania i data audytu: 15 października 2026 r. Ocena wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE (UCPD), ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (UPNPR), ustawy o ochronie konkurencji i konsumentów, ustawy o zwalczaniu nieuczciwej konkurencji (UZNK), rozporządzenia (UE) nr 1169/2011 oraz rozporządzenia (UE) 2018/848.' 10.5 $false $true $MUTED 0 8

    $shot = Join-Path $PSScriptRoot 'zrzut_mowi_pl.png'
    if (Test-Path -LiteralPath $shot) {
        $sel.ParagraphFormat.Alignment = 1
        $shape = $sel.InlineShapes.AddPicture($shot)
        $shape.LockAspectRatio = $true
        $shape.Width = $word.CentimetersToPoints(15.4)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.Alignment = 0
    }
    Add-P 'Zrzut górnej części strony https://mowi.com/pl/ z 15 października 2026 r. Hasło wiodące: „W Mowi zapewniamy naszym klientom 8 milionów posiłków dziennie”.' 8.5 $false $true $MUTED 2 10

    $meta = @(
        @('Badany podmiot', 'Mowi Poland S.A., zakład przetwórstwa w Ustce, część grupy Mowi ASA od 2013 r. (informacja ze strony Kariera). Strona kontaktu podaje adres korespondencyjny grupy: PO Box 4102 Sandviken, 5835 Bergen, Norwegia. Witryna /pl/ jest polskojęzycznym kanałem grupy, nie osobną ofertą lokalną bez claimów grupowych.'),
        @('Przedmiot audytu', 'Treść publiczna witryny mowi.com/pl/ zebrana 15 października 2026 r.: strona główna, działy O nas, Produkty, Zrównoważony rozwój i Kariera oraz podstrony wskazane w podtytule. Oceniono zdania widoczne dla odbiorcy. Nie oceniano raportu rocznego grupy, jeśli nie został zacytowany na tej witrynie.'),
        @('Sektor działalności', 'Hodowla, pasza, przetwórstwo i sprzedaż łososia oraz innych owoców morza (FMCG / żywność), kanał detaliczny, HoReCa i przemysłowy. Odbiorca witryny: konsument i klient biznesowy.'),
        @('Reżim prawny', 'Dyrektywa (UE) 2024/825 (ECGT), stosowanie od 27.09.2026, załącznik I pkt 4a, 4b i 10a. Nie stwierdzono na tej witrynie claimu neutralności opartego na offsecie (pkt 4c) ani własnego znaku zrównoważoności grupy (pkt 2a). Dalej: dyrektywa 2005/29/WE, UPNPR art. 5 i 6, ustawa o ochronie konkurencji i konsumentów art. 24 i 106, UZNK art. 16 i 18, rozporządzenie (UE) nr 1169/2011 art. 7, rozporządzenie (UE) 2018/848 art. 30, norma ISO 14067 jako wzorzec śladu węglowego produktu oraz ISO 14021 jako wzorzec deklaracji własnej. Zielone oświadczenia o soi nie zostały ocenione jako naruszenie EUDR, bo witryna nie składa oświadczenia „wolne od wylesiania”.'),
        @('Data audytu', '15 października 2026 r.'),
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
        'Witryna spełnia przesłanki czarnej listy załącznika I do dyrektywy 2005/29/WE w brzmieniu nadanym dyrektywą (UE) 2024/825: pkt 4a (ogólne twierdzenia „przyjazny dla klimatu”, „zrównoważony łosoś”, „dobra dla planety”, „minimalny wpływ na środowisko”, „najbardziej zrównoważony producent białka”), pkt 4b (claim o całej żywności z oceanu i o 100% zbiorów „certyfikowanych w sposób zrównoważony”, podczas gdy jedyny udział ASC podany na tej samej witrynie to 50% hodowli na koniec 2021 r., a linia MOWI Pure jest opisana osobno) oraz pkt 10a („minimalny wpływ na bentos zgodnie z przepisami krajowymi” i „0% witryn powyżej limitu krajowego” jako wskaźniki zrównoważenia).',
        'Osobno witryna porównuje żywność z oceanu z żywnością lądową pod kątem emisji, wody i gruntów bez jednostki funkcjonalnej, granic systemu i metodyki, podaje sprzeczne cele redukcji gazów cieplarnianych (35% dla zakresów 1, 2 i 3 do 2030 r. wobec 50,6%, 27,5% i 33,33% na stronie zobowiązań) oraz utrzymuje cele z terminami 2023 i 2025 bez wyniku na 15 października 2026 r. Na stronie pasz użyto wyrazu „ekologicznymi” bez numeru certyfikatu i jednostki.',
        'Nie stwierdzono twierdzenia o neutralności klimatycznej, net-zero ani kompensacji offsetem, więc pkt 4c nie jest podstawą werdyktu. Zakazy ECGT stosuje się od 27 września 2026 r. Treść strony internetowej nie korzysta z taryfy dla starych zapasów. UPNPR art. 5 i 6 obowiązuje niezależnie od opóźnienia polskiej transpozycji. Utrzymanie tych zdań naraża wydawcę witryny na postępowanie Prezesa UOKiK (art. 24 i 106 ustawy o ochronie konkurencji i konsumentów, kara do 10% obrotu za praktykę) oraz na przeniesienie wadliwego claimu do komunikacji klientów detalicznych raportujących ESRS.'
    ) (Rgb 255 241 242) $DANGER

    Add-H '1. Analiza Wstępna'
    Add-P 'Mowi Poland S.A. prowadzi w Ustce zakład przetwórstwa łososia i jest częścią Mowi ASA, którą witryna opisuje jako największego producenta łososia atlantyckiego, ze zbiorem 475 000 ton w 2023 r. i udziałem około 20% rynku. Strona jest kanałem wobec konsumenta i klienta HoReCa oraz handlu. Narracja łączy skalę („8 milionów posiłków dziennie”), wizję „Leading the Blue Revolution” i obietnicę żywności „zdrowej dla ludzi i dobrej dla planety”.' 10.5 $false $false $DARK 0 8
    Add-P 'Rdzeń środowiskowy nie jest dopiskiem. Hasło „przyjazny dla klimatu” stoi na stronie Produktów, na stronie hodowli łososia, w cytacie o misji na stronie O nas i w ogłoszeniu kariery. Obok niego witryna stawia porównanie całej „niebieskiej żywności” z żywnością lądową, kafelek „100% wielkości zbiorów certyfikowanych w sposób zrównoważony” oraz stopień najwyższy „najbardziej zrównoważony producent białka”. Certyfikaty ASC, GLOBALG.A.P. i BAP są znakami stron trzecich. Nie są uznaną wysoką efektywnością ekologiczną w rozumieniu ECGT (EU Ecolabel albo urzędowo uznany ISO 14024 typ I) i nie uprawniają do ogólnego przymiotnika „zrównoważony” wobec łososia albo całej działalności.' 10.5 $false $false $DARK 0 8
    Add-P 'Na tej samej witrynie liczby się rozjeżdżają. Strona główna i O nas mówią o 8 milionach posiłków, Produkty o ponad 6 milionach. Kafelek klimatu mówi o 35% redukcji zakresów 1, 2 i 3 do 2030 r., a strona zobowiązań o 50,6%, 27,5% i 33,33% wobec bazy 2019. Coller FAIRR jest raz „cztery lata z rzędu (2019–2022)”, a raz „piąty rok z rzędu” za lata 2019–2024. Udział ASC jest podany na koniec 2021 r. (50% hodowli), tabela certyfikacji nosi rok 2023, a cele opakowaniowe i sprzętowe mają terminy 2023 i 2025, które w dniu audytu już minęły, bez wyniku w tym samym kanale. To jest prawda wybiórcza i pominięcie z art. 5 i 6 UPNPR, nałożone na czarną listę ECGT.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'
    $findings = @(
        @('01','„Przyjazny dla klimatu”, „zrównoważony łosoś”, „dobra dla planety”','ECGT zał. I pkt 4a · UPNPR art. 5','BARDZO WYSOKIE','Ogólne hasło środowiskowe bez uznanej wysokiej efektywności'),
        @('02','Niebieska żywność ma niższą emisję i niższy ślad niż żywność lądowa','ECGT zał. I pkt 4b · UZNK art. 16 · UPNPR art. 5 i 6','BARDZO WYSOKIE','Porównanie kategorii bez jednostki, granic systemu i metodyki'),
        @('03','„Najbardziej zrównoważony producent białka” (Coller FAIRR)','ECGT zał. I pkt 4a · UPNPR art. 5','BARDZO WYSOKIE','Superlativ ESG-rankingu; sprzeczne lata 2019–2022 i 2019–2024'),
        @('04','„100% zbiorów certyfikowanych w sposób zrównoważony” (GSSI)','ECGT zał. I pkt 4a i 4b · UPNPR art. 5 i 6','WYSOKIE','Całość zbioru nazwana zrównoważoną; ASC podane jako 50% hodowli na koniec 2021 r.'),
        @('05','„Minimalny wpływ na bentos zgodnie z przepisami”; „0% witryn powyżej limitu”','ECGT zał. I pkt 10a · UPNPR art. 5','WYSOKIE','Zgodność z limitem krajowym jako wskaźnik zrównoważenia'),
        @('06','35% do 2030 r. wobec 50,6% / 27,5% / 33,33%; energia odnawialna','UPNPR art. 5 i 6 · UCPD art. 6 (efektywność przyszła)','WYSOKIE','Sprzeczne cele GHG i brak planu z zasobami oraz weryfikacją'),
        @('07','Opakowania i sprzęt „do 2025” i „do 2023”; kompostowanie','UPNPR art. 5 i 6 · ECGT zał. I pkt 4b','WYSOKIE','Terminy minęły, wyniku brak; „kompostowanie” bez normy'),
        @('08','Pasza „przyjazna dla środowiska” i „standardy ekologiczne”','ECGT zał. I pkt 4a i 4b · rozp. 2018/848 art. 30','WYSOKIE','Ogólny claim fabryki oraz wyraz zastrzeżony bez certyfikatu na stronie'),
        @('09','„100% soi” i „100% surowców morskich” z dopiskiem „lub równoważny”','UPNPR art. 5 i 6','WYSOKIE','Absoluty bez roku, udziału masy i definicji równoważności'),
        @('10','Certyfikaty GFSI w dziale zrównoważenia; zły nagłówek GSSI','UPNPR art. 5 · transfer atrybutów','ŚREDNIE','Bezpieczeństwo żywności podstawione pod profil środowiskowy'),
        @('11','Astaksantyna „naturalna”; omega-3 a zawał, udar i nagła śmierć','Rozp. 1169/2011 art. 7 · UPNPR art. 5','ŚREDNIE','Właściwość i skutek chorobowy bez dawki, pochodzenia i zakresu'),
        @('12','8 milionów i „ponad 6 milionów” posiłków dziennie','UPNPR art. 5','ŚREDNIE','Dwie skale tej samej działalności na jednej witrynie'),
        @('13','„Brak znaczących różnic” płac; „oficjalna płaca wystarczająca”','UPNPR art. 5 i 6','ŚREDNIE','Wynik społeczny bez roku, metody i luki skorygowanej')
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

    Add-Case 'PRZYPADEK 01: Ogólne hasła „przyjazny dla klimatu” i „zrównoważony łosoś”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'Produkty: „Naszym celem jest dostarczanie najwyższej jakości łososia i inspirowanie do zdrowego i przyjaznego dla klimatu stylu życia”.',
            'Hodowla łososia: „Hodowla łososia to produkcja żywności przyjazna dla klimatu.” Oraz: „Owoce morza oferują niski ślad węglowy, wysoką wydajność zatrzymywania energii i białka oraz niski ślad wodny.”',
            'O nas, nagłówek treści: „Zrównoważony łosoś klasy premium”. Cytat CEO na stronie Zrównoważony rozwój: „jesteśmy bardzo dumni z produkcji żywności, która jest zdrowa dla ludzi i dobra dla lokalnych społeczności i planety”.',
            'O nas: „możemy inwestować w innowacje i ulepszenia, które zapewniają zdrowsze życie łososia, minimalny wpływ na środowisko i doskonałą żywność”. Strategia: „pozwala nam wyznaczać ambitne cele w zakresie zrównoważonego rozwoju i wywierać pozytywny wpływ na naszą planetę”. Kariera: „zdrowych i przyjaznych dla klimatu owoców morza” oraz „ochronie planety dla przyszłych pokoleń”.'
        ) `
        'Punkt 4a zakazuje ogólnego twierdzenia środowiskowego, jeżeli w tym samym kanale nie wykazano uznanej wysokiej efektywności ekologicznej istotnej dla hasła, czyli EU Ecolabel albo urzędowo uznanego systemu ISO 14024 typ I. FAQ Komisji do ECGT wymienia „climate-friendly” jako wzór hasła ogólnego. „Przyjazny dla klimatu”, „zrównoważony łosoś”, „dobra dla planety”, „minimalny wpływ” i „pozytywny wpływ na planetę” nie wskazują aspektu, partii, roku ani granic. Certyfikat ASC, BAP albo GLOBALG.A.P. nie jest tą efektywnością. Dopisek na podstronie certyfikatów nie sanuje hasła na stronie Produktów, hodowli i w cytacie otwierającym dział zrównoważenia. „Pozytywny wpływ” w tych zdaniach nie jest claimem offsetowym, więc nie kwalifikuję go jako pkt 4c. Jest ogólną obietnicą korzyści środowiskowej całej działalności.' `
        'Zdjąć z witryny /pl/ przymiotniki „przyjazny dla klimatu”, „zrównoważony” jako określenie łososia, produktu albo stylu życia, „dobra dla planety”, „minimalny wpływ na środowisko” i „pozytywny wpływ na planetę”, dopóki w tym samym widoku nie ma konkretu: partia albo zakład, wskaźnik, rok, granice i metodyka. Wizji „Leading the Blue Revolution” nie używać jako dowodu efektywności.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Łosoś wędzony MOWI z partii [kod], zakład Ustka, zbiór [rok], farma objęta certyfikatem [nazwa standardu, numer, data ważności]. Ślad tej partii: [X] kg CO2e na kg produktu, ISO 14067, granice od paszy do bramy zakładu, dane za [rok]. Zdanie dotyczy tej partii, nie całej działalności Mowi i nie stylu życia konsumenta.”'

    Add-Case 'PRZYPADEK 02: Porównanie żywności z oceanu z żywnością lądową' `
        'Podstawa prawna: ECGT, załącznik I pkt 4b · UZNK art. 16 i 18 · dyrektywa 2006/114/WE · UPNPR art. 5 i 6 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'Zrównoważony rozwój: „Niebieska żywność ma niższą emisję gazów cieplarnianych niż żywność pochodzenia lądowego (BFA, 2021). Zmiana diety w kierunku zwiększonej konsumpcji owoców morza jest uznawana za część rozwiązania problemu zmian klimatu.” Oraz: „Niebieska żywność charakteryzuje się niższym zużyciem wody słodkiej i gruntów w porównaniu z mięsem lądowym.”',
            'Ta sama strona: „Liczba ludności na świecie ma wzrosnąć z 7,7 miliarda obecnie do 9,7 miliarda w 2050 roku.” Przypis: „ONZ, 2020”. Osobno: „Odsetek nadmiernie eksploatowanych łowisk wzrósł z 33,1 do 34,2%. *SOFIA, 2020.”',
            'Certyfikaty, zdanie otwierające: „Żywność z oceanu jest kluczem do zapewnienia pożywnej żywności o mniejszym wpływie na klimat niż produkcja żywności na lądzie. Jest to potrójna wygrana: dla ludzi, dla planety i dla gospodarki.”',
            'Produkty: „Badanie Blue Foods Assessment wykazało, że niebieska żywność pochodząca z oceanu, taka jak łosoś, ma średnio znacznie większe wartości odżywcze niż żywność pochodząca od zwierząt lądowych, a wiele z nich ma również mniejszy wpływ na środowisko.” Hodowla: „Żywność pozyskiwana z oceanu ma niższą emisję gazów cieplarnianych (GHG) niż żywność pochodzenia lądowego”.'
        ) `
        'Punkt 4b zakazuje twierdzenia o całym produkcie albo całej działalności, gdy dowód dotyczy aspektu albo innej kategorii. Witryna przenosi zdanie o „niebieskiej żywności” i o „wielu” produktach z oceny BFA 2021 na hodowlę łososia Mowi i na „potrójną wygraną dla planety”. Jedno zdanie zachowuje kwalifikator „wiele z nich ma również mniejszy wpływ”, a zdania sąsiednie już mówią, że żywność z oceanu ma niższą emisję niż żywność lądowa. To nie jest porównanie porównywalnych: brak jednostki funkcjonalnej (kilogram, białko, kilokaloria), granic systemu i udziału paszy, która jest głównym nośnikiem emisji łososia hodowlanego. UZNK art. 16 wymaga, aby reklama porównawcza była obiektywna i sprawdzalna oraz dotyczyła cech istotnych i reprezentatywnych. Kategoria „mięso lądowe” obejmuje wołowinę i drób o skrajnie różnym śladzie. „Niski ślad węglowy” i „niski ślad wodny” nie mają liczby ani ISO 14067. „7,7 miliarda obecnie” jest opisem danych ONZ z 2020 r. podanym jako stan bieżący w październiku 2026 r. Ukryty kompromis (pasza, w tym soja i surowce morskie, wpływ lokalny na dno i wszy morskie) nie jest w tym samym bloku co „potrójna wygrana”.' `
        'Usunąć porównanie kategorii „ocean kontra ląd” ze stron Produktów, hodowli, certyfikatów i zrównoważenia. Nie zostawiać cytatu Blue Food Assessment jako tekstu dla klienta: to badanie kategorii, nie ślad łososia Mowi. Na karcie produktu użyć zdania poniżej i wpisać własną liczbę. Liczbę ludności i udział przełowionych stad opatrzyć rokiem źródła i nie pisać „obecnie”, jeśli dana nie jest z roku publikacji strony.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Na karcie produktu, obok nazwy, można napisać: Filet z łososia hodowlanego MOWI, [nazwa produktu i gramatura]. Ślad węglowy tej partii wynosi [liczba] kg CO2e na 1 kg produktu. Dane za [rok], metoda ISO 14067. W tej liczbie jest karma, chów ryb i przetworzenie w zakładzie [nazwa zakładu]. W tym samym akapicie nie dodawać zdania, że produkt ma niższy ślad niż mięso, niższy niż żywność z lądu albo że żywność z oceanu jest lepsza dla klimatu.”'

    Add-Case 'PRZYPADEK 03: „Najbardziej zrównoważony producent białka”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'O nas, nagłówek: „Mowi, największy na świecie dostawca łososia hodowlanego i najbardziej zrównoważony producent białka”.',
            'O nas: „Mowi został uznany za najbardziej zrównoważonego producenta białka przez Coller FAIRR Protein Producer Index piąty rok z rzędu.” Przypis: „MOWI zajęło pierwsze miejsce w rankingu Coller FAIRR Protein Producer Index 2019–2024, kompleksowej ocenie największych producentów białka zwierzęcego pod kątem krytycznych kwestii środowiskowych, społecznych i związanych z zarządzaniem (ESG).”',
            'Mowi Feed: „Firma Mowi została uznana za najbardziej zrównoważonego producenta białka na świecie przez Coller FAIRR Initiative cztery lata z rzędu (2019–2022).”'
        ) `
        '„Najbardziej zrównoważony” jest ogólnym twierdzeniem środowiskowym o całej działalności i o produkcie białkowym. Punkt 4a nie uznaje pierwszego miejsca w prywatnym indeksie ESG za wysoką efektywność ekologiczną. Przypis na stronie O nas sam zawęża indeks do oceny największych producentów białka zwierzęcego pod kątem kwestii ESG. Nagłówek tego zawężenia nie powtarza: mówi o producencie białka w ogóle, nie o miejscu wśród dużych spółek giełdowych ocenianych przez FAIRR. Przypis w tym samym widoku nie zdejmuje hasła z czarnej listy, bo nie podaje aspektu środowiskowego, wagi kryteriów ani tego, czego indeks nie mierzy. Na stronie Mowi Feed ten sam fakt ma inny zakres czasowy: cztery lata, 2019–2022, i dopisek „na świecie”, którego przypis na stronie O nas nie zawiera. Dwie daty na jednej witrynie w dniu 15 października 2026 r. są informacją wprowadzającą w błąd co do aktualności i zakresu wyróżnienia. Wolumen „największy dostawca” (475 000 ton w 2023 r., około 20% udziału) jest zdaniem osobnym i tego przypadku nie dotyczy, o ile przy liczbie zostaje rok.' `
        'Zdjąć superlatyw „najbardziej zrównoważony” z nagłówka O nas i ze strony Mowi Feed. Jeśli indeks zostaje, opublikować tylko miejsce, lata i grupę objętą rankingiem. Nie stawiać tego zdania obok hasła o łososiu, klimacie ani emisji. Uzgodnić jedną wersję lat na całej witrynie /pl/.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W Coller FAIRR Protein Producer Index za lata [RRRR–RRRR] Mowi zajęło miejsce [n] wśród producentów białka zwierzęcego objętych tym indeksem.”'

    Add-Case 'PRZYPADEK 04: „100% zbiorów certyfikowanych w sposób zrównoważony”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5 i 6 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Zrównoważony rozwój, kafelek „Certyfikaty zrównoważonego rozwoju”: „100% wielkości zbiorów certyfikowanych w sposób zrównoważony zgodnie ze standardem uznanym przez GSSI”. To samo zdanie, bez daty, jest na stronie Zobowiązania i działania.',
            'Nasze certyfikaty: „Polityka certyfikacji Mowi (link) stanowi, że 100% naszego rocznego odłowu łososia, tj. wszystkie jednostki hodowlane Mowi – farmy słodkowodne i morskie – muszą być certyfikowane zgodnie ze standardami zrównoważonego rozwoju uznanymi przez Global Seafood Sustainability Initiative (GSSI): ASC, GSA BAP i/lub GLOBALG.A.P.” W treści tej strony zwrot „(link)” nie jest odnośnikiem do dokumentu polityki.',
            'Ta sama strona: „Mowi pozostaje liderem w zakresie certyfikacji ASC, posiadając do końca 2021 r. łącznie 133 miejsca hodowli certyfikowane na całym świecie. Stanowi to 50% wszystkich naszych hodowli i 33% wszystkich hodowli łososia atlantyckiego z certyfikatem ASC na całym świecie”. Oraz: „cały nasz łosoś MOWI Pure jest i będzie certyfikowany ASC”. Tabela certyfikacji odesłana z tej strony nosi w nazwie pliku rok 2023.'
        ) `
        'Kafelek „100%” czyta się jako wynik, bo stoi w rzędzie wskaźników obok liczb. Strona certyfikatów mówi coś węższego: polityka wymaga, aby odłów był certyfikowany według jednego ze standardów uznanych przez GSSI. „Muszą być” nie jest tym samym co „są”, a polityki nie da się otworzyć z miejsca, w którym pada „(link)”. GSSI jest narzędziem porównawczym standardów, nie oznakowaniem produktu i nie typem I ISO 14024. Nazwanie zbioru „certyfikowanym w sposób zrównoważony” zamienia udział w prywatnym standardzie w ogólny wniosek o całej działalności (pkt 4a i 4b). Jedyny udział liczbowy ASC na tej witrynie to 50% hodowli na koniec 2021 r. Linia MOWI Pure jest opisana jako ta, która „jest i będzie” certyfikowana ASC, czyli węziej niż cały zbiór. Wzorzec z zarzutów UOKiK wobec Zary (19.01.2026) jest ten sam: certyfikat i odsetek czytane tak, jakby dotyczyły całości. Dane z końca 2021 r. i PDF z 2023 r. pozostawione bez zdania „stan na” w październiku 2026 r. pomijają informację, której odbiorca potrzebuje (art. 6).' `
        'Na kafelku rozdzielić cel polityki od wyniku. Wynik podać z rokiem, udziałem wolumenu (nie tylko liczbą miejsc) i rozbiciem na ASC, BAP i GLOBALG.A.P. Zdjąć przymiotnik „w sposób zrównoważony”. Podlinkować politykę albo usunąć słowo „(link)”. Zdanie o MOWI Pure zostawić przy tej linii, z numerem certyfikatu i datą, i nie przenosić go na cały łosoś grupy.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] [X]% wolumenu zbioru łososia Mowi pochodziło z jednostek z certyfikatem ASC [numer], [Y]% z BAP i [Z]% z GLOBALG.A.P., według tabeli z dnia [data].”'

    Add-Case 'PRZYPADEK 05: Limit krajowy wszy i wpływ na bentos jako cecha zrównoważenia' `
        'Podstawa prawna: ECGT, załącznik I pkt 10a · UPNPR art. 5 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Zrównoważony rozwój, kafelek „Wszy morskie”: „0% witryn powyżej limitu krajowego (średnia miesięczna)”. Na stronie zobowiązań to samo brzmi: „0% witryn powyżej limitu krajowego (średnia miesięczna)”, bez roku i bez słowa „cel”.',
            'Zobowiązania: „100% terenów uprawnych o minimalnym wpływie na bentos (zgodnie z przepisami krajowymi)”.'
        ) `
        'Punkt 10a zakazuje przedstawiania wymogu prawnego jako cechy wyróżniającej. Zdanie o bentosie samo dopisuje „zgodnie z przepisami krajowymi”: „minimalny wpływ” jest opisem zgodności z limitem, nie wynikiem dobrowolnym powyżej prawa. Kafelek wszy mierzy odsetek miejsc, które nie przekraczają limitu krajowego. Utrzymanie się w limicie jest obowiązkiem prowadzenia hodowli, nie przewagą środowiskową oferty. Brak roku przy „0%” pozwala czytać liczbę jako aktualny wynik albo jako cel. W obu odczytach liczba stoi w dziale zrównoważenia, obok klimatu i certyfikatów, więc funkcjonuje jako claim ekologiczny. Tłumaczenie „terenów uprawnych” dla farm łososia dodatkowo zaciera, czego limit dotyczy.' `
        'Z kafelka „Wszy morskie” zdjąć „0%” bez roku. Jeśli liczba zostaje, podać miesiąc, rok i udział miejsc, które nie przekroczyły krajowego limitu zagęszczenia wszy. Nie dopisywać, że limit jest obowiązkiem prawa. Zdanie o „minimalnym wpływie na bentos” zdjąć z kafelka albo podać metodę, próg i rok.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [miesiąc rok] [X]% miejsc hodowli morskiej Mowi nie przekroczyło krajowego limitu zagęszczenia wszy.”'

    Add-Case 'PRZYPADEK 06: Sprzeczne cele emisji i energia odnawialna bez udziału' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · dyrektywa 2005/29/WE art. 6 w brzmieniu ECGT (przyszła efektywność środowiskowa) | Poziom ryzyka: WYSOKIE' `
        @(
            'Zrównoważony rozwój, kafelek „Zmiana klimatu”: „35% redukcja emisji gazów cieplarnianych z zakresu 1, 2 i 3 do 2030 r. i o 72% do 2050 r.”',
            'Zobowiązania i działania: „Mowi zobowiązuje się do zmniejszenia bezwzględnych emisji gazów cieplarnianych zakresu 1 i 2 o 50,6% do 2030 r. w stosunku do roku bazowego 2019. Mowi zobowiązuje się do zmniejszenia bezwzględnych emisji gazów cieplarnianych zakresu 3 o 27,5% w tych samych ramach czasowych. Mowi zobowiązuje się do zmniejszenia bezwzględnych emisji gazów cieplarnianych z zakresu 3 FLAG o 33,33% w tym samym okresie.”',
            'Strategia: „pozyskiwanie coraz większej ilości energii elektrycznej ze źródeł odnawialnych”.'
        ) `
        'Twierdzenie o przyszłej efektywności środowiskowej wolno oprzeć na jasnym, obiektywnym i publicznym zobowiązaniu wykonanym według szczegółowego planu: cele mierzalne, terminy, zasoby i regularna weryfikacja niezależnego eksperta, z wynikami dostępnymi odbiorcy. Dwie podstrony tego samego działu podają inne liczby na 2030 r. 35% łącznie dla zakresów 1, 2 i 3 nie wynika arytmetycznie ze strony zobowiązań i nie ma tu roku bazowego. 72% do 2050 r. nie ma zakresu ani bazy. Odbiorca nie wie, który cel jest aktualny. Planu zasobów i weryfikatora na tych stronach nie ma. Zdanie o energii odnawialnej nie podaje udziału, roku, kraju zużycia ani tego, czy chodzi o energię fizyczną, umowę PPA, czy gwarancje pochodzenia. Praktyka UOKiK w sprawie Bolt (19.01.2026) odrzuca utożsamienie gwarancji pochodzenia z prądem zużytym. Tu nie ma jeszcze hasła „100% OZE”, jest kierunek bez mianownika, i to wystarczy do art. 6: pominięto dane, bez których zdanie o „coraz większej ilości” nie da się sprawdzić.' `
        'Zostawić jeden zestaw celów, z rokiem bazowym, zakresem i datą przyjęcia, oraz odesłaniem do dokumentu, w którym są zasoby i weryfikator. Drugi zestaw zdjąć. Przy energii podać udział procentowy zużycia, rok, spółki objęte i rodzaj instrumentu (umowa na energię fizyczną albo gwarancja pochodzenia). Nie pisać „odnawialna” przy samym certyfikacie atrybutu bez tego rozróżnienia.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Cel grupowy przyjęty [data]: spadek emisji bezwzględnych zakresu 1 i 2 o 50,6% do 2030 r. wobec 2019 r.; zakresu 3 o 27,5%; zakresu 3 FLAG o 33,33%. Wynik za [rok]: [liczby], metodyka [nazwa], weryfikator [nazwa]. W [roku] energia elektryczna objęta umową na dostawę ze źródeł odnawialnych stanowiła [X]% zużycia [spółka, kraj]. Gwarancje pochodzenia, jeśli są, wykazujemy osobno i nie doliczamy ich do tego udziału.”'

    Add-Case 'PRZYPADEK 07: Cele opakowań i sprzętu po upływie terminu' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · ECGT, załącznik I pkt 4b | Poziom ryzyka: WYSOKIE' `
        @(
            'Zobowiązania: „Do 2025 r. 100% naszych plastikowych opakowań będzie nadawało się do ponownego użycia, recyklingu lub kompostowania. Do 2025 r. co najmniej 25% opakowań z tworzyw sztucznych będzie pochodzić z recyklingu. Do 2023 r. cały plastikowy sprzęt rolniczy (sieci, liny, rury zasilające) będzie ponownie wykorzystywany lub poddawany recyklingowi.”',
            'Ta sama strona: „Osiągnięcie do 2025 r. redukcji zużycia wody o 10% w naszych zakładach przetwórczych zlokalizowanych w regionach o średnim i wysokim ryzyku niedoboru wody. Do 2025 r. zero odpadów składowanych na wysypiskach w naszych zakładach przetwórczych. Do 2023 r. minimum 50% naszego stada z monitorowaniem dobrostanu w czasie rzeczywistym.”',
            'Kafelek na stronie Zrównoważony rozwój skraca wodę do: „10% zmniejszenie zużycia wody w naszych zakładach przetwórczych do 2025 r.”, bez ograniczenia do regionów ryzyka i bez roku bazowego 2018, który jest dopiero na stronie zobowiązań.'
        ) `
        'Dzień audytu jest 15 października 2026 r. Terminy „do 2023” i „do 2025” minęły. Strona wciąż używa czasu przyszłego („będzie”) i nie podaje, czy cel został wykonany, w jakim procencie i dla jakich zakładów. To pominięcie istotnej informacji (art. 6) i utrzymywanie zapowiedzi, która po terminie czyta się jak cecha aktualna (art. 5). Spójnik „lub” przy opakowaniu pozwala uznać cel za spełniony przez samą teoretyczną przydatność do recyklingu albo przez kompostowanie. Witryna nie podaje normy, warunku (kompostownia przemysłowa czy warunki domowe) ani kraju, w którym strumień jest realnie zbierany. Checklista FMCG wymaga normy, odsetka i warunku. Kafelek wody rozszerza cel z regionów o niedoborze na „nasze zakłady przetwórcze” (pkt 4b: węższy zakres pokazany jako szerszy). „Cały plastikowy sprzęt” z terminem 2023 bez wyniku jest claimem o całości wyposażenia.' `
        'Przy każdym celu z lat 2023 i 2025 wstawić wynik albo zdjąć zdanie. Przy opakowaniu rozdzielić trzy ścieżki: ponowne użycie, recykling (z krajem i warunkiem opróżnienia) i kompostowanie (norma, rodzaj instalacji). Na kafelku wody powtórzyć ograniczenie geograficzne i rok bazowy 2018. Nie zostawiać czasu przyszłego po dacie celu.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] [X]% masy opakowań z tworzyw wprowadzonych przez [spółka] nadawało się do recyklingu w [kraj], po opróżnieniu, w strumieniu [nazwa]. Udział recyklatu w tych opakowaniach wyniósł [Y]% masy. Opakowań kompostowalnych według [norma, np. EN 13432, kompostownia przemysłowa] było [Z]% masy. Cel „100% do 2025” [został wykonany w zakresie … / nie został wykonany: luka …].”'

    Add-Case 'PRZYPADEK 08: Pasza „przyjazna dla środowiska” i „standardy ekologiczne”' `
        'Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · rozporządzenie (UE) 2018/848 art. 30 · ustawa o rolnictwie ekologicznym (Dz.U. 2022 poz. 1370) · UPNPR art. 5 | Poziom ryzyka: WYSOKIE (Czarna lista oraz termin zastrzeżony)' `
        @(
            'Mowi Feed: „Sprzedajemy pasze dla ryb w Norwegii, Szkocji, Irlandii i na Wyspach Owczych do produkcji zgodnie ze standardami konwencjonalnymi i ekologicznymi.”',
            'Ta sama strona: „Te specjalnie zbudowane fabryki zostały zaprojektowane specjalnie do produkcji pasz o wysokiej gęstości odżywczej dla ryb hodowlanych (…) i są wyposażone we wszystkie urządzenia, które zwiększają wydajność i opłacalność w sposób przyjazny dla środowiska.”',
            '„Naszym celem jest zatem pomoc naszym klientom w produkcji solidnych, zrównoważonych, wysokiej jakości ryb przy użyciu najszerszego dostępnego spektrum zrównoważonych surowców.” Produkcja „ponad 540 000 ton w całym roku 2020”.'
        ) `
        '„Przyjazny dla środowiska” przy wyposażeniu fabryki jest hasłem ogólnym (pkt 4a). Dotyczy wydajności zakładu, a strona przenosi je na „zrównoważone ryby” i „zrównoważone surowce”. Główny wpływ paszy siedzi w surowcach, nie w sprawności linii. To pkt 4b. Tonaż 540 000 ton jest podany za 2020 r. bez zastrzeżenia, że to nie jest wolumen bieżący. Wyraz „ekologicznymi” przy paszy jest terminem zastrzeżonym. Artykuł 30 rozporządzenia 2018/848 pozwala użyć „eko” i pochodnych w reklamie paszy tylko dla produktu zgodnego z tym rozporządzeniem. Na stronie nie ma numeru certyfikatu, jednostki certyfikującej ani kodu. FAQ Komisji do ECGT traktuje „organic” na żywności i paszy jako lex specialis 2018/848, nie jako wyjątek od obowiązku pokazania certyfikatu. Brak tych danych na stronie nie dowodzi, że certyfikatu nie ma. Dowodzi, że kanał, w którym pada słowo, go nie pokazuje.' `
        'Przy „ekologicznych” wstawić w tym samym akapicie produkt, jednostkę, numer i kod albo usunąć wyraz i zostawić „konwencjonalnymi”. Zdjąć „przyjazny dla środowiska” i „zrównoważonych surowców”. Tonaż opatrzyć rokiem albo podać rok późniejszy, jeśli jest w dokumencie źródłowym grupy, którego ta strona nie cytuje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Mowi Feed sprzedaje pasze konwencjonalne w Norwegii, Szkocji, Irlandii i na Wyspach Owczych. Pasza [nazwa handlowa] jest paszą ekologiczną: jednostka [nazwa], numer [numer], kod [kod]. Pozostałe pasze z tej strony nie są ekologiczne. Wolumen podany na stronie, 540 000 ton, dotyczy 2020 r.”'

    Add-Case 'PRZYPADEK 09: „100% soi” i „100% surowców morskich” z dopiskiem „lub równoważny”' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE' `
        @(
            'Zobowiązania: „100% identyfikowalność surowców paszowych. 100% surowców morskich posiada certyfikat (Marin Trust lub równoważny). 100% soi jest certyfikowane (ProTerra lub odpowiednik). W kierunku niższego śladu węglowego surowców paszowych.”',
            'Kafelek „Zrównoważona pasza” na stronie zbiorczej skraca to do: „100% identyfikowalność surowców paszowych”.'
        ) `
        'Cztery absoluty stoją bez roku, bez udziału masy w paszy i bez wskazania, który zakład paszowy jest objęty. „Lub równoważny” i „lub odpowiednik” nie nazywają drugiego standardu, więc odbiorca nie może sprawdzić, co poza ProTerra i Marin Trust wchodzi do licznika 100%. Identyfikowalność „100%” nie podaje systemu ani głębokości (do regionu, plantacji, partii). Kierunek „niższego śladu węglowego surowców” nie ma bazy, liczby ani ISO 14067. To art. 5 i 6. Witryna nie pisze „wolne od wylesiania”, więc tego przypadku nie opieram na rozporządzeniu EUDR. Prywatny certyfikat soi nie zastępuje oświadczenia należytej staranności, ale takiego oświadczenia strona nie składa i nie trzeba go jej dopisywać na siłę. Trzeba zdjąć absolut, którego nie da się odczytać.' `
        'Przy każdym 100% podać rok, masę objętą, standard z nazwy (bez „lub równoważny”, chyba że lista równoważnych standardów jest w tym samym akapicie) i link do zestawienia partii. Zdanie o śladzie surowców albo uzupełnić liczbą i metodyką, albo usunąć.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] [X]% masy soi kupionej do pasz Mowi Feed dla [kraj/spółka] miało certyfikat ProTerra [numer]. [Y]% masy surowców morskich miało certyfikat Marin Trust [numer]. Innych standardów w tym liczniku nie ma. Identyfikowalność obejmuje [poziom: partia / region], system [nazwa]. Ślad surowców: [liczba, jednostka, ISO 14067, rok] albo tego zdania nie publikujemy.”'

    Add-Case 'PRZYPADEK 10: Bezpieczeństwo żywności w roli dowodu zrównoważenia' `
        'Podstawa prawna: UPNPR art. 5 · transfer atrybutów certyfikatu | Poziom ryzyka: ŚREDNIE' `
        @(
            'Nasze certyfikaty, zdanie ramowe po „potrójnej wygranej” klimatycznej: „Kluczem do osiągnięcia naszego celu jest niezależna certyfikacja naszych działań w zakresie zrównoważonego rozwoju i bezpieczeństwa żywności”. Dalej strona zestawia GSSI, GFSI, FSSC 22000, ASC, GLOBALG.A.P. i BAP.',
            'Kluczowe partnerstwa, nagłówek przy logo GSSI: „Globalna inicjatywa na rzecz bezpieczeństwa żywności”. Akapit pod spodem: „Nadal wspieramy inicjatywę Global Sustainable Seafood Initiative, która odgrywa ważną rolę w zapewnianiu jasności w zakresie certyfikacji owoców morza.”',
            'Jakość: „100% naszych zakładów przetwórczych posiada uznany standard Global Food Safety Initiative (GFSI).”'
        ) `
        'FSSC, BRC, IFS i GFSI potwierdzają system bezpieczeństwa żywności. Nie potwierdzają śladu, wpływu na dno ani paszy. Postawienie ich w jednym bloku z „mniejszym wpływem na klimat” przenosi atrybut bezpieczeństwa na atrybut środowiskowy. Na stronie partnerstw nagłówek nazywa GSSI inicjatywą bezpieczeństwa żywności, a treść i adres ourgssi.org opisują Global Sustainable Seafood Initiative. To pomieszanie dwóch różnych ciał (GFSI i GSSI) w dziale, z którego odbiorca ma wynieść wniosek o zrównoważeniu. Sam fakt certyfikacji zakładów według GFSI, podany na stronie Jakość jako fakt o zakładach, nie jest kwestionowany. Kwestionowane jest użycie go jako klucza do celu klimatycznego oraz błędna nazwa na stronie partnerstw.' `
        'Rozdzielić na witrynie blok „bezpieczeństwo żywności” i blok „standardy hodowli”. Na stronie partnerstw poprawić nagłówek na Global Sustainable Seafood Initiative albo przenieść GFSI do bloku bezpieczeństwa. Przy FSSC zostawić zakład, numer i datę, bez zdania, że certyfikat realizuje cel „dla planety”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Zakład [nazwa, kod] ma certyfikat FSSC 22000 nr [numer], ważny do [data]. Certyfikat dotyczy systemu bezpieczeństwa żywności tego zakładu. Nie jest oceną emisji, paszy ani wpływu hodowli na środowisko. GSSI i GFSI są dwoma różnymi inicjatywami.”'

    Add-Case 'PRZYPADEK 11: „Naturalna” astaksantyna i skutki chorobowe omega-3' `
        'Podstawa prawna: rozporządzenie (UE) nr 1169/2011 art. 7 · UPNPR art. 5 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Hodowla łososia i Produkty, odpowiedź na pytanie o barwienie: „Karma dla łososia zawiera naturalne związki, takie jak astaksantyna, które nadają łososiowi różowy odcień.”',
            'Smak i zdrowie: „Kwasy tłuszczowe omega-3 pomagają utrzymać zdrowe serce poprzez obniżenie ciśnienia krwi, a także zmniejszają ryzyko nagłej śmierci, zawału serca, udaru mózgu i choroby wieńcowej serca.” Oraz: „Łosoś jest ulubieńcem konsumentów, superżywnością”. Produkty: „Łosoś jest (…) superżywnością”.'
        ) `
        'Artykuł 7 rozporządzenia 1169/2011 zakazuje informacji wprowadzającej w błąd co do właściwości i przypisywania środkowi spożywczemu działania, którego nie ma, oraz sugerowania szczególnej właściwości w sposób, którego odbiorca nie może sprawdzić. Odpowiedź na zarzut „czy łosoś jest farbowany” używa słowa „naturalne” bez pochodzenia astaksantyny (glony, drożdże albo synteza) i bez udziału w karmie. Samo słowo zamyka wątpliwość konsumenta co do barwy. Zdanie o omega-3 przypisuje „naszemu” kontekstowi produktowemu obniżenie ryzyka nagłej śmierci, zawału i udaru. Strona cytuje opracowania, w tym materiał FDA o oświadczeniach kwalifikowanych, ale nie podaje dawki EPA i DHA w porcji produktu Mowi ani nie oddziela oświadczenia o funkcji serca od oświadczenia o zmniejszeniu ryzyka choroby. „Superżywność” nie jest nazwą składnika ani dawką. Przypisy bibliograficzne nie zastępują zakresu claimu w tym samym akapicie.' `
        'Przy astaksantynie podać pochodzenie i funkcję technologiczną albo zdjąć „naturalne”. Przy omega-3 zostać przy oświadczeniu, które da się odnieść do porcji, albo usunąć listę chorób. Zdjąć „superżywność”.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Różowa barwa tego łososia pochodzi z astaksantyny w karmie. W partii [kod] astaksantyna pochodzi z [źródło]. Porcja [gramatura] dostarcza [mg] EPA i DHA. Nie publikujemy zdania o zmniejszeniu ryzyka zawału, udaru ani nagłej śmierci, dopóki nie wskażemy oświadczenia, dawki i warunków, które dotyczą tego produktu.”'

    Add-Case 'PRZYPADEK 12: Dwie liczby posiłków dziennie' `
        'Podstawa prawna: UPNPR art. 5 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Strona główna: „W Mowi zapewniamy naszym klientom 8 milionów posiłków dziennie”. O nas: „Dostarczamy naszym klientom 8 milionów posiłków dziennie.”',
            'Produkty: „Każdego dnia na całym świecie spożywanych jest ponad 6 milionów posiłków Mowi.”'
        ) `
        'To nie jest hasło środowiskowe. Jest liczbą skali, na której opiera się reszta narracji, w tym zdania o wyżywieniu populacji. Dwie wartości, 8 milionów i ponad 6 milionów, dotyczą tego samego przedmiotu (posiłki Mowi dziennie) i stoją na jednej witrynie bez roku, definicji posiłku i informacji, która liczba jest aktualna. Artykuł 5 UPNPR obejmuje dane ilościowe, które wprowadzają w błąd co do skali działalności. Audyt nie rozstrzyga, która liczba jest bliższa faktom. Rozstrzyga, że kanał nie może podawać obu jako stanu bieżącego.' `
        'Zostawić jedną liczbę, z rokiem i ze zdaniem, jak liczony jest posiłek (gramatura porcji). Drugą usunąć albo opatrzyć rokiem, jeśli obie pochodzą z różnych lat.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] produkcja Mowi odpowiadała [N] milionom porcji dziennie, przy porcji [g] produktu jadalnego. Liczba dotyczy [grupa / marka] i tego roku.”'

    Add-Case 'PRZYPADEK 13: Luka płacowa i „oficjalna płaca wystarczająca” bez metody' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Bezpieczne i znaczące miejsca pracy: „Żaden pracownik nie otrzymuje wynagrodzenia niższego niż oficjalna krajowa płaca wystarczająca na utrzymanie wskazana dla danej lokalizacji.”',
            'Ta sama strona: „Regularnie mapujemy również wynagrodzenia ze względu na płeć. Nasza ogólna analiza nie wykazuje znaczących różnic w wynagrodzeniach kobiet i mężczyzn.”',
            'Zobowiązania, wśród celów społecznych: „30% kobiet na stanowiskach kierowniczych do 2030 r. Stosunek płci pracowników 50/50 do 2030 r. Brak znaczących różnic między płciami w globalnym badaniu zaangażowania pracowników.”'
        ) `
        'Zdanie o płacy wystarczającej jest mocniejsze niż płaca minimalna i byłoby dodatkiem ponad obowiązek ustawowy, gdyby dało się je sprawdzić. Witryna nie podaje, który organ publikuje „oficjalną krajową płacę wystarczającą” w każdym z krajów działania, ani roku porównania. „Nie wykazuje znaczących różnic” nie podaje luki nieskorygowanej, luki skorygowanej, progu istotności, roku ani tego, czy analiza obejmuje zakłady, w tym Ustkę. Cele 30% i 50/50 są w czasie przyszłym do 2030 r. i same w sobie nie są wynikiem. Sąsiadują jednak ze zdaniem o braku różnic, które brzmi jak wynik już osiągnięty. To art. 6: pominięto metodę i rok. Nie kwalifikuję samych celów różnorodności jako pkt 10a, bo próg 30% kobiet w kadrze kierowniczej i parytet 50/50 nie są w tej treści opisane jako wykonanie obowiązku Kodeksu pracy.' `
        'Przy płacy podać źródło stawki, rok i zakres pracowników (własni / agencja). Przy płci podać dwie luki, rok i populację albo usunąć „nie wykazuje”. Cele 2030 zostawić jako cele, z wynikiem za ostatni rok obok.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W [roku] porównaliśmy wynagrodzenie pracowników [spółki] ze stawką [nazwa źródła] dla [lokalizacja]. Wynik: [opis]. Luka płacowa kobiet i mężczyzn, nieskorygowana, wyniosła [X]%; skorygowana o [czynniki] wyniosła [Y]%. Cel 30% kobiet na stanowiskach kierowniczych dotyczy 2030 r. Wynik za [rok]: [Z]%.”'

    Add-H '4. Porównanie twierdzeń Mowi z rzeczywistym stanem prawno-środowiskowym'
    Add-P 'Witryna porównuje żywność z oceanu, łososia i „niebieską żywność” z żywnością oraz mięsem lądowym. Poniżej zestawiono te porównania ze stanem, który da się odczytać z samej witryny i z wymogu sprawdzenia porównania.' 10.5 $false $false $DARK 0 6

    $comp = @(
        @('Emisja i klimat. „Niebieska żywność ma niższą emisję gazów cieplarnianych niż żywność pochodzenia lądowego (BFA, 2021).” „Hodowla łososia to produkcja żywności przyjazna dla klimatu.” „Owoce morza oferują niski ślad węglowy.”', 'BFA 2021 jest badaniem kategorii, z kwalifikatorem „wiele z nich” użytym tylko w jednym zdaniu strony Produktów. Witryna nie podaje śladu partii Mowi, jednostki funkcjonalnej, granic systemu ani udziału paszy. „Przyjazny dla klimatu” i „niski ślad” są hasłami ogólnymi (ECGT pkt 4a). Przeniesienie średniej kategorii na cały łosoś Mowi jest pkt 4b. Porównanie z „żywnością lądową” nie spełnia UZNK art. 16: kategorie nie są porównywalne, a cecha nie jest sprawdzalna na stronie.'),
        @('Woda i grunty. „Niebieska żywność charakteryzuje się niższym zużyciem wody słodkiej i gruntów w porównaniu z mięsem lądowym.” „Niski ślad wodny.”', 'Brak liczby, roku, jednostki i wskazania, czy „mięso lądowe” to wołowina, wieprzowina czy drób. Ślad wodny łososia hodowlanego zależy od paszy roślinnej, której strona w tym bloku nie wlicza. Cel wody na stronie zobowiązań (minus 10% do 2025 r.) dotyczy zakładów w regionach niedoboru wobec 2018 r., a kafelek mówi o zakładach przetwórczych w ogóle. To inny przedmiot niż ślad wody produktu.'),
        @('Certyfikat a całość zbioru. „100% wielkości zbiorów certyfikowanych w sposób zrównoważony” według standardu uznanego przez GSSI. Osobno: ASC to 50% hodowli na koniec 2021 r., a „cały łosoś MOWI Pure jest i będzie certyfikowany ASC”.', 'GSSI nie jest oznakowaniem produktu. Polityka mówi „muszą być”, kafelek mówi „100%”. Udział ASC podany liczbowo jest historyczny i dotyczy miejsc, niekoniecznie wolumenu, oraz połowy hodowli, nie stu procent. Linia MOWI Pure jest węższa niż „zbiór”. Odbiorca detaliczny nie dostaje rozbicia ASC / BAP / GLOBALG.A.P. za rok bieżący.'),
        @('Skala wyżywienia. „8 milionów posiłków dziennie” oraz „ponad 6 milionów posiłków Mowi”. Populacja „7,7 miliarda obecnie”, przypis ONZ 2020.', 'Dwie skale posiłku bez definicji porcji i bez roku. „Obecnie” przy danej z 2020 r. jest stanem nieaktualnym w dniu audytu. Żadna z tych liczb nie dowodzi niższej emisji produktu.')
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
    Add-P '1. Prezes UOKiK. Praktyka naruszająca zbiorowe interesy konsumentów (art. 24 ustawy o ochronie konkurencji i konsumentów) zagrożona jest karą do 10% obrotu osiągniętego w roku poprzedzającym nałożenie kary, za każdą praktykę (art. 106). Zarzuty z 19.01.2026 r. wobec Tchibo i Zary dotyczą ogólnego „eko” i „zrównoważone” oraz certyfikatu czytanego szerzej niż jego zakres. Zarzut wobec Bolta dotyczy energii odnawialnej opartej na certyfikacie atrybutu. Wzorce z witryny /pl/ są z tej samej grupy. ECGT stosuje się od 27.09.2026 r. do treści online bez taryfy dla starej publikacji. Opóźnienie polskiej ustawy (druk Senatu 825, projektowana data 27.09.2027, akt nieogłoszony) nie wstrzymuje art. 5 i 6 UPNPR.' 10.5 $false $false $DARK 2 8
    Add-P '2. Reklama porównawcza. Zdania o niższej emisji, niższym zużyciu wody i gruntów wobec „mięsa lądowego” i „żywności pochodzenia lądowego” są reklamą porównawczą wobec innych środków spożywczych. UZNK art. 16 i 18 dają konkurentowi roszczenie o zaniechanie, usunięcie skutków i naprawienie szkody, gdy porównanie nie jest obiektywne i sprawdzalne. Nazwany konkurent nie jest potrzebny, gdy kategoria jest wskazana.' 10.5 $false $false $DARK 2 8
    Add-P '3. Łańcuch klientów i ESRS. Odbiorcą witryny jest także sieć detaliczna i HoReCa. Hasło „przyjazny dla klimatu”, „100% zbiorów certyfikowanych w sposób zrównoważony” i „najbardziej zrównoważony producent białka” może wejść do danych o łańcuchu wartości klienta raportującego ESRS E1. Klient, który powtórzy te skróty, przenosi wadę na własny raport. Korekta na mowi.com/pl/ ogranicza to przeniesienie. Nie zastępuje danych pierwotnych, których ta witryna nie publikuje.' 10.5 $false $false $DARK 2 8
    Add-P '4. Pasza i termin „ekologiczny”. Użycie „ekologicznymi” bez jednostki i numeru na stronie Mowi Feed jest ryzykiem z rozporządzenia 2018/848 i nadzoru IJHARS, osobnym od kary UOKiK za praktykę rynkową. Certyfikat, jeśli istnieje, ma być w tym samym akapicie. Jeśli obejmuje tylko część wolumenu, reszta pasz nie może stać w tym zdaniu.' 10.5 $false $false $DARK 2 8

    Add-H '6. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIASTOWY PULL-OFF (15 października 2026 r.)' @(
        'Zdjąć z witryny /pl/ hasła: „przyjazny dla klimatu”, „zrównoważony łosoś”, „dobra dla planety”, „minimalny wpływ na środowisko”, „pozytywny wpływ na planetę”, „najbardziej zrównoważony producent białka”, „niski ślad węglowy”, „niski ślad wodny” oraz porównanie „niższa emisja niż żywność lądowa” w formie absolutnej.',
        'Na kafelkach zrównoważenia zdjąć „100% certyfikowanych w sposób zrównoważony”, „0% witryn powyżej limitu” i „minimalny wpływ na bentos” jako cechy środowiskowe. Cele z terminem 2023 i 2025 przełączyć na wynik albo ukryć do czasu wstawienia wyniku.',
        'Na stronie Mowi Feed zdjąć „ekologicznymi” i „przyjazny dla środowiska”, dopóki w akapicie nie ma certyfikatu.'
    ) (Rgb 255 241 242) $DANGER

    Add-Callout 'ETAP 2: KOREKTA STRON (do 22 października 2026 r.)' @(
        'Uzgodnić jedną liczbę posiłków, jeden zestaw celów GHG (z rokiem bazowym 2019 tam, gdzie strona zobowiązań go podaje) i jeden zakres lat FAIRR.',
        'Na stronie certyfikatów: podlinkować politykę albo usunąć „(link)”, rozdzielić GFSI od GSSI, poprawić nagłówek partnerstw, przy ASC podać, że 50% miejsc jest stanem na koniec 2021 r.',
        'Na kafelku wody przywrócić ograniczenie do regionów niedoboru. Przy opakowaniach rozdzielić recykling, recyklat i kompostowanie.',
        'Przy astaksantynie i omega-3 zastosować wzory z przypadków 11. Przy płacach zdjąć „nie wykazuje znaczących różnic”, dopóki nie ma roku i dwóch luk.'
    ) (Rgb 255 247 237) $WARNING

    Add-Callout 'ETAP 3: METODYKA I DOWÓD W TYM SAMYM KANALE (do 14 listopada 2026 r.)' @(
        'Opublikować przy claimie ilościowym: rok, jednostkę, granice systemu, udział paszy, standard (ISO 14067 dla śladu produktu) i dokument, który odbiorca otwiera bez rejestracji.',
        'Opublikować aktualny udział wolumenu z ASC, BAP i GLOBALG.A.P. oraz wynik celów 2023 i 2025 (opakowania, sprzęt, woda, składowanie, monitoring dobrostanu 50%).',
        'Dla soi i surowców morskich podać masę, rok i standard bez zwrotu „lub równoważny”, chyba że lista standardów jest w akapicie. Energię odnawialną rozdzielić na umowę fizyczną i gwarancje pochodzenia.'
    ) (Rgb 239 246 255) $ACCENT

    Add-Callout 'ETAP 4: SAFE HARBOR I STAŁY PRZEGLĄD (od 14 listopada 2026 r.)' @(
        'Zatwierdzić słownik claimów dopuszczalnych: partia, numer certyfikatu, rok, granice, brak przeniesienia na „planetę”, „styl życia” i „całą żywność z oceanu”.',
        'Każdy nowy kafelek liczbowych celów przechodzi sprawdzenie: czy termin nie minął, czy liczba jest na drugiej podstronie inna, czy zgodność z prawem nie jest pokazana jako przewaga.',
        'Wersja polska nie może zostać w tyle za korektą wersji globalnej. To ten kanał czyta odbiorca w Polsce, w tym sieć, która powtórzy zdanie w swoim raporcie.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '7. Wnioski'
    Add-P 'Witryna mowi.com/pl/ w stanie na 15 października 2026 r. nie nadaje się do dalszej publikacji w obecnych hasłach środowiskowych. Ogólne „przyjazny dla klimatu” i „zrównoważony”, porównanie z żywnością lądową, superlatyw FAIRR, kafelek 100% zbiorów oraz limity prawne pokazane jako zrównoważenie wyczerpują pkt 4a, 4b i 10a załącznika I w brzmieniu ECGT. Korekta polega na zawężeniu każdego zdania do partii, roku, standardu i metodyki albo na jego usunięciu.' 10.5 $false $false $DARK 0 8

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
