#Requires -Version 5.1
# Raport audytu witryny https://impel.pl/ zgodnie z INSTRUKCJA_GENEROWANIA_RAPORTOW.md
$ErrorActionPreference = 'Stop'
$finalPath = Join-Path $PSScriptRoot 'Raport_Audytu_Greenwashingu_Grupa_IMPEL.docx'
$outPath = Join-Path $env:TEMP 'Raport_Audytu_Greenwashingu_Grupa_IMPEL.docx'

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
    $footer.Range.Text = 'ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Grupa Impel  ·  4 listopada 2026 r.'
    $footer.Range.Font.Name = 'Arial'
    $footer.Range.Font.Size = 8
    $footer.Range.Font.Color = $MUTED

    $word.ActiveWindow.ActivePane.View.SeekView = 0
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)

    Add-P 'ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO' 9 $true $false $ACCENT 0 2
    Add-P 'Audyt Ryzyk Greenwashingu: GRUPA IMPEL' 20 $true $false $NAVY 2 2
    Add-P 'Weryfikacja komunikacji opublikowanej na https://impel.pl/ (strona główna, standardy jakości, dział ESG: strona zbiorcza, raport, cele, działania, zrównoważony łańcuch wartości, polityki, EcoVadis, organizacje) oraz raportu „ESG_Impel_raport_2025.pdf” i certyfikatów podlinkowanych z tych stron. Do zakresu weszły też zdania środowiskowe na stronach usługowych grupy, do których prowadzi oferta: czystosc.impel.pl (sprzątanie obiektów), serwis-techniczny.impel.pl (efektywność energetyczna) oraz ipro.impel.pl (magazyny energii i audyty energetyczne). Stan zebrania i data audytu: 4 listopada 2026 r. Ocena wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE (UCPD), ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (UPNPR), ustawy o ochronie konkurencji i konsumentów, ustawy o zwalczaniu nieuczciwej konkurencji (UZNK), Kodeksu pracy oraz ustawy o ochronie sygnalistów.' 10.5 $false $true $MUTED 0 8

    $shot = Join-Path $PSScriptRoot 'zrzut_impel_pl.png'
    if (Test-Path -LiteralPath $shot) {
        $sel.ParagraphFormat.Alignment = 1
        $shape = $sel.InlineShapes.AddPicture($shot)
        $shape.LockAspectRatio = $true
        $shape.Width = $word.CentimetersToPoints(15.4)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.Alignment = 0
    }
    Add-P 'Zrzut górnej części strony https://impel.pl/ z 4 listopada 2026 r. Hasło wiodące: „Napędza nas technologia”. W banerze widać turbiny wiatrowe i drzewo zbudowane z obwodów.' 8.5 $false $true $MUTED 2 10

    $meta = @(
        @('Badany podmiot', 'Impel S.A., ul. A. Słonimskiego 1, 50-304 Wrocław, Contact Center +48 800 190 911. KRS 0000004185, NIP 899-01-06-769, REGON 006318849, kapitał zakładowy 64 325 885 zł, status dużego przedsiębiorcy. Prezes Zarządu: Grzegorz Dzik. Witryna jest kanałem Grupy Impel.'),
        @('Przedmiot audytu', 'Treść publiczna wymieniona w podtytule, zebrana 4 listopada 2026 r. Oceniono zdania widoczne dla odbiorcy oraz pliki, które strona każe pobrać jako dowód: raport ESG za 2025 r., politykę środowiskową, certyfikaty EU Ecolabel i certyfikaty ISO 14001. Nie oceniano całego archiwum aktualności ani każdej podstrony branżowej bez zdania środowiskowego.'),
        @('Sektor działalności', 'Outsourcing usług dla firm i instytucji: utrzymanie czystości, ochrona, serwis techniczny budynków, efektywność energetyczna, pralnie, catering, chemia profesjonalna i procesy administracyjne. Odbiorca witryny: klient biznesowy, a kanał jest publiczny.'),
        @('Reżim prawny', 'Ustawa o przeciwdziałaniu nieuczciwym praktykom rynkowym, art. 5, 6 i 7 pkt 2. Dyrektywa (UE) 2024/825 (ECGT), stosowanie od 27.09.2026, załącznik I pkt 2a, 4a, 4b i 10a oraz dyrektywa 2005/29/WE art. 6 ust. 2 lit. d (przyszła efektywność środowiskowa). Ustawa o ochronie konkurencji i konsumentów art. 24 i 106. UZNK art. 16 ust. 3 i art. 18. Kodeks pracy oraz ustawa o ochronie sygnalistów (Dz.U. 2024 poz. 928). Na badanych stronach nie ma zdania o neutralności klimatycznej opartego na kompensacji emisji (pkt 4c). Raport za 2025 r. podaje, że grupa nie kupiła jednostek offsetowych ani kredytów węglowych.'),
        @('Data audytu', '4 listopada 2026 r.'),
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
        'Strona główna stawia „Usługi przyjazne środowisku” jako cechę outsourcingu, a baner pokazuje turbiny wiatrowe. Zezwolenie EU Ecolabel nr PL/052/002 dotyczy węższej usługi: utrzymania czystości w obiektach handlowych i biurowych, bez dezynfekcji i aromatyzacji, realizowanej przez Dział Kontraktów EKO w Impel Facility Services. Na stronie łańcucha wartości grupa nadaje dostawcom własny „Certyfikat Kwalifikowanego Dostawcy w zrównoważonym łańcuchu wartości”. Szkolenia BHP, ocena ryzyka i kanał dla sygnalistów stoją w dziale ESG jako działania grupy.',
        'Grafika „Ślad węglowy Grupy Impel – dane za 2025 r.” podaje zakres 1 i 2. Raport, który strona każe pobrać, mówi, że zakres 3 dojdzie w kolejnych latach, że część energii oszacowano z faktur i cen z internetu, oraz że grupa obecnie nie posiada celów klimatycznych. Strona EcoVadis podaje mimo to cel 20% do 2030 r. dla dwóch spółek. Gwarancje pochodzenia są opisane jako pokrycie 100% emisji z energii elektrycznej Impel S.A., przy koszyku grupy: 90,5% energii ze źródeł kopalnych i 0 MWh własnej produkcji odnawialnej.',
        'Dwa certyfikaty ISO 14001 nadal leżą na stronie standardów, choć ich ważność skończyła się 14 lipca 2026 r. (Climbex) i 31 lipca 2026 r. (Impel Rental i Praxima Krakpol). Zakazy ECGT stosuje się od 27 września 2026 r. Treść strony nie korzysta z taryfy dla starych zapasów. UPNPR art. 5 i 6 obowiązuje niezależnie od daty polskiej ustawy. Utrzymanie tych zdań naraża wydawcę na postępowanie Prezesa UOKiK, z karą do 10% obrotu za praktykę, oraz na zarzut reklamy porównawczej przy „mniejszym stopniu ingerencji” i przy „drugiej firmie w Polsce”.'
    ) (Rgb 255 241 242) $DANGER

    Add-H '1. Analiza Wstępna'
    Add-P 'Impel S.A. prowadzi grupę outsourcingu usług dla firm i instytucji. Strona główna opisuje 35 lat usług w sprzątaniu, ochronie, serwisie technicznym, procesach administracyjnych i usługach dla przemysłu. Narracja sprzedażowa łączy skalę i technologię z cechą środowiskową: wśród wyróżników outsourcingu jest kafelek „Usługi przyjazne środowisku”, a baner otwierający stronę pokazuje turbiny wiatrowe i drzewo z obwodów. Dział ESG, rozbudowany w 2026 r., mówi o zrównoważonym rozwoju jako warunku wzrostu, o dobrowolnym raporcie i o przekładaniu deklaracji na działania.' 10.5 $false $false $DARK 0 8
    Add-P 'W badanym materiale występują twierdzenia sprawdzalne. Mogą one pozostać w komunikacji pod warunkiem ograniczenia ich treści do zakresu objętego przedstawionym dowodem. Raport za 2025 r. podaje, że grupa nie kupiła jednostek offsetowych ani kredytów węglowych. W pralniach w Trzebini i Chojnowie opisano spadek zużycia wody na tunelu Passat z 10 l/kg do 8,5 l/kg i obniżenie temperatury prania do 60–65°C. Zezwolenie EU Ecolabel nr PL/052/002, ważne od 16 lipca 2025 r. do 31 grudnia 2027 r., jest prawdziwym znakiem typu I dla nazwanej usługi Impel Facility Services. Hasła na stronie głównej, na stronie sprzątania i w ofercie energii są szersze niż ten znak, a liczby klimatyczne na stronie EcoVadis nie zgadzają się ze zdaniem raportu, że grupa nie ma celów klimatycznych.' 10.5 $false $false $DARK 0 8
    Add-P 'Na stronie działań grupa pokazuje ślad jako zakres 1 i 2 za 2025 r.: 16 730,0 Mg CO2e oraz 10 798,1 Mg CO2e albo 9 798,0 Mg CO2e w zakresie 2, zależnie od metody. Ten sam raport mówi, że energia grupy w 90,5% pochodzi ze źródeł kopalnych, że własna produkcja odnawialna wynosi 0 MWh, i że zakres 3 ma być doliczony później. Certyfikaty ISO 14001, na które strona standardów powołuje troskę o środowisko całej grupy, są po terminie ważności. Dostawcy dostają własny certyfikat „zrównoważonego łańcucha”, a obowiązki BHP i kanał sygnalisty są wpisane w listę „Co robimy” w ESG.' 10.5 $false $false $DARK 0 8

    Add-H '2. Analiza szczegółowa - tabela zbiorcza naruszeń'
    $findings = @(
        @('01','„Usługi przyjazne środowisku” i baner z turbinami','UPNPR art. 5 · ECGT zał. I pkt 4a i 4b','BARDZO WYSOKIE','Ogólne hasło o outsourcingu; miks energii grupy jest kopalny w 90,5%'),
        @('02','Środki „przyjazne środowisku” i „np. EU Ecolabel”','UPNPR art. 5 i 6 · ECGT zał. I pkt 4a i 4b','WYSOKIE','Zezwolenie węższe niż oferta sprzątania'),
        @('03','Certyfikat kwalifikowanego dostawcy','UPNPR art. 7 pkt 2 · ECGT zał. I pkt 2a i 4a','WYSOKIE','Własny znak zrównoważonego łańcucha'),
        @('04','ISO 14001 jako praktyka całej grupy','UPNPR art. 5 i 7 pkt 2 · ECGT zał. I pkt 4a i 4b','WYSOKIE','Certyfikaty po terminie i tylko u części spółek'),
        @('05','BHP, równe traktowanie i sygnaliści jako ESG','UPNPR art. 5 · ECGT zał. I pkt 10a','WYSOKIE','Obowiązek ustawowy jako działanie grupy'),
        @('06','Ślad węglowy grupy za 2025 r.','UPNPR art. 5 i 6 · ECGT zał. I pkt 4b','WYSOKIE','Zakres 1 i 2 przy haśle o wpływie organizacji; część energii z szacunku'),
        @('07','Cel 20% do 2030 r. dla dwóch spółek','UPNPR art. 5 i 6 · dyrektywa 2005/29/WE art. 6 ust. 2 lit. d','WYSOKIE','Raport mówi, że grupa nie ma celów klimatycznych'),
        @('08','Gwarancje pochodzenia: 100% i około 20%','UPNPR art. 5 i 6','WYSOKIE','Instrument rynkowy opisany jak redukcja emisji'),
        @('09','„Ekologiczność” magazynu energii i spadek CO2','UPNPR art. 5 i 6 · ECGT zał. I pkt 4a','WYSOKIE','Hasło ogólne i skutek emisji bez instalacji i roku'),
        @('10','Audyt, który „pozwoli spełnić wymogi ESG”','UPNPR art. 5 i 6','WYSOKIE','Audyt energii pokazany jak spełnienie ESG i jak pewny spadek CO2'),
        @('11','Produkty Kappa „zgodnie z ISO, Ecolabel”','UPNPR art. 5 i 6','ŚREDNIE','Nazwa znaku przy chemii bez numeru zezwolenia'),
        @('12','Polityka „zgodna z Porozumieniem Paryskim” i medal EcoVadis','UPNPR art. 5 i 6','WYSOKIE','Zgodność z porozumieniem bez celu; medal bez roku i wyniku')
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
        if ($claimTitle -like '*BRAK WZORU*') { $bg = (Rgb 255 241 242); $tc = $DANGER }
        Add-Callout $claimTitle @($claim) $bg $tc
    }

    $a01 = @'
Na stronie głównej czytelnik odnosi wrażenie, że outsourcing Impel jest przyjazny środowisku, a w sprzątaniu użyje środków bezpiecznych dla środowiska. Nad tym stoi baner z turbinami wiatrowymi i drzewem z obwodów. W kafelku brakuje spółki, obiektu, udziału środków i numeru zezwolenia. Ogólne hasło środowiskowe wolno zostawić tylko wtedy, gdy w tym samym miejscu jest EU Ecolabel albo urzędowy znak ISO 14024 typu I, i tylko w zakresie, którego znak dotyczy. Zezwolenie PL/052/002 leży na innej stronie, dotyczy Impel Facility Services i usługi utrzymania czystości w obiektach handlowych oraz biurowych, z wyłączeniem dezynfekcji i aromatyzacji, realizowanej przez Dział Kontraktów EKO. Hasło, które samo tworzy fałszywe wrażenie, ocenia się jak działanie z art. 5 UPNPR: drugie zdanie kafelka i plik na stronie standardów nie poprawiają nagłówka (C-611/14). W raporcie za 2025 r. udział źródeł kopalnych w zużyciu energii grupy wynosi 90,5%, a własna produkcja energii odnawialnej 0 MWh.
'@
    Add-Case 'PRZYPADEK 01: „Usługi przyjazne środowisku” na stronie głównej' `
        'Podstawa prawna: UPNPR art. 5 · ECGT, załącznik I pkt 4a i 4b · C-611/14 | Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)' `
        @(
            'Strona główna, wyróżniki outsourcingu: „Usługi przyjazne środowisku. W usługach sprzątania wykorzystamy środki bezpieczne dla czyszczonych powierzchni oraz środowiska”.',
            'Baner tej samej strony: „Napędza nas technologia”. W kadrze są turbiny wiatrowe i drzewo zbudowane z obwodów.',
            'Raport ESG 2025, tabela zużycia energii: „Udział źródeł kopalnych w całkowitym zużyciu energii 90,5%”. „Zużycie energii odnawialnej produkowanej samodzielnie bez użycia paliwa 0,0 MWh”. „Udział źródeł odnawialnych w całkowitym zużyciu energii 9,5%”.'
        ) `
        $a01 `
        'Zdjąć nagłówek „Usługi przyjazne środowisku” i zdanie o środkach bezpiecznych dla środowiska. Baner z turbinami zdjąć albo przestawić tak, aby nie otwierał oferty grupy. Zdanie o zezwoleniu EU Ecolabel, jeśli zostaje, stoi przy usłudze sprzątania Impel Facility Services, z numerem, datami i wyłączeniem dezynfekcji oraz aromatyzacji. Wyłączenie zostaje w zdaniu, bo strona sprzątania obiektów oferuje „dezynfekcję powierzchni”, a menu witryny czystosc.impel.pl zawiera „Usługi DDD – dezynsekcja, deratyzacja, dezynfekcja”. Zezwolenie tych czynności nie obejmuje. Nie stawiać tego zdania pod nagłówkiem o całym outsourcingu.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Utrzymanie czystości w obiektach handlowych i biurowych, z wyłączeniem dezynfekcji i aromatyzacji, świadczone przez Dział Kontraktów EKO w Impel Facility Services Sp. z o.o., jest objęte zezwoleniem EU Ecolabel nr PL/052/002, ważnym od 16 lipca 2025 r. do 31 grudnia 2027 r.”'

    $a02 = @'
Na stronie sprzątania obiektów czytelnik odnosi wrażenie, że Impel sprząta nowoczesnym sprzętem i środkami przyjaznymi środowisku, a wyróżnikiem są środki, które w mniejszym stopniu ingerują w środowisko i mają certyfikat, na przykład EU Ecolabel. Niżej jest zdanie, że usługę realizuje Biuro Kontraktów EKO przez wypełnianie kryteriów certyfikacji. Brakuje udziału środków, nazwy środka i numeru przy pierwszym haśle. Słowo „na przykład” zostawia inne certyfikaty bez nazwy. Ogólne hasło „przyjazne środowisku” wymaga znaku typu I w tym samym miejscu, a hasło o całej usłudze sprzątania wolno oprzeć tylko na zakresie zezwolenia. Zezwolenie PL/052/002 obejmuje obiekty handlowe i biurowe, z wyłączeniem dezynfekcji i aromatyzacji, u Impel Facility Services. Ta sama strona w opisie usługi wymienia „dezynfekcję powierzchni”, a menu witryny czystosc.impel.pl zawiera „Usługi DDD – dezynsekcja, deratyzacja, dezynfekcja”. Zdanie o zezwoleniu bez tego wyłączenia czytelnik połączy z dezynfekcją, którą strona już oferuje. Aromatyzacji strona nie nazywa osobno. Zostaje w zdaniu, bo zezwolenie wymienia ją w tej samej granicy usługi co dezynfekcję. Strona oferuje też hale magazynowe, obiekty przemysłowe, hotele i użyteczność publiczną. Raport nazywa ten certyfikat wyróżnikiem rynkowym grupy i podaje numer PL/052/002/001, czyli świadectwo dla centrali mBanku, a w tym samym akapicie wymienia także sklepy Smyk, BOŚ Bank, Galerię Kaskada i Galerię Łódzką. Nagłówek „środki przyjazne środowisku” czyta się przed zdaniem o biurze EKO. To drugie zdanie nie poprawia nagłówka (C-611/14).
'@
    Add-Case 'PRZYPADEK 02: Sprzątanie „środkami przyjaznymi środowisku”' `
        'Podstawa prawna: UPNPR art. 5 i 6 · ECGT, załącznik I pkt 4a i 4b · C-611/14 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'czystosc.impel.pl, sprzątanie obiektów: „Zadbamy o to, by w budynkach Twojej firmy panował idealny porządek, stosując w tym celu nowoczesny sprzęt oraz środki przyjazne środowisku”.',
            'Ta sama strona, „Co nas wyróżnia?”: „Oferujemy usługi sprzątania z wykorzystaniem środków czystości, które w mniejszym stopniu ingerują w środowisko naturalne i posiadają odpowiedni certyfikat, np. EU Ecolabel”.',
            'Pod ofertą: „Usługa realizowana przez Biuro Kontraktów EKO poprzez wypełnianie kryteriów określonych w procesie certyfikacji EU Ecolabel”.',
            'Ta sama strona, opis usługi: „Oferujemy różnego rodzaju profesjonalne usługi sprzątające, np.: codzienne sprzątanie biur, czyszczenie okien, mycie podłóg, usuwanie kurzu, dezynfekcję powierzchni i wiele innych.” Menu witryny czystosc.impel.pl: „Usługi DDD – dezynsekcja, deratyzacja, dezynfekcja”.',
            'Certyfikat główny PL/052/002/005-1 z 16 lipca 2025 r.: zezwolenie dla usługi „Utrzymanie czystości w obiektach handlowych oraz biurowych wyłączając dezynfekcję oraz aromatyzację realizowane przez Dział Kontraktów EKO”, Impel Facility Services Sp. z o.o., ważne do 31 grudnia 2027 r. Raport ESG 2025: „Polskie Centrum Badań i Certyfikacji S.A. udzieliło spółce Impel Facility Services zezwolenia na używanie oznakowania EU Ecolabel nr PL/052/002/001”. Oraz: „Dla Grupy Impel certyfikat EU Ecolabel stanowi wyróżnik rynkowy w obszarze odpowiedzialnych środowiskowo usług utrzymania czystości”.'
        ) `
        $a02 `
        'Z pierwszego akapitu i z listy wyróżników zdjąć „środki przyjazne środowisku” oraz „w mniejszym stopniu ingerują”. Zostawić jedno zdanie o zezwoleniu, z numerem PL/052/002, datami, wyłączeniem dezynfekcji i aromatyzacji oraz nazwą Dział Kontraktów EKO, tak jak na certyfikacie. Wyłączenia nie skracać, dopóki na tej stronie jest „dezynfekcja powierzchni” albo w menu jest „Usługi DDD”. Wyłączenie schodzi ze zdania dopiero wtedy, gdy dezynfekcja i aromatyzacja znikną z witryny czystosc.impel.pl. Numer PL/052/002/001 zostawić tylko przy obiekcie mBanku. Zdanie o wyróżniku rynkowym grupy zdjąć z raportu albo zawęzić do usługi i obiektów objętych zezwoleniem. Porównanie „w mniejszym stopniu” nie wraca bez nazwy środka odniesienia, roku i tej samej metody. Test tego porównania jest w rozdziale 4.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Utrzymanie czystości w obiektach handlowych i biurowych, z wyłączeniem dezynfekcji i aromatyzacji, świadczone przez Dział Kontraktów EKO w Impel Facility Services Sp. z o.o., jest objęte zezwoleniem EU Ecolabel nr PL/052/002, ważnym od 16 lipca 2025 r. do 31 grudnia 2027 r. Dla mBanku, ul. Prosta 18 w Warszawie, numer świadectwa to PL/052/002/001-1.”'

    $a03 = @'
Na stronie łańcucha wartości czytelnik odnosi wrażenie, że łańcuch Impel jest zrównoważony, a dostawca po rocznej kwalifikacji dostaje certyfikat zrównoważonego łańcucha wartości grupy. Na stronie działań jest zdjęcie wręczenia „specjalnego certyfikatu Grupy Impel”. Brakuje progu punktowego, listy kryteriów, udziału dostawców, którzy certyfikat dostali, i jednostki spoza grupy. Znak zrównoważonego charakteru wolno pokazywać, gdy stoi za nim certyfikacja niezależnej strony trzeciej albo organ publiczny. Ten dokument wystawia sama Grupa Impel, po własnym kwestionariuszu i audycie zakupowym. Słowo „certyfikat” i przymiotnik „zrównoważony” w nazwie robią z oceny zakupowej znak środowiskowy. Nagłówek strony mówi o całym łańcuchu, a opis dotyczy kwalifikacji dostawców TIER 1 i TIER 2.
'@
    Add-Case 'PRZYPADEK 03: Własny certyfikat zrównoważonego łańcucha' `
        'Podstawa prawna: UPNPR art. 7 pkt 2 i art. 5 · ECGT, załącznik I pkt 2a i 4a | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Strona „Zrównoważony łańcuch wartości”: „Dostawcy, którzy pozytywnie przejdą proces weryfikacji, otrzymują Certyfikat Kwalifikowanego Dostawcy w zrównoważonym łańcuchu wartości Grupy Impel”.',
            'Ta sama strona: „Kwalifikacja ESG obowiązuje przez rok i po tym czasie podlega ponownej ocenie”. „Proces obejmuje dostawców z poziomu TIER 1 i TIER 2”.',
            'Strona działań, ład korporacyjny: „Kwalifikowani dostawcy otrzymują specjalny certyfikat Grupy Impel”.'
        ) `
        $a03 `
        'Wycofać wzór graficzny i nazwę „Certyfikat Kwalifikowanego Dostawcy w zrównoważonym łańcuchu wartości”. Na stronie i na zdjęciu ze strony działań nie używać słowa „certyfikat” ani przymiotnika „zrównoważony” wobec statusu dostawcy. Tytuł strony zmienić tak, aby nazywał kwalifikację zakupową, a nie cechę łańcucha. Próg, rok oceny i udział dostawców TIER 1 i TIER 2 dopisać przy liczbie, jeśli liczba zostaje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Dostawca TIER 1 albo TIER 2 przeszedł kwalifikację zakupową Grupy Impel na 12 miesięcy, według Regulaminu Kwalifikacji Dostawców, stan na [miesiąc rok].”'

    $a04 = @'
Na stronie standardów czytelnik odnosi wrażenie, że grupa troszczy się o środowisko, a systemy zarządzania, w tym środowiskowy, są codzienną praktyką spółek grupy i potwierdzają standard usług. Pliki ISO 14001 na tej stronie są dwa: Climbex, ważny do 14 lipca 2026 r., oraz Impel Rental i Praxima Krakpol, ważne do 31 lipca 2026 r. Dzień audytu jest 4 listopada 2026 r. Oba terminy minęły, a pliki nadal są pod przyciskiem pobrania. Zakres Climbex to montaż i modernizacja kotłów, rurociągów i instalacji. Zakres Impel Rental i Praximy to pranie i wynajem odzieży. Impel Facility Services, spółka od sprzątania, na tej liście ISO 14001 nie występuje. System ISO 14001 porządkuje zarządzanie. Nie jest EU Ecolabel ani urzędowym znakiem ISO 14024 typu I i nie uprawnia do zdania o trosce środowiskowej całej grupy. Znak zaufania po utracie ważności nie potwierdza aktualnego standardu.
'@
    Add-Case 'PRZYPADEK 04: ISO 14001 jako środowiskowy standard grupy' `
        'Podstawa prawna: UPNPR art. 5 i art. 7 pkt 2 · ECGT, załącznik I pkt 4a i 4b | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Standardy jakości: „Troszczymy się także o środowisko naturalne, a przede wszystkim dbamy o ochronę zdrowia i życia pracowników, poprzez zapewnienie im bezpiecznych i higienicznych warunków pracy. Z tego powodu wdrożyliśmy systemy zarządzania jakością, ochroną środowiska czy bezpieczeństwem i higieną pracy, a gwarantowany standard usług potwierdzamy odpowiednimi certyfikatami. Wszystkie wdrożone systemy są zgodne z międzynarodowymi standardami zarządzania i stanowią codzienną praktykę w działalności spółek należących do Grupy Impel”.',
            'Na tej stronie plik ISO 14001 Climbex Sp. z o.o., nr 63375/B/0001/UK/Po: data ważności 14 lipca 2026 r. Plik ISO 14001 Impel Rental Sp. z o.o. i Praxima Krakpol Sp. z o.o., nr 000814092/7: ważny do 31 lipca 2026 r.'
        ) `
        $a04 `
        'Zdjąć ze strony oba pliki ISO 14001 albo zastąpić je certyfikatem, którego data ważności jest późniejsza niż dzień publikacji. Zdanie o codziennej praktyce wszystkich spółek grupy zawęzić do spółek i zakresów z aktualnego certyfikatu. Zdanie „troszczymy się o środowisko naturalne” zdjąć. Bezpieczne i higieniczne warunki pracy są osobnym przypadkiem 05. Nowego pliku nie publikować, dopóki na stronie nie ma numeru, jednostki, zakresu i dat ważności.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„[Spółka], zakres: [czynność z certyfikatu]. System zarządzania środowiskowego ISO 14001:2015, certyfikat nr [numer], [jednostka], ważny od [data] do [data].”'

    $a05 = @'
W dziale ESG czytelnik odnosi wrażenie, że grupa dba o bezpieczeństwo pracy przez szkolenia BHP, oceny ryzyka i instrukcje, że ogranicza nierówności przez przeciwdziałanie dyskryminacji, oraz że zapewnia platformę dla sygnalistów. Na stronie standardów te same bezpieczne i higieniczne warunki pracy są powodem, dla którego usługi mają gwarantowany standard. Szkolenie BHP na koszt pracodawcy, ocena ryzyka zawodowego oraz bezpieczne i higieniczne warunki pracy są obowiązkiem pracodawcy. Przeciwdziałanie dyskryminacji jest obowiązkiem pracodawcy. Wewnętrzna procedura zgłoszeń dla pracodawcy o statusie dużego przedsiębiorcy jest obowiązkiem ustawy o ochronie sygnalistów. Zakazane jest pokazanie tego obowiązku jako cechy oferty outsourcingu albo jako wkładu w cel „Mniej nierówności”. Odbiorca ma pewność, że czynność się odbywa, gdy zdanie nazywa spółki i odsyła do polityki albo do adresu procedury, które można otworzyć. Liczba z zamkniętego roku wchodzi do zdania, gdy jest w ewidencji. Działania wellbeingowe i wsparcie psychologiczne mogą zostać, gdy mają nazwę, rok i skalę, i gdy nie stoją w jednym punkcie ze szkoleniem BHP.
'@
    Add-Case 'PRZYPADEK 05: Obowiązki BHP, równego traktowania i sygnalistów jako działania ESG' `
        'Podstawa prawna: UPNPR art. 5 · ECGT, załącznik I pkt 10a · Kodeks pracy art. 94 pkt 2b, art. 207 § 2, art. 226 i art. 237³ · ustawa o ochronie sygnalistów (Dz.U. 2024 poz. 928) art. 23 | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Standardy jakości: „dbamy o ochronę zdrowia i życia pracowników, poprzez zapewnienie im bezpiecznych i higienicznych warunków pracy”.',
            'Strona działań, społeczeństwo: „Dbamy o bezpieczeństwo pracy: szkolenia BHP, oceny ryzyka, instrukcje bezpiecznej pracy, kontrole obiektów i działania prewencyjne”.',
            'Nasze cele ESG, „Mniej nierówności”: „Ograniczanie nierówności przekładamy na konkretne działania: polityki, standardy HR, przeciwdziałanie dyskryminacji i wzmacnianie kultury organizacyjnej otwartej na różne potrzeby”.',
            'Strona działań, ład korporacyjny: „Zapewniamy kanały zgłaszania naruszeń, w tym platformę dla sygnalistów oraz procedury ochrony osób zgłaszających”.'
        ) `
        $a05 `
        'Z list „Co robimy” i z celu „Mniej nierówności” zdjąć szkolenia BHP, oceny ryzyka, bezpieczne i higieniczne warunki pracy, przeciwdziałanie dyskryminacji oraz platformę sygnalistów jako osiągnięcia grupy. W ich miejsce, przy Polityce BHP, Polityce różnorodności i przy procedurze zgłoszeń, wstawić zdania ze wzoru. Adres platformy wpisać ten, który jest widoczny przy procedurze. Zdanie o liczbie przeszkolonych w 2025 r. opublikować tylko z liczbą z ewidencji. Przy wsparciu psychologicznym i profilaktyce dopisać nazwę programu i rok.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W spółkach Grupy Impel pracodawca prowadzi szkolenia BHP oraz oceny ryzyka zawodowego. Zasady są w Polityce BHP opublikowanej na stronie impel.pl/esg/polityki/. W spółkach Grupy Impel obowiązuje zakaz dyskryminacji w zatrudnieniu. Zasady równego traktowania są w Polityce różnorodności opublikowanej na stronie impel.pl/esg/polityki/. W Grupie Impel działa wewnętrzna procedura zgłoszeń naruszeń prawa. Zgłoszenie przyjmuje platforma pod adresem [adres widoczny przy procedurze]. W 2025 r. szkolenie BHP odbyło [liczba] pracowników spółek Grupy Impel.”'

    $a06 = @'
Na stronie działań czytelnik odnosi wrażenie, że grupa mierzy ślad węglowy i przez to zarządza wpływem klimatycznym organizacji, a grafika za 2025 r. podaje 16 730,0 Mg CO2e w zakresie 1 oraz dwie wielkości zakresu 2. Sześciokąty nazywają zakres 1 i zakres 2. Nagłówek mówi już o śladzie grupy, a zdanie nad grafiką o wpływie klimatycznym organizacji. W tym samym bloku są transport, środki czyszczące i materiały. Raport za 2025 r. podaje, że wyniki zakresu 1 i 2 publikowane są po raz pierwszy, że zakres 3 grupa planuje dodać w kolejnych latach, oraz że zużycie energii elektrycznej i ciepła oszacowano, dzieląc wartości faktur przez średnie ceny z internetu za 2025 r. Spalanie paliw w transporcie raport opisuje jako dane rzeczywiste. Przy 16 730,0 i przy sumie 27 528 Mg CO2e na grafice brakuje listy spółek i informacji, że część energii jest szacunkiem. Suma na grafice to 27 528 i 26 528 Mg CO2e, a tabela raportu podaje 27 528,1 i 26 528,0. Hasło o wpływie organizacji, oparte na zakresie 1 i 2, jest szersze niż liczba. Na tej samej stronie jest zdanie o 353 maszynach przywróconych do użytku w 2025 r., bez liczby maszyn w parku.
'@
    Add-Case 'PRZYPADEK 06: Ślad węglowy grupy bez granic, które liczba sama podaje' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · ECGT, załącznik I pkt 4b | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'Strona działań: „Naszym celem jest lepsze mierzenie wpływu środowiskowego, odpowiedzialne wykorzystywanie zasobów i stopniowe ograniczanie emisji oraz odpadów”. „Mierzymy ślad węglowy, aby lepiej zarządzać wpływem klimatycznym organizacji”.',
            'Grafika „Ślad węglowy Grupy Impel – dane za 2025 r.”: „16 730,0 Mg CO2e emisje brutto GHG w zakresie 1”; „10 798,1 Mg CO2e emisje brutto GHG w zakresie 2 według metody location-based”; „9 798,0 Mg CO2e emisje brutto GHG w zakresie 2 według metody market-based”; „27 528 Mg CO2e” i „26 528 Mg CO2e” jako sumy zakresu 1 i 2.',
            'Ta sama strona: „w 2025 r. przywróciliśmy do ponownego użytku 353 maszyny”.',
            'Raport ESG 2025: „W roku 2024 po raz pierwszy obliczono emisje Grupy Impel na potrzeby wewnętrzne, a w 2025 r. po raz pierwszy są publikowane wyniki obliczeń emisji z zakresu 1 i 2”. „W kolejnych latach Grupa Impel planuje rozszerzenie zakresu raportowania emisji GHG o zakres 3”. „zastosowano metodę estymacji ilościowej opartą na danych kosztowych” oraz „średnich cen jednostkowych netto dla roku 2025 na podstawie danych dostępnych w sieci”.'
        ) `
        $a06 `
        'Nagłówek grafiki zmienić tak, aby w pierwszym wersie był rok, zakres 1 i 2 oraz nazwa metody. Obok liczby dopisać listę spółek objętych rachunkiem i to, że energia elektryczna oraz ciepło są szacunkiem z faktur i cen za 2025 r. Sumy podać jak w tabeli raportu: 27 528,1 i 26 528,0 Mg CO2e. Zdanie o „wpływie klimatycznym organizacji” zawęzić do zakresów, które liczba obejmuje. 353 maszyny zostawić tylko ze stanem parku w 2025 r. albo zdjąć. Metodyki GHG Protocol nie chować wyłącznie w raporcie, gdy grafika jest na stronie działań.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Emisje brutto GHG Grupy Impel w 2025 r., metoda GHG Protocol: zakres 1 — 16 730,0 Mg CO2e; zakres 2 location-based — 10 798,1 Mg CO2e; zakres 2 market-based — 9 798,0 Mg CO2e. Suma zakresu 1 i 2: 27 528,1 Mg CO2e (location-based) i 26 528,0 Mg CO2e (market-based). Zużycie energii elektrycznej i ciepła w tej kalkulacji oszacowano z wartości faktur i średnich cen jednostkowych za 2025 r. Spółki objęte rachunkiem: [lista].”'

    $a07 = @'
Na stronie EcoVadis czytelnik odnosi wrażenie, że Impel Facility Services ma cel obniżenia emisji zakresu 1 i 2 o 20% do 2030 r. wobec 2025 r., oraz że Praxima Krakpol obniży emisję o 1 300 tCO2e, też o 20% do 2030 r. wobec 2025 r. Przy Facility Services jedyna podana wielkość jest za 2024 r.: 2 065 tCO2e i 1 477 tCO2e. Rok bazowy celu jest 2025, a tej wielkości na stronie nie ma. Przy Praximie 6 501,59 tCO2e stoi w zdaniu o raporcie za 2025 r., bez osobnego wskazania, że jest to wynik roku bazowego. W pliku ESG_Impel_raport_2025.pdf, podlinkowanym z działu ESG, tych trzech wielkości nie ma. Ten sam raport, w części E1-4, podaje, że grupa obecnie nie posiada celów związanych z łagodzeniem zmiany klimatu i z przystosowaniem się do niej. Jedyna „20%” w raporcie dotyczy innego przedmiotu: udziału energii elektrycznej pozostałych spółek objętego gwarancjami pochodzenia. Cel na przyszłość wolno opublikować, gdy w tym samym miejscu są mierzalny wskaźnik, termin, zasoby i regularna weryfikacja niezależnego eksperta, a wynik jest dostępny odbiorcy. Tego planu przy 20% do 2030 r. nie ma, a raport grupy celowi zaprzecza.
'@
    Add-Case 'PRZYPADEK 07: Cel dekarbonizacji 20% do 2030 r.' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · dyrektywa 2005/29/WE art. 6 ust. 2 lit. d | Poziom ryzyka: WYSOKIE' `
        @(
            'Strona EcoVadis, Impel Facility Services: „w raporcie zrównoważonego rozwoju przedstawiamy najważniejsze dane (…) Znajduje się w nim wyliczenie emisyjności naszej organizacji za rok 2024, która wyniosła: w zakresie 1 – 2065 tCO2e, w zakresie 2 – 1477 tCO2e. Jednocześnie wyznaczyliśmy cel dekarbonizacyjny do 2030 roku na poziomie 20%, który dotyczy obniżenia emisyjności organizacji w zakresach 1 i 2 w odniesieniu do bazowego 2025 r.”',
            'Ta sama strona, Praxima Krakpol: „W raporcie ESG za rok 2025 (…) wyliczenie emisyjności naszej organizacji, która wyniosła 6 501,59 tCO2e (zakres 1: 467,96 tCO2e, zakres 2: 6 033,63 tCO2e). Wyznaczyliśmy cel dekarbonizacyjny do 2030 roku na poziomie 20% (…) planujemy obniżyć emisyjność do 2030 roku o 1 300 tCO2e w odniesieniu do roku bazowego 2025”.',
            'Raport ESG 2025, E1-4: „Grupa Impel obecnie nie posiada celów związanych z łagodzeniem zmiany klimatu i przystosowaniem się do niej”.'
        ) `
        $a07 `
        'Zdjąć ze strony EcoVadis oba cele 20% do 2030 r. i wielkość 1 300 tCO2e. Wielkości 2 065, 1 477 i 6 501,59 tCO2e zdjąć do czasu, aż ten sam raport, do którego strona odsyła, poda je z rokiem, zakresem, listą źródeł i metodą. Nowy cel publikować dopiero z datą przyjęcia, rokiem bazowym, zasobami, nazwiskiem lub firmą weryfikatora i miejscem, w którym odbiorca widzi wynik. Do tego czasu strona E1-4 i strona EcoVadis mają mówić to samo.' `
        'BRAK WZORU DO PUBLIKACJI' `
        'Cel 20% do 2030 r. oraz wielkości 2 065 tCO2e, 1 477 tCO2e i 6 501,59 tCO2e schodzą ze strony. Wzorca zastępczego nie ma, dopóki raport grupy i strona EcoVadis podają sprzeczny stan celów.'

    $a08 = @'
W raporcie czytelnik odnosi wrażenie, że gwarancje pochodzenia pokryły 100% emisji z energii elektrycznej Impel S.A. i około 20% takich emisji w pozostałych spółkach, oraz że kolejne gwarancje mają dalej ograniczać emisje. W tej samej tabeli udział źródeł kopalnych wynosi 90,5%, odnawialnych 9,5%, a własna produkcja odnawialna 0 MWh. Przy pozycji odnawialnej w megawatogodzinach wydrukowano „8 113,9%”. Gwarancja pochodzenia jest instrumentem ujęcia rynkowego. Zdanie o „pokryciu 100% emisji” i o „ograniczaniu emisji” czyta się jak spadek emisji w miejscu zużycia. Praktyka Prezesa UOKiK w sprawie Bolt z 19 stycznia 2026 r. dotyczy właśnie energii odnawialnej opartej na certyfikacie atrybutu, bez rozróżnienia od energii fizycznej. Raport nazywa podejście rynkowe, ale nie podaje megawatogodzin gwarancji ani tego, że zakres 2 location-based zostaje 10 798,1 Mg CO2e. Różnica wobec metody rynkowej to 1 000,1 Mg CO2e na poziomie grupy, a nie 100% energii grupy.
'@
    Add-Case 'PRZYPADEK 08: Gwarancje pochodzenia jako pokrycie emisji' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE' `
        @(
            'Raport ESG 2025: „Grupa Impel zakupiła Gwarancje Pochodzenia w celu pokrycia emisji w zakresie 2 związanych z wykorzystaniem energii elektrycznej w ramach podejścia opartego na rynku. W przypadku spółki Impel S.A. instrument ten pozwolił na pokrycie 100% emisji wynikających z zużycia energii elektrycznej, natomiast w pozostałych spółkach Grupy pokryto około 20% tych emisji”.',
            'Ten sam akapit: „W kolejnych latach Grupa Impel planuje stopniowo zwiększać udział odnawialnej energii elektrycznej oraz wykorzystanie instrumentów rynkowych, takich jak Gwarancje Pochodzenia, w celu dalszego ograniczania emisji związanych z zużyciem energii elektrycznej”.',
            'Tabela koszyka: udział źródeł kopalnych 90,5% (77 492,9 MWh), udział odnawialnych 9,5%, własna produkcja odnawialna 0,0 MWh. W wierszu megawatogodzin odnawialnych wydrukowano „8 113,9%”.'
        ) `
        $a08 `
        'Przepisać akapit tak, aby 100% dotyczyło tylko zużycia energii elektrycznej Impel S.A. w ujęciu market-based, z rokiem 2025 i z megawatogodzinami gwarancji. Około 20% zostawić z listą spółek albo zastąpić dokładnym udziałem. Zdanie o dalszym ograniczaniu emisji przez kolejne gwarancje zdjąć. W tabeli poprawić „8 113,9%” na 8 113,9 MWh. Obok metody rynkowej zostawić location-based 10 798,1 Mg CO2e. Tego rozróżnienia nie drukować jako lekcji o systemie elektroenergetycznym. Ma wynikać z tego, którą liczbę strona podaje.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„W 2025 r. gwarancje pochodzenia objęły, w ujęciu market-based, 100% zużycia energii elektrycznej Impel S.A. ([X] MWh) i [Y]% zużycia energii elektrycznej pozostałych spółek grupy ([Z] MWh). Zakres 2 location-based grupy w 2025 r. wyniósł 10 798,1 Mg CO2e, market-based 9 798,0 Mg CO2e. Udział energii ze źródeł odnawialnych w zużyciu energii grupy wyniósł 9,5% (8 113,9 MWh). Własna produkcja energii odnawialnej: 0 MWh.”'

    $a09 = @'
Na stronie magazynów energii czytelnik odnosi wrażenie, że magazyn uniezależnia system od źródeł konwencjonalnych i przez to zmniejsza emisje, oraz że firma kładzie nacisk na ekologiczność i zrównoważony rozwój. Drugi akapit wiąże redukcję dwutlenku węgla z energią z odnawialnych źródeł zgromadzoną w magazynie, ale nagłówek „Redukcja emisji gazów cieplarnianych” stoi jak cecha produktu. Brakuje instalacji, roku, megawatogodzin i tego, skąd pochodzi energia ładowania. Ogólne „ekologiczność” i „zrównoważony rozwój” wymagają znaku typu I w tym samym miejscu. Takiego znaku na stronie nie ma. Magazyn oddaje energię, którą wcześniej pobrał. Sam pojemnik nie zmienia koszyka. W raporcie grupy za 2025 r. ten koszyk jest w 90,5% kopalny, a własna produkcja odnawialna wynosi 0 MWh. Produkcja baterii w tym zdaniu nie występuje.
'@
    Add-Case 'PRZYPADEK 09: Magazyn energii „ekologiczny” i redukujący emisje' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · ECGT, załącznik I pkt 4a | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'ipro.impel.pl, magazyny energii: „systemy energetyczne stają się bardziej elastyczne i mniej zależne od konwencjonalnych źródeł, co przyczynia się do redukcji emisji szkodliwych substancji”.',
            'Ta sama strona: „Redukcja emisji gazów cieplarnianych. Korzystanie z energii elektrycznej zgromadzonej w magazynie, która pochodzi z odnawialnych źródeł energii, pomaga w redukcji emisji dwutlenku węgla i innych gazów cieplarnianych, co przyczynia się do ochrony środowiska naturalnego”.',
            'Lista wyróżników: „Ekologiczne podejście jako firma, kładziemy duży nacisk na ekologiczność i zrównoważony rozwój będąc świadomymi wpływu swoich działań na środowisko”.'
        ) `
        $a09 `
        'Zdjąć „ekologiczne podejście”, „ekologiczność” i „zrównoważony rozwój”. Zdjąć nagłówek „Redukcja emisji gazów cieplarnianych” oraz zdanie o mniejszej zależności od źródeł konwencjonalnych jako cesze każdego magazynu. Jeśli spadek emisji ma zostać przy ofercie, liczyć go dla nazwanej instalacji, z rokiem, czynnikiem emisji i z produkcją baterii w granicach liczby. Test porównania z źródłem konwencjonalnym jest w rozdziale 4.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Magazyn energii [model] w [lokalizacja] w [roku] oddał [X] MWh energii wcześniej pobranej z [umowa albo instalacja, nazwa]. Różnica emisji wobec [scenariusz] wyniosła [Y] Mg CO2e. Metoda [nazwa], granice od produkcji baterii do oddania energii.”'

    $a10 = @'
Na stronie efektywności energetycznej czytelnik odnosi wrażenie, że strategia powstaje z uwzględnieniem zasady zrównoważonego rozwoju, a w rubryce „Co zyskujesz?” jest „Zmniejszenie emisji CO2”. Na stronie audytów energetycznych audyt „pozwoli spełnić wymogi ESG”. Brakuje obiektu, roku, tony CO2 i czynnika emisji. Spadek emisji jest skutkiem po wdrożeniu, a strona stawia go jak cechę zamówienia. Audyt według ustawy o efektywności energetycznej z 2016 r. jest opracowaniem o zużyciu energii. Sam dokument nie spełnia wymogów ESG klienta. „Zasada zrównoważonego rozwoju” w zdaniu o strategii jest ogólnym hasłem środowiskowym bez wskaźnika. Świadectwo efektywności energetycznej, wymienione obok jako możliwość, jest instrumentem tej ustawy. To zdanie może zostać, gdy nie stoi jako wynik środowiskowy.
'@
    Add-Case 'PRZYPADEK 10: Audyt energii jako spełnienie ESG i pewny spadek CO2' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 · ECGT, załącznik I pkt 4a | Poziom ryzyka: WYSOKIE (Czarna lista)' `
        @(
            'serwis-techniczny.impel.pl: „Tworzymy strategie energetyczne uwzględniając trendy technologiczne, prawne oraz zasadę zrównoważonego rozwoju”.',
            'Ta sama strona, „Co zyskujesz?”: „Zmniejszenie emisji CO2”.',
            'ipro.impel.pl, audyty energetyczne: „To właśnie na ich podstawie możliwe jest zaplanowanie działań, które realnie przełożą się na oszczędności, zwiększą niezależność energetyczną i pozwolą spełnić wymogi ESG”.'
        ) `
        $a10 `
        'Z listy korzyści zdjąć „Zmniejszenie emisji CO2”. Ze zdania o audycie zdjąć „pozwolą spełnić wymogi ESG” i „niezależność energetyczną”, dopóki nie ma udziału energii z własnej instalacji. „Zasadę zrównoważonego rozwoju” zdjąć ze zdania o strategii. Zostawić opis audytu jako analizy zużycia według ustawy z 2016 r. Wynik CO2 publikować po wdrożeniu, dla nazwanego obiektu i roku.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Audyt efektywności energetycznej [obiekt, rok] sporządzamy według ustawy o efektywności energetycznej z 2016 r. Po wdrożeniu zalecenia [nazwa] zużycie energii w [roku] wyniosło [X] MWh, a emisje [Y] Mg CO2e, czynnik [nazwa, rok czynnika].”'

    $a11 = @'
W raporcie czytelnik odnosi wrażenie, że środki czystości i papiery Kappa powstają zgodnie z normami jakości i bezpieczeństwa, wśród których wymienia ISO i Ecolabel. Brakuje nazwy produktu, numeru zezwolenia i daty. EU Ecolabel jest znakiem środowiskowym, a zdanie stawia go wśród norm bezpieczeństwa. Jedyne zezwolenie EU Ecolabel w badanym materiale ma numer PL/052/002 i dotyczy usługi utrzymania czystości Impel Facility Services, nie produkcji chemii Kappa. Nazwa znaku przy produktach Kappa przenosi na chemię znak, którego zakresu strona przy tych produktach nie pokazuje.
'@
    Add-Case 'PRZYPADEK 11: „Ecolabel” przy produktach Kappa' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: ŚREDNIE' `
        @(
            'Raport ESG 2025: „Kappa Sp. z o.o. (…) zajmuje się produkcją środków czystości, papierów higienicznych oraz rozwiązań higienicznych dla klientów biznesowych, wytwarzanych w polskich zakładach produkcyjnych zgodnie z normami jakości i bezpieczeństwa (m.in. ISO, Ecolabel)”.'
        ) `
        $a11 `
        'Ze zdania o Kappa zdjąć „Ecolabel” i zbiorcze „ISO”, dopóki przy nazwie handlowej produktu nie ma numeru zezwolenia albo numeru certyfikatu, jednostki i daty. Zezwolenia PL/052/002 nie dopisywać przy chemii. Karty charakterystyki, o których raport mówi niżej, zostawić jako informację o substancjach, bez przymiotnika środowiskowego.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Środek [nazwa handlowa] Kappa Sp. z o.o., zezwolenie EU Ecolabel nr [numer], ważne od [data] do [data], grupa produktowa [nazwa].”'

    $a12 = @'
W polityce środowiskowej czytelnik odnosi wrażenie, że dokument jest zgodny z wytycznymi Celów Zrównoważonego Rozwoju, United Nations Global Compact i Porozumienia Paryskiego, oraz że grupa wprowadza proekologiczne standardy zakupowe i promuje ekologiczne nawyki. Na stronie organizacji uczestnictwo w Global Compact jest opisane jako udział w inicjatywie firm działających zgodnie z zasadami zrównoważonego rozwoju. Na stronie EcoVadis srebrny medal Impel Facility Services wzmacnia, według strony zbiorczej, wiarygodność „naszych spółek”. Polityka nie podaje ścieżki temperatury, roku bazowego ani budżetu redukcji. Raport E1-4 mówi, że celów klimatycznych grupa nie posiada, a koszyk energii jest w 90,5% kopalny. „Zgodna z Porozumieniem Paryskim” czyta się jak cecha aktualna. Uczestnictwo w Global Compact jest zobowiązaniem do zasad, co następne zdanie strony organizacji samo dopowiada. Nie jest znakiem typu I. Srebrny medal EcoVadis jest oceną jednej spółki. Strona nie podaje roku ani wyniku punktowego, a zdanie zbiorcze mówi o spółkach w liczbie mnogiej. Medalu Praximy na tej stronie nie ma.
'@
    Add-Case 'PRZYPADEK 12: Porozumienie paryskie, Global Compact i medal EcoVadis' `
        'Podstawa prawna: UPNPR art. 5 i art. 6 | Poziom ryzyka: WYSOKIE' `
        @(
            'Polityka środowiskowa, § 3: „Polityka jest zgodna z wytycznymi, wynikającymi z Celów Zrównoważonego Rozwoju ONZ United Nations Global Compact oraz Porozumienia Paryskiego w sprawie zmian klimatu”.',
            'Ta sama polityka: „Wprowadzania proekologicznych standardów zakupowych i eksploatacyjnych”. „Organizacja akcji promujących ekologiczne nawyki”.',
            'Strona organizacji: „Jesteśmy uczestnikiem United Nations Global Compact – międzynarodowej inicjatywy ONZ skupiającej firmy działające zgodnie z zasadami zrównoważonego rozwoju”.',
            'Strona zbiorcza ESG: „Ocena EcoVadis wzmacnia wiarygodność naszych spółek”. Strona EcoVadis: „Spółka Impel Facility Services posiada srebrny medal w ratingu EcoVadis”.'
        ) `
        $a12 `
        'Z § 3 polityki zdjąć „zgodna z” Porozumieniem Paryskim, Celami Zrównoważonego Rozwoju i Global Compact. Zdjąć „proekologicznych” i „ekologiczne nawyki”. Na stronie organizacji zostawić fakt uczestnictwa i dziesięć zasad jako tekst inicjatywy, bez zdania, że uczestnik już działa zgodnie z zasadami zrównoważonego rozwoju. Medal EcoVadis publikować z rokiem i wynikiem, tylko przy Impel Facility Services. Zdanie o „naszych spółkach” zawęzić do tej spółki, dopóki druga spółka medalu na stronie nie ma.' `
        'DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)' `
        '„Impel S.A. jest uczestnikiem United Nations Global Compact od [rok]. Impel Facility Services Sp. z o.o. ma w ocenie EcoVadis za [rok] wynik [punkty] na 100, medal srebrny.”'

    Add-H '4. Porównanie środków, certyfikatu i magazynu energii: osobny test porównywalności i sprawdzalności'
    Add-P 'Rozdział 3 opisuje twierdzenia. Ten rozdział sprawdza tylko porównywalność i sprawdzalność zestawień, które materiał sam ogłasza: środki czystości wobec nieopisanej chemii, pozycję EU Ecolabel wobec innych firm sprzątających oraz magazyn energii wobec źródeł konwencjonalnych.' 10.5 $false $false $DARK 0 6

    $comp = @(
        @('Środki. „środki czystości, które w mniejszym stopniu ingerują w środowisko naturalne i posiadają odpowiedni certyfikat, np. EU Ecolabel”. Strona sprzątania obiektów stawia to zdanie jako wyróżnik usługi.', 'Brak nazwy środka Impel i nazwy środka odniesienia, braku roku, jednostki i metody. „Na przykład” nie wskazuje, jaki udział środków ma EU Ecolabel, a jaki inny certyfikat. Zezwolenie PL/052/002 dotyczy usługi utrzymania czystości w obiektach handlowych i biurowych, nie stopnia ingerencji każdego środka w ofercie. Porównanie nie spełnia art. 16 ust. 3 UZNK: nie da się go sprawdzić na stronie i nie zestawia nazwanych produktów.'),
        @('Pozycja na rynku. Raport: „Grupa Impel jest drugą firmą w Polsce z branży usług utrzymania czystości, która legitymuje się tym certyfikatem”. W tym akapicie numer to PL/052/002/001.', 'Brak źródła zestawienia, daty liczenia i grupy, z którą Impel jest porównywany: wszystkie firmy sprzątające, firmy, które złożyły wniosek, czy firmy z zezwoleniem na dzień publikacji. PL/052/002/001 jest świadectwem dla centrali mBanku. Świadectwo główne ma numer PL/052/002/005-1, a numer zezwolenia to PL/052/002. Miejsce „druga” nie da się sprawdzić z informacji w raporcie.'),
        @('Magazyn. „systemy energetyczne stają się bardziej elastyczne i mniej zależne od konwencjonalnych źródeł, co przyczynia się do redukcji emisji szkodliwych substancji”. Osobno: redukcja emisji, „która pochodzi z odnawialnych źródeł energii”.', 'Brak roku, megawatogodzin i koszyka, z którym magazyn jest zestawiony. Magazyn oddaje energię pobraną wcześniej. Koszyk grupy za 2025 r. to 90,5% źródeł kopalnych i 0 MWh własnej produkcji odnawialnej. Produkcja baterii w zdaniu nie występuje. Zestawienie nie nazywa instalacji ani scenariusza odniesienia, więc nie jest sprawdzalne.')
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
    Add-P 'Praktyka naruszająca zbiorowe interesy konsumentów (art. 24 ustawy o ochronie konkurencji i konsumentów) jest zagrożona karą do 10% obrotu z roku poprzedzającego nałożenie kary, za każdą praktykę (art. 106). Zarzuty z 19 stycznia 2026 r. wobec Tchibo i Zary dotyczą ogólnego „eko” i znaku czytanego szerzej niż jego zakres. Zarzut wobec Bolta dotyczy energii odnawialnej opartej na certyfikacie atrybutu. Wzorce z impel.pl są z tej grupy: hasło „przyjazne środowisku”, zezwolenie EU Ecolabel czytane na sprzątanie w ogóle, gwarancje pochodzenia opisane jak pokrycie emisji oraz własny certyfikat dostawcy. ECGT stosuje się od 27 września 2026 r. do treści online. Opóźnienie polskiej ustawy (druk Senatu 825, projektowana data 27 września 2027 r., akt nieogłoszony) nie wstrzymuje art. 5 i 6 UPNPR. Sprzeczność celu 20% do 2030 r. ze zdaniem raportu, że celów klimatycznych nie ma, jest daną, którą pracownik albo dostawca może zgłosić do Prezesa UOKiK.' 10.5 $false $false $DARK 0 8
    Add-P '• 2. Reklama porównawcza i znak zaufania' 11 $true $false $NAVY 8 2
    Add-P 'Zdanie o mniejszej ingerencji środków oraz zdanie o „drugiej firmie w Polsce” pozwalają rozpoznać cudzą chemię i cudze firmy sprzątające. Artykuł 16 ust. 3 UZNK wymaga, aby porównanie było rzetelne, sprawdzalne i dotyczyło cech istotnych, typowych i dających się zbadać. Artykuł 18 daje konkurentowi roszczenie o zaniechanie i naprawienie szkody. Osobno art. 7 pkt 2 UPNPR dotyczy znaku zaufania użytego bez aktualnego uprawnienia. Dwa pliki ISO 14001 na stronie standardów mają datę ważności wcześniejszą niż 4 listopada 2026 r., a strona nadal podaje je jako potwierdzenie standardu.' 10.5 $false $false $DARK 0 8
    Add-P '• 3. Łańcuch klientów i ESRS' 11 $true $false $NAVY 8 2
    Add-P 'Odbiorcą witryny jest firma, która może wpisać dostawcę do własnego raportu ESRS E1. Grafika 16 730,0 Mg CO2e, zdanie o pokryciu 100% emisji Impel S.A. gwarancjami pochodzenia oraz cel 20% do 2030 r. mogą wejść do danych o łańcuchu wartości klienta. Raport grupy sam podaje, że zakres 3 nie jest jeszcze liczony, że część energii jest szacunkiem z cen z internetu i że celów klimatycznych nie ma. Klient, który powtórzy skrót ze strony EcoVadis, przenosi sprzeczność na własny raport. Korekta na impel.pl ogranicza to przeniesienie.' 10.5 $false $false $DARK 0 8
    Add-P '• 4. Filar społeczny i kanał zgłoszeń' 11 $true $false $NAVY 8 2
    Add-P 'Szkolenie BHP, ocena ryzyka i przeciwdziałanie dyskryminacji są obowiązkami Kodeksu pracy. Platforma sygnalistów u dużego przedsiębiorcy jest obowiązkiem ustawy z 14 czerwca 2024 r. Wpisanie ich w listę działań ESG jest praktyką z punktu 10a. Ustawa o sygnalistach otwiera też drogę zewnętrzną do Prezesa UOKiK w sprawach ochrony konsumentów. Rozbieżność między listą „Co robimy” a raportem E1-4 jest zgłoszeniem, które ta droga obejmuje.' 10.5 $false $false $DARK 0 8

    Add-H '6. Rekomendacje naprawcze'
    Add-Callout 'ETAP 1: NATYCHMIASTOWY PULL-OFF (natychmiast)' @(
        'Zdjąć ze strony głównej „Usługi przyjazne środowisku”, zdanie o środkach bezpiecznych dla środowiska oraz baner z turbinami wiatrowymi jako tło oferty grupy.',
        'Zdjąć „środki przyjazne środowisku” i „w mniejszym stopniu ingerują” ze strony sprzątania. Zdjąć nazwę „Certyfikat Kwalifikowanego Dostawcy w zrównoważonym łańcuchu wartości” i słowo „certyfikat” przy dostawcach.',
        'Zdjąć ze strony standardów pliki ISO 14001 Climbex (ważny do 14 lipca 2026 r.) oraz Impel Rental i Praxima Krakpol (ważne do 31 lipca 2026 r.). Zdjąć cel 20% do 2030 r. i wielkości 2 065, 1 477 oraz 6 501,59 tCO2e ze strony EcoVadis. Z raportu zdjąć „pokrycie 100% emisji” i „drugą firmę w Polsce”, dopóki nie ma wzoru z przypadku 08 i źródła zestawienia.'
    ) (Rgb 255 241 242) $DANGER

    Add-Callout 'ETAP 2: KOREKTA STRON I RAPORTU (W ciągu 7 dni od audytu)' @(
        'Wstawić jedno zdanie o zezwoleniu EU Ecolabel nr PL/052/002, z datami 16 lipca 2025 r. – 31 grudnia 2027 r., z wyłączeniem dezynfekcji i aromatyzacji oraz z nazwą Dział Kontraktów EKO. Wyłączenie zostaje, bo strona sprzątania obiektów oferuje „dezynfekcję powierzchni”, a menu witryny czystosc.impel.pl zawiera „Usługi DDD – dezynsekcja, deratyzacja, dezynfekcja”. Numer PL/052/002/001 zostawić przy mBanku.',
        'Na grafice śladu powtórzyć zakres 1 i 2, rok 2025, sumy 27 528,1 i 26 528,0 Mg CO2e oraz zdanie o szacunku energii z faktur. Uzgodnić stronę EcoVadis ze zdaniem E1-4, że celów klimatycznych nie ma.',
        'Przy Polityce BHP, Polityce różnorodności i przy procedurze zgłoszeń wstawić zdania, które nazywają spółki Grupy Impel i odsyłają do pliku na impel.pl/esg/polityki/ albo do adresu platformy. Z list „Co robimy” i z celu „Mniej nierówności” zdjąć te czynności jako osiągnięcia ESG. Z polityki środowiskowej zdjąć „zgodna z Porozumieniem Paryskim”. Z oferty magazynu i audytu zdjąć „ekologiczność”, „redukcję emisji” i „spełnić wymogi ESG”. Przy Kappa zdjąć „Ecolabel”.'
    ) (Rgb 255 247 237) $WARNING

    Add-Callout 'ETAP 3: METODYKA I DOWÓD W TYM SAMYM KANALE (W ciągu 30 dni od audytu)' @(
        'Przy śladzie opublikować listę spółek, granicę organizacyjną i informację, które pozycje są szacunkiem z cen za 2025 r. Gwarancje pochodzenia podać w megawatogodzinach, osobno od koszyka 90,5% / 9,5% i od 0 MWh własnej produkcji.',
        'Jeśli cel redukcji ma wrócić, opublikować datę przyjęcia, rok bazowy, zasoby, weryfikatora i miejsce wyniku. Do tego dnia strona i raport mają jeden stan: cel jest albo go nie ma.',
        'Nowy certyfikat ISO 14001 publikować z numerem, zakresem i datą ważności późniejszą niż dzień wstawienia pliku. Medal EcoVadis publikować z rokiem i wynikiem, przy jednej nazwanej spółce. 353 maszyny podać wobec stanu parku.'
    ) (Rgb 239 246 255) $ACCENT

    Add-Callout 'ETAP 4: SŁOWNIK I STAŁY PRZEGLĄD (po zamknięciu etapu 3)' @(
        'Zatwierdzić słownik zdań dopuszczonych do publikacji: numer zezwolenia, spółka, obiekt, rok, zakres 1 albo 2, metoda. Hasło nie przechodzi na „przyjazne środowisku”, „ekologiczność”, „zrównoważony łańcuch” ani na cały outsourcing.',
        'Każdy plik certyfikatu ma właściciela daty ważności. Plik schodzi ze strony w dniu utraty ważności. Własnej oceny dostawcy nie nazywać certyfikatem.',
        'Obowiązek BHP, równego traktowania i kanału zgłoszeń opisuje się przy polityce albo przy adresie procedury, ze spółką i z odesłaniem do dokumentu. Na liście osiągnięć ESG i przy celu zrównoważonego rozwoju to zdanie nie stoi. Wersja strony usługowej (czystosc, serwis techniczny, ipro) nie zostaje w tyle za korektą impel.pl.'
    ) (Rgb 240 253 250) $SUCCESS

    Add-H '7. Wnioski'
    Add-P 'Witryna impel.pl w stanie na 4 listopada 2026 r. nie nadaje się do dalszej publikacji w obecnych hasłach środowiskowych. „Usługi przyjazne środowisku”, zezwolenie EU Ecolabel czytane szerzej niż obiekty handlowe i biurowe, własny certyfikat dostawcy, wygasłe ISO 14001 oraz obowiązki BHP i sygnalistów wpisane w ESG wyczerpują punkty 4a, 4b, 2a i 10a załącznika I. Korekta polega na zawężeniu każdego zdania do spółki, numeru, roku i zakresu albo na jego usunięciu, oraz na uzgodnieniu strony EcoVadis z raportem, który celów klimatycznych nie podaje i emisji offsetem nie kompensuje.' 10.5 $false $false $DARK 0 8

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
    if ($word -ne $null) { try { $word.Quit() } catch {} }
}
