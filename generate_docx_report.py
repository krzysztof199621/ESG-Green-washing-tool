#!/usr/bin/env python3
"""Raport audytu greenwashingu LIBRA PARTNERS — forma wg INSTRUKCJA_GENEROWANIA_RAPORTOW.md.

Wersja pełna i wersja okrojona (jedno naruszenie, jeden przypadek; w pozostałych
pozycjach listy czytelne jest tylko ryzyko, a punkty 4–6 pokazują po jednej pozycji).
"""

from io import BytesIO
from pathlib import Path

import docx
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
LOGO = ROOT.parent / "ESG Hero Day - Materialy Druk" / "assets" / "esg_logo_white.png"
SCREENSHOT = ROOT / "assets" / "libra_partners_strona.png"
_FONT = "/System/Library/Fonts/Supplemental/Arial.ttf"
_FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"


def _hex_rgb(value):
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def blurred_png(text, width_cm, size_pt=8, bold=False, bg="FFFFFF"):
    """Tekst tej samej długości co w raporcie pełnym, nieczytelny po rozmyciu."""
    scale = 2
    dpi = 96 * scale
    width_px = max(28, int(width_cm / 2.54 * dpi))
    pad = int(5 * scale)
    font_px = max(9, int(size_pt * dpi / 72))
    font = ImageFont.truetype(_FONT_BOLD if bold else _FONT, font_px)
    max_w = max(8, width_px - 2 * pad)
    lines, current = [], ""
    for word in text.replace("\n", " \n ").split(" "):
        if word == "\n":
            lines.append(current)
            current = ""
            continue
        trial = word if not current else current + " " + word
        if font.getlength(trial) <= max_w:
            current = trial
            continue
        if current:
            lines.append(current)
        if font.getlength(word) <= max_w:
            current = word
            continue
        chunk = ""
        for char in word:
            if font.getlength(chunk + char) <= max_w:
                chunk += char
            else:
                if chunk:
                    lines.append(chunk)
                chunk = char
        current = chunk
    if current or not lines:
        lines.append(current)
    line_h = int(font_px * 1.28)
    height_px = pad * 2 + line_h * max(1, len(lines))
    image = Image.new("RGB", (width_px, height_px), _hex_rgb(bg))
    draw = ImageDraw.Draw(image)
    y = pad
    for line in lines:
        if line:
            draw.text((pad, y), line, font=font, fill=(22, 36, 44))
        y += line_h
    image = image.filter(ImageFilter.GaussianBlur(radius=max(6, int(font_px * 0.48))))
    image = ImageOps.autocontrast(image, cutoff=1)
    stream = BytesIO()
    image.save(stream, format="JPEG", quality=74, optimize=True)
    stream.seek(0)
    return stream


def generate_report(out_path=None, preview=False):
    doc = docx.Document()

    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)
        section.header_distance = Cm(0.45)
        section.footer_distance = Cm(0.4)

    PRIMARY = RGBColor(8, 24, 31)
    NAVY = RGBColor(8, 24, 31)
    ACCENT = RGBColor(42, 138, 176)
    DANGER = RGBColor(190, 18, 60)
    WARNING = RGBColor(180, 83, 9)
    SUCCESS = RGBColor(4, 120, 87)
    DARK = RGBColor(22, 36, 44)
    MUTED = RGBColor(90, 107, 119)
    WHITE = RGBColor(255, 255, 255)
    CONTENT = Cm(17.6)

    def shade(cell, hex_color):
        cell._tc.get_or_add_tcPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        )

    def margins(cell, top=80, bottom=80, left=90, right=90):
        cell._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'<w:top w:w="{top}" w:type="dxa"/>'
                f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                f'<w:left w:w="{left}" w:type="dxa"/>'
                f'<w:right w:w="{right}" w:type="dxa"/>'
                f'</w:tcMar>'
            )
        )

    def borders(cell, color="E2E8F0", sz="4", left_sz=None, left_color=None):
        lc = left_color or color
        ls = left_sz or sz
        cell._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
                f'<w:left w:val="single" w:sz="{ls}" w:space="0" w:color="{lc}"/>'
                f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
                f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
                f'</w:tcBorders>'
            )
        )

    def set_run(run, size, bold=False, italic=False, color=DARK, font="Arial"):
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), font)

    def add_p(text, size=10.5, bold=False, italic=False, color=DARK, before=0, after=8, space=1.12):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = space
        set_run(p.add_run(text), size, bold, italic, color)
        return p

    def add_rich(parts, size=10.5, before=0, after=8):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = 1.12
        for text, bold, italic, color in parts:
            set_run(p.add_run(text), size, bold, italic, color)
        return p

    def add_h(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        set_run(p.add_run(text), 14, True, False, NAVY)
        return p

    def callout(title, lines, bg, border, title_color, blur=False):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        cell = table.cell(0, 0)
        cell.width = CONTENT
        shade(cell, bg)
        margins(cell, 120, 120, 140, 140)
        borders(cell, border, "6", "24", border)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0 if blur else 4)
        if blur:
            stream = blurred_png(title + "\n" + "\n".join(lines), 16.7, size_pt=10, bg=bg)
            p.add_run().add_picture(stream, width=Cm(16.7))
        else:
            set_run(p.add_run("■ " + title), 10, True, False, title_color)
            for line in lines:
                bp = cell.add_paragraph()
                bp.paragraph_format.space_before = Pt(2)
                bp.paragraph_format.space_after = Pt(2)
                bp.paragraph_format.line_spacing = 1.12
                set_run(bp.add_run(line), 10, False, False, DARK)
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(6)

    def lock_table(table, widths):
        tbl = table._tbl
        tblPr = tbl.tblPr
        for tag in ("w:tblLayout", "w:tblW"):
            found = tblPr.find(qn(tag))
            if found is not None:
                tblPr.remove(found)
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tblPr.append(layout)
        total = 0
        dxas = []
        for width in widths:
            dxa = int(round(width.cm / 2.54 * 1440))
            dxas.append(dxa)
            total += dxa
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:w"), str(total))
        tblW.set(qn("w:type"), "dxa")
        tblPr.append(tblW)
        grid = tbl.find(qn("w:tblGrid"))
        if grid is None:
            grid = OxmlElement("w:tblGrid")
            tblPr.addnext(grid)
        for child in list(grid):
            grid.remove(child)
        for dxa in dxas:
            col = OxmlElement("w:gridCol")
            col.set(qn("w:w"), str(dxa))
            grid.append(col)
        for row in table.rows:
            for cell, dxa in zip(row.cells, dxas):
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(qn("w:w"), str(dxa))
                tcW.set(qn("w:type"), "dxa")

    def page_field(paragraph):
        run = paragraph.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = " PAGE "
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.append(begin)
        run._r.append(instr)
        run._r.append(end)
        set_run(run, 8, False, False, MUTED)

    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)
        htable = header.add_table(1, 2, CONTENT)
        htable.autofit = False
        left, right = htable.rows[0].cells
        left.width, right.width = Cm(8.2), Cm(9.4)
        for cell in (left, right):
            shade(cell, "08181F")
            margins(cell, 60, 60, 120, 120)
            borders(cell, "08181F", "2", "2", "08181F")
        lp = left.paragraphs[0]
        lp.paragraph_format.space_before = lp.paragraph_format.space_after = Pt(0)
        lp.add_run().add_picture(str(LOGO), width=Cm(4.4))
        rp = right.paragraphs[0]
        rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rp.paragraph_format.space_before = Pt(8)
        rp.paragraph_format.space_after = Pt(0)
        set_run(rp.add_run("AUDYT RYZYK GREENWASHINGU"), 8, True, False, WHITE)
        rp2 = right.add_paragraph()
        rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rp2.paragraph_format.space_before = Pt(0)
        rp2.paragraph_format.space_after = Pt(0)
        set_run(rp2.add_run("esginstitute.eu"), 8, False, False, RGBColor(129, 192, 223))

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fp.paragraph_format.space_before = Pt(2)
        edition = "wersja okrojona  ·  " if preview else ""
        set_run(fp.add_run("ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Libra Partners  ·  " + edition + "str. "), 8, False, False, MUTED)
        page_field(fp)

    add_p(
        "ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO",
        size=9, bold=True, color=ACCENT, after=2,
    )
    add_p("Audyt Ryzyk Greenwashingu: LIBRA PARTNERS", size=20, bold=True, color=PRIMARY, before=2, after=2)
    add_p(
        "Weryfikacja twierdzeń środowiskowych, haseł, specyfikacji, tabeli porównawczej i kalkulacji CO₂ "
        "ze zbioru materiałów strony libra-partners.pl (stan zebrania: 26 września 2026 r.) "
        "wobec dyrektywy (UE) 2024/825 (ECGT), UPNPR, ustawy o ochronie konkurencji i konsumentów, "
        "UZNK oraz rozporządzenia (UE) 2026/1030 (CountEmissionsEU).",
        size=10.5, italic=True, color=MUTED, after=8,
    )
    if preview:
        add_p(
            "Wersja okrojona. W całości widoczne jest jedno naruszenie i jeden przypadek. "
            "W pozostałych wierszach tabeli zbiorczej czytelne jest tylko ryzyko. "
            "W punkcie 4 widoczny jest jeden wiersz, w punkcie 5 jeden punkt, w punkcie 6 jedna rekomendacja. "
            "Reszta tych części jest rozmyta.",
            size=9, italic=True, color=MUTED, after=8,
        )
    shot = doc.add_paragraph()
    shot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shot.paragraph_format.space_before = Pt(2)
    shot.paragraph_format.space_after = Pt(2)
    shot.add_run().add_picture(str(SCREENSHOT), width=Cm(15.4))
    add_p(
        "Zrzut strony głównej libra-partners.pl z 7 października 2026 r. Widoczne hasło: „Technologia z myślą o Ekologii”.",
        size=8.5, italic=True, color=MUTED, after=10,
    )

    meta = [
        ("Badany podmiot", "LIBRA PARTNERS, grupa SZARPOL (Centrum Produkcyjno-Handlowe „Szarpol”), ul. prof. Mariana Rapackiego 3, 98-200 Sieradz"),
        ("Przedmiot audytu", "Zbiór twierdzeń ze strony głównej, podstron oferty, „Zrównoważonego rozwoju”, „O nas”, specyfikacji modeli i tabeli porównawczej z paletą drewnianą. Stan zebrania: 26 września 2026 r."),
        ("Sektor działalności", "Opakowania transportowe: palety z tworzyw sztucznych (rHDPE, HDPE, rPP) dla logistyki, przemysłu spożywczego, farmaceutycznego i automotive"),
        ("Reżim prawny", "Przekaz na libra-partners.pl: dyrektywa (UE) 2024/825 (stosowanie od 27.09.2026), dyrektywa 2005/29/WE, UPNPR art. 5–7, u.o.k.k. art. 24 i 106, UZNK art. 16 i 18, rozporządzenie (UE) 2026/1030 i EN ISO 14083:2023; kontekst wylesiania: rozporządzenie (UE) 2023/1115 (EUDR). Wymogi dla palety w kraju przeznaczenia, w tym fitosanitarne (ISPM 15), zależą od kraju eksportu. W materiale: Wielka Brytania oraz eksport morski do Ameryki Północnej i Południowej. Prawa tych krajów o twierdzeniach środowiskowych nie wchodzą do tej oceny."),
        ("Data audytu", "7 października 2026 r."),
        ("Werdykt operacyjny", "NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)"),
    ]
    meta_table = doc.add_table(rows=len(meta), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    widths = (Cm(4.2), Cm(13.4))
    for i, (k, v) in enumerate(meta):
        ck, cv = meta_table.rows[i].cells
        ck.width, cv.width = widths
        shade(ck, "F8FAFC")
        shade(cv, "FFF1F2" if i == 5 else "FFFFFF")
        margins(ck, 50, 50, 80, 80)
        margins(cv, 50, 50, 80, 80)
        borders(ck)
        borders(cv)
        ck.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pk, pv = ck.paragraphs[0], cv.paragraphs[0]
        pk.paragraph_format.space_before = pk.paragraph_format.space_after = Pt(1)
        pv.paragraph_format.space_before = pv.paragraph_format.space_after = Pt(1)
        set_run(pk.add_run(k), 8.5, True, False, DARK)
        set_run(pv.add_run(v), 8.5, i == 5, False, DANGER if i == 5 else DARK)

    add_p("", size=6, after=8)

    callout(
        "WERDYKT: NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)",
        [
            "Materiał spełnia przesłanki czarnej listy załącznika I do dyrektywy 2005/29/WE w brzmieniu nadanym dyrektywą (UE) 2024/825: pkt 4a (ogólne hasła „ekologiczne”, „eko”, „ESG Friendly”, „zrównoważony” bez uznanej wysokiej efektywności ekologicznej), pkt 4b (Zero Waste i recyklat przypisane całej ofercie oraz całej działalności), pkt 10a (brak fumigacji ISPM 15 jako zaleta ekologiczna) i pkt 23g (krotność trwałości „10 razy” bez weryfikowalnych, porównywalnych kryteriów).",
            "Te zakazy stosuje się od 27 września 2026 r. Polska nowela (druk Senatu 825) przewiduje wejście w życie 27 września 2027 r. i nie jest jeszcze ogłoszona; opóźnienie transpozycji nie zawiesza oceny według ECGT ani odpowiedzialności z art. 5 i 6 UPNPR.",
            "Osobno materiał wprowadza w błąd liczbowo: kalkulacja oparta na EN 16258:2012 i liniowo przeniesiona z 33 palet na relacji Sieradz–Paryż na wolumen zakupu oraz ekwiwalent 600 000 drzew. Tabela porównawcza dyskredytuje paletę drewnianą (UZNK art. 16 ust. 3). Utrzymanie treści na stronie — kanał online bez taryfy zapasów (CPC, czerwiec 2026) — otwiera drogę do zarzutów Prezesa UOKiK (art. 24 i 106 u.o.k.k., do 10% obrotu) i roszczeń konkurencji (UZNK art. 18).",
            "Ta ocena dotyczy przekazu opublikowanego w Polsce. Przepisy, które dotyczą palety jako towaru wwożonego do innego państwa — w szczególności wymogi fitosanitarne opakowań (ISPM 15) i warunki dopuszczenia na rynek — zależą od kraju eksportu. Materiał sam wskazuje Wielką Brytanię oraz eksport morski do Ameryki Północnej i Południowej. Prawa tych krajów o twierdzeniach środowiskowych nie ma w bazie wiedzy narzędzia i nie zostały tu zastosowane.",
        ],
        "FFF1F2", "E11D48", DANGER,
    )

    add_h("1. Analiza Wstępna")
    add_p(
        "LIBRA PARTNERS należy do grupy SZARPOL z Sieradza i produkuje palety o wymiarze 1200 × 800 mm. "
        "W ofercie są modele z regranulatu HDPE (MW01, 9 kg), linia ESD, paleta higieniczna RS-01-FT z HDPE pierwotnego (12 kg), "
        "paleta antybakteryjna MW01-MED z regranulatu HDPE z dodatkiem tlenku cynku oraz paleta automotive z mieszanki PP z recyklingu (11 kg). "
        "Narracja strony ustawia te nośniki jako „ekologiczne palety”, następcę palety drewnianej, model Zero Waste i najprostszy sposób "
        "ograniczenia śladu węglowego oraz „zgodności z ESG”."
    )
    add_p(
        "Część faktów materiałowych da się obronić po zawężeniu języka: niższa masa wybranych modeli, deklarowane obciążenia według ISO 8611, "
        "odbiór zużytych palet kupionych od producenta, wynik badania dwóch szczepów według ISO 22196:2011. "
        "Komunikacja wychodzi poza ten zakres. Hasła ogólne nie mają pokrycia w EU Ecolabel ani w oznakowaniu ISO 14024 typu I. "
        "Twierdzenia o braku jakichkolwiek odpadów, o tym, że rozwiązania „nigdy nie trafiają na wysypisko”, oraz o surowcu z recyklingu i masie 9 kg "
        "obejmują całą ofertę, podczas gdy specyfikacje modeli temu przeczą. Brak fumigacji i anonimowe opinie „jesteśmy bardziej ekologiczni” "
        "przerabiają wyłączenie spod ISPM 15 na cechę środowiskową."
    )
    add_p(
        "Warstwa liczbowa jest odrębnym ryzykiem. Symulacja 400 kg CO₂ miesięcznie dla 33 palet, liczona według EN 16258:2012, "
        "zostaje przemnożona przez wolumen zakupu (do „121,120 ton” przy 10 000 sztuk) i zestawiona z „303 t CO₂” oraz z „600 000 drzew” "
        "przy założeniu 0,5 kg CO₂ na drzewo miesięcznie. Tabela porównawcza przypisuje palecie drewnianej negatywny wpływ na politykę ESG, "
        "trudny recykling i związek z ubytkiem 3,3 mln ha lasów w latach 2010–2015. To praktyki, które Prezes UOKiK już kwestionuje "
        "w sprawach kurierów i sprzedaży (InPost, Allegro, DHL, DPD, Bolt, Tchibo, Zara): ogólne „eko”, ekwiwalent drzew, zero waste i SDG bez skali."
    )

    add_h("2. Analiza szczegółowa - tabela zbiorcza naruszeń")

    findings = [
        ("01", "Kalkulator: 33 palety = 400 kg CO₂; 10 000 szt. = 121,120 t; 303 t miesięcznie", "UPNPR art. 5 ust. 2 pkt 2 i art. 6 · rozp. 2026/1030 · EN ISO 14083", "BARDZO WYSOKIE", "Liniowe przeniesienie jednego scenariusza trasowego na zakup; metodyka inna niż EN ISO 14083"),
        ("02", "„600 000 drzew” przy 0,5 kg CO₂ na drzewo miesięcznie", "UPNPR art. 5 ust. 2 pkt 2 · wzorzec UOKiK (InPost, Allegro)", "BARDZO WYSOKIE", "Ekwiwalent drzew bez metodyki; liczba niespójna z „303 t”"),
        ("03", "Paleta drewniana: trudny recykling, wycinka, „niekorzystny wpływ na politykę ESG”", "UZNK art. 16 ust. 1 pkt 2 i ust. 3 pkt 2, 3 i 5 · UPNPR art. 5 ust. 2 pkt 3", "BARDZO WYSOKIE", "Reklama porównawcza dyskredytująca; dane o lasach bez związku z paletą"),
        ("04", "„Ekologiczne palety”, „eko palety”, „ESG Friendly”, „zrównoważony rozwój”", "ECGT zał. I pkt 4a · UPNPR art. 5", "WYSOKIE", "Ogólne hasło bez EU Ecolabel i bez ISO 14024 typu I"),
        ("05", "„Zero Waste”, „nie ma jakichkolwiek odpadów”, „nigdy nie trafiają na wysypisko”", "ECGT zał. I pkt 4b · UPNPR art. 5 i 6", "WYSOKIE", "Claim o całej działalności przy programie odbioru własnych palet"),
        ("06", "Cała oferta „HDPE z recyklingu” i „tylko 9 kg”", "ECGT zał. I pkt 4b · UPNPR art. 6", "WYSOKIE", "RS-01-FT: HDPE pierwotny, 12 kg; automotive: rPP, 11 kg; ESD bez deklaracji recyklatu"),
        ("07", "„10 razy bardziej trwałe”; 60–70 cykli wobec 6–7 albo 7–8", "ECGT zał. I pkt 23g · UZNK art. 16 ust. 3 pkt 2 i 3", "WYSOKIE", "Krotność trwałości bez badania możliwego do sprawdzenia; dwie różne liczby w jednym materiale"),
        ("08", "Brak fumigacji ISPM 15 jako cecha ekologiczna", "ECGT zał. I pkt 10a · UPNPR art. 5", "WYSOKIE", "Wyłączenie spod normy fitosanitarnej dla drewna sprzedane jako przewaga „eko”"),
        ("09", "„Zgodne z polityką UE oraz ESG”; Agenda 2030; „najprostsze rozwiązanie” raportowania", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5", "WYSOKIE", "Zgodność ESG i cele SDG bez zakresu; wzorzec jak w zarzutach wobec Zary"),
        ("10", "LCA Envirly: „prawie dwukrotnie niższa emisja CO₂ w łańcuchu dostaw”", "UPNPR art. 6 · ISO 14040 / 14044", "ŚREDNIE", "Brak celu, granic, jednostki funkcjonalnej i dostępu do streszczenia"),
        ("11", "PZH, ISO 9001, opinia Politechniki i ISO 22196 obok „ESG Friendly” oraz „99% przez cały cykl życia”", "UPNPR art. 5 ust. 3 pkt 2 i art. 6", "ŚREDNIE", "Transfer atestu i opinii na korzyść środowiskową; uogólnienie badania dwóch szczepów"),
        ("12", "MW01-MED z ZnO przy haśle 100% przerobu i GOZ", "UPNPR art. 6 · ECGT zał. I pkt 4b", "ŚREDNIE", "Dodatek funkcjonalny bez informacji, czy zmienia odbiór i recykling"),
        ("13", "Anonimowe opinie: „będziemy bardziej EKO”, „jesteśmy bardziej ekologiczni”", "UPNPR art. 7 pkt 25 i 26", "NISKIE", "Opinie bez wskazania autora i bez informacji o weryfikacji"),
    ]

    sum_table = doc.add_table(rows=1 + len(findings), cols=5)
    sum_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_table.autofit = False
    sum_table.style = "Table Grid"
    col_w = (Cm(1.1), Cm(5.5), Cm(4.3), Cm(2.3), Cm(4.4))
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota naruszenia"]
    for c, h in enumerate(headers):
        cell = sum_table.rows[0].cells[c]
        cell.width = col_w[c]
        shade(cell, "08181F")
        margins(cell, 50, 50, 50, 50)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
        set_run(p.add_run(h), 8, True, False, RGBColor(255, 255, 255))
    for r, row in enumerate(findings):
        bg = "FFFFFF" if r % 2 == 0 else "F8FAFC"
        for c, val in enumerate(row):
            cell = sum_table.rows[r + 1].cells[c]
            cell.width = col_w[c]
            shade(cell, bg)
            margins(cell, 40, 40, 50, 50)
            borders(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
            if preview and r > 0 and c != 3:
                inner = max(0.7, col_w[c].cm - 0.32)
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
                p.add_run().add_picture(
                    blurred_png(val, inner, size_pt=7.5, bold=(c == 0), bg=bg),
                    width=Cm(inner),
                )
            else:
                run = p.add_run(val)
                if c == 3:
                    color = SUCCESS if val == "NISKIE" else WARNING if val == "ŚREDNIE" else DANGER
                    set_run(run, 7.5, True, False, color)
                elif c == 0:
                    set_run(run, 8, True, False, NAVY)
                else:
                    set_run(run, 7.5, False, False, DARK)
    if preview:
        lock_table(sum_table, col_w)

    add_p("", size=4, after=4)
    add_h("3. Szczegółowa analiza przypadków")

    cases = [
        {
            "title": "PRZYPADEK 01: Kalkulacja CO₂ oparta na jednym scenariuszu i na EN 16258:2012",
            "basis": "Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 i art. 6 · rozporządzenie (UE) 2026/1030 · EN ISO 14083:2023",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (liczbowe wprowadzenie w błąd)",
            "quotes": [
                "„Dzięki przeprowadzonym symulacjom i obliczeniom wg normy EN 16258:2012 (…) Ograniczenie emisji przy wykorzystaniu 33 palet LIBRA PARTNERS w skali miesiąca wynosi 400 kg CO₂.”",
                "„Miesięczny zakup 500 szt. → 6,060 tony mniej CO₂ (…) Miesięczny zakup 10 000 szt. → 121,120 ton mniej CO₂.”",
                "„dzięki naszym staraniom podczas produkcji palet (…) nasi klienci ograniczają emisję CO₂ do atmosfery o ponad 300 ton miesięcznie” oraz „303 t CO₂ mniej w skali miesiąca”.",
            ],
            "assessment": (
                "Rozporządzenie (UE) 2026/1030 stosuje się, gdy emisje GHG z usługi transportowej zaczynającej się lub kończącej w Unii "
                "są liczone i ujawniane dobrowolnie w celu komercyjnym. Metodyką odniesienia jest EN ISO 14083:2023, z pierwszeństwem danych pierwotnych. "
                "Materiał ujawnia liczbę dla odcinka Sieradz–Paryż i odsyła do normy EN 16258:2012. Ta norma nie jest metodyką CountEmissionsEU. "
                "W zapisie polskim „6,060 tony” oznacza 6,060 t. Iloczyn 500 × (400 kg / 33) daje około 6,06 t, a 10 000 sztuk daje rząd 121 t. "
                "Skala jest więc linijką z jednego scenariusza 33 palet, a nie pomiarem pracy przewozowej klienta. Sam zakup palety nie jest usługą transportową. "
                "Tekst o „303 t” mówi raz o produkcji, raz o niższej masie (9 kg wobec 25 kg). Granic systemu, jednostki funkcjonalnej, udziału pustych przebiegów, "
                "well-to-wheel i daty danych nie podano. To prawda cząstkowa podana zwodniczo (art. 5 ust. 2 pkt 2) i pominięcie, które zmienia znaczenie liczby (art. 6). "
                "Ten sam schemat — kalkulator bez pełnej drogi i bez skali floty — jest w zarzutach wobec InPost i DPD."
            ),
            "remediation": (
                "Usunąć ze strony głównej, ze „Zrównoważonego rozwoju” i z kart modeli przelicznik „zakup sztuk = tony CO₂” oraz liczbę 303 t. "
                "Do czasu nowego obliczenia nie publikować żadnej tony CO₂e. Nowe ujawnienie ograniczyć do wskazanego modelu, wskazanej relacji i wskazanego scenariusza, "
                "metodą EN ISO 14083:2023, z granicami, jednostką, źródłem danych i datą w tym samym miejscu co liczba."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Masa modelu MW01 wynosi 9 kg. Różnicę emisji GHG transportu wobec palety odniesienia publikujemy tylko jako wynik według EN ISO 14083:2023 "
                "(rozporządzenie 2026/1030), dla nazwanej relacji, z granicami systemu, jednostką funkcjonalną, informacją, czy użyto danych pierwotnych, i datą obliczenia. "
                "Wynik nie powstaje przez sam zakup palety i nie przelicza się liniowo na dowolny wolumen.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 02: Ekwiwalent 600 000 drzew",
            "basis": "Podstawa prawna: UPNPR art. 5 ust. 2 pkt 2 · praktyka UOKiK (InPost, sygn. komunikatu o zarzutach z 29.07.2025; Allegro)",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (manipulacja wskaźnikowa)",
            "quotes": [
                "„600.000 drzew: tyle drzew potrzebnych jest miesięcznie by wchłonąć dodatkową emisję CO₂ w przypadku stosowania palet drewnianych zamiast palet LIBRA PARTNERS.”",
                "„Przy założeniu, że jedno drzewo wchłania 0,5 kg CO₂ miesięcznie (…) każdego miesiąca potrzeba dodatkowo 600.000 drzew.”",
            ],
            "assessment": (
                "Iloczyn podany w materiale wynosi 600 000 × 0,5 kg = 300 t CO₂. Obok stoi osobna liczba „303 t” i zwrot „ponad 300 ton”. "
                "Założenie 0,5 kg na drzewo miesięcznie nie ma w tekście źródła, gatunku, wieku drzewa ani horyzontu. "
                "Ekwiwalent drzew nie jest metodyką EN ISO 14083. Prezes UOKiK zakwestionował u InPost okleiny typu „redukuje tyle CO₂ co 9 drzew” "
                "oraz u Allegro akcję „sadzimy drzewa” opartą na predykcji i warunku słabo widocznym dla odbiorcy (art. 5 i 6 UPNPR). "
                "Tutaj drzewa nie są nawet deklarowanym projektem nasadzeń. Są przelicznikiem, który powiększa odbiór różnicy emisyjnej. "
                "Punkt 4c załącznika ECGT zakazuje twierdzenia o neutralnym, ograniczonym lub pozytywnym wpływie GHG uzasadnianego kompensacją. "
                "Niniejsze zdanie nie mówi wprost „produkt neutralny dzięki drzewom”, lecz robi z pochłaniania przez drzewa miarę korzyści palety. "
                "W tej funkcji zdanie nie nadaje się do korekty redakcyjnej."
            ),
            "remediation": (
                "Usunąć każde zdanie o drzewach, hektarach lasu potrzebnych do pochłonięcia emisji i przeliczniku 0,5 kg CO₂ na drzewo. "
                "Nie zastępować go innym ekwiwalentem (samochody, loty, gospodarstwa domowe)."
            ),
            "claim_title": "SFORMUŁOWANIE NIE DO URATOWANIA",
            "claim": (
                "Ekwiwalent drzew nie ma dozwolonego odpowiednika. Po korekcie w tym miejscu nie ma claimu. "
                "Jedyna dopuszczalna informacja emisyjna to liczba z przypadku 01, po przeliczeniu według EN ISO 14083:2023."
            ),
            "salvage": False,
        },
        {
            "title": "PRZYPADEK 03: Reklama porównawcza dyskredytująca paletę drewnianą",
            "basis": "Podstawa prawna: UZNK art. 16 ust. 1 pkt 2 oraz ust. 3 pkt 2, 3 i 5 · UPNPR art. 5 ust. 2 pkt 3 · rozporządzenie (UE) 2023/1115 jako kontekst twierdzenia o wylesianiu",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (ryzyko roszczeń konkurencji)",
            "quotes": [
                "„Drewno jest bardzo trudnym materiałem do recyklingu oraz posiada wiele wad.”",
                "„przemysł drzewny wpływa negatywnie na środowisko naturalne, w latach 2010–2015 powierzchnia obszarów leśnych na świecie zmalała o 3,3 mln hektarów.”",
                "„Ponad 100.000 m³ drewna (…) Odpowiada to lasowi o powierzchni około 110 hektarów.”",
                "„Palety drewniane mają niekorzystny wpływ na politykę ESG.” oraz „Produkcja palet drewnianych wiąże się z wycinką drzew.”",
            ],
            "assessment": (
                "Tabela i tekst strony pozwalają rozpoznać towar konkurencyjny: standardową paletę drewnianą typu EUR. "
                "To reklama porównawcza. Jest dozwolona tylko łącznie: gdy nie wprowadza w błąd, porównuje obiektywnie cechy istotne, sprawdzalne i typowe "
                "oraz nie dyskredytuje towaru konkurenta (UZNK art. 16 ust. 3 pkt 2, 3 i 5). UPNPR art. 5 ust. 2 pkt 3 odsyła do tego samego przepisu. "
                "Zdania o „wielu wadach”, „trudnym recyklingu” i „niekorzystnym wpływie na politykę ESG” są oceną całej kategorii, bez kryterium i bez źródła. "
                "Ubytek 3,3 mln ha lasów na świecie w latach 2010–2015 nie jest w tekście powiązany z produkcją palet w Polsce. "
                "Przeliczenie 100 000 m³ na 110 ha nie ma metody. Rozporządzenie (UE) 2023/1115 wymaga należytej staranności wobec drewna wprowadzanego na rynek Unii; "
                "nie stanowi ustalenia, że paleta drewniana jako taka jest tożsama z globalnym wylesianiem. Zestawienie tych danych dyskredytuje surowiec konkurencji."
            ),
            "remediation": (
                "Wycofać tabelę w obecnej postaci oraz akapity o wadach drewna, globalnym ubytku lasów i „niekorzystnym ESG” palety drewnianej. "
                "Porównanie, jeśli zostanie, ograniczyć do cech, które spółka sama zmierzyła i potrafi pokazać: masa własnego modelu, wymiar, obciążenie według ISO 8611, "
                "surowiec tego modelu. Bez wniosku o wylesianiu i bez oceny polityki ESG konkurencyjnego surowca."
            ),
            "claim_title": "SFORMUŁOWANIE NIE DO URATOWANIA",
            "claim": (
                "Zdania o trudnym recyklingu drewna, o 3,3 mln ha, o 110 ha lasu i o niekorzystnym wpływie palety drewnianej na ESG nie mają bezpiecznego brzmienia. "
                "Usunąć je. Parametry własnego wyrobu podawać bez strony „paleta drewniana”."
            ),
            "salvage": False,
        },
        {
            "title": "PRZYPADEK 04: Ogólne hasła ekologiczne bez uznanej wysokiej efektywności",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a (dyrektywa 2005/29/WE) · UPNPR art. 5",
            "risk": "Poziom ryzyka: WYSOKIE (czarna lista)",
            "quotes": [
                "„LIBRA PARTNERS – Palety Nowej Generacji – Ekologiczne palety.”",
                "„Witaj w LIBRA PARTNERS: Technologia z myślą o Ekologii.”",
                "„Poznaj naszą ofertę wysokiej jakości eko palet” oraz „palety przemysłowe wykonane z ekologicznych surowców”.",
                "„Dodatkowo stosując palety plastikowe LIBRA PARTNERS (…) jesteście ESG Friendly!”",
                "„Misją firmy LIBRA PARTNERS jest stać się kluczowym graczem w zakresie zrównoważonego rozwoju.”",
            ],
            "assessment": (
                "Punkt 4a zakazuje ogólnego twierdzenia dotyczącego ekologiczności, gdy przedsiębiorca nie wykazuje uznanej wysokiej efektywności ekologicznej istotnej dla tego twierdzenia. "
                "Motyw 9 dyrektywy 2024/825 wymienia w tej grupie m.in. „ekologiczne” i sformułowania zbliżone, które sugerują wysoką efektywność. "
                "Motyw 10: taką efektywność wykazuje zgodność z rozporządzeniem (WE) nr 66/2010 (EU Ecolabel), urzędowo uznany program EN ISO 14024 typu I "
                "albo najwyższy poziom z innego prawa Unii, i musi ona dotyczyć całego twierdzenia. "
                "„Zrównoważony” i „ESG Friendly” obejmują też cechy inne niż środowiskowe; sam recyklat ich nie niesie. "
                "W materiale nie ma EU Ecolabel ani znaku typu I. Dopisek o HDPE z recyklingu, nawet na tej samej stronie, nie jest tą efektywnością. "
                "FAQ Komisji z 30 czerwca 2026 r.: specyfikacja zdejmuje status hasła ogólnego tylko wtedy, gdy jest jasna i dotyczy tego twierdzenia w tym samym środku przekazu. "
                "Hasło „ekologiczne palety” pozostaje hasłem ogólnym. Wzorce UOKiK: „eko Paczkomat”, listek „eko” Tchibo, „Join Life”."
            ),
            "remediation": (
                "Zdjąć z nagłówków, title, kart produktu, banerów targowych i cytatów słowa: ekologiczny, eko, ekologiczne surowce, ESG Friendly, "
                "technologia z myślą o ekologii, kluczowy gracz zrównoważonego rozwoju. Zostawić zdanie o surowcu i modelu, w tym samym miejscu co nazwa produktu."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Model MW01: paleta 1200 × 800 × 144 mm, masa 9 kg, surowiec — regranulat HDPE. Deklaracja dotyczy tego modelu.” "
                "Bez przymiotników „ekologiczna”, „eko” i „ESG Friendly”."
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 05: Zero Waste i brak jakichkolwiek odpadów jako cecha całej działalności",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4b · UPNPR art. 5 i art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (czarna lista)",
            "quotes": [
                "„W tym modelu działalności nie ma jakichkolwiek odpadów, występuje tu model Gospodarki o Obiegu Zamkniętym.”",
                "„Robimy to poprzez model biznesowy ukierunkowany na ZERO WASTE.”",
                "„Eliminacja odpadów pod koniec okresu ich użytkowania poprzez recykling naszych rozwiązań, które nigdy nie trafiają na wysypisko.”",
                "„Pooling palet (…) optymalizuje puste przebiegi transportowe oraz wpisuje się bezpośrednio w model Gospodarki o Obiegu Zamkniętym (GOZ).”",
            ],
            "assessment": (
                "Punkt 4b zakazuje twierdzenia o całym produkcie lub całej działalności, gdy dotyczy ono tylko aspektu albo niereprezentatywnego elementu. "
                "Motyw 11 podaje jako wzorzec wrażenie, że przedsiębiorca korzysta wyłącznie z określonego rozwiązania, podczas gdy dotyczy to części działalności. "
                "Jedyny mechanizm opisany w materiale to obietnica odbioru zużytych i uszkodzonych palet zakupionych od LIBRA PARTNERS oraz informacja o przerobie ponad 200 ton odpadów HDPE miesięcznie. "
                "Odbiór własnych palet nie jest ustaleniem, że zakład nie wytwarza odpadów, że każda paleta wraca, ani że żadna „nigdy” nie trafia na składowisko. "
                "„Zero Waste” jako hasło całej firmy jest tym samym układem co zarzut wobec Zary: „zero waste we własnych obiektach” przy skali innej niż komunikowana. "
                "Zdanie o poolingu dodaje eliminację pustych przebiegów bez udziału, bez relacji i bez danych o transporcie."
            ),
            "remediation": (
                "Usunąć „Zero Waste”, „nie ma jakichkolwiek odpadów” i „nigdy nie trafiają na wysypisko” oraz zdanie, że pooling optymalizuje puste przebiegi. "
                "Zostawić opis programu odbioru z jego rzeczywistym zakresem: palety kupione od spółki i zwrócone spółce."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Odbieramy zużyte palety zakupione od LIBRA PARTNERS i kierujemy tworzywo do ponownego przerobu. "
                "Program obejmuje palety zwrócone w jego ramach. Nie opisuje wszystkich odpadów zakładu ani palet, które do nas nie wrócą.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 06: Recyklat i masa 9 kg przypisane całej ofercie",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4b · UPNPR art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (czarna lista)",
            "quotes": [
                "„Palety wykonane z surowca HDPE z recyklingu.” „Waga: Nasze palety ważą tylko 9 kg.”",
                "„Zmień palety drewniane na ekologiczne palety LIBRA PARTNERS z surowca HDPE z recyklingu!”",
                "Model RS-01-FT: „Waga: 12 kg” oraz „Surowiec: HDPE pierwotny / dopuszczony do kontaktu”.",
                "Model MW01-PP: „Waga: 11 kg” oraz „Surowiec: Compound PP z recyklingu”. Przy ESD: „Surowiec: HDPE ESD” — udział recyklatu niepodany.",
            ],
            "assessment": (
                "Strona główna i wezwanie do zmiany palet drewnianych przedstawiają ofertę jako palety z HDPE z recyklingu o masie 9 kg. "
                "Specyfikacja RS-01-FT, czyli linii wskazywanej do żywności i farmacji, podaje HDPE pierwotny i 12 kg. "
                "Linia automotive jest z recyklatu PP i waży 11 kg. Przy ESD materiał nie podaje udziału recyklatu. "
                "Kalkulacja CO₂ w tym samym zbiorze przyjmuje różnicę 9 kg wobec 25 kg. Odbiorca, który czyta nagłówek i liczbę emisji, "
                "nie dostaje informacji, że część modeli ma inną masę i inny surowiec. To pominięcie z art. 6 UPNPR i układ z pkt 4b: "
                "twierdzenie o całej ofercie, dowód co najwyżej o części modeli. Wzorzec Tchibo w bazie: wyrób z udziałem recyklatu komunikowany jak „eko” całości."
            ),
            "remediation": (
                "Przy każdym modelu, w tym samym bloku co nazwa, podać surowiec i masę z specyfikacji. "
                "Zdjąć zbiorcze „nasze palety ważą tylko 9 kg” i „surowiec HDPE z recyklingu” z miejsc, które obejmują RS-01-FT, ESD i automotive. "
                "Kalkulacji z przypadku 01 nie stosować do modeli o innej masie."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„MW01 i MW01-R: regranulat HDPE, 9 kg. MW01-MED: regranulat HDPE z dodatkiem tlenku cynku, 9 kg. "
                "RS-01-FT: HDPE pierwotny, 12 kg. MW01-PP: mieszanka PP z recyklingu, 11 kg. ESD: HDPE o właściwościach antystatycznych, 9 kg; "
                "udział recyklatu podamy po jego ustaleniu dla tej linii.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 07: Twierdzenie „10 razy bardziej trwałe”",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 23g · UZNK art. 16 ust. 3 pkt 2 i 3 · UPNPR art. 5 ust. 3 pkt 2",
            "risk": "Poziom ryzyka: WYSOKIE (czarna lista)",
            "quotes": [
                "„Trwałość: Nasze palety są 10 razy bardziej trwałe w porównaniu do drewnianych.”",
                "„można wykorzystywać je 10 razy dłużej (…) Według badań przeprowadzonych przez Instytut Materiałoznawstwa w Warszawie średni cykl życia palety drewnianej wynosi 6-7 cykli (…) cykl życia palety plastikowej wynosi około 60-70 cykli.”",
                "Tabela porównawcza: paleta drewniana „7-8 cykli załadunku / rozładunku (według badań Instytutu Materiałoznawstwa w Warszawie 6-7 cykli)”; paleta LIBRA „10 razy trwalsza od palety drewnianej (około 60-70 cykli)”.",
            ],
            "assessment": (
                "Punkt 23g zakazuje fałszywego twierdzenia, że w normalnych warunkach towar ma określoną trwałość co do czasu lub intensywności użytkowania. "
                "Materiał podaje krotność „10 razy” oraz widełki 60–70 cykli wobec palety drewnianej. "
                "W tym samym zbiorze paleta drewniana ma raz 6–7, raz 7–8 cykli. Badanie Instytutu Materiałoznawstwa w Warszawie nie ma tytułu, daty, próby ani warunków. "
                "Nie da się sprawdzić, czy cykle są porównywalne (załadunek, naprawa, klasa palety drewnianej, środowisko). "
                "Reklama porównawcza wymaga cech sprawdzalnych na podstawie obiektywnych kryteriów (UZNK art. 16 ust. 3 pkt 2 i 3). "
                "Zdanie „nie występuje proces degradacji” wskutek warunków atmosferycznych jest dalej idące niż liczba cykli i również nie ma metody w tekście. "
                "Do czasu badania, które da się odczytać, krotność zostaje na czarnej liście jako twierdzenie o trwałości bez wykazanej podstawy, "
                "a wewnętrzna sprzeczność 6–7 / 7–8 cykli wyklucza obecne brzmienie."
            ),
            "remediation": (
                "Usunąć „10 razy”, „10 razy dłużej”, „60–70 cykli” i zdanie o braku degradacji pod wpływem warunków atmosferycznych. "
                "Przywrócić liczbę dopiero z nazwą badania, datą, definicją cyklu i paletą odniesienia, w tym samym kanale."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Trwałość modelu [nazwa] podajemy jako liczbę cykli [definicja cyklu] według badania [jednostka, data, numer]. "
                "Do publikacji tego badania nie porównujemy trwałości z paletą drewnianą.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 08: Brak fumigacji ISPM 15 jako zaleta ekologiczna",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 10a · UPNPR art. 5",
            "risk": "Poziom ryzyka: WYSOKIE (czarna lista)",
            "quotes": [
                "„Higiena: Idealne dla branży spożywczej i farmaceutycznej. Nie wymagają fumigacji – zgodne z normą ISPM 15.”",
                "„Z uwagi na eksport naszych produktów do UK, musieliśmy kupować drewniane palety fumigowane. (…) nie wymagają raportowania wg normy ISPM15. (…) jesteśmy bardziej ekologiczni.”",
                "„Z uwagi na eksport morski do Ameryki Północnej i Południowej, wybór palet LIBRA PARTNERS będzie zarówno ekonomiczny, jak i ekologiczny!”",
                "„Już na etapie wytwarzania palet, są one bardziej ekologiczne, bowiem nie wymagają poddawania specjalistycznej obróbce i suszeniu.”",
            ],
            "assessment": (
                "Punkt 10a zakazuje przedstawiania wymogu nałożonego prawem na wszystkie produkty danej kategorii na rynku unijnym jako cechy wyróżniającej ofertę. "
                "ISPM 15 dotyczy drewnianego materiału opakowaniowego w obrocie międzynarodowym. Paleta z tworzywa nie wchodzi w zakres tej obróbki. "
                "Brak fumigacji jest skutkiem surowca, nie dobrowolnym wynikiem środowiskowym i nie „zgodnością”, która wyróżnia tę spółkę spośród palet z tworzywa. "
                "Materiał łączy ten fakt ze zdaniem „jesteśmy bardziej ekologiczni” oraz z wyższością już na etapie wytwarzania. "
                "Tak zbudowana cecha jest wyróżnikiem marketingowym z wyłączenia, które dotyczy całej kategorii palet niedrewnianych. "
                "Te same opinie wiążą ten wyróżnik z eksportem do Wielkiej Brytanii oraz z eksportem morskim do Ameryki Północnej i Południowej. "
                "To, czy drewniane opakowanie wymaga obróbki i jak się ją dokumentuje, zależy od kraju, do którego towar jest eksportowany. "
                "Zdanie o braku fumigacji nie zastępuje sprawdzenia wymogów kraju przeznaczenia. "
                "Ocena greenwashingu w tym raporcie pozostaje oceną przekazu opublikowanego w Polsce; prawa krajów eksportu o twierdzeniach środowiskowych nie ma w bazie wiedzy narzędzia."
            ),
            "remediation": (
                "W kartach i opiniach rozdzielić informację fitosanitarną od jakiegokolwiek zdania o ekologii, ESG i niższej emisji. "
                "Przy eksporcie wskazać kraj przeznaczenia i odesłać do wymogów tego kraju, zamiast jednego zdania o „zgodności z ISPM 15” dla wszystkich kierunków."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Paleta z tworzywa nie jest drewnianym materiałem opakowaniowym objętym ISPM 15 i nie przechodzi fumigacji przewidzianej dla drewna. "
                "Wymogi fitosanitarne przesyłki zależą od kraju, do którego towar jest eksportowany. Brak fumigacji nie jest cechą środowiskową.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 09: Zgodność z ESG, Agendą 2030 i raportowaniem niefinansowym",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5 ust. 2 pkt 2",
            "risk": "Poziom ryzyka: WYSOKIE (czarna lista)",
            "quotes": [
                "„Spraw by Twoja firma była zgodna z polityką Zrównoważonego Rozwoju i ESG. (…) Zmiana palet z drewnianych na plastikowe LIBRA PARTNERS to najprostsze rozwiązanie na znaczne ograniczenie emisji CO₂.”",
                "„Wykorzystanie palet plastikowych w przedsiębiorstwie i w łańcuchu dostaw ogranicza emisję CO₂ oraz jest zgodne z polityką UE oraz ESG.”",
                "„Wybrane cele polityki zrównoważonego rozwoju, potwierdzające dopasowanie naszego projektu do celów zawartych w Agendzie 2030” — w tym cel 15.2: „Do 2020 roku promować wdrażanie zrównoważonego zarządzania wszystkimi typami lasów; zahamować proces wylesiania”.",
            ],
            "assessment": (
                "„Zrównoważony” bez specyfikacji w tym samym kanale jest hasłem ogólnym z pkt 4a. Motyw 10 zabrania opierać takie hasło wyłącznie na jednej cesze środowiskowej, "
                "bo obejmuje ono także cechy inne niż środowiskowe. Zdanie, że zmiana palet czyni firmę klienta zgodną z ESG i jest najprostszym rozwiązaniem ograniczenia emisji, "
                "przypisuje jednej decyzji zakupowej skutek dla całej działalności odbiorcy. To pkt 4b. "
                "Lista celów ONZ, w tym 15.2 z terminem 2020, jest przedstawiona jako potwierdzenie dopasowania projektu. Materiał nie podaje wskaźnika, roku bazowego ani zakresu, w jakim cel miałby być realizowany. "
                "Ten układ odpowiada zarzutom wobec Zary z 19 stycznia 2026 r.: SDG, zero waste i hasła ogólne bez skali, ocenione na gruncie art. 5 UPNPR. "
                "Klient, który wpisze „121,120 ton” albo „zgodne z ESG” do własnego raportu niefinansowego, dostaje liczbę i kwalifikację bez metodyki z przypadku 01."
            ),
            "remediation": (
                "Usunąć zaproszenie do „zgodności z ESG”, „polityką UE oraz ESG” i „najprostszego rozwiązania” oraz blok „potwierdzające dopasowanie” do celów 9.4, 12.2, 12.5 i 15.2. "
                "Nie obiecywać skutku w raporcie niefinansowym klienta."
            ),
            "claim_title": "SFORMUŁOWANIE NIE DO URATOWANIA",
            "claim": (
                "Zdanie o zgodności klienta z ESG, z polityką UE i z Agendą 2030 przez zakup palet nie ma bezpiecznej wersji. "
                "Po korekcie spółka nie składa claimu o raporcie niefinansowym odbiorcy. Odbiorca sam decyduje, czy i jak ujmie zweryfikowany wskaźnik we własnym sprawozdaniu."
            ),
            "salvage": False,
        },
        {
            "title": "PRZYPADEK 10: Powołanie na analizę LCA Envirly bez elementów badania",
            "basis": "Podstawa prawna: UPNPR art. 6 · normy ISO 14040 i ISO 14044 jako rama ujawnienia powołanego badania",
            "risk": "Poziom ryzyka: ŚREDNIE (brak weryfikowalności)",
            "quotes": [
                "„opracowany przez firmę Envirly raport LCA dotyczący naszych palet potwierdza prawie dwukrotnie niższą emisję CO₂ w łańcuchu dostaw w porównaniu do standardowych palet drewnianych.”",
                "„W sierpniu 2023 roku powstała analiza Libra Partners przez firmę Envirly (…) Zachęcamy do zapoznania się z przedstawionymi w niej wynikami.”",
            ],
            "assessment": (
                "Materiał powołuje badanie strony trzeciej z sierpnia 2023 r. i wyciąga z niego wniosek „prawie dwukrotnie niższa emisja w łańcuchu dostaw”. "
                "Nie podaje modelu palety, jednostki funkcjonalnej, granic systemu, wskaźnika (CO₂ czy CO₂e), scenariusza palety drewnianej ani miejsca, w którym wynik można przeczytać. "
                "ISO 14040 i ISO 14044 wymagają celu i zakresu oraz granic, zanim wynik nadaje się do porównania. "
                "W tym samym zbiorze jedyne liczby CO₂ są symulacją trasy według EN 16258 i ekstrapolacją zakupu. Odbiorca może wziąć „prawie dwukrotnie” za potwierdzenie kalkulatora. "
                "Tego powiązania tekst nie wykazuje. Brak tych informacji jest pominięciem istotnym (art. 6 UPNPR). Samo istnienie analizy z 2023 r. nie jest zakwestionowane."
            ),
            "remediation": (
                "Do czasu publikacji streszczenia zdjąć „prawie dwukrotnie niższą emisję”. "
                "Streszczenie umieścić przy zdaniu, które na nie powołuje: model, data, jednostka funkcjonalna, moduły cyklu życia, wynik, podmiot i data przeglądu, jeżeli był."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„W sierpniu 2023 r. Envirly opracowało analizę LCA modelu [nazwa]. Streszczenie (cel, granice, jednostka funkcjonalna, wynik w kg CO₂e na jednostkę, data) "
                "jest dostępne [adres]. Wynik dotyczy tego modelu i tych granic. Nie potwierdza kalkulatora zakupowego ani liczby drzew.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 11: Transfer atestu, ISO 9001, opinii uczelni i wyniku ISO 22196 na claim środowiskowy",
            "basis": "Podstawa prawna: UPNPR art. 5 ust. 3 pkt 2 i art. 6",
            "risk": "Poziom ryzyka: ŚREDNIE (transfer atrybutów)",
            "quotes": [
                "„Wytwarzane z surowca HDPE z recyklingu, według ścisłej kontroli systemu zarządzania ISO 9001 w znacznym stopniu do ograniczenia emisji śladu węglowego.”",
                "„Mają atesty od PZH (…) Co więcej, na naszej stronie internetowej sprawdzisz, w jaki sposób dbamy o zrównoważony rozwój naszej firmy i niską emisyjność CO₂.”",
                "„uzyskała opinię o innowacyjności na skalę Świata wystawioną przez Politechnikę Poznańską. (…) jesteśmy pierwszym producentem na świecie palet (…) które są w 99% antybakteryjne przez cały cykl życia palety. Dodatkowo (…) jesteście ESG Friendly!”",
                "„przebadane wg normy ISO 22196:2011: zmniejszenie się liczebności próbek bakterii Escherichia coli (ATCC 8739) o 99,97%” oraz „Staphylococcus aureus (ATCC 6538P) o 99,98%”.",
            ],
            "assessment": (
                "Artykuł 5 ust. 3 pkt 2 UPNPR obejmuje wprowadzenie w błąd co do testów, zezwoleń, nagród i korzyści produktu. "
                "ISO 9001 jest systemem zarządzania jakością. Atest Państwowego Zakładu Higieny dotyczy kontaktu i higieny w zakresie, w jakim został wydany; materiał tego zakresu nie przytacza. "
                "Opinia Politechniki Poznańskiej o innowacyjności nie jest w tekście opisana co do przedmiotu i nie jest dowodem emisji. "
                "Zestawienie tych dokumentów w jednym ciągu ze „ślad węglowy”, „niską emisyjnością CO₂” i „ESG Friendly” przenosi ich znaczenie na profil środowiskowy. "
                "Osobno wynik ISO 22196:2011 dotyczy dwóch wskazanych szczepów w warunkach tej normy. Zdanie „w 99% antybakteryjne przez cały cykl życia” i „pierwszy producent na świecie” "
                "rozszerza ten wynik na cały okres użytkowania i na pozycję rynkową, których badanie w materiale nie opisuje."
            ),
            "remediation": (
                "Rozdzielić bloki. Parametr techniczny i dokument zostają przy swoim przedmiocie. "
                "Zdjąć „w znacznym stopniu do ograniczenia emisji” przy ISO 9001, ciąg PZH–emisyjność oraz „pierwszy na świecie” i „99% przez cały cykl życia”, dopóki nie ma badania o tym zakresie."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Model MW01-MED: w badaniu według ISO 22196:2011 liczebność E. coli ATCC 8739 spadła o 99,97%, a S. aureus ATCC 6538P o 99,98%, w warunkach tego badania. "
                "Wynik nie opisuje całego cyklu życia, innych drobnoustrojów ani emisji GHG. Opinia Politechniki Poznańskiej i atest PZH dotyczą zakresu wskazanego w tych dokumentach.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 12: Dodatek tlenku cynku przy twierdzeniu o pełnym przerobie",
            "basis": "Podstawa prawna: UPNPR art. 6 · ECGT, załącznik I pkt 4b",
            "risk": "Poziom ryzyka: ŚREDNIE (pominięcie istotne)",
            "quotes": [
                "„Dzięki dodaniu podczas procesu produkcyjnego tlenku cynku w odpowiedniej proporcji do surowca HDPE nasza paleta MW01-MED wykazuje bardzo duże właściwości antybakteryjne w całym cyklu życia palety.”",
                "„Surowiec: HDPE z recyklingu + dodatek tlenku cynku (ZnO).”",
                "„surowiec, z którego produkowane są palety jest w 100% możliwy do ponownego przerobienia” oraz w tabeli: „100% mogą być poddane recyklingowi.”",
            ],
            "assessment": (
                "Linia MW01-MED jest w specyfikacji regranulatem HDPE z dodatkiem tlenku cynku, wprowadzanym dla działania przeciwdrobnoustrojowego. "
                "Te same materiały mówią o surowcu palet w 100% zdatnym do ponownego przerobu i o modelu, w którym nie ma jakichkolwiek odpadów. "
                "Tekst nie podaje udziału dodatku, jego statusu ani tego, czy paleta z ZnO wchodzi do tego samego strumienia odbioru i przemiału co MW01 bez dodatku. "
                "Odbiorca oferty zbiorczej nie dostaje informacji, która może zmienzyć znaczenie hasła „100%”. "
                "To pominięcie z art. 6. O ile hasło obejmuje także MW01-MED, jest to również pkt 4b: twierdzenie o całym surowcu przy linii z dodatkiem funkcjonalnym. "
                "Audyt nie przypisuje tlenkowi cynku klasyfikacji zagrożenia, której nie ma w badanym materiale ani w bazie wiedzy narzędzia."
            ),
            "remediation": (
                "Przy MW01-MED, obok informacji o ZnO, napisać, czy ten model jest w programie odbioru i czy dodatek zmienia sposób przerobu. "
                "Do tego zdania nie używać zbiorczego „100%” dla wszystkich modeli."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„MW01-MED zawiera regranulat HDPE i dodatek tlenku cynku. Odbiór zużytej palety tego modelu i sposób jej przerobu opisujemy osobno, "
                "po ustaleniu wpływu dodatku na strumień tworzywa. Twierdzenie o przerobie modelu MW01 bez dodatku nie obejmuje MW01-MED.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 13: Anonimowe opinie o byciu „bardziej EKO”",
            "basis": "Podstawa prawna: UPNPR art. 7 pkt 25 i 26",
            "risk": "Poziom ryzyka: NISKIE",
            "quotes": [
                "„Dzięki lżejszej wadze, oraz temu, że pochodzą z recyklingu – będziemy bardziej EKO” (Kierownik działu ESG, branża spożywcza).",
                "„Dla nas to oszczędność czasu i jesteśmy bardziej ekologiczni.” (Kierownik eksportu, branża spożywcza).",
                "„wybór palet LIBRA PARTNERS będzie zarówno ekonomiczny, jak i ekologiczny!” (Dyrektor sprzedaży, branża wyrobów medycznych).",
            ],
            "assessment": (
                "Artykuł 7 pkt 25 UPNPR zakazuje twierdzenia, że opinie pochodzą od konsumentów, którzy używali produktu lub go nabyli, "
                "jeżeli przedsiębiorca nie podjął uzasadnionych i proporcjonalnych kroków, aby to sprawdzić. Punkt 26 zakazuje zamieszczania nieprawdziwych opinii "
                "oraz zniekształcania opinii w celu promowania produktów. Przepisy te są polskim odpowiednikiem reguł opinii z dyrektywy Omnibus; "
                "w tekście jednolitym UPNPR nie występują jako pkt 23b i 23c załącznika UCPD. "
                "Cytaty są podpisane funkcją i branżą, bez imienia, firmy i daty. Materiał nie opisuje kroku weryfikacji. "
                "Treść opinii powtarza hasła „bardziej EKO” i „ekologiczny”, czyli claim z przypadku 04, oraz wniosek z ISPM 15 z przypadku 08. "
                "Sam opis niższej masy i spalania, gdyby był imienny i sprawdzony, nie jest tym naruszeniem."
            ),
            "remediation": (
                "Zdjąć cytaty z wnioskiem „bardziej EKO”, „bardziej ekologiczni” i „ekologiczny”. "
                "Opinię o masie i spalaniu publikować z oznaczeniem podmiotu oraz informacją, że pochodzi od klienta, który palet używał, po faktycznym sprawdzeniu."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "Nie publikować opinii o „byciu eko”. Dopuszczalny jest cytat imienny, po weryfikacji, ograniczony do faktu: "
                "„[nazwa spółki, funkcja, data]: paleta MW01 waży 9 kg; przy naszych przewozach niższa masa nośnika obniża masę zestawu.” Bez wniosku środowiskowego."
            ),
            "salvage": True,
        },
    ]

    for case in (cases[:1] if preview else cases):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(2)
        h.paragraph_format.keep_with_next = True
        set_run(h.add_run(case["title"]), 12, True, False, PRIMARY)
        meta_line = doc.add_paragraph()
        meta_line.paragraph_format.space_before = Pt(0)
        meta_line.paragraph_format.space_after = Pt(4)
        meta_line.paragraph_format.keep_with_next = True
        set_run(meta_line.add_run(case["basis"] + "  |  " + case["risk"]), 9, True, False, NAVY)
        add_rich([("Cytat źródłowy: ", True, False, DARK)], size=10.5, before=2, after=2)
        for q in case["quotes"]:
            add_p(q, size=10.5, italic=True, after=3)
        add_rich([("Ocena prawna: ", True, False, DARK)], size=10.5, before=4, after=2)
        add_p(case["assessment"], size=10.5, after=6)
        add_rich([("Kierunek naprawy: ", True, False, DARK)], size=10.5, before=2, after=2)
        add_p(case["remediation"], size=10.5, after=4)
        callout(
            case["claim_title"],
            [case["claim"]],
            "E8F4FA" if case["salvage"] else "F4F7F8",
            "2A8AB0" if case["salvage"] else "8AA0AB",
            ACCENT if case["salvage"] else MUTED,
        )

    if preview:
        add_p(
            "Pozostałe przypadki nie są w tej wersji omawiane. Ich poziom ryzyka jest w tabeli zbiorczej; "
            "twierdzenie, podstawa prawna i istota naruszenia są rozmyte.",
            size=10, italic=True, color=MUTED, before=2, after=6,
        )

    add_h("4. Porównanie twierdzeń LIBRA PARTNERS z rzeczywistym stanem prawno-środowiskowym")
    add_p(
        "Sekcja jest w raporcie, ponieważ badany zbiór zawiera tabelę „palety LIBRA PARTNERS wobec palet drewnianych” "
        "oraz zdania porównawcze w tekście strony. Poniżej tylko kryteria środowiskowe i te, które niosą wniosek środowiskowy. "
        "Cechy czysto techniczne (RFID, gwoździe, BHP) nie są przedmiotem tego zestawienia.",
        after=6,
    )

    comparison = [
        ("Surowiec i masa całej oferty",
         "Paleta LIBRA: „tworzywo HDPE z recyklingu”, 9 kg. Paleta drewniana: drewno, 25 kg. Z masy wyprowadzono niższe CO₂.",
         "Specyfikacje: MW01 — rHDPE, 9 kg; RS-01-FT — HDPE pierwotny, 12 kg; automotive — rPP, 11 kg; ESD — HDPE ESD, recyklat niepodany. Zbiorcze 9 kg i rHDPE nie obejmują oferty. Różnica masy nie jest liczbą GHG."),
        ("Trwałość",
         "Drewno: 7–8 cykli, a w innym akapicie 6–7 według Instytutu Materiałoznawstwa w Warszawie. LIBRA: 10 razy trwalsza, około 60–70 cykli, brak degradacji od warunków atmosferycznych.",
         "Jedno badanie nie jest zidentyfikowane (tytuł, data, warunki). Dwie liczby cykli drewna są w tym samym materiale. Krotność „10 razy” nie spełnia pkt 23g ani art. 16 ust. 3 pkt 2 i 3 UZNK."),
        ("Recykling i koniec życia",
         "Drewno: koszt recyklingu czasem po stronie właściciela, brak wartości rezydualnej. LIBRA: „100%” recyklingu, gwarantowany odkup, brak jakichkolwiek odpadów.",
         "Opisany mechanizm to odbiór palet kupionych od spółki. MW01-MED ma dodatek ZnO, którego wpływu na przerób nie podano. „100%” i „brak odpadów” są claimem o całości (pkt 4b)."),
        ("Wylesianie",
         "Produkcja palet drewnianych „wiąże się z wycinką drzew”. Obok: spadek powierzchni lasów na świecie o 3,3 mln ha w latach 2010–2015 oraz 100 000 m³ = około 110 ha miesięcznie w Polsce.",
         "Globalny ubytek z lat 2010–2015 nie jest w tekście przypisany do palet. Przeliczenie m³ na hektary nie ma metody. EUDR (rozp. 2023/1115) to należyta staranność wobec drewna, nie ustalenie, że paleta drewniana równa się temu ubytkowi."),
        ("ESG i ślad węglowy",
         "Palety drewniane „mają niekorzystny wpływ na politykę ESG”. Palety plastikowe „ograniczają emisję CO₂ oraz są zgodne z polityką UE oraz ESG”.",
         "Brak kryterium „polityki ESG”. Liczby CO₂ w materiale pochodzą z EN 16258:2012 i z liniowej skali zakupu, a nie z EN ISO 14083. Hasło zgodności z ESG jest ogólne (pkt 4a i 4b)."),
        ("Higiena biologiczna",
         "Drewno: „niebezpieczeństwo zagnieżdżenia się pasożytów, grzybów i infekcji”. Tworzywo: „brak możliwości penetracji przez grzyby i pasożyty”.",
         "To porównanie kategorii, bez badania przytoczonego w materiale. Bezwzględny „brak możliwości” i przypisanie zagrożenia całemu drewnu nie są cechą sprawdzalną w rozumieniu art. 16 ust. 3 pkt 3 UZNK. Cechę mycia podawać jako cechę własnego modelu, bez wniosku o zagrożeniu biologicznym konkurenta."),
    ]
    comp = doc.add_table(rows=1 + len(comparison), cols=3)
    comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp.autofit = False
    cw = (Cm(3.3), Cm(7.1), Cm(7.2))
    for c, h in enumerate(["Kryterium porównawcze", "Twierdzenie podmiotu", "Rzeczywisty stan prawno-środowiskowy"]):
        cell = comp.rows[0].cells[c]
        cell.width = cw[c]
        shade(cell, "08181F")
        margins(cell, 50, 50, 50, 50)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
        set_run(p.add_run(h), 8, True, False, RGBColor(255, 255, 255))
    for r, (a, b, ctext) in enumerate(comparison):
        bg = "FFFFFF" if r % 2 == 0 else "F8FAFC"
        for c, val in enumerate((a, b, ctext)):
            cell = comp.rows[r + 1].cells[c]
            cell.width = cw[c]
            shade(cell, bg)
            margins(cell, 40, 40, 50, 50)
            borders(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
            if preview and r > 0:
                inner = max(0.9, cw[c].cm - 0.32)
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
                p.add_run().add_picture(
                    blurred_png(val, inner, size_pt=7.5, bold=(c == 0), bg=bg),
                    width=Cm(inner),
                )
            else:
                set_run(p.add_run(val), 7.5, c == 0, False, DARK)
    if preview:
        lock_table(comp, cw)

    add_p("", size=4, after=2)
    add_h("5. Ryzyka prawne i finansowe")

    risks = [
        ("• 1. Prezes UOKiK — do 10% obrotu",
         "Artykuł 24 ustawy o ochronie konkurencji i konsumentów zakazuje praktyk naruszających zbiorowe interesy konsumentów, w tym nieuczciwych praktyk rynkowych. "
         "Artykuł 106 ust. 1 pkt 4 pozwala nałożyć karę do 10% obrotu z roku poprzedzającego nałożenie kary, także przy naruszeniu nieumyślnym. "
         "W komunikatach o zarzutach z 29 lipca 2025 r. (InPost, Allegro, DHL, DPD) i z 19 stycznia 2026 r. (Bolt, Tchibo, Zara) Prezes wskazuje tę karę za każdą praktykę. "
         "Wzorce z tych spraw — ogólne „eko”, kalkulator CO₂, ekwiwalent drzew, zero waste i SDG bez skali — są w badanym materiale. "
         "Od 27 września 2026 r. kanał online nie korzysta z taryfy starych zapasów (wspólne rozumienie CPC). Korekta strony jest natychmiastowa."),
        ("• 2. Roszczenia konkurencji — UZNK art. 16 i 18",
         "Reklama porównawcza palety drewnianej, która wprowadza w błąd albo dyskredytuje ten towar, jest czynem nieuczciwej konkurencji (art. 16 ust. 3). "
         "Przedsiębiorca, którego interes został zagrożony lub naruszony, może żądać zaniechania, usunięcia skutków, oświadczenia, naprawienia szkody, "
         "wydania bezpodstawnie uzyskanych korzyści, a przy czynie zawinionym — odpowiedniej sumy na cel społeczny związany ze wspieraniem kultury polskiej "
         "lub ochroną dziedzictwa narodowego (art. 18 ust. 1). Agencja, która reklamę opracowała, odpowiada obok reklamodawcy (art. 17)."),
        ("• 3. Raportowanie niefinansowe klientów",
         "Materiał wprost zachęca odbiorcę, aby zmianą palet zapewnił sobie zgodność z ESG i ograniczenie emisji na potrzeby raportowania niefinansowego, "
         "oraz podaje gotowe tony CO₂ przy zakupie 500–10 000 sztuk. Klient, który przeniesie te liczby do własnego sprawozdania, przejmuje wskaźnik bez granic systemu "
         "i bez EN ISO 14083. Ryzyko leży po obu stronach: zarzut wprowadzenia w błąd wobec spółki oraz zakwestionowanie danych u klienta, który raportuje według CSRD/ESRS, o ile ten obowiązek go dotyczy."),
        ("• 4. Ujawnianie emisji transportu i trwałości po 27 września 2026 r.",
         "Każda dalsza publikacja tony CO₂ dla relacji transportowej w celu komercyjnym wpada w rozporządzenie 2026/1030. "
         "Krotność trwałości bez badania wpada w pkt 23g. Wytyczne UOKiK o oświadczeniach środowiskowych w stanie bazy na wrzesień 2026 r. pozostają nieopublikowane; "
         "nie zastępują one czarnej listy ani art. 5 i 6 UPNPR."),
    ]
    for index, (title, body) in enumerate(risks):
        if preview and index > 0:
            block = doc.add_paragraph()
            block.paragraph_format.space_before = Pt(8)
            block.paragraph_format.space_after = Pt(4)
            block.paragraph_format.line_spacing = 1.0
            block.add_run().add_picture(
                blurred_png(title + "\n" + body, 17.2, size_pt=10.5),
                width=Cm(17.2),
            )
        else:
            add_p(title, size=11, bold=True, color=PRIMARY, before=8, after=2)
            add_p(body, size=10.5, after=4)

    add_h("6. Rekomendacje naprawcze")
    steps = [
        ("Etap 1: Natychmiastowe zdjęcie treści — 7 października 2026 r.",
         "Z strony głównej, „Zrównoważonego rozwoju”, „O nas”, kart modeli, tabeli porównawczej i cytatów zdjąć: „ekologiczne / eko / ESG Friendly”, "
         "Zero Waste i „brak jakichkolwiek odpadów”, kalkulator sztuk na tony CO₂, 303 t, 600 000 drzew, 3,3 mln ha, „niekorzystny wpływ palet drewnianych na ESG”, "
         "„10 razy trwalsze” oraz opinie z wnioskiem „bardziej EKO”. Kanał online korygować od razu, bez okresu przejściowego."),
        ("Etap 2: Korekta kart i oferty — do 14 października 2026 r.",
         "Przy każdym modelu podać surowiec i masę ze specyfikacji, w tym HDPE pierwotny i 12 kg dla RS-01-FT oraz ZnO dla MW01-MED. "
         "ISPM 15 zostawić wyłącznie jako informację, że paleta z tworzywa nie jest drewnianym materiałem opakowaniowym. "
         "Rozdzielić atest PZH, ISO 9001, opinię Politechniki i wynik ISO 22196 od zdań o emisjach i ESG."),
        ("Etap 3: Metodyka i dowód — do 6 listopada 2026 r.",
         "Jeżeli spółka nadal chce podać liczbę GHG transportu, przeliczyć ją według EN ISO 14083:2023, osobno dla modeli o masie 9, 11 i 12 kg, "
         "i opublikować granice, jednostkę, dane i datę przy liczbie. Opublikować streszczenie analizy Envirly z sierpnia 2023 r. albo nie powoływać się na „prawie dwukrotnie”. "
         "Zidentyfikować badanie cykli albo nie podawać krotności trwałości."),
        ("Etap 4: Stały język dozwolony i przegląd treści",
         "Dalsze teksty, targi i oferty B2B prowadzi się z zamkniętej listy zdań z tego raportu (wzory po korekcie). "
         "Nowe hasło środowiskowe przed publikacją sprawdza się listą: czy jest ogólne (4a), czy obejmuje całą ofertę (4b), czy jest cechą z przepisu (10a), "
         "czy podaje tony, drzewa albo krotność trwałości. Blokady nie zdejmuje dopisek na innej podstronie ani kod QR."),
    ]
    for index, (title, body) in enumerate(steps):
        callout(title, [body], "F8FAFC", "1D4ED8", NAVY, blur=preview and index > 0)

    add_h("7. Wnioski")
    add_p(
        "Zbiór materiałów LIBRA PARTNERS ze stanu na 26 września 2026 r. nie nadaje się do dalszej publikacji. "
        "Zawiera naruszenia czarnej listy ECGT (pkt 4a, 4b, 10a i 23g), liczbową kalkulację CO₂ poza EN ISO 14083 oraz reklamę porównawczą palety drewnianej sprzeczną z art. 16 ust. 3 UZNK."
    )
    add_p(
        "Po zdjęciu tych twierdzeń spółka może komunikować surowiec i masę konkretnego modelu, program odbioru palet zwróconych spółce, "
        "wynik ISO 22196:2011 w jego granicach oraz — dopiero po nowym obliczeniu — emisję transportu według rozporządzenia 2026/1030."
    )

    if out_path is None:
        out_path = ROOT / "Raport_Analizy_Greenwashingu_Libra_Partners.docx"
    out_path = Path(out_path)
    doc.save(out_path)
    print(out_path)
    return out_path


if __name__ == "__main__":
    generate_report(ROOT / "wzory" / "WZOR_raport_Libra_Partners_pelny.docx", preview=False)
    generate_report(ROOT / "wzory" / "WZOR_raport_Libra_Partners_okrojony.docx", preview=True)
