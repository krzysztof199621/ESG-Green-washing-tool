#!/usr/bin/env python3
"""Raport audytu greenwashingu Hi DESIGN — forma wg INSTRUKCJA_GENEROWANIA_RAPORTOW.md."""

from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = ROOT.parent
SHOT = ROOT / "zrzut_strona_glowna.png"
LOGO_CANDIDATES = [
    ROOT / "esg_logo_white.png",
    PROJECT_ROOT.parent / "ESG Hero Day - Materialy Druk" / "assets" / "esg_logo_white.png",
    PROJECT_ROOT / "assets" / "esg_logo_white.png",
]


def _set_run(run, size, bold=False, italic=False, color=RGBColor(22, 36, 44), font="Arial"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)


def shade(cell, hex_color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'))


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


def page_field(paragraph, muted_color):
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
    _set_run(run, 8, False, False, muted_color)


def find_logo():
    for path in LOGO_CANDIDATES:
        if path.exists():
            return path
    return None


def build_report():
    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.4)
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
    MUTED = RGBColor(90, 107, 119)
    DARK = RGBColor(22, 36, 44)
    WHITE = RGBColor(255, 255, 255)
    CONTENT = Cm(17.6)

    def add_p(text, size=10.0, bold=False, italic=False, color=DARK, before=0, after=6):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(before)
        par.paragraph_format.space_after = Pt(after)
        par.paragraph_format.line_spacing = 1.12
        _set_run(par.add_run(text), size, bold, italic, color)
        return par

    def add_h(text):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(16)
        par.paragraph_format.space_after = Pt(6)
        par.paragraph_format.keep_with_next = True
        _set_run(par.add_run(text), 13.5, True, False, NAVY)
        return par

    def callout(title, lines, bg, border, title_color):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        cell = table.cell(0, 0)
        cell.width = CONTENT
        shade(cell, bg)
        margins(cell, 110, 110, 130, 130)
        borders(cell, border, "6", "24", border)
        par = cell.paragraphs[0]
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(4)
        _set_run(par.add_run("■ " + title), 9.5, True, False, title_color)
        for line in lines:
            bp = cell.add_paragraph()
            bp.paragraph_format.space_before = Pt(2)
            bp.paragraph_format.space_after = Pt(2)
            bp.paragraph_format.line_spacing = 1.12
            _set_run(bp.add_run(line), 9.5, False, False, DARK)
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(6)

    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
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
        logo = find_logo()
        if logo:
            lp.add_run().add_picture(str(logo), width=Cm(4.4))
        else:
            _set_run(lp.add_run("ESG INSTITUTE"), 10, True, False, WHITE)
        rp = right.paragraphs[0]
        rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rp.paragraph_format.space_before = Pt(8)
        rp.paragraph_format.space_after = Pt(0)
        _set_run(rp.add_run("AUDYT RYZYK GREENWASHINGU"), 8, True, False, WHITE)
        rp2 = right.add_paragraph()
        rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rp2.paragraph_format.space_before = Pt(0)
        rp2.paragraph_format.space_after = Pt(0)
        _set_run(rp2.add_run("esginstitute.eu"), 8, False, False, RGBColor(129, 192, 223))

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.paragraph_format.space_before = Pt(2)
        _set_run(fp.add_run("ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Hi DESIGN  ·  str. "), 8, False, False, MUTED)
        page_field(fp, MUTED)

    add_p("ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO", size=9, bold=True, color=ACCENT, after=2)
    add_p("Audyt Ryzyk Greenwashingu: HI DESIGN", size=20, bold=True, color=PRIMARY, before=2, after=2)
    add_p(
        "Weryfikacja treści witryny https://hi-design.pl/ (strona główna, 101 podstron, 39 wpisów bloga i 329 kart produktów sklepu) "
        "wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE (UCPD), ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (UPNPR), "
        "ustawy o ochronie konkurencji i konsumentów oraz ustawy o zwalczaniu nieuczciwej konkurencji (UZNK). "
        "Stan zebrania i data audytu: 14 października 2026 r.",
        size=10.0, italic=True, color=MUTED, after=8,
    )

    if SHOT.exists():
        shot_par = doc.add_paragraph()
        shot_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shot_par.paragraph_format.space_before = Pt(2)
        shot_par.paragraph_format.space_after = Pt(2)
        shot_par.add_run().add_picture(str(SHOT), width=Cm(15.4))
        add_p(
            "Zrzut strony głównej https://hi-design.pl/ z 14 października 2026 r. Hero nie zawiera twierdzenia środowiskowego. "
            "Hasła ekologiczne znajdują się w katalogu gadżetów, na podstronie materiałów drukowanych i we wpisach bloga.",
            size=8.5, italic=True, color=MUTED, after=8,
        )

    meta = [
        ("Badany podmiot", "Hi DESIGN Sebastian Siwy, ul. Różana 21, 55-010 Radomierzyce, NIP 8941704766. Punkt kontaktu podany na stronie: ul. Januszowicka 5 lok. 425, 53-135 Wrocław. Telefony: +48 71 307 71 15, +48 505 731 310."),
        ("Przedmiot audytu", "Komunikacja marketingowa i ofertowa opublikowana 14 października 2026 r. na https://hi-design.pl/: strona główna, podstrony usługowe, wpisy bloga oraz karty produktów WooCommerce (329 pozycji, w tym warianty ilościowe powtarzające claim rodzica)."),
        ("Sektor działalności", "Agencja reklamowa: branding, strony internetowe, druk i outdoor oraz sprzedaż gadżetów reklamowych. Kanał B2C (sklep i regulamin konsumencki) oraz B2B."),
        ("Reżim prawny", "Dyrektywa (UE) 2024/825 (ECGT, stosowanie od 27.09.2026, załącznik I pkt 4a i 4b), dyrektywa 2005/29/WE, UPNPR art. 5–7, ustawa o ochronie konkurencji i konsumentów art. 24 i 106, UZNK art. 16 i 18, norma ISO 14021 jako wzorzec deklaracji własnych. Nie stwierdzono claimów offsetowych (pkt 4c) ani własnego znaku zrównoważoności (pkt 2a)."),
        ("Data audytu", "14 października 2026 r."),
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
        _set_run(pk.add_run(k), 8.5, True, False, DARK)
        _set_run(pv.add_run(v), 8.5, i == 5, False, DANGER if i == 5 else DARK)
    add_p("", size=4, after=6)

    callout(
        "WERDYKT: NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)",
        [
            "Katalog gadżetów i wybrane podstrony usługowe spełniają przesłanki czarnej listy załącznika I do dyrektywy 2005/29/WE w brzmieniu nadanym dyrektywą (UE) 2024/825: pkt 4a (ogólne hasła «eko», «ekologiczny», «ekologiczne», «przyjazne dla środowiska», «papiery eko», «biodegradowalne» bez uznanej wysokiej efektywności ekologicznej w tym samym kanale) oraz pkt 4b (przeniesienie cechy jednego aspektu — papier kraft, bambusowe wykończenie, wielorazowość, wymiana grafiki — na cały produkt albo na wizerunek środowiskowy marki klienta).",
            "Najostrzejszy przypadek to torba na wino opisana w jednym zdaniu jako włóknina polipropylenowa i zarazem «przyjazna dla środowiska». Osobno torby kraft są sprzedawane pod nazwą handlową «eko Kraft» i zestawiane z plastikiem jako «ekologiczna alternatywa», bez FSC, udziału recyklatu, granic systemu i porównania cyklu życia.",
            "Dyrektywa ECGT stosuje się od 27 września 2026 r. Treści sklepu internetowego nie korzystają z taryfy dla starych zapasów. UPNPR art. 5 i 6 obowiązuje niezależnie od opóźnienia transpozycji. Utrzymanie tych kart naraża Hi DESIGN Sebastian Siwy na postępowanie Prezesa UOKiK (art. 24 i 106 ustawy o ochronie konkurencji i konsumentów, kara do 10% obrotu) oraz na przeniesienie wadliwego claimu do komunikacji klientów B2B, którzy użyją gadżetu jako dowodu «dbałości o środowisko».",
        ],
        "FFF1F2", "E11D48", DANGER,
    )

    add_h("1. Analiza Wstępna")
    add_p(
        "Hi DESIGN Sebastian Siwy jest wrocławską agencją reklamową. Witryna hi-design.pl sprzedaje strony internetowe, branding, SEO, social media oraz gadżety i materiały outdoorowe. "
        "Strona główna buduje narrację sprzedażową («First conversion. Potem skala. Zawsze.») i nie zawiera hasła środowiskowego. Usługi cyfrowe — hosting, certyfikat SSL, pozycjonowanie, copywriting — również nie przypisują firmie profilu ekologicznego. "
        "Sformułowania «naturalnie» i «naturalny» użyte przy ekspozycji marki albo przy kolorze beżowym bawełny są językiem potocznym, nie twierdzeniem o wpływie na środowisko, i nie zostały zakwalifikowane jako naruszenie."
    )
    add_p(
        "Ryzyko koncentruje się w sklepie z gadżetami. Przeszukano 329 kart produktów, 101 stron i 39 wpisów. Nazwa «eko Kraft» obejmuje torby papierowe beżowe i kolorowe oraz ich warianty 1000, 1500 i 2000 sztuk. "
        "Bidony, torby zakupowe, długopisy z bambusa lub papieru, akcesoria bambusowe i korkowe, termosy i kubki «z bambusowym wykończeniem», bramy pneumatyczne, słupy dmuchane, czapeczki z papieru kredowego oraz torby na wino z polipropylenu są opisywane jako ekologiczne, przyjazne środowisku albo jako dowód, że marka klienta «troszczy się o środowisko». "
        "Na żadnej z tych kart nie ma EU Ecolabel, certyfikatu ISO 14024 typu I, numeru FSC, udziału recyklatu, granic systemu ani streszczenia LCA."
    )
    add_p(
        "Mechanizm jest powtarzalny: jeden aspekt (surowiec papierowy, element bambusowy, możliwość ponownego użycia, technika sublimacji) zostaje nazwany ekologią całego produktu, a następnie przeniesiony na wizerunek zamawiającego. "
        "To greenlighting z pkt 4b oraz ogólne twierdzenie z pkt 4a. Praktyka nadzorcza UOKiK z 19 stycznia 2026 r. wobec Tchibo dotyczy właśnie znacznika «eko» bez czytelnych kryteriów, w tym przy wyrobach z surowców kopalnych. "
        "Nie stwierdzono claimu neutralności klimatycznej, offsetu, własnego znaku zrównoważoności ani przedstawiania obowiązku prawnego jako cechy wyróżniającej."
    )

    add_h("2. Analiza szczegółowa - tabela zbiorcza naruszeń")
    findings = [
        ("01", "Torby na wino z polipropylenu «przyjazne dla środowiska»", "ECGT zał. I pkt 4a · UPNPR art. 5 i 6", "BARDZO WYSOKIE", "Tworzywo sztuczne opisane w tym samym zdaniu jako cecha środowiskowa"),
        ("02", "Nazwa handlowa «eko Kraft» i «ekologiczny nośnik reklamy»", "ECGT zał. I pkt 4a · UPNPR art. 5", "WYSOKIE", "Ogólne «eko» w nazwie i opisie toreb bez certyfikatu typu I"),
        ("03", "Papier kraft jako «ekologiczna alternatywa» dla plastiku", "ECGT zał. I pkt 4a i 4b · UZNK art. 16 · UPNPR art. 6", "WYSOKIE", "Porównanie materiałowe bez cyklu życia i bez porównywalnego zakresu"),
        ("04", "Bidon jako «ekologiczny wybór» i dowód dbałości o środowisko", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5 i 6", "WYSOKIE", "Wielorazowość sprzedana jako profil środowiskowy produktu i marki klienta"),
        ("05", "Torby zakupowe «ekologiczne i wielorazowe»", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5", "WYSOKIE", "Wielorazowa torba bez podanego materiału ma dowodzić, że marka dba o środowisko"),
        ("06", "«Ekologiczne» długopisy z bambusa lub papieru, w zestawieniu z plastikiem", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5 i 6", "WYSOKIE", "Hasło eko bez udziału bambusa i bez oddzielenia od modeli plastikowych"),
        ("07", "Bambus, korek i «naturalny materiał» jako «eko wizerunek»", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5", "WYSOKIE", "Element wykończenia albo jedna pozycja kolekcji niesie claim całej kategorii"),
        ("08", "«Ekologiczny nadruk» sublimacyjny i «bezpieczeństwo dla środowiska»", "ECGT zał. I pkt 4a · UPNPR art. 5 i 6", "WYSOKIE", "Technika druku przedstawiona jako cecha środowiskowa bramy pneumatycznej"),
        ("09", "Słup dmuchany: «Ekologia» przez wymianę grafiki", "ECGT zał. I pkt 4a i 4b · UPNPR art. 6", "WYSOKIE", "Wielokrotne użycie stelaża nazwane ekologią całego nośnika"),
        ("10", "Czapeczki z papieru kredowego «przyjazne dla środowiska»", "ECGT zał. I pkt 4a · UPNPR art. 5 i 6", "WYSOKIE", "Jednorazowy papier kredowy z drukiem UV bez podstawy claimu"),
        ("11", "FAQ druku: «papiery eko» i «materiały biodegradowalne»", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5 i 6", "WYSOKIE", "Brak normy, środowiska rozkładu, udziału oferty i warunków biodegradacji"),
        ("12", "Piersiówki: «skóra ekologiczna» bez wskazania tworzywa", "ECGT zał. I pkt 4a · UPNPR art. 6", "ŚREDNIE", "Przymiotnik «ekologiczna» przy wykończeniu może być odczytany jako cecha środowiskowa"),
        ("13", "Blog: wielorazowość i «bardziej ekologiczne niż konkurencja»", "ECGT zał. I pkt 4a · UZNK art. 16 · UPNPR art. 5", "ŚREDNIE", "Artykuły zrównują reuse z ekologią i instruują claim porównawczy bez metody"),
    ]
    sum_table = doc.add_table(rows=1 + len(findings), cols=5)
    sum_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_table.autofit = False
    sum_table.style = "Table Grid"
    col_w = (Cm(1.1), Cm(5.5), Cm(4.3), Cm(2.5), Cm(4.2))
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota naruszenia"]
    for c, htxt in enumerate(headers):
        cell = sum_table.rows[0].cells[c]
        cell.width = col_w[c]
        shade(cell, "08181F")
        margins(cell, 50, 50, 50, 50)
        par = cell.paragraphs[0]
        par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
        _set_run(par.add_run(htxt), 8, True, False, WHITE)
    for r, row in enumerate(findings):
        bg = "FFFFFF" if r % 2 == 0 else "F8FAFC"
        for c, val in enumerate(row):
            cell = sum_table.rows[r + 1].cells[c]
            cell.width = col_w[c]
            shade(cell, bg)
            margins(cell, 40, 40, 50, 50)
            borders(cell)
            par = cell.paragraphs[0]
            par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
            run = par.add_run(val)
            if c == 3:
                color = WARNING if val == "ŚREDNIE" else DANGER
                _set_run(run, 7.5, True, False, color)
            elif c == 0:
                _set_run(run, 8, True, False, NAVY)
            else:
                _set_run(run, 7.5, False, False, DARK)
    add_p("", size=4, after=4)

    add_h("3. Szczegółowa analiza przypadków")
    cases = [
        {
            "title": "PRZYPADEK 01: Torby na wino z polipropylenu opisane jako przyjazne dla środowiska",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 ust. 1 oraz art. 6 ust. 1",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (Czarna lista)",
            "quotes": [
                "„Wykonane z wysokiej jakości włókniny non woven, torby są odporne na wilgoć i intensywne użytkowanie, a jednocześnie przyjazne dla środowiska.”",
                "„Torby na wino z logo – polipropylenowe.”",
            ],
            "assessment": (
                "Karta https://hi-design.pl/torby-na-wino-z-logo-polipropylenowe/ w jednym akapicie podaje surowiec (włóknina polipropylenowa, non woven) i ogólną korzyść środowiskową. "
                "Polipropylen jest tworzywem sztucznym. «Przyjazne dla środowiska» jest podręcznikowym ogólnym twierdzeniem środowiskowym z pkt 4a: nie wskazuje aspektu, jednostki, certyfikatu typu I ani wyniku cyklu życia. "
                "Możliwość prania w niskiej temperaturze, opisana niżej na karcie, jest informacją użytkową. Nie stanowi uznanej wysokiej efektywności ekologicznej i nie zdejmuje zakazu hasła ogólnego. "
                "Zestawienie polimeru z przymiotnikiem środowiskowym w tym samym zdaniu jest kwalifikowane wprowadzenie w błąd w rozumieniu art. 5 UPNPR oraz pominięciem istotnym z art. 6: odbiorca nie dostaje informacji, że «przyjazność» nie wynika z surowca. "
                "Wzorzec sprawy Tchibo (zarzuty Prezesa UOKiK z 19 stycznia 2026 r.) obejmuje właśnie znacznik eko przy wyrobach powiązanych z surowcami kopalnymi, bez czytelnych kryteriów."
            ),
            "remediation": (
                "Usunąć zwrot «przyjazne dla środowiska» z karty, skrótu, tytułu SEO i wszelkich wariantów. Zostawić nazwę surowca i cechę użytkową (wielorazowość, pranie). "
                "Nie zastępować hasła synonimem «eko», «green» ani ikoną liścia."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Torba na wino z włókniny polipropylenowej (tworzywo sztuczne). Można ją prać w niskiej temperaturze i używać wielokrotnie. Opis nie jest twierdzeniem o przewadze środowiskowej wobec innych opakowań.”",
        },
        {
            "title": "PRZYPADEK 02: Nazwa handlowa «eko Kraft» i hasło ekologicznego nośnika reklamy",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 ust. 1",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Torby papierowe z logo beżowe – eko Kraft.”",
                "„Torby Kraft z nadrukiem to praktyczne i estetyczne rozwiązanie dla firm szukających ekologicznego nośnika reklamy.”",
                "„Eventy firmowe i targi – estetyczne i ekologiczne torby dla klientów.”",
                "„Postaw na torby papierowe z logo beżowe – eko Kraft i wyróżnij swoją markę w praktyczny i ekologiczny sposób.”",
                "„Kolorowe, ekologiczne torby Kraft wykonane z prążkowanego papieru.”",
            ],
            "assessment": (
                "Definicja twierdzenia dotyczącego ekologiczności w ECGT obejmuje nazwę produktu. Człon «eko» w tytule karty, w adresie URL i w wariantach ilościowych (1000, 1500 i 2000 sztuk, osobno dla beżu i koloru) jest ogólnym twierdzeniem środowiskowym. "
                "Na kartach podano gramaturę (90 g na wariantach beżowych), format, sitodruk i paletę Pantone. Nie podano certyfikatu FSC ani PEFC, udziału makulatury, wybielania, pochodzenia włókna ani znaku EU Ecolabel lub innego ISO 14024 typu I. "
                "Sam fakt, że torba jest papierowa, nie jest uznaną wysoką efektywnością ekologiczną. Papier barwiony («kolorowe», «prążkowany») dodatkowo nie jest tożsamy z niebarwionym kraftem, a karta i tak nazywa go ekologicznym. "
                "Dopisek techniczny o druku nie sanuje hasła «eko» w tytule: specyfikacja zdejmuje status hasła ogólnego tylko wtedy, gdy w tym samym kanale wskazuje konkretną, sprawdzalną cechę odpowiadającą temu hasłu. Gramatura i Pantone takiej cechy nie stanowią."
            ),
            "remediation": (
                "Zmienić nazwy produktów, slugi i tytuły SEO: usunąć «eko». Tej samej korekty wymagają warianty ilościowe, które powtarzają nazwę rodzica. "
                "Hasło wolno przywrócić wyłącznie po uzyskaniu certyfikatu typu I albo po publikacji w tym samym widoku karty konkretnego atrybutu (na przykład numeru licencji FSC obejmującej ten papier), z zakresem ograniczonym do papieru, z wyłączeniem nadruku i uchwytów."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Torba papierowa, papier kraft 90 g/m², kolor beżowy, nadruk sitodrukiem w 1–2 kolorach Pantone. Jeśli papier ma certyfikat gospodarki leśnej, podać system, numer licencji i zakres (wyłącznie papier). Bez takiego certyfikatu nie używać słowa «eko».”",
        },
        {
            "title": "PRZYPADEK 03: Papier kraft jako ekologiczna i trwała alternatywa dla plastiku",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 6 · UZNK art. 16",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Torby cateringowe Kraft beżowe – ekologiczne i efektowne rozwiązanie dla branży gastronomicznej.”",
                "„Ekologiczne i trwałe – papier Kraft jest alternatywą dla plastikowych opakowań.”",
                "„wykonanie z ekologicznego i wytrzymałego papieru Kraft (gramatura 70 g).”",
            ],
            "assessment": (
                "Karta toreb cateringowych łączy ogólne «ekologiczne» z porównaniem do plastiku. Porównanie materiałowe jest reklamą porównawczą w rozumieniu art. 16 UZNK i dyrektywy 2006/114/WE: musi być obiektywne, sprawdzalne i obejmować istotne cechy. "
                "Jednorazowa torba papierowa 70 g nie jest z tego tylko powodu lepsza środowiskowo od opakowania z tworzywa. Wynik zależy od masy, liczby użyć, transportu, źródła włókna i końca życia. Tych danych na karcie nie ma. "
                "Jest to ukryty kompromis (art. 6 UPNPR): prawdziwa informacja o surowcu papierowym została podana tak, jakby rozstrzygała cały profil środowiskowy. "
                "«Trwałe» w tym samym punkcie odnosi się do opakowania jednorazowego na żywność i nie może zastąpić dowodu środowiskowego. Warianty 1000, 1500 i 2000 sztuk powielają nazwę i skrót oferty."
            ),
            "remediation": (
                "Usunąć «ekologiczne» z tytułu i z listy zalet oraz zdanie o alternatywie dla plastiku. Jeśli porównanie ma zostać, potrzebne jest zestawienie cyklu życia o tym samym zakresie, jednostce funkcjonalnej i źródle, dostępne z karty produktu. "
                "Do tego czasu opis ograniczyć do funkcji cateringowej i parametrów fizycznych."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Torba cateringowa z papieru kraft 70 g/m², poszerzone dno, obciążenie 4–6 kg, nadruk sitodrukiem w 1–2 kolorach. Opakowanie jednorazowe. Karta nie porównuje go z opakowaniami z tworzyw.”",
        },
        {
            "title": "PRZYPADEK 04: Bidon reklamowy jako ekologiczny wybór i dowód troski o środowisko",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5 i art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Bidony reklamowe z logo to praktyczny i ekologiczny gadżet, który codziennie buduje widoczność Twojej marki.”",
                "„podkreślają dbanie o ekologie i aktywny tryb życia.”",
                "„ekologiczny wybór, który buduje wizerunek. Firmy coraz częściej stawiają na produkty wielorazowe. Dzięki bidonom z logo pokazujesz, że troszczysz się o środowisko i wspierasz zdrowe nawyki.”",
            ],
            "assessment": (
                "Karta modelu SHOT i warianty 20, 50 oraz 100 sztuk nie podają surowca bidonu, pojemności ani udziału recyklatu, a mimo to nazywają produkt ekologicznym. "
                "Wielorazowość jest jednym aspektem użytkowania. Pkt 4b zakazuje twierdzenia o całym produkcie, gdy podstawa dotyczy tylko aspektu, oraz przenoszenia tej cechy na działalność zamawiającego. "
                "Zdanie «pokazujesz, że troszczysz się o środowisko» jest właśnie takim przeniesieniem: gadżet ma dowodzić profilu środowiskowego marki klienta. Klient B2B, który powtórzy to zdanie we własnym raporcie albo na evencie, publikuje claim bez dowodu. "
                "«Dbanie o ekologie» jest ogólnym hasłem z pkt 4a, dodatkowo bez wskazania, czego dotyczy (surowiec, koniec życia, zastąpienie butelki jednorazowej, liczba użyć)."
            ),
            "remediation": (
                "Wycofać «ekologiczny gadżet», «ekologiczny wybór» i zdanie o trosce zamawiającego o środowisko ze wszystkich wariantów ilościowych. "
                "Na karcie podać materiał i pojemność. O wielorazowości pisać jako o cesze użytkowej, z warunkiem rzeczywistego zastępowania opakowań jednorazowych, bez wnioskowania o profilu ESG klienta."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Bidon wielorazowy, model SHOT. Materiał: [podać, np. stal nierdzewna / tworzywo], pojemność: [ml]. Wielokrotne użycie może zmniejszyć liczbę jednorazowych opakowań na napoje tylko wtedy, gdy bidon je faktycznie zastępuje. Karta nie opisuje profilu środowiskowego marki zamawiającego.”",
        },
        {
            "title": "PRZYPADEK 05: Torby zakupowe «ekologiczne i wielorazowe»",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Ekologiczne i wielorazowe. W dobie świadomych zakupów torby wielorazowe są naturalnym wyborem. Pokazujesz, że marka dba o środowisko i oferuje rozwiązania z duchem czasu.”",
                "„dla sklepów chcących wprowadzić torby ekologiczne z własnym stylem.”",
            ],
            "assessment": (
                "Karta toreb TASKE nie wskazuje tkaniny, gramatury ani kraju produkcji, a nagłówek korzyści brzmi «Ekologiczne i wielorazowe». "
                "Wielorazowa torba bawełniana, poliestrowa albo z mieszanki ma różny próg użyć, od którego zastępuje torbę jednorazową. Bez materiału i bez tego progu hasło «ekologiczne» jest ogólne (pkt 4a). "
                "Zdanie o marce, która «dba o środowisko», przenosi cechę gadżetu na cały wizerunek zamawiającego (pkt 4b). "
                "Na tej samej witrynie torby bawełniane są opisane faktograficznie («100% bawełna», «kolor beżowy (naturalny)» jako kolor). Ten opis nie jest naruszeniem. Naruszający jest dopiero przymiotnik środowiskowy na karcie TASKE."
            ),
            "remediation": (
                "Usunąć «ekologiczne» i zdanie o dbałości marki o środowisko. Dodać skład surowcowy modelu TASKE. "
                "Jeśli torba ma certyfikat (na przykład GOTS dla bawełny organicznej albo GRS dla recyklatu), podać go z zakresem. Bez certyfikatu nie przywracać hasła ogólnego."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Torba wielorazowa, model TASKE. Skład: [włókno i udział procentowy]. Przeznaczona do wielokrotnego noszenia. Wielorazowość nie jest certyfikowaną przewagą środowiskową, dopóki karta nie wskaże materiału i — jeśli istnieje — certyfikatu z zakresem.”",
        },
        {
            "title": "PRZYPADEK 06: Ekologiczne długopisy z bambusa lub papieru w ofercie z modelami plastikowymi",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5 i art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„W ofercie znajdziesz szeroki wybór modeli – od klasycznych długopisów plastikowych, przez eleganckie metalowe, po ekologiczne wersje z bambusa lub papieru.”",
                "„długopisy metalowe, ekologiczne i plastikowe.”",
                "„Ekologiczne (papier/bamboo/kraft).”",
            ],
            "assessment": (
                "«Ekologiczne wersje» to ogólne twierdzenie o części oferty (pkt 4a). Bambus albo papier w korpusie nie przesądza o profilu długopisu: wkład, sprężyna, klips i mechanizm są zwykle z metalu lub tworzywa, a karta tego nie rozdziela. "
                "Pkt 4b zabrania claimu o produkcie, gdy dotyczy on tylko aspektu. Skrót «długopisy metalowe, ekologiczne i plastikowe» ustawia «ekologiczne» jako równorzędny typ obok plastiku, bez kryterium, które ten typ wyróżnia. "
                "Odbiorca nie wie, który SKU jest «wersją ekologiczną», jaki jest udział bambusa w masie i czy papier jest z recyklingu. "
                "Odrębnie karta gwarantuje «długą żywotność» długopisów z aluminium, metalu lub tworzyw. To obietnica trwałości bez metody badania. Nie została zakwalifikowana jako pkt 23d–j, bo nie podaje fałszywej krotności ani okresu «niezniszczalności», ale nie może służyć za dowód środowiskowy. "
                "Adres https://hi-design.pl/dlugopisy-reklamowe-z-logo-rozne-modele/ zwraca przekierowanie 301 na https://dlugopisy-reklamowe.hi-design.pl/. Tam filtr oferty nazywa się «Ekologiczne (papier/bamboo/kraft)». Dłuższe zdanie o «ekologicznych wersjach» pozostaje w opublikowanym opisie produktu na hi-design.pl i w wynikach jej wyszukiwarki. Oba miejsca wymagają tej samej korekty."
            ),
            "remediation": (
                "Na karcie zbiorczej usunąć przymiotnik «ekologiczne». Przy konkretnym modelu podać skład korpusu i pozostałych elementów oraz udział masowy bambusa lub papieru. "
                "Plastikowych i metalowych modeli nie grupować pod wspólnym skrótem zawierającym słowo «ekologiczne». "
                "Na https://dlugopisy-reklamowe.hi-design.pl/ zmienić etykietę filtra «Ekologiczne (papier/bamboo/kraft)» na nazwę surowca, bez przymiotnika środowiskowego."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Długopis, korpus z bambusa. Wkład, klips i mechanizm: [materiał]. Udział bambusa w masie wyrobu: [x%]. Claim dotyczy korpusu, nie całego długopisu. Modele plastikowe i metalowe są opisane osobno, bez przymiotnika środowiskowego.”",
        },
        {
            "title": "PRZYPADEK 07: Bambus, korek i «naturalny materiał» jako podstawa eko-wizerunku",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Bambusowy stojak na telefon – prosty, elegancki i ekologiczny dodatek do każdego stanowiska pracy.”",
                "„Naturalne materiały, takie jak bambus, podkreślają nowoczesny i eko wizerunek marki.”",
                "„ekologiczne materiały, w tym bambus.”",
                "„Korkowa podkładka pod mysz – ekologiczna, lekka i idealnie przylegająca do biurka. Naturalny materiał wygląda stylowo i przyjemnie w dotyku.”",
                "„Bambusowe lusterko – lekkie, ekologiczne i zawsze pod ręką.”",
                "„ekologiczne wersje z bambusowym wykończeniem.”",
                "„ekologiczne z bambusowym wykończeniem.”",
            ],
            "assessment": (
                "Ten sam schemat występuje na kartach akcesoriów do telefonu (w tym wariantach 20 i 50 sztuk), akcesoriów do laptopów, gadżetów beauty, termosów i kubków. "
                "«Eko», «ekologiczny» i «ekologiczne materiały» są hasłami ogólnymi. Kolekcje obejmują też kabel USB, organizer z ładowarką, czyścik i nawilżacz, a zdanie o «eko wizerunku marki» odnosi się do zestawu, nie do jednego elementu bambusowego. To pkt 4b. "
                "Przy termosach i kubkach claim jest jeszcze węższy: «bambusowe wykończenie» przy korpusie ze stali lub ceramiki. Wykończenie nie jest całym produktem. "
                "«Naturalny materiał» przy korku sugeruje korzyść środowiskową (surowiec odnawialny) bez informacji o lepiszczu, spodzie podkładki i końcu życia. Odczucie w dotyku nie jest dowodem środowiskowym. "
                "Korek i bambus mogą być opisane jako surowiec. Nie mogą nieść hasła «ekologiczny», dopóki karta nie wskaże udziału tego surowca i nie wyłączy elektroniki oraz tworzyw. "
                "Zdanie o termosach jest na karcie dostępnej bezpośrednio. Adres karty kubków zwraca 301 na kategorię https://gadzety-reklamowe.hi-design.pl/produkt-kategoria/kubki-i-filizanki-termiczne/; sformułowanie «ekologiczne z bambusowym wykończeniem» pozostaje w opisie produktu na hi-design.pl."
            ),
            "remediation": (
                "Na każdej z wymienionych kart usunąć «ekologiczny», «eko wizerunek» i «ekologiczne materiały». "
                "Zostawić nazwę elementu (bambus, korek) i wskazać, których części wyrobu nie obejmuje. Nie obiecywać zamawiającemu, że gadżet buduje profil środowiskowy jego marki."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Stojak na telefon z elementem bambusowym. Bambus dotyczy [podać część]. Złącza, elektronika i tworzywa w tym samym gadżecie nie są opisane jako materiał środowiskowy. Kolekcja zawiera także pozycje bez bambusa; nie obejmuje ich ten opis.”",
        },
        {
            "title": "PRZYPADEK 08: Ekologiczny nadruk sublimacyjny i bezpieczeństwo dla środowiska",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 i art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Ekologiczny nadruk – technologia sublimacji zapewnia intensywne kolory, wysoką jakość oraz bezpieczeństwo dla środowiska.”",
                "„ekologiczna technologia druku sublimacyjnego.”",
            ],
            "assessment": (
                "Karta bram pneumatycznych przypisuje sublimacji trzy skutki: kolor, jakość i «bezpieczeństwo dla środowiska». Dwa pierwsze są cechami druku. Trzeci jest twierdzeniem środowiskowym bez emisji procesu, składu barwnika, substratu tkaniny i końca życia konstrukcji pneumatycznej. "
                "Technika druku nie jest uznaną wysoką efektywnością ekologiczną. Nazwanie jej «ekologiczną» jest hasłem ogólnym (pkt 4a). "
                "Odbiorca może odczytać claim jako cechę całej bramy, choć zdanie mówi o nadruku. Karta nie rozdziela tych zakresów na tyle, by wyłączyć wprowadzenie w błąd co do produktu (art. 6 UPNPR). "
                "Odporność na UV i wielokrotny montaż, opisane obok, są cechami użytkowymi i nie dowodzą bezpieczeństwa środowiskowego. "
                "Publiczny adres karty zwraca 301 na https://outdoor.hi-design.pl/produkt-tag/brama-pneumatyczna/. Archiwum tagu nie powtarza tego zdania w HTML, ale opis jest opublikowany (rekord produktu, status publish, ostatnia zmiana 13 czerwca 2026 r.) i jest zwracany przez wyszukiwarkę hi-design.pl. Korekta dotyczy rekordu produktu, nie tylko szablonu archiwum."
            ),
            "remediation": (
                "Usunąć «ekologiczny nadruk», «bezpieczeństwo dla środowiska» i «ekologiczna technologia druku» ze skrótu oferty i z listy zalet. "
                "Sublimację opisać jako metodę przeniesienia grafiki i trwałości koloru."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Grafika bramy nanoszona metodą sublimacji. Opis dotyczy intensywności i trwałości koloru. Nie jest twierdzeniem o wpływie nadruku ani konstrukcji pneumatycznej na środowisko.”",
        },
        {
            "title": "PRZYPADEK 09: Słup dmuchany — «Ekologia» sprowadzona do wymiany grafiki",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Ekologia i wielokrotne użycie – możliwość wymiany grafik bez konieczności zakupu nowego słupa.”",
            ],
            "assessment": (
                "Na karcie AdColumn Air słowo «Ekologia» jest nazwą korzyści. Jedyną podstawą jest wymiana wydruku na tym samym słupie. "
                "To twierdzenie o całym nośniku oparte na jednym aspekcie użytkowania (pkt 4b). Nowa grafika nadal powstaje, a karta nie podaje materiału poszycia, sposobu zasilania pompki ani końca życia słupa. "
                "Wielokrotne użycie stelaża jest faktem organizacyjnym. Nazwanie go ekologią bez granic systemu jest hasłem ogólnym (pkt 4a), którego nie sanuje sama wzmianka o wymianie grafiki: wzmianka wyjaśnia intencję sprzedawcy, a nie wykazuje wysokiej efektywności ekologicznej. "
                "Adres karty zwraca 301 na kategorię słupków https://outdoor.hi-design.pl/produkt-kategoria/reklama-zewnetrzna-slupki-reklamowe/. Zdanie pozostaje w opublikowanym opisie produktu i w indeksie wyszukiwania hi-design.pl."
            ),
            "remediation": (
                "Zmienić punkt korzyści na opis operacyjny: ten sam słup, nowa grafika. Nie używać nagłówka «Ekologia»."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Konstrukcję słupa AdColumn Air można wykorzystać ponownie, zamawiając nową grafikę. Twierdzenie dotyczy wielokrotnego użycia stelaża. Nie opisuje materiału poszycia ani wpływu nowego wydruku na środowisko.”",
        },
        {
            "title": "PRZYPADEK 10: Czapeczki z papieru kredowego jako ekologiczne i przyjazne dla środowiska",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 5 i art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Nasze Czapeczki Reklamowe z Logo wykonano z wysokiej jakości papieru kredowego 300 g, dodatkowo uszlachetnionego.”",
                "„Druk jednostronny w pełnym kolorze (4/0) w technologii UV.”",
                "„Ekologiczne i efektowne. Nasze papierowe czapeczki są przyjazne dla środowiska, a jednocześnie przyciągają uwagę (…) odbiorców, którzy cenią kreatywne i ekologiczne rozwiązania.”",
            ],
            "assessment": (
                "«Przyjazne dla środowiska» i «ekologiczne rozwiązania» są hasłami ogólnymi z pkt 4a. Karta jednocześnie ujawnia papier kredowy 300 g, uszlachetnienie (bez wskazania, na czym polega) i druk UV. "
                "Jest to produkt eventowy, składany, o charakterze jednorazowym. Papier kredowy z zadrukiem UV nie staje się wyrobem o wysokiej efektywności ekologicznej przez to, że nie jest plastikiem. "
                "Art. 6 UPNPR wymaga informacji, których brak zmienia decyzję: recyklat, możliwość przetworzenia po evencie, sens uszlachetnienia dla recyklingu. Tych informacji nie ma. "
                "Odbiorca «ceniący ekologiczne rozwiązania» jest wprost zachęcany do odczytania gadżetu jako wyboru środowiskowego."
            ),
            "remediation": (
                "Usunąć blok «Ekologiczne i efektowne» oraz «przyjazne dla środowiska». Zostawić papier kredowy, gramaturę i druk UV jako opis wykonania."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Czapeczka eventowa z papieru kredowego 300 g, druk UV 4/0, do złożenia na miejscu. Produkt jednorazowy. Karta nie zawiera twierdzenia środowiskowego.”",
        },
        {
            "title": "PRZYPADEK 11: FAQ druku — papiery eko i materiały biodegradowalne",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a i 4b · UPNPR art. 5 i art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista)",
            "quotes": [
                "„Czy drukujecie ekologiczne materiały reklamowe? Tak! Mamy opcje na papierach eko i materiałach biodegradowalnych – dla firm, które dbają o wizerunek i środowisko.”",
            ],
            "assessment": (
                "Odpowiedź na podstronie materiałów drukowanych zawiera dwa ogólne twierdzenia: «ekologiczne materiały» i «papiery eko», oraz twierdzenie o biodegradacji bez normy, środowiska rozkładu (kompost przemysłowy, gleba, woda), temperatury i czasu. "
                "Biodegradacja bez tych warunków jest pominięciem istotnym (art. 6 UPNPR). Odbiorca nie wie, czy rozkład wymaga instalacji przemysłowej i czy dotyczy papieru, folii czy gadżetu. "
                "Zwrot «dla firm, które dbają o wizerunek i środowisko» przenosi opcję druku na motywację środowiskową zamawiającego (pkt 4b). Reszta strony opisuje ulotki, roll-upy, flagi i namioty bez żadnego atrybutu środowiskowego, więc claim nie dotyczy oferty jako całości, ale jest tak napisany, jakby był jej właściwością. "
                "Dopóki w tym samym akapicie nie ma nazwy papieru, certyfikatu i warunków rozkładu, hasła nie wolno utrzymać. Wniosek Green Claims nie jest prawem obowiązującym; wymóg konkretyzacji wynika już z ECGT i UPNPR."
            ),
            "remediation": (
                "Usunąć odpowiedź FAQ albo przepisać ją jako zapowiedź doboru papieru po otrzymaniu specyfikacji, bez słów «eko» i «biodegradowalne». "
                "Claim o biodegradacji publikować dopiero ze wskazaniem normy, środowiska, czasu i tego, którego elementu zlecenia dotyczy."
            ),
            "salvage": False,
            "claim_title": "SFORMUŁOWANIE NIEMOŻLIWE DO UTRZYMANIA BEZ DOWODU",
            "claim": "Hasła «papiery eko» i «materiały biodegradowalne» należy usunąć. Wolno je zastąpić dopiero zdaniem w rodzaju: «Papier [nazwa], certyfikat [system i numer], dotyczy arkusza, nie uszlachetnienia ani farby. Biodegradacja: norma [numer], środowisko [kompost przemysłowy / inne], czas [x], potwierdzenie [jednostka i data].» Bez tych danych claimu nie publikować.",
        },
        {
            "title": "PRZYPADEK 12: «Skóra ekologiczna» na piersiówkach",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a · UPNPR art. 6",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "„Wykonane z wysokiej jakości materiałów – najczęściej stali nierdzewnej, często uzupełnionej elementami skóry ekologicznej lub drewna.”",
            ],
            "assessment": (
                "W obrocie «skóra ekologiczna» bywa handlową nazwą skóry syntetycznej, a nie certyfikatem środowiskowym. Karta nie wyjaśnia tego znaczenia i nie podaje polimeru. "
                "Przymiotnik «ekologicznej» pada w opisie materiału obok stali i drewna, więc część odbiorców odczyta go jako cechę środowiskową wykończenia. W tym odczytaniu jest to hasło ogólne z pkt 4a, bez typu I i bez składu. "
                "Ryzyko jest średnie, bo zwrot nie jest głównym hasłem karty. Pozostaje jednak pominięcie z art. 6: odbiorca nie wie, czy dotyka tworzywa, czy materiału o deklarowanej przewadze środowiskowej."
            ),
            "remediation": (
                "Zastąpić «skórę ekologiczną» nazwą tworzywa (na przykład skóra syntetyczna, poliuretan) albo dopisać w tym samym zdaniu, że określenie oznacza imitację skóry i nie jest twierdzeniem środowiskowym."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Piersiówka ze stali nierdzewnej, z elementem wykończenia ze skóry syntetycznej (poliuretan) albo z drewna — wariant wskazany przy modelu. Określenie nie oznacza cechy środowiskowej.”",
        },
        {
            "title": "PRZYPADEK 13: Wpisy bloga zrównujące wielorazowość z ekologią i instruujące claim porównawczy",
            "basis": "Podstawa prawna: ECGT, załącznik I pkt 4a · UZNK art. 16 · UPNPR art. 5",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "„Oszczędzasz w ten sposób środki i chronisz środowisko.”",
                "„Jeśli firma komunikuje ekologię i odpowiedzialność, naturalnym wyborem są produkty wielokrotnego użytku lub wykonane z materiałów przyjaznych środowisku.”",
                "„jeśli Twoim klientem są osoby szukające produktów ekologicznych, warto na stronie umieścić informacje o tym, jakie produkty oferujesz i dlaczego są one bardziej ekologiczne niż oferta konkurencji.”",
            ],
            "assessment": (
                "Trzy fragmenty nie są kartą produktu, ale są komunikacją handlową agencji: wpis o trybunce (wymiana grafiki «chroni środowisko»), wpis o gadżetach (wielokrotny użytek lub «materiały przyjazne środowisku» jako naturalny wybór dla firmy komunikującej ekologię) oraz poradnik o stronach www, który zaleca napisanie, dlaczego oferta jest «bardziej ekologiczna niż oferta konkurencji». "
                "Pierwsze dwa powtarzają schemat pkt 4a i 4b w treści doradczej, którą agencja sama stosuje w katalogu. Trzeci zaleca reklamę porównawczą środowiskową bez wymogu obiektywności i sprawdzalności z art. 16 UZNK. "
                "Przykłady InPost, LPP, 4F i Żywiec Zdrój w innym wpisie są ilustracją cudzych programów, nie claimem Hi DESIGN o własnych usługach, i nie zostały ocenione jako osobne naruszenie. "
                "Ryzyko tego przypadku jest średnie, bo nie oznacza konkretnego SKU. Skutek jest praktyczny: agencja uczy klientów formuł, których własny katalog nie potrafi udowodnić."
            ),
            "remediation": (
                "We wpisie o trybunce zastąpić «chronisz środowisko» zdaniem o ponownym użyciu stelaża. "
                "We wpisie o gadżetach dopisać, że wielokrotny użytek i dobór materiału nie są twierdzeniem środowiskowym bez składu, certyfikatu i zakresu. "
                "Z poradnika usunąć zalecenie claimu «bardziej ekologiczne niż konkurencja» albo uzależnić je od porównania, które spełnia art. 16 UZNK."
            ),
            "salvage": True,
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": "„Wymiana samej grafiki pozwala użyć tego samego stelaża przy kolejnej kampanii. To decyzja operacyjna, nie twierdzenie, że stoisko jest ekologiczne. Porównanie środowiskowe z inną ofertą publikujemy tylko z tym samym zakresem, jednostką i źródłem danych, dostępnymi przy haśle.”",
        },
    ]

    for case in cases:
        hpar = doc.add_paragraph()
        hpar.paragraph_format.space_before = Pt(12)
        hpar.paragraph_format.space_after = Pt(2)
        hpar.paragraph_format.keep_with_next = True
        _set_run(hpar.add_run(case["title"]), 11.5, True, False, NAVY)
        meta_line = doc.add_paragraph()
        meta_line.paragraph_format.space_before = Pt(0)
        meta_line.paragraph_format.space_after = Pt(4)
        risk_color = WARNING if "ŚREDNIE" in case["risk"] else DANGER
        _set_run(meta_line.add_run(case["basis"] + "  |  " + case["risk"]), 8.5, True, False, risk_color)
        add_p("Cytat źródłowy:", size=10.0, bold=True, after=2)
        for q in case["quotes"]:
            add_p(q, size=10.0, italic=True, after=3)
        add_p("Ocena prawna:", size=10.0, bold=True, before=4, after=2)
        add_p(case["assessment"])
        add_p("Kierunek naprawy:", size=10.0, bold=True, before=2, after=2)
        add_p(case["remediation"])
        bg = "ECFDF5" if case["salvage"] else "F8FAFC"
        border = "059669" if case["salvage"] else "94A3B8"
        title_color = RGBColor(4, 120, 87) if case["salvage"] else MUTED
        callout(case["claim_title"], [case["claim"]], bg, border, title_color)

    add_h("4. Porównanie twierdzeń HI DESIGN z rzeczywistym stanem prawno-środowiskowym")
    add_p(
        "Zestawienie obejmuje te fragmenty witryny, które porównują materiały, techniki albo oferty albo przenoszą cechę jednego aspektu na cały produkt i na markę klienta.",
        after=6,
    )
    comp = [
        (
            "Torby na wino: włóknina polipropylenowa jest «przyjazna dla środowiska».",
            "Polipropylen jest tworzywem sztucznym. Ogólne «przyjazne dla środowiska» jest zakazane (ECGT pkt 4a), jeśli w tym samym kanale nie ma uznanej wysokiej efektywności ekologicznej. Pranie i wielokrotne użycie są cechą użytkową, nie certyfikatem.",
        ),
        (
            "Torby cateringowe: «papier Kraft jest alternatywą dla plastikowych opakowań» i dlatego jest «ekologiczny».",
            "Porównanie papieru jednorazowego z tworzywem wymaga tego samego zakresu i cyklu życia (UZNK art. 16). Gramatura 70 g i sitodruk takiej oceny nie zastępują. Ukryty kompromis podlega art. 6 UPNPR.",
        ),
        (
            "Nazwa «eko Kraft» oraz «ekologiczny nośnik reklamy» przy torbach 90 g, także barwionych.",
            "«Eko» w nazwie produktu jest twierdzeniem środowiskowym. Brak EU Ecolabel, ISO 14024 typu I, FSC i udziału recyklatu. Papier barwiony nie jest opisany inaczej niż papier beżowy, a oba noszą człon «eko».",
        ),
        (
            "Bidon i torba TASKE: wielorazowość oznacza, że marka klienta «troszczy się o środowisko» / «dba o środowisko».",
            "Wielorazowość jest aspektem użytkowania (pkt 4b). Nie wolno z niej wnosić profilu środowiskowego zamawiającego. Karta bidonu nie podaje nawet surowca.",
        ),
        (
            "Termos i kubek: «ekologiczne wersje z bambusowym wykończeniem» przy korpusie ze stali lub ceramiki.",
            "Wykończenie nie jest całym wyrobem. Claim o produkcie oparty na elemencie dekoracyjnym narusza pkt 4b. Stal i ceramika pozostają bez oceny cyklu życia.",
        ),
        (
            "Brama pneumatyczna: sublimacja zapewnia «bezpieczeństwo dla środowiska».",
            "Sublimacja opisuje przeniesienie barwnika i trwałość koloru. Bez danych o barwniku, tkaninie i konstrukcji nie jest twierdzeniem środowiskowym dopuszczalnym w pkt 4a.",
        ),
        (
            "FAQ druku: «papiery eko i materiały biodegradowalne» dla firm, które «dbają o środowisko».",
            "Brak nazwy papieru, certyfikatu, normy biodegradacji, środowiska rozkładu i czasu. Opcja w ofercie nie może być przedstawiona jako właściwość środowiskowa usługi.",
        ),
        (
            "Poradnik www: napisać, «dlaczego [produkty] są bardziej ekologiczne niż oferta konkurencji».",
            "Porównanie środowiskowe z konkurencją jest reklamą porównawczą (UZNK art. 16). Bez metody, zakresu i źródła jest wprowadzające w błąd. Agencja nie podaje tych warunków.",
        ),
    ]
    ctable = doc.add_table(rows=1 + len(comp), cols=2)
    ctable.alignment = WD_TABLE_ALIGNMENT.CENTER
    ctable.autofit = False
    ctable.style = "Table Grid"
    cw = (Cm(8.8), Cm(8.8))
    for c, htxt in enumerate(["Twierdzenie podmiotu", "Rzeczywisty stan prawno-środowiskowy"]):
        cell = ctable.rows[0].cells[c]
        cell.width = cw[c]
        shade(cell, "08181F")
        margins(cell, 50, 50, 60, 60)
        par = cell.paragraphs[0]
        par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
        _set_run(par.add_run(htxt), 8, True, False, WHITE)
    for r, (left_txt, right_txt) in enumerate(comp):
        bg = "FFFFFF" if r % 2 == 0 else "F8FAFC"
        for c, val in enumerate((left_txt, right_txt)):
            cell = ctable.rows[r + 1].cells[c]
            cell.width = cw[c]
            shade(cell, bg)
            margins(cell, 50, 50, 60, 60)
            borders(cell)
            par = cell.paragraphs[0]
            par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
            _set_run(par.add_run(val), 8, False, False, DARK)

    add_h("5. Ryzyka prawne i finansowe")
    risks = [
        (
            "1. Postępowanie Prezesa UOKiK",
            "Karty gadżetów są ofertą kierowaną także do konsumentów (regulamin zakupów i sklep). Ogólne hasła środowiskowe podlegają art. 5 i 6 UPNPR już teraz, a od 27 września 2026 r. dodatkowo czarnej liście ECGT. Praktyka naruszająca zbiorowe interesy konsumentów (art. 24 ustawy o ochronie konkurencji i konsumentów) może skutkować karą do 10% obrotu osiągniętego w roku poprzedzającym nałożenie kary (art. 106). Zarzuty UOKiK z 19 stycznia 2026 r. wobec Tchibo dotyczą znacznika «eko» bez kryteriów, w tym przy surowcach kopalnych — schemat zbieżny z torbą polipropylenową i z nazwą «eko Kraft».",
        ),
        (
            "2. Roszczenia konkurentów z UZNK",
            "Porównanie papieru kraft z «plastikowymi opakowaniami» oraz poradnik zalecający claim «bardziej ekologiczne niż oferta konkurencji» są reklamą porównawczą. Jeśli nie jest obiektywna i sprawdzalna, konkurent może żądać zaniechania, oświadczenia i naprawienia szkody (UZNK art. 16 i art. 18). Działanie sprzeczne z prawem w komunikacji B2B podlega także klauzuli generalnej art. 3 UZNK.",
        ),
        (
            "3. Przeniesienie claimu na klientów B2B i ich raportowanie",
            "Zdania «pokazujesz, że troszczysz się o środowisko» oraz «eko wizerunek marki» są napisane tak, by zamawiający użył gadżetu jako dowodu własnej odpowiedzialności. Klient objęty CSRD albo składający oświadczenie środowiskowe we własnym kanale przejmuje twierdzenie, którego Hi DESIGN nie udowodnił. Ryzyko wraca do agencji jako odpowiedzialność za treść, którą sama dostarczyła.",
        ),
        (
            "4. Kanał internetowy po 27 września 2026 r.",
            "ECGT stosuje się od 27 września 2026 r. Sklep online nie korzysta z taryfy starych zapasów, która dotyczy partii towaru w opakowaniu fizycznym wprowadzonej wcześniej. Nazwy «eko Kraft» w tytułach, adresach URL i wariantach ilościowych są treścią cyfrową i podlegają korekcie bez okresu przejściowego. Opóźnienie polskiej ustawy transponującej nie wstrzymuje UPNPR ani praktyki UOKiK.",
        ),
    ]
    for title, body in risks:
        add_p("• " + title, size=10.5, bold=True, before=8, after=2)
        add_p(body)

    add_h("6. Rekomendacje naprawcze")
    stages = [
        (
            "Etap 1: Natychmiastowy pull-off (14 października 2026 r.)",
            "1. Zdjąć «przyjazne dla środowiska» z karty toreb na wino z polipropylenu.\n"
            "2. Zdjąć człon «eko» z nazw, tytułów i widocznych nagłówków toreb kraft (beż, kolor, catering) oraz z wariantów 1000 / 1500 / 2000 sztuk. Slugi URL przekierować na nazwy bez «eko».\n"
            "3. Zdjąć «ekologiczny gadżet», «ekologiczny wybór» i zdanie o trosce klienta o środowisko z bidonów (warianty 20 / 50 / 100 sztuk) oraz «ekologiczne i wielorazowe» z toreb TASKE.\n"
            "4. Zdjąć «ekologiczny nadruk», «bezpieczeństwo dla środowiska» i nagłówek «Ekologia» ze słupa AdColumn Air.\n"
            "5. Wyłączyć odpowiedź FAQ o «papierach eko» i «materiałach biodegradowalnych».",
            "FFF1F2", "E11D48", DANGER,
        ),
        (
            "Etap 2: Korekta kart i wpisów (do 21 października 2026 r.)",
            "1. Przepisać długopisy, termosy, kubki, akcesoria bambusowe, podkładkę korkową, lusterko i czapeczki według wzorów z części 3. Zostawić surowiec, usunąć przymiotnik środowiskowy.\n"
            "2. Na piersiówkach zastąpić «skórę ekologiczną» nazwą tworzywa.\n"
            "3. Poprawić trzy wpisy bloga: trybunka, gadżety firmowe, poradnik o stronach www. Usunąć zalecenie claimu «bardziej ekologiczne niż konkurencja».\n"
            "4. Sprawdzić skróty ofertowe (excerpt WooCommerce), które powtarzają hasła niezależnie od opisu długiego.",
            "FFFBEB", "D97706", WARNING,
        ),
        (
            "Etap 3: Metodyka i dowody, jeśli claim ma wrócić (do 13 listopada 2026 r.)",
            "1. Dla papieru: albo sam opis fizyczny, albo licencja FSC/PEFC z numerem i zakresem «wyłącznie papier, bez nadruku i uchwytów», opublikowana na tej samej karcie.\n"
            "2. Dla bambusa i korka: udział masowy elementu i lista części wyłączonych (wkład, elektronika, tworzywo).\n"
            "3. Dla biodegradacji: norma, środowisko rozkładu, czas, jednostka i data badania. Do czasu tych danych claimu nie przywracać.\n"
            "4. Nie publikować porównania z plastikiem ani z konkurencją bez wspólnej jednostki funkcjonalnej i źródła.",
            "F0F9FF", "0284C7", ACCENT,
        ),
        (
            "Etap 4: Safe Harbor i przegląd katalogu (tryb ciągły)",
            "1. Lista zakazana w sklepie: eko, ekologiczny, przyjazny środowisku, zielony, naturalny jako synonim korzyści środowiskowej, biodegradowalny bez normy, «dbamy o środowisko» przypisane klientowi przez sam fakt zakupu gadżetu.\n"
            "2. Nowa karta produktu przechodzi przez tę listę przed publikacją. Warianty ilościowe dziedziczą tekst rodzica, więc korekta rodzica nie zwalnia z kontroli wariantu.\n"
            "3. Raz na kwartał przeszukać witrynę pod «eko», «środowisk», «biodegrad», «bambus» i «kraft». Strona główna i usługi cyfrowe są dziś czyste; katalog nie może ich zepsuć przy kolejnej dostawie opisów.",
            "ECFDF5", "059669", RGBColor(4, 120, 87),
        ),
    ]
    for title, body, bg, border, color in stages:
        callout(title, body.split("\n"), bg, border, color)

    add_h("7. Wnioski")
    add_p(
        "Witryna Hi DESIGN w części usług cyfrowych nie zawiera twierdzeń środowiskowych, natomiast katalog gadżetów i powiązane podstrony druku oraz bloga naruszają zakaz ogólnych haseł ekologicznych (ECGT pkt 4a) i zakaz przenoszenia jednej cechy na cały produkt oraz na markę klienta (pkt 4b). "
        "W obecnej formie te karty i odpowiedzi FAQ nie powinny pozostać opublikowane. Korekta polega na usunięciu przymiotnika środowiskowego i na zostawieniu surowca, gramatury i techniki druku, a nie na dopisaniu kolejnego hasła «eko» z odsyłaczem."
    )

    out = ROOT / "Raport_Audytu_Greenwashingu_Hi_DESIGN.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    path = build_report()
    print(path)
