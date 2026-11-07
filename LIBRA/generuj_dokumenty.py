#!/usr/bin/env python3
"""Materiał źródłowy i raport audytu greenwashingu: Libra Partners (wersja EN).

Skład kolejnych raportów: sklad_raportu.py i wzory/WZOR_raport.docx.
Ten generator zachowuje dotychczasowy skład raportu Libra do czasu przebudowy.
Werdykt według INSTRUKCJA_GENEROWANIA_RAPORTOW.md: do 27.09.2027 czarna lista
EmpCo jest przyszłym standardem. Social washing: knowledge/12_social_washing.md.
"""

from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
LOGO = ROOT.parent.parent / "ESG Hero Day - Materialy Druk" / "assets" / "esg_logo_white.png"
SHOT = ROOT / "zrodla" / "libra-en-hero-2026-10-30.png"


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


def build_material():
    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    NAVY = RGBColor(8, 24, 31)
    MUTED = RGBColor(90, 107, 119)
    DARK = RGBColor(22, 36, 44)

    def p(text, size=10.5, bold=False, italic=False, color=DARK, before=0, after=6):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(before)
        par.paragraph_format.space_after = Pt(after)
        par.paragraph_format.line_spacing = 1.12
        _set_run(par.add_run(text), size, bold, italic, color)
        return par

    def h(text):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(14)
        par.paragraph_format.space_after = Pt(4)
        par.paragraph_format.keep_with_next = True
        _set_run(par.add_run(text), 13, True, False, NAVY)

    p("LIBRA PARTNERS – Zestaw treści i deklaracji ze strony angielskiej", 16, True, False, NAVY, after=2)
    p(
        "Zbiór haseł, liczb, opinii i parametrów zebranych z wersji angielskiej witryny libra-partners.pl. "
        "Plik jest zapisem źródłowym do audytu. Nie zawiera oceny prawnej.",
        10.5, False, True, MUTED, after=8,
    )

    rows = [
        ("Podmiot", "LIBRA PARTNERS, marka w grupie SZARPOL. Kontakt na stronie: Rafał Szarzała, Centrum Produkcyjno-Handlowe „Szarpol”, ul. prof. Mariana Rapackiego 3, 98-200 Sieradz, tel. +48 501 989 969."),
        ("Założyciele podani na stronie", "Maciej Włodarczyk i Rafał Szarzała. Grupa SZARPOL założona w 1994 r. przez dr. Józefa Szarzałę, produkcja włóknin polipropylenowych."),
        ("Witryna badana", "https://libra-partners.pl/en/ oraz podstrony angielskie: /en/zrownowazony-rozwoj/ , /en/o-nas/ , /en/standard-pallet-model-mw01/ , /en/paleta-antybakteryjna-model-mw01-med/ , /en/offer/"),
        ("Data zebrania", "30 października 2026 r."),
        ("Zakres", "Wersja angielska. Wersja polska witryny istnieje (przełącznik języka) i nie jest tym zbiorem."),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = False
    for i, (k, v) in enumerate(rows):
        a, b = table.rows[i].cells
        a.width, b.width = Cm(4.4), Cm(13.0)
        pa, pb = a.paragraphs[0], b.paragraphs[0]
        pa.paragraph_format.space_after = Pt(1)
        pb.paragraph_format.space_after = Pt(1)
        _set_run(pa.add_run(k), 9, True)
        _set_run(pb.add_run(v), 9, False)

    h("1. Strona główna – https://libra-partners.pl/en/")
    p("Tytuł strony: „LIBRA PARTNERS New Generation Pallets - Eco-friendly pallets”.")
    p("Slajd: „WELCOME TO LIBRA PARTNERS”. „Carbon footprint”. „Check how much you can reduce the carbon footprint in your company by choosing ecological LIBRA PARTNERS pallets instead of wooden pallets”.")
    p("Na tym slajdzie, po prawej stronie kadru, widoczna jest okrągła pieczęć z literami „ESG” w środku, gwiazdkami i otokiem, którego czytelny fragment brzmi „ENVIRONMENTAL SOCIAL AND”.")
    p("Drugi slajd: „technology with ecology in mind”.")
    p("Cechy pod slajdem: „Pallets made of recycled HDPE”. „1200x800x144 mm - EUR pallet size”. „Our pallets weigh only 9kg”. „static load 3000 kg - dynamic load 1500 kg - load on racks 750 kg”. „Our pallets are 10 times more durable compared to wooden ones”. „Perfect for the food and pharmaceutical industries. They do not require fumigation - compliant with the ISPM 15 standard”.")
    p("Nagłówek: „100% CIRCULAR ECONOMY”.")
    p("„Because the raw material from which the pallets are made is 100% recyclable, we guarantee our customers that we can come and collect used and damaged pallets purchased from LIBRA PARTNERS. In this business model, there is no waste; a Closed Circuit Economy model is present. Likewise, as a socially responsible and environmentally conscious company, we want our production process to be distinguished by the highest standards in environmental management. We do this through a business model focused on ZERO WASTE.”")
    p("„Replace wooden pallets with ecological LIBRA PARTNERS pallets made of recycled HDPE material! Thanks to simulations and calculations carried out in accordance with the EN 16258:2012 standard for measuring the carbon footprint, we are able to indicate how much the carbon footprint / CO2 emission to the atmosphere will decrease using LIBRA PARTNERS pallets. The reduction of CO2 emissions to the environment with the use of 33 LIBRA PARTNERS pallets per month is 400 kg”.")
    p("Sekcja „Consumer feedback”. Przy wizytówce „Director of ESG, pharmaceutical industry” szablon karty pokazuje pięć wypełnionych gwiazdek. Tekst karty: „We express deep interest in purchasing pallets made from recycled raw materials. This will benefit our company's Sustainability policy.”")
    p("Karta „Head of the ESG department, food industry”, także z pięcioma wypełnionymi gwiazdkami: „We are considering changing our existing wooden pallets to LIBRA PARTNERS pallets. Thanks to the lighter weight, and the fact that they are recycled - we will be more ECO and it will have a very positive impact on the image of our company.”")
    p("„The production of ecological pallets is a brilliant alternative to standard wooden pallets. Wood is a very difficult material to recycle and has many disadvantages, making the possibilities of its use limited. For this reason we have decided to produce modern ecological pallets, which have a positive impact on our environment, the LCA report developed by Envirly on our pallets confirms almost twice lower CO2 emissions in the supply chain compared to standard wooden pallets.”")
    p("„Good quality industrial pallets made of ecological raw materials”. „Already at the stage of manufacturing pallets, they are more environmentally friendly, as they do not need to undergo specialized processing and drying.” „one piece of our eco pallet weighs only 9 kg.”")
    p("Misja: „LIBRA PARTNERS' mission is to become a key player in sustainability, a concept that will undoubtedly shape the company's guidelines in the coming generation.”")
    p("Paleta antybakteryjna na stronie głównej: „The LIBRA PARTNERS antibacterial pallet model MW01-MED received the opinion of the Poznań University of Technology for innovation on a global scale.”")

    h("2. Zrównoważony rozwój – https://libra-partners.pl/en/zrownowazony-rozwoj/")
    p("„Selected goals of the sustainable development policy, confirming the compliance of our project with the goals of the 2030 Agenda.” Dalej cytowane są cele 9.4, 12.2, 12.5 i 15.2 Agendy 2030.")
    p("„200+ tonnes of HDPE waste”. „Over 200 tons of HDPE waste – this is the amount of waste that is recycled monthly, from which our pallets are made.”")
    p("„303 t CO2 less per month”. „Thanks to our efforts in the production of recycled plastic pallets that weigh less than standard wooden pallets, our customers reduce CO2 emissions to the atmosphere by over 300 tons per month.”")
    p("„600.000 trees”. „So many trees are needed per month to absorb additional CO2 emissions in the case of using wooden pallets instead of LIBRA PARTNERS pallets.”")
    p("„Over 100.000 m3 of wood – this is how much wood is used monthly for the production of pallets in Poland. it is like a forest of about 110 hectares.”")
    p("Kalkulator, nagłówek „CASE STUDY” / „MONTHLY PALLET PURCHASE”, po zdaniu o 400 kg dla 33 palet: „500 PCS – 6,060 TONS LESS CO2”. „1.000 PCS – 12,120 TONS LESS CO2”. „2.000 PCS – 24,240 TONS LESS CO2”. „6.000 PCS – 60,600 TONS LESS CO2”. „10.000 PCS – 121,120 TONS LESS CO2”.")
    p("„Make your company compliant with the Sustainable Development and ESG policy. Prepare your company for non-financial ESG reporting now. Changing pallets from wooden to plastic LIBRA PARTNERS is the simplest solution to significantly reduce CO2 emissions and carbon footprint into the atmosphere.”")
    p("„By using recycled material instead of wood like standard wooden pallets, we reduce deforestation. […] Between 2010 and 2015, the world’s forest area decreased by 3.3 million hectares.”")
    p("„Thanks to the monthly production of recycled pallets, LIBRA PARTNERS reduces CO2 emissions by 303 tons per month. This is because the weight of the LIBRA PARTNERS plastic pallet is 9 kg instead of 25 kg. […] Assuming that one tree absorbs 0.5 kg of CO2 per month, it should be emphasized that if EUR wooden pallets are used instead of LIBRA PARTNERS plastic pallets, an additional 600,000 are needed.”")
    p("„According to research conducted by the Institute of Materials Engineering in Warsaw, the average life cycle of a wooden pallet is 6-7 cycles of loading/unloading, then such a pallet requires repair or is useless. […] the life cycle of a plastic pallet is about 60-70 loading / unloading cycles, and it should be emphasized that after this time the pallet is 100% recyclable.”")
    p("„as a socially responsible and ecologically aware company, we want our production process to also stand out with the highest standards in the field of environmental management. We do this through a business model focused on ZERO WASTE. […] There is no waste in this business model, there is a Circular Economy model.”")
    p("„In August 2023, an analysis of Libra Partners was published by Envirly – a leader in measuring the carbon footprint of products on our palette. We encourage you to read the results presented in this analysis.”")

    h("3. O nas – https://libra-partners.pl/en/o-nas/")
    p("„Eliminate end-of-life waste by recycling our solutions, which never end up in a landfill.”")
    p("„LIBRA PARTNERS, thanks to the production of pallets from recycled HDPE raw material, perfectly fits into the principles of Sustainable Development, Circularity and the 6R Principles.”")
    p("„The mission of LIBRA PARTNERS is to become a key player in the field of sustainable development […] by producing innovative transport pallets from recycled materials.”")

    h("4. Paleta standardowa MW01 – https://libra-partners.pl/en/standard-pallet-model-mw01/")
    p("„Manufactured from recycled HDPE raw material, under strict control of the ISO 9001 management system, they significantly reduce carbon footprint emissions in the organization and in the supply chain. Check how much you can reduce the carbon footprint in your company – by choosing ecological LIBRA PARTNERS pallets.”")
    p("Dane: 1200 × 800 × 144 mm, 9 kg, surowiec „Recycled HDPE”, obciążenia według ISO 8611: statyczne 3000 kg, dynamiczne 1500 kg, regałowe 750 kg.")

    h("5. Paleta antybakteryjna MW01-MED – https://libra-partners.pl/en/paleta-antybakteryjna-model-mw01-med/")
    p("„We are proud to announce that thanks to our R&D research, we are the world’s first manufacturer of pallets adapted to the pharmaceutical and medical industries, which have 99% antibacterial properties throughout the entire life cycle of the pallet. In addition, by using LIBRA PARTNERS plastic pallets, you contribute to reducing CO2 emissions to the environment and you are ESG Friendly!”")
    p("„Thanks to the addition of zinc oxide in the right proportion to the HDPE raw material during the production process, our MW01-MED pallet has very high antibacterial properties throughout the entire life cycle of the pallet. In connection with R&D research conducted by LIBRA PARTNERS in cooperation with a certified unit of the Łukasiewicz research network, the MW01-MED pallet model has the following antibacterial properties – tested according to ISO 22196:2011:”")
    p("„DECREASE OF ESCHERICHIA COLI SAMPLES ( ATCC 8739 ) ABOUT 99,97 %”. „DECREASE OF STAPHYLOCOCCUS AUREUS SAMPLES ( ATCC 6538P ) ABOUT 99,98 %”.")
    p("„The LIBRA PARTNERS antibacterial palette model MW01-MED received the opinion of the Poznań University of Technology for innovation on a global scale.”")
    p("Surowiec na tej podstronie: „Recycled HDPE”. Wymiary i obciążenia jak przy MW01. Waga 9 kg.")

    out = ROOT / "Material_Libra_Partners_EN.docx"
    doc.save(out)
    return out


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
        if LOGO.exists():
            lp.add_run().add_picture(str(LOGO), width=Cm(4.4))
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
        _set_run(fp.add_run("ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  ESG Institute Sp. z o.o.  ·  str. "), 8, False, False, MUTED)
        page_field(fp, MUTED)

    add_p("ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO", size=9, bold=True, color=ACCENT, after=2)
    add_p("Audyt Ryzyk Greenwashingu: LIBRA PARTNERS", size=20, bold=True, color=PRIMARY, before=2, after=2)
    add_p(
        "Weryfikacja angielskiej witryny https://libra-partners.pl/en/ oraz podstron zrównoważonego rozwoju, „o nas”, palety MW01 i palety antybakteryjnej MW01-MED, "
        "stan na 30 października 2026 r. Podstawa werdyktu: ustawa o zwalczaniu nieuczciwej konkurencji (reklama wobec klienta biznesowego) oraz, "
        "tam gdzie przekaz jest skierowany do konsumenta, art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. "
        "Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) wchodzi w życie 27 września 2027 r. Do tej daty czarna lista dyrektywy (UE) 2024/825 (EmpCo) "
        "jest przyszłym standardem, a nie stwierdzonym naruszeniem prawa polskiego. Social washing haseł o odpowiedzialności, Agendzie 2030 i opiniach oceniono tą samą miarą.",
        size=10.0, italic=True, color=MUTED, after=8,
    )

    if SHOT.exists():
        shot_par = doc.add_paragraph()
        shot_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shot_par.paragraph_format.space_before = Pt(2)
        shot_par.paragraph_format.space_after = Pt(2)
        shot_par.add_run().add_picture(str(SHOT), width=Cm(15.4))
        add_p(
            "Zrzut slajdu „Carbon footprint” na https://libra-partners.pl/en/ z 30 października 2026 r. Hasło wiodące porównuje palety LIBRA PARTNERS z paletami drewnianymi. Na pełnym slajdzie, na prawo od tego kadru, jest okrągła pieczęć z napisem ESG.",
            size=8.5, italic=True, color=MUTED, after=8,
        )

    meta = [
        (
            "Badany podmiot",
            "LIBRA PARTNERS, prezentowana na stronie jako część grupy SZARPOL z Sieradza. Kontakt: Rafał Szarzała, Centrum Produkcyjno-Handlowe „Szarpol”, ul. prof. Mariana Rapackiego 3, 98-200 Sieradz. Założyciele podani na stronie: Maciej Włodarczyk i Rafał Szarzała. Numer KRS i NIP nie występują w badanym materiale i nie są tu uzupełniane.",
        ),
        (
            "Przedmiot audytu",
            "Angielska strona główna oraz podstrony /en/zrownowazony-rozwoj/, /en/o-nas/, /en/standard-pallet-model-mw01/, /en/paleta-antybakteryjna-model-mw01-med/ i oferta antybakteryjna na /en/offer/. Wersja polska witryny istnieje i nie jest przedmiotem tego raportu.",
        ),
        (
            "Sektor działalności",
            "Produkcja palet transportowych z HDPE (w tym z recyklatu) dla logistyki, przemysłu spożywczego, farmaceutycznego, medycznego, elektronicznego i motoryzacyjnego. Oferta jest skierowana do firm.",
        ),
        (
            "Reżim prawny",
            "Prawo obowiązujące w dniu audytu: ustawa o zwalczaniu nieuczciwej konkurencji, tekst jednolity Dz.U. 2026 poz. 85, art. 16 i art. 18; ustawa o przeciwdziałaniu nieuczciwym praktykom rynkowym, tekst jednolity Dz.U. 2023 poz. 845, art. 5 i 6, gdy odbiorcą jest konsument; ustawa o ochronie konkurencji i konsumentów, tekst jednolity Dz.U. 2025 poz. 1714, art. 24 i art. 106 ust. 1 pkt 4. Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) została ogłoszona 19 października 2026 r. Art. 3: wchodzi w życie 27 września 2027 r. W dniu audytu czarna lista EmpCo nie jest prawem polskim.",
        ),
        ("Data audytu", "30 października 2026 r."),
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
            "Strona sprzedaje paletę firmie i w tym samym miejscu podaje ilość emisji oraz obietnicę zgodności z polityką ESG. Art. 16 ust. 1 pkt 2 i ust. 2 ustawy o zwalczaniu nieuczciwej konkurencji obejmują reklamę, która wprowadza klienta w błąd co do ilości i może wpłynąć na decyzję o nabyciu. Tabela „6,060 TONS” przy podstawie „400 kg” dla 33 palet jest taką ilością na wersji angielskiej. To ustalenie wystarcza do werdyktu czerwonego.",
            "Do tego dochodzą bezwzględne zdania o braku odpadów, o zgodności projektu z Agendą 2030, o „najwyższych standardach” firmy odpowiedzialnej społecznie oraz o palecie „w 99% antybakteryjnej przez cały cykl życia”, przy wyniku badania dwóch szczepów. Hasła o ludziach i o odpowiedzialności oceniono jako social washing tym samym testem reklamy.",
            "Ustawa wdrażająca dyrektywę (UE) 2024/825 wchodzi w życie 27 września 2027 r. Ogólne „eco-friendly”, „key player in sustainability” i pieczęć ESG są do tej daty przyszłym standardem pkt 2a i 4a, a claim o całym produkcie oparty na masie albo na odbiorze własnych palet – przyszłym standardem pkt 4b. Nie są dziś naruszeniem polskiej czarnej listy. Przekaz nie opiera redukcji na kompensacji, więc pkt 4c go nie opisuje.",
        ],
        "FFF1F2", "E11D48", DANGER,
    )

    add_h("1. Analiza Wstępna")
    add_p(
        "LIBRA PARTNERS przedstawia się jako producent palet z HDPE, w modelu grupy SZARPOL z Sieradza, i kieruje ofertę do firm logistycznych, spożywczych, farmaceutycznych i medycznych. "
        "Wersja angielska otwiera się hasłem o śladzie węglowym i zachętą, by firma wybrała „ecological” palety zamiast drewnianych. Niżej strona mówi o gospodarce o obiegu zamkniętym, o firmie odpowiedzialnej społecznie, o zgodności projektu z celami Agendy 2030 i o opiniach nazwanych „Consumer feedback”. "
        "Odbiorcą zdania „your company” jest klient biznesowy. Dla tej decyzji właściwą ustawą już dziś jest ustawa o zwalczaniu nieuczciwej konkurencji. Art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym zostają przy tych fragmentach, które strona sama adresuje do konsumenta, oraz jako miara publicznego przekazu, gdy czyta go konsument."
    )
    add_p(
        "Trzy grupy zdań spełniają przesłanki przepisu obowiązującego. Pierwsza to liczba: 33 palety mają dawać 400 kg CO2 miesięcznie według EN 16258:2012, a kalkulator dla 500 sztuk podaje „6,060 TONS”, po czym strona nazywa zmianę palet najprostszą drogą do zgodności z polityką ESG i do raportowania niefinansowego. "
        "Druga to obieg zamknięty i warstwa społeczna: „100% CIRCULAR ECONOMY”, „there is no waste”, „ZERO WASTE”, „highest standards” firmy „socially responsible” oraz „confirming the compliance” z Agendą 2030. "
        "Trzecia to paleta MW01-MED: wynik ISO 22196 dla dwóch szczepów, około 99,97% i 99,98%, jest na tej samej stronie zdaniem o „99%” przez „entire life cycle”, o „world’s first” i o tym, że nabywca „is ESG Friendly”."
    )
    add_p(
        "Porównanie z paletą drewnianą jest na stronie wprost, więc jest reklamą porównawczą, jeżeli pozwala rozpoznać ten towar. „Wooden pallets”, „standard wooden pallets” i „EUR wooden pallets” pozwalają rozpoznać kategorię konkurencyjnego nośnika 1200 × 800 mm. Samo zestawienie kategorii może zostać, gdy ilość, badanie i założenia da się sprawdzić. "
        "Hasła ogólne „eco-friendly”, „technology with ecology in mind” i „key player in sustainability” oraz pieczęć ESG na slajdzie śladu zbliżają się do zakazów, które w Polsce zaczną obowiązywać 27 września 2027 r. Same nie dają werdyktu czerwonego. Werdykt czerwony dają liczby i zdania bezwzględne opisane niżej. "
        "Kwalifikacja opiera się na tekście ustaw odczytanym w ISAP w dniu audytu. Raport nie powołuje orzeczenia: o wyniku rozstrzyga ten tekst zastosowany do cytatu ze strony. Wyrok nie byłby zresztą wyrokiem wobec LIBRA PARTNERS."
    )

    add_h("2. Analiza szczegółowa - tabela zbiorcza naruszeń")
    findings = [
        ("01", "Kalkulator CO2 „6,060 TONS” i obietnica zgodności z polityką ESG", "UZNK art. 16 ust. 1 pkt 2 i ust. 2 oraz ust. 3", "BARDZO WYSOKIE", "Ilość na stronie angielskiej i zdanie, że zmiana palet daje zgodność ESG oraz nadaje się do raportu"),
        ("02", "„There is no waste”, firma społecznie odpowiedzialna i „compliance” z Agendą 2030", "UZNK art. 16 ust. 1 pkt 2 i ust. 2", "WYSOKIE", "Social washing: brak odpadów, najwyższe standardy i zgodność z celami ONZ bez zakresu"),
        ("03", "Paleta „99% antybakteryjna przez cały cykl życia” i „you are ESG Friendly”", "UZNK art. 16 ust. 1 pkt 2 i ust. 2", "WYSOKIE", "Wynik dwóch szczepów według ISO 22196 rozszerzony na cały cykl życia, pierwszeństwo na świecie i cechę ESG nabywcy"),
        ("04", "„10 razy trwalsze” oraz wycinka i „drewno trudne do recyklingu”", "UZNK art. 16 ust. 1 pkt 2 i ust. 3 pkt 1–3 i 5", "ŚREDNIE", "Porównanie z paletą drewnianą bez badania, które da się odnaleźć, i bez przeliczenia hektarów"),
        ("05", "„Consumer feedback” z pięcioma gwiazdkami przy liście o zainteresowaniu zakupem", "UZNK art. 16 ust. 1 pkt 2 i ust. 2; UPNPR art. 5 ust. 1", "ŚREDNIE", "Social proof: gwiazdki przy wypowiedzi, która sama mówi o zainteresowaniu, a nie o użyciu"),
    ]
    sum_table = doc.add_table(rows=1 + len(findings), cols=5)
    sum_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_table.autofit = False
    sum_table.style = "Table Grid"
    col_w = (Cm(1.1), Cm(5.5), Cm(4.3), Cm(2.3), Cm(4.4))
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota naruszenia"]
    for c, htxt in enumerate(headers):
        cell = sum_table.rows[0].cells[c]
        cell.width = col_w[c]
        shade(cell, "08181F")
        margins(cell, 50, 50, 50, 50)
        par = cell.paragraphs[0]
        par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
        _set_run(par.add_run(htxt), 8, True, False, WHITE)
    risk_color = {
        "BARDZO WYSOKIE": DANGER,
        "WYSOKIE": DANGER,
        "ŚREDNIE": WARNING,
    }
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
                _set_run(run, 7.5, True, False, risk_color.get(val, DARK))
            elif c == 0:
                _set_run(run, 8, True, False, NAVY)
            else:
                _set_run(run, 7.5, False, False, DARK)
    add_p("", size=4, after=4)

    add_h("3. Szczegółowa analiza przypadków")

    cases = [
        {
            "title": "PRZYPADEK 01: Kalkulator emisji i zdanie, że zmiana palet daje zgodność z polityką ESG",
            "basis": "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2, ust. 2 i ust. 3 (Dz.U. 2026 poz. 85). Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4b, a dla hasła ogólnego i pieczęci ESG także pkt 4a i 2a ustawy z 4.09.2026 r. (Dz.U. 2026 poz. 1353) – jeszcze nie naruszenie krajowej czarnej listy.",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE",
            "quotes": [
                "„Check how much you can reduce the carbon footprint in your company by choosing ecological LIBRA PARTNERS pallets instead of wooden pallets.”",
                "„Thanks to simulations and calculations carried out in accordance with the EN 16258:2012 standard for measuring the carbon footprint […] The reduction of CO2 emissions to the environment with the use of 33 LIBRA PARTNERS pallets per month is 400 kg.”",
                "„500 PCS – 6,060 TONS LESS CO2”. „1.000 PCS – 12,120 TONS LESS CO2”. „10.000 PCS – 121,120 TONS LESS CO2.”",
                "„Make your company compliant with the Sustainable Development and ESG policy. Prepare your company for non-financial ESG reporting now. Changing pallets from wooden to plastic LIBRA PARTNERS is the simplest solution to significantly reduce CO2 emissions and carbon footprint into the atmosphere.”",
                "„Manufactured from recycled HDPE raw material, under strict control of the ISO 9001 management system, they significantly reduce carbon footprint emissions in the organization and in the supply chain.”",
                "„the LCA report developed by Envirly on our pallets confirms almost twice lower CO2 emissions in the supply chain compared to standard wooden pallets.”",
            ],
            "assessment": (
                "Art. 16 ust. 3 uznaje za reklamę porównawczą reklamę, która pozwala bezpośrednio albo pośrednio rozpoznać konkurenta albo towar konkurenta. "
                "Zdanie o śladzie stawia paletę LIBRA PARTNERS naprzeciw „wooden pallets”. Jest to kategoria nośnika o tym samym przeznaczeniu. Konkurentów w tej kategorii może być wielu. "
                "Art. 16 ust. 1 pkt 2 obejmuje reklamę, która wprowadza klienta w błąd i może przez to wpłynąć na decyzję o nabyciu. Ust. 2 każe przy tej ocenie uwzględnić wszystkie elementy, zwłaszcza ilość. "
                "Klientem zdania „your company” i „your organization” jest nabywca biznesowy. "
                "Na tej samej stronie 33 palety mają obniżać emisję o 400 kg miesięcznie. Przeliczenie liniowe na 500 palet daje 6,06 tony. Tabela wersji angielskiej zapisuje „6,060 TONS”, a dla 10 000 sztuk „121,120 TONS”. "
                "W tekście angielskim przecinek oddziela tysiące, a jednostką jest słowo TONS. Czytelnik tej wersji może odczytać 6 060 ton, czyli tysiąckrotność 6,06 tony. "
                "Te same cyfry, czytane polskim przecinkiem dziesiętnym, dają 6,060 tony i zgadzają się z ekstrapolacją 400 kg. Strona nie podaje, że przecinek jest znakiem dziesiętnym. "
                "Ilość, którą odbiera odbiorca wersji angielskiej, może zmienić decyzję o zakupie i o wpisaniu liczby do materiału dla klienta końcowego. "
                "Odczyt 6,06 tony też jest na stronie śladem węglowym i „the simplest solution” do zgodności z polityką ESG oraz do przygotowania raportowania niefinansowego. "
                "Strona nazywa EN 16258:2012 normą pomiaru tego śladu. Karta projektu CEN dla EN ISO 14083:2023 podaje, że norma ta zastępuje EN 16258:2012, a data wycofania normy sprzecznej to 31 października 2023 r. "
                "Zakres EN 16258 opisany przy tej karcie dotyczy energii i emisji usługi transportowej. Strona nie podaje trasy, ładunku, jednostki funkcjonalnej ani granic obliczenia. "
                "Obok stoją dwie dalsze wielkości bez uzgodnienia z kalkulatorem: „almost twice lower CO2” z analizy Envirly z sierpnia 2023 r., której granic ten akapit nie podaje, oraz „303 t CO2 less per month” jako efekt po stronie klientów. "
                "Kalkulator przypisuje redukcję każdemu nabywcy z osobna. Zdanie o ISO 9001 stoi w jednym zdaniu ze znaczną redukcją śladu. Strona nie podaje, żeby ten system obejmował pomiar emisji. "
                "Badania Envirly ten raport nie otwierał i nie uznaje go za fałszywe. Ocenia zdanie strony, które z badania robi niemal dwukrotnie niższy ślad łańcucha i zarazem zgodność firmy z polityką ESG. "
                "Art. 2 pkt 4 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym wiąże praktykę rynkową z promocją lub nabyciem produktu przez konsumenta, a art. 5 ust. 1 – z decyzją przeciętnego konsumenta dotyczącą umowy. "
                "Oferta palety na tej stronie jest ofertą dla firmy, więc podstawą spełnioną jest art. 16. Art. 5 zostaje przy publicznym czytaniu tego samego zdania przez konsumenta. "
                "Od 27 września 2027 r. twierdzenie o całym produkcie albo o całej działalności, które dotyczy masy, recyklatu albo jednego wyniku transportowego, trzeba będzie zestawić z art. 7 pkt 4b ustawy zmieniającej. "
                "Hasło „ecological” i tytuł „Eco-friendly pallets”, jeżeli na tym slajdzie nie będzie jasnego opisu, wejdą w pkt 4a. Pieczęć z „ESG” i czytelnym fragmentem otoku „ENVIRONMENTAL SOCIAL AND” jest dobrowolnym znakiem cech środowiskowych i społecznych. "
                "Od 27 września 2027 r. trzeba ją zestawić z pkt 2a, jeżeli nie opiera się na systemie certyfikacji ani na organie publicznym. Art. 3 ustawy zmieniającej odsuwa te punkty do 27 września 2027 r. "
                "Przekaz nie powołuje kompensacji emisji. Pkt 4c go nie opisuje. Dziś żaden z tych punktów nie jest stwierdzonym naruszeniem polskiej czarnej listy."
            ),
            "remediation": (
                "Zdjąć kalkulator i zdanie o zgodności z polityką ESG oraz o przygotowaniu do raportowania niefinansowego do czasu publikacji liczby z jednostką, granicami i metodą. "
                "Jeżeli firma zostawia porównanie transportowe, podać trasę albo tonokilometr, masę ładunku, masę palety drewnianej i jej źródło, granice (paliwo przewozu, a osobno produkcja palety) oraz normę aktualną dla usługi transportowej. "
                "Liczbę zapisać tak, by czytelnik angielski widział tony dziesiętne, na przykład „6.06 tonnes”, albo podać wynik w kilogramach. "
                "Wynik Envirly przywołać z datą, jednostką funkcjonalną i granicami albo usunąć „almost twice”. ISO 9001 zostawić jako system jakości, w zdaniu osobnym od emisji. "
                "Pieczęć ESG zdjąć albo opisać na tym samym slajdzie, kto i za co ją nadał."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„A LIBRA PARTNERS pallet in this comparison weighs 9 kg and is made from recycled HDPE. "
                "For a stated route, load and fuel, the difference in transport fuel emissions against a wooden pallet of a stated weight is [X] kg CO2e per trip, calculated under [method, year, boundaries]. "
                "This figure is not a company carbon footprint, not a CSRD disclosure and not a statement that the buyer’s ESG policy is met.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 02: Brak odpadów, odpowiedzialność społeczna i zgodność z Agendą 2030",
            "basis": "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2 i ust. 2. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a i 4b ustawy z 4.09.2026 r. – jeszcze nie naruszenie krajowej czarnej listy. Ścieżka Krajowego Punktu Kontaktowego OECD: poziom NISKIE, obok tego ustalenia, bez podnoszenia werdyktu.",
            "risk": "Poziom ryzyka: WYSOKIE",
            "quotes": [
                "„100% CIRCULAR ECONOMY.”",
                "„In this business model, there is no waste; a Closed Circuit Economy model is present. Likewise, as a socially responsible and environmentally conscious company, we want our production process to be distinguished by the highest standards in environmental management. We do this through a business model focused on ZERO WASTE.”",
                "„There is no waste in this business model, there is a Circular Economy model.”",
                "„Eliminate end-of-life waste by recycling our solutions, which never end up in a landfill.”",
                "„Selected goals of the sustainable development policy, confirming the compliance of our project with the goals of the 2030 Agenda.”",
            ],
            "assessment": (
                "To jest social washing w tym samym teście co twierdzenie środowiskowe. Hasło o firmie odpowiedzialnej społecznie, o najwyższych standardach zarządzania środowiskiem i o zgodności z celami ONZ mówi o ludziach, o etyce i o wpływie, a odbiorca może wziąć je za fakt szerszy niż mechanizm opisany obok. "
                "Mechanizm na stronie jest jeden: surowiec palety ma być w pełni przetwarzalny, a firma gwarantuje odbiór zużytych i uszkodzonych palet kupionych od LIBRA PARTNERS. "
                "Nagłówki „100% CIRCULAR ECONOMY”, „ZERO WASTE”, „there is no waste” i „never end up in a landfill” obejmują cały model i koniec życia każdej sztuki. "
                "Gwarancja odbioru obejmuje palety kupione od tej firmy i zwrócone. Zdanie „never” oraz „no waste” obejmuje także odpad produkcji, energię i palety, których klient nie odda. "
                "„Highest standards in environmental management” i „socially responsible” nie wskazują zakładu, okresu ani systemu. Strona nie powołuje w tym miejscu ISO 14001 ani innego systemu środowiskowego. "
                "„Confirming the compliance of our project with the goals of the 2030 Agenda” stawia cytat celów 9.4, 12.2, 12.5 i 15.2 jako potwierdzenie zgodności projektu. "
                "Cele te są rezolucją Zgromadzenia Ogólnego ONZ dla państw. Strona nie wskazuje oceny, która miałaby stwierdzać zgodność palety albo zakładu z tymi celami. Cel 15.2 w cytacie ma horyzont „By 2020”. "
                "Klient szukający dostawcy do polityki obiegu zamkniętego albo do odniesienia do celów zrównoważonego rozwoju może na tych zdaniach oprzeć wybór dostawcy. Art. 16 ust. 1 pkt 2 i ust. 2 obejmują taką reklamę, także co do przydatności towaru i zachowania reklamującego. "
                "Od 27 września 2027 r. ogólne „ZERO WASTE” i „100% CIRCULAR”, bez jasnego opisu na tym samym środku przekazu, trzeba będzie zestawić z pkt 4a, a claim o całej działalności oparty na odbiorze własnych palet – z pkt 4b. Dziś te punkty nie są prawem polskim. "
                "Firma nie pisze, że wdraża Wytyczne OECD. Samo hasło odpowiedzialności społecznej może jednak trafić do polskiego Krajowego Punktu Kontaktowego jako zawiadomienie o sposobie informowania. "
                "Postępowanie przed KPK nie jest sprawą sądową i nie kończy się wyrokiem. Skutkiem jest publiczna ocena, zalecenia i efekt wizerunkowy. Ten poziom jest NISKIE i nie podnosi werdyktu. Werdykt tego przypadku niesie art. 16."
            ),
            "remediation": (
                "Zastąpić „there is no waste”, „ZERO WASTE”, „100% CIRCULAR ECONOMY” i „never end up in a landfill” opisem odbioru palet kupionych od LIBRA PARTNERS oraz udziału recyklatu, z tym co dzieje się ze sztuką nieoddaną i z odpadem procesu. "
                "„Socially responsible” i „highest standards” zawęzić do nazwanego zakładu, roku i dokumentu albo usunąć. "
                "Przy Agendzie 2030 napisać, które działania firma odnosi do którego celu, i usunąć słowo „compliance”, dopóki nie ma oceny zewnętrznej. Cel 15.2 cytować razem z jego datą."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„LIBRA PARTNERS collects used and damaged pallets that the customer bought from LIBRA PARTNERS and recycles that returned material. "
                "This take-back does not describe all waste from the plant. The company refers its pallet design to UN Sustainable Development Goals 12.2 and 12.5 as a voluntary framework. "
                "That reference is not a certification and not a finding that the project complies with the 2030 Agenda.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 03: Właściwość antybakteryjna „przez cały cykl życia” oraz „ESG Friendly”",
            "basis": "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2 i ust. 2. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a dla „ESG Friendly” oraz pkt 4b, jeżeli cecha całej palety jest wnioskiem z badania powierzchni – jeszcze nie naruszenie krajowej czarnej listy.",
            "risk": "Poziom ryzyka: WYSOKIE",
            "quotes": [
                "„we are the world’s first manufacturer of pallets adapted to the pharmaceutical and medical industries, which have 99% antibacterial properties throughout the entire life cycle of the pallet. In addition, by using LIBRA PARTNERS plastic pallets, you contribute to reducing CO2 emissions to the environment and you are ESG Friendly!”",
                "„our MW01-MED pallet has very high antibacterial properties throughout the entire life cycle of the pallet. […] tested according to ISO 22196:2011: DECREASE OF ESCHERICHIA COLI SAMPLES ( ATCC 8739 ) ABOUT 99,97 %. DECREASE OF STAPHYLOCOCCUS AUREUS SAMPLES ( ATCC 6538P ) ABOUT 99,98 %.”",
                "„The LIBRA PARTNERS antibacterial palette model MW01-MED received the opinion of the Poznań University of Technology for innovation on a global scale.”",
            ],
            "assessment": (
                "Strona sama drukuje wynik: spadek liczebności próbek dwóch szczepów, około 99,97% i około 99,98%, według ISO 22196:2011, po współpracy z certyfikowaną jednostką sieci Łukasiewicz, przy dodatku tlenku cynku do HDPE. "
                "W tym samym widoku pisze, że paleta ma „99% antibacterial properties throughout the entire life cycle”, że producent jest „the world’s first”, i że używając palety „you are ESG Friendly” oraz ograniczasz emisje CO2. "
                "Wynik dwóch szczepów, zapisany jako „about”, nie jest zdaniem o całym cyklu życia palety w magazynie, po myciu i po ścieraniu. Strona nie podaje czasu kontaktu, liczby próbek ani tego, że ten sam spadek utrzymuje się po latach użytkowania. "
                "„99%” jest zaokrągleniem dwóch różnych wyników i obejmuje już „antibacterial properties” palety jako takiej. "
                "„World’s first” jest zdaniem o pierwszeństwie na świecie. Strona nie wskazuje przeglądu rynku. Raport nie stwierdza, że pierwszeństwo jest fałszywe. Stwierdza, że reklama podaje je jako fakt bez źródła, które klient może sprawdzić. "
                "Opinia Politechniki Poznańskiej jest na stronie opinią o innowacyjności na skalę świata. Strona nie przytacza tej opinii. Postawienie jej w jednym bloku z cyklem życia i ze zdaniem „ESG Friendly” pozwala czytać opinię o innowacji jako potwierdzenie tych dwóch zdań. "
                "Nabywcą jest firma farmaceutyczna albo medyczna. Zdanie o cyklu życia i o „ESG Friendly” może zdecydować o wyborze nośnika do strefy czystej. Art. 16 ust. 1 pkt 2 i ust. 2 obejmują jakość, przydatność i sposób wykonania. "
                "Raportu z badania ISO 22196 i opinii politechniki nie otwierano. Nie ma podstawy, by uznać liczby 99,97% i 99,98% za nieprawdziwe. Podstawą ustalenia jest różnica między tymi liczbami a zdaniami o cyklu życia, o pierwszeństwie i o cesze ESG nabywcy. "
                "Tekstu rozporządzenia w sprawie produktów biobójczych w tej sesji nie otwierano, więc raport nie stwierdza braku pozwolenia na substancję czynną. "
                "Od 27 września 2027 r. „ESG Friendly” bez jasnego opisu na tym samym środku przekazu trzeba będzie zestawić z pkt 4a, a wniosek o całej palecie wyprowadzony z badania powierzchni – z pkt 4b. Dziś to przyszły standard."
            ),
            "remediation": (
                "Zostawić nazwę szczepów, normę, rząd wielkości spadku i jednostkę, która badała. Dopisać warunki badania, które są w raporcie z badania. "
                "Usunąć „throughout the entire life cycle”, „99%” jako cechę całej palety, „world’s first” i „you are ESG Friendly”, dopóki osobne dowody nie obejmują cyklu życia, przeglądu rynku i cechy ESG. "
                "Opinię politechniki cytować jako opinię o innowacyjności, w akapicie osobnym od wyniku mikrobiologicznego i od emisji."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„MW01-MED is recycled HDPE with a zinc oxide addition. In a test according to ISO 22196:2011, run with a certified Łukasiewicz unit, samples showed a reduction of about 99.97% for E. coli ATCC 8739 and about 99.98% for S. aureus ATCC 6538P, under the contact conditions stated in the test report. "
                "The page does not extend this result to the whole service life of the pallet. Poznań University of Technology issued an opinion on the innovation. That opinion is not an ESG certification.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 04: Trwałość „10 razy” oraz zdania o wycince i o recyklingu drewna",
            "basis": "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2 oraz ust. 3 pkt 1–3 i 5. Porównanie może zostać po doprecyzowaniu. Przyszły standard EmpCo od 27.09.2027 nie zastępuje tego testu.",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "„Our pallets are 10 times more durable compared to wooden ones.”",
                "„According to research conducted by the Institute of Materials Engineering in Warsaw, the average life cycle of a wooden pallet is 6-7 cycles of loading/unloading […]. the life cycle of a plastic pallet is about 60-70 loading / unloading cycles.”",
                "„Wood is a very difficult material to recycle and has many disadvantages […]. modern ecological pallets, which have a positive impact on our environment.”",
                "„By using recycled material instead of wood like standard wooden pallets, we reduce deforestation. […] over 100.000 m3 of wood – that’s how much wood is obtained monthly in Poland for the production of wooden pallets. it’s like a forest of about 110 hectares. Between 2010 and 2015, the world’s forest area decreased by 3.3 million hectares.”",
                "„the weight of the LIBRA PARTNERS plastic pallet is 9 kg instead of 25 kg.”",
            ],
            "assessment": (
                "Art. 16 ust. 3 pkt 2 i 3 wymaga, żeby porównanie towarów o tym samym przeznaczeniu było rzetelne, dało się zweryfikować na podstawie obiektywnych kryteriów i obiektywnie zestawiało istotną, sprawdzalną i typową cechę. Pkt 5 wymaga, żeby reklama nie dyskredytowała towaru konkurenta. Pkt 1 odsyła do zakazu reklamy wprowadzającej w błąd. "
                "„10 times more durable” jest na stronie głównej cechą produktu, bez nazwy badania. Nazwa „Institute of Materials Engineering in Warsaw” pada na podstronie zrównoważonego rozwoju, bez tytułu, roku i odesłania. Klient nie może sprawdzić, czy „cykl” oznacza jeden obrót, naprawę, czy koniec życia, i jakiej palety drewnianej dotyczy 6–7 cykli. Stosunek 60–70 do 6–7 jest źródłem liczby „10 razy” na stronie. Bez badania liczba nie spełnia pkt 2 i 3. "
                "Masa 9 kg wobec 25 kg jest cechą, którą da się porównać, gdy strona poda, jakiej palety drewnianej dotyczy 25 kg. Tę parę liczb raport zostawia jako porównanie do utrzymania po dopisku źródła. "
                "„Wood is a very difficult material to recycle” oraz połączenie wyboru palety z ubytkiem 3,3 mln hektarów lasów w latach 2010–2015 i z „lasem około 110 hektarów” ze 100 000 m3 drewna wykracza poza sprawdzalną cechę nośnika. "
                "Strona nie podaje, ile metrów sześciennych przypada na hektar i czy chodzi o roczny przyrost, czy o zręb. Globalny ubytek lasów w cytowanych latach nie jest na stronie powiązany z produkcją palet w Polsce inaczej niż przez sąsiedztwo akapitów. "
                "Takie zestawienie może zniechęcić do palety drewnianej jako do towaru szkodliwego dla lasów i trudnego w recyklingu. To jest dyskredytacja z pkt 5, jeżeli zdania zostają w tej formie. "
                "Porównanie nie jest przez to bezwzględnie niedopuszczalne. Po podaniu badania, definicji cyklu, źródła 25 kg i przeliczenia hektarów, oraz po usunięciu zdania o globalnym ubytku lasów jako argumencie sprzedażowym, porównanie masy i trwałości może zostać. "
                "„Positive impact on our environment” przy całej palecie, oparte na tym porównaniu, od 27 września 2027 r. trzeba będzie zawęzić do pkt 4b, jeżeli dotyczy jednego aspektu. Dziś o wyniku decyduje art. 16."
            ),
            "remediation": (
                "Na stronie głównej albo zdjąć „10 times”, albo w tym samym bloku podać tytuł, rok i jednostkę badania oraz definicję cyklu. "
                "Zostawić 9 kg i 25 kg z dopiskiem, jaka paleta drewniana waży 25 kg. "
                "Usunąć „very difficult material to recycle” albo zastąpić je cechą, którą klient sprawdzi. "
                "Przeliczenie 100 000 m3 na 110 hektarów podać z założeniem albo usunąć. Ubytek 3,3 mln hektarów w latach 2010–2015 odłączyć od oferty palety."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„In [title of the study, institution, year], a wooden pallet in the tested duty reached 6–7 loading cycles before repair or withdrawal, and the plastic pallet reached about 60–70 cycles. "
                "The comparison uses that duty cycle. A LIBRA PARTNERS pallet of this model weighs 9 kg; the wooden pallet in the same comparison weighs 25 kg [source]. "
                "Recyclability of each material depends on the local collection route.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 05: „Consumer feedback” i pięć gwiazdek przy deklaracji zainteresowania",
            "basis": "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2 i ust. 2. Dodatkowo ustawa o przeciwdziałaniu nieuczciwym praktykom rynkowym, art. 5 ust. 1 i ust. 2 pkt 2, bo sekcja jest nazwana opinią konsumencką. Art. 7 pkt 25 tej ustawy w tym materiale nie jest spełniony.",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "Nagłówek sekcji: „Consumer feedback.”",
                "Karta „Director of ESG, pharmaceutical industry”, z pięcioma wypełnionymi gwiazdkami w szablonie: „We express deep interest in purchasing pallets made from recycled raw materials. This will benefit our company's Sustainability policy.”",
                "Karta „Head of the ESG department, food industry”, z pięcioma wypełnionymi gwiazdkami: „We are considering changing our existing wooden pallets to LIBRA PARTNERS pallets. Thanks to the lighter weight, and the fact that they are recycled - we will be more ECO and it will have a very positive impact on the image of our company.”",
            ],
            "assessment": (
                "Sekcja nazywa się opinią konsumentów. Szablon karty stawia pięć wypełnionych gwiazdek przy funkcji, a tekst karty jest deklaracją zainteresowania zakupem albo rozważeniem zmiany palet. "
                "Gwiazdki są elementem reklamy dotyczącym zachowania klienta. Czytelnik może wziąć je za ocenę 5/5 wystawioną po użyciu palety przez dyrektora ESG. Tekst tej samej karty mówi o zainteresowaniu, więc nie stwierdza, że autor już kupił albo używał. "
                "Art. 7 pkt 25 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym dotyczy przedsiębiorcy, który umożliwia dostęp do opinii konsumentów i twierdzi, że pochodzą one od osób, które używały produktu albo go nabyły, choć nie sprawdził tego w sposób uzasadniony i proporcjonalny. "
                "Cytowane karty mówią o zamiarze zakupu. Tego twierdzenia o wcześniejszym użyciu na nich nie ma, więc pkt 25 nie jest tu spełniony. Pkt 26, o nieprawdziwych opiniach, też nie: raport nie ustala, że wypowiedzi są zmyślone. "
                "Zostaje art. 16 ust. 1 pkt 2 i ust. 2 oraz art. 5 ust. 1 ustawy o nieuczciwych praktykach, bo nazwa sekcji jest konsumencka, a gwiazdki mogą skłonić do traktowania karty jak oceny po zakupie. "
                "Ta sama karta zawiera słowa o zainteresowaniu, więc mylne wrażenie da się zdjąć zmianą nagłówka i usunięciem gwiazdek. Stąd poziom średni. "
                "Zdanie „we will be more ECO” w drugiej karcie jest przyszłym ogólnym hasłem środowiskowym rozmówcy, nie faktem o palecie. Od 27 września 2027 r. takie ogólne „ECO” bez opisu trzeba będzie zestawić z pkt 4a, jeżeli zostanie w reklamie LIBRA PARTNERS jako jej własne twierdzenie. Dziś wystarczy zdjąć je razem z gwiazdkami albo opatrzyć zakresem."
            ),
            "remediation": (
                "Usunąć gwiazdki. Zmienić nagłówek na deklaracje zainteresowania, z datą pisma i informacją, czy doszło później do zakupu. "
                "Jeżeli firma nie ma źródła wypowiedzi, sekcję zdjąć. Nie przedstawiać listu intencyjnego jako opinii osoby, która używała palety. "
                "Zwrot „more ECO” w cudzym cytacie albo usunąć, albo zostawić wyłącznie jako cytat z wyraźnym statusem listu, bez własnego wniosku LIBRA PARTNERS, że klient staje się przez zakup bardziej ekologiczny."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Letters of interest, not post-purchase reviews. In [month year] an ESG director in the pharmaceutical industry wrote that the company was interested in buying pallets made from recycled material. "
                "LIBRA PARTNERS does not present this letter as a rating and does not state that a purchase followed.”"
            ),
            "salvage": True,
        },
    ]

    for case in cases:
        add_p(case["title"], size=12, bold=True, color=NAVY, before=12, after=2)
        add_p(case["basis"], size=9, bold=False, color=ACCENT, before=0, after=1)
        add_p(case["risk"], size=9, bold=True, color=DANGER if "BARDZO" in case["risk"] or "WYSOKIE" in case["risk"] else WARNING, before=0, after=4)
        add_p("Cytat źródłowy:", size=10, bold=True, before=2, after=2)
        for q in case["quotes"]:
            add_p(q, size=10, italic=True, before=0, after=3)
        add_p("Ocena prawna:", size=10, bold=True, before=4, after=2)
        add_p(case["assessment"])
        add_p("Kierunek naprawy:", size=10, bold=True, before=2, after=2)
        add_p(case["remediation"])
        if case["salvage"]:
            callout(case["claim_title"], [case["claim"]], "F0FDF4", "16A34A", RGBColor(22, 101, 52))
        else:
            add_p(case["claim"], size=10, italic=True)

    add_h("4. Porównanie twierdzeń LIBRA PARTNERS z rzeczywistym stanem prawno-środowiskowym")
    add_p(
        "Strona porównuje własną paletę z paletą drewnianą i z paletą EUR. Jest to reklama porównawcza w rozumieniu art. 16 ust. 3, bo towar konkurenta da się rozpoznać jako kategorię. "
        "Porównanie może zostać, gdy będzie rzetelne i sprawdzalne. Poniższa tabela nie żąda usunięcia zestawienia. Pokazuje, który element wymaga doprecyzowania, zanim reklama spełni pkt 1–3 i 5."
    )
    comp = [
        (
            "Ilość CO2. Twierdzenie: 33 palety obniżają emisję o 400 kg miesięcznie według EN 16258:2012; 500 sztuk to „6,060 TONS LESS CO2”; zmiana palet jest najprostszą drogą do zgodności z polityką ESG i do raportowania niefinansowego.",
            "400 kg na 33 palety daje liniowo około 6,06 tony na 500 sztuk. Zapis angielski „6,060 TONS” może być odczytany jako 6 060 ton. EN 16258:2012 została zastąpiona przez EN ISO 14083:2023 (karta CEN, wycofanie normy sprzecznej 31.10.2023) i dotyczyła usługi transportowej. Strona nie podaje trasy, ładunku ani granic. Zgodność polityki ESG i gotowość raportu niefinansowego z tej liczby nie wynikają.",
        ),
        (
            "Trwałość. Twierdzenie: paleta jest 10 razy trwalsza od drewnianej. Na podstronie: Instytut Materiałoznawstwa w Warszawie, 6–7 cykli wobec 60–70.",
            "Nazwa instytucji bez tytułu, roku i definicji cyklu nie pozwala sprawdzić porównania. Stosunek widełek jest źródłem „10 razy”. Po publikacji badania i definicji cyklu porównanie może zostać.",
        ),
        (
            "Masa. Twierdzenie: 9 kg zamiast 25 kg, a przez to niższe spalanie i niższy ślad.",
            "Różnica masy jest cechą, którą da się porównać. Brakuje wskazania, jakiej palety drewnianej dotyczy 25 kg, oraz tego, że niższe spalanie dotyczy przewozu, a nie całego śladu firmy.",
        ),
        (
            "Lasy i recykling drewna. Twierdzenie: drewno jest bardzo trudne w recyklingu; wybór recyklatu ogranicza wycinkę; 100 000 m3 miesięcznie to las około 110 ha; w latach 2010–2015 lasy świata zmniejszyły się o 3,3 mln ha.",
            "Przeliczenie hektarów nie ma na stronie założenia (m3/ha, zręb albo przyrost). Globalny ubytek lasów z lat 2010–2015 nie jest wykazany jako skutek polskiej palety drewnianej. Zdanie o trudnym recyklingu drewna dyskredytuje towar konkurenta, dopóki nie wskaże cechy sprawdzalnej.",
        ),
        (
            "Odpady i cel społecznej odpowiedzialności. Twierdzenie: w modelu nie ma odpadów, paleta nigdy nie trafia na składowisko, projekt jest zgodny z Agendą 2030, firma ma najwyższe standardy i jest odpowiedzialna społecznie.",
            "Na stronie jest gwarancja odbioru palet kupionych od LIBRA PARTNERS i twierdzenie o przetwarzalności surowca. To nie opisuje całego odpadu zakładu ani zgodności z rezolucją ONZ. „Compliance” sugeruje ocenę, której strona nie wskazuje.",
        ),
        (
            "Cecha antybakteryjna. Twierdzenie: około 99,97% i 99,98% dla dwóch szczepów według ISO 22196 oraz 99% przez cały cykl życia, pierwszy producent na świecie, nabywca jest „ESG Friendly”.",
            "Liczby dwóch szczepów strona podaje jako wynik badania. Cykl życia, pierwszeństwo na świecie i cecha ESG nabywcy są zdaniami szerszymi niż ten wynik. Opinia Politechniki Poznańskiej jest na stronie opinią o innowacyjności.",
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
            "1. Roszczenia konkurenta z ustawy o zwalczaniu nieuczciwej konkurencji",
            "Reklama porównawcza palety wobec palety drewnianej i zdania o ilości emisji, o braku odpadów oraz o całym cyklu życia właściwości antybakteryjnej spełniają art. 16. Przedsiębiorca, którego interes został zagrożony albo naruszony, może według art. 18 ust. 1 żądać zaniechania, usunięcia skutków, oświadczenia, naprawienia szkody, wydania bezpodstawnie uzyskanych korzyści, a przy zawinieniu także świadczenia na cel społeczny. To jest bezpośrednia droga dla producenta palet drewnianych albo innego nośnika z tej samej kategorii.",
        ),
        (
            "2. Postępowanie Prezesa UOKiK, jeżeli ten sam publiczny przekaz uzna się za praktykę wobec konsumentów",
            "Art. 24 ust. 1 i ust. 2 pkt 2 i 3 ustawy o ochronie konkurencji i konsumentów zakazuje praktyk naruszających zbiorowe interesy konsumentów, w tym naruszenia obowiązku informacji rzetelnej, prawdziwej i pełnej oraz czynu nieuczciwej konkurencji. Art. 106 ust. 1 pkt 4 pozwala nałożyć karę do 10% obrotu z roku poprzedzającego nałożenie kary, także przy naruszeniu nieumyślnym. Oferta palety jest na stronie skierowana do firmy, więc kara nie jest automatyczna. Ścieżka jest otwarta, bo witryna jest publiczna, a sekcja opinii jest nazwana konsumencką. Art. 106a tej ustawy, z karą dla osoby zarządzającej do 2 000 000 zł, dotyczy naruszeń zakazu porozumień ograniczających konkurencję i tego ustalenia nie obejmuje.",
        ),
        (
            "3. Standard EmpCo od 27 września 2027 r.",
            "Od tej daty, zgodnie z art. 3 ustawy z 4 września 2026 r., wejdą pkt 2a, 4a i 4b. Pieczęć ESG bez systemu certyfikacji albo organu publicznego, ogólne „eco-friendly”, „ecology in mind”, „key player in sustainability” i „ESG Friendly” oraz claim o całym produkcie albo całej firmie oparty na masie, recyklacie albo odbiorze własnych palet będą wtedy wymagały zawężenia. Pkt 4c, o wpływie na emisje uzasadnianym kompensacją, tego materiału nie opisuje: strona nie powołuje offsetu. Do 26 września 2027 r. te punkty nie są naruszeniem polskiej czarnej listy i same nie dają werdyktu czerwonego.",
        ),
        (
            "4. Social washing, raport klienta i Krajowy Punkt Kontaktowy OECD",
            "Klient, który wpisze kalkulator albo zdanie o zgodności z polityką ESG do własnego raportu niefinansowego, przenosi ilość ze strony do swojego ujawnienia. Strona wprost zachęca, by tak przygotować firmę do raportowania. LIBRA PARTNERS nie twierdzi na badanych podstronach, że sama podlega CSRD. Osobno: hasło firmy odpowiedzialnej społecznie i „najwyższych standardów”, choć nie powołuje Wytycznych OECD, może być przedmiotem zawiadomienia do polskiego KPK. To postępowanie nie jest sprawą sądową i nie kończy się wyrokiem. Skutkiem bywa publiczna ocena i zalecenia. Poziom tego ryzyka jest NISKIE i nie przesądza werdyktu.",
        ),
    ]
    for title, body in risks:
        add_p("• " + title, size=10.5, bold=True, before=8, after=2)
        add_p(body)

    add_h("6. Rekomendacje naprawcze")
    stages = [
        (
            "Etap 1. Natychmiast, do 13 listopada 2026 r. – tylko prawo już naruszone",
            "1. Zdjąć ze strony angielskiej kalkulator „TONS LESS CO2”, liczbę 400 kg bez trasy i zdanie, że zmiana palet daje zgodność z polityką ESG oraz przygotowuje do raportowania niefinansowego.\n"
            "2. Zdjąć „there is no waste”, „ZERO WASTE”, „100% CIRCULAR ECONOMY”, „never end up in a landfill”, „highest standards” i „confirming the compliance” z Agendą 2030 albo zastąpić je opisem odbioru własnych palet, bez słowa „compliance”.\n"
            "3. Na stronie MW01-MED zdjąć „entire life cycle”, „99%” jako cechę całej palety, „world’s first” i „you are ESG Friendly”. Zostawić wynik dwóch szczepów z nazwą normy.\n"
            "4. Zdjąć pięć gwiazdek i nagłówek „Consumer feedback” przy listach o zainteresowaniu zakupem.",
            "FFF1F2", "E11D48", DANGER,
        ),
        (
            "Etap 2. Przed 27 września 2027 r. – przyszły standard EmpCo",
            "1. Hasła „Eco-friendly pallets”, „technology with ecology in mind”, „key player in sustainability” i „ecological” albo opatrzyć na tym samym slajdzie jasnym opisem (recyklat, masa, odbiór), albo zastąpić opisem czynności.\n"
            "2. Pieczęć ESG na slajdzie śladu usunąć albo powiązać z nazwanym systemem certyfikacji lub z organem publicznym. Własny znak bez takiej podstawy od 27 września 2027 r. wejdzie w pkt 2a.\n"
            "3. Claim o całej palecie i o całej firmie zawęzić do aspektu, którego dotyczy dowód: masa w przewozie, udział recyklatu, odbiór sztuk kupionych od LIBRA PARTNERS. To jest przygotowanie do pkt 4b.\n"
            "4. Nie budować zdania o neutralności ani o pozytywnym wpływie na emisje na kompensacji. Tego zdania na stronie dziś nie ma i nie należy go dodawać.",
            "FFFBEB", "D97706", WARNING,
        ),
        (
            "Etap 3. Metodyka i dowody, do 29 stycznia 2027 r.",
            "1. Jeżeli porównanie transportowe wraca, oprzeć je na metodzie aktualnej dla usługi transportowej, z trasą albo tonokilometrem, ładunkiem, masą palety odniesienia i granicami. Wynik Envirly opublikować z jednostką funkcjonalną albo nie przywoływać „almost twice”.\n"
            "2. Badanie trwałości podać tytułem, rokiem, definicją cyklu i paletą odniesienia. Przeliczenie 100 000 m3 na hektary podać z założeniem albo usunąć.\n"
            "3. Przy MW01-MED dołączyć warunki badania ISO 22196 zgodne z raportem laboratorium. Opinię Politechniki Poznańskiej cytować w zakresie, w jakim została wydana.\n"
            "4. Sprawdzić, czy te same zdania są na wersji polskiej, i wprowadzić tę samą korektę. Ten audyt objął wersję angielską.",
            "F0F9FF", "0284C7", ACCENT,
        ),
        (
            "Etap 4. Sformułowania dopuszczalne i szkolenie, do 27 lutego 2027 r., potem na stałe",
            "1. Wprowadzić do strony wzory z przypadków 01–05. Wersja angielska i polska mają nieść ten sam zakres.\n"
            "2. Przeszkolić osoby prowadzące stronę i targi: ilość bez granic, „compliance”, „zero waste”, „ESG Friendly” i gwiazdki przy liście intencyjnym nie wracają do publikacji.\n"
            "3. Przy haśle o odpowiedzialności społecznej zapisać zakład, rok i dokument. Hasło bez tego zakresu zostaje ryzykiem wizerunkowym także wobec KPK OECD, na poziomie NISKIE.\n"
            "4. Raz na kwartał porównać stronę z kartą badania i z tekstem art. 16. Po 27 września 2027 r. ten przegląd obejmuje także pkt 2a, 4a i 4b.",
            "ECFDF5", "059669", RGBColor(4, 120, 87),
        ),
    ]
    for title, body, bg, border, color in stages:
        callout(title, body.split("\n"), bg, border, color)

    add_h("7. Wnioski")
    add_p(
        "W dniu 30 października 2026 r. angielska witryna LIBRA PARTNERS nie powinna zostać w obecnej formie, bo art. 16 ustawy o zwalczaniu nieuczciwej konkurencji obejmuje kalkulator emisji zapisany jako „TONS”, obietnicę zgodności z polityką ESG, zdania o braku odpadów i o zgodności z Agendą 2030 oraz rozszerzenie wyniku dwóch szczepów na cały cykl życia palety. "
        "Od 27 września 2027 r. dojdzie do tego standard pkt 2a, 4a i 4b ustawy z 4 września 2026 r.; do tej daty nie jest to naruszenie polskiej czarnej listy, a pkt 4c tego przekazu nie dotyczy."
    )

    out = ROOT / "Raport_Audytu_Greenwashingu_Libra_Partners.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build_material())
    print(build_report())
