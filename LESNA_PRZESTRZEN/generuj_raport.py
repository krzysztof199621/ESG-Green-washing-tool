#!/usr/bin/env python3
"""Audyt greenwashingu i social washingu: Fundacja Leśna Przestrzeń.

Skład jak sklad_raportu.py i wzory/WZOR_raport.docx.
Forma według INSTRUKCJA_GENEROWANIA_RAPORTOW.md.
Stan witryny i prawa: 7 listopada 2026 r.
"""

from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
LOGO = (
    ROOT.parents[1]
    / "ESG Experts Day - Komplet Materialow (PDF + PPTX)"
    / "03_DLA_PROWADZACEGO"
    / "assets"
    / "esg_logo_white.png"
)
HERO = ROOT / "assets" / "lesna_hero.png"
REPORT = ROOT / "Raport_Audytu_Greenwashingu_Fundacja_Lesna_Przestrzen.docx"
MATERIAL = ROOT / "Material_Fundacja_Lesna_Przestrzen_tresci_do_audytu.docx"


def set_run(run, size, bold=False, italic=False, color=None, font="Arial"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), font)
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "pl-PL")
    lang.set(qn("w:eastAsia"), "pl-PL")


def shade(cell, hex_color):
    cell._tc.get_or_add_tcPr().append(
        docx.oxml.parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    )


def margins(cell, top=80, bottom=80, left=90, right=90):
    cell._tc.get_or_add_tcPr().append(
        docx.oxml.parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f"</w:tcMar>"
        )
    )


def borders(cell, color="E2E8F0", sz="4", left_sz=None, left_color=None):
    lc = left_color or color
    ls = left_sz or sz
    cell._tc.get_or_add_tcPr().append(
        docx.oxml.parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="{ls}" w:space="0" w:color="{lc}"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f"</w:tcBorders>"
        )
    )


def field(paragraph, instruction, size, color, bold=False, placeholder="1"):
    begin_run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    begin_run._r.append(begin)
    begin_run._r.append(instr)
    begin_run._r.append(separate)
    set_run(begin_run, size, bold, False, color)
    result = paragraph.add_run(placeholder)
    set_run(result, size, bold, False, color)
    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)
    set_run(end_run, size, bold, False, color)


def set_table_width(table, width, col_widths=None, center=False):
    tbl = table._tbl
    tblPr = tbl.tblPr
    twips = str(int(width.twips))
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), twips)
    tblW.set(qn("w:type"), "dxa")
    existing = tblPr.find(qn("w:tblLayout"))
    if existing is not None:
        tblPr.remove(existing)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    if center:
        jc = tblPr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            tblPr.append(jc)
        jc.set(qn("w:val"), "center")
    if not col_widths:
        return
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for col in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(col.twips)))
        grid.append(gc)
    seen = set()
    for row in table.rows:
        for cell, col in zip(row.cells, col_widths):
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(int(col.twips)))
            tcW.set(qn("w:type"), "dxa")


def enable_hyphenation(doc):
    settings = doc.settings.element
    for tag, val in (
        ("w:autoHyphenation", "true"),
        ("w:consecutiveHyphenLimit", "3"),
        ("w:hyphenationZone", "360"),
        ("w:doNotHyphenateCaps", "true"),
        ("w:updateFields", "true"),
    ):
        el = OxmlElement(tag)
        el.set(qn("w:val"), val)
        settings.append(el)


def suppress_hyphen(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    el = OxmlElement("w:suppressAutoHyphens")
    el.set(qn("w:val"), "true")
    pPr.append(el)


def setup(doc):
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)
        section.header_distance = Cm(0.45)
        section.footer_distance = Cm(0.4)
    return doc


def build_report():
    doc = setup(docx.Document())
    enable_hyphenation(doc)
    for section in doc.sections:
        section.top_margin = Cm(2.3)
        section.bottom_margin = Cm(1.85)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        section.header_distance = Cm(0.32)
        section.footer_distance = Cm(0.28)

    NAVY = RGBColor(8, 24, 31)
    DARK = RGBColor(22, 36, 44)
    MUTED = RGBColor(90, 107, 119)
    BURGUNDY = RGBColor(142, 47, 57)
    WHITE = RGBColor(255, 255, 255)
    PALE = RGBColor(186, 214, 224)
    CONTENT = Cm(17.4)
    doc.core_properties.title = "Audyt Ryzyk Greenwashingu: Fundacja Leśna Przestrzeń"
    doc.core_properties.author = "ESG Institute"
    doc.core_properties.subject = "Sygn. ESG/2026/11/07"
    doc.core_properties.category = "Raport audytu"
    doc.core_properties.comments = "Poufne. Wyłącznie do użytku wewnętrznego."

    def add_p(text, size=10.5, bold=False, italic=False, color=DARK, before=0, after=8, align="justify"):
        p = doc.add_paragraph()
        p.alignment = {
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
        }[align]
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.widow_control = True
        set_run(p.add_run(text), size, bold, italic, color)
        return p

    def add_h(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.widow_control = True
        set_run(p.add_run(text), 13, True, False, NAVY)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(11)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.widow_control = True
        set_run(p.add_run(text), 11, True, False, NAVY)
        return p

    def add_lead(label, text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.45)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.widow_control = True
        set_run(p.add_run("•  " + label), 10.5, True, False, NAVY)
        set_run(p.add_run(text), 10.5, False, False, DARK)
        return p

    def add_quote(label, quote):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.45)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.12
        set_run(p.add_run("•  " + label + " "), 10, True, False, NAVY)
        set_run(p.add_run(quote), 10, False, True, DARK)
        return p

    def prepare(cell, fill, border="E4EAE7", sz="4", top=60, bottom=60, left=80, right=80):
        shade(cell, fill)
        margins(cell, top, bottom, left, right)
        borders(cell, border, sz, sz, border)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    def cell_para(cell, first, align="left"):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        p.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }[align]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.08
        suppress_hyphen(p)
        return p

    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)
        hp.paragraph_format.line_spacing = Pt(1)
        set_run(hp.add_run(""), 1, color=WHITE)
        htable = header.add_table(1, 2, CONTENT)
        htable.autofit = False
        set_table_width(htable, CONTENT, (Cm(8.8), Cm(8.6)))
        left, right = htable.rows[0].cells
        left.width, right.width = Cm(8.8), Cm(8.6)
        for cell in (left, right):
            shade(cell, "08181F")
            margins(cell, 46, 46, 120, 120)
            borders(cell, "08181F", "2", "2", "08181F")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        lp = left.paragraphs[0]
        lp.paragraph_format.space_before = lp.paragraph_format.space_after = Pt(0)
        if LOGO.exists():
            lp.add_run().add_picture(str(LOGO), width=Cm(4.35))
        else:
            set_run(lp.add_run("ESG INSTITUTE"), 12, True, False, WHITE)
        rp = right.paragraphs[0]
        rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rp.paragraph_format.space_before = Pt(4)
        rp.paragraph_format.space_after = Pt(0)
        set_run(rp.add_run("Sygn. ESG/2026/11/07"), 9, True, False, WHITE)
        rp2 = right.add_paragraph()
        rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rp2.paragraph_format.space_before = Pt(0)
        rp2.paragraph_format.space_after = Pt(0)
        set_run(rp2.add_run("Raport audytu regulacyjnego i środowiskowego"), 8, False, False, PALE)

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.paragraph_format.space_before = Pt(0)
        fp.paragraph_format.space_after = Pt(0)
        ftable = footer.add_table(1, 3, CONTENT)
        ftable.autofit = False
        set_table_width(ftable, CONTENT, (Cm(6.5), Cm(6.7), Cm(4.2)))
        fcells = ftable.rows[0].cells
        fwidths = (Cm(6.5), Cm(6.7), Cm(4.2))
        for cell, width in zip(fcells, fwidths):
            cell.width = width
            shade(cell, "FFFFFF")
            margins(cell, 50, 10, 40, 40)
            cell._tc.get_or_add_tcPr().append(
                docx.oxml.parse_xml(
                    f'<w:tcBorders {nsdecls("w")}>'
                    f'<w:top w:val="single" w:sz="8" w:space="0" w:color="D5DBD8"/>'
                    f'<w:left w:val="nil"/>'
                    f'<w:bottom w:val="nil"/>'
                    f'<w:right w:val="nil"/>'
                    f"</w:tcBorders>"
                )
            )
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        fl = cell_para(fcells[0], True, "left")
        set_run(fl.add_run("Poufne / Wyłącznie do użytku wewnętrznego"), 8, False, False, MUTED)
        fc = cell_para(fcells[1], True, "center")
        set_run(fc.add_run("Audyt greenwashingu — Fundacja Leśna Przestrzeń"), 8, False, False, MUTED)
        fr = cell_para(fcells[2], True, "right")
        set_run(fr.add_run("7 listopada 2026 r."), 8, False, False, MUTED)
        fr2 = cell_para(fcells[2], False, "right")
        set_run(fr2.add_run("Strona "), 8, False, False, MUTED)
        field(fr2, " PAGE ", 8, MUTED)
        set_run(fr2.add_run(" z "), 8, False, False, MUTED)
        field(fr2, " NUMPAGES ", 8, MUTED)

    add_p(
        "Audyt Ryzyk Greenwashingu: FUNDACJA LEŚNA PRZESTRZEŃ",
        size=18,
        bold=True,
        color=NAVY,
        before=2,
        after=4,
        align="left",
    )
    add_p(
        "Weryfikacja witryny lesnaprzestrzen.pl (strona główna, oferta, projekty i powiązane podstrony; "
        "stan na 7 listopada 2026 r.) pod kątem greenwashingu i social washingu.",
        size=10.5,
        italic=True,
        color=MUTED,
        after=8,
        align="left",
    )

    meta = [
        (
            "Badany podmiot",
            "Fundacja Leśna Przestrzeń, KRS 0001213844, NIP 5130308584, REGON 543601897, "
            "ul. Galicyjska 16B, 32-087 Bibice. Odpis aktualny KRS (stan z 24.06.2026): fundacja, "
            "rejestracja 30.12.2025, bez statusu OPP, oddział w Celestynowie. "
            "Strona wskazuje Monikę Wolnik jako osobę prowadzącą działania.",
            False,
        ),
        (
            "Przedmiot audytu",
            "Treść witryny lesnaprzestrzen.pl zebrana 7.11.2026: strona główna, /o-nas, /oferta, "
            "/projekty, /kapiel-lesna-dla-seniorow-gminy-krzeszowice, "
            "/lesna-przestrzen-–-powrot-do-natury, /kontakt oraz tytuły i opisy tych stron. "
            "Profili Instagram (@lesnaprzestrzen) i Facebook nie analizowano wpis po wpisie.",
            False,
        ),
        (
            "Sektor działalności",
            "Fundacja. Kąpiele leśne (także miejskie i on-line), warsztaty florystyczne oraz zajęcia "
            "dla rodzin, szkół, przedszkoli i seniorów. Usługa wobec konsumenta. "
            "To nie jest żywność, produkt finansowy, wyrób budowlany ani środek transportu.",
            False,
        ),
        (
            "Reżim prawny",
            "Ustawa o przeciwdziałaniu nieuczciwym praktykom rynkowym (t.j. Dz.U. 2023 poz. 845); "
            "ustawa o ochronie konkurencji i konsumentów (t.j. Dz.U. 2025 poz. 1714); "
            "ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353). "
            "Czarna lista EmpCo nie jest w dniu audytu prawem polskim.",
            False,
        ),
        ("Data audytu", "7 listopada 2026 r.", False),
        (
            "Werdykt operacyjny",
            "Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści",
            True,
        ),
    ]
    meta_table = doc.add_table(rows=len(meta), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    set_table_width(meta_table, CONTENT, (Cm(4.15), Cm(13.25)))
    for i, (label, value, alert) in enumerate(meta):
        ck, cv = meta_table.rows[i].cells
        ck.width, cv.width = Cm(4.15), Cm(13.25)
        prepare(ck, "F4F6F5")
        prepare(cv, "FFFFFF")
        ck.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pk = cell_para(ck, True, "left")
        set_run(pk.add_run(label), 9, True, False, NAVY)
        pv = cell_para(cv, True, "left")
        set_run(pv.add_run(value), 9, alert, False, BURGUNDY if alert else DARK)

    add_p("", size=4, after=6, align="left")

    verdict = doc.add_table(rows=1, cols=1)
    verdict.alignment = WD_TABLE_ALIGNMENT.CENTER
    verdict.autofit = False
    set_table_width(verdict, CONTENT, (CONTENT,))
    vc = verdict.cell(0, 0)
    vc.width = CONTENT
    prepare(vc, "FBFCFB", "C5CED4", "8", 110, 110, 140, 140)
    vt = cell_para(vc, True, "left")
    set_run(vt.add_run("Werdykt"), 11, True, False, BURGUNDY)
    vs = cell_para(vc, False, "left")
    vs.paragraph_format.space_after = Pt(4)
    set_run(
        vs.add_run("Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści"),
        10,
        True,
        False,
        BURGUNDY,
    )
    for line in (
        "Werdykt opiera się na art. 5 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, "
        "w tekście jednolitym obowiązującym 7 listopada 2026 r. Strona przedstawia kąpiel leśną jako "
        "praktykę, której mierzalny wpływ na zdrowie fizyczne i psychiczne został udowodniony, oraz "
        "nazywa ofertę terapią zdrowotną i terapią leśną.",
        "Osobno spotkanie on-line, skierowane do osób z trudnościami ruchowymi, po wypadku i w izolacji "
        "z powodu choroby, jest opisane jako wsparcie regulacji układu nerwowego. Przeciętny konsument, "
        "a przy tej ofercie przeciętny członek tak nazwanej grupy, może z tego powodu nawiązać kontakt i zapisać się na zajęcia.",
        "Czarna lista dyrektywy (UE) 2024/825 nie jest stwierdzonym naruszeniem prawa polskiego. "
        "Kalendarz stosowania jest w pkt 1.1. Opis dwóch projektów społecznych, z terminem, miejscem, "
        "źródłem wsparcia i efektem oznaczonym jako zakładany, nie spełnia testu social washingu.",
    ):
        vp = cell_para(vc, False, "justify")
        vp.paragraph_format.space_before = Pt(2)
        vp.paragraph_format.space_after = Pt(2)
        vp.paragraph_format.line_spacing = 1.15
        # body of the verdict is continuous prose; allow hyphenation
        set_run(vp.add_run(line), 10.5, False, False, DARK)

    add_h("1. Analiza wstępna")
    add_h2("1.1. Reżim prawny")
    add_p(
        "Art. 4 ust. 1 dyrektywy Parlamentu Europejskiego i Rady (UE) 2024/825 z 28 lutego 2024 r. "
        "wyznacza 27 września 2026 r. jako datę, od której państwa członkowskie stosują krajowe przepisy "
        "transponujące. Nie jest to data wejścia w życie ustawy polskiej. Ustawa z 4 września 2026 r. "
        "o zmianie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawy o prawach konsumenta "
        "(Dz.U. z 2026 r. poz. 1353) została ogłoszona 19 października 2026 r. Art. 3 tej ustawy stanowi, "
        "że wchodzi ona w życie 27 września 2027 r."
    )
    add_p(
        "Do 26 września 2027 r. włącznie werdykt opiera się na przepisach już obowiązujących. "
        "W tym audycie jest to art. 5 ustawy z 23 sierpnia 2007 r. o przeciwdziałaniu nieuczciwym "
        "praktykom rynkowym (tekst jednolity Dz.U. 2023 poz. 845). Czarna lista EmpCo jest przyszłym "
        "standardem, nie stwierdzonym naruszeniem prawa krajowego. Samo utrzymywanie starszej treści "
        "na stronie internetowej tej daty nie przyspiesza. Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1296), "
        "obowiązująca od 6 listopada 2026 r., zmienia art. 29 i 31 ustawy o ochronie konkurencji i konsumentów "
        "i nie zmienia art. 24 ani art. 106. Rozporządzenie (WE) nr 1924/2006 nie stosuje się, ponieważ "
        "przedmiotem są usługi, nie żywność."
    )

    add_h2("1.2. Profil działalności i kwalifikacje")
    add_p(
        "Fundacja Leśna Przestrzeń, zarejestrowana w KRS 30 grudnia 2025 r., z siedzibą w Bibicach, "
        "prowadzi kąpiele leśne i warsztaty florystyczne dla dorosłych, rodzin, szkół, przedszkoli i seniorów. "
        "Część zajęć opisano jako zrealizowane przy wsparciu Gminy Krzeszowice oraz programu Małopolska Lokalnie 2026. "
        "W odpisie aktualnym KRS, stan z 24 czerwca 2026 r., cel działania obejmuje upowszechnianie wiedzy "
        "o pozytywnym wpływie kontaktu z naturą na zdrowie oraz propagowanie kąpieli leśnych i florystyki "
        "jako formy wspierania dobrostanu i profilaktyki zdrowotnej."
    )
    add_p(
        "Cel statutowy tłumaczy dobór słów. Nie zastępuje testu art. 5 wobec osoby, która czyta ofertę "
        "i może się na nią zapisać. Strona sama nazywa podstawę kwalifikacji: ukończenie kursu Forest Therapy Hub "
        "w zakresie prowadzenia kąpieli leśnych oraz egzamin florystyczny FDF. W ramach audytu nie poddano "
        "tego kursu ocenie jako certyfikatu użytego bez uprawnienia."
    )

    add_h2("1.3. Narracja handlowa")
    add_p(
        "Narracja handlowa idzie dalej niż opis spotkania. Tytuły i opisy stron mówią o holistycznych "
        "terapiach zdrowotnych, terapii leśnej i o poprawie zdrowia psychicznego. Na stronie głównej badania, "
        "w tym prace profesora Qing Li, są przedstawione jako dowód realnego, mierzalnego wpływu regularnego "
        "przebywania w lesie na zdrowie psychiczne i fizyczne, a zaraz potem kąpiel leśna jest nazwana sposobem, "
        "który naprawdę odmienia samopoczucie."
    )
    add_p(
        "Oferta definiuje kąpiel jako praktykę profilaktyki zdrowia, także dla seniorów. "
        "Wersja on-line, prowadzona głosem przy roślinach domowych, jest skierowana do osób wracających "
        "do zdrowia po wypadku i izolowanych z powodu choroby oraz ma wspierać regulację układu nerwowego. "
        "Te zdania spełniają art. 5. Zdanie osobiste „wierzę, że natura może być przestrzenią profilaktyki zdrowia” "
        "jest ograniczone słowami „wierzę” i „może” i samo nie niesie ustalenia."
    )

    fig = doc.add_table(rows=1, cols=1)
    fig.alignment = WD_TABLE_ALIGNMENT.CENTER
    fig.autofit = False
    set_table_width(fig, Cm(9.1), (Cm(9.1),), center=True)
    fc = fig.cell(0, 0)
    fc.width = Cm(9.1)
    prepare(fc, "FFFFFF", "B7BFC4", "4", 70, 50, 80, 80)
    fp = cell_para(fc, True, "center")
    fp.paragraph_format.space_before = Pt(2)
    fp.paragraph_format.space_after = Pt(2)
    if HERO.exists():
        fp.add_run().add_picture(str(HERO), width=Cm(8.2))
    cap = cell_para(fc, False, "center")
    cap.paragraph_format.space_before = Pt(2)
    set_run(
        cap.add_run(
            "Rys. 1. Zrzut ekranu strony głównej lesnaprzestrzen.pl (stan na 07.11.2026 r.). "
            "Źródło: opracowanie własne."
        ),
        8.5,
        False,
        True,
        MUTED,
    )
    add_p("", size=2, after=2, align="left")

    add_h2("1.4. Social washing i wyłączenia")
    add_p(
        "Social washing sprawdzono na tym samym teście prawdy, zakresu i pominięcia. "
        "Poniższe przekazy tego progu nie przekraczają."
    )
    add_lead(
        "Projekty społeczne. ",
        "Projekt dla seniorów Gminy Krzeszowice podaje termin, miejsce (Stadniki), jedno spotkanie, "
        "wsparcie gminy oraz efekt 70% jako efekt zakładany i deklarowany przez uczestników, w bloku „Cel zadania”. "
        "Projekt „Leśna Przestrzeń – powrót do natury” podaje termin, program Małopolska Lokalnie 2026 oraz to, "
        "co dofinansowanie umożliwiło: scenariusze, warsztaty i sprzęt. Te opisy nie są ustaleniem. "
        "Zdanie, że zajęcia dla seniorów przeciwdziałają samotności, mieści się w opisie funkcji spotkania w grupie.",
    )
    add_lead(
        "Zajęcia dla dzieci, art. 9 pkt 5. ",
        "Warsztaty dla szkół i przedszkoli nie zawierają bezpośredniego wezwania dzieci do zakupu "
        "ani do nacisku na dorosłych. „Elementy edukacji ekologicznej” opisują temat zajęć.",
    )
    add_lead(
        "Wytyczne OECD. ",
        "Wytyczne OECD dla przedsiębiorstw wielonarodowych (edycja 2023) nie są na stronie przywołane. "
        "Nie powstaje z nich ustalenie, także na poziomie niskim. Postępowanie przed Krajowym Punktem "
        "Kontaktowym nie jest sprawą sądową i nie kończy się wyrokiem. Skutkiem są publiczna ocena, "
        "zalecenia i ryzyko wizerunkowe. Sam ten poziom nie daje werdyktu czerwonego ani żółtego.",
    )
    add_lead(
        "Reklama porównawcza. ",
        "Zestawienie kąpieli ze spacerem nordic walking odróżnia formę ruchu. Nie wskazuje konkurenta "
        "ani jego usługi. Reklamy porównawczej w tym materiale nie ma.",
    )
    add_lead(
        "Art. 7 pkt 17. ",
        "Przepis wymaga twierdzenia, że produkt leczy choroby, zaburzenia lub wady rozwojowe, "
        "oraz że twierdzenie jest niezgodne z prawdą. Takiego zdania na stronie nie ma.",
    )
    add_lead(
        "Przyszła czarna lista EmpCo. ",
        "Definicja twierdzenia dotyczącego ekologiczności z dyrektywy (UE) 2024/825 obejmuje komunikat, "
        "który stwierdza lub sugeruje pozytywny wpływ na środowisko, brak wpływu, mniejszą szkodliwość "
        "albo poprawę w czasie. „Naturalny sposób”, „naturalne materiały”, „więź z naturą” i tytuł projektu "
        "„powrót do natury” opisują formę spotkania, rośliny jako tworzywo oraz nazwę przedsięwzięcia "
        "wyjaśnioną opisem działań. Na tym brzmieniu nie wchodzą w przyszłą czarną listę.",
    )

    add_h("2. Analiza szczegółowa — tabela zbiorcza naruszeń")
    findings = [
        (
            "01",
            "Kąpiel leśna jako udowodniona profilaktyka i terapia zdrowia, także w tytułach stron i w ofercie dla seniorów",
            "Art. 5 ust. 1, ust. 2 pkt 2, ust. 3 pkt 2 i 6 oraz ust. 4 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. Przyszły standard EmpCo tego zdania nie obejmuje.",
            "WYSOKIE",
            "Dowód badań i nazwa terapii bez zakresu protokołu, grupy i wskaźnika oraz bez świadczenia zdrowotnego",
        ),
        (
            "02",
            "Kąpiel on-line wspiera regulację układu nerwowego u osób po wypadku, z chorobą i z trudnościami ruchowymi",
            "Art. 5 ust. 1 i ust. 3 pkt 2, czytany z perspektywy art. 2 pkt 8",
            "WYSOKIE",
            "Obietnica skutku fizjologicznego wobec grupy, którą strona sama określa przez chorobę, wypadek i izolację",
        ),
    ]
    col_w = (Cm(1.15), Cm(4.85), Cm(4.35), Cm(2.15), Cm(4.9))
    sum_table = doc.add_table(rows=1 + len(findings), cols=5)
    sum_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_table.autofit = False
    set_table_width(sum_table, CONTENT, col_w)
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota naruszenia"]
    for c, htxt in enumerate(headers):
        cell = sum_table.rows[0].cells[c]
        cell.width = col_w[c]
        prepare(cell, "08181F", "08181F", "4", 50, 50, 50, 50)
        p = cell_para(cell, True, "center" if c in (0, 3) else "left")
        set_run(p.add_run(htxt), 8, True, False, WHITE)
    for r, row in enumerate(findings):
        bg = "FFFFFF" if r % 2 == 0 else "F7F9F8"
        for c, val in enumerate(row):
            cell = sum_table.rows[r + 1].cells[c]
            cell.width = col_w[c]
            prepare(cell, bg, "E4EAE7", "4", 46, 46, 50, 50)
            align = "center" if c in (0, 3) else "left"
            p = cell_para(cell, True, align)
            if c == 3:
                set_run(p.add_run(val), 8, True, False, BURGUNDY)
            elif c == 0:
                set_run(p.add_run(val), 8, True, False, NAVY)
            else:
                set_run(p.add_run(val), 8, False, False, DARK)

    add_h("3. Szczegółowa analiza przypadków")

    def case_head(title, basis):
        add_p(title, size=12, bold=True, color=NAVY, before=12, after=4, align="left")
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, CONTENT, (Cm(12.7), Cm(4.7)))
        left_c, right_c = table.rows[0].cells
        left_c.width, right_c.width = Cm(12.7), Cm(4.7)
        prepare(left_c, "F4F6F5")
        prepare(right_c, "F4F6F5")
        a = cell_para(left_c, True, "left")
        set_run(a.add_run("Podstawa prawna"), 8, True, False, MUTED)
        b = cell_para(left_c, False, "left")
        set_run(b.add_run(basis), 9, False, False, DARK)
        c = cell_para(right_c, True, "left")
        set_run(c.add_run("Poziom ryzyka"), 8, True, False, MUTED)
        d = cell_para(right_c, False, "left")
        set_run(d.add_run("WYSOKIE"), 11, True, False, BURGUNDY)
        add_p("", size=2, after=2, align="left")

    def claim_box(text):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, CONTENT, (CONTENT,))
        cell = table.cell(0, 0)
        cell.width = CONTENT
        prepare(cell, "F4F6F5", "D5DBD8", "4", 90, 90, 130, 130)
        h = cell_para(cell, True, "left")
        set_run(h.add_run("Dozwolone sformułowanie (wzór po korekcie)"), 9.5, True, False, NAVY)
        body = cell_para(cell, False, "justify")
        body.paragraph_format.space_before = Pt(3)
        body.paragraph_format.line_spacing = 1.15
        set_run(body.add_run(text), 10, False, True, DARK)
        add_p("", size=2, after=4, align="left")

    case_head(
        "Przypadek 01. Kąpiel leśna jako udowodniona profilaktyka i terapia zdrowia",
        "Art. 5 ust. 1 w zw. z ust. 2 pkt 2, ust. 3 pkt 2 i 6 oraz ust. 4 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (Dz.U. 2023 poz. 845).",
    )
    add_p("Cytat źródłowy", size=10.5, bold=True, before=2, after=2, align="left")
    for label, quote in (
        (
            "Strona główna:",
            "„Japońscy naukowcy, m.in. profesor Qing Li, udowodnili badaniami, że regularne przebywanie w lesie ma realny, mierzalny wpływ na nasze zdrowie – zarówno psychiczne, jak i fizyczne.”",
        ),
        (
            "Strona główna, zdanie następne:",
            "„Kąpiele Leśne to łagodny, naturalny sposób na redukcję stresu, regenerację i odzyskanie energii. Bez wysiłku, bez technik — tylko Ty, las i uważna obecność, która naprawdę odmienia samopoczucie.”",
        ),
        (
            "Oferta, definicja usługi:",
            "„Kąpiele leśne (Shinrin-Yoku) to praktyka profilaktyki zdrowia oparta na świadomym kontakcie z naturą.”",
        ),
        (
            "Oferta dla seniorów:",
            "„Spokojne spacery uważnościowe w naturalnym otoczeniu, stworzone z myślą o seniorach stanowią formę profilaktyki zdrowia – wspierają redukcję stresu, poprawę samopoczucia, równowagę psychiczną oraz łagodną aktywność fizyczną sprzyjającą zachowaniu witalności.”",
        ),
        (
            "Tytuł strony „O nas”:",
            "„O nas - Kąpiele leśne i holistyczne terapie zdrowotne.”",
        ),
        (
            "Opis strony „O nas”:",
            "„Poznaj naszą misję: edukacja ekologiczna, terapia leśna oraz warsztaty uważności dla zdrowia psychicznego i redukcji stresu w kontakcie z naturą.”",
        ),
        (
            "Tytuł stron oferty i projektów:",
            "„Holistyczne Terapie: Kąpiele Leśne i Edukacja.”",
        ),
        (
            "Opis strony oferty:",
            "„Odkryj holistyczne podejście do zdrowia. Uczestnicz w kąpielach leśnych i warsztatach uważności, redukując stres i poprawiając zdrowie psychiczne.”",
        ),
        (
            "Opis strony głównej:",
            "„Odkryj holistyczne podejście do zdrowia poprzez kąpiele leśne i warsztaty florystyczne. Dobrostan z natury.”",
        ),
    ):
        add_quote(label, quote)

    add_p("Ocena prawna", size=10.5, bold=True, before=6, after=2, align="left")
    add_p(
        "Art. 5 ust. 1 uznaje za działanie wprowadzające w błąd praktykę, która powoduje lub może powodować "
        "podjęcie przez przeciętnego konsumenta decyzji dotyczącej umowy, której inaczej by nie podjął. "
        "Ust. 2 pkt 2 wymienia rozpowszechnianie prawdziwych informacji w sposób mogący wprowadzać w błąd. "
        "Ust. 3 pkt 2 obejmuje cechy produktu, w tym spodziewane wyniki jego zastosowania oraz wyniki badań. "
        "Ust. 3 pkt 6 obejmuje kwalifikacje i status przedsiębiorcy. Ust. 4 każe uwzględnić wszystkie elementy "
        "i sposób prezentacji. Decyzja dotycząca umowy, według art. 2 pkt 7, obejmuje decyzję, czy, w jaki sposób "
        "i na jakich warunkach konsument dokona zakupu."
    )
    add_p(
        "Wyrok Trybunału Sprawiedliwości z 19 grudnia 2013 r., C-281/12, Trento Sviluppo i Centrale Adriatica, "
        "dotyczy art. 6 ust. 1 i art. 2 lit. k dyrektywy 2005/29/WE, czyli wzorca tego przepisu. "
        "W sentencji, odczytanej w EUR-Lex, praktyka jest wprowadzająca w błąd, gdy zawiera nieprawdziwe informacje "
        "lub może wprowadzić w błąd przeciętnego konsumenta oraz może skłonić go do decyzji dotyczącej transakcji, "
        "której inaczej by nie podjął. Pojęcie tej decyzji obejmuje każdą decyzję bezpośrednio związaną z decyzją o zakupie. "
        "To omówienie sentencji. Wyrok nie dotyczy Fundacji Leśna Przestrzeń. Podobieństwo polega na tym, że samo "
        "wrażenie co do cechy usługi nie wystarcza: potrzebny jest możliwy wpływ na decyzję o udziale. "
        "Tutaj tą decyzją jest kontakt i zapis na zajęcia przedstawione jako terapia oraz jako praktyka "
        "o udowodnionym wpływie na zdrowie fizyczne."
    )
    add_p(
        "Abstrakty prac Q. Li dostępne w PubMed (m.in. 20487629 i 18394317) opisują małe grupy zdrowych dorosłych, "
        "wyjścia obejmujące około dwugodzinne spacery w lesie, czasem w schemacie dwu- lub trzydniowym, oraz pomiary "
        "wybranych wskaźników, w tym aktywności komórek NK, kortyzolu i adrenaliny. To omówienie abstraktów, "
        "nie wynik tej fundacji i nie cytat z orzeczenia. Strona nie podaje publikacji, grupy, wskaźnika ani schematu. "
        "Czasownik „udowodnili” oraz zwrot „wpływ na zdrowie fizyczne i psychiczne” stawiają te badania jako dowód "
        "praktyki sprzedawanej w następnym zdaniu. Oferowane spotkanie strona sama opisuje jako około 2,5 godziny "
        "uważnej obecności, bez dystansu i bez celu sportowego. Tytuły rozciągają ten dowód na „terapie zdrowotne”, "
        "„terapię leśną” i na warsztaty florystyczne."
    )
    add_p(
        "Na stronie o osobie prowadzącej ta sama osoba pisze, że jest absolwentką kursu w zakresie prowadzenia "
        "kąpieli leśnych. Tytuł „terapia” pokazuje odbiorcy inny status niż kurs prowadzenia kąpieli. "
        "Przy ofercie dla seniorów ocenę przeprowadzono z perspektywy przeciętnego seniora. Wiek jest cechą "
        "wymienioną w art. 2 pkt 8. Słowo „wspierają” przy samopoczuciu, czytane osobno, zostałoby przy opisie spaceru. "
        "W tym układzie stoi w zdaniu, które nazywa spacer formą profilaktyki zdrowia, obok tytułów o terapii "
        "i obok zdania o dowodzie. Zdania „wierzę, że” oraz „może być” na stronie o osobie prowadzącej są ostrożniejsze "
        "i nie są podstawą tego przypadku. Historycznego zdania, że praktyka od lat 80. funkcjonuje w Japonii "
        "jako element narodowego programu wspierania zdrowia, nie zakwalifikowano osobno. Nie zastosowano art. 7 pkt 17: "
        "strona nie twierdzi, że usługa leczy choroby, zaburzenia lub wady rozwojowe."
    )
    add_p("Kierunek naprawy", size=10.5, bold=True, before=2, after=2, align="left")
    add_p(
        "Zdjąć ze strony głównej zdanie, że badania udowodniły mierzalny wpływ na zdrowie fizyczne i psychiczne, "
        "oraz zdanie, że uważna obecność naprawdę odmienia samopoczucie, w tym zestawieniu. "
        "Z tytułów i opisów stron usunąć „holistyczne terapie zdrowotne”, „terapia leśna”, "
        "„holistyczne podejście do zdrowia” i „poprawiając zdrowie psychiczne”. "
        "W ofercie, także dla seniorów, zastąpić definicję „praktyka profilaktyki zdrowia” opisem formy spotkania."
    )
    add_p(
        "Nazwę Forest Therapy Hub pisać tak, jak nazywa się organizacja. Kurs prowadzenia kąpieli zostaje "
        "informacją o kwalifikacji. Nie przerabiać go na uprawnienie do terapii. Jeżeli fundacja chce w ogóle "
        "wspomnieć badania pobytu w lesie, robi to według wzoru poniżej, z publikacją i z granicą: "
        "wynik badań nie jest efektem tego spotkania."
    )
    claim_box(
        "„Kąpiel leśna u nas to około 2,5 godziny spokojnego przebywania w lesie, z prostymi zaproszeniami do uważności. "
        "Prowadząca ukończyła kurs Forest Therapy Hub w zakresie prowadzenia kąpieli leśnych. "
        "To nie jest świadczenie zdrowotne ani terapia. Badania pobytu w lesie, w tym prace Qing Li, mierzyły wybrane "
        "wskaźniki przy określonych schematach wyjścia do lasu. Nie przedstawiamy ich jako dowodu efektu tego spotkania. "
        "Warsztat florystyczny jest nauką kompozycji z roślin.”"
    )

    case_head(
        "Przypadek 02. Kąpiel on-line jako regulacja układu nerwowego w chorobie i po wypadku",
        "Art. 5 ust. 1 i ust. 3 pkt 2 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, z perspektywy art. 2 pkt 8 (Dz.U. 2023 poz. 845).",
    )
    add_p("Cytat źródłowy", size=10.5, bold=True, before=2, after=2, align="left")
    for label, quote in (
        (
            "Oferta, kąpiele on-line:",
            "„To sposób dostarczenia doświadczenia kąpieli leśnych osobom z trudnościami ruchowymi, wracającym do zdrowia po wypadku, będącym w deficycie czasu lub w konieczności izolacji, np. ze względu na chorobę.”",
        ),
        (
            "Oferta, opis skutku:",
            "„Uczestnik otacza się elementami natury, które są dla niego aktualnie dostępne – roślinami domowymi, ogrodem lub balkonem – a przewodnik kąpieli leśnej prowadzi głosem i uważnością do wyobraźni, w której mieszka las. Poprzez pracę z oddechem, zmysłami i uwagą spotkania wspierają regulację układu nerwowego, sprzyjają wyciszeniu oraz poczuciu bezpieczeństwa i zakorzenienia.”",
        ),
        (
            "Zdanie ograniczające, które stoi obok:",
            "„Kąpiele leśne on-line nie zastępują bezpośredniego kontaktu z naturą.”",
        ),
    ):
        add_quote(label, quote)

    add_p("Ocena prawna", size=10.5, bold=True, before=6, after=2, align="left")
    add_p(
        "To osobna usługa i osobna grupa. Strona sama wymienia osoby z trudnościami ruchowymi, wracające do zdrowia "
        "po wypadku oraz pozostające w izolacji z powodu choroby. Art. 2 pkt 8 każe oceniać praktykę z uwzględnieniem "
        "przeciętnego członka dającej się zidentyfikować grupy szczególnie podatnej ze względu na szczególne cechy, "
        "takie jak wiek albo niepełnosprawność fizyczna. Trudność ruchowa mieści się w przykładzie niepełnosprawności "
        "fizycznej. Wypadek i choroba są cechą, którą wskazuje sam tekst oferty."
    )
    add_p(
        "Przeciętny członek tej grupy może odczytać „wspierają regulację układu nerwowego” oraz „poczucie bezpieczeństwa” "
        "jako obietnicę skutku dla swojego stanu, a nie jako opis ćwiczenia wyobraźni. Następne zdanie mówi, "
        "że wersja on-line nie zastępuje lasu. Nie mówi, że nie reguluje układu nerwowego i nie wspiera powrotu do zdrowia. "
        "Usługa, według tego samego akapitu, polega na głosie przewodniczki i na roślinach, które uczestnik ma w domu. "
        "Decyzja o udziale osoby w chorobie albo po wypadku jest decyzją dotyczącą umowy z art. 2 pkt 7."
    )
    add_p(
        "Wyrok C-281/12, przywołany w przypadku 01, stosuje się tu tak samo: możliwy wpływ na decyzję bezpośrednio "
        "związaną z zakupem. Wyrok nie dotyczy tej fundacji. To jest social washing tego przekazu. Hasło opieki "
        "wobec osób chorych, izolowanych i po wypadku jest szersze niż materiał, który strona pokazuje: "
        "ćwiczenie oddechu, zmysłów i wyobraźni. Nie zastosowano art. 7 pkt 17. „Wspierają” nie jest twierdzeniem, "
        "że produkt leczy chorobę, zaburzenie lub wadę rozwojową. Test art. 5 i tak jest spełniony, bo spodziewany "
        "wynik zastosowania (art. 5 ust. 3 pkt 2) został nazwany regulacją układu nerwowego wobec grupy w chorobie "
        "i po wypadku, bez podstawy pokazanej na stronie. Wytyczne OECD nie są tu podstawą. To samo zdanie oceniono "
        "na szczeblu art. 5."
    )
    add_p("Kierunek naprawy", size=10.5, bold=True, before=2, after=2, align="left")
    add_p(
        "Usunąć z oferty on-line adresowanie do osób po wypadku i w izolacji z powodu choroby w połączeniu "
        "ze skutkiem „regulacja układu nerwowego” oraz z „poczuciem bezpieczeństwa” jako wynikiem spotkania. "
        "Zostawić można informację, że forma głosowa jest dostępna dla osoby, która nie wychodzi z domu, "
        "oraz opis, na czym spotkanie polega: głos, oddech, uwaga, rośliny w zasięgu ręki. "
        "Nie łączyć tej formy z leczeniem, rekonwalescencją ani z obietnicą fizjologiczną."
    )
    claim_box(
        "„Spotkanie on-line jest ćwiczeniem uważności prowadzonym głosem, przy roślinach, które masz w domu. "
        "Może z niego skorzystać osoba, która tego dnia nie wychodzi z domu. "
        "Spotkanie nie reguluje układu nerwowego, nie wspiera leczenia po wypadku ani w chorobie "
        "i nie zastępuje kontaktu z lekarzem ani wyjścia do lasu.”"
    )

    add_h("4. Ryzyka prawne i finansowe")
    add_h2("4.1. Kara Prezesa UOKiK")
    add_p(
        "Art. 24 ust. 1 ustawy o ochronie konkurencji i konsumentów zakazuje praktyk naruszających zbiorowe "
        "interesy konsumentów. Ust. 2 pkt 3 uznaje za taką praktykę nieuczciwe praktyki rynkowe. "
        "Art. 106 ust. 1 pkt 4 pozwala nałożyć karę do 10% obrotu osiągniętego w roku obrotowym poprzedzającym "
        "rok nałożenia kary, także gdy naruszenie art. 24 było nieumyślne. Tekst jednolity: Dz.U. 2025 poz. 1714. "
        "Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1296), obowiązująca od 6 listopada 2026 r., dodaje kompetencje "
        "przy akcie o usługach cyfrowych w art. 29 i 31 i nie zmienia tych dwóch przepisów. Dwa ustalenia z art. 5 "
        "dotyczą oferty skierowanej do nieoznaczonego kręgu odbiorców, w tym do seniorów oraz do osób chorych i po wypadku."
    )
    add_h2("4.2. Roszczenia cywilne z art. 12 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym")
    add_p(
        "Konsument, którego interes został zagrożony lub naruszony, może żądać zaniechania praktyki, usunięcia "
        "jej skutków, złożenia oświadczenia, obniżenia ceny, naprawienia szkody na zasadach ogólnych oraz "
        "zasądzenia sumy na cel związany z ochroną konsumentów. Z częścią tych roszczeń może wystąpić także "
        "organizacja konsumencka i Rzecznik Praw Obywatelskich. Ryzyko dotyczy osoby, która zapisała się na zajęcia "
        "w oparciu o zdanie o udowodnionym wpływie na zdrowie albo o regulacji układu nerwowego."
    )
    add_h2("4.3. Ten sam przekaz przy projekcie wspieranym publicznie")
    add_p(
        "Oferta dla seniorów używa formuły profilaktyki zdrowia obok projektu, który strona opisuje jako wsparty "
        "przez Gminę Krzeszowice. Opis projektu sam w sobie podaje jedno spotkanie i efekt zakładany. "
        "Ryzyko powstaje wtedy, gdy język terapii i dowodu ze strony głównej oraz z oferty jest czytany jako wynik "
        "tego zadania. To skutek wizerunkowy i rozliczeniowy wobec grantodawcy, przy tym samym art. 5. "
        "Osobnej podstawy zwrotu dotacji w tym audycie nie stwierdzono, ponieważ umowy wsparcia nie poddano analizie."
    )
    add_h2("4.4. Data 27 września 2027 r. nie odsuwa tych dwóch zdań")
    add_p(
        "Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) wchodzi w życie 27 września 2027 r. "
        "Do tego dnia czarna lista EmpCo nie jest prawem polskim. Sprawdzona definicja twierdzenia dotyczącego "
        "ekologiczności z dyrektywy (UE) 2024/825 nie obejmuje obecnych zdań o lesie, roślinach i projekcie lokalnym. "
        "Zdania o terapii, o udowodnionym wpływie na zdrowie i o regulacji układu nerwowego są problemem już na art. 5. "
        "Po 27 września 2027 r. zostają przy tym przepisie, jeżeli pozostaną na stronie. Wytyczne OECD nie dodają "
        "osobnego ustalenia: fundacja ich nie przywołuje, a opis projektów lokalnych nie jest zdaniem o należytej "
        "staranności w łańcuchu dostaw ani o zgodności raportu z Wytycznymi."
    )

    add_h("5. Rekomendacje naprawcze")
    add_p(
        "Osobnego etapu przed 27 września 2027 r. nie wyznaczono. Materiał nie zawiera twierdzenia objętego "
        "przyszłą czarną listą EmpCo. Etap natychmiastowy dotyczy prawa już naruszonego.",
        italic=True,
        color=MUTED,
    )
    add_h2("5.1. Etap natychmiastowy, do 14 listopada 2026 r.")
    add_p(
        "Zatrzymać publikację czterech sformułowań: „udowodnili badaniami (…) mierzalny wpływ na nasze zdrowie”; "
        "„naprawdę odmienia samopoczucie” w połączeniu z tym dowodem; „holistyczne terapie zdrowotne”, „terapia leśna”, "
        "„holistyczne podejście do zdrowia”, „poprawiając zdrowie psychiczne”; „praktyka profilaktyki zdrowia” "
        "oraz „stanowią formę profilaktyki zdrowia” w ofercie, w tym dla seniorów."
    )
    add_p(
        "Z oferty on-line zdjąć „wspierają regulację układu nerwowego” oraz połączenie tej usługi z powrotem "
        "do zdrowia po wypadku i z izolacją z powodu choroby. Do czasu podmiany nie kierować tych zdań do wyszukiwarki: "
        "tytuł i opis strony są częścią przekazu."
    )
    add_h2("5.2. Korekta materiałów, do 8 grudnia 2026 r.")
    add_p(
        "Przepisać stronę główną, /o-nas, /oferta i opisy stron według wzorów z przypadków 01 i 02. "
        "Kurs Forest Therapy Hub zostaje jako kurs prowadzenia kąpieli leśnych. Nazwę organizacji ujednolicić. "
        "W opisie ukończonego projektu dla seniorów zostawić termin, Stadniki, jedno spotkanie, wsparcie "
        "Gminy Krzeszowice i blok „zakładane efekty”. Procentu 70% nie przenosić do zdania o wyniku zmierzonym."
    )
    add_p(
        "Przy tej samej korekcie technicznej: deklaracja dostępności podaje datę publikacji strony 25 września 2035 r. "
        "i powołuje ustawę o dostępności cyfrowej podmiotów publicznych. Polityka prywatności opisuje konto, towar, "
        "przesyłkę i rękojmię. To nie są ustalenia tego audytu. Data i szablon wymagają poprawy, bo nie opisują "
        "tej strony i tego formularza."
    )
    add_h2("5.3. Metodyka, jeżeli badania zostają w tekście, do 7 lutego 2027 r.")
    add_p(
        "Wzmianka o badaniach pobytu w lesie ma w tym samym akapicie: autora, publikację, grupę, wskaźnik, "
        "schemat czasowy i zdanie, że nie jest to efekt spotkania fundacji. Wyniku spaceru w lesie nie przenosić "
        "na kąpiel miejską w parku ani na spotkanie on-line. Tych dwóch form strona nie wiąże z tym samym protokołem."
    )
    add_p(
        "Słów „terapia” i „profilaktyka zdrowia” nie przywracać na podstawie samego kursu przewodnickiego. "
        "Inna kwalifikacja i inny tekst wymagałyby osobnej oceny."
    )
    add_h2("5.4. Sformułowania i szkolenie, do 7 marca 2027 r.")
    add_p(
        "Osoba prowadząca, zarząd i wolontariusze używają wzorów z przypadków 01 i 02 na stronie, "
        "w ogłoszeniu o naborze i w odpowiedzi na zapytanie o „terapię” albo o „skutek zdrowotny”. "
        "Projekt społeczny opisuje: kto wsparł, kiedy, gdzie, ilu uczestników i czy liczba jest celem, "
        "czy policzonym wynikiem. „Przeciwdziałanie samotności” zostaje tylko jako opis spotkania w grupie, "
        "bez miary, której fundacja nie pokazuje."
    )
    add_p(
        "Profile Instagram i Facebook, wskazane na stronie o osobie prowadzącej, objąć tym samym językiem. "
        "Wpisów z tych profili w tym audycie nie analizowano."
    )

    add_h("6. Wnioski")
    add_p(
        "W dniu 7 listopada 2026 r. witryna narusza art. 5 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym "
        "w dwóch miejscach: w przedstawieniu kąpieli leśnej jako udowodnionej terapii i profilaktyki zdrowia oraz "
        "w przedstawieniu spotkania on-line jako wsparcia regulacji układu nerwowego dla osób chorych i po wypadku. "
        "Opisy projektów dla seniorów i „powrotu do natury” tego progu nie przekraczają."
    )
    add_p(
        "Czarna lista EmpCo stanie się prawem polskim 27 września 2027 r. i na obecnym brzmieniu tych zdań nie przejmuje. "
        "Korekta wynika z art. 5 i jest natychmiastowa."
    )

    add_p("Zatwierdzenie", size=12, bold=True, color=NAVY, before=16, after=2, align="left")
    add_p(
        "Raport sporządził Zespół Audytu Regulacyjnego i Prawa Konsumenckiego, ESG Institute.",
        size=10.5,
        after=10,
        align="left",
    )
    add_p("................................................", size=11, before=8, after=0, align="left")
    add_p("Imię i nazwisko", size=10.5, after=0, align="left")
    add_p("Stanowisko", size=10.5, color=MUTED, after=0, align="left")
    add_p("ESG Institute", size=10.5, after=0, align="left")
    add_p("Data: 7 listopada 2026 r.", size=10.5, after=0, align="left")
    add_p("Sygn. ESG/2026/11/07", size=10.5, after=0, align="left")

    doc.save(REPORT)


def add_simple(doc, text, size=11, bold=False, italic=False, before=0, after=6, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.12
    set_run(p.add_run(text), size, bold, italic, color or RGBColor(22, 36, 44))
    return p


def build_material():
    doc = setup(docx.Document())
    NAVY = RGBColor(8, 24, 31)
    MUTED = RGBColor(90, 107, 119)
    add_simple(
        doc,
        "ESG INSTITUTE · ZBIÓR TREŚCI DO AUDYTU",
        size=9,
        bold=True,
        before=0,
        after=2,
        color=RGBColor(42, 138, 176),
    )
    add_simple(
        doc,
        "Materiał wejściowy: Fundacja Leśna Przestrzeń",
        size=18,
        bold=True,
        before=2,
        after=4,
        color=NAVY,
    )
    add_simple(
        doc,
        "Zbiór dosłownych haseł i opisów z witryny. Bez oceny prawnej. Data zebrania: 7 listopada 2026 r.",
        size=10.5,
        italic=True,
        color=MUTED,
    )
    rows = [
        ("Podmiot", "Fundacja Leśna Przestrzeń, ul. Galicyjska 16B, 32-087 Bibice"),
        ("Identyfikatory podane na stronie", "KRS 0001213844, NIP 5130308584, REGON 543601897"),
        ("Osoba wskazana na stronie", "Monika Wolnik, telefon 506-110-616, kontakt@lesnaprzestrzen.pl"),
        (
            "Adresy źródeł",
            "https://lesnaprzestrzen.pl/ ; /o-nas ; /oferta ; /projekty ; "
            "/kapiel-lesna-dla-seniorow-gminy-krzeszowice ; "
            "/lesna-przestrzen-–-powrot-do-natury ; /kontakt",
        ),
        ("Data zebrania", "7 listopada 2026 r."),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    for i, (k, v) in enumerate(rows):
        table.rows[i].cells[0].text = ""
        table.rows[i].cells[1].text = ""
        set_run(table.rows[i].cells[0].paragraphs[0].add_run(k), 9, True)
        set_run(table.rows[i].cells[1].paragraphs[0].add_run(v), 9, False)

    sections = [
        (
            "Strona główna",
            [
                "Tytuł: „Kąpiele leśne i warsztaty oparte o kontakt z naturą dla Twojego dobrostanu.”",
                "Opis meta: „Odkryj holistyczne podejście do zdrowia poprzez kąpiele leśne i warsztaty florystyczne. Dobrostan z natury.”",
                "„Fundacja Leśna Przestrzeń.”",
                "„Odkryj harmonię zdrowia zanurzając się w naturę.”",
                "„Towarzyszę innym w odkrywaniu głębokiej więzi z przyrodą jako florystka (absolwentka Polskiej Szkoły Florystycznej) i certyfikowana przewodniczka kąpieli leśnych (Forest Therapy Hub). Wierzę, że natura może być przestrzenią profilaktyki zdrowia, regeneracji i budowania wewnętrznej równowagi.”",
                "„Leśna Przestrzeń to miejsce, w którym prowadzimy kąpiele leśne i warszataty florystyczne dla dorosłych i rodzin. Nasze działania wspierają rozwój uważności, integrację i budowanie głębszej więzi z naturą.”",
                "„Kąpiele Leśne – Twoja droga do równowagi i życiowej energii.”",
                "„Kąpiele leśne to spokojne, uważne przebywanie w lesie. Nie jest to spacer nordic walking, bieganie ani żadna forma aktywności sportowej. To czas, w którym po prostu jesteś w środku natury, pozwalając sobie zwolnić, odetchnąć i poczuć las wszystkimi zmysłami.”",
                "„Ta praktyka wywodzi się z Japonii, gdzie od lat 80. XX wieku funkcjonuje pod nazwą Shinrin-Yoku, co oznacza „zanurzenie się w atmosferze lasu”. Japońscy naukowcy, m.in. profesor Qing Li, udowodnili badaniami, że regularne przebywanie w lesie ma realny, mierzalny wpływ na nasze zdrowie – zarówno psychiczne, jak i fizyczne.”",
                "„Kąpiele Leśne to łagodny, naturalny sposób na redukcję stresu, regenerację i odzyskanie energii. Bez wysiłku, bez technik — tylko Ty, las i uważna obecność, która naprawdę odmienia samopoczucie.”",
                "„To twórcze spotkania z naturą, wspierające kreatywne myślenie i naukę tworzenia kompozycji z kwiatów. Podczas pracy z żywymi roślinami uczysz się patrzeć świeżym okiem, eksperymentować z formą i kolorem oraz ufać własnej intuicji. Naturalne materiały pomagają się zatrzymać, wyciszyć i skupić na chwili obecnej.”",
                "„To czas, w którym rozwijasz wyobraźnię, budujesz wewnętrzną równowagę i doświadczasz kojącej mocy kontaktu z naturą.”",
            ],
        ),
        (
            "O mnie",
            [
                "Tytuł: „O nas - Kąpiele leśne i holistyczne terapie zdrowotne.”",
                "Opis meta: „Poznaj naszą misję: edukacja ekologiczna, terapia leśna oraz warsztaty uważności dla zdrowia psychicznego i redukcji stresu w kontakcie z naturą.”",
                "„Nazywam się Monika Wolnik. Jestem certyfikowaną przewodniczką kąpieli leśnych, florystką oraz specjalistką z wieloletnim doświadczeniem w branży IT.”",
                "„Prowadzę Fundację Leśna Przestrzeń, która powstała z potrzeby tworzenia bezpiecznych miejsc wspierających dobrostan psychiczny, równowagę i kontakt z naturą a także kreatywności w działaniach artystycznych. (…) Z doświadczenia wiem, że uważna obecność w naturze może być realnym wsparciem w czasach przeciążenia, nadmiaru bodźców i życia w ciągłym pośpiechu.”",
                "„Certyfikowana przewodniczka kąpieli leśnych oraz florystka. Absolwentka kursu Forest Theraphy Hub w zakresie prowadzenia kąpieli leśnych i Polskiej Szkoły Florystycznej w Poznaniu. Egzamin na tytuł Florystki zdałam w Gelsenkirchen (Niemcy), uzyskując akredytację FDF – Fachverband Deutscher Floristen.”",
                "„Jako przewodniczka kąpieli leśnych realizowałam spotkania w grupach liczących od 4 do 15 uczestników (…). Posiadane kompetencje oraz praktyka w pracy warsztatowej zapewniają wysoką jakość merytoryczną realizowanych działań oraz ich dostosowanie do możliwości i potrzeb uczestników.”",
                "„Gospodyni profilów na Instagramie - @lesnaprzestrzen i na Facebooku - Leśna Przestrzeń, na których dzielę się tym, jak kąpiele leśne, florystyka i natura ratuje nasz zdrowy rozsądek.”",
            ],
        ),
        (
            "Oferta",
            [
                "Tytuł: „Holistyczne Terapie: Kąpiele Leśne i Edukacja.”",
                "Opis meta: „Odkryj holistyczne podejście do zdrowia. Uczestnicz w kąpielach leśnych i warsztatach uważności, redukując stres i poprawiając zdrowie psychiczne.”",
                "Warsztaty florystyczne: „Proces tworzenia kompozycji działa regulująco i wyciszająco, sprzyja koncentracji oraz pozwala doświadczyć przyjemności płynącej z bycia „tu i teraz”.” „To harmonijne połączenie rozwoju osobistego i rzetelnej wiedzy, które wspiera zarówno dobrostan, jak i kreatywną ekspresję.”",
                "„Kąpiele leśne (Shinrin-Yoku) to praktyka profilaktyki zdrowia oparta na świadomym kontakcie z naturą.” „Kąpiele leśne trwa zazwyczaj ok. 2,5 godziny.” „Praktyka ta pochodzi z Japonii, gdzie od lat 80. XX wieku funkcjonuje jako element narodowego programu wspierania zdrowia i redukcji stresu.”",
                "„Miejska kąpiel leśna to spotkanie z naturą, które odbywa się w parkach, zieleniach miejskich i przestrzeniach, gdzie natura wciąż ma swoje miejsce.” „Miejskie kąpiele leśne trwa zazwyczaj 1,5–2 godziny.”",
                "„Leśne warsztaty dla szkół i przedszkoli to zajęcia prowadzone w naturalnym otoczeniu, które wspierają rozwój dzieci poprzez bezpośredni kontakt z przyrodą. (…) Program zajęć łączy elementy edukacji ekologicznej, uważności i swobodnej zabawy.”",
                "Seniorzy: „Spokojne spacery uważnościowe w naturalnym otoczeniu, stworzone z myślą o seniorach stanowią formę profilaktyki zdrowia – wspierają redukcję stresu, poprawę samopoczucia, równowagę psychiczną oraz łagodną aktywność fizyczną sprzyjającą zachowaniu witalności.” „Wspólny kontakt z przyrodą, rozmowa i uważne doświadczanie chwili wzmacniają integrację społeczną, przeciwdziałają samotności i budują poczucie wspólnoty.”",
                "On-line: „To sposób dostarczenia doświadczenia kąpieli leśnych osobom z trudnościami ruchowymi, wracającym do zdrowia po wypadku, będącym w deficycie czasu lub w konieczności izolacji, np. ze względu na chorobę.” „Poprzez pracę z oddechem, zmysłami i uwagą spotkania wspierają regulację układu nerwowego, sprzyjają wyciszeniu oraz poczuciu bezpieczeństwa i zakorzenienia.” „Kąpiele leśne on-line nie zastępują bezpośredniego kontaktu z naturą.”",
            ],
        ),
        (
            "Projekty",
            [
                "„Kąpiele leśne dla seniorów Gminy Krzeszowice. Termin realizacji: kwiecień – lipiec 2026 r. Projekt skierowany do seniorów – mieszkańców Gminy Krzeszowice.”",
                "Podstrona projektu: „Termin: 1 kwietnia 2026 – 31 lipca 2026.” „Projekt to kameralne kąpiele leśne dla seniorów z Gminy Krzeszowice, prowadzone przez certyfikowaną przewodniczkę Forest Theraphy Hub - Monikę Wolnik. Zajęcia odbyły się w Stadnikach i łączyły spokojny spacer, proste ćwiczenia sensoryczne, elementy edukacji zdrowotnej oraz integrację społeczną. Przewidziano jedno spotkanie w maju dla 8–15 uczestników.”",
                "„Cel zadania: Zwiększenie aktywności i integracji społecznej osób w wieku emerytalnym poprzez bezpieczny kontakt z naturą; podniesienie świadomości roli przyrody w dobrostanie seniorów; przeciwdziałanie izolacji i bierności; wzmocnienie nawyku aktywnego, lokalnego wypoczynku. Zakładane efekty obejmują m.in. deklarowaną przez min. 70% uczestników poprawę samopoczucia, poczucia relaksu i przynależności do grupy oraz chęć udziału w podobnych inicjatywach w przyszłości.”",
                "„Projekt realizowany przy wsparciu finansowym Gminy Krzeszowice.”",
                "„Projekt „Leśna Przestrzeń – powrót do natury” w ramach programu Małopolska Lokalnie 2026. Termin realizacji: czerwiec – listopad 2026 r.”",
                "Artykuł projektu: „Termin: 3 czerwca 2026 – 4 listopada 2026.” „(…) zrealizowany dzięki wsparciu programu Małopolska Lokalnie 2026.” „Dzięki otrzymanemu dofinansowaniu mogliśmy (…) przygotować trzy autorskie scenariusze kąpieli leśnych, rozbudować naszą ofertę warsztatów florystycznych i (…) kupić własny sprzęt.” „Podczas kąpieli leśnej w lesie Michałowickim (…). (…) warsztaty florystyczne w Zielonkach.” „Powstał artykuł „Kąpiel leśna jako doświadczenie społeczne”.”",
            ],
        ),
        (
            "Kontakt",
            [
                "„506-110-616.” „kontakt@lesnaprzestrzen.pl.”",
                "„KRS 0001213844. NIP 5130308584. REGON 543601897.”",
                "„Fundacja Leśna Przestrzeń. Galicyjska 16b. Bibice, 32-087.”",
            ],
        ),
    ]
    for title, paras in sections:
        add_simple(doc, title, size=14, bold=True, before=14, after=4, color=NAVY)
        for para in paras:
            add_simple(doc, para, size=10.5, after=4)
    doc.save(MATERIAL)


if __name__ == "__main__":
    build_material()
    build_report()
    print(MATERIAL)
    print(REPORT)
