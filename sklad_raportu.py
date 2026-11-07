#!/usr/bin/env python3
"""Skład graficzny raportu audytu greenwashingu.

Wiążący opis jest w INSTRUKCJA_GENEROWANIA_RAPORTOW.md.
Otwarty wzór formy: wzory/WZOR_raport.docx (generuj_wzor_raportu.py).
Przykład wypełniony: LESNA_PRZESTRZEN/Raport_Audytu_Greenwashingu_Fundacja_Lesna_Przestrzen.docx.

Identyfikacja ESG Institute: pas #08181F, logo po lewej, sygnatura po prawej.
Tekst główny wyjustowany. Tabele do lewej. Najwyżej dwie ramki: werdykt i wzór po korekcie.
"""

from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

TOOL = Path(__file__).resolve().parent

NAVY = RGBColor(8, 24, 31)
DARK = RGBColor(22, 36, 44)
MUTED = RGBColor(90, 107, 119)
BURGUNDY = RGBColor(142, 47, 57)
AMBER = RGBColor(140, 90, 20)
WHITE = RGBColor(255, 255, 255)
PALE = RGBColor(186, 214, 224)

VERDICT_RED = "Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści"
VERDICT_AMBER = "Publikacja warunkowa – rekomendowana korekta przed publikacją"
VERDICT_GREEN = "Dopuszczone do publikacji"

RISK_COLOR = {
    "BARDZO WYSOKIE": BURGUNDY,
    "WYSOKIE": BURGUNDY,
    "ŚREDNIE": AMBER,
    "NISKIE": MUTED,
}


def default_logo():
    candidates = [
        TOOL.parent
        / "ESG Experts Day - Komplet Materialow (PDF + PPTX)"
        / "03_DLA_PROWADZACEGO"
        / "assets"
        / "esg_logo_white.png",
        TOOL.parent / "ESG Hero Day - Materialy Druk" / "assets" / "esg_logo_white.png",
    ]
    for path in candidates:
        if path.exists():
            return path
    return None


def set_run(run, size, bold=False, italic=False, color=None, font="Arial"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "pl-PL")
    lang.set(qn("w:eastAsia"), "pl-PL")


def shade(cell, hex_color):
    cell._tc.get_or_add_tcPr().append(
        docx.oxml.parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    )


def cell_margins(cell, top=80, bottom=80, left=90, right=90):
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


def cell_borders(cell, color="E4EAE7", sz="4"):
    cell._tc.get_or_add_tcPr().append(
        docx.oxml.parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f"</w:tcBorders>"
        )
    )


def field(paragraph, instruction, size, color, placeholder="1"):
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
    set_run(begin_run, size, False, False, color)
    result = paragraph.add_run(placeholder)
    set_run(result, size, False, False, color)
    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)
    set_run(end_run, size, False, False, color)


def set_table_width(table, width, col_widths=None, center=False):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(width.twips)))
    tbl_w.set(qn("w:type"), "dxa")
    existing = tbl_pr.find(qn("w:tblLayout"))
    if existing is not None:
        tbl_pr.remove(existing)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    if center:
        jc = tbl_pr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            tbl_pr.append(jc)
        jc.set(qn("w:val"), "center")
    if not col_widths:
        return
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
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
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(col.twips)))
            tc_w.set(qn("w:type"), "dxa")


def enable_hyphenation(document):
    settings = document.settings.element
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
    p_pr = paragraph._p.get_or_add_pPr()
    el = OxmlElement("w:suppressAutoHyphens")
    el.set(qn("w:val"), "true")
    p_pr.append(el)


class ReportDoc:
    """Dokument A4 w składzie kancelaryjnym ESG Institute."""

    def __init__(self, reference, document_name, audit_date, logo=None):
        self.reference = reference
        self.document_name = document_name
        self.audit_date = audit_date
        self.logo = Path(logo) if logo else default_logo()
        self.content = Cm(17.4)
        self.doc = docx.Document()
        enable_hyphenation(self.doc)
        core = self.doc.core_properties
        core.author = "ESG Institute"
        core.category = "Raport audytu"
        core.subject = reference
        core.comments = "Poufne. Wyłącznie do użytku wewnętrznego."
        for section in self.doc.sections:
            section.page_width = Mm(210)
            section.page_height = Mm(297)
            section.top_margin = Cm(2.3)
            section.bottom_margin = Cm(1.85)
            section.left_margin = Cm(1.8)
            section.right_margin = Cm(1.8)
            section.header_distance = Cm(0.32)
            section.footer_distance = Cm(0.28)
            self._header(section)
            self._footer(section)

    def save(self, path):
        self.doc.core_properties.title = self.document_name
        self.doc.save(path)
        return Path(path)

    def _prepare(self, cell, fill, border="E4EAE7", sz="4", top=60, bottom=60, left=80, right=80):
        shade(cell, fill)
        cell_margins(cell, top, bottom, left, right)
        cell_borders(cell, border, sz)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    def _cell_para(self, cell, first, align="left"):
        paragraph = cell.paragraphs[0] if first else cell.add_paragraph()
        paragraph.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }[align]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(1)
        paragraph.paragraph_format.line_spacing = 1.08
        suppress_hyphen(paragraph)
        return paragraph

    def _header(self, section):
        header = section.header
        header.is_linked_to_previous = False
        lead = header.paragraphs[0]
        lead.paragraph_format.space_before = Pt(0)
        lead.paragraph_format.space_after = Pt(0)
        lead.paragraph_format.line_spacing = Pt(1)
        set_run(lead.add_run(""), 1, color=WHITE)
        table = header.add_table(1, 2, self.content)
        table.autofit = False
        set_table_width(table, self.content, (Cm(8.8), Cm(8.6)))
        left, right = table.rows[0].cells
        for cell in (left, right):
            shade(cell, "08181F")
            cell_margins(cell, 46, 46, 120, 120)
            cell_borders(cell, "08181F", "2")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        logo_p = left.paragraphs[0]
        logo_p.paragraph_format.space_before = logo_p.paragraph_format.space_after = Pt(0)
        if self.logo and self.logo.exists():
            logo_p.add_run().add_picture(str(self.logo), width=Cm(4.35))
        else:
            set_run(logo_p.add_run("ESG INSTITUTE"), 12, True, False, WHITE)
        ref = right.paragraphs[0]
        ref.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ref.paragraph_format.space_before = Pt(4)
        ref.paragraph_format.space_after = Pt(0)
        set_run(ref.add_run(self.reference), 9, True, False, WHITE)
        kind = right.add_paragraph()
        kind.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        kind.paragraph_format.space_before = Pt(0)
        kind.paragraph_format.space_after = Pt(0)
        set_run(kind.add_run("Raport audytu regulacyjnego i środowiskowego"), 8, False, False, PALE)

    def _footer(self, section):
        footer = section.footer
        footer.is_linked_to_previous = False
        lead = footer.paragraphs[0]
        lead.paragraph_format.space_before = Pt(0)
        lead.paragraph_format.space_after = Pt(0)
        table = footer.add_table(1, 3, self.content)
        table.autofit = False
        widths = (Cm(6.5), Cm(6.7), Cm(4.2))
        set_table_width(table, self.content, widths)
        cells = table.rows[0].cells
        for cell, width in zip(cells, widths):
            cell.width = width
            shade(cell, "FFFFFF")
            cell_margins(cell, 50, 10, 40, 40)
            cell._tc.get_or_add_tcPr().append(
                docx.oxml.parse_xml(
                    f'<w:tcBorders {nsdecls("w")}>'
                    f'<w:top w:val="single" w:sz="8" w:space="0" w:color="D5DBD8"/>'
                    f'<w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
                    f"</w:tcBorders>"
                )
            )
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        left = self._cell_para(cells[0], True, "left")
        set_run(left.add_run("Poufne / Wyłącznie do użytku wewnętrznego"), 8, False, False, MUTED)
        mid = self._cell_para(cells[1], True, "center")
        set_run(mid.add_run(self.document_name), 8, False, False, MUTED)
        right = self._cell_para(cells[2], True, "right")
        set_run(right.add_run(self.audit_date), 8, False, False, MUTED)
        pages = self._cell_para(cells[2], False, "right")
        set_run(pages.add_run("Strona "), 8, False, False, MUTED)
        field(pages, " PAGE ", 8, MUTED)
        set_run(pages.add_run(" z "), 8, False, False, MUTED)
        field(pages, " NUMPAGES ", 8, MUTED)

    def paragraph(self, text, size=10.5, bold=False, italic=False, color=DARK, before=0, after=8, align="justify"):
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = {
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
        }[align]
        paragraph.paragraph_format.space_before = Pt(before)
        paragraph.paragraph_format.space_after = Pt(after)
        paragraph.paragraph_format.line_spacing = 1.15
        paragraph.paragraph_format.widow_control = True
        set_run(paragraph.add_run(text), size, bold, italic, color)
        return paragraph

    def title(self, text):
        return self.paragraph(text, size=18, bold=True, color=NAVY, before=2, after=4, align="left")

    def subtitle(self, text):
        return self.paragraph(text, size=10.5, italic=True, color=MUTED, after=8, align="left")

    def h1(self, text):
        paragraph = self.paragraph(text, size=13, bold=True, color=NAVY, before=16, after=6, align="left")
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.line_spacing = 1.05
        return paragraph

    def h2(self, text):
        paragraph = self.paragraph(text, size=11, bold=True, color=NAVY, before=11, after=3, align="left")
        paragraph.paragraph_format.keep_with_next = True
        return paragraph

    def bullet(self, label, text):
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.left_indent = Cm(0.45)
        paragraph.paragraph_format.first_line_indent = Cm(-0.35)
        paragraph.paragraph_format.space_before = Pt(1)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.15
        set_run(paragraph.add_run("•  " + label), 10.5, True, False, NAVY)
        set_run(paragraph.add_run(text), 10.5, False, False, DARK)
        return paragraph

    def quote(self, label, quote):
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.left_indent = Cm(0.45)
        paragraph.paragraph_format.first_line_indent = Cm(-0.35)
        paragraph.paragraph_format.space_before = Pt(1)
        paragraph.paragraph_format.space_after = Pt(3)
        paragraph.paragraph_format.line_spacing = 1.12
        set_run(paragraph.add_run("•  " + label + " "), 10, True, False, NAVY)
        set_run(paragraph.add_run(quote), 10, False, True, DARK)
        return paragraph

    def meta(self, rows):
        table = self.doc.add_table(rows=len(rows), cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = (Cm(4.15), Cm(13.25))
        set_table_width(table, self.content, widths)
        for index, (label, value, alert) in enumerate(rows):
            key, val = table.rows[index].cells
            self._prepare(key, "F4F6F5")
            self._prepare(val, "FFFFFF")
            key.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            key_p = self._cell_para(key, True, "left")
            set_run(key_p.add_run(label), 9, True, False, NAVY)
            val_p = self._cell_para(val, True, "left")
            set_run(val_p.add_run(value), 9, alert, False, BURGUNDY if alert else DARK)
        self.paragraph("", size=4, after=6, align="left")
        return table

    def verdict(self, status, paragraphs):
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, self.content, (self.content,))
        cell = table.cell(0, 0)
        self._prepare(cell, "FBFCFB", "C5CED4", "8", 110, 110, 140, 140)
        heading = self._cell_para(cell, True, "left")
        set_run(heading.add_run("Werdykt"), 11, True, False, BURGUNDY)
        status_p = self._cell_para(cell, False, "left")
        status_p.paragraph_format.space_after = Pt(4)
        set_run(status_p.add_run(status), 10, True, False, BURGUNDY)
        for line in paragraphs:
            body = self._cell_para(cell, False, "justify")
            body.paragraph_format.space_before = Pt(2)
            body.paragraph_format.space_after = Pt(2)
            body.paragraph_format.line_spacing = 1.15
            set_run(body.add_run(line), 10.5, False, False, DARK)
        return table

    def findings(self, rows):
        table = self.doc.add_table(rows=1 + len(rows), cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = (Cm(1.15), Cm(4.85), Cm(4.35), Cm(2.15), Cm(4.9))
        set_table_width(table, self.content, widths)
        headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota naruszenia"]
        for col, text in enumerate(headers):
            cell = table.rows[0].cells[col]
            self._prepare(cell, "08181F", "08181F", "4", 50, 50, 50, 50)
            paragraph = self._cell_para(cell, True, "center" if col in (0, 3) else "left")
            set_run(paragraph.add_run(text), 8, True, False, WHITE)
        for row_index, row in enumerate(rows):
            fill = "FFFFFF" if row_index % 2 == 0 else "F7F9F8"
            for col, value in enumerate(row):
                cell = table.rows[row_index + 1].cells[col]
                self._prepare(cell, fill, "E4EAE7", "4", 46, 46, 50, 50)
                paragraph = self._cell_para(cell, True, "center" if col in (0, 3) else "left")
                if col == 3:
                    set_run(paragraph.add_run(value), 8, True, False, RISK_COLOR.get(value, BURGUNDY))
                elif col == 0:
                    set_run(paragraph.add_run(value), 8, True, False, NAVY)
                else:
                    set_run(paragraph.add_run(value), 8, False, False, DARK)
        return table

    def case(self, title, basis, risk):
        self.paragraph(title, size=12, bold=True, color=NAVY, before=12, after=4, align="left")
        table = self.doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, self.content, (Cm(12.7), Cm(4.7)))
        left, right = table.rows[0].cells
        self._prepare(left, "F4F6F5")
        self._prepare(right, "F4F6F5")
        label = self._cell_para(left, True, "left")
        set_run(label.add_run("Podstawa prawna"), 8, True, False, MUTED)
        body = self._cell_para(left, False, "left")
        set_run(body.add_run(basis), 9, False, False, DARK)
        risk_label = self._cell_para(right, True, "left")
        set_run(risk_label.add_run("Poziom ryzyka"), 8, True, False, MUTED)
        risk_value = self._cell_para(right, False, "left")
        set_run(risk_value.add_run(risk), 11, True, False, RISK_COLOR.get(risk, BURGUNDY))
        self.paragraph("", size=2, after=2, align="left")
        return table

    def claim(self, text):
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, self.content, (self.content,))
        cell = table.cell(0, 0)
        self._prepare(cell, "F4F6F5", "D5DBD8", "4", 90, 90, 130, 130)
        heading = self._cell_para(cell, True, "left")
        set_run(heading.add_run("Dozwolone sformułowanie (wzór po korekcie)"), 9.5, True, False, NAVY)
        body = self._cell_para(cell, False, "justify")
        body.paragraph_format.space_before = Pt(3)
        body.paragraph_format.line_spacing = 1.15
        set_run(body.add_run(text), 10, False, True, DARK)
        self.paragraph("", size=2, after=4, align="left")
        return table

    def figure(self, image_path, caption, width_cm=8.2):
        frame = Cm(width_cm + 0.9)
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, frame, (frame,), center=True)
        cell = table.cell(0, 0)
        self._prepare(cell, "FFFFFF", "B7BFC4", "4", 70, 50, 80, 80)
        picture = self._cell_para(cell, True, "center")
        picture.paragraph_format.space_before = Pt(2)
        picture.paragraph_format.space_after = Pt(2)
        path = Path(image_path) if image_path else None
        if path and path.exists():
            picture.add_run().add_picture(str(path), width=Cm(width_cm))
        else:
            set_run(
                picture.add_run("Tu umieszcza się zrzut witryny, jeżeli materiał z niej pochodzi."),
                9,
                False,
                True,
                MUTED,
            )
        cap = self._cell_para(cell, False, "center")
        cap.paragraph_format.space_before = Pt(2)
        set_run(cap.add_run(caption), 8.5, False, True, MUTED)
        self.paragraph("", size=2, after=2, align="left")
        return table

    def sign_off(self):
        self.paragraph("Zatwierdzenie", size=12, bold=True, color=NAVY, before=16, after=2, align="left")
        self.paragraph(
            "Raport sporządził Zespół Audytu Regulacyjnego i Prawa Konsumenckiego, ESG Institute.",
            size=10.5,
            after=10,
            align="left",
        )
        self.paragraph("................................................", size=11, before=8, after=0, align="left")
        self.paragraph("Imię i nazwisko", size=10.5, after=0, align="left")
        self.paragraph("Stanowisko", size=10.5, color=MUTED, after=0, align="left")
        self.paragraph("ESG Institute", size=10.5, after=0, align="left")
        self.paragraph("Data: " + self.audit_date, size=10.5, after=0, align="left")
        self.paragraph(self.reference, size=10.5, after=0, align="left")
