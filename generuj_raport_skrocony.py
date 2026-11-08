#!/usr/bin/env python3
"""Raport skrócony audytu greenwashingu.

Wzór formy: wzory/WZOR_raport_skrocony.docx
Raport: LIBRA/Raport_Skrocony_Audytu_Greenwashingu_Libra_Partners.docx

Trzy ustalenia rozstrzygające werdykt. Dwa ustalenia o ryzyku średnim zostają
w raporcie pełnym. Kalendarz i podstawy: INSTRUKCJA_GENEROWANIA_RAPORTOW.md.
Identyfikacja: pas #08181F, błękit #3A61B3, akcent #2F8DBC, pole #E8F0FA.
"""

from io import BytesIO
from pathlib import Path

import docx
from PIL import Image
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
LOGO = ROOT.parent / "ESG Hero Day - Materialy Druk" / "assets" / "esg_logo_white.png"
SHOT = ROOT / "LIBRA" / "zrodla" / "libra-en-strona-2026-10-30.png"

NAVY = RGBColor(8, 24, 31)
BLUE = RGBColor(58, 97, 179)
CYAN = RGBColor(47, 141, 188)
CYAN_SOFT = RGBColor(142, 197, 224)
DARK = RGBColor(22, 36, 44)
MUTED = RGBColor(90, 107, 119)
WHITE = RGBColor(255, 255, 255)
DANGER = RGBColor(190, 18, 60)
WARNING = RGBColor(180, 83, 9)
SUCCESS = RGBColor(4, 120, 87)
GREEN_TXT = RGBColor(22, 101, 52)

# A4 21,0 cm. Marginesy lewy i prawy po 2,0 cm — szerokość kolumny treści.
PAGE_W = 17.0


def _shd(cell, hex_color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="{hex_color}"/>'))


def _mar(cell, top=60, bottom=60, left=80, right=80):
    cell._tc.get_or_add_tcPr().append(
        parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f"</w:tcMar>"
        )
    )


def _nil_cell_borders(cell):
    cell._tc.get_or_add_tcPr().append(
        parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="nil"/><w:left w:val="nil"/>'
            f'<w:bottom w:val="nil"/><w:right w:val="nil"/>'
            f"</w:tcBorders>"
        )
    )


def _borders(cell, color="E2E8F0", sz="4", left_sz=None, left_color=None):
    lc = left_color or color
    ls = left_sz or sz
    cell._tc.get_or_add_tcPr().append(
        parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="{ls}" w:space="0" w:color="{lc}"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f"</w:tcBorders>"
        )
    )


def _nil_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tbl_pr.append(borders)


def _run(run, size, bold=False, italic=False, color=DARK, font="Arial", tracking=None, lang="pl-PL"):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font)
    if tracking is not None:
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), str(tracking))
        rpr.append(spacing)
    lang_el = rpr.find(qn("w:lang"))
    if lang_el is None:
        lang_el = OxmlElement("w:lang")
        rpr.append(lang_el)
    lang_el.set(qn("w:val"), lang)
    lang_el.set(qn("w:eastAsia"), lang)


def _lock(table, widths, align="center"):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER if align == "center" else WD_TABLE_ALIGNMENT.LEFT
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    for tag in ("w:tblLayout", "w:tblW", "w:jc", "w:tblInd"):
        found = tbl_pr.find(qn(tag))
        if found is not None:
            tbl_pr.remove(found)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    dxas = [int(round(w / 2.54 * 1440)) for w in widths]
    total = sum(dxas)
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_pr.append(tbl_w)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), align)
    tbl_pr.append(jc)
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    tbl_pr.append(ind)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for dxa in dxas:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(dxa))
        grid.append(col)
    for row in table.rows:
        tr = row._tr
        tr_pr = tr.get_or_add_trPr()
        for cell, dxa in zip(row.cells, dxas):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(dxa))
            tc_w.set(qn("w:type"), "dxa")
        # Drop Word's implicit extra grid cell if merge changed the count.
        _ = tr_pr


def _row_height(row, twips, rule="exact"):
    tr_pr = row._tr.get_or_add_trPr()
    el = tr_pr.find(qn("w:trHeight"))
    if el is None:
        el = OxmlElement("w:trHeight")
        tr_pr.append(el)
    el.set(qn("w:val"), str(twips))
    el.set(qn("w:hRule"), rule)


def _cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def _field(paragraph, instruction, size=8, color=MUTED):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)
    _run(run, size, False, False, color)


def _p_border_bottom(paragraph, color="E6EDF2", sz="8", space="6"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), space)
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _no_snap(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    snap = p_pr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        p_pr.append(snap)
    snap.set(qn("w:val"), "0")


def _widow(paragraph, keep_next=False, keep_lines=False, page_break=False):
    fmt = paragraph.paragraph_format
    fmt.widow_control = True
    fmt.keep_with_next = keep_next
    fmt.keep_together = keep_lines
    if page_break:
        fmt.page_break_before = True
    _no_snap(paragraph)


def _clear_paragraph(paragraph):
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    _no_snap(paragraph)


def _banner():
    image = Image.open(SHOT).convert("RGB")
    crop = image.crop((96, 770, 848, 1044))
    stream = BytesIO()
    crop.save(stream, format="PNG", optimize=True)
    stream.seek(0)
    return stream


def _cell_p(cell, clear=True):
    paragraph = cell.paragraphs[0]
    if clear:
        _clear_paragraph(paragraph)
    return paragraph


def _add_body(cell, text, size=10, bold=False, italic=False, color=DARK, before=0, after=2, align="left", lang="pl-PL"):
    paragraph = cell.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.08
    if align == "justify":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "right":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(paragraph.add_run(text), size, bold, italic, color, lang=lang)
    return paragraph


FINDINGS = [
    (
        "01",
        "Kalkulator CO₂ „6,060 TONS” i obietnica zgodności z polityką ESG",
        "UZNK art. 16 ust. 1 pkt 2, ust. 2 i ust. 3",
        "BARDZO WYSOKIE",
        "Ilość na stronie angielskiej może być odczytana jako tysiące ton. Zmiana palet jest przedstawiona jako zgodność z polityką ESG.",
    ),
    (
        "02",
        "„There is no waste”, odpowiedzialność społeczna i „compliance” z Agendą 2030",
        "UZNK art. 16 ust. 1 pkt 2 i ust. 2",
        "WYSOKIE",
        "Brak odpadów, najwyższe standardy i zgodność z celami ONZ bez wskazania zakresu.",
    ),
    (
        "03",
        "„99%” przez cały cykl życia oraz „you are ESG Friendly”",
        "UZNK art. 16 ust. 1 pkt 2 i ust. 2",
        "WYSOKIE",
        "Wynik dwóch szczepów według ISO 22196 rozszerzony na cały cykl życia, pierwszeństwo na świecie i cechę ESG nabywcy.",
    ),
]

CASES = [
    {
        "no": "01",
        "title": "Kalkulator emisji i zdanie, że zmiana palet daje zgodność z polityką ESG",
        "risk": "BARDZO WYSOKIE",
        "basis": (
            "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2, ust. 2 i ust. 3 "
            "(Dz.U. 2026 poz. 85). Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4b, a dla hasła ogólnego i pieczęci ESG "
            "także pkt 4a i 2a ustawy z 4.09.2026 r. (Dz.U. 2026 poz. 1353). W dniu audytu nie jest to naruszenie krajowej czarnej listy."
        ),
        "quotes": [
            "„Thanks to simulations and calculations carried out in accordance with the EN 16258:2012 standard for measuring the carbon footprint […] The reduction of CO2 emissions to the environment with the use of 33 LIBRA PARTNERS pallets per month is 400 kg.”",
            "„500 PCS – 6,060 TONS LESS CO2”. „10.000 PCS – 121,120 TONS LESS CO2.”",
            "„Make your company compliant with the Sustainable Development and ESG policy. Prepare your company for non-financial ESG reporting now. Changing pallets from wooden to plastic LIBRA PARTNERS is the simplest solution to significantly reduce CO2 emissions and carbon footprint into the atmosphere.”",
        ],
        "assessment": (
            "Art. 16 ust. 3 uznaje za reklamę porównawczą reklamę, która pozwala rozpoznać towar konkurenta. "
            "Zdanie o śladzie stawia paletę LIBRA PARTNERS naprzeciw „wooden pallets”, czyli kategorii nośnika o tym samym przeznaczeniu. "
            "Art. 16 ust. 1 pkt 2 obejmuje reklamę, która wprowadza klienta w błąd i może wpłynąć na decyzję o nabyciu. "
            "Ust. 2 każe uwzględnić wszystkie elementy, zwłaszcza ilość. Klientem zdania „your company” jest nabywca biznesowy.\n"
            "Na tej samej stronie 33 palety mają obniżać emisję o 400 kg miesięcznie. Przeliczenie liniowe na 500 palet daje około 6,06 tony. "
            "Tabela zapisuje „6,060 TONS”, a dla 10 000 sztuk „121,120 TONS”. W tekście angielskim przecinek bywa separatorem tysięcy, a jednostką jest słowo TONS. "
            "Czytelnik tej wersji może odczytać 6 060 ton. Strona nie podaje, że przecinek jest znakiem dziesiętnym. "
            "Ta ilość może zmienić decyzję o zakupie i o wpisaniu liczby do materiału dla klienta. "
            "To samo miejsce nazywa zmianę palet najprostszą drogą do zgodności z polityką ESG i do raportowania niefinansowego.\n"
            "Strona nazywa EN 16258:2012 normą pomiaru tego śladu. Karta projektu CEN dla EN ISO 14083:2023 podaje, że norma ta zastępuje EN 16258:2012, "
            "a data wycofania normy sprzecznej to 31 października 2023 r. Zakres EN 16258 dotyczył energii i emisji usługi transportowej. "
            "Strona nie podaje trasy, ładunku, jednostki funkcjonalnej ani granic. Obok stoją „almost twice lower CO2” z analizy Envirly z sierpnia 2023 r. "
            "oraz „303 t CO2 less per month”, bez uzgodnienia z kalkulatorem. Oferta jest skierowana do firmy, więc podstawą spełnioną jest art. 16.\n"
            "Od 27 września 2027 r. twierdzenie o całym produkcie oparte na masie albo na jednym wyniku transportowym trzeba będzie zestawić z pkt 4b, "
            "hasło „ecological” i tytuł „Eco-friendly” z pkt 4a, a pieczęć ESG z pkt 2a, jeżeli nie opiera się na systemie certyfikacji ani na organie publicznym. "
            "Przekaz nie powołuje kompensacji, więc pkt 4c go nie opisuje."
        ),
        "remediation": (
            "Zdjąć kalkulator i zdanie o zgodności z polityką ESG oraz o przygotowaniu do raportowania niefinansowego do czasu publikacji liczby z jednostką, granicami i metodą. "
            "Jeżeli porównanie transportowe zostaje, podać trasę albo tonokilometr, masę ładunku, masę palety odniesienia, granice oraz normę aktualną dla usługi transportowej. "
            "Liczbę zapisać tak, by czytelnik angielski widział tony dziesiętne, na przykład „6.06 tonnes”, albo podać wynik w kilogramach. "
            "Wynik Envirly przywołać z datą, jednostką funkcjonalną i granicami albo usunąć „almost twice”. Pieczęć ESG zdjąć albo opisać na tym samym slajdzie, kto i za co ją nadał."
        ),
        "claim": (
            "„A LIBRA PARTNERS pallet in this comparison weighs 9 kg and is made from recycled HDPE. "
            "For a stated route, load and fuel, the difference in transport fuel emissions against a wooden pallet of a stated weight is [X] kg CO2e per trip, "
            "calculated under [method, year, boundaries]. This figure is not a company carbon footprint, not a CSRD disclosure and not a statement that the buyer’s ESG policy is met.”"
        ),
    },
    {
        "no": "02",
        "title": "Brak odpadów, odpowiedzialność społeczna i zgodność z Agendą 2030",
        "risk": "WYSOKIE",
        "basis": (
            "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2 i ust. 2. "
            "Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a i 4b ustawy z 4.09.2026 r. W dniu audytu nie jest to naruszenie krajowej czarnej listy. "
            "Ścieżka Krajowego Punktu Kontaktowego OECD ma poziom NISKIE i nie podnosi werdyktu."
        ),
        "quotes": [
            "„In this business model, there is no waste; a Closed Circuit Economy model is present. Likewise, as a socially responsible and environmentally conscious company, we want our production process to be distinguished by the highest standards in environmental management. We do this through a business model focused on ZERO WASTE.”",
            "„Eliminate end-of-life waste by recycling our solutions, which never end up in a landfill.”",
            "„Selected goals of the sustainable development policy, confirming the compliance of our project with the goals of the 2030 Agenda.”",
        ],
        "assessment": (
            "Hasło o firmie odpowiedzialnej społecznie, o najwyższych standardach i o zgodności z celami ONZ odbiorca może wziąć za fakt szerszy niż mechanizm opisany obok. "
            "Mechanizm na stronie jest jeden: surowiec ma być przetwarzalny, a firma gwarantuje odbiór zużytych i uszkodzonych palet kupionych od LIBRA PARTNERS. "
            "„100% CIRCULAR ECONOMY”, „ZERO WASTE”, „there is no waste” i „never end up in a landfill” obejmują cały model i koniec życia każdej sztuki. "
            "Gwarancja odbioru obejmuje palety kupione od tej firmy i zwrócone. „Highest standards” i „socially responsible” nie wskazują zakładu, okresu ani systemu. "
            "Strona nie powołuje w tym miejscu ISO 14001.\n"
            "„Confirming the compliance” stawia cytat celów 9.4, 12.2, 12.5 i 15.2 jako potwierdzenie zgodności projektu. "
            "Cele te są rezolucją Zgromadzenia Ogólnego ONZ dla państw. Cel 15.2 w cytacie ma horyzont „By 2020”. "
            "Klient szukający dostawcy do polityki obiegu zamkniętego może na tych zdaniach oprzeć wybór. Art. 16 ust. 1 pkt 2 i ust. 2 obejmują taką reklamę.\n"
            "Od 27 września 2027 r. ogólne „ZERO WASTE” i „100% CIRCULAR” trzeba będzie zestawić z pkt 4a, a claim o całej działalności oparty na odbiorze własnych palet z pkt 4b. "
            "Hasło odpowiedzialności społecznej może trafić do polskiego Krajowego Punktu Kontaktowego OECD. Postępowanie przed KPK nie jest sprawą sądową i nie kończy się wyrokiem. "
            "Skutkiem jest publiczna ocena, zalecenia i efekt wizerunkowy. Ten poziom jest NISKIE. Werdykt przypadku niesie art. 16."
        ),
        "remediation": (
            "Zastąpić „there is no waste”, „ZERO WASTE”, „100% CIRCULAR ECONOMY” i „never end up in a landfill” opisem odbioru palet kupionych od LIBRA PARTNERS, "
            "z tym, co dzieje się ze sztuką nieoddaną i z odpadem procesu. „Socially responsible” i „highest standards” zawęzić do nazwanego zakładu, roku i dokumentu albo usunąć. "
            "Przy Agendzie 2030 napisać, które działania firma odnosi do którego celu, i usunąć słowo „compliance”, dopóki nie ma oceny zewnętrznej."
        ),
        "claim": (
            "„LIBRA PARTNERS collects used and damaged pallets that the customer bought from LIBRA PARTNERS and recycles that returned material. "
            "This take-back does not describe all waste from the plant. The company refers its pallet design to UN Sustainable Development Goals 12.2 and 12.5 as a voluntary framework. "
            "That reference is not a certification and not a finding that the project complies with the 2030 Agenda.”"
        ),
    },
    {
        "no": "03",
        "title": "Właściwość antybakteryjna „przez cały cykl życia” oraz „ESG Friendly”",
        "risk": "WYSOKIE",
        "basis": (
            "Podstawa prawna: ustawa o zwalczaniu nieuczciwej konkurencji, art. 16 ust. 1 pkt 2 i ust. 2. "
            "Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a dla „ESG Friendly” oraz pkt 4b, jeżeli cecha całej palety jest wnioskiem z badania powierzchni. "
            "W dniu audytu nie jest to naruszenie krajowej czarnej listy."
        ),
        "quotes": [
            "„we are the world’s first manufacturer of pallets adapted to the pharmaceutical and medical industries, which have 99% antibacterial properties throughout the entire life cycle of the pallet. In addition […] you are ESG Friendly!”",
            "„tested according to ISO 22196:2011: DECREASE OF ESCHERICHIA COLI SAMPLES ( ATCC 8739 ) ABOUT 99,97 %. DECREASE OF STAPHYLOCOCCUS AUREUS SAMPLES ( ATCC 6538P ) ABOUT 99,98 %.”",
            "„The LIBRA PARTNERS antibacterial palette model MW01-MED received the opinion of the Poznań University of Technology for innovation on a global scale.”",
        ],
        "assessment": (
            "Strona podaje wynik: spadek liczebności próbek dwóch szczepów, około 99,97% i około 99,98%, według ISO 22196:2011, "
            "po współpracy z certyfikowaną jednostką sieci Łukasiewicz, przy dodatku tlenku cynku do HDPE. "
            "W tym samym widoku pisze, że paleta ma „99% antibacterial properties throughout the entire life cycle”, że producent jest „the world’s first” "
            "i że używając palety „you are ESG Friendly”.\n"
            "Wynik dwóch szczepów, zapisany jako „about”, nie jest zdaniem o całym cyklu życia palety po myciu i po ścieraniu. "
            "„99%” zaokrągla dwa różne wyniki i obejmuje już właściwość palety jako takiej. „World’s first” jest podane jako fakt. "
            "Strona nie wskazuje przeglądu rynku. Raport nie stwierdza, że pierwszeństwo jest fałszywe. Stwierdza, że reklama podaje je bez źródła, które klient może sprawdzić. "
            "Opinia Politechniki Poznańskiej jest na stronie opinią o innowacyjności. Postawienie jej w jednym bloku z cyklem życia i ze zdaniem „ESG Friendly” "
            "pozwala czytać ją jako potwierdzenie tych dwóch zdań.\n"
            "Nabywcą jest firma farmaceutyczna albo medyczna. Zdanie o cyklu życia może zdecydować o wyborze nośnika do strefy czystej. "
            "Art. 16 ust. 1 pkt 2 i ust. 2 obejmują jakość i przydatność. Liczb 99,97% i 99,98% raport nie uznaje za nieprawdziwe. "
            "Podstawą ustalenia jest różnica między tymi liczbami a zdaniami o cyklu życia, o pierwszeństwie i o cesze ESG nabywcy. "
            "Od 27 września 2027 r. „ESG Friendly” trzeba będzie zestawić z pkt 4a, a wniosek o całej palecie wyprowadzony z badania powierzchni z pkt 4b."
        ),
        "remediation": (
            "Zostawić nazwę szczepów, normę, rząd wielkości spadku i jednostkę, która badała. Dopisać warunki badania, które są w raporcie laboratorium. "
            "Usunąć „throughout the entire life cycle”, „99%” jako cechę całej palety, „world’s first” i „you are ESG Friendly”, "
            "dopóki osobne dowody nie obejmują cyklu życia, przeglądu rynku i cechy ESG. Opinię politechniki cytować jako opinię o innowacyjności, w akapicie osobnym od wyniku mikrobiologicznego."
        ),
        "claim": (
            "„MW01-MED is recycled HDPE with a zinc oxide addition. In a test according to ISO 22196:2011, run with a certified Łukasiewicz unit, "
            "samples showed a reduction of about 99.97% for E. coli ATCC 8739 and about 99.98% for S. aureus ATCC 6538P, under the contact conditions stated in the test report. "
            "The page does not extend this result to the whole service life of the pallet. Poznań University of Technology issued an opinion on the innovation. That opinion is not an ESG certification.”"
        ),
    },
]


def build(path: Path, specimen: bool):
    doc = docx.Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = DARK
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), "Arial")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Cm(2.55)
    section.bottom_margin = Cm(1.95)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.header_distance = Cm(0.42)
    section.footer_distance = Cm(0.45)
    grid = section._sectPr.find(qn("w:docGrid"))
    if grid is None:
        grid = OxmlElement("w:docGrid")
        section._sectPr.append(grid)
    grid.set(qn("w:type"), "default")
    grid.set(qn("w:linePitch"), "360")

    settings = doc.settings.element
    for el in settings.findall(".//" + qn("w:compatSetting")):
        if el.get(qn("w:name")) == "compatibilityMode":
            el.set(qn("w:val"), "15")
    theme_lang = settings.find(qn("w:themeFontLang"))
    if theme_lang is not None:
        theme_lang.set(qn("w:val"), "pl-PL")
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings.append(compat)
    compat.append(OxmlElement("w:doNotUseHTMLParagraphAutoSpacing"))

    props = doc.core_properties
    props.author = "ESG Institute"
    props.category = "Audyt greenwashingu"
    props.subject = "Raport skrócony audytu ryzyk greenwashingu — LIBRA PARTNERS"
    props.title = (
        "Wzór raportu skróconego — LIBRA PARTNERS"
        if specimen
        else "Raport skrócony audytu ryzyk greenwashingu — LIBRA PARTNERS"
    )
    props.comments = "ESG Institute · esginstitute.eu · raport skrócony · 30 października 2026 r."

    def para(
        text,
        size=10.5,
        bold=False,
        italic=False,
        color=DARK,
        before=0,
        after=8,
        align="left",
        keep_next=False,
        keep_lines=False,
        page_break=False,
        left=0,
        right=0,
        first=0,
        lang="pl-PL",
        tracking=None,
        space=1.15,
    ):
        paragraph = doc.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(before)
        fmt.space_after = Pt(after)
        fmt.line_spacing = space
        if left:
            fmt.left_indent = Cm(left)
        if right:
            fmt.right_indent = Cm(right)
        if first:
            fmt.first_line_indent = Cm(first)
        if align == "justify":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _widow(paragraph, keep_next, keep_lines, page_break)
        _run(paragraph.add_run(text), size, bold, italic, color, lang=lang, tracking=tracking)
        return paragraph

    def rich(
        parts,
        size=10.5,
        before=0,
        after=8,
        align="left",
        keep_next=False,
        keep_lines=False,
        page_break=False,
        left=0,
        first=0,
        space=1.15,
    ):
        paragraph = doc.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(before)
        fmt.space_after = Pt(after)
        fmt.line_spacing = space
        if left:
            fmt.left_indent = Cm(left)
        if first:
            fmt.first_line_indent = Cm(first)
        if align == "justify":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _widow(paragraph, keep_next, keep_lines, page_break)
        for text, bold, italic, color, lang in parts:
            _run(paragraph.add_run(text), size, bold, italic, color, lang=lang)
        return paragraph

    def body(text, size=10.5, color=DARK, before=2, after=8, keep_next=False, keep_lines=True, first=0.75, lang="pl-PL"):
        return para(
            text,
            size=size,
            color=color,
            before=before,
            after=after,
            align="justify",
            keep_next=keep_next,
            keep_lines=keep_lines,
            first=first,
            lang=lang,
        )

    def start_page():
        paragraph = doc.add_paragraph()
        _clear_paragraph(paragraph)
        _widow(paragraph, page_break=True)
        spacing = paragraph._p.get_or_add_pPr().find(qn("w:spacing"))
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "40")
        spacing.set(qn("w:lineRule"), "exact")
        return paragraph

    def heading(number, title, new_page=False):
        paragraph = doc.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(0 if new_page else 18)
        fmt.space_after = Pt(10)
        fmt.line_spacing = 1.05
        _widow(paragraph, keep_next=True, keep_lines=True, page_break=new_page)
        _run(paragraph.add_run(number + ".  " + title), 14, True, False, NAVY)
        _p_border_bottom(paragraph, "D5DDE3", "12", "8")
        return paragraph

    def label(text):
        return para(text, size=9, bold=True, color=BLUE, before=14, after=5, keep_next=True, keep_lines=True, tracking=32)

    def band(fill, left_color, paragraphs, gap_after=8, new_page=False):
        table = doc.add_table(rows=1, cols=2)
        _lock(table, [0.18, PAGE_W - 0.18])
        _nil_table_borders(table)
        rail, body_cell = table.rows[0].cells
        _shd(rail, left_color)
        _shd(body_cell, fill)
        _mar(rail, 0, 0, 0, 0)
        _mar(body_cell, 100, 100, 160, 150)
        _borders(rail, left_color, "2", "2", left_color)
        _borders(body_cell, fill, "2", "2", fill)
        _clear_paragraph(rail.paragraphs[0])
        first = True
        for text, size, bold, italic, color, lang in paragraphs:
            if first:
                p = body_cell.paragraphs[0]
                _clear_paragraph(p)
                first = False
            else:
                p = body_cell.add_paragraph()
                _no_snap(p)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if (bold or italic) else WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.widow_control = True
            if new_page and p is body_cell.paragraphs[0]:
                p.paragraph_format.page_break_before = True
            _run(p.add_run(text), size, bold, italic, color, lang=lang)
        if body_cell.paragraphs:
            body_cell.paragraphs[-1].paragraph_format.space_after = Pt(0)
        _cant_split(table.rows[0])
        gap = doc.add_paragraph()
        _clear_paragraph(gap)
        gap.paragraph_format.space_after = Pt(gap_after)

    # --- nagłówek i stopka ---
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    _clear_paragraph(hp)
    htable = header.add_table(2, 2, Cm(PAGE_W))
    _lock(htable, [8.4, PAGE_W - 8.4])
    _nil_table_borders(htable)
    left, right = htable.rows[0].cells
    for cell in (left, right):
        _shd(cell, "08181F")
        _mar(cell, 80, 70, 140, 140)
        _nil_cell_borders(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    lp = _cell_p(left)
    if LOGO.exists():
        lp.add_run().add_picture(str(LOGO), width=Cm(4.15))
    else:
        _run(lp.add_run("ESG INSTITUTE"), 11, True, False, WHITE)
    rp = _cell_p(right)
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    who = "WZÓR FORMY" if specimen else "LIBRA PARTNERS"
    _run(rp.add_run("Raport skrócony"), 8, True, False, CYAN_SOFT)
    rp2 = right.add_paragraph()
    _clear_paragraph(rp2)
    rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run(rp2.add_run(who), 8, True, False, WHITE)
    _run(rp2.add_run("   ·   esginstitute.eu"), 8, False, False, CYAN_SOFT)

    stripe_a, stripe_b = htable.rows[1].cells
    stripe_a.merge(stripe_b)
    stripe = htable.rows[1].cells[0]
    _shd(stripe, "2F8DBC")
    _mar(stripe, 0, 0, 0, 0)
    _nil_cell_borders(stripe)
    _clear_paragraph(stripe.paragraphs[0])
    _row_height(htable.rows[1], 70, "exact")

    # Usuń pusty akapit nad tabelą nagłówka, żeby pas przylegał do krawędzi.
    hp_el = hp._p
    if hp_el.getparent() is not None and not hp_el.xpath(".//w:drawing"):
        hp_el.getparent().remove(hp_el)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    _clear_paragraph(fp)
    fp.paragraph_format.space_before = Pt(4)
    p_pr = fp._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(int(round(PAGE_W / 2.54 * 1440))))
    tabs.append(tab)
    p_pr.append(tabs)
    p_bdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "12")
    top.set(qn("w:space"), "8")
    top.set(qn("w:color"), "D5DDE3")
    p_bdr.append(top)
    p_pr.append(p_bdr)
    edition = "wzór formy  ·  " if specimen else ""
    _run(
        fp.add_run("ESG Institute Sp. z o.o.   ·   esginstitute.eu   ·   " + edition + "raport skrócony   ·   Libra Partners"),
        8,
        False,
        False,
        MUTED,
    )
    _run(fp.add_run("\t"), 8, False, False, MUTED)
    _field(fp, "PAGE")
    _run(fp.add_run("  /  "), 8, False, False, MUTED)
    _field(fp, "NUMPAGES")

    # --- treść ---
    if specimen:
        para(
            "WZÓR FORMY. Układ i identyfikacja ESG Institute służą kolejnym raportom skróconym. "
            "Przykład dotyczy LIBRA PARTNERS. Ustalenia nie są oceną innego podmiotu.",
            size=8.5,
            italic=True,
            color=BLUE,
            before=0,
            after=4,
        )

    para(
        "ESG INSTITUTE   ·   RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO",
        size=8.5,
        bold=True,
        color=BLUE,
        before=4,
        after=4,
        keep_next=True,
        keep_lines=True,
        tracking=18,
    )
    para(
        "Audyt Ryzyk Greenwashingu: LIBRA PARTNERS",
        size=20,
        bold=True,
        color=NAVY,
        before=0,
        after=8,
        keep_next=True,
        keep_lines=True,
        space=1.05,
    )

    para(
        "Raport skrócony. Weryfikacja angielskiej witryny https://libra-partners.pl/en/ oraz podstron zrównoważonego rozwoju, "
        "„o nas”, palety MW01 i palety antybakteryjnej MW01-MED, stan na 30 października 2026 r. "
        "Werdykt opiera się na ustawie o zwalczaniu nieuczciwej konkurencji. "
        "Czarna lista dyrektywy (UE) 2024/825 (EmpCo) jest przyszłym standardem od 27 września 2027 r.",
        size=10.5,
        italic=True,
        color=MUTED,
        before=2,
        after=12,
        space=1.2,
    )

    frame = doc.add_table(rows=1, cols=1)
    _lock(frame, [PAGE_W])
    _nil_table_borders(frame)
    fcell = frame.cell(0, 0)
    _shd(fcell, "FFFFFF")
    _mar(fcell, 60, 40, 70, 70)
    _borders(fcell, "E4E8EC", "6")
    fpict = _cell_p(fcell)
    fpict.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fpict.add_run().add_picture(_banner(), width=Cm(PAGE_W - 0.7))
    cap = fcell.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(0)
    cap.paragraph_format.line_spacing = 1.05
    _run(
        cap.add_run(
            "Fragment slajdu „Carbon footprint” na https://libra-partners.pl/en/, 30 października 2026 r. "
            "Hasło porównuje palety LIBRA PARTNERS z paletami drewnianymi. Na pełnym slajdzie, na prawo od tego kadru, jest okrągła pieczęć z napisem ESG."
        ),
        8,
        False,
        True,
        MUTED,
    )
    after_frame = doc.add_paragraph()
    _clear_paragraph(after_frame)
    after_frame.paragraph_format.space_after = Pt(12)

    meta = [
        (
            "Badany podmiot",
            "LIBRA PARTNERS, grupa SZARPOL. Kontakt na stronie: Rafał Szarzała, Centrum Produkcyjno-Handlowe „Szarpol”, ul. prof. Mariana Rapackiego 3, 98-200 Sieradz. Założyciele podani na stronie: Maciej Włodarczyk i Rafał Szarzała.",
        ),
        (
            "Przedmiot audytu",
            "Angielska strona główna oraz podstrony /en/zrownowazony-rozwoj/, /en/o-nas/, /en/standard-pallet-model-mw01/ i /en/paleta-antybakteryjna-model-mw01-med/. Wersja polska witryny istnieje i nie jest przedmiotem tego raportu.",
        ),
        (
            "Sektor działalności",
            "Palety transportowe z HDPE, w tym z recyklatu. Oferta skierowana do firm: logistyka, przemysł spożywczy, farmaceutyczny i medyczny.",
        ),
        (
            "Reżim prawny",
            "Prawo obowiązujące: ustawa o zwalczaniu nieuczciwej konkurencji, Dz.U. 2026 poz. 85, art. 16 i art. 18. Gdy odbiorcą jest konsument: ustawa o przeciwdziałaniu nieuczciwym praktykom rynkowym, Dz.U. 2023 poz. 845, art. 5 i 6. Ustawa o ochronie konkurencji i konsumentów, Dz.U. 2025 poz. 1714, art. 24 i art. 106 ust. 1 pkt 4. Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) wchodzi w życie 27 września 2027 r. W dniu audytu czarna lista EmpCo nie jest prawem polskim.",
        ),
        ("Data audytu", "30 października 2026 r."),
        ("Werdykt operacyjny", "Nie publikować w obecnej formie"),
    ]
    meta_table = doc.add_table(rows=len(meta), cols=2)
    meta_w = [3.9, PAGE_W - 3.9]
    _lock(meta_table, meta_w)
    _nil_table_borders(meta_table)
    for i, (field, value) in enumerate(meta):
        ck, cv = meta_table.rows[i].cells
        last = i == len(meta) - 1
        _shd(ck, "F4F7F8")
        _shd(cv, "FFFFFF")
        _mar(ck, 70, 70, 120, 80)
        _mar(cv, 70, 70, 120, 120)
        _borders(ck, "E4E8EC", "4")
        _borders(cv, "E4E8EC", "4")
        ck.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pk, pv = _cell_p(ck), _cell_p(cv)
        pk.paragraph_format.space_after = Pt(0)
        pv.paragraph_format.space_after = Pt(0)
        pv.paragraph_format.line_spacing = 1.12
        pv.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _run(pk.add_run(field), 8, True, False, NAVY)
        _run(pv.add_run(value), 9, last, False, DARK)
        _cant_split(meta_table.rows[i])

    start_page()
    band(
        "FDF2F4",
        "BE123C",
        [
            ("Werdykt. Nie publikować w obecnej formie.", 12, True, False, NAVY, "pl-PL"),
            (
                "Strona sprzedaje paletę firmie i w tym samym miejscu podaje ilość emisji oraz obietnicę zgodności z polityką ESG. "
                "Art. 16 ust. 1 pkt 2 i ust. 2 ustawy o zwalczaniu nieuczciwej konkurencji obejmują reklamę, która wprowadza klienta w błąd co do ilości i może wpłynąć na decyzję o nabyciu. "
                "Zapis „6,060 TONS” przy podstawie „400 kg” dla 33 palet spełnia tę przesłankę. Do tego dochodzą zdania o braku odpadów i o zgodności z Agendą 2030 oraz rozszerzenie wyniku badania dwóch szczepów na cały cykl życia palety.",
                10.5,
                False,
                False,
                DARK,
                "pl-PL",
            ),
            (
                "Ustawa wdrażająca dyrektywę (UE) 2024/825 wchodzi w życie 27 września 2027 r. Do tej daty ogólne „eco-friendly”, „key player in sustainability” i pieczęć ESG są przyszłym standardem pkt 2a i 4a, a claim o całym produkcie oparty na masie albo na odbiorze własnych palet jest przyszłym standardem pkt 4b. Przekaz nie opiera redukcji na kompensacji, więc pkt 4c go nie opisuje.",
                10.5,
                False,
                False,
                DARK,
                "pl-PL",
            ),
        ],
        gap_after=6,
    )

    heading("1", "Analiza wstępna")
    body(
        "LIBRA PARTNERS przedstawia się jako producent palet z HDPE w grupie SZARPOL i kieruje ofertę do firm logistycznych, spożywczych, farmaceutycznych i medycznych. "
        "Wersja angielska otwiera się porównaniem śladu węglowego z paletą drewnianą, hasłem gospodarki o obiegu zamkniętym oraz opiniami nazwanymi „Consumer feedback”. "
        "Odbiorcą zdania „your company” jest klient biznesowy, więc podstawą już dziś jest ustawa o zwalczaniu nieuczciwej konkurencji. "
        "Trzy twierdzenia spełniają jej przesłanki i rozstrzygają werdykt. Hasła ogólne i pieczęć ESG zbliżają się do zakazów, które w Polsce zaczną obowiązywać 27 września 2027 r. Same nie dają werdyktu czerwonego."
    )

    heading("2", "Ustalenia")
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota"]
    col_w = [1.05, 4.35, 3.55, 2.95, 5.1]
    sum_table = doc.add_table(rows=1 + len(FINDINGS), cols=5)
    _lock(sum_table, col_w)
    _nil_table_borders(sum_table)
    for c, head in enumerate(headers):
        cell = sum_table.rows[0].cells[c]
        _shd(cell, "08181F")
        _mar(cell, 70, 70, 70, 70)
        _borders(cell, "08181F", "4", "4", "08181F")
        p = _cell_p(cell)
        if c in (0, 3):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p.add_run(head), 8, True, False, WHITE)
    risk_color = {"BARDZO WYSOKIE": DANGER, "WYSOKIE": WARNING}
    for r, row in enumerate(FINDINGS):
        bg = "FFFFFF" if r % 2 == 0 else "F7F9FC"
        for c, val in enumerate(row):
            cell = sum_table.rows[r + 1].cells[c]
            _shd(cell, bg)
            _mar(cell, 70, 70, 70, 70)
            _borders(cell, "E4E8EC", "4")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = _cell_p(cell)
            p.paragraph_format.line_spacing = 1.12
            if c in (0, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if c == 0:
                _run(p.add_run(val), 8.5, True, False, NAVY)
            elif c == 3:
                _run(p.add_run(val), 8, True, False, risk_color.get(val, DARK))
            else:
                _run(p.add_run(val), 8, False, False, DARK)
        _cant_split(sum_table.rows[r + 1])
    _cant_split(sum_table.rows[0])
    sum_table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    para(
        "Jest to przykład 3 ryzyk spośród 5 zidentyfikowanych na stronie.",
        size=10,
        italic=True,
        color=MUTED,
        before=12,
        after=2,
    )

    heading("3", "Szczegóły przypadków", new_page=True)
    case = CASES[0]
    para(
        "Przypadek " + case["no"] + ".  " + case["title"],
        size=13,
        bold=True,
        color=NAVY,
        before=12,
        after=4,
        keep_next=True,
        keep_lines=True,
        space=1.08,
    )
    risk_rgb = DANGER if "BARDZO" in case["risk"] else WARNING
    rich(
        [
            ("Ryzyko: ", False, False, MUTED, "pl-PL"),
            (case["risk"].lower() + ".", True, False, risk_rgb, "pl-PL"),
        ],
        size=10.5,
        before=0,
        after=6,
        keep_next=True,
        keep_lines=True,
    )
    body(case["basis"], size=9.5, before=0, after=4, first=0, keep_lines=True)

    label("Cytat źródłowy")
    for quote in case["quotes"]:
        band("F4F7FB", "2F8DBC", [(quote, 10, False, True, DARK, "en-US")], gap_after=6)

    label("Ocena prawna")
    for block in case["assessment"].split("\n"):
        body(block, before=1, after=8, keep_lines=True, first=0.75)

    naprawa = para(
        "Kierunek naprawy",
        size=14,
        bold=True,
        color=NAVY,
        before=0,
        after=8,
        keep_next=True,
        keep_lines=True,
        page_break=True,
        space=1.05,
    )
    _p_border_bottom(naprawa, "D5DDE3", "12", "8")
    body(case["remediation"], before=4, after=6, keep_lines=True, first=0.75)

    label("Dozwolone sformułowanie")
    band("E8F0FA", "3A61B3", [(case["claim"], 10, False, True, DARK, "en-US")], gap_after=4)

    pozostale = para(
        "Pozostałe przypadki",
        size=14,
        bold=True,
        color=NAVY,
        before=0,
        after=10,
        keep_next=True,
        keep_lines=True,
        page_break=True,
        space=1.05,
    )
    _p_border_bottom(pozostale, "D5DDE3", "12", "8")
    for other in CASES[1:]:
        other_risk = DANGER if "BARDZO" in other["risk"] else WARNING
        band(
            "F7F9FC",
            "08181F",
            [
                (other["no"] + "     " + other["title"], 11, True, False, NAVY, "pl-PL"),
                ("Ryzyko: " + other["risk"].lower() + ".", 10, True, False, other_risk, "pl-PL"),
            ],
            gap_after=8,
        )
    band(
        "F4F7FB",
        "3A61B3",
        [
            ("Reszta po zakupie pełnej wersji.", 11, True, False, NAVY, "pl-PL"),
            (
                "Pełna analiza wszystkich przypadków naruszeń wraz z rekomendacjami naprawczymi dostępna w pełnym raporcie.",
                10.5,
                False,
                False,
                DARK,
                "pl-PL",
            ),
        ],
        gap_after=0,
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return path


def main():
    wzor = build(ROOT / "wzory" / "WZOR_raport_skrocony.docx", specimen=True)
    raport = build(ROOT / "LIBRA" / "Raport_Skrocony_Audytu_Greenwashingu_Libra_Partners.docx", specimen=False)
    print(wzor)
    print(raport)


if __name__ == "__main__":
    main()
