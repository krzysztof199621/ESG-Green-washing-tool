#!/usr/bin/env python3
"""Raport audytu greenwashingu i social washingu: Leśne Bycie (lesnebycie.pl).

Forma według INSTRUKCJA_GENEROWANIA_RAPORTOW.md oraz wzorca
wzory/WZOR_raport_Libra_Partners.docx. Treść dotyczy wyłącznie tej witryny.
"""

from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
LOGO = Path(
    "/Users/krzysztofsulinski/Desktop/AI Projects/ESG/"
    "ESG Experts Day - Komplet Materialow (PDF + PPTX)/"
    "03_DLA_PROWADZACEGO/assets/esg_logo_white.png"
)
SHOT = ROOT / "assets" / "lesnebycie_strona_naglowek.png"


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

    p("ESG INSTITUTE – Zestaw treści i deklaracji marketingowych do audytu", 16, True, False, NAVY, after=2)
    p(
        "Zbiór haseł, opisów oferty, wpisów blogowych i polityki prywatności zebranych ze strony lesnebycie.pl. "
        "Plik jest zapisem źródłowym do audytu greenwashingu i social washingu. Nie zawiera oceny prawnej.",
        10.5, False, True, MUTED, after=8,
    )

    rows = [
        ("Podmiot w warstwie strony", "Marka „Leśne Bycie”. Osoba przedstawiana na stronie: Joanna (Asia) Krusiec. E-mail: lesnebyciecelestynow@gmail.com."),
        ("Polityka prywatności", "Jako administratora i „właściciela” strony wskazuje Panią Monikę Wolnik, ul. Galicyjska 16B, Bibice."),
        ("Oficjalna witryna", "https://www.lesnebycie.pl/"),
        ("Materiały źródłowe", "Strona główna, Oferta, O mnie (/o-nas-2), Wydarzenia, Blog i trzy wpisy, Kontakt, Newsletter, Polityka prywatności. Adres /Regulamin zwracał 404."),
        ("Data zebrania", "7 listopada 2026 r."),
        ("Cel zebrania", "Wsad źródłowy do audytu. Ocena jest w osobnym raporcie."),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = False
    for i, (k, v) in enumerate(rows):
        a, b = table.rows[i].cells
        a.width, b.width = Cm(4.6), Cm(12.8)
        pa, pb = a.paragraphs[0], b.paragraphs[0]
        pa.paragraph_format.space_after = Pt(1)
        pb.paragraph_format.space_after = Pt(1)
        _set_run(pa.add_run(k), 9, True)
        _set_run(pb.add_run(v), 9, False)

    h("1. Strona główna")
    p("„Zadbaj o swój wewnętrzny spokój”.")
    p("„Kąpiel leśna to powolne, uważne bycie w lesie.”")
    p("„To prosta a zarazem głęboka praktyka, która pozwala Twojemu układowi nerwowemu odpocząć, a Tobie odnaleźć wewnętrzny spokój.”")
    p("„Dla osób szukających wsparcia w redukcji stresu i poprawie nastroju poprzez kontakt z przyrodą.”")
    p("„Dla tych, którzy chcą zadbać o swój wewnętrzny spokój w naturalny sposób.”")
    p("FAQ, pytanie „Czy kąpiel leśna naprawdę działa?”:")
    p("„Tak, i potwierdzają to liczne badania naukowe. Już 120 minut w lesie znacząco obniża poziom kortyzolu (hormonu stresu), stabilizuje ciśnienie krwi i wzmacnia układ odpornościowy. Jeśli chcesz dowiedzieć się więcej, sprawdź nasze [BADANIA NAUKOWE].”")
    p("„Jako przewodnik pomagam Ci zaangażować zmysły poprzez proste ćwiczenia uważności – tu nie liczą się kilometry, ale Twoja obecność.”")
    p("„Pokonujemy zazwyczaj od 1 do 2 kilometrów w bardzo wolnym tempie.”")

    h("2. Oferta")
    p("„INWESTYCJA W TWÓJ DOBROSTAN NA TRZY SPOSOBY”.")
    p("„LEŚNY WARSZTAT DLA ZESPOŁÓW. Dla firm i organizacji, które szukają prawdziwej regeneracji i głębokiego odpoczynku dla swoich pracowników. WYCENA INDYWIDUALNA”.")
    p("„PRYWATNA KĄPIEL LEŚNA (DO 4 OSÓB) […] 500 zł/grupa”.")
    p("„KĄPIEL LEŚNA W GRUPIE (5-12 OSÓB) […] 150 zł /osoba”.")
    p("„INDYWIDUALNY PROGRAM WSPIERANIA DOBROSTANU. Indywidualnie opracowany 4-6 tygodniowy program, zawierający indywidualne spotkania kąpieli leśnych raz w tygodniu oraz aktywności do pracy własnej. Pozwala na stałe wprowadzić mechanizmy regulacji układu nerwowego do Twojego życia. 1500-2100 zł”.")
    p("„INDYWIDUALNA KĄPIEL LEŚNA. Jednorazowe spotkanie (ok. 2,5 - 3h) […] 300 zł”.")
    p("„Głębsza praca z ciałem: Wykorzystując moje doświadczenie z metodą Lowena i Labana, pomogę Ci poczuć, gdzie w Twoim ciele zapisał się stres i jak pozwolić mu odejść.”")
    p("„Warsztaty to autorskie spotkania tematyczne, podczas których łączę praktykę kąpieli leśnych z elementami psychologii, pracy z ciałem i ekologii głębokiej.”")

    h("3. O mnie")
    p("„Cześć! Nazywam się Asia Krusiec. Jestem ekspertką dobrostanu, przewodniczką po świecie kąpieli leśnych i edukatorką praktykującą interwencje oparte na naturze (Nature-Based Interventions).”")
    p("„Jako certyfikowana przewodniczka kąpieli leśnych i terapii lasem oraz pasjonatka neurobiologii dobrostanu, specjalizuję się w projektowaniu doświadczeń, które wykorzystują prozdrowotne właściwości lasu do redukcji stresu, poprawy odporności i regeneracji układu nerwowego.”")
    p("„Obecnie jestem trenerką w Forest Therapy Hub, gdzie szkolę przyszłych przewodników kąpieli leśnych.”")
    p("„Właścicielka marki i fundacji Leśne Bycie, gdzie od lat promuję ekologię, szacunek i świadomy kontakt z naturą.”")
    p("„Członkini Polskiego Towarzystwa Kąpieli Leśnych i Terapii Leśnej, dbając o standardy tej praktyki w Polsce.”")
    p("„Twórczyni firmy Projects and Trainings Joanna Krusiec […] Obecnie pogłębiam swoją wiedzę na studiach magisterskich z neuropsychologii.”")

    h("4. Blog, 20 lutego 2026 r.")
    p("Tytuł: „Dlaczego Las? O połączeniu neuropsychologii, biznesu i natury.”")
    p("„Kontakt z fitoncydami realnie obniża poziom kortyzolu. To szybki reset dla układu nerwowego, który w biurze rzadko wychodzi ze stanu czuwania.”")
    p("„Wierzę, że w nowoczesnym biznesie kontakt z naturą to nie luksus, ale niezbędny element higieny pracy.”")

    h("5. Blog, 26 lutego 2026 r.")
    p("Tytuł: „Dlaczego Twój mózg potrzebuje lasu bardziej niż nowej listy zadań?”")
    p("„Badania opublikowane w Environmental Health and Preventive Medicine dostarczyły twardych, mierzalnych dowodów […]. W trakcie szeroko zakrojonych analiz wykazano, że osoby praktykujące kąpiele leśne (shinrin-yoku) osiągały o 12,4% niższy poziom kortyzolu […] w porównaniu do osób przebywających w tym samym czasie w środowisku miejskim.”")
    p("Źródło podane przy tym akapicie: „Park, B. J., et al. (2010). The physiological effects of Shinrin-yoku. Environmental Health and Preventive Medicine.”")
    p("„Naukowcy zaobserwowali, że regularne obniżanie poziomu kortyzolu poprzez kontakt z lasem ma bezpośredni wpływ na wzmocnienie odporności organizmu oraz poprawę jakości snu […]. Zatem leśna terapia […] opiera się na udowodnionym naukowo procesie fizycznego wyciszania stanów zapalnych wywołanych przez przewlekłe napięcie psychiczne.”")
    p("„Badania Kaplana udowodniły, że kontakt z naturą przywraca klarowność myślenia i odporność psychiczną, co czyni kąpiele leśne jednym z najskuteczniejszych narzędzi zarządzania energią intelektualną w nowoczesnym biznesie.”")
    p("Drugie źródło: „Kaplan, S. (1995). The restorative effects of nature. Journal of Environmental Psychology.”")

    h("6. Blog, 12 lutego 2026 r.")
    p("Tytuł: „Kąpiele leśne pod lupą nauki. Zapraszamy na Ogólnopolską Konferencję Naukową w Przyłękowie.” Organizator podany na stronie: Polskie Towarzystwo Kąpieli Leśnych i Terapii Leśnej. Daty: 13–15 marca 2026 r.")
    p("„Dzięki badaniom naukowym – m.in. gościa specjalnego konferencji, prof. Qing Li – wiemy już, że czas spędzony wśród drzew to precyzyjne narzędzie regeneracji naszego organizmu.”")
    p("„Poprawa nastroju i witalność: Poznasz mechanizmy, dzięki którym natura skutecznie wspiera profilaktykę zdrowia psychicznego, redukując objawy lęku i depresji.”")
    p("„Dla świata biznesu (Managerowie i HR): Pokażemy, jak Terapia Lasem wspiera pracowników korporacji w walce z przebodźcowaniem, regeneruje uwagę i buduje odporność psychiczną zespołów (resilience), skutecznie zapobiegając wypaleniu zawodowemu.”")

    h("7. Kontakt i polityka prywatności")
    p("Kontakt: „ASIA KRUSIEC”, lesnebyciecelestynow@gmail.com.")
    p("Polityka prywatności: „Administratorem danych osobowych jest Pani Monika Wolnik, zamieszkała przy: Ul. Galicyjska 16B Bibice (zwany dalej: „Właściciel”).”")
    p("Ta sama polityka opisuje m.in. zakupy, płeć, adres dostawy, numer PESEL i NIP oraz stwierdza: „Niniejsza Polityka Prywatności jest zgodna z przepisami wynikającymi z art. 13 ust. 1 i ust. 2 rozporządzenia RODO.”")
    p("Strona Wydarzenia w dniu zebrania zawierała hasło „SPRAWDŹ AKTUALNE WYDARZENIA” bez listy wydarzeń.")

    out = ROOT / "Material_Lesne_Bycie_tresci_do_audytu.docx"
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
        _set_run(
            fp.add_run("ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Leśne Bycie  ·  str. "),
            8, False, False, MUTED,
        )
        page_field(fp, MUTED)

    add_p("ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO", size=9, bold=True, color=ACCENT, after=2)
    add_p("Audyt Ryzyk Greenwashingu: LEŚNE BYCIE", size=20, bold=True, color=PRIMARY, before=2, after=2)
    add_p(
        "Weryfikacja witryny https://www.lesnebycie.pl/ (strona główna, oferta, „O mnie”, wydarzenia, blog, kontakt, newsletter i polityka prywatności), "
        "stan na 7 listopada 2026 r., wobec ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym w tekście jednolitym Dz.U. z 2023 r. poz. 845 "
        "oraz ustawy o ochronie konkurencji i konsumentów w tekście jednolitym Dz.U. z 2025 r. poz. 1714. "
        "Ustawa z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) wchodzi w życie 27 września 2027 r. "
        "Do tej daty czarna lista EmpCo jest przyszłym standardem, a werdykt opiera się na przepisach już obowiązujących. "
        "Social washing oceniono tym samym testem prawdy, zakresu i pominięcia.",
        size=10.0, italic=True, color=MUTED, after=8,
    )

    if SHOT.exists():
        shot_par = doc.add_paragraph()
        shot_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shot_par.paragraph_format.space_before = Pt(2)
        shot_par.paragraph_format.space_after = Pt(2)
        shot_par.add_run().add_picture(str(SHOT), width=Cm(13.4))
        add_p(
            "Zrzut witryny https://www.lesnebycie.pl/ z 7 listopada 2026 r. Hasło wiodące: „Zadbaj o swój wewnętrzny spokój”.",
            size=8.5, italic=True, color=MUTED, after=8,
        )

    meta = [
        (
            "Badany podmiot",
            "Marka „Leśne Bycie”, witryna lesnebycie.pl. W tekstach ofertowych i na stronie „O mnie” występuje Joanna (Asia) Krusiec. "
            "Adres e-mail: lesnebyciecelestynow@gmail.com. Polityka prywatności wskazuje jako administratora i „właściciela” strony Monikę Wolnik, ul. Galicyjska 16B, Bibice. "
            "Na stronach oferty brak NIP i adresu osoby, z którą konsument zawiera umowę.",
        ),
        (
            "Przedmiot audytu",
            "Komunikacja handlowa na lesnebycie.pl: kąpiele leśne indywidualne i grupowe, warsztat dla firm, program 4–6 tygodni, trzy wpisy blogowe oraz polityka prywatności. "
            "Social washing: twierdzenia o fundacji, ekologii, szacunku i wpływie na pracowników.",
        ),
        (
            "Sektor działalności",
            "Usługi dobrostanu i edukacji w przyrodzie, sprzedawane konsumentom oraz firmom (warsztat dla zespołów, wycena indywidualna). "
            "To usługa spędzania czasu w lesie z przewodniczką, z cenami od 150 zł do 1500–2100 zł.",
        ),
        (
            "Reżim prawny",
            "Prawo już stosowane: art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (tekst jednolity Dz.U. 2023 poz. 845, status w ELI: obowiązujący) "
            "oraz art. 24 i art. 106 ust. 1 pkt 4 ustawy o ochronie konkurencji i konsumentów (tekst jednolity Dz.U. 2025 poz. 1714). "
            "Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) ma w art. 3 datę wejścia w życie 27 września 2027 r. "
            "W dniu audytu czarna lista EmpCo nie jest prawem polskim i nie jest podstawą werdyktu.",
        ),
        ("Data audytu", "7 listopada 2026 r."),
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
        pv.paragraph_format.line_spacing = 1.08
        _set_run(pk.add_run(k), 8.5, True, False, DARK)
        _set_run(pv.add_run(v), 8.5, i == 5, False, DANGER if i == 5 else DARK)
    add_p("", size=4, after=6)

    callout(
        "WERDYKT: NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)",
        [
            "Na stronie jest co najmniej jedno naruszenie prawa już obowiązującego o wadze bardzo wysokiej oraz trzy ustalenia o wadze wysokiej. Podstawą jest art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. Czarna lista z ustawy Dz.U. 2026 poz. 1353 wchodzi w życie dopiero 27 września 2027 r. i nie jest tu stwierdzonym naruszeniem.",
            "FAQ sprzedaje dawkę „120 minut” jako potwierdzony, łączny skutek dla kortyzolu, ciśnienia krwi i odporności. Blog i zapowiedź konferencji przenoszą przywołane prace na odporność, stany zapalne, objawy lęku i depresji oraz zapobieganie wypaleniu. Program za 1500–2100 zł obiecuje trwałe wprowadzenie mechanizmów regulacji układu nerwowego. Strona „O mnie” przedstawia Joannę Krusiec jako właścicielkę fundacji Leśne Bycie, która od lat promuje ekologię, podczas gdy fundacja przy adresie z polityki prywatności nazywa się Fundacja Leśna Przestrzeń i została wpisana do KRS 30 grudnia 2025 r.",
            "Samo zbliżenie słów „ekologia” i „ekologia głęboka” do przyszłego zakazu ogólnych twierdzeń środowiskowych nie daje tego werdyktu. Werdykt czerwony wynika z twierdzeń, które już dziś mogą skłonić do zakupu usługi albo do zaufania jej społecznemu charakterowi.",
        ],
        "FFF1F2", "E11D48", DANGER,
    )

    add_h("1. Analiza Wstępna")
    add_p(
        "Leśne Bycie oferuje prowadzone kąpiele leśne: spotkanie indywidualne około 2,5–3 godzin za 300 zł, grupę 5–12 osób za 150 zł od osoby, "
        "grupę do 4 osób za 500 zł, warsztat dla firm wyceniany indywidualnie oraz program 4–6 tygodni za 1500–2100 zł. "
        "Autorką przekazu jest Joanna Krusiec, przedstawiana też jako Asia Krusiec, przewodniczka kąpieli leśnych i terapii lasem. "
        "Narracja łączy uważność w lesie, język neurobiologii i ofertę dla osób przemęczonych oraz dla zespołów. "
        "Na stronie „O mnie” dochodzi warstwa społeczna: własna fundacja, wieloletnie promowanie ekologii i szacunku oraz świadomy kontakt z naturą."
    )
    add_p(
        "Przesłanki art. 5 i art. 6 spełniają cztery twierdzenia. Pierwsze odpowiada w FAQ na pytanie, czy kąpiel „naprawdę działa”, i podaje jedną dawkę oraz trzy skutki fizjologiczne jako wynik licznych badań. "
        "Drugie, na blogu i w zapowiedzi konferencji, dokłada do przywołanej pracy mechanizmy, których ta praca nie opisuje, oraz skutki kliniczne i pracownicze. "
        "Trzecie sprzedaje najdroższy program jako sposób na trwałą regulację układu nerwowego. "
        "Czwarte dotyczy tożsamości: właścicielka fundacji o nazwie, której nie ma w odpisie fundacji siedzącej pod adresem z polityki prywatności, przy równoczesnym wskazaniu innej osoby jako właściciela strony. "
        "To jest social washing w ścisłym sensie tego audytu: odbiorca może wziąć płatną usługę za działalność fundacji o dłuższym i szerszym mandacie społecznym, niż pokazuje rejestr."
    )
    add_p(
        "Poza tymi czterema zdaniami materiał nie dał osobnego ustalenia. Strona nie używa znaku zrównoważonego charakteru, nie twierdzi, że usługa jest neutralna klimatycznie, i nie porównuje się z nazwanym konkurentem. "
        "Odróżnienie kąpieli od spaceru oraz przywołanie lasu i miasta w opisach badań nie jest reklamą porównawczą, dlatego raport nie ma sekcji porównawczej. "
        "Członkostwo w Polskim Towarzystwie Kąpieli Leśnych i Terapii Leśnej oraz certyfikat Forest Therapy Hub pozostają bez ustalenia: audyt nie uzyskał podstawy, by uznać te zdania za nieprawdziwe. "
        "Studia magisterskie z neuropsychologii są na tej samej stronie opisane jako trwające. Słowa „ekologia” i „ekologia głęboka” nie są dziś osobnym naruszeniem art. 5 jako twierdzenie o efektywności środowiskowej usługi. Od 27 września 2027 r. wracają w ryzykach jako przyszły standard, jeżeli zaczną być czytane jako korzyść środowiskowa całej działalności."
    )

    add_h("2. Analiza szczegółowa - tabela zbiorcza naruszeń")
    findings = [
        (
            "01",
            "120 minut w lesie jako łączny, potwierdzony skutek dla kortyzolu, ciśnienia i odporności",
            "UPNPR art. 5 ust. 1, ust. 2 pkt 1 i 2, ust. 3 pkt 2 oraz art. 6 ust. 1 i ust. 3 pkt 1",
            "BARDZO WYSOKIE",
            "Jedna dawka i trzy skutki fizjologiczne w odpowiedzi na pytanie, czy usługa naprawdę działa",
        ),
        (
            "02",
            "Przeniesienie przywołanych badań na odporność, stany zapalne, lęk, depresję i wypalenie",
            "UPNPR art. 5 ust. 1, ust. 2 pkt 1 i 2, ust. 3 pkt 2 oraz art. 6 ust. 1",
            "WYSOKIE",
            "Komentarz do prac wykracza poza ich zakres i sprzedaje szerszy skutek zdrowotny oraz pracowniczy",
        ),
        (
            "03",
            "Trwałe wprowadzenie mechanizmów regulacji układu nerwowego w programie za 1500–2100 zł",
            "UPNPR art. 5 ust. 1, ust. 2 pkt 2 i ust. 3 pkt 2",
            "WYSOKIE",
            "Najdroższy produkt obiecuje trwały skutek fizjologiczny, którego przywołane badania nie opisują",
        ),
        (
            "04",
            "Właścicielka fundacji Leśne Bycie, która od lat promuje ekologię",
            "UPNPR art. 5 ust. 1, ust. 2 pkt 1, ust. 3 pkt 6 oraz art. 6 ust. 1, ust. 3 pkt 1 i ust. 4 pkt 2",
            "WYSOKIE",
            "Nazwa, rola i czas trwania nie pokrywają się z KRS; polityka prywatności wskazuje inną osobę",
        ),
    ]
    sum_table = doc.add_table(rows=1 + len(findings), cols=5)
    sum_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_table.autofit = False
    sum_table.style = "Table Grid"
    col_w = (Cm(1.15), Cm(5.15), Cm(4.15), Cm(2.55), Cm(4.6))
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota naruszenia"]
    for c, htxt in enumerate(headers):
        cell = sum_table.rows[0].cells[c]
        cell.width = col_w[c]
        shade(cell, "08181F")
        margins(cell, 40, 40, 40, 40)
        par = cell.paragraphs[0]
        par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
        _set_run(par.add_run(htxt), 8, True, False, WHITE)
    risk_fill = {
        "BARDZO WYSOKIE": ("FFF1F2", DANGER),
        "WYSOKIE": ("FFF7ED", WARNING),
    }
    for r, row in enumerate(findings):
        bg = "FFFFFF" if r % 2 == 0 else "F8FAFC"
        for c, val in enumerate(row):
            cell = sum_table.rows[r + 1].cells[c]
            cell.width = col_w[c]
            shade(cell, risk_fill[row[3]][0] if c == 3 else bg)
            margins(cell, 40, 40, 40, 40)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            par = cell.paragraphs[0]
            par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
            par.paragraph_format.line_spacing = 1.05
            color = risk_fill[row[3]][1] if c == 3 else DARK
            _set_run(par.add_run(val), 7.5, c == 3, False, color)
    add_p("", size=4, after=4)

    add_h("3. Szczegółowa analiza przypadków")

    add_p("PRZYPADEK 01: Dawka 120 minut jako potwierdzony skutek dla kortyzolu, ciśnienia i odporności", size=12, bold=True, color=NAVY, before=8, after=2)
    add_p(
        "Podstawa prawna: art. 5 ust. 1, ust. 2 pkt 1 i 2 oraz ust. 3 pkt 2, a także art. 6 ust. 1 i ust. 3 pkt 1 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. | Poziom ryzyka: BARDZO WYSOKIE",
        size=9, italic=True, color=MUTED, after=4,
    )
    add_p("Cytat źródłowy:", size=10, bold=True, after=2)
    add_p(
        "„Tak, i potwierdzają to liczne badania naukowe. Już 120 minut w lesie znacząco obniża poziom kortyzolu (hormonu stresu), stabilizuje ciśnienie krwi i wzmacnia układ odpornościowy. Jeśli chcesz dowiedzieć się więcej, sprawdź nasze [BADANIA NAUKOWE].” "
        "Pytanie nad tym zdaniem: „Czy kąpiel leśna naprawdę działa?” Strona główna, FAQ. Odesłanie w nawiasie kwadratowym nie jest linkiem do strony z badaniami. W menu witryny takiej strony nie ma.",
        size=10, italic=True,
    )
    add_p("Ocena prawna:", size=10, bold=True, before=4, after=2)
    add_p(
        "Zdanie stoi przy płatnej usłudze i odpowiada na pytanie o jej skuteczność. Art. 5 ust. 1 uznaje za wprowadzające w błąd działanie, które powoduje albo może powodować podjęcie przez przeciętnego konsumenta decyzji dotyczącej umowy, której inaczej by nie podjął. "
        "Ust. 3 pkt 2 obejmuje korzyści związane z produktem oraz testy i wyniki badań. Ust. 2 pkt 1 obejmuje informację nieprawdziwą, a pkt 2 informację prawdziwą rozpowszechnioną w sposób mogący wprowadzać w błąd. "
        "Art. 6 ust. 1 i ust. 3 pkt 1 obejmują pominięcie istotnych informacji o produkcie, gdy przez to konsument może podjąć inną decyzję."
    )
    add_p(
        "Liczba 120 minut występuje w badaniu White’a i współautorów z 2019 r. w Scientific Reports: co najmniej 120 minut kontaktu z przyrodą w ciągu tygodnia wiązało się z wyższą szansą samoopisu dobrego zdrowia i wysokiego dobrostanu. "
        "To badanie obserwacyjne dotyczy przyrody w ogóle i tygodniowej sumy czasu. Nie mierzy kortyzolu, ciśnienia ani odporności. "
        "Praca Parka, Tsunetsugu, Kasetaniego, Kagawy i Miyazakiego z 2010 r. w Environmental Health and Preventive Medicine, odczytana w PMC, opisuje 24 lasy w Japonii i 280 osób w wieku średnio 21,7 roku. "
        "Kortyzol w ślinie był niższy w lesie o 13,4% po oglądaniu trwającym średnio 14 minut i o 15,8% po marszu trwającym średnio 16 minut, w porównaniu z miastem. Skurczowe ciśnienie było niższe o około 1,7–1,9%, rozkurczowe o około 1,6–2,1%. Ta praca nie mierzy odporności."
    )
    add_p(
        "Wzrost aktywności komórek NK opisują prace Li dotyczące wyjazdów leśnych. W badaniu dziennym (PubMed 20487629) uczestnicy szli po dwie godziny rano i po dwie godziny po południu. Następnego ranka kortyzol był niższy, a aktywność NK wyższa. "
        "Trzydniowe wyjazdy również wiązały się ze wzrostem aktywności NK. Żaden z tych protokołów nie jest dawką „już 120 minut” obejmującą łącznie kortyzol, stabilizację ciśnienia i wzmocnienie układu odpornościowego. "
        "Usługa na tej samej stronie to wolne przejście 1–2 km. FAQ składa z różnych prac jedną dawkę i trzy skutki oraz odsyła do badań, których na stronie nie ma."
    )
    add_p(
        "Test decyzji dotyczącej transakcji jest spełniony. Konsument, który waha się, czy zapłacić 150, 300 albo 500 zł, dostaje odpowiedź, że usługa naprawdę działa, bo 120 minut znacząco zmienia trzy wskaźniki. "
        "Decyzja o zapisie jest decyzją dotyczącą umowy. Tak wyrok TSUE z 19 grudnia 2013 r., C-281/12, Trento Sviluppo i Centrale Adriatica. "
        "Sentencję odczytano w EUR-Lex: praktyka jest wprowadzająca w błąd, gdy zawiera nieprawdziwe informacje albo może wprowadzić w błąd przeciętnego konsumenta oraz może skłonić go do decyzji dotyczącej transakcji, której inaczej by nie podjął. "
        "Pojęcie tej decyzji obejmuje decyzję bezpośrednio związaną z tym, czy konsument dokona zakupu. To omówienie sentencji. Wyrok dotyczy reklamy dostępności towaru w supermarkecie i nie rozstrzyga sytuacji Leśnego Bycia. "
        "Wspólne jest to, że sama atrakcyjność przekazu nie wystarcza: potrzebny jest możliwy wpływ na decyzję o zakupie. Tutaj ten wpływ wynika z odpowiedzi na pytanie o skuteczność, postawionej obok cennika. "
        "Polskiego wyroku o zakresie badań nie powołuję. Baza SN i SAOS nie zwróciły w tej sesji tekstu orzeczenia, które miałoby nieść tę kwalifikację."
    )
    add_p(
        "Wytyczne OECD dla przedsiębiorstw wielonarodowych (edycja 2023) dostają przy tym zdaniu etykietę NISKIE. To zalecenia dotyczące rzetelnej informacji. "
        "Zawiadomienie może trafić do polskiego Krajowego Punktu Kontaktowego. Postępowanie przed KPK nie jest sprawą sądową i nie kończy się wyrokiem. Skutkiem jest publiczna ocena, zalecenia i ryzyko wizerunkowe. Ten poziom nie podnosi wagi ustalenia."
    )
    add_p("Kierunek naprawy:", size=10, bold=True, before=4, after=2)
    add_p(
        "Zdanie w FAQ należy zdjąć z publikacji albo zastąpić opisem praktyki bez dawki i bez trzech skutków naraz. "
        "Jeżeli zostaje jakiekolwiek odniesienie do badań, w tym samym miejscu mają być: autor, rok, czas trwania ekspozycji, grupa, co zmierzono i że wynik jest średnią grupową wobec miasta, a nie pomiarem uczestnika tej sesji. "
        "Odesłanie „BADANIA NAUKOWE” albo prowadzi do takiego opisu, albo znika."
    )
    callout(
        "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
        [
            "„Kąpiel leśna, którą prowadzę, jest wolnym, uważnym pobytem w lesie, zwykle na odcinku 1–2 km. W japońskich badaniach terenowych krótki pobyt w lesie, rzędu kilkunastu minut oglądania albo marszu, wiązał się z niższym średnim stężeniem kortyzolu w ślinie niż pobyt w mieście, u młodych dorosłych uczestników. To wynik grupowy z tamtego protokołu. Nie jest pomiarem Twojego kortyzolu, ciśnienia ani odporności podczas tej sesji i nie obiecuje takiego skutku.”",
        ],
        "F0FDFA", "0F766E", RGBColor(15, 118, 110),
    )

    add_p("PRZYPADEK 02: Rozszerzenie przywołanych badań na odporność, stany zapalne, lęk, depresję i wypalenie", size=12, bold=True, color=NAVY, before=4, after=2)
    add_p(
        "Podstawa prawna: art. 5 ust. 1, ust. 2 pkt 1 i 2 oraz ust. 3 pkt 2, a także art. 6 ust. 1 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. | Poziom ryzyka: WYSOKIE",
        size=9, italic=True, color=MUTED, after=4,
    )
    add_p("Cytat źródłowy:", size=10, bold=True, after=2)
    add_p(
        "Blog z 20 lutego 2026 r.: „Kontakt z fitoncydami realnie obniża poziom kortyzolu. To szybki reset dla układu nerwowego, który w biurze rzadko wychodzi ze stanu czuwania.”",
        size=10, italic=True,
    )
    add_p(
        "Blog z 26 lutego 2026 r., ze źródłem „Park, B. J., et al. (2010)”: „osoby praktykujące kąpiele leśne (shinrin-yoku) osiągały o 12,4% niższy poziom kortyzolu […] w porównaniu do osób przebywających w tym samym czasie w środowisku miejskim.” "
        "Dalej, przy tym samym numerze źródła: „regularne obniżanie poziomu kortyzolu poprzez kontakt z lasem ma bezpośredni wpływ na wzmocnienie odporności organizmu oraz poprawę jakości snu […]. Zatem leśna terapia […] opiera się na udowodnionym naukowo procesie fizycznego wyciszania stanów zapalnych.” "
        "Oraz: „Badania Kaplana udowodniły, że kontakt z naturą przywraca klarowność myślenia i odporność psychiczną, co czyni kąpiele leśne jednym z najskuteczniejszych narzędzi zarządzania energią intelektualną w nowoczesnym biznesie.”",
        size=10, italic=True,
    )
    add_p(
        "Blog z 12 lutego 2026 r.: „natura skutecznie wspiera profilaktykę zdrowia psychicznego, redukując objawy lęku i depresji.” "
        "Oraz: „Terapia Lasem […] skutecznie zapobiegając wypaleniu zawodowemu.” Organizatorem konferencji strona nazywa Polskie Towarzystwo Kąpieli Leśnych i Terapii Leśnej. Tekst stoi na witrynie sprzedażowej Leśnego Bycia.",
        size=10, italic=True,
    )
    add_p("Ocena prawna:", size=10, bold=True, before=4, after=2)
    add_p(
        "To osobne twierdzenie od FAQ, bo podaje inny mechanizm i inne skutki, już z przypisem. Przypis prowadzi do pracy Parka i współautorów z 2010 r. W tekście tej pracy odczytanym w PMC kortyzol po oglądaniu był niższy o 13,4%, a po marszu o 15,8%. "
        "Liczba 12,4% pojawia się w innym omówieniu tego nurtu: 420 osób, 35 lasów, spadek kortyzolu przy siedzeniu w przyrodzie wobec miasta. Sama różnica 12,4% wobec 13,4% albo 15,8% byłaby drobną nieprecyzyjnością i nie dostałaby osobnego przypadku. "
        "Na stronie liczba jest początkiem dalszego zdania: ten sam przypis ma dowieść bezpośredniego wpływu na odporność, sen i „udowodnione” wyciszanie stanów zapalnych. Praca Parka z 2010 r. mierzy kortyzol, tętno, ciśnienie i zmienność rytmu serca. Nie raportuje markerów zapalnych, snu ani odporności."
    )
    add_p(
        "Fitoncydy w pracach Li są wykrywane w powietrzu leśnym. Abstrakty tych prac wiążą je, obok niższego hormonu stresu, z możliwym częściowym udziałem we wzroście aktywności komórek NK po wyjazdach. "
        "Zdanie „kontakt z fitoncydami realnie obniża poziom kortyzolu” przypisuje fitoncydom skutek, którego przywołana na blogu praca Parka nie izoluje. "
        "Kaplan 1995, w brzmieniu źródła podanego na stronie, jest artykułem o teorii przywracania uwagi. Strona nazywa go badaniem, które udowodniło, że kąpiele leśne są jednym z najskuteczniejszych narzędzi zarządzania energią intelektualną w biznesie. "
        "Artykuł teoretyczny z 1995 r. nie jest rankingiem narzędzi dla firm."
    )
    add_p(
        "Zapowiedź konferencji dokłada skutek kliniczny i pracowniczy: redukcja objawów lęku i depresji oraz skuteczne zapobieganie wypaleniu. "
        "Sformułowanie „poznasz mechanizmy” i „pokażemy, jak” opisuje program wydarzenia, a jednocześnie stwierdza, że natura i terapia lasem ten skutek osiągają. "
        "Na tej samej witrynie warsztat dla firm jest sprzedawany jako „prawdziwa regeneracja” pracowników. Czytane łącznie, zdania o lęku, depresji i wypaleniu mogą zdecydować o zakupie warsztatu albo programu. "
        "Test z art. 5 ust. 1 jest spełniony: oczekiwany wynik zastosowania usługi, o którym mowa w ust. 3 pkt 2, jest szerszy niż dowód podany obok. "
        "Wyrok C-281/12 stosuje się tu tak samo jak w przypadku 01: możliwy wpływ na decyzję o zakupie, bez wymogu wykazania, że ktoś już kupił. Wyrok nie dotyczy tego przedsiębiorcy. "
        "Polskiego orzeczenia o zakresie wykorzystania wyników badań nie powołuję, bo jego tekstu nie otworzyłem w bazie SN ani w SAOS."
    )
    add_p(
        "Ścieżka KPK i Wytycznych OECD 2023 pozostaje przy tym przypadku na poziomie NISKIE. Dotyczy rzetelności informacji o usłudze i o wpływie na ludzi w pracy. Nie jest wyrokiem i nie zmienia poziomu WYSOKIE, który wynika z art. 5 i 6."
    )
    add_p("Kierunek naprawy:", size=10, bold=True, before=4, after=2)
    add_p(
        "Przy Parku 2010 zostawić kortyzol, tętno, ciśnienie i zmienność rytmu, z czasem ekspozycji i z informacją, że to średnia młodych dorosłych w Japonii wobec miasta. "
        "Odporność opisać osobno, jako aktywność komórek NK po wskazanym wyjeździe Li, z długością tego wyjazdu. "
        "Zdanie o stanach zapalnych, o fitoncydach jako przyczynie spadku kortyzolu, o „udowodnieniu” przez Kaplana pierwszeństwa kąpieli w biznesie oraz o skutecznej redukcji objawów lęku, depresji i wypalenia zdjąć albo przepisać jako temat dyskusji na konferencji, bez twierdzenia, że usługa ten skutek wywołuje."
    )
    callout(
        "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
        [
            "„W badaniu Parka i współautorów z 2010 r. (Environmental Health and Preventive Medicine) średni kortyzol w ślinie był u młodych dorosłych niższy w lesie niż w mieście: o 13,4% po oglądaniu trwającym średnio 14 minut i o 15,8% po marszu trwającym średnio 16 minut. Ta praca nie mierzy odporności ani stanów zapalnych. Osobne badania wyjazdów leśnych, w tym prace Qing Li, obserwowały zmianę aktywności komórek NK przy dłuższych protokołach, na przykład przy dwóch dwugodzinnych spacerach jednego dnia albo przy wyjeździe trzydniowym. Nie przenoszę tych wyników na sesję, którą prowadzę, i nie twierdzę, że zapobiega ona wypaleniu albo zmniejsza objawy depresji.”",
        ],
        "F0FDFA", "0F766E", RGBColor(15, 118, 110),
    )

    add_p("PRZYPADEK 03: Trwałe wprowadzenie mechanizmów regulacji układu nerwowego", size=12, bold=True, color=NAVY, before=4, after=2)
    add_p(
        "Podstawa prawna: art. 5 ust. 1, ust. 2 pkt 2 i ust. 3 pkt 2 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. | Poziom ryzyka: WYSOKIE",
        size=9, italic=True, color=MUTED, after=4,
    )
    add_p("Cytat źródłowy:", size=10, bold=True, after=2)
    add_p(
        "Oferta, „INDYWIDUALNY PROGRAM WSPIERANIA DOBROSTANU”: „Indywidualnie opracowany 4-6 tygodniowy program, zawierający indywidualne spotkania kąpieli leśnych raz w tygodniu oraz aktywności do pracy własnej. Pozwala na stałe wprowadzić mechanizmy regulacji układu nerwowego do Twojego życia. 1500-2100 zł.”",
        size=10, italic=True,
    )
    add_p("Ocena prawna:", size=10, bold=True, before=4, after=2)
    add_p(
        "To twierdzenie o spodziewanym wyniku najdroższego produktu. „Na stałe” i „mechanizmy regulacji układu nerwowego” nadają programowi skutek fizjologiczny i trwały. "
        "Prace przywołane na tej samej witrynie opisują zmiany średnich wskaźników w trakcie albo krótko po ekspozycji. W omówieniu wyjazdów leśnych wzrost aktywności NK bywał obserwowany jeszcze przez dni, a w części protokołów do około miesiąca po wyjeździe. "
        "Żaden cytowany na stronie opis nie stwierdza, że cztery do sześciu cotygodniowych spotkań oraz praca własna trwale instalują mechanizmy regulacji układu nerwowego. "
        "Słowo „pozwala” nie zdejmuje tego znaczenia, bo stoi przy cenie i przy opisie, czym program różni się od jednorazowej sesji za 300 zł."
    )
    add_p(
        "Konsument może wybrać program za 1500–2100 zł właśnie dlatego, że ma on dawać skutek trwały, a nie jednorazowy pobyt. "
        "To decyzja dotycząca umowy w teście art. 5 ust. 1, w rozumieniu potwierdzonym sentencją C-281/12. Wyrok nie dotyczy Leśnego Bycia. "
        "Waga jest wysoka, a nie bardzo wysoka: zdanie nie podaje fałszywej liczby. Da się je naprawić opisem tego, co program zawiera, bez obietnicy trwałej zmiany fizjologii."
    )
    add_p("Kierunek naprawy:", size=10, bold=True, before=4, after=2)
    add_p(
        "Zostawić liczbę spotkań, czas, cenę i informację, że między sesjami są zadania do pracy własnej. "
        "Skutek opisać jako praktykę uważności, którą uczestnik może powtarzać, albo usunąć. Nie przypisywać programowi trwałej regulacji układu nerwowego."
    )
    callout(
        "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
        [
            "„Program trwa 4–6 tygodni. W każdym tygodniu jest jedno indywidualne spotkanie w lesie oraz zadania do samodzielnej praktyki. Cena: 1500–2100 zł. Program uczy powtarzalnego sposobu uważnego pobytu w lesie. Nie jest pomiarem układu nerwowego i nie obiecuje trwałej zmiany jego regulacji.”",
        ],
        "F0FDFA", "0F766E", RGBColor(15, 118, 110),
    )

    add_p("PRZYPADEK 04: Właścicielka fundacji Leśne Bycie, która od lat promuje ekologię", size=12, bold=True, color=NAVY, before=4, after=2)
    add_p(
        "Podstawa prawna: art. 5 ust. 1, ust. 2 pkt 1 i ust. 3 pkt 6 oraz art. 6 ust. 1, ust. 3 pkt 1 i ust. 4 pkt 2 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. | Poziom ryzyka: WYSOKIE",
        size=9, italic=True, color=MUTED, after=4,
    )
    add_p("Cytat źródłowy:", size=10, bold=True, after=2)
    add_p(
        "Strona „O mnie”: „Właścicielka marki i fundacji Leśne Bycie, gdzie od lat promuję ekologię, szacunek i świadomy kontakt z naturą.”",
        size=10, italic=True,
    )
    add_p(
        "Polityka prywatności: „Administratorem danych osobowych jest Pani Monika Wolnik, zamieszkała przy: Ul. Galicyjska 16B Bibice (zwany dalej: „Właściciel”).” "
        "Kontakt do zapisów: „ASIA KRUSIEC”, lesnebyciecelestynow@gmail.com. Przy cenach i przycisku „ZAPISZ SIĘ” brak imienia przedsiębiorcy w połączeniu z adresem.",
        size=10, italic=True,
    )
    add_p("Ocena prawna:", size=10, bold=True, before=4, after=2)
    add_p(
        "Art. 5 ust. 3 pkt 6 obejmuje informacje o sytuacji prawnej przedsiębiorcy, jego imieniu i nazwisku, statusie, członkostwie i powiązaniach. "
        "Ust. 2 pkt 1 obejmuje informację nieprawdziwą. Art. 6 ust. 4 pkt 2 uznaje za istotne przy propozycji nabycia produktu imię i nazwisko albo nazwę oraz adres przedsiębiorcy i przedsiębiorcy, na którego rzecz działa. "
        "Oferta z ceną i z „ZAPISZ SIĘ” jest taką propozycją. Konsument nie dostaje przy niej adresu. Dostaje dwie różne osoby: Asię Krusiec jako twarz usługi i Monikę Wolnik jako właściciela strony."
    )
    add_p(
        "Odpis aktualny KRS, pobrany 7 listopada 2026 r. z API Ministerstwa Sprawiedliwości, stan na 24 czerwca 2026 r., numer KRS 0001213844, rejestr stowarzyszeń: fundacja o nazwie FUNDACJA LEŚNA PRZESTRZEŃ, NIP 5130308584, "
        "siedziba ul. Galicyjska 16B, 32-087 Bibice, adres elektroniczny monikawolnik1@gmail.com, strona www.lesnaprzestrzen.pl, data rejestracji 30 grudnia 2025 r. "
        "Jest oddział w Celestynowie, ul. Kolejowa 13. Prezesem zarządu jest osoba o imieniu i nazwisku zamaskowanych w API w sposób zgodny z Moniką Marią Wolnik. "
        "W składzie jest też członkini, której zamaskowane imię i nazwisko mają długość zgodną z Joanną Krusiec. Publiczne odwzorowanie odpisu wymienia w zarządzie Justynę Kozacką, Monikę Marię Wolnik i Joannę Krusiec. "
        "Cele widoczne w odczytanym fragmencie to promowanie świadomego kontaktu z naturą, wiedza o jego wpływie na zdrowie oraz rozwój kąpieli leśnych i florystyki jako formy dobrostanu. Fundacja nie nazywa się Leśne Bycie."
    )
    add_p(
        "Zdanie ze strony jest więc nieprawdziwe w trzech punktach, które konsument widzi naraz. Nazwa fundacji na stronie to Leśne Bycie, a w rejestrze przy tym adresie jest Leśna Przestrzeń. "
        "Rola to „właścicielka”, podczas gdy urząd prezesa i adres administratora danych schodzą się z Moniką Wolnik, a Joanna Krusiec – jeżeli potwierdzi to niemaskowany odpis – jest członkinią zarządu, nie właścicielką. "
        "„Od lat” nie opisuje fundacji wpisanej 30 grudnia 2025 r. W dniu audytu miała niespełna jedenaście miesięcy. "
        "Słowo „ekologię” jest szersze niż cele odczytane w odpisie, które dotyczą kontaktu z naturą, zdrowia i samej praktyki kąpieli. "
        "Konsument może wybrać płatny zapis, bo widzi przewodniczkę stojącą na czele wieloletniej fundacji ekologicznej. To decyzja dotycząca umowy. Test art. 5 ust. 1 jest spełniony. Wyrok C-281/12 potwierdza ten próg i nie dotyczy tego podmiotu."
    )
    add_p(
        "To ustalenie jest social washingiem: płatna usługa dostaje wiarygodność organizacji społecznej, szacunku i ekologii. "
        "Wytyczne OECD 2023, jako zalecenie rzetelnego ujawnienia, mają tu poziom NISKIE. Skarga do polskiego KPK nie jest sprawą sądową i nie kończy się wyrokiem. Zostają publiczna ocena, zalecenia i skutek wizerunkowy. Nie podnoszą one wagi ponad art. 5 i 6."
    )
    add_p(
        "Polityka prywatności dodatkowo opisuje sklep, płeć, adres dostawy, PESEL i NIP, podczas gdy formularze na stronie to kontakt i newsletter. "
        "Twierdzi też zgodność z art. 13 ust. 1 i 2 rozporządzenia 2016/679. Tego rozporządzenia nie otwierałem w tej sesji zdanie po zdaniu, więc nie stawiam osobnego ustalenia na jego artykule. "
        "Sprzeczność co do tego, kto jest przedsiębiorcą i właścicielem strony, wystarcza na art. 5 i 6."
    )
    add_p("Kierunek naprawy:", size=10, bold=True, before=4, after=2)
    add_p(
        "Przy cenach podać imię i nazwisko albo firmę, adres i NIP osoby, która świadczy usługę i wystawia rachunek. "
        "Fundację nazywać tak, jak w aktualnym odpisie, z numerem KRS, datą wpisu i funkcją, która w tym odpisie naprawdę jest. "
        "Nie pisać „właścicielka fundacji Leśne Bycie” ani „od lat” o fundacji z 30 grudnia 2025 r. "
        "Politykę prywatności napisać na administratora, który faktycznie dostaje formularz, i wykreślić z niej opis sklepu, jeżeli sklepu nie ma."
    )
    callout(
        "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
        [
            "„Usługę prowadzi Joanna Krusiec, [adres wykonywania działalności], NIP [numer]. Zapis i płatność dotyczą tej działalności. Joanna Krusiec jest członkinią zarządu Fundacji Leśna Przestrzeń, KRS 0001213844, wpisanej 30 grudnia 2025 r., z siedzibą przy ul. Galicyjskiej 16B, 32-087 Bibice – o ile aktualny, niemaskowany odpis nadal tak stanowi. Fundacja nie nazywa się Leśne Bycie. Leśne Bycie jest marką usługi. Prezeską zarządu fundacji jest Monika Maria Wolnik.” "
            "Drugie zdanie o funkcji wolno opublikować dopiero po sprawdzeniu odpisu z dnia korekty. Jeżeli funkcja jest inna, zdanie ma powtórzyć odpis, a nie ten wzór.",
        ],
        "F0FDFA", "0F766E", RGBColor(15, 118, 110),
    )

    add_h("5. Ryzyka prawne i finansowe")
    add_p("• 1. Postępowanie Prezesa UOKiK i kara do 10% obrotu", size=11, bold=True, color=NAVY, before=2, after=2)
    add_p(
        "Art. 24 ust. 1 i ust. 2 pkt 2 oraz 3 ustawy o ochronie konkurencji i konsumentów, w tekście jednolitym Dz.U. 2025 poz. 1714, zakazuje praktyk naruszających zbiorowe interesy konsumentów, w tym naruszenia obowiązku informacji rzetelnej, prawdziwej i pełnej oraz nieuczciwych praktyk rynkowych. "
        "Witryna z cennikiem i formularzem zapisu jest takim przekazem masowym. Art. 106 ust. 1 pkt 4 pozwala Prezesowi Urzędu nałożyć karę do 10% obrotu osiągniętego w roku obrotowym poprzedzającym rok nałożenia kary, także przy naruszeniu nieumyślnym. "
        "Ryzyko dotyczy przypadków 01–04 łącznie, bo wszystkie spełniają art. 5 albo 6 ustawy o nieuczciwych praktykach."
    )
    add_p("• 2. Decyzja zakupowa oparta na obietnicy fizjologicznej", size=11, bold=True, color=NAVY, before=6, after=2)
    add_p(
        "Ceny widoczne na stronie to 150 zł, 300 zł, 500 zł oraz 1500–2100 zł, a warsztat dla firm nie ma ceny z góry. "
        "Konsument albo osoba z działu HR może wybrać droższy wariant, bo FAQ, blog i opis programu przypisują mu skutek, którego protokół badań nie pokrywa. "
        "To jest ten sam mechanizm decyzji dotyczącej umowy, który niesie przypadki 01–03. Nie powołuję dodatkowego przepisu o rękojmi, bo jego tekstu nie otwierałem w tej sesji."
    )
    add_p("• 3. Standard EmpCo od 27 września 2027 r.", size=11, bold=True, color=NAVY, before=6, after=2)
    add_p(
        "Art. 3 ustawy z 4 września 2026 r. (Dz.U. 2026 poz. 1353), odczytany w tekście opublikowanym w ELI, stanowi, że ustawa wchodzi w życie 27 września 2027 r. "
        "Do poprzedniego dnia włącznie nie stwierdzam naruszenia krajowej czarnej listy. "
        "Słowa „promuję ekologię” oraz „ekologia głęboka” w opisie warsztatu nie są dziś osobnym naruszeniem art. 5 jako twierdzenie, że usługa ma wysoką efektywność środowiskową. "
        "Jeżeli przed 27 września 2027 r. zostaną czytane jako korzyść środowiskowa całej działalności albo całej usługi, będą wymagały zawężenia do konkretnej czynności. Sam ten przyszły standard nie niesie werdyktu czerwonego. Werdykt czerwony niosą przypadki 01–04."
    )
    add_p("• 4. Wytyczne OECD i polski KPK – poziom NISKIE", size=11, bold=True, color=NAVY, before=6, after=2)
    add_p(
        "Wytyczne OECD z 2023 r. są zaleceniem. Przy twierdzeniu o skutku zdrowotnym usługi i przy twierdzeniu o fundacji promującej ekologię zawiadomienie może trafić do Krajowego Punktu Kontaktowego w Ministerstwie Funduszy i Polityki Regionalnej. "
        "Postępowanie nie jest sprawą sądową i nie kończy się wyrokiem. Skutkiem jest publiczna ocena, zalecenia i ryzyko wizerunkowe. Ten szczebel nie podnosi werdyktu do czerwonego ani do żółtego. Czerwień wynika z ustawy o nieuczciwych praktykach."
    )

    add_h("6. Rekomendacje naprawcze")
    callout(
        "ETAP 1 – NATYCHMIASTOWY, do 14 listopada 2026 r.",
        [
            "Zdjąć albo zastąpić wzorami z przypadków 01–04: odpowiedź FAQ o 120 minutach, zdania bloga o fitoncydach jako przyczynie spadku kortyzolu, o 12,4% przy Parku 2010, o stanach zapalnych, o „udowodnieniu” przez Kaplana, o redukcji objawów lęku i depresji oraz o skutecznym zapobieganiu wypaleniu, zdanie o trwałej regulacji układu nerwowego oraz zdanie o właścicielce fundacji Leśne Bycie.",
            "Przy każdym „ZAPISZ SIĘ” i przy cenie podać imię, nazwisko albo firmę oraz adres świadczącego usługę. Do czasu korekty nie przyjmować nowych zapisów opartych na tych czterech zdaniach.",
        ],
        "FFF1F2", "E11D48", DANGER,
    )
    callout(
        "ETAP 2 – PRZED 27 WRZEŚNIA 2027 r.",
        [
            "Przejrzeć „ekologię” i „ekologię głęboką”. Zostawić je tylko jako nazwę konkretnej czynności, na przykład rozmowy o relacji z lasem podczas warsztatu, albo usunąć.",
            "Nie budować z nich twierdzenia, że usługa, marka albo cała działalność ma wysoką efektywność środowiskową. Od 27 września 2027 r. takie ogólne twierdzenie wejdzie w krajowe zakazy ustawy Dz.U. 2026 poz. 1353. Dziś nie jest osobnym naruszeniem.",
        ],
        "FFFBEB", "D97706", WARNING,
    )
    callout(
        "ETAP 3 – KOREKTA MATERIAŁÓW I METODYKA DOWODU, do 7 grudnia 2026 r.",
        [
            "Założyć kartę każdego zdania, które zostaje: pełny opis bibliograficzny, czas ekspozycji, liczba i wiek uczestników, wskaźnik, grupa porównawcza i zdanie, że sesja Leśnego Bycia nie jest tym protokołem.",
            "Politykę prywatności napisać od nowa pod formularz kontaktu i newsletter. Usunąć PESEL, płeć, adres dostawy i opis sklepu, jeżeli tych procesów nie ma. Administrator ma być tą osobą albo tym podmiotem, który formularz naprawdę otrzymuje.",
            "Sprawdzić niemaskowany odpis KRS z dnia korekty i wpisać funkcję Joanny Krusiec tak, jak stoi w odpisie.",
        ],
        "F8FAFC", "64748B", NAVY,
    )
    callout(
        "ETAP 4 – SFORMUŁOWANIA STAŁE I SZKOLENIE, do 31 stycznia 2027 r.",
        [
            "Ustalić krótką listę zdań dozwolonych: opis przebiegu spotkania, dystans 1–2 km, czas, cena, imię i adres świadczącego, oraz – jeżeli badania zostają – wzór z przypadku 01 i 02.",
            "Osobno zapisać, czego strona nie mówi: braku obietnicy odporności, ciśnienia, stanów zapalnych, depresji, wypalenia i trwałej regulacji układu nerwowego oraz braku utożsamiania marki z fundacją.",
            "To samo obowiązuje w mailu do firm, w newsletterze i w mediach społecznościowych, jeżeli powtarzają zdania ze strony. Szkolenie dotyczy jednej osoby prowadzącej i każdej osoby, która publikuje w jej imieniu.",
        ],
        "F0FDFA", "0F766E", RGBColor(15, 118, 110),
    )

    add_h("7. Wnioski")
    add_p(
        "W dniu 7 listopada 2026 r. witryna lesnebycie.pl nie powinna zostać w obecnej formie, bo FAQ, blog, program za 1500–2100 zł oraz zdanie o fundacji spełniają art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym. "
        "Czarna lista EmpCo zacznie obowiązywać w Polsce 27 września 2027 r. i nie jest podstawą tego werdyktu. Słowa o ekologii wymagają zawężenia przed tą datą, a social washing fundacji wymaga korekty już teraz."
    )

    out = ROOT / "Raport_Audytu_Greenwashingu_Lesne_Bycie.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build_material())
    print(build_report())
