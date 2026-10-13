#!/usr/bin/env python3
"""Generator raportu audytu greenwashingu dla ESG INSTITUTE.

Zgodny ze specyfikacją INSTRUKCJA_GENEROWANIA_RAPORTOW.md oraz wzorcem
wzory/WZOR_raport_Libra_Partners.docx i LOREAL/generuj_dokumenty.py.
"""

from pathlib import Path

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = ROOT.parent
LOGO = PROJECT_ROOT.parent / "ESG Hero Day - Materialy Druk" / "assets" / "esg_logo_white.png"
SHOT = PROJECT_ROOT / "assets" / "esg_institute_strona.png"


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
        "Zbiór haseł, deklaracji ofertowych, broszur i regulaminów zebranych ze strony esginstitute.eu. "
        "Plik jest zapisem źródłowym do audytu greenwashingu. Nie zawiera oceny prawnej.",
        10.5, False, True, MUTED, after=8,
    )

    rows = [
        ("Podmiot", "ESG Institute Sp. z o.o., Rondo ONZ 1, 00-124 Warszawa. NIP: 5252970752."),
        ("Oficjalna witryna", "https://esginstitute.eu/"),
        ("Materiały źródłowe", "Strona główna, 13 podstron ofertowych, regulamin Collaboration Hub, broszura główna PDF, case study Circulens/Prosave DPP, prezentacje Kobiety w ESG / ESG Ready."),
        ("Data zebrania", "7–13 października 2026 r."),
        ("Cel zebrania", "Wsad źródłowy do audytu zgodności z dyrektywą (UE) 2024/825 (ECGT), UCPD, UPNPR i UOKiK."),
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

    h("1. Hasła wizerunkowe i misja ze strony głównej")
    p("„Misją ESG Institute jest «Tworzyć w pełni zrównoważoną przyszłość, już dziś» poprzez inspirowanie, edukowanie i wspieranie firm, społeczności oraz innych interesariuszy w budowaniu lepszej przyszłości.”")
    p("„Nasze działania opierają się na holistycznym podejściu, które pozwala na głębokie zrozumienie potrzeb i profilu klienta, a następnie dobór najwłaściwszych rozwiązań dla osiągnięcia zrównoważonego sukcesu.”")
    p("„Wierzymy, że razem możemy zmienić świat i że jest to w naszym zasięgu, aby jutro było lepsze niż wczoraj, a dzisiaj to najlepszy moment, aby rozpocząć transformację w kierunku czystej, stabilnej przyszłości.”")

    h("2. Oferta audytu i deklaracja certyfikacji (Broszura oraz esginstitute.eu/esg-cert)")
    p("Z listu Prezeski Agnieszki Orłowskiej w broszurze głównej (strona 2):")
    p("„Przeprowadzimy audyt ESG Państwa firmy i pomożemy uzyskać certyfikaty potwierdzające spełnienie najwyższych standardów.”")
    p("Podstrona oferty: https://esginstitute.eu/esg-cert (Menu: OFERTA / USŁUGA – ESG Cert).")
    p("Opis audytu pogłębionego: „Pogłębiona analiza pod kątem wiarygodności spółki w zakresie ESG (analiza zgodności z regulacjami)”")

    h("3. Pakiety sponsorskie (Prezentacja Kobiety w ESG / ESG Ready)")
    p("Pakiety sponsorskie dla firm komercyjnych:")
    p("„PAKIET BRĄZOWY 5 000 zł: Tytuł Partnera Wydarzenia (...) Prawo do używania logo ESG Institute na stronie internetowej Sponsora.”")
    p("„PAKIET SREBRNY 10 000 zł: Tytuł Partnera Wydarzenia (...) Prawo do używania logo ESG Institute na stronie internetowej Sponsora.”")
    p("„PAKIET ZŁOTY 15 000 zł (GŁÓWNY SPONSOR): Tytuł Sponsora Wydarzenia (...) Prawo do używania logo ESG Institute na stronie internetowej Sponsora.”")

    h("4. Deklaracje dotyczące offsetów i partnerstw leśnych")
    p("Z biogramu dr Beaty Latos (strona główna):")
    p("„W ESG Institute Beata wnosi rozległą wiedzę i doświadczenie w zakresie strategii adaptacyjnych do zmian klimatu, planów dekarbonizacji oraz analiz wpływu środowiskowego przedsięwzięć i propozycji offsetu.”")
    p("Z opisu partnera Fundacja Las Na Zawsze (Collaboration Hub):")
    p("„Fundacja Las na Zawsze zajmuje się ochroną ekosystemów i przywracaniem bioróżnorodności poprzez oddawanie ziemi naturze. Organizacja sadzi nowe, naturalne lasy liściaste (...) Jej działania pozwalają firmom realizować cele z zakresu odpowiedzialności środowiskowej (ESG), realnie wpływając na regenerację przyrody i budowę trwałych zasobów naturalnych dla przyszłych pokoleń.”")

    h("5. Promowanie partnera Libra Partners (Collaboration Hub)")
    p("„Libra Partners produkuje ekologiczne palety z recyklingu, zmniejszając emisję CO2 w łańcuchach dostaw. Wprowadziła również innowacyjną paletę antybakteryjną, skierowaną głównie do branży farmaceutycznej, oferując preferencyjne warunki współpracy z kodem «ESG Institute & Libra Partners».”")

    h("6. Deklaracje dotyczące platformy Green0meter (esginstitute.eu/green0meter-2)")
    p("„Bezpieczeństwo regulacyjne: Gwarancja przygotowania firmy do wymogów prawnych (CSRD, EUDR, NIS2 etc.), chroniąc przed ryzykiem kar finansowych i wykluczenia z rynku.”")
    p("„Spełnienie Wymagań Regulacyjnych: Green0meter pomaga firmom spełniać wymagania międzynarodowych regulacji dotyczących raportowania i zrównoważonego rozwoju, co pozwala uniknąć ryzyka prawnego i reputacyjnego.”")

    h("7. Deklaracje modelu „ESG as a service” (Oferta CSRD i MŚP)")
    p("„przejmując obowiązki raportowania od klienta i wykonując je w imieniu klienta na bazie otrzymanych informacji korzystając z własnego zespołu i narzędzi informatycznych («ESG as a service»).”")

    h("8. Case study Cyfrowego Paszportu Produktu (Circulens / Prosave)")
    p("„CZERWIEC 2027: Bezwzględny wymóg prawny UE (Digital Product Passport). Czerwiec 2027: Ostatni dzwonek na adaptację.”")
    p("„Przewaga Konkurencyjna: Darmowy Moduł PCF. Kalkulacja śladu węglowego to obecnie najdroższy element compliance. Circulens całkowicie zmienia model biznesowy tego segmentu (...) Koszt: 0 PLN (W pełni darmowy w ramach platformy). Zastosowanie: Czyni system paszportowy opłacalnym biznesowo. Może być licencjonowany oddzielnie jako samodzielne narzędzie (szczególnie istotne dla zgodności z CBAM).”")

    h("9. Harmonogram dyrektywy CSRD (esginstitute.eu/csrd-2)")
    p("„Za rok 2024: Firmy, które obecnie podlegają dyrektywie NFRD czyli zatrudniające >500 osób i będące jednostkami zainteresowania publicznego.”")
    p("„Za rok 2027: Firmy, które spełniają kryterium >1000 pracowników i obroty netto >450M EUR.”")
    p("„Za rok 2028: Małe i średnie podmioty notowane niespełniające warunku 1000 pracowników wyłączone z raportowania.”")

    h("10. Deklaracje walki z greenwashingiem (Kobiety w ESG)")
    p("„1. Koniec «Green-Pretend»: Mówimy «nie» pustym hasłom. Naszym filtrem jest radykalna praktyka. Wolimy trudną prawdę o błędach niż wygładzone slajdy o nieistniejących sukcesach.”")
    p("„Pokazujemy ESG w działaniu (...) Bez greenwashingu i pustych teorii. Z naciskiem na praktykę i zrównoważony rozwój.”")

    out = ROOT / "Material_ESG_Institute_tresci_do_audytu.docx"
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

    # Header and kicker
    add_p("ESG INSTITUTE  ·  RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO", size=9, bold=True, color=ACCENT, after=2)
    add_p("Audyt Ryzyk Greenwashingu: ESG INSTITUTE", size=20, bold=True, color=PRIMARY, before=2, after=2)
    add_p(
        "Weryfikacja treści oficjalnej witryny https://esginstitute.eu/, oferty doradczej, broszury głównej, "
        "materiałów platformy Green0meter, case study Cyfrowego Paszportu Produktu (Circulens/Prosave) oraz programu ESG Ready "
        "wobec dyrektywy (UE) 2024/825 (ECGT), dyrektywy 2005/29/WE (UCPD), ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (UPNPR), "
        "ustawy o ochronie konkurencji i konsumentów oraz ustawy o zwalczaniu nieuczciwej konkurencji (UZNK). "
        "Stan zebrania: 7–13 października 2026 r.",
        size=10.0, italic=True, color=MUTED, after=8,
    )

    if SHOT.exists():
        shot_par = doc.add_paragraph()
        shot_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shot_par.paragraph_format.space_before = Pt(2)
        shot_par.paragraph_format.space_after = Pt(2)
        shot_par.add_run().add_picture(str(SHOT), width=Cm(15.4))
        add_p(
            "Zrzut witryny https://esginstitute.eu/ z 13 października 2026 r. Hasło wiodące: „Tworzyć w pełni zrównoważoną przyszłość, już dziś”.",
            size=8.5, italic=True, color=MUTED, after=8,
        )

    # Metadata table
    meta = [
        ("Badany podmiot", "ESG Institute Sp. z o.o., Rondo ONZ 1, 00-124 Warszawa, NIP: 5252970752. Podmiot doradczo-szkoleniowy w formie spółki handlowej."),
        ("Przedmiot audytu", "Komunikacja marketingowa i ofertowa na oficjalnej witrynie esginstitute.eu (strona główna, 13 podstron usługowych: abonament, academy, audytesg, csrd, doradztwo, ecovadis, esg-cert, green0meter, lca, msp, slad-weglowy, hub), broszura ofertowa PDF, case study DPP Prosave/Circulens oraz prezentacje programowe (Kobiety w ESG, ESG Ready)."),
        ("Sektor działalności", "Usługi profesjonalne, doradztwo biznesowe i prawne, edukacja i szkolenia (B2B / think tank) w obszarze Environmental, Social, and Governance (ESG)."),
        ("Reżim prawny", "Dyrektywa (UE) 2024/825 (ECGT – stosowanie od 27.09.2026), Dyrektywa 2005/29/WE (UCPD), UPNPR (art. 4–7), ustawa o ochronie konkurencji i konsumentów (art. 24 i 106, kary do 10% obrotu), UZNK (art. 16), normy ISO 14021 / 14064 / 14040 / 14044 / GHG Protocol, ESRS (EFRAG) oraz Taksonomia UE."),
        ("Data audytu", "13 października 2026 r."),
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

    # Callout werdyktu
    callout(
        "WERDYKT: NIE PUBLIKOWAĆ W OBECNEJ FORMIE (KOD CZERWONY / REJECT & RE-NARROW)",
        [
            "Badany materiał spełnia bezpośrednie przesłanki czarnej listy Załącznika I do Dyrektywy 2005/29/WE w brzmieniu nadanym dyrektywą (UE) 2024/825 (ECGT): pkt 2a (komercyjna sprzedaż prawa do używania logo ESG Institute w płatnych pakietach sponsorskich bez systemu certyfikacji strony trzeciej, obietnica uzyskania certyfikatów przez spółkę doradczą oraz pusta podstrona «ESG Cert»), pkt 4a (ogólne twierdzenia środowiskowe: «w pełni zrównoważona przyszłość», «zrównoważony sukces», «ekologiczne palety»), pkt 4b (promowanie partnerów stosujących greenwashing poprzez transfer autorytetu i claimy o całości produktu) oraz pkt 4c (rekomendowanie propozycji offsetu i sadzenia drzew jako instrumentu realizacji celów ESG i dekarbonizacji).",
            "Materiały zawierają ponadto bezpodstawne zapewnienia o prawnej «gwarancji ochrony przed karami finansowymi» generowanej przez oprogramowanie AI (UPNPR art. 5 ust. 1 i ust. 2 pkt 1), wprowadzające w błąd deklaracje o «przejęciu obowiązków raportowania od klienta» (odpowiedzialność prawna z mocy dyrektywy CSRD i ustawy o rachunkowości jest niezbywalna i spoczywa na organach kierowniczych jednostki) oraz rażące błędy merytoryczne w harmonogramie CSRD (pomieszanie z progami dyrektywy CSDD i mechanizmu CBAM).",
            "Dyrektywa ECGT stosuje się bezwzględnie od 27 września 2026 r. Treści cyfrowe (witryny internetowe, broszury online, prezentacje ofertowe) nie podlegają żadnej taryfie ochronnej dla starych zapasów. Formułowanie takich twierdzeń przez podmiot pozycjonujący się jako instytucja ekspercka i think tank compliance naraża ESG Institute Sp. z o.o. na postępowanie Prezesa UOKiK z art. 24 ustawy o ochronie konkurencji i konsumentów (zagrożenie karą do 10% obrotu), roszczenia konkurentów z art. 16 UZNK oraz krytyczne ryzyko utraty wiarygodności rynkowej.",
        ],
        "FFF1F2", "E11D48", DANGER,
    )

    # 1. Analiza Wstępna
    add_h("1. Analiza Wstępna")
    add_p(
        "ESG Institute Sp. z o.o. jest polskim podmiotem doradczo-szkoleniowym, założonym przez uznanych ekspertów rynku (Agnieszkę Orłowską oraz radcę prawnego Rafała Wiatra). "
        "Firma pozycjonuje się jako wiodący think tank, ośrodek analityczny i platforma współpracy (Collaboration Hub), oferując przedsiębiorstwom kompleksowe wsparcie "
        "w zakresie strategii ESG, raportowania niefinansowego (CSRD, VSME), audytów zgodności, obliczeń śladu węglowego oraz wdrożeń technologicznych (Green0meter, Circulens). "
        "W swojej komunikacji instytut stawia na narrację bezkompromisowej rzetelności, profesjonalizmu i etyki, deklarując walkę z rynkową „eko-ściemą” pod hasłami: „Koniec Green-Pretend” oraz „Bez greenwashingu i pustych teorii”."
    )
    add_p(
        "Szczegółowy audyt regulacyjny materiałów opublikowanych na witrynie esginstitute.eu, powiązanych dokumentów PDF oraz materiałów konferencyjnych ujawnia jednak zjawisko, "
        "które w doktrynie compliance określa się mianem „paradoksu audytora”. Podmiot edukujący rynek z zakresu unikania greenwashingu sam dopuszcza się w swojej ofercie "
        "i materiałach promocyjnych poważnych uchybień regulacyjnych. Najbardziej jaskrawym naruszeniem jest komercyjna sprzedaż prawa do używania logo ESG Institute "
        "w pakietach sponsorskich konferencji (od 5 000 do 15 000 PLN). Z perspektywy prawa ochrony konsumentów stanowi to nieuprawnioną komercjalizację znaku zrównoważoności (sustainability label) "
        "bez jakiejkolwiek procedury certyfikacji czy weryfikacji sponsora (naruszenie pkt 2a czarnej listy ECGT)."
    )
    add_p(
        "Drugim kluczowym obszarem ryzyka jest promowanie partnerów komercyjnych stosujących agresywny greenwashing (w tym Libra Partners, oferującej tzw. „ekologiczne palety z recyklingu” "
        "z powłoką antybakteryjną opartą na biocydach bez rejestracji BPR, zaopatrzone we wspólny kod rabatowy) oraz rekomendowanie firmom „propozycji offsetu” i sadzenia lasów "
        "jako drogi do „realizacji celów środowiskowych ESG”. Zgodnie z dyrektywą ECGT (pkt 4c) oraz standardem ESRS E1, offsety nie mogą stanowić podstawy twierdzeń o neutralności "
        "ani pomniejszać raportowanych emisji brutto. Ponadto w ofercie zidentyfikowano obietnice prawne niemające pokrycia w rzeczywistości (np. „gwarancja ochrony przed karami” przez oprogramowanie AI "
        "czy „przejęcie obowiązków raportowania od klienta”). Poniższa analiza szczegółowa klasyfikuje wykryte naruszenia i definiuje bezwzględny plan naprawczy."
    )

    # 2. Analiza szczegółowa - tabela zbiorcza
    add_h("2. Analiza szczegółowa - tabela zbiorcza naruszeń")
    findings = [
        ("01", "Sprzedaż prawa do logo ESG Institute w pakietach sponsorskich bez certyfikacji", "ECGT zał. I pkt 2a · UPNPR art. 5 i art. 7 pkt 2", "BARDZO WYSOKIE", "Komercyjna sprzedaż znaku zrównoważoności za 5–15 tys. zł bez audytu sponsora"),
        ("02", "Promowanie propozycji offsetu i sadzenia drzew jako realizacji celów ESG", "ECGT zał. I pkt 4c · ESRS E1 · UPNPR art. 5", "BARDZO WYSOKIE", "Offsety i sadzenie lasu prezentowane jako instrument realizacji celów dekarbonizacji"),
        ("03", "Promocja partnera Libra Partners i hasła „ekologiczne palety z recyklingu”", "ECGT zał. I pkt 4a i 4b · UPNPR art. 5 i art. 6", "BARDZO WYSOKIE", "Transfer autorytetu instytutu na produkt greenwashingowy z wadami biocydowymi"),
        ("04", "Oferowanie „certyfikatów najwyższych standardów” i podstrony «ESG Cert»", "ECGT zał. I pkt 2a · UPNPR art. 5 ust. 3 pkt 2", "WYSOKIE", "Kreowanie wrażenia bycia jednostką certyfikującą przez spółkę z o.o. bez akredytacji"),
        ("05", "Twierdzenie o „gwarancji ochrony przed karami finansowymi” (Green0meter)", "UPNPR art. 5 ust. 1 i ust. 2 pkt 1 · UCPD art. 6", "WYSOKIE", "Niedopuszczalne prawnie zapewnienie o gwarancji braku kar przez oprogramowanie AI"),
        ("06", "Deklaracja „przejęcia obowiązków raportowania od klienta” („ESG as a service”)", "UPNPR art. 5 · Dyrektywa CSRD (UE) 2022/2464", "WYSOKIE", "Wprowadzenie w błąd: odpowiedzialność prawna za CSRD spoczywa wyłącznie na zarządzie"),
        ("07", "Ogólne claimy misji: „w pełni zrównoważona przyszłość” i „zrównoważony sukces”", "ECGT zał. I pkt 4a · UPNPR art. 5", "WYSOKIE", "Ogólne hasła środowiskowe bez uznanej wysokiej efektywności ekologicznej"),
        ("08", "Czerwiec 2027 jako „bezwzględny wymóg DPP” oraz moduł PCF „dla zgodności z CBAM”", "UPNPR art. 5 ust. 1 · ESPR 2024/1781 · CBAM 2023/956", "ŚREDNIE", "Przedwczesne sztywne daty przed aktami delegowanymi i mylenie odzieży z reżimem CBAM"),
        ("09", "Chaos merytoryczny i błędy w harmonogramie CSRD oraz wyłączeniach MŚP", "UPNPR art. 5 ust. 1 · Dyrektywy CSRD i CS3D", "ŚREDNIE", "Mylenie progów CS3D (>1000 osób, >450M EUR) z CSRD i dezinformacja o notowanych MŚP"),
        ("10", "Deklaracje „Koniec Green-Pretend” przy braku własnego raportowania ESG", "UPNPR art. 5 i art. 6 · Norma ISO 14021 pkt 5.7", "ŚREDNIE", "Asymetria narracyjna: zwalczanie greenwashingu bez ujawniania własnych danych emisji"),
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

    # 3. Szczegółowa analiza przypadków
    add_h("3. Szczegółowa analiza przypadków")

    cases = [
        {
            "title": "PRZYPADEK 01: Komercyjna sprzedaż prawa do logo ESG Institute w pakietach sponsorskich bez certyfikacji",
            "basis": "Podstawa prawna: ECGT, Załącznik I pkt 2a · UPNPR art. 5 ust. 3 pkt 2 oraz art. 7 pkt 2",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (Czarna lista ECGT)",
            "quotes": [
                "„Pakiety Sponsorskie - Przegląd (...) PAKIET BRĄZOWY 5 000 zł: Prawo do używania logo ESG Institute na stronie internetowej Sponsora.”",
                "„PAKIET SREBRNY 10 000 zł: Prawo do używania logo ESG Institute na stronie internetowej Sponsora.”",
                "„GŁÓWNY SPONSOR - WYŁĄCZNOŚĆ PAKIET ZŁOTY 15 000 zł: Prawo do używania logo ESG Institute na stronie internetowej Sponsora.”",
            ],
            "assessment": (
                "Załącznik I pkt 2a do dyrektywy 2005/29/WE (w brzmieniu ECGT 2024/825) kwalifikuje jako bezwzględnie zakazaną praktykę "
                "«umieszczanie oznakowania dotyczącego zrównoważonego charakteru, które nie jest oparte na systemie certyfikacji ani nie zostało ustanowione przez organy publiczne». "
                "Logo ESG Institute w obrocie gospodarczym komunikuje wiedzę, wiarygodność i wysokie standardy zrównoważonego rozwoju. "
                "Zaoferowanie komercyjnym przedsiębiorstwom prawa do umieszczania tego logo na ich oficjalnych stronach internetowych w zamian za wpłatę sponsoringową "
                "(5 000 – 15 000 zł), bez uprzedniego przeprowadzenia weryfikacji i audytu środowiskowego sponsora w ramach publicznego lub certyfikowanego schematu, "
                "stanowi odpłatną sprzedaż znaku zrównoważoności. Sponsor uzyskuje możliwość sugerowania swoim klientom i kontrahentom, że jego działalność "
                "posiada akredytację lub pozytywną ocenę ESG Institute, co bezpośrednio wyczerpuje znamiona czynu z art. 7 pkt 2 UPNPR oraz pkt 2a ECGT."
            ),
            "remediation": (
                "Natychmiast wykreślić ze wszystkich ofert i prezentacji sponsorskich (w tym konferencji Kobiety w ESG i programu ESG Ready) "
                "świadczenie «Prawo do używania logo ESG Institute na stronie internetowej Sponsora». Sponsor może otrzymać prawo do informacji o sponsorowaniu "
                "konkretnego wydarzenia («Sponsor konferencji Kobiety w ESG 2026»), lecz bez prawa do prezentowania logo instytutu na własnej stronie w sposób sugerujący certyfikację."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Firma [Nazwa] jest sponsorem konferencji Kobiety w ESG 2026 organizowanej przez ESG Institute Sp. z o.o. "
                "Status sponsora dotyczy wyłącznie wsparcia organizacyjnego wydarzenia i nie oznacza certyfikacji, audytu ani oceny działalności firmy pod kątem standardów ESG.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 02: Promowanie propozycji offsetu emisji i sadzenia drzew jako instrumentu realizacji celów ESG",
            "basis": "Podstawa prawna: ECGT, Załącznik I pkt 4c · Standard ESRS E1 pkt AR 55–58 · UPNPR art. 5",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (Czarna lista ECGT)",
            "quotes": [
                "„W ESG Institute Beata wnosi rozległą wiedzę i doświadczenie w zakresie strategii adaptacyjnych do zmian klimatu, planów dekarbonizacji oraz analiz wpływu środowiskowego przedsięwzięć i propozycji offsetu.”",
                "„Fundacja Las na Zawsze (...) sadzi nowe, naturalne lasy liściaste (...) Jej działania pozwalają firmom realizować cele z zakresu odpowiedzialności środowiskowej (ESG), realnie wpływając na regenerację przyrody i budowę trwałych zasobów naturalnych dla przyszłych pokoleń.”",
            ],
            "assessment": (
                "Dyrektywa ECGT w pkt 4c Załącznika I wprowadza całkowity zakaz formułowania twierdzeń, że działalność lub produkt ma neutralny, "
                "zmniejszony lub pozytywny wpływ na środowisko pod względem emisji GHG, w oparciu o kompensację emisji (carbon offsets, sadzenie drzew). "
                "Włączenie do oferty instytutu doradztwa w zakresie «propozycji offsetu» oraz wskazywanie na stronie Hubu, że sadzenie drzew pozwala firmom "
                "«realizować cele z zakresu odpowiedzialności środowiskowej (ESG)», sankcjonuje kompensację jako substytut transformacji niskoemisyjnej. "
                "Zgodnie z GHG Protocol Corporate Standard oraz wytycznymi EFRAG (ESRS E1), offsety węglowe i zalesianie nie mogą pomniejszać wielkości "
                "emisji brutto w Scope 1, Scope 2 i Scope 3 i nie mogą być raportowane jako redukcja śladu węglowego w łańcuchu wartości."
            ),
            "remediation": (
                "Usunąć termin «propozycje offsetu» z opisu kompetencji ekspertów i oferty dekarbonizacyjnej. W opisach współpracy z Fundacją Las Na Zawsze "
                "wyraźnie zastrzec, że wsparcie projektów leśnych ma charakter filantropijnej darowizny na rzecz bioróżnorodności i nie stanowi realizacji celów redukcyjnych GHG ani kompensacji emisji."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„ESG Institute wspiera przedsiębiorstwa w opracowywaniu planów redukcji emisji gazów cieplarnianych w Scope 1, 2 i 3 zgodnie z GHG Protocol. "
                "Współpraca z organizacjami społecznymi w zakresie ochrony lasów stanowi wsparcie bioróżnorodności i nie stanowi kompensacji emisji (offsetu) w rozumieniu ESRS E1.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 03: Promocja partnera Libra Partners i legitymizowanie hasła „ekologiczne palety z recyklingu”",
            "basis": "Podstawa prawna: ECGT, Załącznik I pkt 4a i 4b · UPNPR art. 5 i art. 6 · UZNK art. 16",
            "risk": "Poziom ryzyka: BARDZO WYSOKIE (Czarna lista i transfer autorytetu)",
            "quotes": [
                "„Libra Partners produkuje ekologiczne palety z recyklingu, zmniejszając emisję CO2 w łańcuchach dostaw. Wprowadziła również innowacyjną paletę antybakteryjną, skierowaną głównie do branży farmaceutycznej, oferując preferencyjne warunki współpracy z kodem «ESG Institute & Libra Partners».”",
            ],
            "assessment": (
                "ESG Institute na swojej witrynie promuje Libra Partners jako zaufanego Partnera Eksperckiego z dedykowanym kodem rabatowym, "
                "firmując hasło «produkuje ekologiczne palety z recyklingu, zmniejszając emisję CO2». Zgodnie z audytem Libra Partners, sformułowanie «ekologiczne palety» "
                "stanowi ogólne twierdzenie środowiskowe bez certyfikatu typu I (naruszenie pkt 4a ECGT). Ponadto promowana paleta antybakteryjna oparta jest na biocydach, "
                "które nie posiadały wymaganych prawem rejestracji w Wykazie Produktów Biobójczych (BPR), a deklaracje redukcji CO2 oparto na manipulacji wskaźnikowej. "
                "Publikując ten opis, ESG Institute dopuszcza się greenwashingu przez transfer autorytetu: jako think tank legitymizuje nierzetelne deklaracje środowiskowe podmiotu trzeciego (art. 5 UPNPR)."
            ),
            "remediation": (
                "Natychmiast usunąć ze strony Collaboration Hub niezweryfikowane hasła marketingowe Libry («ekologiczne palety»). Wdrożyć wewnętrzną procedurę due diligence "
                "partnerów Hubu, wykluczającą promowanie podmiotów stosujących greenwashing i oferowanie wspólnych kodów promocyjnych bez weryfikacji kart technicznych i badań LCA."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Libra Partners oferuje palety logistyczne z tworzyw sztucznych produkowane z udziałem surowców wtórnych. "
                "Szczegółowa zawartość recyklatu, parametry techniczne oraz deklaracje zgodności z normami dostępne są w dokumentacji technicznej producenta.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 04: Oferowanie „certyfikatów potwierdzających najwyższe standardy” oraz podstrony «ESG Cert»",
            "basis": "Podstawa prawna: ECGT, Załącznik I pkt 2a · UPNPR art. 5 ust. 3 pkt 2 oraz art. 7 pkt 2",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista ECGT)",
            "quotes": [
                "„Przeprowadzimy audyt ESG Państwa firmy i pomożemy uzyskać certyfikaty potwierdzające spełnienie najwyższych standardów.”",
                "„https://esginstitute.eu/esg-cert (W menu strony: OFERTA / USŁUGA – ESG Cert)”",
                "„Audyt pogłębiony ESG: Pogłębiona analiza pod kątem wiarygodności spółki w zakresie ESG (analiza zgodności z regulacjami)”",
            ],
            "assessment": (
                "Zapewnienie w broszurze, że instytut pomoże uzyskać certyfikaty potwierdzające spełnienie najwyższych standardów, w powiązaniu z zakładką «ESG Cert» "
                "w menu witryny, wywołuje u klientów mylne wrażenie, że ESG Institute jest akredytowaną jednostką certyfikującą (Conformity Assessment Body). "
                "ESG Institute Sp. z o.o. jest podmiotem komercyjnym świadczącym doradztwo biznesowe, a nie jednostką akredytowaną przez Polskie Centrum Akredytacji (PCA) "
                "w rozumieniu rozporządzenia (WE) nr 765/2008. Wystawianie własnych «certyfikatów ESG» przez spółkę doradczą lub sugerowanie takiego rezultatu audytu "
                "narusza zakaz stosowania prywatnych znaków zrównoważoności bez publicznego schematu certyfikacji (pkt 2a ECGT) i wprowadza w błąd co do uprawnień (art. 7 pkt 2 UPNPR)."
            ),
            "remediation": (
                "Zastąpić obietnicę «uzyskania certyfikatów» sformułowaniem «przygotowanie do niezależnej atestacji lub certyfikacji zewnętrznej». "
                "Usunąć lub przedefiniować zakładkę «ESG Cert» na «Wsparcie w certyfikacji», precyzując, że instytut przygotowuje dokumentację, a certyfikację przeprowadzają uprawnione podmioty trzecie."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Przeprowadzimy audyt doradczy ESG Państwa organizacji, oceniając zgodność z regulacjami i przygotowując firmę do niezależnej atestacji sprawozdawczości lub certyfikacji przez akredytowane jednostki zewnętrzne.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 05: Twierdzenie o „gwarancji ochrony przed ryzykiem kar finansowych” (Green0meter)",
            "basis": "Podstawa prawna: UPNPR art. 5 ust. 1 i ust. 2 pkt 1 · Dyrektywa 2005/29/WE art. 6",
            "risk": "Poziom ryzyka: WYSOKIE (Wprowadzenie w błąd co do cech i skutków usługi)",
            "quotes": [
                "„Bezpieczeństwo regulacyjne: Gwarancja przygotowania firmy do wymogów prawnych (CSRD, EUDR, NIS2 etc.), chroniąc przed ryzykiem kar finansowych i wykluczenia z rynku.”",
                "„Spełnienie Wymagań Regulacyjnych: Green0meter pomaga firmom spełniać wymagania międzynarodowych regulacji (...) co pozwala uniknąć ryzyka prawnego i reputacyjnego.”",
            ],
            "assessment": (
                "Użycie kategorycznego terminu «Gwarancja przygotowania (...) chroniąc przed ryzykiem kar finansowych» w ofercie dystrybucji oprogramowania Green0meter "
                "jest prawnie niedopuszczalnym zapewnieniem. Żadne narzędzie informatyczne ani algorytmy AI nie gwarantują zgodności prawnej i nie zdejmują ryzyka kar administracyjnych. "
                "Wymogi dyrektyw CSRD, EUDR czy NIS2 zależą od faktycznych procesów operacyjnych w przedsiębiorstwie, rzetelności wprowadzanych danych oraz oceny organów nadzorczych i biegłych rewidentów. "
                "Obietnica gwarancji ochrony przed sankcjami kreuje fałszywe poczucie bezpieczeństwa prawnego i stanowi rozpowszechnianie informacji wprowadzających w błąd (art. 5 UPNPR)."
            ),
            "remediation": (
                "Bezwzględnie wykreślić słowo «Gwarancja». Zastąpić je sformułowaniem wskazującym na wsparcie w automatyzacji procesów compliance i analizie luk regulacyjnych, "
                "z wyraźnym zastrzeżeniem, że oprogramowanie jest narzędziem analitycznym, a odpowiedzialność za zgodność ponosi zarząd raportującej jednostki."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Wsparcie regulacyjne: Platforma Green0meter wspiera organizację w strukturyzacji danych pod kątem wymogów CSRD, EUDR i NIS2, ułatwiając monitorowanie luk prawnych i przygotowanie dokumentacji.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 06: Deklaracja „przejęcia obowiązków raportowania od klienta” w modelu „ESG as a service”",
            "basis": "Podstawa prawna: UPNPR art. 5 · Dyrektywa CSRD (UE) 2022/2464 · Ustawa o rachunkowości art. 4a i 49b",
            "risk": "Poziom ryzyka: WYSOKIE (Wprowadzenie w błąd co do odpowiedzialności prawnej)",
            "quotes": [
                "„przejmując obowiązki raportowania od klienta i wykonując je w imieniu klienta na bazie otrzymanych informacji korzystając z własnego zespołu i narzędzi informatycznych («ESG as a service»).”",
            ],
            "assessment": (
                "Zgodnie z dyrektywą CSRD oraz przepisami prawa krajowego (art. 4a ustawy o rachunkowości), członkowie zarządu oraz rady nadzorczej jednostki "
                "ponoszą solidarną, ustawową odpowiedzialność za sporządzenie sprawozdania zrównoważonego rozwoju. Odpowiedzialność ta ma charakter publicznoprawny "
                "i jest niezbywalna – nie może zostać w drodze umowy przeniesiona na zewnętrznego doradcę ani zewnętrzną firmę technologiczną. "
                "Twierdzenie, że ESG Institute «przejmuje obowiązki raportowania od klienta», wprowadza kadrę zarządzającą w błąd co do ich osobistej odpowiedzialności prawnej i stanowi nieuczciwą praktykę rynkową."
            ),
            "remediation": (
                "Zmienić opis usługi «ESG as a service»: usunąć zwrot «przejmując obowiązki raportowania». Zdefiniować usługę jako «zewnętrzne wsparcie operacyjne, "
                "analityczne i techniczne w sporządzeniu projektu raportu ESG na rzecz klienta»."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„W modelu operacyjnym zespół ESG Institute realizuje kompleksowe wsparcie analityczne i techniczne, przygotowując projekt raportu ESG na podstawie przekazanych danych, do formalnej akceptacji przez kierownictwo jednostki.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 07: Ogólne, bezwzględne claimy w misji i ofercie: „w pełni zrównoważona przyszłość” i „zrównoważony sukces”",
            "basis": "Podstawa prawna: ECGT, Załącznik I pkt 4a · UPNPR art. 5",
            "risk": "Poziom ryzyka: WYSOKIE (Czarna lista ECGT)",
            "quotes": [
                "„Misją ESG Institute jest «Tworzyć w pełni zrównoważoną przyszłość, już dziś» poprzez inspirowanie, edukowanie i wspieranie firm (...)”",
                "„dobór najwłaściwszych rozwiązań dla osiągnięcia zrównoważonego sukcesu”",
                "„budować pozytywny wizerunek firmy jako lidera w zrównoważonym rozwoju”",
            ],
            "assessment": (
                "Pkt 4a Załącznika I ECGT zakazuje stosowania ogólnych twierdzeń środowiskowych (takich jak «zrównoważony», «eko», «zielony», «świadomy ekologicznie»), "
                "jeżeli przedsiębiorca nie jest w stanie wykazać uznanej wysokiej efektywności ekologicznej w odniesieniu do tego twierdzenia. "
                "Formułowanie hasła «w pełni zrównoważona przyszłość» (claim o charakterze absolutnym) oraz obietnic «osiągnięcia zrównoważonego sukcesu» "
                "w ofercie komercyjnej spółki z o.o., która sama nie posiada certyfikatu typu I (np. EU Ecolabel) i nie ujawnia własnych wskaźników środowiskowych, "
                "jest stosowaniem niepopartych dowodowo ogólnych haseł marketingowych."
            ),
            "remediation": (
                "Przekształcić claimy absolutne w mierzalne deklaracje celów doradczych i edukacyjnych. Zamiast «tworzyć w pełni zrównoważoną przyszłość» "
                "wskazać na «wspieranie transformacji przedsiębiorstw w kierunku zgodności z regulacjami ESG i standardami ochrony środowiska»."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Misją ESG Institute jest wspieranie przedsiębiorstw i instytucji w procesie transformacji w kierunku zrównoważonego rozwoju poprzez profesjonalne doradztwo strategiczne, edukację i wdrażanie standardów ESG.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 08: Przedstawianie daty czerwiec 2027 jako „bezwzględnego wymogu UE dla DPP” oraz modułu PCF „dla zgodności z CBAM”",
            "basis": "Podstawa prawna: UPNPR art. 5 ust. 1 i ust. 2 pkt 1 · Rozporządzenie ESPR (UE) 2024/1781 · Rozporządzenie CBAM (UE) 2023/956",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "„CZERWIEC 2027: Bezwzględny wymóg prawny UE (Digital Product Passport). Czerwiec 2027: Ostatni dzwonek na adaptację (...)”",
                "„Przewaga Konkurencyjna: Darmowy Moduł PCF (...) Koszt: 0 PLN (...) Może być licencjonowany oddzielnie jako samodzielne narzędzie (szczególnie istotne dla zgodności z CBAM).”",
            ],
            "assessment": (
                "W broszurze Prosave/Circulens podano sztywny termin «Czerwiec 2027» jako rzekomy bezwzględny wymóg prawny paszportu produktu dla odzieży roboczej. "
                "Rozporządzenie ESPR 2024/1781 przewiduje, że wymogi DPP dla poszczególnych grup towarowych wejdą w życie dopiero po uchwaleniu aktów delegowanych "
                "przez Komisję Europejską (z 18-miesięcznym vacatio legis). Wyznaczanie sztywnej daty bez wzmianki o procesie legislacyjnym KE kreuje sztuczną presję rynkową. "
                "Ponadto powiązanie modułu odzieżowego PCF z mechanizmem CBAM jest rażącym błędem merytorycznym: CBAM obejmuje wyłącznie ściśle określone surowce importowane "
                "(stal, aluminium, cement, nawozy, wodór, energia), a nie odzież roboczą, i posiada odrębną metodykę obliczania emisji wbudowanych (Rozporządzenie 2023/1773)."
            ),
            "remediation": (
                "Skorygować harmonogram ESPR: wyjaśnić, że wdrożenie DPP dla tekstyliów nastąpi zgodnie z harmonogramem planowanych aktów delegowanych KE. "
                "Wykreślić powiązanie modułu odzieżowego PCF z CBAM lub precyzyjnie wyjaśnić, jakich towarów i jakiego reżimu dotyczy rozporządzenie CBAM."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Cyfrowy Paszport Produktu (DPP) w ramach rozporządzenia ESPR stanie się obowiązkowy dla sektora tekstylnego po przyjęciu dedykowanego aktu delegowanego Komisji Europejskiej. Moduł PCF umożliwia kalkulację śladu węglowego wyrobu zgodnie z normą ISO 14067.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 09: Chaos merytoryczny i błędy w harmonogramie CSRD oraz wyłączeniach MŚP",
            "basis": "Podstawa prawna: UPNPR art. 5 ust. 1 i ust. 2 pkt 1 · Dyrektywa CSRD (UE) 2022/2464 · Dyrektywa CS3D (UE) 2024/1760",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "„Za rok 2027: Firmy, które spełniają kryterium >1000 pracowników i obroty netto >450M EUR.”",
                "„Za rok 2028: Małe i średnie podmioty notowane niespełniające warunku 1000 pracowników wyłączone z raportowania.”",
                "„Za rok 2026: Małe i średnie podmioty notowane spełniające min. 2 z 3 kryteriów: >10 pracowników, obrót netto >700k EUR, suma aktywów >350k EUR”",
            ],
            "assessment": (
                "Na podstronie CSRD doszło do pomieszania harmonogramu dyrektywy CSRD z dyrektywą o należytej staranności w łańcuchu dostaw (CS3D / CSDD), "
                "w której próg wynosi 1000 pracowników i 450 mln EUR. Z kolei stwierdzenie, że za rok 2028 notowane MŚP niespełniające warunku 1000 pracowników "
                "«są wyłączone z raportowania», jest dezinformacją prawną – notowane MŚP podlegają CSRD od roku obrotowego 2026 (z możliwością opt-out do 2028 r.) "
                "niezależnie od kryterium 1000 pracowników. Podawanie błędnych i sprzecznych kryteriów prawnych przez podmiot doradczy wprowadza klientów w błąd (art. 5 UPNPR)."
            ),
            "remediation": (
                "Zaktualizować i ujednolicić harmonogram CSRD we wszystkich materiałach zgodnie z tekstem Dyrektywy 2022/2464 oraz dyrektywy delegowanej 2023/2775. "
                "Wyraźnie oddzielić obowiązki sprawozdawcze CSRD od obowiązków należytej staranności wynikających z dyrektywy CS3D."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„Harmonogram CSRD: za rok 2024 – duże JZP objęte dotychczas NFRD (>500 pracowników); za rok 2025 – pozostałe duże jednostki; za rok 2026 – notowane MŚP (z opcją odroczenia raportowania do 2028 r.). Obowiązki dotyczące łańcucha dostaw dla firm >1000 pracowników wynikają z odrębnej dyrektywy CS3D.”"
            ),
            "salvage": True,
        },
        {
            "title": "PRZYPADEK 10: Deklaracje „Koniec Green-Pretend” i walki z greenwashingiem przy braku własnego raportowania ESG",
            "basis": "Podstawa prawna: UPNPR art. 5 i art. 6 · Norma ISO 14021 pkt 5.7",
            "risk": "Poziom ryzyka: ŚREDNIE",
            "quotes": [
                "„1. Koniec «Green-Pretend»: Mówimy «nie» pustym hasłom. Naszym filtrem jest radykalna praktyka. Wolimy trudną prawdę o błędach niż wygładzone slajdy o nieistniejących sukcesach.”",
                "„Pokazujemy ESG w działaniu (...) Bez greenwashingu i pustych teorii. Z naciskiem na praktykę i zrównoważony rozwój.”",
            ],
            "assessment": (
                "Budowanie przekazu marketingowego w oparciu o radykalne przeciwstawianie się pozorowanym działaniom eko («Koniec Green-Pretend», «bez greenwashingu») "
                "przy jednoczesnym braku publikacji choćby podstawowego oświadczenia o własnym śladzie węglowym (Scope 1 i Scope 2), zużyciu energii czy wskaźnikach DEI "
                "tworzy asymetrię informacyjną. Zgodnie z normą ISO 14021 oraz zasadą staranności zawodowej z art. 4 ust. 1 UPNPR, przedsiębiorca profesjonalnie oceniający "
                "innych powinien sam zapewniać dostęp do transparentnych danych o własnej działalności, aby jego deklaracje etyczne nie miały charakteru pozornego."
            ),
            "remediation": (
                "Opublikować na stronie zwięzłe zestawienie danych środowiskowych i społecznych ESG Institute Sp. z o.o. (np. ślad węglowy operacji, polityka różnorodności, "
                "zasady doboru partnerów komercyjnych), potwierdzające autentyczność deklarowanej misji."
            ),
            "claim_title": "DOZWOLONE SFORMUŁOWANIE (WZÓR PO KOREKCIE)",
            "claim": (
                "„ESG Institute promuje transparentność i rzetelność w komunikacji zrównoważonego rozwoju. Informacje o naszych własnych działaniach organizacyjnych oraz profilu środowiskowym udostępniamy w corocznym komunikacie publicznym.”"
            ),
            "salvage": True,
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
        _set_run(
            meta_line.add_run(case["basis"] + "  |  " + case["risk"]),
            8.5, True, False,
            DANGER if "BARDZO" in case["risk"] or "WYSOKIE" in case["risk"] else WARNING,
        )
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

    # 4. Porównanie twierdzeń z rzeczywistym stanem
    add_h("4. Porównanie twierdzeń ESG INSTITUTE z rzeczywistym stanem prawno-środowiskowym")
    add_p(
        "Zestawienie twierdzeń ofertowych i marketingowych stosowanych przez ESG Institute Sp. z o.o. "
        "z obiektywnym stanem regulacyjnym, normatywnym i metodycznym.",
        after=6,
    )
    comp = [
        (
            "Pakiety sponsorskie: „Prawo do używania logo ESG Institute na stronie internetowej Sponsora” za 5 000 – 15 000 zł.",
            "Zgodnie z ECGT pkt 2a oznakowanie zrównoważonego charakteru wymaga certyfikacji strony trzeciej lub organu publicznego. Komercyjna sprzedaż prawa do posługiwania się logo bez audytu sponsora to bezwzględnie zakazana sprzedaż znaku zaufania.",
        ),
        (
            "Deklaracja certyfikacji: „Przeprowadzimy audyt ESG Państwa firmy i pomożemy uzyskać certyfikaty potwierdzające spełnienie najwyższych standardów” oraz podstrona «ESG Cert».",
            "ESG Institute jest spółką doradczą z o.o., a nie akredytowaną przez PCA jednostką certyfikującą (rozporządzenie 765/2008). Audyt doradczy kończy się raportem luki (gap analysis), a nie formalnym certyfikatem zgodności z prawem.",
        ),
        (
            "Oferta offsetów i partnerstwo z Las Na Zawsze: „propozycje offsetu” oraz „działania pozwalają firmom realizować cele z zakresu odpowiedzialności środowiskowej (ESG)”.",
            "ECGT pkt 4c całkowicie zakazuje claimów środowiskowych opartych na kompensacji i sadzeniu drzew. Standard ESRS E1 zabrania odejmowania offsetów od raportowanych emisji brutto w Scope 1, 2 i 3.",
        ),
        (
            "Rekomendacja Libra Partners: „produkuje ekologiczne palety z recyklingu, zmniejszając emisję CO2 (...) paleta antybakteryjna”.",
            "Hasło «ekologiczne palety» narusza pkt 4a ECGT. Paleta antybakteryjna zawiera biocydy podlegające rozporządzeniu BPR (brak rejestracji w Wykazie Produktów Biobójczych). Promowanie partnera to transfer autorytetu i greenwashing.",
        ),
        (
            "Green0meter: „Gwarancja przygotowania firmy do wymogów prawnych, chroniąc przed ryzykiem kar finansowych”.",
            "Żadne oprogramowanie SaaS/AI nie gwarantuje zgodności prawnej i nie zwalnia z odpowiedzialności. Zapewnienie o gwarancji braku kar narusza art. 5 UPNPR.",
        ),
        (
            "Model «ESG as a service»: „przejmując obowiązki raportowania od klienta i wykonując je w imieniu klienta”.",
            "Zgodnie z CSRD i art. 4a ustawy o rachunkowości, odpowiedzialność prawna za sprawozdawczość ESG jest niezbywalna i spoczywa osobiście na kierownictwie jednostki (zarządzie). Doradca nie może przejąć tego obowiązku.",
        ),
        (
            "Broszura DPP: „CZERWIEC 2027: Bezwzględny wymóg prawny UE (Digital Product Passport)” oraz moduł PCF dla odzieży roboczej „dla zgodności z CBAM”.",
            "ESPR wymaga uprzedniego uchwalenia aktów delegowanych przez KE (z 18-miesięcznym vacatio legis). CBAM dotyczy ściśle określonych surowców przemysłowych (stal, cement, aluminium etc.), a nie gotowej odzieży roboczej.",
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

    # 5. Ryzyka prawne i finansowe
    add_h("5. Ryzyka prawne i finansowe")
    risks = [
        (
            "1. Postępowanie Prezesa UOKiK (naruszenie zbiorowych interesów konsumentów i B2B)",
            "Stosowanie praktyk naruszających art. 5–7 UPNPR (w tym wprowadzające w błąd zapewnienia o gwarancjach ochrony przed karami, bezprawne oferowanie certyfikatów i sprzedaż prawa do logo) stanowi praktykę naruszającą zbiorowe interesy konsumentów i kontrahentów w rozumieniu art. 24 ustawy o ochronie konkurencji i konsumentów. Zgodnie z art. 106 ust. 1 pkt 4 ustawy, Prezes UOKiK może nałożyć na przedsiębiorcę karę pieniężną w wysokości do 10% obrotu osiągniętego w roku obrotowym poprzedzającym rok nałożenia kary. Kary mogą zostać nałożone także osobiście na osoby zarządzające (do 2 000 000 PLN).",
        ),
        (
            "2. Bezpośrednie roszczenia konkurentów z ustawy o zwalczaniu nieuczciwej konkurencji (UZNK)",
            "Przedstawianie oferty doradczej z obietnicą «gwarancji braku kar» oraz «przejęcia odpowiedzialności prawnej od klienta» stanowi czyn nieuczciwej konkurencji w rozumieniu art. 16 ust. 1 pkt 2 UZNK (reklama wprowadzająca klienta w błąd i mogąca przez to wpłynąć na jego decyzję co do nabycia usługi) oraz art. 3 UZNK (działanie sprzeczne z prawem i dobrymi obyczajami). Konkurencyjne firmy doradcze i kancelarie prawne mogą wystąpić z roszczeniami o zaniechanie niedozwolonych działań, usunięcie ich skutków, złożenie publicznego oświadczenia oraz zasądzenie odszkodowania lub wydanie bezpodstawnie uzyskanych korzyści (art. 18 UZNK).",
        ),
        (
            "3. Bezpośrednie stosowanie dyrektywy ECGT w kanale internetowym od 27 września 2026 r.",
            "Dyrektywa (UE) 2024/825 (ECGT) wiąże w obrocie unijnym od 27 września 2026 r. Witryny internetowe, materiały cyfrowe i oferty online nie korzystają z jakiejkolwiek taryfy ochronnej (stare zapasy dotyczą wyłącznie partii towarów wprowadzonych na rynek w opakowaniach fizycznych przed tą datą). Komunikat Komisji Europejskiej i sieci CPC jednoznacznie nakazuje natychmiastowe usunięcie niekwalifikowanych haseł («w pełni zrównoważony», «ekologiczne palety») oraz niecertyfikowanych oznakowań (sprzedaż logo w pakietach sponsorskich) z przestrzeni online.",
        ),
        (
            "4. Ryzyko reputacyjne i utrata mandatu do prowadzenia audytów compliance",
            "Dla podmiotu pozycjonującego się jako instytucja ekspercka, platforma wymiany wiedzy i audytor greenwashingu, ujawnienie wewnętrznych naruszeń regulacyjnych niesie katastrofalne skutki wizerunkowe (efekt „kto audytuje audytora”). Ryzyko to bezpośrednio uderza w zaufanie członków ESG Collaboration Hub, zniechęca partnerów instytucjonalnych (takich jak uczelnie czy fundacje) oraz podważa wiarygodność szkoleń realizowanych w ramach ESG Institute Academy.",
        ),
    ]
    for title, body in risks:
        add_p("• " + title, size=10.5, bold=True, before=8, after=2)
        add_p(body)

    # 6. Rekomendacje naprawcze
    add_h("6. Rekomendacje naprawcze")
    stages = [
        (
            "Etap 1: Natychmiastowy pull-off (do 15 października 2026 r.)",
            "1. Bezwzględnie usunąć zapis o «Prawie do używania logo ESG Institute na stronie internetowej Sponsora» ze wszystkich pakietów sponsorskich (Kobiety w ESG, ESG Ready).\n"
            "2. Wykreślić sformułowanie «Gwarancja przygotowania (...) chroniąc przed ryzykiem kar» z podstrony i broszury Green0meter.\n"
            "3. Usunąć określenie «ekologiczne palety z recyklingu» z opisu Libra Partners na podstronie Hubu i zawiesić promowanie kodu rabatowego do czasu wyjaśnienia kwestii rejestracji BPR.\n"
            "4. Wykreślić termin «propozycje offsetu» z opisu kompetencji dr Beaty Latos i oferty dekarbonizacyjnej.",
            "FFF1F2", "E11D48", DANGER,
        ),
        (
            "Etap 2: Korekta kart usługowych i terminologii prawnej (do 22 października 2026 r.)",
            "1. Zmienić opis modelu «ESG as a service»: zastąpić «przejęcie obowiązków raportowania» terminem «wsparcie operacyjne i analityczne w sporządzeniu projektu raportu».\n"
            "2. Przedefiniować podstronę «ESG Cert» oraz usunąć obietnicę «uzyskania certyfikatów» z broszury głównej – sprecyzować, że audyt jest audytem doradczym luki, a nie certyfikacją.\n"
            "3. Skorygować harmonogram CSRD na stronie: usunąć kryteria CSDD (>1000 pracowników, >450M EUR) i sprostować status notowanych MŚP.\n"
            "4. W broszurze DPP usunąć arbitralną datę «czerwiec 2027» oraz wykreślić mylne powiązanie modułu odzieżowego PCF z rozporządzeniem CBAM.",
            "FFFBEB", "D97706", WARNING,
        ),
        (
            "Etap 3: Wdrożenie procedury Safe Harbor i weryfikacji partnerów (do 15 listopada 2026 r.)",
            "1. Opracować i wdrożyć wewnętrzny kodeks Safe Harbor Claims dla marketingu ESG Institute – zakazujący używania haseł absolutnych («w pełni zrównoważony», «eko»).\n"
            "2. Wdrożyć procedurę weryfikacji greenwashingu dla podmiotów przystępujących do ESG Collaboration Hub – zakaz promowania partnerów stosujących niezweryfikowane deklaracje.\n"
            "3. Opublikować na witrynie zwięzły profil środowiskowy ESG Institute (emisje własne Scope 1–2, zużycie energii) w celu zachowania spójności z deklarowaną misją.",
            "F0F9FF", "0284C7", ACCENT,
        ),
        (
            "Etap 4: Stały nadzór compliance i transparentność instytucjonalna (tryb ciągły)",
            "1. Przeprowadzać kwartalny przegląd wszystkich publicznych treści marketingowych, ofert i prezentacji przez radcę prawnego / oficera compliance.\n"
            "2. Wprowadzić wyraźne rozgraniczenie między komercyjną działalnością doradczą (ESG Institute Sp. z o.o.) a działalnością społeczną think tanku i fundacji partnerskich.\n"
            "3. Regularne szkolenia zespołu marketingu i autorów artykułów z zakresu wytycznych ECGT, UOKiK i dyrektywy Omnibus.",
            "ECFDF5", "059669", RGBColor(4, 120, 87),
        ),
    ]
    for title, body, bg, border, color in stages:
        callout(title, body.split("\n"), bg, border, color)

    # 7. Wnioski
    add_h("7. Wnioski")
    add_p(
        "Komunikacja marketingowa i ofertowa ESG Institute Sp. z o.o. w obecnym kształcie narusza kluczowe zakazy dyrektywy (UE) 2024/825 (ECGT), "
        "w tym zakaz komercjalizacji niecertyfikowanych znaków zrównoważoności (sprzedaż praw do logo w sponsoringu), zakaz ogólnych claimów środowiskowych "
        "oraz zakaz promowania kompensacji emisji (offsetów). W obecnej formie materiały nie mogą być publikowane i wymagają natychmiastowego wdrożenia procedury naprawczej."
    )

    out = ROOT / "Raport_Audytu_Greenwashingu_ESG_Institute.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    print("Budowanie materiału źródłowego...")
    mat = build_material()
    print("Zapisano materiał:", mat)
    print("Budowanie raportu audytu...")
    rep = build_report()
    print("Zapisano raport:", rep)
