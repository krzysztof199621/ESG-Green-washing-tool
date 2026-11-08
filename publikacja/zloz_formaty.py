#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPTX 16:9 i pionowy PDF A4. Tekst zgodny z generuj_prezentacje.py."""

import math
import re
from pathlib import Path

import generuj_prezentacje as g
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
PPTX_OUT = HERE / "ESG_Institute_Audyt_ryzyk_greenwashingu.pptx"
PION_OUT = HERE / "ESG_Institute_Audyt_ryzyk_greenwashingu_pion.pdf"
LOGO = HERE / "_logo_jasne.png"

NAVY = RGBColor(0x14, 0x3B, 0x4E)
INK = RGBColor(0x33, 0x33, 0x33)
MUTED = RGBColor(0x66, 0x66, 0x66)
BLUE = RGBColor(0x33, 0x99, 0xCC)
BLUE_T = RGBColor(0x27, 0x76, 0x9D)
PALE = RGBColor(0xF5, 0xF7, 0xF8)
LINE = RGBColor(0xE3, 0xE6, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

ML = 0.46
SW = 13.333333
SH = 7.5
CW = SW - 2 * ML


def force_arial(run):
    run.font.name = "Arial"
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = OxmlElement("a:latin")
        rPr.append(latin)
    latin.set("typeface", "Arial")


def plain(markup):
    text = re.sub(r"<br\s*/?>", "\n", markup or "")
    text = re.sub(r"<[^>]+>", "", text)
    return text.replace("&nbsp;", " ")


def est_in(markup, width_in, size):
    text = plain(markup)
    cpl = max(8, (width_in * 72) / (size * 0.47))
    lines = 0
    for part in text.split("\n"):
        lines += max(1, math.ceil(len(part) / cpl)) if part else 1
    return lines * (size * 1.3) / 72


def add_markup(paragraph, markup, size, color, bold=False, italic=False):
    tokens = re.split(r"(</?b>|</?i>|<br\s*/?>)", markup or "")
    b, i = bold, italic
    for tok in tokens:
        if tok == "<b>":
            b = True
        elif tok == "</b>":
            b = bold
        elif tok == "<i>":
            i = True
        elif tok == "</i>":
            i = italic
        elif tok.startswith("<br"):
            paragraph.add_line_break()
        elif tok:
            run = paragraph.add_run()
            run.text = tok.replace("&nbsp;", " ")
            run.font.size = Pt(size)
            run.font.bold = b
            run.font.italic = i
            run.font.color.rgb = color
            force_arial(run)


def T(text, size=12, color=INK, bold=False, italic=False):
    return (text, size, color, bold, italic)


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(SW)
        self.prs.slide_height = Inches(SH)
        self.prs.core_properties.title = (
            "Audyt ryzyk greenwashingu. Metodyka, wnioski i rekomendacje"
        )
        self.prs.core_properties.author = "ESG Institute"
        self.prs.core_properties.subject = "Metodyka, wnioski i rekomendacje. 8 listopada 2026 r."
        self.blank = self.prs.slide_layouts[6]
        self.page = 0
        self.total = 13
        self.y = 0.0
        self.s = None
        self.section = ""

    def new(self, section):
        self.page += 1
        self.section = section
        self.s = self.prs.slides.add_slide(self.blank)
        self._chrome(section)
        self.y = 0.76
        return self

    def _shape(self, x, y, w, h, fill, line=True):
        sh = self.s.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(max(h, 0.02))
        )
        sh.shadow.inherit = False
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if line:
            sh.line.color.rgb = LINE
            sh.line.width = Pt(0.75)
        else:
            sh.line.fill.background()
        return sh

    def _chrome(self, section):
        if LOGO.exists():
            self.s.shapes.add_picture(
                str(LOGO), Inches(0.4), Inches(0.12), height=Inches(0.34)
            )
        box = self.s.shapes.add_textbox(Inches(8.2), Inches(0.16), Inches(4.65), Inches(0.28))
        tf = box.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        add_markup(p, section, 12, MUTED)
        self._shape(0, 0.54, SW, 0.012, LINE, line=False)

    def _footer(self):
        self._shape(ML, 7.2, CW, 0.008, LINE, line=False)
        box = self.s.shapes.add_textbox(Inches(ML), Inches(7.24), Inches(9.4), Inches(0.22))
        add_markup(box.text_frame.paragraphs[0], "ESG Institute  –  8 listopada 2026 r.", 10, MUTED)
        box2 = self.s.shapes.add_textbox(Inches(10.5), Inches(7.24), Inches(2.35), Inches(0.22))
        p2 = box2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.RIGHT
        add_markup(p2, f"{self.page}  /  {self.total}", 10, MUTED)

    def _text(self, x, y, w, h, items):
        box = self.s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        for margin in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, margin, Emu(0))
        for i, (markup, size, color, bold, italic) in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_before = Pt(0)
            p.space_after = Pt(2)
            p.line_spacing = 1.02
            add_markup(p, markup, size, color, bold, italic)
        return box

    def title(self, text, lead=None, kicker=None):
        if kicker:
            self._text(ML, self.y, CW, 0.22, [T(kicker, 11, BLUE_T, True)])
            self.y += 0.22
        h = est_in(text, CW, 24) + 0.04
        self._text(ML, self.y, CW, h, [T(text, 24, NAVY, True)])
        self.y += h + 0.05
        self._shape(ML, self.y, 0.36, 0.026, BLUE, line=False)
        self.y += 0.12
        if lead:
            h = est_in(lead, CW, 13) + 0.04
            self._text(ML, self.y, CW, h, [T(lead, 13, MUTED)])
            self.y += h + 0.1
        else:
            self.y += 0.02

    def cards(self, columns, fill=WHITE, gap=0.12, pad=0.1, after=0.08):
        n = len(columns)
        w = (CW - gap * (n - 1)) / n
        inner = w - 2 * pad
        heights = []
        for col in columns:
            h = pad
            for markup, size, color, bold, italic in col:
                h += est_in(markup, inner, size) + 0.06
            heights.append(h + pad * 0.35)
        height = max(heights) + 0.06
        if self.y + height > 7.16:
            print(f"PPTX wychodzi poza slajd: {self.section} y={self.y:.2f} h={height:.2f}")
        x = ML
        for col in columns:
            self._shape(x, self.y, w, height, fill, line=True)
            self._text(x + pad, self.y + pad * 0.55, inner, height - pad, col)
            x += w + gap
        self.y += height + after

    def close_slide(self):
        self._footer()


def build_pptx():
    d = Deck()
    d.page = 1
    d.section = "Okładka"
    d.s = d.prs.slides.add_slide(d.blank)
    if LOGO.exists():
        d.s.shapes.add_picture(str(LOGO), Inches(0.52), Inches(0.36), height=Inches(0.56))
    d._text(0.52, 1.15, 12, 0.62, [T("Audyt ryzyk greenwashingu", 34, NAVY, True)])
    d._shape(0.52, 1.84, 0.4, 0.028, BLUE, line=False)
    d._text(
        0.52,
        2.0,
        11.4,
        0.9,
        [T(
            "Jak ESG Institute ocenia twierdzenia środowiskowe i społeczne, zanim materiał zostanie opublikowany. Metoda, wnioski, przykładowe stwierdzenia i rekomendacje.",
            16,
            INK,
        )],
    )
    d._shape(0, 5.28, SW, 2.22, PALE, line=False)
    d._shape(0, 5.28, SW, 0.018, BLUE, line=False)
    x = 0.52
    for num, label in (
        ("01", "Metoda"),
        ("02", "Reżim"),
        ("03", "Werdykt"),
        ("04", "Przykłady"),
        ("05", "Rekomendacje"),
    ):
        d._text(x, 5.46, 2.3, 0.28, [T(f"{num}   {label}", 13, NAVY, True)])
        x += 2.5
    facts = [
        ("Bez oznaczonego podmiotu", "Zdania w przykładach ułożono na potrzeby tej publikacji. Nie są cytatami z materiałów konkretnej organizacji."),
        ("Dwa reżimy", "Do 26 września 2027 r. werdykt opiera się na przepisach już obowiązujących. Czarna lista EmpCo jest przyszłym standardem."),
        ("Krótka lista ustaleń", "Audyt zostawia twierdzenie tylko wtedy, gdy przesłanki są spełnione. Nie dąży do z góry przyjętej liczby uwag."),
    ]
    x = 0.52
    for title, body in facts:
        d._text(x, 5.92, 3.95, 1.35, [T(title, 13, NAVY, True), T(body, 12, INK)])
        x += 4.2
    d.close_slide()

    d.new("Zakres")
    d.title(
        "Zakres opracowania",
        "Publikacja przedstawia metodę audytu ryzyk greenwashingu stosowaną w ESG Institute. Nie zbiera wyników poszczególnych audytów i nie podaje liczby naruszeń na rynku.",
    )
    d.cards([
        [T("Co wchodzi do audytu", 14, NAVY, True), T("Przekaz przygotowany do publikacji: strona, oferta, specyfikacja, etykieta, prezentacja sprzedażowa albo fragment raportu. Ocenie podlega brzmienie i sposób ujęcia, łącznie z oprawą, która poszerza sens słów.")],
        [T("Co jest wynikiem", 14, NAVY, True), T("Werdykt o publikacji tego przekazu, ustalenia spełniające przesłanki oraz kierunek naprawy. Krótka lista jest wynikiem poprawnym. Audyt nie przyjmuje docelowej liczby uwag.")],
        [T("Czego materiał nie rozstrzyga", 14, NAVY, True), T("Nie ocenia całej działalności organizacji i nie zastępuje analizy konkretnego dokumentu. Stan prawa w tym opracowaniu: 8 listopada 2026 r. Tekst przepisu przy konkretnym audycie sprawdza się w ISAP i w EUR-Lex.")],
    ])
    d.cards([[
        T("Odbiorca opracowania", 14, NAVY, True),
        T("Działy komunikacji, zgodności i osoby zatwierdzające publikację. Przy konkretnym zleceniu audyt czyta wyłącznie przekazany dokument. Szuka wykładni dla podobnego stanu faktycznego, a nie nazwy przedsiębiorcy. Orzeczenie z innej branży wolno przywołać jako analogię, z zaznaczeniem, czego dotyczy i że nie rozstrzyga zgodności z prawem działalności badanego podmiotu."),
    ]])
    d.close_slide()

    d.new("Metoda")
    d.title(
        "Jak czyta się twierdzenie",
        "Cztery reguły metody. Tezy orzeczeń podano jako skrót wykładni przyjętej w standardzie audytu, nie jako cytat z sentencji.",
    )
    for num, title, body in (
        ("01", "Całość przekazu", "Hasło i późniejsze zastrzeżenie czyta się razem. Poprawna informacja niżej albo odesłanie na inną stronę nie zawsze zdejmuje pierwsze wrażenie. W metodzie ten kierunek wyznaczają wyroki TSUE: z 4 czerwca 2015 r., C-195/14, Teekanne, oraz – przy zaniechaniu – z 26 października 2016 r., C-611/14, Canal Digital Danmark."),
        ("02", "Decyzja dotycząca transakcji", "Przy art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym sprawdza się, czy przeciętny odbiorca może podjąć decyzję, której inaczej by nie podjął. Rzeczywisty zakup nie jest potrzebny. Sama nieścisłość albo brak liczby tego progu nie spełnia. Tak wyrok TSUE z 19 grudnia 2013 r., C-281/12, Trento Sviluppo."),
        ("03", "Zgodność zakresu z dowodem", "Informacja prawdziwa może mylić przez selektywny układ, pominięte założenie albo komentarz szerszy niż dowód. Cel nie jest wynikiem osiągniętym. Sąd Najwyższy w wyroku z 9 października 2014 r., IV CSK 56/14, oddalił skargę kasacyjną i nie uznał, że w tamtej sprawie samo użycie wyników badań było deliktem. Sprawdza się, czy zakres twierdzenia pokrywa się z zakresem dowodu."),
        ("04", "Przepis szczególny przed klauzulą ogólną", "Tam, gdzie temat ma własny akt, audyt zaczyna od niego. Przy oświadczeniach zdrowotnych i żywieniowych pierwszeństwo wobec dyrektywy 2005/29/WE potwierdza wyrok TSUE z 10 września 2020 r., C-363/19, Mezina. Finanse, transport, moda i budownictwo również czyta się najpierw przez regulację właściwą dla tematu."),
    ):
        d.cards([[T(f"{num}    {title}", 14, BLUE_T, True), T(body, 12)]], pad=0.08)
    d.close_slide()

    d.new("Prawo")
    d.title(
        "Reżim prawny na 8 listopada 2026 r.",
        "Dwie daty zostają rozdzielone. Werdykt do 26 września 2027 r. włącznie opiera się na przepisach już obowiązujących.",
    )
    d.cards([
        [T("Prawo, na którym stoi werdykt", 14, NAVY, True), T("– art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, z testem decyzji dotyczącej transakcji;<br/>– ustawa o ochronie konkurencji i konsumentów, gdy praktyka narusza zbiorowe interesy konsumentów; kara pieniężna Prezesa UOKiK może sięgnąć 10% obrotu;<br/>– ustawa o zwalczaniu nieuczciwej konkurencji, gdy przekaz jest reklamą porównawczą;<br/>– rozporządzenia sektorowe stosowane bezpośrednio, w tym przy produkcji ekologicznej i przy oświadczeniach zdrowotnych.", 12)],
        [T("Przyszły standard EmpCo", 14, BLUE_T, True), T("Dyrektywa (UE) 2024/825 z 28 lutego 2024 r. Art. 4 ust. 1 wyznacza 27 września 2026 r. jako datę, od której państwa członkowskie stosują krajowe przepisy transponujące. To data dyrektywy.<br/><br/>Ustawa z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) została ogłoszona 19 października 2026 r. Wchodzi w życie 27 września 2027 r.<br/><br/>Do tego dnia czarnej listy nie traktuje się jako już obowiązującego zakazu krajowego. Utrzymywanie starszej treści w internecie tej daty nie przyspiesza.", 12)],
    ])
    d.cards([[T("Wniosek Komisji COM(2023) 166 w sprawie oświadczeń środowiskowych nie jest prawem. W audycie może opisać dobrą praktykę dowodową, wyłącznie z tym zastrzeżeniem. Wytyczne OECD dla przedsiębiorstw wielonarodowych, edycja 2023, są zaleceniami. Na skali audytu dostają poziom niski. Postępowanie przed polskim Krajowym Punktem Kontaktowym kończy się publiczną oceną i zaleceniami. Nie jest sprawą sądową i nie kończy się wyrokiem.", 12)]], fill=PALE)
    d.close_slide()

    d.new("Werdykt")
    d.title(
        "Skala ryzyka i werdykt",
        "Skala mierzy naruszenie przepisu obowiązującego w dniu audytu. Samo zbliżenie do czarnej listy oraz samo powołanie Wytycznych OECD nie dają werdyktu wysokiego ani publikacji warunkowej.",
    )
    for i, (label, text) in enumerate((
        ("Bardzo wysokie.", "Przepis już obowiązujący, spełnione przesłanki, w tym test decyzji dotyczącej transakcji, oraz liczbowe wprowadzenie w błąd, dyskredytowanie konkurencji albo realne ryzyko kary lub powództwa."),
        ("Wysokie.", "Bezpośrednie naruszenie art. 5 lub art. 6 albo już obowiązującego zakazu sektorowego."),
        ("Średnie.", "Brak weryfikowalności albo porównanie, które spełnia przesłanki reklamy porównawczej lub zaniechania i da się naprawić przez doprecyzowanie."),
        ("Niskie.", "Drobna nieprecyzyjność poniżej progu naruszenia. Nie tworzy osobnego przypadku. Wytyczne OECD zostają na tym poziomie."),
        ("Przyszły standard.", "Do 26 września 2027 r. etykieta „Przyszły standard EmpCo (od 27.09.2027)”. Nie jest poziomem wysokim i nie jest stwierdzonym naruszeniem krajowej czarnej listy."),
    )):
        d.cards(
            [[T(f"<b>{label}</b>  {text}", 12, INK)]],
            fill=PALE if i % 2 == 0 else WHITE,
            pad=0.05,
            after=0.04,
        )
    d._text(ML, d.y, CW, 0.28, [T("Werdykt operacyjny do 26 września 2027 r.", 14, NAVY, True)])
    d.y += 0.32
    d.cards([
        [T("Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści", 12, NAVY, True), T("Co najmniej jedno ustalenie bardzo wysokie albo wysokie na prawie już obowiązującym. Sam przyszły standard tego werdyktu nie daje.", 11)],
        [T("Publikacja warunkowa – rekomendowana korekta przed publikacją", 12, NAVY, True), T("Brak takiego ustalenia, ale jest ryzyko średnie albo twierdzenie dziś mieszczące się w art. 5 i 6, które od 27 września 2027 r. spełni zakaz i wcześniej wymaga zawężenia.", 11)],
        [T("Dopuszczone do publikacji", 12, NAVY, True), T("Brak istotnych naruszeń prawa obowiązującego. Krótka wzmianka o przyszłym standardzie, która nie wymaga wstrzymania, zostaje we wniosku.", 11)],
    ], pad=0.08)
    d.close_slide()

    d.new("Tok audytu")
    d.title("Tok audytu", "Sześć kroków przy jednym dokumencie. Data audytu wskazuje reżim zapisany potem w metryczce.")
    steps = [
        ("1. Dokument i data", "Kanał, odbiorca, dosłowne brzmienie, oprawa i data. Data rozstrzyga, czy czarna lista jest już prawem polskim."),
        ("2. Wypisanie twierdzeń", "Zdania środowiskowe i społeczne w brzmieniu z materiału, z miejscem. Oprawa wchodzi do sensu, gdy poszerza słowa."),
        ("3. Test prawa obowiązującego", "Najpierw przepis sektorowy. Potem art. 5 i art. 6, w tym decyzja dotycząca transakcji."),
        ("4. Osobno standard z 2027 r.", "To samo zdanie może dostać drugą podstawę. Do 27 września 2027 r. nie jest naruszeniem krajowej czarnej listy."),
        ("5. Zostają spełnione przesłanki", "Twierdzenie poniżej progu nie dostaje numeru. Może wrócić przy przyszłym standardzie, jeżeli przed tą datą wymaga zawężenia."),
        ("6. Naprawa i wzór", "Kierunek korekty oraz zdanie, które da się opublikować. Gdy twierdzenia nie da się uratować, wzór odpada."),
    ]
    for row in range(3):
        pair = steps[row * 2 : row * 2 + 2]
        d.cards([[T(a, 14, BLUE_T, True), T(b, 12)] for a, b in pair], pad=0.1)
    d.close_slide()

    d.new("Wnioski")
    d.title("Wnioski z metody", "Siedem wniosków opisuje sposób oceny. Nie są ustaleniami wobec żadnej organizacji.")
    items = [
        "Przedmiotem werdyktu jest publikacja konkretnego przekazu. Audyt nie wystawia oceny całej organizacji.",
        "Ogólne hasło środowiskowe nie jest dziś zakazane z samego brzmienia. Krajowy zakaz takiego hasła, gdy brak uznanej wysokiej efektywności środowiskowej istotnej dla twierdzenia, zaczyna się 27 września 2027 r. Dla żywności terminy „eko”, „bio” i „organic” czyta rozporządzenie (UE) 2018/848.",
        "Neutralność, ograniczenie albo dodatni wpływ na emisje oparte na kompensacji czyta się do tej daty przez art. 5 i art. 6. Własny zakaz wchodzi 27 września 2027 r. Wcześniej nie jest bezwzględnym zakazem krajowym.",
        "Hasło o całym produkcie albo o całej działalności, oparte na jednym aspekcie, przechodzi ten sam test wprowadzenia w błąd. Od 27 września 2027 r. dochodzi przyszły standard o takim rozjechaniu zakresu. Dowód z innej kategorii produktów nie jest przesłanką tego punktu. Tę niewspółmierność czyta się osobno na art. 5 i art. 6.",
        "Porównanie zostaje, gdy jest obiektywne i sprawdzalne. Brak nazwy konkurenta nie wyklucza reklamy porównawczej. Samo zestawienie kategorii jej nie przesądza. Naprawą jest ujawnienie założeń, które rozstrzygają wynik.",
        "Poinformowanie o zgodności z limitem nałożonym na producenta nie jest dziś naruszeniem. Od 27 września 2027 r. przyszły standard obejmuje przedstawianie jako przewagi wymogu, który dotyczy wszystkich produktów danej kategorii na rynku unijnym.",
        "Hasło o ludziach, pracy albo etyce przechodzi test wprowadzenia w błąd. Samo powołanie Wytycznych OECD zostaje przy ryzyku niskim: publiczna ocena Krajowego Punktu Kontaktowego i skutek wizerunkowy, bez wyroku i bez podniesienia werdyktu.",
    ]
    for i, text in enumerate(items, 1):
        d.cards([[T(f"{i}.  {text}", 12, INK)]], fill=PALE if i % 2 else WHITE, pad=0.06)
    d.close_slide()

    def example(section, kicker, title, context, quotes, status, assessment, repairs):
        d.new(section)
        d.title(title, context, kicker)
        d.cards([[T(q, size, NAVY if size > 13 else MUTED, False, True) for q, size in quotes]], fill=PALE, pad=0.1)
        d.cards([[T(status, 12)]], pad=0.08)
        d.cards([
            [T("Ocena", 14, NAVY, True), T(assessment, 12)],
            [T("Dozwolone sformułowanie", 14, NAVY, True), *repairs],
        ])
        d.close_slide()

    example(
        "Przykład 1 z 4",
        "Przykład konstrukcyjny",
        "Neutralność oparta na kompensacji",
        None,
        [("„Jesteśmy firmą neutralną klimatycznie.”", 16), ("„Neutralność wynika z kompensacji emisji.”", 12)],
        "<b>Ilustracja poziomu: wysokie.</b> Tylko wtedy, gdy spełniony jest test decyzji dotyczącej transakcji. Samo zbliżenie do zakazu z 27 września 2027 r., bez dzisiejszego wprowadzenia w błąd, tego poziomu nie daje. Może prowadzić do publikacji warunkowej, jeżeli przed tą datą twierdzenie wymaga zawężenia.",
        "Odbiorca może odczytać, że działalność nie zostawia emisji, a kompensacja stoi w dopisku. Do 26 września 2027 r. podstawą jest art. 5 albo art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, jeżeli przeciętny odbiorca może na tej podstawie podjąć decyzję dotyczącą transakcji, której inaczej by nie podjął. Od 27 września 2027 r. twierdzenie o neutralnym wpływie na emisje oparte na kompensacji wchodzi w krajowy zakaz. Wcześniej audyt opisuje je jako przyszły standard.",
        [T("„W [roku] emisje zakresu 1 i 2 były o [X]% niższe niż w [roku bazowym]. Emisji pozostałych w tym zdaniu nie kompensujemy. Rachunek i granice są pod [adresem].”", 12, INK, False, True), T("Zdania „nie emitujemy gazów cieplarnianych”, gdy emisje występują, nie da się uratować dopiskiem. Wychodzą z materiału.", 11)],
    )
    example(
        "Przykład 2 z 4",
        "Przykład konstrukcyjny",
        "Hasło o całości, dowód o fragmencie",
        "W przykładzie jedyny opublikowany rachunek dotyczy jednego zakładu, jednego roku oraz zakresów 1 i 2. Hasło obejmuje całą ofertę.",
        [("„Cała oferta jest zeroemisyjna.”", 16)],
        "<b>Ilustracja poziomu: wysokie.</b> Tylko wtedy, gdy odbiorca może podjąć decyzję dotyczącą transakcji w przekonaniu, że wynik dotyczy wszystkiego, co kupuje. Pochodzenie dowodu z innej kategorii produktów ocenia się osobno i nie dopisuje się go do przyszłego punktu o rozjechaniu zakresu.",
        "Do 26 września 2027 r. jest to pytanie o art. 5 albo art. 6. Od 27 września 2027 r. dochodzi przyszły standard: twierdzenie środowiskowe o całym produkcie albo o całej działalności, które w istocie dotyczy tylko określonego aspektu albo konkretnego rodzaju działalności. Dowód z innej kategorii produktów nie jest przesłanką tego punktu. Niewspółmierność takiej kategorii czyta się osobno, jeżeli spełniony jest test decyzji dotyczącej transakcji.",
        [T("„Zakład w [miejscowości] obniżył w [roku] emisje zakresu 1 i 2 o [X]% względem [roku bazowego]. Wynik nie obejmuje pozostałych zakładów ani zakresu 3.”", 12, INK, False, True), T("Nawiasy uzupełnia nadawca danymi, które już ma. Hasło o całej ofercie bez takiego rachunku nie wraca do materiału w poprzednim brzmieniu.", 11)],
    )
    example(
        "Przykład 3 z 4",
        "Przykład konstrukcyjny",
        "Porównanie w kategorii",
        "W przykładzie przy haśle nie ma metody, roku, jednostki, granic analizy ani listy produktów objętych zestawieniem.",
        [("„Nasz ślad węglowy jest najniższy w kategorii.”", 16)],
        "<b>Ilustracja poziomu: średnie.</b> Gdy porównanie spełnia przesłanki i da się je naprawić doprecyzowaniem. Poziom rośnie, gdy liczba jest myląca albo dyskredytuje konkurencję. Kierunkiem naprawy jest uzupełnienie założeń, nie usunięcie zestawienia z samego faktu, że kategorie zestawiono.",
        "Brak nazwy konkurenta nie zamyka pytania o reklamę porównawczą. Potrzebna jest możliwość zidentyfikowania konkurenta albo jego towarów oraz to, że towary są konkurencyjne. Tak wyrok TSUE z 19 kwietnia 2007 r., C-381/05, De Landtsheer. Zestawienie może zostać, jeżeli jest obiektywne i sprawdzalne. Tak wyrok TSUE z 19 września 2006 r., C-356/04, Lidl Belgium. Założenia, które rozstrzygają wynik, mają być czytelne w samym przekazie. Z wyroku TSUE z 8 lutego 2017 r., C-562/15, Carrefour Hypermarchés, nie wynika ogólny obowiązek powołania określonej normy ISO w każdej reklamie. Liczba bez istotnego założenia może być zaniechaniem z art. 6, gdy wpływa na decyzję dotyczącą transakcji.",
        [T("„Według [metody], za [rok], w granicach od [etapu] do [etapu], jeden [jednostka] ma ślad [X] kg CO2e. Zestawienie obejmuje [liczbę] produktów sprzedawanych w [kanale]. Założenia są pod [adresem].”", 12, INK, False, True), T("Korzystne, rzetelne porównanie zostaje. Nawiasy uzupełnia nadawca.", 11)],
    )

    d.new("Przykład 4 z 4")
    d.title("Znak, obowiązek prawny i etyka", "Trzy kolejne zdania ułożone na potrzeby publikacji. Każde idzie inną ścieżką skali.", "Przykłady konstrukcyjne")
    for label, color, quote, body, repair in (
        ("Przyszły standard znaku", BLUE_T, "„Certyfikat zrównoważonego rozwoju” jako plakietka ułożona przez samego nadawcę, bez systemu certyfikacji i bez ustanowienia przez organ publiczny.", "Od 27 września 2027 r. taki znak wchodzi w przyszły standard. Dziś pyta się, czy odbiorca odbiera plakietkę jako potwierdzenie osoby trzeciej albo organu i czy może na tej podstawie podjąć decyzję dotyczącą transakcji. Brak numeru certyfikatu w tym samym akapicie sam tego nie dowodzi. Przy żywności terminy „eko”, „bio” i „organic” czyta rozporządzenie (UE) 2018/848, nie ogólny zakaz haseł środowiskowych.", "Po korekcie: „Parametr [nazwa] dla [produktu] w [roku] wynosi [X], według [metody].” Plakietki bez systemu certyfikacji albo organu publicznego nie zostawia się."),
        ("Przyszły standard przewagi", BLUE_T, "„Jako jedyni spełniamy unijny limit dla tej kategorii.”", "Samo poinformowanie o zgodności z limitem nie wystarcza do naruszenia. Od 27 września 2027 r. przyszły standard obejmuje przedstawianie jako cechy wyróżniającej wymogu nałożonego prawem na wszystkie produkty danej kategorii na rynku unijnym. Trzeba ustalić obie rzeczy: czy obowiązek naprawdę dotyczy wszystkich produktów właściwej kategorii oraz czy przekaz sprzedaje go jako przewagę.", "Po korekcie: „Produkty tej kategorii podlegają limitowi [nazwa]. Spełniamy go.” Limitu wspólnego dla całej kategorii nie przedstawia się jako cechy wyróżniającej."),
        ("Poziom niski", NAVY, "„Wdrożyliśmy Wytyczne OECD w całym łańcuchu dostaw.”", "To zalecenie, nie wyrok. Na skali audytu zostaje przy ryzyku niskim, o ile to samo zdanie nie spełnia równocześnie art. 5 albo art. 6. Wtedy przepis obowiązujący ocenia się na jego szczeblu, a ścieżkę Krajowego Punktu Kontaktowego dopisuje się osobno. Skutek tej ścieżki: publiczna ocena, zalecenia i ryzyko wizerunkowe.", "Po korekcie: „W [roku] opisaliśmy działania wobec [grupy dostawców] w [dokumencie].” Zdanie nie poświadcza zgodności z Wytycznymi OECD, jeżeli takiego sprawdzenia nie było."),
    ):
        d.cards([[T(label, 13, color, True), T(quote, 12, MUTED, False, True), T(body, 11), T(repair, 11, INK, False, True)]], fill=PALE, pad=0.07)
    d.close_slide()

    d.new("Rekomendacje")
    d.title(
        "Rekomendacje w czterech etapach",
        "Terminy liczy się od daty konkretnego audytu. W tym opracowaniu takiej daty nie ma, bo nie ma oznaczonego materiału. Kolejność etapów zostaje.",
    )
    stages = [
        ("1. Od daty audytu", "Etap natychmiastowy", "Wstrzymanie obejmuje wyłącznie twierdzenia, które już spełniają przesłanki prawa obowiązującego. Hasła objęte tylko standardem z 27 września 2027 r. do tego etapu nie wchodzą."),
        ("2. Przed 27 września 2027 r.", "Zawężenie przyszłego standardu", "Korekta haseł, które od tej daty spełnią krajowy zakaz: ogólne twierdzenie bez uznanej efektywności istotnej dla niego, własny znak bez systemu certyfikacji, kompensacja jako podstawa neutralności, całość oparta na jednym aspekcie, obowiązek prawny przedstawiony jako przewaga."),
        ("3. Przy korekcie materiałów", "Jedno brzmienie we wszystkich kanałach", "Ta sama wersja twierdzenia na stronie, w ofercie, w specyfikacji i w raporcie. Oprawa graficzna nie poszerza zdania. Założenia porównania stoją przy liczbie. Wzór po korekcie jest zdaniem, które da się obronić dowodem już posiadanym."),
        ("4. Na stałe", "Metodyka, słownik i szkolenie", "Właściciel dowodu, data, granice pomiaru i słownik sformułowań dopuszczalnych. Przegląd przed publikacją obejmuje osoby, które piszą oferty i odpowiedzi na ankiety. Twierdzenie o ludziach i o etyce przechodzi tę samą ścieżkę."),
    ]
    for row in range(2):
        pair = stages[row * 2 : row * 2 + 2]
        d.cards([[T(a, 13, BLUE_T, True), T(b, 14, NAVY, True), T(c, 12)] for a, b, c in pair])
    d.close_slide()

    d.new("Nota")
    d.title("Granice metody", "Pięć granic, które audyt stosuje przy każdym materiale. Krótka lista ustaleń jest wtedy wynikiem poprawnym.")
    for i, text in enumerate((
        "Przykładu z jednego audytu nie przenosi się na inny podmiot. Test prawny zostaje, przypisanie przypadku odpada.",
        "Przyszły standard EmpCo, do 26 września 2027 r., opisuje się osobno od naruszenia prawa już obowiązującego.",
        "Same Wytyczne OECD zostają przy ryzyku niskim i same nie podnoszą werdyktu.",
        "Porównania nie usuwa się wyłącznie dlatego, że zestawiono kategorie. Uzupełnia się założenia albo zostawia zestawienie, gdy jest rzetelne.",
        "Braku numeru certyfikatu w jednym akapicie nie traktuje się jako dowodu naruszenia.",
    ), 1):
        d.cards([[T(f"{i}.  {text}", 13)]], fill=PALE if i % 2 else WHITE, pad=0.06)
    d.cards([[
        T("ESG Institute", 14, NAVY, True),
        T("info@esginstitute.eu    –    esginstitute.eu    –    Sygn. ESG/2026/11/08-P", 12, INK),
        T("Teksty ustaw sprawdza się w ISAP, prawo Unii w EUR-Lex, orzecznictwo Trybunału Sprawiedliwości Unii Europejskiej w EUR-Lex, orzecznictwo polskie w SAOS i w bazie Sądu Najwyższego, a w sprawach administracyjnych także w bazie NSA. Opracowanie streszcza metodę według stanu na 8 listopada 2026 r. Audyt konkretnego materiału powstaje na podstawie przekazanego dokumentu. To opracowanie go nie zastępuje.", 12),
    ]], fill=PALE)
    d.close_slide()
    d.prs.save(str(PPTX_OUT))
    print("Zapisano", PPTX_OUT)


def tune_portrait():
    g.register_fonts()
    g.ST = g.styles()
    g.ST["h1"].fontSize = 18
    g.ST["h1"].leading = 22
    g.ST["body"].fontSize = 10.5
    g.ST["body"].leading = 14
    g.ST["body_s"].fontSize = 10
    g.ST["body_s"].leading = 13.2
    g.ST["lead"].fontSize = 10.5
    g.ST["lead"].leading = 14
    g.ST["quote"].fontSize = 12.5
    g.ST["quote"].leading = 16.5
    g.ST["quote_s"].fontSize = 10.5
    g.ST["quote_s"].leading = 14
    g.ST["repair"].fontSize = 10.5
    g.ST["repair"].leading = 14
    g.ST["h2"].fontSize = 12
    g.ST["h2"].leading = 15.5
    g.ST["h2_blue"].fontSize = 12
    g.ST["h2_blue"].leading = 15.5
    g.LOGO, g.LOGO_RATIO = g.prepare_logo()
    g.W, g.H = A4
    g.ML = 36
    g.CW = g.W - 2 * g.ML


def stack(c, y, blocks, fill=None, pad=9, gap=7):
    fill = g.WHITE if fill is None else fill
    h = g.panels_row(c, g.ML, y, g.CW, 0, [{"fill": fill, "blocks": blocks}], pad=pad, inner_gap=3)
    y -= h + gap
    if y < 48:
        print(f"PION wychodzi w stopkę, y={y:.0f}")
    return y


def p_cover(c, page, total):
    c.bookmarkPage("p1")
    c.addOutlineEntry("Okładka", "p1", level=0, closed=False)
    c.setFillColor(g.WHITE)
    c.rect(0, 0, g.W, g.H, fill=1, stroke=0)
    g.draw_logo(c, 40, g.H - 78, 36)
    y = g.H - 118
    p, h = g.mp("Audyt ryzyk greenwashingu", g.ST["cover_h"], g.CW)
    g.ST["cover_h"].fontSize = 26
    g.ST["cover_h"].leading = 31
    p, h = g.mp("Audyt ryzyk greenwashingu", g.ST["cover_h"], g.W - 80)
    p.drawOn(c, 40, y - h)
    y -= h + 12
    c.setFillColor(g.BLUE)
    c.rect(40, y, 28, 2, fill=1, stroke=0)
    y -= 16
    g.ST["cover_sub"].fontSize = 12
    g.ST["cover_sub"].leading = 16.5
    p, h = g.mp(
        "Jak ESG Institute ocenia twierdzenia środowiskowe i społeczne, zanim materiał zostanie opublikowany. Metoda, wnioski, przykładowe stwierdzenia i rekomendacje.",
        g.ST["cover_sub"],
        g.W - 80,
    )
    p.drawOn(c, 40, y - h)
    band_h = 250
    c.setFillColor(g.PALE)
    c.rect(0, 0, g.W, band_h, fill=1, stroke=0)
    c.setFillColor(g.BLUE)
    c.rect(0, band_h, g.W, 1.5, fill=1, stroke=0)
    ix, iy = 40, band_h + 22
    for num, label in (("01", "Metoda"), ("02", "Reżim"), ("03", "Werdykt"), ("04", "Przykłady"), ("05", "Rekomendacje")):
        c.setFillColor(g.BLUE_TEXT)
        c.setFont("Arial-Bold", 10)
        c.drawString(ix, iy, num)
        c.setFillColor(g.NAVY)
        c.setFont("Arial", 10)
        c.drawString(ix + 20, iy, label)
        ix += 108
        if num == "03":
            ix, iy = 40, iy - 16
    facts = [
        ("Bez oznaczonego podmiotu", "Zdania w przykładach ułożono na potrzeby tej publikacji. Nie są cytatami z materiałów konkretnej organizacji."),
        ("Dwa reżimy", "Do 26 września 2027 r. werdykt opiera się na przepisach już obowiązujących. Czarna lista EmpCo jest przyszłym standardem."),
        ("Krótka lista ustaleń", "Audyt zostawia twierdzenie tylko wtedy, gdy przesłanki są spełnione. Nie dąży do z góry przyjętej liczby uwag."),
    ]
    yy = band_h - 18
    for title, body in facts:
        p, h = g.mp(title, g.ST["h2"], g.W - 80)
        p.drawOn(c, 40, yy - h)
        yy -= h + 2
        p, h = g.mp(body, g.ST["body_s"], g.W - 80)
        p.drawOn(c, 40, yy - h)
        yy -= h + 10
    g.finish("pion okładka", yy)


def p_scope(c, page, total):
    y = g.chrome(c, "Zakres", page, total, "Zakres opracowania")
    y = g.section_title(c, "Zakres opracowania", "Publikacja przedstawia metodę audytu ryzyk greenwashingu stosowaną w ESG Institute. Nie zbiera wyników poszczególnych audytów i nie podaje liczby naruszeń na rynku.", y)
    for title, body in (
        ("Co wchodzi do audytu", "Przekaz przygotowany do publikacji: strona, oferta, specyfikacja, etykieta, prezentacja sprzedażowa albo fragment raportu. Ocenie podlega brzmienie i sposób ujęcia, łącznie z oprawą, która poszerza sens słów."),
        ("Co jest wynikiem", "Werdykt o publikacji tego przekazu, ustalenia spełniające przesłanki oraz kierunek naprawy. Krótka lista jest wynikiem poprawnym. Audyt nie przyjmuje docelowej liczby uwag."),
        ("Czego materiał nie rozstrzyga", "Nie ocenia całej działalności organizacji i nie zastępuje analizy konkretnego dokumentu. Stan prawa w tym opracowaniu: 8 listopada 2026 r. Tekst przepisu przy konkretnym audycie sprawdza się w ISAP i w EUR-Lex."),
        ("Odbiorca opracowania", "Działy komunikacji, zgodności i osoby zatwierdzające publikację. Przy konkretnym zleceniu audyt czyta wyłącznie przekazany dokument. Szuka wykładni dla podobnego stanu faktycznego, a nie nazwy przedsiębiorcy. Orzeczenie z innej branży wolno przywołać jako analogię, z zaznaczeniem, czego dotyczy i że nie rozstrzyga zgodności z prawem działalności badanego podmiotu."),
    ):
        y = stack(c, y, [(title, g.ST["h2"]), (body, g.ST["body"])])
    g.finish("pion zakres", y)


def p_principles(c, page, total):
    y = g.chrome(c, "Metoda", page, total, "Jak czyta się twierdzenie")
    y = g.section_title(c, "Jak czyta się twierdzenie", "Cztery reguły metody. Tezy orzeczeń podano jako skrót wykładni przyjętej w standardzie audytu, nie jako cytat z sentencji.", y)
    rows = (
        ("01  Całość przekazu", "Hasło i późniejsze zastrzeżenie czyta się razem. Poprawna informacja niżej albo odesłanie na inną stronę nie zawsze zdejmuje pierwsze wrażenie. W metodzie ten kierunek wyznaczają wyroki TSUE: z 4 czerwca 2015 r., C-195/14, Teekanne, oraz – przy zaniechaniu – z 26 października 2016 r., C-611/14, Canal Digital Danmark."),
        ("02  Decyzja dotycząca transakcji", "Przy art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym sprawdza się, czy przeciętny odbiorca może podjąć decyzję, której inaczej by nie podjął. Rzeczywisty zakup nie jest potrzebny. Sama nieścisłość albo brak liczby tego progu nie spełnia. Tak wyrok TSUE z 19 grudnia 2013 r., C-281/12, Trento Sviluppo."),
        ("03  Zgodność zakresu z dowodem", "Informacja prawdziwa może mylić przez selektywny układ, pominięte założenie albo komentarz szerszy niż dowód. Cel nie jest wynikiem osiągniętym. Sąd Najwyższy w wyroku z 9 października 2014 r., IV CSK 56/14, oddalił skargę kasacyjną i nie uznał, że w tamtej sprawie samo użycie wyników badań było deliktem. Sprawdza się, czy zakres twierdzenia pokrywa się z zakresem dowodu."),
        ("04  Przepis szczególny przed klauzulą ogólną", "Tam, gdzie temat ma własny akt, audyt zaczyna od niego. Przy oświadczeniach zdrowotnych i żywieniowych pierwszeństwo wobec dyrektywy 2005/29/WE potwierdza wyrok TSUE z 10 września 2020 r., C-363/19, Mezina. Finanse, transport, moda i budownictwo również czyta się najpierw przez regulację właściwą dla tematu."),
    )
    for title, body in rows:
        y = stack(c, y, [(title, g.ST["h2_blue"]), (body, g.ST["body"])], pad=8)
    g.finish("pion zasady", y)


def p_regime(c, page, total):
    y = g.chrome(c, "Prawo", page, total, "Reżim prawny")
    y = g.section_title(c, "Reżim prawny na 8 listopada 2026 r.", "Dwie daty zostają rozdzielone. Werdykt do 26 września 2027 r. włącznie opiera się na przepisach już obowiązujących.", y)
    y = stack(c, y, [
        ("Prawo, na którym stoi werdykt", g.ST["h2"]),
        ("– art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, z testem decyzji dotyczącej transakcji;<br/>– ustawa o ochronie konkurencji i konsumentów, gdy praktyka narusza zbiorowe interesy konsumentów; kara pieniężna Prezesa UOKiK może sięgnąć 10% obrotu;<br/>– ustawa o zwalczaniu nieuczciwej konkurencji, gdy przekaz jest reklamą porównawczą;<br/>– rozporządzenia sektorowe stosowane bezpośrednio, w tym przy produkcji ekologicznej i przy oświadczeniach zdrowotnych.", g.ST["body"]),
    ])
    y = stack(c, y, [
        ("Przyszły standard EmpCo", g.ST["h2_blue"]),
        ("Dyrektywa (UE) 2024/825 z 28 lutego 2024 r. Art. 4 ust. 1 wyznacza 27 września 2026 r. jako datę, od której państwa członkowskie stosują krajowe przepisy transponujące. To data dyrektywy.<br/><br/>Ustawa z 4 września 2026 r. o zmianie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawy o prawach konsumenta (Dz.U. z 2026 r. poz. 1353) została ogłoszona 19 października 2026 r. Wchodzi w życie 27 września 2027 r.<br/><br/>Do tego dnia czarnej listy nie traktuje się jako już obowiązującego zakazu krajowego. Utrzymywanie starszej treści w internecie tej daty nie przyspiesza.", g.ST["body"]),
    ], fill=g.PALE)
    y = stack(c, y, [("Wniosek Komisji COM(2023) 166 w sprawie oświadczeń środowiskowych nie jest prawem. W audycie może opisać dobrą praktykę dowodową, wyłącznie z tym zastrzeżeniem. Wytyczne OECD dla przedsiębiorstw wielonarodowych, edycja 2023, są zaleceniami. Na skali audytu dostają poziom niski. Postępowanie przed polskim Krajowym Punktem Kontaktowym kończy się publiczną oceną i zaleceniami. Nie jest sprawą sądową i nie kończy się wyrokiem.", g.ST["body_s"])], fill=g.PALE)
    g.finish("pion reżim", y)


def p_verdict(c, page, total):
    y = g.chrome(c, "Werdykt", page, total, "Skala ryzyka i werdykt")
    y = g.section_title(c, "Skala ryzyka i werdykt", "Skala mierzy naruszenie przepisu obowiązującego w dniu audytu. Samo zbliżenie do czarnej listy oraz samo powołanie Wytycznych OECD nie dają werdyktu wysokiego ani publikacji warunkowej.", y)
    rows = (
        ("Bardzo wysokie", "Przepis już obowiązujący, spełnione przesłanki, w tym test decyzji dotyczącej transakcji, oraz liczbowe wprowadzenie w błąd, dyskredytowanie konkurencji albo realne ryzyko kary lub powództwa."),
        ("Wysokie", "Bezpośrednie naruszenie art. 5 lub art. 6 albo już obowiązującego zakazu sektorowego."),
        ("Średnie", "Brak weryfikowalności albo porównanie, które spełnia przesłanki reklamy porównawczej lub zaniechania i da się naprawić przez doprecyzowanie."),
        ("Niskie", "Drobna nieprecyzyjność poniżej progu naruszenia. Nie tworzy osobnego przypadku. Wytyczne OECD zostają na tym poziomie."),
        ("Przyszły standard EmpCo (od 27.09.2027)", "Do 26 września 2027 r. osobna etykieta. Nie jest poziomem wysokim i nie jest stwierdzonym naruszeniem krajowej czarnej listy."),
    )
    for i, (label, text) in enumerate(rows):
        y = stack(c, y, [(label, g.ST["h2"]), (text, g.ST["body_s"])], fill=g.PALE if i % 2 == 0 else g.WHITE, pad=7, gap=5)
    y = stack(c, y, [("Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści", g.ST["h2"]), ("Co najmniej jedno ustalenie bardzo wysokie albo wysokie na prawie już obowiązującym. Sam przyszły standard tego werdyktu nie daje.", g.ST["body_s"])], pad=8)
    y = stack(c, y, [("Publikacja warunkowa – rekomendowana korekta przed publikacją", g.ST["h2"]), ("Brak takiego ustalenia, ale jest ryzyko średnie albo twierdzenie dziś mieszczące się w art. 5 i 6, które od 27 września 2027 r. spełni zakaz i wcześniej wymaga zawężenia.", g.ST["body_s"])], pad=8)
    y = stack(c, y, [("Dopuszczone do publikacji", g.ST["h2"]), ("Brak istotnych naruszeń prawa obowiązującego. Krótka wzmianka o przyszłym standardzie, która nie wymaga wstrzymania, zostaje we wniosku.", g.ST["body_s"])], pad=8)
    g.finish("pion werdykt", y)


def p_steps(c, page, total):
    y = g.chrome(c, "Tok audytu", page, total, "Tok audytu")
    y = g.section_title(c, "Tok audytu", "Sześć kroków przy jednym dokumencie. Data audytu wskazuje reżim zapisany potem w metryczce.", y)
    for title, body in (
        ("1. Dokument i data", "Kanał, odbiorca, dosłowne brzmienie, oprawa i data. Data rozstrzyga, czy czarna lista jest już prawem polskim."),
        ("2. Wypisanie twierdzeń", "Zdania środowiskowe i społeczne w brzmieniu z materiału, z miejscem. Oprawa wchodzi do sensu, gdy poszerza słowa."),
        ("3. Test prawa obowiązującego", "Najpierw przepis sektorowy. Potem art. 5 i art. 6, w tym decyzja dotycząca transakcji."),
        ("4. Osobno standard z 2027 r.", "To samo zdanie może dostać drugą podstawę. Do 27 września 2027 r. nie jest naruszeniem krajowej czarnej listy."),
        ("5. Zostają spełnione przesłanki", "Twierdzenie poniżej progu nie dostaje numeru. Może wrócić przy przyszłym standardzie, jeżeli przed tą datą wymaga zawężenia."),
        ("6. Naprawa i wzór", "Kierunek korekty oraz zdanie, które da się opublikować. Gdy twierdzenia nie da się uratować, wzór odpada."),
    ):
        y = stack(c, y, [(title, g.ST["h2_blue"]), (body, g.ST["body"])], pad=8)
    g.finish("pion tok", y)


def p_conclusions(c, page, total):
    y = g.chrome(c, "Wnioski", page, total, "Wnioski z metody")
    y = g.section_title(c, "Wnioski z metody", "Siedem wniosków opisuje sposób oceny. Nie są ustaleniami wobec żadnej organizacji.", y)
    items = [
        "Przedmiotem werdyktu jest publikacja konkretnego przekazu. Audyt nie wystawia oceny całej organizacji.",
        "Ogólne hasło środowiskowe nie jest dziś zakazane z samego brzmienia. Krajowy zakaz takiego hasła, gdy brak uznanej wysokiej efektywności środowiskowej istotnej dla twierdzenia, zaczyna się 27 września 2027 r. Dla żywności terminy „eko”, „bio” i „organic” czyta rozporządzenie (UE) 2018/848.",
        "Neutralność, ograniczenie albo dodatni wpływ na emisje oparte na kompensacji czyta się do tej daty przez art. 5 i art. 6. Własny zakaz wchodzi 27 września 2027 r. Wcześniej nie jest bezwzględnym zakazem krajowym.",
        "Hasło o całym produkcie albo o całej działalności, oparte na jednym aspekcie, przechodzi ten sam test wprowadzenia w błąd. Od 27 września 2027 r. dochodzi przyszły standard o takim rozjechaniu zakresu. Dowód z innej kategorii produktów nie jest przesłanką tego punktu. Tę niewspółmierność czyta się osobno na art. 5 i art. 6.",
        "Porównanie zostaje, gdy jest obiektywne i sprawdzalne. Brak nazwy konkurenta nie wyklucza reklamy porównawczej. Samo zestawienie kategorii jej nie przesądza. Naprawą jest ujawnienie założeń, które rozstrzygają wynik.",
        "Poinformowanie o zgodności z limitem nałożonym na producenta nie jest dziś naruszeniem. Od 27 września 2027 r. przyszły standard obejmuje przedstawianie jako przewagi wymogu, który dotyczy wszystkich produktów danej kategorii na rynku unijnym.",
        "Hasło o ludziach, pracy albo etyce przechodzi test wprowadzenia w błąd. Samo powołanie Wytycznych OECD zostaje przy ryzyku niskim: publiczna ocena Krajowego Punktu Kontaktowego i skutek wizerunkowy, bez wyroku i bez podniesienia werdyktu.",
    ]
    for i, text in enumerate(items, 1):
        y = stack(c, y, [(f"<font color='#27769D'><b>{i}.</b></font>  {text}", g.ST["body"])], fill=g.PALE if i % 2 else g.WHITE, pad=7, gap=4)
    g.finish("pion wnioski", y)


def p_example(c, page, total, section, outline, title, context, quotes, status, assessment, repairs):
    y = g.chrome(c, section, page, total, outline)
    y = g.section_title(c, title, context, y, kicker="Przykład konstrukcyjny")
    y = stack(c, y, quotes, fill=g.PALE)
    y = stack(c, y, [(status, g.ST["body_s"])], pad=8)
    y = stack(c, y, [("Ocena", g.ST["h2"]), (assessment, g.ST["body"])])
    y = stack(c, y, [("Dozwolone sformułowanie", g.ST["h2"]), *repairs], fill=g.PALE)
    g.finish("pion " + title, y)


def p_ex1(c, page, total):
    p_example(
        c, page, total, "Przykład 1 z 4", "Przykład: kompensacja",
        "Neutralność oparta na kompensacji", None,
        [("„Jesteśmy firmą neutralną klimatycznie.”", g.ST["quote"]), ("„Neutralność wynika z kompensacji emisji.”", g.ST["quote_s"])],
        "<b>Ilustracja poziomu: wysokie.</b> Tylko wtedy, gdy spełniony jest test decyzji dotyczącej transakcji. Samo zbliżenie do zakazu z 27 września 2027 r., bez dzisiejszego wprowadzenia w błąd, tego poziomu nie daje. Może prowadzić do publikacji warunkowej, jeżeli przed tą datą twierdzenie wymaga zawężenia.",
        "Odbiorca może odczytać, że działalność nie zostawia emisji, a kompensacja stoi w dopisku. Do 26 września 2027 r. podstawą jest art. 5 albo art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, jeżeli przeciętny odbiorca może na tej podstawie podjąć decyzję dotyczącą transakcji, której inaczej by nie podjął. Od 27 września 2027 r. twierdzenie o neutralnym wpływie na emisje oparte na kompensacji wchodzi w krajowy zakaz. Wcześniej audyt opisuje je jako przyszły standard.",
        [("„W [roku] emisje zakresu 1 i 2 były o [X]% niższe niż w [roku bazowym]. Emisji pozostałych w tym zdaniu nie kompensujemy. Rachunek i granice są pod [adresem].”", g.ST["repair"]), ("Zdania „nie emitujemy gazów cieplarnianych”, gdy emisje występują, nie da się uratować dopiskiem. Wychodzą z materiału.", g.ST["body_s"])],
    )


def p_ex2(c, page, total):
    p_example(
        c, page, total, "Przykład 2 z 4", "Przykład: zakres twierdzenia",
        "Hasło o całości, dowód o fragmencie",
        "W przykładzie jedyny opublikowany rachunek dotyczy jednego zakładu, jednego roku oraz zakresów 1 i 2. Hasło obejmuje całą ofertę.",
        [("„Cała oferta jest zeroemisyjna.”", g.ST["quote"])],
        "<b>Ilustracja poziomu: wysokie.</b> Tylko wtedy, gdy odbiorca może podjąć decyzję dotyczącą transakcji w przekonaniu, że wynik dotyczy wszystkiego, co kupuje. Pochodzenie dowodu z innej kategorii produktów ocenia się osobno i nie dopisuje się go do przyszłego punktu o rozjechaniu zakresu.",
        "Do 26 września 2027 r. jest to pytanie o art. 5 albo art. 6. Od 27 września 2027 r. dochodzi przyszły standard: twierdzenie środowiskowe o całym produkcie albo o całej działalności, które w istocie dotyczy tylko określonego aspektu albo konkretnego rodzaju działalności. Dowód z innej kategorii produktów nie jest przesłanką tego punktu. Niewspółmierność takiej kategorii czyta się osobno, jeżeli spełniony jest test decyzji dotyczącej transakcji.",
        [("„Zakład w [miejscowości] obniżył w [roku] emisje zakresu 1 i 2 o [X]% względem [roku bazowego]. Wynik nie obejmuje pozostałych zakładów ani zakresu 3.”", g.ST["repair"]), ("Nawiasy uzupełnia nadawca danymi, które już ma. Hasło o całej ofercie bez takiego rachunku nie wraca do materiału w poprzednim brzmieniu.", g.ST["body_s"])],
    )


def p_ex3(c, page, total):
    p_example(
        c, page, total, "Przykład 3 z 4", "Przykład: porównanie",
        "Porównanie w kategorii",
        "W przykładzie przy haśle nie ma metody, roku, jednostki, granic analizy ani listy produktów objętych zestawieniem.",
        [("„Nasz ślad węglowy jest najniższy w kategorii.”", g.ST["quote"])],
        "<b>Ilustracja poziomu: średnie.</b> Gdy porównanie spełnia przesłanki i da się je naprawić doprecyzowaniem. Poziom rośnie, gdy liczba jest myląca albo dyskredytuje konkurencję. Kierunkiem naprawy jest uzupełnienie założeń, nie usunięcie zestawienia z samego faktu, że kategorie zestawiono.",
        "Brak nazwy konkurenta nie zamyka pytania o reklamę porównawczą. Potrzebna jest możliwość zidentyfikowania konkurenta albo jego towarów oraz to, że towary są konkurencyjne. Tak wyrok TSUE z 19 kwietnia 2007 r., C-381/05, De Landtsheer. Zestawienie może zostać, jeżeli jest obiektywne i sprawdzalne. Tak wyrok TSUE z 19 września 2006 r., C-356/04, Lidl Belgium. Założenia, które rozstrzygają wynik, mają być czytelne w samym przekazie. Z wyroku TSUE z 8 lutego 2017 r., C-562/15, Carrefour Hypermarchés, nie wynika ogólny obowiązek powołania określonej normy ISO w każdej reklamie. Liczba bez istotnego założenia może być zaniechaniem z art. 6, gdy wpływa na decyzję dotyczącą transakcji.",
        [("„Według [metody], za [rok], w granicach od [etapu] do [etapu], jeden [jednostka] ma ślad [X] kg CO2e. Zestawienie obejmuje [liczbę] produktów sprzedawanych w [kanale]. Założenia są pod [adresem].”", g.ST["repair"]), ("Korzystne, rzetelne porównanie zostaje. Nawiasy uzupełnia nadawca.", g.ST["body_s"])],
    )


def p_ex4(c, page, total):
    y = g.chrome(c, "Przykład 4 z 4", page, total, "Przykład: znak, obowiązek, etyka")
    y = g.section_title(c, "Znak, obowiązek prawny i etyka", "Trzy kolejne zdania ułożone na potrzeby publikacji. Każde idzie inną ścieżką skali.", y, kicker="Przykłady konstrukcyjne")
    bands = (
        ("Przyszły standard znaku", g.ST["h2_blue"], "„Certyfikat zrównoważonego rozwoju” jako plakietka ułożona przez samego nadawcę, bez systemu certyfikacji i bez ustanowienia przez organ publiczny.", "Od 27 września 2027 r. taki znak wchodzi w przyszły standard. Dziś pyta się, czy odbiorca odbiera plakietkę jako potwierdzenie osoby trzeciej albo organu i czy może na tej podstawie podjąć decyzję dotyczącą transakcji. Brak numeru certyfikatu w tym samym akapicie sam tego nie dowodzi. Przy żywności terminy „eko”, „bio” i „organic” czyta rozporządzenie (UE) 2018/848, nie ogólny zakaz haseł środowiskowych.", "Po korekcie: „Parametr [nazwa] dla [produktu] w [roku] wynosi [X], według [metody].” Plakietki bez systemu certyfikacji albo organu publicznego nie zostawia się."),
        ("Przyszły standard przewagi", g.ST["h2_blue"], "„Jako jedyni spełniamy unijny limit dla tej kategorii.”", "Samo poinformowanie o zgodności z limitem nie wystarcza do naruszenia. Od 27 września 2027 r. przyszły standard obejmuje przedstawianie jako cechy wyróżniającej wymogu nałożonego prawem na wszystkie produkty danej kategorii na rynku unijnym. Trzeba ustalić obie rzeczy: czy obowiązek naprawdę dotyczy wszystkich produktów właściwej kategorii oraz czy przekaz sprzedaje go jako przewagę.", "Po korekcie: „Produkty tej kategorii podlegają limitowi [nazwa]. Spełniamy go.” Limitu wspólnego dla całej kategorii nie przedstawia się jako cechy wyróżniającej."),
        ("Poziom niski", g.ST["h2"], "„Wdrożyliśmy Wytyczne OECD w całym łańcuchu dostaw.”", "To zalecenie, nie wyrok. Na skali audytu zostaje przy ryzyku niskim, o ile to samo zdanie nie spełnia równocześnie art. 5 albo art. 6. Wtedy przepis obowiązujący ocenia się na jego szczeblu, a ścieżkę Krajowego Punktu Kontaktowego dopisuje się osobno. Skutek tej ścieżki: publiczna ocena, zalecenia i ryzyko wizerunkowe.", "Po korekcie: „W [roku] opisaliśmy działania wobec [grupy dostawców] w [dokumencie].” Zdanie nie poświadcza zgodności z Wytycznymi OECD, jeżeli takiego sprawdzenia nie było."),
    )
    for label, style, quote, body, repair in bands:
        y = stack(c, y, [(label, style), (quote, g.ST["quote_s"]), (body, g.ST["body_s"]), (f"<i>{repair}</i>", g.ST["body_s"])], fill=g.PALE, pad=8)
    g.finish("pion trzy", y)


def p_reco(c, page, total):
    y = g.chrome(c, "Rekomendacje", page, total, "Rekomendacje")
    y = g.section_title(c, "Rekomendacje w czterech etapach", "Terminy liczy się od daty konkretnego audytu. W tym opracowaniu takiej daty nie ma, bo nie ma oznaczonego materiału. Kolejność etapów zostaje.", y)
    for when, title, body in (
        ("1. Od daty audytu", "Etap natychmiastowy", "Wstrzymanie obejmuje wyłącznie twierdzenia, które już spełniają przesłanki prawa obowiązującego. Hasła objęte tylko standardem z 27 września 2027 r. do tego etapu nie wchodzą."),
        ("2. Przed 27 września 2027 r.", "Zawężenie przyszłego standardu", "Korekta haseł, które od tej daty spełnią krajowy zakaz: ogólne twierdzenie bez uznanej efektywności istotnej dla niego, własny znak bez systemu certyfikacji, kompensacja jako podstawa neutralności, całość oparta na jednym aspekcie, obowiązek prawny przedstawiony jako przewaga."),
        ("3. Przy korekcie materiałów", "Jedno brzmienie we wszystkich kanałach", "Ta sama wersja twierdzenia na stronie, w ofercie, w specyfikacji i w raporcie. Oprawa graficzna nie poszerza zdania. Założenia porównania stoją przy liczbie. Wzór po korekcie jest zdaniem, które da się obronić dowodem już posiadanym."),
        ("4. Na stałe", "Metodyka, słownik i szkolenie", "Właściciel dowodu, data, granice pomiaru i słownik sformułowań dopuszczalnych. Przegląd przed publikacją obejmuje osoby, które piszą oferty i odpowiedzi na ankiety. Twierdzenie o ludziach i o etyce przechodzi tę samą ścieżkę."),
    ):
        y = stack(c, y, [(when, g.ST["h2_blue"]), (title, g.ST["h2"]), (body, g.ST["body"])])
    g.finish("pion rekomendacje", y)


def p_close(c, page, total):
    y = g.chrome(c, "Nota", page, total, "Granice metody")
    y = g.section_title(c, "Granice metody", "Pięć granic, które audyt stosuje przy każdym materiale. Krótka lista ustaleń jest wtedy wynikiem poprawnym.", y)
    for i, text in enumerate((
        "Przykładu z jednego audytu nie przenosi się na inny podmiot. Test prawny zostaje, przypisanie przypadku odpada.",
        "Przyszły standard EmpCo, do 26 września 2027 r., opisuje się osobno od naruszenia prawa już obowiązującego.",
        "Same Wytyczne OECD zostają przy ryzyku niskim i same nie podnoszą werdyktu.",
        "Porównania nie usuwa się wyłącznie dlatego, że zestawiono kategorie. Uzupełnia się założenia albo zostawia zestawienie, gdy jest rzetelne.",
        "Braku numeru certyfikatu w jednym akapicie nie traktuje się jako dowodu naruszenia.",
    ), 1):
        y = stack(c, y, [(f"<font color='#27769D'><b>{i}.</b></font>  {text}", g.ST["body"])], fill=g.PALE if i % 2 else g.WHITE, pad=7, gap=4)
    y = stack(c, y, [
        ("ESG Institute", g.ST["h2"]),
        ("info@esginstitute.eu&nbsp;&nbsp;–&nbsp;&nbsp;esginstitute.eu&nbsp;&nbsp;–&nbsp;&nbsp;Sygn. ESG/2026/11/08-P", g.ST["body_s"]),
        ("Teksty ustaw sprawdza się w ISAP, prawo Unii w EUR-Lex, orzecznictwo Trybunału Sprawiedliwości Unii Europejskiej w EUR-Lex, orzecznictwo polskie w SAOS i w bazie Sądu Najwyższego, a w sprawach administracyjnych także w bazie NSA. Opracowanie streszcza metodę według stanu na 8 listopada 2026 r. Audyt konkretnego materiału powstaje na podstawie przekazanego dokumentu. To opracowanie go nie zastępuje.", g.ST["body_s"]),
    ], fill=g.PALE)
    g.finish("pion zamknięcie", y)


PION = [p_cover, p_scope, p_principles, p_regime, p_verdict, p_steps, p_conclusions, p_ex1, p_ex2, p_ex3, p_ex4, p_reco, p_close]


def build_pion():
    tune_portrait()
    c = canvas.Canvas(str(PION_OUT), pagesize=A4)
    c.setTitle("Audyt ryzyk greenwashingu. Metodyka, wnioski i rekomendacje")
    c.setAuthor("ESG Institute")
    c.setSubject("Metodyka, wnioski i rekomendacje. 8 listopada 2026 r.")
    total = len(PION)
    for i, fn in enumerate(PION, 1):
        if i > 1:
            c.showPage()
        fn(c, i, total)
    c.save()
    print("Zapisano", PION_OUT)


def assert_clean(paths):
    banned = (
        "Materiał do publikacji na stronie",
        "OPRACOWANIE PUBLICZNE",
        "Żaden z tych wyroków nie dotyczy",
        "Hasło i dopisek ułożono",
        "Opracowanie publiczne",
    )
    blob = ""
    for path in paths:
        data = Path(path).read_bytes()
        blob += data.decode("latin-1", errors="ignore")
    for phrase in banned:
        if phrase in blob:
            print("ZOSTAŁO:", phrase, "w", path)


def main():
    g.main()
    build_pion()
    build_pptx()
    assert_clean([g.OUT, PION_OUT, PPTX_OUT])


if __name__ == "__main__":
    main()
