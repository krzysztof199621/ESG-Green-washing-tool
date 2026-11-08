#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Publiczna prezentacja metody audytu ryzyk greenwashingu.

PDF do strony ESG Institute. Bez nazw przedsiębiorców.
Zdania w przykładach są ułożone na potrzeby tej publikacji.
Kalendarz: standard audytu na 8 listopada 2026 r.
"""

import io
from pathlib import Path

from PIL import Image
from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

W, H = 1280, 720
ML = 40
CW = W - 2 * ML
FOOTER_SAFE = 46

NAVY = HexColor("#143B4E")
INK = HexColor("#333333")
MUTED = HexColor("#666666")
BLUE = HexColor("#3399CC")
BLUE_TEXT = HexColor("#27769D")
PALE = HexColor("#F5F7F8")
QUOTE_BG = HexColor("#F5F7F8")
LINE = HexColor("#E3E6E8")
HAIR = HexColor("#ECEEEF")
WHITE = white

FONT_DIR = Path("/System/Library/Fonts/Supplemental")
LOGO_SRC = Path(
    "/Users/krzysztofsulinski/Desktop/AI Projects/ESG"
    "/ESG Experts Day - Komplet Materialow (PDF + PPTX)"
    "/03_DLA_PROWADZACEGO/assets/esg_logo_white.png"
)
OUT = Path(__file__).resolve().parent / "ESG_Institute_Audyt_ryzyk_greenwashingu.pdf"

LOGO = None
LOGO_RATIO = 3.46


def register_fonts():
    pdfmetrics.registerFont(TTFont("Arial", str(FONT_DIR / "Arial.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Bold", str(FONT_DIR / "Arial Bold.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Italic", str(FONT_DIR / "Arial Italic.ttf")))
    pdfmetrics.registerFont(
        TTFont("Arial-BoldItalic", str(FONT_DIR / "Arial Bold Italic.ttf"))
    )
    registerFontFamily(
        "Arial",
        normal="Arial",
        bold="Arial-Bold",
        italic="Arial-Italic",
        boldItalic="Arial-BoldItalic",
    )


def styles():
    def ps(name, font, size, color, leading):
        return ParagraphStyle(
            name,
            fontName=font,
            fontSize=size,
            textColor=color,
            leading=leading,
            alignment=TA_LEFT,
            splitLongWords=False,
        )

    return {
        "h1": ps("h1", "Arial-Bold", 26, NAVY, 31),
        "lead": ps("lead", "Arial", 12.5, MUTED, 17),
        "h2": ps("h2", "Arial-Bold", 13, NAVY, 17),
        "h2_blue": ps("h2_blue", "Arial-Bold", 12.5, BLUE_TEXT, 16),
        "body": ps("body", "Arial", 12, INK, 16.4),
        "body_s": ps("body_s", "Arial", 11, INK, 14.8),
        "small": ps("small", "Arial", 9.5, MUTED, 12.6),
        "kicker": ps("kicker", "Arial-Bold", 10, BLUE_TEXT, 13),
        "quote": ps("quote", "Arial-Italic", 15, NAVY, 20),
        "quote_s": ps("quote_s", "Arial-Italic", 11.5, MUTED, 15.5),
        "repair": ps("repair", "Arial-Italic", 12, INK, 16.4),
        "cover_kicker": ps("cover_kicker", "Arial", 11, MUTED, 14),
        "cover_h": ps("cover_h", "Arial-Bold", 36, NAVY, 42),
        "cover_sub": ps("cover_sub", "Arial", 15, INK, 21),
        "cover_note": ps("cover_note", "Arial", 12, MUTED, 16.5),
        "cover_h2": ps("cover_h2", "Arial-Bold", 12, NAVY, 15),
        "cover_body": ps("cover_body", "Arial", 11, INK, 14.6),
        "foot": ps("foot", "Arial", 9, MUTED, 12),
    }


ST = None


def prepare_logo():
    """Znak na jasne tło: błękitny brylant i granatowy napis."""
    im = Image.open(LOGO_SRC).convert("RGBA")
    pix = im.load()
    w, h = im.size
    split = 220
    for yy in range(h):
        for xx in range(w):
            r, g, b, _a = pix[xx, yy]
            lum = max(r, g, b) / 255
            if lum < 0.12:
                pix[xx, yy] = (0, 0, 0, 0)
                continue
            if xx < split:
                t = (lum - 0.12) / 0.88
                c0, c1, c2 = (20, 74, 102), (51, 153, 204), (214, 236, 245)
                if t < 0.55:
                    u = t / 0.55
                    col = tuple(int(c0[i] + (c1[i] - c0[i]) * u) for i in range(3))
                else:
                    u = (t - 0.55) / 0.45
                    col = tuple(int(c1[i] + (c2[i] - c1[i]) * u) for i in range(3))
                pix[xx, yy] = (*col, 255)
            else:
                alpha = int(min(255, max(0, (lum - 0.12) / 0.88 * 255)))
                pix[xx, yy] = (20, 59, 78, alpha)
    im = im.crop(im.getbbox())
    im.save(Path(__file__).resolve().parent / "_logo_jasne.png", format="PNG")
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return ImageReader(buf), im.size[0] / im.size[1]


def mp(text, style, width):
    p = Paragraph(text, style)
    _w, h = p.wrap(width, 5000)
    return p, h


def panel(c, x, y, w, h, fill, border=LINE, sw=0.7):
    c.setFillColor(fill)
    c.setStrokeColor(border)
    c.setLineWidth(sw)
    c.rect(x, y, w, h, fill=1, stroke=1)


def panels_row(c, x, top, col_w, gap, columns, pad=13, inner_gap=5, min_h=0):
    prepared = []
    for col in columns:
        inner = col_w - 2 * pad
        measured = [mp(t, s, inner) for t, s in col["blocks"]]
        content = sum(h for _p, h in measured) + inner_gap * max(0, len(measured) - 1)
        prepared.append((col, measured, content + 2 * pad))
    height = max([min_h] + [h for _c, _m, h in prepared])
    for i, (col, measured, _h) in enumerate(prepared):
        cx = x + i * (col_w + gap)
        bottom = top - height
        panel(c, cx, bottom, col_w, height, col.get("fill", WHITE), col.get("border", LINE), sw=0.6)
        y = top - pad
        for p, ph in measured:
            p.drawOn(c, cx + pad, y - ph)
            y -= ph + inner_gap
    return height


def draw_logo(c, x, y, height):
    if LOGO is None:
        c.setFillColor(NAVY)
        c.setFont("Arial-Bold", 13)
        c.drawString(x, y + 8, "ESG INSTITUTE")
        return
    c.drawImage(
        LOGO,
        x,
        y,
        width=height * LOGO_RATIO,
        height=height,
        mask="auto",
        preserveAspectRatio=True,
        anchor="sw",
    )


def chrome(c, section, page, total, outline):
    c.setFillColor(WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.bookmarkPage(f"p{page}")
    c.addOutlineEntry(outline, f"p{page}", level=0, closed=False)
    draw_logo(c, 36, H - 46, 26)
    c.setFillColor(MUTED)
    c.setFont("Arial", 10)
    c.drawRightString(W - 36, H - 34, section)
    c.setStrokeColor(LINE)
    c.setLineWidth(0.8)
    c.line(0, H - 56, W, H - 56)
    c.setStrokeColor(HAIR)
    c.setLineWidth(0.5)
    c.line(ML, 34, W - ML, 34)
    c.setFillColor(MUTED)
    c.setFont("Arial", 9)
    c.drawString(ML, 16, "ESG Institute   –   8 listopada 2026 r.")
    c.drawRightString(W - ML, 16, f"{page}  /  {total}")
    return H - 78


def finish(name, y):
    print(f"{name}: dolna krawędź treści {y:.0f}")
    if y < FOOTER_SAFE:
        print(f"  UWAGA: treść wchodzi w stopkę ({name})")


def section_title(c, title, lead, top, kicker=None):
    y = top
    if kicker:
        _p, h = mp(kicker, ST["kicker"], CW)
        _p.drawOn(c, ML, y - h)
        y -= h + 3
    _p, h = mp(title, ST["h1"], CW)
    _p.drawOn(c, ML, y - h)
    y -= h + 8
    c.setFillColor(BLUE)
    c.rect(ML, y, 28, 2, fill=1, stroke=0)
    y -= 14
    if lead:
        _p, h = mp(lead, ST["lead"], CW)
        _p.drawOn(c, ML, y - h)
        y -= h + 14
    return y


def col_width(n, gap):
    return (CW - gap * (n - 1)) / n


def slide_cover(c, page, total, section, outline):
    c.bookmarkPage(f"p{page}")
    c.addOutlineEntry(outline, f"p{page}", level=0, closed=False)
    c.setFillColor(WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    draw_logo(c, 56, H - 88, 40)
    y = H - 140
    p, h = mp("Audyt ryzyk greenwashingu", ST["cover_h"], 1000)
    p.drawOn(c, 56, y - h)
    y -= h + 14
    c.setFillColor(BLUE)
    c.rect(56, y, 32, 2, fill=1, stroke=0)
    y -= 22
    p, h = mp(
        "Jak ESG Institute ocenia twierdzenia środowiskowe i społeczne, zanim materiał zostanie opublikowany. Metoda, wnioski, przykładowe stwierdzenia i rekomendacje.",
        ST["cover_sub"],
        860,
    )
    p.drawOn(c, 56, y - h)
    y -= h

    band_h = 176
    index = [("01", "Metoda"), ("02", "Reżim"), ("03", "Werdykt"), ("04", "Przykłady"), ("05", "Rekomendacje")]
    ix = 56
    for num, label in index:
        c.setFillColor(BLUE_TEXT)
        c.setFont("Arial-Bold", 11)
        c.drawString(ix, band_h + 28, num)
        c.setFillColor(INK)
        c.setFont("Arial", 12)
        c.drawString(ix + 24, band_h + 28, label)
        ix += 168
    c.setFillColor(PALE)
    c.rect(0, 0, W, band_h, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.rect(0, band_h, W, 1.5, fill=1, stroke=0)
    facts = [
        (
            "Bez oznaczonego podmiotu",
            "Zdania w przykładach ułożono na potrzeby tej publikacji. Nie są cytatami z materiałów konkretnej organizacji.",
        ),
        (
            "Dwa reżimy",
            "Do 26 września 2027 r. werdykt opiera się na przepisach już obowiązujących. Czarna lista EmpCo jest przyszłym standardem.",
        ),
        (
            "Krótka lista ustaleń",
            "Audyt zostawia twierdzenie tylko wtedy, gdy przesłanki są spełnione. Nie dąży do z góry przyjętej liczby uwag.",
        ),
    ]
    gap = 28
    inner_w = (W - 56 - 48 - 2 * gap) / 3
    for i, (title, body) in enumerate(facts):
        x = 56 + i * (inner_w + gap)
        p, h = mp(title, ST["cover_h2"], inner_w)
        p.drawOn(c, x, band_h - 28 - h)
        p, _bh = mp(body, ST["cover_body"], inner_w)
        p.drawOn(c, x, band_h - 36 - h - _bh)
        if i:
            c.setStrokeColor(LINE)
            c.setLineWidth(0.6)
            c.line(x - gap / 2, 28, x - gap / 2, band_h - 24)
    finish("okładka", band_h)


def slide_scope(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Zakres opracowania",
        "Publikacja przedstawia metodę audytu ryzyk greenwashingu stosowaną w ESG Institute. Nie zbiera wyników poszczególnych audytów i nie podaje liczby naruszeń na rynku.",
        y,
    )
    gap = 14
    width = col_width(3, gap)
    cols = [
        {
            "fill": WHITE,
            "blocks": [
                ("Co wchodzi do audytu", ST["h2"]),
                (
                    "Przekaz przygotowany do publikacji: strona, oferta, specyfikacja, etykieta, prezentacja sprzedażowa albo fragment raportu. Ocenie podlega brzmienie i sposób ujęcia, łącznie z oprawą, która poszerza sens słów.",
                    ST["body"],
                ),
            ],
        },
        {
            "fill": WHITE,
            "blocks": [
                ("Co jest wynikiem", ST["h2"]),
                (
                    "Werdykt o publikacji tego przekazu, ustalenia spełniające przesłanki oraz kierunek naprawy. Krótka lista jest wynikiem poprawnym. Audyt nie przyjmuje docelowej liczby uwag.",
                    ST["body"],
                ),
            ],
        },
        {
            "fill": WHITE,
            "blocks": [
                ("Czego materiał nie rozstrzyga", ST["h2"]),
                (
                    "Nie ocenia całej działalności organizacji i nie zastępuje analizy konkretnego dokumentu. Stan prawa w tym opracowaniu: 8 listopada 2026 r. Tekst przepisu przy konkretnym audycie sprawdza się w ISAP i w EUR-Lex.",
                    ST["body"],
                ),
            ],
        },
    ]
    y -= panels_row(c, ML, y, width, gap, cols, pad=16)
    y -= 14
    y -= panels_row(
        c,
        ML,
        y,
        CW,
        0,
        [
            {
                "fill": WHITE,
                "accent": BLUE_TEXT,
                "blocks": [
                    ("Odbiorca opracowania", ST["h2"]),
                    (
                        "Działy komunikacji, zgodności i osoby zatwierdzające publikację. Przy konkretnym zleceniu audyt czyta wyłącznie przekazany dokument. Szuka wykładni dla podobnego stanu faktycznego, a nie nazwy przedsiębiorcy. Orzeczenie z innej branży wolno przywołać jako analogię, z zaznaczeniem, czego dotyczy i że nie rozstrzyga zgodności z prawem działalności badanego podmiotu.",
                        ST["body"],
                    ),
                ],
            }
        ],
        pad=14,
    )
    finish("zakres", y)


def slide_principles(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Jak czyta się twierdzenie",
        "Cztery reguły metody. Tezy orzeczeń podano jako skrót wykładni przyjętej w standardzie audytu, nie jako cytat z sentencji.",
        y,
    )
    rows = [
        (
            "01",
            "Całość przekazu",
            "Hasło i późniejsze zastrzeżenie czyta się razem. Poprawna informacja niżej albo odesłanie na inną stronę nie zawsze zdejmuje pierwsze wrażenie. W metodzie ten kierunek wyznaczają wyroki TSUE: z 4 czerwca 2015 r., C-195/14, Teekanne, oraz – przy zaniechaniu – z 26 października 2016 r., C-611/14, Canal Digital Danmark.",
        ),
        (
            "02",
            "Decyzja dotycząca transakcji",
            "Przy art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym sprawdza się, czy przeciętny odbiorca może podjąć decyzję, której inaczej by nie podjął. Rzeczywisty zakup nie jest potrzebny. Sama nieścisłość albo brak liczby tego progu nie spełnia. Tak wyrok TSUE z 19 grudnia 2013 r., C-281/12, Trento Sviluppo.",
        ),
        (
            "03",
            "Zgodność zakresu z dowodem",
            "Informacja prawdziwa może mylić przez selektywny układ, pominięte założenie albo komentarz szerszy niż dowód. Cel nie jest wynikiem osiągniętym. Sąd Najwyższy w wyroku z 9 października 2014 r., IV CSK 56/14, oddalił skargę kasacyjną i nie uznał, że w tamtej sprawie samo użycie wyników badań było deliktem. Sprawdza się, czy zakres twierdzenia pokrywa się z zakresem dowodu.",
        ),
        (
            "04",
            "Przepis szczególny przed klauzulą ogólną",
            "Tam, gdzie temat ma własny akt, audyt zaczyna od niego. Przy oświadczeniach zdrowotnych i żywieniowych pierwszeństwo wobec dyrektywy 2005/29/WE potwierdza wyrok TSUE z 10 września 2020 r., C-363/19, Mezina. Finanse, transport, moda i budownictwo również czyta się najpierw przez regulację właściwą dla tematu.",
        ),
    ]
    for num, title, body in rows:
        tp, th = mp(title, ST["h2"], CW - 86)
        bp, bh = mp(body, ST["body"], CW - 86)
        row_h = max(78, 14 + th + 3 + bh + 12)
        panel(c, ML, y - row_h, CW, row_h, WHITE)
        c.setFillColor(BLUE_TEXT)
        c.setFont("Arial-Bold", 15)
        c.drawString(ML + 14, y - 28, num)
        tp.drawOn(c, ML + 58, y - 12 - th)
        bp.drawOn(c, ML + 58, y - 12 - th - 3 - bh)
        y -= row_h + 8
    finish("zasady", y)


def slide_regime(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Reżim prawny na 8 listopada 2026 r.",
        "Dwie daty zostają rozdzielone. Werdykt do 26 września 2027 r. włącznie opiera się na przepisach już obowiązujących.",
        y,
    )
    gap = 14
    width = col_width(2, gap)
    cols = [
        {
            "fill": WHITE,
            "accent": NAVY,
            "blocks": [
                ("Prawo, na którym stoi werdykt", ST["h2"]),
                (
                    "– art. 5 i art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, z testem decyzji dotyczącej transakcji;<br/>"
                    "– ustawa o ochronie konkurencji i konsumentów, gdy praktyka narusza zbiorowe interesy konsumentów; kara pieniężna Prezesa UOKiK może sięgnąć 10% obrotu;<br/>"
                    "– ustawa o zwalczaniu nieuczciwej konkurencji, gdy przekaz jest reklamą porównawczą;<br/>"
                    "– rozporządzenia sektorowe stosowane bezpośrednio, w tym przy produkcji ekologicznej i przy oświadczeniach zdrowotnych.",
                    ST["body"],
                ),
            ],
        },
        {
            "fill": QUOTE_BG,
            "accent": BLUE,
            "blocks": [
                ("Przyszły standard EmpCo", ST["h2_blue"]),
                (
                    "Dyrektywa (UE) 2024/825 z 28 lutego 2024 r. Art. 4 ust. 1 wyznacza 27 września 2026 r. jako datę, od której państwa członkowskie stosują krajowe przepisy transponujące. To data dyrektywy.<br/><br/>"
                    "Ustawa z 4 września 2026 r. o zmianie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawy o prawach konsumenta (Dz.U. z 2026 r. poz. 1353) została ogłoszona 19 października 2026 r. Wchodzi w życie 27 września 2027 r.<br/><br/>"
                    "Do tego dnia czarnej listy nie traktuje się jako już obowiązującego zakazu krajowego. Utrzymywanie starszej treści w internecie tej daty nie przyspiesza.",
                    ST["body"],
                ),
            ],
        },
    ]
    y -= panels_row(c, ML, y, width, gap, cols, pad=14, inner_gap=8)
    y -= 12
    y -= panels_row(
        c,
        ML,
        y,
        CW,
        0,
        [
            {
                "fill": PALE,
                "blocks": [
                    (
                        "Wniosek Komisji COM(2023) 166 w sprawie oświadczeń środowiskowych nie jest prawem. W audycie może opisać dobrą praktykę dowodową, wyłącznie z tym zastrzeżeniem. Wytyczne OECD dla przedsiębiorstw wielonarodowych, edycja 2023, są zaleceniami. Na skali audytu dostają poziom niski. Postępowanie przed polskim Krajowym Punktem Kontaktowym kończy się publiczną oceną i zaleceniami. Nie jest sprawą sądową i nie kończy się wyrokiem.",
                        ST["body_s"],
                    )
                ],
            }
        ],
        pad=12,
    )
    finish("reżim", y)


def slide_verdict(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Skala ryzyka i werdykt",
        "Skala mierzy naruszenie przepisu obowiązującego w dniu audytu. Samo zbliżenie do czarnej listy oraz samo powołanie Wytycznych OECD nie dają werdyktu wysokiego ani publikacji warunkowej.",
        y,
    )
    rows = [
        (
            "BARDZO WYSOKIE",
            NAVY,
            "Przepis już obowiązujący, spełnione przesłanki, w tym test decyzji dotyczącej transakcji, oraz liczbowe wprowadzenie w błąd, dyskredytowanie konkurencji albo realne ryzyko kary lub powództwa.",
        ),
        (
            "WYSOKIE",
            NAVY,
            "Bezpośrednie naruszenie art. 5 lub art. 6 albo już obowiązującego zakazu sektorowego.",
        ),
        (
            "ŚREDNIE",
            NAVY,
            "Brak weryfikowalności albo porównanie, które spełnia przesłanki reklamy porównawczej lub zaniechania i da się naprawić przez doprecyzowanie.",
        ),
        (
            "NISKIE",
            MUTED,
            "Drobna nieprecyzyjność poniżej progu naruszenia. Nie tworzy osobnego przypadku. Wytyczne OECD zostają na tym poziomie.",
        ),
        (
            "PRZYSZŁY STANDARD",
            BLUE_TEXT,
            "Do 26 września 2027 r. osobna etykieta: „Przyszły standard EmpCo (od 27.09.2027)”. Nie jest poziomem wysokim i nie jest stwierdzonym naruszeniem krajowej czarnej listy.",
        ),
    ]
    label_w = 196
    for i, (label, color, text) in enumerate(rows):
        _p, th = mp(text, ST["body_s"], CW - label_w - 28)
        row_h = max(36, th + 14)
        fill = PALE if i % 2 == 0 else WHITE
        panel(c, ML, y - row_h, CW, row_h, fill, border=fill, sw=0)
        c.setFillColor(color)
        c.setFont("Arial-Bold", 10)
        c.drawString(ML + 16, y - row_h / 2 - 3.5, label)
        _p.drawOn(c, ML + label_w, y - (row_h - th) / 2 - th)
        y -= row_h + 5
    y -= 6
    _p, h = mp("Werdykt operacyjny do 26 września 2027 r.", ST["h2"], CW)
    _p.drawOn(c, ML, y - h)
    y -= h + 8
    gap = 12
    width = col_width(3, gap)
    cards = [
        {
            "fill": WHITE,
            "blocks": [
                (
                    "Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści",
                    ST["h2"],
                ),
                (
                    "Co najmniej jedno ustalenie bardzo wysokie albo wysokie na prawie już obowiązującym. Sam przyszły standard tego werdyktu nie daje.",
                    ST["body_s"],
                ),
            ],
        },
        {
            "fill": WHITE,
            "blocks": [
                (
                    "Publikacja warunkowa – rekomendowana korekta przed publikacją",
                    ST["h2"],
                ),
                (
                    "Brak takiego ustalenia, ale jest ryzyko średnie albo twierdzenie dziś mieszczące się w art. 5 i 6, które od 27 września 2027 r. spełni zakaz i wcześniej wymaga zawężenia.",
                    ST["body_s"],
                ),
            ],
        },
        {
            "fill": WHITE,
            "blocks": [
                ("Dopuszczone do publikacji", ST["h2"]),
                (
                    "Brak istotnych naruszeń prawa obowiązującego. Krótka wzmianka o przyszłym standardzie, która nie wymaga wstrzymania, zostaje we wniosku.",
                    ST["body_s"],
                ),
            ],
        },
    ]
    y -= panels_row(c, ML, y, width, gap, cards, pad=10, inner_gap=4, min_h=118)
    y -= 8
    _p, h = mp(
        "Od 27 września 2027 r. spełniony zakaz z ustawy może uzasadniać wstrzymanie publikacji. Nadal tylko przy pełnych przesłankach przepisu.",
        ST["small"],
        CW,
    )
    _p.drawOn(c, ML, y - h)
    y -= h
    finish("werdykt", y)


def slide_steps(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Tok audytu",
        "Sześć kroków przy jednym dokumencie. Data audytu wskazuje reżim zapisany potem w metryczce.",
        y,
    )
    steps = [
        (
            "1",
            "Dokument i data",
            "Kanał, odbiorca, dosłowne brzmienie, oprawa i data. Data rozstrzyga, czy czarna lista jest już prawem polskim.",
        ),
        (
            "2",
            "Wypisanie twierdzeń",
            "Zdania środowiskowe i społeczne w brzmieniu z materiału, z miejscem. Oprawa wchodzi do sensu, gdy poszerza słowa.",
        ),
        (
            "3",
            "Test prawa obowiązującego",
            "Najpierw przepis sektorowy. Potem art. 5 i art. 6, w tym decyzja dotycząca transakcji.",
        ),
        (
            "4",
            "Osobno standard z 2027 r.",
            "To samo zdanie może dostać drugą podstawę. Do 27 września 2027 r. nie jest naruszeniem krajowej czarnej listy.",
        ),
        (
            "5",
            "Zostają spełnione przesłanki",
            "Twierdzenie poniżej progu nie dostaje numeru. Może wrócić przy przyszłym standardzie, jeżeli przed tą datą wymaga zawężenia.",
        ),
        (
            "6",
            "Naprawa i wzór",
            "Kierunek korekty oraz zdanie, które da się opublikować. Gdy twierdzenia nie da się uratować, wzór odpada.",
        ),
    ]
    gap_x, gap_y = 14, 12
    width = col_width(2, gap_x)
    cells = []
    for num, title, body in steps:
        blocks = [
            (f"{num}.  {title}", ST["h2_blue"]),
            (body, ST["body"]),
        ]
        cells.append({"fill": WHITE, "blocks": blocks})
    for row in range(3):
        pair = cells[row * 2 : row * 2 + 2]
        h = panels_row(c, ML, y, width, gap_x, pair, pad=14, inner_gap=4)
        y -= h + gap_y
    finish("tok", y)


def slide_conclusions(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Wnioski z metody",
        "Siedem wniosków opisuje sposób oceny. Nie są ustaleniami wobec żadnej organizacji.",
        y,
    )
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
        rich = f"<font color='#1F5E7E'><b>{i}.</b></font>  {text}"
        p, h = mp(rich, ST["body"], CW - 28)
        row_h = h + 12
        if i % 2 == 1:
            c.setFillColor(PALE)
            c.rect(ML, y - row_h, CW, row_h, fill=1, stroke=0)
        p.drawOn(c, ML + 14, y - 6 - h)
        y -= row_h
    finish("wnioski", y)


def render_example(
    c, page, total, section, outline, title, context, quote_blocks, status, status_accent, assessment, repair_blocks
):
    y = chrome(c, section, page, total, outline)
    y = section_title(c, title, context, y, kicker="Przykład konstrukcyjny")
    h = panels_row(
        c,
        ML,
        y,
        CW,
        0,
        [{"fill": QUOTE_BG, "accent": BLUE, "blocks": quote_blocks}],
        pad=14,
        inner_gap=3,
    )
    y -= h + 10
    h = panels_row(
        c,
        ML,
        y,
        CW,
        0,
        [{"fill": WHITE, "accent": status_accent, "blocks": [(status, ST["body_s"])]}],
        pad=10,
    )
    y -= h + 12
    gap = 14
    width = col_width(2, gap)
    cols = [
        {
            "fill": WHITE,
            "blocks": [("Ocena", ST["h2"]), (assessment, ST["body"])],
        },
        {
            "fill": PALE,
            "blocks": [("Dozwolone sformułowanie", ST["h2"]), *repair_blocks],
        },
    ]
    y -= panels_row(c, ML, y, width, gap, cols, pad=14, inner_gap=6)
    finish(title, y)


def slide_ex_offset(c, page, total, section, outline):
    render_example(
        c,
        page,
        total,
        section,
        outline,
        "Neutralność oparta na kompensacji",
        None,
        [
            ("„Jesteśmy firmą neutralną klimatycznie.”", ST["quote"]),
            ("„Neutralność wynika z kompensacji emisji.”", ST["quote_s"]),
        ],
        "<b>Ilustracja poziomu: wysokie.</b> Tylko wtedy, gdy spełniony jest test decyzji dotyczącej transakcji. Samo zbliżenie do zakazu z 27 września 2027 r., bez dzisiejszego wprowadzenia w błąd, tego poziomu nie daje. Może prowadzić do publikacji warunkowej, jeżeli przed tą datą twierdzenie wymaga zawężenia.",
        NAVY,
        "Odbiorca może odczytać, że działalność nie zostawia emisji, a kompensacja stoi w dopisku. Do 26 września 2027 r. podstawą jest art. 5 albo art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, jeżeli przeciętny odbiorca może na tej podstawie podjąć decyzję dotyczącą transakcji, której inaczej by nie podjął. Od 27 września 2027 r. twierdzenie o neutralnym wpływie na emisje oparte na kompensacji wchodzi w krajowy zakaz. Wcześniej audyt opisuje je jako przyszły standard.",
        [
            (
                "„W [roku] emisje zakresu 1 i 2 były o [X]% niższe niż w [roku bazowym]. Emisji pozostałych w tym zdaniu nie kompensujemy. Rachunek i granice są pod [adresem].”",
                ST["repair"],
            ),
            (
                "Zdania „nie emitujemy gazów cieplarnianych”, gdy emisje występują, nie da się uratować dopiskiem. Wychodzą z materiału.",
                ST["body_s"],
            ),
        ],
    )


def slide_ex_scope(c, page, total, section, outline):
    render_example(
        c,
        page,
        total,
        section,
        outline,
        "Hasło o całości, dowód o fragmencie",
        "W przykładzie jedyny opublikowany rachunek dotyczy jednego zakładu, jednego roku oraz zakresów 1 i 2. Hasło obejmuje całą ofertę.",
        [("„Cała oferta jest zeroemisyjna.”", ST["quote"])],
        "<b>Ilustracja poziomu: wysokie.</b> Tylko wtedy, gdy odbiorca może podjąć decyzję dotyczącą transakcji w przekonaniu, że wynik dotyczy wszystkiego, co kupuje. Pochodzenie dowodu z innej kategorii produktów ocenia się osobno i nie dopisuje się go do przyszłego punktu o rozjechaniu zakresu.",
        NAVY,
        "Do 26 września 2027 r. jest to pytanie o art. 5 albo art. 6. Od 27 września 2027 r. dochodzi przyszły standard: twierdzenie środowiskowe o całym produkcie albo o całej działalności, które w istocie dotyczy tylko określonego aspektu albo konkretnego rodzaju działalności. Dowód z innej kategorii produktów nie jest przesłanką tego punktu. Niewspółmierność takiej kategorii czyta się osobno, jeżeli spełniony jest test decyzji dotyczącej transakcji.",
        [
            (
                "„Zakład w [miejscowości] obniżył w [roku] emisje zakresu 1 i 2 o [X]% względem [roku bazowego]. Wynik nie obejmuje pozostałych zakładów ani zakresu 3.”",
                ST["repair"],
            ),
            (
                "Nawiasy uzupełnia nadawca danymi, które już ma. Hasło o całej ofercie bez takiego rachunku nie wraca do materiału w poprzednim brzmieniu.",
                ST["body_s"],
            ),
        ],
    )


def slide_ex_compare(c, page, total, section, outline):
    render_example(
        c,
        page,
        total,
        section,
        outline,
        "Porównanie w kategorii",
        "W przykładzie przy haśle nie ma metody, roku, jednostki, granic analizy ani listy produktów objętych zestawieniem.",
        [("„Nasz ślad węglowy jest najniższy w kategorii.”", ST["quote"])],
        "<b>Ilustracja poziomu: średnie.</b> Gdy porównanie spełnia przesłanki i da się je naprawić doprecyzowaniem. Poziom rośnie, gdy liczba jest myląca albo dyskredytuje konkurencję. Kierunkiem naprawy jest uzupełnienie założeń, nie usunięcie zestawienia z samego faktu, że kategorie zestawiono.",
        BLUE_TEXT,
        "Brak nazwy konkurenta nie zamyka pytania o reklamę porównawczą. Potrzebna jest możliwość zidentyfikowania konkurenta albo jego towarów oraz to, że towary są konkurencyjne. Tak wyrok TSUE z 19 kwietnia 2007 r., C-381/05, De Landtsheer. Zestawienie może zostać, jeżeli jest obiektywne i sprawdzalne. Tak wyrok TSUE z 19 września 2006 r., C-356/04, Lidl Belgium. Założenia, które rozstrzygają wynik, mają być czytelne w samym przekazie. Z wyroku TSUE z 8 lutego 2017 r., C-562/15, Carrefour Hypermarchés, nie wynika ogólny obowiązek powołania określonej normy ISO w każdej reklamie. Liczba bez istotnego założenia może być zaniechaniem z art. 6, gdy wpływa na decyzję dotyczącą transakcji.",
        [
            (
                "„Według [metody], za [rok], w granicach od [etapu] do [etapu], jeden [jednostka] ma ślad [X] kg CO2e. Zestawienie obejmuje [liczbę] produktów sprzedawanych w [kanale]. Założenia są pod [adresem].”",
                ST["repair"],
            ),
            (
                "Korzystne, rzetelne porównanie zostaje. Nawiasy uzupełnia nadawca.",
                ST["body_s"],
            ),
        ],
    )


def slide_ex_three(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Znak, obowiązek prawny i etyka",
        "Trzy kolejne zdania ułożone na potrzeby publikacji. Każde idzie inną ścieżką skali.",
        y,
        kicker="Przykłady konstrukcyjne",
    )
    bands = [
        (
            BLUE,
            QUOTE_BG,
            "Przyszły standard znaku",
            "„Certyfikat zrównoważonego rozwoju” jako plakietka ułożona przez samego nadawcę, bez systemu certyfikacji i bez ustanowienia przez organ publiczny.",
            "Od 27 września 2027 r. taki znak wchodzi w przyszły standard. Dziś pyta się, czy odbiorca odbiera plakietkę jako potwierdzenie osoby trzeciej albo organu i czy może na tej podstawie podjąć decyzję dotyczącą transakcji. Brak numeru certyfikatu w tym samym akapicie sam tego nie dowodzi. Przy żywności terminy „eko”, „bio” i „organic” czyta rozporządzenie (UE) 2018/848, nie ogólny zakaz haseł środowiskowych.",
            "Po korekcie: „Parametr [nazwa] dla [produktu] w [roku] wynosi [X], według [metody].” Plakietki bez systemu certyfikacji albo organu publicznego nie zostawia się.",
        ),
        (
            BLUE,
            QUOTE_BG,
            "Przyszły standard przewagi",
            "„Jako jedyni spełniamy unijny limit dla tej kategorii.”",
            "Samo poinformowanie o zgodności z limitem nie wystarcza do naruszenia. Od 27 września 2027 r. przyszły standard obejmuje przedstawianie jako cechy wyróżniającej wymogu nałożonego prawem na wszystkie produkty danej kategorii na rynku unijnym. Trzeba ustalić obie rzeczy: czy obowiązek naprawdę dotyczy wszystkich produktów właściwej kategorii oraz czy przekaz sprzedaje go jako przewagę.",
            "Po korekcie: „Produkty tej kategorii podlegają limitowi [nazwa]. Spełniamy go.” Limitu wspólnego dla całej kategorii nie przedstawia się jako cechy wyróżniającej.",
        ),
        (
            MUTED,
            PALE,
            "Poziom niski",
            "„Wdrożyliśmy Wytyczne OECD w całym łańcuchu dostaw.”",
            "To zalecenie, nie wyrok. Na skali audytu zostaje przy ryzyku niskim, o ile to samo zdanie nie spełnia równocześnie art. 5 albo art. 6. Wtedy przepis obowiązujący ocenia się na jego szczeblu, a ścieżkę Krajowego Punktu Kontaktowego dopisuje się osobno. Skutek tej ścieżki: publiczna ocena, zalecenia i ryzyko wizerunkowe.",
            "Po korekcie: „W [roku] opisaliśmy działania wobec [grupy dostawców] w [dokumencie].” Zdanie nie poświadcza zgodności z Wytycznymi OECD, jeżeli takiego sprawdzenia nie było.",
        ),
    ]
    for accent, fill, label, quote, body, repair in bands:
        blocks = [
            (label, ST["h2_blue"] if accent == BLUE else ST["h2"]),
            (quote, ST["quote_s"]),
            (body, ST["body_s"]),
            (f"<i>{repair}</i>", ST["body_s"]),
        ]
        h = panels_row(
            c,
            ML,
            y,
            CW,
            0,
            [{"fill": fill, "accent": accent, "blocks": blocks}],
            pad=10,
            inner_gap=3,
        )
        y -= h + 8
    finish("trzy przykłady", y)


def slide_recommendations(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Rekomendacje w czterech etapach",
        "Terminy liczy się od daty konkretnego audytu. W tym opracowaniu takiej daty nie ma, bo nie ma oznaczonego materiału. Kolejność etapów zostaje.",
        y,
    )
    stages = [
        (
            "1",
            "Od daty audytu",
            "Etap natychmiastowy",
            "Wstrzymanie obejmuje wyłącznie twierdzenia, które już spełniają przesłanki prawa obowiązującego. Hasła objęte tylko standardem z 27 września 2027 r. do tego etapu nie wchodzą.",
            NAVY,
            "h2_blue",
        ),
        (
            "2",
            "Przed 27 września 2027 r.",
            "Zawężenie przyszłego standardu",
            "Korekta haseł, które od tej daty spełnią krajowy zakaz: ogólne twierdzenie bez uznanej efektywności istotnej dla niego, własny znak bez systemu certyfikacji, kompensacja jako podstawa neutralności, całość oparta na jednym aspekcie, obowiązek prawny przedstawiony jako przewaga.",
            BLUE,
            "h2_blue",
        ),
        (
            "3",
            "Przy korekcie materiałów",
            "Jedno brzmienie we wszystkich kanałach",
            "Ta sama wersja twierdzenia na stronie, w ofercie, w specyfikacji i w raporcie. Oprawa graficzna nie poszerza zdania. Założenia porównania stoją przy liczbie. Wzór po korekcie jest zdaniem, które da się obronić dowodem już posiadanym.",
            NAVY,
            "h2_blue",
        ),
        (
            "4",
            "Na stałe",
            "Metodyka, słownik i szkolenie",
            "Właściciel dowodu, data, granice pomiaru i słownik sformułowań dopuszczalnych. Przegląd przed publikacją obejmuje osoby, które piszą oferty i odpowiedzi na ankiety. Twierdzenie o ludziach i o etyce przechodzi tę samą ścieżkę.",
            NAVY,
            "h2_blue",
        ),
    ]
    gap = 14
    width = col_width(2, gap)
    for row in range(2):
        cols = []
        for num, when, title, body, accent, when_style in stages[row * 2 : row * 2 + 2]:
            cols.append(
                {
                    "fill": WHITE,
                    "accent": accent,
                    "blocks": [
                        (f"{num}.  {when}", ST[when_style]),
                        (title, ST["h2"]),
                        (body, ST["body"]),
                    ],
                }
            )
        h = panels_row(c, ML, y, width, gap, cols, pad=14, inner_gap=4)
        y -= h + 12
    finish("rekomendacje", y)


def slide_close(c, page, total, section, outline):
    y = chrome(c, section, page, total, outline)
    y = section_title(
        c,
        "Granice metody",
        "Pięć granic, które audyt stosuje przy każdym materiale. Krótka lista ustaleń jest wtedy wynikiem poprawnym.",
        y,
    )
    limits = [
        "Przykładu z jednego audytu nie przenosi się na inny podmiot. Test prawny zostaje, przypisanie przypadku odpada.",
        "Przyszły standard EmpCo, do 26 września 2027 r., opisuje się osobno od naruszenia prawa już obowiązującego.",
        "Same Wytyczne OECD zostają przy ryzyku niskim i same nie podnoszą werdyktu.",
        "Porównania nie usuwa się wyłącznie dlatego, że zestawiono kategorie. Uzupełnia się założenia albo zostawia zestawienie, gdy jest rzetelne.",
        "Braku numeru certyfikatu w jednym akapicie nie traktuje się jako dowodu naruszenia.",
    ]
    for i, text in enumerate(limits, 1):
        rich = f"<font color='#1F5E7E'><b>{i}.</b></font>  {text}"
        p, h = mp(rich, ST["body"], CW - 28)
        row_h = h + 10
        c.setFillColor(PALE if i % 2 else WHITE)
        c.rect(ML, y - row_h, CW, row_h, fill=1, stroke=0)
        p.drawOn(c, ML + 14, y - 5 - h)
        y -= row_h
    blocks = [
        ("ESG Institute", ST["cover_h2"]),
        (
            "info@esginstitute.eu&nbsp;&nbsp;–&nbsp;&nbsp;esginstitute.eu&nbsp;&nbsp;–&nbsp;&nbsp;Sygn. ESG/2026/11/08-P",
            ST["cover_body"],
        ),
        (
            "Teksty ustaw sprawdza się w ISAP, prawo Unii w EUR-Lex, orzecznictwo Trybunału Sprawiedliwości Unii Europejskiej w EUR-Lex, orzecznictwo polskie w SAOS i w bazie Sądu Najwyższego, a w sprawach administracyjnych także w bazie NSA. Opracowanie streszcza metodę według stanu na 8 listopada 2026 r. Audyt konkretnego materiału powstaje na podstawie przekazanego dokumentu. To opracowanie go nie zastępuje.",
            ST["cover_body"],
        ),
    ]
    pad, inner_gap = 16, 5
    inner = CW - 2 * pad
    heights = [mp(text, style, inner)[1] for text, style in blocks]
    panel_h = sum(heights) + inner_gap * (len(heights) - 1) + 2 * pad
    top = panel_h + 50
    if top > y - 14:
        top = y - 14
    panels_row(
        c,
        ML,
        top,
        CW,
        0,
        [{"fill": PALE, "border": LINE, "blocks": blocks}],
        pad=pad,
        inner_gap=inner_gap,
    )
    finish("zamknięcie", 50)


SLIDES = [
    (slide_cover, "", "Okładka"),
    (slide_scope, "Zakres", "Zakres opracowania"),
    (slide_principles, "Metoda", "Jak czyta się twierdzenie"),
    (slide_regime, "Prawo", "Reżim prawny"),
    (slide_verdict, "Werdykt", "Skala ryzyka i werdykt"),
    (slide_steps, "Tok audytu", "Tok audytu"),
    (slide_conclusions, "Wnioski", "Wnioski z metody"),
    (slide_ex_offset, "Przykład 1 z 4", "Przykład: kompensacja"),
    (slide_ex_scope, "Przykład 2 z 4", "Przykład: zakres twierdzenia"),
    (slide_ex_compare, "Przykład 3 z 4", "Przykład: porównanie"),
    (slide_ex_three, "Przykład 4 z 4", "Przykład: znak, obowiązek, etyka"),
    (slide_recommendations, "Rekomendacje", "Rekomendacje"),
    (slide_close, "Nota", "Granice metody"),
]


def main():
    global ST, LOGO, LOGO_RATIO
    register_fonts()
    ST = styles()
    LOGO, LOGO_RATIO = prepare_logo()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=(W, H))
    c.setTitle("Audyt ryzyk greenwashingu. Metodyka, wnioski i rekomendacje")
    c.setAuthor("ESG Institute")
    c.setCreator("ESG Institute")
    c.setSubject(
        "Metodyka, wnioski i rekomendacje. 8 listopada 2026 r."
    )
    c.setKeywords(
        "audyt ryzyk greenwashingu, ESG Institute, EmpCo, komunikacja środowiskowa"
    )
    total = len(SLIDES)
    for i, (fn, section, outline) in enumerate(SLIDES, 1):
        if i > 1:
            c.showPage()
        fn(c, i, total, section, outline)
    c.save()
    print(f"Zapisano {OUT}")


if __name__ == "__main__":
    main()
