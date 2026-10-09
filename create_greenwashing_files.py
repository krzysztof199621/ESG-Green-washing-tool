"""
Skrypt do wygenerowania dwóch plików testowych z greenwashingiem dla Green Washing Tool:
1. Plik PDF: Sektor E-commerce, Logistyka kurierska i Moda (EcoExpress Logistics & GreenMarket).
2. Plik DOCX: Sektor Nieruchomości, Budownictwo i Finanse ESG (GreenHorizon Development & Aura Impact Fund).
"""

import os
import sys
from pathlib import Path

# Dodanie ścieżki do webapp, aby móc użyć engine do walidacji
BASE_DIR = Path("/Users/krzysztofsulinski/Desktop/AI Projects/ESG/Green Washing Tool")
sys.path.insert(0, str(BASE_DIR / "webapp"))

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# --- REJESTRACJA CZCIONEK SYSTEMOWYCH Z OBSŁUGĄ POLSKICH ZNAKÓW ---
pdfmetrics.registerFont(TTFont('Arial', '/System/Library/Fonts/Supplemental/Arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Bold', '/System/Library/Fonts/Supplemental/Arial Bold.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Italic', '/System/Library/Fonts/Supplemental/Arial Italic.ttf'))
pdfmetrics.registerFont(TTFont('Arial-BoldItalic', '/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf'))


class NumberedCanvas(canvas.Canvas):
    """Canvas z automatyczną numeracją stron i stopką."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_header_footer(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_header_footer(self, page_count):
        self.saveState()
        self.setFont("Arial", 8)
        self.setFillColor(colors.HexColor("#64748B"))
        
        # Nagłówek (od strony 2)
        if self._pageNumber > 1:
            self.drawString(40, 810, "EcoExpress Global & GreenMarket S.A. | Raport Zrównoważonego Rozwoju i Oferta 2026")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(40, 804, 555, 804)

        # Stopka
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.setLineWidth(0.5)
        self.line(40, 45, 555, 45)
        self.drawString(40, 32, "DOKUMENT POGLĄDOWY – MATERIAŁ ANALITYCZNY TESTUJĄCY ZGODNOŚĆ Z ECGT I UOKIK")
        self.drawRightString(555, 32, f"Strona {self._pageNumber} z {page_count}")
        self.restoreState()


def build_pdf_document(pdf_path: Path):
    """Tworzy profesjonalny plik PDF z błędami greenwashingu."""
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=50,
        bottomMargin=55
    )

    styles = getSampleStyleSheet()
    
    # Style niestandardowe
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Arial-Bold',
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#065F46"),
        spaceAfter=6
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Arial',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#475569"),
        spaceAfter=15
    )
    h1_style = ParagraphStyle(
        'H1',
        parent=styles['Normal'],
        fontName='Arial-Bold',
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#0F172A"),
        spaceBefore=14,
        spaceAfter=8,
        keepWithNext=True
    )
    h2_style = ParagraphStyle(
        'H2',
        parent=styles['Normal'],
        fontName='Arial-Bold',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#047857"),
        spaceBefore=10,
        spaceAfter=4,
        keepWithNext=True
    )
    body_style = ParagraphStyle(
        'Body',
        parent=styles['Normal'],
        fontName='Arial',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=6
    )
    bold_body = ParagraphStyle(
        'BoldBody',
        parent=body_style,
        fontName='Arial-Bold'
    )
    callout_style = ParagraphStyle(
        'Callout',
        parent=styles['Normal'],
        fontName='Arial',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#064E3B")
    )
    small_style = ParagraphStyle(
        'Small',
        parent=styles['Normal'],
        fontName='Arial-Italic',
        fontSize=7.5,
        leading=10,
        textColor=colors.HexColor("#64748B")
    )
    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Arial-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#FFFFFF")
    )
    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Arial',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#1E293B")
    )

    story = []

    # --- NAGŁÓWEK TYTUŁOWY ---
    badge_table = Table(
        [[Paragraph("<font color='#047857'><b>KAMPANIA „ZIELONA PRZYSZŁOŚĆ 2026”</b></font> | STRATEGIA ZRÓWNOWAŻONEGO ROZWOJU", ParagraphStyle('B', fontName='Arial-Bold', fontSize=8, leading=10, textColor=colors.HexColor("#047857")))]],
        colWidths=[515]
    )
    badge_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#ECFDF5")),
        ('PADDING', (0,0), (-1,-1), 6),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#A7F3D0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(badge_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph("EcoExpress Global & GreenMarket S.A.", title_style))
    story.append(Paragraph("<b>Kompleksowy Raport Środowiskowy i Nowa Oferta E-Commerce: Zielona Dostawa, Bezemisyjna Flota i Zrównoważona Kolekcja Odzieży</b>", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#059669"), spaceBefore=0, spaceAfter=12))

    # --- SEKCJA 1: WPROWADZENIE I OGÓLNE HASŁA EKOLOGICZNE (ECGT 4a, 4b) ---
    story.append(Paragraph("1. Nasza Misja: Ekologiczny i Przyjazny Środowisku E-Commerce", h1_style))
    story.append(Paragraph(
        "Z dumą ogłaszamy, że od 2026 roku <b>nasza firma jest w 100% ekologiczna</b>. Wdrażamy kompleksowe rozwiązania zorientowane na "
        "klimat, oferując konsumentom w całym kraju rewolucyjną usługę: <b>zielona dostawa</b> do każdego domu oraz w pełni <b>zeroemisyjny e-commerce</b>. "
        "Jesteśmy dumni, że jako lider rynku staliśmy się podmiotem całkowicie <b>climate friendly</b> oraz <b>eco-friendly</b>, tworząc przyjazne środowisku "
        "doświadczenie zakupowe na każdym etapie obsługi.",
        body_style
    ))

    # Callout z hasłem (ECGT 4a)
    callout_1 = Table(
        [[Paragraph("🌿 <b>Hasło Przewodnie Kampanii:</b> Wybierz świadomie! Zrównoważona przyszłość i zielona dostawa – u nas każdy zakup i każda przesyłka są z definicji eko i przyjazne dla planety.", callout_style)]],
        colWidths=[515]
    )
    callout_1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F0FDF4")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#86EFAC")),
        ('PADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(callout_1)
    story.append(Spacer(1, 10))

    # --- SEKCJA 2: NEUTRALNOŚĆ KLIMATYCZNA I OFFSETÓW (ECGT 4c, UOKiK, COUNTEMISSIONS) ---
    story.append(Paragraph("2. Przesyłka Climate Neutral i Sadzenie Drzew: Bezwzględna Neutralność GHG", h1_style))
    story.append(Paragraph(
        "Każda przesyłka nadana w naszej sieci posiada gwarantowany certyfikat: <b>przesyłka climate neutral thanks to offsets</b>! "
        "W ramach autorskiego programu ekologicznego <b>sadzimy drzewa za każde zamówienie</b>, a nieuniknione emisje gazów cieplarnianych "
        "związane z transportem i pakowaniem są w całości kompensowane przez certyfikowane <b>kredyty węglowe (carbon credits)</b> oraz <b>offsety węglowe</b> "
        "pochodzące z projektów ochrony lasów w dorzeczu Amazonki. W ten sposób oferujemy pełną <b>neutralność klimatyczną</b> oraz paczki w 100% <b>węglowo neutralne</b>.",
        body_style
    ))
    story.append(Paragraph(
        "Według naszych wewnętrznych wyliczeń, <b>zredukowaliśmy emisję o 4,8 kg CO2 na każdej przesyłce</b>, osiągając status transportu bezemisyjnego w standardzie net-zero.",
        body_style
    ))

    # Ramka ostrzegawcza w stylu dokumentu z QR kodem (SAME-MEDIUM-SPEC)
    callout_2 = Table(
        [[Paragraph("ℹ️ <b>Weryfikacja i Dane:</b> Szczegóły na www oraz po zeskanowaniu dedykowanego kodu QR na etykiecie przesyłki. Pełna metodologia offsetów i kompensacji CO₂ dostępna online.", callout_style)]],
        colWidths=[515]
    )
    callout_2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EFF6FF")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#93C5FD")),
        ('PADDING', (0,0), (-1,-1), 7),
    ]))
    story.append(callout_2)
    story.append(Spacer(1, 10))

    # --- SEKCJA 3: FLOTA I ENERGIA ODNAWIALNA (ECGT 4b, EAC-NOT-PHYSICAL-ENERGY / UOKIK BOLT) ---
    story.append(Paragraph("3. Zeroemisyjna Flota i Zielona Energia w Centrach Logistycznych", h1_style))
    story.append(Paragraph(
        "Realizując plan dekarbonizacji, z dumą ogłaszamy, że cała nasza <b>zielona flota</b> to nowoczesny transport miejski. "
        "Nasza <b>zeroemisyjna flota</b> dociera do tysięcy klientów dziennie. Flota: zakupiliśmy <b>jeden van elektryczny na odcinek last mile</b> w ścisłym centrum Warszawy "
        "(podczas gdy pozostałe 940 kurierskich samochodów dostawczych to sprawdzona flota spalinowa z normą Euro 5/6).",
        body_style
    ))
    story.append(Paragraph(
        "Dodatkowo nasze magazyny i wielkie huby przeładunkowe są zasilane w <b>100% energii odnawialnej</b>! Osiągnęliśmy to poprzez zakup <b>certyfikatów EAC (Energy Attribute Certificates)</b> "
        "oraz międzynarodowych gwarancji pochodzenia. Dzięki temu cała nasza infrastruktura magazynowa jest w 100% OZE i wolna od emisyjnej energii.",
        body_style
    ))
    story.append(Spacer(1, 6))

    # --- SEKCJA 4: KOLEKCJA MODOWA, OPAKOWANIA I PRYWATNE ZNAKI (ECGT 4b, 2a, 10a, 23f) ---
    story.append(Paragraph("4. Kolekcja GreenWear Join Life i Standard Opakowań", h1_style))
    story.append(Paragraph(
        "Dla miłośników mody wprowadziliśmy na rynek prestiżową kolekcję <b>Join Life Conscious Collection</b>: to nasza w 100% <b>zrównoważona kolekcja</b> z odzieżą przyszłości. "
        "Wszystkie ubrania z linii EcoWear produkowane są w oparciu o materiały z odzysku – linia zawiera <b>8% recyklatu</b> tworzyw sztucznych w mieszance włókien. "
        "Dodatkowo deklarujemy <b>zero waste we własnych obiektach</b> oraz w całym łańcuchu handlowym (program realizowany pilotażowo w 1 z 45 sklepów stacjonarnych).",
        body_style
    ))
    story.append(Paragraph(
        "Dbając o zaufanie konsumentów, nasze wyroby odznaczyliśmy specjalnym symbolem: <b>nasz własny znak Eco Certified Planet Hero</b> gwarantuje najwyższe standardy etyczne i środowiskowe! "
        "Znak ten został ustanowiony przez naszą wewnętrzną radę ds. ekologii i jest gwarantem zielonej jakości.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Wybierz odpowiedzialną modę:</b> Nasz T-shirt jest całkowicie <b>wolny od substancji zakazanych w UE oraz bez substancji zakazanych w UE</b> i spełnia wszystkie wymogi obowiązującego prawa chemicznego w UE! "
        "To nasza unikalna przewaga konkurencyjna i dowód troski o zdrowie konsumentów.",
        body_style
    ))
    story.append(Paragraph(
        "Ponadto nasze firmowe ekotorby i opakowania kurierskie są <b>niezniszczalne i trwają na lata</b>, co czyni je wiecznym rozwiązaniem dla obiegu zamkniętego.",
        body_style
    ))
    story.append(Spacer(1, 6))

    # --- SEKCJA 5: FMCG, ŻYWNOŚĆ I KAWA (ORG-848-30, EUDR) ---
    story.append(Paragraph("5. Strefa Kawiarniana i Zdrowa Żywność: Kawa Eko Bio Organic", h1_style))
    story.append(Paragraph(
        "W każdym naszym punkcie odbioru paczek klienci mogą spróbować wyjątkowych produktów spożywczych: serwujemy <b>eko kawę bio organic</b> oraz przekąski prosto z plantacji "
        "<b>rainforest friendly</b> i w 100% <b>zero deforestation (bez wylesiania)</b>. Nasza kawa to czysta natura, uprawiana bez grama chemii w sercu Ameryki Południowej "
        "(produkt pakowany bez urzędowego unijnego certyfikatu rolnictwa ekologicznego ani procedury due diligence EUDR).",
        body_style
    ))
    story.append(Spacer(1, 8))

    # --- TABELA PODSUMOWUJĄCA OŚWIADCZENIA I STATUS ---
    story.append(Paragraph("Tabela Deklaracji Środowiskowych Kampanii 2026", h2_style))
    table_data = [
        [
            Paragraph("Obszar", table_header_style),
            Paragraph("Oficjalne Twierdzenie Marketingowe", table_header_style),
            Paragraph("Zadeklarowany Mechanizm / Zakres", table_header_style),
        ],
        [
            Paragraph("Dostawy kurierskie", table_cell_style),
            Paragraph("<b>Zielona dostawa</b>, climate neutral, zeroemisyjny e-commerce", table_cell_style),
            Paragraph("Kompensacja emisji: sadzimy drzewa za paczkę + certyfikaty offsetowe", table_cell_style),
        ],
        [
            Paragraph("Flota drogowa", table_cell_style),
            Paragraph("<b>Zeroemisyjna flota</b> samochodów kurierskich", table_cell_style),
            Paragraph("1 van elektryczny na centrum Warszawy (ostatnia mila); 940 vanów spalinowych", table_cell_style),
        ],
        [
            Paragraph("Magazyny i Huby", table_cell_style),
            Paragraph("<b>100% energii odnawialnej (OZE)</b> w obiektach", table_cell_style),
            Paragraph("Zakup certyfikatów atrybutu EAC (gwarancje pochodzenia) na rynku", table_cell_style),
        ],
        [
            Paragraph("Kolekcja ubrań", table_cell_style),
            Paragraph("<b>Zrównoważona kolekcja</b> Join Life, eko odzież", table_cell_style),
            Paragraph("8% recyklatu poliestru; własne logo Eco Certified bez audytu zewnętrznego", table_cell_style),
        ],
        [
            Paragraph("Żywność i Kawa", table_cell_style),
            Paragraph("<b>Eko kawa bio organic</b>, rainforest friendly, zero deforestation", table_cell_style),
            Paragraph("Deklaracja marketingowa; szczegóły i warunki pod kodem QR i na stronie www", table_cell_style),
        ],
    ]
    summary_table = Table(table_data, colWidths=[110, 210, 195])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#065F46")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#F8FAFC"), colors.HexColor("#FFFFFF")]),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(summary_table)

    story.append(Spacer(1, 14))
    story.append(Paragraph(
        "<i>Dokument sporządzony przez Dział Marketingu i Relacji Zrównoważonych EcoExpress Global S.A. "
        "Wszystkie hasła podlegają ocenie pod kątem Dyrektywy ECGT (UE 2024/825), UPNPR oraz UOKiK.</i>",
        small_style
    ))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Wygenerowano PDF: {pdf_path}")


def build_docx_document(docx_path: Path):
    """Tworzy profesjonalny plik DOCX z innym zestawem twierdzeń greenwashingu (Nieruchomości, Budownictwo, Finanse SFDR)."""
    doc = docx.Document()

    # Ustawienie marginesów strony (2.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # Kolory w palecie biznesowo-ekologicznej
    PRIMARY_COLOR = RGBColor(12, 74, 96)     # Deep Teal #0C4A60
    SECONDARY_COLOR = RGBColor(16, 124, 65)  # Forest Green #107C41
    DARK_TEXT = RGBColor(30, 41, 59)         # Slate 800
    MUTED_TEXT = RGBColor(100, 116, 139)     # Slate 500

    # Tytuł
    p_badge = doc.add_paragraph()
    r_badge = p_badge.add_run("PROSPEKT INWESTYCYJNY & OFERTA ZRÓWNOWAŻONEGO BUDOWNICTWA 2026")
    r_badge.font.name = 'Arial'
    r_badge.font.size = Pt(8.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = SECONDARY_COLOR

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("GreenHorizon Development & Aura Impact Property Fund")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = PRIMARY_COLOR

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Inwestycja „Zielona Oaza Residence & Business Park”: Ekologiczny i Bezemisyjny Kompleks Przyszłości Finansowany przez Fundusz ESG Impact")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = MUTED_TEXT

    # Funkcja pomocnicza do dodawania nagłówków
    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = PRIMARY_COLOR
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(10.5)
        run.font.bold = True
        run.font.color.rgb = SECONDARY_COLOR
        return p

    def add_p(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Arial'
            rb.font.size = Pt(9.5)
            rb.font.bold = True
            rb.font.color.rgb = DARK_TEXT
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(9.5)
        r.font.color.rgb = DARK_TEXT
        return p

    def add_callout(text, bg_hex="F0FDF4", border_hex="86EFAC"):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        cell = table.cell(0, 0)
        cell.width = Inches(6.5)
        
        # XML dla koloru tła i obramowania
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>')
        tcPr.append(shd)
        
        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="8" w:space="0" w:color="{border_hex}"/>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>'
            f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{border_hex}"/>'
            f'<w:right w:val="single" w:sz="8" w:space="0" w:color="{border_hex}"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(4)
        cp.paragraph_format.space_after = Pt(4)
        crun = cp.add_run(text)
        crun.font.name = 'Arial'
        crun.font.size = Pt(9)
        crun.font.color.rgb = DARK_TEXT
        
        # Pusty odstęp po tabeli
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(6)

    # --- ROZDZIAŁ 1: WPROWADZENIE I OGÓLNE HASŁA EKOLOGICZNE (ECGT 4a) ---
    add_h1("1. O Inwestycji: Ekologiczna Przestrzeń Nowej Generacji")
    add_p(
        "Z dumą prezentujemy najbardziej zielony projekt deweloperski na polskim rynku. Nasza inwestycja „Zielona Oaza Residence & Business Park” "
        "jest w pełni ekologiczna, całkowicie zielona oraz stworzona w 100% w zgodzie z naturą. Oferujemy przyszłym mieszkańcom oraz najemcom "
        "przyjazny środowisku apartamentowiec oraz zrównoważone przestrzenie biurowe, które w każdym aspekcie są climate-friendly oraz eco-friendly. "
        "To wybór dla każdego, komu bliski jest zrównoważony rozwój i troska o planetę."
    )
    add_callout("🌿 Misja Dewelopera: Zielona Oaza to synonim zrównoważenia. Wybierając nasze lokale, wspierasz w 100% zrównoważony styl życia w ekologicznym sercu nowoczesnego miasta.")

    # --- ROZDZIAŁ 2: BEZEMISYJNOŚĆ I PRZENOSZENIE CERTYFIKACJI BREEAM / LEED (NO-CERT-TRANSFER / EPBD ART. 11) ---
    add_h1("2. Certyfikacja Budynku: W 100% Bezemisyjny Budynek Przyszłości")
    add_p(
        "Nasza inwestycja to w 100% bezemisyjny budynek oraz w pełni zeroemisyjna inwestycja budowlana. Fakt ten jednoznacznie potwierdza uzyskany przez "
        "nas prestiżowy certyfikat BREEAM NC v7.1 na poziomie Excellent oraz precertyfikacja LEED v4.1 BD+C Gold! Otrzymanie tak wysokiej punktacji BREEAM i LEED "
        "stanowi bezsporny dowód, że obiekt osiągnął pełną bezemisyjność operacyjną i wbudowaną, wyprzedzając wymogi unijne."
    )
    add_p(
        "Warto podkreślić, że w ramach zabezpieczenia ciągłości ogrzewania w budynku zainstalowano nowoczesne kotły gazowe jako zasilanie szczytowe, "
        "jednak dzięki doskonałym parametrom izolacji certyfikat BREEAM i LEED gwarantuje, że budynek kwalifikuje się jako całkowicie zeroemisyjny."
    )

    # --- ROZDZIAŁ 3: KOMPENSACJA, BETON, 100% OZE I FLOTA (ECGT 4c, 4b, EAC) ---
    add_h1("3. Zrównoważony Beton, Net-Zero Construction i Energia Odnawialna")
    add_p(
        "Proces powstawania Zielonej Oazy to rewolucja w inżynierii lądowej. Do budowy fundamentów i konstrukcji wykorzystaliśmy zrównoważony beton eko-friendly, "
        "w którym aż 7% kruszywa naturalnego zastąpiono recyklatem ze skruszonego gruzu porozbiórkowego."
    )
    add_p(
        "Ponadto sam proces budowy został zrealizowany jako całkowicie neutralny dla klimatu: z dumą ogłaszamy net-zero construction! "
        "Wszelkie emisje GHG wynikające ze spalania paliw maszyn budowlanych i produkcji stali skompensowaliśmy poprzez międzynarodowe certyfikaty i offsety węglowe, "
        "a w ramach akcji Zielony Ślad sadzimy drzewa za każdy sprzedany metr kwadratowy lokalu. W ten sposób inwestycja osiągnęła pełną neutralność klimatyczną!"
    )
    add_p(
        "Wszystkie części wspólne, windy oraz oświetlenie zewnętrzne są zasilane w 100% energii odnawialnej (100% OZE). Osiągnęliśmy to poprzez zakup "
        "certyfikatów EAC (Energy Attribute Certificates) oraz gwarancji pochodzenia z elektrowni wodnych. Bez konieczności instalowania kosztownych i skomplikowanych "
        "paneli fotowoltaicznych na dachu, nasz obiekt funkcjonuje wyłącznie w oparciu o czystą, zieloną energię."
    )
    add_p(
        "Również transport budowlany opiera się na idei czystej mobilności: stworzyliśmy zieloną flotę logistyczną – nasza bezemisyjna flota budowlana obejmuje "
        "jeden elektryczny wózek widłowy na placu magazynowym (podczas gdy 150 ciężarówek i betoniarek to tradycyjne pojazdy spalinowe z silnikiem diesla). "
        "Wprowadziliśmy także zero waste we własnych obiektach – całkowicie bezodpadowe biuro dewelopera i plac inwestycyjny (w biurze sprzedaży postawiono kosz do segregacji papieru)."
    )

    # --- ROZDZIAŁ 4: FINANSOWANIE I FUNDUSZ ESG IMPACT (SFDR-8-VS-9 & ESMA) ---
    add_h1("4. Finansowanie: Fundusz Aura ESG Impact Property Fund")
    add_p(
        "Głównym partnerem finansowym inwestycji jest renomowany fundusz „Aura ESG Impact Property Fund”. Fundusz został formalnie zaklasyfikowany "
        "pod artykuł 8 SFDR (Rozporządzenie o ujawnianiu informacji związanych ze zrównoważonym rozwojem w sektorze usług finansowych). "
        "Jednocześnie fundusz definiuje jako swój główny cel bezpośrednie inwestycje zrównoważone oraz generowanie mierzalnego, pozytywnego wpływu klimatycznego "
        "(ESG impact fund / cel zrównoważony w rozumieniu artykułu 9)."
    )
    add_p(
        "Portfel funduszu inwestuje 55% kapitału w nieruchomości spełniające kryteria środowiskowe, co pozwala nam promować fundusz jako pełnoprawny, "
        "wiodący na rynku zielony fundusz klimatyczny ESG Impact, generujący maksymalną stopę zwrotu przy ratowaniu klimatu."
    )

    # --- ROZDZIAŁ 5: WYMOGI PRAWNE, WŁASNY ZNAK, TRWAŁOŚĆ I KAWA (ECGT 2a, 10a, 23f, ORG-848-30, EUDR) ---
    add_h1("5. Standardy Bezpieczeństwa, Własna Certyfikacja i Gastronomia")
    add_p(
        "Kupując apartament w Zielonej Oazie, zyskujesz pewność i spokój. Nasze lokale spełniają wszystkie wymogi prawa budowlanego oraz warunki techniczne WT "
        "wymagane w UE i są całkowicie zgodne z obowiązującym prawem budowlanym oraz przepisami środowiskowymi – przedstawiamy to jako naszą unikalną przewagę ekologiczną i korzyść rynkową!"
    )
    add_p(
        "Inwestycja została także wyróżniona prestiżowym odznaczeniem: nasz własny znak „Green Eco-Certified Residence” stanowi gwarancję najwyższego luksusu "
        "w zgodzie z ekologią. Znak ten został nadany przez wewnętrzną komisję oceny jakości naszej spółki deweloperskiej."
    )
    add_p(
        "Zastosowane w lokalach okna trzyszybowe i drzwi są absolutnie niezniszczalne i trwają na lata, zapewniając wieczną bezobsługowość i odporność na wszelkie czynniki atmosferyczne."
    )
    add_p(
        "Dla wygody mieszkańców w lobby apartamentowca prowadzimy ekskluzywny punkt kawiarniany: serwujemy w nim mieszkańcom pyszną eko kawę i bio herbatę organic prosto z plantacji "
        "rainforest friendly, w 100% zero deforestation (bez wylesiania) ze słonecznej Kolumbii (kawa serwowana jest na podstawie deklaracji dostawcy, bez urzędowego certyfikatu "
        "rolnictwa ekologicznego UE 2018/848 ani oświadczeń due diligence EUDR)."
    )

    # --- ROZDZIAŁ 6: WERYFIKACJA I QR KOD (SAME-MEDIUM-SPEC) ---
    add_h1("6. Dostęp do Szczegółowych Danych i Raportów")
    add_p(
        "Wszystkie szczegółowe dane, certyfikaty, analizy środowiskowe oraz raporty zrównoważonego rozwoju są dostępne na stronie www dewelopera oraz pod kodem QR "
        "umieszczonym na banerze przed wejściem na teren budowy. Zeskanuj kod QR i poznaj szczegóły naszej zielonej transformacji!"
    )
    add_callout(
        "📲 Szczegóły na www oraz zeskanuj kod QR: Pełna dokumentacja deklaracji ekologicznych znajduje się pod dedykowanym adresem internetowym.",
        bg_hex="EFF6FF",
        border_hex="93C5FD"
    )

    # Tabela zbiorcza
    add_h2("Zestawienie Parametrów Środowiskowych Inwestycji Zielona Oaza")
    table = doc.add_table(rows=7, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    headers = ["Kategoria", "Oświadczenie w Prospekcie", "Rzeczywisty Zakres i Podstawa"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.width = Inches(2.1)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="0C4A60"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = 'Arial'
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    rows_data = [
        ("Efektywność budynku", "Bezemisyjny budynek, zeroemisyjny obiekt", "Certyfikat BREEAM v7.1 i LEED BD+C (zainstalowane kotły gazowe)"),
        ("Materiały budowlane", "Zrównoważony beton eko-friendly", "7% kruszywa z recyklingu gruzu budowlanego w mieszance"),
        ("Ślad węglowy budowy", "Net-zero construction, neutralność klimatyczna", "Zakup certyfikatów kompensacji CO2 (offsety) i sadzenie drzew"),
        ("Zasilanie energią", "100% energii odnawialnej (OZE)", "Certyfikaty atrybutu EAC (gwarancje pochodzenia), brak PV"),
        ("Finansowanie", "Fundusz pod artykuł 8 SFDR z celem zrównoważonym / ESG Impact", "Promocja art. 8 językiem art. 9; 55% alokacji (poniżej 80% ESMA)"),
        ("Znak jakości", "Green Eco-Certified Residence", "Własny znak marketingowy dewelopera bez certyfikacji 3rd party"),
    ]

    for row_idx, data in enumerate(rows_data, start=1):
        bg_color = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(data):
            cell = table.cell(row_idx, col_idx)
            cell.width = Inches(2.1)
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            cell._tc.get_or_add_tcPr().append(shd)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(text)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = DARK_TEXT

    # Stopka
    doc.save(str(docx_path))
    print(f"Wygenerowano DOCX: {docx_path}")


if __name__ == "__main__":
    pdf_out = BASE_DIR / "Przyklad_Greenwashing_E-commerce_Logistyka.pdf"
    docx_out = BASE_DIR / "Przyklad_Greenwashing_Nieruchomosci_i_Finanse.docx"

    build_pdf_document(pdf_out)
    build_docx_document(docx_out)
    print("Zakończono generowanie plików.")
