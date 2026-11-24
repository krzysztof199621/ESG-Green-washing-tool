#!/usr/bin/env python3
"""Generator materiału wejściowego oraz raportów (pełnego i skróconego) dla MOWI ASA (https://mowi.com/).

Zgodny z INSTRUKCJA_GENEROWANIA_RAPORTOW.md, .cursor/rules/greenwashing-raport.mdc,
składem pełnym ReportDoc (sklad_raportu.py) oraz składem raportu skróconego (generuj_raport_skrocony.py).
Stan witryny i prawa: 24 listopada 2026 r.
"""

from io import BytesIO
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import docx
from PIL import Image
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

from sklad_raportu import (
    DARK,
    MUTED,
    NAVY,
    VERDICT_RED,
    WHITE,
    ReportDoc,
    default_logo,
    set_run,
    set_table_width,
)

MOWI_DIR = Path(__file__).resolve().parent
SCREENSHOT = MOWI_DIR / "zrzut_mowi_com.png"
AUDIT_DATE = "24 listopada 2026 r."
AUDIT_REF = "Sygn. ESG/2026/11/24"

OUT_MATERIAL = MOWI_DIR / "Material_MOWI_com_tresci_do_audytu.docx"
OUT_FULL = MOWI_DIR / "Raport_Audytu_Greenwashingu_MOWI_com.docx"
OUT_SHORT = MOWI_DIR / "Raport_Skrocony_Audytu_Greenwashingu_MOWI_com.docx"


def clean_screenshot_stream(crop_bottom=810):
    """Zwraca kadr strony głównej bez dolnego paska cookies jako strumień PNG."""
    image = Image.open(SCREENSHOT).convert("RGB")
    w, h = image.size
    crop = image.crop((0, 0, w, min(h, crop_bottom)))
    stream = BytesIO()
    crop.save(stream, format="PNG", optimize=True)
    stream.seek(0)
    return stream


# ==============================================================================
# 1. MATERIAŁ WEJŚCIOWY (wzór: wzory/WZOR_material_Libra_Partners.docx)
# ==============================================================================


def build_material():
    """Tworzy uporządkowany zbiór dosłownych treści z witryny https://mowi.com/ (bez oceny prawnej)."""
    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    def add_p(text, size=10.5, bold=False, italic=False, color=DARK, after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = 1.15
        set_run(p.add_run(text), size, bold, italic, color)
        return p

    def add_h1(text):
        p = doc.add_paragraph(style="Heading 1")
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        for r in p.runs:
            set_run(r, 12.5, True, False, NAVY)
        if not p.runs:
            set_run(p.add_run(text), 12.5, True, False, NAVY)
        else:
            p.runs[0].text = text
        return p

    def add_bullet(text):
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        set_run(p.add_run(text), 10, False, False, DARK)
        return p

    add_p("MOWI ASA – Zestaw Treści i Deklaracji Środowiskowych oraz Żywieniowych", size=15, bold=True, color=NAVY, after=2)
    add_p("Kompleksowy zbiór tekstów marketingowych, specyfikacji i twierdzeń z witryny https://mowi.com/ do audytu w Green Washing Tool", size=10.5, italic=True, color=MUTED, after=10)

    meta_lines = [
        "Podmiot: Mowi ASA, Sandviksboder 77AB, 5035 Bergen, Norwegia (adres korespondencyjny: PO Box 4102 Sandviken, 5835 Bergen, Norwegia)",
        "Oficjalna witryna: https://mowi.com/",
        "Podstrony objęte zbiorem: /about-us/, /about-us/our-strategy/, /about-us/our-structure/mowi-feed/ (/sourcing/, /strategy-and-ambitions/), /products/ (/salmon-farming/, /taste-health/, /quality/, /food-safety/, /traceability/, /mowi-4-0-smart-farming/), /sustainability/ (/commitments/, /certifications/, /the-sustainable-development-goals-sdgs/, /how-mowi-uses-freshwater-responsibly/, /key-partnerships/, /policies/), /knowledge/does-salmon-farming-have-an-impact-on-the-seabed/, /contact/",
        "Data zebrania materiałów: 2026-11-24",
        "Przeznaczenie pliku: Bezpośredni wsad źródłowy do analizy greenwashingu. Zawiera dosłowne brzmienie cytatów bez oceny prawnej.",
    ]
    for line in meta_lines:
        add_p(line, size=9.5, after=3)

    sections = [
        (
            "1. STRONA GŁÓWNA I O NAS – https://mowi.com/ oraz https://mowi.com/about-us/",
            [
                "At Mowi we provide our customers with 9.7 million meals a day",
                "We are the most sustainable protein producer. Mowi has been ranked the most sustainable protein producer by the Coller FAIRR Protein Producer Index for the fifth year in a row. *MOWI ranked #1 in the 2019-2024 Coller FAIRR Protein Producer Index, a comprehensive assessment of the largest animal protein producers on critical environmental, social and governance (ESG) issues.",
                "Mowi, the world’s largest supplier of farm-raised salmon and most sustainable protein producer",
                "Mowi ASA is one of the largest seafood companies in the world, and the world’s largest producer of Atlantic salmon with harvest volumes of 559 000 tonnes in 2025, equivalent to a global market share of approximately 20%.",
                "In 2025 the company had a turnover of EUR 5.7 billion. Mowi is listed on the Oslo Stock Exchange (OSE) and its share also trade on the US OTC market.",
                "Dane liczbowe na stronie głównej i /about-us/: 12,500 Employees (na dole strony /about-us/ w sekcji Global Presence: employ 11 500 people; Asia: 7 processing facilities na stronie głównej vs 6 processing facilities na /about-us/).",
                "Cytat CEO (Ivan Vindheim): „Food security and climate change are two of the most pressing challenges facing humanity. As a seafood producer, Mowi is unlocking the potential of the ocean to produce healthy and climate-friendly food for a growing world population.”",
            ],
        ),
        (
            "2. PRODUKTY, ZDROWIE I HODOWLA – https://mowi.com/products/ i podstrony",
            [
                "Our goal is to deliver top-quality salmon and inspire a healthy and climate-friendly lifestyle",
                "Salmon is a consumer favourite, a superfood, highly nutritious and promoted by nutritionists worldwide",
                "The Blue Foods Assessment* found that blue foods from the ocean, such as salmon, on average have much greater nutritional benefits than terrestrial animal foods, and many also have a smaller environmental footprint. * https://bluefood.earth/",
                "Salmon farming is climate friendly food production.",
                "Blue or Aquatic Foods have lower greenhouse gas (GHG) emissions than land-based foods and, therefore, salmon farming can help bridge the gap between the potential of the ocean and the production of sustainable food that is healthy for people.",
                "A true superfood (/products/taste-health/)",
                "Marine Omega-3 prevent development of cardio-vascular disease.",
                "Omega-3 fatty acids help to maintain healthy heart by lowering blood pressure, as well as reduce the risk of sudden death, heart attack, stroke 1,2,3 , and coronary heart disease 3,4",
                "Omega-3 fatty acids support brain function and development in infants. They also possibly prevent psychiatric diseases, particularly cognitive decline in the elderly 5 .",
                "Omega-3 fatty acids possibly prevent inflammation and reduce the risk of arthritis 5,6 .",
                "Przypis 4 na /taste-health/: U.S. Food and Drug Administration. Summary of qualified health claims subject to enforcement discretion. 2014.",
                "Tabela wartości odżywczych (/taste-health/): Total fat 16,0g / 100g (Mowi Salmon) vs 8,0g / 100g (Wild atlantic salmon); Total EPA + DHA 1,3g / 100g (539% RDI) vs 1,6g / 100g (640% RDI); Vitamin B12 3,8µg / 100g vs 6,8µg / 100g; Vitamin D 6,5µg / 100g vs 11,0µg / 100g.",
                "Salmon feed contains natural compounds like astaxanthin, which give salmon their pink hue.",
            ],
        ),
        (
            "3. ZRÓWNOWAŻONY ROZWÓJ, CERTYFIKATY I CELE – https://mowi.com/sustainability/ i podstrony",
            [
                "Blue foods have lower freshwater use and land use compared to terrestrial meats. *BFA, 2021",
                "Blue foods have lower GHG emissions than land-based foods (BFA, 2021). Dietary shifts towards increased seafood consumption is recognised as part of the solution to climate change.",
                "Increased consumption of blue foods may reduce the consumption of terrestrial meats, consequently reducing diet-related chronic disease like hypertension, obesity and certain types of cancer. *BFA, 2021",
                "Kafelki celów na /sustainability/: 50,6% reduction in scope 1 and 2 emissions and 27,5% reduction in scope 3 emissions by 2030*; 10% reduction on water intensity at our processing plants by 2030*; waste to landfill at our direct operations by 2030* (bez cyfry „0”); 0% of sites above national limit (monthly average) [Sea lice — bez widocznego „0%” w tekście kafelka na głównej siatce]; 100% of harvest volumes sustainably certified by a GSSI-recognised standard (bez słowa „target”).",
                "Podstrona /commitments/climate-change/: Mowi commits to reduce absolute scope 1 and 2 GHG emissions 50.6% by 2030 from a 2019 base year; reduce absolute scope 3 GHG emissions 27.5% within the same timeframe; reduce absolute scope 3 FLAG GHG emissions 33.33% within the same timeframe.",
                "Podstrona /commitments/plastics/: By 2030, 100% of our EU plastic packaging adhering to PPWR (Recyclability: recycling performance grade >70%; Recycled content: contact sensitive PET >30%; contact sensitivee other plastic >10%; other plastic packaging >35%).",
                "Podstrona /certifications/: Mowi continues to be a leader on ASC certification, having by the end of 2021 a total of 133 farming sites certified worldwide. This represents 50% of all our farming sites and accounts for 33% of all the ASC certified Atlantic salmon sites worldwide...",
                "Podstrona /key-partnerships/: nagłówek „Global Food Safety Initiative” nad akapitem opisującym „Global Sustainable Seafood Initiative”.",
            ],
        ),
        (
            "4. PASZE, WODA SŁODKA, DNO MORSKIE I SDG – podstrony szczegółowe",
            [
                "Artykuł z 23.11.2026 r. (/sustainability/how-mowi-uses-freshwater-responsibly/): „In fact, salmon production has the lowest water consumption per kilogram of edible meat compared with other sources of farmed animal protein.”",
                "Artykuł z 23.11.2026 r. (/sustainability/how-mowi-uses-freshwater-responsibly/): „For example, all of our soy is from organic-certified sources, with standards that include water management and good agricultural practices.”",
                "Podstrona /commitments/sustainable-feed/: „100% of soy and palm oil is certified (Proterra, RSPO or equivalent)” oraz „KPIs: % Soy Proterra, Palm oil RSPO or eqivalent certified”.",
                "Podstrona /about-us/our-structure/mowi-feed/: „We sell fish feeds in Norway, Scotland, Ireland, Iceland and the Faroe Islands for production according to conventional standards.”",
                "Podstrona /about-us/our-structure/mowi-feed/sourcing/: „The soy used in our feed is 100% deforestation-free.” oraz „We are working towards the 2023 FAO global roadmap to achieve the SDG 2 without breaching the 1.5° C threshold.”",
                "Podstrona /sustainability/: „Our commitments are set to make our business future-proof and are aligned with the UN Sustainable Development Goals (SDGs).”",
                "Podstrona /sustainability/the-sustainable-development-goals-sdgs/: wykazuje „0 programs” przy SDG 2 (Zero hunger), SDG 6 (Clean water and sanitation) oraz SDG 7 (Affordable and clean energy).",
                "Podstrona /commitments/ethical-business-conduct/: „No breach of human rights principles or policy identified”.",
                "Artykuł /knowledge/does-salmon-farming-have-an-impact-on-the-seabed/: „Mowi adheres to ongoing rigorous operations to ensure no negative impacts.” oraz „By farming within the environment’s capacity, we know we leave no permanent impact.” (w kontekście 3-letniego badania po zamknięciu farmy Isle of Ewe / Loch Ewe w Szkocji).",
            ],
        ),
    ]

    for title, items in sections:
        add_h1(title)
        for item in items:
            add_bullet(item)

    OUT_MATERIAL.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT_MATERIAL))
    return OUT_MATERIAL


# ==============================================================================
# 2. DANE AUDYTOWE (WSPÓLNA BAZA DLA RAPORTU PEŁNEGO I SKRÓCONEGO)
# ==============================================================================

CASES_FULL = [
    {
        "no": "01",
        "title": "Niedozwolone oświadczenia o zapobieganiu chorobom i zmniejszaniu ryzyka chorób oraz nieswoiste określenie „superfood”",
        "short_claim": "Oświadczenia o zapobieganiu chorobom układu krążenia, zawałowi, udarowi, chorobom psychicznym i nowotworom oraz hasło „superfood”",
        "short_basis": "Rozp. (UE) 1169/2011 art. 7 ust. 3–4; rozp. (WE) 1924/2006 art. 5, 6, 10, 14; rozp. (UE) 432/2012; UPNPR art. 5; UZNK art. 16",
        "table_basis": "Rozp. (UE) nr 1169/2011 art. 7 ust. 3–4; rozp. (WE) nr 1924/2006 art. 5, 6, 10 i 14 ust. 1 lit. a; rozp. (UE) nr 432/2012; UPNPR art. 5; UZNK art. 16",
        "case_basis": "Rozporządzenie (UE) nr 1169/2011 art. 7 ust. 3 i 4; rozporządzenie (WE) nr 1924/2006 art. 5, 6, 10 ust. 1–3 i art. 14 ust. 1 lit. a w zw. z rozporządzeniem Komisji (UE) nr 432/2012; posiłkowo art. 5 ust. 1 i ust. 2 pkt 1 ustawy z 23.08.2007 r. o przeciwdziałaniu nieuczciwym praktykom rynkowym (Dz.U. z 2023 r. poz. 845), art. 24 ustawy o ochronie konkurencji i konsumentów (Dz.U. z 2025 r. poz. 1714), art. 16 ust. 1 pkt 1–2 ustawy o zwalczaniu nieuczciwej konkurencji (Dz.U. z 2026 r. poz. 85)",
        "risk": "BARDZO WYSOKIE",
        "issue": "Przypisywanie żywności właściwości zapobiegania i zmniejszania ryzyka chorób (choroby krążenia, zawał, udar, nowotwory) z powołaniem na FDA oraz brak warunków wymaganych przez rozporządzenie (UE) nr 432/2012.",
        "quotes": [
            ("https://mowi.com/products/taste-health/:", "„Marine Omega-3 prevent development of cardio-vascular disease.”"),
            ("https://mowi.com/products/taste-health/:", "„Omega-3 fatty acids help to maintain healthy heart by lowering blood pressure, as well as reduce the risk of sudden death, heart attack, stroke 1,2,3 , and coronary heart disease 3,4”"),
            ("https://mowi.com/products/taste-health/:", "„They also possibly prevent psychiatric diseases, particularly cognitive decline in the elderly 5 .” oraz „Omega-3 fatty acids possibly prevent inflammation and reduce the risk of arthritis 5,6 .”"),
            ("https://mowi.com/products/ oraz /products/taste-health/:", "„Salmon is a consumer favourite, a superfood, highly nutritious and promoted by nutritionists worldwide” oraz „A true superfood”"),
            ("https://mowi.com/sustainability/:", "„Increased consumption of blue foods may reduce the consumption of terrestrial meats, consequently reducing diet-related chronic disease like hypertension, obesity and certain types of cancer. *BFA, 2021”"),
        ],
        "assessment": [
            "Ocena oświadczeń dotyczących właściwości zdrowotnych środków spożywczych podlega w pierwszej kolejności przepisom sektorowym jako regulacji szczególnej wobec ogólnych norm o nieuczciwych praktykach rynkowych. Jak wskazał Trybunał Sprawiedliwości Unii Europejskiej w wyroku z dnia 10 września 2020 r., C-363/19, Konsumentombudsmannen przeciwko Mezina AB (omówienie tezy; stan faktyczny dotyczył oświadczeń zdrowotnych na suplementach diety, a wyrok nie dotyczy Mowi ASA), szczególne zasady rozporządzenia (WE) nr 1924/2006 mają pierwszeństwo przed dyrektywą 2005/29/WE, a oświadczenia zdrowotne muszą opierać się na ogólnie uznanych dowodach naukowych i spełniać warunki tego rozporządzenia. Ponadto art. 7 ust. 3 i 4 rozporządzenia (UE) nr 1169/2011 ustanawia bezwzględny zakaz przypisywania środkom spożywczym w oznakowaniu, prezentacji i reklamie właściwości zapobiegania chorobom lub leczenia chorób ludzi bądź odwoływania się do takich właściwości. Sformułowania „prevent development of cardio-vascular disease”, „possibly prevent psychiatric diseases”, „possibly prevent inflammation” oraz sugestia ograniczania zachorowań na nadciśnienie i niektóre typy nowotworów naruszają ten zakaz wprost.",
            "Drugą warstwę naruszenia stanowią nieautoryzowane oświadczenia o zmniejszeniu ryzyka choroby w rozumieniu art. 14 ust. 1 lit. a rozporządzenia (WE) nr 1924/2006 („reduce the risk of sudden death, heart attack, stroke, and coronary heart disease”, „reduce the risk of arthritis”). W Unii Europejskiej każde oświadczenie o zmniejszeniu ryzyka choroby wymaga indywidualnego zezwolenia Komisji Europejskiej po ocenie naukowej EFSA, a odesłanie w przypisie nr 4 do amerykańskich „qualified health claims” agencji FDA nie legalizuje takiego komunikatu na rynku unijnym. Również twierdzenie o „lowering blood pressure” wykracza poza brzmienie dopuszczonego w rozporządzeniu Komisji (UE) nr 432/2012 oświadczenia („EPA and DHA contribute to the maintenance of normal blood pressure”), przy czym dla oświadczenia dotyczącego ciśnienia krwi unijny wykaz wymaga dziennego spożycia aż 3 g EPA i DHA oraz podania obowiązkowych zastrzeżeń dla konsumenta i wyłączenia żywności dla dzieci. Tymczasem porcja 100 g łososia Mowi zawiera według tabeli na tej samej podstronie 1,3 g EPA i DHA, a na stronie brak obowiązkowej informacji z art. 10 ust. 2 rozporządzenia nr 1924/2006 o znaczeniu zrównoważonej diety i ilości niezbędnej do uzyskania deklarowanego efektu.",
            "Określenie „superfood” użyte zarówno na stronie głównej produktów, jak i na podstronie zdrowotnej stanowi ogólne, nieswoiste odniesienie do korzyści dla zdrowia w rozumieniu art. 10 ust. 3 rozporządzenia (WE) nr 1924/2006. Jak wynika z wyroku TSUE z dnia 30 stycznia 2020 r., C-524/18, Dr. Willmar Schwabe przeciwko Queisser Pharma (omówienie tezy; sprawa dotyczyła relacji między ogólnym hasłem na froncie opakowania a szczegółowym oświadczeniem na odwrocie, a wyrok nie dotyczy Mowi ASA), samo słowo ogólne nie jest bezwzględnie zakazane, lecz jest dopuszczalne wyłącznie wtedy, gdy towarzyszy mu wyraźnie powiązane konkretne oświadczenie zdrowotne z unijnego wykazu. Na podstronie Mowi wyraz „A true superfood” otwiera właśnie listę nieautoryzowanych twierdzeń chorobowych i niespełnionych warunków, co czyni takie użycie bezprawnym. Analogicznie w orzecznictwie polskim Sąd Okręgowy w Łodzi, X Wydział Gospodarczy, w wyroku z dnia 5 września 2014 r., X GC 1024/13 (omówienie tezy w drodze analogii do przekazu o suplemencie diety; data 19 września 2014 r. w bazie SAOS jest datą publikacji uzasadnienia, a wyrok nie dotyczy Mowi ASA) oddzielił stanowcze zapewnienia o zwalczaniu schorzeń od dopuszczalnych informacji o działaniu wspomagającym.",
            "Komunikacja na witrynie mowi.com kierowana jest zarówno do konsumentów końcowych (przepisy kulinarne, profile odżywcze, kody QR na opakowaniach detalicznych), jak i do partnerów handlowych z sektora detalicznego i HoReCa. Pomocniczo należy wskazać wyrok TSUE z dnia 14 lipca 2016 r., C-19/15, Innova Vital (omówienie tezy; sprawa dotyczyła reklamy żywności kierowanej do lekarzy, a wyrok nie dotyczy Mowi ASA), z którego wynika, że rozporządzenie nr 1924/2006 obejmuje także komunikaty komercyjne kierowane do profesjonalistów, jeżeli dotyczą żywności przeznaczonej dla konsumenta końcowego. W płaszczyźnie horyzontalnej przypisywanie łososiowi działania zapobiegającego chorobom śmiertelnym spełnia również test wpływu na decyzję dotyczącą transakcji przeciętnego odbiorcy zgodnie z wyrokiem TSUE z dnia 19 grudnia 2013 r., C-281/12, Trento Sviluppo i Centrale Adriatica (omówienie tezy; sprawa dotyczyła promocji centrum handlowego, a wyrok nie dotyczy Mowi ASA) na gruncie art. 5 UPNPR oraz art. 16 ust. 1 pkt 1–2 UZNK.",
        ],
        "remediation": (
            "Należy niezwłocznie usunąć ze stron /products/taste-health/, /products/ oraz /sustainability/ wszelkie sformułowania mówiące o zapobieganiu chorobom układu krążenia, zmniejszaniu ryzyka zawału serca, udaru, nagłej śmierci, chorób psychicznych, zapalenia stawów, otyłości czy nowotworów, a także odwołania do praktyki amerykańskiej FDA. Przy haśle „superfood” należy albo z niego zrezygnować, albo bezpośrednio powiązać je wyłącznie z dopuszczonymi oświadczeniami z rozporządzenia (UE) nr 432/2012, w szczególności dotyczącymi prawidłowego funkcjonowania serca przy progu 250 mg EPA i DHA dziennie, uzupełnionymi o obowiązkowe pouczenia z art. 10 ust. 2 rozporządzenia nr 1924/2006."
        ),
        "claim": (
            "„Mowi Atlantic salmon is a source of long-chain omega-3 fatty acids (1.3 g of EPA and DHA per 100 g). "
            "EPA and DHA contribute to the normal function of the heart. The beneficial effect is obtained with a daily intake of 250 mg of EPA and DHA. "
            "Remember the importance of a varied and balanced diet and a healthy lifestyle.”"
        ),
    },
    {
        "no": "02",
        "title": "Sprzeczne twierdzenie „all of our soy is from organic-certified sources” wobec standardu ProTerra i konwencjonalnego profilu pasz Mowi Feed",
        "short_claim": "Deklaracja, że cała soja pochodzi ze źródeł z certyfikatem „organic”, przy standardzie ProTerra i paszach konwencjonalnych",
        "short_basis": "Rozp. (UE) 2018/848 art. 30; UPNPR art. 5; UZNK art. 16; u.o.k.k. art. 24. Przyszły standard EmpCo od 27.09.2027: pkt 4b",
        "table_basis": "Rozp. (UE) 2018/848 art. 30 ust. 1, 2 i 6; UPNPR art. 5; UZNK art. 16; u.o.k.k. art. 24. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4b ustawy z 4.09.2026 r.",
        "case_basis": "Rozporządzenie Parlamentu Europejskiego i Rady (UE) 2018/848 z 30.05.2018 r. art. 30 ust. 1, 2 i 6; art. 5 ust. 1 i ust. 2 pkt 1 UPNPR (Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 UZNK (Dz.U. z 2026 r. poz. 85); art. 24 u.o.k.k. (Dz.U. z 2025 r. poz. 1714); kontekstowo rozporządzenie (UE) 2023/1115 (EUDR). Przyszły standard EmpCo od 27.09.2027 r.: art. 7 pkt 4b ustawy z 4.09.2026 r. (Dz.U. z 2026 r. poz. 1353) — w dniu audytu nie stanowi naruszenia krajowej czarnej listy.",
        "risk": "BARDZO WYSOKIE",
        "issue": "Użycie ustawowo zastrzeżonego terminu „organic-certified” dla całej soi w publikacji z 23.11.2026 r., mimo że polityka paszowa wskazuje certyfikację ProTerra i produkcję konwencjonalną.",
        "quotes": [
            ("https://mowi.com/sustainability/how-mowi-uses-freshwater-responsibly/ (publikacja z 23.11.2026 r., eksponowana na stronie głównej):", "„For example, all of our soy is from organic-certified sources, with standards that include water management and good agricultural practices.”"),
            ("https://mowi.com/sustainability/commitments/sustainable-feed/:", "„100% of soy and palm oil is certified (Proterra, RSPO or equivalent)” oraz „KPIs: % Soy Proterra, Palm oil RSPO or eqivalent certified”"),
            ("https://mowi.com/about-us/our-structure/mowi-feed/:", "„We sell fish feeds in Norway, Scotland, Ireland, Iceland and the Faroe Islands for production according to conventional standards.”"),
            ("https://mowi.com/about-us/our-structure/mowi-feed/sourcing/:", "„The soy used in our feed is 100% deforestation-free.”"),
        ],
        "assessment": [
            "Na mocy art. 30 ust. 1, 2 i 6 rozporządzenia (UE) 2018/848 określenia odnoszące się do produkcji ekologicznej, w tym angielski termin „organic”, są zastrzeżone wyłącznie dla żywności, pasz oraz składników spełniających rygorystyczne wymogi produkcji ekologicznej określone w tym rozporządzeniu. Pomocniczo należy przywołać wyrok TSUE z dnia 5 listopada 2014 r., C-137/13, Herbaria Kräuterparadies (omówienie tezy na gruncie wcześniejszego rozporządzenia nr 834/2007; aktualnym aktem jest rozporządzenie 2018/848, a wyrok nie dotyczy Mowi ASA), zgodnie z którym możliwość oznaczania produktu jako ekologicznego zależy ściśle od spełnienia ustawowych reguł produkcji ekologicznej i nie może być zastąpiona uzasadnieniem marketingowym. Zarazem należy wyraźnie zastrzec, że ani art. 30, ani art. 32 ust. 1 lit. a rozporządzenia (UE) 2018/848 nie nakładają obowiązku zamieszczania numeru certyfikatu w tym samym akapicie artykułu na stronie internetowej; istotą naruszenia jest merytoryczna niezgodność kwalifikacji „all of our soy is from organic-certified sources” ze stanem opisanym przez samą spółkę.",
            "Zestawienie publikacji z 23 listopada 2026 r. (wyróżnionej w górnej sekcji strony głównej mowi.com) z podstroną /commitments/sustainable-feed/ oraz stroną oddziału Mowi Feed ujawnia bezpośrednią sprzeczność. W zobowiązaniach paszowych Mowi wskazuje certyfikację ProTerra (lub równoważną), która jest standardem zrównoważonej soi wolnej od GMO i wylesiania, lecz nie jest certyfikacją rolnictwa ekologicznego w rozumieniu rozporządzenia (UE) 2018/848. Ponadto sam dział Mowi Feed deklaruje, że wytwarza pasze w zakładach w Valsneset i Kyleakin dla produkcji prowadzonej według standardów konwencjonalnych, jedynie wzmiankując w jednym ze zdań rozwój linii pasz ekologicznych obok konwencjonalnych. Przypisanie certyfikacji „organic” całej soi wykorzystywanej w łańcuchu dostaw grupy wprowadza w błąd co do metody produkcji i właściwości surowca na gruncie art. 5 ust. 1 i ust. 2 pkt 1 UPNPR oraz art. 16 ust. 1 pkt 2 UZNK.",
            "W świetle testu materialności z wyroku TSUE z dnia 19 grudnia 2013 r., C-281/12, Trento Sviluppo i Centrale Adriatica (omówienie tezy; stan faktyczny dotyczył promocji centrum handlowego, a wyrok nie dotyczy Mowi ASA), informacja o ekologicznym pochodzeniu całego surowca białkowego w diecie ryb ma charakter istotny i może wpłynąć na wybór produktu przez konsumenta poszukującego żywności ekologicznej oraz przez nabywcę detalicznego w relacji B2B. Osobno hasło „100% deforestation-free” na podstronie /sourcing/ należy osadzić w kontekście rozporządzenia (UE) 2023/1115 (EUDR) w sprawie produktów wolnych od wylesiania: bez wskazania daty granicznej oraz odniesienia do procedury należytej staranności takie bezwarunkowe zapewnienie wymaga doprecyzowania, choć samo rozporządzenie EUDR nie jest tu naruszone samym brakiem opisu DDS w akapicie marketingowym.",
            "W horyzoncie od 27 września 2027 r. w Polsce (art. 3 ustawy z dnia 4 września 2026 r., Dz.U. z 2026 r. poz. 1353) przeniesienie cechy właściwej jedynie wybranej partii surowca lub linii produktów ekologicznych na całą grupę surowcową („all of our soy”) będzie dodatkowo podlegać ocenie jako przyszły standard z pkt 4b załącznika do ustawy (twierdzenie o całej działalności lub całym produkcie, gdy dotyczy ono tylko konkretnego aspektu). W dniu audytu (24 listopada 2026 r.) nie stanowi to jednak jeszcze naruszenia polskiej czarnej listy, a podstawę werdyktu stanowią obowiązujące przepisy rozporządzenia 2018/848, UPNPR i UZNK.",
        ],
        "remediation": (
            "Należy niezwłocznie sprostować artykuł z 23 listopada 2026 r., usuwając określenie „organic-certified” w odniesieniu do całości soi stosowanej przez Mowi. W jego miejsce należy wskazać rzeczywiste standardy certyfikacji łańcucha dostaw soi (np. ProTerra lub równoważne standardy non-GMO i braku wylesiania z podaniem daty granicznej) oraz wyraźnie oddzielić ewentualną odrębną linię produktów lub pasz posiadających certyfikację rolnictwa ekologicznego zgodną z rozporządzeniem (UE) 2018/848."
        ),
        "claim": (
            "„All soy used in Mowi’s conventional salmon feed is sourced under third-party sustainability and deforestation-free standards such as ProTerra or equivalent schemes, which include water management and agricultural criteria. "
            "(Organic-certified ingredients according to Regulation (EU) 2018/848 are used exclusively in certified organic feed lines.)”"
        ),
    },
    {
        "no": "03",
        "title": "Prezentacja celu „100% of harvest volumes sustainably certified by a GSSI-recognised standard” jako stanu osiągniętego przy danych ASC z 2021 r.",
        "short_claim": "Kafelek „100% of harvest volumes sustainably certified” bez oznaczenia celu przy 50% certyfikacji ASC na koniec 2021 r.",
        "short_basis": "UPNPR art. 5 i 6; UZNK art. 16; rozp. (UE) 1169/2011 art. 7; u.o.k.k. art. 24. Przyszły standard EmpCo od 27.09.2027: pkt 4a i 4b",
        "table_basis": "UPNPR art. 5 ust. 1–2 pkt 1 i art. 6 ust. 1; UZNK art. 16 ust. 1 pkt 2; rozp. 1169/2011 art. 7 ust. 1; u.o.k.k. art. 24. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a i 4b",
        "case_basis": "Art. 5 ust. 1 i ust. 2 pkt 1 oraz art. 6 ust. 1 i ust. 3 pkt 1 UPNPR (Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 UZNK (Dz.U. z 2026 r. poz. 85); art. 7 ust. 1 rozporządzenia (UE) nr 1169/2011; art. 24 u.o.k.k. (Dz.U. z 2025 r. poz. 1714). Przyszły standard EmpCo od 27.09.2027 r.: art. 7 pkt 4a i 4b ustawy z 4.09.2026 r. (Dz.U. z 2026 r. poz. 1353) — w dniu audytu nie stanowi naruszenia krajowej czarnej listy.",
        "risk": "WYSOKIE",
        "issue": "Zatarcie różnicy między celem a faktem dokonanym w głównym kafelku certyfikacji, nieaktualne dane ASC z 2021 r. oraz pomylenie skrótów GFSI i GSSI na podstronie partnerstw.",
        "quotes": [
            ("https://mowi.com/sustainability/ (kafelek w sekcji Commitments & Actions, bez gwiazdki i bez słowa „target”):", "„100% of harvest volumes sustainably certified by a GSSI-recognised standard”"),
            ("https://mowi.com/sustainability/commitments/sustainability-certifications/:", "„Our target is to have 100% of our yearly harvest volumes 100% certified with a Global Seafood Sustainability Initiative (GSSI)-recognised standard: the Aquaculture Stewardship Council (ASC), Best Aquaculture Practices (BAP) or GlobalGAP.”"),
            ("https://mowi.com/sustainability/certifications/:", "„Mowi continues to be a leader on ASC certification, having by the end of 2021 a total of 133 farming sites certified worldwide. This represents 50% of all our farming sites and accounts for 33% of all the ASC certified Atlantic salmon sites worldwide...”"),
            ("https://mowi.com/sustainability/key-partnerships/:", "Nagłówek „Global Food Safety Initiative” umieszczony nad akapitem zaczynającym się od słów „We continue to support the Global Sustainable Seafood Initiative...”"),
        ],
        "assessment": [
            "W głównym bloku zobowiązań na stronie /sustainability/ deklaracja „100% of harvest volumes sustainably certified by a GSSI-recognised standard” jest prezentowana bez wskazania, że stanowi cel polityki certyfikacyjnej („Our target”), bez podania roku sprawozdawczego oraz bez gwiazdki odsyłającej do warunków, mimo że sąsiednie kafelki zawierają datę „by 2030*”. Jak wskazał Sąd Najwyższy w wyroku z dnia 27 października 2021 r., I NSKP 11/21 (omówienie tezy w drodze analogii z obszaru prezentacji kosztów kredytu; wyrok nie dotyczy Mowi ASA), ocena komunikatu wymaga zbadania realnego sensu przekazywanego odbiorcy, a nie jedynie formalnej poprawności poszczególnych zdań — cel polityki nie jest równoznaczny z wynikiem już osiągniętym. Podobnie wyrok Sądu Najwyższego z dnia 4 marca 2014 r., III SK 34/13 (omówienie tezy; sprawa dotyczyła reklamy mobilnego internetu, a wyrok nie dotyczy Mowi ASA) potwierdza, że zastrzeżenia na innej podstronie nie zawsze usuwają mylące wrażenie dominującego hasła.",
            "Jednocześnie, stosując standard wyroku TSUE z dnia 26 października 2016 r., C-611/14, Canal Digital Danmark (omówienie tezy; sprawa dotyczyła prezentacji ceny abonamentu telewizyjnego, a wyrok nie dotyczy Mowi ASA), należy odnotować, że na podstronie /sustainability/certifications/ udostępniono odnośnik do tabeli „2026 Certification Table” oraz interaktywnego rejestru „ASC Dashboard”. Mimo tego udostępnienia sam tekst opisowy na tej samej podstronie nadal podaje jako aktualny stan dane z końca 2021 r. („by the end of 2021 a total of 133 farming sites... 50% of all our farming sites”), co w listopadzie 2026 r. tworzy pięcioletnią lukę informacyjną i zmusza odbiorcę do zestawiania treści marketingowej z plikiem PDF, aby ustalić rzeczywisty podział między ASC, GSA BAP i GLOBALG.A.P.",
            "Ponadto zrównanie uznania schematu przez inicjatywę GSSI — będącą narzędziem benchmarkingu, a nie jednolitym znakiem „zrównoważonego produktu” — z ogólnym hasłem „sustainably certified” całej rocznej produkcji może wprowadzać w błąd co do zakresu i równoważności wymogów środowiskowych poszczególnych programów, szczególnie wobec selektywnej ekspozycji odsetka 100% w nagłówku (por. wyrok Sądu Najwyższego z dnia 23 października 2019 r., I NSK 66/18 — omówienie tezy; sprawa dotyczyła ubezpieczeń z UFK i zakończyła się uchyleniem wyroku do ponownego rozpoznania, a wyrok nie dotyczy Mowi ASA). Na podstronie /key-partnerships/ występuje dodatkowo błąd merytoryczny mylący standard bezpieczeństwa żywności GFSI z inicjatywą GSSI. W teście z wyroku TSUE C-281/12 (Trento Sviluppo) deklaracja „100% zrównoważonej certyfikacji” może wpływać zarówno na decyzję konsumenta, jak i na ocenę dostawcy w zamówieniach sieci handlowych (art. 5 i 6 UPNPR, art. 16 ust. 1 pkt 2 UZNK).",
            "W ujęciu prospektywnym od 27 września 2027 r. (art. 3 ustawy z 4.09.2026 r., Dz.U. z 2026 r. poz. 1353) posługiwanie się ogólnym przymiotnikiem „sustainably certified” bez spełnienia przesłanek uznanej wysokiej efektywności środowiskowej oraz przedstawianie certyfikacji częściowej jako cechy całych wolumenów będzie oceniane przez pryzmat pkt 4a i 4b przyszłej czarnej listy EmpCo. W dniu audytu pozostaje to przyszłym standardem regulacyjnym.",
        ],
        "remediation": (
            "W kafelku na /sustainability/ należy wyraźnie wskazać, czy wartość 100% oznacza cel polityki czy zmierzony udział za konkretny rok (np. 2025), oraz podać w samym tekście podstrony aktualny rozkład procentowy wolumenów certyfikowanych oddzielnie w systemach ASC, GSA BAP i GLOBALG.A.P. Należy zaktualizować pięcioletnie dane ASC z końca 2021 r. do stanu z 2025/2026 r. oraz poprawić błędny nagłówek „Global Food Safety Initiative” na „Global Sustainable Seafood Initiative” na podstronie /key-partnerships/."
        ),
        "claim": (
            "„2025 certification status: [X]% of Mowi’s annual harvest volumes were certified under GSSI-recognised standards (ASC: [A]%, GSA BAP: [B]%, GLOBALG.A.P.: [C]%; see 2026 Certification Table). "
            "Mowi’s policy target is 100% certification of annual harvest volumes under these schemes.”"
        ),
    },
    {
        "no": "04",
        "title": "Uogólnione porównania śladu środowiskowego „blue foods / salmon” z „land-based foods / terrestrial meats” bez założeń metodycznych",
        "short_claim": "Twierdzenia o niższych emisjach GHG niż „land-based foods” i najniższym zużyciu wody „salmon production” bez granic LCA i jednostki",
        "short_basis": "UPNPR art. 5 i 6; UZNK art. 16 ust. 1 pkt 2 i ust. 3; rozp. 1169/2011 art. 7. Przyszły standard EmpCo od 27.09.2027: pkt 4b",
        "table_basis": "UPNPR art. 5 ust. 2 pkt 1 i art. 6 ust. 1; UZNK art. 16 ust. 1 pkt 2 i ust. 3; rozp. 1169/2011 art. 7. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4b",
        "case_basis": "Art. 5 ust. 1 i ust. 2 pkt 1 oraz art. 6 ust. 1 UPNPR (Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 oraz ust. 3 UZNK (Dz.U. z 2026 r. poz. 85); art. 7 ust. 1 lit. a rozporządzenia (UE) nr 1169/2011; art. 24 u.o.k.k. (Dz.U. z 2025 r. poz. 1714). Przyszły standard EmpCo od 27.09.2027 r.: art. 7 pkt 4b ustawy z 4.09.2026 r. (Dz.U. z 2026 r. poz. 1353) w zakresie uogólnienia cechy wybranego etapu na cały produkt — w dniu audytu nie stanowi naruszenia krajowej czarnej listy.",
        "risk": "WYSOKIE",
        "issue": "Przeniesienie średnich wyników szerokiej kategorii „blue foods” na łososia hodowlanego oraz brak granic systemu (pominięcie zużycia wody przy uprawie paszy) przy twierdzeniu o najniższym śladzie wodnym.",
        "quotes": [
            ("https://mowi.com/sustainability/:", "„Blue foods have lower GHG emissions than land-based foods (BFA, 2021).” oraz „Blue foods have lower freshwater use and land use compared to terrestrial meats. *BFA, 2021”"),
            ("https://mowi.com/products/salmon-farming/:", "„Blue or Aquatic Foods have lower greenhouse gas (GHG) emissions than land-based foods and, therefore, salmon farming can help bridge the gap between the potential of the ocean and the production of sustainable food that is healthy for people.”"),
            ("https://mowi.com/sustainability/how-mowi-uses-freshwater-responsibly/ (publikacja z 23.11.2026 r.):", "„In fact, salmon production has the lowest water consumption per kilogram of edible meat compared with other sources of farmed animal protein.”"),
            ("https://mowi.com/products/ (odmienne, ostrożniejsze ujęcie dla kontrastu):", "„The Blue Foods Assessment* found that blue foods from the ocean, such as salmon, on average have much greater nutritional benefits than terrestrial animal foods, and many also have a smaller environmental footprint.”"),
        ],
        "assessment": [
            "W pierwszej kolejności należy rozstrzygnąć kwalifikację tych zestawień jako reklamy porównawczej w świetle art. 16 ust. 3 UZNK oraz wyroku TSUE z dnia 19 kwietnia 2007 r., C-381/05, De Landtsheer (omówienie tezy; sprawa dotyczyła reklamy piwa odwołującej się do szampana, a wyrok nie dotyczy Mowi ASA). Brak wskazania konkurenta z nazwy nie wyklucza reklamy porównawczej, ale samo zestawienie szerokich kategorii nie przesądza jej automatycznie — konieczna jest możliwość pośredniego zidentyfikowania konkurencyjnych produktów lub przedsiębiorców o tym samym przeznaczeniu w konkretnym kontekście. Zestawienia „terrestrial meats” oraz „other sources of farmed animal protein” pozwalają konsumentowi i kupcowi zidentyfikować konkurencyjne produkty mięsne (np. wołowinę, wieprzowinę, drób), dlatego spełniają przesłanki reklamy porównawczej.",
            "Jak wynika z wyroku TSUE z dnia 19 września 2006 r., C-356/04, Lidl Belgium (omówienie tezy; sprawa dotyczyła koszyka cenowego sieci handlowych, a wyrok nie dotyczy Mowi ASA), zbiorcze porównanie kategorii nie jest bezwzględnie zakazane, o ile pozostaje obiektywne, reprezentatywne i sprawdzalne. Z kolei wyrok TSUE z dnia 8 lutego 2017 r., C-562/15, Carrefour Hypermarchés (omówienie tezy; sprawa dotyczyła porównywania cen różnych formatów sklepów, a wyrok nie dotyczy Mowi ASA) wskazuje, że przemilczenie istotnych założeń wpływających na wynik porównania może wprowadzać w błąd, przy czym wyrok ten nie nakłada prawnego obowiązku stosowania jednej wskazanej normy ISO w każdej reklamie. Podobną regułę potwierdza Sąd Najwyższy w wyroku z dnia 9 października 2014 r., IV CSK 56/14 (omówienie tezy; sprawa dotyczyła wyników badań technicznych producentów okien, a wyrok nie dotyczy Mowi ASA): nawet prawdziwe dane badawcze mogą wprowadzać w błąd przez sposób ujęcia, niepełność lub rozszerzający komentarz.",
            "Na stronie mowi.com dochodzi do istotnego przemilczenia założeń porównania w trzech płaszczyznach. Po pierwsze, raport Blue Food Assessment (2021) obejmuje całą różnorodną grupę żywności wodnej (w tym nieżywione paszą małże i wodorosty o skrajnie niskim śladzie), co sama podstrona /products/ oddaje ostrożnie słowem „many also have”, podczas gdy strony /sustainability/ i /salmon-farming/ uogólniają ten wynik na całe „blue foods” i wprost na hodowlę łososia. Po drugie, przeciwstawienie łososia całej kategorii „land-based foods” (bez ograniczenia do mięsa lądowego) obejmuje również żywność roślinną, wobec której twierdzenie o niższych emisjach GHG łososia hodowlanego zniekształca rzeczywiste dane. Po trzecie, kategoryczne zdanie z 23 listopada 2026 r., że produkcja łososia ma najniższe zużycie wody na kilogram jadalnego mięsa, umieszczono tuż pod opisem bezpośredniego zużycia wody w wylęgarniach RAS i zakładach przetwórczych, bez ujawnienia, czy uwzględniono pośredni ślad wodny upraw rolnych (np. soi) wchodzących w skład paszy.",
            "Pominięcie granic cyklu życia (zakres bezpośredni vs łańcuch paszowy) oraz zrównanie całej kategorii „land-based foods” z mięsem lądowym stanowi zaniechanie i wprowadzające w błąd przedstawienie porównania na podstawie art. 5 ust. 2 pkt 1 i art. 6 ust. 1 UPNPR oraz art. 16 ust. 1 pkt 2 i ust. 3 UZNK, spełniając test wpływu na decyzję dotyczącą transakcji (TSUE C-281/12). Od 27 września 2027 r. przypisywanie całemu produktowi cechy wynikającej z wybranego etapu bezpośredniego zużycia wody w obiektach własnych będzie również objęte zakresem przyszłego pkt 4b (bez przypisywania do pkt 4b przesłanki pochodzenia dowodu z innej kategorii).",
        ],
        "remediation": (
            "Porównania nie trzeba usuwać w całości, lecz należy je doprecyzować: zastąpić szerokie pojęcie „land-based foods” konkretnymi gatunkami białka zwierzęcego (np. wołowiną lub wieprzowiną), odnieść wartości bezpośrednio do łososia atlantyckiego zamiast uśrednionej grupy „blue foods”, wskazać jednostkę funkcjonalną (np. kg części jadalnej lub 100 g białka) oraz jasno określić granice analizy (w szczególności, czy wskaźnik wody i GHG obejmuje produkcję paszy oraz transport) wraz z pełnym cytatem źródła i roku badania."
        ),
        "claim": (
            "„According to [cited peer-reviewed LCA / Blue Food Assessment dataset, year], farmed Atlantic salmon has lower average GHG emissions per [functional unit, e.g. kg edible weight / 100 g protein] than [specified terrestrial meats, e.g. beef and pork], within [cradle-to-farm-gate / processing] system boundaries including feed production. "
            "Direct freshwater consumption in Mowi’s RAS smolt facilities does not represent the total supply-chain water footprint of feed crops.”"
        ),
    },
    {
        "no": "05",
        "title": "Absolutny superlatyw „the most sustainable protein producer” (Coller FAIRR 2019–2024) oraz zapewnienia o braku negatywnego i trwałego wpływu na środowisko",
        "short_claim": "Nagłówki „the most sustainable protein producer” (2019–2024 w listopadzie 2026 r.) i zapewnienia „ensure no negative impacts / leave no permanent impact”",
        "short_basis": "UPNPR art. 5 i 6; UZNK art. 16 ust. 1 pkt 2; u.o.k.k. art. 24. Przyszły standard EmpCo od 27.09.2027: pkt 4a i 4b",
        "table_basis": "UPNPR art. 5 ust. 1–2 pkt 1 i art. 6 ust. 1; UZNK art. 16 ust. 1 pkt 2; u.o.k.k. art. 24. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a i 4b",
        "case_basis": "Art. 5 ust. 1 i ust. 2 pkt 1 oraz art. 6 ust. 1 UPNPR (Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 UZNK (Dz.U. z 2026 r. poz. 85); art. 24 u.o.k.k. (Dz.U. z 2025 r. poz. 1714). Przyszły standard EmpCo od 27.09.2027 r.: art. 7 pkt 4a i 4b ustawy z 4.09.2026 r. (Dz.U. z 2026 r. poz. 1353) — w dniu audytu nie stanowi naruszenia krajowej czarnej listy.",
        "risk": "WYSOKIE",
        "issue": "Uogólnienie wyniku indeksu 60 producentów białka zwierzęcego na wszystkich producentów białka (z niespójną datą 2019–2024) oraz wywodzenie braku trwałego wpływu z badania jednej zamkniętej farmy.",
        "quotes": [
            ("https://mowi.com/ oraz https://mowi.com/about-us/:", "„We are the most sustainable protein producer. Mowi has been ranked the most sustainable protein producer by the Coller FAIRR Protein Producer Index for the fifth year in a row. *MOWI ranked #1 in the 2019-2024 Coller FAIRR Protein Producer Index, a comprehensive assessment of the largest animal protein producers on critical environmental, social and governance (ESG) issues.”"),
            ("https://mowi.com/about-us/ (nagłówek H1 bez przypisu ograniczającego):", "„Mowi, the world’s largest supplier of farm-raised salmon and most sustainable protein producer”"),
            ("https://mowi.com/knowledge/does-salmon-farming-have-an-impact-on-the-seabed/:", "„Mowi adheres to ongoing rigorous operations to ensure no negative impacts.” oraz „By farming within the environment’s capacity, we know we leave no permanent impact.”"),
        ],
        "assessment": [
            "W zestawieniu z wyrokiem Sądu Najwyższego z dnia 9 października 2014 r., IV CSK 56/14 (omówienie tezy; sprawa dotyczyła okien, a wyrok nie dotyczy Mowi ASA), nawet obiektywnie prawdziwy wynik rankingu może wprowadzać w błąd przez sposób jego skomentowania i rozszerzenie zakresu poza badaną próbę. Indeks Coller FAIRR ocenia wybraną grupę największych notowanych producentów białka zwierzęcego (60 spółek), nie obejmuje natomiast producentów białka roślinnego ani wszystkich producentów białka na świecie. Tymczasem nagłówek na stronie głównej i nagłówek H1 na podstronie /about-us/ prezentują bezwarunkową tezę „We are the most sustainable protein producer”.",
            "Ponadto między zdaniem głównym („for the fifth year in a row”) a przypisem („2019-2024”, co obejmuje sześć lat kalendarzowych) zachodzi niespójność redakcyjna, a na dzień 24 listopada 2026 r. witryna nie wyjaśnia statusu spółki w kolejnych edycjach indeksu po 2024 r. Jak wskazuje wyrok TSUE z dnia 4 czerwca 2015 r., C-195/14, Teekanne (omówienie tezy na gruncie oznakowania herbaty; wyrok nie dotyczy Mowi ASA) oraz wyrok Sądu Najwyższego z dnia 4 marca 2014 r., III SK 34/13, drobny przypis nie zawsze neutralizuje kategoryczne wrażenie wywołane przez tytuł i lead strony, szczególnie gdy w nagłówku H1 podstrony /about-us/ przypisu w ogóle brak.",
            "Równoległym naruszeniem jest absolutyzujące zapewnienie w artykule o wpływie na dno morskie (/knowledge/does-salmon-farming-have-an-impact-on-the-seabed/), iż procedury Mowi „ensure no negative impacts” oraz „we know we leave no permanent impact”. Podstawą dowodową przywołaną w tym samym tekście jest trzyletnie badanie regeneracji dna po zamknięciu jednej farmy w Szkocji (Isle of Ewe / Loch Ewe, gdzie odsetek próbek spełniających status „Good Ecological” wzrósł z 67% w 2020 r. do 100% w 2023 r. po zaprzestaniu produkcji). Wywodzenie z regeneracji po zamknięciu pojedynczego obiektu tezy o braku jakichkolwiek negatywnych lub trwałych oddziaływań w skali globalnej akwakultury Mowi w siedmiu krajach stanowi nieuprawnione uogólnienie, które w teście z wyroku TSUE C-281/12 może zniekształcać decyzje konsumentów i kontrahentów (art. 5 ust. 1 i ust. 2 pkt 1 oraz art. 6 ust. 1 UPNPR, art. 16 ust. 1 pkt 2 UZNK).",
            "Od 27 września 2027 r. oba komunikaty wejdą dodatkowo w obszar zastosowania przyszłych punktów 4a (ogólne twierdzenie o zrównoważoności bez uznanej wysokiej efektywności ekologicznej) oraz 4b (twierdzenie o całej działalności na podstawie wycinka lub pojedynczego aspektu) nowelizacji UPNPR z dnia 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353).",
        ],
        "remediation": (
            "W nagłówkach na stronie głównej i /about-us/ należy doprecyzować zakres rankingu już w zdaniu głównym (wskazując ocenianą grupę największych notowanych producentów białka zwierzęcego w indeksie Coller FAIRR), ujednolicić wskazanie lat edycji oraz uzupełnić status na 2025/2026 r. W artykule o dnie morskim należy usunąć kategoryczne zwroty „ensure no negative impacts” i „leave no permanent impact”, pozostawiając rzetelny opis wyników badania regeneracji po zamknięciu farmy Loch Ewe oraz okresowego ugorowania stanowisk."
        ),
        "claim": (
            "„Mowi ranked #1 among the major listed animal protein producers assessed in the Coller FAIRR Protein Producer Index [years of editions]. "
            "In a three-year post-closure monitoring study at the former Loch Ewe site in Scotland (2020–2023), benthic samples meeting ‘Good Ecological’ status increased from 67% to 100% after farming operations ceased.”"
        ),
    },
    {
        "no": "06",
        "title": "Ogólne określenia „climate-friendly” oraz prezentacja celów na 2030 r. (GHG, odpady, PPWR) bez widocznego postępu i planu na podstronach",
        "short_claim": "Hasła „climate-friendly food/lifestyle” oraz cele na 2030 r. i PPWR bez bieżącego postępu na podstronach i z błędami kafelków",
        "short_basis": "UPNPR art. 5 i 6; UZNK art. 16. Przyszły standard EmpCo od 27.09.2027: pkt 4a, 10a oraz weryfikacja planu celów",
        "table_basis": "UPNPR art. 5 ust. 1 i art. 6 ust. 1; UZNK art. 16 ust. 1 pkt 2. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a, 10a i wymóg planu celów",
        "case_basis": "Art. 5 ust. 1 i art. 6 ust. 1 UPNPR (Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 UZNK (Dz.U. z 2026 r. poz. 85). Przyszły standard EmpCo od 27.09.2027 r.: art. 7 pkt 4a, pkt 10a (w zakresie regulacji UE PPWR) oraz wymóg planu wdrażania i weryfikacji eksperckiej dla celów przyszłej efektywności ekologicznej z ustawy z 4.09.2026 r. (Dz.U. z 2026 r. poz. 1353) — w dniu audytu nie stanowi naruszenia krajowej czarnej listy.",
        "risk": "ŚREDNIE",
        "issue": "Niesprecyzowane hasła „climate-friendly”, cele na 2030 r. bez aktualnego stopnia realizacji obok kafelków, ucięte liczby w siatce celów i przedstawianie wymogu UE PPWR jako celu własnego.",
        "quotes": [
            ("https://mowi.com/products/salmon-farming/ oraz /products/traceability/:", "„Salmon farming is climate friendly food production.”"),
            ("https://mowi.com/products/:", "„Our goal is to deliver top-quality salmon and inspire a healthy and climate-friendly lifestyle”"),
            ("https://mowi.com/sustainability/commitments/climate-change/:", "„Mowi commits to reduce absolute scope 1 and 2 GHG emissions 50.6% by 2030 from a 2019 base year. Mowi commits to reduce absolute scope 3 GHG emissions 27.5% within the same timeframe. Mowi commits to reduce absolute scope 3 FLAG GHG emissions 33.33% within the same timeframe”"),
            ("https://mowi.com/sustainability/commitments/plastics/:", "„By 2030, 100% of our EU plastic packaging adhering to PPWR (Recyclability: recycling performance grade >70%; Recycled content: contact sensitive PET >30%; contact sensitivee other plastic >10%; other plastic packaging >35%)”"),
            ("https://mowi.com/sustainability/ (usterki prezentacji i spójności):", "Kafelek „waste to landfill at our direct operations by 2030*” bez widocznej wartości liczbowej „0”, odnośniki „*” bez legendy na tej podstronie oraz rozbieżność danych korporacyjnych („12,500 Employees” i „7” zakładów w Azji na stronie głównej wobec „11 500 people” i „6” zakładów w Azji na /about-us/)."),
        ],
        "assessment": [
            "W porównaniu z niejednolitą komunikacją historyczną należy odnotować na korzyść badanej witryny, że na dzień 24 listopada 2026 r. cele redukcji gazów cieplarnianych (50,6% dla zakresów 1 i 2, 27,5% dla zakresu 3 oraz 33,33% dla zakresu 3 FLAG do 2030 r. względem roku bazowego 2019) są na stronie głównej /sustainability/ i podstronie /commitments/climate-change/ spójne liczbowo, a w sekcji ESG Library spółka udostępnia raport CDP 2025 oraz ujawnienie IFRS S2. Zgodnie z wyrokiem TSUE z dnia 26 października 2016 r., C-611/14, Canal Digital Danmark (omówienie tezy; wyrok nie dotyczy Mowi ASA), przy ocenie zaniechania uwzględnia się udostępnienie informacji w innych częściach witryny.",
            "Pomimo tej dokumentacji w bibliotece ESG bezpośrednia komunikacja produktowa i zobowiązaniowa wykazuje mankamenty uzasadniające ocenę średniego ryzyka i pilną korektę techniczną. Po pierwsze, bezwarunkowe hasło „Salmon farming is climate friendly food production” nie precyzuje na podstronach produktowych żadnych parametrów śladu węglowego ani odniesienia, a także nie jest wsparte uznaną wysoką efektywnością ekologiczną. Po drugie, na podstronach /commitments/ przedstawiono liczne cele na 2030 r. bez wskazania bieżącego poziomu wykonania za ostatni zamknięty rok (2025), a na samej stronie /sustainability/ kafelki „waste to landfill” i „sea lice” ucinają wartość liczbową „0” / „0%” oraz odsyłają do gwiazdki bez legendy. Po trzecie, podstrona /commitments/plastics/ prezentuje jako dobrowolny cel spółki dostosowanie opakowań do rozporządzenia UE w sprawie opakowań i odpadów opakowaniowych (PPWR) do 2030 r., co dotyczy obowiązkowego reżimu unijnego dla całego rynku UE.",
            "W reżimie obowiązującym w Polsce w dniu audytu (art. 5 i 6 UPNPR, art. 16 UZNK) niekompletna informacja na kafelku nie przesądza automatycznie wysokiego naruszenia (por. TSUE C-281/12 oraz SN I NSKP 11/21), jednak od 27 września 2027 r. (ustawa z 4.09.2026 r., Dz.U. z 2026 r. poz. 1353) sytuacja ta ulegnie zaostrzeniu. Hasło „climate-friendly” bez bezpośredniej specyfikacji będzie objęte przyszłym zakazem pkt 4a, prezentowanie przyszłego ustawowego standardu PPWR jako wyróżnika oferty będzie oceniane pod kątem pkt 10a, a utrzymanie celów redukcji GHG i zero odpadów na 2030 r. w komunikacji do konsumentów wymagać będzie szczegółowego, realistycznego planu wdrażania z mierzalnymi celami pośrednimi, przypisanymi zasobami oraz regularną weryfikacją przez niezależnego eksperta zewnętrznego podaną do wiadomości publicznej.",
        ],
        "remediation": (
            "Należy zastąpić ogólne określenie „climate-friendly” w opisach hodowli i produktów konkretnymi danymi o intensywności emisji lub odesłaniem do zweryfikowanego śladu węglowego; uzupełnić siatkę celów na /sustainability/ o brakujące wartości liczbowe i legendę gwiazdek; dodać przy każdym celu na 2030 r. aktualny wynik wykonania (np. za 2025 r.) oraz bezpośredni link do planu przejścia i raportu z weryfikacji zewnętrznej; oddzielić obowiązkowe wymogi PPWR od dobrowolnych celów ponadustawowych oraz ujednolicić dane o zatrudnieniu i liczbie zakładów."
        ),
        "claim": (
            "„Mowi targets a 50.6% reduction in absolute Scope 1 and 2 GHG emissions, a 27.5% reduction in Scope 3 emissions, and a 33.33% reduction in Scope 3 FLAG emissions by 2030 compared to the 2019 baseline (status in 2025: [X]%; progress and transition plan published in Mowi Annual Report / IFRS S2 Disclosure). "
            "Mowi is aligning its EU branded packaging with the mandatory EU Packaging and Packaging Waste Regulation (PPWR) timelines.”"
        ),
    },
    {
        "no": "07",
        "title": "Deklaracja zgodności z Celami Zrównoważonego Rozwoju ONZ (SDGs) przy sprzeczności podstrony programów oraz ogólne komunikaty praw człowieka",
        "short_claim": "Zgodność z SDG ONZ przy „0 programs” dla SDG 2, 6 i 7 oraz ogólne twierdzenie o braku naruszeń praw człowieka",
        "short_basis": "UPNPR art. 5 i 6; UZNK art. 16 (ŚREDNIE). Wytyczne OECD (2023) i ścieżka KPK OECD — poziom NISKIE",
        "table_basis": "UPNPR art. 5 ust. 1 i art. 6 ust. 1; UZNK art. 16 ust. 1 pkt 2. Wytyczne OECD (2023) / KPK OECD — poziom NISKIE (nie podnosi werdyktu)",
        "case_basis": "Art. 5 ust. 1 i art. 6 ust. 1 UPNPR (Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 UZNK (Dz.U. z 2026 r. poz. 85) — poziom ŚREDNIE. Wytyczne OECD dla przedsiębiorstw wielonarodowych (edycja 2023) w zakresie należytej staranności i komunikacji o prawach człowieka — poziom NISKIE (zalecenia; ścieżka przed Krajowym Punktem Kontaktowym OECD nie jest sprawą sądową i nie kończy się wyrokiem).",
        "risk": "ŚREDNIE",
        "issue": "Wewnętrzna sprzeczność między deklaracją działań na rzecz SDG 2, wody i OZE a wykazem „0 programs” na stronie SDG oraz ogólna deklaracja braku naruszeń praw człowieka.",
        "quotes": [
            ("https://mowi.com/sustainability/:", "„Our commitments are set to make our business future-proof and are aligned with the UN Sustainable Development Goals (SDGs).”"),
            ("https://mowi.com/about-us/our-structure/mowi-feed/sourcing/:", "„We are working towards the 2023 FAO global roadmap to achieve the SDG 2 without breaching the 1.5° C threshold.”"),
            ("https://mowi.com/sustainability/the-sustainable-development-goals-sdgs/:", "Wykaz liczników programów wskazuje „0 programs” przy SDG 2 (Zero hunger), SDG 6 (Clean water and sanitation) oraz SDG 7 (Affordable and clean energy)."),
            ("https://mowi.com/sustainability/commitments/ethical-business-conduct/:", "„No breach of human rights principles or policy identified”"),
        ],
        "assessment": [
            "Cele Zrównoważonego Rozwoju ONZ (Agenda 2030) stanowią program polityczny skierowany w pierwszej kolejności do państw, a nie system certyfikacji przedsiębiorstwa. Na witrynie mowi.com deklaracja „aligned with the UN Sustainable Development Goals” oraz odwołanie do realizacji SDG 2 na stronie paszowej Mowi Feed zderzają się z dedykowaną podstroną /the-sustainable-development-goals-sdgs/, która przy SDG 2 (Zero hunger), SDG 6 (Clean water and sanitation — mimo rozbudowanej polityki wodnej i wyniku CDP A) oraz SDG 7 (Affordable and clean energy — mimo celów dotyczących OZE) wyświetla „0 programs”. Taka wewnętrzna sprzeczność techniczno-merytoryczna na własnej witrynie osłabia wiarygodność komunikacji i wprowadza niejasność co do rzeczywistej skali wkładu w cele ONZ na gruncie art. 5 ust. 1 i art. 6 ust. 1 UPNPR oraz art. 16 ust. 1 pkt 2 UZNK.",
            "W praktyce nadzorczej Prezesa UOKiK z dnia 19 stycznia 2026 r. (komunikaty o postawieniu zarzutów m.in. w sprawie ogólnych odwołań do SDG bez wskazania konkretnego zakresu i skali) widoczna jest zwiększona kontrola tego typu haseł; należy jednak podkreślić, że postawienie zarzutów przez organ nie jest równoznaczne z prawomocnym wyrokiem sądowym ani nie dotyczy Mowi ASA. Po uwzględnieniu testu materialności z wyroku TSUE C-281/12 (Trento Sviluppo) sama niespójność licznika na podstronie SDG uzasadnia kwalifikację na poziomie ryzyka ŚREDNIE.",
            "Osobno w odniesieniu do deklaracji „No breach of human rights principles or policy identified” na podstronie /ethical-business-conduct/ ocena z perspektywy odpowiedzialnego prowadzenia biznesu odwołuje się do Wytycznych OECD dla przedsiębiorstw wielonarodowych (edycja 2023), które otrzymują etykietę ryzyka NISKIE. Są to zalecenia miękkiego prawa; ewentualne zakwestionowanie braku wskazania okresu, metodologii weryfikacji i zakresu łańcucha dostaw może trafić do właściwego Krajowego Punktu Kontaktowego OECD (w Polsce przy ministrze właściwym ds. rozwoju). Postępowanie przed KPK OECD nie jest sprawą sądową i nie kończy się wyrokiem, lecz publiczną oceną, zaleceniami oraz skutkiem wizerunkowym, dlatego nie podnosi werdyktu raportu.",
        ],
        "remediation": (
            "Należy ujednolicić bazę programów na podstronie /the-sustainable-development-goals-sdgs/, przypisując faktyczne programy do SDG 2, SDG 6 i SDG 7 albo usuwając niespójne liczniki „0 programs”, oraz zastąpić ogólne „aligned with the UN SDGs” precyzyjnym sformułowaniem wskazującym na dobrowolne odniesienie wybranych polityk grupy do celów ONZ bez sugerowania certyfikacji. Przy oświadczeniu o braku naruszeń praw człowieka należy dodać rok sprawozdawczy, zakres podmiotowy (operacje własne vs dostawcy) oraz odniesienie do opublikowanego raportu należytej staranności."
        ),
        "claim": (
            "„Mowi uses selected UN Sustainable Development Goals (in particular SDGs 2, 3, 6, 8, 12, 13 and 14) as a voluntary reference framework for its internal sustainability programmes; this reference does not constitute third-party certification. "
            "In [reporting year], no substantiated breach of Mowi’s Human Rights Policy was identified within own operations covered by internal grievance and due-diligence monitoring (details: Mowi Transparency Act Statement).”"
        ),
    },
]


# ==============================================================================
# 3. RAPORT PEŁNY (skład: ReportDoc z sklad_raportu.py)
# ==============================================================================


def build_full_report():
    """Generuje pełny raport audytu greenwashingu dla MOWI ASA (https://mowi.com/)."""
    report = ReportDoc(
        reference=AUDIT_REF,
        document_name="Audyt greenwashingu — MOWI ASA (mowi.com)",
        audit_date=AUDIT_DATE,
    )
    report.doc.core_properties.subject = "Audyt ryzyk greenwashingu — MOWI ASA (mowi.com)"

    report.title("Audyt Ryzyk Greenwashingu: MOWI ASA")
    report.subtitle(
        "Weryfikacja komunikacji środowiskowej, żywieniowej, zdrowotnej i certyfikacyjnej na witrynie https://mowi.com/ "
        "(strona główna, O nas, Produkty, Mowi Feed, Zrównoważony rozwój i powiązane podstrony; stan na 24 listopada 2026 r.)."
    )

    report.meta(
        [
            (
                "Badany podmiot",
                "Mowi ASA, Sandviksboder 77AB, 5035 Bergen, Norwegia (adres korespondencyjny: PO Box 4102 Sandviken, 5835 Bergen, Norwegia). "
                "Spółka notowana na Oslo Stock Exchange, obecna w 26 krajach, oferująca produkty w ponad 70 krajach (w Polsce przez spółkę zależną Mowi Poland S.A.).",
                False,
            ),
            (
                "Przedmiot audytu",
                "Publiczna treść anglojęzycznej witryny https://mowi.com/ (stan na 24 listopada 2026 r.): strona główna, /about-us/, /about-us/our-strategy/, "
                "/about-us/our-structure/mowi-feed/ (w tym /sourcing/, /strategy-and-ambitions/), /products/ (w tym /salmon-farming/, /taste-health/, /quality/, "
                "/food-safety/, /traceability/, /mowi-4-0-smart-farming/), /sustainability/ (w tym /commitments/, /certifications/, /the-sustainable-development-goals-sdgs/, "
                "/how-mowi-uses-freshwater-responsibly/ z 23.11.2026 r., /key-partnerships/) oraz /knowledge/does-salmon-farming-have-an-impact-on-the-seabed/. "
                "Nie badano raportu rocznego Annual Report 2025 ani dokumentów PDF z biblioteki ESG Library.",
                False,
            ),
            (
                "Sektor działalności",
                "Akwakultura, produkcja pasz, przetwórstwo i sprzedaż łososia atlantyckiego oraz owoców morza (sektor spożywczy / FMCG; kanały detaliczny, HoReCa i przemysłowy). "
                "Komunikacja kierowana jest zarówno do konsumentów (B2C), jak i do kontrahentów biznesowych (B2B).",
                False,
            ),
            (
                "Reżim prawny",
                "Rozporządzenie (WE) nr 1924/2006; rozporządzenie Komisji (UE) nr 432/2012; rozporządzenie (UE) nr 1169/2011; rozporządzenie (UE) 2018/848; "
                "ustawa z 23.08.2007 r. o przeciwdziałaniu nieuczciwym praktykom rynkowym (Dz.U. z 2023 r. poz. 845); "
                "ustawa z 16.04.1993 r. o zwalczaniu nieuczciwej konkurencji (Dz.U. z 2026 r. poz. 85); "
                "ustawa z 16.02.2007 r. o ochronie konkurencji i konsumentów (Dz.U. z 2025 r. poz. 1714); kontekstowo rozporządzenie (UE) 2023/1115 (EUDR) i Wytyczne OECD (2023). "
                "W dniu audytu (24 listopada 2026 r.) czarna lista dyrektywy (UE) 2024/825 (EmpCo) nie jest jeszcze obowiązującym prawem polskim.",
                False,
            ),
            ("Data audytu", AUDIT_DATE, False),
            ("Werdykt operacyjny", VERDICT_RED, True),
        ]
    )

    report.verdict(
        VERDICT_RED,
        [
            "Witryna https://mowi.com/ narusza już dziś bezwzględnie obowiązujące przepisy prawa żywnościowego oraz ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym i ustawy o zwalczaniu nieuczciwej konkurencji. "
            "Na podstronach produktowych i zrównoważonego rozwoju przypisano łososiowi oraz kwasom omega-3 właściwości zapobiegania chorobom i zmniejszania ryzyka zawału, udaru, nagłej śmierci i nowotworów (art. 7 ust. 3–4 rozporządzenia (UE) nr 1169/2011, art. 10 i 14 rozporządzenia (WE) nr 1924/2006), a w publikacji z 23 listopada 2026 r. zadeklarowano, że cała soja pochodzi ze źródeł „organic-certified” wbrew standardowi ProTerra i konwencjonalnemu profilowi pasz Mowi Feed (art. 30 rozporządzenia (UE) 2018/848, art. 5 UPNPR, art. 16 UZNK).",
            "Dodatkowe naruszenia przepisów już obowiązujących obejmują prezentowanie celu 100% certyfikacji GSSI jako stanu osiągniętego przy danych ASC z 2021 r., niekompletne porównania środowiskowe z żywnością lądową oraz rozszerzanie zakresu rankingu Coller FAIRR (2019–2024) i badania dna morskiego jednej zamkniętej farmy na całą działalność grupy.",
            "Nowe zakazy czarnej listy dyrektywy (UE) 2024/825 (EmpCo) wchodzą w życie w Polsce 27 września 2027 r. (ustawa z 4 września 2026 r., Dz.U. z 2026 r. poz. 1353). W dniu audytu (24 listopada 2026 r.) nie stanowią one podstawy werdyktu krajowego i zostały ujęte wyłącznie jako przyszły standard regulacyjny wymagający dostosowania w wyznaczonym etapie przed tą datą.",
        ],
    )

    # 1. Analiza wstępna
    report.h1("1. Analiza wstępna")

    report.h2("1.1. Reżim prawny")
    report.paragraph(
        "Ocenę prawną przeprowadzono według stanu prawnego na dzień 24 listopada 2026 r., rozdzielając przepisy już obowiązujące od przyszłego standardu wynikającego z dyrektywy Parlamentu Europejskiego i Rady (UE) 2024/825 (EmpCo). "
        "Art. 4 ust. 1 dyrektywy (UE) 2024/825 wyznaczył 27 września 2026 r. jako datę rozpoczęcia stosowania krajowych przepisów transponujących w państwach członkowskich Unii Europejskiej. "
        "Polska ustawa z dnia 4 września 2026 r. o zmianie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawy o prawach konsumenta (Dz.U. z 2026 r. poz. 1353), ogłoszona w dniu 19 października 2026 r., wyznacza w art. 3 dzień wejścia w życie na 27 września 2027 r. "
        "Z tego względu do dnia 26 września 2027 r. włącznie nowe pozycje czarnej listy nie stanowią jeszcze w Polsce obowiązujących przepisów sankcyjnych, a samo utrzymywanie treści na stronie internetowej nie przyspiesza terminu wejścia w życie polskiej nowelizacji."
    )
    report.paragraph(
        "W konsekwencji werdykt niniejszego audytu oparto wyłącznie na przepisach już obowiązujących w dniu badania: bezpośrednio stosowanych rozporządzeniach sektorowych UE (rozporządzeniu (WE) nr 1924/2006 o oświadczeniach żywieniowych i zdrowotnych, rozporządzeniu Komisji (UE) nr 432/2012, rozporządzeniu (UE) nr 1169/2011 oraz rozporządzeniu (UE) 2018/848 o produkcji ekologicznej), a także art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (UPNPR), art. 16 ustawy o zwalczaniu nieuczciwej konkurencji (UZNK) oraz art. 24 i 106 ustawy o ochronie konkurencji i konsumentów (u.o.k.k.). "
        "Ponieważ jednak witryna https://mowi.com/ ma charakter globalny w języku angielskim i jest dostępna w państwach członkowskich UE, które mogły wdrożyć dyrektywę 2024/825 od 27 września 2026 r. (w tym na rynkach zbytu grupy w Europie kontynentalnej i Irlandii), w raporcie wskazano dodatkowo przyszły standard EmpCo jako istotne ryzyko transgraniczne oraz obowiązkowy kierunek dostosowania przed 27 września 2027 r. "
        "W odniesieniu do regulacji sektorowych zastosowano zasadę pierwszeństwa norm szczególnych (lex specialis) przed klauzulą generalną nieuczciwych praktyk rynkowych."
    )

    report.h2("1.2. Profil działalności")
    report.paragraph(
        "Mowi ASA z siedzibą w Bergen (Norwegia) jest największym na świecie producentem łososia atlantyckiego, z wolumenem zbiorów 559 000 ton w 2025 r. (około 20% udziału w rynku globalnym) oraz przychodami 5,7 mld EUR. "
        "Grupa prowadzi zintegrowany pionowo model biznesowy w trzech segmentach: Mowi Feed (wytwórnie pasz w Valsneset w Norwegii i Kyleakin w Szkocji), Mowi Farming (hodowla i przetwórstwo pierwotne w Norwegii, Szkocji, Kanadzie, Chile, Irlandii, na Wyspach Owczych i Islandii) oraz Mowi Sales & Marketing (przetwórstwo wtórne i sprzedaż w ponad 70 krajach, w tym poprzez polską spółkę Mowi Poland S.A. w Ustce). "
        "Strona główna komunikuje dostarczanie 9,7 miliona posiłków dziennie i kieruje przekaz zarówno do konsumentów (przepisy kulinarne, informacje zdrowotne, śledzenie pochodzenia kodem QR), jak i do klientów instytucjonalnych z branży retail i HoReCa."
    )
    report.paragraph(
        "Centralną osią komunikacji wizerunkowej jest strategia „Leading the Blue Revolution”, przedstawiająca akwakulturę morską jako rozwiązanie globalnych wyzwań żywnościowych i klimatycznych oraz pozycjonująca grupę jako najbardziej zrównoważonego producenta białka na świecie. "
        "O ile część przekazu opiera się na mierzalnych działaniach technicznych (np. systemy recyrkulacji wody RAS, raportowanie CDP i IFRS S2 w bibliotece ESG, wykaz certyfikatów na ASC Dashboard), o tyle w warstwie haseł produktowych i artykułów wizerunkowych zidentyfikowano istotne naruszenia regulacyjne. "
        "Najpoważniejsze z nich dotyczą przypisywania łososiowi i kwasom omega-3 właściwości zapobiegania chorobom śmiertelnym oraz błędnego przypisania certyfikacji „organic” całej soi wykorzystywanej w paszach konwencjonalnych."
    )

    report.h2("1.3. Materiał z witryny")
    report.paragraph(
        "Dokumentacja dowodowa została zebrana bezpośrednio z publicznie dostępnych podstron https://mowi.com/ według stanu na 24 listopada 2026 r., z uwzględnieniem najnowszych publikacji aktualności (w tym artykułu z 23 listopada 2026 r. o gospodarce wodnej oraz komunikatu o wynikach Q3 2026 z 4 listopada 2026 r.). "
        "Poniżej przedstawiono zrzut górnej części strony głównej witryny."
    )
    # Zapisujemy przycięty obraz bez paska cookies do tymczasowego bufora nie jest potrzebny w ReportDoc.figure
    # bo figure przyjmuje ścieżkę; możemy przygotować czysty plik lub przekazać oryginalny zrzut.
    # Zadbajmy o spójność z planem: MOWI/zrzut_mowi_com.png w pełnym raporcie (width_cm=11.0).
    report.figure(
        SCREENSHOT,
        "Rys. 1. Zrzut ekranu witryny https://mowi.com/ (stan na 24 listopada 2026 r.). Źródło: opracowanie własne.",
        width_cm=11.0,
    )

    report.h2("1.4. Wyłączenia")
    report.paragraph(
        "W ramach audytu zweryfikowano także obszary witryny, które nie spełniają przesłanek naruszenia w dniu badania i nie zostały włączone do listy przypadków:"
    )
    report.bullet(
        "Brak deklaracji neutralności opartej na offsecie (przyszły pkt 4c): ",
        "Witryna mowi.com nie zawiera twierdzeń o „neutralności klimatycznej”, „zeroemisyjności netto” ani kompensacji emisji projektami leśnymi czy kredytami węglowymi, wobec czego nie ma zastosowania przesłanka kompensacyjna.",
    )
    report.bullet(
        "Brak własnego znaku zrównoważonego charakteru (przyszły pkt 2a): ",
        "Spółka posługuje się rozpoznawalnymi zewnętrznymi programami certyfikacji (ASC, GSA BAP, GLOBALG.A.P., ProTerra, RSPO, MarinTrust) i nie ustanowiła własnej quasi-certyfikacyjnej pieczęci ekologicznej.",
    )
    report.bullet(
        "Zgodność z krajowymi limitami hodowli (wszy morskie i dno morskie): ",
        "Informowanie o celu „0% of sites above national limit (monthly average)” dla wszy morskich lub o przywróceniu stanu dna morskiego zgodnie z przepisami krajowymi nie narusza przyszłego pkt 10a, ponieważ dotyczy limitów prawa krajowego w wybranych jurysdykcjach hodowlanych, a nie wymogu nałożonego prawem na wszystkie produkty danej kategorii w całej Unii Europejskiej.",
    )
    report.bullet(
        "Wzmianka o naturalnej astaksantynie w paszy: ",
        "Zdanie na /products/salmon-farming/, że pasza zawiera naturalne związki takie jak astaksantyna nadające łososiowi różowy kolor, nie zostało zakwalifikowane jako naruszenie w oparciu o sam brak opisu syntezy; pomocniczo w świetle wyroku TSUE z 4 kwietnia 2000 r., C-465/98, Darbo (omówienie na tle wcześniejszej dyrektywy, aktualnie art. 7 rozporządzenia nr 1169/2011; wyrok nie dotyczy Mowi ASA), ocena zależy od rzeczywistego składu i odbioru konsumenta.",
    )
    report.bullet(
        "Umiejscowienie kodów certyfikatów: ",
        "Nie zakwalifikowano jako naruszenia samego braku numeru certyfikatu w tym samym akapicie strony internetowej, ponieważ art. 30 i art. 32 ust. 1 lit. a rozporządzenia (UE) 2018/848 nie ustanawiają takiego obowiązku dla ogólnych opisów na stronie WWW.",
    )
    report.bullet(
        "Wyłączenie raportów PDF: ",
        "Poza zakresem audytu pozostały sprawozdania finansowe, Annual Report 2025, pliki PDF z sekcji ESG Library oraz lokalne wersje językowe witryny (w tym mowi.com/pl/).",
    )

    # 2. Tabela zbiorcza
    report.h1("2. Analiza szczegółowa — tabela zbiorcza naruszeń")
    report.findings(
        [
            (
                c["no"],
                c["short_claim"],
                c["table_basis"],
                c["risk"],
                c["issue"],
            )
            for c in CASES_FULL
        ]
    )

    # 3. Szczegółowa analiza przypadków
    report.h1("3. Szczegółowa analiza przypadków")
    for c in CASES_FULL:
        report.case(f"Przypadek {c['no']}. {c['title']}", c["case_basis"], c["risk"])
        report.paragraph("Cytat źródłowy", size=10.5, bold=True, before=2, after=2, align="left")
        for loc, q in c["quotes"]:
            report.quote(loc, q)
        report.paragraph("Ocena prawna", size=10.5, bold=True, before=6, after=2, align="left")
        for para_text in c["assessment"]:
            report.paragraph(para_text)
        report.paragraph("Kierunek naprawy", size=10.5, bold=True, before=2, after=2, align="left")
        report.paragraph(c["remediation"])
        report.claim(c["claim"])

    # 4. Porównanie
    report.h1("4. Porównanie — weryfikacja twierdzeń porównawczych")
    report.paragraph(
        "Witryna https://mowi.com/ posługuje się zarówno zestawieniami kategorii żywności (żywność morska wobec lądowej), jak i porównaniem wartości odżywczych łososia hodowlanego i dzikiego. "
        "Zgodnie z art. 16 ust. 3 UZNK oraz wyrokiem TSUE z dnia 19 kwietnia 2007 r., C-381/05, De Landtsheer (omówienie tezy; wyrok nie dotyczy Mowi ASA), samo zestawienie kategorii nie przesądza automatycznie o reklamie porównawczej, lecz kwalifikację taką przyjmuje się wówczas, gdy kontekst umożliwia bezpośrednie lub pośrednie zidentyfikowanie konkurencyjnych produktów o tym samym przeznaczeniu. "
        "Warunek ten jest spełniony przy porównaniach łososia z mięsem zwierząt lądowych („terrestrial meats”, „other sources of farmed animal protein”) oraz z dzikim łososiem atlantyckim. "
        "Jak wskazują wyroki TSUE z dnia 19 września 2006 r., C-356/04, Lidl Belgium oraz z dnia 8 lutego 2017 r., C-562/15, Carrefour Hypermarchés (omówienia tez; wyroki nie dotyczą Mowi ASA), porównania kategorii nie są bezwzględnie zakazane i nie wymagają wskazywania jednej normy ISO, jednak muszą być obiektywne, sprawdzalne oraz ujawniać w samym przekazie założenia decydujące o wyniku."
    )

    comp_rows = [
        (
            "Emisje gazów cieplarnianych (kryterium: poziom emisji GHG całej żywności morskiej wobec żywności lądowej): „Blue foods have lower GHG emissions than land-based foods (BFA, 2021).”",
            "Porównanie zniekształcające: obejmuje całą kategorię „land-based foods” (w tym produkty roślinne o niższym śladzie niż ryby drapieżne) oraz przypisuje łososiowi atlantyckiemu średnie wyniki szerokiej grupy „blue foods” (obejmującej małże i algi). Brak jednostki funkcjonalnej i granic cyklu życia (pasza, transport).",
        ),
        (
            "Zużycie wody słodkiej (kryterium: zużycie wody na kg jadalnego mięsa): „salmon production has the lowest water consumption per kilogram of edible meat compared with other sources of farmed animal protein.”",
            "Przekaz niepełny: stanowi reklamę porównawczą identyfikującą hodowlane białko lądowe. Twierdzenie umieszczono w kontekście recyrkulacji RAS w wylęgarniach własnych, bez ujawnienia, czy porównanie uwzględnia ślad wodny upraw rolnych (np. soi i pszenicy) w paszy dla łososia.",
        ),
        (
            "Profil odżywczy (kryterium: zawartość makro- i mikroskładników w 100 g produktu): tabela porównująca „Mowi Salmon” oraz „Wild atlantic salmon” na /products/taste-health/.",
            "Porównanie co do zasady rzetelne i dopuszczalne: podaje konkretne wartości liczbowe w tej samej jednostce (100 g), wskazuje źródło dla dzikiego łososia (IMR seafood data) oraz rzetelnie wykazuje wyższą zawartość EPA+DHA (1,6 g vs 1,3 g) i witaminy D u łososia dzikiego przy wyższej zawartości tłuszczu u łososia hodowlanego. Wymaga jedynie usunięcia towarzyszących bezprawnych oświadczeń chorobowych z Przypadku 01.",
        ),
    ]

    comp_table = report.doc.add_table(rows=1 + len(comp_rows), cols=2)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_table.autofit = False
    c_widths = (Cm(8.5), Cm(8.9))
    set_table_width(comp_table, report.content, c_widths)
    for col_idx, head_txt in enumerate(["Twierdzenie podmiotu (z kryterium porównawczym)", "Rzeczywisty stan prawno-środowiskowy"]):
        cell = comp_table.rows[0].cells[col_idx]
        report._prepare(cell, "08181F", "08181F", "4", 55, 55, 65, 65)
        p = report._cell_para(cell, True, "left")
        set_run(p.add_run(head_txt), 8.5, True, False, WHITE)
    for r_idx, (c_left, c_right) in enumerate(comp_rows):
        fill = "FFFFFF" if r_idx % 2 == 0 else "F7F9F8"
        for col_idx, val in enumerate((c_left, c_right)):
            cell = comp_table.rows[r_idx + 1].cells[col_idx]
            report._prepare(cell, fill, "E4EAE7", "4", 55, 55, 65, 65)
            p = report._cell_para(cell, True, "left")
            set_run(p.add_run(val), 8.5, False, False, DARK)
    report.paragraph("", size=3, after=4, align="left")
    report.paragraph(
        "Wnioski z powyższego zestawienia nie uzasadniają bezwarunkowego usunięcia porównań środowiskowych, lecz ich doprecyzowanie przez zawężenie grupy odniesienia do wskazanych mięs lądowych, podanie jednostki funkcjonalnej oraz ujawnienie granic systemu (łącznie z produkcją paszy)."
    )

    # 5. Ryzyka prawne i finansowe
    report.h1("5. Ryzyka prawne i finansowe")

    report.h2("5.1. Sankcje administracyjne (Prezes UOKiK oraz organy nadzoru żywnościowego)")
    report.paragraph(
        "Stwierdzone w Przypadkach 01–05 naruszenia już obowiązujących przepisów art. 5 i 6 UPNPR oraz bezwzględnych norm prawa żywnościowego UE narażają komunikację skierowaną na rynek polski i unijny na postępowania w sprawie praktyk naruszających zbiorowe interesy konsumentów (art. 24 ust. 1 i 2 pkt 3 u.o.k.k.). "
        "Na podstawie art. 106 ust. 1 pkt 4 u.o.k.k. Prezes UOKiK może nałożyć karę pieniężną w wysokości do 10% obrotu osiągniętego w roku obrotowym poprzedzającym rok nałożenia kary za każdą odrębną praktykę, a także nakazać publikację decyzji na koszt przedsiębiorcy. "
        "Równolegle naruszenia rozporządzeń (WE) nr 1924/2006, (UE) nr 1169/2011 i (UE) 2018/848 podlegają kontroli organów jakości handlowej i bezpieczeństwa żywności (w Polsce IJHARS oraz Państwowa Inspekcja Sanitarna), które mogą wydać nakaz usunięcia nieprawidłowych oświadczeń lub nałożyć kary pieniężne."
    )

    report.h2("5.2. Roszczenia cywilnoprawne konkurentów na podstawie UZNK")
    report.paragraph(
        "Zakwalifikowanie porównań śladu wodnego i emisyjnego z białkiem zwierzęcym jako reklamy porównawczej niespełniającej wymogu rzetelności i kompletności założeń (Przypadek 04) oraz wprowadzających w błąd twierdzeń o ekologiczności soi i 100% certyfikacji (Przypadki 02 i 03) otwiera drogę do roszczeń z art. 18 ust. 1 w zw. z art. 16 UZNK. "
        "Konkurencyjni producenci żywności lub organizacje przedsiębiorców mogą żądać zaniechania niedozwolonych działań, usunięcia ich skutków, złożenia wielokrotnego oświadczenia w mediach, naprawienia wyrządzonej szkody oraz zasądzenia odpowiedniej sumy pieniężnej na określony cel społeczny."
    )

    report.h2("5.3. Ryzyko regulacyjne w horyzoncie EmpCo (27 września 2027 r. w Polsce oraz 27 września 2026 r. w innych państwach UE)")
    report.paragraph(
        "Z dniem 27 września 2027 r. wejdzie w życie polska ustawa z dnia 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353), wprowadzająca bezwzględne zakazy czarnej listy EmpCo (m.in. pkt 4a dla ogólnych haseł „climate-friendly” i „sustainably certified”, pkt 4b dla twierdzeń rozciągających cechę części działalności na całość, pkt 10a dla wymogów ustawowych takich jak PPWR oraz wymóg weryfikowanego planu wdrażania dla celów na 2030 r.). "
        "Ponieważ jednak https://mowi.com/ obsługuje rynek ogólnounijny, w państwach członkowskich UE, które rozpoczęły stosowanie transpozycji dyrektywy (UE) 2024/825 od 27 września 2026 r., organy sieci współpracy ochrony konsumentów (CPC) mogą kwestionować te treści już obecnie bez okresu przejściowego dla publikacji internetowych."
    )

    report.h2("5.4. Ryzyko w łańcuchu dostaw B2B i ścieżka KPK OECD")
    report.paragraph(
        "Niespójności dotyczące certyfikacji soi („organic-certified” vs ProTerra) oraz brak rozbicia udziałów ASC/BAP/GLOBALG.A.P. stwarzają ryzyko zakwestionowania wiarygodności danych przekazywanych odbiorcom sieciowym na potrzeby ich własnych ujawnień zrównoważonego rozwoju (CSRD / ESRS) i procedur zakupowych. "
        "Dodatkowo ogólne deklaracje braku naruszeń praw człowieka (Przypadek 07) bez wskazania zakresu i okresu weryfikacji niosą ryzyko zgłoszenia do właściwego Krajowego Punktu Kontaktowego OECD w trybie Wytycznych OECD (2023) — jest to ścieżka o poziomie ryzyka NISKIE, niebędąca postępowaniem sądowym, której skutkiem są publiczne zalecenia i oddziaływanie wizerunkowe."
    )

    # 6. Rekomendacje naprawcze
    report.h1("6. Rekomendacje naprawcze")
    report.paragraph(
        "Harmonogram działań naprawczych uporządkowano w czterech etapach liczonych od daty audytu (24 listopada 2026 r.), rozdzielając natychmiastowe usunięcie naruszeń prawa już obowiązującego od dostosowań przygotowujących witrynę do standardu EmpCo przed dniem 27 września 2027 r."
    )

    report.h2("6.1. Etap natychmiastowy: usunięcie naruszeń prawa obowiązującego (0–14 dni, do 8 grudnia 2026 r.)")
    report.paragraph(
        "W trybie pilnym należy wycofać ze stron /products/taste-health/, /products/ oraz /sustainability/ wszelkie oświadczenia o zapobieganiu i zmniejszaniu ryzyka chorób (zawał, udar, nagła śmierć, choroby krążenia, choroby psychiczne, zapalenie stawów, nowotwory) oraz powołania na amerykańską FDA, zastępując je wyłącznie dopuszczonymi oświadczeniami z rozporządzenia (UE) nr 432/2012 z obowiązkową informacją z art. 10 ust. 2 rozporządzenia nr 1924/2006 (Przypadek 01). "
        "Równocześnie w artykule z 23 listopada 2026 r. należy usunąć błędną deklarację „all of our soy is from organic-certified sources” i zastąpić ją wskazaniem faktycznych standardów certyfikacji soi (ProTerra lub równoważny), bez przypisywania statusu ekologicznego produkcji konwencjonalnej (Przypadek 02)."
    )

    report.h2("6.2. Korekta materiałów, liczb i ujawnień porównawczych (15–45 dni, do 8 stycznia 2027 r.)")
    report.paragraph(
        "Należy skorygować kafelek „100% of harvest volumes sustainably certified” na /sustainability/ przez wyraźne oznaczenie charakteru docelowego albo podanie aktualnego wyniku za ostatni zamknięty rok w rozbiciu na ASC, GSA BAP i GLOBALG.A.P., zaktualizować pięcioletnie dane ASC z 2021 r. oraz poprawić błędny nagłówek GFSI/GSSI na /key-partnerships/ (Przypadek 03). "
        "Należy zawęzić porównania środowiskowe do nazwanych gatunków mięsa lądowego, dodać jednostkę funkcjonalną oraz granice analizy (z uwzględnieniem wody i emisji z pasz) w miejsce ogólnych odniesień do „land-based foods” i bezwzględnego „lowest water consumption” (Przypadek 04). "
        "Ponadto należy doprecyzować nagłówki rankingu Coller FAIRR, usunąć zapewnienia o „braku negatywnych i trwałych wpływów” z artykułu o dnie morskim (Przypadek 05), naprawić ucięte wartości w kafelkach oraz usunąć sprzeczność „0 programs” na podstronie SDG i rozbieżności liczby pracowników oraz zakładów (Przypadki 06 i 07)."
    )

    report.h2("6.3. Dostosowanie do standardu EmpCo przed wejściem w życie ustawy polskiej (do 30 czerwca 2027 r., przed 27 września 2027 r.)")
    report.paragraph(
        "Przed dniem 27 września 2027 r. (wejście w życie ustawy z 4 września 2026 r., Dz.U. z 2026 r. poz. 1353) należy wyeliminować z całej witryny ogólne hasła „climate-friendly” oraz „sustainably certified” lub bezpośrednio w tym samym kanale przekazu powiązać je z mierzalnymi i zweryfikowanymi parametrami (przyszły pkt 4a i 4b). "
        "W odniesieniu do celów redukcji GHG i zero odpadów na 2030 r. należy opublikować bezpośrednio przy zobowiązaniach link do szczegółowego, realistycznego planu wdrażania z mierzalnymi kamieniami milowymi, alokacją zasobów oraz okresowym raportem weryfikacji przez niezależnego eksperta zewnętrznego, a deklarację dostosowania do PPWR oddzielić od celów dobrowolnych (przyszły pkt 10a)."
    )

    report.h2("6.4. Metodyka, dowody, stałe sformułowania i szkolenie (45–90 dni i tryb ciągły)")
    report.paragraph(
        "Rekomenduje się wdrożenie wewnętrznej macierzy zatwierdzania treści marketingowych i artykułów w dziale News Center (w tym weryfikację spójności z politykami Mowi Feed, danymi działu Sustainability oraz wymogami prawa żywnościowego UE przed każdą publikacją). "
        "Zespoły komunikacji korporacyjnej, marketingu produktowego oraz sprzedaży B2B w całej grupie powinny przejść obowiązkowe szkolenie z rozporządzeń (WE) nr 1924/2006, (UE) nr 1169/2011, (UE) 2018/848, reguł reklamy porównawczej oraz standardów dyrektywy (UE) 2024/825, opierając przyszłe publikacje na zatwierdzonych wzorach sformułowań z niniejszego raportu."
    )

    # 7. Wnioski
    report.h1("7. Wnioski")
    report.paragraph(
        "Według stanu prawnego na dzień 24 listopada 2026 r. witryna https://mowi.com/ narusza już obowiązujące przepisy prawa żywnościowego UE oraz art. 5 i 6 UPNPR i art. 16 UZNK — w szczególności poprzez niedozwolone oświadczenia chorobowe dotyczące kwasów omega-3, sprzeczne twierdzenie o „organic-certified” soi oraz prezentowanie celu 100% certyfikacji i porównań środowiskowych bez istotnych założeń — co uzasadnia wstrzymanie publikacji kwestionowanych treści i ich niezwłoczną korektę. "
        "Natomiast nowe zakazy czarnej listy dyrektywy (UE) 2024/825 (EmpCo) wchodzą w życie w Polsce dopiero 27 września 2027 r. (Dz.U. z 2026 r. poz. 1353), stanowiąc w dniu audytu przyszły standard regulacyjny wymagający przygotowania planów weryfikacji celów na 2030 r. oraz usunięcia ogólnych haseł ekologicznych przed tą datą."
    )

    report.sign_off()
    OUT_FULL.parent.mkdir(parents=True, exist_ok=True)
    report.save(str(OUT_FULL))
    return OUT_FULL


# ==============================================================================
# 4. RAPORT SKRÓCONY (wzór: wzory/WZOR_raport_skrocony.docx i generuj_raport_skrocony.py)
# ==============================================================================

BLUE = RGBColor(58, 97, 179)
CYAN = RGBColor(47, 141, 188)
CYAN_SOFT = RGBColor(142, 197, 224)
DANGER = RGBColor(190, 18, 60)
WARNING = RGBColor(180, 83, 9)
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


def _srun(run, size, bold=False, italic=False, color=DARK, font="Arial", tracking=None, lang="pl-PL"):
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
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:w"), str(sum(dxas)))
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
        for cell, dxa in zip(row.cells, dxas):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(dxa))
            tc_w.set(qn("w:type"), "dxa")


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
    _srun(run, size, False, False, color)


def _p_border_bottom(paragraph, color="D5DDE3", sz="12", space="8"):
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


def _cell_p(cell, clear=True):
    paragraph = cell.paragraphs[0]
    if clear:
        _clear_paragraph(paragraph)
    return paragraph


def build_short_report():
    """Generuje raport skrócony audytu greenwashingu dla MOWI ASA (mowi.com)."""
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
    props.subject = "Raport skrócony audytu ryzyk greenwashingu — MOWI ASA (mowi.com)"
    props.title = "Raport skrócony audytu ryzyk greenwashingu — MOWI ASA"
    props.comments = "ESG Institute · esginstitute.eu · raport skrócony · 24 listopada 2026 r."

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
        _srun(paragraph.add_run(text), size, bold, italic, color, lang=lang, tracking=tracking)
        return paragraph

    def rich(parts, size=10.5, before=0, after=8, align="left", keep_next=False, keep_lines=False, space=1.15):
        paragraph = doc.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(before)
        fmt.space_after = Pt(after)
        fmt.line_spacing = space
        _widow(paragraph, keep_next, keep_lines, False)
        for text, bold, italic, color, lang in parts:
            _srun(paragraph.add_run(text), size, bold, italic, color, lang=lang)
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
        fmt.space_before = Pt(0 if new_page else 16)
        fmt.space_after = Pt(8)
        fmt.line_spacing = 1.05
        _widow(paragraph, keep_next=True, keep_lines=True, page_break=new_page)
        _srun(paragraph.add_run(number + ".  " + title), 14, True, False, NAVY)
        _p_border_bottom(paragraph, "D5DDE3", "12", "8")
        return paragraph

    def label(text):
        return para(text, size=9, bold=True, color=BLUE, before=12, after=5, keep_next=True, keep_lines=True, tracking=32)

    def band(fill, left_color, paragraphs, gap_after=8):
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
            _srun(p.add_run(text), size, bold, italic, color, lang=lang)
        if body_cell.paragraphs:
            body_cell.paragraphs[-1].paragraph_format.space_after = Pt(0)
        _cant_split(table.rows[0])
        gap = doc.add_paragraph()
        _clear_paragraph(gap)
        gap.paragraph_format.space_after = Pt(gap_after)

    # --- Nagłówek i stopka ---
    logo_path = default_logo()
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
    if logo_path and logo_path.exists():
        lp.add_run().add_picture(str(logo_path), width=Cm(4.15))
    else:
        _srun(lp.add_run("ESG INSTITUTE"), 11, True, False, WHITE)
    rp = _cell_p(right)
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _srun(rp.add_run("Raport skrócony"), 8, True, False, CYAN_SOFT)
    rp2 = right.add_paragraph()
    _clear_paragraph(rp2)
    rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _srun(rp2.add_run("MOWI ASA"), 8, True, False, WHITE)
    _srun(rp2.add_run("   ·   esginstitute.eu"), 8, False, False, CYAN_SOFT)

    stripe_a, stripe_b = htable.rows[1].cells
    stripe_a.merge(stripe_b)
    stripe = htable.rows[1].cells[0]
    _shd(stripe, "2F8DBC")
    _mar(stripe, 0, 0, 0, 0)
    _nil_cell_borders(stripe)
    _clear_paragraph(stripe.paragraphs[0])
    _row_height(htable.rows[1], 70, "exact")

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
    _srun(fp.add_run("ESG Institute Sp. z o.o.   ·   esginstitute.eu   ·   raport skrócony   ·   MOWI ASA"), 8, False, False, MUTED)
    _srun(fp.add_run("\t"), 8, False, False, MUTED)
    _field(fp, "PAGE")
    _srun(fp.add_run("  /  "), 8, False, False, MUTED)
    _field(fp, "NUMPAGES")

    # --- Strona 1 ---
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
        "Audyt Ryzyk Greenwashingu: MOWI ASA",
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
        "Raport skrócony. Weryfikacja publicznej witryny https://mowi.com/ (strona główna, Produkty, Mowi Feed, Zrównoważony rozwój), stan na 24 listopada 2026 r. "
        "Werdykt opiera się na obowiązującym prawie żywnościowym UE (rozp. 1924/2006, 432/2012, 1169/2011, 2018/848), UPNPR i UZNK; "
        "czarna lista dyrektywy (UE) 2024/825 (EmpCo) wchodzi w życie w Polsce 27 września 2027 r.",
        size=10.5,
        italic=True,
        color=MUTED,
        before=2,
        after=10,
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
    fpict.add_run().add_picture(clean_screenshot_stream(crop_bottom=780), width=Cm(PAGE_W - 0.7))
    cap = fcell.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(0)
    cap.paragraph_format.line_spacing = 1.05
    _srun(
        cap.add_run(
            "Fragment strony głównej https://mowi.com/ z 24 listopada 2026 r. Witryna komunikuje dostarczanie 9,7 miliona posiłków dziennie "
            "oraz kieruje przekazy produktowe, zdrowotne i środowiskowe do konsumentów i partnerów B2B."
        ),
        8,
        False,
        True,
        MUTED,
    )
    after_frame = doc.add_paragraph()
    _clear_paragraph(after_frame)
    after_frame.paragraph_format.space_after = Pt(10)

    meta = [
        (
            "Badany podmiot",
            "Mowi ASA, Sandviksboder 77AB, 5035 Bergen, Norwegia (PO Box 4102 Sandviken, 5835 Bergen). Spółka notowana na Oslo Stock Exchange, obecna w 26 krajach, w tym w Polsce przez Mowi Poland S.A.",
        ),
        (
            "Przedmiot audytu",
            "Globalna witryna https://mowi.com/ (stan na 24.11.2026 r.): strona główna, /about-us/, /products/ (w tym /taste-health/, /salmon-farming/), /about-us/our-structure/mowi-feed/ oraz /sustainability/ (/commitments/, /certifications/, /how-mowi-uses-freshwater-responsibly/).",
        ),
        (
            "Sektor działalności",
            "Akwakultura, produkcja pasz, przetwórstwo i sprzedaż łososia atlantyckiego oraz owoców morza (sektor spożywczy / FMCG; kanały retail, HoReCa i B2B).",
        ),
        (
            "Reżim prawny",
            "Prawo obowiązujące: rozporządzenia (WE) nr 1924/2006, (UE) nr 432/2012, (UE) nr 1169/2011, (UE) 2018/848; UPNPR (Dz.U. z 2023 r. poz. 845); UZNK (Dz.U. z 2026 r. poz. 85); u.o.k.k. (Dz.U. z 2025 r. poz. 1714). "
            "Ustawa z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) wchodzi w życie 27 września 2027 r. — w dniu audytu czarna lista EmpCo nie jest jeszcze prawem polskim.",
        ),
        ("Data audytu", AUDIT_DATE),
        ("Werdykt operacyjny", VERDICT_RED),
    ]
    meta_table = doc.add_table(rows=len(meta), cols=2)
    meta_w = [3.9, PAGE_W - 3.9]
    _lock(meta_table, meta_w)
    _nil_table_borders(meta_table)
    for i, (m_field, m_val) in enumerate(meta):
        ck, cv = meta_table.rows[i].cells
        last = i == len(meta) - 1
        _shd(ck, "F4F7F8")
        _shd(cv, "FFFFFF")
        _mar(ck, 60, 60, 120, 80)
        _mar(cv, 60, 60, 120, 120)
        _borders(ck, "E4E8EC", "4")
        _borders(cv, "E4E8EC", "4")
        ck.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pk, pv = _cell_p(ck), _cell_p(cv)
        pk.paragraph_format.space_after = Pt(0)
        pv.paragraph_format.space_after = Pt(0)
        pv.paragraph_format.line_spacing = 1.12
        pv.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _srun(pk.add_run(m_field), 8, True, False, NAVY)
        _srun(pv.add_run(m_val), 8.5, last, False, DANGER if last else DARK)
        _cant_split(meta_table.rows[i])

    # --- Strona 2 ---
    start_page()
    band(
        "FDF2F4",
        "BE123C",
        [
            ("Werdykt. " + VERDICT_RED + ".", 12, True, False, NAVY, "pl-PL"),
            (
                "Witryna https://mowi.com/ narusza już dziś bezwzględnie obowiązujące przepisy prawa żywnościowego UE (art. 7 ust. 3–4 rozporządzenia nr 1169/2011, art. 10 i 14 rozporządzenia nr 1924/2006 oraz art. 30 rozporządzenia 2018/848), a także art. 5 i 6 UPNPR oraz art. 16 UZNK. "
                "Na podstronach produktowych przypisano kwasom omega-3 i łososiowi właściwości zapobiegania chorobom układu krążenia, zawałowi, udarowi i nowotworom, a w artykule z 23 listopada 2026 r. zadeklarowano, że cała soja pochodzi ze źródeł „organic-certified” wbrew certyfikacji ProTerra i konwencjonalnej produkcji pasz Mowi Feed.",
                10.5,
                False,
                False,
                DARK,
                "pl-PL",
            ),
            (
                "Należy rozdzielić dwie daty regulacyjne: art. 4 ust. 1 dyrektywy (UE) 2024/825 (EmpCo) wyznaczył 27 września 2026 r. dla stosowania transpozycji w państwach UE, natomiast polska ustawa z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) wchodzi w życie 27 września 2027 r. W dniu audytu czarna lista EmpCo opisuje przyszły standard regulacyjny w Polsce i nie stanowi jeszcze naruszenia krajowej czarnej listy.",
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
        "Mowi ASA prezentuje się na globalnej witrynie https://mowi.com/ jako największy na świecie producent łososia atlantyckiego (559 000 ton zbiorów i 5,7 mld EUR przychodów w 2025 r., 9,7 miliona posiłków dziennie) oraz „najbardziej zrównoważony producent białka”. "
        "Zintegrowany model grupy obejmuje własną genetykę, produkcję pasz (Mowi Feed), hodowlę (Mowi Farming) i sprzedaż w ponad 70 krajach do konsumentów oraz odbiorców z kanału detalicznego i HoReCa."
    )
    body(
        "Przegląd witryny według stanu na 24 listopada 2026 r. wykazał siedem obszarów ryzyka prawnego, z których trzy kluczowe rozstrzygają o czerwonym werdykcie operacyjnym już na gruncie obowiązujących rozporządzeń żywnościowych UE, UPNPR oraz UZNK. "
        "Hasła ogólne („climate-friendly”, „sustainably certified”) i cele na 2030 r. wymagają natomiast dodatkowo dostosowania do przyszłego standardu dyrektywy (UE) 2024/825 przed 27 września 2027 r."
    )

    heading("2", "Ustalenia")
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota"]
    col_w = [1.0, 4.35, 3.85, 2.85, 4.95]
    short_cases = CASES_FULL[:3]
    sum_table = doc.add_table(rows=1 + len(short_cases), cols=5)
    _lock(sum_table, col_w)
    _nil_table_borders(sum_table)
    for c_idx, head in enumerate(headers):
        cell = sum_table.rows[0].cells[c_idx]
        _shd(cell, "08181F")
        _mar(cell, 70, 70, 70, 70)
        _borders(cell, "08181F", "4", "4", "08181F")
        p = _cell_p(cell)
        if c_idx in (0, 3):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _srun(p.add_run(head), 8, True, False, WHITE)
    risk_color = {"BARDZO WYSOKIE": DANGER, "WYSOKIE": WARNING}
    for r_idx, c_item in enumerate(short_cases):
        bg = "FFFFFF" if r_idx % 2 == 0 else "F7F9FC"
        row_vals = (c_item["no"], c_item["short_claim"], c_item["short_basis"], c_item["risk"], c_item["issue"])
        for c_idx, val in enumerate(row_vals):
            cell = sum_table.rows[r_idx + 1].cells[c_idx]
            _shd(cell, bg)
            _mar(cell, 70, 70, 70, 70)
            _borders(cell, "E4E8EC", "4")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = _cell_p(cell)
            p.paragraph_format.line_spacing = 1.12
            if c_idx in (0, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if c_idx == 0:
                _srun(p.add_run(val), 8.5, True, False, NAVY)
            elif c_idx == 3:
                _srun(p.add_run(val), 8, True, False, risk_color.get(val, DARK))
            else:
                _srun(p.add_run(val), 8, False, False, DARK)
        _cant_split(sum_table.rows[r_idx + 1])
    _cant_split(sum_table.rows[0])
    sum_table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    para(
        "Zestawienie prezentuje 3 kluczowe ryzyka spośród 7 zidentyfikowanych w pełnym audycie witryny mowi.com (stan na 24 listopada 2026 r.).",
        size=10,
        italic=True,
        color=MUTED,
        before=10,
        after=2,
    )

    # --- Strona 3: Szczegóły przypadku 01 ---
    heading("3", "Szczegóły przypadków", new_page=True)
    case0 = CASES_FULL[0]
    para(
        "Przypadek " + case0["no"] + ".  " + case0["title"],
        size=13,
        bold=True,
        color=NAVY,
        before=10,
        after=4,
        keep_next=True,
        keep_lines=True,
        space=1.08,
    )
    rich(
        [
            ("Ryzyko: ", False, False, MUTED, "pl-PL"),
            (case0["risk"].lower() + ".", True, False, DANGER, "pl-PL"),
        ],
        size=10.5,
        before=0,
        after=6,
        keep_next=True,
        keep_lines=True,
    )
    body("Podstawa prawna: " + case0["case_basis"] + ".", size=9.5, before=0, after=6, first=0, keep_lines=True)

    label("Cytat źródłowy")
    for _, q in case0["quotes"][:4]:
        band("F4F7FB", "2F8DBC", [(q, 10, False, True, DARK, "en-US")], gap_after=5)

    label("Ocena prawna")
    for block in case0["assessment"][:3]:
        body(block, before=1, after=7, keep_lines=True, first=0.75)

    # --- Strona 4: Kierunek naprawy i Dozwolone sformułowanie dla Przypadku 01 ---
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
    body(case0["remediation"], before=4, after=8, keep_lines=True, first=0.75)

    label("Dozwolone sformułowanie")
    band("E8F0FA", "3A61B3", [(case0["claim"], 10, False, True, DARK, "en-US")], gap_after=4)

    # --- Strona 5: Pozostałe przypadki + CTA ---
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
    for other in CASES_FULL[1:3]:
        other_risk = DANGER if "BARDZO" in other["risk"] else WARNING
        summary_line = (
            "Artykuł z 23.11.2026 r. deklaruje, że cała soja pochodzi ze źródeł „organic-certified”, podczas gdy podstrona paszowa wskazuje certyfikację ProTerra i produkcję konwencjonalną (naruszenie art. 30 rozp. 2018/848, art. 5 UPNPR i art. 16 UZNK)."
            if other["no"] == "02"
            else "Kafelek na /sustainability/ przedstawia cel 100% certyfikacji GSSI jako stan osiągnięty przy danych ASC z 2021 r. (50% farm), bez rozbicia udziału ASC/BAP/GLOBALG.A.P. i z błędem GFSI/GSSI."
        )
        band(
            "F7F9FC",
            "08181F",
            [
                (other["no"] + "     " + other["title"], 11, True, False, NAVY, "pl-PL"),
                ("Ryzyko: " + other["risk"].lower() + ".", 10, True, False, other_risk, "pl-PL"),
                (summary_line, 10, False, False, DARK, "pl-PL"),
            ],
            gap_after=8,
        )

    band(
        "F4F7FB",
        "3A61B3",
        [
            ("Reszta po zakupie pełnej wersji.", 11, True, False, NAVY, "pl-PL"),
            (
                "Pełna analiza wszystkich 7 przypadków naruszeń (w tym porównań emisyjnych i wodnych, rankingu Coller FAIRR, celów na 2030 r., PPWR i SDG), "
                "weryfikacja porównawcza, analiza ryzyk prawno-finansowych oraz czterostopniowy harmonogram działań naprawczych przed 27 września 2027 r. dostępne są w pełnym raporcie audytu ESG Institute.",
                10.5,
                False,
                False,
                DARK,
                "pl-PL",
            ),
        ],
        gap_after=0,
    )

    OUT_SHORT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT_SHORT))
    return OUT_SHORT


def main():
    p_mat = build_material()
    p_full = build_full_report()
    p_short = build_short_report()
    print("Wygenerowano:")
    print(" -", p_mat)
    print(" -", p_full)
    print(" -", p_short)


if __name__ == "__main__":
    main()
