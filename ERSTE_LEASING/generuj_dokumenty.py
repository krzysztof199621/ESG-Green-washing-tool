#!/usr/bin/env python3
"""Generator dokumentów audytu ryzyk greenwashingu dla Erste Leasing S.A.

Tworzy trzy pliki w katalogu ERSTE_LEASING/:
1. Material_Erste_Leasing_tresci_do_audytu.docx — czysty zbiór materiałów źródłowych.
2. Raport_Audytu_Greenwashingu_Erste_Leasing.docx — pełny raport (klasa ReportDoc z sklad_raportu.py).
3. Raport_Skrocony_Audytu_Greenwashingu_Erste_Leasing.docx — raport skrócony (układ wg WZOR_raport_skrocony.docx).

Stan witryny i data audytu: 24 listopada 2026 r.
Sygnatura: Sygn. ESG/2026/11/24.
"""

from io import BytesIO
from pathlib import Path
import sys

import docx
from PIL import Image
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
TOOL = ROOT.parent
if str(TOOL) not in sys.path:
    sys.path.insert(0, str(TOOL))

from sklad_raportu import ReportDoc, VERDICT_RED, default_logo  # noqa: E402

SHOT_HOME = ROOT / "zrodla" / "erste_leasing_home.png"
SHOT_ZIELONE = ROOT / "zrodla" / "erste_leasing_zielone_rozwiazania.png"

MATERIAL_OUT = ROOT / "Material_Erste_Leasing_tresci_do_audytu.docx"
FULL_OUT = ROOT / "Raport_Audytu_Greenwashingu_Erste_Leasing.docx"
SHORT_OUT = ROOT / "Raport_Skrocony_Audytu_Greenwashingu_Erste_Leasing.docx"

AUDIT_DATE = "24 listopada 2026 r."
REFERENCE = "Sygn. ESG/2026/11/24"


# ---------------------------------------------------------------------------
# Wspólne dane merytoryczne audytu (bez pierwszej osoby, zwięzły styl prawny)
# ---------------------------------------------------------------------------

FINDINGS = [
    (
        "01",
        "Kwalifikacja pojazdów spalinowych, hybrydowych, LNG/CNG oraz nieograniczonych emisyjnie maszyn jako „Zielone rozwiązania” i transportu „bez zanieczyszczania powietrza spalinami”",
        "Art. 5 ust. 1, ust. 2 pkt 1 i 2, ust. 3 pkt 2 i ust. 4 oraz art. 6 ust. 1 i ust. 3 pkt 1 UPNPR; art. 16 ust. 1 pkt 2 i ust. 2 UZNK. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a i 4b UPNPR (Dz.U. z 2026 r. poz. 1353).",
        "WYSOKIE",
        "Pod szyldem „Zielone rozwiązania” i „floty niespalinowe” promowana jest ogólna oferta Multisalon24, pojazdy LNG/CNG, hybrydy, ciężarówki i klimatyzatory bez progów emisyjnych i klas energetycznych.",
    ),
    (
        "02",
        "Hasła „Wspólnie jesteśmy BIO” / „Naturalnie, że podpisane!”, etykiety „eko” dla narzędzi cyfrowych oraz przelicznik „50 drzew / tlen dla ok. 100 osób” bez metodyki i z pominięciem dojazdu doradcy oraz KSeF",
        "Art. 5 ust. 1, ust. 2 pkt 1 i 2, ust. 3 pkt 2 i ust. 4 oraz art. 6 ust. 1 UPNPR; art. 16 ust. 1 pkt 2 i ust. 2 UZNK. Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 2a, 4a i 4b UPNPR (Dz.U. z 2026 r. poz. 1353).",
        "WYSOKIE",
        "Ekwiwalent drzew i tlenu bez źródeł i z pominięciem śladu węglowego stacjonarnego spotkania z doradcą przy BioPodpisie oraz prezentacja eFaktury jako dobrowolnej akcji „eko” mimo wdrożenia KSeF.",
    ),
    (
        "03",
        "Deklaracja „precyzyjnego mierzenia »zielonego wpływu« portfela” (System Oznaczania i Weryfikacji Transakcji Green w kontekście CSRD i Taksonomii UE) bez metodyki i skali portfela",
        "Art. 16 ust. 1 pkt 2 i ust. 2 UZNK; art. 5 ust. 2 pkt 2 i art. 6 ust. 1 UPNPR; kontekst art. 3 i 8 rozp. (UE) 2020/852 (Taksonomia UE). Przyszły standard EmpCo od 27.09.2027: art. 7 pkt 4a i 4b UPNPR.",
        "ŚREDNIE",
        "Powołanie autorskiego systemu oceny zielonych transakcji i CSRD bez podania kryteriów kwalifikacji i udziału w portfelu, podczas gdy raport Grupy za 2025 r. wyłącza leasing z kalkulacji PCAF i wykazuje GAR 3,28%–3,56%.",
    ),
    (
        "04",
        "Przedstawianie ustawowego obowiązku wpisu do rejestru BDO jako wyrazu „świadomej dbałości o środowisko” oraz niespójność numeru rejestrowego BDO",
        "Art. 5 ust. 2 pkt 2 i ust. 4 oraz art. 6 ust. 1 UPNPR; art. 16 ust. 1 pkt 2 i ust. 2 UZNK w zw. z art. 50 ust. 1 i art. 63 ustawy o odpadach. Zastrzeżenie co do przyszłego pkt 10a dyrektywy 2024/825 w uzasadnieniu.",
        "ŚREDNIE",
        "Ustawowy obowiązek rejestracji w BDO ujęty jako dobrowolna troska o środowisko, połączony z błędnym 8-cyfrowym zapisem numeru w FAQ wobec 9-cyfrowego numeru w stopce.",
    ),
]

CASES = [
    {
        "no": "01",
        "title": "Kwalifikacja pojazdów spalinowych, hybrydowych, LNG/CNG oraz maszyn bez progów emisyjnych jako „Zielone rozwiązania” i transportu „bez zanieczyszczania powietrza spalinami”",
        "short_title": "Pojazdy spalinowe, hybrydowe i LNG/CNG pod hasłem „Zielone rozwiązania” i „floty niespalinowe”",
        "basis": (
            "Art. 5 ust. 1, ust. 2 pkt 1 i 2, ust. 3 pkt 2 i ust. 4 oraz art. 6 ust. 1 i ust. 3 pkt 1 "
            "ustawy z dnia 23 sierpnia 2007 r. o przeciwdziałaniu nieuczciwym praktykom rynkowym "
            "(t.j. Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 i ust. 2 ustawy z dnia 16 kwietnia 1993 r. "
            "o zwalczaniu nieuczciwej konkurencji (t.j. Dz.U. z 2026 r. poz. 85). "
            "Przyszły standard EmpCo od 27 września 2027 r. (w dniu audytu jeszcze nie naruszenie krajowej czarnej listy): "
            "art. 7 pkt 4a i pkt 4b UPNPR w brzmieniu ustawy z dnia 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353)."
        ),
        "risk": "WYSOKIE",
        "quotes": [
            (
                "Podstrona /zielone-rozwiazania (lista przedmiotów):",
                "„Wybrane przedmioty, które możesz u nas sfinansować: Ciągniki siodłowe i balastowe, Pojazdy hybrydowe, Samochody elektryczne, (…) Samochody ciężarowe, (…) Urządzenia i aparaty chłodnicze, Urządzenia wentylacyjne, Urządzenia klimatyzacyjne, (…) Wozy asenizacyjne z aplikatorami, (…) Rozrzutniki obornika.”",
            ),
            (
                "Podstrony /zielone-rozwiazania oraz /multisalon24-szeroki-wybor-aut-od-autoryzowanych-dealerow:",
                "„Multisalon24 to miejsce, gdzie znajdziesz bogatą ofertę pojazdów – również elektrycznych i hybrydowych – wyłącznie od autoryzowanych dealerów (…) Szeroki wybór - setki pojazdów różnych marek i modeli – zarówno nowych, jak i używanych.”",
            ),
            (
                "Podstrona /zielone-rozwiazania/leasing-i-pozyczka-na-niskoemisyjny-transport:",
                "„Szukasz sposobu, żeby Twoja firma jeździła nowocześnie, ekonomicznie i bez zanieczyszczania powietrza spalinami? (…) finansujemy floty niespalinowe zarówno dla prywatnych firm, jak i sektora publicznego.” Oraz niżej na tej samej stronie: „Lekkie pojazdy niskoemisyjne – hybrydy typu plug-in, lekkie pojazdy użytkowe i specjalistyczne z ograniczoną emisją CO₂. Transport ciężki nowej generacji – pojazdy zasilane LNG, CNG, wodorem lub w technologii hybrydowej (…).”",
            ),
            (
                "Podstrona /zielone-rozwiazania/pozyczka-na-zielone-maszyny-rolnicze:",
                "„Niezależnie od wyboru pojazdu, innowacyjna technologia daje przewagę: pracujesz wydajniej, oszczędzasz paliwo i zostawiasz konkurencję daleko w tyle – a środowisko tylko na tym zyskuje.”",
            ),
        ],
        "assessment": [
            (
                "Zestawienie tytułu działu „Zielone rozwiązania” i podstrony o niskoemisyjnym transporcie z katalogiem finansowanych przedmiotów ujawnia wewnętrzną sprzeczność przekazu. "
                "Nagłówek zachęca do jazdy „bez zanieczyszczania powietrza spalinami” oraz zapowiada finansowanie „flot niespalinowych”, podczas gdy w wykazie tej samej podstrony znajdują się pojazdy hybrydowe typu plug-in oraz ciężarówki zasilane skroplonym i sprężonym gazem ziemnym (LNG i CNG). "
                "Silniki spalinowe zasilane gazem ziemnym lub układy hybrydowe spalają paliwa kopalne i emitują spaliny w trakcie eksploatacji. "
                "Co więcej, zarówno w sekcji „Dla biznesu”, jak i w sekcji „Dla Ciebie” umieszczono ogólny portal Multisalon24, w którym pojazdy elektryczne i hybrydowe są jedynie częścią szerszego katalogu aut spalinowych."
            ),
            (
                "W całej sekcji brak jest jakichkolwiek progów technicznych pozwalających zweryfikować, kiedy dany środek trwały kwalifikuje się jako „zielony” lub „wydajny energetycznie”. "
                "Dla ciągników siodłowych, samochodów ciężarowych, urządzeń klimatyzacyjnych i maszyn rolniczych nie podano norm emisji spalin (g CO₂/km), klas efektywności energetycznej ani odesłania do kryteriów Taksonomii UE. "
                "Sformułowanie „bez zanieczyszczania powietrza spalinami” przypisane bez zastrzeżeń także pojazdom elektrycznym pomija emisje pośrednie związane z wytwarzaniem energii elektrycznej w polskim miksie energetycznym oraz ślad węglowy cyklu życia pojazdu. "
                "W krajowej praktyce nadzorczej Prezes UOKiK w postępowaniach z lipca 2025 r. oraz stycznia 2026 r. (m.in. w sprawach DPD i Bolt — opis praktyki organu, nie wyrok sądowy) kwestionował posługiwanie się bezwarunkowymi hasłami „zeroemisyjności” przy pominięciu źródła zasilania i pełnego cyklu życia floty."
            ),
            (
                "Przekaz spełnia kumulatywne przesłanki działania wprowadzającego w błąd oraz zaniechania wprowadzającego w błąd z art. 5 i art. 6 UPNPR w relacji z konsumentami, a także art. 16 ust. 1 pkt 2 i ust. 2 UZNK w relacji z przedsiębiorcami. "
                "Oferta „Zielone rozwiązania” obejmuje wprost osoby fizyczne nieprowadzące działalności gospodarczej (sekcja „Dla Ciebie”, finansowanie aut dla osób zatrudnionych na umowę o pracę, rencistów i emerytów), a także przedsiębiorców deklarujących troskę o zrównoważony rozwój i poszukujących finansowania niskoemisyjnego. "
                "Zgodnie z omówieniem tezy wyroku TSUE z 19 grudnia 2013 r., C-281/12 (Trento Sviluppo i Centrale Adriatica; wyrok nie dotyczy Erste Leasing S.A.), działanie wprowadzające w błąd zachodzi wówczas, gdy mylący przekaz może spowodować podjęcie przez przeciętnego odbiorcę decyzji dotyczącej umowy, której inaczej by nie podjął. "
                "Rozważając odbiór całości strony, należy uwzględnić również linię orzeczniczą wynikającą z omówienia wyroków TSUE z 4 czerwca 2015 r., C-195/14 (Teekanne) oraz Sądu Najwyższego z 4 marca 2014 r., III SK 34/13 i z 23 października 2019 r., I NSK 66/18 (w którym SN uchylił zaskarżony wyrok i przekazał sprawę do ponownego rozpoznania; żadne z tych orzeczeń nie dotyczy badanego podmiotu): dominujące hasło tytułowe i układ graficzny mogą wywoływać mylne wrażenie nawet wtedy, gdy w dalszej części tekstu pojawiają się zastrzeżenia szczegółowe."
            ),
            (
                "Niezależnie od oceny na gruncie przepisów już obowiązujących, komunikacja ta znajduje się w bezpośrednim obszarze przyszłego standardu dyrektywy (UE) 2024/825, który w Polsce wejdzie w życie 27 września 2027 r. na mocy ustawy z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353). "
                "Od tej daty ogólne sformułowania środowiskowe („zielone rozwiązania”, „zielone maszyny”, „finansowanie bliższe naturze”) bez uznanej wysokiej efektywności ekologicznej w rozumieniu motywu 10 dyrektywy 2024/825 zostaną objęte przyszłym zakazem z art. 7 pkt 4a UPNPR, przy czym posiadany przez spółkę certyfikat ISO 14001:2015 odnosi się do systemu zarządzania środowiskowego organizacji, a nie do produktu lub usługi finansowej. "
                "Ponadto przypisywanie całej ofercie portalu Multisalon24 albo pełnej gamie maszyn cechy „zielonej”, mimo że jedynie wycinek oferty dotyczy pojazdów zeroemisyjnych lub rozwiązań ograniczających zużycie paliwa, będzie podlegało ocenie na gruncie przyszłego pkt 4b."
            ),
        ],
        "remediation": (
            "Należy wycofać z sekcji „Zielone rozwiązania” ogólne odesłania do pełnej oferty Multisalon24 oraz usunąć sformułowania „bez zanieczyszczania powietrza spalinami” i „floty niespalinowe” z podstron obejmujących pojazdy LNG, CNG, hybrydy i ciągniki spalinowe. "
            "Każdą kategorię przedmiotów w dziale finansowania transformacji należy powiązać z konkretnym, weryfikowalnym kryterium technicznym (np. pojazdy wyłącznie elektryczne lub wodorowe na etapie jazdy, klasa efektywności energetycznej urządzeń chłodniczych i klimatyzacyjnych, próg emisji g CO₂/km lub wymogi właściwego programu NFOŚiGW). "
            "Przy pojazdach elektrycznych należy precyzyjnie zaznaczyć, że brak emisji spalin dotyczy wyłącznie etapu użytkowania pojazdu (tank-to-wheel)."
        ),
        "claim": (
            "„Finansowanie pojazdów bezemisyjnych w miejscu użytkowania (BEV/FCEV) oraz wybranych inwestycji ograniczających zużycie energii. "
            "Kwalifikacja przedmiotu do oferty odbywa się według jawnych kryteriów technicznych: wyłącznie napędy elektryczne lub wodorowe (brak emisji spalin z rury wydechowej w trakcie jazdy; pełny ślad węglowy zależy od źródła energii elektrycznej) "
            "albo urządzenia spełniające klasę efektywności energetycznej [X] lub kryteria programu dotacyjnego NFOŚiGW. "
            "Pojazdy hybrydowe, spalinowe oraz zasilane LNG/CNG finansowane są w ramach standardowej oferty transportowej.”"
        ),
    },
    {
        "no": "02",
        "title": "Hasła „Wspólnie jesteśmy BIO” / „Naturalnie, że podpisane!”, etykiety „eko” dla narzędzi cyfrowych oraz przelicznik „50 drzew / tlen dla ok. 100 osób” bez metodyki i z pominięciem dojazdu doradcy oraz KSeF",
        "short_title": "Hasła „BIO”/„EKO”, etykiety „eko” i ekwiwalent „50 drzew / tlen dla 100 osób”",
        "basis": (
            "Art. 5 ust. 1, ust. 2 pkt 1 i 2, ust. 3 pkt 2 i ust. 4 oraz art. 6 ust. 1 ustawy z dnia 23 sierpnia 2007 r. "
            "o przeciwdziałaniu nieuczciwym praktykom rynkowym (t.j. Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 i ust. 2 "
            "ustawy z dnia 16 kwietnia 1993 r. o zwalczaniu nieuczciwej konkurencji (t.j. Dz.U. z 2026 r. poz. 85). "
            "Przyszły standard EmpCo od 27 września 2027 r. (jeszcze nie naruszenie krajowej czarnej listy): "
            "art. 7 pkt 2a, pkt 4a i pkt 4b UPNPR w brzmieniu ustawy z dnia 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353)."
        ),
        "risk": "WYSOKIE",
        "quotes": [
            (
                "Strona główna /erste-leasing, menu główne „Cyfrowe rozwiązania”:",
                "Zielone etykiety graficzne z napisem „eko” umieszczone przy czterech pozycjach menu: „Leasing bySMS eko”, „Jednorazowy podpis kwalifikowany OTF eko”, „BioPodpis eko”, „eFaktura eko”.",
            ),
            (
                "Podstrona /erste-leasing/biopodpis:",
                "„Naturalnie, że podpisane! Wspólnie jesteśmy BIO. Zmniejsz koszty związane z wydrukiem dokumentów i zużyj mniej papieru!”; „eliminuje zużycie papieru”; oraz: „Czy wiesz, że rocznie archiwizujemy ponad 40 000 umów? Zużywamy na to ponad pół miliona kartek! By wyprodukować taką ilość papieru, potrzeba aż 50 drzew! Tyle drzew może zapewnić tlen na rok dla ok. 100 osób.”",
            ),
            (
                "Podstrona /erste-leasing/biopodpis (opis procesu zawarcia umowy):",
                "„Jeśli wszystko się zgadza, doradca leasingowy umówi się z Tobą na spotkanie (…) Doradca leasingowy przyjedzie z kompletnym rozwiązaniem, rysikiem i laptopem.”",
            ),
            (
                "Podstrona /erste-leasing/e-faktura:",
                "„wspólnie jesteśmy EKO – to jest nasz wspólny wkład w ochronę środowiska.” (podczas gdy sekcja KSeF na stronie głównej i podstrona /krajowy-system-e-faktur informują o wdrożeniu obowiązkowego Krajowego Systemu e-Faktur od 1 lutego 2026 r.).",
            ),
        ],
        "assessment": [
            (
                "Komunikat o „50 drzewach” potrzebnych do wyprodukowania pół miliona kartek papieru oraz o „tlenie na rok dla ok. 100 osób” ma charakter niezweryfikowanego ekwiwalentu przyrodniczego. "
                "Na stronie nie podano źródła przelicznika, gatunku i wieku drzew, gramatury ani pochodzenia papieru (np. udziału makulatury), a tym bardziej roku bazowego dla ponad 40 000 archiwizowanych umów. "
                "W praktyce nadzorczej Prezes UOKiK w komunikatach o postawieniu zarzutów z lipca 2025 r. (sprawy InPost i Allegro — przywołane wyłącznie jako opis linii nadzorczej, nie jako rozstrzygnięcie sądowe) zakwestionował posługiwanie się przelicznikami liczby drzew w marketingu usług bez rzetelnych podstaw dowodowych."
            ),
            (
                "Istotne zastrzeżenie dotyczy także przemilczenia rzeczywistego przebiegu procesu „BioPodpis”. "
                "Nazwa usługi wywodzi się od biometrycznej weryfikacji nacisku i prędkości rysika, lecz w nagłówku wykorzystano ją do budowy skojarzenia ekologicznego („Wspólnie jesteśmy BIO”, „Naturalnie, że podpisane!”) oraz zapewnienia, że rozwiązanie „eliminuje zużycie papieru”. "
                "Tymczasem zgodnie z opisem zamieszczonym niżej na tej samej podstronie, zawarcie umowy BioPodpisem wymaga osobistego przyjazdu doradcy z laptopem i rysikiem do klienta, a skonsolidowany raport ESG Grupy za 2025 r. (s. 246) wskazuje, że flota Grupy składa się w 90% z samochodów hybrydowych i w 9% z aut benzynowych. "
                "Pominięcie emisji transportowych dojazdu doradcy oraz energii zużywanej przez sprzęt IT powoduje, że informacja o korzyści środowiskowej staje się niepełna w świetle analogii z omówienia wyroku SN z 9 października 2014 r., IV CSK 56/14 (w którym SN oddalił skargę kasacyjną, wskazując jednak, że nawet prawdziwa informacja może wprowadzać w błąd przez sposób ujęcia; wyrok nie dotyczy Erste Leasing S.A.)."
            ),
            (
                "Analogiczne zastrzeżenie dotyczy oznaczania usługi „eFaktura” zieloną etykietą „eko” i hasłem „wspólnie jesteśmy EKO – to jest nasz wspólny wkład w ochronę środowiska”. "
                "W świetle wejścia w życie obowiązkowego Krajowego Systemu e-Faktur (KSeF) od lutego 2026 r., elektroniczny obieg faktur staje się w obrocie gospodarczym standardem wynikającym z ustawy, a nie ponadnormatywną inicjatywą ekologiczną. "
                "Z uwagi na to, że z cyfrowych rozwiązań korzystają również klienci indywidualni i jednoosobowi przedsiębiorcy, wykreowanie fałszywego wrażenia skali korzyści ekologicznych może wpływać na wybór formy zawarcia umowy lub postrzeganie oferty jako całości (test decyzji transakcyjnej z omówienia wyroku TSUE C-281/12)."
            ),
            (
                "Od 27 września 2027 r. (ustawa z 4 września 2026 r., Dz.U. z 2026 r. poz. 1353) zielone pigułki „eko” w menu nawigacyjnym witryny oraz hasła „Wspólnie jesteśmy BIO / EKO” będą dodatkowo weryfikowane pod kątem zakazu ogólnych twierdzeń środowiskowych bez uznanej wysokiej efektywności (art. 7 pkt 4a UPNPR) oraz zakazu prezentowania własnych oznaczeń graficznych przypominających znak zrównoważonego charakteru bez niezależnego systemu certyfikacji (art. 7 pkt 2a UPNPR). "
                "Na dzień 24 listopada 2026 r. regulacje te wyznaczają przyszły standard dostosowawczy, nie zaś naruszenie krajowej czarnej listy."
            ),
        ],
        "remediation": (
            "Należy natychmiast usunąć z podstrony /biopodpis przelicznik „50 drzew” i „tlen dla ok. 100 osób”, a także wycofać hasła „Wspólnie jesteśmy BIO”, „Naturalnie, że podpisane!” (w kontekście ekologicznym) oraz „wspólnie jesteśmy EKO” na podstronie /e-faktura. "
            "W menu nawigacyjnym należy zdjąć zielone etykiety „eko” przy czterech usługach cyfrowych. "
            "Zamiast haseł ogólnych należy wskazać wyłącznie weryfikowalną cechę techniczną procesu: brak drukowania dokumentów papierowych w kanałach w pełni zdalnych (bySMS, OTF, eFaktura) oraz brak konieczności druku po stronie klienta przy podpisie biometrycznym składanym w trakcie wizyty doradcy."
        ),
        "claim": (
            "„Umowę możesz zawrzeć elektronicznie bez drukowania dokumentów: całkowicie zdalnie przez Leasing bySMS lub jednorazowy podpis kwalifikowany OTF w eBOK24, "
            "albo na urządzeniu doradcy podczas spotkania z wykorzystaniem podpisu biometrycznego (BioPodpis). "
            "Dokumenty księgowe udostępniamy w formacie elektronicznym w portalu eBOK24 oraz za pośrednictwem KSeF zgodnie z obowiązującymi przepisami.”"
        ),
    },
    {
        "no": "03",
        "title": "Deklaracja „precyzyjnego mierzenia »zielonego wpływu« portfela” (System Oznaczania i Weryfikacji Transakcji Green w kontekście CSRD i Taksonomii UE) bez metodyki i skali portfela",
        "short_title": "„System Oznaczania i Weryfikacji Transakcji Green” oraz deklaracje CSRD/Taksonomii UE bez metodyki",
        "basis": (
            "Art. 16 ust. 1 pkt 2 i ust. 2 ustawy z dnia 16 kwietnia 1993 r. o zwalczaniu nieuczciwej konkurencji "
            "(t.j. Dz.U. z 2026 r. poz. 85); w zakresie komunikatów dostępnych również konsumentom art. 5 ust. 2 pkt 2 i art. 6 ust. 1 "
            "ustawy z dnia 23 sierpnia 2007 r. o przeciwdziałaniu nieuczciwym praktykom rynkowym (t.j. Dz.U. z 2023 r. poz. 845); "
            "kontekst sektorowy: art. 3 i art. 8 rozporządzenia (UE) 2020/852 (Taksonomia UE) oraz wytyczne nadzorcze ESA z 4 czerwca 2024 r. "
            "Przyszły standard EmpCo od 27 września 2027 r.: art. 7 pkt 4a i pkt 4b UPNPR (Dz.U. z 2026 r. poz. 1353)."
        ),
        "risk": "ŚREDNIE",
        "quotes": [
            (
                "Podstrona /erste-leasing/nagrody (Lider Transformacji Energetycznej 2026):",
                "„W 2026 podczas Polskiego Kongresu Klimatycznego otrzymaliśmy tytuł Lidera Transformacji Energetycznej za stworzenie i wdrożenie pionierskiego w branży leasingowej Systemu Oznaczania i Weryfikacji Transakcji Green. To autorskie rozwiązanie pozwala precyzyjnie mierzyć »zielony wpływ« portfela, co jest kluczowe w dobie rosnącej presji regulacyjnej (CSRD, Taksonomia UE).”",
            ),
            (
                "Podstrona /erste-leasing/o-nas:",
                "„Wspólnie budujemy »zieloną« przyszłość. Pomagamy klientom przejść na »zieloną stronę mocy«. Finansujemy zrównoważone inwestycje i aktywnie edukujemy rynek w obszarze ESG. Publikujemy raporty o zielonych inwestycjach oraz inicjujemy debaty, które kształtują standardy nowoczesnej gospodarki.”",
            ),
            (
                "Skonsolidowane sprawozdanie zarządu za 2025 r. (linkowane z /zielone-rozwiazania, s. 254, 259 i 279):",
                "Wyłączenie z kalkulacji finansowanych emisji PCAF pozycji „pojazdy silnikowe (leasing w kategorii 13 zakresu 3 finansowanych emisji)”; wskaźnik zielonych aktywów Grupy (GAR) na 31.12.2025 r.: 3,28% (obrót) oraz 3,56% (nakłady inwestycyjne); wskazanie, że finansowanie zgodne z Taksonomią UE stanowi podgrupę szerszej wewnętrznej klasyfikacji SFICS.",
            ),
        ],
        "assessment": [
            (
                "Opis wdrożonego w 2026 r. „Systemu Oznaczania i Weryfikacji Transakcji Green” wprost łączy autorską klasyfikację spółki z wymogami dyrektywy CSRD oraz Taksonomii UE, sugerując precyzyjny pomiar środowiskowego wpływu portfela leasingowego. "
                "Na całej witrynie nie udostępniono jednak zasad tej klasyfikacji: brak informacji o kryteriach kwalifikacji transakcji, progach technicznych, spełnianiu zasady DNSH z art. 3 rozporządzenia (UE) 2020/852 ani o tym, jaki odsetek z portfela aktywów (8,7 mld zł sfinansowanych w 2025 r.) spełnia te kryteria. "
                "W raportach końcowych Europejskich Urzędów Nadzoru (ESA: EBA, ESMA, EIOPA) z 4 czerwca 2024 r. podkreślono, że komunikacja dotycząca zrównoważonego finansowania nie może opierać się na nieprzejrzystym mechanizmie („black box”), lecz wymaga jawności metodyki i ograniczeń danych."
            ),
            (
                "Jednocześnie weryfikacja materiałów źródłowych załączonych na podstronie /zielone-rozwiazania wskazuje na istotną lukę informacyjną względem zapowiedzi z zakładki /o-nas. "
                "Choć spółka deklaruje „Publikujemy raporty o zielonych inwestycjach”, przycisk „Przeczytaj więcej o dokumentach ESG” prowadzi do zewnętrznego słownika pojęć na portalu wastebusters.pl, a link „Sprawdzam szczegóły” przy raporcie skonsolidowanym otwiera sprawozdanie za 2025 r. opublikowane jeszcze pod nazwą Santander Bank Polska S.A. "
                "Z lektury tego sprawozdania wynika ponadto, że wskaźnik zielonych aktywów (GAR) całej Grupy wynosi 3,28% (obrót) i 3,56% (CapEx), a ekspozycje leasingowe dotyczące pojazdów i maszyn zostały wyłączone z kalkulacji emisji finansowanych PCAF ze względu na brak wypracowanej metodyki dla przedmiotów innych niż auta."
            ),
            (
                "W relacjach z przedsiębiorcami (art. 16 ust. 1 pkt 2 i ust. 2 UZNK) podanie informacji o autorskim systemie pomiaru powiązanym z CSRD i Taksonomią UE bez ujawnienia jego zakresu może wpływać na wybór partnera finansującego przez spółki poszukujące danych do własnego raportowania zrównoważonego rozwoju. "
                "Przy ocenie zaniechania informacyjnego należy uwzględnić całość kanału komunikacji (omówienie tezy wyroku TSUE z 26 października 2016 r., C-611/14, Canal Digital Danmark; wyrok nie dotyczy Erste Leasing S.A.): samo odesłanie na innej podstronie do obszernego dokumentu korporacyjnego poprzedniej marki nie zastępuje jasnej informacji o metodzie i skali tam, gdzie usługa jest reklamowana. "
                "Ponieważ zidentyfikowane uchybienie wynika z braku publikacji założeń metodycznych i błędnego odesłania, poziom ryzyka ustalono jako średni i możliwy do wyeliminowania poprzez uzupełnienie dokumentacji."
            ),
        ],
        "remediation": (
            "Należy opublikować na podstronie /zielone-rozwiazania zwięzłą kartę metodyczną „Systemu Oznaczania i Weryfikacji Transakcji Green”, wskazując wyraźnie, w jakim zakresie opiera się on na wewnętrznej klasyfikacji SFICS, a w jakim spełnia cztery łączne kryteria art. 3 rozporządzenia (UE) 2020/852 (w tym zasadę DNSH). "
            "W tym samym miejscu należy podać procentowy udział finansowania spełniającego kryteria w rocznej sprzedaży lub portfelu spółki leasingowej oraz naprawić odnośnik „Przeczytaj więcej o dokumentach ESG”, kierując go do faktycznych raportów spółki."
        ),
        "claim": (
            "„Erste Leasing S.A. stosuje wewnętrzny System Oznaczania i Weryfikacji Transakcji oparty na klasyfikacji SFICS Grupy Erste Bank Polska. "
            "Zasady kwalifikacji umów, zakres przedmiotowy oraz odrębny udział finansowania zgodnego z art. 3 rozporządzenia (UE) 2020/852 (Taksonomia UE) publikujemy w karcie metodycznej [link] oraz w Skonsolidowanym oświadczeniu o zrównoważonym rozwoju za dany rok obrotowy.”"
        ),
    },
    {
        "no": "04",
        "title": "Przedstawianie ustawowego obowiązku wpisu do rejestru BDO jako wyrazu „świadomej dbałości o środowisko” oraz niespójność numeru rejestrowego BDO",
        "short_title": "Prezentacja obowiązkowego wpisu BDO jako „świadomej dbałości o środowisko” i rozbieżność numeru",
        "basis": (
            "Art. 5 ust. 2 pkt 2 i ust. 4 oraz art. 6 ust. 1 ustawy z dnia 23 sierpnia 2007 r. o przeciwdziałaniu nieuczciwym praktykom rynkowym "
            "(t.j. Dz.U. z 2023 r. poz. 845); art. 16 ust. 1 pkt 2 i ust. 2 ustawy z dnia 16 kwietnia 1993 r. o zwalczaniu nieuczciwej konkurencji "
            "(t.j. Dz.U. z 2026 r. poz. 85) w zw. z art. 50 ust. 1 i art. 63 ustawy z dnia 14 grudnia 2012 r. o odpadach."
        ),
        "risk": "ŚREDNIE",
        "quotes": [
            (
                "Podstrona /erste-leasing/e-faktura (FAQ „Dlaczego mam na fakturze nr BDO? / Co to jest numer BDO na fakturze?”):",
                "„Świadomie dbamy o środowisko. Dlatego jesteśmy w bazie danych o podmiotach, które wprowadzają do obrotu produkty w opakowaniach oraz gospodarują odpadami. Erste Leasing otrzymał w bazie indywidualny numer rejestrowy BDO: 00048510. Zgodnie z ustawą o odpadach, mamy obowiązek umieszczania tego numeru na naszych dokumentach, także księgowych.”",
            ),
            (
                "Stopka serwisu /erste-leasing:",
                "„Nasz numer rejestrowy BDO 000485101.”",
            ),
        ],
        "assessment": [
            (
                "Sekcja FAQ na podstronie eFaktury otwiera odpowiedź na pytanie o numer BDO zdaniem „Świadomie dbamy o środowisko. Dlatego jesteśmy w bazie danych…”, dopiero w ostatnim zdaniu wspominając o ustawowym obowiązku umieszczania numeru na dokumentach. "
                "Wpis do rejestru BDO (art. 50 ust. 1 ustawy z 14 grudnia 2012 r. o odpadach) oraz posługiwanie się nadanym numerem rejestrowym na dokumentach sporządzanych w związku z prowadzoną działalnością (art. 63 tej ustawy) nie są dobrowolną inicjatywą proekologiczną przedsiębiorcy, lecz bezwzględnym obowiązkiem publicznoprawnym. "
                "Odwrócenie relacji przyczynowej może wywoływać u odbiorcy mylne przekonanie, że obecność w bazie BDO stanowi ponadstandardowe działanie środowiskowe spółki, co odpowiada mechanizmowi podania prawdziwej informacji w sposób mogący wprowadzać w błąd (art. 5 ust. 2 pkt 2 UPNPR, art. 16 ust. 1 pkt 2 UZNK; analogia z omówienia wyroku SN z 9 października 2014 r., IV CSK 56/14, w którym SN oddalił skargę kasacyjną — wyrok nie dotyczy Erste Leasing S.A.)."
            ),
            (
                "Należy wyraźnie zastrzec, że w kontekście przyszłego standardu dyrektywy (UE) 2024/825 (wchodzącego w życie w Polsce 27 września 2027 r.), zakaz z projektowanego pkt 10a załącznika obejmuje przedstawianie jako cechy wyróżniającej wymogów nałożonych prawem na wszystkie produkty danej kategorii na rynku Unii Europejskiej (motyw 15 dyrektywy 2024/825). "
                "Ponieważ obowiązek wpisu do BDO jest wymogiem krajowej ustawy o odpadach uzależnionym od spełnienia przesłanek podmiotowych, samo odwołanie się do BDO nie przesądza automatycznie spełnienia przesłanek przyszłego pkt 10a, lecz podlega ocenie na podstawie obowiązujących już dziś przepisów art. 5–6 UPNPR i art. 16 UZNK. "
                "Dodatkową nieprawidłowością rzetelności informacyjnej jest brak ostatniej cyfry numeru rejestrowego w treści FAQ („00048510”) wobec 9-cyfrowego numeru podanego w stopce serwisu („000485101”). "
                "Ze względu na ograniczoną wagę sekcji FAQ wobec oferty głównej oraz łatwość redakcyjnego skorygowania akapitu, ryzyko oceniono jako średnie."
            ),
        ],
        "remediation": (
            "Z odpowiedzi FAQ na podstronie /e-faktura należy usunąć zdania „Świadomie dbamy o środowisko. Dlatego jesteśmy w bazie danych…” i rozpocząć wyjaśnienie od neutralnego wskazania obowiązku wynikającego z ustawy z dnia 14 grudnia 2012 r. o odpadach. "
            "Należy jednocześnie poprawić omyłkę pisarską w numerze BDO, ujednolicając 9-cyfrowy identyfikator „000485101”."
        ),
        "claim": (
            "„Numer BDO 000485101 jest indywidualnym numerem rejestrowym Erste Leasing S.A. w Bazie danych o produktach i opakowaniach oraz o gospodarce odpadami. "
            "Zgodnie z przepisami ustawy z dnia 14 grudnia 2012 r. o odpadach jesteśmy zobowiązani do wpisu do rejestru oraz do umieszczania tego numeru na dokumentach sporządzanych w związku z prowadzoną działalnością, w tym na fakturach.”"
        ),
    },
]


# ---------------------------------------------------------------------------
# 1. Materiał wejściowy do audytu (bez oceny prawnej)
# ---------------------------------------------------------------------------

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

    def _run(r, size, bold=False, italic=False, color=DARK):
        r.font.name = "Arial"
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Arial")

    def p(text, size=10, bold=False, italic=False, color=DARK, before=0, after=5):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(before)
        par.paragraph_format.space_after = Pt(after)
        par.paragraph_format.line_spacing = 1.12
        _run(par.add_run(text), size, bold, italic, color)
        return par

    def h(text):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(14)
        par.paragraph_format.space_after = Pt(4)
        par.paragraph_format.keep_with_next = True
        _run(par.add_run(text), 12.5, True, False, NAVY)

    p("ESG INSTITUTE – Zestaw treści i deklaracji marketingowych do audytu", 15, True, False, NAVY, after=2)
    p(
        "Zbiór dosłownych haseł, opisów ofertowych, elementów nawigacji oraz fragmentów dokumentów załączonych w serwisie "
        "https://www.erste.pl/erste-leasing. Dokument stanowi czysty zapis źródłowy na stan 24 listopada 2026 r. i nie zawiera oceny prawnej.",
        10, False, True, MUTED, after=8,
    )

    rows = [
        ("Badany podmiot", "Erste Leasing S.A. (grupa Erste Bank Polska S.A., poprzednio Santander Leasing S.A.), ul. Kolorowa 8, 60-198 Poznań, KRS 0000026084, NIP 521-12-95-779, REGON 012162265, nr BDO 000485101."),
        ("Oficjalna witryna", "https://www.erste.pl/erste-leasing (stopka podaje również historyczny adres www.ersteleasing.pl)."),
        ("Podstrony objęte zbiorem", "/zielone-rozwiazania wraz z 10 podstronami ofertowymi; /biopodpis; /e-faktura; /leasing-bysms; /jednorazowy-podpis-kwalifikowany; /o-nas; /nagrody; /leasing/leasing-na-instalacje-fotowoltaiczne; /pozyczka/pozyczka-na-instalacje-fotowoltaiczne; /leasing/finansowanie-z-ebi-leasing; /pozyczka/finansowanie-z-ebi-pozyczka; /obsluga-bez-barier; /krajowy-system-e-faktur."),
        ("Załączone pliki PDF", "Certyfikat ISO 14001:2015 (DEKRA, nr 000821168/2); Kodeks ESG dla dostawców spółek leasingowych Erste Bank Polska; Skonsolidowane sprawozdanie zarządu Grupy Kapitałowej Santander Bank Polska S.A. za 2025 r.; Sprawozdanie finansowe za 2025 r."),
        ("Data zebrania", AUDIT_DATE),
        ("Przeznaczenie zbioru", "Wsad źródłowy do raportu pełnego i raportu skróconego ESG Institute."),
    ]
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.autofit = False
    for i, (k, v) in enumerate(rows):
        a, b = tbl.rows[i].cells
        a.width, b.width = Cm(4.4), Cm(13.0)
        pa, pb = a.paragraphs[0], b.paragraphs[0]
        pa.paragraph_format.space_after = Pt(2)
        pb.paragraph_format.space_after = Pt(2)
        _run(pa.add_run(k), 9, True, False, NAVY)
        _run(pb.add_run(v), 9, False, False, DARK)

    sections = [
        ("1. Strona główna (/erste-leasing) i nawigacja", [
            "Zakładki górnej nawigacji: „Klient firmowy Erste Leasing”, „Klient indywidualny Erste F24”, „Agent ubezpieczeniowy Erste Finanse”, „ESG” (link do /erste-leasing/zielone-rozwiazania).",
            "Formularz zapytania ofertowego w stopce posiada przełącznik: „Firma” / „Klient Indywidualny”.",
            "Menu „Cyfrowe rozwiązania” — etykiety graficzne badge--light-green z tekstem „eko” przy pozycjach: „Leasing bySMS eko”, „Jednorazowy podpis kwalifikowany OTF eko”, „BioPodpis eko”, „eFaktura eko”.",
            "Kafelek w sekcji „Sprawdź nasze portale” oraz „Co nowego w Erste Leasing”: „Zielone rozwiązania - dla Ciebie, biznesu, środowiska. Korzystaj z nowoczesnych i odpowiedzialnych rozwiązań uwzględniających dobro środowiska i oszczędzaj na kosztach eksploatacji.”",
            "Komunikat KSeF: „Zmiana sposobu wystawiania faktur - Krajowy System e-Faktur. Sprawdź, co się zmieni od 1 lutego 2026”.",
            "Stopka rejestrowa: „Erste Leasing S.A. Nasza siedziba jest w Poznaniu, adres siedziby to ul. Kolorowa 8, 60-198 Poznań (…) numer KRS 0000026084, NIP 521-12-95-779, REGON 012162265 (…) Nasz numer rejestrowy BDO 000485101. Wysokość kapitału zakładowego i wpłaconego to 310 000 000 zł.”",
        ]),
        ("2. Hub „Zielone rozwiązania” (/erste-leasing/zielone-rozwiazania)", [
            "Nagłówek główny: „Zielone rozwiązania - dla Ciebie, biznesu, środowiska. Korzystaj z nowoczesnych technologii uwzględniających dobro środowiska i oszczędzaj na kosztach eksploatacji.”",
            "Sekcja „Wybrane przedmioty, które możesz u nas sfinansować”: Ciągniki siodłowe i balastowe; Pojazdy hybrydowe; Samochody elektryczne; Dźwigniki, wciągarki i wciągniki przejezdne oraz nieprzejezdne; Oświetlenia LED; Samochody ciężarowe; Rowery, rowery elektryczne; Wózki jezdniowe; Urządzenia i aparaty chłodnicze; Urządzenia wentylacyjne; Urządzenia klimatyzacyjne; Instalacje fotowoltaiczne; Stacje transformatorowe; Panele solarne; Magazyny energii; Stacje transformatorowe nietrwale związane z podłożem; Pompy ciepła; Pielniki; Siewniki do poplonów; Pługi dłutowe; Agregaty do uprawy bezorkowej (gruber); Agregaty uprawowo-siewne do uprawy pasowej; Precyzyjne opryskiwacze sensorowe i recyrkulacyjne; Systemy rolnictwa precyzyjnego (GPS); Rozsiewacze nawozów i wapna ze zmiennym dawkowaniem; Wozy asenizacyjne z aplikatorami; Mulczery - kosiarki bijakowe; Siewnik, agregat nawozowo-siewny, agregat uprawowo-siewny; Rozrzutniki obornika.",
            "Sekcja „Zielone światło dla Twojego biznesu”: „W Erste Leasing pokazujemy, że zrównoważony rozwój może iść w parze z efektywnością i realnymi korzyściami. Nie musisz być ekspertem, żeby zrobić pierwszy krok. Wystarczy ciekawość i otwartość na praktyczne rozwiązania ESG, bo z nami wprowadzanie zielonych rozwiązań dla domu lub firmy jest prostsze, niż myślisz.”",
            "Sekcja „Małe wielkie kroki w ramach współpracy z naturą”: „W Erste Leasing każdego dnia podejmujemy świadome decyzje, by minimalizować nasz negatywny wpływ na środowisko.” — opisuje Kwalifikowany Podpis Elektroniczny (SimplySign), Podpis Biometryczny (BioPodpis) oraz eFakturę.",
            "Sekcja „Aktualne programy grantowe”: eHDV — dofinansowanie na zeroemisyjne ciężarówki (N2 i N3) do 750 000 zł na pojazd, wnioski do 30 czerwca 2029 r., link do oficjalnej strony NFOŚiGW; informacja o wcześniejszej roli operatora programu „Mój elektryk” i wsparciu w „NaszEauto”.",
            "Sekcja „Dla szerokiego grona klientów”: „Program jest skierowany zarówno do osób fizycznych, jak i jednoosobowych działalności gospodarczych. Każdy, kto chce zmniejszyć swój ślad węglowy i korzystać z nowoczesnych rozwiązań, może skorzystać z tej oferty.”",
            "Sekcja „ESG i zrównoważony rozwój”: „Słownik pojęć ESG” (link zewnętrzny do wastebusters.pl/artykuly/slownik-pojec-esg/); „Raport skonsolidowany - część ESG” (link do PDF sprawozdania zarządu za 2025 r., s. 201).",
            "Sekcja „Polityki i regulaminy”: Kodeks ESG dla dostawców spółek leasingowych (PDF) oraz Certyfikat ISO 14001:2015 (PDF).",
        ]),
        ("3. Podstrony ofertowe w sekcji „Zielone rozwiązania”", [
            "/multisalon24-szeroki-wybor-aut-od-autoryzowanych-dealerow: „Multisalon24 to miejsce, gdzie znajdziesz bogatą ofertę pojazdów – również elektrycznych i hybrydowych – wyłącznie od autoryzowanych dealerów (…) Szeroki wybór - setki pojazdów różnych marek i modeli – zarówno nowych, jak i używanych.”",
            "/multisalon24-szeroki-wybor-aut-elektrycznych-i-hybrydowych: „Oferta dla osób fizycznych - Multisalon24 to rozwiązanie nie tylko dla firm! Z oferty mogą skorzystać osoby, których źródłem dochodu jest umowa o pracę, renta, emerytura czy umowa zlecenie.”",
            "/leasing-i-pozyczka-na-niskoemisyjny-transport: „Szukasz sposobu, żeby Twoja firma jeździła nowocześnie, ekonomicznie i bez zanieczyszczania powietrza spalinami? (…) finansujemy floty niespalinowe zarówno dla prywatnych firm, jak i sektora publicznego.”; „Lekkie pojazdy niskoemisyjne – hybrydy typu plug-in, lekkie pojazdy użytkowe i specjalistyczne z ograniczoną emisją CO₂.”; „Transport ciężki nowej generacji – pojazdy zasilane LNG, CNG, wodorem lub w technologii hybrydowej, przeznaczone do przewozu towarów na dłuższych trasach.”",
            "/pozyczka-na-zielone-maszyny-rolnicze: „Niezależnie od wyboru pojazdu, innowacyjna technologia daje przewagę: pracujesz wydajniej, oszczędzasz paliwo i zostawiasz konkurencję daleko w tyle – a środowisko tylko na tym zyskuje.”",
            "/leasing-i-pozyczka-na-maszyny-i-urzadzenia-wydajne-energetycznie: finansowanie elektrycznych wózków widłowych oraz urządzeń klimatyzacyjnych i chłodniczych „wysokiej klasy energetycznej” bez podania konkretnych klas.",
            "/pozyczka-na-instalacje-fotowoltaiczne: w sekcji „Co zyskujesz, wybierając naszą pożyczkę?” użyto zapisu „Swoboda finansowania – leasing nawet do 120 miesięcy”.",
            "/solarlease-leasing-instalacji-fotowoltaicznych: „Fotowoltaika to strategiczny krok dla firm, które budują odporność biznesową i chcą rzetelnie realizować cele redukcji emisji.”",
            "/pozyczka-na-domowe-elektrownie-wiatrowe-i-inne-zielone-technologie-finansowanie-blizsze-naturze: „dzięki własnej instalacji i magazynowi energii możesz sam kontrolować zużycie energii a w razie awarii sieci korzystasz z własnych zapasów, więc przerwy w dostawie nie są już problemem.”",
        ]),
        ("4. Usługi cyfrowe (/biopodpis, /e-faktura, /leasing-bysms, /jednorazowy-podpis-kwalifikowany)", [
            "/biopodpis: „Naturalnie, że podpisane! Wspólnie jesteśmy BIO. Zmniejsz koszty związane z wydrukiem dokumentów i zużyj mniej papieru!”; „eliminuje zużycie papieru”; „doradca leasingowy umówi się z Tobą na spotkanie (…) Doradca leasingowy przyjedzie z kompletnym rozwiązaniem, rysikiem i laptopem.”; „Czy wiesz, że rocznie archiwizujemy ponad 40 000 umów? Zużywamy na to ponad pół miliona kartek! By wyprodukować taką ilość papieru, potrzeba aż 50 drzew! Tyle drzew może zapewnić tlen na rok dla ok. 100 osób.”",
            "/e-faktura: „wspólnie jesteśmy EKO – to jest nasz wspólny wkład w ochronę środowiska.”; FAQ BDO: „Świadomie dbamy o środowisko. Dlatego jesteśmy w bazie danych o podmiotach, które wprowadzają do obrotu produkty w opakowaniach oraz gospodarują odpadami. Erste Leasing otrzymał w bazie indywidualny numer rejestrowy BDO: 00048510. Zgodnie z ustawą o odpadach, mamy obowiązek umieszczania tego numeru na naszych dokumentach, także księgowych.”",
            "/leasing-bysms oraz /jednorazowy-podpis-kwalifikowany: sekcja „Ekologia — W trosce o środowisko zachęcamy do korzystania z podpisów elektronicznych. Pomagasz w ten sposób oszczędzać (tysiące kartek) papieru i redukować transport dokumentów.”",
        ]),
        ("5. Podstrony /o-nas, /nagrody oraz załączone dokumenty PDF", [
            "/o-nas: „Wspólnie budujemy »zieloną« przyszłość. Pomagamy klientom przejść na »zieloną stronę mocy«. Finansujemy zrównoważone inwestycje i aktywnie edukujemy rynek w obszarze ESG. Publikujemy raporty o zielonych inwestycjach oraz inicjujemy debaty, które kształtują standardy nowoczesnej gospodarki.”",
            "/nagrody: „Lider Transformacji Energetycznej. W 2026 podczas Polskiego Kongresu Klimatycznego otrzymaliśmy tytuł Lidera Transformacji Energetycznej za stworzenie i wdrożenie pionierskiego w branży leasingowej Systemu Oznaczania i Weryfikacji Transakcji Green. To autorskie rozwiązanie pozwala precyzyjnie mierzyć »zielony wpływ« portfela, co jest kluczowe w dobie rosnącej presji regulacyjnej (CSRD, Taksonomia UE).”",
            "Certyfikat ISO 14001:2015 (DEKRA Certification Sp. z o.o., nr 000821168/2, ważny 27.04.2026 – 11.08.2027): zakres „Finansowanie w formie leasingu i pożyczki”, siedziba Poznań, ul. Kolorowa 8.",
            "Kodeks ESG dla dostawców spółek leasingowych Erste Bank Polska (6 stron): dotyczy Erste Leasing S.A., Erste Finanse sp. z o.o. i Erste F24 S.A.; odwołuje się do Celów Zrównoważonego Rozwoju ONZ, Porozumienia paryskiego, UN Global Compact, UNEP FI, Wytycznych OECD dla przedsiębiorstw wielonarodowych i konwencji MOP.",
            "Skonsolidowane sprawozdanie zarządu Grupy Kapitałowej Santander Bank Polska S.A. za 2025 r.: w 2025 r. spółka leasingowa sfinansowała aktywa netto o wartości 8,7 mld zł; wskaźnik GAR Grupy na 31.12.2025 r. wynosił 3,28% (obrót) oraz 3,56% (nakłady inwestycyjne); ekspozycje leasingowe wyłączono z kalkulacji finansowanych emisji PCAF (s. 254); flota samochodowa Grupy składała się w 90% z hybryd i w 9% z aut benzynowych (s. 246).",
        ]),
    ]

    for title, items in sections:
        h(title)
        for item in items:
            p("•  " + item, size=9.5, after=4)

    doc.save(MATERIAL_OUT)
    return MATERIAL_OUT


# ---------------------------------------------------------------------------
# 2. Raport pełny (klasa ReportDoc z sklad_raportu.py)
# ---------------------------------------------------------------------------

def build_full_report():
    report = ReportDoc(
        reference=REFERENCE,
        document_name="Audyt greenwashingu — Erste Leasing S.A.",
        audit_date=AUDIT_DATE,
    )
    report.doc.core_properties.subject = "Audyt Ryzyk Greenwashingu — Erste Leasing S.A."
    report.title("Audyt Ryzyk Greenwashingu: ERSTE LEASING S.A.")
    report.subtitle(
        "Weryfikacja publicznej witryny https://www.erste.pl/erste-leasing, podstron „Zielone rozwiązania”, kart usług cyfrowych i dokumentów załączonych w serwisie (stan na 24 listopada 2026 r.). "
        "Ocena wobec ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, ustawy o zwalczaniu nieuczciwej konkurencji oraz rozporządzeń sektorowych."
    )

    report.meta(
        [
            (
                "Badany podmiot",
                "Erste Leasing S.A. (grupa Erste Bank Polska S.A., wcześniej Santander Leasing S.A.; od stycznia 2026 r. głównym akcjonariuszem Erste Bank Polska S.A. jest Erste Group Bank AG), ul. Kolorowa 8, 60-198 Poznań, KRS 0000026084, NIP 521-12-95-779, REGON 012162265, nr rejestrowy BDO 000485101, kapitał zakładowy i wpłacony 310 000 000 zł, status dużego przedsiębiorcy.",
                False,
            ),
            (
                "Przedmiot audytu",
                "Strona główna https://www.erste.pl/erste-leasing, podstrona /zielone-rozwiazania wraz z 10 podstronami ofertowymi, podstrony usług cyfrowych (/biopodpis, /e-faktura, /leasing-bysms, /jednorazowy-podpis-kwalifikowany), /o-nas, /nagrody, /obsluga-bez-barier, /krajowy-system-e-faktur, karty PV i EBI oraz załączone dokumenty PDF (Certyfikat ISO 14001:2015, Kodeks ESG dla dostawców, Skonsolidowane sprawozdanie zarządu za 2025 r.). Stan zebrania: 24 listopada 2026 r.",
                False,
            ),
            (
                "Sektor działalności",
                "Usługi finansowe: leasing operacyjny i finansowy, pożyczki oraz najem pojazdów, maszyn, urządzeń i instalacji OZE. Oferta kierowana zarówno do przedsiębiorców (B2B), jak i do osób fizycznych nieprowadzących działalności gospodarczej (konsumentów — B2C).",
                False,
            ),
            (
                "Reżim prawny",
                "Ustawa z 23 sierpnia 2007 r. o przeciwdziałaniu nieuczciwym praktykom rynkowym (t.j. Dz.U. z 2023 r. poz. 845); ustawa z 16 lutego 2007 r. o ochronie konkurencji i konsumentów (t.j. Dz.U. z 2025 r. poz. 1714); ustawa z 16 kwietnia 1993 r. o zwalczaniu nieuczciwej konkurencji (t.j. Dz.U. z 2026 r. poz. 85); ustawa z 14 grudnia 2012 r. o odpadach; rozporządzenie (UE) 2020/852 (Taksonomia UE). W dniu audytu (24 listopada 2026 r.) czarna lista dyrektywy (UE) 2024/825 nie jest jeszcze obowiązującym prawem polskim.",
                False,
            ),
            ("Data audytu", AUDIT_DATE, False),
            ("Werdykt operacyjny", VERDICT_RED, True),
        ]
    )

    report.verdict(
        VERDICT_RED,
        [
            "Dwa ustalenia (Przypadki 01 i 02) spełniają w dniu audytu (24 listopada 2026 r.) przesłanki naruszenia obowiązujących art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym (w relacji z konsumentami) oraz art. 16 ust. 1 pkt 2 i ust. 2 ustawy o zwalczaniu nieuczciwej konkurencji (w relacji z przedsiębiorcami), w tym test wpływu na decyzję dotyczącą umowy.",
            "Dotyczy to kwalifikowania pojazdów spalinowych, hybrydowych oraz zasilanych LNG i CNG, a także maszyn i urządzeń bez progów emisyjnych pod szyldem „Zielone rozwiązania” oraz hasła transportu „bez zanieczyszczania powietrza spalinami”, jak również posługiwania się etykietami „eko”, hasłem „Wspólnie jesteśmy BIO” i niezweryfikowanym przelicznikiem „50 drzew / tlen dla ok. 100 osób” przy pominięciu dojazdu doradcy oraz obowiązkowości KSeF.",
            "Nowe zakazy czarnej listy dyrektywy (UE) 2024/825, wdrożone ustawą z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353), wejdą w życie w Polsce dopiero 27 września 2027 r., dlatego nie stanowią podstawy bieżącego werdyktu, lecz wyznaczają obowiązkowy horyzont dalszego dostosowania serwisu.",
        ],
    )

    report.h1("1. Analiza wstępna")

    report.h2("1.1. Reżim prawny")
    report.paragraph(
        "W ocenie regulacyjnej na dzień 24 listopada 2026 r. należy wyraźnie rozdzielić dwie daty wynikające z procesu implementacji dyrektywy Parlamentu Europejskiego i Rady (UE) 2024/825 z dnia 28 lutego 2024 r. (EmpCo). "
        "Art. 4 ust. 1 dyrektywy wyznacza 27 września 2026 r. jako datę, od której państwa członkowskie miały stosować krajowe przepisy transponujące, lecz nie jest to data wejścia w życie ustawy polskiej. "
        "Ustawa z dnia 4 września 2026 r. o zmianie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawy o prawach konsumenta (Dz.U. z 2026 r. poz. 1353) została ogłoszona 19 października 2026 r., a zgodnie z jej art. 3 wchodzi w życie 27 września 2027 r. "
        "Do 26 września 2027 r. włącznie nowe punkty czarnej listy (projektowane art. 7 pkt 2a, 4a, 4b, 4c i 10a UPNPR) stanowią wyłącznie przyszły standard regulacyjny i nie mogą być kwalifikowane jako już dokonane naruszenie krajowej czarnej listy."
    )
    report.paragraph(
        "Werdykt audytu opiera się na przepisach obowiązujących w dniu badania. "
        "Ponieważ Erste Leasing S.A. kieruje ofertę podstrony „Zielone rozwiązania” oraz usług cyfrowych nie tylko do firm, lecz także wyraźnie do osób fizycznych nieprowadzących działalności gospodarczej (zakładki „Dla Ciebie”, finansowanie samochodów dla osób zatrudnionych na umowę o pracę, rencistów i emerytów), zastosowanie mają łącznie przepisy chroniące konsumentów (art. 5 i art. 6 UPNPR, t.j. Dz.U. z 2023 r. poz. 845; art. 24 i art. 106 ust. 1 pkt 4 u.o.k.k., t.j. Dz.U. z 2025 r. poz. 1714) oraz przepisy dotyczące obrotu profesjonalnego B2B (art. 16 i art. 18 UZNK, t.j. Dz.U. z 2026 r. poz. 85). "
        "W zakresie odwołań do klasyfikacji zielonych aktywów uwzględniono regulacje sektorowe: art. 3 i art. 8 rozporządzenia (UE) 2020/852 (Taksonomia UE) oraz raporty końcowe Europejskich Urzędów Nadzoru (ESA) z 4 czerwca 2024 r. dotyczące ryzyk greenwashingu, a w odniesieniu do rejestru BDO — art. 50 ust. 1 i art. 63 ustawy z dnia 14 grudnia 2012 r. o odpadach."
    )

    report.h2("1.2. Profil działalności i narracja serwisu")
    report.paragraph(
        "Erste Leasing S.A. (działająca wcześniej pod nazwą Santander Leasing S.A., w grupie Erste Bank Polska S.A., której głównym akcjonariuszem od stycznia 2026 r. jest Erste Group Bank AG) prowadzi działalność w zakresie leasingu operacyjnego i finansowego, pożyczek oraz najmu środków trwałych. "
        "Zgodnie ze sprawozdaniem finansowym i sprawozdaniem zarządu za 2025 r., spółka sfinansowała w 2025 r. aktywa netto o wartości 8,7 mld zł, a na koniec roku obsługiwała ponad 158 tys. aktywnych umów. "
        "Na witrynie https://www.erste.pl/erste-leasing wyodrębniono rozbudowany dział „Zielone rozwiązania - dla Ciebie, biznesu, środowiska” (podpięty bezpośrednio pod przycisk „ESG” w głównej nawigacji) oraz oznaczono zieloną etykietą „eko” cztery usługi cyfrowego zawierania umów i fakturowania."
    )
    report.paragraph(
        "Analiza wykazała, że obok rzetelnie udokumentowanych programów wsparcia (EBI, eHDV NFOŚiGW) i certyfikacji ISO 14001:2015, przekaz marketingowy serwisu w czterech obszarach przekracza granice rzetelności informacyjnej. "
        "Najwyższe ryzyko dotyczy objęcia parasolową nazwą „Zielone rozwiązania” i hasłami o jeździe „bez zanieczyszczania powietrza spalinami” również samochodów spalinowych z portalu Multisalon24, pojazdów zasilanych LNG/CNG oraz hybryd, a także stosowania niezweryfikowanego przelicznika drzew i tlenu na podstronie BioPodpisu. "
        "Dodatkowe zastrzeżenia dotyczą braku jawnej metodyki „Systemu Oznaczania i Weryfikacji Transakcji Green” oraz przedstawiania obowiązkowego wpisu do rejestru BDO jako dobrowolnej dbałości o środowisko."
    )

    report.h2("1.3. Materiał z witryny")
    report.paragraph(
        "Poniżej przedstawiono główny widok sekcji „Zielone rozwiązania” w serwisie Erste Leasing, w którym połączono hasło szerokiej transformacji środowiskowej z wykazem finansowanych pojazdów i maszyn obejmującym również środki transportu o napędzie spalinowym i hybrydowym."
    )
    report.figure(
        SHOT_ZIELONE,
        "Rys. 1. Zrzut ekranu podstrony https://www.erste.pl/erste-leasing/zielone-rozwiazania (stan na 24 listopada 2026 r.). Źródło: opracowanie własne.",
        width_cm=8.2,
    )

    report.h2("1.4. Wyłączenia")
    report.paragraph(
        "W ramach audytu zweryfikowano pozostałe elementy komunikacji ESG w serwisie i nie zakwalifikowano ich jako naruszeń prawa obowiązującego:"
    )
    report.bullet(
        "Certyfikat ISO 14001:2015. ",
        "Spółka udostępnia aktualny certyfikat wydany przez DEKRA Certification Sp. z o.o. (nr 000821168/2, ważny od 27 kwietnia 2026 r. do 11 sierpnia 2027 r., zakres: finansowanie w formie leasingu i pożyczki dla siedziby przy ul. Kolorowej 8 w Poznaniu). Sama publikacja certyfikatu jest rzetelna; nie zastępuje on jednak oznakowania EN ISO 14024 typu I ani EU Ecolabel na potrzeby przyszłego standardu EmpCo od 27 września 2027 r.",
    )
    report.bullet(
        "Programy dotacyjne i finansowanie z EBI. ",
        "Opisy programu eHDV NFOŚiGW (dotacje do 750 000 zł dla pojazdów kategorii N2 i N3, nabór do 30 czerwca 2029 r., odesłanie do oficjalnej strony NFOŚiGW), doświadczeń z programów „Mój elektryk” i „NaszEauto” oraz warunków finansowania z Europejskiego Banku Inwestycyjnego (wraz z publiczną listą wyłączeń kodów NACE/PKD) są konkretne i weryfikowalne.",
    )
    report.bullet(
        "Brak reklamy porównawczej. ",
        "W badanym materiale nie stwierdzono porównania umożliwiającego bezpośrednie lub pośrednie zidentyfikowanie konkurencyjnego leasingodawcy lub jego oferty w rozumieniu art. 16 ust. 3 UZNK (por. omówienie wyroku TSUE z 19 kwietnia 2007 r., C-381/05, De Landtsheer; wyrok nie dotyczy Erste Leasing S.A.). Z tego względu raport nie zawiera odrębnej sekcji analizy porównawczej.",
    )
    report.bullet(
        "Obszar społeczny, dostępność i Wytyczne OECD. ",
        "Informacje o dostosowaniu portalu eBOK24 do WCAG 2.1, tłumaczu PJM, ramkach do podpisu, benefitach pracowniczych oraz Kodeks ESG dla dostawców (odwołujący się m.in. do Wytycznych OECD dla przedsiębiorstw wielonarodowych) nie naruszają art. 5–6 UPNPR. Same Wytyczne OECD mają charakter zaleceń (poziom niski); ewentualne zgłoszenie do polskiego Krajowego Punktu Kontaktowego OECD nie kończy się wyrokiem sądowym i nie wpływa na podniesienie werdyktu.",
    )
    report.bullet(
        "Drobne uwagi redakcyjne i oferty OZE. ",
        "Na karcie pożyczki fotowoltaicznej omyłkowo użyto słowa „leasing” („leasing nawet do 120 miesięcy”), a w opisie przydomowych wiatraków niezależność od przerw w dostawie prądu warto doprecyzować warunkiem posiadania magazynu energii z funkcją zasilania awaryjnego; kwestie te ujęto w drugim etapie rekomendacji naprawczych.",
    )

    report.h1("2. Analiza szczegółowa — tabela zbiorcza naruszeń")
    report.findings(FINDINGS)

    report.h1("3. Szczegółowa analiza przypadków")
    for c in CASES:
        report.case(f"Przypadek {c['no']}. {c['title']}", c["basis"], c["risk"])
        report.paragraph("Cytat źródłowy", size=10.5, bold=True, before=2, after=2, align="left")
        for label, quote_text in c["quotes"]:
            report.quote(label, quote_text)
        report.paragraph("Ocena prawna", size=10.5, bold=True, before=6, after=2, align="left")
        for para_text in c["assessment"]:
            report.paragraph(para_text)
        report.paragraph("Kierunek naprawy", size=10.5, bold=True, before=2, after=2, align="left")
        report.paragraph(c["remediation"])
        report.claim(c["claim"])

    report.h1("4. Ryzyka prawne i finansowe")
    report.h2("4.1. Odpowiedzialność przed Prezesem UOKiK za praktyki naruszające zbiorowe interesy konsumentów")
    report.paragraph(
        "Ponieważ sekcja „Zielone rozwiązania” oraz usługi cyfrowe kierowane są wprost także do osób fizycznych nieprowadzących działalności gospodarczej (m.in. Multisalon24 i pożyczki OZE „dla domu”), zastrzeżenia z Przypadków 01 i 02 podlegają nadzorowi Prezesa UOKiK na podstawie art. 24 ust. 1 i ust. 2 pkt 3 ustawy o ochronie konkurencji i konsumentów w związku z art. 5 i art. 6 UPNPR. "
        "Zgodnie z art. 106 ust. 1 pkt 4 tej ustawy, Prezes UOKiK może nałożyć na przedsiębiorcę karę pieniężną w wysokości do 10% obrotu osiągniętego w roku obrotowym poprzedzającym rok nałożenia kary za każdą praktykę naruszającą prawo już obowiązujące, a także zobowiązać spółkę do zaniechania praktyki i usunięcia jej skutków."
    )
    report.h2("4.2. Roszczenia z ustawy o zwalczaniu nieuczciwej konkurencji i oczekiwania nadzorcze ESA")
    report.paragraph(
        "W segmencie klientów firmowych (B2B) przekazy zakwalifikowane w Przypadkach 01–04 wypełniają przesłanki nieuczciwej reklamy wprowadzającej w błąd z art. 16 ust. 1 pkt 2 i ust. 2 UZNK. "
        "Na podstawie art. 18 ust. 1 UZNK przedsiębiorca, którego interes został zagrożony lub naruszony, może żądać zaniechania niedozwolonych działań, usunięcia ich skutków, złożenia oświadczenia odpowiedniej treści i formy, a także naprawienia szkody lub wydania bezpodstawnie uzyskanych korzyści. "
        "Niezależnie od sporów cywilnych, dla podmiotu z grupy bankowej kluczowe znaczenie ma spójność komunikacji produktowej z wytycznymi nadzorczymi Europejskich Urzędów Nadzoru (ESA z 4 czerwca 2024 r.), wymagającymi rzetelności i jawności metodologicznej twierdzeń zrównoważonych."
    )
    report.h2("4.3. Dostosowanie do standardu dyrektywy (UE) 2024/825 przed 27 września 2027 r.")
    report.paragraph(
        "Z dniem 27 września 2027 r. wejdzie w życie ustawa z dnia 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353), wprowadzająca do polskiego prawa zakazy czarnej listy EmpCo stosowane w każdych okolicznościach bez konieczności badania testu decyzji transakcyjnej w relacji B2C. "
        "Utrzymywanie w tym terminie w internecie zielonych etykiet „eko” bez systemu certyfikacji (art. 7 pkt 2a), ogólnych haseł „Wspólnie jesteśmy BIO / EKO” i „finansowanie bliższe naturze” bez uznanej wysokiej efektywności ekologicznej (pkt 4a) oraz parasolowego oznaczania pełnej oferty Multisalon24 jako „zielonej” (pkt 4b) będzie stanowiło bezwzględnie zakazaną praktykę rynkową."
    )
    report.h2("4.4. Ryzyko wtórne dla raportowania CSRD i Taksonomii UE klientów korporacyjnych")
    report.paragraph(
        "Podstrona /nagrody oraz podstrona /o-nas łączą ofertę leasingową wprost z obowiązkami sprawozdawczości zrównoważonego rozwoju (CSRD) i Taksonomią UE, wskazując na możliwość precyzyjnego mierzenia „zielonego wpływu” portfela. "
        "Brak publikacji kryteriów „Systemu Oznaczania i Weryfikacji Transakcji Green” oraz wyłączenie ekspozycji leasingowych z kalkulacji finansowanych emisji PCAF w sprawozdaniu Grupy za 2025 r. stwarza ryzyko rozbieżności oczekiwań u klientów korporacyjnych, którzy mogą błędnie zakładać uzyskanie gotowych danych taksonomicznych lub emisyjnych do własnych ujawnień niefinansowych."
    )

    report.h1("5. Rekomendacje naprawcze")
    report.h2("5.1. Etap natychmiastowy, do 14 dni od daty audytu (do 8 grudnia 2026 r.)")
    report.paragraph(
        "W terminie 14 dni należy skorygować treści naruszające przepisy już obowiązujące (Przypadki 01 i 02). "
        "Obejmuje to wyłączenie ogólnej oferty Multisalon24 oraz pojazdów LNG, CNG, hybrydowych i ciągników spalinowych spod parasola „Zielone rozwiązania” i haseł „bez zanieczyszczania powietrza spalinami” / „floty niespalinowe”, a także usunięcie z podstrony /biopodpis przelicznika „50 drzew / tlen dla ok. 100 osób”, haseł „Wspólnie jesteśmy BIO” i „Naturalnie, że podpisane!” oraz twierdzeń „wspólnie jesteśmy EKO” z podstrony /e-faktura."
    )
    report.h2("5.2. Korekta ujawnień metodycznych, BDO oraz ofert OZE, do 30 dni (do 24 grudnia 2026 r.)")
    report.paragraph(
        "W terminie 30 dni należy zrealizować zalecenia z Przypadków 03 i 04 oraz poprawki redakcyjne. "
        "Należy opublikować skrótową kartę metodyczną Systemu Oznaczania i Weryfikacji Transakcji Green (wskazując relację do wewnętrznej klasyfikacji SFICS i art. 3 rozporządzenia 2020/852), naprawić odnośnik prowadzący do zewnętrznego słownika na wastebusters.pl, przeredagować akapit FAQ dotyczący rejestru BDO z ujednoliceniem 9-cyfrowego numeru „000485101”, poprawić omyłkę „leasing” na karcie pożyczki PV oraz doprecyzować warunek posiadania magazynu energii przy deklaracjach odporności na przerwy w sieci."
    )
    report.h2("5.3. Dostosowanie do standardu EmpCo przed 27 września 2027 r. (do 30 czerwca 2027 r.)")
    report.paragraph(
        "Z odpowiednim wyprzedzeniem przed wejściem w życie ustawy z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) należy przeprowadzić pełny przegląd etykiet graficznych i haseł parasolowych w całym serwisie, usuwając zielone pigułki „eko” z menu nawigacyjnego „Cyfrowe rozwiązania” oraz zamieniając ogólne zwroty („zielone światło”, „zielone technologie”, „finansowanie bliższe naturze”) na precyzyjne opisy parametrów technicznych finansowanych aktywów."
    )
    report.h2("5.4. Stałe sformułowania, kontrola publikacji i szkolenie sieci sprzedaży (I–II kwartał 2027 r.)")
    report.paragraph(
        "Należy wdrożyć do procesu akceptacji materiałów marketingowych i produktowych stałe wzory sformułowań zawarte w rozdziale 3 niniejszego raportu oraz listę kontrolną zgodności z art. 5–6 UPNPR, art. 16 UZNK i rozporządzeniem (UE) 2020/852. "
        "W pierwszym półroczu 2027 r. zaleca się przeprowadzenie szkolenia sieci sprzedaży i doradców terenowych w zakresie rozróżniania pojazdów bezemisyjnych w miejscu użytkowania od niskoemisyjnych i spalinowych oraz zasad komunikowania korzyści cyfrowego obiegu dokumentów."
    )

    report.h1("6. Wnioski")
    report.paragraph(
        "W świetle przepisów obowiązujących w dniu 24 listopada 2026 r. serwis Erste Leasing S.A. wymaga pilnej korekty kwalifikacji pojazdów i maszyn w dziale „Zielone rozwiązania” oraz usunięcia niezweryfikowanego przelicznika drzew i haseł „BIO”/„EKO” w usługach cyfrowych, aby wyeliminować ryzyko interwencji Prezesa UOKiK i roszczeń konkurencji. "
        "Równoległe uporządkowanie metodyki systemu oznaczania transakcji oraz wycofanie ogólnych etykiet „eko” pozwoli dostosować witrynę przed 27 września 2027 r., kiedy to wejdą w życie nowe zakazy czarnej listy wynikające z ustawy z dnia 4 września 2026 r."
    )

    report.sign_off()
    report.save(FULL_OUT)
    return FULL_OUT


# ---------------------------------------------------------------------------
# 3. Raport skrócony (układ wg generuj_raport_skrocony.py / WZOR_raport_skrocony.docx)
# ---------------------------------------------------------------------------

def build_short_report():
    NAVY = RGBColor(8, 24, 31)
    BLUE = RGBColor(58, 97, 179)
    CYAN_SOFT = RGBColor(142, 197, 224)
    DARK = RGBColor(22, 36, 44)
    MUTED = RGBColor(90, 107, 119)
    WHITE = RGBColor(255, 255, 255)
    DANGER = RGBColor(190, 18, 60)
    WARNING = RGBColor(180, 83, 9)
    PAGE_W = 17.0
    LOGO = default_logo()

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
        image = Image.open(SHOT_ZIELONE).convert("RGB")
        # Kadr 1280x465 ukazujący hasło główne "Zielone rozwiązania - dla Ciebie, biznesu, środowiska" i drogę w lesie
        crop = image.crop((80, 130, 1360, 595))
        stream = BytesIO()
        crop.save(stream, format="PNG", optimize=True)
        stream.seek(0)
        return stream

    def _cell_p(cell, clear=True):
        paragraph = cell.paragraphs[0]
        if clear:
            _clear_paragraph(paragraph)
        return paragraph

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
    props.subject = "Raport skrócony audytu ryzyk greenwashingu — ERSTE LEASING S.A."
    props.title = "Raport skrócony audytu ryzyk greenwashingu — ERSTE LEASING S.A."
    props.comments = f"ESG Institute · esginstitute.eu · raport skrócony · {AUDIT_DATE}"

    def para(
        text, size=10.5, bold=False, italic=False, color=DARK, before=0, after=8,
        align="left", keep_next=False, keep_lines=False, page_break=False,
        left=0, right=0, first=0, lang="pl-PL", tracking=None, space=1.15,
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

    def rich(parts, size=10.5, before=0, after=8, align="left", keep_next=False, keep_lines=False, space=1.15):
        paragraph = doc.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(before)
        fmt.space_after = Pt(after)
        fmt.line_spacing = space
        if align == "justify":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _widow(paragraph, keep_next, keep_lines, False)
        for text, bold, italic, color, lang in parts:
            _run(paragraph.add_run(text), size, bold, italic, color, lang=lang)
        return paragraph

    def body(text, size=10.5, color=DARK, before=2, after=8, keep_next=False, keep_lines=True, first=0.75, lang="pl-PL"):
        return para(
            text, size=size, color=color, before=before, after=after,
            align="justify", keep_next=keep_next, keep_lines=keep_lines, first=first, lang=lang,
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
        first_p = True
        for text, size, bold, italic, color, lang in paragraphs:
            if first_p:
                p = body_cell.paragraphs[0]
                _clear_paragraph(p)
                first_p = False
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
    if LOGO and LOGO.exists():
        lp.add_run().add_picture(str(LOGO), width=Cm(4.15))
    else:
        _run(lp.add_run("ESG INSTITUTE"), 11, True, False, WHITE)
    rp = _cell_p(right)
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run(rp.add_run("Raport skrócony"), 8, True, False, CYAN_SOFT)
    rp2 = right.add_paragraph()
    _clear_paragraph(rp2)
    rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run(rp2.add_run("ERSTE LEASING S.A."), 8, True, False, WHITE)
    _run(rp2.add_run("   ·   esginstitute.eu"), 8, False, False, CYAN_SOFT)

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
    _run(
        fp.add_run("ESG Institute Sp. z o.o.   ·   esginstitute.eu   ·   raport skrócony   ·   Erste Leasing S.A."),
        8, False, False, MUTED,
    )
    _run(fp.add_run("\t"), 8, False, False, MUTED)
    _field(fp, "PAGE")
    _run(fp.add_run("  /  "), 8, False, False, MUTED)
    _field(fp, "NUMPAGES")

    # --- strona 1 ---
    para(
        "ESG INSTITUTE   ·   RAPORT AUDYTU REGULACYJNEGO I ŚRODOWISKOWEGO",
        size=8.5, bold=True, color=BLUE, before=4, after=4, keep_next=True, keep_lines=True, tracking=18,
    )
    para(
        "Audyt Ryzyk Greenwashingu: ERSTE LEASING S.A.",
        size=20, bold=True, color=NAVY, before=0, after=8, keep_next=True, keep_lines=True, space=1.05,
    )
    para(
        "Raport skrócony. Weryfikacja witryny https://www.erste.pl/erste-leasing, podstron „Zielone rozwiązania” oraz kart usług cyfrowych, stan na 24 listopada 2026 r. "
        "Werdykt opiera się na ustawie o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawie o zwalczaniu nieuczciwej konkurencji. "
        "Czarna lista dyrektywy (UE) 2024/825 (ustawa z 4 września 2026 r., Dz.U. z 2026 r. poz. 1353) jest przyszłym standardem od 27 września 2027 r.",
        size=10.5, italic=True, color=MUTED, before=2, after=12, space=1.2,
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
            "Fragment podstrony https://www.erste.pl/erste-leasing/zielone-rozwiazania z 24 listopada 2026 r. "
            "Sekcja kieruje przekaz jednocześnie do konsumentów („dla Ciebie”) oraz przedsiębiorców („biznesu”), obejmując również pojazdy spalinowe i hybrydowe."
        ),
        8, False, True, MUTED,
    )
    after_frame = doc.add_paragraph()
    _clear_paragraph(after_frame)
    after_frame.paragraph_format.space_after = Pt(12)

    meta = [
        (
            "Badany podmiot",
            "Erste Leasing S.A. (grupa Erste Bank Polska S.A., wcześniej Santander Leasing S.A.), ul. Kolorowa 8, 60-198 Poznań, KRS 0000026084, NIP 521-12-95-779, nr BDO 000485101.",
        ),
        (
            "Przedmiot audytu",
            "Strona główna /erste-leasing, sekcja /zielone-rozwiazania z 10 podstronami ofertowymi, podstrony usług cyfrowych (/biopodpis, /e-faktura, /leasing-bysms, /jednorazowy-podpis-kwalifikowany), /o-nas, /nagrody oraz załączone dokumenty PDF.",
        ),
        (
            "Sektor działalności",
            "Leasing operacyjny i finansowy, pożyczki oraz najem pojazdów, maszyn i instalacji OZE. Oferta kierowana jednocześnie do firm (B2B) oraz osób fizycznych — konsumentów (B2C).",
        ),
        (
            "Reżim prawny",
            "Prawo obowiązujące: UPNPR (t.j. Dz.U. z 2023 r. poz. 845, art. 5 i 6), u.o.k.k. (t.j. Dz.U. z 2025 r. poz. 1714, art. 24 i 106), UZNK (t.j. Dz.U. z 2026 r. poz. 85, art. 16 i 18), rozporządzenie (UE) 2020/852 (Taksonomia UE). Ustawa z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) wchodzi w życie 27 września 2027 r. — w dniu audytu czarna lista EmpCo nie jest jeszcze prawem polskim.",
        ),
        ("Data audytu", AUDIT_DATE),
        ("Werdykt operacyjny", VERDICT_RED),
    ]
    meta_table = doc.add_table(rows=len(meta), cols=2)
    _lock(meta_table, [3.9, PAGE_W - 3.9])
    _nil_table_borders(meta_table)
    for i, (field_k, value) in enumerate(meta):
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
        _run(pk.add_run(field_k), 8, True, False, NAVY)
        _run(pv.add_run(value), 9, last, False, DANGER if last else DARK)
        _cant_split(meta_table.rows[i])

    # --- strona 2 ---
    start_page()
    band(
        "FDF2F4",
        "BE123C",
        [
            ("Werdykt. " + VERDICT_RED + ".", 12, True, False, NAVY, "pl-PL"),
            (
                "Serwis kieruje ofertę „Zielone rozwiązania” zarówno do konsumentów („Dla Ciebie”), jak i przedsiębiorców („Dla biznesu”), dlatego już w dniu audytu (24 listopada 2026 r.) podlega łącznie art. 5 i 6 UPNPR oraz art. 16 UZNK. "
                "Kwalifikacja ogólnego katalogu pojazdów Multisalon24, ciężarówek LNG/CNG i hybryd pod hasłami transportu „bez zanieczyszczania powietrza spalinami” oraz „flot niespalinowych”, a także niezweryfikowany ekwiwalent „50 drzew / tlen dla ok. 100 osób” w usłudze BioPodpis spełniają przesłanki wprowadzenia w błąd.",
                10.5, False, False, DARK, "pl-PL",
            ),
            (
                "Ustawa z 4 września 2026 r. (Dz.U. z 2026 r. poz. 1353) wdrażająca dyrektywę (UE) 2024/825 wchodzi w życie 27 września 2027 r. "
                "Do tej daty etykiety „eko” w menu oraz hasła „Wspólnie jesteśmy BIO / EKO” stanowią przyszły standard art. 7 pkt 2a, 4a i 4b UPNPR, nie zaś już obowiązujące naruszenie krajowej czarnej listy.",
                10.5, False, False, DARK, "pl-PL",
            ),
        ],
        gap_after=6,
    )

    heading("1", "Analiza wstępna")
    body(
        "Erste Leasing S.A. (poprzednio Santander Leasing S.A.) należy do grupy Erste Bank Polska S.A. i w 2025 r. sfinansowała aktywa netto o wartości 8,7 mld zł. "
        "Na witrynie erste.pl/erste-leasing wyeksponowano dział „Zielone rozwiązania - dla Ciebie, biznesu, środowiska” oraz oznaczono etykietą „eko” cztery narzędzia cyfrowe (Leasing bySMS, OTF, BioPodpis, eFaktura)."
    )
    body(
        "Ponieważ finansowanie aut i instalacji OZE kierowane jest również do osób fizycznych (konsumentów), ocena łączy reżim B2C (UPNPR, u.o.k.k.) z reżimem B2B (UZNK). "
        "Dwa ustalenia o ryzyku wysokim rozstrzygają werdykt już na gruncie prawa obowiązującego 24 listopada 2026 r., natomiast dwa ustalenia średnie dotyczą przejrzystości systemu oceny zielonych transakcji oraz opisu rejestru BDO."
    )

    heading("2", "Ustalenia")
    short_findings = [
        (
            "01",
            CASES[0]["short_title"],
            "art. 5 i 6 UPNPR; art. 16 ust. 1 pkt 2 i ust. 2 UZNK",
            "WYSOKIE",
            "Ogólny katalog aut Multisalon24, LNG/CNG i hybrydy pod szyldem „Zielone rozwiązania” i „bez spalin” bez progów emisyjnych.",
        ),
        (
            "02",
            CASES[1]["short_title"],
            "art. 5 i 6 UPNPR; art. 16 ust. 1 pkt 2 i ust. 2 UZNK",
            "WYSOKIE",
            "Przelicznik drzew i tlenu bez metodyki, pominięcie dojazdu doradcy przy BioPodpisie oraz etykiety „eko” dla obowiązkowego KSeF.",
        ),
        (
            "03",
            CASES[2]["short_title"],
            "art. 16 UZNK; art. 5–6 UPNPR; art. 3 i 8 rozp. 2020/852",
            "ŚREDNIE",
            "Brak kryteriów i skali portfela dla „Systemu Green” przy wyłączeniu leasingu z PCAF i wskaźniku GAR Grupy 3,28%–3,56%.",
        ),
    ]
    headers = ["ID", "Analizowane twierdzenie", "Podstawa prawna", "Ryzyko", "Istota"]
    col_w = [1.05, 4.35, 3.55, 2.35, 5.7]
    sum_table = doc.add_table(rows=1 + len(short_findings), cols=5)
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
        _run(p.add_run(head), 8, True, False, WHITE)
    risk_color = {"BARDZO WYSOKIE": DANGER, "WYSOKIE": DANGER, "ŚREDNIE": WARNING}
    for r_idx, row in enumerate(short_findings):
        bg = "FFFFFF" if r_idx % 2 == 0 else "F7F9FC"
        for c_idx, val in enumerate(row):
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
                _run(p.add_run(val), 8.5, True, False, NAVY)
            elif c_idx == 3:
                _run(p.add_run(val), 8, True, False, risk_color.get(val, DARK))
            else:
                _run(p.add_run(val), 8, False, False, DARK)
        _cant_split(sum_table.rows[r_idx + 1])
    _cant_split(sum_table.rows[0])
    sum_table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    para(
        "Jest to przykład 3 ryzyk spośród 4 zidentyfikowanych w serwisie Erste Leasing S.A. (w tym 2 o poziomie wysokim i 2 o poziomie średnim).",
        size=10, italic=True, color=MUTED, before=12, after=2,
    )

    # --- strona 3–4: Szczegóły Przypadku 01 ---
    heading("3", "Szczegóły przypadków", new_page=True)
    case = CASES[0]
    para(
        "Przypadek " + case["no"] + ".  " + case["short_title"],
        size=13, bold=True, color=NAVY, before=12, after=4, keep_next=True, keep_lines=True, space=1.08,
    )
    rich(
        [
            ("Ryzyko: ", False, False, MUTED, "pl-PL"),
            (case["risk"].lower() + ".", True, False, DANGER, "pl-PL"),
        ],
        size=10.5, before=0, after=6, keep_next=True, keep_lines=True,
    )
    body("Podstawa prawna: " + case["basis"], size=9.5, before=0, after=4, first=0, keep_lines=True)

    label("Cytat źródłowy")
    for src_label, q_text in case["quotes"]:
        band("F4F7FB", "2F8DBC", [(src_label + " " + q_text, 9.8, False, True, DARK, "pl-PL")], gap_after=6)

    label("Ocena prawna")
    for block in case["assessment"]:
        body(block, before=1, after=8, keep_lines=True, first=0.75)

    naprawa = para(
        "Kierunek naprawy",
        size=14, bold=True, color=NAVY, before=0, after=8,
        keep_next=True, keep_lines=True, page_break=True, space=1.05,
    )
    _p_border_bottom(naprawa, "D5DDE3", "12", "8")
    body(case["remediation"], before=4, after=6, keep_lines=True, first=0.75)

    label("Dozwolone sformułowanie")
    band("E8F0FA", "3A61B3", [(case["claim"], 10, False, True, DARK, "pl-PL")], gap_after=4)

    # --- ostatnia strona: Pozostałe przypadki + CTA ---
    pozostale = para(
        "Pozostałe przypadki",
        size=14, bold=True, color=NAVY, before=0, after=10,
        keep_next=True, keep_lines=True, page_break=True, space=1.05,
    )
    _p_border_bottom(pozostale, "D5DDE3", "12", "8")

    previews = [
        (
            "02",
            CASES[1]["short_title"],
            "wysokie",
            DANGER,
            "Przelicznik „50 drzew” i „tlen dla ok. 100 osób” nie ma podanej metodyki, nazwa „BioPodpis” łączona jest z hasłem „Wspólnie jesteśmy BIO” mimo wymogu przyjazdu doradcy (flota w 99% spalinowo-hybrydowa), a eFaktura oznaczana jest jako dobrowolny wkład „eko” po wdrożeniu obowiązkowego KSeF.",
        ),
        (
            "03",
            CASES[2]["short_title"],
            "średnie",
            WARNING,
            "Serwis powołuje autorski „System Oznaczania i Weryfikacji Transakcji Green” w kontekście CSRD i Taksonomii UE bez podania kryteriów i udziału w portfelu, podczas gdy raport skonsolidowany za 2025 r. wyłącza leasing z kalkulacji emisji PCAF i wykazuje wskaźnik GAR Grupy na poziomie 3,28%–3,56%.",
        ),
    ]
    for no, title, r_txt, r_col, summary in previews:
        band(
            "F7F9FC",
            "08181F",
            [
                (no + "     " + title, 11, True, False, NAVY, "pl-PL"),
                ("Ryzyko: " + r_txt + ".", 10, True, False, r_col, "pl-PL"),
                (summary, 9.8, False, False, DARK, "pl-PL"),
            ],
            gap_after=8,
        )

    band(
        "F4F7FB",
        "3A61B3",
        [
            ("Reszta po zakupie pełnej wersji.", 11, True, False, NAVY, "pl-PL"),
            (
                "Pełna analiza wszystkich 4 przypadków naruszeń (w tym nieprawidłowości wokół rejestru BDO w Przypadku 04), ocena ryzyk prawno-finansowych oraz czterostopniowy harmonogram działań naprawczych dostępne są w pełnym raporcie audytu ESG Institute.",
                10.5, False, False, DARK, "pl-PL",
            ),
        ],
        gap_after=0,
    )

    doc.save(SHORT_OUT)
    return SHORT_OUT


def main():
    m = build_material()
    f = build_full_report()
    s = build_short_report()
    print("Generated:", m)
    print("Generated:", f)
    print("Generated:", s)


if __name__ == "__main__":
    main()
