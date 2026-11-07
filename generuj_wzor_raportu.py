#!/usr/bin/env python3
"""Buduje wzory/WZOR_raport.docx — skład, nie treść do przeniesienia na inny podmiot."""

from pathlib import Path

from sklad_raportu import VERDICT_RED, ReportDoc

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "wzory" / "WZOR_raport.docx"


def build():
    report = ReportDoc(
        reference="Sygn. ESG/RRRR/MM/DD",
        document_name="Audyt greenwashingu — nazwa podmiotu",
        audit_date="data audytu",
    )
    report.doc.core_properties.subject = "Wzór formy raportu audytu greenwashingu"
    report.title("Audyt Ryzyk Greenwashingu: NAZWA PODMIOTU")
    report.subtitle(
        "Wzór formy. Dwie linie zakresu: co zweryfikowano, z jakiego źródła i na jaki stan. "
        "Kalendarz ustaw nie wchodzi do podtytułu. Treści w nawiasach ostrych nie są ustaleniami."
    )
    report.meta(
        [
            ("Badany podmiot", "<nazwa, identyfikatory, adres, osoba wskazana w materiale>", False),
            ("Przedmiot audytu", "<dokument albo adresy podstron i data zebrania>", False),
            ("Sektor działalności", "<sektor i zdanie, czy jest to usługa wobec konsumenta>", False),
            (
                "Reżim prawny",
                "Lista aktów faktycznie zastosowanych, oddzielona średnikami. "
                "Jedno zdanie, czy czarna lista EmpCo jest w dniu audytu prawem polskim. "
                "Daty stosowania przenosi się do pkt 1.1.",
                False,
            ),
            ("Data audytu", "<dzień miesiąca rok>", False),
            ("Werdykt operacyjny", VERDICT_RED, True),
        ]
    )
    report.verdict(
        VERDICT_RED,
        [
            "Dwa do czterech zdań. Najpierw przepis obowiązujący w dniu audytu. "
            "Przyszłego standardu EmpCo nie opisuje się jako naruszenia już stwierdzonego.",
            "Etykiety operacyjne, bez kodów angielskich: ryzyko wysokie i wstrzymanie publikacji; "
            "publikacja warunkowa po korekcie; dopuszczone do publikacji.",
            "Ta ramka jest jedyną ramką werdyktu. Obramowanie jest cienkie. Nie dodaje się znaku kwadratu.",
        ],
    )
    report.h1("1. Analiza wstępna")
    report.h2("1.1. Reżim prawny")
    report.paragraph(
        "Tu wchodzi polemika o datach: art. 4 ust. 1 dyrektywy (UE) 2024/825 wyznacza 27 września 2026 r. "
        "jako datę stosowania przepisów krajowych, a ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353) "
        "wchodzi w życie 27 września 2027 r. Do tej daty werdykt opiera się na przepisach już obowiązujących."
    )
    report.paragraph(
        "Akapit ma cztery do sześciu zdań i jest wyjustowany. Kolejny argument dostaje osobny akapit. "
        "Forma jest bezosobowa: „ocenę przeprowadzono”, „w ramach audytu nie poddano analizie”."
    )
    report.h2("1.2. Profil działalności")
    report.paragraph("<Czym zajmuje się podmiot i która narracja spełnia przesłanki przepisu obowiązującego.>")
    report.h2("1.3. Materiał z witryny")
    report.paragraph("<Jeżeli materiał pochodzi ze strony, zrzut stoi tutaj albo w załączniku, nie nad metryczką.>")
    report.figure(
        None,
        "Rys. 1. Zrzut ekranu <witryna> (stan na <data>). Źródło: opracowanie własne.",
    )
    report.h2("1.4. Wyłączenia")
    report.paragraph("To, co sprawdzono i nie jest ustaleniem, zapisuje się krótszymi punktami.")
    report.bullet("Projekty i porównania. ", "<Zdanie, czy materiał porównuje konkurenta albo spełnia test social washingu.>")
    report.bullet("Wytyczne OECD. ", "Same wytyczne mają etykietę niską i nie podnoszą werdyktu.")

    report.h1("2. Analiza szczegółowa — tabela zbiorcza naruszeń")
    report.findings(
        [
            (
                "01",
                "<skrót twierdzenia>",
                "<przepis obowiązujący. Przyszły standard EmpCo, jeżeli dotyczy, w osobnym zdaniu.>",
                "WYSOKIE",
                "<istota w jednym zdaniu>",
            )
        ]
    )

    report.h1("3. Szczegółowa analiza przypadków")
    report.case(
        "Przypadek 01. <tytuł twierdzenia>",
        "<przepis obowiązujący, publikator>",
        "WYSOKIE",
    )
    report.paragraph("Cytat źródłowy", size=10.5, bold=True, before=2, after=2, align="left")
    report.quote("Miejsce w materiale:", "„<dosłowny cytat>”")
    report.paragraph("Ocena prawna", size=10.5, bold=True, before=6, after=2, align="left")
    report.paragraph(
        "<Akapit o przesłankach przepisu i o teście decyzji dotyczącej transakcji, gdy podstawą jest art. 5 lub 6.>"
    )
    report.paragraph(
        "<Osobny akapit o orzeczeniu, tylko gdy zmienia kwalifikację. Omówienie tezy, nie cytat, "
        "chyba że zdanie sprawdzono dosłownie. Wyrok nie dotyczy audytowanego podmiotu.>"
    )
    report.paragraph("Kierunek naprawy", size=10.5, bold=True, before=2, after=2, align="left")
    report.paragraph("<Co zdjąć, co zawęzić, czego nie przenosić na inną formę usługi.>")
    report.claim("„<konkretny wzór po korekcie albo wyjaśnienie, że twierdzenia nie da się uratować>”")

    report.h1("4. Ryzyka prawne i finansowe")
    report.paragraph(
        "Sekcja porównania wchodzi przed ryzyka tylko wtedy, gdy materiał porównuje. "
        "Nie zostawia się pustego numeru. Nagłówki ryzyk są wielopoziomowe, bez kropki i cyfry naraz."
    )
    report.h2("4.1. <ryzyko wynikające z ustaleń>")
    report.paragraph("<Akapit. Kara Prezesa UOKiK do 10% obrotu tylko przy praktyce naruszającej prawo już obowiązujące.>")
    report.h2("4.2. <kolejne ryzyko>")
    report.paragraph("<Akapit.>")

    report.h1("5. Rekomendacje naprawcze")
    report.paragraph(
        "Cztery etapy są nagłówkami, nie kolorowymi ramkami. Etap natychmiastowy obejmuje tylko prawo już naruszone. "
        "Twierdzenie, które stanie się zakazane 27 września 2027 r., dostaje etap przed tą datą."
    )
    report.h2("5.1. Etap natychmiastowy, <termin od daty audytu>")
    report.paragraph("<Treść dopasowana do ustaleń.>")
    report.h2("5.2. Korekta materiałów, <termin>")
    report.paragraph("<Treść.>")
    report.h2("5.3. Metodyka i dowody, <termin>")
    report.paragraph("<Treść.>")
    report.h2("5.4. Sformułowania i szkolenie, <termin>")
    report.paragraph("<Treść.>")

    report.h1("6. Wnioski")
    report.paragraph(
        "<Jedno albo dwa zdania. Stan prawny dnia audytu osobno od 27 września 2027 r., "
        "jeżeli przyszły standard ma znaczenie.>"
    )
    report.sign_off()
    report.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
