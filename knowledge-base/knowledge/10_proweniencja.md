# Proweniencja i ograniczenia

## Mapowanie `Przepisy.docx` → baza

| Zapis w dokumencie | Jak wykorzystano |
|---|---|
| UPNPR, UOKiK, UZNK | Pełne PDF + PDF-y ISAP użytkownika (`D20071206L`, `D20070331Lj`, `D19930211Lj`); D20071206L = tekst ogłoszony, nie Lj |
| Dyrektywa 2024/825 | Pełny tekst PL, załącznik czarnej listy + FAQ 30.06.2026 + CPC old stock |
| Transpozycja PL | Druk 2799 (rząd, Art. 3 = 2026) → sprawozdanie 3043 (Art. 3 = 2027) → Senat 825 (Art. 3 = 27.09.2027); proces Senatu 2149 |
| UOKiK praktyka | Komunikaty + PDF zarzutów: InPost/Allegro/DHL/DPD (VII 2025); Bolt/Tchibo/Zara (I 2026) |
| „Projekt Dyrektywy Green Claims” (link SWD 52022SC0085) | Pobrano SWD **oraz** COM(2023) 166; status 2026: wniosek wstrzymany |
| 1169/2011, 834/2007, 2018/848, ustawa ekologiczna, EUDR | PDF-y; 834/2007 uchylone; ustawa PL = `D20221370L.pdf` z ISAP |
| SFDR, Taksonomia, KNF, ESMA | PDF-y + huby HTML; z huba KNF dociągnięto raporty ESA 2024 |
| CountEmissionsEU | Rozp. 2026/1030 PL i EN |
| Strategia tekstyliów | Strona + PDF przewodnika + COM(2022) 141 HTML |
| ESPR, CSDDD | PDF 2024/1781 i 2024/1760 + strona MRiT (Omnibus I) |
| CPR, BREEAM, LEED, EPBD | CPR i EPBD w całości; LEED v4.1 BD+C PDF publiczny; BREEAM — FAQ v7.1, manual za loginem |
| UCPD Guidance 2021 | `extracts/ucpd_guidance_2021.txt` (§ 4.1.1); EUR-Lex PDF bywał pusty |

## Korekty względem brzmienia maila w docx

1. Art. 4 ust. 1 dyrektywy (UE) 2024/825 wyznacza **27 września 2026** jako datę stosowania krajowych przepisów transponujących. To nie jest data wejścia w życie ustawy polskiej. Ustawa z 4 września 2026 r. (**Dz.U. 2026 poz. 1353**, ogłoszona 19 października 2026 r.) wchodzi w życie **27 września 2027 r.** (art. 3). Do tej daty nowe zakazy krajowe nie są naruszeniem już stwierdzonym. Zdanie o braku ogłoszenia w Dz.U. jest nieaktualne. Raport: `INSTRUKCJA_GENEROWANIA_RAPORTOW.md`.
2. Link „Green Claims” prowadził do SWD(2022) 85, nie do projektu dyrektywy.
3. 834/2007 nie jest aktualnym reżimem „bio/eko”.
4. Manuale BREEAM NC v7.1 nie są publicznie pobieralne; jest FAQ + LEED v4.1 BD+C PDF.
5. `D20071206L.pdf` z ISAP to **tekst ogłoszony 2007**, nie ujednolicony — do stosowania UPNPR używać `upnpr_2.pdf`.
6. Wytyczne UOKiK o oświadczeniach środowiskowych (IX 2026) **nie są opublikowane** — nie wynajdywać ich treści.

## Licencje i cytowanie

Teksty aktów UE: EUR-Lex, publikacja w Dz.U. UE.  
Ustawy PL: ISAP / Dziennik Ustaw.  
Raporty ESA: copyright UE / ESA.  
Przewodnik Kampania 17 Celów: materiał organizacji; użyty jako źródło opisowe strategii, nie jako prawo.

Przy budowie narzędzia AI: cytować **artykuł + plik w `extracts/`**, nie parafrazować kar i definicji z pamięci modelu.

## Social washing, wpis 30 października 2026

Teksty otwarte w API Sejmu albo w Publications Office (cellar) i na nocie EUR-Lex. Do `12_social_washing.md` weszły publikatory, daty i te artykuły, które odczytano.

| Akt | Skąd |
|---|---|
| UPNPR, Dz.U. 2023 poz. 845 | `extracts/sw_upnpr_DU2023_845.html` |
| Dyrektywa 2005/29/WE | `originals/sw_ucpd_32005L0029.pdf` |
| Dyrektywa (UE) 2019/2161 | `originals/sw_omnibus_32019L2161.pdf`; nota EUR-Lex: L 328 z 18.12.2019, in force |
| Dyrektywa 2006/114/WE | `originals/sw_mcad_32006L0114.pdf` |
| Kodeks cywilny, Dz.U. 2026 poz. 795 | `originals/sw_kc_DU2026_795.pdf` |
| Ustawa o języku polskim, Dz.U. 2026 poz. 81 i nowela poz. 676 | `originals/sw_jezyk_DU2026_81.pdf`, `sw_jezyk_DU2026_676.pdf` |
| Równe traktowanie, Dz.U. 2026 poz. 1196 | `originals/sw_rowne_DU2026_1196.pdf` |
| Prawa konsumenta, Dz.U. 2026 poz. 1244 | `originals/sw_upk_DU2026_1244.pdf` |
| Dyrektywa 2011/83/UE, konsolidacja 02011L0083-20220528 | `originals/sw_crd_02011L0083-20220528.pdf`. Cytat operacyjny: ustawa krajowa |
| RODO, 32016R0679 | `extracts/sw_gdpr_art_sample.html` |
| Ustawa o ochronie danych osobowych | Dz.U. 2018 poz. 1000 i tekst jednolity Dz.U. 2023 poz. 1206 |
| Konwencja o prawach dziecka | Dz.U. 1991 nr 120 poz. 526 |
| Radiofonia i telewizja | Dz.U. 2022 poz. 1722 |
| Ochrona małoletnich | Dz.U. 2026 poz. 110 |
| Kodeks pracy | Dz.U. 2026 poz. 1245, w tym odnośniki do Dz.U. 2026 poz. 1046 |

Czego ten wpis nie potwierdza: artykułów CSRD, ESRS, ustawy o sygnalistach, dyrektywy (UE) 2023/970, PPK, prawa zamówień publicznych, konwencji MOP ani standardów dobrowolnych innych niż Wytyczne OECD. Dyrektywy 89/391/EWG nie otwierano; deklarację BHP przykłada się do Kodeksu pracy. Art. 13 RODO i ustawy o prawie autorskim nie cytowano zdanie po zdaniu. Próg zgody dziecka: w otwartych tekstach ustawy o ochronie danych osobowych nie ma obniżenia wieku z art. 8 RODO.

## Wytyczne OECD, dopisek 30 października 2026

Otwarte: polski przekład edycji 2023 (`originals/sw_oecd_wytyczne_2023_PL.pdf`, strona MFiPR), procedura KPK z 2024 r. (`originals/sw_kpk_procedura_2024.pdf`) oraz oceny: Valeo 24.01.2024, OLX 13.06.2019, LOTOS 17.02.2023, PGE 30.09.2025, LPP 27.05.2026. Numerację punktów ze spraw sprzed edycji 2023 przepisano tak, jak stoi w ocenie KPK. Nie przenumerowano jej na edycję 2023.
