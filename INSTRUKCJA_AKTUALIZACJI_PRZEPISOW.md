# Instrukcja aktualizacji przepisów

Ten plik jest wiążącą procedurą komendy **Sprawdź nowe przepisy**. Przed odpowiedzią odczytaj go i wykonaj. Dotyczy też prośby wprost równoważnej: sprawdź, czy od ostatniej aktualizacji bazy doszły nowe przepisy o twierdzeniach środowiskowych.

Samo sprawdzenie niczego nie zmienia. Do odpowiedzi użytkownika nie edytuj `knowledge-base/`, `rules.json`, `INSTRUKCJA_GENEROWANIA_RAPORTOW.md` ani reguły raportu. Wdrożenie zaczyna się dopiero po zgodzie.

Przy pisaniu raportu audytu stosuj `INSTRUKCJA_GENEROWANIA_RAPORTOW.md`. Tej procedury przy raporcie nie uruchamiaj.

# CO JEST „NOWYM PRZEPISEM”

Pozycja trafia do odpowiedzi, gdy spełnia oba warunki:

1. Jest ustanowiona albo ogłoszona po dacie `knowledge/rules.json` → `meta.as_of` (albo jest zmianą aktu z listy nadzorowanej po tej dacie).
2. Jej publikatora nie ma jeszcze w `knowledge/`, w `rules.json` ani w `INSTRUKCJA_GENEROWANIA_RAPORTOW.md`.

Ustawa z 4 września 2026 r. (Dz.U. 2026 poz. 1353, ogłoszenie 19 października 2026 r., wejście w życie 27 września 2027 r.) jest już opisana. Nie zgłaszaj jej ponownie.

Rozdziel cztery kategorie i nie nazywaj wytycznych ustawą:

| Kategoria | Co to jest | Dokąd wchodzi po zgodzie |
|---|---|---|
| Akt ustanowiony | Ustawa, rozporządzenie, dyrektywa ogłoszone w Dzienniku Ustaw albo Dzienniku Urzędowym UE | Baza i — gdy zmienia reżim oceny twierdzeń — instrukcja raportu |
| Zmiana aktu z listy | Nowelizacja, akt delegowany albo wykonawczy do pozycji z listy nadzorowanej | Plik sektorowy, `rules.json` tylko gdy zmienia test claimu |
| Tekst jednolity | Obwieszczenie o tekście jednolitym bez nowej normy | `10_proweniencja.md` i `INDEX.md`: który plik jest tekstem do stosowania |
| Praktyka | Komunikat UOKiK, FAQ Komisji, CPC, wytyczne ESMA albo KNF | `11_praktyka_nadzorcza.md`; do instrukcji raportu tylko, gdy ma zmienić werdykt |

Wniosek, który nie został przyjęty, zostaje wnioskiem. Green Claims, COM(2023) 166, nie staje się prawem przez samą wzmiankę w prasie.

# LISTA NADZOROWANA

Sprawdzaj te identyfikatory. Nowy akt spoza listy dodaj tylko wtedy, gdy dotyczy dobrowolnego twierdzenia środowiskowego, certyfikatu, oznakowania ekologicznego albo nazwy produktu zrównoważonego.

| ID w bazie | Identyfikator | Adres sprawdzenia |
|---|---|---|
| upnpr | ELI DU/2007/1206 | `https://api.sejm.gov.pl/eli/acts/DU/2007/1206` |
| uokk | ELI DU/2007/331 | `https://api.sejm.gov.pl/eli/acts/DU/2007/331` |
| uznk | ELI DU/1993/211 | `https://api.sejm.gov.pl/eli/acts/DU/1993/211` |
| empco_pl | ELI DU/2026/1353 | `https://api.sejm.gov.pl/eli/acts/DU/2026/1353` |
| organic_pl | ELI DU/2022/1370 | `https://api.sejm.gov.pl/eli/acts/DU/2022/1370` |
| ecgt | CELEX 32024L0825 | EUR-Lex, nota i akty zmieniające |
| ucpd | CELEX 32005L0029 | EUR-Lex |
| green_claims_proposal | CELEX 52023PC0166 | EUR-Lex, status procedury |
| fir_1169 | CELEX 32011R1169 | EUR-Lex |
| organic_848 | CELEX 32018R0848 | EUR-Lex |
| eudr | CELEX 32023R1115 | EUR-Lex |
| sfdr | CELEX 32019R2088 | EUR-Lex |
| taxonomy | CELEX 32020R0852 | EUR-Lex |
| countemissions | CELEX 32026R1030 | EUR-Lex |
| espr | CELEX 32024R1781 | EUR-Lex, w tym akty delegowane |
| csddd | CELEX 32024L1760 | EUR-Lex |
| cpr | CELEX 32024R3110 | EUR-Lex |
| epbd | CELEX 32024L1275 | EUR-Lex |
| ucpd_g | CELEX 52021XC1229(05) | Czy notice został zastąpiony po EmpCo |

# SPRAWDZENIE

Otwórz źródło. Przepisz tytuł, publikator, datę ogłoszenia i datę wejścia w życie z odpowiedzi. Nie uzupełniaj braków z pamięci. Gdy strona albo API się nie otwiera, napisz, którego adresu nie potwierdzono, i nie zgaduj treści.

**Polska.** Dla każdego roku od roku `meta.as_of` do roku bieżącego pobierz wyszukiwanie API Sejmu i zostaw pozycje z `promulgation` późniejszym niż `meta.as_of`:

- `https://api.sejm.gov.pl/eli/acts/search?publisher=DU&year=ROK&title=nieuczciw&limit=20`
- to samo z `title=ekologiczn`
- to samo z `title=konsument`

Zostaw ustawę albo rozporządzenie, którego tytuł dotyczy nieuczciwych praktyk, praw konsumenta, konkurencji, rolnictwa ekologicznego albo oświadczeń środowiskowych. Odrzuć obwieszczenia i akty już opisane w bazie. Dla pozycji nowej otwórz `https://api.sejm.gov.pl/eli/acts/DU/ROK/POZ` i odczytaj `title`, `displayAddress`, `promulgation`, `entryIntoForce`, `status`.

**Unia.** Dla każdego CELEX z listy otwórz notę na EUR-Lex (`https://eur-lex.europa.eu/legal-content/PL/ALL/?uri=CELEX:NUMER`). Zapisz akt zmieniający, akt delegowany albo nową konsolidację z datą późniejszą niż `meta.as_of`, jeżeli tego CELEX-u nie ma w bazie. Osobno sprawdź status procedury COM(2023) 166.

**Praktyka.** Wyszukaj komunikaty UOKiK, FAQ Komisji do dyrektywy 2024/825, stanowiska CPC oraz wytyczne ESMA nowsze niż `meta.as_of`. Wytyczne UOKiK o oświadczeniach środowiskowych były nieopublikowane we wrześniu 2026 r. Jeżeli nadal ich nie ma, napisz to jednym zdaniem, bez streszczania domniemanej treści.

# ODPOWIEDŹ

Pisz po polsku, krótko. Najpierw data sprawdzenia i data bazy (`meta.as_of`). Potem lista albo zdanie, że w otwartych źródłach nie ma nowego aktu.

Każda pozycja:

1. Tytuł i publikator (Dz.U. albo CELEX), data ogłoszenia, data wejścia w życie albo stosowania.
2. Status w jednym określeniu: obowiązuje; ogłoszony, jeszcze się nie stosuje; wniosek; wytyczne.
3. Czego dotyczy przy twierdzeniach środowiskowych: sektor (horyzontalne, FMCG, finanse, transport, moda, budownictwo) i czy zmienia werdykt raportu, czy tylko doprecyzowuje akt już stosowany.
4. Które pliki dostałyby zapis po zgodzie.

Na końcu, gdy lista nie jest pusta, zadaj jedno pytanie i zatrzymaj się:

„Czy wdrożyć te pozycje do bazy wiedzy i do instrukcji raportu? Można zatwierdzić całość, same numery, tylko bazę albo nic.”

Nie proponuj gotowej treści reguł w tym pytaniu. Gdy lista jest pusta, nie pytaj o wdrożenie. Napisz, które źródła się otworzyły.

# WDROŻENIE PO ZGODZIE

Wykonaj tylko pozycje, które użytkownik wskazał.

Dla każdej potwierdź jeszcze raz tekst w API Sejmu, ISAP albo EUR-Lex. Do plików wchodzi publikator, data i skutek dla claimu. Nie dopisuj numeru artykułu, którego nie odczytałeś.

Kolejność zapisu:

1. `knowledge-base/sources/manifest.json` — pozycja albo `note`, z adresem źródła.
2. `knowledge-base/INDEX.md` — wiersz katalogu.
3. Plik merytoryczny: `00_mapa_i_hierarchia.md` przy dacie albo hierarchii; `01`–`07` przy sektorze; `11_praktyka_nadzorcza.md` przy praktyce; `02_zakazane_praktyki_ecgt.md` przy czarnej liście EmpCo.
4. `knowledge/10_proweniencja.md` — skąd pochodzi tekst i czego nie potwierdzono.
5. `knowledge/rules.json` — podbij `meta.as_of` na dzień wdrożenia. Regułę claimu dopisz albo zmień tylko wtedy, gdy nowy przepis zmienia test już obecny w `immediate_reject`, `always_on_tests` albo `sector_rules`. Data ogłoszenia sama nie jest regułą.
6. `knowledge-base/README.md` — zdanie o stanie katalogu.
7. Instrukcja raportu i reguła `.cursor/rules/greenwashing-raport.mdc` — tylko gdy pozycja zmienia datę, reżim werdyktu albo podstawę, którą raport ma stosować. Zdanie w raporcie ma rozdzielać prawo już obowiązujące od przepisu ogłoszonego, który jeszcze się nie stosuje.
8. Gotowych raportów DOCX i MD nie przerabiaj. Nowy reżim obowiązuje od następnego audytu, chyba że użytkownik wskaże konkretny plik.

Po zapisie wymień zmienione pliki i datę `meta.as_of`.
