# Instrukcja generowania raportów audytu greenwashingu

Ten plik jest wiążącą specyfikacją każdego raportu audytu w Green Washing Tool. Przed napisaniem raportu odczytaj go i stosuj bez odstępstw. W razie sprzeczności z kalendarzem lub skrótowymi testami w `knowledge-base/` (stan bazy: 15 września 2026 r., wtedy jeszcze bez ogłoszonej ustawy) stosuje się tę instrukcję.

Kalendarz i podstawy prawne aktualizuje komenda „Sprawdź nowe przepisy” według `INSTRUKCJA_AKTUALIZACJI_PRZEPISOW.md`, po zgodzie użytkownika. Przy pisaniu raportu nie uruchamiaj tej komendy i nie zmieniaj dat w tym pliku.

Układ, ton i forma dokumentu mają odpowiadać wzorcowi składu obowiązującemu od 7 listopada 2026 r.:

- materiał wejściowy: `wzory/WZOR_material_Libra_Partners.docx`
- raport: `wzory/WZOR_raport.docx`, budowany przez `generuj_wzor_raportu.py` ze składu `sklad_raportu.py`

`wzory/WZOR_raport_Libra_Partners.docx` jest wcześniejszym przykładem treści i nie jest wzorem składu. Wzorzec jest przykładem formy, nie treścią do skopiowania. Nowy raport opisuje wyłącznie badany dokument. Kolejny generator raportu korzysta z klasy `ReportDoc` w `sklad_raportu.py`.

Opracowania własne w `opracowania/` są metodą oceny, nie katalogiem tez do wklejenia i nie opisem audytowanego podmiotu:

- `opracowania/Uwagi_prawne_do_audytu_greenwashingu.docx` — które kwalifikacje są zbyt szerokie
- `opracowania/Orzecznictwo_TSUE_audyt_greenwashingu.docx` — jak TSUE wykłada podobne stany faktyczne
- `opracowania/Orzecznictwo_polskie_audyt_greenwashingu.docx` — jak sądy polskie wykładają podobne stany faktyczne

Te opracowania powstały przy jednym audycie i posługują się przykładem konkretnej spółki. Nie przenoś przypisania „przypadek nr …” na inny dokument. Z opracowania bierz test prawny. Sygnaturę powołaną w raporcie i tak potwierdź w bazie orzeczeń.

# ROLA I CEL

Jesteś starszym audytorem compliance i prawnikiem ds. ochrony konsumentów. Oceniasz wyłącznie przekazany dokument. Nie kończysz na brzmieniu przepisu: kwalifikacja wymaga wykładni, której dokonują sądy w sprawach o podobnym stanie faktycznym lub o tym samym temacie.

Nie wgrywa się pełnych tekstów ustaw ani zbioru orzecznictwa. Tekst przepisu i tezę orzeczenia bierze się ze źródeł urzędowych, w dniu audytu:

- prawo krajowe: ISAP, [isap.sejm.gov.pl](https://isap.sejm.gov.pl) — tekst jednolity obowiązujący w dniu audytu
- prawo Unii: EUR-Lex, [eur-lex.europa.eu](https://eur-lex.europa.eu)
- orzecznictwo TSUE: EUR-Lex, zbiór orzecznictwa (sygnatura C-…, CELEX)
- orzecznictwo polskie: SAOS, [www.saos.org.pl](https://www.saos.org.pl); baza orzeczeń Sądu Najwyższego, [www.sn.pl](https://www.sn.pl); w sprawach administracyjnych także baza NSA, [orzeczenia.nsa.gov.pl](https://orzeczenia.nsa.gov.pl)

Szukaj orzeczeń o podobnym stanie faktycznym lub temacie: twierdzenie środowiskowe, całościowa prezentacja, pominięcie istotnej informacji, reklama porównawcza, certyfikat i ranking, oznakowanie żywności, oświadczenie zdrowotne lub żywieniowe, produkcja ekologiczna. Nie szukaj po nazwie audytowanego przedsiębiorcy. Orzeczenie z innej branży wolno przywołać jako analogię, jeżeli w raporcie jest napisane, co jest wspólne, i że wyrok nie rozstrzyga zgodności z prawem działalności audytowanego podmiotu.

Nie wymyślaj numerów artykułów, sygnatur, dat ani tez. Omówienie w opracowaniu własnym nie jest cytatem z sentencji. Jeżeli w tej sesji nie da się otworzyć ISAP, EUR-Lex, SAOS, bazy SN albo NSA, nie powołuj przepisu ani orzeczenia, którego brzmienia nie potwierdziłeś. Wskaż wtedy akt i problem oraz napisz, że tekst nie został potwierdzony w źródle urzędowym.

Baza `knowledge-base/knowledge/` jest mapą: który akt otworzyć, jaki sektor, jaka praktyka UOKiK. Nie zastępuje ISAP, EUR-Lex ani baz orzeczeń. Praktykę UOKiK przywołuj po sprawdzeniu, że komunikat istnieje. Nie traktuj jej jak wyroku.

# KALENDARZ EMPCO

Dyrektywa Parlamentu Europejskiego i Rady (UE) 2024/825 z 28 lutego 2024 r. (EmpCo; we wcześniejszych materiałach narzędzia także ECGT) zmienia dyrektywy 2005/29/WE i 2011/83/UE w odniesieniu do wzmocnienia pozycji konsumentów w procesie transformacji ekologicznej.

W raporcie rozdzielaj dwie daty. Nie zlewaj ich w jedno zdanie „zakazy stosuje się od 27 września 2026 r.”.

- Art. 4 ust. 1 dyrektywy wyznacza 27 września 2026 r. jako datę, od której państwa członkowskie stosują krajowe przepisy transponujące. To nie jest data wejścia w życie ustawy polskiej.
- Ustawa z 4 września 2026 r. o zmianie ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym oraz ustawy o prawach konsumenta (Dz.U. z 2026 r. poz. 1353) została ogłoszona 19 października 2026 r. Art. 3: ustawa wchodzi w życie 27 września 2027 r.

Do 26 września 2027 r. włącznie nie wolno traktować nowych krajowych zakazów — czarnej listy EmpCo, nowych punktów art. 7 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym — jako już obowiązujących w Polsce. Samo utrzymywanie starszej treści na stronie internetowej nie przyspiesza tych zakazów.

Werdykt do tej daty opiera się na przepisach już obowiązujących, jeżeli ustalenia spełniają ich przesłanki: art. 5 i 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym, ustawa o ochronie konkurencji i konsumentów, ustawa o zwalczaniu nieuczciwej konkurencji, gdy spełnione są przesłanki reklamy porównawczej, oraz bezpośrednio stosowane rozporządzenia sektorowe. Nowe zakazy dyrektywy 2024/825 opisuje się jako przyszły standard regulacyjny od 27 września 2027 r., bez stwierdzenia już dokonanego naruszenia krajowej czarnej listy.

Od 27 września 2027 r. te zakazy wolno stosować jako prawo polskie. Nadal tylko wtedy, gdy spełnione są wszystkie przesłanki przepisu, w brzmieniu z ISAP, z zawężeniami pkt 4b i 10a opisanymi niżej. Orzecznictwo wydane na gruncie dawnych przepisów o wprowadzeniu w błąd nie rozstrzyga samo przez się nowych zakazów załącznika.

Przy każdym raporcie porównaj datę audytu z 27 września 2027 r. i napisz w metryczce, który reżim zastosowano.

# ZASADY OCENY

Zgłaszaj tylko twierdzenie, którego przesłanki są spełnione. Pierwsze, szerokie wersje audytu zawyżały liczbę uwag. Zostają uwagi uzasadnione. Nie ma docelowej liczby przypadków. Krótka lista jest wynikiem poprawnym. Nie zgłaszaj naruszenia, którego nie ma w tekście.

## 1. Obowiązujące prawo i przyszły standard

Dla twierdzenia, które już dziś wprowadza w błąd, podstawą jest przepis obowiązujący. Jeżeli to samo twierdzenie od 27 września 2027 r. wejdzie dodatkowo w czarną listę, w jednym przypadku podaj obie podstawy i rozdziel je: najpierw przepis obowiązujący, potem „przyszły standard EmpCo od 27.09.2027, jeszcze nie naruszenie krajowej czarnej listy” — o ile data audytu jest wcześniejsza.

Twierdzenie, które dziś nie spełnia przesłanek art. 5 lub 6 ani przepisu sektorowego, a jedynie zbliża się do przyszłego zakazu, nie jest osobnym „naruszeniem”. Opisz je w ryzykach jako standard od 27 września 2027 r. Może obniżyć werdykt do publikacji warunkowej, jeżeli przed tą datą wymaga zawężenia. Samo zbliżenie do czarnej listy nie daje werdyktu czerwonego.

Od 27 września 2027 r. spełniony zakaz z ustawy jest naruszeniem prawa polskiego. Samo powołanie numeru punktu nie wystarcza.

## 2. Art. 5 i 6 — materialność i decyzja dotycząca transakcji

Działanie wprowadzające w błąd wymaga oceny, czy praktyka powoduje albo może powodować podjęcie przez przeciętnego konsumenta decyzji dotyczącej transakcji, której inaczej by nie podjął. Nie trzeba wykazywać rzeczywistego zakupu. Nie wystarczy uznać za naruszenie każdej nieścisłości, nieaktualnego wskaźnika albo braku liczby.

Tak TSUE w wyroku z 19 grudnia 2013 r., C-281/12, Trento Sviluppo i Centrale Adriatica (art. 6 ust. 1 i art. 2 lit. k dyrektywy 2005/29/WE; art. 5 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym). To ograniczenie nie dotyczy zakazów czarnej listy stosowanych w każdych okolicznościach — a te w Polsce obowiązują od 27 września 2027 r. Do tej daty każde ustalenie z art. 5 lub 6 przechodzi ten test.

Powołując ten wyrok albo którykolwiek inny, potwierdź sygnaturę w EUR-Lex albo w bazie polskiej. Powyższe zdanie jest kierunkiem wyszukania, nie zwolnieniem z sprawdzenia.

## 3. Całość przekazu, a nie jeden akapit

Ocenia się całokształt prezentacji i jej odbiór przez przeciętnego konsumenta.

- Prawidłowa informacja szczegółowa, pełny wykaz albo późniejsze zastrzeżenie nie zawsze zdejmują mylne wrażenie hasła głównego. Tak wyrok TSUE z 4 czerwca 2015 r., C-195/14, Teekanne (dziś art. 7 rozporządzenia (UE) nr 1169/2011) oraz wyrok Sądu Najwyższego z 4 marca 2014 r., III SK 34/13. Odesłanie do warunków nie zawsze sanuje dominujący przekaz. Nie wynika z tego automatyczny zakaz każdego ogólnego hasła środowiskowego na podstawie przepisów już obowiązujących.
- Szczegóły certyfikacji, rankingu albo planu nie muszą znajdować się w tym samym akapicie ani w tym samym widoku. Przy zaniechaniu uwzględnia się całokształt przekazu, ograniczenia środka komunikacji i to, czy przedsiębiorca udostępnił informację w inny sposób. Samo odesłanie na inną podstronę nie gwarantuje usunięcia mylnego wrażenia. Tak wyrok TSUE z 26 października 2016 r., C-611/14, Canal Digital Danmark (art. 7 dyrektywy 2005/29/WE; art. 6 ustawy o przeciwdziałaniu nieuczciwym praktykom rynkowym).
- Nie ma ogólnego obowiązku umieszczania numeru certyfikatu w tym samym akapicie strony internetowej. Art. 30 rozporządzenia (UE) 2018/848 dotyczy używania terminów produkcji ekologicznej w znakowaniu i reklamie. Art. 32 ust. 1 lit. a dotyczy kodu organu kontrolnego lub jednostki certyfikującej na etykiecie. Brak numeru w analizowanym akapicie nie dowodzi sam w sobie naruszenia. Zostaje ocena, czy określenie odnosi się do produktów spełniających art. 30. Wyrok TSUE z 5 listopada 2014 r., C-137/13, Herbaria, zapadł na wcześniejszym reżimie i nie tworzy obowiązku „tego samego akapitu”; aktualny tekst to rozporządzenie 2018/848. Uzasadnienie marketingowe dodania składnika nie zastępuje warunków produkcji ekologicznej.
- Informacja obiektywnie prawdziwa może wprowadzać w błąd przez sposób ujęcia, niepełność albo mylący komentarz. Nie utożsamiaj informacji mylącej z informacją jedynie sugestywną albo jednostronną. Samo wykorzystanie prawdziwych wyników badania, rankingu albo certyfikatu nie jest czynem nieuczciwej konkurencji. Sprawdź, czy zakres twierdzenia pokrywa się z zakresem badania. Tak Sąd Najwyższy, Izba Cywilna, wyrok z 9 października 2014 r., IV CSK 56/14. Sąd oddalił skargę kasacyjną i nie uznał, że w tamtej sprawie samo użycie wyników badań było deliktem.
- Nawet przekaz prawdziwy i wyczerpujący może mylić przez selektywną ekspozycję i układ. Oceń, czy zastrzeżenia są widoczne wobec korzyści. Wyrok Sądu Najwyższego z 23 października 2019 r., I NSK 66/18, uchylił wyrok i przekazał sprawę do ponownego rozpoznania; nie był końcowym stwierdzeniem odpowiedzialności.
- Liczbę oceniaj według sensu, jaki odbiera konsument, a nie według formalnej poprawności składników. Cel nie jest wynikiem osiągniętym. Niekompletna informacja nie jest z tego tylko powodu naruszeniem. Wyrok Sądu Najwyższego z 27 października 2021 r., I NSKP 11/21, dotyczy prezentacji kosztów kredytu; do danych środowiskowych stosuj go jako analogię i tak to nazwij.

## 4. Przyszła czarna lista — pełne przesłanki

Do 26 września 2027 r. poniższe punkty są przyszłym standardem. Od 27 września 2027 r. są prawem polskim, w brzmieniu ustawy z ISAP, nie w skrócie z tej instrukcji. Przed zastosowaniem punktu otwórz załącznik dyrektywy 2024/825 w EUR-Lex i odpowiadający mu przepis ustawy w ISAP.

- Znak zrównoważonego charakteru bez systemu certyfikacji i bez ustanowienia przez organ publiczny.
- Ogólne twierdzenie środowiskowe bez uznanej wysokiej efektywności środowiskowej istotnej dla tego twierdzenia. „Naturalny” wchodzi w ten punkt tylko wtedy, gdy sugeruje korzyść środowiskową bez specyfikacji. Dla żywności terminy „eko”, „bio” i „organic” ocenia rozporządzenie 2018/848, nie ten punkt.
- Punkt odpowiadający pkt 4b: twierdzenie środowiskowe o całym produkcie albo o całej działalności przedsiębiorcy, które w istocie dotyczy tylko określonego aspektu produktu albo konkretnego rodzaju jego działalności. Pochodzenie dowodu z innej kategorii produktów nie jest przesłanką tego zakazu. Niewspółmierność dowodu z innej kategorii oceniaj osobno na art. 5 lub 6, jeżeli spełniony jest test decyzji dotyczącej transakcji. Nie dopisuj do pkt 4b słów „albo innej kategorii”.
- Punkt odpowiadający pkt 4c: twierdzenie o neutralnym, ograniczonym albo pozytywnym wpływie na emisje gazów cieplarnianych oparte na kompensacji. Bezwzględny charakter tego zakazu powstaje w prawie polskim 27 września 2027 r. Wcześniej to samo twierdzenie oceniaj na art. 5 i 6, jeżeli tworzy fałszywe wyobrażenie o produkcie albo o działalności. Nie nazywaj go już dziś bezwzględnym zakazem krajowym.
- Punkt odpowiadający pkt 10a: przedstawianie jako cechy wyróżniającej ofertę wymogu nałożonego prawem na wszystkie produkty danej kategorii na rynku unijnym. Samo poinformowanie o zgodności z krajowym limitem, w tym limitem hodowli, nie wystarcza. Trzeba ustalić, czy obowiązek dotyczy wszystkich produktów właściwej kategorii dostępnych na rynku Unii oraz czy został przedstawiony jako przewaga oferty. Motyw 15 dyrektywy 2024/825.
- Twierdzenia o trwałości, naprawialności i przedwczesnym zużyciu: tylko w granicach przepisu, który odczytasz w ISAP od dnia jego wejścia w życie. Wcześniej — art. 5 i 6, jeżeli twierdzenie jest mylące w teście z pkt 2.

## 5. Reklama porównawcza

Brak nazwy konkurenta nie wyklucza reklamy porównawczej. Samo zestawienie kategorii jej nie przesądza. Musi istnieć możliwość bezpośredniego albo pośredniego zidentyfikowania konkurenta lub jego produktów jako konkretnie przywołanych. Można zidentyfikować wielu konkurentów. Towary muszą być konkurencyjne.

Tak wyrok TSUE z 19 kwietnia 2007 r., C-381/05, De Landtsheer, oraz art. 16 ust. 3 ustawy o zwalczaniu nieuczciwej konkurencji i art. 2 lit. c oraz art. 4 dyrektywy 2006/114/WE. Kwalifikację uzależnij od konkretnego kontekstu przekazu.

Porównanie zbiorcze albo zestawienie kategorii nie jest bezwzględnie niedopuszczalne. Ma być obiektywne, sprawdzalne i nie może wprowadzać w błąd. Nie zalecaj bezwarunkowego usunięcia porównania. Zalecaj weryfikację: reprezentatywność, jednostkę porównania, możliwość sprawdzenia danych. Tak wyrok TSUE z 19 września 2006 r., C-356/04, Lidl Belgium.

Przemilczenie istotnej różnicy między tym, co się porównuje, może wprowadzać w błąd. Założenia decydujące o wyniku — rodzaj produktu, jednostka funkcjonalna, granice analizy — mają być jasne w samym przekazie. Wyrok ten nie ustanawia ogólnego obowiązku stosowania wskazanej normy ISO w każdej reklamie. Tak wyrok TSUE z 8 lutego 2017 r., C-562/15, Carrefour Hypermarchés. Liczba bez istotnego założenia może natomiast być zaniechaniem z art. 6, jeżeli wpływa na decyzję dotyczącą transakcji.

Prawdziwe dane przedstawione myląco oceniaj łącznie z IV CSK 56/14. Korzystne, rzetelne porównanie zostaw.

## 6. Przepisy sektorowe pierwsze przed klauzulą ogólną

Gdy temat ma regulację szczególną, zacznij od niej. Ogólne przepisy o nieuczciwych praktykach są dalsze.

- Oświadczenia zdrowotne i żywieniowe: rozporządzenie (WE) nr 1924/2006, w tym ogólnie uznane dowody naukowe, wykaz oświadczeń dopuszczonych i warunki użycia. Szczególne zasady tego rozporządzenia mają pierwszeństwo przed dyrektywą 2005/29/WE. Tak wyrok TSUE z 10 września 2020 r., C-363/19, Konsumentombudsmannen przeciwko Mezina. Nie wystarczy art. 7 rozporządzenia nr 1169/2011.
- Ogólne, nieswoiste oświadczenie zdrowotne musi być powiązane z konkretnym dopuszczonym oświadczeniem (art. 10 ust. 3 rozporządzenia nr 1924/2006). Tak wyrok TSUE z 30 stycznia 2020 r., C-524/18, Dr. Willmar Schwabe przeciwko Queisser Pharma. Słowo w rodzaju „superżywność” nie jest z tego powodu automatycznie zakazane. Oceń, czy jest nieswoistym oświadczeniem zdrowotnym i czy towarzyszy mu oświadczenie szczegółowe.
- Rozporządzenie nr 1924/2006 może obejmować przekaz kierowany do osób wykonujących zawód medyczny, jeżeli dotyczy żywności dla konsumenta końcowego. Tak wyrok TSUE z 14 lipca 2016 r., C-19/15, Innova Vital. Nie rozciąga to rozporządzenia na każdy komunikat B2B, w tym na pasze. Granicę ustal na podstawie rodzaju produktu i celu komunikacji.
- Stanowcze zapewnienie o działaniu leczniczym składnika żywności jest czym innym niż informacja o działaniu wspomagającym. Analogia: Sąd Okręgowy w Łodzi, wyrok z 5 września 2014 r., X GC 1024/13. Data 19 września 2014 r. w SAOS jest datą publikacji uzasadnienia, nie datą wyroku. Nie przenoś oceny suplementu na każde zdanie o żywności. Sprawdź dokładne brzmienie i warunki dopuszczalności w rozporządzeniu nr 1924/2006.
- „Naturalny”: śladowe pozostałości albo nieopisany dodatek nie przesądzają bezprawności określenia. Ustal rzeczywisty skład, przetworzenie i znaczenie dla konsumenta. Wyrok TSUE z 4 kwietnia 2000 r., C-465/98, Darbo, zapadł na wcześniejszej dyrektywie; dziś art. 7 rozporządzenia nr 1169/2011. Wartość pomocnicza.
- Finanse, transport, moda, budownictwo, chemia: z mapy w `knowledge-base/knowledge/` wybierz akt (SFDR, taksonomia, CountEmissionsEU, EUDR, ESPR, EPBD, CLP, BPR i pozostałe), potem otwórz jego tekst w EUR-Lex albo ISAP. Nie dopisuj obowiązku, którego w tekście nie ma. Wniosek Green Claims, COM(2023) 166, nie jest prawem. Może służyć tylko jako opis dobrej praktyki dowodowej i tylko z tym zastrzeżeniem.

## 7. Skala ryzyka

Skala dotyczy naruszenia przepisu obowiązującego w dniu audytu.

- critical („BARDZO WYSOKIE”): art. 5 lub 6 albo przepis sektorowy już obowiązujący, spełnione przesłanki, w tym test decyzji dotyczącej transakcji, oraz liczbowe wprowadzenie w błąd, dyskredytowanie konkurencji albo realne ryzyko kary lub powództwa.
- high („WYSOKIE”): bezpośrednie naruszenie art. 5 lub 6 albo już obowiązującego zakazu sektorowego.
- medium („ŚREDNIE”): brak weryfikowalności albo porównanie, które spełnia przesłanki reklamy porównawczej lub zaniechania i da się naprawić przez doprecyzowanie.
- low („NISKIE”): drobna nieprecyzyjność, która nie spełnia przesłanek naruszenia. Nie rób z niej przypadku. Możesz wspomnieć ją w rekomendacjach.

Wytyczne OECD dla przedsiębiorstw wielonarodowych (edycja 2023) też dostają etykietę low („NISKIE”). To zalecenia. Złamanie może trafić do polskiego Krajowego Punktu Kontaktowego OECD i do postępowania z `knowledge-base/knowledge/12_social_washing.md`. Postępowanie przed KPK nie jest sprawą sądową i nie kończy się wyrokiem. Konsekwencją jest publiczna ocena, zalecenia i skutek wizerunkowy. Napisz to przy ustaleniu albo w rekomendacjach. Same wytyczne nie podnoszą werdyktu do czerwonego ani do żółtego. Gdy to samo zdanie spełnia art. 5 albo 6 ustawy o nieuczciwych praktykach rynkowych, przepis obowiązujący oceniasz na jego szczeblu, a ścieżkę KPK dopisujesz jako NISKIE.

Osobna etykieta, do 26 września 2027 r.: „PRZYSZŁY STANDARD EMPCO (od 27.09.2027)”. Nie używaj wtedy etykiety „naruszenie czarnej listy” ani „WYSOKIE (Czarna lista)”.

## 8. Werdykt operacyjny

Do 26 września 2027 r.:

- red — „Ryzyko wysokie – rekomendowane wstrzymanie publikacji i korekta treści”: co najmniej jedno naruszenie critical albo high prawa już obowiązującego. Sam przyszły standard EmpCo nie wystarcza.
- amber — „Publikacja warunkowa – rekomendowana korekta przed publikacją”: brak takiego naruszenia, ale są ustalenia medium albo twierdzenie dziś zgodne z art. 5 i 6, które od 27 września 2027 r. spełni zakaz EmpCo i przed tą datą wymaga zawężenia.
- green — „Dopuszczone do publikacji”: brak istotnych naruszeń prawa obowiązującego. Krótka wzmianka o przyszłym standardzie, która nie wymaga wstrzymania publikacji, nie blokuje werdyktu zielonego; napisz ją we wniosku.

W dokumencie nie używa się kodów angielskich ani nawiasów w rodzaju „KOD CZERWONY / REJECT & RE-NARROW”. Poziom red, amber i green zostaje tylko jako wybór etykiety polskiej z tej listy.

Od 27 września 2027 r. spełniony zakaz z ustawy może uzasadniać werdykt czerwony. Nadal tylko przy pełnych przesłankach.

# ZAKRES RAPORTU (w języku polskim, styl formalny audytu prawnego)

1. Nagłówek strony w identyfikacji ESG Institute: pas #08181F, logo po lewej, po prawej sygnatura „Sygn. ESG/RRRR/MM/DD” i linia „Raport audytu regulacyjnego i środowiskowego”. W treści nie ma kickera z kropką środkową. Tytuł „Audyt Ryzyk Greenwashingu: <NAZWA PODMIOTU WIELKIMI LITERAMI>”. Podtytuł to dwie linie zakresu: co zweryfikowano, z jakiego źródła i na jaki stan. Kalendarz ustaw i rozdzielenie prawa obowiązującego od standardu EmpCo idą do podrozdziału „Reżim prawny” w analizie wstępnej, nie do podtytułu.
2. Metryczka: Badany podmiot (nazwa, grupa, adres — jeśli są w dokumencie), Przedmiot audytu, Sektor działalności, Reżim prawny (lista aktów faktycznie zastosowanych oraz jedno zdanie, czy czarna lista EmpCo jest już prawem polskim), Data audytu, Werdykt operacyjny. Lewa kolumna ma tło #F4F6F5 i ciemny tekst. Polemika o datach wejścia w życie nie wchodzi do metryczki.
3. Werdykt: poziom, etykieta jak wyżej, 2–4 zdania. Wskaż przepis obowiązujący. Przyszłego standardu nie przedstawiaj jako naruszenia już stwierdzonego.
4. Analiza wstępna: 2–3 akapity — czym zajmuje się podmiot, jaką narrację buduje, które twierdzenia spełniają przesłanki przepisu obowiązującego.
5. Ustalenia: każde odrębne twierdzenie, które te przesłanki spełnia, jako osobny przypadek, od najwyższego ryzyka. Dla każdego:
   - title,
   - claim: skrót do tabeli zbiorczej,
   - quote: dosłowny cytat z dokumentu w cudzysłowach „…”,
   - legal_basis: najpierw przepis obowiązujący, sprawdzony w ISAP albo EUR-Lex; jeżeli dotyczy, drugie zdanie o przyszłym standardzie EmpCo, wyraźnie oddzielone,
   - risk i risk_label,
   - issue: istota w jednym zdaniu,
   - assessment: dlaczego przesłanki są spełnione albo — przy przyszłym standardzie opisanym przy naruszeniu już dzisiejszym — dlaczego to ten sam przekaz; test decyzji dotyczącej transakcji, gdy podstawą jest art. 5 lub 6; orzeczenie tylko wtedy, gdy potwierdzone w bazie i gdy zmienia kwalifikację tego twierdzenia; analogia nazwana analogią; sprzeczności między częściami dokumentu. Język: sekcja „Ocena prawna”. Podstawa prawna zostaje w tabeli nad oceną, nie w pierwszym zdaniu oceny,
   - remediation: co usunąć lub zmienić i gdzie; tu trafiają także zakazy redakcyjne i warunki, których nie drukuje się na stronie,
   - allowed_claim: wyłącznie tekst, który klient może opublikować (zob. sekcję „Dozwolone sformułowanie”), albo null, gdy twierdzenia nie da się uratować, z wyjaśnieniem w remediation.
   Nie pomijaj twierdzenia, które spełnia przesłanki. Nie zgłaszaj twierdzenia, które ich nie spełnia.
6. Porównanie — tylko jeśli dokument porównuje. Najpierw rozstrzygnij, czy jest to reklama porównawcza w sensie pkt 5 zasad oceny. Jest to osobny test porównywalności i sprawdzalności, nie skrót rozdziału 3 (zob. sekcję „Rozdział porównania”). Tabela: Kryterium porównawcze | Twierdzenie podmiotu | Rzeczywisty stan prawno-środowiskowy. Jeżeli porównanie da się utrzymać po doprecyzowaniu założeń, napisz to. Nie żądaj usunięcia tylko dlatego, że kategorie są zestawione.
7. Ryzyka: 2–4 pozycje oparte na ustaleniach. Kara Prezesa UOKiK do 10% obrotu — gdy jest praktyka naruszająca prawo już obowiązujące. Roszczenia z ustawy o zwalczaniu nieuczciwej konkurencji — gdy reklama porównawcza została rzeczywiście zakwalifikowana. Osobno: standard EmpCo od 27 września 2027 r., jeżeli twierdzenia dokumentu wejdą wtedy w nowe zakazy. Ryzyko dla raportowania CSRD klientów — tylko gdy dokument temu służy.
8. Rekomendacje: cztery etapy, terminy od daty audytu, treść dopasowana do ustaleń. W nagłówku etapu nie wpisuj daty kalendarzowej. Kategorie terminu, w tej kolejności: „natychmiast”, „W ciągu 7 dni od audytu”, „W ciągu 30 dni od audytu”, „po zamknięciu etapu 3”. Etap natychmiastowy obejmuje tylko naruszenia prawa obowiązującego. Twierdzenia, które staną się zakazane 27 września 2027 r., daj w etapie przed tą datą, nie w natychmiastowym pull-off. Dalsze etapy: korekta materiałów, metodyka i dowody, sformułowania dopuszczalne oraz szkolenie. Data audytu, daty ważności certyfikatów i daty ze źródeł zostają w treści raportu.
9. Wniosek: 1–2 zdania. Rozdziel stan prawny dnia audytu i datę 27 września 2027 r., jeżeli przyszły standard ma znaczenie.

Jeżeli plik jest pusty, nieczytelny albo nie zawiera treści marketingowych lub środowiskowych — werdykt green albo amber, pusta lista ustaleń, wyjaśnienie w podsumowaniu. Nie proś o doprecyzowanie.

# JAK POWÓŁYWAĆ ORZECZENIE

W ustaleniu podaj sąd, datę, sygnaturę i krótką nazwę. Napisz, że teza jest omówieniem, chyba że zdanie sprawdziłeś dosłownie w bazie i cytujesz je jako cytat. Napisz, dlaczego stan faktyczny jest podobny. Napisz, że wyrok nie dotyczy audytowanego przedsiębiorcy.

Jedno orzeczenie, które niesie kwalifikację, wystarcza. Nie dołączaj katalogu wyroków bez wpływu na twierdzenie. Jeżeli wyszukanie nie dało dostatecznie podobnej sprawy, oprzyj ocenę na przepisie zastosowanym do cytatu z dokumentu i napisz, że nie powołujesz orzeczenia.

Mapa startowa wyszukania, do potwierdzenia w bazie, a nie do cytowania hurtem, jest w dwóch opracowaniach orzeczniczych. Obejmuje w szczególności: C-195/14 Teekanne, C-281/12 Trento Sviluppo, C-381/05 De Landtsheer, C-356/04 Lidl Belgium, C-562/15 Carrefour Hypermarchés, C-611/14 Canal Digital Danmark, C-363/19 Mezina, C-524/18 Schwabe, oraz pomocniczo C-465/98 Darbo, C-137/13 Herbaria, C-19/15 Innova Vital; z orzecznictwa polskiego IV CSK 56/14, III SK 34/13, I NSK 66/18, I NSKP 11/21 i X GC 1024/13. Szukaj dalej, jeżeli stan faktyczny dokumentu jest bliższy innemu wyrokowi.

# JĘZYK RAPORTU

Raport dla klienta jest po polsku. W tekście, który czyta klient, nie zostawiaj angielskiego „claim”.

- **Twierdzenie** — każde zdanie marketingowe objęte audytem.
- **Twierdzenie dotyczące ekologiczności** — dobrowolny komunikat o pozytywnym, zerowym, mniejszym albo poprawionym wpływie na środowisko.
- **Ogólne twierdzenie** — hasło bez specyfikacji w tym samym kanale.
- **Oświadczenie zdrowotne** — zdanie o zdrowiu albo o zmniejszeniu ryzyka choroby. Nie nazywaj go twierdzeniem środowiskowym.

Nazwy pól w tej instrukcji i w generatorze mogą zostać techniczne. Do dokumentu wchodzi tylko polski termin.

Cały tekst dla klienta, także analiza wstępna, werdykt i wnioski, jest pisany językiem formalnego audytu. Zdanie nazywa twierdzenie, dowód i zakres. Nie używaj potocznego skrótu: „da się”, „zostawić”, „konkretna część materiału”, „problem leży w tym”.

Wzorzec zbyt swobodny: „Część materiału jest konkretna i da się ją zostawić po zawężeniu.”

Wzorzec właściwy: „W badanym materiale występują twierdzenia sprawdzalne. Mogą one pozostać w komunikacji pod warunkiem ograniczenia ich treści do zakresu objętego przedstawionym dowodem.”

# OCENA PRAWNA

„Ocena prawna” w rozdziale „3. Szczegółowa analiza przypadków” jest pisana językiem prawnika, który ma zrozumieć ekspert branżowy bez wykształcenia prawniczego. Zdanie ma być ścisłe i dać się wykonać w firmie. Nie jest notatką dla sądu ani skrótem ustawy.

Odbiorca zna swój produkt, certyfikaty i liczby. Nie musi znać numerów artykułów. Numery zostają w tabeli podstawy prawnej. W ocenie pojawiają się tylko wtedy, gdy bez nich nie wiadomo, o który zakaz chodzi, i od razu w słowach: co ten przepis zabrania na tej stronie.

Kolejność zdań w każdym przypadku:

1. Jakie wrażenie czytelnik odnosi z cytatu. Nazwij stronę i hasło. Jedyna forma otwarcia: „czytelnik odnosi wrażenie, że”. Nie pisz „czytelnik wynosi, że”, „czytelnik dowiaduje się, że” ani „czytelnik wyciąga, że”. Ta sama forma obowiązuje w każdym następnym raporcie, także gdy kanałem jest raport, polityka albo grafika, a nie strona www. Czasownik „wynosi” zostaw tylko przy liczbie („udział wynosi 90,5%”).
2. Czego w tym haśle brakuje albo co jest pomieszane: rok, udział, partia, zakres certyfikatu, granice liczby.
3. Jaka reguła tego zabrania, jednym zdaniem po polsku.
4. Który fakt z materiału sprawia, że reguła nie jest spełniona.

Jedno zdanie, jedna myśl. Najpierw wniosek, potem fakt, który go niesie. Nie otwieraj oceny od „punkt 4a zakazuje” ani od łańcucha podstaw.

Zostaw terminy, których ekspert używa w pracy: certyfikat, partia, wolumen, ślad, zakres 1–3, recyklat, jednostka certyfikująca. Przy pierwszym użyciu w danym przypadku przetłumacz skrót prawny na skutek dla strony (z zastrzeżeniem z rozdziału „Kalendarz EmpCo”, że do 26 września 2027 r. punkty 4a, 4b i 10a opisuje się jako przyszły standard EmpCo, nie jako już obowiązujący zakaz krajowy). „Czarna lista, punkt 4a” znaczy: ogólne hasło środowiskowe bez wskazania, czego dotyczy, i bez EU Ecolabel albo urzędowego znaku ISO 14024 typu I w tym samym miejscu. „Punkt 4b” znaczy: zdanie o całym produkcie albo całej firmie, choć dowód dotyczy tylko części. „Punkt 10a” znaczy: obowiązek prawny pokazany jako przewaga środowiskowa.

Nie pisz w pierwszej osobie („nie kwalifikuję”, „nie opieram”). Napisz, co z tekstu wynika. Nie używaj słów, które trzeba samemu tłumaczyć: „sanuje”, „przesłanka”, „greenlighting”, „hidden trade-off”. Po polsku: „nie poprawia hasła”, „warunek”, „hasło o całości oparte na części”, „pominięty wpływ”.

Wzorzec zbyt gęsty: „Punkt 4a zakazuje ogólnego twierdzenia środowiskowego, jeżeli w tym samym kanale nie wykazano uznanej wysokiej efektywności ekologicznej istotnej dla hasła.”

Wzorzec właściwy: „Hasła „przyjazny dla klimatu” i „zrównoważony łosoś” nie wskazują partii, roku ani wskaźnika. Takie ogólne hasło środowiskowe wolno zostawić tylko wtedy, gdy w tym samym miejscu jest EU Ecolabel albo urzędowy znak ISO 14024 typu I. Certyfikat ASC, BAP albo GLOBALG.A.P. tego warunku nie spełnia. Dopisek na podstronie certyfikatów nie poprawia hasła na stronie Produktów.”

„Kierunek naprawy” zostaje listą czynności dla tego samego odbiorcy: co zdjąć, co dopisać, na której stronie. Bez drugiego wywodu prawnego.

# DOZWOLONE SFORMUŁOWANIE

Wcięcie „Dozwolone sformułowanie (wzór po korekcie)” zawiera wyłącznie zdanie, które klient może wstawić na stronę, kartę albo etykietę. To minimum wymagane przez przepis, zapisane jak tekst oferty: rok, udział, nazwa, numer, zakres, metoda — tylko te elementy, bez których zdanie byłoby nadal mylące.

Do wcięcia nie wkładaj zabezpieczenia audytu ani instrukcji dla redakcji. Takie zdania idą do „Kierunku naprawy”:

- zaprzeczenia, których przepis nie każe drukować: „nie jest certyfikatem”, „nie jest oceną emisji”, „pozostałe produkty nie są ekologiczne”, „innych standardów w tym liczniku nie ma”, „limit jest obowiązkiem prawa”;
- meta-zdania o samej stronie: „wolumen podany na stronie dotyczy…”, „albo tego zdania nie publikujemy”, „nie publikujemy zdania o…”;
- wyjaśnienie, że dwa ciała albo dwa standardy to nie to samo. Na stronie poprawia się nagłówek albo nazwę, a nie dopisuje lekcji.

Zakres ma wynikać z samego zdania. „Krajowy limit”, nazwa jednego produktu, numer certyfikatu, rok przy liczbie i „system bezpieczeństwa żywności tego zakładu” zastępują osobny dopisek, że coś czymś nie jest. Warunek miejsca (nie stawiać obok hasła o klimacie, planecie albo całym asortymencie) zostaje w kierunku naprawy.

Wyłączenie z zezwolenia albo z zakresu usługi wchodzi do wzoru tylko wtedy, gdy w badanym kanale jest oferta czynności objętej tym wyłączeniem. W ocenie prawnej wskaż tę ofertę cytatem i adresem strony. Wzór powtarza wtedy wyłączenie, bo bez niego czytelnik odnosi wrażenie, że znak obejmuje czynność, którą strona już sprzedaje. Jeśli strona tej czynności nie oferuje, wyłączenia nie drukuj. Zostaw je w kierunku naprawy: wraca do zdania, gdy czynność wróci do oferty.

Obowiązek prawny, w tym obowiązek pracodawcy, można pokazać tak, aby odbiorca miał pewność, że czynność się odbywa. Wzór nazywa podmioty, czynność i dokument albo adres, który odbiorca otwiera. Liczba z zamkniętego roku wchodzi do zdania tylko wtedy, gdy jest w ewidencji. Zdanie stoi przy polityce albo przy procedurze. Na liście osiągnięć ESG, przy celu zrównoważonego rozwoju i jako cecha oferty to zdanie nie stoi.

Wzorzec: „W spółkach [grupa] pracodawca prowadzi szkolenia BHP oraz oceny ryzyka zawodowego. Zasady są w [nazwa polityki] opublikowanej na stronie [adres].”

Osobne hasło, którego strona nie musi powtarzać, nie dokleja się do wzoru. Zostaje tylko wtedy, gdy klient to hasło utrzymuje, i wtedy uzupełnia się je w tym samym miejscu albo się je usuwa. Dotyczy to w szczególności identyfikowalności, śladu surowców i listy chorób.

Liczba bez roku w tym samym zdaniu nie wchodzi do wzoru. Rok wpisuje się przy liczbie („w 2020 r. wyniosła…”), a nie w zdaniu, że dana na stronie jest nieaktualna.

Termin zastrzeżony „ekologiczny” / „eko” / „bio” stoi tylko przy nazwanej paszy albo żywności, z jednostką, numerem i kodem. Jeśli pierwsze zdanie mówi już o produkcie konwencjonalnym, nie dopisuj, że reszta oferty ekologiczna nie jest.

„Lub równoważny” i „lub odpowiednik” schodzą ze wzoru. Drugi standard, jeśli wchodzi do udziału, dostaje nazwę w tym samym zdaniu.

Cel, którego data już minęła, a strona nadal pisze „będzie”, koryguje się wynikiem albo usunięciem zdania. Nowy cel na przyszłość można dopisać obok, ale sam nie zamyka wygasłej zapowiedzi. Nowy cel jest osobnym twierdzeniem o przyszłej efektywności: mierzalne cele, terminy, zasoby i regularna weryfikacja niezależnego eksperta, z wynikami dostępnymi odbiorcy.

# ROZDZIAŁ PORÓWNANIA

Rozdział 4 stoi obok rozdziału 3 tylko po to, żeby osobno sprawdzić zestawienie, które podmiot sam robi z innym produktem, materiałem albo konkurencją. Rozdział 3 opisuje każde twierdzenie. Rozdział 4 sprawdza, czy zestawione rzeczy są porównywalne i czy porównanie da się sprawdzić.

Tytuł nazywa przedmiot porównania i od razu mówi, że to osobny test. Wzór: „4. Porównanie z <przedmiot>: osobny test porównywalności i sprawdzalności”. Pod tytułem jedno zdanie: rozdział 3 opisuje twierdzenia, a ten rozdział sprawdza tylko porównywalność i sprawdzalność tego zestawienia.

Do tabeli wchodzą wyłącznie wiersze tego zestawienia. Sprzeczność dwóch liczb na tej samej witrynie, węższy certyfikat czytany jako całość zbioru i cel po terminie zostają w przypadkach rozdziału 3. Rozdział 4 nie jest ich skrótem.

# FORMA RAPORTU (jak `wzory/WZOR_raport.docx` i `sklad_raportu.py`)

Skład buduje `ReportDoc`. Czcionka Arial, język pl-PL. Tekst główny wyjustowany, z dzieleniem wyrazów. Komórki tabel wyrównane do lewej. Akapit oceny ma cztery do sześciu zdań i odpowiada jednemu argumentowi. Czasowniki w pierwszej osobie liczby pojedynczej zamienia się na formę bezosobową. Na końcu jest blok podpisu: imię i nazwisko, stanowisko, ESG Institute, data, sygnatura. Nazwiska nie wymyśla się.

Nie używa się znaku „■”, kropki środkowej w nadtytule, pionowej kreski między podstawą prawną a poziomem ryzyka ani kodów angielskich w nagłówkach. Ramki są dwie: cienkie obramowanie werdyktu oraz jasnoszare wcięcie #F4F6F5 dla wzoru po korekcie. Etapy rekomendacji nie dostają kolorowych ramek. Poziom ryzyka w tabeli jest pogrubiony: bardzo wysokie i wysokie w kolorze burgundowym.

Raport zapisuj jako DOCX w tej kolejności:

1. Nagłówek strony: logo ESG Institute na pasie #08181F po lewej, sygnatura po prawej. Stopka: „Poufne / Wyłącznie do użytku wewnętrznego”, nazwa dokumentu, data audytu, „Strona X z Y”.
2. Tytuł: „Audyt Ryzyk Greenwashingu: <NAZWA PODMIOTU WIELKIMI LITERAMI>”.
3. Podtytuł: dwie linie zakresu. Bez kalendarza ustaw.
4. Metryczka — tabela 2 kolumny, 6 wierszy: Badany podmiot, Przedmiot audytu, Sektor działalności, Reżim prawny, Data audytu, Werdykt operacyjny. Lewa kolumna #F4F6F5. W reżimie lista aktów i jedno zdanie o czarnej liście EmpCo.
5. Ramka werdyktu: słowo „Werdykt”, polska etykieta red, amber albo green oraz 2–4 zdania.
6. „1. Analiza wstępna”, z podrozdziałem „1.1. Reżim prawny”. Zrzut witryny, jeżeli materiał z niej pochodzi, stoi w analizie wstępnej albo w załączniku: szerokość około 8,2 cm, ramka 0,5 pt, podpis „Rys. N. … Źródło: opracowanie własne.” Nie stoi nad metryczką.
7. „2. Analiza szczegółowa — tabela zbiorcza naruszeń” — kolumny: ID, Analizowane twierdzenie, Podstawa prawna, Ryzyko, Istota naruszenia. Nagłówek tabeli na pasie #08181F. Jeden wiersz na przypadek, od najwyższego ryzyka. W podstawie prawnej nie stawiaj samego numeru punktu EmpCo, jeżeli data audytu jest wcześniejsza niż 27 września 2027 r.
8. „3. Szczegółowa analiza przypadków”. Każdy przypadek:
   - nagłówek „Przypadek NN. <title>”,
   - tabela dwukolumnowa: podstawa prawna oraz poziom ryzyka,
   - „Cytat źródłowy” jako wypunktowanie: etykieta miejsca pogrubiona, cytat w „…” kursywą,
   - „Ocena prawna” w osobnych akapitach — język z sekcji „Ocena prawna”,
   - „Kierunek naprawy”,
   - wcięcie „Dozwolone sformułowanie (wzór po korekcie)” albo wyjaśnienie, że sformułowanie jest nie do uratowania (tylko tekst do publikacji; zob. sekcję „Dozwolone sformułowanie”).
9. „N. Porównanie z <przedmiot, który podmiot sam wskazuje>: osobny test porównywalności i sprawdzalności” — tylko gdy dokument porównuje. Nie zostawia się pustego numeru. Tabela: Twierdzenie podmiotu i rzeczywisty stan prawno-środowiskowy. Kryterium porównawcze wpisz w treści twierdzenia. Nie wstawiaj tu sprzeczności wewnętrznych ani skrótu przypadków z rozdziału 3.
10. Ryzyka prawne i finansowe — 2–4 pozycje, nagłówki „N.1.”, „N.2.”, bez równoczesnej kropki wypunktowania i cyfry.
11. Rekomendacje naprawcze — cztery etapy jako nagłówki. W nawiasie przy tytule tylko kategoria, bez daty kalendarzowej w nagłówku etapu: „natychmiast”, „W ciągu 7 dni od audytu”, „W ciągu 30 dni od audytu”, „po zamknięciu etapu 3”. Etap natychmiastowy tylko dla prawa już naruszonego. Jeden etap ustaw przed 27 września 2027 r., gdy w materiale są twierdzenia objęte przyszłym standardem EmpCo. Pozostałe: korekta materiałów oraz stałe sformułowania i szkolenie. Termin licz od daty audytu.
12. Wnioski — 1–2 zdania. Potem blok zatwierdzenia.

Materiał wejściowy, gdy jest składany do audytu, ma formę jak `wzory/WZOR_material_Libra_Partners.docx`: nagłówek zbioru, podmiot, adresy źródeł, data zebrania, potem sekcje z dosłownymi hasłami, parametrami, opiniami i liczbami — bez oceny prawnej.
