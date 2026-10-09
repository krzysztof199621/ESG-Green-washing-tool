"""Skaner greenwashingu na podstawie knowledge-base/knowledge/rules.json."""

from __future__ import annotations

import json
import re
import zipfile
from dataclasses import asdict, dataclass
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RULES_PATH = ROOT / "knowledge-base" / "knowledge" / "rules.json"

RE_FLAGS = re.IGNORECASE | re.DOTALL


@dataclass
class Finding:
    rule_id: str
    code: str
    title: str
    severity: str  # high | medium | low
    message: str
    snippet: str
    advice: str
    source: str


def load_rules() -> dict:
    return json.loads(RULES_PATH.read_text(encoding="utf-8"))


def _collapse(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


class _HTMLText(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []
        self.skip = False

    def handle_starttag(self, tag, attrs):
        if tag in {"script", "style", "noscript"}:
            self.skip = True
        if tag in {"p", "div", "li", "h1", "h2", "h3", "br", "tr"}:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in {"script", "style", "noscript"}:
            self.skip = False

    def handle_data(self, data):
        if not self.skip and data.strip():
            self.parts.append(data)


def extract_html(raw: bytes) -> str:
    parser = _HTMLText()
    parser.feed(raw.decode("utf-8", errors="replace"))
    return _collapse("".join(parser.parts))


def extract_docx(raw: bytes) -> str:
    with zipfile.ZipFile(BytesIO(raw)) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = re.sub(r"<[^>]+>", "", xml)
    xml = xml.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return _collapse(xml)


def extract_pdf(raw: bytes) -> str:
    import pymupdf

    doc = pymupdf.open(stream=raw, filetype="pdf")
    parts = [page.get_text() for page in doc]
    return _collapse("\n\n".join(parts))


def extract_pptx(raw: bytes) -> str:
    texts: list[str] = []
    with zipfile.ZipFile(BytesIO(raw)) as zf:
        names = sorted(n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
        for name in names:
            xml = zf.read(name).decode("utf-8", errors="replace")
            bits = re.findall(r"<a:t[^>]*>([^<]*)</a:t>", xml)
            texts.append(" ".join(bits))
    return _collapse("\n".join(texts))


def extract_file(filename: str, raw: bytes) -> str:
    name = filename.lower()
    if name.endswith(".pdf"):
        return extract_pdf(raw)
    if name.endswith(".docx"):
        return extract_docx(raw)
    if name.endswith(".pptx"):
        return extract_pptx(raw)
    if name.endswith((".html", ".htm")):
        return extract_html(raw)
    if name.endswith((".txt", ".md", ".csv")):
        return _collapse(raw.decode("utf-8", errors="replace"))
    # sniff
    if raw[:4] == b"%PDF":
        return extract_pdf(raw)
    if raw[:2] == b"PK":
        try:
            return extract_docx(raw)
        except Exception:
            return extract_pptx(raw)
    return _collapse(raw.decode("utf-8", errors="replace"))


def _snip(text: str, start: int, end: int, radius: int = 110) -> str:
    a = max(0, start - radius)
    b = min(len(text), end + radius)
    chunk = re.sub(r"\s+", " ", text[a:b]).strip()
    if a:
        chunk = "…" + chunk
    if b < len(text):
        chunk = chunk + "…"
    return chunk


def _window(text: str, start: int, end: int, radius: int = 220) -> str:
    a = max(0, start - radius)
    b = min(len(text), end + radius)
    return text[a:b]


def _search(pattern: str, text: str):
    return list(re.finditer(pattern, text, RE_FLAGS))


SPECIFIERS = re.compile(
    r"("
    r"eu ecolabel|ecolabel|iso\s*14024|iso\s*14083|typ(?:u)?\s*i\b|"
    r"certyfikat(?:u|em)?\s+nr|"
    r"2018/848|1169/2011|"
    r"\bepd\b|en\s*15804|"
    r"dotyczy (?:wyłącznie )?(?:opakowani|tego modelu|tej linii|warstwy)|"
    r"rpet|recyklingu mechanicznego|"
    r"akredytowan|niezależn\w+\s+stron|"
    r"countemissions"
    r")",
    RE_FLAGS,
)

LEGAL_DOC = re.compile(
    r"(dz\.\s*u\.|dyrektyw[ay]|rozporządzeni|ustaw[ay]\s+z\s+dnia|art\.\s*\d+|załącznik\s+i)",
    RE_FLAGS,
)


def _looks_like_legal_act(text: str) -> bool:
    hits = len(LEGAL_DOC.findall(text[:8000]))
    return hits >= 8


GENERIC_PATTERNS = [
    (r"\bekologiczn\w*", "ekologiczny"),
    (r"\beko\b", "eko"),
    (r"przyjazn\w{0,8}\s+(?:dla\s+)?(?:środowiska|planety)", "przyjazny środowisku"),
    (r"zrównoważon\w{0,6}(?:\s+(?:rozwój|kolekcj\w*|produkt\w*|firm\w*|mod\w*))?", "zrównoważony"),
    (r"climate[\s-]?friendly", "climate friendly"),
    (r"eco[\s-]?friendly", "eco-friendly"),
    (r"join[\s-]?life", "Join Life"),
    (r"\bconscious\b", "Conscious"),
    (r"zielon[aeyą]\s+(?:dostaw\w*|flot\w*|energi\w*|produkt\w*|pacz\w*|opcj\w*)", "zielona [usługa/produkt]"),
    (r"zero[\s-]?waste", "zero waste"),
    (r"net[\s-]?zero", "net-zero"),
    (r"climate[\s-]?neutral|węglowo[\s-]?neutraln\w*|neutraln\w+\s+(?:klimatycznie|dla środowiska)", "neutralność klimatyczna"),
    (r"climate[\s-]?positive|climate[\s-]?positiv", "climate positive"),
    (r"zeroemisyjn\w*|bezemisyjn\w*", "zero-/bezemisyjny"),
    (r"100\s*%\s*(?:oze|energii odnawialnej|odnawialn)", "100% OZE"),
]


def analyze_text(text: str, filename: str = "") -> dict:
    rules = load_rules()
    text = _collapse(text)
    findings: list[Finding] = []
    seen: set[tuple[str, str]] = set()

    def add(f: Finding) -> None:
        key = (f.rule_id, f.snippet[:80])
        if key in seen:
            return
        seen.add(key)
        findings.append(f)

    if not text:
        return _pack(filename, text, findings, rules, extra={"note": "Brak tekstu do analizy."})

    legalish = _looks_like_legal_act(text)

    # --- 4a generic claims ---
    for pat, label in GENERIC_PATTERNS:
        for m in _search(pat, text):
            win = _window(text, m.start(), m.end())
            specified = bool(SPECIFIERS.search(win))
            if specified:
                continue
            # skip isolated "eko" inside long legal definitions
            if legalish and label in {"eko", "ekologiczny"}:
                continue
            add(
                Finding(
                    rule_id="ECGT-4a",
                    code="4a",
                    title="Ogólne hasło ekologiczne",
                    severity="high",
                    message=f"Wykryto ogólne hasło „{label}” bez specyfikacji w sąsiedztwie (%, zakres, Ecolabel, ISO 14024, certyfikat).",
                    snippet=_snip(text, m.start(), m.end()),
                    advice="Zamień na twierdzenie skonkretyzowane w tym samym kanale albo usuń hasło. FAQ Komisji: brak miejsca na opakowaniu = nie stawiaj claimu.",
                    source="ECGT zał. I pkt 4a · FAQ ECGT 2026 · UPNPR art. 5",
                )
            )

    # --- 4c offsets ---
    offset_pat = r"(offset\w*|kompensac\w*|kredyt\w*\s+węgl\w*|carbon credit\w*|sadzim[yę]\s+drzew\w*|nasadzen\w*|eac\b|gwarancj\w+ pochodzenia)"
    for m in _search(
        r"(neutraln\w+\s+(?:klimatycznie|dla środowiska)|climate[\s-]?neutral|net[\s-]?zero|climate[\s-]?positive|węglowo[\s-]?neutraln\w*|zerow[aą]\s+emisj\w*)",
        text,
    ):
        win = _window(text, m.start(), m.end(), 320)
        if re.search(offset_pat, win, RE_FLAGS):
            add(
                Finding(
                    rule_id="ECGT-4c",
                    code="4c",
                    title="Neutralność oparta na offsecie",
                    severity="high",
                    message="Claim GHG o produkcie/usłudze wygląda na uzasadniony kompensacją (offset, drzewa, EAC/GO), a nie redukcją w łańcuchu wartości.",
                    snippet=_snip(text, m.start(), m.end(), 140),
                    advice="Zakaz bez wyjątków. Offset można opisać osobno, nigdy jako „produkt/przesyłka/firma neutralna”.",
                    source="ECGT 4c · FAQ Komisji · UOKiK (Allegro drzewa, Bolt EAC)",
                )
            )

    # EAC / 100% OZE without explanation
    for m in _search(r"100\s*%\s*(?:oze|energii odnawialnej|odnawialnej energii|renewable)", text):
        win = _window(text, m.start(), m.end(), 280)
        if re.search(r"\b(eac|go\b|gwarancj\w+ pochodzenia|certyfikat\w* atrybut|energy attribute)", win, RE_FLAGS) or not re.search(
            r"(ppa|on[\s-]?site|własn\w+\s+instalacj|fizyczn\w+\s+dostaw)", win, RE_FLAGS
        ):
            if re.search(r"\b(eac|go\b|gwarancj\w+ pochodzenia)", win, RE_FLAGS) or "100%" in m.group(0):
                # only flag 100% OZE if certificates mentioned OR no physical source at all
                if re.search(r"\b(eac|go\b|gwarancj\w+ pochodzenia)", win, RE_FLAGS) or not re.search(
                    r"(ppa|fotowolta|własn\w+ źródł|on[\s-]?site)", win, RE_FLAGS
                ):
                    add(
                        Finding(
                            rule_id="EAC-NOT-PHYSICAL-ENERGY",
                            code="UOKiK",
                            title="100% OZE bez jasnego źródła energii",
                            severity="high",
                            message="Hasło „100% energii odnawialnej” bez wskazania, czy to energia fizyczna (PPA/on-site), czy certyfikaty atrybutu (EAC/GO).",
                            snippet=_snip(text, m.start(), m.end(), 130),
                            advice="Ujawnij mechanizm. UOKiK vs Bolt: EAC bez wyjaśnienia wprowadza w błąd.",
                            source="UOKiK Bolt 19.01.2026 · UPNPR art. 5–6",
                        )
                    )

    # --- 4b greenlighting / whole vs aspect ---
    whole_pats = [
        r"zrównoważon\w+\s+kolekcj\w*",
        r"cał[aą]\s+(?:firm\w*|działalność|kolekcj\w*|flot\w*)",
        r"(?:nasza|nasz)\s+firma\s+(?:jest|to)\s+(?:eko|zielon|zeroemisyj|zrównoważ)",
        r"zielon[aą]\s+dostaw\w*",
        r"zeroemisyjn\w+\s+(?:e[\s-]?commerce|flot\w*|firm\w*)",
        r"dowozimy zeroemisyjn",
        r"zero waste we własnych obiektach",
        r"join[\s-]?life",
    ]
    aspect_hint = re.compile(
        r"("
        r"\d[\d,\.]{0,3}\s*%|"
        r"ostatni(?:ej)? mili|last[\s-]?mile|"
        r"jeden magazyn|1 z \d+|fragment|"
        r"tylko(?:\s+na)?\s+(?:opakowani|etykiet|last)"
        r")",
        RE_FLAGS,
    )
    for pat in whole_pats:
        for m in _search(pat, text):
            win = _window(text, m.start(), m.end(), 300)
            aspect = aspect_hint.search(win)
            add(
                Finding(
                    rule_id="ECGT-4b",
                    code="4b",
                    title="Claim o całości przy węższym zakresie",
                    severity="high" if aspect else "medium",
                    message=(
                        "Hasło dotyczy całości (kolekcja, firma, dostawa, flota), a w sąsiedztwie widać tylko aspekt"
                        if aspect
                        else "Hasło dotyczy całości produktu lub działalności. UOKiK wymaga podania skali i zakresu — inaczej 4b / art. 5 UPNPR."
                    ),
                    snippet=_snip(text, m.start(), m.end(), 130),
                    advice="Zawęź język 1:1 do dowodu: % masy modelu, odcinek trasy, liczba sklepów. Nie nazywaj całości „zrównoważoną”.",
                    source="ECGT 4b · UOKiK Zara/InPost/DPD/Tchibo",
                )
            )

    # small recycled % sold as eco
    for m in _search(r"(eko|ekologiczn\w*|zrównoważon\w*).{0,80}?(\d{1,2})\s*%\s*(?:recykl|recycl|z recyklingu)", text):
        pct = int(m.group(2))
        if pct <= 30:
            add(
                Finding(
                    rule_id="ECGT-4b",
                    code="4b",
                    title="Niski udział recyklatu jako „eko”",
                    severity="high",
                    message=f"Oznaczenie ekologiczne przy udziale recyklatu {pct}% — wzorzec UOKiK vs Tchibo (próg 10% jako „eko”).",
                    snippet=_snip(text, m.start(), m.end()),
                    advice="Podaj % i której części produktu dotyczy. Nie używaj znacznika „eko” / listka przy trywialnym progu.",
                    source="UOKiK Tchibo 19.01.2026 · ECGT 4a/4b",
                )
            )

    # --- 2a private labels ---
    for m in _search(
        r"(nasz(?:e|a)?\s+(?:eko[\s-]?logo|certyfikat)|eco[\s-]?certified|self[\s-]?declared|"
        r"własn\w+\s+(?:znak|etykiet\w+|certyfikat)|znak zrównoważoności)",
        text,
    ):
        win = _window(text, m.start(), m.end())
        if not re.search(r"(stron[ay] trzec|akredytowan|organ publiczn|eu ecolabel|iso\s*14024)", win, RE_FLAGS):
            add(
                Finding(
                    rule_id="ECGT-2a",
                    code="2a",
                    title="Znak bez systemu certyfikacji",
                    severity="high",
                    message="Własny znak / „eco certified” bez niezależnego systemu strony trzeciej ani organu publicznego.",
                    snippet=_snip(text, m.start(), m.end()),
                    advice="Usuń znak albo oprzyj go na systemie spełniającym definicję ECGT (weryfikacja 3rd party, kryteria publiczne).",
                    source="ECGT 2a · UPNPR art. 7 pkt 2",
                )
            )

    # --- 10a legal requirement as benefit ---
    for m in _search(
        r"(bez substancji zakazanych(?:\s+w\s+UE)?|zgodn\w+\s+z\s+obowiązującym prawem|"
        r"spełnia wszystkie wymogi prawa(?:\s+UE)?|woln\w+\s+od\s+substancji\s+zakazanych)",
        text,
    ):
        add(
            Finding(
                rule_id="ECGT-10a",
                code="10a",
                title="Obowiązek prawny jako przewaga",
                severity="medium",
                message="Cecha wymagana prawem wygląda na sprzedawaną jako wyróżnik marketingowy.",
                snippet=_snip(text, m.start(), m.end()),
                advice="Nie prezentuj wymogu prawnego jako korzyści. To czarna lista 10a.",
                source="ECGT 10a · UPNPR analogia",
            )
        )

    # QR / www as only specification
    for m in _search(r"(szczegóły na (?:stronie|www)|więcej na|zeskanuj kod qr|qr kod)", text):
        win = _window(text, m.start(), m.end(), 200)
        if re.search(r"\b(eko|zielon|zrównoważ|climate|eco-friendly)\w*", win, RE_FLAGS):
            add(
                Finding(
                    rule_id="SAME-MEDIUM-SPEC",
                    code="FAQ",
                    title="Specyfikacja poza kanałem hasła",
                    severity="medium",
                    message="Odsyłacz do www/QR przy haśle ekologicznym. FAQ Komisji: kwalifikacja musi być w tym samym kanale.",
                    snippet=_snip(text, m.start(), m.end()),
                    advice="Albo dopisz zakres obok hasła, albo zdejmij ogólne hasło z opakowania/reklamy.",
                    source="FAQ ECGT 2026 · UCPD Guidance § 4.1.1",
                )
            )

    # durability / planned obsolescence weak signal
    for m in _search(r"(trwa na lata|niezniszczaln\w+|gwarantowan\w+\s+żywotność(?!\s+\d))", text):
        add(
            Finding(
                rule_id="ECGT-23f",
                code="23f–j",
                title="Fałszywa trwałość (sygnał)",
                severity="low",
                message="Ogólne zapewnienie o trwałości bez okresu, testu lub warunków. Sprawdź, czy nie wpada w 23f–j.",
                snippet=_snip(text, m.start(), m.end()),
                advice="Podaj mierzalny okres, normę i ograniczenia. Ukryta wadliwość konstrukcyjna = zakaz.",
                source="ECGT 23f–j",
            )
        )

    # --- sector filters ---
    sectors = _detect_sectors(text)
    if "fmcg" in sectors:
        for m in _search(r"\b(eko|bio|organic|ekologiczn\w*)\b", text):
            win = _window(text, m.start(), m.end(), 200)
            foodish = re.search(r"(żywnoś|mleko|jogurt|kawa|kakao|olej|mięso|warzyw|owoc|pieczyw|wino)", text, RE_FLAGS)
            if foodish and not re.search(r"(2018/848|certyfikat|jednostk\w+ certyfik|kod pl[\s-]?eko)", win, RE_FLAGS):
                add(
                    Finding(
                        rule_id="ORG-848-30",
                        code="848",
                        title="Eko/bio na żywności bez certyfikatu",
                        severity="high",
                        message="Terminy eko/bio/organic na żywności są zastrzeżone. Potrzebny certyfikat 2018/848 i jednostka upoważniona.",
                        snippet=_snip(text, m.start(), m.end()),
                        advice="Bez certyfikatu — usuń eko/bio/organic i liść sugerujący produkcję ekologiczną.",
                        source="rozp. 2018/848 art. 30 · Dz.U. 2022 poz. 1370",
                    )
                )
        if re.search(r"(zero deforestation|bez wylesiania|rainforest friendly)", text, RE_FLAGS) and not re.search(
            r"(due diligence|dds|geolokaliz)", text, RE_FLAGS
        ):
            m = _search(r"(zero deforestation|bez wylesiania|rainforest friendly)", text)[0]
            add(
                Finding(
                    rule_id="EUDR",
                    code="EUDR",
                    title="Claim o wylesianiu bez due diligence",
                    severity="medium",
                    message="Twierdzenie deforestation-free bez śladu systemu due diligence (EUDR).",
                    snippet=_snip(text, m.start(), m.end()),
                    advice="Albo masz DDS i geolokalizację, albo nie używasz tego hasła.",
                    source="rozp. (UE) 2023/1115",
                )
            )

    if "finance" in sectors:
        if re.search(r"art(?:ykuł|\.)?\s*8", text, RE_FLAGS) and re.search(
            r"(zrównoważon\w+\s+inwestyc|impact|art(?:ykuł|\.)?\s*9|cel(?:em)? zrównoważon)", text, RE_FLAGS
        ):
            add(
                Finding(
                    rule_id="SFDR-8-vs-9",
                    code="SFDR",
                    title="Art. 8 sprzedawany jak art. 9",
                    severity="high",
                    message="Dokument łączy kwalifikację art. 8 SFDR z językiem celu zrównoważonych inwestycji / impact.",
                    snippet="W tekście występują jednocześnie odwołania do art. 8 i języka art. 9 / impact.",
                    advice="Nie nazywaj produktu art. 8 „sustainable / ESG impact”. Próg nazw ESMA: 80% + wykluczenia.",
                    source="SFDR · wytyczne ESMA nazw funduszy",
                )
            )

    if "transport" in sectors:
        if re.search(r"(emisj\w+|co2|co₂|ślad węgl)", text, RE_FLAGS) and not re.search(
            r"(iso\s*14083|countemissions)", text, RE_FLAGS
        ):
            add(
                Finding(
                    rule_id="COUNTEMISSIONS-2026-1030",
                    code="14083",
                    title="Liczba GHG bez wspólnej metodyki",
                    severity="medium",
                    message="Komunikat o emisjach transportu bez ISO 14083 / CountEmissionsEU.",
                    snippet="Wykryto język emisji/CO₂ przy transporcie, bez normy ISO 14083.",
                    advice="Jeśli podajesz liczbę — metodyka EN ISO 14083, granice, dane pierwotne vs default.",
                    source="rozp. (UE) 2026/1030",
                )
            )

    if "construction" in sectors:
        if re.search(r"(breeam|leed)", text, RE_FLAGS) and re.search(
            r"(zeroemisyjn|bezemisyjn|net[\s-]?zero|climate[\s-]?neutral)", text, RE_FLAGS
        ):
            add(
                Finding(
                    rule_id="NO-CERT-TRANSFER",
                    code="EPBD",
                    title="Certyfikat budynku ≠ definicja bezemisyjności",
                    severity="high",
                    message="BREEAM/LEED zestawione z językiem zeroemisyjności. EPBD art. 11 to osobna definicja prawna.",
                    snippet="W tekście są jednocześnie BREEAM/LEED oraz claim zero-/bezemisyjny.",
                    advice="Podaj schemat, wersję, poziom i zakres przestrzenny. „Bezemisyjny” tylko po checkliście EPBD art. 11.",
                    source="EPBD art. 11 · ECGT 4b",
                )
            )

    verdict = _verdict(findings, text)
    # keep at most 3 snippets per rule so the UI stays readable
    capped: list[Finding] = []
    per: dict[str, int] = {}
    for f in findings:
        per[f.rule_id] = per.get(f.rule_id, 0) + 1
        if per[f.rule_id] <= 3:
            capped.append(f)
    findings = capped
    return _pack(filename, text, findings, rules, extra={"legal_document_heuristic": legalish, "sectors": sectors, "verdict": verdict})


def _detect_sectors(text: str) -> list[str]:
    t = text.lower()
    found: list[str] = []
    if re.search(r"żywnoś|fmcg|mleko|kawa|kakao|organic|rolnictw", t):
        found.append("fmcg")
    if re.search(r"fundusz|sfdr|taksonom|esg\s+fundusz|artykuł\s*8|knf", t):
        found.append("finance")
    if re.search(r"dostaw|kurier|paczkomat|flot\w+|przesyłk|e-commerce|logist", t):
        found.append("transport")
    if re.search(r"kolekcj|tekstyl|odzież|bawełn|recyklat|join life|moda", t):
        found.append("textiles")
    if re.search(r"budyn|deweloper|breeam|leed|epbd|nieruchom|beton|epd", t):
        found.append("construction")
    return found


def _verdict(findings: list[Finding], text: str) -> dict:
    highs = [f for f in findings if f.severity == "high"]
    meds = [f for f in findings if f.severity == "medium"]
    has_claim = bool(findings) or bool(re.search(r"\b(eko|zrównoważ|climate|zeroemisyj|ecolabel)\w*", text, RE_FLAGS))

    if any(f.code in {"4c", "2a", "848"} for f in highs):
        key, label, color = "reject", "Nie publikuj", "rose"
        why = "Czarna lista (offset, własny znak lub zastrzeżone eko/bio) albo analogia UOKiK o wysokim ryzyku."
    elif highs:
        key, label, color = "reject", "Nie publikuj / natychmiast zawęź", "rose"
        why = "Ogólne hasło lub claim o całości bez zakresu. Od 27.09.2026 to ECGT; UOKiK karze to już na UPNPR."
    elif meds:
        key, label, color = "narrow", "Zawęź język", "amber"
        why = "Dowód albo metodyka są węższe niż hasło. Najczęstsza naprawa 4b."
    elif not has_claim:
        key, label, color = "clean", "Brak twierdzenia ekologicznego", "forest"
        why = "Nie wykryto haseł ESG do oceny. To nie jest certyfikat „czystości” dokumentu."
    else:
        key, label, color = "publish", "Do weryfikacji dowodu 1:1", "forest"
        why = "Hasła są konkretniejsze, ale i tak sprawdź zakres, datę dowodu i spójność kanałów."

    return {"key": key, "label": label, "color": color, "why": why, "high": len(highs), "medium": len(meds), "low": len(findings) - len(highs) - len(meds)}


def _pack(filename: str, text: str, findings: list[Finding], rules: dict, extra: dict) -> dict:
    order = {"high": 0, "medium": 1, "low": 2}
    findings.sort(key=lambda f: (order.get(f.severity, 9), f.code))
    preview = re.sub(r"\s+", " ", text)[:500]
    meta = rules.get("meta", {})
    return {
        "filename": filename,
        "chars": len(text),
        "preview": preview,
        "findings": [asdict(f) for f in findings],
        "counts": {
            "total": len(findings),
            "high": sum(1 for f in findings if f.severity == "high"),
            "medium": sum(1 for f in findings if f.severity == "medium"),
            "low": sum(1 for f in findings if f.severity == "low"),
        },
        "as_of": meta.get("as_of"),
        "ecgt_apply_from": meta.get("ecgt_apply_from"),
        "disclaimer": "Narzędzie analityczne, nie porada prawna. Silnik reguł z bazy wiedzy (stan 15.09.2026). Nie ocenia prawdziwości faktów poza tekstem pliku.",
        **extra,
    }
