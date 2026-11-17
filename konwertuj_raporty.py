#!/usr/bin/env python3
"""Konwertuje raporty współpracownika (.docx) na aktualny skład ESG Institute (sklad_raportu.py)."""

import re
from pathlib import Path

import docx
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph

from sklad_raportu import (
    DARK,
    MUTED,
    NAVY,
    VERDICT_AMBER,
    VERDICT_GREEN,
    VERDICT_RED,
    WHITE,
    ReportDoc,
    set_run,
    set_table_width,
)

ROOT = Path(__file__).resolve().parent

JOBS = [
    {
        "src": ROOT / "Agnieszka DYSK " / "Greenwashing tool" / "KRUK SA" / "Raport_Audytu_Greenwashingu_Grupa_KRUK_2025.docx",
        "dst": ROOT / "KRUK SA" / "Raport_Audytu_Greenwashingu_Grupa_KRUK_2025.docx",
        "ref": "Sygn. ESG/2026/10/12",
        "doc_name": "Audyt greenwashingu — Grupa Kapitałowa KRUK",
        "date": "12 października 2026 r.",
        "img": None,
    },
    {
        "src": ROOT / "Agnieszka DYSK " / "Greenwashing tool" / "KRUK SA" / "Raport_Audytu_Greenwashingu_KRUK_partnerzy_biznesowi.docx",
        "dst": ROOT / "KRUK SA" / "Raport_Audytu_Greenwashingu_KRUK_partnerzy_biznesowi.docx",
        "ref": "Sygn. ESG/2026/10/13",
        "doc_name": "Audyt greenwashingu — KRUK S.A. (Partnerzy biznesowi)",
        "date": "13 października 2026 r.",
        "img": ROOT / "KRUK SA" / "zrzut_partnerzy_biznesowi.png",
    },
    {
        "src": ROOT / "Agnieszka DYSK " / "Greenwashing tool" / "HI DESIGN" / "Raport_Audytu_Greenwashingu_Hi_DESIGN.docx",
        "dst": ROOT / "HI DESIGN" / "Raport_Audytu_Greenwashingu_Hi_DESIGN.docx",
        "ref": "Sygn. ESG/2026/10/14",
        "doc_name": "Audyt greenwashingu — HI DESIGN",
        "date": "14 października 2026 r.",
        "img": ROOT / "HI DESIGN" / "zrzut_strona_glowna.png",
    },
    {
        "src": ROOT / "Agnieszka DYSK " / "Greenwashing tool" / "MOWI" / "Raport_Audytu_Greenwashingu_MOWI_POLAND.docx",
        "dst": ROOT / "MOWI" / "Raport_Audytu_Greenwashingu_MOWI_POLAND.docx",
        "ref": "Sygn. ESG/2026/10/15",
        "doc_name": "Audyt greenwashingu — MOWI Poland",
        "date": "15 października 2026 r.",
        "img": ROOT / "MOWI" / "zrzut_mowi_pl.png",
    },
    {
        "src": ROOT / "Agnieszka DYSK " / "Raporty" / "Raport_Audytu_Greenwashingu_C-Select.docx",
        "dst": ROOT / "C-SELECT" / "Raport_Audytu_Greenwashingu_C-Select.docx",
        "ref": "Sygn. ESG/2026/11/05-CS",
        "doc_name": "Audyt greenwashingu — C-Select",
        "date": "5 listopada 2026 r.",
        "img": ROOT / "C-SELECT" / "zrzut_strony.png",
    },
    {
        "src": ROOT / "Agnieszka DYSK " / "Raporty" / "Raport_Audytu_Greenwashingu_PGF.docx",
        "dst": ROOT / "PGF" / "Raport_Audytu_Greenwashingu_PGF.docx",
        "ref": "Sygn. ESG/2026/11/10",
        "doc_name": "Audyt greenwashingu — Polska Grupa Farmaceutyczna S.A.",
        "date": "10 listopada 2026 r.",
        "img": ROOT / "PGF" / "zrzut_strony.png",
    },
    {
        "src": ROOT / "Agnieszka DYSK " / "Raporty" / "Raport_Audytu_Greenwashingu_ING_Bank_Slaski.docx",
        "dst": ROOT / "ING_BANK_SLASKI" / "Raport_Audytu_Greenwashingu_ING_Bank_Slaski.docx",
        "ref": "Sygn. ESG/2026/11/16",
        "doc_name": "Audyt greenwashingu — ING Bank Śląski S.A.",
        "date": "16 listopada 2026 r.",
        "img": ROOT / "ING_BANK_SLASKI" / "image2.png",
    },
]


def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield ("P", Paragraph(child, doc))
        elif isinstance(child, CT_Tbl):
            yield ("TBL", Table(child, doc))


def map_verdict(text):
    up = text.upper()
    if "DOPUSZCZONE" in up or "ZIELONY" in up:
        return VERDICT_GREEN, False
    if "WARUNKOW" in up or "ŻÓŁT" in up or "POMARAŃCZ" in up:
        return VERDICT_AMBER, True
    return VERDICT_RED, True


def styled_table(report, headers, rows):
    ncols = len(headers)
    table = report.doc.add_table(rows=1 + len(rows), cols=ncols)
    table.autofit = False
    if ncols == 2:
        widths = (Cm(5.8), Cm(11.6))
    elif ncols == 3:
        widths = (Cm(4.5), Cm(6.45), Cm(6.45))
    elif ncols == 4:
        widths = (Cm(1.4), Cm(5.0), Cm(5.5), Cm(5.5))
    else:
        w = Cm(17.4 / ncols)
        widths = tuple(w for _ in range(ncols))
    set_table_width(table, report.content, widths)
    for col, text in enumerate(headers):
        cell = table.rows[0].cells[col]
        report._prepare(cell, "08181F", "08181F", "4", 55, 55, 60, 60)
        p = report._cell_para(cell, True, "left")
        set_run(p.add_run(text), 8.5, True, False, WHITE)
    for r_idx, row in enumerate(rows):
        fill = "FFFFFF" if r_idx % 2 == 0 else "F7F9F8"
        for col in range(ncols):
            val = row[col] if col < len(row) else ""
            cell = table.rows[r_idx + 1].cells[col]
            report._prepare(cell, fill, "E4EAE7", "4", 50, 50, 60, 60)
            p = report._cell_para(cell, True, "left")
            set_run(p.add_run(val), 8.5, col == 0, False, NAVY if col == 0 else DARK)
    report.paragraph("", size=3, after=4, align="left")


def clean_p_text(p):
    return re.sub(r"\s+", " ", p.text).strip()


def convert_one(job):
    src = docx.Document(str(job["src"]))
    blocks = list(iter_blocks(src))

    report = ReportDoc(
        reference=job["ref"],
        document_name=job["doc_name"],
        audit_date=job["date"],
    )

    idx = 0
    n = len(blocks)
    title_done = False
    subtitle_done = False
    pending_figure_caption = None
    figure_inserted = False

    def maybe_insert_figure():
        nonlocal figure_inserted, pending_figure_caption
        if not figure_inserted and job["img"] and job["img"].exists():
            cap = pending_figure_caption or f"Zrzut ekranu analizowanej witryny ({job['date']}). Źródło: opracowanie własne."
            report.figure(job["img"], cap, width_cm=11.5)
            figure_inserted = True
            pending_figure_caption = None

    while idx < n:
        kind, obj = blocks[idx]
        if kind == "P":
            txt = clean_p_text(obj)
            if not txt:
                idx += 1
                continue
            if txt.startswith("ESG INSTITUTE") and "RAPORT AUDYTU" in txt.upper():
                idx += 1
                continue
            if not title_done and txt.startswith("Audyt"):
                report.title(txt)
                title_done = True
                idx += 1
                continue
            if title_done and not subtitle_done and not re.match(r"^\d+\.\s", txt):
                if txt.startswith("Zrzut "):
                    pending_figure_caption = txt
                    idx += 1
                    continue
                report.subtitle(txt)
                subtitle_done = True
                idx += 1
                continue
            if txt.startswith("Zrzut ") and not figure_inserted and idx < 12:
                pending_figure_caption = txt
                idx += 1
                continue

            # Skip old footer/sign-off blocks at the very end
            if txt.startswith("Raport sporządził") or txt.startswith("Poufne /"):
                idx += 1
                continue

            # Section H1: e.g. "1. Analiza Wstępna", "2. Analiza szczegółowa...", "3. Szczegółowa analiza..."
            if re.match(r"^[1-9]\.\s+[A-ZŻŹĆĄŚĘŁÓŃ]", txt) and len(txt) < 110:
                if txt.startswith("1. Analiza"):
                    report.h1("1. Analiza wstępna")
                elif txt.startswith("2. Analiza szczegółowa"):
                    maybe_insert_figure()
                    report.h1("2. Analiza szczegółowa — tabela zbiorcza naruszeń")
                else:
                    report.h1(txt)
                idx += 1
                continue

            # Subsections like "1. Kara Prezesa UOKiK" or "1.1. ..."
            if re.match(r"^(\d+\.\d+\.|\d+\.)\s+", txt) and len(txt) < 110 and any(r.bold for r in obj.runs):
                report.h2(txt)
                idx += 1
                continue

            # Case block: "PRZYPADEK 01: ..." or "Przypadek 01..."
            if re.match(r"^PRZYPADEK\s+\d+", txt, re.IGNORECASE):
                basis = "Ustawa o przeciwdziałaniu nieuczciwym praktykom rynkowym"
                risk = "WYSOKIE"
                if idx + 1 < n and blocks[idx + 1][0] == "P":
                    nxt = clean_p_text(blocks[idx + 1][1])
                    if "Podstawa prawna:" in nxt:
                        m_risk = re.search(r"Poziom ryzyka:\s*([A-ZŻŹĆĄŚĘŁÓŃ\s]+)$", nxt)
                        if m_risk:
                            risk = m_risk.group(1).strip()
                            basis = nxt[: m_risk.start()].replace("Podstawa prawna:", "").strip(" |·")
                        else:
                            basis = nxt.replace("Podstawa prawna:", "").strip(" |·")
                        idx += 1
                report.case(txt, basis, risk)
                idx += 1
                continue

            # Sub-labels inside a case
            if txt in ("Cytat źródłowy:", "Cytat źródłowy", "Ocena prawna:", "Ocena prawna", "Kierunek naprawy:", "Kierunek naprawy"):
                label = txt.rstrip(":")
                before = 2 if label == "Cytat źródłowy" else 6
                report.paragraph(label, size=10.5, bold=True, color=NAVY, before=before, after=2, align="left")
                idx += 1
                continue

            # Bold heading inside Section 3 (e.g. C-Select "Karta roli na stronie głównej")
            if len(txt) < 85 and any(r.bold for r in obj.runs) and not txt.endswith(".") and not txt.startswith("„"):
                report.h2(txt)
                idx += 1
                continue

            # Quote paragraph
            if txt.startswith("„") and (txt.endswith("”") or txt.endswith("”.")) and len(txt) < 500:
                report.quote("Cytat:", txt)
                idx += 1
                continue

            # Bullet-like paragraph
            if txt.startswith("•") or txt.startswith("- "):
                clean = txt.lstrip("•- ").strip()
                if ":" in clean[:60]:
                    lbl, rest = clean.split(":", 1)
                    report.bullet(lbl + ": ", rest.strip())
                else:
                    report.bullet("", clean)
                idx += 1
                continue

            report.paragraph(txt)
            idx += 1

        elif kind == "TBL":
            rows_txt = [
                [re.sub(r"\s+", " ", c.text).strip() for c in r.cells]
                for r in obj.rows
            ]
            if not rows_txt:
                idx += 1
                continue

            first_cell = rows_txt[0][0] if rows_txt[0] else ""

            # 1) Meta table
            if len(obj.columns) == 2 and first_cell == "Badany podmiot":
                meta_rows = []
                for r in rows_txt:
                    k, v = r[0], r[1]
                    if k == "Werdykt operacyjny":
                        v_mapped, alert = map_verdict(v)
                        meta_rows.append((k, v_mapped, alert))
                    else:
                        meta_rows.append((k, v, False))
                report.meta(meta_rows)
                idx += 1
                continue

            # 2) Verdict box
            if len(obj.rows) == 1 and len(obj.columns) == 1 and "WERDYKT" in first_cell.upper():
                paras = [
                    re.sub(r"\s+", " ", p.text).strip()
                    for p in obj.cell(0, 0).paragraphs
                    if p.text.strip()
                ]
                status, _ = map_verdict(paras[0] if paras else first_cell)
                body_paras = paras[1:] if len(paras) > 1 else [status]
                report.verdict(status, body_paras)
                idx += 1
                continue

            # 3) Claim box (Dozwolone sformułowanie)
            if len(obj.rows) == 1 and len(obj.columns) == 1 and "DOZWOLONE SFORMUŁOWANIE" in first_cell.upper():
                paras = [
                    re.sub(r"\s+", " ", p.text).strip()
                    for p in obj.cell(0, 0).paragraphs
                    if p.text.strip()
                ]
                body = " ".join(paras[1:]) if len(paras) > 1 else first_cell
                report.claim(body)
                idx += 1
                continue

            # 4) Findings table (5 cols: ID | Analizowane twierdzenie | ...)
            if len(obj.columns) == 5 and first_cell == "ID":
                data_rows = []
                for r in rows_txt[1:]:
                    if any(cell.strip() for cell in r):
                        data_rows.append(tuple(r[:5]))
                if data_rows:
                    report.findings(data_rows)
                idx += 1
                continue

            # 5) Any other table (e.g., Section 4 comparison table)
            headers = rows_txt[0]
            data_rows = rows_txt[1:]
            styled_table(report, headers, data_rows)
            idx += 1

    maybe_insert_figure()
    report.sign_off()
    job["dst"].parent.mkdir(parents=True, exist_ok=True)
    report.save(str(job["dst"]))
    print(f"OK -> {job['dst']}")


if __name__ == "__main__":
    for j in JOBS:
        convert_one(j)
