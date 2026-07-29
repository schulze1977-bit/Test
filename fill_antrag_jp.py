#!/usr/bin/env python3
"""
fill_antrag_jp.py – Anlage 1A & 1B Gläubigerverzeichnisse erstellen
                    (juristische Personen, AG Charlottenburg)

Verwendung:
    python fill_antrag_jp.py data.json Anlage_1A.docx Anlage_1B.docx
    python fill_antrag_jp.py data.json Anlage_1A.docx Anlage_1B.docx antrag.pdf antrag_ausgefuellt.pdf

Voraussetzung:
    pip install python-docx pypdf
"""

import json
import sys
from copy import deepcopy
from datetime import date
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("FEHLER: python-docx nicht installiert. Bitte 'pip install python-docx' ausführen.")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ─────────────────────────────────────────────────────────────

def eur(val) -> str:
    try:
        return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "0,00"


def ja_nein(val) -> str:
    if str(val).lower() in ("ja", "yes", "1", "true", "x"):
        return "X"
    return ""


def fmt_date(val) -> str:
    """Konvertiert ISO-Datum (2025-12-31) in deutsches Format (31.12.2025)."""
    if not val:
        return ""
    try:
        d = date.fromisoformat(str(val))
        return d.strftime("%d.%m.%Y")
    except ValueError:
        return str(val)


def set_cell_text(cell, text: str, bold: bool = False, font_size: int = 9):
    cell.text = ""
    para = cell.paragraphs[0]
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)


def shade_cell(cell, fill_hex: str = "D9D9D9"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


# ─────────────────────────────────────────────────────────────
# ANLAGE 1A – Vereinfachtes Gläubigerverzeichnis
# ─────────────────────────────────────────────────────────────

ANLAGE_1A_HEADERS = [
    "Nr.",
    "Name / Kurzbezeichnung und\nAnschrift des Gläubigers",
    "Nahe-\nstehende\nPerson\n§ 138",
    "Forderungs-\ngrund",
    "Haupt-\nforderung\nin EUR",
    "Zinsen\nHöhe in\nEUR",
    "Zinsen\nber. bis\nzum",
    "Kosten\nin EUR",
    "Forder.\ntitulie\n-rt",
    "Forder.\ndurch Sonder-\nrechte ges.",
    "Summe\nHauptford.\nin EUR",
]

COL_WIDTHS_1A_CM = [0.8, 5.5, 1.3, 2.8, 2.0, 1.8, 1.8, 1.8, 1.3, 1.8, 2.0]


def build_anlage_1a(data: dict, output_path: str, template_path: str = None):
    doc = Document(template_path) if template_path and Path(template_path).exists() else Document()

    firma   = data.get("firmaName", "")
    datum   = date.today().strftime("%d.%m.%Y")

    # ── Kopfzeilen ──────────────────────────────────────────────
    doc.add_heading("Anlage 1 A", level=1)
    doc.add_paragraph(f"Zum Eröffnungsantrag der/des: {firma}     vom: {datum}")
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("Vereinfachtes Gläubiger- und Forderungsverzeichnis")
    run.bold = True
    run.font.size = Pt(12)
    doc.add_paragraph("(Verzeichnis der Gläubiger und ihrer gegen den Schuldner gerichteten Forderungen)")
    doc.add_paragraph()

    glaeubiger = data.get("glaeubiger", [])

    # ── Tabelle ─────────────────────────────────────────────────
    table = doc.add_table(rows=1, cols=len(ANLAGE_1A_HEADERS))
    table.style = "Table Grid"

    # Kopfzeile
    hdr_row1 = table.rows[0]
    for ci, hdr in enumerate(ANLAGE_1A_HEADERS):
        cell = hdr_row1.cells[ci]
        set_cell_text(cell, hdr, bold=True, font_size=8)
        shade_cell(cell, "BDD7EE")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Datenzeilen
    for i, g in enumerate(glaeubiger):
        row = table.add_row()
        anschrift = " ".join(filter(None, [g.get("anschrift",""), g.get("plz",""), g.get("ort","")]))
        name_full = "\n".join(filter(None, [g.get("name",""), anschrift]))
        gesamt = (g.get("hauptforderung",0) or 0) + (g.get("zinsen",0) or 0) + (g.get("kosten",0) or 0)

        cells_data = [
            str(i + 1),
            name_full,
            ja_nein(g.get("nahestehendeP138", "")),
            g.get("grund", ""),
            eur(g.get("hauptforderung", 0)),
            eur(g.get("zinsen", 0)),
            fmt_date(g.get("zinsenBis", "")),
            eur(g.get("kosten", 0)),
            ja_nein(g.get("forderungTituliert", "")),
            ja_nein(g.get("forderungGesichert", "")),
            eur(gesamt),
        ]
        for ci, val in enumerate(cells_data):
            cell = row.cells[ci]
            set_cell_text(cell, val, font_size=9)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if ci != 1 and ci != 3 else WD_ALIGN_PARAGRAPH.LEFT

    # Summenzeile
    sum_haupt = sum(g.get("hauptforderung", 0) or 0 for g in glaeubiger)
    sum_zinsen = sum(g.get("zinsen", 0) or 0 for g in glaeubiger)
    sum_kosten = sum(g.get("kosten", 0) or 0 for g in glaeubiger)
    sum_ges    = sum_haupt + sum_zinsen + sum_kosten

    sum_row = table.add_row()
    shade_cell(sum_row.cells[0], "F2F2F2")
    set_cell_text(sum_row.cells[0], "Σ", bold=True, font_size=9)
    set_cell_text(sum_row.cells[1], f"Gesamt: {len(glaeubiger)} Gläubiger", bold=True, font_size=9)
    set_cell_text(sum_row.cells[4], eur(sum_haupt), bold=True, font_size=9)
    set_cell_text(sum_row.cells[5], eur(sum_zinsen), bold=True, font_size=9)
    set_cell_text(sum_row.cells[7], eur(sum_kosten), bold=True, font_size=9)
    set_cell_text(sum_row.cells[10], eur(sum_ges), bold=True, font_size=9)
    for ci in range(len(ANLAGE_1A_HEADERS)):
        shade_cell(sum_row.cells[ci], "F2F2F2")

    # Fußnote
    doc.add_paragraph()
    fn = doc.add_paragraph()
    fn.add_run(
        "Ich versichere die Richtigkeit und Vollständigkeit der in diesem Gläubigerverzeichnis "
        "enthaltenen Angaben."
    ).font.size = Pt(9)
    doc.add_paragraph()
    doc.add_paragraph("(Ort, Datum)                                    (Unterschrift)")

    doc.save(output_path)
    print(f"  ✓ Anlage 1A gespeichert: {output_path}  ({len(glaeubiger)} Gläubiger)")


# ─────────────────────────────────────────────────────────────
# ANLAGE 1B – Qualifiziertes Gläubigerverzeichnis
# ─────────────────────────────────────────────────────────────

KATEGORIEN_1B = [
    ("absonderung",    "Absonderungsberechtigte Gläubiger (§§ 49 ff. InsO)",    "DCE6F1"),
    ("aussonderung",   "Aussonderungsberechtigte Gläubiger (§ 47 InsO)",        "EBF1DD"),
    ("masse",          "Massegläubiger (§ 53 InsO)",                            "FFF2CC"),
    ("nichtNachrangig","Nicht nachrangige Insolvenzgläubiger (§ 38 InsO)",      "F2F2F2"),
    ("nachrangig",     "Nachrangige Insolvenzgläubiger (§ 39 InsO)",            "FDE9D9"),
    ("sonstige",       "Sonstige Gläubiger",                                    "FFFFFF"),
]

ANLAGE_1B_HEADERS = [
    "Nr.",
    "Name / Anschrift des Gläubigers",
    "Nahe-\nstehende\nPerson\n§ 138",
    "Forderungs-\ngrund",
    "Haupt-\nforderung\nin EUR",
    "Zinsen\nHöhe EUR",
    "Zinsen\nber. bis",
    "Kosten\nEUR",
    "Tituliert",
    "Gesichert",
    "Summe EUR",
]


def build_anlage_1b(data: dict, output_path: str):
    doc = Document()

    firma = data.get("firmaName", "")
    datum = date.today().strftime("%d.%m.%Y")

    doc.add_heading("Anlage 1 B", level=1)
    doc.add_paragraph(f"Zum Eröffnungsantrag der/des: {firma}     vom: {datum}")
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("Qualifiziertes Gläubiger- und Forderungsverzeichnis")
    run.bold = True
    run.font.size = Pt(12)
    doc.add_paragraph()

    glaeubiger = data.get("glaeubiger", [])

    # ── Gläubiger nach Kategorie gruppieren ────────────────────
    kategorien_map = {k: [] for k, _, _ in KATEGORIEN_1B}
    for g in glaeubiger:
        kat = g.get("kategorie", "")
        if kat in kategorien_map:
            kategorien_map[kat].append(g)
        else:
            kategorien_map["sonstige"].append(g)

    lfd_nr = 1
    gesamt_summe = 0.0

    for kat_key, kat_label, kat_farbe in KATEGORIEN_1B:
        gruppe = kategorien_map.get(kat_key, [])
        if not gruppe:
            continue

        # Abschnittskopf
        p = doc.add_paragraph()
        run = p.add_run(kat_label)
        run.bold = True
        run.font.size = Pt(10)
        p.paragraph_format.space_before = Pt(12)

        # Tabelle für diese Gruppe
        table = doc.add_table(rows=1, cols=len(ANLAGE_1B_HEADERS))
        table.style = "Table Grid"

        # Kopfzeile
        hdr = table.rows[0]
        for ci, h in enumerate(ANLAGE_1B_HEADERS):
            cell = hdr.cells[ci]
            set_cell_text(cell, h, bold=True, font_size=8)
            shade_cell(cell, kat_farbe)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        gruppe_summe = 0.0
        for g in gruppe:
            row = table.add_row()
            anschrift = " ".join(filter(None, [g.get("anschrift",""), g.get("plz",""), g.get("ort","")]))
            name_full = "\n".join(filter(None, [g.get("name",""), anschrift]))
            gesamt = (g.get("hauptforderung",0) or 0) + (g.get("zinsen",0) or 0) + (g.get("kosten",0) or 0)
            gruppe_summe += gesamt

            cells_data = [
                str(lfd_nr),
                name_full,
                ja_nein(g.get("nahestehendeP138", "")),
                g.get("grund", ""),
                eur(g.get("hauptforderung", 0)),
                eur(g.get("zinsen", 0)),
                fmt_date(g.get("zinsenBis", "")),
                eur(g.get("kosten", 0)),
                ja_nein(g.get("forderungTituliert", "")),
                ja_nein(g.get("forderungGesichert", "")),
                eur(gesamt),
            ]
            for ci, val in enumerate(cells_data):
                cell = row.cells[ci]
                set_cell_text(cell, val, font_size=9)
                cell.paragraphs[0].alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if ci in (1, 3) else WD_ALIGN_PARAGRAPH.CENTER
                )
            lfd_nr += 1

        # Gruppensum
        sum_row = table.add_row()
        set_cell_text(sum_row.cells[1], f"Zwischensumme {kat_label.split('(')[0].strip()}", bold=True, font_size=9)
        set_cell_text(sum_row.cells[10], eur(gruppe_summe), bold=True, font_size=9)
        for ci in range(len(ANLAGE_1B_HEADERS)):
            shade_cell(sum_row.cells[ci], "F2F2F2")

        gesamt_summe += gruppe_summe
        doc.add_paragraph()

    # Gesamtsumme
    p = doc.add_paragraph()
    run = p.add_run(f"Gesamtforderungen aller Gläubiger:  {eur(gesamt_summe)} EUR")
    run.bold = True
    run.font.size = Pt(11)

    # Fußnote
    doc.add_paragraph()
    fn = doc.add_paragraph()
    fn.add_run(
        "Ich versichere die Richtigkeit und Vollständigkeit der in diesem Gläubigerverzeichnis "
        "enthaltenen Angaben."
    ).font.size = Pt(9)
    doc.add_paragraph()
    doc.add_paragraph("(Ort, Datum)                                    (Unterschrift)")

    doc.save(output_path)
    kategorien_mit_inhalt = sum(1 for k, _, _ in KATEGORIEN_1B if kategorien_map.get(k))
    print(f"  ✓ Anlage 1B gespeichert: {output_path}  ({lfd_nr-1} Gläubiger, {kategorien_mit_inhalt} Kategorien)")


# ─────────────────────────────────────────────────────────────
# PDF-Antrag befüllen (optional, wie fill_insolvenz.py)
# ─────────────────────────────────────────────────────────────

def fill_pdf_antrag(data: dict, template_path: str, output_path: str) -> int:
    try:
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import NameObject, create_string_object
    except ImportError:
        print("FEHLER: pypdf nicht installiert. Bitte 'pip install pypdf' ausführen.")
        return 0

    firma      = data.get("firmaName", "")
    rechtsform = data.get("rechtsform", "GmbH")
    gf_name    = data.get("gfName", "")
    gf_funkt   = data.get("gfFunktion", "Geschäftsführer")
    strasse    = " ".join(filter(None, [data.get("strasse",""), data.get("hausnummer","")]))
    plz_ort    = " ".join(filter(None, [data.get("plz",""), data.get("ort","")]))
    ort        = data.get("ort", "")
    heute      = date.today().strftime("%d.%m.%Y")

    gruende = data.get("insolvenzgruende", [])
    verm    = data.get("vermoegen", {})
    def vb(key): return eur(verm.get(key, {}).get("betrag", 0)) if verm.get(key, {}).get("vorhanden") else "0,00"

    mapping = {
        "In meiner  unserer Eigenschaft als": gf_funkt,
        "Firma 1":  firma,
        "Firma 2":  f"{rechtsform} – {strasse}, {plz_ort}",
        "in":       ort,
        "Ort Datum": f"{ort}, {heute}",
        "Name":     gf_name,
        "Amtsgericht":                  data.get("hrGericht", ""),
        "1 Allgemeine AngabenRow1_2":   data.get("hrNummer", ""),
        "Firma":                        firma,
        "Zahlungsunfähigkeit":          "Ja" if "zahlungsunfaehigkeit" in gruende else "Nein",
        "drohende Zahlungsunfähigkeit": "Ja" if "drohend"              in gruende else "Nein",
        "Überschuldung":                "Ja" if "ueberschuldung"        in gruende else "Nein",
        "Wert in EUR Gesamtbetrag":     vb("kasse"),
        "Wert in EUR GesamtbetragB":    vb("betriebsmittel"),
        "Wert in EUR GesamtbetragC":    vb("auftraege"),
        "Wert in EUR GesamtbetragD":    vb("aussenstaende"),
        "Wert in EUR GesamtbetragE":    vb("beteiligungen"),
        "Wert in EUR GesamtbetragF":    vb("grundstuecke"),
        "Anzahl der Arbeitnehmer":      str(data.get("maAnzahl", "")),
    }

    # Anlage 1A – erste 10 Gläubiger direkt ins PDF
    glaeubiger = data.get("glaeubiger", [])
    for i, g in enumerate(glaeubiger[:10]):
        row_suf = "" if i == 0 else f"_{i+1}"
        anschrift = " ".join(filter(None, [g.get("anschrift",""), g.get("plz",""), g.get("ort","")]))
        gesamt = (g.get("hauptforderung",0) or 0) + (g.get("zinsen",0) or 0) + (g.get("kosten",0) or 0)
        mapping.update({
            f"NrRow{i+1}":                                            str(i + 1),
            f"NameKurzbezeichnung und Anschrift des GläubigersRow{i+1}":
                " ".join(filter(None, [g.get("name",""), anschrift])),
            f"ForderungsgrundO{row_suf}":    g.get("grund",""),
            f"Hauptforderung in EURO{row_suf}": eur(g.get("hauptforderung",0)),
            f"Höhe in EURO{row_suf}":        eur(g.get("zinsen",0)),
            f"KostenO{row_suf}":             eur(g.get("kosten",0)),
            f"Summe aller Hauptforderun gen des Gläu bigers in EURO{row_suf}": eur(gesamt),
        })

    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    filled = 0
    for page in writer.pages:
        annotations = page.get("/Annots")
        if not annotations:
            continue
        for annot_ref in annotations:
            try:
                annot = annot_ref.get_object() if hasattr(annot_ref, "get_object") else annot_ref
            except Exception:
                continue
            if annot.get("/Subtype") != "/Widget":
                continue
            field_name = annot.get("/T")
            if not field_name:
                continue
            fname = str(field_name)
            value = mapping.get(fname)
            if value is None:
                for k, v in mapping.items():
                    if k in fname or fname in k:
                        value = v
                        break
            if value is not None:
                annot.update({
                    NameObject("/V"): create_string_object(str(value)),
                    NameObject("/AP"): create_string_object(""),
                })
                filled += 1

    with open(output_path, "wb") as f:
        writer.write(f)
    return filled


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    json_path  = sys.argv[1]
    out_1a     = sys.argv[2]
    out_1b     = sys.argv[3]
    pdf_in     = sys.argv[4] if len(sys.argv) > 4 else None
    pdf_out    = sys.argv[5] if len(sys.argv) > 5 else None

    if not Path(json_path).exists():
        print(f"FEHLER: Datei nicht gefunden: {json_path}")
        sys.exit(1)

    print(f"Lade Daten aus {json_path} …")
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    glaeubiger = data.get("glaeubiger", [])
    print(f"  {len(glaeubiger)} Gläubiger geladen")

    print("\nErstelle Anlage 1A …")
    build_anlage_1a(data, out_1a)

    print("Erstelle Anlage 1B …")
    build_anlage_1b(data, out_1b)

    if pdf_in and pdf_out:
        if not Path(pdf_in).exists():
            print(f"FEHLER: PDF-Vorlage nicht gefunden: {pdf_in}")
        else:
            print(f"\nBefülle Antrag-PDF {pdf_in} …")
            count = fill_pdf_antrag(data, pdf_in, pdf_out)
            print(f"  ✓ PDF-Antrag gespeichert: {pdf_out}  ({count} Felder)")

    print("\n✓ Fertig!")


if __name__ == "__main__":
    main()
