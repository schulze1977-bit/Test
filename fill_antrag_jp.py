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
    """Befüllt den amtlichen Insolvenzantrag (juristische Personen) mit Falldaten."""
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
    hr_nr      = data.get("hrNummer", "")
    gruende    = data.get("insolvenzgruende", [])
    verm       = data.get("vermoegen", {})
    ort_datum  = f"{ort}, {heute}" if ort else heute

    def vb(key):
        v = verm.get(key, {})
        return eur(v.get("betrag", 0)) if v.get("vorhanden") else ""

    def vv(key):
        return verm.get(key, {}).get("vorhanden", False)

    # ── Textfelder ──────────────────────────────────────────────────────────
    text_map = {
        # Antragsteller (Deckblatt)
        "In meiner  unserer Eigenschaft als": gf_funkt,
        "Firma 1":    firma,
        "Firma 2":    rechtsform,
        "in":         ort,
        "Name":       gf_name,
        "Anschrift":  data.get("gfAnschrift", ""),
        "Tel":        data.get("gfTelefon", ""),
        "Telefon mobil": data.get("telefon", ""),
        "Telefon":    data.get("telefonFest", ""),
        "Telefax":    data.get("fax", ""),
        "email":      data.get("email", ""),
        # Verfahrensbevollmächtigter (Anwalt)
        "1_2":        data.get("anwalt", ""),
        # Allgemeine Angaben (Abschnitt 1)
        "1 Allgemeine AngabenRow1_2": firma,
        "1 Allgemeine AngabenRow2_2": f"{strasse}, {plz_ort}",
        "1 Allgemeine AngabenRow3_2": rechtsform,
        # Handelsregister
        "Amtsgericht": data.get("hrGericht", ""),
        "undefined_3": hr_nr,
        # Unternehmensdaten (Abschnitt 2)
        "Das Unternehmen ist tätig im Bereich":
            data.get("branche", ""),
        "Das Unternehmen ist allgemein anwaltlich vertreten durch":
            data.get("anwalt", ""),
        "Das Unternehmen ist steuerlich beraten durch":
            data.get("steuerberater", ""),
        "Gründungsgesellschafter waren":
            data.get("gesellschafterInfo", ""),
        "in voller Höhe bitte Belege beifügen":
            data.get("stammkapital", ""),
        # Betrieb
        "undefined_7": data.get("betriebEingestelltDatum", ""),
        # Mitarbeiter (Abschnitt 4)
        "1 Wie viele Mitarbeiter sind derzeit noch bei der Schuldnerin beschäftigt":
            str(data.get("maAnzahl", "")),
        "Auszubildende 1": str(data.get("azubiAnzahl", "")),
        "undefined_8":  data.get("kuendigungDatum", ""),
        # SV-Rückstände
        "undefined_9":  data.get("svBetrag", ""),
        "bei":          data.get("svTraeger", ""),
        # Lohn-Rückstände
        "undefined_10": data.get("lohnBetrag", ""),
        "für folgende Mitarbeiter": data.get("lohnRueckstandNamen", ""),
        # Geschäftsräume (Abschnitt 5)
        "undefined_11": data.get("mieteBetrag", ""),
        "undefined_12": data.get("vermieterName", ""),
        # Vermögen (Abschnitt 7) – Gesamtbeträge
        "Wert in EUR Gesamtbetrag":  vb("kasse"),
        "Wert in EUR GesamtbetragA": vb("betriebsmittel"),
        "Wert in EUR GesamtbetragB": vb("auftraege"),
        "Wert in EUR GesamtbetragC": vb("aussenstaende"),
        "Wert in EUR GesamtbetragD": vb("beteiligungen"),
        "Wert in EUR GesamtbetragE": vb("grundstuecke"),
        "Wert in EUR GesamtbetragF": vb("sonstiges"),
    }

    # Ort/Datum auf allen Formularseiten
    for suf in ("", "_2", "_3", "_4", "_5", "_6", "_7", "_8", "_9", "_10", "_11"):
        text_map[f"Ort Datum{suf}"] = ort_datum

    # ── Checkboxen ──────────────────────────────────────────────────────────
    check_map = {
        # Art des Antragstellers
        "Geschäftsführerininnen":
            gf_funkt in ("Geschäftsführer", "Geschäftsführerin"),
        "persönlich haftender Gesellschafterininnen":
            "haftender Gesellschafter" in gf_funkt,
        # Verfahrensbevollmächtigter
        "Verfahrensbevollmächtigter für das vorliegende Verfahren ist":
            bool(data.get("anwalt")),
        # Handelsregister ja/nein
        "ja":   bool(hr_nr),
        "nein": not bool(hr_nr),
        "HRB":  rechtsform in ("GmbH", "UG", "UG (haftungsbeschränkt)", "AG", "SE"),
        "HRA":  rechtsform in ("KG", "OHG", "GmbH & Co. KG", "KGaA"),
        # Insolvenzgründe
        "zahlungsunfähig":
            "zahlungsunfaehigkeit" in gruende,
        "voraussichtlich nicht in der Lage die bestehenden Zahlungspflichten bei Fälligkeit":
            "drohend" in gruende,
        "überschuldet":
            "ueberschuldung" in gruende,
        # Stammkapital
        "Das Stammkapital in eingezahlt": bool(data.get("stammkapital")),
        # Betrieb
        "noch nicht eingestellt":
            data.get("betriebStatus", "laufend") == "laufend",
        "eingestellt seit":
            data.get("betriebStatus", "laufend") == "eingestellt",
        # Mitarbeiter Kündigung
        "nein_2": not bool(data.get("kuendigungDatum")),
        "ja zum":  bool(data.get("kuendigungDatum")),
        # SV-Rückstände
        "keine Rückstände":     not data.get("svRueckstaende"),
        "Rückstände i H v EUR": bool(data.get("svRueckstaende")),
        # Lohn-Rückstände
        "keine Rückstände_2":     not data.get("lohnRueckstaende"),
        "Rückstände i H v EUR_2": bool(data.get("lohnRueckstaende")),
        # Geschäftsräume
        "befinden sich noch unter der o g Anschrift": bool(data.get("vermieterName")),
        "angemietet": data.get("mietePacht", "miete") != "pacht",
        "gepachtet zu einem monatlichen Entgelt i H v EUR":
            data.get("mietePacht") == "pacht",
        "Vermieter  Verpächter ist": bool(data.get("vermieterName")),
        "nicht vorhanden": not data.get("vermieterpfandrecht"),
        "vorhanden i H v EUR": bool(data.get("vermieterpfandrecht")),
        # Gläubigerverzeichnis
        "nach Anlage 1A einfaches Gläubigerverzeichnis": True,
        "nach Anlage 1B qualifiziertes Gläubigerverzeichnis nach  13 Absatz 1 Satz 4 bzw":
            True,
        # Vermögen vorhanden
        "ja in Höhe":   vv("kasse"),
        "ja in Höhe_2": vv("betriebsmittel"),
        "ja in Höhe_3": vv("auftraege"),
        "ja in Höhe_4": vv("aussenstaende"),
        "ja in Höhe_5": vv("beteiligungen"),
        "ja in Höhe_6": vv("grundstuecke"),
        "ja in Höhe_7": vv("sonstiges"),
    }

    # ── Anlage 1A – erste 10 Gläubiger direkt ins PDF ───────────────────────
    glaeubiger = data.get("glaeubiger", [])
    for i, g in enumerate(glaeubiger[:10]):
        n   = i + 1
        suf = "" if n == 1 else f"_{n}"
        adr = ", ".join(filter(None, [
            g.get("anschrift",""), g.get("plz",""), g.get("ort","")
        ]))
        name_adr = " – ".join(filter(None, [g.get("name",""), adr]))
        gesamt = sum(float(g.get(k, 0) or 0)
                     for k in ("hauptforderung", "zinsen", "kosten"))
        text_map.update({
            f"NrRow{n}":
                str(n),
            f"NameKurzbezeichnung und Anschrift des GläubigersRow{n}":
                name_adr,
            f"ForderungsgrundO{suf}":
                g.get("grund", ""),
            f"Hauptforderung in EURO{suf}":
                eur(g.get("hauptforderung", 0)),
            f"Höhe in EURO{suf}":
                eur(g.get("zinsen", 0)),
            f"berechnet bis zumO{suf}":
                g.get("zinsenBis", ""),
            f"KostenO{suf}":
                eur(g.get("kosten", 0)),
            f"Forderung durch Sonderrechte gesichertO{suf}":
                g.get("forderungGesichert", "nein"),
            f"Summe aller Hauptforderun gen des Gläu bigers in EURO{suf}":
                eur(gesamt),
        })

    # ── PDF schreiben ────────────────────────────────────────────────────────
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
            fname = str(annot.get("/T", ""))
            ftype = str(annot.get("/FT", ""))

            if ftype == "/Tx" and fname in text_map and text_map[fname]:
                annot.update({
                    NameObject("/V"): create_string_object(str(text_map[fname])),
                })
                filled += 1
            elif ftype == "/Btn" and fname in check_map:
                state = NameObject("/On") if check_map[fname] else NameObject("/Off")
                annot.update({
                    NameObject("/V"):  state,
                    NameObject("/AS"): state,
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
