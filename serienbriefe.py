#!/usr/bin/env python3
"""
serienbriefe.py – Word-Serienbriefe für alle Gläubiger erstellen

Verwendung:
    python serienbriefe.py data.json
    python serienbriefe.py data.json --vorlage vorlage_brief.txt

Ausgabe:
    Unterordner  Serienbriefe\\Brief_01_GlaeubigerName.docx  usw.

Optionale Textvorlage (vorlage_brief.txt):
    Zeile 1 = Betreff
    Rest    = Brieftext mit Platzhaltern

Platzhalter:
    {GLAEUBIGER_NAME}   {ANSCHRIFT}   {PLZ_ORT}
    {FORDERUNGSBETRAG}  {ZINSEN}      {KOSTEN}     {GESAMT}
    {FORDERUNGSGRUND}   {VERTRETER}
    {FIRMA_NAME}        {GF_NAME}     {DATUM}

Voraussetzung:
    pip install python-docx
"""

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("FEHLER: python-docx nicht installiert. Bitte 'pip install python-docx' ausführen.")
    sys.exit(1)


DEFAULT_BETREFF = "Insolvenzverfahren {FIRMA_NAME} – Forderungsanmeldung"

DEFAULT_VORLAGE = """
Sehr geehrte Damen und Herren,

hiermit zeige ich an, dass ich das Insolvenzverfahren über das Vermögen der

    {FIRMA_NAME}

vor dem Amtsgericht Charlottenburg, Berlin, beantrage.

Ihre Forderungen gegen die Schuldnerin lauten nach meinen Unterlagen wie folgt:

    Forderungsgrund:    {FORDERUNGSGRUND}
    Hauptforderung:     {FORDERUNGSBETRAG}
    Zinsen:             {ZINSEN}
    Kosten:             {KOSTEN}
    ─────────────────────────────────────
    Gesamtbetrag:       {GESAMT}

Ich bitte Sie, mir Ihre Forderungen zur Anmeldung im Insolvenzverfahren mitzuteilen,
insbesondere etwaige Sicherungsrechte an den Aktiven der Schuldnerin anzugeben.

Mit freundlichen Grüßen

{GF_NAME}
– im Auftrag von –
{FIRMA_NAME}
""".strip()


def eur(val) -> str:
    try:
        return f"{float(val):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "0,00 €"


def sanitize(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name)[:60]


def load_vorlage(path: str | None) -> tuple[str, str]:
    """Gibt (betreff, text) zurück."""
    if not path or not Path(path).exists():
        return DEFAULT_BETREFF, DEFAULT_VORLAGE
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    betreff = lines[0].strip() if lines else DEFAULT_BETREFF
    text    = "\n".join(lines[1:]).strip() if len(lines) > 1 else DEFAULT_VORLAGE
    return betreff, text


def render(template: str, ctx: dict) -> str:
    for k, v in ctx.items():
        template = template.replace(f"{{{k}}}", str(v))
    return template


def build_doc(betreff: str, text: str, ctx: dict) -> Document:
    doc = Document()

    # Seitenränder
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    # Absender
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"{ctx['FIRMA_NAME']} · {ctx.get('DATUM','')}")
    run.font.size = Pt(10)
    run.font.color.rgb = __import__("docx.shared", fromlist=["RGBColor"]).RGBColor(0x6b, 0x7a, 0x99)

    doc.add_paragraph()

    # Empfänger
    empf_lines = [ctx["GLAEUBIGER_NAME"]]
    if ctx.get("ANSCHRIFT"): empf_lines.append(ctx["ANSCHRIFT"])
    if ctx.get("PLZ_ORT"):   empf_lines.append(ctx["PLZ_ORT"])
    for line in empf_lines:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # Datum
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(ctx["DATUM"])

    doc.add_paragraph()

    # Betreff
    p = doc.add_paragraph()
    run = p.add_run(render(betreff, ctx))
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()

    # Brieftext
    for paragraph in render(text, ctx).split("\n\n"):
        lines = paragraph.strip()
        if lines:
            p = doc.add_paragraph(lines)
            p.paragraph_format.space_after = Pt(6)

    return doc


def main():
    parser = argparse.ArgumentParser(description="Serienbriefe für Insolvenzgläubiger erstellen")
    parser.add_argument("data_json", help="Pfad zur data.json Datei")
    parser.add_argument("--vorlage", help="Optionale Textvorlage (vorlage_brief.txt)", default=None)
    args = parser.parse_args()

    data_path = Path(args.data_json)
    if not data_path.exists():
        print(f"FEHLER: {data_path} nicht gefunden")
        sys.exit(1)

    with open(data_path, encoding="utf-8") as f:
        data = json.load(f)

    glaeubiger = data.get("glaeubiger", [])
    if not glaeubiger:
        print("HINWEIS: Keine Gläubiger in data.json gefunden.")
        sys.exit(0)

    betreff, text_tmpl = load_vorlage(args.vorlage)

    out_dir = data_path.parent / "Serienbriefe"
    out_dir.mkdir(exist_ok=True)

    datum    = date.today().strftime("%d.%m.%Y")
    firma    = data.get("firmaName", "Unbekannt")
    gf_name  = data.get("gfName", "")

    for i, g in enumerate(glaeubiger, 1):
        gesamt = (g.get("hauptforderung",0) or 0) + (g.get("zinsen",0) or 0) + (g.get("kosten",0) or 0)
        ctx = {
            "GLAEUBIGER_NAME":  g.get("name", ""),
            "ANSCHRIFT":        g.get("anschrift", ""),
            "PLZ_ORT":          f"{g.get('plz','')} {g.get('ort','')}".strip(),
            "FORDERUNGSBETRAG": eur(g.get("hauptforderung", 0)),
            "ZINSEN":           eur(g.get("zinsen", 0)),
            "KOSTEN":           eur(g.get("kosten", 0)),
            "GESAMT":           eur(gesamt),
            "FORDERUNGSGRUND":  g.get("grund", ""),
            "VERTRETER":        g.get("vertreter", ""),
            "FIRMA_NAME":       firma,
            "GF_NAME":          gf_name,
            "DATUM":            datum,
        }

        doc = build_doc(betreff, text_tmpl, ctx)
        name_safe = sanitize(g.get("name","Glaeubiger"))
        filename  = out_dir / f"Brief_{i:02d}_{name_safe}.docx"
        doc.save(filename)
        print(f"  ✓ {filename.name}")

    print(f"\n{len(glaeubiger)} Briefe erstellt → {out_dir}")


if __name__ == "__main__":
    main()
