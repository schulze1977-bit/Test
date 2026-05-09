#!/usr/bin/env python3
"""
fill_insolvenz.py – AG Charlottenburg Insolvenzformular befüllen

Verwendung:
    python fill_insolvenz.py originalformular.pdf data.json ausgefuelltes_formular.pdf

Voraussetzung:
    pip install pypdf
"""

import json
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import NameObject, create_string_object
except ImportError:
    print("FEHLER: pypdf nicht installiert. Bitte 'pip install pypdf' ausführen.")
    sys.exit(1)


def eur(val) -> str:
    """Formatiert einen Betrag als deutsche Währungsstring."""
    try:
        return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "0,00"


def join(*parts) -> str:
    return " ".join(p for p in parts if p)


def load_data(json_path: str) -> dict:
    with open(json_path, encoding="utf-8") as f:
        return json.load(f)


def map_fields(d: dict) -> dict:
    """Mappt data.json-Felder auf die Feldnamen des AG-Charlottenburg-Formulars."""
    firma      = d.get("firmaName", "")
    rechtsform = d.get("rechtsform", "GmbH")
    gf_name    = d.get("gfName", "")
    gf_funkt   = d.get("gfFunktion", "Geschäftsführer")
    strasse    = join(d.get("strasse",""), d.get("hausnummer",""))
    plz_ort    = join(d.get("plz",""), d.get("ort",""))
    hr_gericht = d.get("hrGericht", "")
    hr_nummer  = d.get("hrNummer", "")

    gruende = d.get("insolvenzgruende", [])
    datum   = __import__("datetime").date.today().strftime("%d.%m.%Y")
    ort     = d.get("ort", "")

    # Vermögen
    verm   = d.get("vermoegen", {})
    def vb(key): return eur(verm.get(key, {}).get("betrag", 0)) if verm.get(key, {}).get("vorhanden") else "0,00"

    mapping = {
        # Seite 1
        "In meiner  unserer Eigenschaft als": gf_funkt,
        "Firma 1":  firma,
        "Firma 2":  f"{rechtsform} – {strasse}, {plz_ort}",
        "in":       ort,
        "Ort Datum": f"{ort}, {datum}",
        "Name":     gf_name,

        # Seite 2 – HR-Daten
        "Amtsgericht":                   hr_gericht,
        "1 Allgemeine AngabenRow1_2":    hr_nummer,
        "Firma":                         firma,

        # Insolvenzgründe (Checkboxen)
        "Zahlungsunfähigkeit":           "Ja" if "zahlungsunfaehigkeit" in gruende else "Nein",
        "drohende Zahlungsunfähigkeit":  "Ja" if "drohend" in gruende else "Nein",
        "Überschuldung":                 "Ja" if "ueberschuldung" in gruende else "Nein",

        # Seite 7 – Vermögen
        "Wert in EUR Gesamtbetrag":       vb("kasse"),
        "Wert in EUR GesamtbetragB":      vb("betriebsmittel"),
        "Wert in EUR GesamtbetragC":      vb("auftraege"),
        "Wert in EUR GesamtbetragD":      vb("aussenstaende"),
        "Wert in EUR GesamtbetragE":      vb("beteiligungen"),
        "Wert in EUR GesamtbetragF":      vb("grundstuecke"),
        "Wert in EUR GesamtbetragE":      vb("sonstiges"),

        # Betrieb / MA
        "Anzahl der Arbeitnehmer":        str(d.get("maAnzahl", "")),
    }

    # Anlage 1A – Gläubigerliste (bis 10)
    glaeubiger = d.get("glaeubiger", [])
    for i, g in enumerate(glaeubiger[:10]):
        row = "" if i == 0 else f"_{i+1}"
        nr  = f"NrRow{i+1}"
        name_field   = f"NameKurzbezeichnung und Anschrift des GläubigersRow{i+1}"
        grund_field  = f"ForderungsgrundO{row}"
        haupt_field  = f"Hauptforderung in EURO{row}"
        zinsen_field = f"Höhe in EURO{row}"
        kosten_field = f"KostenO{row}"
        summe_field  = f"Summe aller Hauptforderun gen des Gläu bigers in EURO{row}"

        anschrift = join(g.get("anschrift",""), g.get("plz",""), g.get("ort",""))
        gesamt    = (g.get("hauptforderung",0) or 0) + (g.get("zinsen",0) or 0) + (g.get("kosten",0) or 0)

        mapping.update({
            nr:           str(i+1),
            name_field:   join(g.get("name",""), anschrift),
            grund_field:  g.get("grund",""),
            haupt_field:  eur(g.get("hauptforderung",0)),
            zinsen_field: eur(g.get("zinsen",0)),
            kosten_field: eur(g.get("kosten",0)),
            summe_field:  eur(gesamt),
        })

    return mapping


def fill_pdf(template_path: str, data: dict, output_path: str) -> int:
    reader = PdfReader(template_path)
    writer = PdfWriter()

    # append() kopiert Seiten + Formularstruktur korrekt (clone_reader_document_root
    # schlägt bei komplexen PDFs mit IndexError fehl)
    writer.append(reader)

    mapping = map_fields(data)
    filled  = 0

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
            field_name_str = str(field_name)

            # Direkte Übereinstimmung
            value = mapping.get(field_name_str)
            if value is None:
                # Teilstring-Matching für robustere Zuordnung
                for k, v in mapping.items():
                    if k in field_name_str or field_name_str in k:
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


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(1)

    template, data_file, output = sys.argv[1], sys.argv[2], sys.argv[3]

    for p in [template, data_file]:
        if not Path(p).exists():
            print(f"FEHLER: Datei nicht gefunden: {p}")
            sys.exit(1)

    print(f"Lade Daten aus {data_file} …")
    data = load_data(data_file)

    print(f"Befülle Formular {template} …")
    count = fill_pdf(template, data, output)

    print(f"✓ Fertig! {count} Felder befüllt → {output}")


if __name__ == "__main__":
    main()
