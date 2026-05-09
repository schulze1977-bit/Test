#!/usr/bin/env python3
"""
import_adobe_pdf.py – Adobe-PDF oder XFDF in data.json umwandeln

Unterstützt:
  • Ausgefülltes Adobe Acrobat PDF  (pypdf liest Formularfelder)
  • XFDF-Export aus Adobe Acrobat   (XML, Standard-Python)

Verwendung:
    python import_adobe_pdf.py ausgefuelltes_formular.pdf data.json
    python import_adobe_pdf.py formular.xfdf             data.json

Voraussetzung:
    pip install pypdf
"""

import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


# ── Mapping Adobe-Feldname → data.json-Schlüssel ───────────────────────────
#
# Format: "Adobe-Feldname (oder Teilstring)": "json-Schlüssel"
# Teilstrings werden case-insensitiv verglichen.

FIELD_MAP = {
    # Firmendaten
    "firma":                    "firmaName",
    "firmenname":               "firmaName",
    "rechtsform":               "rechtsform",
    "straße":                   "strasse",
    "strasse":                  "strasse",
    "hausnummer":               "hausnummer",
    "plz":                      "plz",
    "postleitzahl":             "plz",
    "ort":                      "ort",
    "telefon":                  "telefon",
    "e-mail":                   "email",
    "email":                    "email",

    # Handelsregister
    "amtsgericht":              "hrGericht",
    "hr-nummer":                "hrNummer",
    "hrb":                      "hrNummer",
    "handelsregisternummer":    "hrNummer",
    "stammkapital":             "stammkapital",

    # Geschäftsführer
    "geschäftsführer":          "gfName",
    "vertretungsberechtigte":   "gfName",
    "gf-name":                  "gfName",
    "funktion":                 "gfFunktion",

    # Insolvenzgrund
    "zahlungsunfähigkeit":      "_grund_zahlungsunfaehigkeit",
    "drohende":                 "_grund_drohend",
    "überschuldung":            "_grund_ueberschuldung",

    # Branche
    "branche":                  "branche",
    "wirtschaftszweig":         "branche",

    # Mitarbeiter
    "arbeitnehmer":             "maAnzahl",
    "mitarbeiter":              "maAnzahl",

    # Mietverhältnis
    "vermieter":                "vermieterName",
    "mietrückstand":            "mietRueckstaendeBetrag",
    "mietrückstände":           "mietRueckstaendeBetrag",

    # Vermögen
    "kasse":                    "_verm_kasse",
    "betriebsmittel":           "_verm_betriebsmittel",
    "aufträge":                 "_verm_auftraege",
    "auftraege":                "_verm_auftraege",
    "außenstände":              "_verm_aussenstaende",
    "aussenstaende":            "_verm_aussenstaende",
    "beteiligungen":            "_verm_beteiligungen",
    "grundstücke":              "_verm_grundstuecke",
    "grundstuecke":             "_verm_grundstuecke",
    "sonstiges vermögen":       "_verm_sonstiges",
}

# Gläubigerfelder (Anlage 1A)
GL_NAME_PATTERNS  = ["namekurzbezeichnung", "gläubiger"]
GL_GRUND_PATTERNS = ["forderungsgrund"]
GL_HAUPT_PATTERNS = ["hauptforderung"]
GL_ZINS_PATTERNS  = ["höhe in euro", "zinsen"]
GL_KOST_PATTERNS  = ["kosten"]
GL_NR_PATTERNS    = ["nrrow"]


# ── PDF lesen ──────────────────────────────────────────────────────────────

def read_pdf_fields(path: str) -> dict:
    try:
        from pypdf import PdfReader
    except ImportError:
        print("FEHLER: pypdf nicht installiert. Bitte 'pip install pypdf' ausführen.")
        sys.exit(1)

    reader = PdfReader(path)
    fields = reader.get_fields()
    if not fields:
        print("HINWEIS: Keine Formularfelder gefunden. Das PDF könnte nicht interaktiv sein.")
        return {}

    return {name: (fld.get("/V") or "") for name, fld in fields.items()}


# ── XFDF lesen ─────────────────────────────────────────────────────────────

def read_xfdf_fields(path: str) -> dict:
    tree   = ET.parse(path)
    root   = tree.getroot()
    ns     = {"xfdf": "http://ns.adobe.com/xfdf/"}
    fields = {}

    # XFDF hat <fields><field name="…"><value>…</value></field></fields>
    for field in root.findall(".//xfdf:field", ns) or root.findall(".//field"):
        name  = field.get("name", "")
        value_el = field.find("{http://ns.adobe.com/xfdf/}value") or field.find("value")
        value = value_el.text or "" if value_el is not None else ""
        if name:
            fields[name] = value

    return fields


# ── Mapping ────────────────────────────────────────────────────────────────

def map_fields_to_data(raw: dict) -> dict:
    """Wandelt die rohen Formularfelder in das data.json-Schema um."""
    data = {
        "firmaName": "", "rechtsform": "GmbH",
        "strasse": "", "hausnummer": "", "plz": "", "ort": "",
        "telefon": "", "email": "", "aktenzeichen": "",
        "hrGericht": "", "hrNummer": "", "stammkapital": "",
        "gfName": "", "gfFunktion": "Geschäftsführer", "gfAnschrift": "",
        "gesellschafter": [],
        "insolvenzgruende": [],
        "branche": "", "anwalt": "", "steuerberater": "", "hinweise": "",
        "betriebStatus": "laufend", "betriebEingestelltDatum": "",
        "maAnzahl": "", "maKurzarbeit": "",
        "svRueckstaende": False, "svBetrag": "", "svZeitraum": "",
        "lohnRueckstaende": False, "lohnBetrag": "", "lohnZeitraum": "",
        "mietePacht": "miete",
        "vermieterName": "", "vermieterAnschrift": "", "mieteBetrag": "",
        "kuendigungStatus": "nicht_gekuendigt", "kuendigungDatum": "",
        "mietRueckstaendeBetrag": "", "vermieterpfandrecht": False,
        "glaeubiger": [],
        "vermoegen": {
            k: {"vorhanden": False, "betrag": ""}
            for k in ["kasse","betriebsmittel","auftraege","aussenstaende","beteiligungen","grundstuecke","sonstiges"]
        },
        "_status": "neu", "notizen": "",
    }

    grund_flags  = {"zahlungsunfaehigkeit": False, "drohend": False, "ueberschuldung": False}
    gl_raw: dict[int, dict] = {}  # idx → {name, grund, haupt, zinsen, kosten}

    for field_name, value in raw.items():
        if not value or str(value).strip() in ("", "/Off", "Off"):
            continue

        value_str = str(value).strip()
        fn_lower  = field_name.lower()

        # Direkte Zuordnung via FIELD_MAP
        matched = False
        for pattern, json_key in FIELD_MAP.items():
            if pattern in fn_lower:
                if json_key.startswith("_grund_"):
                    grund = json_key[len("_grund_"):]
                    if value_str.lower() not in ("nein", "no", "false", "off"):
                        grund_flags[grund] = True
                elif json_key.startswith("_verm_"):
                    verm_key = json_key[len("_verm_"):]
                    try:
                        betrag = float(value_str.replace(",",".").replace("€","").strip())
                        data["vermoegen"][verm_key] = {"vorhanden": True, "betrag": str(betrag)}
                    except ValueError:
                        pass
                else:
                    if json_key in data and not data[json_key]:
                        data[json_key] = value_str
                matched = True
                break

        if matched:
            continue

        # Gläubiger-Felder (Anlage 1A)
        # Versuche Zeilenindex aus Feldnamen zu extrahieren (Row1..Row10 oder _2.._10)
        row_idx = None
        for pat in ["row", "Row"]:
            if pat in field_name:
                try:
                    part = field_name.split(pat)[-1]
                    row_idx = int("".join(c for c in part if c.isdigit())) - 1
                except (ValueError, IndexError):
                    pass
        if row_idx is None and "_" in field_name:
            try:
                row_idx = int(field_name.split("_")[-1]) - 1
            except ValueError:
                pass

        if row_idx is not None and 0 <= row_idx < 10:
            gl = gl_raw.setdefault(row_idx, {
                "name":"","grund":"","anschrift":"","plz":"","ort":"",
                "hauptforderung":0,"zinsen":0,"kosten":0,"vertreter":""
            })
            if any(p in fn_lower for p in GL_NAME_PATTERNS):
                gl["name"] = value_str
            elif any(p in fn_lower for p in GL_GRUND_PATTERNS):
                gl["grund"] = value_str
            elif any(p in fn_lower for p in GL_HAUPT_PATTERNS):
                try: gl["hauptforderung"] = float(value_str.replace(",","."))
                except ValueError: pass
            elif any(p in fn_lower for p in GL_ZINS_PATTERNS):
                try: gl["zinsen"] = float(value_str.replace(",","."))
                except ValueError: pass
            elif any(p in fn_lower for p in GL_KOST_PATTERNS):
                try: gl["kosten"] = float(value_str.replace(",","."))
                except ValueError: pass

    # Insolvenzgründe
    data["insolvenzgruende"] = [k for k, v in grund_flags.items() if v]

    # Gläubiger sortiert
    for idx in sorted(gl_raw.keys()):
        g = gl_raw[idx]
        if g.get("name"):
            data["glaeubiger"].append(g)

    return data


# ── Hauptprogramm ──────────────────────────────────────────────────────────

def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)

    input_path  = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    if not input_path.exists():
        print(f"FEHLER: Eingabedatei nicht gefunden: {input_path}")
        sys.exit(1)

    ext = input_path.suffix.lower()

    if ext == ".pdf":
        print(f"Lese PDF-Formularfelder aus {input_path} …")
        raw = read_pdf_fields(str(input_path))
    elif ext in (".xfdf", ".xml"):
        print(f"Lese XFDF-Daten aus {input_path} …")
        raw = read_xfdf_fields(str(input_path))
    else:
        print(f"FEHLER: Unbekanntes Dateiformat '{ext}'. Erwartet: .pdf oder .xfdf")
        sys.exit(1)

    print(f"{len(raw)} Felder gefunden. Erstelle data.json …")
    data = map_fields_to_data(raw)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    gl_count = len(data["glaeubiger"])
    print(f"✓ Fertig! {gl_count} Gläubiger importiert → {output_path}")
    print(f"  Firma: {data.get('firmaName','–')}")
    print(f"  GF:    {data.get('gfName','–')}")
    print(f"  Insolvenzgründe: {', '.join(data.get('insolvenzgruende',[])  or ['–'])}")


if __name__ == "__main__":
    main()
