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
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


# ── Mapping Adobe-Feldname → data.json-Schlüssel ───────────────────────────
#
# Format: "Adobe-Feldname (oder Teilstring)": "json-Schlüssel"
# Teilstrings werden case-insensitiv verglichen.

FIELD_MAP = {
    # ── Mandantenfragebogen-Felder (a_ bis k_) ─────────────────────────────
    # A – Firmendaten
    "a_firma":          "firmaName",
    "a_sitz":           "ort",
    "a_anschrift":      "_a_anschrift_komplett",
    "a_tel_mobil":      "telefon",
    "a_tel_fest":       "telefonFest",
    "a_fax":            "fax",
    "a_email":          "email",
    "a_branche":        "branche",
    "a_hr_amt":         "hrGericht",
    "a_hr_hrb_nr":      "hrNummer",
    "a_hr_hra_nr":      "hrNummer",
    # B – Geschäftsführer
    "b_gf_name":        "gfName",
    "b_gf_anschr":      "gfAnschrift",
    "b_gf_tel":         "gfTelefon",
    # C – Stammkapital
    "c_gruender":       "gesellschafterInfo",
    "c_sk_betrag":      "stammkapital",
    # D – Insolvenzgründe (Checkboxen)
    "d_zahlungsunf":    "_grund_zahlungsunfaehigkeit",
    "d_ueberschuld":    "_grund_ueberschuldung",
    "d_drohend":        "_grund_drohend",
    # E – Berater
    "e_anw_name":       "anwalt",
    "e_stb_name":       "steuerberater",
    # F – Betrieb
    "f_aktiv_nein":     "_flag_betrieb_eingestellt",
    "f_eingest_seit":   "betriebEingestelltDatum",
    # G – Mitarbeiter
    "g_an_anzahl":      "maAnzahl",
    "g_azubi_anzahl":   "azubiAnzahl",
    "g_kuend_dat":      "kuendigungDatum",
    "g_kuend_namen":    "kuendigungNamen",
    "g_mschutz_info":   "mutterschutzInfo",
    "g_sv_rueck":       "_flag_sv_rueck",
    "g_sv_betrag":      "_sv_betrag",
    "g_sv_traeger":     "svTraeger",
    "g_lohn_rueck":     "_flag_lohn_rueck",
    "g_lohn_betrag":    "_lohn_betrag",
    "g_lohn_namen":     "lohnRueckstandNamen",
    # H – Geschäftsräume
    "h_raum_adr":       "_h_raum_adr",
    "h_miete":          "mieteBetrag",
    "h_vermieter":      "vermieterName",
    "h_vertrag_dat":    "mietVertragDatum",
    "h_mietr_betrag":   "mietRueckstaendeBetrag",
    # I1 – Kasse/Bank
    "i1_betrag":        "_verm_kasse",
    "i1_detail":        "_verm_kasse_detail",
    # I2 – Inventar/Ausstattung
    "i2_buero_wert":    "_verm_buero_wert",
    "i2_laden_wert":    "_verm_laden_wert",
    "i2_masch_wert":    "_verm_maschinen_wert",
    "i2_waren_wert":    "_verm_waren_wert",
    "i2_laden_sich_info":  "_verm_sicherung_laden",
    "i2_masch_sich_info":  "_verm_sicherung_maschinen",
    # I3 – Fahrzeuge
    "i3_fz1_marke":     "_fz1_marke",
    "i3_fz1_bj":        "_fz1_bj",
    "i3_fz1_kz":        "_fz1_kz",
    "i3_fz1_wert":      "_fz1_wert",
    "i3_fz1_bel":       "_fz1_belastung",
    "i3_fz2_marke":     "_fz2_marke",
    "i3_fz2_bj":        "_fz2_bj",
    "i3_fz2_kz":        "_fz2_kz",
    "i3_fz2_wert":      "_fz2_wert",
    "i3_fz2_bel":       "_fz2_belastung",
    "i3_fz3_marke":     "_fz3_marke",
    "i3_fz3_wert":      "_fz3_wert",
    # I4 – Umsatz/Aufträge
    "i4_auft_detail":   "auftraegeDetail",
    "i4_ums_akt":       "umsatzAktuell",
    "i4_gew_akt":       "gewinnAktuell",
    "i4_ums_vor":       "umsatzVorjahr",
    "i4_gew_vor":       "gewinnVorjahr",
    # I5 – Außenstände
    "i5_detail":        "_verm_aussenstaende_detail",
    # I6 – Beteiligungen
    "i6_detail":        "_verm_beteiligungen_detail",
    # I7 – Immobilien
    "i7_im1_adr":       "_immo1_adr",
    "i7_im1_anteil":    "_immo1_anteil",
    "i7_im1_wert":      "_verm_grundstuecke",
    "i7_im1_gb":        "_immo1_grundbuch",
    "i7_im1_bel_info":  "_immo1_belastung",
    "i7_im2_adr":       "_immo2_adr",
    "i7_im2_wert":      "_immo2_wert",
    # I8 – Sonstiges
    "i8_immat_detail":  "_verm_immateriell_detail",
    "i8_sonst_detail":  "_verm_sonstiges_detail",
    # J – Gläubiger (Freitext-Felder)
    "j_finanzamt":      "_gl_finanzamt_info",
    "j_svtraeger":      "_gl_sv_info",
    "j_banken":         "_gl_banken_info",
    "j_gl_liste":       "_gl_sonstige_liste",
    # K – Ort/Datum
    "k_ort_dat":        "_ort_datum",

    # ── AG-Charlottenburg-Formular (generische Teilstring-Zuordnung) ────────
    "firma":                    "firmaName",
    "firmenname":               "firmaName",
    "rechtsform":               "rechtsform",
    "straße":                   "strasse",
    "strasse":                  "strasse",
    "hausnummer":               "hausnummer",
    "plz":                      "plz",
    "postleitzahl":             "plz",
    "telefon":                  "telefon",
    "e-mail":                   "email",
    "amtsgericht":              "hrGericht",
    "hr-nummer":                "hrNummer",
    "handelsregisternummer":    "hrNummer",
    "stammkapital":             "stammkapital",
    "geschäftsführer":          "gfName",
    "vertretungsberechtigte":   "gfName",
    "gf-name":                  "gfName",
    "funktion":                 "gfFunktion",
    "zahlungsunfähigkeit":      "_grund_zahlungsunfaehigkeit",
    "drohende":                 "_grund_drohend",
    "überschuldung":            "_grund_ueberschuldung",
    "branche":                  "branche",
    "wirtschaftszweig":         "branche",
    "arbeitnehmer":             "maAnzahl",
    "mitarbeiter":              "maAnzahl",
    "vermieter":                "vermieterName",
    "mietrückstand":            "mietRueckstaendeBetrag",
    "mietrückstände":           "mietRueckstaendeBetrag",
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

def _is_checked(value_str: str) -> bool:
    return value_str.lower() not in ("", "/off", "off", "nein", "no", "false")


def _parse_eur(value_str: str) -> float:
    try:
        return float(value_str.replace(".", "").replace(",", ".").replace("€", "").strip())
    except ValueError:
        return 0.0


def map_fields_to_data(raw: dict) -> dict:
    """Wandelt die rohen Formularfelder in das data.json-Schema um."""
    data = {
        "firmaName": "", "rechtsform": "GmbH",
        "strasse": "", "hausnummer": "", "plz": "", "ort": "",
        "telefon": "", "telefonFest": "", "fax": "", "email": "",
        "aktenzeichen": "",
        "hrGericht": "", "hrNummer": "", "stammkapital": "",
        "gfName": "", "gfFunktion": "Geschäftsführer", "gfAnschrift": "",
        "gfTelefon": "", "gesellschafterInfo": "",
        "insolvenzgruende": [],
        "branche": "", "anwalt": "", "steuerberater": "", "hinweise": "",
        "betriebStatus": "laufend", "betriebEingestelltDatum": "",
        "maAnzahl": "", "azubiAnzahl": "",
        "svRueckstaende": False, "svBetrag": "", "svTraeger": "",
        "lohnRueckstaende": False, "lohnBetrag": "", "lohnRueckstandNamen": "",
        "kuendigungDatum": "", "kuendigungNamen": "", "mutterschutzInfo": "",
        "mietePacht": "miete",
        "vermieterName": "", "mieteBetrag": "", "mietVertragDatum": "",
        "mietRueckstaendeBetrag": "", "vermieterpfandrecht": False,
        "umsatzAktuell": "", "gewinnAktuell": "", "umsatzVorjahr": "", "gewinnVorjahr": "",
        "auftraegeDetail": "",
        "fahrzeuge": [],
        "glaeubiger": [],
        "vermoegen": {
            k: {"vorhanden": False, "betrag": ""}
            for k in ["kasse","betriebsmittel","auftraege","aussenstaende","beteiligungen","grundstuecke","sonstiges"]
        },
        "_status": "neu", "notizen": "",
    }

    grund_flags = {"zahlungsunfaehigkeit": False, "drohend": False, "ueberschuldung": False}
    gl_raw: dict[int, dict] = {}
    rcs_gl_raw: dict[int, dict] = {}
    _extras: dict = {}  # temporärer Speicher für _-Schlüssel

    # ── Checkbox-Flags für Betrieb/Mitarbeiter ──────────────────────────────
    betrieb_eingestellt = False
    sv_rueck = False
    lohn_rueck = False

    for field_name, value in raw.items():
        if not value or str(value).strip() in ("", "/Off", "Off"):
            continue

        value_str = str(value).strip()
        fn_lower  = field_name.lower()

        # ── Mandantenfragebogen: exakte Feldnamen (höchste Priorität) ───────
        if field_name in FIELD_MAP:
            json_key = FIELD_MAP[field_name]
            if json_key.startswith("_flag_"):
                if _is_checked(value_str):
                    flag = json_key[len("_flag_"):]
                    if flag == "betrieb_eingestellt":
                        betrieb_eingestellt = True
                    elif flag == "sv_rueck":
                        sv_rueck = True
                    elif flag == "lohn_rueck":
                        lohn_rueck = True
            elif json_key.startswith("_grund_"):
                if _is_checked(value_str):
                    grund_flags[json_key[len("_grund_"):]] = True
            elif json_key.startswith("_verm_"):
                verm_key = json_key[len("_verm_"):]
                if verm_key in data["vermoegen"]:
                    betrag = _parse_eur(value_str)
                    if betrag:
                        data["vermoegen"][verm_key] = {"vorhanden": True, "betrag": str(betrag)}
                else:
                    _extras[json_key] = value_str
            elif json_key.startswith("_"):
                _extras[json_key] = value_str
            elif json_key in data and not data[json_key]:
                data[json_key] = value_str
            continue

        # ── Teilstring-Zuordnung (AG-Charlottenburg + Fallback) ──────────────
        matched = False
        for pattern, json_key in FIELD_MAP.items():
            if pattern in fn_lower:
                if json_key.startswith("_grund_"):
                    if _is_checked(value_str):
                        grund_flags[json_key[len("_grund_"):]] = True
                elif json_key.startswith("_verm_"):
                    verm_key = json_key[len("_verm_"):]
                    if verm_key in data["vermoegen"]:
                        betrag = _parse_eur(value_str)
                        if betrag:
                            data["vermoegen"][verm_key] = {"vorhanden": True, "betrag": str(betrag)}
                elif json_key.startswith("_"):
                    _extras.setdefault(json_key, value_str)
                else:
                    if json_key in data and not data[json_key]:
                        data[json_key] = value_str
                matched = True
                break

        if matched:
            continue

        # ── GL1-GL15: RCS-Fragebogen Gläubigerstruktur ──────────────────────
        m_gl = re.fullmatch(r"GL(\d+)_(.+)", field_name, re.IGNORECASE)
        if m_gl:
            gl_n = int(m_gl.group(1))
            gl_sub = m_gl.group(2).lower()
            if 1 <= gl_n <= 15:
                rcs = rcs_gl_raw.setdefault(gl_n, {})
                if gl_sub == "name":
                    rcs["name"] = value_str
                elif gl_sub == "grund":
                    rcs["grund"] = value_str
                elif gl_sub == "haupt":
                    rcs["hauptforderung"] = _parse_eur(value_str)
                elif gl_sub == "zinsen":
                    rcs["zinsen"] = _parse_eur(value_str)
                elif gl_sub == "zinsen_bis":
                    rcs["zinsenBis"] = value_str
                elif gl_sub == "kosten":
                    rcs["kosten"] = _parse_eur(value_str)
                elif gl_sub == "tituliert":
                    rcs["forderungTituliert"] = "ja" if _is_checked(value_str) else "nein"
                elif gl_sub == "sonder":
                    rcs["forderungGesichert"] = "ja" if _is_checked(value_str) else "nein"
                elif gl_sub == "nahestehend":
                    rcs["nahestehendeP138"] = "ja" if _is_checked(value_str) else "nein"
            continue

        # ── Gläubiger-Felder (Anlage 1A im AG-Formular) ─────────────────────
        row_idx = None
        for pat in ["row", "Row"]:
            if pat in field_name:
                try:
                    part = field_name.split(pat)[-1]
                    row_idx = int("".join(c for c in part if c.isdigit())) - 1
                except (ValueError, IndexError):
                    pass
        if row_idx is not None and 0 <= row_idx < 10:
            gl = gl_raw.setdefault(row_idx, {
                "name": "", "grund": "", "anschrift": "", "plz": "", "ort": "",
                "hauptforderung": 0, "zinsen": 0, "kosten": 0, "vertreter": ""
            })
            if any(p in fn_lower for p in GL_NAME_PATTERNS):
                gl["name"] = value_str
            elif any(p in fn_lower for p in GL_GRUND_PATTERNS):
                gl["grund"] = value_str
            elif any(p in fn_lower for p in GL_HAUPT_PATTERNS):
                gl["hauptforderung"] = _parse_eur(value_str)
            elif any(p in fn_lower for p in GL_ZINS_PATTERNS):
                gl["zinsen"] = _parse_eur(value_str)
            elif any(p in fn_lower for p in GL_KOST_PATTERNS):
                gl["kosten"] = _parse_eur(value_str)

    # ── Nachbearbeitung ─────────────────────────────────────────────────────

    data["insolvenzgruende"] = [k for k, v in grund_flags.items() if v]

    if betrieb_eingestellt:
        data["betriebStatus"] = "eingestellt"
    if sv_rueck:
        data["svRueckstaende"] = True
        if _extras.get("_sv_betrag"):
            data["svBetrag"] = _extras["_sv_betrag"]
    if lohn_rueck:
        data["lohnRueckstaende"] = True
        if _extras.get("_lohn_betrag"):
            data["lohnBetrag"] = _extras["_lohn_betrag"]

    # Anschrift aus a_anschrift parsen (z.B. "Straße 42, 10719 Berlin")
    if _extras.get("_a_anschrift_komplett") and not data["strasse"]:
        raw_adr = _extras["_a_anschrift_komplett"]
        parts = [p.strip() for p in raw_adr.split(",")]
        if parts:
            strteil = parts[0].rsplit(" ", 1)
            data["strasse"] = strteil[0] if len(strteil) > 1 else parts[0]
            data["hausnummer"] = strteil[1] if len(strteil) > 1 else ""
        if len(parts) > 1:
            plz_ort = parts[1].strip().split(" ", 1)
            data["plz"] = plz_ort[0] if plz_ort else ""
            data["ort"] = plz_ort[1] if len(plz_ort) > 1 else data.get("ort", "")

    # Geschäftsraumadresse
    if _extras.get("_h_raum_adr"):
        data["geschaeftsraumAdr"] = _extras["_h_raum_adr"]

    # Fahrzeuge aus i3_fz1_*/i3_fz2_*/i3_fz3_* zusammenführen
    for n in ("1", "2", "3"):
        marke = _extras.get(f"_fz{n}_marke", "")
        wert  = _extras.get(f"_fz{n}_wert", "")
        if marke or wert:
            data["fahrzeuge"].append({
                "marke":     marke,
                "baujahr":   _extras.get(f"_fz{n}_bj", ""),
                "kennzeichen": _extras.get(f"_fz{n}_kz", ""),
                "wert":      wert,
                "belastung": _extras.get(f"_fz{n}_belastung", ""),
            })
            # Fahrzeugwert zum Vermögen "betriebsmittel" addieren
            betrag_fz = _parse_eur(wert)
            if betrag_fz:
                bm = data["vermoegen"]["betriebsmittel"]
                bm_betrag = float(bm["betrag"]) if bm["betrag"] else 0
                bm_betrag += betrag_fz
                data["vermoegen"]["betriebsmittel"] = {"vorhanden": True, "betrag": str(bm_betrag)}

    # Inventarwerte (Büro + Laden + Maschinen + Waren) → betriebsmittel
    for xkey in ("_verm_buero_wert", "_verm_laden_wert", "_verm_maschinen_wert"):
        betrag = _parse_eur(_extras.get(xkey, ""))
        if betrag:
            bm = data["vermoegen"]["betriebsmittel"]
            bm_betrag = float(bm["betrag"]) if bm["betrag"] else 0
            data["vermoegen"]["betriebsmittel"] = {"vorhanden": True, "betrag": str(bm_betrag + betrag)}

    # Warenbestand → auftraege (Vorräte)
    waren = _parse_eur(_extras.get("_verm_waren_wert", ""))
    if waren:
        data["vermoegen"]["auftraege"] = {"vorhanden": True, "betrag": str(waren)}

    # Außenstände aus Detail-Text (i5_detail enthält oft EUR-Betrag)
    if _extras.get("_verm_aussenstaende_detail"):
        detail = _extras["_verm_aussenstaende_detail"]
        m = re.search(r"([\d.,]+)\s*(?:EUR|€)", detail)
        if m and not data["vermoegen"]["aussenstaende"]["betrag"]:
            betrag = _parse_eur(m.group(1))
            if betrag:
                data["vermoegen"]["aussenstaende"] = {"vorhanden": True, "betrag": str(betrag)}

    # Notizen aus Gläubiger-Freitextfeldern
    gl_info_parts = []
    for k in ("_gl_finanzamt_info", "_gl_sv_info", "_gl_banken_info", "_gl_sonstige_liste"):
        if _extras.get(k):
            gl_info_parts.append(_extras[k])
    if gl_info_parts:
        data["notizen"] = "Gläubigerangaben aus Fragebogen:\n" + "\n".join(gl_info_parts)

    # Gläubiger aus Anlage 1A (AG-Formular)
    for idx in sorted(gl_raw.keys()):
        g = gl_raw[idx]
        if g.get("name"):
            data["glaeubiger"].append(g)

    # Gläubiger aus GL1-GL15 (RCS-Mandantenfragebogen)
    for idx in sorted(rcs_gl_raw.keys()):
        g = rcs_gl_raw[idx]
        if g.get("name"):
            entry = {
                "name":              g.get("name", ""),
                "grund":             g.get("grund", ""),
                "anschrift":         "",
                "plz":               "",
                "ort":               "",
                "hauptforderung":    g.get("hauptforderung", 0),
                "zinsen":            g.get("zinsen", 0),
                "zinsenBis":         g.get("zinsenBis", ""),
                "kosten":            g.get("kosten", 0),
                "vertreter":         "",
                "forderungTituliert": g.get("forderungTituliert", "nein"),
                "forderungGesichert": g.get("forderungGesichert", "nein"),
                "nahestehendeP138":   g.get("nahestehendeP138", "nein"),
                "kategorie":          "",
            }
            data["glaeubiger"].append(entry)

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
