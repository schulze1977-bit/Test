# Insolvenz-Manager v2.0

Desktop-Anwendung für Insolvenzanträge beim Amtsgericht Charlottenburg.

## Schnellstart

**Doppelklick auf `Start-Insolvenz-Manager.bat`**

Der Browser öffnet sich automatisch unter `http://localhost:5000`.

---

## Dateien

| Datei | Funktion |
|---|---|
| `Start-Insolvenz-Manager.bat` | Windows-Starter (Doppelklick) |
| `insolvenz_server.py` | Lokaler HTTP-Server + REST-API |
| `insolvenz_db.py` | SQLite-Datenbankschicht |
| `Insolvenz-Manager.html` | Hauptanwendung (auch ohne Server per Doppelklick) |
| `Mandanten-Fragebogen.html` | Mobiler Mandantenfragebogen (per E-Mail versendbar) |
| `fill_insolvenz.py` | PDF-Formular des AG Charlottenburg befüllen |
| `serienbriefe.py` | Word-Serienbriefe für alle Gläubiger |
| `import_adobe_pdf.py` | Adobe-PDF / XFDF in data.json umwandeln |
| `insolvenz.db` | SQLite-Datenbank (wird automatisch erstellt) |

---

## Python-Skripte

```bat
# PDF-Formular befüllen
python fill_insolvenz.py originalformular.pdf data.json ausgefuelltes_formular.pdf

# Serienbriefe erstellen
python serienbriefe.py data.json

# Adobe-PDF importieren
python import_adobe_pdf.py ausgefuelltes_adobe_formular.pdf data.json
python import_adobe_pdf.py formular.xfdf data.json
```

## Abhängigkeiten

```bat
pip install pypdf python-docx
```

## Mandanten-Fragebogen

Die Datei `Mandanten-Fragebogen.html` an Mandanten schicken oder per Link teilen.
Am Ende öffnet sich automatisch das E-Mail-Programm mit den kodierten Daten.

Im Insolvenz-Manager: **„Fragebogen importieren"** → Text aus E-Mail einfügen.
