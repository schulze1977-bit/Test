#!/usr/bin/env python3
"""
Insolvenz-Manager – Lokaler Server
Startet per Doppelklick oder via Start-Insolvenz-Manager.bat.
Kein externer Server, kein Internet erforderlich.

Endpunkte:
  GET    /                       → Insolvenz-Manager.html
  GET    /fragebogen              → Mandanten-Fragebogen.html
  GET    /api/health              → Statusprüfung
  GET    /api/cases               → Alle Fälle (Liste)
  POST   /api/cases               → Neuer Fall
  GET    /api/cases/{id}          → Fall abrufen
  PUT    /api/cases/{id}          → Fall aktualisieren
  DELETE /api/cases/{id}          → Fall löschen
  GET    /api/cases/{id}/versions → Versionsverlauf
  POST   /api/cases/{id}/restore/{vid} → Version wiederherstellen
  GET    /api/cases/{id}/export   → data.json herunterladen
"""

import http.server
import json
import os
import re
import sys
import tempfile
import threading
import webbrowser
from urllib.parse import urlparse

PORT = 5000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, BASE_DIR)
import insolvenz_db as db
import import_adobe_pdf as pdf_import

MIME = {
    ".html": "text/html; charset=utf-8",
    ".css":  "text/css; charset=utf-8",
    ".js":   "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".pdf":  "application/pdf",
    ".png":  "image/png",
    ".ico":  "image/x-icon",
}


class Handler(http.server.BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors(200)

    def do_GET(self):
        path = urlparse(self.path).path.rstrip("/") or "/"
        if path == "/":
            self._serve_html("Insolvenz-Manager.html")
        elif path == "/fragebogen":
            self._serve_html("Mandanten-Fragebogen.html")
        elif path.startswith("/api/"):
            self._api_get(path)
        else:
            self._serve_static(path)

    def do_POST(self):
        path = urlparse(self.path).path
        self._api_post(path)

    def do_PUT(self):
        path = urlparse(self.path).path
        self._api_put(path)

    def do_DELETE(self):
        path = urlparse(self.path).path
        self._api_delete(path)

    # ── API GET ───────────────────────────────────────────────────────────────

    def _api_get(self, path):
        if path == "/api/health":
            self._json({"status": "ok", "version": "2.0"})

        elif path == "/api/cases":
            self._json(db.list_cases())

        elif m := re.fullmatch(r"/api/cases/(\d+)", path):
            case = db.get_case(int(m.group(1)))
            self._json(case) if case else self._err(404, "Fall nicht gefunden")

        elif m := re.fullmatch(r"/api/cases/(\d+)/versions", path):
            self._json(db.get_versions(int(m.group(1))))

        elif m := re.fullmatch(r"/api/cases/(\d+)/antrag-pdf", path):
            case = db.get_case(int(m.group(1)))
            if not case:
                self._err(404, "Fall nicht gefunden")
                return
            self._serve_antrag_pdf(case)

        elif m := re.fullmatch(r"/api/cases/(\d+)/export", path):
            case = db.get_case(int(m.group(1)))
            if not case:
                self._err(404, "Fall nicht gefunden")
                return
            body = json.dumps(case["daten"], ensure_ascii=False, indent=2).encode("utf-8")
            firma = case["firma_name"].replace(" ", "_")[:40]
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Disposition",
                             f'attachment; filename="data_{firma}.json"')
            self.send_header("Content-Length", len(body))
            self._add_cors()
            self.end_headers()
            self.wfile.write(body)

        else:
            self._err(404, "Unbekannter Endpunkt")

    # ── API POST ──────────────────────────────────────────────────────────────

    def _api_post(self, path):
        if path == "/api/cases":
            daten = self._body_json()
            case = db.create_case(daten)
            self._json(case, 201)

        elif m := re.fullmatch(r"/api/cases/(\d+)/restore/(\d+)", path):
            result = db.restore_version(int(m.group(2)), int(m.group(1)))
            self._json(result) if result else self._err(404, "Version nicht gefunden")

        elif path == "/api/import-pdf":
            self._api_import_pdf()

        else:
            self._err(404, "Unbekannter Endpunkt")

    # ── API PUT ───────────────────────────────────────────────────────────────

    def _api_put(self, path):
        if m := re.fullmatch(r"/api/cases/(\d+)", path):
            daten = self._body_json()
            case = db.update_case(int(m.group(1)), daten)
            self._json(case) if case else self._err(404, "Fall nicht gefunden")
        else:
            self._err(404, "Unbekannter Endpunkt")

    # ── API DELETE ────────────────────────────────────────────────────────────

    def _api_delete(self, path):
        if m := re.fullmatch(r"/api/cases/(\d+)", path):
            db.delete_case(int(m.group(1)))
            self._json({"ok": True})
        else:
            self._err(404, "Unbekannter Endpunkt")

    # ── Amtlicher Insolvenzantrag (PDF befüllen) ──────────────────────────────

    def _serve_antrag_pdf(self, case: dict):
        template = os.path.join(BASE_DIR, "antrag_vorlage_jp.pdf")
        if not os.path.exists(template):
            self._err(404, "Vorlage antrag_vorlage_jp.pdf nicht gefunden")
            return
        import fill_antrag_jp as filler
        firma = case.get("firma_name", "Insolvenzantrag").replace(" ", "_")[:40]
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp_path = tmp.name
            daten = case.get("daten", {})
            if isinstance(daten, str):
                import json as _json
                daten = _json.loads(daten)
            filled = filler.fill_pdf_antrag(daten, template, tmp_path)
            with open(tmp_path, "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition",
                             f'attachment; filename="Insolvenzantrag_{firma}.pdf"')
            self.send_header("Content-Length", len(body))
            self._add_cors()
            self.end_headers()
            self.wfile.write(body)
        except Exception as exc:
            self._err(500, f"PDF-Erstellung fehlgeschlagen: {exc}")
        finally:
            if tmp_path:
                try: os.unlink(tmp_path)
                except OSError: pass

    # ── PDF-Import ────────────────────────────────────────────────────────────

    def _api_import_pdf(self):
        length = int(self.headers.get("Content-Length", 0))
        if length == 0:
            self._err(400, "Keine PDF-Daten empfangen")
            return
        pdf_bytes = self.rfile.read(length)
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(pdf_bytes)
                tmp_path = tmp.name
            raw = pdf_import.read_pdf_fields(tmp_path)
            if not raw:
                self._err(422, "Keine Formularfelder im PDF gefunden – "
                               "bitte nur ausfüllbare Adobe-PDFs verwenden")
                return
            data = pdf_import.map_fields_to_data(raw)
            gl_count = len(data.get("glaeubiger", []))
            data["_import_info"] = {
                "felder": len(raw),
                "glaeubiger": gl_count,
                "quelle": "pdf",
            }
            self._json(data)
        except Exception as exc:
            self._err(500, f"PDF-Analyse fehlgeschlagen: {exc}")
        finally:
            if tmp_path:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

    # ── Hilfsfunktionen ───────────────────────────────────────────────────────

    def _serve_html(self, filename):
        filepath = os.path.join(BASE_DIR, filename)
        if not os.path.exists(filepath):
            self._err(404, f"{filename} nicht gefunden")
            return
        with open(filepath, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self._add_cors()
        self.end_headers()
        self.wfile.write(body)

    def _serve_static(self, path):
        filepath = os.path.join(BASE_DIR, path.lstrip("/"))
        if not os.path.exists(filepath) or not os.path.isfile(filepath):
            self._err(404, "Datei nicht gefunden")
            return
        ext = os.path.splitext(filepath)[1].lower()
        ctype = MIME.get(ext, "application/octet-stream")
        with open(filepath, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self._add_cors()
        self.end_headers()
        self.wfile.write(body)

    def _err(self, code, msg):
        self._json({"fehler": msg}, code)

    def _cors(self, status=200):
        self.send_response(status)
        self._add_cors()
        self.end_headers()

    def _add_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _body_json(self) -> dict:
        length = int(self.headers.get("Content-Length", 0))
        if length == 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def log_message(self, fmt, *args):
        # Nur API-Aufrufe loggen, statische Dateien unterdrücken
        if "/api/" in args[0] if args else False:
            super().log_message(fmt, *args)


def _open_browser():
    import time
    time.sleep(1.2)
    webbrowser.open(f"http://localhost:{PORT}")


if __name__ == "__main__":
    print("=" * 60)
    print("  Insolvenz-Manager v2.0  –  Lokaler Server")
    print(f"  http://localhost:{PORT}")
    print("  Strg+C zum Beenden")
    print("=" * 60)

    threading.Thread(target=_open_browser, daemon=True).start()

    server = http.server.HTTPServer(("localhost", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer beendet.")
