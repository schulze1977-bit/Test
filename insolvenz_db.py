"""
Insolvenz-Manager – Datenbankschicht (SQLite)
Speichert alle Insolvenzfälle mit Versionsverlauf.
"""

import sqlite3
import json
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "insolvenz.db")


def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    with _conn() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS cases (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                firma_name   TEXT    NOT NULL DEFAULT '',
                rechtsform   TEXT    NOT NULL DEFAULT 'GmbH',
                status       TEXT    NOT NULL DEFAULT 'neu',
                aktenzeichen TEXT    NOT NULL DEFAULT '',
                erstellt     TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
                geaendert    TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
                daten        TEXT    NOT NULL DEFAULT '{}'
            );

            CREATE TABLE IF NOT EXISTS versionen (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                fall_id   INTEGER NOT NULL,
                gespeichert TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                daten     TEXT NOT NULL,
                FOREIGN KEY (fall_id) REFERENCES cases(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS dokumente (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                fall_id   INTEGER NOT NULL,
                typ       TEXT NOT NULL,
                dateiname TEXT NOT NULL,
                erstellt  TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                FOREIGN KEY (fall_id) REFERENCES cases(id) ON DELETE CASCADE
            );
        """)
        c.commit()


# ── CRUD Fälle ────────────────────────────────────────────────────────────────

def list_cases():
    with _conn() as c:
        rows = c.execute(
            "SELECT id, firma_name, rechtsform, status, aktenzeichen, erstellt, geaendert "
            "FROM cases ORDER BY geaendert DESC"
        ).fetchall()
        return [dict(r) for r in rows]


def get_case(case_id: int):
    with _conn() as c:
        row = c.execute("SELECT * FROM cases WHERE id=?", (case_id,)).fetchone()
        if not row:
            return None
        result = dict(row)
        result["daten"] = json.loads(result["daten"])
        return result


def create_case(daten: dict) -> dict:
    firma_name   = daten.get("firmaName", daten.get("firma_name", "Unbekannt"))
    rechtsform   = daten.get("rechtsform", "GmbH")
    status       = daten.get("_status", "neu")
    aktenzeichen = daten.get("aktenzeichen", "")
    with _conn() as c:
        cur = c.execute(
            "INSERT INTO cases (firma_name, rechtsform, status, aktenzeichen, daten) "
            "VALUES (?,?,?,?,?)",
            (firma_name, rechtsform, status, aktenzeichen,
             json.dumps(daten, ensure_ascii=False))
        )
        c.commit()
        return get_case(cur.lastrowid)


def update_case(case_id: int, daten: dict) -> dict | None:
    firma_name   = daten.get("firmaName", daten.get("firma_name", ""))
    rechtsform   = daten.get("rechtsform", "GmbH")
    status       = daten.get("_status", "neu")
    aktenzeichen = daten.get("aktenzeichen", "")
    jetzt        = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as c:
        old = c.execute("SELECT daten FROM cases WHERE id=?", (case_id,)).fetchone()
        if not old:
            return None
        # Snapshot für Versionsverlauf
        c.execute(
            "INSERT INTO versionen (fall_id, daten) VALUES (?,?)",
            (case_id, old["daten"])
        )
        c.execute(
            "UPDATE cases SET firma_name=?, rechtsform=?, status=?, aktenzeichen=?, "
            "geaendert=?, daten=? WHERE id=?",
            (firma_name, rechtsform, status, aktenzeichen,
             jetzt, json.dumps(daten, ensure_ascii=False), case_id)
        )
        c.commit()
    return get_case(case_id)


def delete_case(case_id: int) -> bool:
    with _conn() as c:
        c.execute("DELETE FROM cases WHERE id=?", (case_id,))
        c.commit()
    return True


# ── Versionsverlauf ───────────────────────────────────────────────────────────

def get_versions(case_id: int) -> list:
    with _conn() as c:
        rows = c.execute(
            "SELECT id, gespeichert FROM versionen WHERE fall_id=? ORDER BY gespeichert DESC",
            (case_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def restore_version(version_id: int, case_id: int) -> dict | None:
    with _conn() as c:
        row = c.execute(
            "SELECT daten FROM versionen WHERE id=? AND fall_id=?",
            (version_id, case_id)
        ).fetchone()
        if not row:
            return None
        daten = json.loads(row["daten"])
    return update_case(case_id, daten)


# ── Dokumente ─────────────────────────────────────────────────────────────────

def add_dokument(case_id: int, typ: str, dateiname: str) -> dict:
    with _conn() as c:
        cur = c.execute(
            "INSERT INTO dokumente (fall_id, typ, dateiname) VALUES (?,?,?)",
            (case_id, typ, dateiname)
        )
        c.commit()
        row = c.execute("SELECT * FROM dokumente WHERE id=?", (cur.lastrowid,)).fetchone()
        return dict(row)


def get_dokumente(case_id: int) -> list:
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM dokumente WHERE fall_id=? ORDER BY erstellt DESC",
            (case_id,)
        ).fetchall()
        return [dict(r) for r in rows]


# Datenbank beim Import initialisieren
init_db()
