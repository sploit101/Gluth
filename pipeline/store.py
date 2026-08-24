#!/usr/bin/env python3
"""
Persistent state for the workspace: inquiries, drafts, corrections, triage.

SQLite, single file, no server. Seeded from the real corpus — the projects in
the queue are the projects on disk, not fixtures.
"""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parent / "out"
DB = OUT / "workspace.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS inquiry (
    project_no   TEXT PRIMARY KEY,
    project_name TEXT,
    customer     TEXT,
    stage        TEXT,
    status       TEXT,
    confidence   TEXT,
    assigned_to  TEXT,
    doc_count    INTEGER DEFAULT 0,
    chunk_count  INTEGER DEFAULT 0,
    updated_at   TEXT
);
CREATE TABLE IF NOT EXISTS correction (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    project_no  TEXT,
    artifact    TEXT,
    field       TEXT,
    action      TEXT,
    route_tag   TEXT,
    reviewer    TEXT,
    created_at  TEXT
);
CREATE TABLE IF NOT EXISTS decision (
    project_no  TEXT PRIMARY KEY,
    verdict     TEXT,
    criteria    TEXT,
    assigned_to TEXT,
    note        TEXT,
    created_at  TEXT
);
CREATE TABLE IF NOT EXISTS review (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    project_no  TEXT,
    artifact    TEXT,
    outcome     TEXT,          -- approved | corrected
    reviewer    TEXT,
    created_at  TEXT
);
"""

ROUTING = {
    "BMW": "M. Weber",
    "Bosch": "K. Schmidt",
    "Dräxlmaier": "H. Becker",
    "BMG / ArGe": "H. Becker",
}


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def connect() -> sqlite3.Connection:
    OUT.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    return conn


def derive_stage(phases: set[str]) -> tuple[str, str]:
    """
    Map the folders actually present to a workflow stage.

    Only 2_/3_/7_/8_Nachträge are workflow steps. 4_Kalkulation, 5_Zukaufteile
    and 6_Korrespondenz are supporting folders that exist throughout, so they
    must not be read as progress.
    """
    has = lambda pref: any(p.startswith(pref) for p in phases)  # noqa: E731
    nachtrag = any(p.startswith("8_") and "Nachtr" in p for p in phases)

    if nachtrag:
        return "8_Nachträge", "Won"
    if has("7_"):
        return "7_Auftrag", "Won"
    if has("3_"):
        return "3_Angebot", "Awaiting review"
    if has("2_"):
        return "2_Anfrage", "New"
    return "2_Anfrage", "New"


def seed_from_corpus(conn: sqlite3.Connection, chunks: list[dict]) -> int:
    """Create one inquiry per project actually present in the corpus."""
    stats: dict[str, dict] = {}
    for c in chunks:
        pno = c.get("project_no")
        if not pno:
            continue
        s = stats.setdefault(pno, {
            "project_name": c.get("project_name", ""),
            "customer": c.get("customer", ""),
            "files": set(), "chunks": 0, "phases": set(),
        })
        s["files"].add(c["source_path"])
        s["chunks"] += 1
        if c.get("phase_no"):
            s["phases"].add(f'{c["phase_no"]}_{c["phase_name"]}')

    added = 0
    for pno, s in stats.items():
        existing = conn.execute(
            "SELECT 1 FROM inquiry WHERE project_no=?", (pno,)).fetchone()
        if existing:
            conn.execute(
                "UPDATE inquiry SET doc_count=?, chunk_count=? WHERE project_no=?",
                (len(s["files"]), s["chunks"], pno))
            continue
        stage, status = derive_stage(s["phases"])
        conn.execute(
            "INSERT INTO inquiry (project_no, project_name, customer, stage,"
            " status, confidence, assigned_to, doc_count, chunk_count, updated_at)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)",
            (pno, s["project_name"], s["customer"], stage, status, "",
             ROUTING.get(s["customer"], "Admin triage"),
             len(s["files"]), s["chunks"], now()))
        added += 1
    conn.commit()
    return added


# ------------------------------------------------------------------ queries
def list_inquiries(conn) -> list[dict]:
    rows = conn.execute("SELECT * FROM inquiry ORDER BY project_no").fetchall()
    out = []
    for r in rows:
        d = dict(r)
        dec = conn.execute(
            "SELECT verdict FROM decision WHERE project_no=?",
            (d["project_no"],)).fetchone()
        d["verdict"] = dec["verdict"] if dec else ""
        d["corrections"] = conn.execute(
            "SELECT COUNT(*) c FROM correction WHERE project_no=?",
            (d["project_no"],)).fetchone()["c"]
        out.append(d)
    return out


def add_correction(conn, **kw) -> int:
    cur = conn.execute(
        "INSERT INTO correction (project_no, artifact, field, action,"
        " route_tag, reviewer, created_at) VALUES (?,?,?,?,?,?,?)",
        (kw.get("project_no", ""), kw.get("artifact", ""), kw.get("field", ""),
         kw.get("action", ""), kw.get("route_tag", "Content"),
         kw.get("reviewer", "M. Weber"), now()))
    conn.execute(
        "INSERT INTO review (project_no, artifact, outcome, reviewer, created_at)"
        " VALUES (?,?,?,?,?)",
        (kw.get("project_no", ""), kw.get("artifact", ""), "corrected",
         kw.get("reviewer", "M. Weber"), now()))
    conn.commit()
    return cur.lastrowid or 0


def list_corrections(conn, route_tag: str | None = None) -> list[dict]:
    sql = "SELECT * FROM correction"
    args: tuple = ()
    if route_tag and route_tag != "ALL":
        sql += " WHERE route_tag=?"
        args = (route_tag,)
    sql += " ORDER BY id DESC"
    return [dict(r) for r in conn.execute(sql, args).fetchall()]


def approve(conn, project_no: str, artifact: str, reviewer: str = "M. Weber"):
    conn.execute(
        "INSERT INTO review (project_no, artifact, outcome, reviewer, created_at)"
        " VALUES (?,?,?,?,?)", (project_no, artifact, "approved", reviewer, now()))
    conn.execute("UPDATE inquiry SET status='Approved', updated_at=?"
                 " WHERE project_no=?", (now(), project_no))
    conn.commit()


def set_decision(conn, project_no: str, verdict: str, criteria: list,
                 assigned_to: str = "", note: str = ""):
    conn.execute(
        "INSERT INTO decision (project_no, verdict, criteria, assigned_to, note,"
        " created_at) VALUES (?,?,?,?,?,?)"
        " ON CONFLICT(project_no) DO UPDATE SET verdict=excluded.verdict,"
        " criteria=excluded.criteria, assigned_to=excluded.assigned_to,"
        " note=excluded.note, created_at=excluded.created_at",
        (project_no, verdict, json.dumps(criteria), assigned_to, note, now()))
    if assigned_to:
        conn.execute("UPDATE inquiry SET assigned_to=? WHERE project_no=?",
                     (assigned_to, project_no))
    conn.commit()


def review_stats(conn) -> dict:
    total = conn.execute("SELECT COUNT(*) c FROM review").fetchone()["c"]
    approved = conn.execute(
        "SELECT COUNT(*) c FROM review WHERE outcome='approved'").fetchone()["c"]
    corr = conn.execute("SELECT COUNT(*) c FROM correction").fetchone()["c"]
    drift = conn.execute(
        "SELECT COUNT(*) c FROM correction WHERE route_tag='Drift'").fetchone()["c"]
    top = conn.execute(
        "SELECT field, COUNT(*) c FROM correction GROUP BY field"
        " ORDER BY c DESC LIMIT 1").fetchone()
    return {
        "reviews": total, "approved": approved, "corrections": corr,
        "drift": drift,
        "top_field": top["field"] if top else "—",
        "top_field_count": top["c"] if top else 0,
    }
