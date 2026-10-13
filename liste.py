#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
liste.py — macht aus out/requirements.jsonl eine lesbare Excel-Mappe.

Das ist die Uebergabe an einen Menschen. Alles davor in der Kette erzeugt
technische Dateien; hier entsteht das, was ein Kollege oeffnen, filtern,
kommentieren und per Mail weitergeben kann.

Je Projekt ein Tabellenblatt, dazu ein Blatt "Uebersicht" mit den Zahlen.

Zwei Spalten sind wichtiger als die anderen:
  Pruefen   sicher  = klares Signalwort, vollstaendiger Satz
            pruefen = Verdachtsfall, offener Punkt, Satz ohne Signalwort
                      oder abgeschnitten. Hier zuerst hinsehen.
  Warum     welche Regel die Verbindlichkeit entschieden hat. Damit ist
            jede Einstufung nachvollziehbar und bestreitbar.

Kein Sprachmodell. Liest nur, aendert nichts am Bestand.

    python liste.py
    python liste.py --projekt A07616000
    python liste.py --datei LH_04K_000051
    python liste.py --nur-muss

Ausgabe
    out/anforderungen_zahlen.txt   nur Zahlen      -> darf geteilt werden
    out/anforderungen.xlsx         mit Wortlauten  -> VERTRAULICH
"""

from __future__ import annotations

import argparse
import collections
import json
import re
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

# Reihenfolge, in der die Verbindlichkeitsstufen sortiert werden:
# Pflichten zuerst, Hinweise zuletzt.
RANG = {"MUSS": 0, "VERBOT": 1, "SOLL": 2, "KANN": 3, "INFO": 4}

# Zeichen, die Excel in Blattnamen nicht erlaubt
VERBOTEN = re.compile(r"[\[\]:*?/\\]")


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def blattname(roh: str, belegt: set[str]) -> str:
    name = VERBOTEN.sub("-", norm(roh)) or "ohne Projekt"
    name = name[:31]
    i = 2
    while name.lower() in belegt:
        kuerzel = f" ({i})"
        name = name[:31 - len(kuerzel)] + kuerzel
        i += 1
    belegt.add(name.lower())
    return name


def wert_text(r: dict) -> str:
    if r.get("value") is None:
        return ""
    t = f'{r["value"]:g}'
    if r.get("value_max") is not None:
        t += f' bis {r["value_max"]:g}'
    return t


def lade_regel():
    """binding_rule aus requirements.py, wenn erreichbar — sonst ohne."""
    try:
        sys.path.insert(0, str(HERE))
        from requirements import binding_rule       # type: ignore
        return binding_rule
    except Exception as e:                          # pragma: no cover
        print("HINWEIS: requirements.py nicht ladbar, Spalte 'Warum' bleibt leer")
        print("        ", ascii_safe(e))
        return None


def bewerten(r: dict, regel_fn) -> tuple[str, str]:
    """Gibt (Pruefen, Warum) zurueck."""
    warum = ""
    if regel_fn:
        try:
            warum = norm(regel_fn(r.get("text", "")))
        except Exception:
            warum = ""
    gruende = []
    if r.get("suspect"):
        gruende.append(norm(r["suspect"]))
    if r.get("needs_question"):
        gruende.append("offener Punkt")
    if "kein Signalwort" in warum or "kein englisches Signalwort" in warum:
        gruende.append("kein Signalwort im Satz")
    if "Spezifikationszeile" in warum:
        gruende.append("Tabellenzeile ohne Verb")
    pruefen = ("pruefen: " + "; ".join(gruende)) if gruende else "sicher"
    return pruefen, warum


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--req", default=str(OUT / "requirements.jsonl"))
    ap.add_argument("--projekt", default=None, help="nur dieses Projekt")
    ap.add_argument("--datei", default=None,
                    help="nur Dateien, deren Name diesen Text enthaelt")
    ap.add_argument("--nur-muss", action="store_true",
                    help="nur MUSS und VERBOT")
    ap.add_argument("--aus", default=None, help="Zieldatei (.xlsx)")
    a = ap.parse_args()

    pfad = Path(a.req)
    if not pfad.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(pfad))
        print("Erst 'python requirements.py' laufen lassen.")
        return 2

    rows = []
    for line in pfad.open(encoding="utf-8"):
        r = json.loads(line)
        if a.projekt and r.get("project_no") != a.projekt:
            continue
        if a.datei and a.datei.lower() not in str(r.get("file", "")).lower():
            continue
        if a.nur_muss and r.get("binding") not in ("MUSS", "VERBOT"):
            continue
        rows.append(r)

    if not rows:
        print("Keine Anforderungen gefunden. Filter zu eng?")
        return 1

    regel_fn = lade_regel()

    # ------------------------------------------------------------- anreichern
    for r in rows:
        r["_pruefen"], r["_warum"] = bewerten(r, regel_fn)

    nach_projekt: dict[str, list[dict]] = collections.defaultdict(list)
    for r in rows:
        nach_projekt[r.get("project_no") or "ohne Projektnummer"].append(r)

    def sortschluessel(r: dict):
        return (str(r.get("file") or ""),
                RANG.get(r.get("binding", ""), 9),
                str(r.get("page") or "").zfill(4))

    # ---------------------------------------------------------------- Zahlen
    bind = collections.Counter(r.get("binding") for r in rows)
    fach = collections.Counter(r.get("category") for r in rows)
    pruef = collections.Counter("sicher" if r["_pruefen"] == "sicher" else "pruefen"
                                for r in rows)
    L = [f"ANFORDERUNGSLISTE  {datetime.now():%Y-%m-%d %H:%M}", "=" * 62, ""]
    if a.projekt:
        L.append(f"  nur Projekt            {a.projekt}")
    if a.datei:
        L.append(f"  nur Dateien mit        '{a.datei}'")
    if a.nur_muss:
        L.append("  nur MUSS und VERBOT")
    L += [f"  Anforderungen          {len(rows)}",
          f"  Projekte               {len(nach_projekt)}",
          f"  Dokumente              {len({r.get('file') for r in rows})}",
          f"  mit Grenzwert          {sum(1 for r in rows if r.get('value') is not None)}",
          "", "  Verbindlichkeit"]
    for k, n in sorted(bind.items(), key=lambda x: RANG.get(x[0], 9)):
        L.append(f"    {n:>6}  {k}")
    L += ["", "  Fach"]
    for k, n in fach.most_common():
        L.append(f"    {n:>6}  {k}")
    L += ["", "  Pruefhinweis"]
    for k, n in pruef.most_common():
        L.append(f"    {n:>6}  {k}   ({100*n/len(rows):.0f} %)")
    L += ["", "Diese Datei enthaelt keine Dateinamen und keine Saetze."]

    OUT.mkdir(exist_ok=True)
    (OUT / "anforderungen_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
    for z in L:
        print(ascii_safe(z))

    # ----------------------------------------------------------------- Excel
    try:
        import xlsxwriter
    except ImportError:
        print()
        print("HINWEIS: xlsxwriter fehlt, Excel wurde nicht geschrieben.")
        print("  .\\.venv\\Scripts\\python.exe -m pip install xlsxwriter")
        return 0

    ziel = Path(a.aus) if a.aus else OUT / "anforderungen.xlsx"
    wb = xlsxwriter.Workbook(str(ziel), {"constant_memory": False})
    f_kopf = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1,
                            "text_wrap": True, "valign": "top"})
    f_text = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    f_muss = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                            "bold": True})
    f_grau = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                            "font_color": "#777777"})
    f_t1 = wb.add_format({"bold": True, "font_size": 13})
    f_t2 = wb.add_format({"bold": True})

    # --- Blatt 1: Uebersicht
    ws = wb.add_worksheet("Uebersicht")
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 16)
    z = 0
    ws.write(z, 0, "Anforderungen aus dem Projektbestand", f_t1); z += 2
    ws.write(z, 0, "Stand", f_t2); ws.write(z, 1, f"{datetime.now():%d.%m.%Y %H:%M}"); z += 1
    ws.write(z, 0, "Anforderungen gesamt", f_t2); ws.write(z, 1, len(rows)); z += 1
    ws.write(z, 0, "Dokumente", f_t2)
    ws.write(z, 1, len({r.get("file") for r in rows})); z += 2
    ws.write(z, 0, "Verbindlichkeit", f_t2); z += 1
    for k, n in sorted(bind.items(), key=lambda x: RANG.get(x[0], 9)):
        ws.write(z, 0, k); ws.write(z, 1, n); z += 1
    z += 1
    ws.write(z, 0, "Fach", f_t2); z += 1
    for k, n in fach.most_common():
        ws.write(z, 0, k); ws.write(z, 1, n); z += 1
    z += 1
    ws.write(z, 0, "Pruefhinweis", f_t2); z += 1
    for k, n in pruef.most_common():
        ws.write(z, 0, k); ws.write(z, 1, n); z += 1
    z += 2
    for satz in [
        "Jede Zeile ist ein woertliches Zitat aus dem genannten Dokument.",
        "Es wurde nichts umformuliert und nichts ergaenzt.",
        "Spalte 'Pruefen': 'pruefen' heisst, hier zuerst nachsehen.",
        "Spalte 'Warum': welche Regel die Verbindlichkeitsstufe entschieden hat.",
        "Die Spalte 'Bemerkung' ist leer und fuer Ihre Anmerkungen gedacht.",
    ]:
        ws.write(z, 0, satz); z += 1

    # --- je Projekt ein Blatt
    spalten = ["Nr", "Verbindlichkeit", "Anforderung (Originalwortlaut)",
               "Operator", "Wert", "Einheit", "Fach", "Sprache",
               "Datei", "Seite", "Pruefen", "Warum", "Bemerkung"]
    breiten = [5, 14, 72, 9, 10, 10, 20, 8, 34, 7, 30, 34, 26]

    belegt: set[str] = set()
    for projekt in sorted(nach_projekt):
        eintraege = sorted(nach_projekt[projekt], key=sortschluessel)
        ws = wb.add_worksheet(blattname(projekt, belegt))
        for i, (s, b) in enumerate(zip(spalten, breiten)):
            ws.write(0, i, s, f_kopf)
            ws.set_column(i, i, b)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(eintraege), len(spalten) - 1)
        for n, r in enumerate(eintraege, start=1):
            fmt = f_muss if r.get("binding") in ("MUSS", "VERBOT") else f_text
            if r["_pruefen"] != "sicher":
                fmt = f_grau
            ws.write_row(n, 0, [
                n,
                r.get("binding", ""),
                norm(r.get("text")),
                r.get("operator") or "",
                wert_text(r),
                r.get("unit") or "",
                r.get("category", ""),
                r.get("lang", ""),
                r.get("file", ""),
                r.get("page", ""),
                r["_pruefen"],
                r["_warum"],
                "",
            ], fmt)

    wb.close()
    print()
    print("geschrieben:")
    print("  out\\anforderungen_zahlen.txt   nur Zahlen - darf geteilt werden")
    print(f"  {ziel.name}        mit Wortlauten - VERTRAULICH")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
