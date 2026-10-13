#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
manifest_check.py — beantwortet die drei Fragen, die die Bestandsaufnahme
offengelassen hat. Aendert nichts, liest nur.

  1. Welche 37 Dateien brauchen OCR, und wie gross sind sie?
  2. Was sind die 32 uebersprungenen Dateien, und mit welcher Begruendung?
  3. Warum stehen im Manifest 197 gelesene Dateien, aber nur 152 im Chunk-Bestand?

Schreibt wieder zwei Dateien:
  out/manifest_check_zahlen.txt   nur Zahlen und Endungen   -> darf geteilt werden
  out/manifest_check_voll.txt     mit Dateinamen            -> VERTRAULICH

Aufruf:   python manifest_check.py
"""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"


def ascii_safe(s: str) -> str:
    return s.encode("ascii", "replace").decode("ascii")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--manifest", default=str(OUT / "manifest.json"))
    ap.add_argument("--chunks", default=str(OUT / "chunks.jsonl"))
    a = ap.parse_args()

    mp, cp = Path(a.manifest), Path(a.chunks)
    if not mp.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(str(mp)))
        return 2

    man = json.loads(mp.read_text(encoding="utf-8"))
    if isinstance(man, dict):                       # falls doch ein Objekt
        for v in man.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                man = v
                break
    if not isinstance(man, list):
        print("FEHLER: manifest.json hat eine unerwartete Form.")
        return 2

    # ------------------------------------------------------------- Chunkseite
    pfade, namen = set(), set()
    if cp.exists():
        with cp.open(encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                c = json.loads(line)
                if c.get("source_path"):
                    pfade.add(str(c["source_path"]))
                if c.get("filename"):
                    namen.add(str(c["filename"]))

    # ------------------------------------------------------------- Manifest
    nach_status = defaultdict(list)
    for e in man:
        if isinstance(e, dict):
            nach_status[str(e.get("status", "?"))].append(e)

    def summe(eintraege, feld):
        s = 0
        for e in eintraege:
            v = e.get(feld)
            if isinstance(v, (int, float)):
                s += v
        return s

    Z, V = [], []
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    Z.append(f"MANIFEST-PRUEFUNG  {stamp}")
    Z.append("=" * 62)
    Z.append("")
    Z.append(f"Eintraege im Manifest        {len(man)}")
    Z.append(f"verschiedene Pfade in chunks {len(pfade)}")
    Z.append(f"verschiedene Dateinamen      {len(namen)}"
             "   <- wenn kleiner als die Pfade, gibt es gleiche Namen in mehreren Ordnern")
    Z.append("")

    for status in sorted(nach_status, key=lambda s: -len(nach_status[s])):
        eintraege = nach_status[status]
        Z.append(f"--- status {status}: {len(eintraege)} Dateien")
        Z.append(f"      Seiten gesamt   {summe(eintraege,'pages')}")
        Z.append(f"      Zeichen gesamt  {summe(eintraege,'chars')}")
        Z.append(f"      Chunks gesamt   {summe(eintraege,'chunks')}")
        ext = Counter(str(e.get("ext", "?")).lower() for e in eintraege)
        Z.append("      Endungen        " +
                 ", ".join(f"{k} {v}" for k, v in ext.most_common(8)))
        gruende = Counter(str(e.get("reason") or "-") for e in eintraege)
        Z.append("      Begruendung     " +
                 "; ".join(f"{k} ({v})" for k, v in gruende.most_common(6)))
        proj = Counter(str(e.get("project_no") or "-") for e in eintraege)
        Z.append("      Projekte        " +
                 ", ".join(f"{k} {v}" for k, v in proj.most_common(6)))
        Z.append("")

        V.append(f"--- status {status}: {len(eintraege)} Dateien")
        for e in sorted(eintraege, key=lambda x: -(x.get("pages") or 0)):
            V.append(f"   {str(e.get('pages') or '-'):>5} S. "
                     f"{str(e.get('chars') or '-'):>8} Z.  "
                     f"{e.get('project_no','-'):<12} {e.get('filename','?')}"
                     f"   [{e.get('reason') or '-'}]")
        V.append("")

    # ------------------------------------------- ok, aber nicht in den Chunks
    ok = nach_status.get("ok", [])
    fehlt = [e for e in ok
             if str(e.get("source_path", "")) not in pfade
             and str(e.get("filename", "")) not in namen]
    Z.append(f"--- als ok gemeldet, aber im Chunk-Bestand nicht auffindbar: {len(fehlt)}")
    if fehlt:
        Z.append("      Endungen        " +
                 ", ".join(f"{k} {v}" for k, v in
                           Counter(str(e.get("ext","?")).lower()
                                   for e in fehlt).most_common(8)))
        Z.append(f"      Zeichen gesamt  {summe(fehlt,'chars')}")
        Z.append(f"      Chunks gemeldet {summe(fehlt,'chunks')}"
                 "   <- steht hier 0, hat die Datei nie Text geliefert")
        V.append("--- als ok gemeldet, aber nicht im Chunk-Bestand")
        for e in fehlt:
            V.append(f"   {str(e.get('chars') or '-'):>8} Z.  "
                     f"{str(e.get('chunks') or '-'):>4} Chunks  "
                     f"{e.get('filename','?')}")
    Z.append("")
    Z.append("Diese Datei enthaelt keine Dateinamen und keine Saetze.")

    OUT.mkdir(exist_ok=True)
    (OUT / "manifest_check_zahlen.txt").write_text("\n".join(Z), encoding="utf-8")
    (OUT / "manifest_check_voll.txt").write_text(
        "\n".join(Z + ["", "=" * 62,
                       "DATEINAMEN --- VERTRAULICH", "=" * 62, ""] + V),
        encoding="utf-8")

    for z in Z:
        print(ascii_safe(z))
    print()
    print("geschrieben:")
    print("  out\\manifest_check_zahlen.txt   nur Zahlen - darf geteilt werden")
    print("  out\\manifest_check_voll.txt     mit Dateinamen - VERTRAULICH")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
