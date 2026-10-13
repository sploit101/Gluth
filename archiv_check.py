#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
archiv_check.py — beantwortet die eine offene Frage zum .7z-Archiv:

  Ist der Inhalt des Archivs NEU, oder liegt er ohnehin schon als lose
  Datei im Bestand? Davon haengt ab, ob das Archiv gelesen werden soll
  oder ob es 7 % Dubletten in jede Statistik traegt.

Aendert nichts. Liest nur out/chunks.jsonl.

  out/archiv_check_zahlen.txt   nur Zahlen        -> darf geteilt werden
  out/archiv_check_voll.txt     mit Dateinamen    -> VERTRAULICH

Aufruf:   python archiv_check.py
"""

from __future__ import annotations

import argparse
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

try:
    from requirements import split_sentences, load_chunks, chunk_meta
except Exception as e:  # pragma: no cover
    print("FEHLER: requirements.py konnte nicht geladen werden:", e)
    raise SystemExit(2)

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

ARCHIV = re.compile(r"\.(7z|zip|rar|tar|gz)#", re.I)


def ascii_safe(s: str) -> str:
    return s.encode("ascii", "replace").decode("ascii")


def norm(s: str) -> str:
    return re.sub(r"\W+", " ", s).strip().lower()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--chunks", default=str(OUT / "chunks.jsonl"))
    a = ap.parse_args()

    cpath = Path(a.chunks)
    if not cpath.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(str(cpath)))
        return 2

    arch_saetze, lose_saetze = [], set()
    arch_chunks = lose_chunks = 0
    innen = Counter()          # innerer Dateiname -> Anzahl Chunks
    lose_namen = set()
    archive = Counter()        # Archivname -> Anzahl Chunks

    beispiel_id = ""
    for c in load_chunks(cpath):
        meta = chunk_meta(c)
        text = c.get("text", "") or ""
        # Die Archivkennung steht je nach Extraktor in einem anderen Feld —
        # bei diesem Bestand im Feld "id", nicht in "source_path".
        pfad = ""
        for feld in ("id", "source_path", "filename", "page"):
            v = str(c.get(feld) or "")
            if ARCHIV.search(v):
                pfad = v
                break
        if not beispiel_id:
            beispiel_id = str(c.get("id") or "")
        if pfad:
            arch_chunks += 1
            kopf, _, rest = pfad.partition("#")
            archive[re.split(r"[\\/]", kopf)[-1]] += 1
            # innerer Name bis zum ersten Trennzeichen der Blattangabe
            inner = re.split(r"[↑→:|]", rest)[0].strip()
            innen[Path(inner).name or rest[:60]] += 1
            for s in split_sentences(text):
                arch_saetze.append(norm(s))
        else:
            lose_chunks += 1
            if meta.get("file"):
                lose_namen.add(str(meta["file"]).lower())
            for s in split_sentences(text):
                lose_saetze.add(norm(s))

    doppelt = sum(1 for s in arch_saetze if s in lose_saetze)
    n_arch = len(arch_saetze)
    innen_doppelt = [n for n in innen if n.lower() in lose_namen]

    L = []
    L.append(f"ARCHIV-PRUEFUNG  {datetime.now():%Y-%m-%d %H:%M}")
    L.append("=" * 62)
    L.append("")
    L.append("1 UMFANG")
    L.append(f"  Chunks aus Archiven          {arch_chunks}")
    L.append(f"  Chunks ausserhalb            {lose_chunks}")
    L.append(f"  Anteil Archiv                "
             f"{100*arch_chunks/max(arch_chunks+lose_chunks,1):.1f} %")
    L.append(f"  Saetze aus Archiven          {n_arch}")
    L.append("")
    L.append("2 ARCHIVE")
    for name, n in archive.most_common():
        L.append(f"    {n:>5} Chunks   {name}")
    L.append("")
    L.append("3 ENTHALTENE DOKUMENTE")
    L.append(f"  verschiedene Dokumente im Archiv   {len(innen)}")
    L.append(f"  davon auch lose im Bestand         {len(innen_doppelt)}"
             f"   ({100*len(innen_doppelt)/max(len(innen),1):.0f} %)")
    L.append("")
    L.append("4 SATZ-DUBLETTEN   (die entscheidende Zahl)")
    L.append(f"  Archivsaetze, die WORTGLEICH auch ausserhalb stehen   {doppelt}")
    L.append(f"  Anteil                                                "
             f"{100*doppelt/max(n_arch,1):.1f} %")
    L.append("")
    L.append("  Lesart:")
    L.append("    hoher Anteil  -> das Archiv traegt nur Dubletten, es kann")
    L.append("                     aus dem Suchpfad genommen werden")
    L.append("    tiefer Anteil -> im Archiv steckt eigener Inhalt, es bleibt")
    L.append("                     drin und braucht saubere Quellenangabe")
    L.append("")
    L.append("Diese Datei enthaelt keine Dateinamen und keine Saetze.")

    V = list(L)
    V += ["", "=" * 62, "DATEINAMEN --- VERTRAULICH", "=" * 62, ""]
    V.append("--- Dokumente im Archiv (Chunks je Dokument, * = liegt auch lose vor)")
    for name, n in innen.most_common():
        mark = " *" if name.lower() in lose_namen else "  "
        V.append(f"  {n:>5}{mark}  {name}")

    OUT.mkdir(exist_ok=True)
    (OUT / "archiv_check_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
    (OUT / "archiv_check_voll.txt").write_text("\n".join(V), encoding="utf-8")

    for z in L:
        print(ascii_safe(z))
    if arch_chunks == 0:
        print()
        print("HINWEIS: kein Archiv erkannt. Geprueft wurden die Felder")
        print("  id, source_path, filename, page auf das Muster '.7z#'.")
        print("  Beispiel aus dem Feld id:")
        print("   ", ascii_safe(beispiel_id[:160]))
        print("  Passt das Muster nicht, bitte diese Zeile schicken.")
    print()
    print("geschrieben:")
    print("  out\\archiv_check_zahlen.txt   nur Zahlen - darf geteilt werden")
    print("  out\\archiv_check_voll.txt     mit Dateinamen - VERTRAULICH")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
