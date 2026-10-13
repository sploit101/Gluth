#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bestandsaufnahme.py — misst die Loecher, bevor irgendetwas geaendert wird.

Diese Datei AENDERT NICHTS. Sie liest nur out/chunks.jsonl und (falls
vorhanden) out/manifest.json und zaehlt. Keine Installation noetig,
nur die Standardbibliothek und requirements.py, das daneben liegt.

Gemessen wird mit DERSELBEN Logik, die die Pipeline im Betrieb benutzt —
is_noise(), is_requirement(), detect_lang() werden importiert, nicht
nachgebaut. Deshalb beschreiben die Zahlen das echte Verhalten und nicht
eine zweite Meinung darueber.

Ergebnis sind ZWEI Dateien:

  out/bestandsaufnahme_voll.txt    mit Beispielsaetzen  -> VERTRAULICH, bleibt lokal
  out/bestandsaufnahme_zahlen.txt  nur Zahlen           -> darf geteilt werden

Aufruf:
    python bestandsaufnahme.py
    python bestandsaufnahme.py --chunks out\\chunks.jsonl --beispiele 20
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

try:
    from requirements import (
        split_sentences, detect_lang, classify_binding, is_noise,
        is_requirement, extract_measure, load_chunks, chunk_meta,
    )
except Exception as e:  # pragma: no cover
    print("FEHLER: requirements.py konnte nicht geladen werden:", e)
    print("Die Datei muss im selben Ordner liegen wie bestandsaufnahme.py.")
    raise SystemExit(2)

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

# Gruende aus is_noise(), die auf eine Tabellenzeile hindeuten
TAB_GRUENDE = {
    "ueberwiegend Ziffern (Tabelle)",
    "Tabellen- oder Fusszeile (Trennstriche)",
    "zu wenige echte Woerter",
}
# Gruende, die direkt am Ueberlappungsfehler in extract.py haengen
SCHNITT_GRUENDE = {
    "Satz bricht mitten in der Aussage ab (Chunk-Ende)",
    "beginnt klein — an der Chunk-Grenze abgeschnitten",
}


def ascii_safe(s: str) -> str:
    """Windows-Konsole vertraegt nicht jedes Zeichen. Nur fuer die Anzeige."""
    return s.encode("ascii", "replace").decode("ascii")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--chunks", default=str(OUT / "chunks.jsonl"))
    ap.add_argument("--manifest", default=str(OUT / "manifest.json"))
    ap.add_argument("--beispiele", type=int, default=15,
                    help="Beispielsaetze je Kategorie in der vertraulichen Datei")
    ap.add_argument("--seiten-schwelle", type=int, default=150,
                    help="Zeichen pro Seite, unter denen eine Datei als Scan gilt")
    a = ap.parse_args()

    cpath = Path(a.chunks)
    if not cpath.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(str(cpath)))
        print("Zuerst extract.py laufen lassen.")
        return 2

    chunks = load_chunks(cpath)

    # ---------------------------------------------------------------- zaehlen
    n_saetze = 0
    lang = Counter()
    lang_pro_projekt = defaultdict(Counter)
    verworfen = Counter()
    verworfen_mit_wert = Counter()      # verworfen, obwohl Zahl+Einheit drin
    durchgelassen = 0
    kein_signal = 0                     # sauber, aber keine Anforderung
    beispiele = defaultdict(list)

    datei_zeichen = defaultdict(int)
    datei_seiten = defaultdict(set)
    datei_projekt = {}
    projekte = set()

    for c in chunks:
        meta = chunk_meta(c)
        text = c.get("text", "") or ""
        datei = meta["file"] or "(ohne Namen)"
        datei_zeichen[datei] += len(text)
        if meta.get("page") is not None:
            datei_seiten[datei].add(meta["page"])
        datei_projekt[datei] = meta["project_no"]
        if meta["project_no"]:
            projekte.add(meta["project_no"])

        for sent in split_sentences(text):
            n_saetze += 1
            sp = detect_lang(sent)
            lang[sp] += 1
            lang_pro_projekt[meta["project_no"] or "(ohne Projekt)"][sp] += 1

            noise, grund = is_noise(sent)
            if noise:
                verworfen[grund] += 1
                m = extract_measure(sent)
                if m["value"] is not None and m["unit"]:
                    verworfen_mit_wert[grund] += 1
                    if len(beispiele["WERT_VERWORFEN"]) < a.beispiele:
                        beispiele["WERT_VERWORFEN"].append((grund, sent))
                if grund in SCHNITT_GRUENDE and len(beispiele["SCHNITT"]) < a.beispiele:
                    beispiele["SCHNITT"].append((grund, sent))
                if grund in TAB_GRUENDE and len(beispiele["TABELLE"]) < a.beispiele:
                    beispiele["TABELLE"].append((grund, sent))
                continue

            binding, _neg, _regel = classify_binding(sent)
            if is_requirement(sent, binding):
                durchgelassen += 1
            else:
                kein_signal += 1
                if len(beispiele["KEIN_SIGNAL"]) < a.beispiele:
                    beispiele["KEIN_SIGNAL"].append((binding, sent))

    # ------------------------------------------------------------ Scanverdacht
    scan_verdacht = []
    for datei, zeichen in datei_zeichen.items():
        seiten = len(datei_seiten.get(datei, ())) or 1
        pro_seite = zeichen / seiten
        if pro_seite < a.seiten_schwelle:
            scan_verdacht.append((datei, seiten, round(pro_seite)))
    scan_verdacht.sort(key=lambda t: t[2])

    # ---------------------------------------------------------------- manifest
    man_zeilen = []
    mpath = Path(a.manifest)
    if mpath.exists():
        try:
            man = json.loads(mpath.read_text(encoding="utf-8"))
            if isinstance(man, dict):
                man_zeilen.append(f"Schluessel im manifest: {sorted(man.keys())[:12]}")
                for k, v in man.items():
                    if isinstance(v, list):
                        man_zeilen.append(f"  {k}: {len(v)} Eintraege")
                    elif isinstance(v, (int, float, str)) and len(str(v)) < 60:
                        man_zeilen.append(f"  {k}: {v}")
            elif isinstance(man, list):
                man_zeilen.append(f"manifest ist eine Liste mit {len(man)} Eintraegen")
                if man and isinstance(man[0], dict):
                    man_zeilen.append(f"  Felder je Eintrag: {sorted(man[0].keys())}")
                    status = Counter(str(e.get("status") or e.get("ok") or
                                         e.get("result") or "?") for e in man
                                     if isinstance(e, dict))
                    for s, n in status.most_common(10):
                        man_zeilen.append(f"  status {s}: {n}")
        except Exception as e:
            man_zeilen.append(f"manifest.json nicht lesbar: {e}")
    else:
        man_zeilen.append("manifest.json nicht gefunden - uebersprungen")

    # ----------------------------------------------------------------- Bericht
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    n_verworfen = sum(verworfen.values())
    en_anteil = 100 * lang["en"] / max(n_saetze, 1)
    tab_verworfen = sum(n for g, n in verworfen.items() if g in TAB_GRUENDE)
    tab_mit_wert = sum(n for g, n in verworfen_mit_wert.items() if g in TAB_GRUENDE)
    schnitt = sum(n for g, n in verworfen.items() if g in SCHNITT_GRUENDE)
    schnitt_mit_wert = sum(n for g, n in verworfen_mit_wert.items()
                           if g in SCHNITT_GRUENDE)

    Z = []   # Zahlen - teilbar
    Z.append(f"BESTANDSAUFNAHME  {stamp}")
    Z.append("=" * 62)
    Z.append("")
    Z.append("1 BESTAND")
    Z.append(f"  Dateien                      {len(datei_zeichen)}")
    Z.append(f"  Projekte                     {len(projekte)}")
    Z.append(f"  Chunks                       {len(chunks)}")
    Z.append(f"  Saetze gesamt                {n_saetze}")
    Z.append("")
    Z.append("2 SPRACHE  (je Satz, ueber die Funktionswoerter)")
    Z.append(f"  deutsch                      {lang['de']}")
    Z.append(f"  englisch                     {lang['en']}   ({en_anteil:.1f} %)")
    Z.append("  je Projekt (nur Projekte mit >= 50 Saetzen):")
    for p, c in sorted(lang_pro_projekt.items(),
                       key=lambda t: -sum(t[1].values())):
        ges = sum(c.values())
        if ges >= 50:
            Z.append(f"    {p:<16} {ges:>6} Saetze   englisch "
                     f"{100*c['en']/ges:5.1f} %")
    Z.append("")
    Z.append("3 WAS VERWORFEN WIRD  (Grund aus is_noise)")
    Z.append(f"  verworfen gesamt             {n_verworfen}"
             f"   ({100*n_verworfen/max(n_saetze,1):.1f} % aller Saetze)")
    Z.append(f"  davon MIT Zahl und Einheit   {sum(verworfen_mit_wert.values())}"
             "   <- das sind vermutlich echte Anforderungen")
    Z.append("")
    Z.append("  Gruende, haeufigster zuerst   (in Klammern: mit Zahl+Einheit)")
    for grund, n in verworfen.most_common():
        Z.append(f"    {n:>6}  ({verworfen_mit_wert.get(grund,0):>4})  {grund}")
    Z.append("")
    Z.append("4 DIE DREI LOECHER")
    Z.append(f"  A Tabellenzeilen verworfen   {tab_verworfen}"
             f"   davon mit Zahl+Einheit {tab_mit_wert}")
    Z.append(f"  B Chunkgrenzen-Schaden       {schnitt}"
             f"   davon mit Zahl+Einheit {schnitt_mit_wert}")
    Z.append(f"  C Scanverdacht (Dateien)     {len(scan_verdacht)}"
             f" von {len(datei_zeichen)}"
             f"   ({100*len(scan_verdacht)/max(len(datei_zeichen),1):.1f} %)"
             f"   Schwelle {a.seiten_schwelle} Zeichen/Seite")
    Z.append("")
    Z.append("5 WAS DURCHGEHT")
    Z.append(f"  als Anforderung erkannt      {durchgelassen}")
    Z.append(f"  sauber, aber kein Signal     {kein_signal}")
    Z.append("")
    Z.append("6 MANIFEST")
    for z in man_zeilen:
        Z.append("  " + z)
    Z.append("")
    Z.append("Diese Datei enthaelt keine Saetze aus Kundendokumenten.")

    V = list(Z)   # Vollfassung - vertraulich
    V.append("")
    V.append("=" * 62)
    V.append("BEISPIELE  ---  VERTRAULICH, enthaelt Text aus Kundendokumenten")
    V.append("=" * 62)
    titel = {
        "WERT_VERWORFEN": "Verworfen, obwohl Zahl und Einheit drinstehen",
        "TABELLE": "Als Tabellenzeile verworfen",
        "SCHNITT": "An der Chunkgrenze zerschnitten (Ueberlappungsfehler)",
        "KEIN_SIGNAL": "Sauber, aber kein Anforderungssignal gefunden",
    }
    for key, tit in titel.items():
        V.append("")
        V.append(f"--- {tit}  ({len(beispiele[key])} Beispiele)")
        for grund, sent in beispiele[key]:
            V.append(f"  [{grund}]")
            V.append(f"    {sent[:300]}")
    V.append("")
    V.append("--- Scanverdacht: Dateien mit sehr wenig Text je Seite")
    for datei, seiten, pro in scan_verdacht[:40]:
        V.append(f"  {pro:>5} Zeichen/Seite   {seiten:>4} Seiten   {datei}")

    OUT.mkdir(exist_ok=True)
    (OUT / "bestandsaufnahme_zahlen.txt").write_text("\n".join(Z), encoding="utf-8")
    (OUT / "bestandsaufnahme_voll.txt").write_text("\n".join(V), encoding="utf-8")

    for z in Z:
        print(ascii_safe(z))
    print()
    print("geschrieben:")
    print("  out\\bestandsaufnahme_zahlen.txt   nur Zahlen - darf geteilt werden")
    print("  out\\bestandsaufnahme_voll.txt     mit Beispielsaetzen - VERTRAULICH")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
