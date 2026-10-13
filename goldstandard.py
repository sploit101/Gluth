#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
goldstandard.py — die fehlende Messung: ist das Erkannte richtig?

Zwei Betriebsarten.

  python goldstandard.py --ziehen
      Zieht eine Stichprobe und schreibt eine Excel-taugliche Datei zum
      Ausfuellen von Hand. Die Datei enthaelt NICHT, was das System
      entschieden hat - sonst wird das Urteil davon beeinflusst
      ("Ankereffekt"). Das Systemurteil wird getrennt abgelegt.

  python goldstandard.py --messen out\\goldstandard_vorlage.csv
      Vergleicht die ausgefuellte Datei mit dem Systemurteil und rechnet
      Trefferquote und Sauberkeit getrennt aus - inklusive der Frage,
      MIT WELCHER BEGRUENDUNG die uebersehenen Anforderungen verworfen wurden.
      Das ist die Zahl, die entscheidet, ob der Tabellenmodus sich lohnt.

Aendert nichts am Bestand. Schreibt nur nach out\\.

Dateien:
  out/goldstandard_vorlage.csv   zum Ausfuellen   -> VERTRAULICH (enthaelt Saetze)
  out/goldstandard_system.json   Systemurteil     -> ohne Satztext
  out/goldstandard_ergebnis.txt  Messergebnis     -> nur Zahlen, darf geteilt werden
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

try:
    from requirements import (
        split_sentences, detect_lang, classify_binding, is_noise,
        is_requirement, extract_measure, detect_operator,
        load_chunks, chunk_meta,
    )
except Exception as e:  # pragma: no cover
    print("FEHLER: requirements.py konnte nicht geladen werden:", e)
    raise SystemExit(2)

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

ARCHIV = re.compile(r"\.(7z|zip|rar|tar|gz)#", re.I)


def _datei(a, art: str) -> Path:
    """goldstandard_vorlage.csv  bzw. mit --name X: goldstandard_X_vorlage.csv.
    Damit ueberschreibt eine zweite Stichprobe nie die erste."""
    name = getattr(a, "name", "") or ""
    return OUT / (f"goldstandard_{name}_{art}" if name else f"goldstandard_{art}")

SPALTEN = [
    "id", "datei", "seite", "sprache", "satz",
    "ANFORDERUNG_ja_nein", "VERBINDLICHKEIT", "OPERATOR", "WERT", "EINHEIT",
    "BEMERKUNG",
]

JA = {"ja", "j", "y", "yes", "x", "1", "true", "wahr"}
NEIN = {"nein", "n", "no", "0", "false", "falsch"}


def ascii_safe(s: str) -> str:
    return s.encode("ascii", "replace").decode("ascii")


def system_urteil(satz: str) -> dict:
    """Was die Pipeline heute mit diesem Satz macht — ohne Router."""
    noise, grund = is_noise(satz)
    if noise:
        return {"verdikt": "verworfen", "grund": grund, "binding": "",
                "operator": "", "wert": "", "einheit": ""}
    binding, _neg, _regel = classify_binding(satz)
    m = extract_measure(satz)
    op = detect_operator(satz) or ""
    treffer = is_requirement(satz, binding)
    return {
        "verdikt": "erkannt" if treffer else "kein_signal",
        "grund": "",
        "binding": binding,
        "operator": op or "",
        "wert": "" if m.get("value") is None else str(m["value"]),
        "einheit": m.get("unit") or "",
    }


# ===================================================================== ziehen
def ziehen(a) -> int:
    cpath = Path(a.chunks)
    if not cpath.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(str(cpath)))
        return 2

    alle = []
    uebersprungen = 0
    for c in load_chunks(cpath):
        # Archivinhalte sind nachweislich Dubletten des uebrigen Bestands
        # (archiv_check.py: 99,9 %). Wer sie mitzieht, misst dieselben
        # Saetze doppelt. Mit --mit-archiv laesst sich das abschalten.
        if not a.mit_archiv:
            treffer = False
            for feld in ("id", "source_path", "filename", "page"):
                if ARCHIV.search(str(c.get(feld) or "")):
                    treffer = True
                    break
            if treffer:
                uebersprungen += 1
                continue
        meta = chunk_meta(c)
        for satz in split_sentences(c.get("text", "") or ""):
            alle.append((satz, meta))

    print(f"{len(alle)} Saetze im Bestand.")
    if uebersprungen:
        print(f"{uebersprungen} Chunks aus Archiven uebersprungen (Dubletten).")

    urteile = [system_urteil(s) for s, _ in alle]
    erkannt_idx = [i for i, u in enumerate(urteile) if u["verdikt"] == "erkannt"]

    rnd = random.Random(a.seed)
    # Haelfte aus ALLEN Saetzen  -> misst die Trefferquote (was wird uebersehen)
    # Haelfte aus den ERKANNTEN  -> misst die Sauberkeit  (was ist falsch drin)
    # WICHTIG: Die Trefferquote darf NUR aus dem Zufallstopf gerechnet werden.
    # Im Erkannt-Topf ist ein Uebersehen per Konstruktion unmoeglich; wer beide
    # Toepfe zusammenrechnet, bekommt eine geschoente Trefferquote.
    # Die Reihenfolge der rnd-Aufrufe darf sich nie aendern - sonst ziehen
    # gleiche Seeds andere Saetze und bereits ausgefuellte Dateien passen nicht mehr.
    n_alle = a.n // 2
    n_erk = a.n - n_alle
    topf_zufall = set(rnd.sample(range(len(alle)), min(n_alle, len(alle))))
    auswahl = set(topf_zufall)
    if erkannt_idx:
        auswahl |= set(rnd.sample(erkannt_idx, min(n_erk, len(erkannt_idx))))
    auswahl = sorted(auswahl)
    rnd.shuffle(auswahl)     # damit die Reihenfolge nichts verraet

    zeilen, system = [], {}
    for lauf, i in enumerate(auswahl, 1):
        satz, meta = alle[i]
        # Der Satz wird genau so abgelegt, wie er gemessen wurde - sonst
        # beurteilt --messen spaeter einen anderen Text als --ziehen.
        satz = " ".join(satz.split())[:2000]
        urteile[i] = system_urteil(satz)
        sid = f"S{lauf:04d}"
        zeilen.append({
            "id": sid,
            "datei": meta.get("file", ""),
            "seite": meta.get("page") if meta.get("page") is not None else "",
            "sprache": detect_lang(satz),
            "satz": satz,
            "ANFORDERUNG_ja_nein": "", "VERBINDLICHKEIT": "", "OPERATOR": "",
            "WERT": "", "EINHEIT": "", "BEMERKUNG": "",
        })
        urteile[i]["topf"] = "zufall" if i in topf_zufall else "erkannt"
        system[sid] = urteile[i]

    OUT.mkdir(exist_ok=True)
    vpath = _datei(a, "vorlage.csv")
    # utf-8-sig + Semikolon: so oeffnet deutsches Excel die Datei richtig
    with vpath.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=SPALTEN, delimiter=";",
                           quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(zeilen)
    _datei(a, "system.json").write_text(
        json.dumps({"seed": a.seed, "n": len(zeilen), "urteile": system},
                   ensure_ascii=False, indent=1), encoding="utf-8")

    print()
    print(f"{len(zeilen)} Saetze gezogen (seed {a.seed}).")
    print(f"  out\\{_datei(a, 'vorlage.csv').name}   <- diese Datei ausfuellen (VERTRAULICH)")
    print(f"  out\\{_datei(a, 'system.json').name}   <- nicht anfassen, nicht ansehen")
    print()
    print("Auszufuellen sind nur die GROSS geschriebenen Spalten:")
    print("  ANFORDERUNG_ja_nein   ja / nein   (Pflicht)")
    print("  VERBINDLICHKEIT       MUSS SOLL KANN VERBOT INFO   (wenn ja)")
    print("  OPERATOR              LE GE LT GT EQ RANGE   (wenn ein Grenzwert da ist)")
    print("  WERT / EINHEIT        z.B. 32 / s            (wenn vorhanden)")
    print("  BEMERKUNG             frei, fuer Zweifelsfaelle")
    return 0


# ======================================== verworfene Saetze mit Messwert
def ziehen_verworfen(a) -> int:
    """
    Zieht gezielt aus den Saetzen, die das System VERWORFEN hat und die
    trotzdem Zahl und Einheit enthalten. Das ist der direkte Test der
    Tabellen-Vermutung: sind das echte, verlorene Anforderungen oder nicht?
    """
    cpath = Path(a.chunks)
    if not cpath.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(str(cpath)))
        return 2

    kandidaten = []
    for c in load_chunks(cpath):
        if not a.mit_archiv:
            if any(ARCHIV.search(str(c.get(f) or ""))
                   for f in ("id", "source_path", "filename", "page")):
                continue
        meta = chunk_meta(c)
        for satz in split_sentences(c.get("text", "") or ""):
            satz = " ".join(satz.split())[:2000]
            noise, grund = is_noise(satz)
            if not noise:
                continue
            m = extract_measure(satz)
            if m.get("value") is None or not m.get("unit"):
                continue
            kandidaten.append((satz, meta, grund))

    print(f"{len(kandidaten)} verworfene Saetze mit Zahl UND Einheit im Bestand.")
    print("Das ist die Grundgesamtheit - der gemessene Anteil laesst sich")
    print("auf diese Zahl hochrechnen.")
    if not kandidaten:
        return 1

    rnd = random.Random(a.seed)
    idx = rnd.sample(range(len(kandidaten)), min(a.n, len(kandidaten)))
    rnd.shuffle(idx)

    zeilen, system = [], {}
    for lauf, i in enumerate(idx, 1):
        satz, meta, grund = kandidaten[i]
        sid = f"V{lauf:04d}"
        zeilen.append({
            "id": sid,
            "datei": meta.get("file", ""),
            "seite": meta.get("page") if meta.get("page") is not None else "",
            "sprache": detect_lang(satz),
            "satz": satz,
            "ANFORDERUNG_ja_nein": "", "VERBINDLICHKEIT": "", "OPERATOR": "",
            "WERT": "", "EINHEIT": "", "BEMERKUNG": "",
        })
        u = system_urteil(satz)
        u["topf"] = "zufall"
        system[sid] = u

    OUT.mkdir(exist_ok=True)
    with (OUT / "goldstandard_verworfen_vorlage.csv").open(
            "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=SPALTEN, delimiter=";",
                           quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(zeilen)
    (OUT / "goldstandard_verworfen_system.json").write_text(
        json.dumps({"seed": a.seed, "n": len(zeilen),
                    "grundgesamtheit": len(kandidaten), "urteile": system},
                   ensure_ascii=False, indent=1), encoding="utf-8")

    print()
    print(f"{len(zeilen)} Saetze gezogen (seed {a.seed}).")
    print("  out\\goldstandard_verworfen_vorlage.csv   <- ausfuellen (VERTRAULICH)")
    print("  out\\goldstandard_verworfen_system.json")
    print()
    print("Messen spaeter mit:")
    print("  python goldstandard.py --messen out\\goldstandard_verworfen_gefuellt.csv \\")
    print("         --stand out\\goldstandard_verworfen_system.json")
    return 0


# ===================================================================== messen
def _ja(v: str) -> bool | None:
    s = (v or "").strip().lower()
    if s in JA:
        return True
    if s in NEIN:
        return False
    return None


def messen(a) -> int:
    fpath = Path(a.messen)
    spath = Path(a.stand) if a.stand else _datei(a, "system.json")
    if not fpath.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(str(fpath)))
        return 2
    if not spath.exists():
        print("FEHLER: out\\goldstandard_system.json fehlt - erst --ziehen.")
        return 2

    alt = json.loads(spath.read_text(encoding="utf-8"))["urteile"]

    with fpath.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh, delimiter=";"))

    # Das Systemurteil wird aus dem Satztext NEU berechnet, nicht aus der
    # Datei von der Ziehung gelesen. Nur so misst dieselbe ausgefuellte
    # Stichprobe nach jeder Codeaenderung den AKTUELLEN Stand. Die alten
    # Urteile dienen als Vergleich "vorher/nachher".
    system = {}
    for r in rows:
        sid = (r.get("id") or "").strip()
        satz = (r.get("satz") or "").strip()
        if sid and satz:
            system[sid] = system_urteil(satz)
        elif sid in alt:
            system[sid] = alt[sid]

    a_tp = a_fp = a_fn = 0
    offen = 0
    tp = fp = fn = tn = 0
    verpasst_grund = Counter()
    falsch_drin_grund = Counter()
    bind_richtig = bind_falsch = 0
    bind_verwechslung = Counter()
    wert_richtig = wert_falsch = 0
    n_bewertet = 0

    for r in rows:
        sid = (r.get("id") or "").strip()
        u = system.get(sid)
        if not u:
            continue
        mensch = _ja(r.get("ANFORDERUNG_ja_nein", ""))
        if mensch is None:
            offen += 1
            continue
        n_bewertet += 1
        maschine = u["verdikt"] == "erkannt"

        av = alt.get(sid, {}).get("verdikt")
        if av is not None:
            if mensch and av == "erkannt":
                a_tp += 1
            elif mensch:
                a_fn += 1
            elif av == "erkannt":
                a_fp += 1

        if mensch and maschine:
            tp += 1
            mb = (r.get("VERBINDLICHKEIT") or "").strip().upper()
            if mb:
                if mb == u["binding"]:
                    bind_richtig += 1
                else:
                    bind_falsch += 1
                    bind_verwechslung[f"Mensch {mb} / System {u['binding']}"] += 1
            mw = (r.get("WERT") or "").strip().replace(",", ".")
            me = (r.get("EINHEIT") or "").strip()
            if mw or me:
                sw = (u["wert"] or "").strip()
                gleich = True
                if mw:
                    try:
                        gleich &= abs(float(mw) - float(sw or "nan")) < 1e-9
                    except ValueError:
                        gleich = False
                if me:
                    gleich &= me.lower() == (u["einheit"] or "").lower()
                wert_richtig += 1 if gleich else 0
                wert_falsch += 0 if gleich else 1
        elif mensch and not maschine:
            fn += 1
            verpasst_grund[u["grund"] or f"kein Signal ({u['binding']})"] += 1
        elif not mensch and maschine:
            fp += 1
            falsch_drin_grund[u["binding"] or "?"] += 1
        else:
            tn += 1

    # --- Trefferquote sauber: nur die zufaellig gezogenen Zeilen ----------
    z_tp = z_fn = 0
    z_bekannt = 0
    for r in rows:
        sid = (r.get("id") or "").strip()
        u = system.get(sid)
        topf = alt.get(sid, {}).get("topf")
        if not u or topf is None:
            continue
        z_bekannt += 1
        if topf != "zufall":
            continue
        mensch = _ja(r.get("ANFORDERUNG_ja_nein", ""))
        if mensch is None or not mensch:
            continue
        if u["verdikt"] == "erkannt":
            z_tp += 1
        else:
            z_fn += 1

    def quote(z, n):
        return f"{100*z/n:5.1f} %" if n else "   —  "

    L = []
    L.append(f"GOLDSTANDARD-MESSUNG  {datetime.now():%Y-%m-%d %H:%M}")
    L.append("=" * 62)
    L.append(f"  Zeilen in der Datei          {len(rows)}")
    L.append(f"  davon bewertet               {n_bewertet}")
    L.append(f"  noch offen (leer gelassen)   {offen}")
    L.append("")
    L.append("1 IST ES UEBERHAUPT EINE ANFORDERUNG")
    L.append(f"  richtig erkannt      (TP)    {tp}")
    L.append(f"  uebersehen           (FN)    {fn}")
    L.append(f"  faelschlich erkannt  (FP)    {fp}")
    L.append(f"  richtig verworfen    (TN)    {tn}")
    L.append("")
    L.append(f"  SAUBERKEIT   (precision)     {quote(tp, tp+fp)}"
             "   <- wie viel von dem, was ankommt, stimmt")
    L.append("")
    if z_bekannt and (z_tp + z_fn):
        L.append("  TREFFERQUOTE (recall)        "
                 f"{quote(z_tp, z_tp+z_fn)}"
                 "   <- NUR aus dem Zufallstopf gerechnet")
        L.append(f"     Grundlage: {z_tp+z_fn} Anforderungen im Zufallstopf"
                 f"  ({z_tp} gefunden, {z_fn} uebersehen)")
        if (tp + fn) != (z_tp + z_fn):
            L.append(f"     zum Vergleich ueber ALLE Zeilen: {quote(tp, tp+fn)}"
                     " - geschoent, NICHT verwenden:")
            L.append("     im Erkannt-Topf ist ein Uebersehen unmoeglich.")
    elif not z_bekannt:
        L.append("  TREFFERQUOTE                 nicht berechenbar")
        L.append("     Die Datei goldstandard_system.json stammt aus einer aelteren")
        L.append("     Fassung ohne Topf-Kennzeichnung. Einmal --ziehen mit demselben")
        L.append("     Seed wiederholen; die Stichprobe bleibt identisch.")
    else:
        L.append("  TREFFERQUOTE                 keine Anforderung im Zufallstopf")
    if (a_tp, a_fp, a_fn) != (tp, fp, fn):
        L.append("")
        L.append("  Stand bei der Ziehung  (zum Vergleich - vorher/nachher)")
        L.append(f"    Trefferquote               {quote(a_tp, a_tp+a_fn)}"
                 f"   jetzt {quote(tp, tp+fn)}")
        L.append(f"    Sauberkeit                 {quote(a_tp, a_tp+a_fp)}"
                 f"   jetzt {quote(tp, tp+fp)}")
    # --- Gegenrechnung ohne die abgeschnittenen Saetze -------------------
    # Ein abgeschnittener Satz ist fuer den Leser wertlos, also zaehlt er im
    # Hauptergebnis als Fehlalarm. Die Ursache liegt aber NICHT im Regelwerk,
    # sondern im Ueberlappungsfehler von extract.py. Ohne diese Trennung
    # wuerde man an den Regeln schrauben, um einen Extraktionsfehler
    # auszugleichen. Die Differenz der beiden Bloecke ist der Preis des Fehlers.
    o_tp = o_fp = o_fn = o_tn = 0
    oz_tp = oz_fn = 0          # nur Zufallstopf - fuer die Trefferquote
    for r in rows:
        if "abgeschnitten" in (r.get("BEMERKUNG") or "").lower():
            continue
        sid = (r.get("id") or "").strip()
        u = system.get(sid)
        if not u:
            continue
        mensch = _ja(r.get("ANFORDERUNG_ja_nein", ""))
        if mensch is None:
            continue
        maschine = u["verdikt"] == "erkannt"
        if mensch and maschine:
            o_tp += 1
        elif mensch:
            o_fn += 1
        elif maschine:
            o_fp += 1
        else:
            o_tn += 1
        if alt.get(sid, {}).get("topf") == "zufall" and mensch:
            if maschine:
                oz_tp += 1
            else:
                oz_fn += 1
    weg = n_bewertet - (o_tp + o_fp + o_fn + o_tn)
    if weg:
        L.append("")
        L.append("1b OHNE DIE ABGESCHNITTENEN SAETZE")
        L.append(f"  ausgeklammert                {weg} Zeilen"
                 " (Bemerkung enthaelt 'abgeschnitten')")
        L.append(f"  SAUBERKEIT   (precision)     {quote(o_tp, o_tp+o_fp)}"
                 f"   mit ihnen: {quote(tp, tp+fp)}")
        if oz_tp + oz_fn:
            L.append(f"  TREFFERQUOTE (recall)        {quote(oz_tp, oz_tp+oz_fn)}"
                     f"   mit ihnen: {quote(z_tp, z_tp+z_fn)}"
                     "   (nur Zufallstopf)")
        else:
            L.append("  TREFFERQUOTE                 hier nicht ausgewiesen")
            L.append("     (braucht die Topf-Kennzeichnung aus --ziehen)")
        L.append("  Die Differenz ist der Preis des Ueberlappungsfehlers in")
        L.append("  extract.py - nicht ein Fehler des Regelwerks.")

    L.append("")
    L.append("2 WARUM DIE UEBERSEHENEN VERLOREN GINGEN")
    if verpasst_grund:
        for g, n in verpasst_grund.most_common():
            L.append(f"    {n:>4}  {g}")
    else:
        L.append("    keine uebersehen")
    L.append("")
    L.append("3 WAS FAELSCHLICH DRIN IST  (Einstufung des Systems)")
    if falsch_drin_grund:
        for g, n in falsch_drin_grund.most_common():
            L.append(f"    {n:>4}  {g}")
    else:
        L.append("    nichts faelschlich drin")
    L.append("")
    L.append("4 VERBINDLICHKEIT  (nur bei uebereinstimmend erkannten)")
    L.append(f"  richtig                      {bind_richtig}")
    L.append(f"  falsch                       {bind_falsch}")
    L.append(f"  Genauigkeit                  "
             f"{quote(bind_richtig, bind_richtig+bind_falsch)}")
    for g, n in bind_verwechslung.most_common(8):
        L.append(f"    {n:>4}  {g}")
    L.append("")
    L.append("5 WERT UND EINHEIT  (nur wo der Mensch etwas eingetragen hat)")
    L.append(f"  richtig                      {wert_richtig}")
    L.append(f"  falsch                       {wert_falsch}")
    L.append(f"  Genauigkeit                  "
             f"{quote(wert_richtig, wert_richtig+wert_falsch)}")
    L.append("")
    L.append("Diese Datei enthaelt keine Saetze aus Kundendokumenten.")

    OUT.mkdir(exist_ok=True)
    _datei(a, "ergebnis.txt").write_text("\n".join(L), encoding="utf-8")
    for z in L:
        print(ascii_safe(z))
    print()
    print(f"geschrieben: out\\{_datei(a, 'ergebnis.txt').name}   (nur Zahlen - darf geteilt werden)")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--chunks", default=str(OUT / "chunks.jsonl"))
    ap.add_argument("--ziehen", action="store_true")
    ap.add_argument("--messen", metavar="CSV", default=None)
    ap.add_argument("--n", type=int, default=100)
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--ziehen-verworfen", action="store_true",
                    help="gezielt verworfene Saetze mit Zahl und Einheit ziehen")
    ap.add_argument("--name", default="",
                    help="eigener Name fuer eine weitere Stichprobe, z.B. --name pruefung2")
    ap.add_argument("--stand", default=None,
                    help="andere Systemurteil-Datei (fuer die zweite Stichprobe)")
    ap.add_argument("--mit-archiv", action="store_true",
                    help="Archivinhalte mitziehen (Standard: ueberspringen)")
    a = ap.parse_args()
    if a.messen:
        return messen(a)
    if a.ziehen_verworfen:
        return ziehen_verworfen(a)
    if a.ziehen:
        return ziehen(a)
    ap.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
