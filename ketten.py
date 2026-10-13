#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ketten.py — zeigt, welche Anforderung in mehreren Fassungen vorliegt.

Beantwortet die Frage, um die es in der Entwicklungsakte geht:
"Dieselbe Sache steht in zwei Dokumentfassungen — ist sie gleich geblieben
oder hat sie sich geaendert, und wo genau?"

Grundlage ist der subject_key aus requirements.py: die Kennung des
Gegenstands OHNE die Zahlenwerte. Dadurch gehoeren
  "Die Taktzeit muss max. 32 s betragen."   und
  "Die Taktzeit muss max. 28 s betragen."
zur selben Kette, obwohl kein Textvergleich sie zusammenbringt.

Kein Sprachmodell. Liest nur out/requirements.jsonl, aendert nichts.

    python ketten.py
    python ketten.py --projekt A07616000
    python ketten.py --datei LH_04K_000051          # nur Dateien mit diesem Namensteil
    python ketten.py --nur-aenderungen              # unveraenderte weglassen

Ausgabe
    out/ketten_zahlen.txt   nur Zahlen       -> darf geteilt werden
    out/ketten_voll.xlsx    mit Wortlauten   -> VERTRAULICH

------------------------------------------------------------------------
Korrekturen 02.10.2026, nachdem die erste Fassung an echten Daten geprueft
wurde:

1. Paare wurden auch INNERHALB derselben Datei gebildet. Von 427 Paaren
   stammten 215 aus ein und demselben Dokument — das ist kein Verlauf,
   sondern eine Wiederholung. Es werden jetzt ausschliesslich Paare ueber
   DOKUMENTGRENZEN hinweg gebildet.

2. "Text geaendert" meldete 44 Faelle, von denen 36 keine inhaltliche
   Aenderung waren, sondern Unterschiede der Textgewinnung: ein
   Aufzaehlungspunkt, der in der einen Fassung als Zeichen und in der
   anderen als Leerzeichen ankommt, eine angehaengte Kapitelnummer, ein
   am Rand abgeschnittener Buchstabe. Solche Faelle heissen jetzt
   "nur Darstellung" und stehen nicht mehr bei den Aenderungen.

   Die Regel dafuer steht in unterschied() und ist bewusst konservativ:
   Sobald sich eine Zahl unterscheidet, gilt es IMMER als Aenderung.
------------------------------------------------------------------------
"""

from __future__ import annotations

import argparse
import collections
import difflib
import json
import re
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

# Zeichen, die nur aus der PDF-Textgewinnung stammen: Aufzaehlungspunkte in
# verschiedenen Schriften, Schloss- und Stiftsymbole aus Schutzvermerken.
ARTEFAKT = re.compile(r"[•▪►●·✎✓"
                      r"\U0001F512\U0001F50F\U0001F513]")
# Angehaengte oder vorangestellte Gliederungsnummern: " 12.", "3.4 "
RANDNUMMER = re.compile(r"^[\s\d.,;:()]+|[\s\d.,;:()]+$")
ZAHL = re.compile(r"\d+(?:[.,]\d+)?")

# Ab dieser Aehnlichkeit gilt ein Unterschied ohne abweichende Zahl als
# Darstellungsfrage. 0.98 heisst: hoechstens zwei Prozent der Zeichen
# weichen ab — das sind einzelne Buchstaben, keine Aussagen.
AEHNLICH_GRENZE = 0.98


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def kern(s) -> str:
    """Text ohne Extraktionsartefakte und ohne Randnummern."""
    t = re.sub(r"\s+", " ", ARTEFAKT.sub(" ", str(s or ""))).strip()
    return RANDNUMMER.sub("", t).strip().lower()


def wert(r: dict) -> str:
    if r.get("value") is None:
        return ""
    teile = [r.get("operator") or "", f'{r["value"]:g}']
    if r.get("value_max") is not None:
        teile.append(f'bis {r["value_max"]:g}')
    teile.append(r.get("unit") or "")
    return " ".join(t for t in teile if t).strip()


def unterschied(a: dict, b: dict) -> str:
    """Wie unterscheiden sich zwei Fassungen derselben Anforderung?"""
    ka, kb = kern(a["text"]), kern(b["text"])
    zahlen_gleich = ZAHL.findall(ka) == ZAHL.findall(kb)

    if wert(a) != wert(b):
        return "WERT GEAENDERT"
    if a.get("binding") != b.get("binding"):
        return "VERBINDLICHKEIT GEAENDERT"
    if ka == kb:
        return ("unveraendert" if norm(a["text"]) == norm(b["text"])
                else "nur Darstellung")
    if zahlen_gleich and difflib.SequenceMatcher(None, ka, kb).ratio() >= AEHNLICH_GRENZE:
        return "nur Darstellung"
    return "Text geaendert"


def paare_bilden(eintraege: list[dict]) -> list[tuple[dict, dict, bool]]:
    """
    Paare NUR ueber Dokumentgrenzen. Gibt (alt, neu, mehrdeutig) zurueck.

    Liegen in einer Fassung mehrere Treffer zum selben Gegenstand, ist die
    Zuordnung nicht eindeutig — z. B. zwei Heizungen mit verschiedenen
    Temperaturbereichen im selben Dokument. Solche Paare werden als
    mehrdeutig markiert statt stillschweigend geraten.
    """
    nach_datei: dict[str, list[dict]] = collections.defaultdict(list)
    for e in eintraege:
        nach_datei[str(e.get("file") or "")].append(e)
    dateien = sorted(nach_datei)
    if len(dateien) < 2:
        return []
    raus = []
    for i in range(len(dateien) - 1):
        a = sorted(nach_datei[dateien[i]], key=lambda e: str(e.get("page") or ""))
        b = sorted(nach_datei[dateien[i + 1]], key=lambda e: str(e.get("page") or ""))
        mehrdeutig = len(a) != 1 or len(b) != 1
        for k in range(min(len(a), len(b))):
            raus.append((a[k], b[k], mehrdeutig))
    return raus


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--req", default=str(OUT / "requirements.jsonl"))
    ap.add_argument("--projekt", default=None)
    ap.add_argument("--datei", default=None)
    ap.add_argument("--nur-aenderungen", action="store_true")
    ap.add_argument("--aus", default=None)
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
        rows.append(r)

    gruppen: dict[str, list[dict]] = collections.defaultdict(list)
    for r in rows:
        gruppen[r.get("subject_key", "")].append(r)

    paare = []
    ketten = 0
    for key, eintraege in gruppen.items():
        p = paare_bilden(eintraege)
        if p:
            ketten += 1
        for x, y, mehrdeutig in p:
            st = unterschied(x, y)
            paare.append((x, y, st, mehrdeutig))

    zaehler = collections.Counter(p[2] for p in paare)
    mehrdeutige = sum(1 for p in paare if p[3])
    echte = [p for p in paare if p[2] not in ("unveraendert", "nur Darstellung")]
    seiten = sum(1 for p in paare if str(p[0].get("page")) != str(p[1].get("page")))

    if a.nur_aenderungen:
        paare = echte

    # ---------------------------------------------------------------- Zahlen
    L = [f"VERSIONSKETTEN  {datetime.now():%Y-%m-%d %H:%M}", "=" * 62, ""]
    if a.projekt:
        L.append(f"  nur Projekt            {a.projekt}")
    if a.datei:
        L.append(f"  nur Dateien mit        '{a.datei}'")
    L += [f"  Anforderungen betrachtet          {len(rows)}",
          f"  Ketten ueber Dokumentgrenzen      {ketten}",
          f"  Vergleichspaare                   {len(paare)}",
          f"  davon Seitenzahl verschoben       {seiten}",
          f"  davon Zuordnung mehrdeutig        {mehrdeutige}",
          "", "  Was die Paare zeigen:"]
    for k, n in zaehler.most_common():
        L.append(f"    {n:>5}  {k}")
    L += ["", "  Lesart:",
          "    unveraendert     wortgleich in beiden Fassungen",
          "    nur Darstellung  Unterschied stammt aus der Textgewinnung",
          "                     (Aufzaehlungszeichen, Kapitelnummer) —",
          "                     inhaltlich dieselbe Anforderung",
          "    Text / Wert / Verbindlichkeit geaendert  -> hier hinsehen",
          "    mehrdeutig       in einer Fassung mehrere Treffer zum selben",
          "                     Gegenstand; die Zuordnung ist nicht sicher",
          "", "Diese Datei enthaelt keine Dateinamen und keine Saetze."]

    OUT.mkdir(exist_ok=True)
    (OUT / "ketten_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
    for z in L:
        print(ascii_safe(z))

    # ----------------------------------------------------------------- Excel
    try:
        import xlsxwriter
    except ImportError:
        print("\nHINWEIS: xlsxwriter fehlt, Excel wurde nicht geschrieben.")
        return 0

    ziel = Path(a.aus) if a.aus else OUT / "ketten_voll.xlsx"
    wb = xlsxwriter.Workbook(str(ziel))
    kopf = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1,
                          "text_wrap": True, "valign": "top"})
    zelle = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    warn = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                          "bold": True})
    grau = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                          "font_color": "#888888"})
    ws = wb.add_worksheet("Versionsketten")
    spalten = ["Nr", "Status", "Zuordnung", "Datei A", "Seite A", "Verb. A",
               "Wert A", "Wortlaut A", "Datei B", "Seite B", "Verb. B",
               "Wert B", "Wortlaut B"]
    breiten = [5, 22, 12, 30, 8, 9, 14, 58, 30, 8, 9, 14, 58]
    for i, (s, b) in enumerate(zip(spalten, breiten)):
        ws.write(0, i, s, kopf)
        ws.set_column(i, i, b)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(paare), 1), len(spalten) - 1)

    rang = {"WERT GEAENDERT": 0, "VERBINDLICHKEIT GEAENDERT": 1,
            "Text geaendert": 2, "nur Darstellung": 3, "unveraendert": 4}
    paare.sort(key=lambda p: (rang.get(p[2], 9), str(p[0].get("file")),
                              str(p[0].get("page"))))
    for n, (x, y, st, md) in enumerate(paare, start=1):
        f = warn if rang.get(st, 9) <= 2 else (grau if st == "unveraendert" else zelle)
        ws.write_row(n, 0, [
            n, st, "mehrdeutig" if md else "eindeutig",
            x.get("file", ""), x.get("page", ""), x.get("binding", ""), wert(x),
            norm(x["text"]),
            y.get("file", ""), y.get("page", ""), y.get("binding", ""), wert(y),
            norm(y["text"]),
        ], f)
    wb.close()

    print()
    print("geschrieben:")
    print("  out\\ketten_zahlen.txt   nur Zahlen - darf geteilt werden")
    print(f"  {ziel.name}        mit Wortlauten - VERTRAULICH")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
