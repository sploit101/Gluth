#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pruefen_lastenheft.py — prueft, ob in den vorhandenen Lastenheften wirklich
das steht, was wir fuer die Stationsakte dort zu finden erwarten.

Hintergrund
-----------
zielfelder.csv sagt zu 32 Feldern "automatisch befuellbar". Das ist eine
Einschaetzung, abgeleitet aus den Hinweistexten der Vorlage. Dieses Skript
macht daraus eine Tatsache: Es sucht im bereits extrahierten Bestand nach
jeder Inhaltsart und zaehlt, wie oft sie vorkommt.

Ergebnis ist nicht "die Zuordnung ist richtig". Ergebnis ist:
"diese Inhaltsart existiert im Lastenheft / existiert nicht".
Ein Feld, dessen Inhaltsart gar nicht vorkommt, ist theoretisch befuellbar
und praktisch nicht - und das will man wissen, BEVOR man die Zuordnung baut.

Ehrliche Einschraenkung: ein Lastenheft von einem Kunden ist keine
Stichprobe. "Nicht gefunden" heisst nicht "gibt es nie", sondern
"koennen wir jetzt nicht zeigen".

Gesucht wird in zwei Bestaenden:
  out/requirements.jsonl   die erkannten Anforderungen (gefilterte Saetze)
  out/chunks.jsonl         der komplette Text (auch Tabellen, Kopfzeilen)
Der Unterschied ist wichtig: Zeichnungsnummern stehen meist in Tabellen und
sind damit keine Anforderung - sie tauchen nur im zweiten Bestand auf.

Aufruf
------
    python pruefen_lastenheft.py
    python pruefen_lastenheft.py --projekt A07616000
    python pruefen_lastenheft.py --beispiele 8

Ausgabe
-------
    out/pruefung_zahlen.txt        nur Zahlen        -> darf geteilt werden
    out/pruefung_beispiele.xlsx    mit Wortlauten    -> VERTRAULICH
"""

from __future__ import annotations

import argparse
import collections
import json
import re
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

# ---------------------------------------------------------------------------
# Die Inhaltsarten. Je Eintrag:
#   schluessel: (Klartext, Regex, [betroffene Felder in zielfelder.csv])
#
# Die Regex sind absichtlich weit gefasst. Lieber ein paar Fehltreffer, die
# man in den Beispielen sofort sieht, als eine Inhaltsart faelschlich als
# "nicht vorhanden" melden.
# ---------------------------------------------------------------------------

ARTEN: dict[str, tuple[str, re.Pattern, list[str]]] = {

    # ------------------------------------------------------------------
    # ZWEISPRACHIG. Die Lastenhefte im Bestand sind teils deutsch, teils
    # englisch (LH_04K_000051_V1.02 ist komplett englisch). Die erste
    # Fassung dieser Muster war nur deutsch - dadurch meldete sie fuer
    # "Anlieferform" und "Kalibrierablaeufe" faelschlich NICHT GEFUNDEN
    # und fand bei "Prozessbezeichnungen" fast nichts, obwohl stripping,
    # crimping und cutting im Dokument stehen.
    # ------------------------------------------------------------------

    "zeichnungsnummer": (
        "Zeichnungsnummern",
        re.compile(r"(zeichnungs?\s*-?\s*(nr|nummer)|\bzeichnung\b|"
                   r"drawing\s*(no|nr|number)|\bdwg\b|"
                   r"part\s*(no|nr|number)|item\s*(no|number)|\bp/n\b|"
                   r"teile?\s*-?\s*(nr|nummer)|sachnummer|"
                   r"\b\d{2,4}[-/.]\d{3,6}([-/.]\d+)?\b)", re.I),
        ["Bauteile / Zeichungsnr."],
    ),

    "bauteilname": (
        "Bauteilbezeichnungen",
        re.compile(r"\b(kundenbauteil\w*|bauteil\w*|werkst[uü]ck\w*|"
                   r"geh[aä]use\w*|cover\w*|deckel\w*|tr[aä]?[ae]ger\w*|"
                   r"component\w*|housing\w*|connector\w*|terminal\w*|"
                   r"grommet\w*|harness\w*|contact\w*|seal\w*|"
                   r"part\s+name|sub[-\s]?assembly)", re.I),
        ["Bauteile / Bauteilname", "Bauteile / Kundenbauteile"],
    ),

    "variante": (
        "Varianten und Typen",
        re.compile(r"\b(variante[nr]?\w*|ausf[uü]hrung\w*|derivat\w*|"
                   r"variant\w*|derivative\w*|"
                   r"typ\s*[a-z0-9]\b|type\s*[a-z0-9]\b|"
                   r"product\s+variant|per\s+variant|"
                   r"different\s+type)", re.I),
        ["Bauteile / Existieren Varianten", "Bauteile/Varianten (Grafik)"],
    ),

    "toleranz": (
        "Toleranzen und Grenzwerte",
        re.compile(r"(toleranz\w*|tolerance\w*|\u00b1|\+/-|"
                   r"repeat\s+accuracy|genauigkeit|accuracy|"
                   r"\bmax\.?\b|\bmin\.?\b|h[oö]chstens|mindestens|"
                   r"at\s+least|not\s+exceed|must\s+not\s+be\s+(more|less)|"
                   r"<=|>=|\u2264|\u2265)", re.I),
        ["Faehigkeitstabelle / Toleranz"],
    ),

    "faehigkeit": (
        "Faehigkeitskennwerte (MSA, Cmk, Cpk, Grr)",
        re.compile(r"\b(msa\s*\d?|c\s?mk|c\s?pk?|c\s?gk?|g\s?rr|"
                   r"r\s?&\s?r|gauge\s*r|mgf|mfu|pfu|"
                   r"capability\s+(index|test)|process\s+capability|"
                   r"machine\s+capability|measurement\s+system\s+analysis|"
                   r"prozessf[aä]higkeit|maschinenf[aä]higkeit|"
                   r"messmittelf[aä]higkeit|messf[aä]higkeit|"
                   r"wiederholbarkeit|repeatability)\b", re.I),
        ["Faehigkeitstabelle / Verfahren", "Faehigkeitstabelle / Cmk-cpk-grr"],
    ),

    "prozess": (
        "Prozessbezeichnungen",
        re.compile(r"\b(schraub\w*|verschraub\w*|f[uü]ge\w*|press\w*|"
                   r"niet\w*|kleb\w*|schwei[sß]\w*|pr[uü]f\w*|messung|"
                   r"messen|montage|handling|zuf[uü]hr\w*|dosier\w*|"
                   r"etikett\w*|lasern?|markier\w*|"
                   r"stripping|strip\w*|crimp\w*|cutting|cut\s+to|"
                   r"assembl\w*|testing|inspection|measurement|"
                   r"welding|solder\w*|glu\w*|dispens\w*|labell?\w*|"
                   r"marking|insertion|screw\w*|pressing|"
                   r"vision\s+check\w*|feeding)\b", re.I),
        ["Faehigkeitstabelle / Prozess", "Aufgabenstellung / Prozess"],
    ),

    "taktzeit": (
        "Taktzeit und Ausbringung",
        re.compile(r"\b(taktzeit|zykluszeit|ausbringung|"
                   r"cycle\s*time|machine\s+cycle|throughput|"
                   r"output\s+per|units?\s+per\s+(hour|shift)|"
                   r"st[uü]ck\s*/\s*(h|std|stunde)|teile\s*/\s*(h|std|min)|"
                   r"pcs\s*/\s*(h|hour|min)|\bopm\b|\bjph\b)\b", re.I),
        ["Auslegung * / Zeit  (offene Frage an Daniel)"],
    ),

    "anlieferform": (
        "Anlieferform der Bauteile",
        re.compile(r"\b(sch[uü]ttgut|blister|tray|werkst[uü]cktr[aä]ger|"
                   r"\bwt\b|gitterbox|palette|pallet|magazin|magazine|"
                   r"beh[aä]lter|kleinladungstr[aä]ger|\bklt\b|"
                   r"bulk|carrier|container|\bbin\b|\breel\b|\bcoil\b|"
                   r"\bbox\b|packaging|feeder|bowl\s+feeder)\b", re.I),
        ["Bauteile / Schuettgut-Blister"],
    ),

    "station": (
        "Stationsbezeichnungen",
        re.compile(r"\b(station\s*\d*|\bop\s*\d{2,}|arbeitsplatz|"
                   r"bearbeitungsstation|pr[uü]fstation|montagestation|"
                   r"workstation|work\s*station|workplace|"
                   r"assembly\s+station|test\s+station|"
                   r"input\s+station|output\s+station|"
                   r"production\s+(line|cell))\b", re.I),
        ["Stationsbeschriebung / Stationsnummer und Bezeichnung"],
    ),

    "kalibrieren": (
        "Kalibrier- und Einrichtablaeufe",
        re.compile(r"\b(kalibrier\w*|einricht\w*|referenzier\w*|"
                   r"justier\w*|nullpunkt|teach\w*|\bctc\b|"
                   r"calibrat\w*|\bsetup\b|set[-\s]up|"
                   r"reference\s+(sample|master|part)\w*|zero\s+point|"
                   r"adjustment\w*|master\s+part\w*|golden\s+sample\w*)", re.I),
        ["Kalibrier-/Einrichtteile / Beschreibung des Ablaufs"],
    ),

    "aufgabe_lang": (
        "Lange Anforderungssaetze (>120 Zeichen, MUSS/VERBOT)",
        None,                                   # Sonderfall, siehe unten
        ["Aufgabenstellung_ausfuehrlich / Beschreibung der Aufgabe",
         "Stationsbeschriebung / Beschreibung der Stationsinhalte"],
    ),
}


# Dateien, die zum Projekt selbst gehoeren und keine Kundenunterlagen sind.
# Sie standen bisher mit im Bestand und haben die Messung verfaelscht:
# LEARNBOOK.md lieferte z.B. alle fuenf Beispiele fuer "lange
# Anforderungssaetze". Werden jetzt standardmaessig ausgeschlossen.
EIGENE_DATEI = re.compile(
    r"(readme|learnbook|build_manual|anleitung_|manifest|"
    r"\.py$|\.md$|\.ipynb$|requirements\.txt$)", re.I)


# requirements.jsonl und chunks.jsonl benennen Datei und Seite nicht gleich.
# Beim ersten Lauf stand deshalb bei allen Volltext-Treffern "ohne Dateiname"
# und ueberall "Seite None". Statt einen Namen zu raten werden mehrere
# Kandidaten geprueft und der tatsaechlich belegte gemeldet.
KEYS_DATEI = ["file", "filename", "source_file", "source", "source_path",
              "path", "doc", "document", "dateiname", "datei"]
KEYS_SEITE = ["page", "page_no", "pageno", "page_number", "page_start",
              "seite", "seitenzahl"]


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


def feldname_finden(zeilen: list[dict], kandidaten: list[str]) -> str | None:
    """Welcher der Kandidaten ist in den Daten tatsaechlich belegt?"""
    for k in kandidaten:
        if any(r.get(k) not in (None, "", "None") for r in zeilen[:500]):
            return k
    return None


def wert_von(r: dict, key: str | None, kandidaten: list[str]):
    """Erst den erkannten Feldnamen, dann die anderen Kandidaten."""
    if key and r.get(key) not in (None, "", "None"):
        return r[key]
    for k in kandidaten:
        if r.get(k) not in (None, "", "None"):
            return r[k]
    return None


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def lade(pfad: Path, projekt: str | None, nur: str | None,
         mit_eigenen: bool) -> tuple[list[dict], int]:
    """Liest den Bestand. Gibt (Zeilen, Anzahl ausgeschlossener Zeilen)."""
    if not pfad.exists():
        return [], 0
    raus, weg = [], 0
    for line in pfad.open(encoding="utf-8"):
        line = line.strip()
        if not line:
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            continue
        if projekt and r.get("project_no") != projekt:
            continue
        datei = str(wert_von(r, None, KEYS_DATEI) or "")
        if not mit_eigenen and EIGENE_DATEI.search(datei):
            weg += 1
            continue
        if nur and nur.lower() not in datei.lower():
            continue
        raus.append(r)
    return raus, weg


def text_von(r: dict) -> str:
    for k in ("text", "sentence", "content", "chunk"):
        if r.get(k):
            return norm(r[k])
    return ""


def treffer(saetze: list[dict], art: str, muster) -> list[dict]:
    if art == "aufgabe_lang":
        return [r for r in saetze
                if len(text_von(r)) > 120
                and r.get("binding") in ("MUSS", "VERBOT")]
    return [r for r in saetze if muster.search(text_von(r))]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--req", default=str(OUT / "requirements.jsonl"))
    ap.add_argument("--chunks", default=str(OUT / "chunks.jsonl"))
    ap.add_argument("--projekt", default=None)
    ap.add_argument("--nur-datei", default=None, dest="nur",
                    help="nur Dateien, deren Name diesen Text enthaelt")
    ap.add_argument("--mit-eigenen", action="store_true",
                    help="auch die Projektdateien (README, LEARNBOOK, .py) "
                         "mitzaehlen - normalerweise NICHT gewollt")
    ap.add_argument("--beispiele", type=int, default=5,
                    help="wie viele Beispielsaetze je Inhaltsart")
    ap.add_argument("--aus-txt", dest="aus_txt", default=None,
                    help="Zieldatei fuer die Zahlen (Standard: "
                         "out/pruefung_zahlen.txt)")
    ap.add_argument("--aus-xlsx", dest="aus_xlsx", default=None,
                    help="Zieldatei fuer die Beispiele (Standard: "
                         "out/pruefung_beispiele.xlsx)")
    a = ap.parse_args()

    reqs, weg_req = lade(Path(a.req), a.projekt, a.nur, a.mit_eigenen)
    chunks, weg_chk = lade(Path(a.chunks), a.projekt, a.nur, a.mit_eigenen)

    if not reqs and not chunks:
        print("FEHLER: weder", ascii_safe(a.req), "noch",
              ascii_safe(a.chunks), "gefunden oder beide leer.")
        print("Erst 'python extract.py' und 'python requirements.py' laufen lassen.")
        return 2

    k_datei_req = feldname_finden(reqs, KEYS_DATEI)
    k_datei_chk = feldname_finden(chunks, KEYS_DATEI)
    k_seite_req = feldname_finden(reqs, KEYS_SEITE)
    k_seite_chk = feldname_finden(chunks, KEYS_SEITE)

    def datei_von(r):
        return wert_von(r, k_datei_req or k_datei_chk, KEYS_DATEI)

    def seite_von(r):
        return wert_von(r, k_seite_req or k_seite_chk, KEYS_SEITE)

    dateien_req = {datei_von(r) for r in reqs if datei_von(r)}
    dateien_chk = {datei_von(r) for r in chunks if datei_von(r)}

    ergebnis = []
    beispiele: dict[str, list[tuple[str, dict]]] = {}
    quellen: dict[str, collections.Counter] = {}

    for schluessel, (klartext, muster, felder) in ARTEN.items():
        tr_req = treffer(reqs, schluessel, muster)
        tr_chk = treffer(chunks, schluessel, muster) if schluessel != "aufgabe_lang" else []
        n_req, n_chk = len(tr_req), len(tr_chk)

        # Urteil: wonach wir die Inhaltsart als vorhanden ansehen
        if n_req >= 5 or n_chk >= 10:
            urteil = "VORHANDEN"
        elif n_req >= 1 or n_chk >= 1:
            urteil = "selten"
        else:
            urteil = "NICHT GEFUNDEN"

        ergebnis.append((schluessel, klartext, n_req, n_chk, urteil, felder))
        quellen[schluessel] = collections.Counter(
            str(datei_von(r) or "ohne Dateiname") for r in tr_req + tr_chk)

        # Entdoppeln: in der ersten Fassung belegte derselbe Satz bis zu
        # vier der acht Beispielplaetze.
        b, gesehen = [], set()
        for bestand, liste in (("Anforderung", tr_req), ("Volltext", tr_chk)):
            for r in liste:
                if len(b) >= a.beispiele:
                    break
                kurz = re.sub(r"\W+", " ", text_von(r)).strip().lower()[:160]
                if kurz in gesehen:
                    continue
                gesehen.add(kurz)
                b.append((bestand, r))
        beispiele[schluessel] = b

    # ----------------------------------------------------------- Zahlendatei
    L = [f"PRUEFUNG ZIELFELDER GEGEN LASTENHEFT  {datetime.now():%Y-%m-%d %H:%M}",
         "=" * 72, ""]
    if a.projekt:
        L.append(f"  nur Projekt                 {a.projekt}")
    if a.nur:
        L.append(f"  nur Dateien mit             '{a.nur}'")
    L += [f"  Anforderungen betrachtet    {len(reqs)}",
          f"  Textabschnitte betrachtet   {len(chunks)}",
          f"  Dokumente (Anforderungen)   {len(dateien_req)}",
          f"  Dokumente (Volltext)        {len(dateien_chk)}"]
    sprachen = collections.Counter(
        str(r.get("lang") or "?") for r in reqs)
    if sprachen:
        L.append("  Sprache der Anforderungen   " + ", ".join(
            f"{k}: {n}" for k, n in sprachen.most_common(4)))
    L += [f"  Feldname Datei              "
          f"{k_datei_req or '-'} (Anf.) / {k_datei_chk or '-'} (Volltext)",
          f"  Feldname Seite              "
          f"{k_seite_req or '-'} (Anf.) / {k_seite_chk or '-'} (Volltext)"]
    if a.mit_eigenen:
        L.append("  ACHTUNG: eigene Projektdateien sind MITGEZAEHLT")
    else:
        L.append(f"  eigene Projektdateien weg   {weg_req} Anf. / "
                 f"{weg_chk} Abschnitte")
    # Wie viele Treffer haben keine Seitenzahl? Das betrifft direkt die
    # Zusage "Datei und Seite nachvollziehbar".
    # Derselbe Satz unter mehreren Dateinamen = Dokument liegt doppelt im
    # Bestand (z.B. einzeln UND in einem Archiv). Verdoppelt jede Zahl.
    nach_text: dict[str, set] = collections.defaultdict(set)
    for r in reqs:
        t = text_von(r)
        if len(t) > 60:
            nach_text[t].add(str(datei_von(r) or ""))
    doppelt = sum(1 for t, d in nach_text.items() if len(d) > 1)
    ohne_seite = sum(1 for r in reqs if seite_von(r) is None)
    if reqs:
        L.append(f"  Anforderungen ohne Seite    {ohne_seite} "
                 f"({100*ohne_seite/len(reqs):.0f} %)")
        L.append(f"  Saetze in mehreren Dateien  {doppelt} "
                 f"(Hinweis auf doppelte Dokumente)")
    L += ["",
          "  Inhaltsart                                 Anf.  Volltext  Urteil",
          "  " + "-" * 68]
    for _s, klartext, n_req, n_chk, urteil, _f in ergebnis:
        L.append(f"  {klartext[:40]:<40} {n_req:>6} {n_chk:>9}  {urteil}")
    L += ["", "  Lesart:",
          "    VORHANDEN       Inhaltsart kommt belegbar vor -> Feld ist real",
          "                    befuellbar, Zuordnung lohnt sich",
          "    selten          vereinzelt gefunden -> Beispiele ansehen, oft",
          "                    sind es Fehltreffer",
          "    NICHT GEFUNDEN  Inhaltsart kommt in DIESEN Dokumenten nicht vor.",
          "                    Das Feld bleibt vorerst leer. Kein Fehler des",
          "                    Programms, sondern eine Eigenschaft der Quelle.",
          "",
          "  Wichtig: ein Lastenheft von einem Kunden ist keine Stichprobe.",
          "  'Nicht gefunden' heisst 'koennen wir jetzt nicht zeigen',",
          "  nicht 'gibt es nie'.",
          "",
          "Diese Datei enthaelt keine Dateinamen und keine Saetze."]

    OUT.mkdir(exist_ok=True)
    txt_ziel = Path(a.aus_txt) if a.aus_txt else OUT / "pruefung_zahlen.txt"
    txt_ziel.parent.mkdir(parents=True, exist_ok=True)
    txt_ziel.write_text("\n".join(L), encoding="utf-8")
    for z in L:
        print(ascii_safe(z))

    # ----------------------------------------------------------- Beispiele
    try:
        import xlsxwriter
    except ImportError:
        print()
        print("HINWEIS: xlsxwriter fehlt, Beispieldatei wurde nicht geschrieben.")
        print("  .\\.venv\\Scripts\\python.exe -m pip install xlsxwriter")
        return 0

    ziel = Path(a.aus_xlsx) if a.aus_xlsx else OUT / "pruefung_beispiele.xlsx"
    ziel.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb = xlsxwriter.Workbook(str(ziel))
    except PermissionError:
        print()
        print("FEHLER: ", ascii_safe(ziel.name), "ist in Excel offen.")
        print("        Excel schliessen und nochmal starten.")
        return 1

    kopf = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1,
                          "text_wrap": True, "valign": "top"})
    zelle = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    stark = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                           "bold": True})
    t1 = wb.add_format({"bold": True, "font_size": 13})
    t2 = wb.add_format({"bold": True})

    # Blatt 1: Uebersicht mit Platz fuer dein Urteil
    ws = wb.add_worksheet("Urteil")
    for i, b in enumerate([40, 9, 10, 16, 14, 46, 40]):
        ws.set_column(i, i, b)
    ws.write(0, 0, "Pruefung der Zielfelder gegen die vorhandenen Lastenhefte", t1)
    ws.write(1, 0, "Spalte 'Dein Urteil' ausfuellen: ja / nein / unklar", t2)
    kopfzeile = ["Inhaltsart", "Anf.", "Volltext", "Urteil Programm",
                 "Dein Urteil", "Betroffene Felder in zielfelder.csv",
                 "Deine Bemerkung"]
    for i, s in enumerate(kopfzeile):
        ws.write(3, i, s, kopf)
    ws.freeze_panes(4, 0)
    for n, (_s, klartext, n_req, n_chk, urteil, felder) in enumerate(ergebnis):
        f = stark if urteil == "NICHT GEFUNDEN" else zelle
        ws.write_row(4 + n, 0, [klartext, n_req, n_chk, urteil, "",
                                "\n".join(felder), ""], f)

    # Blatt 2: woher die Treffer kommen. Wichtig fuer das Urteil - wenn alle
    # Treffer aus einem einzigen Dokument stammen, ist die Inhaltsart nicht
    # "im Lastenheft vorhanden", sondern "in diesem einen Dokument".
    ws = wb.add_worksheet("Quellen")
    for i, b in enumerate([40, 46, 10]):
        ws.set_column(i, i, b)
    for i, s in enumerate(["Inhaltsart", "Datei", "Treffer"]):
        ws.write(0, i, s, kopf)
    ws.freeze_panes(1, 0)
    z = 1
    for schluessel, (klartext, _m, _f) in ARTEN.items():
        for datei, anzahl in quellen.get(schluessel,
                                         collections.Counter()).most_common(5):
            ws.write_row(z, 0, [klartext, datei, anzahl], zelle)
            z += 1
        if not quellen.get(schluessel):
            ws.write_row(z, 0, [klartext, "keine Treffer", 0], stark)
            z += 1

    # Je Inhaltsart ein Blatt mit den Beispielsaetzen
    for schluessel, (klartext, _m, _f) in ARTEN.items():
        name = schluessel[:31]
        ws = wb.add_worksheet(name)
        for i, b in enumerate([12, 30, 8, 14, 96]):
            ws.set_column(i, i, b)
        for i, s in enumerate(["Bestand", "Datei", "Seite", "Verbindlichkeit",
                               "Wortlaut"]):
            ws.write(0, i, s, kopf)
        ws.freeze_panes(1, 0)
        for n, (bestand, r) in enumerate(beispiele.get(schluessel, []), start=1):
            ws.write_row(n, 0, [bestand, datei_von(r) or "", seite_von(r) or "",
                                r.get("binding", ""), text_von(r)], zelle)
        if not beispiele.get(schluessel):
            ws.write(1, 0, "keine Treffer", stark)

    wb.close()

    print()
    print("geschrieben:")
    print(f"  {txt_ziel}      nur Zahlen - darf geteilt werden")
    print(f"  {ziel}  mit Wortlauten - VERTRAULICH")
    print()
    print("Naechster Schritt: pruefung_beispiele.xlsx oeffnen, Blatt 'Urteil',")
    print("Spalte 'Dein Urteil' ausfuellen (ja / nein / unklar).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
