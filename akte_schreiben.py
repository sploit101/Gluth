#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
akte_schreiben.py — schreibt die Zuordnungen in eine Stationsakte.

Das ist der letzte Schritt der Erstbefuellung. Eingang ist
out/zuordnung.jsonl, Ausgang eine ausgefuellte PowerPoint.

Vier Regeln ohne Ausnahme
-------------------------
1. Das Original wird NIE angefasst. Es wird kopiert, und geschrieben wird
   ausschliesslich in die Kopie.
2. JEDE Aenderung wird protokolliert: Folie, Feld, Kennung, alter Wert,
   neuer Wert, Quelle. Ohne Protokoll ist eine automatische Aenderung in
   einem Haftungsdokument nicht vertretbar.
3. Farbe ist nie der einzige Hinweis. Jeder geschriebene Eintrag traegt
   zusaetzlich seine Kennung und die Quellenangabe im Text. Wer die Datei
   schwarzweiss ausdruckt, sieht trotzdem, was automatisch entstand.
4. Vor dem Suchen werden Textteile zusammengefuehrt. PowerPoint zerlegt
   Saetze unsichtbar; wer das uebersieht, findet vorhandenen Text nicht.

Was dieses Skript bewusst NICHT tut
-----------------------------------
* Es ordnet nichts einer Station zu. Im Lastenheft steht nicht, welcher
  Satz zu welcher Station gehoert. Mit --station wird eine Angabe des
  MENSCHEN eingetragen und als solche vermerkt.
* Es fuellt keine Bildplatzhalter. Welches Bild wohin gehoert, laesst sich
  nicht verlaesslich entscheiden.
* Es gruppiert die Faehigkeitstabelle nicht. Jede Fundstelle wird zu einer
  eigenen Zeile, weil nicht bekannt ist, welche Toleranz zu welchem Prozess
  gehoert. Das Zusammenfuehren bleibt Aufgabe eines Menschen.

Aufruf
------
    python akte_schreiben.py
    python akte_schreiben.py --station "Station 20 Abisolieren"
    python akte_schreiben.py --nur-sichere

Ausgabe
-------
    out/Stationsakte_gefuellt.pptx   die ausgefuellte Akte   VERTRAULICH
    out/schreibprotokoll.csv         jede Aenderung einzeln  VERTRAULICH
    out/schreiben_zahlen.txt         nur Zahlen -> darf geteilt werden
"""

from __future__ import annotations

import argparse
import collections
import copy
import csv
import json
import re
import shutil
from datetime import datetime
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Cm, Pt
except ImportError:                                            # pragma: no cover
    raise SystemExit("python-pptx fehlt:  python -m pip install python-pptx")

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

FARBE_NEU = RGBColor(0x1A, 0x5C, 0x8A)      # dunkelblau: automatisch gefuellt
FARBE_QUELLE = RGBColor(0x77, 0x77, 0x77)   # grau: Kennung und Herkunft
FARBE_LUECKE = RGBColor(0x9A, 0x64, 0x10)   # braun: offene Stelle

MARKE_LUECKE = "[offen: {}]"
# Fuer Vorschlaege, die es gibt, die aber wegen --nur-sichere nicht
# geschrieben wurden. NICHT "[offen]" - das hiesse "nichts gefunden".
MARKE_ZURUECK = "[{n} Vorschläge zur Prüfung – siehe zuordnung.xlsx]"
MARKE_WEITERE = "+ {n} weitere Vorschläge zur Prüfung – siehe zuordnung.xlsx"

FREIGABE_JA = {"ja"}
FREIGABE_NEIN = {"nein"}


def freigabe_schluessel(kennung, folienart, feld_id, text, datei, seite, station) -> tuple:
    """Identify one proposed mapping in the exact draft that was reviewed."""
    return tuple(str(value or "").strip() for value in
                 (kennung, folienart, feld_id, text, datei, seite, station))


def freigaben_laden(pfad: Path, zuordnungen: list[dict] | None = None) -> tuple[dict, str]:
    """
    Liest die Spalte "Freigabe" aus einer geprueften zuordnung.xlsx.
    Rueckgabe: ({schluessel: "ja"|"nein"}, Fehlertext oder "").
    If mappings are supplied, every row must match the current draft exactly.
    """
    try:
        import openpyxl
    except ImportError:
        return {}, "openpyxl fehlt"
    try:
        wb = openpyxl.load_workbook(str(pfad), read_only=True, data_only=True)
    except Exception as e:
        return {}, str(e)
    try:
        ws = wb["Zuordnung"] if "Zuordnung" in wb.sheetnames else wb.active
        zeilen = ws.iter_rows(values_only=True)
        kopf = [str(x or "").strip() for x in next(zeilen, [])]
        noetig = ["Kennung", "Folienart", "Feld", "Wortlaut", "Datei",
                  "Seite", "Station", "Freigabe"]
        if any(n not in kopf for n in noetig):
            return {}, "Spalten fehlen: " + ", ".join(
                n for n in noetig if n not in kopf)
        i = {n: kopf.index(n) for n in noetig}
        fr = {}
        alle = []
        for z in zeilen:
            if not z or not any(v is not None for v in z):
                continue
            if max(i.values()) >= len(z):
                return {}, "Eine Zeile des Pruefblatts ist unvollstaendig."
            key = freigabe_schluessel(*(z[i[name]] for name in
                                       ("Kennung", "Folienart", "Feld", "Wortlaut",
                                        "Datei", "Seite", "Station")))
            alle.append(key)
            wert = str(z[i["Freigabe"]] or "").strip().lower()
            if wert in FREIGABE_JA:
                wert = "ja"
            elif wert in FREIGABE_NEIN:
                wert = "nein"
            elif wert:
                return {}, f"Ungueltige Freigabe in Zeile {len(alle) + 1}: {wert[:30]}"
            else:
                continue
            if key in fr:
                return {}, "Doppelte Freigabe fuer dieselbe Zuordnung."
            fr[key] = wert
        if zuordnungen is not None:
            erwartet = [freigabe_schluessel(*(r.get(name, "") for name in
                        ("kennung", "folienart", "feld_id", "text", "datei",
                         "seite", "station"))) for r in zuordnungen]
            if collections.Counter(alle) != collections.Counter(erwartet):
                return {}, "Das Pruefblatt gehoert nicht zur aktuellen Zuordnung. Bitte das neue Blatt herunterladen."
        return fr, ""
    finally:
        wb.close()


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


# Aufzaehlungszeichen aus Symbol-/Wingdings-Schriften. Die PDF-Extraktion
# liefert sie als Zeichen aus dem privaten Unicode-Bereich (U+F0B7 usw.).
# In jeder normalen Schrift erscheinen sie als leeres Kaestchen. Ersetzt wird
# NUR das Zeichen durch einen gewoehnlichen Punkt - kein Wort, keine Zahl.
# Im Protokoll und in zuordnung.xlsx bleibt der Rohtext unveraendert.
LISTENZEICHEN = {"": "•", "": "•", "": "•",
                 "": "•", "": "•", "": "•"}


def anzeige(text: str) -> str:
    t = norm(text)
    for alt, neu in LISTENZEICHEN.items():
        t = t.replace(alt, neu)
    return t


def text_zusammengefuehrt(tf) -> str:
    """Regel 4: runs zusammenfuehren, bevor irgendetwas gesucht wird."""
    return "\n".join("".join(r.text for r in p.runs) or p.text or ""
                     for p in tf.paragraphs)


def vorlage_finden(angabe: str | None) -> Path | None:
    if angabe:
        p = Path(angabe)
        return p if p.exists() else None
    fest = HERE / "vorlage_stationsakte.pptx"
    if fest.exists():
        return fest
    k = sorted(HERE.glob("*.pptx"), key=lambda p: p.stat().st_mtime,
               reverse=True)
    return k[0] if k else None


def katalog_laden(pfad: Path) -> dict:
    if not pfad.exists():
        return {}
    kat = {}
    with pfad.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";"):
            kat[(r["folienart"], r["feld_id"])] = r
    return kat


# ---------------------------------------------------------------------------
# Schreibhilfen
# ---------------------------------------------------------------------------

def absatz_schreiben(tf, text: str, farbe, groesse=None, erster=False):
    """Einen Absatz anlegen oder den ersten ueberschreiben."""
    p = tf.paragraphs[0] if erster else tf.add_paragraph()
    for r in list(p.runs):                      # vorhandene runs entfernen
        r._r.getparent().remove(r._r)
    r = p.add_run()
    r.text = text
    r.font.color.rgb = farbe
    if groesse:
        r.font.size = Pt(groesse)
    return p


def tabellenzeile_anhaengen(table):
    """
    python-pptx kann keine Tabellenzeile anlegen. Deshalb wird die letzte
    Zeile im XML kopiert, angehaengt und geleert. Formatierung bleibt dadurch
    erhalten - was gewollt ist, die Akte soll nicht auseinanderfallen.
    """
    tbl = table._tbl
    neu = copy.deepcopy(tbl.tr_lst[-1])
    tbl.append(neu)
    zeile = table.rows[len(table.rows) - 1]
    for c in zeile.cells:
        c.text_frame.clear()
    return len(table.rows) - 1


def erste_leere_zeile(table, ab=1) -> int | None:
    for ri in range(ab, len(table.rows)):
        if all(not norm(c.text) for c in table.rows[ri].cells):
            return ri
    return None


def zelle_schreiben(table, ri, ci, text, farbe=FARBE_NEU):
    if ci >= len(table.columns) or ri >= len(table.rows):
        return False
    c = table.rows[ri].cells[ci]
    c.text_frame.clear()
    absatz_schreiben(c.text_frame, text, farbe, groesse=10, erster=True)
    return True


def tabelle_finden(folie, kopf_stichwort: str):
    for sh in folie.shapes:
        if sh.has_table and len(sh.table.rows):
            kopf = " ".join(norm(c.text).lower()
                            for c in sh.table.rows[0].cells)
            if kopf_stichwort.lower() in kopf:
                return sh.table
    return None


# Schriftgroessen. Fest gesetzt, damit die Platzberechnung unten ueberhaupt
# stimmen kann - geerbte Groessen sind von aussen nicht zuverlaessig lesbar.
PT_TEXT = 11
PT_QUELLE = 8

# Faustwerte fuer die Platzberechnung bei diesen Groessen. Lieber etwas zu
# vorsichtig: eine halbleere Folie ist besser als Text, der unten herauslaeuft.
CM_PRO_ZEICHEN = 0.185          # mittlere Zeichenbreite bei 11 pt (Arial,
                                # gemessen 0,174 - plus Rand)
CM_PRO_ZEILE = 0.48             # Zeilenhoehe bei 11 pt (gemessen 0,46)
ZEILE_QUELLE = 0.75             # eine 8-pt-Zeile kostet so viel wie 0,75 Zeilen
SICHERHEIT = 0.92               # Rand gegen Schaetzfehler


def platz_budget(sh) -> tuple[int, float]:
    """
    Wie viel passt in diese Form?

    Gibt (Zeichen pro Zeile, verfuegbare Zeilen) zurueck. python-pptx kann
    nicht messen, wie PowerPoint umbricht - deshalb wird aus Breite, Hoehe und
    Schriftgroesse geschaetzt. Das ist ungenau, aber es verhindert den Fehler,
    der ohne jede Schaetzung sicher auftritt: Text laeuft unten aus der Folie.
    """
    try:
        breite = (sh.width or 0) / 360000.0
        hoehe = (sh.height or 0) / 360000.0
    except Exception:
        return 90, 12.0
    if breite <= 0 or hoehe <= 0:
        return 90, 12.0
    je_zeile = max(20, int((breite - 0.6) / CM_PRO_ZEICHEN))
    zeilen = max(1.0, ((hoehe - 0.4) / CM_PRO_ZEILE) * SICHERHEIT)
    return je_zeile, zeilen


def kosten(text: str, je_zeile: int) -> float:
    """Wie viele Zeilen kostet ein Eintrag samt seiner Quellenzeile?"""
    import math
    return math.ceil(max(1, len(text)) / je_zeile) + ZEILE_QUELLE


def rahmen_vorbereiten(tf):
    """Umbruch an, und PowerPoints eigene Verkleinerung als zweites Netz."""
    try:
        from pptx.enum.text import MSO_AUTO_SIZE
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass


def ist_klein(sh) -> bool:
    """
    Kleine Kaestchen (Stunden, Budget, Nr./Index, Ja/nein) sind in der
    Vorlage nur wenige Zentimeter breit. Ein langer Satz darin laeuft
    sichtbar ueber den Rand. Ab hier wird gekuerzt.
    """
    try:
        return sh.width is not None and sh.width < 5 * 360000
    except Exception:
        return False


def kuerzen(text: str, n: int) -> str:
    text = norm(text)
    return text if len(text) <= n else text[: n - 1].rstrip() + "\u2026"


# Tabellen-Hoehenbudget. Zellen werden mit 10 pt geschrieben.
FUSS_RESERVE_CM = 1.3          # Fusszeile der Gluth-Vorlage freihalten
                               # (Linie liegt 0,95 cm ueber dem Rand)
MIN_ZEILE_CM = 1.03            # Zeilenhoehe der Vorlage
CM_PRO_ZEICHEN_10PT = 0.18
CM_PRO_ZEILE_10PT = 0.42
ZELLRAND_CM = 0.30             # Innenabstand oben + unten


def zeilenhoehe_cm(texte: dict, breiten: list[float]) -> float:
    """
    Geschaetzte Hoehe einer Tabellenzeile: die hoechste Zelle bestimmt sie.
    Wie beim Textfeld eine Schaetzung - bewusst vorsichtig.
    """
    import math
    hoehe = MIN_ZEILE_CM
    for ci, text in texte.items():
        if ci >= len(breiten) or not text:
            continue
        je_zeile = max(6, int((breiten[ci] - 0.25) / CM_PRO_ZEICHEN_10PT))
        zeilen = sum(max(1, math.ceil(len(teil) / je_zeile))
                     for teil in str(text).split("\n"))
        hoehe = max(hoehe, zeilen * CM_PRO_ZEILE_10PT + ZELLRAND_CM)
    return hoehe


def tabelle_fortsetzen(prs, quell_folie, quell_shape, teil: int):
    """
    Legt eine neue Folie derselben Folienart an und kopiert die Tabelle
    dorthin - mit Kopfzeile und genau einer leeren Datenzeile. Formatierung
    bleibt erhalten, weil die Tabelle im XML kopiert wird.
    Gibt (folie, tabelle) zurueck oder None.
    """
    try:
        neu = prs.slides.add_slide(quell_folie.slide_layout)
        el = copy.deepcopy(quell_shape._element)
        # eindeutige Form-ID auf der neuen Folie
        ids = [s.shape_id for s in neu.shapes]
        el.nvGraphicFramePr.cNvPr.id = (max(ids) if ids else 1) + 1
        neu.shapes._spTree.append(el)
        tab = None
        for s in neu.shapes:
            if s.has_table and s.shape_id == el.nvGraphicFramePr.cNvPr.id:
                tab = s.table
        if tab is None:
            return None
        tbl = tab._tbl
        for tr in list(tbl.tr_lst)[2:]:
            tbl.remove(tr)
        if len(tab.rows) < 2:
            tabellenzeile_anhaengen(tab)
        for c in tab.rows[1].cells:
            c.text_frame.clear()
        # Titel uebernehmen und kennzeichnen
        alt_titel = None
        for s in quell_folie.shapes:
            if s.is_placeholder and s.placeholder_format.idx == 0 \
                    and s.has_text_frame:
                alt_titel = norm(s.text_frame.text)
        for s in neu.shapes:
            if s.is_placeholder and s.placeholder_format.idx == 0 \
                    and s.has_text_frame:
                absatz_schreiben(s.text_frame,
                                 f"{alt_titel or 'Fähigkeitstabelle'} "
                                 f"(Fortsetzung {teil})",
                                 FARBE_QUELLE, erster=True)
        return neu, tab
    except Exception:                                          # pragma: no cover
        return None


def bekannte_beispiele_entfernen(prs, merken) -> None:
    """Remove only identifiable example values from the supplied sample template."""
    beispiele = (
        "Schraubprozess mit max. 38Nm Drehmoment",
        "Kalibrierabläufe mit Rahmenbedingungen xy",
        "CTC Ablauf nach Schema xy",
    )
    for nr, slide in enumerate(prs.slides, start=1):
        labels = {norm(shape.text) for shape in slide.shapes
                  if shape.has_text_frame and shape.text.strip()}
        bauteil_demo = all(f"Kundenbauteil Cover Typ {letter}" in labels
                           for letter in "ABCD")
        for shape in list(slide.shapes):
            if shape.has_text_frame and shape.text.strip():
                old = norm(shape.text)
                if old.startswith(beispiele) or (bauteil_demo and
                        old.startswith("Kundenbauteil Cover Typ ")):
                    shape.text_frame.clear()
                    merken(nr, slide.slide_layout.name, shape.name, "-",
                           "Beispielinhalt der Vorlage geleert", old, "", "")
            if bauteil_demo and shape.name.startswith("Ellipse"):
                shape._element.getparent().remove(shape._element)
                merken(nr, slide.slide_layout.name, shape.name, "-",
                       "Beispielgrafik der Vorlage entfernt", "Ellipse", "", "")
            if shape.has_table and len(shape.table.rows):
                header = " ".join(norm(cell.text).lower()
                                  for cell in shape.table.rows[0].cells)
                if "version" not in header or "verantwortlich" not in header:
                    continue
                for row in list(shape.table.rows)[1:]:
                    old = " | ".join(norm(cell.text) for cell in row.cells)
                    if "Änderungen xy" in old or "Dokument erstellt;" in old:
                        for cell in row.cells:
                            cell.text_frame.clear()
                        merken(nr, slide.slide_layout.name, shape.name, "-",
                               "Beispielversion der Vorlage geleert", old, "", "")


def platzhalter_finden(folie, feld_id: str):
    idx = int(feld_id[2:]) if feld_id.startswith("ph") else None
    if idx is None:
        return None
    for sh in folie.shapes:
        if sh.is_placeholder and sh.placeholder_format.idx == idx:
            return sh
    return None


def folien_nach_art(prs) -> dict:
    d = collections.defaultdict(list)
    for nr, s in enumerate(prs.slides, start=1):
        d[s.slide_layout.name].append((nr, s))
    return d


def layout_nach_name(prs, name: str):
    for L in prs.slide_layouts:
        if L.name == name:
            return L
    return None


# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--vorlage", default=None)
    ap.add_argument("--zuordnung", default=str(OUT / "zuordnung.jsonl"))
    ap.add_argument("--katalog", default=str(HERE / "zielfelder.csv"))
    ap.add_argument("--aus", default=str(OUT / "Stationsakte_gefuellt.pptx"))
    ap.add_argument("--station", default=None,
                    help="Stationsbezeichnung. Wird als Angabe des MENSCHEN "
                         "eingetragen und so vermerkt.")
    ap.add_argument("--nur-sichere", action="store_true", dest="nur_sichere",
                    help="nur Zuordnungen ohne Pruefvermerk schreiben")
    ap.add_argument("--max-je-feld", type=int, default=12, dest="max_je_feld",
                    help="hoechstens so viele Eintraege je Textfeld")
    ap.add_argument("--freigabe", default=None,
                    help="gepruefte zuordnung.xlsx mit Spalte 'Freigabe' "
                         "(ja/nein), passend zur aktuellen Zuordnung")
    ap.add_argument("--beispiele-behalten", action="store_true",
                    dest="beispiele_behalten",
                    help="bekannte Beispielinhalte der Vorlage nicht entfernen")
    ap.add_argument("--verantwortlich", default="automatisch vorbefuellt",
                    help="Eintrag in der Spalte Verantwortlich")
    a = ap.parse_args()

    vorlage = vorlage_finden(a.vorlage)
    if vorlage is None:
        print("FEHLER: keine Vorlage gefunden. --vorlage angeben.")
        return 2

    zpfad = Path(a.zuordnung)
    if not zpfad.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(zpfad))
        print("Erst 'python zuordnen.py' laufen lassen.")
        return 2

    zuord = [json.loads(l) for l in zpfad.open(encoding="utf-8") if l.strip()]

    # ------------------------------------------- Freigaben eines Menschen
    # Wer zuordnung.xlsx geprueft und in der Spalte "Freigabe" ja/nein
    # eingetragen hat, gibt sie mit --freigabe an. "ja" = gilt als geprueft, wird ohne PRUEFEN
    # geschrieben. "nein" = wird nicht geschrieben. Leer = wie bisher.
    # Das ganze Blatt muss zur aktuellen Zuordnung passen. Eine alte Datei
    # darf nicht stillschweigend Entscheidungen fuer einen neuen Lauf liefern.
    freigabe_pfad = Path(a.freigabe) if a.freigabe else None
    freigabe_info = {"datei": "", "ja": 0, "nein": 0, "nicht_gefunden": 0}
    abgelehnt = []
    if freigabe_pfad and not freigabe_pfad.exists():
        print("FEHLER: Freigabe-Datei nicht gefunden:", ascii_safe(freigabe_pfad))
        return 2
    if freigabe_pfad:
        fr, fehler = freigaben_laden(freigabe_pfad, zuord)
        if fehler:
            print("FEHLER beim Lesen der Freigaben:", ascii_safe(fehler))
            print("Datei in Excel geschlossen? Spalten 'Folienart', 'Feld',")
            print("'Wortlaut' und 'Freigabe' vorhanden?")
            return 2
        freigabe_info["datei"] = freigabe_pfad.name
        benutzt = set()
        for z in zuord:
            k = freigabe_schluessel(*(z.get(name, "") for name in
                    ("kennung", "folienart", "feld_id", "text", "datei",
                     "seite", "station")))
            wert = fr.get(k)
            if wert == "ja":
                z["freigabe"] = "ja"
                z["pruefen"] = ""
                freigabe_info["ja"] += 1
                benutzt.add(k)
            elif wert == "nein":
                z["freigabe"] = "nein"
                freigabe_info["nein"] += 1
                benutzt.add(k)
        freigabe_info["nicht_gefunden"] = len(set(fr) - benutzt)
        abgelehnt = [z for z in zuord if z.get("freigabe") == "nein"]
        zuord = [z for z in zuord if z.get("freigabe") != "nein"]

    # Zurueckgehaltene Vorschlaege werden gezaehlt, nicht verschwiegen:
    # Das Feld bekommt einen Hinweis "N Vorschlaege zur Pruefung" statt eines
    # irrefuehrenden "[offen]".
    zurueck: collections.Counter = collections.Counter()
    if a.nur_sichere:
        for z in zuord:
            if z.get("pruefen"):
                zurueck[(z["folienart"], z["feld_id"])] += 1
        zuord = [z for z in zuord if not z.get("pruefen")]
    if not zuord and not zurueck:
        print("Keine Zuordnungen zu schreiben.")
        return 1

    kat = katalog_laden(Path(a.katalog))

    # ------------------------------------------------- Regel 1: kopieren
    OUT.mkdir(exist_ok=True)
    ziel = Path(a.aus)
    if ziel.resolve() == vorlage.resolve():
        print("FEHLER: Ausgabe darf nicht die Vorlage sein.")
        return 2
    if ziel.exists():
        print("FEHLER: Ausgabe existiert bereits:", ascii_safe(ziel))
        return 2
    try:
        shutil.copy2(vorlage, ziel)
    except PermissionError:
        print("FEHLER:", ascii_safe(ziel.name), "ist in PowerPoint offen.")
        return 1
    prs = Presentation(str(ziel))

    protokoll = []

    def merken(folie, folienart, feld_id, kennung, aktion, alt, neu, quelle):
        protokoll.append({
            "zeit": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
            "folie": folie, "folienart": folienart, "feld_id": feld_id,
            "kennung": kennung, "aktion": aktion,
            "alt": norm(alt)[:160], "neu": norm(neu)[:300],
            "quelle": quelle,
        })

    if not a.beispiele_behalten:
        bekannte_beispiele_entfernen(prs, merken)

    for z in abgelehnt:
        merken("-", z["folienart"], z["feld_id"], z.get("kennung", "-"),
               "abgelehnt (Freigabe nein), nicht geschrieben", "", z["text"],
               f"{z.get('datei','')} S.{z.get('seite','')}")

    nach_feld: dict[tuple, list] = collections.defaultdict(list)
    for z in zuord:
        nach_feld[(z["folienart"], z["feld_id"])].append(z)

    arten = folien_nach_art(prs)
    neue_folien = []
    geschrieben, uebersprungen = 0, []

    # ============================================================ Textfelder
    fortsetzungen = set()        # Folien, die nur Fortsetzung sind

    textfelder = sorted(set(nach_feld) | {k for k in zurueck
                                          if k[1].startswith("ph")})
    for (folienart, feld_id) in textfelder:
        if not feld_id.startswith("ph"):
            continue
        liste = list(nach_feld.get((folienart, feld_id), []))
        n_zurueck = zurueck.get((folienart, feld_id), 0)

        ziel_folie, nr = None, None
        for n, s_ in arten.get(folienart, []):
            sh = platzhalter_finden(s_, feld_id)
            if sh is not None:
                ziel_folie, nr = s_, n
                break

        # Keine passende Folie vorhanden -> aus der Folienart neu anlegen
        if ziel_folie is None:
            L = layout_nach_name(prs, folienart)
            if L is None:
                uebersprungen.append((folienart, feld_id,
                                      "Folienart nicht im Master"))
                continue
            ziel_folie = prs.slides.add_slide(L)
            nr = len(prs.slides)
            arten[folienart].append((nr, ziel_folie))
            neue_folien.append((nr, folienart))
            merken(nr, folienart, "-", "-", "Folie angelegt", "", folienart, "")

        sh = platzhalter_finden(ziel_folie, feld_id)
        if sh is None or not sh.has_text_frame:
            uebersprungen.append((folienart, feld_id,
                                  "Platzhalter fehlt oder kein Textfeld"))
            continue

        # Nur zurueckgehaltene Vorschlaege, nichts zu schreiben: ehrlicher
        # Hinweis statt "[offen]". "[offen]" hiesse "nichts gefunden" - das
        # waere hier falsch.
        if not liste:
            alt = text_zusammengefuehrt(sh.text_frame)
            sh.text_frame.clear()
            marke = (f"[{n_zurueck} pruefen]" if ist_klein(sh)
                     else MARKE_ZURUECK.format(n=n_zurueck))
            absatz_schreiben(sh.text_frame, marke, FARBE_LUECKE,
                             groesse=10, erster=True)
            merken(nr, folienart, feld_id, "-",
                   "Vorschlaege zur Pruefung zurueckgehalten", alt,
                   marke, "")
            continue

        zeile_kat = kat.get((folienart, feld_id), {})
        feldtyp = zeile_kat.get("typ", "")
        klein = ist_klein(sh)

        # Ein Ja/Nein-Feld hat EINE Antwort. Vier Belege fuer dasselbe "Ja"
        # ergaben vorher vier Folien mit je zwei Woertern - Unsinn. Dasselbe
        # gilt fuer kleine Kaestchen (Nr./Index, Anzahl VE): dort steht ein
        # Wert, keine Liste. Die uebrigen Belege stehen im Protokoll und in
        # zuordnung.xlsx, sie gehen nicht verloren.
        einwertig = feldtyp == "Ja/Nein" or klein
        if einwertig:
            offen = liste[:1]
            for z in liste[1:]:
                merken(nr, folienart, feld_id, z["kennung"],
                       "Beleg nicht geschrieben (Feld ist einwertig)", "",
                       z["text"], f"{z.get('datei','')} S.{z.get('seite','')}")
        else:
            offen = liste[:a.max_je_feld]
        uebrig_gesamt = (len(liste) - len(offen)) if not einwertig else 0
        teil_nr = 0

        # Solange noch Eintraege offen sind, Folien fuellen. Passt nichts mehr
        # auf die Folie, wird eine weitere derselben Folienart angelegt.
        while offen:
            tf = sh.text_frame
            alt = text_zusammengefuehrt(tf)
            tf.clear()
            rahmen_vorbereiten(tf)
            je_zeile, platz = platz_budget(sh)
            verbraucht, erster = 0.0, True
            geschrieben_hier = 0

            while offen:
                z = offen[0]
                if feldtyp == "Ja/Nein":
                    inhalt = "Ja"
                elif klein:
                    inhalt = kuerzen(anzeige(z["text"]), 40)
                else:
                    inhalt = anzeige(z["text"])

                # Ein einzelner Eintrag, der allein schon zu lang ist, wird
                # gekuerzt statt uebersprungen. Vollstaendig steht er im
                # Protokoll und in zuordnung.xlsx.
                if kosten(inhalt, je_zeile) > platz and erster:
                    inhalt = kuerzen(inhalt, int(je_zeile * max(1, platz - 2)))

                if verbraucht + kosten(inhalt, je_zeile) > platz and not erster:
                    break                      # voll - Rest auf Folgefolie

                absatz_schreiben(tf, inhalt, FARBE_NEU,
                                 groesse=None if klein else PT_TEXT,
                                 erster=erster)
                quelle = f"{z['kennung']} \u00b7 {z.get('datei','')}"
                if z.get("seite") not in (None, "", "None"):
                    quelle += f", S. {z['seite']}"
                if z.get("pruefen"):
                    quelle += "  \u00b7  PRUEFEN: " + norm(z["pruefen"])[:70]
                elif z.get("freigabe") == "ja":
                    quelle += "  \u00b7  freigegeben (Mensch)"
                if klein:
                    quelle = z["kennung"]
                absatz_schreiben(tf, quelle, FARBE_QUELLE, groesse=PT_QUELLE)

                merken(nr, folienart, feld_id, z["kennung"],
                       "Text geschrieben (freigegeben)"
                       if z.get("freigabe") == "ja" else "Text geschrieben",
                       alt if erster else "", z["text"],
                       f"{z.get('datei','')} S.{z.get('seite','')}")
                verbraucht += kosten(inhalt, je_zeile)
                geschrieben += 1
                geschrieben_hier += 1
                erster = False
                offen.pop(0)

            if not offen:
                break
            if einwertig:
                break                          # nie fortsetzen
            if geschrieben_hier == 0:
                break                          # Form zu klein, Endlosschleife
            if teil_nr >= 4:                   # Notbremse
                absatz_schreiben(tf, f"\u2026 {len(offen)} weitere Eintraege, "
                                     f"siehe zuordnung.xlsx",
                                 FARBE_LUECKE, groesse=PT_QUELLE)
                offen = []
                break

            # Folgefolie anlegen
            L = layout_nach_name(prs, folienart)
            if L is None:
                break
            teil_nr += 1
            ziel_folie = prs.slides.add_slide(L)
            nr = len(prs.slides)
            fortsetzungen.add(nr)
            neue_folien.append((nr, folienart + f" (Fortsetzung {teil_nr})"))
            merken(nr, folienart, "-", "-", "Fortsetzungsfolie angelegt", "",
                   f"{len(offen)} Eintraege passten nicht mehr", "")
            # Titel der Folgefolie kennzeichnen, falls vorhanden
            titel = platzhalter_finden(ziel_folie, "ph0")
            if titel is not None and titel.has_text_frame:
                absatz_schreiben(titel.text_frame,
                                 f"{zeile_kat.get('feldname','')[:40]} "
                                 f"(Fortsetzung {teil_nr})",
                                 FARBE_QUELLE, groesse=14, erster=True)
            sh = platzhalter_finden(ziel_folie, feld_id)
            if sh is None or not sh.has_text_frame:
                break

        if uebrig_gesamt:
            merken(nr, folienart, feld_id, "-", "nicht geschrieben", "",
                   f"{uebrig_gesamt} Eintraege ueber --max-je-feld hinaus", "")

        # Einige geschrieben, andere zurueckgehalten: am Ende eine Zeile, damit
        # niemand das Feld fuer vollstaendig haelt.
        if n_zurueck and not einwertig and sh is not None \
                and sh.has_text_frame:
            hinweis = MARKE_WEITERE.format(n=n_zurueck)
            absatz_schreiben(sh.text_frame, hinweis, FARBE_LUECKE,
                             groesse=PT_QUELLE)
            merken(nr, folienart, feld_id, "-",
                   "Vorschlaege zur Pruefung zurueckgehalten", "", hinweis, "")

    # ===================================================== Faehigkeitstabelle
    #
    # Drei Dinge sind hier anders als frueher:
    #
    # 1. Beispielzeilen der Vorlage werden geleert (und protokolliert).
    #    Gluths Vorlage enthaelt Beispielwerte aus einem anderen Projekt
    #    ("Messung Lagersitztiefe ... 30 µm"). Neben unseren Eintraegen sahen
    #    sie aus wie Ergebnisse - das ist gefaehrlicher als jede Optik.
    #    Mit --beispiele-behalten bleiben sie stehen.
    #
    # 2. Ein Satz = eine Zeile. Passt ein Satz zu Verfahren UND Kennwert,
    #    stehen beide in derselben Zeile statt in zwei Zeilen. Das ist keine
    #    inhaltliche Gruppierung (die bleibt Handarbeit), sondern nur: gleiche
    #    Quelle, gleiche Zeile.
    #
    # 3. In den Zellen steht ein WOERTLICHER Auszug der Fundstelle
    #    (z. B. "Cgk ≥ 1.33") statt maschineller Werte wie "GE 60.0 mm".
    #    Und die Tabelle bekommt ein Hoehenbudget: was nicht mehr auf die
    #    Folie passt, kommt auf eine Fortsetzungsfolie mit gleichem Tabellen-
    #    kopf, statt unten aus der Folie zu laufen.
    fk = [z for z in zuord if z["feld_id"].startswith("tab:Faehigkeit")]
    spalte = {"tab:Faehigkeit:Prozess": 0, "tab:Faehigkeit:Verfahren": 1,
              "tab:Faehigkeit:Kennwert": 2, "tab:Faehigkeit:Toleranz": 3}
    n_tab_zurueck = sum(n for k_, n in zurueck.items()
                        if k_[1].startswith("tab:Faehigkeit"))
    if fk or n_tab_zurueck:
        tab, tab_shape, folie_fk, nr_fk = None, None, None, None
        for nr, s in enumerate(prs.slides, start=1):
            for sh_ in s.shapes:
                if sh_.has_table and len(sh_.table.rows) and "cmk" in " ".join(
                        norm(c.text).lower() for c in sh_.table.rows[0].cells):
                    tab, tab_shape, folie_fk, nr_fk = sh_.table, sh_, s, nr
                    break
            if tab is not None:
                break
        if tab is None:
            uebersprungen.append(("Fähigkeitstabelle", "-",
                                  "Tabelle in der Datei nicht gefunden"))
        elif len(tab.columns) < 6:
            uebersprungen.append(("Fähigkeitstabelle", "-",
                                  "Tabelle hat weniger als 6 Spalten"))
        else:
            # --- 1. Beispielzeilen leeren
            if not a.beispiele_behalten:
                for ri in range(1, len(tab.rows)):
                    alt = " | ".join(norm(c.text) for c in tab.rows[ri].cells)
                    if alt.replace("|", "").strip():
                        for c in tab.rows[ri].cells:
                            c.text_frame.clear()
                        merken(nr_fk, "Fähigkeitstabelle", "-", "-",
                               "Beispielzeile der Vorlage geleert", alt, "", "")

            # --- 2. ein Satz = eine Zeile
            zeilen = collections.OrderedDict()
            for z in fk:
                ci = spalte.get(z["feld_id"])
                if ci is None:
                    continue
                k = re.sub(r"\W+", " ", norm(z["text"])).strip().lower()
                if k not in zeilen:
                    zeilen[k] = {"zellen": {}, "eintraege": []}
                inhalt = z.get("auszug") or ""
                if not inhalt or ci == 0:
                    inhalt = kuerzen(anzeige(z["text"]), 70)
                else:
                    inhalt = kuerzen(anzeige(inhalt), 70)
                if ci not in zeilen[k]["zellen"]:
                    zeilen[k]["zellen"][ci] = inhalt
                zeilen[k]["eintraege"].append(z)

            # --- 3. Hoehenbudget
            breiten = [c.width / 360000.0 for c in tab.columns]
            verfuegbar = ((prs.slide_height - tab_shape.top) / 360000.0
                          - FUSS_RESERVE_CM) * SICHERHEIT
            belegt = sum(r.height / 360000.0 for r in tab.rows)
            # leere Zeilen der Vorlage zaehlen nicht, sie werden gefuellt
            leer_vorhanden = sum(1 for ri in range(1, len(tab.rows))
                                 if not "".join(c.text for c in
                                                tab.rows[ri].cells).strip())
            belegt -= leer_vorhanden * MIN_ZEILE_CM
            teil = 0

            for k, zeile in zeilen.items():
                eintraege = zeile["eintraege"]
                kennungen = ", ".join(dict.fromkeys(e["kennung"]
                                                    for e in eintraege))
                seiten = ", ".join(dict.fromkeys(str(e.get("seite", ""))
                                                 for e in eintraege))
                hinweis = f"S. {seiten}"
                if any(e.get("pruefen") for e in eintraege):
                    hinweis += " · PRUEFEN"
                texte = dict(zeile["zellen"])
                texte[4] = hinweis
                texte[5] = kennungen
                h = zeilenhoehe_cm(texte, breiten)

                if belegt + h > verfuegbar:
                    # Fortsetzungsfolie mit gleichem Tabellenkopf
                    if teil >= 100:
                        uebersprungen.append(("Fähigkeitstabelle", "-",
                                              "Mehr als 100 Fortsetzungsfolien noetig"))
                        break
                    teil += 1
                    neu = tabelle_fortsetzen(prs, folie_fk, tab_shape, teil)
                    if neu is None:
                        uebersprungen.append(("Fähigkeitstabelle", "-",
                                              "Fortsetzung nicht moeglich"))
                        break
                    folie_neu, tab = neu
                    nr_fk = len(prs.slides)
                    fortsetzungen.add(nr_fk)
                    neue_folien.append((nr_fk, f"Fähigkeitstabelle "
                                               f"(Fortsetzung {teil})"))
                    merken(nr_fk, "Fähigkeitstabelle", "-", "-",
                           "Fortsetzungsfolie angelegt", "",
                           "Tabelle passte nicht mehr auf die Folie", "")
                    belegt = sum(r.height / 360000.0 for r in tab.rows) \
                        - MIN_ZEILE_CM          # eine leere Zeile ist da
                    if belegt + h > verfuegbar:
                        uebersprungen.append(("Fähigkeitstabelle", "-",
                                              f"Tabellenzeile {kennungen} passt nicht auf eine Folie"))
                        continue

                ri = erste_leere_zeile(tab)
                if ri is None:
                    ri = tabellenzeile_anhaengen(tab)
                    merken(nr_fk, "Fähigkeitstabelle", "-", "-",
                           "Tabellenzeile angelegt", "", f"Zeile {ri}", "")
                tab.rows[ri].height = Cm(h)
                for ci, inhalt in zeile["zellen"].items():
                    zelle_schreiben(tab, ri, ci, inhalt)
                zelle_schreiben(tab, ri, 4, hinweis, FARBE_QUELLE)
                zelle_schreiben(tab, ri, 5, kennungen, FARBE_QUELLE)
                belegt += h
                for e in eintraege:
                    ci = spalte.get(e["feld_id"])
                    merken(nr_fk, "Fähigkeitstabelle", e["feld_id"],
                           e["kennung"],
                           "Tabellenzelle geschrieben (freigegeben)"
                           if e.get("freigabe") == "ja"
                           else "Tabellenzelle geschrieben", "",
                           zeile["zellen"].get(ci, ""),
                           f"{e.get('datei','')} S.{e.get('seite','')}")
                    geschrieben += 1

            # zurueckgehaltene Tabellen-Vorschlaege als letzte Zeile nennen
            if n_tab_zurueck:
                h = MIN_ZEILE_CM
                if belegt + h > verfuegbar and teil < 100:
                    teil += 1
                    neu = tabelle_fortsetzen(prs, folie_fk, tab_shape, teil)
                    if neu is not None:
                        _, tab = neu
                        nr_fk = len(prs.slides)
                        fortsetzungen.add(nr_fk)
                        neue_folien.append((nr_fk, f"Fähigkeitstabelle (Fortsetzung {teil})"))
                        belegt = sum(r.height / 360000.0 for r in tab.rows) - MIN_ZEILE_CM
                if belegt + h <= verfuegbar:
                    ri = erste_leere_zeile(tab)
                    if ri is None:
                        ri = tabellenzeile_anhaengen(tab)
                    tab.rows[ri].height = Cm(h)
                    hinweis = MARKE_WEITERE.format(n=n_tab_zurueck)
                    zelle_schreiben(tab, ri, 0, hinweis, FARBE_LUECKE)
                    merken(nr_fk, "Fähigkeitstabelle", "-", "-",
                           "Vorschlaege zur Pruefung zurueckgehalten", "",
                           hinweis, "")
                else:
                    uebersprungen.append(("Fähigkeitstabelle", "-",
                                          "Hinweis auf zurueckgehaltene Vorschlaege passte nicht"))

    # ==================================================== Versionsuebersicht
    vers_tab, nr_v = None, None
    for nr, s in enumerate(prs.slides, start=1):
        t = tabelle_finden(s, "verantwortlich")
        if t is not None:
            vers_tab, nr_v = t, nr
            break
    if vers_tab is not None:
        ri = erste_leere_zeile(vers_tab)
        if ri is None:
            ri = tabellenzeile_anhaengen(vers_tab)
        txt = (f"{geschrieben} Eintraege automatisch vorbefuellt "
               f"(Kennungen R-0001 ff.)")
        if a.station:
            txt += f"; Station: {a.station} (Angabe Mensch)"
        if zurueck:
            txt += (f"; {sum(zurueck.values())} Vorschlaege zur Pruefung "
                    f"zurueckgehalten (siehe zuordnung.xlsx)")
        if freigabe_info["datei"]:
            txt += (f"; Freigaben Mensch: {freigabe_info['ja']} ja, "
                    f"{freigabe_info['nein']} nein")
        zelle_schreiben(vers_tab, ri, 0, f"{datetime.now():%d.%m.%Y}")
        zelle_schreiben(vers_tab, ri, 1, "auto-1")
        zelle_schreiben(vers_tab, ri, 2, txt)
        zelle_schreiben(vers_tab, ri, 3, a.verantwortlich)
        merken(nr_v, "Versionsübersicht", "tab:Version", "-",
               "Versionszeile angelegt", "", txt, "")

    # ======================================================= Luecken markieren
    luecken = []
    for nr, s_ in enumerate(prs.slides, start=1):
        if nr in fortsetzungen:
            continue             # auf Fortsetzungsfolien keine Luecken melden
        art = s_.slide_layout.name
        for sh in s_.shapes:
            if not sh.is_placeholder or not sh.has_text_frame:
                continue
            fid = f"ph{sh.placeholder_format.idx}"
            zeile = kat.get((art, fid))
            if not zeile or zeile["befuellbar"] not in ("ja", "teilweise"):
                continue
            # Bildplatzhalter nicht beschriften. Ihr "teilweise" bedeutet
            # "Bild technisch extrahierbar, Auswahl nicht automatisierbar" -
            # das ist keine Luecke, die wir nicht belegen konnten, und ein
            # Texthinweis in einem Bildrahmen sieht nur kaputt aus.
            if zeile["quelle"] == "Bild" or zeile["typ"] == "Bild":
                continue
            if norm(text_zusammengefuehrt(sh.text_frame)):
                continue
            name = zeile["feldname"][:48]
            marke = "[offen]" if ist_klein(sh) else MARKE_LUECKE.format(
                kuerzen(name, 38))
            sh.text_frame.clear()
            absatz_schreiben(sh.text_frame, marke,
                             FARBE_LUECKE, groesse=10, erster=True)
            luecken.append((nr, art, fid, name))
            merken(nr, art, fid, "-", "Luecke markiert", "",
                   MARKE_LUECKE.format(name), "")

    prs.save(str(ziel))

    # ---------------------------------------------------------- Protokoll
    pp = ziel.with_name(ziel.stem + "_schreibprotokoll.csv")
    with pp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, delimiter=";", fieldnames=[
            "zeit", "folie", "folienart", "feld_id", "kennung", "aktion",
            "alt", "neu", "quelle"])
        w.writeheader()
        w.writerows(protokoll)

    # ------------------------------------------------------------- Zahlen
    je_aktion = collections.Counter(p["aktion"] for p in protokoll)
    L = [f"SCHREIBEN  {datetime.now():%Y-%m-%d %H:%M}", "=" * 66, ""]
    L += [f"  Vorlage                     {vorlage.name}",
          f"  Zuordnungen eingelesen      {len(zuord)}",
          f"  Eintraege geschrieben       {geschrieben}",
          f"  Folien neu angelegt         {len(neue_folien)}",
          f"  Luecken markiert            {len(luecken)}",
          f"  Protokollzeilen             {len(protokoll)}"]
    if a.station:
        L.append(f"  Station (Angabe Mensch)     {a.station}")
    else:
        L.append("  Station                     nicht zugeordnet")
    if a.nur_sichere:
        L.append("  nur Zuordnungen ohne Pruefvermerk")
        L.append(f"  zurueckgehalten (pruefen)   {sum(zurueck.values())}"
                 f"  -> im Feld als 'Vorschlaege zur Pruefung' vermerkt")
    if freigabe_info["datei"]:
        L.append(f"  Freigaben aus               {freigabe_info['datei']}")
        L.append(f"    ja (geschrieben)          {freigabe_info['ja']}")
        L.append(f"    nein (weggelassen)        {freigabe_info['nein']}")
        if freigabe_info["nicht_gefunden"]:
            L.append(f"    passten zu keinem Eintrag {freigabe_info['nicht_gefunden']}"
                     f"  (anderer Lauf/anderes Projekt?)")
    L += ["", "  Was getan wurde:"]
    for k, n in je_aktion.most_common():
        L.append(f"    {n:>5}  {k}")
    if neue_folien:
        L += ["", "  Neu angelegte Folien:"]
        for nr, art in neue_folien:
            L.append(f"    {nr:>3}  {art}")
    if uebersprungen:
        L += ["", "  Uebersprungen:"]
        for fa, fid, warum in uebersprungen:
            L.append(f"    {fa} / {fid}: {warum}")
    L += ["", "  Lesart:",
          "    Das Original wurde nicht veraendert. Geschrieben wurde in eine",
          "    Kopie.",
          "    Jeder blaue Eintrag ist automatisch entstanden und traegt",
          "    darunter in grau seine Kennung, die Quelldatei und die Seite.",
          "    Braune Eintraege sind offene Stellen, die wir nicht belegen",
          "    konnten.",
          "    Die Faehigkeitstabelle ist NICHT inhaltlich gruppiert: jeder",
          "    Quellsatz ist eine eigene Zeile, weil nicht bekannt ist, welche",
          "    Toleranz zu welchem Prozess gehoert. In den Zellen steht ein",
          "    woertlicher Auszug; der ganze Satz steht in zuordnung.xlsx.",
          "    Beispielzeilen der Vorlage wurden geleert (siehe Protokoll),",
          "    ausser mit --beispiele-behalten.",
          "    Aufzaehlungszeichen aus Symbolschriften (sonst leere Kaestchen)",
          "    wurden als normaler Punkt dargestellt - sonst nichts veraendert.",
          "", "Diese Datei enthaelt keine Saetze."]

    (OUT / "schreiben_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
    # Dieselben Zahlen maschinenlesbar fuer die Weboberflaeche. Nur Zahlen
    # und Dateinamen, keine Saetze.
    (OUT / "schreiben_zahlen.json").write_text(json.dumps({
        "ausgabe": ziel.name,
        "geschrieben": geschrieben,
        "zurueckgehalten": sum(zurueck.values()),
        "nur_sichere": bool(a.nur_sichere),
        "freigabe_datei": freigabe_info["datei"],
        "freigabe_ja": freigabe_info["ja"],
        "freigabe_nein": freigabe_info["nein"],
        "freigabe_ohne_treffer": freigabe_info["nicht_gefunden"],
        "luecken": len(luecken),
        "folien_neu": len(neue_folien),
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    for z in L:
        print(ascii_safe(z))

    print()
    print("geschrieben:")
    print(f"  {ziel}                 VERTRAULICH")
    print(f"  {pp}        VERTRAULICH")
    print(f"  {OUT / 'schreiben_zahlen.txt'}   nur Zahlen")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
