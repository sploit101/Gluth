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
    from pptx.util import Pt
except ImportError:                                            # pragma: no cover
    raise SystemExit("python-pptx fehlt:  python -m pip install python-pptx")

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

FARBE_NEU = RGBColor(0x1A, 0x5C, 0x8A)      # dunkelblau: automatisch gefuellt
FARBE_QUELLE = RGBColor(0x77, 0x77, 0x77)   # grau: Kennung und Herkunft
FARBE_LUECKE = RGBColor(0x9A, 0x64, 0x10)   # braun: offene Stelle

MARKE_LUECKE = "[offen: {}]"


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


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
CM_PRO_ZEICHEN = 0.205          # mittlere Zeichenbreite bei 11 pt
CM_PRO_ZEILE = 0.52             # Zeilenhoehe bei 11 pt
ZEILE_QUELLE = 0.75             # eine 8-pt-Zeile kostet so viel wie 0,75 Zeilen
SICHERHEIT = 0.88               # Rand gegen Schaetzfehler


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
    if a.nur_sichere:
        zuord = [z for z in zuord if not z.get("pruefen")]
    if not zuord:
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

    nach_feld: dict[tuple, list] = collections.defaultdict(list)
    for z in zuord:
        nach_feld[(z["folienart"], z["feld_id"])].append(z)

    arten = folien_nach_art(prs)
    neue_folien = []
    geschrieben, uebersprungen = 0, []

    # ============================================================ Textfelder
    fortsetzungen = set()        # Folien, die nur Fortsetzung sind

    for (folienart, feld_id), liste in sorted(nach_feld.items()):
        if not feld_id.startswith("ph"):
            continue

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
            neue_folien.append((nr, folienart))
            merken(nr, folienart, "-", "-", "Folie angelegt", "", folienart, "")

        sh = platzhalter_finden(ziel_folie, feld_id)
        if sh is None or not sh.has_text_frame:
            uebersprungen.append((folienart, feld_id,
                                  "Platzhalter fehlt oder kein Textfeld"))
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
                    inhalt = kuerzen(z["text"], 40)
                else:
                    inhalt = norm(z["text"])

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
                if klein:
                    quelle = z["kennung"]
                absatz_schreiben(tf, quelle, FARBE_QUELLE, groesse=PT_QUELLE)

                merken(nr, folienart, feld_id, z["kennung"], "Text geschrieben",
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

    # ===================================================== Faehigkeitstabelle
    fk = [z for z in zuord if z["feld_id"].startswith("tab:Faehigkeit")]
    spalte = {"tab:Faehigkeit:Prozess": 0, "tab:Faehigkeit:Verfahren": 1,
              "tab:Faehigkeit:Kennwert": 2, "tab:Faehigkeit:Toleranz": 3,
              "tab:Faehigkeit:Kommentar": 4, "tab:Faehigkeit:Kennung": 5}
    if fk:
        tab = None
        nr_fk = None
        for nr, s in enumerate(prs.slides, start=1):
            t = tabelle_finden(s, "cmk")
            if t is not None:
                tab, nr_fk = t, nr
                break
        if tab is None:
            uebersprungen.append(("Fähigkeitstabelle", "-",
                                  "Tabelle in der Datei nicht gefunden"))
        else:
            for z in fk:
                ci = spalte.get(z["feld_id"])
                if ci is None:
                    continue
                ri = erste_leere_zeile(tab)
                if ri is None:
                    ri = tabellenzeile_anhaengen(tab)
                    merken(nr_fk, "Fähigkeitstabelle", "-", "-",
                           "Tabellenzeile angelegt", "", f"Zeile {ri}", "")
                wert = norm(z["text"])
                if z.get("value") is not None:
                    wert = " ".join(x for x in [z.get("operator") or "",
                                                str(z["value"]),
                                                z.get("unit") or ""] if x)
                # Lange Saetze sprengen die Tabelle. Vollstaendig stehen sie
                # in der Herkunftsliste und im Protokoll.
                zelle_schreiben(tab, ri, ci, kuerzen(wert, 90))
                zelle_schreiben(tab, ri, 5, z["kennung"], FARBE_QUELLE)
                hinweis = f"S. {z.get('seite','')}"
                if z.get("pruefen"):
                    hinweis += " · PRUEFEN"
                zelle_schreiben(tab, ri, 4, hinweis, FARBE_QUELLE)
                merken(nr_fk, "Fähigkeitstabelle", z["feld_id"], z["kennung"],
                       "Tabellenzelle geschrieben", "", wert,
                       f"{z.get('datei','')} S.{z.get('seite','')}")
                geschrieben += 1

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
          "    Die Faehigkeitstabelle ist NICHT gruppiert: jede Fundstelle ist",
          "    eine eigene Zeile, weil nicht bekannt ist, welche Toleranz zu",
          "    welchem Prozess gehoert. Das Zusammenfuehren bleibt Handarbeit.",
          "", "Diese Datei enthaelt keine Saetze."]

    (OUT / "schreiben_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
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
