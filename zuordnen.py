#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zuordnen.py — entscheidet, welche extrahierte Anforderung in welches Feld
der Stationsakte gehoert.

Das ist der Kern der Erstbefuellung. Davor steht die Extraktion, danach
der Schreiber. Hier wird nichts in die PowerPoint geschrieben - diese
Datei trifft nur Entscheidungen und legt sie offen.

Grundsaetze
-----------
1. Kein Sprachmodell. Nur Regeln, und jede Zuordnung nennt die Regel,
   die sie erzeugt hat. Eine Zuordnung ohne nachvollziehbaren Grund ist
   wertlos, weil niemand sie pruefen kann.

2. Es wird nur fuer Felder zugeordnet, die in zielfelder.csv als
   befuellbar ("ja") gelten. Die Einschaetzung dort stammt aus der
   Pruefung gegen LH_04K_000051_V1.02 - 8 Fundstellen je Inhaltsart
   gezaehlt. Felder mit "teilweise" bekommen Zuordnungen, aber jede wird
   als "pruefen" markiert. Felder mit "nein" bekommen keine.

3. Der grosse Teil der Anforderungen landet NIRGENDS. Das ist richtig
   und wird gezaehlt, nicht versteckt. Ein Lastenheft enthaelt Abnahme-,
   Dokumentations- und Sicherheitsanforderungen, fuer die es in der
   Stationsakte kein Feld gibt.

4. Die STATIONSZUORDNUNG wird nicht geraten. Das Lastenheft ist nach
   Themen gegliedert, die Stationsakte nach Stationen - welcher Satz zu
   welcher Station gehoert, steht nirgends im Dokument. Solange keine
   Stationsliste vorliegt, steht in jeder Zeile "nicht zugeordnet".
   Mit --station kann eine Station gesetzt werden; dann ist es eine
   Angabe des Menschen und wird auch so protokolliert.

Aufruf
------
    python zuordnen.py --datei LH_04K_000051_V1.02
    python zuordnen.py --datei LH_04K_000051_V1.02 --station "Station 10 Abisolieren"
    python zuordnen.py --nur-sichere          # nur Felder mit befuellbar=ja

Ausgabe
-------
    out/zuordnung.jsonl        maschinenlesbar, Eingabe fuer den Schreiber
    out/zuordnung_zahlen.txt   nur Zahlen        -> darf geteilt werden
    out/zuordnung.xlsx         mit Wortlauten    -> VERTRAULICH
"""

from __future__ import annotations

import argparse
import collections
import csv
import json
import re
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

# ---------------------------------------------------------------------------
# MUELL. Das Lastenheft enthaelt auf jeder Seite Kopf- und Fusszeilen, ein
# Inhaltsverzeichnis und eine Versionstabelle. Die hat die Extraktion als
# Anforderungen mitgenommen. Wuerden wir sie zuordnen, stuenden in der Akte
# Seitenfusszeilen. Diese Saetze werden vorher verworfen und gezaehlt.
# ---------------------------------------------------------------------------
# --- harte Faelle: Kopf- und Fusszeilen, Inhaltsverzeichnis, Versionstabelle
MUELL_HART = re.compile(
    r"(printouts?\s+are\s+effective|not\s+subject\s+to\s+change\s+service|"
    r"all\s+rights\s+reserved|alle\s+rechte\s+vorbehalten|"
    r"distribution\s+within|weitergabe\s+sowie\s+vervielf|"
    r"template:\s*v\d|\brs\s*no\.?:|dok\.?-?id|"
    r"version\s+information\s+version\s+date|"
    r"\.{6,}|"                               # Inhaltsverzeichnis
    r"confirmation\s*\.{3,}|"
    r"^\s*[\d\s.,;:()-]+$)", re.I)         # nur Zahlen und Zeichen

# --- Signale fuer zerfallene Tabellen und aneinandergeklebte Ueberschriften
#
# Hintergrund: Die Textgewinnung aus PDF macht aus einer Tabellenzeile und aus
# einem Inhaltsverzeichnis eine Zeichenkette, die oberflaechlich wie ein Satz
# aussieht. Beispiele aus dem echten Lastenheft:
#   "80 - 500 ESD component Yes No Technical cleanliness requirements Yes ..."
#   "Key contacts DRAEXLMAIER Group 6 General project description A ..."
# Beide wuerden sonst in der Akte landen und waeren das Erste, was ein Leser
# sieht.
#
# Statt immer neuer Einzelregeln werden messbare Signale gezaehlt. Feuern
# mindestens zwei, gilt der Satz als nicht verwertbar. Welche Signale gefeuert
# haben, steht in out/zuordnung_verworfen.xlsx - damit ist jede Entscheidung
# nachpruefbar und nichts verschwindet stillschweigend.

SATZZEICHEN = re.compile(r'[.,;:!?()\[\]\u201e\u201c\u00ab\u00bb"]')
JA_NEIN = re.compile(r"\b(yes|no|ja|nein)\b", re.I)
VERB = re.compile(
    r"\b(must|shall|should|may|can|is|are|was|were|has|have|had|will|be|been|"
    r"muss|müssen|soll|sollen|darf|dürfen|kann|können|ist|sind|war|waren|"
    r"wird|werden|hat|haben|zu\s+\w+en)\b", re.I)
KAPITELNUMMER = re.compile(
    r"[A-ZÄÖÜ][a-zäöüß]{2,}\s+\d{1,2}(\.\d{1,2})?\s+[A-ZÄÖÜ]")
GLIEDERUNG = re.compile(r"\b\d{1,2}\.\d{1,2}(\.\d{1,2})?\s")
ZAHL_TOKEN = re.compile(r"^[\d.,/-]+$")
ZAHLBEREICH = re.compile(r"\d+\s*[-\u2013]\s*\d+")
GROSS_TOKEN = re.compile(r"^[A-ZÄÖÜ]")

MIN_FUER_SIGNALE = 120          # kurze Saetze werden nicht so geprueft


def muell_signale(text: str) -> list[str]:
    """Welche Signale sprechen gegen diesen Satz? Leere Liste = verwertbar."""
    t = norm(text)
    if MUELL_HART.search(t):
        return ["Kopfzeile oder Verzeichnis"]
    if len(t) < MIN_FUER_SIGNALE:
        return []

    sig = []

    # A  Satzzeichen-Armut. Fliesstext hat Zeichensetzung, Tabellenzeilen nicht.
    zeichen = len(SATZZEICHEN.findall(t))
    if zeichen == 0 or len(t) / zeichen > 90:
        sig.append("kaum Satzzeichen")

    # B  Mehrere Ja/Nein in einem Satz = Ankreuzspalten einer Tabelle
    if len(JA_NEIN.findall(t)) >= 2:
        sig.append("mehrere Ja/Nein")

    # C  Hoher Zahlenanteil = Messwert- oder Kalkulationstabelle
    tok = t.split()
    if tok:
        anteil = sum(1 for w in tok if ZAHL_TOKEN.match(w)) / len(tok)
        if anteil > 0.15:
            sig.append(f"Zahlenanteil {anteil:.0%}")

    # D  Kein Verb = kein Satz
    if not VERB.search(t):
        sig.append("kein Verb")

    # E  Kapitelnummer mitten im Text: "... Group 6 General ..."
    if KAPITELNUMMER.search(t):
        sig.append("Kapitelnummer im Text")

    # F  Mehrere Gliederungsnummern = Inhaltsverzeichnis
    if len(GLIEDERUNG.findall(t)) >= 2:
        sig.append("mehrere Gliederungsnummern")

    # G  Tabellenrest am ANFANG. Der Anfang entscheidet den ersten Eindruck:
    #    "80 - 500 ESD component Yes No Technical cleanliness requirements ..."
    #    Kein Satzzeichen in den ersten 70 Zeichen UND entweder ein
    #    Zahlenbereich, ein Ja/Nein oder mehrere Grossschreibungen.
    kopf = t[:70]
    if not SATZZEICHEN.search(kopf):
        kopf_tok = kopf.split()[1:]
        gross = sum(1 for w in kopf_tok if GROSS_TOKEN.match(w))
        if (ZAHLBEREICH.search(kopf) or JA_NEIN.search(kopf) or gross >= 3):
            sig.append("Tabellenrest am Anfang")

    # H  Aneinandergeklebte Aufzaehlung: viele grossgeschriebene Woerter
    #    mitten im Text, weil aus Stichpunkten eine Zeile wurde.
    if tok and len(tok) > 8:
        anteil_gross = sum(1 for w in tok[1:] if GROSS_TOKEN.match(w)) / len(tok)
        if anteil_gross > 0.33:
            sig.append(f"Aufzaehlung geklebt ({anteil_gross:.0%} gross)")

    return sig


# Mindestlaenge, damit ein Satz ueberhaupt Inhalt tragen kann
MIN_LAENGE = 40

# ---------------------------------------------------------------------------
# Die Regeln. Je Regel: Name, Zielfeld (Folienart, Feld-Kennung), Test,
# Sicherheit. Reihenfolge ist bedeutsam: spezifische Regeln vor allgemeinen.
#
# Ein Satz kann in die Faehigkeitstabelle UND in ein Textfeld gehen -
# "Toleranz der Abisolierlaenge +/- 0,2 mm" gehoert in die Tabelle (als
# kurzer Auszug) und in die Aufgabenbeschreibung (als ganzer Satz). Unter den
# TEXTfeldern bekommt ihn aber nur eines, siehe TEXTFELD_VORRANG.
# ---------------------------------------------------------------------------

VERFAHREN = re.compile(
    r"\b(msa\s*\d?|mfu|mgf|pfu|gauge\s*r\s*&?\s*r|g\s?rr|"
    r"measurement\s+system\s+analysis|machine\s+capability|"
    r"process\s+capability|maschinenf[aä]higkeit|prozessf[aä]higkeit|"
    r"messmittelf[aä]higkeit)\b", re.I)

# Kennwert braucht Buchstabenkuerzel UND eine Zahl - "Cpk" allein ist eine
# Erwaehnung, "Cpk >= 1,67" ist ein Wert.
KENNWERT = re.compile(
    r"\b(c\s?[gmp]k?|c\s?p)\b[^.;]{0,20}?"
    r"(>=|<=|≥|≤|=|von|of)?\s*\d+[.,]?\d*", re.I)

# Toleranz braucht ein Toleranzzeichen oder das Wort, plus Zahl plus Einheit
TOLERANZ = re.compile(
    r"((±|\+/-|\+\s*/\s*-)\s*\d+[.,]?\d*\s*(mm|µm|um|%|s|sec|nm|bar|°)|"
    r"\btoleran[cz]\w*\b[^.;]{0,40}?\d+[.,]?\d*\s*(mm|µm|um|%|s|sec|nm|bar|°)|"
    r"\b(repeat\s+accuracy|wiederholgenauigkeit)\b[^.;]{0,40}?\d)", re.I)

PROZESS_NAME = re.compile(
    r"\b(stripping|crimp\w*|cutting|insertion|positioning|"
    r"vision\s+(system|check\w*)|inspection|measurement|testing|"
    r"screw\w*|pressing|welding|solder\w*|labell?\w*|marking|feeding|"
    r"abisolier\w*|crimpen|schneiden|einlegen|positionier\w*|"
    r"sichtpr[uü]f\w*|messung|verschrauben|einpress\w*)\b", re.I)

KALIBRIEREN = re.compile(
    r"\b(calibrat\w*|re-?calibration|kalibrier\w*|"
    r"reference\s+(sample|master|part)\w*|referenzmuster|"
    r"zero\s+point|nullpunkt|teach\w*|einricht\w*)\b", re.I)

VARIANTE = re.compile(
    r"\b(per\s+variant|each\s+product\s+variant|product\s+variant|"
    r"je\s+variante|produktvariante|variante[nr]?\b|"
    r"different\s+type\s+of)\b", re.I)

STATIONSINHALT = re.compile(
    r"\b(assembly\s+station|test\s+station|output\s+station|"
    r"input\s+station|workstation|production\s+line|"
    r"montagestation|pr[uü]fstation|arbeitsplatz)\b", re.I)

# Saetze, die zwar lang und verbindlich sind, aber nicht die Aufgabe
# beschreiben: Abnahme, Dokumentation, Recht, Sicherheit allgemein.
KEINE_AUFGABE = re.compile(
    r"\b(ce\s+certificate|operating\s+instructions|maintenance\s+plan|"
    r"documentation\s+must\s+contain|applicable\s+documents|"
    r"legally\s+binding|warranty|guarantee|liability|haftung|"
    r"acceptance\s+(test|protocol)|abnahme|gew[aä]hrleistung|"
    r"machinery\s+directive|maschinenrichtlinie|"
    r"insurance|versicherung|training\s+courses)\b", re.I)


# Saetze ueber Dokumente, Termine und Lieferumfang. Sie enthalten oft die
# Woerter "inspection", "calibration" oder "MFU", beschreiben aber keinen
# Prozess und kein Kalibrierteil, sondern etwas, das geliefert oder
# nachgewiesen werden muss. Beispiel aus dem echten Lauf: ein Satz ueber ein
# Pruefzertifikat im Lieferumfang stand als "Prozess" in der Faehigkeitstabelle.
DOKUMENT = re.compile(
    r"\b(certificat\w*|zertifikat\w*|checklist\w*|schedule|terminplan|"
    r"included\s+in\s+the\s+delivery|lieferumfang|spare\s+parts|"
    r"ersatzteil\w*|acceptance\s+run|machine\s+acceptance|"
    r"documentation|dokumentation)\b", re.I)

PRUEFWORT = re.compile(
    r"\b(monitor\w*|check\w*|inspect\w*|test\w*|"
    r"[uü]berwach\w*|pr[uü]f\w*|mess\w*)\b", re.I)


def _stamm(wort: str) -> str:
    return wort.lower()[:4]


def prozess_und_pruefung(t: str) -> bool:
    """
    Prozessname UND ein Pruefwort - aber zwei VERSCHIEDENE Woerter.

    Vorher genuegte "Inspection certificate ...", weil "inspection" zugleich
    als Prozessname und als Pruefwort zaehlte. Jetzt muss das Pruefwort ein
    anderes Wort mit anderem Wortstamm sein ("crimps ... monitoring" ja,
    "testing ... test" nein).
    """
    # Paarweise: irgendein Prozesswort P und irgendein Pruefwort C, die sich
    # nicht ueberlappen und verschiedene Wortstaemme haben. So zaehlt
    # "inspection ... vision system" (C = inspection, P = vision system),
    # aber nicht "Inspection certificate" (nur ein einziges Wort).
    prozesse = list(PROZESS_NAME.finditer(t))
    pruefungen = list(PRUEFWORT.finditer(t))
    for p in prozesse:
        ps, pe = p.span()
        for c in pruefungen:
            cs, ce = c.span()
            if cs < pe and ps < ce:
                continue
            if _stamm(p.group(0)) != _stamm(c.group(0)):
                return True
    return False


def auszug(t: str, muster: re.Pattern, vorlauf: int = 0) -> str:
    """
    Woertlicher Ausschnitt fuer eine Tabellenzelle.

    Eine 3,9 cm breite Zelle vertraegt keinen ganzen Satz. Statt eines
    maschinell umformulierten Werts ("GE 60.0 mm") steht hier der Original-
    wortlaut der Fundstelle, bei Bedarf mit etwas Text davor, damit klar ist,
    worauf sich der Wert bezieht. Jeder Ausschnitt ist ein exaktes Teilstueck
    des Satzes - es wird nichts umformuliert. Mehrere Fundstellen werden mit
    " | " getrennt.
    """
    teile, gesehen = [], set()
    for m in muster.finditer(t):
        a, b = m.span()
        if vorlauf:
            start = max(0, a - vorlauf)
            if start > 0:
                leer = t.find(" ", start)
                start = leer + 1 if 0 <= leer < a else start
            stueck = t[start:b].strip()
            if start > 0:
                stueck = "…" + stueck
        else:
            stueck = t[a:b].strip()
        schluessel = stueck.lower()
        if stueck and schluessel not in gesehen:
            gesehen.add(schluessel)
            teile.append(stueck)
    return " | ".join(teile)


# Wenn ein Satz zu mehreren TEXTfeldern passt, bekommt ihn nur das
# spezifischste. Vorher stand derselbe Satz auf zwei Folien (z. B. unter
# "Aufgabenstellung" und unter "Stationsbeschreibung"). Die anderen passenden
# Felder werden in der Spalte "Auch passend" vermerkt - es geht nichts
# verloren, aber in der Akte steht jeder Satz nur einmal.
#
# Nicht betroffen: die Faehigkeitstabelle (dort steht nur ein kurzer Auszug)
# und das Varianten-Feld (dort steht nur "Ja", der Satz ist der Beleg).
TEXTFELD_VORRANG = {
    "kalibrieren": 1,
    "stationsinhalt": 2,
    "aufgabe_ausfuehrlich": 3,
}

# Welche Regel liefert welchen Auszug fuer die Tabellenzelle
AUSZUG_JE_REGEL = {
    "faehigkeit_verfahren": lambda t: auszug(t, VERFAHREN),
    "faehigkeit_kennwert": lambda t: auszug(t, KENNWERT),
    "faehigkeit_toleranz": lambda t: auszug(t, TOLERANZ, vorlauf=35),
}


def regeln(kat: dict) -> list[tuple]:
    """
    Liefert die Regelliste. kat ist der Zielfeld-Katalog, damit jede Regel
    weiss, ob ihr Feld als "ja" oder "teilweise" eingeschaetzt ist.
    """
    def hat(folienart, feld):
        return (folienart, feld) in kat

    R = []

    def add(name, folienart, feld, test, sicherheit):
        if hat(folienart, feld):
            R.append((name, folienart, feld, test, sicherheit))

    # --- Faehigkeitstabelle: der staerkste Befund der Pruefung
    add("faehigkeit_verfahren", "Fähigkeitstabelle", "tab:Faehigkeit:Verfahren",
        lambda t, r: bool(VERFAHREN.search(t)) and not DOKUMENT.search(t),
        "hoch")

    add("faehigkeit_kennwert", "Fähigkeitstabelle", "tab:Faehigkeit:Kennwert",
        lambda t, r: bool(KENNWERT.search(t)), "hoch")

    add("faehigkeit_toleranz", "Fähigkeitstabelle", "tab:Faehigkeit:Toleranz",
        lambda t, r: bool(TOLERANZ.search(t)), "mittel")

    # Prozessname nur dann, wenn der Satz auch nach Pruefen oder Messen
    # klingt - sonst landet jede Erwaehnung von "assembly" in der Tabelle.
    #
    # UND nicht, wenn es ein Faehigkeitsverfahren ist: "measurement system
    # analysis" und "measuring equipment capability" sind Verfahren, keine
    # Prozessbezeichnungen. Ohne diese Ausnahme landete jeder MSA-Satz
    # zusaetzlich in der Spalte "Prozess".
    add("faehigkeit_prozess", "Fähigkeitstabelle", "tab:Faehigkeit:Prozess",
        lambda t, r: (prozess_und_pruefung(t)
                      and not VERFAHREN.search(t)
                      and not DOKUMENT.search(t)
                      and not re.search(r"measur\w*\s+(equipment|system)", t, re.I)),
        "mittel")

    # --- Aufgabenbeschreibung: lange verbindliche Saetze, die wirklich die
    #     Aufgabe beschreiben
    add("aufgabe_ausfuehrlich", "Aufgabenstellung_ausführlich", "ph11",
        lambda t, r: (len(t) > 120
                      and r.get("binding") in ("MUSS", "VERBOT")
                      and not KEINE_AUFGABE.search(t)
                      and not DOKUMENT.search(t)
                      and bool(PROZESS_NAME.search(t))),
        "hoch")

    # --- Stationsbeschreibung
    add("stationsinhalt", "Stationsbeschriebung", "ph11",
        lambda t, r: (bool(STATIONSINHALT.search(t))
                      and r.get("binding") in ("MUSS", "SOLL", "VERBOT")
                      and len(t) > 60),
        "mittel")

    # --- Kalibrieren und Einrichten
    add("kalibrieren", "Kalibrier-/Einrichtteile", "ph11",
        lambda t, r: (bool(KALIBRIEREN.search(t)) and len(t) > 60
                      and not DOKUMENT.search(t)
                      and not KEINE_AUFGABE.search(t)),
        "mittel")

    # --- Varianten: hier wird kein Satz eingetragen, sondern ein Ja/Nein
    #     belegt. Der Satz ist der Nachweis.
    add("varianten_ja", "Bauteile", "ph22",
        lambda t, r: bool(VARIANTE.search(t)), "hoch")

    return R


SPALTE_FREIGABE = "Freigabe"


def freigaben_in_datei(pfad: Path) -> int:
    """Wie viele Zeilen der Datei haben eine Freigabe eingetragen?"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(pfad), read_only=True, data_only=True)
        ws = wb["Zuordnung"] if "Zuordnung" in wb.sheetnames else wb.active
        zeilen = ws.iter_rows(values_only=True)
        kopf = [str(x or "").strip() for x in next(zeilen, [])]
        if SPALTE_FREIGABE not in kopf:
            wb.close()
            return 0
        i = kopf.index(SPALTE_FREIGABE)
        n = sum(1 for z in zeilen
                if i < len(z) and str(z[i] or "").strip())
        wb.close()
        return n
    except Exception:
        return -1           # unbekannt -> der Aufrufer sichert vorsichtshalber


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def text_von(r: dict) -> str:
    for k in ("text", "sentence", "content", "chunk"):
        if r.get(k):
            return norm(r[k])
    return ""


def wert_von(r: dict, kandidaten: list[str]):
    for k in kandidaten:
        if r.get(k) not in (None, "", "None"):
            return r[k]
    return None


def katalog_laden(pfad: Path) -> dict:
    """(Folienart, Feld) -> Zeile, nur Felder mit befuellbar ja/teilweise."""
    if not pfad.exists():
        return {}
    kat = {}
    with pfad.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r.get("befuellbar") in ("ja", "teilweise"):
                kat[(r["folienart"], r["feld_id"])] = r
    return kat


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--req", default=str(OUT / "requirements.jsonl"))
    ap.add_argument("--katalog", default=str(HERE / "zielfelder.csv"))
    ap.add_argument("--datei", default=None,
                    help="nur Dateien, deren Name diesen Text enthaelt")
    ap.add_argument("--station", default=None,
                    help="Stationsbezeichnung. Wird als MENSCHLICHE Angabe "
                         "protokolliert, nicht aus dem Dokument abgeleitet.")
    ap.add_argument("--nur-sichere", action="store_true", dest="nur_sichere",
                    help="nur Felder mit befuellbar=ja")
    ap.add_argument("--ab-kennung", type=int, default=1,
                    help="erste Kennungsnummer (Standard 1 -> R-0001)")
    a = ap.parse_args()

    kpfad = Path(a.katalog)
    kat = katalog_laden(kpfad)
    if not kat:
        print("FEHLER: Zielfeld-Katalog nicht gefunden oder leer:",
              ascii_safe(kpfad))
        print("Erst 'python zielfelder_bauen.py' laufen lassen.")
        return 2
    if a.nur_sichere:
        kat = {k: v for k, v in kat.items() if v["befuellbar"] == "ja"}

    rpfad = Path(a.req)
    if not rpfad.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(rpfad))
        print("Erst 'python requirements.py' laufen lassen.")
        return 2

    rows, weg_datei = [], 0
    for line in rpfad.open(encoding="utf-8"):
        line = line.strip()
        if not line:
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            continue
        datei = str(wert_von(r, ["file", "filename", "source_file",
                                 "source", "path"]) or "")
        if a.datei and a.datei.lower() not in datei.lower():
            weg_datei += 1
            continue
        r["_datei"] = datei
        r["_seite"] = wert_von(r, ["page", "page_no", "seite", "page_number"])
        rows.append(r)

    if not rows:
        print("Keine Anforderungen uebrig. Filter zu eng?")
        return 1

    R = regeln(kat)

    # ------------------------------------------------------------- zuordnen
    treffer = []
    verworfen = []
    muell, zu_kurz, ohne_feld = 0, 0, 0
    mehrfach_text = 0           # Textfeld-Treffer, die dem Vorrang weichen
    ohne_feld_liste = []        # fuer die Pruefung, was die Regeln uebersehen
    gesehen = set()
    signal_zaehler = collections.Counter()

    for r in rows:
        t = text_von(r)
        sig = muell_signale(t)
        if len(sig) >= 2 or (sig and sig[0] == "Kopfzeile oder Verzeichnis"):
            muell += 1
            for x in sig:
                signal_zaehler[x] += 1
            verworfen.append({"text": t, "signale": "; ".join(sig),
                              "datei": r["_datei"], "seite": r["_seite"],
                              "binding": r.get("binding", "")})
            continue
        if len(t) < MIN_LAENGE:
            zu_kurz += 1
            continue

        wohin = []
        for name, folienart, feld, test, sicher in R:
            try:
                if test(t, r):
                    wohin.append((name, folienart, feld, sicher))
            except Exception:                    # pragma: no cover
                continue
        if not wohin:
            ohne_feld += 1
            ohne_feld_liste.append({"text": t, "datei": r["_datei"],
                                    "seite": r["_seite"],
                                    "binding": r.get("binding", "")})
            continue

        # Unter den Textfeldern nur das spezifischste behalten
        text_treffer = [w for w in wohin if w[0] in TEXTFELD_VORRANG]
        auch_passend = ""
        if len(text_treffer) > 1:
            text_treffer.sort(key=lambda w: TEXTFELD_VORRANG[w[0]])
            weg = text_treffer[1:]
            auch_passend = ", ".join(f"{w[1]} / {w[2]}" for w in weg)
            wohin = [w for w in wohin if w not in weg]
            mehrfach_text += len(weg)

        for name, folienart, feld, sicher in wohin:
            schluessel = (folienart, feld,
                          re.sub(r"\W+", " ", t).strip().lower()[:200])
            if schluessel in gesehen:
                continue
            gesehen.add(schluessel)
            zeile = kat[(folienart, feld)]
            pruefen = []
            if zeile["befuellbar"] == "teilweise":
                pruefen.append("Feld nur teilweise befuellbar")
            if zeile["sicherheit"] == "niedrig":
                pruefen.append("Einschaetzung unsicher")
            if sicher == "mittel":
                pruefen.append("Regel mittelsicher")
            if r.get("suspect"):
                pruefen.append(norm(r["suspect"]))
            if r.get("needs_question"):
                pruefen.append("offener Punkt")
            treffer.append({
                "folienart": folienart,
                "feld_id": feld,
                "feldname": zeile["feldname"],
                "regel": name,
                "sicherheit_regel": sicher,
                "sicherheit_feld": zeile["sicherheit"],
                "befuellbar": zeile["befuellbar"],
                "pruefen": "; ".join(pruefen),
                "station": a.station or "nicht zugeordnet",
                "station_quelle": "Mensch" if a.station else "-",
                "binding": r.get("binding", ""),
                "operator": r.get("operator") or "",
                "value": r.get("value"),
                "unit": r.get("unit") or "",
                "datei": r["_datei"],
                "seite": r["_seite"],
                "text": t,
                "auszug": (AUSZUG_JE_REGEL[name](t)
                           if name in AUSZUG_JE_REGEL else ""),
                "auch_passend": (auch_passend
                                 if name in TEXTFELD_VORRANG else ""),
                "req_id": r.get("req_id", ""),
                "subject_key": r.get("subject_key", ""),
            })

    # Kennungen vergeben: stabile Reihenfolge nach Folienart, Feld, Seite
    treffer.sort(key=lambda z: (z["folienart"], z["feld_id"],
                                str(z["seite"]).zfill(5), z["text"][:60]))
    for i, z in enumerate(treffer, start=a.ab_kennung):
        z["kennung"] = f"R-{i:04d}"

    # ---------------------------------------------------------------- Zahlen
    je_feld = collections.Counter((z["folienart"], z["feld_id"])
                                  for z in treffer)
    je_regel = collections.Counter(z["regel"] for z in treffer)
    zu_pruefen = sum(1 for z in treffer if z["pruefen"])
    betrachtet = len(rows)
    verwertet = betrachtet - muell - zu_kurz - ohne_feld

    L = [f"ZUORDNUNG  {datetime.now():%Y-%m-%d %H:%M}", "=" * 68, ""]
    if a.datei:
        L.append(f"  nur Dateien mit             '{a.datei}'")
    if a.station:
        L.append(f"  Station (Angabe Mensch)     {a.station}")
    else:
        L.append("  Station                     nicht zugeordnet "
                 "(keine Stationsliste vorhanden)")
    if a.nur_sichere:
        L.append("  nur Felder mit befuellbar=ja")
    L += [f"  Anforderungen betrachtet    {betrachtet}",
          f"  davon als Muell verworfen   {muell}",
          f"  davon zu kurz weg           {zu_kurz}",
          f"  davon kein passendes Feld   {ohne_feld}",
          f"  davon mit passendem Feld    {verwertet} "
          f"({100*verwertet/betrachtet:.0f} %)",
          "",
          f"  Eintraege fuer die Akte     {len(treffer)}",
          "    (ein Satz kann Tabelle + Textfeld bekommen; gleicher Wortlaut",
          "     aus mehreren Dateien/Seiten zaehlt nur einmal)",
          f"  Textfeld-Doppel vermieden   {mehrfach_text}",
          f"  davon als 'pruefen' markiert {zu_pruefen} "
          f"({100*zu_pruefen/max(len(treffer),1):.0f} %)",
          f"  Zielfelder im Katalog       {len(kat)}",
          f"  Zielfelder tatsaechlich befuellt {len(je_feld)}",
          "", "  Je Zielfeld:"]
    for (fa, fid), n in je_feld.most_common():
        L.append(f"    {n:>5}  {fa} / {fid}")
    L += ["", "  Je Regel:"]
    for k, n in je_regel.most_common():
        L.append(f"    {n:>5}  {k}")
    if signal_zaehler:
        L += ["", "  Warum verworfen wurde (Mehrfachnennung moeglich):"]
        for k, n in signal_zaehler.most_common():
            L.append(f"    {n:>5}  {k}")
    L += ["", "  Lesart:",
          "    Verworfen wird erst, wenn MINDESTENS ZWEI Signale sprechen",
          "    (oder der Satz eindeutig Kopfzeile bzw. Verzeichnis ist). Alle",
          "    verworfenen Saetze stehen mit Begruendung in",
          "    out/zuordnung_verworfen.xlsx - bitte stichprobenartig pruefen,",
          "    ob dort etwas Brauchbares gelandet ist.",
          "",
          "    'kein passendes Feld' ist kein Fehler. Ein Lastenheft enthaelt",
          "    Abnahme-, Dokumentations- und Rechtsanforderungen, fuer die es",
          "    in der Stationsakte kein Feld gibt.",
          "",
          "    'pruefen' heisst: die Zuordnung ist plausibel, aber nicht",
          "    sicher. Ein Mensch muss sie bestaetigen, bevor sie in die",
          "    Akte geschrieben wird.",
          "",
          "    Die STATION ist nicht abgeleitet. Das Lastenheft ist nach",
          "    Themen gegliedert, die Stationsakte nach Stationen. Ohne",
          "    Stationsliste kann kein Programm das entscheiden.",
          "", "Diese Datei enthaelt keine Dateinamen und keine Saetze."]

    OUT.mkdir(exist_ok=True)
    (OUT / "zuordnung_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
    # maschinenlesbar fuer die Weboberflaeche - nur Zahlen, keine Saetze
    (OUT / "zuordnung_zahlen.json").write_text(json.dumps({
        "betrachtet": betrachtet, "muell": muell, "zu_kurz": zu_kurz,
        "ohne_feld": ohne_feld, "mit_feld": verwertet,
        "eintraege": len(treffer), "pruefen": zu_pruefen,
    }, indent=1), encoding="utf-8")
    for z in L:
        print(ascii_safe(z))

    # ------------------------------------------------------------- jsonl
    with (OUT / "zuordnung.jsonl").open("w", encoding="utf-8") as f:
        for z in treffer:
            f.write(json.dumps(z, ensure_ascii=False) + "\n")

    # ------------------------------------------------------------- Excel
    try:
        import xlsxwriter
    except ImportError:
        print()
        print("HINWEIS: xlsxwriter fehlt, Excel wurde nicht geschrieben.")
        return 0

    ziel = OUT / "zuordnung.xlsx"

    # Schutz fuer menschliche Arbeit: Wer in der alten zuordnung.xlsx schon
    # Freigaben eingetragen hat, verliert sie sonst beim naechsten Lauf.
    # Die alte Datei wird deshalb vorher gesichert.
    if ziel.exists():
        anzahl = freigaben_in_datei(ziel)
        if anzahl:
            sicherung = OUT / f"zuordnung_freigaben_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
            try:
                import shutil
                shutil.copy2(ziel, sicherung)
                print()
                if anzahl > 0:
                    print(f"HINWEIS: Die alte zuordnung.xlsx enthielt {anzahl} "
                          f"Freigaben. Gesichert als:")
                else:
                    print("HINWEIS: Die alte zuordnung.xlsx konnte nicht auf "
                          "Freigaben geprueft werden. Vorsichtshalber "
                          "gesichert als:")
                print("  ", ascii_safe(sicherung.name))
                print("   Fuer den Schreiber als out\\zuordnung_geprueft.xlsx "
                      "speichern.")
            except OSError as e:
                print("FEHLER: Sicherung der Freigaben fehlgeschlagen:",
                      ascii_safe(e))
                print("Abbruch, damit keine Freigaben verloren gehen.")
                return 1

    wb = xlsxwriter.Workbook(str(ziel))

    kopf = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1,
                          "text_wrap": True, "valign": "top"})
    zelle = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    warn = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                          "bg_color": "#FDF4E3"})

    ws = wb.add_worksheet("Zuordnung")
    spalten = ["Kennung", "Folienart", "Feld", "Feldname", "Prüfen",
               SPALTE_FREIGABE,
               "Regel", "Verbindlichkeit", "Operator", "Wert", "Einheit",
               "Station", "Datei", "Seite", "Wortlaut", "Auszug (Tabelle)",
               "Auch passend", "Bemerkung"]
    breiten = [10, 26, 24, 26, 30, 12, 20, 13, 9, 10, 9, 22, 34, 7, 90, 30, 30, 24]
    for i, (s, b) in enumerate(zip(spalten, breiten)):
        ws.write(0, i, s, kopf)
        ws.set_column(i, i, b)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(treffer), 1), len(spalten) - 1)
    for n, z in enumerate(treffer, start=1):
        f = warn if z["pruefen"] else zelle
        ws.write_row(n, 0, [
            z["kennung"], z["folienart"], z["feld_id"], z["feldname"],
            z["pruefen"] or "sicher", "",
            z["regel"], z["binding"],
            z["operator"], z["value"] if z["value"] is not None else "",
            z["unit"], z["station"], z["datei"], z["seite"], z["text"],
            z.get("auszug", ""), z.get("auch_passend", ""), "",
        ], f)
    if treffer:
        sp = spalten.index(SPALTE_FREIGABE)
        ws.data_validation(1, sp, len(treffer), sp, {
            "validate": "list", "source": ["ja", "nein"],
            "input_title": "Freigabe",
            "input_message": "ja = in die Akte, nein = weglassen, "
                             "leer = noch nicht entschieden"})
    try:
        wb.close()
    except Exception as e:                       # Datei in Excel geoeffnet
        print()
        print("FEHLER:", ascii_safe(ziel.name),
              "konnte nicht geschrieben werden (in Excel offen?)",
              ascii_safe(e))
        return 1

    # Saetze ohne passendes Feld - damit ein Mensch pruefen kann, was die
    # Regeln uebersehen. Nur so laesst sich die Trefferquote ueberhaupt
    # schaetzen: Was genommen wurde, sieht man in der Akte; was fehlt, nur hier.
    oz = OUT / "zuordnung_ohne_feld.xlsx"
    try:
        wbo = xlsxwriter.Workbook(str(oz))
        kopf_wbo = wbo.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1, "text_wrap": True, "valign": "top"})
        zelle_wbo = wbo.add_format({"text_wrap": True, "valign": "top", "border": 1})
        wso = wbo.add_worksheet("Ohne Feld")
        oz_spalten = ["Gehört in die Akte? (Feld eintragen)", "Verbindlichkeit",
                      "Datei", "Seite", "Wortlaut"]
        for i, (sp_, br) in enumerate(zip(oz_spalten, [34, 13, 34, 7, 110])):
            wso.write(0, i, sp_, kopf_wbo)
            wso.set_column(i, i, br)
        wso.freeze_panes(1, 0)
        wso.autofilter(0, 0, max(len(ohne_feld_liste), 1), len(oz_spalten) - 1)
        for n, v in enumerate(ohne_feld_liste, start=1):
            wso.write_row(n, 0, ["", v["binding"], v["datei"], v["seite"],
                                 v["text"]], zelle_wbo)
        wbo.close()
        print()
        print(f"  {oz}  Saetze ohne Feld - VERTRAULICH")
    except Exception as e:
        print("HINWEIS: zuordnung_ohne_feld.xlsx nicht geschrieben:",
              ascii_safe(e))

    if verworfen:
        vz = OUT / "zuordnung_verworfen.xlsx"
        try:
            wbv = xlsxwriter.Workbook(str(vz))
            kopf_wbv = wbv.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1, "text_wrap": True, "valign": "top"})
            zelle_wbv = wbv.add_format({"text_wrap": True, "valign": "top", "border": 1})
            wsv = wbv.add_worksheet("Verworfen")
            for i, (sp, br) in enumerate(zip(
                    ["Signale", "Verbindlichkeit", "Datei", "Seite", "Wortlaut"],
                    [40, 13, 34, 7, 110])):
                wsv.write(0, i, sp, kopf_wbv)
                wsv.set_column(i, i, br)
            wsv.freeze_panes(1, 0)
            wsv.autofilter(0, 0, len(verworfen), 4)
            for n, v in enumerate(verworfen, start=1):
                wsv.write_row(n, 0, [v["signale"], v["binding"], v["datei"],
                                     v["seite"], v["text"]], zelle_wbv)
            wbv.close()
            print()
            print(f"  {vz}  verworfene Saetze - VERTRAULICH")
        except PermissionError:
            print("HINWEIS: zuordnung_verworfen.xlsx ist in Excel offen.")

    print()
    print("geschrieben:")
    print("  out\\zuordnung_zahlen.txt   nur Zahlen - darf geteilt werden")
    print("  out\\zuordnung.jsonl        Eingabe fuer den Schreiber")
    print("  out\\zuordnung.xlsx         mit Wortlauten - VERTRAULICH")
    print()
    print("Die gelb hinterlegten Zeilen brauchen eine menschliche Pruefung.")
    print("Freigabe: ja = bestaetigen, nein = ablehnen, leer = unveraendert.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
