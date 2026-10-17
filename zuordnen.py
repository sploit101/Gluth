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
# Ein Satz kann in mehrere Felder gehen - "Toleranz der Abisolierlaenge
# +/- 0,2 mm" gehoert in die Faehigkeitstabelle UND in die ausfuehrliche
# Aufgabenbeschreibung. Das ist gewollt.
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
        lambda t, r: bool(VERFAHREN.search(t)), "hoch")

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
        lambda t, r: (bool(PROZESS_NAME.search(t))
                      and not VERFAHREN.search(t)
                      and not re.search(r"measur\w*\s+(equipment|system)", t, re.I)
                      and bool(
            re.search(r"\b(monitor\w*|check\w*|inspect\w*|test\w*|"
                      r"[uü]berwach\w*|pr[uü]f\w*|mess\w*)\b", t, re.I))),
        "mittel")

    # --- Aufgabenbeschreibung: lange verbindliche Saetze, die wirklich die
    #     Aufgabe beschreiben
    add("aufgabe_ausfuehrlich", "Aufgabenstellung_ausführlich", "ph11",
        lambda t, r: (len(t) > 120
                      and r.get("binding") in ("MUSS", "VERBOT")
                      and not KEINE_AUFGABE.search(t)
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
        lambda t, r: bool(KALIBRIEREN.search(t)) and len(t) > 60, "mittel")

    # --- Varianten: hier wird kein Satz eingetragen, sondern ein Ja/Nein
    #     belegt. Der Satz ist der Nachweis.
    add("varianten_ja", "Bauteile", "ph22",
        lambda t, r: bool(VARIANTE.search(t)), "hoch")

    return R


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
            continue

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
          f"  davon zugeordnet            {verwertet} "
          f"({100*verwertet/betrachtet:.0f} %)",
          "",
          f"  Zuordnungen insgesamt       {len(treffer)}",
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
    try:
        wb = xlsxwriter.Workbook(str(ziel))
    except PermissionError:
        print()
        print("FEHLER:", ascii_safe(ziel.name), "ist in Excel offen.")
        return 1

    kopf = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1,
                          "text_wrap": True, "valign": "top"})
    zelle = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    warn = wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                          "bg_color": "#FDF4E3"})

    ws = wb.add_worksheet("Zuordnung")
    spalten = ["Kennung", "Folienart", "Feld", "Feldname", "Prüfen",
               "Regel", "Verbindlichkeit", "Operator", "Wert", "Einheit",
               "Station", "Seite", "Wortlaut", "Bemerkung"]
    breiten = [10, 26, 24, 26, 30, 20, 13, 9, 10, 9, 22, 7, 90, 24]
    for i, (s, b) in enumerate(zip(spalten, breiten)):
        ws.write(0, i, s, kopf)
        ws.set_column(i, i, b)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(treffer), 1), len(spalten) - 1)
    for n, z in enumerate(treffer, start=1):
        f = warn if z["pruefen"] else zelle
        ws.write_row(n, 0, [
            z["kennung"], z["folienart"], z["feld_id"], z["feldname"],
            z["pruefen"] or "sicher", z["regel"], z["binding"],
            z["operator"], z["value"] if z["value"] is not None else "",
            z["unit"], z["station"], z["seite"], z["text"], "",
        ], f)
    wb.close()

    if verworfen:
        vz = OUT / "zuordnung_verworfen.xlsx"
        try:
            wbv = xlsxwriter.Workbook(str(vz))
            wsv = wbv.add_worksheet("Verworfen")
            for i, (sp, br) in enumerate(zip(
                    ["Signale", "Verbindlichkeit", "Datei", "Seite", "Wortlaut"],
                    [40, 13, 34, 7, 110])):
                wsv.write(0, i, sp, kopf)
                wsv.set_column(i, i, br)
            wsv.freeze_panes(1, 0)
            wsv.autofilter(0, 0, len(verworfen), 4)
            for n, v in enumerate(verworfen, start=1):
                wsv.write_row(n, 0, [v["signale"], v["binding"], v["datei"],
                                     v["seite"], v["text"]], zelle)
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
    print("Die gelb hinterlegten Zeilen in der Excel sind die, die ein")
    print("Mensch bestaetigen muss. Spalte 'Bemerkung' ist fuer dich leer.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
