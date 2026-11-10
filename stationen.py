#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stationen.py — beantwortet die Frage: "Gilt dieser Satz fuer MEINE Station?"

Warum es das gibt
-----------------
Das Lastenheft ist nach Themen gegliedert, die Stationsakte nach Stationen.
Bisher bekam jede Zeile einfach den Stationsnamen angehaengt, den ein Mensch
eingegeben hat. Ein Konstrukteur fuer Station 20 bekam deshalb 39 Zitate aus
der ganzen Anlage, von denen 1-2 seine Station betrafen.

Dieses Modul bewertet jeden Satz in vier Klassen:

    Station          nennt die Stationsnummer ODER ein Verfahren dieser Station
    Anlagenweit      gilt fuer alle (Sicherheit, Faehigkeit, Dokumentation,
                     Abnahme, Elektrik ...) oder sagt ausdruecklich "alle Stationen"
    Andere Station   nennt eine andere Stationsnummer oder ein Verfahren, das
                     nicht zu dieser Station gehoert
    Unklar           kein Hinweis -> ein Mensch muss entscheiden

Grundsaetze (wie im Rest des Projekts)
--------------------------------------
1. Kein Sprachmodell. Reine Regeln. Jede Bewertung nennt ihren Grund
   ("Verfahren Abisolieren", "Station 30 genannt") und ist damit bestreitbar.
2. Es wird nichts erfunden. Die Kurzfassung ist aus Bindung, Thema und den
   woertlich gefundenen Kennwerten zusammengesetzt; der Originalsatz bleibt
   der Beleg.
3. Ist keine echte Station bekannt ("Gesamtanlage ..."), wird NICHT bewertet.
   Dann gibt es nur das Thema je Satz. Besser keine Bewertung als eine geratene.
4. Sobald eine Stationsliste vorliegt (Nummer;Name;Stichworte), wird jeder
   Satz gegen ALLE Stationen gestellt und der Treffer benannt.

    python stationen.py --station "Station 20 Abisolieren" --text "Die Abisolierlaenge ist mit +/-0,2 mm einzuhalten."
    python stationen.py --selftest
    python stationen.py --vorschlag out/zuordnung.jsonl      # Stationsliste vorschlagen
"""

from __future__ import annotations

import argparse
import collections
import csv
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Textnormalisierung: Umlaute falten, damit "Zuführung" und "Zufuehrung" gleich
# behandelt werden. Alle Muster unten stehen deshalb in gefalteter Schreibweise.
# ---------------------------------------------------------------------------
_FALTEN = str.maketrans({"ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
                         "Ä": "ae", "Ö": "oe", "Ü": "ue"})


def falten(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").translate(_FALTEN).lower()).strip()


@dataclass(frozen=True)
class Thema:
    key: str
    de: str
    art: str                      # "prozess" = eigene Station moeglich, "quer" = gilt breit
    muster: re.Pattern


def _t(key, de, art, muster) -> Thema:
    return Thema(key, de, art, re.compile(muster, re.I))


# ---------------------------------------------------------------------------
# LEXIKON. Zweisprachig, weil die Lastenhefte englisch sind und die Stationen
# deutsch heissen. Wer ein Verfahren ergaenzen will, haengt hier einen Eintrag
# an - an der Logik darunter aendert sich nichts.
# ---------------------------------------------------------------------------
PROZESSE: list[Thema] = [
    _t("ABISOLIEREN", "Abisolieren", "prozess",
       r"abisolier|\bstrip(ping|ped|s)?\b|insulation\s+strip|strip(ping)?\s+length"),
    _t("CRIMPEN", "Crimpen", "prozess",
       r"crimp|pull[\s-]?(out|off)\s+force|auszugskraft|abzugskraft|quetschverbind"),
    _t("ABLAENGEN", "Ablängen / Schneiden", "prozess",
       r"ablaeng|\bcut(ting|s)?\b|cut[\s-]?to[\s-]?length|zuschnitt|\bschneid|cable\s+length|kabellaenge|laengenwiederhol|length\s+(repeat|accuracy)"),
    _t("EINLEGEN", "Einlegen / Positionieren", "prozess",
       r"\binsert(ion|ing|ed)?\b|einleg|positionier|\bpositioning\b|\bplacement\b|\bplace(d)?\s+(in|into|on)\b|einfuehr"),
    _t("EINPRESSEN", "Einpressen", "prozess",
       r"press[\s-]?(in|fit)\b|einpress|press(ing)?\s+force|kraft[\s/-]*weg|force[\s/-]*(displacement|travel|stroke)"),
    _t("SICHTPRUEFUNG", "Sichtprüfung / Bildverarbeitung", "prozess",
       r"\bvision\b|\bcamera\b|kamera|sichtpr|bildverarbeit|optical\s+inspect|visual\s+inspect|360[\s-]*(grad|degree|deg)"),
    _t("ZUFUEHRUNG", "Zuführung / Bereitstellung", "prozess",
       r"\bfeed(ing|er|s)?\b|zufuehr|magazin|\bconveyor|foerder|bereitstell|\bhopper\b|\breel\b|\bspool\b|\btrommel\b"),
    _t("VERSCHRAUBEN", "Verschrauben", "prozess",
       r"schraub|\bscrew|\btorque\b|drehmoment|anzugsmoment"),
    _t("FUEGEN", "Schweißen / Löten / Kleben", "prozess",
       r"\bweld|\bsolder|schweiss|loet|\bglue|\badhesive|kleb"),
    _t("BESCHRIFTEN", "Beschriften / Kennzeichnen", "prozess",
       r"\blabel|\bmarking\b|laser[\s-]?mark|beschrift|markier|kennzeichn|data[\s-]?matrix|\bengrav"),
    _t("ELEKTR_PRUEFUNG", "Elektrische Prüfung", "prozess",
       r"continuity|durchgangspr|hochspannungspr|high[\s-]?voltage\s+test|\bhv[\s-]?test|elektrische\s+pr|electrical\s+test|isolationswiderstand"),
    _t("AUSSCHLEUSEN", "Ausschleusen / Verpacken", "prozess",
       r"ausschleus|\breject|\bscrap\b|schlechtteil|\bng[\s-]?part|\bgood\s+part|gutteil|verpack|\bpackag|\bpacking\b|output\s+station|ablage"),
    _t("HANDLING", "Handhabung / Greifer / Roboter", "prozess",
       r"\bgripper|greifer|\brobot|handhabung|\bhandling\b|pick[\s-]?and[\s-]?place|\bgantry\b|portal"),
]

QUERSCHNITT: list[Thema] = [
    _t("FAEHIGKEIT", "Fähigkeitsnachweis", "quer",
       r"\bmsa\b|\bmfu\b|\bmgf\b|\bpfu\b|\bcgk?\b|\bcpk?\b|\bgrr\b|gauge\s+r\s*&?\s*r|capabilit|faehigkeit|measurement\s+system\s+analysis"),
    _t("SICHERHEIT", "Sicherheit", "quer",
       r"\bsafety\b|sicherheit|\be-?stop\b|emergency\s+stop|not[\s-]?halt|not[\s-]?aus|light\s+curtain|lichtvorhang|lichtgitter|schutz(tuer|zaun|einricht|haube)|protective|guard(ing)?\b|performance\s+level|13849|machiner(y|ie)\s+directive|maschinenrichtlinie|\bce[\s-](mark|certif|kennz)|risk\s+assessment|risikobeurteil"),
    _t("DOKUMENTATION", "Dokumentation", "quer",
       r"documentation|dokumentation|operating\s+instruction|betriebsanleitung|\bmanual\b|handbuch|circuit\s+diagram|schaltplan|\bdrawings?\b|zeichnung|spare\s+parts?|ersatzteil|maintenance\s+plan|wartungsplan"),
    _t("ABNAHME", "Abnahme / Inbetriebnahme", "quer",
       r"acceptance|abnahme|\bfat\b|\bsat\b|run[\s-]?off|commissioning|inbetriebnahme|buy[\s-]?off|ramp[\s-]?up|\bhochlauf"),
    _t("STEUERUNG", "Elektrik / Steuerung / Schnittstellen", "quer",
       r"control\s+cabinet|switch\s*cabinet|schaltschrank|\bplc\b|\bsps\b|profinet|profibus|ethernet|\bopc\b|\bmes\b|interface|schnittstelle|\bvoltage\b|spannung|\bsoftware\b|\bhmi\b|bedienung|\bterminal\s+block"),
    _t("MEDIEN", "Medien / Druckluft", "quer",
       r"compressed\s+air|druckluft|pneumat|\bcoolant|kuehlmittel|\bvacuum|vakuum"),
    _t("UMGEBUNG", "Aufstellung / Umgebung", "quer",
       r"\bnoise|laerm|db\s*\(a\)|ambient|umgebung|aufstell|footprint|\bhall\b|floor|dimension|abmessung|\btemperature|temperatur|humidity|feuchte"),
    _t("TAKT", "Taktzeit / Ausbringung", "quer",
       r"cycle\s+time|taktzeit|throughput|ausbringung|parts\s+per\s+(hour|minute)|stueck\s*/\s*(h|std)|\boee\b|availability|verfuegbarkeit"),
    _t("RUECKVERFOLG", "Rückverfolgbarkeit / Daten", "quer",
       r"traceab|rueckverfolg|data\s+storage|datenspeicher|archiv|\blogging\b|protokollier"),
    _t("SERVICE", "Service / Gewährleistung / Schulung", "quer",
       r"\btraining\b|schulung|warranty|gewaehrleistung|guarantee|maintenance|wartung|\bservice\b|\bliability|haftung"),
]

ALLE_THEMEN = {t.key: t for t in PROZESSE + QUERSCHNITT}

# Ausdruecklich anlagenweite Aussagen
GLOBAL = re.compile(
    r"\b(all|each|every)\s+(the\s+)?(stations?|machines?|workstations?|units?)\b|"
    r"\b(entire|whole|complete|overall)\s+(line|plant|system|installation|machine)\b|"
    r"\b(line|plant)[\s-]?wide\b|"
    r"\balle\s+stationen\b|\bjede\s+station\b|\bsaemtliche\s+stationen\b|"
    r"\bgesamte\s+anlage\b|\bgesamtanlage\b|\banlagenweit\b|\bgesamtlinie\b", re.I)

# Stationsnummern im Text: "Station 20", "stations 20/30", "Stn. 40", "OP 50"
STATION_NR = re.compile(
    r"\b(?:station(?:en|s)?|stn|op|arbeitsplatz)\.?[\s-]*"
    r"(\d{1,3}(?:\s*(?:,|/|und|and|&|bis|to|-)\s*\d{1,3})*)\b", re.I)

# Woerter, die einen Stationsnamen NICHT kennzeichnen. Steht nur so etwas im
# Namen ("Gesamtanlage (Stationszuordnung folgt)"), ist keine Station bekannt.
GENERISCH = {
    "station", "stationen", "stn", "op", "arbeitsplatz", "und", "and", "der",
    "die", "das", "mit", "fuer", "zur", "zum", "von", "the", "of", "for", "with",
    "gesamtanlage", "gesamt", "anlage", "linie", "line", "alle", "stationszuordnung",
    "zuordnung", "folgt", "nicht", "zugeordnet", "unbekannt", "projekt", "plant",
    "whole", "entire", "tbd", "tba", "offen", "noch", "unklar", "komplett",
    "bereich", "modul", "einheit", "unit",
}


def _nummern(roh: str) -> set[str]:
    """'20/30' -> {20, 30}; '10 bis 30' -> {10,20,30}? Nein: nur genannte Zahlen
    und kleine Bereiche in Zehnerschritten werden nicht erfunden."""
    zahlen = [int(x) for x in re.findall(r"\d+", roh)]
    raus = {str(z) for z in zahlen}
    if re.search(r"\b(bis|to)\b|-", roh) and len(zahlen) == 2:
        a, b = zahlen
        if 0 < b - a <= 10:                       # "Station 1-3" -> 1,2,3
            raus |= {str(x) for x in range(a, b + 1)}
    return raus


def stationsnummern(text: str) -> set[str]:
    t = falten(text)
    raus: set[str] = set()
    for m in STATION_NR.finditer(t):
        raus |= _nummern(m.group(1))
    return raus


def themen_in(text: str) -> list[Thema]:
    t = falten(text)
    return [th for th in PROZESSE + QUERSCHNITT if th.muster.search(t)]


# ---------------------------------------------------------------------------
# Profile: was macht eine Station aus?
# ---------------------------------------------------------------------------
@dataclass
class Profil:
    name: str
    nummer: str | None = None
    prozesse: set[str] = field(default_factory=set)
    eigene: list[re.Pattern] = field(default_factory=list)
    stichworte: list[str] = field(default_factory=list)

    @property
    def bewertbar(self) -> bool:
        """Kann man Saetze gegen diese Station pruefen?"""
        return bool(self.nummer or self.prozesse or self.eigene)

    @property
    def kurz(self) -> str:
        return self.name


def profil_aus_name(name: str, stichworte: str | list[str] | None = None,
                    nummer: str | None = None) -> Profil | None:
    """
    Baut das Profil aus dem Stationsnamen, z. B. "Station 20 Abisolieren".
    Gibt None zurueck, wenn der Name keine echte Station benennt
    ("Gesamtanlage (Stationszuordnung folgt)").
    """
    name = (name or "").strip()
    if not name:
        return None
    gefaltet = falten(name)
    p = Profil(name=name)
    if nummer:
        p.nummer = str(nummer).strip()
    else:
        m = re.search(r"\b(?:station|stn|op|st)?\.?\s*(\d{1,3})\b", gefaltet)
        if m:
            p.nummer = m.group(1)

    quelle = gefaltet
    zusatz: list[str] = []
    if isinstance(stichworte, str):
        zusatz = [s.strip() for s in re.split(r"[,;|]", stichworte) if s.strip()]
    elif stichworte:
        zusatz = [str(s).strip() for s in stichworte if str(s).strip()]
    p.stichworte = zusatz
    quelle_alle = quelle + " " + " ".join(falten(z) for z in zusatz)

    belegt = []                                  # Textstellen, die das Lexikon erklaert
    for th in PROZESSE:
        m = th.muster.search(quelle_alle)
        if m:
            p.prozesse.add(th.key)
            belegt.append(m.group(0))

    # Eigene Stichworte: Woerter des Namens, die das Lexikon nicht kennt
    for w in re.findall(r"[a-z]{5,}", quelle):
        if w in GENERISCH or any(b in w or w in b for b in belegt):
            continue
        p.eigene.append(re.compile(r"\b" + re.escape(w[:6]), re.I))
    for z in zusatz:
        zf = falten(z)
        if len(zf) >= 3 and not any(th.muster.search(zf) for th in PROZESSE):
            p.eigene.append(re.compile(r"\b" + re.escape(zf), re.I))
    return p if p.bewertbar else None


def stationsliste_laden(pfad: Path) -> list[Profil]:
    """
    CSV oder JSON mit den Stationen eines Projekts.

    CSV (Semikolon oder Komma, Kopfzeile optional):
        Nummer;Name;Stichworte
        10;Zuführung;Magazin, Vereinzeler
        20;Abisolieren;
        30;Crimpen;

    JSON: [{"nummer": "20", "name": "Abisolieren", "stichworte": ["strip"]}, ...]
    """
    pfad = Path(pfad)
    roh = pfad.read_text(encoding="utf-8-sig")
    eintraege: list[dict] = []
    if pfad.suffix.lower() == ".json":
        daten = json.loads(roh)
        if isinstance(daten, dict):
            daten = daten.get("stationen", [])
        eintraege = [dict(x) for x in daten if isinstance(x, dict)]
    else:
        zeilen = [z for z in roh.splitlines() if z.strip() and not z.lstrip().startswith("#")]
        if not zeilen:
            return []
        delim = ";" if zeilen[0].count(";") >= zeilen[0].count(",") else ","
        leser = list(csv.reader(zeilen, delimiter=delim))
        kopf = [falten(c) for c in leser[0]]
        hat_kopf = any(k in ("nummer", "nr", "name", "station", "stichworte") for k in kopf)
        daten = leser[1:] if hat_kopf else leser
        if not hat_kopf:
            kopf = ["nummer", "name", "stichworte"]
        for z in daten:
            d = {kopf[i] if i < len(kopf) else f"spalte{i}": v.strip()
                 for i, v in enumerate(z)}
            eintraege.append(d)
    profile = []
    for e in eintraege:
        e = {falten(k): v for k, v in e.items()}
        nr = str(e.get("nummer") or e.get("nr") or "").strip() or None
        nm = str(e.get("name") or e.get("station") or "").strip()
        sw = e.get("stichworte") or e.get("keywords") or ""
        if not nm and not nr:
            continue
        voll = f"Station {nr} {nm}".strip() if nr else nm
        p = profil_aus_name(nm, sw, nummer=nr)
        if p is None:
            # Name ohne Lexikontreffer und ohne eigene Woerter: wenigstens die Nummer
            p = Profil(name=voll, nummer=nr)
        p.name = voll
        profile.append(p)
    return profile


# ---------------------------------------------------------------------------
# Bewertung
# ---------------------------------------------------------------------------
@dataclass
class Bewertung:
    geltung: str = ""            # Station | Anlagenweit | Andere Station | Unklar | ""
    warum: str = ""
    themen: list[str] = field(default_factory=list)       # deutsche Bezeichnungen
    themen_keys: list[str] = field(default_factory=list)
    andere: str = ""             # Name der vermutlich gemeinten Station

    @property
    def thema(self) -> str:
        return " / ".join(self.themen) if self.themen else "Allgemein"


def _passt(profil: Profil, text_gefaltet: str, nummern: set[str], keys: set[str]) -> str:
    """Grund, warum der Text zu diesem Profil passt - oder leer."""
    if profil.nummer and profil.nummer in nummern:
        return f"Station {profil.nummer} genannt"
    gemeinsam = [ALLE_THEMEN[k].de for k in sorted(keys & profil.prozesse)]
    if gemeinsam:
        return "Verfahren " + ", ".join(gemeinsam)
    for muster in profil.eigene:
        m = muster.search(text_gefaltet)
        if m:
            return f"Stichwort '{m.group(0)}'"
    return ""


def bewerte(text: str, profil: Profil | None,
            alle: list[Profil] | None = None) -> Bewertung:
    """
    Bewertet einen Satz gegenueber der Station `profil`.
    Ohne Profil wird nur das Thema bestimmt (geltung bleibt leer).
    """
    t = falten(text)
    themen = themen_in(text)
    prozess_keys = {th.key for th in themen if th.art == "prozess"}
    quer_keys = {th.key for th in themen if th.art == "quer"}
    # Reihenfolge: erst Verfahren, dann Querschnittsthemen
    ordnung = [th for th in themen if th.art == "prozess"] + \
              [th for th in themen if th.art == "quer"]
    b = Bewertung(themen=[th.de for th in ordnung[:2]],
                  themen_keys=[th.key for th in ordnung])
    if profil is None or not profil.bewertbar:
        return b

    nummern = stationsnummern(text)
    global_marker = bool(GLOBAL.search(t))

    # Gibt es in der Stationsliste eine bessere Heimat fuer den Satz?
    heimat = ""
    if alle:
        beste = (0, "")
        for q in alle:
            if q.name == profil.name:
                continue
            grund = _passt(q, t, nummern, prozess_keys)
            if not grund:
                continue
            punkte = 10 if (q.nummer and q.nummer in nummern) else 1 + len(
                prozess_keys & q.prozesse)
            if punkte > beste[0]:
                beste = (punkte, f"{q.name} ({grund})")
        heimat = beste[1]

    # 1. Eine Stationsnummer im Text entscheidet
    if nummern and not global_marker:
        if profil.nummer and profil.nummer in nummern:
            b.geltung, b.warum = "Station", f"Station {profil.nummer} genannt"
        elif profil.nummer:
            b.geltung = "Andere Station"
            b.warum = "Station " + "/".join(sorted(nummern, key=int)) + " genannt"
            b.andere = heimat or b.warum
        # Profil ohne Nummer: Nummer im Text beweist hier nichts -> weiter unten
        if b.geltung:
            return b

    # 2. Ausdruecklich "alle Stationen"
    if global_marker:
        b.geltung, b.warum = "Anlagenweit", "Text nennt ausdruecklich alle Stationen / die Gesamtanlage"
        return b

    # 3. Passt das Verfahren oder ein Stichwort zu dieser Station?
    grund = _passt(profil, t, set(), prozess_keys)
    if grund:
        b.geltung, b.warum = "Station", grund
        return b

    # 4. Anderes Verfahren oder andere Station aus der Liste
    if heimat:
        b.geltung, b.warum, b.andere = "Andere Station", "passt besser zu " + heimat, heimat
        return b
    fremd = [ALLE_THEMEN[k].de for k in sorted(prozess_keys - profil.prozesse)]
    if fremd and profil.prozesse:
        b.geltung = "Andere Station"
        b.warum = "Verfahren " + ", ".join(fremd) + " gehoert nicht zu dieser Station"
        b.andere = ", ".join(fremd)
        return b

    # 5. Querschnittsthema ohne Verfahren
    if quer_keys:
        b.geltung = "Anlagenweit"
        b.warum = "Querschnittsthema " + ", ".join(ALLE_THEMEN[k].de for k in sorted(quer_keys))
        return b

    b.geltung, b.warum = "Unklar", "kein Hinweis auf eine Station oder ein Thema"
    return b


# ---------------------------------------------------------------------------
# Kurzfassung auf Deutsch. KEINE Uebersetzung des Satzes, sondern ein
# Stichwortzettel aus Bindung, Thema und den woertlich gefundenen Kennwerten.
# ---------------------------------------------------------------------------
BINDUNG_DE = {"MUSS": "Muss", "SOLL": "Soll", "KANN": "Kann",
              "VERBOT": "Verbot", "INFO": "Info"}

OPERATOR_DE = {">=": "≥", "<=": "≤", "=>": "≥", "=<": "≤", "==": "=",
               "min": "mind.", "max": "max.", "approx": "ca.", "~": "ca."}

# (Muster auf gefaltetem Text, deutscher Fachbegriff)
FACHBEGRIFFE: list[tuple[re.Pattern, str]] = [(re.compile(m, re.I), de) for m, de in [
    (r"strip(ping)?\s+length|abisolierlaenge", "Abisolierlänge"),
    (r"crimp\s+height|crimphoehe", "Crimphöhe"),
    (r"pull[\s-]?(out|off)\s+force|auszugskraft|abzugskraft", "Auszugskraft"),
    (r"press[\s-]?(in|fit)\s+force|einpresskraft", "Einpresskraft"),
    (r"cycle\s+time|taktzeit", "Taktzeit"),
    (r"repeat(ing|ability)?\s+accuracy|wiederholgenauigkeit|repeatability", "Wiederholgenauigkeit"),
    (r"length\s+(accuracy|tolerance)", "Längentoleranz"),
    (r"\btolerance", "Toleranz"),
    (r"calibrat", "Kalibrierung"),
    (r"reference\s+(sample|part|master)|referenzmuster", "Referenzmuster"),
    (r"product\s+variant|\bvariants?\b", "Produktvariante"),
    (r"measurement\s+system\s+analysis|\bmsa\b", "Messsystemanalyse (MSA)"),
    (r"machine\s+capability|\bmfu\b", "Maschinenfähigkeit (MFU)"),
    (r"process\s+capability|\bpfu\b", "Prozessfähigkeit (PFU)"),
    (r"\bmgf\b|measuring\s+(equipment|device)\s+capability", "Messmittelfähigkeit (MGF)"),
    (r"gauge\s+r\s*&?\s*r|\bgrr\b", "Messsystem-Streuung (GRR)"),
    (r"\bcgk\b", "Cgk (Maschinenfähigkeitsindex, zentriert)"),
    (r"\bcpk\b", "Cpk (Prozessfähigkeitsindex, zentriert)"),
    (r"light\s+curtain", "Lichtvorhang"),
    (r"emergency\s+stop|\be-?stop\b", "Not-Halt"),
    (r"control\s+cabinet|switch\s*cabinet", "Schaltschrank"),
    (r"compressed\s+air", "Druckluft"),
    (r"operating\s+instructions", "Betriebsanleitung"),
    (r"spare\s+parts?", "Ersatzteile"),
    (r"acceptance", "Abnahme"),
    (r"commissioning", "Inbetriebnahme"),
    (r"warranty", "Gewährleistung"),
    (r"\bcables?\b|\bwires?\b", "Kabel / Leitung"),
    (r"connector", "Stecker"),
    (r"\bgripper", "Greifer"),
    (r"vision\s+system|camera", "Bildverarbeitung / Kamera"),
    (r"traceab", "Rückverfolgbarkeit"),
    (r"\breject|\bscrap\b|\bng[\s-]?part", "Ausschuss (NIO-Teil)"),
    (r"good\s+part", "Gutteil (IO)"),
    (r"cleanliness", "Sauberkeit"),
    (r"insulation", "Isolation"),
    (r"conductor|strand", "Leiter / Litze"),
    (r"terminal", "Kontakt / Klemme"),
    (r"\bseal\b", "Dichtung"),
    (r"data\s+matrix|barcode|\bqr\b", "Datamatrix / Barcode"),
]]


def fachbegriffe(text: str, maximal: int = 6) -> list[str]:
    t = falten(text)
    raus: list[str] = []
    for muster, de in FACHBEGRIFFE:
        if muster.search(t) and de not in raus:
            raus.append(de)
        if len(raus) >= maximal:
            break
    return raus


def zahl_de(wert) -> str:
    try:
        return f"{float(wert):g}".replace(".", ",")
    except (TypeError, ValueError):
        return str(wert)


def kennwert_text(operator, value, unit, value_max=None) -> str:
    """Strukturierter Messwert als kurzer deutscher Ausdruck, z. B. '≥ 1,33'."""
    if value is None:
        return ""
    op = OPERATOR_DE.get(str(operator or "").strip().lower(), str(operator or "").strip())
    teile = [op, zahl_de(value)]
    if value_max is not None:
        teile += ["bis", zahl_de(value_max)]
    if unit:
        teile.append(str(unit))
    return " ".join(t for t in teile if t)


CAP_WERT = re.compile(
    r"\b(c[gmp]k?)\b\s*(>=|<=|≥|≤|=|>|<)?\s*(\d+(?:[.,]\d+)?)", re.I)
TOL_WERT = re.compile(
    r"(±|\+\s*/\s*-)\s*(\d+(?:[.,]\d+)?)\s*(mm|µm|um|%|s|sec|nm|bar|°)(?![a-z])", re.I)


def kompakt(text: str) -> dict:
    """
    Kennwerte als kurze, vergleichbare Ausdruecke:
        'Cgk >= 1.33'            -> 'Cgk ≥ 1,33'
        'tolerance of +/- 0.2 mm' -> '± 0,2 mm'
    Es werden nur Zahlen uebernommen, die woertlich im Satz stehen.
    """
    t = str(text or "")
    kennwert, toleranz = [], []
    for m in CAP_WERT.finditer(t):
        name = m.group(1)
        name = name[0].upper() + name[1:].lower()
        op = OPERATOR_DE.get(m.group(2) or "", m.group(2) or "")
        zahl = m.group(3).replace(".", ",")
        x = " ".join(p for p in (name, op, zahl) if p)
        if x not in kennwert:
            kennwert.append(x)
    for m in TOL_WERT.finditer(t):
        x = f"± {m.group(2).replace('.', ',')} {m.group(3).replace('um', 'µm')}"
        if x not in toleranz:
            toleranz.append(x)
    return {"kennwert": kennwert, "toleranz": toleranz}


def kurzfassung(bindung: str, bewertung: Bewertung, kennwerte: list[str] | None = None,
                begriffe: list[str] | None = None) -> str:
    """
    'Muss · Abisolieren · Toleranz ±0,2 mm'. Nur aus gefundenen Bestandteilen;
    fehlt etwas, fehlt es auch hier.
    """
    teile = [BINDUNG_DE.get(str(bindung or "").upper(), "")]
    if bewertung.themen:
        teile.append(", ".join(bewertung.themen[:2]))
    elif begriffe:
        teile.append(", ".join(begriffe[:2]))
    kw = [k for k in (kennwerte or []) if k]
    if kw:
        teile.append("; ".join(dict.fromkeys(kw))[:90])
    return " · ".join(t for t in teile if t)


# ---------------------------------------------------------------------------
# Stationsvorschlag: Welche Verfahren kommen im Lastenheft vor?
# ---------------------------------------------------------------------------
def vorschlag(saetze: list[str]) -> list[dict]:
    """Zaehlt je Verfahren die Saetze. Grundlage fuer die Stationsliste."""
    zaehler: collections.Counter = collections.Counter()
    for s in saetze:
        for th in themen_in(s):
            if th.art == "prozess":
                zaehler[th.key] += 1
    return [{"key": k, "name": ALLE_THEMEN[k].de, "saetze": n}
            for k, n in zaehler.most_common()]


# ---------------------------------------------------------------------------
# Selbsttest
# ---------------------------------------------------------------------------
def _selbsttest() -> int:
    p20 = profil_aus_name("Station 20 Abisolieren")
    assert p20 and p20.nummer == "20" and "ABISOLIEREN" in p20.prozesse, p20
    assert profil_aus_name("Gesamtanlage (Stationszuordnung folgt)") is None
    liste = [profil_aus_name("Station 10 Zuführung"), p20,
             profil_aus_name("Station 30 Crimpen"),
             profil_aus_name("Station 40 Einpressen")]
    faelle = [
        ("The stripping length must be held within +/-0.2 mm.", "Station"),
        ("Die Abisolierlänge ist mit ±0,2 mm einzuhalten.", "Station"),
        ("The crimp height shall be monitored continuously.", "Andere Station"),
        ("Press-in forces must be monitored.", "Andere Station"),
        ("Station 30 shall be equipped with a crimp force monitor.", "Andere Station"),
        ("Station 20 must reject parts with damaged insulation.", "Station"),
        ("All stations must reach Cgk >= 1.33 (MFU).", "Anlagenweit"),
        ("The machine shall comply with the Machinery Directive.", "Anlagenweit"),
        ("Operating instructions must be delivered in German.", "Anlagenweit"),
        ("The customer reserves the right to change the colour.", "Unklar"),
    ]
    fehler = 0
    for text, soll in faelle:
        ist = bewerte(text, p20, liste)
        ok = ist.geltung == soll
        fehler += not ok
        print(f"{'ok  ' if ok else 'FEHL'} {soll:<15} ist {ist.geltung:<15} {text[:55]}  [{ist.warum}]")
    # Ohne Station wird nur das Thema bestimmt
    ohne = bewerte("Camera inspection of the crimp is required.", None)
    ok = ohne.geltung == "" and "Crimpen" in ohne.themen
    fehler += not ok
    print(f"{'ok  ' if ok else 'FEHL'} ohne Station nur Thema: {ohne.themen}")
    print("\nAlle bestanden." if not fehler else f"\n{fehler} Fehler.")
    return 1 if fehler else 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Gilt dieser Satz fuer meine Station?")
    ap.add_argument("--station")
    ap.add_argument("--stationsliste", type=Path)
    ap.add_argument("--text")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--vorschlag", type=Path, metavar="ZUORDNUNG_JSONL",
                    help="aus requirements.jsonl / zuordnung.jsonl ein Stationsgeruest vorschlagen")
    a = ap.parse_args()
    if a.selftest:
        return _selbsttest()
    if a.vorschlag:
        saetze = []
        for z in a.vorschlag.open(encoding="utf-8"):
            if z.strip():
                d = json.loads(z)
                saetze.append(d.get("text") or d.get("sentence") or "")
        for v in vorschlag(saetze):
            print(f"{v['saetze']:>5}  {v['name']}")
        return 0
    if a.text:
        profil = profil_aus_name(a.station or "")
        alle = stationsliste_laden(a.stationsliste) if a.stationsliste else None
        b = bewerte(a.text, profil, alle)
        print(f"Geltung: {b.geltung or '(keine Station)'}\nGrund:   {b.warum}\nThema:   {b.thema}")
        return 0
    ap.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
