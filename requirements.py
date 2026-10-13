#!/usr/bin/env python3
"""
requirements.py — Anforderungsobjekte aus dem Chunk-Bestand.

Schliesst die zentrale Luecke der Pipeline: heute geht die Kette
Dokument -> Chunk -> Satz -> Anzeige. Fuer die Entwicklungsakte, die
Konfliktpruefung und den Excel-Export braucht es dazwischen ein Objekt
mit Identitaet: die Anforderung.

Kein Sprachmodell. Nur Regeln, regulaere Ausdruecke und Woerterbuecher.
Nur Standardbibliothek.

    python requirements.py                      # alles, schreibt out/requirements.jsonl
    python requirements.py --project A07614000  # nur ein Projekt
    python requirements.py --show 15            # Tabelle auf der Konsole
    python requirements.py --self-test          # ohne Korpus pruefbar

Ausgabefelder je Anforderung
----------------------------
req_id        stabile Kennung dieses Wortlauts
subject_key   Kennung des *Gegenstands* ohne Zahlenwerte.
              Damit findet die Akte spaeter zusammen, was dieselbe
              Anforderung in zwei Fassungen ist:
              "Taktzeit max. 32 s"  und  "Taktzeit max. 28 s"
              haben denselben subject_key, aber verschiedene req_id.
binding       MUSS | SOLL | KANN | VERBOT | INFO
negated       True, wenn der Satz eine Verneinung enthaelt
operator      LE | GE | LT | GT | EQ | RANGE | None
value         Zahl (deutsches Format korrekt gelesen)
value_max     zweite Zahl bei RANGE
unit          mm, s, bar, V, IP.. usw.
tolerance     z. B. "+/- 0,05 mm"
clause        Gliederungsnummer, falls der Satz mit einer beginnt
category      Disziplin bzw. Station (kommt aus router.py)
measurable    True, wenn Operator, Wert und Einheit vorliegen
needs_question offene Stelle beim Kunden ("wird noch festgelegt", "t.b.d.")
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from pathlib import Path

try:  # router liegt daneben; die Datei laeuft auch ohne ihn
    from router import Router
except Exception:  # pragma: no cover
    Router = None  # type: ignore

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"


# ===================================================================== 1 Saetze
# Abkuerzungen, hinter denen der Punkt KEIN Satzende ist. Ohne diese Liste
# zerschneidet der Splitter "max. 32 s" oder "z. B. Schutzart" mitten im Satz.
ABBREV = {
    "z", "b", "bzw", "ca", "evtl", "ggf", "inkl", "exkl", "max", "min", "mind",
    "nr", "abs", "art", "vgl", "sog", "u", "a", "d", "h", "s", "o", "usw",
    "etc", "tel", "fa", "hr", "dr", "ing", "pos", "art", "kap", "abb", "tab",
    "rev", "ver", "zzgl", "lt", "gem", "bspw", "einschl",
}

_SENT_END = re.compile(r"([.!?:;])\s+(?=[A-ZÄÖÜ0-9])")


def split_sentences(text: str) -> list[str]:
    """Deutscher Satz-Splitter, der Abkuerzungen und Dezimalzahlen respektiert."""
    text = re.sub(r"\s+", " ", (text or "").strip())
    if not text:
        return []

    # Punkte schuetzen, die keine Satzenden sind
    protected = text
    # Dezimal- und Tausendertrennung: 1.070,5
    protected = re.sub(r"(\d)\.(\d)", "\\1\x00\\2", protected)
    # Gliederungsnummern: 3.2.1
    protected = re.sub(r"(\d)\.(?=\d)", "\\1\x00", protected)
    # Abkuerzungen
    def _abbr(m: re.Match) -> str:
        word = m.group(1)
        return f"{word}\x00 " if word.lower().strip(" ") in ABBREV else m.group(0)
    protected = re.sub(r"\b([A-Za-zÄÖÜäöüß]{1,7})\.\s", _abbr, protected)

    parts = _SENT_END.split(protected)
    # re.split mit Gruppe liefert [satz, zeichen, satz, zeichen, ...]
    sentences, buf = [], ""
    for i, piece in enumerate(parts):
        if i % 2 == 0:
            buf = piece
        else:
            sentences.append((buf + piece).replace("\x00", "."))
            buf = ""
    if buf:
        sentences.append(buf.replace("\x00", "."))
    return [s.strip() for s in sentences if len(s.strip()) > 12]


# ============================================================ 2 Verbindlichkeit
# WICHTIG: Verneinung wird VOR der KANN-Liste geprueft.
# "darf nicht verwendet werden" ist ein Verbot, keine Option.
# Laeuft es ueber die KANN-Liste, wird aus einem Verbot eine Freiheit —
# das ist der teuerste Fehler, den diese Datei machen koennte.

NEG = re.compile(
    r"\b(nicht|kein|keine|keinen|keiner|keinem|keines|nie|niemals|"
    r"unzul(?:ae|[aä])ssig|untersagt|verboten|ausgeschlossen|ohne)\b", re.I)

# "muß" ist die Schreibweise vor 1996 und steht in vielen aelteren
# Lastenheften. Ohne sie faellt jede Pflicht in solchen Dokumenten durch.
MUSS = re.compile(
    r"\b(mu(?:ss|ß)|m[uü](?:ss|ß)en|muessen|zwingend\w{0,3}|"
    r"verbindlich\w{0,4}|erforderlich\w{0,4}|vorgeschrieben\w{0,4}|"
    r"zu gew[aä]hrleisten|einzuhalten|sicherzustellen|verpflichtend\w{0,3}|"
    r"wird gefordert|gefordert wird|ben[oö]tig\w*|benoetig\w*)\b", re.I)

SOLL = re.compile(r"\b(soll|sollen|sollte|sollten|angestrebt|vorzugsweise)\b", re.I)

KANN = re.compile(
    r"\b(kann|k[oö]nnen|darf|d[uü]rfen|optional|wahlweise|nach Bedarf|"
    r"bei Bedarf|freigestellt|Option\b)\b", re.I)

# --- Welche Sprache hat der Satz? -----------------------------------------
# "can" und "may" sind englische Signalwoerter — aber "CAN" ist zugleich der
# Name des Bussystems. Beides laesst sich nur unterscheiden, wenn man den
# ganzen Satz anschaut: Steht der Satz in englischen Funktionswoertern, ist
# "can" ein Modalverb. Steht er in deutschen, ist es das Bussystem.
STOP_DE = {"der", "die", "das", "den", "dem", "des", "ein", "eine", "einer",
           "und", "oder", "ist", "sind", "wird", "werden", "nicht", "kein",
           "fuer", "für", "mit", "von", "vom", "auf", "im", "in", "am", "an",
           "zu", "zur", "zum", "bei", "als", "aus", "durch", "nach", "ueber",
           "über", "muss", "müssen", "soll", "sollen", "kann", "können",
           "darf", "dürfen", "sich", "auch", "wenn", "dass", "diese", "dieser"}
STOP_EN = {"the", "and", "or", "is", "are", "was", "were", "be", "been",
           "shall", "must", "should", "may", "can", "of", "to", "for", "with",
           "by", "in", "on", "at", "this", "that", "these", "not", "all",
           "any", "such", "which", "from", "has", "have", "will", "it",
           # Runde B: kurze englische Saetze ("We require a written
           # confirmation.") kamen auf 0:0 und wurden als Deutsch gelesen.
           # Nur Woerter, die es im Deutschen nicht gibt — "an", "also",
           # "war", "die" waeren hier fatal.
           "a", "we", "our", "us", "you", "your", "they", "their", "its",
           "if", "when", "as", "but", "than", "there", "those", "into",
           "between", "within", "during", "after", "before", "other", "only"}

_WORDS = re.compile(r"[a-zA-ZäöüÄÖÜß]+")


def detect_lang(sentence: str) -> str:
    """de | en — anhand der Funktionswoerter, nicht anhand einzelner Begriffe."""
    words = [w.lower() for w in _WORDS.findall(sentence)]
    de = sum(w in STOP_DE for w in words)
    en = sum(w in STOP_EN for w in words)
    if en > de:
        return "en"
    return "de"          # im Zweifel Deutsch — der Bestand ist ueberwiegend deutsch


# Englische Signalwoerter. Sie greifen NUR, wenn der Satz als englisch
# erkannt wurde. Dadurch bleibt "CAN-Bus" im deutschen Satz unberuehrt und
# "the supplier can provide" im englischen Satz wird trotzdem erkannt.
MUSS_EN = re.compile(r"\b(shall|must|is required to|are required to|"
                     r"has to|have to|is to be|are to be|mandatory)\b", re.I)
SOLL_EN = re.compile(r"\b(should|is recommended|are recommended|preferably)\b", re.I)
KANN_EN = re.compile(r"\b(may|can|option|optional|optionally|at the discretion)\b", re.I)
VERBOT_EN = re.compile(r"\b(shall not|must not|may not|is not permitted|"
                       r"are not permitted|is prohibited|are prohibited|"
                       r"is not allowed|are not allowed)\b", re.I)

# Das deutsche Gerundivum: "ist in Schutzart IP54 auszufuehren",
# "ist nach DIN EN 60204 auszufuehren", "sind vorzusehen", "ist einzuhalten".
# In Lastenheften die haeufigste MUSS-Formulierung ueberhaupt — und sie wird
# von der Wortliste oben nicht erfasst, weil zwischen "ist" und dem
# zu-Infinitiv beliebig viele Woerter stehen koennen.
# Ein echter zu-Infinitiv steht am Satz- oder Teilsatzende: "ist zu liefern."
# Ohne diese Bedingung trifft das Muster auch den Dativ Plural —
# "ist zu großen Teilen vormontiert" waere sonst eine Pflicht.
# Der Abstand ist bewusst gross: in Lastenheften stehen zwischen "ist" und
# dem Infinitiv oft zehn Woerter ("ist in deutscher Sprache in zweifacher
# Ausfertigung spaetestens vier Wochen vor der Abnahme zu uebergeben").
_INFINITIV = (r"(?:zu\s+[a-zäöüß]{3,}n|[a-zäöüß]{2,}zu[a-zäöüß]{2,}n)"
              r"(?=\s*[.,;:)]|\s+(?:und|oder|bzw|sowie|jedoch|sofern)\b|\s*$)")

GERUND = re.compile(
    r"\b(?:ist|sind|war|waren|hat|haben|hatte|hatten)\s+(?:\S+\s+){0,15}?" + _INFINITIV)

GERUND_VERBOT = re.compile(
    r"\b(?:ist|sind)\s+nicht\s+(?:\S+\s+){0,12}?" + _INFINITIV)

VERBOT_DIREKT = re.compile(
    r"(darf\s+(?:\w+\s+){0,3}?nicht|d(?:ue|[uü])rfen\s+(?:\w+\s+){0,3}?nicht|"
    r"ist\s+(?:\w+\s+){0,2}?unzul(?:ae|[aä])ssig|sind\s+(?:\w+\s+){0,2}?unzul(?:ae|[aä])ssig|"
    r"ist untersagt|sind untersagt|ist verboten|nicht zul(?:ae|[aä])ssig|"
    r"nicht gestattet|nicht erlaubt|shall not|must not|darf kein)", re.I)


# --- Runde A (22.09.2026): Regeln aus der Goldstandard-Messung -------------
# Jede der folgenden Regeln geht auf einen gemessenen Fehler zurueck.
#
# ENGLISCH "can": In Lastenheften beschreibt "can" fast immer eine Moeglichkeit
# oder eine Eigenschaft ("This can lead to a longer cycle time", "materials that
# can be installed") — keine Freiheit fuer den Lieferanten. In der Stichprobe
# waren ALLE englischen KANN-Einstufungen falsch. "can" zaehlt deshalb nur noch
# als KANN, wenn der Lieferant Subjekt ist ("The supplier can ...").
# Verneintes "can" ("cannot be determined") ist eine Feststellung -> INFO, nicht
# VERBOT: echte englische Verbote kommen mit "shall not / must not / may not".
CAN_NEG_EN = re.compile(r"\b(?:can\s*not|cannot|can't|can’t|could\s*not|couldn't)\b"
                        r"|\bno\s+\w+\s+(?:can|could)\b", re.I)
REL_CAN_EN = re.compile(r"\b(?:which|that|who|what)\s+(?:\w+\s+)?(?:can|could|may|might)\b", re.I)
POSSIB_EN = re.compile(r"\b(?:can|could|may|might)\s+(?:also\s+|possibly\s+)?(?:lead|cause|result|"
                       r"occur|happen|affect|arise|damage|reduce|increase|be caused|be affected)\b", re.I)
SUPPLIER_EN = re.compile(r"\b(?:supplier|contractor|manufacturer|vendor|bidder|seller)\b"
                         r"[^.;]{0,40}?\b(?:can|may)\b", re.I)
MAY_EN = re.compile(r"\bmay\b", re.I)
KANN_WORT_EN = re.compile(r"\b(?:optional|optionally|at the discretion|at the option of)\b", re.I)
# Pflicht ohne Modalverb: "Deviations require written approval", Imperativ
# "For each compartment, provide ..."
MUSS_EN_2 = re.compile(r"\b(?:requires?|(?:is|are) required|needs? to|it is necessary)\b", re.I)
BEDINGUNG_EN = re.compile(r"\b(?:if|in case(?: that)?|when|where|as|unless)\b[^,;.]{0,80}?"
                          r"\b(?:is |are )?required\b", re.I)
IMPERATIV_EN = re.compile(
    r"^\s*(?:\d+(?:\.\d+)*\s+)?(?:(?:for|in|at|on|during|after|before)\s[^,]{1,60},\s*)?"
    # nur Verben, die am Satzanfang kaum als Substantiv vorkommen ("Use of ...",
    # "Supply voltage ...", "Design review ..." waeren sonst Imperative)
    r"(?:provide|ensure|install|submit|observe|comply|verify|equip|label|deliver|"
    r"include|maintain|apply)\b(?!\s+of\b)", re.I)
SOLL_EN_2 = re.compile(r"\b(?:(?:it\s+)?(?:is|are)\s+expected|desirable|is desired|preferred|"
                       r"to favou?r|recommended)\b", re.I)

# DEUTSCH "kann": Moeglichkeit ("kann zu Fehlern fuehren") und Verweis ("kann dem
# Anhang entnommen werden") sind keine Freiheit fuer den Lieferanten.
MOEGLICH_DE = re.compile(r"\b(?:kann|k[oö]nnen|k[oö]nnte|k[oö]nnten)\b[^;]{0,80}?\b(?:f(?:ue|[uü])hren|"
                         r"verursachen|auftreten|entstehen|kommen|beeintr(?:ae|[aä])chtigen|beeinflussen)\b", re.I)
VERWEIS_DE = re.compile(r"\b(?:kann|k[oö]nnen)\b[^;]{0,120}?\b(?:entnommen|eingesehen|nachgelesen|"
                        r"gefunden|abgerufen)\s+werden\b", re.I)
# Eine Verneinung macht aus "kann" nur dann ein Verbot, wenn sie beim Modalverb
# steht ("kann nicht", "darf keine"). Frueher genuegte ein "kein" irgendwo im
# Satz — "(kein Umfang dieses Angebots)" wurde so zum VERBOT.
KANN_NEG_NAH = re.compile(
    # hoechstens zwei Woerter Abstand und nie ueber ein Komma hinweg:
    # "kann optional vorgesehen werden, jedoch nicht an Station 3" bleibt KANN
    r"\b(?:kann|k[oö]nnen|darf|d[uü]rfen)\s+(?:[^\s,;]+\s+){0,2}?(?:nicht|kein\w*|nie|niemals)\b"
    r"|\b(?:nicht|kein\w*|nie|niemals)\s+(?:[^\s,;]+\s+){0,2}?(?:kann|k[oö]nnen|darf|d[uü]rfen)\b", re.I)
# Abstimmungspflicht: "IP-Adressen werden in Abstimmung mit dem AG vergeben"
KOORD_DE = re.compile(r"\b(?:in Abstimmung mit|nach Absprache mit|mit Zustimmung|nach Freigabe durch|"
                      r"im Einvernehmen mit)\s+(?:dem|der|des|den)?\s*(?:AG|Auftraggeber\w*|Kunden?)\b", re.I)
FINIT_DE = re.compile(r"\b(?:wird|werden|ist|sind|erfolgt|erfolgen)\b", re.I)
# Lieferunterlagen aus der Gliederung: "10.2.10.5 Weitere Dokumentationsunterlagen
# Sicherheitsdatenblaetter fuer alle Gefahrstoffe ..." — in Liefervorschriften ist
# ein nummerierter Punkt zu Unterlagen eine Lieferpflicht, auch ohne Modalverb.
LIEFERPUNKT_DE = re.compile(r"^\s*(?:\d{1,2}\.){1,5}\d{0,2}\s+(?:\S+\s+){0,3}?\S*(?:unterlagen|"
                            r"dokumentation|nachweise|dokumente|lieferumfang)\b", re.I)


# --- Runde B (22.09.2026): Regeln aus der Pruefstichprobe -------------------
# Gemessen wurde: Sauberkeit 88,6 %, Trefferquote 38,5 %. Alle 8 uebersehenen
# Anforderungen waren "kein Signal" — Pflichten OHNE Modalverb. Runde B zielt
# ausschliesslich auf diese Klasse.

# 1 · Verweise auf mitgeltende Unterlagen, jetzt auch englisch
VERWEIS_EN = re.compile(
    r"\b(?:can|could|may|are|is|shall)\s+(?:\w+\s+){0,3}?(?:be\s+)?"
    r"(?:found|described|specified|defined|listed|stated)\s+in\b"
    r"|\bas\s+(?:specified|described|defined)\s+in\b"
    r"|\brefer\s+to\b|\bsee\s+(?:chapter|section|annex|appendix|document|table)\b", re.I)

# 2 · aktive Erwartung und Forderung ("We expect our partners to ...")
WE_EXPECT_EN = re.compile(r"\bwe\s+(?:expect|would\s+like|prefer)\b", re.I)
WE_REQUIRE_EN = re.compile(r"\bwe\s+(?:require|demand)\b", re.I)

# 3 · Einschraenkung: "darf nur", "may only" — eine Erlaubnis MIT Bedingung ist
#     in Wahrheit eine Pflicht. Wurde bisher als Freiheit (KANN) gelesen.
EINSCHRAENKUNG_DE = re.compile(
    r"\b(?:darf|d(?:ue|[uü])rfen|kann|k(?:oe|[oö])nnen)\s+(?:[^\s,;]+\s+){0,3}?nur\b", re.I)
EINSCHRAENKUNG_EN = re.compile(r"\b(?:may|can|shall|must)\s+only\b|\bonly\s+\w+\s+(?:may|can)\b", re.I)

# 4 · Indikativ als Vorgabe im Deutschen. Bewusst eng gefasst — nur drei
#     Formen, die in Lastenheften eindeutig normativ gemeint sind.
NUR_NACH_DE = re.compile(
    r"\b(?:erfolgt|erfolgen|beginnt|beginnen|startet|ist|sind)\b[^.;]{0,60}?"
    r"\b(?:nur|erst)\s+(?:nach|mit|bei|wenn)\b", re.I)
VORBEHALT_DE = re.compile(r"\bbeh(?:ae|[aä])lt\s+sich\b[^.;]{0,40}?\bvor\b"
                          r"|\bvorbehalten\b", re.I)
BEISTELLUNG_DE = re.compile(
    r"\b(?:wird|werden)\b[^.;]{0,80}?\b(?:beigestellt|zur\s+Verf(?:ue|[uü])gung\s+gestellt"
    r"|bauseits\s+gestellt|vom\s+AG\s+geliefert|durch\s+den\s+AG\s+geliefert)\b", re.I)


def _entscheide(sentence: str) -> tuple[str, bool, str, str]:
    """
    Die EINZIGE Stelle, an der die Verbindlichkeit entschieden wird.
    Gibt (Verbindlichkeit, verneint, Sprache, Regel) zurueck.
    classify_binding() und binding_rule() lesen beide hier ab — dadurch kann
    die Erklaerung nie von der tatsaechlichen Entscheidung abweichen.
    """
    lang = detect_lang(sentence)
    negated = bool(NEG.search(sentence))

    def hit(rx):
        m = rx.search(sentence)
        return m.group(0) if m else ""

    # ---------------------------------------------------------- englisch
    if lang == "en":
        if hit(VERBOT_EN):
            return "VERBOT", True, lang, f'englisches Verbotswort: "{hit(VERBOT_EN)}"'
        if hit(MUSS_EN):
            return "MUSS", negated, lang, f'englisches Pflichtwort: "{hit(MUSS_EN)}"'
        # "if required / in case that ... is required" ist eine Bedingung, keine Pflicht
        ohne_bedingung = BEDINGUNG_EN.sub(" ", sentence)
        m2 = MUSS_EN_2.search(ohne_bedingung)
        if m2:
            return "MUSS", negated, lang, f'englische Pflicht ohne Modalverb: "{m2.group(0)}"'
        if hit(WE_REQUIRE_EN):
            return "MUSS", negated, lang, f'aktive Forderung: "{hit(WE_REQUIRE_EN)}"'
        if hit(EINSCHRAENKUNG_EN):
            return "MUSS", negated, lang, f'Einschraenkung: "{hit(EINSCHRAENKUNG_EN)}"'
        if hit(IMPERATIV_EN):
            return "MUSS", negated, lang, f'englischer Imperativ: "{hit(IMPERATIV_EN).strip()}"'
        if hit(SOLL_EN) or hit(SOLL_EN_2) or hit(WE_EXPECT_EN):
            w = hit(SOLL_EN) or hit(SOLL_EN_2) or hit(WE_EXPECT_EN)
            return "SOLL", negated, lang, f'englisches Sollwort: "{w}"'
        if hit(VERWEIS_EN):
            return "INFO", negated, lang, f'Verweis auf Unterlage: "{hit(VERWEIS_EN)[:60]}"'
        if hit(CAN_NEG_EN):
            return "INFO", True, lang, f'verneintes "can" = Feststellung: "{hit(CAN_NEG_EN)}"'
        if hit(POSSIB_EN):
            return "INFO", negated, lang, f'Moeglichkeit, keine Erlaubnis: "{hit(POSSIB_EN)}"'
        if hit(SUPPLIER_EN):
            return "KANN", False, lang, f'Lieferant darf: "{hit(SUPPLIER_EN)}"'
        if hit(REL_CAN_EN):
            return "INFO", negated, lang, f'Relativsatz beschreibt Eigenschaft: "{hit(REL_CAN_EN)}"'
        if hit(KANN_WORT_EN):
            return "KANN", False, lang, f'englisches Kannwort: "{hit(KANN_WORT_EN)}"'
        if hit(MAY_EN):
            return "KANN", False, lang, 'englisches Kannwort: "may"'
        if re.search(r"\bcan\b", sentence, re.I):
            return "INFO", negated, lang, '"can" ohne Lieferantenbezug = Beschreibung'
        if spez_vorgabe(sentence):
            return "MUSS", negated, lang, "Spezifikationszeile ohne Verb (Tabellenzeile)"
        return "INFO", negated, lang, "kein englisches Signalwort gefunden"

    # ---------------------------------------------------------- deutsch
    # 1. Verbot hat immer Vorrang
    if hit(VERBOT_DIREKT):
        return "VERBOT", True, lang, f'Verbot: "{hit(VERBOT_DIREKT)}"'
    if hit(GERUND_VERBOT):
        return "VERBOT", True, lang, f'verneintes Gerundivum: "{hit(GERUND_VERBOT)}"'
    # 1b. Gerundivum ist eine MUSS-Formulierung
    if hit(GERUND):
        return "MUSS", negated, lang, f'Gerundivum (ist ... zu ...): "{hit(GERUND)}"'
    # 2. Verneintes MUSS bleibt ein Verbot ("es muss vermieden werden")
    if hit(MUSS):
        if negated and re.search(r"\b(vermieden|unterbleiben|verhindert|ausgeschlossen)\b",
                                 sentence, re.I):
            return "VERBOT", True, lang, f'Pflichtwort mit Vermeidung -> Verbot: "{hit(MUSS)}"'
        return "MUSS", negated, lang, f'Pflichtwort: "{hit(MUSS)}"'
    if hit(SOLL):
        return "SOLL", negated, lang, f'Sollwort: "{hit(SOLL)}"'
    # 2a. Einschraenkung: "darf nur im Einrichtbetrieb" ist keine Freiheit.
    #     Vorrang hat aber die Moeglichkeitsaussage und der Verweis — genau
    #     dieselbe Reihenfolge wie unten im KANN-Zweig. Ohne diese Klammer
    #     wurde "... kann zu Fehlern fuehren, dies kann nur durch eine
    #     Aenderung verhindert werden" als Pflicht gelesen (Risikohinweis).
    if hit(EINSCHRAENKUNG_DE) and not (hit(MOEGLICH_DE) or hit(VERWEIS_DE)):
        return "MUSS", negated, lang, f'Einschraenkung: "{hit(EINSCHRAENKUNG_DE)}"'
    # 2b. Abstimmungspflicht mit dem Auftraggeber
    if hit(KOORD_DE) and hit(FINIT_DE):
        return "MUSS", negated, lang, f'Abstimmungspflicht: "{hit(KOORD_DE)}"'
    # 3. "kann" — erst Moeglichkeit und Verweis aussortieren
    if hit(KANN):
        if hit(VERWEIS_DE):
            return "INFO", negated, lang, f'Verweis auf Unterlage: "{hit(VERWEIS_DE)[:60]}"'
        if hit(MOEGLICH_DE):
            return "INFO", negated, lang, f'Moeglichkeit, keine Erlaubnis: "{hit(MOEGLICH_DE)[:60]}"'
        if hit(KANN_NEG_NAH):
            return "VERBOT", True, lang, f'Kannwort mit Verneinung -> Verbot: "{hit(KANN_NEG_NAH)}"'
        return "KANN", False, lang, f'Kannwort: "{hit(KANN)}"'
    # 3b. Indikativ als Vorgabe: "Die Inbetriebnahme erfolgt nur nach Freigabe"
    if hit(NUR_NACH_DE):
        return "MUSS", negated, lang, f'Bedingung als Vorgabe: "{hit(NUR_NACH_DE)[:50]}"'
    if hit(VORBEHALT_DE):
        return "INFO", negated, lang, f'Vorbehalt des Auftraggebers: "{hit(VORBEHALT_DE)[:40]}"'
    if hit(BEISTELLUNG_DE):
        return "INFO", negated, lang, f'Beistellung durch den Auftraggeber: "{hit(BEISTELLUNG_DE)[:50]}"'
    # 4. nummerierter Punkt zu Lieferunterlagen
    if hit(LIEFERPUNKT_DE):
        return "MUSS", negated, lang, f'Gliederungspunkt zu Lieferunterlagen: "{hit(LIEFERPUNKT_DE).strip()}"'
    if spez_vorgabe(sentence):
        return "MUSS", negated, lang, "Spezifikationszeile ohne Verb (Tabellenzeile)"
    return "INFO", negated, lang, "kein Signalwort gefunden"


def classify_binding(sentence: str) -> tuple[str, bool, str]:
    """Gibt (Verbindlichkeit, verneint, Sprache) zurueck."""
    b, neg, lang, _ = _entscheide(sentence)
    return b, neg, lang


# --- Welche Regel hat entschieden? ---------------------------------------
# Liest dieselbe Entscheidung ab wie classify_binding — kann also nicht abweichen.
def binding_rule(sentence: str) -> str:
    return _entscheide(sentence)[3]


def explain(sentence: str, router=None, header: str = "") -> str:
    """Die vollstaendige Entscheidungskette fuer einen einzigen Satz."""
    L = []
    L.append("SATZ")
    L.append(f"  {sentence}")

    bad, why = is_noise(sentence)
    L.append("\n1 · WIRD DER SATZ ÜBERHAUPT GELESEN?")
    L.append(f"  {'NEIN — verworfen: ' + why if bad else 'ja, sieht wie ein Satz aus'}")
    if bad:
        L.append("\n  -> Das Programm sieht diesen Satz nicht. Ist das falsch,")
        L.append("     sag Bescheid: dann ist ein Filter zu scharf eingestellt.")
        return "\n".join(L)

    lang = detect_lang(sentence)
    words = [w.lower() for w in _WORDS.findall(sentence)]
    de = sum(w in STOP_DE for w in words)
    en = sum(w in STOP_EN for w in words)
    L.append("\n2 · WELCHE SPRACHE?")
    L.append(f"  {'Deutsch' if lang == 'de' else 'Englisch'}  "
             f"(deutsche Funktionswörter: {de} · englische: {en})")
    L.append("  Davon hängt ab, ob \"can\" ein Modalverb ist oder der CAN-Bus.")

    binding, negated, _ = classify_binding(sentence)
    L.append("\n3 · VERBINDLICHKEIT")
    L.append(f"  {binding}      (verneint: {'ja' if negated else 'nein'})")
    L.append(f"  Regel: {binding_rule(sentence)}")

    m = extract_measure(sentence)
    L.append("\n4 · ZAHLENWERT")
    if m["value"] is None:
        L.append("  kein Zahlenwert mit Einheit gefunden")
    else:
        L.append(f'  Operator: {m["operator"] or "keiner"}')
        L.append(f'  Wert:     {m["value"]:g}' +
                 (f' bis {m["value_max"]:g}' if m["value_max"] is not None else ""))
        L.append(f'  Einheit:  {m["unit"]}')
        if m["tolerance"]:
            L.append(f'  Toleranz: {m["tolerance"]}')
        L.append(f'  messbar:  {"ja" if m["operator"] and m["unit"] else "nein — Operator oder Einheit fehlt"}')

    L.append("\n5 · KATEGORIE")
    if router is None:
        L.append("  (kein Router geladen — routing_config.json fehlt?)")
    else:
        L.append(f"  {router.explain(sentence, header)}")

    L.append("\n6 · KENNUNGEN")
    L.append(f'  req_id      {_key(sentence)}   (dieser Wortlaut)')
    L.append(f'  subject_key {subject_key(sentence)}   (der Gegenstand ohne Zahlen)')
    L.append("  Zwei Sätze mit gleichem subject_key sind dieselbe Anforderung")
    L.append("  in zwei Fassungen — darauf baut die Entwicklungsakte auf.")

    sus = suspect_reason(sentence, binding, m)
    L.append("\n7 · VERDACHT")
    L.append(f"  {sus if sus else 'kein Verdacht'}")

    L.append("\n8 · ERGEBNIS")
    L.append(f"  {'wird aufgenommen' if is_requirement(sentence, binding) else 'wird NICHT aufgenommen (kein Signalwort, kein Wert, kein Verdacht)'}")
    return "\n".join(L)


# --- Was passiert bei einem Tippfehler im Signalwort? ---------------------
# Antwort: nichts Sichtbares — der Satz faellt still auf INFO zurueck und die
# Pflicht geht verloren. Das ist die gefaehrlichste Fehlerart, weil sie sich
# nicht meldet. Statt zu raten, welche Tippfehler es geben koennte, wird der
# Verdacht sichtbar gemacht: Saetze, die wie eine Anforderung aussehen, aber
# kein Signalwort haben, werden markiert. Ein Mensch entscheidet.
# Woerter, die dem Aehnlichkeitsvergleich aehneln, aber voellig korrektes
# Deutsch sind. Ohne diese Liste meldet die Verdachtssuche staendig
# "kennen ähnelt können" — und wird dadurch unbrauchbar.
KEIN_TIPPFEHLER = {
    "kennen", "kennt", "kannte", "kannten", "beinhalten", "beinhaltet",
    "erhalten", "enthalten", "enthaelt", "verhalten", "unterhalten",
    "gelten", "halten", "haelt", "koennen", "konnte", "konnten",
    "wollen", "sollen", "sollten", "dienen", "denen", "innen",
}

SIGNAL_WORDS = ("muss", "muß", "müssen", "soll", "sollen", "sollte", "kann",
                "können", "darf", "dürfen", "erforderlich", "zwingend",
                "vorgeschrieben", "verbindlich", "einzuhalten", "optional")


def suspect_reason(sentence: str, binding: str, measure: dict) -> str:
    """Leerer String heisst: kein Verdacht."""
    if binding != "INFO":
        return ""
    # Runde A: Die Signalwortliste ist deutsch. Im englischen Satz meldete der
    # Vergleich "option ähnelt optional" — ein Fehlalarm, kein Tippfehler.
    # Und englische Spezifikationen verwenden verlaesslich shall/must: ein
    # nummerierter englischer Satz ohne Signalwort ist Beschreibung.
    deutsch = detect_lang(sentence) == "de"
    import difflib
    for tok in (re.findall(r"[A-Za-zÄÖÜäöüß]{3,13}", sentence) if deutsch else []):
        low = tok.lower()
        if low in KEIN_TIPPFEHLER:
            continue
        near = difflib.get_close_matches(low, SIGNAL_WORDS, n=1, cutoff=0.85)
        # Nur melden, wenn sich die Woerter auch am Anfang gleichen — sonst
        # meldet der Vergleich zufaellige Aehnlichkeiten in der Wortmitte.
        if near and low != near[0] and low[:2] == near[0][:2]:
            return f'moeglicher Tippfehler: "{tok}" ähnelt "{near[0]}"'
    if measure.get("value") is not None and measure.get("unit"):
        if not satzfoermig(sentence):
            return "Spezifikationszeile ohne Verb — Verbindlichkeit pruefen"
        return "Zahlenwert mit Einheit, aber kein Signalwort"
    if deutsch and CLAUSE.match(sentence):
        return "Gliederungsnummer, aber kein Signalwort"
    return ""


# ================================================================== 3 Zahlen
# Deutsches Format: 1.070 sind tausendsiebzig, 0,05 sind null Komma null fuenf.
# Wird das amerikanisch gelesen, liegen Werte um Faktor 1000 daneben.

UNITS = (
    # Zeitraeume zuerst und lange Formen vor kurzen: "15 years", "10 Jahre",
    # "8 Wochen". Fehlten bis Runde A — Gewaehrleistung, Ersatzteilverfuegbarkeit
    # und Lieferfristen fielen dadurch komplett durch.
    r"Jahren|Jahre|Jahr|Monaten|Monate|Monat|Wochen|Woche|Tagen|Tage|Tag|"
    r"Stunden|Stunde|Minuten|years|year|months|month|weeks|week|days|day|hours|hour|"
    r"minutes|"
    r"mm²|cm²|m²|mm³|m³|"
    r"µm|um|mm|cm|dm|m\b|km|"
    r"mg|g\b|kg|t\b|"
    r"ms|sec|sek|s\b|min|h\b|Std|"
    r"°C|°|K\b|"
    r"mbar|bar|Pa|kPa|MPa|N/mm²|Nm|N\b|"
    r"mV|V\b|kV|mA|A\b|kW|MW|W\b|VA|Hz|kHz|"
    r"dB\(A\)|dB|"
    r"ml|l/min|l\b|m/s|mm/s|U/min|min-1|"
    r"%|Stk|St\.|Zoll|IP\d{2}|IP\s?\d{2}"
)
# Schweizer Schreibweise 1'070 und typografischer Apostroph 1’070
_NUM = r"\d{1,3}(?:[.\'’]\d{3})+(?:,\d+)?|\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?|\d+(?:\.\d+)?"

# Achtung "min": das ist zugleich eine Einheit (Minuten, l/min, U/min).
# Deshalb hier nur "min." mit Punkt oder "mind"/"mindestens" — und nie
# direkt hinter einem Schraegstrich. Sonst wird aus "200 l/min" ein
# "mindestens 200" und die Anforderung kippt in ihr Gegenteil.
OP_LE = re.compile(r"(?<![/\w])((?<!\d)(?<!\d )max\.|maximal|h[oö]chstens|nicht mehr als|bis zu|begrenzt|beschr[aä]nkt|nicht gr[oö][ßs]er als|kleiner gleich|up\s+to(?!\s+date)|no more than|not more than|at most|≤|<=)", re.I)
OP_GE = re.compile(r"(?<![/\w])((?<!\d)(?<!\d )min\.|(?<!\d)(?<!\d )mind\.|mindestens|wenigstens|nicht weniger als|gr[oö]ßer gleich|(?<![/\w])from\s+(?=[\d±+-])|not less than|no less than|at least|≥|>=)", re.I)
# Formulierungen, die den Operator im Verb tragen. Sie werden zuerst geprueft,
# weil sie eine Verneinung enthalten und sonst falsch herum gelesen wuerden.
PHRASE_LE = re.compile(r"nicht\s+(?:zu\s+)?(?:ue|[uü])(?:berschreit|bersteig|berschritten)", re.I)
PHRASE_GE = re.compile(r"nicht\s+(?:zu\s+)?(?:unterschreit|unterschritten|unterbiet)", re.I)
# dasselbe auf Englisch — bisher wurde "should not exceed 30 s" ohne Operator gelesen
PHRASE_LE_EN = re.compile(r"\bnot\s+(?:to\s+)?exceed|\bnot\s+(?:be\s+)?(?:higher|greater|more)\s+than", re.I)
PHRASE_GE_EN = re.compile(r"\bnot\s+(?:to\s+)?fall\s+below|\bnot\s+(?:be\s+)?(?:lower|less)\s+than", re.I)
OP_LT = re.compile(r"(\bkleiner als\s+(?=[\d±+-])|\bunter\s+(?=[\d±+-])|\bweniger als\s+(?=[\d±+-])|<(?!=))", re.I)
OP_GT = re.compile(r"(\bgr[oö][ßs]er als\s+(?=[\d±+-])|\b[uü]ber\s+(?=[\d±+-])|\bmehr als\s+(?=[\d±+-])|>(?!=))", re.I)
OP_EQ = re.compile(r"\b(genau|exakt|betr(?:ae|[aä])g(?:t|en)|entspricht|=)\b", re.I)
NOLETTER = r"(?![A-Za-zÄÖÜäöüß])"
# Vor der Zahl darf kein Buchstabe und keine Ziffer stehen, sonst wird
# aus der Projektnummer A07616000A ein Wert von 7.616.000 Ampere.
NOPREFIX = r"(?<![A-Za-z0-9ÄÖÜäöüß])"
RANGE = re.compile(rf"{NOPREFIX}({_NUM})\s*(?:bis|\.\.\.|–|-|to)\s*({_NUM})\s*({UNITS}){NOLETTER}")
# _NUM enthaelt selbst Alternativen (|). Ohne eigene Klammer bricht die
# Alternation nach aussen aus und das Muster trifft jede beliebige Zahl —
# so wurde aus "IP54" eine Toleranz von 54.
TOL = re.compile(rf"([+±]\s*/?\s*-?\s*(?:{_NUM})\s*(?:{UNITS})?|\+\s*(?:{_NUM})\s*/\s*-\s*(?:{_NUM}))")
VALUE = re.compile(rf"{NOPREFIX}({_NUM})\s*({UNITS}){NOLETTER}")
IPCODE = re.compile(r"\bIP\s?(\d{2})\b")
# Nennweite: "DN 50" — die Zahl steht hinter der Einheit, deshalb ein
# eigenes Muster wie bei IP. Kommt in Rohrleitungs- und Ventilvorgaben
# staendig vor und wurde bisher gar nicht gelesen.
DNCODE = re.compile(r"\bDN\s?(\d{2,4})\b")


def parse_german_number(raw: str) -> float | None:
    """'1.070,5' -> 1070.5 · '0,05' -> 0.05 · '1.070' -> 1070 · '3.2' -> 3.2"""
    s = raw.strip().replace("'", "").replace("\u2019", "")   # Schweizer Trennzeichen
    if not s:
        return None
    if "," in s:                      # Komma ist immer das Dezimaltrennzeichen
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", s):   # 1.070 / 12.500.000
        s = s.replace(".", "")
    # alles andere (3.2) bleibt, wie es ist
    try:
        return float(s)
    except ValueError:
        return None


def extract_measure(sentence: str) -> dict:
    """Operator, Wert, Einheit, Toleranz aus einem Satz."""
    out = {"operator": None, "value": None, "value_max": None,
           "unit": None, "tolerance": None}

    m = TOL.search(sentence)
    if m:
        out["tolerance"] = m.group(1).strip()

    m = RANGE.search(sentence)
    if m:
        out.update(operator="RANGE",
                   value=parse_german_number(m.group(1)),
                   value_max=parse_german_number(m.group(2)),
                   unit=m.group(3).strip())
        return out

    m = IPCODE.search(sentence)
    if m:
        out.update(value=float(m.group(1)), unit="IP")
        out["operator"] = detect_operator(sentence) or "EQ"
        return out

    m = DNCODE.search(sentence)
    if m:
        out.update(value=float(m.group(1)), unit="DN")
        out["operator"] = detect_operator(sentence) or "EQ"
        return out

    m = VALUE.search(sentence)
    if m:
        out["value"] = parse_german_number(m.group(1))
        out["unit"] = m.group(2).strip()

    out["operator"] = detect_operator(sentence)
    return out


def detect_operator(sentence: str) -> str | None:
    # Verbformulierungen zuerst — sie tragen eine Verneinung in sich
    if PHRASE_LE_EN.search(sentence):
        return "LE"
    if PHRASE_GE_EN.search(sentence):
        return "GE"
    if PHRASE_LE.search(sentence):
        return "LE"
    if PHRASE_GE.search(sentence):
        return "GE"
    for rx, op in ((OP_LE, "LE"), (OP_GE, "GE"), (OP_LT, "LT"),
                   (OP_GT, "GT"), (OP_EQ, "EQ")):
        if rx.search(sentence):
            return op
    return None


# ============================================================== 4 Sonstiges
CLAUSE = re.compile(r"^\s*((?:\d{1,2}\.){1,4}\d{0,2})\s+")
OPEN_Q = re.compile(
    r"\b(t\.?b\.?d\.?|tbd|noch (?:zu )?(?:kl[aä]ren|festzulegen|festlegen|definieren)|"
    r"wird noch|offen|n\.\s?n\.|nach Absprache|in Abstimmung|klärungsbedarf|"
    r"vom Kunden (?:zu )?(?:liefern|benennen)|k\.\s?A\.|"
    # Runde A: englische offene Punkte — 41 % des Bestands ist englisch
    r"to be (?:defined|clarified|determined|specified|agreed|confirmed)|"
    r"not yet (?:been )?(?:defined|provided|clarified|specified|known|decided|available)|"
    r"has not yet|have not yet|open point|still open|pending)\b", re.I)

# Saetze, die keine Anforderung sind — Kopfzeilen, Verteiler, Fusszeilen
NOISE = re.compile(
    r"^(seite|page|datum|date|von:|an:|cc:|betreff|subject|tel|fax|e-?mail|"
    r"inhaltsverzeichnis|anlage\s*\d|©|copyright|blatt)\b", re.I)

# --- Extraktionshygiene ---------------------------------------------------
# Aus PDF-Dateien kommt nicht nur Fliesstext. Inhaltsverzeichnisse, Fusszeilen,
# Tabellenzeilen und an der Chunk-Grenze abgeschnittene Woerter sehen fuer eine
# Regel wie ein Satz aus. Ohne diese Filter landet Muell in der Liste — und
# genau daran erkennt ein Vertriebsingenieur, dass er dem Werkzeug nicht
# trauen kann.
# E-Mails sind der groesste Einzelposten an Fehltreffern: Anreden,
# Grussformeln und Vertraulichkeitshinweise sehen wie Anforderungen aus.
# "This email may contain material that is confidential" wurde zur Option
# erklaert — der Satz enthaelt "may".
MAIL_NOISE = re.compile(
    r"^(hallo|guten (tag|morgen)|sehr geehrte|liebe[rs]?\s|mit freundlichen|"
    r"viele gr[uü]sse|beste gr[uü]sse|freundliche gr[uü]sse|von:|gesendet:|"
    r"betreff:|wichtiger hinweis|disclaimer)\b|"
    r"(this (e-?mail|message) (may |might )?(contain|is intended)|"
    r"if you (are not|have received)|sole use of the intended recipient|"
    r"unbefugte weitergabe|nicht der richtige empf[aä]nger|"
    r"klicken sie nicht auf links)", re.I)

# Schutzvermerk nach DIN ISO 16016 und aehnliche Fussnoten: stehen auf jeder
# Seite jedes Dokuments und wurden als VERBOT in die Liste geschrieben.
SCHUTZVERMERK = re.compile(
    r"(weitergabe sowie vervielf[aä]ltigung|all rights reserved|"
    r"alle rechte\b[^.;]{0,60}\bvorbehalten|"
    r"es ist vertraulich und darf|"
    r"darf ohne (?:vorherige |ausdr[uü]ckliche |schriftliche )*(?:genehmigung|zustimmung)\b"
    r"[^.]{0,60}(?:reproduziert|vervielf[aä]ltigt|ver[oö]ffentlicht|weitergegeben))", re.I)

# Ein Satz, der mitten in der Pflicht abbricht, ist keine Anforderung,
# sondern ein Chunk-Ende: "These parts ... also have to"
ABBRUCH_WORTE = {
    "ist", "sind", "hat", "haben", "wird", "werden", "muss", "müssen", "soll",
    "sollen", "kann", "können", "darf", "dürfen", "zu", "und", "oder", "der",
    "die", "das", "den", "dem", "des", "für", "mit", "von", "bei", "als",
    "has", "have", "must", "shall", "should", "may", "can", "be", "to", "the",
    "of", "and", "or", "for", "with", "is", "are",
    # Runde A: Praepositionen und Artikel am Ende ohne Satzzeichen sind fast
    # immer ein Chunk-Ende ("The number of parts has to be given in")
    "in", "on", "at", "by", "from", "into", "a", "an", "as", "than",
    "auf", "an", "im", "am", "zum", "zur", "nach", "ueber", "über", "unter",
    "gemäß", "gemaess", "durch", "ein", "eine", "einer", "eines",
}
_TAIL = re.compile(r"[A-Za-zÄÖÜäöüß]+")


def bricht_ab(s: str) -> bool:
    """Endet der Satz mitten in der Aussage? 'These parts also have to 🔏'"""
    woerter = _TAIL.findall(s)
    if not woerter:
        return True
    # Steht hinter dem letzten Wort noch ein Satzzeichen, ist der Satz zu Ende
    rest = s[s.rfind(woerter[-1]) + len(woerter[-1]):]
    if any(c in ".!?" for c in rest):
        return False
    return woerter[-1].lower() in ABBRUCH_WORTE


# Aus E-Mails kommen Web-Adressen, Sicherheitslinks und base64-Ketten mit.
# Sie enthalten Prozentzeichen und Ziffern und wurden dadurch als Messwerte
# gelesen — "...4%3D" wurde zu "6 %".
WEBMUELL = re.compile(
    r"https?://|www\.[a-z]|mailto:|&reserved=|&sdata=|%3D|%2F|%7C|%2B|"
    r"(?=[A-Za-z0-9+/]{28,})[A-Za-z0-9+/]*[0-9][A-Za-z0-9+/]*", re.I)
TRENNLINIE = re.compile(r"[-_=]{6,}")

TOC_DOTS = re.compile(r"\.{4,}|…{2,}")            # "Packaging ........ 32"
TOC_START = re.compile(r"^\d+(\.\d+)*\s+\S")      # "6.1 Packaging"
PIPES = re.compile(r"\|")                          # Fuss- und Kopfzeilen
WORD3 = re.compile(r"[A-Za-zÄÖÜäöüß]{3,}")


def is_noise(sentence: str) -> tuple[bool, str]:
    """Gibt (ist_Muell, Grund) zurueck — der Grund ist fuer die Fehlersuche da."""
    s = sentence.strip()
    if NOISE.match(s):
        return True, "Kopf- oder Fusszeile"
    if WEBMUELL.search(s):
        return True, "Web-Adresse, Sicherheitslink oder base64-Kette"
    if TRENNLINIE.search(s):
        return True, "Trennlinie aus E-Mail oder Fusszeile"
    if MAIL_NOISE.search(s):
        return True, "E-Mail-Anrede, Grussformel oder Vertraulichkeitshinweis"
    if SCHUTZVERMERK.search(s):
        return True, "Schutzvermerk (Vertraulichkeit/Urheberrecht)"
    if bricht_ab(s):
        return True, "Satz bricht mitten in der Aussage ab (Chunk-Ende)"
    if TOC_DOTS.search(s):
        return True, "Inhaltsverzeichnis (Punktfuehrung)"
    if len(PIPES.findall(s)) >= 2:
        return True, "Tabellen- oder Fusszeile (Trennstriche)"
    if s[:1].islower():
        return True, "beginnt klein — an der Chunk-Grenze abgeschnitten"
    words = WORD3.findall(s)
    if len(words) < 4:
        return True, "zu wenige echte Woerter"
    letters = sum(c.isalpha() for c in s)
    digits = sum(c.isdigit() for c in s)
    if letters and digits / max(len(s), 1) > 0.30:
        return True, "ueberwiegend Ziffern (Tabelle)"
    caps = sum(c.isupper() for c in s if c.isalpha())
    if letters > 20 and caps / letters > 0.60:
        return True, "ueberwiegend Grossbuchstaben (Ueberschrift)"
    if TOC_START.match(s) and len(words) < 7 and not s.rstrip().endswith((".", "!", "?")):
        return True, "Ueberschriftzeile"
    if s.rstrip().endswith(":") and len(words) < 8:
        return True, "Aufzaehlungskopf"
    return False, ""


def is_requirement(sentence: str, binding: str) -> bool:
    noise, _ = is_noise(sentence)
    if noise:
        return False
    if len(sentence) < 20 or len(sentence) > 600:
        return False
    if binding in ("MUSS", "SOLL", "KANN", "VERBOT"):
        return True
    # INFO nur, wenn ein messbarer Wert drinsteht oder ein Verdacht besteht,
    # dass ein Signalwort verstuemmelt ist — sonst ist es Fliesstext
    # Offene Punkte sind das Wertvollste am ganzen Bestand: daraus wird die
    # Rueckfragenliste an den Kunden. Sie haben oft kein Signalwort
    # ("Die Lackfarbe wird noch festgelegt.") und fielen bisher durch.
    m = extract_measure(sentence)
    grund = suspect_reason(sentence, binding, m)
    # Moeglicher Tippfehler im Signalwort: immer zur Pruefung aufnehmen
    if grund.startswith("moeglicher Tippfehler"):
        return True
    # Runde A: Ohne Signalwort wird ein Satz nur aufgenommen, wenn er ein
    # echter Satz ist — mit Verb und ohne Aufzaehlungszeichen. Vorher genuegte
    # eine Zahl mit Einheit; so kamen Folienfetzen ("Handover Unit 100mm 70mm
    # Feeding Unit ...") und Kalkulationszeilen herein. Das war die Haelfte
    # aller Fehlalarme in der Stichprobe.
    # Parameterzeile ohne Verb — die Ausnahme von der Satzregel
    if spezifikationszeile(sentence, m):
        return True
    # Verweis, Vorbehalt und Beistellung ebenfalls: sie sind kaufmaennisch
    # bindend, tragen aber nicht immer ein Verb aus der Liste unten
    if (VERWEIS_DE.search(sentence) or VERWEIS_EN.search(sentence)
            or VORBEHALT_DE.search(sentence) or BEISTELLUNG_DE.search(sentence)):
        return True
    if not satzfoermig(sentence):
        return False
    if OPEN_Q.search(sentence):
        return True
    # Verweis auf eine mitgeltende Unterlage ("kann dem Grenzmusterkatalog im
    # Anhang entnommen werden") gehoert in die Liste — als INFO, nicht als KANN
    if m["value"] is not None and m["unit"]:
        return True
    return bool(grund)


VERB_HINWEIS = re.compile(
    r"\b(?:ist|sind|war|waren|wird|werden|wurde|wurden|hat|haben|betr(?:ae|[aä])gt|betragen|"
    r"liegt|liegen|erfolgt|erfolgen|gilt|gelten|entspricht|entsprechen|besteht|bestehen|"
    r"umfasst|umfassen|enth(?:ae|[aä])lt|enthalten|verf(?:ue|[uü])gt|arbeitet|arbeiten|"
    r"l(?:ae|[aä])uft|laufen|dient|dienen|erreicht|erreichen|"
    r"is|are|was|were|be|been|has|have|had|will|shall|must|should|amounts|equals|"
    r"contains|includes|operates|runs|consists)\b", re.I)
LISTENZEICHEN = re.compile(r"[•▪►]|\s−\s|\||\?{3,}")   # − ist das Minuszeichen aus PDF-Aufzaehlungen


def spezifikationszeile(sentence: str, m: dict) -> bool:
    """
    Parameterzeile ohne Verb: "Cooling time adjustable between 1 - 15 sec",
    "Ball valve up to DN 50". Das ist die haeufigste Form, in der technische
    Vorgaben in Lastenheften stehen, wenn sie aus einer Tabelle stammen.

    Abgegrenzt gegen Folienfetzen ("Handover Unit 100mm 70mm Feeding Unit"):
    Es muss ein Grenzwert ODER ein Bereich erkannt worden sein — eine blosse
    Zahl mit Einheit genuegt nicht. Dazu kurze Zeile, keine Aufzaehlungszeichen.
    """
    if LISTENZEICHEN.search(sentence) or len(sentence) > 250:
        return False
    if m.get("value") is None or not m.get("unit"):
        return False
    if not (m.get("operator") or m.get("value_max") is not None):
        return False
    return 2 <= len(WORD3.findall(sentence)) <= 25


# Messbericht statt Vorgabe: "Rundlauf 0,05 mm gemessen", "actual value 3.2 bar".
# Eine solche Zeile ist eine Feststellung und bleibt INFO.
MESSBERICHT = re.compile(
    r"\b(?:gemessen|ermittelt|festgestellt|Istwert|Ist-Wert|Messwert|Messung|"
    r"measured|actual\s+value|as\s+built|result)\b", re.I)


def spez_vorgabe(sentence: str) -> bool:
    """
    Verblose Parameterzeile mit Grenzwert -> Vorgabe, also MUSS statt INFO.

    Begruendung aus der Sache, nicht aus der Stichprobe: eine Zeile wie
    "Cooling time adjustable between 1 - 15 sec" stammt aus einer
    Anforderungstabelle. Es gibt keine Lesart, in der ein solcher Grenzwert
    unverbindlich waere. Umgekehrt ist eine Zeile mit Messwort
    ("... gemessen") eine Feststellung und bleibt INFO.

    Nur die Verbindlichkeit wird hier angehoben — aufgenommen wird die Zeile
    ohnehin schon durch spezifikationszeile() in is_requirement().
    """
    if VERB_HINWEIS.search(sentence) or MESSBERICHT.search(sentence):
        return False
    return spezifikationszeile(sentence, extract_measure(sentence))


def satzfoermig(sentence: str) -> bool:
    """Echter Satz mit Verb — keine Aufzaehlung, keine Tabellenzeile, kein Folienfetzen."""
    return bool(VERB_HINWEIS.search(sentence)) and not LISTENZEICHEN.search(sentence)


def _key(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12]


_STRIP_NUM = re.compile(rf"({_NUM})\s*({UNITS})")
_STRIP_WS = re.compile(r"[^a-zäöüß ]+")


def subject_key(sentence: str, project_no: str = "") -> str:
    """
    Kennung des Gegenstands, unabhaengig von den Zahlenwerten.

    Genau das macht die Entwicklungsakte ohne KI moeglich:
    "Die Taktzeit muss max. 32 s betragen."   ->  subject_key X
    "Die Taktzeit muss max. 28 s betragen."   ->  subject_key X  (dieselbe Anforderung)
    Die beiden Saetze bekommen verschiedene req_id, aber denselben subject_key —
    daraus baut store_akte.py die Versionskette.
    """
    s = sentence.lower()
    s = _STRIP_NUM.sub(" ", s)          # Zahlen samt Einheit raus
    s = re.sub(r"\d+", " ", s)          # restliche Ziffern raus
    s = _STRIP_WS.sub(" ", s)           # Satzzeichen raus
    stop = {"der", "die", "das", "den", "dem", "des", "ein", "eine", "einer",
            "eines", "und", "oder", "von", "vom", "für", "fuer", "mit", "bei",
            "ist", "sind", "wird", "werden", "muss", "müssen", "soll", "sollen",
            "kann", "können", "darf", "dürfen", "zu", "im", "in", "am", "an",
            "auf", "als", "nicht", "max", "min", "mindestens", "maximal", "es"}
    words = sorted({w for w in s.split() if len(w) > 3 and w not in stop})
    return _key(project_no + "|" + " ".join(words[:12]))


# ================================================================ 5 Objekt
@dataclass
class Requirement:
    req_id: str
    subject_key: str
    text: str
    binding: str
    negated: bool
    clause: str = ""
    operator: str | None = None
    value: float | None = None
    value_max: float | None = None
    unit: str | None = None
    tolerance: str | None = None
    category: str = ""
    category_hits: list[str] = field(default_factory=list)
    measurable: bool = False
    needs_question: bool = False
    lang: str = "de"
    suspect: str = ""          # nicht leer = bitte von Hand ansehen
    # Herkunft — jede Anforderung bleibt bis zur Seite belegbar
    project_no: str = ""
    project_name: str = ""
    customer: str = ""
    phase: str = ""
    file: str = ""
    source_path: str = ""
    page: str | int | None = None
    chunk_id: str | int | None = None
    extracted_at: str = ""


def from_sentence(sentence: str, meta: dict, router=None) -> Requirement | None:
    binding, negated, lang = classify_binding(sentence)
    if not is_requirement(sentence, binding):
        return None

    clause_m = CLAUSE.match(sentence)
    clause = clause_m.group(1) if clause_m else ""
    body = sentence[clause_m.end():] if clause_m else sentence

    meas = extract_measure(body)
    cat, hits = ("", [])
    if router is not None:
        cat, hits = router.route(body, header=meta.get("header", ""))

    return Requirement(
        req_id=_key(f'{meta.get("project_no","")}|{body}'),
        subject_key=subject_key(body, meta.get("project_no", "")),
        text=body.strip(),
        binding=binding,
        negated=negated,
        clause=clause,
        operator=meas["operator"],
        value=meas["value"],
        value_max=meas["value_max"],
        unit=meas["unit"],
        tolerance=meas["tolerance"],
        category=cat,
        category_hits=hits,
        measurable=bool(meas["operator"] and meas["value"] is not None and meas["unit"]),
        needs_question=bool(OPEN_Q.search(body)),
        lang=lang,
        suspect=suspect_reason(body, binding, meas),
        project_no=meta.get("project_no", ""),
        project_name=meta.get("project_name", ""),
        customer=meta.get("customer", ""),
        phase=meta.get("phase", ""),
        file=meta.get("file", ""),
        source_path=meta.get("source_path", ""),
        page=meta.get("page"),
        chunk_id=meta.get("chunk_id"),
        extracted_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
    )


# ============================================================== 6 Korpuslauf
def load_chunks(path: Path) -> list[dict]:
    rows = []
    with path.open(encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def chunk_meta(c: dict) -> dict:
    phase = ""
    if c.get("phase_no") or c.get("phase_name"):
        phase = f'{c.get("phase_no","")}_{c.get("phase_name","")}'.strip("_")
    return {
        "project_no": c.get("project_no", ""),
        "project_name": c.get("project_name", ""),
        "customer": c.get("customer", ""),
        "phase": phase,
        "file": c.get("filename") or Path(str(c.get("source_path", ""))).name,
        "source_path": c.get("source_path", ""),
        "page": c.get("page"),
        "chunk_id": c.get("id"),
        "header": c.get("header", ""),
    }


def build(chunks_path: Path, project: str | None = None,
          router=None) -> list[Requirement]:
    out: list[Requirement] = []
    seen: set[str] = set()
    for c in load_chunks(chunks_path):
        if project and c.get("project_no") != project:
            continue
        meta = chunk_meta(c)
        for sent in split_sentences(c.get("text", "")):
            r = from_sentence(sent, meta, router)
            # Dubletten: gleicher Wortlaut bis auf Leerzeichen und Gross-
            # schreibung zaehlt INNERHALB EINES DOKUMENTS als dieselbe
            # Anforderung. Chunk-Ueberlappung und wiederholte Textbausteine
            # erzeugen sonst Karteileichen.
            #
            # Runde A: die Projektnummer kam in den Schluessel. Vorher lief
            # "seen" ueber den ganzen Bestand — ein wortgleicher Satz in zwei
            # Projekten wurde nur dem ersten zugeschlagen und fehlte im zweiten.
            #
            # 01.10.2026: die QUELLE gehoert ebenfalls in den Schluessel.
            # Mit Projektnummer allein wurde ein Satz, der in zwei Fassungen
            # desselben Lastenhefts gleich lautet, nur aus der ersten Fassung
            # behalten. Damit war die Versionskette — der eigentliche Zweck
            # von subject_key und req_id — auf echten Bestaenden unbrauchbar:
            # von 247 Anforderungen aus LH_..._V1.01 und 51 aus V1.02 liessen
            # sich nur 6 verknuepfen, und keine davon war wortgleich.
            # Jetzt bleibt jede Anforderung einmal JE DOKUMENT erhalten:
            #   gleiche req_id in zwei Fassungen  -> unveraendert uebernommen
            #   gleicher subject_key, andere req_id -> geaendert
            # Projektnummer bleibt im Schluessel: fehlt der Pfad, wuerden zwei
            # gleichnamige Dateien aus verschiedenen Projekten sonst wieder
            # zusammenfallen — genau der Fehler, den Runde A behoben hat.
            quelle = ((r.project_no or "") + "|" +
                      (r.source_path or r.file or "")) if r else ""
            norm = quelle + "|" + \
                re.sub(r"\W+", " ", (r.text if r else "")).strip().lower()
            id_je_quelle = quelle + "|" + (r.req_id if r else "")
            if r and id_je_quelle not in seen and norm not in seen:
                seen.add(id_je_quelle)
                seen.add(norm)
                out.append(r)
    return out


# ================================================================ 7 Selbsttest
SELF_TEST = [
    # (Satz, Verbindlichkeit, Wert, Einheit, Operator)
    ("Die Taktzeit muss max. 32 s betragen.", "MUSS", 32.0, "s", "LE"),
    ("Die Taktzeit muss max. 28 s betragen.", "MUSS", 28.0, "s", "LE"),
    ("Der Schaltschrank darf nicht ausserhalb der Halle aufgestellt werden.",
     "VERBOT", None, None, None),
    ("Eine Beschichtung kann optional aufgebracht werden.", "KANN", None, None, None),
    ("Der Bauraum ist auf 1.070 mm begrenzt.", "INFO", 1070.0, "mm", "LE"),
    # Sprache: CAN im deutschen Satz ist das Bussystem, nicht "koennen"
    ("Der Sensor ist gegen den CAN-Bus isoliert.", "INFO", None, None, None),
    ("Die Daten werden ueber den CAN-Bus uebertragen.", "INFO", None, None, None),
    # dieselben Woerter im englischen Satz sind sehr wohl Modalverben
    ("The supplier can provide an optional cover.", "KANN", None, None, None),
    ("The cabinet shall be designed for IP54 protection.", "MUSS", 54.0, "IP", "EQ"),
    ("The machine must not be operated without a guard.", "VERBOT", None, None, None),
    # alte Rechtschreibung
    ("Die Taktzeit muß max. 32 s betragen.", "MUSS", 32.0, "s", "LE"),
    ("Die Toleranz betraegt +/- 0,05 mm.", "INFO", 0.05, "mm", "EQ"),
    ("Die Schutzart soll mindestens IP54 erreichen.", "SOLL", 54.0, "IP", "GE"),
    ("Der Druck liegt zwischen 4 bis 6 bar.", "INFO", 4.0, "bar", "RANGE"),
    ("Die Lackfarbe wird noch festgelegt.", "INFO", None, None, None),
    # Die zwei Faelle, an denen sich der Operator leicht verdreht:
    ("Der Luftverbrauch soll 200 l/min nicht ueberschreiten.", "SOLL", 200.0,
     "l/min", "LE"),
    ("Die Haltekraft darf 50 N nicht unterschreiten.", "VERBOT", 50.0, "N", "GE"),
    # Gerundivum — in Lastenheften die haeufigste MUSS-Formulierung:
    ("Der Schaltschrank ist in Schutzart IP54 auszufuehren.", "MUSS", 54.0, "IP", "EQ"),
    ("Die Verkabelung ist nach DIN EN 60204 auszufuehren.", "MUSS", None, None, None),
    ("Eine Rueckfallebene ist vorzusehen.", "MUSS", None, None, None),
    ("Die Anlage ist nicht ohne Freigabe in Betrieb zu nehmen.", "VERBOT", None, None, None),
    # Verben auf -ern/-eln wurden vom Gerundivum uebersehen
    ("Der Zugang in den Gefahrenbereich ist komplett zu sichern.", "MUSS", None, None, None),
    # "zu" als Steigerung darf keine Pflicht erzeugen
    ("Das Ergebnis ist zu Beginn noch unklar.", "INFO", None, None, None),
    ("Der Abstand ist zu gross fuer die Vorrichtung.", "INFO", None, None, None),
    # Zahl aus einer Projektnummer ist kein Messwert
    ("A07616000A Automatische Kabelkonfektionierung der Anlage.", "INFO", None, None, None),
    # Vergleichszeichen und englische Vergleichsformulierungen
    ("Bei einer Anschlussleistung > 5 kW ist eine Erfassung vorzusehen.", "MUSS", 5.0, "kW", "GT"),
    ("Not less than protection type IP54 has to be provided.", "MUSS", 54.0, "IP", "GE"),
    # "cannot" ist beschreibend, kein Verbot
    ("Otherwise new machines cannot be integrated into the system.", "INFO", None, None, None),
    # die Kernformulierung jeder Anfrage
    ("Wir benoetigen eine Pruefstation mit Visualisierung nach Richtlinie.", "MUSS", None, None, None),
    # Dativ Plural auf -n ist kein Infinitiv und damit keine Pflicht
    ("Die Anlage ist zu großen Teilen vormontiert.", "INFO", None, None, None),
    ("Der Preis ist zu diesen Konditionen unveraendert.", "INFO", None, None, None),
    # ... echte Gerundive muessen weiter greifen, auch mit vielen Woertern dazwischen
    ("Die Dokumentation ist in deutscher Sprache in zweifacher Ausfertigung "
     "spaetestens vier Wochen vor der Abnahme zu uebergeben.", "MUSS", None, None, None),
    ("Die Teile sind zu liefern und zu montieren.", "MUSS", None, None, None),
    # "ueber"/"unter" sind Praepositionen, kein Vergleich
    ("Die Ansteuerung erfolgt über eine SPS mit 24 V.", "INFO", 24.0, "V", None),
    ("Der Druck liegt über 6 bar.", "INFO", 6.0, "bar", "GT"),
    # "min." ist Operator oder Einheit — die Zahl davor entscheidet
    ("Die Laufzeit liegt unter 5 min.", "INFO", 5.0, "min", "LT"),
    ("Der Druck muss min. 5 bar betragen.", "MUSS", 5.0, "bar", "GE"),
    # "hat ... zu ..." — die haeufigste Pflichtformel in Liefervorschriften
    ("Der Auftragnehmer hat die Anlage vollstaendig zu liefern.", "MUSS", None, None, None),
    ("Der AN hat saemtliche Unterlagen beizustellen.", "MUSS", None, None, None),
    ("Der Lieferant hat zu großen Teilen beigetragen.", "INFO", None, None, None),
    # Schweizer Tausendertrennung
    ("Der Bauraum ist auf 1'070 mm begrenzt.", "INFO", 1070.0, "mm", "LE"),
    # --- Runde A: englisches "can" ist meist Moeglichkeit oder Eigenschaft
    ("This can lead to a longer cycle time.", "INFO", None, None, None),
    ("No statement can yet be made about the selectivity.", "INFO", None, None, None),
    ("The material list contains all materials that can be installed.", "INFO", None, None, None),
    ("The contractor may propose alternative solutions.", "KANN", None, None, None),
    # --- Runde A: englische Pflicht ohne Modalverb, Imperativ, Soll
    ("Deviations require written approval by the client.", "MUSS", None, None, None),
    ("For each compartment, provide easily accessible valves.", "MUSS", None, None, None),
    ("Use of solvents is described in chapter four.", "INFO", None, None, None),
    ("Supply voltage is 400 V.", "INFO", 400.0, "V", None),
    ("In case that cooling water is required, it is expected that heat exchangers are used.",
     "SOLL", None, None, None),
    # --- Runde A: deutsches "kann" als Verweis / Moeglichkeit / ferne Verneinung
    ("Die Fehlerspezifikation kann dem Anhang (Kapitel 11.01.19) entnommen werden.",
     "INFO", None, None, None),
    ("Schlupf kann zu Fehlern fuehren, dies ist kein Umfang dieses Angebots.",
     "INFO", None, None, None),
    ("Die Kamera kann optional vorgesehen werden, jedoch nicht an Station drei.",
     "KANN", None, None, None),
    ("Das Teil kann ohne Werkzeug nicht montiert werden.", "VERBOT", None, None, None),
    # --- Runde A: Abstimmungspflicht und Lieferunterlagen ohne Modalverb
    ("IP-Adressen werden in Abstimmung mit dem AG vergeben.", "MUSS", None, None, None),
    ("10.2.5 Weitere Dokumentationsunterlagen Pruefprotokolle aller Stationen.",
     "MUSS", None, None, None),
    # --- Runde A: Zeitraeume
    ("Ersatzteile muessen mindestens 10 Jahre verfuegbar sein.", "MUSS", 10.0, "Jahre", "GE"),
    ("All parts have to be available for at least 15 years.", "MUSS", 15.0, "years", "GE"),
    # --- Runde B: Pflichten ohne Modalverb
    ("Bei der Montage duerfen nur Elektrostapler eingesetzt werden.", "MUSS", None, None, None),
    ("The cell may only be entered during setup mode.", "MUSS", None, None, None),
    ("Die Inbetriebnahme erfolgt nur nach erfolgreicher Freigabe.", "MUSS", None, None, None),
    ("Der Kunde behaelt sich vor, Sublieferanten abzulehnen.", "INFO", None, None, None),
    ("Das Werkzeug wird dem Lieferanten zur Verfuegung gestellt.", "INFO", None, None, None),
    ("The definition could be found in standard N54.", "INFO", None, None, None),
    ("We expect our partners to use resources sparingly.", "SOLL", None, None, None),
    ("We require a written confirmation.", "MUSS", None, None, None),
    ("The cycle time should not exceed 30 s.", "SOLL", 30.0, "s", "LE"),
    # verblose Parameterzeilen: Vorgabe aus einer Anforderungstabelle -> MUSS
    ("Ball valve up to DN 50 for the main line.", "MUSS", 50.0, "DN", "LE"),
    ("Cooling time adjustable between 1 - 15 sec.", "MUSS", 1.0, "sec", "RANGE"),
    # dieselbe Zeile mit Messwort bleibt eine Feststellung
    ("Rundlauf max. 0,05 mm gemessen", "INFO", 0.05, "mm", "LE"),
    # Gegenprobe: eine blosse Zahl ohne Grenzwert bleibt draussen (siehe unten)
]

# Runde A: Wird der Satz ueberhaupt aufgenommen? (Satz, erwartet)
AUFNAHME_TEST = [
    ("The number of parts has to be given in", False),            # abgeschnitten
    ("Handover Unit 100mm 70mm Feeding Unit Cut Unit", False),     # Folienfetzen
    ("Holder Cable − Pneumatic Cylinder − Gripper 0,1 mm steps", False),  # Aufzaehlung
    ("Es ist vertraulich und darf ohne Genehmigung weder reproduziert noch "
     "veröffentlicht werden.", False),                             # Schutzvermerk
    ("Der Betriebsdruck der Zylinder beträgt 6 bar.", True),       # echter Satz mit Wert
    ("Das Bedienpanel wird noch festgelegt.", True),               # offener Punkt
    ("The NOK definition has not yet been provided by the customer.", True),  # offener Punkt EN
    ("Die Fehlerspezifikation kann dem Grenzmusterkatalog im Anhang entnommen werden.", True),  # Verweis
    # --- Runde B
    ("Cooling time adjustable between 1 - 15 sec.", True),       # Parameterzeile
    ("Ball valve up to DN 50 for the main line.", True),          # Parameterzeile
    ("Handover Unit 100mm 70mm variable Feeding Unit Cut Unit", False),  # Folienfetzen
    ("Das Werkzeug wird dem Lieferanten zur Verfuegung gestellt.", True),  # Beistellung
    ("Der Kunde behaelt sich vor, Sublieferanten abzulehnen.", True),      # Vorbehalt
    ("The definition could be found in standard N54.", True),     # Verweis EN
    ("Die Taktzeit muss max. 32 s betragen.", True),
]


def self_test() -> int:
    print("Selbsttest requirements.py\n" + "-" * 68)
    fails = 0
    for sent, exp_bind, exp_val, exp_unit, exp_op in SELF_TEST:
        bind, neg, _lang = classify_binding(sent)
        meas = extract_measure(sent)
        ok_b = bind == exp_bind
        ok_v = (meas["value"] == exp_val)
        ok_u = ((meas["unit"] or None) == exp_unit)
        ok_o = ((meas["operator"] or None) == exp_op)
        flag = "ok  " if (ok_b and ok_v and ok_u and ok_o) else "FEHL"
        if flag == "FEHL":
            fails += 1
        print(f'{flag} {bind:<7} {str(meas["operator"]):<6} {str(meas["value"]):<8} '
              f'{str(meas["unit"]):<7} {sent[:44]}')
        if not ok_b:
            print(f"      erwartet Verbindlichkeit {exp_bind}")
        if not ok_v:
            print(f"      erwartet Wert {exp_val}")
        if not ok_u:
            print(f"      erwartet Einheit {exp_unit}")
        if not ok_o:
            print(f"      erwartet Operator {exp_op}")

    print("-" * 68)
    for sent, erwartet in AUFNAHME_TEST:
        ist = is_requirement(sent, classify_binding(sent)[0])
        flag = "ok  " if ist == erwartet else "FEHL"
        fails += ist != erwartet
        print(f'{flag} aufgenommen={str(ist):<5} {sent[:52]}')

    # Dubletten: derselbe Satz in zwei Projekten muss in BEIDEN erscheinen
    import tempfile
    with tempfile.NamedTemporaryFile("w", suffix=".jsonl", delete=False,
                                     encoding="utf-8") as fh:
        for pn in ("A001", "B002"):
            fh.write(json.dumps({"id": pn, "project_no": pn, "filename": "x.pdf",
                                 "text": "Die Anlage muss nach DIN EN 60204 verdrahtet sein."}) + "\n")
        tmp = fh.name
    projekte = sorted(r.project_no for r in build(Path(tmp)))
    Path(tmp).unlink(missing_ok=True)
    ok_d = projekte == ["A001", "B002"]
    fails += not ok_d
    print(f'{"ok  " if ok_d else "FEHL"} gleicher Satz in zwei Projekten -> {projekte}')

    # Versionskette: derselbe Satz in zwei FASSUNGEN desselben Projekts muss
    # ZWEIMAL erscheinen — sonst gibt es keine Kette. Zusaetzlich: derselbe
    # Satz zweimal IN DERSELBEN Datei bleibt eine einzige Anforderung.
    with tempfile.NamedTemporaryFile("w", suffix=".jsonl", delete=False,
                                     encoding="utf-8") as fh:
        for datei in ("LH_V1.01.pdf", "LH_V1.02.pdf"):
            for _ in range(2):          # je Datei doppelt
                fh.write(json.dumps({
                    "id": datei, "project_no": "A001", "filename": datei,
                    "source_path": "P/A001/" + datei,
                    "text": "Die Anlage muss nach DIN EN 60204 verdrahtet sein."}) + "\n")
        tmp = fh.name
    fassungen = build(Path(tmp))
    Path(tmp).unlink(missing_ok=True)
    dateien = sorted(r.file or r.source_path for r in fassungen)
    kette = len({r.subject_key for r in fassungen}) == 1
    ok_v = len(fassungen) == 2 and kette
    fails += not ok_v
    print(f'{"ok  " if ok_v else "FEHL"} gleicher Satz in zwei Fassungen -> '
          f'{len(fassungen)} Eintraege, subject_key gleich: {kette}')

    # Der wichtigste Test: gleiche Anforderung, anderer Wert -> gleicher subject_key
    a = subject_key("Die Taktzeit muss max. 32 s betragen.")
    b = subject_key("Die Taktzeit muss max. 28 s betragen.")
    c = subject_key("Die Schutzart soll mindestens IP54 erreichen.")
    print("-" * 68)
    print(f'{"ok  " if a == b else "FEHL"} subject_key gleich bei geaendertem Wert')
    print(f'{"ok  " if a != c else "FEHL"} subject_key verschieden bei anderer Sache')
    fails += (a != b) + (a == c)

    print("-" * 68)
    print("bestanden" if fails == 0 else f"{fails} Abweichung(en)")
    return 1 if fails else 0


# ==================================================================== 8 CLI
def main() -> int:
    ap = argparse.ArgumentParser(description="Anforderungen aus dem Chunk-Bestand")
    ap.add_argument("--chunks", default=str(OUT / "chunks.jsonl"))
    ap.add_argument("--out", default=str(OUT / "requirements.jsonl"))
    ap.add_argument("--config", default=str(HERE / "routing_config.json"))
    ap.add_argument("--project", default=None)
    ap.add_argument("--show", type=int, default=0, help="N Zeilen auf der Konsole")
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--sample", type=int, default=0,
                    help="N zufaellige Treffer im Volltext — zum Nachzaehlen von Hand")
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--explain", metavar="SATZ", default=None,
                    help="einen einzelnen Satz Schritt für Schritt erklären")
    ap.add_argument("--trace", metavar="DATEINAME", default=None,
                    help="jeden Satz einer Datei zeigen: gelesen oder verworfen, mit Grund")
    ap.add_argument("--suspects", type=int, default=0,
                    help="N Verdachtsfälle zeigen (mögliche Tippfehler im Signalwort)")
    ap.add_argument("--noise", type=int, default=0,
                    help="N Saetze zeigen, die als Muell verworfen wurden")
    a = ap.parse_args()

    if a.self_test:
        return self_test()

    if a.explain:
        r = None
        if Router is not None and Path(a.config).exists():
            r = Router.from_file(Path(a.config))
        print(explain(a.explain, r))
        return 0

    chunks_path = Path(a.chunks)
    if not chunks_path.exists():
        print(f"Nicht gefunden: {chunks_path}\n"
              f"Zuerst extract.py und index.py laufen lassen, oder --chunks setzen.",
              file=sys.stderr)
        return 2

    router = None
    if Router is not None and Path(a.config).exists():
        router = Router.from_file(Path(a.config))

    if a.trace:
        router = None
        if Router is not None and Path(a.config).exists():
            router = Router.from_file(Path(a.config))
        treffer = [c for c in load_chunks(chunks_path)
                   if a.trace.lower() in str(c.get("filename") or
                                             c.get("source_path", "")).lower()]
        if not treffer:
            print(f'Keine Datei gefunden, die "{a.trace}" enthält.', file=sys.stderr)
            print("Vorhandene Dateien:", file=sys.stderr)
            namen = sorted({str(c.get("filename") or Path(str(c.get("source_path",""))).name)
                            for c in load_chunks(chunks_path)})
            for n in namen[:40]:
                print("   ", n, file=sys.stderr)
            return 2

        print(f'DURCHSICHT: {a.trace}   ({len(treffer)} Textabschnitte)')
        print("=" * 100)
        gelesen = behalten = 0
        gesehen: set[str] = set()
        for c in treffer:
            meta = chunk_meta(c)
            print(f'\n--- Abschnitt {meta["chunk_id"]}, Seite {meta["page"]} '
                  f'{"· Überschrift: " + meta["header"] if meta.get("header") else ""}')
            for sent in split_sentences(c.get("text", "")):
                gelesen += 1
                bad, why = is_noise(sent)
                if bad:
                    print(f'  [ ] verworfen ({why})')
                    print(f'      {sent[:110]}')
                    continue
                r = from_sentence(sent, meta, router)
                if r is None:
                    print(f'  [ ] nicht aufgenommen (kein Signalwort, kein Wert)')
                    print(f'      {sent[:110]}')
                    continue
                norm = re.sub(r"\W+", " ", r.text).strip().lower()
                if norm in gesehen:
                    print(f'  [=] Dublette — steht schon weiter oben, wird nur einmal gespeichert')
                    print(f'      {r.text[:110]}')
                    continue
                gesehen.add(norm)
                behalten += 1
                val = ""
                if r.value is not None:
                    val = f' · {r.operator or ""} {r.value:g} {r.unit or ""}'.replace("  ", " ")
                print(f'  [x] {r.binding} · {r.category}{val}'
                      + (f' · {r.lang}' if r.lang == "en" else ""))
                print(f'      {r.text[:110]}')
                print(f'      Regel: {binding_rule(r.text)}')
                if r.suspect:
                    print(f'      Verdacht: {r.suspect}')
        print("\n" + "=" * 100)
        print(f'{gelesen} Sätze gelesen · {behalten} aufgenommen · '
              f'{gelesen - behalten} verworfen oder übergangen')
        print("Prüfe die [x]-Zeilen gegen das Original — Seite steht jeweils dabei.")
        return 0

    if a.noise:
        shown = 0
        print("Verworfene Saetze (mit Begruendung) — zum Pruefen der Filter\n" + "=" * 92)
        for c in load_chunks(chunks_path):
            if a.project and c.get("project_no") != a.project:
                continue
            for sent in split_sentences(c.get("text", "")):
                bad, why = is_noise(sent)
                if bad:
                    print(f"[{why}]\n  {sent[:150]}")
                    shown += 1
                    if shown >= a.noise:
                        return 0
        return 0

    reqs = build(chunks_path, a.project, router)

    out_path = Path(a.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as fh:
        for r in reqs:
            fh.write(json.dumps(asdict(r), ensure_ascii=False) + "\n")

    # Zusammenfassung
    from collections import Counter
    bind = Counter(r.binding for r in reqs)
    cat = Counter(r.category or "—" for r in reqs)
    print(f"{len(reqs)} Anforderungen -> {out_path}")
    print("  Verbindlichkeit:", dict(bind))
    print("  Kategorie:      ", dict(cat))
    print(f"  messbar:         {sum(r.measurable for r in reqs)}")
    print(f"  mit Rückfrage:   {sum(r.needs_question for r in reqs)}")
    print(f"  englisch:        {sum(r.lang == 'en' for r in reqs)}")
    print(f"  Verdachtsfälle:  {sum(bool(r.suspect) for r in reqs)}  "
          f"(mit --suspects ansehen)")
    print(f"  Gegenstände:     {len({r.subject_key for r in reqs})} "
          f"(mehrfach belegt: "
          f"{len(reqs) - len({r.subject_key for r in reqs})})")

    if a.suspects:
        v = [r for r in reqs if r.suspect]
        print(f"\n{len(v)} Verdachtsfälle — Sätze, die wie eine Anforderung aussehen,")
        print("aber kein sauberes Signalwort haben. Ein Mensch entscheidet.\n" + "=" * 92)
        for r in v[:a.suspects]:
            print(f"\n[{r.suspect}]")
            print(f"  {r.text[:150]}")
            print(f"  Quelle: {r.file}, S. {r.page}")
        return 0

    if a.sample:
        import random
        random.seed(a.seed)
        pick = random.sample(reqs, min(a.sample, len(reqs)))
        print("\nSTICHPROBE — bitte je Zeile ankreuzen: Anforderung ja / nein")
        print("=" * 92)
        for i, r in enumerate(pick, 1):
            val = ""
            if r.value is not None:
                val = f' [{r.operator or ""} {r.value:g} {r.unit or ""}]'.replace("  ", " ")
            print(f"\n{i:>3}. {r.binding} · {r.category}{val}")
            print(f"     {r.text}")
            print(f"     Quelle: {r.file}, S. {r.page}  ({r.project_no})")
        print("\n" + "=" * 92)
        print(f"Von {len(pick)} Zeilen sind ____ echte Anforderungen.")
        print("Diese Zahl geteilt durch die Stichprobengroesse ist die Treffergenauigkeit.")
        print("Ohne diese Zaehlung ist jede Aussage ueber die Qualitaet geraten.")
        return 0

    if a.show:
        print("\n" + "-" * 100)
        for r in reqs[:a.show]:
            val = ""
            if r.value is not None:
                val = f'{r.operator or ""} {r.value:g} {r.unit or ""}'.strip()
            print(f'{r.binding:<7} {r.category[:12]:<13} {val:<16} '
                  f'{r.text[:52]:<54} {r.file[:22]}')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
