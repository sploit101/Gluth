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
    r"unzul[aä]ssig|untersagt|verboten|ausgeschlossen|ohne)\b", re.I)

# "muß" ist die Schreibweise vor 1996 und steht in vielen aelteren
# Lastenheften. Ohne sie faellt jede Pflicht in solchen Dokumenten durch.
MUSS = re.compile(
    r"\b(mu(?:ss|ß)|m[uü](?:ss|ß)en|muessen|ist zu\b|sind zu\b|hat zu\b|haben zu\b|zwingend\w{0,3}|"
    r"verbindlich\w{0,4}|erforderlich\w{0,4}|vorgeschrieben\w{0,4}|"
    r"zu gew[aä]hrleisten|einzuhalten|sicherzustellen|verpflichtend\w{0,3}|"
    r"wird gefordert|gefordert wird)\b", re.I)

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
           "any", "such", "which", "from", "has", "have", "will", "it"}

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
KANN_EN = re.compile(r"\b(may|can|optional|optionally|at the discretion)\b", re.I)
VERBOT_EN = re.compile(r"\b(shall not|must not|may not|cannot|is not permitted|"
                       r"are not permitted|is prohibited|are prohibited|"
                       r"is not allowed|are not allowed)\b", re.I)

# Das deutsche Gerundivum: "ist in Schutzart IP54 auszufuehren",
# "ist nach DIN EN 60204 auszufuehren", "sind vorzusehen", "ist einzuhalten".
# In Lastenheften die haeufigste MUSS-Formulierung ueberhaupt — und sie wird
# von der Wortliste oben nicht erfasst, weil zwischen "ist" und dem
# zu-Infinitiv beliebig viele Woerter stehen koennen.
GERUND = re.compile(
    r"\b(?:ist|sind|war|waren)\s+(?:\S+\s+){0,7}?"
    r"(?:zu\s+\w{3,}en|\w{2,}zu\w{2,}en)\b", re.I)

GERUND_VERBOT = re.compile(
    r"\b(?:ist|sind)\s+nicht\s+(?:\S+\s+){0,5}?"
    r"(?:zu\s+\w{3,}en|\w{2,}zu\w{2,}en)\b", re.I)

VERBOT_DIREKT = re.compile(
    r"(darf\s+(?:\w+\s+){0,3}?nicht|d[uü]rfen\s+(?:\w+\s+){0,3}?nicht|"
    r"ist\s+(?:\w+\s+){0,2}?unzul[aä]ssig|sind\s+(?:\w+\s+){0,2}?unzul[aä]ssig|"
    r"ist untersagt|sind untersagt|ist verboten|nicht zul[aä]ssig|"
    r"nicht gestattet|nicht erlaubt|shall not|must not|darf kein)", re.I)


def classify_binding(sentence: str) -> tuple[str, bool, str]:
    """Gibt (Verbindlichkeit, verneint, Sprache) zurueck."""
    lang = detect_lang(sentence)
    negated = bool(NEG.search(sentence))

    # ---------------------------------------------------------- englisch
    if lang == "en":
        if VERBOT_EN.search(sentence):
            return "VERBOT", True, lang
        if MUSS_EN.search(sentence):
            return "MUSS", negated, lang
        if SOLL_EN.search(sentence):
            return "SOLL", negated, lang
        if KANN_EN.search(sentence):
            return "KANN", False, lang
        return "INFO", negated, lang

    # ---------------------------------------------------------- deutsch
    # 1. Verbot hat immer Vorrang
    if VERBOT_DIREKT.search(sentence) or GERUND_VERBOT.search(sentence):
        return "VERBOT", True, lang

    # 1b. Gerundivum ist eine MUSS-Formulierung
    if GERUND.search(sentence):
        return "MUSS", negated, lang

    # 2. Verneintes MUSS bleibt ein Verbot ("es muss vermieden werden")
    if MUSS.search(sentence):
        return ("VERBOT", True, lang) if negated and re.search(
            r"\b(vermieden|unterbleiben|verhindert|ausgeschlossen)\b", sentence, re.I
        ) else ("MUSS", negated, lang)

    if SOLL.search(sentence):
        return "SOLL", negated, lang

    # 3. KANN erst hier — und nur, wenn der Satz nicht verneint ist
    if KANN.search(sentence) and not negated:
        return "KANN", False, lang
    if KANN.search(sentence) and negated:
        return "VERBOT", True, lang

    return "INFO", negated, lang


# --- Welche Regel hat entschieden? ---------------------------------------
# Dieselbe Reihenfolge wie in classify_binding, aber mit Namen. Damit laesst
# sich jede einzelne Entscheidung nachvollziehen, statt sie glauben zu muessen.
def binding_rule(sentence: str) -> str:
    lang = detect_lang(sentence)
    if lang == "en":
        if VERBOT_EN.search(sentence):
            return f'englisches Verbotswort: "{VERBOT_EN.search(sentence).group(0)}"'
        if MUSS_EN.search(sentence):
            return f'englisches Pflichtwort: "{MUSS_EN.search(sentence).group(0)}"'
        if SOLL_EN.search(sentence):
            return f'englisches Sollwort: "{SOLL_EN.search(sentence).group(0)}"'
        if KANN_EN.search(sentence):
            return f'englisches Kannwort: "{KANN_EN.search(sentence).group(0)}"'
        return "kein englisches Signalwort gefunden"
    if VERBOT_DIREKT.search(sentence):
        return f'Verbot: "{VERBOT_DIREKT.search(sentence).group(0)}"'
    if GERUND_VERBOT.search(sentence):
        return f'verneintes Gerundivum: "{GERUND_VERBOT.search(sentence).group(0)}"'
    if GERUND.search(sentence):
        return f'Gerundivum (ist ... zu ...): "{GERUND.search(sentence).group(0)}"'
    if MUSS.search(sentence):
        return f'Pflichtwort: "{MUSS.search(sentence).group(0)}"'
    if SOLL.search(sentence):
        return f'Sollwort: "{SOLL.search(sentence).group(0)}"'
    if KANN.search(sentence):
        m = KANN.search(sentence).group(0)
        return (f'Kannwort: "{m}"' if not NEG.search(sentence)
                else f'Kannwort "{m}" mit Verneinung -> Verbot')
    return "kein Signalwort gefunden"


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
    import difflib
    for tok in re.findall(r"[A-Za-zÄÖÜäöüß]{3,13}", sentence):
        low = tok.lower()
        if low in KEIN_TIPPFEHLER:
            continue
        near = difflib.get_close_matches(low, SIGNAL_WORDS, n=1, cutoff=0.85)
        # Nur melden, wenn sich die Woerter auch am Anfang gleichen — sonst
        # meldet der Vergleich zufaellige Aehnlichkeiten in der Wortmitte.
        if near and low != near[0] and low[:2] == near[0][:2]:
            return f'moeglicher Tippfehler: "{tok}" ähnelt "{near[0]}"'
    if measure.get("value") is not None and measure.get("unit"):
        return "Zahlenwert mit Einheit, aber kein Signalwort"
    if CLAUSE.match(sentence):
        return "Gliederungsnummer, aber kein Signalwort"
    return ""


# ================================================================== 3 Zahlen
# Deutsches Format: 1.070 sind tausendsiebzig, 0,05 sind null Komma null fuenf.
# Wird das amerikanisch gelesen, liegen Werte um Faktor 1000 daneben.

UNITS = (
    r"mm²|cm²|m²|mm³|m³|"
    r"µm|um|mm|cm|dm|m\b|km|"
    r"mg|g\b|kg|t\b|"
    r"ms|s\b|sek|min|h\b|Std|"
    r"°C|°|K\b|"
    r"mbar|bar|Pa|kPa|MPa|N/mm²|Nm|N\b|"
    r"mV|V\b|kV|mA|A\b|kW|MW|W\b|VA|Hz|kHz|"
    r"dB\(A\)|dB|"
    r"ml|l/min|l\b|m/s|mm/s|U/min|min-1|"
    r"%|Stk|St\.|Zoll|IP\d{2}|IP\s?\d{2}"
)
_NUM = r"\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?|\d+(?:\.\d+)?"

# Achtung "min": das ist zugleich eine Einheit (Minuten, l/min, U/min).
# Deshalb hier nur "min." mit Punkt oder "mind"/"mindestens" — und nie
# direkt hinter einem Schraegstrich. Sonst wird aus "200 l/min" ein
# "mindestens 200" und die Anforderung kippt in ihr Gegenteil.
OP_LE = re.compile(r"(?<![/\w])(max\.|maximal|h[oö]chstens|nicht mehr als|bis zu|begrenzt|beschr[aä]nkt|nicht gr[oö][ßs]er als|kleiner gleich|≤|<=)", re.I)
OP_GE = re.compile(r"(?<![/\w])(min\.|mind\.|mindestens|wenigstens|nicht weniger als|gr[oö]ßer gleich|≥|>=)", re.I)
# Formulierungen, die den Operator im Verb tragen. Sie werden zuerst geprueft,
# weil sie eine Verneinung enthalten und sonst falsch herum gelesen wuerden.
PHRASE_LE = re.compile(r"nicht\s+(?:zu\s+)?(?:ue|[uü])(?:berschreit|bersteig|berschritten)", re.I)
PHRASE_GE = re.compile(r"nicht\s+(?:zu\s+)?(?:unterschreit|unterschritten|unterbiet)", re.I)
OP_LT = re.compile(r"\b(kleiner als|unter|weniger als|<)\b", re.I)
OP_GT = re.compile(r"\b(gr[oö]ßer als|[uü]ber|mehr als|>)\b", re.I)
OP_EQ = re.compile(r"\b(genau|exakt|betr(?:ae|[aä])g(?:t|en)|entspricht|=)\b", re.I)
NOLETTER = r"(?![A-Za-zÄÖÜäöüß])"
RANGE = re.compile(rf"({_NUM})\s*(?:bis|\.\.\.|–|-|to)\s*({_NUM})\s*({UNITS}){NOLETTER}")
# _NUM enthaelt selbst Alternativen (|). Ohne eigene Klammer bricht die
# Alternation nach aussen aus und das Muster trifft jede beliebige Zahl —
# so wurde aus "IP54" eine Toleranz von 54.
TOL = re.compile(rf"([+±]\s*/?\s*-?\s*(?:{_NUM})\s*(?:{UNITS})?|\+\s*(?:{_NUM})\s*/\s*-\s*(?:{_NUM}))")
VALUE = re.compile(rf"({_NUM})\s*({UNITS}){NOLETTER}")
IPCODE = re.compile(r"\bIP\s?(\d{2})\b")


def parse_german_number(raw: str) -> float | None:
    """'1.070,5' -> 1070.5 · '0,05' -> 0.05 · '1.070' -> 1070 · '3.2' -> 3.2"""
    s = raw.strip()
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

    m = VALUE.search(sentence)
    if m:
        out["value"] = parse_german_number(m.group(1))
        out["unit"] = m.group(2).strip()

    out["operator"] = detect_operator(sentence)
    return out


def detect_operator(sentence: str) -> str | None:
    # Verbformulierungen zuerst — sie tragen eine Verneinung in sich
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
    r"vom Kunden (?:zu )?(?:liefern|benennen)|k\.\s?A\.)\b", re.I)

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

# Ein Satz, der mitten in der Pflicht abbricht, ist keine Anforderung,
# sondern ein Chunk-Ende: "These parts ... also have to"
ABBRUCH_WORTE = {
    "ist", "sind", "hat", "haben", "wird", "werden", "muss", "müssen", "soll",
    "sollen", "kann", "können", "darf", "dürfen", "zu", "und", "oder", "der",
    "die", "das", "den", "dem", "des", "für", "mit", "von", "bei", "als",
    "has", "have", "must", "shall", "should", "may", "can", "be", "to", "the",
    "of", "and", "or", "for", "with", "is", "are",
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


TOC_DOTS = re.compile(r"\.{4,}|…{2,}")            # "Packaging ........ 32"
TOC_START = re.compile(r"^\d+(\.\d+)*\s+\S")      # "6.1 Packaging"
PIPES = re.compile(r"\|")                          # Fuss- und Kopfzeilen
WORD3 = re.compile(r"[A-Za-zÄÖÜäöüß]{3,}")


def is_noise(sentence: str) -> tuple[bool, str]:
    """Gibt (ist_Muell, Grund) zurueck — der Grund ist fuer die Fehlersuche da."""
    s = sentence.strip()
    if NOISE.match(s):
        return True, "Kopf- oder Fusszeile"
    if MAIL_NOISE.search(s):
        return True, "E-Mail-Anrede, Grussformel oder Vertraulichkeitshinweis"
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
    if OPEN_Q.search(sentence):
        return True
    m = extract_measure(sentence)
    if m["value"] is not None and m["unit"]:
        return True
    return bool(suspect_reason(sentence, binding, m))


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
            # schreibung zaehlt als dieselbe Anforderung. Chunk-Ueberlappung
            # und wiederholte Textbausteine erzeugen sonst Karteileichen.
            norm = re.sub(r"\W+", " ", (r.text if r else "")).strip().lower()
            if r and r.req_id not in seen and norm not in seen:
                seen.add(r.req_id)
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
