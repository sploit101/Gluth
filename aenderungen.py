#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aenderungen.py — "Es kam eine neue Fassung des Lastenhefts. Was ist fuer
meine Station anders?"

Das ist Daniels Schritt 2 (Aktualisierung der Akte). ketten.py zeigt nur
Paare, die denselben subject_key haben. Hier wird die ganze Fassung
verglichen und jede Anforderung bekommt genau einen Status:

    NEU                       nur in der neuen Fassung
    ENTFALLEN                 nur in der alten Fassung
    WERT GEAENDERT            gleiche Anforderung, andere Zahl
    VERBINDLICHKEIT GEAENDERT z. B. SOLL -> MUSS
    TEXT GEAENDERT            gleiche Anforderung, anderer Wortlaut
    unveraendert              (auch: nur Darstellung - Aufzaehlungszeichen o. ae.)

Zuordnung alt <-> neu in drei Stufen, jede nachvollziehbar:
    1. gleicher Kerntext (ohne Aufzaehlungszeichen und Randnummern)
    2. gleicher subject_key (Gegenstand ohne Zahlen), nur wenn eindeutig
    3. aehnlicher Text ohne Zahlen (difflib, Schwelle AEHNLICH), bestes Paar zuerst

Jede Aenderung wird fuer die Station eingeordnet (stationen.py) und mit einer
kurzen deutschen Beschreibung versehen: "Abisolierlänge (Toleranz): ± 0,3 mm -> ± 0,2 mm".
Kein Sprachmodell.

    python aenderungen.py --datei LH_04K_000051                  # letzte zwei Fassungen
    python aenderungen.py --alt V1.01 --neu V1.02 --station "Station 20 Abisolieren"
    python aenderungen.py --datei LH_04K --stationsliste stationen.csv

Ausgabe
    out/aenderungen.xlsx         mit Wortlauten   -> VERTRAULICH
    out/aenderungen_zahlen.txt   nur Zahlen       -> darf geteilt werden
"""

from __future__ import annotations

import argparse
import collections
import difflib
import json
import re
from datetime import datetime
from pathlib import Path

import stationen as st
from ketten import kern, norm, wert

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"

AEHNLICH = 0.72                # Mindest-Aehnlichkeit fuer Stufe 3
ZAHL = re.compile(r"\d+(?:[.,]\d+)?")

STATUS_RANG = {"VERBINDLICHKEIT GEAENDERT": 0, "WERT GEAENDERT": 1, "NEU": 2,
               "ENTFALLEN": 3, "TEXT GEAENDERT": 4, "unveraendert": 9}
GELTUNG_RANG = {"Station": 0, "Anlagenweit": 1, "Unklar": 2, "Andere Station": 3, "": 1}

# Versionskennungen im Dateinamen: _V1.02, V2.1, Rev 3, _008, -r2, Version 1.01
VERSION = re.compile(
    r"(?:[\s_-]*(?:v|ver|version|rev|r|stand)\.?\s*(\d+(?:[._]\d+)*)|_(\d{2,3}))\s*$", re.I)


def ascii_safe(s) -> str:
    return str(s).encode("ascii", "replace").decode("ascii")


# --------------------------------------------------------------- Fassungen
def familie(dateiname: str) -> tuple[str, tuple[int, ...]]:
    """'LH_04K_000051_V1.02.pdf' -> ('lh_04k_000051', (1, 2))."""
    stamm = Path(dateiname).stem.strip()
    m = VERSION.search(stamm)
    if not m:
        return stamm.lower(), ()
    roh = m.group(1) or m.group(2)
    nummer = tuple(int(x) for x in re.split(r"[._]", roh) if x.isdigit())
    return stamm[:m.start()].strip(" _-").lower(), nummer


def fassungspaare(dateien: list[str]) -> list[tuple[str, str]]:
    """Je Dokumentfamilie mit mehreren Fassungen: (vorletzte, letzte) und davor."""
    fam: dict[str, list[tuple[tuple[int, ...], str]]] = collections.defaultdict(list)
    for d in sorted(set(dateien)):
        basis, nr = familie(d)
        if nr:
            fam[basis].append((nr, d))
    paare = []
    for basis, liste in fam.items():
        liste.sort()
        for (_, alt), (_, neu) in zip(liste, liste[1:]):
            paare.append((alt, neu))
    return paare


def vorgaenger(datei: str, dateien: list[str]) -> str | None:
    """Die naechstaeltere Fassung derselben Familie, falls vorhanden."""
    basis, nr = familie(datei)
    if not nr:
        return None
    kandidaten = sorted((familie(d)[1], d) for d in set(dateien)
                        if familie(d)[0] == basis and familie(d)[1] and familie(d)[1] < nr)
    return kandidaten[-1][1] if kandidaten else None


# --------------------------------------------------------------- Vergleich
def _skelett(text: str) -> str:
    return re.sub(r"\s+", " ", ZAHL.sub("#", kern(text))).strip()


def _woerter(text: str) -> set[str]:
    return set(re.findall(r"[a-zäöüß]{3,}", st.falten(text)))


def wortdiff(alt: str, neu: str, laenge: int = 220) -> str:
    """Kurzer Wortunterschied: [-alt-] {+neu+}, nur die geaenderten Stellen."""
    a, b = norm(alt).split(), norm(neu).split()
    teile = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b).get_opcodes():
        if op == "equal":
            continue
        if i2 > i1:
            teile.append("[-" + " ".join(a[i1:i2]) + "-]")
        if j2 > j1:
            teile.append("{+" + " ".join(b[j1:j2]) + "+}")
    x = " ".join(teile)
    return x if len(x) <= laenge else x[:laenge - 1] + "…"


def status(a: dict, b: dict) -> str:
    ka, kb = kern(a["text"]), kern(b["text"])
    if ka == kb:
        return "unveraendert"
    if ZAHL.findall(ka) != ZAHL.findall(kb) or wert(a) != wert(b):
        return "WERT GEAENDERT"
    if (a.get("binding") or "") != (b.get("binding") or ""):
        return "VERBINDLICHKEIT GEAENDERT"
    if difflib.SequenceMatcher(None, ka, kb).ratio() >= 0.98:
        return "unveraendert"                     # nur Darstellung
    return "TEXT GEAENDERT"


def wertaenderung(alt: str, neu: str) -> str:
    """'Abisolierlänge (Toleranz): ± 0,3 mm → ± 0,2 mm' - aus den Kennwerten beider Fassungen."""
    pa = {p["parameter"]: p for p in st.parameter(alt)}
    pb = {p["parameter"]: p for p in st.parameter(neu)}
    teile = []
    for name in pa.keys() & pb.keys():
        x, y = pa[name], pb[name]
        fx = f"{'' if x['bedingung'] == '=' else x['bedingung'] + ' '}{x['wert']} {x['einheit']}"
        fy = f"{'' if y['bedingung'] == '=' else y['bedingung'] + ' '}{y['wert']} {y['einheit']}"
        if fx != fy:
            teile.append(f"{name}: {fx} → {fy}")
    ka, kb = st.kompakt(alt)["kennwert"], st.kompakt(neu)["kennwert"]
    if ka != kb and (ka or kb):
        teile.append(f"{', '.join(ka) or '–'} → {', '.join(kb) or '–'}")
    if teile:
        return "; ".join(teile)
    za, zb = ZAHL.findall(kern(alt)), ZAHL.findall(kern(neu))
    if za != zb:
        weg = [z for z in za if z not in zb]
        dazu = [z for z in zb if z not in za]
        return f"Zahl {', '.join(weg) or '–'} → {', '.join(dazu) or '–'}"
    return ""


def vergleiche(alt: list[dict], neu: list[dict]) -> list[dict]:
    """
    Liefert eine Liste von Aenderungen {status, alt, neu, stufe}.
    alt/neu sind Anforderungszeilen aus requirements.jsonl (oder None).
    """
    rest_a = list(range(len(alt)))
    rest_b = list(range(len(neu)))
    paare: list[tuple[int, int, str]] = []

    # Stufe 1: gleicher Kerntext
    nach_kern: dict[str, list[int]] = collections.defaultdict(list)
    for i in rest_a:
        nach_kern[kern(alt[i]["text"])].append(i)
    for j in list(rest_b):
        k = kern(neu[j]["text"])
        if nach_kern.get(k):
            i = nach_kern[k].pop(0)
            paare.append((i, j, "gleicher Text"))
            rest_a.remove(i)
            rest_b.remove(j)

    # Stufe 2: gleicher subject_key, nur 1:1
    sk_a = collections.defaultdict(list)
    sk_b = collections.defaultdict(list)
    for i in rest_a:
        if alt[i].get("subject_key"):
            sk_a[alt[i]["subject_key"]].append(i)
    for j in rest_b:
        if neu[j].get("subject_key"):
            sk_b[neu[j]["subject_key"]].append(j)
    for k in sk_a.keys() & sk_b.keys():
        if len(sk_a[k]) == 1 and len(sk_b[k]) == 1:
            i, j = sk_a[k][0], sk_b[k][0]
            paare.append((i, j, "gleicher Gegenstand"))
            rest_a.remove(i)
            rest_b.remove(j)

    # Stufe 3: aehnlicher Text ohne Zahlen, bestes Paar zuerst
    kandidaten = []
    wa = {i: _woerter(alt[i]["text"]) for i in rest_a}
    wb = {j: _woerter(neu[j]["text"]) for j in rest_b}
    sa = {i: _skelett(alt[i]["text"]) for i in rest_a}
    sb = {j: _skelett(neu[j]["text"]) for j in rest_b}
    for i in rest_a:
        for j in rest_b:
            gemeinsam = len(wa[i] & wb[j]) / max(len(wa[i] | wb[j]), 1)
            if gemeinsam < 0.35:
                continue
            sm = difflib.SequenceMatcher(None, sa[i], sb[j])
            if sm.quick_ratio() < AEHNLICH:
                continue
            r = sm.ratio()
            if r >= AEHNLICH:
                kandidaten.append((r, i, j))
    vergeben_a, vergeben_b = set(), set()
    for r, i, j in sorted(kandidaten, reverse=True):
        if i in vergeben_a or j in vergeben_b:
            continue
        vergeben_a.add(i)
        vergeben_b.add(j)
        paare.append((i, j, f"aehnlicher Text ({r:.0%})"))
    rest_a = [i for i in rest_a if i not in vergeben_a]
    rest_b = [j for j in rest_b if j not in vergeben_b]

    raus = []
    for i, j, stufe in paare:
        a, b = alt[i], neu[j]
        stt = status(a, b)
        raus.append({"status": stt, "alt": a, "neu": b, "stufe": stufe,
                     "detail": (wertaenderung(a["text"], b["text"]) if stt == "WERT GEAENDERT" else
                                f"{a.get('binding', '')} → {b.get('binding', '')}"
                                if stt == "VERBINDLICHKEIT GEAENDERT" else ""),
                     "wortdiff": wortdiff(a["text"], b["text"]) if stt != "unveraendert" else ""})
    for j in rest_b:
        raus.append({"status": "NEU", "alt": None, "neu": neu[j], "stufe": "-",
                     "detail": "; ".join(st.kennwerte_de(neu[j]["text"])), "wortdiff": ""})
    for i in rest_a:
        raus.append({"status": "ENTFALLEN", "alt": alt[i], "neu": None, "stufe": "-",
                     "detail": "; ".join(st.kennwerte_de(alt[i]["text"])), "wortdiff": ""})
    return raus


def einordnen(aenderungen: list[dict], profil: st.Profil | None,
              alle: list[st.Profil] | None) -> list[dict]:
    """Haengt Station-Bewertung und Kurztext an jede Aenderung (Kopie), sortiert."""
    raus = []
    for a in aenderungen:
        z = a["neu"] or a["alt"]
        b = st.bewerte(z["text"], profil, alle, z.get("section") or "")
        if a["status"] == "WERT GEAENDERT" and a["detail"]:
            kurz = a["detail"]
        elif a["status"] == "VERBINDLICHKEIT GEAENDERT":
            kurz = f"{b.thema}: {a['detail']}"
        elif a["status"] in ("NEU", "ENTFALLEN"):
            kurz = st.kurzfassung(z.get("binding", ""), b, st.kennwerte_de(z["text"]),
                                  st.fachbegriffe(z["text"]))
        elif a["status"] == "TEXT GEAENDERT":
            kurz = f"{b.thema}: Wortlaut geändert"
        else:
            kurz = ""
        raus.append({**a, "bew": b, "kurz": kurz})
    raus.sort(key=lambda a: (a["status"] == "unveraendert",
                             GELTUNG_RANG.get(a["bew"].geltung, 9),
                             STATUS_RANG.get(a["status"], 9),
                             str((a["neu"] or a["alt"]).get("page") or "").zfill(5)))
    return raus


# --------------------------------------------------------------- Ausgabe
SPALTEN = ["Status", "Einordnung", "Thema", "Änderung (kurz, DE)", "Wortunterschied",
           "Seite alt", "Seite neu", "Verb. alt", "Verb. neu", "Wortlaut alt", "Wortlaut neu",
           "Kapitel", "Warum diese Einordnung", "Zuordnung alt↔neu", "Stimmt?"]
BREITEN = [22, 14, 22, 40, 50, 9, 9, 9, 9, 60, 60, 28, 34, 22, 9]


def blatt_schreiben(ws, liste: list[dict], fm: dict) -> None:
    """Schreibt die Aenderungstabelle in ein vorhandenes xlsxwriter-Blatt."""
    for i, (s_, b_) in enumerate(zip(SPALTEN, BREITEN)):
        ws.write(0, i, s_, fm["kopf"])
        ws.set_column(i, i, b_)
    ws.freeze_panes(1, 1)
    n = 0
    for a in liste:
        if a["status"] == "unveraendert":
            continue
        n += 1
        alt, neu, b = a["alt"] or {}, a["neu"] or {}, a["bew"]
        f = fm["grau"] if b.geltung == "Andere Station" else (
            fm["gruen"] if b.geltung == "Station" else fm["zelle"])
        ws.write_row(n, 0, [
            a["status"], b.geltung or "–", b.thema, a["kurz"], a["wortdiff"],
            alt.get("page", ""), neu.get("page", ""), alt.get("binding", ""),
            neu.get("binding", ""), norm(alt.get("text", "")), norm(neu.get("text", "")),
            (neu or alt).get("section", ""), b.warum, a["stufe"], ""], f)
    ws.autofilter(0, 0, max(n, 1), len(SPALTEN) - 1)
    if n:
        ws.data_validation(1, len(SPALTEN) - 1, n, len(SPALTEN) - 1,
                           {"validate": "list", "source": ["ja", "nein"]})
    else:
        ws.write(1, 0, "Keine inhaltlichen Änderungen gefunden.", fm["zelle"])


def zaehlen(liste: list[dict]) -> dict:
    z = collections.Counter(a["status"] for a in liste)
    je = collections.Counter((a["bew"].geltung, a["status"]) for a in liste
                             if a["status"] != "unveraendert")
    return {"status": dict(z), "geltung_status": {f"{g or '-'}|{s}": n for (g, s), n in je.items()}}


def lade(pfad: Path) -> list[dict]:
    rows = []
    for line in pfad.open(encoding="utf-8"):
        line = line.strip()
        if line:
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return rows


def datei_von(r: dict) -> str:
    return str(r.get("file") or r.get("filename") or r.get("source_file") or "")


def fuer_datei(rows: list[dict], datei: str) -> list[dict]:
    return [r for r in rows if datei_von(r) == datei]


def main() -> int:
    ap = argparse.ArgumentParser(description="Was hat sich zwischen zwei Fassungen geaendert?")
    ap.add_argument("--req", default=str(OUT / "requirements.jsonl"))
    ap.add_argument("--datei", help="nur Dokumentfamilien, deren Name dies enthaelt")
    ap.add_argument("--alt", help="Namensteil der alten Fassung")
    ap.add_argument("--neu", help="Namensteil der neuen Fassung")
    ap.add_argument("--station")
    ap.add_argument("--stationsliste")
    ap.add_argument("--aus", default=str(OUT / "aenderungen.xlsx"))
    a = ap.parse_args()

    pfad = Path(a.req)
    if not pfad.exists():
        print("FEHLER: nicht gefunden:", ascii_safe(pfad), "- erst requirements.py laufen lassen.")
        return 2
    rows = lade(pfad)
    dateien = sorted({datei_von(r) for r in rows if datei_von(r)})

    if a.alt or a.neu:
        if not (a.alt and a.neu):
            print("FEHLER: --alt und --neu gehoeren zusammen.")
            return 2
        alt_d = [d for d in dateien if a.alt.lower() in d.lower()]
        neu_d = [d for d in dateien if a.neu.lower() in d.lower()]
        if len(alt_d) != 1 or len(neu_d) != 1:
            print(f"FEHLER: --alt passt auf {len(alt_d)} Dateien, --neu auf {len(neu_d)}. "
                  "Es muss genau eine sein.")
            for d in alt_d + neu_d:
                print("   ", ascii_safe(d))
            return 2
        paare = [(alt_d[0], neu_d[0])]
    else:
        auswahl = [d for d in dateien if not a.datei or a.datei.lower() in d.lower()]
        paare = fassungspaare(auswahl)
    if not paare:
        print("Keine zwei Fassungen desselben Dokuments gefunden "
              "(Versionskennung im Dateinamen, z. B. _V1.01 / _V1.02).")
        return 1

    alle = st.stationsliste_laden(Path(a.stationsliste)) if a.stationsliste else []
    profil = st.profil_aus_name(a.station) if a.station else None
    if profil and alle:
        profil = next((p for p in alle if (p.nummer and p.nummer == profil.nummer)
                       or st.falten(p.name) == st.falten(profil.name)), profil)

    import xlsxwriter
    OUT.mkdir(exist_ok=True)
    wb = xlsxwriter.Workbook(a.aus)
    fm = {
        "kopf": wb.add_format({"bold": True, "bg_color": "#1F3B57", "font_color": "#FFFFFF",
                               "border": 1, "text_wrap": True, "valign": "top"}),
        "zelle": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
        "gruen": wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                                "bg_color": "#E8F5EE"}),
        "grau": wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                               "font_color": "#777777"}),
    }
    L = [f"AENDERUNGEN  {datetime.now():%Y-%m-%d %H:%M}", "=" * 62, ""]
    if profil:
        L.append(f"  Station (Angabe Mensch)   {ascii_safe(profil.name)}")
    for n, (alt_d, neu_d) in enumerate(paare, start=1):
        liste = einordnen(vergleiche(fuer_datei(rows, alt_d), fuer_datei(rows, neu_d)),
                          profil, alle or None)
        ws = wb.add_worksheet(f"Vergleich {n}")
        blatt_schreiben(ws, liste, fm)
        z = zaehlen(liste)
        L += ["", f"  Vergleich {n}  (Dateinamen stehen nur in der Excel-Datei)"]
        for s_ in sorted(z["status"], key=lambda k: STATUS_RANG.get(k, 9)):
            L.append(f"    {z['status'][s_]:>5}  {s_}")
        if profil:
            fs = sum(v for k, v in z["geltung_status"].items() if k.startswith("Station|"))
            L.append(f"    davon fuer die Station relevant: {fs}")
        print(f"Vergleich {n}: {ascii_safe(alt_d)} -> {ascii_safe(neu_d)}")
    wb.close()
    L += ["", "Diese Datei enthaelt keine Dateinamen und keine Saetze."]
    (OUT / "aenderungen_zahlen.txt").write_text("\n".join(L), encoding="utf-8")
    for z_ in L:
        print(ascii_safe(z_))
    print(f"\n  {ascii_safe(a.aus)}  mit Wortlauten - VERTRAULICH")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
