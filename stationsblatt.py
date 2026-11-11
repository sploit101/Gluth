#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stationsblatt.py — das, was ein Konstrukteur fuer SEINE Station wirklich liest.

Die Zuordnungsliste (zuordnung.xlsx) ist ein Pruefblatt fuer den, der die Akte
befuellt: eine Zeile je Vorschlag, mit Freigabe-Spalte. Ein Konstrukteur, der
Station 20 auslegt, will etwas anderes:

    1. Faehigkeitsmatrix   Verfahren | Kennwert | Toleranz | Quelle   (eine Tabelle)
    2. Diese Station       nur Saetze, die zu seiner Station gehoeren
    3. Anlagenweit         was fuer alle gilt - einmal, nach Thema sortiert
    4. Unklar              Saetze, bei denen kein Hinweis auf eine Station besteht
    5. Andere Stationen    zur Kontrolle, dass nichts Falsches aussortiert wurde
    6. Luecken             was in den Dokumenten NICHT gefunden wurde

Alles stammt woertlich aus den Quellen (Spalte "Wortlaut", mit Datei und Seite).
Die deutsche Kurzfassung ist ein Stichwortzettel aus Regeln, keine Uebersetzung -
massgeblich bleibt immer der Wortlaut.

Die letzte Spalte "Stimmt?" ist fuer Rueckmeldungen gedacht: Wer ja/nein
eintraegt, liefert die Messgrundlage, um die Regeln zu verbessern.
"""

from __future__ import annotations

import collections
from datetime import datetime
from pathlib import Path

import aenderungen as aend
import stationen as st

BINDUNG_RANG = {"MUSS": 0, "VERBOT": 1, "SOLL": 2, "KANN": 3, "INFO": 4}

GELTUNG_TEXT = {
    "Station": "Gilt für diese Station (Stationsnummer oder Verfahren genannt)",
    "Anlagenweit": "Gilt für alle Stationen (Sicherheit, Fähigkeit, Dokumentation, Abnahme, Elektrik ...)",
    "Unklar": "Kein Hinweis auf eine Station – bitte von Hand entscheiden",
    "Andere Station": "Gehört zu einer anderen Station – nur zur Kontrolle aufgeführt",
}


def _seiten(seiten) -> str:
    teile = sorted({str(s) for s in seiten if s not in (None, "")},
                   key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else 0, x))
    return ", ".join(teile)


def bewerten(saetze: list[dict], profil: st.Profil | None,
             alle: list[st.Profil] | None) -> list[dict]:
    """Haengt Bewertung und Kurzfassung an jeden Satz (Kopie)."""
    raus = []
    for s in saetze:
        b = st.bewerte(s["text"], profil, alle, s.get("kapitel", ""))
        begriffe = st.fachbegriffe(s["text"])
        raus.append({**s, "bew": b, "begriffe": begriffe,
                     "kurz": st.kurzfassung(s.get("binding", ""), b,
                                            s.get("kennwerte"), begriffe)})
    return raus


def _rang(s: dict):
    return (s["bew"].thema, BINDUNG_RANG.get(str(s.get("binding") or "").upper(), 5),
            s.get("seite_min", 0), s["text"][:40])


def faehigkeitszeilen(saetze: list[dict]) -> list[dict]:
    """
    Faehigkeitsmatrix: eine Zeile je Verfahren/Kennwert/Toleranz-Kombination.
    Gleiche Kombinationen von verschiedenen Seiten werden zusammengefuehrt,
    damit "Cgk >= 1,33" nicht sechsmal untereinander steht.
    """
    gruppen: dict[tuple, dict] = {}
    for s in saetze:
        f = s.get("faehigkeit") or {}
        if not any(f.get(k) for k in ("verfahren", "kennwert", "toleranz")):
            continue
        schluessel = (st.falten(f.get("verfahren", "")), st.falten(f.get("kennwert", "")),
                      st.falten(f.get("toleranz", "")), s["bew"].geltung)
        g = gruppen.get(schluessel)
        if g is None:
            gruppen[schluessel] = g = {
                "geltung": s["bew"].geltung, "verfahren": f.get("verfahren", ""),
                "kennwert": f.get("kennwert", ""), "toleranz": f.get("toleranz", ""),
                "thema": s["bew"].thema, "kurz": s["kurz"], "wortlaut": s["text"],
                "dateien": set(), "seiten": [], "binding": s.get("binding", ""),
                "weitere": 0,
            }
        else:
            g["weitere"] += 1
        g["dateien"].add(s.get("datei", ""))
        g["seiten"] += list(s.get("seiten", []))
    reihenfolge = {"Station": 0, "Anlagenweit": 1, "Unklar": 2, "Andere Station": 3, "": 1}
    return sorted(gruppen.values(), key=lambda g: (
        reihenfolge.get(g["geltung"], 9), g["verfahren"].lower(), g["kennwert"].lower()))


def kennwertzeilen(saetze: list[dict]) -> list[dict]:
    """Parameter-Tabelle: gleiche Parameter/Werte von mehreren Seiten werden vereint."""
    gruppen: dict[tuple, dict] = {}
    for s in saetze:
        for p in st.parameter(s["text"]):
            k = (p["parameter"], p["bedingung"], p["wert"], p["einheit"], s["bew"].geltung)
            g = gruppen.get(k)
            if g is None:
                gruppen[k] = g = {**p, "geltung": s["bew"].geltung, "thema": s["bew"].thema,
                                  "sicherheit": s["bew"].sicherheit, "seiten": [],
                                  "dateien": set(), "wortlaut": s["text"]}
            g["seiten"] += list(s.get("seiten", []))
            g["dateien"].add(s.get("datei", ""))
    reihenfolge = {"Station": 0, "Anlagenweit": 1, "Unklar": 2, "Andere Station": 3, "": 1}
    return sorted(gruppen.values(), key=lambda g: (
        reihenfolge.get(g["geltung"], 9), g["parameter"], g["wert"]))


def luecken(kat: dict, treffer: list[dict], kennwerte: list[dict] | None = None) -> list[str]:
    """
    Befuellbare Zielfelder, zu denen nichts zugeordnet wurde. Steht ein
    passender Kennwert trotzdem im Text (Feldname enthaelt den Parameter),
    wird das dazugeschrieben - "nicht gefunden" waere dann falsch.
    """
    belegt = {(t["folienart"], t["feld_id"]) for t in treffer}
    raus = []
    for (fa, fid), z in sorted(kat.items()):
        if z.get("befuellbar") == "ja" and (fa, fid) not in belegt:
            feld = z.get("feldname") or fid
            eintrag = f"{fa} / {feld}"
            treffer_ = [k for k in (kennwerte or [])
                        if st.falten(k["parameter"].split(" (")[0]) in st.falten(feld + " " + fa)]
            if treffer_:
                k = treffer_[0]
                op = "" if k["bedingung"] == "=" else k["bedingung"] + " "
                eintrag += (f" – steht im Text ({k['parameter']} {op}{k['wert']} {k['einheit']}, "
                            f"S. {_seiten(k['seiten'])}), aber keinem Feld zugeordnet")
            raus.append(eintrag)
    return raus


def aufbereiten(profil: st.Profil | None, alle: list[st.Profil] | None,
                saetze_roh: list[dict], treffer: list[dict], kat: dict,
                fassungsvergleich: list[tuple] | None = None) -> dict:
    """
    Alles, was Stationsblatt (xlsx) und Stationsfolien (pptx) zeigen - einmal
    berechnet, damit beide garantiert dasselbe sagen.
    """
    saetze = bewerten(saetze_roh, profil, alle)
    kennw = kennwertzeilen(saetze)
    return {
        "profil": profil,
        "saetze": saetze,
        "zaehler": collections.Counter(s["bew"].geltung for s in saetze),
        "faehigkeiten": faehigkeitszeilen(saetze),
        "kennwerte": kennw,
        "luecken": luecken(kat, treffer, kennw),
        "quellen": sorted({Path(s.get("datei", "")).name for s in saetze if s.get("datei")}),
        "aenderungen": [(alt, neu, aend.einordnen(liste, profil, alle))
                        for alt, neu, liste in (fassungsvergleich or [])],
    }


def schreibe(ziel: Path, profil: st.Profil | None, alle: list[st.Profil] | None,
             saetze_roh: list[dict], treffer: list[dict], kat: dict,
             fassungsvergleich: list[tuple] | None = None) -> dict:
    """
    Schreibt das Stationsblatt und gibt die aufbereiteten Daten zurueck
    (fuer die Stationsfolien). Ohne Profil ("Gesamtanlage") gibt es keine
    Geltung, sondern Themen.
    """
    import xlsxwriter

    daten = aufbereiten(profil, alle, saetze_roh, treffer, kat, fassungsvergleich)
    saetze, zaehler = daten["saetze"], daten["zaehler"]
    wb = xlsxwriter.Workbook(str(ziel))

    fm = {
        "titel": wb.add_format({"bold": True, "font_size": 15}),
        "kopf": wb.add_format({"bold": True, "bg_color": "#1F3B57", "font_color": "#FFFFFF",
                               "border": 1, "text_wrap": True, "valign": "top"}),
        "zelle": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
        "fett": wb.add_format({"text_wrap": True, "valign": "top", "border": 1, "bold": True}),
        "grau": wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                               "font_color": "#777777"}),
        "gelb": wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                               "bg_color": "#FDF4E3"}),
        "gruen": wb.add_format({"text_wrap": True, "valign": "top", "border": 1,
                                "bg_color": "#E8F5EE"}),
        "text": wb.add_format({"text_wrap": True, "valign": "top"}),
        "klein": wb.add_format({"italic": True, "font_color": "#666666", "text_wrap": True,
                                "valign": "top"}),
    }

    # ------------------------------------------------------------ Uebersicht
    ws = wb.add_worksheet("Übersicht")
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 95)
    titel = f"Stationsblatt – {profil.name}" if profil else "Anforderungen der Anlage nach Thema"
    ws.write(0, 0, titel, fm["titel"])
    ws.write(1, 0, f"Stand {datetime.now():%d.%m.%Y %H:%M} · VERTRAULICH · Entwurf, nur intern",
             fm["klein"])
    dateien = sorted({s.get("datei", "") for s in saetze if s.get("datei")})
    ws.write(2, 0, "Quellen: " + (", ".join(Path(d).name for d in dateien) or "-"), fm["klein"])
    z = 4
    if profil:
        ws.write_row(z, 0, ["Einordnung", "Anzahl", "Bedeutung"], fm["kopf"])
        z += 1
        for g in ("Station", "Anlagenweit", "Unklar", "Andere Station"):
            ws.write_row(z, 0, [g, zaehler.get(g, 0), GELTUNG_TEXT[g]],
                         fm["gruen"] if g == "Station" else fm["zelle"])
            z += 1
        for alt_d, neu_d, liste in daten["aenderungen"]:
            echt = [x for x in liste if x["status"] != "unveraendert"]
            hier = sum(1 for x in echt if x["bew"].geltung == "Station")
            z += 1
            ws.merge_range(z, 0, z, 2,
                           f"Änderungen {Path(alt_d).name} → {Path(neu_d).name}: {len(echt)} insgesamt, "
                           f"davon {hier} für diese Station (Blatt 'Änderungen').", fm["fett"])
            z += 1
        z += 1
        ws.write(z, 0, "Wie lesen?", fm["fett"])
        z += 1
        for zeile in (
            "1. Blätter 'Fähigkeiten' und 'Kennwerte': Verfahren, Cgk/Cpk, Toleranzen, Taktzeit, Kräfte als Tabelle (Parameter, Bedingung, Wert, Einheit).",
            "2. Blatt 'Diese Station': nur Sätze, die diese Station betreffen. Hier beginnt die Auslegung.",
            "3. Blatt 'Anlagenweit': gilt für alle Stationen, einmal und nach Thema sortiert.",
            "4. Blatt 'Unklar': bitte kurz ansehen – hier entscheidet ein Mensch, ob es zur Station gehört.",
            "5. 'Andere Stationen' dient nur der Kontrolle: Steht hier etwas, das doch zu dieser Station gehört, bitte 'nein' in 'Stimmt?' eintragen.",
            "Die deutsche Kurzfassung ist ein Stichwortzettel aus Regeln. Maßgeblich ist immer der englische/deutsche Wortlaut mit Seite.",
        ):
            ws.merge_range(z, 0, z, 2, zeile, fm["text"])
            ws.set_row(z, 30)
            z += 1
    else:
        ws.merge_range(z, 0, z, 2,
                       "Es ist noch keine Station bekannt. Die Sätze sind deshalb nur nach Thema sortiert. "
                       "Sobald die Stationsliste vorliegt (Nummer;Name), entsteht daraus je Station ein eigenes Blatt.",
                       fm["text"])
        ws.set_row(z, 45)
        z += 2
        ws.write_row(z, 0, ["Thema", "Sätze", ""], fm["kopf"])
        z += 1
        je_thema = collections.Counter(s["bew"].thema for s in saetze)
        for thema, n in je_thema.most_common():
            ws.write_row(z, 0, [thema, n, ""], fm["zelle"])
            z += 1
    z += 1
    ws.write(z, 0, "Nicht gefunden (Lücken)", fm["fett"])
    z += 1
    fehlt = daten["luecken"]
    if fehlt:
        for eintrag in fehlt:
            ws.merge_range(z, 0, z, 2, "– " + eintrag, fm["text"])
            z += 1
        ws.merge_range(z, 0, z, 2,
                       "Diese Angaben stehen nicht im Lastenheft (oder wurden nicht erkannt). "
                       "Sie müssen aus Prozessspezifikation, Zeichnung oder Rückfrage kommen.", fm["klein"])
        ws.set_row(z, 30)
    else:
        ws.write(z, 0, "Keine – alle als befüllbar eingestuften Felder haben mindestens einen Vorschlag.",
                 fm["text"])

    # ---------------------------------------------------------- Faehigkeiten
    zeilen = daten["faehigkeiten"]
    ws = wb.add_worksheet("Fähigkeiten")
    sp = ["Einordnung", "Verfahren", "Kennwert", "Toleranz / Genauigkeit", "Thema",
          "Kurzfassung (DE)", "Datei", "Seite(n)", "Wortlaut (Beleg)", "Stimmt?"]
    br = [14, 24, 22, 28, 22, 34, 26, 10, 80, 9]
    for i, (s_, b_) in enumerate(zip(sp, br)):
        ws.write(0, i, s_, fm["kopf"])
        ws.set_column(i, i, b_)
    ws.freeze_panes(1, 0)
    for n, g in enumerate(zeilen, start=1):
        f = fm["gruen"] if g["geltung"] == "Station" else (
            fm["grau"] if g["geltung"] == "Andere Station" else fm["zelle"])
        ws.write_row(n, 0, [
            g["geltung"] or "–", g["verfahren"], g["kennwert"], g["toleranz"], g["thema"],
            g["kurz"], ", ".join(Path(d).name for d in sorted(g["dateien"]) if d),
            _seiten(g["seiten"]) + (f" (+{g['weitere']} weitere)" if g["weitere"] else ""),
            g["wortlaut"], ""], f)
    ws.autofilter(0, 0, max(len(zeilen), 1), len(sp) - 1)
    if not zeilen:
        ws.write(1, 0, "Keine Fähigkeitsangaben gefunden.", fm["text"])
    else:
        ws.data_validation(1, len(sp) - 1, len(zeilen), len(sp) - 1,
                           {"validate": "list", "source": ["ja", "nein"]})

    # ------------------------------------------------------------ Kennwerte
    kz = daten["kennwerte"]
    ws = wb.add_worksheet("Kennwerte")
    sp = ["Einordnung", "Parameter", "Bedingung", "Wert", "Einheit", "Thema", "Datei",
          "Seite(n)", "Wortlaut (Beleg)", "Stimmt?"]
    br = [14, 28, 10, 10, 9, 22, 26, 10, 80, 9]
    for i, (s_, b_) in enumerate(zip(sp, br)):
        ws.write(0, i, s_, fm["kopf"])
        ws.set_column(i, i, b_)
    ws.freeze_panes(1, 0)
    for n, g in enumerate(kz, start=1):
        f = fm["gruen"] if g["geltung"] == "Station" else (
            fm["grau"] if g["geltung"] == "Andere Station" else fm["zelle"])
        ws.write_row(n, 0, [g["geltung"] or "–", g["parameter"], g["bedingung"], g["wert"],
                            g["einheit"], g["thema"],
                            ", ".join(Path(d).name for d in sorted(g["dateien"]) if d),
                            _seiten(g["seiten"]), g["wortlaut"], ""], f)
    ws.autofilter(0, 0, max(len(kz), 1), len(sp) - 1)
    if kz:
        ws.data_validation(1, len(sp) - 1, len(kz), len(sp) - 1,
                           {"validate": "list", "source": ["ja", "nein"]})
    else:
        ws.write(1, 0, "Keine Zahlenwerte mit bekanntem Parameter gefunden.", fm["text"])

    # ------------------------------------------------------------ Aenderungen
    for n, (alt_d, neu_d, liste) in enumerate(daten["aenderungen"], start=1):
        ws = wb.add_worksheet("Änderungen" if n == 1 else f"Änderungen {n}")
        aend.blatt_schreiben(ws, liste, fm)

    # -------------------------------------------------- Satzlisten je Einordnung
    def satzblatt(name: str, auswahl: list[dict], farbe: str):
        ws_ = wb.add_worksheet(name)
        sp_ = ["Nr", "Thema", "Sicherheit", "Verb.", "Kurzfassung (DE)", "Kennwerte", "Fachbegriffe (DE)",
               "Wortlaut (Beleg)", "Datei", "Seite(n)", "Kapitel", "Warum diese Einordnung",
               "Feld in der Akte", "Stimmt?"]
        br_ = [5, 22, 11, 8, 40, 30, 28, 80, 26, 10, 30, 34, 28, 9]
        for i, (s_, b_) in enumerate(zip(sp_, br_)):
            ws_.write(0, i, s_, fm["kopf"])
            ws_.set_column(i, i, b_)
        ws_.freeze_panes(1, 2)
        for n, s in enumerate(sorted(auswahl, key=_rang), start=1):
            b = s["bew"]
            warum = b.warum + (f" → {b.andere}" if b.andere and b.andere not in b.warum else "")
            ws_.write_row(n, 0, [
                n, b.thema, b.sicherheit or "-", s.get("binding", ""), s["kurz"], "; ".join(s.get("kennwerte") or []),
                ", ".join(s["begriffe"]), s["text"], Path(s.get("datei", "")).name,
                _seiten(s.get("seiten", [])), s.get("kapitel", ""), warum,
                "; ".join(s.get("felder", [])), ""],
                fm[farbe])
        ws_.autofilter(0, 0, max(len(auswahl), 1), len(sp_) - 1)
        if auswahl:
            ws_.data_validation(1, len(sp_) - 1, len(auswahl), len(sp_) - 1,
                                {"validate": "list", "source": ["ja", "nein"]})
        else:
            ws_.write(1, 1, "Keine Einträge.", fm["text"])

    if profil:
        je = collections.defaultdict(list)
        for s in saetze:
            je[s["bew"].geltung].append(s)
        satzblatt("Diese Station", je["Station"], "gruen")
        satzblatt("Anlagenweit", je["Anlagenweit"], "zelle")
        satzblatt("Unklar", je["Unklar"], "gelb")
        satzblatt("Andere Stationen", je["Andere Station"], "grau")
    else:
        satzblatt("Nach Thema", saetze, "zelle")
        # Stationsvorschlag
        vs = st.vorschlag([s["text"] for s in saetze])
        ws = wb.add_worksheet("Stationsvorschlag")
        for i, (s_, b_) in enumerate(zip(["Nummer", "Name (Vorschlag)", "Sätze im Lastenheft",
                                          "Stichworte", "Bitte ergänzen"], [9, 36, 18, 36, 40])):
            ws.write(0, i, s_, fm["kopf"])
            ws.set_column(i, i, b_)
        for n, v in enumerate(vs, start=1):
            ws.write_row(n, 0, ["", v["name"], v["saetze"], "", ""], fm["zelle"])
        ws.write(len(vs) + 2, 1,
                 "Diese Liste zählt nur, welche Verfahren im Lastenheft vorkommen. Welche Station "
                 "was macht und wie sie nummeriert ist, entscheidet der Konstrukteur (Spalte Nummer).",
                 fm["klein"])
    wb.close()
    return daten


def vorschlag_csv(ziel: Path, saetze: list[dict]) -> int:
    """Schreibt die vorgeschlagene Stationsliste als Vorlage zum Ausfuellen."""
    vs = st.vorschlag([s["text"] for s in saetze])
    with ziel.open("w", encoding="utf-8-sig", newline="") as f:
        f.write("Nummer;Name;Stichworte;Saetze_im_Lastenheft\n")
        for v in vs:
            f.write(f";{v['name']};;{v['saetze']}\n")
    return len(vs)
