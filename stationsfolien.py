#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stationsfolien.py — die Stationsuebersicht als PowerPoint, wie im Zielbild.

Bisher war das Zielbild ("Station 20", "Prozessanforderungen als Tabelle",
"Gilt fuer alle Stationen") von Hand gebaut. Diese Datei erzeugt genau das
aus denselben Daten wie das Stationsblatt:

    1  Titel        Station, Quellen, Zahlen, Farbschluessel
    2  Prozessanforderungen   Parameter | Anforderung | gilt fuer | Seite
    3  Anforderungen dieser Station   Thema | Kurzfassung | Auszug | Seite
    4  Aenderungen ggue. Vorfassung   (nur wenn eine aeltere Fassung vorliegt)
    5  Gilt fuer alle Stationen        nach Thema
    6  Offen / zu klaeren              Luecken, unklare Saetze

Farbschluessel wie in der Stationsakte: blau = automatisch aus dem Dokument,
grau = Quelle, braun = nicht gefunden. Jede Zeile nennt ihre Seite. Die
Folien sind ein Entwurf zum Besprechen, keine Freigabe - massgeblich bleibt
der Wortlaut im Stationsblatt.

Wird von zuordnen.py aufgerufen; kein eigener Aufruf noetig.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt

BLAU = RGBColor(0x1A, 0x5C, 0x8A)
GRAU = RGBColor(0x77, 0x77, 0x77)
BRAUN = RGBColor(0x9A, 0x64, 0x10)
DUNKEL = RGBColor(0x1F, 0x3B, 0x57)
WEISS = RGBColor(0xFF, 0xFF, 0xFF)
HELL = RGBColor(0xF3, 0xF6, 0xF8)
GRUEN = RGBColor(0xE8, 0xF5, 0xEE)

BREITE, HOEHE = Cm(33.867), Cm(19.05)          # 16:9


def _kurz(text: str, n: int) -> str:
    text = " ".join(str(text or "").split())
    return text if len(text) <= n else text[:n - 1].rstrip() + "…"


def _seiten(seiten, n: int = 4) -> str:
    teile = sorted({str(s) for s in seiten if s not in (None, "")},
                   key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else 0, x))
    return ", ".join(teile[:n]) + (" …" if len(teile) > n else "")


class Deck:
    def __init__(self, fusszeile: str):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = BREITE, HOEHE
        self.leer = self.prs.slide_layouts[6]
        self.fusszeile = fusszeile

    def text(self, folie, x, y, b, h, inhalt, groesse=12, farbe=None, fett=False):
        box = folie.shapes.add_textbox(x, y, b, h)
        tf = box.text_frame
        tf.word_wrap = True
        zeilen = inhalt if isinstance(inhalt, list) else [inhalt]
        for i, z in enumerate(zeilen):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            if isinstance(z, tuple):
                z, f_, g_ = (z + (None, None))[:3]
            else:
                f_, g_ = None, None
            r = p.add_run()
            r.text = z
            r.font.size = Pt(g_ or groesse)
            r.font.bold = fett
            if f_ or farbe:
                r.font.color.rgb = f_ or farbe
        return box

    def folie(self, titel: str, untertitel: str = ""):
        f = self.prs.slides.add_slide(self.leer)
        balken = f.shapes.add_shape(1, 0, 0, BREITE, Cm(2.3))
        balken.fill.solid()
        balken.fill.fore_color.rgb = DUNKEL
        balken.line.fill.background()
        self.text(f, Cm(1.2), Cm(0.45), Cm(28), Cm(1.4), titel, 24, WEISS, True)
        if untertitel:
            self.text(f, Cm(1.2), Cm(2.5), Cm(31), Cm(0.9), untertitel, 11, GRAU)
        self.text(f, Cm(1.2), Cm(18.1), Cm(31), Cm(0.8), self.fusszeile, 8, GRAU)
        return f

    def tabelle(self, folie, kopf: list[str], zeilen: list[list], breiten_cm: list[float],
                y=Cm(3.4), groesse=10, hervorheben=None):
        t = folie.shapes.add_table(len(zeilen) + 1, len(kopf), Cm(1.2), y,
                                   Cm(sum(breiten_cm)), Cm(0.8 * (len(zeilen) + 1))).table
        for i, b in enumerate(breiten_cm):
            t.columns[i].width = Cm(b)
        for i, k in enumerate(kopf):
            c = t.cell(0, i)
            c.text = k
            c.fill.solid()
            c.fill.fore_color.rgb = DUNKEL
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(groesse)
                    r.font.bold = True
                    r.font.color.rgb = WEISS
        for zi, zeile in enumerate(zeilen, start=1):
            gruen = hervorheben(zeile) if hervorheben else False
            for si, wert in enumerate(zeile):
                c = t.cell(zi, si)
                farbe = None
                if isinstance(wert, tuple):
                    wert, farbe = wert
                c.text = str(wert)
                c.fill.solid()
                c.fill.fore_color.rgb = GRUEN if gruen else (HELL if zi % 2 else WEISS)
                for p in c.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(groesse)
                        r.font.color.rgb = farbe or BLAU
        return t

    def seitenweise(self, titel, untertitel, kopf, zeilen, breiten, je_folie, **kw):
        if not zeilen:
            f = self.folie(titel, untertitel)
            self.text(f, Cm(1.2), Cm(4), Cm(31), Cm(1.5), "Nichts gefunden.", 14, BRAUN)
            return
        teile = [zeilen[i:i + je_folie] for i in range(0, len(zeilen), je_folie)]
        for n, teil in enumerate(teile, start=1):
            t_ = titel if len(teile) == 1 else f"{titel} ({n}/{len(teile)})"
            self.tabelle(self.folie(t_, untertitel), kopf, teil, breiten, **kw)

    def speichern(self, ziel: Path):
        self.prs.save(str(ziel))


GILT = {"Station": "diese Station", "Anlagenweit": "alle Stationen", "Unklar": "unklar",
        "Andere Station": "andere Station", "": "–"}


def schreibe(ziel: Path, daten: dict) -> Path:
    profil = daten["profil"]
    saetze = daten["saetze"]
    z = daten["zaehler"]
    quellen = ", ".join(daten["quellen"]) or "–"
    name = profil.name if profil else "Gesamtanlage"
    deck = Deck(f"Entwurf · nur intern · automatisch zusammengestellt am {datetime.now():%d.%m.%Y} "
                f"aus {quellen} · maßgeblich ist der Wortlaut mit Seite (Stationsblatt)")

    # 1 Titel ------------------------------------------------------------
    f = deck.folie(name, "Anforderungsübersicht aus den Kundendokumenten – Entwurf zum Besprechen")
    zeilen = [("Quellen: " + quellen, GRAU, 12), ("", None, 8)]
    if profil:
        zeilen += [(f"{z.get('Station', 0)} Anforderungen für diese Station", BLAU, 20),
                   (f"{z.get('Anlagenweit', 0)} gelten für alle Stationen", BLAU, 16),
                   (f"{z.get('Unklar', 0)} ohne erkennbaren Stationsbezug – bitte zuordnen", BRAUN, 16),
                   (f"{z.get('Andere Station', 0)} gehören zu anderen Stationen (nicht auf diesen Folien)",
                    GRAU, 14)]
    else:
        zeilen += [(f"{len(saetze)} Anforderungen, nach Thema sortiert", BLAU, 20),
                   ("Noch keine Station bekannt – Vorschlag der Stationsfolge auf der letzten Folie.",
                    BRAUN, 14)]
    for alt_d, neu_d, liste in daten["aenderungen"]:
        echt = [x for x in liste if x["status"] != "unveraendert"]
        hier = sum(1 for x in echt if x["bew"].geltung == "Station")
        zeilen.append((f"{len(echt)} Änderungen ggü. {Path(alt_d).name}"
                       + (f", davon {hier} für diese Station" if profil else ""), BLAU, 16))
    zeilen += [("", None, 8), ("Farben: blau = automatisch aus dem Dokument · grau = Quelle · "
                               "braun = nicht gefunden / offen", GRAU, 11)]
    deck.text(f, Cm(1.2), Cm(4), Cm(31), Cm(12), zeilen)

    relevant = ("Station", "Anlagenweit", "") if profil else ("",)

    # 2 Prozessanforderungen --------------------------------------------
    import stationen as st
    prozess = []
    mit_parameter = {g["wortlaut"] for g in daten["kennwerte"]}
    for g in daten["faehigkeiten"]:
        if g["geltung"] not in relevant:
            continue
        # Nur eine Toleranz ohne Verfahren/Kennwert steht schon als Kennwert da
        if not g["verfahren"] and not g["kennwert"] and g["wortlaut"] in mit_parameter:
            continue
        verfahren_de = ", ".join(b for b in st.fachbegriffe(g["verfahren"])
                                 if "(" in b) if g["verfahren"] else ""
        g = {**g, "verfahren": verfahren_de or g["verfahren"]}
        anforderung = " · ".join(x for x in (g["kennwert"], g["toleranz"]) if x) or "gefordert"
        prozess.append([_kurz(g["verfahren"] or g["thema"], 40), _kurz(anforderung, 60),
                        GILT.get(g["geltung"], "–"), (_seiten(g["seiten"]), GRAU)])
    for g in daten["kennwerte"]:
        if g["geltung"] not in relevant:
            continue
        op = "" if g["bedingung"] == "=" else g["bedingung"] + " "
        prozess.append([g["parameter"], f"{op}{g['wert']} {g['einheit']}",
                        GILT.get(g["geltung"], "–"), (_seiten(g["seiten"]), GRAU)])
    deck.seitenweise("Prozessanforderungen", "Kennwerte, Fähigkeiten und Toleranzen – je Zeile eine Fundstelle",
                     ["Parameter / Verfahren", "Anforderung", "gilt für", "Seite"], prozess,
                     [11, 13, 5.5, 2.5], 10, groesse=12,
                     hervorheben=lambda r: r[2] == "diese Station")

    # 3 Diese Station / nach Thema --------------------------------------
    def satzzeilen(auswahl):
        auswahl = sorted(auswahl, key=lambda s: (s["bew"].thema, s.get("seite_min", 0)))
        return [[_kurz(s["bew"].thema, 26), _kurz(s["kurz"], 70), (_kurz(s["text"], 150), GRAU),
                 (_seiten(s.get("seiten", [])), GRAU)] for s in auswahl]

    if profil:
        deck.seitenweise(f"Anforderungen {name}", "Nur Sätze mit Bezug zu dieser Station "
                         "(Stationsnummer, Verfahren oder Kapitel)",
                         ["Thema", "Kurzfassung", "Originalauszug", "Seite"],
                         satzzeilen([s for s in saetze if s["bew"].geltung == "Station"]),
                         [5, 9.5, 14.5, 2.5], 6, groesse=12)
    else:
        deck.seitenweise("Anforderungen nach Thema", "Noch ohne Stationszuordnung",
                         ["Thema", "Kurzfassung", "Originalauszug", "Seite"],
                         satzzeilen(saetze), [5, 9.5, 14.5, 2.5], 6, groesse=10)

    # 4 Aenderungen ------------------------------------------------------
    for alt_d, neu_d, liste in daten["aenderungen"]:
        zeilen_ = []
        for x in liste:
            if x["status"] == "unveraendert" or (profil and x["bew"].geltung not in relevant):
                continue
            seite = (x["neu"] or x["alt"]).get("page", "")
            zeilen_.append([x["status"].replace("GEAENDERT", "GEÄNDERT"), GILT.get(x["bew"].geltung, "–"),
                            _kurz(x["kurz"] or x["bew"].thema, 80),
                            (_kurz(x["wortdiff"] or (x["neu"] or x["alt"])["text"], 110), GRAU),
                            (str(seite), GRAU)])
        deck.seitenweise(f"Änderungen ggü. {Path(alt_d).name}",
                         f"Neue Fassung: {Path(neu_d).name} – nur was diese Station oder alle Stationen betrifft",
                         ["Status", "gilt für", "Änderung", "Wortunterschied", "Seite"], zeilen_,
                         [5, 4, 10, 10.5, 2], 7, groesse=10,
                         hervorheben=lambda r: r[1] == "diese Station")

    # 5 Gilt fuer alle Stationen ----------------------------------------
    if profil:
        deck.seitenweise("Gilt für alle Stationen", "Sicherheit, Fähigkeit, Dokumentation, Abnahme, "
                         "Elektrik … – einmal, nach Thema",
                         ["Thema", "Kurzfassung", "Originalauszug", "Seite"],
                         satzzeilen([s for s in saetze if s["bew"].geltung == "Anlagenweit"]),
                         [5, 9.5, 14.5, 2.5], 7, groesse=10)

    # 6 Offen ------------------------------------------------------------
    f = deck.folie("Offen / zu klären", "Was in den Dokumenten nicht gefunden wurde oder "
                   "keiner Station zugeordnet werden konnte")
    zeilen = []
    if daten["luecken"]:
        zeilen.append(("Nicht gefunden – muss aus Prozessspezifikation, Zeichnung oder Rückfrage kommen:",
                       DUNKEL, 14))
        zeilen += [("–  " + l_, BRAUN, 12) for l_ in daten["luecken"][:14]]
        if len(daten["luecken"]) > 14:
            zeilen.append((f"… und {len(daten['luecken']) - 14} weitere (Stationsblatt, Übersicht)",
                           BRAUN, 12))
    if profil and z.get("Unklar"):
        zeilen += [("", None, 8), (f"{z['Unklar']} Sätze ohne erkennbaren Stationsbezug – im Stationsblatt, "
                                   "Blatt 'Unklar', kurz ansehen.", BRAUN, 14)]
    if not profil:
        vs = st.vorschlag([s["text"] for s in saetze])
        zeilen += [("", None, 8), ("Verfahren im Lastenheft – Vorschlag für die Stationsfolge:", DUNKEL, 14)]
        zeilen += [(f"–  {v['name']}  ({v['saetze']} Sätze)", BLAU, 12) for v in vs[:12]]
        zeilen.append(("Nummern und Namen trägt der Konstrukteur ein (stationsliste_vorschlag.csv).",
                       GRAU, 11))
    if not zeilen:
        zeilen = [("Keine offenen Punkte erkannt.", BLAU, 14)]
    deck.text(f, Cm(1.2), Cm(3.6), Cm(31), Cm(14), zeilen)

    deck.speichern(ziel)
    return ziel
