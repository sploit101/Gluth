#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
uebersicht_bauen.py — macht aus zielfelder.csv eine lesbare Uebersicht.

Erzeugt eine einzelne HTML-Datei, die man im Browser oeffnet. Darin steht
jede der 21 Folienarten als eigene Karte mit allen ihren Feldern, farblich
getrennt nach: automatisch befuellbar / teilweise / nicht befuellbar.

Zweck: man sieht auf einen Blick, was das Werkzeug fuellen kann und was
nicht - ohne PowerPoint und ohne sich durch 176 Excel-Zeilen zu scrollen.

Keine Internetverbindung, kein externes Skript, alles in einer Datei.
Die Datei enthaelt nur Feldbezeichnungen aus der Vorlage, keine
Kundeninhalte - sie bleibt aber trotzdem im Projektordner.

    python uebersicht_bauen.py
    python uebersicht_bauen.py --csv zielfelder.csv --aus uebersicht.html
"""

from __future__ import annotations

import argparse
import collections
import csv
import html
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent

FARBE = {
    "ja":        ("#1a7f4b", "#e8f5ee", "fuellbar"),
    "teilweise": ("#9a6410", "#fdf4e3", "teilweise"),
    "nein":      ("#6b6b6b", "#f1f1f1", "nicht fuellbar"),
}

# Reihenfolge der Karten: erst was wir bauen, dann der Rest.
def kartenrang(gruppe: list[dict]) -> tuple:
    ja = sum(1 for z in gruppe if z["befuellbar"] == "ja")
    tw = sum(1 for z in gruppe if z["befuellbar"] == "teilweise")
    im_beispiel = max(int(z.get("im_beispiel") or 0) for z in gruppe)
    return (-ja, -tw, -im_beispiel, gruppe[0]["folienart"].lower())


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", default=str(HERE / "zielfelder.csv"))
    ap.add_argument("--aus", default=str(HERE / "uebersicht_folienarten.html"))
    a = ap.parse_args()

    pfad = Path(a.csv)
    if not pfad.exists():
        print("FEHLER: nicht gefunden:", pfad)
        print("Erst 'python zielfelder_bauen.py' laufen lassen.")
        return 2

    with pfad.open(encoding="utf-8-sig", newline="") as f:
        zeilen = list(csv.DictReader(f, delimiter=";"))

    gruppen: dict[str, list[dict]] = collections.defaultdict(list)
    for z in zeilen:
        gruppen[z["folienart"]].append(z)

    bf = collections.Counter(z["befuellbar"] for z in zeilen)
    n = len(zeilen)
    fragen = []
    for z in zeilen:
        if z["frage"] and z["frage"] not in [q[0] for q in fragen]:
            fragen.append((z["frage"], z["folienart"], z["feldname"]))

    karten = sorted(gruppen.values(), key=kartenrang)

    T = []
    A = T.append
    A("<!doctype html><html lang='de'><head><meta charset='utf-8'>")
    A("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    A("<title>Stationsakte - Folienarten und Felder</title><style>")
    A("""
:root{--bg:#ffffff;--fg:#1b1b1b;--mut:#666;--lin:#e2e2e2;--kar:#fafafa;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);
 font:15px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;}
.wrap{max-width:1180px;margin:0 auto;padding:32px 16px 64px}
h1{font-size:25px;margin:0 0 4px;font-weight:650}
.sub{color:var(--mut);font-size:14px;margin:0 0 28px}
.zahlen{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 26px}
.z{border:1px solid var(--lin);border-radius:8px;padding:11px 15px;min-width:130px}
.z b{display:block;font-size:22px;font-weight:650;line-height:1.2}
.z span{font-size:12.5px;color:var(--mut)}
.legende{display:flex;flex-wrap:wrap;gap:16px;margin:0 0 30px;
 padding:13px 15px;border:1px solid var(--lin);border-radius:8px;font-size:13.5px}
.leg{display:flex;align-items:center;gap:7px}
.pkt{width:11px;height:11px;border-radius:50%;flex:none}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:16px}
.karte{border:1px solid var(--lin);border-radius:10px;overflow:hidden;
 background:var(--kar);display:flex;flex-direction:column}
.kopf{padding:12px 15px;border-bottom:1px solid var(--lin);background:#fff}
.kopf h2{font-size:15px;margin:0;font-weight:650;letter-spacing:-.1px}
.marken{display:flex;flex-wrap:wrap;gap:6px;margin-top:7px}
.marke{font-size:11px;padding:2px 7px;border-radius:20px;border:1px solid var(--lin);
 color:var(--mut);background:#fff}
.felder{list-style:none;margin:0;padding:6px 0;flex:1}
.felder li{display:flex;gap:9px;padding:6px 15px;align-items:flex-start;
 font-size:13.5px;border-top:1px solid #efefef}
.felder li:first-child{border-top:none}
.felder .pkt{margin-top:5px}
.fname{flex:1}
.qt{font-size:11px;color:var(--mut);white-space:nowrap;padding-top:1px}
.frage{font-size:11.5px;color:#9a6410;display:block;margin-top:2px}
h3{font-size:17px;margin:40px 0 12px;font-weight:650}
table{border-collapse:collapse;width:100%;font-size:13.5px}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--lin);
 vertical-align:top}
th{font-weight:600;color:var(--mut);font-size:12.5px}
.hinweis{margin-top:36px;padding:14px 16px;border:1px solid var(--lin);
 border-radius:8px;font-size:13.5px;color:#444}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
 --bg:#141414;--fg:#ececec;--mut:#9a9a9a;--lin:#2e2e2e;--kar:#1b1b1b;}
 :root:not([data-theme="light"]) .kopf{background:#191919}
 :root:not([data-theme="light"]) .marke{background:#1b1b1b}
 :root:not([data-theme="light"]) .felder li{border-top-color:#242424}
 :root:not([data-theme="light"]) .frage{color:#d8a55a}}
@media print{.karte{break-inside:avoid}body{background:#fff}}
    """)
    A("</style></head><body><div class='wrap'>")

    A("<h1>Stationsakte &ndash; Folienarten und Felder</h1>")
    A(f"<p class='sub'>Erzeugt aus der Vorlage von Gluth, "
      f"{datetime.now():%d.%m.%Y %H:%M}. Jede Karte ist eine Folienart aus dem "
      f"Folienmaster. Jede Zeile darin ist ein Platzhalter mit dem Hinweistext, "
      f"den Gluth selbst hineingeschrieben hat.</p>")

    # Karten, die aus einem echten Folienlayout stammen, von den
    # nachgetragenen Tabellen und Grafikfeldern trennen.
    echte = sum(1 for g in karten
                if any(z["feld_id"].startswith("ph") or z["feld_id"] == "-"
                       for z in g))
    zusatz = len(karten) - echte

    A("<div class='zahlen'>")
    A(f"<div class='z'><b>{echte}</b><span>Folienarten im Master</span></div>")
    if zusatz:
        A(f"<div class='z'><b>{zusatz}</b><span>Tabellen und Grafiken "
          f"nachgetragen</span></div>")
    A(f"<div class='z'><b>{n}</b><span>Felder insgesamt</span></div>")
    for k in ("ja", "teilweise", "nein"):
        c = bf.get(k, 0)
        A(f"<div class='z'><b>{c}</b><span>{FARBE[k][2]} &middot; "
          f"{100*c/n:.0f}&nbsp;%</span></div>")
    A("</div>")

    A("<div class='legende'>")
    for k, (fg, _bgc, txt) in FARBE.items():
        A(f"<div class='leg'><span class='pkt' style='background:{fg}'></span>"
          f"<span><b>{txt}</b> &ndash; " + {
              "ja": "Inhalt steht im Lastenheft oder in der Spezifikation",
              "teilweise": "kann dort stehen, muss aber nicht",
              "nein": "Konstruktion, Kalkulation, Einkauf oder Bild",
          }[k] + "</span></div>")
    A("</div>")

    A("<div class='grid'>")
    for g in karten:
        name = g[0]["folienart"]
        ja = sum(1 for z in g if z["befuellbar"] == "ja")
        imb = max(int(z.get("im_beispiel") or 0) for z in g)
        A("<div class='karte'><div class='kopf'>")
        A(f"<h2>{html.escape(name)}</h2><div class='marken'>")
        A(f"<span class='marke'>{len(g)} Felder</span>")
        if ja:
            A(f"<span class='marke' style='color:{FARBE['ja'][0]};"
              f"border-color:{FARBE['ja'][0]}'>{ja}&times; f&uuml;llbar</span>")
        A("<span class='marke'>" +
          (f"im Beispiel {imb}&times;" if imb else "nicht im Beispiel") +
          "</span>")
        A("</div></div><ul class='felder'>")
        rang = {"ja": 0, "teilweise": 1, "nein": 2}
        for z in sorted(g, key=lambda r: (rang.get(r["befuellbar"], 3),
                                          r["feld_id"])):
            fg = FARBE.get(z["befuellbar"], FARBE["nein"])[0]
            A(f"<li><span class='pkt' style='background:{fg}'></span>")
            A(f"<span class='fname'>{html.escape(z['feldname'])}")
            if z["frage"]:
                A("<span class='frage'>offene Frage an Daniel</span>")
            A("</span>")
            A(f"<span class='qt'>{html.escape(z['quelle'])}</span></li>")
        A("</ul></div>")
    A("</div>")

    A("<h3>Die sechs offenen Fragen</h3>")
    A("<table><tr><th>Frage</th><th>Folienart</th><th>Feld</th></tr>")
    for q, fa, fn in fragen:
        A(f"<tr><td>{html.escape(q)}</td><td>{html.escape(fa)}</td>"
          f"<td>{html.escape(fn)}</td></tr>")
    A("</table>")

    A("<div class='hinweis'><b>Wie das zu lesen ist.</b> "
      "Gr&uuml;n hei&szlig;t: diesen Platz kann das Programm aus einem "
      "Kundendokument f&uuml;llen. Grau hei&szlig;t: hier steht etwas, das "
      "kein Kundendokument enth&auml;lt &ndash; Stunden, Budget, Hersteller, "
      "Drehzahlen. Das ist Gluths eigene Arbeit.<br><br>"
      "Die Einteilung ist eine <b>begr&uuml;ndete Einsch&auml;tzung</b>, "
      "abgeleitet aus den Hinweistexten der Vorlage, nicht aus einer "
      "ausgef&uuml;llten Akte. Sie geh&ouml;rt gegengelesen.<br><br>"
      "<b>Achtung bei den Prozentzahlen:</b> hier z&auml;hlt jedes Feld "
      "gleich. &bdquo;U/min&ldquo; und &bdquo;Beschreibung der Aufgabe "
      "ausf&uuml;hrlich&ldquo; sind je ein Feld, obwohl im zweiten zw&ouml;lf "
      "S&auml;tze stehen. Nach Textmenge gerechnet ist der f&uuml;llbare "
      "Anteil deutlich gr&ouml;&szlig;er.</div>")

    A("</div></body></html>")

    ziel = Path(a.aus)
    ziel.write_text("\n".join(T), encoding="utf-8")
    print("geschrieben:", ziel.name)
    print(f"  {len(karten)} Folienarten, {n} Felder")
    print(f"  fuellbar {bf.get('ja',0)} / teilweise {bf.get('teilweise',0)}"
          f" / nicht {bf.get('nein',0)}")
    print("  Im Browser oeffnen (Doppelklick).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
