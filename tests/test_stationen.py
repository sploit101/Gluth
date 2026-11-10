"""Rules that decide whether a sentence belongs to a station (no model, no network)."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import stationen as st

LISTE = [st.profil_aus_name(n) for n in
         ("Station 10 Zuführung", "Station 20 Abisolieren", "Station 30 Crimpen",
          "Station 40 Einpressen")]
S20 = LISTE[1]


class BewertungTests(unittest.TestCase):
    def check(self, text, soll, profil=S20):
        ist = st.bewerte(text, profil, LISTE)
        self.assertEqual(ist.geltung, soll, f"{text} -> {ist.geltung} ({ist.warum})")
        return ist

    def test_station_by_process_english_and_german(self):
        self.check("The stripping length must be held within +/-0.2 mm.", "Station")
        self.check("Die Abisolierlänge ist mit ±0,2 mm einzuhalten.", "Station")
        self.check("Die Abisolierlaenge ist mit ±0,2 mm einzuhalten.", "Station")

    def test_station_number_wins(self):
        self.check("Station 20 must reject parts with damaged insulation.", "Station")
        b = self.check("Station 30 needs a crimp force monitor.", "Andere Station")
        self.assertIn("30", b.warum)

    def test_other_station_by_process(self):
        b = self.check("The crimp height shall be monitored.", "Andere Station")
        self.assertIn("Station 30", b.andere)

    def test_plant_wide(self):
        self.check("All stations must reach Cgk >= 1.33.", "Anlagenweit")
        self.check("The machine shall comply with the Machinery Directive.", "Anlagenweit")
        self.check("Operating instructions must be delivered in German.", "Anlagenweit")

    def test_unclear_is_not_guessed(self):
        self.check("The customer may change the colour at short notice.", "Unklar")

    def test_no_real_station_means_no_rating(self):
        self.assertIsNone(st.profil_aus_name("Gesamtanlage (Stationszuordnung folgt)"))
        self.assertIsNone(st.profil_aus_name(""))
        b = st.bewerte("The crimp height shall be monitored.", None)
        self.assertEqual(b.geltung, "")
        self.assertIn("Crimpen", b.themen)

    def test_custom_station_name_uses_own_keyword(self):
        p = st.profil_aus_name("Station 60 Gehäusemontage")
        b = st.bewerte("Die Gehäusemontage erfolgt halbautomatisch.", p)
        self.assertEqual(b.geltung, "Station")

    def test_compact_values_only_from_text(self):
        k = st.kompakt("Cgk >= 1.33 and Cpk>=1.67, tolerance +/- 0.2 mm")
        self.assertEqual(k["kennwert"], ["Cgk ≥ 1,33", "Cpk ≥ 1,67"])
        self.assertEqual(k["toleranz"], ["± 0,2 mm"])
        self.assertEqual(st.kompakt("The machine must be safe."),
                         {"kennwert": [], "toleranz": []})

    def test_station_list_csv(self):
        with tempfile.TemporaryDirectory() as d:
            f = Path(d) / "s.csv"
            f.write_text("Nummer;Name;Stichworte\n20;Abisolieren;\n30;Crimpen;Presse\n",
                         encoding="utf-8")
            profile = st.stationsliste_laden(f)
        self.assertEqual([p.nummer for p in profile], ["20", "30"])
        self.assertIn("CRIMPEN", profile[1].prozesse)

    def test_proposal_counts_processes(self):
        v = st.vorschlag(["Crimp height.", "crimping force", "Camera inspection."])
        self.assertEqual(v[0]["name"], "Crimpen")
        self.assertEqual(v[0]["saetze"], 2)


if __name__ == "__main__":
    unittest.main()
