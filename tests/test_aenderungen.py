"""Version comparison and chapter detection (synthetic text, no customer data)."""
from __future__ import annotations

import unittest

import aenderungen as ae
import requirements as rq
import stationen as st


def req(text, binding="MUSS", section="", page=1):
    return {"text": text, "binding": binding, "section": section, "page": page}


class FassungTests(unittest.TestCase):
    def test_family_and_version(self):
        self.assertEqual(ae.familie("LH_04K_000051_V1.02.pdf"), ("lh_04k_000051", (1, 2)))
        self.assertEqual(ae.familie("Spec Rev 3.docx"), ("spec", (3,)))
        self.assertEqual(ae.familie("Y152058_LH_008.pdf"), ("y152058_lh", (8,)))
        dateien = ["LH_V1.01.pdf", "LH_V1.02.pdf", "LH_V1.10.pdf", "Other.pdf"]
        self.assertEqual(ae.vorgaenger("LH_V1.10.pdf", dateien), "LH_V1.02.pdf")
        self.assertEqual(ae.fassungspaare(dateien),
                         [("LH_V1.01.pdf", "LH_V1.02.pdf"), ("LH_V1.02.pdf", "LH_V1.10.pdf")])

    def test_diff_statuses(self):
        alt = [req("The stripping length must be held within +/- 0.3 mm."),
               req("The blade should be monitored for wear.", "SOLL"),
               req("A manual test bench is to be provided."),
               req("Operating instructions must be delivered in German.")]
        neu = [req("The stripping length must be held within +/- 0.2 mm."),
               req("The blade must be monitored for wear.", "MUSS"),
               req("A camera shall check every stripped wire end."),
               req("• Operating instructions must be delivered in German.")]
        s = {x["status"]: x for x in ae.vergleiche(alt, neu)}
        self.assertEqual(set(s), {"WERT GEAENDERT", "VERBINDLICHKEIT GEAENDERT", "NEU",
                                  "ENTFALLEN", "unveraendert"})
        self.assertIn("± 0,3 mm → ± 0,2 mm", s["WERT GEAENDERT"]["detail"])
        self.assertIn("{+must+}", s["VERBINDLICHKEIT GEAENDERT"]["wortdiff"])

    def test_changes_rated_for_station(self):
        p = st.profil_aus_name("Station 20 Abisolieren")
        liste = ae.einordnen(ae.vergleiche(
            [req("The crimp height must be 1.2 mm.")],
            [req("The crimp height must be 1.1 mm."),
             req("The stripping length must be 5 mm.")]), p, None)
        self.assertEqual(liste[0]["bew"].geltung, "Station")       # Station zuerst
        self.assertEqual(liste[0]["status"], "NEU")


class KapitelTests(unittest.TestCase):
    def test_headings_and_path(self):
        text = ("Lastenheft LH_04K\n3 Process requirements\n3.2 Stripping unit\n"
                "The tolerance shall be +/- 0.2 mm and has a length of\n12 Millimeters overall.\n"
                "3.4 Press-in\nForces must be monitored.\n3.2 Stripping ........ 12\n6. Safety\nGuards.")
        teile = rq.abschnitte(text, "")
        kapitel = [k for k, _ in teile]
        self.assertIn("3 Process requirements › 3.2 Stripping unit", kapitel)
        self.assertIn("3 Process requirements › 3.4 Press-in", kapitel)
        self.assertIn("6 Safety", kapitel)
        self.assertFalse(any("12 Millimeters" in k for k in kapitel))   # umbrochener Satz
        self.assertFalse(any("....." in k for k in kapitel))            # Verzeichnis

    def test_limit_is_not_prohibition(self):
        self.assertEqual(rq.classify_binding("The cycle time must not exceed 6 s.")[0], "MUSS")
        self.assertEqual(rq.classify_binding("Die Taktzeit darf 32 s nicht überschreiten.")[0], "MUSS")
        self.assertEqual(rq.classify_binding("The guard must not be removed.")[0], "VERBOT")


class SelbsttestTests(unittest.TestCase):
    """Die eingebauten Selbsttests laufen mit der Testsuite mit."""

    def test_requirements_self_test(self):
        import contextlib
        import io
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(rq.self_test(), 0)

    def test_stationen_self_test(self):
        import contextlib
        import io
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(st._selbsttest(), 0)


if __name__ == "__main__":
    unittest.main()
