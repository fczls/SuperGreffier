import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from energie import estimate_energy, format_range, pc_minutes, format_duration, minutes_evitees, format_hm, DEMO_G_CO2


class TestEnergie(unittest.TestCase):
    def test_zero(self):
        r = estimate_energy(0, 0)
        self.assertEqual(r["wh"], (0, 0))
        self.assertEqual(r["g_co2"], (0, 0))

    def test_fourchette_ordonnee(self):
        r = estimate_energy(5000, 6000)
        self.assertLess(r["wh"][0], r["wh"][1])
        self.assertLess(r["g_co2"][0], r["g_co2"][1])

    def test_ordre_de_grandeur_document_10_pages(self):
        # ~5k tokens en entrée, ~6k en sortie : quelques Wh, pas des centaines
        low, high = estimate_energy(5000, 6000)["wh"]
        self.assertGreater(low, 1)
        self.assertLess(high, 100)

    def test_sortie_pese_plus_que_entree(self):
        self.assertGreater(
            estimate_energy(0, 1000)["wh"][1], estimate_energy(1000, 0)["wh"][1]
        )

    def test_tokens_negatifs_ignores(self):
        self.assertEqual(estimate_energy(-5, -5)["wh"], (0, 0))

    def test_format(self):
        self.assertEqual(format_range(3.2, 38.9, "Wh"), "3.2 à 39 Wh")
        self.assertEqual(format_range(0.25, 0.9, "g"), "0.25 à 0.90 g")

    def test_pc_minutes_demo(self):
        bas, haut = pc_minutes(*DEMO_G_CO2)
        self.assertAlmostEqual(bas, 0.24, places=2)   # 0,32 g à 80 g/h
        self.assertAlmostEqual(haut, 42.0, places=1)  # 28 g à 40 g/h

    def test_format_duration(self):
        self.assertEqual(format_duration(0.24), "14 s")
        self.assertEqual(format_duration(42), "42 min")
        self.assertEqual(format_duration(150), "2.5 h")

    def test_minutes_evitees_4h_demo(self):
        # 4 h - 42 min (pire cas démo : 28 g à 40 g/h) = 3 h 18
        self.assertEqual(format_hm(minutes_evitees(4, DEMO_G_CO2[1])), "3 h 18")

    def test_minutes_evitees_jamais_negatif(self):
        self.assertEqual(minutes_evitees(0.5, 1000), 0)


if __name__ == "__main__":
    unittest.main()
