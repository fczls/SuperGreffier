import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from anonymisation import apply_anonymization, find_residual_names

TABLE = {"Malo Royen": "Thomas Bertrand", "L.K.": "A.B."}


class TestAnonymisation(unittest.TestCase):
    def test_nom_complet(self):
        self.assertEqual(apply_anonymization("Malo Royen parle.", TABLE), "Thomas Bertrand parle.")

    def test_prenom_seul(self):
        self.assertEqual(apply_anonymization("Malo a parlé.", TABLE), "Thomas a parlé.")

    def test_nom_seul(self):
        self.assertEqual(apply_anonymization("Royen a parlé.", TABLE), "Bertrand a parlé.")

    def test_casse(self):
        self.assertEqual(apply_anonymization("MALO ROYEN, royen.", TABLE), "THOMAS BERTRAND, Bertrand.")

    def test_mots_entiers(self):
        self.assertEqual(apply_anonymization("Royenne et Malory", TABLE), "Royenne et Malory")

    def test_initiales(self):
        self.assertEqual(apply_anonymization("Me L.K. plaide.", TABLE), "Me A.B. plaide.")

    def test_pas_de_chainage(self):
        table = {"Alice Martin": "Bob Durand", "Bob Durand": "Carl Petit"}
        self.assertEqual(apply_anonymization("Alice Martin et Bob Durand", table), "Bob Durand et Carl Petit")

    def test_entree_explicite_prioritaire(self):
        table = {"Malo Royen": "Thomas Bertrand", "Malo": "Tom"}
        self.assertEqual(apply_anonymization("Malo Royen, puis Malo", table), "Thomas Bertrand, puis Tom")

    def test_nom_partage_ambigu_non_devine(self):
        table = {"Anne Martin": "Julie Lemoine", "Paul Martin": "Marc Roux"}
        out = apply_anonymization("Martin a parlé", table)
        self.assertEqual(out, "Martin a parlé")
        self.assertIn("Martin", find_residual_names(out, table))

    def test_table_vide(self):
        self.assertEqual(apply_anonymization("Rien à faire", {}), "Rien à faire")

    def test_residuel_vide_apres_anonymisation(self):
        out = apply_anonymization("Malo Royen puis Royen puis MALO", TABLE)
        self.assertEqual(find_residual_names(out, TABLE), [])


if __name__ == "__main__":
    unittest.main()
