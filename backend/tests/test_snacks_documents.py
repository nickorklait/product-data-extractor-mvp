import unittest

from backend.app.excel import populate_ipd_rows
from backend.app.parser import extract_structured_data


class SnacksDocumentTests(unittest.TestCase):
    def test_pipe_nutrition_and_named_allergens_are_mapped(self) -> None:
        text = """
Produkt | Anyday Dry Roasted Nut Mix
Allergener | Produktet inneholder cashew, peanøtter, mandler og hasselnøtter. Kan inneholde spor av andre nøtter.
Datomerking pose | Produksjonsdato + 270 dager.
Næringsinnhold | Per 100 g | Per porsjon (30g)
Energi (kJ/kcal) | 2536/605 | 762/182
Fett | 49 g | 15 g
hvorav mettede fettsyrer | 7.1 g | 2.1 g
Karbohydrater | 18 g | 5.5 g
hvorav sukkerarter | 4.5 g | 1.3 g
Kostfiber | 7.5 g | 2.2 g
Protein | 20 g | 6.0 g
Salt | 0.8 g | 0.2 g
"""

        result = populate_ipd_rows(extract_structured_data(text, "anyday.docx"))
        mapped = {row.attribute: row.data for row in result.ipd_rows}

        self.assertEqual(result.nutrition.energy_kj, "2536")
        self.assertEqual(result.nutrition.energy_kcal, "605")
        self.assertEqual(result.nutrition.saturated_fat_g, "7.1")
        self.assertEqual(result.shelf_life.shelf_life_value, "270")
        self.assertEqual(result.allergens.peanuts, "Yes")
        self.assertEqual(result.allergens.nuts, "Yes")
        self.assertIn("andre nøtter", result.allergens.allergens_may_contain or "")
        self.assertEqual(mapped["Cashew"], "Yes")
        self.assertEqual(mapped["Hazelnut"], "Yes")
        self.assertEqual(mapped["Almonds"], "Yes")
        self.assertEqual(mapped["Other nuts"], "Traces")

    def test_finished_product_shelf_life_wins_and_process_temperature_is_ignored(self) -> None:
        text = """
Produkt | KiMs Snack Chips
Allergener | Melk
Holdbarhet pellets | 1 år
Veiledende Set.temp | 199 °C
Datomerking pose | PD + 180 dager
NÆRINGSINNHOLD Per 100 g 30g
Energi (kJ/kcal) 1957/467 581/139
Fett 18g 5,4g
hvorav mettede fettsyrer 1,6g 0,5g
Karbohydrat 68g 20g
hvorav sukkerarter 219 0,6g
Kostfiber 50g 15g
Protein 569 1,7g
Salt 239 0,7g
"""

        result = extract_structured_data(text, "snack-chips.docx")

        self.assertEqual(result.shelf_life.shelf_life_value, "180")
        self.assertEqual(result.shelf_life.shelf_life_unit, "dager")
        self.assertIsNone(result.storage.storage_temperature)
        self.assertEqual(result.allergens.milk, "Yes")
        self.assertEqual(result.nutrition.fat_g, "18")
        self.assertEqual(result.nutrition.carbohydrate_g, "68")
        self.assertIsNone(result.nutrition.sugars_g)
        self.assertIsNone(result.nutrition.fibre_g)
        self.assertIsNone(result.nutrition.protein_g)
        self.assertIsNone(result.nutrition.salt_g)

    def test_fragmented_nutrition_keeps_only_values_with_evidence(self) -> None:
        text = """
Produkt | IKEA Cashews
Allergener | Cashew, kan inneholde spor av andre nøtter.
£ Storage: Store atreomteme
Oppbevaring: Oppbevares ved romtemperatur.
word/media/image8.png:
PD+1 år Datomerkes under forhåndstrykket "best før".
Næringsinnhold pr. 100 g
Energi / Energi 2494 kJ
597 kcal
Fedt / Fett
heraf mættede fedtsyrer / hvorav mettede fettsyrer
Kulhydrat / Karbohydrater
heraf sukkerarter / hvorav sukkerarter
Protein / Protein 21g
Salt / Salt 0.30g
"""

        result = populate_ipd_rows(extract_structured_data(text, "ikea-cashews.docx"))
        mapped = {row.attribute: row.data for row in result.ipd_rows}

        self.assertEqual(result.nutrition.energy_kj, "2494")
        self.assertEqual(result.nutrition.energy_kcal, "597")
        self.assertEqual(result.nutrition.protein_g, "21")
        self.assertEqual(result.nutrition.salt_g, "0.30")
        self.assertIsNone(result.nutrition.fat_g)
        self.assertIsNone(result.nutrition.carbohydrate_g)
        self.assertEqual(mapped["Cashew"], "Yes")
        self.assertEqual(mapped["Other nuts"], "Traces")
        self.assertEqual(result.shelf_life.shelf_life_text, 'PD+1 år Datomerkes under forhåndstrykket "best før".')
        self.assertIn("Oppbevares ved romtemperatur", result.storage.storage_text or "")

    def test_no_allergens_sets_explicit_no_values(self) -> None:
        result = populate_ipd_rows(
            extract_structured_data("Produkt | Chips\nAllergener | Ingen allergener", "chips.docx")
        )
        allergen_rows = [row for row in result.ipd_rows if row.node == "Allergens"]

        self.assertTrue(allergen_rows)
        self.assertTrue(all(row.data == "No" for row in allergen_rows))


if __name__ == "__main__":
    unittest.main()
