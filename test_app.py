"""
Test suite para verificar la logica de procesamiento y generacion de reportes.
"""

import unittest
import os
import io
import pandas as pd
import processor
import pdf_generator

class TestCommissionsProcessing(unittest.TestCase):

    def setUp(self):
        self.sample_file1 = os.path.join(os.path.dirname(__file__), 'sample_data', 'NLG_PayableGrossCommission_11232026.xlsx')
        self.sample_file2 = os.path.join(os.path.dirname(__file__), 'sample_data', 'NLG_PayableGrossCommission_11172026.xlsx')

    def test_single_file_extraction(self):
        if not os.path.exists(self.sample_file1):
            self.skipTest("Sample file 1 not found.")
        
        df, exported_by, exported_on = processor.extract_excel_data(self.sample_file1)
        self.assertEqual(len(df), 20)
        self.assertEqual(exported_by, "Gutierrez, Manuel Hernandez")
        self.assertEqual(exported_on, "11/23/2026")
        self.assertEqual(list(df.columns), processor.REQUIRED_COLUMNS)

    def test_split_camel_case(self):
        self.assertEqual(processor.split_camel_case("AlexanderArellano"), "Alexander Arellano")
        self.assertEqual(processor.split_camel_case("JuanLopez Bravo"), "Juan Lopez Bravo")
        self.assertEqual(processor.split_camel_case("KarenBaquero"), "Karen Baquero")
        self.assertEqual(processor.split_camel_case("Arellano,Alexander"), "Arellano, Alexander")
        self.assertEqual(processor.split_camel_case("Gutierrez,ManuelHernandez"), "Gutierrez, Manuel Hernandez")
        self.assertEqual(processor.split_camel_case("ALEXANDER ARELLANO"), "ALEXANDER ARELLANO")

    def test_name_matching(self):
        # Gutierrez, Manuel Hernandez <-> HERNANDEZ GUTIERREZ MANUEL
        self.assertTrue(processor.names_match("Gutierrez, Manuel Hernandez", "HERNANDEZ GUTIERREZ MANUEL"))
        self.assertTrue(processor.names_match("HERNANDEZ GUTIERREZ MANUEL", "Gutierrez, Manuel Hernandez"))
        self.assertTrue(processor.names_match("Betancourt, Carol", "BETANCOURT CAROL A"))
        # Prueba con CamelCase pegado en Exported By:
        self.assertTrue(processor.names_match("AlexanderArellano", "ARELLANO ALEXANDER"))
        self.assertTrue(processor.names_match("ARELLANO ALEXANDER", "AlexanderArellano"))
        self.assertTrue(processor.names_match("JuanLopez Bravo", "LOPEZ BRAVO JUAN"))
        self.assertFalse(processor.names_match("Gutierrez, Manuel Hernandez", "ABREU ANABEL"))
        self.assertFalse(processor.names_match("Gutierrez, Manuel Hernandez", "HERRERA YUDEMA"))

    def test_clean_currency(self):
        self.assertEqual(processor.clean_currency("$1,530.00"), 1530.0)
        self.assertEqual(processor.clean_currency("($202.50)"), -202.5)
        self.assertEqual(processor.clean_currency("-45.34"), -45.34)
        self.assertEqual(processor.clean_currency(64), 64.0)
        self.assertEqual(processor.clean_currency(None), 0.0)
        self.assertEqual(processor.clean_currency(""), 0.0)

    def test_commission_classification_and_text(self):
        if not os.path.exists(self.sample_file1):
            self.skipTest("Sample file 1 not found.")

        df, exported_by, exported_on = processor.extract_excel_data(self.sample_file1)
        unified_df, final_by, final_on = processor.unify_dataframes([(df, exported_by, exported_on)])
        result = processor.process_commissions(unified_df, final_by)

        principal = result['principal']
        team = result['team']

        # 3 principal insureds: Rafael Aguilar Perez, Emilio O Olivera, Gabino Ruiz
        principal_dict = {p['insured']: p['status'] for p in principal}
        self.assertEqual(principal_dict.get('Rafael Aguilar Perez'), 'Activa')
        self.assertEqual(principal_dict.get('Emilio O Olivera'), 'Activa')
        self.assertEqual(principal_dict.get('Gabino Ruiz'), 'Chargeback')

        # Check team chargebacks
        team_dict = {t['insured']: t['status'] for t in team}
        self.assertEqual(team_dict.get('Erick Jr Lucero'), 'Chargeback')
        self.assertEqual(team_dict.get('Elly Maria'), 'Chargeback')
        self.assertEqual(team_dict.get('Franki J Leon Calzadilla'), 'Chargeback')
        self.assertEqual(team_dict.get('Sofia Moreno Aguilera'), 'Activa')

        text = processor.generate_prebuilt_text(result)
        self.assertIn("Comision y Chargeback NL", text)
        self.assertIn("* Rafael Aguilar Perez: Activa", text)
        self.assertIn("* Gabino Ruiz: Chargeback", text)
        self.assertIn("Comision y Chargeback de Equipo NL", text)
        self.assertIn("* Sofia Moreno Aguilera: Activa", text)

    def test_pdf_generation(self):
        if not os.path.exists(self.sample_file1):
            self.skipTest("Sample file 1 not found.")

        df, exported_by, exported_on = processor.extract_excel_data(self.sample_file1)
        pdf_buf = pdf_generator.build_unified_pdf(df, exported_by, exported_on)
        pdf_bytes = pdf_buf.getvalue()
        self.assertGreater(len(pdf_bytes), 10000)
        # Check standard filename
        fname = pdf_generator.clean_filename(exported_by)
        self.assertEqual(fname, "NLG_Gutierrez, Manuel Hernandez_PayableGrossComission.pdf")

    def test_two_files_unification(self):
        if not (os.path.exists(self.sample_file1) and os.path.exists(self.sample_file2)):
            self.skipTest("Sample files not found.")

        d1 = processor.extract_excel_data(self.sample_file1)
        d2 = processor.extract_excel_data(self.sample_file2)
        unified_df, final_by, final_on = processor.unify_dataframes([d1, d2])
        self.assertEqual(len(unified_df), 20 + 40)
        self.assertEqual(len(unified_df.columns), 10)

if __name__ == '__main__':
    unittest.main()
