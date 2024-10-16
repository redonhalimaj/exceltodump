# tests/test_reader.py

import unittest
from exceltodump.reader import categorize_param_value, extract_operations_params

class TestReader(unittest.TestCase):

    def test_categorize_param_value_numeric(self):
        self.assertEqual(categorize_param_value("123"), "Numeric")
        self.assertEqual(categorize_param_value("45.67"), "Numeric")

    def test_categorize_param_value_text(self):
        self.assertEqual(categorize_param_value('"A2"'), "Text")
        self.assertEqual(categorize_param_value('"TV_Signal1"'), "Text")
        self.assertEqual(categorize_param_value("10 sec"), "Text")

    def test_categorize_param_value_comparison(self):
        self.assertEqual(categorize_param_value('"=="'), "Comparison")
        self.assertEqual(categorize_param_value('">="'), "Comparison")

    def test_categorize_param_value_empty(self):
        self.assertEqual(categorize_param_value(""), "Text")  # Defaults to 'Text'

    def test_extract_operations_params(self):
        text = '#op[Set_Step_Precondition](#p[1]) [Testvorbedingung] [Betriebszustand]'
        categorized_params = {"Text": set(), "Numeric": set(), "Comparison": set()}
        descriptions, ops, categorized_params = extract_operations_params(text, categorized_params)
        self.assertEqual(len(descriptions), 2)
        self.assertEqual(len(ops), 1)
        self.assertIn("1", categorized_params["Numeric"])

if __name__ == '__main__':
    unittest.main()
