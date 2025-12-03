import unittest
import os
import pandas as pd
from tools.pdf_data_extractor_tool import PDFDataExtractorFrame
import customtkinter as ctk

class TestPDFDataExtractor(unittest.TestCase):
    def setUp(self):
        # Create a hidden root for the frame
        self.root = ctk.CTk()
        self.extractor = PDFDataExtractorFrame(master=self.root)
        
    def tearDown(self):
        self.root.destroy()

    def test_pattern_matching(self):
        # Simulate extracted text
        text = """
        FORD OTOSAN
        Sicil No: 12345
        Adı Soyadı: Ahmet Yılmaz
        Tutar: 1000 TL
        Tarih: 01.01.2023
        """
        
        # Test pattern matching logic
        val1 = self.extractor.find_value_for_pattern(text, "Sicil No:")
        self.assertEqual(val1, "12345")
        
        val2 = self.extractor.find_value_for_pattern(text, "Adı Soyadı:")
        self.assertEqual(val2, "Ahmet Yılmaz")
        
        val3 = self.extractor.find_value_for_pattern(text, "Tutar:")
        self.assertEqual(val3, "1000 TL")
        
        # Test non-existent pattern
        val4 = self.extractor.find_value_for_pattern(text, "Olmayan:")
        self.assertEqual(val4, "")

    def test_case_insensitivity(self):
        text = "Sicil No: 98765"
        val = self.extractor.find_value_for_pattern(text, "sicil no:")
        self.assertEqual(val, "98765")

if __name__ == '__main__':
    unittest.main()
