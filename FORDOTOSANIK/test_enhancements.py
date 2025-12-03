import unittest
import os
import shutil
import customtkinter as ctk
from tools.pdf_renamer_tool import PDFRenamerFrame
from utils.backup_manager import BackupManager

class TestEnhancements(unittest.TestCase):
    def setUp(self):
        self.root = ctk.CTk()
        self.renamer = PDFRenamerFrame(master=self.root)
        
        # Create dummy file for backup test
        self.test_file = "test_backup.txt"
        with open(self.test_file, "w") as f:
            f.write("test content")
            
    def tearDown(self):
        self.root.destroy()
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        if os.path.exists("Yedek"):
            shutil.rmtree("Yedek")

    def test_backup_manager(self):
        # Test backup creation
        backed_up = BackupManager.create_backup([os.path.abspath(self.test_file)])
        self.assertTrue(len(backed_up) > 0)
        self.assertTrue(os.path.exists(backed_up[0]))
        self.assertIn("Yedek", backed_up[0])

    def test_renamer_pattern_logic(self):
        text = "TC: 12345678901\nAd Soyad: Mehmet Demir"
        
        # Test finding values
        tc = self.renamer.find_value_for_pattern(text, "TC:")
        self.assertEqual(tc, "12345678901")
        
        ad = self.renamer.find_value_for_pattern(text, "Ad Soyad:")
        self.assertEqual(ad, "Mehmet Demir")
        
        # Test multiline
        text_multiline = "TC:\n98765432109\n"
        tc_multi = self.renamer.find_value_for_pattern(text_multiline, "TC:")
        self.assertEqual(tc_multi, "98765432109")

        # Test flexible whitespace and optional colon
        text_spaced = "S i c i l  N o : 55555"
        # User enters "Sicil No", text has spaces and colon
        val_spaced = self.renamer.find_value_for_pattern(text_spaced, "Sicil No")
        self.assertEqual(val_spaced, "55555")
        
        text_no_colon = "Ad Soyad 12345"
        # User enters "Ad Soyad:", text has no colon
        val_no_colon = self.renamer.find_value_for_pattern(text_no_colon, "Ad Soyad:")
        self.assertEqual(val_no_colon, "12345")

        # Test formatting logic (simulation)
        values = {"TC": tc, "Ad Soyad": ad}
        format_str = "{TC}_{Ad Soyad}"
        new_name = format_str.format(**values)
        self.assertEqual(new_name, "12345678901_Mehmet Demir")

if __name__ == '__main__':
    unittest.main()
