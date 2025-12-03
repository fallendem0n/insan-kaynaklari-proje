import unittest
import os
import json
from unittest.mock import MagicMock, patch
from utils.template_manager import TemplateManager
from PIL import Image

class TestTemplateManager(unittest.TestCase):
    def setUp(self):
        # Create a dummy template
        self.template_name = "test_template"
        self.fields = [
            {'name': 'Title', 'x': 10, 'y': 10, 'w': 100, 'h': 20, 'page_width': 1000, 'page_height': 1000},
            {'name': 'Date', 'x': 10, 'y': 50, 'w': 50, 'h': 20, 'page_width': 1000, 'page_height': 1000}
        ]
        TemplateManager.save_template(self.template_name, self.fields)

    def tearDown(self):
        # Clean up
        path = os.path.join(TemplateManager.get_template_dir(), f"{self.template_name}.json")
        if os.path.exists(path):
            os.remove(path)

    def test_save_and_load_template(self):
        loaded = TemplateManager.load_template(self.template_name)
        self.assertEqual(len(loaded), 2)
        self.assertEqual(loaded[0]['name'], 'Title')

    @patch('utils.template_manager.convert_from_path')
    @patch('utils.template_manager.pytesseract')
    def test_extract_data(self, mock_pytesseract, mock_convert):
        # Mock image
        mock_img = Image.new('RGB', (1000, 1000), color='white')
        mock_convert.return_value = [mock_img]
        
        # Mock OCR result
        mock_pytesseract.image_to_string.side_effect = ["My Title", "01.01.2023"]
        
        data = TemplateManager.extract_data_with_template("dummy.pdf", self.template_name)
        
        self.assertEqual(data['Title'], "My Title")
        self.assertEqual(data['Date'], "01.01.2023")
        
        # Verify crop called with correct coords
        # Since we mocked image, we can't easily check crop calls on the mock object created inside the function 
        # unless we mock Image.open or similar. But we can check if pytesseract was called.
        self.assertEqual(mock_pytesseract.image_to_string.call_count, 2)

if __name__ == '__main__':
    unittest.main()
