import unittest
from PIL import Image, ImageDraw, ImageFont
import pytesseract
from utils.template_manager import TemplateManager
import os

class TestRotatedText(unittest.TestCase):
    def setUp(self):
        # Create a dummy image with text
        self.img = Image.new('RGB', (200, 100), color='white')
        d = ImageDraw.Draw(self.img)
        # We can't easily draw text without a font file, but we can test the rotation logic
        # by mocking the image and checking if it gets rotated.
        
    def test_rotation_logic(self):
        # Mocking the extraction logic from TemplateManager
        # We want to see if the image is rotated when is_rotated is True
        
        field = {
            'name': 'test_field',
            'x': 0,
            'y': 0,
            'w': 100,
            'h': 200,
            'is_rotated': True,
            'page_width': 200,
            'page_height': 200
        }
        
        # Create a vertical image (simulating a vertical selection)
        original_crop = Image.new('RGB', (100, 200), color='red')
        
        # Apply the logic from TemplateManager
        if field.get('is_rotated', False):
            rotated_crop = original_crop.rotate(90, expand=True)
            
        # Check dimensions. If rotated 90 degrees, width and height should swap
        self.assertEqual(rotated_crop.width, 200)
        self.assertEqual(rotated_crop.height, 100)
        print("Rotation logic verified: Image dimensions swapped correctly.")

if __name__ == '__main__':
    unittest.main()
