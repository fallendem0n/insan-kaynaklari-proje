import json
import os
import sys
from PIL import Image
try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None

class TemplateManager:
    TEMPLATE_DIR = "templates"

    @staticmethod
    def get_template_dir():
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        template_path = os.path.join(base_path, TemplateManager.TEMPLATE_DIR)
        if not os.path.exists(template_path):
            os.makedirs(template_path)
        return template_path

    @staticmethod
    def save_template(name, fields):
        """
        Saves a template to a JSON file.
        fields: list of dicts {'name': str, 'x': int, 'y': int, 'w': int, 'h': int, 'page_width': int, 'page_height': int}
        """
        path = os.path.join(TemplateManager.get_template_dir(), f"{name}.json")
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(fields, f, ensure_ascii=False, indent=4)

    @staticmethod
    def load_template(name):
        path = os.path.join(TemplateManager.get_template_dir(), f"{name}.json")
        if not os.path.exists(path):
            return []
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    @staticmethod
    def get_all_templates():
        directory = TemplateManager.get_template_dir()
        files = [f.replace(".json", "") for f in os.listdir(directory) if f.endswith(".json")]
        return sorted(files)

    @staticmethod
    def extract_data_with_template(pdf_path, template_name, poppler_path=None, tesseract_path=None):
        """
        Extracts data from a PDF using the specified template.
        """
        fields = TemplateManager.load_template(template_name)
        if not fields:
            return {}

        if not convert_from_path or not pytesseract:
            raise ImportError("OCR kütüphaneleri (pdf2image, pytesseract) eksik.")

        if tesseract_path:
             pytesseract.pytesseract.tesseract_cmd = tesseract_path

        try:
            # Convert first page to image
            images = convert_from_path(pdf_path, poppler_path=poppler_path, first_page=1, last_page=1)
            if not images:
                return {}
            
            image = images[0]
            img_w, img_h = image.size
            
            extracted_data = {}
            
            for field in fields:
                # Scale coordinates if necessary
                # The template was saved with a specific page size (page_width, page_height)
                # We need to scale the coordinates to the current image size
                
                orig_w = field.get('page_width', img_w)
                orig_h = field.get('page_height', img_h)
                
                scale_x = img_w / orig_w
                scale_y = img_h / orig_h
                
                x = int(field['x'] * scale_x)
                y = int(field['y'] * scale_y)
                w = int(field['w'] * scale_x)
                h = int(field['h'] * scale_y)
                
                # Crop
                cropped = image.crop((x, y, x + w, y + h))
                
                # OCR
                text = pytesseract.image_to_string(cropped, lang='tur+eng', config='--psm 7') # PSM 7: Treat as single text line
                extracted_data[field['name']] = text.strip()
                
            return extracted_data
            
        except Exception as e:
            print(f"Template extraction error: {e}")
            return {}
