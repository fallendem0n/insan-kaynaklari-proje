import json
import os
import sys
from PIL import Image
try:
    import pytesseract
    import fitz # PyMuPDF
except ImportError:
    pytesseract = None
    fitz = None

# şablon yönetim sınıfı
class TemplateManager:
    TEMPLATE_DIR = "templates"

    # şablon klasörünün yolunu buluyoruz
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

    # şablonu json dosyası olarak kaydediyoruz
    @staticmethod
    def save_template(name, fields):
        """
        şablonu json dosyasına kaydeder.
        fields: alanların listesi (koordinatlar ve isimler)
        """
        path = os.path.join(TemplateManager.get_template_dir(), f"{name}.json")
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(fields, f, ensure_ascii=False, indent=4)

    # şablonu dosyadan yüklüyoruz
    @staticmethod
    def load_template(name):
        path = os.path.join(TemplateManager.get_template_dir(), f"{name}.json")
        if not os.path.exists(path):
            return []
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    # tüm kayıtlı şablonları listeliyoruz
    @staticmethod
    def get_all_templates():
        directory = TemplateManager.get_template_dir()
        files = [f.replace(".json", "") for f in os.listdir(directory) if f.endswith(".json")]
        return sorted(files)

    # şablon kullanarak pdf'den veri çıkarıyoruz
    # şablon kullanarak pdf'den veri çıkarıyoruz
    @staticmethod
    @staticmethod
    def extract_data_with_template(pdf_path, template_name, tesseract_path=None):
        """
        belirtilen şablonu kullanarak pdf'den veri çeker.
        """
        fields = TemplateManager.load_template(template_name)
        if not fields:
            return {}

        if not fitz or not pytesseract:
            raise ImportError("ocr kütüphaneleri (PyMuPDF, pytesseract) eksik.")

        if tesseract_path:
             pytesseract.pytesseract.tesseract_cmd = tesseract_path

        try:
            # pdf'in ilk sayfasını resme çeviriyoruz
            doc = fitz.open(pdf_path)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # 2x zoom
            
            import io
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            
            img_w, img_h = image.size
            
            extracted_data = {}
            
            for field in fields:
                # koordinatları ölçekliyoruz
                # şablon kaydedilirken sayfa boyutu da kaydedilmişti
                # şu anki resim boyutu farklı olabilir, o yüzden oranlıyoruz
                
                orig_w = field.get('page_width', img_w)
                orig_h = field.get('page_height', img_h)
                
                scale_x = img_w / orig_w
                scale_y = img_h / orig_h
                
                x = int(field['x'] * scale_x)
                y = int(field['y'] * scale_y)
                w = int(field['w'] * scale_x)
                h = int(field['h'] * scale_y)
                
                # resmi kırpıyoruz (ilgili alanı alıyoruz)
                cropped = image.crop((x, y, x + w, y + h))
                
                # eğer döndürülmüş ise resmi çeviriyoruz
                if field.get('is_rotated', False):
                    # görsel şablon aracında olduğu gibi 90 derece sola çeviriyoruz
                    cropped = cropped.rotate(90, expand=True)
                
                # ocr ile metni okuyoruz
                # psm 7: tek satır metin olarak algıla
                text = pytesseract.image_to_string(cropped, lang='tur+eng', config='--psm 7') 
                extracted_data[field['name']] = text.strip()
                
            return extracted_data
            
        except Exception as e:
            print(f"şablon veri çıkarma hatası: {e}")
            return {}
