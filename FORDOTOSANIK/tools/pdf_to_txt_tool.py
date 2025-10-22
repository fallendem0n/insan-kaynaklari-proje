import customtkinter as ctk
from tkinter import filedialog
import os
import threading

try:
    from PyPDF2 import PdfReader
except ImportError:
    try:
        from pypdf import PdfReader
    except ImportError:
        print("PyPDF2/pypdf kütüphanesi bulunamadı.")
        PdfReader = None 

try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
except ImportError:
    print("Tesseract veya pdf2image kütüphaneleri eksik.")
    pytesseract = None
    convert_from_path = None

class PDFToTXTFrame(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)

        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        tesseract_path = os.path.join(base_dir, "tesseract", "tesseract.exe")
        if pytesseract and os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        elif pytesseract:
            print(f"Uyarı: Tesseract yolu bulunamadı: {tesseract_path}")

        self.poppler_path = os.path.join(base_dir, "poppler", "Library", "bin")
        if not os.path.exists(self.poppler_path):
             print(f"Uyarı: Poppler yolu bulunamadı: {self.poppler_path}")

        self.grid_columnconfigure(0, weight=1)
        
        self.title_label = ctk.CTkLabel(self, text="PDF'den Metin Çıkarıcı", font=ctk.CTkFont(size=16, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(10, 10), sticky="ew")

        self.file_path_label = ctk.CTkLabel(self, text="Henüz PDF seçilmedi")
        self.file_path_label.grid(row=1, column=0, padx=20, pady=5, sticky="ew")

        self.select_file_button = ctk.CTkButton(self, text="PDF Dosyası Seç", command=self.select_pdf)
        self.select_file_button.grid(row=2, column=0, padx=20, pady=10)

        self.ocr_var = ctk.StringVar(value="on")
        self.ocr_check = ctk.CTkCheckBox(self, text="Gerekirse OCR kullan (Açık kalması önerilir.)", variable=self.ocr_var, onvalue="on", offvalue="off")
        self.ocr_check.grid(row=3, column=0, padx=20, pady=5, sticky="w")

        self.convert_button = ctk.CTkButton(self, text="Metne Dönüştür", command=self.start_conversion_thread, state="disabled")
        self.convert_button.grid(row=4, column=0, padx=20, pady=10)

        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.grid(row=5, column=0, padx=20, pady=(10, 20), sticky="ew")
        
        self.selected_pdf_path = ""

    def select_pdf(self):
        file_path = filedialog.askopenfilename(
            title="PDF dosyası seçin",
            filetypes=(("PDF Dosyaları", "*.pdf"), ("Tüm Dosyalar", "*.*"))
        )
        if file_path:
            self.selected_pdf_path = file_path
            self.file_path_label.configure(text=os.path.basename(file_path))
            self.convert_button.configure(state="normal")
            self.status_label.configure(text="")
        else:
            self.selected_pdf_path = ""
            self.file_path_label.configure(text="Henüz PDF seçilmedi")
            self.convert_button.configure(state="disabled")

    def start_conversion_thread(self):
        """
        Dönüştürme işlemini arayüzü kilitlememesi için
        ayrı bir thread'de başlatır.
        """
        if not self.selected_pdf_path:
            self.status_label.configure(text="Lütfen önce bir PDF dosyası seçin.", text_color="orange")
            return
            
        self.convert_button.configure(state="disabled")
        self.select_file_button.configure(state="disabled")
        self.status_label.configure(text="Dönüştürme işlemi sürüyor, lütfen bekleyin...", text_color="cyan")

        conversion_thread = threading.Thread(target=self.convert_to_txt)
        conversion_thread.daemon = True 
        conversion_thread.start()

    def convert_to_txt(self):
        try:
            full_text = ""
            
            if PdfReader:
                try:
                    reader = PdfReader(self.selected_pdf_path)
                    for page in reader.pages:
                        extracted = page.extract_text()
                        if extracted:
                            full_text += extracted + "\n"
                except Exception as e:
                    print(f"PyPDF2 hatası: {e}")
                    full_text = "" 

            use_ocr = self.ocr_var.get() == "on"
            if use_ocr and (not full_text or len(full_text) < 1024):
                if not pytesseract or not convert_from_path:
                    self.status_label.configure(text="Hata: OCR kütüphaneleri yüklenemedi.", text_color="red")
                    self.ui_reset()
                    return

                self.status_label.configure(text="Metin bulunamadı, OCR deneniyor...", text_color="cyan")
                
                try:
                    images = convert_from_path(
                        self.selected_pdf_path,
                        poppler_path=self.poppler_path
                    )
                    
                    full_text = "" 
                    for i, img in enumerate(images):
                        self.status_label.configure(text=f"OCR Sayfa {i+1}/{len(images)} işleniyor...", text_color="cyan")
                        text = pytesseract.image_to_string(img, lang='tur+eng')
                        full_text += text + "\n\n"
                        
                except Exception as ocr_error:
                    self.status_label.configure(text=f"OCR Hatası: {ocr_error}", text_color="red")
                    print(f"OCR Hatası: {ocr_error}")
                    self.ui_reset()
                    return

            if not full_text.strip():
                self.status_label.configure(text="PDF'den metin çıkarılamadı.", text_color="orange")
                self.ui_reset()
                return

            output_path = os.path.splitext(self.selected_pdf_path)[0] + ".txt"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(full_text)
            
            self.status_label.configure(text=f"Başarılı! Dosya şuraya kaydedildi: \n{output_path}", text_color="green")
            
        except Exception as e:
            self.status_label.configure(text=f"Genel Hata: {e}", text_color="red")
            print(f"Hata: {e}")
            
        finally:
            self.ui_reset()

    def ui_reset(self):
        """Arayüz bileşenlerini sıfırlar."""
        self.convert_button.configure(state="normal")
        self.select_file_button.configure(state="normal")