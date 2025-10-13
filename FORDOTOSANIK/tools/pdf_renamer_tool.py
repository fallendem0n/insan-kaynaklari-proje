import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import re
import pypdf
from threading import Thread
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import sys

class PDFRenamerFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        self.tesseract_path = os.path.join(application_path, 'tesseract', 'tesseract.exe')
        self.poppler_path = os.path.join(application_path, 'poppler', 'Library', 'bin')
        
        try:
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
        except Exception:
            pass

        self.create_widgets()
        self.selected_files = []

    def create_widgets(self):
        # (Bu fonksiyon değişmedi)
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(padx=10, pady=10, fill="x")
        select_button = ctk.CTkButton(file_frame, text="PDF Dosyaları Seç", command=self.select_pdfs)
        select_button.pack(pady=10, padx=10)
        self.file_list_box = ctk.CTkTextbox(self, height=150, state="disabled")
        self.file_list_box.pack(padx=10, pady=5, fill="x", expand=True)
        format_frame = ctk.CTkFrame(self)
        format_frame.pack(padx=10, pady=10, fill="x")
        format_label = ctk.CTkLabel(format_frame, text="Yeni Dosya Adı Formatı:")
        format_label.pack(side="left", padx=10)
        self.format_combo = ctk.CTkComboBox(format_frame, values=["{TC} - {ADSOYAD}", "{SICIL} - {ADSOYAD}", "{ADSOYAD}", "{TC}"])
        self.format_combo.pack(side="left", padx=10, fill="x", expand=True)
        self.format_combo.set("{TC} - {ADSOYAD}")
        action_frame = ctk.CTkFrame(self)
        action_frame.pack(padx=10, pady=10, fill="x")
        self.rename_button = ctk.CTkButton(action_frame, text="Yeniden Adlandırmayı Başlat", command=self.start_rename_thread, state="disabled")
        self.rename_button.pack(pady=10, padx=10, fill="x")
        self.progress_bar = ctk.CTkProgressBar(action_frame)
        self.progress_bar.pack(pady=5, padx=10, fill="x")
        self.progress_bar.set(0)
        self.status_label = ctk.CTkLabel(self, text="Lütfen dosyaları seçin...")
        self.status_label.pack(pady=5, padx=10)

    def select_pdfs(self):
        # (Bu fonksiyon değişmedi)
        self.selected_files = filedialog.askopenfilenames(title="İşlem Yapılacak PDF Dosyalarını Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if not self.selected_files:
            self.rename_button.configure(state="disabled")
            return
        self.file_list_box.configure(state="normal")
        self.file_list_box.delete("1.0", "end")
        for f in self.selected_files:
            self.file_list_box.insert("end", os.path.basename(f) + "\n")
        self.file_list_box.configure(state="disabled")
        self.rename_button.configure(state="normal")
        self.status_label.configure(text=f"{len(self.selected_files)} dosya seçildi.")
    
    def find_info_in_text(self, text):
        info = {'TC': None, 'ADSOYAD': None, 'SICIL': None}
        
        tc_patterns = [
            re.compile(r'(\b[1-9][0-9]{10})\b'),
            re.compile(r'(\b[1-9][0-9]{2}\s?[0-9]{3}\s?[0-9]{3}\s?[0-9]{2}\b)')
        ]
        
        adsoyad_patterns = [
            re.compile(r'AD SOYAD\s+([A-ZÇĞİÖŞÜ\s]+?)\s+İŞYERİ', re.IGNORECASE),
            re.compile(r'^(?:Ad Soyad|Adı Soyadı)\s*[:\-]\s*([^\n\r]+)', re.IGNORECASE | re.MULTILINE)
        ]

        sicil_patterns = [
            re.compile(r'(?:Sicil|Personel|Dosya)\s*(?:No|Numarası)?\s*[:\-]?\s*([A-Za-z0-9-]+)', re.IGNORECASE)
        ]

        for pattern in tc_patterns:
            match = pattern.search(text)
            if match:
                info['TC'] = re.sub(r'\s+', '', match.group(1).strip())
                break

        for pattern in adsoyad_patterns:
            match = pattern.search(text)
            if match:
                adsoyad = match.group(1).strip()
                adsoyad_clean = re.sub(r'\s+', ' ', adsoyad)
                info['ADSOYAD'] = ' '.join(word.capitalize() for word in adsoyad_clean.split())
                break

        for pattern in sicil_patterns:
            match = pattern.search(text)
            if match:
                info['SICIL'] = match.group(1).strip()
                break
                
        return info

    # --- YENİ ve EN KARARLI YÖN BULMA FONKSİYONU ---
    def ocr_with_orientation_check(self, image):
        """
        Bir resmi 4 farklı açıda (0, 90, 180, 270) OCR'dan geçirir ve
        içinde anahtar kelimeler bulunan ilk anlamlı metni döndürür.
        """
        angles = [0, 270, 180, 90]  # Denenecek açılar
        keywords = ["Ad", "Soyad", "Kimlik", "T.C", "İŞYERİ", "Sicil"]
        
        best_text = ""
        
        for angle in angles:
            try:
                if angle == 0:
                    rotated_image = image
                else:
                    rotated_image = image.rotate(angle, expand=True)
                
                text = pytesseract.image_to_string(rotated_image, lang='tur', config='--psm 6')
                
                # İlk denemede (0 derece) metni varsayılan olarak ayarla
                if angle == 0:
                    best_text = text

                # Metnin içinde anahtar kelimeler var mı diye kontrol et
                for key in keywords:
                    if re.search(key, text, re.IGNORECASE):
                        print(f"Anlamlı metin {angle} derece açıda bulundu.")
                        return text # Anlamlı metin bulununca hemen döndür
            except Exception as e:
                print(f"{angle} derece denenirken hata: {e}")
                continue # Hata olursa bir sonraki açıya geç
        
        print("Anahtar kelime bulunamadı, en iyi tahmin kullanılıyor.")
        return best_text # Hiçbir şey bulunamazsa 0 derecedeki sonucu döndür

    def extract_info_from_pdf(self, pdf_path):
        info = {'TC': 'TC-YOK', 'ADSOYAD': 'ISIM-YOK', 'SICIL': 'SICIL-YOK'}
        text_from_pdf = ""
        try:
            reader = pypdf.PdfReader(pdf_path)
            for page in reader.pages[:2]:
                text_from_pdf += page.extract_text() or ""
        except Exception:
            text_from_pdf = ""
        
        found_info = self.find_info_in_text(text_from_pdf)
        info.update({k: v for k, v in found_info.items() if v})

        if info['ADSOYAD'] == 'ISIM-YOK' or info['TC'] == 'TC-YOK':
            try:
                images = convert_from_path(pdf_path, poppler_path=self.poppler_path, first_page=1, last_page=1, dpi=300)
                if images:
                    # --- GÜNCELLENEN KISIM: Yeni ve kararlı yön bulma fonksiyonu çağrılıyor ---
                    text_from_ocr = self.ocr_with_orientation_check(images[0])
                    print(f"--- OCR Sonucu ({os.path.basename(pdf_path)}): ---\n{text_from_ocr}\n--------------------")
                    
                    found_info_ocr = self.find_info_in_text(text_from_ocr)
                    if info['ADSOYAD'] == 'ISIM-YOK' and found_info_ocr.get('ADSOYAD'):
                        info['ADSOYAD'] = found_info_ocr['ADSOYAD']
                    if info['TC'] == 'TC-YOK' and found_info_ocr.get('TC'):
                        info['TC'] = found_info_ocr['TC']
                    if info['SICIL'] == 'SICIL-YOK' and found_info_ocr.get('SICIL'):
                        info['SICIL'] = found_info_ocr['SICIL']
            except Exception as e:
                print(f"OCR Hatası ({os.path.basename(pdf_path)}): {e}")
        return info

    def start_rename_thread(self):
        # (Bu fonksiyon değişmedi)
        self.rename_button.configure(state="disabled")
        self.progress_bar.set(0)
        thread = Thread(target=self.rename_process)
        thread.start()

    def rename_process(self):
        # (Bu fonksiyon değişmedi)
        name_format = self.format_combo.get()
        total_files = len(self.selected_files)
        processed_count = 0
        for i, file_path in enumerate(self.selected_files):
            self.status_label.configure(text=f"İşleniyor: {os.path.basename(file_path)}")
            extracted_data = self.extract_info_from_pdf(file_path)
            
            try:
                new_name = name_format.format(**extracted_data)
                new_name = re.sub(r'[\\/*?:"<>|]', "", new_name) + ".pdf"
                directory = os.path.dirname(file_path)
                new_file_path = os.path.join(directory, new_name)
                
                if file_path.lower() != new_file_path.lower() and os.path.exists(new_file_path):
                     base, ext = os.path.splitext(new_file_path)
                     counter = 1
                     while os.path.exists(new_file_path):
                         new_file_path = f"{base}_{counter}{ext}"
                         counter += 1
                
                if file_path.lower() != new_file_path.lower():
                    os.rename(file_path, new_file_path)
                processed_count += 1
            except Exception as e:
                self.status_label.configure(text=f"Hata: {os.path.basename(file_path)} adlandırılamadı.")
                print(f"Adlandırma hatası: {e}")
                
            self.progress_bar.set((i + 1) / total_files)
        
        messagebox.showinfo("İşlem Tamamlandı", f"{processed_count} / {total_files} dosya başarıyla yeniden adlandırıldı.")
        self.status_label.configure(text="İşlem tamamlandı. Yeni dosyalar seçebilirsiniz.")
        self.rename_button.configure(state="normal")
        self.file_list_box.configure(state="normal")
        self.file_list_box.delete("1.0", "end")
        self.file_list_box.configure(state="disabled")

if __name__ == "__main__":
    app = ctk.CTk()
    app.title("PDF Yeniden Adlandırma Aracı")
    app.geometry("500x400")
    
    frame = PDFRenamerFrame(app)
    frame.pack(fill="both", expand=True)
    
    app.mainloop()