
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
import re
from utils.backup_manager import BackupManager
from utils.template_manager import TemplateManager
from PIL import Image
import sys

# ocr kütüphanelerini yüklemeye çalışıyoruz
try:
    import pytesseract
    import fitz # PyMuPDF
except ImportError:
    # eğer yüklü değilse none yapıyoruz ki hata vermesin
    pytesseract = None
    fitz = None

# pdf okuma kütüphanesini yüklüyoruz
try:
    from PyPDF2 import PdfReader
except ImportError:
    try:
        from pypdf import PdfReader
    except ImportError:
        PdfReader = None

# pdf yeniden adlandırma aracı sınıfı
class PDFRenamerFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # değişkenleri tanımlıyoruz
        self.patterns = []
        self.selected_files = []
        
        # ocr ayarlarını yapıyoruz
        # eğer uygulama donmuşsa (exe ise) yolu ona göre alıyoruz
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # tesseract yolunu belirtiyoruz
        self.tesseract_path = os.path.join(application_path, 'tesseract', 'tesseract.exe')
        
        # eğer pytesseract yüklüyse yolunu ayarlıyoruz
        if pytesseract:
            try:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            except Exception:
                pass

        # arayüz elemanlarını oluşturuyoruz
        self.create_widgets()
        
    def create_widgets(self):
        # ızgara düzenini ayarlıyoruz
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # --- sol taraf: desen yönetimi ---
        left_frame = ctk.CTkFrame(self)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(2, weight=1) 
        
        # --- üst kısım: mod seçimi ---
        mode_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        mode_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        
        self.mode_var = ctk.StringVar(value="regex")
        
        # mod değiştirme butonları
        ctk.CTkRadioButton(mode_frame, text="Dinamik Desen (Regex)", variable=self.mode_var, value="regex", command=self.toggle_mode).pack(side="left", padx=10, pady=10)
        ctk.CTkRadioButton(mode_frame, text="Görsel Şablon (Template)", variable=self.mode_var, value="template", command=self.toggle_mode).pack(side="left", padx=10, pady=10)

        # --- regex çerçevesi ---
        self.regex_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        self.regex_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(self.regex_frame, text="Desen Ekle (Örn: TC, Ad Soyad):").pack(side="left", padx=5)
        self.pattern_entry = ctk.CTkEntry(self.regex_frame, width=200)
        self.pattern_entry.pack(side="left", padx=5)
        ctk.CTkButton(self.regex_frame, text="Ekle", width=60, command=self.add_pattern).pack(side="left", padx=5)

        # --- şablon çerçevesi ---
        self.template_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        # başlangıçta gizli, mod değişiminde gösterilecek
        
        ctk.CTkLabel(self.template_frame, text="Şablon Seçiniz:").pack(side="left", padx=5)
        self.template_combobox = ctk.CTkComboBox(self.template_frame, values=[])
        self.template_combobox.pack(fill="x", padx=5, pady=5)
        ctk.CTkButton(self.template_frame, text="Yenile", command=self.refresh_templates).pack(fill="x", padx=5, pady=5)
        ctk.CTkButton(self.template_frame, text="Yenile", width=60, command=self.refresh_templates).pack(side="left", padx=5)

        # --- desen listesi (ortak alan) ---
        self.pattern_list_frame = ctk.CTkScrollableFrame(left_frame, label_text="Kullanılabilir Desenler")
        self.pattern_list_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
        
        # --- sağ taraf: dosya ve format işlemleri ---
        right_frame = ctk.CTkFrame(self)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(right_frame, text="Dosya ve Format İşlemleri", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 10))
        
        self.select_files_btn = ctk.CTkButton(right_frame, text="PDF Dosyaları Seç", command=self.select_files)
        self.select_files_btn.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.file_count_label = ctk.CTkLabel(right_frame, text="Seçilen Dosya: 0")
        self.file_count_label.grid(row=2, column=0, padx=20, pady=5)
        
        ctk.CTkLabel(right_frame, text="Yeni Dosya Adı Formatı:", anchor="w").grid(row=3, column=0, padx=20, pady=(15, 5), sticky="w")
        
        self.format_entry = ctk.CTkEntry(right_frame, placeholder_text="Örn: {TC}_{Ad Soyad}")
        self.format_entry.grid(row=4, column=0, padx=20, pady=5, sticky="ew")
        
        ctk.CTkLabel(right_frame, text="* Desen isimlerini süslü parantez içinde kullanın.", font=ctk.CTkFont(size=11), text_color="gray").grid(row=5, column=0, padx=20, pady=0, sticky="w")
        
        self.rename_btn = ctk.CTkButton(right_frame, text="Yeniden Adlandır", command=self.start_rename_thread, state="disabled", fg_color="green")
        self.rename_btn.grid(row=6, column=0, padx=20, pady=20, sticky="ew")
        
        self.status_label = ctk.CTkLabel(right_frame, text="", text_color="gray")
        self.status_label.grid(row=7, column=0, padx=20, pady=10)

    # mod değiştirme fonksiyonu (regex veya şablon)
    def toggle_mode(self):
        mode = self.mode_var.get()
        if mode == "regex":
            # regex moduna geçiş
            self.template_frame.grid_forget()
            self.regex_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
            self.pattern_list_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
            self.refresh_pattern_list() # regex desenlerini geri getir
        else:
            # şablon moduna geçiş
            self.regex_frame.grid_forget()
            self.template_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
            self.pattern_list_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew") # şablon alanlarını göster
            self.refresh_templates()

    # şablonları yenileme fonksiyonu
    def refresh_templates(self):
        templates = TemplateManager.get_all_templates()
        self.template_combobox.configure(values=templates, command=self.on_template_select)
        if templates:
            self.template_combobox.set(templates[0])
            self.on_template_select(templates[0])
        else:
            self.template_combobox.set("")
            self.clear_pattern_list()

    # şablon seçildiğinde çalışacak fonksiyon
    def on_template_select(self, choice):
        if not choice:
            return
        
        try:
            # seçilen şablonun alanlarını yüklüyoruz
            fields = TemplateManager.load_template(choice)
            self.display_template_fields(fields)
        except Exception as e:
            print(f"şablon alanları yüklenirken hata: {e}")

    # desen listesini temizleme
    def clear_pattern_list(self):
        for widget in self.pattern_list_frame.winfo_children():
            widget.destroy()

    # şablon alanlarını ekranda gösterme
    def display_template_fields(self, fields):
        self.clear_pattern_list()
        
        ctk.CTkLabel(self.pattern_list_frame, text="Şablondaki Alanlar (Kopyalamak için tıklayın):", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=5, padx=5, anchor="w")
        
        for field in fields:
            name = field['name']
            # her alan için bir buton oluşturuyoruz, tıklanınca kopyalanacak
            btn = ctk.CTkButton(self.pattern_list_frame, text=f"{{{name}}}", fg_color="gray", hover_color="darkgray",
                              command=lambda n=name: self.copy_to_format(n))
            btn.pack(fill="x", pady=2, padx=5)

    # format girişine metin kopyalama
    def copy_to_format(self, text):
        current = self.format_entry.get()
        self.format_entry.insert(len(current), f"{{{text}}}")
            
    def tkraise(self, aboveThis=None):
        super().tkraise(aboveThis)
        self.refresh_templates()

    # yeni desen ekleme
    def add_pattern(self):
        pattern = self.pattern_entry.get().strip()
        if not pattern:
            return
            
        if pattern in self.patterns:
            messagebox.showwarning("Uyarı", "Bu desen zaten ekli.")
            return
            
        self.patterns.append(pattern)
        self.refresh_pattern_list()
        self.pattern_entry.delete(0, "end")
        
    # desen silme
    def remove_pattern(self, pattern):
        if pattern in self.patterns:
            self.patterns.remove(pattern)
            self.refresh_pattern_list()
            
    # desen listesini yenileme (regex modu için)
    def refresh_pattern_list(self):
        for widget in self.pattern_list_frame.winfo_children():
            widget.destroy()
            
        for pattern in self.patterns:
            row_frame = ctk.CTkFrame(self.pattern_list_frame, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)
            
            lbl = ctk.CTkLabel(row_frame, text=pattern, anchor="w")
            lbl.pack(side="left", padx=5, fill="x", expand=True)
            
            del_btn = ctk.CTkButton(row_frame, text="X", width=30, fg_color="red", hover_color="darkred", 
                                  command=lambda p=pattern: self.remove_pattern(p))
            del_btn.pack(side="right", padx=5)

    # dosya seçme işlemi
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="PDF Dosyaları Seç",
            filetypes=[("PDF Dosyaları", "*.pdf")]
        )
        if files:
            self.selected_files = list(files)
            self.file_count_label.configure(text=f"Seçilen Dosya: {len(self.selected_files)}")
            self.update_rename_button()
            
    # yeniden adlandırma butonunu güncelleme
    def update_rename_button(self):
        if self.selected_files:
            self.rename_btn.configure(state="normal")
        else:
            self.rename_btn.configure(state="disabled")

    # yeniden adlandırma işlemini başlatma (thread içinde)
    def start_rename_thread(self):
        format_str = self.format_entry.get().strip()
        if not format_str:
            messagebox.showwarning("Uyarı", "Lütfen bir dosya adı formatı girin.")
            return
            
        if not self.selected_files:
            return
            
        self.rename_btn.configure(state="disabled")
        self.select_files_btn.configure(state="disabled")
        self.status_label.configure(text="İşlem başlıyor...", text_color="cyan")
        
        thread = threading.Thread(target=self.rename_process, args=(format_str,))
        thread.daemon = True
        thread.start()

    # pdf'den metin çıkarma
    def extract_text_from_pdf(self, pdf_path):
        text = ""
        try:
            # 1. önce pypdf ile deniyoruz
            if PdfReader:
                reader = PdfReader(pdf_path)
                for page in reader.pages:
                    extracted = page.extract_text()
                    if extracted:
                        text += extracted + "\n"
            
            # 2. eğer metin boşsa veya çok kısaysa ocr deniyoruz
            if not text.strip() or len(text) < 50:
                if pytesseract and fitz:
                    print(f"ocr deneniyor: {os.path.basename(pdf_path)}")
                    try:
                        doc = fitz.open(pdf_path)
                        for page in doc:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                            import io
                            img = Image.open(io.BytesIO(pix.tobytes("png")))
                            ocr_text = pytesseract.image_to_string(img, lang='tur+eng')
                            text += ocr_text + "\n"
                    except Exception as ocr_e:
                        print(f"ocr hatası: {ocr_e}")
                else:
                    print("ocr kütüphaneleri eksik, sadece metin çıkarıldı.")
                    
        except Exception as e:
            print(f"pdf okuma hatası ({os.path.basename(pdf_path)}): {e}")
        return text

    # desen için değer bulma
    def find_value_for_pattern(self, text, pattern):
        # 1. deseni temizliyoruz: sonunda iki nokta varsa siliyoruz
        clean_pattern = pattern.strip()
        if clean_pattern.endswith(":"):
            clean_pattern = clean_pattern[:-1].strip()
            
        # 2. esnek regex oluşturuyoruz
        # boşlukları esnek hale getiriyoruz
        escaped_chars = [re.escape(c) for c in clean_pattern]
        flexible_pattern = r"[ \t]*".join(escaped_chars)
        
        # 3. sona opsiyonel iki nokta ve boşluk ekliyoruz
        final_regex_base = fr"{flexible_pattern}[ \t]*:?"
        
        # önce aynı satırda arıyoruz
        match = re.search(fr"{final_regex_base}[ \t]*([^\n]*)", text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            if value: 
                # dosya adında olmaması gereken karakterleri temizliyoruz
                return re.sub(r'[\\/*?:"<>|]', "", value)
        
        # eğer bulunamazsa bir alt satıra bakıyoruz
        regex_multi = fr"{final_regex_base}[ \t]*[\r\n]+\s*([^\n]+)"
        match_multiline = re.search(regex_multi, text, re.IGNORECASE)
        if match_multiline:
            value = match_multiline.group(1).strip()
            return re.sub(r'[\\/*?:"<>|]', "", value)
            
        return ""

    # asıl yeniden adlandırma işlemi
    def rename_process(self, format_str):
        try:
            # 1. dosyaları yedekliyoruz
            self.status_label.configure(text="Yedekleme yapılıyor...")
            BackupManager.create_backup(self.selected_files)
            
            total = len(self.selected_files)
            processed_count = 0
            mode = self.mode_var.get()
            template_name = self.template_combobox.get()
            
            if mode == "template" and not template_name:
                messagebox.showwarning("Uyarı", "Lütfen bir şablon seçin.")
                return

            for i, pdf_path in enumerate(self.selected_files):
                self.status_label.configure(text=f"İşleniyor ({i+1}/{total}): {os.path.basename(pdf_path)}")
                
                values = {}
                
                if mode == "regex":
                    text = self.extract_text_from_pdf(pdf_path)
                    # tüm desenler için değerleri buluyoruz
                    for pattern in self.patterns:
                        val = self.find_value_for_pattern(text, pattern)
                        values[pattern] = val
                        if pattern.endswith(":"):
                            clean_key = pattern[:-1].strip()
                            values[clean_key] = val
                else:
                    # şablon modu
                    try:
                        values = TemplateManager.extract_data_with_template(
                            pdf_path, template_name, 
                            tesseract_path=self.tesseract_path
                        )
                    except Exception as e:
                        print(f"şablon çıkarma hatası {pdf_path}: {e}")
                
                try:
                    # yeni dosya adını oluşturuyoruz
                    new_filename = format_str.format(**values)
                    if not new_filename.strip():
                        raise ValueError("Oluşan dosya adı boş.")
                        
                    new_filename += ".pdf"
                    
                    directory = os.path.dirname(pdf_path)
                    new_path = os.path.join(directory, new_filename)
                    
                    # aynı isimde dosya varsa sonuna numara ekliyoruz
                    if os.path.exists(new_path) and new_path.lower() != pdf_path.lower():
                        base, ext = os.path.splitext(new_filename)
                        counter = 1
                        while os.path.exists(os.path.join(directory, f"{base}_{counter}{ext}")):
                            counter += 1
                        new_path = os.path.join(directory, f"{base}_{counter}{ext}")
                    
                    # dosyayı yeniden adlandırıyoruz
                    if new_path.lower() != pdf_path.lower():
                        os.rename(pdf_path, new_path)
                        processed_count += 1
                        
                except KeyError as e:
                    print(f"format hatası: {e} anahtarı bulunamadı.")
                except Exception as e:
                    print(f"yeniden adlandırma hatası ({pdf_path}): {e}")
            
            self.status_label.configure(text=f"Tamamlandı! {processed_count}/{total} dosya adlandırıldı.", text_color="green")
            messagebox.showinfo("Başarılı", "İşlem tamamlandı.")
            
        except Exception as e:
            self.status_label.configure(text=f"Hata: {str(e)}", text_color="red")
            messagebox.showerror("Hata", f"İşlem sırasında hata: {e}")
            
        finally:
            self.rename_btn.configure(state="normal")
            self.select_files_btn.configure(state="normal")