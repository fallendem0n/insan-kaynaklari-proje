import customtkinter as ctk
import sys
import os
import re
import threading
import pandas as pd
from tkinter import filedialog, messagebox
from utils.template_manager import TemplateManager

try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None

try:
    from PyPDF2 import PdfReader
except ImportError:
    try:
        from pypdf import PdfReader
    except ImportError:
        PdfReader = None

class PDFDataExtractorFrame(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)
        
        self.patterns = []
        self.selected_files = []
        
        # OCR Configuration
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        self.tesseract_path = os.path.join(application_path, 'tesseract', 'tesseract.exe')
        self.poppler_path = os.path.join(application_path, 'poppler', 'Library', 'bin')
        
        if pytesseract:
            try:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            except Exception:
                pass

        self.create_widgets()
        
    def create_widgets(self):
        # Layout configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # --- Left Side: Pattern Management ---
        left_frame = ctk.CTkFrame(self)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(3, weight=1) # Pattern list expands
        
        # Mode Selection
        mode_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        mode_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        
        self.mode_var = ctk.StringVar(value="regex")
        ctk.CTkRadioButton(mode_frame, text="Dinamik Desen (Regex)", variable=self.mode_var, value="regex", command=self.toggle_mode).pack(side="left", padx=10)
        ctk.CTkRadioButton(mode_frame, text="Görsel Şablon", variable=self.mode_var, value="template", command=self.toggle_mode).pack(side="left", padx=10)

        # Regex Frame
        self.regex_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        self.regex_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(self.regex_frame, text="Desen Ekle (Örn: 'Sicil No:'):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=5)
        
        input_frame = ctk.CTkFrame(self.regex_frame, fg_color="transparent")
        input_frame.pack(fill="x", pady=5)
        
        self.pattern_entry = ctk.CTkEntry(input_frame, placeholder_text="Desen girin...")
        self.pattern_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        add_btn = ctk.CTkButton(input_frame, text="Ekle", width=60, command=self.add_pattern)
        add_btn.pack(side="right")

        # Template Frame
        self.template_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        # Don't grid initially
        
        ctk.CTkLabel(self.template_frame, text="Şablon Seçiniz:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=5)
        self.template_combobox = ctk.CTkComboBox(self.template_frame, values=[], command=lambda x: self.update_process_button())
        self.template_combobox.pack(fill="x", padx=5, pady=5)
        ctk.CTkButton(self.template_frame, text="Yenile", command=self.refresh_templates).pack(fill="x", padx=5, pady=5)

        # Pattern List (Only for Regex mode)
        self.pattern_list_frame = ctk.CTkScrollableFrame(left_frame, label_text="Kullanılabilir Desenler")
        self.pattern_list_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        
        # --- Right Side: File Selection & Action ---
        right_frame = ctk.CTkFrame(self)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(right_frame, text="PDF İşlemleri", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 20))
        
        self.select_files_btn = ctk.CTkButton(right_frame, text="PDF Dosyaları Seç", command=self.select_files)
        self.select_files_btn.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.file_count_label = ctk.CTkLabel(right_frame, text="Seçilen Dosya: 0")
        self.file_count_label.grid(row=2, column=0, padx=20, pady=5)
        
        self.process_btn = ctk.CTkButton(right_frame, text="Verileri Çıkar ve Excel'e Kaydet", command=self.start_extraction, state="disabled", fg_color="green")
        self.process_btn.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        
        self.status_label = ctk.CTkLabel(right_frame, text="", text_color="gray")
        self.status_label.grid(row=4, column=0, padx=20, pady=10)
        
        # Instructions
        info_text = (
            "NASIL KULLANILIR?\n\n"
            "1. Sol taraftan PDF içinde aranacak başlıkları ekleyin.\n"
            "   Örn: 'Sicil No:', 'Adı Soyadı:', 'Tutar:'\n\n"
            "2. 'PDF Dosyaları Seç' butonu ile işlem yapılacak dosyaları seçin.\n\n"
            "3. 'Verileri Çıkar...' butonuna basın.\n\n"
            "Program, her PDF için bu başlıkların karşısındaki değerleri\n"
            "bulup tek bir Excel dosyasında birleştirecektir."
        )
        ctk.CTkLabel(right_frame, text=info_text, justify="left", text_color="gray").grid(row=5, column=0, padx=20, pady=20, sticky="w")

    def toggle_mode(self):
        mode = self.mode_var.get()
        if mode == "regex":
            self.template_frame.grid_forget()
            self.regex_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
            self.pattern_list_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        else:
            self.regex_frame.grid_forget()
            self.pattern_list_frame.grid_forget()
            self.template_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
            self.refresh_templates()
        self.update_process_button()

    def refresh_templates(self):
        templates = TemplateManager.get_all_templates()
        self.template_combobox.configure(values=templates)
        if templates:
            self.template_combobox.set(templates[0])
        else:
            self.template_combobox.set("")
            
    def tkraise(self, aboveThis=None):
        super().tkraise(aboveThis)
        self.refresh_templates()

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
        
    def remove_pattern(self, pattern):
        if pattern in self.patterns:
            self.patterns.remove(pattern)
            self.refresh_pattern_list()
            
    def refresh_pattern_list(self):
        # Clear existing widgets
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

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="PDF Dosyaları Seç",
            filetypes=[("PDF Dosyaları", "*.pdf")]
        )
        if files:
            self.selected_files = list(files)
            self.file_count_label.configure(text=f"Seçilen Dosya: {len(self.selected_files)}")
            self.update_process_button()
            
    def update_process_button(self):
        has_files = bool(self.selected_files)
        mode = self.mode_var.get()
        
        has_criteria = False
        if mode == "regex":
            has_criteria = bool(self.patterns)
        else:
            has_criteria = bool(self.template_combobox.get())
            
        if has_files and has_criteria:
            self.process_btn.configure(state="normal")
        else:
            self.process_btn.configure(state="disabled")

    def start_extraction(self):
        if not self.selected_files:
            messagebox.showwarning("Uyarı", "Lütfen en az bir PDF dosyası seçin.")
            return
            
        mode = self.mode_var.get()
        template_name = self.template_combobox.get()
        
        if mode == "regex" and not self.patterns:
            messagebox.showwarning("Uyarı", "Lütfen en az bir desen ekleyin.")
            return
            
        if mode == "template" and not template_name:
            messagebox.showwarning("Uyarı", "Lütfen bir şablon seçin.")
            return
            
        self.status_label.configure(text="İşlem başlıyor...")
        
        # Run in thread
        threading.Thread(target=self.extraction_process, args=(mode, template_name), daemon=True).start()

    def extraction_process(self, mode, template_name):
        try:
            all_data = []
            total = len(self.selected_files)
            
            # Get columns
            if mode == "regex":
                columns = ["Dosya Adı"] + self.patterns
            else:
                # Load template fields to get columns
                fields = TemplateManager.load_template(template_name)
                columns = ["Dosya Adı"] + [f['name'] for f in fields]
            
            for i, pdf_path in enumerate(self.selected_files):
                self.status_label.configure(text=f"İşleniyor ({i+1}/{total}): {os.path.basename(pdf_path)}")
                
                row_data = {"Dosya Adı": os.path.basename(pdf_path)}
                
                if mode == "regex":
                    text = self.extract_text_from_pdf(pdf_path)
                    for pattern in self.patterns:
                        val = self.find_value_for_pattern(text, pattern)
                        row_data[pattern] = val
                else:
                    # Template Mode
                    try:
                        extracted = TemplateManager.extract_data_with_template(
                            pdf_path, template_name,
                            poppler_path=self.poppler_path,
                            tesseract_path=self.tesseract_path
                        )
                        row_data.update(extracted)
                    except Exception as e:
                        print(f"Template extraction error: {e}")
                
                all_data.append(row_data)
            
            # Create DataFrame
            df = pd.DataFrame(all_data, columns=columns)
            
            # Save to Excel
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Dosyası", "*.xlsx")])
            if save_path:
                df.to_excel(save_path, index=False)
                self.status_label.configure(text="Tamamlandı!", text_color="green")
                messagebox.showinfo("Başarılı", f"Veriler kaydedildi:\n{save_path}")
            else:
                self.status_label.configure(text="İptal edildi.", text_color="orange")
                
        except Exception as e:
            self.status_label.configure(text=f"Hata: {str(e)}", text_color="red")
            messagebox.showerror("Hata", f"İşlem sırasında bir hata oluştu:\n{e}")
            print(f"Process error: {e}")
            
        finally:
            self.process_btn.configure(state="normal")
            self.select_files_btn.configure(state="normal")

    def extract_text_from_pdf(self, pdf_path):
        text = ""
        try:
            # 1. Try pypdf first
            if PdfReader:
                reader = PdfReader(pdf_path)
                for page in reader.pages:
                    extracted = page.extract_text()
                    if extracted:
                        text += extracted + "\n"
            
            # 2. If text is empty or too short, try OCR
            if not text.strip() or len(text) < 50:
                if pytesseract and convert_from_path:
                    print(f"OCR deneniyor: {os.path.basename(pdf_path)}")
                    try:
                        images = convert_from_path(pdf_path, poppler_path=self.poppler_path)
                        for img in images:
                            ocr_text = pytesseract.image_to_string(img, lang='tur+eng')
                            text += ocr_text + "\n"
                    except Exception as ocr_e:
                        print(f"OCR hatası: {ocr_e}")
                else:
                    print("OCR kütüphaneleri eksik, sadece metin çıkarıldı.")
                    
        except Exception as e:
            print(f"PDF okuma hatası ({os.path.basename(pdf_path)}): {e}")
        return text

    def find_value_for_pattern(self, text, pattern):
        # 1. Clean up the pattern: remove trailing colon if exists, we will add it optionally in regex
        clean_pattern = pattern.strip()
        if clean_pattern.endswith(":"):
            clean_pattern = clean_pattern[:-1].strip()
            
        # 2. Build flexible regex
        # Escape special chars first
        escaped_chars = [re.escape(c) for c in clean_pattern]
        # Join with [ \t]* to allow whitespace between any characters (e.g. "S i c i l")
        flexible_pattern = r"[ \t]*".join(escaped_chars)
        
        # 3. Add optional colon and whitespace at the end
        # Pattern + optional whitespace + optional colon + optional whitespace
        final_regex_base = fr"{flexible_pattern}[ \t]*:?"
        
        # Try finding on the same line first
        # Look for: Pattern + (optional colon) + whitespace + (Value)
        match = re.search(fr"{final_regex_base}[ \t]*([^\n]*)", text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            if value:
                return value
        
        # If not found or empty, try looking at the next line
        match_multiline = re.search(fr"{final_regex_base}[ \t]*[\r\n]+\s*([^\n]+)", text, re.IGNORECASE)
        if match_multiline:
            value = match_multiline.group(1).strip()
            return value
            
        return ""
