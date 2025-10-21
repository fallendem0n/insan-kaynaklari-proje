import sys
import traceback

print(f"--- Hata Ayıklama Başlangıcı ---")
print(f"Python Yolu (sys.path): {sys.path}")
print(f"Python Sürümü: {sys.version}")

try:
    print("pandas import deneniyor...")
    import pandas as pd
    print("pandas başarıyla import edildi.")
except ImportError as e:
    print(f"HATA: pandas import edilemedi!")
    print(f"Detay: {e}")
    traceback.print_exc() # Daha detaylı hata çıktısı için
except Exception as e:
    print(f"BEKLENMEDİK HATA (pandas): {e}")
    traceback.print_exc()

try:
    print("openpyxl import deneniyor...")
    import openpyxl # openpyxl'i kullanmasak da Excel okumak için pandas'ın ihtiyacı var
    print("openpyxl başarıyla import edildi.")
except ImportError as e:
    print(f"HATA: openpyxl import edilemedi!")
    print(f"Detay: {e}")
    traceback.print_exc()
except Exception as e:
    print(f"BEKLENMEDİK HATA (openpyxl): {e}")
    traceback.print_exc()

try:
    print("docx_mailmerge import deneniyor...")
    from docx_mailmerge import MailMerge
    print("docx_mailmerge başarıyla import edildi.")
except ImportError as e:
    print(f"HATA: docx_mailmerge import edilemedi!")
    print(f"Detay: {e}")
    traceback.print_exc() # Daha detaylı hata çıktısı için
except Exception as e:
    print(f"BEKLENMEDİK HATA (docx_mailmerge): {e}")
    traceback.print_exc()

print(f"--- Hata Ayıklama Sonu ---")

# --- Orijinal kodunuzun geri kalanı buradan devam ediyor ---
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os

class MailMergerFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        
        self.excel_path = None
        self.word_path = None
        self.excel_columns = []
        self.word_fields = []

        self.create_widgets()

    def create_widgets(self):
        # --- BİLGİLENDİRME KUTUSU ---
        self.info_textbox = ctk.CTkTextbox(self, height=100, wrap="word")
        self.info_textbox.pack(padx=10, pady=(10, 5), fill="x", expand=False)

        bilgi_metni = """
Bu araç, bir Excel veri dosyasını bir Word şablonuyla birleştirir (Mail Merge).

1. Excel veri kaynağınızı seçin.
2. `.docx` formatındaki Word şablonunuzu seçin.
3. **ÖNEMLİ:** Word şablonunuzdaki birleştirme alanları `{{Alan_Adi}}` şeklinde olmalıdır. (Çift süslü parantez)
4. Excel dosyanızdaki sütun başlığı (`Alan_Adi`) ile Word'deki alan adı birebir aynı olmalıdır.
5. Araç, eşleşmelere göre Excel'deki her satır için ayrı bir Word belgesi oluşturur.
        """
        self.info_textbox.configure(state="normal")
        self.info_textbox.insert("1.0", bilgi_metni.strip())
        self.info_textbox.configure(state="disabled")
        # --- BİLGİLENDİRME KUTUSU SONU ---

        main_frame = ctk.CTkFrame(self, corner_radius=10)
        main_frame.pack(pady=10, padx=20, fill="both", expand=True)

        title_label = ctk.CTkLabel(main_frame, text="Mail Merge Oluşturucu", font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(10, 10))

        # --- EXCEL DOSYA SEÇİMİ ---
        excel_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        excel_frame.pack(pady=5, padx=10, fill="x")
        
        excel_button = ctk.CTkButton(excel_frame, text="1. Excel Veri Dosyası Seç", width=180, command=self.select_excel)
        excel_button.pack(side="left", padx=(0, 10))
        
        self.excel_label = ctk.CTkLabel(excel_frame, text="Veri dosyası seçilmedi...", text_color="gray")
        self.excel_label.pack(side="left", fill="x", expand=True)

        # --- WORD DOSYA SEÇİMİ ---
        word_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        word_frame.pack(pady=5, padx=10, fill="x")
        
        word_button = ctk.CTkButton(word_frame, text="2. Word Şablon Dosyası Seç", width=180, command=self.select_word)
        word_button.pack(side="left", padx=(0, 10))
        
        self.word_label = ctk.CTkLabel(word_frame, text="Şablon dosyası seçilmedi...", text_color="gray")
        self.word_label.pack(side="left", fill="x", expand=True)
        
        # --- KONTROL BUTONU ---
        self.check_files_button = ctk.CTkButton(main_frame, text="3. Dosyaları Kontrol Et ve Alanları Eşleştir", command=self.check_files, state="disabled")
        self.check_files_button.pack(pady=(10, 5), padx=10, fill="x")

        # --- ALAN GÖSTERGE ÇERÇEVESİ ---
        fields_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        fields_frame.pack(pady=5, padx=10, fill="both", expand=True)
        
        fields_frame.grid_columnconfigure(0, weight=1)
        fields_frame.grid_columnconfigure(1, weight=1)
        fields_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(fields_frame, text="Excel Sütunları").grid(row=0, column=0, padx=5, pady=(0, 5))
        ctk.CTkLabel(fields_frame, text="Word Birleştirme Alanları ({{...}})").grid(row=0, column=1, padx=5, pady=(0, 5))

        self.excel_cols_textbox = ctk.CTkTextbox(fields_frame, wrap="none", height=100)
        self.excel_cols_textbox.grid(row=1, column=0, padx=(0, 5), sticky="nsew")
        
        self.word_fields_textbox = ctk.CTkTextbox(fields_frame, wrap="none", height=100)
        self.word_fields_textbox.grid(row=1, column=1, padx=(5, 0), sticky="nsew")

        self.excel_cols_textbox.configure(state="disabled")
        self.word_fields_textbox.configure(state="disabled")

        # --- İŞLEM BUTONU VE DURUM ---
        self.process_button = ctk.CTkButton(main_frame, text="4. Birleştirme İşlemini Başlat", height=40, command=self.process_merge, state="disabled")
        self.process_button.pack(pady=(10, 10), padx=10, fill="x")

        self.status_label = ctk.CTkLabel(main_frame, text="", font=ctk.CTkFont(size=12))
        self.status_label.pack(pady=(0, 10))

    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="Excel Veri Dosyasını Seçin",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path = file_path
            self.excel_label.configure(text=os.path.basename(file_path), text_color="white")
            self.update_check_button_state()
            self.process_button.configure(state="disabled") # Yeni dosya seçildi, tekrar kontrol edilmeli

    def select_word(self):
        file_path = filedialog.askopenfilename(
            title="Word Şablon Dosyasını Seçin",
            filetypes=[("Word Belgeleri", "*.docx")]
        )
        if file_path:
            self.word_path = file_path
            self.word_label.configure(text=os.path.basename(file_path), text_color="white")
            self.update_check_button_state()
            self.process_button.configure(state="disabled") # Yeni dosya seçildi, tekrar kontrol edilmeli

    def update_check_button_state(self):
        if self.excel_path and self.word_path:
            self.check_files_button.configure(state="normal")
        else:
            self.check_files_button.configure(state="disabled")

    def check_files(self):
        self.status_label.configure(text="Dosyalar kontrol ediliyor...", text_color="yellow")
        self.process_button.configure(state="disabled")
        try:
            # Excel Sütunlarını Oku
            df = pd.read_excel(self.excel_path, nrows=0) # Sadece başlıkları oku
            self.excel_columns = sorted(list(df.columns))
            
            self.excel_cols_textbox.configure(state="normal")
            self.excel_cols_textbox.delete("1.0", "end")
            self.excel_cols_textbox.insert("1.0", "\n".join(self.excel_columns))
            self.excel_cols_textbox.configure(state="disabled")

            # Word Alanlarını Oku
            doc = MailMerge(self.word_path)
            self.word_fields = sorted(list(doc.get_merge_fields()))
            
            self.word_fields_textbox.configure(state="normal")
            self.word_fields_textbox.delete("1.0", "end")
            self.word_fields_textbox.insert("1.0", "\n".join(self.word_fields))
            self.word_fields_textbox.configure(state="disabled")

            # Eşleşmeleri Kontrol Et
            if not self.word_fields:
                self.status_label.configure(text="HATA: Word şablonunda {{alan_adi}} gibi bir alan bulunamadı.", text_color="red")
                messagebox.showerror("Hata", "Word şablonunda hiç birleştirme alanı ({{alan_adi}} gibi) bulunamadı.")
                return

            excel_set = set(self.excel_columns)
            word_set = set(self.word_fields)
            matches = excel_set.intersection(word_set)
            missing_in_excel = word_set - excel_set

            if not matches:
                self.status_label.configure(text="HATA: Word alanları ile Excel sütunları arasında eşleşme yok.", text_color="red")
                messagebox.showwarning("Eşleşme Yok", "Word şablonundaki alanlar ile Excel'deki sütun başlıkları arasında hiçbir eşleşme bulunamadı.\nLütfen dosyaları kontrol edin.")
                return

            if missing_in_excel:
                self.status_label.configure(text=f"UYARI: Bazı Word alanları Excel'de yok. ({len(missing_in_excel)} adet)", text_color="orange")
                messagebox.showwarning("Eksik Sütunlar", f"Word şablonundaki şu alanlar Excel dosyasında bulunamadı:\n\n{', '.join(missing_in_excel)}\n\nBu alanlar boş bırakılacaktır.")
            
            self.status_label.configure(text=f"Kontrol başarılı. {len(matches)} adet alan eşleşti. İşlemi başlatabilirsiniz.", text_color="lightgreen")
            self.process_button.configure(state="normal")

        except Exception as e:
            self.status_label.configure(text=f"Bir hata oluştu: {e}", text_color="red")
            messagebox.showerror("Hata", f"Dosyalar okunurken bir hata oluştu:\n{e}")

    def process_merge(self):
        if not self.excel_path or not self.word_path:
            messagebox.showerror("Hata", "Lütfen önce Excel ve Word dosyalarını seçin.")
            return

        if not self.excel_columns or not self.word_fields:
            messagebox.showerror("Hata", "Lütfen önce 'Dosyaları Kontrol Et' butonuna basın.")
            return

        # Çıktı klasörünü seçtir
        output_dir = filedialog.askdirectory(title="Birleştirilmiş belgeler nereye kaydedilsin?")
        if not output_dir:
            self.status_label.configure(text="İşlem iptal edildi.", text_color="gray")
            return

        try:
            self.status_label.configure(text="Excel verisi okunuyor...", text_color="yellow")
            self.update_idletasks()
            
            # Excel verisinin tamamını oku
            df = pd.read_excel(self.excel_path)
            # DataFrame'i docx-mailmerge'in istediği formata (dict listesi) çevir
            # Not: Excel'deki 'NaN' (boş) değerleri None (Python boş değeri) ile değiştiriyoruz.
            data_to_merge = df.where(pd.notnull(df), None).to_dict('records')

            self.status_label.configure(text=f"İşlem başladı... {len(data_to_merge)} adet belge oluşturuluyor...", text_color="yellow")
            self.update_idletasks()

            word_base_name = os.path.splitext(os.path.basename(self.word_path))[0]

            # Her bir satır için döngü başlat
            for i, row_data in enumerate(data_to_merge):
                # Şablonu her seferinde yeniden aç
                document = MailMerge(self.word_path)
                
                # Veriyi şablonla birleştir
                # Not: Sadece eşleşen alanları değil, tüm satırı gönderiyoruz.
                # Kütüphane, Word'de bulamadığı verileri göz ardı eder.
                document.merge(**row_data)

                # Çıktı dosyasını isimlendir
                # İlk sütundaki veriyi dosya adı yapmak riskli olabilir (geçersiz karakterler içerebilir)
                # Bu yüzden sıra numarası kullanalım:
                output_filename = f"{word_base_name}_kayit_{i+1:03d}.docx" # Örn: Sablon_kayit_001.docx
                output_filepath = os.path.join(output_dir, output_filename)
                
                # Dosyayı kaydet
                document.write(output_filepath)
                document.close()
            
            self.status_label.configure(text=f"İşlem tamamlandı! {len(data_to_merge)} belge oluşturuldu.", text_color="lightgreen")
            messagebox.showinfo("Başarılı", f"İşlem tamamlandı.\n\n'{output_dir}' klasörüne {len(data_to_merge)} adet Word belgesi oluşturuldu.")

        except Exception as e:
            self.status_label.configure(text=f"Bir hata oluştu: {e}", text_color="red")
            messagebox.showerror("Hata", f"Birleştirme sırasında bir hata oluştu:\n{e}")