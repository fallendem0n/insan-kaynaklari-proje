import customtkinter as ctk
import os
import json # json modülünü import edin
from tools.pdf_splitter_tool import PDFSplitterFrame
from tools.pdf_renamer_tool import PDFRenamerFrame
from tools.egitim_sertifikasi_tool import EgitimSertifikasiFrame
from tools.pdf_to_txt_tool import PDFToTXTFrame

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- TEMA YÜKLEME (GÜNCELLENMİŞ KISIM) ---
        theme_path = None # Başlangıçta None yapalım
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            theme_path = os.path.join(script_dir, "modern_theme.json")

            # 1. Dosyanın var olup olmadığını ve boş olup olmadığını kontrol et
            if not os.path.exists(theme_path):
                print(f"HATA: Tema dosyası bulunamadı: {theme_path}")
                theme_path = None # Hata durumunda yolu sıfırla
            elif os.path.getsize(theme_path) == 0:
                print(f"HATA: Tema dosyası boş: {theme_path}")
                theme_path = None # Hata durumunda yolu sıfırla
            else:
                # 2. Dosyayı Python'un json modülü ile okumayı dene (UTF-8 olarak)
                try:
                    with open(theme_path, 'r', encoding='utf-8') as f:
                        theme_data = json.load(f)
                        print(f"Tema dosyası başarıyla okundu: {theme_path}")
                        # (Opsiyonel) Okunan verinin küçük bir kısmını yazdırabiliriz:
                        # print("Okunan tema verisi (başlangıç):", str(theme_data)[:100])
                except json.JSONDecodeError as e:
                    print(f"HATA: Tema dosyası okunurken JSON hatası oluştu: {theme_path}")
                    print(f"Hata Detayı: {e}")
                    theme_path = None # Hata durumunda yolu sıfırla
                except Exception as e:
                    print(f"HATA: Tema dosyası okunurken genel bir hata oluştu: {theme_path}")
                    print(f"Hata Detayı: {e}")
                    theme_path = None # Hata durumunda yolu sıfırla

            # 3. CustomTkinter'a temayı ayarla (eğer bir hata oluşmadıysa)
            if theme_path:
                ctk.set_default_color_theme(theme_path)
                print("CustomTkinter teması ayarlandı.")
            else:
                print("Özel tema yüklenemediği için varsayılan 'blue' teması kullanılıyor.")
                ctk.set_default_color_theme("blue")

        except Exception as e:
            print(f"Tema yükleme sırasında beklenmedik bir hata oluştu: {e}")
            print("Varsayılan 'blue' teması kullanılıyor.")
            ctk.set_default_color_theme("blue")
        # --- GÜNCELLENMİŞ KISIM SONU ---


        self.title("Ofis Yardımcısı")
        self.geometry("700x575")

        # Görünüm modu (Tema renginden ayrıdır)
        ctk.set_appearance_mode("Dark") # Veya "Light"

        # --- ARAYÜZ OLUŞTURMA ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True)

        self.tab_view = ctk.CTkTabview(self.main_frame, width=250)
        self.tab_view.pack(padx=20, pady=10, fill="both", expand=True)

        self.tab_view.add("PDF Bölücü")
        self.tab_view.add("PDF Yeniden Adlandır")
        self.tab_view.add("Eğitim Sertifikası Formatı")
        self.tab_view.add("PDF to TXT")
        self.tab_view.add("Yapım Aşamasında #1")

        # Sekme içeriklerini oluştur ve yerleştir
        self.pdf_splitter_frame = PDFSplitterFrame(master=self.tab_view.tab("PDF Bölücü"))
        self.pdf_splitter_frame.pack(fill="both", expand=True)

        self.pdf_renamer_frame = PDFRenamerFrame(master=self.tab_view.tab("PDF Yeniden Adlandır"))
        self.pdf_renamer_frame.pack(fill="both", expand=True)

        self.egitim_sertifikasi_frame = EgitimSertifikasiFrame(master=self.tab_view.tab("Eğitim Sertifikası Formatı"))
        self.egitim_sertifikasi_frame.pack(fill="both", expand=True)

        self.pdf_to_txt_frame = PDFToTXTFrame(master=self.tab_view.tab("PDF to TXT"))
        self.pdf_to_txt_frame.pack(fill="both", expand=True)

        # --- ALT KISIM (TEMA DEĞİŞTİRME BUTONU) ---
        self.bottom_frame = ctk.CTkFrame(self.main_frame)
        self.bottom_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))

        self.theme_switch = ctk.CTkSwitch(
            self.bottom_frame,
            text="Karanlık Mod",
            command=self.toggle_theme
        )
        self.theme_switch.pack(side="right", padx=10, pady=5)

        # Başlangıçta tema anahtarını mevcut moda göre ayarla
        if ctk.get_appearance_mode() == "Dark":
            self.theme_switch.select()
        else:
            self.theme_switch.deselect()

    def toggle_theme(self):
        # Görünüm modunu değiştir
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")

# Uygulamayı başlatmak için (genellikle main.py içinde yapılır)
# if __name__ == "__main__":
#     app = App()
#     app.mainloop()