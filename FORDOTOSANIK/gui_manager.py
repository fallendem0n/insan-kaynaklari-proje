import customtkinter as ctk
import os
import json 
from tools.pdf_splitter_tool import PDFSplitterFrame
from tools.pdf_renamer_tool import PDFRenamerFrame
from tools.egitim_sertifikasi_tool import EgitimSertifikasiFrame
from tools.pdf_to_txt_tool import PDFToTXTFrame
from tools.pdf_to_txt_tool import PDFToTXTFrame
from tools.mail_merger_tool import MailMergerFrame  

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        theme_path = None 
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            theme_path = os.path.join(script_dir, "modern_theme.json")

            if not os.path.exists(theme_path):
                print(f"HATA: Tema dosyası bulunamadı: {theme_path}")
                theme_path = None 
            elif os.path.getsize(theme_path) == 0:
                print(f"HATA: Tema dosyası boş: {theme_path}")
                theme_path = None 
            else:
                try:
                    with open(theme_path, 'r', encoding='utf-8') as f:
                        theme_data = json.load(f)
                        print(f"Tema dosyası başarıyla okundu: {theme_path}")
                except json.JSONDecodeError as e:
                    print(f"HATA: Tema dosyası okunurken JSON hatası oluştu: {theme_path}")
                    print(f"Hata Detayı: {e}")
                    theme_path = None 
                except Exception as e:
                    print(f"HATA: Tema dosyası okunurken genel bir hata oluştu: {theme_path}")
                    print(f"Hata Detayı: {e}")
                    theme_path = None 

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


        self.title("Ofis Yardımcısı")
        self.geometry("1000x775")

        ctk.set_appearance_mode("Dark")

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True)

        self.tab_view = ctk.CTkTabview(self.main_frame, width=250)
        self.tab_view.pack(padx=20, pady=10, fill="both", expand=True)

        self.tab_view.add("PDF Bölücü")
        self.tab_view.add("PDF Yeniden Adlandır")
        self.tab_view.add("Eğitim Sertifikası Formatı")
        self.tab_view.add("PDF to TXT")
        self.tab_view.add("Merge Oluşturucu")

        self.pdf_splitter_frame = PDFSplitterFrame(master=self.tab_view.tab("PDF Bölücü"))
        self.pdf_splitter_frame.pack(fill="both", expand=True)

        self.pdf_renamer_frame = PDFRenamerFrame(master=self.tab_view.tab("PDF Yeniden Adlandır"))
        self.pdf_renamer_frame.pack(fill="both", expand=True)

        self.egitim_sertifikasi_frame = EgitimSertifikasiFrame(master=self.tab_view.tab("Eğitim Sertifikası Formatı"))
        self.egitim_sertifikasi_frame.pack(fill="both", expand=True)

        self.pdf_to_txt_frame = PDFToTXTFrame(master=self.tab_view.tab("PDF to TXT"))
        self.pdf_to_txt_frame.pack(fill="both", expand=True)

        self.mail_merger_frame = MailMergerFrame(master=self.tab_view.tab("Merge Oluşturucu"))
        self.mail_merger_frame.pack(fill="both", expand=True)

        self.bottom_frame = ctk.CTkFrame(self.main_frame)
        self.bottom_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))

        self.theme_switch = ctk.CTkSwitch(
            self.bottom_frame,
            text="Karanlık Mod",
            command=self.toggle_theme
        )
        self.theme_switch.pack(side="right", padx=10, pady=5)

        if ctk.get_appearance_mode() == "Dark":
            self.theme_switch.select()
        else:
            self.theme_switch.deselect()

    def toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")
