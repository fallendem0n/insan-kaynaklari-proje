import customtkinter as ctk
from tools.pdf_splitter_tool import PDFSplitterFrame
from tools.pdf_renamer_tool import PDFRenamerFrame
# Yeni modülleri import edin:
from tools.egitim_sertifikasi_tool import EgitimSertifikasiFrame
from tools.pdf_to_txt_tool import PDFToTXTFrame

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Ofis Yardımcısı")
        self.geometry("700x550")

        ctk.set_appearance_mode("Dark")

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True)

        self.tab_view = ctk.CTkTabview(self.main_frame, width=250)
        self.tab_view.pack(padx=20, pady=10, fill="both", expand=True)

        self.tab_view.add("PDF Bölücü")
        self.tab_view.add("PDF Yeniden Adlandır")
        self.tab_view.add("Eğitim Sertifikası Formatı") # Bu sekme adı zaten vardı
        self.tab_view.add("PDF to TXT")               # Bu sekme adı zaten vardı
        self.tab_view.add("Yapım Aşamasında #1")

        # PDF Bölücü
        self.pdf_splitter_frame = PDFSplitterFrame(master=self.tab_view.tab("PDF Bölücü"))
        self.pdf_splitter_frame.pack(fill="both", expand=True)
        
        # PDF Yeniden Adlandır
        self.pdf_renamer_frame = PDFRenamerFrame(master=self.tab_view.tab("PDF Yeniden Adlandır"))
        self.pdf_renamer_frame.pack(fill="both", expand=True)
        
        # --- YENİ EKLENEN KISIMLAR ---
        
        # Eğitim Sertifikası
        self.egitim_sertifikasi_frame = EgitimSertifikasiFrame(master=self.tab_view.tab("Eğitim Sertifikası Formatı"))
        self.egitim_sertifikasi_frame.pack(fill="both", expand=True)
        
        # PDF to TXT
        self.pdf_to_txt_frame = PDFToTXTFrame(master=self.tab_view.tab("PDF to TXT"))
        self.pdf_to_txt_frame.pack(fill="both", expand=True)
        
        # --- YENİ EKLENEN KISIMLARIN SONU ---
        
        self.bottom_frame = ctk.CTkFrame(self.main_frame)
        self.bottom_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))

        self.theme_switch = ctk.CTkSwitch(
            self.bottom_frame, 
            text="Karanlık Mod", 
            command=self.toggle_theme
        )
        self.theme_switch.pack(side="right", padx=10, pady=5)
        
        self.theme_switch.select()

    def toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")

# Bu kısım, eğer main.py yerine doğrudan bu dosyayı çalıştırıyorsanız gereklidir:
# if __name__ == "__main__":
#     app = App()
#     app.mainloop()