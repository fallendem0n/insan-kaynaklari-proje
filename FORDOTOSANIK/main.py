import customtkinter as ctk
from gui_manager import App

# bu kısım programın başlangıç noktası
# eğer bu dosya doğrudan çalıştırılırsa uygulama başlar
if __name__ == "__main__":
    # ana uygulama sınıfını oluşturuyoruz
    app = App()
    # uygulamayı çalıştırıp ekranda kalmasını sağlıyoruz
    app.mainloop()