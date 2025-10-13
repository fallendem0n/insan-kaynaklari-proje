import customtkinter as ctk
from tkinter import filedialog, messagebox
import pypdf
import os

class PDFSplitterFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.selected_pdf_path = None
        self.create_widgets()

    def create_widgets(self):
        main_frame = ctk.CTkFrame(self, corner_radius=10)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        title_label = ctk.CTkLabel(main_frame, text="PDF Sayfa Bölme Aracı", font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(10, 20))
        file_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        file_frame.pack(pady=10, padx=10, fill="x")
        self.file_path_label = ctk.CTkLabel(file_frame, text="Lütfen bir PDF dosyası seçin...", text_color="gray")
        self.file_path_label.pack(side="left", padx=(0, 10), expand=True, fill="x")
        select_button = ctk.CTkButton(file_frame, text="Dosya Seç", width=120, command=self.select_pdf)
        select_button.pack(side="right")
        settings_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        settings_frame.pack(pady=20, padx=10, fill="x")
        pages_label = ctk.CTkLabel(settings_frame, text="Her bir dosyada kaç sayfa olsun?")
        pages_label.pack(side="left")
        self.pages_entry = ctk.CTkEntry(settings_frame, width=80, justify="center")
        self.pages_entry.pack(side="left", padx=10)
        self.pages_entry.insert(0, "1")
        self.process_button = ctk.CTkButton(main_frame, text="PDF'i Böl ve Kaydet", height=40, command=self.process_pdf, state="disabled")
        self.process_button.pack(pady=20, padx=10, fill="x")
        self.status_label = ctk.CTkLabel(main_frame, text="", font=ctk.CTkFont(size=12))
        self.status_label.pack(pady=10)

    def select_pdf(self):
        file_path = filedialog.askopenfilename(title="Bir PDF Dosyası Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if file_path:
            self.selected_pdf_path = file_path
            file_name = os.path.basename(file_path)
            self.file_path_label.configure(text=file_name, text_color="white")
            self.process_button.configure(state="normal")
            self.status_label.configure(text="")

    def process_pdf(self):
        if not self.selected_pdf_path:
            messagebox.showerror("Hata", "Lütfen önce bir PDF dosyası seçin.")
            return
        try:
            pages_per_file = int(self.pages_entry.get())
            if pages_per_file <= 0: raise ValueError
        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli bir sayfa sayısı girin (0'dan büyük bir tam sayı).")
            return
        try:
            self.status_label.configure(text="PDF işleniyor, lütfen bekleyin...", text_color="yellow")
            self.update_idletasks()
            reader = pypdf.PdfReader(self.selected_pdf_path)
            total_pages = len(reader.pages)
            output_dir = filedialog.askdirectory(title="Bölünen dosyaları nereye kaydetmek istersiniz?")
            if not output_dir:
                self.status_label.configure(text="İşlem iptal edildi.", text_color="gray")
                return
            base_name, _ = os.path.splitext(os.path.basename(self.selected_pdf_path))
            for i in range(0, total_pages, pages_per_file):
                writer = pypdf.PdfWriter()
                end_page = min(i + pages_per_file, total_pages)
                for page_num in range(i, end_page):
                    writer.add_page(reader.pages[page_num])
                output_filename = f"{base_name}_sayfa_{i+1}-{end_page}.pdf"
                output_filepath = os.path.join(output_dir, output_filename)
                with open(output_filepath, 'wb') as output_file:
                    writer.write(output_file)
            self.status_label.configure(text=f"İşlem tamamlandı! {total_pages // pages_per_file + 1} dosya oluşturuldu.", text_color="lightgreen")
            messagebox.showinfo("Başarılı", f"PDF başarıyla bölündü ve dosyalar '{output_dir}' klasörüne kaydedildi.")
        except Exception as e:
            self.status_label.configure(text=f"Bir hata oluştu: {e}", text_color="red")
            messagebox.showerror("Hata", f"PDF işlenirken bir hata oluştu:\n{e}")