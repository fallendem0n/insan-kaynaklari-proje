import customtkinter as ctk
from tkinter import filedialog, messagebox, Canvas
import os
import sys
from PIL import Image, ImageTk
from utils.template_manager import TemplateManager

# ocr kütüphanelerini yüklemeye çalışıyoruz
try:
    import pytesseract
    import fitz # PyMuPDF
except ImportError:
    # eğer yüklü değilse none yapıyoruz
    pytesseract = None
    fitz = None

# görsel şablon oluşturma aracı sınıfı
class VisualTemplateFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # değişkenleri tanımlıyoruz
        self.pdf_path = None
        self.image = None
        self.tk_image = None
        self.rect_start_x = None
        self.rect_start_y = None
        self.current_rect = None
        self.fields = [] # alan listesi: {'name', 'rect_id', 'x', 'y', 'w', 'h'}
        self.scale_factor = 1.0
        
        # yolları ayarlıyoruz
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        self.tesseract_path = os.path.join(application_path, 'tesseract', 'tesseract.exe')
        
        # tesseract yolunu ayarlıyoruz
        if pytesseract:
            try:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            except Exception:
                pass

        # arayüzü oluşturuyoruz
        self.create_widgets()

    def create_widgets(self):
        # düzen: sol taraf (kontroller), sağ taraf (tuval)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # --- sol menü ---
        sidebar = ctk.CTkFrame(self, width=250)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        sidebar.grid_propagate(False)
        
        ctk.CTkLabel(sidebar, text="Şablon Oluşturucu", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        ctk.CTkButton(sidebar, text="PDF Yükle", command=self.load_pdf).pack(pady=10, padx=10, fill="x")
        
        self.template_name_entry = ctk.CTkEntry(sidebar, placeholder_text="Şablon Adı")
        self.template_name_entry.pack(pady=5, padx=10, fill="x")
        
        ctk.CTkButton(sidebar, text="Seçili Alanı Test Et (OCR)", command=self.test_ocr, fg_color="orange").pack(pady=5, padx=10, fill="x")
        
        ctk.CTkButton(sidebar, text="Şablonu Kaydet", command=self.save_template, fg_color="green").pack(pady=10, padx=10, fill="x")
        
        ctk.CTkLabel(sidebar, text="Alanlar Listesi:", anchor="w").pack(pady=(20, 5), padx=10, fill="x")
        
        self.fields_frame = ctk.CTkScrollableFrame(sidebar)
        self.fields_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        ctk.CTkLabel(sidebar, text="Nasıl Kullanılır?", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 5))
        info_text = "1. PDF yükleyin.\n2. Mouse ile alan seçin.\n3. Alana isim verin.\n4. Şablonu kaydedin."
        ctk.CTkLabel(sidebar, text=info_text, justify="left", text_color="gray", font=ctk.CTkFont(size=11)).pack(pady=5, padx=10)

        # --- tuval alanı ---
        canvas_frame = ctk.CTkFrame(self)
        canvas_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        # kaydırma çubukları
        v_scroll = ctk.CTkScrollbar(canvas_frame, orientation="vertical")
        h_scroll = ctk.CTkScrollbar(canvas_frame, orientation="horizontal")
        
        self.canvas = Canvas(canvas_frame, bg="#333333", highlightthickness=0, 
                             yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        v_scroll.configure(command=self.canvas.yview)
        h_scroll.configure(command=self.canvas.xview)
        
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # fare olaylarını bağlıyoruz
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

    # pdf yükleme işlemi
    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Dosyaları", "*.pdf")])
        if not file_path:
            return
            
        try:
            if not fitz:
                messagebox.showerror("Hata", "PyMuPDF (fitz) kütüphanesi eksik.")
                return

            # pdf'in ilk sayfasını resme çeviriyoruz
            doc = fitz.open(file_path)
            page = doc.load_page(0) # ilk sayfa
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # 2x zoom ile daha net görüntü
            
            # pixmap'i PIL Image'a çeviriyoruz
            img_data = pix.tobytes("png")
            import io
            self.image = Image.open(io.BytesIO(img_data))
            
            if self.image:
                self.pdf_path = file_path
                self.display_image()
                self.fields = []
                self.refresh_fields_list()
                self.canvas.delete("all")
                self.display_image() # resmi tekrar çiziyoruz
        except Exception as e:
            messagebox.showerror("Hata", f"PDF yüklenirken hata: {e}")

    # resmi ekranda gösterme
    def display_image(self):
        if not self.image:
            return
            
        # şimdilik resmi olduğu gibi gösteriyoruz
        self.tk_image = ImageTk.PhotoImage(self.image)
        
        self.canvas.config(scrollregion=(0, 0, self.image.width, self.image.height))
        self.canvas.create_image(0, 0, image=self.tk_image, anchor="nw")

    # fare tıklandığında (seçim başlangıcı)
    def on_mouse_down(self, event):
        if not self.image:
            return
        # kaydırmayı hesaba katıyoruz
        self.rect_start_x = self.canvas.canvasx(event.x)
        self.rect_start_y = self.canvas.canvasy(event.y)
        self.current_rect = self.canvas.create_rectangle(
            self.rect_start_x, self.rect_start_y, self.rect_start_x, self.rect_start_y,
            outline="red", width=2
        )

    # fare sürüklendiğinde (seçim devamı)
    def on_mouse_drag(self, event):
        if not self.current_rect:
            return
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.current_rect, self.rect_start_x, self.rect_start_y, cur_x, cur_y)

    # fare bırakıldığında (seçim bitişi)
    def on_mouse_up(self, event):
        if not self.current_rect:
            return
            
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        
        # koordinatları normalize ediyoruz (x1 < x2, y1 < y2)
        x1, x2 = sorted([self.rect_start_x, cur_x])
        y1, y2 = sorted([self.rect_start_y, cur_y])
        
        w = x2 - x1
        h = y2 - y1
        
        if w < 5 or h < 5: # çok küçükse iptal et
            self.canvas.delete(self.current_rect)
            self.current_rect = None
            return
            
        # kullanıcıdan alan adı ve özelliklerini istiyoruz
        dialog = FieldDefinitionDialog(self)
        self.wait_window(dialog)
        
        if dialog.result:
            name = dialog.result['name']
            is_rotated = dialog.result['is_rotated']
            
            # alanı kaydediyoruz
            field = {
                'name': name,
                'rect_id': self.current_rect,
                'x': int(x1),
                'y': int(y1),
                'w': int(w),
                'h': int(h),
                'is_rotated': is_rotated
            }
            self.fields.append(field)
            
            # kalıcı dikdörtgen ve etiket çiziyoruz
            # döndürülmüş ise farklı renk veya işaret koyabiliriz
            color = "orange" if is_rotated else "red"
            self.canvas.itemconfig(self.current_rect, outline=color)
            
            label_text = name + (" (Döndürülmüş)" if is_rotated else "")
            self.canvas.create_text(x1, y1-10, text=label_text, fill=color, anchor="sw")
            self.refresh_fields_list()
        else:
            self.canvas.delete(self.current_rect)
            
        self.current_rect = None

    # ocr testi yapma
    def test_ocr(self):
        if not self.image:
            messagebox.showwarning("Uyarı", "Lütfen önce bir PDF yükleyin.")
            return
            
        if not self.fields:
            messagebox.showwarning("Uyarı", "Lütfen test etmek için en az bir alan seçin.")
            return
            
        # son eklenen alanı test ediyoruz
        field = self.fields[-1]
        
        try:
            cropped = self.image.crop((field['x'], field['y'], field['x']+field['w'], field['y']+field['h']))
            
            # eğer döndürülmüş ise resmi çeviriyoruz
            if field.get('is_rotated', False):
                # metin 90 derece dik ise, okumak için -90 (veya 270) çevirmemiz gerekebilir
                # genelde aşağıdan yukarıya yazılmışsa 90, yukarıdan aşağıya ise -90
                # varsayılan olarak 90 derece (sağa yatık) kabul edip düzeltmek için sola çevirelim
                # kullanıcı deneyimine göre bu değişebilir, şimdilik 90 derece sola çeviriyoruz (expand=True önemli)
                cropped = cropped.rotate(90, expand=True)
                
            text = pytesseract.image_to_string(cropped, lang='tur+eng', config='--psm 7')
            messagebox.showinfo("OCR Sonucu", f"Alan: {field['name']}\nOkunan Değer: '{text.strip()}'")
        except Exception as e:
            messagebox.showerror("Hata", f"OCR Hatası: {e}")

    # şablonu kaydetme
    def save_template(self):
        name = self.template_name_entry.get().strip()
        if not name:
            messagebox.showwarning("Uyarı", "Lütfen şablon adı girin.")
            return
            
        if not self.fields:
            messagebox.showwarning("Uyarı", "Lütfen en az bir alan tanımlayın.")
            return
            
        # sayfa boyutlarını ekliyoruz (ölçekleme için)
        page_w = self.image.width
        page_h = self.image.height
        
        fields_to_save = []
        for f in self.fields:
            fields_to_save.append({
                'name': f['name'],
                'x': f['x'],
                'y': f['y'],
                'w': f['w'],
                'h': f['h'],
                'is_rotated': f.get('is_rotated', False),
                'page_width': page_w,
                'page_height': page_h
            })
            
        TemplateManager.save_template(name, fields_to_save)
        messagebox.showinfo("Başarılı", f"'{name}' şablonu kaydedildi.")

    # alan listesini yenileme
    def refresh_fields_list(self):
        for widget in self.fields_frame.winfo_children():
            widget.destroy()
            
        for i, field in enumerate(self.fields):
            f_frame = ctk.CTkFrame(self.fields_frame)
            f_frame.pack(fill="x", pady=2)
            
            ctk.CTkLabel(f_frame, text=field['name']).pack(side="left", padx=5)
            
            del_btn = ctk.CTkButton(f_frame, text="Sil", width=40, fg_color="red", 
                                  command=lambda idx=i: self.delete_field(idx))
            del_btn.pack(side="right", padx=5)

    # alan silme
    def delete_field(self, index):
        field = self.fields.pop(index)
        self.canvas.delete(field['rect_id'])
        # tüm çizimleri temizleyip tekrar çiziyoruz
        self.canvas.delete("all")
        self.display_image()
        for f in self.fields:
            color = "orange" if f.get('is_rotated', False) else "red"
            rect = self.canvas.create_rectangle(f['x'], f['y'], f['x']+f['w'], f['y']+f['h'], outline=color, width=2)
            label_text = f['name'] + (" (Döndürülmüş)" if f.get('is_rotated', False) else "")
            self.canvas.create_text(f['x'], f['y']-10, text=label_text, fill=color, anchor="sw")
            f['rect_id'] = rect
            
        self.refresh_fields_list()

# alan tanımlama penceresi
class FieldDefinitionDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Alan Tanımla")
        self.geometry("300x200")
        self.resizable(False, False)
        
        self.result = None
        
        # pencereyi merkeze alıyoruz
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        
    def create_widgets(self):
        ctk.CTkLabel(self, text="Alan Adı:").pack(pady=(20, 5))
        self.name_entry = ctk.CTkEntry(self)
        self.name_entry.pack(pady=5, padx=20, fill="x")
        self.name_entry.focus()
        
        self.rotated_var = ctk.BooleanVar(value=False)
        self.rotated_check = ctk.CTkCheckBox(self, text="Döndürülmüş Metin (90°)", variable=self.rotated_var)
        self.rotated_check.pack(pady=10)
        
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20, fill="x", padx=20)
        
        ctk.CTkButton(btn_frame, text="İptal", fg_color="gray", command=self.destroy, width=100).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Kaydet", command=self.save, width=100).pack(side="right", padx=5)
        
        self.bind("<Return>", lambda e: self.save())
        self.bind("<Escape>", lambda e: self.destroy())

    def save(self):
        name = self.name_entry.get().strip()
        if name:
            self.result = {
                'name': name,
                'is_rotated': self.rotated_var.get()
            }
            self.destroy()
        else:
            self.name_entry.configure(placeholder_text="Lütfen isim girin")



