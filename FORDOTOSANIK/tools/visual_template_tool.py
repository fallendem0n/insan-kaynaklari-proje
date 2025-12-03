import customtkinter as ctk
from tkinter import filedialog, messagebox, Canvas
import os
import sys
from PIL import Image, ImageTk
from utils.template_manager import TemplateManager

try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None

class VisualTemplateFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        self.pdf_path = None
        self.image = None
        self.tk_image = None
        self.rect_start_x = None
        self.rect_start_y = None
        self.current_rect = None
        self.fields = [] # List of {'name', 'rect_id', 'x', 'y', 'w', 'h'}
        self.scale_factor = 1.0
        
        # Paths
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.poppler_path = os.path.join(application_path, 'poppler', 'Library', 'bin')
        self.tesseract_path = os.path.join(application_path, 'tesseract', 'tesseract.exe')
        
        if pytesseract:
            try:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            except Exception:
                pass

        self.create_widgets()

    def create_widgets(self):
        # Layout: Left sidebar (controls), Right (Canvas)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # --- Sidebar ---
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

        # --- Canvas Area ---
        canvas_frame = ctk.CTkFrame(self)
        canvas_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        # Scrollbars for canvas
        v_scroll = ctk.CTkScrollbar(canvas_frame, orientation="vertical")
        h_scroll = ctk.CTkScrollbar(canvas_frame, orientation="horizontal")
        
        self.canvas = Canvas(canvas_frame, bg="#333333", highlightthickness=0, 
                             yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        v_scroll.configure(command=self.canvas.yview)
        h_scroll.configure(command=self.canvas.xview)
        
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Bind events
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Dosyaları", "*.pdf")])
        if not file_path:
            return
            
        try:
            if not convert_from_path:
                messagebox.showerror("Hata", "pdf2image kütüphanesi eksik.")
                return

            images = convert_from_path(file_path, poppler_path=self.poppler_path, first_page=1, last_page=1)
            if images:
                self.pdf_path = file_path
                self.image = images[0]
                self.display_image()
                self.fields = []
                self.refresh_fields_list()
                self.canvas.delete("all")
                self.display_image() # Redraw image
        except Exception as e:
            messagebox.showerror("Hata", f"PDF yüklenirken hata: {e}")

    def display_image(self):
        if not self.image:
            return
            
        # Resize if too large to fit comfortably? No, let's keep original resolution for accuracy, but maybe scrollable.
        # Actually, displaying full res might be too big. Let's scale down for display but keep coords relative.
        
        # For now, let's just display as is.
        self.tk_image = ImageTk.PhotoImage(self.image)
        
        self.canvas.config(scrollregion=(0, 0, self.image.width, self.image.height))
        self.canvas.create_image(0, 0, image=self.tk_image, anchor="nw")

    def on_mouse_down(self, event):
        if not self.image:
            return
        # Account for scrolling
        self.rect_start_x = self.canvas.canvasx(event.x)
        self.rect_start_y = self.canvas.canvasy(event.y)
        self.current_rect = self.canvas.create_rectangle(
            self.rect_start_x, self.rect_start_y, self.rect_start_x, self.rect_start_y,
            outline="red", width=2
        )

    def on_mouse_drag(self, event):
        if not self.current_rect:
            return
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.current_rect, self.rect_start_x, self.rect_start_y, cur_x, cur_y)

    def on_mouse_up(self, event):
        if not self.current_rect:
            return
            
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        
        # Normalize coords (x1 < x2, y1 < y2)
        x1, x2 = sorted([self.rect_start_x, cur_x])
        y1, y2 = sorted([self.rect_start_y, cur_y])
        
        w = x2 - x1
        h = y2 - y1
        
        if w < 5 or h < 5: # Too small
            self.canvas.delete(self.current_rect)
            self.current_rect = None
            return
            
        # Ask for name
        dialog = ctk.CTkInputDialog(text="Alan Adı:", title="Alan Tanımla")
        name = dialog.get_input()
        
        if name:
            # Save field
            field = {
                'name': name,
                'rect_id': self.current_rect,
                'x': int(x1),
                'y': int(y1),
                'w': int(w),
                'h': int(h)
            }
            self.fields.append(field)
            
            # Draw permanent rect with label
            self.canvas.create_text(x1, y1-10, text=name, fill="red", anchor="sw")
            self.refresh_fields_list()
        else:
            self.canvas.delete(self.current_rect)
            
        self.current_rect = None

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

    def delete_field(self, index):
        field = self.fields.pop(index)
        self.canvas.delete(field['rect_id'])
        # Also delete text label? We didn't save its ID. 
        # For simplicity, let's just redraw everything or ignore the text label (it will stay until refresh).
        # Better: redraw all fields.
        self.canvas.delete("all")
        self.display_image()
        for f in self.fields:
            rect = self.canvas.create_rectangle(f['x'], f['y'], f['x']+f['w'], f['y']+f['h'], outline="red", width=2)
            self.canvas.create_text(f['x'], f['y']-10, text=f['name'], fill="red", anchor="sw")
            f['rect_id'] = rect
            
        self.refresh_fields_list()

    def test_ocr(self):
        if not self.image:
            messagebox.showwarning("Uyarı", "Lütfen önce bir PDF yükleyin.")
            return
            
        if not self.fields:
            messagebox.showwarning("Uyarı", "Lütfen test etmek için en az bir alan seçin.")
            return
            
        # Test the last added field or selected field? 
        # For simplicity, let's test the last added field since we don't have selection logic for list items yet.
        field = self.fields[-1]
        
        try:
            cropped = self.image.crop((field['x'], field['y'], field['x']+field['w'], field['y']+field['h']))
            text = pytesseract.image_to_string(cropped, lang='tur+eng', config='--psm 7')
            messagebox.showinfo("OCR Sonucu", f"Alan: {field['name']}\nOkunan Değer: '{text.strip()}'")
        except Exception as e:
            messagebox.showerror("Hata", f"OCR Hatası: {e}")

    def save_template(self):
        name = self.template_name_entry.get().strip()
        if not name:
            messagebox.showwarning("Uyarı", "Lütfen şablon adı girin.")
            return
            
        if not self.fields:
            messagebox.showwarning("Uyarı", "Lütfen en az bir alan tanımlayın.")
            return
            
        # Add page dimensions to fields for scaling later
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
                'page_width': page_w,
                'page_height': page_h
            })
            
        TemplateManager.save_template(name, fields_to_save)
        messagebox.showinfo("Başarılı", f"'{name}' şablonu kaydedildi.")
