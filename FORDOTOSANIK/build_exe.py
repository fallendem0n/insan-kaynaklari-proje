import PyInstaller.__main__
import os
import shutil

import stat
import time

def remove_readonly(func, path, excinfo):
    # Read-only dosyaların silinmesini sağlar
    os.chmod(path, stat.S_IWRITE)
    func(path)

def force_remove_dir(dir_path):
    if os.path.exists(dir_path):
        try:
            # Python 3.12+ için onexc, eskiler için onerror (gerçi kullanıcı 3.14 kullanıyor görünüyor)
            shutil.rmtree(dir_path, onexc=remove_readonly)
        except TypeError:
            # Eski sürüm uyumluluğu
            shutil.rmtree(dir_path, onerror=remove_readonly)
        except Exception as e:
            print(f"Klasör silinirken hata: {e}")
            print("Lütfen 'dist' veya 'build' klasörünü kullanan programları kapatın.")
            time.sleep(1)
            try:
                shutil.rmtree(dir_path, onexc=remove_readonly)
            except:
                pass

def build_exe():
    # temizlik yapalım
    force_remove_dir('dist')
    force_remove_dir('build')

    # pyinstaller komutunu hazırlıyoruz
    args = [
        'main.py',  # ana dosya
        '--name=OfisAsistaniPro',  # exe adı
        '--windowed',  # konsol penceresi açılmasın
        '--onedir',  # klasör olarak çıkar (daha hızlı açılır)
        '--icon=NONE', # ikon varsa buraya eklenir
        
        # veri dosyalarını ekliyoruz
        '--add-data=modern_theme.json;.',
        # 'poppler', 'tesseract' ve 'templates' klasörlerini exe içine gömmüyoruz, yanına kopyalayacağız
        
        # gerekli importlar
        '--hidden-import=PIL',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=customtkinter',
        '--hidden-import=pdf2image',
        '--hidden-import=pytesseract',
        '--hidden-import=fitz',
        '--hidden-import=pymupdf',
        
        # temiz bir build için
        '--clean',
        '--noconfirm',
    ]

    print("EXE oluşturuluyor, lütfen bekleyin...")
    PyInstaller.__main__.run(args)
    
    # klasörleri kopyalama işlemi
    print("Harici klasörler kopyalanıyor...")
    # onedir modunda dist/OfisAsistaniPro klasörü oluşur
    dist_folder = os.path.join('dist', 'OfisAsistaniPro')
    
    folders_to_copy = ['tesseract', 'templates']
    
    for folder in folders_to_copy:
        src = folder
        dst = os.path.join(dist_folder, folder)
        
        if os.path.exists(src):
            if os.path.exists(dst):
                force_remove_dir(dst)
            shutil.copytree(src, dst)
            print(f"'{folder}' kopyalandı.")
        else:
            print(f"UYARI: '{folder}' bulunamadı!")

    print("İşlem tamamlandı! 'dist' klasörünü kontrol edin.")

if __name__ == "__main__":
    build_exe()
