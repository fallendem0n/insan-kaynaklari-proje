# Ofis Asistanı Pro

Ofis Asistanı Pro, PDF dosyalarıyla çalışmayı kolaylaştıran kapsamlı bir araçtır. PDF'lerden veri çıkarma, toplu yeniden adlandırma, görsel şablon oluşturma ve metne dönüştürme gibi özellikler sunar.

## Özellikler

- **PDF Veri Çıkarıcı:** Regex veya görsel şablonlar kullanarak PDF'lerden veri ayıklayın ve Excel'e kaydedin.
- **PDF Yeniden Adlandırıcı:** PDF içeriklerine göre dosyaları otomatik olarak yeniden adlandırın.
- **Görsel Şablon Oluşturucu:** PDF üzerinde alanları görsel olarak seçerek veri çıkarma şablonları oluşturun.
- **PDF'den Metne:** PDF dosyalarını metin (.txt) formatına dönüştürün (OCR destekli).
- **OCR Desteği:** Tesseract ve PyMuPDF kullanarak taranmış dokümanlardan metin okuma.

## Kurulum

1. Python 3.10 veya üzeri bir sürümün yüklü olduğundan emin olun.
2. Gerekli kütüphaneleri yükleyin:

```bash
pip install -r requirements.txt
```

*Not: Eğer `requirements.txt` dosyanız yoksa, aşağıdaki komutla temel paketleri yükleyebilirsiniz:*

```bash
pip install customtkinter Pillow pandas openpyxl pytesseract pymupdf
```

3. **Tesseract OCR:** Proje klasörü içinde `tesseract` adında bir klasör bulunmalı ve içinde `tesseract.exe` olmalıdır. Veya sisteminizde Tesseract yüklü olmalıdır.

## Kullanım

Uygulamayı başlatmak için:

```bash
python main.py
```

## EXE Oluşturma (Build)

Uygulamayı tek bir klasör halinde (taşınabilir) paketlemek için:

```bash
python build_exe.py
```

Bu işlem `dist/OfisAsistaniPro` klasöründe çalıştırılabilir dosyaları oluşturacaktır.

## Lisans

Bu proje özel kullanım içindir.
