import os
import shutil
import datetime

# yedekleme işlemlerini yöneten sınıf
class BackupManager:
    @staticmethod
    def create_backup(file_paths, backup_folder_name="Yedek"):
        """
        seçilen dosyaların yedeğini oluşturur.
        
        args:
            file_paths (list): yedeklenecek dosyaların tam yolları.
            backup_folder_name (str): yedek klasörünün adı.
            
        returns:
            list: yedeklenen dosyaların yeni yolları.
        """
        if not file_paths:
            return []
            
        # yedek klasörünü, ilk dosyanın olduğu yere açıyoruz
        base_dir = os.path.dirname(file_paths[0])
        backup_dir = os.path.join(base_dir, backup_folder_name)
        
        # klasör yoksa oluşturuyoruz
        if not os.path.exists(backup_dir):
            try:
                os.makedirs(backup_dir)
            except OSError as e:
                print(f"yedek klasörü oluşturulamadı: {e}")
                return []
                
        backed_up_files = []
        
        # dosya isimleri çakışırsa kullanmak için zaman damgası
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        for file_path in file_paths:
            try:
                if not os.path.exists(file_path):
                    continue
                    
                filename = os.path.basename(file_path)
                
                # hedef dosya yolunu belirliyoruz
                backup_path = os.path.join(backup_dir, filename)
                
                # eğer aynı isimde dosya varsa, sonuna tarih ekliyoruz
                # böylece eski yedeğin üzerine yazmamış oluruz
                if os.path.exists(backup_path):
                    base, ext = os.path.splitext(filename)
                    backup_path = os.path.join(backup_dir, f"{base}_{timestamp}{ext}")
                    
                # dosyayı kopyalıyoruz (meta verileriyle birlikte)
                shutil.copy2(file_path, backup_path)
                backed_up_files.append(backup_path)
                
            except Exception as e:
                print(f"dosya yedeklenirken hata ({file_path}): {e}")
                
        return backed_up_files
