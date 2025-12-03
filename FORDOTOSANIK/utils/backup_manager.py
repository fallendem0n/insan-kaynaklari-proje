import os
import shutil
import datetime

class BackupManager:
    @staticmethod
    def create_backup(file_paths, backup_folder_name="Yedek"):
        """
        Seçilen dosyaların yedeğini oluşturur.
        
        Args:
            file_paths (list): Yedeklenecek dosyaların tam yolları.
            backup_folder_name (str): Yedek klasörünün adı.
            
        Returns:
            list: Yedeklenen dosyaların yeni yolları.
        """
        if not file_paths:
            return []
            
        # İlk dosyanın bulunduğu dizini baz alarak yedek klasörü oluştur
        base_dir = os.path.dirname(file_paths[0])
        backup_dir = os.path.join(base_dir, backup_folder_name)
        
        if not os.path.exists(backup_dir):
            try:
                os.makedirs(backup_dir)
            except OSError as e:
                print(f"Yedek klasörü oluşturulamadı: {e}")
                return []
                
        backed_up_files = []
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        for file_path in file_paths:
            try:
                if not os.path.exists(file_path):
                    continue
                    
                filename = os.path.basename(file_path)
                # Dosya adının çakışmaması için timestamp ekleyebiliriz veya direkt kopyalayabiliriz.
                # Kullanıcı "orjinal halleri" dediği için direkt kopyalamak daha mantıklı, 
                # ama aynı isimde varsa üzerine yazmamak için kontrol edelim.
                
                backup_path = os.path.join(backup_dir, filename)
                
                if os.path.exists(backup_path):
                    base, ext = os.path.splitext(filename)
                    backup_path = os.path.join(backup_dir, f"{base}_{timestamp}{ext}")
                    
                shutil.copy2(file_path, backup_path)
                backed_up_files.append(backup_path)
                
            except Exception as e:
                print(f"Dosya yedeklenirken hata ({file_path}): {e}")
                
        return backed_up_files
