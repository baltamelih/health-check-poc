# prepare_assets.py
import os
import base64
from glob import glob
from pathlib import Path
from cryptography.fernet import Fernet
from dotenv import load_dotenv

# --- Ayarlar ve Sabitler ---
ENV_KEY_NAME = "SECRET_KEY"
INPUT_DIR = Path("txt")
OUTPUT_DIR = Path("src")
CONFIG_FILE = Path("config.py")
ENV_FILE = Path(".env")
OUTPUT_EXT = ".enc"

# --- Ana Fonksiyonlar ---

def get_or_create_key():
    """
    .env dosyasından anahtarı okur. Eğer dosya veya anahtar yoksa,
    yeni bir anahtar oluşturur, .env dosyasına yazar ve yeni anahtarı döndürür.
    Anahtarı her zaman bytes formatında döndürür.
    """
    # Mevcut .env dosyasını yüklemeyi dene
    load_dotenv(dotenv_path=ENV_FILE)
    
    # 1) Ortamdan veya .env'den anahtarı oku
    key_str = os.getenv(ENV_KEY_NAME)
    if key_str:
        print(f"🔑 Mevcut anahtar {ENV_FILE} dosyasından okundu.")
        return key_str.encode("utf-8")

    # 2) Anahtar bulunamadıysa, yeni bir tane oluştur
    print(f"⚠️ {ENV_KEY_NAME} bulunamadı, yeni bir anahtar oluşturuluyor...")
    new_key_bytes = Fernet.generate_key()
    
    # Yeni anahtarı .env dosyasına yaz
    # Dosyaya yazarken string'e çevirmemiz gerekir
    key_for_env = new_key_bytes.decode('utf-8')
    
    # Varsa sonuna ekle, yoksa yeni oluştur
    mode = "a" if ENV_FILE.exists() else "w"
    with ENV_FILE.open(mode, encoding="utf-8") as f:
        f.write(f"\n{ENV_KEY_NAME}={key_for_env}\n")
        
    print(f"✨ Yeni anahtar oluşturuldu ve {ENV_FILE} dosyasına kaydedildi.")
    return new_key_bytes

def update_config_file(key_bytes: bytes):
    """
    Verilen anahtarı kullanarak config.py dosyasını oluşturur veya günceller.
    """
    config_content = f"""# -*- coding: utf-8 -*-
#
# BU DOSYA OTOMATİK OLARAK prepare_assets.py TARAFINDAN OLUŞTURULMUŞTUR.
# DOĞRUDAN DÜZENLEMEYİN!
#

# Fernet tarafından kullanılacak şifreleme anahtarı
SECRET_KEY = {key_bytes}
"""
    CONFIG_FILE.write_text(config_content, encoding='utf-8')
    print(f"✅ '{CONFIG_FILE}' dosyası güncel anahtar ile oluşturuldu/güncellendi.")

def encode_and_encrypt_all(key_bytes: bytes):
    """
    Verilen anahtarı kullanarak txt/ klasöründeki tüm dosyaları şifreler.
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    cipher = Fernet(key_bytes)

    txt_files = sorted(glob(str(INPUT_DIR / "*.txt")))
    if not txt_files:
        print(f"🟡 '{INPUT_DIR}/' içinde şifrelenecek .txt dosyası bulunamadı.")
        return

    print(f"\n--- Dosyalar şifreleniyor ---")
    for in_path in txt_files:
        in_path = Path(in_path)
        raw = in_path.read_bytes()
        b64_str = base64.b64encode(raw).decode("utf-8")
        encrypted = cipher.encrypt(b64_str.encode("utf-8"))

        out_name = in_path.stem + OUTPUT_EXT
        out_path = OUTPUT_DIR / out_name
        out_path.write_bytes(encrypted)
        print(f"✅ {in_path.name} → {out_name}")

# --- Script'in Çalıştırılması ---

if __name__ == "__main__":
    print("--- Varlık Hazırlama Script'i Başlatıldı ---")
    
    # 1. Adım: Anahtarı al veya oluştur. Bu anahtar artık tek doğru kaynaktır.
    app_key = get_or_create_key()
    
    # 2. Adım: Aynı anahtarı kullanarak config.py dosyasını güncelle.
    update_config_file(app_key)
    
    # 3. Adım: Yine aynı anahtarı kullanarak tüm dosyaları şifrele.
    encode_and_encrypt_all(app_key)
    
    print("\n🎉 Tüm hazırlık işlemleri başarıyla tamamlandı!")