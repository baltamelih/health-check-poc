# prepare_config.py

import os
from pathlib import Path
from dotenv import load_dotenv

print("Building config file from .env...")

# Proje ana dizinindeki .env dosyasını yükle
dotenv_path = Path('.') / '.env'
if not dotenv_path.exists():
    raise FileNotFoundError("'.env' file not found! Cannot create config.py.")

load_dotenv(dotenv_path=dotenv_path)

# SECRET_KEY'i .env'den oku
secret_key = os.getenv("SECRET_KEY")

if not secret_key:
    raise ValueError("SECRET_KEY not found in .env file or it is empty.")

# config.py dosyasının içeriğini hazırla
# Anahtarı doğrudan bir bytes literali (b'...') olarak yazmak en sağlıklısıdır.
# Bu, encoding problemlerinin önüne geçer.
config_content = f"""# -*- coding: utf-8 -*-
#
# BU DOSYA OTOMATİK OLARAK prepare_config.py TARAFINDAN OLUŞTURULMUŞTUR.
# DOĞRUDAN DÜZENLEMEYİN!
#

# Fernet tarafından kullanılacak şifreleme anahtarı
SECRET_KEY = {secret_key.encode('utf-8')}
"""

# config.py dosyasını yaz
config_file_path = Path('config.py')
config_file_path.write_text(config_content, encoding='utf-8')

print(f"✅ Successfully created/updated '{config_file_path}' with the key from .env.")