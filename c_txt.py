import base64
from cryptography.fernet import Fernet

# 1️⃣ Ortak secret-key oluştur (bir kere üret, sonra sakla)
secret_key = Fernet.generate_key()
cipher = Fernet(secret_key)

print("SECRET KEY:", secret_key.decode())
# Bu key'i bir defa üretip güvenli bir yerde sakla. 
# Aynı key ile hem şifreleme hem çözme yapacaksın.

# 2️⃣ TXT dosyasını oku
with open("txt/drop.txt", "rb") as file:
    raw_data = file.read()

# 3️⃣ Base64 encode et
encoded = base64.b64encode(raw_data).decode("utf-8")

# 4️⃣ Base64 edilmiş veriyi şifrele
encrypted = cipher.encrypt(encoded.encode("utf-8"))

# 5️⃣ Şifreli çıktıyı dosyaya yaz
with open("src/schema_cleanup.txt", "wb") as file:
    file.write(encrypted)

print("Dosya başarıyla Base64 + Fernet ile şifrelendi ve kaydedildi.")
