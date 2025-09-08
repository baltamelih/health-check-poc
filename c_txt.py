import base64
import os

# 1️⃣ TXT dosyasını Base64 olarak kodla
with open("txt/xpcmdshell.txt", "rb") as file:
    encoded = base64.b64encode(file.read()).decode("utf-8")

# Ensure src/ directory exists
os.makedirs("src", exist_ok=True)

# 2️⃣ Şifrelenmiş içeriği yeni bir Python dosyasına yaz
with open("src/sql_execute_extended.txt", "w") as file:
    file.write(f'SECURITY_SCRIPT11_BASE64 = """{encoded}"""')


print("Dosya başarıyla Base64 formatına çevrildi ve src/schema_cleanup.txt dosyasına kaydedildi.")


