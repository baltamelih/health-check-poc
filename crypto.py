import base64

# 1️⃣ TXT dosyasını Base64 olarak kodla
with open("HCv54.txt", "rb") as file:
    encoded = base64.b64encode(file.read()).decode("utf-8")

# 2️⃣ Şifrelenmiş içeriği yeni bir Python dosyasına yaz
with open("encoded_script.py", "w") as file:
    file.write(f'SECURITY_SCRIPT_BASE64 = """{encoded}"""')



print("Dosya başarıyla Base64 formatına çevrildi ve encoded_script.py dosyasına kaydedildi.")


