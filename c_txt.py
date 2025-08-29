import base64

# 1️⃣ TXT dosyasını Base64 olarak kodla
with open("txt/drop.txt", "rb") as file:
    encoded = base64.b64encode(file.read()).decode("utf-8")



# 2️⃣ Şifrelenmiş içeriği yeni bir Python dosyasına yaz
with open("schema_cleanup.py", "w") as file:
    file.write(f'SECURITY_SCRIPT8_BASE64 = """{encoded}"""')



print("Dosya başarıyla Base64 formatına çevrildi ve dp_script dosyasına kaydedildi.")


