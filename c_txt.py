# encode_and_encrypt_all.py
import os
import base64
from glob import glob
from pathlib import Path
from cryptography.fernet import Fernet

# İsteğe bağlı: .env kullanmak istersen yorumdan çıkar
# from dotenv import load_dotenv
# load_dotenv()

ENV_KEY_NAME = "SECRET_KEY"
INPUT_DIR = Path("txt")
OUTPUT_DIR = Path("src")
OUTPUT_EXT = ".enc"  # istersen ".txt" yapabilirsin

def get_or_create_key():
    # 1) .env veya environment'tan oku
    key = os.getenv(ENV_KEY_NAME)
    if key:
        return key.encode("utf-8")

    # 2) Yoksa oluştur ve .env'ye yaz (güvenli saklama için önerilir)
    key = Fernet.generate_key()
    # .env üret (varsa append etme mantığını basit tuttuk)
    env_path = Path(".env")
    if env_path.exists():
        # Varsa, ENV_KEY_NAME zaten tanımlıysa dokunma
        content = env_path.read_text(encoding="utf-8")
        if f"{ENV_KEY_NAME}=" not in content:
            with env_path.open("a", encoding="utf-8") as f:
                f.write(f"\n{ENV_KEY_NAME}={key.decode('utf-8')}\n")
    else:
        env_path.write_text(f"{ENV_KEY_NAME}={key.decode('utf-8')}\n", encoding="utf-8")

    return key

def encode_and_encrypt_all():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    key = get_or_create_key()
    cipher = Fernet(key)

    txt_files = sorted(glob(str(INPUT_DIR / "*.txt")))
    if not txt_files:
        print("⚠️ txt/ içinde .txt dosyası bulunamadı.")
        return

    for in_path in txt_files:
        in_path = Path(in_path)
        # 1) ham içeriği oku
        raw = in_path.read_bytes()

        # 2) Base64'e çevir (not: doğrudan raw'ı şifrelemek de mümkündü)
        b64_str = base64.b64encode(raw).decode("utf-8")

        # 3) Fernet ile şifrele
        encrypted = cipher.encrypt(b64_str.encode("utf-8"))

        # 4) çıktı adı: aynı dosya adı + .enc
        out_name = in_path.stem + OUTPUT_EXT
        out_path = OUTPUT_DIR / out_name
        out_path.write_bytes(encrypted)

        print(f"✅ {in_path.name} → {out_name} (tek SECRET_KEY ile şifrelendi)")

    print("\n🎉 Tüm dosyalar şifrelendi. Key kaynağı: '.env' içindeki SECRET_KEY")

def decrypt_one(encrypted_path: Path, output_path: Path):
    key = os.getenv(ENV_KEY_NAME)
    if not key:
        raise RuntimeError("SECRET_KEY environment veya .env içinde bulunamadı.")
    cipher = Fernet(key.encode("utf-8"))

    enc = encrypted_path.read_bytes()
    b64_str = cipher.decrypt(enc).decode("utf-8")
    raw = base64.b64decode(b64_str)
    output_path.write_bytes(raw)
    print(f"🔓 Çözüldü: {encrypted_path.name} → {output_path.name}")

if __name__ == "__main__":
    encode_and_encrypt_all()
    # Örnek çözme:
    # decrypt_one(Path("src/schema_cleanup.enc"), Path("decoded/schema_cleanup.txt"))
