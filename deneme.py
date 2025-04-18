from sec_script import SECURITY_SCRIPT3_BASE64  # İlk script
import traceback
import base64
import time
import sys
import os
import getpass
import re
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
from PyPDF2 import PdfReader, PdfWriter

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import numpy as np
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader
import chardet
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Drawing, String, Path, Circle
from reportlab.graphics import renderPDF
from reportlab.lib.colors import HexColor
from math import cos, sin, radians
from PIL import Image
from reportlab.lib.utils import ImageReader

import io
import hashlib
import json
import getpass
from cryptography.fernet import Fernet
from reportlab.graphics.renderPDF import drawToString
from PyQt5.QtWidgets import (

    QApplication, QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QCheckBox, QGridLayout
)
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QGridLayout, QLabel, QLineEdit, QComboBox,
    QCheckBox, QPushButton, QMessageBox, QDateEdit, QListWidget, QRadioButton, QStackedWidget
)
from PyQt5.QtCore import QDateTime
from PyQt5.QtCore import Qt, QDate
import pyodbc
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from charset_normalizer import detect


class SQLServerConnectionUI(QWidget):
    def __init__(self):
        super().__init__()
        self.current_user = self.get_windows_user()
        self.connection = None
        self.key = self.load_or_generate_key()  # sifreleme anahtarını yükle veya olustur
        self.fernet = Fernet(self.key)
        self.initUI()

    def load_or_generate_key(self):
        """sifreleme anahtarını olusturur veya mevcut olanı yükler."""
        try:
            key_file = "credentials_folder/secret.key"
            if not os.path.exists("credentials_folder"):
                os.makedirs("credentials_folder")
            if not os.path.exists(key_file):
                key = Fernet.generate_key()
                with open(key_file, "wb") as file:
                    file.write(key)
            else:
                with open(key_file, "rb") as file:
                    key = file.read()
            return key
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Anahtar dosyasini yuklerken bir hata olustu:\n{e}")
            return None

    def encrypt_data(self, data):
        """Veriyi sifreler."""
        return self.fernet.encrypt(data.encode()).decode()

    def decrypt_data(self, data):
        """sifreli veriyi çözer."""
        return self.fernet.decrypt(data.encode()).decode()

    def initUI(self):
        self.setWindowTitle("SQL Server baglantisi")
        self.resize(450, 300)

        # Layout
        grid = QGridLayout()
        self.setLayout(grid)

        # Server Type
        grid.addWidget(QLabel("Server type:"), 0, 0)
        self.server_type = QComboBox()
        self.server_type.addItems(["Database Engine"])
        grid.addWidget(self.server_type, 0, 1)

        # Server Name
        grid.addWidget(QLabel("Server name:"), 1, 0)
        self.server_name = QLineEdit()
        grid.addWidget(self.server_name, 1, 1)

        # Authentication
        grid.addWidget(QLabel("Authentication:"), 2, 0)
        self.auth_type = QComboBox()
        self.auth_type.addItems(["SQL Server Authentication", "Windows Authentication"])
        self.auth_type.currentIndexChanged.connect(self.toggle_authentication)
        grid.addWidget(self.auth_type, 2, 1)

        # Login
        grid.addWidget(QLabel("Login:"), 3, 0)
        self.login = QLineEdit()
        grid.addWidget(self.login, 3, 1)

        # Password
        grid.addWidget(QLabel("Password:"), 4, 0)
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        grid.addWidget(self.password, 4, 1)

        # Remember Password
        self.remember_password = QCheckBox("Remember password")
        self.remember_password.stateChanged.connect(self.save_credentials)
        grid.addWidget(self.remember_password, 5, 1)

        # Connect Button
        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.test_connection)
        grid.addWidget(self.connect_button, 6, 0, 1, 2)

        # Export Button
        self.export_button = QPushButton("Run Scripts and Export")
        self.export_button.clicked.connect(self.run_scripts_and_export)
        self.export_button.setDisabled(True)
        grid.addWidget(self.export_button, 7, 0, 1, 2)
        # Report Comparison Button


        # Load Saved Credentials
        self.load_saved_credentials()
        self.toggle_authentication()
        self.show()

    def open_report_comparison(self):
        try:
            server_name = self.server_name.text().strip()  # Giristeki server_name bilgisini al
            if not server_name:
                QMessageBox.warning(self, "Error", "Server name cannot be empty.")
                return


        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

    def get_windows_user(self):
        """Windows kullanıcısını alma."""
        username = getpass.getuser()
        domain = os.environ.get('USERDOMAIN', '')
        if domain:
            return f"{domain}\{username}"
        return username

    def toggle_authentication(self):
        """Authentication türüne göre Login ve Password kutularını kontrol eder."""
        if self.auth_type.currentText() == "Windows Authentication":
            self.login.setText(self.get_windows_user())
            self.login.setDisabled(True)
            self.password.setDisabled(True)
        else:
            self.login.setText("")
            self.login.setDisabled(False)
            self.password.setDisabled(False)

    def save_credentials(self):
        """Kullanıcı bilgilerini sifreleyerek kaydeder."""
        try:
            if self.remember_password.isChecked():
                server_name = self.server_name.text()
                if not server_name:
                    QMessageBox.warning(self, "Error", "Server name cannot be empty.")
                    self.remember_password.setChecked(False)
                    return

                data = {"server_name": server_name}

                if self.auth_type.currentText() == "SQL Server Authentication":
                    login = self.login.text()
                    password = self.password.text()
                    if not login or not password:
                        QMessageBox.warning(self, "Error", "Login and Password cannot be empty.")
                        self.remember_password.setChecked(False)
                        return
                    data.update({
                        "login": login,
                        "password": password
                    })

                self.save_to_file(data)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Bilgileri kaydederken bir hata olustu:\n{e}")
            self.remember_password.setChecked(False)

    def load_saved_credentials(self):
        """Kaydedilmis bilgileri yükler ve çözer."""
        try:
            if os.path.exists("credentials_folder/credentials.json"):
                with open("credentials_folder/credentials.json", "r") as file:
                    data = json.load(file)
                    self.server_name.setText(data.get("server_name", ""))
                    if "login" in data and "password" in data:
                        self.login.setText(data.get("login", ""))
                        self.password.setText(data.get("password", ""))
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaydedilmis bilgileri yuklerken bir hata olustu:\n{e}")

    def save_to_file(self, data):
        """Bilgileri düz metin olarak dosyaya kaydeder."""
        os.makedirs("credentials_folder", exist_ok=True)
        with open("credentials_folder/credentials.json", "w") as file:
            json.dump(data, file)

    def test_connection(self):
        """SQL Server baglantisini test eder."""
        server = self.server_name.text()
        auth = self.auth_type.currentText()

        try:
            if auth == "Windows Authentication":
                # Windows Authentication ile baglanti
                self.conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server},1433;Trusted_Connection=yes;autocommit=True;"
            else:
                # SQL Server Authentication ile baglanti
                login = self.login.text()
                password = self.password.text()
                self.conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};UID={login};PWD={password};autocommit=True;"

            self.connection = pyodbc.connect(self.conn_str,autocommit=True)
            QMessageBox.information(self, "BASARILI", "SQL Server baglantisi BASARILI!")
            self.export_button.setDisabled(False)


        except Exception as e:
            QMessageBox.critical(self, "Hata", f"baglanti Basarisiz:\n{e}")


    def summarize_dataframe(self,grouped_df):

        summary = {}
        for group in grouped_df["Group"].unique():
            group_df = grouped_df[grouped_df["Group"] == group]
            success_count = (group_df["Status"] == "Successful").sum()
            failed_count = (group_df["Status"] == "Failed").sum()
            warning_count = (group_df["Status"]== "Warning").sum()
            summary[group] = {"Successful": success_count, "Warning": warning_count, "Failed": failed_count}
        return summary

    def get_resource_path(self,relative_path):
        """Kaynak dosyaların dogru yolunu döndürür."""
        if getattr(sys, 'frozen', False):  # Eger exe olarak çalısıyorsa
            base_path = sys._MEIPASS  # PyInstaller'ın geçici dosyaları koydugu klasör
        else:  # Python dosyası olarak çalısıyorsa
            base_path = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_path, relative_path)

    def group_dataframe(self,summary_df):
        """
        Group the DataFrame into categories for better organization in the PDF report.
        """
        group_mapping = {
            "Config": ["SecHC_Version", "SecHC_Xpcmdshell","SecHC_CLR","SPCumulativeUpdate"],
            "Logins": ["SecHC_DisableLogins", "SecHc_Sysadminlogin"],
            "Users": [
                "SecHC_BlankUsernamePassword",
                "SecHC_ComplexPassword",
                "SecHC_dbowner",
                "SecHC_OrphanUser",
                "SecHC_SaAccount",
                "SecHC_SameSQLUsernameasPassword",
                "SecHc_ServiceAccount",

            ],
        }

        grouped_data = []
        for group, items in group_mapping.items():
            for item in items:
                row = summary_df[summary_df["control_column_name"] == item]
                if not row.empty:
                    # Status kontrolü
                    if row["status"].iloc[0] == 1:
                        status = "Successful"
                    elif row["status"].iloc[0] == 2:
                        status = "Warning"
                    else:
                        status = "Failed"

                    # İsimleri temizle
                    item = re.sub(r'^SecHC_', '', item)
                    item = re.sub(r'^SecHc_', '', item)

                    # "Sysadminlogin" özel durum
                    if item == 'Sysadminlogin':
                        item = 'HighPriviligeLogin'

                    # Sonuçları listeye ekle
                    grouped_data.append([group, item, status])

        # Convert to DataFrame for easier handling
        grouped_df = pd.DataFrame(grouped_data, columns=["Group", "Description", "Status"])
        #print(grouped_df)
        return pd.DataFrame(grouped_df)


    def clean_text(self,text):
        if isinstance(text,str):
            return text.encode('utf-8', errors='replace').decode('utf-8')
        return text

    def clean_column(self,value):
        return value.replace(" ", "").lower()
    def clean_sheet_name(self,sheet_name):
        return sheet_name.replace("SecHC_", "").lower()


    def extract_table_names(self, sql_script):
        """
        SQL script'inden tüm SELECT ifadelerini ve ilgili tablo adlarını çıkarır.
        Örnek: SELECT * FROM #SecHC_ComplexPassword → SecHC_ComplexPassword
        """
        table_names = re.findall(r"SELECT\s+\*.*?FROM\s+#([a-zA-Z0-9_]+)", sql_script, re.IGNORECASE)

        return table_names

    def run_scripts_and_export(self):
        try:
            username = self.server_name.text()
            rundate = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"{username}_{rundate}.xlsx"

            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            csv_dir = os.path.join(base_path, "csv")
            os.makedirs(csv_dir, exist_ok=True)
            output_path = os.path.join(csv_dir, file_name)
            used_sheet_names = set()

            cursor = self.connection.cursor()
            encoded_scripts = [SECURITY_SCRIPT3_BASE64]

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                sheet_added = False  # 📌 En az bir sayfa yazıldı mı kontrolü
                for encoded_script in encoded_scripts:
                    decoded_bytes = base64.b64decode(encoded_script)
                    detected = chardet.detect(decoded_bytes)
                    encoding = detected.get('encoding', 'utf-8')
                    decoded_script = decoded_bytes.decode(encoding, errors='ignore')

                    # 'USE master; GO' ifadelerini bölme
                    parts = re.split(r'USE master;\s*GO', decoded_script, flags=re.IGNORECASE)

                    if len(parts) > 1:
                        script_to_execute = parts[1]  # İlk 'USE master; GO' sonrası kısmı al
                    else:
                        script_to_execute = decoded_script  # Eğer 'USE master; GO' yoksa, tüm scripti çalıştır

                    # Önce tüm betiği çalıştır
                    statements = script_to_execute.split("GO")
                    for statement in statements:
                        statement = statement.strip()
                        if statement:
                            try:
                                cursor.execute(statement)
                            except Exception as e:
                                print(f"SQL Execution Error: {e}")

                    # Şimdi sadece SELECT ifadelerini çalıştır ve verileri çek
                    select_statements = re.findall(r"(SELECT\s+\*.*?FROM\s+#\w+;)", script_to_execute, re.IGNORECASE)
                    table_names = self.extract_table_names(script_to_execute)

                    for i, select in enumerate(select_statements):
                        table_name = table_names[i] if i < len(table_names) else f"Sheet{i + 1}"

                        if table_name in used_sheet_names:
                            count = 1
                            while f"{table_name}_{count}" in used_sheet_names:
                                count += 1
                            table_name = f"{table_name}_{count}"
                        used_sheet_names.add(table_name)

                        try:
                            print(f"Executing SELECT: {select}")
                            cursor.execute(select)

                            if cursor.description:
                                columns = [desc[0] for desc in cursor.description]
                                rows = cursor.fetchall()
                                data = pd.DataFrame.from_records(rows, columns=columns)

                                if not data.empty:
                                    data.to_excel(writer, sheet_name=table_name[:31], index=False)
                                    sheet_added = True  # 📌 En az bir sayfa yazıldı
                        except Exception as e:
                            self.log_error(e, "windows_group_error")
                            continue

                # 📌 Eğer hiç veri yazılmadıysa, boş bir sayfa ekleyelim
                if not sheet_added:
                    df_placeholder = pd.DataFrame({"No Data": ["No queries returned results"]})
                    df_placeholder.to_excel(writer, sheet_name="Placeholder", index=False)

            QMessageBox.information(self, "SUCCESS", f"The report has been saved in the {output_path}")
        except Exception as e:
            self.log_error(e, "run_script")
            QMessageBox.critical(self, "Error", f"An occurred error:\n{e}")



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SQLServerConnectionUI()
    sys.exit(app.exec_())



