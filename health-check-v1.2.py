from encoded_script import SECURITY_SCRIPT_BASE64  # İlk script
from vlfcount import SECURITY_SCRIPT54_BASE64
from encoded_script2 import SECURITY_SCRIPT2_BASE64 #Hcv999
from hcscript import SECURITY_SCRIPT4_BASE64# İkinci script
from backuplogrunning import SECURITY_SCRIPT5_BASE64
from policynotcheckedlogin import SECURITY_SCRIPT6_BASE64
from samepasswordlogins import  SECURITY_SCRIPT7_BASE64
from full_script import SECURITY_SCRIPT9_BASE64
from drop import SECURITY_SCRIPT8_BASE64
from new_script import SECURITY_SCRIPT10_BASE64
from xpcmdshell import SECURITY_SCRIPT11_BASE64
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
    QProgressDialog,QFileDialog,QDialog,QSplashScreen,QApplication, QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QCheckBox, QGridLayout,QScrollArea
)
from PyQt5.QtWidgets import (
    QProgressBar,QTextEdit,QWidget, QVBoxLayout, QGridLayout, QLabel, QLineEdit, QComboBox,
    QCheckBox, QPushButton, QMessageBox, QDateEdit, QListWidget, QRadioButton, QStackedWidget,QTableWidget
)
from PyQt5.QtCore import QDateTime
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QDate,QTimer
import pyodbc
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from charset_normalizer import detect
from PyQt5.QtWidgets import (
    QWidget, QLabel, QLineEdit, QComboBox, QCheckBox, QPushButton,
    QVBoxLayout, QHBoxLayout, QFormLayout, QFrame
)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QListWidget,
    QProgressBar
)

from PyQt5.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QStackedWidget, QLabel,QHeaderView,QTableWidgetItem
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt

from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QStackedWidget, QLabel
)
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QSizePolicy

class MainUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setWindowTitle("DAPLAIT HEALTH CHECK")
        self.setFixedSize(1080, 720)

        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Sol Panel (%40)
        left_panel = self.build_logo_panel()
        left_panel.setFixedWidth(360)  # 40%
        main_layout.addWidget(left_panel)

        # Sağ Panel (%60) – Sağ tarafı stack + info frame içerecek bir dikey kutu yapalım
        self.right_container = QWidget()
        self.right_layout = QVBoxLayout(self.right_container)
        self.right_layout.setContentsMargins(0, 0, 0, 0)

        # Stack
        self.stack = QStackedWidget()
        self.stack.setMinimumWidth(720)  # 60%
        self.stack.setMaximumWidth(720)
        self.stack.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.stack.setStyleSheet("background-color: #F6F6F6;")

        # Info panel (başlangıçta görünmesin)
        self.server_info_frame = QFrame()

        self.server_info_frame.setVisible(False)
        self.server_info_frame.setStyleSheet("""
            QFrame {
                background-color: #f9f9f9;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 16px;
            }
        """)

        # Sağ panele ekle (önce info, sonra stack)
        self.right_layout.addWidget(self.server_info_frame)
        self.right_layout.addWidget(self.stack)
        main_layout.addWidget(self.right_container)

        # Sayfalar
        self.progress_overlay = self.create_progress_overlay()
        self.checklist_page = CheckListWindow()
        self.consent_page = ConsentWindow(self.switch_to_connection)
        self.connection_page = SQLServerConnectionUI(self.switch_to_checklist, self)

        self.progress_overlay.hide()
        self.stack.addWidget(self.consent_page)
        self.stack.addWidget(self.connection_page)
        self.stack.addWidget(self.checklist_page)
        self.stack.addWidget(self.progress_overlay)
        self.stack.setCurrentIndex(0)

    def display_server_info(self, rows, columns):
        try:
            layout = QVBoxLayout(self.server_info_frame)
            layout.setContentsMargins(20, 20, 20, 20)
            layout.setSpacing(15)

            info_box = QFrame()
            info_box.setStyleSheet("""
                QFrame {
                    background-color: #ffffff;
                    border: 1px solid #e0e0e0;
                    border-radius: 12px;
                    padding: 20px;
                }
            """)
            info_box.setMinimumHeight(240)

            info_layout = QGridLayout(info_box)
            info_layout.setHorizontalSpacing(25)
            info_layout.setVerticalSpacing(15)

            display_keys = ['Sunucu', 'Edition', 'Version', 'CPU', 'RAM (MB)', 'SQL Version']
            font = QFont("Segoe UI", 10)

            for idx, key in enumerate(display_keys):
                if key in columns:
                    col_index = columns.index(key)
                    val = rows[0][col_index]

                    box = QFrame()
                    box.setMinimumWidth(300)  # 🧩 Genişlik artırıldı
                    box.setStyleSheet("""
                        QFrame {
                            background-color: #f9f9f9;
                            border: 1px solid #d0d0d0;
                            border-radius: 10px;
                            padding: 8px 12px;
                        }
                    """)

                    hbox = QHBoxLayout(box)
                    hbox.setContentsMargins(0, 0, 0, 0)
                    hbox.setSpacing(6)

                    label_title = QLabel(f"{key}:")
                    label_title.setFont(QFont("Segoe UI", 10, QFont.Bold))
                    label_title.setStyleSheet("color: #333;")
                    hbox.addWidget(label_title)

                    label_value = QLabel(str(val))
                    label_value.setFont(font)
                    label_value.setStyleSheet("color: #444;")
                    label_value.setWordWrap(True)  # ✅ Kelime taşmalarını engeller
                    label_value.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                    hbox.addWidget(label_value)

                    info_layout.addWidget(box, idx // 2, idx % 2)

            layout.addWidget(info_box)
            self.server_info_frame.setVisible(True)

        except Exception as e:
            import traceback
            print("Error in display_server_info:", traceback.format_exc())

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_pos = event.globalPos()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(self.pos() + event.globalPos() - self.drag_pos)
            self.drag_pos = event.globalPos()
            event.accept()

    def mouseReleaseEvent(self, event):
        self.drag_pos = None
        event.accept()
    def create_progress_overlay(self):
        overlay = QWidget()

        layout = QVBoxLayout(overlay)
        layout.setAlignment(Qt.AlignCenter)

        self.progress_label = QLabel("Generating PDF, please wait...")
        self.progress_label.setStyleSheet("color: black; font-size: 16px;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(300)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                color: black;
                border: 1px solid #444;
                border-radius: 6px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4caf50;
                width: 20px;
            }
        """)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        return overlay

    def build_logo_panel(self):
        logo_panel = QWidget()
        logo_layout = QVBoxLayout(logo_panel)
        logo_layout.setContentsMargins(40, 40, 40, 40)
        logo_layout.setSpacing(30)
        logo_layout.setAlignment(Qt.AlignTop)

        # Başlık
        #title = QLabel("HCAI")
        #title.setFont(QFont("Segoe UI", 36, QFont.Bold))
        #title.setAlignment(Qt.AlignCenter)

        # Logo
        logo = QLabel()
        pix = QPixmap("dplogo1.png")
        logo.setPixmap(pix.scaled(480, 480, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        logo.setAlignment(Qt.AlignCenter)

        # Exit butonu
        #exit_button = QPushButton("Exit")
        #exit_button.setFont(QFont("Segoe UI", 12, QFont.Bold))
        #exit_button.setStyleSheet("""
        #    QPushButton {
        #        background-color: #c62828;
        #        color: white;
        #        padding: 12px;
        #        border-radius: 8px;
        #    }
        #    QPushButton:hover {
        #        background-color: #e53935;
        #    }
        #""")
        #exit_button.setFixedHeight(50)
        #exit_button.clicked.connect(QApplication.instance().quit)

        #logo_layout.addWidget(title)
        logo_layout.addWidget(logo)
        logo_layout.addStretch()
        # logo_layout.addWidget(exit_button)


        logo_panel.setStyleSheet("background-color: #a9cfeb;")
        return logo_panel

    def switch_to_checklist(self):
        self.stack.setCurrentIndex(2)

    def switch_to_connection(self):
        self.stack.setCurrentIndex(1)

    def switch_to_consent(self):
        self.stack.setCurrentIndex(0)

class ConsentWindow(QWidget):
    def __init__(self, switch_callback):
        super().__init__()

        self.switch_callback = switch_callback  # Sayfa geçiş fonksiyonu
        self.setFixedWidth(720)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.initUI()

    def initUI(self):

        # Main Layout
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        self.setLayout(layout)
        # Scrollable area for the long consent text
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)

        # Consent text with full legal details
        self.terms_text = QTextEdit(self)
        self.terms_text.setPlainText(
            "📜 **Privacy & Data Processing Consent**\n\n"
            "Welcome to [DAPLAIT HEALTH CHECK]. Before you proceed, please carefully read and agree to the terms below regarding "
            "the collection, processing, and protection of your data in accordance with **General Data Protection Regulation (GDPR), "
            "KVKK (Turkey), DIFC Data Protection Law (UAE), and the Australian Privacy Act**.\n\n"
            "🔍 **Why Do We Collect Your Data?**\n"
            "Our platform provides **database health monitoring and anomaly detection services** to improve performance and security. "
            "To achieve this, certain system-related data must be processed.\n\n"
            "✅ **Types of Data We Collect:**\n"
            "- **Server Performance Metrics:** CPU, Memory, Disk Usage, Network Traffic.\n"
            "- **System Logs & Error Reports:** To analyze database health and detect failures.\n"
            "- **Database Query Execution Time & Performance Metrics.**\n"
            "- **User Access Credentials (for authentication purposes only).**\n"
            "- **Metadata related to database health and anomalies.**\n\n"
            "🔒 **How We Use Your Data:**\n"
            "All collected data is used exclusively for:\n"
            "- **Real-time monitoring of database health and performance.**\n"
            "- **Identifying slow or failing queries and suggesting optimizations.**\n"
            "- **Detecting anomalies and security threats.**\n"
            "- **Compliance with regulatory and security standards.**\n"
            "- **Enhancing user experience through analytics-based recommendations.**\n\n"
            "🔐 **Data Security & Retention:**\n"
            "Your data is stored in encrypted and access-controlled environments. Key measures include:\n"
            "- **AES-256 encryption** for data at rest.\n"
            "- **End-to-end SSL/TLS encryption** for data transmission.\n"
            "- **Role-based access control (RBAC)** ensuring only authorized personnel access the data.\n"
            "- **Automated data retention policy:** Data will be **deleted or anonymized after [Retention Period]**.\n\n"
            "📜 **Your Legal Rights Under Data Protection Laws:**\n"
            "According to GDPR (EU & UK), KVKK (Turkey), DIFC (UAE), and the Australian Privacy Act, you have the right to:\n"
            "✔ **Access Your Data:** Request a copy of the data stored about you.\n"
            "✔ **Correct Inaccurate Data:** Request updates to incorrect or incomplete data.\n"
            "✔ **Delete Your Data:** Request permanent deletion of your personal data.\n"
            "✔ **Restrict Processing:** Request limitations on how your data is processed.\n"
            "✔ **Object to Processing:** Opt out of certain uses of your data.\n"
            "✔ **Data Portability:** Request transfer of your data to another service provider.\n\n"
            "📞 **Contact & Support:**\n"
            "For any privacy concerns, you can reach us at **  ** or visit our full **[Privacy Policy]**.\n\n"
            "🔹 By checking the box below, you acknowledge and consent to the processing of your data under the above terms."
        )
        self.terms_text.setReadOnly(True)

        # Adding the consent text inside a scrollable area
        self.scroll_area.setWidget(self.terms_text)

        # User agreement checkbox
        self.consent_checkbox = QCheckBox('I have read and agree to the above terms.', self)

        # Continue button
        self.continue_button = QPushButton('Continue', self)
        self.continue_button.clicked.connect(self.check_consent)
        self.continue_button.setStyleSheet("""
                   QPushButton {
                       background-color: #e53935;
                       color: white;
                       padding: 12px;
                       border-radius: 8px;
                   }
                   QPushButton:hover {
                       background-color: #e53935;
                   }
                """)
        # Exit butonu
        self.exit_button = QPushButton("Exit")
        self.exit_button.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.exit_button.setStyleSheet("""
           QPushButton {
               background-color: #c62828;
               color: white;
               padding: 12px;
               border-radius: 8px;
           }
           QPushButton:hover {
               background-color: #e53935;
           }
        """)
        self.exit_button.setFixedHeight(50)
        self.exit_button.clicked.connect(QApplication.instance().quit)


        # Adding widgets to the layout
        layout.addWidget(self.scroll_area)
        layout.addWidget(self.consent_checkbox)
        layout.addWidget(self.continue_button)
        layout.addWidget(self.exit_button)
        self.setLayout(layout)

    def check_consent(self):
        """Ensures the user provides explicit consent before proceeding."""
        if self.consent_checkbox.isChecked():
            self.switch_callback()  # sayfayı değiştir
        else:
            QMessageBox.warning(
                self, 'Consent Required',
                'You must accept the terms to proceed.',
                QMessageBox.Ok
            )

    def accept_consent(self):
        """Closes the consent window and launches the main application."""
        self.close()
        self.open_main_application()

    def open_main_application(self):
        """Launches the main application after consent approval."""
        self.main_app = SQLServerConnectionUI()
        self.main_app.show()

class CheckListWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setFixedWidth(720)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.is_dark_mode = True
        self.final_progress = False
        self.output_path = None


        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(16)

        # Theme button (top right)
        ##self.theme_button = QPushButton("🌞")
        ##self.theme_button.setFixedSize(38, 38)
        ##self.theme_button.setStyleSheet("""
        ##    QPushButton {
        ##        font-size: 18px;
        ##        border-radius: 19px;
        ##        background-color: #9e9e9e;
        ##        color: white;
        ##    }
        ##    QPushButton:hover {
        ##        background-color: #bdbdbd;
        ##    }
        ##""")
        ##self.theme_button.clicked.connect(self.toggle_theme)

        ##top_bar = QHBoxLayout()
        ##top_bar.addStretch()
        ##top_bar.addWidget(self.theme_button)
        ##main_layout.addLayout(top_bar)

        # Query list
        self.query_list = QListWidget()
        self.query_list.setFont(QFont("Segoe UI", 10))
        self.query_list.setStyleSheet("border-radius: 8px;")
        main_layout.addWidget(self.query_list)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        main_layout.addWidget(self.progress_bar)

        # Output folder button
        self.open_folder_button = QPushButton("Open Output Folder")
        self.open_folder_button.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.open_folder_button.setDisabled(True)
        self.open_folder_button.setFixedHeight(40)
        self.open_folder_button.clicked.connect(self.open_output_folder)
        main_layout.addWidget(self.open_folder_button)



        self.setLayout(main_layout)
        #self.apply_theme()

    ##def apply_theme(self):
    ##    if self.is_dark_mode:
    ##        self.setStyleSheet("""
    ##            QWidget {
    ##                background-color: #121212;
    ##                color: white;
    ##            }
    ##            QListWidget {
    ##                background-color: #1e1e1e;
    ##                color: white;
    ##                border: 1px solid #333;
    ##            }
    ##            QProgressBar {
    ##                background-color: #2e2e2e;
    ##                color: white;
    ##                border: 1px solid #444;
    ##                border-radius: 6px;
    ##                text-align: center;
    ##            }
    ##            QProgressBar::chunk {
    ##                background-color: #4caf50;
    ##                width: 20px;
    ##            }
    ##            QPushButton {
    ##                background-color: #1e88e5;
    ##                color: white;
    ##                border: none;
    ##                border-radius: 6px;
    ##                padding: 10px;
    ##            }
    ##            QPushButton:hover {
    ##                background-color: #42a5f5;
    ##            }
    ##        """)
    ##        self.theme_button.setText("🌞")
    ##    else:
    ##        self.setStyleSheet("""
    ##            QWidget {
    ##                background-color: #f5f5f5;
    ##                color: black;
    ##            }
    ##            QListWidget {
    ##                background-color: white;
    ##                color: black;
    ##                border: 1px solid #ccc;
    ##            }
    ##            QProgressBar {
    ##                background-color: #e0e0e0;
    ##                color: black;
    ##                border: 1px solid #aaa;
    ##                border-radius: 6px;
    ##                text-align: center;
    ##            }
    ##            QProgressBar::chunk {
    ##                background-color: #4caf50;
    ##                width: 20px;
    ##            }
    ##            QPushButton {
    ##                background-color: #43a047;
    ##                color: white;
    ##                border: none;
    ##                border-radius: 6px;
    ##                padding: 10px;
    ##            }
    ##            QPushButton:hover {
    ##                background-color: #66bb6a;
    ##            }
    ##        """)
    ##        self.theme_button.setText("🌙")

    ##def toggle_theme(self):
    ##    self.is_dark_mode = not self.is_dark_mode
    ##    self.apply_theme()

    def update_list(self, query, success):
        self.increment_progress()
        display_text = query[:2] + '*' * 8
        symbol = "✅" if success else "❌"
        self.query_list.addItem(f"{symbol} {display_text}")

    def increment_progress(self):
        if not self.final_progress:
            current_value = self.progress_bar.value()
            new_value = min(99, current_value + 3)
            self.progress_bar.setValue(new_value)

    def finalize_progress(self):
        self.progress_bar.setValue(100)
        self.progress_bar.hide()
        self.final_progress = True

    def set_output_path(self, output_path):
        self.output_path = output_path
        self.open_folder_button.setDisabled(False)

    def open_output_folder(self):
        if self.output_path:
            os.startfile(os.path.dirname(self.output_path))

class SQLServerConnectionUI(QWidget):
    def __init__(self, switch_callback,main_ui):
        super().__init__()
        self.switch_callback = switch_callback  # MainUI'den gelen sayfa geçiş fonksiyonu
        self.main_ui = main_ui
        self.setFixedWidth(720)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # Bu fonksiyonlar initUI()'den once gelmeli
        self.current_user = self.get_windows_user()
        self.connection = None
        self.key = self.load_or_generate_key()
        self.fernet = Fernet(self.key)
        self.checklist_window = None
        self.is_dark_mode = False
        self.initUI()  # UI en sona alınmalıydı

    def log_error(self, exception: Exception, function_name: str = "Unknown"):
        try:
            error_message = f"{datetime.now()} - Error: {str(exception)}\n"
            error_message += f"Function: {function_name}\n"
            error_message += f"Stack Trace:\n{traceback.format_exc()}\n"

            # Log dosyasına yaz
            with open("error_log.txt", "a", encoding="utf-8") as file:
                file.write(error_message + "\n")

        except Exception as log_error_exception:
            print(f"Log writing error: {log_error_exception}")

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
            self.log_error(e, 'load_or_generate_key')
            QMessageBox.critical(self, "Error", f"An error occurred while loading the key file:\n{e}")
            return None

    def encrypt_data(self, data):
        """Veriyi sifreler."""
        return self.fernet.encrypt(data.encode()).decode()

    def decrypt_data(self, data):
        """sifreli veriyi çözer."""
        return self.fernet.decrypt(data.encode()).decode()

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode

        if self.is_dark_mode:
            # Dark Theme
            self.setStyleSheet("""
                QWidget {
                    background-color: #f4f4f4;
                    color: black;
                }
                QLineEdit, QComboBox {
                    background-color: white;
                    color: black;
                    border: 1px solid #cccccc;
                    border-radius: 6px;
                    padding: 6px;
                }
                QCheckBox {
                    padding: 4px;
                }
            """)
            self.form_frame.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 30px;
                    padding: 48px;
                }
            """)
            self.theme_button.setText("Dark")
            self.theme_button.setToolTip("Switch to Dark Mode")

        else:
            # Light Theme (as default)
            self.setStyleSheet("""
                QWidget {
                    background-color: #121212;
                    color: white;
                }
                QLineEdit, QComboBox {
                    background-color: #1e1e1e;
                    color: white;
                    border: 1px solid #3a3a3a;
                    border-radius: 6px;
                    padding: 6px;
                }
                QCheckBox {
                    padding: 4px;
                }
            """)
            self.form_frame.setStyleSheet("""
                QFrame {
                    background-color: #1c1c1c;
                    border-radius: 30px;
                    padding: 48px;
                }
            """)
            self.theme_button.setText("Light")
            self.theme_button.setToolTip("Switch to Light Mode")

    ##def mousePressEvent(self, event):
    ##    if event.button() == Qt.LeftButton:
    ##        self.drag_pos = event.globalPos()
    ##        event.accept()
##
    ##def mouseMoveEvent(self, event):
    ##    if event.buttons() == Qt.LeftButton:
    ##        self.move(self.pos() + event.globalPos() - self.drag_pos)
    ##        self.drag_pos = event.globalPos()
    ##        event.accept()
##
    ##def mouseReleaseEvent(self, event):
    ##    self.drag_pos = None
    ##    event.accept()

    def initUI(self):
        try:
            main_layout = QVBoxLayout(self)
            main_layout.setAlignment(Qt.AlignVCenter)  # Ortala

            # Form Alanları
            self.form_frame = QFrame()
            form_layout = QFormLayout()
            form_layout.setSpacing(16)

            label_font = QFont("Segoe UI", 16, QFont.Bold)  # Label font büyütüldü
            input_font = QFont("Segoe UI", 14)

            self.server_name = QLineEdit()
            self.server_name.setPlaceholderText("e.g. DESKTOP-U2CP0FH")
            self.server_name.setFont(input_font)
            self.server_name.setMinimumHeight(40)

            self.auth_type = QComboBox()
            self.auth_type.addItems(["SQL Server Authentication", "Windows Authentication"])
            self.auth_type.setFont(input_font)
            self.auth_type.setMinimumHeight(40)
            self.auth_type.currentIndexChanged.connect(self.toggle_authentication)

            self.login = QLineEdit()
            self.login.setPlaceholderText("Enter username")
            self.login.setFont(input_font)
            self.login.setMinimumHeight(40)

            self.password = QLineEdit()
            self.password.setPlaceholderText("Enter password")
            self.password.setFont(input_font)
            self.password.setEchoMode(QLineEdit.Password)
            self.password.setMinimumHeight(40)

            self.remember_password = QCheckBox("Remember password")
            self.remember_password.setFont(input_font)
            self.remember_password.stateChanged.connect(self.save_credentials)

            server_label = QLabel("Server name:")
            server_label.setFont(label_font)
            server_label.setMinimumHeight(40)

            auth_label = QLabel("Authentication:")
            auth_label.setFont(label_font)
            auth_label.setMinimumHeight(40)

            login_label = QLabel("Login:")
            login_label.setFont(label_font)
            login_label.setMinimumHeight(40)

            password_label = QLabel("Password:")
            password_label.setFont(label_font)
            password_label.setMinimumHeight(40)
            form_layout.addRow(server_label, self.server_name)
            form_layout.addRow(auth_label, self.auth_type)
            form_layout.addRow(login_label, self.login)
            form_layout.addRow(password_label, self.password)
            form_layout.addRow(QLabel(""), self.remember_password)

            self.form_frame.setLayout(form_layout)

            main_layout.addWidget(self.form_frame)

            # Butonlar
            button_layout = QHBoxLayout()
            # Export button
            self.export_button = QPushButton("Run Scripts and Export PDF")
            self.export_button.setFont(QFont("Segoe UI", 11, QFont.Bold))
            self.export_button.setStyleSheet("""
                        QPushButton {
                            background-color: #1e88e5;
                            color: white;
                            padding: 12px 20px;
                            border-radius: 8px;
                        }
                        QPushButton:hover {
                            background-color: #42a5f5;
                        }
                    """)

            main_layout.addStretch()
            main_layout.addWidget(self.export_button)

            self.export_button.setVisible(False)
            self.export_button.clicked.connect(self.run_scripts_and_export)


            self.connect_button = QPushButton("Connect")  # Connect artık Next görevi görecek
            self.connect_button.setFont(QFont("Segoe UI", 11, QFont.Bold))
            self.connect_button.setStyleSheet("""
                QPushButton {
                    background-color: #43a047;
                    color: white;
                    padding: 12px 20px;
                    border-radius: 8px;
                }
                QPushButton:hover {
                    background-color: #66bb6a;
                }
            """)

            button_layout.addWidget(self.connect_button)
            self.connect_button.clicked.connect(self.test_connection)

            button_layout.setSpacing(20)
            main_layout.addLayout(button_layout)
            # Exit butonu
            self.exit_button = QPushButton("Exit")
            self.exit_button.setFont(QFont("Segoe UI", 12, QFont.Bold))
            self.exit_button.setStyleSheet("""
                       QPushButton {
                           background-color: #c62828;
                           color: white;
                           padding: 12px;
                           border-radius: 8px;
                       }
                       QPushButton:hover {
                           background-color: #e53935;
                       }
                    """)
            self.exit_button.setFixedHeight(50)
            self.exit_button.clicked.connect(QApplication.instance().quit)
            button_layout.addWidget(self.exit_button)
            self.load_saved_credentials()
            self.toggle_authentication()

        except Exception as e:
            self.log_error(e, "initUI")


    def get_windows_user(self):
        try:
            """Windows kullanıcısını alma."""
            username = getpass.getuser()
            domain = os.environ.get('USERDOMAIN', '')
            if domain:
                return f"{domain}\{username}"
            return username
        except Exception as e:
            self.log_error(e, "get_windows_user")
            return

    def toggle_authentication(self):
        try:
            """Authentication türüne göre Login ve Password kutularını kontrol eder."""
            if self.auth_type.currentText() == "Windows Authentication":
                self.login.setText(self.get_windows_user())
                self.login.setDisabled(True)
                self.password.setDisabled(True)
            else:
                self.login.setText("")
                self.login.setDisabled(False)
                self.password.setDisabled(False)
        except Exception as e:
            self.log_error(e, "toggle_authentication")
            return

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
            self.log_error(e, "save_credentials")
            QMessageBox.critical(self, "Error", f"An error occurred while saving the information:\n{e}")
            self.remember_password.setChecked(False)
            return

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
            self.log_error(e, "load_saved_credentials")
            QMessageBox.critical(self, "Warning", f"An error occurred while loading the saved information:\n{e}")
            return

    def save_to_file(self, data):
        try:
            """Bilgileri düz metin olarak dosyaya kaydeder."""
            os.makedirs("credentials_folder", exist_ok=True)
            with open("credentials_folder/credentials.json", "w") as file:
                json.dump(data, file)
        except Exception as e:
            self.log_error(e, "save_to_file")
            QMessageBox.critical(self, "Warning", f"An error occurred while loading the saved information:\n{e}")
            return

    def test_connection(self):
        """SQL Server bağlantısını test eder ve uygun ODBC sürücüsünü otomatik olarak bulur."""
        server = self.server_name.text()
        auth = self.auth_type.currentText()

        # Kullanılabilecek olası ODBC sürücüleri (en güncellerden eskiye doğru sıralı)
        odbc_drivers = [
            "ODBC Driver 18 for SQL Server",
            "ODBC Driver 17 for SQL Server",
            "ODBC Driver 13 for SQL Server",
            "ODBC Driver 11 for SQL Server",
            "SQL Server Native Client 11.0",
            "SQL Server"  # En eski sürüm
        ]

        # Sistemde yüklü ODBC sürücülerini al
        available_drivers = pyodbc.drivers()
        valid_driver = None

        # Mevcut sürücülerden en güncel olanı seç
        for driver in odbc_drivers:
            if driver in available_drivers:
                valid_driver = driver
                break  # İlk bulunan sürücüyü kullan

        if not valid_driver:
            QMessageBox.critical(self, "Error", "No suitable SQL Server ODBC driver found!")
            return

        try:
            if auth == "Windows Authentication":
                # Windows Authentication ile bağlantı
                self.conn_str = f"DRIVER={{{valid_driver}}};SERVER={server},1433;Trusted_Connection=yes;autocommit=True;"
            else:
                # SQL Server Authentication ile bağlantı
                login = self.login.text()
                password = self.password.text()
                self.conn_str = f"DRIVER={{{valid_driver}}};SERVER={server};UID={login};PWD={password};autocommit=True;"

            self.connection = pyodbc.connect(self.conn_str, autocommit=True)
            cursor = self.connection.cursor()
            QMessageBox.information(self, "SUCCESS",
                                    f"SQL Server connection successful!")

            # Server info SQL script'i dosyadan oku
            try:
                with open("serverinfo.txt", "r", encoding="utf-8") as f:
                    server_info_sql = f.read()

                try:
                    cursor.execute(server_info_sql)
                    cursor.execute("SELECT * FROM #ServerInfo")
                    columns = [desc[0] for desc in cursor.description]
                    rows = cursor.fetchall()
                    # MainUI çağrısı:
                    self.main_ui.display_server_info(rows, columns)

                except Exception as e:
                    self.log_error(e, "execute_serverinfo_statements")

            except Exception as e:
                self.log_error(e, "execute_serverinfo_statements")
            self.connect_button.setVisible(False)
            self.form_frame.setVisible(False)
            self.export_button.setVisible(True)



            self.connect_button.clicked.connect(self.switch_callback)  # Direkt geçiş yapar
            #self.export_button.setDisabled(False)

        except Exception as e:
            self.log_error(e, "Tried ODBC: {valid_driver} or failed test_connection")
            QMessageBox.critical(self, "Error", f"Connection failed:\n{e}\nTried ODBC: {valid_driver}")
            return

    def get_resource_path(self, relative_path):
        try:
            """Kaynak dosyaların dogru yolunu döndürür."""
            if getattr(sys, 'frozen', False):  # Eger exe olarak çalısıyorsa
                base_path = sys._MEIPASS  # PyInstaller'ın geçici dosyaları koydugu klasör
            else:  # Python dosyası olarak çalısıyorsa
                base_path = os.path.dirname(os.path.abspath(__file__))
            return os.path.join(base_path, relative_path)
        except Exception as e:
            self.log_error(e, "get_resource_path")
            return

    def clean_text(self, text):
        try:
            if isinstance(text, str):
                return text.encode('utf-8', errors='replace').decode('utf-8')
            return text
        except Exception as e:
            self.log_error(e, "create_gauge_chart")
            return

    def clean_column(self, value):
        try:
            return value.replace(" ", "").lower()
        except Exception as e:
            self.log_error(e, "clean_column")
            return

    def clean_sheet_name(self, sheet_name):
        try:
            return sheet_name.replace("SecHC_", "").lower()
        except Exception as e:
            self.log_error(e, "clean_sheet_name")
            return

    def score_table(self,summary_df):
        try:
            data = {
                "Parameter": [
                    "AlwaysOn", "SPCU", "SQLServerVersion", "MaxMemory", "MinMemory", "SQLFiles", "TempDB",
                    "LogFiles", "PowerShellOutput", "LocalSecurityPolicy", "ServerConfig", "Antivirus",
                    "CompressionBackup", "BackupManagement", "DatabaseSize", "BackupStats", "Deadlock", "CheckDB",
                    "PageVerify", "CompatibilityLevel", "AutoShrink", "AutoClose", "RecoveryModel", "Storage",
                    "SystemDatabases", "JobHistory", "CPU", "ServiceAccount", "ServiceAccountPermission", "SaAccount",
                    "OrphanUser", "VLFCount", "EmptyPasswordLogins", "PolicyNotCheckedLogins", "SamePasswordLogins",
                    "DisableLogins", "Sysadminlogin", "SQLServerBrowserService", "XpCmdShell", "UpdateStats",
                    "ReIndex", "LeftoverFakeIndex", "ClusteredIndex", "MissingIndex", "BadIndex"
                ],
                "Warning": [
                    2, 3, 3, 2, 1, 2, 2, 2, 1, 1, 2, 1, 1, 1, 1, 1, 1, 2, 1, 2, 1, 1, 1, 2, 1, 1, 2, 1, 1, 2,
                    1, 2, 1, 1, 3, 1, 2, 1, 1, 3, 3, 2, 2, 2, 1
                ],
                "Failed": [
                    2, 3, 3, 3, 1, 3, 3, 2, 1, 2, 2, 1, 1, 3, 2, 2, 2, 3, 2, 2, 2, 2, 3, 2, 1, 3, 2, 2, 1, 2,
                    2, 3, 1, 1, 3, 3, 3, 1, 1, 4, 4, 3, 3, 3, 2
                ]
            }

            # Create the DataFrame
            scoring_df = pd.DataFrame(data)
            merged_df = summary_df.merge(scoring_df, left_on='control_column_name', right_on='Parameter', how='left')
            merged_df.drop('Parameter', axis=1, inplace=True)

            merged_df['Score'] = merged_df.apply(
                lambda row: row['Failed'] if row['status'] == 0 else (row['Warning'] if row['status'] == 2 else 0), axis=1)
            total_score = merged_df['Score'].sum()


            return total_score
        except Exception as e:
            self.log_error(e, "score_table")
            return


    def generate_dataframe_from_excel(self, excel_path):
        try:
            import pandas as pd

            # Initialize an empty list to collect rows
            rows = []

            # Excel dosyasını oku
            excel_data = pd.ExcelFile(excel_path)
            ServerInfoControl = pd.read_excel(excel_data, sheet_name="ServerInfo")
            version = ServerInfoControl['Version']
            version = str(version.iloc[0])


            # Sheet isimlerine göre analiz yap
            for sheet_name in excel_data.sheet_names:
                sheet_df = pd.read_excel(excel_data, sheet_name=sheet_name)


                if sheet_name == "VLFCountResults":
                    try:
                        rows.append({"control_column_name": "Antivirus", "status": 2})

                        failed = sheet_df[sheet_df['VLFCount'] > 200]
                        warning = sheet_df[(sheet_df['VLFCount'] > 120) & (sheet_df['VLFCount'] <= 200)]
                        success = sheet_df[sheet_df['VLFCount'] <= 120]

                        # Durumları kontrol et
                        if not failed.empty:
                            status = 0  # Eğer bir tane bile failed varsa, status 0
                        elif not warning.empty:
                            status = 2  # Failed yok, fakat warning varsa, status 2
                        else:
                            status = 1  # Ne failed ne warning varsa, status 1

                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "VLFCount", "status": status})
                    except Exception as e:
                        self.log_error(e, "VLFCountResults")

                elif sheet_name == "MissingIndexCount":
                    try:
                        # IndexCount değeri 10'dan fazla olan veritabanlarını filtrele
                        high_missing_indexes = sheet_df[sheet_df["IndexCount"] > 10]
                        # Bu veritabanlarının sayısını hesapla
                        db_count = len(high_missing_indexes)
                        # Durumu belirle ve rows listesine ekle
                        if db_count > 10:
                            rows.append({
                                "control_column_name": "MissingIndex",
                                "status": 0
                            })
                        elif 5 <= db_count <= 10:
                            rows.append({
                                "control_column_name": "MissingIndex",
                                "status": 2
                            })
                        else:
                            rows.append({
                                "control_column_name": "MissingIndex",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "MissingIndex")

                elif sheet_name == "ClusteredIndexes":
                    try:
                        # Clustered index bulunmayan tabloları filtrele (IndexType değeri boş veya NaN olanlar)
                        tables_without_clustered_index = sheet_df[sheet_df["IndexType"].isna()]

                        # Clustered index bulunmayan tablo sayısını hesapla
                        missing_index_count = len(tables_without_clustered_index)

                        # Durumu belirle ve rows listesine ekle
                        if missing_index_count > 20:
                            rows.append({
                                "control_column_name": "ClusteredIndex",
                                "status": 0
                            })
                        elif 10 <= missing_index_count <= 20:
                            rows.append({
                                "control_column_name": "ClusteredIndex",
                                "status": 2
                            })
                        else:
                            rows.append({
                                "control_column_name": "ClusteredIndex",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "ClusteredIndexes")
                elif sheet_name == "SQLServerBrowser":
                    try:

                        if str(version).startswith("12"):
                            rows.append({
                                "control_column_name": "SQLServerBrowserService",
                                "status": 3
                            })
                        else:
                            # Servis durumunu al
                            sql_browser_state = sheet_df["State"].iloc[0]  # İlk satırdaki durum

                            # Kontrol mekanizması ve rows listesine ekleme
                            if sql_browser_state.lower() == "running":
                                rows.append({
                                    "control_column_name": "SQLServerBrowserService",
                                    "status": 0
                                })
                            else:
                                rows.append({
                                    "control_column_name": "SQLServerBrowserService",
                                    "status": 1
                                })
                    except Exception as e:
                        self.log_error(e, "ClusteredIndexes")

                elif sheet_name == "UpdateStats":
                    try:
                        # Bugünün tarihini al
                        today = datetime.today()
                        # LastStatsUpdateTime sütununu datetime formatına çevir
                        sheet_df["LastStatsUpdateTime"] = pd.to_datetime(sheet_df["LastStatsUpdateTime"],
                                                                            errors='coerce')
                        # 1 haftadan fazla geçmiş olanları filtrele
                        outdated_tables = sheet_df[today - sheet_df["LastStatsUpdateTime"] > pd.Timedelta(days=7)]
                        # Güncellenmesi gereken tablo sayısını hesapla
                        outdated_count = len(outdated_tables)
                        # Durumu belirle ve rows listesine ekle
                        if outdated_count > 10:
                            rows.append({
                                "control_column_name": "UpdateStats",
                                "status": 0
                            })
                        elif 3 <= outdated_count <= 10:
                            rows.append({
                                "control_column_name": "UpdateStats",
                                "status": 2
                            })
                        else:
                            rows.append({
                                "control_column_name": "UpdateStats",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "UpdateStats")

                elif sheet_name =="ReIndex":
                    try:
                        # %50'den az fragmantasyona sahip indexleri filtrele
                        low_fragmentation_indexes = sheet_df[sheet_df["AvgFragmentationPercent"] < 50]
                        # Düşük fragmantasyona sahip index sayısını hesapla
                        low_frag_count = len(low_fragmentation_indexes)

                        # Durumu belirle ve rows listesine ekle
                        if low_frag_count > 10:
                            rows.append({
                                "control_column_name": "ReIndex",
                                "status": 0
                            })
                        elif 3 <= low_frag_count <= 10:
                            rows.append({
                                "control_column_name": "ReIndex",
                                "status": 2
                            })
                        else:
                            rows.append({
                                "control_column_name": "ReIndex",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "ReIndex")
                elif sheet_name == "LeftoverFakeIndex":
                    try:
                        # Toplam index sayısını hesapla
                        index_count = len(sheet_df)
                        # Durumu belirle ve rows listesine ekle
                        if index_count > 10:
                            rows.append({
                                "control_column_name": "LeftoverFakeIndex",
                                "status": 0
                            })
                        elif 3 <= index_count <= 10:
                            rows.append({
                                "control_column_name": "LeftoverFakeIndex",
                                "status": 2
                            })
                        else:
                            rows.append({
                                "control_column_name": "LeftoverFakeIndex",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "LeftoverFakeIndex")
                elif sheet_name == "Sysadminlogin":
                    try:
                        # Yalnızca "sysadmin" rolüne sahip olanları filtrele
                        sysadmin_users = sheet_df[sheet_df["AssociatedServerRole"] == "sysadmin"]

                        # "NT SERVICE" veya "NT Service" ile başlayanları hariç tut
                        filtered_users = sysadmin_users[
                            ~sysadmin_users["Name"].str.startswith(("NT SERVICE", "NT Service"))]

                        # Kullanıcı sayısını hesapla
                        user_count = len(filtered_users)

                        # Durumu belirle ve rows listesine ekle
                        if user_count < 4:
                            rows.append({
                                "control_column_name": "Sysadminlogin",
                                "status": 1
                            })
                        elif 4 <= user_count <= 7:
                            rows.append({
                                "control_column_name": "Sysadminlogin",
                                "status": 2
                            })
                        else:
                            rows.append({
                                "control_column_name": "Sysadminlogin",
                                "status": 0
                            })
                    except Exception as e:
                        self.log_error(e, "LeftoverFakeIndex")


                elif sheet_name == "DBCCResults":
                    try:
                        DatabaseSizeInfo = pd.read_excel(excel_data, sheet_name="DatabaseSizeInfo")

                        # Birleştirme işlemi
                        merged_df = pd.merge(sheet_df, DatabaseSizeInfo, left_on='DBName', right_on='DatabaseName',
                                             how='inner')

                        # Sonuç sütunu ekleyerek koşulları doğrudan kontrol et
                        merged_df['ConditionResult'] = [
                            "FAILED: DBCC CHECKDB required" if "1900-01-01" in str(row["LastCleanDBCCDate"]) and row[
                                "TotalSizeGB"] > 10
                            else "WARNING: Heavy load possible" if row["TotalSizeGB"] > 50
                            else "SUCCESS"
                            for index, row in merged_df.iterrows()
                        ]

                        # Failed olan sonuçları filtrele
                        failing_databases = merged_df[merged_df['ConditionResult'].str.contains("FAILED")][
                            "DatabaseName"].tolist()

                        # Sonuçları rows listesine ekle
                        if failing_databases:
                            rows.append({
                                "control_column_name": "CheckDB",
                                "status": 0
                            })
                        else:
                            rows.append({
                                "control_column_name": "CheckDB",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "CheckDB")
                elif sheet_name == "BackupLogRunning":
                    try:
                        DatabaseSizeInfo = pd.read_excel(excel_data, sheet_name="DatabaseSizeInfo")
                        BackupDetails = pd.read_excel(excel_data, sheet_name="BackupDetails")
                        DatabaseLogInfo = pd.read_excel(excel_data, sheet_name="DatabaseLogInfo")

                        # Log boyutu eşiği (10 GB) ve zaman eşiği (son 6 saat)
                        log_threshold_gb = 10
                        time_threshold = datetime.now() - timedelta(hours=6)

                        # Recovery Model FULL olan veritabanlarını filtrele
                        full_recovery_dbs = DatabaseLogInfo[DatabaseLogInfo['Recovery Model'] == 'FULL'][
                            'Database Name'].tolist()

                        # Başlangıçta tüm durumları başarılı kabul edelim
                        overall_status = 1  # Success olarak başlatıyoruz

                        # Aktif çalışan backup job kontrolü
                        if sheet_df.empty:
                            overall_status = 1  # Backup işlemi yoksa success sayılabilir

                        # Log Backup ve LogSize kontrolleri
                        for db in full_recovery_dbs:
                            log_size = DatabaseSizeInfo.loc[DatabaseSizeInfo['DatabaseName'] == db, 'LogSizeGB'].max()
                            if log_size and log_size > log_threshold_gb:
                                db_logs = BackupDetails[(BackupDetails['database_name'] == db) &
                                                        (BackupDetails['backup_type'].str.lower() == 'log')]
                                db_logs['backup_finish_date'] = pd.to_datetime(db_logs['backup_finish_date'],
                                                                               errors='coerce')
                                recent_log_backup = db_logs['backup_finish_date'].max()

                                if pd.isna(recent_log_backup) or recent_log_backup < time_threshold:
                                    overall_status = 0  # Failed durumu
                                    break  # Bir tane bile failed varsa hemen durdur

                                overall_status = max(overall_status, 2)  # Warning durumuna geçiş olabilir

                        # Recovery Model SIMPLE olan veritabanlarını filtrele
                        simple_recovery_dbs = DatabaseLogInfo[DatabaseLogInfo['Recovery Model'] == 'SIMPLE'][
                            'Database Name'].tolist()

                        for db in simple_recovery_dbs:
                            db_full_backup = BackupDetails[(BackupDetails['database_name'] == db) &
                                                           (BackupDetails['backup_type'].str.lower() == 'full')]
                            db_full_backup['backup_finish_date'] = pd.to_datetime(db_full_backup['backup_finish_date'],
                                                                                  errors='coerce')
                            recent_full_backup = db_full_backup['backup_finish_date'].max()

                            if pd.isna(recent_full_backup) or recent_full_backup < time_threshold:
                                overall_status = 0  # Failed durumu
                                break  # Bir tane bile failed varsa hemen durdur

                            overall_status = max(overall_status, 2)  # Warning durumuna geçiş olabilir

                        # Tek bir sonuç döndürme
                        rows.append({
                            "control_column_name": "RecoveryModel",
                            "status": overall_status
                        })

                    except Exception as e:
                        self.log_error(e, "RecoveryModel")

                elif sheet_name == "EmptyPasswordLogins":
                    try:
                        SamePasswordLogins = pd.read_excel(excel_data, sheet_name="SamePasswordLogins")
                        PolicyNotCheckedLogins = pd.read_excel(excel_data, sheet_name="PolicyNotCheckedLogins")

                        # DF24 kontrolü: Boş şifre kullanan kullanıcılar
                        if not sheet_df.empty:
                            rows.append({
                                "control_column_name": "EmptyPasswordLogins",
                                "status": 0
                            })
                        else:
                            rows.append({
                                "control_column_name": "EmptyPasswordLogins",
                                "status": 1
                            })
                        # DF25 kontrolü: Kullanıcı adını şifre olarak kullananlar
                        if not SamePasswordLogins.empty:
                            tag = 'Failed'
                            rows.append({
                                "control_column_name": "SamePasswordLogins",
                                "status": 0
                            })
                        else:
                            rows.append({
                                "control_column_name": "SamePasswordLogins",
                                "status": 1
                            })
                        # DF26 kontrolü: Şifre policy kontrolü kapalı olanlar
                        if not PolicyNotCheckedLogins.empty:
                            tag = 'Failed'
                            rows.append({
                                "control_column_name": "PolicyNotCheckedLogins",
                                "status": 0
                            })
                        else:
                            rows.append({
                                "control_column_name": "PolicyNotCheckedLogins",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "EmptyPasswordLogins-SamePasswordLogins-PolicyNotCheckedLogins")


                elif sheet_name == "ServerInfo": #version kontrolü
                    try:
                        sqlversion = sheet_df['SQL Version']
                        version_number = re.search(r'\d+', sqlversion[0]).group()  # Sadece yıl kısmını al
                        cu = sheet_df['Cumulative Update']
                        # SQL Server versiyonundan ilgili CSV dosya adını oluştur
                        base_path = os.path.dirname(os.path.abspath(__file__))
                        csv_filename = f"spcu\\sqlserver-{version_number}.csv"
                        csv_filename_original = os.path.join(base_path, csv_filename)

                        # CSV dosyasını oku
                        cu_data = pd.read_csv(csv_filename_original)

                        # En güncel CU'yu al (ilk satırdaki CU adı)
                        latest_cu = cu_data.iloc[0, 0]

                        # Eğer mevcut CU, en güncel CU ile aynıysa günceliz
                        if cu[0] == latest_cu:
                            status = 0

                        else:
                            # Mevcut CU kaç sıra geride onu hesapla
                            cu_list = cu_data.iloc[:, 0].tolist()
                            if cu[0] in cu_list:
                                cu_gerilik = abs(cu_list.index(latest_cu) - cu_list.index(cu[0]))  # Mutlak değerini al
                                # Durum kontrolü
                                if cu_gerilik > 5:
                                    status = 0
                                elif 3 <= cu_gerilik <= 5:
                                    status = 2
                                else:
                                    status = 1

                        rows.append({"control_column_name": "SPCU", "status": status})
                    except Exception as e:
                        self.log_error(e, "SPCU")
                    try:
                        supported_versions = ["SQL Server 2019", "SQL Server 2022"]
                        if not sheet_df.empty:
                            # SQL Version sütununu kontrol et
                            sql_version = sheet_df['SQL Version'].iloc[0]  # İlk satırdaki SQL Version değeri

                            if sql_version in supported_versions:
                                status = 1  # SQL sürümü destekleniyor
                            else:
                                status = 0  # SQL sürümü desteklenmiyor
                            # Durumu rows listesine ekle
                            rows.append({"control_column_name": "SQLServerVersion", "status": status})
                    except Exception as e:
                        self.log_error(e, "SQLServerVersion")

                elif sheet_name == "ServerConfiguration":
                    try:
                        if not sheet_df.empty:
                            if sheet_df['ConfigValue'][2] == 1:
                                oran = sheet_df['ConfigValue'][1] / sheet_df['ConfigValue'][0]
                                if oran == 0.025:
                                    status = 1  # Başarılı
                                else:
                                    status = 0  # Başarısız
                            else:
                                status = 0  # Başarısız
                            # Durumu rows listesine ekle
                            rows.append({"control_column_name": "ServerConfig", "status": status})
                    except Exception as e:
                        self.log_error(e, "ServerConfig")

                elif sheet_name == "MinServerMemory":
                    try:
                        if not sheet_df.empty:
                            degisken = sheet_df['config_value'][0]
                            if degisken == 0:
                                status = 1  # Başarılı
                            else:
                                status = 0  # Başarısız
                            # Durumu rows listesine ekle
                            rows.append({"control_column_name": "MinMemory", "status": status})
                    except Exception as e:
                        self.log_error(e, "MinMemory")

                elif sheet_name == "MaxServerMemory":
                    import pandas as pd

                    try:
                        # Excel dosyasından verileri oku
                        TotalMemory = pd.read_excel(excel_data, sheet_name="TotalMemory")
                        total_memory_mb = TotalMemory['TotalMemory_MB'][0]
                        default = 2147483647
                        config_value = sheet_df['config_value'][0]

                        # config_value değerini kontrol et
                        if config_value == default:
                            status = 0  # Başarısız
                        else:
                            diff = total_memory_mb - config_value

                            # config_value için belirtilen aralıklara göre kontrol et
                            if diff <= 6000:
                                status = 0  # Başarısız
                            else:
                                if 0 <= config_value < 64000:
                                    if diff < 8:
                                        status = 0  # Başarısız
                                    else:
                                        status = 1  # Potansiyel başarı
                                elif 64000 <= config_value < 128000:
                                    if diff < 11:
                                        status = 0  # Başarısız
                                    else:
                                        status = 1  # Potansiyel başarı
                                elif config_value >= 128000:
                                    if not (8 < diff < 17):
                                        status = 0  # Başarısız
                                    else:
                                        status = 1  # Potansiyel başarı

                                # final kontrol total_memory'ye göre
                                if status == 1:
                                    if total_memory_mb > 100000:
                                        if total_memory_mb * 0.9 < config_value:
                                            status = 1  # Başarılı
                                        else:
                                            status = 2  # Başarısız
                                    else:
                                        if total_memory_mb * 0.8 < config_value:
                                            status = 1  # Başarılı
                                        else:
                                            status = 2  # Başarısız

                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "MaxMemory", "status": status})
                    except Exception as e:
                        self.log_error(e, "MaxMemory")



                elif sheet_name == "BackupCompressionInfo":
                    try:
                        if not sheet_df.empty:
                            if sheet_df['CompressionBackup'][0] == 1:
                                status = 1  # Başarılı
                            else:
                                status = 0  # Başarısız
                            # Durumu rows listesine ekle
                            rows.append({"control_column_name": "CompressionBackup", "status": status})
                    except Exception as e:
                        self.log_error(e, "CompressionBackup")

                elif sheet_name == "DatabaseFileInfo":
                    try:

                        if str(version).startswith("12"):
                            rows.append({
                                "control_column_name": "SQLFiles",
                                "status": 3
                            })
                        else:

                            SystemDrive = pd.read_excel(excel_data, sheet_name="SystemDrive")
                            if not SystemDrive.empty:
                                disk_row = SystemDrive.dropna(subset=['Output']).iloc[0]
                                target_disk = disk_row['Output'][0]  # İlk harf: disk adı (C:, D: vs.)
                                exclude_list = ['master', 'model', 'msdb']
                                # `Database Name` sütununda exclude_list içinde olmayanları filtrele
                                df_filtered = sheet_df[~sheet_df['Database Name'].isin(exclude_list)]
                                conflict_rows = df_filtered[df_filtered['physical_name'].str.startswith(target_disk)]

                                if not conflict_rows.empty:
                                    status = 0  # Başarısız
                                else:
                                    status = 1  # Başarılı
                                # Durumu rows listesine ekle
                                rows.append({"control_column_name": "SQLFiles", "status": status})
                    except Exception as e:
                        self.log_error(e, "SQLFiles")

                elif sheet_name == "DatabaseFileInfo": #tempdb
                    try:
                        temp_files_df = sheet_df[
                            sheet_df['name'].str.contains('tempdb', case=False, na=False) &
                            ~sheet_df['name'].str.contains('templog', case=False, na=False)]

                        if temp_files_df['growth'].nunique() == 1:
                            growth_value = temp_files_df['growth'].iloc[0]
                            status = 1  # Başarılı
                        else:
                            most_common_growth = temp_files_df['growth'].mode()[0]
                            differing_rows = temp_files_df[temp_files_df['growth'] != most_common_growth]
                            status = 0  # Başarısız
                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "TempDB", "status": status})
                    except Exception as e:
                        self.log_error(e, "SQLFiles")

                elif sheet_name == "DatabaseSizeInfo":
                    try:
                        failing_databases = []
                        for index, row in sheet_df.iterrows():
                            if row["LogSizeGB"] > 100:  # Log size 100 GB'dan büyükse
                                sheet_df.at[index, 'ConditionResult'] = False
                                failing_databases.append(row["DatabaseName"])
                            elif row["TotalSizeGB"] > 10 and row["LogSizeGB"] > (
                                    row["TotalSizeGB"] * 0.5):  # Total size > 10 GB ve log size > %50
                                sheet_df.at[index, 'ConditionResult'] = False
                                failing_databases.append(row["DatabaseName"])
                            else:  # Diğer durumlar
                                sheet_df.at[index, 'ConditionResult'] = True

                        # Sonuçları kullanıcıya uygun şekilde yazdır
                        if failing_databases:
                            status = 0  # Başarısız
                        else:
                            status = 1  # Başarılı
                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "DatabaseSize", "status": status})
                    except Exception as e:
                        self.log_error(e, "DatabaseSize")


                elif sheet_name == "DeadlockPerformance":
                    try:
                        if not sheet_df.empty:
                            deadlock = sheet_df['CounterValue_Per_Day'].iloc[0]  # İlk deadlock sayısını al
                            if int(deadlock) > 14:
                                status = 0  # Başarısız
                            else:
                                status = 1  # Başarılı
                            # Durumu rows listesine ekle
                            rows.append({"control_column_name": "Deadlock", "status": status})
                    except Exception as e:
                        self.log_error(e, "DatabaseSize")

                elif sheet_name == "DatabasePageVerifyInfo":
                    try:
                        non_checksum_dbs = sheet_df[
                            sheet_df['Page Verify Option'] != 'CHECKSUM']

                        if non_checksum_dbs.empty:
                            status = 1  # Başarılı
                        else:
                            status = 0  # Başarısız
                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "PageVerify", "status": status})
                    except Exception as e:
                        self.log_error(e, "PageVerifyInfo")

                elif sheet_name == "DatabaseCompatibilityInfo":
                    try:
                        # Master veritabanının uyumluluk seviyesini referans olarak belirle
                        reference_level = sheet_df[sheet_df['Database Name'] == 'master'][
                            'DB Compatibility Level'].iloc[0]

                        # FAILED ve WARNING durumlarını belirle
                        failed_dbs = []
                        warning_dbs = []
                        for index, row in sheet_df.iterrows():
                            diff = abs(row['DB Compatibility Level'] - reference_level)
                            if diff > 20:
                                tag = 'Failed'
                                failed_dbs.append(row)
                            elif 10 <= diff <= 20:
                                tag = 'Warning'
                                warning_dbs.append(row)
                            else:
                                sheet_df.at[index, 'Status'] = 'SUCCESS'
                        # Durumu rows listesine ekle
                        status = 0 if tag == 'Failed' else (2 if tag == 'Warning' else 1)
                        rows.append({"control_column_name": "CompatibilityLevel", "status": status})
                    except Exception as e:
                        self.log_error(e, "CompatibilityLevel")

                elif sheet_name == "JobHistory":
                    try:
                        if not sheet_df.empty:
                            success_outputs = []
                            failed_outputs = []
                            warning_outputs = []
                            # Check each row for job status
                            for _, row in sheet_df.iterrows():
                                if row['RunStatus'] == 'Succeeded':
                                    success_outputs.append(f"{row['JobName']} ----> SUCCESS")
                                elif row['RunStatus'] == 'Failed':
                                    if row['Status'] == 'Disable':
                                        success_outputs.append(f"{row['JobName']} ----> SUCCESS")
                                    elif row['Status'] == 'Enable':
                                        failed_outputs.append(f"{row['JobName']} ----> FAILED")
                                else:
                                    if row['Status'] == 'Enable':
                                        warning_outputs.append(f"{row['JobName']} ----> UNKNOWN")
                            # Determine overall job status
                            if failed_outputs:
                                status = 0  # There are failed jobs
                            elif success_outputs and not failed_outputs:
                                status = 1  # All jobs are successful
                            elif warning_outputs and not failed_outputs:
                                status = 2  # Other / unknown status

                            # Append the overall status to the rows list
                            rows.append({"control_column_name": "JobHistory", "status": status})
                    except Exception as e:
                        self.log_error(e, "JobHistory")
                elif sheet_name == "ServerLogins":
                    try:
                        if sheet_df.empty:
                            rows.append({"control_column_name": "DisableLogins", "status": 1})

                        high_privileges = ['sysadmin', 'serveradmin', 'db_owner']
                        issue_found = False

                        for index, row in sheet_df.iterrows():
                            # Replace NaN values with an empty string and split roles
                            server_roles = row['server_roles'] if not pd.isna(row['server_roles']) else ''
                            db_roles = row['db_roles'] if not pd.isna(row['db_roles']) else ''
                            roles = server_roles.split(',') + db_roles.split(',')
                            roles = [role.strip() for role in roles]  # Strip spaces

                            # Check if the user has high privileges
                            if any(role in high_privileges for role in roles):
                                issue_found = True
                                status = 0  # High privileges found, failed status
                            else:
                                status = 2  # Low privileges found, warning status


                        if not issue_found:
                            rows.append({"control_column_name": "DisableLogins", "status": 1})
                        rows.append({"control_column_name": "DisableLogins", "status": status})
                    except Exception as e:
                        self.log_error(e, "DisableLogins")

                elif sheet_name == 'PowerShellOutput':
                    try:
                        if 'Output' not in sheet_df.columns:
                            rows.append({"control_column_name": "PowerShellOutput", "status": 0,})

                        power_scheme = sheet_df['Output'].iloc[0]  # Get the output value from the first row
                        # Check if the power scheme is set to 'High Performance'
                        if "HighPerformance" in power_scheme or "YüksekPerformans" in power_scheme:
                            status = 1  # Correct power scheme, successful status
                        else:
                            status = 2  # Incorrect power scheme, warning status

                        # Append the result and message to the rows list
                        rows.append({"control_column_name": "PowerShellOutput", "status": status})
                    except Exception as e:
                        self.log_error(e, "PowerShellOutput")

                elif sheet_name == 'ORPHANUSER':
                    try:
                        if sheet_df.empty or sheet_df['USERNAME'].isnull().all():
                            status = 1  # No orphan users found, successful status
                        else:
                            orphan_users = sheet_df['USERNAME'].dropna().unique().tolist()
                            if orphan_users:
                                status = 0  # Orphan users found, failed status
                            else:
                                status = 1  # No orphan users found after checking, successful status
                            # Append the overall status to the rows list
                        rows.append({"control_column_name": "OrphanUser", "status": status})
                    except Exception as e:
                        self.log_error(e, "OrphanUser")

                elif sheet_name == "DatabaseAutoCloseInfo":
                    try:
                        # Auto Close ayarı açık (1) olan veritabanlarını filtrele
                        auto_close_on = sheet_df[sheet_df['is_auto_close_on'] == 1]
                        if auto_close_on.empty:
                            status = 1  # Başarılı
                        else:
                            status = 0  # Başarısız
                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "AutoClose", "status": status})
                    except Exception as e:
                        self.log_error(e, "AutoClose")

                elif sheet_name == "AlwaysOnInfo":
                    try:
                        if sheet_df.empty:
                            status = 0  # Standalone modunda çalışıldığını belirten status
                        else:
                            status = 1  # Always On modunun aktif olduğunu belirten status
                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "AlwaysOn", "status": status})
                    except Exception as e:
                        self.log_error(e, "AlwaysOn")

                elif sheet_name =="BackupDetails":
                    try:
                        unique_databases = sheet_df[
                            'database_name'].unique().tolist()  # Benzersiz veritabanı isimlerini liste olarak saklıyoruz

                        for i in range(len(sheet_df)):
                            # Eğer backup_type 'Full Database' ise
                            if sheet_df['backup_type'].iloc[i] == 'Full Database':
                                # database_name değeri unique listesinde bulunuyorsa
                                if sheet_df['database_name'].iloc[i] in unique_databases:
                                    # Listeden kaldır
                                    unique_databases.remove(sheet_df['database_name'].iloc[i])

                        # Backup analiz sonuçlarını rows listesine ekle
                        if unique_databases:
                            status = 0  # Full Backup eksik olan veritabanları için durum
                            rows.append({"control_column_name": "BackupManagement"
                                         , "status": status})
                        else:
                            status = 1  # Tüm veritabanlarında Full Backup mevcut
                            rows.append({"control_column_name": "BackupManagement", "status": status})
                    except Exception as e:
                        self.log_error(e, "BackupManagement")

                elif sheet_name == "BackupStats":
                    try:
                        if sheet_df.empty:
                            # Veri bulunmadığında durumu rows listesine ekle
                            rows.append({"control_column_name": "BackupStats", "status": 0})
                        else:
                            # İlgili sütunları al (DatabaseName hariç)
                            columns = [col for col in sheet_df.columns if col != "DatabaseName"]
                            all_changes = []  # Tüm değişim oranlarını burada saklayacağız

                            # Her sütunun bir önceki sütuna göre değişimini kontrol et
                            for i in range(len(columns) - 1):
                                col1, col2 = columns[i], columns[i + 1]
                                for index, row in sheet_df.iterrows():
                                    val1, val2 = row[col1], row[col2]
                                    # Eğer herhangi bir değer NaN ise atla
                                    if pd.isna(val1) or pd.isna(val2):
                                        continue
                                    # Değişim oranını hesapla
                                    change = ((val1 - val2) / val2) * 100 if val2 != 0 else 0
                                    all_changes.append(abs(change))

                            # Tüm değişim oranlarına bakarak genel bir durum değerlendirmesi yap
                            if all_changes:
                                max_change = max(all_changes)
                                if max_change < 5:
                                    status = 1
                                elif 5 <= max_change < 15:
                                    status = 2
                                else:
                                    status = 0
                            else:
                                status = 1  # Değişim olmadığı varsayılır

                            # Son durumu rows listesine tek bir kayıt olarak ekle
                            rows.append({"control_column_name": "BackupStats", "status": status})
                    except Exception as e:
                        self.log_error(e, "BackupStats")


                elif sheet_name == "SystemDrive":
                    try:
                        if str(version).startswith("12"):
                            rows.append({
                                "control_column_name": "SystemDrive",
                                "status": 3
                            })
                        else:
                            DatabaseFileInfo = pd.read_excel(excel_data, sheet_name="DatabaseFileInfo")
                            # LOG olanları filtrele
                            log_files = DatabaseFileInfo[DatabaseFileInfo["type_desc"] == "LOG"]

                            # SystemDrive DataFrame'inden sistem sürücüsünü al
                            system_drive = sheet_df.iloc[0, 0]  # İlk satırdaki sürücü bilgisi

                            # Log dosyalarını kontrol et ve hatalı kayıt olup olmadığını takip et
                            incorrect_found = False
                            for index, row in log_files.iterrows():
                                physical_path = row["physical_name"]
                                file_drive = physical_path.split(":")[0] + ":"  # Fiziksel adı parçalayarak diski bul

                                # Disk sürücüsü ve yol kontrolü
                                if file_drive == system_drive or not physical_path.startswith(system_drive):
                                    incorrect_found = True  # Hatalı kayıt bulundu
                                    break

                            # Sonuçları rows listesine ekle
                            if incorrect_found:
                                rows.append({"control_column_name": "LogFiles", "status": 0})
                            else:
                                rows.append({"control_column_name": "LogFiles", "status": 1})
                    except Exception as e:
                        self.log_error(e, "LogFiles")

                elif sheet_name == "DatabaseFileInfo":
                    try:
                        # Kontrol edilecek sistem veritabanları
                        system_databases = ["msdb", "model", "master"]

                        # Sistem veritabanlarına ait log veya data dosyalarını filtrele
                        system_db_files = sheet_df[sheet_df["Database Name"].isin(system_databases)]

                        # Fiziksel yolları al ve disk sürücülerini çıkar
                        disk_drives = system_db_files["physical_name"].apply(lambda x: x.split(":")[0] + ":").unique()

                        # Eğer tek bir disk varsa, sistem veritabanları aynı disktedir
                        if len(disk_drives) == 1:
                            rows.append({"control_column_name": "SistemDosyalari", "status": 1})
                        else:
                            rows.append({
                                "control_column_name": "SistemDosyalari",
                                "status": 0
                            })
                    except Exception as e:
                        self.log_error(e, "SistemDosyalari")

                elif sheet_name == "CPUInfo":
                    try:
                        # CPU bilgilerini al
                        logical_cpu = sheet_df["Logical CPU"].iloc[0]
                        physical_cpu = sheet_df["Physical CPU"].iloc[0]
                        socket_count = sheet_df["Socket Count"].iloc[0]
                        # Core/Socket oranını hesapla
                        core_socket_rate = logical_cpu / socket_count
                        core_socket_special = [[2,16],[2,8],[2,32],[1,4],[2,24],[2,12],[2,24],[1,8],[2,8],[4,16]]
                        if [socket_count,logical_cpu] in core_socket_special:
                            rows.append({
                                "control_column_name": "CPU",
                                "status": 1
                            })
                        else:
                            # Kontroller
                            if (core_socket_rate != 8 and logical_cpu > 8) or (logical_cpu <= 4 and socket_count > 1):
                                rows.append({
                                    "control_column_name": "CPU",
                                    "status": 1
                                })
                            else:
                                rows.append({
                                    "control_column_name": "CPU",
                                    "status": 0
                                })
                    except Exception as e:
                        self.log_error(e, "CPU")

                elif sheet_name == "ServiceAccount":
                    try:
                        # Service account sütununu kontrol et
                        if sheet_df["service_account"].isna().all():
                            rows.append({
                                "control_column_name": "ServiceAccount",
                                "status": 0
                            })
                        else:
                            rows.append({
                                "control_column_name": "ServiceAccount",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "ServiceAccount")

                elif sheet_name == "SaAccount":
                    try:
                        # SA hesabını kontrol et
                        sa_name = sheet_df["Name"].iloc[0]  # İlk satırdaki Name değeri
                        sa_status = sheet_df["Status"].iloc[0]  # İlk satırdaki Status değeri

                        # Kontrol mekanizması
                        if sa_name.lower() == "sa":
                            rows.append({
                                "control_column_name": "SaAccount",
                                "status": 0
                            })
                        elif sa_status.lower() == "enabled":
                            rows.append({
                                "control_column_name": "SaAccount",
                                "status": 0
                            })
                        else:
                            rows.append({
                                "control_column_name": "SaAccount",
                                "status": 1
                            })
                    except Exception as e:
                        self.log_error(e, "SaAccount")

                elif sheet_name == "Xpcmdshell":
                    try:
                        if str(version).startswith("12"):
                            rows.append({
                                "control_column_name": "XpCmdShell",
                                "status": 3
                            })
                        else:
                            # xp_cmdshell durumunu kontrol et
                            xp_cmdshell_status = sheet_df["xp_cmdshell_configuredvalue"].iloc[0]  # İlk satırdaki değer
                            # Kontrol mekanizması
                            if xp_cmdshell_status == 1:
                                rows.append({
                                    "control_column_name": "XpCmdShell",
                                    "status": 0
                                })
                            else:
                                rows.append({
                                    "control_column_name": "XpCmdShell",
                                    "status": 1
                                })
                    except Exception as e:
                        self.log_error(e, "XpCmdShell")

                elif sheet_name == "BadIndex":
                    try:
                        # Eğer tablo boşsa mesaj döndür
                        if sheet_df.empty:
                            rows.append({
                                "control_column_name": "BadIndex",
                                "status": 1
                            })
                        else:
                            # Database bazlı gruplama (TableName üzerinden)
                            db_index_counts = sheet_df.groupby("TableName").size()

                            # Fail ve Success durumlarını belirle
                            fail_databases = db_index_counts[db_index_counts > 10]
                            success_databases = db_index_counts[db_index_counts <= 10]

                            if not fail_databases.empty:
                                fail_indexes = sheet_df[sheet_df["TableName"].isin(fail_databases.index)][
                                    ["TableName", "IndexName"]]
                                fail_indexes_list = fail_indexes.to_dict(
                                    'records')  # DataFrame'i dictionary listesine çevir

                                rows.append({
                                    "control_column_name": "BadIndex",
                                    "status": 0
                                })
                            else:
                                rows.append({
                                    "control_column_name": "BadIndex",
                                    "status": 1
                                })
                    except Exception as e:
                        self.log_error(e, "BadIndex")

                elif sheet_name == "DatabaseAutoShrinkInfo":
                    try:
                        auto_shrink_on = sheet_df[sheet_df['is_auto_shrink_on'] == 1]
                        if auto_shrink_on.empty:
                            status = 1  # Başarılı
                        else:
                            status = 0  # Başarısız
                        # Durumu rows listesine ekle
                        rows.append({"control_column_name": "AutoShrink", "status": status})
                    except Exception as e:
                        self.log_error(e, "AutoShrink")


            # Create DataFrame from rows
            summary_df = pd.DataFrame(rows, columns=["control_column_name", "status"])
            return summary_df
        except Exception as e:
            self.log_error(e,"generate_dataframe_from_excel")
            QMessageBox.critical(self, "Error", f"Export to Excel failed:\n{e}")
            return


    def summarize_dataframe(self,grouped_df):
        try:
            summary = {}
            for group in grouped_df["Group"].unique():
                group_df = grouped_df[grouped_df["Group"] == group]
                success_count = (group_df["Status"] == "Successful").sum()
                failed_count = (group_df["Status"] == "Failed").sum()
                warning_count = (group_df["Status"]== "Warning").sum()
                summary[group] = {"Successful": success_count, "Warning": warning_count, "Failed": failed_count}
            return summary
        except Exception as e:
            self.log_error(e,"summarize_dataframe")
            return

    def group_dataframe(self,summary_df):
        try:
            """
            Group the DataFrame into categories for better organization in the PDF report.
            """
            group_mapping = {
                "SQL Server Configuration": ["MaxMemory", "MinMemory","SQLFiles","TempDB","LogFiles","SQLConfiguration","CompressionBackup","Storage","ServerConfig","SistemDosyalari"],
                "VM Configuration": ["HighAvailability", "SPCU", "SQLServerVersion", "OSPerformance", "Antivirus",
                                     "Local Security", "IOPerformance", "AlwaysOn"],
                "Performance": ["Deadlock","JobHistory","PowerShellOutput","VLFCount"],
                "Security": ["HighPriviligeLogin","EmptyPasswordLogins","SamePasswordLogins","PolicyNotCheckedLogins","ServerLogins","DisableLogins","CPU","ServiceAccount","ServiceAccountPermission","SaAccount","BUILTINAdministratorsGroup","OrphanUser","ServerAuthenticationMode","ComplexPassword","SameSQLUsernameAsPassword","Xpcmdshell"],
                "Query Performance": ["UpdateStats","ReIndex","LeftoverFakeIndex","ClusteredIndex","MissingIndex","BadIndex","SQLServerBrowserService"],
                "Database Config": [
                    "DatabaseSize",
                    "BackupManagement",
                    "CheckDB",
                    "PageVerify",
                    "CompatibilityLevel",
                    "AutoShrink",
                    "AutoClose",
                    "RecoveryModel",
                    "BackupStats"

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
        except Exception as e:
            self.log_error(e,"group_dataframe")
            return

    def create_gauge_chart(self, value, category, save_path):
        try:
            if value > 65:
                color = "green"
            elif 35 <= value <= 65:
                color = "orange"
            else:
                color = "red"

            # Kare seklinde bir figsize belirle
            fig, ax = plt.subplots(figsize=(9, 6))  # Oranları esit tutmak için 6x6 kullandık
            ax.axis("off")

            # Gauge
            angles = np.linspace(0, np.pi, 300)

            radius = 1
            ax.plot(
                radius * np.cos(angles),
                radius * np.sin(angles),
                color="lightgrey",
                linewidth=35,
                solid_capstyle="round",
            )
            ax.plot(
                radius * np.cos(angles[: int(len(angles) * value / 100)]),
                radius * np.sin(angles[: int(len(angles) * value / 100)]),
                color=color,
                linewidth=35,
                solid_capstyle="round",
            )

            # Add category label
            ax.text(0, -0.2, category, fontsize=30, ha="center", color="black", fontweight="bold")
            ax.text(0, -0.4, f"{value:.2f}%", fontsize=34, ha="center", color="black", fontweight="bold")

            # Grafigi kaydet
            plt.savefig(save_path, dpi=150, bbox_inches="tight", pad_inches=0.5)
            plt.close()
        except Exception as e:
            self.log_error(e,"create_gauge_chart")
            return
    def draw_security_level(self,pdf_canvas, x, y, level, score, stars):
        try:
            # Çizim için ölçüler
            inner_radius = 50
            # Renkler ve açı aralıkları
            # Yarım çember için ikinci katman
            half_outer_radius = 80
            half_segments = [
                (HexColor("#F44336"), 0, 180),  # Yesil
                (HexColor("#FF9800"), 0, 135),  # Sarı
                (HexColor("#FFEB3B"), 0, 90),  # Turuncu
                (HexColor("#4CAF50"), 0, 45 ) # Kırmızı
            ]

            for color, start_angle, end_angle in half_segments:
                pdf_canvas.setFillColor(color)
                pdf_canvas.wedge(
                    x + half_outer_radius, y + half_outer_radius,  # Sol alt
                    x - half_outer_radius, y - half_outer_radius,  # Sag üst
                    start_angle, end_angle, stroke=0, fill=1
                )

                # Determine the angle for the indicator based on the level color

            # İç dairenin rengini puana göre belirle
            if score <= 25:
                inner_color = HexColor("#F44336") # Kırmızı
                indicator_angle = 157.5
            elif 25 < score <= 50:
                inner_color = HexColor("#FF9800") # Turuncu
                indicator_angle = 112.5
            elif 50 < score <= 75:
                inner_color = HexColor("#FFEB3B")  # Sarı
                indicator_angle = 67.5
            elif 75 < score <=100:
                inner_color = HexColor("#4CAF50")  # Yesil
                indicator_angle = 22.5
                # Convert the angle to radians
            else:
                inner_color = HexColor("#00000")
                indicator_angle = 90
            indicator_angle = 180-(180*score/100)
            angle_rad = radians(indicator_angle)

            # Calculate the endpoint of the indicator line
            line_length = half_outer_radius -10  # Length of the indicator line
            end_x = x + line_length * cos(angle_rad)
            end_y = y + line_length * sin(angle_rad)

            # Draw the indicator line
            pdf_canvas.setStrokeColor(HexColor("#000000"))  # Black line
            pdf_canvas.setLineWidth(2)
            pdf_canvas.line(x, y, end_x, end_y)
            # İç çember
            pdf_canvas.setFillColor(inner_color)
            pdf_canvas.circle(x, y, inner_radius, stroke=1, fill=1)

            # Orta metinler (Level ve Score)
            pdf_canvas.setFont("Helvetica-Bold", 14)
            pdf_canvas.setFillColor(HexColor("#000000"))
            pdf_canvas.drawCentredString(x, y + 10, level)  # Level text
            pdf_canvas.setFont("Helvetica", 12)
            pdf_canvas.drawCentredString(x, y - 5, f"Score: {score}")  # Score text

            # Yıldızlar
            star_x = x - (stars * 7)  # Yıldızları ortalamak için baslangıç noktasi
            pdf_canvas.setFillColor(HexColor("#00000"))  # Gold color
            for _ in range(stars):
                pdf_canvas.drawString(star_x, y - 20, "★")
                star_x += 14
        except Exception as e:
            self.log_error(e,"draw_security_level")
            return
    def add_watermark_to_pdf(self,input_pdf, output_pdf):
        try:
            base_path = os.path.dirname(os.path.abspath(__file__))

            # Watermark dosyasının yolu (main.py ile aynı dizinde)
            watermark_image = os.path.join(base_path, "dplogo1.png")
            # Mevcut PDF'yi oku
            existing_pdf = PdfReader(input_pdf)
            total_pages = len(existing_pdf.pages)  # Mevcut PDF'deki toplam sayfa sayısını al

            # Filigranı küçült ve soluk hale getir
            img = Image.open(watermark_image).convert("RGBA")
            img = img.resize((int(A4[0] / 3), int(A4[1] / 3)))  # Sayfanın 1/3 boyutunda olacak

            # Şeffaflık ekle
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: p * 0.2)  # %20 opaklık ekleyerek soluk hale getir
            img.putalpha(alpha)

            # Yeni PNG olarak kaydet (geçici dosya)
            temp_watermark = "temp_watermark1.png"
            img.save(temp_watermark)

            # Yeni PDF dosyası oluştur
            output = PdfWriter()

            for page_num in range(total_pages):
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=A4)
                width, height = A4

                # Filigranı merkeze ekle (küçültülmüş)
                watermark = ImageReader(temp_watermark)
                can.drawImage(watermark, width / 3, height / 3, width=A4[0] / 3, height=A4[1] / 3, mask='auto')

                # Sayfayı kaydet
                can.save()

                # Filigranı yeni bir PDF sayfası olarak ekle
                packet.seek(0)
                watermark_pdf = PdfReader(packet)
                watermark_page = watermark_pdf.pages[0]

                # Mevcut PDF sayfasını al ve filigranı üzerine yerleştir
                existing_page = existing_pdf.pages[page_num]
                existing_page.merge_page(watermark_page)

                # Güncellenmiş sayfayı çıktı PDF'sine ekle
                output.add_page(existing_page)

            # Yeni PDF'yi kaydet
            with open(output_pdf, "wb") as outputStream:
                output.write(outputStream)
        except Exception as e:
            self.log_error(e,"add_watermark_to_pdf")
            return
    def generate_pdf_report(self,excel_path, output_dir,username,rundate):

        try:
            main_ui = self.parent().parent()
            main_ui.progress_overlay.show()
            main_ui.adjustSize()
            main_ui.progress_bar.setValue(0)
            QApplication.processEvents()
            time.sleep(0.5)
            QApplication.processEvents()  # UI güncellenmesini sağla
            font_path = self.get_resource_path("arialbd.ttf")
            pdfmetrics.registerFont(TTFont('ArialBlack', font_path))
            summary_df = self.generate_dataframe_from_excel(excel_path)
            grouped_df = self.group_dataframe(summary_df)
            summary = self.summarize_dataframe(grouped_df)
            allScore = self.score_table(summary_df)
        except Exception as e:
            self.log_error(e,'font_path')
        try:
            with pd.ExcelWriter(excel_path, mode='a', engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='DataFrame', index=False)
        except Exception as e:
            self.log_error(e,'excel_writer')
            QMessageBox.critical('Error','An error occurred, the Excel file could not be read.')

        try:
            main_ui.progress_bar.setValue(10)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)
            # Username ve rundate bilgilerini al

            pdf_file_name = f"{username}_{rundate}.pdf"
            pdf_path = os.path.join(output_dir, pdf_file_name)

            # PDF için canvas olustur
            pdf = canvas.Canvas(pdf_path, pagesize=A4)
            width, height = A4

            image_path = self.get_resource_path("dpoint1.png")  # Dosyanın yolu
            img_width = 50  # Resmin genişliği (isteğe bağlı olarak değiştirilebilir)
            img_height = 50  # Resmin yüksekliği
            img_x = width / 2 - img_width   # Ortalamak için sol tarafa kaydır
            img_y = height - 50  # Resmin konumu

            pdf.drawImage(ImageReader(image_path), img_x, img_y, width=img_width, height=img_height)

            main_ui.progress_bar.setValue(20)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)

            pdf.setTitle("CREATED BY DAPLAIT INFORMATION SYSTEM")
            pdf.setTitle("SQL SERVER HEALTH CHECK")

            # Kenar boslukları
            margin_left = 30
            margin_right = 30

            # Baslık Arka Planı
            pdf.setFillColor(HexColor("#003366"))  # Lacivert arka plan
            pdf.rect(0, height - 90, width, 50, fill=1)

            # Baslık Yazısı
            pdf.setFillColor(HexColor("#FFFFFF"))  # Beyaz yazı rengi
            pdf.setFont("Helvetica-Bold", 5)  # Daha profesyonel bir yazı tipi
            pdf.drawCentredString(width / 2, height - 55, "CREATED BY DAPLAIT INFORMATION SYSTEM")
            pdf.setFont("Helvetica-Bold", 18)  # Daha profesyonel bir yazı tipi
            pdf.drawCentredString(width / 2, height - 70, "SQL SERVER HEALTH CHECK")
            # Alt Baslık (Tarih veya Kullanıcı Adı gibi ek bilgiler)
            pdf.setFont("Helvetica", 12)
            pdf.setFillColor(HexColor("#FFFFFF"))
            pdf.drawCentredString(width / 2, height - 85, f"Prepared by: {username} | Date: {rundate}")
            main_ui.progress_bar.setValue(30)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)
            excel_data = pd.ExcelFile(excel_path)
            ServerInfo = pd.read_excel(excel_data, sheet_name="ServerInfo")

            serverinfodata = [
                ['SERVER', 'EDITION', 'VERSION', 'COLLATION', 'CPU', 'RAM (MB)', 'CLUSTER', 'HA', 'VM SERVER', 'OS',
                 'Cumulative Update', 'SQL Version']]  # Başlık
            serverinfodata.extend(ServerInfo.values.tolist())
            df = pd.DataFrame(serverinfodata[1:], columns=serverinfodata[0])
            transposed_df = df.transpose()

            # DataFrame'i table verisine dönüştür
            transposed_data = [['SERVER INFO']]  # Başlık
            transposed_data.extend(transposed_df.itertuples(index=True, name=None))

            tableserverinfo = Table(transposed_data, colWidths=[width * 0.15, width * 0.2])
            styleserverinfo = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), HexColor("#68abe8")),  # Başlık arka planı
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Başlık metin rengi
                ('FONTNAME', (0, 0), (-1, 0), 'ArialBlack'),  # Başlık fontu
                ('FONTSIZE', (0, 0), (-1, 0), 8),  # Başlık font boyutu
                ('FONTSIZE', (0, 1), (-1, -1), 6),  # Gövde font boyutu
                ('GRID', (0, 0), (-1, -1), 1, colors.white),  # Izgara çizgileri
                ('TOPPADDING', (0, 0), (-1, -1), 0),  # Üst dolgu
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),  # Alt dolgu
                ('LEFTPADDING', (0, 0), (-1, -1), 1),  # Sol dolgu
                ('RIGHTPADDING', (0, 0), (-1, -1), 1),  # Sağ dolgu
                ('ALIGN', (0, 0), (-1, 0), 'LEFT'),  # Sadece başlık satırını merkeze hizala
                ('ALIGN', (0, 1), (-1, -1), 'LEFT'),  # Tüm hücreler için hizalama
            ])

            x_position = 30  # 200, göstergenin genisligi
            y_position = height - 260
            tableserverinfo.setStyle(styleserverinfo)
            tableserverinfo.wrapOn(pdf, width - 60, height - 25)  # Sağdan ve soldan boşluk bırakarak wrap
            tableserverinfo.drawOn(pdf, x_position, y_position)
            main_ui.progress_bar.setValue(40)  # %20 tamamlandı
            QApplication.processEvents()
            if not summary_df.empty:
                count_true = 100-allScore

                if 0 < count_true <= 25:
                    level = 'WORSE'
                    stars = 1
                elif 25 < count_true <= 50:
                    level = 'MEDIUM'
                    stars = 2
                elif 50 < count_true <= 75:
                    level = 'GOOD'
                    stars = 3
                elif 75 <count_true <= 100:
                    level = 'VERY GOOD'
                    stars = 4
                else:
                    level = 'NONE'
                    stars = 0
                # Güvenlik seviyesi göstergesinin konumunu ortalamak için düzenleme
                x_position = width / 1.25  # 200, göstergenin genisligi
                y_position = height - 200  # Y konumu, üstten olan mesafe
                self.draw_security_level(pdf, x_position, y_position, level=level, score=count_true, stars=stars)

                # Baslık Yazısı
                pdf.setFillColor(HexColor("#00000"))  # Beyaz yazı rengi
                pdf.setFont("Helvetica-Bold", 10)  # Daha profesyonel bir yazı tipi
                pdf.drawCentredString(width / 1.25, height - 110, "HEALTH CHECK LEVEL")

            main_ui.progress_bar.setValue(50)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)

            #security level yazisi
            # Baslık Arka Planı
            pdf.setFillColor(HexColor("#9bb5e8"))
            pdf.rect(0, height - 300, width , 30, fill=1, stroke=0)
            pdf.setFillColor(HexColor("#FFFFFF"))  # Beyaz yazı rengi
            pdf.setFont("Helvetica-Bold", 12)  # Daha profesyonel bir yazı tipi
            pdf.drawCentredString(width / 2, height - 290, "HEALTH CHECK LEVEL")

            # Excel dosyasını oku
            sheets = pd.ExcelFile(excel_path)
            y_position = height - 280  # İlk tablo konumu

            # Reportlab Paragraph stilini al
            styles = getSampleStyleSheet()
            style = styles["BodyText"]
            main_ui.progress_bar.setValue(60)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)
            #smcid ye göre iki sheet'i birlestirip dataframe'e atama
            #score_table = sheets.parse('ScoreMasterDetail')
            #description_table = sheets.parse('ScoreMasterConfig')

            # Kolon isimlerini kucük harfe dönüstür
            #score_table.columns = [col.lower() for col in score_table.columns]
            #description_table.columns = [col.lower() for col in description_table.columns]

            # SmcID kolonunu kullanarak birlestirme
            #merged_table = pd.merge(description_table, score_table, on="smcid", how="left")
            #merged_table['description'] = merged_table['description'].apply(self.clean_column)



            # Baslangıç y-konumu ve satırda kaç görsel/tablo oldugunu takip eden degiskenler
            main_ui.progress_bar.setValue(70)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)
            x_position = 80
            images_per_row = 3
            image_count = 0

            for category, metrics in summary.items():
                # Add gauge chart

                gauge_path = f"{output_dir}/{category}_gauge.png"
                self.create_gauge_chart(round((metrics["Successful"] + (metrics["Warning"]*0.4)) /(metrics["Warning"] + metrics["Successful"] + metrics["Failed"] ) * 100, 1), category,
                                        gauge_path)

                # Görseli PDF'e yerlestir
                pdf.drawImage(gauge_path, x_position, y_position - 150, width=150, height=100)
                # Verileri düzenleme
                data = [

                    ["Success", "Failed", "Warning"],  # İkinci satır başlıklar
                    [metrics["Successful"], metrics["Failed"], metrics["Warning"]],  # Üçüncü satır veriler
                ]
                # Sütun genişliklerini ayarla (başlık satırı için tek genişlik)
                # Tabloyu oluştur
                table = Table(data, colWidths=[50, 50, 50])
                table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), HexColor("#1f3242")),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("FONTNAME", (0, 0), (-1, 0), "ArialBlack"),
                            ("FONTSIZE", (0, 0), (-1, -1), 10),
                            ("GRID", (0, 0), (-1, -1), 1, colors.black),
                        ]
                    )
                )

                table.wrapOn(pdf, width, height)
                table.drawOn(pdf, x_position, y_position - 200)  # Görselin altına tabloyu yerlestir

                # X pozisyonunu kaydır ve sayaç artır
                x_position += 160  # Görseller ve tablolar arası yatay bosluk için
                image_count += 1

                # Görseller ve tablolar satırı dolunca bir alt satıra geç
                if image_count % images_per_row == 0:
                    x_position = 80  # Sola sıfırla
                    y_position -= 250  # Bir satır asagı in (görsel + tablo yüksekligi)

                # Sayfa sonuna ulasırsa yeni sayfa aç ve bastan basla
                if y_position < 150:
                    pdf.showPage()
                    x_position = 80
                    y_position = height - 100

            main_ui.progress_bar.setValue(80)  # %20 tamamlandı
            QApplication.processEvents()
            time.sleep(0.5)
            y_position = height - 100
            pdf.setFillColor(HexColor("#9bb5e8"))
            pdf.rect(0, y_position , width, 30, fill=1 , stroke=0)
            # Baslık Yazısı
            pdf.setFillColor(HexColor("#FFFFFF"))  # Beyaz yazı rengi
            pdf.setFont("Helvetica-Bold", 12)  # Daha profesyonel bir yazı tipi
            pdf.drawCentredString(width / 2, y_position+12 , "HEALTH CHECK RESULT")
            y_position -= 2

            main_ui.progress_bar.setValue(90)  # %20 tamamlandı
            QApplication.processEvents()
            # NaN değerlerini doldur (isteğe bağlı olarak "N/A" ile doldurabilirsiniz)

            data2 = grouped_df.iterrows()

            # Başlıkları ekleyin (ilk satır)
            data2 = [["Group", "Description", "Current Status"]]

            last_group = None  # Tekrarlayan 'Group' değerlerini kontrol etmek için
            first_group_seen = False
            # Birleştirilmiş DataFrame'i döngü ile işleyin
            unique_groups = grouped_df["Group"].unique()

            for group in unique_groups:
                group_df = grouped_df[grouped_df["Group"] == group]
                for _, row in grouped_df.iterrows():
                    # 'Group' sütununda tekrar eden değerleri boş bırak
                    group_value = row["Group"] if row["Group"] != last_group else ""
                    # Satırdaki değerleri alın ve data2'ye ekleyin
                    if row["Group"] != "SQL Server Configuration":
                        description_value = row["Description"] #blurlanacak yer
                    else:
                        description_value = row["Description"]
                    row_data = [group_value, description_value, row["Status"]] + row[3:].tolist()
                    data2.append(row_data)
                    last_group = row["Group"]  # Son görülen grubu güncelle

                margin_left = 30
                table_width = width - (2 * margin_left)
                col_count = len(data2[0])
                col_width = table_width / col_count

                table2 = Table(data2, colWidths=[col_width] * col_count)

                table_styles = [
                    ('BACKGROUND', (0, 0), (-1, 0), HexColor("#68abe8")),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.white),
                ]

                for idx, row in enumerate(data2[1:], start=1):
                    for col_idx, value in enumerate(row[2:], start=2):
                        if value == "Successful":
                            table_styles.append(('BACKGROUND', (col_idx, idx), (col_idx, idx), HexColor("#52bf58")))
                            table_styles.append(('TEXTCOLOR', (col_idx, idx), (col_idx, idx), colors.whitesmoke))
                        elif value == "Failed":
                            table_styles.append(('BACKGROUND', (col_idx, idx), (col_idx, idx), HexColor("#ed283c")))
                            table_styles.append(('TEXTCOLOR', (col_idx, idx), (col_idx, idx), colors.whitesmoke))
                        elif value == "Warning":
                            table_styles.append(('BACKGROUND', (col_idx, idx), (col_idx, idx), HexColor("#FFA500")))
                            table_styles.append(('TEXTCOLOR', (col_idx, idx), (col_idx, idx), colors.whitesmoke))

                table2.setStyle(TableStyle(table_styles))
                table_max_width = width - (margin_left * 2)
                y_position = height - 100  # Sayfa başında boşluk

                # Tablonun sayfaya sığmasını kontrol et
                table_splits2 = table2.split(table_max_width, y_position)
                for split_table in table_splits2:
                    split_width, split_height = split_table.wrap(table_max_width, y_position)
                    if y_position - split_height < 50:
                        pdf.showPage()
                        y_position = height - 50
                    split_table.drawOn(pdf, margin_left, y_position - split_height)
                    y_position -= split_height + 20

                # Her grup için yeni sayfa
                pdf.showPage()
            main_ui.progress_bar.setValue(100)  # %20 tamamlandı
            QApplication.processEvents()
            #pdf.drawImage(image_path, img_x, img_y, width=img_width, height=img_height)
            main_ui.progress_overlay.hide()

            # PDF'yi kaydet
            pdf.save()
            self.add_watermark_to_pdf(pdf_path, pdf_path)

            return pdf_path
        except Exception as e:
            self.log_error(e, 'generate_pdf_report')
            return

    def extract_table_names(self, sql_script):
        try:
            """
            SQL script'inden tüm SELECT ifadelerini ve ilgili tablo adlarını çıkarır.
            Örnek: SELECT * FROM #SecHC_ComplexPassword → SecHC_ComplexPassword
            """
            table_names = re.findall(r"SELECT\s+\*.*?FROM\s+#([a-zA-Z0-9_]+)", sql_script, re.IGNORECASE)

            return table_names
        except Exception as e:
            self.log_error(e, 'extract_table_names')
            return

    def run_scripts_and_export(self):

        try:
            self.main_ui.server_info_frame.setVisible(False)
            username = self.server_name.text()
            valid_username = username.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace(
                "?", "_").replace("\"", "_").replace("<", "_").replace(">", "_").replace("|", "_")

            rundate = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"{valid_username}_{rundate}.xlsx"

            # Programın çalıştığı yerdeki 'csv' klasörünün yolu
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            csv_dir = os.path.join(base_path, "csv")
            os.makedirs(csv_dir, exist_ok=True)

            # Kullanıcıdan kaydetme konumunu seçmesini iste
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Report", os.path.join(csv_dir, file_name),
                                                       "Excel Files (*.xlsx);;All Files (*)", options=options)
            if not file_path:
                return  # Kullanıcı "Cancel" dediyse işlemi iptal et

            output_path = file_path  # Kullanıcının seçtiği dosya yolu
            used_sheet_names = set()


            self.main_ui.checklist_page.set_output_path(output_path)
            self.main_ui.stack.setCurrentWidget(self.main_ui.checklist_page)

            encoded_scripts = [SECURITY_SCRIPT54_BASE64,SECURITY_SCRIPT_BASE64, SECURITY_SCRIPT2_BASE64, SECURITY_SCRIPT5_BASE64,
                               SECURITY_SCRIPT6_BASE64, SECURITY_SCRIPT7_BASE64, SECURITY_SCRIPT10_BASE64,SECURITY_SCRIPT11_BASE64, SECURITY_SCRIPT8_BASE64]

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                sheet_added = False
                for encoded_script in encoded_scripts:
                    cursor = self.connection.cursor()
                    decoded_bytes = base64.b64decode(encoded_script)
                    detected = chardet.detect(decoded_bytes)
                    encoding = detected.get('encoding', 'utf-8')
                    decoded_script = decoded_bytes.decode(encoding, errors='ignore')
                    statements = decoded_script.split("GO")
                    select_statements = re.findall(r"(SELECT\s+\*.*?FROM\s+#\w+;)", decoded_script, re.IGNORECASE)
                    table_names = self.extract_table_names(decoded_script)

                    for statement in statements:
                        statement = statement.strip()
                        if statement:
                            try:
                                cursor.execute(statement)
                            except Exception as e:
                                self.log_error(e, "SQL Execution Error")

                    for i, select in enumerate(select_statements):
                        table_name = table_names[i] if i < len(table_names) else f"Sheet{i + 1}"

                        if table_name in used_sheet_names:
                            count = 1
                            while f"{table_name}_{count}" in used_sheet_names:
                                count += 1
                            table_name = f"{table_name}_{count}"
                        used_sheet_names.add(table_name)
                        selected = re.sub(r"SELECT \* FROM #*", "", select)
                        try:
                            cursor.execute(select)

                            QApplication.processEvents()
                            try:
                                if cursor.description:
                                    columns = [desc[0] for desc in cursor.description]
                                    rows = cursor.fetchall()
                                    data = pd.DataFrame.from_records(rows, columns=columns)
                                    if data.empty and columns:
                                        data = pd.DataFrame(
                                            columns=columns)  # Boş DataFrame, sadece sütun başlıkları ile
                                    if not data.empty or columns:
                                        data.to_excel(writer, sheet_name=table_name[:31], index=False, header=True)
                                        sheet_added = True

                                    self.main_ui.checklist_page.update_list(selected, True)

                            except Exception as e:
                                self.log_error(e, "cursor_error")

                        except Exception as e:
                            self.checklist_window.update_list(selected, False)
                            self.log_error(e, "windows_group_error")
                            continue
                        time.sleep(0.3)
                    if not sheet_added:
                        df_placeholder = pd.DataFrame({"No Data": ["No queries returned results"]})
                        df_placeholder.to_excel(writer, sheet_name="Placeholder", index=False)
            self.main_ui.checklist_page.finalize_progress()
            self.main_ui.checklist_page.setVisible(False)
            QMessageBox.information(self, "SUCCESS",
                                    f"The report has been saved in the {output_path}\n\n You can access the location where the file is saved by clicking the \n->Open Output Folder\n  button on the Execution Checklist page.")

            if getattr(sys, 'frozen', False):  # Eger derlenmiş bir .exe dosyası ise
                base_path = os.path.dirname(sys.executable)
            else:  # Eger bir Python dosyası olarak çalıştırılıyorsa
                base_path = os.path.dirname(os.path.abspath(__file__))

            output_dir = os.path.join(base_path, "report")
            os.makedirs(output_dir, exist_ok=True)  # Klasör yoksa olustur


            pdf_path = self.generate_pdf_report(output_path, output_dir, valid_username, rundate)
            self.parent().parent().stack.setCurrentWidget(self.parent().parent().progress_overlay)

            QMessageBox.information(self, "SUCCESS",
                                    f"The report has been saved in the {output_dir} folder as {pdf_path}.")

        except Exception as e:
            self.log_error(e, "run_script")
            QMessageBox.critical(self, "Error", f"An occurred error:\n{e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Ana UI
    main_window = MainUI()

    # Splash ekran yükle
    base_path = os.path.dirname(os.path.abspath(__file__))
    splash_path = os.path.join(base_path, "dp-splash1.png")

    if os.path.exists(splash_path):
        pixmap = QPixmap(splash_path)
        splash = QSplashScreen(pixmap)
        splash.show()
        QTimer.singleShot(3000, lambda: (splash.close(), main_window.show()))
    else:
        print("Splash image not found:", splash_path)
        main_window.show()

    sys.exit(app.exec_())



