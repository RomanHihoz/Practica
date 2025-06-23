import sys
import sqlite3
import os
from PyQt6.QtWidgets import (
    QApplication, QDialog, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QCheckBox, QMessageBox
)
from PyQt6.QtCore import Qt

DB_PATH = os.path.join(os.path.dirname(__file__), "tech.db")

class LoginWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Вход — Учет техники")
        self.setFixedSize(420, 360)
        self.setStyleSheet("""
            QWidget {
                background-color: #fff;
                color: #333;
                font-family: Segoe UI, sans-serif;
            }
            QLineEdit {
                background-color: #e3f2fd;
                color: #1565c0;
                border: 1px solid #90caf9;
                padding: 6px;
                border-radius: 4px;
            }
            QPushButton {
                background-color: #1565c0;
                color: #fff;
                border: none;
                padding: 6px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
            QCheckBox {
                color: #1565c0;
                padding-left: 2px;
            }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Заголовок (бело-синий)
        title = QLabel("ООО «Разработчики Программ»")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #1565c0;
                margin-bottom: 2px;
            }
        """)
        layout.addWidget(title)

        # Подзаголовок
        subtitle = QLabel("Учет техники производства")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #1976d2;
                font-style: italic;
                margin-bottom: 15px;
            }
        """)
        layout.addWidget(subtitle)

        # Надпись "Вход"
        login_title = QLabel("Вход")
        login_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        login_title.setStyleSheet("font-size: 14pt; margin-bottom: 12px; color: #1565c0;")
        layout.addWidget(login_title)

        # Поля
        self.username = QLineEdit()
        self.username.setPlaceholderText("Логин")
        layout.addWidget(self.username)

        self.password = QLineEdit()
        self.password.setPlaceholderText("Пароль")
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password)

        self.show_pass = QCheckBox("Показать пароль")
        self.show_pass.stateChanged.connect(self.toggle_password)
        layout.addWidget(self.show_pass)

        # Кнопка "Войти"
        self.login_btn = QPushButton("Войти")
        self.login_btn.setFixedWidth(180)
        self.login_btn.clicked.connect(self.try_login)
        login_btn_layout = QHBoxLayout()
        login_btn_layout.addStretch()
        login_btn_layout.addWidget(self.login_btn)
        login_btn_layout.addStretch()
        layout.addLayout(login_btn_layout)

        # Кнопка "Регистрация"
        self.register_btn = QPushButton("Регистрация")
        self.register_btn.setFixedWidth(180)
        self.register_btn.clicked.connect(self.register_user)
        register_btn_layout = QHBoxLayout()
        register_btn_layout.addStretch()
        register_btn_layout.addWidget(self.register_btn)
        register_btn_layout.addStretch()
        layout.addLayout(register_btn_layout)

    def toggle_password(self, state):
        self.password.setEchoMode(QLineEdit.EchoMode.Normal if state else QLineEdit.EchoMode.Password)

    def try_login(self):
        login = self.username.text().strip()
        password = self.password.text().strip()

        if not login or not password:
            QMessageBox.warning(self, "Ошибка", "Введите логин и пароль.")
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT пароль FROM ПОЛЬЗОВАТЕЛИ WHERE логин = ?", (login,))
            row = cursor.fetchone()
            conn.close()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка БД", str(e))
            return

        if row:
            if row[0] == password:
                QMessageBox.information(self, "Добро пожаловать", f"{login}, вход выполнен.")
                self.accept()
            else:
                QMessageBox.warning(self, "Ошибка", "Неверный пароль.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пользователь не найден.")

    def register_user(self):
        login = self.username.text().strip()
        password = self.password.text().strip()

        if not login or not password:
            QMessageBox.warning(self, "Ошибка", "Введите логин и пароль.")
            return

        if len(login) < 6 or len(password) < 6:
            QMessageBox.warning(self, "Ошибка", "Логин и пароль должны содержать минимум 6 символов.")
            return

        if login.isdigit() or password.isdigit():
            QMessageBox.warning(self, "Ошибка", "Логин и пароль не могут состоять только из цифр.")
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO ПОЛЬЗОВАТЕЛИ (логин, пароль) VALUES (?, ?)", (login, password))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Регистрация", "Пользователь зарегистрирован.")
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Ошибка", "Логин уже существует.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка БД", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LoginWindow()
    if window.exec():
        print("Вход выполнен.")
    else:
        print("Вход отменён.")