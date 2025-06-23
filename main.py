import sys
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QTabWidget, 
                            QLabel, QFrame)
from PyQt6.QtCore import Qt
from login_window import LoginWindow
from db import create_connection
from widgets import TableWidget

TABLES = [
    ("Оборудование", "ОБОРУДОВАНИЕ"),
    ("Рабочие места", "РАБОЧЕЕ_МЕСТО"),
    ("Сотрудники", "СОТРУДНИК"),
    ("Выдача техники", "ВЫДАЧА_ТЕХНИКИ"),
    ("Обслуживание", "ОБСЛУЖИВАНИЕ"),
]

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Учет техники производства")
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # 1. Заголовок компании (бело-синий)
        title_label = QLabel("ООО «Разработчики Программ»")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #1565c0;
                margin-bottom: 5px;
            }
        """)
        main_layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignLeft)

        # 2. Подзаголовок
        subtitle_label = QLabel("Учет техники производства")
        subtitle_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #555;
                font-style: italic;
                margin-bottom: 10px;
            }
        """)
        main_layout.addWidget(subtitle_label, alignment=Qt.AlignmentFlag.AlignLeft)

        # 3. Табы с таблицами (бело-синий стиль)
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.TabPosition.North)
        self.tabs.setMovable(True)
        self.tabs.setStyleSheet("""
            QTabWidget::pane { 
                border: 1px solid #1565c0; 
                border-radius: 8px; 
                background: #fff;
            }
            QTabBar::tab {
                background: #e3f2fd;
                color: #1565c0;
                border: 1px solid #90caf9;
                border-bottom: none;
                min-width: 110px;
                min-height: 26px;
                font-size: 13px;
                padding: 6px;
                margin-right: 2px;
                border-radius: 0;
                font-weight: 500;
            }
            QTabBar::tab:first {
                border-top-left-radius: 8px;
                border-top-right-radius: 0;
            }
            QTabBar::tab:last {
                border-top-left-radius: 0;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background: #1565c0;
                color: #fff;
                font-weight: bold;
            }
            QTabBar::tab:hover {
                background: #1976d2;
                color: #fff;
            }
            QWidget {
                background: #fff;
                color: #333;
            }
        """)
        
        for label, table in TABLES:
            self.tabs.addTab(TableWidget(table), label)

        main_layout.addWidget(self.tabs, stretch=1)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    create_connection()

    login = LoginWindow()
    if login.exec():
        window = MainWindow()
        window.showMaximized()
        sys.exit(app.exec())
    else:
        sys.exit()