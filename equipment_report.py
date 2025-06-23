from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QDateEdit, QPushButton, QFileDialog, QMessageBox
)
from PyQt6.QtGui import (
    QTextDocument, QTextCursor, QTextTableFormat, QTextCharFormat,
    QFont, QTextTableCellFormat, QTextLength, QColor
)
from PyQt6.QtCore import Qt, QDate, QMarginsF
from PyQt6.QtPrintSupport import QPrinter
import os

class EquipmentReportGenerator(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Отчёт по оборудованию")
        self.setMinimumSize(600, 500)
        self.model = model

        self.init_ui()
        self.load_filters()
        self.load_fields()

    def init_ui(self):
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Выберите поля для отчёта:"))
        self.fields_list = QListWidget()
        self.fields_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.fields_list)

        # Фильтры
        layout.addWidget(QLabel("Фильтрация:"))

        self.type_filter = QListWidget()
        self.type_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Тип оборудования:"))
        layout.addWidget(self.type_filter)

        self.status_filter = QListWidget()
        self.status_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Статус:"))
        layout.addWidget(self.status_filter)

        self.manufacturer_filter = QListWidget()
        self.manufacturer_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Производитель:"))
        layout.addWidget(self.manufacturer_filter)

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd.MM.yyyy")
        self.date_to.setDisplayFormat("dd.MM.yyyy")

        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Дата покупки от:"))
        date_layout.addWidget(self.date_from)
        date_layout.addWidget(QLabel("до:"))
        date_layout.addWidget(self.date_to)
        layout.addLayout(date_layout)

        # Кнопки
        button_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Сформировать PDF")
        self.cancel_btn = QPushButton("Отмена")
        button_layout.addWidget(self.generate_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        self.generate_btn.clicked.connect(self.generate_report)
        self.cancel_btn.clicked.connect(self.reject)

    def load_filters(self):
        type_col = status_col = manufacturer_col = purchase_col = None

        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if "тип" in header:
                type_col = col
            elif "статус" in header:
                status_col = col
            elif "производитель" in header:
                manufacturer_col = col
            elif "дата покупки" in header:
                purchase_col = col

        def fill_unique_values(col_index, widget):
            seen = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, col_index))
                if val and val not in seen:
                    widget.addItem(val)
                    seen.add(val)

        if type_col is not None:
            fill_unique_values(type_col, self.type_filter)
        if status_col is not None:
            fill_unique_values(status_col, self.status_filter)
        if manufacturer_col is not None:
            fill_unique_values(manufacturer_col, self.manufacturer_filter)

        # Определим минимальную и максимальную дату
        dates = []
        if purchase_col is not None:
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, purchase_col))
                dt = QDate.fromString(val, "dd.MM.yyyy")
                if dt.isValid():
                    dates.append(dt)
        self.date_from.setDate(min(dates) if dates else QDate(2000, 1, 1))
        self.date_to.setDate(max(dates) if dates else QDate.currentDate())

    def load_fields(self):
        self.fields_list.clear()
        for col in range(self.model.columnCount()):
            header = self.model.headerData(col, Qt.Orientation.Horizontal)
            if not header or str(header).lower() == "id":
                continue
            item = QListWidgetItem(str(header))
            item.setData(Qt.ItemDataRole.UserRole, col)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.fields_list.addItem(item)

    def get_selected_fields(self):
        selected = []
        for i in range(self.fields_list.count()):
            item = self.fields_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                col = item.data(Qt.ItemDataRole.UserRole)
                header = item.text()
                selected.append((col, header))
        return selected

    def generate_report(self):
        selected_fields = self.get_selected_fields()
        if not selected_fields:
            QMessageBox.warning(self, "Нет полей", "Выберите хотя бы одно поле для отчёта.")
            return

        doc = QTextDocument()
        cursor = QTextCursor(doc)

        # Заголовок
        title_format = QTextCharFormat()
        title_font = QFont("Arial", 16)
        title_font.setBold(True)
        title_format.setFont(title_font)
        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText("Отчёт по оборудованию\n\n")

        # Таблица
        table_format = QTextTableFormat()
        table_format.setCellPadding(4)
        table_format.setBorder(1)
        table_format.setBorderBrush(QColor("black"))
        table_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        table_format.setColumnWidthConstraints([
            QTextLength(QTextLength.Type.PercentageLength, 100 / (len(selected_fields) + 1))
            for _ in range(len(selected_fields) + 1)
        ])
        table = cursor.insertTable(1, len(selected_fields) + 1, table_format)

        # Формат заголовков
        header_format = QTextCharFormat()
        header_font = QFont("Arial", 10)
        header_font.setBold(True)
        header_format.setFont(header_font)

        # Формат обычных ячеек
        cell_format = QTextTableCellFormat()
        cell_format.setBorder(0.5)
        cell_format.setBorderBrush(QColor("black"))

        cell_text_format = QTextCharFormat()
        small_font = QFont("Arial", 8)
        cell_text_format.setFont(small_font)

        # Заголовки
        table.cellAt(0, 0).setFormat(cell_format)
        cursor0 = table.cellAt(0, 0).firstCursorPosition()
        cursor0.setCharFormat(header_format)
        cursor0.insertText("#")

        for i, (_, hdr) in enumerate(selected_fields, 1):
            table.cellAt(0, i).setFormat(cell_format)
            cursor_i = table.cellAt(0, i).firstCursorPosition()
            cursor_i.setCharFormat(header_format)
            cursor_i.insertText(hdr)

        # Получаем выбранные значения фильтров
        types = [item.text() for item in self.type_filter.selectedItems()]
        statuses = [item.text() for item in self.status_filter.selectedItems()]
        mans = [item.text() for item in self.manufacturer_filter.selectedItems()]
        d_from = self.date_from.date()
        d_to = self.date_to.date()

        # Индексы колонок
        type_col = status_col = man_col = purchase_col = None
        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if "тип" in header:
                type_col = col
            elif "статус" in header:
                status_col = col
            elif "производитель" in header:
                man_col = col
            elif "дата покупки" in header:
                purchase_col = col

        for row in range(self.model.rowCount()):
            # Фильтрация
            t = self.model.data(self.model.index(row, type_col)) if type_col is not None else ""
            s = self.model.data(self.model.index(row, status_col)) if status_col is not None else ""
            m = self.model.data(self.model.index(row, man_col)) if man_col is not None else ""
            d_str = self.model.data(self.model.index(row, purchase_col)) if purchase_col is not None else ""
            d_obj = QDate.fromString(d_str, "dd.MM.yyyy")

            if types and t not in types:
                continue
            if statuses and s not in statuses:
                continue
            if mans and m not in mans:
                continue
            if d_obj.isValid() and (d_obj < d_from or d_obj > d_to):
                continue

            # Добавляем строку
            table.appendRows(1)
            r_idx = table.rows() - 1

            table.cellAt(r_idx, 0).setFormat(cell_format)
            cursor0 = table.cellAt(r_idx, 0).firstCursorPosition()
            cursor0.setCharFormat(cell_text_format)
            cursor0.insertText(str(r_idx))

            for i, (col, _) in enumerate(selected_fields, 1):
                val = self.model.data(self.model.index(row, col))
                table.cellAt(r_idx, i).setFormat(cell_format)
                cursor_i = table.cellAt(r_idx, i).firstCursorPosition()
                cursor_i.setCharFormat(cell_text_format)
                cursor_i.insertText(str(val) if val else "")

        # Сохранение в PDF
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", os.path.expanduser("~/Отчет_Оборудование.pdf"), "PDF Files (*.pdf)"
        )
        if not file_path:
            return

        printer.setOutputFileName(file_path)
        printer.setPageMargins(QMarginsF(10, 10, 10, 10))
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertBlock()

        footer_format = QTextCharFormat()
        footer_font = QFont("Arial", 10)
        footer_font.setItalic(True)
        footer_format.setFont(footer_font)

        cursor.setCharFormat(footer_format)
        cursor.insertText("\nРуководитель отдела информационных технологий: ____________ /Фамилия И.О./")
        doc.print(printer)

        QMessageBox.information(self, "Успешно", f"Отчёт сохранён:\n{file_path}")
        self.accept()