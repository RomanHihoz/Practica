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

class WriteOffReportGenerator(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Отчёт по списанному оборудованию")
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

        layout.addWidget(QLabel("Фильтрация:"))

        self.equipment_filter = QListWidget()
        self.equipment_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Оборудование:"))
        layout.addWidget(self.equipment_filter)

        self.reason_filter = QListWidget()
        self.reason_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Причина списания:"))
        layout.addWidget(self.reason_filter)

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd.MM.yyyy")
        self.date_to.setDisplayFormat("dd.MM.yyyy")

        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Дата списания от:"))
        date_layout.addWidget(self.date_from)
        date_layout.addWidget(QLabel("до:"))
        date_layout.addWidget(self.date_to)
        layout.addLayout(date_layout)

        button_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Сформировать PDF")
        self.cancel_btn = QPushButton("Отмена")
        button_layout.addWidget(self.generate_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        self.generate_btn.clicked.connect(self.generate_report)
        self.cancel_btn.clicked.connect(self.reject)

    def load_filters(self):
        equipment_col = reason_col = date_col = None
        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if "оборудование" in header:
                equipment_col = col
            elif "причина" in header:
                reason_col = col
            elif "дата списания" in header:
                date_col = col

        def fill_unique_values(col_index, widget):
            seen = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, col_index))
                if val and val not in seen:
                    widget.addItem(val)
                    seen.add(val)

        if equipment_col is not None:
            fill_unique_values(equipment_col, self.equipment_filter)
        if reason_col is not None:
            fill_unique_values(reason_col, self.reason_filter)

        # Даты
        dates = []
        if date_col is not None:
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, date_col))
                dt = QDate.fromString(val, "dd.MM.yyyy")
                if dt.isValid():
                    dates.append(dt)
        self.date_from.setDate(min(dates) if dates else QDate(2020, 1, 1))
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

        title_format = QTextCharFormat()
        title_font = QFont("Arial", 16)
        title_font.setBold(True)
        title_format.setFont(title_font)
        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText("Отчёт по списанному оборудованию\n\n")

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

        header_format = QTextCharFormat()
        header_font = QFont("Arial", 10)
        header_font.setBold(True)
        header_format.setFont(header_font)

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

        # Фильтры
        equipments = [item.text() for item in self.equipment_filter.selectedItems()]
        reasons = [item.text() for item in self.reason_filter.selectedItems()]
        d_from = self.date_from.date()
        d_to = self.date_to.date()

        equipment_col = reason_col = date_col = None
        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if "оборудование" in header:
                equipment_col = col
            elif "причина" in header:
                reason_col = col
            elif "дата списания" in header:
                date_col = col

        for row in range(self.model.rowCount()):
            eq = self.model.data(self.model.index(row, equipment_col)) if equipment_col is not None else ""
            rs = self.model.data(self.model.index(row, reason_col)) if reason_col is not None else ""
            d_str = self.model.data(self.model.index(row, date_col)) if date_col is not None else ""
            d_obj = QDate.fromString(d_str, "dd.MM.yyyy")

            if equipments and eq not in equipments:
                continue
            if reasons and rs not in reasons:
                continue
            if d_obj.isValid() and (d_obj < d_from or d_obj > d_to):
                continue

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

        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", os.path.expanduser("~/Отчет_СписаниеОборудования.pdf"), "PDF Files (*.pdf)"
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
        cursor.insertText("\nРуководитель отдела ИТ: ____________ /Фамилия И.О./")
        doc.print(printer)

        QMessageBox.information(self, "Успешно", f"Отчёт сохранён:\n{file_path}")
        self.accept()