from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QDateEdit, QPushButton, QFileDialog, QMessageBox
)
from PyQt6.QtGui import (
    QTextDocument, QTextCursor, QTextTableFormat, QTextCharFormat,
    QFont, QTextTableCellFormat, QTextLength, QColor
)
from PyQt6.QtCore import Qt, QMarginsF, QDate
from PyQt6.QtPrintSupport import QPrinter
import os

class WorkplaceReportGenerator(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Отчёт по рабочим местам")
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

        self.building_filter = QListWidget()
        self.building_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Корпус:"))
        layout.addWidget(self.building_filter)

        self.floor_filter = QListWidget()
        self.floor_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Этаж:"))
        layout.addWidget(self.floor_filter)

        self.room_filter = QListWidget()
        self.room_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Кабинет:"))
        layout.addWidget(self.room_filter)

        self.status_filter = QListWidget()
        self.status_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Статус:"))
        layout.addWidget(self.status_filter)

        self.desk_filter = QListWidget()
        self.desk_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Стол:"))
        layout.addWidget(self.desk_filter)

        button_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Сформировать PDF")
        self.cancel_btn = QPushButton("Отмена")
        button_layout.addWidget(self.generate_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        self.generate_btn.clicked.connect(self.generate_report)
        self.cancel_btn.clicked.connect(self.reject)

    def load_filters(self):
        building_col = floor_col = room_col = desk_col = status_col = None

        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if "корпус" in header:
                building_col = col
            elif "этаж" in header:
                floor_col = col
            elif "кабинет" in header:
                room_col = col
            elif "стол" in header:
                desk_col = col
            elif "статус" in header:
                status_col = col

        def fill_unique(col_index, widget):
            if col_index is None:
                return
            seen = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, col_index))
                if val and val not in seen:
                    widget.addItem(str(val))
                    seen.add(val)

        fill_unique(building_col, self.building_filter)
        fill_unique(floor_col, self.floor_filter)
        fill_unique(room_col, self.room_filter)
        fill_unique(desk_col, self.desk_filter)
        fill_unique(status_col, self.status_filter)

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
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одно поле для отчёта.")
            return

        doc = QTextDocument()
        cursor = QTextCursor(doc)

        title_format = QTextCharFormat()
        title_font = QFont("Arial", 16)
        title_font.setBold(True)
        title_format.setFont(title_font)
        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText("Отчёт по рабочим местам\n\n")

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
        table.cellAt(0, 0).firstCursorPosition().setCharFormat(header_format)
        table.cellAt(0, 0).firstCursorPosition().insertText("#")

        for i, (_, hdr) in enumerate(selected_fields, 1):
            table.cellAt(0, i).setFormat(cell_format)
            table.cellAt(0, i).firstCursorPosition().setCharFormat(header_format)
            table.cellAt(0, i).firstCursorPosition().insertText(hdr)

        # Фильтры
        buildings = [item.text() for item in self.building_filter.selectedItems()]
        floors = [item.text() for item in self.floor_filter.selectedItems()]
        rooms = [item.text() for item in self.room_filter.selectedItems()]
        desks = [item.text() for item in self.desk_filter.selectedItems()]
        statuses = [item.text() for item in self.status_filter.selectedItems()]

        # Определим индексы
        col_indices = {}
        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            col_indices[header] = col

        for row in range(self.model.rowCount()):
            vals = {}
            for key in ["корпус", "этаж", "кабинет", "стол"]:
                col = col_indices.get(key)
                val = self.model.data(self.model.index(row, col)) if col is not None else ""
                vals[key] = str(val)
                status_val = self.model.data(self.model.index(row, col_indices.get("статус"))) or ""

            if buildings and vals["корпус"] not in buildings:
                continue
            if floors and vals["этаж"] not in floors:
                continue
            if rooms and vals["кабинет"] not in rooms:
                continue
            if desks and vals["стол"] not in desks:
                continue
            if statuses and status_val not in statuses:
                continue

            # Строка проходит фильтры — добавляем
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
            self, "Сохранить отчёт", os.path.expanduser("~/Отчет_РабочиеМеста.pdf"), "PDF Files (*.pdf)"
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
