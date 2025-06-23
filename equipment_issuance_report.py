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

class EquipmentIssuanceReportGenerator(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Отчёт по выдаче техники")
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

        self.employee_filter = QListWidget()
        self.employee_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Сотрудник (ФИО):"))
        layout.addWidget(self.employee_filter)

        self.equipment_filter = QListWidget()
        self.equipment_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Оборудование:"))
        layout.addWidget(self.equipment_filter)

        self.workplace_filter = QListWidget()
        self.workplace_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Рабочее место:"))
        layout.addWidget(self.workplace_filter)

        self.return_filter = QListWidget()
        self.return_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Состояние возврата:"))
        layout.addWidget(self.return_filter)

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd.MM.yyyy")
        self.date_to.setDisplayFormat("dd.MM.yyyy")

        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Дата выдачи от:"))
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
        # Индексы всех колонок
        col_indices = {
            str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower(): col
            for col in range(self.model.columnCount())
        }

        # Единая функция: очищает виджет и добавляет уникальные значения
        def fill_unique(col_name, widget):
            col = col_indices.get(col_name)
            if col is None:
                widget.clear()
                return
            widget.clear()
            seen = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, col))
                if val and val not in seen:
                    widget.addItem(str(val))
                    seen.add(val)

        # Фильтры
        fill_unique("оборудование", self.equipment_filter)
        fill_unique("сотрудник", self.employee_filter)
        fill_unique("рабочее место", self.workplace_filter)
        fill_unique("состояние возврата", self.return_filter)

        # Диапазон даты выдачи
        date_col = col_indices.get("дата выдачи")
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
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одно поле для отчёта.")
            return

        doc = QTextDocument()
        cursor = QTextCursor(doc)

        title_format = QTextCharFormat()
        title_font = QFont("Arial", 16)
        title_font.setBold(True)
        title_format.setFont(title_font)

        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText("Отчёт по выдаче техники\n\n")

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

        # Индексы колонок
        col_map = {
            str(self.model.headerData(i, Qt.Orientation.Horizontal)).strip().lower(): i
            for i in range(self.model.columnCount())
        }

        # Выбранные фильтры
        employees = [item.text() for item in self.employee_filter.selectedItems()]
        equipments = [item.text() for item in self.equipment_filter.selectedItems()]
        workplaces = [item.text() for item in self.workplace_filter.selectedItems()]
        returns = [item.text() for item in self.return_filter.selectedItems()]
        d_from = self.date_from.date()
        d_to = self.date_to.date()

        for row in range(self.model.rowCount()):
            fio_val = self.model.data(self.model.index(row, col_map.get("сотрудник", -1))) or ""
            equipment_val = self.model.data(self.model.index(row, col_map.get("оборудование", -1))) or ""
            workplace_val = self.model.data(self.model.index(row, col_map.get("рабочее место", -1))) or ""
            return_val = self.model.data(self.model.index(row, col_map.get("состояние возврата", -1))) or ""
            date_str = self.model.data(self.model.index(row, col_map.get("дата выдачи", -1))) or ""
            date_obj = QDate.fromString(date_str, "dd.MM.yyyy")

            if employees and fio_val not in employees:
                continue
            if equipments and equipment_val not in equipments:
                continue
            if workplaces and workplace_val not in workplaces:
                continue
            if returns and return_val not in returns:
                continue
            if date_obj.isValid() and (date_obj < d_from or date_obj > d_to):
                continue

            # Добавляем строку в таблицу
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

        # Сохранение PDF
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", os.path.expanduser("~/Отчет_ВыдачаТехники.pdf"), "PDF Files (*.pdf)"
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

