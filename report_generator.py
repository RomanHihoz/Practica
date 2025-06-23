from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QFileDialog, QMessageBox, QDateEdit
)
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtGui import (
    QTextDocument, QTextCursor, QTextTableFormat, QTextCharFormat,
    QTextBlockFormat, QTextTableCellFormat, QFont, QTextLength, QColor
)
from PyQt6.QtCore import Qt, QMarginsF, QDate
import os

class ReportGenerator(QDialog):
    def __init__(self, table_name, model, parent=None):
        super().__init__(parent)
        self.table_name = table_name
        self.model = model
        self.setWindowTitle(f"Генерация отчета: {table_name}")
        self.setMinimumSize(600, 500)

        self.init_ui()
        self.load_employee_filters()
        self.load_fields()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Заголовок
        title = QLabel(f"Выберите поля для отчета (таблица: {self.table_name})")
        title.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(title)

        # Список полей
        self.fields_list = QListWidget()
        self.fields_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.fields_list)

        # Фильтры
        filter_layout = QVBoxLayout()
        layout.addLayout(filter_layout)

        filter_layout.addWidget(QLabel("Фильтрация по сотрудникам:"))

        self.name_filter = QListWidget()
        self.name_filter.setMaximumHeight(80)
        self.name_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        filter_layout.addWidget(QLabel("Сотрудник:"))
        filter_layout.addWidget(self.name_filter)

        self.department_filter = QListWidget()
        self.department_filter.setMaximumHeight(60)
        self.department_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        filter_layout.addWidget(QLabel("Отдел:"))
        filter_layout.addWidget(self.department_filter)

        self.position_filter = QListWidget()
        self.position_filter.setMaximumHeight(60)
        self.position_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        filter_layout.addWidget(QLabel("Должность:"))
        filter_layout.addWidget(self.position_filter)

        # Диапазон даты рождения
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd.MM.yyyy")
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("dd.MM.yyyy")

        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Дата рождения от:"))
        date_layout.addWidget(self.date_from)
        date_layout.addWidget(QLabel("до:"))
        date_layout.addWidget(self.date_to)
        filter_layout.addLayout(date_layout)

        # Кнопки
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(self.select_all_fields)
        btn_layout.addWidget(self.select_all_btn)

        self.deselect_all_btn = QPushButton("Снять все")
        self.deselect_all_btn.clicked.connect(self.deselect_all_fields)
        btn_layout.addWidget(self.deselect_all_btn)

        self.generate_btn = QPushButton("Сгенерировать отчет")
        self.generate_btn.clicked.connect(self.generate_report)
        btn_layout.addWidget(self.generate_btn)

        self.cancel_btn = QPushButton("Отмена")
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.cancel_btn)

        layout.addLayout(btn_layout)

    def load_employee_filters(self):
        self.name_filter.clear()
        self.department_filter.clear()
        self.position_filter.clear()

        name_col_fam = name_col_name = name_col_patronym = dept_col = pos_col = birth_col = None

        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if header == "фамилия":
                name_col_fam = col
            elif header == "имя":
                name_col_name = col
            elif header == "отчество":
                name_col_patronym = col
            elif "отдел" in header:
                dept_col = col
            elif "должн" in header:
                pos_col = col
            elif "дата рождения" in header or "рожд" in header:
                birth_col = col

        seen_names = set()
        for row in range(self.model.rowCount()):
            fam = self.model.data(self.model.index(row, name_col_fam)) or ""
            name = self.model.data(self.model.index(row, name_col_name)) or ""
            pat = self.model.data(self.model.index(row, name_col_patronym)) or ""
            full_name = f"{fam} {name} {pat}".strip()
            if full_name and full_name not in seen_names:
                self.name_filter.addItem(full_name)
                seen_names.add(full_name)

        if dept_col is not None:
            seen_depts = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, dept_col))
                if val and val not in seen_depts:
                    self.department_filter.addItem(val)
                    seen_depts.add(val)

        if pos_col is not None:
            seen_positions = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, pos_col))
                if val and val not in seen_positions:
                    self.position_filter.addItem(val)
                    seen_positions.add(val)

        dates = []
        if birth_col is not None:
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, birth_col))
                date_obj = QDate.fromString(val, "yyyy-MM-dd")
                if date_obj.isValid():
                    dates.append(date_obj)
        self.date_from.setDate(min(dates) if dates else QDate(1950, 1, 1))
        self.date_to.setDate(max(dates) if dates else QDate.currentDate())

        # Вспомогательная функция заполнения
        def populate_unique_values(col_index, widget):
            if col_index is None:
                return
            seen = set()
            for row in range(self.model.rowCount()):
                fam = self.model.data(self.model.index(row, name_col_fam)) or ""
                name = self.model.data(self.model.index(row, name_col_name)) or ""
                patronym = self.model.data(self.model.index(row, name_col_patronym)) or ""
                full_name = f"{fam} {name} {patronym}".strip()
                if full_name and full_name not in seen:
                    self.name_filter.addItem(full_name)
                    seen.add(full_name)

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

    def select_all_fields(self):
        for i in range(self.fields_list.count()):
            item = self.fields_list.item(i)
            item.setCheckState(Qt.CheckState.Checked)

    def deselect_all_fields(self):
        for i in range(self.fields_list.count()):
            item = self.fields_list.item(i)
            item.setCheckState(Qt.CheckState.Unchecked)

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
            QMessageBox.warning(self, "Ошибка", "Не выбрано ни одного поля для отчета!")
            return

        doc = QTextDocument()
        cursor = QTextCursor(doc)

        title_format = QTextCharFormat()
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_format.setFont(title_font)

        subtitle_format = QTextCharFormat()
        subtitle_font = QFont()
        subtitle_font.setPointSize(12)
        subtitle_format.setFont(subtitle_font)

        header_format = QTextCharFormat()
        header_font = QFont()
        header_font.setPointSize(10)
        header_font.setBold(True)
        header_format.setFont(header_font)

        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText(f"Отчет по таблице: {self.table_name}\n\n", subtitle_format)

        table_format = QTextTableFormat()
        table_format.setCellPadding(4)
        table_format.setBorder(1)
        table_format.setBorderBrush(QColor("black"))
        table_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        table_format.setAlignment(Qt.AlignmentFlag.AlignLeft)

        table = cursor.insertTable(1, len(selected_fields) + 1, table_format)
        for i in range(len(selected_fields) + 1):
            table_format.setColumnWidthConstraints([
                QTextLength(QTextLength.Type.PercentageLength, 100 / (len(selected_fields) + 1))
                for _ in range(len(selected_fields) + 1)
            ])

        cell_format = QTextTableCellFormat()
        cell_format.setBorder(0.5)
        cell_format.setBorderBrush(QColor("black"))

        table.cellAt(0, 0).setFormat(cell_format)
        table.cellAt(0, 0).firstCursorPosition().insertText("#", header_format)

        for i, (_, header_text) in enumerate(selected_fields, 1):
            table.cellAt(0, i).setFormat(cell_format)
            table.cellAt(0, i).firstCursorPosition().insertText(header_text, header_format)

        # Ищем индексы нужных колонок
        name_col_fam = name_col_name = name_col_patronym = dept_col = pos_col = birth_col = None

        for col in range(self.model.columnCount()):
            header = str(self.model.headerData(col, Qt.Orientation.Horizontal)).strip().lower()
            if header == "фамилия":
                name_col_fam = col
            elif header == "имя":
                name_col_name = col
            elif header == "отчество":
                name_col_patronym = col
            elif "отдел" in header:
                dept_col = col
            elif "должн" in header:
                pos_col = col
            elif "дата рождения" in header or "рожд" in header:
                birth_col = col

        if name_col_fam is None or name_col_name is None or name_col_patronym is None:
            QMessageBox.warning(self, "Ошибка", "Не найдены все колонки ФИО в таблице.")
            return

        names = [item.text() for item in self.name_filter.selectedItems()]
        depts = [item.text() for item in self.department_filter.selectedItems()]
        positions = [item.text() for item in self.position_filter.selectedItems()]
        date_start = self.date_from.date()
        date_end = self.date_to.date()

        for row in range(self.model.rowCount()):
            fam = self.model.data(self.model.index(row, name_col_fam)) or ""
            name = self.model.data(self.model.index(row, name_col_name)) or ""
            pat = self.model.data(self.model.index(row, name_col_patronym)) or ""
            full_name = f"{fam} {name} {pat}".strip()

            dept_val = self.model.data(self.model.index(row, dept_col)) if dept_col is not None else ""
            pos_val = self.model.data(self.model.index(row, pos_col)) if pos_col is not None else ""
            date_val = self.model.data(self.model.index(row, birth_col)) if birth_col is not None else ""
            birth_date = QDate.fromString(date_val, "dd.MM.yyyy") if date_val else QDate()

            if names and full_name not in names:
                continue
            if depts and dept_val not in depts:
                continue
            if positions and pos_val not in positions:
                continue
            if birth_date.isValid() and (birth_date < date_start or birth_date > date_end):
                continue

            table.appendRows(1)
            row_index = table.rows() - 1
            table.cellAt(row_index, 0).setFormat(cell_format)
            table.cellAt(row_index, 0).firstCursorPosition().insertText(str(row_index))  # row_index = 1, 2, 3…

            for i, (col, _) in enumerate(selected_fields, 1):
                val = self.model.data(self.model.index(row, col))
                table.cellAt(row_index, i).setFormat(cell_format)
                table.cellAt(row_index, i).firstCursorPosition().insertText(str(val) if val else "")

        # Сохранение
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчет как PDF",
            os.path.expanduser(f"~/Отчет_{self.table_name}.pdf"),
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return

        printer.setOutputFileName(file_path)
        printer.setPageMargins(QMarginsF(15, 15, 15, 15))
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertBlock()

        footer_format = QTextCharFormat()
        footer_font = QFont("Arial", 10)
        footer_font.setItalic(True)
        footer_format.setFont(footer_font)

        cursor.setCharFormat(footer_format)
        cursor.insertText("\nРуководитель отдела информационных технологий: ____________ /Фамилия И.О./")
        doc.print(printer)

        QMessageBox.information(self, "Успешно", f"Отчет сохранен:\n{file_path}")
        self.accept()
