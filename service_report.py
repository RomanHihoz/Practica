from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QDateEdit, QPushButton, QFileDialog,
    QMessageBox, QDoubleSpinBox
)
from PyQt6.QtGui import (
    QTextDocument, QTextCursor, QTextTableFormat, QTextCharFormat,
    QFont, QTextTableCellFormat, QTextLength, QColor
)
from PyQt6.QtCore import Qt, QMarginsF, QDate
from PyQt6.QtPrintSupport import QPrinter
import os
import re 

class ServiceReportGenerator(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Отчёт по обслуживанию")
        self.setMinimumSize(650, 550)
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

        self.type_filter = QListWidget()
        self.type_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Тип работы:"))
        layout.addWidget(self.type_filter)

        self.equipment_filter = QListWidget()
        self.equipment_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Оборудование:"))
        layout.addWidget(self.equipment_filter)

        self.technician_filter = QListWidget()
        self.technician_filter.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(QLabel("Техник:"))
        layout.addWidget(self.technician_filter)


        # Диапазон дат
        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd.MM.yyyy")
        self.date_to.setDisplayFormat("dd.MM.yyyy")

        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Дата обслуживания от:"))
        date_layout.addWidget(self.date_from)
        date_layout.addWidget(QLabel("до:"))
        date_layout.addWidget(self.date_to)
        layout.addLayout(date_layout)

        # Диапазон стоимости
        self.cost_from = QDoubleSpinBox()
        self.cost_to = QDoubleSpinBox()
        self.cost_from.setPrefix("от ")
        self.cost_to.setPrefix("до ")
        self.cost_from.setRange(0, 1_000_000)
        self.cost_to.setRange(0, 1_000_000)
        self.cost_from.setDecimals(2)
        self.cost_to.setDecimals(2)

        cost_layout = QHBoxLayout()
        cost_layout.addWidget(QLabel("Стоимость обслуживания:"))
        cost_layout.addWidget(self.cost_from)
        cost_layout.addWidget(self.cost_to)
        layout.addLayout(cost_layout)

        button_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Сформировать PDF")
        self.cancel_btn = QPushButton("Отмена")
        button_layout.addWidget(self.generate_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        self.generate_btn.clicked.connect(self.generate_report)
        self.cancel_btn.clicked.connect(self.reject)

    def load_filters(self):
        col_indices = {
            str(self.model.headerData(i, Qt.Orientation.Horizontal)).strip().lower(): i
            for i in range(self.model.columnCount())
        }

        def fill_unique(col_name, widget):
            col = col_indices.get(col_name)
            widget.clear()
            if col is None:
                return
            seen = set()
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, col))
                if val and val not in seen:
                    widget.addItem(str(val))
                    seen.add(val)

        fill_unique("тип работы", self.type_filter)
        fill_unique("оборудование", self.equipment_filter)
        fill_unique("техник", self.technician_filter)

        # Дата обслуживания
        date_col = col_indices.get("дата")
        dates = []
        if date_col is not None:
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, date_col))
                dt = QDate.fromString(val, "dd.MM.yyyy")
                if dt.isValid():
                    dates.append(dt)
        self.date_from.setDate(min(dates) if dates else QDate(2020, 1, 1))
        self.date_to.setDate(max(dates) if dates else QDate.currentDate())

        # Стоимость обслуживания
        cost_col = col_indices.get("стоимость")
        costs = []
        if cost_col is not None:
            for row in range(self.model.rowCount()):
                cost_str = self.model.data(self.model.index(row, cost_col)) or ""
                cleaned = re.sub(r"[^\d,.-]", "", cost_str).replace(",", ".")
                try:
                    cost = float(cleaned)
                    costs.append(cost)
                except:
                    continue
                
        if cost_col is not None:
            for row in range(self.model.rowCount()):
                val = self.model.data(self.model.index(row, cost_col))
                try:
                    cost = float(cleaned)
                except:
                    continue
        self.cost_from.setValue(min(costs) if costs else 0.0)
        self.cost_to.setValue(max(costs) if costs else 0.0)

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
        return [
            (self.fields_list.item(i).data(Qt.ItemDataRole.UserRole),
             self.fields_list.item(i).text())
            for i in range(self.fields_list.count())
            if self.fields_list.item(i).checkState() == Qt.CheckState.Checked
        ]

    def generate_report(self):
        selected_fields = self.get_selected_fields()
        if not selected_fields:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одно поле для отчёта.")
            return

        col_map = {
            str(self.model.headerData(i, Qt.Orientation.Horizontal)).strip().lower(): i
            for i in range(self.model.columnCount())
        }

        types = [item.text() for item in self.type_filter.selectedItems()]
        equipments = [item.text() for item in self.equipment_filter.selectedItems()]
        technicians = [item.text() for item in self.technician_filter.selectedItems()]
        d_from = self.date_from.date()
        d_to = self.date_to.date()
        cost_from = self.cost_from.value()
        cost_to = self.cost_to.value()

        doc = QTextDocument()
        cursor = QTextCursor(doc)

        title_format = QTextCharFormat()
        title_font = QFont("Arial", 16)
        title_font.setBold(True)
        title_format.setFont(title_font)
        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText("Отчёт по обслуживанию\n\n")

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
        cell_text_format.setFont(QFont("Arial", 8))

        table.cellAt(0, 0).setFormat(cell_format)
        table.cellAt(0, 0).firstCursorPosition().setCharFormat(header_format)
        table.cellAt(0, 0).firstCursorPosition().insertText("#")

        for i, (_, hdr) in enumerate(selected_fields, 1):
            table.cellAt(0, i).setFormat(cell_format)
            cursor_i = table.cellAt(0, i).firstCursorPosition()
            cursor_i.setCharFormat(header_format)
            cursor_i.insertText(hdr)

        for row in range(self.model.rowCount()):
            type_val = self.model.data(self.model.index(row, col_map.get("тип работы", -1))) or ""
            equipment_val = self.model.data(self.model.index(row, col_map.get("оборудование", -1))) or ""
            technician_val = self.model.data(self.model.index(row, col_map.get("техник", -1))) or ""
            status_val = self.model.data(self.model.index(row, col_map.get("статус", -1))) or ""
            date_str = self.model.data(self.model.index(row, col_map.get("дата", -1))) or ""
            date_obj = QDate.fromString(date_str, "dd.MM.yyyy")
            cost_str = self.model.data(self.model.index(row, col_map.get("стоимость", -1))) or ""
            cleaned = re.sub(r"[^\d,.-]", "", cost_str).replace(",", ".")
            try:
                cost = float(cleaned)
            except:
                cost = 0.0

            # Применяем фильтры
            if types and type_val not in types:
                continue
            if equipments and equipment_val not in equipments:
                continue
            if technicians and technician_val not in technicians:
                continue
            if date_obj.isValid() and (date_obj < d_from or date_obj > d_to):
                continue
            if not (cost_from <= cost <= cost_to):
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

        # Сохраняем отчёт в PDF
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", os.path.expanduser("~/Отчет_Обслуживание.pdf"), "PDF Files (*.pdf)"
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
