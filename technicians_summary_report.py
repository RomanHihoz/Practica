from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QPushButton, QFileDialog, QMessageBox
)
import re
from PyQt6.QtGui import (
    QTextDocument, QTextCursor, QTextTableFormat, QTextCharFormat,
    QFont, QTextTableCellFormat, QTextLength, QColor
)
from PyQt6.QtCore import Qt, QMarginsF, QDate
from PyQt6.QtPrintSupport import QPrinter
from collections import defaultdict
import os

class TechniciansSummaryReport(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сводный отчёт по техникам")
        self.setMinimumSize(500, 400)
        self.model = model

        layout = QVBoxLayout(self)
        self.generate_btn = QPushButton("Сформировать сводный PDF")
        self.cancel_btn = QPushButton("Отмена")
        layout.addWidget(self.generate_btn)
        layout.addWidget(self.cancel_btn)

        self.generate_btn.clicked.connect(self.generate_report)
        self.cancel_btn.clicked.connect(self.reject)

    def generate_report(self):
        col_map = {
            str(self.model.headerData(i, Qt.Orientation.Horizontal)).strip().lower(): i
            for i in range(self.model.columnCount())
        }
        print("Найденные колонки:", list(col_map.keys()))

        tech_col = col_map.get("техник")
        eq_col = col_map.get("оборудование")
        date_col = col_map.get("дата")
        cost_col = col_map.get("стоимость")

        if None in (tech_col, eq_col, date_col, cost_col):
            QMessageBox.warning(self, "Ошибка", "Не найдены все нужные поля в таблице.")
            return

        # Собираем данные по техникам
        summary = defaultdict(lambda: {
            "count": 0,
            "equipment": set(),
            "dates": [],
            "sum": 0.0
        })

        for row in range(self.model.rowCount()):
            tech = self.model.data(self.model.index(row, tech_col)) or ""
            eq = self.model.data(self.model.index(row, eq_col)) or ""
            date_str = self.model.data(self.model.index(row, date_col)) or ""
            cost_str = self.model.data(self.model.index(row, cost_col)) or ""
            cleaned = re.sub(r"[^\d,.-]", "", cost_str).replace(",", ".")
            try:
                cost = float(cleaned)
            except:
                cost = 0.0

            if tech:
                summary[tech]["count"] += 1
                summary[tech]["equipment"].add(eq)
                summary[tech]["sum"] += cost

                # 🔽 Добавь вот эту часть:
                formats = ["dd.MM.yyyy", "yyyy-MM-dd", "dd/MM/yyyy", "yyyy.MM.dd"]
                dt = QDate()
                for fmt in formats:
                    test = QDate.fromString(date_str, fmt)
                    if test.isValid():
                        dt = test
                        break
                if dt.isValid():
                    summary[tech]["dates"].append(dt)

        # Готовим документ
        doc = QTextDocument()
        cursor = QTextCursor(doc)

        title_format = QTextCharFormat()
        title_font = QFont("Arial", 16)
        title_font.setBold(True)
        title_format.setFont(title_font)
        cursor.insertText("ООО «Разработчики Программ»\n", title_format)
        cursor.insertText("Сводный отчёт по техникам\n\n")

        table_format = QTextTableFormat()
        table_format.setCellPadding(4)
        table_format.setBorder(1)
        table_format.setBorderBrush(QColor("black"))
        table_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        table_format.setColumnWidthConstraints([
            QTextLength(QTextLength.Type.PercentageLength, perc) for perc in [20, 10, 25, 10, 10, 25]
        ])
        table = cursor.insertTable(1, 6, table_format)

        headers = ["Техник", "Заявок", "Устройства", "Сумма (₽)", "Средняя", "Период"]
        header_format = QTextCharFormat()
        header_font = QFont("Arial", 10)
        header_font.setBold(True)
        header_format.setFont(header_font)

        cell_format = QTextTableCellFormat()
        cell_format.setBorder(0.5)
        cell_format.setBorderBrush(QColor("black"))

        cell_text_format = QTextCharFormat()
        cell_text_format.setFont(QFont("Arial", 8))

        for i, h in enumerate(headers):
            table.cellAt(0, i).setFormat(cell_format)
            cursor_h = table.cellAt(0, i).firstCursorPosition()
            cursor_h.setCharFormat(header_format)
            cursor_h.insertText(h)

        # Заполняем строки
        for tech, data in summary.items():
            if data["dates"]:
                min_d = min(data["dates"]).toString("dd.MM.yyyy")
                max_d = max(data["dates"]).toString("dd.MM.yyyy")
                period = f"{min_d} — {max_d}"
            table.appendRows(1)
            r_idx = table.rows() - 1

            total = data["sum"]
            count = data["count"]
            avg = round(total / count, 2) if count else 0
            period = ""
            if data["dates"]:
                min_d = min(data["dates"]).toString("dd.MM.yyyy")
                max_d = max(data["dates"]).toString("dd.MM.yyyy")
                period = f"{min_d} — {max_d}"

            values = [
                tech,
                str(count),
                ", ".join(sorted(data["equipment"])),
                f"{total:.2f}",
                f"{avg:.2f}",
                period
            ]

            for i, val in enumerate(values):
                table.cellAt(r_idx, i).setFormat(cell_format)
                cursor_i = table.cellAt(r_idx, i).firstCursorPosition()
                cursor_i.setCharFormat(cell_text_format)
                cursor_i.insertText(val)

        # Сохраняем PDF
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить сводный отчёт", os.path.expanduser("~/Сводка_по_техникам.pdf"), "PDF Files (*.pdf)"
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
