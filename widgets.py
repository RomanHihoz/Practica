from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton,
    QTableView, QLabel, QComboBox, QMessageBox, QDialog, QFormLayout, QDialogButtonBox, QStyledItemDelegate, QStyleOptionViewItem, QDateEdit, QHeaderView, QFrame
)
from PyQt6.QtSql import QSqlTableModel, QSqlRelationalTableModel, QSqlRelation
from PyQt6.QtSql import QSqlQuery
from PyQt6.QtCore import Qt, QRegularExpression, QDate, QItemSelectionModel, QItemSelection
from PyQt6.QtGui import QRegularExpressionValidator, QBrush, QColor
import re
from PyQt6.QtCore import QSortFilterProxyModel
from PyQt6.QtSql    import QSqlQueryModel

DEPARTMENTS = [
    "ИТ", "Бухгалтерия", "Отдел кадров"
]
POSITIONS_BY_DEPARTMENT = {
    "ИТ": ["Системный администратор", "Инженер-программист", "Техник"],
    "Бухгалтерия": ["Бухгалтер", "Главный бухгалтер", "Кассир"],
    "Отдел кадров": ["HR-менеджер", "Специалист по кадрам", "Руководитель отдела кадров"]
}
EQUIP_TYPES = ["Компьютер", "Монитор", "Принтер", "Сканер", "Другое"]
EQUIP_STATUS = ["В эксплуатации", "На складе", "Списано", "В ремонте"]
WORKPLACE_STATUS = ["Активно", "Не используется", "В ремонте", "Зарезервировано"]
WORK_TYPES = [
    "Профилактика",
    "Ремонт",
    "Замена комплектующих",
    "Тестирование",
    "Настройка",
    "Другое"
]
class CostModel(QSqlQueryModel):
    """
    QSqlQueryModel, который форматирует колонку "Стоимость" с двумя десятичными
    и добавляет символ ₽. Сохраняет исходные данные в модели как числа.
    """
    def __init__(self, sql_query: str, parent=None):
        super().__init__(parent)
        self._sql = sql_query
        # при установке запроса сразу вычисляем индекс колонки «Стоимость»
        self.setQuery(self._sql)

    def setQuery(self, query: str):
        # переопределяем, чтобы после каждого запроса найти, где «Стоимость»
        super().setQuery(query)
        self._calc_cost_column()

    def _calc_cost_column(self):
        # ищем колонку «Стоимость» по её заголовку
        self.cost_col = None
        for col in range(self.columnCount()):
            hdr = self.headerData(col, Qt.Orientation.Horizontal)
            if str(hdr) == "Стоимость":
                self.cost_col = col
                break

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        # если это вывод в ячейку и наша колонка «Стоимость» — форматируем
        if role == Qt.ItemDataRole.DisplayRole and index.column() == self.cost_col:
            raw = super().data(index, role)
            if raw is None or raw == "":
                return ""
            # пытаемся превратить в число
            try:
                amount = float(raw)
            except Exception:
                return raw
            # форматируем: разделяем тысячи пробелами, 2 знака после точки
            formatted = f"{amount:,.2f}".replace(",", " ")
            return f"{formatted} ₽"
        # всё остальное возвращаем без изменений
        return super().data(index, role)
    
def beautify_header(header):
    header = header.replace('_', ' ').strip()
    if header.lower().endswith(' id'):
        header = header[:-3]
    return header.capitalize()

def is_valid_email(email):
    # Примитивная проверка на валидный email с доменом
    return re.match(r"^[^@]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email)

def is_valid_phone(phone):
    # Российский формат: +7XXXXXXXXXX или 8XXXXXXXXXX
    return re.match(r"^(\+7|8)\d{10}$", phone)

class MaintenanceFilterProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.filters = []  # [(value_or_tuple, source_col), ...]

    def filterAcceptsRow(self, src_row, src_parent):
        model = self.sourceModel()
        for value, col in self.filters:
            idx = model.index(src_row, col, src_parent)
            cell = str(idx.data() or "")
            # диапазон дат — кортеж (dd.MM.yyyy, dd.MM.yyyy)
            if isinstance(value, tuple):
                fr, to = value
                if fr and to:
                    if not (fr <= cell <= to):
                        return False
            else:
                if value and value.lower() not in cell.lower():
                    return False
        return True

    def set_maintenance_filters(self, flts):
        self.filters = flts
        self.invalidateFilter()

class PhoneDelegate(QStyledItemDelegate):
    def displayText(self, value, locale):
        val = str(value)
        if val and not val.startswith("+"):
            return "+" + val
        return val

class HighlightDelegate(QStyledItemDelegate):
    def __init__(self, search_text="", parent=None):
        super().__init__(parent)
        self.search_text = search_text.lower()

    def paint(self, painter, option, index):
        text = str(index.data()) if index.data() is not None else ""
        if self.search_text and self.search_text in text.lower():
            option_new = QStyleOptionViewItem(option)  # обязательно копия!
            option_new.backgroundBrush = QBrush(QColor("#3a7bd5"))
            super().paint(painter, option_new, index)
        else:
            super().paint(painter, option, index)

class CombinedDelegate(QStyledItemDelegate):
    def __init__(self, search_text, base_delegate, parent=None):
        super().__init__(parent)
        self.search_text = search_text.lower()
        self.base_delegate = base_delegate

    def paint(self, painter, option, index):
        text = str(index.data()) if index.data() is not None else ""
        if self.search_text and self.search_text in text.lower():
            option_new = QStyleOptionViewItem(option)  # обязательно копия!
            option_new.backgroundBrush = QBrush(QColor("#3a7bd5"))
            self.base_delegate.paint(painter, option_new, index)
        else:
            self.base_delegate.paint(painter, option, index)

    def displayText(self, value, locale):
        return self.base_delegate.displayText(value, locale)

class RowHighlightDelegate(QStyledItemDelegate):
    def __init__(self, search_text="", parent=None):
        super().__init__(parent)
        self.search_text = search_text.lower()

    def paint(self, painter, option, index):
        if self.search_text:
            model = index.model()
            row = index.row()
            found = False
            for col in range(model.columnCount()):
                # Используем sibling для корректной работы с прокси-моделями
                cell = index.sibling(row, col).data()
                cell_str = str(cell) if cell is not None else ""
                if self.search_text in cell_str.lower():
                    found = True
                    break
            if found:
                option_new = QStyleOptionViewItem(option)
                option_new.backgroundBrush = QBrush(QColor("#ffe066"))  # Ярко-жёлтый
                super().paint(painter, option_new, index)
                return
        super().paint(painter, option, index)



class TableWidget(QWidget):
    def __init__(self, table_name, parent=None):
        super().__init__(parent)
        self.table_name = table_name
        self.editable_table = table_name
        self._sql = ""
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        search_panel = QWidget()
        search_layout = QHBoxLayout(search_panel)
        search_layout.setContentsMargins(0, 0, 0, 0)
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Поиск...")
        search_layout.addWidget(QLabel("Поиск:"))
        search_layout.addWidget(self.search_field)
        main_layout.addWidget(search_panel)

        filter_panel = QWidget()
        filter_layout = QHBoxLayout(filter_panel)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        self.filter_btn = QPushButton("Фильтры")
        self.reset_filter_btn = QPushButton("Сбросить фильтры")
        self.filter_btn.clicked.connect(self.open_filter_dialog)
        self.reset_filter_btn.clicked.connect(self.reset_filters)
        filter_layout.addWidget(self.filter_btn)
        filter_layout.addWidget(self.reset_filter_btn)
        main_layout.addWidget(filter_panel)

        self.table_view = QTableView()
        main_layout.addWidget(self.table_view, stretch=1)

        control_panel = QWidget()
        control_layout = QHBoxLayout(control_panel)
        control_layout.setContentsMargins(0, 0, 0, 0)
        self.add_btn = QPushButton("Добавить")
        self.edit_btn = QPushButton("Редактировать")
        self.del_btn = QPushButton("Удалить")
        self.report_btn = QPushButton("Отчет")
        for btn in [self.add_btn, self.edit_btn, self.del_btn, self.report_btn]:
            btn.setMinimumWidth(150)
            btn.setMinimumHeight(40)
        control_layout.addWidget(self.add_btn)
        control_layout.addWidget(self.edit_btn)
        control_layout.addWidget(self.del_btn)
        control_layout.addWidget(self.report_btn)
        main_layout.addWidget(control_panel)

        self.search_field.textChanged.connect(self.apply_filter)
        self.add_btn.clicked.connect(self.add_row)
        self.edit_btn.clicked.connect(self.edit_row)
        self.del_btn.clicked.connect(self.delete_row)
        self.report_btn.clicked.connect(self.generate_report)

        self.setup_table()

    def generate_report(self):
        """Открывает диалог генерации отчета"""
        from report_generator import ReportGenerator
        if self.table_name == "СОТРУДНИК":
            dialog = ReportGenerator(self.table_name, self.model, self)
            dialog.exec()
        elif self.table_name == "ОБОРУДОВАНИЕ":
            from equipment_report import EquipmentReportGenerator
            dialog = EquipmentReportGenerator(self.model, self)
            dialog.exec()
        elif self.table_name == "РАБОЧЕЕ_МЕСТО":
            from workplace_report import WorkplaceReportGenerator
            dialog = WorkplaceReportGenerator(self.model, self)
            dialog.exec()
        elif self.table_name == "ВЫДАЧА_ТЕХНИКИ":
            from equipment_issuance_report import EquipmentIssuanceReportGenerator
            dialog = EquipmentIssuanceReportGenerator(self.model, self)
            dialog.exec()
        elif self.table_name == "ОБСЛУЖИВАНИЕ":
            from PyQt6.QtWidgets import QMessageBox
            msg = QMessageBox(self)
            msg.setWindowTitle("Выбор типа отчёта")
            msg.setText("Выберите, какой отчёт сформировать:")
            detail_btn = msg.addButton("Подробный отчёт", QMessageBox.ButtonRole.AcceptRole)
            summary_btn = msg.addButton("Сводка по техникам", QMessageBox.ButtonRole.AcceptRole)
            cancel_btn = msg.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
            msg.exec()

            if msg.clickedButton() == detail_btn:
                from service_report import ServiceReportGenerator
                dialog = ServiceReportGenerator(self.model, self)
                dialog.exec()
            elif msg.clickedButton() == summary_btn:
                from technicians_summary_report import TechniciansSummaryReport
                report = TechniciansSummaryReport(self.model, self)
                report.generate_report()  # <-- вызывает метод напрямую без диалога

    def setup_table(self):
        # --- Используем QSqlRelationalTableModel для ВЫДАЧА_ТЕХНИКИ ---
        if self.table_name == "ОБСЛУЖИВАНИЕ":
            self._sql = """
                SELECT
                m.id AS id,
                o.название||' ('||o.производитель||', SN:'||o.серийный_номер||')' AS Оборудование,
                m.дата        AS Дата,
                m.тип_работы  AS "Тип работы",
                m.описание    AS Описание,
                s.фамилия||' '||s.имя||COALESCE(' '||s.отчество,'') AS Техник,
                m.стоимость   AS Стоимость
                FROM ОБСЛУЖИВАНИЕ m
                JOIN ОБОРУДОВАНИЕ o ON m.оборудование_id = o.id
                LEFT JOIN СОТРУДНИК s ON m.техник_id = s.id
            """
            self.model = CostModel(self._sql, self)
            self.proxy = MaintenanceFilterProxy(self)
            self.proxy.setSourceModel(self.model)
        elif self.table_name == "ВЫДАЧА_ТЕХНИКИ":
            self._sql = """
                SELECT
                w.id AS id,
                o.название||' ('||o.производитель||', SN:'||o.серийный_номер||')' AS Оборудование,
                s.фамилия||' '||s.имя||COALESCE(' '||s.отчество,'') AS Сотрудник,
                wp.корпус||' | этаж '||wp.этаж||' | каб. '||wp.кабинет||' | стол '||wp.стол AS Рабочее_место,
                w.дата_выдачи        AS Дата_выдачи,
                w.состояние_при_возврате AS Состояние_возврата
                FROM ВЫДАЧА_ТЕХНИКИ w
                JOIN ОБОРУДОВАНИЕ  o  ON w.оборудование_id   = o.id
                JOIN СОТРУДНИК     s  ON w.сотрудник_id      = s.id
                JOIN РАБОЧЕЕ_МЕСТО wp ON w.рабочее_место_id = wp.id
            """
            self.model = QSqlQueryModel(self)
            self.model.setQuery(self._sql)
            self.proxy = IssuanceFilterProxy(self)
            self.proxy.setSourceModel(self.model)
        else:
            self.model = QSqlTableModel(self)
            self.model.setTable(self.table_name)
            self.model.select()
            self.proxy = QSortFilterProxyModel(self)
            self.proxy.setSourceModel(self.model)

        self.proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.table_view.setModel(self.proxy)
        self.table_view.setSortingEnabled(True)

        self.id_column = None
        for col in range(self.model.columnCount()):
            hdr = self.model.headerData(col, Qt.Orientation.Horizontal)
            if str(hdr).lower() == "id":
                self.id_column = col
                break

        if self.id_column is not None:
            self.table_view.hideColumn(self.id_column)

        # Красивые заголовки
        for col in range(self.model.columnCount()):
            hdr = self.model.headerData(col, Qt.Orientation.Horizontal)
            self.model.setHeaderData(col, Qt.Orientation.Horizontal, beautify_header(str(hdr)))

        # Авто-подгонка размеров
        self.table_view.resizeColumnsToContents()
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.resizeRowsToContents()

        # Особый делегат для телефона
        if self.table_name == "СОТРУДНИК":
            for col in range(self.model.columnCount()):
                hdr = str(self.model.headerData(col, Qt.Orientation.Horizontal)).lower()
                if "номер телефона" in hdr:
                    self.table_view.setItemDelegateForColumn(col, PhoneDelegate(self))

        self.table_view.setStyleSheet("""
            QTableView::item:selected {
                background: #ffe066;
                color: #222;
            }
        """)

        # Настройка ширины столбцов
        header = self.table_view.horizontalHeader()
        for col in range(self.model.columnCount()):
            header_name = self.model.headerData(col, Qt.Orientation.Horizontal)
            if header_name in ("Оборудование", "Сотрудник", "Рабочее место"):
                header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
            else:
                header.setSectionResizeMode(col, QHeaderView.ResizeMode.Stretch)


    def apply_filter(self):
        text = self.search_field.text().strip()
        selection_model = self.table_view.selectionModel()
        selection_model.clearSelection()

        if not text:
            return

        # Проверяем формат "Заголовок значение"
        if " " in text:
            header_part, value_part = text.split(" ", 1)
            header_part = header_part.strip().lower()
            value_part = value_part.strip().lower()
            # Найдём столбец по заголовку
            target_col = None
            for col in range(self.proxy.columnCount()):
                header = str(self.proxy.headerData(col, Qt.Orientation.Horizontal)).lower()
                if header_part in header:
                    target_col = col
                    break
            if target_col is not None:
                matches = []
                for row in range(self.proxy.rowCount()):
                    idx = self.proxy.index(row, target_col)
                    cell = idx.data()
                    cell_str = str(cell) if cell is not None else ""
                    if value_part in cell_str.lower():
                        matches.append(row)
                for row in matches:
                    selection_model.select(
                        self.proxy.index(row, 0),
                        QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows
                    )
                return  # Не ищем по всем столбцам, если найден шаблон

        # Обычный поиск по всем столбцам
        matches = []
        text = text.lower()
        for row in range(self.proxy.rowCount()):
            for col in range(self.proxy.columnCount()):
                idx = self.proxy.index(row, col)
                cell = idx.data()
                cell_str = str(cell) if cell is not None else ""
                if text in cell_str.lower():
                    matches.append(row)
                    break

        for row in matches:
            selection_model.select(
                self.proxy.index(row, 0),
                QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows
            )

    def add_row(self):
        # 1) Ветка для ВЫДАЧА_ТЕХНИКИ (QSqlQueryModel)
        if self.table_name == "ВЫДАЧА_ТЕХНИКИ":
            headers = ["Оборудование", "Сотрудник", "Рабочее место", "Дата выдачи", "Состояние возврата"]
            dialog = AddRowDialog(headers, self.editable_table, self)
            if not dialog.exec():
                return
            vals = dialog.get_data(self.table_name)

            q = QSqlQuery()
            q.prepare("""
                INSERT INTO ВЫДАЧА_ТЕХНИКИ
                  (оборудование_id, сотрудник_id, рабочее_место_id, дата_выдачи, состояние_при_возврате)
                VALUES (?, ?, ?, ?, ?)
            """)
            for i, v in enumerate(vals):
                q.bindValue(i, v)
            if not q.exec():
                QMessageBox.critical(self, "Ошибка добавления", q.lastError().text())
            else:
                # пересоздаём результат запроса
                self.model.setQuery(self._sql)
            return

        # 2) Ветка для ОБСЛУЖИВАНИЕ (CostModel → QSqlQueryModel)
        if self.table_name == "ОБСЛУЖИВАНИЕ":
            headers = ["Оборудование", "Дата", "Тип работы", "Описание", "Техник", "Стоимость"]
            dialog = AddRowDialog(headers, self.editable_table, self)
            if not dialog.exec():
                return
            vals = dialog.get_data(self.table_name)

            q = QSqlQuery()
            q.prepare("""
                INSERT INTO ОБСЛУЖИВАНИЕ
                  (оборудование_id, дата, тип_работы, описание, техник_id, стоимость)
                VALUES (?, ?, ?, ?, ?, ?)
            """)
            for i, v in enumerate(vals):
                q.bindValue(i, v)
            if not q.exec():
                QMessageBox.critical(self, "Ошибка добавления", q.lastError().text())
            else:
                # заново прогоняем SQL и CostModel найдёт колонку «Стоимость»
                self.model.setQuery(self._sql)
            return

        # 3) Всё остальное — QSqlTableModel
        # собираем headers/columns
        headers = []
        columns = []
        for col in range(self.model.columnCount()):
            if col == self.id_column:
                continue
            headers.append(self.model.headerData(col, Qt.Orientation.Horizontal))
            columns.append(col)

        dialog = AddRowDialog(headers, self.editable_table, self)
        if not dialog.exec():
            return

        values = dialog.get_data(self.table_name)
        edit_model = QSqlTableModel(self)
        edit_model.setTable(self.editable_table)
        edit_model.select()
        row = edit_model.rowCount()
        edit_model.insertRow(row)
        for i, col in enumerate(columns):
            edit_model.setData(edit_model.index(row, col), values[i])
        edit_model.submitAll()

        # обновляем представление
        self.model.select()

        # если это СОТРУДНИК — заполним человеко-читаемое рабочее место
        if self.table_name == "СОТРУДНИК":
            query = QSqlQuery("SELECT id, корпус, этаж, кабинет, стол FROM РАБОЧЕЕ_МЕСТО")
            wp = {}
            while query.next():
                wid = query.value(0)
                label = f"{query.value(1)} | этаж {query.value(2)} | каб. {query.value(3)} | стол {query.value(4)}"
                wp[wid] = label
            for r in range(self.model.rowCount()):
                idx = self.model.index(r, self.model.fieldIndex("рабочее_место_id"))
                wid = self.model.data(idx)
                if wid in wp:
                    self.model.setData(idx, wp[wid])
        elif self.table_name == "ОБСЛУЖИВАНИЕ":
            headers = ["Оборудование", "Дата", "Тип работы", "Описание", "Техник", "Стоимость"]
            dialog = AddRowDialog(headers, self.editable_table, self)
            if not dialog.exec():
                return

            vals = dialog.get_data(self.table_name)
            q = QSqlQuery()
            q.prepare("""
                INSERT INTO ОБСЛУЖИВАНИЕ
                (оборудование_id, дата, тип_работы, описание, техник_id, стоимость)
                VALUES (?, ?, ?, ?, ?, ?)
            """)
            for i, v in enumerate(vals):
                q.bindValue(i, v)
            if not q.exec():
                QMessageBox.critical(self, "Ошибка добавления", q.lastError().text())
            else:
                self.model.setQuery(self._sql)
            return

    def edit_row(self):
        proxy_index = self.table_view.currentIndex()
        if not proxy_index.isValid():
            QMessageBox.warning(self, "Редактирование", "Выберите строку для редактирования.")
            return

        source_index = self.proxy.mapToSource(proxy_index)
        row = source_index.row()

        # --- ВЫДАЧА_ТЕХНИКИ ---
        if self.table_name == "ВЫДАЧА_ТЕХНИКИ":
            rec = self.model.record(row)
            rec_id = rec.value("id")
            q0 = QSqlQuery()
            q0.prepare("""
                SELECT оборудование_id, сотрудник_id, рабочее_место_id,
                    дата_выдачи, состояние_при_возврате
                FROM ВЫДАЧА_ТЕХНИКИ WHERE id = ?
            """)
            q0.bindValue(0, rec_id)
            if not q0.exec() or not q0.next():
                QMessageBox.critical(self, "Ошибка", "Не удалось загрузить данные.")
                return
            orig = [q0.value(i) for i in range(5)]
            headers = ["Оборудование", "Сотрудник", "Рабочее место", "Дата выдачи", "Состояние возврата"]
            dialog = EditRowDialog(headers, orig, self.editable_table, self)
            if not dialog.exec():
                return
            new_vals = dialog.get_data(self.table_name)
            q = QSqlQuery()
            q.prepare("""
                UPDATE ВЫДАЧА_ТЕХНИКИ
                SET оборудование_id = ?, сотрудник_id = ?, рабочее_место_id = ?,
                    дата_выдачи = ?, состояние_при_возврате = ?
                WHERE id = ?
            """)
            for i, v in enumerate(new_vals):
                q.bindValue(i, v)
            q.bindValue(len(new_vals), rec_id)
            if not q.exec():
                QMessageBox.critical(self, "Ошибка редактирования", q.lastError().text())
            else:
                self.model.setQuery(self._sql)
            return

        # --- ОБСЛУЖИВАНИЕ ---
        if self.table_name == "ОБСЛУЖИВАНИЕ":
            rec = self.model.record(row)
            rec_id = rec.value("id")
            q0 = QSqlQuery()
            q0.prepare("""
                SELECT оборудование_id, дата, тип_работы, описание, техник_id, стоимость
                FROM ОБСЛУЖИВАНИЕ WHERE id = ?
            """)
            q0.bindValue(0, rec_id)
            if not q0.exec() or not q0.next():
                QMessageBox.critical(self, "Ошибка", "Не удалось загрузить данные.")
                return
            orig = [q0.value(i) for i in range(6)]
            headers = ["Оборудование", "Дата", "Тип работы", "Описание", "Техник", "Стоимость"]
            dialog = EditRowDialog(headers, orig, self.editable_table, self)
            if not dialog.exec():
                return
            new_vals = dialog.get_data(self.table_name)
            q = QSqlQuery()
            q.prepare("""
                UPDATE ОБСЛУЖИВАНИЕ
                SET оборудование_id = ?, дата = ?, тип_работы = ?,
                    описание = ?, техник_id = ?, стоимость = ?
                WHERE id = ?
            """)
            for i, v in enumerate(new_vals):
                q.bindValue(i, v)
            q.bindValue(len(new_vals), rec_id)
            if not q.exec():
                QMessageBox.critical(self, "Ошибка редактирования", q.lastError().text())
            else:
                self.model.setQuery(self._sql)
            return

        # --- Все остальные таблицы ---
        headers = []
        columns = []
        values = []
        for col in range(self.model.columnCount()):
            if col == self.id_column:
                continue
            headers.append(self.model.headerData(col, Qt.Orientation.Horizontal))
            columns.append(col)
            values.append(self.model.data(self.model.index(row, col)))

        dialog = EditRowDialog(headers, values, self.editable_table, self)
        if dialog.exec():
            new_values = dialog.get_data(self.table_name)
            for i, col in enumerate(columns):
                self.model.setData(self.model.index(row, col), new_values[i])
            self.model.submitAll()
            self.model.select()

            if self.table_name == "СОТРУДНИК":
                query = QSqlQuery("SELECT id, корпус, этаж, кабинет, стол FROM РАБОЧЕЕ_МЕСТО")
                workplace_map = {}
                while query.next():
                    wid = query.value(0)
                    label = f"{query.value(1)} | этаж {query.value(2)} | каб. {query.value(3)} | стол {query.value(4)}"
                    workplace_map[wid] = label
                for r in range(self.model.rowCount()):
                    idx2 = self.model.index(r, self.model.fieldIndex("рабочее_место_id"))
                    wid = self.model.data(idx2)
                    if wid in workplace_map:
                        self.model.setData(idx2, workplace_map[wid])

    def delete_row(self):
        from PyQt6.QtSql import QSqlQuery

        idx = self.table_view.currentIndex()
        proxy_index = self.table_view.currentIndex()
        if not proxy_index.isValid():
            QMessageBox.warning(self, "Редактирование", "Выберите строку для редактирования.")
            return

        source_index = self.proxy.mapToSource(proxy_index)
        row = source_index.row()

        if self.table_name == "ВЫДАЧА_ТЕХНИКИ":
            rec = self.model.record(source_index.row())
            rec_id = rec.value("id")
            ans = QMessageBox.question(
                self, "Удаление", "Удалить выбранную запись?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if ans == QMessageBox.StandardButton.Yes:
                q = QSqlQuery()
                if not q.exec(f"DELETE FROM ВЫДАЧА_ТЕХНИКИ WHERE id = {rec_id}"):
                    QMessageBox.critical(self, "Ошибка удаления", q.lastError().text())
                else:
                    self.model.setQuery(self._sql)
            return

        if self.table_name == "ОБСЛУЖИВАНИЕ":
            rec = self.model.record(source_index.row())
            rec_id = rec.value("id")
            ans = QMessageBox.question(
                self, "Удаление", "Удалить выбранную запись?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if ans == QMessageBox.StandardButton.Yes:
                q = QSqlQuery()
                if not q.exec(f"DELETE FROM ОБСЛУЖИВАНИЕ WHERE id = {rec_id}"):
                    QMessageBox.critical(self, "Ошибка удаления", q.lastError().text())
                else:
                    self.model.setQuery(self._sql)
            return
        
        ans = QMessageBox.question(
            self, "Удаление", "Удалить выбранную запись?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if ans == QMessageBox.StandardButton.Yes:
            self.model.removeRow(idx.row())
            self.model.submitAll()
            self.model.select()

    def open_filter_dialog(self):
        headers = []
        columns = []
        real_headers = []
        for col in range(self.model.columnCount()):
            if col == self.id_column:
                continue
            header = self.model.headerData(col, Qt.Orientation.Horizontal)
            headers.append(header)
            columns.append(col)
            # Получаем реальное имя столбца из record()
            real_headers.append(self.model.record().fieldName(col))
        dialog = FilterDialog(headers, self.table_name, real_headers, self)
        if dialog.exec():
            filters = dialog.get_filters()
            self.apply_filters(filters, columns, real_headers)

    def apply_filters(self, filters, columns, real_headers):
        # 1) Спец-ветка для JOIN-модели ВЫДАЧА_ТЕХНИКИ — фильтруем через прокси
        if self.table_name in ["ВЫДАЧА_ТЕХНИКИ", "ОБСЛУЖИВАНИЕ"]:
            flts = []
            # собираем только непустые фильтры
            for value, col in zip(filters, columns):
                # tuple → диапазон дат
                if isinstance(value, tuple):
                    # (from_str, to_str), даже если один из них пустой
                    flts.append((value, col))
                # простая строка или число
                elif value.strip():
                    flts.append((value, col))
            # передаём в прокси
            if self.table_name == "ВЫДАЧА_ТЕХНИКИ":
                self.proxy.set_issuance_filters(flts)
            else:
                self.proxy.set_maintenance_filters(flts)
            return


        # 2) Базовая сборка SQL-фильтра для QSqlTableModel
        filter_strings = []
        for value, col, real_header in zip(filters, columns, real_headers):
            # диапазон дат
            if isinstance(value, tuple):
                date_from, date_to = value
                sql_from = self._to_sql_date(date_from) if date_from else ""
                sql_to   = self._to_sql_date(date_to)   if date_to   else ""
                if sql_from and sql_to:
                    filter_strings.append(
                        "(substr([{0}],7,4)||'-'||substr([{0}],4,2)||'-'||substr([{0}],1,2) >= '{1}' "
                        "AND substr([{0}],7,4)||'-'||substr([{0}],4,2)||'-'||substr([{0}],1,2) <= '{2}')"
                        .format(real_header, sql_from, sql_to)
                    )
                elif sql_from:
                    filter_strings.append(
                        "substr([{0}],7,4)||'-'||substr([{0}],4,2)||'-'||substr([{0}],1,2) >= '{1}'"
                        .format(real_header, sql_from)
                    )
                elif sql_to:
                    filter_strings.append(
                        "substr([{0}],7,4)||'-'||substr([{0}],4,2)||'-'||substr([{0}],1,2) <= '{1}'"
                        .format(real_header, sql_to)
                    )

            # простой текст или число
            elif value.strip():
                if value.isdigit():
                    filter_strings.append(f"[{real_header}] = '{value}'")
                else:
                    filter_strings.append(f"[{real_header}] LIKE '%{value}%'")

        filter_query = " AND ".join(filter_strings)

        # 3) Применяем к QSqlTableModel
        self.model.setFilter(filter_query)
        self.model.select()

    def _to_sql_date(self, date_str):
        # Преобразует dd.MM.yyyy в yyyy-MM-dd для SQLite
        d, m, y = date_str.split('.')
        return f"{y}-{m}-{d}"
    
    def reset_filters(self):
        """
        - для QSqlQueryModel (CostModel и QSqlQueryModel) — перезапускаем SQL и очищаем прокси-фильтры
        - для QSqlTableModel — очищаем setFilter и делаем select()
        """
        # 1) CostModel и QSqlQueryModel
        if isinstance(self.model, QSqlQueryModel) and self._sql:
            # возвращаем исходный JOIN-запрос
            self.model.setQuery(self._sql)
            # а потом сбрасываем фильтры прокси
            if self.table_name == "ОБСЛУЖИВАНИЕ":
                self.proxy.set_maintenance_filters([])
            else:
                self.proxy.set_issuance_filters([])
            return

        # 2) Обычные таблицы (QSqlTableModel)
        if hasattr(self.model, "setFilter"):
            self.model.setFilter("")
        if hasattr(self.model, "select"):
            self.model.select()

class IssuanceFilterProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.filters = []   # [(value_or_tuple, source_col), ...]

    def filterAcceptsRow(self, src_row, src_parent):
        model = self.sourceModel()
        for value, col in self.filters:
            idx = model.index(src_row, col, src_parent)
            cell = str(idx.data() or "")
            # диапазон дат — кортеж (dd.MM.yyyy, dd.MM.yyyy)
            if isinstance(value, tuple):
                fr, to = value
                if fr and to:
                    if not (fr <= cell <= to):
                        return False
            else:
                if value and value.lower() not in cell.lower():
                    return False
        return True

    def set_issuance_filters(self, flts):
        self.filters = flts
        self.invalidateFilter()

class AddRowDialog(QDialog):
    def __init__(self, headers, table_name, parent=None):
        super().__init__(parent)
        self.headers = headers  # <--- добавьте эту строку!
        self.setWindowTitle("Добавить запись")
        self.inputs = []
        layout = QFormLayout(self)

        if table_name == "ОБОРУДОВАНИЕ":
            for header in headers:
                if header.lower() == "тип":
                    combo = QComboBox(self)
                    combo.addItems(EQUIP_TYPES)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                elif header.lower() == "статус":
                    combo = QComboBox(self)
                    combo.addItems(EQUIP_STATUS)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                elif header.lower() == "дата покупки":
                    date_edit = QDateEdit(self)
                    date_edit.setDisplayFormat("dd.MM.yyyy")
                    date_edit.setCalendarPopup(True)
                    date_edit.setDate(QDate.currentDate())
                    layout.addRow(header, date_edit)
                    self.inputs.append(date_edit)
                else:
                    edit = QLineEdit(self)
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
        elif table_name == "ОБСЛУЖИВАНИЕ":
            for header in headers:
                header_str = str(header).lower()
                # Оборудование
                if "оборудование" in header_str:
                    combo = QComboBox(self)
                    from PyQt6.QtSql import QSqlQuery
                    query = QSqlQuery("SELECT id, название, производитель, серийный_номер FROM ОБОРУДОВАНИЕ")
                    self.equip_ids = []
                    while query.next():
                        eqid = query.value(0)
                        label = f"{query.value(1)} ({query.value(2)}, {query.value(3)})"
                        self.equip_ids.append(eqid)
                        combo.addItem(label)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                # Дата
                elif "дата" in header_str:
                    date_edit = QDateEdit(self)
                    date_edit.setDisplayFormat("dd.MM.yyyy")
                    date_edit.setCalendarPopup(True)
                    date_edit.setDate(QDate.currentDate())
                    layout.addRow(header, date_edit)
                    self.inputs.append(date_edit)
                # Тип работы
                elif "тип работы" in header_str:
                    combo = QComboBox(self)
                    combo.addItems(WORK_TYPES)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                # Техник - используем тот же список, что и для сотрудников
                elif "техник" in header_str:
                    combo = QComboBox(self)
                    from PyQt6.QtSql import QSqlQuery
                    query = QSqlQuery("SELECT id, фамилия, имя FROM СОТРУДНИК")
                    self.tech_ids = [None]  # Первое значение - None (не указано)
                    combo.addItem("Не указано")
                    while query.next():
                        tid = query.value(0)
                        label = f"{query.value(1)} {query.value(2)}"
                        self.tech_ids.append(tid)
                        combo.addItem(label)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                # Стоимость
                elif "стоимость" in header_str:
                    edit = QLineEdit(self)
                    edit.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d*\.?\d*$")))
                    edit.setPlaceholderText("0.00")
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
                else:
                    edit = QLineEdit(self)
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
        elif table_name == "РАБОЧЕЕ_МЕСТО":
            for header in headers:
                if header.lower() == "статус":
                    combo = QComboBox(self)
                    combo.addItems(WORKPLACE_STATUS)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                elif header.lower() == "этаж":
                    edit = QLineEdit(self)
                    edit.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,2}$")))
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
                else:
                    edit = QLineEdit(self)
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
        elif table_name == "ВЫДАЧА_ТЕХНИКИ":
            for header in headers:
                header_str = str(header).lower()
                # Сотрудник
                if "сотрудник" in header_str:
                    combo = QComboBox(self)
                    from PyQt6.QtSql import QSqlQuery
                    query = QSqlQuery("SELECT id, фамилия, имя, отчество FROM СОТРУДНИК")
                    self.employee_ids = []
                    while query.next():
                        eid = query.value(0)
                        patr = query.value(3)
                        label = f"{query.value(1)} {query.value(2)}"
                        if patr and patr.strip() != "-":
                            label += f" {patr}"
                        self.employee_ids.append(eid)
                        combo.addItem(label)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                # Оборудование
                elif "оборудование" in header_str:
                    combo = QComboBox(self)
                    from PyQt6.QtSql import QSqlQuery
                    query = QSqlQuery("SELECT id, название, производитель, серийный_номер FROM ОБОРУДОВАНИЕ")
                    self.equip_ids = []
                    while query.next():
                        eqid = query.value(0)
                        label = f"{query.value(1)} ({query.value(2)}, {query.value(3)})"
                        self.equip_ids.append(eqid)
                        combo.addItem(label)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                # Рабочее место (улучшено)
                elif "рабочее" in header_str and "место" in header_str:
                    combo = QComboBox(self)
                    from PyQt6.QtSql import QSqlQuery
                    query = QSqlQuery("SELECT id, корпус, этаж, кабинет, стол FROM РАБОЧЕЕ_МЕСТО")
                    self.workplace_ids = []
                    while query.next():
                        wid = query.value(0)
                        label = f"{query.value(1)} | этаж {query.value(2)} | каб. {query.value(3)} | стол {query.value(4)}"
                        self.workplace_ids.append(wid)
                        combo.addItem(label)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                # Даты
                elif "дата" in header_str:
                    date_edit = QDateEdit(self)
                    date_edit.setDisplayFormat("dd.MM.yyyy")
                    date_edit.setCalendarPopup(True)
                    if "выдачи" in header_str:
                        date_edit.setDate(QDate.currentDate())
                    layout.addRow(header, date_edit)
                    self.inputs.append(date_edit)
                # Состояние при возврате (улучшено)
                elif "состояние" in header_str and "возврат" in header_str:
                    combo = QComboBox(self)
                    combo.addItems([
                        "Исправно",
                        "Требует ремонта",
                        "Утеряно",
                        "Повреждено",
                        "Другое"
                    ])
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                else:
                    edit = QLineEdit(self)
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
        else:
            for header in headers:
                if header.lower() == "отдел":
                    self.department_combo = QComboBox(self)
                    self.department_combo.addItems(DEPARTMENTS)
                    layout.addRow(header, self.department_combo)
                    self.inputs.append(self.department_combo)
                elif header.lower() == "должность":
                    self.position_combo = QComboBox(self)
                    self.position_combo.addItems(POSITIONS_BY_DEPARTMENT[DEPARTMENTS[0]])
                    layout.addRow(header, self.position_combo)
                    self.inputs.append(self.position_combo)
                elif header.lower() == "рабочее место":
                    combo = QComboBox(self)
                    from PyQt6.QtSql import QSqlQuery
                    query = QSqlQuery("SELECT id, корпус, этаж, кабинет, стол FROM РАБОЧЕЕ_МЕСТО")
                    self.workplace_ids = []
                    while query.next():
                        wid = query.value(0)
                        label = f"{query.value(1)} | этаж {query.value(2)} | каб. {query.value(3)} | стол {query.value(4)}"
                        self.workplace_ids.append(wid)
                        combo.addItem(label)
                    if not self.workplace_ids:
                        combo.addItem("Нет")
                        self.workplace_ids.append(None)
                    layout.addRow(header, combo)
                    self.inputs.append(combo)
                elif header.lower() == "дата рождения":
                    date_edit = QDateEdit(self)
                    date_edit.setDisplayFormat("dd.MM.yyyy")
                    date_edit.setCalendarPopup(True)
                    date_edit.setDate(QDate.currentDate().addYears(-18))  # по умолчанию 18 лет назад
                    layout.addRow(header, date_edit)
                    self.inputs.append(date_edit)
                elif header.lower() == "номер телефона":
                    edit = QLineEdit(self)
                    # Разрешить только + и 11 цифр, максимум 12 символов (+79991234567)
                    regex = QRegularExpression(r"^\+?\d{0,11}$")
                    validator = QRegularExpressionValidator(regex)
                    edit.setValidator(validator)
                    edit.setMaxLength(12)
                    edit.setPlaceholderText("+79991234567")
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
                else:
                    edit = QLineEdit(self)
                    layout.addRow(header, edit)
                    self.inputs.append(edit)
            # Связываем изменение отдела с обновлением должностей
            if self.department_combo and self.position_combo:
                self.department_combo.currentIndexChanged.connect(self.update_positions)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def update_positions(self):
        department = self.department_combo.currentText()
        self.position_combo.clear()
        self.position_combo.addItems(POSITIONS_BY_DEPARTMENT.get(department, []))

    def get_data(self, table_name=None):
        if table_name == "ВЫДАЧА_ТЕХНИКИ":
            result = []
            for idx, inp in enumerate(self.inputs):
                header_str = str(self.headers[idx]).lower()
                if "сотрудник" in header_str:
                    value = self.employee_ids[inp.currentIndex()]
                elif "оборудование" in header_str:
                    value = self.equip_ids[inp.currentIndex()]
                elif "рабочее" in header_str and "место" in header_str:
                    value = self.workplace_ids[inp.currentIndex()]
                elif isinstance(inp, QDateEdit):
                    value = inp.date().toString("dd.MM.yyyy")
                elif isinstance(inp, QComboBox):
                    value = inp.currentText()
                else:
                    value = inp.text()
                result.append(value)
            return result
        elif table_name == "ОБСЛУЖИВАНИЕ":
            result = []
            for idx, inp in enumerate(self.inputs):
                header_str = str(self.headers[idx]).lower()
                if "оборудование" in header_str:
                    value = self.equip_ids[inp.currentIndex()]
                elif "техник" in header_str:
                    value = self.tech_ids[inp.currentIndex()]
                elif isinstance(inp, QDateEdit):
                    value = inp.date().toString("dd.MM.yyyy")
                elif isinstance(inp, QComboBox):
                    value = inp.currentText()
                elif "стоимость" in header_str:
                    value = inp.text() or "0"
                else:
                    value = inp.text()
                result.append(value)
            return result

        result = []
        for idx, inp in enumerate(self.inputs):
            if isinstance(inp, QComboBox):
                value = inp.currentText()
            elif isinstance(inp, QDateEdit):
                value = inp.date().toString("dd.MM.yyyy")
            else:
                value = inp.text()
            if table_name == "СОТРУДНИК":
                header_str = str(self.headers[idx]).lower()
                if "номер телефона" in header_str:
                    if value and not value.startswith("+"):
                        value = "+" + value
                elif "рабочее" in header_str and "место" in header_str:
                    if hasattr(self, "workplace_ids"):
                        value = self.workplace_ids[inp.currentIndex()]
            result.append(value)
        return result

    def validate(self):
        # Для сотрудников
        if hasattr(self, "department_combo"):
            surname = get_input_value(self.inputs[0])
            name = get_input_value(self.inputs[1])
            patronymic = get_input_value(self.inputs[2])
            # дата рождения = self.inputs[3] (QDateEdit)
            email = get_input_value(self.inputs[6])
            phone = get_input_value(self.inputs[7])

            # Фамилия, Имя: 2-33 символа, только буквы
            for value, label in [(surname, "Фамилия"), (name, "Имя")]:
                if not (2 <= len(value) <= 33) or not value.isalpha():
                    QMessageBox.warning(self, "Ошибка", f"{label} должна содержать только буквы и быть длиной от 2 до 33 символов.")
                    return False
            # Отчество: если не пустое и не прочерк, те же правила
            if patronymic and patronymic != "-" and (not (2 <= len(patronymic) <= 33) or not patronymic.isalpha()):
                QMessageBox.warning(self, "Ошибка", "Отчество должно содержать только буквы и быть длиной от 2 до 33 символов.")
                return False
            # Email: обязательно, валидный
            if not email or not is_valid_email(email):
                QMessageBox.warning(self, "Ошибка", "Введите корректный email (например, user@mail.ru).")
                return False
            # Телефон: обязательно, российский формат
            if not phone or not is_valid_phone(phone):
                QMessageBox.warning(self, "Ошибка", "Введите корректный российский номер телефона (например, +79991234567 или 89991234567).")
                return False
        # Для оборудования
        if hasattr(self, "inputs") and self.parent().table_name == "ОБОРУДОВАНИЕ":
            name = get_input_value(self.inputs[0])
            inv = get_input_value(self.inputs[1])
            equip_type = get_input_value(self.inputs[2])
            manufacturer = get_input_value(self.inputs[3])
            model = get_input_value(self.inputs[4])
            serial = get_input_value(self.inputs[5])

            if not (2 <= len(name) <= 50):
                QMessageBox.warning(self, "Ошибка", "Название должно быть от 2 до 50 символов.")
                return False
            if not inv or len(inv) > 20:
                QMessageBox.warning(self, "Ошибка", "Инвентарный номер обязателен и до 20 символов.")
                return False
            if not equip_type:
                QMessageBox.warning(self, "Ошибка", "Выберите тип оборудования.")
                return False
            if len(manufacturer) > 50:
                QMessageBox.warning(self, "Ошибка", "Производитель — до 50 символов.")
                return False
            if len(model) > 50:
                QMessageBox.warning(self, "Ошибка", "Модель — до 50 символов.")
                return False
            if len(serial) > 30:
                QMessageBox.warning(self, "Ошибка", "Серийный номер — до 30 символов.")
                return False
        # Для рабочих мест
        if self.parent().table_name == "РАБОЧЕЕ_МЕСТО":
            name = get_input_value(self.inputs[0])
            room = get_input_value(self.inputs[1])
            floor = get_input_value(self.inputs[2])
            building = get_input_value(self.inputs[3])
            # статус = self.inputs[4] (QComboBox)
            if not (2 <= len(name) <= 50):
                QMessageBox.warning(self, "Ошибка", "Название должно быть от 2 до 50 символов.")
                return False

            if floor and not floor.isdigit():
                QMessageBox.warning(self, "Ошибка", "Этаж должен быть числом.")
                return False
            # Можно добавить другие проверки по желанию
        return True

    def accept(self):
        if self.validate():
            super().accept()

class EditRowDialog(AddRowDialog):
    def __init__(self, headers, values, table_name, parent=None):
        super().__init__(headers, table_name, parent)
        for idx, (inp, val) in enumerate(zip(self.inputs, values)):
            if isinstance(inp, QLineEdit):
                if table_name == "СОТРУДНИК":
                    header = headers[idx].lower()
                    if "номер телефона" in header:
                        val_str = str(val)
                        if val_str and not val_str.startswith("+"):
                            val_str = "+" + val_str
                        inp.setText(val_str)
                        continue
                inp.setText(str(val))
            elif isinstance(inp, QComboBox):
                header = headers[idx].lower()
                if table_name == "ВЫДАЧА_ТЕХНИКИ":
                    if "сотрудник" in header and hasattr(self, "employee_ids"):
                        if val in self.employee_ids:
                            inp.setCurrentIndex(self.employee_ids.index(val))
                    elif "оборудование" in header and hasattr(self, "equip_ids"):
                        if val in self.equip_ids:
                            inp.setCurrentIndex(self.equip_ids.index(val))
                    elif "рабочее_место" in header and hasattr(self, "workplace_ids"):
                        if val in self.workplace_ids:
                            inp.setCurrentIndex(self.workplace_ids.index(val))
                    else:
                        idx_combo = inp.findText(str(val))
                        if idx_combo >= 0:
                            inp.setCurrentIndex(idx_combo)
                elif table_name == "СОТРУДНИК" and "рабочее_место" in header and hasattr(self, "workplace_ids"):
                    if val in self.workplace_ids:
                        inp.setCurrentIndex(self.workplace_ids.index(val))
                    else:
                        inp.setCurrentIndex(0)
                elif table_name == "ОБСЛУЖИВАНИЕ":
                    if "оборудование" in header and hasattr(self, "equip_ids"):
                        if val in self.equip_ids:
                            inp.setCurrentIndex(self.equip_ids.index(val))
                    elif "техник" in header and hasattr(self, "tech_ids"):
                        if val in self.tech_ids:
                            inp.setCurrentIndex(self.tech_ids.index(val))
                        else:
                            inp.setCurrentIndex(0)
                    else:
                        idx_combo = inp.findText(str(val))
                        if idx_combo >= 0:
                            inp.setCurrentIndex(idx_combo)
                else:
                    idx_combo = inp.findText(str(val))
                    if idx_combo >= 0:
                        inp.setCurrentIndex(idx_combo)
            elif isinstance(inp, QDateEdit):
                try:
                    date = QDate.fromString(str(val), "dd.MM.yyyy")
                    if date.isValid():
                        inp.setDate(date)
                except Exception:
                    pass

class FilterDialog(QDialog):
    def __init__(self, headers, table_name, real_headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Фильтрация")
        self.inputs = []
        self.headers = headers
        self.real_headers = real_headers
        self.table_name = table_name

        layout = QFormLayout(self)

        # Собираем уникальные значения для выпадающих списков
        self.values_by_header = {}
        from PyQt6.QtSql import QSqlQuery, QSqlQueryModel

        # Если это окно фильтрации для JOIN-модели ВЫДАЧА_ТЕХНИКИ — берём данные из модели
        if self.table_name in ("ВЫДАЧА_ТЕХНИКИ", "ОБСЛУЖИВАНИЕ") \
           and isinstance(parent.model, QSqlQueryModel):
            model = parent.model
            rec  = model.record()  # по нему найдём поле
            for idx, header in enumerate(headers):
                field = real_headers[idx]             # имя поля из SELECT
                col   = rec.indexOf(field)           # его реальный индекс
                vals  = set()
                for row in range(model.rowCount()):
                    cell = model.index(row, col).data()
                    if cell not in (None, ""):
                        vals.add(str(cell))
                self.values_by_header[header] = sorted(vals)

        else:
            # Всё остальное — как у вас было, через SELECT DISTINCT
            for header, real_header in zip(headers, real_headers):
                query = QSqlQuery(f"SELECT DISTINCT [{real_header}] FROM {table_name}")
                vals = []
                while query.next():
                    v = query.value(0)
                    if v not in (None, ""):
                        vals.append(str(v))
                self.values_by_header[header] = sorted(vals)

        # Строим поля фильтрации
        for i, header in enumerate(headers):
            # Спец. диапазон дат для сотрудников
            if table_name == "СОТРУДНИК" and header.lower() == "дата рождения":
                hbox = QHBoxLayout()
                date_from = QDateEdit(self)
                date_from.setDisplayFormat("dd.MM.yyyy")
                date_from.setCalendarPopup(True)
                date_from.setDate(QDate(1950, 1, 1))
                date_to = QDateEdit(self)
                date_to.setDisplayFormat("dd.MM.yyyy")
                date_to.setCalendarPopup(True)
                date_to.setDate(QDate.currentDate())
                hbox.addWidget(date_from)
                hbox.addWidget(QLabel("—"))
                hbox.addWidget(date_to)
                layout.addRow(header, hbox)
                self.inputs.append((date_from, date_to))
            else:
                combo = QComboBox(self)
                combo.addItem("Любое значение")
                combo.addItems(self.values_by_header.get(header, []))
                combo.currentIndexChanged.connect(self.update_cascades)
                layout.addRow(header, combo)
                self.inputs.append(combo)

        # Кнопки ОК/Отмена
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def update_cascades(self):
        # для JOIN-моделей (QueryModel) — ни в коем случае не чистим комбо
        if (self.table_name in ("ВЫДАЧА_ТЕХНИКИ", "ОБСЛУЖИВАНИЕ")
                and isinstance(self.parent().model, QSqlQueryModel)):
            return

        # остальной ваш код каскадов для остальных таблиц
        selected = {}
        for i, inp in enumerate(self.inputs):
            if isinstance(inp, tuple):
                continue
            txt = inp.currentText()
            if txt != "Любое значение":
                selected[self.headers[i]] = txt

        from PyQt6.QtSql import QSqlQuery
        for i, (header, real_header) in enumerate(zip(self.headers, self.real_headers)):
            if isinstance(self.inputs[i], tuple):
                continue

            where_clauses = []
            for j in range(i):
                if isinstance(self.inputs[j], tuple):
                    continue
                sel = self.inputs[j].currentText()
                if sel != "Любое значение":
                    where_clauses.append(f"[{self.real_headers[j]}]='{sel}'")
            where_sql = " AND ".join(where_clauses)
            sql = f"SELECT DISTINCT [{real_header}] FROM {self.table_name}"
            if where_sql:
                sql += " WHERE " + where_sql

            qry = QSqlQuery(sql)
            vals = [str(qry.value(0)) for _ in iter(lambda: qry.next(), False) if qry.value(0) not in (None, "")]
            
            combo = self.inputs[i]
            cur   = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("Любое значение")
            combo.addItems(sorted(vals))
            idx = combo.findText(cur)
            if idx >= 0:
                combo.setCurrentIndex(idx)
            combo.blockSignals(False)

    def get_filters(self):
        result = []
        for inp in self.inputs:
            if isinstance(inp, tuple):  # диапазон дат
                date_from = inp[0].date().toString("dd.MM.yyyy")
                date_to   = inp[1].date().toString("dd.MM.yyyy")
                result.append((date_from, date_to))
            else:
                val = inp.currentText()
                # Если ничего не выбрано, используем пустую строку
                result.append("" if val == "Любое значение" else val)
        return result
    
def get_input_value(inp):
    if isinstance(inp, QComboBox):
        return inp.currentText()
    elif isinstance(inp, QDateEdit):
        return inp.date().toString("dd.MM.yyyy")
    elif hasattr(inp, "text"):
        return inp.text()
    else:
        return ""
