"""Настольный интерфейс. Сетевые операции выполняются вне потока Qt."""
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from PySide6.QtCore import QObject, QRunnable, QThreadPool, QTimer, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog,
    QFormLayout, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QPushButton,
    QSpinBox, QSplitter, QTabWidget, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget,
)

from .domain import Status
from .excel import read_table, sheet_names
from .export import preview, write_xlsx
from .google_source import GoogleSource, new_account, spreadsheet_id
from .service import Service
from .theme import apply_theme
from . import __version__


class WorkerSignals(QObject):
    done = Signal(object)
    failed = Signal(str)


class Worker(QRunnable):
    def __init__(self, fn):
        super().__init__()
        self.fn = fn
        self.signals = WorkerSignals()

    def run(self):
        try:
            result = self.fn()
        except Exception as exc:
            self.signals.failed.emit(str(exc))
        else:
            self.signals.done.emit(result)


def button(text, action, layout):
    widget = QPushButton(text)
    widget.clicked.connect(action)
    layout.addWidget(widget)
    return widget


def table(headers):
    widget = QTableWidget(0, len(headers))
    widget.setHorizontalHeaderLabels(headers)
    widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    widget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    widget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
    widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    widget.horizontalHeader().setStretchLastSection(True)
    widget.verticalHeader().hide()
    widget.setAlternatingRowColors(True)
    return widget


def fill(widget, rows):
    widget.setRowCount(len(rows))
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            widget.setItem(r, c, QTableWidgetItem("" if value is None else str(value)))


def page(tabs, name):
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(18, 18, 18, 18)
    tabs.addTab(widget, name)
    return layout


class ExportDialog(QDialog):
    def __init__(self, snapshot, parent):
        super().__init__(parent)
        self.setWindowTitle("Предварительный просмотр экспорта — локальное состояние")
        self.resize(1000, 650)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Экспорт использует показанный снимок. Google повторно не читается."))
        tabs = QTabWidget()
        layout.addWidget(tabs)
        for sheet in snapshot["sheets"]:
            host = QWidget()
            pane = QVBoxLayout(host)
            pane.addWidget(QLabel("Последнее успешное чтение (UTC): " + (sheet["last_read"] or "не выполнялось")))
            grid = table(sheet["headers"])
            fill(grid, sheet["rows"])
            pane.addWidget(grid)
            tabs.addTab(host, f"{sheet['name']} ({len(sheet['rows'])})")
        notes = QTextEdit()
        notes.setReadOnly(True)
        notes.setPlainText("\n".join(snapshot["warnings"]) or "Предупреждений нет")
        notes.setMaximumHeight(160)
        layout.addWidget(notes)
        controls = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        controls.button(QDialogButtonBox.StandardButton.Save).setText("Сохранить XLSX")
        controls.button(QDialogButtonBox.StandardButton.Cancel).setText("Отмена")
        controls.accepted.connect(self.accept)
        controls.rejected.connect(self.reject)
        layout.addWidget(controls)


class MainWindow(QMainWindow):
    def __init__(self, root):
        super().__init__()
        apply_theme(QApplication.instance())
        self.service = Service(root)
        self.service.store.recover()
        self.root = Path(root)
        self.pool = QThreadPool(self)
        self.pool.setMaxThreadCount(3)
        self.busy = set()
        self.workers = {}
        self.next_reads = {}
        self.records = []
        self.setWindowTitle(f"Сверка УИК {__version__} • Excel и Google Таблицы")
        self.resize(1280, 840)
        self.setMinimumSize(980, 650)
        self._build()
        self.load_projects()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)

    @property
    def project_id(self):
        return self.projects.currentData()

    @property
    def pair_id(self):
        return self.pairs.currentData()

    def _build(self):
        host = QWidget()
        layout = QVBoxLayout(host)
        self.setCentralWidget(host)
        top = QHBoxLayout()
        top.addWidget(QLabel("Проект"))
        self.projects = QComboBox()
        self.projects.setMinimumWidth(220)
        top.addWidget(self.projects)
        button("Новый проект", self.new_project, top)
        top.addSpacing(20)
        top.addWidget(QLabel("Рабочая пара"))
        self.pairs = QComboBox()
        top.addWidget(self.pairs)
        top.addStretch()
        self.check_now = button("Проверить сейчас", self.refresh_current, top)
        layout.addLayout(top)
        self.operation = QLabel("Готово к работе")
        self.operation.setWordWrap(True)
        layout.addWidget(self.operation)
        self.project_hint = QLabel("Чтобы подключить таблицы, нажмите «Новый проект» и задайте название. Все разделы можно открыть уже сейчас.")
        self.project_hint.setObjectName("projectHint")
        self.project_hint.setWordWrap(True)
        layout.addWidget(self.project_hint)
        self.tabs = QTabWidget()
        self.tabs.setUsesScrollButtons(True)
        layout.addWidget(self.tabs)
        self._summary_page()
        self._sources_page()
        self._mapping_page()
        self._checks_page()
        self._history_page()
        self._confirmations_page()
        self._export_page()
        self._settings_page()
        self._log_page()
        self.projects.currentIndexChanged.connect(self.project_changed)
        self.pairs.currentIndexChanged.connect(self.pair_changed)
        self.statusBar().showMessage(f"Рабочая папка: {self.root}")

    def _summary_page(self):
        layout = page(self.tabs, "Главная сводка")
        filters = QHBoxLayout()
        self.uik_filter = QLineEdit()
        self.uik_filter.setPlaceholderText("Поиск по УИК")
        filters.addWidget(self.uik_filter)
        self.pair_filter = QComboBox()
        self.pair_filter.addItem("Все пары", None)
        for i in range(1, 4):
            self.pair_filter.addItem(f"Пара {i}", i)
        filters.addWidget(self.pair_filter)
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Все результаты", *[s.value for s in Status]])
        filters.addWidget(self.status_filter)
        self.confirm_filter = QComboBox()
        self.confirm_filter.addItems(["Все подтверждения", "Подтверждено", "Не подтверждено"])
        filters.addWidget(self.confirm_filter)
        layout.addLayout(filters)
        self.summary_grid = table(["УИК", "Пара 1", "Пара 2", "Пара 3", "Итог"])
        self.summary_grid.cellDoubleClicked.connect(self.open_uik)
        layout.addWidget(self.summary_grid)
        layout.addWidget(QLabel("Дважды нажмите строку, чтобы открыть историю выбранной УИК."))
        self.uik_filter.textChanged.connect(self.show_summary)
        for combo in (self.pair_filter, self.status_filter, self.confirm_filter):
            combo.currentIndexChanged.connect(self.show_summary)

    def _sources_page(self):
        layout = page(self.tabs, "Источники данных")
        form = QFormLayout()
        self.active = QCheckBox("Пара используется в сводке и экспорте")
        form.addRow("Активность", self.active)
        self.excel_info = QLabel("Excel отсутствует")
        self.excel_info.setWordWrap(True)
        form.addRow("Excel", self.excel_info)
        self.document = QLineEdit()
        self.document.setPlaceholderText("https://docs.google.com/spreadsheets/d/…")
        form.addRow("Google Таблица", self.document)
        self.google_sheet = QComboBox()
        self.google_sheet.setEditable(True)
        form.addRow("Лист Google", self.google_sheet)
        layout.addLayout(form)
        row = QHBoxLayout()
        button("Импортировать Excel…", self.import_excel, row)
        button("Сохранить источник", self.save_source, row)
        button("Войти в Google…", self.authorize, row)
        button("Получить листы", self.load_sheets, row)
        button("Прочитать структуру", self.discover, row)
        layout.addLayout(row)
        self.source_info = QTextEdit()
        self.source_info.setReadOnly(True)
        layout.addWidget(self.source_info)
        hint = QLabel("Первое подключение: сохраните ссылку и название листа → войдите в Google → прочитайте структуру → настройте поля → «Проверить сейчас».\nExcel необязателен. Первая строка каждого листа должна содержать уникальные непустые заголовки.")
        hint.setWordWrap(True)
        layout.addWidget(hint)

    def _mapping_page(self):
        layout = page(self.tabs, "Сопоставление полей")
        form = QFormLayout()
        self.google_uik = QComboBox()
        self.id_field = QComboBox()
        self.time_field = QComboBox()
        form.addRow("Номер УИК в Google", self.google_uik)
        form.addRow("Уникальный ID отправки", self.id_field)
        form.addRow("Время отправки", self.time_field)
        self.time_format = QComboBox()
        self.time_format.setEditable(True)
        self.time_format.addItems(["%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "ISO"])
        form.addRow("Формат времени", self.time_format)
        self.ties_by_row = QCheckBox("При одинаковом времени использовать порядок строк Google")
        form.addRow("Равное время", self.ties_by_row)
        layout.addLayout(form)
        self.mapping_grid = table(["Поле Excel", "Поле Google (пусто — хранить отдельно)"])
        layout.addWidget(self.mapping_grid)
        self.mapping_info = QLabel()
        self.mapping_info.setWordWrap(True)
        layout.addWidget(self.mapping_info)
        button("Сохранить схему и сверить локальные данные", self.save_mapping, layout)

    def _checks_page(self):
        layout = page(self.tabs, "Проверка данных")
        button("Сверить сохранённые данные без чтения Google", self.recheck, layout)
        self.checks_grid = table(["УИК", "ID", "Результат", "Различия / причина"])
        layout.addWidget(self.checks_grid)

    def _history_page(self):
        layout = page(self.tabs, "История отправок")
        self.history_filter = QLineEdit()
        self.history_filter.setPlaceholderText("Фильтр УИК (пусто — все)")
        self.history_filter.textChanged.connect(self.show_history)
        layout.addWidget(self.history_filter)
        self.history_grid = table(["УИК", "ID", "Версия", "Время отправки", "Результат", "Подтверждение", "Источник / версия"])
        self.history_grid.itemSelectionChanged.connect(self.show_details)
        self.details = QTextEdit()
        self.details.setReadOnly(True)
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.addWidget(self.history_grid)
        splitter.addWidget(self.details)
        splitter.setSizes([400, 220])
        layout.addWidget(splitter)

    def _confirmations_page(self):
        layout = page(self.tabs, "Подтверждения")
        layout.addWidget(QLabel("Выберите актуальную отправку. Для расхождений и отсутствующего Excel требуется отдельное согласие."))
        self.confirm_grid = table(["УИК", "ID", "Время отправки", "Результат", "Подтверждено (UTC)"])
        layout.addWidget(self.confirm_grid)
        controls = QHBoxLayout()
        button("Показать подробности", self.confirm_details, controls)
        button("Подтвердить выбранную", self.confirm_selected, controls)
        button("Снять подтверждение", self.revoke_selected, controls)
        layout.addLayout(controls)

    def _export_page(self):
        layout = page(self.tabs, "Экспорт")
        layout.addWidget(QLabel("Уникальные поля Excel текущей пары, включаемые в экспорт:"))
        self.export_fields = QListWidget()
        layout.addWidget(self.export_fields)
        button("Сохранить выбор полей Excel", self.save_export, layout)
        layout.addWidget(QLabel("Общие поля и все уникальные поля Google включаются автоматически. Каждая активная пара — отдельный лист XLSX."))
        button("Предпросмотр и экспорт всего проекта…", self.export_project, layout)

    def _settings_page(self):
        layout = page(self.tabs, "Настройки")
        self.auto = QCheckBox("Автоматически читать текущую пару Google")
        layout.addWidget(self.auto)
        form = QFormLayout()
        self.interval = QSpinBox()
        self.interval.setRange(10, 86400)
        self.interval.setSuffix(" сек.")
        form.addRow("Интервал", self.interval)
        layout.addLayout(form)
        button("Сохранить настройки таймера", self.save_timer, layout)
        self.timer_info = QLabel()
        self.timer_info.setWordWrap(True)
        layout.addWidget(self.timer_info)
        path = QLabel(f"Рабочие данные: {self.root}\nДля переноса закройте программу и скопируйте папку приложения вместе с data.\nТокены Google хранятся в защищённом хранилище ОС; на другом компьютере нужен повторный вход.")
        path.setWordWrap(True)
        layout.addWidget(path)
        layout.addStretch()

    def _log_page(self):
        layout = page(self.tabs, "Журнал операций")
        button("Обновить журнал", self.show_log, layout)
        self.log_grid = table(["Время (UTC)", "Пара", "Событие", "Описание"])
        layout.addWidget(self.log_grid)

    def error(self, text):
        QMessageBox.warning(self, "Не удалось выполнить операцию", str(text))

    def guard(self, fn):
        try:
            return fn()
        except Exception as exc:
            self.error(str(exc))
            return None

    def load_projects(self, selected=None):
        self.projects.blockSignals(True)
        self.projects.clear()
        for row in self.service.store.query("SELECT * FROM projects ORDER BY id"):
            self.projects.addItem(row["name"], row["id"])
        if selected is not None:
            self.projects.setCurrentIndex(self.projects.findData(selected))
        self.projects.blockSignals(False)
        self.project_changed()

    def new_project(self):
        name, ok = QInputDialog.getText(self, "Новый проект", "Название проекта:")
        if ok and name.strip():
            pid = self.guard(lambda: self.service.store.create_project(name))
            if pid:
                self.load_projects(pid)

    def project_changed(self):
        self.pairs.blockSignals(True)
        self.pairs.clear()
        if self.project_id:
            for row in self.service.store.query("SELECT id,slot,active FROM pairs WHERE project_id=? ORDER BY slot", (self.project_id,)):
                self.pairs.addItem(f"Пара {row['slot']}" + ("" if row["active"] else " (выключена)"), row["id"])
        self.pairs.blockSignals(False)
        # Навигация доступна и до создания проекта. Ограничиваются только действия с данными.
        self.project_hint.setVisible(not bool(self.pair_id))
        self.pairs.setEnabled(bool(self.pair_id))
        for kind in (QPushButton, QLineEdit, QComboBox, QCheckBox, QSpinBox):
            for control in self.tabs.findChildren(kind):
                control.setEnabled(bool(self.pair_id))
        self.check_now.setEnabled(bool(self.pair_id))
        self.pair_changed()

    def pair_changed(self):
        if not self.pair_id:
            self.show_summary()
            return
        pair = self.service.store.pair(self.pair_id)
        config = pair["config"]
        self.active.setChecked(bool(pair["active"]))
        self.document.setText(config.get("document", ""))
        self.google_sheet.clear()
        self.google_sheet.addItems(config.get("sheets", []))
        self.google_sheet.setCurrentText(config.get("google_sheet", ""))
        self.excel_info.setText(f"{config.get('excel_status', 'Отсутствует')}; лист: {config.get('excel_sheet', '—')}; строк: {len(config.get('excel_rows', []))}\nВнутренняя копия: {config.get('excel_file', '—')}")
        self.populate_mapping(config)
        self.auto.setChecked(config.get("auto", False))
        self.interval.setValue(config.get("interval", 60))
        self.load_views()

    def populate_mapping(self, config):
        headers = config.get("google_headers", [])
        for widget, key in ((self.google_uik, "google_uik"), (self.id_field, "id_field"), (self.time_field, "time_field")):
            widget.clear()
            widget.addItem("— выберите поле —", "")
            for h in headers:
                widget.addItem(h, h)
            index = widget.findData(config.get(key, ""))
            widget.setCurrentIndex(max(0, index))
        self.time_format.setCurrentText(config.get("time_format", "%d.%m.%Y %H:%M:%S"))
        self.ties_by_row.setChecked(config.get("ties_by_row", False))
        mapping = config.get("mapping", {})
        excel_headers = config.get("excel_headers", [])
        self.mapping_grid.setRowCount(len(excel_headers))
        for i, e in enumerate(excel_headers):
            self.mapping_grid.setItem(i, 0, QTableWidgetItem(e))
            combo = QComboBox()
            combo.addItem("— хранить отдельно —", "")
            for h in headers:
                combo.addItem(h, h)
            if e in mapping:
                idx = combo.findData(mapping[e])
                if idx < 0:
                    combo.addItem(mapping[e] + " (поле исчезло)", mapping[e])
                    idx = combo.count() - 1
                combo.setCurrentIndex(idx)
            self.mapping_grid.setCellWidget(i, 1, combo)
        change = config.get("schema_change", {})
        self.mapping_info.setText("Схема не подбирается автоматически. Поля Google без пары сохраняются отдельно.\n"
                                  + ("Изменение структуры: " + json.dumps(change, ensure_ascii=False) if change else ""))
        self.export_fields.clear()
        for h in excel_headers:
            if h not in mapping:
                item = QListWidgetItem(h)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if h in config.get("export_excel", []) else Qt.CheckState.Unchecked)
                self.export_fields.addItem(item)

    def load_views(self):
        if not self.pair_id:
            return
        self.records = self.service.history(self.pair_id)
        config = self.service.store.pair(self.pair_id)["config"]
        self.source_info.setPlainText(
            "Последнее успешное чтение (UTC): " + config.get("last_read", "не выполнялось")
            + "\nПоследняя ошибка: " + (config.get("last_error") or "нет")
            + "\nСтолбцы Google: " + ", ".join(config.get("google_headers", []))
            + f"\nСохранено версий отправок: {len(self.records)}"
        )
        current = [r for r in self.records if r["current"]]
        fill(self.checks_grid, [[r["uik"], r["source_key"], r["check_result"]["status"],
                               json.dumps(r["check_result"].get("differences") or r["check_result"].get("reason", ""), ensure_ascii=False)] for r in current])
        self.confirm_records = current
        fill(self.confirm_grid, [[r["uik"], r["source_key"], r["source_time"], r["check_result"]["status"], r["confirmed_at"] or "Нет"] for r in current])
        self.show_history()
        self.show_summary()
        self.show_log()
        self.update_timer_label()

    def show_summary(self):
        if not self.project_id:
            self.summary_grid.setRowCount(0)
            return
        try:
            pairs, summary = self.service.summary(self.project_id)
            rows = []
            scope = self.pair_filter.currentData()
            for item in summary:
                if self.uik_filter.text().strip() not in item["uik"]:
                    continue
                relevant = [item["pairs"][p["id"]] for p in pairs if scope is None or p["slot"] == scope]
                if not relevant:
                    continue
                wanted = self.status_filter.currentText()
                if wanted != "Все результаты" and not any(c["status"] == wanted or (wanted == Status.READ_ERROR and c.get("last_error")) for c in relevant):
                    continue
                confirmed = self.confirm_filter.currentIndex()
                if confirmed == 1 and not any(c["confirmed"] for c in relevant):
                    continue
                if confirmed == 2 and not any(not c["confirmed"] for c in relevant):
                    continue
                row = [item["uik"], "Выключена", "Выключена", "Выключена", item["total"]]
                for p in pairs:
                    cell = item["pairs"][p["id"]]
                    text = ("Подтверждено · " if cell["confirmed"] else "") + cell["status"]
                    if cell["warnings"]:
                        text += " ⚠"
                    if cell.get("last_error"):
                        text += " · Ошибка чтения"
                    row[p["slot"]] = text
                rows.append(row)
            fill(self.summary_grid, rows)
            for r, row in enumerate(rows):
                self.summary_grid.item(r, 4).setForeground(QColor("#16724b" if row[4] == "Готово" else "#9b5814"))
        except Exception as exc:
            self.operation.setText("Не удалось построить сводку: " + str(exc))

    def open_uik(self, row, column):
        if column in (1, 2, 3):
            self.pairs.setCurrentIndex(column - 1)
        self.history_filter.setText(self.summary_grid.item(row, 0).text())
        self.tabs.setCurrentIndex(4)

    def show_history(self):
        wanted = self.history_filter.text().strip()
        self.history_records = [r for r in self.records if wanted in r["uik"]]
        fill(self.history_grid, [[r["uik"], r["source_key"], r["revision"], r["source_time"] or "Нет в источнике",
                                 r["check_result"]["status"], r["confirmed_at"] or "Нет",
                                 ("Актуальная" if r["current"] else "Предыдущая версия") + ("" if r["present"] else " / отсутствует в снимке")]
                                for r in self.history_records])
        self.details.clear()

    def show_details(self):
        row = self.history_grid.currentRow()
        if 0 <= row < len(self.history_records):
            r = self.history_records[row]
            config = self.service.store.pair(r["pair_id"])["config"]
            data = {"Отправка Google (исходные значения)": r["payload"],
                    "Excel": self.service.excel_index(config).get(r["uik"]),
                    "Проверка": r["check_result"], "Время чтения (UTC)": r["observed_at"],
                    "Отпечаток": r["fingerprint"]}
            self.details.setPlainText(json.dumps(data, ensure_ascii=False, indent=2))

    def confirm_details(self):
        index = self.confirm_grid.currentRow()
        if index < 0:
            return
        record = self.confirm_records[index]
        self.history_filter.setText(record["uik"])
        self.tabs.setCurrentIndex(4)
        for i, r in enumerate(self.history_records):
            if r["id"] == record["id"]:
                self.history_grid.selectRow(i)
                self.show_details()
                break

    def confirm_selected(self):
        index = self.confirm_grid.currentRow()
        if index < 0:
            return self.error("Выберите отправку")
        record = self.confirm_records[index]
        problem = record["check_result"]["status"] != Status.MATCH
        if problem and QMessageBox.question(self, "Подтверждение с предупреждением",
                f"УИК {record['uik']}, ID {record['source_key']}: {record['check_result']['status']}.\nПодтвердить эту отправку вручную для экспорта?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        self.guard(lambda: self.service.confirm(record["id"], allow_problem=problem))
        self.load_views()

    def revoke_selected(self):
        index = self.confirm_grid.currentRow()
        if index >= 0:
            self.guard(lambda: self.service.revoke(self.confirm_records[index]["id"]))
            self.load_views()

    def ensure_idle(self):
        if self.pair_id in self.busy:
            self.error("Дождитесь завершения операции текущей пары")
            return False
        return bool(self.pair_id)

    def save_source(self):
        if not self.ensure_idle():
            return False
        try:
            value = self.document.text().strip()
            if value:
                value = spreadsheet_id(value)
            self.service.configure(self.pair_id, {"document": value, "google_sheet": self.google_sheet.currentText().strip()}, self.active.isChecked())
            self.load_views()
            return True
        except Exception as exc:
            self.error(str(exc))
            return False

    def import_excel(self):
        if not self.ensure_idle():
            return
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel", "", "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        def perform():
            names = sheet_names(path)
            sheet, ok = QInputDialog.getItem(self, "Лист Excel", "Выберите лист:", names, 0, False)
            if not ok:
                return
            headers, rows = read_table(path, sheet)
            uik = ""
            if headers:
                uik, ok = QInputDialog.getItem(self, "Номер УИК", "Столбец УИК Excel:", headers, 0, False)
                if not ok:
                    return
            pid = self.pair_id
            self.run_job(pid, lambda: self.service.import_excel(pid, path, sheet, uik), lambda _: self.pair_changed())
        self.guard(perform)

    def account(self, pid):
        config = self.service.store.pair(pid)["config"]
        if not config.get("account"):
            self.service.configure(pid, {"account": new_account()})
        return self.service.store.pair(pid)["config"]["account"]

    def authorize(self):
        if not self.ensure_idle():
            return
        path, _ = QFileDialog.getOpenFileName(self, "OAuth client JSON (тип Desktop app)", "", "JSON (*.json)")
        if path:
            pid = self.pair_id
            account = self.account(pid)
            self.run_job(pid, lambda: GoogleSource(self.root, account).authorize(path), lambda _: self.operation.setText("Вход Google выполнен"))

    def load_sheets(self):
        if not self.save_source():
            return
        pid = self.pair_id
        config = self.service.store.pair(pid)["config"]
        source = GoogleSource(self.root, self.account(pid))
        def done(names):
            self.service.configure(pid, {"sheets": names})
            if self.pair_id == pid:
                old = self.google_sheet.currentText()
                self.google_sheet.clear()
                self.google_sheet.addItems(names)
                if old:
                    self.google_sheet.setCurrentText(old)
        self.run_job(pid, lambda: source.sheets(config.get("document", "")), done)

    def discover(self):
        if not self.save_source():
            return
        pid = self.pair_id
        config = self.service.store.pair(pid)["config"]
        source = GoogleSource(self.root, self.account(pid))
        def done(result):
            headers, _ = result
            # Структура не подменяет сохранённый снимок и его last_read.
            old = config.get("google_headers", [])
            change = {"added": [h for h in headers if h not in old], "removed": [h for h in old if h not in headers]}
            self.service.configure(pid, {"google_headers": headers, "schema_change": change})
            if self.pair_id == pid:
                self.pair_changed()
                self.tabs.setCurrentIndex(2)
        self.run_job(pid, lambda: source.read(config.get("document", ""), config.get("google_sheet", "")), done)

    def save_mapping(self):
        if not self.ensure_idle():
            return
        def perform():
            mapping = {}
            for i in range(self.mapping_grid.rowCount()):
                e = self.mapping_grid.item(i, 0).text()
                g = self.mapping_grid.cellWidget(i, 1).currentData()
                if g:
                    mapping[e] = g
            if len(set(mapping.values())) != len(mapping):
                raise ValueError("Один столбец Google нельзя сопоставить нескольким полям Excel")
            fields = [w.currentData() for w in (self.google_uik, self.id_field, self.time_field)]
            if not all(fields):
                raise ValueError("Выберите поля УИК, ID и времени отправки Google")
            self.service.configure(self.pair_id, {"mapping": mapping, "google_uik": fields[0], "id_field": fields[1],
                "time_field": fields[2], "time_format": self.time_format.currentText(), "ties_by_row": self.ties_by_row.isChecked()})
            self.pair_changed()
        self.guard(perform)

    def run_job(self, pid, fn, on_done=None, quiet=False):
        if pid in self.busy:
            return
        self.busy.add(pid)
        self.operation.setText(f"Выполняется операция пары {self.service.store.pair(pid)['slot']}…")
        worker = Worker(fn)
        self.workers[pid] = worker
        def finish(result=None, error=None):
            self.busy.discard(pid)
            self.workers.pop(pid, None)
            config = self.service.store.pair(pid)["config"]
            self.next_reads[pid] = datetime.now(timezone.utc) + timedelta(seconds=config.get("interval", 60))
            if error:
                self.operation.setText("Ошибка: " + error)
                if not quiet:
                    self.error(error)
            else:
                self.operation.setText("Операция завершена")
                if on_done:
                    self.guard(lambda: on_done(result))
            self.load_views()
        worker.signals.done.connect(lambda result: finish(result=result))
        worker.signals.failed.connect(lambda error: finish(error=error))
        self.pool.start(worker)

    def refresh_current(self):
        if self.pair_id:
            self.refresh_pair(self.pair_id)

    def refresh_pair(self, pid, quiet=False):
        if pid in self.busy:
            return
        config = self.service.store.pair(pid)["config"]
        if not all(config.get(k) for k in ("document", "google_sheet", "google_uik", "id_field", "time_field", "account")):
            if not quiet:
                self.error("Сначала настройте источник, вход Google и поля УИК, ID, времени")
            self.next_reads[pid] = datetime.now(timezone.utc) + timedelta(seconds=config.get("interval", 60))
            return
        source = GoogleSource(self.root, config["account"])
        self.run_job(pid, lambda: self.service.refresh(pid, lambda: source.read(config["document"], config["google_sheet"])), quiet=quiet)

    def recheck(self):
        if self.ensure_idle():
            self.guard(lambda: self.service.recheck(self.pair_id))
            self.load_views()

    def save_export(self):
        if not self.ensure_idle():
            return
        chosen = [self.export_fields.item(i).text() for i in range(self.export_fields.count())
                  if self.export_fields.item(i).checkState() == Qt.CheckState.Checked]
        self.guard(lambda: self.service.configure(self.pair_id, {"export_excel": chosen}))
        self.load_views()

    def export_project(self):
        def perform():
            snapshot = preview(self.service, self.project_id)
            dialog = ExportDialog(snapshot, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                path, _ = QFileDialog.getSaveFileName(self, "Сохранить экспорт", "Сверка.xlsx", "Excel (*.xlsx)")
                if path:
                    if not path.lower().endswith(".xlsx"):
                        path += ".xlsx"
                    write_xlsx(snapshot, path)
                    self.operation.setText("Экспорт сохранён: " + path)
        self.guard(perform)

    def save_timer(self):
        if self.ensure_idle():
            self.guard(lambda: self.service.configure(self.pair_id, {"auto": self.auto.isChecked(), "interval": self.interval.value()}))
            self.next_reads[self.pair_id] = datetime.now(timezone.utc) + timedelta(seconds=self.interval.value())
            self.update_timer_label()

    def tick(self):
        current = datetime.now(timezone.utc)
        for row in self.service.store.query("SELECT id,active,config FROM pairs"):
            config = json.loads(row["config"])
            if not row["active"] or not config.get("auto"):
                continue
            due = self.next_reads.setdefault(row["id"], current + timedelta(seconds=config.get("interval", 60)))
            if current >= due and row["id"] not in self.busy:
                self.refresh_pair(row["id"], quiet=True)
        self.update_timer_label()

    def update_timer_label(self):
        if not self.pair_id:
            return
        config = self.service.store.pair(self.pair_id)["config"]
        due = self.next_reads.get(self.pair_id)
        self.timer_info.setText(
            "Последнее успешное чтение (UTC): " + config.get("last_read", "не выполнялось")
            + "\nСледующее чтение (UTC): " + (due.isoformat(timespec="seconds") if due and config.get("auto") else "таймер выключен / ожидает планирования")
            + "\nОперация: " + ("выполняется" if self.pair_id in self.busy else "ожидание")
            + "\nОшибка: " + (config.get("last_error") or "нет")
        )

    def show_log(self):
        if not self.project_id:
            return
        rows = self.service.store.query("""SELECT e.at,p.slot,e.kind,e.detail FROM events e
             JOIN pairs p ON p.id=e.pair_id WHERE p.project_id=? ORDER BY e.id DESC LIMIT 1000""", (self.project_id,))
        labels = {"settings": "Настройки", "excel": "Импорт Excel", "read": "Чтение Google", "read_error": "Ошибка чтения",
                  "check": "Сверка", "confirmed": "Подтверждение", "revoked": "Снятие подтверждения",
                  "confirmation_revoked": "Изменение записи", "interrupted": "Прерывание"}
        fill(self.log_grid, [[r["at"], r["slot"], labels.get(r["kind"], r["kind"]), r["detail"]] for r in rows])

    def closeEvent(self, event):
        if self.busy:
            self.error("Дождитесь завершения операций перед закрытием. При аварийном завершении сохранённые данные восстановятся.")
            event.ignore()
            return
        self.timer.stop()
        event.accept()
