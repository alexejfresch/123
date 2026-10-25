import time
import threading

import pytest
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtTest import QTest

from uik_reconcile.ui import ExportDialog, MainWindow
from uik_reconcile.export import preview
from conftest import ingest, setup_pair, submission


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def test_window_project_views_and_export(app, service):
    setup_pair(service)
    ingest(service, [submission()])
    service.confirm(service.history(1)[0]["id"], allow_problem=True)
    window = MainWindow(service.store.root)
    window.show()
    app.processEvents()
    assert window.tabs.count() == 9
    assert window.summary_grid.rowCount() == 1
    assert window.history_grid.rowCount() == 1
    assert window.confirm_grid.rowCount() == 1
    window.history_grid.selectRow(0)
    window.show_details()
    assert "305" in window.details.toPlainText()
    dialog = ExportDialog(preview(service, 1), window)
    assert dialog.windowTitle().startswith("Предварительный")
    dialog.close()
    window.close()


def test_worker_completes_and_releases_pair(app, service):
    window = MainWindow(service.store.root)
    result = []
    callback_threads = []
    def done(value):
        callback_threads.append(threading.get_ident())
        result.append(value)
    window.run_job(1, lambda: "готово", done)
    deadline = time.monotonic() + 5
    while window.busy and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(0.01)
    assert not window.busy
    assert result == ["готово"]
    assert callback_threads == [threading.get_ident()]
    window.close()


def test_ui_setup_read_confirm_export_workflow(app, service, workbook, monkeypatch, tmp_path):
    from PySide6.QtWidgets import QDialog, QFileDialog
    from uik_reconcile.google_source import GoogleSource
    from openpyxl import load_workbook
    path = workbook(rows=[["305", 100]], headers=["Номер", "Всего"])
    service.import_excel(1, path, "Данные", "Номер")
    window = MainWindow(service.store.root)
    window.document.setText("https://docs.google.com/spreadsheets/d/test-document/edit")
    window.google_sheet.setCurrentText("Ответы")
    assert window.save_source()
    monkeypatch.setattr(GoogleSource, "read", lambda *a: (["УИК", "ID", "Время", "Избиратели"], [submission()]))
    def wait():
        deadline = time.monotonic() + 5
        while window.busy and time.monotonic() < deadline:
            app.processEvents()
            time.sleep(0.01)
        assert not window.busy
    window.discover()
    wait()
    for combo, value in [(window.google_uik, "УИК"), (window.id_field, "ID"), (window.time_field, "Время")]:
        combo.setCurrentIndex(combo.findData(value))
    window.time_format.setCurrentText("ISO")
    for i, value in enumerate(["УИК", "Избиратели"]):
        combo = window.mapping_grid.cellWidget(i, 1)
        combo.setCurrentIndex(combo.findData(value))
    window.save_mapping()
    window.refresh_current()
    wait()
    assert window.history_grid.rowCount() == 1
    window.confirm_grid.selectRow(0)
    window.confirm_selected()
    assert service.history(1)[0]["confirmed_at"]
    out = tmp_path / "Экспорт.xlsx"
    monkeypatch.setattr(ExportDialog, "exec", lambda _: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **kw: (str(out), ""))
    window.export_project()
    book = load_workbook(out)
    assert book["Пара 1"]["A2"].value == "305"
    assert book["Пара 1"]["B2"].value == "100"
    book.close()
    window.close()


def test_timer_schedules_only_active_enabled_pairs(app, service, monkeypatch):
    window = MainWindow(service.store.root)
    service.configure(1, {"auto": True, "interval": 10})
    service.configure(2, {"auto": True, "interval": 10}, active=False)
    from datetime import datetime, timedelta, timezone
    window.next_reads[1] = datetime.now(timezone.utc) - timedelta(seconds=1)
    window.next_reads[2] = window.next_reads[1]
    calls = []
    monkeypatch.setattr(window, "refresh_pair", lambda pid, quiet=False: calls.append(pid))
    window.tick()
    assert calls == [1]
    window.close()


@pytest.mark.parametrize("with_project", [False, True])
def test_all_tabs_can_be_clicked_before_and_after_project_creation(app, tmp_path, with_project):
    from uik_reconcile.storage import Store
    root = tmp_path / "navigation"
    if with_project:
        Store(root).create_project("Проверка навигации")
    window = MainWindow(root)
    window.resize(2000, 850)
    window.show()
    app.processEvents()
    bar = window.tabs.tabBar()
    for index in range(window.tabs.count()):
        assert window.tabs.isTabEnabled(index)
        QTest.mouseClick(bar, Qt.MouseButton.LeftButton, pos=bar.tabRect(index).center())
        app.processEvents()
        assert window.tabs.currentIndex() == index
        assert window.tabs.currentWidget().isVisible()
    assert window.project_hint.isVisible() is (not with_project)
    assert window.document.isEnabled() is with_project
    if not with_project:
        from PySide6.QtWidgets import QInputDialog
        from unittest.mock import patch
        with patch.object(QInputDialog, "getText", return_value=("Новый проект", True)):
            window.new_project()
        assert window.document.isEnabled()
        assert not window.project_hint.isVisible()
    window.close()


def contrast(first, second):
    alpha = first.alphaF()
    first = QColor.fromRgbF(
        first.redF() * alpha + second.redF() * (1 - alpha),
        first.greenF() * alpha + second.greenF() * (1 - alpha),
        first.blueF() * alpha + second.blueF() * (1 - alpha),
    )
    def luminance(color):
        values = [color.redF(), color.greenF(), color.blueF()]
        linear = [v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4 for v in values]
        return sum(v * weight for v, weight in zip(linear, [0.2126, 0.7152, 0.0722]))
    a, b = sorted([luminance(first), luminance(second)])
    return (b + 0.05) / (a + 0.05)


def test_dark_system_theme_cannot_leave_light_text_on_light_background(app, service):
    dark = QPalette()
    for role in (QPalette.ColorRole.Window, QPalette.ColorRole.Base, QPalette.ColorRole.Button):
        dark.setColor(role, QColor("#202020"))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        dark.setColor(role, QColor("#ffffff"))
    app.styleHints().setColorScheme(Qt.ColorScheme.Dark)
    app.setPalette(dark)
    window = MainWindow(service.store.root)
    window.show()
    window.tabs.setCurrentIndex(1)
    dialog = ExportDialog(preview(service, 1), window)
    dialog.show()
    app.processEvents()
    # Offscreen не сообщает цветовую схему ОС; проверяем реально применённые цвета.
    assert app.palette().color(QPalette.ColorRole.Window).lightnessF() > 0.9
    pairs = [
        (QPalette.ColorRole.WindowText, QPalette.ColorRole.Window),
        (QPalette.ColorRole.Text, QPalette.ColorRole.Base),
        (QPalette.ColorRole.Text, QPalette.ColorRole.AlternateBase),
        (QPalette.ColorRole.ButtonText, QPalette.ColorRole.Button),
        (QPalette.ColorRole.HighlightedText, QPalette.ColorRole.Highlight),
        (QPalette.ColorRole.ToolTipText, QPalette.ColorRole.ToolTipBase),
        (QPalette.ColorRole.PlaceholderText, QPalette.ColorRole.Base),
    ]
    for widget in (window, window.tabs, window.document, window.source_info, window.history_grid,
                   window.summary_grid.horizontalHeader(), window.check_now, dialog):
        for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive, QPalette.ColorGroup.Disabled):
            for text, background in pairs:
                palette = widget.palette()
                assert contrast(palette.color(group, text), palette.color(group, background)) >= 4.5, (type(widget).__name__, group, text)
    dialog.close()
    window.close()


def test_placeholder_is_readable_when_no_project_exists(app, tmp_path):
    window = MainWindow(tmp_path / "empty")
    window.tabs.setCurrentIndex(1)
    window.show()
    app.processEvents()
    palette = window.document.palette()
    group = QPalette.ColorGroup.Disabled
    assert not window.document.isEnabled()
    assert contrast(palette.color(group, QPalette.ColorRole.PlaceholderText),
                    palette.color(group, QPalette.ColorRole.Base)) >= 4.5
    window.close()
