import time
import threading

import pytest
from PySide6.QtWidgets import QApplication

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
