import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from openpyxl import Workbook

from uik_reconcile.service import Service


@pytest.fixture
def service(tmp_path):
    s = Service(tmp_path / "data")
    s.store.create_project("Тестовый проект")
    return s


@pytest.fixture
def workbook(tmp_path):
    def make(rows=None, headers=None, name="input.xlsx"):
        path = tmp_path / name
        book = Workbook()
        book.active.title = "Данные"
        if headers is not None:
            book.active.append(headers)
        for row in rows or []:
            book.active.append(row)
        book.save(path)
        book.close()
        return path
    return make


def setup_pair(service, pair_id=1, active=True):
    service.configure(pair_id, {
        "google_uik": "УИК", "id_field": "ID", "time_field": "Время", "time_format": "ISO",
        "mapping": {"Номер": "УИК", "Всего": "Избиратели"},
    }, active=active)


HEADERS = ["УИК", "ID", "Время", "Избиратели", "Примечание"]


def submission(key="a", uik="305", value="100", minute=0, note="исходное"):
    return {"УИК": uik, "ID": key, "Время": f"2026-10-25T12:{minute:02d}:00", "Избиратели": value, "Примечание": note}


def ingest(service, rows, pair_id=1, headers=None):
    service.refresh(pair_id, lambda: (headers if headers is not None else HEADERS, rows))
