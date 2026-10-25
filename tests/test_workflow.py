import json
import shutil
import sqlite3
import threading
import subprocess
import sys

import pytest
from openpyxl import load_workbook

from uik_reconcile.domain import Status, fingerprint, normalize
from uik_reconcile.excel import read_table, sheet_names
from uik_reconcile.export import preview, write_xlsx
from uik_reconcile.service import Service
from uik_reconcile.storage import now
from conftest import HEADERS, ingest, setup_pair, submission


def excel(service, workbook, pair_id=1, rows=None):
    path = workbook(rows=rows or [["305", 100, "только Excel"]], headers=["Номер", "Всего", "Excel-поле"], name=f"input{pair_id}.xlsx")
    service.import_excel(pair_id, path, "Данные", "Номер")
    return path


def test_one_pair(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission()])
    record = service.history(1)[0]
    assert record["check_result"]["status"] == Status.MATCH
    service.confirm(record["id"])
    pairs, rows = service.summary(1)
    assert len(pairs) == 1 and rows[0]["total"] == "Готово"
    assert preview(service, 1)["sheets"][0]["rows"][0][:2] == ["305", "100"]


def test_three_pairs_union(service, workbook):
    for pid, uik in enumerate(["291", "292", "305"], 1):
        setup_pair(service, pid)
        excel(service, workbook, pid, [[uik, 100, "x"]])
        ingest(service, [submission(str(pid), uik)], pid)
    pairs, rows = service.summary(1)
    assert len(pairs) == 3
    assert {r["uik"] for r in rows} == {"291", "292", "305"}
    assert rows[0]["pairs"][2]["status"] == Status.NO_GOOGLE
    assert len(preview(service, 1)["sheets"]) == 3


def test_excel_only_and_google_only_uiks(service, workbook):
    setup_pair(service)
    excel(service, workbook, rows=[["291", 100, "x"]])
    ingest(service, [submission(uik="305")])
    _, rows = service.summary(1)
    cells = {r["uik"]: r["pairs"][1]["status"] for r in rows}
    assert cells == {"291": Status.NO_GOOGLE, "305": Status.NO_EXCEL}


@pytest.mark.parametrize("empty", [False, True])
def test_missing_or_empty_excel(service, workbook, empty):
    setup_pair(service)
    if empty:
        service.import_excel(1, workbook(), "Данные", "")
        assert service.store.pair(1)["config"]["excel_status"] == "Пуст"
    ingest(service, [submission()])
    record = service.history(1)[0]
    assert record["check_result"]["status"] == Status.NO_EXCEL
    with pytest.raises(ValueError):
        service.confirm(record["id"])
    service.confirm(record["id"], allow_problem=True)
    assert len(preview(service, 1)["sheets"][0]["rows"]) == 1


def test_dedup_all_submissions_and_latest_bad(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    a, b = submission(), submission("b", value="999", minute=1)
    ingest(service, [a])
    first_id = service.history(1)[0]["id"]
    service.confirm(first_id)
    ingest(service, [a, b])
    ingest(service, [a, b])
    assert len(service.history(1)) == 2
    chosen, notes = service.latest_confirmed(1, "305")
    assert chosen["id"] == first_id
    assert any("Более поздняя" in n and "Несовпадение" in n for n in notes)
    snapshot = preview(service, 1)
    assert snapshot["sheets"][0]["source_ids"] == [first_id]
    assert any("Более поздняя" in n for n in snapshot["warnings"])


def test_manual_mismatch_and_latest_confirmed(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission(), submission("b", value="999", minute=1)])
    records = service.history(1)
    for r in records:
        service.confirm(r["id"], allow_problem=True)
    chosen, notes = service.latest_confirmed(1, "305")
    assert chosen["source_key"] == "b"
    assert "999" in preview(service, 1)["sheets"][0]["rows"][0]


def test_content_change_revokes_only_changed_record(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    a, b = submission(), submission("b", uik="306", minute=1)
    ingest(service, [a, b])
    for r in service.history(1):
        service.confirm(r["id"], allow_problem=True)
    changed = dict(a, Избиратели="101")
    ingest(service, [changed, b])
    current = {r["source_key"]: r for r in service.history(1, current_only=True)}
    assert current["a"]["confirmed_at"] is None
    assert current["b"]["confirmed_at"] is not None
    assert len(service.history(1)) == 3
    assert service.latest_confirmed(1, "305")[0] is None
    assert service.store.query("SELECT * FROM events WHERE kind='confirmation_revoked'")


def test_spaces_keep_confirmation_and_original_values(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    a = submission()
    ingest(service, [a])
    service.confirm(service.history(1)[0]["id"])
    a["УИК"] = " 305 "
    a["Избиратели"] = " 100 "
    ingest(service, [a])
    current = service.history(1, current_only=True)[0]
    assert current["confirmed_at"]
    assert current["check_result"]["status"] == Status.MATCH
    assert current["payload"]["УИК"] == " 305 "
    assert normalize(305) == normalize(" 305 ")
    assert normalize("0305") != normalize(305)
    assert normalize("1 000") != normalize("1000")


def test_network_failure_retains_cache(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission()])
    record = service.history(1)[0]
    service.confirm(record["id"])
    last_read = service.store.pair(1)["config"]["last_read"]
    def offline():
        raise OSError("Интернет недоступен")
    with pytest.raises(OSError):
        service.refresh(1, offline)
    assert service.store.pair(1)["config"]["last_read"] == last_read
    assert service.latest_confirmed(1, "305")[0]["id"] == record["id"]
    assert preview(service, 1)["sheets"][0]["rows"]
    assert service.store.query("SELECT state FROM operations ORDER BY id DESC")[0]["state"] == "failed"


@pytest.mark.parametrize("kind", ["count", "rename"])
def test_schema_change(service, workbook, kind):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission()])
    row = submission()
    if kind == "count":
        row["Новое"] = "новое значение"
        headers = HEADERS + ["Новое"]
    else:
        row["Другое имя"] = row.pop("Избиратели")
        headers = ["Другое имя" if h == "Избиратели" else h for h in HEADERS]
    ingest(service, [row], headers=headers)
    config = service.store.pair(1)["config"]
    assert config["schema_change"]["added"]
    if kind == "rename":
        assert service.history(1)[0]["check_result"]["status"] == Status.UNMAPPABLE
        service.configure(1, {"mapping": {"Номер": "УИК", "Всего": "Другое имя"}})
    assert service.history(1)[0]["check_result"]["status"] == Status.MATCH


def test_invalid_batch_is_atomic(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission()])
    before = service.history(1)
    with pytest.raises(ValueError, match="повторяется ID"):
        ingest(service, [submission(value="111"), submission(value="222")])
    assert service.history(1) == before


def test_interrupted_read_recovered(service):
    setup_pair(service)
    ingest(service, [submission()])
    def interrupted():
        raise KeyboardInterrupt()
    with pytest.raises(KeyboardInterrupt):
        service.refresh(1, interrupted)
    assert service.store.query("SELECT state FROM operations ORDER BY id DESC")[0]["state"] == "running"
    reopened = Service(service.store.root)
    reopened.store.recover()
    assert reopened.store.query("SELECT state FROM operations ORDER BY id DESC")[0]["state"] == "interrupted"
    assert len(reopened.history(1)) == 1


def test_transaction_rollback(service):
    with pytest.raises(RuntimeError):
        with service.store.transaction() as db:
            db.execute("UPDATE projects SET name='не сохранять'")
            raise RuntimeError("Аварийное закрытие")
    assert service.store.query("SELECT name FROM projects")[0]["name"] == "Тестовый проект"


def test_real_process_crash_preserves_completed_state(service):
    setup_pair(service)
    ingest(service, [submission()])
    service.confirm(service.history(1)[0]["id"], allow_problem=True)
    script = """
import os, sqlite3, sys
db = sqlite3.connect(sys.argv[1])
db.execute('BEGIN IMMEDIATE')
db.execute("UPDATE submissions SET confirmed_at=NULL")
db.execute("UPDATE projects SET name='НЕ СОХРАНЯТЬ'")
os._exit(23)
"""
    crashed = subprocess.run([sys.executable, "-c", script, str(service.store.path)])
    assert crashed.returncode == 23
    reopened = Service(service.store.root)
    assert reopened.history(1)[0]["confirmed_at"]
    assert reopened.store.query("SELECT name FROM projects")[0]["name"] == "Тестовый проект"
    assert reopened.store.query("PRAGMA integrity_check")[0]["integrity_check"] == "ok"


def test_portable_internal_copy_and_reopen(service, workbook, tmp_path):
    setup_pair(service)
    original = excel(service, workbook)
    ingest(service, [submission()])
    service.confirm(service.history(1)[0]["id"])
    original.unlink()
    moved = tmp_path / "другая папка" / "data"
    shutil.copytree(service.store.root, moved)
    reopened = Service(moved)
    config = reopened.store.pair(1)["config"]
    assert not __import__('pathlib').Path(config["excel_file"]).is_absolute()
    assert sheet_names(moved / config["excel_file"]) == ["Данные"]
    assert reopened.history(1)[0]["confirmed_at"]
    assert preview(reopened, 1)["sheets"][0]["rows"]


def test_export_offline_snapshot_unique_fields_and_no_side_effects(service, workbook, tmp_path, monkeypatch):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission(note="=HYPERLINK(\"x\")")])
    service.confirm(service.history(1)[0]["id"])
    service.configure(1, {"export_excel": ["Excel-поле"]})
    before = service.history(1)
    def forbidden(*a, **kw):
        raise AssertionError("Экспорт не должен читать Google или выполнять проверку")
    monkeypatch.setattr(service, "refresh", forbidden)
    monkeypatch.setattr(service, "recheck", forbidden)
    monkeypatch.setattr(service, "_check", forbidden)
    snapshot = preview(service, 1)
    assert service.history(1) == before
    assert snapshot["sheets"][0]["headers"] == ["Номер", "Всего", "ID", "Время", "Примечание", "Excel-поле"]
    out = tmp_path / "export.xlsx"
    write_xlsx(snapshot, out)
    book = load_workbook(out)
    assert book["Пара 1"]["E2"].data_type == "s"
    assert book["Пара 1"]["F2"].value == "только Excel"
    assert book["Сведения"]["B2"].value
    book.close()


def test_order_uses_time_not_read_or_row_order(service):
    setup_pair(service)
    ingest(service, [submission("new", minute=2), submission("old", minute=1)])
    for r in service.history(1):
        service.confirm(r["id"], allow_problem=True)
    assert service.latest_confirmed(1, "305")[0]["source_key"] == "new"


def test_tied_times_require_explicit_decision(service):
    setup_pair(service)
    ingest(service, [submission("a"), submission("b")])
    for r in service.history(1):
        service.confirm(r["id"], allow_problem=True)
    assert service.latest_confirmed(1, "305")[0] is None
    assert not preview(service, 1)["sheets"][0]["rows"]
    service.configure(1, {"ties_by_row": True})
    assert service.latest_confirmed(1, "305")[0]["source_key"] == "b"


def test_removed_source_record_retained_with_warning(service):
    setup_pair(service)
    ingest(service, [submission()])
    service.confirm(service.history(1)[0]["id"], allow_problem=True)
    ingest(service, [])
    chosen, notes = service.latest_confirmed(1, "305")
    assert chosen and any("отсутствует" in n for n in notes)


def test_id_moves_without_losing_confirmation(service):
    setup_pair(service)
    a, b = submission(), submission("b", minute=1)
    ingest(service, [a, b])
    service.confirm(service.history(1)[1]["id"], allow_problem=True)
    ingest(service, [b, a])
    assert len(service.history(1)) == 2
    assert next(r for r in service.history(1) if r["source_key"] == "a")["confirmed_at"]


def test_no_concurrent_read_for_same_pair(service):
    setup_pair(service)
    entered, release = threading.Event(), threading.Event()
    failures = []
    def reader():
        entered.set()
        release.wait(timeout=5)
        return HEADERS, [submission()]
    def work():
        try:
            service.refresh(1, reader)
        except Exception as exc:
            failures.append(exc)
    worker = threading.Thread(target=work)
    worker.start()
    assert entered.wait(timeout=5)
    try:
        with pytest.raises(RuntimeError, match="уже обрабатывается"):
            service.refresh(1, reader)
    finally:
        release.set()
        worker.join(timeout=5)
    assert not failures
    assert len(service.history(1)) == 1


def test_duplicate_excel_uik_rejected(service, workbook):
    setup_pair(service)
    path = workbook(rows=[["305", 1], [" 305 ", 2]], headers=["Номер", "Всего"])
    with pytest.raises(ValueError, match="несколько строк"):
        service.import_excel(1, path, "Данные", "Номер")


def test_missing_id_or_time_does_not_substitute_read_time(service):
    setup_pair(service)
    with pytest.raises(ValueError, match="время"):
        ingest(service, [dict(submission(), Время="")])
    assert not service.history(1)
    assert "last_read" not in service.store.pair(1)["config"]


def test_cannot_confirm_stale_revision(service):
    setup_pair(service)
    ingest(service, [submission()])
    old = service.history(1)[0]["id"]
    ingest(service, [submission(value="200")])
    with pytest.raises(ValueError, match="предыдущая версия"):
        service.confirm(old, allow_problem=True)


def test_preserve_unique_google_field_removed_from_latest_structure(service):
    setup_pair(service)
    ingest(service, [submission()])
    service.confirm(service.history(1)[0]["id"], allow_problem=True)
    # Старой отправки больше нет в листе, но её подтверждённые уникальные значения не теряются.
    ingest(service, [], headers=[h for h in HEADERS if h != "Примечание"])
    snapshot = preview(service, 1)
    headers = snapshot["sheets"][0]["headers"]
    assert "Примечание" in headers
    assert snapshot["sheets"][0]["rows"][0][headers.index("Примечание")] == "исходное"


def test_export_same_header_different_sources(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    row = dict(submission(), **{"Excel-поле": "только Google"})
    ingest(service, [row], headers=HEADERS + ["Excel-поле"])
    service.confirm(service.history(1)[0]["id"])
    service.configure(1, {"export_excel": ["Excel-поле"]})
    sheet = preview(service, 1)["sheets"][0]
    assert "Excel-поле (Google)" in sheet["headers"]
    assert "Excel-поле (Excel)" in sheet["headers"]
    assert "только Google" in sheet["rows"][0]
    assert "только Excel" in sheet["rows"][0]


def test_invalid_time_format_change_rolls_back_settings(service):
    setup_pair(service)
    ingest(service, [submission()])
    with pytest.raises(ValueError):
        service.configure(1, {"time_format": "%d.%m.%Y"})
    assert service.store.pair(1)["config"]["time_format"] == "ISO"


def test_latest_successful_read_with_changed_excel_warns_for_manual_confirmation(service, workbook):
    setup_pair(service)
    excel(service, workbook)
    ingest(service, [submission()])
    service.confirm(service.history(1)[0]["id"])
    excel(service, workbook, rows=[["305", 101, "изменённое"]])
    assert service.history(1)[0]["confirmed_at"]
    assert any("Ручное подтверждение" in s for s in preview(service, 1)["warnings"])


def test_timezone_offsets_are_used_for_order(service):
    setup_pair(service)
    a = dict(submission("a"), Время="2026-10-25T15:00:00+05:00")
    b = dict(submission("b"), Время="2026-10-25T11:00:00+00:00")
    ingest(service, [a, b])
    for r in service.history(1):
        service.confirm(r["id"], allow_problem=True)
    assert service.latest_confirmed(1, "305")[0]["source_key"] == "b"


def test_mixed_timezone_metadata_does_not_silently_change_order(service):
    setup_pair(service)
    with pytest.raises(ValueError, match="часовым поясом"):
        ingest(service, [submission("a"), dict(submission("b"), Время="2026-10-25T11:00:00Z")])
    assert not service.history(1)
