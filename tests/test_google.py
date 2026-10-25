from unittest.mock import MagicMock

import pytest

from uik_reconcile.google_source import GoogleSource, spreadsheet_id


def test_spreadsheet_url():
    assert spreadsheet_id("https://docs.google.com/spreadsheets/d/abc-123/edit#gid=0") == "abc-123"
    with pytest.raises(ValueError):
        spreadsheet_id("not a sheet")


def test_official_api_all_rows_blank_position_and_quoted_title(monkeypatch, tmp_path):
    api = MagicMock()
    call = api.spreadsheets.return_value.values.return_value.get
    call.return_value.execute.return_value = {"values": [["ID", "УИК"], ["a", "305"], [], ["b", "306"]]}
    monkeypatch.setattr(GoogleSource, "_api", lambda _: api)
    source = GoogleSource(tmp_path, "test")
    headers, rows = source.read("abc", "Ответы 'формы'")
    assert headers == ["ID", "УИК"]
    assert len(rows) == 3 and rows[1] == {"ID": None, "УИК": None}
    assert call.call_args.kwargs["range"] == "'Ответы ''формы'''"
    assert call.call_args.kwargs["valueRenderOption"] == "FORMATTED_VALUE"
    api.close.assert_called_once()


def test_sheets_failure_is_russian_and_closes_client(monkeypatch, tmp_path):
    api = MagicMock()
    api.spreadsheets.return_value.values.return_value.get.return_value.execute.side_effect = OSError("offline")
    monkeypatch.setattr(GoogleSource, "_api", lambda _: api)
    with pytest.raises(RuntimeError, match="Сохранённые данные доступны"):
        GoogleSource(tmp_path, "test").read("abc", "Лист")
    api.close.assert_called_once()


def test_no_plaintext_token_fallback(monkeypatch, tmp_path):
    def unavailable(*args):
        raise RuntimeError("нет сейфа")
    monkeypatch.setattr("uik_reconcile.google_source.keyring.get_password", unavailable)
    with pytest.raises(RuntimeError, match="защищённое хранилище"):
        GoogleSource(tmp_path, "test").authorize(tmp_path / "client.json")
    assert list(tmp_path.iterdir()) == []
