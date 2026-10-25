"""Экспорт только локального состояния: ни чтения источников, ни перепроверки."""
import os
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from .history import select_latest


def preview(service, project_id):
    sheets, warnings = [], []
    # Один согласованный локальный снимок для всех пар, включая подтверждения.
    with service.store.transaction() as db:
        pairs = db.execute("SELECT id,slot,config FROM pairs WHERE project_id=? AND active=1 ORDER BY slot", (project_id,)).fetchall()
        import json
        for pair in pairs:
            config = json.loads(pair["config"])
            records = [dict(r) for r in db.execute("SELECT * FROM submissions WHERE pair_id=? AND current=1", (pair["id"],))]
            for r in records:
                r["payload"] = json.loads(r["payload"])
                r["check_result"] = json.loads(r["check_result"])
            mapping = config.get("mapping", {})
            if not config.get("last_read"):
                warnings.append(f"Пара {pair['slot']}: Google ещё не прочитан")
            excel_headers = config.get("excel_headers", [])
            google_headers = list(config.get("google_headers", []))
            # Удалённые из источника столбцы остаются в сохранённых отправках.
            for record in records:
                for field in record["payload"]:
                    if field not in google_headers:
                        google_headers.append(field)
            google_extra = [h for h in google_headers if h not in mapping.values()]
            excel_extra = [h for h in config.get("export_excel", []) if h in excel_headers and h not in mapping]
            common_order = [e for e in excel_headers if e in mapping]
            common_order += [e for e in mapping if e not in common_order]
            columns = [(e, "common", mapping[e]) for e in common_order]
            columns += [(g, "google", g) for g in google_extra]
            columns += [(e, "excel", e) for e in excel_extra]
            # Префиксы только для настоящих коллизий заголовков.
            names = [c[0] for c in columns]
            headers = [f"{name} ({'Google' if kind == 'google' else 'Excel'})" if names.count(name) > 1 else name
                       for name, kind, _ in columns]
            if len(headers) != len(set(headers)):
                raise ValueError("Коллизия заголовков экспорта. Уточните названия полей или сопоставление.")
            rows, source_ids = [], []
            excel_index = service.excel_index(config)
            for uik in sorted({r["uik"] for r in records}):
                group = [r for r in records if r["uik"] == uik]
                # Этот метод только читает БД. Конфигурация не может измениться пока удерживается транзакция.
                selected, notes = select_latest(group, config)
                warnings.extend(f"Пара {pair['slot']}, УИК {uik}: {n}" for n in notes)
                if selected is None:
                    warnings.append(f"Пара {pair['slot']}, УИК {uik}: нет выбранной подтверждённой отправки")
                    continue
                google, excel = selected["payload"], excel_index.get(uik, {})
                rows.append([google.get(field) if kind in ("common", "google") else excel.get(field)
                             for _, kind, field in columns])
                source_ids.append(selected["id"])
            if config.get("last_error"):
                warnings.append(f"Пара {pair['slot']}: последнее чтение не удалось; используется прежний снимок")
            sheets.append({"name": f"Пара {pair['slot']}", "headers": headers, "rows": rows,
                           "last_read": config.get("last_read"), "source_ids": source_ids})
    return {"sheets": sheets, "warnings": warnings}


def write_xlsx(snapshot, destination):
    """Записывается именно показанный предпросмотр, без повторного вычисления."""
    path = Path(destination)
    if not snapshot["sheets"]:
        raise ValueError("Нет активных пар для экспорта")
    book = Workbook()
    book.remove(book.active)
    for sheet in snapshot["sheets"]:
        ws = book.create_sheet(sheet["name"])
        for row in [sheet["headers"], *sheet["rows"]]:
            ws.append(row)
        # Источник может содержать строки, начинающиеся на '=': сохраняем их текстом.
        for row in ws:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="245D7C")
        ws.freeze_panes = "A2"
        if sheet["headers"]:
            ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(48, max(16, max(len(str(c.value or "")) for c in col) + 2))
    meta = book.create_sheet("Сведения")
    meta.append(["Пара", "Последнее успешное чтение (UTC)"])
    for sheet in snapshot["sheets"]:
        meta.append([sheet["name"], sheet["last_read"] or "Не выполнялось"])
    meta.append(["Предупреждения"])
    for note in snapshot["warnings"]:
        meta.append([note])
    for row in meta:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
    fd, temporary = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)
    try:
        book.save(temporary)
        os.replace(temporary, path)
    finally:
        book.close()
        Path(temporary).unlink(missing_ok=True)
