"""Чистые правила: только обрезка внешних пробелов, без числовых эвристик."""
import hashlib
import json
from datetime import date, datetime
from enum import StrEnum


class Status(StrEnum):
    MATCH = "Совпадение"
    MISMATCH = "Несовпадение"
    NO_EXCEL = "Отсутствие данных Excel"
    NO_GOOGLE = "Отсутствие данных Google"
    READ_ERROR = "Ошибка чтения"
    UNMAPPABLE = "Невозможность сопоставления"


def dumps(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=_json_default)


def _json_default(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    raise TypeError(f"Неподдерживаемый тип данных: {type(value).__name__}")


def normalize(value):
    if value is None:
        return None
    if isinstance(value, (date, datetime)):
        return value.isoformat().strip()
    return str(value).strip()


def fingerprint(row):
    content = {k: normalize(v) for k, v in row.items()}
    return hashlib.sha256(dumps(content).encode("utf-8")).hexdigest()


def compare(excel, google, mapping, excel_headers, google_headers):
    if google is None:
        return {"status": Status.NO_GOOGLE, "differences": []}
    if excel is None:
        return {"status": Status.NO_EXCEL, "differences": []}
    if not mapping or any(e not in excel_headers or g not in google_headers for e, g in mapping.items()):
        return {"status": Status.UNMAPPABLE, "differences": [], "reason": "Проверьте схему и названия столбцов"}
    differences = []
    for e, g in mapping.items():
        left, right = excel.get(e), google.get(g)
        if normalize(left) != normalize(right):
            differences.append({"excel_field": e, "google_field": g, "excel": left, "google": right})
    return {"status": Status.MISMATCH if differences else Status.MATCH, "differences": differences}
