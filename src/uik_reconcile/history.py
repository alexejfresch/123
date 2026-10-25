"""Порядок отправок и единое правило выбора для сводки и экспорта."""
from datetime import datetime, timezone

from .domain import Status, normalize


def parse_time(value, config):
    text = normalize(value)
    if not text:
        raise ValueError("У отправки отсутствует время")
    try:
        fmt = config.get("time_format", "%d.%m.%Y %H:%M:%S")
        dt = datetime.fromisoformat(text.replace("Z", "+00:00")) if fmt == "ISO" else datetime.strptime(text, fmt)
    except ValueError as exc:
        raise ValueError("Не удалось разобрать время отправки. Проверьте формат в настройках пары.") from exc
    return dt


def time_key(value, config):
    dt = parse_time(value, config)
    if dt.tzinfo:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt.isoformat(timespec="microseconds")


def ordered_records(records, config):
    return sorted(records, key=lambda r: (time_key(r["source_time"], config), r["position"]))


def select_latest(records, config):
    ordered = ordered_records(records, config)
    valid = [r for r in ordered if r["confirmed_at"] and r["confirmation_hash"] == r["fingerprint"]]
    if not valid:
        return None, []
    selected = valid[-1]
    if sum(time_key(r["source_time"], config) == time_key(selected["source_time"], config) for r in ordered) > 1 and not config.get("ties_by_row"):
        return None, ["Одинаковое время отправок: порядок неоднозначен. Разрешите порядок по строкам в настройках."]
    notes = [f"Более поздняя отправка ID {r['source_key']}: {r['check_result']['status']}; не подтверждена"
             for r in ordered[ordered.index(selected)+1:]]
    if not selected["present"]:
        notes.append("Подтверждённая отправка отсутствует в последнем снимке источника")
    if selected["check_result"]["status"] != Status.MATCH:
        notes.append("Ручное подтверждение: " + selected["check_result"]["status"])
    return selected, notes
