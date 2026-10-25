"""Прикладные операции. Получение данных отделено от проверки и экспорта."""
import json
import threading

from .domain import Status, compare, dumps, fingerprint, normalize
from .excel import copy_internal, read_table, validate_headers
from .storage import Store, now
from .history import ordered_records, parse_time, select_latest, time_key

_locks_guard = threading.Lock()
_locks = {}


def pair_lock(path, pair_id):
    with _locks_guard:
        return _locks.setdefault((str(path), pair_id), threading.RLock())


class Service:
    def __init__(self, root):
        self.store = Store(root)

    def configure(self, pair_id, changes, active=None):
        with pair_lock(self.store.path, pair_id):
            pair = self.store.pair(pair_id)
            config = pair["config"]
            # Смена документа/листа/ID при наличии истории может ошибочно переиспользовать ID.
            identity_fields = ("document", "google_sheet", "id_field", "google_uik")
            has_history = self.store.query("SELECT 1 FROM submissions WHERE pair_id=? LIMIT 1", (pair_id,))
            if has_history and any(k in changes and changes[k] != config.get(k) for k in identity_fields):
                raise ValueError("Источник или столбец ID уже имеет историю. Создайте новый проект для другого источника.")
            if active is False and pair["active"]:
                others = self.store.query("SELECT id FROM pairs WHERE project_id=? AND id<>? AND active=1", (pair["project_id"], pair_id))
                if not others:
                    raise ValueError("В проекте должна оставаться хотя бы одна активная пара")
            if has_history and any(k in changes and changes[k] != config.get(k) for k in ("time_field", "time_format")):
                candidate = dict(config, **changes)
                for r in self.history(pair_id, current_only=True):
                    time_key(r["payload"].get(candidate.get("time_field")), candidate)
            config.update(changes)
            with self.store.transaction() as db:
                Store.write_config(db, pair_id, config)
                if active is not None:
                    db.execute("UPDATE pairs SET active=? WHERE id=?", (int(active), pair_id))
                if has_history and "time_field" in changes:
                    for r in db.execute("SELECT id,payload FROM submissions WHERE pair_id=? AND current=1", (pair_id,)).fetchall():
                        value = json.loads(r["payload"]).get(config["time_field"])
                        db.execute("UPDATE submissions SET source_time=? WHERE id=?", (normalize(value), r["id"]))
                self._recheck(db, pair_id, config)
                Store.event(db, pair_id, "settings", "Настройки сохранены; выполнена локальная сверка")

    def import_excel(self, pair_id, path, sheet, uik_field):
        with pair_lock(self.store.path, pair_id):
            headers, rows = read_table(path, sheet)
            if headers and uik_field not in headers:
                raise ValueError("Выберите столбец УИК Excel")
            keys = [normalize(r.get(uik_field)) for r in rows]
            if any(not key for key in keys):
                raise ValueError("В Excel есть строка без номера УИК")
            if len(set(keys)) != len(keys):
                raise ValueError("В Excel несколько строк одной УИК. Устраните неоднозначность до импорта.")
            relative = copy_internal(self.store.root, path)
            # Снимок хранится в БД вместе с относительным путём копии.
            config = self.store.pair(pair_id)["config"]
            config.update(excel_file=relative, excel_sheet=sheet, excel_uik=uik_field,
                          excel_headers=headers, excel_rows=rows,
                          excel_status="Пуст" if not rows else "Загружен")
            with self.store.transaction() as db:
                Store.write_config(db, pair_id, config)
                self._recheck(db, pair_id, config)
                Store.event(db, pair_id, "excel", f"Импортировано строк: {len(rows)}; лист: {sheet}")

    @staticmethod
    def excel_index(config):
        return {normalize(r.get(config.get("excel_uik"))): r for r in config.get("excel_rows", [])}

    def _check(self, payload, config, index=None):
        google_uik = config.get("google_uik")
        if not google_uik or google_uik not in payload or not normalize(payload.get(google_uik)):
            return {"status": Status.UNMAPPABLE, "differences": [], "reason": "Отсутствует поле или номер УИК Google"}
        excel = (index if index is not None else self.excel_index(config)).get(normalize(payload[google_uik]))
        return compare(excel, payload, config.get("mapping", {}), config.get("excel_headers", []),
                       config.get("google_headers", []))

    def refresh(self, pair_id, reader):
        lock = pair_lock(self.store.path, pair_id)
        if not lock.acquire(blocking=False):
            raise RuntimeError("Эта пара уже обрабатывается")
        op_id = None
        try:
            config = self.store.pair(pair_id)["config"]
            with self.store.transaction() as db:
                op_id = db.execute("INSERT INTO operations(pair_id,started,state) VALUES(?,?,'running')", (pair_id, now())).lastrowid
            headers, rows = reader()  # Сеть вне транзакции. При сбое прежний снимок не меняется.
            self._ingest(pair_id, config, headers, rows, op_id)
        except Exception as exc:
            if op_id is not None:
                with self.store.transaction() as db:
                    db.execute("UPDATE operations SET state='failed',finished=?,error=? WHERE id=?",
                               (now(), str(exc), op_id))
                    config = json.loads(db.execute("SELECT config FROM pairs WHERE id=?", (pair_id,)).fetchone()[0])
                    config["last_error"] = str(exc)
                    Store.write_config(db, pair_id, config)
                    Store.event(db, pair_id, "read_error", str(exc))
            raise
        finally:
            lock.release()

    def _ingest(self, pair_id, config, headers, rows, op_id):
        validate_headers(headers)
        for field in ("google_uik", "id_field", "time_field"):
            if not config.get(field) or config[field] not in headers:
                raise ValueError("Выберите существующие столбцы УИК, ID и времени Google; структура могла измениться")
        prepared, seen, timezone_modes = [], set(), set()
        for position, row in enumerate(rows, 2):
            if not any(v is not None and str(v) != "" for v in row.values()):
                continue
            key = normalize(row.get(config["id_field"]))
            uik = normalize(row.get(config["google_uik"]))
            if not key or not uik:
                raise ValueError(f"Строка {position}: отсутствует ID или номер УИК")
            if key in seen:
                raise ValueError(f"Строка {position}: повторяется ID отправки; снимок не сохранён")
            seen.add(key)
            stamp = parse_time(row.get(config["time_field"]), config)
            timezone_modes.add(stamp.tzinfo is not None)
            prepared.append((key, uik, row, position))
        for prior in self.history(pair_id, current_only=True):
            if prior["source_key"] not in seen:
                timezone_modes.add(parse_time(prior["source_time"], config).tzinfo is not None)
        if len(timezone_modes) > 1:
            raise ValueError("Время отправок смешивает значения с часовым поясом и без него. Приведите время источника к единому формату.")
        old_headers = config.get("google_headers", [])
        config["google_headers"] = headers
        config["schema_change"] = ({"removed": [h for h in old_headers if h not in headers],
                                    "added": [h for h in headers if h not in old_headers]}
                                   if old_headers != headers and old_headers else {})
        config["last_read"] = now()
        config["last_error"] = ""
        with self.store.transaction() as db:
            db.execute("UPDATE submissions SET present=0 WHERE pair_id=? AND current=1", (pair_id,))
            excel_index = self.excel_index(config)
            for key, uik, row, position in prepared:
                old = db.execute("SELECT * FROM submissions WHERE pair_id=? AND source_key=? AND current=1", (pair_id, key)).fetchone()
                fp = fingerprint(row)
                payload = dumps(row)
                result = dumps(self._check(row, config, excel_index))
                if old and old["payload"] == payload:
                    db.execute("UPDATE submissions SET present=1,position=?,check_result=? WHERE id=?", (position, result, old["id"]))
                    continue
                confirmed = old["confirmed_at"] if old and old["fingerprint"] == fp else None
                if old:
                    db.execute("UPDATE submissions SET current=0 WHERE id=?", (old["id"],))
                    if old["confirmed_at"] and confirmed is None:
                        Store.event(db, pair_id, "confirmation_revoked", f"ID {key}: содержимое изменилось; подтверждение версии {old['revision']} снято")
                db.execute("""INSERT INTO submissions(pair_id,source_key,revision,uik,fingerprint,payload,source_time,
                           position,observed_at,confirmed_at,confirmation_hash,check_result)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                           (pair_id, key, old["revision"] + 1 if old else 1, uik, fp, payload,
                            normalize(row.get(config["time_field"])), position, now(), confirmed,
                            fp if confirmed else None, result))
            Store.write_config(db, pair_id, config)
            # Даже для отсутствующих в новом снимке записей сохраняются история и проверки.
            self._recheck(db, pair_id, config)
            db.execute("UPDATE operations SET state='success',finished=? WHERE id=?", (now(), op_id))
            Store.event(db, pair_id, "read", f"Прочитано отправок: {len(prepared)}")

    def _recheck(self, db, pair_id, config):
        index = self.excel_index(config)
        for record in db.execute("SELECT id,payload FROM submissions WHERE pair_id=? AND current=1", (pair_id,)).fetchall():
            result = self._check(json.loads(record["payload"]), config, index)
            db.execute("UPDATE submissions SET check_result=? WHERE id=?", (dumps(result), record["id"]))

    def recheck(self, pair_id):
        with pair_lock(self.store.path, pair_id), self.store.transaction() as db:
            config = json.loads(db.execute("SELECT config FROM pairs WHERE id=?", (pair_id,)).fetchone()[0])
            self._recheck(db, pair_id, config)
            Store.event(db, pair_id, "check", "Выполнена локальная сверка без чтения Google")

    def confirm(self, submission_id, allow_problem=False):
        record = self.store.query("SELECT pair_id FROM submissions WHERE id=?", (submission_id,))
        if not record:
            raise ValueError("Отправка не найдена")
        pair_id = record[0]["pair_id"]
        with pair_lock(self.store.path, pair_id), self.store.transaction() as db:
            row = db.execute("SELECT * FROM submissions WHERE id=?", (submission_id,)).fetchone()
            if not row["current"]:
                raise ValueError("Это предыдущая версия изменённой отправки. Выберите актуальную версию.")
            result = json.loads(row["check_result"])
            if result["status"] == Status.UNMAPPABLE:
                raise ValueError("Сначала настройте сопоставление полей")
            if result["status"] != Status.MATCH and not allow_problem:
                raise ValueError("Есть расхождения или отсутствует Excel. Требуется явное ручное подтверждение.")
            db.execute("UPDATE submissions SET confirmed_at=?,confirmation_hash=fingerprint WHERE id=?", (now(), submission_id))
            Store.event(db, pair_id, "confirmed", f"ID {row['source_key']}, версия {row['revision']}: {result['status']}")

    def revoke(self, submission_id):
        with self.store.transaction() as db:
            row = db.execute("SELECT pair_id,source_key FROM submissions WHERE id=?", (submission_id,)).fetchone()
            if row:
                db.execute("UPDATE submissions SET confirmed_at=NULL,confirmation_hash=NULL WHERE id=?", (submission_id,))
                Store.event(db, row["pair_id"], "revoked", f"ID {row['source_key']}: подтверждение снято пользователем")

    def history(self, pair_id, uik=None, current_only=False):
        sql = "SELECT * FROM submissions WHERE pair_id=?"
        args = [pair_id]
        if uik is not None:
            sql += " AND uik=?"
            args.append(uik)
        if current_only:
            sql += " AND current=1"
        records = self.store.query(sql + " ORDER BY id DESC", args)
        for row in records:
            row["payload"] = json.loads(row["payload"])
            row["check_result"] = json.loads(row["check_result"])
        return records

    def ordered(self, records, config):
        return ordered_records(records, config)

    def latest_confirmed(self, pair_id, uik, records=None):
        config = self.store.pair(pair_id)["config"]
        records = records if records is not None else self.history(pair_id, uik, current_only=True)
        return select_latest(records, config)

    def summary(self, project_id):
        pairs = self.store.query("SELECT id,slot FROM pairs WHERE project_id=? AND active=1 ORDER BY slot", (project_id,))
        uiks, by_pair = set(), {}
        for pair in pairs:
            config = self.store.pair(pair["id"])["config"]
            index = self.excel_index(config)
            records = self.history(pair["id"], current_only=True)
            uiks.update(k for k in index if k)
            uiks.update(r["uik"] for r in records)
            by_pair[pair["id"]] = (config, index, records)
        result = []
        for uik in sorted(uiks):
            cells = {}
            for pair in pairs:
                config, index, records = by_pair[pair["id"]]
                group = [r for r in records if r["uik"] == uik]
                if not group:
                    cells[pair["id"]] = {"status": Status.NO_GOOGLE, "confirmed": False, "warnings": [], "last_error": config.get("last_error", "")}
                    continue
                latest = self.ordered(group, config)[-1]
                chosen, warnings = self.latest_confirmed(pair["id"], uik, group)
                cells[pair["id"]] = {"status": latest["check_result"]["status"], "confirmed": bool(chosen),
                                      "warnings": warnings, "last_error": config.get("last_error", "")}
            ready = all(c["confirmed"] and not c["warnings"] and not c.get("last_error") for c in cells.values())
            partial = any(c["confirmed"] for c in cells.values())
            result.append({"uik": uik, "pairs": cells, "total": "Готово" if ready else "Частично" if partial else "Требует внимания"})
        return pairs, result
