"""SQLite: короткие транзакции, история версий и восстановление операций."""
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from .domain import dumps


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


SCHEMA = """
CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY, name TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS pairs(
 id INTEGER PRIMARY KEY, project_id INTEGER NOT NULL REFERENCES projects(id),
 slot INTEGER NOT NULL CHECK(slot BETWEEN 1 AND 3), active INTEGER NOT NULL DEFAULT 1,
 config TEXT NOT NULL DEFAULT '{}', UNIQUE(project_id,slot));
CREATE TABLE IF NOT EXISTS submissions(
 id INTEGER PRIMARY KEY, pair_id INTEGER NOT NULL REFERENCES pairs(id),
 source_key TEXT NOT NULL, revision INTEGER NOT NULL, uik TEXT NOT NULL,
 fingerprint TEXT NOT NULL, payload TEXT NOT NULL, source_time TEXT,
 position INTEGER NOT NULL, current INTEGER NOT NULL DEFAULT 1,
 present INTEGER NOT NULL DEFAULT 1, observed_at TEXT NOT NULL,
 confirmed_at TEXT, confirmation_hash TEXT, check_result TEXT NOT NULL,
 UNIQUE(pair_id,source_key,revision));
CREATE INDEX IF NOT EXISTS submissions_lookup ON submissions(pair_id,uik,current);
CREATE UNIQUE INDEX IF NOT EXISTS submissions_current ON submissions(pair_id,source_key) WHERE current=1;
CREATE TABLE IF NOT EXISTS events(
 id INTEGER PRIMARY KEY, at TEXT NOT NULL, pair_id INTEGER, kind TEXT NOT NULL, detail TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS operations(
 id INTEGER PRIMARY KEY, pair_id INTEGER NOT NULL, started TEXT NOT NULL,
 finished TEXT, state TEXT NOT NULL, error TEXT);
PRAGMA user_version=1;
"""


class Store:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.path = self.root / "state.sqlite3"
        with self.transaction() as db:
            if db.execute("PRAGMA user_version").fetchone()[0] > 1:
                raise RuntimeError("База создана более новой версией программы. Откройте её соответствующей версией.")
            db.executescript(SCHEMA)

    @contextmanager
    def transaction(self):
        db = sqlite3.connect(self.path, timeout=30)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys=ON")
        db.execute("PRAGMA journal_mode=WAL")
        db.execute("PRAGMA synchronous=FULL")
        try:
            db.execute("BEGIN IMMEDIATE")
            yield db
            db.commit()
        except BaseException:
            db.rollback()
            raise
        finally:
            db.close()

    def query(self, sql, args=()):
        with self.transaction() as db:
            return [dict(r) for r in db.execute(sql, args)]

    def create_project(self, name):
        name = name.strip()
        if not name:
            raise ValueError("Укажите название проекта")
        with self.transaction() as db:
            pid = db.execute("INSERT INTO projects(name) VALUES(?)", (name,)).lastrowid
            for slot in range(1, 4):
                db.execute("INSERT INTO pairs(project_id,slot,active) VALUES(?,?,?)", (pid, slot, int(slot == 1)))
            return pid

    def pair(self, pair_id):
        rows = self.query("SELECT * FROM pairs WHERE id=?", (pair_id,))
        if not rows:
            raise ValueError("Пара не найдена")
        row = rows[0]
        row["config"] = json.loads(row["config"])
        return row

    def save_config(self, pair_id, config, active=None):
        with self.transaction() as db:
            self.write_config(db, pair_id, config)
            if active is not None:
                db.execute("UPDATE pairs SET active=? WHERE id=?", (int(active), pair_id))
            self.event(db, pair_id, "settings", "Настройки пары сохранены")

    @staticmethod
    def write_config(db, pair_id, config):
        db.execute("UPDATE pairs SET config=? WHERE id=?", (dumps(config), pair_id))

    @staticmethod
    def event(db, pair_id, kind, detail):
        db.execute("INSERT INTO events(at,pair_id,kind,detail) VALUES(?,?,?,?)", (now(), pair_id, kind, detail))

    def recover(self):
        with self.transaction() as db:
            pending = db.execute("SELECT id,pair_id FROM operations WHERE state='running'").fetchall()
            for op in pending:
                db.execute("UPDATE operations SET state='interrupted',finished=?,error=? WHERE id=?",
                           (now(), "Чтение прервано; прежний снимок сохранён", op["id"]))
                self.event(db, op["pair_id"], "interrupted", "Прерванное чтение не считается успешным")
