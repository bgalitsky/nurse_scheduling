"""Журнал учёта (приказ РКН № 140, п. 1.7). Без персональных данных: партии, DOC-номера, хеши,
версии, режимы, результаты проверок и дальнейшая судьба наборов (выдача, копирование, удаление)."""
from __future__ import annotations

import csv
import datetime as dt
import io
import json
import sqlite3
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs(
  run_id TEXT PRIMARY KEY, started_at TEXT, finished_at TEXT, run_by TEXT, basis TEXT,
  config_sha256 TEXT, fingerprint TEXT, tool_version TEXT, model_version TEXT,
  dict_version TEXT, dict_sha256 TEXT, status TEXT, summary TEXT
);
CREATE TABLE IF NOT EXISTS files(
  run_id TEXT, input_sha256 TEXT, output_name TEXT, output_sha256 TEXT, doc_type TEXT,
  patient_code TEXT, counts TEXT, warnings TEXT, status TEXT
);
CREATE TABLE IF NOT EXISTS events(
  ts TEXT NOT NULL, user TEXT NOT NULL, event TEXT NOT NULL, batch TEXT, details TEXT
);
"""
# столбцы, добавленные в 2.1 (миграция старого журнала)
MIGRATIONS = {
    "runs": {"batch": "TEXT", "kind": "TEXT", "config_json": "TEXT", "auto_check": "TEXT",
             "manual_check": "TEXT", "manual_check_by": "TEXT", "manual_check_at": "TEXT",
             "manual_check_note": "TEXT"},
    "files": {"doc_id": "TEXT", "pages_in": "INTEGER", "pages_text": "INTEGER", "images": "INTEGER"},
}

EVENT_TYPES = ("выдача_набора", "копирование_набора", "удаление_набора", "удаление_исходников",
               "отзыв_согласия", "ручная_проверка", "раскрытие_ключа")


def _now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


class Journal:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(str(path))
        self.db.executescript(SCHEMA)
        for table, cols in MIGRATIONS.items():
            have = {r[1] for r in self.db.execute(f"PRAGMA table_info({table})")}
            for c, t in cols.items():
                if c not in have:
                    self.db.execute(f"ALTER TABLE {table} ADD COLUMN {c} {t}")
        self.db.commit()

    # ---------- партии
    def new_batch_id(self) -> str:
        day = dt.date.today().isoformat()
        n = self.db.execute("SELECT COUNT(*) FROM runs WHERE batch LIKE ?", (day + "-%",)).fetchone()[0] + 1
        return f"{day}-{n:03d}"

    def start(self, run_id, cfg, fingerprint, dict_sha, batch, kind):
        self.db.execute(
            "INSERT INTO runs(run_id, started_at, run_by, basis, config_sha256, fingerprint, tool_version,"
            " model_version, dict_version, dict_sha256, status, batch, kind, config_json)"
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                run_id, _now(), cfg.run_by, json.dumps(cfg.basis, ensure_ascii=False),
                cfg.sha256(), fingerprint, cfg.tool_version,
                (f"{cfg.llm.name}@{cfg.llm.digest[:16]}" if cfg.llm else "правила"),
                cfg.dict_version, dict_sha, "выполняется", batch, kind,
                json.dumps(cfg.raw, ensure_ascii=False, default=str)))   # конфигурация без ПДн
        self.db.commit()

    def file(self, run_id, rec: dict):
        self.db.execute(
            "INSERT INTO files(run_id, input_sha256, output_name, output_sha256, doc_type, patient_code,"
            " counts, warnings, status, doc_id, pages_in, pages_text, images) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                run_id, rec["input_sha256"], rec.get("output_name"), rec.get("output_sha256"),
                rec.get("doc_type"), rec.get("patient_code"),
                json.dumps(rec.get("counts", {}), ensure_ascii=False),
                json.dumps(rec.get("warnings", []), ensure_ascii=False), rec["status"],
                rec.get("doc_id"), rec.get("pages_in"), rec.get("pages_text"), rec.get("images")))
        self.db.commit()

    def finish(self, run_id, status, summary, auto_check):
        self.db.execute("UPDATE runs SET finished_at=?, status=?, summary=?, auto_check=? WHERE run_id=?",
                        (_now(), status, json.dumps(summary, ensure_ascii=False), auto_check, run_id))
        self.db.commit()

    def set_manual_check(self, batch, user, ok: bool, note: str):
        self.db.execute("UPDATE runs SET manual_check=?, manual_check_by=?, manual_check_at=?, manual_check_note=?"
                        " WHERE batch=?", ("пройдена" if ok else "не пройдена", user, _now(), note, batch))
        self.event(user, "ручная_проверка", batch, {"результат": "пройдена" if ok else "не пройдена",
                                                     "комментарий": note})

    # ---------- события
    def event(self, user, event, batch=None, details=None):
        self.db.execute("INSERT INTO events VALUES (?,?,?,?,?)",
                        (_now(), user, event, batch, json.dumps(details or {}, ensure_ascii=False)))
        self.db.commit()

    # ---------- чтение (окно администратора)
    def runs(self):
        cur = self.db.execute(
            "SELECT batch, kind, started_at, run_by, basis, status, summary, auto_check, manual_check,"
            " tool_version, model_version, dict_version, run_id FROM runs ORDER BY started_at DESC")
        return cur.fetchall()

    def files_of(self, batch):
        return self.db.execute(
            "SELECT f.doc_id, f.doc_type, f.status, substr(f.input_sha256,1,16), substr(f.output_sha256,1,16),"
            " f.pages_in, f.pages_text, f.images, f.counts, f.warnings, f.patient_code"
            " FROM files f JOIN runs r ON r.run_id=f.run_id WHERE r.batch=? ORDER BY f.doc_id", (batch,)).fetchall()

    def config_of(self, batch):
        r = self.db.execute("SELECT config_json, basis, fingerprint FROM runs WHERE batch=?", (batch,)).fetchone()
        return r

    def events(self, kinds=None):
        q = "SELECT ts, user, event, batch, details FROM events"
        args = ()
        if kinds:
            q += " WHERE event IN (%s)" % ",".join("?" * len(kinds)); args = tuple(kinds)
        return self.db.execute(q + " ORDER BY ts", args).fetchall()

    def export_csv(self, kinds=None) -> str:
        buf = io.StringIO()
        w = csv.writer(buf, delimiter=";")
        w.writerow(["дата_время_utc", "пользователь", "событие", "партия", "подробности"])
        for row in self.events(kinds):
            w.writerow(row)
        return buf.getvalue()

    def close(self):
        self.db.close()
