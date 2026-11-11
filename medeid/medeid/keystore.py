"""Хранилище ключей «код ↔ пациент» — отдельная БД, отдельно от результатов и журнала.

Схема (SQLite):
  patients(code, identity_hash, fio, birth, created_at, created_by, basis)
  date_shifts(identity_hash, shift_days, created_at)      — только при «хранить_в_ключе»
  access_log(ts, user, action, code, reason)               — каждое раскрытие ключа
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import sqlite3
import stat
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS patients(
  code TEXT PRIMARY KEY,
  identity_hash TEXT UNIQUE NOT NULL,
  fio TEXT NOT NULL,
  birth TEXT,
  created_at TEXT NOT NULL,
  created_by TEXT NOT NULL,
  basis TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS date_shifts(
  identity_hash TEXT PRIMARY KEY,
  shift_days INTEGER NOT NULL,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS access_log(
  ts TEXT NOT NULL, user TEXT NOT NULL, action TEXT NOT NULL, code TEXT, reason TEXT
);
"""


def _now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def identity_hash(identity: str) -> str:
    return hashlib.sha256(identity.encode("utf-8")).hexdigest()


class KeyStore:
    def __init__(self, path: Path, user: str, basis: dict, forbidden_dirs: list[Path] = ()):
        path = path.resolve()
        for d in forbidden_dirs:
            d = d.resolve()
            if d == path.parent or d in path.parents:
                raise RuntimeError(f"хранилище ключей {path} не должно находиться внутри {d}")
        new = not path.exists()
        path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(str(path))
        if new:
            os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)     # 0600
        self.db.executescript(SCHEMA)
        self.user, self.basis = user, json.dumps(basis, ensure_ascii=False, sort_keys=True)
        self.path = path

    def code_for(self, identity: str, fio: str, birth: str | None) -> str:
        h = identity_hash(identity)
        row = self.db.execute("SELECT code, fio FROM patients WHERE identity_hash=?", (h,)).fetchone()
        if row:
            if len(fio) > len(row[1]):        # «Мазуренко Г.В.» → «Мазуренко Галина Викторовна»
                self.db.execute("UPDATE patients SET fio=? WHERE code=?", (fio, row[0]))
                self.db.commit()
            return row[0]
        n = self.db.execute("SELECT COUNT(*) FROM patients").fetchone()[0] + 1
        code = f"ПАЦИЕНТ_{n:04d}"
        self.db.execute("INSERT INTO patients VALUES (?,?,?,?,?,?,?)",
                        (code, h, fio, birth, _now(), self.user, self.basis))
        self.db.execute("INSERT INTO access_log VALUES (?,?,?,?,?)", (_now(), self.user, "создание", code, None))
        self.db.commit()
        return code

    def get_shift(self, identity: str) -> int | None:
        row = self.db.execute("SELECT shift_days FROM date_shifts WHERE identity_hash=?",
                              (identity_hash(identity),)).fetchone()
        return row[0] if row else None

    def put_shift(self, identity: str, days: int):
        self.db.execute("INSERT OR IGNORE INTO date_shifts VALUES (?,?,?)", (identity_hash(identity), days, _now()))
        self.db.commit()

    def reveal(self, code: str, reason: str) -> tuple | None:
        self.db.execute("INSERT INTO access_log VALUES (?,?,?,?,?)", (_now(), self.user, "раскрытие", code, reason))
        self.db.commit()
        return self.db.execute("SELECT code, fio, birth, created_at, created_by FROM patients WHERE code=?",
                               (code,)).fetchone()

    def close(self):
        self.db.close()
