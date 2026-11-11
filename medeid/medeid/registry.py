"""Реестр хранилища исходников: внутренний номер документа ↔ исходный файл, отзыв согласия.

Хранится ВНУТРИ папки исходников (<исходники>/.medeid/registry.sqlite), потому что содержит
имена исходных файлов (в них бывают ФИО и номера историй). В журнал попадают только DOC-номера.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import sqlite3
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS docs(
  doc_id TEXT PRIMARY KEY, sha256 TEXT UNIQUE NOT NULL, original_name TEXT NOT NULL,
  added_at TEXT NOT NULL, added_by TEXT NOT NULL, deleted_at TEXT, deleted_by TEXT
);
CREATE TABLE IF NOT EXISTS excluded_patients(
  identity_hash TEXT PRIMARY KEY, added_at TEXT NOT NULL, added_by TEXT NOT NULL, basis TEXT
);
"""


def _now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def identity_hash(identity: str) -> str:
    return hashlib.sha256(("medeid-exclusion|" + identity).encode()).hexdigest()


class Registry:
    def __init__(self, source_dir: Path, user: str):
        d = source_dir / ".medeid"
        d.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(str(d / "registry.sqlite"))
        self.db.executescript(SCHEMA)
        self.user = user
        self.source_dir = source_dir
        self.new_ids: list[str] = []          # зарегистрированы в этом запуске

    def doc_id(self, sha: str, path: Path) -> str:
        row = self.db.execute("SELECT doc_id FROM docs WHERE sha256=?", (sha,)).fetchone()
        if row:
            return row[0]
        n = self.db.execute("SELECT COUNT(*) FROM docs").fetchone()[0] + 1
        did = f"DOC-{n:06d}"
        self.db.execute("INSERT INTO docs(doc_id, sha256, original_name, added_at, added_by) VALUES (?,?,?,?,?)",
                        (did, sha, str(path.relative_to(self.source_dir)), _now(), self.user))
        self.db.commit()
        self.new_ids.append(did)
        return did

    def path_of(self, doc_id: str) -> Path | None:
        row = self.db.execute("SELECT original_name, deleted_at FROM docs WHERE doc_id=?", (doc_id,)).fetchone()
        if not row or row[1]:
            return None
        return self.source_dir / row[0]

    def mark_deleted(self, doc_id: str):
        self.db.execute("UPDATE docs SET deleted_at=?, deleted_by=? WHERE doc_id=?", (_now(), self.user, doc_id))
        self.db.commit()

    # ---------- отзыв согласия
    def exclude(self, identity: str, basis: str):
        self.db.execute("INSERT OR IGNORE INTO excluded_patients VALUES (?,?,?,?)",
                        (identity_hash(identity), _now(), self.user, basis))
        self.db.commit()

    def is_excluded(self, identity: str) -> bool:
        return self.db.execute("SELECT 1 FROM excluded_patients WHERE identity_hash=?",
                               (identity_hash(identity),)).fetchone() is not None

    def excluded_count(self) -> int:
        return self.db.execute("SELECT COUNT(*) FROM excluded_patients").fetchone()[0]

    def close(self):
        self.db.close()
