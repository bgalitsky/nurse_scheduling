"""Конвейер партии: документы → текст → поиск → замена → контроль → DOC-xxxxxx.txt + манифест + журнал.

Партия = один запуск с одной конфигурацией и одним основанием. Результаты партии лежат в
<папка результатов>/<номер партии>/. В журнал и манифест попадают только DOC-номера и хеши.
"""
from __future__ import annotations

import hashlib
import json
import re
import secrets
import uuid
from pathlib import Path
from typing import Callable

from .config import Config
from .detect import Detector, Dictionaries, NerBackend, PatientInfo
from .extract import SUPPORTED, ExtractError, detect_doc_type, extract, fold
from .journal import Journal
from .keystore import KeyStore
from .llm import Entity, LLMClient, LLMError
from .registry import Registry
from .transform import Transformer, integrity_check

REVIEW_DIR = "_на_проверку"


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256_dir(d: Path) -> str:
    h = hashlib.sha256()
    if d.is_dir():
        for p in sorted(x for x in d.rglob("*") if x.is_file()):
            h.update(p.relative_to(d).as_posix().encode()); h.update(b"\0")
            h.update(p.read_bytes()); h.update(b"\0")
    return h.hexdigest()


def check_storages(named: dict[str, Path | None]):
    """Четыре раздельных хранилища: ни одно не может лежать внутри другого."""
    items = [(k, Path(v).resolve()) for k, v in named.items() if v]
    for i, (a, pa) in enumerate(items):
        for b, pb in items[i + 1:]:
            if pa == pb or pa in pb.parents or pb in pa.parents:
                raise RuntimeError(f"хранилища «{a}» и «{b}» должны быть раздельными: {pa} / {pb}")


class Cancelled(Exception):
    pass


class Pipeline:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.dicts = Dictionaries.load(cfg.dict_dir)
        self.det = Detector(self.dicts, NerBackend())
        self.dict_sha = sha256_dir(cfg.dict_dir)
        self.llm: LLMClient | None = None
        self.model_digest = None
        self._shifts: dict[str, int] = {}     # в памяти; при «удалить» исчезает после запуска

    # ------------------------------------------------------------------
    def run(self, input_dir: Path, out_root: Path, journal=True, kind="обработка",
            progress: Callable[[int, int, dict], None] | None = None,
            cancel: Callable[[], bool] | None = None) -> dict:
        cfg = self.cfg
        input_dir, out_root = Path(input_dir), Path(out_root)
        if not input_dir.is_dir():
            raise RuntimeError(f"папка исходных документов не найдена: {input_dir}")
        cache = cfg.llm.cache if cfg.llm else None
        check_storages({"исходники": input_dir, "результаты": out_root,
                        "ключи": cfg.keystore.parent if cfg.keystore else None,
                        "настройки": cfg.dict_dir.parent.parent})
        if cache and cfg.keystore and cache.parent != cfg.keystore.parent:
            check_storages({"исходники": input_dir, "результаты": out_root, "кэш модели": cache.parent})

        reg = Registry(input_dir, cfg.run_by)
        inputs: dict[str, Path] = {}
        for p in sorted(input_dir.rglob("*")):
            if ".medeid" in p.parts:
                continue
            if p.is_file() and p.suffix.lower() in SUPPORTED:
                inputs.setdefault(sha256_file(p), p)       # дубликаты по содержимому — один раз
        order = sorted(inputs)                               # порядок не зависит от имён файлов
        doc_ids = {h: reg.doc_id(h, inputs[h]) for h in order}

        if cfg.llm:
            try:
                self.llm = LLMClient(cfg.llm, forbidden_dirs=[out_root, input_dir])
                self.model_digest = self.llm.verify_model()
            except LLMError as e:
                reg.close()
                raise RuntimeError(f"модель: {e}")
        fingerprint = hashlib.sha256(json.dumps(
            [cfg.sha256(), self.dict_sha, self.model_digest, order]).encode()).hexdigest()
        run_id = str(uuid.uuid4())

        ks = None
        if cfg.keystore and (cfg.through_code or (cfg.modes["даты"] == "сдвиг" and cfg.shift_after == "хранить_в_ключе")):
            ks = KeyStore(cfg.keystore, cfg.run_by, cfg.basis, forbidden_dirs=[out_root, input_dir])
        jr = Journal(cfg.journal) if journal else None
        batch = jr.new_batch_id() if jr else "без-журнала"
        out_dir = out_root / batch
        if out_dir.exists() and any(out_dir.iterdir()):
            raise RuntimeError(f"папка партии {out_dir} уже существует и не пуста")
        out_dir.mkdir(parents=True, exist_ok=True)
        if jr:
            jr.start(run_id, cfg, fingerprint, self.dict_sha, batch, kind)
            if reg.new_ids:
                jr.event(cfg.run_by, "регистрация_документов", batch,
                         {"документы": reg.new_ids, "количество": len(reg.new_ids)})
            if cfg.through_code:
                jr.event(cfg.run_by, "сквозной_код_включён", batch,
                         {"хранилище_ключей": "задано", "основание": cfg.basis})

        records, status = [], "завершён"
        try:
            for i, h in enumerate(order, 1):
                if cancel and cancel():
                    status = "прерван пользователем"
                    break
                try:
                    rec = self._one(inputs[h], h, doc_ids[h], out_dir, ks, reg)
                except Exception as e:                  # без текста ошибки: в нём может быть фрагмент документа
                    rec = {"input_sha256": h, "doc_id": doc_ids[h], "status": "ошибка_обработки",
                           "warnings": [f"внутренняя ошибка {type(e).__name__}"], "counts": {}}
                records.append(rec)
                if jr:
                    jr.file(run_id, rec)
                if progress:
                    progress(i, len(order), rec)
        finally:
            if ks:
                ks.close()
            if self.llm:
                self.llm.close()
            reg.close()
            self._shifts.clear()                             # величина сдвига удаляется

        summary: dict[str, int] = {}
        for r in records:
            summary[r["status"]] = summary.get(r["status"], 0) + 1
        problems = sum(v for k, v in summary.items()
                       if k not in ("готово", "отклонён_тип_документа", "исключён_отзыв_согласия"))
        auto_check = "пройдена" if problems == 0 and records else f"замечания: {problems} док."
        manifest = {
            "партия": batch, "run_id": run_id, "вид": kind, "запустил": cfg.run_by, "основание": cfg.basis,
            "отпечаток_запуска": fingerprint, "config_sha256": cfg.sha256(),
            "версия_инструмента": cfg.tool_version, "извлечение": cfg.extraction,
            "модель": ({**cfg.llm.fingerprint(), "дайджест_на_сервере": self.model_digest,
                        "версия_сервера": self.llm.server_version,
                        "статистика": self.llm.stats} if self.llm else None),
            "версия_словарей": cfg.dict_version,
            "dict_sha256": self.dict_sha, "параметры": cfg.processing_fingerprint_part(),
            "сдвиг_дат_сохранён": cfg.modes["даты"] == "сдвиг" and cfg.shift_after == "хранить_в_ключе",
            "статус": status, "автопроверка": auto_check, "итог": summary, "файлы": records,
        }
        (out_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2, default=str),
                                               encoding="utf-8")
        if jr:
            jr.finish(run_id, status, summary, auto_check)
            jr.close()
        manifest["папка_партии"] = str(out_dir)
        return manifest

    # ------------------------------------------------------------------
    def _one(self, path: Path, h: str, doc_id: str, out_dir: Path, ks: KeyStore | None, reg: Registry) -> dict:
        cfg = self.cfg
        rec = {"input_sha256": h, "doc_id": doc_id, "status": "", "warnings": [], "counts": {}}
        try:
            ex = extract(path, cfg.min_text_chars)
        except ExtractError as e:
            rec.update(status="отклонён", warnings=[str(e)]); return rec
        rec.update(pages_in=ex.pages, pages_text=ex.pages_with_text, images=ex.images)
        text = ex.text
        ftext = fold(text)                     # для поиска; замены применяются к исходному тексту

        doc_type = detect_doc_type(ftext)
        rec["doc_type"] = doc_type
        if doc_type not in cfg.doc_types:
            rec["status"] = "отклонён_тип_документа"; return rec

        warnings = list(ex.warnings)
        entities = None
        if self.llm:
            ents, w = self.llm.extract(text)
            warnings += w
            entities = [Entity(fold(e.text), e.category, e.specialty) for e in ents]
        patient = self.det.find_patient(ftext)
        if patient is None and entities:
            patient = self.det.patient_from_llm(ftext, entities)

        if patient and reg.is_excluded(patient.identity):
            rec.update(status="исключён_отзыв_согласия"); return rec

        needs_patient = cfg.through_code or cfg.modes["даты"] == "сдвиг"
        if patient is None:
            if needs_patient:
                rec.update(status="отклонён_пациент_не_определён",
                           warnings=["не найдено поле ФИО пациента — нельзя присвоить код/сдвиг"]); return rec
            warnings.append("пациент не определён: падежные формы фамилии не отслеживались; "
                            "проверка отзыва согласия невозможна")
        elif not (patient.birth_date or patient.birth_year):
            warnings.append("дата рождения не найдена: ключ пациента только по ФИО (риск совпадения однофамильцев)")

        code = None
        if cfg.through_code:
            code = ks.code_for(patient.identity, patient.display,
                               patient.birth_date.isoformat() if patient.birth_date else
                               (str(patient.birth_year) if patient.birth_year else None))
        rec["patient_code"] = code

        shift = self._shift_for(patient, ks) if cfg.modes["даты"] == "сдвиг" else None

        spans, w = self.det.detect(ftext, patient, entities, use_rules=cfg.extraction != "модель")
        warnings += w
        dates = [s.data["дата"] for s in spans if s.category == "даты" and not s.data.get("рождение")
                 and s.sub != "гггг" and not s.data.get("некорректная")]
        ref = max(dates) if dates else None

        tr = Transformer(cfg.modes, cfg.over_89, code, shift, ref)
        out_text, counts = tr.apply(text, spans)
        warnings += integrity_check(text, out_text, spans, tr.replaced)
        warnings += self._residual(fold(out_text), patient)

        name = f"{doc_id}.txt"
        target_dir = out_dir / REVIEW_DIR if warnings else out_dir
        target_dir.mkdir(exist_ok=True)
        data = out_text.encode("utf-8")
        (target_dir / name).write_bytes(data)
        rec.update(status="требует_проверки" if warnings else "готово",
                   output_name=(f"{REVIEW_DIR}/{name}" if warnings else name),
                   output_sha256=hashlib.sha256(data).hexdigest(),
                   counts=counts, warnings=warnings)
        return rec

    def _shift_for(self, patient: PatientInfo, ks: KeyStore | None) -> int:
        ident = patient.identity
        if ident in self._shifts:
            return self._shifts[ident]
        if ks and self.cfg.shift_after == "хранить_в_ключе":
            stored = ks.get_shift(ident)
            if stored is not None:
                self._shifts[ident] = stored; return stored
        lo, hi = self.cfg.shift_range
        rnd = secrets.SystemRandom()
        days = 0
        while days == 0:
            days = rnd.randint(lo, hi)
        self._shifts[ident] = days
        if ks and self.cfg.shift_after == "хранить_в_ключе":
            ks.put_shift(ident, days)
        return days

    def _residual(self, out_text: str, patient: PatientInfo | None) -> list[str]:
        """Повторный поиск по уже обезличенному тексту: всё найденное — повод для ручной проверки."""
        spans, _ = self.det.detect(out_text, patient)        # только правила: независимая проверка
        found: dict[str, int] = {}
        date_mode = self.cfg.modes["даты"]
        for s in spans:
            frag = out_text[s.start:s.end].strip()
            if re.fullmatch(r"(?:\[[^\[\]]+\][\s,;]*)+", frag):
                continue                                   # это уже наша метка
            if s.category == "даты":
                if date_mode in ("оставить", "сдвиг") or s.sub in ("гггг", "месяц гггг", "мм.гггг", "мм-мм.гггг"):
                    continue
            if s.category == "возраст" and s.data["лет"] <= 89:
                continue
            if s.category == "адрес" and out_text[s.start:s.end] == (s.data.get("город") or ""):
                continue
            found[s.category] = found.get(s.category, 0) + 1
        return [f"остаток после обработки: {k} ×{v}" for k, v in sorted(found.items())]


# ---------------------------------------------------------------- операции с готовыми партиями

def delete_sources(cfg: Config, input_dir: Path, batch: str, user: str) -> int:
    """Удалить исходные файлы партии из хранилища исходников (после обработки) с записью в журнал."""
    jr = Journal(cfg.journal)
    reg = Registry(input_dir, user)
    ids = [r[0] for r in jr.files_of(batch) if r[0]]
    deleted = []
    for did in ids:
        p = reg.path_of(did)
        if p and p.exists():
            p.unlink()
            reg.mark_deleted(did)
            deleted.append(did)
    jr.event(user, "удаление_исходников", batch, {"документы": deleted, "количество": len(deleted)})
    reg.close(); jr.close()
    return len(deleted)


def delete_outputs(cfg: Config, out_root: Path, batch: str, user: str, reason: str) -> int:
    d = out_root / batch
    n = 0
    if d.is_dir():
        for p in sorted(d.rglob("*"), reverse=True):
            if p.is_file():
                p.unlink(); n += 1
            elif p.is_dir():
                p.rmdir()
        d.rmdir()
    jr = Journal(cfg.journal)
    jr.event(user, "удаление_набора", batch, {"файлов": n, "основание": reason})
    jr.close()
    return n


def verify_batch(input_dir: Path, batch_dir: Path, journal_path: Path, base_dir: Path | None = None) -> dict:
    """Повторить партию с её же конфигурацией (из журнала) во временной папке и сравнить побайтно."""
    import tempfile
    from .config import load_raw
    mf = Path(batch_dir) / "manifest.json"
    if not mf.exists():
        raise RuntimeError(f"в папке {batch_dir} нет manifest.json — это не папка партии")
    prev = json.loads(mf.read_text(encoding="utf-8"))
    jr = Journal(journal_path)
    row = jr.config_of(prev["партия"])
    jr.close()
    if not row or not row[0]:
        raise RuntimeError(f"в журнале нет конфигурации партии {prev['партия']}")
    cfg = load_raw(json.loads(row[0]), base_dir)
    with tempfile.TemporaryDirectory() as tmp:
        m = Pipeline(cfg).run(Path(input_dir), Path(tmp), kind="проверка_повторяемости")
    old = {r["input_sha256"]: r.get("output_sha256") for r in prev["файлы"]}
    new = {r["input_sha256"]: r.get("output_sha256") for r in m["файлы"]}
    id_of = {r["input_sha256"]: r.get("doc_id") for r in prev["файлы"] + m["файлы"]}
    diff = sorted(id_of[h] or h[:12] for h in old.keys() | new.keys() if old.get(h) != new.get(h))
    expected = cfg.modes["даты"] == "сдвиг" and cfg.shift_after == "удалить"
    return {"партия": prev["партия"], "проверка": m["партия"], "всего": len(old),
            "совпадает_отпечаток": m["отпечаток_запуска"] == prev["отпечаток_запуска"],
            "различаются": diff, "ожидаемо_из_за_сдвига": expected and bool(diff)}
