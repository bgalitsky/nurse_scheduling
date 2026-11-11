"""Загрузка и строгая проверка конфигурации.

Любое расхождение версий или недопустимый режим — ошибка до начала обработки:
повторяемость важнее, чем «как-нибудь отработать».
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from . import __version__

DOC_TYPES = {"выписной_эпикриз", "протокол_консилиума", "результаты_обследования", "осмотр_врача"}

NAME_MODES = {"замена_меткой", "удаление"}
MODES = {
    "фио_пациента": NAME_MODES | {"сквозной_код"},
    "фио_родственников": NAME_MODES,
    "фио_врачей": NAME_MODES | {"специальность"},
    "фио_прочие": NAME_MODES,
    "адрес": {"обобщение_до_города", "замена_меткой", "удаление"},
    "телефон": {"удаление", "замена_меткой"},
    "email": {"удаление", "замена_меткой"},
    "снилс_полис_паспорт": {"удаление", "замена_меткой"},
    "номер_истории": {"удаление", "замена_меткой"},
    "номера_документов": {"замена_меткой", "удаление"},
    "место_работы": {"замена_меткой", "удаление"},
    "даты": {"оставить", "только_год", "квартал", "сдвиг", "замена_меткой", "удаление"},
    "возраст": {"как_есть", "возрастная_группа"},
}
DEFAULTS = {
    "фио_прочие": "замена_меткой",   # ФИО, роль которых не удалось определить
    "номера_документов": "замена_меткой",
    "место_работы": "замена_меткой",
    "возраст": "как_есть",
    "возраст_старше_89": "90+",
}
YES = {"да", "yes", "true", True}
NO = {"нет", "no", "false", False}


class ConfigError(Exception):
    pass


@dataclass
class Config:
    raw: dict
    run_by: str
    basis: dict
    doc_types: list[str]
    modes: dict[str, str]
    over_89: str
    through_code: bool
    shift_range: tuple[int, int]
    shift_after: str          # "удалить" | "хранить_в_ключе"
    keystore: Path | None
    journal: Path
    tool_version: str
    extraction: str           # правила | модель | модель+правила
    llm: object | None        # LLMSettings
    dict_version: str
    dict_dir: Path
    models_dir: Path
    min_text_chars: int = 200
    extra: dict = field(default_factory=dict)

    def processing_fingerprint_part(self) -> dict:
        """Всё, что влияет на результат. Кто запустил и основание — не влияют."""
        return {
            "типы_документов": sorted(self.doc_types),
            "режимы": dict(sorted(self.modes.items())),
            "возраст_старше_89": self.over_89,
            "сквозной_код_пациента": self.through_code,
            "сдвиг_дат": {"диапазон_дней": list(self.shift_range), "после_обработки": self.shift_after},
            "версия_инструмента": self.tool_version,
            "извлечение": self.extraction,
            "модель": self.llm.fingerprint() if self.llm else None,
            "версия_словарей": self.dict_version,
        }

    def sha256(self) -> str:
        blob = json.dumps(self.processing_fingerprint_part(), ensure_ascii=False, sort_keys=True)
        return hashlib.sha256(blob.encode()).hexdigest()


def _bool(v, name):
    if v in YES or (isinstance(v, str) and v.lower() in YES):
        return True
    if v in NO or (isinstance(v, str) and v.lower() in NO):
        return False
    raise ConfigError(f"{name}: ожидается да/нет, получено {v!r}")


def load(path: str | Path, base_dir: Path | None = None) -> Config:
    raw = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    return load_raw(raw, base_dir)


def load_raw(raw: dict, base_dir: Path | None = None) -> Config:
    """Проверка конфигурации, собранной в программе (например, в окне приложения)."""
    raw = json.loads(json.dumps(raw, ensure_ascii=False, default=str))   # копия
    base_dir = base_dir or Path(__file__).resolve().parent.parent
    errs: list[str] = []

    run_by = str(raw.get("запустил") or "").strip()
    if not run_by:
        errs.append("запустил: не указан пользователь")

    basis = raw.get("основание") or {}
    if isinstance(basis, str):
        basis = {"текст": basis}
    if not basis or any(str(v).strip("_ ") == "" for v in basis.values()):
        errs.append("основание: заполните приказ (номер, дата, цель, отделение) — пустые '___' не допускаются")

    doc_types = raw.get("типы_документов") or []
    bad = [t for t in doc_types if t not in DOC_TYPES]
    if not doc_types or bad:
        errs.append(f"типы_документов: допустимы {sorted(DOC_TYPES)}, ошибка в {bad or 'пустой список'}")

    cats = dict(raw.get("категории_и_режимы") or {})
    over_89 = str(cats.pop("возраст_старше_89", DEFAULTS["возраст_старше_89"]))
    modes: dict[str, str] = {}
    for cat, allowed in MODES.items():
        val = cats.pop(cat, DEFAULTS.get(cat))
        if val is None:
            errs.append(f"категории_и_режимы.{cat}: не задан режим (допустимо: {sorted(allowed)})")
        elif val not in allowed:
            errs.append(f"категории_и_режимы.{cat}: режим {val!r} недопустим (допустимо: {sorted(allowed)})")
        else:
            modes[cat] = val
    if cats:
        errs.append(f"категории_и_режимы: неизвестные категории {sorted(cats)}")

    try:
        through = _bool(raw.get("сквозной_код_пациента", "нет"), "сквозной_код_пациента")
    except ConfigError as e:
        errs.append(str(e)); through = False
    if through and modes.get("фио_пациента") != "сквозной_код":
        errs.append("сквозной_код_пациента: да → фио_пациента должно быть 'сквозной_код'")
    if not through and modes.get("фио_пациента") == "сквозной_код":
        errs.append("фио_пациента: сквозной_код требует сквозной_код_пациента: да")

    sd = raw.get("сдвиг_дат") or {}
    rng = tuple(sd.get("диапазон_дней", [-365, 365]))
    if len(rng) != 2 or int(rng[0]) >= int(rng[1]):
        errs.append("сдвиг_дат.диапазон_дней: ожидается [мин, макс], мин < макс")
    shift_after = sd.get("после_обработки", "удалить")
    if shift_after not in {"удалить", "хранить_в_ключе"}:
        errs.append("сдвиг_дат.после_обработки: удалить | хранить_в_ключе")

    needs_keystore = through or (modes.get("даты") == "сдвиг" and shift_after == "хранить_в_ключе")
    ks = raw.get("хранилище_ключей")
    keystore = Path(ks).expanduser() if ks else None
    if needs_keystore and not keystore:
        errs.append("хранилище_ключей: обязательно при сквозном коде или хранении сдвига дат")

    journal = Path(raw.get("журнал") or "journal.sqlite").expanduser()

    tool_v = str(raw.get("версия_инструмента") or "")
    if tool_v != __version__:
        errs.append(f"версия_инструмента: в конфигурации {tool_v!r}, установлена {__version__!r}")
    if "версия_модели" in raw:
        errs.append("версия_модели: в 2.0 заменено на «извлечение» и секцию «модель» (см. config.example.yaml)")
    extraction = str(raw.get("извлечение") or "")
    llm = None
    if extraction not in ("правила", "модель", "модель+правила"):
        errs.append("извлечение: правила | модель | модель+правила")
    elif extraction != "правила":
        from .llm import LLMSettings
        m = raw.get("модель") or {}
        try:
            llm = LLMSettings.from_raw(m)
        except (TypeError, ValueError) as e:
            errs.append(f"модель: {e}")
        else:
            if llm.engine not in ("ollama", "openai"):
                errs.append("модель.движок: ollama | openai")
            if not llm.name:
                errs.append("модель.имя: не указано (например qwen3:8b)")
            if not llm.digest:
                errs.append("модель.дайджест: обязателен для повторяемости; узнать: python -m medeid check-model -c <конфиг>")
            if not (500 <= llm.chunk <= 20000):
                errs.append("модель.размер_фрагмента: 500…20000 символов")
    dict_v = str(raw.get("версия_словарей") or "")
    dict_dir = Path(raw.get("каталог_словарей") or base_dir / "dicts") / dict_v
    if not dict_v or not dict_dir.is_dir():
        errs.append(f"версия_словарей: каталог {dict_dir} не найден")
    models_dir = Path(raw.get("каталог_моделей") or base_dir / "models")

    if errs:
        raise ConfigError("Ошибки конфигурации:\n  - " + "\n  - ".join(errs))

    return Config(
        raw=raw, run_by=run_by, basis=basis, doc_types=list(doc_types), modes=modes,
        over_89=over_89, through_code=through, shift_range=(int(rng[0]), int(rng[1])),
        shift_after=shift_after, keystore=keystore, journal=journal,
        tool_version=tool_v, extraction=extraction, llm=llm, dict_version=dict_v,
        dict_dir=dict_dir, models_dir=models_dir,
        min_text_chars=int(raw.get("минимум_символов_текста", 200)),
    )
