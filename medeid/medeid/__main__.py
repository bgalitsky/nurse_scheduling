"""Командная строка.

  python -m medeid run    -c config.yaml -i исходники/ -o результаты/          → результаты/<партия>/
  python -m medeid verify -c config.yaml -i исходники/ -o результаты/<партия>/ (проверка повторяемости)
  python -m medeid app                                                          (окно приложения)
  python -m medeid reveal --keystore keys.sqlite --code ПАЦИЕНТ_0042 --user petrov.pp --reason "..."
"""
from __future__ import annotations

import argparse
import getpass
import json
import sys
import tempfile
from pathlib import Path

from . import __version__
from .config import ConfigError, load
from .keystore import KeyStore
from .pipeline import Pipeline


def cmd_run(a):
    cfg = load(a.config)
    cfg.run_by = getpass.getuser()          # кто запустил — учётная запись ОС, а не поле конфигурации
    m = Pipeline(cfg).run(Path(a.input), Path(a.output))
    print(f"партия {m['партия']}  отпечаток {m['отпечаток_запуска'][:16]}…  → {m['папка_партии']}")
    for k, v in m["итог"].items():
        print(f"  {k}: {v}")
    print(f"  автопроверка: {m['автопроверка']}")
    if m["итог"].get("требует_проверки"):
        print(f"  ! файлы с замечаниями лежат в {Path(m['папка_партии']) / '_на_проверку'}")
    return 0


def cmd_verify(a):
    from .pipeline import verify_batch
    cfg = load(a.config)
    try:
        r = verify_batch(Path(a.input), Path(a.output), cfg.journal)
    except RuntimeError as e:
        print(e); return 2
    if not r["совпадает_отпечаток"]:
        print("ОТЛИЧАЕТСЯ отпечаток запуска: изменились версии, словари, модель или входные файлы.")
    if not r["различаются"]:
        print(f"ПОВТОРЯЕМО: партия {r['партия']}, {r['всего']} док., результаты побайтно совпадают.")
        return 0
    print(f"НЕ СОВПАДАЮТ {len(r['различаются'])} из {r['всего']}: {', '.join(r['различаются'][:20])}")
    if r["ожидаемо_из_за_сдвига"]:
        print("  Ожидаемо: величина сдвига дат удалена после обработки.")
    return 1


def cmd_check_model(a):
    """Проверка установки модели: сервер, дайджест, пробное извлечение на синтетическом тексте."""
    import yaml
    from .llm import LLMClient, LLMError, LLMSettings
    raw = yaml.safe_load(Path(a.config).read_text(encoding="utf-8")) or {}
    s = LLMSettings.from_raw(raw.get("модель") or {})
    s.cache = None
    try:
        c = LLMClient(s)
        print(f"сервер {s.server}: доступен, адрес внутренний")
        try:
            d = c.verify_model()
            print(f"модель {s.name}: дайджест совпадает ({d[:16]}…)")
        except LLMError as e:
            print(f"модель {s.name}: {e}")
            if "укажите модель.дайджест" not in str(e):
                return 1
        sample = ("ВЫПИСНОЙ ЭПИКРИЗ\nФИО пациента: Тестова Мария Олеговна, тел. +7 900 000-00-00\n"
                  "Адрес: г. Казань, ул. Пробная, д. 1, кв. 2. Консилиум № 77.\n"
                  "Лечащий врач: Проверкин А.А. (кардиолог). Дочь Тестова О.И. проинформирована.\n")
        ents, w = c.extract(sample)
        print("пробное извлечение:")
        for e in ents:
            print(f"  {e.category:18} {e.text}" + (f"  [{e.specialty}]" if e.specialty else ""))
        for x in w:
            print("  !", x)
        print(f"статистика: {c.stats}")
        return 0 if ents else 1
    except LLMError as e:
        print(f"ошибка: {e}"); return 1


def cmd_reveal(a):
    ks = KeyStore(Path(a.keystore), a.user, {})
    row = ks.reveal(a.code, a.reason)
    ks.close()
    if not row:
        print("код не найден"); return 1
    print(f"{row[0]}: {row[1]}, дата рождения {row[2] or '—'} (создан {row[3]}, {row[4]})")
    return 0


def main(argv=None):
    p = argparse.ArgumentParser(prog="medeid", description=f"Обезличивание медицинских PDF, v{__version__}")
    p.add_argument("--version", action="version", version=__version__)
    sub = p.add_subparsers(dest="cmd", required=True)
    for name, fn in (("run", cmd_run), ("verify", cmd_verify)):
        s = sub.add_parser(name)
        s.add_argument("-c", "--config", required=True)
        s.add_argument("-i", "--input", required=True, help="каталог с PDF")
        s.add_argument("-o", "--output", required=True, help="каталог результатов")
        s.set_defaults(fn=fn)
    s = sub.add_parser("app", help="окно приложения (оператор / администратор)")
    s.set_defaults(fn=lambda a: __import__("medeid.app", fromlist=["main"]).main())
    s = sub.add_parser("check-model", help="проверить, что локальная модель установлена и отвечает")
    s.add_argument("-c", "--config", required=True)
    s.set_defaults(fn=cmd_check_model)
    s = sub.add_parser("reveal", help="раскрыть код пациента (запись в журнал доступа)")
    s.add_argument("--keystore", required=True); s.add_argument("--code", required=True)
    s.add_argument("--user", required=True); s.add_argument("--reason", required=True)
    s.set_defaults(fn=cmd_reveal)
    a = p.parse_args(argv)
    try:
        return a.fn(a)
    except ConfigError as e:
        print(e, file=sys.stderr); return 2
    except RuntimeError as e:
        print(f"ошибка: {e}", file=sys.stderr); return 2


if __name__ == "__main__":
    sys.exit(main())
