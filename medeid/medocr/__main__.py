"""python -m medocr -i scans/ -o source/   — распознать сканы локально, результат — исходники для medeid.

Параметры по умолчанию можно задать в папке настроек medeid, файл ocr.yaml (те же ключи, что DEFAULTS).
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .core import DEFAULTS, OcrError, process_folder


def _settings_cfg() -> dict:
    try:
        import yaml
        from medeid.settings import settings_dir
        p = settings_dir() / "ocr.yaml"
        if p.exists():
            return yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    except Exception:
        pass
    return {}


def main(argv=None):
    ap = argparse.ArgumentParser(prog="medocr", description=f"Локальное распознавание сканов PDF, v{__version__}")
    ap.add_argument("--version", action="version", version=__version__)
    ap.add_argument("-i", "--input", required=True, help="папка со сканами PDF")
    ap.add_argument("-o", "--output", required=True, help="папка результата (исходники для medeid)")
    ap.add_argument("--tesseract", help="путь к tesseract(.exe)")
    ap.add_argument("--lang", help="языки, по умолчанию rus+eng")
    ap.add_argument("--dpi", type=int)
    ap.add_argument("--threshold", type=float, help="порог средней уверенности страницы, %%")
    ap.add_argument("--force", action="store_true", help="распознавать и страницы с хорошим текстовым слоем")
    a = ap.parse_args(argv)

    cfg = {**DEFAULTS, **_settings_cfg()}
    for key, val in (("tesseract", a.tesseract), ("языки", a.lang), ("dpi", a.dpi), ("порог_уверенности", a.threshold)):
        if val is not None:
            cfg[key] = val

    def show(i, n, r):
        extra = f" — низкая уверенность на стр. {', '.join(map(str, r.low_pages))}" if r.low_pages else ""
        conf = [p.confidence for p in r.pages if p.confidence is not None]
        c = f", уверенность {min(conf):.0f}–{max(conf):.0f}%" if conf else ""
        print(f"[{i}/{n}] {r.output}: {r.status}{c}{extra}{(' — ' + r.note) if r.note else ''}")

    try:
        rep = process_folder(Path(a.input), Path(a.output), cfg, a.force, progress=show)
    except OcrError as e:
        print(f"ошибка: {e}", file=sys.stderr); return 2
    print(f"Tesseract: {rep['tesseract']}; итог: " + ", ".join(f"{k}: {v}" for k, v in rep["итог"].items()))
    print(f"отчёт: {Path(a.output) / 'ocr_report.json'}, соответствие файлов: {Path(a.output) / 'ocr_index.csv'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
