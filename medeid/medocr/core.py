"""Распознавание сканов локальным Tesseract.

Для каждой страницы решается, что с ней делать:
  text     — хороший текстовый слой, крупных изображений нет → страница копируется без изменений;
  scan     — текста нет (скан) → распознаётся;
  mixed    — есть текст и крупное изображение (вклеенный скан, снимок с подписью) → распознаётся целиком,
             чтобы текст внутри изображения тоже попал в текстовый слой;
  garbled  — текстовый слой «иероглифами» (нестандартный шрифт) → распознаётся как изображение.

Распознанная страница = изображение страницы + невидимый текстовый слой Tesseract.
Для каждой страницы сохраняется уверенность распознавания; страницы ниже порога записываются
в метаданные PDF (/MedocrLowConfidence), medeid по ним отправляет документ на ручную проверку.

Данные не покидают компьютер: Tesseract вызывается локально, временные изображения страниц
создаются внутри папки результатов (защищённое хранилище исходников) и удаляются сразу после страницы.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Callable

from . import __version__

DEFAULTS = {
    "tesseract": "tesseract",       # путь к tesseract.exe, если его нет в PATH
    "языки": "rus+eng",
    "dpi": 300,
    "порог_уверенности": 70,        # средняя уверенность страницы, %, ниже — на ручную проверку
    "доля_изображения": 0.15,       # изображение больше этой доли страницы → страница «смешанная» сразу
    "мин_доля_проверки": 0.01,      # изображения от этой доли страницы проверяются: есть ли в них текст
    "слов_в_изображении": 3,        # столько распознанных слов в изображении → страница «смешанная»
    "минимум_символов": 50,         # меньше символов в текстовом слое → страница считается сканом
    "sha256_языков": {},            # {"rus": "<sha256 rus.traineddata>"} — проверка источника (4.3)
}


class OcrError(Exception):
    pass


@dataclass
class PageReport:
    page: int
    action: str                      # text | scan | mixed | garbled
    confidence: float | None = None  # средняя уверенность Tesseract, %
    words: int = 0
    low_words: int = 0               # слов с уверенностью < 50
    rotation: int = 0


@dataclass
class DocReport:
    output: str
    status: str                      # ok | review | copied | error
    pages: list[PageReport] = field(default_factory=list)
    low_pages: list[int] = field(default_factory=list)
    note: str = ""


# ---------------------------------------------------------------- Tesseract

class Tesseract:
    def __init__(self, exe: str = "tesseract", langs: str = "rus+eng", expected_sha: dict | None = None):
        self.exe = shutil.which(exe) or exe
        self.langs = langs
        try:
            out = subprocess.run([self.exe, "--version"], capture_output=True, text=True, timeout=30)
        except (OSError, subprocess.TimeoutExpired):
            raise OcrError(f"Tesseract не найден ({exe}). Установите его (см. medocr/INSTALL_OCR.md) "
                           f"или укажите путь: --tesseract C:\\Program Files\\Tesseract-OCR\\tesseract.exe")
        self.version = (out.stdout or out.stderr).splitlines()[0].strip()
        have = set(self._run(["--list-langs"]).splitlines()[1:])
        need = set(langs.split("+"))
        if need - have:
            raise OcrError(f"нет языковых данных Tesseract: {', '.join(sorted(need - have))} "
                           f"(нужны файлы {', '.join(sorted(need - have))}.traineddata)")
        self.has_osd = "osd" in have
        self.tessdata = self._tessdata_dir()
        self.lang_sha = {}
        for lang in sorted(need | ({"osd"} if self.has_osd else set())):
            p = self.tessdata / f"{lang}.traineddata" if self.tessdata else None
            if p and p.exists():
                self.lang_sha[lang] = hashlib.sha256(p.read_bytes()).hexdigest()
        for lang, sha in (expected_sha or {}).items():
            got = self.lang_sha.get(lang)
            if not got or not got.startswith(str(sha).lower()):
                raise OcrError(f"контрольная сумма {lang}.traineddata не совпадает с утверждённой "
                               f"(есть {str(got)[:16]}…, ожидается {str(sha)[:16]}…)")

    def _run(self, args, timeout=600) -> str:
        r = subprocess.run([self.exe, *args], capture_output=True, text=True, timeout=timeout)
        return r.stdout + r.stderr

    def _tessdata_dir(self) -> Path | None:
        txt = self._run(["--list-langs"])
        m = re.search(r'"(.+?)"', txt.splitlines()[0]) if txt else None
        return Path(m.group(1)) if m else None

    def orientation(self, img: Path) -> int:
        """Поворот страницы по OSD (0/90/180/270), если есть osd.traineddata."""
        if not self.has_osd:
            return 0
        out = self._run([str(img), "stdout", "--psm", "0"], timeout=120)
        m = re.search(r"Rotate:\s*(\d+)", out)
        c = re.search(r"Orientation confidence:\s*([\d.]+)", out)
        if m and c and float(c.group(1)) >= 2.0:
            return int(m.group(1)) % 360
        return 0

    def words(self, img: Path, dpi: int) -> list[tuple[str, float]]:
        """Только TSV: есть ли текст в изображении (быстрая проверка вклеенных картинок)."""
        base = img.with_suffix("")
        self._run([str(img), str(base), "-l", self.langs, "--dpi", str(dpi), "--psm", "11", "tsv"])
        return self._read_tsv(base.with_suffix(".tsv"))

    @staticmethod
    def _read_tsv(tsv: Path) -> list[tuple[str, float]]:
        words = []
        if tsv.exists():
            rows = csv.DictReader(io.StringIO(tsv.read_text(encoding="utf-8", errors="replace")),
                                  delimiter="\t", quoting=csv.QUOTE_NONE)
            for r in rows:
                t = (r.get("text") or "").strip()
                try:
                    conf = float(r.get("conf") or -1)
                except ValueError:
                    conf = -1
                if t and conf >= 0:
                    words.append((t, conf))
            tsv.unlink(missing_ok=True)
        return words

    def ocr(self, img: Path, dpi: int) -> tuple[bytes, list[tuple[str, float]]]:
        """Один проход Tesseract: PDF страницы (изображение + невидимый текст) и TSV с уверенностями."""
        base = img.with_suffix("")
        self._run([str(img), str(base), "-l", self.langs, "--dpi", str(dpi), "--psm", "3",
                   "-c", "textonly_pdf=0", "pdf", "tsv"])
        pdf, tsv = base.with_suffix(".pdf"), base.with_suffix(".tsv")
        if not pdf.exists():
            raise OcrError("Tesseract не создал страницу")
        words = self._read_tsv(tsv)
        data = pdf.read_bytes()
        pdf.unlink(missing_ok=True)
        return data, words


# ---------------------------------------------------------------- анализ страниц

def _garbled(t: str) -> bool:
    letters = [c for c in t if c.isalpha()]
    if len(letters) < 50:
        return False
    ok = sum(1 for c in letters if "а" <= c.lower() <= "я" or c.lower() == "ё" or "a" <= c.lower() <= "z")
    return ok / len(letters) < 0.6


def classify_pages(path: Path, min_chars: int, image_share: float, check_share: float):
    """Действие для каждой страницы и изображения, которые надо проверить на наличие текста.
    Возвращает (actions, candidates): candidates[i] — список bbox (x0, top, x1, bottom) в пунктах."""
    import pdfplumber
    actions, candidates = [], []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            area = float(page.width * page.height) or 1.0
            img_area, boxes = 0.0, []
            for im in page.images:
                x0, x1 = max(im["x0"], 0), min(im["x1"], page.width)
                top, bottom = max(im["top"], 0), min(im["bottom"], page.height)
                a = max(0.0, x1 - x0) * max(0.0, bottom - top)
                img_area += a
                if a / area >= check_share:
                    boxes.append((float(x0), float(top), float(x1), float(bottom)))
            share = img_area / area
            n = len(text.strip())
            if n < min_chars:
                actions.append("scan")
            elif _garbled(text):
                actions.append("garbled")
            elif share >= image_share:
                actions.append("mixed")
            else:
                actions.append("text")
            candidates.append(boxes if actions[-1] == "text" else [])
    return actions, candidates


# ---------------------------------------------------------------- имена файлов

_TR = dict(zip("абвгдеёжзийклмнопрстуфхцчшщъыьэюя",
               ["a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s",
                "t", "u", "f", "kh", "ts", "ch", "sh", "shch", "", "y", "", "e", "yu", "ya"]))


def ascii_name(name: str) -> str:
    """Имя файла только латиницей (правило проекта): транслитерация + безопасные символы."""
    out = []
    for ch in name:
        low = ch.lower()
        if low in _TR:
            t = _TR[low]
            out.append(t.capitalize() if ch.isupper() and t else t)
        else:
            out.append(ch)
    s = unicodedata.normalize("NFKD", "".join(out)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("._")
    return s or "document"


# ---------------------------------------------------------------- документ

def process_pdf(src: Path, dst: Path, tess: Tesseract, cfg: dict, work_dir: Path,
                force: bool = False) -> DocReport:
    import pypdfium2 as pdfium
    from pypdf import PdfReader, PdfWriter

    rep = DocReport(output=dst.name, status="ok")
    try:
        reader = PdfReader(str(src))
        if reader.is_encrypted:
            raise OcrError("PDF защищён паролем")
        actions, candidates = classify_pages(src, int(cfg["минимум_символов"]), float(cfg["доля_изображения"]),
                                             float(cfg["мин_доля_проверки"]))
    except OcrError:
        raise
    except Exception as e:
        raise OcrError(f"PDF повреждён или не читается ({type(e).__name__})")
    if force:
        actions = [a if a != "text" else "scan" for a in actions]

    dpi = int(cfg["dpi"])
    doc = pdfium.PdfDocument(str(src))
    # изображения на текстовых страницах: если внутри есть слова — страница «смешанная»
    for i, boxes in enumerate(candidates):
        if actions[i] != "text" or not boxes:
            continue
        page = doc[i]
        pw, ph = page.get_size()
        for (x0, top, x1, bottom) in boxes:
            with tempfile.TemporaryDirectory(dir=work_dir) as tmp:
                img = Path(tmp) / "img.png"
                page.render(scale=dpi / 72, grayscale=True,
                            crop=(x0, ph - bottom, pw - x1, top)).to_pil().save(img, dpi=(dpi, dpi))
                found = [w for w, c in tess.words(img, dpi) if c >= 60 and len(w) > 1]
            if len(found) >= int(cfg["слов_в_изображении"]):
                actions[i] = "mixed"
                break

    if all(a == "text" for a in actions):
        doc.close()
        shutil.copy2(src, dst)
        rep.status = "copied"
        rep.pages = [PageReport(i + 1, "text") for i in range(len(actions))]
        return rep

    thr = float(cfg["порог_уверенности"])
    writer = PdfWriter()
    try:
        for i, action in enumerate(actions):
            pr = PageReport(i + 1, action)
            if action == "text":
                writer.add_page(reader.pages[i])
            else:
                with tempfile.TemporaryDirectory(dir=work_dir) as tmp:
                    img = Path(tmp) / f"p{i + 1:04d}.png"
                    pil = doc[i].render(scale=dpi / 72, grayscale=True).to_pil()
                    pil.save(img, dpi=(dpi, dpi))
                    rot = tess.orientation(img)
                    if rot:
                        pil.rotate(-rot, expand=True).save(img, dpi=(dpi, dpi))
                    pr.rotation = rot
                    page_pdf, words = tess.ocr(img, dpi)
                writer.add_page(PdfReader(io.BytesIO(page_pdf)).pages[0])
                confs = [c for _, c in words]
                pr.words = len(words)
                pr.low_words = sum(1 for c in confs if c < 50)
                pr.confidence = round(sum(confs) / len(confs), 1) if confs else 0.0
                if pr.confidence < thr:
                    rep.low_pages.append(pr.page)
            rep.pages.append(pr)
    finally:
        doc.close()

    # метаданные: только служебные сведения без ПДн (исходные свойства документа не переносятся)
    low = {str(p.page): p.confidence for p in rep.pages if p.page in rep.low_pages}
    writer.add_metadata({
        "/Producer": f"medocr {__version__}",
        "/MedocrVersion": __version__,
        "/MedocrTesseract": tess.version,
        "/MedocrPages": json.dumps({str(p.page): p.action for p in rep.pages}),
        "/MedocrLowConfidence": json.dumps(low),
        "/MedocrThreshold": str(thr),
    })
    with open(dst, "wb") as f:
        writer.write(f)
    rep.status = "review" if rep.low_pages else "ok"
    return rep


# ---------------------------------------------------------------- папка

def process_folder(src_dir: Path, out_dir: Path, cfg: dict | None = None, force: bool = False,
                   progress: Callable[[int, int, DocReport], None] | None = None,
                   cancel: Callable[[], bool] | None = None) -> dict:
    cfg = {**DEFAULTS, **(cfg or {})}
    src_dir, out_dir = Path(src_dir).resolve(), Path(out_dir).resolve()
    if not src_dir.is_dir():
        raise OcrError(f"папка со сканами не найдена: {src_dir}")
    if src_dir == out_dir or src_dir in out_dir.parents:
        raise OcrError("папка результата OCR не должна совпадать с папкой сканов или лежать внутри неё")
    tess = Tesseract(cfg["tesseract"], cfg["языки"], cfg.get("sha256_языков"))
    out_dir.mkdir(parents=True, exist_ok=True)
    work = out_dir / ".medocr_tmp"                    # временные изображения — внутри защищённой папки
    work.mkdir(exist_ok=True)

    files = sorted(p for p in src_dir.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
    index_path = out_dir / "ocr_index.csv"
    known = {}
    if index_path.exists():
        with open(index_path, encoding="utf-8-sig") as f:
            known = {r["sha256_источника"]: r for r in csv.DictReader(f, delimiter=";")}
    used = {p.name for p in out_dir.glob("*.pdf")}
    rows, reports = [], []
    try:
        for i, src in enumerate(files, 1):
            if cancel and cancel():
                break
            sha = hashlib.sha256(src.read_bytes()).hexdigest()
            if sha in known and (out_dir / known[sha]["файл_результата"]).exists():
                rep = DocReport(known[sha]["файл_результата"], "skipped", note="уже распознан ранее")
            else:
                name = ascii_name(src.stem) + ".pdf"
                n = 2
                while name in used:
                    name = f"{ascii_name(src.stem)}_{n}.pdf"; n += 1
                used.add(name)
                try:
                    rep = process_pdf(src, out_dir / name, tess, cfg, work, force)
                except OcrError as e:
                    rep = DocReport(name, "error", note=str(e))
                except Exception as e:                     # без текста ошибки — в нём может быть фрагмент
                    rep = DocReport(name, "error", note=f"внутренняя ошибка {type(e).__name__}")
                if rep.status != "error":
                    rows.append({"sha256_источника": sha, "файл_источника": str(src.relative_to(src_dir)),
                                 "файл_результата": rep.output, "статус": rep.status,
                                 "страниц": len(rep.pages),
                                 "распознано_страниц": sum(1 for p in rep.pages if p.action != "text"),
                                 "низкая_уверенность": " ".join(map(str, rep.low_pages))})
            reports.append(rep)
            if progress:
                progress(i, len(files), rep)
    finally:
        shutil.rmtree(work, ignore_errors=True)

    if rows:
        new = not index_path.exists()
        with open(index_path, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0]), delimiter=";")
            if new:
                w.writeheader()
            w.writerows(rows)
    summary: dict[str, int] = {}
    for r in reports:
        summary[r.status] = summary.get(r.status, 0) + 1
    report = {"medocr": __version__, "tesseract": tess.version, "языки": cfg["языки"],
              "sha256_языков": tess.lang_sha, "dpi": cfg["dpi"], "порог_уверенности": cfg["порог_уверенности"],
              "итог": summary, "документы": [asdict(r) for r in reports]}
    (out_dir / "ocr_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report
