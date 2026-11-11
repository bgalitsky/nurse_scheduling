"""Извлечение текста из PDF/DOCX и детерминированная нормализация.

В результат попадает только текст страниц (и колонтитулов DOCX). Метаданные, комментарии,
поля форм, закладки, вложения и изображения не переносятся; на месте изображения ставится
метка [ИЗОБРАЖЕНИЕ УДАЛЕНО]. Страницы без текстового слоя (сканы) отмечаются для ручной проверки.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber

SUPPORTED = {".pdf", ".docx"}
IMAGE_MARK = "[ИЗОБРАЖЕНИЕ УДАЛЕНО]"


class ExtractError(Exception):
    pass


@dataclass
class Extracted:
    text: str
    pages: int | None = None         # страниц на входе (для DOCX — None: страницы не определены)
    pages_with_text: int | None = None
    images: int = 0
    warnings: list[str] = field(default_factory=list)


def extract(path: Path, min_chars: int = 200) -> Extracted:
    if path.suffix.lower() == ".docx":
        text, images = _docx_text(path)
        res = Extracted(text, images=images)
    else:
        res = _pdf(path)
    if len(res.text.strip()) < min_chars:
        raise ExtractError("в документе почти нет текста (скан?) — нужен OCR")
    res.text = normalize(res.text)
    return res


def extract_text(path: Path, min_chars: int = 200) -> str:          # совместимость с 2.0
    return extract(path, min_chars).text


def _pdf(path: Path) -> Extracted:
    try:
        pdf = pdfplumber.open(str(path))
    except Exception as e:
        name = type(e).__name__
        if "Password" in name or "Encrypt" in name:
            raise ExtractError("PDF защищён паролем")
        raise ExtractError(f"PDF повреждён или не читается ({name})")
    parts, warnings = [], []
    with_text = images = 0
    with pdf:
        n = len(pdf.pages)
        for i, page in enumerate(pdf.pages, 1):
            t = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            imgs = len(page.images)
            images += imgs
            if len(t.strip()) >= 20:
                with_text += 1
            elif imgs:
                warnings.append(f"страница {i}: нет текстового слоя, только изображение (скан) — нужен OCR и ручная проверка")
            else:
                warnings.append(f"страница {i}: текст не извлечён")
            if t and _garbled(t):
                warnings.append(f"страница {i}: текстовый слой похож на «иероглифы» (нестандартный шрифт) — нужен OCR")
            parts.append(t + ("\n" + "\n".join([IMAGE_MARK] * imgs) if imgs else ""))
    return Extracted("\n\f\n".join(parts), pages=n, pages_with_text=with_text, images=images, warnings=warnings)


def _garbled(t: str) -> bool:
    letters = [c for c in t if c.isalpha()]
    if len(letters) < 50:
        return False
    ok = sum(1 for c in letters if "а" <= c.lower() <= "я" or c.lower() == "ё" or "a" <= c.lower() <= "z")
    return ok / len(letters) < 0.6


def _docx_text(path: Path) -> tuple[str, int]:
    """Колонтитулы (уникальные; колонки, разделённые табуляцией, — отдельными строками),
    затем тело документа в исходном порядке абзацев и таблиц.
    Колонтитулы обязательны: в выгрузках МИС там часто ФИО, дата рождения, полис, адрес."""
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    try:
        d = docx.Document(str(path))
    except Exception as e:
        raise ExtractError(f"DOCX повреждён или не читается ({type(e).__name__})")
    parts: list[str] = []
    seen: set[str] = set()
    images = 0

    def n_images(el) -> int:
        return len(el.xpath(".//*[local-name()='drawing' or local-name()='pict']"))

    for sec in d.sections:
        for hf in (sec.first_page_header, sec.header, sec.even_page_header,
                   sec.first_page_footer, sec.footer, sec.even_page_footer):
            try:
                blocks = [(p.text, n_images(p._p)) for p in hf.paragraphs]
                blocks += [(c.text, 0) for t in hf.tables for r in t.rows for c in r.cells]
            except Exception:
                continue
            for b, k in blocks:
                for piece in b.split("\t"):
                    piece = piece.strip()
                    if piece and piece not in seen:
                        seen.add(piece)
                        parts.append(piece)
                if k and ("img", b) not in seen:
                    seen.add(("img", b)); images += k
                    parts.extend([IMAGE_MARK] * k)

    for child in d.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            k = n_images(child)
            images += k
            parts.append(Paragraph(child, d).text + "".join(" " + IMAGE_MARK for _ in range(k)))
        elif tag == "tbl":
            for row in Table(child, d).rows:
                prev = None
                for cell in row.cells:
                    if cell._tc is prev:              # объединённые ячейки повторяются
                        continue
                    prev = cell._tc
                    k = n_images(cell._tc)
                    images += k
                    parts.append(cell.text + "".join(" " + IMAGE_MARK for _ in range(k)))
    return "\n".join(parts).replace("\t", "\n"), images


_HYPHEN_BREAK = re.compile(r"([а-яё])-\n([а-яё])")
_GLUED = re.compile(r"([а-яё])([А-ЯЁ][а-яё])")


def normalize(text: str) -> str:
    text = unicodedata.normalize("NFC", text)
    text = text.replace("Nº", "№").replace(" ", " ").replace(" ", " ").replace("‑", "-")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _HYPHEN_BREAK.sub(r"\1\2", text)          # перенос «Мазу-\nренко» → «Мазуренко»
    text = _GLUED.sub(r"\1 \2", text)              # «НиколаевичПациент» → «Николаевич Пациент»
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip() + "\n"


# Латинские буквы, похожие на кириллические (ошибки OCR и «маскировка»): «Мaзypенко».
_HOMO = str.maketrans("AaBEeKkMHOoPpCcTyXx", "АаВЕеКкМНОоРрСсТуХх")
_WORD = re.compile(r"[A-Za-zА-Яа-яЁё]+")
_CYR = re.compile(r"[А-Яа-яЁё]")


def fold(text: str) -> str:
    """Текст для ПОИСКА (не для вывода): латиница внутри кириллических слов → кириллица, ё → е.
    Длина и позиции символов сохраняются, поэтому найденные фрагменты применяются к исходному тексту."""
    t = _WORD.sub(lambda m: m.group(0).translate(_HOMO) if _CYR.search(m.group(0)) else m.group(0), text)
    return t.replace("ё", "е").replace("Ё", "Е")


DOC_TYPE_PATTERNS = [
    ("осмотр_врача", re.compile(r"(?:осмотр|консультаци\w+|при[её]м)\s+(?:врача|специалиста)", re.I)),
    ("протокол_консилиума", re.compile(r"консилиум", re.I)),
    ("выписной_эпикриз", re.compile(r"выписн\w*\s+эпикриз|эпикриз\s+выписн|выписка\s+из\s+(?:медицинской\s+карты|истории\s+болезни)", re.I)),
    ("результаты_обследования", re.compile(r"протокол\s+(?:исследования|обследования)|результат\w*\s+(?:исследования|обследования|анализ)|заключение\s+(?:УЗИ|КТ|МРТ|ЭКГ|ЭхоКГ)|\bМРТ\b|\bКТ\b", re.I)),
]


def detect_doc_type(text: str) -> str | None:
    """Тип по заголовку; если в заголовке не найден — по началу текста."""
    lines = [l for l in text.splitlines() if l.strip()]
    title_area = "\n".join(lines[:12])
    for area in (title_area, text[:1500]):
        for name, rx in DOC_TYPE_PATTERNS:
            if rx.search(area):
                return name
    return None
