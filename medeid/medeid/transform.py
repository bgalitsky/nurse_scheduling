"""Применение режимов обезличивания к найденным фрагментам."""
from __future__ import annotations

import datetime as dt
import re

from .detect import MONTH_GEN, Span

LABELS = {
    "фио_пациента": "[ФИО_ПАЦИЕНТА]",
    "фио_родственников": "[ФИО_РОДСТВЕННИКА]",
    "фио_врачей": "[ФИО_ВРАЧА]",
    "фио_прочие": "[ФИО]",
    "адрес": "[АДРЕС]",
    "телефон": "[ТЕЛЕФОН]",
    "email": "[EMAIL]",
    "номер_истории": "[НОМЕР_ИСТОРИИ]",
    "даты": "[ДАТА]",
    "номера_документов": "[num]",
    "место_работы": "[МЕСТО_РАБОТЫ]",
}
TYPE_NAMES = {
    "фио_пациента": "ФИО пациента", "фио_родственников": "ФИО родственника", "фио_врачей": "ФИО врача",
    "фио_прочие": "ФИО", "адрес": "адрес", "телефон": "телефон", "email": "e-mail",
    "номер_истории": "номер истории болезни", "номера_документов": "номер документа",
    "место_работы": "место работы", "даты": "дата",
    "снилс": "СНИЛС", "полис": "полис", "паспорт": "паспорт", "инн": "ИНН",
}
SUB_LABELS = {"снилс": "[СНИЛС]", "полис": "[ПОЛИС]", "паспорт": "[ПАСПОРТ]", "инн": "[ИНН]"}
ROMAN_Q = {1: "I", 2: "II", 3: "III", 4: "IV"}


class Transformer:
    def __init__(self, modes: dict[str, str], over_89: str, patient_code: str | None,
                 shift_days: int | None, reference_date: dt.date | None):
        self.modes = modes
        self.over_89 = over_89
        self.code = patient_code
        self.shift = shift_days
        self.ref = reference_date

    def apply(self, text: str, spans: list[Span]) -> tuple[str, dict[str, int]]:
        """Меняются только найденные фрагменты; остальной текст переносится без изменений."""
        counts: dict[str, int] = {}
        out, pos, olen = [], 0, 0
        self.replaced: list[tuple[int, int]] = []      # позиции замен в выходном тексте
        for sp in spans:
            seg = text[pos: sp.start]
            out.append(seg); olen += len(seg)
            rep = self._replace(text[sp.start: sp.end], sp)
            out.append(rep); self.replaced.append((olen, olen + len(rep))); olen += len(rep)
            pos = sp.end
            key = sp.category if sp.category != "снилс_полис_паспорт" else f"{sp.category}.{sp.sub}"
            counts[key] = counts.get(key, 0) + 1
        out.append(text[pos:])
        return "".join(out), dict(sorted(counts.items()))

    def _deleted(self, sp: Span) -> str:
        key = sp.sub if sp.category == "снилс_полис_паспорт" else sp.category
        return f"[УДАЛЕНО: {TYPE_NAMES.get(key, key)}]"

    # ------------------------------------------------------------------
    def _replace(self, frag: str, sp: Span) -> str:
        cat = sp.category
        if cat == "возраст":
            return self._age(sp.data["лет"])
        if cat == "даты":
            return self._date(sp, frag)
        mode = self.modes[cat]
        if mode == "удаление":
            return self._deleted(sp)
        if cat == "фио_пациента" and mode == "сквозной_код":
            return self.code
        if cat == "фио_врачей" and mode == "специальность":
            return sp.data.get("специальность", "[ВРАЧ]")
        if cat == "адрес" and mode == "обобщение_до_города":
            return sp.data.get("город") or "[НАСЕЛЁННЫЙ_ПУНКТ]"
        if cat == "снилс_полис_паспорт":
            return SUB_LABELS.get(sp.sub, "[ДОКУМЕНТ]")
        return LABELS[cat]

    def _age(self, years: int) -> str:
        if years > 89:
            return self.over_89
        if self.modes["возраст"] == "возрастная_группа":
            lo = years // 10 * 10
            return f"{lo}–{lo + 9}"
        return str(years)

    def _date(self, sp: Span, frag: str) -> str:
        d: dt.date = sp.data["дата"]
        fmt = sp.sub
        if sp.data.get("рождение") and self.ref:
            age = self.ref.year - d.year - ((self.ref.month, self.ref.day) < (d.month, d.day))
            if age > 89:
                return f"[{self.over_89}]"
        mode = self.modes["даты"]
        if mode == "оставить":
            return frag
        if mode == "удаление":
            return self._deleted(sp)
        if mode == "замена_меткой":
            return LABELS["даты"]
        if mode == "только_год":
            return str(d.year)
        if mode == "квартал":
            if sp.data.get("некорректная"):
                return str(d.year)
            return f"{ROMAN_Q[(d.month - 1) // 3 + 1]} кв. {d.year}"
        if mode == "сдвиг":
            if sp.data.get("некорректная"):
                return LABELS["даты"]             # сдвигать несуществующую дату бессмысленно
            if self.shift is None:
                raise RuntimeError("сдвиг дат без определённого пациента")
            if fmt == "мм-мм.гггг":
                a = d + dt.timedelta(days=self.shift)
                b = sp.data["дата2"] + dt.timedelta(days=self.shift)
                return f"{a:%m.%Y}–{b:%m.%Y}"
            return self._fmt(d + dt.timedelta(days=self.shift), fmt)
        raise ValueError(mode)

    @staticmethod
    def _fmt(d: dt.date, fmt: str) -> str:
        if fmt == "гггг-мм-дд":
            return d.isoformat()
        if fmt == "д месяц гггг":
            return f"{d.day} {MONTH_GEN[d.month - 1]} {d.year}"
        if fmt in ("месяц гггг", "мм.гггг"):
            return f"{d.month:02d}.{d.year}"
        if fmt == "гггг":
            return str(d.year)
        return d.strftime("%d.%m.%Y")


# ---------------------------------------------------------------- контроль целостности
MED_RX = re.compile(
    r"\d+(?:[.,]\d+)?\s*(?:мг/м2|мг/м²|мг/кг|мг|мкг|мл|ммоль/л|мкмоль/л|г/л|ЕД|ед\.?|МЕ|Гр|сГр|мм|см|кг|%|×10\^?9/л)"
    r"(?![А-Яа-яЁё])", re.I)


def integrity_check(original: str, output: str, spans: list[Span], replaced: list[tuple[int, int]]) -> list[str]:
    """Независимая проверка требования «медицинский текст не меняется»:
    1) текст между заменами совпадает с исходным посимвольно;
    2) дозы, единицы и результаты анализов (число + единица) — те же и в том же порядке."""
    problems = []
    pos_in, pos_out = 0, 0
    for sp, (a, b) in zip(spans, replaced):
        if original[pos_in: sp.start] != output[pos_out: a]:
            problems.append("текст вне найденных фрагментов изменён")
            break
        pos_in, pos_out = sp.end, b
    else:
        if original[pos_in:] != output[pos_out:]:
            problems.append("текст вне найденных фрагментов изменён")

    def outside(rx, t, ranges):
        return [m.group(0) for m in rx.finditer(t) if all(m.end() <= x or m.start() >= y for x, y in ranges)]
    a = outside(MED_RX, original, [(s.start, s.end) for s in spans])
    b = outside(MED_RX, output, replaced)
    if a != b:
        problems.append(f"изменились дозы/единицы/результаты анализов ({len(a)} → {len(b)})")
    return problems
