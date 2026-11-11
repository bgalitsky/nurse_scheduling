"""Поиск идентифицирующих сведений: правила + словари + (опционально) NER-модель.

Все правила детерминированы: одинаковый текст → одинаковый набор фрагментов.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------- словари

def _read_list(path: Path) -> list[str]:
    if not path.exists():
        return []
    return [l.strip() for l in path.read_text(encoding="utf-8").splitlines()
            if l.strip() and not l.lstrip().startswith("#")]


def name_stem(word: str) -> str:
    w = word
    for suf in ("ий", "ей", "ая", "ой", "ь"):
        if w.endswith(suf) and len(w) - len(suf) >= 3:
            return w[: -len(suf)]
    if w[-1] in "аяйоыи" and len(w) > 3:
        return w[:-1]
    return w


@dataclass
class Dictionaries:
    first_names: list[str]
    cities: list[str]
    eponyms: set[str]
    first_name_rx: str = ""
    city_rx: list[tuple[re.Pattern, str]] = field(default_factory=list)

    @classmethod
    def load(cls, d: Path) -> "Dictionaries":
        names = _read_list(d / "имена.txt")
        cities = _read_list(d / "города.txt")
        epon = {w.lower() for w in _read_list(d / "эпонимы.txt")}
        obj = cls(names, cities, epon)
        stems = sorted({re.escape(name_stem(n)) for n in names}, key=lambda s: (-len(s), s))
        obj.first_name_rx = r"(?:" + "|".join(stems) + r")[а-яё]{0,3}" if stems else r"(?!x)x"
        for c in sorted(cities, key=lambda s: (-len(s), s)):
            parts = c.split()
            parts[-1] = name_stem(parts[-1]) if len(parts) == 1 and "-" not in c else parts[-1]
            core = r"\s+".join(re.escape(p) for p in parts)
            obj.city_rx.append((re.compile(r"(?<![А-Яа-яЁё])" + core + r"[а-яё]{0,3}(?![А-Яа-яЁё])"), c))
        return obj


# ---------------------------------------------------------------- фрагменты

@dataclass
class Span:
    start: int
    end: int
    category: str            # ключ из категории_и_режимы (или 'дата', 'возраст')
    sub: str = ""            # уточнение: снилс/полис/паспорт, формат даты и т. п.
    data: dict = field(default_factory=dict)
    source: str = "правило"

    @property
    def length(self):
        return self.end - self.start


PRIORITY = {"номера_документов": 6.5, "место_работы": 3, "номер_истории": 0, "снилс_полис_паспорт": 0, "email": 1, "телефон": 2,
            "адрес": 3, "фио_пациента": 4, "фио_родственников": 4, "фио_врачей": 4,
            "фио_прочие": 5, "даты": 6, "возраст": 7}

# ---------------------------------------------------------------- регулярные выражения

U, L = "А-ЯЁ", "а-яё"
SURNAME = rf"[{U}][{L}]+(?:-[{U}][{L}]+)?"
INITIALS = rf"[{U}]\.[ ]?[{U}]\."
SP = r"[ \t]+"   # пробел внутри ФИО — без перевода строки
PATRONYMIC = rf"[{U}][{L}]+?(?:ович|евич|ьич|ич|овн|евн|ичн)(?:а|у|ем|е|ым|ом|ы|ой|ою)?"
SURNAME_SUFFIX = rf"[{U}][{L}]+(?:ов|ев|ёв|ин|ын|ск|цк|ко|ук|юк|ян|дзе|швили|ых|их)[{L}]{{0,3}}"
B = rf"(?<![{U}{L}])"     # граница слова для кириллицы
E = rf"(?![{U}{L}])"

ROLE_PATIENT = re.compile(
    r"(?:Ф\.?\s?И\.?\s?О\.?|пациент\w*|больн(?:ой|ая|ого|ую)|пациент\(ка\))", re.I)
ROLE_DOCTOR = re.compile(
    r"(?:врач\w*|зав\.|заведующ\w*|председател\w*|секретар\w*|член\w*\s+консилиума|"
    r"консультант\w*|исполнител\w*|профессор\w*|доцент\w*|д\.\s?м\.\s?н\.|к\.\s?м\.\s?н\.|"
    r"хирург\w*|терапевт\w*|кардиолог\w*|невролог\w*|онколог\w*|анестезиолог\w*|реаниматолог\w*|"
    r"рентгенолог\w*|радиолог\w*|эндокринолог\w*|уролог\w*|гинеколог\w*|офтальмолог\w*|"
    r"(?<!мед)сестр[аы]\s+(?:палатн|процедурн|старш)\w*|медсестр\w*|фельдшер\w*|ординатор\w*|"
    r"подпись|осмотрен\w*|осмотрел\w*|консультирован\w*|консультировал\w*|оперировал\w*|"
    r"прооперирован\w*|выполнил\w*\s+(?:операци|исследовани)\w*)", re.I)
ROLE_RELATIVE = re.compile(
    r"(?:\bмат(?:ь|ери|ерью)\b|\bот(?:ец|ца|цу|цом)\b|\bмуж(?:а|ем|у)?\b|\bжен(?:а|ы|ой)\b|супруг\w*|\bсын\w*|\bдоч(?:ь|ери)\w*|"
    r"\bбрат\w*|(?<!мед)\bсестр(?:а|ы|ой)\b|бабушк\w*|дедушк\w*|внук\w*|внучк\w*|родственни\w*|"
    r"законн\w+\s+представител\w*|представител\w*|опекун\w*|контактн\w+\s+лиц\w*)", re.I)

# Специальности врачей (именительный падеж = основа для поиска падежных форм)
SPECIALTIES = [
    "анестезиолог-реаниматолог", "акушер-гинеколог", "травматолог-ортопед", "сердечно-сосудистый хирург",
    "химиотерапевт", "радиотерапевт", "онколог-химиотерапевт", "онколог-радиолог", "нейрохирург",
    "оториноларинголог", "гастроэнтеролог", "аллерголог-иммунолог", "кардиолог", "невролог", "онколог",
    "хирург", "терапевт", "анестезиолог", "реаниматолог", "рентгенолог", "радиолог", "эндокринолог",
    "уролог", "гинеколог", "офтальмолог", "пульмонолог", "нефролог", "гематолог", "ревматолог",
    "дерматовенеролог", "дерматолог", "психиатр", "психотерапевт", "педиатр", "инфекционист",
    "патологоанатом", "колопроктолог", "маммолог", "эндоскопист", "трансфузиолог", "реабилитолог",
    "физиотерапевт", "диетолог", "стоматолог", "фтизиатр", "генетик", "гериатр",
]
SPEC_RX = re.compile(
    r"(?i)(?<![а-яё])(?:врач[а-яё]*\s*-\s*)?(" + "|".join(re.escape(x) for x in SPECIALTIES) + r")(?:а|у|ом|е|ы|ов|ами|ам)?(?![а-яё])")
POSITIONS = [
    (re.compile(r"(?i)зав(?:\.|едующ\w*)\s*(?:отделени\w*|отд\.)"), "ЗАВ_ОТДЕЛЕНИЕМ"),
    (re.compile(r"(?i)зам\w*\.?\s+главн\w*\s+врач\w*"), "ЗАМ_ГЛАВНОГО_ВРАЧА"),
    (re.compile(r"(?i)главн\w*\s+врач\w*"), "ГЛАВНЫЙ_ВРАЧ"),
    (re.compile(r"(?i)председател\w*"), "ПРЕДСЕДАТЕЛЬ_КОНСИЛИУМА"),
    (re.compile(r"(?i)секретар\w*"), "СЕКРЕТАРЬ_КОНСИЛИУМА"),
    (re.compile(r"(?i)лечащ\w*\s+врач\w*"), "ЛЕЧАЩИЙ_ВРАЧ"),
    (re.compile(r"(?i)дежурн\w*\s+врач\w*"), "ДЕЖУРНЫЙ_ВРАЧ"),
    (re.compile(r"(?i)ординатор\w*"), "ОРДИНАТОР"),
    (re.compile(r"(?i)(?<![а-яё])(?:мед(?:ицинская|\.)?\s*)?сестр\w*|медсестр\w*"), "МЕДСЕСТРА"),
]
DOCTOR_FIELD = re.compile(r"(?i)(?:^|[\n\t])[ ]*(?:лечащ\w*\s+)?врач[а-яё\-]*\s*:\s*$")


def doctor_label(text: str, start: int, end: int, title_spec: str | None) -> tuple[str, bool]:
    """Метка врача: специальность из строки → должность из строки → специальность из заголовка → ВРАЧ.
    Второе значение — найдена ли конкретная метка (а не общая «ВРАЧ»)."""
    ls = text.rfind("\n", 0, start) + 1
    le = text.find("\n", end); le = len(text) if le < 0 else le
    right = text[end: min(le, end + 45)]
    m = re.match(r"\s*[,–—-]?\s*\(?\s*" + SPEC_RX.pattern[4:], right, re.I)
    if m:
        return f"[ВРАЧ-{m.group(1).upper()}]", True
    left = text[max(ls, start - 120): start]
    best, pos = None, -1
    for m in SPEC_RX.finditer(left):
        if m.end() > pos:
            best, pos = f"[ВРАЧ-{m.group(1).upper()}]", m.end()
    for rx, name in POSITIONS:
        for m in rx.finditer(left):
            if m.end() > pos:
                best, pos = f"[{name}]", m.end()
    if best:
        return best, True
    if title_spec:
        return f"[ВРАЧ-{title_spec.upper()}]", True
    return "[ВРАЧ]", False


MONTHS = ["январ", "феврал", "март", "апрел", "ма[йя]", "июн", "июл", "август",
          "сентябр", "октябр", "ноябр", "декабр"]
MONTH_RX = "(" + "|".join(m + r"\w*" for m in MONTHS) + ")"
MONTH_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа",
             "сентября", "октября", "ноября", "декабря"]

DOB_BEFORE = re.compile(r"(?:дата\s+рождения|д\.\s?р\.|родил(?:ся|ась)|рожд\.?|г\.\s?р\.|\d{1,3}\s*(?:лет|года?)\s*\()\s*[:\-–]?\s*$", re.I)
DOB_AFTER = re.compile(r"^\s*(?:г\.?\s?р\.|года\s+рождения)", re.I)


def _month_index(word: str) -> int | None:
    w = word.lower()
    for i, m in enumerate(MONTHS):
        if re.match(m, w):
            return i + 1
    return None


def _year(y: str) -> int:
    v = int(y)
    if len(y) == 2:
        v += 2000 if v <= 30 else 1900
    return v


# ---------------------------------------------------------------- NER

class NerBackend:
    name = "нет"

    def persons(self, text: str) -> list[tuple[int, int]]:
        return []


# ---------------------------------------------------------------- пациент документа

@dataclass
class PatientInfo:
    surname: str
    first: str = ""          # имя или инициал
    patronymic: str = ""     # отчество или инициал
    birth_date: dt.date | None = None
    birth_year: int | None = None

    @property
    def identity(self) -> str:
        """Нормализованный идентификатор: ФАМИЛИЯ И.О. + дата рождения."""
        i = (self.first[:1] + "." if self.first else "") + (self.patronymic[:1] + "." if self.patronymic else "")
        b = self.birth_date.isoformat() if self.birth_date else (str(self.birth_year) if self.birth_year else "?")
        return f"{self.surname.upper()} {i.upper()}|{b}"

    @property
    def display(self) -> str:
        if len(self.first) == 1:
            return f"{self.surname} {self.first}.{self.patronymic}."
        return " ".join(x for x in (self.surname, self.first, self.patronymic) if x)


# ---------------------------------------------------------------- детектор

class Detector:
    def __init__(self, dicts: Dictionaries, ner: NerBackend):
        self.d, self.ner = dicts, ner
        FN = dicts.first_name_rx
        ANY = rf"[{U}][{L}]{{2,}}"            # имя не из словаря: опорой служит отчество
        self.rx_fio = [
            # Фамилия Имя Отчество
            ("фио_полное", re.compile(rf"{B}{SURNAME}{SP}{FN}{SP}{PATRONYMIC}{E}")),
            # Фамилия Имя Отчество, имя не из словаря (фамилия с типичным суффиксом)
            ("фио_полное_любое", re.compile(rf"{B}{SURNAME_SUFFIX}{SP}{ANY}{SP}{PATRONYMIC}{E}")),
            # Имя Отчество Фамилия (фамилия с типичным суффиксом)
            ("фио_ио_ф", re.compile(rf"{B}{FN}{SP}{PATRONYMIC}{SP}{SURNAME_SUFFIX}{E}")),
            # Фамилия И.О.
            ("фио_ф_ио", re.compile(rf"{B}{SURNAME}{SP}{INITIALS}")),
            # И.О. Фамилия
            ("фио_ио_ф_иниц", re.compile(rf"{B}{INITIALS}[ ]?{SURNAME}{E}")),
            # Имя Отчество (обращение), в т. ч. имя не из словаря
            ("фио_ио", re.compile(rf"{B}{FN}{SP}{PATRONYMIC}{E}")),
            ("фио_ио_любое", re.compile(rf"{B}{ANY}{SP}{PATRONYMIC}{E}")),
            # подпись /Сидоров/
            ("фио_подпись", re.compile(rf"/[ ]*({SURNAME}(?:{SP}{INITIALS})?)[ ]*/")),
        ]
        # для расширения найденной фамилии до полного ФИО
        self.rx_after = re.compile(rf"(?:{SP}(?:{ANY}{SP}{PATRONYMIC}{E}|{INITIALS}))")
        self.rx_before = re.compile(rf"(?:(?:{ANY}{SP}{PATRONYMIC}|{INITIALS}){SP}?)$")
        self.rx_patient_field = re.compile(
            r"(?im)^[ \t]*(?:Ф\.?\s?И\.?\s?О\.?(?:\s+(?:пациент\w*|больно\w*))?|Пациент(?:ка|\(ка\))?|Больн(?:ой|ая))"
            r"\s*[:\-–—]\s*(?P<v>[^\n]+)")
        self.rx_dob_field = re.compile(
            r"(?i)(?:дата\s+рождения|д\.\s?р\.)\s*[:\-–]?\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4}|\d{4})")
        self.rx_dob_gr = re.compile(r"(?i)(?<!\d)(?:\d{1,2}[./]\d{1,2}[./])?(\d{4})\s*г\.?\s?р\.")

    # ---------- пациент
    def find_patient(self, text: str) -> PatientInfo | None:
        for m in self.rx_patient_field.finditer(text):
            val = m.group("v")
            p = self._parse_name(val)
            if p:
                self._find_dob(text, p)
                return p
        return None

    def _parse_name(self, s: str) -> PatientInfo | None:
        FN = self.d.first_name_rx
        m = re.match(rf"\s*({SURNAME}){SP}({FN}){SP}({PATRONYMIC}){E}", s)
        if m:
            return PatientInfo(m.group(1), m.group(2), m.group(3))
        m = re.match(rf"\s*({SURNAME}){SP}([{U}])\.[ ]?([{U}])\.", s)
        if m:
            return PatientInfo(m.group(1), m.group(2), m.group(3))
        m = re.match(rf"\s*([{U}])\.[ ]?([{U}])\.[ ]?({SURNAME}){E}", s)
        if m:
            return PatientInfo(m.group(3), m.group(1), m.group(2))
        m = re.match(rf"\s*({SURNAME}){SP}([{U}][{L}]+){SP}([{U}][{L}]+){E}", s)   # имя не из словаря
        if m:
            return PatientInfo(m.group(1), m.group(2), m.group(3))
        return None

    def _find_dob(self, text, p: PatientInfo):
        m = self.rx_dob_field.search(text)
        if m:
            v = m.group(1)
            if re.fullmatch(r"\d{4}", v):
                p.birth_year = int(v)
            else:
                d, mo, y = re.split(r"[./]", v)
                try:
                    p.birth_date = dt.date(_year(y), int(mo), int(d)); p.birth_year = p.birth_date.year
                except ValueError:
                    pass
            return
        m = self.rx_dob_gr.search(text)
        if m:
            p.birth_year = int(m.group(1)); return
        # шапка МИС: «50 лет(14 сентября 1976)» — берём дату, а если она некорректна, хотя бы год
        m = re.search(r"(?i)\d{1,3}\s*(?:лет|года?)\s*\(([^)\n]{4,30})\)", text)
        if m:
            inner = m.group(1)
            d = re.fullmatch(rf"\s*(\d{{1,2}})[ .]+{MONTH_RX}[ ]+(\d{{4}})\s*", inner, re.I) or \
                re.fullmatch(r"\s*(\d{1,2})[./](\d{1,2})[./](\d{4})\s*", inner)
            y = re.search(r"(19|20)\d{2}", inner)
            if d:
                mo = _month_index(d.group(2)) if not d.group(2).isdigit() else int(d.group(2))
                try:
                    p.birth_date = dt.date(int(d.group(3)), mo, int(d.group(1)))
                except (ValueError, TypeError):
                    pass
            if y:
                p.birth_year = int(y.group(0))

    def surname_rx(self, p: PatientInfo) -> re.Pattern | None:
        s = p.surname
        stem = s
        for suf in ("ая", "ий", "ой", "ый"):
            if s.endswith(suf):
                stem = s[:-2]; break
        else:
            if s[-1] in "аяы" and len(s) > 4:
                stem = s[:-1]
        if len(stem) < 4:
            return None
        return re.compile(rf"{B}{re.escape(stem)}[{L}]{{0,3}}{E}")

    # ---------- основной проход
    def detect(self, text: str, patient: PatientInfo | None, llm_entities=None,
               use_rules: bool = True) -> tuple[list[Span], list[str]]:
        spans: list[Span] = []
        warnings: list[str] = []
        # даты и возраст — всегда правилами: их нужно разобрать, чтобы обобщить или сдвинуть
        spans += self._dates(text)
        spans += self._ages(text)
        if use_rules:
            spans += self._ids(text)
            spans += self._contacts(text)
            spans += self._addresses(text)
            spans += self._names(text, patient, warnings)
            spans += self._doc_numbers(text)
            spans += self._workplace(text)
        if llm_entities:
            spans += self._llm_spans(text, llm_entities)
        if patient:
            spans += self._patient_name_parts(text, patient)
        spans = self._expand_names(text, resolve(spans))
        spans = resolve(spans)
        self._label_doctors(text, spans)
        return spans, warnings

    # ---------- полное ФИО вокруг найденной фамилии
    def _expand_names(self, text: str, spans: list[Span]) -> list[Span]:
        """Если найдена только фамилия (моделью или по распространению), а рядом стоят
        имя с отчеством или инициалы, заменяется всё ФИО целиком."""
        out = []
        for sp in spans:
            if sp.category.startswith("фио"):
                m = self.rx_after.match(text, sp.end)
                if m:
                    sp = Span(sp.start, m.end(), sp.category, sp.sub, dict(sp.data), sp.source)
                ls = text.rfind("\n", 0, sp.start) + 1
                m = self.rx_before.search(text, ls, sp.start)
                if m and m.end() == sp.start:
                    sp = Span(m.start(), sp.end, sp.category, sp.sub, dict(sp.data), sp.source)
            out.append(sp)
        return out

    def _patient_name_parts(self, text: str, p: PatientInfo) -> list[Span]:
        """Имя и отчество пациента по отдельности («Со слов Эдуарда…», «Петровна жалуется»)."""
        out = []
        ends = "аяуюеиоыйьмх"
        for word, kind in ((p.first, "имя"), (p.patronymic, "отчество")):
            if len(word) < 4:                     # инициал — не ищем
                continue
            st = name_stem(word) if kind == "имя" else (word[:-1] if word.endswith("а") else word)
            if len(st) < 4:
                continue
            for m in re.finditer(rf"{B}{re.escape(st)}[{ends}]{{0,3}}{E}", text):
                out.append(Span(m.start(), m.end(), "фио_пациента", sub=f"пациент_{kind}"))
        return out

    # ---------- сущности от языковой модели
    def _llm_spans(self, text: str, entities) -> list[Span]:
        from .llm import LLM_CATEGORIES, find_all
        order = list(LLM_CATEGORIES)
        votes: dict[str, dict[str, int]] = {}
        spec: dict[str, str] = {}
        for e in entities:
            votes.setdefault(e.text, {}).setdefault(e.category, 0)
            votes[e.text][e.category] += 1
            if e.specialty:
                spec.setdefault(e.text, e.specialty)
        out = []
        for t in sorted(votes):
            cat = sorted(votes[t].items(), key=lambda kv: (-kv[1], order.index(kv[0])))[0][0]
            category, sub = LLM_CATEGORIES[cat]
            short_number = re.fullmatch(r"[\d\s/\-.]+", t) and len(re.sub(r"\D", "", t)) < 6
            base_category = category
            for s0, e0 in find_all(text, t):
                category = base_category
                if short_number and not re.search(r"(?:№|\bN|\bNo\.?)\s*$", text[max(0, s0 - 5): s0]):
                    continue        # «15» заменяем только как номер, а не везде, где встретилось число 15
                if re.search(r"(?:\bим\.|\bимени)\s*$", text[max(0, s0 - 8): s0]):
                    continue
                if category.startswith("фио"):
                    # явная подпись поля в строке сильнее догадки модели
                    ls = text.rfind("\n", 0, s0) + 1
                    left = text[max(0, ls - 1): s0]
                    if DOCTOR_FIELD.search(left):
                        category = "фио_врачей"
                    elif self.rx_patient_field.search(text[ls: e0]) and \
                            self.rx_patient_field.search(text[ls: e0]).start("v") <= s0 - ls + 1:
                        category = "фио_пациента"
                if category == "адрес":
                    sp = self._addr_span(text, s0, e0)
                else:
                    sp = Span(s0, e0, category, sub=sub)
                if category == "фио_врачей" and t in spec:
                    sp.data["спец_модели"] = f"[ВРАЧ-{spec[t].upper().replace(' ', '-')}]"
                sp.source = "модель"
                out.append(sp)
        return out

    def patient_from_llm(self, text: str, entities) -> PatientInfo | None:
        from .llm import find_all
        cands = []
        for e in entities:
            if e.category == "фио_пациента":
                occ = find_all(text, e.text)
                if occ:
                    cands.append((occ[0][0], e.text))
        for _, t in sorted(cands):
            p = self._parse_name(t)
            if p:
                self._find_dob(text, p)
                return p
        return None

    # ---------- специальность врача
    def _label_doctors(self, text: str, spans: list[Span]):
        lines = [l for l in text.splitlines() if l.strip()][:40]   # колонтитулы идут первыми
        title_spec = None
        for l in lines:
            if re.search(r"(?i)(?:осмотр|консультаци\w*|при[её]м|заключени\w*)\s+(?:врача|специалиста)", l):
                m = SPEC_RX.search(l)
                if m:
                    title_spec = m.group(1); break
        docs = [sp for sp in spans if sp.category == "фио_врачей"]
        known: dict[str, str] = {}
        for sp in docs:
            label, specific = doctor_label(text, sp.start, sp.end, title_spec)
            if sp.data.get("спец_модели") and not (specific and label.startswith("[ВРАЧ-")):
                label, specific = sp.data["спец_модели"], True
            sp.data["специальность"] = label
            sur = re.match(rf"[{U}][{L}]+", text[sp.start:sp.end])
            if specific and sur:
                known.setdefault(name_stem(sur.group(0)), label)
        for sp in docs:                     # «/Сидоров/» в подписи → метка того же врача из шапки
            if sp.data["специальность"] == "[ВРАЧ]":
                sur = re.match(rf"[{U}][{L}]+", text[sp.start:sp.end])
                if sur and name_stem(sur.group(0)) in known:
                    sp.data["специальность"] = known[name_stem(sur.group(0))]

    # ---------- номера документов: консилиумы, гистология, ИГХ, протоколы, «№ …»
    def _doc_numbers(self, text) -> list[Span]:
        out = []
        pats = [
            r"(?:№|(?<![A-Za-z])No\.?)[ ]*(\d[\w/\-]*)",                    # № 049-002-2-2025-0007427, №20-2655, № 17
            r"(?<![A-Za-z0-9])N[ ]?(\d[\w/\-]{2,})",                          # N 12345 (латинская N, от 3 знаков)
            r"(?<![\w.,])(\d{2,}-?[A-ZА-Я]\d{4,}(?:-\d+)?)(?!\w)",            # 2266K25000807-1, 2266-I25002303-2
            r"(?<![\w.,\-])(\d{2}-\d{4,})(?![\w\-]|[.,]\d)",                 # 20-6124, 22-3290
            r"(?<![\w.,\-])(\d{6,})(?![\w]|[.,]\d)",                          # 206124 и другие длинные числа
        ]
        for p in pats:
            for m in re.finditer(p, text):
                out.append(Span(m.start(1), m.end(1), "номера_документов"))
        return out

    # ---------- ФИО
    def _role(self, text: str, start: int, end: int, patient: PatientInfo | None, frag: str) -> str:
        other_person = False
        line_start = text.rfind("\n", 0, start) + 1
        if DOCTOR_FIELD.search(text[max(0, line_start - 1): start]):
            return "фио_врачей"            # «Врач: …» — это врач, даже если ФИО совпадает с пациентом
        if patient and self._is_patient_name(frag, patient):
            if not self._initials_conflict(frag, patient):
                return "фио_пациента"
            other_person = True        # та же фамилия, другие имя/отчество → родственник или однофамилец
        line_start = text.rfind("\n", 0, start) + 1
        left = text[max(line_start, start - 120): start]
        best, pos = None, -1
        roles = (("фио_врачей", ROLE_DOCTOR), ("фио_родственников", ROLE_RELATIVE))
        if not other_person:
            roles = (("фио_пациента", ROLE_PATIENT),) + roles
        for role, rx in roles:
            for m in rx.finditer(left):
                if m.end() > pos:
                    best, pos = role, m.end()
        if best:
            return best
        right = text[end: end + 40]
        m = re.match(r"\s*[,(–—-]\s*\(?\s*(\w+)", right)
        if m and ROLE_DOCTOR.match(m.group(1)):
            return "фио_врачей"
        if m and ROLE_RELATIVE.match(m.group(1)):
            return "фио_родственников"
        return "фио_прочие"

    def _initials_conflict(self, frag: str, p: PatientInfo) -> bool:
        """Есть ли во фрагменте имя/отчество (или инициалы), не совпадающие с пациентом."""
        if not p.first:
            return False
        words = re.findall(rf"[{U}][{L}]*", frag)
        sur = self.surname_rx(p)
        rest = [w for w in words if not (sur and sur.fullmatch(w))]
        if not rest:
            return False
        got = "".join(w[0] for w in rest[:2])
        want = (p.first[:1] + p.patronymic[:1])[: len(got)]
        return got != want

    def _is_patient_name(self, frag: str, p: PatientInfo) -> bool:
        rx = self.surname_rx(p)
        if rx and rx.search(frag):
            return True
        if p.first and len(p.first) > 1 and p.patronymic and len(p.patronymic) > 1:
            return (name_stem(p.first) in frag) and (p.patronymic[:-2] in frag)
        return False

    def institution_names(self, text) -> set[str]:
        """Фамилии из названий учреждений («им. А.С. Логинова») — по всему документу,
        чтобы и строка колонтитула «А.С. Логинова Департамента…» не считалась ФИО."""
        rx = re.compile(rf"(?:\bим\.?|\bимени)\s*(?:{INITIALS}[ ]?|[{U}]\.[ ]?[{U}][ ]?)?({SURNAME})")
        return {m.group(1) for m in rx.finditer(text)}

    def _names(self, text, patient, warnings) -> list[Span]:
        out = []
        inst = self.institution_names(text)
        for kind, rx in self.rx_fio:
            for m in rx.finditer(text):
                s, e = (m.start(1), m.end(1)) if kind == "фио_подпись" else (m.start(), m.end())
                frag = text[s:e]
                if re.search(r"(?:\bим\.?|\bимени)\s*$", text[max(0, s - 8): s]):
                    continue                      # «МКНЦ им. А.С. Логинова» — учреждение, не человек
                if inst and any(re.search(rf"{B}{re.escape(n)}{E}", frag) for n in inst) and \
                        not (patient and self._is_patient_name(frag, patient)):
                    continue
                role = self._role(text, s, e, patient, frag)
                if kind == "фио_подпись" and role == "фио_прочие":
                    role = "фио_врачей"
                out.append(Span(s, e, role, sub=kind))
        if patient:
            rx = self.surname_rx(patient)
            if rx is None:
                warnings.append("фамилия пациента слишком короткая для поиска падежных форм — нужна ручная проверка")
            else:
                if patient.surname.lower() in self.d.eponyms or name_stem(patient.surname).lower() in {name_stem(x) for x in self.d.eponyms}:
                    warnings.append("фамилия пациента совпадает с эпонимом — упоминания в названиях болезней оставлены, нужна ручная проверка")
                for m in rx.finditer(text):
                    if m.group(0).lower() in self.d.eponyms:
                        continue
                    out.append(Span(m.start(), m.end(), "фио_пациента", sub="фамилия_распространение"))
        for s, e in self.ner.persons(text):
            frag = text[s:e]
            out.append(Span(s, e, self._role(text, s, e, patient, frag), sub="ner", source="модель"))
        return out

    # ---------- документы, контакты
    def _ids(self, text) -> list[Span]:
        out = []
        for m in re.finditer(r"(?<!\d)\d{3}-\d{3}-\d{3}[ -]\d{2}(?!\d)", text):
            out.append(Span(m.start(), m.end(), "снилс_полис_паспорт", sub="снилс"))
        for m in re.finditer(r"(?i)СНИЛС\s*[:№]?\s*(\d[\d -]{9,13}\d)", text):
            out.append(Span(m.start(1), m.end(1), "снилс_полис_паспорт", sub="снилс"))
        for m in re.finditer(r"(?i)(?:полис\w*(?:\s+(?:ОМС|ДМС))?|\bОМС\b|\bДМС\b)[^\n\d]{0,25}((?:серия\s*\S+\s*)?(?:№[ ]*)?\d[\d -]{6,24}\d)", text):
            out.append(Span(m.start(1), m.end(1), "снилс_полис_паспорт", sub="полис"))
        for m in re.finditer(r"(?i)\bИНН\s*[:№]?\s*(\d{12}|\d{10})(?!\d)", text):
            out.append(Span(m.start(1), m.end(1), "снилс_полис_паспорт", sub="инн"))
        for m in re.finditer(r"(?im)паспорт\w*(?:\s+РФ|\s+гражданина\s+РФ)?\s*[:№]?\s*([^\n]*\d[^\n]*)$", text):
            out.append(Span(m.start(1), m.end(1), "снилс_полис_паспорт", sub="паспорт"))
        for m in re.finditer(
                r"(?i)(?:истори[яиюей]\s+болезни|медицинск\w+\s+карт\w*(?:\s+стационарного\s+больного|\s+амбулаторного\s+больного|\s+пациента)?|"
                r"мед\.?\s?карт\w*|амбулаторн\w+\s+карт\w*|\bи/б\b|\bИБ\b|\bМК\b)\s*(?:№|N|#|No\.?)\s*([A-ZА-Яа-я]{0,3}[-/]?\d[\w/\-]*)", text):
            out.append(Span(m.start(1), m.end(1), "номер_истории"))
        return out

    def _contacts(self, text) -> list[Span]:
        out = []
        for m in re.finditer(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", text):
            out.append(Span(m.start(), m.end(), "email"))
        for m in re.finditer(r"(?<![\d\w])(?:\+7|8)[ \-(]*\d{3}[ \-)]*\d{3}[ \-]*\d{2}[ \-]*\d{2}(?!\d)", text):
            out.append(Span(m.start(), m.end(), "телефон"))
        for m in re.finditer(r"(?i)\bтел(?:ефон|\.)?\w*\s*[:№]?\s*(\+?[\d][\d ()\-]{4,18}\d)", text):
            out.append(Span(m.start(1), m.end(1), "телефон"))
        return out

    # ---------- место работы / учёбы
    def _workplace(self, text) -> list[Span]:
        out = []
        rx = re.compile(r"(?im)^[ \t]*(?:место\s+(?:работы|учёбы|учебы|службы)|работает|учится|служит|должность|профессия)"
                        r"\s*[:\-–]\s*([^\n]+)$")
        for m in rx.finditer(text):
            v = m.group(1).strip()
            if v.lower().rstrip(".") in ("не работает", "пенсионер", "пенсионерка", "нет", "-", "—", "инвалид"):
                continue
            out.append(Span(m.start(1), m.start(1) + len(m.group(1).rstrip()), "место_работы"))
        return out

    # ---------- адреса
    STREET = (r"(?:ул\.|улица|пр-т|просп\.|проспект|пер\.|переулок|ш\.|шоссе|б-р|бульвар|наб\.|набережная|"
              r"мкр\.?|микрорайон|пл\.|площадь|проезд|тупик)")

    def _addresses(self, text) -> list[Span]:
        out = []
        field = re.compile(
            r"(?im)^[ \t]*(?:адрес(?:\s+(?:регистрации|проживания|места\s+жительства|фактического\s+проживания))?|"
            r"место\s+(?:жительства|регистрации)|проживает(?:\s+по\s+адресу)?|зарегистрирован\w*(?:\s+по\s+адресу)?)"
            r"\s*[:\-–]\s*([^\n]+)$")
        for m in field.finditer(text):
            out.append(self._addr_span(text, m.start(1), m.end(1)))
        inline = re.compile(
            rf"(?:\d{{6}},?\s*)?(?:[{U}][{L}-]+\s+(?:обл\.|область|край|респ\.)\s*,?\s*)?"
            rf"(?:(?:г\.|город|пос\.|с\.|д\.)\s?[{U}][{L}-]+(?:\s[{U}][{L}-]+)?\s*,?\s*)?"
            rf"{self.STREET}\s*[{U}0-9][^,\n]*(?:,\s*(?:д\.|дом|корп\.|к\.|стр\.|кв\.|квартира|под\.)\s*[\w/-]+)*")
        for m in inline.finditer(text):
            out.append(self._addr_span(text, m.start(), m.end()))
        inline2 = re.compile(
            rf"(?:\d{{6}},?\s*)?(?:[{U}][{L}-]+(?:-[{L}]+-[{U}][{L}]+)?,\s*)?"
            rf"[{U}][{L}-]+(?:[ ][{U}{L}][{L}-]+)?[ ](?:улица|проспект|шоссе|переулок|бульвар|набережная|проезд|площадь|тупик)"
            rf"(?:,\s*(?:д\.|дом|корп\.|к\.|стр\.|кв\.|квартира|под\.)\s*[\w/-]+)+")
        for m in inline2.finditer(text):
            out.append(self._addr_span(text, m.start(), m.end()))
        return out

    def _addr_span(self, text, s, e) -> Span:
        frag = text[s:e]
        city = None
        for rx, name in self.d.city_rx:
            if rx.search(frag):
                city = name; break
        return Span(s, e, "адрес", data={"город": city})

    # ---------- даты
    def _dates(self, text) -> list[Span]:
        out = []

        def add(m, y, mo, d, fmt, s=None, e=None):
            invalid = False
            try:
                date = dt.date(y, mo, d)
            except ValueError:
                if not (1900 <= y <= 2100 and 1 <= mo <= 12):
                    return
                date, invalid = dt.date(y, mo, 1), True   # «37 сентября» — опечатка, но всё равно дата
            s = m.start() if s is None else s
            e = m.end() if e is None else e
            dob = bool(DOB_BEFORE.search(text[max(0, s - 30): s]) or DOB_AFTER.match(text[e: e + 15]))
            out.append(Span(s, e, "даты", sub=fmt, data={"дата": date, "рождение": dob, "некорректная": invalid}))

        for m in re.finditer(r"(?<![\d.])(\d{1,2})[./](\d{1,2})[./](\d{4}|\d{2})(?![\d.]\d)", text):
            add(m, _year(m.group(3)), int(m.group(2)), int(m.group(1)), "дд.мм.гггг")
        # опечатка с запятой: «12,05.2020», «12.05,2020»
        for m in re.finditer(r"(?<![\d.,])(\d{1,2})(?:,(\d{1,2})\.|\.(\d{1,2}),)(\d{4})(?![\d.,]\d)", text):
            add(m, int(m.group(4)), int(m.group(2) or m.group(3)), int(m.group(1)), "дд.мм.гггг")
        # «08.04 2025»
        for m in re.finditer(r"(?<![\d.,])(\d{1,2})\.(\d{1,2})\.?[ ]((?:19|20)\d{2})(?![\d.,]\d)", text):
            add(m, int(m.group(3)), int(m.group(2)), int(m.group(1)), "дд.мм.гггг")
        # «09.2021г», «(03.2022г)» — месяц.год
        for m in re.finditer(r"(?<![\d.,])(0?[1-9]|1[0-2])\.((?:19|20)\d{2})(?![\d.,]\d)", text):
            add(m, int(m.group(2)), int(m.group(1)), 15, "мм.гггг")
        # «с 08.23 по 07.2024» — месяц.гг только после предлога и не перед единицами измерения
        UNITS = r"(?![ ]*(?:мм|см|мг|мл|г\b|%|ед|кг|л\b|раз|мкг|ммоль))"
        for m in re.finditer(rf"(?i)(?<![А-Яа-яЁё])(?:с|по|до|от|в|на)\s+(0[1-9]|1[0-2])\.(\d{{2}})(?!\d|[.,]\d){UNITS}", text):
            add(m, _year(m.group(2)), int(m.group(1)), 15, "мм.гггг", s=m.start(1), e=m.end(2))
        # диапазон «09.24-01.25»
        for m in re.finditer(rf"(?<![\d.,])(0[1-9]|1[0-2])\.(\d{{2}})[ ]?[-–][ ]?(0[1-9]|1[0-2])\.(\d{{2}})(?!\d|[.,]\d){UNITS}", text):
            add(m, _year(m.group(2)), int(m.group(1)), 15, "мм.гггг", s=m.start(1), e=m.end(2))
            add(m, _year(m.group(4)), int(m.group(3)), 15, "мм.гггг", s=m.start(3), e=m.end(4))
        # диапазон месяцев «05-07.2023»
        for m in re.finditer(r"(?<![\d.,\-])(0?[1-9]|1[0-2])[ ]?[-–][ ]?(0?[1-9]|1[0-2])\.((?:19|20)\d{2})(?!\d|[.,]\d)", text):
            y = int(m.group(3))
            if int(m.group(1)) <= int(m.group(2)):
                add(m, y, int(m.group(1)), 15, "мм-мм.гггг")
                if out and out[-1].start == m.start():
                    out[-1].data["дата2"] = dt.date(y, int(m.group(2)), 15)
        # «05-2023»
        for m in re.finditer(r"(?<![\d.,\-])(0[1-9]|1[0-2])-((?:19|20)\d{2})(?![\d\-])", text):
            add(m, int(m.group(2)), int(m.group(1)), 15, "мм.гггг")
        # опечатка «07.08..23»
        for m in re.finditer(r"(?<![\d.,])(\d{1,2})\.(\d{1,2})\.\.(\d{4}|\d{2})(?![\d])", text):
            add(m, _year(m.group(3)), int(m.group(2)), int(m.group(1)), "дд.мм.гггг")
        for m in re.finditer(r"(?<!\d)(\d{4})-(\d{2})-(\d{2})(?!\d)", text):
            add(m, int(m.group(1)), int(m.group(2)), int(m.group(3)), "гггг-мм-дд")
        for m in re.finditer(rf"(?i)(?<!\d)(\d{{1,2}})\s+{MONTH_RX}\s+(\d{{4}})", text):
            mo = _month_index(m.group(2))
            if mo:
                add(m, int(m.group(3)), mo, int(m.group(1)), "д месяц гггг")
        for m in re.finditer(rf"(?i){B}{MONTH_RX}\s+(\d{{4}})(?!\d)", text):
            mo = _month_index(m.group(1))
            if mo:
                add(m, int(m.group(2)), mo, 15, "месяц гггг")
        for m in re.finditer(r"(?<![\d./])((?:19|20)\d{2})(?=\s*(?:г\.|года?\b|гг\.))", text):
            add(m, int(m.group(1)), 7, 1, "гггг")
        return out

    # ---------- возраст
    def _ages(self, text) -> list[Span]:
        out = []
        pats = [
            r"(?i)(?:возраст\w*|в\s+возрасте)\s*[:\-–]?\s*(\d{1,3})(?![\d+])",
            r"(?i)(?:пациент\w*|больн\w+|мужчин\w*|женщин\w*)\s*,?\s*(\d{1,3})(?!\+)\s*(?:лет|года?)\b",
            r"(?i)(?<![\d.])(\d{1,3})(?!\+)\s*-?\s*летн\w*",
            r"(?i)(?<![\d.])(\d{1,3})\s+(?:лет|года?)\s*(?:\)|,)?\s*(?:от\s+роду|возраст)",
            # любое «N лет/года» при N ≥ 90 — для клинических сроков таких значений не бывает
            r"(?i)(?<![\d.,])(9\d|1[0-2]\d)(?![\d+])\s*(?:лет|года?)\b",
        ]
        for p in pats:
            for m in re.finditer(p, text):
                v = int(m.group(1))
                if 0 < v < 130:
                    out.append(Span(m.start(1), m.end(1), "возраст", data={"лет": v}))
        return out


def resolve(spans: list[Span]) -> list[Span]:
    """Снятие пересечений: сначала более длинный фрагмент, затем приоритет категории, затем позиция."""
    chosen: list[Span] = []
    for sp in sorted(spans, key=lambda s: (-s.length, PRIORITY.get(s.category, 9),
                                           0 if s.source == "модель" else 1, s.start, s.category, s.sub)):
        if all(sp.end <= c.start or sp.start >= c.end for c in chosen):
            chosen.append(sp)
    return sorted(chosen, key=lambda s: s.start)
