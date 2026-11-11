"""Окно приложения medeid (tkinter).

Роли определяются по учётной записи Windows/Linux (getpass.getuser()) и файлу roles.yaml:
  оператор       — выбирает папки исходников и результатов, запускает партии, отмечает проверку,
                   удаляет исходники/результаты партии, вносит отзыв согласия;
  администратор  — видит журнал, события и таблицу ключей (без ФИО; раскрытие — только с основанием
                   и записью в журнал), выгружает отчёты. Документов не видит: в его окне нет
                   доступа к папкам исходников и результатов.
Окно — не граница безопасности: доступ к папкам ограничивается правами NTFS (см. GAPS.md).

Запуск:  python -m medeid app
"""
from __future__ import annotations

import sys as _sys
from pathlib import Path as _Path

if __name__ == "__main__" and not __package__:
    # запуск файлом (python medeid/app.py, кнопка Run в PyCharm) → запускаем как модуль пакета
    _sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
    import runpy
    runpy.run_module("medeid.app", run_name="__main__")
    raise SystemExit

import getpass
import json
import os
import queue
import sqlite3
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import yaml

from . import __version__
from .config import DOC_TYPES, MODES, ConfigError, load_raw
from .journal import Journal

APP_DIR = Path(__file__).resolve().parent.parent            # хранилище «правила и настройки»
HOME = Path.home()
DEFAULT_SETTINGS = {
    "исходники": str(HOME / "medeid_data" / "исходники"),
    "результаты": str(HOME / "medeid_data" / "результаты"),
    "хранилище_ключей": str(HOME / "medeid_secure" / "keys.sqlite"),
    "кэш_модели": str(HOME / "medeid_secure" / "llm_cache.sqlite"),
    "журнал": str(HOME / "medeid_journal" / "journal.sqlite"),
}
DOC_TYPE_NAMES = {
    "выписной_эпикриз": "Выписной эпикриз", "протокол_консилиума": "Протокол консилиума",
    "результаты_обследования": "Результаты обследования", "осмотр_врача": "Осмотр врача",
}
CATEGORY_NAMES = {
    "фио_пациента": "ФИО пациента", "фио_родственников": "ФИО родственников", "фио_врачей": "ФИО врачей",
    "фио_прочие": "ФИО прочих лиц", "адрес": "Адрес", "телефон": "Телефон", "email": "E-mail",
    "снилс_полис_паспорт": "СНИЛС, полис, паспорт, ИНН", "номер_истории": "Номер истории болезни",
    "номера_документов": "Номера документов", "место_работы": "Место работы/учёбы", "даты": "Даты",
    "возраст": "Возраст (до 90 лет)",
}
PURPOSES = ["контроль качества", "научное исследование", "обучение модели на обезличенных данных"]


# ---------------------------------------------------------------- файлы настроек

def _yaml(path: Path, default):
    if path.exists():
        return yaml.safe_load(path.read_text(encoding="utf-8")) or default
    return default


def load_settings() -> dict:
    s = dict(DEFAULT_SETTINGS)
    s.update(_yaml(APP_DIR / "app_settings.yaml", {}))
    return s


def save_settings(s: dict):
    (APP_DIR / "app_settings.yaml").write_text(
        "# Папки по умолчанию. Пути хранилищ ключей, кэша и журнала меняет администратор.\n"
        + yaml.safe_dump(s, allow_unicode=True, sort_keys=False), encoding="utf-8")


def base_config() -> dict:
    p = APP_DIR / "config.yaml"
    if not p.exists():
        p = APP_DIR / "config.example.yaml"
    return _yaml(p, {})


def roles_for(user: str) -> tuple[list[str], str | None]:
    p = APP_DIR / "roles.yaml"
    if not p.exists():
        p.write_text("# Роли по учётным записям. Файл ведёт отдел ИБ.\n"
                     f"операторы: [{user}]\nадминистраторы: [{user}]\n", encoding="utf-8")
        note = (f"Создан файл ролей {p}: пользователю {user} назначены обе роли. "
                f"Отредактируйте его перед работой с реальными документами.")
    else:
        note = None
    r = _yaml(p, {})
    roles = []
    if user in (r.get("операторы") or []):
        roles.append("оператор")
    if user in (r.get("администраторы") or []):
        roles.append("администратор")
    return roles, note


# ---------------------------------------------------------------- общие элементы

def safe_error(e: BaseException) -> str:
    """Сообщение об ошибке без фрагментов документов: свои ошибки — текстом, чужие — только тип."""
    if isinstance(e, (ConfigError, RuntimeError, ValueError)) and e.__class__.__module__.startswith(("medeid", "builtins")):
        return str(e)
    return f"внутренняя ошибка {type(e).__name__}"


def tree(parent, cols: list[tuple[str, int]], height=10) -> ttk.Treeview:
    frame = ttk.Frame(parent)
    tv = ttk.Treeview(frame, columns=[c for c, _ in cols], show="headings", height=height)
    for c, w in cols:
        tv.heading(c, text=c); tv.column(c, width=w, stretch=True)
    sb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    tv.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
    frame.pack(fill="both", expand=True, padx=6, pady=4)
    return tv


def open_folder(path: Path):
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform.startswith("win"):
        os.startfile(str(path))                      # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


# ---------------------------------------------------------------- окно оператора

class OperatorWindow:
    def __init__(self, root: tk.Tk, user: str):
        self.root, self.user = root, user
        self.settings = load_settings()
        self.base = base_config()
        self.q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_flag = threading.Event()
        root.title(f"medeid {__version__} — оператор ({user})")

        nb = ttk.Notebook(root); nb.pack(fill="both", expand=True)
        self._tab_run(ttk.Frame(nb), nb)
        self._tab_modes(ttk.Frame(nb), nb)
        self._tab_batches(ttk.Frame(nb), nb)
        self._tab_consent(ttk.Frame(nb), nb)
        self._after = root.after(150, self._poll)
        root.bind("<Destroy>", lambda e: self._stop_poll() if e.widget is root else None)

    def _stop_poll(self):
        try:
            self.root.after_cancel(self._after)
        except (tk.TclError, ValueError):
            pass

    # ---------- вкладка «Партия»
    def _tab_run(self, f, nb):
        nb.add(f, text="Партия")
        g = ttk.LabelFrame(f, text="Хранилища"); g.pack(fill="x", padx=8, pady=4)
        self.v_in = tk.StringVar(value=self.settings["исходники"])
        self.v_out = tk.StringVar(value=self.settings["результаты"])
        for r, (label, var) in enumerate((("Исходные документы:", self.v_in), ("Обезличенные результаты:", self.v_out))):
            ttk.Label(g, text=label).grid(row=r, column=0, sticky="w", padx=4, pady=2)
            ttk.Entry(g, textvariable=var, width=70).grid(row=r, column=1, sticky="we", padx=4)
            ttk.Button(g, text="Обзор…", command=lambda v=var: self._browse(v)).grid(row=r, column=2, padx=4)
        ttk.Button(g, text="Сохранить как папки по умолчанию", command=self._save_dirs).grid(row=2, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(g, foreground="gray", text=f"Ключи: {Path(self.settings['хранилище_ключей']).parent}   "
                  f"Журнал: {self.settings['журнал']}").grid(row=3, column=0, columnspan=3, sticky="w", padx=4)
        g.columnconfigure(1, weight=1)

        b = ttk.LabelFrame(f, text="Основание (обязательно)"); b.pack(fill="x", padx=8, pady=4)
        bas = self.base.get("основание") or {}
        clean = lambda v: "" if str(v).strip("_ ") == "" else str(v)
        self.v_order = tk.StringVar(value=clean(bas.get("приказ_номер", "")))
        self.v_date = tk.StringVar(value=clean(bas.get("приказ_дата", "")))
        self.v_purpose = tk.StringVar(value=clean(bas.get("цель", PURPOSES[0])) or PURPOSES[0])
        self.v_dept = tk.StringVar(value=clean(bas.get("отделение", "")))
        for c, (label, var, w) in enumerate((("Приказ №", self.v_order, 10), ("от", self.v_date, 12),
                                             ("Отделение", self.v_dept, 18))):
            ttk.Label(b, text=label).grid(row=0, column=2 * c, padx=4, pady=2, sticky="e")
            ttk.Entry(b, textvariable=var, width=w).grid(row=0, column=2 * c + 1, padx=4, sticky="w")
        ttk.Label(b, text="Цель").grid(row=1, column=0, padx=4, sticky="e")
        ttk.Combobox(b, textvariable=self.v_purpose, values=PURPOSES, width=40).grid(row=1, column=1, columnspan=5, sticky="w", padx=4, pady=2)

        t = ttk.LabelFrame(f, text="Типы документов (остальные отклоняются)"); t.pack(fill="x", padx=8, pady=4)
        chosen = set(self.base.get("типы_документов") or [])
        self.v_types = {}
        for i, dtp in enumerate(sorted(DOC_TYPES)):
            v = tk.BooleanVar(value=dtp in chosen)
            self.v_types[dtp] = v
            ttk.Checkbutton(t, text=DOC_TYPE_NAMES.get(dtp, dtp), variable=v).grid(row=0, column=i, padx=6, sticky="w")

        m = ttk.LabelFrame(f, text="Поиск фрагментов"); m.pack(fill="x", padx=8, pady=4)
        ext = str(self.base.get("извлечение", "правила"))
        self.v_llm = tk.BooleanVar(value=ext != "правила")
        self.v_llm_mode = tk.StringVar(value="модель" if ext == "модель" else "модель+правила")
        ttk.Checkbutton(m, text="Использовать локальную языковую модель (LLM)", variable=self.v_llm,
                        command=self._llm_toggle).grid(row=0, column=0, sticky="w", padx=4)
        self.rb1 = ttk.Radiobutton(m, text="модель + правила (рекомендуется)", value="модель+правила", variable=self.v_llm_mode)
        self.rb2 = ttk.Radiobutton(m, text="только модель", value="модель", variable=self.v_llm_mode)
        self.rb1.grid(row=0, column=1, padx=6); self.rb2.grid(row=0, column=2, padx=6)
        self.btn_check = ttk.Button(m, text="Проверить модель", command=self._check_model)
        self.btn_check.grid(row=0, column=3, padx=6)
        mod = self.base.get("модель") or {}
        ttk.Label(m, foreground="gray", text=f"Модель: {mod.get('имя', '—')}  сервер: {mod.get('сервер', '—')}"
                  ).grid(row=1, column=0, columnspan=4, sticky="w", padx=4)
        self._llm_toggle()

        k = ttk.LabelFrame(f, text="Обратимость"); k.pack(fill="x", padx=8, pady=4)
        self.v_code = tk.BooleanVar(value=str(self.base.get("сквозной_код_пациента", "нет")).lower() in ("да", "true", "yes"))
        ttk.Checkbutton(k, text="Сквозной код пациента (ПАЦИЕНТ_0042). Создаётся ключ «код ↔ пациент» — "
                        "данные остаются персональными", variable=self.v_code, command=self._code_toggle
                        ).pack(anchor="w", padx=4)

        a = ttk.Frame(f); a.pack(fill="x", padx=8, pady=6)
        self.btn_run = ttk.Button(a, text="▶ Запустить партию", command=self._run)
        self.btn_run.pack(side="left")
        self.btn_stop = ttk.Button(a, text="■ Остановить", command=self.cancel_flag.set, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        ttk.Button(a, text="Открыть папку результатов", command=lambda: open_folder(Path(self.v_out.get()))).pack(side="left", padx=6)
        self.pb = ttk.Progressbar(f, mode="determinate"); self.pb.pack(fill="x", padx=8)
        self.log = tk.Text(f, height=12, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=4)
        ttk.Label(f, foreground="gray", text="В этом окне показываются только внутренние номера документов (DOC-…) и статусы."
                  ).pack(anchor="w", padx=8)

    # ---------- вкладка «Режимы»
    def _tab_modes(self, f, nb):
        nb.add(f, text="Режимы")
        cats = self.base.get("категории_и_режимы") or {}
        self.v_modes = {}
        ttk.Label(f, text="Режим по умолчанию для научных наборов — необратимый (удаление, замена меткой, обобщение).",
                  foreground="gray").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=6)
        r = 1
        for cat, allowed in MODES.items():
            vals = sorted(allowed)
            default = cats.get(cat) or {"фио_прочие": "замена_меткой", "номера_документов": "замена_меткой",
                                        "место_работы": "замена_меткой", "возраст": "как_есть"}.get(cat, vals[0])
            v = tk.StringVar(value=default)
            self.v_modes[cat] = v
            ttk.Label(f, text=CATEGORY_NAMES.get(cat, cat)).grid(row=r, column=0, sticky="w", padx=8, pady=2)
            cb = ttk.Combobox(f, textvariable=v, values=vals, state="readonly", width=24)
            cb.grid(row=r, column=1, sticky="w", padx=4)
            if cat == "фио_пациента":
                self.cb_patient = cb
            r += 1
        self.v_over89 = tk.StringVar(value=str(cats.get("возраст_старше_89", "90+")))
        ttk.Label(f, text="Возраст старше 89 →").grid(row=r, column=0, sticky="w", padx=8, pady=2)
        ttk.Entry(f, textvariable=self.v_over89, width=10).grid(row=r, column=1, sticky="w", padx=4); r += 1
        sd = self.base.get("сдвиг_дат") or {}
        rng = sd.get("диапазон_дней", [-365, 365])
        self.v_shift_lo = tk.IntVar(value=int(rng[0])); self.v_shift_hi = tk.IntVar(value=int(rng[1]))
        self.v_shift_after = tk.StringVar(value=sd.get("после_обработки", "удалить"))
        ttk.Label(f, text="Сдвиг дат: диапазон дней").grid(row=r, column=0, sticky="w", padx=8, pady=2)
        sf = ttk.Frame(f); sf.grid(row=r, column=1, sticky="w")
        ttk.Spinbox(sf, from_=-3650, to=0, textvariable=self.v_shift_lo, width=7).pack(side="left")
        ttk.Label(sf, text="…").pack(side="left")
        ttk.Spinbox(sf, from_=1, to=3650, textvariable=self.v_shift_hi, width=7).pack(side="left"); r += 1
        ttk.Label(f, text="Величина сдвига после обработки").grid(row=r, column=0, sticky="w", padx=8, pady=2)
        ttk.Combobox(f, textvariable=self.v_shift_after, values=["удалить", "хранить_в_ключе"], state="readonly",
                     width=24).grid(row=r, column=1, sticky="w", padx=4)
        self._code_toggle(ask=False)

    # ---------- вкладка «Партии»
    def _tab_batches(self, f, nb):
        nb.add(f, text="Партии")
        bar = ttk.Frame(f); bar.pack(fill="x", padx=6, pady=4)
        for text, cmd in (("Обновить", self._load_batches), ("Отметить ручную проверку…", self._manual_check),
                          ("Проверить повторяемость", self._verify), ("Открыть папку партии", self._open_batch),
                          ("Удалить исходники партии…", self._del_sources), ("Удалить результаты партии…", self._del_outputs)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=3)
        self.tv_batches = tree(f, [("партия", 110), ("вид", 120), ("начало (UTC)", 150), ("кто", 90),
                                   ("статус", 110), ("итог", 260), ("автопроверка", 120), ("ручная проверка", 110)], 18)
        self._load_batches()

    # ---------- вкладка «Отзыв согласия»
    def _tab_consent(self, f, nb):
        nb.add(f, text="Отзыв согласия")
        ttk.Label(f, wraplength=760, justify="left", text=(
            "Пациент исключается из БУДУЩИХ партий, обрабатываемых из текущей папки исходных документов. "
            "Из уже выданных необратимых наборов удалить его нельзя (фиксируется в локальном документе). "
            "В реестр хранилища исходников записывается только отпечаток «ФАМИЛИЯ И.О.|дата рождения»; "
            "в журнал — только факт и основание. Не указывайте ФИО в поле «Основание».")).pack(anchor="w", padx=8, pady=8)
        g = ttk.Frame(f); g.pack(anchor="w", padx=8)
        self.v_ex_fio = tk.StringVar(); self.v_ex_dob = tk.StringVar(); self.v_ex_basis = tk.StringVar()
        for r, (label, var) in enumerate((("ФИО (Фамилия Имя Отчество или Фамилия И.О.)", self.v_ex_fio),
                                          ("Дата рождения (дд.мм.гггг)", self.v_ex_dob),
                                          ("Основание (например, заявление № 15 от 10.11.2026)", self.v_ex_basis))):
            ttk.Label(g, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(g, textvariable=var, width=40).grid(row=r, column=1, padx=6)
        ttk.Button(g, text="Исключить из будущих партий", command=self._exclude).grid(row=3, column=1, sticky="w", pady=6)

    # ---------- действия
    def _browse(self, var):
        d = filedialog.askdirectory(initialdir=var.get() or str(HOME))
        if d:
            var.set(d)

    def _save_dirs(self):
        self.settings["исходники"], self.settings["результаты"] = self.v_in.get(), self.v_out.get()
        save_settings(self.settings)
        messagebox.showinfo("medeid", "Папки сохранены как папки по умолчанию.")

    def _llm_toggle(self):
        st = "normal" if self.v_llm.get() else "disabled"
        for w in (self.rb1, self.rb2, self.btn_check):
            w.configure(state=st)

    def _code_toggle(self, ask=True):
        if not hasattr(self, "cb_patient"):
            return
        if self.v_code.get():
            if ask and not messagebox.askyesno("Сквозной код пациента", (
                    "Будет создан ключ «код ↔ пациент». Пока ключ существует, данные по закону остаются "
                    "персональными. Включение фиксируется в журнале.\n\nВключить?")):
                self.v_code.set(False)
                return
            self.v_modes["фио_пациента"].set("сквозной_код")
            self.cb_patient.configure(state="disabled")
        else:
            if self.v_modes["фио_пациента"].get() == "сквозной_код":
                self.v_modes["фио_пациента"].set("замена_меткой")
            self.cb_patient.configure(state="readonly")

    def _raw_config(self) -> dict:
        raw = json.loads(json.dumps(self.base, ensure_ascii=False, default=str))
        raw["запустил"] = self.user
        raw["основание"] = {"приказ_номер": self.v_order.get().strip(), "приказ_дата": self.v_date.get().strip(),
                            "цель": self.v_purpose.get().strip(), "отделение": self.v_dept.get().strip()}
        raw["типы_документов"] = [t for t, v in self.v_types.items() if v.get()]
        cats = {k: v.get() for k, v in self.v_modes.items()}
        cats["возраст_старше_89"] = self.v_over89.get().strip()
        raw["категории_и_режимы"] = cats
        raw["сквозной_код_пациента"] = "да" if self.v_code.get() else "нет"
        raw["сдвиг_дат"] = {"диапазон_дней": [self.v_shift_lo.get(), self.v_shift_hi.get()],
                            "после_обработки": self.v_shift_after.get()}
        raw["извлечение"] = self.v_llm_mode.get() if self.v_llm.get() else "правила"
        raw["хранилище_ключей"] = self.settings["хранилище_ключей"]
        raw["журнал"] = self.settings["журнал"]
        raw.setdefault("модель", {})["кэш"] = self.settings["кэш_модели"]
        return raw

    def _say(self, line: str):
        self.log.configure(state="normal"); self.log.insert("end", line + "\n"); self.log.see("end")
        self.log.configure(state="disabled")

    def _run(self):
        try:
            cfg = load_raw(self._raw_config(), APP_DIR)
        except ConfigError as e:
            messagebox.showerror("Конфигурация", str(e)); return
        inp, out = Path(self.v_in.get()), Path(self.v_out.get())
        if not inp.is_dir():
            messagebox.showerror("medeid", f"Папка исходных документов не найдена:\n{inp}"); return
        self.cancel_flag.clear()
        self.btn_run.configure(state="disabled"); self.btn_stop.configure(state="normal")
        self.pb.configure(value=0)
        self._say(f"— запуск: {cfg.extraction}, режим ФИО пациента: {cfg.modes['фио_пациента']}")

        def work():
            from .pipeline import Pipeline
            try:
                m = Pipeline(cfg).run(inp, out, progress=lambda i, n, r: self.q.put(("progress", i, n, r)),
                                      cancel=self.cancel_flag.is_set)
                self.q.put(("done", m))
            except BaseException as e:                 # noqa: BLE001
                self.q.put(("error", safe_error(e)))
        self.worker = threading.Thread(target=work, daemon=True); self.worker.start()

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, i, n, r = msg
                    self.pb.configure(maximum=n, value=i)
                    extra = f" — {len(r['warnings'])} замеч." if r.get("warnings") else ""
                    self._say(f"{r.get('doc_id', '?')}: {r['status']}{extra}")
                elif msg[0] == "done":
                    m = msg[1]
                    self._say(f"— партия {m['партия']}: {m['статус']}; автопроверка: {m['автопроверка']}; "
                              + ", ".join(f"{k}: {v}" for k, v in m["итог"].items()))
                    self._finish()
                    self._load_batches()
                    messagebox.showinfo("Партия обработана",
                                        f"Партия {m['партия']}\nПапка: {m['папка_партии']}\nАвтопроверка: {m['автопроверка']}")
                elif msg[0] == "error":
                    self._say(f"— ошибка: {msg[1]}"); self._finish()
                    messagebox.showerror("medeid", msg[1])
                elif msg[0] == "info":
                    self._finish(); messagebox.showinfo(msg[1], msg[2])
        except queue.Empty:
            pass
        except tk.TclError:                            # окно закрыто
            return
        try:
            self._after = self.root.after(150, self._poll)
        except tk.TclError:
            pass

    def _finish(self):
        self.btn_run.configure(state="normal"); self.btn_stop.configure(state="disabled")
        self.btn_check.configure(state="normal" if self.v_llm.get() else "disabled")

    def _check_model(self):
        from .llm import LLMClient, LLMError, LLMSettings
        s = LLMSettings.from_raw(self.base.get("модель") or {}); s.cache = None
        self.btn_check.configure(state="disabled")

        def work():
            try:
                c = LLMClient(s)
                d = c.verify_model()
                ents, w = c.extract("Пациент: Тестова Мария Олеговна, тел. +7 900 000-00-00. "
                                    "Лечащий врач: Проверкин А.А. (кардиолог).\n")
                txt = (f"Сервер {s.server} доступен (внутренний адрес).\nДайджест совпадает: {d[:16]}…\n"
                       f"Пробное извлечение (вымышленный текст): найдено {len(ents)}:\n"
                       + "\n".join(f"  {e.category}: {e.text}" for e in ents) + ("\n" + "\n".join(w) if w else ""))
                self.q.put(("info", "Модель", txt))
            except LLMError as e:
                self.q.put(("error", f"модель: {e}"))
            except BaseException as e:                 # noqa: BLE001
                self.q.put(("error", safe_error(e)))
        threading.Thread(target=work, daemon=True).start()

    # ---------- партии
    def _journal(self) -> Journal:
        return Journal(Path(self.settings["журнал"]))

    def _load_batches(self):
        tv = self.tv_batches
        tv.delete(*tv.get_children())
        if not Path(self.settings["журнал"]).exists():
            return
        jr = self._journal()
        for r in jr.runs():
            batch, kind, started, who, _basis, status, summary, auto, manual = r[:9]
            s = ", ".join(f"{k}: {v}" for k, v in json.loads(summary or "{}").items())
            tv.insert("", "end", iid=f"{batch}|{r[12]}", values=(batch, kind, started, who, status, s, auto or "", manual or ""))
        jr.close()

    def _selected_batch(self) -> str | None:
        sel = self.tv_batches.selection()
        if not sel:
            messagebox.showinfo("medeid", "Выберите партию в списке."); return None
        return sel[0].split("|")[0]

    def _manual_check(self):
        b = self._selected_batch()
        if not b:
            return
        ok = messagebox.askyesnocancel("Ручная проверка", f"Партия {b}: выборочная ручная проверка пройдена?\n"
                                       "Да — пройдена, Нет — не пройдена.")
        if ok is None:
            return
        note = simpledialog.askstring("Комментарий", "Комментарий (без ФИО и номеров документов):") or ""
        jr = self._journal(); jr.set_manual_check(b, self.user, ok, note); jr.close()
        self._load_batches()

    def _open_batch(self):
        b = self._selected_batch()
        if b:
            open_folder(Path(self.v_out.get()) / b)

    def _verify(self):
        b = self._selected_batch()
        if not b:
            return
        from .pipeline import verify_batch
        inp, bdir = Path(self.v_in.get()), Path(self.v_out.get()) / b
        self._say(f"— проверка повторяемости партии {b}…")

        def work():
            try:
                r = verify_batch(inp, bdir, Path(self.settings["журнал"]), APP_DIR)
                if not r["различаются"]:
                    txt = f"Партия {b}: ПОВТОРЯЕМО. {r['всего']} док., результаты побайтно совпадают."
                else:
                    txt = (f"Партия {b}: НЕ СОВПАДАЮТ {len(r['различаются'])} из {r['всего']}: "
                           f"{', '.join(r['различаются'][:15])}")
                    if r["ожидаемо_из_за_сдвига"]:
                        txt += "\nОжидаемо: величина сдвига дат была удалена после обработки."
                if not r["совпадает_отпечаток"]:
                    txt += "\nОтпечаток запуска отличается: изменились версии, модель, словари или исходные файлы."
                self.q.put(("info", "Повторяемость", txt))
            except BaseException as e:                 # noqa: BLE001
                self.q.put(("error", safe_error(e)))
        threading.Thread(target=work, daemon=True).start()

    def _del_sources(self):
        b = self._selected_batch()
        if not b or not messagebox.askyesno("Удаление исходников", (
                f"Удалить из папки исходных документов все файлы, обработанные в партии {b}?\n"
                "Действие необратимо и записывается в журнал (для акта об уничтожении).")):
            return
        from .pipeline import delete_sources
        try:
            cfg = load_raw(self._raw_config_for_ops(), APP_DIR)
            n = delete_sources(cfg, Path(self.v_in.get()), b, self.user)
            messagebox.showinfo("medeid", f"Удалено исходных файлов: {n}")
        except BaseException as e:                     # noqa: BLE001
            messagebox.showerror("medeid", safe_error(e))

    def _del_outputs(self):
        b = self._selected_batch()
        if not b:
            return
        reason = simpledialog.askstring("Удаление результатов", f"Основание удаления результатов партии {b} "
                                        "(приказ, акт; без ФИО):")
        if not reason:
            return
        from .pipeline import delete_outputs
        try:
            cfg = load_raw(self._raw_config_for_ops(), APP_DIR)
            n = delete_outputs(cfg, Path(self.v_out.get()), b, self.user, reason)
            messagebox.showinfo("medeid", f"Удалено файлов результата: {n}")
        except BaseException as e:                     # noqa: BLE001
            messagebox.showerror("medeid", safe_error(e))

    def _raw_config_for_ops(self) -> dict:
        raw = self._raw_config()
        if not all(raw["основание"].values()):          # для удаления основание партии уже в журнале
            raw["основание"] = {"операция": "удаление"}
        if not raw["типы_документов"]:
            raw["типы_документов"] = sorted(DOC_TYPES)
        return raw

    def _exclude(self):
        from .detect import Detector, Dictionaries, NerBackend
        from .extract import fold
        from .registry import Registry
        import datetime as dt
        fio, dob, basis = self.v_ex_fio.get().strip(), self.v_ex_dob.get().strip(), self.v_ex_basis.get().strip()
        if not (fio and dob and basis):
            messagebox.showerror("medeid", "Заполните ФИО, дату рождения и основание."); return
        try:
            d, mo, y = (int(x) for x in dob.split("."))
            birth = dt.date(y, mo, d)
        except ValueError:
            messagebox.showerror("medeid", "Дата рождения в формате дд.мм.гггг"); return
        det = Detector(Dictionaries.load(APP_DIR / "dicts" / str(self.base.get("версия_словарей", ""))), NerBackend())
        p = det._parse_name(fold(fio))
        if not p:
            messagebox.showerror("medeid", "Не удалось разобрать ФИО. Пример: Иванова Мария Петровна или Иванова М.П."); return
        p.birth_date, p.birth_year = birth, birth.year
        inp = Path(self.v_in.get())
        if not inp.is_dir():
            messagebox.showerror("medeid", "Папка исходных документов не найдена."); return
        reg = Registry(inp, self.user); reg.exclude(p.identity, basis); n = reg.excluded_count(); reg.close()
        jr = self._journal(); jr.event(self.user, "отзыв_согласия", None, {"основание": basis}); jr.close()
        self.v_ex_fio.set(""); self.v_ex_dob.set("")
        messagebox.showinfo("medeid", f"Пациент исключён из будущих партий. Всего исключено: {n}.")


# ---------------------------------------------------------------- окно администратора

class AdminWindow:
    def __init__(self, root: tk.Tk, user: str):
        self.root, self.user = root, user
        self.settings = load_settings()
        root.title(f"medeid {__version__} — администратор ({user})")
        nb = ttk.Notebook(root); nb.pack(fill="both", expand=True)
        self._tab_journal(ttk.Frame(nb), nb)
        self._tab_events(ttk.Frame(nb), nb)
        self._tab_keys(ttk.Frame(nb), nb)

    def _journal(self) -> Journal | None:
        p = Path(self.settings["журнал"])
        if not p.exists():
            messagebox.showinfo("medeid", f"Журнал ещё не создан: {p}"); return None
        return Journal(p)

    # ---------- журнал партий
    def _tab_journal(self, f, nb):
        nb.add(f, text="Журнал партий")
        bar = ttk.Frame(f); bar.pack(fill="x", padx=6, pady=4)
        ttk.Button(bar, text="Обновить", command=self._load_runs).pack(side="left", padx=3)
        ttk.Button(bar, text="Выгрузить журнал событий (CSV)…", command=lambda: self._export(None)).pack(side="left", padx=3)
        self.tv_runs = tree(f, [("партия", 110), ("вид", 120), ("начало (UTC)", 150), ("кто", 90), ("основание", 220),
                                ("статус", 100), ("автопроверка", 110), ("ручная", 90), ("версии", 220)], 8)
        self.tv_runs.bind("<<TreeviewSelect>>", self._show_run)
        self.tv_files = tree(f, [("документ", 100), ("тип", 140), ("статус", 140), ("sha256 вход", 130),
                                 ("sha256 выход", 130), ("стр. вход/текст", 100), ("изобр.", 60),
                                 ("замены", 260), ("замечания", 260)], 8)
        self.txt_cfg = tk.Text(f, height=7, wrap="word", state="disabled")
        self.txt_cfg.pack(fill="both", padx=6, pady=4)
        self._load_runs()

    def _load_runs(self):
        tv = self.tv_runs
        tv.delete(*tv.get_children())
        jr = self._journal()
        if not jr:
            return
        for r in jr.runs():
            batch, kind, started, who, basis, status, _s, auto, manual, tv_, mv, dv, run_id = r
            b = json.loads(basis or "{}")
            btxt = f"№{b.get('приказ_номер', '')} от {b.get('приказ_дата', '')}, {b.get('цель', '')}"
            tv.insert("", "end", iid=f"{batch}|{run_id}", values=(batch, kind, started, who, btxt, status,
                                                                    auto or "", manual or "", f"{tv_} / {mv} / {dv}"))
        jr.close()

    def _show_run(self, _e=None):
        sel = self.tv_runs.selection()
        if not sel:
            return
        batch = sel[0].split("|")[0]
        jr = self._journal()
        if not jr:
            return
        self.tv_files.delete(*self.tv_files.get_children())
        for r in jr.files_of(batch):
            did, dtype, st, hin, hout, pin, ptxt, imgs, counts, warns, _code = r
            c = ", ".join(f"{k}: {v}" for k, v in json.loads(counts or "{}").items())
            w = "; ".join(json.loads(warns or "[]"))
            self.tv_files.insert("", "end", values=(did, dtype or "", st, hin or "", hout or "",
                                                    f"{pin if pin is not None else '—'}/{ptxt if ptxt is not None else '—'}",
                                                    imgs if imgs is not None else "—", c, w))
        cfg = jr.config_of(batch)
        jr.close()
        self.txt_cfg.configure(state="normal"); self.txt_cfg.delete("1.0", "end")
        if cfg:
            try:
                pretty = yaml.safe_dump(json.loads(cfg[0] or "{}"), allow_unicode=True, sort_keys=False)
            except ValueError:
                pretty = cfg[0] or ""
            self.txt_cfg.insert("end", f"Отпечаток запуска: {cfg[2]}\nКонфигурация партии (без ПДн):\n{pretty}")
        self.txt_cfg.configure(state="disabled")

    # ---------- события
    def _tab_events(self, f, nb):
        nb.add(f, text="События и выдача")
        bar = ttk.Frame(f); bar.pack(fill="x", padx=6, pady=4)
        ttk.Button(bar, text="Обновить", command=self._load_events).pack(side="left", padx=3)
        ttk.Button(bar, text="Записать выдачу набора…", command=lambda: self._record("выдача_набора")).pack(side="left", padx=3)
        ttk.Button(bar, text="Записать копирование…", command=lambda: self._record("копирование_набора")).pack(side="left", padx=3)
        ttk.Button(bar, text="Выгрузка для акта об уничтожении (CSV)…",
                   command=lambda: self._export(["удаление_набора", "удаление_исходников"])).pack(side="left", padx=3)
        self.tv_events = tree(f, [("дата/время (UTC)", 150), ("пользователь", 100), ("событие", 160),
                                  ("партия", 110), ("подробности", 520)], 20)
        self._load_events()

    def _load_events(self):
        tv = self.tv_events
        tv.delete(*tv.get_children())
        jr = self._journal()
        if not jr:
            return
        for ts, user, ev, batch, det in jr.events():
            d = json.loads(det or "{}")
            tv.insert("", "end", values=(ts, user, ev, batch or "", "; ".join(f"{k}: {v}" for k, v in d.items())))
        jr.close()

    def _record(self, kind):
        batch = simpledialog.askstring("Партия", "Номер партии (например 2026-11-11-001):")
        if not batch:
            return
        who = simpledialog.askstring("Кому / куда", "Кому выдан набор или куда скопирован (подразделение, система):")
        basis = simpledialog.askstring("Основание", "Тема или приказ:")
        if not who or not basis:
            return
        jr = self._journal()
        if jr:
            jr.event(self.user, kind, batch, {"получатель_или_место": who, "основание": basis}); jr.close()
        self._load_events()

    def _export(self, kinds):
        jr = self._journal()
        if not jr:
            return
        p = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                         initialfile="журнал_удаления.csv" if kinds else "журнал_событий.csv")
        if p:
            Path(p).write_text(jr.export_csv(kinds), encoding="utf-8-sig")
            jr.event(self.user, "выгрузка_журнала", None, {"события": kinds or "все"})
        jr.close()

    # ---------- ключи
    def _tab_keys(self, f, nb):
        nb.add(f, text="Ключи")
        ttk.Label(f, foreground="gray", wraplength=900, justify="left", text=(
            "Таблица «код ↔ пациент» хранится отдельно. ФИО здесь не показываются; раскрытие кода — только "
            "с основанием, каждое раскрытие записывается в журнал доступа хранилища ключей и в журнал событий."
        )).pack(anchor="w", padx=8, pady=6)
        bar = ttk.Frame(f); bar.pack(fill="x", padx=6)
        ttk.Button(bar, text="Обновить", command=self._load_keys).pack(side="left", padx=3)
        ttk.Button(bar, text="Раскрыть код…", command=self._reveal).pack(side="left", padx=3)
        self.tv_keys = tree(f, [("код", 140), ("создан (UTC)", 170), ("кем", 120), ("сдвиг дат хранится", 140)], 12)
        ttk.Label(f, text="Журнал доступа к ключам").pack(anchor="w", padx=8)
        self.tv_access = tree(f, [("дата/время (UTC)", 170), ("пользователь", 120), ("действие", 120),
                                  ("код", 140), ("основание", 400)], 8)
        self._load_keys()

    def _ks_conn(self):
        p = Path(self.settings["хранилище_ключей"])
        if not p.exists():
            return None
        return sqlite3.connect(f"file:{p}?mode=ro", uri=True)

    def _load_keys(self):
        for tv in (self.tv_keys, self.tv_access):
            tv.delete(*tv.get_children())
        db = self._ks_conn()
        if not db:
            return
        shifts = {r[0] for r in db.execute("SELECT identity_hash FROM date_shifts")}
        for code, h, created, by in db.execute("SELECT code, identity_hash, created_at, created_by FROM patients ORDER BY code"):
            self.tv_keys.insert("", "end", values=(code, created, by, "да" if h in shifts else "нет"))
        for row in db.execute("SELECT ts, user, action, code, reason FROM access_log ORDER BY ts"):
            self.tv_access.insert("", "end", values=tuple(x or "" for x in row))
        db.close()

    def _reveal(self):
        sel = self.tv_keys.selection()
        if not sel:
            messagebox.showinfo("medeid", "Выберите код в таблице."); return
        code = self.tv_keys.item(sel[0])["values"][0]
        reason = simpledialog.askstring("Основание раскрытия", f"Основание раскрытия кода {code} (приказ, запрос):")
        if not reason:
            return
        from .keystore import KeyStore
        ks = KeyStore(Path(self.settings["хранилище_ключей"]), self.user, {})
        row = ks.reveal(code, reason); ks.close()
        jr = self._journal()
        if jr:
            jr.event(self.user, "раскрытие_ключа", None, {"код": code, "основание": reason}); jr.close()
        if row:
            messagebox.showinfo(f"Код {code}", f"{row[1]}\nДата рождения: {row[2] or '—'}\n\n"
                                "Раскрытие записано в журнал. Не копируйте эти сведения в незащищённые места.")
        self._load_keys()


# ---------------------------------------------------------------- запуск

def main():
    user = getpass.getuser()
    root = tk.Tk()
    root.geometry("1100x780")
    root.report_callback_exception = lambda et, ev, tb: messagebox.showerror("medeid", safe_error(ev))
    roles, note = roles_for(user)
    if note:
        messagebox.showwarning("medeid", note)
    if not roles:
        messagebox.showerror("medeid", f"Учётной записи {user} не назначена роль (roles.yaml). Обратитесь в отдел ИБ.")
        root.destroy(); return
    role = roles[0]
    if len(roles) > 1:
        role = "оператор" if messagebox.askyesno("Роль", f"{user}: войти как оператор?\n"
                                                 "Да — оператор (обработка), Нет — администратор (журнал и ключи).") \
            else "администратор"
    (OperatorWindow if role == "оператор" else AdminWindow)(root, user)
    root.mainloop()


if __name__ == "__main__":
    main()
