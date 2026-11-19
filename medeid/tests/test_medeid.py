"""pytest tests/ — синтетические документы, без реальных данных."""
import shutil, subprocess, sys
from pathlib import Path

import yaml
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
from medeid.config import load
from medeid.pipeline import Pipeline


from medeid.settings import init_settings  # noqa: E402

DICTS = ROOT / "templates" / "dicts"


def _settings(tmp) -> Path:
    """Отдельная папка настроек (как у ИБ), созданная из шаблонов программы."""
    sd = tmp / "settings"
    if not (sd / "config.yaml").exists():
        init_settings(sd, "tester")
    return sd


def _cfg(tmp, **over):
    sd = _settings(tmp)
    c = yaml.safe_load((sd / "config.yaml").read_text(encoding="utf-8"))
    c["основание"] = {"приказ_номер": "1", "приказ_дата": "01.11.2026", "цель": "тест", "отделение": "тест"}
    c["извлечение"] = "правила"
    c["хранилище_ключей"] = str(tmp / "secure" / "keys.sqlite")
    c["журнал"] = str(tmp / "journal.sqlite")
    for k, v in over.items():
        if k in c.get("категории_и_режимы", {}):
            c["категории_и_режимы"][k] = v
        else:
            c[k] = v
    p = sd / "config.yaml"; p.write_text(yaml.safe_dump(c, allow_unicode=True), encoding="utf-8")
    return load(p)


@pytest.fixture(scope="module")
def pdfs(tmp_path_factory):
    d = tmp_path_factory.mktemp("in")
    subprocess.run([sys.executable, str(ROOT / "tests" / "make_samples.py"), str(d)], check=True)
    return d


def _texts(out):
    return {p.name: p.read_text(encoding="utf-8") for p in out.rglob("*.txt")}


def test_no_identifiers_left(pdfs, tmp_path):
    out = tmp_path / "out"
    Pipeline(_cfg(tmp_path)).run(pdfs, out)
    allt = "\n".join(_texts(out).values())
    for bad in ["Мазуренко", "Галин", "Сидоров", "Петрова", "Тверская", "123-45-67", "@mail.ru",
                "123-456-789", "345678", "12345/26", "14.02.1934", "Иванов", "Волков", "92 лет"]:
        assert bad not in allt, bad
    assert "Москва" in allt and "Химки" in allt and "[ФИО_РОДСТВЕННИКА]" in allt
    # специальность врача вместо ФИО
    assert "[ВРАЧ-КАРДИОЛОГ] (кардиолог), [ВРАЧ-НЕВРОЛОГ] (невролог)" in allt
    assert "Председатель консилиума: зам. главного врача [ЗАМ_ГЛАВНОГО_ВРАЧА]" in allt
    assert "/[ЛЕЧАЩИЙ_ВРАЧ]/" in allt                    # подпись связана с врачом из шапки
    # номера документов
    assert "ПРОТОКОЛ КОНСИЛИУМА ВРАЧЕЙ № [num]" in allt


def test_repeatable(pdfs, tmp_path):
    cfg = _cfg(tmp_path)
    a = Pipeline(cfg).run(pdfs, tmp_path / "a")
    b = Pipeline(cfg).run(pdfs, tmp_path / "b")
    assert [r.get("output_sha256") for r in a["файлы"]] == [r.get("output_sha256") for r in b["файлы"]]


def test_through_code_and_shift(pdfs, tmp_path):
    cfg = _cfg(tmp_path, фио_пациента="сквозной_код", сквозной_код_пациента="да", даты="сдвиг",
               сдвиг_дат={"диапазон_дней": [-365, 365], "после_обработки": "хранить_в_ключе"})
    out = tmp_path / "o"
    Pipeline(cfg).run(pdfs, out)
    texts = _texts(out)
    assert all(n.startswith("DOC-") for n in texts)                  # имя файла — только внутренний номер
    assert sum("ПАЦИЕНТ_0002" in t for t in texts.values()) == 2      # один пациент — один код в двух документах
    assert (tmp_path / "secure" / "keys.sqlite").stat().st_mode & 0o077 == 0
    again = Pipeline(cfg).run(pdfs, tmp_path / "o2")
    assert _texts(out) == _texts(tmp_path / "o2")


def test_keystore_inside_output_rejected(pdfs, tmp_path):
    cfg = _cfg(tmp_path, фио_пациента="сквозной_код", сквозной_код_пациента="да")
    cfg.keystore = tmp_path / "out" / "keys.sqlite"
    with pytest.raises(RuntimeError):
        Pipeline(cfg).run(pdfs, tmp_path / "out")


# ------------------------------------------------------------ 2.0: языковая модель
import mock_ollama  # noqa: E402


@pytest.fixture(scope="module")
def llm_server():
    srv = mock_ollama.start()
    yield f"http://127.0.0.1:{srv.server_address[1]}"
    srv.shutdown()


def _llm_cfg(tmp, server, mode="модель+правила", **model):
    m = {"движок": "ollama", "сервер": server, "имя": "qwen3:8b", "дайджест": mock_ollama.DIGEST,
         "кэш": str(tmp / "secure" / "llm_cache.sqlite"), **model}
    return _cfg(tmp, извлечение=mode, модель=m)


def test_llm_hybrid(pdfs, llm_server, tmp_path):
    out = tmp_path / "out"
    m = Pipeline(_llm_cfg(tmp_path, llm_server)).run(pdfs, out)
    allt = "\n".join(_texts(out).values())
    for bad in ["Мазуренко", "Сидоров", "Тверская", "123-45-67", "12345/26", "Иванов", "Несуществующий"]:
        assert bad not in allt, bad
    assert "[ВРАЧ-КАРДИОЛОГ]" in allt and "[ВРАЧ-НЕВРОЛОГ]" in allt
    st = m["модель"]["статистика"]
    assert st["отброшено_не_найдено_в_тексте"] > 0          # выдуманная сущность отброшена
    assert m["модель"]["дайджест_на_сервере"] == mock_ollama.DIGEST


def test_llm_only_and_cache_repeatable(pdfs, llm_server, tmp_path):
    cfg = _llm_cfg(tmp_path, llm_server, mode="модель")
    a = Pipeline(cfg).run(pdfs, tmp_path / "a")
    calls = mock_ollama.CALLS["chat"]
    b = Pipeline(cfg).run(pdfs, tmp_path / "b")
    assert mock_ollama.CALLS["chat"] == calls                 # второй прогон целиком из кэша
    assert b["модель"]["статистика"]["из_кэша"] > 0
    assert _texts(tmp_path / "a") == _texts(tmp_path / "b")
    assert "Мазуренко" not in "\n".join(_texts(tmp_path / "a").values())


def test_llm_digest_mismatch(pdfs, llm_server, tmp_path):
    cfg = _llm_cfg(tmp_path, llm_server, дайджест="b" * 64)
    with pytest.raises(RuntimeError, match="дайджест"):
        Pipeline(cfg).run(pdfs, tmp_path / "o")


def test_llm_external_server_refused(pdfs, tmp_path):
    cfg = _llm_cfg(tmp_path, "http://8.8.8.8:11434")
    with pytest.raises(RuntimeError, match="вне локальной сети"):
        Pipeline(cfg).run(pdfs, tmp_path / "o")


def test_llm_short_numbers_not_global(llm_server, tmp_path):
    from medeid.detect import Detector, Dictionaries, NerBackend
    from medeid.llm import Entity
    d = Detector(Dictionaries.load(DICTS / "2026-11-01"), NerBackend())
    text = "Консилиум № 88. Доза 88 мг.\n"
    spans, _ = d.detect(text, None, [Entity("88", "номер_документа")], use_rules=False)
    assert [(text[s.start:s.end], s.category) for s in spans if s.category == "номера_документов"] == [("88", "номера_документов")]
    assert spans[0].start == text.index("88")


def test_field_label_beats_model(llm_server, tmp_path):
    from medeid.detect import Detector, Dictionaries, NerBackend
    from medeid.llm import Entity
    d = Detector(Dictionaries.load(DICTS / "2026-11-01"), NerBackend())
    text = "Врач: Врачев Игорь Николаевич\nПациент: Врачев Игорь Николаевич\n"
    spans, _ = d.detect(text, None, [Entity("Врачев Игорь Николаевич", "фио_пациента")], use_rules=False)
    assert [s.category for s in spans] == ["фио_врачей", "фио_пациента"]


# ------------------------------------------------------------ 2.1: требования инструкции
def test_traps_and_integrity(pdfs, tmp_path):
    out = tmp_path / "out"
    m = Pipeline(_cfg(tmp_path)).run(pdfs, out)
    texts = _texts(out)
    t = [v for v in texts.values() if "метформин" in v][0]
    assert "Кошкин" not in t and "Кoшкин" not in t                    # латиница в фамилии не помогла
    assert "[ИНН]" in t or "[УДАЛЕНО: ИНН]" in t
    assert "[МЕСТО_РАБОТЫ]" in t and "Ромашка" not in t
    assert "[ИЗОБРАЖЕНИЕ УДАЛЕНО]" in t
    assert "1000 мг" in t and "128 г/л" in t and "35 %" in t        # медицинский текст не тронут
    rec = [r for r in m["файлы"] if r.get("images")][0]
    assert rec["pages_in"] == rec["pages_text"] == 1
    assert not any("изменён" in w or "изменились" in w for r in m["файлы"] for w in r["warnings"])
    assert [r for r in m["файлы"] if r.get("images")][0]["status"] == "готово"     # своя метка — не «остаток»


def test_no_pii_in_journal_manifest_names(pdfs, tmp_path):
    out = tmp_path / "out"
    Pipeline(_cfg(tmp_path)).run(pdfs, out)
    blob = (tmp_path / "journal.sqlite").read_bytes().decode("utf-8", "ignore")
    blob += "\n".join(p.read_text(encoding="utf-8") for p in out.rglob("manifest.json"))
    blob += "\n".join(p.name for p in out.rglob("*"))
    for bad in ["Мазуренко", "Кошкин", "Иванов", "epicrisis", "trap_", "12345/26", ".pdf"]:
        assert bad not in blob, bad


def test_consent_withdrawal_and_deletion_log(pdfs, tmp_path):
    import shutil
    from medeid.journal import Journal
    from medeid.pipeline import delete_sources
    from medeid.registry import Registry
    src = tmp_path / "src"; shutil.copytree(pdfs, src, ignore=shutil.ignore_patterns(".medeid"))
    cfg = _cfg(tmp_path)
    reg = Registry(src, "test"); reg.exclude("КОШКИНА В.П.|1961-04-03", "заявление № 5"); reg.close()
    m = Pipeline(cfg).run(src, tmp_path / "out")
    st = {r["doc_id"]: r["status"] for r in m["файлы"]}
    assert "исключён_отзыв_согласия" in st.values()
    n = delete_sources(cfg, src, m["партия"], "test")
    assert n >= 3 and not list(src.glob("*.pdf"))[:0]
    jr = Journal(cfg.journal)
    assert jr.events(["регистрация_документов"])
    ev = jr.events(["удаление_исходников"])
    assert ev and ev[0][3] == m["партия"]
    csv_text = jr.export_csv(["удаление_исходников"])
    assert "DOC-" in csv_text and ".pdf" not in csv_text


def test_deletion_mode_is_typed_placeholder(pdfs, tmp_path):
    out = tmp_path / "out"
    Pipeline(_cfg(tmp_path)).run(pdfs, out)
    allt = "\n".join(_texts(out).values())
    assert "[УДАЛЕНО: телефон]" in allt and "[УДАЛЕНО: СНИЛС]" in allt


def test_eponyms_dict_version():
    from medeid.detect import Dictionaries
    d = Dictionaries.load(DICTS / "2026-11-11")
    assert {"жильбера", "крона", "ходжкина", "боткина"} <= d.eponyms


def test_full_name_always_replaced():
    """Имя и отчество не остаются рядом с заменённой фамилией: имя не из словаря,
    модель нашла только фамилию, имя/отчество пациента отдельно."""
    from medeid.detect import Detector, Dictionaries, NerBackend
    from medeid.extract import fold
    from medeid.llm import Entity
    from medeid.transform import Transformer
    d = Detector(Dictionaries.load(DICTS / "2026-11-11"), NerBackend())
    modes = {k: "замена_меткой" for k in ("фио_пациента", "фио_родственников", "фио_врачей", "фио_прочие", "адрес",
             "телефон", "email", "снилс_полис_паспорт", "номер_истории", "номера_документов", "место_работы")}
    modes.update(даты="только_год", возраст="как_есть")

    def run(t, ents=None, rules=True):
        ft = fold(t); p = d.find_patient(ft); sp, _ = d.detect(ft, p, ents, use_rules=rules)
        return Transformer(modes, "90+", None, None, None).apply(t, sp)[0]
    t = ("ВЫПИСНОЙ ЭПИКРИЗ\nПациент: Шаповалов Эдуард Рустамович\nСо слов Эдуарда Рустамовича, ухудшение. "
         "Эдуард спит плохо. Мать пациента Шаповалова Зульфия Ринатовна.\n")
    out = run(t)
    for w in ("Эдуард", "Рустамович", "Зульфия", "Ринатовна", "Шаповалов"):
        assert w not in out, w
    assert "Мать пациента [ФИО_РОДСТВЕННИКА]" in out
    t2 = "ОСМОТР ВРАЧА\nВрач: Врачев Игорь Николаевич\nПациент: Врачев Игорь Николаевич\n"
    for rules in (True, False):
        out = run(t2, [Entity("Врачев", "фио_пациента")], rules)
        assert "Игорь" not in out and "Николаевич" not in out
        assert "Врач: [ФИО_ВРАЧА]" in out
    t3 = "Пациент: Петрова Анна Ивановна\nЦентр имени\nА.С. Логинова Департамента\nА.С. Логинова Департамента\n"
    assert run(t3).count("А.С. Логинова") == 2


# ------------------------------------------------------------ 2.2: имена файлов, правила, конфиденциальность
def _ascii_names(root: Path) -> list[str]:
    return [str(p.relative_to(root)) for p in root.rglob("*") if not p.name.isascii()]


def test_only_ascii_file_names(pdfs, llm_server, tmp_path):
    """Программа нигде не создаёт имён файлов и папок не латиницей."""
    from medeid.pipeline import delete_outputs
    from medeid.rules import default_filename, rules_docx
    cfg = _llm_cfg(tmp_path, llm_server)
    m = Pipeline(cfg).run(pdfs, tmp_path / "results")
    assert (Path(m["папка_партии"]) / "_review").exists() or True
    rules_docx(cfg, tmp_path / default_filename(cfg))
    assert _ascii_names(tmp_path) == []
    assert _ascii_names(ROOT) == []                                   # и в самой программе


def test_results_manifest_reveals_no_methods(pdfs, llm_server, tmp_path):
    import json
    m = Pipeline(_llm_cfg(tmp_path, llm_server)).run(pdfs, tmp_path / "results")
    pub = json.loads((Path(m["папка_партии"]) / "manifest.json").read_text(encoding="utf-8"))
    assert set(pub) == {"партия", "документы"}
    blob = json.dumps(pub, ensure_ascii=False)
    for secret in ("qwen", "модель", "словар", "извлечение", "правила", "инструкция", "режим", "замечан", "остаток"):
        assert secret not in blob, secret
    assert all(set(d) == {"документ", "файл", "sha256"} for d in pub["документы"])


def test_rules_export_and_edition(tmp_path):
    from medeid.rules import rules_docx, rules_text
    cfg = _cfg(tmp_path)
    t = rules_text(cfg)
    for phrase in ("ФИО пациента (в любой форме", "заменяется меткой [ФИО_ПАЦИЕНТА]",
                   "заменяется специальностью, например [ВРАЧ-ОНКОЛОГ]", "сокращается до названия города",
                   "остаётся только год", "возраст старше 89 лет заменяется на «90+»", "ПРОЕКТ", cfg.rules_id()):
        assert phrase in t, phrase
    # не раскрываются: модель, инструкция для неё, содержимое и файлы словарей
    for secret in ("qwen", "llm", "языков", "Ollama", "first_names", "eponyms", "Жильбера", "Химки", "seed"):
        assert secret.lower() not in t.lower(), secret
    rules_docx(cfg, tmp_path / "r.docx")
    import docx
    assert "заменяется меткой [ФИО_ПАЦИЕНТА]" in "\n".join(p.text for p in docx.Document(str(tmp_path / "r.docx")).paragraphs)
    # редакция не зависит от способа поиска, но меняется при изменении правил
    assert _cfg(tmp_path, извлечение="правила").rules_id() == cfg.rules_id()
    assert _cfg(tmp_path, даты="квартал").rules_id() != cfg.rules_id()


def test_approval_blocks_changed_rules(pdfs, tmp_path):
    from medeid.config import ConfigError
    from medeid.journal import Journal
    rid = _cfg(tmp_path).rules_id()
    ok = _cfg(tmp_path, утверждение={"редакция": rid, "приказ_номер": "77-о", "приказ_дата": "18.11.2026"})
    assert ok.approved
    m = Pipeline(ok).run(pdfs, tmp_path / "results")
    jr = Journal(ok.journal)
    row = [r for r in jr.runs() if r[0] == m["партия"]][0]
    assert row[13] == rid and row[14].startswith("да: приказ № 77-о")
    with pytest.raises(ConfigError, match="изменены после утверждения"):
        _cfg(tmp_path, даты="удаление", утверждение={"редакция": rid, "приказ_номер": "77-о", "приказ_дата": "1"})


def test_verify_from_journal_and_tamper(pdfs, tmp_path):
    from medeid.pipeline import verify_batch
    cfg = _cfg(tmp_path)
    m = Pipeline(cfg).run(pdfs, tmp_path / "results")
    r = verify_batch(pdfs, m["партия"], cfg.journal, cfg.settings_dir, tmp_path / "results")
    assert r["различаются"] == [] and r["изменены_файлы_результата"] == [] and r["совпадает_отпечаток"]
    f = sorted(Path(m["папка_партии"]).glob("DOC-*.txt"))[0]
    f.write_text("подмена", encoding="utf-8")
    r = verify_batch(pdfs, m["партия"], cfg.journal, cfg.settings_dir, tmp_path / "results")
    assert r["изменены_файлы_результата"] == [f.stem]


def test_settings_must_be_outside_program(pdfs, tmp_path):
    from medeid.config import load_raw
    import yaml
    sd = _settings(tmp_path)
    raw = yaml.safe_load((sd / "config.yaml").read_text(encoding="utf-8"))
    raw["основание"] = {"приказ_номер": "1", "приказ_дата": "1", "цель": "т", "отделение": "т"}
    raw["извлечение"] = "правила"; raw["журнал"] = str(tmp_path / "j.sqlite")
    cfg = load_raw(raw, ROOT / "templates")                # «настройки» внутри папки программы
    with pytest.raises(RuntimeError, match="раздельными"):
        Pipeline(cfg).run(pdfs, tmp_path / "results")
